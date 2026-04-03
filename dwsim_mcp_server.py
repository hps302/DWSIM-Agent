import clr
import os
import sys
import asyncio
import io
import json
import time

# ── Load DWSIM .NET libraries ────────────────────────────────
DWSIM_PATH = r"C:\Users\himan\AppData\Local\DWSIM"
sys.path.append(DWSIM_PATH)

# ── Suppress DWSIM startup warnings from polluting MCP JSON stream ──
_stdout = sys.stdout
sys.stdout = io.StringIO()  # capture stdout during DWSIM loading

clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.Automation.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.Interfaces.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.GlobalSettings.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.SharedClasses.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.Thermodynamics.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.UnitOperations.dll"))
clr.AddReference(os.path.join(DWSIM_PATH, "DWSIM.FlowsheetSolver.dll"))

from DWSIM.Automation import Automation3

# ── Load simulation file ─────────────────────────────────────
SIM_FILE = r"C:\Current Items\AI agents\DWSIM Automation\Test_File_Manual.dwxmz"
sim = Automation3()
flowsheet = sim.LoadFlowsheet(SIM_FILE)

# ── Restore stdout for MCP communication ────────────────────
sys.stdout = _stdout

# ── MCP Server setup ─────────────────────────────────────────
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

server = Server("dwsim-server")

# ── Helper function ──────────────────────────────────────────
from DWSIM.Thermodynamics.Streams import MaterialStream
from DWSIM.UnitOperations.UnitOperations import DistillationColumn

def get_object(name):
    for obj in flowsheet.SimulationObjects.Values:
        if obj.GraphicObject.Tag == name:
            return obj
    raise ValueError(f"Object '{name}' not found in flowsheet.")

def get_stream(name):
    for obj in flowsheet.SimulationObjects.Values:
        if obj.GraphicObject.Tag == name:
            return obj
    raise ValueError(f"Stream '{name}' not found in flowsheet.")

def get_column(name):
    for obj in flowsheet.SimulationObjects.Values:
        if obj.GraphicObject.Tag == name:
            return obj
    raise ValueError(f"Column '{name}' not found in flowsheet.")

# ── Tool 1: List all flowsheet objects ───────────────────────
@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="list_objects",
            description="List all simulation object names in the DWSIM flowsheet",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="get_stream_composition",
            description="Get the mole fractions of all compounds in a named material stream",
            inputSchema={
                "type": "object",
                "properties": {"stream_name": {"type": "string"}},
                "required": ["stream_name"]
            }
        ),
        Tool(
            name="get_stream_flowrate",
            description="Get the mass flow rate of a named material stream in kg/s",
            inputSchema={
                "type": "object",
                "properties": {"stream_name": {"type": "string"}},
                "required": ["stream_name"]
            }
        ),
        Tool(
            name="get_stream_temperature",
            description="Get the temperature of a named material stream in Kelvin",
            inputSchema={
                "type": "object",
                "properties": {"stream_name": {"type": "string"}},
                "required": ["stream_name"]
            }
        ),
        Tool(
            name="get_column_duties",
            description="Get the condenser and reboiler duties of the distillation column in kW",
            inputSchema={
                "type": "object",
                "properties": {"column_name": {"type": "string"}},
                "required": ["column_name"]
            }
        ),
        Tool(
            name="get_column_spec",
            description="Get the current reflux ratio and boilup ratio of the distillation column",
            inputSchema={
                "type": "object",
                "properties": {"column_name": {"type": "string"}},
                "required": ["column_name"]
            }
        ),
        Tool(
            name="set_reflux_ratio",
            description="Set the reflux ratio of the distillation column and recalculate the flowsheet",
            inputSchema={
                "type": "object",
                "properties": {
                    "column_name": {"type": "string"},
                    "value": {"type": "number"}
                },
                "required": ["column_name", "value"]
            }
        ),
        Tool(
            name="set_feed_temperature",
            description="Set the temperature of the feed stream in Kelvin and recalculate",
            inputSchema={
                "type": "object",
                "properties": {
                    "stream_name": {"type": "string"},
                    "temperature_K": {"type": "number"}
                },
                "required": ["stream_name", "temperature_K"]
            }
        ),
        Tool(
            name="solve_flowsheet",
            description="Trigger a full recalculation of the DWSIM flowsheet",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="get_full_summary",
            description="Get a complete summary of the distillation column including feed, distillate, bottoms compositions, flowrates, and column duties",
            inputSchema={"type": "object", "properties": {}}
        ),
        Tool(
            name="debug_stream",
            description="Debug a stream object to see available attributes",
            inputSchema={"type": "object", "properties": {"stream_name": {"type": "string"}}, "required": ["stream_name"]}
        ),
    ]

# ── Tool implementations ─────────────────────────────────────
@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:

    if name == "list_objects":
        names = [obj.GraphicObject.Tag for obj in flowsheet.SimulationObjects.Values]
        return [TextContent(type="text", text="\n".join(names))]

    elif name == "get_stream_composition":
        stream = get_stream(arguments["stream_name"])
        methanol = stream.GetPropertyValue("PROP_MS_102/Methanol")
        water = stream.GetPropertyValue("PROP_MS_102/Water")
        return [TextContent(type="text", text=f"Methanol: {round(float(methanol), 4)} mol frac\nWater: {round(float(water), 4)} mol frac")]

    elif name == "get_stream_flowrate":
        stream = get_stream(arguments["stream_name"])
        flow = stream.GetPropertyValue("PROP_MS_2")
        return [TextContent(type="text", text=f"Mass flow: {round(float(flow), 4)} kg/s")]

    elif name == "get_stream_temperature":
        stream = get_stream(arguments["stream_name"])
        temp = stream.GetPropertyValue("PROP_MS_0")
        return [TextContent(type="text", text=f"Temperature: {round(float(temp), 2)} K")]

    elif name == "get_column_duties":
        col = get_column(arguments["column_name"])
        cond = round(float(col.GetPropertyValue("PROP_DC_5")) / 1000, 2)
        reb = round(float(col.GetPropertyValue("PROP_DC_6")) / 1000, 2)
        return [TextContent(type="text", text=f"Condenser Duty: {cond} kW\nReboiler Duty: {reb} kW")]

    elif name == "get_column_spec":
        col = get_column(arguments["column_name"])
        rr = float(col.GetPropertyValue("Condenser_Specification_Value"))
        br = float(col.GetPropertyValue("Reboiler_Specification_Value"))
        return [TextContent(type="text", text=f"Reflux Ratio: {rr}\nBoilup Ratio: {br}")]

    elif name == "set_reflux_ratio":
        col = get_column(arguments["column_name"])
        dist = get_stream("DIST")
        bot = get_stream("BOTTOMS")

        old_rr = float(col.GetPropertyValue("Condenser_Specification_Value"))
        old_methanol = round(float(dist.GetPropertyValue("PROP_MS_102/Methanol")), 4)

        col.SetPropertyValue("Condenser_Specification_Value", float(arguments["value"]))
        flowsheet.RequestCalculationAndWait()

        new_rr = float(col.GetPropertyValue("Condenser_Specification_Value"))
        new_methanol = round(float(dist.GetPropertyValue("PROP_MS_102/Methanol")), 4)
        methanol_bot = round(float(bot.GetPropertyValue("PROP_MS_102/Methanol")), 4)
        cond = round(float(col.GetPropertyValue("PROP_DC_5")) / 1000, 2)
        reb = round(float(col.GetPropertyValue("PROP_DC_6")) / 1000, 2)

        converged = abs(new_rr - float(arguments["value"])) < 0.001

        result = {
            "converged": converged,
            "reflux_ratio_requested": float(arguments["value"]),
            "reflux_ratio_actual": new_rr,
            "distillate_methanol_before": old_methanol,
            "distillate_methanol_after": new_methanol,
            "bottoms_methanol": methanol_bot,
            "condenser_duty_kW": cond,
            "reboiler_duty_kW": reb,
            "note": "" if converged else f"WARNING: Solver did not converge at RR={arguments['value']}. DWSIM reverted to RR={new_rr}."
        }
        return [TextContent(type="text", text=json.dumps(result, indent=2))]

    elif name == "set_feed_temperature":
        stream = get_stream(arguments["stream_name"])
        stream.SetPropertyValue("PROP_MS_0", float(arguments["temperature_K"]))
        flowsheet.RequestCalculationAndWait()
        new_temp = float(stream.GetPropertyValue("PROP_MS_0"))
        return [TextContent(type="text", text=f"Feed temperature set to {round(new_temp, 2)} K and flowsheet recalculated.")]

    elif name == "debug_stream":
        for obj in flowsheet.SimulationObjects.Values:
            if obj.GraphicObject.Tag == arguments["stream_name"]:
                try:
                    props = list(obj.GetProperties(0))
                    return [TextContent(type="text", text="\n".join(str(p) for p in props[:50]))]
                except Exception as e1:
                    try:
                        props = list(obj.GetProperties())
                        return [TextContent(type="text", text="\n".join(str(p) for p in props[:50]))]
                    except Exception as e2:
                        test_props = ["temperature", "pressure", "massflow",
                                      "molarflow", "Methanol", "Water"]
                        results = {}
                        for p in test_props:
                            try:
                                results[p] = str(obj.GetPropertyValue(p))
                            except Exception as e:
                                results[p] = f"Error: {str(e)}"
                        return [TextContent(type="text", text=str(results))]

    elif name == "solve_flowsheet":
        flowsheet.RequestCalculationAndWait()
        return [TextContent(type="text", text="Flowsheet recalculated successfully.")]

    elif name == "get_full_summary":
        feed = get_stream("MeOH_Water")
        dist = get_stream("DIST")
        bot  = get_stream("BOTTOMS")
        col  = get_column("T1")

        summary = {
            "Feed": {
                "Methanol_mol_frac": round(float(feed.GetPropertyValue("PROP_MS_102/Methanol")), 4),
                "Water_mol_frac": round(float(feed.GetPropertyValue("PROP_MS_102/Water")), 4),
                "mass_flow_kg_s": round(float(feed.GetPropertyValue("PROP_MS_2")), 4),
                "temperature_K": round(float(feed.GetPropertyValue("PROP_MS_0")), 2)
            },
            "Distillate": {
                "Methanol_mol_frac": round(float(dist.GetPropertyValue("PROP_MS_102/Methanol")), 4),
                "Water_mol_frac": round(float(dist.GetPropertyValue("PROP_MS_102/Water")), 4),
                "mass_flow_kg_s": round(float(dist.GetPropertyValue("PROP_MS_2")), 4)
            },
            "Bottoms": {
                "Methanol_mol_frac": round(float(bot.GetPropertyValue("PROP_MS_102/Methanol")), 4),
                "Water_mol_frac": round(float(bot.GetPropertyValue("PROP_MS_102/Water")), 4),
                "mass_flow_kg_s": round(float(bot.GetPropertyValue("PROP_MS_2")), 4)
            },
            "Column": {
                "reflux_ratio": float(col.GetPropertyValue("Condenser_Specification_Value")),
                "boilup_ratio": float(col.GetPropertyValue("Reboiler_Specification_Value")),
                "condenser_duty_kW": round(float(col.GetPropertyValue("PROP_DC_5")) / 1000, 2),
                "reboiler_duty_kW": round(float(col.GetPropertyValue("PROP_DC_6")) / 1000, 2)
            }
        }
        return [TextContent(type="text", text=json.dumps(summary, indent=2))]

    else:
        return [TextContent(type="text", text=f"Unknown tool: {name}")]

# ── Run server ───────────────────────────────────────────────
async def main():
    async with stdio_server() as streams:
        await server.run(
            streams[0],
            streams[1],
            server.create_initialization_options()
        )

if __name__ == "__main__":
    asyncio.run(main())