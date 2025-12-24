import asyncio
import logging
from typing import Optional
import win32com.client
import pythoncom
from mcp.server.models import InitializationOptions
import mcp.types as types
from mcp.server import NotificationOptions, Server
import mcp.server.stdio

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("autocad-server")

# Global AutoCAD Application instance
acad = None

def get_autocad():
    """Connect to the running AutoCAD application."""
    global acad
    try:
        # Re-connect if lost (conceptually, though dispatch usually persists or throws)
        # Note: In a thread/async loop, we might need CoInitialize
        pythoncom.CoInitialize() 
        acad = win32com.client.Dispatch("AutoCAD.Application")
        return acad
    except Exception as e:
        logger.error(f"Failed to connect to AutoCAD: {e}")
        return None

def APoint(x, y, z=0):
    """Helper to create an ActiveX point."""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

async def main():
    server = Server("autocad-mcp")

    @server.list_tools()
    async def handle_list_tools() -> list[types.Tool]:
        return [
            types.Tool(
                name="send_command",
                description="Send a command string to the AutoCAD command line. Use this for complex actions not covered by other tools.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "command": {"type": "string", "description": "The command string to execute (e.g. '_LINE 0,0 100,100 ')"},
                    },
                    "required": ["command"],
                },
            ),
             types.Tool(
                name="draw_line",
                description="Draw a straight line between two points.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "start_point": {
                            "type": "array",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3,
                            "description": "Start [x, y, z] coordinate"
                        },
                        "end_point": {
                            "type": "array",
                             "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3,
                            "description": "End [x, y, z] coordinate"
                        },
                    },
                    "required": ["start_point", "end_point"],
                },
            ),
            types.Tool(
                name="draw_circle",
                description="Draw a circle given a center point and radius.",
                inputSchema={
                    "type": "object",
                     "properties": {
                        "center": {
                            "type": "array",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3,
                            "description": "Center [x, y, z] coordinate"
                        },
                        "radius": {"type": "number", "description": "Radius of the circle"},
                    },
                    "required": ["center", "radius"],
                },
            ),
             types.Tool(
                name="add_text",
                description="Add a text object to the drawing.",
                inputSchema={
                    "type": "object",
                     "properties": {
                        "text": {"type": "string", "description": "The content string"},
                        "insertion_point": {
                            "type": "array",
                            "items": {"type": "number"},
                            "minItems": 2,
                            "maxItems": 3,
                            "description": "Insertion point [x, y, z]"
                        },
                        "height": {"type": "number", "description": "Text height"},
                    },
                    "required": ["text", "insertion_point", "height"],
                },
            ),
             types.Tool(
                name="get_active_document_name",
                description="Get the name of the currently active drawing.",
                inputSchema={"type": "object", "properties": {}},
            ),
             types.Tool(
                name="read_layout",
                description="Read the current layout from ModelSpace. Returns a JSON list of lines, circles, and text.",
                inputSchema={"type": "object", "properties": {}},
            ),
        ]

    @server.call_tool()
    async def handle_call_tool(
        name: str, arguments: dict | None
    ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
        app = get_autocad()
        if not app:
             return [types.TextContent(type="text", text="Error: Could not connect to AutoCAD. Is it running?")]
        
        doc = app.ActiveDocument
        msp = doc.ModelSpace
        arguments = arguments or {}

        try:
            if name == "send_command":
                cmd = arguments.get("command", "")
                # SendCommand is synchronous in COM but sends to command line buffer
                # It behaves like typing in the GUI. A trailing space/newline often needed.
                doc.SendCommand(cmd + "\n") 
                return [types.TextContent(type="text", text=f"Command sent: {cmd}")]

            elif name == "draw_line":
                start = arguments["start_point"]
                end = arguments["end_point"]
                # Ensure 3D points
                if len(start) == 2: start.append(0.0)
                if len(end) == 2: end.append(0.0)
                
                line = msp.AddLine(APoint(*start), APoint(*end))
                return [types.TextContent(type="text", text=f"Line drawn. Handle: {line.Handle}")]

            elif name == "draw_circle":
                center = arguments["center"]
                radius = arguments["radius"]
                if len(center) == 2: center.append(0.0)
                
                circle = msp.AddCircle(APoint(*center), radius)
                return [types.TextContent(type="text", text=f"Circle drawn. Handle: {circle.Handle}")]
            
            elif name == "add_text":
                text_content = arguments["text"]
                point = arguments["insertion_point"]
                height = arguments["height"]
                if len(point) == 2: point.append(0.0)
                
                text_obj = msp.AddText(text_content, APoint(*point), height)
                return [types.TextContent(type="text", text=f"Text added. Handle: {text_obj.Handle}")]

            elif name == "get_active_document_name":
                return [types.TextContent(type="text", text=f"Active Document: {doc.Name}")]

            elif name == "read_layout":
                # Extract entities from ModelSpace
                # Limit to avoid overwhelming context window for now (e.g. 500 entities)
                max_entities = 500
                entities_data = []
                count = 0
                
                for entity in msp:
                    if count >= max_entities: 
                        break
                    
                    try:
                        obj_name = entity.ObjectName
                        data = {"type": obj_name, "handle": entity.Handle}
                        
                        if obj_name == "AcDbLine":
                            # StartPoint/EndPoint return tuples/variants
                            data["start"] = list(entity.StartPoint)
                            data["end"] = list(entity.EndPoint)
                            entities_data.append(data)
                            count += 1
                        
                        elif obj_name == "AcDbCircle":
                            data["center"] = list(entity.Center)
                            data["radius"] = entity.Radius
                            entities_data.append(data)
                            count += 1
                            
                        elif obj_name == "AcDbText" or obj_name == "AcDbMText":
                            data["text"] = entity.TextString
                            data["insertion_point"] = list(entity.InsertionPoint)
                            entities_data.append(data)
                            count += 1
                            
                        elif obj_name == "AcDbPolyline":
                            # Polyline coords are slightly more complex (flat list of x,y)
                            # Skipping complex parsing for brevity in this demo, just noting existence
                            data["coordinates"] = "Polyline data (omitted)"
                            entities_data.append(data)
                            count += 1
                            
                    except Exception as entity_err:
                        # Skip entity if read fails
                        continue

                import json
                return [types.TextContent(type="text", text=json.dumps(entities_data, indent=2))]

            else:
                 raise ValueError(f"Unknown tool: {name}")

        except Exception as e:
            return [types.TextContent(type="text", text=f"Error executing {name}: {str(e)}")]

    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="autocad-mcp",
                server_version="0.1.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    asyncio.run(main())
