import win32com.client
import pythoncom
import json

def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def smart_architect():
    try:
        print("1. Connecting to AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        
        # --- READ PHASE ---
        print("2. Reading current layout...")
        # Simulate what the MCP tool 'read_layout' does:
        entities = []
        for entity in msp:
            try:
                obj_name = entity.ObjectName
                if obj_name == "AcDbText":
                    entities.append({
                        "type": "text",
                        "text": entity.TextString,
                        "point": list(entity.InsertionPoint)
                    })
                elif obj_name == "AcDbLine":
                    entities.append({
                        "type": "line",
                        "start": list(entity.StartPoint),
                        "end": list(entity.EndPoint)
                    })
            except:
                pass
        
        print(f"   Found {len(entities)} entities.")
        
        # --- ANALYZE & GENERATE RULES PHASE ---
        print("3. Generating Design Rules...")
        # Rule 1: Find "Living Room" (客厅) and place a "Sofa" and "TV"
        # Rule 2: Find "Bedroom" (卧室) and place a "Bed"
        # Rule 3: Find "Dining" (餐厅) and draw a table (Circle)
        
        new_items = []
        
        for e in entities:
             if e["type"] == "text":
                text = e["text"]
                x, y, z = e["point"]
                
                if "客厅" in text or "Living" in text:
                    print("   - Found Living Room! Designing furniture...")
                    # Add Sofa (Rectangle represented by lines)
                    new_items.append({"type": "text", "text": "New Sofa", "x": x, "y": y - 1000, "h": 200})
                    # Add TV
                    new_items.append({"type": "text", "text": "New TV Unit", "x": x, "y": y + 1000, "h": 200})
                    
                if "卧室" in text or "Bedroom" in text:
                    print("   - Found Bedroom! Adding Bed...")
                    new_items.append({"type": "text", "text": "King Size Bed", "x": x, "y": y - 1000, "h": 200})
                    
                if "餐厅" in text or "Dining" in text:
                    print("   - Found Dining! Adding Table...")
                    new_items.append({"type": "circle", "x": x, "y": y - 500, "r": 600})

        # --- DRAW PHASE ---
        print("4. Executing New Design...")
        if not new_items:
            print("   No recognizable rooms found to furnish. (Make sure you ran the previous floor plan script!)")
        
        for item in new_items:
            if item["type"] == "text":
                msp.AddText(item["text"], APoint(item["x"], item["y"]), item["h"])
            elif item["type"] == "circle":
                msp.AddCircle(APoint(item["x"], item["y"]), item["r"])
        
        print("All done! New design elements added.")
        doc.SendCommand("_ZOOM _E \n")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    smart_architect()
