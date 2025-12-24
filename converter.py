import win32com.client
import pythoncom
import os

def extract_entities_from_file(filepath):
    """
    Opens a DWG file (or uses active one if path matches) and extracts entities.
    Returns a list of entity dictionaries.
    """
    entities = []
    doc = None
    try:
        # REQUIRED for Streamlit (threading): Initialize COM
        pythoncom.CoInitialize()
        
        # 1. Convert to absolute path (COM requires full paths)
        filepath = os.path.abspath(filepath)
        print(f"DEBUG: Attempting to read: {filepath}")

        # 2. Connect to AutoCAD
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application")
            acad.Visible = True
        except Exception as e:
             print(f"ERROR: Could not connect to AutoCAD Application. {e}")
             return []

        # 3. Open Logic
        # Try to find if it's already open
        opened_here = False # Did we open it?
        
        for d in acad.Documents:
            if d.FullName.lower() == filepath.lower():
                doc = d
                print("DEBUG: Document was already open.")
                break
        
        if not doc:
            try:
                # IMPORTANT: active document interaction can sometimes interfere
                # We simply ask to open. 
                # Note: DWG must exist.
                if not os.path.exists(filepath):
                    print(f"ERROR: File does not exist on disk: {filepath}")
                    return []
                    
                doc = acad.Documents.Open(filepath)
                opened_here = True
                print("DEBUG: Document opened successfully.")
            except Exception as e:
                print(f"ERROR: Failed to open document via COM. {e}")
                # Fallback: Maybe try to access ActiveDocument if the user just opened it?
                # But creating a new doc is safer.
                return []

        # 4. Read Entities from ModelSpace
        msp = doc.ModelSpace
        print(f"DEBUG: ModelSpace accessed. Count: {msp.Count}")
        
        # Limit processing for performance
        max_limit = 1000
        count = 0
        
        for entity in msp:
            if count >= max_limit: break
            try:
                obj_name = entity.ObjectName
                # print(f"DEBUG: Found entity {obj_name}") # Verbose
                
                data = {'type': obj_name}
                
                if obj_name == "AcDbLine":
                    data['start'] = list(entity.StartPoint)
                    data['end'] = list(entity.EndPoint)
                    entities.append(data)
                    count += 1
                elif obj_name == "AcDbCircle":
                    data['center'] = list(entity.Center)
                    data['radius'] = entity.Radius
                    entities.append(data)
                    count += 1
                elif obj_name == "AcDbText" or obj_name == "AcDbMText":
                    data['text'] = entity.TextString
                    # MText and Text handle coords slightly differently but usually both have InsertionPoint
                    data['point'] = list(entity.InsertionPoint)
                    data['height'] = entity.Height
                    entities.append(data)
                    count += 1
            except Exception as ent_err:
                print(f"WARN: Failed to read entity: {ent_err}")
                pass
                
        # 5. Cleanup
        # If we opened it solely for reading, close it to avoid clutter
        # But for debugging now, let's leave it open or handle carefully
        if opened_here:
            print("DEBUG: Closing document after read.")
            # doc.Close(False) # False = Do not save. commented out for safety/debug
 
        
    except Exception as e:
        print(f"Error reading DWG: {e}")
        return []

    return entities

def generate_python_skill(entities):
    """
    Generates a Python script string that recreates the given entities.
    """
    code = []
    code.append("import win32com.client")
    code.append("import pythoncom")
    code.append("")
    code.append("def APoint(x, y, z=0):")
    code.append("    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))")
    code.append("")
    code.append("def generated_skill():")
    code.append("    try:")
    code.append("        acad = win32com.client.Dispatch('AutoCAD.Application')")
    code.append("        doc = acad.ActiveDocument")
    code.append("        msp = doc.ModelSpace")
    code.append("        print('Executing generated skill...')")
    code.append("")
    
    for e in entities:
        if e['type'] == "AcDbLine":
            s = e['start']
            end = e['end']
            code.append(f"        msp.AddLine(APoint({s[0]}, {s[1]}, {s[2]}), APoint({end[0]}, {end[1]}, {end[2]}))")
        elif e['type'] == "AcDbCircle":
            c = e['center']
            r = e['radius']
            code.append(f"        msp.AddCircle(APoint({c[0]}, {c[1]}, {c[2]}), {r})")
        elif e['type'] == "AcDbText":
            t = e['text'].replace("'", "\\'") # Escape basic quotes
            p = e['point']
            h = e['height']
            code.append(f"        msp.AddText('{t}', APoint({p[0]}, {p[1]}, {p[2]}), {h})")
            
    code.append("")
    code.append("        print('Skill execution complete.')")
    code.append("        doc.SendCommand('_ZOOM _E \\n')")
    code.append("    except Exception as e:")
    code.append("        print(f'Error: {e}')")
    code.append("")
    code.append("if __name__ == '__main__':")
    code.append("    generated_skill()")
    
    return "\n".join(code)
