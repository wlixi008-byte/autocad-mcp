import win32com.client
import sys

def test_autocad_connection():
    try:
        print("Attempting to connect to AutoCAD...")
        # specific version progID might be needed if multiple versions are installed
        # but usually AutoCAD.Application maps to the latest registered one
        acad = win32com.client.Dispatch("AutoCAD.Application") 
        print(f"Connected to AutoCAD Version: {acad.Version}")
        doc = acad.ActiveDocument
        print(f"Active Document: {doc.Name}")
        
        # Try a simple command to verify control
        # Draw a line from (0,0,0) to (100,100,0) concept
        # We won't actually draw yet, just reading properties is enough for connection test
        # But let's try to get the ModelSpace object
        msp = doc.ModelSpace
        print(f"ModelSpace object obtained: {msp}")
        
        return True
    except Exception as e:
        print(f"Failed to connect: {e}")
        return False

if __name__ == "__main__":
    if test_autocad_connection():
        print("SUCCESS: AutoCAD connection verified.")
    else:
        print("FAILURE: Could not connect to AutoCAD. Make sure it is running.")
