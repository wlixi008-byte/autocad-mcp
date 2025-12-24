import win32com.client
import pythoncom

def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def draw_floor_plan():
    try:
        print("Connecting to AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        
        # Define some dimensions (assuming generic units, e.g., mm or cm)
        # Structure: A simple 2-bedroom apart layout
        # Outer bounds: 0,0 to 12000, 8000
        
        print("Drawing walls...")
        
        # Outer Walls (Simple Rectangle)
        # Using specific lines instead of polyline for simplicity in this script
        # Bottom
        msp.AddLine(APoint(0,0), APoint(12000,0))
        # Top
        msp.AddLine(APoint(0,8000), APoint(12000,8000))
        # Left
        msp.AddLine(APoint(0,0), APoint(0,8000))
        # Right
        msp.AddLine(APoint(12000,0), APoint(12000,8000))

        # Internal Walls
        # Middle vertical wall splitting Living/Bedrooms
        msp.AddLine(APoint(6000, 0), APoint(6000, 8000))
        
        # Horizontal wall splitting Bedroom 1 and Bedroom 2 (Right side)
        msp.AddLine(APoint(6000, 4000), APoint(12000, 4000))
        
        # Bathroom (Top Left corner of Living room area)
        msp.AddLine(APoint(0, 5000), APoint(3000, 5000)) # Bottom wall of bath
        msp.AddLine(APoint(3000, 5000), APoint(3000, 8000)) # Right wall of bath

        print("Adding text labels...")
        # Text Labels
        height = 300
        msp.AddText("LIVING ROOM", APoint(3000, 2500), height)
        msp.AddText("BATHROOM", APoint(1000, 6500), height)
        msp.AddText("BEDROOM 1", APoint(8000, 6000), height)
        msp.AddText("BEDROOM 2", APoint(8000, 2000), height)

        print("Zooming extents...")
        doc.SendCommand("_ZOOM _E \n")
        
        print("Floor plan drawn successfully!")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    draw_floor_plan()
