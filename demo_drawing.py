import win32com.client
import pythoncom
import math

def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def demo_draw():
    try:
        print("Connecting to AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        print(f"Connected to {doc.Name}")

        # Draw a fancy star or something visible
        print("Drawing demo shapes...")
        
        # 1. Circle
        center = APoint(100, 100, 0)
        radius = 50
        circle = msp.AddCircle(center, radius)
        print(" - Circle drawn")

        # 2. Text
        text_string = "Hello from Antigravity AI!"
        insertion_point = APoint(100, 160, 0)
        height = 10
        text = msp.AddText(text_string, insertion_point, height)
        print(" - Text added")

        # 3. Line
        line = msp.AddLine(APoint(50, 50, 0), APoint(150, 150, 0))
        print(" - Line drawn")
        
        # Zoom extents to show what we did
        # SendCommand needs a newline to execute
        doc.SendCommand("_ZOOM _E \n")
        print(" - Zoom Extents command sent")

        print("Demo completed successfully!")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    demo_draw()
