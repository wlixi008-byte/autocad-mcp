import win32com.client
import pythoncom

def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def draw_specific_plan():
    try:
        print("Connecting to AutoCAD...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        msp = doc.ModelSpace
        
        # Dimensions based on the image (typical layout 10m x 11m)
        width = 10000
        depth = 11000
        
        print("Drawing specific layout...")
        
        # 1. Outer Frame (Main Walls)
        # Main rectangle 10000 x 11000
        # Offsets for Porch (Doorway) at bottom
        
        # Let's draw the main bounding box first
        msp.AddLine(APoint(0,0), APoint(width,0))       # Bottom (Main) - simplistic
        msp.AddLine(APoint(width,0), APoint(width,depth)) # Right
        msp.AddLine(APoint(width,depth), APoint(0,depth)) # Top
        msp.AddLine(APoint(0,depth), APoint(0,0))       # Left

        # 2. Internal Layout (Approximation)
        
        # Vertical Split (Left/Right)
        # Left side: Staircase(Top), Living Room(Bottom)
        # Right side: Kitchen(Top), Bathroom(Mid), Bedroom(Bottom)
        mid_x = 5000 
        
        # Vertical wall roughly in the middle
        msp.AddLine(APoint(mid_x, 0), APoint(mid_x, depth))
        
        # Left Side:
        # Stairs area at top left (approx top 4000)
        stair_y = 7000
        msp.AddLine(APoint(0, stair_y), APoint(mid_x, stair_y))
        
        # Right Side:
        # Kitchen at top right (approx top 4000)
        kitchen_y = 7000
        msp.AddLine(APoint(mid_x, kitchen_y), APoint(width, kitchen_y))
        
        # Bathrooms in the middle right (approx 2000 wide strip?)
        # Let's say Bathroom is between y=4500 and y=7000
        bath_y_bottom = 4500
        msp.AddLine(APoint(mid_x, bath_y_bottom), APoint(width, bath_y_bottom))
        
        # Bedroom at bottom right
        # (Already defined by bounds and bath_y_bottom)

        # 3. Features
        
        # Porch (Men Ku) - Pop out at bottom middle? 
        # Usually cuts into the bottom line. Let's add a small extension.
        porch_w = 2000
        porch_h = 1500
        porch_x = (width - porch_w) / 2 # Centered
        # msp.AddLine(APoint(porch_x, 0), APoint(porch_x, -porch_h))
        # msp.AddLine(APoint(porch_x + porch_w, 0), APoint(porch_x + porch_w, -porch_h))
        # msp.AddLine(APoint(porch_x, -porch_h), APoint(porch_x + porch_w, -porch_h))

        # 4. Text Labels (Chinese)
        text_h = 250
        
        # Top Left: Stairs (楼梯)
        msp.AddText("楼梯 (Stairs)", APoint(1500, 8500), text_h)
        
        # Center/Top: Dining (餐厅)
        # Often in the middle area if open plan, or specific room.
        # Image seems to show Dining in middle top or left.
        # Let's put Dining near Kitchen
        msp.AddText("餐厅 (Dining)", APoint(3500, 7500), text_h)
        
        # Top Right: Kitchen (厨房)
        msp.AddText("厨房 (Kitchen)", APoint(7500, 9000), text_h)
        
        # Middle Right: Bathroom (卫生间)
        msp.AddText("卫生间 (Bath)", APoint(7500, 5500), text_h)
        
        # Bottom Left: Living Room (客厅)
        msp.AddText("客厅 (Living Room)", APoint(2500, 3500), text_h)
        
        # Bottom Right: Bedroom (卧室)
        msp.AddText("卧室 (Bedroom)", APoint(7500, 2500), text_h)
        
        # Bottom Center: Porch (门廊)
        msp.AddText("门廊 (Porch)", APoint(5000, 500), text_h)

        print("Zooming extents...")
        doc.SendCommand("_ZOOM _E \n")
        print("Drawing complete.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    draw_specific_plan()
