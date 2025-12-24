import win32com.client
import pythoncom
import math

class SmartLayoutGenerator:
    def __init__(self):
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            try:
                self.doc = self.acad.ActiveDocument
            except:
                print("No active document found. Creating new...")
                self.doc = self.acad.Documents.Add()
                
            self.msp = self.doc.ModelSpace
        except Exception as e:
            print(f"Failed to connect to AutoCAD: {e}")
            self.acad = None

    def _apoint(self, x, y, z=0):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def draw_rect(self, x, y, width, depth, layer="WALL"):
        """Draws a rectangle (room) and returns the polyline."""
        # Points: Bottom-Left -> Bottom-Right -> Top-Right -> Top-Left -> Close
        points = [x, y, 0, x+width, y, 0, x+width, y+depth, 0, x, y+depth, 0, x, y, 0]
        # In simple mode we adhere to using Lines for max compatibility or light polylines
        # Let's use AddLine for clarity in this demo 
        l1 = self.msp.AddLine(self._apoint(x, y), self._apoint(x+width, y))
        l2 = self.msp.AddLine(self._apoint(x+width, y), self._apoint(x+width, y+depth))
        l3 = self.msp.AddLine(self._apoint(x+width, y+depth), self._apoint(x, y+depth))
        l4 = self.msp.AddLine(self._apoint(x, y+depth), self._apoint(x, y))
        return [l1, l2, l3, l4]

    def add_label(self, text, x, y, height=250):
        self.msp.AddText(text, self._apoint(x, y), height)

    def draw_room(self, name, x, y, w, d, text_offset_x=None, text_offset_y=None):
        """Draws a room boundary and adds a label."""
        self.draw_rect(x, y, w, d)
        
        # Calculate center for text if offsets not provided
        tx = x + w/2 if text_offset_x is None else x + text_offset_x
        ty = y + d/2 if text_offset_y is None else y + text_offset_y
        
        # Approximate centering logic for text (assuming char width ~height)
        tx -= (len(name) * 250 / 2) 
        
        self.add_label(name, tx, ty)

    def generate(self, area):
        if not self.acad:
            print("AutoCAD not connected.")
            return

        print(f"Generating layout for {area} sqm...")
        self.doc.Utility.Prompt(f"\\nGenerative AI: Creating {area}sqm layout...\\n")

        # ---------------------------------------------------------
        # Logic Branching (The "Architectural Intelligence")
        # ---------------------------------------------------------

        if 85 <= area <= 105:
            self._layout_90_compact()
        elif 110 <= area <= 135:
            self._layout_120_standard()
        elif 140 <= area <= 180:
            self._layout_160_luxury()
        else:
            print("Area out of trained range (90-160). Defaulting to 120.")
            self._layout_120_standard()
            
        self.doc.SendCommand("_ZOOM _E \\n")
        print("Done.")

    def _layout_90_compact(self):
        """
        90 sqm Compact: 3 Bed 1 Bath
        Structure: ~10m x 9m
        Typical High-rise Layout:
        South: Living (3.6) + Bed1 (3.2) + Bed2 (3.0) -> No, usually 90sqm is 3 bays or 2 bays.
        Let's do a classic 3-Bay South (San-Kai-Jian)
        Widths: Living(3600), Bed(3300), Bed(3000)
        Depths: Varies
        """
        base_x, base_y = 0, 0
        
        # --- South Bay ---
        # Master Bedroom (Right-South)
        self.draw_room("主卧 (Master)", base_x + 6900, base_y, 3300, 4200)
        # Living Room (Center-South) - Connects to Dining usually
        self.draw_room("客厅 (Living)", base_x + 3300, base_y, 3600, 4200) # Only drawing main box
        # Second Bedroom (Left-South)
        self.draw_room("次卧 (Bed 2)", base_x, base_y, 3300, 3600)
        
        # --- North Bay ---
        # Bathroom (Left-North) - Behind Bed 2
        self.draw_room("卫 (Bath)", base_x, base_y + 3600, 2000, 2400)
        # Kitchen (Left-North-Top) - Behind Bath? Or Next to Dining?
        # Let's put Kitchen next to Entry
        # Dining (Center-North) - Behind Living
        self.draw_room("餐厅 (Dining)", base_x + 3300, base_y + 4200, 3600, 3000)
        # Kitchen (Right-North) - Behind Master
        self.draw_room("厨房 (Kitchen)", base_x + 6900, base_y + 4200, 2500, 3000)
        # Study/Small Bed (North-Far-Right) -- Maybe next to kitchen?
        # Let's add a small room behind master bed?
        self.draw_room("书房 (Study)", base_x + 6900 + 2500, base_y + 4200, 2500, 3000) # Bit awkward, simplified
        
        self.add_label("90㎡ 紧凑三居", base_x + 5000, base_y - 1000, 400)

    def _layout_120_standard(self):
        """
        120 sqm Standard: 3 Bed 2 Bath
        Comfortable sizes.
        South: Master (3.6), Living (4.2), Bed2 (3.3)
        """
        base_x, base_y = 15000, 0 # Offset so we don't overlap if drawing multiple
        
        # --- South Zone ---
        # Bed 2 (Left)
        self.draw_room("次卧 (Bed)", base_x, base_y, 3300, 3900)
        # Living (Center)
        self.draw_room("客厅 (Living)", base_x + 3300, base_y, 4200, 5000)
        # Master Zone (Right)
        # Master Bed
        self.draw_room("主卧 (Master)", base_x + 7500, base_y, 3600, 4500)
        # Master Bath (North of Master)
        self.draw_room("主卫", base_x + 7500, base_y + 4500, 2000, 2400)
        
        # --- North Zone ---
        # Dining (Center, North of Living)
        self.draw_room("餐厅 (Dining)", base_x + 3300, base_y + 5000, 4200, 3500)
        # Kitchen (Right of Dining, North of Master Bath)
        self.draw_room("厨房 (Kitchen)", base_x + 7500, base_y + 4500 + 2400, 3000, 3000) # L-shape filler
        # Guest Bath (Left of Dining, North of Bed 2)
        self.draw_room("客卫", base_x + 1300, base_y + 3900, 2000, 2400)
        # North Bed (Study)
        self.draw_room("北卧 (Guest)", base_x + 3300, base_y + 8500, 3300, 3300) # Pop out at top
        
        self.add_label("120㎡ 舒适三居", base_x + 5000, base_y - 1000, 400)

    def _layout_160_luxury(self):
        """
        160 sqm Luxury: 4 Bed 2 Bath (Wide Living)
        Feature: Horizontal Hall (Heng Ting)
        """
        base_x, base_y = 30000, 0
        
        # Wide Living + Dining Combo (South)
        # 7.0m Wide
        self.draw_room("横厅 (Grand Living)", base_x + 3600, base_y, 7000, 6000)
        
        # Master Suite (East/Right)
        self.draw_room("主卧套房", base_x + 10600, base_y, 3900, 5000)
        self.draw_room("衣帽间", base_x + 10600, base_y + 5000, 2000, 3000)
        self.draw_room("主卫", base_x + 12600, base_y + 5000, 1900, 3000)
        
        # West Wing (Bedrooms)
        self.draw_room("南次卧", base_x, base_y, 3600, 4000)
        self.draw_room("客卫", base_x, base_y + 4000, 2400, 3000) # Public bath accessible from living?
        
        # North Zone
        # Kitchen (Behind Dining/Living)
        self.draw_room("中西厨 (Kitchen)", base_x + 5000, base_y + 6000, 3000, 4000) # Center
        # North Bed 1
        self.draw_room("北次卧", base_x, base_y + 7000, 3300, 3600)
        # North Bed 2 (or Study)
        self.draw_room("书房/保姆间", base_x + 8000, base_y + 6000, 2600, 3000)
        
        self.add_label("160㎡ 豪华横厅", base_x + 5000, base_y - 1000, 400)

if __name__ == "__main__":
    generator = SmartLayoutGenerator()
    
    # For demo, generate all three to show capability
    # In real use, this would take an arg
    print("DEMO MODE: Generating complete set.")
    generator.generate(90)
    generator.generate(120)
    generator.generate(160)
