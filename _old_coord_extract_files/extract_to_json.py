import win32com.client
import os
import json
import math
import tempfile
import subprocess
from datetime import datetime

# Where to save the output so the GUI can find it
DUMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipe_dumps")

def hr(char="─", width=60):
    print(char * width)

def extract_coords_via_vbs():
    """Extracts X, Y, Z coordinates securely via VBScript."""
    vbs_code = """
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    Set doc = CATIA.ActiveDocument
    Set sel = doc.Selection
    Set spa = doc.GetWorkbench("SPAWorkbench")
    
    For i = 1 To sel.Count
        Set ref = sel.Item(i).Reference
        Set meas = spa.GetMeasurable(ref)
        
        Dim c(2)
        Err.Clear
        meas.GetCenter c
        
        ' Fallback if it is a corner point instead of a circular edge
        If Err.Number <> 0 Then
            Err.Clear
            meas.GetPoint c
        End If
        
        If Err.Number = 0 Then
            x = Round(c(0), 2)
            y = Round(c(1), 2)
            z = Round(c(2), 2)
            WScript.Echo "NODE|" & i & "|" & x & "|" & y & "|" & z
        Else
            WScript.Echo "ERROR_NODE|" & i & "|" & Err.Description
        End If
    Next
    """
    temp_dir = tempfile.gettempdir()
    vbs_path = os.path.join(temp_dir, "extract_coords.vbs")
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(vbs_code)
        
    result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
    try: os.remove(vbs_path)
    except: pass
        
    return result.stdout.strip().split("\n")

def extract_radii_via_vbs():
    """Extracts the radius of selected curved edges securely via VBScript."""
    vbs_code = """
    On Error Resume Next
    Set CATIA = GetObject(, "CATIA.Application")
    Set doc = CATIA.ActiveDocument
    Set sel = doc.Selection
    Set spa = doc.GetWorkbench("SPAWorkbench")
    
    For i = 1 To sel.Count
        Set ref = sel.Item(i).Reference
        Set meas = spa.GetMeasurable(ref)
        
        Err.Clear
        r = meas.Radius
        
        If Err.Number = 0 Then
            WScript.Echo "RAD|" & i & "|" & Round(r, 2)
        Else
            WScript.Echo "ERROR_RAD|" & i & "|" & Err.Description
        End If
    Next
    """
    temp_dir = tempfile.gettempdir()
    vbs_path = os.path.join(temp_dir, "extract_rad.vbs")
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(vbs_code)
        
    result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
    try: os.remove(vbs_path)
    except: pass
        
    return result.stdout.strip().split("\n")


def main():
    hr("═")
    print("  CATIA REVERSE-ENGINEER TO JSON EXTRACTOR")
    hr("═")
    
    # 1. Connect to CATIA
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        doc = catia.ActiveDocument
        part = doc.Part
        selection = doc.Selection
        spa = doc.GetWorkbench("SPAWorkbench")
        print("[SETUP] Connected to CATIA successfully.")
    except Exception as e:
        print(f"[ERROR] Could not connect to CATIA: {e}")
        return

    part_name = input("\nWhat do you want to name this extracted pipe? (e.g., Extracted_Line_1): ").strip()
    if not part_name: part_name = "Extracted_Pipe"

    # =========================================================================
    # STEP 1: ROUTE EXTRACTION
    # =========================================================================
    print("\n[STEP 1] Route Extraction")
    print("         Select the CIRCULAR EDGES at the start,")
    print("         bends, and end of your pipe.")
    print("         (Hold CTRL to select multiple in order).")
    input("         Press ENTER here in the console once selected...")

    if selection.Count < 2:
        print("[ERROR] You must select at least 2 edges to define a pipe.")
        return

    nodes = []
    print("\nExtracting center coordinates...")
    output_lines = extract_coords_via_vbs()
    
    for line in output_lines:
        line = line.strip()
        if not line: continue
        parts = line.split("|")
        if parts[0] == "NODE":
            idx = int(parts[1])
            x, y, z = float(parts[2]), float(parts[3]), float(parts[4])
            print(f"  Node {idx} Center: X={x}, Y={y}, Z={z}")
            nodes.append({"point": len(nodes) + 1, "x": x, "y": y, "z": z})
        elif parts[0] == "ERROR_NODE":
            print(f"  [ERROR] Failed to measure selection {parts[1]}. Make sure it is a valid edge.")

    selection.Clear()
    if len(nodes) < 2:
        print("\n[ERROR] Not enough valid coordinates extracted. Aborting.")
        return

    # =========================================================================
    # STEP 2: OUTER DIAMETER
    # =========================================================================
    print("\n[STEP 2] Outer Diameter")
    print("         Select a straight CYLINDRICAL FACE on the PartBody.")
    input("         Press ENTER here in the console once selected...")
    
    outer_dia = 30.0  # Fallback
    try:
        if selection.Count > 0:
            ref = selection.Item(1).Reference
            measurable = spa.GetMeasurable(ref)
            radius = measurable.Radius
            outer_dia = round(radius * 2.0, 2)
            print(f"  Extracted Outer Diameter: {outer_dia} mm")
        else:
            print(f"  [WARNING] Nothing selected. Defaulting to {outer_dia} mm.")
    except Exception as e:
        print(f"  [WARNING] Could not read cylinder radius. Defaulting to {outer_dia} mm. ({e})")
    
    inner_dia = max(0.1, outer_dia - 4.0)
    selection.Clear()

    # =========================================================================
    # STEP 3: BEND RADII
    # =========================================================================
    print("\n[STEP 3] Bend Radii")
    print("         How would you like to define the bend radii?")
    print("         [1] Select curved arcs/edges in CATIA to DETECT them")
    print("         [2] Enter a single default radius manually")
    choice = input("         Choice [1 or 2]: ").strip()
    
    default_bend_radius = 30.0
    detected_radii = []
    
    if choice == "1":
        print("\n         Select the curved ARCS of the bends in order.")
        print("         (Hold CTRL to select multiple).")
        input("         Press ENTER here in the console once selected...")
        
        if selection.Count > 0:
            print("\nExtracting bend radii...")
            output_lines = extract_radii_via_vbs()
            raw_radii = []
            for line in output_lines:
                line = line.strip()
                if not line: continue
                parts = line.split("|")
                if parts[0] == "RAD":
                    raw_radii.append(float(parts[2]))
                elif parts[0] == "ERROR_RAD":
                    print(f"  [ERROR] Could not measure radius for selection {parts[1]}.")
            selection.Clear()
            
            if raw_radii:
                print("\n         Where are these selected arcs located on the pipe?")
                print("         [C] Centerline (Neutral axis)")
                print("         [O] Outer seam (Outside of the bend)")
                print("         [I] Inner seam (Inside of the bend)")
                loc = input("         Choice [C/O/I] (Default C): ").strip().upper()
                
                tube_rad = outer_dia / 2.0
                for i, r in enumerate(raw_radii):
                    if loc == 'O':
                        adjusted_r = r - tube_rad
                    elif loc == 'I':
                        adjusted_r = r + tube_rad
                    else:
                        adjusted_r = r
                        
                    if adjusted_r < 0.1: adjusted_r = 0.1
                        
                    detected_radii.append(round(adjusted_r, 2))
                    print(f"  -> Bend {i+1} Raw: {r}mm | True Centerline Radius: {round(adjusted_r, 2)}mm")
        else:
            print("  [WARNING] Nothing selected.")
            
        # Fallback if they selected fewer radii than there are bends
        if not detected_radii or len(detected_radii) < (len(nodes) - 2):
            if detected_radii: 
                print(f"\n  [WARNING] You extracted {len(detected_radii)} radii, but there are {len(nodes)-2} bends.")
            b_rad_input = input(f"\n         Enter a fallback bend radius for any missing corners (mm) [Default 30.0]: ").strip()
            default_bend_radius = float(b_rad_input) if b_rad_input else 30.0
    else:
        b_rad_input = input(f"\n         Enter the default bend radius for corners (mm) [Default 30.0]: ").strip()
        default_bend_radius = float(b_rad_input) if b_rad_input else 30.0

    # =========================================================================
    # JSON BUILDER
    # =========================================================================
    segments = []
    bend_idx = 0
    for i in range(len(nodes) - 1):
        p_start = nodes[i]
        p_end = nodes[i + 1]
        
        chord = math.sqrt((p_end["x"] - p_start["x"])**2 + 
                          (p_end["y"] - p_start["y"])**2 + 
                          (p_end["z"] - p_start["z"])**2)
        
        # Map detected bend radii sequentially to the corners
        if i > 0:
            if bend_idx < len(detected_radii):
                current_bend = detected_radii[bend_idx]
                bend_idx += 1
            else:
                current_bend = default_bend_radius
        else:
            current_bend = default_bend_radius
            
        segments.append({
            "segment": i + 1,
            "from_point": i + 1,
            "to_point": i + 2,
            "type": "curved",
            "bend_radius": current_bend,
            "outer_dia": outer_dia,
            "inner_dia": inner_dia,
            "chord_length": round(chord, 4)
        })

    for node in nodes:
        node["outer_dia"] = outer_dia
        node["inner_dia"] = inner_dia

    pipe_def = {
        "meta": {
            "created": datetime.now().isoformat(timespec="seconds"),
            "part_name": part_name,
            "num_points": len(nodes),
            "num_segments": len(segments),
            "default_outer_dia": outer_dia,
            "default_inner_dia": inner_dia,
        },
        "nodes": nodes,
        "segments": segments
    }

    os.makedirs(DUMP_DIR, exist_ok=True)
    fname = f"{part_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    fpath = os.path.join(DUMP_DIR, fname)
    
    with open(fpath, "w") as f:
        json.dump(pipe_def, f, indent=2)

    hr("═")
    print(f"[SUCCESS] JSON generated perfectly for the Pipe Builder GUI!")
    print(f"File saved to: {fpath}")
    hr("═")


if __name__ == "__main__":
    main()