import win32com.client
import os
import json
import math
import tempfile
import subprocess
import time
from datetime import datetime

# =============================================================================
#  THE PHOENIX PIPELINE: Autonomous Vector Routing & 3D Sweep
# =============================================================================

DUMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipe_dumps")

def hr(char="─", width=70):
    print(char * width)

# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 1: THE EXTRACTOR (VBScript & Vector Math)
# ─────────────────────────────────────────────────────────────────────────────

def get_deflection_angle(p1, p2, p3):
    """Calculates the angle (in degrees) between three 3D coordinates."""
    v1_x = p2["x"] - p1["x"]
    v1_y = p2["y"] - p1["y"]
    v1_z = p2["z"] - p1["z"]
    
    v2_x = p3["x"] - p2["x"]
    v2_y = p3["y"] - p2["y"]
    v2_z = p3["z"] - p2["z"]
    
    mag1 = math.sqrt(v1_x**2 + v1_y**2 + v1_z**2)
    mag2 = math.sqrt(v2_x**2 + v2_y**2 + v2_z**2)
    
    if mag1 == 0 or mag2 == 0:
        return 0.0 
        
    dot_product = (v1_x * v2_x) + (v1_y * v2_y) + (v1_z * v2_z)
    cos_theta = dot_product / (mag1 * mag2)
    cos_theta = max(-1.0, min(1.0, cos_theta)) 
    
    return math.degrees(math.acos(cos_theta))

def orthogonal_snap(nodes):
    """
    Forces all pipe segments to align perfectly with the X, Y, or Z axes.
    This eliminates click-inaccuracies and guarantees perfect 90-degree bends.
    """
    snapped = [nodes[0].copy()]
    
    for i in range(1, len(nodes)):
        prev = snapped[-1]
        curr = nodes[i].copy()
        
        dx = curr["x"] - prev["x"]
        dy = curr["y"] - prev["y"]
        dz = curr["z"] - prev["z"]
        
        # Skip if the points are identical (accidental double-click)
        if abs(dx) < 0.1 and abs(dy) < 0.1 and abs(dz) < 0.1:
            continue
            
        # Find the dominant axis of movement
        abs_deltas = {'x': abs(dx), 'y': abs(dy), 'z': abs(dz)}
        dominant_axis = max(abs_deltas, key=abs_deltas.get)
        
        # Lock the other two coordinates to the previous point to create a pure 90-degree turn
        if dominant_axis == 'x':
            curr["y"] = prev["y"]
            curr["z"] = prev["z"]
        elif dominant_axis == 'y':
            curr["x"] = prev["x"]
            curr["z"] = prev["z"]
        else:
            curr["x"] = prev["x"]
            curr["y"] = prev["y"]
            
        snapped.append(curr)
        
    # Re-number the points sequentially in case duplicates were removed
    for i, n in enumerate(snapped):
        n["point"] = i + 1
        
    return snapped

def extract_coords_via_vbs():
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


def extract_pipe_data(catia):
    doc       = catia.ActiveDocument
    selection = doc.Selection
    spa       = doc.GetWorkbench("SPAWorkbench")

    part_name = input("\nName the new parametric pipe (e.g., Line_1): ").strip()
    if not part_name: part_name = "Parametric_Pipe"

    # ── Step 1: Extract Nodes ──────────────────────────────────────────────────
    print("\n[STEP 1] ROUTE EXTRACTION")
    print("  Select the CIRCULAR EDGES/POINTS at the start, bends, and end.")
    print("  (Hold CTRL to select multiple in physical order).")
    input("  Press ENTER in this console once selected...")

    if selection.Count < 2:
        print("[ERROR] You must select at least 2 points to define a pipe.")
        return None

    nodes = []
    print("  Extracting pure mathematical coordinates...")
    output_lines = extract_coords_via_vbs()

    for line in output_lines:
        line = line.strip()
        if not line: continue
        parts = line.split("|")
        if parts[0] == "NODE":
            x, y, z = float(parts[2]), float(parts[3]), float(parts[4])
            nodes.append({"point": len(nodes) + 1, "x": x, "y": y, "z": z, "bend_radius": 0.0})
            print(f"    ✔ Node {len(nodes)}: X={x}, Y={y}, Z={z}")

    selection.Clear()
    if len(nodes) < 2: return None

    # ── Step 2: Global Pipe Size ───────────────────────────────────────────────
    print("\n[STEP 2] PIPE DIAMETER")
    print("  Select a straight CYLINDRICAL FACE on the pipe body.")
    input("  Press ENTER in this console once selected...")

    outer_dia = 30.0
    try:
        if selection.Count > 0:
            ref    = selection.Item(1).Reference
            radius = spa.GetMeasurable(ref).Radius
            outer_dia = round(radius * 2.0, 2)
            print(f"  ✔ Extracted Outer Diameter: {outer_dia} mm")
        else:
            outer_dia = float(input("  Nothing selected. Enter Outer Diameter manually (mm): ").strip())
    except:
        outer_dia = float(input("  Could not read radius. Enter Outer Diameter manually (mm): ").strip())

    inner_dia = max(0.1, outer_dia - 4.0)
    selection.Clear()

    # ── Step 3: Autonomous Corner Routing (Vector Math) ────────────────────────
    print("\n[STEP 3] AUTONOMOUS ROUTING DEFINITION")
    
    # NEW: Ask the user if they want to square off the pipe
    force_90 = input("  Force perfect 90° bends? (Snaps clicks to global X,Y,Z axes) (y/n) [default: y]: ").strip().lower()
    if force_90 != 'n':
        nodes = orthogonal_snap(nodes)
        print("    ✔ Route orthogonally squared. All corners locked to 90°.")
    
    default_br_input = input("  Enter the default bend radius for curved corners (mm) [default 30]: ").strip()
    global_bend_radius = float(default_br_input) if default_br_input else 30.0

    if len(nodes) > 2:
        print("\n  Analyzing vector collinearity...")
        for i in range(1, len(nodes) - 1):
            angle = get_deflection_angle(nodes[i-1], nodes[i], nodes[i+1])
            
            # If the angle is less than 1 degree, it's a straight line
            if angle < 1.0:
                nodes[i]["bend_radius"] = 0.0  # Sharp straight joint
                print(f"    ✔ Node {i+1}: Straight Line detected (Angle: {angle:.2f}°)")
            else:
                nodes[i]["bend_radius"] = global_bend_radius # Sweeping elbow
                print(f"    ✔ Node {i+1}: Curve detected     (Angle: {angle:.2f}°)")

    # Apply diameters to nodes for the JSON
    for node in nodes:
        node["outer_dia"] = outer_dia
        node["inner_dia"] = inner_dia

    pipe_def = {
        "meta": {
            "part_name":    part_name,
            "num_points":   len(nodes)
        },
        "nodes": nodes
    }

    # Save to JSON
    os.makedirs(DUMP_DIR, exist_ok=True)
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    fpath = os.path.join(DUMP_DIR, f"{part_name}_{ts}.json")
    with open(fpath, "w") as f: json.dump(pipe_def, f, indent=2)
    
    print(f"\n  [DNA EXTRACTED] Blueprint saved: {fpath}")
    return pipe_def


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 2: THE BUILDER (Solid Sweep)
# ─────────────────────────────────────────────────────────────────────────────

def build_parametric_pipe(catia, pipe_def):
    nodes     = pipe_def["nodes"]
    part_name = pipe_def["meta"]["part_name"]

    print("\n[STEP 4] RESURRECTION (Generating Parametric Solid)")

    catia.Documents.Add("Part")
    part_doc = catia.ActiveDocument
    part     = part_doc.Part
    try: part_doc.Product.PartNumber = part_name
    except: pass

    hb            = part.HybridBodies
    wireframe_set = hb.Add()
    wireframe_set.Name = "Pipe_Wireframe"
    hs = part.HybridShapeFactory

    # 1. Build Polyline
    print("  -> Drafting Polyline Skeleton...")
    polyline = hs.AddNewPolyline()
    pt_refs  = []
    for i, node in enumerate(nodes):
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        wireframe_set.AppendHybridShape(pt)
        pt_ref = part.CreateReferenceFromObject(pt)
        polyline.InsertElement(pt_ref, i + 1)
        pt_refs.append(pt_ref)

    # 2. Apply Bend Radii (Based purely on the Vector Math calculation!)
    for i in range(1, len(nodes) - 1):
        if nodes[i]["bend_radius"] > 0:
            polyline.SetRadius(i + 1, nodes[i]["bend_radius"])

    polyline.Name = "Pipe_Centerline"
    wireframe_set.AppendHybridShape(polyline)
    part.Update()

    # 3. Create Normal Plane
    print("  -> Orienting Profile Plane...")
    plane = hs.AddNewPlaneNormal(part.CreateReferenceFromObject(polyline), pt_refs[0])
    wireframe_set.AppendHybridShape(plane)
    part.Update()

    # 4. Sketch Profile
    print("  -> Drawing Hollow Annular Profile...")
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(part.CreateReferenceFromObject(plane))
    sketch.Name = "Pipe_Master_Profile"
    fd = sketch.OpenEdition()
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["outer_dia"] / 2.0)
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["inner_dia"] / 2.0)
    sketch.CloseEdition()
    part.Update()

    # 5. Trojan VBScript Rib
    print("  -> Sweeping 3D Solid (Trojan Protocol)...")
    vbs  = 'Set CATIA = GetObject(, "CATIA.Application")\n'
    vbs += "Set part1 = CATIA.ActiveDocument.Part\n"
    vbs += "Set sf = part1.ShapeFactory\n"
    vbs += "part1.InWorkObject = part1.MainBody\n"
    vbs += 'Set sk1 = part1.MainBody.Sketches.Item("Pipe_Master_Profile")\n'
    vbs += 'Set spl = part1.HybridBodies.Item("Pipe_Wireframe").HybridShapes.Item("Pipe_Centerline")\n'
    vbs += "Set rib1 = sf.AddNewRib(part1.CreateReferenceFromObject(sk1), part1.CreateReferenceFromObject(spl))\n"
    vbs += f'rib1.Name = "{part_name}_3D_Solid"\n'
    vbs += "part1.Update\n"
    vbs += 'WScript.Echo ">>> PIPE COMPLETE!"\n'

    temp_dir = tempfile.gettempdir()
    vbs_path = os.path.join(temp_dir, "build_pipe_rib.vbs")
    with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)

    time.sleep(1)
    subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True)
    try: os.remove(vbs_path)
    except: pass

    hr("═")
    print(f"  >>> SUCCESS: {part_name} FULLY RECONSTRUCTED! <<<")
    hr("═")

def main():
    hr("═")
    print("  THE PHOENIX PIPELINE: Autonomous Vector Routing")
    hr("═")

    try:
        catia = win32com.client.Dispatch("CATIA.Application")
    except:
        print("[ERROR] Could not connect to CATIA. Is it running?")
        return

    pipe_def = extract_pipe_data(catia)
    if pipe_def:
        build_parametric_pipe(catia, pipe_def)

if __name__ == "__main__":
    main()