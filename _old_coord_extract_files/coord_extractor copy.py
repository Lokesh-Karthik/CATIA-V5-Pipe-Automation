import win32com.client
import os
import json
import math
import csv
import tempfile
import subprocess
import time
from datetime import datetime

# ══════════════════════════════════════════════════════════════════════════════
#  CATIA COORD EXTRACTOR  v1.5 (Diagnostic Edge Engine)
#
#  Bypasses face topology entirely. Uses raw edge diagnostics + Nearest
#  Neighbor pathing to generate a perfect piper2.py JSON dump.
# ══════════════════════════════════════════════════════════════════════════════

DUMP_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipe_dumps")
PIPER2_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "piper2.py")

SCRIPT_WIREFRAME_NAMES = ["Hose_Wireframe", "Pipe_Wireframe"]
SCRIPT_SKETCH_NAMES    = ["Hose_Master_Profile", "Pipe_Master_Profile"]

def hr(char="─", width=62):
    print(char * width)

def ask_float(prompt, default=None, min_val=None):
    hint = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"  {prompt}{hint}: ").strip()
        if not raw and default is not None: return float(default)
        try:
            val = float(raw)
            if min_val is not None and val < min_val:
                print(f"    ✗ Must be ≥ {min_val}")
                continue
            return val
        except ValueError:
            print("    ✗ Enter a number.")

def get_dist(p1, p2):
    return math.sqrt((p1[0]-p2[0])**2 + (p1[1]-p2[1])**2 + (p1[2]-p2[2])**2)

# ─────────────────────────────────────────────────────────────────────────────
#  CONNECT
# ─────────────────────────────────────────────────────────────────────────────

def connect_to_catia():
    print("  Attaching to CATIA (must already be open with a part loaded)...")
    try:
        catia = win32com.client.GetActiveObject("CATIA.Application")
        part_doc = catia.ActiveDocument
        part = part_doc.Part
        try: part_name = part_doc.Product.PartNumber
        except: part_name = os.path.splitext(part_doc.Name)[0]
        print(f"  ✔ Connected  →  {part_name}")
        return catia, part_doc, part, part_name
    except:
        print("  ✗ Connection failed. Open a .CATPart in CATIA first.")
        return None, None, None, None

def get_point_coords(shape):
    try: return float(shape.X.Value), float(shape.Y.Value), float(shape.Z.Value)
    except: pass
    try: return float(shape.X), float(shape.Y), float(shape.Z)
    except: pass
    try:
        pos = shape.GetPosition()
        return float(pos[0]), float(pos[1]), float(pos[2])
    except: pass
    return None, None, None

# ─────────────────────────────────────────────────────────────────────────────
#  MODE 1 & 2 (Retained for Compatibility)
# ─────────────────────────────────────────────────────────────────────────────

def mode1_named_entities(part):
    hr()
    print("  MODE 1 — Named entity extraction")
    hr()
    wireframe_body = None
    hbs = part.HybridBodies
    for name in SCRIPT_WIREFRAME_NAMES:
        try:
            wireframe_body = hbs.Item(name)
            break
        except: pass
    if wireframe_body is None:
        for i in range(1, hbs.Count + 1):
            hb = hbs.Item(i)
            for j in range(1, hb.HybridShapes.Count + 1):
                if hb.HybridShapes.Item(j).Name.startswith("Node_"):
                    wireframe_body = hb
                    break
            if wireframe_body: break
    if wireframe_body is None:
        print("  ✗ No suitable wireframe body found.")
        return None, None, None

    node_shapes = [wireframe_body.HybridShapes.Item(i) for i in range(1, wireframe_body.HybridShapes.Count + 1) if wireframe_body.HybridShapes.Item(i).Name.startswith("Node_")]
    node_shapes.sort(key=lambda s: int(s.Name.rsplit("_", 1)[-1]) if s.Name.rsplit("_", 1)[-1].isdigit() else 9999)

    nodes = []
    for s in node_shapes:
        x, y, z = get_point_coords(s)
        if x is not None:
            nodes.append({"point": len(nodes) + 1, "x": round(x, 1), "y": round(y, 1), "z": round(z, 1)})
    outer_dia, inner_dia = _read_sketch_diameters(part, SCRIPT_SKETCH_NAMES)
    return nodes, outer_dia, inner_dia

def _read_sketch_diameters(part, sketch_name_list):
    sketches = part.MainBody.Sketches
    for sk_name in sketch_name_list:
        try:
            geom = sketches.Item(sk_name).GeometricElements
            radii = [round(geom.Item(i).Radius * 2.0, 1) for i in range(1, geom.Count + 1) if geom.Item(i).Radius > 0.0]
            if len(radii) >= 2:
                radii.sort(reverse=True)
                return radii[0], radii[1]
            elif len(radii) == 1:
                return radii[0], round(radii[0] * 0.8, 1)
        except: pass
    return None, None

def mode2_hybrid_scan(part):
    hr()
    print("  MODE 2 — Generic HybridShape point scan")
    hr()
    all_pts = []
    for i in range(1, part.HybridBodies.Count + 1):
        shapes = part.HybridBodies.Item(i).HybridShapes
        for j in range(1, shapes.Count + 1):
            s = shapes.Item(j)
            x, y, z = get_point_coords(s)
            if x is not None:
                all_pts.append({"shape": s.Name, "x": round(x, 1), "y": round(y, 1), "z": round(z, 1)})
    if not all_pts: return None, None, None
    for idx, p in enumerate(all_pts):
        print(f"  {idx+1:>3}  {p['shape']:<22} X:{p['x']:>8.1f} Y:{p['y']:>8.1f} Z:{p['z']:>8.1f}")
    raw = input("\n  Route points (e.g. '1 3 5' or ENTER for all): ").strip()
    selected = [all_pts[int(t)-1] for t in raw.split() if 0 <= int(t)-1 < len(all_pts)] if raw else all_pts
    nodes = [{"point": i+1, "x": p["x"], "y": p["y"], "z": p["z"]} for i, p in enumerate(selected)]
    return nodes, None, None

# ─────────────────────────────────────────────────────────────────────────────
#  MODE 3  —  DIAGNOSTIC EDGE ENGINE (RAW CLOUD -> JSON)
# ─────────────────────────────────────────────────────────────────────────────

def mode3_diagnostic_edge_engine():
    hr()
    print("  MODE 3 — Diagnostic Edge Engine (Raw Cloud -> JSON)")
    hr()

    temp_dir = tempfile.gettempdir()
    csv_path = os.path.join(temp_dir, "catia_diag_extract.csv")
    vbs_path = os.path.join(temp_dir, "catia_diag_scan.vbs")
    csv_vbs  = csv_path.replace("\\", "\\\\")

    # Uses the exact robust edge scan you saw working in Mode 4
    vbs  = 'Set CATIA = GetObject(, "CATIA.Application")\n'
    vbs += 'Set sel = CATIA.ActiveDocument.Selection\n'
    vbs += 'Set spa = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")\n'
    vbs += 'sel.Clear\nsel.Search "Topology.Edge,all"\n'
    vbs += f'Set ts = CreateObject("Scripting.FileSystemObject").CreateTextFile("{csv_vbs}", True)\n'
    vbs += 'ts.WriteLine "radius,cog_x,cog_y,cog_z"\n'
    vbs += 'For i = 1 To sel.Count\n'
    vbs += '    On Error Resume Next\n'
    vbs += '    Set meas = spa.GetMeasurable(sel.Item(i).Reference)\n'
    vbs += '    r = meas.Radius\n'
    vbs += '    If Err.Number = 0 And r > 0 Then\n'
    vbs += '        Dim cog(2)\nmeas.GetCOG cog\n'
    vbs += '        ts.WriteLine r & "," & cog(0) & "," & cog(1) & "," & cog(2)\n'
    vbs += '    End If\n    On Error GoTo 0\nNext\nts.Close\nsel.Clear\n'

    with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)

    print("  Running raw edge diagnostics via VBScript...")
    time.sleep(0.5)
    subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True)

    try: os.remove(vbs_path)
    except: pass

    if not os.path.exists(csv_path):
        return None, None, None

    edges = []
    with open(csv_path, newline="") as f:
        for row in csv.DictReader(f):
            try: edges.append({"radius": float(row["radius"]), "cog": [float(row["cog_x"]), float(row["cog_y"]), float(row["cog_z"])]})
            except: pass

    try: os.remove(csv_path)
    except: pass

    if not edges:
        print("  ✗ No circular edges detected.")
        return None, None, None

    print(f"  ✔ {len(edges)} circular edge(s) found.")

    # Find the dominant outer radius
    unique_radii = sorted(set(round(c["radius"], 1) for c in edges), reverse=True)
    outer_r = unique_radii[0]
    outer_dia = round(outer_r * 2.0, 1)
    inner_dia = round(unique_radii[1] * 2.0, 1) if len(unique_radii) > 1 else round(outer_dia * 0.8, 1)

    print(f"  ✔ Inferred OD = {outer_dia} mm, ID = {inner_dia} mm")
    
    # Isolate edges that belong to the outer diameter
    outer_edges = [c["cog"] for c in edges if abs(c["radius"] - outer_r) < 0.5]

    # Cluster identical COGs (overlapping seam edges) into unique Nodes
    unique_cogs = []
    for cog in outer_edges:
        if not any(get_dist(cog, u) < (outer_dia * 0.5) for u in unique_cogs):
            unique_cogs.append(cog)

    print(f"  ✔ Clustered into {len(unique_cogs)} unique structural nodes.")

    if len(unique_cogs) < 2:
        return None, None, None

    print("  Routing path via Nearest Neighbor sorting...")

    # 1. Find the two nodes furthest apart (this guarantees we find at least one absolute endpoint of the pipe)
    max_d = -1
    start_node = unique_cogs[0]
    for n1 in unique_cogs:
        for n2 in unique_cogs:
            d = get_dist(n1, n2)
            if d > max_d:
                max_d = d
                start_node = n1

    # 2. Nearest Neighbor Sort (Trace the pipe from the starting end to the other end)
    ordered_points = [start_node]
    unvisited = [n for n in unique_cogs if n != start_node]

    while unvisited:
        current = ordered_points[-1]
        # Find the unvisited node closest to the current node
        nearest = min(unvisited, key=lambda n: get_dist(current, n))
        ordered_points.append(nearest)
        unvisited.remove(nearest)

    # 3. Format to piper2 JSON standard
    nodes = []
    for i, pt in enumerate(ordered_points):
        nodes.append({
            "point": i + 1, 
            "x": round(pt[0], 1), 
            "y": round(pt[1], 1), 
            "z": round(pt[2], 1)
        })

    return nodes, outer_dia, inner_dia

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD SEGMENTS
# ─────────────────────────────────────────────────────────────────────────────

def build_segments(nodes, outer_dia, inner_dia, straight_thresh_deg=3.0):
    segments = []
    for i in range(len(nodes) - 1):
        p1, p2 = nodes[i], nodes[i + 1]
        dx, dy, dz = p2["x"]-p1["x"], p2["y"]-p1["y"], p2["z"]-p1["z"]
        chord = math.sqrt(dx**2 + dy**2 + dz**2)

        seg_type, bend_radius = "straight", None

        if i > 0:
            pp = nodes[i - 1]
            vp = (p1["x"]-pp["x"], p1["y"]-pp["y"], p1["z"]-pp["z"])
            vc = (dx, dy, dz)
            lp, lc = math.sqrt(sum(a**2 for a in vp)), chord

            if lp > 0 and lc > 0:
                dot = max(-1.0, min(1.0, sum(a*b for a, b in zip(vp, vc)) / (lp * lc)))
                angle_deg = math.degrees(math.acos(dot))
                if angle_deg > straight_thresh_deg:
                    seg_type = "curved"
                    half = math.radians(angle_deg / 2.0)
                    if half > 0:
                        bend_radius = max(round(chord / (2.0 * math.sin(half)), 1), round(outer_dia / 2.0 + 0.1, 1))

        segments.append({
            "segment": i + 1, "from_point": i + 1, "to_point": i + 2, 
            "type": seg_type, "bend_radius": bend_radius, 
            "outer_dia": p1.get("outer_dia", outer_dia), "inner_dia": p1.get("inner_dia", inner_dia), 
            "chord_length": round(chord, 1)
        })
    return segments

# ─────────────────────────────────────────────────────────────────────────────
#  SAVE EXACT JSON TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────

def save_piper2_dump(nodes, segments, part_name, outer_dia, inner_dia):
    os.makedirs(DUMP_DIR, exist_ok=True)
    
    ts_iso = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    ts_file = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in part_name)
    path = os.path.join(DUMP_DIR, f"{safe_name}_{ts_file}.json")

    pipe_def = {
        "meta": {
            "created": ts_iso,
            "part_name": part_name,
            "num_points": len(nodes),
            "num_segments": len(segments),
            "default_outer_dia": outer_dia,
            "default_inner_dia": inner_dia
        },
        "nodes": nodes,
        "segments": segments
    }

    with open(path, "w") as f:
        json.dump(pipe_def, f, indent=2)

    print(f"\n  ✔ JSON Dump saved  →  {path}")
    return path

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    hr("═")
    print("  CATIA COORD EXTRACTOR  v1.5 (Diagnostic Edge Engine)")
    hr("═")
    print()

    catia, part_doc, part, part_name = connect_to_catia()
    if part is None: return

    while True:
        print("\n  SELECT EXTRACTION MODE")
        hr()
        print("  1  Named Entities    — parts built by hose_automation / piper2")
        print("  2  HybridShape Scan  — any part with points in open bodies")
        print("  3  Edge Analysis     — solid PartBody or imported dumb solid")
        hr()
        choice = input("  Mode [1 / 2 / 3]: ").strip()
        
        if choice in ("1", "2", "3"): 
            break
        print("  ✗ Enter 1, 2, or 3.")

    if choice == "1":
        nodes, outer_dia, inner_dia = mode1_named_entities(part)
    elif choice == "2":
        nodes, outer_dia, inner_dia = mode2_hybrid_scan(part)
    else:
        nodes, outer_dia, inner_dia = mode3_diagnostic_edge_engine()

    if not nodes or len(nodes) < 2:
        print("\n  ✗ Extraction failed.")
        return

    if outer_dia is None or inner_dia is None:
        print()
        hr()
        print("  CROSS-SECTION")
        hr()
        outer_dia = ask_float("Outer diameter (mm)", default=25.0, min_val=1.0)
        inner_dia = ask_float("Inner diameter (mm)", default=round(outer_dia * 0.8, 1), min_val=0.1)

    for node in nodes:
        node["outer_dia"] = outer_dia
        node["inner_dia"] = inner_dia

    segments = build_segments(nodes, outer_dia, inner_dia)

    print()
    hr("═")
    print("  NAME YOUR JSON EXPORT")
    hr("═")
    custom_name = input(f"  Enter custom part name (or press Enter to keep '{part_name}'): ").strip()
    if custom_name:
        part_name = custom_name

    dump_path = save_piper2_dump(nodes, segments, part_name, outer_dia, inner_dia)

    print()
    go = input("  Launch piper2.py with this dump right now? [y/n]: ").strip().lower()
    if go == "y" and os.path.exists(PIPER2_PY):
        subprocess.Popen(["python", PIPER2_PY], creationflags=subprocess.CREATE_NEW_CONSOLE)
        print("\n  piper2.py launched in a new console window.")
        print(f"  → Choose [load] and paste:\n    {dump_path}")
    
    hr("═")
    print("  DONE!")
    hr("═")

if __name__ == "__main__":
    main()