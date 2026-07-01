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
#  CATIA COORD EXTRACTOR  v1.1
#
#  Reads an ALREADY-OPEN CATIA part and produces a piper2.py-compatible
#  JSON dump that can be loaded directly via piper2's "load" mode.
#
#  EXTRACTION MODES:
#    Mode 1 — Named Entities   : parts built by hose_automation / piper2
#    Mode 2 — HybridShape Scan : any part that has points in open bodies
#    Mode 3 — Edge Analysis    : closed solid PartBody / imported dumb solid
#                                (Reads circular seams to find centerline COGs)
#    Mode 4 — Diagnostic       : Scans selected topology to verify measurables
# ══════════════════════════════════════════════════════════════════════════════

# Folder where piper2.py saves (and loads) its JSON dumps
DUMP_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipe_dumps")

# Path to piper2.py (same folder as this script)
PIPER2_PY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "piper2.py")

# Wireframe body names and sketch names created by our scripts
SCRIPT_WIREFRAME_NAMES = ["Hose_Wireframe", "Pipe_Wireframe"]
SCRIPT_SKETCH_NAMES    = ["Hose_Master_Profile", "Pipe_Master_Profile"]


# ─────────────────────────────────────────────────────────────────────────────
#  SMALL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def hr(char="─", width=62):
    print(char * width)


def ask_float(prompt, default=None, min_val=None):
    hint = f" [{default}]" if default is not None else ""
    while True:
        raw = input(f"  {prompt}{hint}: ").strip()
        if not raw and default is not None:
            return float(default)
        try:
            val = float(raw)
            if min_val is not None and val < min_val:
                print(f"    ✗ Must be ≥ {min_val}")
                continue
            return val
        except ValueError:
            print("    ✗ Enter a number.")


# ─────────────────────────────────────────────────────────────────────────────
#  CONNECT
# ─────────────────────────────────────────────────────────────────────────────

def connect_to_catia():
    print("  Attaching to CATIA (must already be open with a part loaded)...")
    try:
        catia    = win32com.client.GetActiveObject("CATIA.Application")
        part_doc = catia.ActiveDocument
        part     = part_doc.Part

        try:
            part_name = part_doc.Product.PartNumber
        except Exception:
            part_name = os.path.splitext(part_doc.Name)[0]

        try:
            part_file = part_doc.FullName
        except Exception:
            part_file = "Unknown"

        print(f"  ✔ Connected  →  {part_name}  ({part_file})")
        return catia, part_doc, part, part_name, part_file

    except Exception as e:
        print(f"  ✗ Connection failed: {e}")
        print("    Open a .CATPart in CATIA first, then re-run this script.")
        return None, None, None, None, None


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
#  MODE 1  —  NAMED ENTITY EXTRACTION
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
            print(f"  ✔ Wireframe body found : '{name}'")
            break
        except Exception: pass

    if wireframe_body is None:
        print("  ℹ Standard wireframe name not found — scanning all HybridBodies...")
        for i in range(1, hbs.Count + 1):
            hb = hbs.Item(i)
            shapes = hb.HybridShapes
            for j in range(1, shapes.Count + 1):
                if shapes.Item(j).Name.startswith("Node_"):
                    wireframe_body = hb
                    print(f"  ✔ Found body with Node_* shapes: '{hb.Name}'")
                    break
            if wireframe_body: break

    if wireframe_body is None:
        print("  ✗ No suitable wireframe body found. Try Mode 2 or Mode 3.")
        return None, None, None

    shapes = wireframe_body.HybridShapes
    node_shapes = [shapes.Item(i) for i in range(1, shapes.Count + 1) if shapes.Item(i).Name.startswith("Node_")]

    def _node_sort_key(s):
        tail = s.Name.rsplit("_", 1)[-1]
        return int(tail) if tail.isdigit() else 9999

    node_shapes.sort(key=_node_sort_key)

    if not node_shapes:
        print("  ✗ No Node_* shapes inside the wireframe body.")
        return None, None, None

    nodes = []
    for s in node_shapes:
        x, y, z = get_point_coords(s)
        if x is None:
            print(f"    ⚠ Could not read coords for '{s.Name}' — skipping.")
            continue
        idx = len(nodes) + 1
        nodes.append({
            "point": idx, "x": round(x, 4), "y": round(y, 4), "z": round(z, 4),
            "outer_dia": 0.0, "inner_dia": 0.0,
        })
        print(f"    Node {idx:3d}: ({x:10.4f}, {y:10.4f}, {z:10.4f})")

    outer_dia, inner_dia = _read_sketch_diameters(part, SCRIPT_SKETCH_NAMES)
    return nodes, outer_dia, inner_dia


def _read_sketch_diameters(part, sketch_name_list):
    sketches = part.MainBody.Sketches
    for sk_name in sketch_name_list:
        try:
            geom = sketches.Item(sk_name).GeometricElements
            radii = []
            for i in range(1, geom.Count + 1):
                try:
                    r = geom.Item(i).Radius
                    if r > 0.0: radii.append(round(r * 2.0, 4))
                except: pass
            if len(radii) >= 2:
                radii.sort(reverse=True)
                print(f"  ✔ Profile sketch '{sk_name}': OD={radii[0]} mm  /  ID={radii[1]} mm")
                return radii[0], radii[1]
            elif len(radii) == 1:
                print(f"  ✔ Profile sketch '{sk_name}': OD={radii[0]} mm  (no inner circle)")
                return radii[0], round(radii[0] * 0.8, 4)
        except: pass
    print("  ℹ No profile sketch found — diameters will be entered manually.")
    return None, None


# ─────────────────────────────────────────────────────────────────────────────
#  MODE 2  —  GENERIC HYBRIDSHAPE SCAN
# ─────────────────────────────────────────────────────────────────────────────

def mode2_hybrid_scan(part):
    hr()
    print("  MODE 2 — Generic HybridShape point scan")
    hr()

    hbs = part.HybridBodies
    all_pts = []
    for i in range(1, hbs.Count + 1):
        hb = hbs.Item(i)
        shapes = hb.HybridShapes
        for j in range(1, shapes.Count + 1):
            s = shapes.Item(j)
            x, y, z = get_point_coords(s)
            if x is not None:
                all_pts.append({
                    "body": hb.Name, "shape": s.Name, "x": round(x, 4), "y": round(y, 4), "z": round(z, 4)
                })

    if not all_pts:
        print("  ✗ No readable point shapes found.")
        return None, None, None

    print(f"\n  Found {len(all_pts)} readable point shape(s):\n")
    print(f"  {'#':>3}  {'Body':<22} {'Shape':<22} {'X':>10} {'Y':>10} {'Z':>10}")
    hr("-")
    for idx, p in enumerate(all_pts):
        print(f"  {idx+1:>3}  {p['body']:<22} {p['shape']:<22} {p['x']:>10.3f} {p['y']:>10.3f} {p['z']:>10.3f}")
    hr("-")

    raw = input("\n  Route points (e.g. '1 3 5' or ENTER for all): ").strip()
    if raw:
        try:
            indices = [int(t) - 1 for t in raw.split()]
            selected = [all_pts[i] for i in indices if 0 <= i < len(all_pts)]
        except:
            selected = all_pts
    else:
        selected = all_pts

    if len(selected) < 2:
        print("  ✗ Need at least 2 points to define a pipe route.")
        return None, None, None

    nodes = []
    for p in selected:
        idx = len(nodes) + 1
        nodes.append({"point": idx, "x": p["x"], "y": p["y"], "z": p["z"], "outer_dia": 0.0, "inner_dia": 0.0})
    return nodes, None, None


# ─────────────────────────────────────────────────────────────────────────────
#  MODE 3  —  SOLID BODY EDGE ANALYSIS
#
#  Dumps ALL circular edge measurements (radius, COG) to a temp CSV via VBScript.
#  The COGs of these junction seams provide the exact nodes of our pipe!
# ─────────────────────────────────────────────────────────────────────────────

def mode3_edge_analysis():
    hr()
    print("  MODE 3 — Solid body edge analysis  (VBScript + SPAWorkbench)")
    hr()

    temp_dir = tempfile.gettempdir()
    csv_path = os.path.join(temp_dir, "catia_edge_extract.csv")
    vbs_path = os.path.join(temp_dir, "catia_edge_scan.vbs")
    csv_vbs  = csv_path.replace("\\", "\\\\")

    # Build VBScript targeting Topology.Edge
    vbs  = 'Set CATIA   = GetObject(, "CATIA.Application")\n'
    vbs += 'Set partDoc = CATIA.ActiveDocument\n'
    vbs += 'Set sel     = partDoc.Selection\n'
    vbs += 'Set spa     = partDoc.GetWorkbench("SPAWorkbench")\n'
    vbs += 'sel.Clear\n'
    vbs += 'sel.Search "Topology.Edge,all"\n'
    vbs += f'Set fso = CreateObject("Scripting.FileSystemObject")\n'
    vbs += f'Set ts  = fso.CreateTextFile("{csv_vbs}", True)\n'
    vbs += 'ts.WriteLine "idx,radius,cog_x,cog_y,cog_z"\n'
    vbs += 'Dim i\n'
    vbs += 'For i = 1 To sel.Count\n'
    vbs += '    On Error Resume Next\n'
    vbs += '    Err.Clear\n'
    vbs += '    Dim ref, meas, r\n'
    vbs += '    Set ref  = sel.Item(i).Reference\n'
    vbs += '    Set meas = spa.GetMeasurable(ref)\n'
    vbs += '    r = meas.Radius\n'
    vbs += '    If Err.Number = 0 And r > 0 Then\n'
    vbs += '        Dim cog(2)\n'
    vbs += '        meas.GetCOG cog\n'
    vbs += '        ts.WriteLine i & "," & r & "," & cog(0) & "," & cog(1) & "," & cog(2)\n'
    vbs += '    End If\n'
    vbs += '    On Error GoTo 0\n'
    vbs += 'Next\n'
    vbs += 'ts.Close\n'
    vbs += 'sel.Clear\n'
    vbs += 'WScript.Echo "Edge scan complete."\n'

    with open(vbs_path, "w", encoding="ascii") as f:
        f.write(vbs)

    print("  Running SPAWorkbench edge scan via VBScript...")
    time.sleep(0.5)

    result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
    if result.stderr:
        print(f"  VBScript error: {result.stderr.strip()}")

    try: os.remove(vbs_path)
    except: pass

    if not os.path.exists(csv_path):
        print(f"  ✗ Output CSV not found: {csv_path}")
        return None, None, None

    circular_edges = []
    try:
        with open(csv_path, newline="") as f:
            for row in csv.DictReader(f):
                try:
                    circular_edges.append({
                        "radius": float(row["radius"]),
                        "cog":  [float(row["cog_x"]), float(row["cog_y"]), float(row["cog_z"])],
                    })
                except: pass
    finally:
        try: os.remove(csv_path)
        except: pass

    if not circular_edges:
        print("  ✗ No circular edges detected. Not a pipe body, or body is empty.")
        return None, None, None

    print(f"  ✔ {len(circular_edges)} circular edge(s) found.")

    unique_radii = sorted(set(round(c["radius"], 1) for c in circular_edges), reverse=True)
    print(f"  Detected radii (mm) : {unique_radii}")

    outer_r   = unique_radii[0]
    inner_r   = unique_radii[1] if len(unique_radii) > 1 else outer_r * 0.8
    outer_dia = round(outer_r * 2.0, 3)
    inner_dia = round(inner_r * 2.0, 3)
    print(f"  Inferred  OD = {outer_dia} mm   ID = {inner_dia} mm")

    # Filter to outer-radius edges only (the defining boundary)
    outer_edges = [c for c in circular_edges if abs(c["radius"] - outer_r) < 0.5]

    # Sort along the primary axis of progression
    def _spread(axis_idx):
        vals = [c["cog"][axis_idx] for c in outer_edges]
        return max(vals) - min(vals) if vals else 0.0

    primary_axis = max(range(3), key=_spread)
    axis_labels  = ["X", "Y", "Z"]
    print(f"  Sorting nodes along primary axis : {axis_labels[primary_axis]}")
    outer_edges.sort(key=lambda c: c["cog"][primary_axis])

    # Cluster overlapping seam edges into single nodes
    nodes = _cluster_to_nodes(outer_edges, tolerance=outer_dia * 0.5)

    if not nodes:
        return None, None, None

    print(f"  ✔ Reconstructed {len(nodes)} route node(s):")
    for n in nodes:
        print(f"    Node {n['point']:3d}: ({n['x']:.3f}, {n['y']:.3f}, {n['z']:.3f})")

    return nodes, outer_dia, inner_dia


def _cluster_to_nodes(edges, tolerance):
    used  = [False] * len(edges)
    nodes = []

    for i, edge in enumerate(edges):
        if used[i]: continue
        cluster = [edge["cog"]]
        used[i] = True
        for j in range(i + 1, len(edges)):
            if not used[j] and math.dist(edge["cog"], edges[j]["cog"]) < tolerance:
                cluster.append(edges[j]["cog"])
                used[j] = True

        cx = sum(p[0] for p in cluster) / len(cluster)
        cy = sum(p[1] for p in cluster) / len(cluster)
        cz = sum(p[2] for p in cluster) / len(cluster)

        nodes.append({
            "point": len(nodes) + 1, "x": round(cx, 4), "y": round(cy, 4), "z": round(cz, 4),
            "outer_dia": 0.0, "inner_dia": 0.0,
        })
    return nodes


# ─────────────────────────────────────────────────────────────────────────────
#  MODE 4  —  DIAGNOSTIC INSPECTOR
# ─────────────────────────────────────────────────────────────────────────────

def mode4_diagnostic(part):
    hr()
    print("  MODE 4 — Diagnostic Inspector")
    hr()
    catia = win32com.client.Dispatch("CATIA.Application")
    doc   = catia.ActiveDocument
    sel   = doc.Selection
    spa   = doc.GetWorkbench("SPAWorkbench")

    print("  Scanning active part topologies...")
    
    for topo in ["Face", "Edge", "Vertex"]:
        sel.Clear()
        sel.Search(f"Topology.{topo},all")
        count = sel.Count
        print(f"\n  Found {count} {topo}(s).")
        
        if count > 0:
            print(f"  Testing SPAWorkbench measurables on a sample of up to 10 {topo}s:")
            for i in range(1, min(10, count) + 1):
                ref = sel.Item(i).Reference
                try:
                    meas = spa.GetMeasurable(ref)
                    props = []
                    try: props.append(f"Radius: {meas.Radius:.2f}")
                    except: pass
                    try: props.append(f"Length: {meas.Length:.2f}")
                    except: pass
                    try: props.append(f"Area: {meas.Area:.2f}")
                    except: pass
                    
                    out = ", ".join(props) if props else "No standard measurables (Radius/Length/Area)"
                    print(f"    Item {i:2d}: {out}")
                except Exception:
                    print(f"    Item {i:2d}: GetMeasurable failed.")
    
    sel.Clear()
    print("\n  Diagnostic complete.")


# ─────────────────────────────────────────────────────────────────────────────
#  BUILD SEGMENTS
# ─────────────────────────────────────────────────────────────────────────────

def build_segments(nodes, outer_dia, inner_dia, straight_thresh_deg=3.0):
    segments = []
    n = len(nodes)

    for i in range(n - 1):
        p1 = nodes[i]
        p2 = nodes[i + 1]
        dx, dy, dz = p2["x"]-p1["x"], p2["y"]-p1["y"], p2["z"]-p1["z"]
        chord = math.sqrt(dx**2 + dy**2 + dz**2)

        seg_type    = "straight"
        bend_radius = None

        if i > 0:
            pp = nodes[i - 1]
            vp = (p1["x"]-pp["x"], p1["y"]-pp["y"], p1["z"]-pp["z"])
            vc = (dx, dy, dz)
            lp = math.sqrt(sum(a**2 for a in vp))
            lc = chord

            if lp > 0 and lc > 0:
                dot = max(-1.0, min(1.0, sum(a*b for a, b in zip(vp, vc)) / (lp * lc)))
                angle_deg = math.degrees(math.acos(dot))

                if angle_deg > straight_thresh_deg:
                    seg_type = "curved"
                    half = math.radians(angle_deg / 2.0)
                    if half > 0:
                        bend_radius = round(chord / (2.0 * math.sin(half)), 2)
                        min_bend = round(outer_dia / 2.0 + 0.1, 2)
                        bend_radius = max(bend_radius, min_bend)

        od  = p1.get("outer_dia") or outer_dia or 25.0
        id_ = p1.get("inner_dia") or inner_dia or 20.0

        segments.append({
            "segment": i + 1, "from_point": i + 1, "to_point": i + 2, "type": seg_type,
            "bend_radius": bend_radius, "outer_dia": od, "inner_dia": id_, "chord_length": round(chord, 4),
        })
    return segments


# ─────────────────────────────────────────────────────────────────────────────
#  SAVE AS PIPER2-COMPATIBLE JSON DUMP
# ─────────────────────────────────────────────────────────────────────────────

def save_piper2_dump(nodes, segments, part_name, outer_dia, inner_dia, source_file):
    os.makedirs(DUMP_DIR, exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in part_name)
    path = os.path.join(DUMP_DIR, f"{safe}_extracted_{ts}.json")

    pipe_def = {
        "meta": {
            "created": datetime.now().isoformat(timespec="seconds"), "part_name": part_name,
            "num_points": len(nodes), "num_segments": len(segments),
            "default_outer_dia": outer_dia, "default_inner_dia": inner_dia,
            "extracted_from": source_file, "extractor_version": "1.1",
        },
        "nodes": nodes, "segments": segments,
    }

    with open(path, "w") as f:
        json.dump(pipe_def, f, indent=2)

    print(f"\n  ✔ Dump saved  →  {path}")
    return path


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    hr("═")
    print("  CATIA COORD EXTRACTOR  v1.1")
    print("  Active CATIA part  →  piper2.py-compatible JSON dump")
    hr("═")
    print()

    catia, part_doc, part, part_name, part_file = connect_to_catia()
    if part is None: return

    while True:
        print()
        print("  SELECT EXTRACTION MODE")
        hr()
        print("  1  Named Entities    — parts built by hose_automation / piper2")
        print("  2  HybridShape Scan  — any part with points in open bodies")
        print("  3  Edge Analysis     — solid PartBody or imported dumb solid")
        print("                         (SPAWorkbench circular seam edge detection)")
        print("  4  Diagnostic        — inspect available topology and measurements")
        hr()
        choice = input("  Mode [1 / 2 / 3 / 4]: ").strip()
        
        if choice == "4":
            mode4_diagnostic(part)
            input("\n  Press Enter to return to menu...")
            continue
        elif choice in ("1", "2", "3"):
            break
        print("  ✗ Enter 1, 2, 3, or 4.")

    print()
    if choice == "1":
        nodes, outer_dia, inner_dia = mode1_named_entities(part)
    elif choice == "2":
        nodes, outer_dia, inner_dia = mode2_hybrid_scan(part)
    else:
        nodes, outer_dia, inner_dia = mode3_edge_analysis()

    if not nodes or len(nodes) < 2:
        print("\n  ✗ Fewer than 2 nodes extracted — cannot build a pipe definition.")
        return

    if outer_dia is None or inner_dia is None:
        print()
        hr()
        print("  CROSS-SECTION  (could not be read from the part automatically)")
        hr()
        outer_dia = ask_float("Outer diameter (mm)", default=25.0, min_val=1.0)
        inner_dia = ask_float("Inner diameter (mm)", default=round(outer_dia * 0.8, 1), min_val=0.1)

    for node in nodes:
        if not node.get("outer_dia"): node["outer_dia"] = outer_dia
        if not node.get("inner_dia"): node["inner_dia"] = inner_dia

    segments = build_segments(nodes, outer_dia, inner_dia)

    print()
    hr("═")
    print("  EXTRACTION SUMMARY")
    hr("═")
    print(f"  Part      : {part_name}")
    print(f"  Nodes     : {len(nodes)}")
    print(f"  Segments  : {len(segments)}")
    print(f"  OD / ID   : {outer_dia} / {inner_dia} mm")
    print()
    for n in nodes:
        print(f"    Node {n['point']:3d}:  ({n['x']:10.4f},  {n['y']:10.4f},  {n['z']:10.4f})")
    print()
    for seg in segments:
        br = f"  bend_r = {seg['bend_radius']} mm" if seg["bend_radius"] else ""
        print(f"    Seg {seg['segment']}:  {seg['type']:8s}  chord = {seg['chord_length']:8.2f} mm{br}")
    hr("═")

    custom = input(f"\n  Output part name [{part_name}]: ").strip()
    if custom: part_name = custom

    dump_path = save_piper2_dump(nodes, segments, part_name, outer_dia, inner_dia, part_file)

    print()
    go = input("  Launch piper2.py with this dump right now? [y/n]: ").strip().lower()
    if go == "y":
        if os.path.exists(PIPER2_PY):
            subprocess.Popen(["python", PIPER2_PY], creationflags=subprocess.CREATE_NEW_CONSOLE)
            print("\n  piper2.py launched in a new console window.")
            print("  → When prompted for start mode, choose  [load]  and paste:")
            print(f"    {dump_path}")
        else:
            print(f"  ✗ piper2.py not found at {PIPER2_PY}")
    else:
        print("\n  To rebuild this pipe in CATIA later:")
        print("    python piper2.py   →   choose [load]   →   paste path:")
        print(f"    {dump_path}")

    hr("═")
    print("  DONE!")
    hr("═")

if __name__ == "__main__":
    main()