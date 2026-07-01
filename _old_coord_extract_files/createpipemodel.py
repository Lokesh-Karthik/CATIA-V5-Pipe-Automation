import win32com.client
import os
import json
import math
import tempfile
import subprocess
import time
from datetime import datetime

# ══════════════════════════════════════════════════════════════════════════════
#  PIPE BUILDER — Interactive Prompt → Dump File → CATIA
#  Flow:
#    1. Ask user for pipe definition (nodes, segments, diameters)
#    2. Save everything to a JSON dump file (reusable / auditable)
#    3. Build the pipe in CATIA: spline centerline + per-segment profiles + rib
# ══════════════════════════════════════════════════════════════════════════════

DUMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pipe_dumps")


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def hr(char="─", width=60):
    print(char * width)

def ask_float(prompt, min_val=None, max_val=None):
    while True:
        try:
            val = float(input(f"  {prompt}: ").strip())
            if min_val is not None and val < min_val:
                print(f"    ✗ Must be ≥ {min_val}. Try again.")
                continue
            if max_val is not None and val > max_val:
                print(f"    ✗ Must be ≤ {max_val}. Try again.")
                continue
            return val
        except ValueError:
            print("    ✗ Enter a number.")

def ask_int(prompt, min_val=1, max_val=None):
    while True:
        try:
            val = int(input(f"  {prompt}: ").strip())
            if val < min_val:
                print(f"    ✗ Must be ≥ {min_val}. Try again.")
                continue
            if max_val is not None and val > max_val:
                print(f"    ✗ Must be ≤ {max_val}. Try again.")
                continue
            return val
        except ValueError:
            print("    ✗ Enter a whole number.")

def ask_choice(prompt, choices):
    choices_lower = [c.lower() for c in choices]
    display = " / ".join(f"[{c}]" for c in choices)
    while True:
        val = input(f"  {prompt} {display}: ").strip().lower()
        if val in choices_lower:
            return val
        print(f"    ✗ Enter one of: {', '.join(choices)}.")

def ask_str(prompt, default=None):
    hint = f" (default: {default})" if default else ""
    val = input(f"  {prompt}{hint}: ").strip()
    if not val and default:
        return default
    return val


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 1 — INTERACTIVE INPUT
# ─────────────────────────────────────────────────────────────────────────────

def collect_pipe_data():
    hr("═")
    print("  CATIA PIPE BUILDER — Interactive Definition")
    hr("═")
    print()

    # ── Pipe name ──────────────────────────────────────────────────────────────
    part_name = ask_str("Pipe / Part name", default="Custom_Pipe")

    # ── Global defaults (user can override per segment) ────────────────────────
    print()
    hr()
    print("  DEFAULT CROSS-SECTION  (can be overridden per segment)")
    hr()
    default_outer = ask_float("Default outer diameter (mm)", min_val=1.0)
    default_inner = ask_float("Default inner diameter (mm)", min_val=0.1,
                               max_val=default_outer - 0.5)

    # ── Number of bends / nodes ────────────────────────────────────────────────
    print()
    hr()
    print("  ROUTE DEFINITION")
    print("  A pipe is defined by N points joined by (N-1) segments.")
    print("  Each segment is either  'straight'  or  'curved'.")
    hr()
    num_points = ask_int("Number of route points (min 2)", min_val=2)
    num_segments = num_points - 1

    # ── Collect nodes ──────────────────────────────────────────────────────────
    nodes = []
    print()
    print(f"  Enter coordinates for each of the {num_points} route points (mm):")
    for i in range(num_points):
        print()
        hr("-")
        print(f"  Point {i + 1} of {num_points}")
        hr("-")
        x = ask_float(f"  X")
        y = ask_float(f"  Y")
        z = ask_float(f"  Z")
        nodes.append({"point": i + 1, "x": x, "y": y, "z": z})

    # ── Collect segment properties ─────────────────────────────────────────────
    segments = []
    print()
    print(f"  Now define each of the {num_segments} segment(s) between your points:")

    for i in range(num_segments):
        p_start = nodes[i]
        p_end   = nodes[i + 1]
        seg_len = math.sqrt(
            (p_end["x"] - p_start["x"]) ** 2 +
            (p_end["y"] - p_start["y"]) ** 2 +
            (p_end["z"] - p_start["z"]) ** 2
        )
        print()
        hr("-")
        print(f"  Segment {i + 1}: Point {i+1} → Point {i+2}   "
              f"(chord length ≈ {seg_len:.1f} mm)")
        hr("-")

        seg_type = ask_choice("Segment type?", ["straight", "curved"])

        bend_radius = None
        if seg_type == "curved":
            min_bend = round(default_outer / 2.0 + 0.1, 2)
            print(f"    ℹ Bend radius must be > half the outer diameter "
                  f"({default_outer/2:.1f} mm).")
            bend_radius = ask_float("Bend radius (mm)", min_val=min_bend)

        # Per-segment diameter override?
        override = ask_choice("Override cross-section for this segment?", ["yes", "no"])
        if override == "yes":
            outer = ask_float("  Outer diameter (mm)", min_val=1.0)
            inner = ask_float("  Inner diameter (mm)", min_val=0.1,
                               max_val=outer - 0.5)
        else:
            outer = default_outer
            inner = default_inner

        segments.append({
            "segment":      i + 1,
            "from_point":   i + 1,
            "to_point":     i + 2,
            "type":         seg_type,
            "bend_radius":  bend_radius,
            "outer_dia":    outer,
            "inner_dia":    inner,
            "chord_length": round(seg_len, 4),
        })

    # ── Summary ────────────────────────────────────────────────────────────────
    print()
    hr("═")
    print("  SUMMARY")
    hr("═")
    print(f"  Part name   : {part_name}")
    print(f"  Route points: {num_points}")
    print(f"  Segments    : {num_segments}")
    for seg in segments:
        br = f", bend_r={seg['bend_radius']} mm" if seg["bend_radius"] else ""
        print(f"    Seg {seg['segment']}: {seg['type']:8s}  "
              f"OD={seg['outer_dia']} / ID={seg['inner_dia']}{br}")
    print()

    confirm = ask_choice("Proceed and build in CATIA?", ["yes", "no"])
    if confirm == "no":
        print("  Aborted by user.")
        return None

    # ── Attach per-node diameters (needed for CATIA profile placement) ─────────
    # Node diameter = the segment that STARTS from that node; last node uses last seg.
    for idx, node in enumerate(nodes):
        seg = segments[idx] if idx < num_segments else segments[-1]
        node["outer_dia"] = seg["outer_dia"]
        node["inner_dia"] = seg["inner_dia"]

    pipe_def = {
        "meta": {
            "created":      datetime.now().isoformat(timespec="seconds"),
            "part_name":    part_name,
            "num_points":   num_points,
            "num_segments": num_segments,
            "default_outer_dia": default_outer,
            "default_inner_dia": default_inner,
        },
        "nodes":    nodes,
        "segments": segments,
    }

    return pipe_def


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 2 — DUMP TO JSON
# ─────────────────────────────────────────────────────────────────────────────

def save_dump(pipe_def):
    os.makedirs(DUMP_DIR, exist_ok=True)
    ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
    name  = pipe_def["meta"]["part_name"].replace(" ", "_")
    fname = f"{name}_{ts}.json"
    fpath = os.path.join(DUMP_DIR, fname)
    with open(fpath, "w") as f:
        json.dump(pipe_def, f, indent=2)
    print(f"  ✔ Pipe definition saved → {fpath}")
    return fpath


def load_dump(fpath):
    with open(fpath) as f:
        return json.load(f)


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 3 — CONNECT TO CATIA
# ─────────────────────────────────────────────────────────────────────────────

def connect_to_catia(part_name):
    print()
    hr("═")
    print("  CATIA CONNECTION")
    hr("═")
    print("  Connecting to CATIA...")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        catia.Documents.Add("Part")
        part_doc = catia.ActiveDocument
        part     = part_doc.Part

        try:
            part_doc.Product.PartNumber = part_name
        except:
            pass

        hb = part.HybridBodies
        wireframe_set = hb.Add()
        wireframe_set.Name = "Pipe_Wireframe"

        print(f"  ✔ Connected. New part: {part_name}")
        return catia, part, wireframe_set
    except Exception as e:
        print(f"  ✗ CATIA connection failed: {e}")
        return None, None, None


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 4 — CENTERLINE SPINE (POLYLINE)
#  Uses Polyline to guarantee straight legs with localized corner radii.
# ─────────────────────────────────────────────────────────────────────────────

def create_spine(part, wireframe_set, nodes, segments):
    print()
    hr()
    print("  STEP 1 / 4 — Building Polyline centerline...")
    hr()
    hs = part.HybridShapeFactory

    # We use a Polyline to guarantee straight legs between nodes
    polyline = hs.AddNewPolyline()

    pt_refs = []
    for i, node in enumerate(nodes):
        # 1. Create the 3D Point
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"
        wireframe_set.AppendHybridShape(pt)
        
        # 2. Add it to the polyline (CATIA indexing is 1-based)
        pt_ref = part.CreateReferenceFromObject(pt)
        polyline.InsertElement(pt_ref, i + 1)
        pt_refs.append(pt_ref)
        
        print(f"    Node {node['point']}: ({node['x']}, {node['y']}, {node['z']})")

    # 3. Apply the bend radii strictly to the corners (middle points)
    #    We skip the first and last points because they are the start/end of the pipe
    for i in range(1, len(nodes) - 1):
        # In our JSON, the segment leaving the node holds the curve data
        seg = segments[i]
        if seg["type"] == "curved" and seg["bend_radius"]:
            # Set the radius at this specific corner vertex
            polyline.SetRadius(i + 1, seg["bend_radius"])
            print(f"    Applied {seg['bend_radius']}mm bend radius at Node {i + 1}")

    polyline.Name = "Pipe_Centerline"
    wireframe_set.AppendHybridShape(polyline)
    part.Update()
    
    print(f"  ✔ Centerline built with {len(nodes)} nodes.")
    return polyline, pt_refs

# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 5 — NORMAL PLANES AT EACH NODE
# ─────────────────────────────────────────────────────────────────────────────

def create_planes(part, wireframe_set, spline, pt_refs):
    print()
    hr()
    print("  STEP 2 / 4 — Creating normal planes at each node...")
    hr()
    hs       = part.HybridShapeFactory
    spline_r = part.CreateReferenceFromObject(spline)
    plane_refs = []

    for idx, pt_ref in enumerate(pt_refs):
        plane = hs.AddNewPlaneNormal(spline_r, pt_ref)
        plane.Name = f"Profile_Plane_{idx + 1}"
        wireframe_set.AppendHybridShape(plane)
        plane_refs.append(part.CreateReferenceFromObject(plane))
        print(f"    Plane {idx + 1} at node {idx + 1}")

    part.Update()
    print(f"  ✔ {len(plane_refs)} planes created.")
    return plane_refs


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 6 — CROSS-SECTION SKETCH (annular profile)
#  Uses the FIRST node's diameters for the master sweep profile.
#  Variable diameter would require a multi-section solid (future upgrade).
# ─────────────────────────────────────────────────────────────────────────────

def create_master_profile(part, plane_refs, nodes):
    print()
    hr()
    print("  STEP 3 / 4 — Drawing hollow annular cross-section...")
    hr()
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(plane_refs[0])
    sketch.Name = "Pipe_Master_Profile"

    fd = sketch.OpenEdition()
    outer_r = nodes[0]["outer_dia"] / 2.0
    inner_r = nodes[0]["inner_dia"] / 2.0
    fd.CreateClosedCircle(0.0, 0.0, outer_r)
    fd.CreateClosedCircle(0.0, 0.0, inner_r)
    sketch.CloseEdition()
    part.Update()

    print(f"  ✔ Profile: OD={nodes[0]['outer_dia']} mm / "
          f"ID={nodes[0]['inner_dia']} mm")
    return sketch


# ─────────────────────────────────────────────────────────────────────────────
#  PHASE 7 — SOLID RIB (VBScript bridge — same proven pattern as original)
# ─────────────────────────────────────────────────────────────────────────────

def create_solid_rib(pipe_name):
    print()
    hr()
    print("  STEP 4 / 4 — Sweeping profile → 3-D solid Rib (VBScript)...")
    hr()
    try:
        vbs  = 'Set CATIA = GetObject(, "CATIA.Application")\n'
        vbs += "Dim part1\n"
        vbs += "Set part1 = CATIA.ActiveDocument.Part\n"
        vbs += "Dim sf\n"
        vbs += "Set sf = part1.ShapeFactory\n"
        vbs += "part1.InWorkObject = part1.MainBody\n"

        vbs += "Dim sk1\n"
        vbs += 'Set sk1 = part1.MainBody.Sketches.Item("Pipe_Master_Profile")\n'
        vbs += "Dim refProfile\n"
        vbs += "Set refProfile = part1.CreateReferenceFromObject(sk1)\n"

        vbs += "Dim wf\n"
        vbs += 'Set wf = part1.HybridBodies.Item("Pipe_Wireframe")\n'
        vbs += "Dim spl\n"
        vbs += 'Set spl = wf.HybridShapes.Item("Pipe_Centerline")\n'
        vbs += "Dim refCurve\n"
        vbs += "Set refCurve = part1.CreateReferenceFromObject(spl)\n"

        vbs += "Dim rib1\n"
        vbs += "Set rib1 = sf.AddNewRibFromRef(refProfile, refCurve)\n"
        vbs += f'rib1.Name = "{pipe_name}_3D_Pipe"\n'
        vbs += "part1.Update\n"
        vbs += 'WScript.Echo ">>> 3D PIPE COMPLETE! <<<"\n'

        temp_dir = tempfile.gettempdir()
        vbs_path = os.path.join(temp_dir, "build_pipe_rib.vbs")
        with open(vbs_path, "w", encoding="ascii") as f:
            f.write(vbs)

        time.sleep(1)
        result = subprocess.run(
            ["cscript", "//Nologo", vbs_path],
            capture_output=True, text=True
        )
        print("  VBScript output:", result.stdout.strip())
        if result.stderr:
            print("  VBScript error:", result.stderr.strip())
        os.remove(vbs_path)

    except Exception as e:
        print(f"  ✗ VBScript failed: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    hr("═")
    print("  CATIA PIPE BUILDER  v1.0")
    print("  Prompted input → JSON dump → CATIA solid pipe")
    hr("═")
    print()

    # ── Option: load existing dump or define new ───────────────────────────────
    mode = ask_choice("Start mode?  New pipe  or  Load saved dump?",
                      ["new", "load"])

    if mode == "load":
        dump_path = ask_str("Full path to dump JSON file").strip('"')
        if not os.path.exists(dump_path):
            print(f"  ✗ File not found: {dump_path}")
            return
        pipe_def = load_dump(dump_path)
        print(f"  ✔ Loaded: {dump_path}")
    else:
        pipe_def = collect_pipe_data()
        if pipe_def is None:
            return
        dump_path = save_dump(pipe_def)

    nodes    = pipe_def["nodes"]
    segments = pipe_def["segments"]
    meta     = pipe_def["meta"]
    part_name = meta["part_name"]

    # ── Print re-loaded summary ────────────────────────────────────────────────
    print()
    hr()
    print(f"  Building: {part_name}")
    print(f"  Points  : {meta['num_points']}   Segments: {meta['num_segments']}")
    hr()

    # ── CATIA pipeline ─────────────────────────────────────────────────────────
    catia, part, wireframe_set = connect_to_catia(part_name)
    if part is None:
        return

    spline, pt_refs   = create_spine(part, wireframe_set, nodes, segments)
    if not spline:
        print("  ✗ Spine failed. Aborting.")
        return

    plane_refs        = create_planes(part, wireframe_set, spline, pt_refs)
    if not plane_refs:
        print("  ✗ Planes failed. Aborting.")
        return

    sketch            = create_master_profile(part, plane_refs, nodes)
    if not sketch:
        print("  ✗ Profile sketch failed. Aborting.")
        return

    create_solid_rib(part_name)

    hr("═")
    print(f"  DONE.  Dump saved to: {dump_path}")
    hr("═")


if __name__ == "__main__":
    main()
