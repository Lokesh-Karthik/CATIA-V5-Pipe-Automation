"""
CATIA Pipe Automation Suite  v2.1
Fixes vs v2.0:
  • pythoncom.CoInitialize() per worker thread  → pipeline runs unlimited times
  • Per-tab _running flags                       → each tab is independent
  • Typed done-sentinel with tab-id              → safe queue protocol
  • Dassault / CATIA color scheme               → dark navy + signature orange
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import sys, os, json, math, time, tempfile, subprocess
from datetime import datetime


# ═══════════════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS  — Dassault Systèmes / CATIA V5 color language
#  References: CATIA 3DX dark shell, DS brand orange #FF6600, viewport navy
# ═══════════════════════════════════════════════════════════════════════════════
BG_DEEP    = "#131C27"   # viewport / deepest background
BG_PANEL   = "#1C2A3A"   # sidebar, card panels
BG_RAISED  = "#253546"   # inputs, table rows, raised cells
BG_BORDER  = "#324356"   # dividers, borders
ACCENT     = "#E8700A"   # Dassault signature orange
ACCENT_DIM = "#9E4B04"   # pressed / hover orange
ACCENT2    = "#1A7EC8"   # secondary DS blue (used sparingly)
SUCCESS    = "#3DB87F"   # green for success lines
WARNING    = "#F5C518"   # amber for warnings
ERROR      = "#E05252"   # red for errors
TEXT_HI    = "#EAF0F8"   # primary text
TEXT_MID   = "#7A90AA"   # secondary / labels
TEXT_DIM   = "#3D5060"   # disabled / dim

FNT_UI     = ("Segoe UI", 10)
FNT_UI_B   = ("Segoe UI", 10, "bold")
FNT_UI_L   = ("Segoe UI", 12, "bold")
FNT_TITLE  = ("Segoe UI", 17, "bold")
FNT_CODE   = ("Consolas", 9)
FNT_SMALL  = ("Segoe UI", 8)

# ── Typed done-sentinel carries the tab-id (1/2/3) so the poll knows which
#    tab just finished without risking a widget == string comparison.
class _Done:
    __slots__ = ("tab",)
    def __init__(self, tab): self.tab = tab


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — EXTRACT & REBUILD
# ═══════════════════════════════════════════════════════════════════════════════
def run_extract_rebuild(log, pipe_radius=15.0, straightness_tol=0.5, dedup_tol=5.0):
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed. Run on Windows with CATIA V5.", "error"); return

    log("[SETUP] Connecting to CATIA V5...", "info")
    try:
        catia         = win32com.client.Dispatch("CATIA.Application")
        part_document = catia.ActiveDocument
        part          = part_document.Part
        hs_factory    = part.HybridShapeFactory
        shape_factory = part.ShapeFactory
        spa           = part_document.GetWorkbench("SPAWorkbench")
        selection     = part_document.Selection
        hybrid_bodies = part.HybridBodies
    except Exception as e:
        log(f"[ERROR] Cannot connect to CATIA: {e}", "error")
        log("  Make sure CATIA V5 is open with a Part document active.", "warn"); return

    log("[SETUP] Connected successfully.", "success")

    # ── PHASE 1 ────────────────────────────────────────────────────────────────
    log("\n── PHASE 1: Extracting cylinder faces ──────────────────────────", "head")
    try:
        old = hybrid_bodies.Item("Extracted_Cylinders")
        selection.Clear(); selection.Add(old); selection.Delete()
        log("[Phase 1] Cleared old 'Extracted_Cylinders' set.", "dim")
    except: pass

    extraction_set = hybrid_bodies.Add(); extraction_set.Name = "Extracted_Cylinders"
    selection.Clear(); selection.Search("Topology.Face,all")
    total_faces = selection.Count
    log(f"[Phase 1] Found {total_faces} total faces in model.", "info")

    extracted_count = 0; all_extractions = []
    for i in range(1, total_faces + 1):
        item = selection.Item(i); shape_type = str(item.Type); face_ref = item.Reference
        if "ylinder" in shape_type:
            ext = hs_factory.AddNewExtract(face_ref); ext.Name = f"Cyl_Face_{extracted_count+1}"
            extraction_set.AppendHybridShape(ext); all_extractions.append(ext); extracted_count += 1

    if extracted_count == 0:
        log("[Phase 1] No cylinder type detected — falling back to ALL faces.", "warn")
        selection.Clear(); selection.Search("Topology.Face,all")
        for i in range(1, selection.Count + 1):
            face_ref = selection.Item(i).Reference
            ext = hs_factory.AddNewExtract(face_ref); ext.Name = f"Face_{i}"
            extraction_set.AppendHybridShape(ext); all_extractions.append(ext); extracted_count += 1

    try:
        part.Update(); log(f"[Phase 1] {extracted_count} faces in 'Extracted_Cylinders'.", "success")
    except Exception as e:
        log(f"[Phase 1] Update warning: {e}", "warn")

    if extracted_count == 0:
        log("[Phase 1] FAILED: No faces found.", "error"); return

    # ── PHASE 2 ────────────────────────────────────────────────────────────────
    log("\n── PHASE 2: Computing centerlines ──────────────────────────────", "head")
    try:
        old_cl = hybrid_bodies.Item("Extracted_Centerlines")
        selection.Clear(); selection.Add(old_cl); selection.Delete()
    except: pass

    centerline_set = hybrid_bodies.Add(); centerline_set.Name = "Extracted_Centerlines"
    faces = extraction_set.HybridShapes; all_lines = []

    for i in range(1, faces.Count + 1):
        current_face = faces.Item(i)
        selection.Clear(); selection.Add(current_face)
        try: selection.Search("Topology.Edge,sel")
        except: continue
        if selection.Count == 0: continue

        points_data = []
        for j in range(1, selection.Count + 1):
            try:
                raw_ref = selection.Item(j).Reference
                ee = hs_factory.AddNewExtract(raw_ref); centerline_set.AppendHybridShape(ee)
                pt = hs_factory.AddNewPointCenter(part.CreateReferenceFromObject(ee))
                centerline_set.AppendHybridShape(pt); part.UpdateObject(pt); points_data.append(pt)
            except: continue

        if len(points_data) < 2: continue
        max_dist = -1; best_p1 = best_p2 = None
        for p_a in points_data:
            meas_a = spa.GetMeasurable(part.CreateReferenceFromObject(p_a))
            for p_b in points_data:
                try:
                    d = meas_a.GetMinimumDistance(part.CreateReferenceFromObject(p_b))
                    if d > max_dist: max_dist = d; best_p1 = p_a; best_p2 = p_b
                except: pass

        if not best_p1 or max_dist <= 1.0: continue
        line = hs_factory.AddNewLinePtPt(part.CreateReferenceFromObject(best_p1),
                                          part.CreateReferenceFromObject(best_p2))
        centerline_set.AppendHybridShape(line); part.UpdateObject(line)
        midpt = hs_factory.AddNewPointOnCurveFromPercent(part.CreateReferenceFromObject(line), 0.5, True)
        centerline_set.AppendHybridShape(midpt); part.UpdateObject(midpt)

        face_ref = part.CreateReferenceFromObject(current_face)
        dist_mid = spa.GetMeasurable(part.CreateReferenceFromObject(midpt)).GetMinimumDistance(face_ref)
        dist_p1  = spa.GetMeasurable(part.CreateReferenceFromObject(best_p1)).GetMinimumDistance(face_ref)

        if abs(dist_mid - dist_p1) < straightness_tol:
            all_lines.append({"line": line, "p1": best_p1, "p2": best_p2, "midpoint": midpt})
        else:
            selection.Clear(); selection.Add(line); selection.Add(midpt)
            try: selection.Delete()
            except: pass

    log(f"[Phase 2] Generated {len(all_lines)} raw axis lines.", "info")
    if not all_lines:
        log("[Phase 2] FAILED: No straight lines found.", "error"); return

    log("[Phase 2] Deduplicating overlapping axis lines...", "info")
    unique_lines = []
    for ld in all_lines:
        meas = spa.GetMeasurable(part.CreateReferenceFromObject(ld["midpoint"])); is_dup = False
        for ul in unique_lines:
            try:
                if meas.GetMinimumDistance(part.CreateReferenceFromObject(ul["midpoint"])) < dedup_tol:
                    is_dup = True
                    selection.Clear()
                    for o in [ld["line"], ld["p1"], ld["p2"], ld["midpoint"]]: selection.Add(o)
                    try: selection.Delete()
                    except: pass
                    break
            except: pass
        if not is_dup: unique_lines.append(ld)

    for ul in unique_lines:
        selection.Clear(); selection.Add(ul["midpoint"])
        try: selection.Delete()
        except: pass

    log(f"[Phase 2] {len(unique_lines)} unique pipe segments after dedup.", "success")
    if len(unique_lines) < 2:
        log("[Phase 2] FAILED: Need at least 2 segments.", "error"); return

    log("[Phase 2] Sorting segments tip-to-tail...", "info")
    meas_ref = spa.GetMeasurable(part.CreateReferenceFromObject(unique_lines[0]["p1"]))
    max_d = -1; start_line = unique_lines[0]
    for ld in unique_lines:
        try:
            d = meas_ref.GetMinimumDistance(part.CreateReferenceFromObject(ld["line"]))
            if d > max_d: max_d = d; start_line = ld
        except: pass

    sorted_lines = [start_line]; unvisited = [ld for ld in unique_lines if ld is not start_line]
    active_tip = start_line["p1"]
    if unvisited:
        mp1 = spa.GetMeasurable(part.CreateReferenceFromObject(start_line["p1"]))
        mp2 = spa.GetMeasurable(part.CreateReferenceFromObject(start_line["p2"]))
        md1 = md2 = float("inf")
        for ld in unvisited:
            ref_l = part.CreateReferenceFromObject(ld["line"])
            try:
                d1 = mp1.GetMinimumDistance(ref_l); d2 = mp2.GetMinimumDistance(ref_l)
                if d1 < md1: md1 = d1
                if d2 < md2: md2 = d2
            except: pass
        active_tip = start_line["p2"] if md2 < md1 else start_line["p1"]

    while unvisited:
        meas_tip = spa.GetMeasurable(part.CreateReferenceFromObject(active_tip))
        best = None; bd = float("inf"); bp1d = bp2d = float("inf")
        for ld in unvisited:
            try:
                d1 = meas_tip.GetMinimumDistance(part.CreateReferenceFromObject(ld["p1"]))
                d2 = meas_tip.GetMinimumDistance(part.CreateReferenceFromObject(ld["p2"]))
                lm = min(d1, d2)
                if lm < bd: bd = lm; best = ld; bp1d = d1; bp2d = d2
            except: pass
        if not best: break
        sorted_lines.append(best); unvisited.remove(best)
        active_tip = best["p2"] if bp1d < bp2d else best["p1"]

    log(f"[Phase 2] Sorted {len(sorted_lines)} segments.", "success")

    # ── PHASE 3 ────────────────────────────────────────────────────────────────
    log("\n── PHASE 3: Bridging bends & creating solid Rib ────────────────", "head")
    connects = []
    for i in range(len(sorted_lines) - 1):
        l1 = sorted_lines[i]; l2 = sorted_lines[i + 1]
        p1a, p1b = l1["p1"], l1["p2"]; p2a, p2b = l2["p1"], l2["p2"]
        mp1a = spa.GetMeasurable(part.CreateReferenceFromObject(p1a))
        mp1b = spa.GetMeasurable(part.CreateReferenceFromObject(p1b))
        daa = mp1a.GetMinimumDistance(part.CreateReferenceFromObject(p2a))
        dab = mp1a.GetMinimumDistance(part.CreateReferenceFromObject(p2b))
        dba = mp1b.GetMinimumDistance(part.CreateReferenceFromObject(p2a))
        dbb = mp1b.GetMinimumDistance(part.CreateReferenceFromObject(p2b))
        min_d = min(daa, dab, dba, dbb)
        if   min_d == daa: pt1, pt2 = p1a, p2a
        elif min_d == dab: pt1, pt2 = p1a, p2b
        elif min_d == dba: pt1, pt2 = p1b, p2a
        else:              pt1, pt2 = p1b, p2b
        ori1 = 1 if pt1.Name == l1["p2"].Name else -1
        ori2 = 1 if pt2.Name == l2["p1"].Name else -1
        try:
            conn = hs_factory.AddNewConnect(
                part.CreateReferenceFromObject(l1["line"]), part.CreateReferenceFromObject(pt1), ori1, 1, 1.0,
                part.CreateReferenceFromObject(l2["line"]), part.CreateReferenceFromObject(pt2), ori2, 1, 1.0,
                False)
            conn.Name = f"Bend_Connect_{i+1}"; centerline_set.AppendHybridShape(conn); part.UpdateObject(conn)
            connects.append(conn); log(f"  Bend {i+1} connected (ori: {ori1}, {ori2})", "info")
        except Exception as e:
            log(f"  WARNING: Bend {i+1} failed: {e}", "warn")

    log(f"[Phase 3] {len(connects)} bend(s) bridged.", "success")

    join_elements = [sorted_lines[0]["line"]]
    for i in range(len(connects)):
        join_elements.append(connects[i]); join_elements.append(sorted_lines[i+1]["line"])

    if len(join_elements) < 2:
        log("[Phase 3] FAILED: Not enough elements.", "error"); return

    try:
        master_join = hs_factory.AddNewJoin(
            part.CreateReferenceFromObject(join_elements[0]),
            part.CreateReferenceFromObject(join_elements[1]))
        for elem in join_elements[2:]:
            master_join.AddElement(part.CreateReferenceFromObject(elem))
        master_join.Name = "Master_Centerline_Join"
        centerline_set.AppendHybridShape(master_join); part.UpdateObject(master_join)
        log("[Phase 3] Master centerline join created.", "success")
    except Exception as e:
        log(f"[Phase 3] FAILED during Join: {e}", "error"); return

    join_ref = part.CreateReferenceFromObject(master_join)
    try:
        pt_start = hs_factory.AddNewPointOnCurveFromPercent(join_ref, 0.0, True)
        centerline_set.AppendHybridShape(pt_start)
        plane = hs_factory.AddNewPlaneNormal(join_ref, part.CreateReferenceFromObject(pt_start))
        centerline_set.AppendHybridShape(plane)
        circle = hs_factory.AddNewCircleCtrRad(
            part.CreateReferenceFromObject(pt_start),
            part.CreateReferenceFromObject(plane), False, pipe_radius)
        circle.Name = "Pipe_Profile_Circle"
        centerline_set.AppendHybridShape(circle); part.Update()
        log(f"[Phase 3] Profile circle created (radius = {pipe_radius} mm).", "success")
    except Exception as e:
        log(f"[Phase 3] FAILED creating profile circle: {e}", "error"); return

    try:
        part.InWorkObject = part.MainBody
        shape_factory.AddNewRibFromRef(part.CreateReferenceFromObject(circle), join_ref)
        part.Update(); log("[Phase 3] >>> SOLID RIB GENERATED SUCCESSFULLY! <<<", "success")
    except Exception as e:
        log(f"[Phase 3] FAILED during Rib: {e}", "error"); return

    try:
        part_document.Save(); log("[DONE] Document saved.", "success")
    except Exception as e:
        log(f"[DONE] Save warning: {e}", "warn")

    log(f"\n{'='*55}", "head")
    log(f"  PIPELINE COMPLETE", "head")
    log(f"  Segments extracted : {extracted_count}", "head")
    log(f"  Unique axis lines  : {len(unique_lines)}", "head")
    log(f"  Bends connected    : {len(connects)}", "head")
    log(f"  Solid pipe         : Rib in PartBody", "head")
    log(f"{'='*55}", "head")


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — BUILD FROM COORDINATES
# ═══════════════════════════════════════════════════════════════════════════════
def run_build_from_coords(log, hose_data):
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed.", "error"); return

    log("[SETUP] Connecting to CATIA V5...", "info")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True; catia.Documents.Add("Part")
        part_doc = catia.ActiveDocument; part = part_doc.Part
        try: part_doc.Product.PartNumber = "AC_Hose_Automated"
        except: pass
        wf_set = part.HybridBodies.Add(); wf_set.Name = "Hose_Wireframe"
        log("[SETUP] New Part created: AC_Hose_Automated", "success")
    except Exception as e:
        log(f"[ERROR] CATIA connection failed: {e}", "error"); return

    log("\n── Building Spline Centerline ──────────────────────────────────", "head")
    hs = part.HybridShapeFactory
    spline = hs.AddNewSpline(); spline.SetSplineType(0)
    pt_refs = []
    for node in hose_data:
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"; wf_set.AppendHybridShape(pt)
        pt_ref = part.CreateReferenceFromObject(pt)
        spline.AddPoint(pt_ref); pt_refs.append(pt_ref)
        log(f"  Node {node['point']}: ({node['x']}, {node['y']}, {node['z']})", "info")
    spline.Name = "Hose_Centerline"; wf_set.AppendHybridShape(spline); part.Update()
    log(f"[OK] Spline created with {len(hose_data)} nodes.", "success")

    log("\n── Creating Normal Planes ──────────────────────────────────────", "head")
    spline_ref = part.CreateReferenceFromObject(spline); plane_refs = []
    for idx, pt_ref in enumerate(pt_refs):
        plane = hs.AddNewPlaneNormal(spline_ref, pt_ref)
        plane.Name = f"Profile_Plane_{idx+1}"; wf_set.AppendHybridShape(plane)
        plane_refs.append(part.CreateReferenceFromObject(plane))
    part.Update(); log(f"[OK] {len(plane_refs)} planes created.", "success")

    log("\n── Drawing Hollow Master Profile ───────────────────────────────", "head")
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(plane_refs[0]); sketch.Name = "Hose_Master_Profile"
    fd = sketch.OpenEdition()
    fd.CreateClosedCircle(0.0, 0.0, hose_data[0]["outer_dia"] / 2.0)
    fd.CreateClosedCircle(0.0, 0.0, hose_data[0]["inner_dia"] / 2.0)
    sketch.CloseEdition(); part.Update()
    log(f"[OK] Profile: OD={hose_data[0]['outer_dia']} / ID={hose_data[0]['inner_dia']} mm", "success")

    log("\n── Sweeping Rib via VBScript ────────────────────────────────────", "head")
    try:
        vbs  = 'Set CATIA = GetObject(, "CATIA.Application")\n'
        vbs += "Dim part1 : Set part1 = CATIA.ActiveDocument.Part\n"
        vbs += "Dim sf : Set sf = part1.ShapeFactory\n"
        vbs += "part1.InWorkObject = part1.MainBody\n"
        vbs += 'Dim sk1 : Set sk1 = part1.MainBody.Sketches.Item("Hose_Master_Profile")\n'
        vbs += "Dim refP : Set refP = part1.CreateReferenceFromObject(sk1)\n"
        vbs += 'Dim wf : Set wf = part1.HybridBodies.Item("Hose_Wireframe")\n'
        vbs += 'Dim spl : Set spl = wf.HybridShapes.Item("Hose_Centerline")\n'
        vbs += "Dim refC : Set refC = part1.CreateReferenceFromObject(spl)\n"
        vbs += "Dim rib1 : Set rib1 = sf.AddNewRibFromRef(refP, refC)\n"
        vbs += 'rib1.Name = "Hose_3D_Sweep"\npart1.Update\n'
        vbs += 'WScript.Echo ">>> 3D HOLLOW HOSE COMPLETE! <<<"\n'
        vbs_path = os.path.join(tempfile.gettempdir(), f"build_hose_{os.getpid()}.vbs")
        with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)
        time.sleep(1)
        result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
        log("  VBScript: " + result.stdout.strip(),
            "success" if "COMPLETE" in result.stdout else "info")
        if result.stderr: log("  VBScript error: " + result.stderr.strip(), "warn")
        try: os.remove(vbs_path)
        except: pass
        log("[DONE] Hose complete!", "success")
    except Exception as e:
        log(f"[ERROR] VBScript failed: {e}", "error")


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — INTERACTIVE BUILDER
# ═══════════════════════════════════════════════════════════════════════════════
def run_interactive_builder(log, pipe_def):
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed.", "error"); return

    nodes = pipe_def["nodes"]; segments = pipe_def["segments"]
    meta  = pipe_def["meta"];  part_name = meta["part_name"]

    log(f"[SETUP] Connecting to CATIA — part: {part_name}", "info")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True; catia.Documents.Add("Part")
        part_doc = catia.ActiveDocument; part = part_doc.Part
        try: part_doc.Product.PartNumber = part_name
        except: pass
        wf_set = part.HybridBodies.Add(); wf_set.Name = "Pipe_Wireframe"
        log(f"[SETUP] New Part '{part_name}' created.", "success")
    except Exception as e:
        log(f"[ERROR] CATIA connection failed: {e}", "error"); return

    hs = part.HybridShapeFactory

    log("\n── Building Polyline Centerline ────────────────────────────────", "head")
    polyline = hs.AddNewPolyline(); pt_refs = []
    for i, node in enumerate(nodes):
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"; wf_set.AppendHybridShape(pt)
        pt_ref = part.CreateReferenceFromObject(pt)
        polyline.InsertElement(pt_ref, i + 1); pt_refs.append(pt_ref)
        log(f"  Node {node['point']}: ({node['x']}, {node['y']}, {node['z']})", "info")

    for i in range(1, len(nodes) - 1):
        seg = segments[i]
        if seg["type"] == "curved" and seg.get("bend_radius"):
            polyline.SetRadius(i + 1, seg["bend_radius"])
            log(f"  Bend radius {seg['bend_radius']} mm at Node {i+1}", "info")

    polyline.Name = "Pipe_Centerline"; wf_set.AppendHybridShape(polyline); part.Update()
    log(f"[OK] Centerline built with {len(nodes)} nodes.", "success")

    log("\n── Creating Normal Planes ──────────────────────────────────────", "head")
    spline_ref = part.CreateReferenceFromObject(polyline); plane_refs = []
    for idx, pt_ref in enumerate(pt_refs):
        plane = hs.AddNewPlaneNormal(spline_ref, pt_ref)
        plane.Name = f"Profile_Plane_{idx+1}"; wf_set.AppendHybridShape(plane)
        plane_refs.append(part.CreateReferenceFromObject(plane))
    part.Update(); log(f"[OK] {len(plane_refs)} planes created.", "success")

    log("\n── Drawing Annular Cross-Section ───────────────────────────────", "head")
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(plane_refs[0]); sketch.Name = "Pipe_Master_Profile"
    fd = sketch.OpenEdition()
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["outer_dia"] / 2.0)
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["inner_dia"] / 2.0)
    sketch.CloseEdition(); part.Update()
    log(f"[OK] Profile: OD={nodes[0]['outer_dia']} / ID={nodes[0]['inner_dia']} mm", "success")

    log("\n── Sweeping Rib via VBScript ────────────────────────────────────", "head")
    try:
        vbs  = 'Set CATIA = GetObject(, "CATIA.Application")\n'
        vbs += "Dim part1 : Set part1 = CATIA.ActiveDocument.Part\n"
        vbs += "Dim sf : Set sf = part1.ShapeFactory\n"
        vbs += "part1.InWorkObject = part1.MainBody\n"
        vbs += 'Dim sk1 : Set sk1 = part1.MainBody.Sketches.Item("Pipe_Master_Profile")\n'
        vbs += "Dim refP : Set refP = part1.CreateReferenceFromObject(sk1)\n"
        vbs += 'Dim wf : Set wf = part1.HybridBodies.Item("Pipe_Wireframe")\n'
        vbs += 'Dim spl : Set spl = wf.HybridShapes.Item("Pipe_Centerline")\n'
        vbs += "Dim refC : Set refC = part1.CreateReferenceFromObject(spl)\n"
        vbs += "Dim rib1 : Set rib1 = sf.AddNewRibFromRef(refP, refC)\n"
        vbs += f'rib1.Name = "{part_name}_3D_Pipe"\npart1.Update\n'
        vbs += 'WScript.Echo ">>> 3D PIPE COMPLETE! <<<"\n'
        vbs_path = os.path.join(tempfile.gettempdir(), f"build_pipe_{os.getpid()}.vbs")
        with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)
        time.sleep(1)
        result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
        log("  VBScript: " + result.stdout.strip(),
            "success" if "COMPLETE" in result.stdout else "info")
        if result.stderr: log("  VBScript error: " + result.stderr.strip(), "warn")
        try: os.remove(vbs_path)
        except: pass
        log("[DONE] Pipe complete!", "success")
    except Exception as e:
        log(f"[ERROR] VBScript failed: {e}", "error")


# ═══════════════════════════════════════════════════════════════════════════════
#  DEFAULT HOSE DATA
# ═══════════════════════════════════════════════════════════════════════════════
DEFAULT_HOSE = [
    {"point": 1, "x": 0.0,   "y": 0.0,   "z": 0.0,  "outer_dia": 25.0, "inner_dia": 20.0},
    {"point": 2, "x": 100.0, "y": 50.0,  "z": 20.0, "outer_dia": 25.0, "inner_dia": 20.0},
    {"point": 3, "x": 150.0, "y": 150.0, "z": 50.0, "outer_dia": 20.0, "inner_dia": 15.0},
    {"point": 4, "x": 250.0, "y": 100.0, "z": 80.0, "outer_dia": 20.0, "inner_dia": 15.0},
]


# ═══════════════════════════════════════════════════════════════════════════════
#  GUI APPLICATION
# ═══════════════════════════════════════════════════════════════════════════════
class CatiaPipeApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("CATIA Pipe Automation Suite")
        self.geometry("1100x740")
        self.minsize(880, 600)
        self.configure(bg=BG_DEEP)
        self._apply_style()

        self._log_queue = queue.Queue()

        # ── FIX: one independent running flag per tab — tabs never block each other
        self._running = {1: False, 2: False, 3: False}

        self._build_layout()
        self._poll_log()

    # ─────────────────────────────────────────────────────────────────────────
    #  TTK STYLE
    # ─────────────────────────────────────────────────────────────────────────
    def _apply_style(self):
        s = ttk.Style(self); s.theme_use("clam")

        s.configure("TFrame",        background=BG_PANEL)
        s.configure("Deep.TFrame",   background=BG_DEEP)
        s.configure("Raised.TFrame", background=BG_RAISED)
        s.configure("TLabel",        background=BG_PANEL, foreground=TEXT_HI, font=FNT_UI)
        s.configure("Dim.TLabel",    background=BG_PANEL, foreground=TEXT_MID, font=FNT_SMALL)

        s.configure("TNotebook",     background=BG_DEEP, borderwidth=0, tabmargins=[0,0,0,0])
        s.configure("TNotebook.Tab", background=BG_RAISED, foreground=TEXT_MID,
                    font=FNT_UI_B, padding=[18, 9], borderwidth=0)
        s.map("TNotebook.Tab",
              background=[("selected", ACCENT_DIM)],
              foreground=[("selected", TEXT_HI)])

        s.configure("TEntry",
            fieldbackground=BG_RAISED, foreground=TEXT_HI, insertcolor=TEXT_HI,
            bordercolor=BG_BORDER, lightcolor=BG_BORDER, darkcolor=BG_BORDER,
            font=FNT_UI, padding=5)
        s.configure("TCombobox",
            fieldbackground=BG_RAISED, foreground=TEXT_HI,
            selectbackground=ACCENT_DIM, bordercolor=BG_BORDER, font=FNT_UI)
        s.configure("TScrollbar",
            background=BG_RAISED, troughcolor=BG_PANEL,
            arrowcolor=TEXT_MID, bordercolor=BG_PANEL)
        s.configure("TSeparator", background=BG_BORDER)

    # ─────────────────────────────────────────────────────────────────────────
    #  LAYOUT SKELETON
    # ─────────────────────────────────────────────────────────────────────────
    def _build_layout(self):
        # Header
        hdr = tk.Frame(self, bg=BG_PANEL, height=58)
        hdr.pack(fill=tk.X, side=tk.TOP); hdr.pack_propagate(False)

        # DS orange bar (3px accent stripe top of header)
        tk.Frame(hdr, bg=ACCENT, height=3).place(relx=0, rely=0, relwidth=1)

        tk.Label(hdr, text="◈", bg=BG_PANEL, fg=ACCENT,
                 font=("Segoe UI", 22, "bold")).pack(side=tk.LEFT, padx=(16, 6), pady=(10, 6))
        tk.Label(hdr, text="CATIA Pipe Automation Suite",
                 bg=BG_PANEL, fg=TEXT_HI, font=FNT_TITLE).pack(side=tk.LEFT, pady=(10, 6))
        tk.Label(hdr, text="Dassault Systèmes  |  CATIA V5",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FNT_SMALL).pack(side=tk.RIGHT, padx=18)

        # Tabs
        body = ttk.Frame(self, style="Deep.TFrame")
        body.pack(fill=tk.BOTH, expand=True)
        nb = ttk.Notebook(body)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        t1 = ttk.Frame(nb); nb.add(t1, text="  ① Extract & Rebuild  "); self._build_tab_extract(t1)
        t2 = ttk.Frame(nb); nb.add(t2, text="  ② Build from Coords  "); self._build_tab_coords(t2)
        t3 = ttk.Frame(nb); nb.add(t3, text="  ③ Interactive Builder  "); self._build_tab_interactive(t3)

        # Status bar
        sbar = tk.Frame(self, bg=BG_RAISED, height=26)
        sbar.pack(fill=tk.X, side=tk.BOTTOM); sbar.pack_propagate(False)
        self._status_var = tk.StringVar(value="Ready — all tabs independent")
        tk.Label(sbar, textvariable=self._status_var,
                 bg=BG_RAISED, fg=TEXT_MID, font=FNT_SMALL, anchor="w").pack(side=tk.LEFT, padx=10)
        self._dots = {}
        for tab_id, label in [(1,"T1"),(2,"T2"),(3,"T3")]:
            dot = tk.Label(sbar, text=f"● {label}", bg=BG_RAISED, fg=SUCCESS, font=FNT_SMALL)
            dot.pack(side=tk.RIGHT, padx=6)
            self._dots[tab_id] = dot

    # ─────────────────────────────────────────────────────────────────────────
    #  SHARED WIDGET FACTORIES
    # ─────────────────────────────────────────────────────────────────────────
    def _make_console(self, parent):
        frame = tk.Frame(parent, bg="#0D1520", bd=0)
        txt = tk.Text(frame, bg="#0D1520", fg=TEXT_HI, font=FNT_CODE,
                      insertbackground=TEXT_HI, relief=tk.FLAT, padx=10, pady=8,
                      selectbackground=ACCENT_DIM, wrap=tk.WORD, state=tk.DISABLED)
        scr = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=scr.set)
        scr.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        txt.tag_configure("head",    foreground=ACCENT)
        txt.tag_configure("success", foreground=SUCCESS)
        txt.tag_configure("warn",    foreground=WARNING)
        txt.tag_configure("error",   foreground=ERROR)
        txt.tag_configure("dim",     foreground=TEXT_DIM)
        txt.tag_configure("info",    foreground=TEXT_HI)
        return frame, txt

    def _console_write(self, txt, text, tag="info"):
        txt.configure(state=tk.NORMAL)
        txt.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n", tag)
        txt.see(tk.END); txt.configure(state=tk.DISABLED)

    def _clear_console(self, txt):
        txt.configure(state=tk.NORMAL); txt.delete("1.0", tk.END); txt.configure(state=tk.DISABLED)

    def _run_btn(self, parent, text, cmd):
        return tk.Button(parent, text=text, command=cmd,
                         bg=ACCENT, fg="#FFFFFF", activebackground=ACCENT_DIM,
                         activeforeground="#FFFFFF", relief=tk.FLAT, bd=0,
                         font=FNT_UI_B, padx=20, pady=8, cursor="hand2")

    def _clear_btn(self, parent, txt):
        return tk.Button(parent, text="Clear Log",
                         command=lambda: self._clear_console(txt),
                         bg=BG_RAISED, fg=TEXT_MID, activebackground=BG_BORDER,
                         activeforeground=TEXT_HI, relief=tk.FLAT, bd=0,
                         font=FNT_UI, padx=12, pady=5, cursor="hand2")

    def _section_label(self, parent, text):
        f = tk.Frame(parent, bg=BG_BORDER, height=1)
        f.pack(fill=tk.X, padx=12, pady=(8, 0))
        lf = tk.Frame(parent, bg=BG_PANEL)
        lf.pack(fill=tk.X, padx=12)
        tk.Label(lf, text=text, bg=BG_PANEL, fg=ACCENT, font=FNT_UI_B).pack(
            anchor="w", pady=(4, 2))

    def _lbl_entry(self, parent, row, label, var, width=14, col=0):
        tk.Label(parent, text=label, bg=BG_PANEL, fg=TEXT_MID, font=FNT_UI,
                 anchor="e").grid(row=row, column=col, sticky="e", padx=(4, 8), pady=5)
        e = ttk.Entry(parent, textvariable=var, width=width)
        e.grid(row=row, column=col+1, sticky="w", pady=5)
        return e

    # ─────────────────────────────────────────────────────────────────────────
    #  TAB 1 — EXTRACT & REBUILD
    # ─────────────────────────────────────────────────────────────────────────
    def _build_tab_extract(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=5, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # LEFT
        left = tk.Frame(pane, bg=BG_PANEL, width=290); pane.add(left, minsize=250)

        tk.Label(left, text="Configuration", bg=BG_PANEL, fg=ACCENT,
                 font=FNT_UI_L).pack(anchor="w", padx=14, pady=(14, 4))

        cfg = tk.Frame(left, bg=BG_PANEL); cfg.pack(fill=tk.X, padx=14, pady=4)
        cfg.columnconfigure(1, weight=1)
        self._e_radius       = tk.StringVar(value="15.0")
        self._e_straight_tol = tk.StringVar(value="0.5")
        self._e_dedup_tol    = tk.StringVar(value="5.0")
        self._lbl_entry(cfg, 0, "Pipe Radius (mm)",       self._e_radius)
        self._lbl_entry(cfg, 1, "Straightness Tol (mm)",  self._e_straight_tol)
        self._lbl_entry(cfg, 2, "Dedup Distance (mm)",    self._e_dedup_tol)

        tk.Frame(left, bg=BG_BORDER, height=1).pack(fill=tk.X, padx=14, pady=12)

        tk.Label(left,
            text="① Open a pipe model in CATIA V5.\n\n"
                 "② Adjust parameters above.\n\n"
                 "③ Click Run Pipeline.\n\n"
                 "Pipeline will:\n"
                 "  • Extract cylinder faces\n"
                 "  • Compute & deduplicate\n    centerline axes\n"
                 "  • Bridge bends with Connect\n"
                 "  • Sweep a solid Rib\n"
                 "  • Save the document",
            bg=BG_PANEL, fg=TEXT_MID, font=FNT_SMALL,
            justify=tk.LEFT, wraplength=230).pack(anchor="nw", padx=14, pady=4)

        btn_f = tk.Frame(left, bg=BG_PANEL); btn_f.pack(fill=tk.X, padx=14, pady=10, side=tk.BOTTOM)
        self._btn_extract = self._run_btn(btn_f, "▶  Run Pipeline", self._run_extract)
        self._btn_extract.pack(fill=tk.X)

        # RIGHT
        right = tk.Frame(pane, bg=BG_DEEP); pane.add(right, minsize=420)
        hdr2 = tk.Frame(right, bg=BG_PANEL, height=36); hdr2.pack(fill=tk.X); hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con1_frame, self._con1 = self._make_console(right)
        self._con1_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._clear_btn(hdr2, self._con1).pack(side=tk.RIGHT, padx=8, pady=5)
        self._console_write(self._con1, "Extract & Rebuild Mode — ready.", "head")
        self._console_write(self._con1, "Open a Part in CATIA V5, then click Run Pipeline.", "dim")

    def _run_extract(self):
        if self._running[1]: return             # ← only blocks this tab
        try:
            radius = float(self._e_radius.get())
            s_tol  = float(self._e_straight_tol.get())
            d_tol  = float(self._e_dedup_tol.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Enter valid numeric values."); return

        self._clear_console(self._con1)
        self._set_tab_running(1, True)
        con = self._con1

        def worker():
            # ── FIX: CoInitialize must be called in each new thread ──────────
            try:
                import pythoncom; pythoncom.CoInitialize()
            except ImportError: pass
            def log(m, t="info"): self._log_queue.put((con, m, t))
            try:
                run_extract_rebuild(log, pipe_radius=radius,
                                    straightness_tol=s_tol, dedup_tol=d_tol)
            except Exception as e:
                log(f"[FATAL] Unexpected error: {e}", "error")
            finally:
                # ── FIX: typed sentinel carries tab-id — no string comparison risk
                self._log_queue.put(_Done(1))
                try:
                    import pythoncom; pythoncom.CoUninitialize()
                except ImportError: pass

        threading.Thread(target=worker, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────────────
    #  TAB 2 — BUILD FROM COORDINATES
    # ─────────────────────────────────────────────────────────────────────────
    def _build_tab_coords(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=5, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # LEFT
        left = tk.Frame(pane, bg=BG_PANEL, width=390); pane.add(left, minsize=340)

        tk.Label(left, text="Route Nodes", bg=BG_PANEL, fg=ACCENT,
                 font=FNT_UI_L).pack(anchor="w", padx=14, pady=(14, 2))
        tk.Label(left, text="All dimensions in millimetres",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FNT_SMALL).pack(anchor="w", padx=14)

        # Column headers
        hdr_row = tk.Frame(left, bg=BG_RAISED)
        hdr_row.pack(fill=tk.X, padx=10, pady=(8, 0))
        for col_txt, w in [("Pt",3),("X",7),("Y",7),("Z",7),("OD",7),("ID",7)]:
            tk.Label(hdr_row, text=col_txt, bg=BG_RAISED, fg=ACCENT,
                     font=FNT_SMALL, width=w, anchor="center").pack(side=tk.LEFT, padx=2, pady=3)

        # Scrollable node canvas
        nc = tk.Canvas(left, bg=BG_PANEL, bd=0, highlightthickness=0)
        ns = ttk.Scrollbar(left, orient="vertical", command=nc.yview)
        self._node_frame = tk.Frame(nc, bg=BG_PANEL)
        self._node_frame.bind("<Configure>",
            lambda e: nc.configure(scrollregion=nc.bbox("all")))
        nc.create_window((0, 0), window=self._node_frame, anchor="nw")
        nc.configure(yscrollcommand=ns.set)
        nc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        ns.pack(side=tk.RIGHT, fill=tk.Y)

        self._node_vars = []
        for d in DEFAULT_HOSE: self._add_node_row(d)

        btn_bar = tk.Frame(left, bg=BG_PANEL); btn_bar.pack(fill=tk.X, padx=10, pady=4)
        tk.Button(btn_bar, text="+ Node", command=self._add_node_row,
                  bg=BG_RAISED, fg=SUCCESS, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_bar, text="− Remove", command=self._remove_node_row,
                  bg=BG_RAISED, fg=ERROR, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT)

        tk.Frame(left, bg=BG_BORDER, height=1).pack(fill=tk.X, padx=10, pady=6)
        self._run_coords_btn = self._run_btn(left, "▶  Build Hose in CATIA", self._run_coords)
        self._run_coords_btn.pack(fill=tk.X, padx=10, pady=(0, 10))

        # RIGHT
        right = tk.Frame(pane, bg=BG_DEEP); pane.add(right, minsize=380)
        hdr2 = tk.Frame(right, bg=BG_PANEL, height=36); hdr2.pack(fill=tk.X); hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con2_frame, self._con2 = self._make_console(right)
        self._con2_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._clear_btn(hdr2, self._con2).pack(side=tk.RIGHT, padx=8, pady=5)
        self._console_write(self._con2, "Build from Coordinates — ready.", "head")
        self._console_write(self._con2, "Define nodes in the table, then click Build.", "dim")

    def _add_node_row(self, data=None):
        idx = len(self._node_vars) + 1
        if data is None:
            data = {"point": idx, "x": 0.0, "y": 0.0, "z": 0.0,
                    "outer_dia": 25.0, "inner_dia": 20.0}
        bg = BG_RAISED if len(self._node_vars) % 2 == 0 else BG_PANEL
        v  = {k: tk.StringVar(value=str(data[k2]))
              for k, k2 in [("x","x"),("y","y"),("z","z"),("od","outer_dia"),("id","inner_dia")]}
        f  = tk.Frame(self._node_frame, bg=bg); f.pack(fill=tk.X, pady=1)
        tk.Label(f, text=str(data["point"]), bg=bg, fg=TEXT_MID,
                 font=FNT_SMALL, width=3, anchor="center").pack(side=tk.LEFT, padx=2)
        for key in ["x","y","z","od","id"]:
            tk.Entry(f, textvariable=v[key], width=7,
                     bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                     relief=tk.FLAT, font=FNT_CODE).pack(side=tk.LEFT, padx=2, pady=2)
        self._node_vars.append(v)

    def _remove_node_row(self):
        if len(self._node_vars) <= 2: return
        self._node_vars.pop()
        kids = self._node_frame.winfo_children()
        if kids: kids[-1].destroy()

    def _collect_hose_data(self):
        result = []
        for i, v in enumerate(self._node_vars):
            try:
                result.append({"point": i+1,
                    "x": float(v["x"].get()), "y": float(v["y"].get()), "z": float(v["z"].get()),
                    "outer_dia": float(v["od"].get()), "inner_dia": float(v["id"].get())})
            except ValueError:
                messagebox.showerror("Invalid Input", f"Row {i+1} has non-numeric values."); return None
        return result

    def _run_coords(self):
        if self._running[2]: return             # ← only blocks Tab 2
        data = self._collect_hose_data()
        if data is None: return
        self._clear_console(self._con2)
        self._set_tab_running(2, True)
        con = self._con2

        def worker():
            try:
                import pythoncom; pythoncom.CoInitialize()
            except ImportError: pass
            def log(m, t="info"): self._log_queue.put((con, m, t))
            try:
                run_build_from_coords(log, data)
            except Exception as e:
                log(f"[FATAL] Unexpected error: {e}", "error")
            finally:
                self._log_queue.put(_Done(2))
                try:
                    import pythoncom; pythoncom.CoUninitialize()
                except ImportError: pass

        threading.Thread(target=worker, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────────────
    #  TAB 3 — INTERACTIVE BUILDER
    # ─────────────────────────────────────────────────────────────────────────
    def _build_tab_interactive(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=5, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # LEFT
        left = tk.Frame(pane, bg=BG_PANEL, width=390); pane.add(left, minsize=340)

        tk.Label(left, text="Pipe Definition", bg=BG_PANEL, fg=ACCENT,
                 font=FNT_UI_L).pack(anchor="w", padx=14, pady=(14, 4))

        # Global settings
        gs = tk.LabelFrame(left, text=" Global Settings ",
                           bg=BG_PANEL, fg=TEXT_MID, font=FNT_SMALL,
                           bd=1, relief=tk.RIDGE); gs.pack(fill=tk.X, padx=12, pady=(4, 0))
        gs.columnconfigure(1, weight=1)
        self._iv_name = tk.StringVar(value="Custom_Pipe")
        self._iv_od   = tk.StringVar(value="25.0")
        self._iv_id   = tk.StringVar(value="20.0")
        self._lbl_entry(gs, 0, "Part Name",       self._iv_name, width=18)
        self._lbl_entry(gs, 1, "Default OD (mm)", self._iv_od,   width=10)
        self._lbl_entry(gs, 2, "Default ID (mm)", self._iv_id,   width=10)

        # Segment list header
        sh = tk.Frame(left, bg=BG_PANEL); sh.pack(fill=tk.X, padx=12, pady=(10, 2))
        tk.Label(sh, text="Segments", bg=BG_PANEL, fg=TEXT_HI, font=FNT_UI_B).pack(side=tk.LEFT)
        tk.Button(sh, text="+ Segment", command=self._add_seg_row,
                  bg=BG_RAISED, fg=SUCCESS, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(sh, text="− Remove", command=self._remove_seg_row,
                  bg=BG_RAISED, fg=ERROR, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.RIGHT, padx=4)

        # Segment scroll
        sc = tk.Canvas(left, bg=BG_PANEL, bd=0, highlightthickness=0, height=280)
        ss = ttk.Scrollbar(left, orient="vertical", command=sc.yview)
        self._seg_frame = tk.Frame(sc, bg=BG_PANEL)
        self._seg_frame.bind("<Configure>",
            lambda e: sc.configure(scrollregion=sc.bbox("all")))
        sc.create_window((0, 0), window=self._seg_frame, anchor="nw")
        sc.configure(yscrollcommand=ss.set)
        sc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))
        ss.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4))

        self._seg_vars = []
        for _ in range(3): self._add_seg_row()

        # JSON + run buttons
        act = tk.Frame(left, bg=BG_PANEL); act.pack(fill=tk.X, padx=12, pady=(6, 4))
        tk.Button(act, text="📂  Load JSON", command=self._load_json,
                  bg=BG_RAISED, fg=TEXT_MID, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(act, text="💾  Save JSON", command=self._save_json,
                  bg=BG_RAISED, fg=TEXT_MID, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT)

        tk.Frame(left, bg=BG_BORDER, height=1).pack(fill=tk.X, padx=12, pady=6)
        self._run_ib_btn = self._run_btn(left, "▶  Build Pipe in CATIA", self._run_interactive)
        self._run_ib_btn.pack(fill=tk.X, padx=12, pady=(0, 10))

        # RIGHT
        right = tk.Frame(pane, bg=BG_DEEP); pane.add(right, minsize=380)
        hdr2 = tk.Frame(right, bg=BG_PANEL, height=36); hdr2.pack(fill=tk.X); hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con3_frame, self._con3 = self._make_console(right)
        self._con3_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        self._clear_btn(hdr2, self._con3).pack(side=tk.RIGHT, padx=8, pady=5)
        self._console_write(self._con3, "Interactive Builder — ready.", "head")
        self._console_write(self._con3, "Define segments, then click Build.", "dim")

    def _add_seg_row(self):
        idx = len(self._seg_vars) + 1
        bg  = BG_RAISED if idx % 2 == 0 else BG_PANEL
        f   = tk.LabelFrame(self._seg_frame, text=f" Segment {idx} ",
                            bg=bg, fg=ACCENT, font=FNT_SMALL, bd=1, relief=tk.GROOVE)
        f.pack(fill=tk.X, padx=4, pady=4); f.columnconfigure(1, weight=1)

        v = {k: tk.StringVar(value=d) for k, d in [
            ("p1x","0.0"),("p1y","0.0"),("p1z","0.0"),
            ("p2x","100.0"),("p2y","0.0"),("p2z","0.0"),
            ("type","straight"),("bend",""),("od","25.0"),("id","20.0")]}

        for row_i, (lbl, keys) in enumerate([
            ("Start (X Y Z)", ["p1x","p1y","p1z"]),
            ("End   (X Y Z)", ["p2x","p2y","p2z"])]):
            tk.Label(f, text=lbl, bg=bg, fg=TEXT_MID,
                     font=FNT_SMALL).grid(row=row_i, column=0, sticky="e", padx=(6,2), pady=2)
            for ci, k in enumerate(keys):
                tk.Entry(f, textvariable=v[k], width=7, bg=BG_RAISED, fg=TEXT_HI,
                         insertbackground=TEXT_HI, relief=tk.FLAT,
                         font=FNT_CODE).grid(row=row_i, column=ci+1, padx=2, pady=2)

        tk.Label(f, text="Type", bg=bg, fg=TEXT_MID,
                 font=FNT_SMALL).grid(row=2, column=0, sticky="e", padx=(6,2), pady=2)
        ttk.Combobox(f, textvariable=v["type"], values=["straight","curved"],
                     width=9, state="readonly").grid(row=2, column=1, columnspan=2, sticky="w", pady=2)
        tk.Label(f, text="Bend R", bg=bg, fg=TEXT_MID,
                 font=FNT_SMALL).grid(row=2, column=3, sticky="e", padx=(8,2))
        tk.Entry(f, textvariable=v["bend"], width=7, bg=BG_RAISED, fg=TEXT_HI,
                 insertbackground=TEXT_HI, relief=tk.FLAT,
                 font=FNT_CODE).grid(row=2, column=4, padx=2, pady=2)

        tk.Label(f, text="OD / ID", bg=bg, fg=TEXT_MID,
                 font=FNT_SMALL).grid(row=3, column=0, sticky="e", padx=(6,2), pady=(2,6))
        tk.Entry(f, textvariable=v["od"], width=7, bg=BG_RAISED, fg=TEXT_HI,
                 insertbackground=TEXT_HI, relief=tk.FLAT,
                 font=FNT_CODE).grid(row=3, column=1, pady=(2,6))
        tk.Label(f, text="/", bg=bg, fg=TEXT_DIM,
                 font=FNT_UI).grid(row=3, column=2)
        tk.Entry(f, textvariable=v["id"], width=7, bg=BG_RAISED, fg=TEXT_HI,
                 insertbackground=TEXT_HI, relief=tk.FLAT,
                 font=FNT_CODE).grid(row=3, column=3, pady=(2,6))

        self._seg_vars.append({"vars": v, "frame": f})

    def _remove_seg_row(self):
        if len(self._seg_vars) <= 1: return
        entry = self._seg_vars.pop(); entry["frame"].destroy()

    def _collect_pipe_def(self):
        try:
            part_name = self._iv_name.get().strip() or "Custom_Pipe"
            def_od = float(self._iv_od.get()); def_id = float(self._iv_id.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Check global settings."); return None

        nodes = []; segments = []
        for i, entry in enumerate(self._seg_vars):
            v = entry["vars"]
            try:
                p1 = tuple(float(v[k].get()) for k in ["p1x","p1y","p1z"])
                p2 = tuple(float(v[k].get()) for k in ["p2x","p2y","p2z"])
                od = float(v["od"].get()); id_ = float(v["id"].get())
                seg_type = v["type"].get()
                bend_r = float(v["bend"].get()) if v["bend"].get().strip() else None
            except ValueError:
                messagebox.showerror("Invalid Input", f"Segment {i+1} has invalid values."); return None

            chord = math.sqrt(sum((b-a)**2 for a,b in zip(p1,p2)))
            if not nodes or (nodes[-1]["x"],nodes[-1]["y"],nodes[-1]["z"]) != p1:
                nodes.append({"point":len(nodes)+1,"x":p1[0],"y":p1[1],"z":p1[2],
                               "outer_dia":od,"inner_dia":id_})
            nodes.append({"point":len(nodes)+1,"x":p2[0],"y":p2[1],"z":p2[2],
                           "outer_dia":od,"inner_dia":id_})
            segments.append({"segment":i+1,"from_point":i+1,"to_point":i+2,
                              "type":seg_type,"bend_radius":bend_r,
                              "outer_dia":od,"inner_dia":id_,"chord_length":round(chord,4)})

        uniq = [nodes[0]]
        for n in nodes[1:]:
            if (n["x"],n["y"],n["z"]) != (uniq[-1]["x"],uniq[-1]["y"],uniq[-1]["z"]):
                uniq.append(n)
        for i, n in enumerate(uniq): n["point"] = i + 1

        return {"meta":{"created":datetime.now().isoformat(timespec="seconds"),
                        "part_name":part_name,"num_points":len(uniq),
                        "num_segments":len(segments),
                        "default_outer_dia":def_od,"default_inner_dia":def_id},
                "nodes":uniq,"segments":segments}

    def _load_json(self):
        path = filedialog.askopenfilename(title="Load Pipe Definition",
                                          filetypes=[("JSON","*.json"),("All","*.*")])
        if not path: return
        try:
            with open(path) as f: data = json.load(f)
            self._iv_name.set(data["meta"]["part_name"])
            self._iv_od.set(str(data["meta"]["default_outer_dia"]))
            self._iv_id.set(str(data["meta"]["default_inner_dia"]))
            messagebox.showinfo("Loaded",
                f"Loaded: {data['meta']['part_name']}\n"
                f"Points: {data['meta']['num_points']}   "
                f"Segments: {data['meta']['num_segments']}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load JSON:\n{e}")

    def _save_json(self):
        pipe_def = self._collect_pipe_def()
        if pipe_def is None: return
        path = filedialog.asksaveasfilename(title="Save Pipe Definition",
                                            defaultextension=".json",
                                            filetypes=[("JSON","*.json")])
        if not path: return
        try:
            with open(path, "w") as f: json.dump(pipe_def, f, indent=2)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save:\n{e}")

    def _run_interactive(self):
        if self._running[3]: return             # ← only blocks Tab 3
        pipe_def = self._collect_pipe_def()
        if pipe_def is None: return
        self._clear_console(self._con3)
        self._set_tab_running(3, True)
        con = self._con3
        self._console_write(con, f"Part Name : {pipe_def['meta']['part_name']}", "head")
        self._console_write(con, f"Nodes     : {pipe_def['meta']['num_points']}", "head")
        self._console_write(con, f"Segments  : {pipe_def['meta']['num_segments']}", "head")

        def worker():
            try:
                import pythoncom; pythoncom.CoInitialize()
            except ImportError: pass
            def log(m, t="info"): self._log_queue.put((con, m, t))
            try:
                run_interactive_builder(log, pipe_def)
            except Exception as e:
                log(f"[FATAL] Unexpected error: {e}", "error")
            finally:
                self._log_queue.put(_Done(3))
                try:
                    import pythoncom; pythoncom.CoUninitialize()
                except ImportError: pass

        threading.Thread(target=worker, daemon=True).start()

    # ─────────────────────────────────────────────────────────────────────────
    #  PER-TAB STATE HELPERS
    # ─────────────────────────────────────────────────────────────────────────
    def _set_tab_running(self, tab_id: int, running: bool):
        """Lock/unlock only the button for tab_id. Other tabs are unaffected."""
        self._running[tab_id] = running
        btn_map = {1: self._btn_extract, 2: self._run_coords_btn, 3: self._run_ib_btn}
        btn = btn_map[tab_id]
        btn.configure(state="disabled" if running else "normal",
                      bg=BG_BORDER if running else ACCENT)
        dot = self._dots[tab_id]
        dot.configure(fg=WARNING if running else SUCCESS)

        # Update status bar: list which tabs are currently active
        active = [f"T{t}" for t, r in self._running.items() if r]
        if active:
            self._status_var.set(f"Running: {', '.join(active)}")
        else:
            self._status_var.set("Ready — all tabs independent")

    # ─────────────────────────────────────────────────────────────────────────
    #  LOG QUEUE POLL  — runs on the Tk main thread every 60 ms
    # ─────────────────────────────────────────────────────────────────────────
    def _poll_log(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                # ── FIX: isinstance check — safe, no widget == string comparison
                if isinstance(item, _Done):
                    self._set_tab_running(item.tab, False)
                else:
                    # item is (console_widget, message, tag)
                    con, msg, tag = item
                    self._console_write(con, msg, tag)
        except queue.Empty:
            pass
        self.after(60, self._poll_log)


# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = CatiaPipeApp()
    app.mainloop()
