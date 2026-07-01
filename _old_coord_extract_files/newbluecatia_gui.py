"""
CATIA Pipe Automation Suite — Professional GUI
Bundles three automation modes:
  1. Extract & Rebuild  — scans an open CATIA model and generates a solid pipe
  2. Build from Coords  — sweeps a hose from hardcoded 3-D coordinate list
  3. Interactive Builder — full prompts: points, segments, diameters → JSON → CATIA
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import threading
import queue
import sys
import os
import io
import json
import math
import time
import tempfile
import subprocess
from datetime import datetime


# ═══════════════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS  (dark industrial / CAD software aesthetic)
# ═══════════════════════════════════════════════════════════════════════════════
BG_DEEP    = "#0B0F19"   # deepest background
BG_PANEL   = "#111827"   # card / panel background
BG_RAISED  = "#1A2236"   # inputs, raised surfaces
BG_BORDER  = "#243049"   # borders, dividers
ACCENT     = "#1D8CF8"   # primary blue
ACCENT_DIM = "#0F4C8A"   # hover/pressed blue
SUCCESS    = "#00D4AA"   # green-teal for success
WARNING    = "#FFB547"   # amber for warnings
ERROR      = "#FF4D6A"   # red for errors
TEXT_HI    = "#E8EEFF"   # primary text
TEXT_MID   = "#8896B3"   # secondary text
TEXT_DIM   = "#4A5A7A"   # disabled / dim text

FNT_UI     = ("Segoe UI", 10)
FNT_UI_B   = ("Segoe UI", 10, "bold")
FNT_UI_L   = ("Segoe UI", 12, "bold")
FNT_TITLE  = ("Segoe UI", 18, "bold")
FNT_CODE   = ("Consolas", 9)
FNT_SMALL  = ("Segoe UI", 8)


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — EXTRACT & REBUILD
# ═══════════════════════════════════════════════════════════════════════════════
def run_extract_rebuild(log, pipe_radius=15.0, straightness_tol=0.5, dedup_tol=5.0):
    """Phase 1→2→3 pipeline. Reads open CATIA model, produces solid Rib."""
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed. This must run on Windows with CATIA V5.", "error")
        return

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
        log("  Make sure CATIA V5 is open with a Part document active.", "warn")
        return

    log("[SETUP] Connected successfully.", "success")

    # ── PHASE 1 ────────────────────────────────────────────────────────────────
    log("\n── PHASE 1: Extracting cylinder faces ──────────────────────────", "head")
    try:
        old = hybrid_bodies.Item("Extracted_Cylinders")
        selection.Clear(); selection.Add(old); selection.Delete()
        log("[Phase 1] Cleared old 'Extracted_Cylinders' set.", "dim")
    except: pass

    extraction_set      = hybrid_bodies.Add()
    extraction_set.Name = "Extracted_Cylinders"

    selection.Clear()
    selection.Search("Topology.Face,all")
    total_faces = selection.Count
    log(f"[Phase 1] Found {total_faces} total faces in model.", "info")

    extracted_count = 0
    all_extractions = []

    for i in range(1, total_faces + 1):
        item       = selection.Item(i)
        shape_type = str(item.Type)
        face_ref   = item.Reference
        if "ylinder" in shape_type:
            ext      = hs_factory.AddNewExtract(face_ref)
            ext.Name = f"Cyl_Face_{extracted_count + 1}"
            extraction_set.AppendHybridShape(ext)
            all_extractions.append(ext)
            extracted_count += 1

    if extracted_count == 0:
        log("[Phase 1] No cylinder type detected — falling back to ALL faces.", "warn")
        selection.Clear(); selection.Search("Topology.Face,all")
        for i in range(1, selection.Count + 1):
            face_ref = selection.Item(i).Reference
            ext      = hs_factory.AddNewExtract(face_ref)
            ext.Name = f"Face_{i}"
            extraction_set.AppendHybridShape(ext)
            all_extractions.append(ext)
            extracted_count += 1

    try:
        part.Update()
        log(f"[Phase 1] {extracted_count} faces stored in 'Extracted_Cylinders'.", "success")
    except Exception as e:
        log(f"[Phase 1] Update warning: {e}", "warn")

    if extracted_count == 0:
        log("[Phase 1] FAILED: No faces found. Load a pipe model.", "error"); return

    # ── PHASE 2 ────────────────────────────────────────────────────────────────
    log("\n── PHASE 2: Computing centerlines ──────────────────────────────", "head")
    try:
        old_cl = hybrid_bodies.Item("Extracted_Centerlines")
        selection.Clear(); selection.Add(old_cl); selection.Delete()
    except: pass

    centerline_set      = hybrid_bodies.Add()
    centerline_set.Name = "Extracted_Centerlines"

    faces    = extraction_set.HybridShapes
    all_lines = []

    for i in range(1, faces.Count + 1):
        current_face = faces.Item(i)
        selection.Clear(); selection.Add(current_face)
        try:   selection.Search("Topology.Edge,sel")
        except: continue
        if selection.Count == 0: continue

        points_data = []
        for j in range(1, selection.Count + 1):
            try:
                raw_ref = selection.Item(j).Reference
                ee      = hs_factory.AddNewExtract(raw_ref)
                centerline_set.AppendHybridShape(ee)
                pt = hs_factory.AddNewPointCenter(part.CreateReferenceFromObject(ee))
                centerline_set.AppendHybridShape(pt)
                part.UpdateObject(pt)
                points_data.append(pt)
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

        line = hs_factory.AddNewLinePtPt(
            part.CreateReferenceFromObject(best_p1),
            part.CreateReferenceFromObject(best_p2)
        )
        centerline_set.AppendHybridShape(line); part.UpdateObject(line)

        midpt = hs_factory.AddNewPointOnCurveFromPercent(
            part.CreateReferenceFromObject(line), 0.5, True
        )
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

    # Deduplication
    log("[Phase 2] Deduplicating overlapping axis lines...", "info")
    unique_lines = []
    for ld in all_lines:
        meas = spa.GetMeasurable(part.CreateReferenceFromObject(ld["midpoint"]))
        is_dup = False
        for ul in unique_lines:
            try:
                if meas.GetMinimumDistance(part.CreateReferenceFromObject(ul["midpoint"])) < dedup_tol:
                    is_dup = True
                    selection.Clear()
                    for o in [ld["line"], ld["p1"], ld["p2"], ld["midpoint"]]:
                        selection.Add(o)
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

    # Tip-to-tail sort
    log("[Phase 2] Sorting segments tip-to-tail...", "info")
    meas_ref = spa.GetMeasurable(part.CreateReferenceFromObject(unique_lines[0]["p1"]))
    max_d = -1; start_line = unique_lines[0]
    for ld in unique_lines:
        try:
            d = meas_ref.GetMinimumDistance(part.CreateReferenceFromObject(ld["line"]))
            if d > max_d: max_d = d; start_line = ld
        except: pass

    sorted_lines = [start_line]
    unvisited    = [ld for ld in unique_lines if ld is not start_line]
    active_tip   = start_line["p1"]

    if unvisited:
        mp1 = spa.GetMeasurable(part.CreateReferenceFromObject(start_line["p1"]))
        mp2 = spa.GetMeasurable(part.CreateReferenceFromObject(start_line["p2"]))
        md1 = md2 = float("inf")
        for ld in unvisited:
            ref_l = part.CreateReferenceFromObject(ld["line"])
            try:
                d1 = mp1.GetMinimumDistance(ref_l)
                d2 = mp2.GetMinimumDistance(ref_l)
                if d1 < md1: md1 = d1
                if d2 < md2: md2 = d2
            except: pass
        active_tip = start_line["p2"] if md2 < md1 else start_line["p1"]

    while unvisited:
        meas_tip  = spa.GetMeasurable(part.CreateReferenceFromObject(active_tip))
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
        p1a, p1b = l1["p1"], l1["p2"]
        p2a, p2b = l2["p1"], l2["p2"]

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
                False
            )
            conn.Name = f"Bend_Connect_{i+1}"
            centerline_set.AppendHybridShape(conn); part.UpdateObject(conn)
            connects.append(conn)
            log(f"  Bend {i+1} connected (orientations: {ori1}, {ori2})", "info")
        except Exception as e:
            log(f"  WARNING: Bend {i+1} connect failed: {e}", "warn")

    log(f"[Phase 3] {len(connects)} bend(s) bridged.", "success")

    log("[Phase 3] Joining all elements into Master_Centerline_Join...", "info")
    join_elements = [sorted_lines[0]["line"]]
    for i in range(len(connects)):
        join_elements.append(connects[i])
        join_elements.append(sorted_lines[i + 1]["line"])

    if len(join_elements) < 2:
        log("[Phase 3] FAILED: Not enough elements.", "error"); return

    try:
        master_join = hs_factory.AddNewJoin(
            part.CreateReferenceFromObject(join_elements[0]),
            part.CreateReferenceFromObject(join_elements[1])
        )
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
            part.CreateReferenceFromObject(plane),
            False, pipe_radius
        )
        circle.Name = "Pipe_Profile_Circle"
        centerline_set.AppendHybridShape(circle)
        part.Update()
        log(f"[Phase 3] Profile circle created (radius = {pipe_radius} mm).", "success")
    except Exception as e:
        log(f"[Phase 3] FAILED creating profile circle: {e}", "error"); return

    try:
        part.InWorkObject = part.MainBody
        rib = shape_factory.AddNewRibFromRef(
            part.CreateReferenceFromObject(circle), join_ref
        )
        part.Update()
        log("[Phase 3] >>> SOLID RIB GENERATED SUCCESSFULLY! <<<", "success")
    except Exception as e:
        log(f"[Phase 3] FAILED during Rib: {e}", "error"); return

    try:
        part_document.Save()
        log("[DONE] Document saved.", "success")
    except Exception as e:
        log(f"[DONE] Save warning: {e}", "warn")

    log(f"\n{'='*55}", "head")
    log(f" PIPELINE COMPLETE", "head")
    log(f"  Segments extracted : {extracted_count}", "head")
    log(f"  Unique axis lines  : {len(unique_lines)}", "head")
    log(f"  Bends connected    : {len(connects)}", "head")
    log(f"  Solid pipe         : Rib in PartBody", "head")
    log(f"{'='*55}", "head")


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — BUILD FROM COORDINATES
# ═══════════════════════════════════════════════════════════════════════════════
def run_build_from_coords(log, hose_data):
    """Builds a hollow hose from a list of coordinate+diameter nodes."""
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed. This must run on Windows with CATIA V5.", "error")
        return

    log("[SETUP] Connecting to CATIA V5...", "info")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        catia.Documents.Add("Part")
        part_doc  = catia.ActiveDocument
        part      = part_doc.Part
        try:
            part_doc.Product.PartNumber = "AC_Hose_Automated"
        except: pass
        hb       = part.HybridBodies
        wf_set   = hb.Add(); wf_set.Name = "Hose_Wireframe"
        log("[SETUP] New Part created: AC_Hose_Automated", "success")
    except Exception as e:
        log(f"[ERROR] CATIA connection failed: {e}", "error"); return

    # Spine
    log("\n── Building Spline Centerline ──────────────────────────────────", "head")
    hs  = part.HybridShapeFactory
    spline = hs.AddNewSpline(); spline.SetSplineType(0)
    pt_refs = []

    for node in hose_data:
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"
        wf_set.AppendHybridShape(pt)
        pt_ref = part.CreateReferenceFromObject(pt)
        spline.AddPoint(pt_ref)
        pt_refs.append(pt_ref)
        log(f"  Node {node['point']}: ({node['x']}, {node['y']}, {node['z']})", "info")

    spline.Name = "Hose_Centerline"
    wf_set.AppendHybridShape(spline); part.Update()
    log(f"[OK] Spline created with {len(hose_data)} nodes.", "success")

    # Planes
    log("\n── Creating Normal Planes ──────────────────────────────────────", "head")
    spline_ref = part.CreateReferenceFromObject(spline)
    plane_refs = []
    for idx, pt_ref in enumerate(pt_refs):
        plane = hs.AddNewPlaneNormal(spline_ref, pt_ref)
        plane.Name = f"Profile_Plane_{idx+1}"
        wf_set.AppendHybridShape(plane)
        plane_refs.append(part.CreateReferenceFromObject(plane))
    part.Update()
    log(f"[OK] {len(plane_refs)} planes created.", "success")

    # Master Profile
    log("\n── Drawing Hollow Master Profile ───────────────────────────────", "head")
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(plane_refs[0])
    sketch.Name = "Hose_Master_Profile"
    fd = sketch.OpenEdition()
    outer_r = hose_data[0]["outer_dia"] / 2.0
    inner_r = hose_data[0]["inner_dia"] / 2.0
    fd.CreateClosedCircle(0.0, 0.0, outer_r)
    fd.CreateClosedCircle(0.0, 0.0, inner_r)
    sketch.CloseEdition(); part.Update()
    log(f"[OK] Profile: OD={hose_data[0]['outer_dia']} / ID={hose_data[0]['inner_dia']} mm", "success")

    # Rib via VBScript
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
        vbs += 'rib1.Name = "Hose_3D_Sweep"\n'
        vbs += "part1.Update\n"
        vbs += 'WScript.Echo ">>> 3D HOLLOW HOSE COMPLETE! <<<"\n'

        vbs_path = os.path.join(tempfile.gettempdir(), "build_hose_rib.vbs")
        with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)
        time.sleep(1)
        result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
        log("  VBScript: " + result.stdout.strip(), "success" if "COMPLETE" in result.stdout else "info")
        if result.stderr: log("  VBScript error: " + result.stderr.strip(), "warn")
        os.remove(vbs_path)
        log("[DONE] Hose complete!", "success")
    except Exception as e:
        log(f"[ERROR] VBScript failed: {e}", "error")


# ═══════════════════════════════════════════════════════════════════════════════
#  CATIA PIPELINE — INTERACTIVE BUILDER
# ═══════════════════════════════════════════════════════════════════════════════
def run_interactive_builder(log, pipe_def):
    """Builds a pipe from a fully-specified pipe_def dict (from the form)."""
    try:
        import win32com.client
    except ImportError:
        log("[ERROR] pywin32 not installed. This must run on Windows with CATIA V5.", "error")
        return

    nodes     = pipe_def["nodes"]
    segments  = pipe_def["segments"]
    meta      = pipe_def["meta"]
    part_name = meta["part_name"]

    log(f"[SETUP] Connecting to CATIA — creating part: {part_name}", "info")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        catia.Documents.Add("Part")
        part_doc = catia.ActiveDocument
        part     = part_doc.Part
        try: part_doc.Product.PartNumber = part_name
        except: pass
        wf_set = part.HybridBodies.Add(); wf_set.Name = "Pipe_Wireframe"
        log(f"[SETUP] New Part '{part_name}' created.", "success")
    except Exception as e:
        log(f"[ERROR] CATIA connection failed: {e}", "error"); return

    hs = part.HybridShapeFactory

    # Polyline spine
    log("\n── Building Polyline Centerline ────────────────────────────────", "head")
    polyline = hs.AddNewPolyline()
    pt_refs  = []

    for i, node in enumerate(nodes):
        pt = hs.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"
        wf_set.AppendHybridShape(pt)
        pt_ref = part.CreateReferenceFromObject(pt)
        polyline.InsertElement(pt_ref, i + 1)
        pt_refs.append(pt_ref)
        log(f"  Node {node['point']}: ({node['x']}, {node['y']}, {node['z']})", "info")

    for i in range(1, len(nodes) - 1):
        seg = segments[i]
        if seg["type"] == "curved" and seg.get("bend_radius"):
            polyline.SetRadius(i + 1, seg["bend_radius"])
            log(f"  Bend radius {seg['bend_radius']} mm applied at Node {i+1}", "info")

    polyline.Name = "Pipe_Centerline"
    wf_set.AppendHybridShape(polyline); part.Update()
    log(f"[OK] Centerline built with {len(nodes)} nodes.", "success")

    # Normal planes
    log("\n── Creating Normal Planes ──────────────────────────────────────", "head")
    spline_ref = part.CreateReferenceFromObject(polyline)
    plane_refs = []
    for idx, pt_ref in enumerate(pt_refs):
        plane = hs.AddNewPlaneNormal(spline_ref, pt_ref)
        plane.Name = f"Profile_Plane_{idx+1}"
        wf_set.AppendHybridShape(plane)
        plane_refs.append(part.CreateReferenceFromObject(plane))
    part.Update()
    log(f"[OK] {len(plane_refs)} planes created.", "success")

    # Master profile
    log("\n── Drawing Annular Cross-Section ───────────────────────────────", "head")
    part.InWorkObject = part.MainBody
    sketch = part.MainBody.Sketches.Add(plane_refs[0])
    sketch.Name = "Pipe_Master_Profile"
    fd = sketch.OpenEdition()
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["outer_dia"] / 2.0)
    fd.CreateClosedCircle(0.0, 0.0, nodes[0]["inner_dia"] / 2.0)
    sketch.CloseEdition(); part.Update()
    log(f"[OK] Profile: OD={nodes[0]['outer_dia']} / ID={nodes[0]['inner_dia']} mm", "success")

    # Rib via VBScript
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
        vbs += f'rib1.Name = "{part_name}_3D_Pipe"\n'
        vbs += "part1.Update\n"
        vbs += 'WScript.Echo ">>> 3D PIPE COMPLETE! <<<"\n'

        vbs_path = os.path.join(tempfile.gettempdir(), "build_pipe_rib.vbs")
        with open(vbs_path, "w", encoding="ascii") as f: f.write(vbs)
        time.sleep(1)
        result = subprocess.run(["cscript", "//Nologo", vbs_path], capture_output=True, text=True)
        log("  VBScript: " + result.stdout.strip(), "success" if "COMPLETE" in result.stdout else "info")
        if result.stderr: log("  VBScript error: " + result.stderr.strip(), "warn")
        os.remove(vbs_path)
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
        self.geometry("1080x720")
        self.minsize(860, 600)
        self.configure(bg=BG_DEEP)
        self._apply_style()

        self._log_queue = queue.Queue()
        self._running   = False

        self._build_layout()
        self._poll_log()

    # ── Ttk style ──────────────────────────────────────────────────────────────
    def _apply_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame",       background=BG_PANEL)
        style.configure("Deep.TFrame",  background=BG_DEEP)
        style.configure("Raised.TFrame",background=BG_RAISED)

        style.configure("TLabel",
            background=BG_PANEL, foreground=TEXT_HI,
            font=FNT_UI)
        style.configure("Dim.TLabel",   background=BG_PANEL, foreground=TEXT_MID, font=FNT_SMALL)
        style.configure("Head.TLabel",  background=BG_DEEP,  foreground=TEXT_HI,  font=FNT_UI_L)
        style.configure("Sub.TLabel",   background=BG_PANEL, foreground=TEXT_MID, font=FNT_UI)

        style.configure("TNotebook", background=BG_DEEP, borderwidth=0, tabmargins=[0,0,0,0])
        style.configure("TNotebook.Tab",
            background=BG_RAISED, foreground=TEXT_MID, font=FNT_UI_B,
            padding=[18, 8], borderwidth=0)
        style.map("TNotebook.Tab",
            background=[("selected", ACCENT_DIM)],
            foreground=[("selected", TEXT_HI)])

        style.configure("TEntry",
            fieldbackground=BG_RAISED, foreground=TEXT_HI,
            insertcolor=TEXT_HI, bordercolor=BG_BORDER,
            lightcolor=BG_BORDER, darkcolor=BG_BORDER,
            font=FNT_UI, padding=6)

        style.configure("TSpinbox",
            fieldbackground=BG_RAISED, foreground=TEXT_HI,
            insertcolor=TEXT_HI, arrowcolor=TEXT_MID,
            bordercolor=BG_BORDER, lightcolor=BG_BORDER,
            darkcolor=BG_BORDER, font=FNT_UI, padding=6)

        style.configure("TCombobox",
            fieldbackground=BG_RAISED, foreground=TEXT_HI,
            selectbackground=ACCENT_DIM, bordercolor=BG_BORDER,
            font=FNT_UI)

        style.configure("TScrollbar",
            background=BG_RAISED, troughcolor=BG_PANEL,
            arrowcolor=TEXT_MID, bordercolor=BG_PANEL)

        style.configure("TSeparator", background=BG_BORDER)

    # ── Layout skeleton ────────────────────────────────────────────────────────
    def _build_layout(self):
        # ── Header bar
        hdr = tk.Frame(self, bg=BG_PANEL, height=56)
        hdr.pack(fill=tk.X, side=tk.TOP)
        hdr.pack_propagate(False)

        tk.Label(hdr, text="⬡", bg=BG_PANEL, fg=ACCENT,
                 font=("Segoe UI", 22, "bold")).pack(side=tk.LEFT, padx=(18, 6), pady=10)
        tk.Label(hdr, text="CATIA Pipe Automation Suite",
                 bg=BG_PANEL, fg=TEXT_HI, font=FNT_TITLE).pack(side=tk.LEFT, pady=10)
        tk.Label(hdr, text="v2.0  |  CATIA V5 Automation",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FNT_SMALL).pack(side=tk.RIGHT, padx=18)

        sep = ttk.Separator(self, orient="horizontal")
        sep.pack(fill=tk.X)

        # ── Body: tabs on left, content on right via notebook
        body = ttk.Frame(self, style="Deep.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        nb = ttk.Notebook(body)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # ── Tab 1: Extract & Rebuild
        tab1 = ttk.Frame(nb)
        nb.add(tab1, text="  ① Extract & Rebuild  ")
        self._build_tab_extract(tab1)

        # ── Tab 2: Build from Coordinates
        tab2 = ttk.Frame(nb)
        nb.add(tab2, text="  ② Build from Coords  ")
        self._build_tab_coords(tab2)

        # ── Tab 3: Interactive Builder
        tab3 = ttk.Frame(nb)
        nb.add(tab3, text="  ③ Interactive Builder  ")
        self._build_tab_interactive(tab3)

        # ── Status bar
        sbar = tk.Frame(self, bg=BG_RAISED, height=26)
        sbar.pack(fill=tk.X, side=tk.BOTTOM)
        sbar.pack_propagate(False)
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(sbar, textvariable=self._status_var,
                 bg=BG_RAISED, fg=TEXT_MID,
                 font=FNT_SMALL, anchor="w").pack(side=tk.LEFT, padx=10)
        self._dot = tk.Label(sbar, text="●", bg=BG_RAISED, fg=SUCCESS, font=FNT_SMALL)
        self._dot.pack(side=tk.RIGHT, padx=10)

    # ── Shared console widget factory ─────────────────────────────────────────
    def _make_console(self, parent):
        frame = tk.Frame(parent, bg=BG_DEEP, bd=0)
        txt = tk.Text(frame, bg="#060B14", fg=TEXT_HI,
                      font=FNT_CODE, insertbackground=TEXT_HI,
                      relief=tk.FLAT, padx=10, pady=8,
                      selectbackground=ACCENT_DIM,
                      wrap=tk.WORD, state=tk.DISABLED)
        scr = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=scr.set)
        scr.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Colour tags
        txt.tag_configure("head",    foreground=ACCENT)
        txt.tag_configure("success", foreground=SUCCESS)
        txt.tag_configure("warn",    foreground=WARNING)
        txt.tag_configure("error",   foreground=ERROR)
        txt.tag_configure("dim",     foreground=TEXT_DIM)
        txt.tag_configure("info",    foreground=TEXT_HI)
        return frame, txt

    def _console_write(self, txt_widget, text, tag="info"):
        txt_widget.configure(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        txt_widget.insert(tk.END, f"[{ts}] {text}\n", tag)
        txt_widget.see(tk.END)
        txt_widget.configure(state=tk.DISABLED)

    def _clear_console(self, txt_widget):
        txt_widget.configure(state=tk.NORMAL)
        txt_widget.delete("1.0", tk.END)
        txt_widget.configure(state=tk.DISABLED)

    # ── Shared run-button factory ──────────────────────────────────────────────
    def _run_btn(self, parent, text, cmd, color=ACCENT):
        btn = tk.Button(parent, text=text, command=cmd,
                        bg=color, fg=TEXT_HI, activebackground=ACCENT_DIM,
                        activeforeground=TEXT_HI, relief=tk.FLAT, bd=0,
                        font=FNT_UI_B, padx=20, pady=8, cursor="hand2")
        return btn

    def _clear_btn(self, parent, txt_widget):
        return tk.Button(parent, text="Clear Log",
                         command=lambda: self._clear_console(txt_widget),
                         bg=BG_RAISED, fg=TEXT_MID, activebackground=BG_BORDER,
                         activeforeground=TEXT_HI, relief=tk.FLAT, bd=0,
                         font=FNT_UI, padx=14, pady=8, cursor="hand2")

    # ── Helper label + entry row ───────────────────────────────────────────────
    def _lbl_entry(self, parent, row, label, textvariable, width=14, col=0):
        tk.Label(parent, text=label, bg=BG_PANEL, fg=TEXT_MID, font=FNT_UI,
                 anchor="e").grid(row=row, column=col, sticky="e", padx=(0, 8), pady=4)
        e = ttk.Entry(parent, textvariable=textvariable, width=width)
        e.grid(row=row, column=col+1, sticky="w", pady=4)
        return e

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 1 — EXTRACT & REBUILD
    # ══════════════════════════════════════════════════════════════════════════
    def _build_tab_extract(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=6, sashpad=0, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # ── LEFT: settings panel ───────────────────────────────────────────────
        left = tk.Frame(pane, bg=BG_PANEL, width=280)
        pane.add(left, minsize=240)

        tk.Label(left, text="Configuration", bg=BG_PANEL,
                 fg=ACCENT, font=FNT_UI_L).pack(anchor="w", padx=16, pady=(14, 6))

        cfg_frm = tk.Frame(left, bg=BG_PANEL)
        cfg_frm.pack(fill=tk.X, padx=16, pady=4)
        cfg_frm.columnconfigure(1, weight=1)

        self._e_radius       = tk.StringVar(value="15.0")
        self._e_straight_tol = tk.StringVar(value="0.5")
        self._e_dedup_tol    = tk.StringVar(value="5.0")

        self._lbl_entry(cfg_frm, 0, "Pipe Radius (mm)",        self._e_radius,       col=0)
        self._lbl_entry(cfg_frm, 1, "Straightness Tol (mm)",   self._e_straight_tol, col=0)
        self._lbl_entry(cfg_frm, 2, "Dedup Distance (mm)",     self._e_dedup_tol,    col=0)

        sep = tk.Frame(left, bg=BG_BORDER, height=1)
        sep.pack(fill=tk.X, padx=16, pady=12)

        tk.Label(left,
            text=(
                "① Open a pipe model in CATIA V5.\n\n"
                "② Adjust parameters above.\n\n"
                "③ Click Run Pipeline.\n\n"
                "The script will:\n"
                "  • Extract cylinder faces\n"
                "  • Compute & deduplicate\n    centerline axes\n"
                "  • Bridge bends\n"
                "  • Sweep a solid Rib\n"
                "  • Save the document"
            ),
            bg=BG_PANEL, fg=TEXT_MID, font=FNT_SMALL,
            justify=tk.LEFT, anchor="nw", wraplength=220
        ).pack(anchor="nw", padx=16, pady=4)

        btn_frame = tk.Frame(left, bg=BG_PANEL)
        btn_frame.pack(fill=tk.X, padx=16, pady=12, side=tk.BOTTOM)
        self._e_console1_ref = None  # set after console is built

        self._btn_extract = self._run_btn(btn_frame, "▶  Run Pipeline",
                                          self._run_extract)
        self._btn_extract.pack(fill=tk.X, pady=(0, 6))

        # ── RIGHT: console ─────────────────────────────────────────────────────
        right = tk.Frame(pane, bg=BG_DEEP)
        pane.add(right, minsize=400)

        hdr2 = tk.Frame(right, bg=BG_PANEL, height=38)
        hdr2.pack(fill=tk.X)
        hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con1_frame, self._con1 = self._make_console(right)
        self._con1_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        clr = self._clear_btn(hdr2, self._con1)
        clr.pack(side=tk.RIGHT, padx=8, pady=4)

        self._console_write(self._con1,
            "CATIA Pipe Automation Suite — Extract & Rebuild Mode", "head")
        self._console_write(self._con1,
            "Open a pipe model in CATIA V5, configure parameters, then click Run.", "dim")

    def _run_extract(self):
        if self._running: return
        try:
            radius  = float(self._e_radius.get())
            s_tol   = float(self._e_straight_tol.get())
            d_tol   = float(self._e_dedup_tol.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numeric values.")
            return
        self._clear_console(self._con1)
        self._set_running(True)
        console = self._con1

        def worker():
            def log(msg, tag="info"):
                self._log_queue.put((console, msg, tag))
            try:
                run_extract_rebuild(log, pipe_radius=radius,
                                    straightness_tol=s_tol, dedup_tol=d_tol)
            finally:
                self._log_queue.put(("__done__",))

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 2 — BUILD FROM COORDINATES
    # ══════════════════════════════════════════════════════════════════════════
    def _build_tab_coords(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=6, sashpad=0, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # ── LEFT: node table ───────────────────────────────────────────────────
        left = tk.Frame(pane, bg=BG_PANEL, width=380)
        pane.add(left, minsize=340)

        tk.Label(left, text="Route Nodes", bg=BG_PANEL,
                 fg=ACCENT, font=FNT_UI_L).pack(anchor="w", padx=16, pady=(14, 4))
        tk.Label(left, text="All dimensions in millimetres",
                 bg=BG_PANEL, fg=TEXT_DIM, font=FNT_SMALL).pack(anchor="w", padx=16)

        # Column headers
        hdr = tk.Frame(left, bg=BG_RAISED)
        hdr.pack(fill=tk.X, padx=10, pady=(8, 0))
        for i, col in enumerate(["Pt", "X", "Y", "Z", "OD", "ID"]):
            tk.Label(hdr, text=col, bg=BG_RAISED, fg=ACCENT,
                     font=FNT_SMALL, width=7, anchor="center").grid(
                row=0, column=i, padx=2, pady=4)

        # Scrollable node rows
        canvas = tk.Canvas(left, bg=BG_PANEL, bd=0, highlightthickness=0)
        scr    = ttk.Scrollbar(left, orient="vertical", command=canvas.yview)
        self._node_frame = tk.Frame(canvas, bg=BG_PANEL)
        self._node_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._node_frame, anchor="nw")
        canvas.configure(yscrollcommand=scr.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        scr.pack(side=tk.RIGHT, fill=tk.Y)

        self._node_vars = []
        for row_data in DEFAULT_HOSE:
            self._add_node_row(row_data)

        btn_bar = tk.Frame(left, bg=BG_PANEL)
        btn_bar.pack(fill=tk.X, padx=10, pady=6)
        tk.Button(btn_bar, text="+ Add Node", command=self._add_node_row,
                  bg=BG_RAISED, fg=SUCCESS, activebackground=BG_BORDER,
                  activeforeground=TEXT_HI, font=FNT_UI, relief=tk.FLAT,
                  cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_bar, text="− Remove Last", command=self._remove_node_row,
                  bg=BG_RAISED, fg=ERROR, activebackground=BG_BORDER,
                  activeforeground=TEXT_HI, font=FNT_UI, relief=tk.FLAT,
                  cursor="hand2").pack(side=tk.LEFT)

        sep = tk.Frame(left, bg=BG_BORDER, height=1)
        sep.pack(fill=tk.X, padx=10, pady=6)

        run_bar = tk.Frame(left, bg=BG_PANEL)
        run_bar.pack(fill=tk.X, padx=10, pady=(0, 10))
        self._run_coords_btn = self._run_btn(run_bar, "▶  Build Hose in CATIA",
                                              self._run_coords)
        self._run_coords_btn.pack(fill=tk.X)

        # ── RIGHT: console ─────────────────────────────────────────────────────
        right = tk.Frame(pane, bg=BG_DEEP)
        pane.add(right, minsize=380)

        hdr2 = tk.Frame(right, bg=BG_PANEL, height=38)
        hdr2.pack(fill=tk.X)
        hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con2_frame, self._con2 = self._make_console(right)
        self._con2_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        clr = self._clear_btn(hdr2, self._con2)
        clr.pack(side=tk.RIGHT, padx=8, pady=4)

        self._console_write(self._con2,
            "CATIA Pipe Automation Suite — Build from Coordinates", "head")
        self._console_write(self._con2,
            "Define your nodes in the table, then click Build.", "dim")

    def _add_node_row(self, data=None):
        idx = len(self._node_vars) + 1
        if data is None:
            data = {"point": idx, "x": 0.0, "y": 0.0, "z": 0.0,
                    "outer_dia": 25.0, "inner_dia": 20.0}

        row = len(self._node_vars)
        bg  = BG_PANEL if row % 2 == 0 else BG_RAISED

        vars_ = {
            "x": tk.StringVar(value=str(data["x"])),
            "y": tk.StringVar(value=str(data["y"])),
            "z": tk.StringVar(value=str(data["z"])),
            "od": tk.StringVar(value=str(data["outer_dia"])),
            "id": tk.StringVar(value=str(data["inner_dia"])),
        }

        f = tk.Frame(self._node_frame, bg=bg)
        f.pack(fill=tk.X, pady=1)

        tk.Label(f, text=str(data["point"]), bg=bg, fg=TEXT_MID,
                 font=FNT_SMALL, width=3, anchor="center").pack(side=tk.LEFT, padx=2)
        for key in ["x", "y", "z", "od", "id"]:
            e = tk.Entry(f, textvariable=vars_[key], width=7,
                         bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                         relief=tk.FLAT, font=FNT_CODE, bd=0)
            e.pack(side=tk.LEFT, padx=2, pady=2)

        self._node_vars.append(vars_)

    def _remove_node_row(self):
        if len(self._node_vars) <= 2: return
        self._node_vars.pop()
        children = self._node_frame.winfo_children()
        if children: children[-1].destroy()

    def _collect_hose_data(self):
        result = []
        for i, v in enumerate(self._node_vars):
            try:
                result.append({
                    "point":     i + 1,
                    "x":         float(v["x"].get()),
                    "y":         float(v["y"].get()),
                    "z":         float(v["z"].get()),
                    "outer_dia": float(v["od"].get()),
                    "inner_dia": float(v["id"].get()),
                })
            except ValueError:
                messagebox.showerror("Invalid Input",
                    f"Row {i+1} contains non-numeric values.")
                return None
        return result

    def _run_coords(self):
        if self._running: return
        data = self._collect_hose_data()
        if data is None: return
        self._clear_console(self._con2)
        self._set_running(True)
        console = self._con2

        def worker():
            def log(msg, tag="info"):
                self._log_queue.put((console, msg, tag))
            try:
                run_build_from_coords(log, data)
            finally:
                self._log_queue.put(("__done__",))

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════════
    #  TAB 3 — INTERACTIVE BUILDER
    # ══════════════════════════════════════════════════════════════════════════
    def _build_tab_interactive(self, parent):
        parent.configure(style="TFrame")
        pane = tk.PanedWindow(parent, orient=tk.HORIZONTAL,
                              bg=BG_DEEP, sashwidth=6, sashpad=0, sashrelief=tk.FLAT)
        pane.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # ── LEFT: form ─────────────────────────────────────────────────────────
        left = tk.Frame(pane, bg=BG_PANEL, width=380)
        pane.add(left, minsize=340)

        tk.Label(left, text="Pipe Definition", bg=BG_PANEL,
                 fg=ACCENT, font=FNT_UI_L).pack(anchor="w", padx=16, pady=(14, 4))

        # Global settings
        gset = tk.LabelFrame(left, text=" Global Settings ",
                             bg=BG_PANEL, fg=TEXT_MID, font=FNT_SMALL,
                             bd=1, relief=tk.RIDGE)
        gset.pack(fill=tk.X, padx=12, pady=(6, 0))
        gset.columnconfigure(1, weight=1)

        self._iv_name  = tk.StringVar(value="Custom_Pipe")
        self._iv_od    = tk.StringVar(value="25.0")
        self._iv_id    = tk.StringVar(value="20.0")

        self._lbl_entry(gset, 0, "Part Name",         self._iv_name, width=18)
        self._lbl_entry(gset, 1, "Default OD (mm)",   self._iv_od,   width=10)
        self._lbl_entry(gset, 2, "Default ID (mm)",   self._iv_id,   width=10)

        # Segment list
        seg_hdr = tk.Frame(left, bg=BG_PANEL)
        seg_hdr.pack(fill=tk.X, padx=12, pady=(10, 2))
        tk.Label(seg_hdr, text="Segments", bg=BG_PANEL,
                 fg=TEXT_HI, font=FNT_UI_B).pack(side=tk.LEFT)
        tk.Button(seg_hdr, text="+ Segment",
                  command=self._add_seg_row,
                  bg=BG_RAISED, fg=SUCCESS, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.RIGHT)
        tk.Button(seg_hdr, text="− Remove",
                  command=self._remove_seg_row,
                  bg=BG_RAISED, fg=ERROR, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.RIGHT, padx=4)

        # Segments scroll
        seg_canvas = tk.Canvas(left, bg=BG_PANEL, bd=0, highlightthickness=0, height=260)
        seg_scr    = ttk.Scrollbar(left, orient="vertical", command=seg_canvas.yview)
        self._seg_frame = tk.Frame(seg_canvas, bg=BG_PANEL)
        self._seg_frame.bind("<Configure>",
            lambda e: seg_canvas.configure(scrollregion=seg_canvas.bbox("all")))
        seg_canvas.create_window((0, 0), window=self._seg_frame, anchor="nw")
        seg_canvas.configure(yscrollcommand=seg_scr.set)
        seg_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 0))
        seg_scr.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4))

        self._seg_vars = []
        # Default: 3 segments, 4 points
        for i in range(3):
            self._add_seg_row()

        # Actions
        act = tk.Frame(left, bg=BG_PANEL)
        act.pack(fill=tk.X, padx=12, pady=(8, 6))

        tk.Button(act, text="📂  Load JSON",
                  command=self._load_json,
                  bg=BG_RAISED, fg=TEXT_MID, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(act, text="💾  Save JSON",
                  command=self._save_json,
                  bg=BG_RAISED, fg=TEXT_MID, relief=tk.FLAT, font=FNT_UI,
                  cursor="hand2").pack(side=tk.LEFT)

        sep = tk.Frame(left, bg=BG_BORDER, height=1)
        sep.pack(fill=tk.X, padx=12, pady=6)

        self._run_ib_btn = self._run_btn(left, "▶  Build Pipe in CATIA",
                                          self._run_interactive)
        self._run_ib_btn.pack(fill=tk.X, padx=12, pady=(0, 10))

        # ── RIGHT: console ─────────────────────────────────────────────────────
        right = tk.Frame(pane, bg=BG_DEEP)
        pane.add(right, minsize=380)

        hdr2 = tk.Frame(right, bg=BG_PANEL, height=38)
        hdr2.pack(fill=tk.X)
        hdr2.pack_propagate(False)
        tk.Label(hdr2, text="Console Output", bg=BG_PANEL,
                 fg=TEXT_MID, font=FNT_UI_B).pack(side=tk.LEFT, padx=12, pady=8)
        self._con3_frame, self._con3 = self._make_console(right)
        self._con3_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        clr = self._clear_btn(hdr2, self._con3)
        clr.pack(side=tk.RIGHT, padx=8, pady=4)

        self._console_write(self._con3,
            "CATIA Pipe Automation Suite — Interactive Builder", "head")
        self._console_write(self._con3,
            "Define segments in the form, then click Build.", "dim")

    def _add_seg_row(self):
        idx = len(self._seg_vars) + 1
        row = len(self._seg_vars)
        bg  = BG_PANEL if row % 2 == 0 else BG_RAISED

        f = tk.LabelFrame(self._seg_frame, text=f" Segment {idx} ",
                          bg=bg, fg=ACCENT, font=FNT_SMALL,
                          bd=1, relief=tk.GROOVE)
        f.pack(fill=tk.X, padx=4, pady=4)
        f.columnconfigure(1, weight=1); f.columnconfigure(3, weight=1)

        v = {
            "p1x": tk.StringVar(value="0.0"),
            "p1y": tk.StringVar(value="0.0"),
            "p1z": tk.StringVar(value="0.0"),
            "p2x": tk.StringVar(value="100.0"),
            "p2y": tk.StringVar(value="0.0"),
            "p2z": tk.StringVar(value="0.0"),
            "type": tk.StringVar(value="straight"),
            "bend": tk.StringVar(value=""),
            "od":   tk.StringVar(value="25.0"),
            "id":   tk.StringVar(value="20.0"),
        }

        tk.Label(f, text="Start (X Y Z)", bg=bg, fg=TEXT_MID, font=FNT_SMALL).grid(
            row=0, column=0, sticky="e", padx=(6,2), pady=2)
        for ci, k in enumerate(["p1x","p1y","p1z"]):
            tk.Entry(f, textvariable=v[k], width=7,
                     bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                     relief=tk.FLAT, font=FNT_CODE).grid(row=0, column=ci+1, padx=2, pady=2)

        tk.Label(f, text="End (X Y Z)", bg=bg, fg=TEXT_MID, font=FNT_SMALL).grid(
            row=1, column=0, sticky="e", padx=(6,2), pady=2)
        for ci, k in enumerate(["p2x","p2y","p2z"]):
            tk.Entry(f, textvariable=v[k], width=7,
                     bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                     relief=tk.FLAT, font=FNT_CODE).grid(row=1, column=ci+1, padx=2, pady=2)

        tk.Label(f, text="Type", bg=bg, fg=TEXT_MID, font=FNT_SMALL).grid(
            row=2, column=0, sticky="e", padx=(6,2), pady=2)
        cb = ttk.Combobox(f, textvariable=v["type"], values=["straight","curved"],
                          width=9, state="readonly")
        cb.grid(row=2, column=1, columnspan=2, sticky="w", pady=2)

        tk.Label(f, text="Bend R (mm)", bg=bg, fg=TEXT_MID, font=FNT_SMALL).grid(
            row=2, column=3, sticky="e", padx=(8,2))
        tk.Entry(f, textvariable=v["bend"], width=7,
                 bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                 relief=tk.FLAT, font=FNT_CODE).grid(row=2, column=4, padx=2, pady=2)

        tk.Label(f, text="OD / ID (mm)", bg=bg, fg=TEXT_MID, font=FNT_SMALL).grid(
            row=3, column=0, sticky="e", padx=(6,2), pady=(2,6))
        tk.Entry(f, textvariable=v["od"], width=7,
                 bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                 relief=tk.FLAT, font=FNT_CODE).grid(row=3, column=1, pady=(2,6))
        tk.Label(f, text="/", bg=bg, fg=TEXT_DIM).grid(row=3, column=2)
        tk.Entry(f, textvariable=v["id"], width=7,
                 bg=BG_RAISED, fg=TEXT_HI, insertbackground=TEXT_HI,
                 relief=tk.FLAT, font=FNT_CODE).grid(row=3, column=3, pady=(2,6))

        self._seg_vars.append({"vars": v, "frame": f})

    def _remove_seg_row(self):
        if len(self._seg_vars) <= 1: return
        entry = self._seg_vars.pop()
        entry["frame"].destroy()

    def _collect_pipe_def(self):
        try:
            part_name = self._iv_name.get().strip() or "Custom_Pipe"
            def_od    = float(self._iv_od.get())
            def_id    = float(self._iv_id.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Check global settings.")
            return None

        nodes = []; segments = []

        for i, entry in enumerate(self._seg_vars):
            v = entry["vars"]
            try:
                p1 = (float(v["p1x"].get()), float(v["p1y"].get()), float(v["p1z"].get()))
                p2 = (float(v["p2x"].get()), float(v["p2y"].get()), float(v["p2z"].get()))
                od = float(v["od"].get())
                id_ = float(v["id"].get())
                seg_type = v["type"].get()
                bend_r   = float(v["bend"].get()) if v["bend"].get().strip() else None
            except ValueError:
                messagebox.showerror("Invalid Input", f"Segment {i+1} has invalid values.")
                return None

            chord = math.sqrt(sum((b-a)**2 for a,b in zip(p1,p2)))

            # Add start node (skip if same as previous end)
            if not nodes or (nodes[-1]["x"], nodes[-1]["y"], nodes[-1]["z"]) != p1:
                nodes.append({"point": len(nodes)+1, "x": p1[0], "y": p1[1], "z": p1[2],
                               "outer_dia": od, "inner_dia": id_})
            # Add end node
            nodes.append({"point": len(nodes)+1, "x": p2[0], "y": p2[1], "z": p2[2],
                           "outer_dia": od, "inner_dia": id_})

            segments.append({
                "segment": i+1, "from_point": i+1, "to_point": i+2,
                "type": seg_type, "bend_radius": bend_r,
                "outer_dia": od, "inner_dia": id_,
                "chord_length": round(chord, 4)
            })

        # Deduplicate consecutive nodes
        unique_nodes = [nodes[0]]
        for n in nodes[1:]:
            if (n["x"], n["y"], n["z"]) != (unique_nodes[-1]["x"],
                                             unique_nodes[-1]["y"],
                                             unique_nodes[-1]["z"]):
                unique_nodes.append(n)
        for i, n in enumerate(unique_nodes): n["point"] = i + 1

        return {
            "meta": {
                "created": datetime.now().isoformat(timespec="seconds"),
                "part_name": part_name,
                "num_points": len(unique_nodes),
                "num_segments": len(segments),
                "default_outer_dia": def_od,
                "default_inner_dia": def_id,
            },
            "nodes": unique_nodes,
            "segments": segments,
        }

    def _load_json(self):
        path = filedialog.askopenfilename(
            title="Load Pipe Definition",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path: return
        try:
            with open(path) as f: data = json.load(f)
            messagebox.showinfo("Loaded",
                f"Loaded: {data['meta']['part_name']}\n"
                f"Points: {data['meta']['num_points']}  "
                f"Segments: {data['meta']['num_segments']}")
            self._iv_name.set(data["meta"]["part_name"])
            self._iv_od.set(str(data["meta"]["default_outer_dia"]))
            self._iv_id.set(str(data["meta"]["default_inner_dia"]))
        except Exception as e:
            messagebox.showerror("Error", f"Could not load JSON:\n{e}")

    def _save_json(self):
        pipe_def = self._collect_pipe_def()
        if pipe_def is None: return
        path = filedialog.asksaveasfilename(
            title="Save Pipe Definition",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        if not path: return
        try:
            with open(path, "w") as f: json.dump(pipe_def, f, indent=2)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save:\n{e}")

    def _run_interactive(self):
        if self._running: return
        pipe_def = self._collect_pipe_def()
        if pipe_def is None: return

        self._clear_console(self._con3)
        self._set_running(True)
        console = self._con3

        self._console_write(console, f"Part Name : {pipe_def['meta']['part_name']}", "head")
        self._console_write(console, f"Nodes     : {pipe_def['meta']['num_points']}", "head")
        self._console_write(console, f"Segments  : {pipe_def['meta']['num_segments']}", "head")

        def worker():
            def log(msg, tag="info"):
                self._log_queue.put((console, msg, tag))
            try:
                run_interactive_builder(log, pipe_def)
            finally:
                self._log_queue.put(("__done__",))

        threading.Thread(target=worker, daemon=True).start()

    # ── Thread / status helpers ────────────────────────────────────────────────
    def _set_running(self, running):
        self._running = running
        state = "disabled" if running else "normal"
        color = BG_BORDER if running else ACCENT
        dot_c = WARNING if running else SUCCESS
        for btn in [self._btn_extract, self._run_coords_btn, self._run_ib_btn]:
            btn.configure(state=state, bg=color)
        self._dot.configure(fg=dot_c)
        self._status_var.set("Running pipeline…" if running else "Ready")

    def _poll_log(self):
        try:
            while True:
                item = self._log_queue.get_nowait()
                if item[0] == "__done__":
                    self._set_running(False)
                else:
                    console, msg, tag = item
                    self._console_write(console, msg, tag)
        except queue.Empty:
            pass
        self.after(60, self._poll_log)


# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = CatiaPipeApp()
    app.mainloop()