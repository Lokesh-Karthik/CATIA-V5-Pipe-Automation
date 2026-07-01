import win32com.client
import math

# =============================================================================
#  CATIA INTERACTIVE CENTERLINE EXTRACTOR & RIB GENERATOR
#  Phase 1 : Interactively extract straight cylinder faces from the 3D model
#  Phase 2 : Compute centerlines, deduplicate, sort tip-to-tail
#  Phase 3 : Bridge bends, join into master curve, create solid Rib
# =============================================================================

def run_full_pipeline():
    print("=" * 60)
    print(" CATIA INTERACTIVE CENTERLINE EXTRACTOR & RIB GENERATOR")
    print("=" * 60)

    # =========================================================================
    # COMMON SETUP
    # =========================================================================
    print("\n[SETUP] Connecting to CATIA...")
    try:
        catia             = win32com.client.Dispatch("CATIA.Application")
        part_document     = catia.ActiveDocument          
        part              = part_document.Part            
        hs_factory        = part.HybridShapeFactory       
        shape_factory     = part.ShapeFactory             
        spa_workbench     = part_document.GetWorkbench("SPAWorkbench")  
        selection         = part_document.Selection       
        hybrid_bodies     = part.HybridBodies             
    except Exception as e:
        print(f"[SETUP] FAILED: Could not connect to CATIA.\n  -> {e}")
        return

    print("[SETUP] Connected successfully.\n")

    # =========================================================================
    # PHASE 1 — INTERACTIVE FACE EXTRACTION
    # =========================================================================
    print("-" * 60)
    print(" PHASE 1: Interactive Face Extraction")
    print("-" * 60)

    try:
        old_set = hybrid_bodies.Item("Extracted_Cylinders")
        selection.Clear()
        selection.Add(old_set)
        selection.Delete()
    except:
        pass  

    extraction_set      = hybrid_bodies.Add()
    extraction_set.Name = "Extracted_Cylinders"

    print("[Phase 1] Waiting for user selection in CATIA...")
    print("          Go to CATIA and select the STRAIGHT CYLINDRICAL FACES of the pipe.")
    print("          (Hold CTRL to select multiple straight segments. Do NOT select the curved bends).")
    input("          Press ENTER here in the console once selected...")

    if selection.Count < 2:
        print("[Phase 1] FAILED: You must select at least 2 straight faces to route a pipe.")
        return

    extracted_count = 0
    
    # Safely copy references to memory to avoid context loss
    refs = []
    for i in range(1, selection.Count + 1):
        refs.append(selection.Item(i).Reference)
    selection.Clear()

    # Extract the faces WITHOUT tangency so they remain isolated straight segments
    for i, ref in enumerate(refs):
        extraction = hs_factory.AddNewExtract(ref)
        extraction.Name = f"Cyl_Face_{i+1}"
        extraction.PropagationType = 0  # 0 = No tangency propagation
        extraction_set.AppendHybridShape(extraction)
        extracted_count += 1

    part.Update()
    print(f"[Phase 1] Successfully extracted {extracted_count} straight faces.")

    # =========================================================================
    # PHASE 2 — COMPUTE CENTERLINES, DEDUPLICATE, SORT
    # =========================================================================
    print("\n" + "-" * 60)
    print(" PHASE 2: Computing centerlines")
    print("-" * 60)

    try:
        old_cl = hybrid_bodies.Item("Extracted_Centerlines")
        selection.Clear()
        selection.Add(old_cl)
        selection.Delete()
    except:
        pass

    centerline_set      = hybrid_bodies.Add()
    centerline_set.Name = "Extracted_Centerlines"

    print("[Phase 2] Analyzing each face to find straight axis lines...")
    faces    = extraction_set.HybridShapes   
    all_lines = []

    for i in range(1, faces.Count + 1):
        current_face = faces.Item(i)

        selection.Clear()
        selection.Add(current_face)
        try:
            selection.Search("Topology.Edge,sel")  
        except:
            continue

        if selection.Count == 0:
            continue

        points_data = []
        for j in range(1, selection.Count + 1):
            try:
                raw_edge_ref  = selection.Item(j).Reference
                edge_extract  = hs_factory.AddNewExtract(raw_edge_ref)
                centerline_set.AppendHybridShape(edge_extract)

                edge_ref_obj  = part.CreateReferenceFromObject(edge_extract)
                pt            = hs_factory.AddNewPointCenter(edge_ref_obj)
                centerline_set.AppendHybridShape(pt)
                part.UpdateObject(pt)
                points_data.append(pt)
            except:
                continue

        if len(points_data) < 2:
            continue

        max_dist        = -1
        best_p1, best_p2 = None, None

        for p_a in points_data:
            ref_a  = part.CreateReferenceFromObject(p_a)
            meas_a = spa_workbench.GetMeasurable(ref_a)
            for p_b in points_data:
                ref_b = part.CreateReferenceFromObject(p_b)
                try:
                    dist = meas_a.GetMinimumDistance(ref_b)
                    if dist > max_dist:
                        max_dist = dist
                        best_p1  = p_a
                        best_p2  = p_b
                except:
                    pass

        if not best_p1 or not best_p2 or max_dist <= 1.0:
            continue  

        line = hs_factory.AddNewLinePtPt(
            part.CreateReferenceFromObject(best_p1),
            part.CreateReferenceFromObject(best_p2)
        )
        centerline_set.AppendHybridShape(line)
        part.UpdateObject(line)

        midpoint = hs_factory.AddNewPointOnCurveFromPercent(
            part.CreateReferenceFromObject(line), 0.5, True
        )
        centerline_set.AppendHybridShape(midpoint)
        part.UpdateObject(midpoint)

        face_ref = part.CreateReferenceFromObject(current_face)
        meas_mid = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(midpoint))
        dist_mid = meas_mid.GetMinimumDistance(face_ref)
        meas_p1  = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(best_p1))
        dist_p1  = meas_p1.GetMinimumDistance(face_ref)

        STRAIGHTNESS_TOL = 0.5
        if abs(dist_mid - dist_p1) < STRAIGHTNESS_TOL:
            all_lines.append({
                "line":     line,
                "p1":       best_p1,
                "p2":       best_p2,
                "midpoint": midpoint
            })
        else:
            selection.Clear()
            selection.Add(line)
            selection.Add(midpoint)
            try: selection.Delete()
            except: pass

    print(f"[Phase 2] Generated {len(all_lines)} raw axis lines.")

    if not all_lines:
        print("[Phase 2] FAILED: No straight lines found.")
        return

    print("[Phase 2] Deduplicating overlapping axis lines...")
    DEDUP_TOL   = 5.0   
    unique_lines = []

    for line_data in all_lines:
        ref_mid  = part.CreateReferenceFromObject(line_data["midpoint"])
        meas_mid = spa_workbench.GetMeasurable(ref_mid)
        is_dup   = False

        for u_line in unique_lines:
            ref_u = part.CreateReferenceFromObject(u_line["midpoint"])
            try:
                dist = meas_mid.GetMinimumDistance(ref_u)
                if dist < DEDUP_TOL:
                    is_dup = True
                    selection.Clear()
                    selection.Add(line_data["line"])
                    selection.Add(line_data["p1"])
                    selection.Add(line_data["p2"])
                    selection.Add(line_data["midpoint"])
                    try: selection.Delete()
                    except: pass
                    break
            except:
                pass

        if not is_dup:
            unique_lines.append(line_data)

    for u_line in unique_lines:
        selection.Clear()
        selection.Add(u_line["midpoint"])
        try: selection.Delete()
        except: pass

    print(f"[Phase 2] Reduced to {len(unique_lines)} unique pipe segments.")

    if len(unique_lines) < 2:
        print("[Phase 2] FAILED: Need at least 2 segments to build a route.")
        return

    print("[Phase 2] Sorting segments tip-to-tail...")
    meas_ref  = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(unique_lines[0]["p1"]))
    max_d      = -1
    start_line = unique_lines[0]

    for ld in unique_lines:
        try:
            d = meas_ref.GetMinimumDistance(part.CreateReferenceFromObject(ld["line"]))
            if d > max_d:
                max_d      = d
                start_line = ld
        except:
            pass

    sorted_lines = [start_line]
    unvisited    = [ld for ld in unique_lines if ld is not start_line]
    active_tip = start_line["p1"]   

    if unvisited:
        meas_p1   = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(start_line["p1"]))
        meas_p2   = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(start_line["p2"]))
        min_d1 = min_d2 = float("inf")

        for ld in unvisited:
            ref_line = part.CreateReferenceFromObject(ld["line"])
            try:
                d1 = meas_p1.GetMinimumDistance(ref_line)
                d2 = meas_p2.GetMinimumDistance(ref_line)
                if d1 < min_d1: min_d1 = d1
                if d2 < min_d2: min_d2 = d2
            except:
                pass

        active_tip = start_line["p2"] if min_d2 < min_d1 else start_line["p1"]

    while unvisited:
        meas_tip   = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(active_tip))
        best_line  = None
        best_dist  = float("inf")
        best_p1_d  = float("inf")
        best_p2_d  = float("inf")

        for ld in unvisited:
            try:
                d1 = meas_tip.GetMinimumDistance(part.CreateReferenceFromObject(ld["p1"]))
                d2 = meas_tip.GetMinimumDistance(part.CreateReferenceFromObject(ld["p2"]))
                local_min = min(d1, d2)
                if local_min < best_dist:
                    best_dist  = local_min
                    best_line  = ld
                    best_p1_d  = d1
                    best_p2_d  = d2
            except:
                pass

        if not best_line: break

        sorted_lines.append(best_line)
        unvisited.remove(best_line)
        active_tip = best_line["p2"] if best_p1_d < best_p2_d else best_line["p1"]

    print(f"[Phase 2] Sorted {len(sorted_lines)} segments into pipe order.")

    # =========================================================================
    # PHASE 3 — BRIDGE BENDS, JOIN, CREATE SOLID RIB
    # =========================================================================
    print("\n" + "-" * 60)
    print(" PHASE 3: Bridging bends and creating solid Rib")
    print("-" * 60)

    print("[Phase 3] Creating Connect curves at bends...")
    connects = []

    for i in range(len(sorted_lines) - 1):
        l1, l2 = sorted_lines[i], sorted_lines[i + 1]
        p1_a, p1_b = l1["p1"], l1["p2"]
        p2_a, p2_b = l2["p1"], l2["p2"]

        meas_p1a = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(p1_a))
        d_aa = meas_p1a.GetMinimumDistance(part.CreateReferenceFromObject(p2_a))
        d_ab = meas_p1a.GetMinimumDistance(part.CreateReferenceFromObject(p2_b))

        meas_p1b = spa_workbench.GetMeasurable(part.CreateReferenceFromObject(p1_b))
        d_ba = meas_p1b.GetMinimumDistance(part.CreateReferenceFromObject(p2_a))
        d_bb = meas_p1b.GetMinimumDistance(part.CreateReferenceFromObject(p2_b))

        min_d = min(d_aa, d_ab, d_ba, d_bb)
        if   min_d == d_aa: pt1_obj, pt2_obj = p1_a, p2_a
        elif min_d == d_ab: pt1_obj, pt2_obj = p1_a, p2_b
        elif min_d == d_ba: pt1_obj, pt2_obj = p1_b, p2_a
        else:               pt1_obj, pt2_obj = p1_b, p2_b

        ori1 = 1 if pt1_obj.Name == l1["p2"].Name else -1
        ori2 = 1 if pt2_obj.Name == l2["p1"].Name else -1

        try:
            connect = hs_factory.AddNewConnect(
                part.CreateReferenceFromObject(l1["line"]), part.CreateReferenceFromObject(pt1_obj), ori1, 1, 1.0,
                part.CreateReferenceFromObject(l2["line"]), part.CreateReferenceFromObject(pt2_obj), ori2, 1, 1.0, False
            )
            connect.Name = f"Bend_Connect_{i + 1}"
            centerline_set.AppendHybridShape(connect)
            part.UpdateObject(connect)
            connects.append(connect)
        except Exception as e:
            print(f"  -> WARNING: Bend {i + 1} connect failed: {e}")

    print(f"[Phase 3] {len(connects)} bend(s) bridged successfully.")

    print("[Phase 3] Joining all elements into Master_Centerline_Join...")
    join_elements = [sorted_lines[0]["line"]]
    for i in range(len(connects)):
        join_elements.append(connects[i])
        join_elements.append(sorted_lines[i + 1]["line"])

    try:
        master_join = hs_factory.AddNewJoin(
            part.CreateReferenceFromObject(join_elements[0]),
            part.CreateReferenceFromObject(join_elements[1])
        )
        for elem in join_elements[2:]:
            master_join.AddElement(part.CreateReferenceFromObject(elem))

        master_join.Name = "Master_Centerline_Join"
        centerline_set.AppendHybridShape(master_join)
        part.UpdateObject(master_join)
    except Exception as e:
        print(f"[Phase 3] FAILED during Join: {e}")
        return

    join_ref = part.CreateReferenceFromObject(master_join)

    # Allow user to input custom radius to prevent Twisted Volume crashes
    rad_input = input("Enter the pipe outer radius (mm) [Default 5.0]: ").strip()
    PIPE_RADIUS = float(rad_input) if rad_input else 5.0

    try:
        pt_start = hs_factory.AddNewPointOnCurveFromPercent(join_ref, 0.0, True)
        centerline_set.AppendHybridShape(pt_start)

        plane = hs_factory.AddNewPlaneNormal(join_ref, part.CreateReferenceFromObject(pt_start))
        centerline_set.AppendHybridShape(plane)

        circle = hs_factory.AddNewCircleCtrRad(
            part.CreateReferenceFromObject(pt_start),
            part.CreateReferenceFromObject(plane),
            False,          
            PIPE_RADIUS
        )
        circle.Name = "Pipe_Profile_Circle"
        centerline_set.AppendHybridShape(circle)
        part.Update()
        print(f"[Phase 3] Profile circle created (radius = {PIPE_RADIUS} mm).")
    except Exception as e:
        print(f"[Phase 3] FAILED creating profile circle: {e}")
        return

    try:
        part.InWorkObject = part.MainBody   
        circle_ref = part.CreateReferenceFromObject(circle)
        rib        = shape_factory.AddNewRibFromRef(circle_ref, join_ref)
        part.UpdateObject(rib)
        print("[Phase 3] >>> SUCCESS: Solid Rib generated! <<<")
    except Exception as e:
        print(f"[Phase 3] FAILED during Rib creation: {e}")
        print("  -> Try entering a smaller radius to prevent self-intersection at the corners!")
        return

    print("\n" + "=" * 60)
    print(" PIPELINE COMPLETE")
    print(f"  Segments extracted : {extracted_count}")
    print(f"  Unique axis lines  : {len(unique_lines)}")
    print(f"  Bends connected    : {len(connects)}")
    print(f"  Solid pipe         : Rib in PartBody")
    print("=" * 60)

if __name__ == "__main__":
    run_full_pipeline()