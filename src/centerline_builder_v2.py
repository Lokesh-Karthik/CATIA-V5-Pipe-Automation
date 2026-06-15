"""
Centerline Builder V2 — point-based centerline reconstruction.

This is the workflow ported from the ``tube_parametric_creation_v2``
prototype into the project's win32com / ``CatiaSession`` style:

  1. User selects the outer cylindrical face of the pipe.
  2. A tangent-continuity Extract grabs the whole pipe surface (incl. bends).
  3. User selects ONE circular edge to establish the reference radius.
  4. Every edge on the extract matching that radius gets an
     ``AddNewPointCenter`` (geometry-native center point).
  5. User picks the starting point; remaining points are nearest-neighbour
     sorted.
  6. Consecutive points closer than ``arc_threshold`` are bridged with
     Connect curves (bends); farther pairs become straight Lines.
  7. All lines/connects are joined into a single centerline spine.
"""
from src.utils import setup_logger

logger = setup_logger(__name__)


# =========================================================================
#  HELPERS
# =========================================================================

def _spa_distance(part, spa, pt_a, pt_b):
    """Measure distance between two point objects using CATIA SPA.
    Tries multiple COM methods for compatibility across CATIA versions."""
    ref_a = part.CreateReferenceFromObject(pt_a)
    ref_b = part.CreateReferenceFromObject(pt_b)

    # Method 1: GetMeasurableBetween (pycatia-style)
    try:
        mb = spa.GetMeasurableBetween(ref_a, ref_b)
        return mb.MinimumDistance
    except Exception:
        pass

    # Method 2: Measurable.GetMinimumDistance(Reference)
    try:
        meas_a = spa.GetMeasurable(ref_a)
        return meas_a.GetMinimumDistance(ref_b)
    except Exception:
        pass

    # Method 3: Measurable.GetMinimumDistance(Measurable)
    try:
        meas_a = spa.GetMeasurable(ref_a)
        meas_b = spa.GetMeasurable(ref_b)
        return meas_a.GetMinimumDistance(meas_b)
    except Exception:
        pass

    return None


# =========================================================================
#  1. USER SELECTIONS
# =========================================================================

def select_outer_surface(session):
    """Prompt the user to pick the outer cylindrical face of the pipe."""
    selection = session.selection
    selection.Clear()
    logger.info("Waiting for outer tube surface selection in CATIA...")
    status = selection.SelectElement2(("Face",), "Select the outer tube surface", False)
    if status == "Cancel":
        raise RuntimeError("User cancelled surface selection.")

    face_ref = selection.Item(1).Reference
    selection.Clear()
    return face_ref


def create_multiextract(part, hs_factory, geo_set, face_ref):
    """Tangent-continuity Extract of the whole pipe surface (incl. bends)."""
    extract = hs_factory.AddNewExtract(face_ref)
    extract.PropagationType = 2  # tangent continuity
    try:
        extract.ComplementaryExtract = False
    except Exception:
        pass
    extract.Name = "Pipe_Surface_Extract"
    geo_set.AppendHybridShape(extract)
    part.UpdateObject(extract)
    logger.info("Tangent multiextract created.")
    return extract


def select_reference_radius(session, part, spa_workbench, hs_factory, geo_set):
    """Prompt the user for one circular edge and return its radius."""
    selection = session.selection
    selection.Clear()
    logger.info("Waiting for reference edge selection...")
    status = selection.SelectElement2(
        ("Edge",), "Select a circular edge of the tube (sets the reference radius)", False
    )
    if status == "Cancel":
        raise RuntimeError("User cancelled edge selection.")

    edge_ref = selection.Item(1).Reference
    edge_extract = hs_factory.AddNewExtract(edge_ref)
    edge_extract.Name = "Reference_Edge_Extract"
    geo_set.AppendHybridShape(edge_extract)
    part.UpdateObject(edge_extract)

    extract_ref = part.CreateReferenceFromObject(edge_extract)
    try:
        radius = spa_workbench.GetMeasurable(extract_ref).Radius
    except Exception as e:
        raise RuntimeError(f"Could not measure radius of selected edge: {e}") from e

    selection.Clear()
    logger.info("Reference radius = %.3f mm", radius)
    return radius


# =========================================================================
#  2. CENTER POINTS FROM MATCHING EDGES
# =========================================================================

def extract_centerpoints(part, hs_factory, spa_workbench, selection,
                          extract, geo_set, reference_radius, tolerance=0.001):
    """
    Search every edge on the multiextract; for each edge whose radius
    matches ``reference_radius`` (within ``tolerance``), create a
    geometry-native center point via ``AddNewPointCenter``.

    All created extracts and points are left in the CATIA tree untouched.
    Deduplication uses CATIA distance measurement.
    """
    selection.Clear()
    selection.Add(extract)
    try:
        selection.Search("Topology.Edge,sel")
    except Exception as e:
        raise RuntimeError(f"Edge search on extract failed: {e}") from e

    edge_count = selection.Count
    logger.info("Scanning %d edges on pipe extract (ref radius=%.3f)...",
                edge_count, reference_radius)

    center_points = []
    match_count = 0
    no_radius_count = 0

    for i in range(1, edge_count + 1):
        edge_ref = selection.Item(i).Reference

        # --- Step A: extract the individual edge ---
        try:
            edge_extract = hs_factory.AddNewExtract(edge_ref)
            geo_set.AppendHybridShape(edge_extract)
            part.Update()
        except Exception as exc:
            logger.debug("Edge %d/%d: extract failed: %s", i, edge_count, exc)
            continue

        # --- Step B: measure radius ---
        try:
            extract_ref = part.CreateReferenceFromObject(edge_extract)
            radius = spa_workbench.GetMeasurable(extract_ref).Radius
        except Exception:
            no_radius_count += 1
            continue

        # --- Step C: radius match check ---
        if abs(radius - reference_radius) >= tolerance:
            continue

        match_count += 1

        # --- Step D: create PointCenter ---
        try:
            pt = hs_factory.AddNewPointCenter(extract_ref)
            geo_set.AppendHybridShape(pt)
            part.Update()
        except Exception as exc:
            logger.warning("Edge %d: PointCenter creation/update failed: %s",
                           i, exc)
            continue

        center_points.append({
            "point": pt,
            "name": pt.Name,
            "radius": radius,
        })

    selection.Clear()

    logger.info("Edge scan done: %d total, %d no-radius, %d matched, "
                "%d points created.",
                edge_count, no_radius_count, match_count, len(center_points))

    # Deduplicate coincident points using CATIA distance measurement.
    filtered = _dedup_points_by_distance(part, spa_workbench, center_points)

    logger.info("Generated %d center points", len(filtered))
    return filtered


def _dedup_points_by_distance(part, spa, center_points, tol=0.01):
    """Remove duplicate points that are closer than ``tol`` mm apart."""
    if not center_points:
        return center_points

    kept = [center_points[0]]
    for p in center_points[1:]:
        is_dup = False
        for k in kept:
            d = _spa_distance(part, spa, p["point"], k["point"])
            if d is not None and d < tol:
                is_dup = True
                break
        if not is_dup:
            kept.append(p)
    return kept


# =========================================================================
#  3. ORDERING
# =========================================================================

def select_starting_point(session, center_points):
    """Let the user pick which center point is the pipe's starting end."""
    selection = session.selection
    selection.Clear()
    for p in center_points:
        selection.Add(p["point"])

    logger.info("Waiting for starting point selection...")
    status = selection.SelectElement2(("Point",), "Select the starting point of the centerline", False)
    if status == "Cancel":
        raise RuntimeError("User cancelled starting point selection.")

    selected = selection.Item(1).Value
    name = selected.Name
    selection.Clear()

    # Try exact name match first
    for i, p in enumerate(center_points):
        if p["name"] == name:
            logger.info("✓ Starting point selected: %s", name)
            return i

    # User may have picked a duplicate that was deduped out —
    # find the closest kept point by SPA distance.
    part = session.part
    spa = session.spa_workbench
    best_i, best_d = 0, float("inf")
    for i, p in enumerate(center_points):
        d = _spa_distance(part, spa, selected, p["point"])
        if d is not None and d < best_d:
            best_d, best_i = d, i
    logger.info("✓ Starting point selected: %s (matched to %s, dist=%.3f mm)",
                name, center_points[best_i]["name"], best_d)
    return best_i


def nearest_neighbour_sort(part, spa, points, start_idx=0):
    """Order points into a continuous path via greedy nearest-neighbour
    using CATIA SPA distance measurement (no coordinates needed)."""
    logger.info("   Sorting points using nearest neighbor...")
    n = len(points)
    used = [False] * n
    ordered = []
    idx = start_idx

    for _ in range(n):
        used[idx] = True
        ordered.append(points[idx])

        best_dist, best_j = float("inf"), -1
        for j in range(n):
            if used[j]:
                continue
            d = _spa_distance(part, spa, points[idx]["point"], points[j]["point"])
            if d is not None and d < best_dist:
                best_dist, best_j = d, j
        if best_j < 0:
            break
        idx = best_j

    logger.info("   Total points sorted: %d", len(ordered))
    return ordered


# =========================================================================
#  4. LINES + BEND CONNECTS
# =========================================================================

def build_segments(part, hs_factory, spa, ordered_points, geo_set, arc_threshold):
    """
    For each consecutive pair of points: a straight Line if the distance
    is >= arc_threshold, otherwise leave the gap for ``bridge_bends`` to
    fill with a Connect.

    Before creating geometry, the user is asked whether they want to
    manually choose straight/curved for each segment.
    """
    import tkinter.messagebox as mbox

    logger.info("   Distance threshold: %.1f mm", arc_threshold)
    logger.info("   Logic: Distance < %.1f = ARC (connect), Distance >= %.1f = STRAIGHT (line)",
                arc_threshold, arc_threshold)

    # --- Measure all distances first ---
    seg_data = []
    for i in range(len(ordered_points) - 1):
        p1, p2 = ordered_points[i], ordered_points[i + 1]
        dist = _spa_distance(part, spa, p1["point"], p2["point"])
        if dist is None:
            dist = arc_threshold
        auto_is_arc = dist < arc_threshold
        seg_data.append({
            "p1": p1, "p2": p2, "dist": dist,
            "auto_is_arc": auto_is_arc,
        })

    # --- Show distances in the log ---
    logger.info("   Analyzing segments...")
    logger.info("   " + "=" * 50)
    logger.info("   SEGMENT DISTANCES:")
    for i, sd in enumerate(seg_data):
        label = "ARC (connect)" if sd["auto_is_arc"] else "STRAIGHT (line)"
        logger.info("   Segment %3d: %s → %s", i + 1,
                     sd["p1"]["name"], sd["p2"]["name"])
        logger.info("               Distance: %8.3f mm  [%s]", sd["dist"], label)
    logger.info("   " + "=" * 50)

    # --- Ask user if they want manual control ---
    manual_mode = mbox.askyesno(
        "Segment Type Selection",
        f"There are {len(seg_data)} segments detected.\n\n"
        f"Do you want to manually choose Straight or Curved\n"
        f"for each segment?\n\n"
        f"Yes  →  Choose per segment (recommended for complex pipes)\n"
        f"No   →  Use automatic threshold ({arc_threshold:.0f} mm)",
        icon="question"
    )

    # --- Determine segment types ---
    segments = []
    for i, sd in enumerate(seg_data):
        if manual_mode:
            auto_label = "CURVED" if sd["auto_is_arc"] else "STRAIGHT"
            use_straight = mbox.askyesno(
                f"Segment {i + 1} of {len(seg_data)}",
                f"{sd['p1']['name']}  →  {sd['p2']['name']}\n"
                f"Distance: {sd['dist']:.3f} mm\n"
                f"Auto suggestion: {auto_label}\n\n"
                f"Yes  →  Straight line\n"
                f"No   →  Curved bend (connect)",
                icon="question"
            )
            is_arc = not use_straight
        else:
            is_arc = sd["auto_is_arc"]

        seg = {"p1": sd["p1"]["point"], "p2": sd["p2"]["point"],
               "dist": sd["dist"], "is_arc": is_arc,
               "line": None, "connect": None}

        if not is_arc:
            line = hs_factory.AddNewLinePtPt(
                part.CreateReferenceFromObject(sd["p1"]["point"]),
                part.CreateReferenceFromObject(sd["p2"]["point"])
            )
            line.Name = f"Spine_Line_{i + 1}"
            geo_set.AppendHybridShape(line)
            seg["line"] = line

        segments.append(seg)

    line_count = sum(1 for s in segments if s["line"])
    arc_count = sum(1 for s in segments if s["is_arc"])

    part.Update()
    logger.info("   ✓ Created %d lines (straight segments)",  line_count)
    logger.info("   ℹ %d arc segments to process", arc_count)
    return segments


def bridge_bends(part, hs_factory, spa_workbench, selection, segments, geo_set,
                  continuity=1, tension=1.0):
    """
    Fill each arc segment with a Connect curve between the two neighbouring
    straight lines, trying all 4 orientation combinations and keeping the
    shortest result (avoids twisted connects).

    When a connect cannot be created, the user is prompted to choose between
    a straight-line fallback or skipping the segment.
    """
    import tkinter.messagebox as mbox

    logger.info("   Creating connects for arc segments...")
    connects_made = 0
    fallback_lines = 0

    for i, seg in enumerate(segments):
        if not seg["is_arc"]:
            continue

        prev_curve = next((segments[j]["line"] or segments[j]["connect"]
                           for j in range(i - 1, -1, -1)
                           if segments[j]["line"] or segments[j]["connect"]), None)
        next_curve = next((segments[j]["line"] or segments[j]["connect"]
                           for j in range(i + 1, len(segments))
                           if segments[j]["line"] or segments[j]["connect"]), None)

        best_connect = None

        if prev_curve is not None and next_curve is not None:
            best_length = float("inf")

            for o1, o2 in ((1, 1), (1, -1), (-1, 1), (-1, -1)):
                try:
                    candidate = hs_factory.AddNewConnect(
                        part.CreateReferenceFromObject(prev_curve),
                        part.CreateReferenceFromObject(seg["p1"]),
                        o1, continuity, tension,
                        part.CreateReferenceFromObject(next_curve),
                        part.CreateReferenceFromObject(seg["p2"]),
                        o2, continuity, tension,
                        False
                    )
                    part.UpdateObject(candidate)

                    length = spa_workbench.GetMeasurable(
                        part.CreateReferenceFromObject(candidate)
                    ).Length

                    if length < best_length:
                        if best_connect is not None:
                            selection.Clear()
                            selection.Add(best_connect)
                            try:
                                selection.Delete()
                            except Exception:
                                pass
                        best_connect, best_length = candidate, length
                    else:
                        selection.Clear()
                        selection.Add(candidate)
                        try:
                            selection.Delete()
                        except Exception:
                            pass
                except Exception:
                    pass

            selection.Clear()

        if best_connect is not None:
            best_connect.Name = f"Bend_Connect_{i + 1}"
            geo_set.AppendHybridShape(best_connect)
            part.UpdateObject(best_connect)
            seg["connect"] = best_connect
            connects_made += 1
        else:
            # Ask the user what to do with this failed bend
            p1_name = seg["p1"].Name if hasattr(seg["p1"], "Name") else "?"
            p2_name = seg["p2"].Name if hasattr(seg["p2"], "Name") else "?"
            dist_str = f"{seg['dist']:.1f}" if seg.get("dist") else "?"

            use_line = mbox.askyesno(
                "Bend Connect Failed",
                f"Segment {i + 1}: {p1_name} → {p2_name}\n"
                f"Distance: {dist_str} mm\n\n"
                f"Connect curve could not be created for this bend.\n\n"
                f"Yes  →  Use a straight line (recommended)\n"
                f"No   →  Skip this segment (leave a gap)",
                icon="question"
            )

            if use_line:
                try:
                    line = hs_factory.AddNewLinePtPt(
                        part.CreateReferenceFromObject(seg["p1"]),
                        part.CreateReferenceFromObject(seg["p2"])
                    )
                    line.Name = f"Bend_Line_{i + 1}"
                    geo_set.AppendHybridShape(line)
                    part.UpdateObject(line)
                    seg["line"] = line
                    fallback_lines += 1
                    logger.info("   Bend %d: straight line created (user choice).", i + 1)
                except Exception as exc:
                    logger.warning("   Bend %d: line creation failed: %s", i + 1, exc)
            else:
                logger.info("   Bend %d: skipped by user.", i + 1)

    logger.info("   ✓ Created %d connects for arcs, %d fallback lines", connects_made, fallback_lines)
    return segments


# =========================================================================
#  5. JOIN INTO SINGLE SPINE
# =========================================================================

def join_spine(part, hs_factory, segments, geo_set):
    """Join all Line/Connect segments into a single centerline curve."""
    curves = [s["line"] or s["connect"] for s in segments if (s["line"] or s["connect"])]
    if not curves:
        return None
    if len(curves) == 1:
        return curves[0]

    join = hs_factory.AddNewJoin(
        part.CreateReferenceFromObject(curves[0]),
        part.CreateReferenceFromObject(curves[1])
    )
    for c in curves[2:]:
        join.AddElement(part.CreateReferenceFromObject(c))

    join.SetConnex(True)
    join.SetManifold(True)
    # Increase merge tolerance to handle small gaps from fallback lines
    try:
        join.SetTolerance(0.1)
    except Exception:
        pass
    join.Name = "Centerline_Spine"
    geo_set.AppendHybridShape(join)

    try:
        part.UpdateObject(join)
    except Exception:
        try:
            part.Update()
        except Exception as e:
            logger.warning("Join update warning: %s (continuing anyway)", e)

    logger.info("✓ Joined %d curves into single path", len(curves))
    return join


# =========================================================================
#  6. FULL PIPELINE
# =========================================================================

def build_centerline(session, geo_set, arc_threshold):
    """
    Run the full interactive centerline workflow.

    Returns (spine, reference_radius). Raises RuntimeError on failure.
    """
    part = session.part
    hs_factory = session.hybrid_shape_factory
    spa = session.spa_workbench
    selection = session.selection

    face_ref = select_outer_surface(session)
    extract = create_multiextract(part, hs_factory, geo_set, face_ref)

    ref_radius = select_reference_radius(session, part, spa, hs_factory, geo_set)
    center_points = extract_centerpoints(part, hs_factory, spa, selection, extract, geo_set, ref_radius)

    # --- Ask for starting point FIRST ---
    logger.info("Select starting point for curve...")
    start_idx = select_starting_point(session, center_points)

    # --- Then sort, create lines, connects, join ---
    logger.info("Creating tube spine (distance-based line/connect selection)")
    ordered = nearest_neighbour_sort(part, spa, center_points, start_idx)

    segments = build_segments(part, hs_factory, spa, ordered, geo_set, arc_threshold)
    segments = bridge_bends(part, hs_factory, spa, selection, segments, geo_set)

    line_count = sum(1 for s in segments if s["line"])
    connect_count = sum(1 for s in segments if s["connect"])
    total = line_count + connect_count
    logger.info("✓ Tube spine complete: %d lines (straight), %d connects (arcs)", line_count, connect_count)
    logger.info("   Total curves for joining: %d", total)

    logger.info("Joining all curves...")
    spine = join_spine(part, hs_factory, segments, geo_set)
    if spine is None:
        raise RuntimeError("Failed to join centerline segments into a spine.")

    return spine, ref_radius