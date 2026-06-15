"""
Surface Builder — create Sweep, Fill, and Join surfaces from the centerline.

These are used when the user wants a surface-based parametric result
(Geometrical Set → Sweep.1 / Fill.1 / Fill.2 / Join.1) instead of
or in addition to a solid Rib.
"""
from src.utils import setup_logger

logger = setup_logger(__name__)


def create_profile_circle(part, hs_factory, master_join, geo_set, radius=2.0):
    """
    Create a circle profile at the start of the master centerline spine.

    Returns (circle, start_point, plane).
    """
    join_ref = part.CreateReferenceFromObject(master_join)

    # Point at 0 % along the spine = pipe start
    pt_start = hs_factory.AddNewPointOnCurveFromPercent(join_ref, 0.0, True)
    pt_start.Name = "Spine_Start_Point"
    geo_set.AppendHybridShape(pt_start)

    # Plane normal to spine at start
    plane = hs_factory.AddNewPlaneNormal(
        join_ref, part.CreateReferenceFromObject(pt_start)
    )
    plane.Name = "Profile_Plane"
    geo_set.AppendHybridShape(plane)

    # Circle on the normal plane
    circle = hs_factory.AddNewCircleCtrRad(
        part.CreateReferenceFromObject(pt_start),
        part.CreateReferenceFromObject(plane),
        False,   # not geodesic
        radius,
    )
    circle.Name = "Pipe_Profile_Circle"
    geo_set.AppendHybridShape(circle)
    part.Update()

    logger.info("Profile circle created (radius = %.2f mm).", radius)
    return circle, pt_start, plane


def create_sweep(part, hs_factory, spine_ref, geo_set, radius=2.0):
    """
    Create a surface Sweep along the spine with a circular cross-section.

    Uses AddNewSweepCircle which is reliable across V5 R19-R21.

    Returns the sweep HybridShape object, or None on failure.
    """
    try:
        sweep = hs_factory.AddNewSweepCircle(spine_ref)
        sweep.Name = "Sweep.1"
        sweep.SetRadius(radius)
        geo_set.AppendHybridShape(sweep)
        part.Update()
        logger.info("Sweep surface created.")
        return sweep
    except Exception as e:
        logger.warning("SweepCircle failed: %s — trying explicit sweep.", e)

    # Fallback: explicit profile sweep
    try:
        # Create the profile circle on the spine start
        pt_start = hs_factory.AddNewPointOnCurveFromPercent(spine_ref, 0.0, True)
        geo_set.AppendHybridShape(pt_start)
        plane = hs_factory.AddNewPlaneNormal(
            spine_ref, part.CreateReferenceFromObject(pt_start)
        )
        geo_set.AppendHybridShape(plane)
        circle = hs_factory.AddNewCircleCtrRad(
            part.CreateReferenceFromObject(pt_start),
            part.CreateReferenceFromObject(plane),
            False, radius,
        )
        geo_set.AppendHybridShape(circle)
        part.Update()

        circle_ref = part.CreateReferenceFromObject(circle)
        sweep = hs_factory.AddNewSweepExplicit(circle_ref, spine_ref)
        sweep.Name = "Sweep.1"
        geo_set.AppendHybridShape(sweep)
        part.Update()
        logger.info("Sweep surface created (explicit profile).")
        return sweep
    except Exception as e2:
        logger.error("Explicit sweep also failed: %s", e2)
        return None


def create_end_fills(part, hs_factory, selection, sweep, geo_set):
    """
    Create Fill surfaces to cap the open ends of the swept pipe surface.

    Strategy:
      1. Select the sweep, search for boundary edges
      2. Extract each boundary edge
      3. Create a Fill from the boundary

    Returns list of fill objects (may be empty if fill creation fails).
    """
    fills = []
    try:
        selection.Clear()
        selection.Add(sweep)
        selection.Search("Topology.CGMEdge,sel")

        if selection.Count >= 2:
            # Extract the first and last boundary edges
            for idx, label in [(1, "Fill.1"), (selection.Count, "Fill.2")]:
                try:
                    edge_ref = selection.Item(idx).Reference
                    edge_extract = hs_factory.AddNewExtract(edge_ref)
                    edge_extract.Name = f"End_Edge_{idx}"
                    geo_set.AppendHybridShape(edge_extract)
                    part.UpdateObject(edge_extract)

                    extract_ref = part.CreateReferenceFromObject(edge_extract)
                    fill = hs_factory.AddNewFill()
                    fill.AddBound(extract_ref)
                    fill.Name = label
                    geo_set.AppendHybridShape(fill)
                    part.UpdateObject(fill)
                    fills.append(fill)
                    logger.info("%s created.", label)
                except Exception as e:
                    logger.warning("Fill at edge %d failed: %s", idx, e)
        else:
            logger.warning("Could not find boundary edges for fills (%d edges).", selection.Count)
    except Exception as e:
        logger.warning("End fill search failed: %s", e)

    selection.Clear()
    return fills


def join_surfaces(part, hs_factory, sweep, fills, geo_set):
    """
    Join the sweep surface and fill caps into one closed surface (Join.1).

    Returns the join HybridShape object, or None on failure.
    """
    surfaces = [sweep] + fills
    if len(surfaces) < 2:
        logger.warning("Not enough surfaces to join (%d). Skipping.", len(surfaces))
        return None

    try:
        join = hs_factory.AddNewJoin(
            part.CreateReferenceFromObject(surfaces[0]),
            part.CreateReferenceFromObject(surfaces[1]),
        )
        for s in surfaces[2:]:
            join.AddElement(part.CreateReferenceFromObject(s))

        join.Name = "Join.1"
        geo_set.AppendHybridShape(join)
        part.Update()
        logger.info("Surface join created (%d surfaces).", len(surfaces))
        return join
    except Exception as e:
        logger.error("Surface join failed: %s", e)
        return None
