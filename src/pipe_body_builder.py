"""
Pipe Body Builder — turns a centerline spine into the final body.

Three output modes (selected from the GUI):

  "rib"     — solid Rib swept along the spine (hollow via a two-circle
              sketch profile if requested).
  "surface" — Sweep surface + optional end Fills + Join + CloseSurface
              (+ Shell if hollow).
  "sweep"   — Sweep surface only, no caps / solidification.
"""
from src.utils import setup_logger
from src.surface_builder import create_sweep, join_surfaces
from src.solid_converter import create_target_body, create_close_surface, create_shell

logger = setup_logger(__name__)


def _spine_start_geometry(part, hs_factory, geo_set, spine_ref):
    """Point + normal plane at the start (0%) of the spine."""
    pt_start = hs_factory.AddNewPointOnCurveFromPercent(spine_ref, 0.0, True)
    pt_start.Name = "Spine_Start_Point"
    geo_set.AppendHybridShape(pt_start)

    plane = hs_factory.AddNewPlaneNormal(spine_ref, part.CreateReferenceFromObject(pt_start))
    plane.Name = "Profile_Plane"
    geo_set.AppendHybridShape(plane)

    part.Update()
    return pt_start, plane


# =========================================================================
#  RIB (SOLID) MODE
# =========================================================================

def build_rib_body(part, hs_factory, shape_factory, geo_set, spine, radius, hollow, thickness):
    """Solid Rib — circular (or hollow two-circle) profile swept along the spine."""
    spine_ref = part.CreateReferenceFromObject(spine)
    pt_start, plane = _spine_start_geometry(part, hs_factory, geo_set, spine_ref)
    plane_ref = part.CreateReferenceFromObject(plane)

    body = part.Bodies.Add()
    body.Name = "Pipe_Solid_Body"
    part.InWorkObject = body
    part.Update()

    if hollow and thickness > 0:
        sketches = body.Sketches
        sketch = sketches.Add(plane_ref)
        factory_2d = sketch.OpenEdition()
        factory_2d.CreateClosedCircle(0, 0, radius)
        factory_2d.CreateClosedCircle(0, 0, max(radius - thickness, 0.01))
        sketch.CloseEdition()
        part.Update()
        profile_ref = part.CreateReferenceFromObject(sketch)
        logger.info("Hollow profile sketch created (outer=%.2f mm, inner=%.2f mm).",
                     radius, radius - thickness)
    else:
        circle = hs_factory.AddNewCircleCtrRad(
            part.CreateReferenceFromObject(pt_start), plane_ref, False, radius
        )
        circle.Name = "Pipe_Profile_Circle"
        geo_set.AppendHybridShape(circle)
        part.Update()
        profile_ref = part.CreateReferenceFromObject(circle)
        logger.info("Solid profile circle created (radius=%.2f mm).", radius)

    rib = shape_factory.AddNewRibFromRef(profile_ref, spine_ref)
    rib.Name = "Pipe_Rib"
    part.Update()
    logger.info("Rib body created in '%s'.", body.Name)
    return body, rib


# =========================================================================
#  SURFACE / SWEEP MODES
# =========================================================================

def _create_end_fills(part, hs_factory, selection, sweep, geo_set, cap_front, cap_back):
    """Create Fill caps on the requested ends of the sweep surface."""
    fills = []
    selection.Clear()
    selection.Add(sweep)
    try:
        selection.Search("Topology.CGMEdge,sel")
    except Exception:
        selection.Clear()
        return fills

    edge_count = selection.Count
    targets = []
    if cap_front and edge_count >= 1:
        targets.append((1, "Fill_Front"))
    if cap_back and edge_count >= 2:
        targets.append((edge_count, "Fill_Back"))
    elif cap_back and edge_count == 1 and not cap_front:
        targets.append((1, "Fill_Back"))

    for idx, label in targets:
        try:
            edge_ref = selection.Item(idx).Reference
            edge_extract = hs_factory.AddNewExtract(edge_ref)
            edge_extract.Name = f"{label}_Edge"
            geo_set.AppendHybridShape(edge_extract)
            part.UpdateObject(edge_extract)

            fill = hs_factory.AddNewFill()
            fill.AddBound(part.CreateReferenceFromObject(edge_extract))
            fill.Name = label
            geo_set.AppendHybridShape(fill)
            part.UpdateObject(fill)
            fills.append(fill)
            logger.info("%s created.", label)
        except Exception as e:
            logger.warning("%s creation failed: %s", label, e)

    selection.Clear()
    return fills


def build_surface_body(part, hs_factory, shape_factory, selection, geo_set,
                        spine, radius, hollow, thickness,
                        cap_front=True, cap_back=True, sweep_only=False):
    """
    Sweep surface, with optional end Fill caps, Join, CloseSurface,
    and Shell (if hollow). Returns (body, sweep, fills, join, (close_surface, shell)).

    ``body``/``join``/(close_surface, shell) are None where the
    corresponding step was skipped (e.g. sweep_only, or no caps requested).
    """
    spine_ref = part.CreateReferenceFromObject(spine)

    sweep = create_sweep(part, hs_factory, spine_ref, geo_set, radius=radius)
    if sweep is None:
        raise RuntimeError("Sweep surface creation failed.")

    if sweep_only:
        return None, sweep, [], None, None

    fills = _create_end_fills(part, hs_factory, selection, sweep, geo_set, cap_front, cap_back)

    if not fills:
        logger.info("No end caps requested/created — leaving as open sweep surface.")
        return None, sweep, fills, None, None

    join = join_surfaces(part, hs_factory, sweep, fills, geo_set)
    if join is None:
        return None, sweep, fills, None, None

    # Only fully solidify (CloseSurface/Shell) if BOTH ends are capped —
    # otherwise the joined surface isn't a closed volume.
    if not (cap_front and cap_back):
        logger.info("Only one end capped — surface left open (no CloseSurface/Shell).")
        return None, sweep, fills, join, None

    body = create_target_body(part)
    join_ref = part.CreateReferenceFromObject(join)
    close_surface = create_close_surface(part, shape_factory, join_ref)

    shell = None
    if close_surface and hollow and thickness > 0:
        shell = create_shell(part, shape_factory, thickness)

    part.Update()
    return body, sweep, fills, join, (close_surface, shell)
