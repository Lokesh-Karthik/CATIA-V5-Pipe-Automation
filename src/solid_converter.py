"""
Solid Converter — transform a joined closed surface into a parametric solid.

Creates CloseSurface + Shell features in a dedicated Body, matching the
output structure shown in the user's "after" screenshot:

    Body.4
    ├── CloseSurface.1
    └── Shell.1
"""
from src.utils import setup_logger

logger = setup_logger(__name__)


def create_target_body(part):
    """
    Create a new Body for the solid features and set it as InWorkObject.

    Returns the new body.
    """
    try:
        bodies = part.Bodies
        new_body = bodies.Add()
        new_body.Name = "Pipe_Solid_Body"
        part.InWorkObject = new_body
        logger.info("Created target body: %s", new_body.Name)
        return new_body
    except Exception as e:
        # Fallback: use MainBody
        logger.warning("Could not create new body (%s), using MainBody.", e)
        part.InWorkObject = part.MainBody
        return part.MainBody


def create_close_surface(part, shape_factory, join_ref):
    """
    Convert a closed joined surface into a solid via CloseSurface.

    Falls back to ThickSurface if CloseSurface fails (happens when
    the surface has gaps).

    Returns the solid feature object.
    """
    # Try CloseSurface first
    try:
        close_surface = shape_factory.AddNewCloseSurface(join_ref)
        close_surface.Name = "CloseSurface.1"
        part.Update()
        logger.info("CloseSurface created successfully.")
        return close_surface
    except Exception as e:
        logger.warning("CloseSurface failed: %s — trying ThickSurface fallback.", e)

    # Fallback: ThickSurface (works even if surface isn't perfectly closed)
    try:
        thick = shape_factory.AddNewThickSurface(join_ref, 1, 1.0, 1.0)
        thick.Name = "ThickSurface_Fallback"
        part.Update()
        logger.info("ThickSurface fallback created (1mm each side).")
        return thick
    except Exception as e2:
        logger.error("ThickSurface fallback also failed: %s", e2)
        return None


def create_shell(part, shape_factory, thickness=1.0):
    """
    Create a Shell feature to hollow out the solid pipe.

    Parameters
    ----------
    thickness : float
        Wall thickness in mm.

    Returns the Shell feature object.
    """
    try:
        # Shell with internal offset (hollows the solid)
        shell = shape_factory.AddNewShell(thickness)
        shell.Name = "Shell.1"
        part.Update()
        logger.info("Shell created (thickness = %.2f mm).", thickness)
        return shell
    except Exception as e:
        logger.error("Shell creation failed: %s", e)
        return None


def convert_to_solid(part, shape_factory, hs_factory, join, config):
    """
    Full surface-to-solid pipeline:
      1. Create target body
      2. CloseSurface (with ThickSurface fallback)
      3. Shell
      4. part.Update()

    Returns (body, close_surface, shell) — any may be None on failure.
    """
    thickness = config.get("pipe", {}).get("default_wall_thickness", 1.0)

    body = create_target_body(part)

    join_ref = part.CreateReferenceFromObject(join)
    close_surface = create_close_surface(part, shape_factory, join_ref)

    shell = None
    if close_surface:
        shell = create_shell(part, shape_factory, thickness)

    try:
        part.Update()
    except Exception as e:
        logger.warning("Final update warning: %s", e)

    return body, close_surface, shell
