"""
GUI Pipeline — orchestrates centerline extraction + body creation for the
Tkinter GUI (see src/gui_app.py).
"""
from src.catia_connection import CatiaSession
from src.centerline_builder_v2 import build_centerline
from src.pipe_body_builder import build_rib_body, build_surface_body
from src.utils import setup_logger

logger = setup_logger(__name__)


def run_pipe_builder(params, log):
    """
    Run the full centerline + body pipeline against the active CATIA document.

    params: dict with keys
        radius        (float)  — pipe profile radius in mm (0 = use detected radius)
        arc_threshold (float)  — bend vs. straight distance threshold in mm
        hollow        (bool)   — create a hollow profile / shell
        thickness     (float)  — wall thickness in mm (used if hollow)
        cap_front     (bool)   — add an end Fill at the front (surface/rib modes)
        cap_back      (bool)   — add an end Fill at the back
        mode          (str)    — "rib" | "surface" | "sweep"

    log: callable(str) — progress messages
    """
    with CatiaSession() as session:
        part = session.part
        hs_factory = session.hybrid_shape_factory
        shape_factory = session.shape_factory
        selection = session.selection

        log("Creating construction geometry set...")
        geo_set = session.create_geometrical_set("Construction_Geometry")

        # Route centerline builder logs to the GUI text area so the user
        # can see exactly what is happening during edge scanning.
        import logging
        _cl_logger = logging.getLogger("src.centerline_builder_v2")
        class _GuiLogHandler(logging.Handler):
            def emit(self, record):
                log(f"   [{record.levelname}] {record.getMessage()}")
        _gui_handler = _GuiLogHandler()
        _cl_logger.addHandler(_gui_handler)

        log("Step 1/3: Building centerline spine — follow the prompts in CATIA:")
        log("   1) Select the outer tube surface")
        log("   2) Select one circular edge (sets the reference radius)")
        log("   3) Select the starting point of the centerline")
        try:
            spine, ref_radius = build_centerline(session, geo_set, params["arc_threshold"])
        finally:
            _cl_logger.removeHandler(_gui_handler)
        log(f"  ✓ Centerline complete. Detected pipe radius ≈ {ref_radius:.3f} mm")

        radius = params["radius"] if params["radius"] > 0 else ref_radius

        mode = params["mode"]
        log(f"Step 2/3: Building final body (mode = {mode})...")

        if mode == "rib":
            body, rib = build_rib_body(
                part, hs_factory, shape_factory, geo_set, spine,
                radius, params["hollow"], params["thickness"]
            )
            log(f"  ✓ Solid rib created in body '{body.Name}'.")

        elif mode == "surface":
            body, sweep, fills, join, solidify = build_surface_body(
                part, hs_factory, shape_factory, selection, geo_set, spine,
                radius, params["hollow"], params["thickness"],
                cap_front=params["cap_front"], cap_back=params["cap_back"],
                sweep_only=False,
            )
            if body:
                log(f"  ✓ Surface solidified in body '{body.Name}'.")
            elif join:
                log("  ✓ Sweep + end fill(s) joined (open surface — both ends not capped).")
            else:
                log("  ✓ Sweep surface created (no end caps requested).")

        elif mode == "sweep":
            build_surface_body(
                part, hs_factory, shape_factory, selection, geo_set, spine,
                radius, params["hollow"], params["thickness"],
                cap_front=False, cap_back=False, sweep_only=True,
            )
            log("  ✓ Sweep surface created.")

        else:
            raise ValueError(f"Unknown mode: {mode!r}")

        log("Step 3/3: Final update...")
        try:
            part.Update()
        except Exception as e:
            log(f"  ⚠ Final update warning: {e}")

        log("Done.")
