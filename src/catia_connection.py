"""
CATIA V5 Connection Manager.

Provides a ``CatiaSession`` context manager that connects to a running
CATIA V5 instance and exposes the most commonly used COM objects.

Usage::

    with CatiaSession() as session:
        part = session.part
        hs   = session.hybrid_shape_factory
        ...
"""
import os
import win32com.client

from src.utils import setup_logger

logger = setup_logger(__name__)


class CatiaSession:
    """Context-managed connection to a running CATIA V5 instance."""

    def __init__(self):
        self._catia = None
        self._document = None
        self._part = None

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------
    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # We intentionally do NOT close CATIA — the user keeps working.
        return False

    # ------------------------------------------------------------------
    # Connection
    # ------------------------------------------------------------------
    def connect(self):
        """Connect to a running CATIA V5 instance."""
        try:
            self._catia = win32com.client.Dispatch("CATIA.Application")
            self._document = self._catia.ActiveDocument
            self._part = self._document.Part
            logger.info("Connected to CATIA V5 — %s", self._document.Name)
        except Exception as e:
            raise RuntimeError(
                f"Could not connect to CATIA V5. Is it running with a Part open?\n  → {e}"
            ) from e

    # ------------------------------------------------------------------
    # Properties (lazy, cached from connection)
    # ------------------------------------------------------------------
    @property
    def catia(self):
        return self._catia

    @property
    def document(self):
        return self._document

    @property
    def part(self):
        return self._part

    @property
    def hybrid_shape_factory(self):
        return self._part.HybridShapeFactory

    @property
    def shape_factory(self):
        return self._part.ShapeFactory

    @property
    def spa_workbench(self):
        return self._document.GetWorkbench("SPAWorkbench")

    @property
    def selection(self):
        return self._document.Selection

    @property
    def hybrid_bodies(self):
        return self._part.HybridBodies

    # ------------------------------------------------------------------
    # Document operations
    # ------------------------------------------------------------------
    def open_part(self, filepath):
        """Open a .CATPart file in CATIA."""
        filepath = os.path.abspath(filepath)
        if not os.path.isfile(filepath):
            raise FileNotFoundError(f"File not found: {filepath}")
        try:
            self._catia.Documents.Open(filepath)
            self._document = self._catia.ActiveDocument
            self._part = self._document.Part
            logger.info("Opened: %s", filepath)
        except Exception as e:
            raise RuntimeError(f"Failed to open {filepath}: {e}") from e

    def save_part(self):
        """Save the active document."""
        try:
            self._document.Save()
            logger.info("Document saved.")
        except Exception as e:
            logger.warning("Save warning: %s", e)

    def save_part_as(self, filepath):
        """SaveAs the active document to a new path."""
        filepath = os.path.abspath(filepath)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        try:
            self._document.SaveAs(filepath)
            logger.info("Saved as: %s", filepath)
        except Exception as e:
            logger.warning("SaveAs warning: %s", e)

    def create_new_part(self, name="New_Part"):
        """Create a new empty Part document in CATIA."""
        try:
            self._catia.Documents.Add("Part")
            self._document = self._catia.ActiveDocument
            self._part = self._document.Part
            try:
                self._document.Product.PartNumber = name
            except Exception:
                pass
            logger.info("Created new part: %s", name)
        except Exception as e:
            raise RuntimeError(f"Failed to create new part: {e}") from e

    def close_document(self):
        """Close the active document (without saving)."""
        try:
            self._document.Close()
            logger.info("Document closed.")
        except Exception as e:
            logger.warning("Close warning: %s", e)

    # ------------------------------------------------------------------
    # Geometrical set helpers
    # ------------------------------------------------------------------
    def create_geometrical_set(self, name):
        """
        Create (or recreate) a Geometrical Set with the given name.

        If a set with that name already exists, it is deleted first
        so re-runs start clean.
        """
        sel = self.selection
        hbs = self.hybrid_bodies

        # Delete existing set if present
        try:
            old_set = hbs.Item(name)
            sel.Clear()
            sel.Add(old_set)
            sel.Delete()
            logger.info("Deleted existing '%s' set (clean re-run).", name)
        except Exception:
            pass  # set didn't exist — fine

        new_set = hbs.Add()
        new_set.Name = name
        logger.info("Created geometrical set: %s", name)
        return new_set
