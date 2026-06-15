"""
Tkinter GUI for the CATIA Parametric Pipe Builder.

Run with:  python main.py gui

Style: white background, a navy -> orange gradient header banner, and a
green "Run" button.
"""
import traceback
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from src.gui_pipeline import run_pipe_builder
from src.utils import setup_logger

logger = setup_logger(__name__)

# ---------------------------------------------------------------------------
#  Palette
# ---------------------------------------------------------------------------
NAVY = "#1B2C4F"
ORANGE = "#E8700A"
GREEN = "#43A047"
GREEN_HOVER = "#388E3C"
WHITE = "#FFFFFF"
LIGHT_GRAY = "#F4F4F4"
TEXT_GRAY = "#444444"
MUTED_GRAY = "#AAAAAA"


class GradientHeader(tk.Canvas):
    """A horizontal navy -> orange gradient banner with a title."""

    def __init__(self, parent, text, height=64, **kwargs):
        super().__init__(parent, height=height, highlightthickness=0,
                          bg=WHITE, bd=0, **kwargs)
        self._text = text
        self._height = height
        self.bind("<Configure>", self._redraw)

    def _redraw(self, event=None):
        self.delete("all")
        width = self.winfo_width() or 620
        height = self._height

        r1, g1, b1 = self.winfo_rgb(NAVY)
        r2, g2, b2 = self.winfo_rgb(ORANGE)
        steps = max(width, 1)

        for i in range(steps):
            t = i / steps
            r = int((r1 + (r2 - r1) * t)) >> 8
            g = int((g1 + (g2 - g1) * t)) >> 8
            b = int((b1 + (b2 - b1) * t)) >> 8
            self.create_line(i, 0, i, height, fill=f"#{r:02x}{g:02x}{b:02x}")

        self.create_text(20, height // 2, text=self._text, anchor="w",
                          fill="white", font=("Segoe UI", 16, "bold"))


class PipeBuilderGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CATIA Parametric Pipe Builder")
        self.geometry("640x660")
        self.configure(bg=WHITE)
        self.resizable(False, False)
        self._build_ui()

    # ------------------------------------------------------------------
    #  UI construction
    # ------------------------------------------------------------------
    def _build_ui(self):
        header = GradientHeader(self, "CATIA Parametric Pipe Builder")
        header.pack(fill="x")

        form = tk.Frame(self, bg=WHITE, padx=24, pady=16)
        form.pack(fill="x")

        # Pipe radius
        self.radius_var = tk.StringVar(value="2.0")
        self._labeled_entry(form, "Pipe Radius (mm):", self.radius_var, 0,
                             hint="0 = use radius detected from selected edge")

        # Arc / bend threshold
        self.threshold_var = tk.StringVar(value="20.0")
        self._labeled_entry(form, "Arc / Bend Threshold (mm):", self.threshold_var, 1,
                             hint="gaps below this become bends (Connects)")

        # Hollow toggle + thickness
        self.hollow_var = tk.BooleanVar(value=False)
        self.thickness_var = tk.StringVar(value="1.0")
        self._hollow_row(form, 2)

        # Front / Back end treatment
        self.front_var = tk.StringVar(value="Open")
        self.back_var = tk.StringVar(value="Open")
        self._combo_row(form, "Front End:", self.front_var, ["Open", "Capped"], 3)
        self._combo_row(form, "Back End:", self.back_var, ["Open", "Capped"], 4)

        # Final body mode
        self.mode_var = tk.StringVar(value="Rib (Solid)")
        self._combo_row(
            form, "Final Body Mode:", self.mode_var,
            ["Rib (Solid)", "Surface (Sweep + Fill + Join)", "Sweep Only"],
            5, width=26, span=True
        )

        # Hint
        hint = tk.Label(
            form,
            text=("After clicking Run, switch to CATIA and follow the prompts:\n"
                  "select the outer tube surface, a circular edge, then the\n"
                  "starting point of the centerline."),
            bg=WHITE, fg=TEXT_GRAY, justify="left", font=("Segoe UI", 9)
        )
        hint.grid(row=6, column=0, columnspan=3, sticky="w", pady=(12, 0))

        # Run button
        self.run_btn = tk.Button(
            self, text="▶  Run", command=self.on_run,
            bg=GREEN, fg="white", activebackground=GREEN_HOVER,
            activeforeground="white", disabledforeground="white",
            font=("Segoe UI", 12, "bold"), relief="flat",
            cursor="hand2", height=2, bd=0
        )
        self.run_btn.pack(fill="x", padx=24, pady=(10, 6))

        # Log area
        log_frame = tk.Frame(self, bg=WHITE, padx=24)
        log_frame.pack(fill="both", expand=True)
        tk.Label(log_frame, text="Log:", bg=WHITE, anchor="w",
                 font=("Segoe UI", 9, "bold")).pack(fill="x")
        self.log = ScrolledText(log_frame, wrap="word", height=14,
                                 state="disabled", bg=LIGHT_GRAY, relief="flat",
                                 font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, pady=(2, 8))

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status = tk.Label(self, textvariable=self.status_var, anchor="w",
                           bg=LIGHT_GRAY, fg=TEXT_GRAY, padx=10, pady=4)
        status.pack(fill="x")

    # ------------------------------------------------------------------
    #  Form helpers
    # ------------------------------------------------------------------
    def _labeled_entry(self, parent, label, var, row, hint=None):
        tk.Label(parent, text=label, bg=WHITE, anchor="w", width=22,
                 font=("Segoe UI", 10)).grid(row=row, column=0, sticky="w", pady=4)
        tk.Entry(parent, textvariable=var, width=12,
                 font=("Segoe UI", 10)).grid(row=row, column=1, sticky="w", pady=4)
        if hint:
            tk.Label(parent, text=hint, bg=WHITE, fg=MUTED_GRAY,
                     font=("Segoe UI", 8)).grid(row=row, column=2, sticky="w", padx=(8, 0))

    def _hollow_row(self, parent, row):
        chk = tk.Checkbutton(
            parent, text="Hollow pipe (Shell)", variable=self.hollow_var,
            bg=WHITE, font=("Segoe UI", 10), command=self._toggle_thickness,
            activebackground=WHITE, selectcolor=WHITE
        )
        chk.grid(row=row, column=0, sticky="w", pady=4)

        self.thickness_label = tk.Label(parent, text="Wall Thickness (mm):", bg=WHITE,
                                         anchor="w", font=("Segoe UI", 10))
        self.thickness_entry = tk.Entry(parent, textvariable=self.thickness_var,
                                         width=10, font=("Segoe UI", 10))
        self.thickness_label.grid(row=row, column=1, sticky="w", pady=4)
        self.thickness_entry.grid(row=row, column=2, sticky="w", pady=4, padx=(8, 0))
        self._toggle_thickness()

    def _toggle_thickness(self):
        if self.hollow_var.get():
            self.thickness_entry.configure(state="normal")
            self.thickness_label.configure(fg=TEXT_GRAY)
        else:
            self.thickness_entry.configure(state="disabled")
            self.thickness_label.configure(fg=MUTED_GRAY)

    def _combo_row(self, parent, label, var, values, row, width=12, span=False):
        tk.Label(parent, text=label, bg=WHITE, anchor="w", width=22,
                 font=("Segoe UI", 10)).grid(row=row, column=0, sticky="w", pady=4)
        combo = ttk.Combobox(parent, textvariable=var, width=width, state="readonly",
                              values=values, font=("Segoe UI", 10))
        if span:
            combo.grid(row=row, column=1, columnspan=2, sticky="w", pady=4)
        else:
            combo.grid(row=row, column=1, sticky="w", pady=4)

    # ------------------------------------------------------------------
    #  Logging
    # ------------------------------------------------------------------
    def log_message(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    #  Run
    # ------------------------------------------------------------------
    def _validate(self):
        radius = float(self.radius_var.get())
        threshold = float(self.threshold_var.get())
        hollow = self.hollow_var.get()
        thickness = float(self.thickness_var.get()) if hollow else 0.0

        if radius < 0:
            raise ValueError("Pipe radius cannot be negative.")
        if threshold <= 0:
            raise ValueError("Arc/bend threshold must be positive.")
        if hollow:
            if thickness <= 0:
                raise ValueError("Wall thickness must be positive for a hollow pipe.")
            if radius > 0 and thickness >= radius:
                raise ValueError("Wall thickness must be less than the pipe radius.")

        return radius, threshold, hollow, thickness

    def on_run(self):
        try:
            radius, threshold, hollow, thickness = self._validate()
        except ValueError as e:
            self.status_var.set(f"Input error: {e}")
            self.log_message(f"ERROR: {e}")
            return

        mode_map = {
            "Rib (Solid)": "rib",
            "Surface (Sweep + Fill + Join)": "surface",
            "Sweep Only": "sweep",
        }

        params = {
            "radius": radius,
            "arc_threshold": threshold,
            "hollow": hollow,
            "thickness": thickness,
            "cap_front": self.front_var.get() == "Capped",
            "cap_back": self.back_var.get() == "Capped",
            "mode": mode_map[self.mode_var.get()],
        }

        self.clear_log()
        self.run_btn.configure(state="disabled", bg=GREEN_HOVER)
        self.status_var.set("Running — switch to CATIA and follow the selection prompts...")

        try:
            run_pipe_builder(params, self.log_message)
            self.status_var.set("✓ Done")
            self.log_message("✓ Completed successfully.")
        except Exception as e:
            self.status_var.set("Error — see log")
            self.log_message(f"ERROR: {e}")
            self.log_message(traceback.format_exc())
        finally:
            self.run_btn.configure(state="normal", bg=GREEN)


def main():
    app = PipeBuilderGUI()
    app.log_message("Ready. Configure parameters and click Run.")
    app.mainloop()


if __name__ == "__main__":
    main()
