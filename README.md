# CATIA Parametric Pipe Builder

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![CATIA V5](https://img.shields.io/badge/CATIA-V5-005386)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![COM](https://img.shields.io/badge/Automation-win32com%20COM-orange)

A Python desktop application that automates **parametric pipe and tube creation** in CATIA V5. Point the tool at an existing pipe geometry, and it will extract the centerline, reconstruct the spine with straight segments and smooth bends, and build a fully parametric pipe body — all driven from an intuitive Tkinter GUI.

> **Fully parametric output** — changing the reference circle radius in CATIA automatically updates the entire pipe geometry.

---

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Quick Start](#quick-start)
  - [Step-by-Step Workflow](#step-by-step-workflow)
  - [GUI Parameters](#gui-parameters)
- [Project Structure](#project-structure)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [License](#license)

---

## Features

| Category | Details |
|---|---|
| **Centerline Extraction** | Automatic edge scanning with tangent-continuity surface extraction; geometry-native center points via `AddNewPointCenter` |
| **Point Ordering** | Nearest-neighbor sorting algorithm ensures correct sequential point ordering along the pipe path |
| **Bend Detection** | Configurable distance threshold distinguishes straight segments from bends — all four orientation combinations are tested to find the shortest, twist-free Connect curve |
| **Interactive Fallback** | When a Connect curve fails, a dialog lets the user choose a straight-line fallback or skip the segment |
| **Body Modes** | **Rib** (solid swept profile), **Surface** (Sweep + Fill + Join + CloseSurface), or **Sweep Only** (open surface) |
| **Hollow Pipes** | Shell-based hollowing with configurable wall thickness; Rib mode uses a dual-circle sketch profile |
| **End Caps** | Independent front/back end capping (Fill surfaces) — fully closed surfaces are automatically solidified |
| **GUI Console** | Real-time, detailed logging of every pipeline step directly in the application window |
| **Robust COM Layer** | `CatiaSession` context manager with automatic connection, multiple SPA measurement fallbacks, and clean error handling |

---

## Prerequisites

| Requirement | Details |
|---|---|
| **CATIA V5** | Installed and running with a `.CATPart` document open (tested on V5 R19–R32) |
| **Python** | 3.10 or later, **64-bit** (must match CATIA's architecture) |
| **OS** | Windows 10 / 11 |

### Python Packages

| Package | Purpose |
|---|---|
| `pywin32` ≥ 306 | COM automation bridge to CATIA V5 (`win32com.client`) |
| `pyyaml` ≥ 6.0 | Configuration file parsing (`config/defaults.yaml`) |

> [!IMPORTANT]
> Your Python installation's bitness (32-bit vs 64-bit) **must match** your CATIA V5 installation. A mismatch will cause `win32com.client.Dispatch("CATIA.Application")` to fail silently or raise a COM error.

---

## Installation

1. **Clone or download** this repository:

   ```bash
   git clone <repository-url>
   cd forfix
   ```

2. **Create a virtual environment** (recommended):

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

   This installs `pywin32` and `pyyaml`.

---

## Usage

### Quick Start

1. Open **CATIA V5** and load the `.CATPart` containing the pipe/tube you want to recreate parametrically.
2. Launch the tool:

   ```bash
   python main.py
   ```

3. Configure parameters in the GUI, click **▶ Run**, then switch to CATIA and follow the selection prompts.

### Step-by-Step Workflow

```
┌─────────────────────────────────────────────────────────┐
│                    Tkinter GUI                          │
│  1. Set pipe radius, threshold, body mode, caps, etc.   │
│  2. Click ▶ Run                                         │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│              CATIA V5 — User Selections                 │
│  3. Select the outer cylindrical surface of the pipe    │
│  4. Select one circular edge (sets reference radius)    │
│  5. Select the starting point of the centerline         │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────┐
│              Automatic Pipeline                         │
│  6. Extract all edges from the pipe surface             │
│  7. Find circular edges matching the reference radius   │
│  8. Create center points for each matching edge         │
│  9. Sort points by nearest-neighbor traversal           │
│ 10. Create straight lines (far points) and Connect      │
│     curves (close points = bends)                       │
│ 11. Join segments into a single centerline spine        │
│ 12. Build the parametric pipe body (Rib/Surface/Sweep)  │
└─────────────────────────────────────────────────────────┘
```

### GUI Parameters

| Parameter | Default | Description |
|---|---|---|
| **Pipe Radius** | `2.0 mm` | Cross-section radius for the new pipe. Set to `0` to auto-detect from the selected reference edge. |
| **Arc / Bend Threshold** | `20.0 mm` | Center-point gaps **below** this distance are treated as bends (Connect curves); gaps **at or above** this become straight line segments. |
| **Hollow Pipe** | Off | Enable to create a hollow pipe with a shell offset. |
| **Wall Thickness** | `1.0 mm` | Shell wall thickness (only active when Hollow is enabled). Must be less than the pipe radius. |
| **Front End** | Open | `Capped` adds a Fill surface to seal the front end of the pipe. |
| **Back End** | Open | `Capped` adds a Fill surface to seal the back end of the pipe. |
| **Final Body Mode** | Rib (Solid) | See below. |

#### Body Modes

| Mode | Output | Best For |
|---|---|---|
| **Rib (Solid)** | Solid body created via `AddNewRibFromRef` with a circular (or dual-circle hollow) sketch profile | Final manufacturing-ready solid geometry |
| **Surface (Sweep + Fill + Join)** | Sweep surface + optional end Fill caps + Join; solidified via `CloseSurface` if both ends are capped | Surfacing workflows, QA checks on surface quality |
| **Sweep Only** | Single circular Sweep surface, no caps or solidification | Quick preview or downstream surface operations |

---

## Project Structure

```
forfix/
├── main.py                          # Entry point — launches the GUI
├── requirements.txt                 # Python dependencies (pywin32, pyyaml)
├── config/
│   └── defaults.yaml                # Default configuration values
├── src/
│   ├── __init__.py
│   ├── gui_app.py                   # Tkinter GUI (PipeBuilderGUI class)
│   ├── gui_pipeline.py              # Orchestrates GUI → centerline → body pipeline
│   ├── catia_connection.py          # CatiaSession context manager for CATIA V5 COM
│   ├── centerline_builder_v2.py     # Core centerline extraction and reconstruction
│   ├── pipe_body_builder.py         # Creates solid/surface pipe body (Rib, Surface, Sweep)
│   ├── surface_builder.py           # Sweep, Fill, and Join surface creation
│   ├── solid_converter.py           # CloseSurface, ThickSurface fallback, Shell
│   └── utils.py                     # Logging, configuration, math helpers
└── output/                          # Output directory
```

### Module Responsibilities

| Module | Responsibility |
|---|---|
| `gui_app.py` | Tkinter window with gradient header, parameter form, Run button, scrollable log console, and status bar. Validates user input and invokes the pipeline. |
| `gui_pipeline.py` | Connects the GUI to the backend — passes parameters into `build_centerline()` and then into the appropriate body builder, routing log messages back to the GUI console. |
| `catia_connection.py` | `CatiaSession` context manager that connects to a running CATIA V5 instance via `win32com.client.Dispatch("CATIA.Application")`. Exposes `part`, `hybrid_shape_factory`, `shape_factory`, `selection`, and `spa_workbench` properties. Provides document management helpers (open, save, close, create). |
| `centerline_builder_v2.py` | Full interactive centerline workflow: surface extract → edge scanning → radius matching → center point creation → nearest-neighbor sorting → straight lines + Connect curves for bends → Join into a single spine. |
| `pipe_body_builder.py` | Builds the final parametric pipe body in one of three modes (Rib, Surface, Sweep). Handles hollow profiles, end caps, and solidification. |
| `surface_builder.py` | Creates `SweepCircle` (with explicit sweep fallback), profile circles, end Fill caps, and surface Joins. |
| `solid_converter.py` | Converts a closed joined surface into a solid body via `CloseSurface` (with `ThickSurface` fallback) and applies `Shell` for hollow pipes. |
| `utils.py` | Shared utilities: `setup_logger()` for timestamped console logging, `load_config()` for YAML configuration with hardcoded fallback defaults, `com_safe` decorator for COM error handling, `distance_3d()` math helper. |

---

## Configuration

Default parameters are stored in [`config/defaults.yaml`](config/defaults.yaml):

```yaml
# Centerline detection
centerline:
  straightness_tolerance: 0.5      # mm — midpoint/endpoint face-distance match threshold
  dedup_tolerance: 5.0             # mm — midpoints closer than this are duplicates
  min_segment_length: 1.0          # mm — segments shorter than this are discarded

# Pipe profile
pipe:
  default_radius: 2.0              # mm — default cross-section radius for Rib
  default_wall_thickness: 1.0      # mm — default Shell offset

# Connect curves (bend bridging)
connect:
  continuity: 1                    # 0 = point, 1 = tangent, 2 = curvature
  tension: 1.0                     # tension factor for Connect curves
```

If the YAML file is missing or `pyyaml` is not installed, the application falls back to equivalent hardcoded defaults.

---

## Troubleshooting

### Connection Issues

| Problem | Solution |
|---|---|
| `Could not connect to CATIA V5` | Ensure CATIA V5 is **running** with a `.CATPart` document **open and active** before clicking Run. |
| `win32com` import error | Run `pip install pywin32` and then `python Scripts/pywin32_postinstall.py -install` from your Python directory. |
| COM error / `Dispatch` fails | Verify your Python bitness matches CATIA (both must be 64-bit or both 32-bit). Run `python -c "import struct; print(struct.calcsize('P')*8)"` to check. |

### Runtime Errors

| Problem | Solution |
|---|---|
| `Edge search on extract failed` | The selected surface may not be a valid cylindrical face. Select the **outer cylindrical surface** of the pipe, not a planar face or fillet. |
| No center points found | The reference edge radius may not match any edges on the extract. Try selecting a different circular edge. |
| Connect curve fails for all orientations | This can happen at tight bends or branching geometry. Use the dialog to fall back to a straight line or skip the segment. |
| `CloseSurface failed` | The joined surface may have small gaps. The tool automatically falls back to `ThickSurface`. Ensure both ends are capped for a valid closed volume. |
| `Join update warning` | Minor tolerance issues during the final join. The tool continues despite the warning — check the CATIA tree to verify the result. |

### General Tips

- **Set Pipe Radius to `0`** to let the tool auto-detect the radius from your selected reference edge — useful when you don't know the exact dimension.
- **Increase the Arc/Bend Threshold** if straight segments are being incorrectly classified as bends.
- **Decrease the Arc/Bend Threshold** if bends are being treated as straight lines, producing sharp corners.
- Check the **GUI log console** for detailed, timestamped output of every pipeline step — this is the fastest way to diagnose issues.

---

## License

This project is licensed under the [MIT License](LICENSE).
