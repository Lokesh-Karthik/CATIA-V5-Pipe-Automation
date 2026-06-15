# CATIA Parametric Pipe Builder — How-To Guide

> **Version:** 1.0  
> **Last Updated:** June 2026  
> **Tool:** Python Tkinter GUI for automated parametric pipe/tube creation in CATIA V5

---

## Table of Contents

1. [Setup & Installation](#1-setup--installation)  
2. [GUI Overview](#2-gui-overview)  
3. [Step-by-Step Usage](#3-step-by-step-usage)  
4. [Understanding the Log Output](#4-understanding-the-log-output)  
5. [Troubleshooting](#5-troubleshooting)  
6. [Tips & Best Practices](#6-tips--best-practices)  

---

## 1. Setup & Installation

### 1.1 Prerequisites

Before running the tool, ensure the following software is installed and available on your system:

| Requirement | Details |
|---|---|
| **CATIA V5** | R19 – R21 (or later), installed and **running** on Windows |
| **Python** | 3.10 or later, **64-bit** — must match CATIA's architecture (both 64-bit) |
| **pywin32** | ≥ 306 — provides the `win32com` bridge used to communicate with CATIA via COM |
| **PyYAML** | ≥ 6.0 — reads the default configuration file (`config/defaults.yaml`) |

> **Important:** Python's bit-width (32-bit vs. 64-bit) **must** match CATIA's. If CATIA V5 is installed as a 64-bit application, you must use a 64-bit Python interpreter. A mismatch will cause the COM connection to fail silently or with cryptic errors.

### 1.2 Installing Dependencies

Open a terminal (PowerShell or Command Prompt) and navigate to the project directory:

```bash
cd "d:\Loki\affluent internship\final_working\forfix"
pip install -r requirements.txt
```

The `requirements.txt` file installs:
- `pywin32>=306`
- `pyyaml>=6.0`

### 1.3 Launching the Tool

```bash
python main.py
```

This opens the **CATIA Parametric Pipe Builder** GUI window. CATIA V5 must already be running with a `.CATPart` file open before you click **Run**.

---

## 2. GUI Overview

When the application launches, you will see a window with a navy-to-orange gradient header, a parameter form, a green **Run** button, and a log panel. Below is a description of every configurable field.

### 2.1 Pipe Radius (mm)

| | |
|---|---|
| **Default** | `2.0` |
| **Purpose** | Sets the radius of the pipe's circular cross-section profile. |
| **Auto-detect** | Enter `0` to let the tool automatically use the radius detected from the circular edge you select in CATIA during the interactive prompts. |
| **Override** | Any positive value overrides the auto-detected radius with a fixed value. |

### 2.2 Arc / Bend Threshold (mm)

| | |
|---|---|
| **Default** | `20.0` |
| **Purpose** | Controls how the tool distinguishes between **straight segments** and **bends** when building the centerline. |

The tool measures the distance between each pair of consecutive center points along the pipe:

- **Distance ≥ threshold** → the segment is treated as a **straight line** (a `LinePtPt` is created).
- **Distance < threshold** → the segment is treated as a **bend / arc** (a `Connect` curve is created to smoothly bridge the gap with tangent continuity).

**Typical values:**
- `15–40 mm` depending on pipe complexity.
- Lower values (10–15 mm) → fewer arcs, more straight lines — better for complex pipes with tight bends.
- Higher values (25–40 mm) → more arcs — better for simple pipes with gentle, sweeping bends.

### 2.3 Hollow Pipe (Shell)

| | |
|---|---|
| **Default** | Unchecked (solid pipe) |
| **Purpose** | When enabled, the pipe is created as a hollow tube instead of a solid cylinder. |

- **Unchecked** — produces a solid pipe body.
- **Checked** — enables the **Wall Thickness (mm)** field. The tool creates either:
  - A two-circle sketch profile (inner + outer circles) for **Rib** mode, or
  - A `Shell` feature applied to the closed surface body for **Surface** mode.

**Wall Thickness (mm):** The thickness of the pipe wall. Must be a positive value less than the pipe radius. Default: `1.0 mm`.

### 2.4 Front End / Back End

| | |
|---|---|
| **Default** | Both set to `Open` |
| **Options** | `Open` or `Capped` |

These dropdowns control what happens at the two ends of the pipe:

- **Open** — the pipe end is left open (no fill surface).
- **Capped** — a `Fill` surface is created to close the pipe end.

> **Note:** In **Surface** mode, if both ends are set to **Capped**, the tool will create a fully closed volume and solidify it with a `CloseSurface`. If only one end is capped, the result remains an open surface body.

### 2.5 Final Body Mode

| Mode | Internal Key | Description |
|---|---|---|
| **Rib (Solid)** | `rib` | Creates a solid body by sweeping a circular profile along the centerline spine using the CATIA **Rib** feature. This is the simplest, fastest option and produces an editable solid body. |
| **Surface (Sweep + Fill + Join)** | `surface` | Creates the pipe as a set of surfaces: a **Sweep** for the tube wall, optional **Fill** caps at each end, and a **Join** to merge everything. If both ends are capped and the resulting surface is a closed volume, it is solidified via `CloseSurface` (and optionally shelled). |
| **Sweep Only** | `sweep` | Creates only the **Sweep** surface along the centerline — no end caps, no joining, no solidification. Useful for inspection or when you want to manually finish the part. |

---

## 3. Step-by-Step Usage

Follow this walkthrough to convert a "dumb" solid pipe part into a fully parametric, editable model.

### Step 1 — Open the Pipe Part in CATIA V5

1. Launch **CATIA V5**.
2. Open the `.CATPart` file that contains the existing pipe/tube geometry you want to rebuild parametrically.
3. Make sure the part is the **active document** (its window should be in the foreground).

### Step 2 — Launch the Tool

Open a terminal and run:

```bash
cd "d:\Loki\affluent internship\final_working\forfix"
python main.py
```

The GUI window will appear. You should see the message **"Ready. Configure parameters and click Run."** in the log panel.

### Step 3 — Configure Parameters in the GUI

Set the following parameters based on your pipe geometry:

1. **Pipe Radius** — enter the desired pipe radius, or leave at `0` to auto-detect it from the selected edge.
2. **Arc / Bend Threshold** — start with the default (`20.0 mm`). Adjust later if needed (see [Tips & Best Practices](#6-tips--best-practices)).
3. **Hollow Pipe** — check the box if you need a hollow pipe; set the wall thickness accordingly.
4. **Front End / Back End** — choose `Open` or `Capped` for each end.
5. **Final Body Mode** — select `Rib (Solid)`, `Surface (Sweep + Fill + Join)`, or `Sweep Only`.

### Step 4 — Click the Green ▶ Run Button

- The button will disable (turn dark green) and the status bar will display:  
  *"Running — switch to CATIA and follow the selection prompts..."*
- The log panel will show the pipeline steps as they begin.

### Step 5 — Switch to CATIA and Follow the Selection Prompts

The tool requires three interactive selections inside CATIA. Switch to the CATIA window — you will see selection prompts appear in the CATIA status bar at the bottom of the screen.

#### Prompt 1: Select the Outer Tube Surface

- **What to click:** The main cylindrical face of the pipe/tube.
- **What happens:** The tool creates a tangent-continuity **Extract** of the entire pipe surface (including bends). This extract captures all connected faces belonging to the pipe.

#### Prompt 2: Select a Circular Edge

- **What to click:** One of the circular edges at the end of the pipe, or at the junction of a bend.
- **What happens:** The tool measures the **radius** of this edge. It then scans every edge on the extracted surface and generates a center point (`PointCenter`) for each edge whose radius matches. These center points form the raw centerline of the pipe.

#### Prompt 3: Select the Starting Point

- **What to click:** One of the generated center points — pick the point at the **end** of the pipe where you want the centerline to begin.
- **What happens:** The tool uses this point as the first node and sorts all remaining center points into a continuous path using nearest-neighbour ordering.

> **Tip:** All generated center points will be highlighted in CATIA for you to choose from. Pick a point at one of the two pipe ends (not in the middle).

### Step 6 — Watch the Automatic Processing

After the three selections, the tool processes automatically. Watch the log panel for progress updates:

1. **Edge scanning** — the tool scans all edges on the extract and creates center points for matching edges.
2. **Deduplication** — coincident/duplicate points are removed.
3. **Nearest-neighbour sorting** — points are ordered into a continuous path.
4. **Segment analysis** — each pair of consecutive points is classified as `STRAIGHT` or `ARC` based on the threshold.
5. **Line/Connect creation** — straight lines and Connect curves (for bends) are created.
6. **Joining** — all segments are joined into a single centerline spine.
7. **Body creation** — the final Rib, Surface, or Sweep is created.
8. **Final update** — the CATIA part tree is updated.

### Step 7 — Handle Bend Connect Failures (if any)

If a Connect curve cannot be created for a particular bend segment, a dialog box will appear:

```
Bend Connect Failed

Segment 5: PointCenter.12 → PointCenter.13
Distance: 18.4 mm

Connect curve could not be created for this bend.

Yes  →  Use a straight line (recommended)
No   →  Skip this segment (leave a gap)
```

- Click **Yes** to substitute a straight line for the failed bend. This is recommended — the result will be slightly less smooth at that bend but the overall pipe will still be valid.
- Click **No** to skip the segment entirely, which will leave a gap in the centerline. Only use this if you plan to manually fix the geometry in CATIA.

### Step 8 — Check the Result in CATIA

Once the log shows **"✓ Completed successfully."** and the status bar reads **"✓ Done"**:

1. Switch to CATIA.
2. Expand the feature tree — you should see a new `Construction_Geometry` geometrical set containing:
   - `Pipe_Surface_Extract` — the tangent-continuity extract of the pipe surface.
   - `Reference_Edge_Extract` — the edge you selected.
   - Center points (`PointCenter.X`) — the detected centerline nodes.
   - `Spine_Line_X` — straight segments of the centerline.
   - `Bend_Connect_X` — curved bend segments.
   - `Centerline_Spine` — the joined single-curve centerline.
   - `Pipe_Profile_Circle` — the profile used for the sweep/rib (if applicable).
3. Depending on the selected mode:
   - **Rib mode:** A new body named `Pipe_Solid_Body` containing `Pipe_Rib`.
   - **Surface mode:** Sweep, Fill, and Join surfaces, potentially solidified into a body.
   - **Sweep Only:** Just the Sweep surface in the geometrical set.

---

## 4. Understanding the Log Output

The log panel provides real-time feedback on every stage of the pipeline. Below are the key log line types you will see, with examples.

### 4.1 Connection & Setup

```
Connected to CATIA V5 — MyPipe.CATPart
Creating construction geometry set...
Created geometrical set: Construction_Geometry
```

Confirms a successful COM connection to the running CATIA instance and creation of the working geometrical set.

### 4.2 Selection Prompts

```
Step 1/3: Building centerline spine — follow the prompts in CATIA:
   1) Select the outer tube surface
   2) Select one circular edge (sets the reference radius)
   3) Select the starting point of the centerline
```

Tells you the tool is now waiting for your selections inside CATIA. Switch to the CATIA window.

### 4.3 Edge Scan Progress

```
   [INFO] Scanning 142 edges on pipe extract (ref radius=8.500)...
   [INFO] Edge scan done: 142 total, 38 no-radius, 24 matched, 24 points created.
```

- **142 total** — the total number of edges found on the extracted pipe surface.
- **38 no-radius** — edges that are not circular (e.g., seam edges, straight edges) — safely ignored.
- **24 matched** — circular edges whose radius matched the reference edge.
- **24 points created** — center points successfully generated.

### 4.4 Deduplication & Point Count

```
   [INFO] Generated 12 center points
```

After removing duplicate/coincident points, this is the final count of unique center points that will form the centerline.

### 4.5 Segment Distance Analysis Table

```
   [INFO] Distance threshold: 20.0 mm
   [INFO] Logic: Distance < 20.0 = ARC (connect), Distance >= 20.0 = STRAIGHT (line)
   [INFO] Analyzing segments...
   [INFO] ==================================================
   [INFO] SEGMENT DISTANCES:
   [INFO] Segment   1: PointCenter.3 → PointCenter.7
   [INFO]              Distance:   45.230 mm  [STRAIGHT (line)]
   [INFO] Segment   2: PointCenter.7 → PointCenter.11
   [INFO]              Distance:   12.841 mm  [ARC (connect)]
   [INFO] Segment   3: PointCenter.11 → PointCenter.15
   [INFO]              Distance:   62.500 mm  [STRAIGHT (line)]
   [INFO] ==================================================
```

This is the most important diagnostic output. Each segment shows:
- The **two points** being connected.
- The measured **distance** between them in millimeters.
- The **classification**: `STRAIGHT (line)` or `ARC (connect)`, determined by comparison against the threshold.

### 4.6 Line & Connect Creation

```
   [INFO] ✓ Created 8 lines (straight segments, distance >= 20.0 mm)
   [INFO] ℹ 3 arc segments to process (distance < 20.0 mm)
   [INFO] Creating connects for arc segments...
   [INFO] ✓ Created 3 connects for arcs, 0 fallback lines
```

Summary of what was created: how many straight lines, how many Connect curves for bends, and how many fallback straight lines (substituted for failed Connect curves).

### 4.7 Join Result

```
   [INFO] ✓ Joined 11 curves into single path
```

Confirms that all straight lines and Connect curves were successfully joined into a single continuous centerline spine.

### 4.8 Final Body Creation

```
Step 2/3: Building final body (mode = rib)...
   [INFO] Solid profile circle created (radius=8.50 mm).
   [INFO] Rib body created in 'Pipe_Solid_Body'.
  ✓ Solid rib created in body 'Pipe_Solid_Body'.
```

Confirms the final pipe geometry was created in the selected mode.

### 4.9 Warnings

```
  ⚠ Final update warning: <error message>
```

A warning during the final `Part.Update()` call. This is usually non-critical — the geometry may still be valid. Check the CATIA feature tree for any broken features (highlighted in orange/red).

---

## 5. Troubleshooting

### "Could not connect to CATIA V5. Is it running with a Part open?"

| Cause | Solution |
|---|---|
| CATIA V5 is not running | Launch CATIA V5 and open a `.CATPart` file before clicking **Run**. |
| No Part document is active | Open or create a `.CATPart` in CATIA. The active document must be a Part, not a Product or Drawing. |
| Python/CATIA architecture mismatch | Ensure both Python and CATIA are the same bit-width (both 64-bit or both 32-bit). |

### "Edge search on extract failed"

| Cause | Solution |
|---|---|
| The selected surface did not extract correctly | Ensure you clicked on the main cylindrical face of the pipe. Avoid selecting internal faces, planar faces, or edges. |
| The extract captured no edges | Try selecting a different face on the pipe, or ensure the pipe geometry is a proper solid/surface. |

### "Could not measure radius of selected edge"

| Cause | Solution |
|---|---|
| The selected edge is not circular | Select a clearly circular edge at the pipe end or at a bend. Avoid seam edges, straight edges, or half-arcs. |
| SPA workbench issue | Restart CATIA and try again. |

### "UpdateObject failed" or Part Update Errors

| Cause | Solution |
|---|---|
| Geometry conflicts in the CATIA model | Try adjusting the arc threshold — a different threshold may produce segments that don't conflict with existing geometry. |
| Connect curve self-intersection | Lower the arc threshold to reduce the number of connect curves (more segments become straight lines). |

### Multiple Failed Connects

If many bend connects fail (you see multiple dialog boxes asking about fallback lines):

- **Try a lower arc threshold** (5–15 mm). This reduces the number of segments classified as arcs, converting more of them to straight lines. Fewer connects = fewer potential failures.
- **Try a different starting point** if you re-run. The nearest-neighbour sorting order can affect how segments align with the available geometry.

### Pipe Geometry Doesn't Look Right

| Symptom | Solution |
|---|---|
| Pipe has kinks or sharp bends | Lower the arc threshold to reduce connect curves, or raise it to make more segments arcs for smoother transitions. |
| Pipe radius is wrong | Check the Pipe Radius field — if it's non-zero, it overrides auto-detection. Set to `0` to use the detected radius. |
| Centerline follows the wrong path | Try selecting a different starting point. The nearest-neighbour sorting can produce different paths depending on where you start. |
| Pipe ends are open when they should be closed | Check the Front End / Back End settings. Both must be set to **Capped** to produce a fully closed solid in Surface mode. |

---

## 6. Tips & Best Practices

### General Workflow

1. **Start with default settings** and adjust only if the result is not satisfactory. The defaults (`Radius = 2.0`, `Threshold = 20.0`, `Solid` pipe, both ends `Open`, `Rib` mode) work well for many common pipe geometries.

2. **Run the tool once with defaults**, examine the log output and the generated geometry in CATIA, then adjust parameters and re-run if needed. Each re-run deletes and recreates the `Construction_Geometry` set, so you get a clean start.

### Arc / Bend Threshold Selection

| Pipe Type | Recommended Threshold | Rationale |
|---|---|---|
| Complex pipes with many tight bends | **10–15 mm** | A lower threshold means fewer segments are classified as arcs, reducing the chance of Connect failures. More segments become straight lines, which are always reliable. |
| Simple pipes with few, gentle bends | **25–40 mm** | A higher threshold allows more segments to be treated as smooth arcs, producing a more faithful reproduction of gentle curves. |
| Pipes with very long straight runs | **30–50 mm** | Large gaps between center points on straight sections need to stay above the threshold so they are not misclassified as bends. |

### Edge Selection

- **Always select a clean, full circular edge** for the reference radius — not a half-arc, seam edge, or a boundary that only partially encircles the pipe cross-section.
- The best edges to select are at the pipe's open ends or at the junction between a straight section and a bend.

### Starting Point Selection

- **Pick a center point at one of the two pipe ends**, not one in the middle. Starting from an end gives the nearest-neighbour sort the best chance of producing a correct path through the pipe.
- If the resulting centerline path looks wrong (e.g., it doubles back on itself), re-run and try starting from the **opposite end**.

### Radius Parameter

- **Set Pipe Radius to `0`** to auto-detect the radius from the edge you select. This is the recommended default.
- **Enter a positive value** only when you want to deliberately override the detected radius (e.g., to create a pipe with a different cross-section size than the original).

### Mode Selection

| Use Case | Recommended Mode |
|---|---|
| Quick solid pipe, simplest workflow | **Rib (Solid)** |
| Need end caps, surface decomposition, or hollow shell | **Surface (Sweep + Fill + Join)** |
| Just want to inspect the sweep before committing | **Sweep Only** |

### Performance

- The **edge scan** is the slowest step (each edge is individually extracted and measured). For pipes with hundreds of edges, this can take several minutes. The log will show progress as each edge is processed.
- If the process seems stuck, check the CATIA window — CATIA may be showing an error dialog or waiting for input.

---

*This guide is part of the CATIA Parametric Pipe Builder project. For the project README and developer documentation, see [README.md](README.md).*
