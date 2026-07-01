import win32com.client
import os
import tempfile
import subprocess
import time  # Added this to give Windows time to catch up

# ==========================================
# PHASE 2: THE DATA STRUCTURE 
# ==========================================
hose_data = [
    {"point": 1, "x": 0.0, "y": 0.0, "z": 0.0, "outer_dia": 25.0, "inner_dia": 20.0},
    {"point": 2, "x": 100.0, "y": 50.0, "z": 20.0, "outer_dia": 25.0, "inner_dia": 20.0},
    {"point": 3, "x": 150.0, "y": 150.0, "z": 50.0, "outer_dia": 20.0, "inner_dia": 15.0},
    {"point": 4, "x": 250.0, "y": 100.0, "z": 80.0, "outer_dia": 20.0, "inner_dia": 15.0}
]

# ==========================================
# PHASE 3: THE HANDSHAKE
# ==========================================
def connect_to_catia():
    print("Attempting to connect to CATIA...")
    try:
        catia = win32com.client.Dispatch("CATIA.Application")
        catia.Visible = True
        catia.Documents.Add("Part")
        part_document = catia.ActiveDocument
        part = part_document.Part

        try:
            part_document.Product.PartNumber = "AC_Hose_Automated"
        except:
            pass

        hybrid_bodies = part.HybridBodies
        wireframe_set = hybrid_bodies.Add()
        wireframe_set.Name = "Hose_Wireframe"

        return catia, part, wireframe_set
    except Exception as e:
        print(f"FAILED to connect: {e}")
        return None, None, None

# ==========================================
# PHASE 4a: THE SPINE
# ==========================================
def create_spine(part, wireframe_set, data):
    print("Generating 3D points and centerline...")
    hs_factory = part.HybridShapeFactory
    spline = hs_factory.AddNewSpline()
    spline.SetSplineType(0)

    point_references = []
    for node in data:
        pt = hs_factory.AddNewPointCoord(node["x"], node["y"], node["z"])
        pt.Name = f"Node_{node['point']}"
        wireframe_set.AppendHybridShape(pt)

        pt_ref = part.CreateReferenceFromObject(pt)
        spline.AddPoint(pt_ref)
        point_references.append(pt_ref)

    spline.Name = "Hose_Centerline"
    wireframe_set.AppendHybridShape(spline)
    part.Update()
    return spline, point_references

# ==========================================
# PHASE 4b: THE PLANES
# ==========================================
def create_planes(part, wireframe_set, spline, point_references):
    print("Generating sketch planes...")
    hs_factory = part.HybridShapeFactory
    spline_ref = part.CreateReferenceFromObject(spline)
    plane_references = []

    for idx, pt_ref in enumerate(point_references):
        plane = hs_factory.AddNewPlaneNormal(spline_ref, pt_ref)
        plane.Name = f"Profile_Plane_{idx + 1}"
        wireframe_set.AppendHybridShape(plane)

        plane_ref = part.CreateReferenceFromObject(plane)
        plane_references.append(plane_ref)

    part.Update()
    return plane_references

# ==========================================
# PHASE 4c: THE MASTER PROFILE
# ==========================================
def create_master_profile(part, plane_references, data):
    print("Drawing the hollow master profile...")
    part.InWorkObject = part.MainBody
    sketches = part.MainBody.Sketches

    first_plane_ref = plane_references[0]
    sketch = sketches.Add(first_plane_ref)
    sketch.Name = "Hose_Master_Profile"

    factory_2d = sketch.OpenEdition()

    outer_rad = data[0]["outer_dia"] / 2.0
    inner_rad = data[0]["inner_dia"] / 2.0

    factory_2d.CreateClosedCircle(0.0, 0.0, outer_rad)
    factory_2d.CreateClosedCircle(0.0, 0.0, inner_rad)

    sketch.CloseEdition()
    part.Update()
    return sketch

# ==========================================
# PHASE 5: THE SOLID RIB (EXTERNAL SCRIPT)
# ==========================================
def create_solid_rib():
    print("Sweeping the 2D sketch into a 3D solid Rib via VBScript...")
    try:
        # 1. Write the VBScript exactly as you designed it
        vba_code  = "Set CATIA = GetObject(, \"CATIA.Application\")\n"
        vba_code += "Dim part1\n"
        vba_code += "Set part1 = CATIA.ActiveDocument.Part\n"
        vba_code += "Dim shapeFactory1\n"
        vba_code += "Set shapeFactory1 = part1.ShapeFactory\n"
        vba_code += "part1.InWorkObject = part1.MainBody\n"

        # --- Sketch → Reference (THE FIX: AddNewRib needs References, not raw objects) ---
        vba_code += "Dim sketch1\n"
        vba_code += "Set sketch1 = part1.MainBody.Sketches.Item(\"Hose_Master_Profile\")\n"
        vba_code += "Dim refProfile\n"
        vba_code += "Set refProfile = part1.CreateReferenceFromObject(sketch1)\n"

        # --- Spline → Reference (was already correct) ---
        vba_code += "Dim wireframe\n"
        vba_code += "Set wireframe = part1.HybridBodies.Item(\"Hose_Wireframe\")\n"
        vba_code += "Dim spline1\n"
        vba_code += "Set spline1 = wireframe.HybridShapes.Item(\"Hose_Centerline\")\n"
        vba_code += "Dim refCurve\n"
        vba_code += "Set refCurve = part1.CreateReferenceFromObject(spline1)\n"

        # --- AddNewRib: both args are now proper Reference objects ---
        vba_code += "Dim rib1\n"
        vba_code += "Set rib1 = shapeFactory1.AddNewRibFromRef(refProfile, refCurve)\n"
        vba_code += "rib1.Name = \"Hose_3D_Sweep\"\n"
        vba_code += "part1.Update\n"
        vba_code += "WScript.Echo \"3D HOLLOW HOSE COMPLETE!\"\n"

        # 2. Save it to a temporary file
        temp_dir = tempfile.gettempdir()
        vbs_path = os.path.join(temp_dir, "build_rib.vbs")

        with open(vbs_path, "w", encoding="ascii") as f:
            f.write(vba_code)

        # 3. CRITICAL: Wait 1 second so Windows recognizes CATIA is ready
        time.sleep(1)

        # 4. Run it externally via cscript
        result = subprocess.run(
            ["cscript", "//Nologo", vbs_path],
            capture_output=True, text=True
        )

        print("Output:", result.stdout.strip())
        if result.stderr:
            print("Script Error:", result.stderr)

        # 5. Clean up the file
        os.remove(vbs_path)

    except Exception as e:
        print(f"Error executing VBScript: {e}")



# ==========================================
# RUN THE SCRIPT
# ==========================================
if __name__ == "__main__":
    catia_app, my_part, my_wireframe = connect_to_catia()

    if my_part is not None:
        my_spline, my_pt_refs = create_spine(my_part, my_wireframe, hose_data)

        if my_spline and my_pt_refs:
            my_plane_refs = create_planes(my_part, my_wireframe, my_spline, my_pt_refs)

            if my_plane_refs:
                my_sketch = create_master_profile(my_part, my_plane_refs, hose_data)

                if my_sketch:
                    # Fire the external script!
                    create_solid_rib()
