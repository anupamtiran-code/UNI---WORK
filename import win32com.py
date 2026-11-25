import math
import win32com.client
import pythoncom

# ---------- dimensions (mm) - change these ----------
DISC_DIAMETER   = 120
DISC_THICKNESS  = 15
HUB_DIAMETER    = 60
HUB_HEIGHT      = 20
BORE_DIAMETER   = 30
BOLT_COUNT      = 6
BOLT_CIRCLE_DIA = 90
BOLT_HOLE_DIA   = 10
FILLET_RADIUS   = 2
SAVE_PATH       = r"C:\Users\anupa\breakout-game\flange.SLDPRT"

def mm(value):
    """SolidWorks API uses meters."""
    return value / 1000.0

# ---------- helpers ----------
def connect():
    sw = win32com.client.Dispatch("SldWorks.Application")
    sw.Visible = True
    return sw

def new_part(sw):
    template = sw.GetUserPreferenceStringValue(8)  # default part template
    return sw.NewDocument(template, 0, 0, 0)

def select(model, name, kind, x=0, y=0, z=0, append=False):
    callout = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    return model.Extension.SelectByID2(name, kind, x, y, z, append, 0, callout, 0)

def start_sketch(model, plane="Front Plane"):
    """Sketch on a named plane."""
    model.ClearSelection2(True)
    select(model, plane, "PLANE")
    model.SketchManager.InsertSketch(True)

def start_sketch_on_face(model, x, y, z):
    """Sketch on the face that contains the point (x, y, z) in meters."""
    model.ClearSelection2(True)
    ok = select(model, "", "FACE", x, y, z)
    if not ok:
        raise RuntimeError(f"Could not select a face at ({x}, {y}, {z})")
    model.SketchManager.InsertSketch(True)

def end_sketch(model):
    model.SketchManager.InsertSketch(True)

def extrude(model, depth_m):
    """Add material, blind, depth in meters."""
    return model.FeatureManager.FeatureExtrusion2(
        True, False, False, 0, 0, depth_m, 0,
        False, False, False, False, 0, 0,
        False, False, False, False,
        True, True, True, 0, 0, False)

def cut(model, depth_m, flip=False):
    """Remove material, blind, depth in meters."""
    return model.FeatureManager.FeatureCut3(
        True, flip, False, 0, 0, depth_m, 0,
        False, False, False, False, 0, 0,
        False, False, False, False,
        True, True, True, True, False, False)

# ---------- build the flange ----------
sw = connect()
model = new_part(sw)

# 1) Base disc on the Front Plane
start_sketch(model, "Front Plane")
model.SketchManager.CreateCircleByRadius(0, 0, 0, mm(DISC_DIAMETER / 2))
end_sketch(model)
extrude(model, mm(DISC_THICKNESS))
print("Step 1: base disc done")

# 2) Raised hub on top of the disc
#    (pick a point on the disc's top face, outside the hub area)
face_x = mm(DISC_DIAMETER / 2 * 0.8)
try:
    start_sketch_on_face(model, face_x, 0, mm(DISC_THICKNESS))
except RuntimeError:
    # fallback: sketch on the Front Plane and extrude from an offset (see notes)
    raise
model.SketchManager.CreateCircleByRadius(0, 0, 0, mm(HUB_DIAMETER / 2))
end_sketch(model)
extrude(model, mm(HUB_HEIGHT))
print("Step 2: hub done")

# 3) Center bore, cut from the top of the hub all the way down
top_z = mm(DISC_THICKNESS + HUB_HEIGHT)
start_sketch_on_face(model, mm(HUB_DIAMETER / 2 * 0.7), 0, top_z)
model.SketchManager.CreateCircleByRadius(0, 0, 0, mm(BORE_DIAMETER / 2))
end_sketch(model)
cut(model, mm(DISC_THICKNESS + HUB_HEIGHT))
print("Step 3: center bore done")

# 4) Bolt holes on the disc's top face (outside the hub)
start_sketch_on_face(model, face_x, 0, mm(DISC_THICKNESS))
bolt_r = mm(BOLT_CIRCLE_DIA / 2)
hole_r = mm(BOLT_HOLE_DIA / 2)
for i in range(BOLT_COUNT):
    angle = math.radians(i * 360 / BOLT_COUNT)
    cx = bolt_r * math.cos(angle)
    cy = bolt_r * math.sin(angle)
    model.SketchManager.CreateCircleByRadius(cx, cy, 0, hole_r)
end_sketch(model)
cut(model, mm(DISC_THICKNESS))
print("Step 4: bolt holes done")

# 5) Optional fillet where the hub meets the disc
try:
    model.ClearSelection2(True)
    select(model, "", "EDGE", mm(HUB_DIAMETER / 2), 0, mm(DISC_THICKNESS))
    model.FeatureManager.FeatureFillet(
        195, mm(FILLET_RADIUS), 0, 0, None, None, None)
    print("Step 5: fillet done")
except Exception as e:
    print("Step 5: fillet skipped:", e)

model.ClearSelection2(True)
model.ViewZoomtofit2()
model.SaveAs3(SAVE_PATH, 0, 1)
print("Saved:", SAVE_PATH)