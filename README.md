# SOLIDWORKS 2025 Straight Bevel Gear Generator

A VBA macro for **SOLIDWORKS 2025** that generates approximate straight bevel gears for CAD prototyping and 3D printing.

It can create either a single bevel gear or a calculated mating pair, including non-90° shaft angles.

> [!IMPORTANT]
> The tooth surfaces are a CAD/printing-oriented approximation. They are **not** true manufacturing-generated Gleason/Klingelnberg bevel tooth surfaces and should not be treated as production gear geometry.

## Features

- Single gear or mating-pair mode
- Straight bevel teeth
- Selectable shaft angle
- Selectable pressure angle
- Configurable face width
- Configurable common bore diameter
- Automatic mating-pair pitch-cone calculation
- Automatic pitch-diameter/module/tooth-count calculation
- Involute-based tooth-space approximation
- Automatic circular tooth-space pattern
- Automatic Gear 2 phase offset for a mating pair
- Automatic pair positioning when using the default origin plane
- Supports a selected reference plane or planar face for gear generation
- Input validation for face width, bore size, tooth count, and parameter consistency

## Tested With

- **SOLIDWORKS 2025**
- VBA macro project (`.swp`)

Other SOLIDWORKS versions may work, but they have not been verified.

---

## Recommended Download / Installation

The easiest distribution format is the ready-to-run:

```text
BevelGearGenerator.swp
```

Download the `.swp`, then in SOLIDWORKS:

1. Open or create a **Part**.
2. Go to **Tools → Macro → Run**.
3. Select `BevelGearGenerator.swp`.
4. Run `main`.

### Plane selection

Before starting the macro:

- Select a **reference plane** or **planar face** to generate the gear there.
- If nothing valid is selected, the macro uses the first origin reference plane.

For a mating pair, **automatic Gear 2 positioning is currently performed only when no custom plane/face was selected before launching the macro**. With a custom plane/face, both matched gears are generated but automatic 3D pair placement is skipped.

---

## Parameters

### Mode

- **Single gear**
- **Mating pair**

### Parameter locking

Module, tooth count, and pitch diameter obey:

\[
d = mz
\]

where:

- \(d\) = pitch diameter
- \(m\) = module
- \(z\) = tooth count

The macro therefore lets you lock any two and calculates the third:

- Module + Teeth → Pitch diameter
- Pitch diameter + Teeth → Module
- Module + Pitch diameter → Teeth

### Tooth count

Sets the number of teeth on each gear.

In pair mode, Gear 1 and Gear 2 can have different tooth counts.

### Shaft angle

The angle between the two gear shafts.

Typical miter pair:

```text
90°
```

The pitch-cone angles for a mating pair are calculated from the tooth counts and shaft angle.

### Pressure angle

Selectable involute pressure angle.

Default:

```text
20°
```

### Bore diameter

A single common bore diameter is used for both gears in pair mode.

### Face width

`Face width` is measured **along the bevel/cone surface**, not directly between the two planar end faces.

The approximate axial separation of those faces is:

\[
t_{\text{axial}} = W_{\text{face}}\cos(\delta)
\]

where:

- \(W_{\text{face}}\) = entered face width
- \(\delta\) = pitch-cone angle

For a 45° gear:

| Face width | Axial separation |
|---:|---:|
| 3 mm | 2.12 mm |
| 4 mm | 2.83 mm |
| 5 mm | 3.54 mm |

The macro warns when face width becomes large compared with the cone distance.

---

## Mating-Pair Geometry

For a pair with shaft angle \(\Sigma\), the macro calculates the two pitch-cone angles from the tooth counts.

For the common 90° case with equal tooth counts:

```text
Gear 1 = 45°
Gear 2 = 45°
```

Gear 2's tooth-space geometry is generated with a half-tooth phase offset so a tooth on one gear faces a space on the other.

Pair placement uses separate Move/Copy Body operations for:

1. shaft-angle rotation;
2. translation to the common bevel apex.

---

## How the Gear Is Built

The macro creates each gear using normal SOLIDWORKS features:

1. Create the outer and inner bevel reference profiles.
2. Loft the conical gear blank.
3. Capture the clean blank's bevel axis/apex.
4. Create the bore.
5. Generate an involute-derived tooth-space profile at the large end.
6. Scale the corresponding profile toward the small end.
7. Loft-cut the tooth space.
8. Circular-pattern the tooth-space cut.

This produces an editable feature tree rather than an imported mesh.

---

## 3D Printing Notes

These gears are intended primarily for prototyping and printable mechanisms.

The current macro creates **nominal tooth geometry**. It does not yet have a dedicated backlash/printing-clearance input.

For FDM printing, you may need to account for:

- printer dimensional accuracy;
- material shrinkage;
- layer orientation;
- bore clearance;
- tooth backlash;
- surface finish.

Do not assume two zero-clearance CAD gears will rotate freely straight off the printer.

---

## Building the `.swp` From Source

The `/src` directory contains the VBA source used by the macro.

### 1. Create a new macro

In SOLIDWORKS:

**Tools → Macro → New**

Save the project as:

```text
BevelGearGenerator.swp
```

### 2. Standard module

Import:

```text
src/modBevelGear.bas
```

or paste its contents into a standard module named:

```text
modBevelGear
```

### 3. UserForm

Choose:

**Insert → UserForm**

Leave/name it:

```text
UserForm1
```

Open the form's code window and paste the complete contents of:

```text
src/UserForm1_CODE.txt
```

The UI controls are generated dynamically at runtime, so you do **not** need to manually place buttons or text boxes.

### 4. Event class

Choose:

**Insert → Class Module**

Name it:

```text
cBevelFormEvents
```

Either:

- paste the contents of `src/cBevelFormEvents_CODE.txt`, or
- import `src/cBevelFormEvents.cls` directly.

Do not paste the exported `.cls` metadata lines such as `VERSION 1.0 CLASS` into a code window. Those lines are valid only when VBA imports the `.cls` file itself.

### 5. Compile

Run:

**Debug → Compile VBAProject**

Then run:

```text
modBevelGear.main
```

---

## Repository Layout

```text
.
├── BevelGearGenerator.swp        # recommended ready-to-run distributable
├── README.md
├── .gitattributes
├── src/
│   ├── modBevelGear.bas
│   ├── UserForm1_CODE.txt
│   ├── cBevelFormEvents.cls
│   └── cBevelFormEvents_CODE.txt
├── docs/
│   └── TECHNICAL_NOTES.md
└── dist/
    └── PLACE_WORKING_SWP_HERE.txt
```

The `.swp` is binary, so GitHub cannot show useful line-by-line diffs for it. Keeping the exported VBA source in `/src` makes changes reviewable.

---

## GitHub Distribution

A good release workflow is:

1. Commit the tested `.swp` in the repository root.
2. Keep the source files under `/src`.
3. Tag a version, for example:
   ```text
   v1.0.0
   ```
4. Create a **GitHub Release**.
5. Attach `BevelGearGenerator.swp` to the release.

That gives normal users one obvious file to download while still keeping the source visible.

---

## Current Limitations

- Tested specifically on SOLIDWORKS 2025.
- Approximate straight-bevel tooth geometry, not manufacturing-generated surfaces.
- No dedicated backlash/printing-clearance parameter yet.
- Automatic mating-pair placement currently assumes the default origin-plane construction path.
- A custom selected plane/face can be used to generate gears, but pair auto-positioning is skipped.
- Very small tooth counts can produce poor or undercut-prone geometry.
- Face width must remain physically reasonable relative to cone distance.

---

## Source Version

The source in this repository corresponds to the current working **v15.1** macro path.
