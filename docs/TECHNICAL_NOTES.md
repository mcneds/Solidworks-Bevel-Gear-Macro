# Technical Notes

## Geometry

The macro uses an approximate straight-bevel construction intended for CAD and printing:

- conical lofted blank;
- standard involute-derived angular tooth-space calculation at the large end;
- corresponding radial scaling toward the small end;
- lofted tooth-space cut;
- circular feature pattern.

Gear 2 is pre-phased in pair mode by shifting the tooth-space sketches by half one tooth pitch.

## Pair pitch-cone calculation

For shaft angle `sigma` and tooth counts `z1`, `z2`, the macro uses:

```text
delta1 = atan2(sin(sigma), z2/z1 + cos(sigma))
delta2 = atan2(sin(sigma), z1/z2 + cos(sigma))
```

For equal tooth counts at 90 degrees:

```text
delta1 = delta2 = 45 degrees
```

## SOLIDWORKS feature/API approach

The current implementation uses feature-based SOLIDWORKS construction including:

- reference planes;
- sketches;
- boss loft;
- lofted cuts;
- reference axis;
- circular feature patterns;
- Move/Copy Body features.

The source contains additional validation around COM selection behavior because SOLIDWORKS selection-return booleans do not always reflect the resulting SelectionManager state reliably.

## Distribution

The final `.swp` must be saved by SOLIDWORKS itself. The source files in `/src` are provided so the binary macro can be audited and rebuilt.
