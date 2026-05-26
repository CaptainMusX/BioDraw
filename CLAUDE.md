# BioDraw — AI Development Guide

## Project Overview

BioDraw is a PowerPoint VSTO add-in (.NET Framework 4.7.2) that adds a custom Ribbon tab. The project type is VSTO Office Add-in; it can only be built in Visual Studio with the Office workload installed.

## File Architecture (Post Code-Split)

| File | Role |
|---|---|
| `BioDraw/Ribbon1.cs` | Ribbon UI callbacks, preset UI, material library UI, settings persistence (~4200 lines) |
| `BioDraw/Ribbon1.xml` | Ribbon XML defining the BioDraw tab layout |
| `BioDraw/Models.cs` | Data classes: `ImageReplacePreset`, `MaterialEntry` |
| `BioDraw/ImageReplacePipeline.cs` | Core image color replacement pipeline (ImageMagick execution, shape replacement, geometry restoration, animation preservation) |
| `BioDraw/MaterialLibraryService.cs` | Material library: file scanning, thumbnail generation, material insertion |
| `BioDraw/PresetManager.cs` | Static utilities for preset management, color option handling, config parsing |
| `BioDraw/NativeMethods.cs` | Win32 P/Invoke declarations for screen color picking |
| `BioDraw/PictureConverter.cs` | AxHost-based `Image → stdole.IPictureDisp` converter |

## COM Interop Constraints

The project uses `Microsoft.Office.Core` COM reference v2.7 with `EmbedInteropTypes=true`. The following types are NOT available and must NOT be used as property types or static enum references:

- `MsoAnimEffect`
- `MsoAnimTriggerType`
- `MsoAnimateByLevel`

These can only be used as `(int)` casts on `dynamic` values, never as typed properties. `CapturedEffect` uses `int` for `EffectType` and `TriggerType` for this reason. The `msoAnimateLevelNone` value is `0`.

Safe types (available in v2.7): `MsoTriState`, `MsoZOrderCmd`.

---

## Image Replacement Pipeline — CRITICAL SECTION

This section describes the core "图色替换" (Image Color Replacement) feature. **Do not reorder or refactor the geometry restoration sequence without understanding every constraint below.**

### High-Level Flow

1. **Extract original image** from the PPTX zip (reads the slide XML to find the media reference, then extracts the embedded file from the ZIP)
2. **Run ImageMagick** on the extracted image (`-fuzz N% -transparent/opaque`) with configurable `imageMagickPath`
3. **Replace the shape** on the slide: delete old picture shape, insert new picture via `AddPicture`, restore all properties

### Shape Replacement — Geometry Restoration Order (DO NOT CHANGE)

The method `TryReplaceShapeImageInPlace` in `ImageReplacePipeline.cs` restores a deleted shape's geometry onto a newly inserted picture. The order of property restoration is **critical and non-negotiable**:

```
1. AddPicture(outputPath, ..., insertionLeft, insertionTop, insertionWidth, insertionHeight)
2. newShape.Name = shapeName
3. newShape.Rotation = rotation
4. newShape.AlternativeText = altText
5. Hyperlink restoration (ActionSettings)
6. newShape.LockAspectRatio = 0          ← TEMPORARILY DISABLE
7. newShape.PictureFormat.CropLeft/Top/Right/Bottom = savedCrop   ← SET CROP FIRST
8. newShape.Left/Top/Width/Height = savedValues                    ← THEN GEOMETRY
9. newShape.LockAspectRatio = savedLock  ← RESTORE ORIGINAL LOCK
10. newShape.Apply()                      ← PickUp formatting LAST
11. RestoreShapeEffects (animations)
12. Z-Order fix (bidirectional)
```

### Why This Order Matters (Known Bugs)

**Bug 1 — Crop stretch**: Insertion dimensions for cropped shapes use the "external model" formula:
```csharp
insertionWidth  = width + cropLeft + cropRight;   // width = visible frame width
insertionHeight = height + cropTop + cropBottom;
```
This creates a shape larger than the visible area. Step 7-8 collapses it back by setting crop values, then Width/Height. PPT reconciles the two to produce the correct visible area. If Width/Height are NOT set (were skipped in an earlier broken version), the shape stays at the oversized insertion dimensions.

**Bug 2 — LockAspectRatio interference**: If `LockAspectRatio` is not disabled (step 6), setting Width (step 8) on a locked shape causes PPT to auto-adjust Height to maintain the image's natural aspect ratio. The image's natural ratio may differ from the shape's current ratio if the user applied non-proportional scaling. Disabling the lock during geometry ensures Width/Height are set independently to their exact saved values.

**Bug 3 — PickUp/Apply before geometry**: `Shape.PickUp()` captures visual formatting which may include crop/picture format data. If `Apply()` runs before geometry restoration (step 10 before 6-9), the old crop values are applied to the uncorrected new shape dimensions. Subsequent Width/Height changes trigger PPT's internal proportional adjustment, permanently distorting the visible area. `Apply()` MUST run after all geometry is finalized.

### Insertion Math Constants

For **non-cropped** shapes: `insertionLeft = left`, `insertionWidth = width` (standard model)

For **cropped** shapes (hasCrop = any `|crop| > 0.01`):
```csharp
insertionLeft   = left - cropLeft;
insertionTop    = top - cropTop;
insertionWidth  = Math.Max(1f, width + cropLeft + cropRight);
insertionHeight = Math.Max(1f, height + cropTop + cropBottom);
```

This formula assumes PowerPoint's `shape.Width` returns the **visible** frame width and crop values extend **beyond** the shape bounds (external model). The crop restoration + Width/Height combination reconciles this to the correct visible result.

### Other Preserved Properties

| Property | How Restored |
|---|---|
| Visual formatting (shadow, glow, reflection, 3D) | `shape.PickUp()` before delete → `newShape.Apply()` after geometry |
| Animations | `CaptureShapeEffects()` iterates `TimeLine.MainSequence` + `InteractiveSequences`, `RestoreShapeEffects()` calls `mainSeq.AddEffect()` |
| AlternativeText | Saved from `shape.AlternativeText` (fallback: `shape.Title`) |
| Hyperlink | Saved from `shape.ActionSettings[1].Hyperlink` (Address, SubAddress) |
| Z-Order | Bidirectional msoSendBackward/msoBringForward loop with 2048 guard |
| Grouped pictures | Recursive `FlattenToPictures()` — group-internal pictures show error, not silently broken |

### ImageMagick Output Format

- Non-transparent mode: preserves input extension
- Transparent mode: switches JPEG/BMP to PNG (transparency unsupported in those formats)
- JPEG output: `-quality 100 -sampling-factor 4:4:4` (no chroma subsampling)

### Material Library

- `MaterialLibraryService.TryInsertMaterialToCurrentSlide()` calls `Ribbon1.TryExecuteMso()` which must remain `internal static`
- Material preview cache limit: 300 entries, LRU eviction clears half
- `materialFileExtensions` field was removed during code split; extensions are now in `MaterialLibraryService.MaterialFileExtensions`

### Build Notes

- Project type: VSTO Office Add-in, `TargetFrameworkVersion: v4.7.2`
- Output type: Library (`.dll`)
- Host application: PowerPoint
- Cannot be built with `dotnet build` — requires Visual Studio with Office/SharePoint workload
- Debug: attaches to `powerpnt.exe`
