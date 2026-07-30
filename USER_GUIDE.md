# User Guide

## Open the launcher

Run:

```text
VFE_CorelProductionTools27.VFE_CorelProductionTools27
```

The launcher provides four separate tools.

## 1. Dimension Tool

Use this tool before exporting when dimensions must be visible in the final artwork.

### Measurement target

- **Every selected object separately** — each selected object receives its own dimensions.
- **Entire selection as one object** — the combined selection bounding box receives one overall measurement.

### Dimension choices

- Width
- Height
- Width and height
- Display units: mm, cm, or inches
- Decimal precision
- Font name and font size
- Distance between objects and dimension lines
- Colour
- Remove existing objects from the `DIMENSIONS` layer
- Group newly created dimension objects

Dimensions are added to the active document so they can be inspected or adjusted before export.

## 2. Export Workflow

### Workflow profiles

Use built-in settings or save custom profiles. Profiles contain export settings only; dimensions are prepared separately.

### Process scope

- Each page as a separate file
- All pages in one document
- Current page

### Filename source

- Object named `EXPORT_NAME`
- Document name
- Page name
- Custom prefix

For reliable page-by-page naming, create a text object named `EXPORT_NAME` on each relevant page.

### Output formats

- Current-version CDR
- Version 15 CDR
- Curved current-version CDR
- Curved version 15 CDR
- Curved PDF

### CDR page handling

- Preserve the source page
- Fit grouped objects with no margin
- Fit grouped objects with a custom margin
- Custom page size while keeping object positions
- Custom page size and centre objects

### PDF page handling

- Fit grouped objects with no margin
- Fit grouped objects with a custom margin
- Preserve the source page
- Custom page size while keeping object positions
- Custom page size and centre objects

For fitted modes, page objects are grouped temporarily in the export copy and treated as one object for page-size calculation. The source document should not remain grouped after export.

### DIMENSIONS layer

Enable **Include DIMENSIONS layer in fitted page size** when existing dimensions must be included in the page bounding box.

### Output safety

- Choose a dedicated output folder.
- Enable automatic unique names to add `_2`, `_3`, and similar suffixes when files already exist.
- Review the workflow summary before running.
- Run Preflight before important production jobs.
- Press Cancel or Esc when a running workflow must be stopped.

## 3. Convert Text to Curves — Current Page

This changes text on the active page into curves, including supported text inside nested groups and PowerClips.

This operation modifies the source document. Save an editable copy first.

## 4. Convert Text to Curves — All Pages

This converts supported text throughout the document, including nested groups and PowerClips.

This operation modifies the source document. Save an editable copy first.

## Recommended production sequence

1. Save an editable source copy.
2. Add dimensions manually if required.
3. Inspect dimensions and object positions.
4. Open Export Workflow.
5. Load or configure a profile.
6. Run Preflight.
7. Confirm filenames, output formats, page handling, and output folder.
8. Run Workflow.
9. Inspect all generated files before delivery.
