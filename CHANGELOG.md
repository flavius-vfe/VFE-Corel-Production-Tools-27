# Changelog

## 0.4.0 — Repository release

- Added single visual launcher.
- Added separate Dimension Tool and Export Workflow interfaces.
- Added current-page and all-pages recursive text-to-curves commands.
- Added page-preservation, fitted-page, margin, and custom-size modes for CDR and PDF.
- Added saved/custom export profiles, preflight, progress, cancellation, and logging.
- Added VFE branding and copyright notice.
- Added proprietary licence and EULA.
- Added prebuilt GMS distribution and source-build documentation.

### Repository packaging notes

- Source filenames were normalized to remove upload suffixes.
- Source headers reference VFE Proprietary Software License v1.1.
- The source UI installer exposes only `VFE_InstallVisualInterface`.
- The supplied prebuilt GMS is preserved byte-for-byte and should be rebuilt from repository source when exact binary/source parity is required.
