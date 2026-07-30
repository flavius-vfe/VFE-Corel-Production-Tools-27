# Packaging Audit

- GMS header begins with `GMS` and contains an OLE compound-document signature.
- GMS size: 658,450 bytes.
- Static string inspection found the expected forms: `frmVFELauncher`, `frmVFEDimensionTool`, and `frmVFEExportWorkflow`.
- Engine module name: `VFE_CorelProductionTools27`.
- Installer module name: `VFE_UI_Installer`.
- Repository installer exposes only `VFE_InstallVisualInterface`.
- Copyright branding is present in source.
- No CorelDRAW runtime execution or VBA compilation was possible in the packaging environment.

## Integrity

- `dist/VFE-Corel-Production-Tools-27.gms` — `bbf3ca829d7b8e213ab81fa552e69d578963dc71a8a461dff53ed2202770f577`
- `src/VFE_CorelProductionTools27.bas` — `0f07b61862aaa1de3289aceebc996231dfa4223c72858e2a1a95179b27ac20c4`
- `src/VFE_UI_Installer.bas` — `cdd4d425460b0efc126b905ed561aeba754cf2ce4b3589213331aede82fe7df6`
- `release-assets/VFE-Corel-Production-Tools-27-v0.4.0.zip` — `a9d3931cf2f64a0a32fc2e71b2960bd5570734765d20962ffc917cd2de76892f`
