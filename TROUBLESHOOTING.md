# Troubleshooting

## The GMS does not appear in CorelDRAW

- Use **Tools → Scripts → Scripts → Visual Basic for Applications → Load** and select the GMS explicitly.
- Confirm the file extension is `.gms`, not `.gms.zip`.
- For automatic loading, confirm the file is in the current user's `Draw\GMS` folder.
- Restart CorelDRAW after copying the file.

## The macro is blocked

Only enable macros from a trusted VFE release. Review CorelDRAW's VBA security settings. Do not reduce security globally unless you understand the risk.

## The launcher does not open

Run the fully qualified entry point:

```text
VFE_CorelProductionTools27.VFE_CorelProductionTools27
```

If the form is missing in a source-built project, rebuild the forms using `VFE_InstallVisualInterface` before removing the installer module.

## The interface installer fails

The installer requires programmatic access to the active VBA project.

- Confirm the correct standalone GMS project is selected.
- Do not install into `GlobalMacros.gms`.
- Temporarily enable **Trust access to Visual Basic project** during the controlled build process.
- Confirm both source modules are in the same project.

## The wrong project receives the imported module

In Project Explorer, select the standalone VFE project immediately before every **File → Import File** operation. Remove accidental copies from `GlobalMacros.gms`.

## Run-time error 13 while adding dimensions

Use the current engine source or GMS release. An earlier build incorrectly resolved a page's parent document and could produce a type mismatch.

## Exported page size is not correct

- Verify the selected CDR and PDF page modes independently.
- Use **Preserve source page** when the original page dimensions and object positions must remain unchanged.
- Use fitted modes only when the page should be calculated from grouped object bounds.
- Confirm whether the `DIMENSIONS` layer should participate in fitted bounds.
- Check custom width and height values are positive and in millimetres.

## Text inside a PowerClip was not converted

- Confirm the object is actual CorelDRAW text.
- Check whether the object or containing layer is locked or protected.
- Test on a copy and report the object structure if recursive conversion skips it.

## A file already exists

Enable unique filenames so the tool adds `_2`, `_3`, and later suffixes, or choose a clean output folder.

## A workflow stops partway through

- Check the export log in the selected output folder.
- Confirm the folder is writable.
- Check available disk space.
- Test the failing page separately.
- Look for invalid filenames, locked objects, unsupported effects, or a damaged document.

## Profiles do not appear on another computer

Saved profile settings may be stored in the current Windows user's VBA/application settings rather than inside the GMS. Recreate or re-save profiles for each Windows account as needed.

## Report a bug

Include:

- CorelDRAW full version/build;
- Windows version;
- VFE tool version;
- exact steps;
- error number and message;
- screenshot;
- relevant workflow settings;
- a sanitized sample CDR when possible;
- export log.
