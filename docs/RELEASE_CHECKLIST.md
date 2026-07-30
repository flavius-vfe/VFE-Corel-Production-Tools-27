# Release Checklist

## Source

- [ ] Engine and installer import into a standalone GMS project.
- [ ] Installer creates all three VFE UserForms.
- [ ] `Debug → Compile` succeeds without errors.
- [ ] Copyright text is correct.
- [ ] Installer module is removed from the final GMS.
- [ ] No code was built inside `GlobalMacros.gms`.

## Functional tests

- [ ] Launcher opens.
- [ ] Dimensions: one object.
- [ ] Dimensions: multiple objects separately.
- [ ] Dimensions: entire selection.
- [ ] Current-page text conversion.
- [ ] All-pages text conversion.
- [ ] PowerClip text conversion.
- [ ] Each-page separate export.
- [ ] All-pages export.
- [ ] Preserve original CDR page.
- [ ] Preserve original PDF page.
- [ ] Fit no margin.
- [ ] Fit custom margin.
- [ ] Custom page size, keep positions.
- [ ] Custom page size, centre objects.
- [ ] Existing-file naming.
- [ ] Cancellation.
- [ ] Export log.
- [ ] Custom profile survives form restart.

## Packaging

- [ ] GMS copied while CorelDRAW is closed or through Scripts docker Copy To.
- [ ] Release GMS tested independently from development project.
- [ ] SHA-256 manifest generated.
- [ ] Licence, EULA, copyright, and installation files included.
- [ ] No confidential CDR documents included.
- [ ] GitHub release notes match the embedded version.
