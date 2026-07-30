# Publish This Repository on GitHub

## Upload repository contents

1. Create the GitHub repository.
2. Extract the repository ZIP.
3. Upload the contents of the repository folder—not the outer ZIP—to the repository root.
4. Commit with a message such as:

```text
Initial repository release: VFE Corel Production Tools 27 v0.4.0
```

## Create a GitHub release

1. Open **Releases → Draft a new release**.
2. Create tag `v0.4.0`.
3. Use title `VFE Corel Production Tools 27 v0.4.0`.
4. Attach:

```text
dist/VFE-Corel-Production-Tools-27.gms
release-assets/VFE-Corel-Production-Tools-27-v0.4.0.zip
checksums/SHA256SUMS.txt
```

5. Copy the relevant section from `CHANGELOG.md` into the release notes.
6. State that users should back up documents and test on disposable files.

## Source-visible proprietary notice

Do not choose an open-source licence in GitHub's licence selector. The repository already contains the custom `LICENSE` and `LICENSE.md` files. Repository visibility does not grant modification, redistribution, or resale rights.

## Before publishing

The supplied GMS was not executed in this packaging environment. Complete `docs/RELEASE_CHECKLIST.md` in CorelDRAW before marking the release as production-ready.
