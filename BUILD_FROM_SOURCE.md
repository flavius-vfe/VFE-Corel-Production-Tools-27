# Build a Single Ready-to-Run GMS from Source

This guide creates one standalone file:

```text
VFE-Corel-Production-Tools-27.gms
```

The final GMS contains the main VBA engine and three UserForms. The temporary UI installer should be removed before publishing the final GMS.

## Source files

```text
src/VFE_CorelProductionTools27.bas
src/VFE_UI_Installer.bas
```

The installer source in this repository exposes only:

```text
VFE_InstallVisualInterface
```

## 1. Create a standalone GMS project

Do not build the product inside `GlobalMacros.gms`.

1. Open CorelDRAW.
2. Open **Tools → Scripts → Scripts**.
3. Select **Visual Basic for Applications**.
4. Click **New → New Macro Project**.
5. Save the new project as:

```text
VFE-Corel-Production-Tools-27.gms
```

The GMS file exists as soon as the Save dialog completes. Choose either a development folder or the user CorelDRAW GMS folder.

## 2. Open the VBA editor

1. Press **Alt+F11**.
2. Press **Ctrl+R** if Project Explorer is hidden.
3. Select the project associated with `VFE-Corel-Production-Tools-27.gms`.
4. Do not select `GlobalMacros.gms`.

A new project normally contains only `ThisMacroStorage`.

## 3. Import the engine

With the standalone VFE project selected:

1. Choose **File → Import File**.
2. Import `src/VFE_CorelProductionTools27.bas`.

The project should show:

```text
Modules
└── VFE_CorelProductionTools27
```

## 4. Import the temporary UI installer

Select the same standalone project again and import:

```text
src/VFE_UI_Installer.bas
```

The project should show:

```text
Modules
├── VFE_CorelProductionTools27
└── VFE_UI_Installer
```

## 5. Permit programmatic form creation if required

The installer creates UserForms through the VBA project object model. If CorelDRAW blocks this:

1. Open CorelDRAW's VBA/security options.
2. Enable **Trust access to Visual Basic project** only for the controlled build session.
3. Restart CorelDRAW if required.
4. Run the installer.
5. Disable the setting again after the forms are generated and saved.

Do not leave elevated macro trust enabled unnecessarily.

## 6. Generate the three UserForms

1. Open `VFE_UI_Installer` in the VBA editor.
2. Put the cursor inside:

```vb
Public Sub VFE_InstallVisualInterface()
```

3. Press **F5**.
4. Confirm the target project name when prompted.

The installer creates:

```text
Forms
├── frmVFELauncher
├── frmVFEDimensionTool
└── frmVFEExportWorkflow
```

Save immediately with **Ctrl+S**.

## 7. Test before removing the installer

Run:

```text
VFE_CorelProductionTools27.VFE_CorelProductionTools27
```

Test on disposable documents:

- launcher and all four buttons;
- dimensions for one object and multiple objects;
- text conversion on the current page;
- text conversion in PowerClips;
- text conversion across all pages;
- CDR and PDF exports;
- all page-sizing modes;
- saved custom profiles;
- cancellation and error handling.

## 8. Remove the temporary installer

After successful testing:

1. Right-click `VFE_UI_Installer` in Project Explorer.
2. Choose **Remove VFE_UI_Installer**.
3. Choose **No** when VBA asks whether to export it.

The final structure should be:

```text
VFE_Corel_Production_Tools_27
├── CorelDRAW Objects
│   └── ThisMacroStorage
├── Modules
│   └── VFE_CorelProductionTools27
└── Forms
    ├── frmVFELauncher
    ├── frmVFEDimensionTool
    └── frmVFEExportWorkflow
```

Removing the installer module does not remove the forms it created.

## 9. Compile and save

1. Choose **Debug → Compile** for the VFE project.
2. Resolve every compile error.
3. Press **Ctrl+S**.
4. Close the VBA editor.
5. Exit CorelDRAW completely so the GMS is written to disk.

## 10. Test the standalone GMS

1. Make a copy of the finished GMS using Windows File Explorer while CorelDRAW is closed, or use **Copy To** from the Scripts docker.
2. Temporarily move the development GMS out of the auto-load folder.
3. Open CorelDRAW.
4. Use **Tools → Scripts → Scripts → Load** to load the release copy.
5. Run the launcher and complete a minimal test.

## 11. Publish

Recommended GitHub release files:

```text
VFE-Corel-Production-Tools-27.gms
VFE-Corel-Production-Tools-27-v0.4.0.zip
SHA256SUMS.txt
```

Keep the source modules in the repository, but end users should install the GMS from the release package.
