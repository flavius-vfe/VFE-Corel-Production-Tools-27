# Installation

## Requirements

- Windows
- CorelDRAW version 27 with VBA support installed
- Permission to load a local GMS macro project

The normal end-user installation uses only this file:

```text
VFE-Corel-Production-Tools-27.gms
```

End users do **not** need to import `.bas` files or run the visual-interface installer.

## Method A — Load for the current CorelDRAW session

1. Save `VFE-Corel-Production-Tools-27.gms` to a trusted local folder.
2. Open CorelDRAW.
3. Open **Tools → Scripts → Scripts**. The shortcut is commonly **Alt+Shift+F11**.
4. Select **Visual Basic for Applications** in the Scripts docker.
5. Click **Load** or right-click **Visual Basic for Applications** and choose **Load macro project**.
6. Select `VFE-Corel-Production-Tools-27.gms`.
7. Expand the loaded project and run:

```text
VFE_CorelProductionTools27.VFE_CorelProductionTools27
```

A manually loaded GMS normally remains loaded only until CorelDRAW closes.

## Method B — Install for automatic loading

1. Close CorelDRAW.
2. Locate the current user's CorelDRAW GMS folder. The general pattern is:

```text
%APPDATA%\Corel\<CorelDRAW suite folder>\Draw\GMS\
```

3. If the `GMS` folder does not exist, create it under the correct `Draw` folder.
4. Copy `VFE-Corel-Production-Tools-27.gms` into that folder.
5. Start CorelDRAW.
6. Open **Tools → Scripts → Scripts** and locate the VFE project.
7. Run:

```text
VFE_CorelProductionTools27.VFE_CorelProductionTools27
```

Do not place the GMS into `GlobalMacros.gms`, a CDR document, or an unrelated macro project.

## Macro-security prompt

CorelDRAW may ask whether macros should be enabled. Only enable the project when the GMS was obtained from a trusted VFE release and its checksum matches the release manifest.

Do not lower macro security globally just to install the tool. Prefer loading the trusted GMS explicitly or placing it in the trusted user GMS folder.

## Verify the installation

The launcher should show four actions:

1. Dimension Tool
2. Export Workflow
3. Convert Text to Curves — Current Page
4. Convert Text to Curves — All Pages

The interface should display:

```text
Copyright (c) 2026 VFE Flavius. All rights reserved.
```

## Assign a keyboard shortcut

1. Open **Tools → Scripts → Scripts**.
2. Right-click `VFE_CorelProductionTools27` or the launcher macro.
3. Choose **Assign Keyboard Shortcut**.
4. Select a key combination that is not already assigned.
5. Save the workspace customization.

## Update

1. Close all VFE forms and CorelDRAW.
2. Back up the existing GMS and any important custom profile information.
3. Replace the old GMS with the new release file.
4. Restart CorelDRAW.
5. Test the launcher and one disposable export job.

## Uninstall

### Manually loaded project

Open the Scripts docker, right-click the VFE macro project, and choose **Unload macro project**.

### Automatically loaded project

1. Close CorelDRAW.
2. Remove `VFE-Corel-Production-Tools-27.gms` from the user `Draw\GMS` folder.
3. Restart CorelDRAW.

Removing the GMS does not delete CDR documents or files previously exported by the tool.
