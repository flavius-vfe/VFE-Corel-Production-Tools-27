Attribute VB_Name = "VFE_UI_Installer"
Option Explicit

' SPDX-License-Identifier: LicenseRef-VFE-Proprietary-1.0
' Licensed under the VFE Proprietary Software License v1.0.
' See LICENSE.md in the distribution package.

' VFE Corel Production Tools 27 v0.4.0 - Visual Interface Installer
' Copyright (c) 2026 VFE Flavius. All rights reserved.
' Author and copyright holder: VFE Flavius
'
' Run VFE_InstallVisualInterface after importing this module and
' VFE_CorelProductionTools27.bas into the target GMS project.

' Backward-compatible installer entry point.
Public Sub CPT27_InstallVisualInterface()
    VFE_InstallVisualInterface
End Sub

Public Sub VFE_InstallVisualInterface()
    Dim vbProj As Object
    Dim errText As String

    On Error GoTo InstallFail
    Set vbProj = Application.VBE.ActiveVBProject

    If vbProj Is Nothing Then Err.Raise vbObjectError + 2730, , "No active VBA project."

    If MsgBox("Install or replace the VFE Corel Production Tools 27 visual interface in project:" & vbCrLf & _
              vbProj.Name & "?", vbYesNo + vbQuestion, "VFE CPT27 v0.4.0 Interface Installer") <> vbYes Then Exit Sub

    CPT27_CreateLauncher vbProj
    CPT27_CreateDimensionForm vbProj
    CPT27_CreateExportForm vbProj

    MsgBox "The visual interface was installed successfully." & vbCrLf & vbCrLf & _
           "Run VFE_CorelProductionTools27.VFE_CorelProductionTools27 to open it.", _
           vbInformation, "VFE CPT27 v0.4.0 Interface Installer"
    Exit Sub

InstallFail:
    errText = Err.Description
    MsgBox "The interface could not be installed:" & vbCrLf & errText & vbCrLf & vbCrLf & _
           "Make sure the target GMS project is selected and programmatic access to the VBA project is allowed.", _
           vbCritical, "VFE CPT27 v0.4.0 Interface Installer"
End Sub

Private Sub CPT27_CreateLauncher(ByVal vbProj As Object)
    Dim comp As Object
    Dim designer As Object
    Dim ctl As Object

    CPT27_RemoveComponent vbProj, "frmCPTLauncher"
    CPT27_RemoveComponent vbProj, "frmVFELauncher"
    Set comp = vbProj.VBComponents.Add(3)
    comp.Name = "frmVFELauncher"
    Set designer = comp.designer

    CPT27_SetComponentProperty comp, "Caption", "VFE Corel Production Tools 27"
    CPT27_SetComponentProperty comp, "Width", 390#
    CPT27_SetComponentProperty comp, "Height", 315#
    CPT27_SetComponentProperty comp, "StartUpPosition", 1

    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblTitle", "VFE COREL PRODUCTION TOOLS 27", 18#, 12#, 342#, 18#)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblInfo", "Prepare dimensions, convert source text when required, or run a saved/custom export workflow.", 18#, 33#, 342#, 27#)
    CPT27_SetControlProperty ctl, "WordWrap", True
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdDimensions", "1.  Dimension Tool", 42#, 69#, 288#, 33#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdExport", "2.  Export Workflow", 42#, 108#, 288#, 33#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdCurvesCurrent", "3.  Convert Text to Curves - Current Page", 42#, 147#, 288#, 33#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdCurvesAll", "4.  Convert Text to Curves - All Pages", 42#, 186#, 288#, 33#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdClose", "Close", 282#, 231#, 48#, 21#)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblCopyright", "Copyright (c) 2026 VFE Flavius. All rights reserved.", 18#, 267#, 342#, 13.5)
    CPT27_SetControlProperty ctl, "TextAlign", 2

    CPT27_SetCode comp, CPT27_frmVFELauncher_Code()
End Sub

Private Sub CPT27_CreateDimensionForm(ByVal vbProj As Object)
    Dim comp As Object
    Dim designer As Object
    Dim ctl As Object

    CPT27_RemoveComponent vbProj, "frmDimensionTool"
    CPT27_RemoveComponent vbProj, "frmVFEDimensionTool"
    Set comp = vbProj.VBComponents.Add(3)
    comp.Name = "frmVFEDimensionTool"
    Set designer = comp.designer

    CPT27_SetComponentProperty comp, "Caption", "VFE Corel Production Tools 27 - Dimension Tool"
    CPT27_SetComponentProperty comp, "Width", 462#
    CPT27_SetComponentProperty comp, "Height", 399#
    CPT27_SetComponentProperty comp, "StartUpPosition", 1

    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblHeader", "DIMENSION TOOL", 15#, 9#, 414#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblInstructions", "Select objects in CorelDRAW, choose how they should be measured, then add editable dimensions to the DIMENSIONS layer.", 15#, 27#, 414#, 25.5)
    CPT27_SetControlProperty ctl, "WordWrap", True
    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraTarget", "Measure", 15#, 57#, 201#, 63#)
    Set ctl = CPT27_AddControl(designer.Controls("fraTarget"), "Forms.OptionButton.1", "optEachObject", "Every selected object separately", 12#, 18#, 171#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraTarget"), "Forms.OptionButton.1", "optCompleteSelection", "Entire selection as one object", 12#, 39#, 171#, 15#)
    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraAxes", "Create", 225#, 57#, 204#, 63#)
    Set ctl = CPT27_AddControl(designer.Controls("fraAxes"), "Forms.CheckBox.1", "chkWidth", "Width dimension", 12#, 18#, 165#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraAxes"), "Forms.CheckBox.1", "chkHeight", "Height dimension", 12#, 39#, 165#, 15#)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblUnit", "Display unit", 21#, 138#, 75#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.ComboBox.1", "cboUnit", "", 99#, 135#, 66#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblDecimals", "Decimals", 21#, 162#, 75#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.ComboBox.1", "cboDecimals", "", 99#, 159#, 66#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblFont", "Font name", 21#, 186#, 75#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.TextBox.1", "txtFont", "", 99#, 183#, 117#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblFontSize", "Font size (pt)", 234#, 138#, 87#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.TextBox.1", "txtFontSize", "", 324#, 135#, 66#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblSpacing", "Line distance (mm)", 234#, 162#, 87#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.TextBox.1", "txtSpacing", "", 324#, 159#, 66#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblColor", "Colour", 234#, 186#, 87#, 13.5)
    Set ctl = CPT27_AddControl(designer, "Forms.ComboBox.1", "cboColor", "", 324#, 183#, 66#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer, "Forms.CheckBox.1", "chkRemoveOld", "Remove existing objects from the DIMENSIONS layer first", 21#, 216#, 300#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.CheckBox.1", "chkGroupCreated", "Group the newly created dimension objects", 21#, 237#, 300#, 16.5)
    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraSelection", "Current selection", 15#, 261#, 414#, 45#)
    Set ctl = CPT27_AddControl(designer.Controls("fraSelection"), "Forms.Label.1", "lblSelection", "No objects selected", 12#, 18#, 297#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraSelection"), "Forms.CommandButton.1", "cmdRefresh", "Refresh", 324#, 15#, 66#, 18#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdAdd", "Add Dimensions", 258#, 312#, 102#, 24#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdClose", "Close", 369#, 312#, 60#, 24#)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblCopyright", "Copyright (c) 2026 VFE Flavius. All rights reserved.", 15#, 348#, 414#, 13.5)
    CPT27_SetControlProperty ctl, "TextAlign", 2

    CPT27_SetCode comp, CPT27_frmVFEDimensionTool_Code()
End Sub

Private Sub CPT27_CreateExportForm(ByVal vbProj As Object)
    Dim comp As Object
    Dim designer As Object
    Dim ctl As Object

    CPT27_RemoveComponent vbProj, "frmExportWorkflow"
    CPT27_RemoveComponent vbProj, "frmVFEExportWorkflow"
    Set comp = vbProj.VBComponents.Add(3)
    comp.Name = "frmVFEExportWorkflow"
    Set designer = comp.designer

    CPT27_SetComponentProperty comp, "Caption", "VFE Corel Production Tools 27 - Export Workflow"
    CPT27_SetComponentProperty comp, "Width", 672#
    CPT27_SetComponentProperty comp, "Height", 555#
    CPT27_SetComponentProperty comp, "StartUpPosition", 1

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraProfile", "Workflow profile", 12#, 9#, 636#, 45#)
    Set ctl = CPT27_AddControl(designer.Controls("fraProfile"), "Forms.ComboBox.1", "cboProfile", "", 9#, 16.5, 234#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer.Controls("fraProfile"), "Forms.CommandButton.1", "cmdNewProfile", "New", 255#, 15#, 51#, 18#)
    Set ctl = CPT27_AddControl(designer.Controls("fraProfile"), "Forms.CommandButton.1", "cmdSaveProfile", "Save As", 312#, 15#, 60#, 18#)
    Set ctl = CPT27_AddControl(designer.Controls("fraProfile"), "Forms.CommandButton.1", "cmdDeleteProfile", "Delete", 378#, 15#, 51#, 18#)
    Set ctl = CPT27_AddControl(designer.Controls("fraProfile"), "Forms.Label.1", "lblProfileHelp", "Profiles contain export settings only. Dimensions are prepared separately.", 441#, 15#, 174#, 18#)
    CPT27_SetControlProperty ctl, "WordWrap", True

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraScopeName", "Job and naming", 12#, 60#, 204#, 105#)
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.Label.1", "lblScope", "Process", 9#, 18#, 57#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.ComboBox.1", "cboScope", "", 69#, 15#, 120#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.Label.1", "lblNameMode", "Filename source", 9#, 42#, 66#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.ComboBox.1", "cboNameMode", "", 78#, 39#, 111#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.Label.1", "lblCustomName", "Custom prefix", 9#, 66#, 66#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.TextBox.1", "txtCustomName", "", 78#, 63#, 111#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraScopeName"), "Forms.CheckBox.1", "chkRenamePages", "Rename source pages from resolved names", 9#, 84#, 177#, 16.5)

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraFormats", "Output formats", 225#, 60#, 195#, 105#)
    Set ctl = CPT27_AddControl(designer.Controls("fraFormats"), "Forms.CheckBox.1", "chkCurrentCDR", "Current-version CDR", 9#, 15#, 168#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraFormats"), "Forms.CheckBox.1", "chkV15CDR", "Version 15 CDR", 9#, 33#, 168#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraFormats"), "Forms.CheckBox.1", "chkCurvesCDR", "Curved current-version CDR", 9#, 51#, 168#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraFormats"), "Forms.CheckBox.1", "chkCurvesV15", "Curved version 15 CDR", 9#, 69#, 168#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraFormats"), "Forms.CheckBox.1", "chkPDF", "Curved PDF", 9#, 87#, 168#, 15#)

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraSummary", "Workflow summary", 429#, 60#, 219#, 246#)
    Set ctl = CPT27_AddControl(designer.Controls("fraSummary"), "Forms.TextBox.1", "txtSummary", "", 9#, 18#, 201#, 219#)
    CPT27_SetControlProperty ctl, "Locked", True
    CPT27_SetControlProperty ctl, "MultiLine", True
    CPT27_SetControlProperty ctl, "ScrollBars", 2

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraCDR", "CDR page handling", 12#, 174#, 204#, 132#)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.ComboBox.1", "cboCDRPageMode", "", 9#, 18#, 180#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.Label.1", "lblCDRMargin", "Margin (mm)", 9#, 45#, 72#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.TextBox.1", "txtCDRMargin", "", 84#, 42#, 48#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.Label.1", "lblCDRWidth", "Width", 9#, 69#, 36#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.TextBox.1", "txtCDRWidth", "", 45#, 66#, 45#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.Label.1", "lblCDRHeight", "Height", 99#, 69#, 42#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.TextBox.1", "txtCDRHeight", "", 144#, 66#, 45#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.Label.1", "lblCDRUnits", "Custom size in mm", 9#, 87#, 93#, 12#)
    Set ctl = CPT27_AddControl(designer.Controls("fraCDR"), "Forms.CheckBox.1", "chkIncludeOutlines", "Include outline width in fitted bounds", 9#, 105#, 180#, 15#)

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraPDF", "PDF page handling", 225#, 174#, 195#, 132#)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.ComboBox.1", "cboPDFPageMode", "", 9#, 18#, 177#, 16.5)
    CPT27_SetControlProperty ctl, "Style", 2
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.Label.1", "lblPDFMargin", "Margin (mm)", 9#, 45#, 72#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.TextBox.1", "txtPDFMargin", "", 84#, 42#, 48#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.Label.1", "lblPDFWidth", "Width", 9#, 69#, 36#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.TextBox.1", "txtPDFWidth", "", 45#, 66#, 45#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.Label.1", "lblPDFHeight", "Height", 99#, 69#, 42#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.TextBox.1", "txtPDFHeight", "", 144#, 66#, 42#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.Label.1", "lblPDFUnits", "Custom size in mm", 9#, 87#, 93#, 12#)
    Set ctl = CPT27_AddControl(designer.Controls("fraPDF"), "Forms.CheckBox.1", "chkIncludeDimensions", "Include DIMENSIONS layer in fitted page size", 9#, 102#, 177#, 27#)
    CPT27_SetControlProperty ctl, "WordWrap", True

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraOutput", "Output and safety", 12#, 315#, 408#, 114#)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.Label.1", "lblFolder", "Output folder", 9#, 18#, 66#, 13.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.TextBox.1", "txtOutputFolder", "", 75#, 15#, 252#, 16.5)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.CommandButton.1", "cmdBrowse", "Browse", 333#, 14.25, 57#, 18#)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.CheckBox.1", "chkEnsureUnique", "Add _2, _3, etc. when files already exist", 9#, 45#, 186#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.CheckBox.1", "chkAutoScale", "Automatically scale oversized PDF pages to 1:10", 207#, 45#, 183#, 15#)
    Set ctl = CPT27_AddControl(designer.Controls("fraOutput"), "Forms.Label.1", "lblPDFNote", "Preserve mode extracts an exact duplicate of the source page. Fit modes group objects for sizing. Custom modes use the width and height entered above.", 9#, 66#, 381#, 36#)
    CPT27_SetControlProperty ctl, "WordWrap", True

    Set ctl = CPT27_AddControl(designer, "Forms.Frame.1", "fraProgress", "Progress", 429#, 315#, 219#, 108#)
    CPT27_SetControlProperty ctl, "Visible", False
    Set ctl = CPT27_AddControl(designer.Controls("fraProgress"), "Forms.Label.1", "lblProgressText", "Ready", 9#, 15#, 201#, 36#)
    CPT27_SetControlProperty ctl, "WordWrap", True
    Set ctl = CPT27_AddControl(designer.Controls("fraProgress"), "Forms.Label.1", "lblProgressFill", "", 9#, 54#, 3#, 12#)
    CPT27_SetControlProperty ctl, "BackColor", &H8000000D
    CPT27_SetControlProperty ctl, "BorderStyle", 1
    Set ctl = CPT27_AddControl(designer.Controls("fraProgress"), "Forms.CommandButton.1", "cmdCancel", "Cancel", 153#, 78#, 57#, 18#)
    CPT27_SetControlProperty ctl, "Enabled", False

    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdPreflight", "Preflight", 360#, 444#, 75#, 24#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdRun", "Run Workflow", 444#, 444#, 102#, 24#)
    Set ctl = CPT27_AddControl(designer, "Forms.CommandButton.1", "cmdClose", "Close", 555#, 444#, 75#, 24#)
    Set ctl = CPT27_AddControl(designer, "Forms.Label.1", "lblCopyright", "Copyright (c) 2026 VFE Flavius. All rights reserved.", 12#, 486#, 636#, 13.5)
    CPT27_SetControlProperty ctl, "TextAlign", 2

    CPT27_SetCode comp, CPT27_frmVFEExportWorkflow_Code()
End Sub

Private Sub CPT27_RemoveComponent(ByVal vbProj As Object, ByVal componentName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = vbProj.VBComponents(componentName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    On Error GoTo 0
End Sub

Private Function CPT27_AddControl( _
    ByVal parent As Object, _
    ByVal progId As String, _
    ByVal controlName As String, _
    ByVal captionText As String, _
    ByVal leftValue As Double, _
    ByVal topValue As Double, _
    ByVal widthValue As Double, _
    ByVal heightValue As Double) As Object

    Dim ctl As Object
    Set ctl = parent.Controls.Add(progId, controlName, True)
    On Error Resume Next
    ctl.Caption = captionText
    ctl.Left = leftValue
    ctl.Top = topValue
    ctl.Width = widthValue
    ctl.Height = heightValue
    On Error GoTo 0
    Set CPT27_AddControl = ctl
End Function

Private Sub CPT27_SetControlProperty(ByVal ctl As Object, ByVal propertyName As String, ByVal propertyValue As Variant)
    On Error Resume Next
    CallByName ctl, propertyName, VbLet, propertyValue
    On Error GoTo 0
End Sub

Private Sub CPT27_SetComponentProperty(ByVal comp As Object, ByVal propertyName As String, ByVal propertyValue As Variant)
    On Error Resume Next
    comp.Properties(propertyName).value = propertyValue
    On Error GoTo 0
End Sub

Private Sub CPT27_SetCode(ByVal comp As Object, ByVal codeText As String)
    Dim cm As Object
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString codeText
End Sub

Private Function CPT27_frmVFELauncher_Code() As String
    CPT27_frmVFELauncher_Code = CPT27_frmVFELauncher_CodePart1()
End Function

Private Function CPT27_frmVFELauncher_CodePart1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDimensions_Click()" & vbCrLf
    s = s & "    CPT_ShowDimensionTool" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdExport_Click()" & vbCrLf
    s = s & "    CPT_ShowExportWorkflow" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCurvesCurrent_Click()" & vbCrLf
    s = s & "    CPT_ConvertTextCurrentPage" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCurvesAll_Click()" & vbCrLf
    s = s & "    CPT_ConvertTextAllPages" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CPT27_frmVFELauncher_CodePart1 = s
End Function

Private Function CPT27_frmVFEDimensionTool_Code() As String
    CPT27_frmVFEDimensionTool_Code = CPT27_frmVFEDimensionTool_CodePart1()
End Function

Private Function CPT27_frmVFEDimensionTool_CodePart1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    optEachObject.Value = True" & vbCrLf
    s = s & "    chkWidth.Value = True" & vbCrLf
    s = s & "    chkHeight.Value = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboUnit.AddItem ""mm""" & vbCrLf
    s = s & "    cboUnit.AddItem ""cm""" & vbCrLf
    s = s & "    cboUnit.AddItem ""in""" & vbCrLf
    s = s & "    cboUnit.ListIndex = 1" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboDecimals.AddItem ""0""" & vbCrLf
    s = s & "    cboDecimals.AddItem ""1""" & vbCrLf
    s = s & "    cboDecimals.AddItem ""2""" & vbCrLf
    s = s & "    cboDecimals.AddItem ""3""" & vbCrLf
    s = s & "    cboDecimals.AddItem ""4""" & vbCrLf
    s = s & "    cboDecimals.ListIndex = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboColor.AddItem ""Blue""" & vbCrLf
    s = s & "    cboColor.AddItem ""Red""" & vbCrLf
    s = s & "    cboColor.AddItem ""Green""" & vbCrLf
    s = s & "    cboColor.AddItem ""Black""" & vbCrLf
    s = s & "    cboColor.ListIndex = 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtFont.Text = ""Arial""" & vbCrLf
    s = s & "    txtFontSize.Text = ""10""" & vbCrLf
    s = s & "    txtSpacing.Text = ""5""" & vbCrLf
    s = s & "    chkRemoveOld.Value = False" & vbCrLf
    s = s & "    chkGroupCreated.Value = True" & vbCrLf
    s = s & "    RefreshSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRefresh_Click()" & vbCrLf
    s = s & "    RefreshSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdAdd_Click()" & vbCrLf
    s = s & "    Dim targetMode As Long" & vbCrLf
    s = s & "    Dim fontSize As Double" & vbCrLf
    s = s & "    Dim spacing As Double" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TryReadNumber(txtFontSize.Text, fontSize) Then" & vbCrLf
    s = s & "        MsgBox ""Enter a valid font size."", vbExclamation, ""Dimension Tool""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not TryReadNumber(txtSpacing.Text, spacing) Then" & vbCrLf
    s = s & "        MsgBox ""Enter a valid line distance."", vbExclamation, ""Dimension Tool""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If optCompleteSelection.Value Then targetMode = 1 Else targetMode = 2" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    CPT_UI_AddDimensions targetMode, chkWidth.Value, chkHeight.Value, _" & vbCrLf
    s = s & "        cboUnit.Value, CLng(cboDecimals.Value), txtFont.Text, fontSize, spacing, _" & vbCrLf
    s = s & "        chkRemoveOld.Value, chkGroupCreated.Value, cboColor.Value" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    RefreshSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    Unload Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Activate()" & vbCrLf
    s = s & "    RefreshSelection" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub RefreshSelection()" & vbCrLf
    s = s & "    lblSelection.Caption = CPT_UI_GetSelectionSummary()" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TryReadNumber(ByVal valueText As String, ByRef result As Double) As Boolean" & vbCrLf
    s = s & "    Dim normalized As String" & vbCrLf
    s = s & "    normalized = Replace(Trim$(valueText), "","", ""."")" & vbCrLf
    s = s & "    If Len(normalized) = 0 Then Exit Function" & vbCrLf
    s = s & "    If normalized Like ""*[!0-9.-]*"" Then Exit Function" & vbCrLf
    s = s & "    On Error GoTo InvalidValue" & vbCrLf
    s = s & "    result = Val(normalized)" & vbCrLf
    s = s & "    TryReadNumber = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "InvalidValue:" & vbCrLf
    s = s & "    TryReadNumber = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    CPT27_frmVFEDimensionTool_CodePart1 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_Code() As String
    CPT27_frmVFEExportWorkflow_Code = CPT27_frmVFEExportWorkflow_CodePart1() & CPT27_frmVFEExportWorkflow_CodePart2() & CPT27_frmVFEExportWorkflow_CodePart3() & CPT27_frmVFEExportWorkflow_CodePart4() & CPT27_frmVFEExportWorkflow_CodePart5() & CPT27_frmVFEExportWorkflow_CodePart6() & CPT27_frmVFEExportWorkflow_CodePart7()
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart1() As String
    Dim s As String
    s = s & "Option Explicit" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Const UI_APP As String = ""CorelProductionTools27_UI""" & vbCrLf
    s = s & "Private Const INDEX_SECTION As String = ""Profiles""" & vbCrLf
    s = s & "Private Const INDEX_KEY As String = ""Names""" & vbCrLf
    s = s & "Private mLoading As Boolean" & vbCrLf
    s = s & "Private mRunning As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "    CPT_UI_RegisterExportForm Me" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboScope.AddItem ""Each page as separate files""" & vbCrLf
    s = s & "    cboScope.AddItem ""All pages in one document""" & vbCrLf
    s = s & "    cboScope.AddItem ""Current page only""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboNameMode.AddItem ""Object named EXPORT_NAME""" & vbCrLf
    s = s & "    cboNameMode.AddItem ""Document filename""" & vbCrLf
    s = s & "    cboNameMode.AddItem ""Page name""" & vbCrLf
    s = s & "    cboNameMode.AddItem ""Custom prefix""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboCDRPageMode.AddItem ""Preserve source page exactly""" & vbCrLf
    s = s & "    cboCDRPageMode.AddItem ""Fit grouped objects - no margin""" & vbCrLf
    s = s & "    cboCDRPageMode.AddItem ""Fit grouped objects - custom margin""" & vbCrLf
    s = s & "    cboCDRPageMode.AddItem ""Custom page size - keep positions""" & vbCrLf
    s = s & "    cboCDRPageMode.AddItem ""Custom page size - center objects""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboPDFPageMode.AddItem ""Fit grouped objects - no margin""" & vbCrLf
    s = s & "    cboPDFPageMode.AddItem ""Fit grouped objects - custom margin""" & vbCrLf
    s = s & "    cboPDFPageMode.AddItem ""Preserve source page exactly""" & vbCrLf
    s = s & "    cboPDFPageMode.AddItem ""Custom page size - keep positions""" & vbCrLf
    s = s & "    cboPDFPageMode.AddItem ""Custom page size - center objects""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    LoadProfileList" & vbCrLf
    s = s & "    ApplyBuiltInProfile ""Production Export""" & vbCrLf
    s = s & "    SelectProfileCaption ""Built-in: Production Export""" & vbCrLf
    s = s & "    If Len(CPT_UI_DefaultOutputFolder()) > 0 Then txtOutputFolder.Text = CPT_UI_DefaultOutputFolder()" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    fraProgress.Visible = False" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "    UpdateControlState" & vbCrLf
    s = s & "    UpdateSummary" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_Terminate()" & vbCrLf
    s = s & "    CPT_UI_UnregisterExportForm Me" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadProfileList()" & vbCrLf
    s = s & "    Dim namesText As String" & vbCrLf
    s = s & "    Dim names() As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboProfile.Clear" & vbCrLf
    s = s & "    cboProfile.AddItem ""Built-in: Production Export""" & vbCrLf
    s = s & "    cboProfile.AddItem ""Built-in: Customer Preview""" & vbCrLf
    s = s & "    cboProfile.AddItem ""Built-in: All Pages Archive""" & vbCrLf
    s = s & "    cboProfile.AddItem ""Built-in: Curved PDF Per Page""" & vbCrLf
    s = s & "    cboProfile.AddItem ""Built-in: 3 Formats Per Page""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    namesText = GetSetting(UI_APP, INDEX_SECTION, INDEX_KEY, """")" & vbCrLf
    s = s & "    If Len(namesText) > 0 Then" & vbCrLf
    s = s & "        names = Split(namesText, ""|"")" & vbCrLf
    s = s & "        For i = LBound(names) To UBound(names)" & vbCrLf
    s = s & "            If Len(Trim$(names(i))) > 0 Then cboProfile.AddItem ""Custom: "" & names(i)" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SelectProfileCaption(ByVal captionText As String)" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "    For i = 0 To cboProfile.ListCount - 1" & vbCrLf
    s = s & "        If StrComp(cboProfile.List(i), captionText, vbTextCompare) = 0 Then" & vbCrLf
    s = s & "            cboProfile.ListIndex = i" & vbCrLf
    s = s & "            Exit Sub" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboProfile_Change()" & vbCrLf
    s = s & "    Dim captionText As String" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "    captionText = cboProfile.Value" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "    If Left$(captionText, 10) = ""Built-in: "" Then" & vbCrLf
    s = s & "        ApplyBuiltInProfile Mid$(captionText, 11)" & vbCrLf
    s = s & "    ElseIf Left$(captionText, 8) = ""Custom: "" Then" & vbCrLf
    s = s & "        LoadCustomProfile Mid$(captionText, 9)" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "    UpdateControlState" & vbCrLf
    s = s & "    UpdateSummary" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub ApplyBuiltInProfile(ByVal profileName As String)" & vbCrLf
    s = s & "    SetBasicDefaults" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case profileName" & vbCrLf
    s = s & "        Case ""Customer Preview""" & vbCrLf
    s = s & "            cboScope.ListIndex = 2" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart1 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart2() As String
    Dim s As String
    s = s & "            cboNameMode.ListIndex = 2" & vbCrLf
    s = s & "            chkCurrentCDR.Value = False" & vbCrLf
    s = s & "            chkV15CDR.Value = False" & vbCrLf
    s = s & "            chkCurvesCDR.Value = False" & vbCrLf
    s = s & "            chkCurvesV15.Value = False" & vbCrLf
    s = s & "            chkPDF.Value = True" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 0" & vbCrLf
    s = s & "            cboPDFPageMode.ListIndex = 0" & vbCrLf
    s = s & "            chkRenamePages.Value = False" & vbCrLf
    s = s & "        Case ""All Pages Archive""" & vbCrLf
    s = s & "            cboScope.ListIndex = 1" & vbCrLf
    s = s & "            cboNameMode.ListIndex = 1" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 0" & vbCrLf
    s = s & "            cboPDFPageMode.ListIndex = 2" & vbCrLf
    s = s & "            chkRenamePages.Value = False" & vbCrLf
    s = s & "        Case ""Curved PDF Per Page""" & vbCrLf
    s = s & "            cboScope.ListIndex = 0" & vbCrLf
    s = s & "            cboNameMode.ListIndex = 0" & vbCrLf
    s = s & "            chkCurrentCDR.Value = False" & vbCrLf
    s = s & "            chkV15CDR.Value = False" & vbCrLf
    s = s & "            chkCurvesCDR.Value = False" & vbCrLf
    s = s & "            chkCurvesV15.Value = False" & vbCrLf
    s = s & "            chkPDF.Value = True" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 0" & vbCrLf
    s = s & "            cboPDFPageMode.ListIndex = 1" & vbCrLf
    s = s & "            txtPDFMargin.Text = ""5""" & vbCrLf
    s = s & "        Case ""3 Formats Per Page""" & vbCrLf
    s = s & "            cboScope.ListIndex = 0" & vbCrLf
    s = s & "            cboNameMode.ListIndex = 0" & vbCrLf
    s = s & "            chkCurrentCDR.Value = True" & vbCrLf
    s = s & "            chkV15CDR.Value = False" & vbCrLf
    s = s & "            chkCurvesCDR.Value = True" & vbCrLf
    s = s & "            chkCurvesV15.Value = False" & vbCrLf
    s = s & "            chkPDF.Value = True" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 0" & vbCrLf
    s = s & "            cboPDFPageMode.ListIndex = 2" & vbCrLf
    s = s & "        Case Else" & vbCrLf
    s = s & "            ' Production Export defaults already applied." & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetBasicDefaults()" & vbCrLf
    s = s & "    cboScope.ListIndex = 0" & vbCrLf
    s = s & "    cboNameMode.ListIndex = 0" & vbCrLf
    s = s & "    txtCustomName.Text = """"" & vbCrLf
    s = s & "    chkRenamePages.Value = False" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    chkCurrentCDR.Value = True" & vbCrLf
    s = s & "    chkV15CDR.Value = True" & vbCrLf
    s = s & "    chkCurvesCDR.Value = True" & vbCrLf
    s = s & "    chkCurvesV15.Value = True" & vbCrLf
    s = s & "    chkPDF.Value = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboCDRPageMode.ListIndex = 2" & vbCrLf
    s = s & "    txtCDRMargin.Text = ""5""" & vbCrLf
    s = s & "    txtCDRWidth.Text = ""210""" & vbCrLf
    s = s & "    txtCDRHeight.Text = ""297""" & vbCrLf
    s = s & "    chkIncludeOutlines.Value = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboPDFPageMode.ListIndex = 1" & vbCrLf
    s = s & "    txtPDFMargin.Text = ""5""" & vbCrLf
    s = s & "    txtPDFWidth.Text = ""210""" & vbCrLf
    s = s & "    txtPDFHeight.Text = ""297""" & vbCrLf
    s = s & "    chkIncludeDimensions.Value = True" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    chkEnsureUnique.Value = True" & vbCrLf
    s = s & "    chkAutoScale.Value = True" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdNewProfile_Click()" & vbCrLf
    s = s & "    mLoading = True" & vbCrLf
    s = s & "    SetBasicDefaults" & vbCrLf
    s = s & "    cboProfile.ListIndex = -1" & vbCrLf
    s = s & "    mLoading = False" & vbCrLf
    s = s & "    UpdateControlState" & vbCrLf
    s = s & "    UpdateSummary" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdSaveProfile_Click()" & vbCrLf
    s = s & "    Dim profileName As String" & vbCrLf
    s = s & "    profileName = Trim$(InputBox(""Enter a name for this export profile:"", ""Save Export Profile"", CurrentProfileName()))" & vbCrLf
    s = s & "    If Len(profileName) = 0 Then Exit Sub" & vbCrLf
    s = s & "    profileName = Replace(profileName, ""|"", ""-"")" & vbCrLf
    s = s & "    SaveCustomProfile profileName" & vbCrLf
    s = s & "    LoadProfileList" & vbCrLf
    s = s & "    SelectProfileCaption ""Custom: "" & profileName" & vbCrLf
    s = s & "    MsgBox ""Export profile saved: "" & profileName, vbInformation, ""Export Workflow""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdDeleteProfile_Click()" & vbCrLf
    s = s & "    Dim captionText As String" & vbCrLf
    s = s & "    Dim profileName As String" & vbCrLf
    s = s & "    captionText = cboProfile.Value" & vbCrLf
    s = s & "    If Left$(captionText, 8) <> ""Custom: "" Then" & vbCrLf
    s = s & "        MsgBox ""Only custom profiles can be deleted."", vbInformation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    profileName = Mid$(captionText, 9)" & vbCrLf
    s = s & "    If MsgBox(""Delete profile '"" & profileName & ""'?"", vbYesNo + vbQuestion, ""Export Workflow"") <> vbYes Then Exit Sub" & vbCrLf
    s = s & "    DeleteCustomProfile profileName" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart2 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart3() As String
    Dim s As String
    s = s & "    LoadProfileList" & vbCrLf
    s = s & "    ApplyBuiltInProfile ""Production Export""" & vbCrLf
    s = s & "    SelectProfileCaption ""Built-in: Production Export""" & vbCrLf
    s = s & "    UpdateControlState" & vbCrLf
    s = s & "    UpdateSummary" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SaveCustomProfile(ByVal profileName As String)" & vbCrLf
    s = s & "    Dim section As String" & vbCrLf
    s = s & "    section = ProfileSection(profileName)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""Scope"", CStr(cboScope.ListIndex + 1)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""NameMode"", CStr(cboNameMode.ListIndex + 1)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CustomName"", txtCustomName.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""OutputFolder"", txtOutputFolder.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CurrentCDR"", BoolText(chkCurrentCDR.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""V15CDR"", BoolText(chkV15CDR.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CurvesCDR"", BoolText(chkCurvesCDR.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CurvesV15"", BoolText(chkCurvesV15.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""PDF"", BoolText(chkPDF.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CDRModeV4"", CStr(cboCDRPageMode.ListIndex + 1)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CDRMargin"", txtCDRMargin.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CDRWidth"", txtCDRWidth.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""CDRHeight"", txtCDRHeight.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""PDFMode"", CStr(cboPDFPageMode.ListIndex + 1)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""PDFMargin"", txtPDFMargin.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""PDFWidth"", txtPDFWidth.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""PDFHeight"", txtPDFHeight.Text" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""IncludeOutlines"", BoolText(chkIncludeOutlines.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""IncludeDimensions"", BoolText(chkIncludeDimensions.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""RenamePages"", BoolText(chkRenamePages.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""EnsureUnique"", BoolText(chkEnsureUnique.Value)" & vbCrLf
    s = s & "    SaveSetting UI_APP, section, ""AutoScale"", BoolText(chkAutoScale.Value)" & vbCrLf
    s = s & "    AddProfileName profileName" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub LoadCustomProfile(ByVal profileName As String)" & vbCrLf
    s = s & "    Dim section As String" & vbCrLf
    s = s & "    Dim cdrModeText As String" & vbCrLf
    s = s & "    Dim oldCDRMode As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    section = ProfileSection(profileName)" & vbCrLf
    s = s & "    SetBasicDefaults" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cboScope.ListIndex = CLng(GetSetting(UI_APP, section, ""Scope"", ""1"")) - 1" & vbCrLf
    s = s & "    cboNameMode.ListIndex = CLng(GetSetting(UI_APP, section, ""NameMode"", ""1"")) - 1" & vbCrLf
    s = s & "    txtCustomName.Text = GetSetting(UI_APP, section, ""CustomName"", """")" & vbCrLf
    s = s & "    txtOutputFolder.Text = GetSetting(UI_APP, section, ""OutputFolder"", txtOutputFolder.Text)" & vbCrLf
    s = s & "    chkCurrentCDR.Value = SettingBool(section, ""CurrentCDR"", True)" & vbCrLf
    s = s & "    chkV15CDR.Value = SettingBool(section, ""V15CDR"", True)" & vbCrLf
    s = s & "    chkCurvesCDR.Value = SettingBool(section, ""CurvesCDR"", True)" & vbCrLf
    s = s & "    chkCurvesV15.Value = SettingBool(section, ""CurvesV15"", True)" & vbCrLf
    s = s & "    chkPDF.Value = SettingBool(section, ""PDF"", True)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    cdrModeText = GetSetting(UI_APP, section, ""CDRModeV4"", """")" & vbCrLf
    s = s & "    If Len(cdrModeText) > 0 Then" & vbCrLf
    s = s & "        cboCDRPageMode.ListIndex = CLng(cdrModeText) - 1" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        oldCDRMode = CLng(GetSetting(UI_APP, section, ""CDRMode"", ""0""))" & vbCrLf
    s = s & "        If oldCDRMode = 0 Then" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 0" & vbCrLf
    s = s & "        Else" & vbCrLf
    s = s & "            cboCDRPageMode.ListIndex = 2" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    txtCDRMargin.Text = GetSetting(UI_APP, section, ""CDRMargin"", ""5"")" & vbCrLf
    s = s & "    txtCDRWidth.Text = GetSetting(UI_APP, section, ""CDRWidth"", ""210"")" & vbCrLf
    s = s & "    txtCDRHeight.Text = GetSetting(UI_APP, section, ""CDRHeight"", ""297"")" & vbCrLf
    s = s & "    cboPDFPageMode.ListIndex = CLng(GetSetting(UI_APP, section, ""PDFMode"", ""2"")) - 1" & vbCrLf
    s = s & "    txtPDFMargin.Text = GetSetting(UI_APP, section, ""PDFMargin"", ""5"")" & vbCrLf
    s = s & "    txtPDFWidth.Text = GetSetting(UI_APP, section, ""PDFWidth"", ""210"")" & vbCrLf
    s = s & "    txtPDFHeight.Text = GetSetting(UI_APP, section, ""PDFHeight"", ""297"")" & vbCrLf
    s = s & "    chkIncludeOutlines.Value = SettingBool(section, ""IncludeOutlines"", True)" & vbCrLf
    s = s & "    chkIncludeDimensions.Value = SettingBool(section, ""IncludeDimensions"", True)" & vbCrLf
    s = s & "    chkRenamePages.Value = SettingBool(section, ""RenamePages"", False)" & vbCrLf
    s = s & "    chkEnsureUnique.Value = SettingBool(section, ""EnsureUnique"", True)" & vbCrLf
    s = s & "    chkAutoScale.Value = SettingBool(section, ""AutoScale"", True)" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub DeleteCustomProfile(ByVal profileName As String)" & vbCrLf
    s = s & "    Dim namesText As String" & vbCrLf
    s = s & "    Dim names() As String" & vbCrLf
    s = s & "    Dim newText As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    DeleteSetting UI_APP, ProfileSection(profileName)" & vbCrLf
    s = s & "    On Error GoTo 0" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    namesText = GetSetting(UI_APP, INDEX_SECTION, INDEX_KEY, """")" & vbCrLf
    s = s & "    If Len(namesText) = 0 Then Exit Sub" & vbCrLf
    s = s & "    names = Split(namesText, ""|"")" & vbCrLf
    s = s & "    For i = LBound(names) To UBound(names)" & vbCrLf
    s = s & "        If StrComp(names(i), profileName, vbTextCompare) <> 0 Then" & vbCrLf
    s = s & "            If Len(newText) > 0 Then newText = newText & ""|""" & vbCrLf
    s = s & "            newText = newText & names(i)" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next i" & vbCrLf
    s = s & "    SaveSetting UI_APP, INDEX_SECTION, INDEX_KEY, newText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart3 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart4() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Private Sub AddProfileName(ByVal profileName As String)" & vbCrLf
    s = s & "    Dim namesText As String" & vbCrLf
    s = s & "    Dim names() As String" & vbCrLf
    s = s & "    Dim i As Long" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    namesText = GetSetting(UI_APP, INDEX_SECTION, INDEX_KEY, """")" & vbCrLf
    s = s & "    If Len(namesText) > 0 Then" & vbCrLf
    s = s & "        names = Split(namesText, ""|"")" & vbCrLf
    s = s & "        For i = LBound(names) To UBound(names)" & vbCrLf
    s = s & "            If StrComp(names(i), profileName, vbTextCompare) = 0 Then Exit Sub" & vbCrLf
    s = s & "        Next i" & vbCrLf
    s = s & "        namesText = namesText & ""|"" & profileName" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        namesText = profileName" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    SaveSetting UI_APP, INDEX_SECTION, INDEX_KEY, namesText" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ProfileSection(ByVal profileName As String) As String" & vbCrLf
    s = s & "    Dim value As String" & vbCrLf
    s = s & "    value = profileName" & vbCrLf
    s = s & "    value = Replace(value, ""\"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, ""/"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, "":"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, ""*"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, ""?"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, Chr$(34), ""_"")" & vbCrLf
    s = s & "    value = Replace(value, ""<"", ""_"")" & vbCrLf
    s = s & "    value = Replace(value, "">"", ""_"")" & vbCrLf
    s = s & "    ProfileSection = ""Profile_"" & value" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function BoolText(ByVal value As Boolean) As String" & vbCrLf
    s = s & "    If value Then BoolText = ""1"" Else BoolText = ""0""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SettingBool(ByVal section As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean" & vbCrLf
    s = s & "    Dim value As String" & vbCrLf
    s = s & "    value = GetSetting(UI_APP, section, key, BoolText(defaultValue))" & vbCrLf
    s = s & "    SettingBool = (value = ""1"" Or LCase$(value) = ""true"")" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdBrowse_Click()" & vbCrLf
    s = s & "    Dim folderPath As String" & vbCrLf
    s = s & "    folderPath = CPT_UI_BrowseFolder(""Choose export output folder"")" & vbCrLf
    s = s & "    If Len(folderPath) > 0 Then txtOutputFolder.Text = folderPath" & vbCrLf
    s = s & "    UpdateSummary" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdPreflight_Click()" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "    If Not ValidateNumericSettings() Then Exit Sub" & vbCrLf
    s = s & "    result = CPT_UI_Preflight(cboScope.ListIndex + 1, cboNameMode.ListIndex + 1, _" & vbCrLf
    s = s & "        txtCustomName.Text, txtOutputFolder.Text, chkCurrentCDR.Value, chkV15CDR.Value, _" & vbCrLf
    s = s & "        chkCurvesCDR.Value, chkCurvesV15.Value, chkPDF.Value)" & vbCrLf
    s = s & "    MsgBox result, vbInformation, ""Export Workflow Preflight""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdRun_Click()" & vbCrLf
    s = s & "    Dim result As String" & vbCrLf
    s = s & "    Dim cdrMargin As Double" & vbCrLf
    s = s & "    Dim cdrWidth As Double" & vbCrLf
    s = s & "    Dim cdrHeight As Double" & vbCrLf
    s = s & "    Dim pdfMargin As Double" & vbCrLf
    s = s & "    Dim pdfWidth As Double" & vbCrLf
    s = s & "    Dim pdfHeight As Double" & vbCrLf
    s = s & "    On Error GoTo RunFail" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not ValidateNumericSettings() Then Exit Sub" & vbCrLf
    s = s & "    result = CPT_UI_Preflight(cboScope.ListIndex + 1, cboNameMode.ListIndex + 1, _" & vbCrLf
    s = s & "        txtCustomName.Text, txtOutputFolder.Text, chkCurrentCDR.Value, chkV15CDR.Value, _" & vbCrLf
    s = s & "        chkCurvesCDR.Value, chkCurvesV15.Value, chkPDF.Value)" & vbCrLf
    s = s & "    If Left$(result, 6) = ""ERROR:"" Then" & vbCrLf
    s = s & "        MsgBox result, vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Sub" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TryReadNumber(txtCDRMargin.Text, cdrMargin) Then Exit Sub" & vbCrLf
    s = s & "    If Not TryReadNumber(txtCDRWidth.Text, cdrWidth) Then Exit Sub" & vbCrLf
    s = s & "    If Not TryReadNumber(txtCDRHeight.Text, cdrHeight) Then Exit Sub" & vbCrLf
    s = s & "    If Not TryReadNumber(txtPDFMargin.Text, pdfMargin) Then Exit Sub" & vbCrLf
    s = s & "    If Not TryReadNumber(txtPDFWidth.Text, pdfWidth) Then Exit Sub" & vbCrLf
    s = s & "    If Not TryReadNumber(txtPDFHeight.Text, pdfHeight) Then Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    SetRunning True" & vbCrLf
    s = s & "    CPT_UI_RunExport CurrentProfileName(), cboScope.ListIndex + 1, cboNameMode.ListIndex + 1, _" & vbCrLf
    s = s & "        txtCustomName.Text, txtOutputFolder.Text, chkCurrentCDR.Value, chkV15CDR.Value, _" & vbCrLf
    s = s & "        chkCurvesCDR.Value, chkCurvesV15.Value, chkPDF.Value, _" & vbCrLf
    s = s & "        cboCDRPageMode.ListIndex + 1, cdrMargin, cdrWidth, cdrHeight, _" & vbCrLf
    s = s & "        cboPDFPageMode.ListIndex + 1, pdfMargin, pdfWidth, pdfHeight, _" & vbCrLf
    s = s & "        chkIncludeOutlines.Value, chkIncludeDimensions.Value, chkRenamePages.Value, _" & vbCrLf
    s = s & "        chkAutoScale.Value, chkEnsureUnique.Value" & vbCrLf
    s = s & "    SetRunning False" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "RunFail:" & vbCrLf
    s = s & "    SetRunning False" & vbCrLf
    s = s & "    MsgBox ""The workflow interface stopped: "" & Err.Description, vbCritical, ""Export Workflow""" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart4 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart5() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdCancel_Click()" & vbCrLf
    s = s & "    CPT_UI_RequestCancel" & vbCrLf
    s = s & "    lblProgressText.Caption = ""Cancellation requested. Completing the current safe step...""" & vbCrLf
    s = s & "    cmdCancel.Enabled = False" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cmdClose_Click()" & vbCrLf
    s = s & "    If mRunning Then" & vbCrLf
    s = s & "        MsgBox ""Cancel the current workflow before closing this window."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        Unload Me" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    s = s & "    If mRunning Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        MsgBox ""Cancel the current workflow before closing this window."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub BeginProgress(ByVal message As String)" & vbCrLf
    s = s & "    fraProgress.Visible = True" & vbCrLf
    s = s & "    lblProgressFill.Width = 60" & vbCrLf
    s = s & "    lblProgressText.Caption = message" & vbCrLf
    s = s & "    cmdCancel.Enabled = True" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Public Sub UpdateProgress(ByVal percentValue As Long, ByVal message As String)" & vbCrLf
    s = s & "    Dim maxWidth As Single" & vbCrLf
    s = s & "    If percentValue < 0 Then percentValue = 0" & vbCrLf
    s = s & "    If percentValue > 100 Then percentValue = 100" & vbCrLf
    s = s & "    maxWidth = 3900" & vbCrLf
    s = s & "    lblProgressFill.Width = CSng(maxWidth * percentValue / 100#)" & vbCrLf
    s = s & "    If lblProgressFill.Width < 60 Then lblProgressFill.Width = 60" & vbCrLf
    s = s & "    lblProgressText.Caption = CStr(percentValue) & ""% - "" & message" & vbCrLf
    s = s & "    DoEvents" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub SetRunning(ByVal runningValue As Boolean)" & vbCrLf
    s = s & "    mRunning = runningValue" & vbCrLf
    s = s & "    cmdRun.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdPreflight.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdSaveProfile.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdDeleteProfile.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdNewProfile.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdClose.Enabled = Not runningValue" & vbCrLf
    s = s & "    cmdCancel.Enabled = runningValue" & vbCrLf
    s = s & "    fraProgress.Visible = runningValue Or fraProgress.Visible" & vbCrLf
    s = s & "    If Not runningValue Then" & vbCrLf
    s = s & "        cmdCancel.Enabled = False" & vbCrLf
    s = s & "        UpdateSummary" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function ValidateNumericSettings() As Boolean" & vbCrLf
    s = s & "    Dim cdrMargin As Double" & vbCrLf
    s = s & "    Dim cdrWidth As Double" & vbCrLf
    s = s & "    Dim cdrHeight As Double" & vbCrLf
    s = s & "    Dim pdfMargin As Double" & vbCrLf
    s = s & "    Dim pdfWidth As Double" & vbCrLf
    s = s & "    Dim pdfHeight As Double" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If cboScope.ListIndex < 0 Or cboNameMode.ListIndex < 0 Then" & vbCrLf
    s = s & "        MsgBox ""Choose a process scope and filename source."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TryReadNumber(txtCDRMargin.Text, cdrMargin) Or cdrMargin < 0# Then" & vbCrLf
    s = s & "        MsgBox ""Enter a valid CDR margin."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Not TryReadNumber(txtPDFMargin.Text, pdfMargin) Or pdfMargin < 0# Then" & vbCrLf
    s = s & "        MsgBox ""Enter a valid PDF margin."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TryReadNumber(txtCDRWidth.Text, cdrWidth) Or _" & vbCrLf
    s = s & "       Not TryReadNumber(txtCDRHeight.Text, cdrHeight) Then" & vbCrLf
    s = s & "        MsgBox ""Enter valid custom CDR page dimensions."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If (cboCDRPageMode.ListIndex = 3 Or cboCDRPageMode.ListIndex = 4) And _" & vbCrLf
    s = s & "       (cdrWidth <= 0# Or cdrHeight <= 0#) Then" & vbCrLf
    s = s & "        MsgBox ""Custom CDR page width and height must be greater than zero."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If Not TryReadNumber(txtPDFWidth.Text, pdfWidth) Or _" & vbCrLf
    s = s & "       Not TryReadNumber(txtPDFHeight.Text, pdfHeight) Then" & vbCrLf
    s = s & "        MsgBox ""Enter valid custom PDF page dimensions."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If (cboPDFPageMode.ListIndex = 3 Or cboPDFPageMode.ListIndex = 4) And _" & vbCrLf
    s = s & "       (pdfWidth <= 0# Or pdfHeight <= 0#) Then" & vbCrLf
    s = s & "        MsgBox ""Custom PDF page width and height must be greater than zero."", vbExclamation, ""Export Workflow""" & vbCrLf
    s = s & "        Exit Function" & vbCrLf
    s = s & "    End If" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart5 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart6() As String
    Dim s As String
    s = s & "" & vbCrLf
    s = s & "    ValidateNumericSettings = True" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function TryReadNumber(ByVal valueText As String, ByRef result As Double) As Boolean" & vbCrLf
    s = s & "    Dim normalized As String" & vbCrLf
    s = s & "    normalized = Replace(Trim$(valueText), "","", ""."")" & vbCrLf
    s = s & "    If Len(normalized) = 0 Then Exit Function" & vbCrLf
    s = s & "    If normalized Like ""*[!0-9.-]*"" Then Exit Function" & vbCrLf
    s = s & "    On Error GoTo InvalidValue" & vbCrLf
    s = s & "    result = Val(normalized)" & vbCrLf
    s = s & "    TryReadNumber = True" & vbCrLf
    s = s & "    Exit Function" & vbCrLf
    s = s & "InvalidValue:" & vbCrLf
    s = s & "    TryReadNumber = False" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function CurrentProfileName() As String" & vbCrLf
    s = s & "    Dim captionText As String" & vbCrLf
    s = s & "    captionText = cboProfile.Value" & vbCrLf
    s = s & "    If Left$(captionText, 10) = ""Built-in: "" Then" & vbCrLf
    s = s & "        CurrentProfileName = Mid$(captionText, 11)" & vbCrLf
    s = s & "    ElseIf Left$(captionText, 8) = ""Custom: "" Then" & vbCrLf
    s = s & "        CurrentProfileName = Mid$(captionText, 9)" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        CurrentProfileName = ""Unsaved Custom Workflow""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateControlState()" & vbCrLf
    s = s & "    Dim cdrCustom As Boolean" & vbCrLf
    s = s & "    Dim pdfCustom As Boolean" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtCustomName.Enabled = (cboNameMode.ListIndex = 3)" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtCDRMargin.Enabled = (cboCDRPageMode.ListIndex = 2)" & vbCrLf
    s = s & "    cdrCustom = (cboCDRPageMode.ListIndex = 3 Or cboCDRPageMode.ListIndex = 4)" & vbCrLf
    s = s & "    txtCDRWidth.Enabled = cdrCustom" & vbCrLf
    s = s & "    txtCDRHeight.Enabled = cdrCustom" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtPDFMargin.Enabled = (cboPDFPageMode.ListIndex = 1)" & vbCrLf
    s = s & "    pdfCustom = (cboPDFPageMode.ListIndex = 3 Or cboPDFPageMode.ListIndex = 4)" & vbCrLf
    s = s & "    txtPDFWidth.Enabled = pdfCustom" & vbCrLf
    s = s & "    txtPDFHeight.Enabled = pdfCustom" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub UpdateSummary()" & vbCrLf
    s = s & "    Dim formats As String" & vbCrLf
    s = s & "    Dim cdrText As String" & vbCrLf
    s = s & "    Dim pdfText As String" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    If mLoading Then Exit Sub" & vbCrLf
    s = s & "    If chkCurrentCDR.Value Then formats = formats & ""Current CDR, """ & vbCrLf
    s = s & "    If chkV15CDR.Value Then formats = formats & ""V15 CDR, """ & vbCrLf
    s = s & "    If chkCurvesCDR.Value Then formats = formats & ""Curved CDR, """ & vbCrLf
    s = s & "    If chkCurvesV15.Value Then formats = formats & ""Curved V15, """ & vbCrLf
    s = s & "    If chkPDF.Value Then formats = formats & ""Curved PDF, """ & vbCrLf
    s = s & "    If Len(formats) > 2 Then formats = Left$(formats, Len(formats) - 2)" & vbCrLf
    s = s & "    If Len(formats) = 0 Then formats = ""None""" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case cboCDRPageMode.ListIndex" & vbCrLf
    s = s & "        Case 0: cdrText = ""Preserve source page exactly""" & vbCrLf
    s = s & "        Case 1: cdrText = ""Fit grouped objects, no margin""" & vbCrLf
    s = s & "        Case 2: cdrText = ""Fit grouped objects + "" & txtCDRMargin.Text & "" mm""" & vbCrLf
    s = s & "        Case 3: cdrText = ""Custom "" & txtCDRWidth.Text & "" x "" & txtCDRHeight.Text & "" mm, keep positions""" & vbCrLf
    s = s & "        Case 4: cdrText = ""Custom "" & txtCDRWidth.Text & "" x "" & txtCDRHeight.Text & "" mm, center objects""" & vbCrLf
    s = s & "        Case Else: cdrText = ""Not selected""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    Select Case cboPDFPageMode.ListIndex" & vbCrLf
    s = s & "        Case 0: pdfText = ""Fit grouped objects, no margin""" & vbCrLf
    s = s & "        Case 1: pdfText = ""Fit grouped objects + "" & txtPDFMargin.Text & "" mm""" & vbCrLf
    s = s & "        Case 2: pdfText = ""Preserve source page exactly""" & vbCrLf
    s = s & "        Case 3: pdfText = ""Custom "" & txtPDFWidth.Text & "" x "" & txtPDFHeight.Text & "" mm, keep positions""" & vbCrLf
    s = s & "        Case 4: pdfText = ""Custom "" & txtPDFWidth.Text & "" x "" & txtPDFHeight.Text & "" mm, center objects""" & vbCrLf
    s = s & "        Case Else: pdfText = ""Not selected""" & vbCrLf
    s = s & "    End Select" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "    txtSummary.Text = ""PROFILE"" & vbCrLf & CurrentProfileName() & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""PROCESS"" & vbCrLf & SafeComboText(cboScope) & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""NAMING"" & vbCrLf & SafeComboText(cboNameMode) & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""FORMATS"" & vbCrLf & formats & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""CDR PAGE"" & vbCrLf & cdrText & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""PDF PAGE"" & vbCrLf & pdfText & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""DIMENSIONS IN FIT"" & vbCrLf & IIf(chkIncludeDimensions.Value, ""Included"", ""Ignored for fitted size"") & vbCrLf & vbCrLf & _" & vbCrLf
    s = s & "        ""OUTPUT"" & vbCrLf & txtOutputFolder.Text" & vbCrLf
    s = s & "End Sub" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Function SafeComboText(ByVal cbo As Object) As String" & vbCrLf
    s = s & "    If cbo.ListIndex >= 0 Then SafeComboText = cbo.Value Else SafeComboText = ""Not selected""" & vbCrLf
    s = s & "End Function" & vbCrLf
    s = s & "" & vbCrLf
    s = s & "Private Sub cboScope_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub cboNameMode_Change(): UpdateControlState: UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub cboCDRPageMode_Change(): UpdateControlState: UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub cboPDFPageMode_Change(): UpdateControlState: UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtCustomName_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtCDRMargin_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtCDRWidth_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtCDRHeight_Change(): UpdateSummary: End Sub" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart6 = s
End Function

Private Function CPT27_frmVFEExportWorkflow_CodePart7() As String
    Dim s As String
    s = s & "Private Sub txtPDFMargin_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtPDFWidth_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtPDFHeight_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub txtOutputFolder_Change(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkCurrentCDR_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkV15CDR_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkCurvesCDR_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkCurvesV15_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkPDF_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkIncludeOutlines_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkIncludeDimensions_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkRenamePages_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkEnsureUnique_Click(): UpdateSummary: End Sub" & vbCrLf
    s = s & "Private Sub chkAutoScale_Click(): UpdateSummary: End Sub" & vbCrLf
    CPT27_frmVFEExportWorkflow_CodePart7 = s
End Function
