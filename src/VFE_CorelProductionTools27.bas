Attribute VB_Name = "VFE_CorelProductionTools27"
Option Explicit

' SPDX-License-Identifier: LicenseRef-VFE-Proprietary-1.0
' Licensed under the VFE Proprietary Software License v1.0.
' See LICENSE.md in the distribution package.

' ============================================================================
' VFE Corel Production Tools 27 - v0.4.0 FULL
' Unified CorelDRAW VBA macro for CorelDRAW 2026 / automation version 27
'
' Copyright (c) 2026 VFE Flavius. All rights reserved.
' Author and copyright holder: VFE Flavius
'
' Combines these workflows:
'   - Three-format export per page
'   - Five-format export per page
'   - Five-format export with all pages in one document
'   - Non-destructive text-to-curves CDR/PDF export
'   - Separate visual Dimension Tool
'   - Single visual Export Workflow with saved profiles
'
' Main entry point:
'   VFE_CorelProductionTools27
' Legacy compatibility entry point:
'   CorelProductionTools27
'
' Progress:
'   Uses CorelDRAW's native status-bar progress display.
'   Press Esc while an export is running to request cancellation.
' ============================================================================

Private Const CPT_VERSION As String = "0.4.0"
Private Const CPT_COPYRIGHT As String = "Copyright (c) 2026 VFE Flavius. All rights reserved."
Private Const NAME_SHAPE As String = "EXPORT_NAME"
Private Const CURVES_SUFFIX As String = "_curves"
Private Const V15_SUFFIX As String = "_v15"
Private Const DEFAULT_MARGIN_MM As Double = 5#
Private Const PDF_SAFE_MAX_MM As Double = 5080#
Private Const COREL_WARNING_MM As Double = 5000#
Private Const PDF_SCALE_RATIO As Double = 10#
Private Const DIMENSION_LAYER As String = "DIMENSIONS"
Private Const ERR_CANCELLED As Long = vbObjectError + 2701

Private Enum CPTExportScope
    cptScopeEachPage = 1
    cptScopeAllPages = 2
    cptScopeCurrentPage = 3
End Enum

Private Enum CPTNameMode
    cptNameExportField = 1
    cptNameDocument = 2
    cptNamePage = 3
    cptNameCustom = 4
End Enum

Private Enum CPTCDRPageMode
    cptCDRKeepOriginalPage = 1
    cptCDRFitNoMargin = 2
    cptCDRFitCustomMargin = 3
    cptCDRCustomKeepPositions = 4
    cptCDRCustomCenterObjects = 5
End Enum

Private Enum CPTPDFPageMode
    cptPDFFitNoMargin = 1
    cptPDFFitCustomMargin = 2
    cptPDFKeepOriginalPage = 3
    cptPDFCustomKeepPositions = 4
    cptPDFCustomCenterObjects = 5
End Enum

Private Enum CPTDimensionTarget
    cptDimensionCompleteSelection = 1
    cptDimensionEachObject = 2
End Enum

Private Enum CPTDimensionAxes
    cptDimensionWidthOnly = 1
    cptDimensionHeightOnly = 2
    cptDimensionWidthAndHeight = 3
End Enum

Private Enum CPTDimensionMode
    cptDimensionPermanentSource = 1
    cptDimensionTemporarySource = 2
    cptDimensionExportCopiesOnly = 3
End Enum

Private Type CPTExportOptions
    Scope As CPTExportScope
    NameMode As CPTNameMode
    customName As String
    outputFolder As String

    exportCurrentCDR As Boolean
    exportV15CDR As Boolean
    exportCurvesCurrentCDR As Boolean
    exportCurvesV15CDR As Boolean
    exportCurvesPDF As Boolean

    ' CDR page preparation. FitPageToArtwork/ MarginMM are retained for
    ' backward compatibility with the legacy prompt interface.
    FitPageToArtwork As Boolean
    includeOutlines As Boolean
    marginMM As Double
    CDRPageMode As CPTCDRPageMode
    cdrCustomWidthMM As Double
    cdrCustomHeightMM As Double

    PDFPageMode As CPTPDFPageMode
    pdfMarginMM As Double
    pdfCustomWidthMM As Double
    pdfCustomHeightMM As Double

    renameSourcePages As Boolean
    autoScaleLargePDF As Boolean
    ensureUniqueNames As Boolean
End Type

Private Type CPTDimensionOptions
    Enabled As Boolean
    targetMode As CPTDimensionTarget
    Axes As CPTDimensionAxes
    mode As CPTDimensionMode
    unitText As String
    decimals As Long
    fontName As String
    fontSizePt As Double
    lineOffsetMM As Double
    IncludeInPageFit As Boolean
    RemoveOldDimensions As Boolean
    groupCreated As Boolean
    ColorR As Long
    ColorG As Long
    ColorB As Long
End Type

Private Type CPTBoundingBox
    x As Double
    y As Double
    w As Double
    h As Double
End Type

Private Type CPTWorkflowProfile
    Name As String
    BuiltIn As Boolean
    DimensionOptions As CPTDimensionOptions
    ExportOptions As CPTExportOptions
End Type

' Legacy registry key retained so existing v0.4.0 profiles continue to load.
Private Const PROFILE_APP As String = "CorelProductionTools27"
Private Const PROFILE_INDEX_SECTION As String = "Profiles"
Private Const PROFILE_INDEX_KEY As String = "Names"

Private gWorkflowActive As Boolean
Private gWorkflowName As String
Private gWorkflowDimensionOptions As CPTDimensionOptions
Private gWorkflowDimensionSourcePage As Long
Private gWorkflowDimensionSpecCount As Long
Private gWorkflowDimensionSpecs() As CPTBoundingBox

Private gCancelRequested As Boolean
Private gExportUI As Object
Private gUIExportActive As Boolean
Private gUIIncludeDimensionsInFit As Boolean

' ============================================================================
' MAIN MENU
' ============================================================================
' Backward-compatible launcher for existing toolbar shortcuts.
Public Sub CorelProductionTools27()
    VFE_CorelProductionTools27
End Sub

Public Sub VFE_CorelProductionTools27()
    On Error GoTo FormFail
    frmVFELauncher.Show vbModeless
    Exit Sub

FormFail:
    MsgBox "The visual interface could not be opened." & vbCrLf & _
           "Run VFE_UI_Installer.VFE_InstallVisualInterface, then try again." & vbCrLf & vbCrLf & _
           "The legacy menu will open now.", vbExclamation, "VFE Corel Production Tools 27"
    CPT_LegacyMenu
End Sub

Public Sub CPT_LegacyMenu()
    Dim answer As String

    answer = InputBox( _
        "VFE COREL PRODUCTION TOOLS 27 - v" & CPT_VERSION & vbCrLf & _
        CPT_COPYRIGHT & vbCrLf & vbCrLf & _
        "Choose a workflow:" & vbCrLf & _
        "1  Export 3 formats per page" & vbCrLf & _
        "2  Export 5 formats per page" & vbCrLf & _
        "3  Export 5 formats, all pages in one document" & vbCrLf & _
        "4  Export non-destructive curves CDR + PDF" & vbCrLf & _
        "5  Add dimensions" & vbCrLf & _
        "6  Custom export" & vbCrLf & _
        "7  Convert text to curves on current page" & vbCrLf & _
        "8  Convert text to curves on all pages" & vbCrLf & vbCrLf & _
        "Press Esc during export to cancel.", _
        "VFE Corel Production Tools 27", "6")

    answer = Trim$(answer)
    If Len(answer) = 0 Then Exit Sub

    Select Case answer
        Case "1": CPT_Preset3FormatsPerPage
        Case "2": CPT_Preset5FormatsPerPage
        Case "3": CPT_Preset5FormatsAllPages
        Case "4": CPT_PresetCurvesCopy
        Case "5": CPT_AddWidthHeightDimensions
        Case "6": CPT_CustomExport
        Case "7": CPT_ConvertTextCurrentPage
        Case "8": CPT_ConvertTextAllPages
        Case Else
            MsgBox "Choose a number from 1 to 8.", vbExclamation, "VFE Corel Production Tools 27"
    End Select
End Sub

' ============================================================================
' PRODUCT INFORMATION
' ============================================================================
Public Function CPT_ProductCopyright() As String
    CPT_ProductCopyright = CPT_COPYRIGHT
End Function

' ============================================================================
' VISUAL INTERFACE API
' ============================================================================
Public Sub CPT_ShowDimensionTool()
    If Not CPT_HasActiveDocument() Then Exit Sub
    frmVFEDimensionTool.Show vbModeless
End Sub

Public Sub CPT_ShowExportWorkflow()
    If Not CPT_HasActiveDocument() Then Exit Sub
    frmVFEExportWorkflow.Show vbModeless
End Sub

Public Sub CPT_ConvertTextCurrentPage()
    CPT_ConvertTextInSource False
End Sub

Public Sub CPT_ConvertTextAllPages()
    CPT_ConvertTextInSource True
End Sub

Private Sub CPT_ConvertTextInSource(ByVal allPages As Boolean)
    Dim doc As Document
    Dim oldOptimization As Boolean
    Dim oldEvents As Boolean
    Dim originalPageIndex As Long
    Dim commandStarted As Boolean
    Dim converted As Long
    Dim i As Long
    Dim scopeText As String
    Dim answer As VbMsgBoxResult
    Dim errNumber As Long
    Dim errDescription As String

    If Not CPT_HasActiveDocument() Then Exit Sub
    Set doc = ActiveDocument
    originalPageIndex = doc.ActivePage.Index

    If allPages Then
        scopeText = "all pages"
    Else
        scopeText = "the current page"
    End If

    answer = MsgBox( _
        "Convert every editable text object on " & scopeText & " to curves?" & vbCrLf & vbCrLf & _
        "Text inside groups and PowerClips is included." & vbCrLf & _
        "This changes the source document. You can undo the operation as one step.", _
        vbYesNo + vbExclamation, "Convert Text to Curves")
    If answer <> vbYes Then Exit Sub

    oldOptimization = Application.Optimization
    oldEvents = Application.EventsEnabled
    On Error GoTo ConvertFail

    doc.BeginCommandGroup "VFE CPT27 Convert Text to Curves"
    commandStarted = True
    Application.Optimization = True
    Application.EventsEnabled = False

    If allPages Then
        For i = 1 To doc.Pages.Count
            doc.Pages(i).Activate
            converted = converted + CPT_ConvertAllTextOnPage(doc.Pages(i))
            DoEvents
        Next i
    Else
        converted = CPT_ConvertAllTextOnPage(doc.ActivePage)
    End If

ConvertExit:
    On Error Resume Next
    If commandStarted Then doc.EndCommandGroup
    doc.Pages(originalPageIndex).Activate
    Application.EventsEnabled = oldEvents
    Application.Optimization = oldOptimization
    ActiveWindow.Refresh
    On Error GoTo 0

    If errNumber <> 0 Then
        MsgBox "Text conversion stopped." & vbCrLf & vbCrLf & _
               "Error " & CStr(errNumber) & ": " & errDescription, _
               vbCritical, "VFE Corel Production Tools 27"
    Else
        MsgBox CStr(converted) & " text object(s) converted to curves on " & scopeText & "." & vbCrLf & _
               "Locked or protected text objects were skipped.", _
               vbInformation, "VFE Corel Production Tools 27"
    End If
    Exit Sub

ConvertFail:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume ConvertExit
End Sub

Public Sub CPT_UI_RegisterExportForm(ByVal formRef As Object)
    Set gExportUI = formRef
End Sub

Public Sub CPT_UI_UnregisterExportForm(ByVal formRef As Object)
    On Error Resume Next
    If gExportUI Is formRef Then Set gExportUI = Nothing
    On Error GoTo 0
End Sub

Public Sub CPT_UI_RequestCancel()
    gCancelRequested = True
End Sub

Public Function CPT_UI_BrowseFolder(ByVal title As String) As String
    CPT_UI_BrowseFolder = CPT_GetFolderFromUser(title)
End Function

Public Function CPT_UI_DefaultOutputFolder() As String
    Dim fullName As String
    Dim p As Long

    If Documents.Count = 0 Then Exit Function
    On Error Resume Next
    fullName = ActiveDocument.FullFileName
    On Error GoTo 0
    p = InStrRev(fullName, "\")
    If p > 0 Then CPT_UI_DefaultOutputFolder = Left$(fullName, p)
End Function

Public Function CPT_UI_GetSelectionSummary() As String
    Dim sr As ShapeRange
    Dim x As Double, y As Double, w As Double, h As Double
    Dim oldUnit As cdrUnit

    If Documents.Count = 0 Then
        CPT_UI_GetSelectionSummary = "No active document"
        Exit Function
    End If
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        CPT_UI_GetSelectionSummary = "No objects selected"
        Exit Function
    End If

    oldUnit = ActiveDocument.Unit
    On Error GoTo SummaryFail
    ActiveDocument.Unit = cdrMillimeter
    sr.GetBoundingBox x, y, w, h, True
    ActiveDocument.Unit = oldUnit
    CPT_UI_GetSelectionSummary = CStr(sr.Count) & " object(s) selected | " & _
                                 Format$(w, "0.00") & " x " & Format$(h, "0.00") & " mm"
    Exit Function
SummaryFail:
    On Error Resume Next
    ActiveDocument.Unit = oldUnit
    On Error GoTo 0
    CPT_UI_GetSelectionSummary = CStr(sr.Count) & " object(s) selected"
End Function

Public Sub CPT_UI_AddDimensions( _
    ByVal targetMode As Long, _
    ByVal addWidth As Boolean, _
    ByVal addHeight As Boolean, _
    ByVal unitText As String, _
    ByVal decimals As Long, _
    ByVal fontName As String, _
    ByVal fontSizePt As Double, _
    ByVal lineOffsetMM As Double, _
    ByVal removeOld As Boolean, _
    ByVal groupCreated As Boolean, _
    ByVal colorName As String)

    On Error GoTo AddDimensionsFail

    Dim dimOpt As CPTDimensionOptions
    Dim specs() As CPTBoundingBox
    Dim specCount As Long
    Dim createdShape As Shape

    If Not CPT_HasActiveDocument() Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Select one or more objects first.", vbExclamation, "Dimension Tool"
        Exit Sub
    End If
    If Not addWidth And Not addHeight Then
        MsgBox "Select Width, Height, or both.", vbExclamation, "Dimension Tool"
        Exit Sub
    End If

    CPT_SetDefaultDimensionOptions dimOpt
    If targetMode = 1 Then
        dimOpt.targetMode = cptDimensionCompleteSelection
    Else
        dimOpt.targetMode = cptDimensionEachObject
    End If

    If addWidth And addHeight Then
        dimOpt.Axes = cptDimensionWidthAndHeight
    ElseIf addWidth Then
        dimOpt.Axes = cptDimensionWidthOnly
    Else
        dimOpt.Axes = cptDimensionHeightOnly
    End If

    unitText = LCase$(Trim$(unitText))
    If unitText <> "mm" And unitText <> "cm" And unitText <> "in" Then unitText = "cm"
    dimOpt.unitText = unitText
    If decimals < 0 Then decimals = 0
    If decimals > 4 Then decimals = 4
    dimOpt.decimals = decimals
    dimOpt.fontName = Trim$(fontName)
    If Len(dimOpt.fontName) = 0 Then dimOpt.fontName = "Arial"
    If fontSizePt <= 0# Then fontSizePt = 10#
    dimOpt.fontSizePt = fontSizePt
    If lineOffsetMM <= 0# Then lineOffsetMM = 5#
    dimOpt.lineOffsetMM = lineOffsetMM
    dimOpt.RemoveOldDimensions = removeOld
    dimOpt.groupCreated = groupCreated
    CPT_SetDimensionColor dimOpt, colorName

    If Not CPT_CaptureSelectionSpecs(ActiveSelectionRange, dimOpt.targetMode, specs, specCount) Then Exit Sub
    If dimOpt.RemoveOldDimensions Then CPT_RemoveDimensionLayerContents ActiveDocument.ActivePage
    Set createdShape = CPT_CreateDimensionsFromSpecs(ActiveDocument.ActivePage, dimOpt, specs, specCount)

    ActiveWindow.Refresh
    If createdShape Is Nothing Then
        MsgBox "No dimensions were created.", vbExclamation, "Dimension Tool"
    Else
        MsgBox "Dimensions added on layer '" & DIMENSION_LAYER & "'." & vbCrLf & _
               "Targets measured: " & CStr(specCount), vbInformation, "Dimension Tool"
    End If
    Exit Sub

AddDimensionsFail:
    MsgBox "The dimensions could not be created." & vbCrLf & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, _
           vbCritical, "VFE Corel Production Tools 27 - Dimension Tool"
End Sub

Public Sub CPT_UI_RunExport( _
    ByVal profileName As String, _
    ByVal scopeValue As Long, _
    ByVal nameModeValue As Long, _
    ByVal customName As String, _
    ByVal outputFolder As String, _
    ByVal exportCurrentCDR As Boolean, _
    ByVal exportV15CDR As Boolean, _
    ByVal exportCurvesCurrentCDR As Boolean, _
    ByVal exportCurvesV15CDR As Boolean, _
    ByVal exportCurvesPDF As Boolean, _
    ByVal cdrPageModeValue As Long, _
    ByVal cdrMarginMM As Double, _
    ByVal cdrCustomWidthMM As Double, _
    ByVal cdrCustomHeightMM As Double, _
    ByVal pdfPageModeValue As Long, _
    ByVal pdfMarginMM As Double, _
    ByVal pdfCustomWidthMM As Double, _
    ByVal pdfCustomHeightMM As Double, _
    ByVal includeOutlines As Boolean, _
    ByVal includeDimensionsInFit As Boolean, _
    ByVal renameSourcePages As Boolean, _
    ByVal autoScaleLargePDF As Boolean, _
    ByVal ensureUniqueNames As Boolean)

    Dim opt As CPTExportOptions

    CPT_SetDefaultOptions opt
    opt.Scope = scopeValue
    opt.NameMode = nameModeValue
    opt.customName = customName
    opt.outputFolder = outputFolder
    opt.exportCurrentCDR = exportCurrentCDR
    opt.exportV15CDR = exportV15CDR
    opt.exportCurvesCurrentCDR = exportCurvesCurrentCDR
    opt.exportCurvesV15CDR = exportCurvesV15CDR
    opt.exportCurvesPDF = exportCurvesPDF

    opt.CDRPageMode = cdrPageModeValue
    opt.marginMM = cdrMarginMM
    opt.cdrCustomWidthMM = cdrCustomWidthMM
    opt.cdrCustomHeightMM = cdrCustomHeightMM
    opt.FitPageToArtwork = (opt.CDRPageMode = cptCDRFitNoMargin Or opt.CDRPageMode = cptCDRFitCustomMargin)

    opt.PDFPageMode = pdfPageModeValue
    opt.pdfMarginMM = pdfMarginMM
    opt.pdfCustomWidthMM = pdfCustomWidthMM
    opt.pdfCustomHeightMM = pdfCustomHeightMM

    opt.includeOutlines = includeOutlines
    opt.renameSourcePages = renameSourcePages
    opt.autoScaleLargePDF = autoScaleLargePDF
    opt.ensureUniqueNames = ensureUniqueNames

    gCancelRequested = False
    gUIExportActive = True
    gUIIncludeDimensionsInFit = includeDimensionsInFit
    gWorkflowActive = True
    gWorkflowName = Trim$(profileName)
    CPT_SetDefaultDimensionOptions gWorkflowDimensionOptions
    gWorkflowDimensionOptions.Enabled = False

    CPT_RunExport opt

    CPT_ClearWorkflowContext
    gUIExportActive = False
End Sub

Public Function CPT_UI_Preflight( _
    ByVal scopeValue As Long, _
    ByVal nameModeValue As Long, _
    ByVal customName As String, _
    ByVal outputFolder As String, _
    ByVal exportCurrentCDR As Boolean, _
    ByVal exportV15CDR As Boolean, _
    ByVal exportCurvesCurrentCDR As Boolean, _
    ByVal exportCurvesV15CDR As Boolean, _
    ByVal exportCurvesPDF As Boolean) As String

    Dim doc As Document
    Dim firstPage As Long, lastPage As Long, i As Long
    Dim warnings As String, details As String, rawName As String, safeName As String
    Dim names As Object
    Dim fileTotal As Long, formatCount As Long

    If Documents.Count = 0 Then
        CPT_UI_Preflight = "ERROR: No active CorelDRAW document."
        Exit Function
    End If
    If Not (exportCurrentCDR Or exportV15CDR Or exportCurvesCurrentCDR Or exportCurvesV15CDR Or exportCurvesPDF) Then
        CPT_UI_Preflight = "ERROR: No export format selected."
        Exit Function
    End If
    outputFolder = CPT_NormalizeFolder(outputFolder)
    If Len(outputFolder) = 0 Or Not CPT_FolderExists(outputFolder) Then
        CPT_UI_Preflight = "ERROR: Choose an existing output folder."
        Exit Function
    End If
    If nameModeValue = cptNameCustom And Len(Trim$(customName)) = 0 Then
        CPT_UI_Preflight = "ERROR: Enter a custom filename or prefix."
        Exit Function
    End If

    Set doc = ActiveDocument
    If scopeValue = cptScopeCurrentPage Then
        firstPage = doc.ActivePage.Index: lastPage = firstPage
    Else
        firstPage = 1: lastPage = doc.Pages.Count
    End If

    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    For i = firstPage To lastPage
        If doc.Pages(i).Shapes.Count = 0 Then
            CPT_AppendLine warnings, "Page " & CStr(i) & " is empty."
        End If

        If nameModeValue = cptNameExportField Then
            If Len(Trim$(CPT_FindTextByShapeName(doc.Pages(i).Shapes, NAME_SHAPE))) = 0 Then
                If scopeValue <> cptScopeAllPages Or i = firstPage Then
                    CPT_AppendLine warnings, "Page " & CStr(i) & " has no text object named " & NAME_SHAPE & "; a fallback filename will be used."
                End If
            End If
        End If

        If scopeValue = cptScopeAllPages Then
            If i = firstPage Then
                rawName = CPT_UI_ResolveAllNameForPreflight(doc, nameModeValue, customName)
            Else
                rawName = ""
            End If
        Else
            rawName = CPT_UI_ResolveNameForPreflight(doc, doc.Pages(i), nameModeValue, customName, i)
        End If

        If Len(rawName) > 0 Then
            safeName = CPT_MakeSafeFileName(rawName)
            If Len(safeName) = 0 Then
                CPT_AppendLine warnings, "Page " & CStr(i) & " produces an empty filename."
            ElseIf names.Exists(safeName) Then
                CPT_AppendLine warnings, "Duplicate filename: " & safeName & " (pages " & names(safeName) & " and " & CStr(i) & ")."
            Else
                names.Add safeName, CStr(i)
            End If
        End If
    Next i

    If exportCurrentCDR Then formatCount = formatCount + 1
    If exportV15CDR Then formatCount = formatCount + 1
    If exportCurvesCurrentCDR Then formatCount = formatCount + 1
    If exportCurvesV15CDR Then formatCount = formatCount + 1
    If exportCurvesPDF Then formatCount = formatCount + 1
    If scopeValue = cptScopeAllPages Then
        fileTotal = formatCount
    Else
        fileTotal = (lastPage - firstPage + 1) * formatCount
    End If

    details = "PRE-FLIGHT COMPLETE" & vbCrLf & String$(34, "-") & vbCrLf & _
              "Pages to process: " & CStr(lastPage - firstPage + 1) & vbCrLf & _
              "Formats selected: " & CStr(formatCount) & vbCrLf & _
              "Expected output files: " & CStr(fileTotal) & vbCrLf & _
              "Output folder: " & outputFolder
    If Len(warnings) > 0 Then
        details = details & vbCrLf & vbCrLf & "WARNINGS" & vbCrLf & warnings
    Else
        details = details & vbCrLf & vbCrLf & "No blocking problems found."
    End If
    CPT_UI_Preflight = details
End Function

Private Function CPT_UI_ResolveAllNameForPreflight( _
    ByVal doc As Document, ByVal nameModeValue As Long, ByVal customName As String) As String

    Dim opt As CPTExportOptions
    CPT_SetDefaultOptions opt
    opt.NameMode = nameModeValue
    opt.customName = customName
    CPT_UI_ResolveAllNameForPreflight = CPT_ResolveAllPagesBaseName(doc, opt)
End Function

Private Function CPT_UI_ResolveNameForPreflight( _
    ByVal doc As Document, ByVal pg As Page, ByVal nameModeValue As Long, _
    ByVal customName As String, ByVal pageIndex As Long) As String

    Dim opt As CPTExportOptions
    CPT_SetDefaultOptions opt
    opt.NameMode = nameModeValue
    opt.customName = customName
    CPT_UI_ResolveNameForPreflight = CPT_ResolvePageName(doc, pg, opt, pageIndex)
End Function

Private Sub CPT_SetDimensionColor(ByRef dimOpt As CPTDimensionOptions, ByVal colorName As String)
    Select Case LCase$(Trim$(colorName))
        Case "red": dimOpt.ColorR = 255: dimOpt.ColorG = 0: dimOpt.ColorB = 0
        Case "green": dimOpt.ColorR = 0: dimOpt.ColorG = 160: dimOpt.ColorB = 0
        Case "black": dimOpt.ColorR = 0: dimOpt.ColorG = 0: dimOpt.ColorB = 0
        Case Else: dimOpt.ColorR = 0: dimOpt.ColorG = 0: dimOpt.ColorB = 255
    End Select
End Sub

' ============================================================================
' PRESETS THAT REPLACE THE ORIGINAL MACROS
' ============================================================================
Public Sub CPT_Preset3FormatsPerPage()
    Dim opt As CPTExportOptions

    If Not CPT_HasActiveDocument() Then Exit Sub

    CPT_SetDefaultOptions opt
    opt.Scope = cptScopeEachPage
    opt.NameMode = cptNameExportField
    opt.exportCurrentCDR = True
    opt.exportCurvesCurrentCDR = True
    opt.exportCurvesPDF = True
    opt.FitPageToArtwork = False
    opt.CDRPageMode = cptCDRKeepOriginalPage
    opt.PDFPageMode = cptPDFKeepOriginalPage
    opt.renameSourcePages = True

    opt.outputFolder = CPT_GetFolderFromUser("Choose folder for 3-format page export")
    If Len(opt.outputFolder) = 0 Then Exit Sub

    CPT_RunExport opt
End Sub

Public Sub CPT_Preset5FormatsPerPage()
    Dim opt As CPTExportOptions
    Dim mode As Long

    If Not CPT_HasActiveDocument() Then Exit Sub

    mode = CPT_AskNameModeExportOrDocument()
    If mode = 0 Then Exit Sub

    CPT_SetDefaultOptions opt
    opt.Scope = cptScopeEachPage
    opt.NameMode = mode
    opt.exportCurrentCDR = True
    opt.exportV15CDR = True
    opt.exportCurvesCurrentCDR = True
    opt.exportCurvesV15CDR = True
    opt.exportCurvesPDF = True
    opt.FitPageToArtwork = True
    opt.CDRPageMode = cptCDRFitCustomMargin
    opt.PDFPageMode = cptPDFFitCustomMargin
    opt.pdfMarginMM = DEFAULT_MARGIN_MM
    opt.renameSourcePages = True

    opt.outputFolder = CPT_GetFolderFromUser("Choose folder for 5-format page export")
    If Len(opt.outputFolder) = 0 Then Exit Sub

    CPT_RunExport opt
End Sub

Public Sub CPT_Preset5FormatsAllPages()
    Dim opt As CPTExportOptions

    If Not CPT_HasActiveDocument() Then Exit Sub

    CPT_SetDefaultOptions opt
    opt.Scope = cptScopeAllPages
    opt.NameMode = cptNameDocument
    opt.exportCurrentCDR = True
    opt.exportV15CDR = True
    opt.exportCurvesCurrentCDR = True
    opt.exportCurvesV15CDR = True
    opt.exportCurvesPDF = True
    opt.FitPageToArtwork = True
    opt.CDRPageMode = cptCDRFitCustomMargin
    opt.PDFPageMode = cptPDFFitCustomMargin
    opt.pdfMarginMM = DEFAULT_MARGIN_MM
    opt.renameSourcePages = True

    opt.outputFolder = CPT_GetFolderFromUser("Choose folder for all-pages 5-format export")
    If Len(opt.outputFolder) = 0 Then Exit Sub

    CPT_RunExport opt
End Sub

Public Sub CPT_PresetCurvesCopy()
    Dim opt As CPTExportOptions

    If Not CPT_HasActiveDocument() Then Exit Sub

    CPT_SetDefaultOptions opt
    opt.Scope = cptScopeAllPages
    opt.NameMode = cptNameDocument
    opt.exportCurvesCurrentCDR = True
    opt.exportCurvesPDF = True
    opt.FitPageToArtwork = False
    opt.CDRPageMode = cptCDRKeepOriginalPage
    opt.PDFPageMode = cptPDFKeepOriginalPage
    opt.renameSourcePages = False

    opt.outputFolder = CPT_GetFolderFromUser("Choose folder for curves CDR and PDF")
    If Len(opt.outputFolder) = 0 Then Exit Sub

    CPT_RunExport opt
End Sub

' ============================================================================
' CUSTOM EXPORT PROMPT INTERFACE
' ============================================================================
Public Sub CPT_CustomExport()
    Dim opt As CPTExportOptions

    If Not CPT_HasActiveDocument() Then Exit Sub
    CPT_SetDefaultOptions opt

    If Not CPT_PromptExportOptions(opt, "Custom Export") Then Exit Sub
    CPT_RunExport opt
End Sub

' ============================================================================
' EXPORT CONTROLLER
' ============================================================================
Private Sub CPT_RunExport(ByRef opt As CPTExportOptions)
    Dim srcDoc As Document
    Dim oldOptimization As Boolean
    Dim oldEvents As Boolean
    Dim originalPageIndex As Long
    Dim fileCount As Long
    Dim textCount As Long
    Dim warningText As String
    Dim logText As String
    Dim logPath As String
    Dim errNumber As Long
    Dim errDescription As String
    Dim doneMessage As String

    If Not CPT_HasActiveDocument() Then Exit Sub
    If Not CPT_AnyFormatSelected(opt) Then
        MsgBox "No output format was selected.", vbExclamation
        Exit Sub
    End If

    opt.outputFolder = CPT_NormalizeFolder(opt.outputFolder)
    If Len(opt.outputFolder) = 0 Then Exit Sub
    If Not CPT_FolderExists(opt.outputFolder) Then
        MsgBox "The output folder does not exist:" & vbCrLf & opt.outputFolder, vbExclamation
        Exit Sub
    End If

    Set srcDoc = ActiveDocument
    originalPageIndex = srcDoc.ActivePage.Index
    oldOptimization = Application.Optimization
    oldEvents = Application.EventsEnabled

    On Error GoTo CleanFail

    Application.Optimization = True
    Application.EventsEnabled = False
    CPT_ProgressBegin "Preparing Corel Production Tools export..."

    logText = "VFE Corel Production Tools 27 - v" & CPT_VERSION & vbCrLf & _
              CPT_COPYRIGHT & vbCrLf & _
              "Started: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Source: " & CPT_GetDocumentDisplayName(srcDoc) & vbCrLf & _
              "Folder: " & opt.outputFolder & vbCrLf

    If gWorkflowActive Then
        logText = logText & "Workflow profile: " & gWorkflowName & vbCrLf
        If gWorkflowDimensionOptions.Enabled Then
            logText = logText & "Dimensions: " & CPT_DimensionOptionsSummary(gWorkflowDimensionOptions) & vbCrLf
        End If
    End If
    logText = logText & "CDR page mode: " & CPT_CDRPageModeText(opt.CDRPageMode, opt.marginMM, opt.cdrCustomWidthMM, opt.cdrCustomHeightMM) & vbCrLf & _
              "PDF page mode: " & CPT_PDFPageModeText(opt.PDFPageMode, opt.pdfMarginMM, opt.pdfCustomWidthMM, opt.pdfCustomHeightMM) & vbCrLf & _
              String$(72, "-") & vbCrLf

    Select Case opt.Scope
        Case cptScopeEachPage
            CPT_ExportSeparatePages srcDoc, opt, False, fileCount, textCount, warningText, logText
        Case cptScopeCurrentPage
            CPT_ExportSeparatePages srcDoc, opt, True, fileCount, textCount, warningText, logText
        Case cptScopeAllPages
            CPT_ExportAllPagesTogether srcDoc, opt, fileCount, textCount, warningText, logText
        Case Else
            Err.Raise vbObjectError + 2702, , "Unknown export scope."
    End Select

    logText = logText & String$(72, "-") & vbCrLf & _
              "Completed: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
              "Files created: " & CStr(fileCount) & vbCrLf & _
              "Text objects converted in temporary copies: " & CStr(textCount) & vbCrLf

    If Len(warningText) > 0 Then
        logText = logText & "Warnings:" & vbCrLf & warningText & vbCrLf
    End If

    logPath = opt.outputFolder & "VFE_CPT27_export_log_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    CPT_WriteTextFile logPath, logText

CleanExit:
    On Error Resume Next
    srcDoc.Pages(originalPageIndex).Activate
    Application.Status.EndProgress
    On Error Resume Next
    If Not gExportUI Is Nothing Then gExportUI.UpdateProgress 100, "Workflow finished"
    On Error GoTo 0
    Application.EventsEnabled = oldEvents
    Application.Optimization = oldOptimization
    ActiveWindow.Refresh
    On Error GoTo 0

    If errNumber = ERR_CANCELLED Then
        MsgBox "Export cancelled." & vbCrLf & _
               "Any files completed before cancellation were kept.", _
               vbInformation, "VFE Corel Production Tools 27"
    ElseIf errNumber <> 0 Then
        MsgBox "Export stopped:" & vbCrLf & errDescription, vbCritical, "VFE Corel Production Tools 27"
    Else
        doneMessage = "Export complete." & vbCrLf & vbCrLf & _
                      "Files created: " & CStr(fileCount) & vbCrLf & _
                      "Text objects converted in temporary copies: " & CStr(textCount) & vbCrLf & _
                      "Log: " & logPath

        If Len(warningText) > 0 Then
            doneMessage = doneMessage & vbCrLf & vbCrLf & "Large-format notes were written to the log."
        End If

        MsgBox doneMessage, vbInformation, "VFE Corel Production Tools 27"
    End If
    Exit Sub

CleanFail:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume CleanExit
End Sub

' ============================================================================
' SEPARATE PAGE EXPORT
' ============================================================================
Private Sub CPT_ExportSeparatePages( _
    ByVal srcDoc As Document, _
    ByRef opt As CPTExportOptions, _
    ByVal currentPageOnly As Boolean, _
    ByRef fileCount As Long, _
    ByRef textCount As Long, _
    ByRef warningText As String, _
    ByRef logText As String)

    Dim firstIndex As Long
    Dim lastIndex As Long
    Dim pageIndex As Long
    Dim pageNumberInJob As Long
    Dim totalPagesInJob As Long
    Dim srcPage As Page
    Dim normalDoc As Document
    Dim curvesDoc As Document
    Dim pdfDoc As Document
    Dim rawName As String
    Dim safeName As String
    Dim uniqueName As String
    Dim pageWarning As String
    Dim converted As Long
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo PageFail

    If currentPageOnly Then
        firstIndex = srcDoc.ActivePage.Index
        lastIndex = firstIndex
    Else
        firstIndex = 1
        lastIndex = srcDoc.Pages.Count
    End If

    totalPagesInJob = lastIndex - firstIndex + 1

    For pageIndex = firstIndex To lastIndex
        pageNumberInJob = pageIndex - firstIndex + 1
        CPT_CheckCancelled

        Set srcPage = srcDoc.Pages(pageIndex)
        srcPage.Activate

        rawName = CPT_ResolvePageName(srcDoc, srcPage, opt, pageIndex)
        safeName = CPT_MakeSafeFileName(rawName)
        uniqueName = CPT_ResolveUniqueBaseName(opt.outputFolder, safeName, opt)

        If opt.renameSourcePages Then
            On Error Resume Next
            srcPage.Name = safeName
            On Error GoTo 0
        End If

        CPT_ProgressSet pageNumberInJob, totalPagesInJob, _
            "Page " & CStr(pageNumberInJob) & " of " & CStr(totalPagesInJob) & ": " & uniqueName

        ' -------------------- Normal CDR copies --------------------
        If opt.exportCurrentCDR Or opt.exportV15CDR Then
            Set normalDoc = CPT_CreateSinglePageDocument(srcDoc, srcPage)
            normalDoc.ActivePage.Name = uniqueName
            CPT_ApplyWorkflowDimensionsToTempPage normalDoc.ActivePage, pageIndex

            CPT_PrepareCDRPage normalDoc, normalDoc.ActivePage, opt

            CPT_SaveNormalFormats normalDoc, opt.outputFolder, uniqueName, opt, fileCount, logText
            CPT_CloseTemporaryDocument normalDoc
        End If

        CPT_CheckCancelled

        ' -------------------- Curved CDR copies --------------------
        If opt.exportCurvesCurrentCDR Or opt.exportCurvesV15CDR Then
            Set curvesDoc = CPT_CreateSinglePageDocument(srcDoc, srcPage)
            curvesDoc.ActivePage.Name = uniqueName & CURVES_SUFFIX
            CPT_ApplyWorkflowDimensionsToTempPage curvesDoc.ActivePage, pageIndex

            converted = CPT_ConvertAllTextOnPage(curvesDoc.ActivePage)
            textCount = textCount + converted

            CPT_PrepareCDRPage curvesDoc, curvesDoc.ActivePage, opt

            CPT_SaveCurvedFormats curvesDoc, opt.outputFolder, uniqueName, opt, fileCount, logText
            CPT_CloseTemporaryDocument curvesDoc
        End If

        CPT_CheckCancelled

        ' -------------------- PDF has its own temporary copy --------------------
        ' This guarantees large-format PDF scaling never changes curved CDR output.
        If opt.exportCurvesPDF Then
            Set pdfDoc = CPT_CreateSinglePageDocument(srcDoc, srcPage)
            pdfDoc.ActivePage.Name = uniqueName & CURVES_SUFFIX
            CPT_ApplyWorkflowDimensionsToTempPage pdfDoc.ActivePage, pageIndex

            converted = CPT_ConvertAllTextOnPage(pdfDoc.ActivePage)
            textCount = textCount + converted

            CPT_PreparePDFPage pdfDoc, pdfDoc.ActivePage, opt

            pageWarning = CPT_HandleLargePDFPage(pdfDoc, pdfDoc.ActivePage, uniqueName, opt.autoScaleLargePDF)
            If Len(pageWarning) > 0 Then
                CPT_AppendLine warningText, pageWarning
                CPT_AppendLine logText, "WARNING: " & pageWarning
            End If

            CPT_SetupPDFSettings pdfDoc
            pdfDoc.PublishToPDF opt.outputFolder & uniqueName & CURVES_SUFFIX & ".pdf"
            fileCount = fileCount + 1
            CPT_AppendLine logText, "PDF: " & opt.outputFolder & uniqueName & CURVES_SUFFIX & ".pdf"

            CPT_CloseTemporaryDocument pdfDoc
        End If
    Next pageIndex

    Exit Sub

PageFail:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    CPT_CloseTemporaryDocument normalDoc
    CPT_CloseTemporaryDocument curvesDoc
    CPT_CloseTemporaryDocument pdfDoc
    Err.Raise savedErrNumber, "CPT_ExportSeparatePages", savedErrDescription
End Sub

' ============================================================================
' ALL PAGES IN ONE DOCUMENT EXPORT
' ============================================================================
Private Sub CPT_ExportAllPagesTogether( _
    ByVal srcDoc As Document, _
    ByRef opt As CPTExportOptions, _
    ByRef fileCount As Long, _
    ByRef textCount As Long, _
    ByRef warningText As String, _
    ByRef logText As String)

    Dim normalDoc As Document
    Dim curvesDoc As Document
    Dim pdfDoc As Document
    Dim rawName As String
    Dim safeName As String
    Dim uniqueName As String
    Dim converted As Long
    Dim pdfWarnings As String
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error GoTo AllPagesFail

    rawName = CPT_ResolveAllPagesBaseName(srcDoc, opt)
    safeName = CPT_MakeSafeFileName(rawName)
    uniqueName = CPT_ResolveUniqueBaseName(opt.outputFolder, safeName, opt)

    If opt.renameSourcePages Then CPT_RenamePagesFromExportField srcDoc

    CPT_ProgressSet 1, 3, "Preparing normal all-pages document: " & uniqueName

    If opt.exportCurrentCDR Or opt.exportV15CDR Then
        Set normalDoc = CPT_DuplicateWholeDocument(srcDoc)
        CPT_RenamePagesFromExportField normalDoc
        CPT_ApplyWorkflowDimensionsToDocument normalDoc

        CPT_PrepareAllCDRPages normalDoc, opt

        CPT_SaveNormalFormats normalDoc, opt.outputFolder, uniqueName, opt, fileCount, logText
        CPT_CloseTemporaryDocument normalDoc
    End If

    CPT_CheckCancelled
    CPT_ProgressSet 2, 3, "Preparing curved all-pages document: " & uniqueName

    If opt.exportCurvesCurrentCDR Or opt.exportCurvesV15CDR Then
        Set curvesDoc = CPT_DuplicateWholeDocument(srcDoc)
        CPT_RenamePagesFromExportField curvesDoc
        CPT_ApplyWorkflowDimensionsToDocument curvesDoc
        converted = CPT_ConvertAllTextInDocument(curvesDoc)
        textCount = textCount + converted

        CPT_PrepareAllCDRPages curvesDoc, opt

        CPT_SaveCurvedFormats curvesDoc, opt.outputFolder, uniqueName, opt, fileCount, logText
        CPT_CloseTemporaryDocument curvesDoc
    End If

    CPT_CheckCancelled
    CPT_ProgressSet 3, 3, "Publishing curved all-pages PDF: " & uniqueName

    If opt.exportCurvesPDF Then
        Set pdfDoc = CPT_DuplicateWholeDocument(srcDoc)
        CPT_RenamePagesFromExportField pdfDoc
        CPT_ApplyWorkflowDimensionsToDocument pdfDoc
        converted = CPT_ConvertAllTextInDocument(pdfDoc)
        textCount = textCount + converted

        CPT_PrepareAllPDFPages pdfDoc, opt

        pdfWarnings = CPT_HandleLargePDFDocument(pdfDoc, opt.autoScaleLargePDF)
        If Len(pdfWarnings) > 0 Then
            CPT_AppendLine warningText, pdfWarnings
            CPT_AppendLine logText, "WARNING: " & pdfWarnings
        End If

        CPT_SetupPDFSettings pdfDoc
        pdfDoc.PublishToPDF opt.outputFolder & uniqueName & CURVES_SUFFIX & ".pdf"
        fileCount = fileCount + 1
        CPT_AppendLine logText, "PDF: " & opt.outputFolder & uniqueName & CURVES_SUFFIX & ".pdf"

        CPT_CloseTemporaryDocument pdfDoc
    End If

    Exit Sub

AllPagesFail:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    CPT_CloseTemporaryDocument normalDoc
    CPT_CloseTemporaryDocument curvesDoc
    CPT_CloseTemporaryDocument pdfDoc
    Err.Raise savedErrNumber, "CPT_ExportAllPagesTogether", savedErrDescription
End Sub

' ============================================================================
' SAVE CDR FORMATS
' ============================================================================
Private Sub CPT_SaveNormalFormats( _
    ByVal doc As Document, _
    ByVal folderPath As String, _
    ByVal baseName As String, _
    ByRef opt As CPTExportOptions, _
    ByRef fileCount As Long, _
    ByRef logText As String)

    Dim saveCurrent As New StructSaveAsOptions
    Dim saveV15 As New StructSaveAsOptions
    Dim hasLinkedSave As Boolean
    Dim path As String

    CPT_SetupSaveOptions saveCurrent, cdrCurrentVersion
    CPT_SetupSaveOptions saveV15, cdrVersion15

    If opt.exportCurrentCDR Then
        path = folderPath & baseName & ".cdr"
        doc.SaveAs path, saveCurrent
        hasLinkedSave = True
        fileCount = fileCount + 1
        CPT_AppendLine logText, "CDR: " & path
    End If

    If opt.exportV15CDR Then
        path = folderPath & baseName & V15_SUFFIX & ".cdr"
        If hasLinkedSave Then
            doc.SaveAsCopy path, saveV15
        Else
            doc.SaveAs path, saveV15
            hasLinkedSave = True
        End If
        fileCount = fileCount + 1
        CPT_AppendLine logText, "CDR v15: " & path
    End If
End Sub

Private Sub CPT_SaveCurvedFormats( _
    ByVal doc As Document, _
    ByVal folderPath As String, _
    ByVal baseName As String, _
    ByRef opt As CPTExportOptions, _
    ByRef fileCount As Long, _
    ByRef logText As String)

    Dim saveCurrent As New StructSaveAsOptions
    Dim saveV15 As New StructSaveAsOptions
    Dim hasLinkedSave As Boolean
    Dim path As String

    CPT_SetupSaveOptions saveCurrent, cdrCurrentVersion
    CPT_SetupSaveOptions saveV15, cdrVersion15

    If opt.exportCurvesCurrentCDR Then
        path = folderPath & baseName & CURVES_SUFFIX & ".cdr"
        doc.SaveAs path, saveCurrent
        hasLinkedSave = True
        fileCount = fileCount + 1
        CPT_AppendLine logText, "Curves CDR: " & path
    End If

    If opt.exportCurvesV15CDR Then
        path = folderPath & baseName & CURVES_SUFFIX & V15_SUFFIX & ".cdr"
        If hasLinkedSave Then
            doc.SaveAsCopy path, saveV15
        Else
            doc.SaveAs path, saveV15
            hasLinkedSave = True
        End If
        fileCount = fileCount + 1
        CPT_AppendLine logText, "Curves CDR v15: " & path
    End If
End Sub

Private Sub CPT_SetupSaveOptions(ByRef saveOpt As StructSaveAsOptions, ByVal fileVersion As cdrFileVersion)
    With saveOpt
        .EmbedICCProfile = False
        .EmbedVBAProject = False
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Overwrite = True
        .Range = cdrAllPages
        .Version = fileVersion
    End With
End Sub

' ============================================================================
' TEMPORARY DOCUMENT CREATION / CLEANUP
' ============================================================================
Private Function CPT_CreateSinglePageDocument(ByVal srcDoc As Document, ByVal srcPage As Page) As Document
    Dim newDoc As Document
    Dim sourcePageIndex As Long
    Dim sourceUnit As cdrUnit
    Dim sourceWidthMM As Double
    Dim sourceHeightMM As Double
    Dim i As Long

    sourcePageIndex = srcPage.Index
    sourceUnit = srcDoc.Unit
    CPT_GetPageSizeMM srcPage, sourceWidthMM, sourceHeightMM

    ' Duplicate the complete document first, then remove the other pages.
    ' This preserves the selected page's exact dimensions, orientation, object
    ' coordinates, layers, guides, and page-specific settings much more reliably
    ' than copying shapes into a newly created blank document.
    Set newDoc = CPT_DuplicateWholeDocument(srcDoc)

    For i = newDoc.Pages.Count To 1 Step -1
        If i <> sourcePageIndex Then newDoc.Pages(i).Delete
    Next i

    newDoc.Pages(1).Activate
    newDoc.Unit = cdrMillimeter
    newDoc.Pages(1).SetSize sourceWidthMM, sourceHeightMM
    newDoc.Unit = sourceUnit

    Set CPT_CreateSinglePageDocument = newDoc
End Function

Private Function CPT_DuplicateWholeDocument(ByVal srcDoc As Document) As Document
    Dim tempPath As String
    Dim saveOpt As New StructSaveAsOptions

    Randomize
    CPT_SetupSaveOptions saveOpt, cdrCurrentVersion

    tempPath = Environ$("TEMP") & "\VFE_CPT27_" & _
               Format$(Now, "yyyymmdd_hhnnss") & "_" & _
               CStr(Int((Rnd * 1000000) + 1)) & ".cdr"

    ' SaveAsCopy is intentionally used so the source document keeps its filename/link.
    srcDoc.SaveAsCopy tempPath, saveOpt
    Set CPT_DuplicateWholeDocument = OpenDocument(tempPath)

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0
End Function

Private Sub CPT_CloseTemporaryDocument(ByRef doc As Document)
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Dirty = False
        doc.Close
        Set doc = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' PAGE PREPARATION
' ============================================================================
Private Sub CPT_PrepareAllCDRPages(ByVal doc As Document, ByRef opt As CPTExportOptions)
    Dim i As Long

    For i = 1 To doc.Pages.Count
        CPT_CheckCancelled
        CPT_PrepareCDRPage doc, doc.Pages(i), opt
    Next i
End Sub

Private Sub CPT_PrepareCDRPage(ByVal doc As Document, ByVal pg As Page, ByRef opt As CPTExportOptions)
    Select Case opt.CDRPageMode
        Case cptCDRFitNoMargin
            CPT_PreparePageForExport doc, pg, 0#, opt.includeOutlines, CPT_WorkflowIncludesDimensionsInFit()
        Case cptCDRFitCustomMargin
            CPT_PreparePageForExport doc, pg, opt.marginMM, opt.includeOutlines, CPT_WorkflowIncludesDimensionsInFit()
        Case cptCDRCustomKeepPositions
            CPT_SetCustomPageSize doc, pg, opt.cdrCustomWidthMM, opt.cdrCustomHeightMM, False
        Case cptCDRCustomCenterObjects
            CPT_SetCustomPageSize doc, pg, opt.cdrCustomWidthMM, opt.cdrCustomHeightMM, True
        Case Else
            ' The page was extracted from an exact duplicate of the source
            ' document, so its original page size and object positions are kept.
    End Select
End Sub

Private Sub CPT_PrepareAllPages( _
    ByVal doc As Document, _
    ByVal marginMM As Double, _
    ByVal includeOutlines As Boolean, _
    Optional ByVal includeDimensions As Boolean = True)

    Dim i As Long

    doc.Unit = cdrMillimeter
    For i = 1 To doc.Pages.Count
        CPT_CheckCancelled
        doc.Pages(i).Activate
        CPT_PreparePageForExport doc, doc.Pages(i), marginMM, includeOutlines, includeDimensions
    Next i
End Sub

Private Sub CPT_PrepareAllPDFPages(ByVal doc As Document, ByRef opt As CPTExportOptions)
    Dim i As Long

    For i = 1 To doc.Pages.Count
        CPT_CheckCancelled
        CPT_PreparePDFPage doc, doc.Pages(i), opt
    Next i
End Sub

Private Sub CPT_PreparePDFPage(ByVal doc As Document, ByVal pg As Page, ByRef opt As CPTExportOptions)
    Select Case opt.PDFPageMode
        Case cptPDFFitNoMargin
            CPT_PreparePageForExport doc, pg, 0#, opt.includeOutlines, CPT_WorkflowIncludesDimensionsInFit()
        Case cptPDFFitCustomMargin
            CPT_PreparePageForExport doc, pg, opt.pdfMarginMM, opt.includeOutlines, CPT_WorkflowIncludesDimensionsInFit()
        Case cptPDFKeepOriginalPage
            ' Keep the source page width, height and object positions unchanged.
        Case cptPDFCustomKeepPositions
            CPT_SetCustomPageSize doc, pg, opt.pdfCustomWidthMM, opt.pdfCustomHeightMM, False
        Case cptPDFCustomCenterObjects
            CPT_SetCustomPageSize doc, pg, opt.pdfCustomWidthMM, opt.pdfCustomHeightMM, True
    End Select
End Sub

Private Sub CPT_SetCustomPageSize( _
    ByVal doc As Document, _
    ByVal pg As Page, _
    ByVal widthMM As Double, _
    ByVal heightMM As Double, _
    ByVal centerObjects As Boolean)

    Dim sr As ShapeRange
    Dim grp As Shape

    If widthMM <= 0# Or heightMM <= 0# Then
        Err.Raise vbObjectError + 2717, "CPT_SetCustomPageSize", _
                  "Custom page width and height must be greater than zero."
    End If

    doc.Unit = cdrMillimeter
    pg.Activate
    pg.SetSize widthMM, heightMM

    If centerObjects And pg.Shapes.Count > 0 Then
        Set sr = pg.Shapes.All

        If sr.Count > 1 Then
            Set grp = sr.Group
        Else
            Set grp = sr(1)
        End If

        ' Use the real CorelDRAW page center. Do not assume page origin = 0,0.
        grp.CenterX = pg.CenterX
        grp.CenterY = pg.CenterY
    End If
End Sub

Private Sub CPT_GetPageSizeMM(ByVal pg As Page, ByRef widthMM As Double, ByRef heightMM As Double)
    Dim doc As Document
    Dim oldUnit As cdrUnit

    Set doc = pg.parent.parent
    oldUnit = doc.Unit
    On Error GoTo SizeFail

    doc.Unit = cdrMillimeter
    widthMM = pg.SizeWidth
    heightMM = pg.SizeHeight
    doc.Unit = oldUnit
    Exit Sub

SizeFail:
    On Error Resume Next
    doc.Unit = oldUnit
    On Error GoTo 0
    Err.Raise vbObjectError + 2718, "CPT_GetPageSizeMM", "Could not read the source page dimensions."
End Sub

Private Sub CPT_PreparePageForExport( _
    ByVal doc As Document, _
    ByVal pg As Page, _
    ByVal marginMM As Double, _
    ByVal includeOutlines As Boolean, _
    Optional ByVal includeDimensions As Boolean = True)

    Dim sr As ShapeRange
    Dim grp As Shape
    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double
    Dim pageW As Double
    Dim pageH As Double
    Dim artworkCenterX As Double
    Dim artworkCenterY As Double
    Dim moveX As Double
    Dim moveY As Double

    doc.Unit = cdrMillimeter
    pg.Activate

    If pg.Shapes.Count = 0 Then Exit Sub
    If marginMM < 0# Then marginMM = 0#

    Set sr = pg.Shapes.All

    ' ------------------------------------------------------------
    ' FITTED GROUPED OBJECTS:
    ' 1) Group all page objects.
    ' 2) Measure the required artwork bounding box.
    ' 3) Resize the page to the artwork + optional margin.
    ' 4) Center the group on the REAL page center.
    ' ------------------------------------------------------------
    If includeDimensions Then
        If sr.Count > 1 Then
            Set grp = sr.Group
        Else
            Set grp = sr(1)
        End If

        grp.GetBoundingBox x, y, w, h, includeOutlines
    Else
        ' If dimensions are excluded from page-fit calculation, calculate the
        ' artwork-only box first, then group ALL objects so their relative
        ' positions remain unchanged while the complete page contents move.
        If Not CPT_GetPageBoundingBox(pg, includeOutlines, False, x, y, w, h) Then Exit Sub

        If sr.Count > 1 Then
            Set grp = sr.Group
        Else
            Set grp = sr(1)
        End If
    End If

    If grp Is Nothing Then Exit Sub
    If w <= 0# Or h <= 0# Then Exit Sub

    pageW = w + (marginMM * 2#)
    pageH = h + (marginMM * 2#)

    If pageW <= 0# Or pageH <= 0# Then Exit Sub

    pg.SetSize pageW, pageH

    artworkCenterX = x + (w / 2#)
    artworkCenterY = y + (h / 2#)

    moveX = pg.CenterX - artworkCenterX
    moveY = pg.CenterY - artworkCenterY

    grp.Move moveX, moveY

    ' Final correction after SetSize/Move. This protects against page-coordinate
    ' offsets and small CorelDRAW rounding differences.
    If includeDimensions Then
        grp.GetBoundingBox x, y, w, h, includeOutlines
    Else
        If Not CPT_GetPageBoundingBox(pg, includeOutlines, False, x, y, w, h) Then Exit Sub
    End If

    artworkCenterX = x + (w / 2#)
    artworkCenterY = y + (h / 2#)

    moveX = pg.CenterX - artworkCenterX
    moveY = pg.CenterY - artworkCenterY

    If Abs(moveX) > 0.0001 Or Abs(moveY) > 0.0001 Then
        grp.Move moveX, moveY
    End If
End Sub

Private Function CPT_GetPageBoundingBox( _
    ByVal pg As Page, _
    ByVal includeOutlines As Boolean, _
    ByVal includeDimensions As Boolean, _
    ByRef x As Double, _
    ByRef y As Double, _
    ByRef w As Double, _
    ByRef h As Double) As Boolean

    Dim lyr As Layer
    Dim s As Shape
    Dim sx As Double
    Dim sy As Double
    Dim sw As Double
    Dim sh As Double
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim found As Boolean
    Dim allRange As ShapeRange

    For Each lyr In pg.Layers
        If UCase$(Trim$(lyr.Name)) <> UCase$(DIMENSION_LAYER) Then
            For Each s In lyr.Shapes
                s.GetBoundingBox sx, sy, sw, sh, includeOutlines
                If sw > 0# And sh > 0# Then
                    If Not found Then
                        minX = sx: minY = sy: maxX = sx + sw: maxY = sy + sh
                        found = True
                    Else
                        If sx < minX Then minX = sx
                        If sy < minY Then minY = sy
                        If sx + sw > maxX Then maxX = sx + sw
                        If sy + sh > maxY Then maxY = sy + sh
                    End If
                End If
            Next s
        End If
    Next lyr

    If Not found Then
        Set allRange = pg.Shapes.All
        allRange.GetBoundingBox x, y, w, h, includeOutlines
        CPT_GetPageBoundingBox = (w > 0# And h > 0#)
        Exit Function
    End If

    x = minX: y = minY: w = maxX - minX: h = maxY - minY
    CPT_GetPageBoundingBox = True
End Function

' ============================================================================
' LARGE-FORMAT PDF HANDLING
' ============================================================================
Private Function CPT_HandleLargePDFDocument(ByVal doc As Document, ByVal autoScale As Boolean) As String
    Dim i As Long
    Dim pageNote As String
    Dim allNotes As String

    For i = 1 To doc.Pages.Count
        CPT_CheckCancelled
        doc.Pages(i).Activate
        pageNote = CPT_HandleLargePDFPage(doc, doc.Pages(i), doc.Pages(i).Name, autoScale)
        If Len(pageNote) > 0 Then CPT_AppendLine allNotes, pageNote
    Next i

    CPT_HandleLargePDFDocument = allNotes
End Function

Private Function CPT_HandleLargePDFPage( _
    ByVal doc As Document, _
    ByVal pg As Page, _
    ByVal displayName As String, _
    ByVal autoScale As Boolean) As String

    Dim w As Double
    Dim h As Double
    Dim maxSide As Double
    Dim scaleRatio As Double
    Dim finalW As Double
    Dim finalH As Double
    Dim note As String

    doc.Unit = cdrMillimeter
    pg.Activate

    w = pg.SizeWidth
    h = pg.SizeHeight

    If w > h Then
        maxSide = w
    Else
        maxSide = h
    End If

    If maxSide > COREL_WARNING_MM Then
        note = displayName & ": " & _
               FormatNumber(w, 2) & " x " & _
               FormatNumber(h, 2) & " mm"
    End If

    If maxSide > PDF_SAFE_MAX_MM Then
        If autoScale Then
            ' Production convention: scale the PDF-only working copy by 1:10.
            ' If 1:10 is still too large, continue 1:100, 1:1000, etc.
            scaleRatio = 1#

            Do While (maxSide / scaleRatio) > PDF_SAFE_MAX_MM
                scaleRatio = scaleRatio * PDF_SCALE_RATIO

                If scaleRatio > 1000000000# Then
                    Err.Raise vbObjectError + 2719, _
                              "CPT_HandleLargePDFPage", _
                              "Could not calculate a safe PDF scale."
                End If
            Loop

            CPT_ScalePageForPDF doc, pg, scaleRatio

            finalW = pg.SizeWidth
            finalH = pg.SizeHeight

            If Len(note) = 0 Then
                note = displayName & ": " & _
                       FormatNumber(w, 2) & " x " & _
                       FormatNumber(h, 2) & " mm"
            End If

            note = note & _
                   " -> PDF working copy scaled 1:" & CStr(scaleRatio) & _
                   " -> " & FormatNumber(finalW, 2) & _
                   " x " & FormatNumber(finalH, 2) & " mm"
        Else
            If Len(note) = 0 Then
                note = displayName & ": " & _
                       FormatNumber(w, 2) & " x " & _
                       FormatNumber(h, 2) & " mm"
            End If

            note = note & _
                   " -> exceeds recommended PDF maximum of about " & _
                   CStr(PDF_SAFE_MAX_MM) & " mm"
        End If
    ElseIf maxSide > COREL_WARNING_MM Then
        note = note & " -> above recommended large-format workflow size"
    End If

    CPT_HandleLargePDFPage = note
End Function

Private Sub CPT_ScalePageForPDF( _
    ByVal doc As Document, _
    ByVal pg As Page, _
    ByVal scaleRatio As Double)

    Dim sr As ShapeRange
    Dim grp As Shape
    Dim oldW As Double
    Dim oldH As Double
    Dim newW As Double
    Dim newH As Double
    Dim scaleFactor As Double
    Dim anchorX As Double
    Dim anchorY As Double

    If scaleRatio <= 1# Then Exit Sub
    If pg.Shapes.Count = 0 Then Exit Sub

    doc.Unit = cdrMillimeter
    pg.Activate

    oldW = pg.SizeWidth
    oldH = pg.SizeHeight

    Set sr = pg.Shapes.All

    ' Fitted modes should already contain one grouped object. Group again only
    ' when necessary so the complete PDF working page scales as one unit.
    If sr.Count > 1 Then
        Set grp = sr.Group
    Else
        Set grp = sr(1)
    End If

    If grp Is Nothing Then Exit Sub

    scaleFactor = 1# / scaleRatio

    ' Scale around the actual page center, not Document.ReferencePoint.
    anchorX = pg.CenterX
    anchorY = pg.CenterY

    grp.StretchEx anchorX, anchorY, scaleFactor, scaleFactor

    newW = oldW * scaleFactor
    newH = oldH * scaleFactor

    pg.SetSize newW, newH

    ' SetSize can change page coordinates. Re-center after resizing.
    grp.CenterX = pg.CenterX
    grp.CenterY = pg.CenterY
End Sub

Private Sub CPT_SetupPDFSettings(ByVal doc As Document)
    On Error Resume Next

    ' Publish the complete temporary document. Each-page exports already use a
    ' temporary one-page document, while all-pages exports contain all pages.
    doc.PDFSettings.PublishRange = pdfWholeDocument

    doc.PDFSettings.EmbedFonts = False
    doc.PDFSettings.TextAsCurves = False

    ' IMPORTANT:
    ' Do NOT set UsePageBoundingBox here.
    ' The previous fix enabled that property and produced blank PDFs in this
    ' workflow. The macro now physically prepares the real CorelDRAW page
    ' (grouped, fitted, centered and scaled), so PublishToPDF can use the
    ' prepared page normally.
    On Error GoTo 0
End Sub

' ============================================================================
' TEXT TO CURVES
' ============================================================================
Private Function CPT_ConvertAllTextInDocument(ByVal doc As Document) As Long
    Dim i As Long
    Dim total As Long

    For i = 1 To doc.Pages.Count
        CPT_CheckCancelled
        doc.Pages(i).Activate
        total = total + CPT_ConvertAllTextOnPage(doc.Pages(i))
    Next i

    CPT_ConvertAllTextInDocument = total
End Function

Private Function CPT_ConvertAllTextOnPage(ByVal pg As Page) As Long
    CPT_ConvertAllTextOnPage = CPT_ConvertShapesRecursive(pg.Shapes)
End Function

Private Function CPT_ConvertShapesRecursive(ByVal shapeList As Shapes) As Long
    Dim i As Long
    Dim s As Shape
    Dim pc As PowerClip
    Dim total As Long

    For i = shapeList.Count To 1 Step -1
        Set s = shapeList(i)

        If s.Type = cdrGroupShape Then
            total = total + CPT_ConvertShapesRecursive(s.Shapes)
        End If

        Set pc = Nothing
        On Error Resume Next
        Set pc = s.PowerClip
        On Error GoTo 0
        If Not pc Is Nothing Then
            total = total + CPT_ConvertShapesRecursive(pc.Shapes)
        End If

        If s.Type = cdrTextShape Then
            If Not s.Locked Then
                On Error Resume Next
                Err.Clear
                s.ConvertToCurves
                If Err.Number = 0 Then total = total + 1
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next i

    CPT_ConvertShapesRecursive = total
End Function

' ============================================================================
' PAGE AND FILE NAMING
' ============================================================================
Private Function CPT_ResolvePageName( _
    ByVal doc As Document, _
    ByVal pg As Page, _
    ByRef opt As CPTExportOptions, _
    ByVal pageIndex As Long) As String

    Dim value As String
    Dim docBase As String

    docBase = CPT_GetDocumentBaseName(doc)
    If Len(docBase) = 0 Then docBase = "Unnamed"

    Select Case opt.NameMode
        Case cptNameExportField
            value = CPT_FindTextByShapeName(pg.Shapes, NAME_SHAPE)
            If Len(Trim$(value)) = 0 Then value = docBase & "_" & Format$(pageIndex, "000")

        Case cptNameDocument
            If doc.Pages.Count = 1 Then
                value = docBase
            Else
                value = docBase & "_" & Format$(pageIndex, "000")
            End If

        Case cptNamePage
            value = Trim$(pg.Name)
            If Len(value) = 0 Then value = "Page_" & Format$(pageIndex, "000")

        Case cptNameCustom
            If doc.Pages.Count = 1 Then
                value = opt.customName
            Else
                value = opt.customName & "_" & Format$(pageIndex, "000")
            End If

        Case Else
            value = "Page_" & Format$(pageIndex, "000")
    End Select

    CPT_ResolvePageName = value
End Function

Private Function CPT_ResolveAllPagesBaseName(ByVal doc As Document, ByRef opt As CPTExportOptions) As String
    Dim value As String

    Select Case opt.NameMode
        Case cptNameCustom
            value = opt.customName
        Case cptNamePage
            value = doc.ActivePage.Name
        Case cptNameExportField
            value = CPT_FindTextByShapeName(doc.Pages(1).Shapes, NAME_SHAPE)
        Case Else
            value = CPT_GetDocumentBaseName(doc)
    End Select

    If Len(Trim$(value)) = 0 Then value = CPT_GetDocumentBaseName(doc)
    If Len(Trim$(value)) = 0 Then value = "Export"

    CPT_ResolveAllPagesBaseName = value
End Function

Private Sub CPT_RenamePagesFromExportField(ByVal doc As Document)
    Dim i As Long
    Dim value As String

    For i = 1 To doc.Pages.Count
        value = CPT_FindTextByShapeName(doc.Pages(i).Shapes, NAME_SHAPE)
        If Len(Trim$(value)) = 0 Then value = Trim$(doc.Pages(i).Name)
        If Len(Trim$(value)) = 0 Then value = "Page_" & Format$(i, "000")

        On Error Resume Next
        doc.Pages(i).Name = CPT_MakeSafeFileName(value)
        On Error GoTo 0
    Next i
End Sub

Private Function CPT_FindTextByShapeName(ByVal shapeList As Shapes, ByVal targetName As String) As String
    Dim i As Long
    Dim s As Shape
    Dim result As String
    Dim pc As PowerClip
    Dim shapeName As String

    For i = 1 To shapeList.Count
        Set s = shapeList(i)

        shapeName = ""
        On Error Resume Next
        shapeName = Trim$(s.Name)
        On Error GoTo 0

        If StrComp(shapeName, Trim$(targetName), vbTextCompare) = 0 Then
            result = CPT_GetShapeTextSafe(s)
            If Len(result) > 0 Then
                CPT_FindTextByShapeName = result
                Exit Function
            End If
        End If

        If s.Type = cdrGroupShape Then
            result = CPT_FindTextByShapeName(s.Shapes, targetName)
            If Len(result) > 0 Then
                CPT_FindTextByShapeName = result
                Exit Function
            End If
        End If

        Set pc = Nothing
        On Error Resume Next
        Set pc = s.PowerClip
        On Error GoTo 0
        If Not pc Is Nothing Then
            result = CPT_FindTextByShapeName(pc.Shapes, targetName)
            If Len(result) > 0 Then
                CPT_FindTextByShapeName = result
                Exit Function
            End If
        End If
    Next i

    CPT_FindTextByShapeName = ""
End Function

Private Function CPT_GetShapeTextSafe(ByVal s As Shape) As String
    On Error GoTo NoText
    If s.Type = cdrTextShape Then
        CPT_GetShapeTextSafe = Trim$(s.Text.Story.Text)
        Exit Function
    End If
NoText:
    CPT_GetShapeTextSafe = ""
End Function

Private Function CPT_MakeSafeFileName(ByVal rawValue As String) As String
    Dim badChars As Variant
    Dim reserved As Variant
    Dim i As Long
    Dim value As String
    Dim testUpper As String

    value = Trim$(rawValue)
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        value = Replace(value, badChars(i), "_")
    Next i

    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")

    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop

    value = Trim$(value)
    Do While Len(value) > 0 And (Right$(value, 1) = "." Or Right$(value, 1) = " ")
        value = Left$(value, Len(value) - 1)
    Loop

    If Len(value) = 0 Then value = "Unnamed"
    If Len(value) > 150 Then value = Left$(value, 150)

    reserved = Array("CON", "PRN", "AUX", "NUL", _
                     "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
                     "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    testUpper = UCase$(value)

    For i = LBound(reserved) To UBound(reserved)
        If testUpper = reserved(i) Then
            value = "_" & value
            Exit For
        End If
    Next i

    CPT_MakeSafeFileName = value
End Function

Private Function CPT_ResolveUniqueBaseName( _
    ByVal folderPath As String, _
    ByVal baseName As String, _
    ByRef opt As CPTExportOptions) As String

    Dim candidate As String
    Dim n As Long

    If Not opt.ensureUniqueNames Then
        CPT_ResolveUniqueBaseName = baseName
        Exit Function
    End If

    candidate = baseName
    n = 1

    Do While CPT_BaseNameConflicts(folderPath, candidate, opt)
        n = n + 1
        candidate = baseName & "_" & CStr(n)
    Loop

    CPT_ResolveUniqueBaseName = candidate
End Function

Private Function CPT_BaseNameConflicts( _
    ByVal folderPath As String, _
    ByVal baseName As String, _
    ByRef opt As CPTExportOptions) As Boolean

    If opt.exportCurrentCDR Then
        If CPT_FileExists(folderPath & baseName & ".cdr") Then CPT_BaseNameConflicts = True: Exit Function
    End If
    If opt.exportV15CDR Then
        If CPT_FileExists(folderPath & baseName & V15_SUFFIX & ".cdr") Then CPT_BaseNameConflicts = True: Exit Function
    End If
    If opt.exportCurvesCurrentCDR Then
        If CPT_FileExists(folderPath & baseName & CURVES_SUFFIX & ".cdr") Then CPT_BaseNameConflicts = True: Exit Function
    End If
    If opt.exportCurvesV15CDR Then
        If CPT_FileExists(folderPath & baseName & CURVES_SUFFIX & V15_SUFFIX & ".cdr") Then CPT_BaseNameConflicts = True: Exit Function
    End If
    If opt.exportCurvesPDF Then
        If CPT_FileExists(folderPath & baseName & CURVES_SUFFIX & ".pdf") Then CPT_BaseNameConflicts = True: Exit Function
    End If
End Function

' ============================================================================
' DIMENSION TOOL
' ============================================================================
Public Sub CPT_AddWidthHeightDimensions()
    Dim dimOpt As CPTDimensionOptions
    Dim specs() As CPTBoundingBox
    Dim specCount As Long
    Dim createdGroup As Shape

    If Not CPT_HasActiveDocument() Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Select one or more objects first.", vbExclamation, "Add Dimensions"
        Exit Sub
    End If

    CPT_SetDefaultDimensionOptions dimOpt
    dimOpt.mode = cptDimensionPermanentSource
    If Not CPT_PromptDimensionOptions(dimOpt, False) Then Exit Sub
    If Not CPT_CaptureSelectionSpecs(ActiveSelectionRange, dimOpt.targetMode, specs, specCount) Then Exit Sub

    If dimOpt.RemoveOldDimensions Then CPT_RemoveDimensionLayerContents ActiveDocument.ActivePage
    Set createdGroup = CPT_CreateDimensionsFromSpecs(ActiveDocument.ActivePage, dimOpt, specs, specCount)

    If createdGroup Is Nothing Then
        MsgBox "No dimensions were created.", vbExclamation, "VFE Corel Production Tools 27"
    Else
        ActiveWindow.Refresh
        MsgBox "Dimensions added on layer '" & DIMENSION_LAYER & "'." & vbCrLf & _
               "Targets measured: " & CStr(specCount), vbInformation, "VFE Corel Production Tools 27"
    End If
End Sub

Private Sub CPT_SetDefaultDimensionOptions(ByRef dimOpt As CPTDimensionOptions)
    dimOpt.Enabled = True
    dimOpt.targetMode = cptDimensionEachObject
    dimOpt.Axes = cptDimensionWidthAndHeight
    dimOpt.mode = cptDimensionExportCopiesOnly
    dimOpt.unitText = "cm"
    dimOpt.decimals = 2
    dimOpt.fontName = "Arial"
    dimOpt.fontSizePt = 10#
    dimOpt.lineOffsetMM = 5#
    dimOpt.IncludeInPageFit = True
    dimOpt.RemoveOldDimensions = False
    dimOpt.groupCreated = True
    dimOpt.ColorR = 0
    dimOpt.ColorG = 0
    dimOpt.ColorB = 255
End Sub

Private Function CPT_PromptDimensionOptions(ByRef dimOpt As CPTDimensionOptions, ByVal askMode As Boolean) As Boolean
    Dim value As String
    Dim ans As VbMsgBoxResult

    value = Trim$(InputBox( _
        "Dimension target:" & vbCrLf & _
        "1 = Complete selection as one measurement" & vbCrLf & _
        "2 = Every selected object separately", _
        "Dimension Options", CStr(dimOpt.targetMode)))
    If Len(value) = 0 Then Exit Function
    If value = "1" Then
        dimOpt.targetMode = cptDimensionCompleteSelection
    ElseIf value = "2" Then
        dimOpt.targetMode = cptDimensionEachObject
    Else
        MsgBox "Invalid dimension target.", vbExclamation
        Exit Function
    End If

    value = Trim$(InputBox( _
        "Measurements:" & vbCrLf & _
        "1 = Width only" & vbCrLf & _
        "2 = Height only" & vbCrLf & _
        "3 = Width and height", _
        "Dimension Options", CStr(dimOpt.Axes)))
    If Len(value) = 0 Then Exit Function
    Select Case value
        Case "1": dimOpt.Axes = cptDimensionWidthOnly
        Case "2": dimOpt.Axes = cptDimensionHeightOnly
        Case "3": dimOpt.Axes = cptDimensionWidthAndHeight
        Case Else: MsgBox "Invalid measurement option.", vbExclamation: Exit Function
    End Select

    value = LCase$(Trim$(InputBox("Display unit: mm, cm, or in", "Dimension Options", dimOpt.unitText)))
    If Len(value) = 0 Then Exit Function
    If value <> "mm" And value <> "cm" And value <> "in" Then
        MsgBox "Unit must be mm, cm, or in.", vbExclamation
        Exit Function
    End If
    dimOpt.unitText = value

    value = Trim$(InputBox("Number of decimal places (0 to 4):", "Dimension Options", CStr(dimOpt.decimals)))
    If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
    dimOpt.decimals = CLng(value)
    If dimOpt.decimals < 0 Then dimOpt.decimals = 0
    If dimOpt.decimals > 4 Then dimOpt.decimals = 4

    dimOpt.fontName = Trim$(InputBox("Dimension font name:", "Dimension Options", dimOpt.fontName))
    If Len(dimOpt.fontName) = 0 Then Exit Function

    value = Trim$(InputBox("Dimension text size in points:", "Dimension Options", CStr(dimOpt.fontSizePt)))
    If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
    dimOpt.fontSizePt = CDbl(value)
    If dimOpt.fontSizePt < 1# Then dimOpt.fontSizePt = 10#

    value = Trim$(InputBox("Object-to-dimension-line distance in millimetres:", "Dimension Options", CStr(dimOpt.lineOffsetMM)))
    If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
    dimOpt.lineOffsetMM = CDbl(value)
    If dimOpt.lineOffsetMM <= 0# Then dimOpt.lineOffsetMM = 5#

    ans = MsgBox("Include generated dimensions when calculating fitted page size?", vbYesNoCancel + vbQuestion, "Dimension Options")
    If ans = vbCancel Then Exit Function
    dimOpt.IncludeInPageFit = (ans = vbYes)

    ans = MsgBox("Remove existing objects from the DIMENSIONS layer before adding new dimensions?", vbYesNoCancel + vbQuestion, "Dimension Options")
    If ans = vbCancel Then Exit Function
    dimOpt.RemoveOldDimensions = (ans = vbYes)

    If askMode Then
        value = Trim$(InputBox( _
            "Dimension placement:" & vbCrLf & _
            "1 = Add permanently to source document" & vbCrLf & _
            "2 = Add temporarily, export, then remove" & vbCrLf & _
            "3 = Add only inside export copies", _
            "Dimension Options", CStr(dimOpt.mode)))
        If Len(value) = 0 Then Exit Function
        Select Case value
            Case "1": dimOpt.mode = cptDimensionPermanentSource
            Case "2": dimOpt.mode = cptDimensionTemporarySource
            Case "3": dimOpt.mode = cptDimensionExportCopiesOnly
            Case Else: MsgBox "Invalid dimension placement.", vbExclamation: Exit Function
        End Select
    End If

    CPT_PromptDimensionOptions = True
End Function

Private Function CPT_CaptureSelectionSpecs( _
    ByVal sr As ShapeRange, _
    ByVal targetMode As CPTDimensionTarget, _
    ByRef specs() As CPTBoundingBox, _
    ByRef specCount As Long) As Boolean

    Dim doc As Document
    Dim oldUnit As cdrUnit
    Dim i As Long
    Dim savedErrDescription As String

    If sr Is Nothing Then Exit Function
    If sr.Count = 0 Then Exit Function

    Set doc = ActiveDocument
    oldUnit = doc.Unit
    On Error GoTo CaptureFail
    doc.Unit = cdrMillimeter

    If targetMode = cptDimensionCompleteSelection Then
        specCount = 1
        ReDim specs(1 To 1)
        sr.GetBoundingBox specs(1).x, specs(1).y, specs(1).w, specs(1).h, False
    Else
        specCount = sr.Count
        ReDim specs(1 To specCount)
        For i = 1 To specCount
            sr(i).GetBoundingBox specs(i).x, specs(i).y, specs(i).w, specs(i).h, False
        Next i
    End If

    doc.Unit = oldUnit
    CPT_CaptureSelectionSpecs = True
    Exit Function

CaptureFail:
    savedErrDescription = Err.Description
    On Error Resume Next
    doc.Unit = oldUnit
    On Error GoTo 0
    MsgBox "Could not read the selected object sizes:" & vbCrLf & savedErrDescription, vbCritical, "VFE Corel Production Tools 27"
End Function

Private Function CPT_CreateDimensionsFromSpecs( _
    ByVal pg As Page, _
    ByRef dimOpt As CPTDimensionOptions, _
    ByRef specs() As CPTBoundingBox, _
    ByVal specCount As Long) As Shape

    Dim doc As Document
    Dim oldUnit As cdrUnit
    Dim lyr As Layer
    Dim i As Long
    Dim prefix As String
    Dim master As Shape
    Dim commandStarted As Boolean
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    If specCount <= 0 Then Exit Function
    ' Page.Parent is the Pages collection; Pages.Parent is the owning Document.
    ' Assigning Page.Parent directly to Document raises VBA Run-time error 13.
    Set doc = pg.parent.parent
    oldUnit = doc.Unit

    On Error GoTo DimensionFail
    doc.BeginCommandGroup "VFE CPT27 Create Dimensions"
    commandStarted = True
    doc.Unit = cdrMillimeter
    pg.Activate

    Set lyr = CPT_GetOrCreateDimensionLayer(pg)
    prefix = "CPT_DIM_" & Format$(Now, "hhnnss") & "_" & CStr(Int(Timer * 100#)) & "_"

    doc.ClearSelection
    For i = 1 To specCount
        CPT_CreateDimensionsForBox lyr, dimOpt, specs(i), prefix & CStr(i)
    Next i

    If ActiveSelectionRange.Count > 0 Then
        If dimOpt.groupCreated And ActiveSelectionRange.Count > 1 Then
            Set master = ActiveSelectionRange.Group
            master.Name = prefix & "GROUP"
            master.CreateSelection
        Else
            Set master = ActiveSelectionRange(1)
        End If
    End If

    doc.Unit = oldUnit
    doc.EndCommandGroup
    commandStarted = False
    Set CPT_CreateDimensionsFromSpecs = master
    Exit Function

DimensionFail:
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    doc.Unit = oldUnit
    If commandStarted Then doc.EndCommandGroup
    On Error GoTo 0
    Err.Raise savedErrNumber, "CPT_CreateDimensionsFromSpecs", savedErrDescription
End Function

Private Sub CPT_CreateDimensionsForBox( _
    ByVal lyr As Layer, _
    ByRef dimOpt As CPTDimensionOptions, _
    ByRef box As CPTBoundingBox, _
    ByVal shapePrefix As String)

    Dim lineOffsetMM As Double
    Dim textOffsetMM As Double
    Dim extensionMM As Double
    Dim dimY As Double
    Dim txtY As Double
    Dim dimX As Double
    Dim txtX As Double
    Dim value As Double
    Dim s As Shape

    lineOffsetMM = dimOpt.lineOffsetMM
    textOffsetMM = lineOffsetMM * 2#
    extensionMM = 1.5
    dimY = box.y - lineOffsetMM
    txtY = box.y - textOffsetMM
    dimX = box.x - lineOffsetMM
    txtX = box.x - textOffsetMM

    If dimOpt.Axes = cptDimensionWidthOnly Or dimOpt.Axes = cptDimensionWidthAndHeight Then
        Set s = lyr.CreateLineSegment(box.x, box.y, box.x, dimY - extensionMM)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_W1", False, dimOpt
        Set s = lyr.CreateLineSegment(box.x + box.w, box.y, box.x + box.w, dimY - extensionMM)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_W2", False, dimOpt
        Set s = lyr.CreateLineSegment(box.x, dimY, box.x + box.w, dimY)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_WL", False, dimOpt

        value = CPT_ConvertMMToDisplayUnit(box.w, dimOpt.unitText)
        Set s = lyr.CreateArtisticText(box.x + (box.w / 2#), txtY, CPT_FormatMeasurement(value, dimOpt.decimals) & " " & dimOpt.unitText)
        s.CenterX = box.x + (box.w / 2#)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_WT", True, dimOpt
    End If

    If dimOpt.Axes = cptDimensionHeightOnly Or dimOpt.Axes = cptDimensionWidthAndHeight Then
        Set s = lyr.CreateLineSegment(box.x, box.y, dimX - extensionMM, box.y)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_H1", False, dimOpt
        Set s = lyr.CreateLineSegment(box.x, box.y + box.h, dimX - extensionMM, box.y + box.h)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_H2", False, dimOpt
        Set s = lyr.CreateLineSegment(dimX, box.y, dimX, box.y + box.h)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_HL", False, dimOpt

        value = CPT_ConvertMMToDisplayUnit(box.h, dimOpt.unitText)
        Set s = lyr.CreateArtisticText(txtX, box.y + (box.h / 2#), CPT_FormatMeasurement(value, dimOpt.decimals) & " " & dimOpt.unitText)
        s.RotationAngle = 90#
        s.CenterY = box.y + (box.h / 2#)
        CPT_NameAndSelectDimensionShape s, shapePrefix & "_HT", True, dimOpt
    End If
End Sub

Private Sub CPT_NameAndSelectDimensionShape( _
    ByVal s As Shape, _
    ByVal shapeName As String, _
    ByVal isText As Boolean, _
    ByRef dimOpt As CPTDimensionOptions)

    s.Name = shapeName
    If isText Then
        CPT_SetDimensionTextStyle s, dimOpt
    Else
        CPT_SetDimensionLineStyle s, dimOpt
    End If
    s.AddToSelection
End Sub

Private Function CPT_GetOrCreateDimensionLayer(ByVal pg As Page) As Layer
    Dim lyr As Layer

    On Error Resume Next
    Set lyr = pg.Layers.Find(DIMENSION_LAYER)
    On Error GoTo 0
    If lyr Is Nothing Then Set lyr = pg.CreateLayer(DIMENSION_LAYER)
    Set CPT_GetOrCreateDimensionLayer = lyr
End Function

Private Sub CPT_RemoveDimensionLayerContents(ByVal pg As Page)
    Dim lyr As Layer

    On Error Resume Next
    Set lyr = pg.Layers.Find(DIMENSION_LAYER)
    If Not lyr Is Nothing Then
        If lyr.Shapes.Count > 0 Then lyr.Shapes.All.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub CPT_SetDimensionLineStyle(ByVal s As Shape, ByRef dimOpt As CPTDimensionOptions)
    On Error Resume Next
    s.Fill.ApplyNoFill
    s.Outline.Width = 0
    s.Outline.Color.RGBAssign dimOpt.ColorR, dimOpt.ColorG, dimOpt.ColorB
    On Error GoTo 0
End Sub

Private Sub CPT_SetDimensionTextStyle(ByVal s As Shape, ByRef dimOpt As CPTDimensionOptions)
    On Error Resume Next
    s.Outline.Width = 0
    s.Text.Story.Size = dimOpt.fontSizePt
    s.Text.Story.Font = dimOpt.fontName
    If Err.Number <> 0 Then
        Err.Clear
        s.Text.Story.Font = "Arial"
    End If
    s.Text.Story.Fill.UniformColor.RGBAssign dimOpt.ColorR, dimOpt.ColorG, dimOpt.ColorB
    On Error GoTo 0
End Sub

Private Function CPT_ConvertMMToDisplayUnit(ByVal valueMM As Double, ByVal unitText As String) As Double
    Select Case LCase$(unitText)
        Case "cm": CPT_ConvertMMToDisplayUnit = valueMM / 10#
        Case "in": CPT_ConvertMMToDisplayUnit = valueMM / 25.4
        Case Else: CPT_ConvertMMToDisplayUnit = valueMM
    End Select
End Function

Private Function CPT_FormatMeasurement(ByVal value As Double, ByVal decimals As Long) As String
    If decimals <= 0 Then
        CPT_FormatMeasurement = Format$(value, "0")
    Else
        CPT_FormatMeasurement = Format$(value, "0." & String$(decimals, "0"))
    End If
End Function

' ============================================================================
' WORKFLOW PROFILES
' ============================================================================
Public Sub CPT_RunWorkflowProfile()
    Dim profile As CPTWorkflowProfile
    Dim value As String
    Dim names As String

    If Not CPT_HasActiveDocument() Then Exit Sub

    names = CPT_GetProfileNames()
    value = Trim$(InputBox( _
        "Choose a workflow profile:" & vbCrLf & _
        "1 = Dimensions + 5 Formats" & vbCrLf & _
        "2 = Dimensions + PDF Only" & vbCrLf & _
        "3 = Production Export" & vbCrLf & _
        "4 = Customer Preview" & vbCrLf & _
        "5 = All Pages Archive" & vbCrLf & _
        "6 = Saved custom profile" & vbCrLf & _
        "7 = Build, save and run a new custom profile", _
        "Run Workflow Profile", "1"))
    If Len(value) = 0 Then Exit Sub

    Select Case value
        Case "1", "2", "3", "4", "5"
            CPT_LoadBuiltInProfile CLng(value), profile
        Case "6"
            If Len(names) = 0 Then
                MsgBox "No custom profiles are saved yet.", vbInformation, "Workflow Profiles"
                Exit Sub
            End If
            If Not CPT_ChooseSavedProfile(profile, "Run Saved Profile") Then Exit Sub
        Case "7"
            If Not CPT_BuildCustomProfile(profile) Then Exit Sub
            CPT_SaveWorkflowProfile profile
        Case Else
            MsgBox "Choose a number from 1 to 7.", vbExclamation, "Workflow Profiles"
            Exit Sub
    End Select

    CPT_ExecuteWorkflowProfile profile
End Sub

Public Sub CPT_ManageWorkflowProfiles()
    Dim value As String
    Dim profile As CPTWorkflowProfile
    Dim nameToDelete As String
    Dim ans As VbMsgBoxResult

    value = Trim$(InputBox( _
        "Manage workflow profiles:" & vbCrLf & _
        "1 = Create or update a custom profile" & vbCrLf & _
        "2 = View a saved profile" & vbCrLf & _
        "3 = Delete a saved profile" & vbCrLf & _
        "4 = Delete all saved custom profiles", _
        "Manage Workflow Profiles", "1"))
    If Len(value) = 0 Then Exit Sub

    Select Case value
        Case "1"
            If Not CPT_BuildCustomProfile(profile) Then Exit Sub
            CPT_SaveWorkflowProfile profile
            MsgBox "Profile saved: " & profile.Name, vbInformation, "Workflow Profiles"
        Case "2"
            If Not CPT_ChooseSavedProfile(profile, "View Saved Profile") Then Exit Sub
            MsgBox CPT_ProfileSummary(profile), vbInformation, "Profile: " & profile.Name
        Case "3"
            If Not CPT_ChooseSavedProfile(profile, "Delete Saved Profile") Then Exit Sub
            nameToDelete = profile.Name
            ans = MsgBox("Delete profile '" & nameToDelete & "'?", vbYesNo + vbQuestion, "Delete Profile")
            If ans = vbYes Then
                CPT_DeleteWorkflowProfile nameToDelete
                MsgBox "Profile deleted.", vbInformation, "Workflow Profiles"
            End If
        Case "4"
            ans = MsgBox("Delete every saved custom profile?", vbYesNo + vbExclamation, "Delete Profiles")
            If ans = vbYes Then
                CPT_DeleteAllWorkflowProfiles
                MsgBox "All custom profiles were deleted.", vbInformation, "Workflow Profiles"
            End If
        Case Else
            MsgBox "Choose a number from 1 to 4.", vbExclamation, "Workflow Profiles"
    End Select
End Sub

Private Sub CPT_LoadBuiltInProfile(ByVal profileNumber As Long, ByRef profile As CPTWorkflowProfile)
    CPT_SetDefaultOptions profile.ExportOptions
    CPT_SetDefaultDimensionOptions profile.DimensionOptions
    profile.BuiltIn = True

    Select Case profileNumber
        Case 1
            profile.Name = "Dimensions + 5 Formats"
            profile.DimensionOptions.Enabled = True
            profile.DimensionOptions.targetMode = cptDimensionEachObject
            profile.DimensionOptions.Axes = cptDimensionWidthAndHeight
            profile.DimensionOptions.mode = cptDimensionExportCopiesOnly
            profile.DimensionOptions.IncludeInPageFit = True

            With profile.ExportOptions
                .Scope = cptScopeEachPage
                .NameMode = cptNameExportField
                .exportCurrentCDR = True
                .exportV15CDR = True
                .exportCurvesCurrentCDR = True
                .exportCurvesV15CDR = True
                .exportCurvesPDF = True
                .FitPageToArtwork = True
                .CDRPageMode = cptCDRFitCustomMargin
                .marginMM = DEFAULT_MARGIN_MM
                .PDFPageMode = cptPDFFitCustomMargin
                .pdfMarginMM = DEFAULT_MARGIN_MM
            End With

        Case 2
            profile.Name = "Dimensions + PDF Only"
            profile.DimensionOptions.Enabled = True
            profile.DimensionOptions.targetMode = cptDimensionEachObject
            profile.DimensionOptions.Axes = cptDimensionWidthAndHeight
            profile.DimensionOptions.mode = cptDimensionExportCopiesOnly
            profile.DimensionOptions.IncludeInPageFit = True

            With profile.ExportOptions
                .Scope = cptScopeCurrentPage
                .NameMode = cptNameExportField
                .exportCurvesPDF = True
                .PDFPageMode = cptPDFFitNoMargin
            End With

        Case 3
            profile.Name = "Production Export"
            profile.DimensionOptions.Enabled = False
            With profile.ExportOptions
                .Scope = cptScopeEachPage
                .NameMode = cptNameExportField
                .exportCurrentCDR = True
                .exportV15CDR = True
                .exportCurvesCurrentCDR = True
                .exportCurvesV15CDR = True
                .exportCurvesPDF = True
                .FitPageToArtwork = True
                .CDRPageMode = cptCDRFitCustomMargin
                .marginMM = DEFAULT_MARGIN_MM
                .PDFPageMode = cptPDFFitCustomMargin
                .pdfMarginMM = DEFAULT_MARGIN_MM
            End With

        Case 4
            profile.Name = "Customer Preview"
            profile.DimensionOptions.Enabled = False
            With profile.ExportOptions
                .Scope = cptScopeCurrentPage
                .NameMode = cptNameExportField
                .exportCurvesPDF = True
                .PDFPageMode = cptPDFFitNoMargin
            End With

        Case 5
            profile.Name = "All Pages Archive"
            profile.DimensionOptions.Enabled = False
            With profile.ExportOptions
                .Scope = cptScopeAllPages
                .NameMode = cptNameDocument
                .exportCurrentCDR = True
                .exportV15CDR = True
                .exportCurvesCurrentCDR = True
                .exportCurvesV15CDR = True
                .exportCurvesPDF = True
                .FitPageToArtwork = False
                .CDRPageMode = cptCDRKeepOriginalPage
                .PDFPageMode = cptPDFKeepOriginalPage
            End With
    End Select
End Sub

Private Function CPT_BuildCustomProfile(ByRef profile As CPTWorkflowProfile) As Boolean
    Dim ans As VbMsgBoxResult

    CPT_SetDefaultOptions profile.ExportOptions
    CPT_SetDefaultDimensionOptions profile.DimensionOptions
    profile.BuiltIn = False

    profile.Name = Trim$(InputBox("Profile name:", "Custom Workflow Profile", "My Production Profile"))
    If Len(profile.Name) = 0 Then Exit Function
    profile.Name = Replace(profile.Name, "|", "-")

    ans = MsgBox("Should this profile add dimensions before export?", vbYesNoCancel + vbQuestion, "Custom Workflow Profile")
    If ans = vbCancel Then Exit Function
    profile.DimensionOptions.Enabled = (ans = vbYes)

    If profile.DimensionOptions.Enabled Then
        If Not CPT_PromptDimensionOptions(profile.DimensionOptions, True) Then Exit Function
    End If

    ' Start custom profiles with the common five-format production selection.
    profile.ExportOptions.exportCurrentCDR = True
    profile.ExportOptions.exportV15CDR = True
    profile.ExportOptions.exportCurvesCurrentCDR = True
    profile.ExportOptions.exportCurvesV15CDR = True
    profile.ExportOptions.exportCurvesPDF = True
    profile.ExportOptions.FitPageToArtwork = True
    profile.ExportOptions.CDRPageMode = cptCDRFitCustomMargin
    profile.ExportOptions.PDFPageMode = cptPDFFitCustomMargin

    If Not CPT_PromptExportOptions(profile.ExportOptions, "Profile: " & profile.Name) Then Exit Function
    CPT_BuildCustomProfile = True
End Function

Private Sub CPT_ExecuteWorkflowProfile(ByRef profile As CPTWorkflowProfile)
    Dim specs() As CPTBoundingBox
    Dim specCount As Long
    Dim originalSelection As ShapeRange
    Dim sourceDimensionGroup As Shape
    Dim sourcePage As Page
    Dim ans As VbMsgBoxResult
    Dim workflowError As String

    If Not CPT_HasActiveDocument() Then Exit Sub
    On Error GoTo WorkflowFail

    If profile.DimensionOptions.Enabled Then
        If ActiveSelectionRange.Count = 0 Then
            MsgBox "Select the objects to dimension before running this profile.", vbExclamation, "Workflow Profile"
            Exit Sub
        End If
        Set originalSelection = ActiveSelectionRange
        Set sourcePage = ActiveDocument.ActivePage
        If Not CPT_CaptureSelectionSpecs(originalSelection, profile.DimensionOptions.targetMode, specs, specCount) Then Exit Sub
    End If

    If Len(profile.ExportOptions.outputFolder) = 0 Or Not CPT_FolderExists(CPT_NormalizeFolder(profile.ExportOptions.outputFolder)) Then
        profile.ExportOptions.outputFolder = CPT_GetFolderFromUser("Choose output folder for " & profile.Name)
        If Len(profile.ExportOptions.outputFolder) = 0 Then Exit Sub
    End If

    ans = MsgBox(CPT_ProfileSummary(profile) & vbCrLf & vbCrLf & "Run this workflow now?", _
                 vbYesNo + vbQuestion, "Workflow Profile")
    If ans <> vbYes Then Exit Sub

    CPT_ClearWorkflowContext
    gWorkflowActive = True
    gWorkflowName = profile.Name
    CPT_CopyDimensionOptions profile.DimensionOptions, gWorkflowDimensionOptions

    If profile.DimensionOptions.Enabled Then
        gWorkflowDimensionSourcePage = sourcePage.Index
        CPT_CopySpecsToWorkflow specs, specCount

        Select Case profile.DimensionOptions.mode
            Case cptDimensionPermanentSource
                If profile.DimensionOptions.RemoveOldDimensions Then CPT_RemoveDimensionLayerContents sourcePage
                Set sourceDimensionGroup = CPT_CreateDimensionsFromSpecs(sourcePage, profile.DimensionOptions, specs, specCount)

            Case cptDimensionTemporarySource
                ' Existing dimensions are preserved so the source can be restored safely.
                gWorkflowDimensionOptions.RemoveOldDimensions = False
                Set sourceDimensionGroup = CPT_CreateDimensionsFromSpecs(sourcePage, profile.DimensionOptions, specs, specCount)

            Case cptDimensionExportCopiesOnly
                ' Dimensions are created independently in every temporary export document.
        End Select
    End If

    CPT_RunExport profile.ExportOptions
    GoTo WorkflowCleanup

WorkflowFail:
    workflowError = Err.Description

WorkflowCleanup:
    On Error Resume Next
    If profile.DimensionOptions.Enabled Then
        If profile.DimensionOptions.mode = cptDimensionTemporarySource Then
            If Not sourceDimensionGroup Is Nothing Then sourceDimensionGroup.Delete
        End If
        If Not originalSelection Is Nothing Then originalSelection.CreateSelection
    End If
    CPT_ClearWorkflowContext
    ActiveWindow.Refresh
    On Error GoTo 0
    If Len(workflowError) > 0 Then
        MsgBox "Workflow stopped:" & vbCrLf & workflowError, vbCritical, "VFE Corel Production Tools 27"
    End If
End Sub

Private Sub CPT_CopyDimensionOptions(ByRef source As CPTDimensionOptions, ByRef target As CPTDimensionOptions)
    target.Enabled = source.Enabled
    target.targetMode = source.targetMode
    target.Axes = source.Axes
    target.mode = source.mode
    target.unitText = source.unitText
    target.decimals = source.decimals
    target.fontName = source.fontName
    target.fontSizePt = source.fontSizePt
    target.lineOffsetMM = source.lineOffsetMM
    target.IncludeInPageFit = source.IncludeInPageFit
    target.RemoveOldDimensions = source.RemoveOldDimensions
End Sub

Private Sub CPT_CopySpecsToWorkflow(ByRef specs() As CPTBoundingBox, ByVal specCount As Long)
    Dim i As Long

    gWorkflowDimensionSpecCount = specCount
    If specCount <= 0 Then Exit Sub
    ReDim gWorkflowDimensionSpecs(1 To specCount)
    For i = 1 To specCount
        gWorkflowDimensionSpecs(i).x = specs(i).x
        gWorkflowDimensionSpecs(i).y = specs(i).y
        gWorkflowDimensionSpecs(i).w = specs(i).w
        gWorkflowDimensionSpecs(i).h = specs(i).h
    Next i
End Sub

Private Sub CPT_ClearWorkflowContext()
    gWorkflowActive = False
    gWorkflowName = ""
    gWorkflowDimensionSpecCount = 0
    Erase gWorkflowDimensionSpecs
    CPT_SetDefaultDimensionOptions gWorkflowDimensionOptions
    gWorkflowDimensionOptions.Enabled = False
    gWorkflowDimensionSourcePage = 0
End Sub

Private Function CPT_WorkflowIncludesDimensionsInFit() As Boolean
    If gUIExportActive Then
        CPT_WorkflowIncludesDimensionsInFit = gUIIncludeDimensionsInFit
    ElseIf gWorkflowActive And gWorkflowDimensionOptions.Enabled Then
        CPT_WorkflowIncludesDimensionsInFit = gWorkflowDimensionOptions.IncludeInPageFit
    Else
        CPT_WorkflowIncludesDimensionsInFit = True
    End If
End Function

Private Sub CPT_ApplyWorkflowDimensionsToTempPage(ByVal pg As Page, ByVal sourcePageIndex As Long)
    Dim created As Shape

    If Not gWorkflowActive Then Exit Sub
    If Not gWorkflowDimensionOptions.Enabled Then Exit Sub
    If gWorkflowDimensionOptions.mode <> cptDimensionExportCopiesOnly Then Exit Sub
    If sourcePageIndex <> gWorkflowDimensionSourcePage Then Exit Sub
    If gWorkflowDimensionSpecCount <= 0 Then Exit Sub

    If gWorkflowDimensionOptions.RemoveOldDimensions Then CPT_RemoveDimensionLayerContents pg
    Set created = CPT_CreateDimensionsFromSpecs(pg, gWorkflowDimensionOptions, gWorkflowDimensionSpecs, gWorkflowDimensionSpecCount)
End Sub

Private Sub CPT_ApplyWorkflowDimensionsToDocument(ByVal doc As Document)
    If Not gWorkflowActive Then Exit Sub
    If Not gWorkflowDimensionOptions.Enabled Then Exit Sub
    If gWorkflowDimensionOptions.mode <> cptDimensionExportCopiesOnly Then Exit Sub
    If gWorkflowDimensionSourcePage < 1 Or gWorkflowDimensionSourcePage > doc.Pages.Count Then Exit Sub

    CPT_ApplyWorkflowDimensionsToTempPage doc.Pages(gWorkflowDimensionSourcePage), gWorkflowDimensionSourcePage
End Sub

' ============================================================================
' CUSTOM EXPORT PROMPTS
' ============================================================================
Private Function CPT_PromptExportOptions(ByRef opt As CPTExportOptions, ByVal dialogTitle As String) As Boolean
    Dim value As String
    Dim ans As VbMsgBoxResult

    value = Trim$(InputBox( _
        "Export scope:" & vbCrLf & _
        "1 = Each page as separate files" & vbCrLf & _
        "2 = All pages in one document" & vbCrLf & _
        "3 = Current page only", _
        dialogTitle & " - Scope", CStr(opt.Scope)))
    If Len(value) = 0 Then Exit Function
    Select Case value
        Case "1": opt.Scope = cptScopeEachPage
        Case "2": opt.Scope = cptScopeAllPages
        Case "3": opt.Scope = cptScopeCurrentPage
        Case Else: MsgBox "Invalid export scope.", vbExclamation: Exit Function
    End Select

    value = Trim$(InputBox( _
        "Filename source:" & vbCrLf & _
        "1 = Text object named EXPORT_NAME" & vbCrLf & _
        "2 = Document filename" & vbCrLf & _
        "3 = Page name" & vbCrLf & _
        "4 = Custom prefix", _
        dialogTitle & " - Names", CStr(opt.NameMode)))
    If Len(value) = 0 Then Exit Function
    Select Case value
        Case "1": opt.NameMode = cptNameExportField
        Case "2": opt.NameMode = cptNameDocument
        Case "3": opt.NameMode = cptNamePage
        Case "4"
            opt.NameMode = cptNameCustom
            opt.customName = Trim$(InputBox("Enter the custom filename or prefix:", dialogTitle & " - Prefix", opt.customName))
            If Len(opt.customName) = 0 Then Exit Function
        Case Else: MsgBox "Invalid filename source.", vbExclamation: Exit Function
    End Select

    opt.exportCurrentCDR = (MsgBox("Export normal current-version CDR?", vbYesNo + vbQuestion, dialogTitle & " - Format 1") = vbYes)
    opt.exportV15CDR = (MsgBox("Export normal version-15 CDR?", vbYesNo + vbQuestion, dialogTitle & " - Format 2") = vbYes)
    opt.exportCurvesCurrentCDR = (MsgBox("Export curved current-version CDR?", vbYesNo + vbQuestion, dialogTitle & " - Format 3") = vbYes)
    opt.exportCurvesV15CDR = (MsgBox("Export curved version-15 CDR?", vbYesNo + vbQuestion, dialogTitle & " - Format 4") = vbYes)
    opt.exportCurvesPDF = (MsgBox("Export curved PDF?", vbYesNo + vbQuestion, dialogTitle & " - Format 5") = vbYes)

    If Not CPT_AnyFormatSelected(opt) Then
        MsgBox "No output format was selected.", vbExclamation
        Exit Function
    End If

    If opt.exportCurrentCDR Or opt.exportV15CDR Or opt.exportCurvesCurrentCDR Or opt.exportCurvesV15CDR Then
        value = Trim$(InputBox( _
            "CDR page mode:" & vbCrLf & _
            "1 = Preserve source page exactly" & vbCrLf & _
            "2 = Fit grouped objects, no margin" & vbCrLf & _
            "3 = Fit grouped objects, custom margin" & vbCrLf & _
            "4 = Custom page size, keep object positions" & vbCrLf & _
            "5 = Custom page size, centre objects", _
            dialogTitle & " - CDR Page", CStr(opt.CDRPageMode)))
        If Len(value) = 0 Then Exit Function
        Select Case value
            Case "1": opt.CDRPageMode = cptCDRKeepOriginalPage
            Case "2": opt.CDRPageMode = cptCDRFitNoMargin: opt.marginMM = 0#
            Case "3"
                opt.CDRPageMode = cptCDRFitCustomMargin
                value = Trim$(InputBox("CDR margin on all four sides in millimetres:", dialogTitle & " - CDR Margin", CStr(opt.marginMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.marginMM = CDbl(value)
                If opt.marginMM < 0# Then opt.marginMM = 0#
            Case "4", "5"
                If value = "4" Then opt.CDRPageMode = cptCDRCustomKeepPositions Else opt.CDRPageMode = cptCDRCustomCenterObjects
                value = Trim$(InputBox("Custom CDR page width in millimetres:", dialogTitle & " - CDR Width", CStr(opt.cdrCustomWidthMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.cdrCustomWidthMM = CDbl(value)
                value = Trim$(InputBox("Custom CDR page height in millimetres:", dialogTitle & " - CDR Height", CStr(opt.cdrCustomHeightMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.cdrCustomHeightMM = CDbl(value)
                If opt.cdrCustomWidthMM <= 0# Or opt.cdrCustomHeightMM <= 0# Then Exit Function
            Case Else: MsgBox "Invalid CDR page mode.", vbExclamation: Exit Function
        End Select
        opt.FitPageToArtwork = (opt.CDRPageMode = cptCDRFitNoMargin Or opt.CDRPageMode = cptCDRFitCustomMargin)
    Else
        opt.FitPageToArtwork = False
        opt.CDRPageMode = cptCDRKeepOriginalPage
    End If

    If opt.exportCurvesPDF Then
        value = Trim$(InputBox( _
            "PDF page mode:" & vbCrLf & _
            "1 = Group all objects and fit page, no margin" & vbCrLf & _
            "2 = Group all objects and fit page, custom margin" & vbCrLf & _
            "3 = Preserve source page exactly" & vbCrLf & _
            "4 = Custom page size, keep object positions" & vbCrLf & _
            "5 = Custom page size, centre objects", _
            dialogTitle & " - PDF Page", CStr(opt.PDFPageMode)))
        If Len(value) = 0 Then Exit Function
        Select Case value
            Case "1": opt.PDFPageMode = cptPDFFitNoMargin: opt.pdfMarginMM = 0#
            Case "2"
                opt.PDFPageMode = cptPDFFitCustomMargin
                value = Trim$(InputBox("PDF margin on all four sides in millimetres:", dialogTitle & " - PDF Margin", CStr(opt.pdfMarginMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.pdfMarginMM = CDbl(value)
                If opt.pdfMarginMM < 0# Then opt.pdfMarginMM = 0#
            Case "3": opt.PDFPageMode = cptPDFKeepOriginalPage
            Case "4", "5"
                If value = "4" Then opt.PDFPageMode = cptPDFCustomKeepPositions Else opt.PDFPageMode = cptPDFCustomCenterObjects
                value = Trim$(InputBox("Custom PDF page width in millimetres:", dialogTitle & " - PDF Width", CStr(opt.pdfCustomWidthMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.pdfCustomWidthMM = CDbl(value)
                value = Trim$(InputBox("Custom PDF page height in millimetres:", dialogTitle & " - PDF Height", CStr(opt.pdfCustomHeightMM)))
                If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
                opt.pdfCustomHeightMM = CDbl(value)
                If opt.pdfCustomWidthMM <= 0# Or opt.pdfCustomHeightMM <= 0# Then Exit Function
            Case Else: MsgBox "Invalid PDF page mode.", vbExclamation: Exit Function
        End Select

        ans = MsgBox( _
            "Automatically scale PDF-only working copies to 1:10 when a page exceeds " & _
            CStr(PDF_SAFE_MAX_MM) & " mm?" & vbCrLf & vbCrLf & _
            "Curved CDR files always remain full size.", _
            vbYesNoCancel + vbQuestion, dialogTitle & " - Large PDF")
        If ans = vbCancel Then Exit Function
        opt.autoScaleLargePDF = (ans = vbYes)
    End If

    ans = MsgBox("Rename pages in the original source document?", vbYesNoCancel + vbQuestion, dialogTitle & " - Source Pages")
    If ans = vbCancel Then Exit Function
    opt.renameSourcePages = (ans = vbYes)

    opt.outputFolder = CPT_GetFolderFromUser("Choose output folder for " & dialogTitle)
    If Len(opt.outputFolder) = 0 Then Exit Function

    CPT_PromptExportOptions = True
End Function

' ============================================================================
' PROFILE STORAGE
' ============================================================================
Private Sub CPT_SaveWorkflowProfile(ByRef profile As CPTWorkflowProfile)
    Dim section As String

    section = CPT_ProfileSection(profile.Name)
    SaveSetting PROFILE_APP, section, "Name", profile.Name

    With profile.ExportOptions
        SaveSetting PROFILE_APP, section, "Scope", CStr(.Scope)
        SaveSetting PROFILE_APP, section, "NameMode", CStr(.NameMode)
        SaveSetting PROFILE_APP, section, "CustomName", .customName
        SaveSetting PROFILE_APP, section, "OutputFolder", .outputFolder
        SaveSetting PROFILE_APP, section, "ExportCurrentCDR", CPT_BoolText(.exportCurrentCDR)
        SaveSetting PROFILE_APP, section, "ExportV15CDR", CPT_BoolText(.exportV15CDR)
        SaveSetting PROFILE_APP, section, "ExportCurvesCurrentCDR", CPT_BoolText(.exportCurvesCurrentCDR)
        SaveSetting PROFILE_APP, section, "ExportCurvesV15CDR", CPT_BoolText(.exportCurvesV15CDR)
        SaveSetting PROFILE_APP, section, "ExportCurvesPDF", CPT_BoolText(.exportCurvesPDF)
        SaveSetting PROFILE_APP, section, "FitPageToArtwork", CPT_BoolText(.FitPageToArtwork)
        SaveSetting PROFILE_APP, section, "IncludeOutlines", CPT_BoolText(.includeOutlines)
        SaveSetting PROFILE_APP, section, "MarginMM", CStr(.marginMM)
        SaveSetting PROFILE_APP, section, "CDRPageMode", CStr(.CDRPageMode)
        SaveSetting PROFILE_APP, section, "CDRCustomWidthMM", CStr(.cdrCustomWidthMM)
        SaveSetting PROFILE_APP, section, "CDRCustomHeightMM", CStr(.cdrCustomHeightMM)
        SaveSetting PROFILE_APP, section, "PDFPageMode", CStr(.PDFPageMode)
        SaveSetting PROFILE_APP, section, "PDFMarginMM", CStr(.pdfMarginMM)
        SaveSetting PROFILE_APP, section, "PDFCustomWidthMM", CStr(.pdfCustomWidthMM)
        SaveSetting PROFILE_APP, section, "PDFCustomHeightMM", CStr(.pdfCustomHeightMM)
        SaveSetting PROFILE_APP, section, "RenameSourcePages", CPT_BoolText(.renameSourcePages)
        SaveSetting PROFILE_APP, section, "AutoScaleLargePDF", CPT_BoolText(.autoScaleLargePDF)
        SaveSetting PROFILE_APP, section, "EnsureUniqueNames", CPT_BoolText(.ensureUniqueNames)
    End With

    With profile.DimensionOptions
        SaveSetting PROFILE_APP, section, "DimensionEnabled", CPT_BoolText(.Enabled)
        SaveSetting PROFILE_APP, section, "DimensionTarget", CStr(.targetMode)
        SaveSetting PROFILE_APP, section, "DimensionAxes", CStr(.Axes)
        SaveSetting PROFILE_APP, section, "DimensionMode", CStr(.mode)
        SaveSetting PROFILE_APP, section, "DimensionUnit", .unitText
        SaveSetting PROFILE_APP, section, "DimensionDecimals", CStr(.decimals)
        SaveSetting PROFILE_APP, section, "DimensionFont", .fontName
        SaveSetting PROFILE_APP, section, "DimensionFontSize", CStr(.fontSizePt)
        SaveSetting PROFILE_APP, section, "DimensionOffset", CStr(.lineOffsetMM)
        SaveSetting PROFILE_APP, section, "DimensionIncludeFit", CPT_BoolText(.IncludeInPageFit)
        SaveSetting PROFILE_APP, section, "DimensionRemoveOld", CPT_BoolText(.RemoveOldDimensions)
    End With

    CPT_AddProfileName profile.Name
End Sub

Private Function CPT_LoadWorkflowProfile(ByVal profileName As String, ByRef profile As CPTWorkflowProfile) As Boolean
    Dim section As String

    section = CPT_ProfileSection(profileName)
    If Len(GetSetting(PROFILE_APP, section, "Name", "")) = 0 Then Exit Function

    CPT_SetDefaultOptions profile.ExportOptions
    CPT_SetDefaultDimensionOptions profile.DimensionOptions
    profile.Name = GetSetting(PROFILE_APP, section, "Name", profileName)
    profile.BuiltIn = False

    With profile.ExportOptions
        .Scope = CLng(GetSetting(PROFILE_APP, section, "Scope", CStr(.Scope)))
        .NameMode = CLng(GetSetting(PROFILE_APP, section, "NameMode", CStr(.NameMode)))
        .customName = GetSetting(PROFILE_APP, section, "CustomName", "")
        .outputFolder = GetSetting(PROFILE_APP, section, "OutputFolder", "")
        .exportCurrentCDR = CPT_SettingBool(section, "ExportCurrentCDR", .exportCurrentCDR)
        .exportV15CDR = CPT_SettingBool(section, "ExportV15CDR", .exportV15CDR)
        .exportCurvesCurrentCDR = CPT_SettingBool(section, "ExportCurvesCurrentCDR", .exportCurvesCurrentCDR)
        .exportCurvesV15CDR = CPT_SettingBool(section, "ExportCurvesV15CDR", .exportCurvesV15CDR)
        .exportCurvesPDF = CPT_SettingBool(section, "ExportCurvesPDF", .exportCurvesPDF)
        .FitPageToArtwork = CPT_SettingBool(section, "FitPageToArtwork", .FitPageToArtwork)
        .includeOutlines = CPT_SettingBool(section, "IncludeOutlines", .includeOutlines)
        .marginMM = CDbl(GetSetting(PROFILE_APP, section, "MarginMM", CStr(.marginMM)))
        .CDRPageMode = CLng(GetSetting(PROFILE_APP, section, "CDRPageMode", IIf(.FitPageToArtwork, CStr(cptCDRFitCustomMargin), CStr(cptCDRKeepOriginalPage))))
        .cdrCustomWidthMM = CDbl(GetSetting(PROFILE_APP, section, "CDRCustomWidthMM", CStr(.cdrCustomWidthMM)))
        .cdrCustomHeightMM = CDbl(GetSetting(PROFILE_APP, section, "CDRCustomHeightMM", CStr(.cdrCustomHeightMM)))
        .PDFPageMode = CLng(GetSetting(PROFILE_APP, section, "PDFPageMode", CStr(.PDFPageMode)))
        .pdfMarginMM = CDbl(GetSetting(PROFILE_APP, section, "PDFMarginMM", CStr(.pdfMarginMM)))
        .pdfCustomWidthMM = CDbl(GetSetting(PROFILE_APP, section, "PDFCustomWidthMM", CStr(.pdfCustomWidthMM)))
        .pdfCustomHeightMM = CDbl(GetSetting(PROFILE_APP, section, "PDFCustomHeightMM", CStr(.pdfCustomHeightMM)))
        .renameSourcePages = CPT_SettingBool(section, "RenameSourcePages", .renameSourcePages)
        .autoScaleLargePDF = CPT_SettingBool(section, "AutoScaleLargePDF", .autoScaleLargePDF)
        .ensureUniqueNames = CPT_SettingBool(section, "EnsureUniqueNames", .ensureUniqueNames)
    End With

    With profile.DimensionOptions
        .Enabled = CPT_SettingBool(section, "DimensionEnabled", .Enabled)
        .targetMode = CLng(GetSetting(PROFILE_APP, section, "DimensionTarget", CStr(.targetMode)))
        .Axes = CLng(GetSetting(PROFILE_APP, section, "DimensionAxes", CStr(.Axes)))
        .mode = CLng(GetSetting(PROFILE_APP, section, "DimensionMode", CStr(.mode)))
        .unitText = GetSetting(PROFILE_APP, section, "DimensionUnit", .unitText)
        .decimals = CLng(GetSetting(PROFILE_APP, section, "DimensionDecimals", CStr(.decimals)))
        .fontName = GetSetting(PROFILE_APP, section, "DimensionFont", .fontName)
        .fontSizePt = CDbl(GetSetting(PROFILE_APP, section, "DimensionFontSize", CStr(.fontSizePt)))
        .lineOffsetMM = CDbl(GetSetting(PROFILE_APP, section, "DimensionOffset", CStr(.lineOffsetMM)))
        .IncludeInPageFit = CPT_SettingBool(section, "DimensionIncludeFit", .IncludeInPageFit)
        .RemoveOldDimensions = CPT_SettingBool(section, "DimensionRemoveOld", .RemoveOldDimensions)
    End With

    CPT_LoadWorkflowProfile = True
End Function

Private Function CPT_ChooseSavedProfile(ByRef profile As CPTWorkflowProfile, ByVal title As String) As Boolean
    Dim namesText As String
    Dim names() As String
    Dim prompt As String
    Dim value As String
    Dim i As Long
    Dim selectedIndex As Long

    namesText = CPT_GetProfileNames()
    If Len(namesText) = 0 Then
        MsgBox "No custom profiles are saved.", vbInformation, title
        Exit Function
    End If

    names = Split(namesText, "|")
    For i = LBound(names) To UBound(names)
        prompt = prompt & CStr(i + 1) & " = " & names(i) & vbCrLf
    Next i

    value = Trim$(InputBox("Choose a saved profile:" & vbCrLf & prompt, title, "1"))
    If Len(value) = 0 Or Not IsNumeric(value) Then Exit Function
    selectedIndex = CLng(value) - 1
    If selectedIndex < LBound(names) Or selectedIndex > UBound(names) Then
        MsgBox "Invalid profile number.", vbExclamation, title
        Exit Function
    End If

    CPT_ChooseSavedProfile = CPT_LoadWorkflowProfile(names(selectedIndex), profile)
End Function

Private Sub CPT_DeleteWorkflowProfile(ByVal profileName As String)
    Dim namesText As String
    Dim names() As String
    Dim newNames As String
    Dim i As Long

    On Error Resume Next
    DeleteSetting PROFILE_APP, CPT_ProfileSection(profileName)
    On Error GoTo 0

    namesText = CPT_GetProfileNames()
    If Len(namesText) = 0 Then Exit Sub
    names = Split(namesText, "|")
    For i = LBound(names) To UBound(names)
        If StrComp(names(i), profileName, vbTextCompare) <> 0 Then
            If Len(newNames) > 0 Then newNames = newNames & "|"
            newNames = newNames & names(i)
        End If
    Next i
    SaveSetting PROFILE_APP, PROFILE_INDEX_SECTION, PROFILE_INDEX_KEY, newNames
End Sub

Private Sub CPT_DeleteAllWorkflowProfiles()
    Dim namesText As String
    Dim names() As String
    Dim i As Long

    namesText = CPT_GetProfileNames()
    On Error Resume Next
    If Len(namesText) > 0 Then
        names = Split(namesText, "|")
        For i = LBound(names) To UBound(names)
            DeleteSetting PROFILE_APP, CPT_ProfileSection(names(i))
        Next i
    End If
    DeleteSetting PROFILE_APP, PROFILE_INDEX_SECTION
    On Error GoTo 0
End Sub

Private Sub CPT_AddProfileName(ByVal profileName As String)
    Dim namesText As String
    Dim names() As String
    Dim i As Long

    namesText = CPT_GetProfileNames()
    If Len(namesText) > 0 Then
        names = Split(namesText, "|")
        For i = LBound(names) To UBound(names)
            If StrComp(names(i), profileName, vbTextCompare) = 0 Then Exit Sub
        Next i
        namesText = namesText & "|" & profileName
    Else
        namesText = profileName
    End If
    SaveSetting PROFILE_APP, PROFILE_INDEX_SECTION, PROFILE_INDEX_KEY, namesText
End Sub

Private Function CPT_GetProfileNames() As String
    CPT_GetProfileNames = GetSetting(PROFILE_APP, PROFILE_INDEX_SECTION, PROFILE_INDEX_KEY, "")
End Function

Private Function CPT_ProfileSection(ByVal profileName As String) As String
    Dim value As String

    value = CPT_MakeSafeFileName(profileName)
    value = Replace(value, " ", "_")
    CPT_ProfileSection = "Profile_" & value
End Function

Private Function CPT_BoolText(ByVal value As Boolean) As String
    If value Then CPT_BoolText = "1" Else CPT_BoolText = "0"
End Function

Private Function CPT_SettingBool(ByVal section As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim value As String

    value = GetSetting(PROFILE_APP, section, key, CPT_BoolText(defaultValue))
    CPT_SettingBool = (value = "1" Or LCase$(value) = "true")
End Function

' ============================================================================
' WORKFLOW SUMMARIES
' ============================================================================
Private Function CPT_ProfileSummary(ByRef profile As CPTWorkflowProfile) As String
    Dim formats As String
    Dim scopeText As String
    Dim cdrPageText As String
    Dim dimensionText As String

    With profile.ExportOptions
        If .exportCurrentCDR Then formats = formats & "Current CDR, "
        If .exportV15CDR Then formats = formats & "V15 CDR, "
        If .exportCurvesCurrentCDR Then formats = formats & "Curved CDR, "
        If .exportCurvesV15CDR Then formats = formats & "Curved V15 CDR, "
        If .exportCurvesPDF Then formats = formats & "Curved PDF, "
        If Len(formats) >= 2 Then formats = Left$(formats, Len(formats) - 2)

        Select Case .Scope
            Case cptScopeEachPage: scopeText = "Each page separately"
            Case cptScopeAllPages: scopeText = "All pages in one document"
            Case cptScopeCurrentPage: scopeText = "Current page only"
        End Select

        cdrPageText = CPT_CDRPageModeText(.CDRPageMode, .marginMM, .cdrCustomWidthMM, .cdrCustomHeightMM)
    End With

    If profile.DimensionOptions.Enabled Then
        dimensionText = CPT_DimensionOptionsSummary(profile.DimensionOptions)
    Else
        dimensionText = "None"
    End If

    CPT_ProfileSummary = _
        "Profile: " & profile.Name & vbCrLf & _
        "Dimensions: " & dimensionText & vbCrLf & _
        "Scope: " & scopeText & vbCrLf & _
        "Formats: " & formats & vbCrLf & _
        "CDR pages: " & cdrPageText & vbCrLf & _
        "PDF pages: " & CPT_PDFPageModeText(profile.ExportOptions.PDFPageMode, profile.ExportOptions.pdfMarginMM, profile.ExportOptions.pdfCustomWidthMM, profile.ExportOptions.pdfCustomHeightMM) & vbCrLf & _
        "Output: " & profile.ExportOptions.outputFolder
End Function

Private Function CPT_DimensionOptionsSummary(ByRef dimOpt As CPTDimensionOptions) As String
    Dim targetText As String
    Dim axesText As String
    Dim modeText As String

    If dimOpt.targetMode = cptDimensionEachObject Then targetText = "each selected object" Else targetText = "complete selection"
    Select Case dimOpt.Axes
        Case cptDimensionWidthOnly: axesText = "width"
        Case cptDimensionHeightOnly: axesText = "height"
        Case Else: axesText = "width + height"
    End Select
    Select Case dimOpt.mode
        Case cptDimensionPermanentSource: modeText = "permanent in source"
        Case cptDimensionTemporarySource: modeText = "temporary in source"
        Case Else: modeText = "export copies only"
    End Select

    CPT_DimensionOptionsSummary = axesText & " for " & targetText & ", " & dimOpt.unitText & ", " & modeText
End Function

Private Function CPT_CDRPageModeText( _
    ByVal mode As CPTCDRPageMode, _
    ByVal marginMM As Double, _
    ByVal widthMM As Double, _
    ByVal heightMM As Double) As String

    Select Case mode
        Case cptCDRFitNoMargin
            CPT_CDRPageModeText = "Group all objects and fit page, no margin"
        Case cptCDRFitCustomMargin
            CPT_CDRPageModeText = "Group all objects and fit page with " & CStr(marginMM) & " mm margin"
        Case cptCDRCustomKeepPositions
            CPT_CDRPageModeText = "Custom " & CStr(widthMM) & " x " & CStr(heightMM) & " mm, keep object positions"
        Case cptCDRCustomCenterObjects
            CPT_CDRPageModeText = "Custom " & CStr(widthMM) & " x " & CStr(heightMM) & " mm, centre objects"
        Case Else
            CPT_CDRPageModeText = "Preserve source page exactly"
    End Select
End Function

Private Function CPT_PDFPageModeText( _
    ByVal mode As CPTPDFPageMode, _
    ByVal marginMM As Double, _
    ByVal widthMM As Double, _
    ByVal heightMM As Double) As String

    Select Case mode
        Case cptPDFFitNoMargin
            CPT_PDFPageModeText = "Group all objects and fit page, no margin"
        Case cptPDFFitCustomMargin
            CPT_PDFPageModeText = "Group all objects and fit page with " & CStr(marginMM) & " mm margin"
        Case cptPDFCustomKeepPositions
            CPT_PDFPageModeText = "Custom " & CStr(widthMM) & " x " & CStr(heightMM) & " mm, keep object positions"
        Case cptPDFCustomCenterObjects
            CPT_PDFPageModeText = "Custom " & CStr(widthMM) & " x " & CStr(heightMM) & " mm, centre objects"
        Case Else
            CPT_PDFPageModeText = "Preserve source page exactly"
    End Select
End Function

' ============================================================================
' PROGRESS / CANCELLATION
' ============================================================================
Private Sub CPT_ProgressBegin(ByVal message As String)
    gCancelRequested = False
    On Error Resume Next
    Application.Status.BeginProgress message, True
    Application.Status.Progress = 0
    If Not gExportUI Is Nothing Then gExportUI.BeginProgress message
    On Error GoTo 0
End Sub

Private Sub CPT_ProgressSet(ByVal currentStep As Long, ByVal totalSteps As Long, ByVal message As String)
    Dim pct As Long

    If totalSteps <= 0 Then totalSteps = 1
    pct = CLng((CDbl(currentStep) / CDbl(totalSteps)) * 100#)
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    On Error Resume Next
    Application.Status.SetProgressMessage message
    Application.Status.Progress = pct
    If Not gExportUI Is Nothing Then gExportUI.UpdateProgress pct, message
    On Error GoTo 0

    DoEvents
    CPT_CheckCancelled
End Sub

Private Sub CPT_CheckCancelled()
    Dim wasAborted As Boolean

    DoEvents
    On Error Resume Next
    wasAborted = Application.Status.Aborted
    On Error GoTo 0

    If wasAborted Or gCancelRequested Then
        Err.Raise ERR_CANCELLED, "VFE_CorelProductionTools27", "The user cancelled the export."
    End If
End Sub

' ============================================================================
' GENERAL HELPERS
' ============================================================================
Private Sub CPT_SetDefaultOptions(ByRef opt As CPTExportOptions)
    opt.Scope = cptScopeEachPage
    opt.NameMode = cptNameExportField
    opt.customName = ""
    opt.outputFolder = ""

    opt.exportCurrentCDR = False
    opt.exportV15CDR = False
    opt.exportCurvesCurrentCDR = False
    opt.exportCurvesV15CDR = False
    opt.exportCurvesPDF = False

    opt.FitPageToArtwork = False
    opt.includeOutlines = True
    opt.marginMM = DEFAULT_MARGIN_MM
    opt.CDRPageMode = cptCDRKeepOriginalPage
    opt.cdrCustomWidthMM = 210#
    opt.cdrCustomHeightMM = 297#
    opt.PDFPageMode = cptPDFKeepOriginalPage
    opt.pdfMarginMM = DEFAULT_MARGIN_MM
    opt.pdfCustomWidthMM = 210#
    opt.pdfCustomHeightMM = 297#
    opt.renameSourcePages = False
    opt.autoScaleLargePDF = True
    opt.ensureUniqueNames = True
End Sub

Private Function CPT_AnyFormatSelected(ByRef opt As CPTExportOptions) As Boolean
    CPT_AnyFormatSelected = opt.exportCurrentCDR Or _
                            opt.exportV15CDR Or _
                            opt.exportCurvesCurrentCDR Or _
                            opt.exportCurvesV15CDR Or _
                            opt.exportCurvesPDF
End Function

Private Function CPT_HasActiveDocument() As Boolean
    If Documents.Count = 0 Then
        MsgBox "Open a CorelDRAW document first.", vbExclamation, "VFE Corel Production Tools 27"
        CPT_HasActiveDocument = False
    Else
        CPT_HasActiveDocument = True
    End If
End Function

Private Function CPT_AskNameModeExportOrDocument() As Long
    Dim ans As VbMsgBoxResult

    ans = MsgBox( _
        "Does this job contain values in text objects named EXPORT_NAME?" & vbCrLf & vbCrLf & _
        "YES = use EXPORT_NAME per page" & vbCrLf & _
        "NO = use the document filename plus page number", _
        vbYesNoCancel + vbQuestion, "Export Name Source")

    Select Case ans
        Case vbYes: CPT_AskNameModeExportOrDocument = cptNameExportField
        Case vbNo: CPT_AskNameModeExportOrDocument = cptNameDocument
        Case Else: CPT_AskNameModeExportOrDocument = 0
    End Select
End Function

Private Function CPT_GetFolderFromUser(ByVal title As String) As String
    Dim shellObject As Object
    Dim folderObject As Object
    Dim fallback As String

    On Error GoTo UseFallback
    Set shellObject = CreateObject("Shell.Application")
    Set folderObject = shellObject.BrowseForFolder(0, title, 0, 0)

    If folderObject Is Nothing Then
        CPT_GetFolderFromUser = ""
    Else
        CPT_GetFolderFromUser = CPT_NormalizeFolder(folderObject.Self.path)
    End If
    Exit Function

UseFallback:
    fallback = Trim$(InputBox("Enter the full output folder path:", title))
    CPT_GetFolderFromUser = CPT_NormalizeFolder(fallback)
End Function

Private Function CPT_NormalizeFolder(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If Len(folderPath) = 0 Then
        CPT_NormalizeFolder = ""
    ElseIf Right$(folderPath, 1) = "\" Then
        CPT_NormalizeFolder = folderPath
    Else
        CPT_NormalizeFolder = folderPath & "\"
    End If
End Function

Private Function CPT_FolderExists(ByVal folderPath As String) As Boolean
    Dim testPath As String
    Dim attributes As Long

    testPath = folderPath
    If Right$(testPath, 1) = "\" And Len(testPath) > 3 Then
        testPath = Left$(testPath, Len(testPath) - 1)
    End If

    On Error GoTo NotFound
    attributes = GetAttr(testPath)
    CPT_FolderExists = ((attributes And vbDirectory) = vbDirectory)
    Exit Function
NotFound:
    CPT_FolderExists = False
End Function

Private Function CPT_FileExists(ByVal fullPath As String) As Boolean
    On Error Resume Next
    CPT_FileExists = (Len(Dir$(fullPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    On Error GoTo 0
End Function

Private Function CPT_GetDocumentBaseName(ByVal doc As Document) As String
    Dim value As String
    Dim p As Long

    On Error Resume Next
    value = doc.FileName
    If Len(Trim$(value)) = 0 Then value = doc.Name
    On Error GoTo 0

    value = Trim$(value)
    p = InStrRev(value, "\")
    If p > 0 Then value = Mid$(value, p + 1)

    p = InStrRev(value, ".")
    If p > 1 Then value = Left$(value, p - 1)

    CPT_GetDocumentBaseName = Trim$(value)
End Function

Private Function CPT_GetDocumentDisplayName(ByVal doc As Document) As String
    Dim value As String

    On Error Resume Next
    value = doc.FullFileName
    If Len(Trim$(value)) = 0 Then value = doc.Name
    On Error GoTo 0

    If Len(Trim$(value)) = 0 Then value = "Untitled document"
    CPT_GetDocumentDisplayName = value
End Function

Private Sub CPT_AppendLine(ByRef textValue As String, ByVal lineValue As String)
    If Len(textValue) > 0 And Right$(textValue, 2) <> vbCrLf Then textValue = textValue & vbCrLf
    textValue = textValue & lineValue & vbCrLf
End Sub

Private Sub CPT_WriteTextFile(ByVal fullPath As String, ByVal contents As String)
    Dim fileNumber As Integer

    On Error Resume Next
    fileNumber = FreeFile
    Open fullPath For Output As #fileNumber
    Print #fileNumber, contents
    Close #fileNumber
    On Error GoTo 0
End Sub
