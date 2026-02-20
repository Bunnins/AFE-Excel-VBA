Attribute VB_Name = "ComponentsDue"
Public Sub RunAllSteps()
    On Error GoTo ErrorHandler

    Dim runStart As Double
    Dim stepStart As Double
    Dim stepElapsed As Double
    runStart = Timer

    Const PERF_MODE As Boolean = True

    Dim runStart As Double
    Dim stepStart As Double
    Dim stepElapsed As Double
    runStart = Timer

    gPerfLock = PERF_MODE
    If PERF_MODE Then BeginPerfMode

    Dim result As String
    result = MsgBox("Please ensure Data is Fully Loaded before Proceeding - Is all Data Loaded", vbYesNo)
    If result = vbNo Then
        If PERF_MODE Then EndPerfMode
        Exit Sub
    End If

    ufProgress.LabelProgress.Width = 0
    ufProgress.Show

    Dim totalSteps As Integer
    totalSteps = 26 ' Update this based on the actual number of steps

    Dim currentStep As Integer
    currentStep = 1 ' Initialize step counter

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Creating Workbook")
    stepStart = Timer
    Call Createworkbook
    stepElapsed = Timer - stepStart
    Debug.Print "Createworkbook: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Clearing Data and Data2 sheets")
    stepStart = Timer
    Call ClearDataAndUpdate
    stepElapsed = Timer - stepStart
    Debug.Print "ClearDataAndUpdate: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Removing duplicates from Labour Costing sheet")
    stepStart = Timer
    Call Removeduplicates
    stepElapsed = Timer - stepStart
    Debug.Print "Removeduplicates: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Applying Text to Columns")
    stepStart = Timer
    Call Text_to_Columns
    stepElapsed = Timer - stepStart
    Debug.Print "Text_to_Columns: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Executing Macro1")
    stepStart = Timer
    Call Macro1
    stepElapsed = Timer - stepStart
    Debug.Print "Macro1: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Searching and returning values")
    stepStart = Timer
    Call SearchAndReturnValues
    stepElapsed = Timer - stepStart
    Debug.Print "SearchAndReturnValues: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Running CounterReturn")
    stepStart = Timer
    Call CounterReturn
    stepElapsed = Timer - stepStart
    Debug.Print "CounterReturn: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Running CounterReturn2")
    stepStart = Timer
    Call CounterReturn2
    stepElapsed = Timer - stepStart
    Debug.Print "CounterReturn2: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Removing blank rows from Components Due sheet")
    stepStart = Timer
    Call RemoveBlankRows(Workbooks("AFE Builder from SAP HANA").Sheets("Components Due"))
    stepElapsed = Timer - stepStart
    Debug.Print "RemoveBlankRows: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Sorting data in Components Due sheet")
    stepStart = Timer
    Call Sort_Order
    stepElapsed = Timer - stepStart
    Debug.Print "Sort_Order: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Clearing specific columns from another sheet")
    stepStart = Timer
    Call ClearTable1672ColI
    stepElapsed = Timer - stepStart
    Debug.Print "ClearTable1672ColI: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Extracting matching items")
    stepStart = Timer
    Call ExtractMatchingItems
    stepElapsed = Timer - stepStart
    Debug.Print "ExtractMatchingItems: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Listing occurrences in Labour Costing")
    stepStart = Timer
    Call ListOccurrencesInLabourCosting
    stepElapsed = Timer - stepStart
    Debug.Print "ListOccurrencesInLabourCosting: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Summarizing Labour Data")
    stepStart = Timer
    Call SummarizeLabourData
    stepElapsed = Timer - stepStart
    Debug.Print "SummarizeLabourData: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Summarizing Labour Data2")
    stepStart = Timer
    Call SummarizeLabourData2
    stepElapsed = Timer - stepStart
    Debug.Print "SummarizeLabourData2: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Applying XLOOKUP Formula")
    stepStart = Timer
    Call ApplyXLOOKUPFormula
    stepElapsed = Timer - stepStart
    Debug.Print "ApplyXLOOKUPFormula: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Applying XLOOKUP for Column F")
    stepStart = Timer
    Call ApplyXLOOKUPForColumnF
    stepElapsed = Timer - stepStart
    Debug.Print "ApplyXLOOKUPForColumnF: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Running test")
    stepStart = Timer
    Call test
    stepElapsed = Timer - stepStart
    Debug.Print "test: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Copying Data Worksheet")
    stepStart = Timer
    Call CopyDataWorksheet
    stepElapsed = Timer - stepStart
    Debug.Print "CopyDataWorksheet: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Applying XLOOKUP Down Sheet AFE")
    stepStart = Timer
    Call ApplyXLOOKUPDownSheet_AFE
    stepElapsed = Timer - stepStart
    Debug.Print "ApplyXLOOKUPDownSheet_AFE: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Finalizing Labour")
    stepStart = Timer
    Call LabourFinal
    stepElapsed = Timer - stepStart
    Debug.Print "LabourFinal: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Summarizing Totals")
    stepStart = Timer
    Call SummaryTotals
    stepElapsed = Timer - stepStart
    Debug.Print "SummaryTotals: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Summarizing PMEX Totals")
    stepStart = Timer
    Call SummaryPMEXTotals
    stepElapsed = Timer - stepStart
    Debug.Print "SummaryPMEXTotals: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Cleaning Data")
    stepStart = Timer
    Call DataClean
    stepElapsed = Timer - stepStart
    Debug.Print "DataClean: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Tallying Values Between Yellow Cells")
    stepStart = Timer
    Call TallyValuesBetweenYellowCells
    stepElapsed = Timer - stepStart
    Debug.Print "TallyValuesBetweenYellowCells: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Assigning Vendors")
    stepStart = Timer
    Call VendorAssignment
    stepElapsed = Timer - stepStart
    Debug.Print "VendorAssignment: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    ' Update progress
    Call UpdateProgress(currentStep, totalSteps, "Breaking Links")
    stepStart = Timer
    Call BreakLinks
    stepElapsed = Timer - stepStart
    Debug.Print "BreakLinks: " & Format(stepElapsed, "0.00") & "s"
    currentStep = currentStep + 1

    Debug.Print "RunAllSteps Total: " & Format(Timer - runStart, "0.00") & "s"

    Debug.Print "RunAllSteps Total: " & Format(Timer - runStart, "0.00") & "s"

    ' Final progress update
    Call UpdateProgress(currentStep, totalSteps, "Completed")
    If PERF_MODE Then EndPerfMode
    ufProgress.Hide
    Exit Sub

ErrorHandler:
    If PERF_MODE Then EndPerfMode
    MsgBox "An error occurred: " & Err.Description
    ufProgress.Hide
End Sub

Private Sub BeginPerfMode()
    If gPerfApplied Then Exit Sub

    gPrevScreenUpdating = Application.ScreenUpdating
    gPrevEnableEvents = Application.EnableEvents
    gPrevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    gPerfApplied = True
End Sub

Private Sub EndPerfMode()
    If Not gPerfApplied Then Exit Sub

    Application.ScreenUpdating = gPrevScreenUpdating
    Application.EnableEvents = gPrevEnableEvents
    Application.Calculation = gPrevCalculation

    gPerfApplied = False
    gPerfLock = False
End Sub

Public Sub SetScreenUpdatingSafely(ByVal isEnabled As Boolean)
    If gPerfLock Then Exit Sub
    Application.ScreenUpdating = isEnabled
End Sub

Public Sub SetEnableEventsSafely(ByVal isEnabled As Boolean)
    If gPerfLock Then Exit Sub
    Application.EnableEvents = isEnabled
End Sub

Public Sub SetCalculationSafely(ByVal calcMode As XlCalculation)
    If gPerfLock Then Exit Sub
    Application.Calculation = calcMode
End Sub

Sub UpdateProgress(currentStep As Integer, totalSteps As Integer, stepDescription As String)
    Dim progress As Double
    progress = (currentStep / totalSteps) * 100
    ufProgress.LabelProgress.Width = progress * (ufProgress.FrameProgress.Width / 100)
    ufProgress.Caption = "Progress: " & Round(progress) & "% - " & stepDescription
    DoEvents
End Sub




Sub Createworkbook()
    Dim newWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim templateSheet As Worksheet
    Dim template2Sheet As Worksheet
    Dim template3Sheet As Worksheet
    Dim password As String
    password = "PlanExRavMB" ' Your actual password

    SetScreenUpdatingSafely False

    ' Set the current workbook (the workbook you're working in)
    Set currentWorkbook = ThisWorkbook

    ' Store the references to the template sheets
    Set templateSheet = currentWorkbook.Sheets("Template")
    Set template2Sheet = currentWorkbook.Sheets("Template2")
    Set template3Sheet = currentWorkbook.Sheets("Template3")

    With ufProgress
        .LabelCaption.Caption = "Creating Workbook - 25% Complete"
        .LabelProgress.Width = 0.25 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Unprotect the workbook
    currentWorkbook.Unprotect password:=password

    ' Unhide the sheets temporarily (just to copy them)
    templateSheet.Visible = xlSheetVisible
    template2Sheet.Visible = xlSheetVisible
    template3Sheet.Visible = xlSheetVisible

    ' Create a new workbook
    Set newWorkbook = Workbooks.Add

    ' Copy each template sheet from the current workbook to the new workbook
    templateSheet.Copy Before:=newWorkbook.Sheets(1)
    template2Sheet.Copy Before:=newWorkbook.Sheets(2)
    template3Sheet.Copy Before:=newWorkbook.Sheets(3)

    ' Rename the sheets in the new workbook
    newWorkbook.Sheets("Template").Name = "Summary"
    newWorkbook.Sheets("Template2").Name = "Components Due"
    newWorkbook.Sheets("Template3").Name = "Budget"

    With ufProgress
        .LabelCaption.Caption = "Creating Workbook - 50% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Change the name of the new workbook without saving it
    newWorkbook.Windows(1).Caption = "AFE"  ' Set the window name to "AFE"

    ' Hide the default Sheet1 in the new workbook (if it exists)
    On Error Resume Next ' In case Sheet1 doesn't exist
    newWorkbook.Sheets("Sheet1").Visible = False
    On Error GoTo 0 ' Reset error handling

    With ufProgress
        .LabelCaption.Caption = "Creating Workbook - 75% Complete"
        .LabelProgress.Width = 0.75 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Add the SUM formula to cell C4 in the "Summary" sheet (newWorkbook.Sheets("Summary"))
    newWorkbook.Sheets("Summary").Range("C4").formula = "=SUM('Components Due'!I3:I1048576)"

    ' Save the new workbook as AFE.xlsx in the same directory as the current workbook
    newWorkbook.SaveAs Filename:=currentWorkbook.Path & "\AFE.xlsx", FileFormat:=xlOpenXMLWorkbook

    ' Now, hide the template sheets again in the current workbook
    templateSheet.Visible = xlSheetVeryHidden
    template2Sheet.Visible = xlSheetVeryHidden
    template3Sheet.Visible = xlSheetVeryHidden

    ' Reprotect the workbook
    currentWorkbook.Protect password:=password, Structure:=True

    With ufProgress
        .LabelCaption.Caption = "Creating Workbook - 100% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    SetScreenUpdatingSafely True
End Sub




Sub ClearDataAndUpdate()
    ' Declare variables for the sheets and workbook
    Dim Data As Worksheet
    Dim Data2 As Worksheet
    Dim LCData As Worksheet
    Dim wb As Workbook
    SetScreenUpdatingSafely False
    With ufProgress
        .LabelCaption.Caption = "Clearing Data - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Check if the workbook "AFE Builder from SAP HANA" is open
    On Error Resume Next
    Set wb = Workbooks("AFE Builder from SAP HANA")
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "'AFE Builder from SAP HANA' workbook is not open.", vbCritical
        Exit Sub
    End If
    
    ' Set references to Data, Data2, and LC Data sheets
    Set Data = wb.Sheets("Data")
    Set Data2 = wb.Sheets("Data2")
    Set LCData = wb.Sheets("LC Data")

       With ufProgress
        .LabelCaption.Caption = "Clearing Data - 50" & "% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Check if Data, Data2, and LC Data sheets exist
    If Data Is Nothing Then
        MsgBox "'Data' sheet not found in AFE Builder from SAP HANA workbook.", vbCritical
        Exit Sub
    End If
    If Data2 Is Nothing Then
        MsgBox "'Data2' sheet not found in AFE Builder from SAP HANA workbook.", vbCritical
        Exit Sub
    End If
    If LCData Is Nothing Then
        MsgBox "'LC Data' sheet not found in AFE Builder from SAP HANA workbook.", vbCritical
        Exit Sub
    End If
    
    ' Clear the Data, Data2, and LC Data sheets
    Data.Cells.Clear
    Data2.Cells.Clear
    LCData.Cells.Clear

           With ufProgress
        .LabelCaption.Caption = "Clearing Data - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    
    ' Optional: Display a confirmation message after clearing the sheets
    'MsgBox "Data, Data2, and LC Data sheets have been cleared.", vbInformation
End Sub

Sub Removeduplicates()
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim dataArr As Variant
    Dim outArr() As Variant
    Dim dict As Object
    Dim i As Long, c As Long
    Dim keepCount As Long
    Dim key As String

    SetScreenUpdatingSafely False
    With ufProgress
        .LabelCaption.Caption = "Removing Duplicates - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents

    Set ws = ThisWorkbook.Sheets("Task_list_rep")
    On Error Resume Next
    Set tbl = ws.ListObjects("TASK_LIST_REP")
    On Error GoTo 0

    With ufProgress
        .LabelCaption.Caption = "Removing Duplicates - 50" & "% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents

    If tbl Is Nothing Then Exit Sub
    If tbl.ListColumns.Count < 16 Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    dataArr = tbl.DataBodyRange.Value2
    ReDim outArr(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))
    Set dict = CreateObject("Scripting.Dictionary")

    keepCount = 0
    For i = 1 To UBound(dataArr, 1)
        key = CStr(dataArr(i, 8)) & ChrW(30) & CStr(dataArr(i, 16))
        If Not dict.Exists(key) Then
            dict.Add key, 1
            keepCount = keepCount + 1
            For c = 1 To UBound(dataArr, 2)
                outArr(keepCount, c) = dataArr(i, c)
            Next c
        End If
    Next i

    tbl.DataBodyRange.ClearContents
    If keepCount > 0 Then
        tbl.Resize ws.Range(tbl.HeaderRowRange.Cells(1, 1), tbl.HeaderRowRange.Cells(1, 1).Offset(keepCount, tbl.ListColumns.Count - 1))
        tbl.DataBodyRange.Value2 = outArr
    Else
        tbl.Resize ws.Range(tbl.HeaderRowRange.Cells(1, 1), tbl.HeaderRowRange.Cells(1, 1).Offset(0, tbl.ListColumns.Count - 1))
    End If

    With ufProgress
        .LabelCaption.Caption = "Removing Duplicates - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents
End Sub


Sub Macro1()
    ' Declare variables
    Dim frontSheet As Worksheet
    Dim componentsDueSheet As Worksheet
    Dim dataSheet As Worksheet
    Dim criteriaValue As String
    Dim maintPlansTable As ListObject
    Dim afeBuilderWorkbook As Workbook
    Dim filteredRange As Range

SetScreenUpdatingSafely False
    With ufProgress
        .LabelCaption.Caption = "Copying Data - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    
    ' Ensure "AFE Builder from SAP HANA" workbook is open
    On Error Resume Next
    Set afeBuilderWorkbook = Workbooks("AFE Builder from SAP HANA")
    On Error GoTo 0
    If afeBuilderWorkbook Is Nothing Then
        MsgBox "'AFE Builder from SAP HANA' workbook is not open!", vbCritical
        Exit Sub
    End If

    ' Set the "Front Sheet" from the AFE Builder workbook
    On Error Resume Next
    Set frontSheet = afeBuilderWorkbook.Sheets("Front Sheet")
    On Error GoTo 0
    If frontSheet Is Nothing Then
        MsgBox "'Front Sheet' does not exist in 'AFE Builder from SAP HANA' workbook!", vbCritical
        Exit Sub
    End If

    ' Check if "Components Due" sheet exists in the current workbook
    On Error Resume Next
    Set componentsDueSheet = ThisWorkbook.Sheets("Components Due")
    On Error GoTo 0
    If componentsDueSheet Is Nothing Then
        MsgBox "'Components Due' sheet does not exist in the current workbook!", vbCritical
        Exit Sub
    End If

    ' Check if "Data2" sheet exists in the current workbook
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("Data2")
    On Error GoTo 0
    If dataSheet Is Nothing Then
        MsgBox "'Data2' sheet does not exist in the current workbook!", vbCritical
        Exit Sub
    End If

    With ufProgress
        .LabelCaption.Caption = "Copying Data - 25" & "% Complete"
        .LabelProgress.Width = 0.25 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents


    ' Get the value from Front Sheet I11
    criteriaValue = frontSheet.Range("I11").Value
    If criteriaValue = "" Then
        MsgBox "I11 on 'Front Sheet' is empty!", vbCritical
        Exit Sub
    End If

    ' Check if the "MAINT_PLANS__2" table exists in "Components Due" sheet
    On Error Resume Next
    Set maintPlansTable = componentsDueSheet.ListObjects("MAINT_PLANS")
    On Error GoTo 0
    If maintPlansTable Is Nothing Then
        MsgBox "'MAINT_PLANS' table does not exist on the 'Components Due' sheet.", vbCritical
        Exit Sub
    End If

    With ufProgress
        .LabelCaption.Caption = "Copying Data - 50" & "% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents


    ' Apply the filter in Components Due sheet based on the value in I11
    componentsDueSheet.Activate
    maintPlansTable.Range.AutoFilter Field:=2, Criteria1:=criteriaValue

    ' Get the filtered visible range (excluding hidden rows)
    On Error Resume Next
    Set filteredRange = maintPlansTable.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' If no visible cells after filtering, display a message and exit
    If filteredRange Is Nothing Then
        MsgBox "No data found for the specified criteria.", vbExclamation
        Exit Sub
    End If
    With ufProgress
        .LabelCaption.Caption = "Copying Data - 75" & "% Complete"
        .LabelProgress.Width = 0.75 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    ' Copy only the values from the filtered range
    filteredRange.Copy

    ' Paste the copied data into Data2 sheet (starting at A1)
    dataSheet.Activate
    dataSheet.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Optional: Preserve the column width if needed (uncomment the next line if you want to keep the column widths the same)
    ' dataSheet.Cells.PasteSpecial Paste:=xlPasteColumnWidths

    ' Clear the Cut/Copy mode
    application.CutCopyMode = False

    With ufProgress
        .LabelCaption.Caption = "Copying Data - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Optional: Inform the user that the operation is complete
    'MsgBox "Data has been copied successfully!", vbInformation
End Sub




Sub Text_to_Columns()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("Components Due") ' Reference the "Components Due" sheet for Column G
    SetScreenUpdatingSafely False
    
    With ufProgress
        .LabelCaption.Caption = "Converting Data - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    
    ' Convert text to columns in Column G of "Components Due" starting from row 2
    sheet.Range("G2:G" & sheet.Cells(sheet.Rows.Count, "G").End(xlUp).Row).TextToColumns _
        Destination:=sheet.Range("G2"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True, _
        Space:=False, _
        Other:=False

    With ufProgress
        .LabelCaption.Caption = "Converting Data - 50" & "% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Convert text to columns in Column A of the "Parts Costing" sheet starting from row 2
    Set sheet = ThisWorkbook.Sheets("Parts Costing") ' Reference the "Parts Costing" sheet for Column A
    sheet.Range("A2:A" & sheet.Cells(sheet.Rows.Count, "A").End(xlUp).Row).TextToColumns _
        Destination:=sheet.Range("A2"), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True, _
        Space:=False, _
        Other:=False
        
        With ufProgress
        .LabelCaption.Caption = "Converting Data - 100" & "% Complete"
        .LabelProgress.Width = 100 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
End Sub

Sub SearchAndReturnValues()
    Dim sourceWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim searchValue As String
    Dim lastRow As Long
    Dim outputRow As Long
    Dim i As Long
    Dim valueFromAFEBuild As String ' Variable to hold value from AFE Builder - Front Sheet J7
    Dim valueFromI13 As String ' Variable to hold value from AFE Builder - Front Sheet I13

    ' Disable updates to improve speed
    SetScreenUpdatingSafely False
    SetCalculationSafely xlCalculationManual
    SetEnableEventsSafely False

    With ufProgress
        .LabelCaption.Caption = "Retrieving Values - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Set source and destination workbooks
    Set sourceWorkbook = ThisWorkbook
    Set destinationWorkbook = Workbooks("AFE.xlsx")
    
    ' Ensure the destination workbook is open
    If destinationWorkbook Is Nothing Then
        MsgBox "The AFE workbook is not open.", vbCritical
        Exit Sub
    End If
    
    ' Set source and destination sheets
    Set sourceSheet = sourceWorkbook.Sheets("Data2")
    Set destinationSheet = destinationWorkbook.Sheets("Components Due")
    
    ' Get the value from I12 in the Front Sheet of AFE Builder workbook
    valueFromAFEBuild = Workbooks("AFE Builder from SAP HANA").Sheets("Front Sheet").Range("I12").Value
    If valueFromAFEBuild = "" Then
        MsgBox "No value found in I12 of the Front Sheet.", vbCritical
        Exit Sub
    End If
   With ufProgress
        .LabelCaption.Caption = "Retrieving Values - 20" & "% Complete"
        .LabelProgress.Width = 0.2 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    ' Place the value from I12 into K2 of Components Due in the AFE workbook
    destinationSheet.Range("K2").Value = valueFromAFEBuild

    ' Get the value from I13 in the Front Sheet of AFE Builder workbook
    valueFromI13 = Workbooks("AFE Builder from SAP HANA").Sheets("Front Sheet").Range("I13").Value
    If valueFromI13 = "" Then
        MsgBox "No value found in I13 of the Front Sheet.", vbCritical
        Exit Sub
    End If

   With ufProgress
        .LabelCaption.Caption = "Retrieving Values - 50" & "% Complete"
        .LabelProgress.Width = 0.5 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    ' Place the value from I13 into M2 of Components Due in the AFE workbook
    destinationSheet.Range("M2").Value = valueFromI13

    ' Get the search value from I11 in the Front Sheet of current workbook
    searchValue = sourceWorkbook.Sheets("Front Sheet").Range("I11").Value
    If searchValue = "" Then
        MsgBox "No value found in I11.", vbCritical
        Exit Sub
    End If

    ' Find last row in source sheet and output row in destination sheet
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "B").End(xlUp).Row
    outputRow = destinationSheet.Cells(destinationSheet.Rows.Count, "A").End(xlUp).Row + 1
    If outputRow < 3 Then outputRow = 3 ' Ensure output starts at row 3

   With ufProgress
        .LabelCaption.Caption = "Retrieving Values - 75" & "% Complete"
        .LabelProgress.Width = 0.75 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    ' Loop through source sheet and copy matching values to destination sheet
    For i = 2 To lastRow
        If sourceSheet.Cells(i, 2).Value = searchValue Then
            ' Place the 'Maint Item' from column E of source sheet into column A of destination sheet
            destinationSheet.Cells(outputRow, "A").Value = sourceSheet.Cells(i, "E").Value
            ' Place the corresponding value from column H into column B of destination sheet
            destinationSheet.Cells(outputRow, "B").Value = sourceSheet.Cells(i, "H").Value

            ' Reduce the range scope for XLookup
            destinationSheet.Cells(outputRow, "C").Value = application.WorksheetFunction.XLookup( _
                sourceSheet.Cells(i, "G").Value, sourceSheet.Range("G2:G" & lastRow), sourceSheet.Range("J2:J" & lastRow))

            'destinationSheet.Cells(outputRow, "D").Value = application.WorksheetFunction.XLookup( _
               ' sourceSheet.Cells(i, "G").Value, sourceSheet.Range("G2:G" & lastRow), sourceSheet.Range("s2:s" & lastRow))

          '  destinationSheet.Cells(outputRow, "G").Value = application.WorksheetFunction.XLookup( _
             '   sourceSheet.Cells(i, "G").Value, sourceSheet.Range("G2:G" & lastRow), sourceSheet.Range("Q2:Q" & lastRow))
            
            ' Increment output row
            outputRow = outputRow + 1
        End If
    Next i

    ' Call the function to remove blank rows in the destination sheet
    RemoveBlankRows destinationSheet
   With ufProgress
        .LabelCaption.Caption = "Retrieving Values - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    ' Re-enable updates after completion
    SetScreenUpdatingSafely True
    SetCalculationSafely xlCalculationAutomatic
    SetEnableEventsSafely True

    ' Notify the user that data has been copied
   ' MsgBox "Data has been successfully copied to AFE workbook.", vbInformation
End Sub
Sub CounterReturn()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim formula As String
    Dim measurePointCell As Range
    
    ' Ensure the AFE workbook is open
    On Error Resume Next  ' In case the workbook is not already open
    Set wb = Workbooks("AFE.xlsx")
    On Error GoTo 0  ' Reset error handling
    
    ' If the workbook is not open, open it
    If wb Is Nothing Then
        Set wb = Workbooks.Open("C:\path\to\your\AFE.xlsx")  ' Update path as necessary
    End If
    
    ' Set the worksheet and table
    Set ws = wb.Sheets("Components due")
    Set tbl = ws.ListObjects("Table1672")  ' Table name
    
    ' Find the last row in the "Measure Point" column (Column C) of the table
    lastRow = tbl.ListRows.Count
    
    ' Define the formula to be used
    formula = "=XLOOKUP([@[Measure Point]],'[AFE Builder from SAP HANA.xlsm]MAINT_PLANS (2)'!C9,'[AFE Builder from SAP HANA.xlsm]MAINT_PLANS (2)'!C15)"
    
    ' Loop through each row of the table starting from row 3 (skip header)
    For i = 1 To lastRow  ' i starts from 1 because ListObject indices start at 1
        ' Check if Measure Point in column C is not blank
        Set measurePointCell = tbl.DataBodyRange.Cells(i, 3)  ' Column C in Table1672
        If Not IsEmpty(measurePointCell.Value) Then
            ' Apply the formula in the corresponding row of column D
            tbl.DataBodyRange.Cells(i, 4).FormulaR1C1 = formula  ' Column D in Table1672
        End If
    Next i
    
    ' Optional: You can call other subroutines after this
    ' Call AnotherSubroutine  ' Uncomment if you have other subs to run
End Sub
Sub CounterReturn2()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim formula As String
    Dim measurePointCell As Range
    
    ' Ensure the AFE workbook is open
    On Error Resume Next  ' In case the workbook is not already open
    Set wb = Workbooks("AFE.xlsx")
    On Error GoTo 0  ' Reset error handling
    
    ' If the workbook is not open, open it
    If wb Is Nothing Then
        Set wb = Workbooks.Open("C:\path\to\your\AFE.xlsx")  ' Update path as necessary
    End If
    
    ' Set the worksheet and table
    Set ws = wb.Sheets("Components due")
    Set tbl = ws.ListObjects("Table1672")  ' Table name
    
    ' Find the last row in the "Measure Point" column (Column C) of the table
    lastRow = tbl.ListRows.Count
    
    ' Define the formula to be used
    formula = "=XLOOKUP([@[ID]] & [@[Maint Item Desc]], '[AFE Builder from SAP HANA.xlsm]MAINT_PLANS (2)'!C26, '[AFE Builder from SAP HANA.xlsm]MAINT_PLANS (2)'!C22)"
    
    ' Loop through each row of the table starting from row 3 (skip header)
    For i = 1 To lastRow  ' i starts from 1 because ListObject indices start at 1
        ' Check if Measure Point in column C is not blank
        Set measurePointCell = tbl.DataBodyRange.Cells(i, 3)  ' Column C in Table1672
        If Not IsEmpty(measurePointCell.Value) Then
            ' Apply the formula in the corresponding row of column D
            tbl.DataBodyRange.Cells(i, 7).FormulaR1C1 = formula  ' Column G in Table1672
        End If
    Next i
    
    ' Optional: You can call other subroutines after this
    ' Call AnotherSubroutine  ' Uncomment if you have other subs to run
End Sub





Sub RemoveBlankRows(sheet As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    ' Find the last row with data in column B
    lastRow = sheet.Cells(sheet.Rows.Count, "B").End(xlUp).Row
    
    ' Loop from LastRow to row 3 (to remove rows from 3 downwards)
    For i = lastRow To 3 Step -1
        cellValue = sheet.Cells(i, "B").Value  ' Check column B
        
        ' If the cell in column B is truly blank (ignoring spaces, line breaks, etc.)
        If IsBlankOrEmpty(cellValue) Then
            sheet.Rows(i).Delete
        End If
    Next i
End Sub

' Helper function to check for truly blank cells
Function IsBlankOrEmpty(cellValue As String) As Boolean
    ' Trim spaces and check for empty string or only whitespace
    If Len(Trim(cellValue)) = 0 Then
        IsBlankOrEmpty = True
    Else
        IsBlankOrEmpty = False
    End If
End Function

Sub Sort_Order()
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim table As ListObject
    Dim columnIndex As Integer

    ' Set the destination workbook and sheet
    Set destinationWorkbook = Workbooks("AFE.xlsx")
    Set destinationSheet = destinationWorkbook.Sheets("Components Due")

    ' Set the table (Table1672)
    On Error Resume Next
    Set table = destinationSheet.ListObjects("Table1672")
    On Error GoTo 0
    If table Is Nothing Then
        MsgBox "'Table1672' not found.", vbCritical
        Exit Sub
    End If

    ' Find the column index for "Comp Hrs Remaining @ O/H"
    columnIndex = table.ListColumns("Comp Hrs Remaining @ O/H").Index

    ' Sort the table based on the "Comp Hrs Remaining @ O/H" column (ascending)
    table.Sort.SortFields.Clear
    table.Sort.SortFields.Add Key:=table.ListColumns(columnIndex).Range, order:=xlAscending
    table.Sort.Apply

    ' Copy and paste the values of column D (entire column D)
    With destinationSheet
        .Range("D:D").Copy
        .Range("D1").PasteSpecial Paste:=xlPasteValues
    End With

    ' Optionally clear the clipboard (to remove the "marching ants" effect)
    application.CutCopyMode = False
End Sub





Sub ClearTable1672ColI()
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim table As ListObject
    Dim columnIndex As Integer
    Dim lastRow As Long

    ' Set the destination workbook and sheet
    Set destinationWorkbook = Workbooks("AFE.xlsx")
    Set destinationSheet = destinationWorkbook.Sheets("Components Due")

    ' Set the table (Table1672)
    Set table = destinationSheet.ListObjects("Table1672")

    ' Find the column index for "Parts Cost" in Table1672
    columnIndex = table.ListColumns("Component & Parts Cost").Index

    ' Find the last row of data in Table1672
    lastRow = table.ListRows.Count

    ' Clear the cells in "Parts Cost" column starting from row 1
    table.ListColumns(columnIndex).DataBodyRange.Cells(1, 1).Resize(lastRow, 1).ClearContents
End Sub







