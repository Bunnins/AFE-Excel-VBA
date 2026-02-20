Attribute VB_Name = "Components"

Sub ExtractMatchingItems()
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim taskListRepSheet As Worksheet
    Dim wsData As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTaskList As Long
    Dim i As Long, r As Long
    Dim resultRow As Long, resultRow2 As Long
    Dim searchValue As String
    Dim roundedProgress As Double
    Dim dictIDs As Object
    Dim dictRangeStart As Object
    Dim dictRangeEnd As Object
    Dim dictUniquePairs As Object
    Dim idArr As Variant
    Dim taskArr As Variant
    Dim outputArr() As Variant
    Dim outputCount As Long
    Dim rangeStart As Long, rangeEnd As Long
    Dim currentKey As String
    Dim startRow As Long

    SetScreenUpdatingSafely False

    With ufProgress
        .LabelCaption.Caption = "Retrieving Tasklist Data - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents

    On Error GoTo ErrorHandler

    Set sourceWorkbook = Workbooks("AFE.xlsx")
    Set sourceWorksheet = sourceWorkbook.Sheets("Components Due")
    Set taskListRepSheet = Workbooks("AFE Builder from SAP HANA").Sheets("TASK_LIST_REP")
    Set wsData = Worksheets("Data")

    lastRowSource = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "H").End(xlUp).Row
    If lastRowSource < 3 Then Exit Sub

    idArr = sourceWorksheet.Range("A3:A" & lastRowSource).Value2

    Set dictIDs = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(idArr, 1)
        searchValue = CStr(idArr(i, 1))
        If Len(searchValue) > 0 Then
            If Not dictIDs.Exists(searchValue) Then dictIDs.Add searchValue, 1
        End If
    Next i

    wsData.Columns("A:A").ClearContents
    resultRow = 1
    If dictIDs.Count > 0 Then
        Dim keys As Variant
        keys = dictIDs.Keys
        For i = LBound(keys) To UBound(keys)
            wsData.Cells(resultRow, 1).Value = keys(i)
            resultRow = resultRow + 1
        Next i
    End If

    wsData.Sort.SortFields.Clear
    wsData.Sort.SortFields.Add Key:=wsData.Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsData.Sort
        .SetRange wsData.UsedRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    lastRowTaskList = taskListRepSheet.Cells(taskListRepSheet.Rows.Count, "H").End(xlUp).Row
    If lastRowTaskList < 3 Or resultRow = 1 Then GoTo Finalize

    taskArr = taskListRepSheet.Range("H3:P" & lastRowTaskList).Value2
    Set dictRangeStart = CreateObject("Scripting.Dictionary")
    Set dictRangeEnd = CreateObject("Scripting.Dictionary")

    currentKey = ""
    For r = 1 To UBound(taskArr, 1)
        searchValue = CStr(taskArr(r, 1))
        If Len(searchValue) > 0 Then
            If searchValue <> currentKey Then
                If Len(currentKey) > 0 Then dictRangeEnd(currentKey) = r - 1
                currentKey = searchValue
                dictRangeStart(searchValue) = r
                dictRangeEnd(searchValue) = r
            End If
        End If
    Next r
    If Len(currentKey) > 0 Then dictRangeEnd(currentKey) = UBound(taskArr, 1)

    Set dictUniquePairs = CreateObject("Scripting.Dictionary")
    ReDim outputArr(1 To 9, 1 To 1)
    outputCount = 0

    resultRow2 = 1
    For i = 1 To resultRow - 1
        searchValue = CStr(wsData.Cells(i, 1).Value2)
        If dictRangeStart.Exists(searchValue) Then
            rangeStart = CLng(dictRangeStart(searchValue))
            rangeEnd = CLng(dictRangeEnd(searchValue))
            For r = rangeStart To rangeEnd
                Dim pairKey As String
                pairKey = CStr(taskArr(r, 1)) & CStr(taskArr(r, 9))
                If Not dictUniquePairs.Exists(pairKey) Then
                    dictUniquePairs.Add pairKey, 1
                    outputCount = outputCount + 1
                    If outputCount > UBound(outputArr, 2) Then ReDim Preserve outputArr(1 To 9, 1 To outputCount)
                    Dim c As Long
                    For c = 1 To 9
                        outputArr(c, outputCount) = taskArr(r, c)
                    Next c
                End If
            Next r
        End If

        With ufProgress
            .LabelCaption.Caption = "Retrieving Tasklist Data - " & Format((i / resultRow) * 100, "0") & "% Complete"
            roundedProgress = Round((i / resultRow), 2)
            .LabelProgress.Width = roundedProgress * (.FrameProgress.Width)
            .Repaint
        End With
        DoEvents
    Next i

    wsData.Range("E:M").ClearContents
    If outputCount > 0 Then
        Dim writeArr() As Variant
        ReDim writeArr(1 To outputCount, 1 To 9)
        Dim rr As Long, cc As Long
        For rr = 1 To outputCount
            For cc = 1 To 9
                writeArr(rr, cc) = outputArr(cc, rr)
            Next cc
        Next rr
        wsData.Range("E1").Resize(outputCount, 9).Value2 = writeArr
    End If

Finalize:
    startRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If wsData.Cells(wsData.Rows.Count, "M").End(xlUp).Row > startRow Then
        startRow = wsData.Cells(wsData.Rows.Count, "M").End(xlUp).Row
    End If
    If startRow < 1 Then startRow = 1

    wsData.Columns("F").Insert Shift:=xlToRight
    wsData.Range("F1:F" & startRow).Value2 = wsData.Range("N1:N" & startRow).Value2
    wsData.Columns("I:N").Delete

    With ufProgress
        .LabelCaption.Caption = "Retrieving Tasklist Data - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & "Error Number: " & Err.Number, vbCritical
End Sub

Sub ListOccurrencesInLabourCosting()
    Dim wsLabourCosting As Worksheet
    Dim wsLCData As Worksheet
    Dim tbl As ListObject
    Dim lastRowData As Long
    Dim wbAFEBuilder As Workbook
    Dim wbAFE As Workbook
    Dim wsAFE As Worksheet
    Dim wsComponentsDue As Worksheet
    Dim tblData As Variant
    Dim compData As Variant
    Dim outputArr() As Variant
    Dim outputCount As Long
    Dim i As Long, r As Long
    Dim searchValue As String
    Dim fullSearchValue As String
    Dim dictIndex As Object
    Dim idx As Collection
    Dim idxGroup As Long, idxVendor As Long, idxWork As Long
    Dim idxActivity As Long, idxCost As Long, idxResource As Long

    SetScreenUpdatingSafely False

    On Error Resume Next
    Set wbAFEBuilder = Workbooks("AFE builder from SAP HANA")
    On Error GoTo 0
    If wbAFEBuilder Is Nothing Then
        MsgBox "The 'AFE builder' workbook is not open."
        Exit Sub
    End If

    On Error Resume Next
    Set wbAFE = Workbooks("AFE")
    On Error GoTo 0
    If wbAFE Is Nothing Then
        MsgBox "The 'AFE' workbook is not open."
        Exit Sub
    End If

    Set wsLabourCosting = wbAFEBuilder.Sheets("Labour Costing")
    Set wsLCData = ThisWorkbook.Sheets("LC Data")
    Set wsComponentsDue = wbAFE.Sheets("Components Due")

    On Error Resume Next
    Set tbl = wsLabourCosting.ListObjects("TASK_LIST_REP__2")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Table 'TASK_LIST_REP__2' not found in the 'Labour Costing' worksheet."
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    idxGroup = tbl.ListColumns("Group and Group Counter").Index
    idxVendor = tbl.ListColumns("PLPO-LIFNR Vendor Desc").Index
    idxWork = tbl.ListColumns("PLPO-ARBEI Work (Attr)").Index
    idxActivity = tbl.ListColumns("PLPO-LARNT Activity Type").Index
    idxCost = tbl.ListColumns("PLPO-PEINH Net Price (Attr)").Index
    On Error Resume Next
    idxResource = tbl.ListColumns("CRHD-ARBPL Work Center/Resource Desc").Index
    On Error GoTo 0
    If idxResource = 0 Then
        MsgBox "'CRHD-ARBPL Work Center/Resource Desc' column not found."
        idxResource = idxVendor
    End If

    tblData = tbl.DataBodyRange.Value2

    Set dictIndex = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(tblData, 1)
        searchValue = CStr(tblData(r, idxGroup))
        If Len(searchValue) > 0 Then
            If Not dictIndex.Exists(searchValue) Then Set dictIndex(searchValue) = New Collection
            dictIndex(searchValue).Add r
        End If
    Next r

    lastRowData = wsComponentsDue.Cells(wsComponentsDue.Rows.Count, "A").End(xlUp).Row
    wsLCData.Cells.Clear
    wsLCData.Cells(1, 1).Value = "Group and Group Counter"
    wsLCData.Cells(1, 2).Value = "Vendor Description"
    wsLCData.Cells(1, 3).Value = "Work Hours"
    wsLCData.Cells(1, 4).Value = "Activity Type"
    wsLCData.Cells(1, 5).Value = "PMEX Cost"
    wsLCData.Cells(1, 6).Value = "Vendor Name"

    If lastRowData >= 3 Then
        compData = wsComponentsDue.Range("A3:B" & lastRowData).Value2
        ReDim outputArr(1 To 6, 1 To 1)
        outputCount = 0

        For i = 1 To UBound(compData, 1)
            fullSearchValue = CStr(compData(i, 1)) & " " & CStr(compData(i, 2))
            searchValue = Left(CStr(compData(i, 1)), 10)
            If dictIndex.Exists(searchValue) Then
                Set idx = dictIndex(searchValue)
                Dim n As Long
                For n = 1 To idx.Count
                    r = idx(n)
                    outputCount = outputCount + 1
                    If outputCount > UBound(outputArr, 2) Then ReDim Preserve outputArr(1 To 6, 1 To outputCount)
                    outputArr(1, outputCount) = fullSearchValue
                    outputArr(2, outputCount) = tblData(r, idxVendor)
                    outputArr(3, outputCount) = tblData(r, idxWork)
                    outputArr(4, outputCount) = tblData(r, idxActivity)
                    outputArr(5, outputCount) = tblData(r, idxCost)
                    outputArr(6, outputCount) = tblData(r, idxResource)
                Next n
            End If
        Next i

        If outputCount > 0 Then
            Dim outWrite() As Variant
            ReDim outWrite(1 To outputCount, 1 To 6)
            Dim wr As Long, wc As Long
            For wr = 1 To outputCount
                For wc = 1 To 6
                    outWrite(wr, wc) = outputArr(wc, wr)
                Next wc
            Next wr
            wsLCData.Range("A2").Resize(outputCount, 6).Value2 = outWrite
        End If
    End If

    Set wsAFE = wbAFE.Sheets("Sheet1")
    wsAFE.Name = "Labour Data"
    wsAFE.Cells.Clear
    wsAFE.Range("A1").Resize(wsLCData.UsedRange.Rows.Count, wsLCData.UsedRange.Columns.Count).Value2 = wsLCData.UsedRange.Value2
End Sub


Sub SummarizeLabourData()
    ' Declare variables
    Dim wsAFE As Worksheet
    Dim wsComponents As Worksheet
    Dim lastRowAFE As Long
    Dim lastRowComponents As Long
    Dim codeDict As Object
    Dim i As Long
    Dim code As Variant
    Dim hours As Double
    Dim outputRow As Long
    Dim lookupValue As String
    Dim matchRow As Long
    Dim requiredValue As Variant
    Dim rowNum As Long
    
    ' Set reference to the "Labour Data" and "Components Due" sheets
    Set wsAFE = Workbooks("AFE").Sheets("Labour Data")
    Set wsComponents = Workbooks("AFE").Sheets("Components Due")
    
    ' Find the last row of data in Labour Data sheet
    lastRowAFE = wsAFE.Cells(wsAFE.Rows.Count, "A").End(xlUp).Row
    
    ' Find the last row of data in Components Due sheet
    lastRowComponents = wsComponents.Cells(wsComponents.Rows.Count, "A").End(xlUp).Row
    
    ' Create a dictionary to store codes and summed hours
    Set codeDict = CreateObject("Scripting.Dictionary")

    ' Loop through each row in column F (now Vendor data) and add corresponding hours in column H (Actual Hours) to the dictionary
    For i = 2 To lastRowAFE ' Start from row 2 to skip headers
        code = wsAFE.Cells(i, 6).Value ' Vendor code in column F (previously was column B)
        hours = wsAFE.Cells(i, 8).Value ' Actual Hours in column H (updated from column C)

        If codeDict.Exists(code) Then
            ' If the code already exists in the dictionary, add the hours to the existing value
            codeDict(code) = codeDict(code) + hours
        Else
            ' If the code doesn't exist in the dictionary, add it with the initial hours value
            codeDict.Add code, hours
        End If
    Next i

    ' Now, write the unique codes and summed hours to columns N and O
    outputRow = 2 ' Start writing data in row 2 of columns N and O (row 1 will have headers)

    ' Add headers for the new summarized table
    wsAFE.Cells(1, 14).Value = "Vendor" ' Column N (was "Work Centre")
    wsAFE.Cells(1, 15).Value = "Total Hours" ' Column O (unchanged)
    
    ' Add new headers for columns G, H, and I
    wsAFE.Cells(1, 7).Value = "Required" ' Column G
    wsAFE.Cells(1, 8).Value = "Actual Hours" ' Column H
    wsAFE.Cells(1, 9).Value = "Actual PMEX Value" ' Column I for Vendor Description

    ' Loop through the dictionary to output the data
    For Each code In codeDict.Keys
        wsAFE.Cells(outputRow, 14).Value = code ' Write the unique code in column N (Vendor)
        wsAFE.Cells(outputRow, 15).Value = codeDict(code) ' Write the summed hours in column O (Total Hours)
        
        ' Leave the "Required" column (G) and "Actual Hours" column (H) blank
        wsAFE.Cells(outputRow, 7).Value = "" ' Placeholder for "Required" in column G
        wsAFE.Cells(outputRow, 8).Value = "" ' Leave "Actual Hours" blank in column H
        wsAFE.Cells(outputRow, 9).Value = "" ' Leave Vendor Description blank in column I

        outputRow = outputRow + 1
    Next code

    ' Optional: Format the results as a table
    Dim tblRange As Range
    Set tblRange = wsAFE.Range("N1:O" & outputRow - 1)
    wsAFE.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "SummaryTable"
    
    ' Perform VLOOKUP for values in column A and populate results in column G starting from G2
    rowNum = 2 ' Start from row 2
    
    ' Loop through each row in column A until a blank cell is encountered
    For i = 2 To lastRowAFE ' Start from row 2 to skip headers
        lookupValue = wsAFE.Cells(i, 1).Value ' Concatenated value in column A
        
        ' Extract the first 10 characters and the remaining part
        Dim firstPart As String
        Dim secondPart As String
        firstPart = Left(lookupValue, 10)
        secondPart = Trim(Mid(lookupValue, 12)) ' Adjusted to skip the space
        
        ' Perform the lookup in Components Due sheet
        matchRow = 0
        For j = 2 To lastRowComponents ' Start from row 2 to skip headers
            If wsComponents.Cells(j, 1).Value = firstPart And wsComponents.Cells(j, 2).Value = secondPart Then
                matchRow = j
                Exit For
            End If
        Next j
        
        ' If a match is found, get the value from column H (Required)
        If matchRow > 0 Then
            requiredValue = wsComponents.Cells(matchRow, 8).Value ' Column H in Components Due
            If requiredValue = 1 Then
                wsAFE.Cells(i, 7).Value = "Yes" ' Write "Yes" to column G in Labour Data
            ElseIf requiredValue = -1 Then
                wsAFE.Cells(i, 7).Value = "No" ' Write "No" to column G in Labour Data
            Else
                wsAFE.Cells(i, 7).Value = "No Match" ' Indicate no match found
            End If
        Else
            wsAFE.Cells(i, 7).Value = "No Match" ' Indicate no match found
        End If
    Next i
    
    ' Insert the formula for "Actual Hours" in column H
    For rowNum = 2 To lastRowAFE
        actualHoursFormula = "=IF(G" & rowNum & "=""Yes"",C" & rowNum & ",0)"
        wsAFE.Cells(rowNum, 8).formula = actualHoursFormula ' Insert the formula into column H
    Next rowNum
    
    ' Insert the formula for "Actual PMEX Value" in column I
    For rowNum = 2 To lastRowAFE
        wsAFE.Cells(rowNum, 9).formula = "=IF(G" & rowNum & "=""Yes"",E" & rowNum & ", """")" ' Insert the formula into column I
    Next rowNum
    
    ' Now, insert the formula for "Total Hours" in column O, pulling from the summary of Actual Hours (column H)
    For rowNum = 2 To outputRow - 1
        wsAFE.Cells(rowNum, 15).formula = "=SUMIF(F:F, N" & rowNum & ", H:H)" ' Sum Actual Hours based on Vendor in Column F
    Next rowNum

    ' After populating the summary, clear cells below the table in Column O
    lastRowSummary = outputRow - 1 ' Adjust to the last populated row of the summary table

    ' Clear the cells below the table in Column O
    wsAFE.Range("O" & lastRowSummary + 1 & ":O" & wsAFE.Rows.Count).ClearContents

    ' Auto-fit columns A to O to adjust the width based on the content
    wsAFE.Columns("A:O").AutoFit
End Sub

Sub SummarizeLabourData2()
    ' Declare variables
    Dim wsAFE As Worksheet
    Dim lastRow As Long
    Dim codeDict As Object
    Dim i As Long
    Dim code As Variant
    Dim pmexCost As Variant ' Change to Variant to handle different types (strings, numbers, etc.)
    Dim outputRow As Long
    Dim tblRange As Range
    Dim summaryTable As ListObject
    Dim lastSummaryRow As Long
    Dim vendorDescriptionFormula As String
    Dim pmexCostFormula As String

    ' Set reference to the "Labour Data" sheet in AFE workbook
    Set wsAFE = Workbooks("AFE").Sheets("Labour Data")

    ' Find the last row of data in column B (Vendor Description) and column I (PMEX Cost)
    lastRow = wsAFE.Cells(wsAFE.Rows.Count, "B").End(xlUp).Row

    ' Create a dictionary to store Vendor Description and summed PMEX Cost
    Set codeDict = CreateObject("Scripting.Dictionary")

    ' Loop through each row in column B (Vendor Description) and add corresponding PMEX Cost in column I to the dictionary
    For i = 2 To lastRow ' Start from row 2 to skip headers
        code = wsAFE.Cells(i, 2).Value ' Vendor Description in column B
        pmexCost = wsAFE.Cells(i, 9).Value ' PMEX Cost in column I

        ' Ensure that pmexCost is a number (to avoid type mismatch error)
        If IsNumeric(pmexCost) Then
            pmexCost = CDbl(pmexCost) ' Ensure it is treated as a Double
        Else
            pmexCost = 0 ' If it's not numeric, treat it as 0 (you can adjust this logic)
        End If
        
        ' Only add to dictionary if PMEX Cost is greater than 0 and Vendor Description is not blank
        If pmexCost > 0 And code <> "" Then
            If codeDict.Exists(code) Then
                ' If the Vendor Description already exists in the dictionary, add the PMEX Cost to the existing value
                codeDict(code) = codeDict(code) + pmexCost
            Else
                ' If the Vendor Description doesn't exist in the dictionary, add it with the initial PMEX Cost
                codeDict.Add code, pmexCost
            End If
        End If
    Next i

    ' Now, clear old summary table if exists (in columns Q and R)
    On Error Resume Next
    Set summaryTable = wsAFE.ListObjects("SummaryTable2")
    On Error GoTo 0
    If Not summaryTable Is Nothing Then
        summaryTable.Delete ' Delete the old table if it exists
    End If

    ' Add headers for the new summarized table
    wsAFE.Cells(1, 17).Value = "Vendor Description" ' Column Q
    wsAFE.Cells(1, 18).Value = "PMEX Cost" ' Column R

    ' Start writing data in row 2 of columns Q and R (row 1 will have headers)
    outputRow = 2

    ' Loop through the dictionary to output the data
    For Each code In codeDict.Keys
        wsAFE.Cells(outputRow, 17).Value = code ' Write the unique Vendor Description in column Q
        
        ' Insert SUMIF formula for PMEX Cost in column R
        pmexCostFormula = "=SUMIF(B:B, Q" & outputRow & ", I:I)" ' Sum PMEX Cost based on Vendor Description in column B
        wsAFE.Cells(outputRow, 18).formula = pmexCostFormula

        outputRow = outputRow + 1
    Next code

    ' Now create the dynamic summary table
    Set tblRange = wsAFE.Range("Q1:R" & outputRow - 1)
    Set summaryTable = wsAFE.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    summaryTable.Name = "SummaryTable2"

    ' Auto-fit columns Q and R to adjust the width based on the content
    wsAFE.Columns("Q:R").AutoFit
End Sub






Sub CopyDataWorksheet()
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook

    ' Set the source workbook (this workbook)
    Set sourceWorkbook = ThisWorkbook

    ' Set the destination workbook
    Set destinationWorkbook = Workbooks("AFE.xlsx") ' Ensure this workbook is open

    ' Check if the "Data" sheet exists in the source workbook
    On Error Resume Next
    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceWorkbook.Sheets("Data")
    On Error GoTo 0

    If sourceSheet Is Nothing Then
        MsgBox "Worksheet 'Data' not found in source workbook.", vbExclamation
        Exit Sub
    End If

    ' Copy the "Data" worksheet to the destination workbook
    sourceSheet.Copy After:=destinationWorkbook.Sheets(destinationWorkbook.Sheets.Count)

End Sub


Sub ApplyXLOOKUPFormula()
    Dim resultWorksheet As Worksheet
    Dim lastRow As Long
    Dim ws As Worksheet
    
    ' Reference the "Data" worksheet in the destination workbook
    Set resultWorksheet = Worksheets("Data")
    
    ' Determine the last row in column E
    lastRow = resultWorksheet.Cells(resultWorksheet.Rows.Count, "E").End(xlUp).Row
    
    ' Step 1: Apply the XLOOKUP formula to column I starting from I1
    resultWorksheet.Range("I1:I" & lastRow).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(RC[-3],'Parts Costing'!C[-8],'Parts Costing'!C[-3]), 0) * RC[-2]"
    
    ' Step 2: Format the cells in column I as currency
    resultWorksheet.Range("I1:I" & lastRow).NumberFormat = "$#,##0.00"
    
    ' Step 3: Auto fit columns in the result sheet (optional but recommended)
    resultWorksheet.Columns("I").AutoFit
    
    ' Step 4: Unlock the workbook and paste values in "Data" sheet
    ' Unprotect the workbook (replace "yourpassword" with the actual password if necessary)
    ThisWorkbook.Unprotect "PlanExRavMB"
    
    ' Copy the formulas in column I and paste them as values
    resultWorksheet.Range("I1:I" & lastRow).Copy
    resultWorksheet.Range("I1:I" & lastRow).PasteSpecial Paste:=xlPasteValues
    
        ' Copy the formulas in column J and paste them as values
    resultWorksheet.Range("J1:J" & lastRow).Copy
    resultWorksheet.Range("J1:J" & lastRow).PasteSpecial Paste:=xlPasteValues
    
    ' Step 5: Protect the workbook again (optional: set a password for re-protection)
    ThisWorkbook.Protect "PlanExRavMB"
    
    ' Step 6: Auto fit columns in the result sheet (optional but recommended)
    resultWorksheet.Columns("I").AutoFit
    

End Sub




Sub ApplyXLOOKUPForColumnF()
    Dim resultWorksheet As Worksheet
    Dim lastRow As Long
    
    ' Reference the "GroupCounterResults" worksheet in the destination workbook
     Set resultWorksheet = Worksheets("Data")
    
    ' Determine the last row in column E
    lastRow = resultWorksheet.Cells(resultWorksheet.Rows.Count, "E").End(xlUp).Row
    
    ' Step 1: Apply the XLOOKUP formula to column F starting from F2
    resultWorksheet.Range("J1:J" & lastRow).FormulaR1C1 = _
        "=IFERROR(XLOOKUP(RC[-5],'[AFE Builder from SAP HANA.xlsm]Data2'!C5,'[AFE Builder from SAP HANA.xlsm]Data2'!C7), """")"
    
    ' Step 2: Auto fit columns in the result sheet (optional but recommended)
   ' resultWorksheet.Columns("A:F").AutoFit
    
    ' Notify the user
   ' MsgBox "XLOOKUP formula has been applied to column F of GroupCounterResults."
End Sub



Sub test()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim srcArr As Variant
    Dim outArr() As Variant
    Dim markRows() As Long
    Dim outRow As Long
    Dim i As Long, c As Long
    Dim currentOrder As String
    Dim workHoursTotal As Double
    Dim roundedProgress As Double

    SetScreenUpdatingSafely False
    Set ws = ThisWorkbook.Sheets("Data")

    With ufProgress
        .LabelCaption.Caption = "Retrieving Material Pricing - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents

    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then lastCol = 9

    srcArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2
    ReDim outArr(1 To (lastRow * 2), 1 To lastCol)
    ReDim markRows(1 To (lastRow * 2))

    outRow = 0
    currentOrder = ""
    workHoursTotal = 0

    For i = 1 To lastRow
        If CStr(srcArr(i, 5)) <> currentOrder Then
            If Len(currentOrder) > 0 Then
                outRow = outRow + 1
                outArr(outRow, 8) = currentOrder
                outArr(outRow, 9) = workHoursTotal
                markRows(outRow) = 1
            End If
            currentOrder = CStr(srcArr(i, 5))
            workHoursTotal = 0
        End If

        workHoursTotal = workHoursTotal + Val(srcArr(i, 9))

        outRow = outRow + 1
        For c = 1 To lastCol
            outArr(outRow, c) = srcArr(i, c)
        Next c

        With ufProgress
            .LabelCaption.Caption = "Retrieving Material Pricing  - " & Format((i / lastRow) * 100, "0") & "% Complete"
            roundedProgress = Round((i / lastRow), 2)
            .LabelProgress.Width = roundedProgress * (.FrameProgress.Width)
            .Repaint
        End With
        DoEvents
    Next i

    If Len(currentOrder) > 0 Then
        outRow = outRow + 1
        outArr(outRow, 8) = currentOrder
        outArr(outRow, 9) = workHoursTotal
        markRows(outRow) = 1
    End If

    ws.Cells.Clear
    Dim finalArr() As Variant
    ReDim finalArr(1 To outRow, 1 To lastCol)
    Dim fr As Long, fc As Long
    For fr = 1 To outRow
        For fc = 1 To lastCol
            finalArr(fr, fc) = outArr(fr, fc)
        Next fc
    Next fr
    ws.Range(ws.Cells(1, 1), ws.Cells(outRow, lastCol)).Value2 = finalArr

    For i = 1 To outRow
        If markRows(i) = 1 Then ws.Rows(i).Interior.Color = RGB(255, 255, 153)
    Next i

    With ufProgress
        .LabelCaption.Caption = "Retrieving Material Pricing - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint
    End With
    DoEvents
End Sub

Sub ApplyXLOOKUPDownSheet_AFE()
    ' Declare necessary variables
    Dim afeWorkbook As Workbook
    Dim componentsDueSheet As Worksheet
    Dim componentsCost As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim xLookupFormula As String
    Dim requiredValue As Variant
    
    ' Ensure "AFE.xlsx" workbook is open
    On Error Resume Next
    Set afeWorkbook = Workbooks("AFE.xlsx") ' Ensure this matches the name of the AFE workbook
    On Error GoTo 0
    If afeWorkbook Is Nothing Then
        MsgBox "'AFE.xlsx' workbook is not open!", vbCritical
        Exit Sub
    End If
    
    ' Set the "Components Due" sheet in the "AFE" workbook
    Set componentsDueSheet = afeWorkbook.Sheets("Components Due")
    Set componentsCost = afeWorkbook.Sheets("Data")
    
    ' Find the last row with data in column A (the lookup column) in "Components Due" sheet
    lastRow = componentsDueSheet.Cells(componentsDueSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row starting from row 3 until the last row with data in column A
    For currentRow = 3 To lastRow
        ' Check if Column H ("Required") value is 1
        requiredValue = componentsDueSheet.Cells(currentRow, "H").Value
        
        'If requiredValue = 1 Then
            ' Formula for XLOOKUP (adjusted for the current row)
xLookupFormula = "=IF(H" & currentRow & "=1, IFERROR(VLOOKUP([@ID], Data!H:I, 2, 0), """"), """")"

' Set the formula in the current row of column I in the "Components Due" sheet
componentsDueSheet.Cells(currentRow, 9).formula = xLookupFormula
            
            ' Apply Currency format to the cell in column I (XLOOKUP result)
            componentsDueSheet.Cells(currentRow, 9).NumberFormat = "$#,##0.00"
       ' Else
            ' If Column H is not 1, clear any existing value or formula in Column I for the current row
           ' componentsDueSheet.Cells(currentRow, 9).ClearContents
       ' End If
    Next currentRow
    
    ' Optionally clear the clipboard if any copy/paste operation occurred
    application.CutCopyMode = False
End Sub

Sub LabourFinal()
    ' Declare variables
    Dim wsAFE As Worksheet
    Dim wsLabour As Worksheet
    Dim wsSAP As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim lastLabourRow As Long
    Dim tblRange As Range
    Dim cell As Range
    Dim headerCell As Range
    Dim rowNum As Long
    Dim sapValue As Variant
    Dim activityType As String
    Dim numericValue As Double
    Dim regex As Object
    Dim matches As Object

    ' Set reference to the "Labour Data" sheet in the AFE workbook
    Set wsAFE = Workbooks("AFE").Sheets("Labour Data") ' Reference to Labour Data

    ' Set reference to the "AFE Builder From SAP HANA" sheet
    Set wsSAP = Workbooks("AFE Builder From SAP HANA").Sheets("Front Sheet") ' Reference to the SAP sheet

    ' Create a new sheet named "Labour" and delete any existing one with that name
    On Error Resume Next
    Set wsLabour = Workbooks("AFE").Sheets("Labour")
    On Error GoTo 0

    ' If the "Labour" sheet exists, delete it
    If Not wsLabour Is Nothing Then
        application.DisplayAlerts = False ' Turn off alerts to prevent confirmation message
        wsLabour.Delete
        application.DisplayAlerts = True ' Turn alerts back on
    End If

    ' Add a new sheet at the end of the sheets in the workbook
    Set wsLabour = Workbooks("AFE").Sheets.Add(After:=wsAFE) ' Add sheet after the Labour Data sheet
    wsLabour.Name = "Labour"

    ' Set the background color of the entire sheet to white
    wsLabour.Cells.Interior.Color = RGB(255, 255, 255) ' White background

    ' Find the last row of data in column N (assuming the table is in columns N and O)
    lastRow = wsAFE.Cells(wsAFE.Rows.Count, "N").End(xlUp).Row

    ' Set the range for the entire table data (including headers) in columns N and O
    Set dataRange = wsAFE.Range("N1:O" & lastRow) ' Include header row N1:O1

    ' Copy the data values and paste them into the "Labour" sheet (no formulas)
    wsLabour.Range("A1").Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

    ' Add new headers to columns C, D, and E
    wsLabour.Cells(1, 3).Value = "Activity Type" ' Column C
    wsLabour.Cells(1, 4).Value = "Rate P/H" ' Column D
    wsLabour.Cells(1, 5).Value = "Total" ' Column E

    ' Format the header row with "Aptos Display" font, white text, black background, only for cells with text
    For Each headerCell In wsLabour.Rows(1).Cells
        If headerCell.Value <> "" Then ' Check if the cell has text
            headerCell.Font.Name = "Aptos Display"  ' Header font
            headerCell.Font.Color = RGB(255, 255, 255) ' White font color
            headerCell.Interior.Color = RGB(0, 0, 0) ' Black background
            headerCell.Font.Size = 12 ' Adjust this size as needed
            headerCell.Font.Bold = True
        End If
    Next headerCell

    ' Format the table values: "Aptos Narrow" font, black text, white background
    With wsLabour.Range("A2:E" & lastRow)
        .Font.Name = "Aptos Narrow" ' Table values font
        .Font.Color = RGB(0, 0, 0) ' Black font color
        .Font.Size = 10 ' Adjust size as needed
        .Interior.Color = RGB(255, 255, 255) ' White background
    End With

    ' Convert the text to uppercase in the table (Rows 2 and below)
    For Each cell In wsLabour.Range("A2:E" & lastRow)
        cell.Value = UCase(cell.Value) ' Convert cell value to uppercase
    Next cell

    ' Apply solid black borders around the table and all cells
    With wsLabour.Range("A1:E" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black borders
        .TintAndShade = 0
        .Weight = xlThin ' Adjust weight as needed
    End With

    ' Convert the range into an Excel table for better styling
    Set tblRange = wsLabour.Range("A1:E" & lastRow)
    wsLabour.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "LabourTable"

    ' Add the XLOOKUP formula to column C (Activity Type)
    For rowNum = 2 To lastRow
        wsLabour.Cells(rowNum, 3).formula = "=XLOOKUP([@Vendor],'Labour Data'!F:F,'Labour Data'!D:D)"
    Next rowNum

    ' Add the XLOOKUP formula to column B (Vendor Information from column O in Labour Data)
    For rowNum = 2 To lastRow
        wsLabour.Cells(rowNum, 2).formula = "=IFERROR(XLOOKUP([@Vendor],'Labour Data'!N:N,'Labour Data'!O:O),"""")"
    Next rowNum

    ' Check if C2 down contains "CRAN-S", "CRAN-L", "RIL", or "REL80"/"REL110"/"REL120", etc.
    For rowNum = 2 To lastRow
        activityType = wsLabour.Cells(rowNum, 3).Value

        ' Check for "CRAN-S", "CRAN-L", "RIL" and return the respective values
        If activityType = "CRAN-S" Then
            ' Get the value from I16 in "AFE Builder From SAP HANA"
            sapValue = wsSAP.Range("I16").Value
            wsLabour.Cells(rowNum, 4).Value = sapValue
        ElseIf activityType = "CRAN-L" Then
            ' Get the value from I15 in "AFE Builder From SAP HANA"
            sapValue = wsSAP.Range("I15").Value
            wsLabour.Cells(rowNum, 4).Value = sapValue
        ElseIf activityType = "RIL" Then
            ' Get the value from I14 in "AFE Builder From SAP HANA"
            sapValue = wsSAP.Range("I14").Value
            wsLabour.Cells(rowNum, 4).Value = sapValue
        ElseIf activityType Like "REL*" Then
            ' Check if the value starts with "REL" and contains numbers
            Set regex = CreateObject("VBScript.RegExp")
            regex.IgnoreCase = True
            regex.Global = True
            regex.Pattern = "\d+" ' Match one or more digits

            ' Extract numbers from the activityType string (e.g., "REL80" -> 80)
            Set matches = regex.Execute(activityType)
            If matches.Count > 0 Then
                numericValue = matches(0).Value
                wsLabour.Cells(rowNum, 4).Value = numericValue ' Return the numeric part
            End If
        Else
            ' Optionally return blank if not a matching type
            wsLabour.Cells(rowNum, 4).ClearContents ' Clear contents if not matching
        End If
    Next rowNum

    ' Format all values in column D as currency (with "$" sign)
    wsLabour.Range("D2:D" & lastRow).NumberFormat = "$#,##0.00"

    ' Format column J as currency (with "$" sign)
    wsLabour.Range("J2:J" & lastRow).NumberFormat = "$#,##0.00"

    ' Add formula in column E (Total) as B2 * D2, until blank cell in column B
    For rowNum = 2 To lastRow
        If wsLabour.Cells(rowNum, 2).Value <> "" Then ' Check if there's a value in column B
            wsLabour.Cells(rowNum, 5).formula = "=B" & rowNum & "*D" & rowNum ' Formula: B * D
        Else
            Exit For ' Exit loop if blank cell is encountered in column B
        End If
    Next rowNum

    ' Now, copy data from columns Q and R in Labour Data to columns I and J in Labour
    ' Find the last row in column Q (Labour Data)
    lastRow = wsAFE.Cells(wsAFE.Rows.Count, "Q").End(xlUp).Row

    ' Set range for columns Q and R in Labour Data
    Set dataRange = wsAFE.Range("Q1:R" & lastRow)

    ' Copy the data to columns I and J in the Labour sheet
    wsLabour.Range("I1").Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

    ' Add headers to columns I and J
    wsLabour.Cells(1, 9).Value = "Vendor Description" ' Column I header
    wsLabour.Cells(1, 10).Value = "PMEX Value" ' Column J header

    ' Format the header row with "Aptos Display" font, white text, black background, only for cells with text
    For Each headerCell In wsLabour.Range("I1:J1").Cells
        If headerCell.Value <> "" Then ' Check if the cell has text
            headerCell.Font.Name = "Aptos Display"  ' Header font
            headerCell.Font.Color = RGB(255, 255, 255) ' White font color
            headerCell.Interior.Color = RGB(0, 0, 0) ' Black background
            headerCell.Font.Size = 12 ' Adjust this size as needed
            headerCell.Font.Bold = True
        End If
    Next headerCell

    ' Format the copied data in columns I and J (same as original data)
    With wsLabour.Range("I2:J" & lastRow)
        .Font.Name = "Aptos Narrow" ' Table values font
        .Font.Color = RGB(0, 0, 0) ' Black font color
        .Font.Size = 10 ' Adjust size as needed
        .Interior.Color = RGB(255, 255, 255) ' White background
    End With

    ' Apply borders to columns I and J
    With wsLabour.Range("I1:J" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black borders
        .TintAndShade = 0
        .Weight = xlThin ' Adjust weight as needed
    End With

    ' Auto-fit columns A through E and I through J
    wsLabour.Columns("A:E").AutoFit
    wsLabour.Columns("I:J").AutoFit

    ' Add the formula in column J if there is a value in column I
    For rowNum = 2 To lastRow
        If wsLabour.Cells(rowNum, 9).Value <> "" Then ' Check if there's a value in Column I
            wsLabour.Cells(rowNum, 10).formula = "=SummaryTable2[@[PMEX Cost]]"
        End If
    Next rowNum

End Sub



Sub SummaryTotals()
    ' Declare variables
    Dim wsLabour As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRowLabour As Long
    Dim i As Long
    Dim insertRow As Long
    Dim lastInsertedRow As Long

    ' Set references to the "Labour" and "Summary" sheets
    Set wsLabour = Workbooks("AFE").Sheets("Labour")
    Set wsSummary = Workbooks("AFE").Sheets("Summary")
    
    ' Find the last row in the "Labour" sheet in column A
    lastRowLabour = wsLabour.Cells(wsLabour.Rows.Count, "A").End(xlUp).Row
    
    ' Set the starting point for inserting rows into "Summary" sheet (A6)
    insertRow = 6
    lastInsertedRow = 6 ' Track the last inserted row for cleanup
    
    ' Loop through each row in the Labour data starting from row 2
    For i = 2 To lastRowLabour
        ' Only copy non-blank values in column A from Labour sheet
        If wsLabour.Cells(i, 1).Value <> "" Then
            ' Insert a new row in "Summary" sheet starting at insertRow
            wsSummary.Rows(insertRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Copy the value from "Labour" A to "Summary" A
            wsSummary.Cells(insertRow, 1).Value = wsLabour.Cells(i, 1).Value
            
            ' Paste formula referencing the corresponding cell in Labour sheet (Column E) to the newly inserted row in Summary (Column C)
            wsSummary.Cells(insertRow, 3).formula = "='Labour'!E" & i

            
            ' Add the text "Price Based on SAP Extract" to column D
            wsSummary.Cells(insertRow, 4).Value = "Site Based Contract"
            
            ' Apply accounting format to column C (Summary sheet)
            wsSummary.Cells(insertRow, 3).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""_);_(@_)"
            
            ' Move to the next row in the "Summary" sheet for the next insert
            lastInsertedRow = insertRow ' Update the last inserted row
            insertRow = insertRow + 1
        End If
    Next i

    ' Remove the extra blank row after the last data entry
    If lastInsertedRow > 6 Then
        wsSummary.Rows(lastInsertedRow + 1).Delete
    End If

    ' Clean up
    application.CutCopyMode = False
    wsSummary.Activate ' Optional: Activate the Summary sheet after completing
End Sub

Sub SummaryPMEXTotals()
    Dim wsSummary As Worksheet
    Dim wsLabour As Worksheet
    Dim pmexCell As Range
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim wb As Workbook
    Dim lastInsertedRow As Long
    Dim borderRange As Range
    Dim pmexRow As Long
    Dim externalWb As Workbook
    
    ' Set reference to the "AFE" workbook
    Set wb = Workbooks("AFE")
    
    ' Set references to the "Labour" and "Summary" sheets in the "AFE" workbook
    Set wsLabour = wb.Sheets("Labour")
    Set wsSummary = wb.Sheets("Summary")
    
    ' Set reference to the external workbook (AFE Builder from SAP HANA)
    On Error Resume Next
    Set externalWb = Workbooks("AFE Builder from SAP HANA.xlsm")
    On Error GoTo 0

    ' Check if the external workbook is open
    If externalWb Is Nothing Then
        MsgBox "The external workbook 'AFE Builder from SAP HANA.xlsm' is not open.", vbExclamation
        Exit Sub
    End If
    
    ' Add formula to cell B1 in the Summary sheet
    wsSummary.Range("B1").formula = "=RIGHT('[AFE Builder from SAP HANA.xlsm]Front Sheet'!I11, LEN('[AFE Builder from SAP HANA.xlsm]Front Sheet'!I11) - 4)"
    wsSummary.Range("B1").Value = wsSummary.Range("B1").Value ' Paste the value (remove the formula)
    
    ' Add formula to cell C1 in the Summary sheet
    wsSummary.Range("C1").formula = "='[AFE Builder from SAP HANA.xlsm]Data2'!A1"
    wsSummary.Range("C1").Value = wsSummary.Range("C1").Value ' Paste the value (remove the formula)
    
    ' Add formula to cell C1 in the Summary sheet
    wsSummary.Range("C2").formula = "='[AFE Builder from SAP HANA.xlsm]Front Sheet'!I12"
    wsSummary.Range("C2").Value = wsSummary.Range("C2").Value ' Paste the value (remove the formula)
    
    ' Find "PMEX COSTS" in column A of the Summary sheet
    Set pmexCell = wsSummary.Range("A:A").Find("PMEX COSTS", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not pmexCell Is Nothing Then
        ' Find the row where "PMEX COSTS" is located
        pmexRow = pmexCell.Row
        
        ' Remove the bottom border of the cell with "PMEX COSTS" (and across to column D)
        With wsSummary.Range("A" & pmexRow & ":D" & pmexRow).Borders(xlEdgeBottom)
            .LineStyle = xlNone
        End With
        
        ' Find the first empty cell under "PMEX COSTS"
        targetRow = pmexRow + 1
        
        ' Get the last row of the Labour sheet in column I
        lastRow = wsLabour.Cells(wsLabour.Rows.Count, "I").End(xlUp).Row
        
        ' Loop through the Labour sheet column I starting from row 2
        For i = 2 To lastRow
            ' Skip blank cells (cells with no value or only spaces)
            If Trim(wsLabour.Cells(i, "I").Value) = "" Then
                GoTo NextIteration
            End If
            
            ' Insert a new row at the target position in the Summary sheet
            wsSummary.Rows(targetRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            
            ' Paste value from Labour sheet to the newly inserted row in the Summary sheet (Column A)
            wsSummary.Cells(targetRow, "A").Value = wsLabour.Cells(i, "I").Value
            
            ' Paste formula referencing the corresponding cell in Labour sheet (Column J) to the newly inserted row in Summary (Column C)
            wsSummary.Cells(targetRow, "C").formula = "='Labour'!J" & i
            
            ' Set the value for Column D to "Price Based on SAP Extract"
            wsSummary.Cells(targetRow, "D").Value = "Price Based on SAP Extract"
            
            ' Track the last inserted row
            lastInsertedRow = targetRow
            
            ' Increment targetRow to the next row for the next value
            targetRow = targetRow + 1
            
NextIteration:
        Next i
        
        ' Add a medium border around the entire group of A4:D4 as a block (not individual cells)
        Set borderRange = wsSummary.Range("A4:D4")
        With borderRange.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With borderRange.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With borderRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With borderRange.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        ' Add a medium border around the entire group of newly inserted rows in columns A to D
        If lastInsertedRow > 6 Then
            Set borderRange = wsSummary.Range("A" & pmexCell.Row + 1 & ":D" & lastInsertedRow)
            With borderRange.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With borderRange.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With borderRange.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
        End If
        
        ' Remove the extra blank row after the last data entry
        If lastInsertedRow > 6 Then
            wsSummary.Rows(lastInsertedRow + 1).Delete
        End If
        
        ' Ensure no cell in column A has Wrap Text enabled
        wsSummary.Range("A4:A" & lastInsertedRow).WrapText = False
        
    Else
        MsgBox "'PMEX COSTS' not found in Summary sheet.", vbExclamation
    End If
End Sub

Sub DataClean()
    Dim wsLabourData As Worksheet
    Dim wsPartsList As Worksheet
    Dim wb As Workbook
    Dim lastRowLabour As Long
    Dim lastRowH As Long
    Dim lastRowPartsList As Long
    Dim i As Long
    Dim currentValue As String
    Dim lookupFormula As String
    
    ' Set reference to the "AFE" workbook
    Set wb = Workbooks("AFE")
    
    ' Set reference to the "Labour Data" sheet
    Set wsLabourData = wb.Sheets("Labour Data")
    
    ' Set reference to the "Data" sheet (will be renamed to "Parts List")
    Set wsPartsList = wb.Sheets("Data")
    
    ' Rename the "Data" sheet to "Parts List"
    wsPartsList.Name = "Parts List"
    
    ' Unhide the "Labour Data" sheet
    wsLabourData.Visible = xlSheetVisible
    
    
    ' Insert a new column A in "Labour Data" and perform XLOOKUP
    wsLabourData.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    ' Add the heading to cell A1 in "Labour Data"
    wsLabourData.Cells(1, "A").Value = "Task List Description"
    
    ' Format A1:J1 in "Labour Data"
    With wsLabourData.Range("A1:J1")
        .Interior.Color = RGB(0, 0, 0) ' Black background
        .Font.Color = RGB(255, 255, 255) ' White text
        .Font.Name = "Aptos Display" ' Font name
        .Font.Size = 12 ' Font size
        .HorizontalAlignment = xlCenter ' Center the text
        .VerticalAlignment = xlCenter ' Vertically align the text
    End With
    
    ' Loop through all rows in Column B of "Labour Data" and apply the XLOOKUP formula
    currentValue = "" ' Initialize current value for comparison
        ' Find the last row with data in column B of the "Labour Data" sheet
    lastRowLabour = wsLabourData.Cells(wsLabourData.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRowLabour ' Start from row 2 to skip header
        ' Only apply the lookup formula if the value in B changes from the previous row
        If wsLabourData.Cells(i, "B").Value <> currentValue Then
            lookupFormula = "=IFERROR(XLOOKUP(B" & i & ",'Components Due'!A:A,'Components Due'!B:B), """")"
            wsLabourData.Cells(i, "A").formula = lookupFormula
            currentValue = wsLabourData.Cells(i, "B").Value ' Update current value
        End If
    Next i
    
    ' Delete columns K, L, and M in "Labour Data"
    wsLabourData.Columns("K:M").Delete
    
    ' AutoFit all columns in "Labour Data" to fit the content
    wsLabourData.Columns.AutoFit
    
    ' Freeze the top row (Row 1) in "Labour Data"
    wsLabourData.Activate
    wsLabourData.application.ActiveWindow.FreezePanes = False ' Unfreeze if there was any previous freeze
    wsLabourData.Rows("2:2").Select ' Select row 2 to freeze above it
    wsLabourData.application.ActiveWindow.FreezePanes = True ' Freeze the top row
    
    ' Clear all contents in Column A (Parts List) without deleting the rows
    wsPartsList.Columns("A:A").ClearContents
    
    ' Find the last row with data in column H (for the lookup range)
    lastRowH = wsPartsList.Cells(wsPartsList.Rows.Count, "H").End(xlUp).Row
    
    ' Add the XLOOKUP formula to column A starting from row 1 (A1)
    wsPartsList.Range("A1:A" & lastRowH).formula = "=IFERROR(XLOOKUP(H1, 'Components Due'!A:A, 'Components Due'!B:B), """")"
    
    ' Optionally, you can paste the formulas as values if you don't want to leave them as formulas
    ' wsPartsList.Range("A1:A" & lastRowH).Value = wsPartsList.Range("A1:A" & lastRowH).Value
    
    ' AutoFit columns A-J in "Parts List" to fit the content
    wsPartsList.Columns("A:J").AutoFit
    
    ' Re-hide the "Labour Data" sheet after the operation (so it's still visible on the tab)
    wsLabourData.Visible = xlSheetHidden
    
    ' Re-hide the "Parts List" sheet after the operation (so it's still visible on the tab)
    wsPartsList.Visible = xlSheetHidden
End Sub
Sub TallyValuesBetweenYellowCells()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cell As Range
    Dim startCell As Range
    Dim sumRange As Range
    Dim lastRow As Long
    
    ' Set reference to the "AFE" workbook
    Set wb = Workbooks("AFE")
    
    ' Set reference to the "Parts List" sheet
    Set ws = wb.Sheets("Parts List")
    
    ' Define the last row in column I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Loop through each cell in column I
    For Each cell In ws.Range("I1:I" & lastRow)
        ' Check if the cell is yellow (RGB(255, 255, 153))
        If cell.Interior.Color = RGB(255, 255, 153) Then
            ' If we have a start cell, set the formula
            If Not startCell Is Nothing Then
                startCell.FormulaR1C1 = "=SUM(R[1]C:R[" & cell.Row - startCell.Row - 1 & "]C)"
            End If
            ' Set the new start cell
            Set startCell = cell
        End If
    Next cell
    
    ' Set the formula for the last range if there's a remaining start cell
    If Not startCell Is Nothing Then
        startCell.FormulaR1C1 = "=SUM(R[1]C:R[" & lastRow - startCell.Row & "]C)"
    End If
    
    ' Copy and paste values of column A and column J
    ws.Range("A1:A" & lastRow).Copy
    ws.Range("A1:A" & lastRow).PasteSpecial Paste:=xlPasteValues
    ws.Range("J1:J" & lastRow).Copy
    ws.Range("J1:J" & lastRow).PasteSpecial Paste:=xlPasteValues
    application.CutCopyMode = False
End Sub




Sub VendorAssignment()
    ' Declare variables
    Dim wsAFE As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim vendorDescription As String
    Dim vendorAssigned As String
    Dim newWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim outputRow As Long
    Dim group As String
    Dim groupCounter As String
    Dim msg As String

    ' Set reference to the "Labour Data" sheet in AFE workbook
    Set wsAFE = Workbooks("AFE").Sheets("Labour Data")

    ' Find the last row of data in column C (Vendor Description)
    lastRow = wsAFE.Cells(wsAFE.Rows.Count, "C").End(xlUp).Row

    ' Check if "VENDOR NOT ASSIGNED" is found and create new workbook if needed
    vendorAssigned = False
    For i = 2 To lastRow ' Start from row 2 to skip headers
        vendorDescription = wsAFE.Cells(i, 3).Value ' Vendor Description in Column C
        If vendorDescription = "VENDOR NOT ASSIGNED" Then
            vendorAssigned = True
            Exit For
        End If
    Next i

    If vendorAssigned Then
        ' Show a message box if "VENDOR NOT ASSIGNED" is found
        msg = "TASKLIST HAVE BEEN IDENTIFIED WITHOUT VENDORS ASSIGNED.                                                                                       A LIST HAS BEEN GENERATED INTO A SPREADSHEET, PLEASE ASSIGN VENDORS IN SAP - IA06.                                                   CHANGES CAN BE RELOADED AFTER 4AM - DEAD LINE MONITORING."
        MsgBox msg, vbExclamation, "Vendor Assignment Alert"
        
        ' Create a new workbook for the Vendor Assignment data
        Set newWorkbook = Workbooks.Add
        Set newSheet = newWorkbook.Sheets(1)
        newSheet.Name = "Vendor Assignment"

        ' Add headers to the new sheet
        newSheet.Cells(1, 1).Value = "Group"
        newSheet.Cells(1, 2).Value = "Group"
        newSheet.Cells(1, 3).Value = "Group Counter"
        newSheet.Cells(1, 4).Value = "Vendor Description"

        ' Set row for output in the new sheet
        outputRow = 2 ' Start from row 2 to add data

        ' Loop through "Labour Data" sheet and gather information where "VENDOR NOT ASSIGNED" is found
        For i = 2 To lastRow ' Start from row 2 to skip headers
            vendorDescription = wsAFE.Cells(i, 3).Value ' Vendor Description in Column C
            If vendorDescription = "VENDOR NOT ASSIGNED" Then
                ' Retrieve Group and Group Counter values from Columns B (assuming Group is in column B)
                group = wsAFE.Cells(i, 2).Value ' Group in Column B
                
                ' Write the Group and Vendor Description to the new sheet
                newSheet.Cells(outputRow, 1).Value = group
                newSheet.Cells(outputRow, 4).Value = vendorDescription

                outputRow = outputRow + 1
            End If
        Next i

        ' Step 1: Apply formula in Column B to extract everything except the last 2 characters from Group
        newSheet.Range("B2:B" & outputRow - 1).formula = "=LEFT(A2, LEN(A2)-2)"

        ' Step 2: Apply formula in Column C to extract the last 2 characters from Group
        newSheet.Range("C2:C" & outputRow - 1).formula = "=RIGHT(A2, 2)"

        ' Step 3: Copy the sheet and paste values to remove formulas
        newSheet.Cells.Copy
        newSheet.Cells.PasteSpecial Paste:=xlPasteValues

        ' Step 4: Delete Column A (Group column) as it is no longer needed
        newSheet.Columns("A").Delete

        ' Step 5: Auto-fit all columns to adjust the width based on the content
        newSheet.Cells.Columns.AutoFit

        ' Optionally, save the new workbook with the name "Vendor Assignment"
        ' newWorkbook.SaveAs "C:\Path\To\Save\Vendor Assignment.xlsx" ' Modify the path as needed
        ' newWorkbook.Close False
    End If
End Sub

Sub BreakLinks()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim links As Variant
    Dim i As Long
    
    ' Set reference to the "AFE" workbook
    Set wb = Workbooks("AFE")
    
    ' Get all links in the workbook
    links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
    
    ' Check if there are any links
    If Not IsEmpty(links) Then
        ' Loop through each link and break it
        For i = LBound(links) To UBound(links)
            If InStr(links(i), "AFE Builder from SAP HANA") > 0 Then
                wb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
                ' Debug.Print "Link to " & links(i) & " has been broken."
            End If
        Next i
    Else
        ' Debug.Print "No links found in the workbook.'
    End If
    
    ' Set reference to the "Components Due" sheet
    Set ws = wb.Sheets("Components Due")
    
    ' Highlight column E, copy and paste values
    'ws.Range("E:E").Copy
    'ws.Range("E:E").PasteSpecial Paste:=xlPasteValues
    
    ' Highlight column F, copy and paste values
    'ws.Range("F:F").Copy
    'ws.Range("F:F").PasteSpecial Paste:=xlPasteValues
    
    ' Insert a column between I and J
    ws.Columns("J:J").Insert Shift:=xlToRight
    
    ' Copy the value from K2 and paste it into the new column J2
    ws.Range("K2").Copy
    ws.Range("K2").PasteSpecial Paste:=xlPasteValues
    
    ' Delete the original column J
    ws.Columns("J:J").Delete
    
    ' Insert a column at C in "Components Due" sheet
    ws.Columns("C:C").Insert Shift:=xlToRight
    
    ' Create a heading in C2
    ws.Range("C2").Value = "TaskList and Description"
    
    ' Concatenate A3 and B3 with a space in between from C3 to the bottom of the table
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 3 To lastRow
        ws.Cells(i, 3).Value = ws.Cells(i, 1).Value & " " & ws.Cells(i, 2).Value
    Next i
    
    ' Hide the new column C
    ws.Columns("C:C").Hidden = True
    
    ' Set reference to the "Labour Data" sheet
    Set ws = wb.Sheets("Labour Data")
    
    ' Delete column A
    ws.Columns("A:A").Delete
    
    ' Find the last row with value in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Apply XLOOKUP in column G starting at G2
    For i = 2 To lastRow
        ws.Cells(i, 7).formula = "=IF(XLOOKUP(A" & i & ",'Components Due'!C:C,'Components Due'!I:I,,-1)=1,""Yes"",""No"")"
    Next i
    
    ' Activate the "Summary" tab
    wb.Sheets("Summary").Activate
    
    application.CutCopyMode = False
End Sub



























