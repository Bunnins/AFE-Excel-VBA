Attribute VB_Name = "Components"

Sub ExtractMatchingItems()
    ' Declare workbook and worksheet variables
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorkbook As Workbook
    Dim destinationWorksheet As Worksheet
    Dim resultWorksheet As Worksheet ' For storing the results of the Group and Group Counter search
    Dim lastRowSource As Long, resultRow As Long
    Dim i As Long, j As Long
    Dim searchValue As String
    Dim searchValueFound As Boolean
    Dim RestartRow As Long
    Dim lastRowTaskList As Long, resultRow2 As Long
    Dim taskListRepSheet As Worksheet
    Dim groupCounter As String, columnI As String, columnJ As String, columnP As String
    Dim dict As Object ' Dictionary to track unique values
application.ScreenUpdating = False

   With ufProgress
        .LabelCaption.Caption = "Retrieving Tasklist Data - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents



    On Error GoTo ErrorHandler ' Global error handling

    ' Set references to the workbooks and worksheets
    Set sourceWorkbook = Workbooks("AFE.xlsx")
    Set sourceWorksheet = sourceWorkbook.Sheets("Components Due")

    ' Set the task list sheet for reference
    Set taskListRepSheet = Workbooks("AFE Builder from SAP HANA").Sheets("TASK_LIST_REP")


    ' Define the sorting range
    With taskListRepSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=taskListRepSheet.Range("H:H"), _
                        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
        ' Apply the sort to the entire used range
        .SetRange taskListRepSheet.UsedRange
        .Header = xlYes ' Indicates the data range has headers
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ' Create a new dictionary to track unique values
    Set dict = CreateObject("Scripting.Dictionary")

    ' Step 3: Determine the last row in the source sheet
    lastRowSource = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "H").End(xlUp).Row
    resultRow = 1 ' Start writing results in the first row of the destination sheet

    ' Step 4: Loop through rows in the source sheet starting from row 3

    
    
    For i = 3 To lastRowSource
        ' Check if the value in column H is 1
        'If sourceWorksheet.Cells(i, "H").Value = 1 Then
            ' Step 7: Get the value from column A to use as searchValue
            searchValue = sourceWorksheet.Cells(i, "A").Value

            ' Ensure searchValue is a string (to avoid issues with numeric values)
            searchValue = CStr(searchValue)

            ' Check if the search value is already in the dictionary (i.e., has been added before)
            If Not dict.Exists(searchValue) Then
                ' Add the value to the Data sheet (Column A)
                Worksheets("Data").Cells(resultRow, 1).Value = searchValue
                dict.Add searchValue, 1 ' Mark the value as added (you could store any value, here I use 1)

                resultRow = resultRow + 1
            End If
       'End If
    Next i

    ' Sort the group and group counters
    Worksheets("Data").Columns("A:A").AutoFit
    ' Sort Column A in the worksheet named "Data"
    Worksheets("Data").Sort.SortFields.Clear
    Worksheets("Data").Sort.SortFields.Add Key:=Worksheets("Data").Range("A:A"), _
        SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal

    ' Apply the sort
    With Worksheets("Data").Sort
        .SetRange Worksheets("Data").UsedRange
        .Header = xlNo ' Change to xlYes if the column has a header
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With


On Error Resume Next
    ' Step 7: Now, for each value in column A of the destination sheet, search in TASK_LIST_REP tab
    RestartRow = 1
    resultRow2 = 1 ' Start writing Group and Group Counter results in the new sheet
    lastRowTaskList = taskListRepSheet.Cells(taskListRepSheet.Rows.Count, "H").End(xlUp).Row

    ' Loop through each row in column A of the destination sheet
    For i = 1 To resultRow - 1
        searchValue = Worksheets("Data").Cells(i, 1).Value
        
        ' Ensure searchValue is a string (to avoid issues with numeric values)
        searchValue = CStr(searchValue)
searchValueFound = False
        ' Search for this value in column H of the TASK_LIST_REP sheet
 For j = 3 To lastRowTaskList ' Starting from row 3 in TASK_LIST_REP
    If taskListRepSheet.Cells(j, "H").Value = searchValue Then
    searchValueFound = True
        ' Concatenate the current Group and Group Counter (H & P columns)
        Dim duplicate2 As String
        duplicate2 = taskListRepSheet.Cells(j, "H").Value & taskListRepSheet.Cells(j, "P").Value

        ' Only add if the combination is unique
        If Not dict.Exists(duplicate2) Then
            ' Initialize start and end rows for the dynamic range
            Dim startRow As Long, endRow As Long
            startRow = j
            endRow = j

            ' Extend the range while the value in Column H matches the next row
            Do While taskListRepSheet.Cells(endRow + 1, "H").Value = taskListRepSheet.Cells(endRow, "H").Value _
                And Not IsEmpty(taskListRepSheet.Cells(endRow + 1, "H").Value)
                endRow = endRow + 1
            Loop
            
            ' Copy the dynamic range from Columns H to P
            taskListRepSheet.Range(taskListRepSheet.Cells(startRow, "H"), taskListRepSheet.Cells(endRow, "P")).Copy

            ' Paste into the "Data" worksheet starting at the appropriate row and column
            Worksheets("Data").Cells(resultRow2, 5).PasteSpecial Paste:=xlPasteValues

            ' Update the dictionary to track this combination
            dict.Add duplicate2, 1

            ' Update resultRow2 to account for the pasted rows
            resultRow2 = resultRow2 + (endRow - startRow + 1)

            ' Skip the rows that were already processed
            j = endRow
            RestartRow = endRow
        End If
ElseIf j = lastRowTaskList Then
    If searchValueFound = False Then
        'MsgBox "The search value '" & searchValue & "' was not found in Column H.", vbExclamation, "Search Value Not Found"
        j = RestartRow
        GoTo Newstart
    End If
End If
Next j
Newstart:

       With ufProgress
        .LabelCaption.Caption = "Retrieving Tasklist Data - " & Format((i / resultRow) * 100, "0") & "% Complete"
        Dim roundedProgress As Double
        roundedProgress = Round((i / resultRow), 2) ' Limit to 2 decimal places
        .LabelProgress.Width = roundedProgress * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
    Next i


    ' Cut Column M
    Worksheets("Data").Columns("M").Cut
    
    ' Insert the cut column into Column F
    Worksheets("Data").Columns("F").Insert Shift:=xlToRight

    ' Delete the now-empty Column M
    Worksheets("Data").Columns("I:N").Delete


    ' Step 8: Auto fit columns in the result sheet
   'Worksheets("Data").Columns("A:H").AutoFit

       With ufProgress
        .LabelCaption.Caption = "Retrieving Tasklist Data - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    Exit Sub

ErrorHandler:
    ' Error handling message
    MsgBox "An error occurred: " & Err.Description & vbCrLf & "Error Number: " & Err.Number, vbCritical
End Sub

Sub ListOccurrencesInLabourCosting()
    ' Declare variables
    Dim wsData As Worksheet
    Dim wsLabourCosting As Worksheet
    Dim wsLCData As Worksheet
    Dim tbl As ListObject
    Dim lastRowData As Long
    Dim lastRowLCData As Long
    Dim searchValue As Variant
    Dim fullSearchValue As String ' To hold concatenated value of ID and Main Item Desc
    Dim i As Long, j As Long
    Dim groupCounter As Variant
    Dim workCenter As Variant
    Dim work As Variant
    Dim activityType As Variant
    Dim netPrice As Variant
    Dim CRHDWorkCenterResourceDesc As Variant
    Dim wbAFEBuilder As Workbook
    Dim wbAFE As Workbook
    Dim colIndex As Integer
    Dim wsAFE As Worksheet
    Dim wsComponentsDue As Worksheet ' Added variable for Components Due sheet

    ' Disable screen updating to improve performance
    application.ScreenUpdating = False
    
    ' Ensure the AFE builder workbook is open
    On Error Resume Next
    Set wbAFEBuilder = Workbooks("AFE builder from SAP HANA")
    On Error GoTo 0
    If wbAFEBuilder Is Nothing Then
        MsgBox "The 'AFE builder' workbook is not open."
        Exit Sub
    End If
    
    ' Ensure the AFE workbook is open
    On Error Resume Next
    Set wbAFE = Workbooks("AFE") ' Make sure the AFE workbook is open
    On Error GoTo 0
    If wbAFE Is Nothing Then
        MsgBox "The 'AFE' workbook is not open."
        Exit Sub
    End If
    
    ' Set references to the worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsLabourCosting = wbAFEBuilder.Sheets("Labour Costing")
    Set wsLCData = ThisWorkbook.Sheets("LC Data")
    Set wsComponentsDue = wbAFE.Sheets("Components Due") ' Reference to Components Due sheet
    
    ' Set reference to the table in Labour Costing sheet
    On Error Resume Next
    Set tbl = wsLabourCosting.ListObjects("TASK_LIST_REP__2")
    On Error GoTo 0
    If tbl Is Nothing Then
        MsgBox "Table 'TASK_LIST_REP__2' not found in the 'Labour Costing' worksheet."
        Exit Sub
    End If

    ' Find the last row in column A of Components Due sheet (starting from row 3)
    lastRowData = wsComponentsDue.Cells(wsComponentsDue.Rows.Count, "A").End(xlUp).Row

    ' Initialize the first row for LC Data sheet
    lastRowLCData = 2 ' Start from row 2 (row 1 will have the headers)

    ' Add headers to LC Data sheet
    wsLCData.Cells(1, 1).Value = "Group and Group Counter" ' Column A
    wsLCData.Cells(1, 2).Value = "Vendor Description" ' Column B
    wsLCData.Cells(1, 3).Value = "Work Hours" ' Column C
    wsLCData.Cells(1, 4).Value = "Activity Type" ' Column D
    wsLCData.Cells(1, 5).Value = "PMEX Cost" ' Column E
    wsLCData.Cells(1, 6).Value = "Vendor Name" ' Column F

    ' Loop through each value in column A of Components Due sheet starting from A3
    For i = 3 To lastRowData
        ' Concatenate the values from Col A ("ID") and Col B ("Main Item Desc") with a space in between
        fullSearchValue = wsComponentsDue.Cells(i, 1).Value & " " & wsComponentsDue.Cells(i, 2).Value
        searchValue = Left(wsComponentsDue.Cells(i, 1).Value, 10) ' Use the first 10 characters for the lookup
        
        ' Check if the ID in Components Due sheet (Col A) is not empty
        If Not IsEmpty(searchValue) Then
            ' Loop through all rows in the table to find all occurrences of searchValue in "Group and Group Counter" column (D)
            For j = 1 To tbl.DataBodyRange.Rows.Count
                ' If we find a match in column D (Group and Group Counter)
                If tbl.DataBodyRange.Cells(j, tbl.ListColumns("Group and Group Counter").Index).Value = searchValue Then
                    ' Get the corresponding values from other columns in the table
                    groupCounter = tbl.DataBodyRange.Cells(j, tbl.ListColumns("Group and Group Counter").Index).Value
                    workCenter = tbl.DataBodyRange.Cells(j, tbl.ListColumns("PLPO-LIFNR Vendor Desc").Index).Value ' Updated line
                    work = tbl.DataBodyRange.Cells(j, tbl.ListColumns("PLPO-ARBEI Work (Attr)").Index).Value
                    activityType = tbl.DataBodyRange.Cells(j, tbl.ListColumns("PLPO-LARNT Activity Type").Index).Value
                    netPrice = tbl.DataBodyRange.Cells(j, tbl.ListColumns("PLPO-PEINH Net Price (Attr)").Index).Value
                    
                    ' Find the index of the 'CRHD-ARBPL Work Center/Resource Desc' column
                    On Error Resume Next
                    colIndex = tbl.ListColumns("CRHD-ARBPL Work Center/Resource Desc").Index
                    On Error GoTo 0

                    If colIndex > 0 Then
                        ' Access the value
                        CRHDWorkCenterResourceDesc = tbl.DataBodyRange.Cells(j, colIndex).Value
                    Else
                        MsgBox "'CRHD-ARBPL Work Center/Resource Desc' column not found."
                        CRHDWorkCenterResourceDesc = "Not Found"
                    End If

                    ' Write the concatenated data into LC Data sheet (only in Column A)
                    wsLCData.Cells(lastRowLCData, 1).Value = fullSearchValue ' Concatenated ID and Main Item Desc (from Col A and Col B)
                    wsLCData.Cells(lastRowLCData, 2).Value = workCenter ' PLPO-LIFNR Vendor Desc (Vendor Name)
                    wsLCData.Cells(lastRowLCData, 3).Value = work ' PLPO-ARBEI Work (Attr)
                    wsLCData.Cells(lastRowLCData, 4).Value = activityType ' PLPO-LARNT Activity Type
                    wsLCData.Cells(lastRowLCData, 5).Value = netPrice ' PLPO-PEINH Net Price (Attr)
                    wsLCData.Cells(lastRowLCData, 6).Value = CRHDWorkCenterResourceDesc ' Vendor Name (CRHD-ARBPL Work Center/Resource Desc)

                    ' Increment the row in LC Data sheet for the next entry
                    lastRowLCData = lastRowLCData + 1
                End If
            Next j
        End If
    Next i

    ' Finalize the progress form to 100% when done
    With ufProgress
        .LabelCaption.Caption = "Retrieving Labour Data - 100% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    ' Now copy the entire LC Data sheet content and paste it into AFE workbook Sheet1 as values
    Set wsAFE = wbAFE.Sheets("Sheet1") ' Ensure that Sheet1 exists in AFE workbook
    
    ' Rename the Sheet1 to "Labour Data"
    wsAFE.Name = "Labour Data"
    
    ' Copy the contents of the LC Data sheet
    wsLCData.UsedRange.Copy

    ' Paste as values into the AFE workbook's "Labour Data" sheet
    wsAFE.Cells(1, 1).PasteSpecial Paste:=xlPasteValues

    ' Clear the clipboard
    application.CutCopyMode = False

    ' Re-enable screen updating
    application.ScreenUpdating = True
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
    Dim currentRow As Long
    Dim WorkHourstotal As Double
    Dim workOrder As String
    Dim nextWorkOrder As String
    Dim startRow As Long
    Dim order As Integer
    Dim workhours As Integer
    Dim OrderDescription As Integer
    Dim activity As Integer
    Dim Floc As Integer
    Dim unit As Integer
    Dim lastCol As Long
    Dim x As Long
    Dim startdate As Integer
    Dim columncheck As Boolean
application.ScreenUpdating = False
Set ws = ThisWorkbook.Sheets("Data")

   With ufProgress
        .LabelCaption.Caption = "Retrieving Material Pricing - 0" & "% Complete"
        .LabelProgress.Width = 0 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents

    WorkHourstotal = 0
    startRow = 1
    currentRow = 1
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    ' Loop through each row in column J (Work Order Number)
    Do While currentRow <= lastRow
        workOrder = ws.Cells(currentRow, "E").Value
        nextWorkOrder = ws.Cells((currentRow + 1), "E").Value

        ' Check if the work order changes
        
        If nextWorkOrder <> workOrder Then
        WorkHourstotal = WorkHourstotal + ws.Cells(currentRow, "I").Value
            ' Add total for totalwork on each Work Order Operation and insert a row
            ws.Rows(startRow).Insert Shift:=xlDown
            ws.Cells(startRow, "I").Value = WorkHourstotal
            ws.Cells(startRow, "H").Value = ws.Cells((startRow + 1), "E")
            ws.Rows(startRow).Interior.Color = RGB(255, 255, 153)
            ' Reset variables for the next work order
            WorkHourstotal = 0
            startRow = currentRow + 2
            currentRow = currentRow + 1
             
        Else
            ' Keep adding Work Order Operation hours
            WorkHourstotal = WorkHourstotal + ws.Cells(currentRow, "I").Value
        End If
        currentRow = currentRow + 1
        lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        
        With ufProgress
        .LabelCaption.Caption = "Retrieving Material Pricing  - " & Format((currentRow / lastRow) * 100, "0") & "% Complete"
        Dim roundedProgress As Double
        roundedProgress = Round((currentRow / lastRow), 2) ' Limit to 2 decimal places
        .LabelProgress.Width = roundedProgress * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
    End With
    DoEvents
        
    Loop
    
       With ufProgress
        .LabelCaption.Caption = "Retrieving Material Pricing - 100" & "% Complete"
        .LabelProgress.Width = 1 * (.FrameProgress.Width)
        .Repaint ' Force the form to refresh
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



























