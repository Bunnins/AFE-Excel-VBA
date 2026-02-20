Attribute VB_Name = "Module1"
Sub Createworkbook()
    Dim newWorkbook As Workbook
    Dim currentWorkbook As Workbook

    ' Set the current workbook (the workbook you're working in)
    Set currentWorkbook = ThisWorkbook
    
    ' Create a new workbook
    Set newWorkbook = Workbooks.Add
    
    ' Copy each template sheet from the current workbook to the new workbook
    currentWorkbook.Sheets("Template").Copy Before:=newWorkbook.Sheets(1)
    currentWorkbook.Sheets("Template2").Copy Before:=newWorkbook.Sheets(2)
    currentWorkbook.Sheets("Template3").Copy Before:=newWorkbook.Sheets(3)
    
    ' Rename the sheets in the new workbook
    newWorkbook.Sheets("Template").Name = "Summary"
    newWorkbook.Sheets("Template2").Name = "Components Due"
    newWorkbook.Sheets("Template3").Name = "Budget"
    
    ' Change the name of the new workbook without saving it
    newWorkbook.Windows(1).Caption = "AFE"  ' Set the window name to "AFE"
    
    ' Hide the default Sheet1 in the new workbook (if it exists)
    On Error Resume Next ' In case Sheet1 doesn't exist
    newWorkbook.Sheets("Sheet1").Visible = False
    On Error GoTo 0 ' Reset error handling
    
    ' The workbook is not saved, just opened with the desired name
End Sub
