VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreateNewWorkbookWithTemplates()
    Dim newWorkbook As Workbook
    Dim currentWorkbook As Workbook
    
    ' Set current workbook (the workbook you're working in)
    Set currentWorkbook = ThisWorkbook
    
    ' Create a new workbook (it will initially have 1 sheet)
    Set newWorkbook = Workbooks.Add
    
    ' Copy Template sheet from the current workbook to the new workbook
    currentWorkbook.Sheets("Template").Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
    
    ' Copy Template2 sheet from the current workbook to the new workbook
    currentWorkbook.Sheets("Template2").Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
    
    ' Copy Template3 sheet from the current workbook to the new workbook
    currentWorkbook.Sheets("Template3").Copy After:=newWorkbook.Sheets(newWorkbook.Sheets.Count)
    
    ' Rename the copied sheets in the new workbook
    newWorkbook.Sheets(newWorkbook.Sheets.Count - 2).Name = "Summary"  ' Template sheet
    newWorkbook.Sheets(newWorkbook.Sheets.Count - 1).Name = "Components Due"  ' Template2 sheet
    newWorkbook.Sheets(newWorkbook.Sheets.Count).Name = "Budget"  ' Template3 sheet
    
    ' Change the name of the new workbook without saving it
    newWorkbook.Windows(1).Caption = "AFE"  ' Set the name (without saving)
    
    ' The workbook is not saved, just opened with the desired name.
End Sub


