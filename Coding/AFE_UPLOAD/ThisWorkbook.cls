VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub workbook_open()
    '
    ' Macro1 Macro
    '
    
    ' Refresh all data connections in the active workbook
    application.CommandBars("Queries and Connections").Visible = True
    ActiveWorkbook.RefreshAll
    

End Sub


