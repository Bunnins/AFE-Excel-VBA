VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "Progress Indicator"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5490
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Abort_Click()
Dim result As String
result = MsgBox("Abort AFE Creation?", vbYesNo)
If result = vbYes Then End
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
If MsgBox("Are you sure you want to exit?", vbYesNo) = vbYes Then
Cancel = True
End
Else

End If
End If
End Sub

Private Sub UserForm_Initialize()
    Dim excelCenterX As Long, excelCenterY As Long
    Dim formWidth As Long, formHeight As Long
    
    ' Get the dimensions of the Excel application window
    excelCenterX = application.Left + (application.Width / 2)
    excelCenterY = application.Top + (application.Height / 2)
    
    ' Get the dimensions of the UserForm
    formWidth = Me.Width
    formHeight = Me.Height
    
    ' Position the UserForm in the center of the Excel window
    Me.Left = excelCenterX - (formWidth / 2)
    Me.Top = excelCenterY - (formHeight / 2)
End Sub
