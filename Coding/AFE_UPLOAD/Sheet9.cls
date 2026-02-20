VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub RunIW38Transaction()
    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim objSheet As Object
    Dim COL1 As String
    Dim url As String

    ' Start SAP GUI Scripting
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        MsgBox "PLEASE OPEN AN ACTIVE SAP SESSION"
        Exit Sub
    End If

    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0)
    Set session = connection.Children(0)
    If session Is Nothing Then
        MsgBox "Error getting SAP session"
        Exit Sub
    End If
    On Error GoTo 0

    ' Get active Excel sheet
    Set objSheet = ThisWorkbook.ActiveSheet

    ' Read the value from cell F11
    COL1 = Trim(CStr(objSheet.Cells(11, 6).Value))

    ' Check if the value is valid (not empty)
    If COL1 = "" Then
        MsgBox "Error: Empty value in cell F11"
        Exit Sub
    End If

    ' Inform the user of the process
    MsgBox "Processing SAP transaction for: " & COL1

    ' Run SAP IW38 transaction and input values
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/niw38"
    session.FindById("wnd[0]").SendVKey 0
    
    'Check User Authorisation
        If session.FindById("wnd[0]/sbar").MessageType = "E" Then
        MsgBox (session.FindById("wnd[0]/sbar").Text) & ". Please Conatct Your SAP Administrator For Assistance"
        'Return To Home Screen
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.FindById("wnd[0]").SendVKey 0
        Exit Sub
        End If

    session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtAUFNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtAUART-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtAUART-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTRNO-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTRNO-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtEQUNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtEQUNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSERMAT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSERMAT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtSERIALNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtSERIALNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtDEVICEID-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtDEVICEID-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtQMNUM-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtQMNUM-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGEWRK-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGEWRK-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtVAWRK-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtVAWRK-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtDATUV").Text = ""
    session.FindById("wnd[0]/usr/ctxtDATUB").Text = ""
    session.FindById("wnd[0]/usr/ctxtDY_PARNR").Text = ""
    session.FindById("wnd[0]/usr/ctxtWAERS").Text = ""
    session.FindById("wnd[0]/usr/ctxtLEAD_AUF-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtLEAD_AUF-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtMAUFNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtMAUFNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtIWERK-LOW").Text = "2711"
    session.FindById("wnd[0]/usr/ctxtIWERK-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPRIOK-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPRIOK-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtERNAM-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtERNAM-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtERDAT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtERDAT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTAI1-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTAI1-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTAE1-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTAE1-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtKTEXT-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtKTEXT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtAENAM-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtAENAM-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtAEDAT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtAEDAT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLBD-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLBD-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRP-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRP-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGLTRP-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGLTRP-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtWARPL-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtWARPL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtWAPOS-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtWAPOS-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtREVNR-LOW").Text = COL1
    session.FindById("wnd[0]/usr/ctxtREVNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPSPEL-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPSPEL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtKOSTV-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtKOSTV-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGLTRS-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGLTRS-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtFTRMI-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtFTRMI-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGETRI-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGETRI-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRI-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRI-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRS-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSTRS-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLVD-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLVD-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLNNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLNNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtPLNAL-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtPLNAL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLKNZ-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLKNZ-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtBAUTL-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtBAUTL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSWERK-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSWERK-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtILART-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtILART-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtARBPL-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtARBPL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtKOSTL-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtKOSTL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPROID-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPROID-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtAUFNT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtAUFNT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtVORUE-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtVORUE-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtKDAUF-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtKDAUF-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtKDPOS-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtKDPOS-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtVKORG-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtVKORG-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtVTWEG-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtVTWEG-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSPART-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSPART-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSBER-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtGSBER-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtBUKRS-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtBUKRS-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtANLNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtBEBER-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtBEBER-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTORT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSTORT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtEQFNR-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtEQFNR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtABCKZ-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtABCKZ-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtINGPR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtINGPR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtMSGRP-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtMSGRP-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtAUFPL-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtAUFPL-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLGRP-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPLGRP-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtKUNUM-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtKUNUM-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtGESIST-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtGESIST-HIGH").Text = ""
    session.FindById("wnd[0]/usr/txtGESPLN-LOW").Text = ""
    session.FindById("wnd[0]/usr/txtGESPLN-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtSOGEN-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtSOGEN-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPRCTR-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPRCTR-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtPAGESTAT-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtPAGESTAT-HIGH").Text = ""
    session.FindById("wnd[0]/usr/ctxtMONITOR").Text = ""
    session.FindById("wnd[0]/usr/ctxtREVNR-LOW").SetFocus
    session.FindById("wnd[0]/usr/ctxtREVNR-LOW").caretPosition = 8
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press

    ' Input the value from the Excel sheet into SAP
    session.FindById("wnd[0]/tbar[1]/btn[32]").Press
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").currentCellRow = -1
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").SelectAll
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").doubleClickCurrentCell
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = -1
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectColumn "SELTEXT"
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressColumnHeader "SELTEXT"
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 92
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 81
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "92"
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 65
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 63
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "65"
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").currentCellRow = 99
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").firstVisibleRow = 90
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectedRows = "99"
    session.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").doubleClickCurrentCell
    session.FindById("wnd[0]").SendVKey 0
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    session.FindById("wnd[0]/tbar[1]/btn[16]").Press
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press

    ' Final message
    MsgBox "Process Completed"

    ' Open the hyperlink stored in cell E16
    If objSheet.Hyperlinks.Count > 0 Then
        url = objSheet.Hyperlinks(1).Address
        If url <> "" Then
            ThisWorkbook.FollowHyperlink Address:=url
        Else
            MsgBox "No valid URL found in the hyperlink."
        End If
    Else
        MsgBox "No hyperlink found in cell E16."
    End If

    ' Release SAP GUI Scripting objects
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
End Sub


Private Sub TextBox1_Change()

End Sub
