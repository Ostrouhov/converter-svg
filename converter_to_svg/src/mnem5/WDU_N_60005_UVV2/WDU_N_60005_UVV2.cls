VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDU_N_60005_UVV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub AAA_AI_20_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_20
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource83.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG21.Description = frsReadValue(ps_DataSource83.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CFixPicture_Initialize()
On Error GoTo ErrorHandler
    frszinitPicture Me
    Exit Sub
ErrorHandler:
    frsHandleError

End Sub

Private Sub AAA_AI_11_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CFixPicture_InitializeConfigure()
On Error GoTo ErrorHandler
    frszInitConfig Me
Exit Sub
ErrorHandler:
     frsHandleError
End Sub

Private Sub Group364_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frsOpenUserPicture "WDU_N_60005_UUV2_KMH_deltav.grf", 1, "кмх", "КМХ для УУВ-2"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group364_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost29_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource83.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost29_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE21_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos12_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource83.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos12.Description = frsReadValue(ps_DataSource83.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CFixPicture_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    If Button = vbKeyRButton Then
      'Track #36546 - Removed IEPOP1 on : 12/3/01
      frszCreateAnEPMMenu Me
    End If
Exit Sub
ErrorHandler:
     frsHandleError
End Sub

'This event will update the currentlink
Private Sub CFixPicture_SelectionChanged()
On Error GoTo ErrorHandler
    frszpicCurLink Me.ContainedSelections, ps_CurrentLink
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG12_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG12.Description = frsReadValue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost28_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource76.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE12_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos11_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos11.Description = frsReadValue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost28_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr22_Click()
On Error GoTo ErrorHandler
    If MsgBox("Закончить измерения?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "2", "DVSYS.60005-UUW2-KMH/KMH_START.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr28_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №6?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "6", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClosePic_Click()
On Error GoTo ErrorHandler
    frszClosePictureButton Me
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr27_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №5?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "5", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CPF_4_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "wdu_N_60005_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr26_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №4?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "4", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CPF_4_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Rect111_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frsOpenUserPicture "WDU_N_60005_UUV2_KMH_deltav.grf", 6, "", ""
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub bmpClrErr25_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №3?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "3", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub NavigateButtons_Edit()
On Error GoTo ErrorHandler
    frszopenPrevNextPic (Me.Name)
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr24_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №2?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "2", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Group363_Click()
    On Error GoTo ErrHandler
        frsWriteValue "6", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group363_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group362_Click()
    On Error GoTo ErrHandler
        frsWriteValue "5", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group362_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group361_Click()
    On Error GoTo ErrHandler
        frsWriteValue "4", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group361_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group360_Click()
    On Error GoTo ErrHandler
        frsWriteValue "3", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group360_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group359_Click()
    On Error GoTo ErrHandler
        frsWriteValue "2", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group359_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group356_Click()
    On Error GoTo ErrHandler
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW2-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUW2_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group356_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpNextPicture_Click()
On Error GoTo ErrorHandler
    frszopenNextPicture Me
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr23_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №1?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005-UUW2-KMH/KMH_LINE.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpPreviousPicture_Click()
On Error GoTo ErrorHandler
    frszopenPreviousPicture Me
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CPF_44000_01_Main_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "CPF_00000_VPDD.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr8_Click()
On Error GoTo ErrorHandler
    If MsgBox("Начать измерения?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005-UUW2-KMH/KMH_START.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl61_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV050", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr21_Click()
On Error GoTo ErrorHandler
    If MsgBox("Вывести из учета?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "2", "DVSYS.60005_UUW2_SUM/MLTX1/SELECTOR.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CPF_44000_01_Main_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr20_Click()
On Error GoTo ErrorHandler
    If MsgBox("Ввести в учет?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005_UUW2_SUM/MLTX1/SELECTOR.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub But_VKip3_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "CPF_47000_AirKIP.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub But_VKip4_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "WDU_N_60023_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub wdu_n_fire_3_Edit()
'Name: WDU_N_FRS/wdu_n_fire_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fire_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG7_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_42_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_42
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl57_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV024", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode20_Click()
  On Error GoTo ErrHandler
    If frsReadValue("DVSYS.60005-FIRE-74002/unblock.F_CV") = 1 Then
      frsWriteValue "0", "DVSYS.60005-FIRE-74002/unblock.F_CV"
    Else
      frsWriteValue "1", "DVSYS.60005-FIRE-74002/unblock.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl56_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV023", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl49_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0771", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode21_Click()
  On Error GoTo ErrHandler
    frsWriteValue "1", "DVSYS.60005-FIRE-74002/Reset.F_CV"
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap63_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap63.Description = frsReadValue(ps_DataSource82.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG7.Description = frsReadValue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AAA_AI_4_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_4
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE7_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl48_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0781", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos6_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub TechCap51_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption49.Description = frsReadValue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos6.Description = frsReadValue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost53_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource77.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost53_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_20_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_20
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_21_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_21
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER24_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource109.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap58_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap58.Description = frsReadValue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap58_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap51_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_41_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_41
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap63_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER24.Description = frsReadValue(ps_DataSource109.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub CLICKER25_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource110.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER25.Description = frsReadValue(ps_DataSource110.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub But_VKip6_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "WDU_N_60026.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1595_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60005_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap59_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap59.Description = frsReadValue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption49_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption49.Description = frsReadValue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1595_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption61_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption61.Description = frsReadValue(ps_DataSource82.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption49_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap50_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption48.Description = frsReadValue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption61_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource82.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_29_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_29
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap50_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER26_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource112.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption48_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption48.Description = frsReadValue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption48_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER26.Description = frsReadValue(ps_DataSource112.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub ZD_CS_U_18_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_18
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER27_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource113.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER27.Description = frsReadValue(ps_DataSource113.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub ZD_CS_U_19_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_19
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER28_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource115.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER28.Description = frsReadValue(ps_DataSource115.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER29_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource116.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER29.Description = frsReadValue(ps_DataSource116.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub ZD_CS_L_13_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_22_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_22
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl44_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0741", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_14_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_14
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_40_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_40
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap45_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption44.Description = frsReadValue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap45_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption44_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption44.Description = frsReadValue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect102_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60006_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect102_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption44_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DI_23_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_23
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_12_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_30_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_30
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_24_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_24
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_13_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AAA_AI_2_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_25_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_25
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_14_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_14
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_20_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_20
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_31_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_31
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_26_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_26
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_L_18_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_18
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_21_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_21
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap41_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption41.Description = frsReadValue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap41_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption41_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption41.Description = frsReadValue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption41_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_32_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_32
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_15_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_15
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Polygon1604_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60007_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl41_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0761", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1604_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Polygon1605_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60011_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1605_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_27_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_27
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER30_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource118.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER30.Description = frsReadValue(ps_DataSource118.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub DI_16_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_16
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_33_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_33
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_39_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_39
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_22_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_22
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource119.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG24_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource105.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG24.Description = frsReadValue(ps_DataSource105.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub CLICKER31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER31.Description = frsReadValue(ps_DataSource119.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub DI_17_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_17
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource104.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG23.Description = frsReadValue(ps_DataSource104.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub ZD_CS_U_23_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_23
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_34_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_34
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource103.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG22.Description = frsReadValue(ps_DataSource103.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_24_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_24
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub DI_18_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_18
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_35_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_35
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG20_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource84.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG20.Description = frsReadValue(ps_DataSource84.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_23_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_23
Exit Sub
ErrHandler:
    frsHandleError
End Sub




Private Sub ZD_CS_U_24_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_24
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_19_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_19
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_36_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_36
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG18_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_25_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_25
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_22_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_22
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG18.Description = frsReadValue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_37_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_37
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_26_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_26
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AAA_AI_21_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_21
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG17_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource70.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub ZD_CS_U_38_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_38
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption56_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption56.Description = frsReadValue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG17.Description = frsReadValue(ps_DataSource70.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub TechCap59_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_18_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_18
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption56_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource78.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub AI_TAG16_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource69.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub ZD_CS_U_27_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_27
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG16.Description = frsReadValue(ps_DataSource69.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_16_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_16
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_17_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_17
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Polygon1594_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_00_62131_02_ESh4.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Polygon1594_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost62_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource105.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption57_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption57.Description = frsReadValue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption57_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource79.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_28_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_28
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_16_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_16
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_17_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_17
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV025", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsReadValue(ps_DataSource46.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost61_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource104.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource46.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsReadValue(ps_DataSource46.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource46.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV026", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption19.Description = frsReadValue(ps_DataSource93.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap19_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource93.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption19.Description = frsReadValue(ps_DataSource93.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost60_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource103.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption19_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource93.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV030", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption20.Description = frsReadValue(ps_DataSource35.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap20_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption20.Description = frsReadValue(ps_DataSource35.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost62_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost59_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource84.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption20_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl21_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV036", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption21.Description = frsReadValue(ps_DataSource47.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap21_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource47.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost57_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource72.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption21.Description = frsReadValue(ps_DataSource47.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption21_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource47.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl33_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV037", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost61_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption33.Description = frsReadValue(ps_DataSource59.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap33_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource59.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption33.Description = frsReadValue(ps_DataSource59.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption33_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource59.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost56_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource70.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource97.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource105.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos23.Description = frsReadValue(ps_DataSource105.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub SubCtrl24_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV032", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption24.Description = frsReadValue(ps_DataSource50.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource50.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost60_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption24.Description = frsReadValue(ps_DataSource50.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE24_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption24_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource50.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource6.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost51_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource69.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub SubCtrl26_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV033", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsReadValue(ps_DataSource52.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap26_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource52.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource104.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos22.Description = frsReadValue(ps_DataSource104.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub rAI_Dost59_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG1.Description = frsReadValue(ps_DataSource6.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsReadValue(ps_DataSource52.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption26_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource52.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl34_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV034", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE23_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub TechCap34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption34.Description = frsReadValue(ps_DataSource60.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap34_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource60.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER16.Description = frsReadValue(ps_DataSource97.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub rAI_Dost57_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER17_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource98.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl22_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV031", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource103.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos21.Description = frsReadValue(ps_DataSource103.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub TechCap22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption22.Description = frsReadValue(ps_DataSource48.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption34.Description = frsReadValue(ps_DataSource60.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption34_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource60.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV035", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap37_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption37.Description = frsReadValue(ps_DataSource62.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE22_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub TechCap37_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource62.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption37_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption37.Description = frsReadValue(ps_DataSource62.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap22_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource48.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos20_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource84.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos20.Description = frsReadValue(ps_DataSource84.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub Caption22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption22.Description = frsReadValue(ps_DataSource48.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption37_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource62.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption22_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource48.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER17.Description = frsReadValue(ps_DataSource98.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub rAI_Dost56_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource99.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE20_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER18.Description = frsReadValue(ps_DataSource99.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub CLICKER19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource100.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER19.Description = frsReadValue(ps_DataSource100.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub SubCtrl28_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV039", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsReadValue(ps_DataSource54.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap28_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource54.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsReadValue(ps_DataSource54.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption28_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource54.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl29_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV041", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption29.Description = frsReadValue(ps_DataSource55.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap29_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource55.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos18_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos18.Description = frsReadValue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub Caption29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption29.Description = frsReadValue(ps_DataSource55.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption29_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource55.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl30_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV043", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost51_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption30.Description = frsReadValue(ps_DataSource56.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE18_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap30_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource56.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption30.Description = frsReadValue(ps_DataSource56.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl23_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV038", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption23.Description = frsReadValue(ps_DataSource49.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap23_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource49.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption23.Description = frsReadValue(ps_DataSource49.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos17_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource70.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos17.Description = frsReadValue(ps_DataSource70.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub Caption30_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource56.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption23_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource49.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl25_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV040", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsReadValue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl36_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV045", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE17_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap36_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption36.Description = frsReadValue(ps_DataSource61.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap25_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource51.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsReadValue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource61.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption25_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource51.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption36_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption36.Description = frsReadValue(ps_DataSource61.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos16_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource69.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos16.Description = frsReadValue(ps_DataSource69.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub AI_VALUE1_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource6.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos.Description = frsReadValue(ps_DataSource6.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl27_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV042", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsReadValue(ps_DataSource53.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE16_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub TechCap27_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource53.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource61.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsReadValue(ps_DataSource53.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption27_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource53.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl39_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV047", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap39_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption39.Description = frsReadValue(ps_DataSource64.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost32_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource6.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost32_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap39_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource64.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption39_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption39.Description = frsReadValue(ps_DataSource64.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption39_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource64.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV044", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption35.Description = frsReadValue(ps_DataSource94.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl40_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0711", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap35_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource94.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption35.Description = frsReadValue(ps_DataSource94.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption35_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource94.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub TechCap40_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption40.Description = frsReadValue(ps_DataSource65.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl38_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV046", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap40_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource65.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption40_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption40.Description = frsReadValue(ps_DataSource65.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption40_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource65.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap38_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption38.Description = frsReadValue(ps_DataSource63.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap38_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource63.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption38_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption38.Description = frsReadValue(ps_DataSource63.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption38_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource63.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub CLICKER20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource101.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER20.Description = frsReadValue(ps_DataSource101.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub CLICKER21_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource102.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER21.Description = frsReadValue(ps_DataSource102.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub CLICKER22_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource107.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER22.Description = frsReadValue(ps_DataSource107.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub CLICKER23_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource108.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub CLICKER23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER23.Description = frsReadValue(ps_DataSource108.CurrentValue + "/DESC.A_CV")
End Sub
Private Sub SubCtrl31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV048", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption31.Description = frsReadValue(ps_DataSource57.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap31_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource57.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption31.Description = frsReadValue(ps_DataSource57.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption31_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource57.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl32_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV049", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption32.Description = frsReadValue(ps_DataSource58.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap32_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource58.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption32.Description = frsReadValue(ps_DataSource58.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption32_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource58.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Rect111_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
