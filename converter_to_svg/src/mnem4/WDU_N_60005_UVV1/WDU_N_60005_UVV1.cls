VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDU_N_60005_UVV1"
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

Private Sub bmpClrErr8_Click()
On Error GoTo ErrorHandler
    If MsgBox("Начать измерения?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005_UUW1_KMH/START_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr22_Click()
On Error GoTo ErrorHandler
    If MsgBox("Выполнить сброс всех параметров?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005_UUW1_KMH/RESET_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr23_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №1?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005_UUW1_KMH/NUM_IL_WORK_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr24_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №2?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "2", "DVSYS.60005_UUW1_KMH/NUM_IL_WORK_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub bmpClrErr25_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №3?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "3", "DVSYS.60005_UUW1_KMH/NUM_IL_WORK_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr26_Click()
On Error GoTo ErrorHandler
    If MsgBox("Задать линию №4?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "4", "DVSYS.60005_UUW1_KMH/NUM_IL_WORK_KMH.F_CV"
    End If
Exit Sub
ErrorHandler:
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


Private Sub SubCtrl60_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV012", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl59_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV003", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub SubCtrl58_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV004", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClrErr21_Click()
On Error GoTo ErrorHandler
    If MsgBox("Вывести из учета?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "2", "DVSYS.60005_UUW1_SUM/MLTX1/SELECTOR.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub bmpClrErr20_Click()
On Error GoTo ErrorHandler
    If MsgBox("Ввести в учет?", 1, "Подтверждение") = vbOK Then
    frsWriteValue "1", "DVSYS.60005_UUW1_SUM/MLTX1/SELECTOR.F_CV"
    End If
Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl55_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0721", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Group288_Click()
    On Error GoTo ErrHandler
        frsWriteValue "4", "DVSYS.60005-UUW1-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW1-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUV1_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group288_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap62_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap62.Description = frsReadValue(ps_DataSource81.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap61_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap61.Description = frsReadValue(ps_DataSource80.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group287_Click()
    On Error GoTo ErrHandler
        frsWriteValue "3", "DVSYS.60005-UUW1-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW1-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUV1_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group287_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap60_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap60.Description = frsReadValue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpClosePic_Click()
On Error GoTo ErrorHandler
    frszClosePictureButton Me
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Group296_Click()
    On Error GoTo ErrHandler
        frsWriteValue "1", "DVSYS.60005-UUW1-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW1-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUV1_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group296_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub Group295_Click()
    On Error GoTo ErrHandler
        frsWriteValue "2", "DVSYS.60005-UUW1-KMH/ARX_KMX_OIL.F_CV"
        frsWriteValue "1", "DVSYS.60005-UUW1-KMH/START_ARX.F_CV"
        frszExpertReplacePic 4, "WDU_N_60005_UUV1_KMH"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Group295_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
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

Private Sub TechCap62_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap57_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap57.Description = frsReadValue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub CPF_4_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap61_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
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

Private Sub SubCtrl49_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0771", 6
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

Private Sub ZD_CS_U_40_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_40
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap60_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
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

Private Sub Caption60_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption60.Description = frsReadValue(ps_DataSource81.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption60_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource81.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_U_41_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_41
Exit Sub
ErrHandler:
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
Private Sub TechCap51_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption49.Description = frsReadValue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap57_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption59_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption59.Description = frsReadValue(ps_DataSource80.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption59_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource80.CurrentValue
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

Private Sub Caption58_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption58.Description = frsReadValue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption58_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource79.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_17_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_17
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

Private Sub SubCtrl46_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0751", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption55_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption55.Description = frsReadValue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption55_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource75.CurrentValue
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

Private Sub Caption49_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap48_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption46.Description = frsReadValue(ps_DataSource25.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap48_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource25.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption46_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption46.Description = frsReadValue(ps_DataSource25.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption46_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource25.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_16_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_16
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_15_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_15
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_12_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_12
Exit Sub
ErrHandler:
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
Private Sub Rect36_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60006_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Rect36_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
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
Private Sub Polygon1595_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60005_00.grf"
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
Private Sub ZD_CS_U_8_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_8
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_2_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_13_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_4_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_4
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_6_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl3_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV008", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption3.Description = frsReadValue(ps_DataSource88.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap3_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource88.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption3.Description = frsReadValue(ps_DataSource88.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption3_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource88.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl5_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV009", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption5.Description = frsReadValue(ps_DataSource90.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap5_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource90.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption5.Description = frsReadValue(ps_DataSource90.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption5_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource90.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl7_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV010", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsReadValue(ps_DataSource92.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource92.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsReadValue(ps_DataSource92.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource92.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl9_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV011", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsReadValue(ps_DataSource23.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource23.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsReadValue(ps_DataSource23.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource23.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl14_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV022", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption14.Description = frsReadValue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap14_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption14.Description = frsReadValue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption14_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_1_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U__Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_
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

Private Sub bmpMode21_Click()
  On Error GoTo ErrHandler
    frsWriteValue "1", "DVSYS.60005-FIRE-74002/Reset.F_CV"
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

Private Sub SubCtrl_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV001", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption.Description = frsReadValue(ps_DataSource86.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource86.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption.Description = frsReadValue(ps_DataSource86.CurrentValue + "/DESC.A_CV")
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

Private Sub AI_TAG5_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource18.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource86.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub SubCtrl2_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV002", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsReadValue(ps_DataSource87.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap2_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource87.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsReadValue(ps_DataSource87.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG5.Description = frsReadValue(ps_DataSource18.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption2_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource87.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_10_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_10
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_11_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_12_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost35_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource18.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_15_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_15
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl11_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV018", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption11.Description = frsReadValue(ps_DataSource27.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost35_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap11_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource27.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption11.Description = frsReadValue(ps_DataSource27.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption11_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource27.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl12_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV016", 6
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

Private Sub TechCap12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption12.Description = frsReadValue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap12_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource29.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption12.Description = frsReadValue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
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

Private Sub Caption12_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource29.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl13_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV015", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption13.Description = frsReadValue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap13_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource42.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TAG2_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource10.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG2.Description = frsReadValue(ps_DataSource10.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption13.Description = frsReadValue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption13_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource42.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV021", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AAA_AI__Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_5_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_5
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsReadValue(ps_DataSource45.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TAG13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG13.Description = frsReadValue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG11_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource30.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG11.Description = frsReadValue(ps_DataSource30.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG10_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource28.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG10.Description = frsReadValue(ps_DataSource28.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG9_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource26.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG9.Description = frsReadValue(ps_DataSource26.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_12_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG8_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG8.Description = frsReadValue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_10_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_10
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AAA_AI_9_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG6_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource19.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG6.Description = frsReadValue(ps_DataSource19.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_8_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_8
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsReadValue(ps_DataSource45.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub AAA_AI_7_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AAA_AI_6_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_U_9_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_3_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_14_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_14
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub ZD_CS_U_5_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_5
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_7_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub SubCtrl4_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV020", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption4.Description = frsReadValue(ps_DataSource89.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource89.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE2_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption4.Description = frsReadValue(ps_DataSource89.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost19_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource10.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource89.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl6_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV019", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost19_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption6.Description = frsReadValue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE5_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap6_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource91.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource10.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost45_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource31.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption6.Description = frsReadValue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos4_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource18.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos4.Description = frsReadValue(ps_DataSource18.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost41_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource30.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption6_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource91.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost40_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource28.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost39_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource26.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl8_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV017", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsReadValue(ps_DataSource22.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost38_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource20.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost36_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource19.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap8_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource22.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsReadValue(ps_DataSource22.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption8_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource22.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl10_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV014", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption10.Description = frsReadValue(ps_DataSource24.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap10_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource24.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption10.Description = frsReadValue(ps_DataSource24.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption10_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource24.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV013", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsReadValue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost45_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub Caption15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsReadValue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Caption15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub SubCtrl1_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0731", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost41_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption1.Description = frsReadValue(ps_DataSource93.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap1_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource93.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost40_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption1.Description = frsReadValue(ps_DataSource93.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost39_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption1_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource93.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos13.Description = frsReadValue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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
Private Sub AI_VALUE13_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator1_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator1
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost38_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos10_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource30.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos10.Description = frsReadValue(ps_DataSource30.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub Text156_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE11_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos9_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource28.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos9.Description = frsReadValue(ps_DataSource28.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub rAI_Dost36_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE10_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE9_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos8_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource26.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos8.Description = frsReadValue(ps_DataSource26.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub




Private Sub Ctrl_32_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-PY-RD01", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub TechCap47_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap47.Description = frsReadValue(ps_DataSource15.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub TechCap47_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource15.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub DescCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap2.Description = frsReadValue(ps_DataSource15.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE8_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos7_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos7.Description = frsReadValue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub DescCap2_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource15.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos1.Description = frsReadValue(ps_DataSource10.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE6_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos5_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource19.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos5.Description = frsReadValue(ps_DataSource19.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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
