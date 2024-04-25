VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDU_N_60005_00"
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

Private Sub AAA_AI_11_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_11
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

Private Sub rAI_Dost28_Edit()
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

Private Sub CFixPicture_InitializeConfigure()
On Error GoTo ErrorHandler
    frszInitConfig Me
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

Private Sub CFixPicture_SelectionChanged()
On Error GoTo ErrorHandler
    frszpicCurLink Me.ContainedSelections, ps_CurrentLink
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl54_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0731", 6
    Exit Sub
ErrHandler:
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

Private Sub SubCtrl57_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV024", 6
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

Private Sub SubCtrl55_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0721", 6
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

Private Sub SubCtrl53_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0761", 6
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

Private Sub ZD_CS_U_42_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_42
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

Private Sub TechCap60_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap60.Description = frsReadValue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap59_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap59.Description = frsReadValue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap58_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap58.Description = frsReadValue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap56_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap56.Description = frsReadValue(ps_DataSource32.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap63_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
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

Private Sub TechCap55_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap55.Description = frsReadValue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_41_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_41
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

Private Sub TechCap61_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap59_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
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

Private Sub TechCap58_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub Caption61_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource82.CurrentValue
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

Private Sub TechCap56_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap51_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap55_Click()
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

Private Sub Caption57_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption57.Description = frsReadValue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption57_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource78.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption54_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption54.Description = frsReadValue(ps_DataSource32.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption54_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource32.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption56_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption56.Description = frsReadValue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption56_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_L_18_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_18
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

Private Sub ZD_CS_L_17_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_17
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption53_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption53.Description = frsReadValue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub Caption53_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_L_15_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_15
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

Private Sub ZD_CS_L_14_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_14
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

Private Sub ZD_CS_L_13_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L_11_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_11
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

Private Sub ZD_CS_L_10_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_10
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

Private Sub wdu_n_fire_4_Edit()
'Name: WDU_N_FRS/wdu_n_fire_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fire_4
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

Private Sub wdu_n_fire_1_Edit()
'Name: WDU_N_FRS/wdu_n_fire_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fire_1
Exit Sub
ErrHandler:
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

Private Sub SubCtrl44_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0741", 6
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

Private Sub Caption44_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub AAA_AI_38_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_38
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AAA_AI_37_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_37
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AAA_AI_36_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_36
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub Text282_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub Text281_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE39_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE38_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE37_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub rAI_Dost55_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource116.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost55_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub rAI_Dost54_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource115.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost54_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub rAI_Dost52_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource114.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost52_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG24_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource103.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG24.Description = frsReadValue(ps_DataSource103.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource102.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG23.Description = frsReadValue(ps_DataSource102.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource101.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG22.Description = frsReadValue(ps_DataSource101.CurrentValue + "/DESC.A_CV")
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

Private Sub AAA_AI_24_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_24
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

Private Sub AAA_AI_23_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_23
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


Private Sub AI_TAG17_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource70.CurrentValue
    Exit Sub
ErrorHandler:
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


Private Sub AAA_AI_21_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_21
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


Private Sub AI_TAG16_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource69.CurrentValue
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


Private Sub AI_TAG16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG16.Description = frsReadValue(ps_DataSource69.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AI_TAG13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource36.CurrentValue
    Exit Sub
ErrorHandler:
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


Private Sub AI_TAG13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG13.Description = frsReadValue(ps_DataSource36.CurrentValue + "/DESC.A_CV")
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

Private Sub AAA_AI_12_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_12
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



Private Sub AI_TAG8_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource100.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG8.Description = frsReadValue(ps_DataSource100.CurrentValue + "/DESC.A_CV")
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
Private Sub rAI_Dost62_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource103.CurrentValue, 6
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


Private Sub AI_TAG14_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource38.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG14.Description = frsReadValue(ps_DataSource38.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost61_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource102.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG15_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG15.Description = frsReadValue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost60_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource101.CurrentValue, 6
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

Private Sub rAI_Dost57_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource72.CurrentValue, 6
    Exit Sub
ErrHandler:
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

Private Sub AAA_AI_13_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub AAA_AI_14_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_14
Exit Sub
ErrHandler:
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


Private Sub bmpMode22_Click()
  On Error GoTo ErrHandler
    If frsReadValue("DVSYS.60005-FIRE-74003/unblock.F_CV") = 1 Then
      frsWriteValue "0", "DVSYS.60005-FIRE-74003/unblock.F_CV"
    Else
      frsWriteValue "1", "DVSYS.60005-FIRE-74003/unblock.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode23_Click()
  On Error GoTo ErrHandler
    frsWriteValue "1", "DVSYS.60005-FIRE-74003/Reset.F_CV"
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_15_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_15
Exit Sub
ErrHandler:
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
Private Sub rAI_Dost62_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub bmpMode16_Click()
  On Error GoTo ErrHandler
    If frsReadValue("DVSYS.60005-FIRE-74001/unblock.F_CV") = 1 Then
      frsWriteValue "0", "DVSYS.60005-FIRE-74001/unblock.F_CV"
    Else
      frsWriteValue "1", "DVSYS.60005-FIRE-74001/unblock.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost49_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource36.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub bmpMode17_Click()
  On Error GoTo ErrHandler
    frsWriteValue "1", "DVSYS.60005-FIRE-74001/Reset.F_CV"
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
Private Sub rAI_Dost61_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AAA_AI__Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_
Exit Sub
ErrHandler:
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
Private Sub rAI_Dost60_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AI_TAG7_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
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


Private Sub AAA_AI_5_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_5
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
Private Sub rAI_Dost57_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource103.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos23.Description = frsReadValue(ps_DataSource103.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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

Private Sub AI_TAG4_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource17.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos7_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource100.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub AI_TechPos22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource102.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos22.Description = frsReadValue(ps_DataSource102.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub rAI_Dost44_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource28.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost56_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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
Private Sub AI_VALUE23_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource101.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos21.Description = frsReadValue(ps_DataSource101.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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
Private Sub AI_VALUE22_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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
Private Sub rAI_Dost51_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AI_TechPos7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos7.Description = frsReadValue(ps_DataSource100.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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
Private Sub rAI_Dost49_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
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



Private Sub rAI_Dost42_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource30.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub AI_TAG7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG7.Description = frsReadValue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost43_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource38.CurrentValue, 6
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




Private Sub CaptionPos1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionPos1.Description = frsReadValue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionPos3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionPos3.Description = frsReadValue(ps_DataSource71.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost50_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource39.CurrentValue, 6
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




Private Sub AI_TAG5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG5.Description = frsReadValue(ps_DataSource18.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionPos2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionPos2.Description = frsReadValue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionPos_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionPos.Description = frsReadValue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
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


Private Sub AI_TAG4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG4.Description = frsReadValue(ps_DataSource17.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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
Private Sub AI_TechPos13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource36.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos13.Description = frsReadValue(ps_DataSource36.CurrentValue + "/DESC.A_CV")
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

Private Sub rAI_Dost19_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource10.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost48_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource20.CurrentValue, 6
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


Private Sub rAI_Dost35_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource18.CurrentValue, 6
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

Private Sub rAI_Dost48_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost53_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource100.CurrentValue, 6
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
Private Sub rAI_Dost44_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub rAI_Dost53_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
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

Private Sub rAI_Dost34_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource17.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost43_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost50_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub rAI_Dost42_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub wdu_n_fan_1_Edit()
'Name: WDU_N_FRS/wdu_n_fan_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fan_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-N-FV3", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost34_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost38_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource26.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm160_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub wdu_n_fan_3_Edit()
'Name: WDU_N_FRS/wdu_n_fan_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fan_3
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

Private Sub wdu_n_fan_2_Edit()
'Name: WDU_N_FRS/wdu_n_fan_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fan_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionTech1.Description = frsReadValue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub wdu_n_fan__Edit()
'Name: WDU_N_FRS/wdu_n_fan_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_fan_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos15_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos15.Description = frsReadValue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub rAI_Alm162_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE15_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Ctrl_37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-N-FV2", 6
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

Private Sub AI_TechPos14_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource38.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos14.Description = frsReadValue(ps_DataSource38.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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
Private Sub AI_VALUE14_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
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




Private Sub rAI_Alm161_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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


Private Sub CaptionPos3_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource71.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Ctrl_36_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-N-FV1", 6
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



Private Sub Ctrl_34_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-N-FV4", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionTech3.Description = frsReadValue(ps_DataSource71.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionTech2.Description = frsReadValue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech3_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource71.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CaptionPos2_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CaptionTech2_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CaptionPos1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub rAI_Alm159_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionPos_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CaptionTech_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionTech.Description = frsReadValue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub AI_VALUE7_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
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

Private Sub AI_TechPos6_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos6.Description = frsReadValue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
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

Private Sub AI_VALUE5_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE4_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
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

Private Sub AI_TechPos5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos5.Description = frsReadValue(ps_DataSource19.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos3_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource17.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos3.Description = frsReadValue(ps_DataSource17.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_38_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_38
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

Private Sub SubCtrl40_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0711", 6
    Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_36_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_36
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

Private Sub TechCap40_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption40.Description = frsReadValue(ps_DataSource65.CurrentValue + "/DESC.A_CV")
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

Private Sub SubCtrl37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV035", 6
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

Private Sub TechCap39_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption39.Description = frsReadValue(ps_DataSource64.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_33_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_33
Exit Sub
ErrHandler:
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

Private Sub SubCtrl35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV044", 6
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

Private Sub Text264_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
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

Private Sub AO_Regilator1_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator1
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

Private Sub ZD_CS_U_32_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_32
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DescCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap2.Description = frsReadValue(ps_DataSource15.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap37_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption37.Description = frsReadValue(ps_DataSource62.CurrentValue + "/DESC.A_CV")
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

Private Sub Caption37_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource62.CurrentValue
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

Private Sub TechCap36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource61.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption35.Description = frsReadValue(ps_DataSource94.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap35_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource94.CurrentValue
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

Private Sub Caption36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource61.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption34.Description = frsReadValue(ps_DataSource60.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap34_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource60.CurrentValue
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

Private Sub ZD_CS_U_29_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_29
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

Private Sub Text263_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub But_VKip6_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "WDU_N_60026.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-PY-RD02", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV048", 6
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

Private Sub DescCap1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap1.Description = frsReadValue(ps_DataSource14.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap1_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource14.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap42_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap42.Description = frsReadValue(ps_DataSource14.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap42_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource14.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_U_28_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_28
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

Private Sub ZD_CS_U_27_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_27
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

Private Sub SubCtrl30_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV043", 6
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

Private Sub SubCtrl29_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV041", 6
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

Private Sub ZD_CS_U_26_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_26
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl28_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV039", 6
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

Private Sub ZD_CS_U_25_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_25
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

Private Sub Caption30_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource56.CurrentValue
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

Private Sub ZD_CS_U_24_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_24
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

Private Sub SubCtrl26_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV033", 6
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

Private Sub ZD_CS_U_23_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_23
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

Private Sub SubCtrl25_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV040", 6
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

Private Sub ZD_CS_U_22_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_22
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

Private Sub SubCtrl24_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV032", 6
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

Private Sub ZD_CS_U_21_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_21
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

Private Sub TechCap25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsReadValue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
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

Private Sub SubCtrl23_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV038", 6
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

Private Sub ZD_CS_U_20_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_20
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

Private Sub SubCtrl22_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV031", 6
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

Private Sub Caption25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsReadValue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_19_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_19
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

Private Sub SubCtrl21_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV036", 6
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

Private Sub Caption24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption24.Description = frsReadValue(ps_DataSource50.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_18_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_18
Exit Sub
ErrHandler:
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

Private Sub SubCtrl20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV030", 6
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

Private Sub Caption23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption23.Description = frsReadValue(ps_DataSource49.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub Caption22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption22.Description = frsReadValue(ps_DataSource48.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_17_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_17
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

Private Sub ZD_CS_U_16_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_16
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

Private Sub SubCtrl19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV026", 6
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

Private Sub SubCtrl18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV025", 6
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

Private Sub Caption20_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_U_15_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_15
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

Private Sub Caption19_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource93.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub ZD_CS_U_14_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_14
Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_13_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_13
Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_12_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_12
Exit Sub
ErrHandler:
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

Private Sub TechCap16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsReadValue(ps_DataSource45.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub Caption16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
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

Private Sub TechCap15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsReadValue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_10_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_10
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

Private Sub SubCtrl12_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV016", 6
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

Private Sub ZD_CS_U_9_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_9
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

Private Sub SubCtrl11_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV018", 6
    Exit Sub
ErrHandler:
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

Private Sub SubCtrl10_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV014", 6
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

Private Sub ZD_CS_U_8_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_8
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

Private Sub SubCtrl9_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV011", 6
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

Private Sub Caption12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption12.Description = frsReadValue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_7_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_7
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

Private Sub ZD_CS_U_6_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_6
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

Private Sub DI_11_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_11
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


Private Sub DI_12_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_10_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_10
Exit Sub
ErrHandler:
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

Private Sub CLICKER15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource96.CurrentValue, 6
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


Private Sub CLICKER16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource97.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER14_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource95.CurrentValue, 6
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

Private Sub ZD_CS_U_5_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_5
Exit Sub
ErrHandler:
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

Private Sub TechCap9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsReadValue(ps_DataSource23.CurrentValue + "/DESC.A_CV")
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

Private Sub Caption9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsReadValue(ps_DataSource23.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER17.Description = frsReadValue(ps_DataSource98.CurrentValue + "/DESC.A_CV")
End Sub


Private Sub CLICKER15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER15.Description = frsReadValue(ps_DataSource96.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER16.Description = frsReadValue(ps_DataSource97.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub Caption9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource23.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CLICKER14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER14.Description = frsReadValue(ps_DataSource95.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub TechCap9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource23.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub SubCtrl6_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV019", 6
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

Private Sub ZD_CS_U_3_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_3
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

Private Sub SubCtrl5_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60005-XV-GV009", 6
    Exit Sub
ErrHandler:
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

Private Sub TechCap7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource92.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption6.Description = frsReadValue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap6_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource91.CurrentValue
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

Private Sub TechCap5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption5.Description = frsReadValue(ps_DataSource90.CurrentValue + "/DESC.A_CV")
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

Private Sub Caption6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption6.Description = frsReadValue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap5_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource90.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption4.Description = frsReadValue(ps_DataSource89.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource89.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption3.Description = frsReadValue(ps_DataSource88.CurrentValue + "/DESC.A_CV")
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

Private Sub Caption4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource89.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsReadValue(ps_DataSource87.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub Caption2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsReadValue(ps_DataSource87.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap2_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource87.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub Caption_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource86.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Rect1213_Click()
On Error GoTo ErrHandler
    frszExpertReplacePic 2, "WDU_N_60005_UVV2.grf"
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect1210_Click()
On Error GoTo ErrHandler
    frszExpertReplacePic 2, "WDU_N_60005_UVV1.grf"
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1605_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60007_00.grf"
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

Private Sub Polygon1604_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60011_00.grf"
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

Private Sub Polygon1595_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60018_00.grf"
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

Private Sub Rect1214_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60006_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect1214_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_23_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_24_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_24_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource68.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_23_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource67.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub NavigateButtons_Edit()
On Error GoTo ErrorHandler
    frszopenPrevNextPic (Me.Name)
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DeltaV_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60000_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DeltaV_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
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

Private Sub bmpPreviousPicture_Click()
On Error GoTo ErrorHandler
    frszopenPreviousPicture Me
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub heater_1_Edit()
'Name: tipo_typical_v2/heater_ Version: 1.5
On Error GoTo ErrHandler
    frszdynamo heater_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub heater_3_Edit()
'Name: tipo_typical_v2/heater_ Version: 1.5
On Error GoTo ErrHandler
    frszdynamo heater_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub
