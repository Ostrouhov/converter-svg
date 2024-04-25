VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GKSVD_74_49041_UPTUG_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CPF_46900_Park_of_ChemReagents1_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "DeltaV.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

'This is the init event for our main picture.  It calls the initPicture code
'that is located in a our global project.
Private Sub Ctrl_40_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_40_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource78.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_40_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frsWriteValue "0", "DVSYS.MODULE/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub VLV_SR_3_Edit()
'Name: CPF_Type/VLV_SR_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo VLV_SR_3
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

Private Sub NavigateButtons_Edit()
On Error GoTo ErrorHandler
    frszopenPrevNextPic (Me.Name)
    Exit Sub
ErrorHandler:
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

Private Sub rAI_Alm153_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource78.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm153_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub Oval171_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval167_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval163_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval159_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub ZD_CS_U_12_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl8_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-BV52005", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsreadvalue(ps_DataSource67.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-UA006", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl17_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-UA005", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-UA004", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-UA003", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl7_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-BV52004", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl34_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-BV52003", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl33_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-BV52001", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl32_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-BV52002", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval186_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval190_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap8_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource67.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl38_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-PK006", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval180_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource141.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource141.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-PK005", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval176_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource140.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource140.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl36_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-PK004", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.49207-PK003", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval156_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource138.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource138.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval152_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource137.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource137.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_U_43_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_43
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsreadvalue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_23_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_23
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

Private Sub ZD_CS_U_21_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_21
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

Private Sub ZD_CS_U_20_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_20
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

Private Sub TechCap7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub Oval180_Edit()
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

Private Sub AAA_MB_24_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_24
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_21_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_21
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_23_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_23
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_20_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_20
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_22_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_22
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_28_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_28
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_27_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_27
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_26_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_26
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_25_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_25
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_MB_19_Edit()
'Name: WDU_N_FRS_MB/AAA_MB_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo AAA_MB_19
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval176_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsreadvalue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption17.Description = frsreadvalue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsreadvalue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap17_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsreadvalue(ps_DataSource74.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource75.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource74.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE26_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub TechCap25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsreadvalue(ps_DataSource95.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE25_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsreadvalue(ps_DataSource92.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap25_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource95.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsreadvalue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource92.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval156_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap23_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource91.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap22.Description = frsreadvalue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap21.Description = frsreadvalue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
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

Private Sub DescCap20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap20.Description = frsreadvalue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
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

Private Sub DescCap22_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap19.Description = frsreadvalue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap21_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG26_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource38.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE32_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource36.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE31_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DescCap20_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE30_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text510_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE27_Edit()
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

Private Sub AI_TAG26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG26.Description = frsreadvalue(ps_DataSource38.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DescCap19_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Text509_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text508_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text507_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG23.Description = frsreadvalue(ps_DataSource36.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap35.Description = frsreadvalue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG25_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap35_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval152_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG25.Description = frsreadvalue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsreadvalue(ps_DataSource67.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG32_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource32.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap34.Description = frsreadvalue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG22.Description = frsreadvalue(ps_DataSource35.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap34_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption8_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource67.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG31_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap33.Description = frsreadvalue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap33_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG30_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource30.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG32.Description = frsreadvalue(ps_DataSource32.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap32.Description = frsreadvalue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG24_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource28.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap32_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG27_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource29.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator21_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator21
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG31.Description = frsreadvalue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource27.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG30.Description = frsreadvalue(ps_DataSource30.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator20_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator20
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG24.Description = frsreadvalue(ps_DataSource28.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG27.Description = frsreadvalue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator19_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator19
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG21.Description = frsreadvalue(ps_DataSource27.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AO_Regilator18_Edit()
'Name: WDU_N_FRS/wdu_n_klp2_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo AO_Regilator18
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost44_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource38.CurrentValue, 6
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

Private Sub rAI_Dost24_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource28.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval181_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource141.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource141.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost24_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsreadvalue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost30_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource37.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost30_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval172_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource152.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-H", 6
    'frszDynamoLoadIAFB ps_DataSource152.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost32_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource32.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval168_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource151.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-L", 6
    'frszDynamoLoadIAFB ps_DataSource151.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsreadvalue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
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

Private Sub Oval177_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource140.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource140.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption17.Description = frsreadvalue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost31_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource31.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost43_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource36.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption17_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsreadvalue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
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

Private Sub rAI_Dost31_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource75.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsreadvalue(ps_DataSource74.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsreadvalue(ps_DataSource95.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption28_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource95.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource74.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsreadvalue(ps_DataSource92.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption27_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource92.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsreadvalue(ps_DataSource91.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption26_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource91.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval157_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource138.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource138.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval164_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource150.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-H", 6
    'frszDynamoLoadIAFB ps_DataSource150.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost28_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource30.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost27_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource29.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost23_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource27.CurrentValue, 6
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

Private Sub rAI_Dost29_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource35.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost27_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub rAI_Dost23_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval181_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval160_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource149.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-L", 6
    'frszDynamoLoadIAFB ps_DataSource149.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval177_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval182_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource141.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource141.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval178_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource140.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource140.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval153_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource137.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource137.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval157_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_47_Edit()
'Name: tipo_typical_v2/AAA_AI_ Version: 1.9
On Error GoTo ErrHandler
    frszdynamo AAA_AI_47
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionPos1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionPos1.Description = frsreadvalue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval158_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource138.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource138.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval182_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval153_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval187_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource153.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource153.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval178_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_46_Edit()
'Name: tipo_typical_v2/AAA_AI_ Version: 1.9
On Error GoTo ErrHandler
    frszdynamo AAA_AI_46
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval191_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource154.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource154.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval154_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource137.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource137.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval158_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_45_Edit()
'Name: tipo_typical_v2/AAA_AI_ Version: 1.9
On Error GoTo ErrHandler
    frszdynamo AAA_AI_45
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_44_Edit()
'Name: tipo_typical_v2/AAA_AI_ Version: 1.9
On Error GoTo ErrHandler
    frszdynamo AAA_AI_44
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval172_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub wdu_n_gas__Edit()
'Name: WDU_N_FRS/wdu_n_gas_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo wdu_n_gas_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode16_Click()
   On Error GoTo ErrHandler
    frsWriteValue "0", "DVSYS.49207-GAS/BLOCK.F_CV"
    frsWriteValue "1", "DVSYS.49207-GAS/HS5.F_CV"
Exit Sub
ErrHandler:
    frsHandleError

End Sub

Private Sub bmpMode17_Click()
  On Error GoTo ErrHandler
    If frsreadvalue("DVSYS.49207-GAS/HS5.F_CV") > 0 Then
      frsWriteValue "0", "DVSYS.49207-GAS/HS5.F_CV"
    Else
      frsWriteValue "1", "DVSYS.49207-GAS/BLOCK.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval168_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval154_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval164_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval160_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_49_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_49
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
    frszFPExpert "DVSYS.49207-B002", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_48_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_48
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval187_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval191_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    CaptionTech1.Description = frsreadvalue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CaptionTech1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub CaptionPos1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm173_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_46_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_46
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval179_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource141.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource141.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_45_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_45
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval179_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval155_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource138.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource138.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval175_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource140.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource140.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval173_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource152.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-H", 6
    'frszDynamoLoadIAFB ps_DataSource152.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval175_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval155_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval169_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource151.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-L", 6
    'frszDynamoLoadIAFB ps_DataSource151.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval173_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval169_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval151_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource137.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource137.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval165_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource150.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-H", 6
    'frszDynamoLoadIAFB ps_DataSource150.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval151_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval165_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval174_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource152.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-H", 6
    'frszDynamoLoadIAFB ps_DataSource152.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval161_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource149.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-L", 6
    'frszDynamoLoadIAFB ps_DataSource149.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval174_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval161_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval166_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource150.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-H", 6
    'frszDynamoLoadIAFB ps_DataSource150.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval166_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval170_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource151.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-L", 6
    'frszDynamoLoadIAFB ps_DataSource151.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval162_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource149.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-L", 6
    'frszDynamoLoadIAFB ps_DataSource149.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval170_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval162_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost60_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource83.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval188_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource153.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource153.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost59_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource82.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost58_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource80.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval192_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource154.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource154.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval188_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost57_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource79.CurrentValue, 6
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

Private Sub Oval192_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval189_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource153.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource153.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval193_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource154.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource154.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
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

Private Sub Oval189_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost58_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval193_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
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

Private Sub DI_N_BIG_3_Edit()
'Name: tipo_typical_v2/DI_N_BIG_ Version: 1.3
On Error GoTo ErrHandler
    frszdynamo DI_N_BIG_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Poly_R_ar5_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszExpertReplacePic 2, "CPF_44000_02_HydroCyclonesLine1.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Poly_R_ar5_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenPicExpert Version : 3.0
    frszOpenPicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value138_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link121_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AAA_AI_35_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_35
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_41_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_41
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_40_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_40
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_39_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_39
Exit Sub
ErrHandler:
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

Private Sub AAA_AI_36_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_36
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_34_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_34
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value137_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value136_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link120_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value135_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_60_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_60
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_N_BIG_4_Edit()
'Name: tipo_typical_v2/DI_N_BIG_ Version: 1.3
On Error GoTo ErrHandler
    frszdynamo DI_N_BIG_4
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect586_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource96.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect586_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_59_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_59
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect588_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource97.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_58_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_58
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_62_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_62
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG38_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource48.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DI_61_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_61
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG37_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource47.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Rect588_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG36_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource46.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG35_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG33_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DI_57_Edit()
'Name: tipo_typical_v2/DI_ Version: 1.1
On Error GoTo ErrHandler
    frszdynamo DI_57
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG34_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG29_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource42.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG28_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval167_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource151.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-L", 6
    'frszDynamoLoadIAFB ps_DataSource151.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link93_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_Link92_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval171_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource152.CurrentValue, 6
    'frszFPExpert "49207-PIA-39B-H", 6
    'frszDynamoLoadIAFB ps_DataSource152.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link91_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval163_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource150.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-H", 6
    'frszDynamoLoadIAFB ps_DataSource150.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval159_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource149.CurrentValue, 6
    'frszFPExpert "49207-PIA-31B-L", 6
    'frszDynamoLoadIAFB ps_DataSource149.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub DI_N_BIG__Edit()
'Name: tipo_typical_v2/DI_N_BIG_ Version: 1.3
On Error GoTo ErrHandler
    frszdynamo DI_N_BIG_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect564_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource93.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Rect564_Edit()
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
    frszFPExpert ps_DataSource48.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval190_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource154.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource154.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval186_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource153.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_DataSource153.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost53_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource47.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect566_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource94.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Rect566_Edit()
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
    frszFPExpert ps_DataSource46.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost50_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource44.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_N_BIG_1_Edit()
'Name: tipo_typical_v2/DI_N_BIG_ Version: 1.3
On Error GoTo ErrHandler
    frszdynamo DI_N_BIG_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost51_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource45.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG38_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG38.Description = frsreadvalue(ps_DataSource48.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link99_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_Link85_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Dost49_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource43.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG37_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG37.Description = frsreadvalue(ps_DataSource47.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost48_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource42.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost47_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource41.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG36_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG36.Description = frsreadvalue(ps_DataSource46.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link82_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG35.Description = frsreadvalue(ps_DataSource45.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link83_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG33.Description = frsreadvalue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link84_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link94_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG34.Description = frsreadvalue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link86_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG29.Description = frsreadvalue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Link87_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_Link88_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TAG28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG28.Description = frsreadvalue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_2_Edit()
'Name: tipo_typical/VLV_U_ Version: 3.4
On Error GoTo ErrHandler
    frszdynamo VLV_U_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_1_Edit()
'Name: tipo_typical/VLV_U_ Version: 3.4
On Error GoTo ErrHandler
    frszdynamo VLV_U_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Ctrl_9_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_9_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource7.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frsWriteValue "0", "DVSYS.MODULE/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub rAI_Alm99_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource7.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm99_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub Ctrl_8_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Ctrl_8_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource6.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Ctrl_8_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frsWriteValue "0", "DVSYS.MODULE/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub
Private Sub AI_VALUE34_Edit()
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

Private Sub AI_VALUE38_Edit()
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

Private Sub rAI_Alm98_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource6.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub rAI_Alm98_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub Oval183_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AI_VALUE37_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos26_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource48.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos26.Description = frsreadvalue(ps_DataSource48.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval184_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE29_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Oval185_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource12.CurrentValue, 6
    'frszFPExpert "MODULE", 6
    'frszDynamoLoadIAFB ps_dataSource5.CurrentValue, Me, frsVariables.gs_iafb_dc.CurrentValue
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE36_Edit()
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
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub AI_TechPos25_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource47.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos25.Description = frsreadvalue(ps_DataSource47.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE35_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos22.Description = frsreadvalue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Oval185_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos24_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource46.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos24.Description = frsreadvalue(ps_DataSource46.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_VALUE33_Edit()
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

Private Sub AI_TechPos23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource45.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos23.Description = frsreadvalue(ps_DataSource45.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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

Private Sub LO_LVL__Edit()
'Name: tipo_typical_v2/LO_LVL_ Version: 1.0
On Error GoTo ErrHandler
    frszdynamo LO_LVL_
Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos21.Description = frsreadvalue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos20_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource42.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos20.Description = frsreadvalue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost47_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_VALUE28_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos19_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
Private Sub AI_TechPos19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos19.Description = frsreadvalue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub
