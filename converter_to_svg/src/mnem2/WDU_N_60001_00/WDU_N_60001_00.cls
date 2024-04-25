VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WDU_N_60001_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Bitmap82_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60001_01.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap82_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap84_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60001_02.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap84_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    TechCap1.Description = frsreadvalue(ps_DataSource.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub ZD_CS_U_31_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_31
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl32_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1111", 6
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

Private Sub SubCtrl31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1101", 6
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


Private Sub ZD_CS_U_33_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_33
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption32.Description = frsreadvalue(ps_DataSource54.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap34_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource54.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption32.Description = frsreadvalue(ps_DataSource54.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption32_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource54.CurrentValue
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

Private Sub SubCtrl30_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1081", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-1171", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub SubCtrl34_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-1151", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption31.Description = frsreadvalue(ps_DataSource80.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl29_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1091", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap33_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource80.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption31.Description = frsreadvalue(ps_DataSource80.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl33_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-1141", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption31_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource80.CurrentValue
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

Private Sub ZD_CS_U_27_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_27
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap32_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption30.Description = frsreadvalue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl28_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1041", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap32_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub


Private Sub TechCap36_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption34.Description = frsreadvalue(ps_DataSource85.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption30.Description = frsreadvalue(ps_DataSource40.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption30_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource40.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource85.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap35_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption33.Description = frsreadvalue(ps_DataSource81.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap31_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption29.Description = frsreadvalue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl27_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1051", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap31_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource79.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap35_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource81.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl26_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1001", 6
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

Private Sub Caption29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption29.Description = frsreadvalue(ps_DataSource79.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption29_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource79.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap30_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsreadvalue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap30_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource78.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl25_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1011", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap29_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsreadvalue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub POST_1_Edit()
'Name: temp/POST_ Version: 3.0
On Error GoTo ErrHandler
    frszdynamo POST_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption28.Description = frsreadvalue(ps_DataSource78.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption33_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption33.Description = frsreadvalue(ps_DataSource81.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode18_Click()
   On Error GoTo ErrHandler
    frswritevalue "0", ps_DataSource15.CurrentValue + "/BLOCK.F_CV"
    frswritevalue "1", ps_DataSource15.CurrentValue + "/HS5.F_CV"
Exit Sub
ErrHandler:
    frsHandleError

End Sub

Private Sub bmpMode19_Click()
  On Error GoTo ErrHandler
    If frsreadvalue("DVSYS.60001-GAS/HS5.F_CV") > 0 Then
      frswritevalue "0", ps_DataSource15.CurrentValue + "/HS5.F_CV"
    Else
      frswritevalue "1", ps_DataSource15.CurrentValue + "/BLOCK.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption33_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource81.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption28_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource78.CurrentValue
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

Private Sub TechCap29_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption27.Description = frsreadvalue(ps_DataSource77.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub Caption27_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource77.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl24_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0961", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption34_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption34.Description = frsreadvalue(ps_DataSource85.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption34_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource85.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap28_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsreadvalue(ps_DataSource38.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap28_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource38.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption26.Description = frsreadvalue(ps_DataSource38.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption26_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource38.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Polygon1603_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60002_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1603_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl23_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0971", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsreadvalue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap27_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption25.Description = frsreadvalue(ps_DataSource76.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption25_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource76.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl22_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0931", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption24.Description = frsreadvalue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption23.Description = frsreadvalue(ps_DataSource74.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap26_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource75.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption24.Description = frsreadvalue(ps_DataSource75.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption24_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource75.CurrentValue
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

Private Sub TechCap25_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource74.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl21_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0921", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption22.Description = frsreadvalue(ps_DataSource73.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1061", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap24_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource73.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption23.Description = frsreadvalue(ps_DataSource74.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption23_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource74.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption22.Description = frsreadvalue(ps_DataSource73.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption22_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource73.CurrentValue
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

Private Sub TechCap23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption21.Description = frsreadvalue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap23_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption21.Description = frsreadvalue(ps_DataSource72.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption21_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource72.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1071", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption20.Description = frsreadvalue(ps_DataSource71.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap22_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource71.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption20.Description = frsreadvalue(ps_DataSource71.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption20_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource71.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1211", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption19.Description = frsreadvalue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap21_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub




Private Sub Text373_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub SubCtrl17_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1201", 6
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

Private Sub Caption19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption19.Description = frsreadvalue(ps_DataSource44.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption19_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource44.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1021", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsreadvalue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap20_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1031", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption17.Description = frsreadvalue(ps_DataSource70.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption18.Description = frsreadvalue(ps_DataSource43.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource43.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub TechCap19_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource70.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption17.Description = frsreadvalue(ps_DataSource70.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub





Private Sub Text372_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption17_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource70.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl14_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1191", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsreadvalue(ps_DataSource69.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap18_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource69.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption16.Description = frsreadvalue(ps_DataSource69.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource69.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl13_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0981", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsreadvalue(ps_DataSource68.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap17_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource68.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption15.Description = frsreadvalue(ps_DataSource68.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource68.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl12_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0991", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption14.Description = frsreadvalue(ps_DataSource67.CurrentValue + "/DESC.A_CV")
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

Private Sub TechCap16_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource67.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption14.Description = frsreadvalue(ps_DataSource67.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption14_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource67.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl11_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1181", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption13.Description = frsreadvalue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption12.Description = frsreadvalue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap15_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource42.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption13.Description = frsreadvalue(ps_DataSource42.CurrentValue + "/DESC.A_CV")
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

Private Sub SubCtrl10_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0941", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap14_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
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

Private Sub TechCap13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption11.Description = frsreadvalue(ps_DataSource65.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl9_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0951", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap13_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource65.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption12.Description = frsreadvalue(ps_DataSource66.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption12_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource66.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption11.Description = frsreadvalue(ps_DataSource65.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption11_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource65.CurrentValue
    Exit Sub
ErrorHandler:
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


Private Sub ZD_CS_U_7_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption10.Description = frsreadvalue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
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

Private Sub ZD_CS_U_6_Edit()
'Name: WDU_N_FRS_v2/ZD_CS_U_ Version: 3.9
On Error GoTo ErrHandler
    frszdynamo ZD_CS_U_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap12_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption10.Description = frsreadvalue(ps_DataSource41.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption10_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource41.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-1181", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub SubCtrl8_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1121", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsreadvalue(ps_DataSource64.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap11_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource64.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub SubCtrl36_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0911", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl7_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-0901", 6
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


Private Sub TechCap10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsreadvalue(ps_DataSource63.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption9.Description = frsreadvalue(ps_DataSource64.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub


Private Sub Caption9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource64.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap10_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource63.CurrentValue
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

Private Sub SubCtrl5_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60001-XV-1221", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub TechCap38_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption36.Description = frsreadvalue(ps_DataSource86.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsreadvalue(ps_DataSource62.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption8.Description = frsreadvalue(ps_DataSource63.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption8_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource63.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap38_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource86.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption36_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption36.Description = frsreadvalue(ps_DataSource86.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption36_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource86.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap9_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource62.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption7.Description = frsreadvalue(ps_DataSource62.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource62.CurrentValue
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

Private Sub SubCtrl4_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-0631", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption5.Description = frsreadvalue(ps_DataSource60.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
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

Private Sub TechCap7_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource60.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub bmpMode21_Click()
  On Error GoTo ErrHandler
    frswritevalue "1", "DVSYS.60001-FIRE/Reset.F_CV"
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub bmpMode20_Click()
  On Error GoTo ErrHandler
    If frsreadvalue("DVSYS.60001-FIRE/unblock.F_CV") = 1 Then
      frswritevalue "0", "DVSYS.60001-FIRE/unblock.F_CV"
    Else
      frswritevalue "1", "DVSYS.60001-FIRE/unblock.F_CV"
    End If
  Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub SubCtrl3_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60000-XV-1161", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption4.Description = frsreadvalue(ps_DataSource59.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap86_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60001_03.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap6_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource59.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption5.Description = frsreadvalue(ps_DataSource60.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption5_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource60.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption3.Description = frsreadvalue(ps_DataSource58.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap86_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap5_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource58.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption4.Description = frsreadvalue(ps_DataSource59.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource59.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource4.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsreadvalue(ps_DataSource57.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub



Private Sub Caption1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption1.Description = frsreadvalue(ps_DataSource.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap4_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource57.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub



Private Sub Caption3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption3.Description = frsreadvalue(ps_DataSource58.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption3_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource58.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Caption2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    Caption2.Description = frsreadvalue(ps_DataSource57.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Caption2_Click()
'вызов стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource57.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub





Private Sub Caption1_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Bitmap88_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60001_04.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap88_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap90_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60001_05.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub ZD_CS_L__Edit()
'Name: WDU_N_FRS_v2/ZD_CS_L_1 Version: 4.1
On Error GoTo ErrHandler
    frszdynamo ZD_CS_L_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Bitmap90_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
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

Private Sub Polygon1594_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_01_62000_Main.grf"
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

Private Sub Polygon1596_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60013_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1596_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1598_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60028_FND.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1598_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1599_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60019_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1599_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1600_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_00_62129_02_ESh3"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1600_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1601_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_01_62000_Main.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1601_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1602_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_00_62127"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Polygon1602_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect919_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60000_HR.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect919_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Rect929_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszExpertReplacePic 2, "WDU_N_60004_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_1_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_13_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_13
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_12_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption99_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource56.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_11_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption98_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource53.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_10_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_10
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption97_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource52.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_9_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource56.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost20_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption99_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption99.Description = frsreadvalue(ps_DataSource56.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption96_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource51.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_8_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_8
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption98_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption98.Description = frsreadvalue(ps_DataSource53.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos26_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource56.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos26_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos26.Description = frsreadvalue(ps_DataSource56.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption95_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource50.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value98_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_7_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource53.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos25_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource53.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos25_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos25.Description = frsreadvalue(ps_DataSource53.CurrentValue + "/DESC.A_CV")
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

Private Sub AI_Caption97_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption97.Description = frsreadvalue(ps_DataSource52.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption94_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value97_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_6_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost18_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource52.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos24_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource52.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos24_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos24.Description = frsreadvalue(ps_DataSource52.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost18_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption96_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption96.Description = frsreadvalue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption93_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value96_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_5_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_5
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost17_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource51.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos23_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource51.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos23_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos23.Description = frsreadvalue(ps_DataSource51.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost17_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption95_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption95.Description = frsreadvalue(ps_DataSource50.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption92_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource36.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value95_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_4_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_4
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource50.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos22_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource50.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos22_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos22.Description = frsreadvalue(ps_DataSource50.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost16_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption94_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption94.Description = frsreadvalue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption91_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value94_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption90_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource39.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos21_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource39.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos21_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos21.Description = frsreadvalue(ps_DataSource39.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost15_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption93_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption93.Description = frsreadvalue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_3_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value93_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption89_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost14_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource37.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos20_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource37.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos20_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos20.Description = frsreadvalue(ps_DataSource37.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost14_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost13_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource36.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost13_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value92_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption92_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption92.Description = frsreadvalue(ps_DataSource36.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_2_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value91_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_16_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_16
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos19_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource36.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos19_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos19.Description = frsreadvalue(ps_DataSource36.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost12_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource35.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost12_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption91_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption91.Description = frsreadvalue(ps_DataSource35.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption104_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource32.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value90_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption90_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption90.Description = frsreadvalue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos18_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource35.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos18_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos18.Description = frsreadvalue(ps_DataSource35.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption103_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource29.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_Small_15_Edit()
'Name: temp/AAA_AI_Small_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_Small_15
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost11_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource34.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost11_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption89_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption89.Description = frsreadvalue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost10_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource33.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost10_Edit()
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
    frszProcHist ps_DataSource34.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos17_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos17.Description = frsreadvalue(ps_DataSource34.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value89_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value88_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption102_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource28.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Rect929_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption101_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource27.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Caption104_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption104.Description = frsreadvalue(ps_DataSource32.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos16_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource33.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos16.Description = frsreadvalue(ps_DataSource33.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Caption103_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption103.Description = frsreadvalue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos15_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource32.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos15.Description = frsreadvalue(ps_DataSource32.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value103_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost37_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource32.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost37_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text297_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2.0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos14_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource29.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos14_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos14.Description = frsreadvalue(ps_DataSource29.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value102_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost36_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource29.CurrentValue, 6
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

Private Sub rAI_Dost35_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource28.CurrentValue, 6
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

Private Sub AI_Caption102_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption102.Description = frsreadvalue(ps_DataSource28.CurrentValue + "/DESC.A_CV")
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

Private Sub L_LINE_12_Edit()
'Name: tipo_typical_v2/L_LINE_ Version: 1.2
On Error GoTo ErrHandler
    frszdynamo L_LINE_12
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text354_Edit()
On Error GoTo ErrHandler
'Event added by frmPHVExpert
    frszPHVExpert
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Text355_Edit()
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

Private Sub Ctrl_32_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60004-PY-0582", 6
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

Private Sub Ctrl_31_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert "DVSYS.60004-PY-0592", 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub TechCap3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap3.Description = frsreadvalue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap3_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    TechCap2.Description = frsreadvalue(ps_DataSource30.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub TechCap2_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource30.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap2_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap2.Description = frsreadvalue(ps_DataSource31.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap2_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource31.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap1_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
On Error GoTo ErrorHandler
    DescCap1.Description = frsreadvalue(ps_DataSource30.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub DescCap1_Click()
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource30.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_Value101_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource28.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos13.Description = frsreadvalue(ps_DataSource28.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost34_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource27.CurrentValue, 6
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

Private Sub AI_Caption101_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_Caption101.Description = frsreadvalue(ps_DataSource27.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub L_LINE_11_Edit()
'Name: tipo_typical_v2/L_LINE_ Version: 1.2
On Error GoTo ErrHandler
    frszdynamo L_LINE_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub L_LINE_10_Edit()
'Name: tipo_typical_v2/L_LINE_ Version: 1.2
On Error GoTo ErrHandler
    frszdynamo L_LINE_10
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_Value100_Edit()
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
    frszProcHist ps_DataSource27.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrHandler
    AI_TechPos12.Description = frsreadvalue(ps_DataSource27.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub L_LINE_9_Edit()
'Name: tipo_typical_v2/L_LINE_ Version: 1.2
On Error GoTo ErrHandler
    frszdynamo L_LINE_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub L_LINE_8_Edit()
'Name: tipo_typical_v2/L_LINE_ Version: 1.2
On Error GoTo ErrHandler
    frszdynamo L_LINE_8
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

Private Sub AAA_AI_11_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_11
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

Private Sub AI_TAG12_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost33_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource20.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost33_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG12.Description = frsreadvalue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG11_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource19.CurrentValue
    Exit Sub
ErrorHandler:
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

Private Sub AI_TAG11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG11.Description = frsreadvalue(ps_DataSource19.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos11_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource20.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos11.Description = frsreadvalue(ps_DataSource20.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
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



Private Sub AAA_AI_9_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_9
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

Private Sub rAI_Dost32_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource19.CurrentValue, 6
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



Private Sub AI_TAG13_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource55.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG13.Description = frsreadvalue(ps_DataSource55.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub





Private Sub AI_TAG10_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource18.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG10.Description = frsreadvalue(ps_DataSource18.CurrentValue + "/DESC.A_CV")
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

Private Sub AAA_AI_7_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos10_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource19.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos10.Description = frsreadvalue(ps_DataSource19.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG9_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource16.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos27_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource55.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos27_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos27.Description = frsreadvalue(ps_DataSource55.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub





Private Sub rAI_Dost38_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource55.CurrentValue, 6
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


Private Sub rAI_Dost38_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub





Private Sub rAI_Dost31_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource18.CurrentValue, 6
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

Private Sub AI_TAG9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG9.Description = frsreadvalue(ps_DataSource16.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos9_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource18.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos9.Description = frsreadvalue(ps_DataSource18.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost30_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource16.CurrentValue, 6
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

Private Sub rAI_Dost30_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG8_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource9.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos8_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource16.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos8.Description = frsreadvalue(ps_DataSource16.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG8.Description = frsreadvalue(ps_DataSource9.CurrentValue + "/DESC.A_CV")
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

Private Sub AAA_AI_6_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG7_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource8.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG7.Description = frsreadvalue(ps_DataSource8.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost29_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource9.CurrentValue, 6
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
    frszProcHist ps_DataSource9.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos7.Description = frsreadvalue(ps_DataSource9.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG6_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource7.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG6.Description = frsreadvalue(ps_DataSource7.CurrentValue + "/DESC.A_CV")
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

Private Sub rAI_Dost28_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource8.CurrentValue, 6
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

Private Sub AI_VALUE7_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos6_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource8.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos6.Description = frsreadvalue(ps_DataSource8.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG5_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource6.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG5.Description = frsreadvalue(ps_DataSource6.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost27_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource7.CurrentValue, 6
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
    frszProcHist ps_DataSource7.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos5.Description = frsreadvalue(ps_DataSource7.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AAA_AI_3_Edit()
'Name: temp/AAA_AI_ Version: 2.4
On Error GoTo ErrHandler
    frszdynamo AAA_AI_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub AI_TAG4_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource5.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TAG4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
 On Error GoTo ErrHandler
    AI_TAG4.Description = frsreadvalue(ps_DataSource5.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost26_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource6.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost26_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
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

Private Sub AI_TechPos4_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource6.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos4.Description = frsreadvalue(ps_DataSource6.CurrentValue + "/DESC.A_CV")
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost25_Click()
On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource5.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Dost25_Edit()
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

Private Sub AI_TechPos3_Click()
'открытие стандартных графиков
On Error GoTo ErrorHandler
    frszProcHist ps_DataSource5.CurrentValue
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub AI_TechPos3_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
'всплывающая подсказка
On Error GoTo ErrorHandler
    AI_TechPos3.Description = frsreadvalue(ps_DataSource5.CurrentValue + "/DESC.A_CV")
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

Private Sub DI_9_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER13_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource26.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_8_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_8
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER12_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource25.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_7_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_7
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_6_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_6
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER13_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER13.Description = frsreadvalue(ps_DataSource26.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER11_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource24.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_5_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_5
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER12_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER12.Description = frsreadvalue(ps_DataSource25.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER10_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource23.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER9_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource21.CurrentValue, 6
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

Private Sub DI_11_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_4_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_4
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER16_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource87.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER15_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource84.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER11_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER11.Description = frsreadvalue(ps_DataSource24.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER8_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource17.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DI_3_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_3
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER10_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER10.Description = frsreadvalue(ps_DataSource23.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER16_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER16.Description = frsreadvalue(ps_DataSource87.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER7_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource13.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER15_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER15.Description = frsreadvalue(ps_DataSource84.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER9_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER9.Description = frsreadvalue(ps_DataSource21.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub DI_2_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_2
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER6_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource12.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER8_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER8.Description = frsreadvalue(ps_DataSource17.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub DI_1_Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_1
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER5_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource11.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER7_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER7.Description = frsreadvalue(ps_DataSource13.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub DI__Edit()
'Name: temp/DI_ Version: 2.0
On Error GoTo ErrHandler
    frszdynamo DI_
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER4_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource10.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub CLICKER6_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER6.Description = frsreadvalue(ps_DataSource12.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER5_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER5.Description = frsreadvalue(ps_DataSource11.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub CLICKER4_MouseMove(ByVal Button As Integer, ByVal Shift As Long, ByVal X As Double, ByVal Y As Double)
    CLICKER4.Description = frsreadvalue(ps_DataSource10.CurrentValue + "/DESC.A_CV")
End Sub

Private Sub NavigateButtons_Edit()
On Error GoTo ErrorHandler
    frszopenPrevNextPic (Me.Name)
    Exit Sub
ErrorHandler:
    frsHandleError
End Sub

Private Sub Ctrl_22_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_19_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_21_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_21_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource95.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_20_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_22_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource96.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frswritevalue "0", "DVSYS.PY-0000501/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub Ctrl_21_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frswritevalue "0", "DVSYS.PY-0000521/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub Ctrl_19_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource93.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_19_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frswritevalue "0", "DVSYS.PY-0000511/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub Ctrl_20_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource94.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub Ctrl_20_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Double, ByVal Y As Double)
If Button = vbKeyRButton Then
        frswritevalue "0", "DVSYS.PY-0000561/UNIT/STATE_CHANGE.F_CV"
    End If
End Sub

Private Sub rAI_Alm111_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource96.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm109_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource93.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm111_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm110_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource94.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm109_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm110_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_11_Edit()
'Name: TDA_Type/VLV_U_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo VLV_U_11
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm126_Click()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszFPExpert ps_DataSource95.CurrentValue, 6
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub rAI_Alm126_Edit()
    On Error GoTo ErrHandler
    'Event added by frmOpenFPExpert
    frszLoadFPExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_10_Edit()
'Name: TDA_Type/VLV_U_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo VLV_U_10
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_9_Edit()
'Name: TDA_Type/VLV_U_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo VLV_U_9
Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub VLV_U_8_Edit()
'Name: TDA_Type/VLV_U_ Version: 3.5
On Error GoTo ErrHandler
    frszdynamo VLV_U_8
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

Private Sub DeltaV_Click()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2,0
    frszExpertReplacePic 2, "WDU_N_60000_00.grf"
    Exit Sub
ErrHandler:
    frsHandleError
End Sub

Private Sub DeltaV_Edit()
    On Error GoTo ErrHandler
    'Event added by frmReplacePicExpert Version : 2,0
    frszReplacePicExpert
    Exit Sub
ErrHandler:
    frsHandleError
End Sub
