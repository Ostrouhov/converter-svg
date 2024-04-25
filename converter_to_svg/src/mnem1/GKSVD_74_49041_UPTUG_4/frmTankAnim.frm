VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTankAnim 
   Caption         =   "Tank Dynamo"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   OleObjectBlob   =   "frmTankAnim.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1024"
End
Attribute VB_Name = "frmTankAnim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private NLSStrMgr As Object 'Important!! Do NOT move, edit, or remove this line!
'FormVersion: 1.01
Option Explicit
Public RefreshRate
Public DeadBand
Public Tank As Object
Public OldDataSource As String
Private AnimationObj As Object
Private frmDynamoColor As Object
Private blnError As Boolean
Private blnColorFormCancel As Boolean
Private blnColorFormShow As Boolean
Private blnFormActivate As Boolean
Private LoInTooHigh As Boolean
Private HiInTooHigh As Boolean
Private BadEntryMinPercentValue As Boolean
Private MinPercentTooHigh As Boolean
Private MaxPercentTooHigh As Boolean
Public Sub InitializeDynamo(DynamoName As Object)
    Set Tank = DynamoName
End Sub
Private Sub LaunchColorByForm()
    Dim TankLevelColorObj As Object
    Dim blnHasConnection As Boolean
    Dim lngIndex As Long
    Dim lngStatus As Long
    
    'If the Dynamo ColorBy form has just been activated, we don't want to launch the color by form.
    If blnColorFormShow = False Then
        'Set the flag as to whether the form was shown to true then, copy a local instance of the DynamoColorBy form
        blnColorFormShow = True
        GetFormDynamoColor frmDynamoColor
    End If
    Set TankLevelColorObj = FindLocalObject(Tank, "TankLevelColor")
    frmDynamoColor.InitializeColorByForm TankLevelColorObj, frmTankAnim, blnColorFormCancel
    frmDynamoColor.Show
    If frmDynamoColor.blnCanceled = False Then
        lblTankColorDataSource.Caption = frmDynamoColor.ExpressionEditor1.EditText
    End If
    'Now that the form has been activated, we can set the flag to false and allow launching
    'of the ColorBy form
    blnFormActivate = False
    
    'If the user did not make a connection to animate color, uncheck the Animate Tank Level Color checkbox
    TankLevelColorObj.IsConnected "ForegroundColor", blnHasConnection, lngIndex, lngStatus
    If blnHasConnection = False Then
        cbxAnimateTankLevelColor.Value = False
    End If
End Sub
Private Sub cbxAnimateTankLevelColor_Click()
    If blnFormActivate = False Then
        If cbxAnimateTankLevelColor.Value = True Then
            LaunchColorByForm
        End If
    End If
    blnFormActivate = False
End Sub
Private Sub cmdColorBy_Click()
    If cbxAnimateTankLevelColor.Value <> True Then
        cbxAnimateTankLevelColor.Value = True
    Else
        LaunchColorByForm
    End If
End Sub
Private Sub cbxFetchLimits_Click()
    On Error GoTo ErrorHandler
    If cbxFetchLimits.Value = True Then
        lblHiIn.Enabled = False
        txtHiIn.Enabled = False
        lblLoIn.Enabled = False
        txtLoIn.Enabled = False
    Else
        lblHiIn.Enabled = True
        txtHiIn.Enabled = True
        lblLoIn.Enabled = True
        txtLoIn.Enabled = True
    End If
    Exit Sub
    
ErrorHandler:
    HandleError
End Sub
Private Sub ExpressionEditor1_AfterKillFocus()
    If ExpressionEditor1.EditText <> "" Then
        Dim err As Integer
        err = HandleFetchLimits()
        If (err <> 0) Then
            ExpressionEditor1.SetFocusToComboBox
            MsgBox NLSStrMgr.GetNLSStr(1000)
        End If
    End If
End Sub
Private Function HandleFetchLimits() As String
      If ExpressionEditor1.EditText <> OldDataSource Then
            Dim err As Integer
            Dim HiLimit As Single, LoLimit As Single
            Dim ds As String
            ds = ExpressionEditor1.EditText
            Call FetchLimits(ds, HiLimit, LoLimit, err)
            If (err = 0) Then
                txtLoIn.Value = LoLimit
                txtHiIn.Value = HiLimit
                OldDataSource = ds
                ExpressionEditor1.EditText = ds
            ElseIf (err = 2) Then ' Valid source may not exist? Just do nothing
                err = 0
            End If
            
      End If
      HandleFetchLimits = err
End Function
Private Sub UserForm_Activate()
'Initialize form interface
    Me.clrTankForegroundColor.color = GetTankForegroundColor()
    Me.clrTankBackgroundColor.color = GetTankBackgroundColor()
    Me.clrTankColor.color = GetTankColor()
    Me.ExpressionEditor1.EditText = GetTankDataSource()
    OldDataSource = ExpressionEditor1.EditText
    ExpressionEditor1.RefreshRate = GetTankRefreshRate()
    ExpressionEditor1.DeadBand = GetTankDeadBand()
    
    'Set up the Lo and Hi input and output for tank level
    Dim LoIn As Single, HiIn As Single, LoOut As Single, HiOut As Single
    Call GetInOutValues(LoIn, HiIn, LoOut, HiOut)
    txtLoIn.Text = LoIn
    txtHiIn.Text = HiIn
    txtLoOut.Text = LoOut
    txtHiOut.Text = HiOut
    ExpressionEditor1.SetFocusToComboBox
    
    Dim DataSource As Object
    Set DataSource = FindLocalObject(Tank, "AnimatedTankLevel")
    If DataSource.Autofetch = True Then
        cbxFetchLimits.Value = True
        lblHiIn.Enabled = False
        txtHiIn.Enabled = False
        lblLoIn.Enabled = False
        txtLoIn.Enabled = False
    Else
        cbxFetchLimits.Value = False
        lblHiIn.Enabled = True
        txtHiIn.Enabled = True
        lblLoIn.Enabled = True
        txtLoIn.Enabled = True
    End If
    
    'Set up the tank level color by animation
    Dim TankLevelColorObj As Object
    Dim blnHasConnection As Boolean
    Dim lIndex As Long
    Dim lStatus As Long
    Dim strPropertyName As String
    Dim strExpression As String
    Dim strFullyQualifiedExpression As String
    Dim vtAnimationObjects
    
    Set TankLevelColorObj = FindLocalObject(Tank, "TankLevelColor")
    'Determine if the tank level foreground color is animated
    TankLevelColorObj.IsConnected "ForegroundColor", blnHasConnection, lIndex, lStatus
    If blnHasConnection = True Then
        'Set a flag so the Dynamo ColorBy form does not get launced
        blnFormActivate = True
        cbxAnimateTankLevelColor.Value = True
        TankLevelColorObj.GetConnectionInformation lIndex, strPropertyName, strExpression, strFullyQualifiedExpression, vtAnimationObjects
        lblTankColorDataSource.Caption = vtAnimationObjects(0).Source
    End If
End Sub
Public Sub SetTankDataSource(DataSource As String)
    Dim TankDataSource As Object
    
    On Error GoTo ErrorHandler
    Set TankDataSource = FindLocalObject(Tank, "AnimatedTankLevel")
    TankDataSource.SetSource DataSource, True, ExpressionEditor1.RefreshRate, ExpressionEditor1.DeadBand
    Exit Sub
ErrorHandler:
    blnError = True
    If err.Number = -2147200603 Then
        MsgBox NLSStrMgr.GetNLSStr(1001)
        ExpressionEditor1.SetFocusToComboBox
    Else
        HandleError
    End If
End Sub
Public Sub SetTankColor(color As Long)
    Dim hue As Double, sat As Double, lum As Double
    Dim TankColor As Long
    TankColor = GetTankColor()
    If TankColor <> color Then
        Dim TankShell As Object
        Set TankShell = FindLocalObject(Tank, "TankMain")
        Call frmAdvancedColor.SetGroupGradiant(TankShell.ContainedObjects, color)
    End If
End Sub
Public Sub SetTankLevelForegroundColor(ForegroundColor As Long)
    Dim TankColor As Object
    Dim strTableName As String
    Dim strColorName As String
    Dim blnSys As Boolean
    
    Set TankColor = FindLocalObject(Tank, "TankLevelColor")
    clrTankForegroundColor.GetIndirectionInfo strTableName, strColorName, blnSys
    'If the user selected a named color, get the named color information
    If strTableName <> "" Then
        TankColor.SetupPropertyIndirection "ForegroundColor", strTableName, strColorName
    End If
    TankColor.ForegroundColor = ForegroundColor
End Sub
Public Sub SetTankLevelBackgroundColor(BackgroundColor As Long)
    Dim TankColor As Object
    Dim strTableName As String
    Dim strColorName As String
    Dim blnSys As Boolean
    
    Set TankColor = FindLocalObject(Tank, "TankLevelColor")
    clrTankBackgroundColor.GetIndirectionInfo strTableName, strColorName, blnSys
    'If the user selected a named color, get the named color information
    If strTableName <> "" Then
        TankColor.SetupPropertyIndirection "BackgroundColor", strTableName, strColorName
    End If
    TankColor.BackgroundColor = BackgroundColor
End Sub
Public Sub SetInOutValues(LowIn As Single, HighIn As Single, LowOut As Single, HighOut As Single)
    Dim TankDataSource As Object
    Set TankDataSource = FindLocalObject(Tank, "AnimatedTankLevel")
    TankDataSource.Autofetch = False
    TankDataSource.LoInValue = LowIn
    TankDataSource.HiInValue = HighIn
    TankDataSource.LoOutValue = LowOut
    TankDataSource.HiOutValue = HighOut
End Sub
Public Sub SetAutoFetchInputLimits(LowOut As Single, HighOut As Single)
    Dim TankDataSource As Object
    Set TankDataSource = FindLocalObject(Tank, "AnimatedTankLevel")
    TankDataSource.Autofetch = True
    TankDataSource.LoOutValue = LowOut
    TankDataSource.HiOutValue = HighOut
End Sub
Public Function GetTankDataSource() As String
    Dim AnimatedTankLevel As Object
    Set AnimatedTankLevel = FindLocalObject(Tank, "AnimatedTankLevel")
    GetTankDataSource = AnimatedTankLevel.Source
End Function
    
Public Function GetTankColor() As Long
    Dim TankColor As Object
    Set TankColor = FindLocalObject(Tank, "TankColor")
    GetTankColor = TankColor.ForegroundColor
End Function
    
Public Function GetTankForegroundColor() As Long
    Dim TankColor As Object
    Dim strTableName As String
    Dim strColorName As String
    Dim blnSys As Boolean
    Dim blnIsIndirected As Boolean
    Dim lngEntry As Long
    
    Set TankColor = FindLocalObject(Tank, "TankLevelColor")
    blnIsIndirected = TankColor.IsPropertyIndirected("ForegroundColor")
    If blnIsIndirected = True Then
        TankColor.GetIndirectedProperty "ForegroundColor", strTableName, strColorName
        blnSys = clrTankForegroundColor.SetIndirectionInfo(strTableName, strColorName, lngEntry)
    End If
    GetTankForegroundColor = TankColor.ForegroundColor
End Function
Public Function GetTankBackgroundColor() As Long
    Dim TankColor As Object
    Dim strTableName As String
    Dim strColorName As String
    Dim blnSys As Boolean
    Dim blnIsIndirected As Boolean
    Dim lngEntry As Long
    
    Set TankColor = FindLocalObject(Tank, "TankLevelColor")
    blnIsIndirected = TankColor.IsPropertyIndirected("BackgroundColor")
    If blnIsIndirected = True Then
        TankColor.GetIndirectedProperty "BackgroundColor", strTableName, strColorName
        blnSys = clrTankBackgroundColor.SetIndirectionInfo(strTableName, strColorName, lngEntry)
    End If
    GetTankBackgroundColor = TankColor.BackgroundColor
End Function
Public Sub GetInOutValues(LowIn As Single, HighIn As Single, LowOut As Single, HighOut As Single)
    Dim TankDataSource As Object
    Set TankDataSource = FindLocalObject(Tank, "AnimatedTankLevel")
    LowIn = TankDataSource.LoInValue
    HighIn = TankDataSource.HiInValue
    LowOut = TankDataSource.LoOutValue
    HighOut = TankDataSource.HiOutValue
End Sub
Private Sub cmdCancel_Click()
    Dim msgResult
    Dim TankLevelColorObj As Object
    Dim blnHasConnection As Boolean
    Dim lngIndex As Long
    Dim lngStatus As Long
    If TypeName(frmDynamoColor) <> "Nothing" Then
        Unload frmDynamoColor
    End If
    'Select the Tank group
    Tank.SelectObject (True)
    Unload Me
End Sub
Private Sub cmdOK_Click()
    On_OK
    If blnError = True Then
        blnError = False
        Exit Sub
    End If
    'Select the Tank group
    Tank.SelectObject (True)
    Unload Me
End Sub
Private Sub On_OK()
    Dim lngStatus As Long
    Dim lngIndex As Long
    Dim vtValidObjects
    Dim vtUndefinedObjects
    Dim strFullyQualifiedExpression As String
    Dim vtResults
    Dim vtAttributes
    Dim msgNoSource
    
    ' Check to make sure the user entered something
    If ExpressionEditor1.EditText = "" Then
        ExpressionEditor1.SetFocusToComboBox
        MsgBox (NLSStrMgr.GetNLSStr(1002))
        blnError = True
        Exit Sub
    End If
    
    ' Check the Data source
    Dim ret As Integer
    ret = QuickAdd(ExpressionEditor1.EditText)
    Select Case ret
        Case 0  'Data Source is valid or nothing is entered.
            Dim err As Integer
            err = HandleFetchLimits
            If err <> 0 Then
                ExpressionEditor1.SetFocusToComboBox
                MsgBox (NLSStrMgr.GetNLSStr(1002))
                Exit Sub
            End If
            ExpressionEditor1.SaveToHistoryList (ExpressionEditor1.EditText)
        Case 1  'Invalid Data Source syntax
            ExpressionEditor1.SetFocusToComboBox
            MsgBox (NLSStrMgr.GetNLSStr(1003))
            blnError = True
            Exit Sub
        Case 2  'User Performed QuickAdd (now ds is OK)
            System.ParseConnectionSource "Name", ExpressionEditor1.EditText, lngIndex, vtValidObjects, vtUndefinedObjects, strFullyQualifiedExpression
            System.GetPropertyAttributes strFullyQualifiedExpression, 0, vtResults, vtAttributes, lngStatus
            If lngStatus <> 0 Then
            msgNoSource = MsgBox(NLSStrMgr.GetNLSStr(1004), vbYesNo)
                If msgNoSource = vbNo Then
                    ExpressionEditor1.SetFocusToComboBox
                    blnError = True
                    Exit Sub
                End If
            End If
            Call HandleFetchLimits
            ExpressionEditor1.SetFocusToComboBox
        Case 3 'Type of data source does not match property being animated
            MsgBox (NLSStrMgr.GetNLSStr(1005))
            ExpressionEditor1.SetFocusToComboBox
            blnError = True
            Exit Sub
        Case 4  'Use Anyway on datasource (Could not validate tag, user said OK)
        Case 5  'Could not validate tag, user said do not use
            ExpressionEditor1.SetFocusToComboBox
            blnError = True
            Exit Sub
    End Select
    
    Call SetTankDataSource(ExpressionEditor1.EditText)
    Call SetTankColor(clrTankColor.color)
    Call SetTankLevelForegroundColor(clrTankForegroundColor.color)
    Call SetTankLevelBackgroundColor(clrTankBackgroundColor.color)
    If cbxFetchLimits.Value = True Then
        Call SetAutoFetchInputLimits(txtLoOut.Text, txtHiOut.Text)
    Else
        Call SetInOutValues(txtLoIn.Text, txtHiIn.Text, txtLoOut.Text, txtHiOut.Text)
    End If
    
    Dim TankLevelColorObj As Object
    If TypeName(TankLevelColorObj) <> "Nothing" Then
        If cbxAnimateTankLevelColor.Value = False Then
            'If the user wants to remove a color animation, get an instance of the Dynamo
            'color form if one does not already exist and initialize it.  Then call the
            'KillColorAnimationObject Subroutine to remove the connection
            If TypeName(frmDynamoColor) = "Nothing" Then
                GetFormDynamoColor frmDynamoColor
                Set TankLevelColorObj = FindLocalObject(Tank, "TankLevelColor")
                frmDynamoColor.InitializeColorByForm TankLevelColorObj, frmTankAnim, blnColorFormCancel
            End If
            Call frmDynamoColor.KillColorAnimationObject(TankLevelColorObj)
        End If
    End If
    
    If TypeName(frmDynamoColor) <> "Nothing" Then
        Unload frmDynamoColor
    End If
End Sub
Public Function GetTankRefreshRate()
Dim Tolerance As Double
Dim TankLevelObj As Object
Dim blnHasConnection1 As Boolean
Dim lIndex1 As Long
Dim lStatus1 As Long
Dim strPropertyName1 As String
Dim strExpression1 As String
Dim strFullyQualifiedExpression1 As String
Dim vtAnimationObjects1
Set TankLevelObj = FindLocalObject(Tank, "AnimatedTankLevel")
TankLevelObj.IsConnected "InputValue", blnHasConnection1, lIndex1, lStatus1
TankLevelObj.GetConnectionInformation lIndex1, strPropertyName1, strExpression1, strFullyQualifiedExpression1, vtAnimationObjects1, Tolerance, DeadBand, RefreshRate
GetTankRefreshRate = RefreshRate
End Function
Public Function GetTankDeadBand()
    GetTankDeadBand = DeadBand
End Function
Private Sub UserForm_Initialize()
    Set NLSStrMgr = CreateObject("frmTankAnimRES.NLSStrMgr") 'Important!! Do NOT move, edit, or remove this line!
    NLSStrMgr.NLSContainer Me 'Important!! Do NOT move, edit, or remove this line!
End Sub
