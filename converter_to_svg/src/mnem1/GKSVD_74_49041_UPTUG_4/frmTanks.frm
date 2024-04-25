VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTanks 
   Caption         =   "Static Tank Dynamo Tank Color"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   OleObjectBlob   =   "frmTanks.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1005"
End
Attribute VB_Name = "frmTanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private NLSStrMgr As Object 'Important!! Do NOT move, edit, or remove this line!
'FormVersion: 1.0
Option Explicit
Public TankObject As Object
Public Sub InitializeDynamo(DynamoName As Object)
    Set TankObject = DynamoName
End Sub
Private Sub UserForm_Activate()
    Me.clrTankColor.color = GetTankColor()
End Sub
Public Function GetTankColor() As Long
    Dim TankColor As Object
    Set TankColor = FindLocalObject(TankObject, "TankColor")
    GetTankColor = TankColor.ForegroundColor
End Function
Public Sub SetTankColor(color As Long)
    Dim hue As Double, sat As Double, lum As Double
    Dim TankColor As Long
    TankColor = GetTankColor()
    If TankColor <> color Then
        Call frmAdvancedColor.SetGroupGradiant(TankObject.ContainedObjects, color)
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call SetTankColor(clrTankColor.color)
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Set NLSStrMgr = CreateObject("frmTanksRES.NLSStrMgr") 'Important!! Do NOT move, edit, or remove this line!
    NLSStrMgr.NLSContainer Me 'Important!! Do NOT move, edit, or remove this line!
End Sub
