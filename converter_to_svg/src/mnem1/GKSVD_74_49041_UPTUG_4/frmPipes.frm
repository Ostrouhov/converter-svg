VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPipes 
   Caption         =   "Pipe Dynamo Pipe Color"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   OleObjectBlob   =   "frmPipes.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1003"
End
Attribute VB_Name = "frmPipes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private NLSStrMgr As Object 'Important!! Do NOT move, edit, or remove this line!
'FormVersion: 1.0
Option Explicit
Public Pipe As Object
Public Sub InitializeDynamo(DynamoName As Object)
    Set Pipe = DynamoName
End Sub
Private Sub UserForm_Activate()
    Me.clrPipeColor.color = GetPipeColor()
End Sub
Public Function GetPipeColor() As Long
    Dim PipeColor As Object
    Set PipeColor = FindLocalObject(Pipe, "PipeColor")
    GetPipeColor = PipeColor.ForegroundColor
End Function
Public Sub SetPipeColor(color As Long)
    Dim hue As Double, sat As Double, lum As Double
    Dim PipeColor As Long
    PipeColor = GetPipeColor()
    If PipeColor <> color Then
        Call frmAdvancedColor.SetGroupGradiant(Pipe.ContainedObjects, color)
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Call SetPipeColor(clrPipeColor.color)
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Set NLSStrMgr = CreateObject("frmPipesRES.NLSStrMgr") 'Important!! Do NOT move, edit, or remove this line!
    NLSStrMgr.NLSContainer Me 'Important!! Do NOT move, edit, or remove this line!
End Sub
