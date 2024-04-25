VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdvancedColor 
   Caption         =   "This is a container of methods dealing with color."
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   OleObjectBlob   =   "frmAdvancedColor.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1001"
End
Attribute VB_Name = "frmAdvancedColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private NLSStrMgr As Object 'Important!! Do NOT move, edit, or remove this line!
'FormVersion: 1.0
Public Sub SetGroupGradiant(CObj As Object, color As Long)
Dim Shape As Object
Dim LumMin As Double, LumMax As Double, LumRange As Double
Dim ObjLum As Double, NewLumMin As Double, NewLumRange As Double, NewLum As Double
Dim i As Integer
Dim h As Double, l As Double, s As Double
Call GetHLS(color, h, l, s)
Call GetLumInfo(CObj, LumMin, LumMax)
LumRange = LumMax - LumMin
NewLumMin = l
NewRange = 1 - l
For i = 1 To CObj.Count
    Set Shape = CObj.item(i)
    If (Shape.ClassName = "Group") Then
        Call SetGroupGradiant(Shape.ContainedObjects, color)
    Else
        On Error Resume Next
        ObjLum = GetLum(Shape.ForegroundColor)
        If LumRange = 0 Then
            NewLum = NewLumMin
        Else
            NewLum = ((ObjLum - LumMin) / LumRange) * NewRange + NewLumMin
        End If
        Shape.ForegroundColor = SetHLS(h, NewLum, s)
    End If
Next i
End Sub
Private Sub GetLumInfo(CObj As Object, min As Double, max As Double)
Dim Shape As Object
Dim i As Integer, min1 As Double, max1 As Double, lum As Double
min = 1
max = 0
For i = 1 To CObj.Count
    Set Shape = CObj.item(i)
    If (Shape.ClassName = "Group") Then
        Call GetLumInfo(Shape.ContainedObjects, min1, max1)
        If (min1 < min) Then
          min = min1
        End If
        If (max1 > max) Then
          max = max1
        End If
    Else
        On Error Resume Next
        lum = GetLum(Shape.ForegroundColor)
        If (lum > max) Then
            max = lum
        End If
        If (lum < min) Then
            min = lum
        End If
    End If
Next i
End Sub
Public Sub SetObjHueSat(CObj As Object, h As Double, s As Double)
Dim Shape As Object
Dim i As Integer
For i = 1 To CObj.Count
    Set Shape = CObj.item(i)
    If (Shape.ClassName = "Group") Then
        Call SetObjHueSat(Shape.ContainedObjects, h, s)
    Else
        On Error Resume Next
        Shape.ForegroundColor = SetHueSat(Shape.ForegroundColor, h, s)
    End If
Next i
End Sub
Public Sub SetObjHue(CObj As Object, h As Double)
Dim Shape As Object
Dim i As Integer
For i = 1 To CObj.Count
    Set Shape = CObj.item(i)
    If (Shape.ClassName = "Group") Then
        Call SetObjHue(Shape.ContainedObjects, h)
    Else
        On Error Resume Next
        Shape.ForegroundColor = SetHue(Shape.ForegroundColor, h)
    End If
Next i
End Sub
Public Sub SetObjSat(CObj As Object, s As Double)
Dim Shape As Object
Dim i As Integer
For i = 1 To CObj.Count
    Set Shape = CObj.item(i)
    If (Shape.ClassName = "Group") Then
        Call SetObjSat(Shape.ContainedObjects, s)
    Else
        On Error Resume Next
        Shape.ForegroundColor = SetSat(Shape.ForegroundColor, s)
    End If
Next i
End Sub
Private Sub HLS_To_RGB(r As Long, g As Long, b As Long, hue As Double, lum As Double, sat As Double)
Dim m2 As Double
Dim ml As Double
Dim RValue As Double
Dim GValue As Double
Dim BValue As Double
If (lum <= 0.5) Then
    m2 = lum * (1 + sat)
Else
    m2 = lum + sat - (lum * sat)
End If
ml = 2 * lum - m2
If (sat = 0) Then
    If (hue = -1) Then
        RValue = lum
        GValue = lum
        BValue = lum
    Else
        MsgBox NLSStrMgr.GetNLSStr(1000)
    End If
    
Else
    RValue = hls_func(ml, m2, hue + 120)
    GValue = hls_func(ml, m2, hue)
    BValue = hls_func(ml, m2, hue - 120)
End If
' Convert to RGBMAX (255)
r = RValue * 255
g = GValue * 255
b = BValue * 255
End Sub
Private Function hls_func(ByVal nl As Double, ByVal n2 As Double, ByVal hue As Double) As Double
    If (hue > 360) Then
        hue = hue - 360
    Else
        If hue < 0 Then
            hue = hue + 360
        End If
    End If
    If (hue < 60) Then
        hls_func = nl + (n2 - nl) * hue / 60
    Else
        If (hue < 180) Then
            hls_func = n2
        Else
    
            If (hue < 240) Then
                hls_func = nl + (n2 - nl) * (240 - hue) / 60
            Else
                hls_func = nl
            End If
        End If
    
    End If
End Function
Private Sub RGB_To_HLS(ByVal r As Long, ByVal g As Long, ByVal b As Long, hue As Double, lum As Double, sat As Double)
    Dim min As Double
    Dim max As Double
    Dim RValue As Double
    Dim GValue As Double
    Dim BValue As Double
    RValue = r / 255#
    GValue = g / 255#
    BValue = b / 255#
    max = maximum(RValue, GValue, BValue)
    min = minimum(RValue, GValue, BValue)
    lum = (max + min) / 2
    If (max = min) Then
        sat = 0
        hue = -1   ' Undefined
    Else
        ' Calculate the saturation
    
        If (lum <= 0.5) Then
            sat = (max - min) / (max + min)
        Else
            sat = (max - min) / (2 - max - min)
        End If
        ' next calculate the hue
        Dim delta As Double
        delta = max - min
        If (RValue = max) Then
            hue = (GValue - BValue) / delta ' result between yellow and magenta
        Else
            If (GValue = max) Then
                hue = 2 + (BValue - RValue) / delta   ' between cyan and yellow
            Else
                If (BValue = max) Then
                    hue = 4 + (RValue - GValue) / delta ' between magenta and cyan
                End If
            End If
        End If
        hue = hue * 60  ' convert to degrees
        If (hue < 0#) Then
            hue = hue + 360  ' make non-negative
        End If
    End If
End Sub
Private Function maximum(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    If (a > b) Then
        If (a > c) Then
            maximum = a
        Else
            maximum = c
        End If
    Else
        If (b > c) Then
            maximum = b
        Else
            maximum = c
        End If
    End If
End Function
Private Function minimum(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    If (a < b) Then
        If (a < c) Then
            minimum = a
        Else
            minimum = c
        End If
    Else
        If (b < c) Then
            minimum = b
        Else
            minimum = c
        End If
    End If
End Function
Private Sub GetRGB(ByVal rgbval As Long, r As Long, g As Long, b As Long)
    r = rgbval Mod 256
    rgbval = Int(rgbval / 256)
    g = rgbval Mod 256
    rgbval = Int(rgbval / 256)
    b = rgbval Mod 256
End Sub
Public Sub GetHLS(ByVal rgbval As Long, hue As Double, lum As Double, sat As Double)
    Dim r As Long, g As Long, b As Long
    Call GetRGB(rgbval, r, g, b)
    Call RGB_To_HLS(r, g, b, hue, lum, sat)
    'If sat = 0 Then
        'sat = 0.00000001
    'End If
End Sub
Private Function GetLum(ByVal rgbval As Long) As Double
    Dim r As Long, g As Long, b As Long, hue As Double, sat As Double, lum As Double
    Call GetRGB(rgbval, r, g, b)
    Call RGB_To_HLS(r, g, b, hue, lum, sat)
    GetLum = lum
End Function
Private Function SetHue(rgbval As Long, hue As Double) As Long
    Dim h As Double, l As Double, s As Double
    Dim r As Long, g As Long, b As Long
    Call GetHLS(rgbval, h, l, s)
    Call HLS_To_RGB(r, g, b, hue, l, s)
    SetHue = RGB(r, g, b)
End Function
Private Function SetSat(rgbval As Long, sat As Double) As Long
    Dim h As Double, l As Double, s As Double
    Dim r As Long, g As Long, b As Long
    Call GetHLS(rgbval, h, l, s)
    Call HLS_To_RGB(r, g, b, h, l, sat)
    SetSat = RGB(r, g, b)
End Function
Private Function SetHueSat(rgbval As Long, hue As Double, sat As Double) As Long
    Dim h As Double, l As Double, s As Double
    Dim r As Long, g As Long, b As Long
    Call GetHLS(rgbval, h, l, s)
    Call HLS_To_RGB(r, g, b, hue, l, sat)
    SetHueSat = RGB(r, g, b)
End Function
Private Function SetHLS(hue As Double, lum As Double, sat As Double) As Long
    Dim r As Long, g As Long, b As Long
    Call HLS_To_RGB(r, g, b, hue, lum, sat)
    SetHLS = RGB(r, g, b)
End Function
Private Sub UserForm_Initialize()
    Set NLSStrMgr = CreateObject("frmAdvancedColorRES.NLSStrMgr") 'Important!! Do NOT move, edit, or remove this line!
    NLSStrMgr.NLSContainer Me 'Important!! Do NOT move, edit, or remove this line!
End Sub
