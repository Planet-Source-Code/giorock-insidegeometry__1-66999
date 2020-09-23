VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmGeometry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   56.356
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   85.99
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*********************
' Written by GioRock *
'*********************
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Pt2() As POINTAPI

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Enum WHERE_IS
    Inside = 0
    Outside = 1
    OnBorder = 2
    OnVertex = 3
    DontExist = 4
End Enum

Private CX As Single
Private CY As Single
Private RADIUS As Single
Private ASPECT As Single

Private PI As Double
Private HALPH_PI As Double

Private Function InsidePolygon(PtApi() As POINTAPI, ByVal X As Single, ByVal Y As Single) As WHERE_IS
Dim i As Integer
Dim theta As Double
Dim alfa As Double
Dim xNew As Double
Dim yNew As Double
Dim angle As Double
Dim nCount As Integer
'*********************************
' Thanks to Italian's Programmer *
'*********************************

    nCount = UBound(PtApi)
    
    If (nCount < 3) Then
        InsidePolygon = DontExist ' polygon don't exist
        Exit Function
    End If

    For i = 0 To nCount
        If (CLng(X) = PtApi(i).X And CLng(Y) = PtApi(i).Y) Then
            InsidePolygon = OnVertex ' point on polygon vertex
            Exit Function
        End If
    Next i

    PtApi(nCount).X = PtApi(0).X '  enforce polygon closure
    PtApi(nCount).Y = PtApi(0).Y '
    
    angle = 0
   
    For i = 0 To nCount - 1
        theta = ATan2(CDbl(PtApi(i).X) - X, CDbl(PtApi(i).Y) - Y)
        xNew = (CDbl(PtApi(i + 1).X) - X) * Cos(theta) + (CDbl(PtApi(i + 1).Y) - Y) * Sin(theta)
        yNew = (CDbl(PtApi(i + 1).Y) - Y) * Cos(theta) - (CDbl(PtApi(i + 1).X) - X) * Sin(theta)
        alfa = ATan2(xNew, yNew)
        If Round(alfa, 1) = Round(PI, 1) Then
            InsidePolygon = OnBorder ' point on polygon side
            Exit Function
        End If
        angle = angle + alfa
    Next i
   
    If (Abs(angle) < 0.0001) Then
        InsidePolygon = Outside ' external point
        Exit Function
    ElseIf (Abs(angle) > PI) Then
        InsidePolygon = Inside  ' internal point
        Exit Function
    End If

End Function

Private Function ATan2(ByVal X As Double, ByVal Y As Double) As Single
' A "C" function in VB
'*********************************
' Thanks to Italian's Programmer *
'*********************************

    If X = 0 Then
        If Y > 0 Then
            ATan2 = HALPH_PI
        ElseIf Y < 0 Then
            ATan2 = -HALPH_PI
        Else
            ATan2 = 0
        End If
    ElseIf Y = 0 Then
        If X < 0 Then
            ATan2 = PI
        Else
            ATan2 = 0
        End If
    Else
        If X < 0 Then
            If Y > 0 Then
                ATan2 = Atn(Y / X) + PI
            Else
                ATan2 = Atn(Y / X) - PI
            End If
        Else
            ATan2 = Atn(Y / X)
        End If
    End If
   
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'*********************
' Created by GioRock *
'*********************
    If KeyCode = vbKeyF8 And Shift = 0 Then
        While ShowCursor(False) >= 0
        Wend
    ElseIf KeyCode = vbKeyF9 And Shift = 0 Then
        While ShowCursor(True) < 0
        Wend
    End If
End Sub

Private Sub Form_Load()
Dim Pt As POINTAPI
    
    PI = 4 * Atn(1)      ' pigreco
    HALPH_PI = PI / 2    ' pigreco mezzi
    
    Show
    DoEvents
    
    Pt.X = PixToMM(CX)
    Pt.Y = PixToMM(CY)
    ClientToScreen hwnd, Pt
    
    SetCursorPos Pt.X, Pt.Y
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim WI As WHERE_IS
Dim sPos As String
Dim sCaption As String
'*********************
' Created by GioRock *
'*********************

    AutoRedraw = False
    DrawWidth = 1
    Refresh
    DrawStyle = vbDot
    
    If X >= 5 And Y >= 5 Then
        Line (X, 0)-(X, ScaleHeight)
        Line (0, Y)-(ScaleWidth, Y)
        sPos = "X:" + Str(Round(X - 5, 2)) + ", Y:" + Str(Round(Y - 5, 2))
        CurrentX = ScaleWidth - (TextWidth(sPos) + 1)
        CurrentY = ScaleHeight - (TextHeight(sPos) + 1)
        Print sPos
    End If
    
    WI = InsideEllipse(X, Y)
    If WI = Inside Then
        If sCaption <> "InSide Ellipse" Then
            sCaption = "InSide Ellipse"
        End If
    ElseIf WI = Outside Then
        If sCaption <> "OutSide Ellipse" Then
            sCaption = "OutSide Ellipse"
        End If
    Else
        If sCaption <> "OnBorder Ellipse" Then
            sCaption = "OnBorder Ellipse"
        End If
    End If
    
    WI = InsidePolygon(Pt2(), PixToMM(X), PixToMM(Y))
    If WI = Inside Then
        If sCaption <> sCaption + " and InSide Polygon" Then
            sCaption = sCaption + " and InSide Polygon"
        End If
    ElseIf WI = Outside Then
        If sCaption <> sCaption + " and OutSide Polygon" Then
            sCaption = sCaption + " and OutSide Polygon"
        End If
    ElseIf WI = OnBorder Then
        If sCaption <> sCaption + " and OnBorder Polygon" Then
            sCaption = sCaption + " and OnBorder Polygon"
        End If
    ElseIf WI = OnVertex Then
        If sCaption <> sCaption + " and OnVertex Polygon" Then
            sCaption = sCaption + " and OnVertex Polygon"
        End If
    Else
        If sCaption <> sCaption + " and Polygon Don't Exist" Then
            sCaption = sCaption + " and Polygon Don't Exist"
        End If
    End If
    
    If Me.Caption <> sCaption Then
        Me.Caption = sCaption
    End If
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        CX = (ScaleWidth / 2) + 3
        CY = (ScaleHeight / 2) + 3
        RADIUS = (ScaleHeight / 3)
        ASPECT = 1.2 ' Try to Change ASPECT: from 0.1 to upper value
        DrawStyle = vbSolid
        AutoRedraw = True
        Cls
        DrawWidth = 1
        DrawMeterGuide
        DrawWidth = PixToMM(1) ' 3.78px = 1mm.
        DrawCircle
        DrawRhomb
        ForeColor = QBColor(2)
        ForeColor = QBColor(0)
        Refresh
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowCursor True
    Cls
    Erase Pt2
    End
    Set Form1 = Nothing
End Sub



Private Function InsideEllipse(ByVal X As Single, ByVal Y As Single) As WHERE_IS
Dim TBase As Single
Dim THeight As Single
Dim THypotenuse As Single
'*********************
' Created by GioRock *
'*********************

    'Calculate triangle base and height
    If ASPECT = 1 Then
        TBase = Abs(CX - X)
        THeight = Abs(CY - Y)
    ElseIf ASPECT > 1 Then
        TBase = Abs(CX - X) * ASPECT
        THeight = Abs(CY - Y)
    Else
        TBase = Abs(CX - X)
        THeight = Abs(CY - Y) / ASPECT
    End If
    
    ' We apply the Pitagora's Theorem to find hypotenuse
    THypotenuse = CSng(Sqr(TBase ^ 2 + THeight ^ 2))
    
    'Verify where is the hypotenuse in ellipse area
     
    'Secant
    If THypotenuse < RADIUS - CSng(1 / PixToMM(1)) Then ' Only (RADIUS - 1) if you use pixel ScaleMode and DrawWidth = 1
        InsideEllipse = Inside
    'Extern
    ElseIf THypotenuse > RADIUS + CSng(1 / PixToMM(1)) Then
        InsideEllipse = Outside
    'Tangent
    Else
        InsideEllipse = OnBorder
    End If
    
End Function

Private Sub DrawMeterGuide()
Dim X As Single
Dim Y As Single
Dim iC As Integer
'*********************
' Created by GioRock *
'*********************

    Line (0, 0)-(5, 5), QBColor(7), BF
    CurrentX = 0.2
    CurrentY = 0.5
    Print "mm"
    
    For X = 0 To ScaleWidth Step 1
        If X Mod 10 = 0 Then
            Line (X + 5, 0)-(X + 5, 5)
            If iC > 0 Then
                CurrentX = (X + 5) - CSng(TextWidth(CStr(iC)) / Len(CStr(iC)) + IIf(Len(CStr(iC)) = 1, 0, 1))
                CurrentY = 5
                Print iC
            End If
            iC = iC + 1
        Else
            Line (X + 5, 0)-(X + 5, IIf(X Mod 5 = 0, 3, 1))
        End If
    Next X
    
    iC = 0
    For Y = 0 To ScaleHeight Step 1
        If Y Mod 10 = 0 Then
            Line (0, Y + 5)-(5, Y + 5)
            If iC > 0 Then
                CurrentX = 5
                CurrentY = (Y + 5) - CSng(TextHeight(CStr(iC)) / 2)
                Print iC
            End If
            iC = iC + 1
        Else
            Line (0, Y + 5)-(IIf(Y Mod 5 = 0, 3, 1), Y + 5)
        End If
    Next Y
        
End Sub

Private Sub DrawCircle()
    Me.Circle (CX, CY), RADIUS, QBColor(12), , , ASPECT
    Me.Circle (CX, CY), 1 / 3.78, QBColor(1), , , ASPECT
End Sub

Private Function PixToMM(ByVal Pix As Single) As Single
    PixToMM = Pix * 3.78
End Function

Private Sub DrawRhomb()
'*********************
' Created by GioRock *
'*********************

    ReDim Preserve Pt2(0 To 4)
    
    Pt2(0).X = PixToMM(CX)
    Pt2(0).Y = PixToMM(CY) - IIf(ASPECT <= 1, PixToMM(RADIUS * ASPECT), PixToMM(RADIUS / ASPECT))
    Pt2(1).X = PixToMM(CX) + IIf(ASPECT <= 1, PixToMM(RADIUS * ASPECT), PixToMM(RADIUS / ASPECT))
    Pt2(1).Y = PixToMM(CY)
    Pt2(2).X = PixToMM(CX)
    Pt2(2).Y = PixToMM(CY) + IIf(ASPECT <= 1, PixToMM(RADIUS * ASPECT), PixToMM(RADIUS / ASPECT))
    Pt2(3).X = PixToMM(CX) - IIf(ASPECT <= 1, PixToMM(RADIUS * ASPECT), PixToMM(RADIUS / ASPECT))
    Pt2(3).Y = PixToMM(CY)
    Pt2(4) = Pt2(0)
    
    Polygon hdc, Pt2(0), 5
    
End Sub




