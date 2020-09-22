VERSION 5.00
Begin VB.UserControl GlowBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "GlowBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim X(2) As Integer
Dim change As Boolean
Dim max As Boolean
Dim min As Integer, interval As Integer
Dim r1 As Integer, g1 As Integer, b1 As Integer
Dim red As Integer, green As Integer, blue As Integer

Public Sub start()
Timer1.Enabled = True
End Sub
Public Sub EndGlow()
Timer1.Enabled = False
End Sub
Public Sub BoxGradient(OBJ As Object, R%, G%, b%, RStep%, GStep%, BStep%, Direc As Boolean)
    Dim s%, xpos%, ypos%
    OBJ.ScaleMode = 3 'pixel


    If Direc = True Then
        RStep% = -RStep%
        GStep% = -GStep%
        BStep% = -BStep%
    End If


DoBox:
        s% = s% + 1
        If xpos% < Int(OBJ.ScaleWidth / 2) Then xpos% = s%
        If ypos% < Int(OBJ.ScaleHeight / 2) Then ypos% = s%
        OBJ.Line (xpos%, ypos%)-(OBJ.ScaleWidth - xpos%, OBJ.ScaleHeight - ypos%), RGB(R%, G%, b%), B
        R% = R% - RStep%
        If R% < 0 Then R% = 0
        If R% > 255 Then R% = 255
        G% = G% - GStep%
        If G% < 0 Then G% = 0
        If G% > 255 Then G% = 255
        b% = b% - BStep%
        If b% < 0 Then b% = 0
        If b% > 255 Then b% = 255


        If xpos% = Int(OBJ.ScaleWidth / 2) And ypos% = Int(OBJ.ScaleHeight / 2) Then
            Exit Sub
        End If
        GoTo DoBox
    End Sub
Private Sub Timer1_Timer()
Dim Y As Integer
Dim b As Integer
Dim a As Integer
Timer1.Enabled = False
DoEvents
    If change = True Then
        Randomize
        
        r1 = 0
        g1 = 0
        b1 = 0
        
        For Y = LBound(X()) To UBound(X())
            X(Y) = Int((3 * Rnd) + 0)
        Next Y
        For b = LBound(X()) + 1 To UBound(X())
            a = b
            Do
                a = a - 1
                If X(b) = X(a) Then
                    X(b) = Int((3 * Rnd) + 0)
                    a = b
                End If
            Loop Until a = 0
        Next b
        
        If X(0) = 0 Then red = 0
        If X(0) = 1 Then
            red = Int((253 * Rnd) + 0)
            min = red
        End If
        If X(0) = 2 Then red = 252
        
        If X(1) = 0 Then green = 0
        If X(1) = 1 Then
            green = Int((253 * Rnd) + 0)
            min = green
        End If
        If X(1) = 2 Then green = 252
        
        If X(2) = 0 Then blue = 0
        If X(2) = 1 Then
            blue = Int((253 * Rnd) + 0)
            min = blue
        End If
        If X(2) = 2 Then blue = 252
        
        interval = min / 84
        
        change = False
        
    Else
        If max = False Then
            
            If red = 252 Then
                If r1 <= 252 Then
                    r1 = r1 + 3
                Else
                    max = True
                End If
            ElseIf red > 0 Then
                If r1 <= 252 Then r1 = r1 + interval
            End If
            
            If green = 252 Then
                If g1 <= 252 Then
                    g1 = g1 + 3
                Else
                    max = True
                End If
            ElseIf green > 0 Then
                If g1 <= 252 Then g1 = g1 + interval
            End If
            
            If blue = 252 Then
                If b1 <= 252 Then
                    b1 = b1 + 3
                Else
                    max = True
                End If
            ElseIf blue > 0 Then
                If b1 <= 252 Then b1 = b1 + interval
            End If
            
        End If
        DoEvents
        If max = True Then
            If red = 252 Then
                If r1 >= 3 Then
                    r1 = r1 - 3
                Else
                    max = False
                    change = True
                End If
            ElseIf red > 0 Then
                If r1 >= 3 Then r1 = r1 - interval
            End If
            
            If green = 252 Then
                If g1 >= 3 Then
                    g1 = g1 - 3
                Else
                    max = False
                    change = True
                End If
            ElseIf green > 0 Then
                If g1 >= 3 Then g1 = g1 - interval
            End If
            
            If blue = 252 Then
                If b1 >= 3 Then
                    b1 = b1 - 3
                Else
                    max = False
                    change = True
                End If
            ElseIf blue > 0 Then
                If b1 >= 3 Then b1 = b1 - interval
            End If
        End If
        Picture1.BackColor = RGB(r1, g1, b1)
    End If
    DoEvents
   Timer1.Enabled = True
End Sub

Private Sub UserControl_Initialize()
      change = True
End Sub

Private Sub UserControl_Resize()
      Picture1.Width = UserControl.Width
      Picture1.Height = UserControl.Height
End Sub
