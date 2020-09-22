VERSION 5.00
Object = "{EA2A5653-4D0E-11D3-9DB5-444553540000}#1.0#0"; "FormShaper.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   1440
      Top             =   840
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting....."
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   2760
      Picture         =   "frmSplash.frx":0000
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmSplash.frx":0442
      Top             =   960
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   480
      Picture         =   "frmSplash.frx":0884
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Tag             =   "s"
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      If App.PrevInstance Then
            MsgBox "You are already running KeyServer"
            Unload Me
            Exit Sub
      End If
      Me.MousePointer = 11
      Me.Left = (Screen.Width / 2) - (Me.Width / 2)
      Me.Top = (Screen.Height / 2) - (Me.Height / 2)
      Label1.Caption = "Version " & AppVer
      FormShaper1.ShapeIt
      'FormShaper1.ShapeIt "u"
      frmProgg.Show
      SetForegroundWindow Me.hwnd
End Sub


Sub Collapse()
Do
      If Me.Width > 5 Then
             Me.Width = Me.Width - 5
      ElseIf Me.Height > 5 Then
              Me.Height = Me.Height - 5
              Me.Show
      Else
            Exit Do
      End If
Loop
frmMain.Show
Unload Me
End Sub
