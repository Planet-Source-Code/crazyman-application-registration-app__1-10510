VERSION 5.00
Object = "{EA2A5653-4D0E-11D3-9DB5-444553540000}#1.0#0"; "FormShaper.ocx"
Begin VB.Form frmProgg 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   1200
      Top             =   120
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   45
      Width           =   735
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmProgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      Me.Left = (Screen.Width / 2) - (Me.Width / 2)
      Me.Top = (Screen.Height / 2) - (Me.Height / 2) + (frmSplash.Height / 2)
      FormShaper1.ShapeIt
End Sub

Private Sub Timer1_Timer()
      If Shape2(1).Width >= 3255 Then
            frmMain.Show
            Unload frmSplash
            Unload Me
      Else
            Shape2(1).Width = Shape2(1).Width + 13
            Label1.Caption = CStr(CInt(Round((Shape2(1).Width / 3255) * 100, 0))) & "%"
            Label1.Left = (Shape2(1).Width / 2) - (Label1.Width / 2)
            FormShaper1.ShapeIt
            
           ' Shape2.Width = Label2.Width
           SetForegroundWindow Me.hwnd
            SetForegroundWindow frmSplash.hwnd
         '   Shape2.Left = Label2.Left
         '   Shape2.Height = Label2.Height
          '  Shape2.Top = Label2.Top
      End If
End Sub
