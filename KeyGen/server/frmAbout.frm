VERSION 5.00
Object = "{EA2A5653-4D0E-11D3-9DB5-444553540000}#1.0#0"; "FormShaper.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   1080
      Top             =   1200
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Interval        =   45
      Left            =   1320
      Top             =   360
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4215
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   4200
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Send FeedBack"
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4200
      Picture         =   "frmAbout.frx":014A
      Stretch         =   -1  'True
      ToolTipText     =   "Send FeedBack"
      Top             =   1920
      Width           =   240
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "frmAbout.frx":058C
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   0
      Left            =   4080
      Picture         =   "frmAbout.frx":09CE
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   840
      Picture         =   "frmAbout.frx":0E10
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strAll As String

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.FontUnderline = False
    TogglePics False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontUnderline = False
TogglePics False
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontUnderline = False
TogglePics False
End Sub
Private Sub Image3_Click()
On Error Resume Next
Dim strEmail As String
strEmail = "mailto:" & GetEmail
ShellExecute Me.hwnd, "open", strEmail, "", "", 0
TogglePics False
End Sub
Private Sub Image4_Click()
On Error Resume Next
Dim strEmail As String
strEmail = "mailto:" & GetEmail
ShellExecute Me.hwnd, "open", strEmail, "", "", 0
TogglePics False
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TogglePics True
End Sub

Private Sub Label1_Click()
      Unload Me
End Sub
Sub TogglePics(blnOver As Boolean)
      If blnOver Then
            Image4.Visible = True
            Image3.Visible = False
      Else
            Image3.Visible = True
            Image4.Visible = False
      End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Label1.FontUnderline = True
End Sub

Private Sub Text1_Click()
      ShellExecute Me.hwnd, "mailto", "", "ric_fri@yahoo.com", "", 0
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Label1.FontUnderline = False
         TogglePics False
End Sub

Private Sub Timer1_Timer()
      IncText
End Sub

Private Sub Form_Load()
      Me.Left = (Screen.Width / 2) - (Me.Width / 2)
      Me.Top = (Screen.Height / 2) - (Me.Height / 2)
      strAll = "Original concept and design by Richard Friend" & vbCrLf
      strAll = strAll & "This product is not officially supported, however if you do have any comments or suggestions"
      strAll = strAll & " please send them to " & GetEmail & vbCrLf
      strAll = strAll & "Id like to give credit to some people whos code i have used but unfortunatley i do not"
      strAll = strAll & " have their names, if you recognise your code please drop me a line and i will add your name to the credits."
      FormShaper1.ShapeIt
End Sub


Sub IncText()
      If Len(Text1.Text) < Len(strAll) Then
            Text1.Text = Mid(strAll, 1, Len(Text1.Text) + 1)
            DoEvents
            Text1.SelStart = Len(Text1)
            DoEvents
      Else
            Timer1.Enabled = False
      End If
End Sub
