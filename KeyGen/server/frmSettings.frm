VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "?"
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkSettings 
      Caption         =   "Show on startup"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Tag             =   "ShowOnStart"
      Top             =   840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkSettings 
      Caption         =   "Start on load"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   18
      Tag             =   "StartOnLoad"
      Top             =   360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Okay"
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Tag             =   "LogFile"
      Top             =   2880
      Width           =   3855
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Tag             =   "Conn"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Tag             =   "Divider"
      Text            =   "-"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Tag             =   "DivLen"
      Text            =   "3"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Tag             =   "KeyLength"
      Text            =   "12"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Log File Path"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Connection String"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Divider Character"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Divider Frequency"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Key Length"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Default(0 To 10) As String
Private blnLoaded As Boolean

Private Sub chkSettings_Click(Index As Integer)
        If blnLoaded Then
             cmdAction(1).Enabled = True
        End If
End Sub

Private Sub cmdAction_Click(Index As Integer)
      Select Case Index
            Case 0
                  Unload Me
            Case 1
                  Save
                  cmdAction(1).Enabled = False
            Case 2
                  Save
                  Unload Me
      End Select
End Sub

Private Sub cmdBrowse_Click()
      CommonDialog1.InitDir = App.Path
      CommonDialog1.ShowOpen
      If CommonDialog1.FileName <> "" Then
            txtSettings(4).Text = CommonDialog1.FileName
      End If
End Sub

Private Sub cmdDefault_Click(Index As Integer)
      txtSettings(Index).Text = Default(Index)
End Sub

Private Sub Command1_Click()
      Save
End Sub
Private Sub Save()
Dim X As Integer
Me.MousePointer = 11
For X = txtSettings.LBound To txtSettings.UBound
      SaveSetting App.EXEName, "Settings", txtSettings(X).Tag, txtSettings(X).Text
Next X
For X = chkSettings.LBound To chkSettings.UBound
      SaveSetting App.EXEName, "Settings", chkSettings(X).Tag, chkSettings(X).Value
Next X
Me.MousePointer = 0
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
      Dim X As Integer
      blnLoaded = False
      Default(0) = "12"
      Default(1) = "3"
      Default(2) = "-"
      Default(3) = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & App.Path & "\keys.mdb"
      Default(4) = App.Path & "\ErrLog" & ".log"

For X = txtSettings.LBound To txtSettings.UBound
      txtSettings(X).Text = GetSetting(App.EXEName, "Settings", txtSettings(X).Tag, Default(X))
Next X
For X = chkSettings.LBound To chkSettings.UBound
      chkSettings(X).Value = GetSetting(App.EXEName, "Settings", chkSettings(X).Tag, 1)
Next X
blnLoaded = True
End Sub

Private Sub txtSettings_Change(Index As Integer)
        If blnLoaded Then
             cmdAction(1).Enabled = True
        End If
End Sub
