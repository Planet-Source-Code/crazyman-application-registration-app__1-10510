VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Key Server"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   -120
      Width           =   225
      Begin VB.Timer tmrResetSpark 
         Interval        =   150
         Left            =   240
         Top             =   240
      End
   End
   Begin KeyServer.GlowBar GlowBar4 
      Height          =   30
      Left            =   2280
      TabIndex        =   27
      Top             =   1200
      Width           =   1095
      _extentx        =   1931
      _extenty        =   53
   End
   Begin KeyServer.GlowBar GlowBar3 
      Height          =   1095
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   30
      _extentx        =   53
      _extenty        =   1931
   End
   Begin KeyServer.GlowBar GlowBar2 
      Height          =   30
      Left            =   2280
      TabIndex        =   25
      Top             =   120
      Width           =   1095
      _extentx        =   1931
      _extenty        =   53
   End
   Begin VB.Timer tmrReset 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   -240
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Left            =   2880
      TabIndex        =   19
      Top             =   2400
      Width           =   495
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   22
         ToolTipText     =   "System Status"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "Data In"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   20
         ToolTipText     =   "Data Out"
         Top             =   360
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   2
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1200
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   1
         Left            =   120
         Shape           =   3  'Circle
         Top             =   780
         Width           =   270
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000004&
         BackStyle       =   1  'Opaque
         Height          =   135
         Index           =   0
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Settings"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   1200
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   23
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "View"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Advanced"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   -240
      Top             =   -240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin KeyServer.GlowBar GlowBar5 
      Height          =   1095
      Left            =   3360
      TabIndex        =   28
      Top             =   120
      Width           =   30
      _extentx        =   53
      _extenty        =   1931
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1920
      Picture         =   "frmMain.frx":0884
      Top             =   1440
      Width           =   480
   End
   Begin VB.Shape spark 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2040
      Y1              =   360
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   2565
      Picture         =   "frmMain.frx":0B8E
      Stretch         =   -1  'True
      ToolTipText     =   "About KeyServer"
      Top             =   375
      Width           =   555
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2280
      TabIndex        =   29
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   18
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Index           =   5
      Left            =   255
      TabIndex        =   17
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblStatus 
      Caption         =   "None"
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblStatus 
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Last Error"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Errors"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Failed"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Succeded"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Attempted "
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu mnumain 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu void 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnAll As Boolean
Dim ind As Integer
Dim erg As Long
Dim blnHidden As Boolean
Dim blnAtFront As Boolean
Dim lngTrayHeight As Long
Private Sub cmdAction_Click(Index As Integer)
      Select Case Index
            Case 0
                  frmRecordOptions.Show
            Case 2
                  frmView.Show
      End Select
End Sub
Public Sub CreateIcon()
Dim Tic As NOTIFYICONDATA

Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd
Tic.uID = 1&
Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE
Tic.hIcon = Picture1.Picture
Tic.szTip = "KeyServer " & AppVer & Chr$(0)
erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
Public Sub ShowMe()
      SetForegroundWindow Me.hwnd
      Me.WindowState = vbNormal
      While (Me.Top >= (Screen.Height - Me.Height) - GetXtra())
            Me.Top = Me.Top - 10
            DoEvents
      Wend
      blnHidden = False
      mnuShow.Enabled = False
      mnuHide.Enabled = True
End Sub
Public Sub HideMe()
                  SetForegroundWindow Me.hwnd
      Me.WindowState = vbNormal
      While (Me.Top <= (Screen.Height + Me.Height))
            Me.Top = Me.Top + 10
            DoEvents
      Wend
      blnHidden = True
      mnuShow.Enabled = True
      mnuHide.Enabled = False
End Sub
Public Sub DeleteIcon()
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = Picture1.hwnd
Tic.uID = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub


Private Sub Command2_Click()
     Animate
End Sub

Private Sub Form_GotFocus()
      blnAtFront = True
End Sub

Private Sub Form_LostFocus()
      blnAtFront = False
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
      If MsgBox("Are you sure you wish to Exit KeyServer?", vbOKCancel + vbQuestion) = vbOK Then
            DeleteIcon
      Else
            Cancel = 1
      End If
End Sub

Private Sub Image1_Click()
      frmAbout.Show
End Sub

Private Sub mnuEnd_Click()
      Dim f As Form
      For Each f In Forms
            If Not f.hwnd = Me.hwnd Then
                  Unload f
            End If
      Next f
      Unload Me
      
End Sub

Private Sub mnuHide_Click()
      Me.Show
      HideMe
End Sub

Private Sub mnuShow_Click()
If blnHidden Then
      ShowMe
Else
      Me.Show
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print x
X = X / Screen.TwipsPerPixelX

Select Case X
Case WM_LBUTTONDOWN
If blnHidden Then
      ShowMe
Else
      Me.Show
End If
Case WM_RBUTTONDOWN
'Caption = "Right Click"
      PopupMenu mnumain
Case WM_MOUSEMOVE
'Caption = "Move"
Case WM_LBUTTONDBLCLK
If blnHidden Then
      ShowMe
Else
      HideMe
End If
End Select
End Sub
Private Sub cmdStart_Click()
      SetServiceState True
End Sub
Private Sub SetServiceState(blnStart As Boolean)
Animate
If blnStart Then
           If StartSocket(sckMain) Then
            Status "Ready on port " & LOCAL_PORT
            SetIndicators 0, COLOR_OK, True
             Glow True
            cmdStart.Enabled = False
            cmdStop.Enabled = True
      Else
            AddError "", -1, "cmdStart", lblStatus(6), "Could not start socket"
            SetIndicators 0, COLOR_BUSY, True
            Glow False
      End If
Else
          If StopSocket(sckMain) Then
            Status "Stopped..."
            SetIndicators 0, COLOR_BUSY, True
            Glow False
            cmdStart.Enabled = True
            cmdStop.Enabled = False
      Else
            Status "Error..."
            SetIndicators 0, COLOR_BUSY, True
            Glow False
      End If
End If
End Sub
Private Sub cmdStop_Click()
  SetServiceState False
End Sub

Private Sub Command1_Click()
      frmSettings.Show vbModal, Me
End Sub


Public Sub Status(strStatus As String)
      lblStatus(6).Caption = strStatus
End Sub
Public Sub SetIndicators(Index As Integer, ByVal color As ColEnum, Optional blnAll As Boolean)
Shape1(Index).BackColor = color
If blnAll Then
      Shape1(0).BackColor = color
      Shape1(1).BackColor = color
      Shape1(2).BackColor = color
End If
Animate
End Sub
Private Sub SetLEDs()
      Shape1(0).Left = (Frame3.Width / 2) - (Shape1(0).Width / 2)
      Shape1(1).Left = (Frame3.Width / 2) - (Shape1(1).Width / 2)
      Shape1(2).Left = (Frame3.Width / 2) - (Shape1(2).Width / 2)

End Sub

Function GetXtra() As Long
Dim hwndTray As Long
Dim rec As RECT
  hwndTray = FindWindow("shell_traywnd", "")
            GetWindowRect hwndTray, rec
            If rec.Top > 0 Then
                  lngTrayHeight = rec.Top * 15
            Else
                  lngTrayHeight = Screen.Height
            End If
GetXtra = Screen.Height - lngTrayHeight
End Function

Private Sub Form_Load()
            Me.Left = Screen.Width - Me.Width
            Me.Top = Screen.Height
            Init
            'GlowBar1.start
         
End Sub
Sub Glow(blnGlow As Boolean)
      If blnGlow Then
            GlowBar2.start
            GlowBar3.start
            GlowBar4.start
            GlowBar5.start
      Else
            GlowBar2.EndGlow
            GlowBar3.EndGlow
            GlowBar4.EndGlow
            GlowBar5.EndGlow
      End If
End Sub

Private Sub Init()
      SetLEDs
      CreateIcon
      SetForegroundWindow Me.hwnd
      If GetSetting(App.EXEName, "Settings", "StartOnLoad", 1) = "1" Then
            SetServiceState True
      End If
      If GetSetting(App.EXEName, "Settings", "ShowOnStart", 1) = "1" Then
            ShowMe
      End If
      GlowBar2.BoxGradient Picture2, 20, 30, 200, 1, 1, 1, True
      'Image1.ZOrder 0
      'BitBlt Me.hDC, GlowBar1.Left, GlowBar1.Top, Picture2.Width, Picture2.Height, Picture2.hDC, Picture2.Width, Picture2, SRCCOPY
End Sub

                  
Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
      Dim strData As String
      SetIndicators 0, COLOR_BUSY
      sckMain.GetData strData
      HandleMessage strData, sckMain.RemoteHostIP
End Sub
Public Sub SetResetAll()
      'SetIndicators 0, COLOR_MED, True
      blnAll = True
      tmrReset.Enabled = True
End Sub
Public Sub SetReset(Index As Integer)
      SetIndicators Index, COLOR_MED
      ind = Index
      blnAll = False
      tmrReset.Enabled = True
End Sub
Public Sub Animate()
      Dim s As Shape
      Load spark(spark.UBound + 1)
      Set s = spark(spark.UBound)
With s
            .Width = spark(0).Width
            .BorderStyle = spark(0).BorderStyle
            .FillStyle = spark(0).FillStyle
            .FillColor = spark(0).FillColor
            .Shape = spark(0).Shape
            .Visible = True
            .Left = Line2.X1 - (.Width / 2) + 14
            tmrAnimate.Enabled = True
End With

End Sub

Private Sub tmrAnimate_Timer()
On Error Resume Next
      Dim X As Integer
      For X = spark.LBound To spark.UBound
            If spark.UBound = 0 Then
                   tmrAnimate.Enabled = False
            End If
            If X <> 0 Then
                  If spark(X).Top > (Line2.Y1 - 34) Then
                        spark(X).Top = spark(X).Top - 13
                  ElseIf spark(X).Left < Line1.X2 Then
                        spark(X).Left = spark(X).Left + 13
                  Else
                        Label7.BackColor = vbRed
                        Unload spark(X)
                        tmrResetSpark = True
                  End If
            Else
                   'tmrAnimate.Enabled = False
            End If
            DoEvents
      Next X
End Sub

Private Sub tmrReset_Timer()
If blnAll Then
      SetIndicators 0, COLOR_OK, True
Else
      SetIndicators ind, COLOR_OK
End If
      tmrReset.Enabled = False
End Sub

Private Sub tmrResetSpark_Timer()
      Label7.BackColor = &HFF0000
      tmrResetSpark.Enabled = False
End Sub
