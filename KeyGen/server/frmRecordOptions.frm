VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecordOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Details"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmRecordOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   4455
      Begin VB.CommandButton cmdCopy 
         Caption         =   "C"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cboTemplate 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCustKey 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Print in Template"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Load Template"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdTemp 
         Caption         =   "Edit Template"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4455
      Begin MSComDlg.CommonDialog cd 
         Left            =   3960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtDetails 
         Height          =   285
         Index           =   5
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Customer Name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Customer Address"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Amount"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmRecordOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAdd_Click()
      Dim strStatus As String
      Dim objDB As New DB
      If Len(txtDetails(0).Text) < 1 Then
            MsgBox "You must enter a name"
                  Exit Sub
      ElseIf Len(txtDetails(5).Text) < 1 Then
            MsgBox "You must enter an amount"
            Exit Sub
      End If
      Me.MousePointer = 11
      objDB.Init Setting("Conn")
      strStatus = objDB.AddLicence(CInt(Val(txtDetails(5).Text)), txtDetails(0).Text, txtDetails(1).Text & txtDetails(2).Text & txtDetails(3).Text)
      txtCustKey.Text = strStatus
      Expand True
      cmdAdd.Enabled = False
      Me.MousePointer = 0
End Sub
Private Sub Expand(blnExpanded As Boolean)
      If blnExpanded Then
            Me.Height = 4815
            txtCustKey.SelStart = 1
            txtCustKey.SelLength = Len(txtCustKey.Text)
           
      Else
            Me.Height = 3105
      End If
End Sub

Private Sub Reset()
      Dim ctlText As Control
      For Each ctlText In Me.Controls
            If TypeOf ctlText Is TextBox Then
                  ctlText.Text = ""
            End If
            Next ctlText
      Expand False
      cmdAdd.Enabled = True
End Sub

Private Sub cmdClose_Click()
      Reset
      Unload Me
End Sub

Private Sub cmdCopy_Click()
      Clipboard.SetText txtCustKey.Text
End Sub

Private Sub cmdReset_Click()
      Reset
End Sub


Private Sub cmdTemp_Click(Index As Integer)
Select Case Index
      Case 0
                 'print
      Case 1
            frmTemplates.Show vbModal
End Select
End Sub

Private Sub Form_Load()
      Expand False
End Sub

Private Sub txtDetails_Change(Index As Integer)
Select Case Index
      Case 5
            If Not IsNumeric(txtDetails(5).Text) Then
                  txtDetails(5).Text = ""
            End If
End Select
End Sub
