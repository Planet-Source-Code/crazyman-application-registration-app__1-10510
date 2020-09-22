VERSION 5.00
Object = "{EA2A5653-4D0E-11D3-9DB5-444553540000}#1.0#0"; "FormShaper.ocx"
Begin VB.Form frmShow 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Details"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   Icon            =   "frmShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   1800
      Top             =   2040
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Client Details"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblDetails 
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblDetails 
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowCurrent(strName As String, strAddress As String, strKey As String)
With lblDetails
      .Item(0).Caption = strName
      .Item(1).Caption = strAddress
      
End With
txtKey.Text = strKey
Me.Show
End Sub

Private Sub Form_Load()
      FormShaper1.ShapeIt
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Label1.FontUnderline = False
End Sub

Private Sub Label1_Click()
      Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
            Label1.FontUnderline = True
End Sub

Private Sub lblDetails_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      Label1.FontUnderline = False
End Sub
