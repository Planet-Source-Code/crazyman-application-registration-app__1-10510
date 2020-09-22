VERSION 5.00
Begin VB.Form frmTemplates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Template Editor"
   ClientHeight    =   5940
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTemplate 
      Height          =   5535
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblDetails 
      Caption         =   "Amount"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Code"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Telephone"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "PostCode"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Address3"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Address2"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Address1"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblDetails 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
   End
End
Attribute VB_Name = "frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrReps(0 To 10) As String

Private Sub Form_Load()
      arrReps(0) = "<!--NAME-->"
      arrReps(1) = "<!--ADD1-->"
      arrReps(2) = "<!--ADD2-->"
      arrReps(3) = "<!--ADD3-->"
      arrReps(4) = "<!--POST-->"
      arrReps(5) = "<!--PHONE-->"
      arrReps(6) = "<!--CODE-->"
      arrReps(7) = "<!--AMOUNT-->"
      
End Sub

Private Sub lblDetails_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 
 '     lblDetails(Index).DragMode = 1
      lblDetails(Index).Drag
End Sub

Private Sub txtTemplate_DragDrop(Source As Control, x As Single, y As Single)
   txtTemplate.Text = txtTemplate.Text & arrReps(Source.Index)
End Sub

