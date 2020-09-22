VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "unregister"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "reset"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "register"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      Dim WithEvents a As AppRegister.KeyGen
Attribute a.VB_VarHelpID = -1


Private Sub a_UserDeRegistration(Success As Boolean, strDesc As String)
      MsgBox "The request success was " & Success
      MsgBox "You are " & IIf(a.IsRegistered(Val(Text3)), "Still ", "No Longer") & " Registred"
      ''test any time using object.IsRegistred(KeyCode)
End Sub

Private Sub a_UserRegistration(Registered As Boolean, strDesc As String, KeyCode As Integer)
      Text3 = KeyCode
      MsgBox "The request success was " & Registered
      MsgBox "You are " & IIf(a.IsRegistered(Val(Text3)), "Now ", "Not ") & " Registred"
      ''you must save key code somewhere, use this to check any where in your app if the user is registred
      '' you can save anywhere as the code its self is useless it just needs it as a part of its decrytption
      'you also need it to de-register
      'a.CloseSocket
     ' Set a = Nothing
End Sub

Private Sub Command1_Click()
Set a = New AppRegister.KeyGen
      Dim b As Boolean
     a.Address = "new york" '//this does not need to match
     a.Name = Text2.Text '//this must exactly math the server entry/caps sensitive
     a.ServerIP = "127.0.0.1" 'servers ip
     a.Code = Text1 'this must exactly match the regestration key on the server
     a.Init 'start the socket..etc..
     a.SendRegistration '''request the regestration
    
End Sub

Private Sub Command2_Click()
      a.CloseSocket 'reset the socket
End Sub

Private Sub Command3_Click()
Set a = New AppRegister.KeyGen
     a.ServerIP = "127.0.0.1"
'          a.init
      a.DeRegister Text1, Text2, Text3 ' send exactly Registration Key,Name,KeyCode that you saved earlier
End Sub

