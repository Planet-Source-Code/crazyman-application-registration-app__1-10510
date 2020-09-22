VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View"
   ClientHeight    =   5730
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8985
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5280
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":0442
            Key             =   "u"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":059C
            Key             =   "d"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":06F6
            Key             =   "n"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin MSComctlLib.ListView Lv 
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Left"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Initial"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unlock Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Attempted"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "--"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuCopyCode 
         Caption         =   "Copy Code"
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnRightClick As Boolean
Private XITem As ListItem
Function PopulateLicences()
      Dim conn As New ADODB.Connection
      Dim rs As ADODB.Recordset
      Dim strSQL As String
      Dim LItem As ListItem
      Me.MousePointer = 11
      conn.Open Setting("Conn")
      strSQL = "select * from licences"
      Set rs = conn.Execute(strSQL)
      Lv.ListItems.Clear
      While Not rs.EOF
            With Lv.ListItems
                        .Add , , rs!Name
                        .Item(.count).ListSubItems.Add , , rs!Address
                        .Item(.count).ListSubItems.Add , , rs!remaining_Licences
                        .Item(.count).ListSubItems.Add , , rs!Original_Licences
                        .Item(.count).ListSubItems.Add , , rs!Key
                        .Item(.count).ListSubItems.Add , , rs!Reg_Attempts
            End With
            rs.MoveNext
      Wend
      Set rs = Nothing
      Me.MousePointer = 0
End Function



Private Sub cmdRefresh_Click()
      PopulateLicences
End Sub

Private Sub Form_Load()
      Dim objColH As ColumnHeader
      For Each objColH In Lv.ColumnHeaders
             objColH.Icon = "n"
      Next objColH
        PopulateLicences
End Sub

Private Sub Lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      Dim objColH As ColumnHeader
      For Each objColH In Lv.ColumnHeaders
             objColH.Icon = "n"
      Next objColH
      If Len(ColumnHeader.Tag) = 0 Then
            ColumnHeader.Tag = "true"
      End If
      ColumnHeader.Tag = CStr(Not CBool(ColumnHeader.Tag))
      If CBool(ColumnHeader.Tag) = True Then
            ColumnHeader.Icon = "u"
      Else
            ColumnHeader.Icon = "d"
      End If
      Select Case ColumnHeader.Index
           Case 3, 4, 6
                  SortListView Lv, ColumnHeader.Index, ldtNumber, ColumnHeader.Tag
           Case Else
                  SortListView Lv, ColumnHeader.Index, ldtString, ColumnHeader.Tag
           
      End Select
End Sub

Private Sub Lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
If blnRightClick Then
      Set XITem = Item
      PopupMenu mnuMain
      blnRightClick = False
Else
      blnRightClick = False
End If
End Sub

Private Sub Lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button
      Case 2
            blnRightClick = True
      Case Else
            blnRightClick = False
End Select
End Sub

Private Sub mnuCopyCode_Click()
      Clipboard.SetText XITem.ListSubItems(4).Text
End Sub

Private Sub mnuDelete_Click()
      Dim objDB As New DB
      objDB.Init Setting("Conn")
      If objDB.DeleteLicencee(XITem.Text, XITem.ListSubItems(4).Text) Then
          Lv.ListItems.Remove XITem.Index
      End If
      Set objDB = Nothing
End Sub

Private Sub mnuShow_Click()
        frmShow.ShowCurrent XITem.Text, XITem.ListSubItems(1).Text, XITem.ListSubItems(4).Text
End Sub
