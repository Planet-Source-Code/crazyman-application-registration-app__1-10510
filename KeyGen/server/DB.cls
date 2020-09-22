VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private conn As ADODB.Connection
Private rs As ADODB.Recordset
Public Sub Init(strConnection As String)
      Set conn = New ADODB.Connection
      conn.Open strConnection
End Sub
Public Function DeleteLicencee(strName As String, strKey As String) As Boolean
      On Error GoTo e
      Dim strSQL As String
      strSQL = "delete from licences where Name='" & strName & "' and key='" & strKey & "'"
      conn.Execute strSQL
      DeleteLicencee = True
Exit Function
e:
AddError Err.Description, Err.Number, "DeleteLicencee"
DeleteLicencee = False
End Function
Public Function AddLicence(Amount As Integer, Name As String, Address As String) As String
      Dim strKey As String
      Dim strSQL As String
      strKey = CreateNewLicenceKey(CInt(Setting("KeyLength")), CInt(Setting("DivLen")), Setting("Divider"))
      strSQL = "insert into licences (Name,Address,Original_Licences,Remaining_Licences,Reg_Attempts,Key)" & _
            "values ('" & Name & "','" & Address & "'," & Amount & "," & Amount & ",0 ,'" & strKey & "')"
      conn.Execute strSQL
      AddLicence = strKey
End Function
Public Function isValidUnRegister(ByVal Name As String, RegKey As String, ByVal UnlockCode As String, ByRef RecId As String) As Boolean
 Dim strSQL As String
 strSQL = "select * from current where Customer_Name = '" & Name & "' and customer_REgKey = '" & RegKey & "' and Decode_Key='" & UnlockCode & "'"
 Set rs = conn.Execute(strSQL)
 If rs.EOF Then
      isValidUnRegister = False
 Else
      strSQL = "select * from licences where Name = '" & Name & "' and key = '" & RegKey & "'"
      Set rs = conn.Execute(strSQL)

      If CInt(rs!remaining_Licences) >= CInt(rs!Original_Licences) Then
            isValidUnRegister = False
      Else
            RecId = rs!id
            isValidUnRegister = True
      End If
 End If
End Function
Public Function UnRegister(Name As String, RecId As String, Key As String) As Boolean
      Dim strSQL As String
      Dim newAmount As Integer
      strSQL = "delete from current where id =" & RecId
      conn.Execute strSQL
      strSQL = "select Remaining_Licences from licences where name='" & Name & "' and key='" & Key & "'"
      Set rs = conn.Execute(strSQL)
      If Not rs.EOF Then
            newAmount = CInt(Val(rs!remaining_Licences)) + 1
            strSQL = "update licences set remaining_Licences = " & newAmount & " where name='" & Name & "' and key='" & Key & "'"
            conn.Execute strSQL
            UnRegister = True
      Else
            UnRegister = False
            Exit Function
      End If
      
End Function
Public Function isValidRequest(ByVal Key As String, ByVal Name As String, Optional ByRef strRemaining As String) As Boolean
      Dim strSQL As String
      Dim strAttempts As String
      strSQL = "select Remaining_Licences ,reg_Attempts,key from licences where key = '" & Key & "' and Name='" & Name & "'"
      Set rs = conn.Execute(strSQL)
      If Not rs.EOF Then
      strAttempts = rs!Reg_Attempts
      If CInt(rs!remaining_Licences) > 0 Then
                  isValidRequest = True
                  strRemaining = CStr(rs!remaining_Licences)
            Else
                  isValidRequest = False
            End If
            strSQL = "update licences set reg_Attempts=" & CInt(Val(strAttempts)) + 1 & " where key = '" & Key & "' and Name='" & Name & "'"
            conn.Execute strSQL
      Else
            isValidRequest = False
      End If
      
End Function
Public Function UserReg(ByVal strName As String, ByVal strAddress As String, ByVal strKey As String, ByVal Code As String, ByVal intCurrentRemaining As Integer, DecodeKey As Integer) As Boolean
      Dim strSQL As String
'On Error GoTo reg_Err
      strSQL = "update licences set remaining_licences =" & CStr(intCurrentRemaining - 1) & " where key='" & strKey & "'"
      conn.Execute strSQL
      strSQL = "insert into current (Customer_Name,Customer_Address,Customer_RegDate,Customer_RandID,Customer_RegKey,Decode_Key)" & _
            " values ('" & strName & "','" & strAddress & "','" & Format(Now, "dd mmm yyyy hh:nn:ss") & "','" & Code & "','" & strKey & "','" & CStr(DecodeKey) & "')"
      conn.Execute strSQL
      UserReg = True
Exit Function
reg_Err:
AddError Err.Description, Err.Number, "UserReg"
End Function
Public Function CreateNewLicenceKey(Length As Integer, divLen As Integer, Divider As String) As String
      Dim strPart As String
      Dim sRand As Single
      Dim count As Integer
      Dim divCount As Integer
      Dim strKey As String
      Dim blnBadNum As Boolean
      Dim blnOkay As Long
      Dim intRand As Integer
      blnOkay = False
      Do While Not blnOkay
      For count = 1 To Length
            blnBadNum = True
            While blnBadNum
                  Randomize
                  sRand = (122 * Rnd()) + 48
                  intRand = CInt(sRand)
                  If intRand < 48 Or (intRand > 57 And intRand < 65) Or (intRand > 90 And intRand < 97) Or intRand > 122 Then
                        blnBadNum = True
                  Else
                        blnBadNum = False
                  End If
            Wend
            strPart = Chr(CInt(sRand))
            If divCount = divLen And divLen <> 0 Then
                  divCount = 0
                  strKey = strKey & Divider
             End If
                  strKey = strKey & strPart
                  divCount = divCount + 1
          
            
      Next count
      Set rs = conn.Execute("select key from licences where key='" & strKey & "'")
      If rs.EOF Then
            blnOkay = True 'not found
      End If
      Loop
      Set rs = Nothing
      CreateNewLicenceKey = strKey
End Function

Private Sub Class_Initialize()
      frmMain.SetIndicators 2, COLOR_BUSY
End Sub

Private Sub Class_Terminate()
      frmMain.SetIndicators 2, COLOR_OK
End Sub
