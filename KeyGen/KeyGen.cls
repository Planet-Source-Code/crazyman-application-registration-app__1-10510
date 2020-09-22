VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function apiSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private WithEvents sck As Winsock
Attribute sck.VB_VarHelpID = -1
Private m_ServIP As String
Private m_RegCode As String
Private m_Name As String
Private m_Address As String
Private m_RandomID As Long
Private Const REGISTER_MESSAGE As String = "REG"
Private Const REGISTER_OKAY As String = "REGOKAY"
Private Const REGISTER_FAIL As String = "REGFAIL"
Private Const UNREGISTER_MESSAGE As String = "UNREG"
Private Const UNREGISTER_OKAY As String = "UNREGOKAY"

Private m_strPath As String
Public Event UserRegistration(Registered As Boolean, strDesc As String, KeyCode As Integer)
Public Event UserDeRegistration(Success As Boolean, strDesc As String)
Public Function SendRegistration()
With sck
      .SendData ConstructRegMessage(m_RegCode)
End With

End Function
Public Sub CloseSocket()
      sck.Close
      Set sck = Nothing
      
End Sub
Public Function DeRegister(strKey As String, strName As String, strUnlockCode As String) As Boolean
      Dim enc As New e
      Dim strtemp As String
      Randomize
      m_RandomID = (1000000 * Rnd() + 1)
      strtemp = UNREGISTER_MESSAGE & "|" & strKey & "|" & strUnlockCode & "|" & strName & "|" & m_RandomID & "|"
      enc.InBuffer = strtemp
      enc.Encrypt
      sck.SendData enc.OutBuffer
      Set enc = Nothing
End Function

Public Function Init() As Boolean
With sck
      .Protocol = sckUDPProtocol
      .RemoteHost = m_ServIP
      .RemotePort = 12101
      .LocalPort = 12000
      .Bind .LocalPort
End With
End Function
Private Function ConstructRegMessage(strKey) As String
Dim strtemp As String
Dim enc As New e
Randomize
m_RandomID = (1000000 * Rnd() + 1)
strtemp = REGISTER_MESSAGE & "|" & strKey & "|" & m_Name & "|" & m_Address & "|" & m_RandomID & "|"
enc.InBuffer = strtemp
enc.Encrypt
ConstructRegMessage = enc.OutBuffer
Set enc = Nothing
End Function
Private Function RegisterClient(key As Integer) As Boolean
On Error GoTo e
      Dim enc As New e
      Dim r As New registry
      Dim sNum As String
      Dim strKeyName As String
      Dim strValName As String
      Dim strVal As String
      Dim intFile As Integer
      Dim sFirst As Single
      Dim sSecond As String
      Dim strFirst As String
      Dim intcount As Integer
      Dim strToPrint As String
      intFile = FreeFile
      sNum = SerialNumber(Left(App.Path, 3))
      Debug.Print SerialNumber(Left(App.Path, 3))
      enc.InBuffer = CStr(sNum)
      enc.Encrypt
      strKeyName = enc.OutBuffer
      r.CreateNewKey HKEY_CURRENT_USER, strKeyName
      enc.InBuffer = CStr(sNum ^ 4)
      enc.Encrypt
      strValName = enc.OutBuffer
      enc.InBuffer = CStr(sNum ^ 6)
      enc.Encrypt
      strVal = enc.OutBuffer
      r.SetKeyValue HKEY_CURRENT_USER, strKeyName, strValName, strVal, REG_SZ
      Randomize
      While sFirst < 0.1
      sFirst = (Rnd())
      Wend
      Debug.Print sFirst
      sSecond = CStr(sNum - sFirst)
      strFirst = sFirst
      If Len(Dir(m_strPath & "\key.txt")) > 0 Then
            SetAttr m_strPath & "\key.txt", vbNormal
            Kill m_strPath & "\key.txt"
      End If
    
                
                Do While Len(CStr(strFirst)) > 0
                  Randomize
                  If intcount = key Then
                        If Left(CStr(strFirst), 1) = "." Then
                              strToPrint = strToPrint & "."
                        Else
                              strVal = CStr(CInt(24 * Rnd()))
                              strVal = strVal & CStr(Left(strFirst, 1))
                              strVal = Chr(strVal)
                              'Debug.Print strVal
                              'Debug.Print Asc(strVal)
                              strToPrint = strToPrint & strVal
                              
                        End If
                        strFirst = Mid(strFirst, 2)
                        intcount = 0
                  Else
                        strToPrint = strToPrint & Chr(CInt((Rnd(255) * 100)))
                  End If
            intcount = intcount + 1
            Loop
        Open m_strPath & "\key.txt" For Binary As intFile
             Put #intFile, , strToPrint
      Close intFile
      intFile = FreeFile
      enc.InBuffer = sSecond
      enc.Encrypt
      If Len(Dir(m_strPath & "\key2.txt")) > 0 Then
            Kill m_strPath & "\key2.txt"
      End If
      Open m_strPath & "\key2.txt" For Append As intFile
                 Print #intFile, , enc.OutBuffer
      Close intFile

RegisterClient = True
Set enc = Nothing
Exit Function
e:
Resume
End Function
Public Function IsRegistered(KeyCode As Integer) As Boolean

      If GetFileInfo(KeyCode) = SerialNumber(Left(App.Path, 3)) And CheckRegKeys Then
            IsRegistered = True
      Else
            IsRegistered = False
      End If

End Function
Private Function GetFileInfo(key As Integer) As String
Dim enc As New e
On Error GoTo e
      Dim intFile As Integer
      Dim strtemp As String
      Dim strData As String
      Dim strTest As String
      Dim intcount As Integer
      Dim strChar As String
      Dim intCode As Integer
      Dim strNum As String
      Dim b() As Byte
      Dim X As Long
      intFile = FreeFile
      If Len(Dir(m_strPath & "\key.txt")) < 1 Then Exit Function
      If Len(Dir(m_strPath & "\key2.txt")) < 1 Then Exit Function
      Open m_strPath & "\key.txt" For Binary As intFile
            Do While Not EOF(intFile)
                   b = InputB(LOF(intFile), 1)
                   'Input #intFile, strTemp
                  'strData = strData & strTemp
            Loop
      Close intFile
      For X = LBound(b) To UBound(b)
            strData = strData & Chr(b(X))
      Next X
      While Len(strData) > 0
            If intcount = key Then
                  strChar = Left(strData, 1)
                  'Debug.Print strChar
                  intCode = Asc(strChar)
                  'Debug.Print intCode
                  strNum = Mid(CStr(intCode), Len(CStr(intCode)), 1)
                  'Debug.Print strNum
                  If Len(strTest) = 1 Then strNum = "."
                  strTest = strTest & strNum
                  intcount = 0
            End If
            strData = Mid(strData, 2)
            intcount = intcount + 1
      Wend
      intFile = FreeFile
      strData = ""
    
      Open m_strPath & "\key2.txt" For Input As intFile
            While Not EOF(intFile)
                  Input #intFile, strtemp
                  strData = strData & strtemp
            Wend
      Close intFile
      enc.InBuffer = strData
      enc.Decrypt
      GetFileInfo = CStr(CLng((enc.OutBuffer) + CSng(strTest)))
      Set enc = Nothing
Exit Function
e:
Stop
Debug.Print Err.Description
Resume Next
End Function
Private Function CheckRegKeys() As Boolean
      Dim r As New registry
      Dim e As New e
      Dim strNum As String
      Dim strNum2 As String
      Dim strNum3 As String
      Dim strNum4 As String
      strNum = SerialNumber(Left(App.Path, 3))
      e.InBuffer = strNum
      e.Encrypt
      strNum2 = e.OutBuffer
      e.InBuffer = CStr(CLng(strNum) ^ 4)
      e.Encrypt
      strNum3 = e.OutBuffer
   '   e.InBuffe =
'      e.Encryptr
      strNum4 = CStr(CLng(strNum) ^ 6)
      e.InBuffer = r.QueryValue(HKEY_CURRENT_USER, strNum2, strNum3)
      e.Decrypt
      CheckRegKeys = (e.OutBuffer = strNum4)
      Set e = Nothing
      Set r = Nothing
End Function

Private Function DestroyReg()
 Dim str As String
 Dim r As New registry
 Dim sNum As String
 Dim strKeyName As String
 str = m_strPath & "\key"
 If Len(Dir(str & ".txt")) > 0 Then
      Kill str & ".txt"
End If
If Len(Dir(str & "2.txt")) > 0 Then
      Kill str & "2.txt"
End If
Dim enc As New e
sNum = SerialNumber(Left(App.Path, 3))
enc.InBuffer = CStr(sNum)
enc.Encrypt
strKeyName = enc.OutBuffer
r.DeleteKey HKEY_CURRENT_USER, "", strKeyName
Set r = Nothing
Set enc = Nothing
End Function


Private Function SerialNumber(ByVal strRoot As String) As String
'---------------------------------------------------------------------------
' FUNCTION: SerialNumber
'
' Returns the serial number of a drive. It returns the number exactly the
' same as DOS does (hexadecimal value e.g. : 1104-224E)
'
' IN:   strRoot      - String containing the root of a drive (e.g. "A:\").
'
' OUT:  SerialNumber - String containing the serial number.
'
' If the function fails (because the drive wasn't ready or something), the
' function returns "0000-0000" as the serial number.
'
'---------------------------------------------------------------------------
'
Dim VolLabel As String
Dim VolSize As Long
Dim SerNum As Long
Dim MaxLen As Long
Dim Flags As Long
Dim Name As String
Dim NameSize As Long
Dim Check As String

If apiSerialNumber(strRoot, VolLabel, VolSize, SerNum, MaxLen, Flags, Name, NameSize) Then
' This function returns a lot more, but I can get that information via another function.

    'Check = Format(Hex(SerNum), "00000000")
    ' Make sure that the length = 8. So convert "123456" to "00123456"
     
    'SerialNumber = Left$(Check, 4) + "-" + Right$(Check, 4)
    ' Split the number in two parts of four and add a "-" between them.
SerialNumber = SerNum
Else
    ' Return "0000-0000" is the function fails.
    SerialNumber = "0000-0000"

End If

End Function



Private Function TrimString(ByRef strIn As String) As String
      Dim intPos As Integer
      intPos = InStr(strIn, "|")
      If intPos > 0 Then
            TrimString = Left(strIn, intPos - 1)
            strIn = Mid(strIn, intPos + 1)
      Else
            TrimString = ""
      End If
End Function

Private Sub Class_Initialize()
      Dim s As String
      Dim X As Integer
      Set sck = frm.sckSend
'      s = "w"
'      Do While s <> ""
'      x = x + 1
'            s = Environ(x)
'            Debug.Print s
'
'      Loop
      m_strPath = Environ("TEMP")
End Sub

Private Sub sck_DataArrival(ByVal bytesTotal As Long)
      Dim strAll As String
      Dim enc As New e
      sck.GetData strAll
      enc.InBuffer = strAll
      enc.Decrypt
      strAll = enc.OutBuffer

      
      HandleResponse strAll
      Set enc = Nothing
End Sub
Private Sub HandleResponse(strAll As String)
      Dim strMessageType As String
      Dim strRandCode As String
      Dim strKey As String
      Dim strReg As String
      strMessageType = TrimString(strAll)
      strRandCode = TrimString(strAll)
      If CLng(strRandCode) = m_RandomID Then
            Select Case strMessageType
                  Case REGISTER_OKAY
                        strReg = TrimString(strAll)
                        strKey = TrimString(strAll)
                        If RegisterClient(CInt(strKey)) = True Then
                              RaiseEvent UserRegistration(True, "Okay", CInt(strKey))
                        Else
                              RaiseEvent UserRegistration(False, "Failed to write details to cache.", CInt(strKey))
                        End If
                  Case REGISTER_FAIL
                        RaiseEvent UserRegistration(False, TrimString(strAll), CInt(Val(strKey)))
                  Case UNREGISTER_OKAY
                        DestroyReg
                        RaiseEvent UserDeRegistration(True, "")
            End Select
      Else
            RaiseEvent UserRegistration(False, "Spoofed Registration Message", CInt(strKey))
      End If
End Sub
Public Property Get ServerIP() As String
      ServerIP = m_ServIP
End Property

Public Property Let ServerIP(ByVal strNewValue As String)
      m_ServIP = strNewValue
End Property

Public Property Get Code() As String
      Code = m_RegCode
End Property

Public Property Let Code(ByVal strCode As String)
      m_RegCode = strCode
End Property

Public Property Get Address() As String
      Address = m_Address
End Property

Public Property Let Address(ByVal strNewValue As String)
      m_Address = strNewValue
End Property
Public Property Get Name() As String
      Name = m_Name
End Property

Public Property Let Name(ByVal strNewValue As String)
      m_Name = strNewValue
End Property

