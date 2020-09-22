Attribute VB_Name = "modMain"
Option Explicit
Public Const LOCAL_PORT As Integer = 12101
Public Const DEFAULT_EMAIL_ADDRESS As String = "ric_fri@yahoo.com"
Public Enum ColEnum
      COLOR_OK = &HFF00&
      COLOR_MED = &HFFFF&
      COLOR_BUSY = &HFF&
End Enum

Public Enum enumEvents
      REG_ATTEMPT = 0
      REG_ACCEPT = 1
      REG_FAIL = 2
      REG_ERROR = 3
End Enum
Private Const REGISTER_MESSAGE As String = "REG"
Private Const REGISTER_OKAY As String = "REGOKAY"
Private Const REGISTER_FAIL As String = "REGFAIL"
Private Const UNREGISTER_MESSAGE As String = "UNREG"
Private Const UNREGISTER_OKAY As String = "UNREGOKAY"
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'Make your own constant, e.g.:
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
  '  Public What As RECT




Public Function Setting(Name As String) As String
      Setting = GetSetting(App.EXEName, "Settings", Name)
End Function
Public Sub IncrementEvent(f As Form, ev As enumEvents, Optional ByVal strLastErr As String)
      f.lblStatus(CInt(ev)).Caption = CStr(CInt(f.lblStatus(CInt(ev)).Caption) + 1)
      If ev = REG_ERROR Then
            f.lblStatus(REG_ERROR).Caption = strLastErr
      End If
End Sub
Public Function StartSocket(sck As Winsock) As Boolean
On Error GoTo sock_err
      With sck
            .Protocol = sckUDPProtocol
            .Bind LOCAL_PORT
      End With
      StartSocket = True
Exit Function
sock_err:
AddError Err.Description, Err.Number, "StartSocket", frmMain.lblStatus(6), "Could not Start Socket"
StartSocket = False
End Function
Public Function StopSocket(sck As Winsock) As Boolean
      On Error GoTo err_Stop_sck
      With sck
            .Close
      End With
StopSocket = True
Exit Function
err_Stop_sck:
AddError Err.Description, Err.Number, "StopSocket"
StopSocket = False
End Function

Public Sub AddError(ByVal strErr As String, ByVal intErrNum As Long, ByVal strCaller As String, Optional objStatus As Label, Optional ByVal strDescription As String)
      Dim intFile As Integer
      Dim strData As String
      intFile = FreeFile
      strData = "[" & Format(Now, "hh:nn:ss") & "] (" & strCaller & ")  " & CStr(intErrNum) & " : " & strErr & "      :" & strDescription & vbCrLf
      Open Setting("LogFile") For Append As #intFile
            Print #intFile, , strData
      Close #intFile
      If Not objStatus Is Nothing Then
            objStatus.Caption = strDescription
      End If
      frmMain.lblStatus(4).Caption = strErr
      frmMain.lblStatus(3) = CInt(Val(frmMain.lblStatus(3))) + 1
End Sub
Public Function AppVer() As String
      AppVer = App.Major & "." & App.Minor & "." & App.Revision
End Function
Public Sub HandleMessage(strMessage As String, Ipaddr As String)
Dim strMessageType As String
Dim strKey As String
Dim strName As String
Dim strAdd As String
Dim strRandId As String
Dim enc As New crypto
Dim intDecode As Integer
Dim objDB As New DB
Dim strRemaining As String
Dim strCode As String
enc.InBuffer = strMessage
enc.Decrypt
strMessage = enc.OutBuffer
strMessageType = TrimString(strMessage)
Select Case strMessageType
      Case REGISTER_MESSAGE
            strKey = TrimString(strMessage)
            strName = TrimString(strMessage)
            strAdd = TrimString(strMessage)
            strRandId = TrimString(strMessage)
            IncrementEvent frmMain, REG_ATTEMPT
            If VerifyRequest(strKey, strName, strRemaining) Then
                  ''okay
                  ''''
                  Randomize
                  intDecode = CInt((1000 * Rnd()) + 400)
                  objDB.Init Setting("Conn")
                  If objDB.UserReg(strName, strAdd, strKey, strRandId, CInt(strRemaining), intDecode) Then
                        IncrementEvent frmMain, REG_ACCEPT
                        SendMessage frmMain.sckMain, Ipaddr, REGISTER_OKAY, strRandId, strKey, intDecode
                        
                  Else
                        IncrementEvent frmMain, REG_FAIL
                        IncrementEvent frmMain, REG_ERROR, "Could not create Record"
                        SendFailMessage frmMain.sckMain, Ipaddr, strRandId, "An Error Occured"
                  End If
            Else
                        IncrementEvent frmMain, REG_FAIL
                        SendFailMessage frmMain.sckMain, Ipaddr, strRandId, "Invalid User"
            End If
      Case UNREGISTER_MESSAGE
            Dim RecId As String
            strKey = TrimString(strMessage)
            strCode = TrimString(strMessage)
            strName = TrimString(strMessage)
            strRandId = TrimString(strMessage)
             objDB.Init Setting("Conn")
             If objDB.isValidUnRegister(strName, strKey, strCode, RecId) Then
                  If objDB.UnRegister(strName, RecId, strKey) Then
                        SendDeRegister frmMain.sckMain, Ipaddr, strRandId
                  Else
                  End If
             Else
             End If
End Select
frmMain.SetResetAll
Set enc = Nothing
Set objDB = Nothing
End Sub
Function SendDeRegister(sck As Winsock, ip As String, strRand As String)
       Dim enc As New crypto
        frmMain.SetIndicators 1, COLOR_BUSY
        enc.InBuffer = UNREGISTER_OKAY & "|" & strRand & "|"
        enc.Encrypt
        sck.RemoteHost = ip
        sck.SendData enc.OutBuffer
      frmMain.SetReset 1
      Set enc = Nothing
End Function
Function SendFailMessage(sck As Winsock, ip As String, strRand As String, Optional ByVal strDescription As String)
      Dim enc As New crypto
      frmMain.SetIndicators 1, COLOR_BUSY
      enc.InBuffer = REGISTER_FAIL & "|" & strRand & "|" & strDescription & "|"
      enc.Encrypt
      sck.RemoteHost = ip
      sck.SendData enc.OutBuffer
      frmMain.SetReset 1
      Set enc = Nothing
End Function
Function SendMessage(sck As Winsock, ip As String, ByVal strMessage As String, strRand As String, strKey As String, Decode As Integer)
      Dim enc As New crypto
      frmMain.SetIndicators 1, COLOR_BUSY
      enc.InBuffer = strMessage & "|" & strRand & "|" & strKey & "|" & CStr(Decode) & "|"
      enc.Encrypt
      sck.RemoteHost = ip
      sck.SendData enc.OutBuffer
      frmMain.SetReset 1
      Set enc = Nothing
End Function
Private Function VerifyRequest(Key As String, Name As String, ByRef strRemaining As String) As Boolean
      Dim objDB As New DB
      objDB.Init Setting("Conn")
      VerifyRequest = objDB.isValidRequest(Key, Name, strRemaining)
      Set objDB = Nothing
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
Public Function GetEmail() As String
      Dim intFile As Integer
      Dim strEmail As String
      intFile = FreeFile
      If Len(Dir(App.Path & "\email.txt")) > 0 Then
            Open App.Path & "\email.txt" For Input As #intFile
                  Line Input #intFile, strEmail
            Close intFile
            GetEmail = strEmail
      Else
            GetEmail = DEFAULT_EMAIL_ADDRESS
      End If
End Function
