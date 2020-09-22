VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Enum reg
  REG_SZ = 1
  REG_BINARY = 3
  REG_DWORD = 4
End Enum

Public Enum HKEY
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_DYN_DATA = &H80000006
End Enum

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKEY As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal HKEY As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKEY As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Byte
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue() As Byte, ByVal cbData As Long) As Long

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
Public Sub EnumKeys(ByVal lPredefinedKey As HKEY, ByVal sKeyName As String, ByRef sKeys() As String)
  
  Dim lRetVal As Long         'result of the API functions
  Dim HKEY As Long            'handle of opened key
  Dim vValue As String * 255
  Dim vClassName As String * 255
  Dim fTime As FILETIME
  Dim i As Long
  i = 0

  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName & Chr$(0), 0, KEY_ALL_ACCESS, HKEY)
    
  Do
    lRetVal = RegEnumKeyEx(HKEY, i, vValue, 255, 0, vClassName, 255, fTime)
    If Not (lRetVal = ERROR_NONE) Then Exit Do
    If (i Mod 100 = 0) Then ReDim Preserve sKeys(i + 100)
    i = i + 1
    sKeys(i) = RemNull(vValue)
  Loop Until (lRetVal = ERROR_NO_MORE_ITEMS)
  
  Call RegCloseKey(HKEY)
  ReDim Preserve sKeys(i)

End Sub
Public Sub EnumValues(ByVal lPredefinedKey As Long, ByVal sKeyName As String, ByRef sKey() As String, ByRef sData() As String)

  Dim lRetVal As Long         'result of the API functions
  Dim HKEY As Long            'handle of opened key
  Dim vValue As String * 255
  Dim vType As Long
  Dim Length As Long
  Dim strData As String * 255 'setting of queried value
  Dim j As Long
  Dim i As Long
  i = 0

  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName & Chr$(0), 0, KEY_ALL_ACCESS, HKEY)
  
  Do
    Length = 255
    lRetVal = RegEnumValue(HKEY, i, vValue, 255, 0, vType, strData, Length)
    If Not (lRetVal = ERROR_NONE) Then Exit Do
    If (i Mod 10 = 0) Then ReDim Preserve sData(i + 10)
    If (i Mod 10 = 0) Then ReDim Preserve sKey(i + 10)
    i = i + 1
    sKey(i) = RemNull(vValue)
    If (vType = REG_SZ) Then
      'For strings.
      sData(i) = RemNull(strData)
    Else
      'For others (DWORDS), I make them into hexadecimal values
      For j = 1 To Length
        sData(i) = sData(i) & CStr(Hex$(Asc(Mid$(strData, j, 1))))
      Next j
    End If
    vValue = ""
    strData = ""
  Loop Until (lRetVal = ERROR_NO_MORE_ITEMS)
  
  Call RegCloseKey(HKEY)
  ReDim Preserve sData(i)
  ReDim Preserve sKey(i)

End Sub
Public Function GetFileTypes(ByRef X() As String) As Integer

  Dim i As Long
  Dim HKEY As Long          'handle of opened key
  Dim lRetVal As Long       'result of the API functions
  Dim sValue As String
  Dim xCount As Integer
  Dim fTime As FILETIME
  Dim vValue As String * 255
  Dim vClassName As String * 255
  
  i = 0
  lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, Chr$(0), 0, KEY_ALL_ACCESS, HKEY)
  
  Do
    lRetVal = RegEnumKeyEx(HKEY, i, vValue, 255, 0, vClassName, 255, fTime)
    If (lRetVal = ERROR_NONE) Then
      sValue = RemNull(vValue)
      If (Left$(sValue, 1) = ".") Then
        If (xCount Mod 100 = 0) Then ReDim Preserve X(xCount + 100)
        xCount = xCount + 1
        X(xCount) = UCase$(Mid$(sValue, 2))
      End If
    End If
    i = i + 1
  Loop Until (lRetVal = ERROR_NO_MORE_ITEMS)
    
  Call RegCloseKey(HKEY)
  ReDim Preserve X(xCount)
  GetFileTypes = xCount
  
End Function
Private Function RemNull(ByVal strString As String) As String
  
  ' Returns a string without any zero terminator.  Typically,
  ' this was a string returned by a Windows API call.
    
  Dim intZeroPos As Integer

  intZeroPos = InStr(strString, Chr$(0))
  If (intZeroPos > 0) Then
    RemNull = Left$(strString, intZeroPos - 1)
  Else
    RemNull = strString
  End If

End Function

Public Function SetKeyValue(lPredefinedKey As HKEY, sKeyName As String, sValueName As String, vValue As Variant, lValueType As reg) As Boolean

  Dim HKEY As Long
  Dim lValue As Long
  Dim sValue As String

  'Open the specified key
  Call RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, HKEY)
  
  Select Case lValueType
  Case REG_SZ
    sValue = vValue & Chr$(0)
    SetKeyValue = (RegSetValueExString(HKEY, sValueName, 0&, lValueType, sValue, Len(sValue)) = 0)
  Case REG_DWORD
    lValue = vValue
    SetKeyValue = (RegSetValueExLong(HKEY, sValueName, 0&, lValueType, lValue, 4) = 0)
  Case REG_BINARY
    SetKeyValue = RegSetValueExBinary(HKEY, sValueName, 0&, lValueType, Hex("" & vValue), Len(Hex("" & vValue)))
  End Select

  'Close the key
  Call RegCloseKey(HKEY)
  
End Function
Function QueryValueEx(ByVal lhKey As HKEY, ByVal szValueName As String, vValue As Variant) As Long

  Dim cch As Long
  Dim lRC As Long
  Dim lType As Long
  Dim lValue As Long
  Dim sValue As String

'  On Local Error GoTo QueryValueExError
    
  'Determine the size and type of data to be read
  If (RegQueryValueExNULL(lhKey, szValueName, 0, lType, 0, cch) <> ERROR_NONE) Then
    Exit Function
  End If
  
  Select Case lType
  Case REG_SZ
    sValue = String(cch, 0)
    lRC = RegQueryValueExString(lhKey, szValueName, 0, lType, sValue, cch)
    If (lRC = ERROR_NONE) Then vValue = Left$(sValue, cch)
  Case REG_DWORD
    lRC = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
    If lRC = ERROR_NONE Then vValue = lValue
  Case REG_BINARY
    Dim lValueType As Long
    Dim byBuffer() As Byte
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    Dim hCurKey As Long

    'Open the key and get number of bytes
    lRegResult = RegQueryValueEx(lhKey, szValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If (lRegResult = 0) Then
      'Initialise buffers and retrieve value
      ReDim byBuffer(lDataBufferSize - 1) As Byte
      lRC = RegQueryValueEx(lhKey, szValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
      lRC = byBuffer(0)
      vValue = lRC
    End If
  Case Else
    lRC = -1
  End Select

QueryValueExExit:
  Exit Function

QueryValueExError:

End Function
Public Function QueryValue(lPredefinedKey As HKEY, sKeyName As String, sValueName As String) As String
    
  Dim HKEY As Long            'handle of opened key
  Dim vValue As Variant       'setting of queried value
  
  If (RegOpenKeyEx(lPredefinedKey, sKeyName & Chr$(0), 0, KEY_ALL_ACCESS, HKEY) = 0) Then
    Call QueryValueEx(HKEY, sValueName, vValue)
    If Not IsEmpty(vValue) Then QueryValue = RemNull(vValue)
  End If
  Call RegCloseKey(HKEY)

End Function
Public Sub DeleteValue(lPredefinedKey As HKEY, sKeyPath As String, sValueName As String)
    
  Dim HKEY As Long            'handle of opened key

  Call RegOpenKeyEx(lPredefinedKey, sKeyPath, 0, KEY_ALL_ACCESS, HKEY)
  Call RegDeleteValue(HKEY, sValueName & Chr$(0))
  Call RegCloseKey(HKEY)

End Sub
Public Sub DeleteKey(lPredefinedKey As HKEY, sKeyPath As String, sKeyName As String)
    
  Dim HKEY As Long            'handle of opened key

  Call RegOpenKeyEx(lPredefinedKey, sKeyPath, 0, KEY_ALL_ACCESS, HKEY)
  Call RegDeleteKey(HKEY, sKeyName & Chr$(0))
  Call RegCloseKey(HKEY)

End Sub
Public Sub CreateNewKey(lPredefinedKey As HKEY, sNewKeyName As String)
    
  Dim hNewKey As Long         'handle to the new key
  Dim lRetVal As Long
  
  Call RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hNewKey, lRetVal)
  Call RegCloseKey(hNewKey)

End Sub


