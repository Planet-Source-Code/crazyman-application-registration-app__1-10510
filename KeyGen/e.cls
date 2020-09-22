VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal HKEY As Long, ByVal dwFlags As Long, phHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal HKEY As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal HKEY As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal HKEY As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, pdwDataLen As Long) As Long
Private Declare Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, phKey As Long) As Long
Private Declare Function CryptGetProvParam Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetUserKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwKeySpec As Long, phUserKey As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptSignHashA Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwKeySpec As Long, ByVal sDescription As String, ByVal dwFlags As Long, ByVal pbSignature As String, pdwSigLen As Long) As Long
Private Declare Function CryptVerifySignatureA Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbSignature As String, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long


Private Const MAX_PATH = 260
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  
Private Declare Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
     
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8

Private Const CRYPT_MACHINE_KEYSET = &H20
Private Const CRYPT_MACHINE_NEWKEYSET = CRYPT_MACHINE_KEYSET Or CRYPT_NEWKEYSET

Private Const PP_CONTAINER = 6
Private Const AT_KEYEXCHANGE = 1
Private Const AT_SIGNATURE = 2

Private Const SIMPLEBLOB = 1

Private Const ALG_CLASS_DATA_ENCRYPT = 24576
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_BLOCK = 1536
Private Const ALG_TYPE_STREAM = 2048
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_MD5 = 3
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC2 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK) Or ALG_SID_RC2)
Private Const CALG_RC4 = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)

Private Const NTE_NO_KEY As Long = -2146893811  '0x8009000DL
Private Const NTE_BAD_SIGNATURE As Long = -2146893818
Private Const CFB_BUSY = 0
Private Const CFB_READY = 1
Private Const CFB_VALID = 2

Private Const ENCRYPT_ALGORITHM = CALG_RC4
Private Const ENCRYPT_BLOCK_SIZE = 1
Private Const CRYPT_EXPORTABLE = 1

Private sInBuffer As String
Private sOutBuffer As String
Private sPassword As String
Private sSignature As String
Private lStatus As Long
Private sLastError As String
Private lKEYSET As Long



Public Property Get InBuffer() As String
  InBuffer = sInBuffer
End Property

Public Property Let InBuffer(vNewValue As String)
  sInBuffer = vNewValue
End Property

Public Property Get LastError() As String
  LastError = sLastError
End Property

Public Property Get OutBuffer() As String
  OutBuffer = sOutBuffer
End Property

Public Property Get Signature() As String
  Signature = sSignature
End Property

Public Property Let Signature(vNewValue As String)
  sSignature = vNewValue
End Property

Public Sub Sign()
  'Create a signature for Inbuffer and place in Signature
   
  Dim sContainer As String
  Dim sDescription As String
  Dim sProvider As String
  Dim lHCryptprov As Long
  Dim lHHash As Long
  Dim lResult As Long
  Dim lSignatureLen As Long

On Error GoTo ErrSign
     
  'switch Status property
  lStatus = CFB_BUSY
   
  'init Signature property
  sSignature = ""
   
  sLastError = ""
   
  'Get handle to the default provider.
  sContainer = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
   
  lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)
  If Not (lResult = 0) Then
    sLastError = "Error " & CStr(GetLastError) & " during CryptAcquireContext!"
    GoTo ReleaseHandles:
  End If
   
  'Create a hash object.
  lResult = CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)
  If Not (lResult = 0) Then
    sLastError = "Error " & CStr(GetLastError) & " during CryptCreateHash!"
    GoTo ReleaseHandles:
  End If
   
  lResult = CryptHashData(lHHash, sInBuffer, Len(sInBuffer), 0)
  If Not (lResult = 0) Then
    sLastError = "Error " & CStr(GetLastError) & " during CryptHashData!"
    GoTo ReleaseHandles:
  End If
   
  'Sign hash object.
  'Determine size of signature.
  sDescription = vbNullChar
  lResult = CryptSignHashA(lHHash, AT_SIGNATURE, sDescription, 0, sSignature, lSignatureLen)
   
  sSignature = String(lSignatureLen, vbNullChar)
   
  'Sign hash object (with signature key).
  lResult = CryptSignHashA(lHHash, AT_SIGNATURE, sDescription, 0, sSignature, lSignatureLen)
  If Not (lResult = 0) Then
    sLastError = "Error " & CStr(GetLastError()) & " during CryptSignHash"
    GoTo ReleaseHandles:
  End If
   
ReleaseHandles:
  'Destroy hash object.
  If lHHash Then lResult = CryptDestroyHash(lHHash)
  'Release provider handle.
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
   
  'switch Status property
  lStatus = CFB_READY
   
Exit Sub
   
ErrSign:
  sLastError = "ErrSign " & Error$
  GoTo ReleaseHandles
End Sub

Public Sub Validate()
  'Validate InBuffer with Signature and assign Status with result
  Dim bValid As Boolean
  Dim sContainer As String
  Dim sDescription As String
  Dim sProvider As String
  Dim lDataLen As Long
  Dim lDataPoint As Long
  Dim lHCryptprov As Long
  Dim lHHash As Long
  Dim lResult As Long
  Dim lSignatureLen As Long
  Dim lHCryptKey As Long

  ReDim aByteData(0) As Byte
   
On Error GoTo ErrValidate

  'switch Status property
  lStatus = CFB_BUSY
   
  sLastError = ""
   
  'init internal valid flag
  bValid = True
   
  'Get handle to the default provider.
  sContainer = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
  lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)
  If Not (lResult = 0) Then
    bValid = False
    sLastError = "Error " & CStr(GetLastError) & " during CryptAcquireContext!"
    GoTo ReleaseHandles:
  End If
   
  'Create a hash object.
  lResult = CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)
  If Not (lResult = 0) Then
    bValid = False
    sLastError = "Error " & CStr(GetLastError) & " during CryptCreateHash!"
    GoTo ReleaseHandles:
  End If
   
  'Add data to hash object.
  lResult = CryptHashData(lHHash, sInBuffer, Len(sInBuffer), 0)
  If Not (lResult = 0) Then
    bValid = False
    sLastError = "Error " & CStr(GetLastError) & " during CryptHashData!"
    GoTo ReleaseHandles:
  End If
   
  'Determine size of signature.
  'sDescription = vbNullChar
  'lResult = CryptSignHashA(lHHash, AT_SIGNATURE, sDescription, 0, 0, lSignatureLen)
   
  'Get handle to signature key.
  lResult = CryptGetUserKey(lHCryptprov, AT_SIGNATURE, lHCryptKey)
  If Not (lResult = 0) Then
    bValid = False
    sLastError = "Error " & CStr(GetLastError) & " during CryptGetUserKey!"
    GoTo ReleaseHandles:
  End If
   
  lSignatureLen = Len(sSignature)
   
  'Verify signature.
  lResult = CryptVerifySignatureA(lHHash, sSignature, lSignatureLen, lHCryptKey, sDescription, 0)
  If Not (lResult = 0) Then
    If GetLastError = NTE_BAD_SIGNATURE Then
      bValid = False
      GoTo ReleaseHandles:
    Else
      bValid = False
      sLastError = "Error " & CStr(GetLastError) & " during CryptVerifySignature!"
      GoTo ReleaseHandles:
    End If
  End If

ReleaseHandles:
  'Release signature key.
  If lHCryptKey Then lResult = CryptDestroyKey(lHCryptKey)
  'Destroy hash object.
  If lHHash Then lResult = CryptDestroyHash(lHHash)
  'Release provider handle.
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
      
  Select Case bValid
    Case True
      lStatus = CFB_VALID
    Case Else
      lStatus = CFB_READY
  End Select

Exit Sub

ErrValidate:
  sLastError = "ErrValidate " & Error$
  Resume

End Sub

Public Sub Encrypt()
  'Encrypt InBuffer into OutBuffer
   
  Dim lHExchgKey As Long
  Dim lHCryptprov As Long
  Dim lHHash As Long
  Dim lhKey As Long
  Dim lResult As Long
  Dim sContainer As String
  Dim sProvider As String
  Dim sCryptBuffer As String
  Dim lCryptLength As Long
  Dim lCryptBufLen As Long
   
  sLastError = ""
   
  If sInBuffer = "" Then Exit Sub
   
  On Error GoTo ErrEncrypt
   
  'switch Status property
  lStatus = CFB_BUSY
   
  'Get handle to the default provider
  sContainer = vbNullChar
  sProvider = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
  lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)
  sLastError = sLastError & lResult & "-"
  
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptAcquireContext!"
    GoTo Done
  End If
   
  'Create a hash object.
  lResult = CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptCreateHash!"
    GoTo Done
  End If
    
  'Hash in the password data.
  lResult = CryptHashData(lHHash, sPassword, Len(sPassword), 0)
  sLastError = sLastError & lResult & "-"
   
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptHashData!"
    GoTo Done
  End If
    
  'Derive a session key from the hash object.
  lResult = CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lhKey)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptDeriveKey!"
    GoTo Done
  End If
   
  'Destroy the hash object.
  lResult = CryptDestroyHash(lHHash)
  sLastError = sLastError & lResult & "-"
   
  lHHash = 0
   
  'Prepare a string buffer for the CryptEncrypt function
  lCryptLength = Len(sInBuffer)
  lCryptBufLen = lCryptLength * 2
  sCryptBuffer = String(lCryptBufLen, vbNullChar)
  LSet sCryptBuffer = sInBuffer
   
  'Encrypt data
  lResult = CryptEncrypt(lhKey, 0, 1, 0, sCryptBuffer, lCryptLength, lCryptBufLen)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptEncrypt!"
  End If
   
  sOutBuffer = Mid$(sCryptBuffer, 1, lCryptLength)
     
Done:

  'Destroy session key.
  If (lhKey) Then lResult = CryptDestroyKey(lhKey)
   
  'Release key exchange key handle.
  If lHExchgKey Then CryptDestroyKey (lHExchgKey)
   
  'Destroy hash object.
  If lHHash Then CryptDestroyHash (lHHash)
   
  'Release provider handle.
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
   
  'switch Status property
  lStatus = CFB_READY
   
Exit Sub

ErrEncrypt:
  sLastError = "ErrEncrypt " & Error$
  Resume

End Sub


Public Sub Decrypt()

  'Decrypt InBuffer into OutBuffer
  Dim lHExchgKey As Long
  Dim lHCryptprov As Long
  Dim lHHash As Long
  Dim lhKey As Long
  Dim lResult As Long
   
  Dim sContainer As String
  Dim sProvider As String
   
  Dim sCryptBuffer As String
  Dim lCryptBufLen As Long
  Dim lCryptPoint As Long
   
  Dim lPasswordPoint As Long
  Dim lPasswordCount As Long
       
  sLastError = ""
       
  If sInBuffer = "" Then Exit Sub
     
On Error GoTo ErrDecrypt

  'switch Status property
  lStatus = CFB_BUSY
           
  'Init sOutBuffer
  sOutBuffer = ""
   
  'Get handle to the default provider.
  sContainer = vbNullChar
  sProvider = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
   
  lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptAcquireContext!"
    GoTo Done
  End If
   
  'Create a hash object.
  lResult = CryptCreateHash(lHCryptprov, CALG_MD5, 0, 0, lHHash)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptCreateHash!"
    GoTo Done
  End If
   
  'Hash in the password data.
  lResult = CryptHashData(lHHash, sPassword, Len(sPassword), 0)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptHashData!"
    GoTo Done
  End If
       
  'Derive a session key from the hash object.
  lResult = CryptDeriveKey(lHCryptprov, ENCRYPT_ALGORITHM, lHHash, 0, lhKey)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    sLastError = sLastError & "Error " & CStr(GetLastError) & " during CryptDeriveKey!"
    GoTo Done
  End If
   
  'Destroy the hash object.
  lResult = CryptDestroyHash(lHHash)
  sLastError = sLastError & lResult & "-"
  lHHash = 0
   
  'Prepare sCryptBuffer for CryptDecrypt
  lCryptBufLen = Len(sInBuffer) * 2
  sCryptBuffer = String(lCryptBufLen, vbNullChar)
  LSet sCryptBuffer = sInBuffer
   
  'Decrypt data
  lResult = CryptDecrypt(lhKey, 0, 1, 0, sCryptBuffer, lCryptBufLen)
  If (lResult = 0) Then
    sLastError = "Error " & CStr(GetLastError) & " during CryptDecrypt!"
    GoTo Done
  End If
   
  'Apply decrypted string from sCryptBuffer to private buffer for OutBuffer property
  sOutBuffer = Mid$(sCryptBuffer, 1, lCryptBufLen / 2)

Done:

  'Destroy session key.
  If (lhKey) Then lResult = CryptDestroyKey(lhKey)
   
  'Release key exchange key handle.
  If lHExchgKey Then lResult = CryptDestroyKey(lHExchgKey)
   
  'Destroy hash object.
  If lHHash Then lResult = CryptDestroyHash(lHHash)
   
  'Release provider handle.
  If lHCryptprov Then lResult = CryptReleaseContext(lHCryptprov, 0)
   
  'switch Status property
  lStatus = CFB_READY
   
Exit Sub

ErrDecrypt:
  sLastError = "ErrDecrypt " & Error$
  GoTo Done

End Sub

Public Property Get Status() As Long
  Status = lStatus
End Property

' Initialize Crypto Object from Computer installed CryptoAPI
Public Function InitUser() As Long
  Dim lHCryptprov As Long
  Dim lHCryptKey As Long
  Dim avProviderData(1000) As Byte
  Dim lProviderDataAddress As Long
  Dim lProviderDataLen As Long
  Dim lDataSize As Long
   
  Dim lResult As Long
  Dim sContainer As String
  Dim sProvider As String
  Dim sUserName As String
  Dim lPoint As Long
   
  Dim lMemHandle As Long
  Dim lReturn As Long
  Dim lError As Long
   
  Dim sBuffer As String
   
  On Error GoTo ErrInitUser
  'prepare string buffers
   
  sContainer = vbNullChar
  sProvider = MS_DEF_PROV & vbNullChar
   
  sLastError = ""
  lKEYSET = 0
   
  'Attempt to acquire a handle to the default key container (HKEY_CURRENT_USER).
  lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, 0)
'Stop:
lResult = 0 'THIS IS FOR TEST ONLY (MACHINE_KEY_SET)
   
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, CRYPT_NEWKEYSET)
'Stop:
lResult = 0 'THIS IS FOR TEST ONLY (MACHINE_KEY_SET)
    sLastError = sLastError & lResult & "-"
  End If
   
  If (lResult = 0) Then
    'Create default key container (HKEY_CURRENT_USER).
    'GHE 1999/10/05: The following mechanism tries to use registry
    '   in HKEY_LOCAL_MACHINE instead of HKEY_CURRENT_USER
    '   Usage: when not logged on interactively
    '   Is implemented in NT4 SP2 or higher

    'GHE: Try to get the container (HKEY_LOCAL_MACHINE)
    lKEYSET = CRYPT_MACHINE_KEYSET
    lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, CRYPT_MACHINE_KEYSET)
    sLastError = sLastError & lResult & "-"
  End If
         
  If (lResult = 0) Then
    'GHE: Try to create the container (HKEY_LOCAL_MACHINE)
    lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, CRYPT_MACHINE_NEWKEYSET)
    sLastError = sLastError & lResult & "-"
  End If
   
  If (lResult = 0) Then
    lError = GetLastError
    sLastError = sLastError & "Error creating key container! [" _
        & CStr(lError) & "] " & GetSystemMessageText(lError)
    Exit Function
  End If
       
  'Get name of default key container.
  lProviderDataLen = Len(avProviderData(0)) * (UBound(avProviderData) + 1)
   
  lResult = CryptGetProvParam(lHCryptprov, PP_CONTAINER, avProviderData(0), lProviderDataLen, lKEYSET)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    lError = GetLastError
    sLastError = sLastError & "Error getting user name!  [" _
        & CStr(lError) & "] " & GetSystemMessageText(lError)
    avProviderData(0) = 0
    Exit Function
  End If
     
  'Get sUserName from avProviderData()
  lPoint = LBound(avProviderData)
  While lPoint <= UBound(avProviderData)
    If avProviderData(lPoint) <> 0 Then
      sUserName = sUserName & Chr$(avProviderData(lPoint))
    Else
      lPoint = UBound(avProviderData)
    End If
    lPoint = lPoint + 1
  Wend
  sLastError = sLastError & sUserName & "-"
     
  'Attempt to get handle to signature key
  lResult = CryptGetUserKey(lHCryptprov, AT_SIGNATURE, lHCryptKey)
  sLastError = sLastError & lResult & "-"
   
  If (lResult = 0) Then
    lError = GetLastError
    If lError = NTE_NO_KEY Then
       
      lResult = CryptGenKey(lHCryptprov, AT_SIGNATURE, 0, lHCryptKey)
      sLastError = sLastError & lResult & "-"
       
      If (lResult = 0) Then
        sLastError = sLastError & "Error during CryptGenKey! [" _
            & CStr(lError) & "] " & GetSystemMessageText(lError)
        Exit Function
      Else
        lResult = CryptDestroyKey(lHCryptprov)
        sLastError = sLastError & lResult & "-"
      End If
    Else
      sLastError = sLastError & "Error during CryptGetUserKey! [" _
          & CStr(lError) & "] " & GetSystemMessageText(lError)
      Exit Function
    End If
  End If

  'Attempt to get handle to exchange key
  lResult = CryptGetUserKey(lHCryptprov, AT_KEYEXCHANGE, lHCryptKey)
  sLastError = sLastError & lResult & "-"
  If (lResult = 0) Then
    lError = GetLastError
    If lError = NTE_NO_KEY Then
      lResult = CryptGenKey(lHCryptprov, AT_KEYEXCHANGE, 0, lHCryptKey)
      lError = GetLastError
      If Not (lResult = 0) Then
        sLastError = sLastError & "Error during CryptGenKey!  [" _
            & CStr(lError) & "] " & GetSystemMessageText(lError)
        Exit Function
      Else
        lResult = CryptDestroyKey(lHCryptprov)
        sLastError = sLastError & lResult & "-"
      End If
    Else
      sLastError = sLastError & "Error during CryptGetUserKey! [" _
          & CStr(lError) & "] " & GetSystemMessageText(lError)
      Exit Function
    End If
  End If
   
  'release handle to provider
  lResult = CryptReleaseContext(lHCryptprov, 0)
  sLastError = sLastError & lResult & "-"
  InitUser = True

Exit Function

ErrInitUser:
  sLastError = sLastError & "ErrInitUser " & Error$
  Resume

End Function

Private Sub Class_Initialize()
  InitUser
End Sub

Public Property Get Password() As String
    Password = sPassword
End Property

Public Property Let Password(vNewValue As String)
    sPassword = vNewValue
End Property

Private Function GetSystemMessageText(msgID As Long) As String
  'Function Incapsulating FormatMessage API call for system errors
  Dim ret As Long
  Dim sBuff As String
    
  sBuff = Space$(MAX_PATH)
  ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, _
                      0&, msgID, 0&, _
                      sBuff, Len(sBuff), 0&)
    
  sBuff = Left$(sBuff, ret)
  If ret Then
      GetSystemMessageText = sBuff
  Else
      GetSystemMessageText = "(Unknown error)"
  End If
End Function


