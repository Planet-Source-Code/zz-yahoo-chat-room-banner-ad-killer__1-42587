Attribute VB_Name = "Module2"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long '
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" _
(ByVal hKey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long








'-------------------------------------------------------------------

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private MyMousePos As POINTAPI 'for getting the mouse positioning

Public Const conHwndTopmost = -1
Public Const conHwndNoTopmost = -2
Public Const conSwpNoActivate = &H10
Public Const conSwpShowWindow = &H40
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long




Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Const KEY_ALL_ACCESS = &H3F
Public Const ERROR_NONE = 0

Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const KEY_CREATE_LINK = &H20
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = (KEY_READ)
Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
   On Error GoTo QueryValueExError
   Dim cch As Long
   Dim lrc As Long
   Dim lType As Long
   Dim lValue As Long
   Dim nLoop As Long
   Dim sValue As String
   Dim sBinaryString As String
   lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then Error 5
   Select Case lType
      Case REG_SZ:
           sValue = String(cch, 0)
           lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
           If lrc = ERROR_NONE Then
              vValue = Left$(sValue, cch - 1)
           Else
              vValue = Empty
           End If
      Case REG_BINARY
           sValue = String(cch, 0)
           lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
           If lrc = ERROR_NONE Then
              vValue = sValue
           Else
              vValue = Empty
           End If
           sBinaryString = ""
           For nLoop = 1 To Len(sValue)
               sBinaryString = sBinaryString & Format$(Hex(Asc(Mid$(vValue, nLoop, 1))), "00") & " "
           Next
           vValue = sBinaryString
      Case REG_DWORD:
           lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
           If lrc = ERROR_NONE Then vValue = lValue
      Case Else
           lrc = -1
   End Select
QueryValueExExit:
   QueryValueEx = lrc
   Exit Function
QueryValueExError:
   Resume QueryValueExExit
End Function
Public Function QueryValue(ByVal hKey As Long, sKeyName As String, sValueName As String) As String
   Dim lRetVal As Long
   Dim vValue As Variant
   lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
   lRetVal = QueryValueEx(hKey, sValueName, vValue)
   QueryValue = vValue
   RegCloseKey (hKey)
End Function
Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
   Dim hNewKey As Long
   Dim lRetVal As Long
   lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
   RegCloseKey (hNewKey)
End Sub
Public Sub SetKeyValue(ByVal hKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
   Dim lRetVal As Long
   lRetVal = RegOpenKeyEx(hKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
   lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
   RegCloseKey (hKey)
End Sub
Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
   Dim lValue As Long
   Dim sValue As String
   Select Case lType
      Case REG_SZ
           sValue = vValue & Chr$(0)
           SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
      Case REG_DWORD
           lValue = vValue
           SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
   End Select
End Function
Private Sub ParseKey(KeyName As String, Keyhandle As Long)
    On Error Resume Next
    Dim rtn
rtn = InStr(KeyName, "\") 'return if "\" is contained in the Keyname
If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then 'if the is a "\" at the end of the Keyname then
   MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName 'display error to the user
   Exit Sub 'exit the procedure
ElseIf rtn = 0 Then 'if the Keyname contains no "\"
  ' Keyhandle = GetMainKeyHandle(Keyname)
   KeyName = "" 'leave Keyname blank
Else 'otherwise, Keyname contains "\"
'   Keyhandle = GetMainKeyHandle(Left(Keyname, rtn - 1)) 'seperate the Keyname
   KeyName = Right(KeyName, Len(KeyName) - rtn)
End If
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
' Set up default value
If Not IsEmpty(default) Then
  GetSettingString = default
Else
  GetSettingString = ""
End If
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lRegResult = ERROR_SUCCESS Then
  If lValueType = REG_SZ Then
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
     intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left$(strBuffer, intZeroPos - 1)
    Else
      GetSettingString = strBuffer
    End If
  End If
Else
End If
lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
If lRegResult <> ERROR_SUCCESS Then
End If
lRegResult = RegCloseKey(hCurKey)
End Sub
Public Sub CreateKey(hKey As Long, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegCreateKey(hKey, strPath, hCurKey)
If lRegResult <> ERROR_SUCCESS Then
End If
lRegResult = RegCloseKey(hCurKey)
End Sub

Sub InitWindow(tForm As Form)
Dim tX As Long
Dim tY As Long
tX = Screen.TwipsPerPixelX
tY = Screen.TwipsPerPixelY
SetWindowPos tForm.hWnd, conHwndTopmost, tForm.Left / tX, tForm.Top / tY, tForm.Width / tX, tForm.Height / tY, conSwpNoActivate Or conSwpShowWindow

End Sub

Sub UnInitWindow(tForm As Form)
Dim tX As Long
Dim tY As Long
tX = Screen.TwipsPerPixelX
tY = Screen.TwipsPerPixelY
SetWindowPos tForm.hWnd, conHwndNoTopmost, tForm.Left / tX, tForm.Top / tY, tForm.Width / tX, tForm.Height / tY, conSwpNoActivate Or conSwpShowWindow

End Sub

Sub WipeFileClean(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, Offset As Long
    'Create two buffers with a specified wipe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wipe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    For iLoop = 1 To Blocks
        Offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, Offset, Block2
    Next iLoop
    Close hFileHandle
    'Now you can delete the file, which contains no sensitive data
    Kill sFileName
End Sub
Function FileExists(Filename As String) As Boolean
    On Error Resume Next
    Dim X As Long
    X = Len(Dir$(Filename))
    If Err Or X = 0 Then FileExists = False Else FileExists = True
End Function
Public Function Get_From_INI(AppName$, KeyName$, Filename$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
Get_From_INI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), Filename$))
'To write to an ini type this
'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\organize\options.ini")

'To read do this
'Color$ = Get_From_INI("ascii", "Color", App.Path + "\organize\options.ini")
'If Color$ = "bbb" Then
'what ever
'Else
'what ever
'end if
End Function

Public Function GenFileFromRes(resID As Long, resSECTION As String, fEXT As String, Optional fPath As String = "", Optional fNAME As String = "temp", Optional FullName As String = "") As String
    On Error GoTo ErrorGenFileFromRes
    Dim resBYTE() As Byte
    If fPath = "" Then fPath = App.Path
    If fNAME = "" Then fNAME = "temp"
    resBYTE = LoadResData(resID, resSECTION)
    If FullName = "" Then
        Open fPath & "\" & fNAME & "." & fEXT For Binary Access Write As #1
    Else
        Open FullName For Binary Access Write As #1
    End If
    Put #1, , resBYTE
    Close #1
    If FullName = "" Then
        GenFileFromRes = fPath & "\" & fNAME & "." & fEXT
    Else
        GenFileFromRes = FullName
    End If
    Exit Function
ErrorGenFileFromRes:
    GenFileFromRes = ""
    MsgBox Err & ":Error in GenFileFromRes.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function

Public Function GetSysDir() As String
    On Error GoTo ErrorGetSysDir
    Dim rSTR As String
    Dim rLEN As Long
    rSTR = String(255, 0)
    rLEN = GetSystemDirectory(rSTR, Len(rSTR))
    If rLEN < Len(rSTR) Then
        rSTR = Left(rSTR, rLEN)
        If Right(rSTR, 1) = "\" Then
            GetSysDir = Left(rSTR, Len(rSTR) - 1)
        Else
            GetSysDir = rSTR
        End If
    Else
        GetSysDir = ""
    End If
    Exit Function
ErrorGetSysDir:
    GetSysDir = ""
    MsgBox Err & ":Error in GetSysDir.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Function
End Function
Function REG_ExportNode(sKeyPath As String, sOutFile As String)
    'Example:
    'REG_ExportNode "HKEY_LOCAL_MACHINE\software
    '     \microsoft","c:\windows\desktop\out.reg"
    '
    '
    '/E (Export) switch
    Shell "regedit /E " & sOutFile & " " & sKeyPath
End Function
Function REG_ImportNode(sInFile As String)
    '
    'Example:
    'REG_ImportNode "c:\windows\desktop\reg.reg"
    '
    '
    '/I (Import) /S (Silent) switchs
    Shell "regedit /I /S " & sInFile
End Function
Function SetDWORDValue(Subkey As String, Entry As String, Value As Long)
'Call ParseKey(Subkey, MainKeyHandle)
If MainKeyHandle Then
   rtn = RegOpenKeyEx(MainKeyHandle, Subkey, 0, KEY_WRITE, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key was open successfully then
     ' rtn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4) 'write the value
      If Not rtn = ERROR_SUCCESS Then   'if there was an error writting the value
         If DisplayErrorMsg = True Then 'if the user want errors displayed
           ' MsgBox ErrorMsg(rtn)        'display the error
         End If
      End If
      rtn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      If DisplayErrorMsg = True Then 'if the user want errors displayed
        ' MsgBox ErrorMsg(rtn) 'display the error
      End If
   End If
End If
End Function
Public Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
Dim lRegResult As Long
lRegResult = RegDeleteKey(hKey, strPath)
End Sub
Public Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegDeleteValue(hCurKey, strValue)
lRegResult = RegCloseKey(hCurKey)
End Sub
