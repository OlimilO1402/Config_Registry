Attribute VB_Name = "Registry"
Option Explicit

Private Const REG_NONE                  As Long = 0 ' No value type
Private Const REG_ERROR_NONE            As Long = 0 ' &H0&
Private Const REG_SZ                    As Long = 1 ' Unicode null terminated string
Private Const REG_EXPAND_SZ             As Long = 2 ' Unicode null terminated string (with environment variable references)
Private Const REG_BINARY                As Long = 3 ' Free form binary
Private Const REG_DWORD                 As Long = 4 ' 32-bit number
Private Const REG_DWORD_LITTLE_ENDIAN   As Long = 4 ' yes, twice 4 ' 32-bit number like REG_DWORD
Private Const REG_DWORD_BIG_ENDIAN      As Long = 5 ' 32-bit number
Private Const REG_LINK                  As Long = 6 ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ              As Long = 7 ' Multiple Unicode strings
Private Const REGCLS_MULTIPLEUSE        As Long = 1 ' &H1&

Private Const REG_OPTION_NON_VOLATILE   As Long = &H0& ' This key is not volatile; this is the default. The information is stored in a file and is preserved when the system is restarted. The RegSaveKey function saves keys that are not volatile.
Private Const REG_OPTION_VOLATILE       As Long = &H1& ' All keys created by the function are volatile. The information is stored in memory and is not preserved when the corresponding registry hive is unloaded. For HKEY_LOCAL_MACHINE, this occurs only when the system initiates a full shutdown. For registry keys loaded by the RegLoadKey function, this occurs when the corresponding RegUnLoadKey is performed. The RegSaveKey function does not save volatile keys. This flag is ignored for keys that already exist. Note  On a user selected shutdown, a fast startup shutdown is the default behavior for the system.
Private Const REG_OPTION_CREATE_LINK    As Long = &H2& ' Note Registry symbolic links should only be used for for application compatibility when absolutely necessary. This key is a symbolic link. The target path is assigned to the L"SymbolicLinkValue" value of the key. The target path must be an absolute registry path.
Private Const REG_OPTION_BACKUP_RESTORE As Long = &H4& ' If this flag is set, the function ignores the samDesired parameter and attempts to open the key with the access required to backup or restore the key. If the calling thread has the SE_BACKUP_NAME privilege enabled, the key is opened with the ACCESS_SYSTEM_SECURITY and KEY_READ access rights. If the calling thread has the SE_RESTORE_NAME privilege enabled, beginning with Windows Vista, the key is opened with the ACCESS_SYSTEM_SECURITY, DELETE and KEY_WRITE access rights. If both privileges are enabled, the key has the combined access rights for both privileges. For more information, see Running with Special Privileges.

Private Const REG_CREATED_NEW_KEY       As Long = &H1& ' The key did not exist and was created.
Private Const REG_OPENED_EXISTING_KEY   As Long = &H2& ' The key existed and was simply opened without being changed.

Private Const ERROR_NO_MORE_ITEMS       As Long = 259&
Private Const ERROR_MORE_DATA           As Long = 234&
Private Const MAX_KEY_LENGTH            As Long = 255&

Private Const GW_CHILD     As Long = 5
Private Const GW_HWNDFIRST As Long = 0
Private Const GW_HWNDLAST  As Long = 1
Private Const GW_HWNDNEXT  As Long = 2
Private Const GW_HWNDPREV  As Long = 3
Private Const GW_OWNER     As Long = 4
Private Const GW_MAX       As Long = 5
Private Const MaxBuff      As Long = 255

Public Enum RegistryHive
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
'  HKCR = HKEY_CLASSES_ROOT
'  HKCU = HKEY_CURRENT_USER
'  HKLM = HKEY_LOCAL_MACHINE
'  HKU = HKEY_USERS
'  HKPD = HKEY_PERFORMANCE_DATA
'  HKCC = HKEY_CURRENT_CONFIG
'  HKDD = HKEY_DYN_DATA
End Enum

' Registry key security options...
Private Const ERROR_SUCCESS            As Long = 0&
Private Const SYNCHRONIZE              As Long = &H100000
Private Const READ_CONTROL             As Long = &H20000
Private Const STANDARD_RIGHTS_ALL      As Long = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE  As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ     As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_WRITE    As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE          As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS   As Long = &H8
Private Const KEY_NOTIFY               As Long = &H10
Private Const KEY_SET_VALUE            As Long = &H2
Private Const KEY_CREATE_SUB_KEY       As Long = &H4
Private Const KEY_CREATE_LINK          As Long = &H20
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

Private Type FILETIME
    dwLoDateTime As Long
    dwHiDateTime As Long
End Type

Private Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or _
                                         READ_CONTROL Or _
                                         KEY_QUERY_VALUE Or _
                                         KEY_SET_VALUE Or _
                                         KEY_CREATE_SUB_KEY Or _
                                         KEY_ENUMERATE_SUB_KEYS Or _
                                         KEY_NOTIFY Or _
                                         KEY_CREATE_LINK) And (Not SYNCHRONIZE))
                                          
Private Const KEY_READ       As Long = ((STANDARD_RIGHTS_READ Or _
                                         KEY_QUERY_VALUE Or _
                                         KEY_ENUMERATE_SUB_KEYS Or _
                                         KEY_NOTIFY) And (Not SYNCHRONIZE))
                                   
Private Const KEY_WRITE      As Long = ((STANDARD_RIGHTS_WRITE Or _
                                         KEY_SET_VALUE Or _
                                         KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
                                         
Private Const KEY_EXECUTE    As Long = (KEY_READ)


#If VBA7 Then

    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryinfokeyw
    'LSTATUS RegQueryInfoKeyW([in] HKEY hKey, [out, optional] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, LPDWORD lpReserved, [out, optional] LPDWORD lpcSubKeys,
    '                         [out, optional] LPDWORD lpcbMaxSubKeyLen, [out, optional] LPDWORD lpcbMaxClassLen, [out, optional] LPDWORD lpcValues, [out, optional] LPDWORD lpcbMaxValueNameLen,
    '                         [out, optional] LPDWORD lpcbMaxValueLen,  [out, optional] LPDWORD lpcbSecurityDescriptor, [out, optional] PFILETIME lpftLastWriteTime
    Private Declare PtrSafe Function RegQueryInfoKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpClass As LongPtr, ByVal lpcbClass As LongPtr, ByVal lpReserved As LongPtr, ByVal lpcSubKeys As LongPtr, _
                                                                      ByVal lpcbMaxSubKeyLen As LongPtr, ByVal lpcbMaxClassLen As LongPtr, ByVal lpcValues As LongPtr, ByVal lpcbMaxValueNameLen As LongPtr, _
                                                                      ByVal lpcbMaxValueLen As LongPtr, ByVal lpcbSecurityDescriptor As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryvalueexw
    'LSTATUS RegQueryValueExW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare PtrSafe Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumkeyexw
    'LSTATUS RegEnumKeyExW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpName, [in, out] LPDWORD lpcchName, LPDWORD lpReserved,
    '                      [in, out] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, [out, optional] PFILETIME lpftLastWriteTime);
    Private Declare PtrSafe Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As LongPtr, ByVal lpcchName As LongPtr, ByVal lpReserved As LongPtr, _
                                                                   ByVal lpClass As LongPtr, ByVal lpcchClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexw
    'LSTATUS RegOpenKeyExW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [in] DWORD ulOptions, [in] REGSAM samDesired, [out] PHKEY phkResult);
    Private Declare PtrSafe Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
    'LSTATUS RegCloseKey([in] HKEY hKey);
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyexw
    'LSTATUS RegCreateKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, DWORD Reserved, [in, optional] LPWSTR lpClass, [in] DWORD dwOptions, [in] REGSAM samDesired, [in, optional] const LPSECURITY_ATTRIBUTES lpSecurityAttributes, [out] PHKEY phkResult, [out, optional] LPDWORD lpdwDisposition);
    Private Declare PtrSafe Function RegCreateKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal Reserved As Long, ByVal lpClass As LongPtr, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As LongPtr, ByVal phkResult As LongPtr, ByVal lpdwDisposition As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyw
    'LSTATUS RegCreateKeyW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [out] PHKEY phkResult);
    Private Declare PtrSafe Function RegCreateKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvaluew
    'LSTATUS RegSetValueW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [in] DWORD dwType, [in] LPCWSTR lpData, [in] DWORD cbData);
    Private Declare PtrSafe Function RegSetValueW Lib "advapi32" (ByVal RtKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvalueexw
    'LSTATUS RegSetValueExW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName, DWORD Reserved, [in] DWORD dwType, [in] const BYTE *lpData, [in] DWORD cbData);
    Private Declare PtrSafe Function RegSetValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regflushkey
    'LSTATUS RegFlushKey([in] HKEY hKey);
    Private Declare PtrSafe Function RegFlushKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyw
    'LSTATUS RegDeleteKeyW([in] HKEY hKey, [in] LPCWSTR lpSubKey);
    Private Declare PtrSafe Function RegDeleteKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletevaluew
    'LSTATUS RegDeleteValueW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName);
    Private Declare PtrSafe Function RegDeleteValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr) As Long
    
#Else
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryinfokeyw
    'LSTATUS RegQueryInfoKeyW([in] HKEY hKey, [out, optional] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, LPDWORD lpReserved, [out, optional] LPDWORD lpcSubKeys,
    '                         [out, optional] LPDWORD lpcbMaxSubKeyLen, [out, optional] LPDWORD lpcbMaxClassLen, [out, optional] LPDWORD lpcValues, [out, optional] LPDWORD lpcbMaxValueNameLen,
    '                         [out, optional] LPDWORD lpcbMaxValueLen,  [out, optional] LPDWORD lpcbSecurityDescriptor, [out, optional] PFILETIME lpftLastWriteTime
    Private Declare Function RegQueryInfoKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpClass As LongPtr, ByVal lpcbClass As LongPtr, ByVal lpReserved As LongPtr, ByVal lpcSubKeys As LongPtr, _
                                                              ByVal lpcbMaxSubKeyLen As LongPtr, ByVal lpcbMaxClassLen As LongPtr, ByVal lpcValues As LongPtr, ByVal lpcbMaxValueNameLen As LongPtr, _
                                                              ByVal lpcbMaxValueLen As LongPtr, ByVal lpcbSecurityDescriptor As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryvalueexw
    'LSTATUS RegQueryValueExW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumkeyexw
    'LSTATUS RegEnumKeyExW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpName, [in, out] LPDWORD lpcchName, LPDWORD lpReserved,
    '                      [in, out] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, [out, optional] PFILETIME lpftLastWriteTime);
    Private Declare Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As LongPtr, ByVal lpcchName As LongPtr, ByVal lpReserved As LongPtr, _
                                                           ByVal lpClass As LongPtr, ByVal lpcchClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexw
    'LSTATUS RegOpenKeyExW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [in] DWORD ulOptions, [in] REGSAM samDesired, [out] PHKEY phkResult);
    Private Declare Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
    'LSTATUS RegCloseKey([in] HKEY hKey);
    Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyexw
    'LSTATUS RegCreateKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, DWORD Reserved, [in, optional] LPWSTR lpClass, [in] DWORD dwOptions, [in] REGSAM samDesired, [in, optional] const LPSECURITY_ATTRIBUTES lpSecurityAttributes, [out] PHKEY phkResult, [out, optional] LPDWORD lpdwDisposition);
    Private Declare Function RegCreateKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal Reserved As Long, ByVal lpClass As LongPtr, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As LongPtr, ByVal phkResult As LongPtr, ByVal lpdwDisposition As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyw
    Private Declare Function RegCreateKeyW Lib "advapi32" (ByVal RtKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvaluew
    Private Declare Function RegSetValueW Lib "advapi32" (ByVal RtKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvalueexw
    Private Declare Function RegSetValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regflushkey
    Private Declare Function RegFlushKey Lib "advapi32" (ByVal RtKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyw
    Private Declare Function RegDeleteKeyW Lib "advapi32" (ByVal RtKey As LongPtr, ByVal lpSubKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletevaluew
    Private Declare Function RegDeleteValueW Lib "advapi32" (ByVal RtKey As LongPtr, ByVal lpValueName As LongPtr) As Long
    
#End If

'private members
Private mRootKey     As LongPtr 'handle to the current rootkey-hive
Private mCurrentKey  As LongPtr 'handle to the current Key
Private mCurrentPath As String  'path to the actual key as String
Private mLazyWrite   As Boolean
Private mAccess      As Long

Public Sub Init()
    mRootKey = HKEY_CURRENT_USER
    mCurrentKey = mRootKey
    mCurrentPath = ""
    mLazyWrite = True
    mAccess = KEY_ALL_ACCESS
End Sub


Public Function App_RegisterAsComServer(ByVal exePath As String, _
                                        ByVal activatorCLSID As String) As Boolean
    Dim ret As Boolean
    Dim hKey As LongPtr
    If RegCreateKeyExW(HKEY_CURRENT_USER, _
                       StrPtr("SOFTWARE\Classes\CLSID\" & activatorCLSID & "\LocalServer32"), _
                       0&, _
                       vbNullString, 0&, KEY_ALL_ACCESS, _
                       0&, hKey, 0&) = REG_ERROR_NONE Then
        If hKey <> 0& Then
            If RegSetValueExW(hKey, vbNullString, _
                              0&, REG_SZ, _
                              StrPtr(exePath & vbNullChar), _
                              Len(exePath)) = REG_ERROR_NONE Then
                ret = True
            End If
            Call RegCloseKey(hKey)
        End If
    End If
    App_RegisterAsComServer = ret
End Function

Public Function App_UnregisterAsComServer(ByVal activatorCLSID As String) As Boolean
    Dim ret As Boolean
    If RegDeleteKeyW(HKEY_CURRENT_USER, _
                    StrPtr("SOFTWARE\Classes\CLSID\" & activatorCLSID & "\LocalServer32")) = REG_ERROR_NONE Then
        If RegDeleteKeyW(HKEY_CURRENT_USER, _
                         StrPtr("SOFTWARE\Classes\CLSID\" & activatorCLSID)) = REG_ERROR_NONE Then
            ret = True
        End If
    End If
    App_UnregisterAsComServer = ret
End Function

Public Function App_IsRegisteredAsComServer(ByVal activatorCLSID As String) As Boolean
    Dim ret As Boolean
    Dim hKey As LongPtr
    If RegOpenKeyExW(HKEY_CURRENT_USER, _
                    StrPtr("SOFTWARE\Classes\CLSID\" & activatorCLSID), _
                    0&, _
                    KEY_QUERY_VALUE, _
                    hKey) = REG_ERROR_NONE Then
        If hKey <> 0& Then
            ret = True
            Call RegCloseKey(hKey)
        End If
    End If
    App_IsRegisteredAsComServer = ret
End Function


'registers and associates a file extension to a certain program
Public Sub RegisterShellFileTypes(ByVal FileExtension As String, _
                                  ByVal sAppReg As String, _
                                  ByVal sAppName As String, _
                                  ByVal aPFN As String, _
                                  ByVal IconId As Long)
Try: On Error GoTo Catch
    Init
    If Left(FileExtension, 1) <> "." Then FileExtension = "." & FileExtension
    
    'associates the file extension
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey FileExtension
    Registry.WriteString vbNullString, sAppReg
    Registry.CloseKey
    
    'generates the new entry sAppReg
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg
    Registry.WriteString vbNullString, sAppReg
    Registry.CloseKey
    
    'creates the link to the file icon
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg & "\DefaultIcon"
    Registry.WriteString vbNullString, """" & aPFN & """" & "," & CStr(IconId)
    Registry.CloseKey
    
    'sets the path to execute the app
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.CreateKey sAppReg & "\shell\open\command"
    Dim StrKeyVal As String: StrKeyVal = """" & aPFN & """" & " %" & CStr(IconId)  'should we add quotes " here?
    Registry.WriteString vbNullString, StrKeyVal
    GoTo Finally
Catch:
    'end silently?
    ErrHandler "RegisterShellFileTypes"
Finally:
    Registry.CloseKey
End Sub

'unregisters a file extension from a certain program
Public Sub UnRegisterShellFileTypes(ByVal FileExtension As String, ByVal sAppReg As String)
Try: On Error GoTo Catch
    Init
    'deletes the path to execute the app
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg & "\shell\open\command"
    Registry.CloseKey
    
    'deletes the link to the icon
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg & "\DefaultIcon"
    Registry.CloseKey
    
    'deletes the entry sAppReg
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey sAppReg
    Registry.CloseKey
    
    'deletes the association to the fil extension
    Registry.RootKey = HKEY_CLASSES_ROOT
    Registry.DeleteKey "." & FileExtension
    GoTo Finally
Catch:
    ErrHandler "UnRegisterShellFileTypes"
Finally:
    Registry.CloseKey
End Sub

Private Function RegistryHive_ToStr(E As RegistryHive) As String
    Dim s As String
    Select Case E
    Case HKEY_CLASSES_ROOT:     s = "HKEY_CLASSES_ROOT"
    Case HKEY_CURRENT_USER:     s = "HKEY_CURRENT_USER"
    Case HKEY_LOCAL_MACHINE:    s = "HKEY_LOCAL_MACHINE"
    Case HKEY_USERS:            s = "HKEY_USERS"
    Case HKEY_PERFORMANCE_DATA: s = "HKEY_PERFORMANCE_DATA"
    Case HKEY_CURRENT_CONFIG:   s = "HKEY_CURRENT_CONFIG"
    Case HKEY_DYN_DATA:         s = "HKEY_DYN_DATA"
    End Select
    RegistryHive_ToStr = s
End Function
Private Function RegistryHive_Parse(ByVal s As String) As RegistryHive
    s = UCase(s)
    Dim E As RegistryHive
    Select Case s
    Case "HKEY_CLASSES_ROOT", "HKCR":     E = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER", "HKCU":     E = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE", "HKLM":    E = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS", "HKU":             E = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA", "HKPD": E = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG", "HKCC":   E = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA", "HKDD":         E = HKEY_DYN_DATA
    End Select
    RegistryHive_Parse = E
End Function

' v ########################## v ' Properties ' v ########################## v '
Public Property Get RootKey() As RegistryHive
    RootKey = mRootKey
End Property
Public Property Let RootKey(Value As RegistryHive)
    mRootKey = Value
End Property
Public Property Get RootKeyToStr() As String
    RootKeyToStr = RegistryHive_ToStr(mRootKey)
End Property

Public Property Get CurrentKey() As LongPtr ' ReadOnly
    CurrentKey = mCurrentKey
End Property

Public Property Get CurrentPath() As String ' ReadOnly
    CurrentPath = RootKeyToStr & "\" & mCurrentPath
End Property

Public Property Get LazyWrite() As Boolean
    LazyWrite = mLazyWrite
End Property
Public Property Let LazyWrite(Value As Boolean)
    mLazyWrite = Value
End Property

Public Property Get access() As Long
    access = mAccess
End Property
Public Property Let access(Value As Long)
    mAccess = Value
End Property
' ^ ########################## ^ ' Properties ' ^ ########################## ^ '

Public Function KeyExists(Key As String) As Boolean
    'asks if the Key exists and closes it immediately
Try: On Error GoTo Catch
    KeyExists = KeyExistsNoClose(Key)
    If KeyExists Then
        Dim lResult As Long: lResult = RegCloseKey(mCurrentKey)
    End If
    Exit Function
Catch:
    ErrHandler "KeyExists", "key: """ & Key & """", lResult
End Function

Public Function KeyExistsNoClose(Key As String) As Boolean
    'asks if the Key exists without closing it
Try: On Error GoTo Catch
    Dim lResult As Long, HandleKey As LongPtr
    lResult = RegOpenKeyExW(mRootKey, StrPtr(Key), 0&, KEY_READ, VarPtr(HandleKey))
    KeyExistsNoClose = (lResult = ERROR_SUCCESS)
    If KeyExistsNoClose Then
        mCurrentKey = HandleKey
        mCurrentPath = Key
    Else
        GoTo Catch
    End If
    Exit Function
Catch:
    ErrHandler "KeyExistsNoClose", "Path: " & CurrentPath & " key: """ & Key & """", lResult
End Function

Public Function CreateKey(Key As String) As Boolean
'Key may be an absolute or a relative key name.
'an absolute Key starts with a backslash and sits directly on the rootkey
'a relative Key is a subkey of the current key (without a backSlash at the beginning)
    Dim lResult As Long, HandleKey As LongPtr, lAction As Long, Class As String '
Try: On Error GoTo Catch
    lResult = RegCreateKeyExW(mRootKey, StrPtr(Key), 0&, StrPtr(Class), REG_OPTION_NON_VOLATILE, mAccess, 0&, VarPtr(HandleKey), VarPtr(lAction))
    If lResult = ERROR_SUCCESS Then
        mCurrentPath = Key
        mCurrentKey = HandleKey
        If RegFlushKey(mCurrentKey) = ERROR_SUCCESS Then
            lResult = RegCloseKey(mCurrentKey)
        End If
        CreateKey = (lAction = REG_CREATED_NEW_KEY)
    'Else
    '    CreateKey = False
    End If
    Exit Function
Catch:
    ErrHandler "CreateKey", "Key: """ & Key & """", lResult
End Function

Public Function OpenKey(Key As String, CanCreate As Boolean) As Boolean
    'opens the Key, if not existing, then cancreate decides to create it or not
Try: On Error GoTo Catch
    OpenKey = KeyExistsNoClose(Key)
    If Not OpenKey Then
        If CanCreate Then
            OpenKey = CreateKey(Key)
        End If
    End If
    Exit Function
Catch:
    ErrHandler "OpenKey", "Key: """ & Key & """"
End Function

Public Function OpenKeyReadOnly(Key As String) As Boolean
Try: On Error GoTo Catch
    Dim mA As Long: mA = mAccess
    mAccess = KEY_READ
    OpenKeyReadOnly = KeyExistsNoClose(Key)
    mAccess = mA
    Exit Function
Catch:
    ErrHandler "OpenKeyReadOnly", "Key: """ & Key & """"
End Function

Public Sub CloseCurrentKey()
    'Diese Methode schreibt den aktuellen Schlüssel in die Registrierdatenbank und schließt ihn.
    RegCloseKey mCurrentKey
    mCurrentKey = 0
End Sub

Public Sub CloseKey()
    RegCloseKey mCurrentKey
    RegCloseKey mRootKey
    mCurrentKey = 0
    mRootKey = 0
End Sub

Public Function DeleteKey(Key As String) As Boolean
Try: On Error GoTo Catch
    Dim hResult As Long
    hResult = RegDeleteKeyW(mRootKey, StrPtr(Key))
    DeleteKey = hResult = ERROR_SUCCESS
    If DeleteKey Then Exit Function
    Exit Function
Catch:
    ErrHandler "DeleteKey", "Key: """ & Key & """", hResult
End Function

Public Sub MoveKey(OldName As String, NewName As String, Delete As Boolean)
  'ToDo
End Sub

' v ############################## v '   For Hive-File   ' v ############################## v '
Public Function LoadKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function

Public Function ReplaceKey(Key As String, FileName As String, BackUpFileName As String) As Boolean
  'ToDo
End Function

Public Function RestoreKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function

Public Function SaveKey(Key As String, FileName As String) As Boolean
  'ToDo
End Function

Public Function UnLoadKey(Key As String) As Boolean
  'ToDo
End Function

Public Function HasSubKeys() As Boolean
  'ToDo
End Function

' v ############################## v '    concerning Values    ' v ############################## v '
Public Function ValueExists(Name As String) As Boolean
Try: On Error GoTo Catch
    Dim HandleKey As LongPtr: HandleKey = mCurrentKey
    ValueExists = ValueExistsNoClose(Name) ', HandleVal)
    If ValueExists Then RegCloseKey HandleKey
    Exit Function
Catch:
    ErrHandler "ValueExists", "Name: """ & Name & """"
End Function

Public Function ValueExistsNoClose(Name As String) As Boolean ', HandleVal As Long) As Boolean
Try: On Error GoTo Catch
    Dim lResult As Long, dwType As Long, buffersize As Long
    Dim HandleKey As LongPtr: HandleKey = mCurrentKey
    lResult = RegQueryValueExW(HandleKey, StrPtr(Name), 0&, VarPtr(dwType), 0&, VarPtr(buffersize))
    ValueExistsNoClose = (lResult = ERROR_SUCCESS)
    If ValueExistsNoClose Then
        mCurrentKey = HandleKey
    End If
    Exit Function
Catch:
    ErrHandler "ValueExistsNoClose", "Name: """ & Name & """", lResult
End Function

Public Function DeleteValue(Name As String) As Boolean
Try: On Error GoTo Catch
    Dim lResult As Long, HandleKey As LongPtr
    DeleteValue = ValueExistsNoClose(Name)
    If DeleteValue Then
        lResult = RegDeleteValueW(mCurrentKey, StrPtr(Name))
        DeleteValue = (lResult = ERROR_SUCCESS)
    End If
    Call RegCloseKey(mCurrentKey)
    Exit Function
Catch:
    ErrHandler "DeleteValue", "Name: """ & Name & """", lResult
End Function

Public Sub RenameValue(OldName As String, NewName As String)
Try: On Error GoTo Catch
    'ToDo
    Exit Sub
Catch:
    ErrHandler "RenameValue", "OldName: """ & OldName & """" & "; NewName: """ & NewName & """" ', hr
End Sub
' ^ ########################### ^ '    concerning Values    ' ^ ########################## ^ '




















' !!!!!!!! '    AB HIER NOCH ANPASSEN ' !!!!!!!! '
'######################### Get- Subs und Functions ########################
'Public Function GetDataInfo(ValueName As String, value As RegDataInfo) As Boolean
'
'End Function

Public Function GetDataSize(valueName As String) As Integer
  'ToDo
End Function

Public Sub GetDataType(valueName As String)
  'ToDo
End Sub

'Public Function GetKeyInfo(value As RegKeyInfo) As Boolean
'
'End Function
'internal unsafe string[] InternalGetSubKeyNames()
'{
'    EnsureNotDisposed();
'    int num = InternalSubKeyCount();
'    string[] array = new string[num];
'    if (num > 0)
'    {
'        char[] array2 = new char[256];
'        fixed (char* ptr = &array2[0])
'        {
'            for (int i = 0; i < num; i++)
'            {
'                int lpcbName = array2.Length;
'                int num2 = Win32Native.RegEnumKeyEx(hkey, i, ptr, ref lpcbName, null, null, null, null);
'                if (num2 != 0)
'                {
'                    Win32Error(num2, null);
'                }
'                array[i] = new string(ptr);
'            }
'        }
'    }
'    return array;
'}

'Private Const ERROR_SUCCESS             As Long = 0&
'Private Const ERROR_NO_MORE_ITEMS       As Long = 259&
'Private Const ERROR_MORE_DATA           As Long = 234&


Public Sub GetKeyNames(StrCol As Collection)
Try: On Error GoTo Catch
    If StrCol Is Nothing Then Set StrCol = New Collection
    Dim c As Long: c = GetKeyCount
    Dim ft As FILETIME
    If c > 0 Then
        'ReDim arr(0 To c - 1) As String
        Dim cbName As Long
        Dim s As String
        Dim i As Long, hr As Long
        For i = 0 To c - 1
            cbName = MAX_KEY_LENGTH
            s = Space(cbName)
            hr = RegEnumKeyExW(mCurrentKey, i, StrPtr(s), VarPtr(cbName), 0&, 0&, 0&, VarPtr(ft))
            If hr <> ERROR_SUCCESS And hr <> ERROR_NO_MORE_ITEMS Then
                GoTo Catch
            End If
            s = Left(s, cbName)
            StrCol.Add s
        Next
    End If
    Exit Sub
Catch:
    ErrHandler "GetKeyNames", , hr
End Sub

'internal int InternalSubKeyCount()
'{
'    EnsureNotDisposed();
'    int lpcSubKeys = 0;
'    int lpcValues = 0;
'    int num = Win32Native.RegQueryInfoKey(hkey, null, null, IntPtr.Zero, ref lpcSubKeys, null, null, ref lpcValues, null, null, null, null);
'    if (num != 0)
'    {
'        Win32Error(num, null);
'    }
'    return lpcSubKeys;
'}
Public Function GetKeyCount() As Long
Try: On Error GoTo Catch
    Dim cValues  As Long, ft As FILETIME, cSubKeys As Long: cSubKeys = 10000
    Dim hr As Long: hr = RegQueryInfoKeyW(mCurrentKey, ByVal 0&, ByVal 0&, ByVal 0&, VarPtr(cSubKeys), VarPtr(cSubKeys), ByVal 0&, VarPtr(cValues), ByVal 0&, ByVal 0&, ByVal 0&, ByVal 0&) 'VarPtr(ft))
    If hr = ERROR_SUCCESS Then
        GetKeyCount = cSubKeys
        Exit Function
    End If
Catch:
    ErrHandler "GetKeyCount:RegQueryInfoKey", "CurrentKey: " & CStr(mCurrentKey), hr
End Function

Public Sub GetValueNames(StrCol As Collection)
  'ToDo
End Sub

'########################### Spezial #####################################
Public Function RegistryConnect(UNCName As String) As Boolean
'Die Methode richtet eine Verbindung zur Registrierdatenbank auf einem anderen Computer ein.
  'ToDo
End Function

'vv############################### ReadFunctions und WriteSubs ########################vv
Public Function ReadCurrency(Name As String) As Currency
  'ToDo
End Function

Public Sub WriteCurrency(Name As String, Value As Currency)
  'ToDo
End Sub

Public Function ReadBinaryData(Name As String, Buffer As Variant, BufSize As Integer) As Integer
  'ToDo
End Function

Public Sub WriteBinaryData(Name As String, Buffer As Variant, BufSize As Integer)
  'ToDo
End Sub

Public Function ReadBool(Name As String) As Boolean
  'ToDo
End Function

Public Sub WriteBool(Name As String, Value As Boolean)
  'ToDo
End Sub

Public Function ReadDate(Name As String) As Date
  'ToDo
End Function

Public Sub WriteDate(Name As String, Value As Date)
  'ToDo
End Sub

Public Function ReadDateTime(Name As String) As Date
  'ToDo
End Function

Public Sub WriteDateTime(Name As String, Value As Date)
  'ToDo
End Sub

Public Function ReadTime(Name As String) As Date
  'ToDo
End Function

Public Sub WriteTime(Name As String, Value As Date)
  'ToDo
End Sub

Public Function ReadFloat(Name As String) As Double
  'ToDo
End Function

Public Sub WriteFloat(Name As String, Value As Double)
  'ToDo
End Sub

Public Function ReadInteger(Name As String) As Long
    Dim LngVal As Long
    If GetValue(mCurrentPath, Name, LngVal) Then
        ReadInteger = LngVal
    'Else
        'MsgBox "Wert: """ & Name & """ konnte nicht gelesen werden"
    End If
End Function

Public Sub WriteInteger(Name As String, ByVal Value As Long)
    Dim LngVal As Long: LngVal = Value
    If Not SetValue(mRootKey, mCurrentPath, Name, LngVal) Then
        'MsgBox "Wert: """ & Name & """ konnte nicht geschrieben werden"
    End If
End Sub

Public Function ReadString(Name As String) As String
    Dim StrVal As String
    If GetValue(mCurrentPath, Name, StrVal) Then
        ReadString = StrVal
    'Else
        'MsgBox "Wert: """ & Name & """ konnte nicht gelesen werden"
    End If
End Function

Public Sub WriteString(Name As String, ByVal Value As String)
    Dim StrVal As String: StrVal = Value
    If Not SetValue(mRootKey, mCurrentPath, Name, StrVal) Then
        'MsgBox "Wert: """ & Name & """ konnte nicht geschrieben werden"
    End If
End Sub

Public Sub WriteExpandString(Name As String, Value As String)
  'ToDo
End Sub

Private Function GetValue(Key As String, ValNam As String, VarVal As Variant) As Boolean
    'GetValue =  KeyExistsNoClose(Key) ', HandleKey)
    If Not KeyExistsNoClose(Key) Then
        Exit Function
    End If
    Dim dwType As Long
    Dim zw As Long, buffersize As Long
    Dim Buffer As String
    Dim lResult As Long: lResult = RegQueryValueExW(mCurrentKey, StrPtr(ValNam), 0&, VarPtr(dwType), ByVal 0&, VarPtr(buffersize))
    GetValue = (lResult = ERROR_SUCCESS)
    If lResult <> ERROR_SUCCESS Then Exit Function ' Field does not exist
    Select Case dwType
    Case REG_SZ, REG_EXPAND_SZ       ' nullterminated String
        Buffer = Space$(buffersize + 1)
        lResult = RegQueryValueExW(mCurrentKey, StrPtr(ValNam), 0&, VarPtr(dwType), StrPtr(Buffer), VarPtr(buffersize))
        GetValue = (lResult = ERROR_SUCCESS)
        If lResult <> ERROR_SUCCESS Then Exit Function ' error during reading the data
        Dim plen As Long
        plen = InStr(1, Buffer, vbNullChar) - 1
        If plen > 0 Then
            VarVal = Left$(Buffer, plen)
        End If
    Case REG_DWORD     ' 32-bit number word
        buffersize = 4       ' = 32 Bit
        lResult = RegQueryValueExW(mCurrentKey, StrPtr(ValNam), 0&, VarPtr(dwType), VarPtr(zw), VarPtr(buffersize))
        GetValue = (lResult = ERROR_SUCCESS)
        If lResult <> ERROR_SUCCESS Then Exit Function ' error during reading the data
        VarVal = zw
        ' Hier könnten auch die weiteren Datentypen behandelt werden, soweit dies sinnvoll ist
    End Select
    RegCloseKey mCurrentKey
    GetValue = (lResult = ERROR_SUCCESS)
End Function

Private Function SetValue(root As LongPtr, Key As String, field As String, Value As Variant) As Boolean
    Dim lResult As Long, keyhandle As LongPtr
    Dim s As String, L As Long
    lResult = RegOpenKeyExW(root, StrPtr(Key), 0, KEY_ALL_ACCESS, VarPtr(keyhandle))
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
        Exit Function
    End If
    Select Case VarType(Value)
    Case vbInteger, vbLong
        L = CLng(Value)
        lResult = RegSetValueExW(keyhandle, StrPtr(field), 0, REG_DWORD, L, 4)
    Case vbString
        's = StrConv(CStr(Value), vbFromUnicode) & vbNullString
        s = CStr(Value) & vbNullString
        lResult = RegSetValueExW(keyhandle, StrPtr(field), 0, REG_SZ, StrPtr(s), LenB(s)) ' + 1)    ' +1 for the trailing 0
        ' here you may save or change any other data type
    End Select
    RegCloseKey keyhandle
    SetValue = (lResult = ERROR_SUCCESS)
End Function

Public Function GetKeyValue(KeyRoot As LongPtr, keyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim rc         As Long                                  ' Rückgabe-Code
    Dim hKey       As LongPtr                               ' Zugriffsnummer für einen offenen Registrierungsschlüssel
    Dim hDepth     As Long                                  '
    Dim KeyValType As Long                                  ' Datentyp eines Registrierungsschlüssels
    Dim tmpVal     As String                                ' Temporärer Speicher eines Registrierungsschlüsselwertes
    Dim KeyValSize As Long                                  ' Größe der Registrierungsschlüsselvariablen
    '------------------------------------------------------------
    ' Registrierungsschlüssel unter KeyRoot {HKEY_LOCAL_MACHINE...} öffnen
    '------------------------------------------------------------
    rc = RegOpenKeyExW(KeyRoot, StrPtr(keyName), 0, KEY_ALL_ACCESS, VarPtr(hKey)) ' Registrierungsschlüssel öffnen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln...
    
    tmpVal = String$(1024, 0)                               ' Platz für Variable reservieren
    KeyValSize = 1024                                       ' Größe der Variable markieren
    
    '------------------------------------------------------------
    ' Registrierungsschlüsselwert abrufen...
    '------------------------------------------------------------
    rc = RegQueryValueExW(hKey, StrPtr(SubKeyRef), 0, VarPtr(KeyValType), VarPtr(tmpVal), VarPtr(KeyValSize))     ' Schlüsselwert abrufen/erstellen
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehler behandeln
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 fügt null-terminierte Zeichenfolge hinzu...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null gefunden, aus Zeichenfolge extrahieren
    Else                                                    ' Keine null-terminierte Zeichenfolge für WinNT...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null nicht gefunden, nur Zeichenfolge extrahieren
    End If
    '------------------------------------------------------------
    ' Schlüsselwerttyp für Konvertierung bestimmen...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Datentypen durchsuchen...
    Case REG_SZ, REG_EXPAND_SZ                              ' Zeichenfolge für Registrierungsschlüsseldatentyp
        KeyVal = tmpVal                                     ' Zeichenfolgenwert kopieren
    Case REG_DWORD                                          ' Registrierungsschlüsseldatentyp DWORD
        Dim i As Long                                           ' Schleifenzähler
        For i = Len(tmpVal) To 1 Step -1                    ' Jedes Bit konvertieren
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' DWORD in Zeichenfolge konvertieren
    End Select
    
    GetKeyValue = True                                      ' Erfolgreiche Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
    Exit Function                                           ' Beenden
    
GetKeyError:      ' Bereinigen, nachdem ein Fehler aufgetreten ist...
    KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
    GetKeyValue = False                                     ' Fehlgeschlagene Ausführung zurückgeben
    rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
End Function

' #################### ' Local ErrHandler  ' #################### '
''copy this same function to every class or form
''the name of the class or form will be added automatically
''in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
'' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    
    If bRetry Then
        
        ErrHandler = MessErrorRetry("Registry", FuncName, AddInfo, WinApiError, bErrLog)
        
    Else
        
        ErrHandler = MessError("Registry", FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)
        
    End If
    
End Function


