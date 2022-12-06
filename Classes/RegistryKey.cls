VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Imports System.Runtime.InteropServices
'
'Public NotInheritable Class RegistryKey
'Inherits MarshalByRefObject
'Implements IDisposable

#If VBA7 Then
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyw
    Private Declare PtrSafe Function RegCreateKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyexw
    'LSTATUS RegCreateKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, DWORD Reserved, [in, optional] LPWSTR lpClass, [in] DWORD dwOptions, [in] REGSAM samDesired, [in, optional] const LPSECURITY_ATTRIBUTES lpSecurityAttributes, [out] PHKEY phkResult, [out, optional] LPDWORD lpdwDisposition);
    Private Declare PtrSafe Function RegCreateKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal Reserved As Long, ByVal lpClass As LongPtr, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As LongPtr, ByVal phkResult As LongPtr, ByVal lpdwDisposition As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexw
    'LSTATUS RegOpenKeyExW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [in] DWORD ulOptions, [in] REGSAM samDesired, [out] PHKEY phkResult);
    Private Declare PtrSafe Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryinfokeyw
    'LSTATUS RegQueryInfoKeyW([in] HKEY hKey, [out, optional] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, LPDWORD lpReserved, [out, optional] LPDWORD lpcSubKeys,
    '                         [out, optional] LPDWORD lpcbMaxSubKeyLen, [out, optional] LPDWORD lpcbMaxClassLen, [out, optional] LPDWORD lpcValues, [out, optional] LPDWORD lpcbMaxValueNameLen,
    '                         [out, optional] LPDWORD lpcbMaxValueLen,  [out, optional] LPDWORD lpcbSecurityDescriptor, [out, optional] PFILETIME lpftLastWriteTime
    Private Declare PtrSafe Function RegQueryInfoKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpClass As LongPtr, ByVal lpcbClass As LongPtr, ByVal lpReserved As LongPtr, ByVal lpcSubKeys As LongPtr, _
                                                                      ByVal lpcbMaxSubKeyLen As LongPtr, ByVal lpcbMaxClassLen As LongPtr, ByVal lpcValues As LongPtr, ByVal lpcbMaxValueNameLen As LongPtr, _
                                                                      ByVal lpcbMaxValueLen As LongPtr, ByVal lpcbSecurityDescriptor As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regflushkey
    Private Declare PtrSafe Function RegFlushKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
    'LSTATUS RegCloseKey([in] HKEY hKey);
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyw
    Private Declare PtrSafe Function RegDeleteKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyexw
    'LSTATUS RegDeleteKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, [in] REGSAM samDesired, DWORD Reserved);
    Private Declare PtrSafe Function RegDeleteKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal samDesired As Long, ByVal Reserved As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumvaluew
    'LSTATUS RegEnumValueW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpValueName, [in, out] LPDWORD lpcchValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare PtrSafe Function RegEnumValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As LongPtr, ByVal lpcchValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryvalueexw
    'LSTATUS RegQueryValueExW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare PtrSafe Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvaluew
    Private Declare PtrSafe Function RegSetValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvalueexw
    Private Declare PtrSafe Function RegSetValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletevaluew
    Private Declare PtrSafe Function RegDeleteValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumkeyexw
    'LSTATUS RegEnumKeyExW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpName, [in, out] LPDWORD lpcchName, LPDWORD lpReserved,
    '                      [in, out] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, [out, optional] PFILETIME lpftLastWriteTime);
    Private Declare PtrSafe Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As LongPtr, ByVal lpcchName As LongPtr, ByVal lpReserved As LongPtr, _
                                                                   ByVal lpClass As LongPtr, ByVal lpcchClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
#Else
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyw
    Private Declare Function RegCreateKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regcreatekeyexw
    'LSTATUS RegCreateKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, DWORD Reserved, [in, optional] LPWSTR lpClass, [in] DWORD dwOptions, [in] REGSAM samDesired, [in, optional] const LPSECURITY_ATTRIBUTES lpSecurityAttributes, [out] PHKEY phkResult, [out, optional] LPDWORD lpdwDisposition);
    Private Declare Function RegCreateKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal Reserved As Long, ByVal lpClass As LongPtr, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As LongPtr, ByVal phkResult As LongPtr, ByVal lpdwDisposition As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regopenkeyexw
    'LSTATUS RegOpenKeyExW([in] HKEY hKey, [in, optional] LPCWSTR lpSubKey, [in] DWORD ulOptions, [in] REGSAM samDesired, [out] PHKEY phkResult);
    Private Declare Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal ulOptions As Long, ByVal samDesired As Long, ByVal phkResult As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryinfokeyw
    'LSTATUS RegQueryInfoKeyW([in] HKEY hKey, [out, optional] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, LPDWORD lpReserved, [out, optional] LPDWORD lpcSubKeys,
    '                         [out, optional] LPDWORD lpcbMaxSubKeyLen, [out, optional] LPDWORD lpcbMaxClassLen, [out, optional] LPDWORD lpcValues, [out, optional] LPDWORD lpcbMaxValueNameLen,
    '                         [out, optional] LPDWORD lpcbMaxValueLen,  [out, optional] LPDWORD lpcbSecurityDescriptor, [out, optional] PFILETIME lpftLastWriteTime
    Private Declare Function RegQueryInfoKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpClass As LongPtr, ByVal lpcbClass As LongPtr, ByVal lpReserved As LongPtr, ByVal lpcSubKeys As LongPtr, _
                                                              ByVal lpcbMaxSubKeyLen As LongPtr, ByVal lpcbMaxClassLen As LongPtr, ByVal lpcValues As LongPtr, ByVal lpcbMaxValueNameLen As LongPtr, _
                                                              ByVal lpcbMaxValueLen As LongPtr, ByVal lpcbSecurityDescriptor As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regflushkey
    Private Declare Function RegFlushKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regclosekey
    'LSTATUS RegCloseKey([in] HKEY hKey);
    Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyw
    Private Declare Function RegDeleteKeyW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletekeyexw
    'LSTATUS RegDeleteKeyExW([in] HKEY hKey, [in] LPCWSTR lpSubKey, [in] REGSAM samDesired, DWORD Reserved);
    Private Declare Function RegDeleteKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal samDesired As Long, ByVal Reserved As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumvaluew
    'LSTATUS RegEnumValueW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpValueName, [in, out] LPDWORD lpcchValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare Function RegEnumValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpValueName As LongPtr, ByVal lpcchValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regqueryvalueexw
    'LSTATUS RegQueryValueExW([in] HKEY hKey, [in, optional] LPCWSTR lpValueName, LPDWORD lpReserved, [out, optional] LPDWORD lpType, [out, optional] LPBYTE lpData, [in, out, optional] LPDWORD lpcbData);
    Private Declare Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpType As LongPtr, ByVal lpData As LongPtr, ByVal lpcbData As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvaluew
    Private Declare Function RegSetValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpSubKey As LongPtr, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regsetvalueexw
    Private Declare Function RegSetValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As LongPtr, ByVal cbData As Long) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regdeletevaluew
    Private Declare Function RegDeleteValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal lpValueName As LongPtr) As Long
    
    'https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumkeyexw
    'LSTATUS RegEnumKeyExW([in] HKEY hKey, [in] DWORD dwIndex, [out] LPWSTR lpName, [in, out] LPDWORD lpcchName, LPDWORD lpReserved,
    '                      [in, out] LPWSTR lpClass, [in, out, optional] LPDWORD lpcchClass, [out, optional] PFILETIME lpftLastWriteTime);
    Private Declare Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As LongPtr, ByVal lpcchName As LongPtr, ByVal lpReserved As LongPtr, _
                                                           ByVal lpClass As LongPtr, ByVal lpcchClass As LongPtr, ByVal lpftLastWriteTime As LongPtr) As Long
    
#End If

' v ############################## v '     ACCESS TYPES  (winnt.h)     ' v ############################## v '
Private Const KEY_QUERY_VALUE           As Long = &H1      ' Required to query the values of a registry key.
Private Const KEY_SET_VALUE             As Long = &H2&     ' Required to create, delete, or set a registry value.
Private Const KEY_CREATE_SUB_KEY        As Long = &H4&     ' Required to create a subkey of a registry key.
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8&     ' Required to enumerate the subkeys of a registry key.
Private Const KEY_NOTIFY                As Long = &H10&    ' Required to request change notifications for a registry key or for subkeys of a registry key.
Private Const KEY_CREATE_LINK           As Long = &H20&    ' Reserved for system use.

Private Const KEY_WOW64_64KEY           As Long = &H100&   ' Indicates that an application on 64-bit Windows should operate on the 64-bit registry view. This flag is ignored by 32-bit Windows. For more information, see Accessing an Alternate Registry View.
                                                           ' This flag must be combined using the OR operator with the other flags in this table that either query or access registry values.
                                                           ' Windows 2000: This flag is not supported.
Private Const KEY_WOW64_32KEY           As Long = &H200&   ' Indicates that an application on 64-bit Windows should operate on the 32-bit registry view. This flag is ignored by 32-bit Windows. For more information, see Accessing an Alternate Registry View.
                                                           ' This flag must be combined using the OR operator with the other flags in this table that either query or access registry values.
                                                           ' Windows 2000: This flag is not supported.

Private Const RIGHT_DELETE              As Long = &H10000   ' The right to delete the object.
Private Const RIGHT_READ_CONTROL        As Long = &H20000   ' The right to read the information in the object's security descriptor, not including the information in the system access control list (SACL).
Private Const RIGHT_WRITE_DAC           As Long = &H40000   ' The right to modify the discretionary access control list (DACL) in the object's security descriptor.
Private Const RIGHT_WRITE_OWNER         As Long = &H80000   ' The right to change the owner in the object's security descriptor.
Private Const RIGHT_SYNCHRONIZE         As Long = &H100000  ' The right to use the object for synchronization. This enables a thread to wait until the object is in the signaled state. Some object types do not support this access right.

Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000   ' Combines DELETE, READ_CONTROL, WRITE_DAC, and WRITE_OWNER access.

Private Const STANDARD_RIGHTS_READ      As Long = (RIGHT_READ_CONTROL) ' Currently defined to equal READ_CONTROL.
Private Const STANDARD_RIGHTS_WRITE     As Long = (RIGHT_READ_CONTROL) ' Currently defined to equal READ_CONTROL.
Private Const STANDARD_RIGHTS_EXECUTE   As Long = (RIGHT_READ_CONTROL) ' Currently defined to equal READ_CONTROL.
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000             ' Combines DELETE, READ_CONTROL, WRITE_DAC, WRITE_OWNER, and SYNCHRONIZE access.
Private Const SPECIFIC_RIGHTS_ALL       As Long = &HFFFF&

Private Const KEY_ALL_ACCESS            As Long = &HF003F  ' Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.
Private Const KEY_READ                  As Long = &H20019  ' Combines the STANDARD_RIGHTS_READ, KEY_QUERY_VALUE, KEY_ENUMERATE_SUB_KEYS, and KEY_NOTIFY values.
Private Const KEY_EXECUTE               As Long = &H20019  ' Equivalent to KEY_READ.
Private Const KEY_WRITE                 As Long = &H20006  ' Combines the STANDARD_RIGHTS_WRITE, KEY_SET_VALUE, and KEY_CREATE_SUB_KEY access rights.
' ^ ############################## ^ '     ACCESS TYPES  (winnt.h)       ' ^ ############################## ^ '


Private Const HKEY_CLASSES_ROOT     As Long = &H80000000
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE    As Long = &H80000002
Private Const HKEY_USERS            As Long = &H80000003
Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004
Private Const HKEY_CURRENT_CONFIG   As Long = &H80000005
Private Const HKEY_DYN_DATA         As Long = &H80000006

Private Const STATE_DIRTY           As Long = 1
Private Const STATE_SYSTEMKEY       As Long = 2
Private Const STATE_WRITEACCESS     As Long = 4
Private Const STATE_PERF_DATA       As Long = 8

Private Const MaxKeyLength          As Long = 255
Private Const MaxValueLength        As Long = 16383

Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = 512
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = 4096
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = 8192

Public Enum RegistryHive
    ClassesRoot = &H80000000 'int.MinValue, '/// <summary>Stellt den Basisschlüssel HKEY_CLASSES_ROOT auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    CurrentUser         '/// <summary>Stellt den Basisschlüssel HKEY_CURRENT_USER auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    LocalMachine        '/// <summary>Stellt den Basisschlüssel HKEY_LOCAL_MACHINE auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    Users               '/// <summary>Stellt den Basisschlüssel HKEY_USERS auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    PerformanceData     '/// <summary>Stellt den Basisschlüssel HKEY_PERFORMANCE_DATA auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    CurrentConfig       '/// <summary>Stellt den Basisschlüssel HKEY_CURRENT_CONFIG auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
    DynData             '/// <summary>Stellt den Basisschlüssel HKEY_DYN_DATA auf einem anderen Computer dar.Dieser Wert kann an die <see cref="M:Microsoft.Win32.RegistryKey.OpenRemoteBaseKey(Microsoft.Win32.RegistryHive,System.String)" />-Methode übergeben werden, um diesen Knoten remote zu öffnen.</summary>
End Enum

Private Enum RegistryInternalCheck
    CheckSubKeyWritePermission
    CheckSubKeyReadPermission
    CheckSubKeyCreatePermission
    CheckSubTreeReadPermission
    CheckSubTreeWritePermission
    CheckSubTreeReadWritePermission
    CheckValueWritePermission
    CheckValueCreatePermission
    CheckValueReadPermission
    CheckKeyReadPermission
    CheckSubTreePermission
    CheckOpenSubKeyWithWritablePermission
    CheckOpenSubKeyPermission
End Enum

Public Enum RegistryKeyPermissionCheck
    rkpcUndefined = -1
    Default             '/// <summary>Der Registrierungsschlüssel erbt den Modus von seinem übergeordneten Element.Beim Versuch, auf Unterschlüssel oder Werte zuzugreifen, werden Sicherheitsüberprüfungen durchgeführt, sofern das übergeordnete Element nicht mit dem <see cref="F:Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree" />-Modus oder dem <see cref="F:Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree" />-Modus geöffnet wurde.</summary>
    ReadSubTree         '/// <summary>Sicherheitsüberprüfungen werden nicht durchgeführt, wenn auf Unterschlüssel oder Werte zugegriffen wird.Beim Versuch, den aktuellen Schlüssel zu öffnen, wird eine Sicherheitsüberprüfung durchgeführt, sofern das übergeordnete Element nicht mit <see cref="F:Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree" /> oder <see cref="F:Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree" /> geöffnet wurde.</summary>
    ReadWriteSubTree    '/// <summary>Sicherheitsüberprüfungen werden nicht durchgeführt, wenn auf Unterschlüssel oder Werte zugegriffen wird.Beim Versuch, den aktuellen Schlüssel zu öffnen, wird eine Sicherheitsüberprüfung durchgeführt, sofern das übergeordnete Element nicht mit <see cref="F:Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree" /> geöffnet wurde.</summary>
End Enum

Public Enum RegistryOptions
    roUndefined = -1
    None     '= 0x0     '/// <summary>Ein nicht flüchtiger Schlüssel.Dies ist die Standardeinstellung.</summary>
    Volatile '= 0x1     '/// <summary>Ein temporärer Schlüssel.Die Informationen werden im Arbeitsspeicher gespeichert und nicht beibehalten, wenn der entsprechende Registrierungshive entladen wird.</summary>
End Enum

Public Enum RegistryValueKind
    rvkNone = -1        '/// <summary>Kein Datentyp.</summary>
    rvkUnknown = 0      '/// <summary>Ein nicht unterstützter Registrierungsdatentyp.Zum Beispiel wird der Microsoft Win32-API-Registrierungsdatentyp REG_RESOURCE_LIST nicht unterstützt.Geben Sie mit diesem Wert an, dass die <see cref="M:Microsoft.Win32.RegistryKey.SetValue(System.String,System.Object)" />-Methode den entsprechenden Registrierungsdatentyp bei der Speicherung eines Name-/Wert-Paares bestimmen soll.</summary>
    rvkString = 1       '/// <summary>Eine auf NULL endende Zeichenfolge.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_SZ.</summary>
    rvkExpandString = 2 '/// <summary>Eine auf NULL endende Zeichenfolge, die nicht erweiterte Verweise auf Umgebungsvariablen (z. B. %PATH%) enthält, die bei Abruf des Werts erweitert werden.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_EXPAND_SZ.</summary>
    rvkBinary = 3       '/// <summary>Binärdaten in beliebiger Form.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_BINARY.</summary>
    rvkDWord = 4        '/// <summary>Eine 32-Bit-Binärzahl.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_DWORD.</summary>
    rvkMultiString = 7  '/// <summary>Ein Array von auf NULL endenden Zeichenfolgen, das auf zwei NULL-Zeichen endet.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_MULTI_SZ.</summary>
    rvkQWord = 11       '/// <summary>Eine 64-Bit-Binärzahl.Dieser Wert entspricht dem Win32-API-Registrierungsdatentyp REG_QWORD.</summary>
End Enum

Public Enum RegistryValueOptions
    None '= 0x0         '/// <summary>Es ist kein optionales Verhalten angegeben.</summary>
    DoNotExpandEnvironmentNames '= 0x1'/// <summary>Ein Wert des Typs <see cref="F:Microsoft.Win32.RegistryValueKind.ExpandString" /> wird abgerufen, ohne die eingebetteten Umgebungsvariablen zu erweitern. </summary>
End Enum

Public Enum RegistryView
    Default = 0         '/// <summary>Die Standardansicht.</summary>
    Registry64 = &H100  '/// <summary>Die 64-Bit-Ansicht.</summary>
    Registry32 = &H200  '/// <summary>Die 32-Bit-Ansicht.</summary>
End Enum

Private Type SECURITY_ATTRIBUTES
    nLength              As Long    ' 4
    lpSecurityDescriptor As Long    ' 4
    bInheritHandle       As Long    ' 4
End Type                      ' Sum: 12

Private Type ACL
    AclRevision         As Byte    ' 1
    Sbz1                As Byte    ' 1
    AclSize             As Integer ' 2
    AceCount            As Integer ' 2
    Sbz2                As Integer ' 2
End Type                      ' Sum: 8

Private Type SECURITY_DESCRIPTOR
    Revision             As Byte    ' 1
    Sbz1                 As Byte    ' 1+2
    Control              As Long    ' 4
    Owner                As Long    ' 4
    Group                As Long    ' 4
    Sacl                 As ACL     ' 8
    Dacl                 As ACL     ' 8
End Type                      ' Sum: 32

Private hkeyNames() As String ' = Array("HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_USERS", "HKEY_PERFORMANCE_DATA", "HKEY_CURRENT_CONFIG", "HKEY_DYN_DATA")

''<SecurityCritical>
Private m_hkey      As LongPtr ' SafeRegistryHandle
Private m_state     As Long
Private m_keyName   As String
Private m_remoteKey As Boolean
Private m_checkMode As RegistryKeyPermissionCheck
Private m_regView   As RegistryView

Private Sub Class_Initialize()
    hkeyNames = SArray("HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER", "HKEY_LOCAL_MACHINE", "HKEY_USERS", "HKEY_PERFORMANCE_DATA", "HKEY_CURRENT_CONFIG", "HKEY_DYN_DATA")
End Sub

Private Sub Class_Terminate()
    Me.Dispose
End Sub

''' <summary>Ruft die Anzahl der Unterschlüssel des aktuellen Schlüssels ab.</summary>
''' <returns>Die Anzahl der Unterschlüssel des aktuellen Schlüssels.</returns>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer besitzt keine Leseberechtigung für den Schlüssel. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, möglicherweise wurde der aktuelle Schlüssel gelöscht.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
Public Property Get SubKeyCount() As Long
    '<SecuritySafeCritical>
    CheckPermission RegistryInternalCheck.CheckKeyReadPermission, vbNullString, False, RegistryKeyPermissionCheck.[Default]
    SubKeyCount = InternalSubKeyCount()
End Property

''' <summary>Ruft die Ansicht ab, mit der der Registrierungsschlüssel erstellt wurde. </summary>
''' <returns>Die Ansicht, mit der der Registrierungsschlüssel erstellt wurde.– oder –<seecref="F:Microsoft.Win32.RegistryView.Default"/>, wenn keine Ansicht verwendet wurde.</returns>
'<ComVisible(False)>
Public Property Get View() As RegistryView
    '<SecuritySafeCritical>
    EnsureNotDisposed
    View = m_regView
End Property

''' <summary>Ruft ein <seecref="T:Microsoft.Win32.SafeHandles.SafeRegistryHandle"/>-Objekt ab, das den Registrierungsschlüssel darstellt, der vom aktuellen <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt gekapselt wird.</summary>
''' <returns>Das Handle für den Registrierungsschlüssel.</returns>
''' <exceptioncref="T:System.ObjectDisposedException">Der Registrierungsschlüssel ist geschlossen.Auf geschlossene Schlüssel kann nicht zugegriffen werden.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, z. B. Löschen des aktuellen Schlüssels.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Lesen des Schlüssels.</exception>
'<ComVisible(False)>
Public Property Get Handle() As LongPtr
    '<SecurityCritical>
    '<SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.UnmanagedCode)>
    EnsureNotDisposed
    Dim errorCode As Long: errorCode = 6
    Dim hkResult As LongPtr '= Nothing
    If IsSystemKey Then
        Dim hKey As LongPtr '= CType(0, IntPtr)
        Select Case m_keyName
        Case "HKEY_CLASSES_ROOT":     hKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER":     hKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE":    hKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS":            hKey = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA": hKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG":   hKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA":         hKey = HKEY_DYN_DATA
        Case Else: Win32Error errorCode, Nothing
        End Select
        
        errorCode = RegOpenKeyExW(hKey, Nothing, 0, GetRegistryKeyAccess(IsWritable) Or CLng(regView), VarPtr(hkResult))
        If errorCode = 0 Then
            If Not hkResult.IsInvalid Then Handle = hkResult
        End If
        Win32Error errorCode, Nothing
        'Throw New IOException(Win32Native.GetMessage(errorCode), errorCode)
    End If
    Handle = hKey
End Property

''' <summary>Ruft die Anzahl der Werte im Schlüssel ab.</summary>
''' <returns>Die Anzahl der Name-Wert-Paare im Schlüssel.</returns>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer besitzt keine Leseberechtigung für den Schlüssel. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, möglicherweise wurde der aktuelle Schlüssel gelöscht.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
Public Property Get ValueCount() As Long
    '<SecuritySafeCritical>
    CheckPermission RegistryInternalCheck.CheckKeyReadPermission, vbNullString, False, RegistryKeyPermissionCheck.[Default]
    ValueCount = InternalValueCount
End Property

''' <summary>Ruft den Namen des Schlüssels ab.</summary>
''' <returns>Der absolute (vollständige) Name des Schlüssels.</returns>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
Public Property Get Name() As String
    '<SecuritySafeCritical>
    EnsureNotDisposed
    Name = m_keyName
End Property

'<SecurityCritical>
Private Sub NewPrivate(ByVal hKey As LongPtr, ByVal Writable As Boolean, ByVal View As RegistryView)
    NewPrivate2 hKey, Writable, False, False, False, View
End Sub

'<SecurityCritical>
Private Sub NewPrivate2(ByVal hKey As LongPtr, ByVal Writable As Boolean, ByVal systemkey As Boolean, ByVal remoteKey As Boolean, ByVal isPerfData As Boolean, ByVal View As RegistryView)
    m_hkey = hKey
    m_keyName = ""
    m_remoteKey = remoteKey
    regView = View
    If systemkey Then
        m_state = m_state Or STATE_SYSTEMKEY
    End If
    If Writable Then
        m_state = m_state Or STATE_WRITEACCESS
    End If
    If isPerfData Then
        m_state = m_state Or STATE_PERF_DATA
    End If
    ValidateKeyView View
End Sub

''' <summary>Schließt den Schlüssel und schreibt diesen auf den Datenträger weg, sofern der Inhalt verändert wurde.</summary>
Public Sub CClose()
    Dispose
End Sub

'<SecuritySafeCritical>
Private Sub DDispose(ByVal disposing As Boolean)
    If m_hkey = 0 Then
        Exit Sub
    End If
    If Not IsSystemKey Then
        'Try
            RegCloseKey m_hkey '.Dispose
        'Catch __unusedIOException1__ As IOException
        'Finally
            m_hkey = 0 'Nothing
        'End Try
    ElseIf disposing Then
        If IsPerfDataKey Then
            RegCloseKey HKEY_PERFORMANCE_DATA
        End If
    End If
End Sub

''' <summary>Schreibt alle Attribute des angegebenen geöffneten Registrierungsschlüssels in die Registrierung.</summary>
'<SecuritySafeCritical>
Public Sub Flush()
    If m_hkey <> 0 Then
        If IsDirty() Then RegFlushKey m_hkey
    End If
End Sub

''' <summary>Gibt alle von der aktuellen Instanz der <seecref="T:Microsoft.Win32.RegistryKey"/>-Klasse verwendeten Ressourcen frei.</summary>
Public Sub Dispose()
    DDispose True
End Sub

Public Function CreateSubKey(ByVal subkey As String, Optional ByVal Writable, _
                                                     Optional ByVal PermissionCheck As RegistryKeyPermissionCheck = RegistryKeyPermissionCheck.rkpcUndefined, _
                                                     Optional ByVal Options As RegistryOptions = RegistryOptions.None, _
                                                     Optional ByVal Reg_Security) As RegistryKey
    If IsMissing(Writable) Then
        If PermissionCheck = RegistryKeyPermissionCheck.rkpcUndefined Then
            PermissionCheck = m_checkMode
        End If
    Else
        PermissionCheck = IIf(CBool(Writable), RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryKeyPermissionCheck.ReadSubTree)
    End If
    Set CreateSubKey = CreateSubKeyInternal(subkey, PermissionCheck, Reg_Security, Options)
End Function

''' <summary>Erstellt einen neuen Unterschlüssel oder öffnet einen vorhandenen Unterschlüssel für Schreibzugriff.  </summary>
''' <returns>Der neu erstellte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.Wenn eine Zeichenfolge der Länge 0 (null) für <paramrefname="subkey"/> angegeben wird, wird das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt zurückgegeben.</returns>
''' <paramname="subkey">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Öffnen des Registrierungsschlüssels. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, für den die Methode aufgerufen wird, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/> kann nicht bearbeitet werden. Möglicherweise wurde der Schlüssel schreibgeschützt geöffnet, oder der Benutzer verfügt nicht über die erforderlichen Zugriffsrechte. </exception>
''' <exceptioncref="T:System.IO.IOException">Die Schachtelungsebene übersteigt 510.- oder -Ein Systemfehler ist aufgetreten. Möglicherweise wurde der Schlüssel gelöscht, oder es wurde versucht, einen Schlüssel im <seecref="F:Microsoft.Win32.Registry.LocalMachine"/>-Stamm zu erstellen.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
'Public Function CreateSubKey1(ByVal subkey As String) As RegistryKey
'    Set CreateSubKey = CreateSubKey2(subkey, m_checkMode)
'End Function

''' <summary>Erstellt einen neuen Unterschlüssel oder öffnet einen vorhandenen Unterschlüssel für Schreibzugriff unter Verwendung der angegebenen Berechtigungsprüfungsoption. </summary>
''' <returns>Der neu erstellte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.Wenn eine Zeichenfolge der Länge 0 (null) für <paramrefname="subkey"/> angegeben wird, wird das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt zurückgegeben.</returns>
''' <paramname="subkey">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Öffnen des Registrierungsschlüssels. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="permissionCheck"/> enthält einen ungültigen Wert.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, für den die Methode aufgerufen wird, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/> kann nicht bearbeitet werden. Möglicherweise wurde der Schlüssel schreibgeschützt geöffnet, oder der Benutzer verfügt nicht über die erforderlichen Zugriffsrechte. </exception>
''' <exceptioncref="T:System.IO.IOException">Die Schachtelungsebene übersteigt 510.- oder -Ein Systemfehler ist aufgetreten. Möglicherweise wurde der Schlüssel gelöscht, oder es wurde versucht, einen Schlüssel im <seecref="F:Microsoft.Win32.Registry.LocalMachine"/>-Stamm zu erstellen.</exception>
'<ComVisible(False)>
'Public Function CreateSubKey2(ByVal subkey As String, ByVal PermissionCheck As RegistryKeyPermissionCheck) As RegistryKey
'    Set CreateSubKey2 = CreateSubKeyInternal(subkey, PermissionCheck, Nothing, RegistryOptions.None)
'End Function

''' <summary>Erstellt oder öffnet einen Unterschlüssel für Schreibzugriff unter Verwendung der angegebenen Berechtigungsprüfungs- und Registrierungsoptionen. </summary>
''' <returns>Der neu erstellte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.</returns>
''' <paramname="subkey">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels. </param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <paramname="options">Die zu verwendende Registrierungsoption, z. B. zum Erstellen eines temporären Schlüssels. </param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> ist null.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt, auf das zugegriffen werden soll, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt kann nicht bearbeitet werden. Möglicherweise wurde der Schlüssel schreibgeschützt geöffnet, oder der Benutzer verfügt nicht über die erforderlichen Zugriffsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Die Schachtelungsebene übersteigt 510.- oder -Ein Systemfehler ist aufgetreten. Möglicherweise wurde der Schlüssel gelöscht, oder es wurde versucht, einen Schlüssel im <seecref="F:Microsoft.Win32.Registry.LocalMachine"/>-Stamm zu erstellen. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Öffnen des Registrierungsschlüssels.</exception>
'<ComVisible(False)>
'Public Function CreateSubKey3(ByVal subkey As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal Options As RegistryOptions) As RegistryKey
'    Set CreateSubKey3 = CreateSubKeyInternal(subkey, PermissionCheck, Nothing, Options)
'End Function

'<ComVisible(False)>
'Public Function CreateSubKey4(ByVal subkey As String, ByVal Writable As Boolean) As RegistryKey
'    Set CreateSubKey4 = CreateSubKeyInternal(subkey, IIf(Not Writable, RegistryKeyPermissionCheck.ReadSubTree, RegistryKeyPermissionCheck.ReadWriteSubTree), Nothing, RegistryOptions.None)
'End Function

'<ComVisible(False)>
'Public Function CreateSubKey5(ByVal subkey As String, ByVal Writable As Boolean, ByVal Options As RegistryOptions) As RegistryKey
'    Set CreateSubKey5 = CreateSubKeyInternal(subkey, IIf(Not Writable, RegistryKeyPermissionCheck.ReadSubTree, RegistryKeyPermissionCheck.ReadWriteSubTree), Nothing, Options)
'End Function

''' <summary>Erstellt einen neuen Unterschlüssel oder öffnet einen vorhandenen Unterschlüssel für Schreibzugriff unter Verwendung der angegebenen Berechtigungsprüfungsoption und der Registrierungssicherheit. </summary>
''' <returns>Der neu erstellte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.Wenn eine Zeichenfolge der Länge 0 (null) für <paramrefname="subkey"/> angegeben wird, wird das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt zurückgegeben.</returns>
''' <paramname="subkey">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <paramname="registrySecurity">Die Zugriffssteuerungssicherheit für den neuen Schlüssel.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Öffnen des Registrierungsschlüssels. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="permissionCheck"/> enthält einen ungültigen Wert.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, für den die Methode aufgerufen wird, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/> kann nicht bearbeitet werden. Möglicherweise wurde der Schlüssel schreibgeschützt geöffnet, oder der Benutzer verfügt nicht über die erforderlichen Zugriffsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Die Schachtelungsebene übersteigt 510.- oder -Ein Systemfehler ist aufgetreten. Möglicherweise wurde der Schlüssel gelöscht, oder es wurde versucht, einen Schlüssel im <seecref="F:Microsoft.Win32.Registry.LocalMachine"/>-Stamm zu erstellen.</exception>
'<ComVisible(False)>
'Public Function CreateSubKey6(ByVal subkey As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal registrySecurity As registrySecurity) As RegistryKey
'    Set CreateSubKey6 = CreateSubKeyInternal(subkey, PermissionCheck, registrySecurity, RegistryOptions.None)
'End Function

''' <summary>Erstellt oder öffnet einen Unterschlüssel für Schreibzugriff unter Verwendung der angegebenen Berechtigungsprüfungsoption, Registrierungsoption und Registrierungssicherheit.</summary>
''' <returns>Der neu erstellte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.  </returns>
''' <paramname="subkey">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.</param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <paramname="registryOptions">Die zu verwendende Registrierungsoption.</param>
''' <paramname="registrySecurity">Die Zugriffssteuerungssicherheit für den neuen Unterschlüssel. </param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> ist null.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt ist geschlossen.Auf geschlossene Schlüssel kann nicht zugegriffen werden.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt kann nicht bearbeitet werden. Möglicherweise wurde der Schlüssel schreibgeschützt geöffnet, oder der Benutzer verfügt nicht über die erforderlichen Zugriffsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Die Schachtelungsebene übersteigt 510.- oder -Ein Systemfehler ist aufgetreten. Möglicherweise wurde der Schlüssel gelöscht, oder es wurde versucht, einen Schlüssel im <seecref="F:Microsoft.Win32.Registry.LocalMachine"/>-Stamm zu erstellen. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Öffnen des Registrierungsschlüssels.</exception>
'<ComVisible(False)>
'Public Function CreateSubKey7(ByVal subkey As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal RegistryOptions As RegistryOptions, ByVal registrySecurity As registrySecurity) As RegistryKey
'    Set CreateSubKey7 = CreateSubKeyInternal(subkey, PermissionCheck, registrySecurity, RegistryOptions)
'End Function

Private Function CreateSubKeyInternal(ByVal subkey As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal registrySecurityObj, ByVal RegistryOptions As RegistryOptions) As RegistryKey
    ValidateKeyOptions RegistryOptions
    ValidateKeyName subkey
    ValidateKeyMode PermissionCheck
    EnsureWriteable
    subkey = FixupName(subkey)
    If Not m_remoteKey Then
        Dim aRegistryKey As RegistryKey: Set aRegistryKey = InternalOpenSubKey(subkey, PermissionCheck <> RegistryKeyPermissionCheck.ReadSubTree)
        If Not aRegistryKey Is Nothing Then
            CheckPermission RegistryInternalCheck.CheckSubKeyWritePermission, subkey, False, RegistryKeyPermissionCheck.[Default]
            CheckPermission RegistryInternalCheck.CheckSubTreePermission, subkey, False, PermissionCheck
            RegistryKey.checkMode = PermissionCheck
            Set CreateSubKeyInternal = aRegistryKey
        End If
    End If
    CheckPermission RegistryInternalCheck.CheckSubKeyCreatePermission, subkey, False, RegistryKeyPermissionCheck.[Default]
    Dim sec_attribs As SECURITY_ATTRIBUTES '= Nothing
    Dim registrySecurity As registrySecurity: registrySecurity = CType(registrySecurityObj, registrySecurity)
    If Not registrySecurity Is Nothing Then
        'sec_attribs = New Win32Native.SECURITY_ATTRIBUTES()
        sec_attribs.nLength = LenB(sec_attribs)  ' Marshal.SizeOf(SECURITY_ATTRIBUTES)
        Dim securityDescriptorBinaryForm() As Byte: securityDescriptorBinaryForm = registrySecurity.GetSecurityDescriptorBinaryForm()
''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.TypeSyntax'.
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 471
'''    at ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\MethodBodyExecutableStatementVisitor.cs:line 52
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingMethodBodyVisitor.cs:line 16
'''
''' Input:
             byte* ptr = stackalloc byte[(int)checked(unchecked((System.nuint)(uint)securityDescriptorBinaryForm.Length) * (System.nuint)1u)];
'''
'''
        Buffer.Memcpy(ptr, 0, securityDescriptorBinaryForm, 0, securityDescriptorBinaryForm.Length)
        sec_attribs.pSecurityDescriptor = Ptr
    End If
    Dim lpdwDisposition As Long '= 0
    Dim hkResult As LongPtr 'SafeRegistryHandle = Nothing
    Dim num As Long: num = RegCreateKeyExW(hKey, StrPtr(subkey), 0, Nothing, CLng(RegistryOptions), GetRegistryKeyAccess(Not PermissionCheck = RegistryKeyPermissionCheck.ReadSubTree) Or CLng(regView), SECURITY_ATTRIBUTES, hkResult, lpdwDisposition)
    If num = 0 Then
        If Not hkResult.IsInvalid Then
            Dim RegistryKey2 As RegistryKey: set RegistryKey2 = New RegistryKey: RegistryKey2.New_ hkResult, permissionCheck IsNot RegistryKeyPermissionCheck.ReadSubTree, False, remoteKey, False, regView
            CheckPermission(RegistryInternalCheck.CheckSubTreePermission, subkey, subKeyWritable:=False, permissionCheck)
            RegistryKey2.checkMode = PermissionCheck
            If subkey.Length = 0 Then
                RegistryKey2.keyName = keyName
            Else
                RegistryKey2.keyName = keyName & "\" & subkey
            End If
            Set CreateSubKeyInternal = RegistryKey2
            Exit Function
        End If
    End If
    If num <> 0 Then
        Win32Error(num, keyName & "\" & subkey)
    End If
    Return Nothing
End Function

'''     private unsafe RegistryKey CreateSubKeyInternal(string subkey, RegistryKeyPermissionCheck permissionCheck, object registrySecurityObj, RegistryOptions registryOptions)
'''     {
'''         RegistryKey.ValidateKeyOptions(registryOptions);
'''         RegistryKey.ValidateKeyName(subkey);
'''         RegistryKey.ValidateKeyMode(permissionCheck);
'''         this.EnsureWriteable();
'''         subkey = RegistryKey.FixupName(subkey);
'''         if (!this.remoteKey)
'''         {
'''             RegistryKey RegistryKey = this.InternalOpenSubKey(subkey, permissionCheck != RegistryKeyPermissionCheck.ReadSubTree);
'''             if (RegistryKey != null)
'''             {
'''                 CheckPermission(RegistryKey.RegistryInternalCheck.CheckSubKeyWritePermission, subkey, subKeyWritable: false, RegistryKeyPermissionCheck.Default);
'''                 this.CheckPermission(RegistryKey.RegistryInternalCheck.CheckSubTreePermission, subkey, subKeyWritable: false, permissionCheck);
'''                 RegistryKey.checkMode = permissionCheck;
'''                 return RegistryKey;
'''             }
'''         }
'''         CheckPermission(RegistryKey.RegistryInternalCheck.CheckSubKeyCreatePermission, subkey, subKeyWritable: false, RegistryKeyPermissionCheck.Default);
'''         Win32Native.SECURITY_ATTRIBUTES sECURITY_ATTRIBUTES = null;
'''         RegistrySecurity registrySecurity = (RegistrySecurity)registrySecurityObj;
'''         if (registrySecurity != null)
'''         {
'''             sECURITY_ATTRIBUTES = new Win32Native.SECURITY_ATTRIBUTES();
'''             sECURITY_ATTRIBUTES.nLength = Marshal.SizeOf(sECURITY_ATTRIBUTES);
'''             byte[] securityDescriptorBinaryForm = registrySecurity.GetSecurityDescriptorBinaryForm();
'''             byte* ptr = stackalloc byte[(int)checked(unchecked((System.nuint)(uint)securityDescriptorBinaryForm.Length) * (System.nuint)1u)];
'''             Buffer.Memcpy(ptr, 0, securityDescriptorBinaryForm, 0, securityDescriptorBinaryForm.Length);
'''             sECURITY_ATTRIBUTES.pSecurityDescriptor = ptr;
'''         }
'''         int lpdwDisposition = 0;
'''         SafeRegistryHandle hkResult = null;
'''         int num = Win32Native.RegCreateKeyEx(this.hkey, subkey, 0, null, (int)registryOptions, RegistryKey.GetRegistryKeyAccess(permissionCheck != RegistryKeyPermissionCheck.ReadSubTree) | (int)this.regView, sECURITY_ATTRIBUTES, out hkResult, out lpdwDisposition);
'''         if (num == 0 && !hkResult.IsInvalid)
'''         {
'''             RegistryKey RegistryKey2 = new RegistryKey(hkResult, permissionCheck != RegistryKeyPermissionCheck.ReadSubTree, systemkey: false, this.remoteKey, isPerfData: false, this.regView);
'''             this.CheckPermission(RegistryKey.RegistryInternalCheck.CheckSubTreePermission, subkey, subKeyWritable: false, permissionCheck);
'''             RegistryKey2.checkMode = permissionCheck;
'''             if (subkey.Length == 0)
'''             {
'''                 RegistryKey2.keyName = this.keyName;
'''             }
'''             else
'''             {
'''                 RegistryKey2.keyName = this.keyName + "\\" + subkey;
'''             }
'''             return RegistryKey2;
'''         }
'''         if (num != 0)
'''         {
'''             this.Win32Error(num, this.keyName + "\\" + subkey);
'''         }
'''         return null;
'''     }
'''
'''

''' <summary>Löscht den angegebenen Unterschlüssel. </summary>
''' <paramname="subkey">Der Name des zu löschenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <exceptioncref="T:System.InvalidOperationException">Der <paramrefname="subkey"/> besitzt untergeordnete Unterschlüssel. </exception>
''' <exceptioncref="T:System.ArgumentException">Der <paramrefname="subkey"/>-Parameter gibt keinen gültigen Registrierungsschlüssel an. </exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Schlüssel zu löschen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
Public Sub DeleteSubKey(ByVal subkey As String)
    DeleteSubKey2 subkey, True
End Sub

''' <summary>Löscht den angegebenen Unterschlüssel und gibt an, ob eine Ausnahme ausgelöst wird, wenn der Unterschlüssel nicht gefunden wird. </summary>
''' <paramname="subkey">Der Name des zu löschenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="throwOnMissingSubKey">Gibt an, ob eine Ausnahme ausgelöst werden soll, wenn der angegebene Unterschlüssel nicht gefunden werden kann.Wenn dieses Argument true und der angegebene Unterschlüssel nicht vorhanden ist, wird eine Ausnahme ausgelöst.Wenn dieses Argument false und der angegebene Unterschlüssel nicht vorhanden ist, findet keine Aktion statt.</param>
''' <exceptioncref="T:System.InvalidOperationException">
'''   <paramrefname="subkey"/> hat untergeordnete Unterschlüssel. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="subkey"/> gibt keinen gültigen Registrierungsschlüssel an, und <paramrefname="throwOnMissingSubKey"/> ist true. </exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Schlüssel zu löschen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Sub DeleteSubKey2(ByVal subkey As String, ByVal throwOnMissingSubKey As Boolean)
    ValidateKeyName subkey
    EnsureWriteable
    subkey = FixupName(subkey)
    CheckPermission RegistryInternalCheck.CheckSubKeyWritePermission, subkey, False, RegistryKeyPermissionCheck.[Default]
    Dim aRegistryKey As RegistryKey: Set aRegistryKey = InternalOpenSubKey(subkey, Writable:=False)
    If Not aRegistryKey Is Nothing Then
        'Try
            If aRegistryKey.InternalSubKeyCount > 0 Then
                'ThrowHelper.ThrowInvalidOperationException (ExceptionResource.InvalidOperation_RegRemoveSubKey)
                Err.Raise 0, "ThrowHelper.ThrowInvalidOperationException (ExceptionResource.InvalidOperation_RegRemoveSubKey)"
            End If
        'Finally
            aRegistryKey.CClose
        'End Try
        Dim num As Long
        'Try
            num = RegDeleteKeyExW(hKey, subkey, CInt(regView), 0)
        'Catch __unusedEntryPointNotFoundException1__ As EntryPointNotFoundException
            num = RegDeleteKeyW(hKey, subkey)
        'End Try
        Select Case num
        Case 0
        Case 2
            If throwOnMissingSubKey Then
                'ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyAbsent)
                Err.Raise "ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyAbsent)"
            End If
        Case Else
            'Win32Error(num, Nothing)
        End Select
    ElseIf throwOnMissingSubKey Then
        ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyAbsent)
    End If
End Sub

''' <summary>Löscht einen Unterschlüssel und alle untergeordneten Unterschlüssel rekursiv. </summary>
''' <paramname="subkey">Der zu löschende Unterschlüssel.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ArgumentException">Es wird versucht, eine Stammstruktur zu löschen.- oder -<paramrefname="subkey"/> gibt keinen gültigen Registrierungsunterschlüssel an. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Schlüssel zu löschen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
Public Sub DeleteSubKeyTree(ByVal subkey As String)
    DeleteSubKeyTree(subkey, throwOnMissingSubKey:=True)
End Sub

''' <summary>Löscht den angegebenen Unterschlüssel und untergeordnete Unterschlüssel rekursiv und gibt an, ob eine Ausnahme ausgelöst wird, wenn der Unterschlüssel nicht gefunden wird. </summary>
''' <paramname="subkey">Der Name des zu löschenden Unterschlüssels.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="throwOnMissingSubKey">Gibt an, ob eine Ausnahme ausgelöst werden soll, wenn der angegebene Unterschlüssel nicht gefunden werden kann.Wenn dieses Argument true und der angegebene Unterschlüssel nicht vorhanden ist, wird eine Ausnahme ausgelöst.Wenn dieses Argument false und der angegebene Unterschlüssel nicht vorhanden ist, findet keine Aktion statt.</param>
''' <exceptioncref="T:System.ArgumentException">Es wurde versucht, den Stammhive der Struktur zu löschen.- oder -<paramrefname="subkey"/> gibt keinen gültigen Registrierungsunterschlüssel an, und <paramrefname="throwOnMissingSubKey"/> ist true.</exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="subkey"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden).</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Schlüssel zu löschen.</exception>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Sub DeleteSubKeyTree(ByVal subkey As String, ByVal throwOnMissingSubKey As Boolean)
    RegistryKey.ValidateKeyName (subkey)
    If subkey.Length = 0 AndAlso IsSystemKey() Then
        ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegKeyDelHive)
    End If
    EnsureWriteable()
    subkey = RegistryKey.FixupName(subkey)
    CheckPermission(RegistryKey.RegistryInternalCheck.CheckSubTreeWritePermission, subkey, subKeyWritable:=False, RegistryKeyPermissionCheck.[Default])
    Dim RegistryKey = InternalOpenSubKey(subkey, writable:=True)
    If RegistryKey IsNot Nothing Then
        Try
            If RegistryKey.InternalSubKeyCount() > 0 Then
                Dim array As String() = RegistryKey.InternalGetSubKeyNames()
                For i = 0 To array.Length - 1
                    RegistryKey.DeleteSubKeyTreeInternal (Array(i))
                Next
            End If

        Finally
            RegistryKey.Close()
        End Try
        Dim num As Long
        Try
            num = Win32Native.RegDeleteKeyEx(hKey, subkey, CInt(regView), 0)
        Catch __unusedEntryPointNotFoundException1__ As EntryPointNotFoundException
            num = Win32Native.RegDeleteKey(hKey, subkey)
        End Try
        If num <> 0 Then
            Win32Error(num, Nothing)
        End If
    ElseIf throwOnMissingSubKey Then
        ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyAbsent)
    End If
End Sub

'<SecurityCritical>
Private Sub DeleteSubKeyTreeInternal(ByVal subkey As String)
    Dim RegistryKey = InternalOpenSubKey(subkey, writable:=True)
    If RegistryKey IsNot Nothing Then
        Try
            If RegistryKey.InternalSubKeyCount() > 0 Then
                Dim array As String() = RegistryKey.InternalGetSubKeyNames()
                For i = 0 To array.Length - 1
                    RegistryKey.DeleteSubKeyTreeInternal (Array(i))
                Next
            End If

        Finally
            RegistryKey.Close()
        End Try
        Dim num As Long
        Try
            num = Win32Native.RegDeleteKeyEx(hKey, subkey, CInt(regView), 0)
        Catch __unusedEntryPointNotFoundException1__ As EntryPointNotFoundException
            num = Win32Native.RegDeleteKey(hKey, subkey)
        End Try
        If num <> 0 Then
            Win32Error(num, Nothing)
        End If
    Else
        ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyAbsent)
    End If
End Sub

''' <summary>Löscht den angegebenen Wert aus diesem Schlüssel.</summary>
''' <paramname="name">Der Name des zu löschenden Werts. </param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="name"/> ist kein gültiger Verweis auf einen Wert. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Wert zu löschen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist schreibgeschützt. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
Public Sub DeleteValue(ByVal Name As String)
    DeleteValue(name, throwOnMissingValue:=True)
End Sub

''' <summary>Löscht den angegebenen Wert aus diesem Schlüssel und gibt an, ob eine Ausnahme ausgelöst wird, wenn der Wert nicht gefunden wird.</summary>
''' <paramname="name">Der Name des zu löschenden Werts. </param>
''' <paramname="throwOnMissingValue">Gibt an, ob eine Ausnahme ausgelöst werden soll, wenn der angegebene Wert nicht gefunden werden kann.Wenn dieses Argument true und der angegebene Wert nicht vorhanden ist, wird eine Ausnahme ausgelöst.Wenn dieses Argument false und der angegebene Wert nicht vorhanden ist, findet keine Aktion statt.</param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="name"/> ist kein gültiger Verweis auf einen Wert, und <paramrefname="throwOnMissingValue"/> ist true. - oder - <paramrefname="name"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Wert zu löschen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist schreibgeschützt. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Sub DeleteValue(ByVal Name As String, ByVal throwOnMissingValue As Boolean)
    EnsureWriteable()
    CheckPermission(RegistryInternalCheck.CheckValueWritePermission, name, subKeyWritable:=False, RegistryKeyPermissionCheck.[Default])
    Dim num As Long: num = Win32Native.RegDeleteValue(hKey, Name)
    If (num = 2 OrElse num = 206) AndAlso throwOnMissingValue Then
        ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegSubKeyValueAbsent)
    End If
End Sub

'<SecurityCritical>
Friend Function GetBaseKey(ByVal hKey As IntPtr) As RegistryKey
    Return GetBaseKey(hKey, RegistryView.Default)
End Function

'<SecurityCritical>
Friend Function GetBaseKey(ByVal hKey As IntPtr, ByVal View As RegistryView) As RegistryKey
    Dim num As Long: num = CInt(hKey) And &HFFFFFFF
    Dim flag = hKey Is RegistryKey.HKEY_PERFORMANCE_DATA
    Dim safeRegistryHandle As LongPtr = New SafeRegistryHandle(hKey, flag)
    Dim RegistryKey As RegistryKey = New RegistryKey(safeRegistryHandle, writable:=True, systemkey:=True, remoteKey:=False, flag, view)
    RegistryKey.checkMode = RegistryKeyPermissionCheck.[Default]
    RegistryKey.keyName = hkeyNames(num)
    Return RegistryKey
End Function

''' <summary>Öffnet einen neuen <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angeforderten Schlüssel auf dem lokalen Computer mit der angegebenen Ansicht darstellt.</summary>
''' <returns>Der angeforderte Registrierungsschlüssel.</returns>
''' <paramname="hKey">Der zu öffnende HKEY.</param>
''' <paramname="view">Die zu verwendende Registrierungsansicht.</param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="hKey"/> oder <paramrefname="view"/> ist ungültig.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über ausreichende Berechtigungen zum Ausführen dieser Aktion.</exception>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function OpenBaseKey(ByVal hKey As RegistryHive, ByVal View As RegistryView) As RegistryKey
    RegistryKey.ValidateKeyView (View)
    CheckUnmanagedCodePermission()
    Return RegistryKey.GetBaseKey(CType(CInt(hKey), IntPtr), view)
End Function

''' <summary>Öffnet einen neuen <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angeforderten Schlüssel für einen Remotecomputer darstellt.</summary>
''' <returns>Der angeforderte Registrierungsschlüssel.</returns>
''' <paramname="hKey">Der zu öffnende HKEY aus der <seecref="T:Microsoft.Win32.RegistryHive"/>-Enumeration. </param>
''' <paramname="machineName">Der Remotecomputer. </param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="hKey"/> ist ungültig.</exception>
''' <exceptioncref="T:System.IO.IOException">
'''   <paramrefname="machineName"/> wurde nicht gefunden.</exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="machineName"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über ausreichende Berechtigungen zum Ausführen dieser Aktion. </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
Public Function OpenRemoteBaseKey(ByVal hKey As RegistryHive, ByVal machineName As String) As RegistryKey
    Return OpenRemoteBaseKey(hKey, machineName, RegistryView.Default)
End Function

''' <summary>Öffnet einen neuen Registrierungsschlüssel, der den angeforderten Schlüssel auf einem Remotecomputer mit der angegebenen Ansicht darstellt.</summary>
''' <returns>Der angeforderte Registrierungsschlüssel.</returns>
''' <paramname="hKey">Der zu öffnende HKEY aus der <seecref="T:Microsoft.Win32.RegistryHive"/>-Enumeration. </param>
''' <paramname="machineName">Der Remotecomputer.</param>
''' <paramname="view">Die zu verwendende Registrierungsansicht.</param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="hKey"/> oder <paramrefname="view"/> ist ungültig.</exception>
''' <exceptioncref="T:System.IO.IOException">
'''   <paramrefname="machineName"/> wurde nicht gefunden.</exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="machineName"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="machineName"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über ausreichende Berechtigungen zum Ausführen dieser Aktion.</exception>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function OpenRemoteBaseKey(ByVal hKey As RegistryHive, ByVal machineName As String, ByVal View As RegistryView) As RegistryKey
    If Equals(machineName, Nothing) Then
        Throw New ArgumentNullException("machineName")
    End If
    Dim num As Long: num = CInt(hKey And CType(268435455, RegistryHive))
    If num < 0 OrElse num >= RegistryKey.hkeyNames.Length OrElse (CULng(hKey) And &HFFFFFFF0UL) <> 2147483648UI Then
        Throw New ArgumentException(Environment.GetResourceString("Arg_RegKeyOutOfRange"))
    End If
    RegistryKey.ValidateKeyView (View)
    RegistryKey.CheckUnmanagedCodePermission()
    Dim result As LongPtr '= Nothing
    Dim num2 As Long: num2 = RegConnectRegistry(machineName, New_SafeRegistryHandle(m_hkey), False, result)
    Select Case num2
        Case 1114
            Throw New ArgumentException(Environment.GetResourceString("Arg_DllInitFailure"))
        Case 0
        Case Else
            RegistryKey.Win32ErrorStatic(num2, Nothing)
    End Select
    If result.IsInvalid Then
        Throw New ArgumentException(Environment.GetResourceString("Arg_RegKeyNoRemoteConnect", machineName))
    End If
    Dim RegistryKey As RegistryKey = New RegistryKey(result, writable:=True, systemkey:=False, remoteKey:=True, CType(CInt(hKey), IntPtr) Is HKEY_PERFORMANCE_DATA, view)
    RegistryKey.checkMode = RegistryKeyPermissionCheck.[Default]
    RegistryKey.keyName = hkeyNames(num)
    Return RegistryKey
End Function

''' <summary>Ruft einen angegebenen Unterschlüssel ab und gibt an, ob Schreibzugriff auf den Schlüssel angewendet werden soll. </summary>
''' <returns>Der angeforderte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.</returns>
''' <paramname="name">Name oder Pfad des zu öffnenden Unterschlüssels. </param>
''' <paramname="writable">Muss für Schreibzugriff auf den Schlüssel auf true festgelegt werden. </param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="name"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um auf den Registrierungsschlüssel im angegebenen Modus zuzugreifen. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Function OpenSubKey(ByVal Name As String, ByVal Writable As Boolean) As RegistryKey
    ValidateKeyName Name
    EnsureNotDisposed
    Name = FixupName(Name)
    CheckPermission RegistryInternalCheck.CheckOpenSubKeyWithWritablePermission, Name, Writable, RegistryKeyPermissionCheck.[Default]
    Dim hkResult As LongPtr '= Nothing
    Dim num As Long: num = RegOpenKeyExW(m_hkey, StrPtr(Name), 0, GetRegistryKeyAccess(Writable) Or CLng(regView), hkResult)
    If num = ERROR_SUCCESS Then
        If Not hkResult.IsInvalid Then
        Dim aRegistryKey As RegistryKey: Set aRegistryKey = New_RegistryKey(hkResult, Writable, False, remoteKey, False, regView)
        RegistryKey.checkMode = GetSubKeyPermissonCheck(Writable)
        RegistryKey.keyName = m_keyName & "\" & Name
        Set OpenSubKey = aRegistryKey
    End If
    If num = 5 Then
        If num = 1346 Then
            'ThrowHelper.ThrowSecurityException (ExceptionResource.Security_RegistryPermission)
            Err.Raise 0, "ThrowHelper.ThrowSecurityException (ExceptionResource.Security_RegistryPermission)"
        End If
    End If
    'Return Nothing
End Function

''' <summary>Ruft den angegebenen Unterschlüssel für Lesezugriff oder Lese-/Schreibzugriff ab.</summary>
''' <returns>Der angeforderte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.</returns>
''' <paramname="name">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.</param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="name"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="permissionCheck"/> enthält einen ungültigen Wert.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Registrierungsschlüssel zu lesen. </exception>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function OpenSubKey(ByVal Name As String, ByVal PermissionCheck As RegistryKeyPermissionCheck) As RegistryKey
    RegistryKey.ValidateKeyMode (PermissionCheck)
    Return Me.InternalOpenSubKey(name, permissionCheck, RegistryKey.GetRegistryKeyAccess(permissionCheck))
End Function

'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function OpenSubKey(ByVal Name As String, ByVal rights As RegistryRights) As RegistryKey
    Return Me.InternalOpenSubKey(name, checkMode, CInt(rights))
End Function

''' <summary>Ruft den angegebenen Unterschlüssel für Lesezugriff oder Lese-/Schreibzugriff ab, und fordert die angegebenen Zugriffsrechte an.</summary>
''' <returns>Der angeforderte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.</returns>
''' <paramname="name">Name oder Pfad des zu erstellenden oder zu öffnenden Unterschlüssels.</param>
''' <paramname="permissionCheck">Einer der Enumerationswerte, der angibt, ob der Schlüssel für Lesezugriff oder für Lese-/Schreibzugriff geöffnet wird.</param>
''' <paramname="rights">Eine bitweise Kombination von Enumerationswerten, die den gewünschten Sicherheitszugriff angeben.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="name"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="permissionCheck"/> enthält einen ungültigen Wert.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.Security.SecurityException">
'''   <paramrefname="rights"/> enthält ungültige Werte für Registrierungsrechte.- oder -Der Benutzer verfügt nicht über die angeforderten Berechtigungen. </exception>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function OpenSubKey(ByVal Name As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal rights As RegistryRights) As RegistryKey
    Return Me.InternalOpenSubKey(name, permissionCheck, CInt(rights))
End Function

'<SecurityCritical>
Private Function InternalOpenSubKey(ByVal Name As String, ByVal PermissionCheck As RegistryKeyPermissionCheck, ByVal rights As Long) As RegistryKey
    ValidateKeyName (Name)
    RegistryKey.ValidateKeyMode (PermissionCheck)
    ValidateKeyRights (rights)
    EnsureNotDisposed
    Name = FixupName(Name)
    Me.CheckPermission(RegistryInternalCheck.CheckOpenSubKeyPermission, name, subKeyWritable:=False, permissionCheck)
    Me.CheckPermission(RegistryInternalCheck.CheckSubTreePermission, name, subKeyWritable:=False, permissionCheck)
    Dim hkResult As LongPtr = Nothing
    Dim num As Long: num = Win32Native.RegOpenKeyEx(hKey, Name, 0, rights Or CInt(regView), hkResult)
    If num = 0 AndAlso Not hkResult.IsInvalid Then
        Dim RegistryKey As RegistryKey = New RegistryKey(hkResult, permissionCheck Is RegistryKeyPermissionCheck.ReadWriteSubTree, systemkey:=False, remoteKey, isPerfData:=False, regView)
        RegistryKey.keyName = m_keyName & "\" & Name
        RegistryKey.checkMode = PermissionCheck
        Return RegistryKey
    End If
    If num = 5 OrElse num = 1346 Then
        ThrowHelper.ThrowSecurityException (ExceptionResource.Security_RegistryPermission)
    End If
    Return Nothing
End Function

'<SecurityCritical>
Friend Function InternalOpenSubKey(ByVal Name As String, ByVal Writable As Boolean) As RegistryKey
    ValidateKeyName (Name)
    EnsureNotDisposed
    Dim hkResult As LongPtr = Nothing
    If Win32Native.RegOpenKeyEx(hkey, name, 0, GetRegistryKeyAccess(writable) Or CInt(regView), hkResult) = 0 AndAlso Not hkResult.IsInvalid Then
        Dim RegistryKey As RegistryKey = New RegistryKey(hkResult, writable, systemkey:=False, remoteKey, isPerfData:=False, regView)
        RegistryKey.keyName = m_keyName & "\" & Name
        Return RegistryKey
    End If
    Return Nothing
End Function

''' <summary>Ruft einen Unterschlüssel als schreibgeschützt ab.</summary>
''' <returns>Der angeforderte Unterschlüssel oder null, wenn bei dem Vorgang ein Fehler aufgetreten ist.</returns>
''' <paramname="name">Der Name oder der Pfad des Unterschlüssels, der schreibgeschützt geöffnet werden soll. </param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="name"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um den Registrierungsschlüssel zu lesen. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
Public Function OpenSubKey(ByVal Name As String) As RegistryKey
    Return OpenSubKey(name, writable:=False)
End Function

''' <summary>Erstellt einen Registrierungsschlüssel aus einem angegebenen Handle.</summary>
''' <returns>Ein Registrierungsschlüssel.</returns>
''' <paramname="handle">Das Handle für den Registrierungsschlüssel.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="handle"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über ausreichende Berechtigungen zum Ausführen dieser Aktion.</exception>
'<SecurityCritical>
'<ComVisible(False)>
'<SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.UnmanagedCode)>
Public Function FromHandle(ByVal Handle As LongPtr) As RegistryKey
    Return FromHandle(handle, RegistryView.Default)
End Function

''' <summary>Erstellt einen Registrierungsschlüssel aus einem angegebenen Handle und einer Registrierungsansichtseinstellung. </summary>
''' <returns>Ein Registrierungsschlüssel.</returns>
''' <paramname="handle">Das Handle für den Registrierungsschlüssel.</param>
''' <paramname="view">Die zu verwendende Registrierungsansicht.</param>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="view"/> ist ungültig.</exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="handle"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über ausreichende Berechtigungen zum Ausführen dieser Aktion.</exception>
'<SecurityCritical>
'<ComVisible(False)>
'<SecurityPermission(SecurityAction.Demand, Flags:=SecurityPermissionFlag.UnmanagedCode)>
Public Function FromHandle(ByVal Handle As LongPtr, ByVal View As RegistryView) As RegistryKey
    If Handle Is Nothing Then
        Throw New ArgumentNullException("handle")
    End If
    RegistryKey.ValidateKeyView (View)
    Return New RegistryKey(handle, writable:=True, view)
End Function

'<SecurityCritical>
Friend Function InternalSubKeyCount() As Long
    EnsureNotDisposed
    Dim lpcSubKeys = 0
    Dim lpcValues = 0
    Dim num As Long: num = Win32Native.RegQueryInfoKey(hKey, Nothing, Nothing, IntPtr.Zero, lpcSubKeys, Nothing, Nothing, lpcValues, Nothing, Nothing, Nothing, Nothing)
    If num <> 0 Then
        Win32Error(num, Nothing)
    End If
    Return lpcSubKeys
End Function

''' <summary>Ruft ein Array von Zeichenfolgen mit den Namen aller Unterschlüssel ab.</summary>
''' <returns>Ein Array von Zeichenfolgen, das die Namen der Unterschlüssel des aktuellen Schlüssels enthält.</returns>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Schlüssel zu lesen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <exceptioncref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, möglicherweise wurde der aktuelle Schlüssel gelöscht.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Function GetSubKeyNames() As String()
    CheckPermission(RegistryInternalCheck.CheckKeyReadPermission, Nothing, subKeyWritable:=False, RegistryKeyPermissionCheck.[Default])
    Return Me.InternalGetSubKeyNames()
End Function

    ''' Cannot convert MethodDeclarationSyntax, System.NotSupportedException: UnsafeKeyword is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\SyntaxKindExtensions.cs:line 188
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifier(SyntaxToken m, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 465
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.MoveNext()
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.ToList()
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifiersCore(IReadOnlyCollection`1 modifiers, TokenContext context, Boolean isConstructor) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 418
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitMethodDeclaration(MethodDeclarationSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 418
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''
'''     [SecurityCritical]
Private Function InternalGetSubKeyNames() As String()
    EnsureNotDisposed
    Dim num As Integer: num = InternalSubKeyCount()
    ReDim SArray(0 To num - 1) As String
    If num > 0 Then
        Dim array2 = New Char(255) {}
        Dim ptr As Char = 0
        If Ptr = array2(0) Then
            For i = 0 To num - 1
                Dim lpcbName = array2.Length
                Dim num2 As Integer: num2 = RegEnumKeyExW(hKey, i, Ptr, lpcbName, Nothing, Nothing, Nothing, Nothing)
                If num2 <> 0 Then
                    Win32Error(num2, Nothing)
                End If
                array2(i) = New String(ptr)
            Next
        End If
    End If
    Return array2
End Function

'<SecurityCritical>
Friend Function InternalValueCount() As Long
    EnsureNotDisposed
    Dim lpcValues = 0
    Dim lpcSubKeys = 0
    Dim num As Long: num = Win32Native.RegQueryInfoKey(hKey, Nothing, Nothing, IntPtr.Zero, lpcSubKeys, Nothing, Nothing, lpcValues, Nothing, Nothing, Nothing, Nothing)
    If num <> 0 Then
        Win32Error(num, Nothing)
    End If
    Return lpcValues

    ''' <summary>Ruft ein Array von Zeichenfolgen ab, das die Namen aller diesem Schlüssel zugeordneten Werte enthält.</summary>
    ''' <returns>Ein Array von Zeichenfolgen, das die Namen der Werte für den aktuellen Schlüssel enthält.</returns>
    ''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
    ''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
    ''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
    ''' <exceptioncref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, möglicherweise wurde der aktuelle Schlüssel gelöscht.</exception>
    ''' <PermissionSet>
    '''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
    ''' </PermissionSet>
End Function

    ''' Cannot convert MethodDeclarationSyntax, System.NotSupportedException: UnsafeKeyword is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\SyntaxKindExtensions.cs:line 188
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifier(SyntaxToken m, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 465
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.MoveNext()
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.ToList()
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifiersCore(IReadOnlyCollection`1 modifiers, TokenContext context, Boolean isConstructor) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 418
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitMethodDeclaration(MethodDeclarationSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 418
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''
'''     /// <summary>Ruft ein Array von Zeichenfolgen ab, das die Namen aller diesem Schlüssel zugeordneten Werte enthält.</summary>
'''     /// <returns>Ein Array von Zeichenfolgen, das die Namen der Werte für den aktuellen Schlüssel enthält.</returns>
'''     /// <exception cref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
'''     /// <exception cref="T:System.ObjectDisposedException">Der zu bearbeitende <see cref="T:Microsoft.Win32.RegistryKey" /> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
'''     /// <exception cref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
'''     /// <exception cref="T:System.IO.IOException">Ein Systemfehler ist aufgetreten, möglicherweise wurde der aktuelle Schlüssel gelöscht.</exception>
'''     /// <PermissionSet>
'''     ///   <IPermission class="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Unrestricted="true" />
'''     /// </PermissionSet>
'''     [SecuritySafeCritical]
'''     public unsafe string[] GetValueNames()
'''     {
'''         CheckPermission(RegistryKey.RegistryInternalCheck.CheckKeyReadPermission, null, subKeyWritable: false, RegistryKeyPermissionCheck.Default);
'''         this.EnsureNotDisposed;
'''         int num = this.InternalValueCount();
'''         string[] array = new string[num];
'''         if (num > 0)
'''         {
'''             char[] array2 = new char[16384];
'''             fixed (char* ptr = &array2[0])
'''             {
'''                 for (int i = 0; i < num; i++)
'''                 {
'''                     int lpcbValueName = array2.Length;
'''                     int num2 = Win32Native.RegEnumValue(this.hkey, i, ptr, ref lpcbValueName, IntPtr.Zero, null, null, null);
'''                     if (num2 != 0 && (!this.IsPerfDataKey() || num2 != 234))
'''                     {
'''                         this.Win32Error(num2, null);
'''                     }
'''                     array[i] = new string(ptr);
'''                 }
'''             }
'''         }
'''         return array;
'''     }
'''
'''

''' <summary>Ruft den Wert ab, der dem angegebenen Namen zugeordnet ist.Gibt null zurück, wenn das Name-Wert-Paar in der Registrierung nicht vorhanden ist.</summary>
''' <returns>Der <paramrefname="name"/> zugeordnete Wert oder null, wenn <paramrefname="name"/> nicht gefunden wurde.</returns>
''' <paramname="name">Der Name des abzurufenden Werts.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.IO.IOException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, wurde zum Löschen markiert. </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Function GetValue(ByVal Name As String) As Object
    CheckPermission RegistryInternalCheck.CheckValueReadPermission, Name, False, RegistryKeyPermissionCheck.Default
    GetValue = InternalGetValue(Name, Nothing, doNotExpand:=False, checkSecurity:=True)
End Function

''' <summary>Ruft den Wert ab, der dem angegebenen Namen zugeordnet ist.Wenn der Name nicht gefunden wird, wird der von Ihnen bereitgestellte Standardwert zurückgegeben.</summary>
''' <returns>Der <paramrefname="name"/> zugeordnete Wert mit allen eingebetteten Umgebungsvariablen, die nicht erweitert wurden, oder <paramrefname="defaultValue"/>, wenn <paramrefname="name"/> nicht gefunden wurde.</returns>
''' <paramname="name">Der Name des abzurufenden Werts.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="defaultValue">Der zurückzugebende Wert, wenn <paramrefname="name"/> nicht vorhanden ist. </param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.IO.IOException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, wurde zum Löschen markiert. </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Function GetValue(ByVal Name As String, ByVal defaultValue As Object) As Object
    CheckPermission(RegistryInternalCheck.CheckValueReadPermission, name, subKeyWritable:=False, RegistryKeyPermissionCheck.Default)
    Return InternalGetValue(name, defaultValue, doNotExpand:=False, checkSecurity:=True)
End Function

''' <summary>Ruft den Wert ab, der dem angegebenen Namen und den Abrufoptionen zugeordnet ist.Wenn der Name nicht gefunden wird, wird der von Ihnen bereitgestellte Standardwert zurückgegeben.</summary>
''' <returns>Der <paramrefname="name"/> zugeordnete Wert, der entsprechend den <paramrefname="options"/> verarbeitet wurde, oder <paramrefname="defaultValue"/>, wenn <paramrefname="name"/> nicht gefunden wird.</returns>
''' <paramname="name">Der Name des abzurufenden Werts.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <paramname="defaultValue">Der zurückzugebende Wert, wenn <paramrefname="name"/> nicht vorhanden ist. </param>
''' <paramname="options">Einer der Enumerationswerte, die eine optionale Verarbeitung des abgerufenen Werts angeben.</param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.IO.IOException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, wurde zum Löschen markiert. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="options"/> ist kein gültiger <seecref="T:Microsoft.Win32.RegistryValueOptions"/>-Wert; ein ungültiger Wert wird z. B. in <seecref="T:Microsoft.Win32.RegistryValueOptions"/> umgewandelt.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function GetValue(ByVal Name As String, ByVal defaultValue As Object, ByVal Options As RegistryValueOptions) As Object
    If options < RegistryValueOptions.None OrElse options > RegistryValueOptions.DoNotExpandEnvironmentNames Then
        Throw New ArgumentException(Environment.GetResourceString("Arg_EnumIllegalVal", CInt(options)), "options")
    End If
    Dim doNotExpand = options Is RegistryValueOptions.DoNotExpandEnvironmentNames
    CheckPermission(RegistryInternalCheck.CheckValueReadPermission, name, subKeyWritable:=False, RegistryKeyPermissionCheck.[Default])
    Return InternalGetValue(name, defaultValue, doNotExpand, checkSecurity:=True)
End Function

'<SecurityCritical>
Friend Function InternalGetValue(ByVal Name As String, ByVal defaultValue As Object, ByVal doNotExpand As Boolean, ByVal checkSecurity As Boolean) As Object
    If checkSecurity Then
        EnsureNotDisposed
    End If
    Dim obj = defaultValue
    Dim lpType = 0
    Dim lpcbData = 0
    Dim num As Long: num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, Nothing, lpcbData)
    If num <> 0 Then
        If IsPerfDataKey() Then
            Dim nums As Long: nums = 65000
            Dim lpcbData2 = num2
            Dim array = New Byte(num2 - 1) {}
            Dim num3 As Long
            While 234 = CSharpImpl.__Assign(num3, Win32Native.RegQueryValueEx(hkey, name, Nothing, lpType, array, lpcbData2))
                If num2 = Integer.MaxValue Then
                    Win32Error(num3, name)
                Else
                    num2 = If(num2 <= 1073741823, num2 * 2, Integer.MaxValue)
                End If
                lpcbData2 = num2
                array = New Byte(num2 - 1) {}
            Wend
            If num3 <> 0 Then
                Win32Error(num3, name)
            End If
            Return array
        End If
        If num <> 234 Then
            Return obj
        End If
    End If
    If lpcbData < 0 Then
        lpcbData = 0
    End If
    Select Case lpType
        Case 0, 3, 5
_Select0_Case0:
            Dim array5 = New Byte(lpcbData - 1) {}
            num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, array5, lpcbData)
            obj = array5
            Exit Select
        Case 11
_Select0_Case11:
            If lpcbData <= 8 Then
                Dim lpData2 = 0L
                num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, lpData2, lpcbData)
                obj = lpData2
                Exit Select
            End If
            GoTo _Select0_Case0
        Case 4
            If lpcbData <= 4 Then
                Dim lpData = 0
                num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, lpData, lpcbData)
                obj = lpData
                Exit Select
            End If
            GoTo _Select0_Case11
        Case 1
            If lpcbData Mod 2 = 1 Then
                Try
                ''' Cannot convert AssignmentExpressionSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitAssignmentExpression(AssignmentExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 969
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''                     lpcbData = checked(lpcbData + 1)
'''                 Catch innerException4 As OverflowException
                    Throw New IOException(Environment.GetResourceString("Arg_RegGetOverflowBug"), innerException4)
                End Try
            End If
            Dim array6 = New Char(lpcbData / 2 - 1) {}
            num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, array6, lpcbData)
            obj = (If(array6.Length = 0 OrElse array6(array6.Length - 1) <> 0, New String(array6), New String(array6, 0, array6.Length - 1)))
            Exit Select
        Case 2
            If lpcbData Mod 2 = 1 Then
                Try
                ''' Cannot convert AssignmentExpressionSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitAssignmentExpression(AssignmentExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 969
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''                     lpcbData = checked(lpcbData + 1)
'''                 Catch innerException3 As OverflowException
                    Throw New IOException(Environment.GetResourceString("Arg_RegGetOverflowBug"), innerException3)
                End Try
            End If
            Dim array4 = New Char(lpcbData / 2 - 1) {}
            num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, array4, lpcbData)
            obj = (If(array4.Length = 0 OrElse array4(array4.Length - 1) <> 0, New String(array4), New String(array4, 0, array4.Length - 1)))
            If Not doNotExpand Then
                obj = Environment.ExpandEnvironmentVariables(CStr(obj))
            End If
            Exit Select
        Case 7
            If lpcbData Mod 2 = 1 Then
                Try
                ''' Cannot convert AssignmentExpressionSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitAssignmentExpression(AssignmentExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 969
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''                     lpcbData = checked(lpcbData + 1)
'''                 Catch innerException As OverflowException
                    Throw New IOException(Environment.GetResourceString("Arg_RegGetOverflowBug"), innerException)
                End Try
            End If
            Dim array2 = New Char(lpcbData / 2 - 1) {}
            num = Win32Native.RegQueryValueEx(hKey, Name, Nothing, lpType, array2, lpcbData)
            If array2.Length <> 0 AndAlso array2(array2.Length - 1) <> 0 Then
                Try
                                            ''' Cannot convert LocalDeclarationStatementSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.RemodelVariableDeclaration(VariableDeclarationSyntax declaration) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 471
'''    at ICSharpCode.CodeConverter.VB.MethodBodyExecutableStatementVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\MethodBodyExecutableStatementVisitor.cs:line 52
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingMethodBodyVisitor.cs:line 16
'''
''' Input:
'''                     char[] array3 = new char[checked(array2.Length + 1)];
'''
'''
                    For i = 0 To array2.Length - 1
                        array3(i) = array2(i)
                    Next
                    array3(array3.Length - 1) = ChrW(0)
                    array2 = array3
                Catch innerException2 As OverflowException
                    Throw New IOException(Environment.GetResourceString("Arg_RegGetOverflowBug"), innerException2)
                End Try
                array2(array2.Length - 1) = ChrW(0)
            End If
            Dim list As IList(Of String) = New List(Of String)()
            Dim num4 = 0
            Dim num5 = array2.Length
            While num = 0 AndAlso num4 < num5
                Dim j As Long
                j = num4

                While j < num5 AndAlso array2(j) <> 0
                    j += 1
                Wend
                If j < num5 Then
                    If j - num4 > 0 Then
                        list.Add(New String(array2, num4, j - num4))
                    ElseIf j <> num5 - 1 Then
                        list.Add(String.Empty)
                    End If
                Else
                    list.Add(New String(array2, num4, num5 - num4))
                End If
                num4 = j + 1
            Wend
            obj = New String(list.Count - 1) {}
            list.CopyTo(CType(obj, String()), 0)
            Exit Select
    End Select
    Return obj
End Function

''' <summary>Ruft den Registrierungsdatentyp des Werts ab, der dem angegebenen Namen zugeordnet ist.</summary>
''' <returns>Der Registrierungsdatentyp des <paramrefname="name"/> zugeordneten Werts.</returns>
''' <paramname="name">Der Name des Werts, dessen Registrierungsdatentyp abgerufen werden soll.Bei dieser Zeichenfolge wird die Groß-/Kleinschreibung nicht berücksichtigt.</param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.IO.IOException">Der Unterschlüssel, der den angegebenen Wert enthält, ist nicht vorhanden.- oder -Das von <paramrefname="name"/> angegebene Name-Wert-Paar ist nicht vorhanden.Diese Ausnahme wird unter Windows 95, Windows 98 oder Windows Millennium Edition nicht ausgelöst.</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der Benutzer verfügt nicht über die notwendigen Registrierungsrechte.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
'<SecuritySafeCritical>
'<ComVisible(False)>
Public Function GetValueKind(ByVal Name As String) As RegistryValueKind
    CheckPermission RegistryInternalCheck.CheckValueReadPermission, Name, False, RegistryKeyPermissionCheck.[Default]
    EnsureNotDisposed
    Dim lpType As LongPtr '= 0
    Dim lpcbData As LongPtr '= 0
    Dim num As Long: num = RegQueryValueExW(hKey, StrPtr(Name), 0, VarPtr(lpType), 0, VarPtr(lpcbData))
    If num <> 0 Then
        Win32Error num, 0
    End If
    If lpType = 0 Then
        GetValueKind = RegistryValueKind.rvkNone
        Exit Function
    End If
    If Not [Enum].IsDefined(GetType(RegistryValueKind), lpType) Then
        GetValueKind = RegistryValueKind.rvkUnknown
        Exit Function
    End If
    GetValueKind = CType(lpType, RegistryValueKind)
End Function

Private Function IsDirty() As Boolean
    IsDirty = (m_state And STATE_DIRTY) <> 0 'STATE_DIRTY           As Long = 1
End Function

Private Function IsSystemKey() As Boolean
    IsSystemKey = (m_state And STATE_SYSTEMKEY) <> 0 'STATE_SYSTEMKEY       As Long = 2
End Function

Private Function IsWritable() As Boolean
    IsWritable = (m_state And STATE_WRITEACCESS) <> 0 'STATE_WRITEACCESS     As Long = 4
End Function

Private Function IsPerfDataKey() As Boolean
    IsPerfDataKey = (m_state And STATE_PERF_DATA) <> 0 'STATE_PERF_DATA       As Long = 8
End Function

Private Sub SetDirty()
    m_state = m_state Or STATE_DIRTY
End Sub

''' <summary>Legt das angegebene Name-Wert-Paar fest.</summary>
''' <paramname="name">Der Name des zu speichernden Werts. </param>
''' <paramname="value">Die zu speichernden Daten. </param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="value"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="value"/> ist ein nicht unterstützter Datentyp. </exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist schreibgeschützt, sodass das Schreiben in den Schlüssel nicht möglich ist. Möglicherweise wurde der Schlüssel ohne Schreibzugriff geöffnet. - oder -Das <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt stellt einen Knoten auf Stammebene dar, und das Betriebssystem ist Windows Millennium Edition oder Windows 98.</exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Ändern von Registrierungsschlüsseln. </exception>
''' <exceptioncref="T:System.IO.IOException">Das <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt stellt einen Knoten auf der Stammebene dar, und das Betriebssystem ist Windows 2000, Windows XP oder Windows Server 2003.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
Public Sub SetValue(ByVal Name As String, ByVal value)
    SetValue Name, value, RegistryValueKind.rvkUnknown

    ''' <summary>Legt mithilfe des angegebenen Registrierungsdatentyps den Wert eines Name-Wert-Paars im Registrierungsschlüssel fest.</summary>
    ''' <paramname="name">Der Name des zu speichernden Werts. </param>
    ''' <paramname="value">Die zu speichernden Daten. </param>
    ''' <paramname="valueKind">Der beim Speichern der Daten zu verwendende Registrierungsdatentyp. </param>
    ''' <exceptioncref="T:System.ArgumentNullException">
    '''   <paramrefname="value"/> hat den Wert null. </exception>
    ''' <exceptioncref="T:System.ArgumentException">Der Typ von <paramrefname="value"/> stimmt nicht mit dem durch <paramrefname="valueKind"/> angegebenen Registrierungsdatentyp überein. Die Daten konnten daher nicht ordnungsgemäß konvertiert werden. </exception>
    ''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
    ''' <exceptioncref="T:System.UnauthorizedAccessException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist schreibgeschützt, sodass das Schreiben in den Schlüssel nicht möglich ist. Möglicherweise wurde der Schlüssel ohne Schreibzugriff geöffnet.- oder -Das <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt stellt einen Knoten auf Stammebene dar, und das Betriebssystem ist Windows Millennium Edition oder Windows 98. </exception>
    ''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Ändern von Registrierungsschlüsseln. </exception>
    ''' <exceptioncref="T:System.IO.IOException">Das <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt stellt einen Knoten auf der Stammebene dar, und das Betriebssystem ist Windows 2000, Windows XP oder Windows Server 2003.</exception>
    ''' <PermissionSet>
    '''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
    '''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
    ''' </PermissionSet>
End Sub

    ''' Cannot convert MethodDeclarationSyntax, System.NotSupportedException: UnsafeKeyword is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\SyntaxKindExtensions.cs:line 188
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifier(SyntaxToken m, TokenContext context) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 465
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.MoveNext()
'''    at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.ToList()
'''    at ICSharpCode.CodeConverter.VB.CommonConversions.ConvertModifiersCore(IReadOnlyCollection`1 modifiers, TokenContext context, Boolean isConstructor) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommonConversions.cs:line 418
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitMethodDeclaration(MethodDeclarationSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 418
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 20
'''
''' Input:
'''
'''     /// <summary>Legt mithilfe des angegebenen Registrierungsdatentyps den Wert eines Name-Wert-Paars im Registrierungsschlüssel fest.</summary>
'''     /// <param name="name">Der Name des zu speichernden Werts. </param>
'''     /// <param name="value">Die zu speichernden Daten. </param>
'''     /// <param name="valueKind">Der beim Speichern der Daten zu verwendende Registrierungsdatentyp. </param>
'''     /// <exception cref="T:System.ArgumentNullException">
'''     ///   <paramref name="value" /> hat den Wert null. </exception>
'''     /// <exception cref="T:System.ArgumentException">Der Typ von <paramref name="value" /> stimmt nicht mit dem durch <paramref name="valueKind" /> angegebenen Registrierungsdatentyp überein. Die Daten konnten daher nicht ordnungsgemäß konvertiert werden. </exception>
'''     /// <exception cref="T:System.ObjectDisposedException">Der <see cref="T:Microsoft.Win32.RegistryKey" />, der den angegebenen Wert enthält, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
'''     /// <exception cref="T:System.UnauthorizedAccessException">Der <see cref="T:Microsoft.Win32.RegistryKey" /> ist schreibgeschützt, sodass das Schreiben in den Schlüssel nicht möglich ist. Möglicherweise wurde der Schlüssel ohne Schreibzugriff geöffnet.- oder -Das <see cref="T:Microsoft.Win32.RegistryKey" />-Objekt stellt einen Knoten auf Stammebene dar, und das Betriebssystem ist Windows Millennium Edition oder Windows 98. </exception>
'''     /// <exception cref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Ändern von Registrierungsschlüsseln. </exception>
'''     /// <exception cref="T:System.IO.IOException">Das <see cref="T:Microsoft.Win32.RegistryKey" />-Objekt stellt einen Knoten auf der Stammebene dar, und das Betriebssystem ist Windows 2000, Windows XP oder Windows Server 2003.</exception>
'''     /// <PermissionSet>
'''     ///   <IPermission class="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Unrestricted="true" />
'''     ///   <IPermission class="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" version="1" Flags="UnmanagedCode" />
'''     /// </PermissionSet>
'''     [SecuritySafeCritical]
'''     [ComVisible(false)]
'''     public unsafe void SetValue(string name, object value, RegistryValueKind valueKind)
'''     {
'''         if (value == null)
'''         {
'''             ThrowHelper.ThrowArgumentNullException(ExceptionArgument.value);
'''         }
'''         if (name != null && name.Length > 16383)
'''         {
'''             throw new ArgumentException(Environment.GetResourceString("Arg_RegValStrLenBug"));
'''         }
'''         if (!Enum.IsDefined(typeof(RegistryValueKind), valueKind))
'''         {
'''             throw new ArgumentException(Environment.GetResourceString("Arg_RegBadKeyKind"), "valueKind");
'''         }
'''         this.EnsureWriteable();
'''         if (!this.remoteKey && this.ContainsRegistryValue(name))
'''         {
'''             CheckPermission(RegistryKey.RegistryInternalCheck.CheckValueWritePermission, name, subKeyWritable: false, RegistryKeyPermissionCheck.Default);
'''         }
'''         else
'''         {
'''             CheckPermission(RegistryKey.RegistryInternalCheck.CheckValueCreatePermission, name, subKeyWritable: false, RegistryKeyPermissionCheck.Default);
'''         }
'''         if (valueKind == RegistryValueKind.Unknown)
'''         {
'''             valueKind = this.CalculateValueKind(value);
'''         }
'''         int num = 0;
'''         try
'''         {
'''             switch (valueKind)
'''             {
'''             case RegistryValueKind.String:
'''             case RegistryValueKind.ExpandString:
'''             {
'''                 string text = value.ToString();
'''                 num = Win32Native.RegSetValueEx(this.hkey, name, 0, valueKind, text, checked(text.Length * 2 + 2));
'''                 break;
'''             }
'''             case RegistryValueKind.MultiString:
'''             {
'''                 string[] array2 = (string[])((string[])value).Clone();
'''                 int num2 = 0;
'''                 for (int i = 0; i < array2.Length; i++)
'''                 {
'''                     if (array2[i] == null)
'''                     {
'''                         ThrowHelper.ThrowArgumentException(ExceptionResource.Arg_RegSetStrArrNull);
'''                     }
'''                     num2 = checked(num2 + (array2[(System.Int32)(i)].Length + 1) * 2);
'''                 }
'''                 num2 = checked(num2 + 2);
'''                 byte[] array3 = new byte[num2];
'''                 try
'''                 {
'''                     fixed (byte* value2 = array3)
'''                     {
'''                         IntPtr intPtr = new IntPtr(value2);
'''                         for (int j = 0; j < array2.Length; j++)
'''                         {
'''                             string.InternalCopy(array2[j], intPtr, checked(array2[(System.Int32)(j)].Length * 2));
'''                             intPtr = new IntPtr((long)intPtr + checked(array2[(System.Int32)(j)].Length * 2));
'''                             *(short*)intPtr.ToPointer() = 0;
'''                             intPtr = new IntPtr((long)intPtr + 2);
'''                         }
'''                         *(short*)intPtr.ToPointer() = 0;
'''                         intPtr = new IntPtr((long)intPtr + 2);
'''                         num = Win32Native.RegSetValueEx(this.hkey, name, 0, RegistryValueKind.MultiString, array3, num2);
'''                     }
'''                 }
'''                 finally
'''                 {
'''                 }
'''                 break;
'''             }
'''             case RegistryValueKind.None:
'''             case RegistryValueKind.Binary:
'''             {
'''                 byte[] array = (byte[])value;
'''                 num = Win32Native.RegSetValueEx(this.hkey, name, 0, (valueKind != RegistryValueKind.None) ? RegistryValueKind.Binary : RegistryValueKind.Unknown, array, array.Length);
'''                 break;
'''             }
'''             case RegistryValueKind.DWord:
'''             {
'''                 int lpData2 = Convert.ToInt32(value, CultureInfo.InvariantCulture);
'''                 num = Win32Native.RegSetValueEx(this.hkey, name, 0, RegistryValueKind.DWord, ref lpData2, 4);
'''                 break;
'''             }
'''             case RegistryValueKind.QWord:
'''             {
'''                 long lpData = Convert.ToInt64(value, CultureInfo.InvariantCulture);
'''                 num = Win32Native.RegSetValueEx(this.hkey, name, 0, RegistryValueKind.QWord, ref lpData, 8);
'''                 break;
'''             }
'''             case RegistryValueKind.Unknown:
'''             case (RegistryValueKind)5:
'''             case (RegistryValueKind)6:
'''             case (RegistryValueKind)8:
'''             case (RegistryValueKind)9:
'''             case (RegistryValueKind)10:
'''                 break;
'''             }
'''         }
'''         catch (OverflowException)
'''         {
'''             ThrowHelper.ThrowArgumentException(ExceptionResource.Arg_RegSetMismatchedKind);
'''         }
'''         catch (InvalidOperationException)
'''         {
'''             ThrowHelper.ThrowArgumentException(ExceptionResource.Arg_RegSetMismatchedKind);
'''         }
'''         catch (FormatException)
'''         {
'''             ThrowHelper.ThrowArgumentException(ExceptionResource.Arg_RegSetMismatchedKind);
'''         }
'''         catch (InvalidCastException)
'''         {
'''             ThrowHelper.ThrowArgumentException(ExceptionResource.Arg_RegSetMismatchedKind);
'''         }
'''         if (num == 0)
'''         {
'''             this.SetDirty();
'''         }
'''         else
'''         {
'''             this.Win32Error(num, null);
'''         }
'''     }
'''
'''

Private Function CalculateValueKind(value) As RegistryValueKind
    Dim vt As VbVarType: vt = VarType(value)
    If vt = vbLong Then
        CalculateValueKind = RegistryValueKind.rvkDWord
        Exit Function
    End If
    If vt And vbArray Then
        vt = vt - vbArray
        If vt = vbByte Then
            CalculateValueKind = RegistryValueKind.rvkBinary
            Exit Function
        End If
        If vt = vbString Then
            CalculateValueKind = RegistryValueKind.rvkMultiString
            Exit Function
        End If
        'Throw New ArgumentException(Environment.GetResourceString("Arg_RegSetBadArrType", value.GetType().Name))
        Err.Raise &H7007, , "Throw New ArgumentException(Environment.GetResourceString(""Arg_RegSetBadArrType"", value.GetType().Name))"
    End If
    CalculateValueKind = RegistryValueKind.rvkString
End Function

''' <summary>Ruft eine Zeichenfolgenentsprechung dieses Schlüssels ab.</summary>
''' <returns>Eine Zeichenfolge, die den Schlüssel darstellt.Wenn der angegebene Schlüssel ungültig ist (nicht gefunden werden kann), wird null zurückgegeben.</returns>
''' <exceptioncref="T:System.ObjectDisposedException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, auf den zugegriffen werden soll, ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden). </exception>
'<SecuritySafeCritical>
Public Function ToString() As String
    EnsureNotDisposed
    ToString = m_keyName
End Function

''' <summary>Gibt die Zugriffssteuerungssicherheit für den aktuellen Registrierungsschlüssel zurück.</summary>
''' <returns>Ein Objekt, das die Zugriffssteuerungsberechtigungen für den durch den aktuellen <seecref="T:Microsoft.Win32.RegistryKey"/> dargestellten Registrierungsschlüssel beschreibt.</returns>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die notwendigen Berechtigungen.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden).</exception>
''' <exceptioncref="T:System.InvalidOperationException">Der aktuelle Schlüssel wurde gelöscht.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
Public Function GetAccessControl() As registrySecurity
    Return GetAccessControl(AccessControlSections.Access Or AccessControlSections.Owner Or AccessControlSections.Group)
End Function

''' <summary>Gibt die angegebenen Abschnitte der Zugriffssteuerungssicherheit für den aktuellen Registrierungsschlüssel zurück.</summary>
''' <returns>Ein Objekt, das die Zugriffssteuerungsberechtigungen für den durch den aktuellen <seecref="T:Microsoft.Win32.RegistryKey"/> dargestellten Registrierungsschlüssel beschreibt.</returns>
''' <paramname="includeSections">Eine bitweise Kombination von Enumerationswerten, die den Typ der abzurufenden Sicherheitsinformationen angeben. </param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die notwendigen Berechtigungen.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden).</exception>
''' <exceptioncref="T:System.InvalidOperationException">Der aktuelle Schlüssel wurde gelöscht.</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Function GetAccessControl(ByVal includeSections As AccessControlSections) As registrySecurity
    EnsureNotDisposed()
    Return New RegistrySecurity(hkey, m_keyName, includeSections)
End Function

''' <summary>Übernimmt Windows-Zugriffssteuerungssicherheit für einen vorhandenen Registrierungsschlüssel.</summary>
''' <paramname="registrySecurity">Die Zugriffssteuerungssicherheit, die für den aktuellen Unterschlüssel angewendet werden soll. </param>
''' <exceptioncref="T:System.UnauthorizedAccessException">Das aktuelle <seecref="T:Microsoft.Win32.RegistryKey"/>-Objekt stellt einen Schlüssel mit Zugriffssteuerungssicherheit dar, und der Aufrufer verfügt nicht über <seecref="F:System.Security.AccessControl.RegistryRights.ChangePermissions"/>-Rechte.</exception>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="registrySecurity"/> hat den Wert null.</exception>
''' <exceptioncref="T:System.ObjectDisposedException">Der zu bearbeitende <seecref="T:Microsoft.Win32.RegistryKey"/> ist geschlossen (auf geschlossene Schlüssel kann nicht zugegriffen werden).</exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
''' </PermissionSet>
'<SecuritySafeCritical>
Public Sub SetAccessControl(ByVal registrySecurity As registrySecurity)
    EnsureWriteable()
    If registrySecurity Is Nothing Then
        Throw New ArgumentNullException("registrySecurity")
    End If
    registrySecurity.Persist(hkey, m_keyName)
End Sub

'<SecuritySafeCritical>
Friend Sub Win32Error(ByVal errorCode As Long, ByVal str As String)
    Select Case errorCode
        Case 5
            If Not Equals(str, Nothing) Then
                Throw New UnauthorizedAccessException(Environment.GetResourceString("UnauthorizedAccess_RegistryKeyGeneric_Key", str))
            End If
            Throw New UnauthorizedAccessException()
        Case 6
            If Not IsPerfDataKey() Then
                hkey.SetHandleAsInvalid()
                m_hkey = Nothing
            End If
        Case 2
            Throw New IOException(Environment.GetResourceString("Arg_RegKeyNotFound"), errorCode)
    End Select
    Throw New IOException(Win32Native.GetMessage(errorCode), errorCode)
End Sub

'<SecuritySafeCritical>
Friend Sub Win32ErrorStatic(ByVal errorCode As Long, ByVal str As String)
    If errorCode = 5 Then
        If Not Equals(str, Nothing) Then
            Throw New UnauthorizedAccessException(Environment.GetResourceString("UnauthorizedAccess_RegistryKeyGeneric_Key", str))
        End If
        Throw New UnauthorizedAccessException()
    End If
    Throw New IOException(Win32Native.GetMessage(errorCode), errorCode)
End Sub

Friend Function FixupName(ByVal Name As String) As String
    If name.IndexOf("\"c) = -1 Then
        Return name
    End If
    Dim stringBuilder As StringBuilder = New StringBuilder(name)
    FixupPath (StringBuilder)
    Dim num As Long: num = StringBuilder.Length - 1
    If num >= 0 AndAlso stringBuilder(num) = "\"c Then
        StringBuilder.Length = num
    End If
    Return stringBuilder.ToString()
End Function

Private Sub FixupPath(ByVal Path As String) 'Builder)
    Dim Length As Long: Length = Len(Path) '.Length
    Dim flag As Boolean '= False
    Dim c As String: c = ChrW(65535)
    Dim i As Long
    For i = 1 To Length - 1 - 1
        If Mid(Path, i) = "\" Then
            i = i + 1
            While i < Length And Mid(Path, i) = "\"
                Mid(Path, i) = c
                i = i + 1
                flag = True
            Wend
        End If
    Next
    If Not flag Then
        Return
    End If
    i = 0
    Dim num As Long: num = 0
    While i < Length
        If Path(i) = c Then
            i += 1
            Continue While
        End If
        Path(num) = Path(i)
        i += 1
        num += 1
    Wend
    path.Length += num - i
End Sub

Private Sub GetSubKeyReadPermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Read
    Path = m_keyName & "\" & subKeyName & "\."
End Sub

Private Sub GetSubKeyWritePermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Write
    Path = m_keyName & "\" & subKeyName & "\."
End Sub

Private Sub GetSubKeyCreatePermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Create
    Path = m_keyName & "\" & subKeyName & "\."
End Sub

Private Sub GetSubTreeReadPermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Read
    Path = m_keyName & "\" & subKeyName & "\"
End Sub

Private Sub GetSubTreeWritePermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Write
    Path = m_keyName & "\" & subKeyName & "\"
End Sub

Private Sub GetSubTreeReadWritePermission(ByVal subKeyName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Read Or RegistryPermissionAccess.Write
    Path = m_keyName & "\" & subKeyName
End Sub

Private Sub GetValueReadPermission(ByVal valueName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Read
    Path = m_keyName & "\" & valueName
End Sub

Private Sub GetValueWritePermission(ByVal valueName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Write
    Path = m_keyName & "\" & valueName
End Sub

Private Sub GetValueCreatePermission(ByVal valueName As String, ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Create
    Path = m_keyName & "\" & valueName
End Sub

Private Sub GetKeyReadPermission(ByRef access_out As RegistryPermissionAccess, ByRef path_out As String)
    access = RegistryPermissionAccess.Read
    Path = m_keyName & "\."
End Sub

'<SecurityCritical>
Private Sub CheckPermission(ByVal check As RegistryInternalCheck, ByVal item As String, ByVal subkeyWritable As Boolean, ByVal subKeyCheck As RegistryKeyPermissionCheck)
    Dim flag As Boolean '= False
    Dim access As RegistryPermissionAccess: access = RegistryPermissionAccess.NoAccess
    Dim Path As String '= Nothing
    If CodeAccessSecurityEngine.QuickCheckForAllDemands() Then
        Exit Sub 'Return
    End If
    Select Case check
    Case RegistryInternalCheck.CheckSubKeyReadPermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
            'Exit Select
        End If
        flag = True
        GetSubKeyReadPermission item, access, Path
    Case RegistryInternalCheck.CheckSubKeyWritePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetSubKeyWritePermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckSubKeyCreatePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetSubKeyCreatePermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckSubTreeReadPermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetSubTreeReadPermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckSubTreeWritePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetSubTreeWritePermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckSubTreeReadWritePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
            'Exit Select
        End If
        flag = True
        GetSubTreeReadWritePermission item, access, Path
    Case RegistryInternalCheck.CheckValueReadPermission
        If checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetValueReadPermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckValueWritePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetValueWritePermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckValueCreatePermission
        If m_remoteKey Then
            CheckUnmanagedCodePermission
        ElseIf checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetValueCreatePermission item, access, Path
        End If
    Case RegistryInternalCheck.CheckKeyReadPermission
        If checkMode Is RegistryKeyPermissionCheck.[Default] Then
            flag = True
            GetKeyReadPermission access, Path
        End If
    Case RegistryInternalCheck.CheckSubTreePermission
        Select Case subKeyCheck
            Case RegistryKeyPermissionCheck.ReadSubTree
                If checkMode Is RegistryKeyPermissionCheck.[Default] Then
                    If m_remoteKey Then
                        CheckUnmanagedCodePermission
                        'Exit Select
                    End If
                    flag = True
                    GetSubTreeReadPermission item, access, Path
                End If
            Case RegistryKeyPermissionCheck.ReadWriteSubTree
                If checkMode <> RegistryKeyPermissionCheck.ReadWriteSubTree Then
                    If m_remoteKey Then
                        CheckUnmanagedCodePermission
                        'Exit Select
                    Else
                        flag = True
                        GetSubTreeReadWritePermission item, access, Path
                    End If
                End If
        End Select
    Case RegistryInternalCheck.CheckOpenSubKeyWithWritablePermission
        If checkMode Is RegistryKeyPermissionCheck.[Default] Then
            If m_remoteKey Then
                CheckUnmanagedCodePermission
                'Exit Select
            Else
                flag = True
                GetSubKeyReadPermission item, access, Path
            End If
        ElseIf subkeyWritable Then
            If checkMode = RegistryKeyPermissionCheck.ReadSubTree Then
                If m_remoteKey Then
                    CheckUnmanagedCodePermission
                    'Exit Select
                End If
            Else
                flag = True
                GetSubTreeReadWritePermission item, access, Path
            End If
        End If
    Case RegistryInternalCheck.CheckOpenSubKeyPermission
        If subKeyCheck = RegistryKeyPermissionCheck.[Default] Then
            If checkMode = RegistryKeyPermissionCheck.[Default] Then
                If m_remoteKey Then
                    CheckUnmanagedCodePermission
                    'Exit Select
                End If
            Else
                flag = True
                GetSubKeyReadPermission item, access, Path
            End If
        End If
    End Select
    If flag Then
        'Call New_RegistryPermission(access, path).Demand
    End If
End Sub

'<SecurityCritical>
Private Sub CheckUnmanagedCodePermission()
    'Call New_SecurityPermission(SecurityPermissionFlag.UnmanagedCode).Demand
End Sub

'<SecurityCritical>
Private Function ContainsRegistryValue(ByVal Name As String) As Boolean
    Dim lpType As LongPtr '= 0
    Dim lpcbData As LongPtr '= 0
    Dim num As Long: num = RegQueryValueExW(hKey, StrPtr(Name), 0, VarPtr(lpType), 0, VarPtr(lpcbData))
    ContainsRegistryValue = (num = ERROR_SUCCESS)
End Function

'<SecurityCritical>
Private Sub EnsureNotDisposed()
    If m_hkey = 0 Then
        'ThrowHelper.ThrowObjectDisposedException(keyName, ExceptionResource.ObjectDisposed_RegKeyClosed)
        Err.Raise 0, m_keyName & ": ThrowHelper.ThrowObjectDisposedException(keyName, ExceptionResource.ObjectDisposed_RegKeyClosed)"
    End If
End Sub

'<SecurityCritical>
Private Sub EnsureWriteable()
    EnsureNotDisposed
    If Not IsWritable() Then
        'ThrowHelper.ThrowUnauthorizedAccessException (ExceptionResource.UnauthorizedAccess_RegistryNoWrite)
        Err.Raise 0, m_keyName & ": ThrowHelper.ThrowUnauthorizedAccessException (ExceptionResource.UnauthorizedAccess_RegistryNoWrite)"
    End If
End Sub

'KEY_ALL_ACCESS: &HF003F = 983103
'KEY_READ      : &H20019 = 131097
'KEY_WRITE     : &H20006 = 131078
'KEY_EXECUTE   : &H20019 = 131097

'Private Function GetRegistryKeyAccess(IsWritable, mode_RegistryKeyPermissionCheck) As Long
'
'End Function

Private Function GetRegistryKeyAccess(ByVal IsWritable As Boolean) As Long
    If Not IsWritable Then
        GetRegistryKeyAccess = KEY_READ 'KEY_READ      : &H20019 = 131097
        Exit Function
    End If
    GetRegistryKeyAccess = KEY_READ Or KEY_WRITE ' 131103
End Function

Private Function GetRegistryKeyAccess2(ByVal mode As RegistryKeyPermissionCheck) As Long
    Dim result As Long ' result = 0
    Select Case mode
    Case RegistryKeyPermissionCheck.Default, RegistryKeyPermissionCheck.ReadSubTree
        result = KEY_READ '131097
    Case RegistryKeyPermissionCheck.ReadWriteSubTree
        result = KEY_READ Or KEY_WRITE ' 131103
    End Select
    GetRegistryKeyAccess = result
End Function

Private Function GetSubKeyPermissonCheck(ByVal subkeyWritable As Boolean) As RegistryKeyPermissionCheck
    If m_checkMode = RegistryKeyPermissionCheck.Default Then GetSubKeyPermissonCheck = m_checkMode: Exit Function
    If subkeyWritable Then GetSubKeyPermissonCheck = RegistryKeyPermissionCheck.ReadWriteSubTree: Exit Function
    GetSubKeyPermissonCheck = RegistryKeyPermissionCheck.ReadSubTree
End Function

Private Sub ValidateKeyName(ByVal Name As String)
    If Len(Name) = 0 Then
        'ThrowHelper.ThrowArgumentNullException (ExceptionArgument.Name)
        Err.Raise 0, "ThrowHelper.ThrowArgumentNullException (ExceptionArgument.Name)"
    End If
    Dim num As Long: num = MString.IndexOf(Name, "\") ', StringComparison.OrdinalIgnoreCase)
    Dim num2 As Long ': nums = 0
    While num <> -1
        If num - num2 > 255 Then
            'ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegKeyStrLenBug)
            Err.Raise &H7007, , "ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegKeyStrLenBug)"
        End If
        num2 = num + 1
        num = MString.IndexOf(Name, "\", num2) ', StringComparison.OrdinalIgnoreCase)
    Wend
    If Len(Name) - num2 > 255 Then
        'ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegKeyStrLenBug)
        Err.Raise &H7007, , "ThrowHelper.ThrowArgumentException (ExceptionResource.Arg_RegKeyStrLenBug)"
    End If
End Sub

Private Sub ValidateKeyMode(ByVal mode As RegistryKeyPermissionCheck)
    If mode < RegistryKeyPermissionCheck.Default Then
        If mode > RegistryKeyPermissionCheck.ReadWriteSubTree Then
            'ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryKeyPermissionCheck, ExceptionArgument.mode)
            Err.Raise &H7007, , "ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryKeyPermissionCheck, ExceptionArgument.mode)"
        End If
    End If
End Sub

Private Sub ValidateKeyOptions(ByVal Options As RegistryOptions)
    If Options < RegistryOptions.None Then
        If Options > RegistryOptions.Volatile Then
            'ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryOptionsCheck, ExceptionArgument.options)
            Err.Raise &H7007, , "ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryOptionsCheck, ExceptionArgument.options)"
        End If
    End If
End Sub

Private Sub ValidateKeyView(ByVal View As RegistryView)
    If View <> 0 Then
        If View <> RegistryView.Registry32 Then
            If View <> RegistryView.Registry64 Then
                'ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryViewCheck, ExceptionArgument.view)
                Err.Raise &H7007, , "ThrowHelper.ThrowArgumentException(ExceptionResource.Argument_InvalidRegistryViewCheck, ExceptionArgument.view)"
            End If
        End If
    End If
End Sub

Private Sub ValidateKeyRights(ByVal rights As Long)
    If (rights And &HFFF0FFC0) <> 0 Then
        'ThrowHelper.ThrowSecurityException (ExceptionResource.Security_RegistryPermission)
        Err.Raise &H7007, , "ThrowHelper.ThrowSecurityException (ExceptionResource.Security_RegistryPermission)"
    End If
End Sub

'Private Class CSharpImpl
'   <Obsolete("Please refactor calling code to use normal Visual Basic assignment")>
'   Function __Assign(Of T)(ByRef target As T, value As T) As T
'       target = value
'       Return value
'   End Function
'End Class
'End Class

'private Class SafeRegistryHandle
'Private Function SafeRegistryHandle_IsInvalid(Handle As LongPtr) As Boolean
Private Function SRH_IsInvalid(Handle As LongPtr) As Boolean
    SRH_IsInvalid = Handle = 0
End Function
Private Function SRH_ReleaseHandle(Handle As LongPtr) As Boolean
    SRH_ReleaseHandle = RegCloseKey(Handle) = ERROR_SUCCESS
End Function
Private Sub SRH_SetHandleAsInvalid(Handle As LongPtr)
    If SRH_ReleaseHandle(Handle) Then Handle = 0
End Sub

