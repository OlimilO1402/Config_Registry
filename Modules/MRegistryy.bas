Attribute VB_Name = "MRegistryy"
Option Explicit
'Imports System.Runtime.InteropServices
''' <summary>Enthält Informationen über die Einstellungen des aktuellen Benutzers.Dieses Feld liest den Basisschlüssel HKEY_CURRENT_USER der Windows-Registrierung.</summary>
Public CurrentUser As RegistryKey

''' <summary>Speichert die Konfigurationsinformationen für den lokalen Computer.Dieses Feld liest den Basisschlüssel HKEY_LOCAL_MACHINE der Windows-Registrierung.</summary>
Public LocalMachine As RegistryKey

''' <summary>Definiert die Typen (oder Klassen) von Dokumenten und die diesen Typen zugeordneten Eigenschaften.Dieses Feld liest den Basisschlüssel HKEY_CLASSES_ROOT der Windows-Registrierung.</summary>
Public ClassesRoot As RegistryKey

''' <summary>Enthält Informationen über die Standardkonfiguration des Benutzer.Dieses Feld liest den Basisschlüssel HKEY_USERS der Windows-Registrierung.</summary>
Public Users As RegistryKey

''' <summary>Enthält Leistungsdaten für Softwarekomponenten.Dieses Feld liest den Basisschlüssel HKEY_PERFORMANCE_DATA der Windows-Registrierung.</summary>
Public PerformanceData As RegistryKey

''' <summary>Enthält benutzerunabhängige Konfigurationsinformationen über die Hardware.Dieses Feld liest den Basisschlüssel HKEY_CURRENT_CONFIG der Windows-Registrierung.</summary>
Public CurrentConfig As RegistryKey

''' <summary>Speichert dynamische Registrierungsdaten.Dieses Feld liest den Basisschlüssel HKEY_DYN_DATA der Windows-Registrierung.</summary>
''' <exceptioncref="T:System.ObjectDisposedException">Das Betriebssystem unterstützt keine dynamischen Daten, d. h., es ist nicht Windows 98, Windows 98 Second Edition oder Windows Millennium Edition (Windows Me).</exception>
'<Obsolete("The DynData registry key only works on Win9x, which is no longer supported by the CLR.  On NT-based operating systems, use the PerformanceData registry key instead.")>
Public DynData As RegistryKey

Public Sub Init()
    CurrentUser = RegistryKey.GetBaseKey(RegistryKey.HKEY_CURRENT_USER)
    LocalMachine = RegistryKey.GetBaseKey(RegistryKey.HKEY_LOCAL_MACHINE)
    ClassesRoot = RegistryKey.GetBaseKey(RegistryKey.HKEY_CLASSES_ROOT)
    Users = RegistryKey.GetBaseKey(RegistryKey.HKEY_USERS)
    PerformanceData = RegistryKey.GetBaseKey(RegistryKey.HKEY_PERFORMANCE_DATA)
    CurrentConfig = RegistryKey.GetBaseKey(RegistryKey.HKEY_CURRENT_CONFIG)
    DynData = RegistryKey.GetBaseKey(RegistryKey.HKEY_DYN_DATA)
End Sub

''' [SecurityCritical]
Private Function GetBaseKeyFromKeyName(ByVal keyName As String, ByRef subKeyName_Out As String) As RegistryKey
    If Len(keyName) = 0 Then
        'Throw New ArgumentNullException("keyName")
    End If
    Dim num As Long: num = MString.IndexOf(keyName, "\")
    Dim text As String: text = IIf((num = -1), LCase(keyName), UCase(MString.Substring(keyName, num)))  '.Substring(0, num).ToUpper(CultureInfo.InvariantCulture));
    Dim aRegistryKey As RegistryKey '= Nothing
    Select Case text
    Case "HKEY_CURRENT_USER":     Set aRegistryKey = Registry.CurrentUser
    Case "HKEY_LOCAL_MACHINE":    Set aRegistryKey = Registry.LocalMachine
    Case "HKEY_CLASSES_ROOT":     Set aRegistryKey = Registry.ClassesRoot
    Case "HKEY_USERS":            Set aRegistryKey = Registry.Users
    Case "HKEY_PERFORMANCE_DATA": Set aRegistryKey = Registry.PerformanceData
    Case "HKEY_CURRENT_CONFIG":   Set aRegistryKey = Registry.CurrentConfig
    Case "HKEY_DYN_DATA":         Set aRegistryKey = RegistryKey.GetBaseKey(RegistryKey.HKEY_DYN_DATA)
    Case Else 'throw new ArgumentException(Environment.GetResourceString("Arg_RegInvalidKeyName", "keyName"))
    End Select
    If num = -1 Then
        If num = keyName.Length Then subKeyName = vbNullString 'String.Empty
    Else
        subKeyName = keyName.Substring(num + 1, keyName.Length - num - 1)
    End If
    Set GetBaseKeyFromKeyName = aRegistryKey
End Function

''' <summary>Ruft den Wert ab, der dem angegebenen Namen im angegebenen Registrierungsschlüssel zugeordnet ist.Wenn der Name im angegebenen Schlüssel nicht gefunden wird, wird ein von Ihnen bereitgestellter Standardwert zurückgegeben, oder null, wenn der angegebene Schlüssel nicht vorhanden ist.</summary>
''' <returns>null, wenn der durch <paramrefname="keyName"/> angegebene Unterschlüssel nicht vorhanden ist, andernfalls der Wert, der <paramrefname="valueName"/> zugeordnet ist, oder <paramrefname="defaultValue"/>, wenn <paramrefname="valueName"/> nicht gefunden wurde.</returns>
''' <paramname="keyName">Der vollständige Registrierungspfad des Schlüssels, beginnend mit einem gültigen Registrierungsstamm (z. B. "HKEY_CURRENT_USER").</param>
''' <paramname="valueName">Der Name des Name-/Wert-Paars.</param>
''' <paramname="defaultValue">Der zurückzugebende Wert, wenn <paramrefname="valueName"/> nicht vorhanden ist.</param>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen, um aus dem Registrierungsschlüssel zu lesen. </exception>
''' <exceptioncref="T:System.IO.IOException">Der <seecref="T:Microsoft.Win32.RegistryKey"/>, der den angegebenen Wert enthält, wurde zum Löschen markiert. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="keyName"/> beginnt nicht mit einem gültigen Registrierungsstamm. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Read="\"/>
''' </PermissionSet>
''' [SecuritySafeCritical]
Public Function GetValue(ByVal keyName As String, ByVal valueName As String, ByVal defaultValue) 'As Object
    Dim subKeyName As String
    Dim baseKeyFromKeyName As RegistryKey: Set baseKeyFromKeyName = GetBaseKeyFromKeyName(keyName, subKeyName)
    Dim aRegistryKey As RegistryKey: Set aRegistryKey = baseKeyFromKeyName.OpenSubKey(subKeyName)
    If aRegistryKey Is Nothing Then
        'Return Nothing
        Exit Function
    End If
Try: On Error GoTo Finally
    GetValue = RegistryKey.GetValue(valueName, defaultValue)
Finally:
    RegistryKey.CClose
End Function

''' <summary>Legt das angegebene Name-/Wert-Paar für den angegebenen Registrierungsschlüssel fest.Wenn der angegebene Schlüssel nicht vorhanden ist, wird er erstellt.</summary>
''' <paramname="keyName">Der vollständige Registrierungspfad des Schlüssels, beginnend mit einem gültigen Registrierungsstamm (z. B. "HKEY_CURRENT_USER").</param>
''' <paramname="valueName">Der Name des Name-/Wert-Paars.</param>
''' <paramname="value">Der zu speichernde Wert.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="value"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="keyName"/> beginnt nicht mit einem gültigen Registrierungsstamm. - oder -<paramrefname="keyName"/> überschreitet die maximal zulässige Länge (255 Zeichen).</exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Die <seecref="T:Microsoft.Win32.RegistryKey"/>-Klasse ist schreibgeschützt. Es ist kein Schreibzugriff möglich, d. h., es handelt sich z. B. um einen Knoten auf Stammebene. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Ändern von Registrierungsschlüsseln. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
'Public Sub SetValue(ByVal keyName As String, ByVal valueName As String, ByVal value)
'    SetValue2 keyName, valueName, value, RegistryValueKind.Unknown
'End Sub

''' <summary>Legt unter Verwendung des angegebenen Registrierungsdatentyps das Name-/Wert-Paar für den angegebenen Registrierungsschlüssel fest.Wenn der angegebene Schlüssel nicht vorhanden ist, wird er erstellt.</summary>
''' <paramname="keyName">Der vollständige Registrierungspfad des Schlüssels, beginnend mit einem gültigen Registrierungsstamm (z. B. "HKEY_CURRENT_USER").</param>
''' <paramname="valueName">Der Name des Name-/Wert-Paars.</param>
''' <paramname="value">Der zu speichernde Wert.</param>
''' <paramname="valueKind">Der beim Speichern der Daten zu verwendende Registrierungsdatentyp.</param>
''' <exceptioncref="T:System.ArgumentNullException">
'''   <paramrefname="value"/> hat den Wert null. </exception>
''' <exceptioncref="T:System.ArgumentException">
'''   <paramrefname="keyName"/> beginnt nicht mit einem gültigen Registrierungsstamm.- oder -<paramrefname="keyName"/> überschreitet die maximal zulässige Länge (255 Zeichen).- oder - Der Typ von <paramrefname="value"/> stimmt nicht mit dem durch <paramrefname="valueKind"/> angegebenen Registrierungsdatentyp überein. Die Daten konnten daher nicht ordnungsgemäß konvertiert werden. </exception>
''' <exceptioncref="T:System.UnauthorizedAccessException">Der <seecref="T:Microsoft.Win32.RegistryKey"/> ist schreibgeschützt. Es ist kein Schreibzugriff möglich, d. h. es handelt sich z. B. um einen Knoten auf Stammebene oder um einen Schlüssel, der nicht mit Schreibzugriff geöffnet wurde. </exception>
''' <exceptioncref="T:System.Security.SecurityException">Der Benutzer verfügt nicht über die erforderlichen Berechtigungen zum Erstellen oder Ändern von Registrierungsschlüsseln. </exception>
''' <PermissionSet>
'''   <IPermissionclass="System.Security.Permissions.RegistryPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Unrestricted="true"/>
'''   <IPermissionclass="System.Security.Permissions.SecurityPermission, mscorlib, Version=2.0.3600.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"version="1"Flags="UnmanagedCode"/>
''' </PermissionSet>
''' [SecuritySafeCritical]
Public Sub SetValue(ByVal keyName As String, ByVal valueName As String, ByVal value, Optional ByVal valueKind As RegistryValueKind = RegistryValueKind.Unknown)
    Dim subKeyName As String
    Dim baseKeyFromKeyName As RegistryKey: Set baseKeyFromKeyName = GetBaseKeyFromKeyName(keyName, subKeyName)
    Dim aRegistryKey As RegistryKey: Set aRegistryKey = baseKeyFromKeyName.CreateSubKey(subKeyName)
Try: On Error GoTo Finally
    aRegistryKey.SetValue valueName, value, valueKind
Finally:
    RegistryKey.CClose
End Sub

