Attribute VB_Name = "modRegistry"
Option Explicit

Public Enum REG_VALUE_TYPE
    REG_NONE = 0&
    REG_SZ = 1&
    REG_EXPAND_SZ = 2&
    REG_BINARY = 3&
    REG_DWORD = 4&
    REG_DWORDLittleEndian = 4&
    REG_DWORDBigEndian = 5&
    REG_LINK = 6&
    REG_MULTI_SZ = 7&
    REG_ResourceList = 8&
    REG_FullResourceDescriptor = 9&
    REG_ResourceRequirementsList = 10&
    REG_QWORD = 11&
    REG_QWORD_LITTLE_ENDIAN = 11&
End Enum

Public Enum FLAG_REG_TYPE   'flags to be able to map bit mask and default registry type constants
    FLAG_REG_ALL = -1&
    FLAG_REG_NONE = 1&
    FLAG_REG_SZ = 2&
    FLAG_REG_EXPAND_SZ = 4&
    FLAG_REG_BINARY = 8&
    FLAG_REG_DWORD = &H10&
    FLAG_REG_DWORDLittleEndian = &H10&
    FLAG_REG_DWORDBigEndian = &H20&
    FLAG_REG_LINK = &H40&
    FLAG_REG_MULTI_SZ = &H80&
    FLAG_REG_ResourceList = &H100&
    FLAG_REG_FullResourceDescriptor = &H200&
    FLAG_REG_ResourceRequirementsList = &H400&
    FLAG_REG_QWORD = &H800&
    FLAG_REG_QWORD_LITTLE_ENDIAN = &H1000&
End Enum

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyW" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Any, lpcbData As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExByte Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Byte, ByRef lpcbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long

Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_MORE_DATA      As Long = 234&
Private Const KEY_QUERY_VALUE      As Long = &H1


Function GetRegData(hHive As REG_HIVES, ByVal KeyName As String, ByVal ValueName As String) As Variant
    On Error GoTo ErrorHandler
    Dim abData()     As Byte
    Dim cData        As Long
    Dim hKey         As Long
    Dim lData        As Long
    Dim lret         As Long
    Dim ordType      As Long
    Dim iPos         As Long
    Dim sData        As String
    Dim vValue       As Variant
    
    If ERROR_SUCCESS <> RegOpenKeyEx(hHive, StrPtr(KeyName), 0&, KEY_QUERY_VALUE, hKey) Then Exit Function
    lret = RegQueryValueExLong(hKey, StrPtr(ValueName), 0&, ordType, 0&, cData)

    If ERROR_SUCCESS <> lret And ERROR_MORE_DATA <> lret Then Exit Function
    
    Select Case ordType
        
        Case REG_DWORD, REG_DWORDLittleEndian
            lret = RegQueryValueExLong(hKey, StrPtr(ValueName), 0&, ordType, lData, cData)
            vValue = lData
        
        Case REG_SZ, REG_MULTI_SZ
            If cData > 1 Then
                sData = String$(cData - 1&, 0&)
                lret = RegQueryValueExStr(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                vValue = Left$(sData, lstrlen(StrPtr(sData)))
            End If
        
        Case REG_EXPAND_SZ
            If cData > 1 Then
                sData = String$(cData - 1&, 0&)
                lret = RegQueryValueExStr(hKey, StrPtr(ValueName), 0&, ordType, StrPtr(sData), cData)
                vValue = ExpandEnvStr(sData)
            End If
    
    End Select
    GetRegData = vValue
ErrorHandler:
    If hKey <> 0& Then RegCloseKey hKey
End Function

Private Function ExpandEnvStr(sData As String) As String
    Dim lret     As Long
    Dim sTemp    As String
    lret = ExpandEnvironmentStrings(StrPtr(sData), StrPtr(sTemp), lret) 'get buffer size needed
    sTemp = Space(lret - 1&)
    lret = ExpandEnvironmentStrings(StrPtr(sData), StrPtr(sTemp), lret)
    If lret Then
        ExpandEnvStr = Left$(sTemp, lret - 1&)
    Else
        ExpandEnvStr = sData
    End If
End Function
