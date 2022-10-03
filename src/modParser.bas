Attribute VB_Name = "modParser"
Option Explicit

Private Const MAX_PATH As Long = 260&
Private Const MAX_PATH_W As Long = 32767&

Public Enum idCodePage
    WIN = 1251
    Dos = 866
    Koi = 20866
    Iso = 28595
    UTF8 = 65001
End Enum

Public Type Environment_Variables
    SysDisk      As String
    SysRoot      As String
    PF_64        As String
    PF_32        As String
    PF_64_Common As String
    PF_32_Common As String
    AppData      As String
    LocalAppData As String
    System32     As String
    SysWow64     As String
    Desktop      As String
    StartMenu    As String
    UserProfile  As String
    sWinDir      As String
    sWinSysDir   As String
    TempCU       As String
End Type

Public OSver    As clsOSInfo
Public Proc     As clsProcess
Public Reg      As clsRegistry
'Public MD5      As clsMD5
Public Env      As Environment_Variables

Enum REG_HIVES
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathW" (ByVal hWndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function PathFindOnPath Lib "Shlwapi" Alias "PathFindOnPathW" (ByVal pszFile As Long, ppszOtherDirs As Any) As Boolean
Private Declare Function AssocQueryString Lib "Shlwapi.dll" Alias "AssocQueryStringW" (ByVal flags As Long, ByVal str As Long, ByVal pszAssoc As Long, ByVal pszExtra As Long, ByVal pszOut As Long, pcchOut As Long) As Long

Public Function AppPath(Optional bGetFullPath As Boolean) As String
    On Error GoTo ErrorHandler

    Static ProcPathFull  As String
    Static ProcPathShort As String
    Dim ProcPath As String
    Dim cnt      As Long
    Dim hProc    As Long
    Dim pos      As Long
    
    'Cache
    If bGetFullPath Then
        If Len(ProcPathFull) <> 0 Then
            AppPath = ProcPathFull
            Exit Function
        End If
    Else
        If Len(ProcPathShort) <> 0 Then
            AppPath = ProcPathShort
            Exit Function
        End If
    End If

    If inIDE Then
        AppPath = GetDOSFilename(App.Path, bReverse:=True)
        'bGetFullPath does not supported in IDE
        Exit Function
    End If

    hProc = GetModuleHandle(0&)
    If hProc < 0 Then hProc = 0

    ProcPath = String$(MAX_PATH, vbNullChar)
    cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath)) 'hproc can be 0 (mean - current process)
    
    If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
        ProcPath = Space$(MAX_PATH_W)
        cnt = GetModuleFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
    End If
    
    If cnt = 0 Then                          'clear path
        ProcPath = App.Path
    Else
        ProcPath = Left$(ProcPath, cnt)
        If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = Env.sWinDir & Mid$(ProcPath, 12)
        If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        
        If Not bGetFullPath Then
            ' trim to path
            pos = InStrRev(ProcPath, "\")
            If pos <> 0 Then ProcPath = Left$(ProcPath, pos - 1)
        End If
    End If
    
    ProcPath = GetDOSFilename(ProcPath, bReverse:=True)     '8.3 -> to Full
    
    AppPath = ProcPath
    
    If bGetFullPath Then
        ProcPathFull = ProcPath
    Else
        ProcPathShort = ProcPath
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.AppPath", "ProcPath:", ProcPath
    If inIDE Then Stop: Resume Next
End Function

'if short name is unavailable, it returns source string anyway
Public Function GetDOSFilename$(sFile$, Optional bReverse As Boolean = False)
    'works for folders too btw
    Dim cnt&, sBuffer$
    If bReverse Then
        sBuffer = Space$(MAX_PATH_W)
        cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If cnt Then
            GetDOSFilename = Left$(sBuffer, cnt)
        Else
            GetDOSFilename = sFile
        End If
    Else
        sBuffer = Space$(MAX_PATH)
        cnt = GetShortPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
        If cnt Then
            GetDOSFilename = Left$(sBuffer, cnt)
        Else
            GetDOSFilename = sFile
        End If
    End If
End Function

Public Function BuildPath$(sPath$, sFile$)
    BuildPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\") & sFile
End Function

Public Function RTrimNull(pStr As String) As String
    Dim pos&
    pos = InStr(pStr, vbNullChar)
    If pos <> 0 Then
        RTrimNull = Left$(pStr, pos - 1)
    Else
        RTrimNull = pStr
    End If
End Function

Public Function EnvironW(ByVal SrcEnv As String) As String
    Dim lr As Long
    Dim buf As String
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        'redirector correction
        If OSver.IsWin64 Then
            If InStr(1, SrcEnv, "%PROGRAMFILES%", 1) <> 0 Then
                SrcEnv = Replace$(SrcEnv, "%PROGRAMFILES%", Env.PF_64, 1, 1, 1)
            End If
            If InStr(1, SrcEnv, "%COMMONPROGRAMFILES%", 1) <> 0 Then
                SrcEnv = Replace$(SrcEnv, "%COMMONPROGRAMFILES%", Env.PF_64_Common, 1, 1, 1)
            End If
        End If
        buf = String$(MAX_PATH, vbNullChar)
        lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), MAX_PATH + 1)
        
        If lr Then
            EnvironW = Left$(buf, lr - 1)
        Else
            EnvironW = SrcEnv
        End If
        
        If InStr(EnvironW, "%") <> 0 Then
            If OSver.MajorMinor <= 6 Then
                If InStr(1, EnvironW, "%ProgramW6432%", 1) <> 0 Then
                    EnvironW = Replace$(EnvironW, "%ProgramW6432%", Env.SysDisk & "\Program Files", 1, -1, 1)
                End If
            End If
        End If
    End If
End Function

' Заполнение переменных окружения и объектов WSH
Public Sub InitVariables()
  On Error GoTo ErrorHandler
    
  Const CSIDL_DESKTOP               As Long = 0
  Const CSIDL_STARTMENU             As Long = 11
  Const CSIDL_LOCAL_APPDATA         As Long = 28&

  Set OSver = New clsOSInfo
  Set Reg = New clsRegistry
  'Set MD5 = New clsMD5
  Set Proc = New clsProcess
    
  ' Раскрытие переменных окружения
  With Env

    Dim lr As Long
    .SysDisk = String$(MAX_PATH, 0)
    lr = GetWindowsDirectory(StrPtr(.SysDisk), MAX_PATH)
    If lr Then
        .SysRoot = Left$(.SysDisk, lr)
        .SysDisk = Left$(.SysDisk, 2)
    Else
        .SysDisk = EnvironW("%SystemDrive%")
        .SysRoot = EnvironW("%SystemRoot%")
    End If
    
    .sWinDir = .SysRoot
    
    If OSver.IsWin64 Then
        ' Do not use Environ(W) here !!!
        If OSver.MajorMinor >= 6.1 Then     'Win 7 and later
            .PF_64 = Environ("ProgramW6432")
        Else
            .PF_64 = .SysDisk & "\Program Files"
        End If
        .PF_32 = Environ("ProgramFiles")
    Else
        .PF_32 = Environ("ProgramFiles")
        .PF_64 = .PF_32
    End If
    
    .PF_64_Common = .PF_64 & "\Common Files"
    .PF_32_Common = .PF_32 & "\Common Files"
    
    .AppData = EnvironW("%AppData%")
    If OSver.IsWindowsVistaOrGreater Then
        .LocalAppData = EnvironW("%LocalAppData%")
    Else
        .LocalAppData = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)
        If Len(.LocalAppData) = 0 Then .LocalAppData = EnvironW("%USERPROFILE%") & "\Local Settings\Application Data"
    End If
    .System32 = .SysRoot & "\System32"
    .sWinSysDir = .System32
    .SysWow64 = .SysRoot & "\SysWOW64"
    .Desktop = GetSpecialFolderPath(CSIDL_DESKTOP)
    .StartMenu = GetSpecialFolderPath(CSIDL_STARTMENU)
    .UserProfile = EnvironW("%UserProfile%")
    
    .TempCU = Environ("temp")
    ' if REG_EXPAND_SZ is missing
    If InStr(.TempCU, "%") <> 0 Then
        If OSver.IsWindowsVistaOrGreater Then
            .TempCU = .UserProfile & "\Local\Temp"
        Else
            .TempCU = .UserProfile & "\Local Settings\Temp"
        End If
    End If
    
  End With
    
  Exit Sub
ErrorHandler:
    ErrorMsg Err, "Engine.InitVariables"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetSpecialFolderPath(CSIDL As Long, Optional hToken As Long = 0&) As String
    On Error GoTo ErrorHandler
    'https://msdn.microsoft.com/en-us/library/windows/desktop/bb762494.aspx
    Const SHGFP_TYPE_CURRENT As Long = &H0&
    Const SHGFP_TYPE_DEFAULT As Long = &H1&
    Dim lr      As Long
    Dim sPath   As String
    sPath = String$(MAX_PATH, vbNullChar)
    ' 3-th parameter - is a token of user (user registry hive must be loaded first! )
    lr = SHGetFolderPath(0&, CSIDL, hToken, SHGFP_TYPE_CURRENT, StrPtr(sPath))
    If lr = 0 Then GetSpecialFolderPath = Left$(sPath, lstrlen(StrPtr(sPath)))
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Engine.GetSpecialFolderPath", "CSIDL:", CSIDL
End Function

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, LastPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(LastPart)), LastPart, 1) = 0)
End Function

Public Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ") As String()
    If 0 = Len(sComplexString) Then
        ReDim ret(0) As String
        SplitSafe = ret
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

' Возвращает true, если искомое значение найдено в одном из элементов массива (lB, uB ограничивает просматриваемый диапазон индексов)
Public Function inArray( _
    Stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to Казанский :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetParentDir(sPath As String) As String
    Dim pos As Long
    pos = InStrRev(sPath, "\")
    If pos <> 0 Then
        GetParentDir = Left$(sPath, pos - 1)
    End If
End Function

Public Function TrimNull(S$) As String
    TrimNull = Left$(S, lstrlen(StrPtr(S)))
End Function

Public Sub ReleaseVariables()
    Set Proc = Nothing
    Set Reg = Nothing
    Set OSver = Nothing
End Sub

Public Function ConvertCodePageW(Src As String, inPage As idCodePage) As String
    On Error GoTo ErrorHandler
    
    Dim buf   As String
    Dim Size  As Long
    
    Size = MultiByteToWideChar(inPage, 0&, Src, Len(Src), 0&, 0&)
    If Size > 0 Then
        buf = String$(Size, 0)
        Size = MultiByteToWideChar(inPage, 0&, Src, Len(Src), StrPtr(buf), Len(Src))
    
        If Size <> 0 Then ConvertCodePageW = Left$(buf, Size)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.ConvertCodePageW", "String:", Src
End Function

Public Function ConvertCodePage(ByVal Src As String, inPage As idCodePage, outPage As idCodePage) As String
    On Error GoTo ErrorHandler
    
    Dim buf   As String
    Dim Dst   As String
    Dim Size  As Long
    
    Size = MultiByteToWideChar(inPage, 0&, Src, Len(Src), 0&, 0&)
    If Size > 0 Then
        buf = String$(Size, 0)
        Size = MultiByteToWideChar(inPage, 0&, Src, Len(Src), StrPtr(buf), Len(Src))
    
        Size = WideCharToMultiByte(outPage, 0&, StrPtr(buf), Size, ByVal 0&, 0&, 0&, 0&)
        If Size > 0 Then
            Dst = String$(Size, 0)
            Size = WideCharToMultiByte(outPage, 0&, StrPtr(buf), Size, Dst, LenB(Dst), 0&, 0&)
            
            If Size <> 0 Then ConvertCodePage = Left$(Dst, lstrlen(StrPtr(Dst)))
        End If
    End If

    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.ConvertCodePage", "String:", Src
End Function

Public Function ReadIniValue(sPath As String, sSection As String, sParam As String) As String
    Dim buf As String
    Dim lr As Long
    
    buf = String$(256&, 0)
    lr = GetPrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(""), StrPtr(buf), Len(buf), StrPtr(sPath))
    
    If Err.LastDllError = 0 Then
        ReadIniValue = Left$(buf, lr)
    End If
End Function

Public Function WriteIniValue(sPath As String, sSection As String, sParam As String, vData As Variant) As Long
    WriteIniValue = WritePrivateProfileString(StrPtr(sSection), StrPtr(sParam), StrPtr(CStr(vData)), StrPtr(sPath))
End Function

Public Function AddToArray(str As String, arr() As String)
    If 0 = Len(arr(UBound(arr))) And UBound(arr) = 0 Then
        arr(0) = str
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        arr(UBound(arr)) = str
    End If
End Function

Public Function isL4dDir(sPath As String) As Boolean
    If 0 <> Len(sPath) Then
        If FileExists(sPath & "\left4dead\missions\hospital.txt") Then
            isL4dDir = True
        End If
    End If
End Function

Public Function GetCollectionKeyByIndex(ByVal Index As Long, Col As Collection) As String ' Thanks to 'The Trick' (А. Кривоус) for this code
    'Fixed by Dragokas
    On Error GoTo ErrorHandler:
    Dim lpSTR As Long, ptr As Long, Key As String
    If Col Is Nothing Then Exit Function
    Select Case Index
    Case Is < 1, Is > Col.Count: Exit Function
    Case Else
        ptr = ObjPtr(Col)
        Do While Index
            GetMem4 ByVal ptr + 24, ptr
            Index = Index - 1
        Loop
    End Select
    GetMem4 ByVal VarPtr(Key), lpSTR
    GetMem4 ByVal ptr + 16, ByVal VarPtr(Key)
    GetCollectionKeyByIndex = Key
    GetMem4 lpSTR, ByVal VarPtr(Key)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetCollectionKeyByIndex"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetCollectionIndexByItem(sItem As String, Col As Collection) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, 1) = 0 Then
            GetCollectionIndexByItem = i
            Exit For
        End If
    Next
End Function

Public Function GetCollectionKeyByItem(sItem As String, Col As Collection) As String
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(Col.Item(i), sItem, 1) = 0 Then
            GetCollectionKeyByItem = GetCollectionKeyByIndex(i, Col)
            Exit For
        End If
    Next
End Function

Public Function isCollectionKeyExists(Key As String, Col As Collection) As Boolean
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, 1) = 0 Then isCollectionKeyExists = True: Exit For
    Next
End Function

Public Function GetCollectionKeyByItemName(Key As String, Col As Collection) As String
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, 1) = 0 Then GetCollectionKeyByItemName = Col.Item(i)
    Next
End Function

Public Function GetCollectionIndexByKey(Key As String, Col As Collection) As Long
    Dim i As Long
    For i = 1 To Col.Count
        If StrComp(GetCollectionKeyByIndex(i, Col), Key, 1) = 0 Then GetCollectionIndexByKey = i
    Next
End Function

Public Function GetPlayerPath() As String
    Dim TXTProg     As String
    Dim TXTClassID  As String
    Dim sFile       As String
    Dim sVerb       As String
    
    TXTClassID = GetRegData(HKEY_CLASSES_ROOT, ".mp3", "")
    
    If TXTClassID <> "" Then
        sVerb = GetRegData(HKEY_CLASSES_ROOT, TXTClassID & "\shell", "")
        
        If sVerb = "" Then sVerb = "open"
    
        TXTProg = EnvironW(GetRegData(HKEY_CLASSES_ROOT, TXTClassID & "\shell\" & sVerb & "\command", ""))
        
        SplitIntoPathAndArgs TXTProg, sFile, , True
        TXTProg = sFile
    End If
    
    If Not FileExists(TXTProg) Then
        TXTProg = "rundll32.exe shell32,ShellExec_RunDLL"
    End If
    GetPlayerPath = TXTProg
End Function

Public Function GetPlayerPath2() As String
    Dim sRegData As String
    Dim sFile As String
    
    sRegData = GetDefaultApp(".MP3")
    
    If sRegData <> "" Then
        SplitIntoPathAndArgs sRegData, sFile, , True
    End If
        
    GetPlayerPath2 = sFile
End Function

Function GetDefaultApp(Protocol As String, Optional out_FriendlyName As String) As String
    On Error GoTo ErrorHandler
    
    Const ASSOCF_INIT_FIXED_PROGID  As Long = 2048
    Const ASSOCF_IS_PROTOCOL        As Long = 4096
    Const ASSOCF_INIT_FOR_FILE      As Long = 8192
    Const ASSOCF_INIT_BYEXENAME     As Long = 2
    Const ASSOCF_INIT_NOREMAPCLSID  As Long = 1
    Const ASSOCF_NOFIXUPS           As Long = &H100
    Const ASSOCF_INIT_IGNOREUNKNOWN As Long = &H400&
    Const ASSOCF_OPEN_BYEXENAME     As Long = 2&
    Const ASSOCF_INIT_DEFAULTTOSTAR As Long = 4&
    
    Const ASSOCSTR_EXECUTABLE       As Long = 2&
    Const ASSOCSTR_FRIENDLYAPPNAME  As Long = 4&
    Const ASSOCSTR_COMMAND          As Long = 1&
    
    Dim HRes    As Long
    Dim buf     As String
    Dim Size    As Long
    Dim buf2    As String
    
    buf = String$(MAX_PATH, vbNullChar)
    Size = MAX_PATH
    
    HRes = AssocQueryString(ASSOCF_INIT_DEFAULTTOSTAR Or ASSOCF_NOFIXUPS, _
        ASSOCSTR_COMMAND, StrPtr(Protocol), 0&, StrPtr(buf), Size)         'StrPtr("Open"),

    If HRes = 0 And Size <> 0 Then
        buf = Left$(buf, Size - 1)
        buf2 = String$(MAX_PATH_W, vbNullChar)
        Size = GetLongPathName(StrPtr(buf), StrPtr(buf2), Len(buf2))
        If Size <> 0 Then
            GetDefaultApp = Left$(buf2, Size)
        Else
            GetDefaultApp = buf
        End If
    Else
        If HRes = -2147023741 Then      'AL_USER
            GetDefaultApp = "(AppID)"
        Else
            GetDefaultApp = "?"
            'Err.Raise 76
        End If
    End If
    
    buf = String$(MAX_PATH, vbNullChar)
    Size = MAX_PATH
    
    HRes = AssocQueryString(ASSOCF_INIT_DEFAULTTOSTAR Or ASSOCF_NOFIXUPS, _
        ASSOCSTR_FRIENDLYAPPNAME, StrPtr(Protocol), 0&, StrPtr(buf), Size)
    
    If HRes = 0 And Size <> 0 Then
        buf = Left$(buf, Size - 1)
        out_FriendlyName = buf
    End If
    
    If 0 = Len(out_FriendlyName) And "(AppID)" = GetDefaultApp Or "?" = GetDefaultApp Then
        GetDefaultApp = ""
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Engine.GetDefaultApp", "Result=", HRes, "Protocol:", Protocol
    If inIDE Then Stop: Resume Next
End Function

Public Sub SplitIntoPathAndArgs(ByVal InLine As String, Path As String, Optional Args As String, Optional bIsRegistryData As Boolean)
    On Error GoTo ErrorHandler
    Dim pos As Long
    Dim sTmp As String
    Dim bFail As Boolean
    
    Path = vbNullString
    Args = vbNullString
    If Len(InLine) = 0& Then Exit Sub
    
    InLine = Trim(InLine)
    If Left$(InLine, 1) = """" Then
        pos = InStr(2, InLine, """")
        If pos <> 0 Then
            Path = Mid$(InLine, 2, pos - 2)
            Args = Trim(Mid$(InLine, pos + 1))
        Else
            Path = Mid$(InLine, 2)
        End If
    Else
        '//TODO: Check correct system behaviour: maybe it uses number of 'space' characters, like, if more than 1 'space', exec bIsRegistryData routine.
    
        If bIsRegistryData Then
            'Expanding paths like: C:\Program Files (x86)\Download Master\dmaster.exe -autorun
            pos = InStrRev(InLine, ".exe", -1, 1)
            If pos <> 0 Then
                Path = Left$(InLine, pos + 3)
                If Not FileExists(Path) Then bFail = True
            End If
        Else
            bFail = True
        End If
        
        If bFail Or Len(Path) = 0 Then
            pos = InStr(InLine, " ")
            If pos <> 0 Then
                Path = Left$(InLine, pos - 1)
                Args = Mid$(InLine, pos + 1)
            Else
                Path = InLine
            End If
        End If
    End If
    If Len(Path) <> 0 Then
        If Not FileExists(Path) Then  'find on %PATH%
            sTmp = FindOnPath(Path)
            If Len(sTmp) <> 0 Then
                Path = sTmp
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "Parser.SplitIntoPathAndArgs", "In Line:", InLine
    If inIDE Then Stop: Resume Next
End Sub

Public Function FindOnPath(sAppName As String, Optional bUseDefaultOnFailure As Boolean) As String
    On Error GoTo ErrorHandler:

    Static Exts
    Static isinit As Boolean
    Dim ProcPath$
    Dim sFile As String
    Dim sFolder As String
    Dim pos As Long
    Dim i As Long
    Dim FoundFile As Boolean
    Dim sFileTry As String
    Dim bFullPath As Boolean
    
    If Not isinit Then
        isinit = True
        Exts = Split(EnvironW("%PathExt%"), ";")
        For i = 0 To UBound(Exts)
            Exts(i) = LCase(Exts(i))
        Next
    End If
    
    If Mid(sAppName, 2, 1) = ":" Then bFullPath = True
    
    If bFullPath Then
        If FileExists(sAppName) Then
            FindOnPath = sAppName
            Exit Function
        End If
    End If
    
    pos = InStrRev(sAppName, "\")
    
    If bFullPath And pos <> 0 Then
        sFolder = Left$(sAppName, pos - 1)
        sFile = Mid$(sAppName, pos + 1)
        
        For i = 0 To UBound(Exts)
            sFileTry = sFolder & "\" & sFile & Exts(i)
            
            If FileExists(sFileTry) Then
                FindOnPath = sFileTry
                Exit Function
            End If
        Next
    Else
        ToggleWow64FSRedirection False
    
        ProcPath = Space$(MAX_PATH)
        LSet ProcPath = sAppName & vbNullChar
        
        If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
            FindOnPath = TrimNull(ProcPath)
        Else
            'go through the extensions list
            
            For i = 0 To UBound(Exts)
                sFileTry = sAppName & Exts(i)
            
                ProcPath = Space$(MAX_PATH)
                LSet ProcPath = sFileTry & vbNullChar
            
                If CBool(PathFindOnPath(StrPtr(ProcPath), 0&)) Then
                    FindOnPath = TrimNull(ProcPath)
                    Exit For
                End If
            
            Next
            
        End If
        
        ToggleWow64FSRedirection True
    End If
    
    If Len(FindOnPath) = 0 And bUseDefaultOnFailure Then
        FindOnPath = sAppName
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "FindOnPath", "AppName: ", sAppName
    ToggleWow64FSRedirection True
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFileNameAndExt(Path As String) As String ' вернет только имя файла вместе с расширением
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then
        GetFileNameAndExt = Mid$(Path, pos + 1)
    Else
        GetFileNameAndExt = Path
    End If
End Function

