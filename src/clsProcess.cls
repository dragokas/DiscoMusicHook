VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' Process Class Module by Alex Dragokas
'
' Thanks to Catstail (Boris Fayfel) & The Trick (Anatolii Krivous) for improvements.
'

Option Explicit

Private Const MAX_PATH As Long = 260&

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile(MAX_PATH - 1) As Integer
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type SAFEARRAY
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
End Type

Private Enum TOKEN_TYPE
    TokenPrimary = 1
    TokenImpersonation
End Enum

Private Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous
    SecurityIdentification
    SecurityImpersonation
    SecurityDelegation
End Enum

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32.dll" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32.dll" Alias "Process32FirstW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32.dll" Alias "Process32NextW" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CommandLineToArgvW Lib "shell32.dll" (ByVal lpCmdLine As Long, pNumArgs As Long) As Long
'Private Declare Function lstrcmpi Lib "kernel32.dll" Alias "lstrcmpiW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (Src As Any, Dst As Any) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
'Private Declare Function SetEnvironmentVariable Lib "kernel32.dll" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpValue As Long) As Long
'Private Declare Function GetEnvironmentVariable Lib "kernel32.dll" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInformation As Any) As Long
Private Declare Function DuplicateTokenEx Lib "advapi32.dll" (ByVal hExistingToken As Long, ByVal dwDesiredAccess As Long, ByVal lpTokenAttributes As Long, ByVal ImpersonationLevel As Long, ByVal TokenType As Long, phNewToken As Long) As Long
Private Declare Function GetShellWindow Lib "user32.dll" () As Long
Private Declare Function CreateProcessWithTokenW Lib "advapi32.dll" (ByVal hToken As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function IUnknown_QueryService Lib "Shlwapi.dll" (ByVal pUnk As IUnknown, ByRef guidService As Any, ByRef riid As Any, ByRef ppvOut As Any) As Long
Private Declare Function EndTask Lib "user32.dll" (ByVal hWnd As Long, ByVal fShutDown As Long, ByVal fForce As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" _
                         Alias "LoadLibraryW" ( _
                         ByVal lpLibFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Private Declare Sub GetStartupInfo Lib "kernel32.dll" Alias "GetStartupInfoW" (lpStartupInfo As STARTUPINFO)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const INFINITE                  As Long = -1
Private Const STILL_ACTIVE              As Long = &H103
Private Const WAIT_FAILED               As Long = -1
Private Const TH32CS_SNAPPROCESS        As Long = 2
Private Const TH32CS_SNAPNOHEAPS        As Long = &H40000000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = 1024&
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
Private Const WM_CLOSE                  As Long = 16&
Private Const WM_QUIT                   As Long = &H12  ' wparam - exitcode
Private Const PROCESS_TERMINATE         As Long = &H1
Private Const ERROR_SUCCESS             As Long = 0&
Private Const SW_SHOWNORMAL = 1

Private si    As STARTUPINFO
Private pi    As PROCESS_INFORMATION
Private inIDE As Boolean


Public Property Let SetProcessID(PID As Long)
    pi.dwProcessId = PID
End Property

Public Property Get GetProcessID() As Long
    GetProcessID = pi.dwProcessId
End Property

Public Property Get GetThreadId() As Long
    GetThreadId = pi.dwThreadId
End Property

Public Property Get GetProcessHandle() As Long
    GetProcessHandle = pi.hProcess
End Property

Public Property Get GetThreadHandle() As Long
    GetThreadHandle = pi.hThread
End Property

Function IsRunned(Optional ByRef ProcessName As String, Optional ByRef ProcessID As Long) As Boolean
    ' ??????? ?? ???????, ????????? ?? ????? ??? ProcessID
    ' (??????? ???????? ?? ???????????)
    ' ???? ??????? ??????, ? ?????????? ProcessID ?????????? PID ?????????? ???????? (???? ????? ????? ?? ?????),
    ' ???? ? ?????????? ProcessName ???????????? ??? ?????????? ???????? (???? ????? ????? ?? PID).

    On Error GoTo ErrorHandler:

    Dim hSnap   As Long
    Dim pe      As PROCESSENTRY32

    If ProcessName = vbNullString And ProcessID = 0 Then
        ProcessID = GetProcessID()
    End If
    
    If ProcessName = vbNullString And ProcessID = 0 Then Exit Function

    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS Or TH32CS_SNAPNOHEAPS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)

    If Process32First(hSnap, pe) Then
        Do
            If Len(ProcessName) = 0 Then
                If pe.th32ProcessID = ProcessID Then
                    IsRunned = True
                    ProcessName = StringFromPtrW(VarPtr(pe.szExeFile(0)))
                    'ProcessName = TrimNull(StrConv(pe.szExeFile, vbFromUnicode))
                    Exit Do
                End If
            Else
                If StrComp(StringFromPtrW(VarPtr(pe.szExeFile(0))), ProcessName, 1) = 0 Then
                    IsRunned = True
                    ProcessID = pe.th32ProcessID
                    Exit Do
                End If
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_IsRunned"
    If inIDE Then Stop: Resume Next
End Function


Function WaitForTerminate( _
                        Optional ByVal WindowHandle As Long, _
                        Optional ByVal ProcessHandle As Long, _
                        Optional ByVal FreezeThread As Boolean, _
                        Optional ByVal TimeOutMs As Long, _
                        Optional ByVal ProcessID As Long) As Long
                        ' ??????? ?????????? ????????                           'Thanks for help to Catstail (?. ???????)
    
    On Error GoTo ErrorHandler:
    
    ' FreezeThread, ???? true - ????????? ??????? ????? ?? ?????, ???? ????????????? ??????? ?? ????? ??????
    ' TimeOut, ???? FreezeThread = false, ?????? ????? ???????? ???????? ?????????? ????????.
    ' ???????????? ????????: ExitCode ????????, ??? -1, ???? ????? ??????? ???????? ??? WaitForSingleObject ?????? ?????? (WAIT_FAILED)
    
    Dim iExit       As Long
    Dim hProc       As Long
    Dim iResult     As Long
    Dim dwThreadId  As Long
    Dim hThread     As Long
    Dim InitialTime As Single
    Dim DeadTime    As Single
    Dim bTimeOutON  As Boolean
    
    Static isinit As Boolean
    Static bIsWinVistaOrLater As Boolean
    
    If Not isinit Then
        isinit = True
        Dim inf(68) As Long
        inf(0) = 276: GetVersionEx inf(0): bIsWinVistaOrLater = (inf(1) >= 6)
    End If
    
    If ProcessID <> 0 Then
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
        
    ElseIf WindowHandle <> 0 Then
        ' ???????? PID ????????
        dwThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        '::: ???????? ????? ????????
        hProc = OpenProcess(IIf(bIsWinVistaOrLater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or SYNCHRONIZE, False, ProcessID)
    ElseIf ProcessHandle <> 0 Then
        hProc = ProcessHandle
    Else
        hProc = GetProcessHandle()
        hThread = GetThreadHandle()
    End If
    
    If hProc = 0 Then Exit Function
    
    If FreezeThread Then
       '::: ???? ?? ????????? ?????
       iResult = WaitForSingleObject(hProc, INFINITE)
       '::: ???? ???????? ??????????? ???????? - ?????? ??? -1
       If (iResult = WAIT_FAILED) Then
           iExit = -1
       Else
           '::: ???????? ??? ??????????
           GetExitCodeProcess hProc, iExit
       End If
    Else
       '::: ??????? ???????? ??? ??????????
       GetExitCodeProcess hProc, iExit
       InitialTime = Timer
       DeadTime = InitialTime + TimeOutMs / 1000
       bTimeOutON = (TimeOutMs <> INFINITE)
       '::: ???? ??????? ??? ??? ? ????????? ??????????,
       '::: ?? ????? ??????? ?????????
       Do While (iExit = STILL_ACTIVE) And IIf(bTimeOutON, DeadTime > Timer, True)
          DoEvents
          Sleep 100
          GetExitCodeProcess hProc, iExit
       Loop
       If bTimeOutON And (Timer > DeadTime) Then iExit = -1
    End If
    '::: ??????? ????? ??? ???????? ?????????? ?????? ??????????, ????? ?? ????? ??????????? ??? ??????????????? ?????????? ????????
    If iExit = ERROR_SUCCESS Then
        CloseHandle hThread
        CloseHandle hProc
    End If
    WaitForTerminate = iExit
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_WaitForTerminate"
    If inIDE Then Stop: Resume Next
End Function

Private Function GetPIDsByName(ProcessName As String) As Long()
    On Error GoTo ErrorHandler:
    Dim pe               As PROCESSENTRY32
    Dim hSnap            As Long
    Dim PID()            As Long
    ReDim PID(0) As Long
    
    
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS Or TH32CS_SNAPNOHEAPS, 0)

    If hSnap = INVALID_HANDLE_VALUE Then Exit Function

    pe.dwSize = Len(pe)
    
    If Process32First(hSnap, pe) Then
        Do
            If StrComp(StringFromPtrW(VarPtr(pe.szExeFile(0))), ProcessName, 1) = 0 Then
                
                ReDim Preserve PID(UBound(PID) + 1) As Long
                PID(UBound(PID)) = pe.th32ProcessID
            End If
        Loop While Process32Next(hSnap, pe)
    End If
    CloseHandle hSnap
    
    GetPIDsByName = PID
    
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_GetPIDsByName"
    If inIDE Then Stop: Resume Next
End Function

Public Function ProcessClose( _
                            Optional ByVal WindowHandle As Long, _
                            Optional ProcessName As String, _
                            Optional Async As Boolean, _
                            Optional TimeOutMs As Long = INFINITE, _
                            Optional ByVal ProcessID As Long, _
                            Optional SendCloseMsg As Boolean) As Long
                            '???????????? ???????? = ExitCode ????????, ??? -1 ? ?????? ???????
    
    On Error GoTo ErrorHandler:

    

    Dim lr               As Long
    Dim hThread          As Long
    Dim hProc            As Long
    Dim dwThreadId       As Long
    Dim PID()            As Long
    Dim i                As Long
    Dim ProcNameCur      As String

    If Len(ProcessName) <> 0 Then

        PID() = GetPIDsByName(ProcessName)
        
        If IsArrDimmed(PID) Then
            For i = 1 To UBound(PID)
                ProcessClose , , True, 0, PID(i), SendCloseMsg 'Async mode
            Next
        End If
        
        If Not Async Then
            If TimeOutMs = INFINITE Then
                For i = 1 To UBound(PID)
                    ProcNameCur = ""
                    Call IsRunned(ProcNameCur, PID(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        Call WaitForTerminate(, , False, INFINITE, PID(i))
                    End If
                Next
            Else
                For i = 1 To UBound(PID)
                    ProcNameCur = ""
                    Call IsRunned(ProcNameCur, PID(i)) 'get name by PID
                    If StrComp(ProcNameCur, ProcessName, vbTextCompare) = 0 Then 'not changed ?
                        ProcessID = PID(i)
                        Exit For ' get only 1 process
                    End If
                Next
                If ProcessID <> 0 Then Call WaitForTerminate(, , False, TimeOutMs, ProcessID)
            End If
        End If
        
        ProcessID = 0
        ProcessClose = IsRunned(ProcessName, ProcessID)
        Exit Function
        
    ElseIf WindowHandle <> 0 Then
        dwThreadId = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE ???????? ????
        Call WaitForTerminate(, hProc, False, 1000)                              ' ???????? 1 ???.
        lr = PostMessage(WindowHandle, WM_QUIT, 0, 0)                            ' WM_QUIT ???????? ????
        If Not Async Then
            lr = WaitForTerminate(WindowHandle, , False, TimeOutMs)              ' ????????, ??? ????????????? ???????? ??????
        Else
            lr = WaitForTerminate(WindowHandle, , False, 500)
        End If
        If IsProcedureAvail("EndTask", "user32.dll") Then
            EndTask WindowHandle, False, True
        End If
        If IsRunned(, ProcessID) Then
            hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
        End If
        
    ElseIf ProcessID <> 0 Then
        hProc = OpenProcess(PROCESS_TERMINATE, False, ProcessID)
    
    Else
        hThread = pi.hThread
        ProcessID = pi.dwProcessId
        hProc = pi.hProcess
    End If
    
    If SendCloseMsg Then
        hThread = GetWindowThreadProcessId(ByVal WindowHandle, ProcessID)
        lr = PostMessage(WindowHandle, WM_CLOSE, 0, 0)                           ' WM_CLOSE ???????? ????
        Call WaitForTerminate(, hProc, False, 1000)
    End If
    
    If hProc <> 0 Then
        TerminateProcess hProc, 0
        If Not (Async) Then
            ProcessClose = WaitForTerminate(, hProc, False, TimeOutMs)
        End If
    End If
    
    ProcessClose = Not IsRunned(, ProcessID)                           ' ??????? ??????? ? ?????? ?????????
    
    If ProcessClose Then
        If hThread <> 0 Then CloseHandle hThread
        If hProc <> 0 Then CloseHandle hProc
    End If
    
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessClose", "ProcessName:", ProcessName
    If inIDE Then Stop: Resume Next
End Function


Public Function ProcessRun( _
                            ByVal FileName As String, _
                            Optional ByVal CommandLine As String, _
                            Optional ByVal CurrentDirectory As String = vbNullString, _
                            Optional WindowStyle As Long = SW_SHOWNORMAL, _
                            Optional CloseHandles As Boolean = False) As Long
    
    On Error GoTo ErrorHandler:
    
    
    Const STARTF_USESHOWWINDOW      As Long = 1
    Const NORMAL_PRIORITY_CLASS     As Long = &H20
    
    Dim n           As Long
    Dim lr          As Long
    Dim CMDLine     As String
    Dim argc        As Long
    Dim argv()      As String
    
    CMDLine = """" & FileName & """"                    ' ??? ????? ? ???????
    
    If CommandLine <> vbNullString Then                 ' ???? ???? ?????????, ?? ???? ????????? ? ????????
    
        ParseCommandLine CommandLine, argc, argv        ' ????????? ????????? ????????? ?????? ?? ????????????
        For n = 1 To argc
            CMDLine = CMDLine & " """ & argv(n) & """"
        Next
    End If
    
    '???? ????????? ?????? ???????????? ??? cmd.exe, ?? ??????? ?????????
    If StrComp(FileName, Environ("ComSpec"), 1) = 0 Or StrComp(FileName, "schtasks.exe", 1) = 0 Then
        CMDLine = """" & FileName & """" & " " & CommandLine
    End If
    
    si.cb = Len(si)
    GetStartupInfo si                                   ' ????????? ????????? ???????? ????????
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = WindowStyle                        ' ????? ????
    
    lr = CreateProcess(ByVal 0&, _
                       StrPtr(CMDLine), _
                       ByVal 0&, _
                       ByVal 0&, _
                       False, _
                       NORMAL_PRIORITY_CLASS, _
                       ByVal 0&, _
                       StrPtr(CurrentDirectory), _
                       si, _
                       pi)

    ProcessRun = lr    ' not 0 is SUCCESS
    
    If CloseHandles Then
        If pi.hProcess <> 0 Then CloseHandle pi.hProcess
        If pi.hThread <> 0 Then CloseHandle pi.hThread
    End If
    
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "clsProcess_ProcessRun", "FileName:", FileName, "CommandLine:", CommandLine
    If inIDE Then Stop: Resume Next
End Function

Private Sub ParseCommandLine(Line As String, argc As Long, argv() As String)
    On Error GoTo ErrorHandler:
    Dim ptr     As Long
    Dim Index   As Long
    Dim strLen  As Long
    Dim strAdr  As Long
    
    ptr = CommandLineToArgvW(StrPtr(Line), argc)
    ReDim argv(argc)
    argv(0) = App.Path & "\" & App.EXEName & ".exe"
    If argc = 0 Then Exit Sub
    
    For Index = 1 To argc
        GetMem4 ByVal ptr + (Index - 1) * 4, strAdr
        strLen = lstrlen(strAdr)
        argv(Index) = Space$(strLen)
        lstrcpyn StrPtr(argv(Index)), strAdr, strLen + 1
    Next
    GlobalFree ptr
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "clsProcess_ParseCommandLine", "Line:", Line
    If inIDE Then Stop: Resume Next
End Sub

Private Function TrimNull(S$) As String
    TrimNull = Left$(S, lstrlen(StrPtr(S)))
End Function

Private Sub Class_Initialize()
    inIDE = (App.LogMode = 0)
End Sub

Public Function IsArrDimmed(vArray As Variant) As Boolean
    IsArrDimmed = (GetArrDims(vArray) > 0)
End Function

Public Function GetArrDims(vArray As Variant) As Integer
    Dim ppSA As Long
    Dim pSA As Long
    Dim VT As Long
    Dim sa As SAFEARRAY
    Const vbByRef As Integer = 16384

    If IsArray(vArray) Then
        GetMem4 ByVal VarPtr(vArray) + 8, ppSA      ' pV -> ppSA (pSA)
        If ppSA <> 0 Then
            GetMem2 vArray, VT
            If VT And vbByRef Then
                GetMem4 ByVal ppSA, pSA                 ' ppSA -> pSA
            Else
                pSA = ppSA
            End If
            If pSA <> 0 Then
                memcpy sa, ByVal pSA, LenB(sa)
                If sa.pvData <> 0 Then
                    GetArrDims = sa.cDims
                End If
            End If
        End If
    End If
End Function

Private Sub Class_Terminate()
    If pi.hThread <> 0 Then
        CloseHandle pi.hThread
    End If
    If pi.hProcess <> 0 Then
        CloseHandle pi.hProcess
    End If
End Sub

Private Function StringFromPtrW(ptr As Long) As String
    Dim strSize As Long
    If 0 <> ptr Then
        strSize = lstrlen(ptr)
        If 0 <> strSize Then
            StringFromPtrW = String$(strSize, 0&)
            lstrcpyn StrPtr(StringFromPtrW), ptr, strSize + 1&
        End If
    End If
End Function


Public Function IsProcedureAvail(ByVal ProcedureName As String, ByVal DllFilename As String) As Boolean
    
    Dim hModule As Long, procAddr As Long
    hModule = LoadLibrary(StrPtr(DllFilename))
    If hModule Then
        procAddr = GetProcAddress(hModule, StrPtr(StrConv(ProcedureName, vbFromUnicode)))
        FreeLibrary hModule
    End If
    IsProcedureAvail = (procAddr <> 0)
    
End Function

