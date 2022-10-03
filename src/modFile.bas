Attribute VB_Name = "modFile"
'
' modFile module by Alex Dragokas
'

Option Explicit

Const MAX_PATH As Long = 260&
Const MAX_PATH_W     As Long = 32767&
Const MAX_FILE_SIZE As Currency = 104857600@

Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 4
End Enum

Enum CACHE_TYPE
    USE_CACHE
    NO_CACHE
End Enum
 
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
 
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
 
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH) As Integer
    lpszAlternate(14) As Integer
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Declare Function PathFileExists Lib "Shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
Private Declare Function GetLongPathNameW Lib "kernel32" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bDontOverwrite As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceW" (ByVal lpDeviceName As Long, ByVal lpTargetPath As Long, ByVal ucchMax As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function QueryFullProcessImageName Lib "kernel32.dll" Alias "QueryFullProcessImageNameW" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As Long, ByVal lpdwSize As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExW" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lpImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32.dll" Alias "GetFullPathNameW" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, lpFilePart As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32.dll" Alias "GetLongPathNameW" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long

Const FILE_SHARE_READ           As Long = &H1&
Const FILE_SHARE_WRITE          As Long = &H2&
Const FILE_SHARE_DELETE         As Long = 4&
Const FILE_READ_ATTRIBUTES      As Long = &H80&
Const OPEN_EXISTING             As Long = 3&
Const CREATE_ALWAYS             As Long = 2&
Const GENERIC_READ              As Long = &H80000000
Const GENERIC_WRITE             As Long = &H40000000
Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Const FILE_ATTRIBUTE_READONLY   As Long = 1&
Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Const ERROR_SUCCESS             As Long = 0&
Const ERROR_FILE_NOT_FOUND      As Long = 2&
Const ERROR_ACCESS_DENIED       As Long = 5&
Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Const NO_ERROR                  As Long = 0&
Const FILE_BEGIN                As Long = 0&
Const FILE_CURRENT              As Long = 1&
Const FILE_END                  As Long = 2&
Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
Const ERROR_PARTIAL_COPY            As Long = 299&

Const DRIVE_FIXED               As Long = 3&
Const DRIVE_RAMDISK             As Long = 6&

Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Const KEY_QUERY_VALUE           As Long = &H1&
Const RegType_DWord             As Long = 4&

Const ch_Dot                    As String = "."
Const ch_DotDot                 As String = ".."
Const ch_Slash                  As String = "\"
Const ch_SlashAsterisk          As String = "\*"

Private lWow64Old               As Long
Private DriveTypeName           As New Collection
Private arrPathFolders()        As String
Private arrPathFiles()          As String
Private Total_Folders           As Long
Private Total_Files             As Long



Public Function FileExists(ByVal sFile$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    Dim Redirect As Boolean
    
    sFile = Trim$(sFile)
    If Len(sFile) = 0 Then Exit Function
    If Left$(sFile, 2) = "\\" Then Exit Function 'DriveType = "REMOTE"
    
    ' use 2 methods for reliability reason (both supported unicode pathes)
    Dim Ex(1) As Boolean
    Dim ret As Long
    
    Dim WFD     As WIN32_FIND_DATA
    Dim hFile   As Long
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFile)
    
    ret = GetFileAttributes(StrPtr(sFile))
    If ret <> INVALID_HANDLE_VALUE And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY)) Then Ex(0) = True
 
    hFile = FindFirstFile(StrPtr(sFile), WFD)
    Ex(1) = (hFile <> INVALID_HANDLE_VALUE) And Not CBool(WFD.dwFileAttributes And vbDirectory)
    FindClose hFile

    ' // here must be enabling of FS redirector
    If Redirect Then Call ToggleWow64FSRedirection(True)

    FileExists = Ex(0) Or Ex(1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileExists", "File:", sFile$
    If inIDE Then Stop: Resume Next
End Function

Public Function FolderExists(ByVal sFolder$, Optional ForceUnderRedirection As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim ret As Long
    sFolder = Trim$(sFolder)
    If Len(sFolder) = 0 Then Exit Function
    If Left$(sFolder, 2) = "\\" Then Exit Function 'network path
    
    '// FS redirection checking
    
    ret = GetFileAttributes(StrPtr(sFolder))
    FolderExists = CBool(ret And vbDirectory) And (ret <> INVALID_FILE_ATTRIBUTES)
    
    '// FS redirection enambling
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FolderExists", "Folder:", sFolder$, "Redirection: ", ForceUnderRedirection
    If inIDE Then Stop: Resume Next
End Function

Function FileLenW(Path As String) As Currency ', Optional DoNotUseCache As Boolean
    On Error GoTo ErrorHandler
'    ' Last cached File
'    Static CachedFile As String
'    Static CachedSize As Currency
    
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency

'    If Not DoNotUseCache Then
'        If StrComp(Path, CachedFile, 1) = 0 Then
'            FileLenW = CachedSize
'            Exit Function
'        End If
'    End If

    hFile = CreateFile(StrPtr(Path), FILE_READ_ATTRIBUTES, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
'        If Not DoNotUseCache Then
'            CachedFile = Path
'            CachedSize = FileLenW
'        End If
        CloseHandle hFile: hFile = 0&
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileLenW", "File:", Path, "hFile:", hFile, "FileSize:", FileSize, "Return:", lr
End Function



Public Function OpenW(FileName As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional MountToMemory As Boolean) As Boolean '// TODO: MountToMemory
    
    Dim FSize As Currency
    
    'Print #ffOpened, FileName
    
    If Access And (FOR_READ Or FOR_READ_WRITE) Then
        If Not FileExists(FileName) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
        
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        'WriteCon "Wrong access mode!", cErr
    End If

    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
    
    ' ограничение на максимально возможный файл для открытия ( > 100 МБ )
    If OpenW Then
        If Access And (FOR_READ Or FOR_READ_WRITE) Then
            FSize = LOFW(retHandle)
            If FSize > MAX_FILE_SIZE Then
                CloseHandle retHandle
                retHandle = INVALID_HANDLE_VALUE
                OpenW = False
                '"Не хочу и не буду открывать этот файл, потому что его размер превышает безопасный максимум"
                Err.Clear: ErrorMsg Err, "modFile.OpenW: " & "Trying to open too big file" & ": (" & (FSize \ 1024 \ 1024) & " MB.) " & FileName
            End If
        End If
    Else
        ErrorMsg Err, "modFile.OpenW: Cannot open file: " & FileName
        Err.Raise 75 ' Path/File Access error
    End If

End Function

                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
                                                                  
    'On Error GoTo ErrorHandler
    
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If INVALID_SET_FILE_POINTER <> SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
        If NO_ERROR = Err.LastDllError Then
            vType = VarType(vOut)
            
            If 0 <> cbToRead Then   'vbError = vType
                lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
            ElseIf vbString = vType Then
                lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
                If Err.LastDllError <> 0 Or lr = 0 Then Err.Raise 52
                
                vOut = StrConv(vOut, vbUnicode)
                If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
            Else
                'do a bit of magik :)
                memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
                Select Case vType
                Case vbByte
                    lr = ReadFile(hFile, ptr, 1&, lBytesRead, 0&)
                Case vbInteger
                    lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
                Case vbLong
                    lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
                Case vbCurrency
                    lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
                Case Else
                    UnknType = True
                    Err.Clear: ErrorMsg Err, "modFile.GetW. type #" & VarType(vOut) & " of buffer is not supported.": Err.Raise 52
                End Select
            End If
            GetW = (0 <> lr)
            If 0 = lr And Not UnknType Then Err.Clear: ErrorMsg Err, "Cannot read file!": Err.Raise 52
        Else
            Err.Clear: ErrorMsg Err, "Cannot set file pointer!": Err.Raise 52
        End If
    Else
        Err.Clear: ErrorMsg Err, "Cannot set file pointer!": Err.Raise 52
    End If
    
'    Exit Function
'ErrorHandler:
'    AppendErrorLogFormat Now, err, "modFile.GetW"
'    Resume Next
End Function

Public Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = Err.LastDllError Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then PutW = True
        
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.PutW"
End Function

Public Function LOFW(hFile As Long) As Currency
    On Error GoTo ErrorHandler
    Dim lr          As Long
    Dim FileSize    As Currency
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then
                LOFW = FileSize * 10000&
            Else
                Err.Clear
                ErrorMsg Now, "File is too big. Size: " & FileSize
            End If
        End If
    End If
ErrorHandler:
End Function

Public Function CloseW(hFile As Long) As Long
    CloseW = CloseHandle(hFile)
End Function

Public Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean

    On Error GoTo ErrorHandler

    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Dim lr                  As Long

    OldStatus = Not IsNotRedirected

    If Not OSver.IsWin64 Then Exit Function

    If Len(PathNecessity) <> 0 Then
        If StrComp(Left$(PathNecessity, Len(Env.sWinDir)), Env.sWinDir, vbTextCompare) <> 0 Then Exit Function
    End If

    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.PutW"
    If inIDE Then Stop: Resume Next
End Function


Public Function GetExtensionName(Path As String) As String  'вернет .ext
    Dim pos As Long
    pos = InStrRev(Path, ".")
    If pos <> 0 Then GetExtensionName = Mid$(Path, pos)
End Function

Public Function GetPathName(Path As String) As String   ' получить родительский каталог
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then GetPathName = Left$(Path, pos - 1)
End Function

' Получить только имя файла (без расширения имени)
Public Function GetFileName(Path As String) As String
    On Error GoTo ErrorHandler
    Dim posDot      As Long
    Dim posSl       As Long
    
    posSl = InStrRev(Path, "\")
    If posSl <> 0 Then
        posDot = InStrRev(Path, ".")
        If posDot < posSl Then posDot = 0
    Else
        posDot = InStrRev(Path, ".")
    End If
    If posDot = 0 Then posDot = Len(Path) + 1
    
    GetFileName = Mid$(Path, posSl + 1, posDot - posSl - 1)
    Exit Function
ErrorHandler:
End Function

'main function to list folders

' Возвращает массив путей.
' Если ничего не найдено - возвращается неинициализированный массив.
Public Function ListSubfolders(Path As String, Optional Recursively As Boolean = False) As String()
    Dim bRedirected As Boolean
    'прежде, чем использовать ListSubfolders_Ex, нужно инициализировать глобальные массивы.
    ReDim arrPathFolders(100) As String
    'при каждом вызове ListSubfolders_Ex следует обнулить глобальный счетчик файлов
    Total_Folders = 0&
    
    If OSver.IsWin64 Then
        If StrBeginWith(Path, Env.sWinDir) Then
            ToggleWow64FSRedirection False
            bRedirected = True
        End If
    End If
    
    'вызов тушки
    Call ListSubfolders_Ex(Path, Recursively)
    If Total_Folders > 0 Then
        Total_Folders = Total_Folders - 1
        ReDim Preserve arrPathFolders(Total_Folders)      '0 to Max -1
        ListSubfolders = arrPathFolders
    End If
    
    If bRedirected Then ToggleWow64FSRedirection True
End Function


Private Sub ListSubfolders_Ex(Path As String, Optional Recursively As Boolean = False)
    On Error GoTo ErrorHandler
    'On Error Resume Next
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim l               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Folders as long
    ' arrPathFolders() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        l = fd.dwFileAttributes And &H600& ' мимо симлинков
        Do While l <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            l = fd.dwFileAttributes And &H600&
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If UBound(arrPathFolders) < Total_Folders Then ReDim Preserve arrPathFolders(UBound(arrPathFolders) + 100&) As String
                        arrPathFolders(Total_Folders) = SubPathName
                        Total_Folders = Total_Folders + 1&
                        If Recursively Then
                            Call ListSubfolders_Ex(SubPathName, Recursively)
                        End If
                    End If
                End If
            End If
        End If
        
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListSubfolders", "Folder:", Path
    Resume Next
End Sub

'main function to list files

Public Function ListFiles(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False) As String()
    Dim bRedirected As Boolean
    'прежде, чем использовать ListFiles_Ex, нужно инициализировать глобальные массивы.
    ReDim arrPathFiles(100) As String
    'при каждом вызове ListFiles_Ex следует обнулить глобальный счетчик файлов
    Total_Files = 0&
    
    If OSver.IsWin64 Then
        If StrBeginWith(Path, Env.sWinDir) Then
            ToggleWow64FSRedirection False
            bRedirected = True
        End If
    End If
    
    'вызов тушки
    Call ListFiles_Ex(Path, Extension, Recursively)
    If Total_Files > 0 Then
        Total_Files = Total_Files - 1
        ReDim Preserve arrPathFiles(Total_Files)      '0 to Max -1
        ListFiles = arrPathFiles
    End If
    
    If bRedirected Then ToggleWow64FSRedirection True
End Function


Private Sub ListFiles_Ex(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False)
    'Example of Extension:
    '".txt" - txt files
    'empty line - all files (by default)

    On Error GoTo ErrorHandler
    'On Error Resume Next
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim l               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Files as long
    ' arrPathFiles() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        l = fd.dwFileAttributes And &H600& ' мимо симлинков
        Do While l <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            l = fd.dwFileAttributes And &H600&
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If Recursively Then
                            Call ListFiles_Ex(SubPathName, Extension, Recursively)
                        End If
                    End If
                End If
            Else
                If inArray(GetExtensionName(PathName), SplitSafe(Extension, ";"), , , 1) Or Len(Extension) = 0 Then
                    SubPathName = Path & "\" & PathName
                    If UBound(arrPathFiles) < Total_Files Then ReDim Preserve arrPathFiles(UBound(arrPathFiles) + 100&) As String
                    arrPathFiles(Total_Files) = SubPathName
                    Total_Files = Total_Files + 1&
                End If
            End If
        End If
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListFiles_Ex", "File:", Path
    Resume Next
End Sub

Public Function GetLocalDisks$()
    Dim lDrives&, i&, sDrive$, sLocalDrives$
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(Asc("A") + i) & ":\"
            Select Case GetDriveType(StrPtr(sDrive))
                Case DRIVE_FIXED, DRIVE_RAMDISK: sLocalDrives = sLocalDrives & Chr$(Asc("A") + i) & " "
            End Select
        End If
    Next i
    GetLocalDisks = Trim$(sLocalDrives)
End Function

Public Function GetLongFilename$(sFilename$)
    Dim sLongFilename$
    If InStr(sFilename, "~") = 0 Then
        GetLongFilename = sFilename
        Exit Function
    End If
    sLongFilename = String(512, 0)
    GetLongPathNameW StrPtr(sFilename), StrPtr(sLongFilename), Len(sLongFilename)
    GetLongFilename = TrimNull(sLongFilename)
End Function

Public Function GetFilePropVersion(sFilename As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, uVFFI As VS_FIXEDFILEINFO, sVersion$
    
    If Not FileExists(sFilename) Then Exit Function
    
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then Exit Function
    
    ReDim uBuf(0 To lDataLen - 1)
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
    
        If 0 <> VerQueryValue(uBuf(0), StrPtr("\"), hData, lDataLen) Then
        
            If hData <> 0 Then
        
                memcpy uVFFI, ByVal hData, Len(uVFFI)
    
                With uVFFI
                    sVersion = .dwFileVersionMSh & "." & _
                        .dwFileVersionMSl & "." & _
                        .dwFileVersionLSh & "." & _
                        .dwFileVersionLSl
                End With
            End If
        End If
    End If
    GetFilePropVersion = sVersion
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropVersion", sFilename
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFilePropCompany(sFilename As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, Stady&
    
    If Not FileExists(sFilename) Then Exit Function
    
    Stady = 1
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then Exit Function
    
    Stady = 2
    ReDim uBuf(0 To lDataLen - 1)
    
    Stady = 3
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
        
        Stady = 4
        VerQueryValue uBuf(0), StrPtr("\VarFileInfo\Translation"), hData, lDataLen
        If lDataLen = 0 Then Exit Function
        
        Stady = 5
        memcpy uCodePage(0), ByVal hData, 4
        
        Stady = 6
        sCodePage = Right$("0" & Hex(uCodePage(1)), 2) & _
                Right$("0" & Hex(uCodePage(0)), 2) & _
                Right$("0" & Hex(uCodePage(3)), 2) & _
                Right$("0" & Hex(uCodePage(2)), 2)
        
        'get CompanyName string
        Stady = 7
        If VerQueryValue(uBuf(0), StrPtr("\StringFileInfo\" & sCodePage & "\CompanyName"), hData, lDataLen) = 0 Then Exit Function
    
        If lDataLen > 0 And hData <> 0 Then
            Stady = 8
            sCompanyName = String$(lDataLen, 0)
            
            Stady = 9
            lstrcpy ByVal StrPtr(sCompanyName), ByVal hData
        End If
        
        Stady = 10
        GetFilePropCompany = RTrimNull(sCompanyName)
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropCompany", sFilename, "DataLen: ", lDataLen, "hData: ", hData, "sCodePage: ", sCodePage, _
        "Buf: ", uCodePage(0), uCodePage(1), uCodePage(2), uCodePage(3), "Stady: ", Stady
    If inIDE Then Stop: Resume Next
End Function

'Delete File with unlock access rights on failure. Return non 0 on success.
Public Function DeleteFileWEx(lpSTR As Long, Optional ForceDeleteMicrosoft As Boolean) As Long
    On Error GoTo ErrorHandler:

    Dim iAttr As Long, lr As Long, sExt As String

    Dim FileName$
    FileName = String$(lstrlen(lpSTR), vbNullChar)
    If Len(FileName) <> 0 Then
        lstrcpy StrPtr(FileName), lpSTR
    Else
        Exit Function
    End If
    
    sExt = GetExtensionName(FileName)
    
    ToggleWow64FSRedirection False, FileName
    
    iAttr = GetFileAttributes(lpSTR)
    If (iAttr And 2048) Then iAttr = iAttr - 2048
    
    If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
    lr = DeleteFileW(lpSTR)
    
    If lr <> 0 Then
        DeleteFileWEx = lr
        ToggleWow64FSRedirection True, FileName
        Exit Function
    End If
    
    If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
        DeleteFileWEx = 1
        ToggleWow64FSRedirection True, FileName
        Exit Function
    End If
    
    If Err.LastDllError = ERROR_ACCESS_DENIED Then
        TryUnlock FileName
        If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
        lr = DeleteFileW(lpSTR)
    End If
    ToggleWow64FSRedirection True, FileName
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.DeleteFileWEx", "File:", FileName
    If inIDE Then Stop: Resume Next
End Function

Sub TryUnlock(ByVal File As String)  'получения прав NTFS + смена владельца на локальную группу "Администраторы"
    On Error GoTo ErrorHandler:
    Dim TakeOwn As String
    Dim Icacls As String
    Dim DosName As String
    
    DosName = GetDOSFilename(File)
    If Len(DosName) <> 0 Then File = DosName
    
    If Not OSver.IsWindowsVistaOrGreater Then Exit Sub
    
    If OSver.Bitness = "x64" And FolderExists(Env.sWinDir & "\sysnative") Then
        TakeOwn = EnvironW("%SystemRoot%") & "\Sysnative\takeown.exe"
        Icacls = EnvironW("%SystemRoot%") & "\Sysnative\icacls.exe"
    Else
        TakeOwn = EnvironW("%SystemRoot%") & "\System32\takeown.exe"
        Icacls = EnvironW("%SystemRoot%") & "\System32\icacls.exe"
    End If
    
    If FileExists(TakeOwn) Then
        Proc.ProcessRun TakeOwn, "/F " & """" & File & """" & " /A", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    End If
    
    If FileExists(Icacls) Then
        Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *S-1-1-0:F", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
        
        Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *S-1-5-32-544:F", , 0
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    
        If 0 <> Len(OSver.SID_CurrentProcess) Then
            Proc.ProcessRun Icacls, """" & File & """" & " /grant:r *" & OSver.SID_CurrentProcess & ":F", , 0
        Else
            Proc.ProcessRun Icacls, """" & File & """" & " /grant:r """ & EnvironW("%UserName%") & """:F", , 0
        End If
        If ERROR_SUCCESS <> Proc.WaitForTerminate(, , , 5000) Then
            Proc.ProcessClose , , True
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "TryUnlock", "File:", File
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetEmptyName(ByVal sFullPath As String) As String

    Dim sExt As String
    Dim sName As String
    Dim sPath As String
    Dim i As Long

    If Not FileExists(sFullPath) Then
        GetEmptyName = sFullPath
    Else
        sExt = GetExtensionName(sFullPath)
        sPath = GetPathName(sFullPath)
        sName = GetFileName(sFullPath)
        Do
            i = i + 1
            sFullPath = BuildPath(sPath, sName & "(" & i & ")" & sExt)
        Loop While FileExists(sFullPath)
        
        GetEmptyName = sFullPath
    End If
End Function

Public Function BackupFile(sFullPath As String) As Boolean
    
    Dim sBackup As String
    
    sBackup = sFullPath & ".bak"
    
    If FileExists(sFullPath) Then
        sBackup = GetEmptyName(sBackup)
        
        BackupFile = CopyFile(StrPtr(sFullPath), StrPtr(sBackup), False)
    End If
    
End Function 'true on success

Public Function ConvertDosDeviceToDriveName(inDosDeviceName As String) As String
    On Error GoTo ErrorHandler:
    'alternatives:
    'RtlNtPathNameToDosPathName (XP+)
    'RtlVolumeDeviceToDosName
    'IOCTL_MOUNTMGR_QUERY_DOS_VOLUME_PATH
    
    Static DosDevices   As New Collection
    Static bInit       As Boolean
    
    If bInit Then
        If DosDevices.Count Then GoTo GetFromCollection
        Exit Function
    End If
    
    Dim aDrive()        As String
    Dim sDrives         As String
    Dim cnt             As Long
    Dim i               As Long
    Dim DosDeviceName   As String
    
    bInit = True
    
    cnt = GetLogicalDriveStrings(0&, StrPtr(sDrives))
    
    sDrives = Space$(cnt)
    
    cnt = GetLogicalDriveStrings(Len(sDrives), StrPtr(sDrives))
    
    If 0 = Err.LastDllError Then
    
        aDrive = Split(Left$(sDrives, cnt - 1), vbNullChar)
    
        For i = 0 To UBound(aDrive)
            
            DosDeviceName = Space$(MAX_PATH)
            
            cnt = QueryDosDevice(StrPtr(Left$(aDrive(i), 2)), StrPtr(DosDeviceName), Len(DosDeviceName))
            
            If cnt <> 0 Then
            
                DosDeviceName = Left$(DosDeviceName, InStr(DosDeviceName, vbNullChar) - 1)

                If Not isCollectionKeyExists(DosDeviceName, DosDevices) Then
                    DosDevices.Add aDrive(i), DosDeviceName
                End If

            End If
            
        Next
    
    End If

GetFromCollection:

    Dim pos As Long
    Dim sDrivePart As String
    Dim sOtherPart As String

    'Extract drive part
    If StrComp(Left$(inDosDeviceName, 8), "\Device\", 1) = 0 Then
        pos = InStr(9, inDosDeviceName, "\")
        If pos = 0 Then
            sDrivePart = inDosDeviceName
        Else
            sDrivePart = Left$(inDosDeviceName, pos - 1)
            sOtherPart = Mid$(inDosDeviceName, pos + 1)
        End If
        If isCollectionKeyExists(sDrivePart, DosDevices) Then
            ConvertDosDeviceToDriveName = BuildPath(DosDevices(sDrivePart), sOtherPart)
        End If
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "ConvertDosDeviceToDriveName"
    If inIDE Then Stop: Resume Next
End Function

Function GetFilePathByPID(PID As Long) As String
    On Error GoTo ErrorHandler:

    Const MAX_PATH_W                        As Long = 32767&
    Const PROCESS_VM_READ                   As Long = 16&
    Const PROCESS_QUERY_INFORMATION         As Long = 1024&
    Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000&
    
    Dim ProcPath    As String
    Dim hProc       As Long
    Dim cnt         As Long
    Dim pos         As Long
    Dim FullPath    As String

    hProc = OpenProcess(IIf(OSver.IsWindowsVistaOrGreater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION) Or PROCESS_VM_READ, 0&, PID)
    
    If hProc = 0 Then
        If Err.LastDllError = ERROR_ACCESS_DENIED Then
            hProc = OpenProcess(IIf(OSver.IsWindowsVistaOrGreater, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_QUERY_INFORMATION), 0&, PID)
        End If
    End If
    
    If hProc <> 0 Then
    
        If OSver.IsWindowsVistaOrGreater Then
            cnt = MAX_PATH_W + 1
            ProcPath = Space$(cnt)
            Call QueryFullProcessImageName(hProc, 0&, StrPtr(ProcPath), VarPtr(cnt))
        End If
        
        If 0 <> Err.LastDllError Or Not OSver.IsWindowsVistaOrGreater Then      'Win 2008 Server (x64) can cause Error 128 if path contains space characters
        
            ProcPath = Space$(MAX_PATH)
            cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
        
            If cnt = MAX_PATH Then 'Path > MAX_PATH -> realloc
                ProcPath = Space$(MAX_PATH_W)
                cnt = GetModuleFileNameEx(hProc, 0&, StrPtr(ProcPath), Len(ProcPath))
            End If
        End If
        
        If cnt <> 0 Then                          'clear path
            ProcPath = Left$(ProcPath, cnt)
            If StrComp("\SystemRoot\", Left$(ProcPath, 12), 1) = 0 Then ProcPath = Environ("SystemRoot") & Mid$(ProcPath, 12)
            If "\??\" = Left$(ProcPath, 4) Then ProcPath = Mid$(ProcPath, 5)
        Else
            ProcPath = ""
        End If
        
        If ERROR_PARTIAL_COPY = Err.LastDllError Or cnt = 0 Then     'because GetModuleFileNameEx cannot access to that information for 64-bit processes on WOW64
            ProcPath = Space$(MAX_PATH)
            cnt = GetProcessImageFileName(hProc, StrPtr(ProcPath), Len(ProcPath))
            
            If cnt <> 0 Then
                ProcPath = Left$(ProcPath, cnt)
                
                ' Convert DosDevice format to Disk drive format
                If StrComp(Left$(ProcPath, 8), "\Device\", 1) = 0 Then
                    pos = InStr(9, ProcPath, "\")
                    If pos <> 0 Then
                        FullPath = ConvertDosDeviceToDriveName(Left$(ProcPath, pos - 1))
                        If Len(FullPath) <> 0 Then
                            ProcPath = FullPath & Mid$(ProcPath, pos + 1)
                        End If
                    End If
                End If
            Else
                ProcPath = ""
            End If
            
        End If
        
        If Len(ProcPath) <> 0 Then    'if process ran with 8.3 style, GetModuleFileNameEx will return 8.3 style on x64 and full pathname on x86
                                      'so wee need to expand it ourself
            
            ProcPath = GetFullPath(ProcPath)
            GetFilePathByPID = GetLongPath(ProcPath)
        End If
        
        CloseHandle hProc
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePathByPID"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFullPath(sFilename As String) As String
    On Error GoTo ErrorHandler
    Dim cnt        As Long
    Dim sFullName  As String
    
    sFullName = String$(MAX_PATH_W, 0)
    cnt = GetFullPathName(StrPtr(sFilename), MAX_PATH_W, StrPtr(sFullName), 0&)
    If cnt Then
        GetFullPath = Left$(sFullName, cnt)
    Else
        GetFullPath = sFilename
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFullPath"
    If inIDE Then Stop: Resume Next
End Function

Public Function GetLongPath(sFile As String) As String '8.3 -> to Full name
    On Error GoTo ErrorHandler
    If InStr(sFile, "~") = 0 Then
        GetLongPath = sFile
        Exit Function
    End If
    Dim sBuffer As String, cnt As Long, pos As Long, sFolder As String
    
    If Not FileExists(sFile) And Not FolderExists(sFile) Then
        'try to convert folder struct instead, like C:\PROGRA~1\MICROS~1\Office15\ONBttnIE.dll (file missing)
        pos = InStrRev(sFile, "\", -1)
        If pos <> 0 Then
            Do
                sFolder = Left$(sFile, pos - 1)
                
                If InStr(sFolder, "~") = 0 Then Exit Do
                
                If FolderExists(sFolder) Then
                    GetLongPath = GetLongPath(sFolder) & "\" & Mid$(sFile, pos + 1)
                    Exit Do
                End If
                
                pos = pos - 1
                If pos <> 0 Then
                    pos = InStrRev(sFile, "\", pos)
                End If
                
            Loop While pos <> 0
        End If
        If GetLongPath = "" Then GetLongPath = sFile
        Exit Function
    End If
    
    sBuffer = String$(MAX_PATH_W, 0&)
    cnt = GetLongPathName(StrPtr(sFile), StrPtr(sBuffer), Len(sBuffer))
    If cnt Then
        GetLongPath = Left$(sBuffer, cnt)
    Else
        GetLongPath = sFile
    End If
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetLongPath", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function
