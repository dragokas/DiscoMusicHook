Attribute VB_Name = "modOpenFolder"
' Open Folder dialogue by The Trick

' Fork by Dragokas:
' Added Unicode Support
' OpenFolderDialog now returns collection of pathes

'// TODO: add code based on IFileOpenDialog interface (for Vista+), so you'll be able to set initial dir as special folder 'This PC'

Option Explicit

Private Type OPENFILENAME_W
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type
Private Enum CdlgExt_Flags
    OFNAllowMultiselect = &H200
    OFNCreatePrompt = &H2000
    OFNexplorer = &H80000
    OFNEnableHook = &H20
    OFNExtensionDifferent = &H400
    OFNFileMustExist = &H1000
    OFNHelpButton = &H10
    OFNHideReadOnly = &H4
    OFNLongNames = &H200000
    OFNNoChangeDir = &H8
    OFNNoDereferenceLinks = &H100000
    OFNNoLongNames = &H40000
    OFNNoReadOnlyReturn = &H8000
    OFNNoValidate = &H100
    OFNOverwritePrompt = &H2
    OFNPathMustExist = &H800
    OFNReadOnly = &H1
    OFNShareAware = &H4000
    OFNEnableIncludeNotify = &H400000
End Enum
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
Private Type LVITEM_W
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
Private Const GWL_WNDPROC = (-4)
 
Private Const WM_INITDIALOG = &H110
Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E
Private Const WM_USER = &H400
Private Const WM_COMMAND = &H111
 
Private Const CDN_FIRST = -601&
Private Const CDN_INITDONE = (CDN_FIRST - 0&)
Private Const CDN_FILEOK = (CDN_FIRST - 5&)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
 
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
 
Private Const BN_CLICKED As Long = &H0
 
Private Const MAX_PATH = 260
 
Private Const IDOK = 1
Private Const IDFILETYPECOMBO = &H470
Private Const IDFILETYPESTATIC = &H441      ' Files of Type
Private Const IDFILENAMESTATIC = &H442      ' File Name
Private Const IDFILELIST = &H460            ' Listbox
Private Const IDFILENAMECOMBO = &H47C       ' Combo
 
Private Const LVM_FIRST = &H1000&
Private Const LVM_GETSELECTEDCOUNT = LVM_FIRST + 50
Private Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
'Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
Private Const LVM_GETITEMTEXTW = 4211
 
Private Const LVIS_SELECTED = &H2&
 
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenFilename As OPENFILENAME_W) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal Count As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMem2 Lib "msvbvm60" (pSrc As Any, pDst As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (pOpenFilename As OPENFILENAME_W) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function SHCreateShellItem Lib "shell32" (ByVal pidlParent As Long, ByVal psfParent As Long, ByVal pidl As Long, ppsi As IShellItem) As Long

Dim OFN As OPENFILENAME_W
Dim OldWndProc As Long
Dim hwndDlg As Long
Dim mFolders As Collection
Dim mPath As String

Public Function OpenFolderDialog(ByVal Owner As Long, Optional InitialFolder As String = "") As String
    Call PickFolder(Owner, False, InitialFolder)
    If mFolders.Count = 0 Then
        OpenFolderDialog = mPath
    Else
        OpenFolderDialog = mFolders(1)
    End If
End Function

Public Function OpenFolderDialogMultiSelect(ByVal Owner As Long, Optional InitialFolder As String = "") As Collection
    Call PickFolder(Owner, True, InitialFolder)
    If mFolders.Count = 0 Then
        mFolders.Add mPath
    End If
    Set OpenFolderDialogMultiSelect = mFolders
End Function

Private Function PickFolder(ByVal Owner As Long, AllowMultiSelect As Boolean, InitialFolder As String) As String
 
    If mFolders Is Nothing Then Set mFolders = New Collection
    Do While mFolders.Count: mFolders.Remove (1): Loop
    
    With OFN
        .lStructSize = Len(OFN)
        '.hwndOwner = Owner
        .hInstance = App.hInstance
        .lpfnHook = lHookAddress(AddressOf DialogHookFunction)
        .flags = OFNexplorer Or OFNNoChangeDir Or OFNEnableHook Or OFNEnableIncludeNotify Or OFNHideReadOnly
        If AllowMultiSelect Then .flags = .flags Or OFNAllowMultiselect
        .lpstrFile = StrPtr(String$(MAX_PATH, 0))
        .nMaxFile = MAX_PATH
        .lpstrFileTitle = StrPtr(String$(MAX_PATH, 0))
        .nMaxFileTitle = MAX_PATH
        .lpstrFilter = StrPtr("Folders" & Chr$(0) & "*." & String$(2, Chr$(0)))
        .lpstrTitle = StrPtr("Выбор папки")
        .nFilterIndex = 0
        .lpstrInitialDir = StrPtr(InitialFolder)
    End With
    mPath = vbNullString
    GetOpenFileName OFN
End Function
 
Private Function lHookAddress(lPtr As Long) As Long
    lHookAddress = lPtr
End Function

Private Function DialogHookFunction(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg
        Case WM_INITDIALOG
            hwndDlg = GetParent(hDlg)
            OldWndProc = SetWindowLong(hwndDlg, GWL_WNDPROC, AddressOf DlgWndProc)
        Case WM_NOTIFY
            Dim tNMH As NMHDR
            CopyMemory tNMH, ByVal lParam, Len(tNMH)
            Select Case tNMH.code
            Case CDN_INITDONE
                SendMessage hwndDlg, CDM_SETCONTROLTEXT, IDOK, ByVal StrPtr("Выбрать")
                SendMessage hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMESTATIC, ByVal StrPtr("") 'Надпись "Имя папки"
                SendMessage hwndDlg, CDM_HIDECONTROL, IDFILETYPECOMBO, ByVal 0&
                SendMessage hwndDlg, CDM_HIDECONTROL, IDFILETYPESTATIC, ByVal 0&
                SendMessage hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMECOMBO, ByVal StrPtr(GetPath)
                SetWindowPos hwndDlg, 0, 100, 100, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
            Case CDN_INCLUDEITEM
                ' Фильтруем
                DialogHookFunction = 0
            Case CDN_SELCHANGE
                SendMessage hwndDlg, CDM_SETCONTROLTEXT, IDFILENAMECOMBO, ByVal StrPtr(GetPath)
            End Select
        Case WM_DESTROY
        Case Else
    End Select
End Function

Private Function DlgWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
    Case WM_COMMAND
        If HiWord(wParam) = BN_CLICKED Then
            Dim hwndPick As Long
                        
            hwndPick = GetDlgItem(hwndDlg, IDOK)
                        
            If lParam = hwndPick Then
                Dim hwndLVParent As Long, hwndLV As Long
                Dim pos As Long, txtLen As Long, tmp As String
                Dim Itm As LVITEM_W
                
                hwndLVParent = FindWindowEx(hwndDlg, ByVal 0&, "SHELLDLL_DefView", vbNullString)
                hwndLV = FindWindowEx(hwndLVParent, ByVal 0&, "SysListView32", vbNullString)
 
                pos = SendMessage(hwndLV, LVM_GETNEXTITEM, -1, ByVal LVIS_SELECTED)
                
                If pos >= 0 Then
                    
                    mPath = String(MAX_PATH, vbNullChar)
                    txtLen = SendMessage(hwndDlg, CDM_GETFOLDERPATH, MAX_PATH, ByVal StrPtr(mPath))
                    mPath = Left(mPath, txtLen - 1)
                    
                    Itm.cchTextMax = MAX_PATH
                    Itm.pszText = StrPtr(String(MAX_PATH, vbNullChar))
                    
                    txtLen = SendMessage(hwndLV, LVM_GETITEMTEXTW, pos, Itm)
                    
                    'mFolders.Add Left(Itm.pszText, txtLen)
                    If txtLen <> 0 Then
                        tmp = String$(txtLen, vbNullChar)
                        lstrcpyn StrPtr(tmp), Itm.pszText, txtLen + 1
                    End If
                    mFolders.Add mPath & "\" & tmp
                    
                    Do Until pos = -1
                        pos = SendMessage(hwndLV, LVM_GETNEXTITEM, pos, ByVal LVIS_SELECTED)
                        txtLen = SendMessage(hwndLV, LVM_GETITEMTEXTW, pos, Itm)
                        'If Pos >= 0 Then mFolders.Add Left(Itm.pszText, txtLen)
                        If pos >= 0 Then
                            If txtLen <> 0 Then
                                tmp = String$(txtLen, vbNullChar)
                                lstrcpyn StrPtr(tmp), Itm.pszText, txtLen + 1
                            End If
                            mFolders.Add mPath & "\" & tmp
                        End If
                    Loop
                    
                    DestroyWindow hwndDlg
                Else
                    mPath = GetPath()
                    If Len(mPath) Then
                        EndDialog hwndDlg, 0
                    End If
                End If
            Else
                DlgWndProc = CallWindowProc(OldWndProc, hwnd, Msg, wParam, lParam)
            End If
        End If
    Case Else
        DlgWndProc = CallWindowProc(OldWndProc, hwnd, Msg, wParam, lParam)
    End Select
End Function
 
Private Function GetPath() As String
    Dim txtLen As Long, tmp As String
        
    tmp = String(MAX_PATH, vbNullChar)
    
    txtLen = SendMessage(hwndDlg, CDM_GETFILEPATH, MAX_PATH, ByVal StrPtr(tmp))
    
    If txtLen > 0 Then
        tmp = Left(tmp, txtLen - 1)
        If GetFileAttributes(StrPtr(tmp)) And vbDirectory Then GetPath = tmp
    End If
End Function
 
Private Function LoWord(ByVal LongIn As Long) As Integer
    GetMem2 LongIn, LoWord
    'Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function HiWord(ByVal LongIn As Long) As Integer
    GetMem2 ByVal VarPtr(LongIn) + 2, HiWord
    'Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function GetOpenFile2(hParent As Long, Optional InitDir As String = "") As String
    'Shows the simplest Open File Dialog
    On Error Resume Next 'A major error is thrown when the user cancels the dialog box
    
    Dim fodSimple As FileOpenDialog
    Dim isiRes As IShellItem
    Dim isiDef As IShellItem
    Dim pidlDef As Long
    Dim lPtr As Long
    Dim FileFilter() As COMDLG_FILTERSPEC
    ReDim FileFilter(0)
    
    FileFilter(0).pszName = "Исполняемые файлы"
    FileFilter(0).pszSpec = "*.exe"
    
    Set fodSimple = New FileOpenDialog
    
    
    If Len(InitDir) <> 0 Then
        pidlDef = GetPIDLFromPathW(InitDir)
    End If
    If pidlDef = 0 Then
        Call SHGetKnownFolderIDList(FOLDERID_ComputerFolder, 0, 0, pidlDef)
    End If
    If pidlDef Then
        Call SHCreateShellItem(0, 0, pidlDef, isiDef)
    End If
    
    With fodSimple
        .SetTitle "Выберите запускаемый файл игры Left4dead"
        .SetFileTypes UBound(FileFilter) + 1, VarPtr(FileFilter(0).pszName)
        '.SetDefaultFolder
        If Not (isiDef Is Nothing) Then .SetFolder isiDef
        
        .Show hParent
        
        .GetResult isiRes
        isiRes.GetDisplayName SIGDN_FILESYSPATH, lPtr
        'GetOpenFile2 = BStrFromLPWStr(lPtr, True)
        GetOpenFile2 = StringFromPtrW(lPtr)
        
    End With
    Set isiRes = Nothing
    Set fodSimple = Nothing
    Set isiDef = Nothing

End Function

Private Function GetPIDLFromPathW(sPath As String) As Long
   GetPIDLFromPathW = ILCreateFromPathW(StrPtr(sPath))
End Function

' Диалог сохранения файла и получить имя выбранного файла
Public Function GetSaveFile(ByVal hwnd As Long) As String
    Dim OFN As OPENFILENAME_W
    Dim Title As String, Out As String
    Dim Filter As String, i As Long
    
    OFN.nMaxFile = 260
    Out = String(260, vbNullChar)
    Title = "Сохранить файл"
    Filter = "Все файлы" & vbNullChar & "*.*" & vbNullChar
    OFN.hWndOwner = hwnd
    OFN.lpstrTitle = StrPtr(Title)
    OFN.lpstrFile = StrPtr(Out)
    OFN.lStructSize = Len(OFN)
    OFN.lpstrFilter = StrPtr(Filter)
    OFN.lpstrInitialDir = StrPtr(App.Path)
    If GetSaveFileName(OFN) Then
        i = InStr(1, Out, vbNullChar, vbBinaryCompare)
        If i Then GetSaveFile = Left$(Out, i - 1)
    End If
End Function

' Диалог открытия файла и получить имя выбранного файла
Public Function GetOpenFile(ByVal hwnd As Long, Optional sInitDir$) As String
    Dim OFN As OPENFILENAME_W
    Dim Title As String, Out As String
    Dim Filter As String, i As Long
    
    OFN.nMaxFile = 260
    Out = String(260, vbNullChar)
    Title = "Выберите запускаемый файл игры Left4dead"
    Filter = "Исполняемые файлы" & vbNullChar & "*.exe" & vbNullChar
    OFN.hWndOwner = hwnd
    OFN.lpstrTitle = StrPtr(Title)
    OFN.lpstrFile = StrPtr(Out)
    OFN.lStructSize = Len(OFN)
    OFN.lpstrFilter = StrPtr(Filter)
    OFN.lpstrInitialDir = IIf(sInitDir = "", StrPtr(App.Path), StrPtr(sInitDir))
    If GetOpenFileName(OFN) Then
        i = InStr(1, Out, vbNullChar, vbBinaryCompare)
        If i Then GetOpenFile = Left$(Out, i - 1)
    End If
End Function

'Private Function BStrFromLPWStr(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
'SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
'If CleanupLPWStr Then CoTaskMemFree lpWStr
'End Function

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

