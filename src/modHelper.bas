Attribute VB_Name = "modHelper"
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
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
End Type

Private Type MIB_IPADDRROW
    dwAddr As Long
    dwIndex As Long
    dwMask As Long
    dwBCastAddr As Long
    dwReasmSize As Long
    unused1 As Integer
    wType As Integer
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenFilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameW" (pOpenFilename As OPENFILENAME) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" (src As Any, dst As Any) As Long
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetIpAddrTable Lib "IPHlpApi.dll" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Private Const SM_CXFULLSCREEN   As Long = 16&
Private Const SM_CYFULLSCREEN   As Long = 17&

Private Const RGN_OR            As Long = 2

Public inIDE    As Boolean
Public fX       As Single
Public fY       As Single

Public USE_DEBUG As Boolean
Public g_InstallState As Boolean
Public g_Separate As Boolean 'Music launched without game


Public Function FindPathL4d() As String
    
    ' we'll use tracing info :)
    
    ' 0. (most priority at first run)
    ' If exe placed on game folder.
    
    ' 1.
    ' "HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
    ' G:\Left 4 Dead\left4dead.exe.FriendlyAppName
    ' = left4dead.exe
    
    ' 2.
    ' HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\LowRegistry\Audio\PolicyConfig\PropertyStore\6f06ec5b_0
    ' def. param
    ' = {2}.\\?\hdaudio#func_01&ven_10ec&dev_0889&subsys_10438418&rev_1000#{6994ad04-93ef-11d0-a3cc-00a0c9223196}\singlelineouttopo/00010001|\Device\HarddiskVolume15\Left 4 Dead\left4dead.exe%b{00000000-0000-0000-0000-000000000000}
    
    ' 3.
    ' HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store
    ' G:\Left 4 Dead\left4dead.exe
    ' = binary data
    
    ' 4.
    ' HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
    ' {4E9C284F-809A-45A4-9831-A568B48751EA}
    ' = v2.25|Action=Allow|Active=TRUE|Dir=In|Protocol=6|App=U:\SteamApps\steamapps\common\Left 4 Dead 2 Beta\left4dead2_beta.exe|Name=Left 4 Dead 2 Beta|
    
    ' TCP Query User{73D5B9EA-D0DB-4CA4-87EE-EBEC0F85AC4F}L:\left 4 dead\left4dead.exe
    ' = v2.10|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Public|App=L:\l4d\left4dead.exe|Name=left4dead|Desc=left4dead|Defer=User|0
    
    ' 5.
    ' By running process: left4dead.exe
    
    'indentifier of L4d: g:\Left 4 Dead\left4dead\missions\hospital.txt
    
    'most actual file: g:\Left 4 Dead\bin\stats.bin
    
    Dim PROCESS_NAME As String
    ReDim Path(0) As String
    Dim sPath As String
    
    PROCESS_NAME = "left4dead.exe"
    
    '0. Application dir = game dir ?
    sPath = AppPath()
    
    If isL4dDir(sPath) Then
        AddToArray sPath, Path()
    End If
    
    '1. MuiCache
    Dim rParam() As String, rData(), rType() As Long, i As Long
    
    For i = 1 To Reg.GetRegValuesAndData(0&, "HKEY_CURRENT_USER\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache", _
        FLAG_REG_SZ, rParam, rData, rType)
        
        If StrComp(rData(i), PROCESS_NAME, 1) = 0 Then
            sPath = rParam(i)
            sPath = GetParentDir(sPath)
            If Not inArray(sPath, Path) Then
                If isL4dDir(sPath) Then AddToArray sPath, Path
            End If
        End If
    Next
    
    '2. Audio\PolicyConfig
    Dim rKey() As String, pos As Long, Disk As String
    
    For i = 1 To Reg.RegEnumSubkeysToArray(0&, _
        "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\LowRegistry\Audio\PolicyConfig\PropertyStore", rKey)
    
        sPath = Reg.GetRegData(0&, "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Internet Explorer\LowRegistry\Audio\PolicyConfig\PropertyStore" & _
            "\" & rKey(i), "")
        
        pos = InStr(1, sPath, "\" & PROCESS_NAME, 1)
        If pos <> 0 Then
            '...|\Device\HarddiskVolume15\Left 4 Dead\left4dead.exe%b...
            
            sPath = Left$(sPath, pos - 1) ' trim file name
            
            pos = InStr(1, sPath, "\Device\HarddiskVolume", 1)
            If pos <> 0 Then
                sPath = Mid$(sPath, pos)
                
                pos = InStr(9, sPath, "\")
                If pos <> 0 Then
                    Disk = Left$(sPath, pos - 1)
                    Disk = ConvertDosDeviceToDriveName(Disk)
                    sPath = Disk & Mid$(sPath, pos + 1)
                    
                    If Not inArray(sPath, Path) Then
                        If isL4dDir(sPath) Then AddToArray sPath, Path
                    End If
                End If
            End If
        End If
    Next
    
    '3. AppCompatFlags
    Erase rParam
    For i = 1 To Reg.GetEnumValuesToArray(0&, "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store", _
        rParam)
        
        If InStr(1, rParam(i), "\" & PROCESS_NAME, 1) <> 0 Then
            
            sPath = GetParentDir(rParam(i))
            
            If Not inArray(sPath, Path) Then
                If isL4dDir(sPath) Then AddToArray sPath, Path
            End If
        End If
    Next
    
    '4. FirewallRules
    Erase rParam, rData, rType
    
    For i = 1 To Reg.GetRegValuesAndData(0&, "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules", _
        FLAG_REG_SZ, rParam, rData, rType)
    
        '...|App=L:\l4d\left4dead.exe|...
        
        If InStr(1, rData(i), "\" & PROCESS_NAME, 1) <> 0 Then
        
            sPath = rData(i)
            pos = InStr(1, sPath, "App=", 1)
            If pos <> 0 Then
                sPath = Mid$(sPath, pos + 4)
                pos = InStr(sPath, "|")
                If pos <> 0 Then
                    sPath = Left$(sPath, pos - 1)
                End If
                sPath = GetParentDir(rParam(i))
                If Not inArray(sPath, Path) Then
                    If isL4dDir(sPath) Then AddToArray sPath, Path
                End If
            End If
        End If
    Next
    
    '5. By launched process
    Dim PID As Long
    If Proc.IsRunned(PROCESS_NAME, PID) Then
        If PID <> 0 Then
            sPath = GetFilePathByPID(PID)
            
            sPath = GetParentDir(sPath)
            
            If Not inArray(sPath, Path) Then
                If isL4dDir(sPath) Then AddToArray sPath, Path
            End If
        End If
    End If
    
    For i = 0 To UBound(Path)
        If 0 <> Len(Path(i)) Then
            FindPathL4d = Path(i)
            Exit Function
        End If
    Next
    
End Function

Public Function UnpackResource(ResourceID As Long, DestinationPath As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim ff      As Integer
    Dim b()     As Byte
    UnpackResource = True
    b = LoadResData(ResourceID, "CUSTOM")
    ff = FreeFile
    Open DestinationPath For Binary Access Write As #ff
        Put #ff, , b
    Close #ff
    Exit Function
ErrorHandler:
    MsgBox Err & " " & "UnpackResource" & " " & "ID: " & ResourceID & " " & "Destination path: " & DestinationPath
    UnpackResource = False
    'If inIDE Then Stop: Resume Next
End Function

Public Sub CenterForm(myForm As Form) ' Центрирование формы на экране с учетом системных панелей
    On Error Resume Next
    Dim Left    As Long
    Dim Top     As Long
    Left = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CXFULLSCREEN) / 2 - myForm.Width / 2
    Top = Screen.TwipsPerPixelY * GetSystemMetrics(SM_CYFULLSCREEN) / 2 - myForm.Height / 2
    myForm.Move Left, Top
End Sub
