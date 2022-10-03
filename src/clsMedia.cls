VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�������� ����� ���������, ������������ ��� ������ � MCI

'� Device Types

'    Device type    Constant                    Description

'    cdaudio        MCI_DEVTYPE_CD_AUDIO        CD audio player
'    dat            MCI_DEVTYPE_DAT             Digital-audio tape player
'    digitalvideo   MCI_DEVTYPE_DIGITAL_VIDEO   Digital video in a window (not GDI-based)
'    other          MCI_DEVTYPE_OTHER           Undefined MCI device
'    overlay        MCI_DEVTYPE_OVERLAY         Overlay device (analog video in a window)
'    scanner        MCI_DEVTYPE_SCANNER         Image scanner
'    sequencer      MCI_DEVTYPE_SEQUENCER       MIDI sequencer
'    vcr            MCI_DEVTYPE_VCR             Video-cassette recorder or player
'    videodisc      MCI_DEVTYPE_VIDEODISC       Videodisc player
'    waveaudio      MCI_DEVTYPE_WAVEFORM_AUDIO  Audio device that plays digitized waveform files


'������������� � �������� ������, ������������ MCI

'� System Commands

'    String         Message                     Description

'    break          MCI_BREAK                   Sets a break key for an MCI device.
'    sysinfo        MCI_SYSINFO                 Returns information about MCI devices.


'� Required Commands

'    String         Message                     Description

'    capability     MCI_GETDEVCAPS              Obtains the capabilities of a device.
'    close          MCI_CLOSE                   Closes the device.
'    info           MCI_INFO                    Obtains textual information from a device.
'    open           MCI_OPEN                    Initializes the device.
'    status         MCI_STATUS                  Obtains status information from the device. Some of this command's flags are
'                                               not required, so it is also a basic command.


'� Basic Commands

'    String         Message                     Description

'    load           MCI_LOAD                    Loads data from a file.
'    pause          MCI_PAUSE                   Stops playing. Playback or recording can be resumed at the current position.
'    play           MCI_PLAY                    Starts transmitting output data.
'    record         MCI_RECORD                  Starts recording input data.
'    resume         MCI_RESUME                  Resumes playing or recording on a paused device.
'    save           MCI_SAVE                    Saves data to a disk file.
'    seek           MCI_SEEK                    Seeks forward or backward.
'    set            MCI_SET                     Sets the operating state of the device.
'    status         MCI_STATUS                  Obtains status information about the device. This is also a required command;
'                                               since some of its flags are not required, it is also listed here.
'                                               (The optional items support devices that use linear media with identifiable positions.)
'    stop           MCI_STOP                    Stops playing.


'� Extended Commands

'    String         Message                     Device types                Description

'    configure      MCI_CONFIGURE               digitalvideo                Displays a configuration dialog box.
'    cue            MCI_CUE                     digitalvideo, waveaudio     Prepares for playing or recording.
'    delete         MCI_DELETE                  waveaudio                   Deletes a data segment from the media file.
'    escape         MCI_ESCAPE                  videodisc                   Sends custom information to a device.
'    freeze         MCI_FREEZE                  overlay                     Disables video acquisition to the frame buffer.
'    put            MCI_PUT                     digitalvideo, overlay       Defines the source, destination, and frame windows.
'    realize        MCI_REALIZE                 digitalvideo                Tells the device to select and realize its palette into a device
'                                                                           context of the displayed window.
'    setaudio       MCI_SETAUDIO                digitalvideo                Sets audio parameters for video.
'    setvideo       MCI_SETVIDEO                digitalvideo                Sets video parameters.
'    signal         MCI_SIGNAL                  digitalvideo                Identifies a specified position with a signal.
'    spin           MCI_SPIN                    videodisc                   Starts the disc spinning or stops the disc from spinning.
'    step           MCI_STEP                    digitalvideo, videodisc     Steps the play one or more frames forward or reverse.
'    unfreeze       MCI_UNFREEZE                overlay                     Enables the frame buffer to acquire video data.
'    update         MCI_UPDATE                  digitalvideo                Repaints the current frame into the device context.
'    where          MCI_WHERE                   digitalvideo, overlay       Obtains the rectangle specifying the source, destination, or frame area.
'    window         MCI_WINDOW                  digitalvideo, overlay       Controls the display window.

'�������������� ���������� ����� ����� �� http://msdn.microsoft.com

Rem ////////////////////////////////////////////////////////////////
Rem     ������: clsMedia (����� ��� ������ � MCI)
Rem     �����: BV (Boris Vorontsov, borisbox@mail.ru)
Rem ////////////////////////////////////////////////////////////////

'��������� ��� ������ MCI � ������
Private Const MCI_ALIAS As String = "Media_"

'�������� ������ (������ ��� ������)
Public Enum Channels
    mciLeftChannel& = 0
    mciRightChannel& = 1
End Enum

'������ ������
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal ErrorNumber As Long, ByVal ReturnBuffer As String, ByVal ReturnBufferSize As Long) As Long

Private Const MCIERR_BASE As Long = 256
Private Const MCIERR_MISSING_COMMAND_STRING As Long = (MCIERR_BASE + 11)
Private Const MCIERR_FILE_NOT_FOUND As Long = (MCIERR_BASE + 19)

Dim bNotify As Boolean '���� Notify
Dim bShareable As Boolean '���� Shareable
Dim bWait As Boolean '���� Wait
Dim bFullscreen As Boolean '���� Fullscreen
Dim bRepeat As Boolean '���� Repeat

Dim lStartPosition As Long '������� ������ ���������������
Dim lEndPosition As Long '������� ����� ���������������

Dim strDeviceType As String '��� ����������
Dim strFileName As String '���������� ��� ����� ��� ������
Dim hWndCallback As Long '����, �������� ����� ������������ ���������
Dim hVideoWnd As PictureBox '����, ������ ����� ���������������� �����
Dim lRetval() As Long '������������ �������� ������� mciSendString

'������������� ������
Private Sub Class_Initialize()
    bNotify = False
    bShareable = False
    bWait = False
    bFullscreen = False
    bRepeat = False
    lStartPosition = -1 '�������� -1 ������� � ���, ��� ��� ����c��� �� ����� ���������
    lEndPosition = -1
    strDeviceType = vbNullString
    strFileName = vbNullString
    hWndCallback = 0
End Sub

'�������� ��������� �����
'�������� strMCIAlias ����� ��������� (�� ��������� MCI_ALIAS)
Public Sub mOpen(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    Dim strOpenCommand As String
    '���� ���� �� ����� ��� �� ����������, ����� ������ ������ � �������
    If strFileName = vbNullString Or Dir$(strFileName) = vbNullString Then
        lRetval(0) = MCIERR_FILE_NOT_FOUND
        Exit Sub
    End If
    '��������� - ���� � ����� ����� ���� �������, �� ���� ��� � �������
    If InStr(strFileName, Chr$(32)) <> 0 Then strFileName = Chr$(34) & strFileName & Chr$(34)
    '���� ����� (������ ����), ��������� ��� �� �������� ����
    If Not hVideoWnd Is Nothing Then
        ReDim lRetval(3) As Long '����� ������ ������� ������ ���-�� ������� �������
        '������ ����������� ������ (�������� �����)
        strOpenCommand = "Open " & strFileName & " alias " & strMCIAlias & " parent " & hVideoWnd.hWnd & _
            " style child"
        '���� �������� "DeviceType" ������, ��������� ��� � ������
        If strDeviceType <> vbNullString Then _
            strOpenCommand = strOpenCommand & " type " & strDeviceType
        '��������� �����, ���� �����...
        If bNotify Then strOpenCommand = strOpenCommand & " notify"
        If bShareable Then strOpenCommand = strOpenCommand & " shareable"
        If bWait Then strOpenCommand = strOpenCommand & " wait"
        
        lRetval(0) = mciSendString(strOpenCommand, vbNullString, 0&, hWndCallback)
        '����������� �����
        Dim strReturn As String * 128
        Dim lWidth As Long, lHeight As Long
        Dim lPos(0 To 1) As Long
        lRetval(1) = mciSendString("Where " & strMCIAlias & " destination", ByVal strReturn, _
            Len(strReturn) - 1, hWndCallback)
        
        '���� ����� ������� �������� �������� ������, ������ ���� �� ��������
        '���������� ������ ����� � ������������ ����������
        If lRetval(1) = 0 Then
            lPos(0) = InStr(1, strReturn, Chr$(32))
            lPos(1) = InStr(lPos(0) + 1, strReturn, Chr$(32))
            lPos(0) = InStr(lPos(1) + 1, strReturn, Chr$(32))
            
            lWidth = Mid$(strReturn, lPos(1), lPos(0) - lPos(1)) * Screen.TwipsPerPixelX
            lHeight = Mid$(strReturn, lPos(0) + 1) * Screen.TwipsPerPixelY
            '��������� ������� (�� ������ ��������� PictureBox)
            lRetval(2) = mciSendString("Put " & strMCIAlias & " window at " & (hVideoWnd.Width - lWidth) \ _
                (2 * Screen.TwipsPerPixelX) & Chr$(32) & (hVideoWnd.Height - lHeight) \ _
                (2 * Screen.TwipsPerPixelY) & Chr$(32) & lWidth \ Screen.TwipsPerPixelX & Chr$(32) & _
                lHeight \ Screen.TwipsPerPixelY, vbNullString, 0&, hWndCallback)
        End If
        '��� ����, ����� ���������� �������� ����� ��� ������ ���� ������, ����������
        '���������� ������ ������� � �������������
        lRetval(3) = mciSendString("Set " & strMCIAlias & " time format milliseconds", _
            vbNullString, 0&, hWndCallback)
    Else
        ReDim lRetval(1) As Long
        '������ ����������� ������, �� ��� ����� ���� ��� �����
        strOpenCommand = "Open " & strFileName & " alias " & strMCIAlias
        
        If strDeviceType <> vbNullString Then _
            strOpenCommand = strOpenCommand & " type " & strDeviceType
        
        If bNotify Then strOpenCommand = strOpenCommand & " notify"
        If bShareable Then strOpenCommand = strOpenCommand & " shareable"
        If bWait Then strOpenCommand = strOpenCommand & " wait"
        
        lRetval(0) = mciSendString(strOpenCommand, vbNullString, 0&, hWndCallback)
        lRetval(1) = mciSendString("Set " & strMCIAlias & " time format milliseconds", _
            vbNullString, 0&, hWndCallback)
    End If
End Sub

'������� ���������������
Public Sub mPlay(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    Dim strPlayCommand As String
    '��� ������ �������� ���������� �������� ����������� ������� ���������������
    strPlayCommand = "Play " & strMCIAlias
    '��������� � strPlayCommand ��������������� �������
    If lStartPosition > -1 Then strPlayCommand = strPlayCommand & " from " & Str$(lStartPosition)
    If lEndPosition > -1 Then strPlayCommand = strPlayCommand & " to " & Str$(lEndPosition)
    '��������� �����...
    If bRepeat Then strPlayCommand = strPlayCommand & " repeat"
    If bFullscreen Then strPlayCommand = strPlayCommand & " fullscreen"
    lRetval(0) = mciSendString(strPlayCommand, vbNullString, 0&, hWndCallback)
End Sub

'����� ��� �������� ���������������
Public Sub mPause(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Pause " & strMCIAlias, vbNullString, 0&, hWndCallback)
End Sub

'����������� �������� ���������������
Public Sub mResume(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Resume " & strMCIAlias, vbNullString, 0&, hWndCallback)
End Sub

'��������� �������� ���������������
Public Sub mStop(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Stop " & strMCIAlias, vbNullString, 0&, hWndCallback)
End Sub

'�������� �������� ���������������
Public Sub mClose(Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Close " & strMCIAlias, vbNullString, 0&, hWndCallback)
    '������� ��� �����
    strFileName = vbNullString
End Sub

'�������� �����
Public Sub mCloseAll()
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Close all", vbNullString, 0&, hWndCallback)
    strFileName = vbNullString
End Sub

'������� �� �������� ������� ��� �������� ���������������
Public Sub mSeek(ByVal sngPosition As Single, _
    Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    lRetval(0) = mciSendString("Seek " & strMCIAlias & " to " & sngPosition, _
        vbNullString, 0&, hWndCallback)
End Sub

'������ ������ � �������: Notify, Shareable, Wait...
Public Property Get Notify() As Boolean
    Notify = bNotify
End Property

Public Property Let Notify(ByVal bNotifyFlag As Boolean)
    bNotify = bNotifyFlag
End Property

Public Property Get Shareable() As Boolean
    Shareable = bShareable
End Property

Public Property Let Shareable(ByVal bShareableFlag As Boolean)
    bShareable = bShareableFlag
End Property

Public Property Get Wait() As Boolean
    Wait = bWait
End Property

Public Property Let Wait(ByVal bWaitFlag As Boolean)
    bWait = bWaitFlag
End Property

Public Property Get Fullscreen() As Boolean
    Fullscreen = bFullscreen
End Property

Public Property Let Fullscreen(ByVal bFullscreenFlag As Boolean)
    bFullscreen = bFullscreenFlag
End Property

Public Property Get Repeat() As Boolean
    Repeat = bRepeat
End Property

Public Property Let Repeat(ByVal bRepeatFlag As Boolean)
    bRepeat = bRepeatFlag
End Property

'������ ������ �� ������� ������ ���������������
Public Property Get StartPosition() As Long
    StartPosition = lStartPosition
End Property

'��������� ������� ������ ���������������
Public Property Let StartPosition(ByVal lNewStartPosition As Long)
    lStartPosition = lNewStartPosition
End Property

'������ ������ �� ������� ����� ���������������
Public Property Get EndPosition() As Long
    EndPosition = lEndPosition
End Property

'��������� ������� ����� ���������������
Public Property Let EndPosition(ByVal lNewEndPosition As Long)
    lEndPosition = lNewEndPosition
End Property

'������ ������ �� ��� ����������
Public Property Get DeviceType() As String
    DeviceType = strDeviceType
End Property

'���������� ���� ����������
Public Property Let DeviceType(ByVal strNewDeviceType As String)
    strDeviceType = strNewDeviceType
End Property

'������ ������ �� ��� �����
Public Property Get FileName() As String
    FileName = strFileName
End Property

'���������� ����� �����
Public Property Let FileName(ByVal strNewFileName As String)
    strFileName = strNewFileName
End Property

'���������� ���� ��������� ������
Public Property Get WndCallback() As Long
    WndCallback = hWndCallback
End Property

'������ ������ �� ���� ��������� ������
Public Property Let WndCallback(ByVal hNewWndCallback As Long)
    hWndCallback = hNewWndCallback
End Property

'������ ������ �� ���� ��� �����
Public Property Get VideoWnd() As PictureBox
    Set VideoWnd = hVideoWnd
End Property

'���������� ���� ��� �����
Public Property Let VideoWnd(ByVal hNewVideoWnd As PictureBox)
    Set hVideoWnd = hNewVideoWnd
End Property

'������ ������ �� ����� ��� �������� ���������������
Public Property Get Length(Optional ByVal strMCIAlias As String = MCI_ALIAS) As Single
    ReDim lRetval(0) As Long
    If strFileName = vbNullString Then lRetval(0) = 1: Exit Property
    Dim intLength As Integer
    Dim strLength As String * 255
    
    lRetval(0) = mciSendString("Status " & strMCIAlias & " length", _
        strLength, 255, hWndCallback)
    '�������� �� ���������� ������ ������ �����
    intLength = InStr(strLength, vbNullChar)
    Length = Val(Left$(strLength, intLength - 1))
End Property

'������ ������ �� ��������� �������� ���������������
Public Property Get Mode(Optional ByVal strMCIAlias As String = MCI_ALIAS) As String
    ReDim lRetval(0) As Long
    If strFileName = vbNullString Then lRetval(0) = 1: Exit Property
    Dim intMode As Integer
    Dim strMode As String * 255
    
    lRetval(0) = mciSendString("Status " & strMCIAlias & " mode", _
        strMode, 255, hWndCallback)

    intMode = InStr(strMode, vbNullChar)
    Mode = Val(Left$(strMode, intMode - 1))
End Property

'������ ������ �� ������� ��� �������� ���������������
Public Property Get Position(Optional ByVal strMCIAlias As String = MCI_ALIAS) As Single
    ReDim lRetval(0) As Long
    If strFileName = vbNullString Then lRetval(0) = 1: Exit Property
    Dim intPosition As Integer
    Dim strPosition As String * 255
    
    lRetval(0) = mciSendString("Status " & strMCIAlias & " position", _
        strPosition, 255, hWndCallback)

    intPosition = InStr(strPosition, vbNullChar)
    Position = Val(Left$(strPosition, intPosition - 1))
End Property

'������ ������ �� ��������� ������������ ���� �� ��������� ����������
Public Property Get WndHandle(Optional ByVal strMCIAlias As String = MCI_ALIAS) As Long
    ReDim lRetval(0) As Long
    Dim strHandle As String * 255
    Dim intPosition As Integer

    lRetval(0) = mciSendString("Status " & strMCIAlias & " window handle", _
        strHandle, 255, hWndCallback)
    
    intPosition = InStr(strHandle, vbNullChar)
    WndHandle = Val(Left$(strHandle, intPosition - 1))
End Property

'������� ��������� ��������� ������
Public Sub DeviceDoor(ByVal bOpen As Boolean, Optional ByVal bUseDefaultDeviceType As Boolean = True)
    ReDim lRetval(0) As Long
    Dim strOldDeviceType As String
    '���� ������ ���� ������������� ���� ���������� �� ���������, ��
    '�������� ������ ��� � ����� "CDAudio"
    If bUseDefaultDeviceType Then
        strOldDeviceType = strDeviceType
        strDeviceType = "CDAudio"
    End If
    '��������/�������� ����� ��� CD
    lRetval(0) = mciSendString("Set " & strDeviceType & " door " & _
        IIf(bOpen, "open", "closed"), vbNullString, 0&, hWndCallback)
    If bUseDefaultDeviceType Then strDeviceType = strOldDeviceType
End Sub

'������� ������������� ��������� ��� �������� ���������������
'�����: ������� ���������� ����� ��������� �����
Public Sub SetVolume(ByVal lLeftVolume As Long, ByVal lRightVolume As Long, _
    Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(1) As Long
    '��� ������ ������
    lRetval(0) = mciSendString("SetAudio " & strMCIAlias & " left volume to " & _
        lLeftVolume, vbNullString, 0&, hWndCallback)
    '��� ������� ������
    lRetval(1) = mciSendString("SetAudio " & strMCIAlias & " right volume to " & _
        lRightVolume, vbNullString, 0&, hWndCallback)
End Sub

'���������/���������� ������� ���������������
'�����: ������� ���������� ����� ��������� �����
Public Sub SetMute(ByRef lChannel As Channels, ByVal bOn As Boolean, _
    Optional ByVal strMCIAlias As String = MCI_ALIAS)
    ReDim lRetval(0) As Long
    Select Case lChannel
        Case mciLeftChannel '����� �����
            lRetval(0) = mciSendString("SetAudio " & strMCIAlias & " left " & _
                IIf(bOn, "on", "off"), vbNullString, 0&, hWndCallback)
        Case mciRightChannel '������ �����
            lRetval(0) = mciSendString("SetAudio " & strMCIAlias & " right " & _
                IIf(bOn, "on", "off"), vbNullString, 0&, hWndCallback)
    End Select
End Sub

'������� �������� MCI ��������� ������������� �������
'���������� �������� �� "lpstrReturnString"
Public Function SendCommand(ByVal strCommand As String, Optional ByVal intReturnLength As Integer = 0&) As String
    ReDim lRetval(0) As Long
    Dim strReturn As String
    Dim intPosition As Integer
    '��������� ������������ ���������
    If strCommand = vbNullString Then
        lRetval(0) = MCIERR_MISSING_COMMAND_STRING
        Exit Function
    End If
    '�������� ������ ��������� �� �������� ��������
    strReturn = Space$(intReturnLength)
    '�������� �������� �������
    lRetval(0) = mciSendString(strCommand, strReturn, intReturnLength, hWndCallback)
    
    intPosition = InStr(strReturn, vbNullChar)
    SendCommand = Left$(strReturn, intPosition - 1)
End Function

'��������, ��������� ������� ������ � ������ ������ - ���� ���� ������, ���������� � ��������
'������� ������������ ����� ������� ������/��������, ����������� mciSendString
Public Property Get IsError() As Long
    Dim i As Long
    '�������� ���� �������� ������� �� ������
    For i = LBound(lRetval) To UBound(lRetval)
        '���� ���� - ���������� �������� ������ � �������
        If lRetval(i) <> 0 Then IsError = lRetval(i): Exit Property
    Next i
    IsError = 0
End Property

'�������� ���������� �������� ������ �� � ����
'(��� ����� ������ �� ������� IsError)
Public Property Get ErrorMessage(ByVal lError As Long) As String
    Dim strErrorString As String * 255
    Dim lRetval As Long
    Dim intPosition As Integer
    '����������� �������� ������ �� ����
    lRetval = mciGetErrorString(lError, strErrorString, 255)
    '���� ������ ��������, ���������� ��������...
    If lRetval Then
        intPosition = InStr(strErrorString, vbNullChar)
        ErrorMessage = Left$(strErrorString, intPosition - 1)
    Else
        ErrorMessage = "Unknown error"
    End If
End Property