VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlay 
   Caption         =   "Disco hook by Alex Dragokas - перехватчик музыки для Left4Dead на базе плагина Disco"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   7695
   Begin VB.Frame FraUpdateCheck 
      Caption         =   "Тип подключения к интернет"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Width           =   7335
      Begin VB.TextBox txtUpdateProxyHost 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   40
         Text            =   "127.0.0.1"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtUpdateProxyPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   39
         Text            =   "8080"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkUpdateUseProxyAuth 
         Caption         =   "С авторизацией"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtUpdateProxyLogin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   37
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtUpdateProxyPass 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5640
         PasswordChar    =   "*"
         TabIndex        =   36
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton optProxyIE 
         Caption         =   "Настройки Internet Explorer"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optProxyManual 
         Caption         =   "Прокси"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkSocks4 
         Caption         =   "Socks4"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         TabIndex        =   33
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton OptProxyDirect 
         Caption         =   "Прямое подключение"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblUpdateServer 
         Caption         =   "Сервер"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblUpdatePort 
         Caption         =   "Порт"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblUpdateLogin 
         Caption         =   "Логин"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2520
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblUpdatePass 
         Caption         =   "Пароль"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   41
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.CheckBox chkUpdates 
      Caption         =   "Проверять обновления"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   6240
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.CommandButton cmdPlayerChoose 
      Caption         =   "Выбрать другой..."
      Enabled         =   0   'False
      Height          =   300
      Left            =   5880
      TabIndex        =   26
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox txtPlayer 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   5880
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Статус"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   20
      Top             =   2280
      Width           =   2655
      Begin VB.Label lblStatusHook 
         Caption         =   "остановлен"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblStatusGame 
         Caption         =   "не запущена"
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Захват:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Игра:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Закрывать это окно при завершении игры"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Тип захвата"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   4575
      Begin VB.TextBox txtFreq 
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Text            =   "2"
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton optForceHookGame 
         Caption         =   "когда игра запущена"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   720
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optForceHookAlways 
         Caption         =   "всегда"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "сек."
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Частота захвата:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdSettings 
      Caption         =   "Другие настройки ..."
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Выполнить тест (скачать и проиграть)"
      Height          =   600
      Left            =   3840
      TabIndex        =   14
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkUseDefPlayer 
      Caption         =   "Играть музыку в плеере, установленном в системе:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Поставьте галочку, если хотите запускать в AIMP или другом внешнем плеере"
      Top             =   5520
      Width           =   5295
   End
   Begin VB.CheckBox chkRepeat 
      Caption         =   "Повтор"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      ToolTipText     =   "Повторять трек по кругу"
      Top             =   960
      Width           =   975
   End
   Begin VB.Timer tmrStatus 
      Interval        =   500
      Left            =   6000
      Top             =   120
   End
   Begin VB.Timer tmrTime 
      Interval        =   500
      Left            =   6480
      Top             =   120
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   0
      ToolTipText     =   "Volume Left"
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   10
      Max             =   1000
      SelStart        =   900
      TickStyle       =   3
      Value           =   900
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "&Stop   X"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pa&use   | |"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play   >"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin MSComctlLib.Slider sldSeek 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Seek"
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Timer tmpStartup 
      Interval        =   50
      Left            =   6960
      Top             =   120
   End
   Begin VB.Timer tmrHook 
      Left            =   7440
      Top             =   120
   End
   Begin VB.Label lblTrack 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1080
      TabIndex        =   13
      Top             =   150
      Width           =   6315
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Файл:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   12
      Top             =   480
      Width           =   510
   End
   Begin VB.Label lblVolume 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Громкость"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4920
      TabIndex        =   11
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1080
      TabIndex        =   10
      Top             =   495
      Width           =   6315
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Общая длительность трека"
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblLabel6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Трек:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TRACK_URL = "https://dragokas.com/music/track.txt"

Private Const MCI_ALIAS = "DiscoHook_Alias_1"
Private Const APP_SECTION = "main"

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Const INFINITE                  As Long = -1

Private m_CurFileName As String
Private m_BeenDownloading As Boolean
Private m_BeenLoadCfg As Boolean
Private mEngine As New clsMedia
Private bPause As Boolean
Private m_PlayerPath As String
Private oShellApp As Object
Private m_UnloadEvent As Boolean
Private sGameExe As String
Private USE_TEST_MODE As Boolean

Private Sub CmdSettings_Click()
    Me.Height = IIf(Me.Height = 2715, 7500, 2715)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    If g_InstallState Then
        Me.Caption = "Настройки и тестирование"
        Me.Top = frmStartup.Top + 500
        Me.Left = frmStartup.Left + 500
    Else
        Me.Height = 2715
        CenterForm Me
        cmdTest.Visible = False
    End If
    
    LoadSettings
    'Set mEngine = New clsMedia
    m_PlayerPath = GetPlayerPath2()
    txtPlayer.Text = m_PlayerPath
    Set oShellApp = CreateObject("Shell.Application")
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error in Form_Load. Err. # " & Err.Number & ": " & Err.Description
End Sub

Private Sub optForceHookAlways_Click()
    SetForceHook True
End Sub

Private Sub optForceHookGame_Click()
    SetForceHook False
End Sub

Private Sub SetForceHook(bForce As Boolean)
    AppSaveSetting APP_SECTION, "hook_force", Abs(bForce)
    
    If bForce Then
        Hook
    Else
        If Not Proc.IsRunned("left4dead.exe") Then
            lblStatusGame.Caption = "не запущена"
            Unhook
            WaitForRun
        Else
            lblStatusGame.Caption = "запущена"
            tmpStartup_Timer 'go in wait loop
        End If
    End If
End Sub

Private Sub txtFreq_Change()
    AppSaveSetting APP_SECTION, "hook_interval", txtFreq.Text
End Sub

Private Sub chkRepeat_Click()
    AppSaveSetting APP_SECTION, "repeat", chkRepeat.Value
End Sub

Private Sub chkUseDefPlayer_Click()
    AppSaveSetting APP_SECTION, "use_default", chkUseDefPlayer.Value
    cmdPlayerChoose.Enabled = chkUseDefPlayer.Value
    txtPlayer.Enabled = chkUseDefPlayer.Value
End Sub

Private Sub cmdTest_Click() 'тестовое скачивание и проигрывание трека
    Unhook
    USE_TEST_MODE = True
    
    m_CurFileName = Env.TempCU & "\track_0.mp3"
    
    If DownloadFile("http://dragokas.com/music/music/dnepropetrovskaja%5Fbratva.mp3", m_CurFileName, True) Then
        lblTrack.Caption = "Днепропетровская братва"
        lblFile.Caption = "dnepropetrovskaja_bratva.mp3"
        
        If chkUseDefPlayer.Value = vbChecked Then
            Call ShellExecute(0&, StrPtr(""), StrPtr(m_CurFileName), 0&, 0&, vbMinimizedNoFocus)
        Else
            cmdPlay_Click
        End If
    End If
End Sub

Private Sub LoadSettings()
    On Error GoTo ErrorHandler
    
    Dim sTmp As String
    
    sTmp = AppGetSetting(APP_SECTION, "l4d_exe")
    sGameExe = IIf(sTmp <> "", sTmp, "left4dead.exe")
    
    sTmp = AppGetSetting(APP_SECTION, "hook_interval")
    txtFreq.Text = IIf(sTmp <> "" And IsNumeric(sTmp), sTmp, 2)
    
    sTmp = AppGetSetting(APP_SECTION, "repeat")
    chkRepeat.Value = IIf(sTmp <> "", sTmp, vbUnchecked)
    
    sTmp = AppGetSetting(APP_SECTION, "use_default")
    chkUseDefPlayer.Value = IIf(sTmp <> "", sTmp, vbUnchecked)
    
    sTmp = AppGetSetting(APP_SECTION, "hook_force")
    optForceHookAlways.Value = IIf(sTmp <> "", True, False)
    
    sTmp = AppGetSetting(APP_SECTION, "volume")
    sldVolume(0).Value = IIf(sTmp <> "", sTmp, 900)

    Exit Sub
ErrorHandler:
    Debug.Print "Error in Form_Load. Err. # " & Err.Number & ": " & Err.Description
    If inIDE Then Stop: Resume Next
End Sub

Private Sub tmpStartup_Timer()
    Dim bRunSuccess As Boolean
    Dim PID As Long
    Dim hProc As Long

    tmpStartup.Interval = 0
    
    If g_InstallState Then Exit Sub
    
    If Proc.IsRunned("left4dead.exe", PID) Then
        bRunSuccess = True
        Proc.SetProcessID = PID
    Else
        If Not g_Separate Then
            If Proc.ProcessRun(BuildPath(App.Path, sGameExe), Command()) Then
                bRunSuccess = True
            End If
        End If
    End If
    
    If USE_DEBUG Or optForceHookAlways.Value Then
        Hook
        Exit Sub
    End If
    
    If bRunSuccess Then
        lblStatusGame.Caption = "запущена"
        Hook
        Call Proc.WaitForTerminate(, Proc.GetProcessHandle, False, INFINITE)
        lblStatusGame.Caption = "не запущена"
        
        If optForceHookAlways.Value = 0 Then
            Unload Me
        End If
    Else
        'Unload Me
        WaitForRun
        If Proc.IsRunned("left4dead.exe") Then
            lblStatusGame.Caption = "запущена"
            tmpStartup_Timer
        End If
    End If
End Sub

Sub WaitForRun()
    Do
        If m_UnloadEvent Then Exit Do
        If optForceHookAlways.Value Then Exit Do
        If Proc.IsRunned("left4dead.exe") Then Exit Do
        DoEvents
        Sleep 500
    Loop
End Sub

Sub Unhook()
    tmrHook.Interval = 0
    lblStatusHook.Caption = "остановлен"
End Sub

Sub Hook()
    On Error GoTo ErrorHandler
    If IsNumeric(txtFreq.Text) Then
        tmrHook.Interval = CLng(txtFreq.Text) * 1000&
        lblStatusHook.Caption = "работает"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Указано некорректное число!"
End Sub

Private Sub tmrHook_Timer()
    On Error GoTo ErrorHandler

    Static bSecond As Boolean
    Static sLastStr As String
    Static TrackIdx As Long
    
    Dim sCurStr As String
    Dim sTrack As String
    Dim pos As Long
    Dim lret As Long
    
    If m_BeenDownloading Or m_BeenLoadCfg Then Exit Sub
    
    If Not USE_DEBUG Then
        If Not bSecond Then
            bSecond = True
            sLastStr = GetTrackURL()
            Exit Sub
        End If
    End If
    
    Debug.Print "Timer"
    
    sCurStr = GetTrackURL()
    
    If sCurStr = sLastStr Then Exit Sub
    
    If Len(sCurStr) <> 0 Then
        sLastStr = sCurStr
    
        Debug.Print sCurStr
        pos = InStr(sCurStr, "*")
        If pos <> 0 Then
            sTrack = Mid$(sCurStr, pos + 1)
            
            'check extension type for security reason
            
            If UCase(GetExtensionName(sTrack)) <> ".MP3" Then
                Debug.Print "Entension name is wrong: " & sTrack
                Exit Sub
            End If
            
            'to prevent access denied when file are being still played
            If TrackIdx = 0 Then
                TrackIdx = 1
            Else
                TrackIdx = 0
            End If
            
            m_CurFileName = Env.TempCU & "\track_" & TrackIdx & ".mp3"
            
            If FileExists(m_CurFileName) Then Kill m_CurFileName
            
            If StrComp(sTrack, "Dummy.mp3", vbTextCompare) = 0 Then
                lblTrack.Caption = ""
                lblFile.Caption = ""
            
                If chkUseDefPlayer.Value = vbChecked Then
                    If m_PlayerPath <> "" Then Proc.ProcessClose , GetFileName(m_PlayerPath) & ".exe"
                Else
                    cmdStop_Click
                End If
                Exit Sub
            End If
            
            m_BeenDownloading = True
            Debug.Print "Downloading ... " & sTrack
            
            lblTrack.Caption = "Скачиваю ..."
            lblFile.Caption = sTrack
            
            If DownloadFile("http://dragokas.com/music/music/" & sTrack, m_CurFileName, True) Then
                
                m_BeenDownloading = False
                lblTrack.Caption = GetSoundTitle(m_CurFileName)
                
                If chkUseDefPlayer.Value = vbChecked Then
                    lret = (32 < ShellExecute(0&, StrPtr(""), StrPtr(m_CurFileName), 0&, 0&, vbMinimizedNoFocus))
                Else
                    cmdPlay_Click
                End If
                
                Debug.Print "PlayMusic ret = " & lret
                
            End If
            m_BeenDownloading = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "Error in Timer1_Timer. Err. # " & Err.Number & ": " & Err.Description
    m_BeenDownloading = False
End Sub

Function GetSoundTitle(sFile As String) As String
    On Error GoTo ErrorHandler:

    Dim oDir As Object:     Set oDir = oShellApp.Namespace(GetParentDir(sFile))
    Dim oFile As Object:    Set oFile = oDir.ParseName(GetFileNameAndExt(sFile))
    
    Dim Artist As String
    Dim Title As String
    Artist = Trim(oDir.GetDetailsOf(oFile, 20))
    Title = Trim(oDir.GetDetailsOf(oFile, 21))
    
    If Artist = "artist" Then Artist = ""
    If Title = "title" Then Title = ""
    
    GetSoundTitle = Artist & IIf(Title <> "", " - " & Title, "")
    Exit Function
ErrorHandler:
    Debug.Print "Error in GetSoundTitle. Err. # " & Err.Number & ": " & Err.Description
End Function

Function GetTrackURL() As String
    m_BeenLoadCfg = True
    GetTrackURL = GetUrl(TRACK_URL)
    m_BeenLoadCfg = False
End Function

'Private Sub chkChannel_Click(Index As Integer)
'    With mEngine
'        If Index = 0 Then
'            .SetMute mciLeftChannel, CBool(chkChannel(Index).Value), _
'                MCI_ALIAS
'        Else
'            .SetMute mciRightChannel, CBool(chkChannel(Index).Value), _
'                MCI_ALIAS
'        End If
'        If .IsError Then MsgBox "Ошибка при вкл./откл. каналов: " & .IsError, _
'            vbCritical Or vbApplicationModal, "Debug"
'    End With
'End Sub

Private Sub cmdPause_Click()
    With mEngine
        If Not bPause Then .mPause (MCI_ALIAS) Else .mResume (MCI_ALIAS)
'        If .IsError Then MsgBox "Ошибка при паузе: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
    End With
    bPause = Not bPause
End Sub

Private Sub cmdPlay_Click()
    If m_BeenDownloading Then Exit Sub
    If m_CurFileName = "" Then Exit Sub
    cmdStop_Click
    With mEngine
        'If chkVideo.Value Then .VideoWnd = picVideo Else .VideoWnd = Nothing
        .FileName = m_CurFileName
        .Fullscreen = False ' CBool(chkFullscreen.Value)
        .Wait = False 'CBool(chkWait.Value)
        .Repeat = CBool(chkRepeat.Value)
        .DeviceType = "MPEGVideo" 'txtDeviceType.Text
        '.Shareable = True
        '.Notify = True
        '.WndCallback = frmMain.hWnd
        .mOpen MCI_ALIAS

        If .IsError Then MsgBox "Ошибка при открытии: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
        
        .mPlay MCI_ALIAS
        
        If .Length(MCI_ALIAS) Then sldSeek.Max = .Length(MCI_ALIAS)
        If .IsError Then MsgBox "Ошибка при воспроизведении: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
    End With
    bPause = False
End Sub

Private Sub cmdStop_Click()
    With mEngine
        .mStop MCI_ALIAS
'        If .IsError Then MsgBox "Ошибка при остановке: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
        .mClose MCI_ALIAS
'        If .IsError Then MsgBox "Ошибка при закрытии: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
    End With
    bPause = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    m_UnloadEvent = True
    If Not (mEngine Is Nothing) Then
        With mEngine
            If .FileName <> vbNullString Then
                .mCloseAll
'            If .IsError Then MsgBox "Ошибка при закрытии всего: " & .IsError, _
                vbCritical Or vbApplicationModal, "Debug"
            End If
        End With
        Set mEngine = Nothing
    End If
End Sub

Private Sub sldSeek_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrStatus.Enabled = False
End Sub

Private Sub sldSeek_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With mEngine
        .mSeek sldSeek.Value, MCI_ALIAS
'        If .IsError Then MsgBox "Ошибка при изменении позиции: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
        .mPlay MCI_ALIAS
'        If .IsError Then MsgBox "Ошибка при воспроизведении: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
    End With
    tmrStatus.Enabled = True
End Sub

Private Sub sldVolume_Change(Index As Integer)
    With mEngine
        Call .SetVolume(sldVolume(0).Value, sldVolume(0).Value, MCI_ALIAS)
'        If .IsError Then MsgBox "Ошибка при изменении громкости: " & .IsError, _
            vbCritical Or vbApplicationModal, "Debug"
    End With
    AppSaveSetting APP_SECTION, "volume", sldVolume(0).Value
End Sub

Private Sub sldVolume_Scroll(Index As Integer)
    Call sldVolume_Change(Index)
End Sub

Private Sub tmrStatus_Timer()
    sldSeek.Value = mEngine.Position(MCI_ALIAS)
End Sub

Private Sub tmrTime_Timer()
    Dim intTime(0 To 1) As Integer

    intTime(0) = mEngine.Position(MCI_ALIAS) / 1000
    intTime(1) = mEngine.Length(MCI_ALIAS) / 1000
    
    lblTime(0).Caption = Format$(intTime(0) \ 3600, "00") & _
        ":" & Format$((intTime(0) \ 60) Mod 60, "00") & ":" & _
        Format$(intTime(0) Mod 60, "00")
    lblTime(1).Caption = Format$(intTime(1) \ 3600, "00") & _
        ":" & Format$((intTime(1) \ 60) Mod 60, "00") & ":" & _
        Format$(intTime(1) Mod 60, "00")
End Sub
