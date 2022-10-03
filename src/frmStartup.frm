VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "DiscoMusic Hook Installer by Alex Dragokas"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Прочие настройки"
      Height          =   345
      Left            =   5520
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Начать установку"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdSelectDir 
      Caption         =   "Обзор ..."
      Height          =   360
      Left            =   6120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtGameFolder 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Text            =   "...\Steam\steamapps\common\left 4 dead\left4dead.exe"
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Создано эксклюзивно для сервера Spartan Witch - 46.174.49.143:27000"
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ver. 1.0 beta"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   1320
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Укажите запускаемый файл Left4Dead"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Установщик перехватчика музыки для плагина Disco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   5400
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const USE_DEBUG = False
Const APP_VERSION = "1.2"

Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Const APP_SECTION = "main"

Private sGamePath As String
Private sGameExe As String

Private Sub Form_Load()
    Dim bForceStart As Boolean
    
    g_InstallState = True
    
    #If USE_DEBUG Then
        bForceStart = True
        USE_DEBUG = True
    #End If
    
    lblVersion.Caption = "ver." & APP_VERSION

    InitVariables
    
    'already installed?
    If bForceStart Or isL4dDir(App.Path) Then
        g_InstallState = False
        If Len(Command$()) = 0 Then g_Separate = True
        frmPlay.Show
        Unload Me
    Else
        Me.WindowState = vbNormal
        sGamePath = FindPathL4d()
        If sGamePath <> "" Then
            txtGameFolder.Text = sGamePath & "\left4dead.exe"
        End If
    End If
End Sub

Private Sub txtGameFolder_Change()
    If isL4dDir(GetParentDir(txtGameFolder.Text)) Then
        sGamePath = GetParentDir(txtGameFolder.Text)
        sGameExe = GetFileNameAndExt(txtGameFolder.Text)
        cmdInstall.Enabled = True
        txtGameFolder.ForeColor = vbBlack
    Else
        txtGameFolder.ForeColor = vbRed
    End If
End Sub

Private Sub cmdInstall_Click()
    
    Dim sExeName$: sExeName = "Disco-Music-Hook"
    If sGameExe = "" Then sGameExe = "left4dead.exe"
    Shell "taskkill.exe /f /im DiscoHook.exe"
    
    FileCopy BuildPath(App.Path, IIf(inIDE, sExeName, App.EXEName) & ".exe"), BuildPath(sGamePath, "DiscoHook.exe")
    
    With CreateObject("WScript.Shell")
      With .CreateShortcut(BuildPath(Env.Desktop, "Left4dead + music.lnk"))
        .WorkingDirectory = sGamePath
        .Description = "Запуск Left4dead + DiscoMusicHook"
        .TargetPath = BuildPath(sGamePath, "DiscoHook.exe")
        .Arguments = "-novid -language russian +connect 46.174.48.12:27243"
        .IconLocation = BuildPath(sGamePath, sGameExe)
        .Save
      End With
      
      With .CreateShortcut(BuildPath(Env.Desktop, "Disco Music.lnk"))
        .WorkingDirectory = sGamePath
        .Description = "Disco Music"
        .TargetPath = BuildPath(sGamePath, "DiscoHook.exe")
        .Arguments = ""
        .IconLocation = BuildPath(sGamePath, "DiscoHook.exe")
        .Save
      End With
    End With
    
    UnpackResource 101, BuildPath(sGamePath, "MSCOMCTL.OCX")
    
    AppSaveSetting APP_SECTION, "l4d_exe", sGameExe
    FileCopy BuildPath(App.Path, SETTINGS_FILENAME), BuildPath(sGamePath, SETTINGS_FILENAME)
    
    If vbYes = MsgBox("Установка завершена." & _
        vbCrLf & vbCrLf & "На рабочем столе создано 2 ярлыка:" & vbCrLf & _
        "'Lef4dead + music' - чтобы запустить и музыку, и игру" & _
        vbCrLf & "'Disco Music' - чтобы запустить только музыку." & vbCrLf & vbCrLf & _
        "Файл DiscoMusicHook.exe больше не нужен." & vbCrLf & "Удалить его сейчас?", vbYesNo Or vbInformation) Then
        
        Dim f$
        f = """" & BuildPath(App.Path, IIf(inIDE, sExeName, App.EXEName) & ".exe") & """"
        Shell "cmd /v/c (set f=" & f & "&for /l %l in () do if exist !f! (del /f/a !f!) else (exit))", 0
        
        If FileExists(BuildPath(App.Path, "MSCOMCTL.OCX")) Then
            f = """" & BuildPath(App.Path, "MSCOMCTL.OCX") & """"
            Shell "cmd /v/c (set f=" & f & "&for /l %l in () do if exist !f! (del /f/a !f!) else (exit))", 0
        End If
    End If
    
    Kill BuildPath(App.Path, SETTINGS_FILENAME)
    
    Unload Me
End Sub

Private Sub cmdSelectDir_Click()
    Dim sPath As String
    
    If OSver.IsWindowsVistaOrGreater Then
        sPath = GetOpenFile2(Me.hwnd, txtGameFolder.Text)
    Else
        sPath = GetOpenFile(Me.hwnd, txtGameFolder.Text)
    End If
    If 0 <> Len(sPath) Then
        If isL4dDir(GetParentDir(sPath)) Then
            txtGameFolder.Text = sPath
            sGamePath = GetParentDir(sPath)
            sGameExe = GetFileNameAndExt(sPath)
            cmdInstall.Enabled = True
            txtGameFolder.ForeColor = vbBlack
        Else
            MsgBox "Указанная папка не является папкой игры!"
        End If
    End If
End Sub

Private Sub cmdSettings_Click()
    UnpackResource 101, BuildPath(App.Path, "MSCOMCTL.OCX")
    frmPlay.Show vbModal
End Sub

Private Sub Form_Initialize()

    On Error Resume Next
    Dim ICC         As tagINITCOMMONCONTROLSEX
    
    Debug.Assert MakeTrue(inIDE)
    
    With ICC
        .dwSize = Len(ICC)
        .dwICC = &HFF&
    End With

    InitCommonControlsEx ICC

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Forms.Count = 1 Then
        ReleaseVariables
    End If
End Sub

Private Function MakeTrue(ByRef bvar As Boolean) As Boolean: bvar = True: MakeTrue = True: End Function
