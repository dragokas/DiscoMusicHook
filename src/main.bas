Attribute VB_Name = "modMain"
Option Explicit

Private Type MOD_NAMES
    L4d2_on_L4d1 As String
    L4d2_on_L4d1_Archive As String
End Type

Private Type STRING_CONSTANTS
    Mods As MOD_NAMES
End Type

Private Type NICK_INFO
    RevIni As String
    Name As String
    UTF8 As Boolean
End Type

Public Type GAME_INFO
    Nick As NICK_INFO
    L4dPath As String
    IsSteam As String
End Type

Public Enum GAME_MODE
    MODE_COOP
    MODE_VERSUS
    MODE_SURVIVAL
End Enum

Public BWitchMOD_Name As String

Public colDifficulty As New Collection
Public colGameType As New Collection
Public colCampany As New Collection
Public colPerson As New Collection

Public Game     As GAME_INFO
Public strConst As STRING_CONSTANTS

Public Sub GetNickName()
    On Error GoTo ErrorHandler

    '//TODO: if firstRun
    
    If True Then
        
        
    Else
        'get from our settings
    
    End If

    Dim aBuf() As Byte
    Dim cpPercent As Long
    Dim CodePage As Long
    Dim ff As Long
    Dim buf As String
    Dim aLines() As String
    Dim sTemp As String
    Dim bUTF8 As Long
    
    Game.Nick.RevIni = Game.L4dPath & "\" & "rev.ini"
    
    If FileExists(Game.Nick.RevIni) Then
    
        ff = FreeFile()
        
        buf = ReadIniValue(Game.Nick.RevIni, "steamclient", "PlayerName")
           
        aBuf() = StrConv(buf, vbFromUnicode, &H419&)
        CodePage = GetEncoding(aBuf, cpPercent)
        
        If (UTF8 = CodePage) And (cpPercent = -1 Or cpPercent > 50) Then bUTF8 = True
        
        If Not bUTF8 Then
            ' try another one check by BOM
            
            Open Game.Nick.RevIni For Input As #ff
                If LOF(ff) >= 3 Then
                    sTemp = String$(3, 0)
                    Input #ff, sTemp
                End If
            Close #ff
        
            If sTemp = Chr$(&HEF) & Chr$(&HBB) & Chr$(&HBF) Then bUTF8 = True
        End If
        
        If bUTF8 Then
            Game.Nick.Name = ConvertCodePageW(buf, UTF8)
            Game.Nick.UTF8 = True
        Else
            Game.Nick.Name = buf
            Game.Nick.UTF8 = False
        End If
        
        Game.IsSteam = False
    Else
        'Steam version
        
        Game.IsSteam = True
    End If
    
    Form1.txtNick.Text = Game.Nick.Name
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "GetNickName"
End Sub

Public Sub SaveNickName(ByVal sNick As String)

    If Game.IsSteam Then
        
    Else
        If Game.Nick.UTF8 Then
            sNick = ConvertCodePage(sNick, WIN, UTF8)
        End If
    
        WriteIniValue Game.Nick.RevIni, "steamclient", "PlayerName", """" & sNick & """"
    End If

End Sub

Public Sub UpdateConnectIP()
    Static ExternalIP As String
    
    If 0 = Len(ExternalIP) Then
        ExternalIP = GetExternalIp()
    End If

    Form1.txtMyIP.Text = "connect " & ExternalIP & ":" & Form1.txtPort.Text
End Sub

'example:

'Survivor Bill
'Survivor Francis
'Survivor Louis
'Survivor Zoey
'Infected

'hl2.exe" -game left4dead +z_difficulty Normal +maxplayers 4 +team_desired "Survivor Bill" +map l4d_hospital01_apartment -console

'$command = " -game left4dead +z_difficulty " & $difficult_var & " +maxplayers " & $maxplayers & ' +team_desired "' & $player & '" +map ' & $map & $console

'$command = ' -game left4dead +team_desired "' & $player & '" +connect ' & $server & $console

'
'
'
''

Public Sub FillInfo()

    Dim it, i As Long
    
    With colDifficulty
        .Add "Легко", "Easy"
        .Add "Нормально", "Normal"
        .Add "Мастер", "Hard"
        .Add "Эксперт", "Impossible"
    End With
    
    For Each it In colDifficulty
        Form1.CbComplex.AddItem it
    Next
    
    With colGameType
        .Add "Кооперация"
        .Add "Сражение"
        .Add "Выживание"
    End With
    
    For Each it In colGameType
        Form1.CbType.AddItem it
    Next
    
    With colCampany
        .Add "НЕТ МИЛОСЕРДИЮ", "hospital"
        .Add "РОКОВОЙ ПОЛЁТ", "garage"
        .Add "ПОХОРОННЫЙ ЗВОН", "smalltown"
        .Add "СМЕРТЬ В ВОЗДУХЕ", "airport"
        .Add "КРОВАВАЯ ЖАТВА", "farm"
        .Add "ЖЕРТВА", "sacrifice" 'note: no .txt info associated
    End With
    
    'survival
    '"ПОСЛЕДНИЙ РУБЕЖ", "lighthouse"
    '"Маяк", "l4d_sv_lighthouse"
    
    For Each it In colCampany
        Form1.CbCampany.AddItem it
    Next
    
    With colPerson
        .Add "Зоя"
        .Add "Луис"
        .Add "Френсис"
        .Add "Билл"
    End With

    For Each it In colPerson
        Form1.CbPlayerName.AddItem it
        Form1.CbPlayerNameRemote.AddItem it
    Next
    
    For i = 1 To 16
        Form1.CbPlayersCnt.AddItem i
    Next
    
    SetDefaultServerSettings
    UpdateConnectIP
    GetNickName
    
    strConst.Mods.L4d2_on_L4d1 = "Глобальный мод L4d2 на L4d1 (Карты, звуки, эффекты)"
    strConst.Mods.L4d2_on_L4d1_Archive = "L4D2_on_L4D1_Mod_(Maps,Sound,Effects).7z"
    
    UpdateModsList
End Sub

Public Sub UpdateModsList()
    lstClearAll Form1.lstMods
    
    With Form1.lstMods
        .AddItem strConst.Mods.L4d2_on_L4d1
        .Selected(0) = True
    End With
    
    If FileExists(BuildPath(AppPath(), strConst.Mods.L4d2_on_L4d1_Archive)) Then
        lstUpdate Form1.lstMods, strConst.Mods.L4d2_on_L4d1, " - Доступен для установки", True
    End If
    
    If FileExists(BuildPath(Game.L4dPath, "left4dead\addons\l4d2_soundscapes.vpk")) Then
        lstUpdate Form1.lstMods, strConst.Mods.L4d2_on_L4d1, " - Установлен", True
    End If
End Sub

Sub SetDefaultServerSettings()
    With Form1
        .CbComplex.Text = "Эксперт"
        .CbType.Text = "Кооперация"
        .CbCampany = "НЕТ МИЛОСЕРДИЮ"
        .CbPlayerName.Text = "Билл"
        .CbPlayerNameRemote.Text = "Билл"
        .CbPlayersCnt.Text = 8
        
    End With
End Sub

Public Function GetConfig()
    Dim sDifficulty As String

    sDifficulty = GetCollectionKeyByName(Form1.CbComplex.Text, colDifficulty)
    
    
    MsgBox sDifficulty

End Function


Public Function GetMapName(sCampany As String, sLevel As String, sMode As GAME_MODE) As String
    Dim sMap As String

    Select Case sCampany
    
      Case "НЕТ МИЛОСЕРДИЮ" 'hospital
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Апартаменты": sMap = "l4d_hospital01_apartment"
        Case "2. Метро": sMap = "l4d_hospital02_subway"             'survival
        Case "3. Канализация": sMap = "l4d_hospital03_sewers"       'survival
        Case "4. Госпиталь": sMap = "l4d_hospital04_interior"       'survival
        Case "5. Крыша": sMap = "l4d_hospital05_rooftop"            'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Апартаменты": sMap = "l4d_vs_hospital01_apartment"
        Case "2. Метро": sMap = "l4d_vs_hospital02_subway"
        Case "3. Канализация": sMap = "l4d_vs_hospital03_sewers"
        Case "4. Госпиталь": sMap = "l4d_vs_hospital04_interior"
        Case "5. Крыша": sMap = "l4d_vs_hospital05_rooftop"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Апартаменты": sMap = ""
        Case "2. Метро": sMap = "l4d_hospital02_subway"
        Case "3. Канализация": sMap = "l4d_hospital03_sewers"
        Case "4. Госпиталь": sMap = "l4d_hospital04_interior"
        Case "5. Крыша": sMap = "l4d_vs_hospital05_rooftop"
        End Select
      End If

      Case "РОКОВОЙ ПОЛЁТ" 'garage
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Переулки": sMap = "l4d_garage01_alleys"            'survival '!!!
        Case "2. Гараж": sMap = "l4d_garage02_lots"                 'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Переулки": sMap = "l4d_garage01_alleys"
        Case "2. Гараж": sMap = "l4d_garage02_lots"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Переулки": sMap = "l4d_garage01_alleys"
        Case "2. Гараж": sMap = "l4d_garage02_lots"
        End Select
      End If

      Case "ПОХОРОННЫЙ ЗВОН" 'smalltown
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Ограждение": sMap = "l4d_smalltown01_caves"
        Case "2. Водосток": sMap = "l4d_smalltown02_drainage"       'survival
        Case "3. Церковь": sMap = "l4d_smalltown03_ranchhouse"      'survival
        Case "4. Город": sMap = "l4d_smalltown04_mainstreet"        'survival
        Case "5. Лодочная станция": sMap = "l4d_smalltown05_houseboat" 'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Ограждение": sMap = "l4d_vs_smalltown01_caves"
        Case "2. Водосток": sMap = "l4d_vs_smalltown02_drainage"
        Case "3. Церковь": sMap = "l4d_vs_smalltown03_ranchhouse"
        Case "4. Город": sMap = "l4d_vs_smalltown04_mainstreet"
        Case "5. Лодочная станция": sMap = "l4d_vs_smalltown05_houseboat"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Ограждение": sMap = ""
        Case "2. Водосток": sMap = "l4d_smalltown02_drainage"
        Case "3. Церковь": sMap = "l4d_smalltown03_ranchhouse"
        Case "4. Город": sMap = "l4d_smalltown04_mainstreet"
        Case "5. Лодочная станция": sMap = "l4d_vs_smalltown05_houseboat"
        End Select
      End If

      Case "СМЕРТЬ В ВОЗДУХЕ" 'airport
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Теплица": sMap = "l4d_airport01_greenhouse"
        Case "2. Кран": sMap = "l4d_airport02_offices"              'survival
        Case "3. Стройка": sMap = "l4d_airport03_garage"            'survival
        Case "4. Терминал": sMap = "l4d_airport04_terminal"         'survival
        Case "5. Взлётная полоса": sMap = "l4d_airport05_runway"    'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Теплица": sMap = "l4d_vs_airport01_greenhouse"
        Case "2. Кран": sMap = "l4d_vs_airport02_offices"
        Case "3. Стройка": sMap = "l4d_vs_airport03_garage"
        Case "4. Терминал": sMap = "l4d_vs_airport04_terminal"
        Case "5. Взлётная полоса": sMap = "l4d_vs_airport05_runway"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Теплица": sMap = ""
        Case "2. Кран": sMap = "l4d_airport02_offices"
        Case "3. Стройка": sMap = "l4d_airport03_garage"
        Case "4. Терминал": sMap = "l4d_airport04_terminal"
        Case "5. Взлётная полоса": sMap = "l4d_vs_airport05_runway"
        End Select
      End If

      Case "КРОВАВАЯ ЖАТВА" 'farm
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Леса": sMap = "l4d_farm01_hilltop"
        Case "2. Тоннель": sMap = "l4d_farm02_traintunnel"          'survival
        Case "3. Мост": sMap = "l4d_farm03_bridge"                  'survival
        Case "4. Ж/Д станция": sMap = "l4d_farm04_barn"
        Case "5. Ферма": sMap = "l4d_farm05_cornfield"              'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Леса": sMap = "l4d_vs_farm01_hilltop"
        Case "2. Тоннель": sMap = "l4d_vs_farm02_traintunnel"
        Case "3. Мост": sMap = "l4d_vs_farm03_bridge"
        Case "4. Ж/Д станция": sMap = "l4d_vs_farm04_barn"
        Case "5. Ферма": sMap = "l4d_vs_farm05_cornfield"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Леса": sMap = ""
        Case "2. Тоннель": sMap = "l4d_farm02_traintunnel"
        Case "3. Мост": sMap = "l4d_farm03_bridge"
        Case "4. Ж/Д станция": sMap = ""
        Case "5. Ферма": sMap = "l4d_vs_farm05_cornfield"
        End Select
      End If
    
      Case "ЖЕРТВА" 'sacrifice
      If sMode = MODE_COOP Then
        Select Case sLevel
        Case "1. Доки": sMap = "l4d_river01_docks"                  'survival
        Case "2. Баржа": sMap = "l4d_river02_barge"
        Case "3. Порт": sMap = "l4d_river03_port"                   'survival
        End Select
      ElseIf sMode = MODE_VERSUS Then
        Select Case sLevel
        Case "1. Доки": sMap = "l4d_river01_docks"
        Case "2. Баржа": sMap = "l4d_river02_barge"
        Case "3. Порт": sMap = "l4d_river03_port"
        End Select
      ElseIf sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "1. Доки": sMap = "l4d_river01_docks"
        Case "2. Баржа": sMap = ""
        Case "3. Порт": sMap = "l4d_river03_port"
        End Select
      End If
    
      Case "ПОСЛЕДНИЙ РУБЕЖ" 'lighthouse
      If sMode = MODE_SURVIVAL Then
        Select Case sLevel
        Case "Маяк": sMap = "l4d_sv_lighthouse"                     'survival only
        End Select
      End If
      
      Case "Спортзал"
        sMap = "tutorial_standards"
    
    End Select
    
    GetMapName = sMap
End Function


