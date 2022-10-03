Attribute VB_Name = "modWriteLog"
Option Explicit

Public Const SETTINGS_FILENAME = "DiscoHook.ini"

Private Const MAX_PATH As Long = 260&

Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As Long, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ErrMessageText(lCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000&
    Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
    
    Dim sRtrnMessage   As String
    Dim lret           As Long
    
    sRtrnMessage = String$(MAX_PATH, vbNullChar)
    lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lCode, 0&, StrPtr(sRtrnMessage), MAX_PATH, ByVal 0&)
    If lret > 0 Then
        ErrMessageText = Left$(sRtrnMessage, lret)
        ErrMessageText = Replace$(ErrMessageText, vbCrLf, vbNullString)
    End If
End Function

Public Function ParseDateTime(myDate As Date) As String
    ParseDateTime = Right$("0" & Day(myDate), 2) & _
        "." & Right$("0" & Month(myDate), 2) & _
        "." & Year(myDate) & _
        " " & Right$("0" & Hour(myDate), 2) & _
        ":" & Right$("0" & Minute(myDate), 2) & _
        ":" & Right$("0" & Second(myDate), 2)
End Function

Public Sub ErrorMsg(ByVal ErrObj As ErrObject, sProcedure As String, ParamArray CodeModule())
    Dim HRESULT     As String
    Dim Other       As String
    Dim i           As Long
    Dim sFormatted  As String
    
    For i = 0 To UBound(CodeModule)
        Other = Other & CodeModule(i) & " "
    Next
    
    HRESULT = ErrMessageText(IIf(ErrObj.Number = 0, ErrObj.LastDllError, ErrObj.Number))
    
    sFormatted = _
        "- " & ParseDateTime(Now) & _
        " - " & sProcedure & _
        " - #" & ErrObj.Number & " " & _
        ErrObj.Description & _
        ". LastDllError = " & ErrObj.LastDllError & _
        IIf(Len(HRESULT), " (" & HRESULT & ")", "") & " " & _
        IIf(Len(Other), "" & Other, "")
    
    Debug.Print sFormatted
End Sub

Public Function AppGetSetting(Section As String, Parameter As String, Optional Default As Variant = vbNullString) As Variant
    Dim lr As Long
    Dim buf As String
    Dim sIniFile As String
    sIniFile = BuildPath(App.Path, SETTINGS_FILENAME)
    buf = Space(255)
    lr = GetPrivateProfileString(Section, Parameter, "error", buf, Len(buf), sIniFile)
    buf = Left(buf, lr)
    If buf <> "error" Then AppGetSetting = buf Else AppGetSetting = Default
End Function

Public Sub AppSaveSetting(Section As String, Parameter As String, Value As Variant)
    Dim lr As Long
    Dim sIniFile As String
    sIniFile = BuildPath(App.Path, SETTINGS_FILENAME)
    lr = WritePrivateProfileString(Section, Parameter, CStr(Value), sIniFile)
    If lr = 0 Then MsgBox "Ошибка записи в секцию " & Parameter & " файла настроек " & sIniFile & ". Код: " & Err.LastDllError
End Sub
