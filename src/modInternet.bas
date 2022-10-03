Attribute VB_Name = "modInternet"
'
' Internet module by Merijn Bellekom & Alex Dragokas
'

Option Explicit

Private Const MAX_HOSTNAME_LEN = 132&
Private Const MAX_DOMAIN_NAME_LEN = 132&
Private Const MAX_SCOPE_ID_LEN = 260&

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type IP_ADDR_STRING
    Next As Long
    IpAddress As String * 16
    IpMask As String * 16
    Context As Long
End Type

Private Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId  As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type
Public Enum COMPUTER_NAME_FORMAT
  ComputerNameNetBIOS
  ComputerNameDnsHostname
  ComputerNameDnsDomain
  ComputerNameDnsFullyQualified
  ComputerNamePhysicalNetBIOS
  ComputerNamePhysicalDnsHostname
  ComputerNamePhysicalDnsDomain
  ComputerNamePhysicalDnsFullyQualified
  ComputerNameMax
End Enum

Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectW" (ByVal InternetSession As Long, ByVal sServerName As Long, ByVal nServerPort As Integer, ByVal sUsername As Long, ByVal sPassword As Long, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenW" (ByVal sAgent As Long, ByVal lAccessType As Long, ByVal sProxyName As Long, ByVal sProxyBypass As Long, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlW" (ByVal hInternetSession As Long, ByVal sURL As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, Buffer As Any, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Public Declare Function InternetReadFileString Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByVal Buffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestW" (ByVal hHttpSession As Long, ByVal sVerb As Long, ByVal sObjectName As Long, ByVal sVersion As Long, ByVal sReferer As Long, lplpszAcceptTypes As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestW" (ByVal hHttpRequest As Long, ByVal sHeaders As Long, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Long
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoW" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Declare Function GetNetworkParams Lib "IPHlpApi.dll" (FixedInfo As Any, pOutBufLen As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)


Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_OVERWRITEPROMPT = &H2

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Const INTERNET_SERVICE_HTTP = 3
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

Private Const ERROR_BUFFER_OVERFLOW = 111&

Public szResponse As String
Public szSubmitUrl As String


Public Sub SendData(szUrl As String, szData As String)
    On Error GoTo ErrorHandler
    Dim szRequest As String
    Dim xmlhttp As Object
    Dim dataLen As Long
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    szRequest = "data=" & URLEncode(szData)

    dataLen = Len(szRequest)
    xmlhttp.Open "POST", szUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion

    xmlhttp.send "" & szRequest
    szResponse = xmlhttp.responseText

    Set xmlhttp = Nothing
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "SendData"
    If inIDE Then Stop: Resume Next
End Sub

Public Function GetUrl(szUrl As String) As String
    On Error GoTo ErrorHandler:
    Dim szRequest As String
    Dim xmlhttp As Object
    Dim dataLen As Long
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")

    dataLen = Len(szRequest)
    xmlhttp.Open "GET", szUrl, False
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    'xmlhttp.setRequestHeader "User-Agent", "HJT.1.99.2" & "|" & sWinVersion & "|" & sMSIEVersion
    
    xmlhttp.send "" & szRequest

    GetUrl = xmlhttp.responseText

    Set xmlhttp = Nothing
    Exit Function

ErrorHandler:

End Function

Function URLEncode(ByVal Text As String) As String
    On Error GoTo ErrorHandler:

    Dim i As Long
    Dim acode As Long
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "URLEncode", "Src:", Text
    If inIDE Then Stop: Resume Next
End Function

Public Function IsOnline() As Boolean

   IsOnline = InternetGetConnectedState(0&, 0&)
     
End Function

' ---------------------------------------------------------------------------------------------------
' StartupList2 routine
' ---------------------------------------------------------------------------------------------------

Public Function DownloadFile(sURL$, sTarget$, Optional bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Const Chunk As Long = 16384

    Dim hInternet&, hFile&, sFile$, lBytesRead&
    Dim sUserAgent$, ff%
    Dim aBuf() As Byte, curPos As Long
    
    sUserAgent = "StartupList v" & "1.0"
    
    hInternet = InternetOpen(StrPtr(sUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
    
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, StrPtr(sURL), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        
        If hFile = 0 And InStr(1, sURL, "https://", 1) <> 0 And OSver.MajorMinor <= 5.2 Then 'XP + https ?
            hFile = InternetOpenUrl(hInternet, StrPtr(Replace$(sURL, "https://", "http://", , , 1)), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        End If
        
        If hFile <> 0 Then
            curPos = -1
            DownloadFile = True
            Do
                ReDim Preserve aBuf(curPos + Chunk)
                InternetReadFile hFile, aBuf(curPos + 1), Chunk, lBytesRead
                If lBytesRead < Chunk Then
                    If curPos + lBytesRead <> -1 Then
                        ReDim Preserve aBuf(curPos + lBytesRead)
                        DownloadFile = True
                    End If
                    Exit Do
                Else
                    curPos = curPos + Chunk
                End If
            Loop Until lBytesRead = 0
            
            InternetCloseHandle hFile
            
            If DownloadFile Then
                ff = FreeFile()
                If FileExists(sTarget) Then Kill sTarget
                Open sTarget For Binary Access Write As #ff
                    Put #ff, , aBuf
                Close #ff
            End If
        Else
            If Not bSilent Then
                'Unable to connect to the Internet.
                MsgBox "Unable to connect to the Internet.", vbCritical
            End If
        End If
        InternetCloseHandle hInternet
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadFile", "URL:", sURL, "Target:", sTarget
    DownloadFile = False
    If inIDE Then Stop: Resume Next
End Function

Public Function DownloadFileToArray(sURL$, aBuf() As Byte, Optional bSilent As Boolean) As Boolean
    On Error GoTo ErrorHandler:

    Const Chunk As Long = 16384

    Dim hInternet&, hFile&, sFile$, lBytesRead&
    Dim sUserAgent$, ff%
    Dim curPos As Long
    
    sUserAgent = "StartupList v" & "1.0"
    
    hInternet = InternetOpen(StrPtr(sUserAgent), INTERNET_OPEN_TYPE_DIRECT, 0&, 0&, 0&)
    
    If hInternet Then
        hFile = InternetOpenUrl(hInternet, StrPtr(sURL), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        
        If hFile = 0 And InStr(1, sURL, "https://", 1) <> 0 And OSver.MajorMinor <= 5.2 Then 'XP + https ?
            hFile = InternetOpenUrl(hInternet, StrPtr(Replace$(sURL, "https://", "http://", , , 1)), 0&, 0&, INTERNET_FLAG_RELOAD, 0&)
        End If
        
        If hFile <> 0 Then
            curPos = -1
            Do
                ReDim Preserve aBuf(curPos + Chunk)
                InternetReadFile hFile, aBuf(curPos + 1), Chunk, lBytesRead
                If lBytesRead < Chunk Then
                    If curPos + lBytesRead <> -1 Then
                        ReDim Preserve aBuf(curPos + lBytesRead)
                        DownloadFileToArray = True
                    Else
                        Erase aBuf
                    End If
                    Exit Do
                Else
                    curPos = curPos + Chunk
                End If
            Loop Until lBytesRead = 0
            
            InternetCloseHandle hFile
        Else
            If Not bSilent Then
                'Unable to connect to the Internet.
                MsgBox "Unable to connect to the Internet.", vbCritical
            End If
        End If
        InternetCloseHandle hInternet
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "DownloadFile", "URL:", sURL
    DownloadFileToArray = False
    If inIDE Then Stop: Resume Next
End Function

Public Function OpenURL(sURL As String) As Boolean
    OpenURL = (32 < ShellExecute(0&, StrPtr("open"), StrPtr(sURL), 0&, 0&, vbNormalFocus))
End Function

