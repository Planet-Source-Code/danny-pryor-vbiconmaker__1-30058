Attribute VB_Name = "ModIcoEd"
Option Explicit



Type BITMAPINFOHEADER
     biSize As Long
     biWidth As Long
     biHeight As Long
     biPlanes As Integer
     biBitCount As Integer
     biCompression As Long
     biSizeImage As Long
     biXPelsPerMeter As Long
     biYPelsPerMeter As Long
     biClrUsed As Long
     biClrImportant As Long
End Type

Type ICONDIR
     idReserved As Integer
     idType As Integer
     idCount As Integer
End Type

Type ICONDIRENTRY
     bWidth As Byte
     bHeight As Byte
     bColorCount As Byte
     bReserved As Byte
     wPlanes As Integer
     wBitCount As Integer
     dwBytesInRes As Long
     dwImageOffset As Long
End Type

Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type

Public BIH As BITMAPINFOHEADER
Public ID As ICONDIR
Public IDE As ICONDIRENTRY

Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, ByVal SecAtts&, phkResult&, lpdwDisp&)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, lpReserved&, lpType&, ByVal lpData$, lpcbData&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpData$, ByVal cbData&)


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_ALL_ACCESS = (&H1F0000 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000)

Public Const Ttl = "vbIconMaker"
Public Sub CheckSettings()
Dim Ret1
    #If Win32 Then
        If Not ScreenDimensionsOk Then End
        If Not ColorModeOk Then End
        'Ret1 = MsgBox("Unless your color settings are changed to 24bit or higher, Icons produced by this program are not viewed properly (Icons will have Gray where transparency should be). Do you want to continue?", vbYesNo, "Improper Settings")
        'If Ret1 = vbNo Then End
        'End If
        Form1.Show
    #Else
        End
    #End If
End Sub
Private Function ScreenDimensionsOk() As Boolean
    Dim Msg As String
    Dim ScreenHeight As Integer
    Dim ScreenWidth As Integer

    ScreenDimensionsOk = True

    ScreenWidth = GetSystemMetrics(0)
    ScreenHeight = GetSystemMetrics(1)

    If ScreenHeight < 480 Or ScreenWidth < 640 Then
       Msg = Ttl & " requires a screen resolution of at least 640 x 480 pixels."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Please adjust your screen resolution and try again."
       MsgBox Msg, vbCritical, Ttl & " - Error"
       ScreenDimensionsOk = False
    End If

End Function
Private Function ColorModeOk() As Boolean

    Dim Msg As String
    Dim hScreenDc As Long

    hScreenDc = GetDC(0)
    
    If GetDeviceCaps(hScreenDc, 12) > 8 Then
       ColorModeOk = True
    Else
       Msg = Ttl & " requires at least high color mode."
       Msg = Msg & vbCrLf & vbCrLf
       Msg = Msg & "Please adjust your color mode and try again."
       MsgBox Msg, vbCritical, Ttl & " - Error"
    End If

End Function
Public Sub SetUpIconDblClick()

    Dim RegData$, hKey&, Rv&

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, ".ico", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "icofile", 7
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "Windows Icon", 12
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile\DefaultIcon", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       RegSetValueEx hKey, vbNullString, 0, 1, "%1", 2
       RegCloseKey hKey
    End If

    Rv = RegCreateKeyEx(HKEY_CLASSES_ROOT, "icofile\Shell\Open\Command", 0, vbNullString, 0, KEY_ALL_ACCESS, 0, hKey, 0)
    If Rv = 0 Then
       If Right(App.Path, 1) = "\" Then
          RegData = App.Path & "vbIconMaker.exe"
       Else
          RegData = App.Path & "\vbIconMaker.exe"
       End If
       RegData = RegData & " /open %1"
       RegSetValueEx hKey, vbNullString, 0, 1, RegData, Len(RegData)
       RegCloseKey hKey
    End If

End Sub

Public Sub PrepIconHeader()

    ID.idReserved = 0
    ID.idType = 1
    ID.idCount = 1

    IDE.bWidth = 32
    IDE.bHeight = 32
    IDE.bColorCount = 0
    IDE.bReserved = 0
    IDE.wPlanes = 1
    IDE.wBitCount = 24
    IDE.dwBytesInRes = 3240
    IDE.dwImageOffset = 22

    BIH.biSize = 40
    BIH.biWidth = 32
    BIH.biHeight = 64
    BIH.biPlanes = 1
    BIH.biBitCount = 24
    BIH.biCompression = 0
    BIH.biSizeImage = 3200
    BIH.biXPelsPerMeter = 0
    BIH.biYPelsPerMeter = 0
    BIH.biClrUsed = 0
    BIH.biClrImportant = 0

End Sub
