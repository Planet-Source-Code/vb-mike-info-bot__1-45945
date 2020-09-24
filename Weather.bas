Attribute VB_Name = "WeatherMod"
Option Explicit
Public Type DialerInfoType
        account As String
        TotalMinutes_Calls As String
        OverallAverage As String
        ThisMonthMinutes_Calls As String
        MonthAverage As String
        DBYMinutes_Calls As String
        YesterdayMinutes_Calls As String
        TodayMinutes_Calls As String
        Cash As String
End Type

Public Type WeatherInfoType
        CurrentC As Integer
        CurrentF As Integer
        CurrentCond As String
        FeelsLike As String
        UVIndex As String
        DewPoint As String
        Humidity As String
        Visibility As String
        Pressure As String
        Wind As String
        City As String
        State As String
        invalid As String
End Type
Public WeatherInfo As WeatherInfoType
Public DialerInfo As DialerInfoType
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
Private Const BUFFER_LEN = 256

Public Function Parse(ByVal sInput As String, lField As Long, sDelimiter As String) As String
    Dim sTemp As String
    Dim lPos, lLen, lCnt, lTmp As Long
    If lField < 0 Then Parse = "": Exit Function
    sTemp = "" & Trim(sInput) & sDelimiter
    For lCnt = 1 To lField
        lLen = Len(sTemp)
        lPos = InStr(1, sTemp, sDelimiter)
        lTmp = lLen - lPos
        sTemp = Right(sTemp, lTmp)
    Next lCnt
    lPos = InStr(1, sTemp, sDelimiter)
    If lPos > 0 Then
        sTemp = Left(sTemp, lPos - 1)
    End If
    Parse = Trim(sTemp)
End Function

Public Function GetSource(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    hSession = InternetOpen("", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    If hInternet Then
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
    iResult = InternetCloseHandle(hInternet)
    GetSource = sData
    sData = ""
End Function

Public Function GetWeatherInfo(ZipCode As String)
On Error Resume Next
Dim Source As String
Dim CurCstr As Integer
Dim WData As String
Dim Temp As String

CurCstr = 0
Source = ""
DoEvents
Source = GetSource("http://www.weather.com/weather/local/" & ZipCode) ' Get the web Source
If InStr(1, Source, "Sorry, the page you requested was not found on weather.com", vbTextCompare) <> 0 Then WeatherInfo.invalid = "Invalid"
' Current City / State
CurCstr = InStr(1, Source, "Local Forecast for ")
WData = Mid(Source, CurCstr + 19, 50)
Temp = Parse(WData, 0, "(")
WeatherInfo.City = (Left(Temp, InStr(1, Temp, ",", vbTextCompare) - 1))
WeatherInfo.State = (Right(Temp, 3))

' Current TempC & F
CurCstr = InStr(Source, "obsTempTextA>")
WData = Mid(Source, CurCstr + 13, 5)
WeatherInfo.CurrentF = CInt(Parse(WData, 0, "&"))
WeatherInfo.CurrentC = CInt((WeatherInfo.CurrentF - 32) * 5 / 9)
'END
'Get Current Condition
CurCstr = InStr(Source, "obsTextA>")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.CurrentCond = Parse(WData, 0, "</B>")
If InStr(WeatherInfo.CurrentCond, "HEAD") Then
WeatherInfo.CurrentCond = " "
WeatherInfo.CurrentF = 0
WeatherInfo.CurrentC = 0
End If
'End
'Get Feels Like
CurCstr = InStr(CurCstr + 9, Source, "obsTextA")
WData = Mid(Source, CurCstr + 23, 30)
WeatherInfo.FeelsLike = CInt(Parse(WData, 0, "&deg"))
If InStr(WeatherInfo.FeelsLike, " ") Or InStr(WeatherInfo.FeelsLike, "HEAD") Then
WeatherInfo.FeelsLike = " "
End If
If WeatherInfo.FeelsLike = WeatherInfo.CurrentF Then
WeatherInfo.FeelsLike = ""
Else
WeatherInfo.FeelsLike = color1 & ", Feels like " & color2 & WeatherInfo.FeelsLike & "º"
End If
'End
'get UV index
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.UVIndex = Parse(Replace(WData, "&nbsp;", " "), 0, "</TD>")
If InStr(WeatherInfo.UVIndex, "HEAD") Then
WeatherInfo.UVIndex = " "
End If
'End
'get Dew Point
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.DewPoint = Parse(WData, 0, "&deg")
If InStr(WeatherInfo.DewPoint, "HEAD") Then
WeatherInfo.DewPoint = " "
End If
'End
'get Humidity
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.Humidity = Parse(WData, 0, "%</TD>")
If InStr(WeatherInfo.Humidity, "HEAD") Then
WeatherInfo.Humidity = " "
End If
'end
'get Visibility
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.Visibility = Parse(WData, 0, "</TD>")
If InStr(WeatherInfo.Visibility, "HEAD") Then
WeatherInfo.Visibility = " "
End If
'end
'get Pressure
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 30)
WeatherInfo.Pressure = Parse(WData, 0, "</TD>")
If InStr(WeatherInfo.Pressure, "HEAD") Then
WeatherInfo.Pressure = " "
End If
'end
'get Wind
CurCstr = InStr(CurCstr + 9, Source, "obsInfo2")
WData = Mid(Source, CurCstr + 9, 40)
WeatherInfo.Wind = Parse(WData, 0, "&nbsp") & " Mph"
If WeatherInfo.Wind = "calm Mph" Then WeatherInfo.Wind = "calm"
If InStr(WeatherInfo.Wind, "HEAD") Then
WeatherInfo.Wind = " "
End If
'END
End Function



Sub SetDialerStats()
Dim Source As String, Start As Integer, Temp As String
Source = Form1.Text3
Start = (InStr(1, Source, "<TD vAlign=center>", vbTextCompare) + 18)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.account = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "Dialer Profits</td>", vbTextCompare) + 25)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.TotalMinutes_Calls = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.OverallAverage = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.ThisMonthMinutes_Calls = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)
DialerInfo.Cash = Left(DialerInfo.ThisMonthMinutes_Calls, InStr(1, DialerInfo.ThisMonthMinutes_Calls, "/", vbTextCompare) - 1)
DialerInfo.Cash = Val(DialerInfo.TotalMinutes_Calls) * 0.2
DialerInfo.Cash = FormatCurrency(DialerInfo.Cash)
Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.MonthAverage = Left(Source, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.DBYMinutes_Calls = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.YesterdayMinutes_Calls = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

Start = (InStr(1, Source, "<td>", vbTextCompare) + 4)
Source = Mid(Source, Start, Len(Source) - Start)
DialerInfo.TodayMinutes_Calls = Mid(Source, 1, InStr(1, Source, "</td>", vbTextCompare) - 1)

End Sub
