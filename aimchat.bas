Attribute VB_Name = "aimchat"

'*************************************************************
' For the 32 bit code that is generated with this spy
' to work, you will need to put these functions/consts
' in your module (*.bas file). Just Select it all,
' copy it, and paste it in.
' thanks to patorjk.com for their apispy
'

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5


Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20



'Public oldlastlyne As String, newlastlyne As String
'Public aim_whosaid As String, aim_whatsaid As String, aim_datesaid As String
'
'

Public Function chatlastlyne(lyne_chatlast As String) As String
Dim last_br As Long, lastlengthnew As Long
Dim verylastlyne As String

last_br = InStrRev(lyne_chatlast, "<BR>", , vbTextCompare) '+ 1

If InStrRev(lyne_chatlast, "<BR>", , vbTextCompare) = False Then
chatlastlyne = verylastlyne
Exit Function
End If

lastlengthnew = Len(lyne_chatlast) - last_br

verylastlyne = Mid(lyne_chatlast, last_br, lastlengthnew)
chatlastlyne = verylastlyne

End Function
Public Function chatlastlyne_nohtml(nohtml_chatlast As String) As String
Dim lynelength As Long
Dim no_html As String, no_html2 As String

lynelength = Len(nohtml_chatlast)
i = 1
Do Until i > lynelength - 1

If Mid(nohtml_chatlast, i, 1) = "<" Then
    While Mid(nohtml_chatlast, i, 1) <> ">"
    i = i + 1
    If i > lynelength - 1 Then Exit Do
    Wend
End If

no_html = no_html + Mid(nohtml_chatlast, i, 1)
i = i + 1

Loop

' thys is to get rid of the closing >

no_html2 = no_html
no_html = ""

lynelength = Len(nohtml_chatlast)
i = 1

Do Until i > lynelength - 1

If Mid(no_html2, i, 1) = ">" Then
    i = i + 1
Else

no_html = no_html + Mid(no_html2, i, 1)
i = i + 1

End If

Loop

chatlastlyne_nohtml = no_html

End Function
Public Function get_chat() As String
'get aim windowtext

Dim aimchatwnd As Long, wndateclass As Long, ateclass As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(ateclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(ateclass, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
' is thys needed anymore?? Clipboard.SetText TheText, vbCFText
get_chat = TheText

End Function

Public Function get_chatsendtext(chat_to_send As String)

'get chatsend text... thys is useless... no wait.. it can be used to..
'get what someone is typing when a bot has to interrupt and send somethyng..
'you get the text being typed... clear the send box.. send what you need.. then put it back

Dim aimchatwnd As Long, wndateclass As Long, ateclass As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(ateclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(ateclass, WM_GETTEXT, TL + 1, TheText)


TheText = Left(TheText, TL)

End Function

Public Function AIM_chatsend(send_text As String, Optional asciii As String)

'set text windows

Dim aimchatwnd As Long, wndateclass As Long, ateclass As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
If Len(send_text) > 2000 Then send_text = color1 & "[" & color2 & "Too much text to send to chat" & color1 & "]"
Call SendMessageByString(ateclass, WM_SETTEXT, 0&, ascii & send_text)

Call clyck_send_button
Call TimeOut7(0.9)
End Function

Public Function clyck_send_button()

Dim aimchatwnd As Long, oscariconbtn As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)

Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_KEYUP, VK_SPACE, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)

End Function

Public Function get_chat_line() As String

Dim scoop As String
Dim scoopsyze As Long

scoopsyze = 500

' get the text
scoop = get_chat()
scoop = Right(scoop, scoopsyze)



' assign text to get_chat_line
get_chat_line = scoop


End Function
