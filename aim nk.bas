Attribute VB_Name = "aimnk"
'aim nk.bas
'by: nk
'website: http://nkillaz.com
'this module was made in VB6 and is
'compatable for AIM 4.0+
'and can be used in VB4+

'i'd like to thank legion for helping
'me with some stuff in this module.

'i only took 5 codes from other
'people's modules. i took one from
'abbotaim2, and 4 from dos32, they got
'full credit. this bas has 87 subs.

'i'd also like to thank pat or jk
'and mavness, because i used there
'api spys for this module.

'shout outs: abbot, zb, progee, coby,
'dane, gleet, liquid, mopa, moss, chuck
'izekial, gpx, tipz, pac, ozzy,
'da joker, mizi, quirk, kloned, kaotic,
'bukataou, smokey, weed, legion
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetDesktopWindow& Lib "user32" ()
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowEnabled& Lib "user32" (ByVal hwnd As Long)
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function MoveWindow& Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HTCAPTION = 2
Public Const HWND_TOPMOST = -1

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

'Public Const LB_SETCURSEL = &H186
'Public Const LB_GETtext = &H189
'Public Const LB_GETTEXTLEN = &H18A
'Public Const LB_GETCOUNT = &H18B
'Public Const LB_FINDSTRING = &H18F
'Public Const LB_FINDSTRINGEXACT = &H1A2

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_RESTORE = 1
Public Const SW_SHOW = 5




Public Function ClickButton(button As Long)
Dim Click As Long
Click& = SendMessageByNum(button, WM_LBUTTONDOWN, &HD, 0)
Click& = SendMessageByNum(button, WM_LBUTTONUP, &HD, 0)
End Function

Public Function IF_FileExist(Path As String) As Boolean
    Dim a As String
    If Trim(Path) = "" Then IF_FileExist = False: Exit Function
    a = Dir(Path)
    If Len(a) = 0 Then
        IF_FileExist = False
    Else
        IF_FileExist = True
    End If
End Function



Function Count_List(lst As Listbox)
Count_List = lst.ListCount
'MsgBox "" & a & ""
End Function

Public Sub SetText(window As Long, Text As String)
Call SendMessageByString(window, WM_SETTEXT, 0&, Text)
End Sub

Sub Mass_IM(List As Listbox, Text As String, close_win As Boolean)
  'call Mass_IM(List1, Text1, True)
   'you may want to edit the message boxes
   'for your program.
   
   If List.ListCount = 0 Then
   MsgBox "You need people to IM.", vbCritical, "aim nk.bas"
   Exit Sub
   ElseIf Text = "" Then
   MsgBox "You need a message to send.", vbCritical, "aim nk.bas"
   Exit Sub
   End If
   Dim nk

    For nk = 0 To List.ListCount - 1
        Call IM_Send(List.List(nk), Text, True)
    Next nk
    
    If close_win = True Then

aimimessage = FindWindow("aim_imessage", vbNullString)
SendMessageLong aimimessage, WM_CLOSE, 0&, 0&
Else

End If
End Sub

Sub Change_IM_Caption(newcaption As String)
Dim aimimessage As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
Call SetText(aimimessage, newcaption)
End Sub

Sub Change_Chat_Caption(newcaption As String)
Dim aimchatwnd As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Call SetText(aimchatwnd, newcaption)
End Sub

Sub Close_AIM()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
SendMessage oscarbuddylistwin&, &H10, 0&, 0&
End Sub

Sub Change_AIM_Caption(newcaption As String)
Dim oscarbuddylistwin As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
Call SetText(oscarbuddylistwin, newcaption)
End Sub
Sub Chat_Close()
'call Chat_Close
'closes the chat.
Dim aimchatwnd As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
SendMessage aimchatwnd&, &H10, 0&, 0&
End Sub
Sub IM_Close()
'call IM_Close
'closes an IM.
Dim aimimessage As Long
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
SendMessage aimimessage&, &H10, 0&, 0&
End Sub

Sub Bot_Attention(Message)
'call Bot_Attention ("aim nk.bas")
'or
'call Bot_Attention (Text1)
ChatSend "(`(`÷·» A T T E N T I O N"
Pause 0.5
ChatSend (Message)
Pause 0.5
ChatSend "(`(`÷·» A T T E N T I O N"
End Sub

Public Function FilterHTML(ByVal HTML As String)

'by abbot, www.abbot3000.com
'he gets full credit for this one.

Dim Tag1 As Long, Tag2 As Long, STest As String
Dim LString As String, RString As String
Dim STest2 As Long, STest3 As Long

Do: DoEvents
    Tag1& = InStr(HTML$, "<")
    Tag2& = InStr(HTML$, ">")
    
    If Tag1& = 0 Or Tag2& = 0 Then Exit Do
Check1:
    DoEvents
    If Tag2& < Tag1& And Tag2& <> 0 Then
        STest$ = Mid$(HTML$, Tag2& + 1)
        STest2& = InStr(STest$, ">")
        If STest2& = 0 Then Exit Do
        Tag2& = Tag2& + STest2&
        GoTo Check1:
    End If
        

checkit:
    DoEvents
    If Tag1& = 0 Or Tag2& = 0 Then Exit Do
    
    STest$ = Mid$(HTML$, Tag1& + 1)
    STest2& = InStr(STest$, "<")
    STest3& = InStr(STest$, ">")
    
    If STest2& < STest3& And STest2& <> 0 Then
        Tag1& = Tag1& + STest2&
        GoTo checkit
    End If
        
    If STest3& = 0 Then Exit Do
    
    LString$ = Left$(HTML$, Tag1& - 1)
    RString$ = Mid$(HTML$, Tag2& + 1)
    HTML$ = LString$ & RString$
    
Loop

HTML$ = Replace(HTML$, "&amp;", "&")
HTML$ = Replace(HTML$, "&lt;", "<")
HTML$ = Replace(HTML$, "&lt;", "<")
    
FilterHTML = HTML$

End Function

Public Function GetChatText() As String

Dim aimchatwnd As Long, wndateclass As Long, ateclass As Long, TheText As String, TL As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
TL = SendMessageLong(ateclass, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(ateclass, WM_gettext, TL + 1, TheText)
TheText = Left(TheText, TL)
GetChatText = TheText

End Function

Public Function LastChatLine() As String


Dim str1 As String, intpos As Long
str1$ = GetChatText
intpos = InStrRev(str1$, "<BR>")
If Fnd <> 0 Then
    str1$ = Mid$(str1$, intpos + 4)
End If
LastChatLine = FilterHTML(str1$)

End Function

Public Function LastChatSN() As String
Dim str1 As String, intpos As Long

str1 = LastChatLine
intpos = InStr(str1, ":")
If intpos <> 0 Then
str1 = Left(str1, -1)
Else
str1 = "None"
End If

LastChatSN = str1

End Function

Public Sub AddRoomToList(Listbox As Listbox, adduser As Boolean)
'AddRoomToList List1, False
'that adds the room.

'AddRoomToList List1,True
'that adds the room including your sn.

Dim aimchatwnd As Long, oscartree As Long, LCount As Long
Dim int1 As Integer, txt1 As String
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
oscartree = FindWindowEx(aimchatwnd, 0&, "_oscar_tree", vbNullString)
LCount = SendMessageLong(oscartree, LB_GETCOUNT, 0&, 0&)
For int1 = 0 To LCount - 1
    txtLength = SendMessageByNum(oscartree&, LB_GETTEXTLEN, int1, 0)
    txt1 = String(txtLength, 1)
    Call SendMessageByString(oscartree&, LB_GETtext, int1, txt1)
    If adduser = False Then
    If txt1 = UserSN Then
        End If
        Else
        Listbox.AddItem txt1
        End If
        
Next

End Sub

Public Sub OnTop(Frm As Form)
    Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub NotOnTop(Frm As Form)
    Call SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub IM_Click_Block()
Dim aimimessage As Long, oscariconbtn As Long
aimimessage = FindWindow("aim_imessage", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimimessage, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn, WM_LBUTTONUP, 0&, 0&)
End Sub



Sub Form_Center(Frm As Form)
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub

Sub File_Delete(File)
Kill (File)
End Sub

Sub GetDate(Label)
'this has the date on a label.
Label.Caption = Date
End Sub

Public Sub ClipCopy(Text As String)
'call ClipCopy (Text1)
    Clipboard.Clear
    Clipboard.SetText (Text)
End Sub

Public Function ClipPaste() As String
    ClipPaste = Clipboard.GetText
End Function

Sub GetTime(Label)
'this puts the time on a label.
Label.Caption = Time
End Sub

Sub IP_Sniffer()
'the easiest way to make an ip sniffer
'is with mswinsock.ocx

'you'll need 2 text boxes, 1 list
'and 1 command button.
'text1 is for the sn to IM, and text2
'is for the fake link.

'put this code in the winsock.ocx:
'List1.AddItem Winsock1.RemoteHostIP

'in the form load put:
'Winsock1.LocalPort = 99 - this can be almost any port
'Winsock1.Listen

'in the command (send) button put:
'Call IM_Send(Text1.Text, "<a href=""" & Winsock1.LocalIP & """>" & Text2.text & "</a>")

'that is the easiest way to make an
'ip sniffer.

'by nk
End Sub

Sub ClearChat(Text As String)
Text = ""
Dim aimchatwnd As Long, WndAte32Class As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
WndAte32Class& = FindWindowEx(aimchatwnd&, 0&, "WndAte32Class", vbNullString)
Ate32Class& = FindWindowEx(WndAte32Class&, 0&, "Ate32Class", vbNullString)
SendMessageByString Ate32Class&, WM_SETTEXT, 0&, Text$
End Sub

Sub GoChat(room As String)
'goto a chat room
'call GoChat (Text1)
Dim oscarbuddylistwin As Long, editx As Long, oscariconbtn As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
editx = FindWindowEx(oscarbuddylistwin, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, "aim:gochat?roomname=" & room)
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn = FindWindowEx(oscarbuddylistwin, 0&, "_oscar_iconbtn", vbNullString)
Call ClickButton(oscariconbtn)
End Sub

Sub Chat_Invite(who As String, room As String, what As String)
'call Chat_Invite ("xnk","nk","come to this cool room.")
'or...
'call Chat_Invite (Text1,Text2,Text3)
'you can send this to multiple people just
'by putting a comma after each name.

Dim oscarbuddylistwin As Long, oscartabgroup As Long, oscariconbtn As Long, aimchatinvitesendwnd As Long, editx As Long


oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup = FindWindowEx(oscarbuddylistwin, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn = FindWindowEx(oscartabgroup, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(oscartabgroup, oscariconbtn, "_oscar_iconbtn", vbNullString)

Call ClickButton(oscariconbtn)


aimchatinvitesendwnd = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, 0&, "edit", vbNullString)

Call SendMessageByString(editx, WM_SETTEXT, 0&, who)


aimchatinvitesendwnd = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, 0&, "edit", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, editx, "edit", vbNullString)

Call SendMessageByString(editx, WM_SETTEXT, 0&, what)


aimchatinvitesendwnd = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, 0&, "edit", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, editx, "edit", vbNullString)
editx = FindWindowEx(aimchatinvitesendwnd, editx, "edit", vbNullString)

Call SendMessageByString(editx, WM_SETTEXT, 0&, room)


aimchatinvitesendwnd = FindWindow("aim_chatinvitesendwnd", vbNullString)
oscariconbtn = FindWindowEx(aimchatinvitesendwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatinvitesendwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimchatinvitesendwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call ClickButton(oscariconbtn)



End Sub







Sub Cancel_Invite()
'clicks the cancel button on the
'invite screen
Dim aimchatinvitesendwnd As Long, oscariconbtn As Long
aimchatinvitesendwnd& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
SendMessage oscariconbtn&, &H201, 0&, 0&
SendMessage oscariconbtn&, &H202, 0&, 0&
End Sub



Sub ChatSend_Bold(Text)
Call ChatSend("<b>" & Text & "</b>")
End Sub

Sub ChatSend_Underlined(Text)
Call ChatSend("<u>" & Text & "</u>")
End Sub

Sub ChatSend_Italic(Text)
Call ChatSend("<i>" & Text & "</i>")
End Sub

Sub IM_Open()
'opens an IM.
Dim oscarbuddylistwin As Long, oscartabgroup As Long, oscariconbtn As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_Oscar_TabGroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_Oscar_IconBtn", vbNullString)
SendMessage oscariconbtn&, &H201, 0&, 0&
SendMessage oscariconbtn&, &H202, 0&, 0&
End Sub


Sub Invite_Open()
'opens a chat invite
Dim oscarbuddylistwin As Long, oscartabgroup As Long, oscariconbtn As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_Oscar_TabGroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
SendMessage oscariconbtn&, &H201, 0&, 0&
SendMessage oscariconbtn&, &H202, 0&, 0&
End Sub

Sub ConnectToTalk()
'clicks the connect to talk button.
Dim oscarbuddylistwin As Long, oscartabgroup As Long, oscariconbtn As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_Oscar_TabGroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
SendMessage oscariconbtn&, &H201, 0&, 0&
SendMessage oscariconbtn&, &H202, 0&, 0&
End Sub

Sub Chat_Less()
'call Chat_Less
'makes the chat room smaller.
Dim aimchatwnd As Long, button As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
button& = FindWindowEx(aimchatwnd&, button&, "Button", vbNullString)
button& = FindWindowEx(aimchatwnd&, button&, "Button", vbNullString)
PostMessage button&, WM_LBUTTONDOWN, 0&, 0&
PostMessage button&, WM_LBUTTONUP, 0&, 0&
End Sub

Sub Chat_More()
'call Chat_More
'shows the advertisments at the buttom of
'the chat room.
Dim aimchatwnd As Long, button As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
button& = FindWindowEx(aimchatwnd&, 0&, "Button", vbNullString)
PostMessage button&, WM_LBUTTONDOWN, 0&, 0&
PostMessage button&, WM_LBUTTONUP, 0&, 0&
End Sub

Sub Hide_Ticker()
'hides the news ticker
'call Hide_Ticker
Dim AIMScrollTickerNewsWnd As Long
AIMScrollTickerNewsWnd& = FindWindow("AIM_ScrollTickerNewsWnd", vbNullString)
ShowWindow AIMScrollTickerNewsWnd&, 0
End Sub

Sub Show_Ticker()
'shows the news ticker
'call Show_Ticker
Dim AIMScrollTickerNewsWnd As Long
AIMScrollTickerNewsWnd& = FindWindow("AIM_ScrollTickerNewsWnd", vbNullString)
ShowWindow AIMScrollTickerNewsWnd&, 1
End Sub

Public Sub Shell_File(File As String)
Call Shell(File)
End Sub





Sub Close_Ticker()
'closes the news ticker
'call Close_Ticker
Dim AIMScrollTickerNewsWnd As Long
AIMScrollTickerNewsWnd& = FindWindow("AIM_ScrollTickerNewsWnd", vbNullString)
PostMessage AIMScrollTickerNewsWnd&, &H10, 0&, 0&
End Sub

Public Function FindAIM() As Long
FindAIM = FindWindow("_Oscar_BuddyListWin", vbNullString)
End Function

Public Function FindIM() As Long
FindIM = FindWindow("AIM_IMessage", vbNullString)
End Function

Public Function AIM_FindChat() As Long
AIM_FindChat = FindWindow("AIM_ChatWnd", vbNullString)
End Function
Sub PercentBar(picPic As Object, lngPercent As Long)
'how to use this: put this in a button
'or Form_Load

'Dim PercentCount As Long

'For PercentCount = 1 To 100 'starting number 1, ending percent 100%. you can change this too.

    'Call PercentBar(Picture1, PercentCount) 'draw percent in picture box
    
    'Call Pause(0.001)

'Next PercentCount
  
  With picPic
    .Cls

    .DrawMode = vbNotXorPen
    .BackColor = vbWhite
    .ForeColor = vbBlue
    .AutoRedraw = True
    .CurrentX = .ScaleWidth / 2 - .TextWidth(CStr(lngPercent&) & "%") / 2
    .CurrentY = .ScaleHeight / 2 - .TextHeight(CStr(lngPercent&) & "%") / 2
    picPic.Print CStr(lngPercent&) & "%"
    picPic.Line (1, 1)-((.Width / 100) * lngPercent&, .Height), vbBlue, BF
    .Refresh
  End With
  
End Sub

Public Function UserSN() As String
'this is how you can use this sub

'Dim X As String
'Call UserSN
'X = UserSN
'Label1.Caption = X
 Dim oscarbuddylistwin As Long
oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(oscarbuddylistwin, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(oscarbuddylistwin, WM_gettext, TL + 1, TheText)
TheText = Left(TheText, TL)
TheText = Left$(TheText, InStr(TheText, "'") - 1)
UserSN = TheText
End Function




Sub SaveText(txtSave As TextBox, Path As String)
'dos32
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Sub SaveText2(Path As String, Save As String)
    Dim a As Long
    On Error GoTo ErrorStop
    If FileExists(Path) = True Then
        Call Kill(Path)
    End If
    a = FreeFile
    Open Path For Binary As #a
    Put #a, 1, Save
ErrorStop:
    Close #a
End Sub

Sub List_Clear(List As Listbox)
'this is pretty simple.
List.Clear
End Sub

Public Sub Form_ExitRight(Form As Form)
Do
    DoEvents
    Form.Left = Trim(Str(Int(Form.Left) + 300))
Loop Until Form.Left > Screen.Width
End Sub

Public Sub Form_ExitLeft(Form As Form)
Do
     DoEvents
     Form.Left = Trim(Str(Int(Form.Left) - 300))
Loop Until Form.Left < -Form.Width
End Sub

Public Function Get_INI(Section As String, Keyword As String, Path As String) As String
    Dim a As String
        a = String(255, Chr(0))
        Get_INI = Left(a, GetPrivateProfileString(Section, Keyword, "", a, Len(a), Path))
End Function

Public Sub Write_INI(Section As String, Keyword As String, NewWord As String, Path As String)
    Call WritePrivateProfileString(Section, Keyword, NewWord, Path)
End Sub





Sub LoadText(txtLoad As TextBox, Path As String)
'dos32
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub



Sub IM_Linker(ScreenName As String, URL As String, Description As String, close_win As Boolean)
'call IM_Linker("xnk","http://nkillaz.com","cool website",True")
'or
'call IM_Linker (Text1, Text2, Text3, True)
'Text1 = the screen name to IM
'Text2 = the web adress of the site
'Text3 = what the link will say
'True,False - close im or dont after sent

Call IM_Send(ScreenName, "<a href=""" + URL + """>" + Description + "", True)

If close_win = True Then

Call SendMessageLong(aimimessage, WM_CLOSE, 0&, 0&)
Else

End If
End Sub

Sub Chat_Linker(URL As String, Description As String)
Call ChatSend("<a href=""" + URL + """>" + Description + "")
End Sub


Sub Form_Unload(Frm As Form)
'call Form_Unload (Me)
Unload Frm
End Sub

Sub Bot_FakeProgram(Name, MadeBy)
ChatSend "(`(`÷·» " & Name
Pause 0.5
ChatSend "(`(`÷·» " & MadeBy
End Sub

Public Sub Form_ExitUp(Frm As Form)
Do
Pause (0.3)
Frm.Move Frm.Left + 0, Frm.Top - 120
Loop Until Frm.Top - Frm.Top > Frm.Top
If Frm.Top - Frm.Top > Frm.Top Then
End
End If
End Sub

Public Sub Form_ExitDown(Frm As Form)
Do
Pause (0.3)
Frm.Move Frm.Left + 0, Frm.Top + 120
Loop Until Frm.Top > Screen.Height
If Frm.Top > Screen.Height Then
End
End If
End Sub



Sub Click_GoToURL()
Dim oscarbuddylistwin As Long, oscariconbtn As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, 0&, "_Oscar_IconBtn", vbNullString)
SendMessage oscariconbtn&, &H201, 0&, 0&
SendMessage oscariconbtn&, &H202, 0&, 0&
End Sub

Sub Play_Wav(Wav)
SoundName$ = Wav
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)
End Sub

Public Sub FormMove(Form As Form)
'put this in the MouseDown
'of a label, Image Box, or Picture Box
ReleaseCapture
x = SendMessage(Form.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Sub Chat_Scroll(Text As String)
'put this in a timer with an interval
'of 1. and have the timer enabled = false
'have a start and stop button.
'in the start button, put
'Timer1.Enabled = True
'in the stop button, put
'Timer1.enabled = false
Call ChatSend(Text)
Pause 0.6
End Sub

Sub IM_FaceHell(ScreenName As String, close_win As Boolean)
'people seem to like doing this :\

Call IM_Send(ScreenName, ":):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):)", True)
Pause 0.6
Call IM_Send(ScreenName, ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:", True)
Pause 0.6
Call IM_Send(ScreenName, ":):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):):)", True)
Pause 0.6
Call IM_Send(ScreenName, ":(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:(:", True)
If close_win = True Then

Call SendMessageLong(aimimessage, WM_CLOSE, 0&, 0&)
Else

End If

End Sub



Sub Chat_Scroll3(Text As String)
Call ChatSend(Text)
Pause 0.5
Call ChatSend(Text)
Pause 0.5
Call ChatSend(Text)
End Sub

Sub Chat_Scroll10(Text As String)
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
Pause 0.6
Call ChatSend(Text)
End Sub

Sub Pause(Time)
    Dim Current
    Current = Timer
    Do While Timer - Current < Val(Interval)
    DoEvents
    Loop
End Sub

Sub IM_Send(who As String, what As String, close_win As Boolean)
'send an instant message
'call SendIM("xnk","nice bas file.",True)
'or
'call SendIM(Text1,Text2,True)

'if you have True at the end, the IM
'closes once you send it. if you have a
'False at the end, the IM sends but
'doesnt close.
Dim oscarbuddylistwin As Long, oscartabgroup As Long, oscariconbtn As Long, aimimessage As Long, oscarpersistantcombo As Long, editx As Long, wndateclass As Long, ateclass As Long

oscarbuddylistwin = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup = FindWindowEx(oscarbuddylistwin, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn = FindWindowEx(oscartabgroup, 0&, "_oscar_iconbtn", vbNullString)

Call ClickButton(oscariconbtn)

aimimessage = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo = FindWindowEx(aimimessage, 0&, "_oscar_persistantcombo", vbNullString)
editx = FindWindowEx(oscarpersistantcombo, 0&, "edit", vbNullString)

Call SendMessageByString(editx, WM_SETTEXT, 0&, who)


wndateclass = FindWindowEx(aimimessage, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimimessage, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass, WM_SETTEXT, 0&, what)
oscariconbtn = FindWindowEx(aimimessage, 0&, "_oscar_iconbtn", vbNullString)
Call ClickButton(oscariconbtn)
If close_win = True Then

Call SendMessageLong(aimimessage, WM_CLOSE, 0&, 0&)
Else

End If
End Sub

Sub Hide_AIM()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowWindow oscarbuddylistwin&, 0
End Sub

Sub Show_AIM()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowWindow oscarbuddylistwin&, 1
End Sub



Sub Hide_Chat()
Dim aimchatwnd As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
ShowWindow aimchatwnd&, 0
End Sub

Sub Show_Chat()
Dim aimchatwnd As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
ShowWindow aimchatwnd&, 1
End Sub

Sub Hide_IM()
Dim aimimessage As Long
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
ShowWindow aimimessage&, 0
End Sub

Sub Hide_Chat_Meter()
Dim aimchatwnd As Long, OscarRateMeter As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
OscarRateMeter& = FindWindowEx(aimchatwnd&, 0&, "_Oscar_RateMeter", vbNullString)
ShowWindow OscarRateMeter&, 0
End Sub

Sub Show_Chat_Meter()
Dim aimchatwnd As Long, OscarRateMeter As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
OscarRateMeter& = FindWindowEx(aimchatwnd&, 0&, "_Oscar_RateMeter", vbNullString)
ShowWindow OscarRateMeter&, 1
End Sub

Sub Show_IM()
Dim aimimessage As Long
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
ShowWindow aimimessage&, 1
End Sub

Sub Show_IM_Meter()
Dim aimimessage As Long, OscarRateMeter As Long
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
OscarRateMeter& = FindWindowEx(aimimessage&, 0&, "_Oscar_RateMeter", vbNullString)
ShowWindow OscarRateMeter&, 1
End Sub
Sub Hide_IM_Meter()
Dim aimimessage As Long, OscarRateMeter As Long
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
OscarRateMeter& = FindWindowEx(aimimessage&, 0&, "_Oscar_RateMeter", vbNullString)
ShowWindow OscarRateMeter&, 0
End Sub

Public Function AIM_chatsend(send_text As String, Optional asciii As String)

'set text windows

Dim aimchatwnd As Long, wndateclass As Long, ateclass As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, 0&, "wndate32class", vbNullString)
wndateclass = FindWindowEx(aimchatwnd, wndateclass, "wndate32class", vbNullString)
ateclass = FindWindowEx(wndateclass, 0&, "ate32class", vbNullString)
'If Len(send_text) > 1000 Then send_text = color1 & "[" & color2 & "Too much text to send to chat" & color1 & "]"
Call SendMessageByString(ateclass, WM_SETTEXT, 0&, ascii & send_text)

Call clyck_send_button
Call TimeOut7(0.9)
End Function

Function FileExists(TheFile) As Boolean
On Error GoTo fileisgone
a = TheFile
b = FileLen(a)
FileExists = True
Exit Function
fileisgone:
FileExists = False
Exit Function
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
