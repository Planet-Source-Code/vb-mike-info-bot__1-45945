Attribute VB_Name = "gayol7"
' Hi thanks for dl'ing Gayol7 by caloric.
' This was made using Visual Basic 6.0 Professional.
' If you have any aol 7.0 code you'd like to
' contribute please e-mail me at fiowmotion@aol.com
'                  Version [1.0 beta]
' This is not completely finished..
' For updates goto:
' http://freefilesonline.crosswinds.net/gayol

' Thanks, caloric(fiowmotion@aol.com)
' Some code was made using partial code
' from DoS32.bas and other various persons.
' Credit is given where credit is due.
Option Explicit


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

'Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const WM_CLEAR7 = &H303
Public Const LB_GETCOUNT7 = &H18B
Public Const LB_GETITEMDATA7 = &H199
Public Const LB_GETTEXT7 = &H189
Public Const LB_GETTEXTLEN7 = &H18A
Public Const LB_SETCURSEL7 = &H186
Public Const LB_SETSEL7 = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1



Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR7 = &H102
Public Const WM_CLOSE7 = &H10
Public Const WM_COMMAND7 = &H111
Public Const WM_GETTEXT7 = &HD
Public Const WM_GETTEXTLENGTH7 = &HE
Public Const WM_KEYDOWN7 = &H100
Public Const WM_KEYUP7 = &H101
Public Const WM_LBUTTONDBLCLK7 = &H203
Public Const WM_LBUTTONDOWN7 = &H201
Public Const WM_LBUTTONUP7 = &H202
Public Const WM_MOVE7 = &HF012
Public Const WM_SETTEXT7 = &HC
Public Const WM_SYSCOMMAND7 = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
'Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        Y As Long
End Type
Public Function FindRoomFull()
' Finds Room is Full window & closes it
Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Function

Public Sub ChatSend(Chat As String, Optional asciii As String)
' Sends text to AOL Chat Room
' ie.  Call Chatsend7("<-- prog name, etc")
Dim AoRoom As Long, AoRich As Long
If Len(Chat) > 1860 Then ChatSend (Left(Chat, 1860)): Chat = "<font color=#fefcfe><i <font color=#0000ff>" & Mid(Chat, 1860, Len(Chat) - 1860) & "<font color=#fefcfe> > ." 'ChatSend (Left(Chat, 1862)): Chat = Mid(Chat, 1862, Len(Chat) - 1862) '"Text is to long to send to chat."
If Len(Chat) > 232 Then Chat = "<font color=#fefcfe><i <font color=#0000ff>" & Chat & "<font color=#fefcfe> > ."
    'If frmMain.chkChatSend.Value = vbUnchecked Then Exit Sub
    AoRoom = FindChat
    If AoRoom = 0 Then Exit Sub
    AoRich = FindWindowEx(AoRoom, 0, "RICHCNTL", vbNullString)
    
    Do
    DoEvents
    Form1.Text1 = GetText(AoRich)
    Loop Until Form1.Text1 = ""
    Call SendMessageByString(AoRich, WM_SETTEXT, 0, asciii & color1 & Chat)
    Call SendMessageLong(AoRich, WM_CHAR, VK_RETURN, 0)
    Call TimeOut7(0.9)
    End Sub
Public Sub ChatSend2(Chat As String)

Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(RICHCNTL, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(RICHCNTL, WM_gettext, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "" Then GoTo justsendchat
Call SendMessageByString(RICHCNTL, WM_CLEAR, 0&, 0&)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, TheText)
Exit Sub
justsendchat:
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub SendToChat(Message As String)

' This was taken from bofen32.bas and modified
' it's supposed to be like ChatSend2, but again
' doesnt work.

Dim RICHCNTL As Long, TextLen As Long, RICHCNTLTxt As String
Dim aolframe As Long, mdiclient As Long, aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
TextLen& = SendMessage(RICHCNTL&, WM_GETTEXTLENGTH, 0&, 0&)
RICHCNTLTxt$ = String(TextLen&, 0&)
Call SendMessageByString(RICHCNTL&, WM_gettext, TextLen& + 1&, RICHCNTLTxt$)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Message$)
Call SendMessageByNum(RICHCNTL&, WM_CHAR, 13&, 0&)
If Len(RICHCNTLTxt$) <> 0& Then Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, RICHCNTLTxt$)
End Sub
Public Sub WaitForTextToLoad(hwnd As Long)

' This sub will wait for all
' the text to load in a text field.
' Taken from bofen32.bas

Dim Count1 As Long, Count2 As Long, Count3 As Long
Do: DoEvents
    Count1& = Len(GetText(hwnd&))
    Call timeout(0.5)
    Count2& = Len(GetText(hwnd&))
    Call timeout(0.5)
    Count3& = Len(GetText(hwnd&))
Loop Until Count2& = Count1& And Count3& = Count1& And Count3& <> 0&
End Sub
Public Sub WriteToINI(AppName As String, KeyName As String, KeyValue As String, FileName As String)

' This function will write to an INI file.  Here's an example...
' WriteToINI("My App", "My Keyname", "My KeyValue", "myfile.ini")
' Taken from bofen32.bas

Call WritePrivateProfileString(AppName$, KeyName$, KeyValue$, FileName$)
End Sub



Function GetChatText()
' Gets all of the Chattext.
' Not sure if this works.
'Dim ChatText
'Dim AORich As Long
'Dim Room As Long
'Room& = FindChat
'AORich& = FindChildByClass(Room&, "RICHCNTL")
'ChatText = GetText(AORich&)
'GetChatText = ChatText
End Function
Public Sub SaveListBox7(Directory As String, TheList As Listbox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub FormOnTop7(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub Keyword7(KW As String)
' Goes to specified Keyword.
' Can also be used to open websites using AOL browser(yuck)
If Form1.AOL9.Value = 1 Then Call Keyword9(KW): Exit Sub
Dim aolframe As Long, aoltoolbar As Long, AOLCombobox As Long
Dim editx As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
AOLCombobox = FindWindowEx(aoltoolbar, 0&, "_aol_combobox", vbNullString)
editx = FindWindowEx(AOLCombobox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, KW$)
Call SendMessageLong(editx&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(editx&, WM_CHAR, VK_RETURN, 0&)

End Sub
Public Sub Keyword9(KW As String)
Dim aolframe As Long, aoltoolbar As Long, aoledit As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aoledit = FindWindowEx(aoltoolbar, 0&, "_aol_edit", vbNullString)
 
Call SendMessageLong(aoledit, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(aoledit, WM_LBUTTONUP, 0&, 0&)
 
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, "")
 
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, KW)
Call SendMessageLong(aoledit, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(aoledit, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub EnterPR(room As String)
' Enters a specified private room
Call Keyword7("aol://2719:2-2-" & room)
End Sub
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(MyString$, (ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, (MyString$), (ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Sub WaitForOKOrRoom(room As String)
' Use this to make a room buster.
' Waits for 'Room is Full' window or
' the chat room itself. If 'Room Is Full' pops up
' it will close it.
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    room$ = (ReplaceString(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindChat&)
        RoomTitle$ = (ReplaceString(room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or room$ = RoomTitle$
    DoEvents
    If FullWindow& <> 0& Then
        Do
            DoEvents
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            FullWindow& = FindWindow("#32770", "America Online")
            FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
        Loop Until FullWindow& = 0& And FullButton& = 0&
    
    End If
    DoEvents

End Sub
Public Sub clickToolbar2(IconNumber&, letter$, letter2$)
' This code provided by DaCrazyOne(ThatsMrPsP2U@aol.com)
Dim aolframe As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim aolicon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
aolicon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
aolicon = FindWindowEx(clickToolbar2, aolicon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
letter2 = Asc(letter2)
Call PostMessage(menu, WM_CHAR, letter, 0&)
Call PostMessage(menu, WM_CHAR, letter2, 0&)
End Sub

Public Sub clickToolbar(IconNumber&, letter$)
' Provided by DaCrazyOne
' ie. Call clicktoolbar("3", "F")
' 3 being the iconnumber, and the letter being
' the underlined letter in a word in a drop
' down menu. Like if you were to press ALT+the underlined letter.
' icon numbers: 0 = Mail, 3 = People
'               6 = Services, 9 = Settings
'               11 = Favorites
Dim aolframe As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim aolicon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
aolicon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
aolicon = FindWindowEx(clickToolbar2, aolicon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter, 0&)
End Sub
Public Sub ChatNow()
' Enter a Lobby
Call clickToolbar("3", "N")
End Sub
Public Sub AddRoom7(List As String)
' This adds the chat room to a listbox
' This is buggy and doesn't want to work
' If you can make a working AddRoom
' sub for aol 7.0 please e-mail me at fiowmotion@aol.com
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aollistbox As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aollistbox = FindWindowEx(aolchild, 0&, "_aol_listbox", vbNullString)
'Call AddAOLListToListbox(aollistbox, list)
End Sub
Public Sub AddAOLListToListbox(ListToGet As Long, ListToPut As Listbox)
  ' Use ADDROOM
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ListItem As String
    Dim psnHold As Long, rBytes As Long, i As Integer
    Dim sThread As Long, mThread As Long
    ' Obtain the identifiers of a thread and process that are associated
    ' with the window. A process is a running application and a thread
    ' is a task that the program is doing (like a program could be doing
    ' several things, each of these things would be a thread).
    sThread = GetWindowThreadProcessId(ListToGet, cProcess)
    ' Open the handle to the existing process
    mThread = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess)
    If mThread <> 0 Then
        For i = 0 To SendMessage(ListToGet, LB_GETCOUNT, 0, 0) - 1
            ListItem = String(4, vbNullChar)
            itmHold = SendMessage(ListToGet, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
            itmHold = itmHold + 24
            ' Read memory from the address space of the process
            Call ReadProcessMemory(mThread, itmHold, ListItem, 4, rBytes)
            Call CopyMemory(psnHold, ByVal ListItem, 4)
            psnHold = psnHold + 6
            ListItem = String(16, vbNullChar)
            Call ReadProcessMemory(mThread, psnHold, ListItem, Len(ListItem), rBytes)
            ' cut nulls off
            ListItem = Left(ListItem, InStr(ListItem, vbNullChar) - 1)
            ListToPut.AddItem ListItem
        Next i
        Call CloseHandle(mThread)
    End If
End Sub


Public Sub ListRemoveBlanks(TheList As Listbox)
' Self-explanitory
Dim Count&, Count2&
If TheList.ListCount = 0 Then Exit Sub
Do
DoEvents
Count& = 1
Do
DoEvents
If TheList.List(Count&) = "" Then TheList.RemoveItem (Count&)
Count& = Count& + 1
Count2& = TheList.ListCount
Loop Until Count& >= Count2&
Loop Until InStr(TheList.hwnd, "") = 0
End Sub
Public Sub KillDupes7(TheList As Listbox)
' Kills duplicates in a listbox.
Dim Count&, Count2&, Count3&
If TheList.ListCount = 0 Then Exit Sub
For Count& = 0 To TheList.ListCount - 1
DoEvents
For Count2& = Count& + 1 To TheList.ListCount - 1
DoEvents
If TheList.List(Count&) = TheList.List(Count2&) Then TheList.RemoveItem (Count2&)
Next Count2&
Next Count&
End Sub
Public Sub TimeOut7(Length&)
Dim Time As Long
Time = Timer
Do
DoEvents
Loop Until Timer - Time >= Length
End Sub
Public Function direxists(search As String) As Boolean
' This has nothing to do with aol
' so version does not matter.
' I got this code from http://www.w4sp.com
' But I am not sure as to who wrote it.
' Checks to see if a given directory exists.
If Right(search$, 1) <> "" + "\" Then
search$ = search$ + "\"
End If
If Dir(search$) <> "" Then
direxists = True
Else
direxists = False
End If
End Function
Public Sub ReadNew()
' Opens New Mail
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call SendMessageLong(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(aolicon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub WriteMail()
' Opens Write Mail
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call SendMessageLong(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(aolicon, WM_KEYUP, VK_SPACE, 0&)
End Sub





Public Sub AddListToListbox(TheList As Long, NewList As Listbox)
    ' This sub will only work with standard listboxes.
    Dim LCount As Long, Item As String, i As Integer, TheNull As Integer
    ' get the item count in the list
    LCount = SendMessageLong(TheList, LB_GETCOUNT, 0&, 0&)
    For i = 0 To LCount - 1
        Item = String(255, Chr(0))
        Call SendMessageByString(TheList, LB_GETtext, i, Item)
        TheNull = InStr(Item, Chr(0))
        ' remove any null characters that might be on the end of the string
        If TheNull <> 0 Then
            NewList.AddItem Mid$(Item, 1, TheNull - 1)
        Else
            NewList.AddItem Item
        End If
    Next
End Sub
Public Function GetUser7()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim UserString As String
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolchild& = FindWindowEx(mdiclient, aolchild, "aol child", vbNullString)
UserString$ = GetCaption(aolchild&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser7 = UserString
        Exit Function
    Else
        Do
            aolchild& = FindWindowEx(mdiclient&, aolchild&, "AOL Child", vbNullString)
            UserString$ = GetCaption(aolchild&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser7 = UserString
                Exit Function
            End If
        Loop Until aolchild& = 0&
    End If
    GetUser7 = ""
End Function
Public Sub ClickIdleOff()
' Closes the window that pops and up and says
' Idle Message is now Off or something like that.
' Kind of useless but thought I'd add it anyway.
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(aolicon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function DateDiffEx(starttime As Date, EndTime As Date) As String
    DateDiffEx = DateDiffExFormat(DateDiff("d", starttime, EndTime) \ 365, "year")
    DateDiffEx = DateDiffEx & DateDiffExFormat((DateDiff("s", starttime, EndTime) \ 86400) _
    Mod 365, "day")
    DateDiffEx = DateDiffEx & DateDiffExFormat((DateDiff("s", starttime, EndTime) \ 3600) _
    Mod 24, "hour")
    DateDiffEx = DateDiffEx & DateDiffExFormat((DateDiff("s", starttime, EndTime) \ 60) _
    Mod 60, "minute")
    DateDiffEx = DateDiffEx & DateDiffExFormat(DateDiff("s", starttime, EndTime) _
    Mod 60, "second")
    


    If Len(DateDiffEx) > 0 Then
        DateDiffEx = Mid(DateDiffEx, 1, Len(DateDiffEx) - 2)
    End If
End Function


Private Function DateDiffExFormat(inputValue As Long, unitValue As String) As String


    If inputValue <> 0 Then
        DateDiffExFormat = inputValue & " " & unitValue & IIf(inputValue <> 1, "s", "") & ", "
    End If
End Function


Public Sub SendIM7(Person As String, Message As String)
Call Keyword7("aol://9293:" & Person & ":" & Message)
End Sub

Public Function icon_ChatIgnore()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatIgnore = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function

Public Function icon_ChatIgnore9()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatIgnore9 = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Public Sub ChatIgnoreUser(strUser As String, Optional blnPartial As Boolean = False, Optional blnTellChat As Boolean = True)
'%% Its Not Case Sensitive :)
'%% Ex:
'%%     Call ChatIgnoreUser("loserdude551")
'
On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    
    'make it lcased
    strUser = LCase(strUser)
    
    'Find List
    rList = ChatPeopleHereList()
    'Open Thread
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    'this 'OpenProcess' API Call allows for us to read the contents of the
    'listbox in a nonconventional way, since conventional methods dont work :)
    'We will open it through mThread&
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    'If its valid, then commence adding
    If mThread& Then
        'Loop through the list items using my ListCount Procedure
        'We must subtract 1 since its zero-based
        For index& = 0 To ListCount(rList) - 1
            'Get the current screen name -the invalid characters
            ScreenName$ = String$(4, vbNullChar)
            'Get's the screen name from the list through itmHold&
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 28
            'Read the data from itmHold with the ReadProcessMemory API Call,
            'this will enable us to get the actual data via the "non-conventional" way :)
            'heh, trying to explain as simple as possible.. basically its reading the data
            'in a special way into th memory
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            'Now we will read that memory-added data
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            'Subtract invalid characters
            ScreenName$ = String$(16, vbNullChar)
            'Read into memory again
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            'Finally, we've got our SN
            ScreenName$ = LCase(Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1))
            
            'If its who we want, ignore/unignore 'em :)
            If (blnPartial = True And InStr(ScreenName, strUser)) Or (blnPartial = False And ScreenName = strUser) Then
                If LCase(GetUser) = ScreenName Then GoTo notme
                'DblClick SN
                Call ListSelect(rList, CInt(index&))
                'Get thier Screen Name
                If Form1.AOL9.Value = 0 Then
                    Call ClickIcon(icon_ChatIgnore)
                    Else
                    Call ClickIcon(icon_ChatIgnore9)
                End If
                'Close the processed memory thread
                Call CloseHandle(mThread)
                If blnTellChat = True Then
                ChatSend (ScreenName & ", has been ignored.")
                Else
                ChatSend (ScreenName & ", has been Unignored.")
                End If
                Exit Sub
            End If
notme:
            'Move on to next SN :)
        Next index&
        'Close the processed memory thread
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub ListSelect(lnglist As Long, intIndex As Integer)
'%% Just selects an item in a list (lngList)
'%% Ex:
'%%     Call ListSelect(AOLList, 4)
'
    'Selection API
    Call SendMessageLong(lnglist, LB_SETCURSEL, intIndex, 0&)
End Sub
Public Sub ClickIcon(Icon As Long)
'%% Clicks an AOL8 Icon using Left-Button Down (WM_LBUTTONDOWN)
'%% Followd by a KeyUp of a SpaceBar, this is necessary b/c AOL8's
'%% Icons are weird :), usually we can just use WM_LBUTTONUP
    Call SendMessageLong(Icon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(Icon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function ChatNumPeople() As Integer
'%% Returns number of people in chat room
'%% Ex:
'%%     intNumPeople = ChatNumPeople()
'%% ..or.. you could just do ListCount(ChatPeopleHereList)
'
    Dim lngroom As Long, aolstatic As Long
    lngroom = FindRoom
    'Find AOL's "people here"
    aolstatic = FindWindowEx(lngroom, 0&, "_aol_static", vbNullString)
    aolstatic = FindWindowEx(lngroom, aolstatic, "_aol_static", vbNullString)
    aolstatic = FindWindowEx(lngroom, aolstatic, "_aol_static", vbNullString)
    'Return number of people by using AOL's "people here" :)
    ChatNumPeople = Int(GetText(aolstatic))
    '..or simply..
    'ChatNumPeople = ListCount(ChatPeopleHereList)
End Function
Public Function ChatPeopleHereList() As Long
'%% Returns the "People Here" list handle in a chat room
'%% Ex:
'%%     Call ListCount(ChatPeopleHereList)
'
    ChatPeopleHereList = FindWindowEx(FindRoom, 0&, "_aol_listbox", vbNullString)
End Function
Public Function ChatEjected() As Long
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aollistbox As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aollistbox = FindWindowEx(aolchild, 0&, "_aol_listbox", vbNullString)
ChatEjected = FindWindowEx(aolchild, aollistbox, "_aol_listbox", vbNullString)

End Function
Public Function ListCount(Listbox As Long) As Long
'%% From PAT or JK's API Spy 5.1
'%% Counts and Returns number of Items in List via API Calls
    ListCount& = SendMessageLong(Listbox&, LB_GETCOUNT, 0&, 0&)
End Function
Public Sub AddAOL8ListToList(ListToGet As Long, ListToAddTo As ComboBox, Optional adduser As Boolean = False)
'%% This sub was made by myst. However, I modified it slightly to allow
'%% for it to work with all AOL8 ListBox's (not just 'People Here' List
'%% I also added my own comments in hopes that you could learn how
'%% it works :) Works much faster and better than the other sub,
'%% so use this unless it stops working for ya.
'%% Ex:
'%%     Call AddAOL8ListToList(ChatPeopleHereList, List1, False) '<- Adds to List1 w/o actual User

    On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long
    Dim sThread As Long, mThread As Long
    
    sThread& = GetWindowThreadProcessId(ListToGet, cProcess&)
    'this 'OpenProcess' API Call allows for us to read the contents of the
    'listbox in a nonconventional way, since conventional methods dont work :)
    'We will read that tasks of the window (thread) into mThread&
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    'If its valid, then commence adding
    If mThread& Then
        'Loop through the list items using my ListCount Procedure
        'We must subtract 1 since its zero-based
        For index& = 0 To ListCount(ListToGet) - 1
            'Get the current screen name -the invalid characters
            ScreenName$ = String$(4, vbNullChar)
            'Get's the screen name from the list through itmHold&
            itmHold& = SendMessage(ListToGet, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 28
            'Read the data from itmHold with the ReadProcessMemory API Call,
            'this will enable us to get the actual data via the "non-conventional" way :)
            'heh, trying to explain as simple as possible.. basically its reading the data
            'in a special way into th memory
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            'Now we will read that memory-added data
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            'Subtract invalid characters
            ScreenName$ = String$(16, vbNullChar)
            'Read into memory again
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            'Finally, we've got our SN
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            'Make sure, if the user doesn't want to add the actual user, then
            'dont add to list, otherwise, add :)
            If ScreenName$ <> GetUser$ Or adduser = True Then
                ListToAddTo.AddItem ScreenName$
            End If
            'Move on to next SN :)
        Next index&
        'Close the processed memory thread
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub ChatEjectUser(strUser As String, Optional blnPartial As Boolean = False, Optional blnTellChat As Boolean = True)
'%% Its Not Case Sensitive :)
'%% Ex:
'%%     Call ChatIgnoreUser("loserdude551")
'
On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    If Form1.AOL9.Value = 0 Then
    Call ClickIcon(FindPeopleHereTab)
    Else
    Call ClickIcon(FindPeopleHereTab9)
    End If
    'make it lcased
    strUser = LCase(strUser)
    
    'Find List
    rList = ChatPeopleHereList()
    'Open Thread
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    'this 'OpenProcess' API Call allows for us to read the contents of the
    'listbox in a nonconventional way, since conventional methods dont work :)
    'We will open it through mThread&
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    'If its valid, then commence adding
    If mThread& Then
        'Loop through the list items using my ListCount Procedure
        'We must subtract 1 since its zero-based
        For index& = 0 To ListCount(rList) - 1
            'Get the current screen name -the invalid characters
            ScreenName$ = String$(4, vbNullChar)
            'Get's the screen name from the list through itmHold&
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 28
            'Read the data from itmHold with the ReadProcessMemory API Call,
            'this will enable us to get the actual data via the "non-conventional" way :)
            'heh, trying to explain as simple as possible.. basically its reading the data
            'in a special way into th memory
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            'Now we will read that memory-added data
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            'Subtract invalid characters
            ScreenName$ = String$(16, vbNullChar)
            'Read into memory again
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            'Finally, we've got our SN
            ScreenName$ = LCase(Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1))
            
            'If its who we want, ignore/unignore 'em :)
            If (blnPartial = True And InStr(ScreenName, strUser)) Or (blnPartial = False And ScreenName = strUser) Then
                If LCase(GetUser) = ScreenName Then GoTo notme
                'DblClick SN
                Call ListSelect(rList, CInt(index&))
                'Get thier Screen Name
                If Form1.AOL9.Value = 0 Then
                Call ClickIcon(icon_ChatEject)
                Else
                Call ClickIcon(icon_ChatEject9)
                End If
                'Close the processed memory thread
                Call CloseHandle(mThread)
                'If blnTellChat = True Then ChatSend (ScreenName & ", has been ignored.")
                Exit Sub
            End If
notme:
            'Move on to next SN :)
        Next index&
        'Close the processed memory thread
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub ChatAllowUser(strUser As String, Optional blnPartial As Boolean = False, Optional blnTellChat As Boolean = True)
'%% Its Not Case Sensitive :)
'%% Ex:
'%%     Call ChatIgnoreUser("loserdude551")
'
On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    If Form1.AOL9.Value = 0 Then
    Call ClickIcon(FindEjectTab)
    Else
    Call ClickIcon(FindEjectTab9)
    End If
    'make it lcased
    strUser = LCase(strUser)
    
    'Find List
    rList = ChatEjected()
    'Open Thread
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    'this 'OpenProcess' API Call allows for us to read the contents of the
    'listbox in a nonconventional way, since conventional methods dont work :)
    'We will open it through mThread&
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    'If its valid, then commence adding
    If mThread& Then
        'Loop through the list items using my ListCount Procedure
        'We must subtract 1 since its zero-based
        For index& = 0 To ListCount(rList) - 1
            'Get the current screen name -the invalid characters
            ScreenName$ = String$(4, vbNullChar)
            'Get's the screen name from the list through itmHold&
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 28
            'Read the data from itmHold with the ReadProcessMemory API Call,
            'this will enable us to get the actual data via the "non-conventional" way :)
            'heh, trying to explain as simple as possible.. basically its reading the data
            'in a special way into th memory
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            'Now we will read that memory-added data
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            'Subtract invalid characters
            ScreenName$ = String$(16, vbNullChar)
            'Read into memory again
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            'Finally, we've got our SN
            ScreenName$ = LCase(Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1))
            
            'If its who we want, ignore/unignore 'em :)
            If (blnPartial = True And InStr(ScreenName, strUser)) Or (blnPartial = False And ScreenName = strUser) Then
                If LCase(GetUser) = ScreenName Then GoTo notme
                'DblClick SN
                
                Call ListSelect(rList, CInt(index&))
                'Get thier Screen Name
                If Form1.AOL9.Value = 0 Then
                Call ClickIcon(icon_ChatAllow)
                Else
                Call ClickIcon(icon_ChatAllow9)
                End If
                'Close the processed memory thread
                Call CloseHandle(mThread)
                'If blnTellChat = True Then ChatSend (ScreenName & ", has been ignored.")
                Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been removed from ejected list.")
                
                Exit Sub
            End If
notme:
            'Move on to next SN :)
        Next index&
        'Close the processed memory thread
        Call CloseHandle(mThread)
    End If
    If Form1.AOL9.Value = 0 Then
    Call ClickIcon(FindPeopleHereTab)
    Else
    Call ClickIcon(FindPeopleHereTab9)
    End If
End Sub
Public Function icon_RoomClosed()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_RoomClosed = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function

Public Function icon_ChatAllow()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatAllow = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Public Function icon_ChatAllow9()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatAllow9 = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Public Function icon_ChatEject()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatEject = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Public Function icon_ChatEject9()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
icon_ChatEject9 = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
End Function
Public Function FindChat() As Long
'%% I used PAT or JK's API Spy 5.1 "Generate Function to Find Window" feature
'%% Incase you don't know how this works, what it does is look for the
'%% 'Siblings' of a Window, in other words, uses the parent to find every
'%% single object under that parent. In this case, the Parent here is the
'%% actual room, and the Siblings are the Chat InputBox, the Entered Chat Text
'%% The "People Here" list, the "Send" button, etc. If all of the siblings
'%% are present (sibling <> 0&, aka it has a handle) under that one window then
'%% we must have found the room, return it, otherwise return 0&
'
    Dim aolframe As Long, mdiclient As Long, aolchild As Long
    aolframe = FindWindow("aol frame25", vbNullString)
    mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
    aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
    Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long
    FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
    Do While FindOtherWin <> 0
           DoEvents
           Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid2 = FindWindowEx(FindOtherWin, 0&, "richcntlreadonly", vbNullString)
           Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
           Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid6 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
           Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
           Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) And (Winkid7 <> 0) And (Winkid8 <> 0) And (Winkid9 <> 0) Then
                  FindChat = FindOtherWin
                  Exit Function
           End If
           FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
    Loop
    FindChat = 0
    ' example on how to use:
    '
    ' Dim TheWin As Long
    ' TheWin = FindRoom()
    '
    ' If TheWin <> 0 Then
        ' What to do if window is there
    ' End If
      End Function
Public Sub ProfileEdit()
Dim strName As String, strLoc As String, sglSex As Single, strMStatus As String, _
    strhobbies As String, strComps As String, strOccu As String, strquote As String, blnIncLinkToHomePage As Boolean
strName = "Mike"
strLoc = "FL"
strComps = "PC"
strOccu = "Stuff"
strquote = "dfsadfasdf"
blnIncLinkToHomePage = False
strhobbies = "liar"
sglSex = 1
strMStatus = "sdfs"
'%% Edits/Sets a user's profile
'%% sglSex has 3 values --> 0=Male, 1=Female, 2=No Response
'%% blnIncLinkToHomePage has 2 --> True/False for including a link to users homepage
'%% Ex:
'%%     Call ProfileEdit("somedude", "USA", 2, "STAT", "I do stuff", "Dell", "hitman", "go to: http://magikweb.cjb.net :)", False)
'
    Dim lngMDir As Long, lngProf As Long, lngname As Long, lngLoc As Long, lngmstatus As Long
    Dim lngComp As Long, lngOccu As Long, lngquote As Long, lnghobbies As Long, lngProceed As Long
    Dim oMale As Long, oFemale As Long, oNoRes As Long
    Dim lngCheckBox As Long
    
    Dim aolframe As Long, aolmodal As Long, aolicon As Long, aolcheckbox As Long
    
    'goto profile
    Call Keyword7("profile")
    
    'find member dir
    Do: DoEvents
        lngMDir = FindMemberDir
    Loop Until lngMDir <> 0&
    
    'click my profile to edit profile
    Call ClickIcon(FindMyProfile)
    
    'find profile window
    Do
        DoEvents
        lngProf = FindProfileEdit
    Loop Until lngProf <> 0&
    
    'we can close the member directory
    Call windowclose(lngMDir)
    
    Pause 0.5
    
    '[kids privacy]
    aolframe = FindWindow("aol frame25", vbNullString)
    aolmodal = FindWindow("_aol_modal", vbNullString)
    aolicon = FindWindowEx(aolmodal, 0&, "_aol_icon", vbNullString)
    If GetText(aolicon) = "Kids' Privacy" Then
        aolicon = FindWindowEx(aolmodal, aolicon, "_aol_icon", vbNullString)
    End If
    Pause 10
    'click profile button
    Call ProfileClickQuote
    Pause 5
    'set quote text
    Call ProfileSetQuote("im in PR <a href=aol://2719:2-2-" & GetText(FindChat) & ">" & GetText(FindChat))
    
    'click save profile
    Call ProfileSave
    'click ok on msg box
    Dim x As Long, button As Long
    Do: DoEvents
        x = FindWindow("#32770", vbNullString)
        button = FindWindowEx(x, 0&, "button", vbNullString)
    Loop Until x <> 0& And button <> 0&
    Call PostMessage(button, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(button, WM_KEYUP, VK_SPACE, 0&)
    
    'check again for kids privacy
    '[kids privacy]
    aolframe = FindWindow("aol frame25", vbNullString)
    aolmodal = FindWindow("_aol_modal", vbNullString)
    aolicon = FindWindowEx(aolmodal, 0&, "_aol_icon", vbNullString)
    If GetText(aolicon) = "Kids' Privacy" Then
        aolicon = FindWindowEx(aolmodal, aolicon, "_aol_icon", vbNullString)
    End If
    If GetText(aolicon) = "OK" Then
        'find the checkbox
        'aolcheckbox = FindWindowEx(aolmodal, 0&, "_aol_checkbox", vbNullString)
        'Call check(aolcheckbox, True)
        ClickIcon (aolicon)
    End If
    '[kids privacy]
End Sub

Public Function FindMemberDir() As Long
'%% I used PAT or JK's API Spy 5.1 "Generate Function to Find Window" feature
'%% Finds the Member Directory window based on its siblings
'%% Works just like FindChat, FindIM, FindWelcomeWin, etc.
'
    Dim aolframe As Long, mdiclient As Long, aolchild As Long
    aolframe = FindWindow("aol frame25", vbNullString)
    mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
    aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
    Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long
    FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
    Do While FindOtherWin <> 0
           DoEvents
           Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid2 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
           Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid6 = FindWindowEx(FindOtherWin, 0&, "_aol_edit", vbNullString)
           Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_edit", vbNullString)
           Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) And (Winkid7 <> 0) And (Winkid8 <> 0) And (Winkid9 <> 0) Then
                  FindMemberDir = FindOtherWin
                  Exit Function
           End If
           FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
    Loop
    FindMemberDir = 0
    ' example on how to use:
    '
    ' Dim TheWin As Long
    ' TheWin = FindMemberDir()
    ' If TheWin <> 0 Then
        ' What to do if window is there
    ' End If
End Function

Public Function FindProfileEdit() As Long
'%% I used PAT or JK's API Spy 5.1 "Generate Function to Find Window" feature
'%% Finds the Profile Edit window based on its siblings
'%% Works just like FindChat, FindIM, FindWelcomeWin, etc.
'
    Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindProfileEdit = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function


Public Sub SetText(hWndToSet As Long, TextToSet As String)
'%% Sets text of a window, can be used to set Captions, Chat Boxes, etc.
    'clear text
    Call SendMessageByString(hWndToSet, WM_SETTEXT, 0&, "")
    Call SendMessageByString(hWndToSet, WM_SETTEXT, 0&, TextToSet)
End Sub

Public Function FindMyProfile()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindMyProfile = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
End Function
Public Function ProfileClickQuote()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
ClickIcon (aolicon)
End Function
Public Function ProfileQuote()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
ProfileQuote = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Public Sub ProfileSetQuote(Msg As String)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Msg)
End Sub
Public Function ProfileSave()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
ClickIcon (aolicon)
End Function
Public Function RandomPR()
StartOver:
Select Case randomnumber(6)

Case 1
If LCase(TrimSpaces(GetText(FindChat))) = "hello" Then GoTo StartOver
Call EnterPR("hello")

Case 2
If LCase(TrimSpaces(GetText(FindChat))) = "hi" Then GoTo StartOver
Call EnterPR("hi")

Case 3
If LCase(TrimSpaces(GetText(FindChat))) = "hey" Then GoTo StartOver
Call EnterPR("hey")

Case 4
If LCase(TrimSpaces(GetText(FindChat))) = "vb" Then GoTo StartOver
Call EnterPR("vb")

Case 5
If LCase(TrimSpaces(GetText(FindChat))) = "killcade" Then GoTo StartOver
Call EnterPR("kill cade")

Case 6
If LCase(TrimSpaces(GetText(FindChat))) = "hi" Then GoTo StartOver
Call EnterPR("hi")

Case Else
Call EnterPR("damn hosts")

End Select
End Function

Public Function ChatCheckIfOwner()
Dim i As Integer
'taken from errorgods bas
Dim CloseRoomButton As Long
CloseRoomButton = FindWindowEx(FindRoom, 0&, "_AOL_Icon", vbNullString)
For i = 0 To 15
CloseRoomButton = FindWindowEx(FindRoom, CloseRoomButton, "_AOL_Icon", vbNullString)
Next
If IsWindowVisible(CloseRoomButton) = 1 Then
ChatCheckIfOwner = True
Else
ChatCheckIfOwner = False
End If
End Function

Public Function find_RoomClosed() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe As Long, aolmodal As Long
aolframe = FindWindow("aol frame25", vbNullString)
aolmodal = FindWindow("_aol_modal", vbNullString)
Dim Winkid1 As Long, Winkid2 As Long, FindOtherWin As Long
FindOtherWin = GetWindow(aolmodal, GW_HWNDFIRST)
Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       If (Winkid1 <> 0) And (Winkid2 <> 0) Then
              find_RoomClosed = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
find_RoomClosed = 0
' example on how to use:
' Dim TheWin As Long
' TheWin = find_aolmodal()
' If TheWin <> 0 Then
' What to do if window is there
' End If
End Function
Public Function CloseChatRoom()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call SendMessageLong(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(aolicon, WM_LBUTTONUP, 0&, 0&)
End Function
