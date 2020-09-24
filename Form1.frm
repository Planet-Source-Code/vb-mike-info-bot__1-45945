VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9F4A5A01-0B99-4B22-805B-E357924F08B3}#1.0#0"; "scanchat.ocx"
Object = "{A6A92A0D-7CA7-4B0F-ACAF-DEBF9D1F0BD8}#1.0#0"; "aimchatscan.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AOL Info Bot"
   ClientHeight    =   5280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   9240
      Top             =   2640
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   3600
      TabIndex        =   61
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   57
      Top             =   5760
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   255
      Left            =   6360
      TabIndex        =   56
      Top             =   6120
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   3600
      TabIndex        =   54
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   49
      Top             =   2760
      Width           =   3375
      Begin VB.CheckBox AOL9 
         Caption         =   "AOL9"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblTrigger 
         AutoSize        =   -1  'True
         Caption         =   " "
         Height          =   195
         Left            =   1200
         TabIndex        =   63
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Trigger set by:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblAscii 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   59
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label11 
         Caption         =   "Ascii set by:"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   53
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Color2 set by:"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   960
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1200
         TabIndex        =   51
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Color1 set by:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">O p t  i o n s>"
      Height          =   4815
      Left            =   9840
      TabIndex        =   48
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4560
      TabIndex        =   47
      Top             =   5880
      Width           =   1455
   End
   Begin aimchatscan.gin_aim_chat aim_chat2 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   2619
      _ExtentY        =   2117
   End
   Begin aimchatscan.gin_aim_chat aim_chat1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   2619
      _ExtentY        =   2117
   End
   Begin nitesChatScan.NiteScan Chat2 
      Height          =   255
      Left            =   8880
      TabIndex        =   45
      Top             =   5400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8520
      TabIndex        =   44
      Top             =   5880
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar Bar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   5025
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3545
            Text            =   "SN's: "
            TextSave        =   "SN's: "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3545
            Text            =   "AIM's: "
            TextSave        =   "AIM's: "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3545
            Text            =   "Voiced: "
            TextSave        =   "Voiced: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3545
            Text            =   "Oped: "
            TextSave        =   "Oped: "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3545
            Text            =   "Room Enter: "
            TextSave        =   "Room Enter: "
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Commands"
      Height          =   4815
      Left            =   10200
      TabIndex        =   17
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   ".eject"
         Height          =   195
         Index           =   25
         Left            =   1560
         TabIndex        =   46
         Top             =   2880
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".french"
         Height          =   195
         Index           =   24
         Left            =   1560
         TabIndex        =   43
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".german"
         Height          =   195
         Index           =   23
         Left            =   1560
         TabIndex        =   42
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".translate"
         Height          =   195
         Index           =   22
         Left            =   1560
         TabIndex        =   40
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".remove"
         Height          =   195
         Index           =   21
         Left            =   1560
         TabIndex        =   39
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".name"
         Height          =   195
         Index           =   20
         Left            =   1560
         TabIndex        =   38
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".link"
         Height          =   195
         Index           =   19
         Left            =   1560
         TabIndex        =   37
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".banlist"
         Height          =   195
         Index           =   18
         Left            =   1560
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".clearquotes"
         Height          =   195
         Index           =   17
         Left            =   1560
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".clearban"
         Height          =   195
         Index           =   16
         Left            =   1560
         TabIndex        =   34
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "add/rem AIM"
         Height          =   195
         Index           =   15
         Left            =   1560
         TabIndex        =   33
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".horoscope"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".synonyms .antonyms"
         Height          =   435
         Index           =   13
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".definition"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   30
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".weather"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "add/rem SN"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".website"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".addquote"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "all search's"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".ban/.unban"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "msgall"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".msg"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".op/.deop"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".voice .devoice"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".del"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   ".handle"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8400
      Top             =   4920
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8160
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
   Begin nitesChatScan.NiteScan Chat1 
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   7680
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin MSComctlLib.ListView lst 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4471
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Handle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SN's"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AIM's"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Locked"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Seen"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Enter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Voiced"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "OP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "URL"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "MSG1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "MSG2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "MSG3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "MSG4"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "MSG5"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "<><"
      Height          =   195
      Left            =   4440
      TabIndex        =   55
      Top             =   2640
      Width           =   270
   End
   Begin VB.Label lblAIMs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   5160
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AIM's:"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblOp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   6960
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Oped:"
      Height          =   195
      Left            =   6480
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblVoiced 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced:"
      Height          =   195
      Left            =   5520
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quotes:"
      Height          =   195
      Left            =   6240
      TabIndex        =   8
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label lblEnter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8280
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RoomEnter:"
      Height          =   195
      Left            =   7320
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Banned:"
      Height          =   195
      Left            =   8400
      TabIndex        =   4
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SN's: "
      Height          =   195
      Left            =   3720
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblSNs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   4200
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuChatSplit 
         Caption         =   "Chat Split"
      End
      Begin VB.Menu mnuDialer 
         Caption         =   "Dialer"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Info"
      End
      Begin VB.Menu mnuAscii 
         Caption         =   "Set ascii"
      End
      Begin VB.Menu mnuSetDialer 
         Caption         =   "Set Dialer Login"
      End
      Begin VB.Menu mnuStatusMSG 
         Caption         =   "Set Status Message"
      End
   End
   Begin VB.Menu mnuLST 
      Caption         =   "Lst"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusMSG As String, BotStart As Date
Dim Ascii As String, Trigger As String

Private Sub aim_chat2_aimchatscan(ScreenName As String, aim_whatsaid As String, aim_datesaid As String)
Dim SNs As String, SNs2 As String, Voiced As Integer, Op As Integer, counter As Integer
Dim sn As String, SN2 As String, msgs As Integer, Temp As String
Dim FROM As String, AIM As Integer

On Error Resume Next:
Chat = LTrim(RTrim(Chat))

For g = 0 To List1.ListCount
If TrimSpaces(LCase(List1.List(g))) = TrimSpaces(LCase(ScreenName)) Then Exit Sub
Next g
  
  For x = 1 To lst.ListItems.Count
  If InStr(1, "," & lst.ListItems(x).SubItems(3), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
  For d = 0 To List1.ListCount - 1
  If LCase(List1.List(d)) = LCase(lst.ListItems(x)) Then Exit Sub
  Next d
  End If
  
  
  Next x
   Dim cSend As String
   Dim cData As String
    Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If InStr(Chat, Trigger & "") = 1& Then
      lngSpace& = InStr(Chat, " ")
      If lngSpace& = 0& Then
         strCommand$ = Chat
      Else
         strCommand$ = Left(Chat, lngSpace& - 1&)
      strCommand$ = Mid(strCommand$, 2, Len(strCommand$))
      End If
   End If
snth4$ = ""
pwth4$ = ""
pw4$ = ""
sn4$ = ""
pwthn$ = ""
snthn$ = ""
snn$ = ""
pwn$ = ""
      Select Case LCase(strCommand$)
      
Case "remove"
    srv$ = (Mid(Chat, Len("remove") + 3, Len(Chat) - Len("remove") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(21).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remove""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    AIM_chatsend (lst.ListItems(n) & ", removed by " & ScreenName)
    Set lst.SelectedItem = lst.ListItems(CLng(n))
    lst.ListItems.Remove (lst.SelectedItem.index)
    Exit Sub
    End If
    Next n
    End If
    Next x
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)
    
    

Case "translate"
srv$ = (Mid(Chat, Len("translate") + 3, Len(Chat) - Len("translate") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "translate""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Translate(Left(srv$, InStr(1, srv$, ",", vbTextCompare) - 1), Mid(srv$, InStr(1, srv$, ",", vbTextCompare) + 1, Len(srv$) - InStr(1, srv$, ",", vbTextCompare)))): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)





Case "lock"
srv$ = (Mid(Chat, Len("lock") + 3, Len(Chat) - Len("lock") + 3))
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
        If Len(lst.ListItems(x).SubItems(7)) > 2 Then
        For n = 1 To lst.ListItems.Count
        If LCase(srv$) = LCase(lst.ListItems(n)) Then
        lst.ListItems(n).SubItems(3) = "Yes"
        Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] is now locked.", Ascii): Exit Sub
        End If
        
        Next n
        Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & srv$ & "" & color1 & "</u>] is an invalid handle.", Ascii): Exit Sub
        End If
        
    End If
Next x



Case "unlock"
srv$ = (Mid(Chat, Len("unlock") + 3, Len(Chat) - Len("unlock") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
        If Len(lst.ListItems(x).SubItems(7)) > 2 Then
        For n = 1 To lst.ListItems.Count
        If LCase(srv$) = LCase(lst.ListItems(n)) And ScreenName = GetUser Then GoTo home
        
        If LCase(srv$) = LCase(lst.ListItems(n)) And Len(lst.ListItems(n).SubItems(7)) > 2 Then
        Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] Op's can't be unlocked.", Ascii): Exit Sub
        End If
        If LCase(srv$) = LCase(lst.ListItems(n)) And Len(lst.ListItems(n).SubItems(7)) = 2 Then
home:
        lst.ListItems(n).SubItems(3) = "No"
        Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] is no longer locked.", Ascii): Exit Sub
        End If
        
        Next n
        Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & srv$ & "" & color1 & "</u>] is an invalid handle.", Ascii): Exit Sub
        End If
        
    End If
    Next x




Case "color1"
    srv$ = (Mid(Chat, Len("color1") + 3, Len(Chat) - Len("color1") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) <> 6 Then Exit Sub
    color1 = "<font color=""#" & srv$ & """>"
    lblColor1 = lst.ListItems(x)
    Call AIM_chatsend("" & color1 & "[" & color2 & "Color1 set" & color1 & "]", Ascii): Exit Sub
    End If
    Next x
    
    

    
    
Case "color2"
    srv$ = (Mid(Chat, Len("color2") + 3, Len(Chat) - Len("color2") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    color2 = "<font color=""#" & srv$ & """>"
    lblColor2 = lst.ListItems(x)
    Call AIM_chatsend("" & color1 & "[" & color2 & "Color2 set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
     
     


Case "ascii"
    srv$ = (Mid(Chat, Len("ascii") + 3, Len(Chat) - Len("ascii") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) <> 6 Then Exit Sub
    Ascii = srv$
    lblAscii = lst.ListItems(x)
    Call AIM_chatsend("" & color1 & "[" & color2 & "Ascii set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
 
 

Case "aimchat"
Dim lstPR As ComboBox
srv$ = (Mid(Chat, Len("aimchat") + 3, Len(Chat) - Len("aimchat") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
Call GoChat(srv$)
Do
DoEvents
Loop Until LCase(TrimSpaces(srv$)) = LCase(TrimSpaces(GetText(FindChat)))
Combo1.Clear
Call AddAOL8ListToList(ChatPeopleHereList, Combo1, False)
For n = 0 To Combo1.ListCount - 1
If InStr(1, LCase(TrimSpaces(Combo1.List(n))), "host", vbTextCompare) <> 0 Then
Call RandomPR
End If
Next n
Call AIM_chatsend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b>" & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b>Entered: <u>" & color2 & "" & GetText(FindChat) & "</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
Combo1.Clear
End If
Next x




Case "spanish"
srv$ = (Mid(Chat, Len("spanish") + 3, Len(Chat) - Len("spanish") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "spanish""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Translate("en|es", srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "french"
srv$ = (Mid(Chat, Len("french") + 3, Len(Chat) - Len("french") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "french""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Translate("en|fr", srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "german"
srv$ = (Mid(Chat, Len("german") + 3, Len(Chat) - Len("german") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "german""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Translate("en|de", srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "add<><"
    srv$ = (Mid(Chat, Len("add<><") + 3, Len(Chat) - Len("add<><") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) < 9 Then Exit Sub
    If InStr(1, srv$, ":", vbTextCompare) = 0 Then Call AIM_chatsend("" & color1 & "[" & color2 & "Invalid <><" & color1 & "]", Ascii): Exit Sub
    List3.AddItem srv$
    Call AIM_chatsend("" & color1 & "[" & color2 & srv$ & color1 & "] Saved to <>< tank.", Ascii): Exit Sub
    End If
    Next x
    
    Case "dead<><"
    srv$ = (Mid(Chat, Len("dead<><") + 3, Len(Chat) - Len("dead<><") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    For n = 0 To List3.ListCount - 1
    If LCase(List3.List(n)) = LCase(srv$) Then List3.RemoveItem (n)
    Next n
    Call AIM_chatsend("" & color1 & "[" & color2 & srv$ & color1 & "] removed from <>< tank.", Ascii): Exit Sub
    End If
    Next x
    

End Select
'=======================================================
If LCase(Chat) = Trigger & "defaultascii" Then
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    lblAscii = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Ascii set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
End If


If Chat = Trigger & "<><" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    Call AIM_chatsend("" & color1 & "[" & color2 & List3.List(randomnumber(List3.ListCount) - 1) & color1 & "] ( <> . . <> )", Ascii): Exit Sub
    End If
    Next x
End If

    
If Chat = Trigger & "colors" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    Call AIM_chatsend("" & color1 & "[" & color2 & "Ascii set by: " & lblAscii & color1 & "]", Ascii)
    Call AIM_chatsend("" & color1 & "[" & color2 & "Color1: " & lblColor1 & color1 & "]", Ascii)
    Call AIM_chatsend("" & color1 & "[" & color2 & "Color2: " & lblColor2 & color1 & "]", Ascii): Exit Sub
    End If
    Next x
End If


If LCase(Chat) = Trigger & "wordoftheday" Then
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend(WordOfTheDay, Ascii): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
End If



If LCase(Chat) = Trigger & "uptime" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("[" & DateDiffEx(BotStart, Date & " " & Time) & "]")
    End If
    Next x
End If

End Sub

Private Sub Chat1_ChatScan(ScreenName As String, Chat As String)
Dim SNs As String, SNs2 As String, Voiced As Integer, Op As Integer, counter As Integer
Dim sn As String, SN2 As String, msgs As Integer, Temp As String
Dim FROM As String, AIM As Integer
For d = 1 To lst.ListItems.Count
SNs = SNs & lst.ListItems(d).SubItems(1)
SNs2 = SNs2 & lst.ListItems(d).SubItems(2)
If Len(lst.ListItems(d).SubItems(6)) > 2 Then Voiced = Voiced + 1
If Len(lst.ListItems(d).SubItems(7)) > 2 Then Op = Op + 1
Next d
lblSNs = CountCharAppearance(SNs, ",", False)
lblAIMs = CountCharAppearance(SNs2, ",", False)
lblVoiced = Voiced
lblOp = Op
On Error Resume Next:
Chat = LTrim(RTrim(Chat))

For g = 0 To List1.ListCount
If TrimSpaces(LCase(List1.List(g))) = TrimSpaces(LCase(ScreenName)) Then Exit Sub
Next g
  
  For x = 1 To lst.ListItems.Count
  If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
  For d = 0 To List1.ListCount - 1
  If LCase(List1.List(d)) = LCase(lst.ListItems(x)) Then Exit Sub
  Next d
  End If
  
  
  Next x
   Dim cSend As String
   Dim cData As String
    Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If InStr(Chat, Trigger) = 1& Then
      lngSpace& = InStr(Chat, " ")
      If lngSpace& = 0& Then
         strCommand$ = Chat
      Else
         strCommand$ = Left(Chat, lngSpace& - 1&)
      strCommand$ = Mid(strCommand$, 2, Len(strCommand$))
      End If
   End If
snth4$ = ""
pwth4$ = ""
pw4$ = ""
sn4$ = ""
pwthn$ = ""
snthn$ = ""
snn$ = ""
pwn$ = ""
      Select Case LCase(strCommand$)
      
      Case "handle"
      If Check1(0).Value = 0 Then ChatSend ("Sorry but we are not accepting any new members at this time."): Exit Sub
    Dim strNew As String
    strNew = ""
    srv$ = (Mid(Chat, Len("handle") + 3, Len(Chat) - Len("handle") + 3))
    srv$ = TrimSpaces(srv$)
    'check that handle is < 16 char
If Len(srv$) > 16 Then ChatSend (ScreenName & ", Please try to keep you handle 16 characters or less."): Exit Sub

'check if SN is new
For d = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
Call ChatSend("<b>[" & color2 & "</b><u>" & ScreenName & "</u>" & color1 & "]</b> your han<u>dl</u>e is <b>[" & color2 & "<u>" & lst.ListItems(d) & "</u>" & color1 & "]</b>", Ascii)
Exit Sub
End If
Next d

'check if handle is new
    For Z = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(Z)) = LCase(TrimSpaces(srv$)) Then strNew = "No"
    Next Z
    If strNew = "" Then GoTo NewHandle
    
'check if handle is locked
    For x = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(x)) = LCase(TrimSpaces(srv$)) And lst.ListItems(x).SubItems(3) = "Yes" Then
    Call ChatSend("<b>[" & color2 & "<u>Handle is l</b>ock<b>ed</u>" & color1 & "]", Ascii)
    Exit Sub
    End If
    Next x
    
'if handle is not locked then
For n = 1 To lst.ListItems.Count
If lst.ListItems(n) = srv$ Then
lst.ListItems(n).SubItems(1) = lst.ListItems(n).SubItems(1) & TrimSpaces(ScreenName) & "," 'set SN
Call ChatSend("<b>Welcome back  [" & color2 & "<u> " & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)

End If
Next n

Exit Sub
'if handle is new then
NewHandle:
Dim objLvi As MSComctlLib.ListItem: Set objLvi = lst.ListItems.Add()
objLvi.Text = srv$ 'set Handle
objLvi.SubItems(1) = TrimSpaces(ScreenName) & "," 'set SN
objLvi.SubItems(2) = "" 'set AIM
objLvi.SubItems(3) = "No" 'set locked
objLvi.SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat)) 'set seen
objLvi.SubItems(5) = "On" 'set enter
objLvi.SubItems(6) = "No" 'set voice
objLvi.SubItems(7) = "No" 'set Op
objLvi.SubItems(9) = "Welcome New member!" 'set Msg 1
Call ChatSend("<b>[" & color2 & "<U>" & TrimSpaces(srv$) & "</u>" & color1 & "]</b> type <b>.</b>he<u>l</u>p for a list of <b>com</b>mands.", Ascii)

Case "del"
srv$ = (Mid(Chat, Len("del") + 3, Len(Chat) - Len("del") + 3))
If LCase(srv$) = "me" Then
If Check1(1).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "del me""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    Call ChatSend("<b>All info for [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] has been removed.", Ascii)
    lst.ListItems.Remove (x)
    Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If


Case "voice"
    srv$ = (Mid(Chat, Len("voice") + 3, Len(Chat) - Len("voice") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) <= 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(2).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "voice""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If Len(lst.ListItems(n).SubItems(6)) <> 2 Then Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], is already Voiced"): Exit Sub
    lst.ListItems(n).SubItems(6) = ("by " & ScreenName)
    lst.ListItems(n).SubItems(3) = ("Yes")
    Call ChatSend("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], Voiced by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii): Exit Sub
    End If
    Next x
    'Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", ascii)
    
    

Case "devoice"
    srv$ = (Mid(Chat, Len("devoice") + 3, Len(Chat) - Len("devoice") + 3))
    If LCase(srv$) = LCase(GetUser) Then List1.AddItem (ScreenName): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
    If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
    
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If LCase(lst.ListItems(x).SubItems(7)) = "by mikestoolz" And ScreenName <> "MikesTooLz" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(2).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "devoice""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    lst.ListItems(n).SubItems(6) = ("No")
    Call ChatSend("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], DeVoiced by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & srv$ & " ,<u>" & color1 & "] is not a current member.")
    
    
    
    
Case "op"
    srv$ = (Mid(Chat, Len("op") + 3, Len(Chat) - Len("op") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(3).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "op""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If Len(lst.ListItems(n).SubItems(7)) <> 2 Then Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], is already OP'd"): Exit Sub
    lst.ListItems(n).SubItems(7) = ("by " & ScreenName)
    lst.ListItems(n).SubItems(3) = ("Yes")
    Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], OP'd by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]"): Exit Sub
    End If
    Next n
    End If
    Next x
    Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)



Case "deop"
    srv$ = (Mid(Chat, Len("deop") + 3, Len(Chat) - Len("deop") + 3))
    If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
    If LCase(srv$) = LCase(GetUser) Then List1.AddItem (ScreenName): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
 
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    
    
    If Check1(3).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "deop""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If lst.ListItems(n).SubItems(7) = "by MikesTooLz" And ScreenName <> "MikesTooLz" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    
    lst.ListItems(n).SubItems(7) = ("No")
    Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], deOP'd by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]"): Exit Sub
    End If
    Next n
    End If
    Next x
    Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)
    
    
    
    
Case "msg"
Dim handle As String, Msg As String
srv$ = (Mid(Chat, Len("msg") + 3, Len(Chat) - Len("msg") + 3))
    For d = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(4).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & """messages""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

FROM = lst.ListItems(d)
handle = Left(srv$, InStr(1, srv$, ",", vbTextCompare) - 1)
Msg = Right(srv$, Len(srv$) - InStr(1, srv$, ",", vbTextCompare))
Msg = LTrim(RTrim(Msg))
If Len(Msg) > 400 Then Call ChatSend("<b>[" & color2 & "<u>" & FROM & "</u>" & color1 & "] please keep the msgs under 400 characters"): Exit Sub
For x = 1 To lst.ListItems.Count
If LCase(lst.ListItems(x)) = LCase(handle) Then

Select Case ""
Case lst.ListItems(x).SubItems(9)
lst.ListItems(x).SubItems(9) = ("[" & FROM & "] " & Msg)
Call ChatSend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 1.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
lst.ListItems(x).SubItems(10) = ("[" & FROM & "] " & Msg)
Call ChatSend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 2.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(11)
lst.ListItems(x).SubItems(11) = ("[" & FROM & "] " & Msg)
Call ChatSend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 3.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(12)
lst.ListItems(x).SubItems(12) = ("[" & FROM & "] " & Msg)
Call ChatSend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 4.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(13)
lst.ListItems(x).SubItems(13) = ("[" & FROM & "] " & Msg)
Call ChatSend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 5.", Ascii): Exit Sub

Case Else
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "'s</u>" & color1 & "] msg slots are full.", Ascii): Exit Sub
End Select
End If
Next x
End If
Next d




Case "msgall"

srv$ = (Mid(Chat, Len("msgall") + 3, Len(Chat) - Len("msgall") + 3))
For d = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If lst.ListItems(d).SubItems(7) = "No" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
If Check1(5).Value = 0 Then Exit Sub

FROM = lst.ListItems(d)
For x = 1 To lst.ListItems.Count
lst.ListItems(x).SubItems(9) = ("[" & FROM & "] " & srv$)
Next x
Call ChatSend("<b>Msg saved in slot 1 of all members.", Ascii)
End If
Next d


Case "ban"
srv$ = (Mid(Chat, Len("ban") + 3, Len(Chat) - Len("ban") + 3))
If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): Call ChatEjectUser(ScreenName, False): ChatSend (color1 & "[" & color2 & ScreenName & color1 & "] has been banned"): Exit Sub
If LCase(srv$) = LCase(GetUser) Then List1.AddItem (ScreenName): Call ChatEjectUser(ScreenName, False): ChatSend (color1 & "[" & color2 & ScreenName & color1 & "] has been banned"): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend(color1 & "<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(6).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "ban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
        If LCase(lst.ListItems(n)) = LCase(srv$) Or InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(srv$) & ",", vbTextCompare) <> 0 Then
        If lst.ListItems(n).SubItems(6) = "by MikesTooLz" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
        If lst.ListItems(n).SubItems(7) = "by MikesTooLz" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
        End If
    Next n
    List1.AddItem (TrimSpaces(LCase(srv$))): Call list_nodupes2(List1)
    ChatSend (color1 & "[" & color2 & srv$ & color1 & "] is now being blocked.")
    End If
    Next x


Case "unban"
srv$ = (Mid(Chat, Len("unban") + 3, Len(Chat) - Len("unban") + 3))
If LCase(srv$) = "mikestoolz" Then List1.AddItem (TrimSpaces(LCase(ScreenName))): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
If LCase(srv$) = LCase(GetUser) Then List1.AddItem (TrimSpaces(LCase(ScreenName))): Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(6).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "unban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For n = 0 To List1.ListCount - 1
    If LCase(List1.List(n)) = LCase(srv$) Then List1.RemoveItem (n)
    Next n
    ChatSend (color1 & "[" & color2 & srv$ & color1 & "] is no longer being blocked.")
    End If
    Next x
    
    
Case "seen"
srv$ = (Mid(Chat, Len("seen") + 3, Len(Chat) - Len("seen") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If LCase(lst.ListItems(n)) = LCase(srv$) Then

Call ChatSend("" & color1 & "I saw " & srv$ & " in PR <b>" & Right(lst.ListItems(n).SubItems(4), Len(lst.ListItems(n).SubItems(4)) - InStr(1, lst.ListItems(n).SubItems(4), "|", vbTextCompare)) & "</b>[<u>" & color2 & " " & DateDiffEx(Left(lst.ListItems(n).SubItems(4), InStr(1, lst.ListItems(n).SubItems(4), "|", vbTextCompare) - 1), Date & " " & Time) & " ago.</u>" & color1 & "]", Ascii): Exit Sub
End If
Next n
Call ChatSend("" & color1 & "[<u>" & color2 & "" & srv$ & ", is an invalid handle." & "" & color1 & "</u>]")
End If
Next x



Case "whois"
srv$ = (Mid(Chat, Len("whois") + 3, Len(Chat) - Len("whois") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If InStr(1, Replace(lst.ListItems(n).SubItems(1), " ", "", 1, Len(lst.ListItems(n).SubItems(1)), vbTextCompare), Replace(srv$, " ", "", 1, Len(srv$)) & ",", vbTextCompare) <> 0 Then
Call ChatSend("" & color1 & "[<u>" & color2 & "" & srv$ & "</u>" & color1 & "] is [<u>" & color2 & "" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii): Exit Sub
End If
Next n
ChatSend (srv$ & ", is not a member.")
End If
Next x



Case "info"

srv$ = (Mid(Chat, Len("info") + 3, Len(Chat) - Len("info") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n), Replace(srv$, " ", "", 1, Len(srv$)), vbTextCompare) <> 0 Then
For d = 1 To lst.ListItems.Count
If LCase(lst.ListItems(d)) = LCase(srv$) Then
SNs = lst.ListItems(d).SubItems(1)


    If lst.ListItems(n).SubItems(8) = "" Then
    Call ChatSend("" & color1 & "[SNs: " & color2 & "" & CountCharAppearance(SNs, ",", False) & "" & color1 & "][AIMs: " & color2 & "" & CountCharAppearance(lst.ListItems(d).SubItems(2), ",", False) & "" & color1 & "][Locked: " & color2 & "" & lst.ListItems(d).SubItems(3) & "" & color1 & "][Voiced: " & color2 & "" & lst.ListItems(d).SubItems(6) & "" & color1 & "]" & "[Oped: " & color2 & "" & lst.ListItems(d).SubItems(7) & "" & color1 & "]", Ascii): Exit Sub
    Else
    Call ChatSend("" & color1 & "[SNs: " & color2 & "" & CountCharAppearance(SNs, ",", False) & "" & color1 & "][AIMs: " & color2 & "" & CountCharAppearance(lst.ListItems(d).SubItems(2), ",", False) & "" & color1 & "][Locked: " & color2 & "" & lst.ListItems(d).SubItems(3) & "" & color1 & "][Voiced: " & color2 & "" & lst.ListItems(d).SubItems(6) & "" & color1 & "]" & "[Oped: " & color2 & "" & lst.ListItems(d).SubItems(7) & "" & color1 & "]" & "[<a href=""" & lst.ListItems(d).SubItems(8) & """>WebSite</a>]", Ascii): Exit Sub
    End If
End If
Next d
End If
Next n
Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is an invalid handle.", Ascii)
End If
Next x



Case "read"
srv$ = (Mid(Chat, Len("read") + 3, Len(Chat) - Len("read") + 3))
If LCase(srv$) = LCase("msg") Then
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
Select Case ""
Case lst.ListItems(x).SubItems(9)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have no me<u>ss</u>ages.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 1 me<u>ss</u>age", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
lst.ListItems(x).SubItems(9) = "": Exit Sub
Case lst.ListItems(x).SubItems(11)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 2 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = "": Exit Sub
Case lst.ListItems(x).SubItems(12)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 3 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
Case lst.ListItems(x).SubItems(13)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 4 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
Case Else
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 5 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(13), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
lst.ListItems(x).SubItems(13) = ""
Exit Sub
End Select
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If



Case "enter"
    srv$ = (Mid(Chat, Len("enter") + 3, Len(Chat) - Len("enter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If LCase(srv$) = "on" Then lst.ListItems(x).SubItems(5) = "On": Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] your chat enter is now <u>on</u>.", Ascii): Exit Sub
    If LCase(srv$) = "off" Then lst.ListItems(x).SubItems(5) = "Off": Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] your chat enter is now <u>off</u>.", Ascii): Exit Sub
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
    
    
Case "roomenter"
    srv$ = (Mid(Chat, Len("roomenter") + 3, Len(Chat) - Len("roomenter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If LCase(srv$) = "on" Then lblEnter = "On": Call ChatSend("<b>""Room Enter"" is now on.", Ascii): Exit Sub
    If LCase(srv$) = "off" Then lblEnter = "Off": Call ChatSend("<b>""Room Enter"" is now off.", Ascii): Exit Sub
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii)
    

Case "link"
    If Check1(19).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "link""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("link") + 3, Len(Chat) - Len("link") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""" & srv$ & """>" & srv$ & "</a>")
    End If
    Next x
    
    
    
Case "google"
If Check1(7).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "google""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("google") + 3, Len(Chat) - Len("google") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""http://www.google.com/search?hl=en&q=" & srv$ & """>" & lst.ListItems(x) & ", you were just googled.</a>")
    End If
    Next x



Case "yahoo"
If Check1(7).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "yahoo""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("Yahoo") + 3, Len(Chat) - Len("yahoo") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""http://search.yahoo.com/bin/search?p=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x

Case "altavista"
If Check1(7).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "altavista""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("altavista") + 3, Len(Chat) - Len("altavista") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""http://www.altavista.com/web/results?q=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x



Case "ebay"
If Check1(7).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "ebay""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("ebay") + 3, Len(Chat) - Len("ebay") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""http://search.ebay.com/search/search.dll?cgiurl=http%3A%2F%2Fcgi.ebay.com%2Fws%2F&krd=1&from=R8&MfcISAPICommand=GetResult&ht=1&SortProperty=MetaEndSort&query=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x
    
    
    
    Case "askjeeves"
    If Check1(7).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "askjeeves""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("askjeeves") + 3, Len(Chat) - Len("askjeeves") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("<a href=""http://web.ask.com/web?q=" & srv$ & """>" & lst.ListItems(x) & ", here's what Jeeves has to say.</a>")
    End If
    Next x
    
    
    
Case "addquote"
    srv$ = (Mid(Chat, Len("addquote") + 3, Len(Chat) - Len("addquote") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Check1(8).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addquote""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List2.AddItem ("[" & lst.ListItems(x) & "] " & srv$)
    Call list_nodupes2(List2)
    ChatSend ("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] quote added.")
    Exit Sub
    End If
    Next x
    
    
    
    
Case "allenter"
    srv$ = (Mid(Chat, Len("allenter") + 3, Len(Chat) - Len("allenter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub  'exit if there not @
    
    For n = 1 To lst.ListItems.Count
    If LCase(srv$) = "on" Then lst.ListItems(n).SubItems(5) = "On"
    If LCase(srv$) = "off" Then lst.ListItems(n).SubItems(5) = "Off"
    Next n
    If LCase(srv$) = "on" Then Call ChatSend("<b>All members ""Room Enter"" had been set to [on].", Ascii)
    If LCase(srv$) = "off" Then Call ChatSend("<b>All members ""Room Enter"" had been set to [off].", Ascii)
    Exit Sub
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
    
    
    
Case "website"
    srv$ = (Mid(Chat, Len("website") + 3, Len(Chat) - Len("website") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Check1(9).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "website""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    lst.ListItems(x).SubItems(8) = srv$
    ChatSend ("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] website saved.")
    End If
    Next x
    
    
    
Case "addsn"
srv$ = (Mid(Chat, Len("addsn") + 3, Len(Chat) - Len("addsn") + 3))
'find illegal letters
If InStr(1, srv$, "<", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ">", vbTextCompare) <> 0 Then Call ChatSend(ScreenName & ", Invalid SN"): Exit Sub
If InStr(1, srv$, "/", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "`", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "~", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "@", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "#", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "^", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "&", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "*", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "(", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ")", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "_", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "+", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "=", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "-", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "[", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "]", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "{", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "}", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "\", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "|", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "?", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ".", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ",", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "'", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, """", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ";", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ":", vbTextCompare) <> 0 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
'check that sn is < 16 char
If Len(srv$) > 16 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Please keep your handle 16 characters or less.", Ascii): Exit Sub

For x = 1 To lst.ListItems.Count
If InStr(1, "," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
Call ChatSend(srv$ & ", is " & lst.ListItems(x) & "'s SN.", Ascii): Exit Sub
End If
If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
Call ChatSend(srv$ & ", is " & lst.ListItems(x) & "'s AIM.", Ascii): Exit Sub
End If
Next x
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(10).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addsn""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
lst.ListItems(n).SubItems(1) = lst.ListItems(n).SubItems(1) & TrimSpaces(srv$) & ","
Call ChatSend("<b>[" & color2 & "<u>" & srv$ & "</u>" & color1 & "] added to handle [" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)
End If
Next n



Case "pr"
Dim lstPR As ComboBox
srv$ = (Mid(Chat, Len("pr") + 3, Len(Chat) - Len("pr") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
Call EnterPR(srv$)
Do
DoEvents
Loop Until LCase(TrimSpaces(srv$)) = LCase(TrimSpaces(GetText(FindChat)))
Combo1.Clear
Call AddAOL8ListToList(ChatPeopleHereList, Combo1, False)
For n = 0 To Combo1.ListCount - 1
If InStr(1, LCase(TrimSpaces(Combo1.List(n))), "host", vbTextCompare) <> 0 Then
Call RandomPR
End If
Next n
Call ChatSend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b> v" & App.Major & "." & App.Minor & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b>Entered: <u>" & color2 & "" & GetText(FindChat) & "</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
Combo1.Clear
End If
Next x


Case "weather"
srv$ = (Mid(Chat, Len("weather") + 3, Len(Chat) - Len("weather") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(11).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "weather""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
GetWeatherInfo (srv$)

If WeatherInfo.State = "m -" Then Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "]Invalid zipcode", Ascii): Exit Sub
If WeatherInfo.CurrentCond = " " Then Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "]Invalid zipcode", Ascii): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "]Invalid zipcode", Ascii): Exit Sub
ChatSend ("" & color2 & "<b>" & WeatherInfo.City & "," & WeatherInfo.State & "</b> " & color1 & "- [Condition: " & color2 & WeatherInfo.CurrentCond & color1 & "] [Temp: " & color2 & "" & WeatherInfo.CurrentF & "º" & WeatherInfo.FeelsLike & "" & color1 & "] [UVIndex: " & color2 & "" & WeatherInfo.UVIndex & "" & color1 & "] [Humidity: " & color2 & "" & WeatherInfo.Humidity & "" & color1 & "] [Wind: " & color2 & "" & WeatherInfo.Wind & "" & color1 & "] [Visibility: " & color2 & "" & WeatherInfo.Visibility & "" & color1 & "]"): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "definition"
srv$ = (Mid(Chat, Len("definition") + 3, Len(Chat) - Len("definition") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(12).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "definition""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

Call ChatSend(Definition(srv$), Ascii): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)


Case "synonyms"
srv$ = (Mid(Chat, Len("synonyms") + 3, Len(Chat) - Len("synonyms") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(13).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "synonyms""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Synonyms(srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "antonyms"
srv$ = (Mid(Chat, Len("antonyms") + 3, Len(Chat) - Len("antonyms") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(13).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "antonyms""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Antonyms(srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "newhandle"
srv$ = (Mid(Chat, Len("newhandle") + 3, Len(Chat) - Len("newhandle") + 3))
srv$ = TrimSpaces(srv$)
'check that handle is < 16 char
If Len(srv$) > 16 Then ChatSend (ScreenName & ", Please try to keep you handle 16 characters or less."): Exit Sub
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
'check if handle is new
    For Z = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(Z)) = LCase(TrimSpaces(srv$)) Then strNew = "No"
    Next Z
    If strNew = "" Then
    ChatSend (lst.ListItems(x) & "'s handle was changed to <b>" & srv$)
    lst.ListItems(x).Text = srv$
    Else
    ChatSend ("A member is already using the handle <b>" & srv$ & "</b>")
    End If
End If
Next x


Case "sn"
    srv$ = (Mid(Chat, Len("sn") + 3, Len(Chat) - Len("sn") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then ChatSend ("[" & Left(lst.ListItems(n).SubItems(1), Len(lst.ListItems(n).SubItems(1)) - 1) & "]"): Exit Sub
    Next n
    ChatSend (color1 & "[" & color2 & ScreenName & color1 & "] there is no member with that handle."): Exit Sub
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
    
    
    
Case "horoscope"
srv$ = (Mid(Chat, Len("horoscope") + 3, Len(Chat) - Len("horoscope") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(14).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "horoscope""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

'If WeatherInfo.State = "m -" Then ChatSend ("Invalid zipcode"): Exit Sub
ChatSend (Horoscope(srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)


Case "remsn"
srv$ = (Mid(Chat, Len("remsn") + 3, Len(Chat) - Len("remsn") + 3))
srv$ = TrimSpaces(srv$)
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(10).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remsn""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

If InStr(1, "," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(srv$) & ",", vbTextCompare) <> 0 Then
Temp = Replace("," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(srv$) & ",", ",", 1, Len("," & LCase(lst.ListItems(x).SubItems(1))), vbTextCompare)
If Left(Temp, 1) = "," Then Temp = Right(Temp, Len(Temp) - 1)
lst.ListItems(x).SubItems(1) = Temp
ChatSend (srv$ & ",was removed"): Exit Sub
End If

End If
Next x



Case "x"
srv$ = (Mid(Chat, Len("x") + 3, Len(Chat) - Len("x") + 3))
If ScreenName = GetUser Then Call ChatIgnoreUser(srv$, True)


Case "unx"
srv$ = (Mid(Chat, Len("unx") + 3, Len(Chat) - Len("unx") + 3))
If ScreenName = GetUser Then Call ChatIgnoreUser(srv$, True, False)



Case "eject"
srv$ = (Mid(Chat, Len("eject") + 3, Len(Chat) - Len("eject") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(25).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "eject""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
If ScreenName = "MikesTooLz" Then Call ChatEjectUser(srv$, True): Exit Sub
If ScreenName = GetUser Then Call ChatEjectUser(srv$, True): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) <> 2 Then Call ChatEjectUser(srv$, True): Exit Sub
End If
Next x


Case "allow"
srv$ = (Mid(Chat, Len("allow") + 3, Len(Chat) - Len("allow") + 3))
If ScreenName = GetUser Then Call ChatAllowUser(srv$, True): Exit Sub
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(25).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "eject""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
If ScreenName = "MikesTooLz" Then Call ChatEjectUser(srv$, True): Exit Sub
If ScreenName = GetUser Then Call ChatEjectUser(srv$, True): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
Call ChatAllowUser(srv$, True): Exit Sub
End If
Next x





Case "addaim"
srv$ = (Mid(Chat, Len("addaim") + 3, Len(Chat) - Len("addaim") + 3))
'find illegal letters
If InStr(1, srv$, "<", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ">", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "/", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "`", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "~", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "@", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "#", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "^", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "&", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "*", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "(", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ")", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "_", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "+", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "=", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "-", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "[", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "]", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "{", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "}", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "\", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "|", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "?", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ".", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ",", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "'", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, """", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ";", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ":", vbTextCompare) <> 0 Then ChatSend (ScreenName & ", Invalid AIM"): Exit Sub
'check that AIM is < 24 char
If Len(srv$) > 24 Then ChatSend (ScreenName & ", Please try to keep you handle 24 characters or less."): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, LCase("," & lst.ListItems(x).SubItems(1)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
    ChatSend (srv$ & ", is " & lst.ListItems(x) & "'s SN."): Exit Sub
    End If
    If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
    ChatSend (srv$ & ", is " & lst.ListItems(x) & "'s AIM."): Exit Sub
    End If
Next x
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(15).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addaim""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
lst.ListItems(n).SubItems(2) = lst.ListItems(n).SubItems(2) & TrimSpaces(srv$) & ","
Call ChatSend("<b>[" & color2 & "<u>" & srv$ & "</u>" & color1 & "] added to handle [" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)
End If
Next n



Case "remaim"
srv$ = (Mid(Chat, Len("remaim") + 3, Len(Chat) - Len("remaim") + 3))
srv$ = TrimSpaces(srv$)
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(15).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remaim""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(srv$) & ",", vbTextCompare) <> 0 Then
Temp = Replace("," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(srv$) & ",", ",", 1, Len("," & LCase(lst.ListItems(x).SubItems(2))), vbTextCompare)
If Left(Temp, 1) = "," Then Temp = Right(Temp, Len(Temp) - 1)
lst.ListItems(x).SubItems(2) = Temp
ChatSend (srv$ & ", was removed"): Exit Sub
End If
End If
Next x



Case "aim"
    srv$ = (Mid(Chat, Len("aim") + 3, Len(Chat) - Len("aim") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If lst.ListItems(n).SubItems(2) = "" Then Call ChatSend("" & color1 & "[" & color2 & "" & lst.ListItems(n) & "" & color1 & "], has no AIMs.", Ascii): Exit Sub
    Call ChatSend("" & color1 & "[" & color2 & "" & Left(lst.ListItems(n).SubItems(2), Len(lst.ListItems(n).SubItems(2)) - 1) & "" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    Call ChatSend("" & color1 & "[" & color2 & "" & ScreenName & "" & color1 & "] there is no member with that handle.", Ascii): Exit Sub
    End If
    Next x
    Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "name"
srv$ = (Mid(Chat, Len("name") + 3, Len(Chat) - Len("name") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(20).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "name""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (WhatNameMeans(srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




    
    
    
    
 
    



     
     
End Select
'========================================================
If LCase(Chat) = Trigger & "lock" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    lst.ListItems(x).SubItems(3) = "Yes"
    Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] is now locked.", Ascii)
    Exit Sub
    End If
    Next x
      Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If



If LCase(Chat) = Trigger & "unlock" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) > 2 Then Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] Op's can't be unlocked. Use " & Trigger & "addsn"" to add a new SN.", Ascii): Exit Sub
    lst.ListItems(x).SubItems(3) = "No"
    Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] is no longer locked.", Ascii)
    Exit Sub
    End If
    Next x
      Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If



If InStr(1, Chat, "has entered the room.", vbTextCompare) <> 0 Then
If ScreenName = "OnlineHost" Then
sn = Left(Chat, InStr(1, Chat, "has entered the room.", vbTextCompare) - 2)

If InStr(1, LCase(TrimSpaces(sn)), "host", vbTextCompare) Then Call RandomPR
'====if there banned then say it
For Z = 1 To lst.ListItems.Count
  If InStr(1, lst.ListItems(Z).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
  SN2 = lst.ListItems(Z)
  For F = 0 To List1.ListCount - 1
  If LCase(List1.List(F)) = LCase(SN2) Then Call ChatSend("" & color1 & "[" & color2 & "" & sn & "" & color1 & "] you are banned. All commands disabled.", Ascii): Call ChatEjectUser(Left(sn, 3), True): Exit Sub
  If LCase(List1.List(F)) = LCase(sn) Then Call ChatSend("" & color1 & "[" & color2 & "" & sn & "" & color1 & "] you are banned. All commands disabled.", Ascii): Call ChatEjectUser(Left(sn, 3), True): Exit Sub
  
  Next F
  End If
Next Z
'======

For x = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
lst.ListItems(x).SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat))
If lst.ListItems(x).SubItems(5) = "Off" Then Exit Sub
If LCase(lblEnter) = "off" Then Exit Sub
Select Case ""
Case lst.ListItems(x).SubItems(9)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have no me<u>ss</u>ages.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 1 me<u>ss</u>age", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
lst.ListItems(x).SubItems(9) = "": Exit Sub
Case lst.ListItems(x).SubItems(11)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 2 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = "": Exit Sub
Case lst.ListItems(x).SubItems(12)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 3 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = "": Exit Sub
Case lst.ListItems(x).SubItems(13)
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 4 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = "": Exit Sub
Case Else
Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 5 me<u>ss</u>age<u>s</u>", Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
Call ChatSend("" & color1 & "" & lst.ListItems(x).SubItems(13), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
lst.ListItems(x).SubItems(13) = "": Exit Sub

End Select
End If
Next x
If Check1(0).Value = 0 Then Exit Sub
If LCase(lblEnter) = "off" Then Exit Sub
Call ChatSend("" & color1 & "" & sn & " type " & Trigger & "handle "" and your handle(name)")
End If
End If


If InStr(1, Chat, "has left the room.", vbTextCompare) <> 0 Then
If ScreenName = "OnlineHost" Then
'===if everyone left get owner
Dim room As String
room = GetText(FindChat)
List4.Clear
Call AddAOLListToListbox(ChatPeopleHereList, List4)
If List4.ListCount = 1 And ChatCheckIfOwner = False Then Call RandomPR: timeout (0.5): Call EnterPR(room)
sn = Left(Chat, InStr(1, Chat, "has left the room.", vbTextCompare) - 2)
For x = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
lst.ListItems(x).SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat))
End If
Next x
End If
End If


If LCase(Chat) = Trigger & "banlist" Then
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(18).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "banlist""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    If List1.ListCount = 0 Then Call ChatSend(ScreenName & ", no one is banned.", Ascii)
    For n = 0 To List1.ListCount - 1
    Call ChatSend("" & color1 & "[" & color2 & "" & n + 1 & "" & color1 & "] " & List1.List(n), Ascii)
    Next n
    End If
    Next x
End If



If LCase(Chat) = Trigger & "quote" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend (List2.List(randomnumber(List2.ListCount) - 1))
    End If
    Next x
End If


If LCase(Chat) = Trigger & "clearban" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(16).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "clearban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List1.Clear
    ChatSend ("Banned list has been cleared."): Exit Sub
    End If
    Next x
End If


If LCase(Chat) = Trigger & "clearquotes" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(17).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "clearquotes""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List2.Clear
    ChatSend ("Quotes list has been cleared."): Exit Sub
    End If
    Next x
End If



If LCase(Chat) = Trigger & "status" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("" & color1 & "Handles:[" & color2 & "" & lst.ListItems.Count & "" & color1 & "] SNs:[" & color2 & "" & lblSNs & "" & color1 & "] AIMs:[" & color2 & "" & lblAIMs & "" & color1 & "] Voiced:[" & color2 & "" & lblVoiced & "" & color1 & "] Oped:[" & color2 & "" & lblOp & "" & color1 & "] Quotes:[" & color2 & "" & List2.ListCount & "" & color1 & "] Banned:[" & color2 & "" & List1.ListCount & "" & color1 & "]" & StatusMSG)
    End If
    Next x
End If




If LCase(Chat) = Trigger & "help" Then
    ChatSend ("[<a href=""http://www.mikestoolz.com/downloads/commands.html"">" & ScreenName & " click here for list of commands</a>]")
End If
    
    
    
    
If LCase(Chat) = Trigger & "oplist" Then
    counter = 0
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If lst.ListItems(x).SubItems(7) = "No" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
    If Len(lst.ListItems(n).SubItems(7)) > 2 Then counter = counter + 1: Call ChatSend("" & color1 & "[" & color2 & "" & counter & "" & color1 & "] " & lst.ListItems(n), Ascii)
    Next n
    End If
    Next x
End If










End Sub







Private Sub chat2_ChatScan(ScreenName As String, Chat As String)
Dim SNs As String, SNs2 As String, Voiced As Integer, Op As Integer, counter As Integer
Dim sn As String, SN2 As String, msgs As Integer, Temp As String
Dim FROM As String, AIM As Integer

On Error Resume Next:
Chat = LTrim(RTrim(Chat))

For g = 0 To List1.ListCount
If TrimSpaces(LCase(List1.List(g))) = TrimSpaces(LCase(ScreenName)) Then Exit Sub
Next g
  
  For x = 1 To lst.ListItems.Count
  If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
  For d = 0 To List1.ListCount - 1
  If LCase(List1.List(d)) = LCase(lst.ListItems(x)) Then Exit Sub
  Next d
  End If
  
  
  Next x
   Dim cSend As String
   Dim cData As String
    Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If InStr(Chat, Trigger) = 1& Then
      lngSpace& = InStr(Chat, " ")
      If lngSpace& = 0& Then
         strCommand$ = Chat
      Else
         strCommand$ = Left(Chat, lngSpace& - 1&)
      strCommand$ = Mid(strCommand$, 2, Len(strCommand$))
      End If
   End If
snth4$ = ""
pwth4$ = ""
pw4$ = ""
sn4$ = ""
pwthn$ = ""
snthn$ = ""
snn$ = ""
pwn$ = ""
      Select Case LCase(strCommand$)
      


Case "remove"
    srv$ = (Mid(Chat, Len("remove") + 3, Len(Chat) - Len("remove") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(21).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remove""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    ChatSend (lst.ListItems(n) & ", removed by " & ScreenName)
    Set lst.SelectedItem = lst.ListItems(CLng(n))
    lst.ListItems.Remove (lst.SelectedItem.index)
    Exit Sub
    End If
    Next n
    End If
    Next x
    Call ChatSend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)
    
    

Case "translate"
srv$ = (Mid(Chat, Len("translate") + 3, Len(Chat) - Len("translate") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "translate""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Translate(Left(srv$, InStr(1, srv$, ",", vbTextCompare) - 1), Mid(srv$, InStr(1, srv$, ",", vbTextCompare) + 1, Len(srv$) - InStr(1, srv$, ",", vbTextCompare)))): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "spanish"
srv$ = (Mid(Chat, Len("spanish") + 3, Len(Chat) - Len("spanish") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "spanish""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Translate("en|es", srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "french"
srv$ = (Mid(Chat, Len("french") + 3, Len(Chat) - Len("french") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "french""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Translate("en|fr", srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "german"
srv$ = (Mid(Chat, Len("german") + 3, Len(Chat) - Len("german") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(22).Value = 0 Then Call ChatSend("<b>Sorry but [<u>" & color2 & "" & Trigger & "german""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
ChatSend (Translate("en|de", srv$)): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "lock"
srv$ = (Mid(Chat, Len("lock") + 3, Len(Chat) - Len("lock") + 3))
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
        If Len(lst.ListItems(x).SubItems(7)) > 2 Then
        For n = 1 To lst.ListItems.Count
        If LCase(srv$) = LCase(lst.ListItems(n)) Then
        lst.ListItems(n).SubItems(3) = "Yes"
        Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] is now locked.", Ascii): Exit Sub
        End If
        
        Next n
        Call ChatSend("" & color1 & "[" & color2 & "<u>" & srv$ & "" & color1 & "</u>] is an invalid handle.", Ascii): Exit Sub
        End If
        
    End If
Next x



Case "unlock"
srv$ = (Mid(Chat, Len("unlock") + 3, Len(Chat) - Len("unlock") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
        If Len(lst.ListItems(x).SubItems(7)) > 2 Then
        For n = 1 To lst.ListItems.Count
        If LCase(srv$) = LCase(lst.ListItems(n)) And ScreenName = GetUser Then GoTo home
        
        If LCase(srv$) = LCase(lst.ListItems(n)) And Len(lst.ListItems(n).SubItems(7)) > 2 Then
        Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] Op's can't be unlocked.", Ascii): Exit Sub
        End If
        If LCase(srv$) = LCase(lst.ListItems(n)) And Len(lst.ListItems(n).SubItems(7)) = 2 Then
home:
        lst.ListItems(n).SubItems(3) = "No"
        Call ChatSend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "" & color1 & "</u>] is no longer locked.", Ascii): Exit Sub
        End If
        
        Next n
        Call ChatSend("" & color1 & "[" & color2 & "<u>" & srv$ & "" & color1 & "</u>] is an invalid handle.", Ascii): Exit Sub
        End If
        
    End If
    Next x



Case "trigger"
    srv$ = (Mid(Chat, Len("trigger") + 3, Len(Chat) - Len("trigger") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) <> 1 Then Exit Sub
    Trigger = srv$
    lblTrigger = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Trigger set" & color1 & "]", Ascii): Exit Sub
    End If
    Next x


Case "color1"
    srv$ = (Mid(Chat, Len("color1") + 3, Len(Chat) - Len("color1") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) <> 6 Then Exit Sub
    color1 = "<font color=""#" & srv$ & """>"
    lblColor1 = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Color1 set" & color1 & "]", Ascii): Exit Sub
    End If
    Next x
    
    
    
Case "color2"
    srv$ = (Mid(Chat, Len("color2") + 3, Len(Chat) - Len("color2") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) <> 6 Then Exit Sub
    color2 = "<font color=""#" & srv$ & """>"
    lblColor2 = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Color2 set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
     
     
Case "ascii"
    srv$ = (Mid(Chat, Len("ascii") + 3, Len(Chat) - Len("ascii") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    Ascii = srv$
    lblAscii = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Ascii set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
     



Case "add<><"
    srv$ = (Mid(Chat, Len("add<><") + 3, Len(Chat) - Len("add<><") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Len(srv$) < 9 Then Exit Sub
    If InStr(1, srv$, ":", vbTextCompare) = 0 Then Call ChatSend("" & color1 & "[" & color2 & "Invalid <><" & color1 & "]", Ascii): Exit Sub
    List3.AddItem srv$
    Call ChatSend("" & color1 & "[" & color2 & srv$ & color1 & "] Saved to <>< tank.", Ascii): Exit Sub
    End If
    Next x
     
     
     
     
    Case "dead<><"
    srv$ = (Mid(Chat, Len("dead<><") + 3, Len(Chat) - Len("dead<><") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    For n = 0 To List3.ListCount - 1
    If LCase(List3.List(n)) = LCase(srv$) Then List3.RemoveItem (n)
    Next n
    Call ChatSend(color1 & "[" & color2 & srv$ & color1 & "] removed from <>< tank.", Ascii): Exit Sub
    End If
    Next x
    
    
    Case "split"
    If ScreenName = GetUser Then
    Dim room As String
    srv$ = (Mid(Chat, Len("split") + 3, Len(Chat) - Len("split") + 3))
    room = GetText(FindChat)
    Call EnterPR(srv$): Call EnterPR(room)
    End If
    
    
End Select
'=======================================================
If LCase(Chat) = Trigger & "vlist" Then
    counter = 0
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If lst.ListItems(x).SubItems(6) = "No" And lst.ListItems(x).SubItems(7) = "No" Then Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
    If Len(lst.ListItems(n).SubItems(6)) > 2 Then counter = counter + 1: Call ChatSend("" & color1 & "[" & color2 & "" & counter & "" & color1 & "] " & lst.ListItems(n), Ascii)
    Next n
    End If
    Next x
End If



If LCase(Chat) = Trigger & "close" Then
If ScreenName = "MikesTooLz" Then
Call ClickIcon(icon_RoomClosed)
timeout (1)
SendKeys ("   ")
End If
End If


If LCase(Chat) = Trigger & "uptime" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    ChatSend ("[" & DateDiffEx(BotStart, Date & " " & Time) & "]")
    End If
    Next x
End If



If LCase(Chat) = Trigger & "deopall" And ScreenName = "MikesTooLz" Then
Dim DeOPcounter As Integer
    For x = 1 To lst.ListItems.Count
    If Len(lst.ListItems(x).SubItems(7)) <> 2 And lst.ListItems(x).SubItems(7) <> "by MikesTooLz" Then lst.ListItems(x).SubItems(7) = "No": DeOPcounter = DeOPcounter + 1
    Next x
Call ChatSend("[" & color2 & DeOPcounter & color1 & "] Liars have been DeOP'd!", Ascii)
End If


If LCase(Chat) = Trigger & "devoiceall" And ScreenName = "MikesTooLz" Then
Dim DeVoicecounter As Integer
    For x = 1 To lst.ListItems.Count
    If Len(lst.ListItems(x).SubItems(6)) <> 2 And lst.ListItems(x).SubItems(6) <> "by MikesTooLz" Then lst.ListItems(x).SubItems(6) = "No": DeVoicecounter = DeVoicecounter + 1
    Next x
Call ChatSend("[" & color2 & DeOPcounter & color1 & "] Liars have been DeVoiced!", Ascii)
End If


If LCase(Chat) = Trigger & "sound" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    Select Case randomnumber(14)
    
    Case 1
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud1")
    Case 2
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud2")
    Case 3
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud3")
    Case 4
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud4")
    Case 5
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud5")
    Case 6
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud6")
    Case 7
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud7")
    Case 8
    ChatSend ("{</h7>s c:\Windows\help\tours\windowsmediaplayer\audio\wav\wmpaud8")
    Case 9
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\media\cagcat10\ELPHRG01")
    Case 10
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\media\cagcat10\J0214098")
    Case 11
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\Office10\Media\CASHREG")
    Case 12
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\Office10\Media\chimes")
    Case 13
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\Office10\Media\drumroll")
    Case 14
    ChatSend ("{</h7>s C:\Program Files\Microsoft Office\Office10\Media\type")
    
    End Select
    End If
    Next x
End If
If LCase(Chat) = Trigger & "defaultascii" Then
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    lblAscii = lst.ListItems(x)
    Call ChatSend("" & color1 & "[" & color2 & "Ascii set " & color1 & "]", Ascii): Exit Sub
    End If
    Next x
End If

If Chat = Trigger & "dialer" And ScreenName = GetUser Then Call GetDialerStats: Call ChatSend(color1 & "<b>dialer stats</b> - [" & "Total min/calls:" & color2 & DialerInfo.TotalMinutes_Calls & color1 & "] [Average:" & color2 & DialerInfo.OverallAverage & color1 & "] [This Month:" & color2 & DialerInfo.ThisMonthMinutes_Calls & color1 & "] [Months Average:" & color2 & DialerInfo.MonthAverage & color1 & "] [Todays Min/Calls:" & color2 & DialerInfo.TodayMinutes_Calls & color1 & "] [Cash:" & color2 & DialerInfo.Cash & color1 & "]", Ascii): Exit Sub

If Chat = Trigger & "<><" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    Call ChatSend("" & color1 & "[" & color2 & List3.List(randomnumber(List3.ListCount) - 1) & color1 & "] ( <> . . <> )", Ascii): Exit Sub
    End If
    Next x
End If

    
    
If Chat = Trigger & "colors" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    Call ChatSend("" & color1 & "[" & color2 & "Ascii set by: " & lblAscii & color1 & "]", Ascii)
    Call ChatSend("" & color1 & "[" & color2 & "Color1 set by: " & lblColor1 & color1 & "]", Ascii)
    Call ChatSend("" & color1 & "[" & color2 & "Color2 set by: " & lblColor2 & color1 & "]", Ascii): Exit Sub
    End If
    Next x
End If

If LCase(Chat) = Trigger & "wordoftheday" Then
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(1)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
Call ChatSend(WordOfTheDay, Ascii): Exit Sub
End If
Next x
Call ChatSend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
End If

If LCase(Chat) = Trigger & "exit" Then
    If ScreenName = GetUser Then End
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = ">O p t  i o n s>" Then
    Form1.Width = Form1.Width + 3800
    Command1.Caption = "<O p t  i o n s<"
Else
    Form1.Width = Form1.Width - 3800
    Command1.Caption = ">O p t  i o n s>"
End If
End Sub







Private Sub Form_Load()
Dim TestHighlight As ItemColourType
BotStart = Date & " " & Time

Call FormOnTop(Me, True)
'Dim objLvi As MSComctlLib.ListItem: Set objLvi = lst.ListItems.Add()
If file_ifileexists(App.Path & "/Records.txt") = True Then
Call LoadLW(lst, App.Path & "/Records.txt")
End If
If file_ifileexists(App.Path & "/banned.txt") = True Then
Call Loadlistbox(App.Path & "/banned.txt", List1)
End If
If file_ifileexists(App.Path & "/quotes.txt") = True Then
Call Loadlistbox(App.Path & "/quotes.txt", List2)
End If
If file_ifileexists(App.Path & "/phish.txt") = True Then
Call Loadlistbox(App.Path & "/phish.txt", List3)
End If

Dim SNs As String
For d = 1 To lst.ListItems.Count
SNs = SNs & "," & lst.ListItems(d).SubItems(1)
Next d
'SNs = Right(SNs, Len(SNs) - 1)
lblSNs = CountCharAppearance(SNs, ",", False)
'Timer1.enabled = True

'===load check box settings===
If file_ifileexists(App.Path & "\settings.ini") = True Then
For x = 0 To Check1.Count - 1
Check1(x).Value = Val(readini("Settings", "Option" & x, App.Path & "\settings.ini"))
Next x
AOL9.Value = Val(readini("Settings", "aol9", App.Path & "\settings.ini"))
StatusMSG = readini("Settings", "Status_MSG", App.Path & "\settings.ini")
lblEnter = readini("Settings", "RoomEnter", App.Path & "\settings.ini")
Ascii = readini("Settings", "ascii", App.Path & "\settings.ini")
Trigger = readini("Settings", "trigger", App.Path & "\settings.ini")
color1 = readini("Settings", "color1", App.Path & "\settings.ini")
color2 = readini("Settings", "color2", App.Path & "\settings.ini")
lblColor1 = readini("Settings", "setcolor1", App.Path & "\settings.ini")
lblColor2 = readini("Settings", "setcolor2", App.Path & "\settings.ini")
lblTrigger = readini("Settings", "settrigger", App.Path & "\settings.ini")
lblAscii = readini("Settings", "setascii", App.Path & "\settings.ini")
DialerLogin = readini("Dialer", "User:Password", App.Path & "\settings.ini")
mnuSetDialer.Caption = "Set Dialer Login / " & DialerLogin

End If
'===

Chat1.ScanON
Chat2.ScanON
aim_chat1.aimchatscan_on
aim_chat2.aimchatscan_on

Call ChatSend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b> v" & App.Major & "." & App.Minor & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "Loaded</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
'Call AIM_chatsend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b>" & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "Loaded</u><b>" & color1 & "]<font color=#fefcfe>", ascii)

'Call ChatSend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b>" & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "Loaded</u><b>" & color1 & "]", ascii)



'ModLVSubClass.Attach Me.hWnd, lst
    
    'ModLVSubClass.UseCustomHighLight True
    'ModLVSubClass.UseAlternatingColour False
    
    'TestHighlight.BackGround = RGB(82, 151, 249)
    'TestHighlight.ForeGround = RGB(255, 255, 255)
    
    'ModLVSubClass.SetHighLightColour TestHighlight
    
    'TestHighlight.BackGround = RGB(82, 151, 249)
    'TestHighlight.ForeGround = RGB(0, 255, 0)
    
    'ModLVSubClass.SetCustomColour TestHighlight
    
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, Y As Single)
If button = vbLeftButton Then
ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
Else
PopupMenu Form1.mnuFile
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call mnuSave_Click
Call ChatSend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b> v" & App.Major & "." & App.Minor & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "UnLoaded</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
Call AIM_chatsend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b> v" & App.Major & "." & App.Minor & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "UnLoaded</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
Chat1.ScanOFF
Chat2.ScanOFF
Call Startrek(Me)
'Call ChatSend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b>" & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b><u>" & color2 & "UnLoaded</u><b>" & color1 & "]", ascii)

End Sub


Private Sub Frame1_MouseDown(button As Integer, Shift As Integer, x As Single, Y As Single)
If button = vbLeftButton Then
ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
Else
PopupMenu Form1.mnuFile
End If
End Sub

Private Sub aim_chat1_aimchatscan(ScreenName As String, Chat As String, aim_datesaid As String)
Dim SNs As String, SNs2 As String, Voiced As Integer, Op As Integer, counter As Integer
Dim sn As String, SN2 As String, msgs As Integer, Temp As String
Dim FROM As String, AIM As Integer
For d = 1 To lst.ListItems.Count
SNs = SNs & lst.ListItems(d).SubItems(1)
SNs2 = SNs2 & lst.ListItems(d).SubItems(2)
If Len(lst.ListItems(d).SubItems(6)) > 2 Then Voiced = Voiced + 1
If Len(lst.ListItems(d).SubItems(7)) > 2 Then Op = Op + 1
Next d
lblSNs = CountCharAppearance(SNs, ",", False)
lblAIMs = CountCharAppearance(SNs2, ",", False)
lblVoiced = Voiced
lblOp = Op
On Error Resume Next:
Chat = LTrim(RTrim(Chat))
Text2 = Chat
For g = 0 To List1.ListCount
If TrimSpaces(LCase(List1.List(g))) = TrimSpaces(LCase(ScreenName)) Then Exit Sub
Next g
  
  For x = 1 To lst.ListItems.Count
  If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
  For d = 0 To List1.ListCount - 1
  If LCase(List1.List(d)) = LCase(lst.ListItems(x)) Then Exit Sub
  Next d
  End If
  
  
  Next x
   Dim cSend As String
   Dim cData As String
    Dim lngSpace As Long, strCommand As String, strArgument1 As String
   Dim strArgument2 As String, lngComma As Long
   If InStr(Chat, Trigger) = 1& Then
      lngSpace& = InStr(Chat, " ")
      If lngSpace& = 0& Then
         strCommand$ = Chat
      Else
         strCommand$ = Left(Chat, lngSpace& - 1&)
      strCommand$ = Mid(strCommand$, 2, Len(strCommand$))
      End If
   End If
snth4$ = ""
pwth4$ = ""
pw4$ = ""
sn4$ = ""
pwthn$ = ""
snthn$ = ""
snn$ = ""
pwn$ = ""
      Select Case LCase(strCommand$)
      
      Case "handle"
      If Check1(0).Value = 0 Then AIM_chatsend ("Sorry but we are not accepting any new members at this time."): Exit Sub
    Dim strNew As String
    strNew = ""
    srv$ = (Mid(Chat, Len("handle") + 3, Len(Chat) - Len("handle") + 3))
    srv$ = TrimSpaces(srv$)
    'check that handle is < 16 char
If Len(srv$) > 16 Then AIM_chatsend (ScreenName & ", Please try to keep you handle 16 characters or less."): Exit Sub

'check if SN is new
For d = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend("<b>[" & color2 & "</b><u>" & ScreenName & "</u>" & color1 & "]</b> your han<u>dl</u>e is <b>[" & color2 & "<u>" & lst.ListItems(d) & "</u>" & color1 & "]</b>", Ascii)
Exit Sub
End If
Next d

For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend("<b>[" & color2 & "</b><u>" & ScreenName & "</u>" & color1 & "]</b> your han<u>dl</u>e is <b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "]</b>", Ascii)
Exit Sub
End If
Next x
'check if handle is new
    For Z = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(Z)) = LCase(TrimSpaces(srv$)) Then strNew = "No"
    Next Z
    If strNew = "" Then GoTo NewHandle
    
'check if handle is locked
    For x = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(x)) = LCase(TrimSpaces(srv$)) And lst.ListItems(x).SubItems(3) = "Yes" Then
    Call AIM_chatsend("<b>[" & color2 & "<u>Handle is l</b>ock<b>ed</u>" & color1 & "]", Ascii)
    Exit Sub
    End If
    Next x
    
'if handle is not locked then
For n = 1 To lst.ListItems.Count
If lst.ListItems(n) = srv$ Then
lst.ListItems(n).SubItems(2) = lst.ListItems(n).SubItems(2) & TrimSpaces(ScreenName) & "," 'set SN
Call AIM_chatsend("<b>Welcome back  [" & color2 & "<u> " & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)

End If
Next n

Exit Sub
'if handle is new then
NewHandle:
Dim objLvi As MSComctlLib.ListItem: Set objLvi = lst.ListItems.Add()
objLvi.Text = srv$ 'set Handle
objLvi.SubItems(1) = "" 'set SN
objLvi.SubItems(2) = TrimSpaces(ScreenName) & "," 'set AIM
objLvi.SubItems(3) = "No" 'set locked
objLvi.SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat)) 'set seen
objLvi.SubItems(5) = "On" 'set enter
objLvi.SubItems(6) = "No" 'set voice
objLvi.SubItems(7) = "No" 'set Op
objLvi.SubItems(9) = "Welcome New member!" 'set Msg 1
Call AIM_chatsend("<b>[" & color2 & "<U>" & TrimSpaces(srv$) & "</u>" & color1 & "]</b> type <b>.</b>he<u>l</u>p for a list of <b>com</b>mands.", Ascii)

Case "del"
srv$ = (Mid(Chat, Len("del") + 3, Len(Chat) - Len("del") + 3))
If LCase(srv$) = "me" Then
If Check1(1).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "del me""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    Call AIM_chatsend("<b>All info for [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] has been removed.", Ascii)
    lst.ListItems.Remove (x)
    Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If


Case "voice"
    srv$ = (Mid(Chat, Len("voice") + 3, Len(Chat) - Len("voice") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) <= 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(2).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "voice""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If Len(lst.ListItems(n).SubItems(6)) <> 2 Then Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], is already Voiced"): Exit Sub
    lst.ListItems(n).SubItems(6) = ("by " & ScreenName)
    lst.ListItems(n).SubItems(3) = ("Yes")
    Call AIM_chatsend("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], Voiced by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
    
    

Case "devoice"
    srv$ = (Mid(Chat, Len("devoice") + 3, Len(Chat) - Len("devoice") + 3))
    If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & ",</u>" & color1 & "] has been banned.", Ascii): Exit Sub
    If LCase(srv$) = LCase(GetUser) Then List1.AddItem (ScreenName): AIM_chatsend (ScreenName & ", has been banned"): Exit Sub
    
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If LCase(lst.ListItems(x).SubItems(7)) = "by mikestoolz" And ScreenName <> "MikesTooLz" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(2).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "devoice""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    lst.ListItems(n).SubItems(6) = ("No")
    Call AIM_chatsend("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], DeVoiced by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & srv$ & " ,<u>" & color1 & "] is not a current member.")
    
    
    
    
Case "op"
    srv$ = (Mid(Chat, Len("op") + 3, Len(Chat) - Len("op") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(3).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "op""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If Len(lst.ListItems(n).SubItems(7)) <> 2 Then Call ChatSend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], is already OP'd"): Exit Sub
    lst.ListItems(n).SubItems(7) = ("by " & ScreenName)
    lst.ListItems(n).SubItems(3) = ("Yes")
    Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], OP'd by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]"): Exit Sub
    End If
    Next n
    End If
    Next x
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)



Case "deop"
    srv$ = (Mid(Chat, Len("deop") + 3, Len(Chat) - Len("deop") + 3))
    If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): AIM_chatsend (ScreenName & ", has been banned"): Exit Sub
    If LCase(srv$) = LCase(GetUser) Then List1.AddItem (ScreenName): AIM_chatsend (ScreenName & ", has been banned"): Exit Sub
    
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(3).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "deop""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If lst.ListItems(n).SubItems(7) = "by MikesTooLz" And ScreenName <> "MikesTooLz" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    lst.ListItems(n).SubItems(7) = ("No")
    Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "], deOP'd by [" & color2 & "<u>" & ScreenName & "</u>" & color1 & "]"): Exit Sub
    End If
    Next n
    End If
    Next x
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is not a current member.", Ascii)
    
    
    
    
Case "msg"
Dim handle As String, Msg As String
srv$ = (Mid(Chat, Len("msg") + 3, Len(Chat) - Len("msg") + 3))
    For d = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(d).SubItems(2), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(4).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & """messages""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

FROM = lst.ListItems(d)
handle = Left(srv$, InStr(1, srv$, ",", vbTextCompare) - 1)
Msg = Right(srv$, Len(srv$) - InStr(1, srv$, ",", vbTextCompare))
Msg = LTrim(RTrim(Msg))
If Len(Msg) > 400 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & FROM & "</u>" & color1 & "] please keep the msgs under 400 characters"): Exit Sub
For x = 1 To lst.ListItems.Count
If LCase(lst.ListItems(x)) = LCase(handle) Then

Select Case ""
Case lst.ListItems(x).SubItems(9)
lst.ListItems(x).SubItems(9) = ("[" & FROM & "] " & Msg)
Call AIM_chatsend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 1.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
lst.ListItems(x).SubItems(10) = ("[" & FROM & "] " & Msg)
Call AIM_chatsend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 2.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(11)
lst.ListItems(x).SubItems(11) = ("[" & FROM & "] " & Msg)
Call AIM_chatsend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 3.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(12)
lst.ListItems(x).SubItems(12) = ("[" & FROM & "] " & Msg)
Call AIM_chatsend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 4.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(13)
lst.ListItems(x).SubItems(13) = ("[" & FROM & "] " & Msg)
Call AIM_chatsend("<b>msg to [" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] Saved in slot 5.", Ascii): Exit Sub

Case Else
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "'s</u>" & color1 & "] msg slots are full.", Ascii): Exit Sub
End Select
End If
Next x
End If
Next d




Case "msgall"

srv$ = (Mid(Chat, Len("msgall") + 3, Len(Chat) - Len("msgall") + 3))
For d = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(d).SubItems(2), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(d).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If lst.ListItems(d).SubItems(7) = "No" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
If Check1(5).Value = 0 Then Exit Sub

FROM = lst.ListItems(d)
For x = 1 To lst.ListItems.Count
lst.ListItems(x).SubItems(9) = ("[" & FROM & "] " & srv$)
Next x
Call AIM_chatsend("<b>Msg saved in slot 1 of all members.", Ascii)
End If
Next d


Case "ban"
srv$ = (Mid(Chat, Len("ban") + 3, Len(Chat) - Len("ban") + 3))
If LCase(srv$) = "mikestoolz" Then List1.AddItem (ScreenName): Call ChatEjectUser(ScreenName, False): AIM_chatsend (color1 & "[" & color2 & ScreenName & color1 & " has been banned"): Exit Sub
If LCase(srv$) = "mike" Then List1.AddItem (ScreenName): Call ChatEjectUser(ScreenName, False): AIM_chatsend (color1 & "[" & color2 & ScreenName & color1 & " has been banned"): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend(color1 & "<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(6).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "ban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Or InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(srv$) & ",", vbTextCompare) <> 0 Then
    If lst.ListItems(n).SubItems(6) = "by MikesTooLz" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If lst.ListItems(n).SubItems(7) = "by MikesTooLz" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    End If
    Next n
    List1.AddItem (TrimSpaces(LCase(srv$))): Call list_nodupes2(List1)
    AIM_chatsend (color1 & "[" & color2 & srv$ & color1 & "] is now being blocked.")
    End If
    Next x


Case "unban"
srv$ = (Mid(Chat, Len("unban") + 3, Len(Chat) - Len("unban") + 3))
If LCase(srv$) = "mikestoolz" Then List1.AddItem (TrimSpaces(LCase(ScreenName))): AIM_chatsend (ScreenName & ", has been banned"): Exit Sub
If LCase(srv$) = "mike" Then List1.AddItem (TrimSpaces(LCase(ScreenName))): AIM_chatsend (ScreenName & ", has been banned"): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(6).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "unban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    For n = 0 To List1.ListCount - 1
    If LCase(List1.List(n)) = LCase(srv$) Then List1.RemoveItem (n)
    Next n
    AIM_chatsend (color1 & "[" & color2 & srv$ & color1 & "] is no longer being blocked.")
    End If
    Next x
    
    
Case "seen"
srv$ = (Mid(Chat, Len("seen") + 3, Len(Chat) - Len("seen") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If LCase(lst.ListItems(n)) = LCase(srv$) Then

Call AIM_chatsend("" & color1 & "I saw " & srv$ & " in PR <b>" & Right(lst.ListItems(n).SubItems(4), Len(lst.ListItems(n).SubItems(4)) - InStr(1, lst.ListItems(n).SubItems(4), "|", vbTextCompare)) & "</b>[<u>" & color2 & " " & DateDiffEx(Left(lst.ListItems(n).SubItems(4), InStr(1, lst.ListItems(n).SubItems(4), "|", vbTextCompare) - 1), Date & " " & Time) & " ago.</u>" & color1 & "]", Ascii): Exit Sub
End If
Next n
Call AIM_chatsend("" & color1 & "[<u>" & color2 & "" & srv$ & ", is an invalid handle." & "" & color1 & "</u>]")
End If
Next x



Case "whois"
srv$ = (Mid(Chat, Len("whois") + 3, Len(Chat) - Len("whois") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If InStr(1, Replace(lst.ListItems(n).SubItems(2), " ", "", 1, Len(lst.ListItems(n).SubItems(2)), vbTextCompare), Replace(srv$, " ", "", 1, Len(srv$)) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend("" & color1 & "[<u>" & color2 & "" & srv$ & "</u>" & color1 & "] is [<u>" & color2 & "" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii): Exit Sub
End If
Next n
AIM_chatsend (srv$ & ", is not a member.")
End If
Next x



Case "info"

srv$ = (Mid(Chat, Len("info") + 3, Len(Chat) - Len("info") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n), Replace(srv$, " ", "", 1, Len(srv$)), vbTextCompare) <> 0 Then
For d = 1 To lst.ListItems.Count
If LCase(lst.ListItems(d)) = LCase(srv$) Then
SNs = lst.ListItems(d).SubItems(1)


    If lst.ListItems(n).SubItems(8) = "" Then
    Call AIM_chatsend("" & color1 & "[SNs: " & color2 & "" & CountCharAppearance(SNs, ",", False) & "" & color1 & "][AIMs: " & color2 & "" & CountCharAppearance(lst.ListItems(d).SubItems(2), ",", False) & "" & color1 & "][Locked: " & color2 & "" & lst.ListItems(d).SubItems(3) & "" & color1 & "][Voiced: " & color2 & "" & lst.ListItems(d).SubItems(6) & "" & color1 & "]" & "[Oped: " & color2 & "" & lst.ListItems(d).SubItems(7) & "" & color1 & "]", Ascii): Exit Sub
    Else
    Call AIM_chatsend("" & color1 & "[SNs: " & color2 & "" & CountCharAppearance(SNs, ",", False) & "" & color1 & "][AIMs: " & color2 & "" & CountCharAppearance(lst.ListItems(d).SubItems(2), ",", False) & "" & color1 & "][Locked: " & color2 & "" & lst.ListItems(d).SubItems(3) & "" & color1 & "][Voiced: " & color2 & "" & lst.ListItems(d).SubItems(6) & "" & color1 & "]" & "[Oped: " & color2 & "" & lst.ListItems(d).SubItems(7) & "" & color1 & "]" & "[<a href=""" & lst.ListItems(d).SubItems(8) & """>WebSite</a>]", Ascii): Exit Sub
    End If
End If
Next d
End If
Next n
Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "] is an invalid handle.", Ascii)
End If
Next x



Case "read"
srv$ = (Mid(Chat, Len("read") + 3, Len(Chat) - Len("read") + 3))
If LCase(srv$) = LCase("msg") Then
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
Select Case ""
Case lst.ListItems(x).SubItems(9)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have no me<u>ss</u>ages.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 1 me<u>ss</u>age", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
lst.ListItems(x).SubItems(9) = "": Exit Sub
Case lst.ListItems(x).SubItems(11)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 2 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = "": Exit Sub
Case lst.ListItems(x).SubItems(12)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 3 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
Case lst.ListItems(x).SubItems(13)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 4 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
Case Else
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 5 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(13), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
lst.ListItems(x).SubItems(13) = ""
Exit Sub
End Select
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If



Case "enter"
    srv$ = (Mid(Chat, Len("enter") + 3, Len(Chat) - Len("enter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If LCase(srv$) = "on" Then lst.ListItems(x).SubItems(5) = "On": Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] your chat enter is now <u>on</u>.", Ascii): Exit Sub
    If LCase(srv$) = "off" Then lst.ListItems(x).SubItems(5) = "Off": Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] your chat enter is now <u>off</u>.", Ascii): Exit Sub
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
    
    
Case "roomenter"
    srv$ = (Mid(Chat, Len("roomenter") + 3, Len(Chat) - Len("roomenter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If LCase(srv$) = "on" Then lblEnter = "On": Call AIM_chatsend("<b>""Room Enter"" is now on.", Ascii): Exit Sub
    If LCase(srv$) = "off" Then lblEnter = "Off": Call AIM_chatsend("<b>""Room Enter"" is now off.", Ascii): Exit Sub
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii)
    

Case "link"
    If Check1(19).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "link""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("link") + 3, Len(Chat) - Len("link") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""" & srv$ & """>" & srv$ & "</a>")
    End If
    Next x
    
    
    
Case "google"
If Check1(7).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "google""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("google") + 3, Len(Chat) - Len("google") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""http://www.google.com/search?hl=en&q=" & srv$ & """>" & lst.ListItems(x) & ", you were just googled.</a>")
    End If
    Next x



Case "yahoo"
If Check1(7).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "yahoo""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("Yahoo") + 3, Len(Chat) - Len("yahoo") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""http://search.yahoo.com/bin/search?p=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x

Case "altavista"
If Check1(7).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "altavista""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("altavista") + 3, Len(Chat) - Len("altavista") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""http://www.altavista.com/web/results?q=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x



Case "ebay"
If Check1(7).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "ebay""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("ebay") + 3, Len(Chat) - Len("ebay") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""http://search.ebay.com/search/search.dll?cgiurl=http%3A%2F%2Fcgi.ebay.com%2Fws%2F&krd=1&from=R8&MfcISAPICommand=GetResult&ht=1&SortProperty=MetaEndSort&query=" & srv$ & """>" & lst.ListItems(x) & ", here's what you wanted.</a>")
    End If
    Next x
    
    
    
    Case "askjeeves"
    If Check1(7).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "askjeeves""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    srv$ = (Mid(Chat, Len("askjeeves") + 3, Len(Chat) - Len("askjeeves") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("<a href=""http://web.ask.com/web?q=" & srv$ & """>" & lst.ListItems(x) & ", here's what Jeeves has to say.</a>")
    End If
    Next x
    
    
    
Case "addquote"
    srv$ = (Mid(Chat, Len("addquote") + 3, Len(Chat) - Len("addquote") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
    If Check1(8).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addquote""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List2.AddItem ("[" & lst.ListItems(x) & "] " & srv$)
    Call list_nodupes2(List2)
    AIM_chatsend ("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] quote added.")
    Exit Sub
    End If
    Next x
    
    
    
    
Case "allenter"
    srv$ = (Mid(Chat, Len("allenter") + 3, Len(Chat) - Len("allenter") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub  'exit if there not @
    
    For n = 1 To lst.ListItems.Count
    If LCase(srv$) = "on" Then lst.ListItems(n).SubItems(5) = "On"
    If LCase(srv$) = "off" Then lst.ListItems(n).SubItems(5) = "Off"
    Next n
    If LCase(srv$) = "on" Then Call AIM_chatsend("<b>All members ""Room Enter"" had been set to [on].", Ascii)
    If LCase(srv$) = "off" Then Call AIM_chatsend("<b>All members ""Room Enter"" had been set to [off].", Ascii)
    Exit Sub
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
    
    
    
Case "website"
    srv$ = (Mid(Chat, Len("website") + 3, Len(Chat) - Len("website") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Check1(9).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "website""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    lst.ListItems(x).SubItems(8) = srv$
    AIM_chatsend ("<b>" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] website saved.")
    End If
    Next x
    
    
    
Case "addsn"
srv$ = (Mid(Chat, Len("addsn") + 3, Len(Chat) - Len("addsn") + 3))
'find illegal letters
If InStr(1, srv$, "<", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ">", vbTextCompare) <> 0 Then Call AIM_chatsend(ScreenName & ", Invalid SN"): Exit Sub
If InStr(1, srv$, "/", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "`", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "~", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "@", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "#", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "^", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "&", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "*", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "(", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ")", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "_", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "+", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "=", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "-", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "[", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "]", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "{", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "}", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "\", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "|", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "?", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ".", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ",", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, "'", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, """", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ";", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
If InStr(1, srv$, ":", vbTextCompare) <> 0 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & ", Invalid SN!", Ascii): Exit Sub
'check that sn is < 16 char
If Len(srv$) > 16 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Please keep your handle 16 characters or less.", Ascii): Exit Sub

For x = 1 To lst.ListItems.Count
If InStr(1, "," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend(srv$ & ", is " & lst.ListItems(x) & "'s SN.", Ascii): Exit Sub
End If
If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
Call AIM_chatsend(srv$ & ", is " & lst.ListItems(x) & "'s AIM.", Ascii): Exit Sub
End If
Next x
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n).SubItems(2), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(10).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addsn""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
lst.ListItems(n).SubItems(1) = lst.ListItems(n).SubItems(1) & TrimSpaces(srv$) & ","
Call AIM_chatsend("<b>[" & color2 & "<u>" & srv$ & "</u>" & color1 & "] added to handle [" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)
End If
Next n



Case "pr"
Dim lstPR As ComboBox
srv$ = (Mid(Chat, Len("pr") + 3, Len(Chat) - Len("pr") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
Call EnterPR(srv$)
Do
DoEvents
Loop Until LCase(TrimSpaces(srv$)) = LCase(TrimSpaces(GetText(FindChat)))
Combo1.Clear
Call AddAOL8ListToList(ChatPeopleHereList, Combo1, False)
For n = 0 To Combo1.ListCount - 1
If InStr(1, LCase(TrimSpaces(Combo1.List(n))), "host", vbTextCompare) <> 0 Then
Call RandomPR
End If
Next n
Call AIM_chatsend("<b>[<u><font color=""#1E0095"">I</b>nf</u>o<u> bot</u><b>" & color1 & "] B</b><i>y</i> Mik<u>e</u> <B>[</b>Entered: <u>" & color2 & "" & GetText(FindChat) & "</u><b>" & color1 & "]<font color=#fefcfe>", Ascii)
Combo1.Clear
End If
Next x


Case "weather"
srv$ = (Mid(Chat, Len("weather") + 3, Len(Chat) - Len("weather") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(11).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "weather""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
GetWeatherInfo (srv$)
If WeatherInfo.State = "m -" Then Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "]Invalid zipcode", Ascii): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then Call AIM_chatsend("" & color1 & "[" & color2 & "" & srv$ & "" & color1 & "]Invalid zipcode", Ascii): Exit Sub
AIM_chatsend ("" & color2 & "<b>" & WeatherInfo.City & "," & WeatherInfo.State & "</b> " & color1 & "- [Condition: " & color2 & "" & WeatherInfo.CurrentCond & "" & color1 & "] [Temp: " & color2 & "" & WeatherInfo.CurrentF & "º" & WeatherInfo.FeelsLike & "" & color1 & "] [UVIndex: " & color2 & "" & WeatherInfo.UVIndex & "" & color1 & "] [Humidity: " & color2 & "" & WeatherInfo.Humidity & "" & color1 & "] [Wind: " & color2 & "" & WeatherInfo.Wind & "" & color1 & "] [Visibility: " & color2 & "" & WeatherInfo.Visibility & "" & color1 & "]"): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "definition"
srv$ = (Mid(Chat, Len("definition") + 3, Len(Chat) - Len("definition") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(12).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "definition""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

Call AIM_chatsend(Definition(srv$), Ascii): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)


Case "synonyms"
srv$ = (Mid(Chat, Len("synonyms") + 3, Len(Chat) - Len("synonyms") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(13).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "synonyms""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Synonyms(srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




Case "antonyms"
srv$ = (Mid(Chat, Len("antonyms") + 3, Len(Chat) - Len("antonyms") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(13).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "antonyms""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (Antonyms(srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "newhandle"
srv$ = (Mid(Chat, Len("newhandle") + 3, Len(Chat) - Len("newhandle") + 3))
srv$ = TrimSpaces(srv$)
'check that handle is < 16 char
If Len(srv$) > 16 Then AIM_chatsend (ScreenName & ", Please try to keep you handle 16 characters or less."): Exit Sub
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
'check if handle is new
    For Z = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(Z)) = LCase(TrimSpaces(srv$)) Then strNew = "No"
    Next Z
    If strNew = "" Then
    AIM_chatsend (lst.ListItems(x) & "'s handle was changed to <b>" & srv$)
    lst.ListItems(x).Text = srv$
    Else
    AIM_chatsend ("A member is already using the handle <b>" & srv$ & "</b>")
    End If
End If
Next x


Case "sn"
    srv$ = (Mid(Chat, Len("sn") + 3, Len(Chat) - Len("sn") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then AIM_chatsend ("[" & Left(lst.ListItems(n).SubItems(1), Len(lst.ListItems(n).SubItems(1)) - 1) & "]"): Exit Sub
    Next n
    AIM_chatsend (color1 & "[" & color2 & ScreenName & color1 & "] there is no member with that handle."): Exit Sub
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)
    
    
    
Case "horoscope"
srv$ = (Mid(Chat, Len("horoscope") + 3, Len(Chat) - Len("horoscope") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(14).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "horoscope""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

AIM_chatsend (color1 & "[" & color2 & "Too much text to send to AIM chat" & color1 & "]"): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)


Case "remsn"
srv$ = (Mid(Chat, Len("remsn") + 3, Len(Chat) - Len("remsn") + 3))
srv$ = TrimSpaces(srv$)
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(10).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remsn""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

If InStr(1, "," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(srv$) & ",", vbTextCompare) <> 0 Then
Temp = Replace("," & LCase(lst.ListItems(x).SubItems(1)), "," & LCase(srv$) & ",", ",", 1, Len("," & LCase(lst.ListItems(x).SubItems(1))), vbTextCompare)
If Left(Temp, 1) = "," Then Temp = Right(Temp, Len(Temp) - 1)
lst.ListItems(x).SubItems(1) = Temp
AIM_chatsend (srv$ & ",was removed"): Exit Sub
End If

End If
Next x



Case "x"
srv$ = (Mid(Chat, Len("x") + 3, Len(Chat) - Len("x") + 3))
If ScreenName = GetUser Then Call ChatIgnoreUser(srv$, True)


Case "unx"
srv$ = (Mid(Chat, Len("unx") + 3, Len(Chat) - Len("unx") + 3))
If ScreenName = GetUser Then Call ChatIgnoreUser(srv$, True, False)



Case "eject"
srv$ = (Mid(Chat, Len("eject") + 3, Len(Chat) - Len("eject") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(25).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "eject""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
If ScreenName = GetUser Then Call ChatEjectUser(srv$, True): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) <> 2 Then Call ChatEjectUser(srv$, True): Exit Sub
End If
Next x


Case "allow"
srv$ = (Mid(Chat, Len("allow") + 3, Len(Chat) - Len("allow") + 3))
If ScreenName = GetUser Then Call ChatAllowUser(srv$, True): Exit Sub
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(1), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(25).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "eject""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
If Len(lst.ListItems(x).SubItems(7)) = 2 Then Exit Sub
If Len(lst.ListItems(x).SubItems(7)) <> 2 Then Call ChatEjectUser(srv$, True): Exit Sub
End If
Next x



Case "addaim"
srv$ = (Mid(Chat, Len("addaim") + 3, Len(Chat) - Len("addaim") + 3))
'find illegal letters
If InStr(1, srv$, "<", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ">", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "/", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "`", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "~", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "@", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "#", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "%", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "^", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "&", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "*", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "(", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ")", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "_", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "+", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "=", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "-", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "[", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "]", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "{", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "}", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "\", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "|", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "?", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ".", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ",", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, "'", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, """", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ";", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
If InStr(1, srv$, ":", vbTextCompare) <> 0 Then AIM_chatsend (ScreenName & ", Invalid AIM"): Exit Sub
'check that AIM is < 24 char
If Len(srv$) > 24 Then AIM_chatsend (ScreenName & ", Please try to keep you handle 24 characters or less."): Exit Sub

For x = 1 To lst.ListItems.Count
    If InStr(1, LCase("," & lst.ListItems(x).SubItems(1)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend (srv$ & ", is " & lst.ListItems(x) & "'s SN."): Exit Sub
    End If
    If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(TrimSpaces(srv$)) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend (srv$ & ", is " & lst.ListItems(x) & "'s AIM."): Exit Sub
    End If
Next x
For n = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(n).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(15).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "addaim""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
lst.ListItems(n).SubItems(2) = lst.ListItems(n).SubItems(2) & TrimSpaces(srv$) & ","
Call AIM_chatsend("<b>[" & color2 & "<u>" & srv$ & "</u>" & color1 & "] added to handle [" & color2 & "<u>" & lst.ListItems(n) & "</u>" & color1 & "]", Ascii)
End If
Next n



Case "remaim"
srv$ = (Mid(Chat, Len("remaim") + 3, Len(Chat) - Len("remaim") + 3))
srv$ = TrimSpaces(srv$)
For x = 1 To lst.ListItems.Count
If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
If Check1(15).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "remaim""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub

If InStr(1, "," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(srv$) & ",", vbTextCompare) <> 0 Then
Temp = Replace("," & LCase(lst.ListItems(x).SubItems(2)), "," & LCase(srv$) & ",", ",", 1, Len("," & LCase(lst.ListItems(x).SubItems(2))), vbTextCompare)
If Left(Temp, 1) = "," Then Temp = Right(Temp, Len(Temp) - 1)
lst.ListItems(x).SubItems(2) = Temp
AIM_chatsend (srv$ & ", was removed"): Exit Sub
End If
End If
Next x



Case "aim"
    srv$ = (Mid(Chat, Len("aim") + 3, Len(Chat) - Len("aim") + 3))
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    For n = 1 To lst.ListItems.Count
    If LCase(lst.ListItems(n)) = LCase(srv$) Then
    If lst.ListItems(n).SubItems(2) = "" Then Call AIM_chatsend("" & color1 & "[" & color2 & "" & lst.ListItems(n) & "" & color1 & "], has no AIMs.", Ascii): Exit Sub
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & Left(lst.ListItems(n).SubItems(2), Len(lst.ListItems(n).SubItems(2)) - 1) & "" & color1 & "]", Ascii): Exit Sub
    End If
    Next n
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & ScreenName & "" & color1 & "] there is no member with that handle.", Ascii): Exit Sub
    End If
    Next x
    Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)



Case "name"
srv$ = (Mid(Chat, Len("name") + 3, Len(Chat) - Len("name") + 3))
For x = 1 To lst.ListItems.Count
If InStr(1, LCase(lst.ListItems(x).SubItems(2)), TrimSpaces(LCase(ScreenName)) & ",", vbTextCompare) <> 0 Then
If Check1(20).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "name""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
AIM_chatsend (WhatNameMeans(srv$)): Exit Sub
End If
Next x
Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] register a handle first!", Ascii)




    
    
    
    
 
    



     
     
End Select
'========================================================
If LCase(Chat) = Trigger & "lock" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    lst.ListItems(x).SubItems(3) = "Yes"
    Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] is now locked.", Ascii)
    Exit Sub
    End If
    Next x
      Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If



If LCase(Chat) = Trigger & "unlock" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) > 2 Then Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] Op's can't be unlocked. Use " & Trigger & "addsn"" to add a new SN.", Ascii): Exit Sub
    lst.ListItems(x).SubItems(3) = "No"
    Call AIM_chatsend("" & color1 & "[" & color2 & "<u>" & lst.ListItems(x) & "" & color1 & "</u>] is no longer locked.", Ascii)
    Exit Sub
    End If
    Next x
      Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "], your not a member!", Ascii)
End If






If InStr(1, Chat, "has left the room.", vbTextCompare) <> 0 Then
'If ScreenName = "" Then
sn = Left(Chat, InStr(1, Chat, "has left the room.", vbTextCompare) - 2)
For x = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(x).SubItems(2), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
lst.ListItems(x).SubItems(4) = (Date & " " & Time & "|" & GetCaption(AIM_FindChat))
End If
Next x
'End If
End If


If LCase(Chat) = Trigger & "banlist" Then
For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(6)) = 2 And Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(18).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "banlist""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    If List1.ListCount = 0 Then Call AIM_chatsend(ScreenName & ", no one is banned.", Ascii)
    For n = 0 To List1.ListCount - 1
    Call AIM_chatsend("" & color1 & "[" & color2 & "" & n + 1 & "" & color1 & "] " & List1.List(n), Ascii)
    Next n
    End If
    Next x
End If



If LCase(Chat) = Trigger & "quote" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend (List2.List(randomnumber(List2.ListCount) - 1))
    End If
    Next x
End If


If LCase(Chat) = Trigger & "clearban" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(16).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "clearban""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List1.Clear
    AIM_chatsend ("Banned list has been cleared."): Exit Sub
    End If
    Next x
End If


If LCase(Chat) = Trigger & "clearquotes" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If Len(lst.ListItems(x).SubItems(7)) = 2 Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    If Check1(17).Value = 0 Then Call AIM_chatsend("<b>Sorry but [<u>" & color2 & "" & Trigger & "clearquotes""</u>" & color1 & "] has been disabled.", Ascii): Exit Sub
    List2.Clear
    AIM_chatsend ("Quotes list has been cleared."): Exit Sub
    End If
    Next x
End If



If LCase(Chat) = Trigger & "status" Then
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    AIM_chatsend ("" & color1 & "Handles:[" & color2 & "" & lst.ListItems.Count & "" & color1 & "] SNs:[" & color2 & "" & lblSNs & "" & color1 & "] AIMs:[" & color2 & "" & lblAIMs & "" & color1 & "] Voiced:[" & color2 & "" & lblVoiced & "" & color1 & "] Oped:[" & color2 & "" & lblOp & "" & color1 & "] Quotes:[" & color2 & "" & List2.ListCount & "" & color1 & "] Banned:[" & color2 & "" & List1.ListCount & "" & color1 & "]" & StatusMSG)
    End If
    Next x
End If




If LCase(Chat) = Trigger & "help" Then
    AIM_chatsend ("[<a href=""http://www.mikestoolz.com/downloads/commands.html"">" & ScreenName & " click here for list of commands</a>]")
End If
    
    
    
    
If LCase(Chat) = Trigger & "oplist" Then
    counter = 0
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If lst.ListItems(x).SubItems(7) = "No" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
    If Len(lst.ListItems(n).SubItems(7)) > 2 Then counter = counter + 1: Call AIM_chatsend("" & color1 & "[" & color2 & "" & counter & "" & color1 & "] " & lst.ListItems(n), Ascii)
    Next n
    End If
    Next x
End If



If LCase(Chat) = Trigger & "vlist" Then
    counter = 0
    For x = 1 To lst.ListItems.Count
    If InStr(1, "," & lst.ListItems(x).SubItems(2), "," & TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Or InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(ScreenName) & ",", vbTextCompare) <> 0 Then
    If lst.ListItems(x).SubItems(6) = "No" And lst.ListItems(x).SubItems(7) = "No" Then Call AIM_chatsend("<b>[" & color2 & "<u>" & ScreenName & "</u>" & color1 & "] Access <U>Denied</u>", Ascii): Exit Sub
    For n = 1 To lst.ListItems.Count
    If Len(lst.ListItems(n).SubItems(6)) > 2 Then counter = counter + 1: Call AIM_chatsend("" & color1 & "[" & color2 & "" & counter & "" & color1 & "] " & lst.ListItems(n), Ascii)
    Next n
    End If
    Next x
End If


If LCase(Chat) = Trigger & "exit" Then
    If ScreenName = GetUser Then End
End If



End Sub


Private Sub Frame2_MouseDown(button As Integer, Shift As Integer, x As Single, Y As Single)
If button = vbLeftButton Then
ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
Else
PopupMenu Form1.mnuFile
End If
End Sub

Private Sub lblAIMs_Change()
Bar1.Panels.Item(2) = "AIM's: " & lblAIMs
End Sub

Private Sub lblEnter_Change()
Bar1.Panels.Item(5) = "Room Enter: " & lblEnter
End Sub

Private Sub lblOp_Change()
Bar1.Panels.Item(4) = "Oped: " & lblOp
End Sub

Private Sub lblSNs_Change()
Bar1.Panels.Item(1) = "SN's: " & lblSNs
End Sub

Private Sub lblVoiced_Change()
Bar1.Panels.Item(3) = "voiced: " & lblVoiced
End Sub

Private Sub List1_DblClick()
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub List2_DblClick()
List2.RemoveItem (List2.ListIndex)
End Sub

Private Sub List3_DblClick()
List3.RemoveItem (List3.ListIndex)
End Sub

Private Sub lst_DblClick()
Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String, str7 As String, str8 As String, str9 As String, str10 As String, str11 As String, str12 As String, str13 As String, str14 As String

str1 = lst.SelectedItem
str2 = lst.SelectedItem.SubItems(1)
str3 = lst.SelectedItem.SubItems(2)
str4 = lst.SelectedItem.SubItems(3)
str5 = lst.SelectedItem.SubItems(4)
str6 = lst.SelectedItem.SubItems(5)
str7 = lst.SelectedItem.SubItems(6)
str8 = lst.SelectedItem.SubItems(7)
str9 = lst.SelectedItem.SubItems(8)
str10 = lst.SelectedItem.SubItems(9)
str11 = lst.SelectedItem.SubItems(10)
str12 = lst.SelectedItem.SubItems(11)
str13 = lst.SelectedItem.SubItems(12)
str14 = lst.SelectedItem.SubItems(13)

str1 = InputBox("Handle: ", "Handle", str1)
str2 = InputBox("SN: ", "SN", str2)
str3 = InputBox("AIM: ", "AIM", str3)
str4 = InputBox("Locked: ", "Locked", str4)
str5 = InputBox("Seen: ", "Seen:", str5)
str6 = InputBox("Enter: ", "Enter", str6)
str7 = InputBox("Voiced: ", "Voiced", str7)
str8 = InputBox("Op: ", "Op", str8)
str9 = InputBox("URL: ", "URL", str9)
str10 = InputBox("MSG1: ", "MSG1", str10)
str11 = InputBox("MSG2: ", "MSG2", str11)
str12 = InputBox("MSG3: ", "MSG3", str12)
str13 = InputBox("MSG4: ", "MSG4", str13)
str14 = InputBox("MSG5: ", "MSG5", str14)

lst.SelectedItem = str1
lst.SelectedItem.SubItems(1) = str2
lst.SelectedItem.SubItems(2) = str3
lst.SelectedItem.SubItems(3) = str4
lst.SelectedItem.SubItems(4) = str5
lst.SelectedItem.SubItems(5) = str6
lst.SelectedItem.SubItems(6) = str7
lst.SelectedItem.SubItems(7) = str8
lst.SelectedItem.SubItems(8) = str9
lst.SelectedItem.SubItems(9) = str10
lst.SelectedItem.SubItems(10) = str11
lst.SelectedItem.SubItems(11) = str12
lst.SelectedItem.SubItems(12) = str13
lst.SelectedItem.SubItems(13) = str14

End Sub

Private Sub NiteScan1_ChatScan(ScreenName As String, Chat As String)
Text1 = Chat
End Sub

Private Sub lst_MouseDown(button As Integer, Shift As Integer, x As Single, Y As Single)
If button = vbLeftButton Then
ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
Else
PopupMenu Form1.mnuLST
End If
End Sub

Private Sub lst_MouseMove(button As Integer, Shift As Integer, x As Single, Y As Single)
'For i = 1 To lst.ListItems.Count 'Goes through all items In the listView
        'checks to see if the mouse is over the
        '     current listView item
        'If (X > lst.ListItems.Item(i).Left) And _
        '(X < (lst.ListItems.Item(i).Left + lst.ListItems.Item(i).Width)) _
        'And (Y > lst.ListItems.Item(i).Top) And _
        '(Y < lst.ListItems.Item(i).Top + lst.ListItems.Item(i).Height) Then
        'if it is, set all to default, in this c
        '     ase, black


        'For b = 1 To lst.ListItems.Count
            'lst.ListItems.Item(b).ForeColor = vbBlack
        'Next b
        'sets the one that the mouse is over to
        '     Blue, can be changed.
        'lst.ListItems.Item(i).ForeColor = vbBlue
        
    'End If
'Next i
End Sub

Private Sub mnuAscii_Click()
Dim strmsg As String
If Ascii = "" Then
strmsg = InputBox("What would you like the ascii to be?", "ascii")
Else
strmsg = InputBox("What would you like the ascii to be?", "ascii", Ascii)
End If
If strmsg = vbcancle Then Exit Sub
If LCase(strmsg) = "off" Then Ascii = "": Exit Sub
Ascii = strmsg
End Sub

Private Sub mnuChatSplit_Click()
Dim daRoom As String, room As String
room = GetText(FindChat)
daRoom = InputBox("What other chat you wanna be in?")
Call EnterPR(daRoom): Call EnterPR(room)
End Sub

Private Sub mnuDelete_Click()
lst.ListItems.Remove (lst.SelectedItem.index)
End Sub

Private Sub mnuDialer_Click()
If DialerLogin = "" Then MsgBox "Set a username and login first.": Exit Sub
GetDialerStats
MsgBox ("dialer stats - [" & "Total min/calls:" & DialerInfo.TotalMinutes_Calls & "] [Average:" & DialerInfo.OverallAverage & "] [This Month:" & DialerInfo.ThisMonthMinutes_Calls & "] [Months Average:" & DialerInfo.MonthAverage & "] [Todays Min/Calls:" & DialerInfo.TodayMinutes_Calls & "] [Cash:" & "$" & DialerInfo.Cash & "]")
End Sub

Private Sub mnuEdit_Click()
Call lst_DblClick
End Sub

Private Sub mnuPause_Click()
If mnuPause.Caption = "Pause" Then
Chat1.ScanOFF
mnuPause.Caption = "UnPause"
ChatSend ("[info bot] - Paused")
Else
Chat1.ScanON
mnuPause.Caption = "Pause"
ChatSend ("[info bot] - UnPaused")
End If
End Sub

Private Sub mnuSave_Click()
Call SaveLW(lst, App.Path & "/Records.txt")
Call SaveListBox(App.Path & "/banned.txt", List1)
Call SaveListBox(App.Path & "/quotes.txt", List2)
Call SaveListBox(App.Path & "/phish.txt", List3)
For x = 0 To Check1.Count - 1
Call writeini("Settings", "Option" & x, "" & Check1(x).Value & "", App.Path & "\settings.ini")
Next x
Call writeini("Settings", "aol9", "" & AOL9.Value & "", App.Path & "\settings.ini")
Call writeini("Settings", "Status_MSG", StatusMSG, App.Path & "\settings.ini")
Call writeini("Settings", "RoomEnter", lblEnter, App.Path & "\settings.ini")
Call writeini("Settings", "ascii", Ascii, App.Path & "\settings.ini")
Call writeini("Settings", "Trigger", Trigger, App.Path & "\settings.ini")
Call writeini("Settings", "color1", color1, App.Path & "\settings.ini")
Call writeini("Settings", "color2", color2, App.Path & "\settings.ini")
Call writeini("Settings", "setcolor1", lblColor1, App.Path & "\settings.ini")
Call writeini("Settings", "setcolor2", lblColor2, App.Path & "\settings.ini")
Call writeini("Settings", "settrigger", lblTrigger, App.Path & "\settings.ini")
Call writeini("Settings", "setascii", lblAscii, App.Path & "\settings.ini")
Call writeini("Dialer", "User:Password", DialerLogin, App.Path & "\settings.ini")

End Sub

Private Sub mnuSetDialer_Click()
DialerLogin = InputBox("User:Password", "User Login")
If DialerLogin = "" Then DialerLogin = ":"
If InStr(1, DialerLogin, ":", vbTextCompare) = 0 Then DialerLogin = ":"
mnuSetDialer.Caption = "Set Dialer Login / " & DialerLogin
End Sub

Private Sub mnuStatusMSG_Click()
Dim strmsg As String
If StatusMSG = "" Then
strmsg = InputBox("What would you like the status msg to be?", "Status Message")
Else
strmsg = InputBox("What would you like the status msg to be?", "Status Message", Mid(StatusMSG, 55, Len(StatusMSG) - 77))
End If
If strmsg = vbcancle Then Exit Sub
If LCase(strmsg) = "off" Then StatusMSG = "": Exit Sub
StatusMSG = "<b>" & color1 & " MSG: [" & color2 & "" & strmsg & "" & color1 & "]"
End Sub

Private Sub Text2_Change()
Dim sn As String
If InStr(1, Text2, "has entered the room.", vbTextCompare) <> 0 Then
'If ScreenName = "OnlineHost" Then
sn = Left(Text2, InStr(1, Text2, "has entered the room.", vbTextCompare) - 2)

If InStr(1, LCase(TrimSpaces(sn)), "host", vbTextCompare) Then Call RandomPR
'====if there banned then say it
For Z = 1 To lst.ListItems.Count
  If InStr(1, lst.ListItems(Z).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
  SN2 = lst.ListItems(Z)
  For F = 0 To List1.ListCount - 1
  If LCase(List1.List(F)) = LCase(SN2) Then Call AIM_chatsend("" & color1 & "[" & color2 & "" & sn & "" & color1 & "] you are banned. All commands disabled.", Ascii): Call ChatEjectUser(sn, False): Exit Sub
  If LCase(List1.List(F)) = LCase(sn) Then Call AIM_chatsend("" & color1 & "[" & color2 & "" & sn & "" & color1 & "] you are banned. All commands disabled.", Ascii): Call ChatEjectUser(sn, False): Exit Sub
  
  Next F
  End If
Next Z
'======

For x = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
lst.ListItems(x).SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat))
If lst.ListItems(x).SubItems(5) = "Off" Then Exit Sub
If LCase(lblEnter) = "off" Then Exit Sub
Select Case ""
Case lst.ListItems(x).SubItems(9)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have no me<u>ss</u>ages.", Ascii): Exit Sub
Case lst.ListItems(x).SubItems(10)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 1 me<u>ss</u>age", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
lst.ListItems(x).SubItems(9) = "": Exit Sub
Case lst.ListItems(x).SubItems(11)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 2 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = "": Exit Sub
Case lst.ListItems(x).SubItems(12)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 3 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
Case lst.ListItems(x).SubItems(13)
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 4 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
Case Else
Call AIM_chatsend("<b>[" & color2 & "<u>" & lst.ListItems(x) & "</u>" & color1 & "] you have 5 me<u>ss</u>age<u>s</u>", Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(9), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(10), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(11), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(12), Ascii)
Call AIM_chatsend("" & color1 & "" & lst.ListItems(x).SubItems(13), Ascii)
lst.ListItems(x).SubItems(9) = ""
lst.ListItems(x).SubItems(10) = ""
lst.ListItems(x).SubItems(11) = ""
lst.ListItems(x).SubItems(12) = ""
lst.ListItems(x).SubItems(13) = ""

End Select
End If
Next x
If Check1(0).Value = 0 Then Exit Sub
If LCase(lblEnter) = "off" Then Exit Sub
Call AIM_chatsend("" & color1 & "" & sn & " type " & Trigger & "handle "" and your handle(name)")
'End If
End If


If InStr(1, Text2, "has left the room.", vbTextCompare) <> 0 Then
'If ScreenName = "OnlineHost" Then
sn = Left(Text2, InStr(1, Text2, "has left the room.", vbTextCompare) - 2)
For x = 1 To lst.ListItems.Count
If InStr(1, lst.ListItems(x).SubItems(1), TrimSpaces(sn) & ",", vbTextCompare) <> 0 Then
lst.ListItems(x).SubItems(4) = (Date & " " & Time & "|" & GetCaption(FindChat))
End If
Next x
'End If
End If

End Sub

Private Sub Text3_Change()
Call SetDialerStats
End Sub

Private Sub Timer1_Timer()
Call SaveLW(lst, App.Path & "/Records.dat")
Call SaveListBox(App.Path & "/banned.txt", List1)
Call SaveListBox(App.Path & "/quotes.txt", List2)
End Sub

Private Sub Timer2_Timer()
Dim room As String
If find_RoomClosed <> 0 Then
room = GetText(FindChat)
ClickIcon (icon_roomclose)
EnterPR (room)
End If
End Sub
