VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MyIM 
   Caption         =   "MyInstant Messenger"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyIM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer BuddyUpdater 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1560
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   6000
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   2715
      TabIndex        =   3
      Top             =   4635
      Width           =   2775
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ready."
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   495
      Picture         =   "MyIM.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Change my Status "
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      Picture         =   "MyIM.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Send a Message "
      Top             =   0
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyIM.frx":06D6
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyIM.frx":0C28
            Key             =   "Offline"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyIM.frx":0F7A
            Key             =   "Away"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyIM.frx":12CC
            Key             =   "DND"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MyIM.frx":161E
            Key             =   "Unknown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6376
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2085
      TabIndex        =   4
      Top             =   150
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileToggleLog 
         Caption         =   "&Log On"
      End
      Begin VB.Menu mnuFileSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "My &Status"
         Enabled         =   0   'False
         Begin VB.Menu mnuStatusOnline 
            Caption         =   "&Online"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuStatusAway 
            Caption         =   "&Away"
         End
         Begin VB.Menu mnuStatusDND 
            Caption         =   "&Do Not Disturb"
         End
         Begin VB.Menu mnuStatusSplit 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStatusInvisible 
            Caption         =   "&Invisible"
         End
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuBuddy 
      Caption         =   "&Buddy"
      Begin VB.Menu mnuBuddyMessage 
         Caption         =   "Send &Message"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBuddySplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBuddyFile 
         Caption         =   "Transfer &File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBuddyChat 
         Caption         =   "Request &Chat"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "MyIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BuddyUpdater_Timer()

Label2.Caption = "Checking buddy... (0/" & TreeView1.Nodes.Count & ")"

For i = 1 To TreeView1.Nodes.Count
   Label2.Caption = "Checking buddy... (" & i & "/" & TreeView1.Nodes.Count & ")"
   Winsock1.SendData ".getstatus " & TreeView1.Nodes(i).Key
   timedPause 1
Next

Label2.Caption = ""

End Sub

Private Sub Command1_Click()

mnuBuddyMessage_Click

End Sub

Private Sub Command2_Click()

PopupMenu mnuStatus, , Command2.Left, Command2.Top + Command2.Height

End Sub

Private Sub Form_Load()

Connect.Show

End Sub

Private Sub Form_Resize()
On Error Resume Next

Shape1.Width = Me.ScaleWidth
Label1.Left = Me.ScaleWidth - Label1.Width - 120

TreeView1.Width = Me.ScaleWidth
TreeView1.Height = Me.ScaleHeight - Shape1.Height - Picture1.Height

End Sub


Private Sub mnuAbout_Click()

MsgBox "MyInstant Messenger" & vbCrLf & vbCrLf & "A fully functional Instant Messenger source code example for Visual Basic 6, created by Evan Christopher Sims of Post for Help. Please visit our website at:" & vbCrLf & vbCrLf & "http://members.xoom.com/esims/postforhelp/" & vbCrLf & vbCrLf & "If you use this code, please provide a link to the Post for Help site so others can partake in it's tutorial goodness. ;)", vbInformation

End Sub

Private Sub mnuBuddyMessage_Click()

On Error Resume Next

   'MsgBox TreeView1.SelectedItem.Tag
   Dim NewIMessage As New IMessage
   NewIMessage.Show ownerform:=Me
   
   NewIMessage.Label2.Caption = TreeView1.SelectedItem
   NewIMessage.RecieversID = TreeView1.SelectedItem.Key

End Sub

Private Sub mnuFileClose_Click()

   Unload Me

End Sub

Private Sub mnuFileToggleLog_Click()

If mnuFileToggleLog.Caption = "&Log Off" Then

   mnuFileToggleLog.Caption = "&Log On"
   Winsock1.Close
   BuddyUpdater.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   Label1.Caption = "Offline"
   TreeView1.Nodes.Clear
   mnuBuddyMessage.Enabled = False
   mnuStatus.Enabled = False


ElseIf mnuFileToggleLog.Caption = "&Log On" Then

   Connect.Show ownerform:=Me

End If

End Sub

Private Sub mnuStatusAway_Click()

mnuStatusOnline.Checked = False
mnuStatusAway.Checked = True
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = False

Label1.Caption = "Away"
Winsock1.SendData ".status Away"

End Sub

Private Sub mnuStatusDND_Click()

mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = True
mnuStatusInvisible.Checked = False

Label1.Caption = "DND"
Winsock1.SendData ".status DND"

End Sub


Private Sub mnuStatusInvisible_Click()

mnuStatusOnline.Checked = False
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = True

Label1.Caption = "Invisible"
Winsock1.SendData ".status Invisible"

End Sub


Private Sub mnuStatusOnline_Click()

mnuStatusOnline.Checked = True
mnuStatusAway.Checked = False
mnuStatusDND.Checked = False
mnuStatusInvisible.Checked = False

Label1.Caption = "Online"
Winsock1.SendData ".status Online"

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

On Error Resume Next

If TreeView1.SelectedItem.Text <> "" Then
   PopupMenu mnuBuddy
End If

End Sub


Private Sub Winsock1_Close()

   mnuStatus.Enabled = False
   mnuBuddyMessage.Enabled = False
   Label1.Caption = "Offline"
   Command1.Enabled = False
   Command2.Enabled = False
   BuddyUpdater.Enabled = False
   mnuFileToggleLog.Caption = "&Log On"
   Winsock1.Close

End Sub

Private Sub Winsock1_Connect()

   Command1.Enabled = True
   Command2.Enabled = True
   mnuFileToggleLog.Caption = "&Log Off"

End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

If Winsock1.State = 7 Then
    Dim UserCommand As String
    Winsock1.GetData UserCommand
    
       If Word(UserCommand, 1) = ".loginok" Then
    
          Connect.Label1.Caption = "Entering Service..."
          Unload Connect

          Winsock1.SendData ".status Online"
          Label1.Caption = "Online"
          mnuStatusOnline.Checked = True
          mnuBuddyMessage.Enabled = True
          mnuStatus.Enabled = True
          timedPause 1

          Winsock1.SendData ".getbuddys"
          'Listing.Show ownerform:=Me

          BuddyUpdater.Enabled = True
    
       ElseIf Word(UserCommand, 1) = ".loginbad" Then
    
          Connect.Label1.Caption = "Oops!"
          Connect.Label2.Caption = "Uh oh! Sorry, but it looks like "
          
          If Word(UserCommand, 2) = "0" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "your account couldn't be found! Try re-entering your username."
          ElseIf Word(UserCommand, 2) = "1" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "your password is wrong! Try re-entering it."
          ElseIf Word(UserCommand, 2) = "2" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "your account has been temporarily banned, for some reason or other. Please contact BinaryPoint customer service for details on the reason, and the remaining ban time.."
          ElseIf Word(UserCommand, 2) = "3" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "your account has been banned, for some reason or other. Please contact BinaryPoint customer service for details on the reason."
          ElseIf Word(UserCommand, 2) = "4" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "your account has been frozen, for some reason or other. Please contact BinaryPoint customer service for details on the reason."
          ElseIf Word(UserCommand, 2) = "5" Then
             Connect.Label2.Caption = Connect.Label2.Caption & "the server is full. Please try again in a little while."
          End If
          
          Connect.Label3.ForeColor = vbBlack
          Connect.Label4.ForeColor = vbBlack
          Connect.Text1.Enabled = True
          Connect.Text2.Enabled = True
          Connect.Command1.Enabled = True
          Connect.Command2.Caption = "&Close"
          Winsock1.Close
    
       ElseIf Word(UserCommand, 1) = ".msg" Then
    
          Dim NewReponseMessage As New GotMessage
          NewReponseMessage.Show ownerform:=Me
    
          NewReponseMessage.Caption = "Message from " & Trim(Replace(Word(UserCommand, 3), "_._", " "))
    
          NewReponseMessage.Label2.Caption = Trim(Replace(Word(UserCommand, 3), "_._", " ")) & " (" & Trim(Word(UserCommand, 2) & ")")
          NewReponseMessage.SenderID = Trim(Word(UserCommand, 2))
          NewReponseMessage.SenderName = Trim(Replace(Word(UserCommand, 3), "_._", " "))

          NewReponseMessage.RichTextBox1.TextRTF = Trim(Replace(SplitString(UserCommand, "..//.."), "//crlf\\", vbCrLf))

       ElseIf Word(UserCommand, 1) = ".pushbuddyupdate" Then
          
          For i = 1 To TreeView1.Nodes.Count
             If TreeView1.Nodes(i).Key = Word(UserCommand, 2) Then
                TreeView1.Nodes(i).Image = Word(UserCommand, 3)
                TreeView1.Nodes(i).SelectedImage = Word(UserCommand, 3)
                TreeView1.Refresh
                Exit For
             End If
          Next

       ElseIf Word(UserCommand, 1) = ".pushbuddy" Then
          Dim BuddyUserID
          Dim BuddyUserTitle
       
          BuddyStatus = Word(UserCommand, 2)
          BuddyUserID = Word(UserCommand, 3)
          BuddyUserTitle = SplitString(UserCommand, Word(UserCommand, 3))
    
          'MsgBox "Server pushed user " & SplitString(UserCommand, ".pushbuddy") & " to me!"
          TreeView1.Nodes.Add , tvwChild, BuddyUserID, BuddyUserTitle, BuddyStatus, BuddyStatus

       End If
End If

End Sub

