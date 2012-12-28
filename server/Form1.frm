VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MyInstant Messenger Server"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Event Log"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin MSWinsockLib.Winsock ServiceSocket 
         Index           =   0
         Left            =   4080
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   6000
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5530
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":27A2
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 /200"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   3
      Top             =   510
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# Connected Users:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intMax As Integer
Public Account_Ref_Path As String

Private m_cIni As New cInifile
Private Sub Command1_Click()
   Unload Me
   End
End Sub

Private Sub Form_Load()

   Label3.Caption = "0 /" & MaxUsers

   ServiceSocket(0).Listen

   Account_Ref_Path = App.Path & "\data\"

End Sub

Private Sub RichTextBox1_Change()

   RichTextBox1.SelStart = Len(RichTextBox1)

End Sub


Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

RichTextBox1.SelStart = Len(RichTextBox1)

End Sub

Private Sub ServiceSocket_Close(Index As Integer)

   ServiceSocket(Index).Close
   UserInfo(Index).InUse = False
   UserInfo(Index).InRoom = ""
   UserInfo(Index).NickColor = ""
   UserInfo(Index).Nickname = ""
   UserInfo(Index).Password = ""
   UserInfo(Index).Status = ""
   UserInfo(Index).UserID = ""
   UserInfo(Index).UserName = ""

   RichTextBox1.SelColor = vbOrange
   RichTextBox1.SelText = Now & ": Connected closed for " & ServiceSocket(Index).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = vbBlack

   Label3.Caption = Word(Label3.Caption, 1) - 1 & " /200"

End Sub

Private Sub ServiceSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        intMax = intMax + 1
        Load ServiceSocket(intMax)

        ServiceSocket(intMax).Accept requestID

        RichTextBox1.SelColor = vbOrange
        RichTextBox1.SelText = Now & ": New connection request from " & ServiceSocket(intMax).RemoteHostIP & vbCrLf
        RichTextBox1.SelColor = vbBlack
    End If
End Sub

Private Sub ServiceSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim UserCommand As String
Dim i
Dim BadAccount As Boolean
Dim IMMyName
ServiceSocket(Index).GetData UserCommand


   If Word(UserCommand, 1) = ".login" Then

      With m_cIni
         .Path = Account_Ref_Path & "reference.txt"
         .Section = "Accounts"
         .Key = Word(UserCommand, 2)
         .Default = "Invalid"
        
         sSavePath = .Value
         If Not (.Success) Then
            BadAccount = True
         End If
      End With

      If sSavePath = "Invalid" Then
         ServiceSocket(Index).SendData ".loginbad 0"
      Else
         If sSavePath = Word(UserCommand, 3) Then
            ServiceSocket(Index).SendData ".loginok"
            Label3.Caption = Word(Label3.Caption, 1) + 1 & " /200"

            RichTextBox1.SelColor = &H80FF&
            RichTextBox1.SelText = Now & ": User " & Word(UserCommand, 2) & " logged in from " & ServiceSocket(intMax).RemoteHostIP & vbCrLf
            RichTextBox1.SelColor = vbBlack

            With m_cIni
               .Path = Account_Ref_Path & Word(UserCommand, 2) & ".txt"
               .Section = "Info"
               .Key = "Nick"
               .Default = "Invalid"
              
               UserInfo(Index).Nickname = .Value
            End With

            UserInfo(Index).UserID = Word(UserCommand, 2)
            UserInfo(Index).Status = "Online"
            
            UserInfo(Index).InUse = True
         Else
            ServiceSocket(Index).SendData ".loginbad 1"
         End If
      End If


   ElseIf Word(UserCommand, 1) = ".status" Then
   
      UserInfo(Index).Status = Word(UserCommand, 2)
      If UserInfo(Index).Status = "Invisible" Then UserInfo(Index).Status = "Offline"


   ElseIf Word(UserCommand, 1) = ".getstatus" Then

            For u = 0 To MaxUsers
              If UserInfo(u).UserID <> "" And UserInfo(u).UserID = Word(UserCommand, 2) Then
                If ServiceSocket(u).State = 7 Then

                      ServiceSocket(Index).SendData ".pushbuddyupdate " & UserInfo(u).UserID & " " & TranslateStatus(UserInfo(u).Status)

                Else

                      ServiceSocket(Index).SendData ".pushbuddyupdate " & UserInfo(u).UserID & " Offline"

                End If
              End If
            Next

   ElseIf Word(UserCommand, 1) = ".msg" Then

      IMMyName = Replace(UserInfo(Index).Nickname, " ", "_._")

            For u = 0 To MaxUsers
              If UserInfo(u).UserID = Word(UserCommand, 2) Then
                If ServiceSocket(u).State = 7 Then

                      timedPause 1
                      ServiceSocket(u).SendData ".msg " & UserInfo(Index).UserID & " " & IMMyName & " ..//.. " & SplitString(UserCommand, "..//..")
                      Exit For

                End If
              End If
            Next

   ElseIf Word(UserCommand, 1) = ".getbuddys" Then
    Dim TotalBuddys
    Dim BuddyUserStatus
    Dim BuddyUserID
    Dim BuddyUserTitle
    Dim BuddyUserAccID
    Dim BuddyUserFound As Boolean

            With m_cIni
               .Path = Account_Ref_Path & UserInfo(Index).UserID & "_bl" & ".txt"
               .Section = "Buddylist"
               .Key = "Total"
               .Default = "0"
              
               TotalBuddys = .Value
            End With

    If TotalBuddys > 0 Then
      For i = 1 To TotalBuddys
        If ServiceSocket(Index).State = 7 Then

            With m_cIni
               .Path = Account_Ref_Path & UserInfo(Index).UserID & "_bl" & ".txt"
               .Section = "Buddy_" & i
               .Key = "UserID"
               .Default = ""
              
               BuddyUserID = .Value

               .Section = "Buddy_" & i
               .Key = "Title"
               .Default = ""
              
               BuddyUserTitle = .Value
            End With

            timedPause 1

                  For u = 0 To MaxUsers
                  BuddyUserFound = False
                  
                    If UserInfo(u).UserID = BuddyUserID Then
                    If Not UserInfo(u).UserID = UserInfo(Index).UserID Then
                      If ServiceSocket(u).State = 7 Then

                         BuddyUserFound = True
                         ServiceSocket(Index).SendData ".pushbuddy " & TranslateStatus(UserInfo(u).Status) & " " & BuddyUserID & " " & BuddyUserTitle
                         Exit For

                      Else
                      
                         BuddyUserFound = True
                         ServiceSocket(Index).SendData ".pushbuddy Offline " & BuddyUserID & " " & BuddyUserTitle
                         Exit For
                      
                      End If
                    End If
                    End If
                  Next

            If BuddyUserFound = False Then
               ServiceSocket(Index).SendData ".pushbuddy Offline " & BuddyUserID & " " & BuddyUserTitle
            End If

            timedPause 1

        End If
      Next
    End If

   End If

End Sub

Private Sub ServiceSocket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

   RichTextBox1.SelColor = vbRed
   RichTextBox1.SelText = Now & ": Encountered error #" & Number & ", " & Description & ", from " & ServiceSocket(intMax).RemoteHostIP & vbCrLf
   RichTextBox1.SelColor = vbBlack

End Sub


