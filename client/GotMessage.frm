VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form GotMessage 
   Caption         =   " Message from %"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GotMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5655
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown (?)"
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You've received a message from:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   3990
      Width           =   5655
      Begin VB.CommandButton Command4 
         Caption         =   "&Close"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Forward"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1090
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Reply"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"GotMessage.frx":014A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "GotMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SenderID
Public SenderName
Private Sub Command1_Click()

   Dim NewIMessage As New IMessage
   NewIMessage.Show
   
   NewIMessage.Label2.Caption = SenderName
   NewIMessage.RecieversID = SenderID

End Sub

Private Sub Command4_Click()

Unload Me

End Sub

Private Sub Form_Resize()
On Error Resume Next

RichTextBox1.Width = Me.ScaleWidth
RichTextBox1.Height = Me.ScaleHeight - Picture1.Height - Picture2.Height

End Sub


Private Sub RichTextBox1_Change()

'RichTextBox1.TextRTF = Trim(RichTextBox1.TextRTF)

End Sub


