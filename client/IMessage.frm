VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form IMessage 
   Caption         =   " Send Message"
   ClientHeight    =   3840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4471
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"IMessage.frx":014A
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
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4950
      TabIndex        =   1
      Top             =   3345
      Width           =   4950
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "IMessage.frx":01CC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      TextRTF         =   $"IMessage.frx":0316
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   4950
      TabIndex        =   2
      Top             =   0
      Width           =   4950
      Begin VB.CommandButton ARButton1 
         Caption         =   "&Send"
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4440
         Shape           =   3  'Circle
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unknown"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   345
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Writing Message For:"
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   180
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   495
         Left            =   960
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Font"
      Begin VB.Menu mnuFontBold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuFontItalic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuFontUnderline 
         Caption         =   "&Underline               "
      End
      Begin VB.Menu mnuFontSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontPT 
         Caption         =   "8 pt"
         Index           =   0
      End
      Begin VB.Menu mnuFontPT 
         Caption         =   "10 pt"
         Index           =   1
      End
      Begin VB.Menu mnuFontPT 
         Caption         =   "11 pt"
         Index           =   2
      End
      Begin VB.Menu mnuFontPT 
         Caption         =   "12 pt"
         Index           =   3
      End
      Begin VB.Menu mnuFontPT 
         Caption         =   "15 pt"
         Index           =   4
      End
   End
End
Attribute VB_Name = "IMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecieversID As String
Private Sub Text1_Change()

End Sub


Private Sub ARButton1_Click()
'Dim NewBuddyName

'MsgBox RecieversID

RichTextBox1.Enabled = False

' Remove control line-feeds ('enters') for correct transmission.
RichTextBox2.Text = Replace(RichTextBox1.TextRTF, vbCrLf, "//crlf\\")

' Send message to other party.
'NewBuddyName = Replace(Label2.Caption, " ", "_._")
MyIM.Winsock1.SendData ".msg " & RecieversID & " ..//.. " & RichTextBox2.Text

RichTextBox1.Enabled = True
Unload Me

End Sub

Private Sub ARButton2_Click()

If RichTextBox1.SelBold = True Then
   mnuFontBold.Checked = True
Else
   mnuFontBold.Checked = False
End If

If RichTextBox1.SelItalic = True Then
   mnuFontItalic.Checked = True
Else
   mnuFontItalic.Checked = False
End If

If RichTextBox1.SelUnderline = True Then
   mnuFontUnderline.Checked = True
Else
   mnuFontUnderline.Checked = False
End If

PopupMenu mnuFont, , ARButton2.Left, Picture2.Top + ARButton2.Height ' + 3590

End Sub

Private Sub Command1_Click()

If RichTextBox1.SelBold = True Then
   mnuFontBold.Checked = True
Else
   mnuFontBold.Checked = False
End If

If RichTextBox1.SelItalic = True Then
   mnuFontItalic.Checked = True
Else
   mnuFontItalic.Checked = False
End If

If RichTextBox1.SelUnderline = True Then
   mnuFontUnderline.Checked = True
Else
   mnuFontUnderline.Checked = False
End If

PopupMenu mnuFont, , Command1.Left, Picture2.Top + Command1.Height

End Sub

Private Sub Form_Load()

mnuFont.Visible = False

End Sub

Private Sub Form_Resize()
On Error Resume Next

RichTextBox1.Height = Me.ScaleHeight - Picture2.Height - 850 ' - Picture1.Height ' - 500
RichTextBox1.Width = Me.ScaleWidth

Shape1.Width = Me.ScaleWidth - ARButton1.Width - 190
Shape2.Left = Shape1.Left + Shape1.Width - Shape2.Width - 120

End Sub


Private Sub mnuFontBold_Click()

If RichTextBox1.SelBold = True Then
   RichTextBox1.SelBold = False
Else
   RichTextBox1.SelBold = True
End If

End Sub

Private Sub mnuFontItalic_Click()

If RichTextBox1.SelItalic = True Then
   RichTextBox1.SelItalic = False
Else
   RichTextBox1.SelItalic = True
End If

End Sub


Private Sub mnuFontPT_Click(Index As Integer)

   RichTextBox1.SelFontSize = Word(mnuFontPT(Index).Caption, 1)

   'RichTextBox1.SelColor = vbOrange
   'RichTextBox1.SelText = Now & ": Connected closed for " & ServiceSocket(intMax).RemoteHostIP & vbCrLf
   'RichTextBox1.SelColor = vbBlack

End Sub


Private Sub mnuFontUnderline_Click()

If RichTextBox1.SelUnderline = True Then
   RichTextBox1.SelUnderline = False
Else
   RichTextBox1.SelUnderline = True
End If

End Sub


Private Sub Picture1_Click()

ARButton1_Click

End Sub


Private Sub RichTextBox1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    If RichTextBox1.SelBold = True Then
       mnuFontBold.Checked = True
    Else
       mnuFontBold.Checked = False
    End If
    
    If RichTextBox1.SelItalic = True Then
       mnuFontItalic.Checked = True
    Else
       mnuFontItalic.Checked = False
    End If
    
    If RichTextBox1.SelUnderline = True Then
       mnuFontUnderline.Checked = True
    Else
       mnuFontUnderline.Checked = False
    End If
    
    PopupMenu mnuFont
End If

End Sub


