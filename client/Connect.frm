VERSION 5.00
Begin VB.Form Connect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Connect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "127.0.0.1"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2025
      TabIndex        =   7
      Top             =   3765
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proceed »"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3765
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Connect.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
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
      Left            =   480
      TabIndex        =   9
      Top             =   3045
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2565
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2205
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Connect.frx":27AE
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   3105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   1920
      Width           =   4215
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Command1.Enabled = False
Command2.Caption = "&Cancel"

Text1.Enabled = False
Text2.Enabled = False

Label3.ForeColor = &H808080
Label4.ForeColor = &H808080

Label1.Caption = "Connecting..."
Label2.Caption = "Great! Now all you have to do is sit back and relax, I'll connect to the server."

MyIM.Winsock1.RemoteHost = Text3.Text
MyIM.Winsock1.Connect

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()

If Command2.Caption = "&Cancel" Then

   Label1.Caption = "Connection Canceled."
   Label2.Caption = "Ok, I've stopped trying to connect. When your ready to try again, hit the Proceed button."

      Label3.ForeColor = vbBlack
      Label4.ForeColor = vbBlack
      Text1.Enabled = True
      Text2.Enabled = True
      Command1.Enabled = True
      Command2.Caption = "&Close"
      MyIM.Winsock1.Close
   
   Command2.Caption = "&Close"

Else

   Unload Me
   Unload MyIM
   End

End If

End Sub

Private Sub Timer1_Timer()

If MyIM.Winsock1.State = sckConnected Then
   Label1.Caption = "Verifying ID and Password..."
   MyIM.Winsock1.SendData ".login " & Text1 & " " & Text2
   Timer1.Enabled = False
End If

End Sub


