Attribute VB_Name = "Module1"
Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
  ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

' =========================

Public AccountStats As bAccount
Public WebBrowserIsOpen As Boolean

Type bAccount
     ChatHost As Boolean
     ChatStyle As Integer

     NickColor As Variant
     FrameAni As Integer
     OldFrameAni As Integer

     BuddyStatus As String
End Type

Global Const MaxChats As Integer = 5
Public Chatroom(MaxChats) As Chatroomdata

Type Chatroomdata
     InUse As Boolean

     Title As String
     Description As String

     AdminID As String

     AttachedURL As String
End Type
Public Sub CloseAllForms(frmOwner As Form)
     '========================================================
     'Closes all forms in the app, closing the owner form last
     '========================================================
     Dim frm As Form
     On Error GoTo CloseAllForms_Err
     For Each frm In Forms
       If Not frm Is frmOwner Then
           Unload frm
           Set frm = Nothing
       End If
     Next frm
     Unload frmOwner
     Set frmOwner = Nothing
CloseAllForms_Exit:
     Exit Sub
CloseAllForms_Err:
     MsgBox "Error: " & Err.Number & " " & Err.Description, vbInformation, "CloseAllForms"
     Resume CloseAllForms_Exit
End Sub
Public Sub Main()

If Online() = True Then

   MDIForm1.Show

Else

   MsgBox "You are currently not connected to the Internet or a LAN." & vbCrLf & vbCrLf & "You will not be able to access any of BinaryPoint's Internet-enabled features during this session.", vbInformation
   MDIForm1.Show

End If

End Sub


