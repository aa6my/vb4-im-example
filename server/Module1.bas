Attribute VB_Name = "Module1"
Global Const MaxUsers As Integer = 100
Public UserInfo(MaxUsers) As UserStatisticalData

Public Type UserStatisticalData
     UserID As String
     UserName As String
     Nickname As String
     Password As String
     
     InRoom As String
     
     NickColor As Variant
     
     InUse As Boolean
     Status As String
End Type
Public Function TranslateStatus(StatusText As String) As String

    If StatusText = "Online" Then
    
       TranslateStatus = StatusText

    ElseIf StatusText = "Away" Then
    
       TranslateStatus = StatusText
    
    ElseIf StatusText = "DND" Then
    
       TranslateStatus = StatusText
    
    ElseIf UserInfo(u).Status = "Invisible" Then

       TranslateStatus = "Offline"
    
    ElseIf UserInfo(u).Status = "WT" Then

       TranslateStatus = "Webtour"
    
    ElseIf UserInfo(u).Status = "WTHost" Then

       TranslateStatus = "WebtourHost"
    
    Else

       ' Don't know or unsupported. Send Offline notice for now.
       TranslateStatus = "Offline"
    
    End If

End Function


