Option Explicit

Function GetCurrentUser() As String
'
'Retorna o usuário Atual do computador
Dim someObject As Object
Set someObject = CreateObject("WScript.Network")
'
Dim textUser As String
textUser = someObject.UserName
Set someObject = Nothing
'
GetCurrentUser = textUser
End Function

Function GetCurrentPC() As String
'
'Retorna o Computador atual
Dim someObject As Object
Set someObject = CreateObject("WScript.Network")
'
Dim textPC As String
textPC = someObject.ComputerName
Set someObject = Nothing
'
GetCurrentPC = textPC
End Function