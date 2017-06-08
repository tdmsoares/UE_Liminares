Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lIdCriança As Long
Dim sNome As String
Dim sNascimento As String
Dim sProcesso As String
Dim nStatus As Integer
'
Public Enum Status
    Aguardo = 1
    Atendida = 2
    Suspensa = 3
End Enum

Public Property Let IdCriança(ByVal IdCriança As Long)
    lIdCriança = IdCriança
End Property

Public Property Get IdCriança() As Long
    IdCriança = lIdCriança
End Property

Public Property Let Nome(ByVal Nome As String)
    sNome = Nome
End Property

Public Property Get Nome() As String
    Nome = sNome
End Property

Public Property Let Nascimento(ByVal Nascimento As String)
    sNascimento = Nascimento
End Property

Public Property Get Nascimento() As String
    Nascimento = sNascimento
End Property

Public Property Let Processo(ByVal Processo As String)
    sProcesso = Processo
End Property

Public Property Get Processo() As String
    Processo = sProcesso
End Property

Public Property Let StatusLiminar(ByVal StatusLiminar As Status)
    nStatus = StatusLiminar
End Property

Public Property Get StatusLiminar() As Status
    StatusLiminar = nStatus
End Property

Function GetNomeStatus(ByVal StatusLiminar As Status) As String
    Select Case StatusLiminar
        Case Aguardo
            GetNomeStatus = "Aguardo"
        Case Atendida
            GetNomeStatus = "Atendida"
        Case Suspensa
            GetNomeStatus = "Suspensa"
    End Select
End Function

Function SetNomeStatus(ByVal StatusLiminar As String) As Integer
    Select Case StatusLiminar
        Case "Aguardo"
            SetNomeStatus = Aguardo
        Case Atendida
            SetNomeStatus = Atendida
        Case Suspensa
            SetNomeStatus = Suspensa
    End Select
End Function

Function CarregarDadosDe(ByVal IdCriança As Long) As Boolean
'
'Carrega os dados das liminares
CarregarDadosDe = False
'
    Dim Db As DAO.Database
    Dim recordsetCriança As Recordset
    '
    Set Db = CurrentDb
    Set recordsetCriança = Db.OpenRecordset("SELECT Código,Nome,Nascimento,Processo,Status,IdCiclo " & _
                                            "FROM Criança " & _
                                            "WHERE Código = " & IdCriança, dbOpenDynaset)
    If (Not recordsetCriança.BOF) Then recordsetCriança.MoveFirst
    '
    If (recordsetCriança.RecordCount > 0) Then
        With recordsetCriança
            Me.IdCriança = !Código
            Me.Nome = !Nome
            Me.Nascimento = !Nascimento
            Me.Processo = !Processo
            Me.StatusLiminar = SetNomeStatus(!Status)
        End With
        CarregarDadosDe = True
    End If
End Function