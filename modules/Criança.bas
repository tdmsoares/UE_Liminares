Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lIdCrian�a As Long
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

Public Property Let IdCrian�a(ByVal IdCrian�a As Long)
    lIdCrian�a = IdCrian�a
End Property

Public Property Get IdCrian�a() As Long
    IdCrian�a = lIdCrian�a
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

Function CarregarDadosDe(ByVal IdCrian�a As Long) As Boolean
'
'Carrega os dados das liminares
CarregarDadosDe = False
'
    Dim Db As DAO.Database
    Dim recordsetCrian�a As Recordset
    '
    Set Db = CurrentDb
    Set recordsetCrian�a = Db.OpenRecordset("SELECT C�digo,Nome,Nascimento,Processo,Status,IdCiclo " & _
                                            "FROM Crian�a " & _
                                            "WHERE C�digo = " & IdCrian�a, dbOpenDynaset)
    If (Not recordsetCrian�a.BOF) Then recordsetCrian�a.MoveFirst
    '
    If (recordsetCrian�a.RecordCount > 0) Then
        With recordsetCrian�a
            Me.IdCrian�a = !C�digo
            Me.Nome = !Nome
            Me.Nascimento = !Nascimento
            Me.Processo = !Processo
            Me.StatusLiminar = SetNomeStatus(!Status)
        End With
        CarregarDadosDe = True
    End If
End Function