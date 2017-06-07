VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Acompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim lIdCriança As Long
Dim sData As String
Dim sOcorrência As String
Dim sOfício As String
'
Public Property Let IdCriança(ByVal IdCriança As Long)
    lIdCriança = IdCriança
End Property

Public Property Get IdCriança() As Long
    IdCriança = lIdCriança
End Property

Public Property Let Data(ByVal Data As String)
    sData = Data
End Property

Public Property Get Data() As String
    Data = sData
End Property

Public Property Let Ocorrência(ByVal Ocorrência As String)
    sOcorrência = Ocorrência
End Property

Public Property Get Ocorrência() As String
    Ocorrência = sOcorrência
End Property

Public Property Let Ofício(ByVal Ofício As String)
    sOfício = Ofício
End Property

Public Property Get Ofício() As String
    Ofício = sOfício
End Property

Function GetOfícioDisponibilidade(ByVal IdCriança As Long) As String
'
'Obtém o Ofício de disponibilidade da Liminar
Dim db As DAO.Database
Dim recordsetAcompanhamento As Recordset
'
Set db = CurrentDb
Set recordsetAcompanhamento = db.OpenRecordset("SELECT Data, Ofício " & _
                                                "FROM Acompanhamento " & _
                                                "WHERE IdCriança = " & IdCriança & " AND Ocorrência = 'Disponibilizada'", dbOpenDynaset)
'
If (Not recordsetAcompanhamento.BOF) Then recordsetAcompanhamento.MoveFirst
'
If (recordsetAcompanhamento.RecordCount > 0) Then
    GetOfícioDisponibilidade = recordsetAcompanhamento!Ofício
End If
'
End Function

