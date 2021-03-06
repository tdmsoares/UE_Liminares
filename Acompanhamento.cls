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

Dim lIdCrian�a As Long
Dim sData As String
Dim sOcorr�ncia As String
Dim sOf�cio As String
'
Public Property Let IdCrian�a(ByVal IdCrian�a As Long)
    lIdCrian�a = IdCrian�a
End Property

Public Property Get IdCrian�a() As Long
    IdCrian�a = lIdCrian�a
End Property

Public Property Let Data(ByVal Data As String)
    sData = Data
End Property

Public Property Get Data() As String
    Data = sData
End Property

Public Property Let Ocorr�ncia(ByVal Ocorr�ncia As String)
    sOcorr�ncia = Ocorr�ncia
End Property

Public Property Get Ocorr�ncia() As String
    Ocorr�ncia = sOcorr�ncia
End Property

Public Property Let Of�cio(ByVal Of�cio As String)
    sOf�cio = Of�cio
End Property

Public Property Get Of�cio() As String
    Of�cio = sOf�cio
End Property

Function GetOf�cioDisponibilidade(ByVal IdCrian�a As Long) As String
'
'Obt�m o Of�cio de disponibilidade da Liminar
Dim db As DAO.Database
Dim recordsetAcompanhamento As Recordset
'
Set db = CurrentDb
Set recordsetAcompanhamento = db.OpenRecordset("SELECT Data, Of�cio " & _
                                                "FROM Acompanhamento " & _
                                                "WHERE IdCrian�a = " & IdCrian�a & " AND Ocorr�ncia = 'Disponibilizada'", dbOpenDynaset)
'
If (Not recordsetAcompanhamento.BOF) Then recordsetAcompanhamento.MoveFirst
'
If (recordsetAcompanhamento.RecordCount > 0) Then
    GetOf�cioDisponibilidade = recordsetAcompanhamento!Of�cio
End If
'
End Function

