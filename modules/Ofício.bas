Option Explicit

'
Dim wordApp As Word.Application
Dim wordDoc As Word.Document
Dim Crian�a As New Crian�a
Dim Acompanhamento As New Acompanhamento
'
Const dirOf�cio = "V:\_2017\Of�cio\"
'
Function N�oMatriculado(ByVal IdCrian�a As Long, Optional SalvarEmDoc As Boolean, Optional SalvarEmPDF As Boolean) As Boolean
N�oMatriculado = False
    '
    'Preenche o Of�cio referente a crian�a ainda n�o matriculada
    'Somente liminares em Aguardo
    '
    '
    Const templateOf�cioFilename = "V:\_Templates\Alunos\Liminar\Of�cio - Liminar - Crian�as N�o Matriculadas.dotx"
    Const bookmarkNDoc = "N_Doc"
    Const bookmarkNome = "Nome"
    Const bookmarkNascimento = "Nascimento"
    Const bookmarkProcesso = "N_Processo"
    Const bookmarkOf�cio = "N_Of�cio"
    '
    Dim Of�cioName As String
    Dim N�meroDocumento As String
    '
    N�meroDocumento = GetN�meroOf�cio
    Of�cioName = GetNomeOf�cio
    '
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add(templateOf�cioFilename)
    '
    wordApp.Visible = True
    wordApp.Activate
    '
    Crian�a.CarregarDadosDe (IdCrian�a)
    '
    InsertInField bookmarkNome, Crian�a.Nome
    InsertInField bookmarkNascimento, Crian�a.Nascimento
    InsertInField bookmarkProcesso, Crian�a.Processo
    InsertInField bookmarkOf�cio, Acompanhamento.GetOf�cioDisponibilidade(IdCrian�a)
    '
    If (SalvarEmDoc) Then
        InsertInField bookmarkNDoc, N�meroDocumento
        wordDoc.SaveAs2 (dirOf�cio & Of�cioName & ".doc")
    End If
    '
    If (SalvarEmPDF) Then
        wordDoc.ExportAsFixedFormat dirOf�cio & Of�cioName & ".pdf", wdExportFormatPDF, True
    End If
    '
    Set wordDoc = Nothing
    Set wordApp = Nothing
    '
N�oMatriculado = True
End Function

Function Desist�ncia(ByVal IdCrian�a As Long, Optional SalvarEmDoc As Boolean, Optional SalvarEmPDF As Boolean) As Boolean
Desist�ncia = False
    '
    Const templateOf�cioFilename = "V:\_Templates\Alunos\Liminar\Of�cio - Liminar - Desist�ncia.dotx"
    Const bookmarkNDoc = "N_Doc"
    Const bookmarkNome = "Nome"
    Const bookmarkNascimento = "Nascimento"
    Const bookmarkProcesso = "N_Processo"
    Const bookmarkOf�cio = "N_Of�cio"
    Const bookmarkDataDesist�ncia = "DataDesist�ncia"
    '
    Dim Of�cioName As String
    Dim N�meroDocumento As String
    '
    Dim DataDesist�ncia As String
    '
    DataDesist�ncia = InputBox("Digite a Data da Desist�ncia", "Desist�ncia - Liminar", Format(DateTime.Now, "dd/mm/yyyy"))
    If (EmptyTextChecker.isEmptyText(DataDesist�ncia)) Then
        MsgBox "Necess�rio apontar a data da desist�ncia", vbExclamation, "Of�cio - Desist�ncia"
        Exit Function
    End If
    '
    N�meroDocumento = GetN�meroOf�cio
    Of�cioName = GetNomeOf�cio
    '
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add(templateOf�cioFilename)
    '
    wordApp.Visible = True
    wordApp.Activate
    '
    Crian�a.CarregarDadosDe (IdCrian�a)
    '
    InsertInField bookmarkNome, Crian�a.Nome
    InsertInField bookmarkNascimento, Crian�a.Nascimento
    InsertInField bookmarkProcesso, Crian�a.Processo
    InsertInField bookmarkOf�cio, Acompanhamento.GetOf�cioDisponibilidade(IdCrian�a)
    InsertInField bookmarkDataDesist�ncia, DataDesist�ncia
    '
    '
    If (SalvarEmDoc) Then
        InsertInField bookmarkNDoc, N�meroDocumento
        wordDoc.SaveAs2 (dirOf�cio & Of�cioName & ".doc")
    End If
    '
    If (SalvarEmPDF) Then
        wordDoc.ExportAsFixedFormat dirOf�cio & Of�cioName & ".pdf", wdExportFormatPDF, True
    End If
    '
    Set wordDoc = Nothing
    Set wordApp = Nothing
    '
Desist�ncia = True
End Function

Sub Teste()
Desist�ncia 139
'N�oMatriculado 171, True, True
'
Dim Db As DAO.Database
Dim recordsetPendentesOf�cio As Recordset
'
Set Db = CurrentDb
Set recordsetPendentesOf�cio = Db.OpenRecordset("tmpOf�cios-13-04-2017", dbOpenDynaset)
'
If (Not recordsetPendentesOf�cio.BOF) Then recordsetPendentesOf�cio.MoveFirst
'
Do
    N�oMatriculado recordsetPendentesOf�cio!IdCrian�a, True, True
    recordsetPendentesOf�cio.MoveNext
Loop Until recordsetPendentesOf�cio.EOF
End Sub

Sub InsertInField(bookmarkName As String, Text As String)
    wordDoc.Bookmarks(bookmarkName).Range.Select
    wordApp.Selection.TypeBackspace
    wordApp.Selection.TypeText Text:=Text
End Sub

Function GetNomeOf�cio() As String
GetNomeOf�cio = GetN�meroOf�cio & "-" & Format(DateTime.Now, "mm") & "-" & Format(DateTime.Now, "yyyy")
End Function

Function GetN�meroOf�cio() As String
Dim N�meroOf�cio As String
N�meroOf�cio = Get�ltimoN�meroOf�cio + 1
Debug.Print Format(DateTime.Now, "mm")
Debug.Print (N�meroOf�cio)
'
Select Case Len(N�meroOf�cio)
    Case 1
        N�meroOf�cio = "00" & N�meroOf�cio
    Case 2
        N�meroOf�cio = "0" & N�meroOf�cio
End Select
'
GetN�meroOf�cio = N�meroOf�cio
'
End Function

Function Get�ltimoN�meroOf�cio() As Integer
On Error GoTo Errado
'
Dim file As String
Dim N_Of�cio() As String
Dim N�mero As Integer
Dim N�meroM�ximo As Integer
'
file = Dir(dirOf�cio)
'
N�meroM�ximo = 0
N�mero = 0
'
While file <> ""
    N_Of�cio = Split(file, "-", 2)
    '
    N�mero = CInt(N_Of�cio(0))
    If (N�mero > N�meroM�ximo) Then
        N�meroM�ximo = N�mero
    End If
    '
    file = Dir  'Move para o Pr�ximo arquivo
Wend
'
Get�ltimoN�meroOf�cio = N�meroM�ximo
'
Errado:
    If (Err.Number = 13) Then Resume Next   'Tipos Incompat�veis (se n�o for poss�vel atribuir N�mero)
End Function