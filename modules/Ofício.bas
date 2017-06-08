Option Explicit

'
Dim wordApp As Word.Application
Dim wordDoc As Word.Document
Dim Criança As New Criança
Dim Acompanhamento As New Acompanhamento
'
Const dirOfício = "V:\_2017\Ofício\"
'
Function NãoMatriculado(ByVal IdCriança As Long, Optional SalvarEmDoc As Boolean, Optional SalvarEmPDF As Boolean) As Boolean
NãoMatriculado = False
    '
    'Preenche o Ofício referente a criança ainda não matriculada
    'Somente liminares em Aguardo
    '
    '
    Const templateOfícioFilename = "V:\_Templates\Alunos\Liminar\Ofício - Liminar - Crianças Não Matriculadas.dotx"
    Const bookmarkNDoc = "N_Doc"
    Const bookmarkNome = "Nome"
    Const bookmarkNascimento = "Nascimento"
    Const bookmarkProcesso = "N_Processo"
    Const bookmarkOfício = "N_Ofício"
    '
    Dim OfícioName As String
    Dim NúmeroDocumento As String
    '
    NúmeroDocumento = GetNúmeroOfício
    OfícioName = GetNomeOfício
    '
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add(templateOfícioFilename)
    '
    wordApp.Visible = True
    wordApp.Activate
    '
    Criança.CarregarDadosDe (IdCriança)
    '
    InsertInField bookmarkNome, Criança.Nome
    InsertInField bookmarkNascimento, Criança.Nascimento
    InsertInField bookmarkProcesso, Criança.Processo
    InsertInField bookmarkOfício, Acompanhamento.GetOfícioDisponibilidade(IdCriança)
    '
    If (SalvarEmDoc) Then
        InsertInField bookmarkNDoc, NúmeroDocumento
        wordDoc.SaveAs2 (dirOfício & OfícioName & ".doc")
    End If
    '
    If (SalvarEmPDF) Then
        wordDoc.ExportAsFixedFormat dirOfício & OfícioName & ".pdf", wdExportFormatPDF, True
    End If
    '
    Set wordDoc = Nothing
    Set wordApp = Nothing
    '
NãoMatriculado = True
End Function

Function Desistência(ByVal IdCriança As Long, Optional SalvarEmDoc As Boolean, Optional SalvarEmPDF As Boolean) As Boolean
Desistência = False
    '
    Const templateOfícioFilename = "V:\_Templates\Alunos\Liminar\Ofício - Liminar - Desistência.dotx"
    Const bookmarkNDoc = "N_Doc"
    Const bookmarkNome = "Nome"
    Const bookmarkNascimento = "Nascimento"
    Const bookmarkProcesso = "N_Processo"
    Const bookmarkOfício = "N_Ofício"
    Const bookmarkDataDesistência = "DataDesistência"
    '
    Dim OfícioName As String
    Dim NúmeroDocumento As String
    '
    Dim DataDesistência As String
    '
    DataDesistência = InputBox("Digite a Data da Desistência", "Desistência - Liminar", Format(DateTime.Now, "dd/mm/yyyy"))
    If (EmptyTextChecker.isEmptyText(DataDesistência)) Then
        MsgBox "Necessário apontar a data da desistência", vbExclamation, "Ofício - Desistência"
        Exit Function
    End If
    '
    NúmeroDocumento = GetNúmeroOfício
    OfícioName = GetNomeOfício
    '
    Set wordApp = CreateObject("Word.Application")
    Set wordDoc = wordApp.Documents.Add(templateOfícioFilename)
    '
    wordApp.Visible = True
    wordApp.Activate
    '
    Criança.CarregarDadosDe (IdCriança)
    '
    InsertInField bookmarkNome, Criança.Nome
    InsertInField bookmarkNascimento, Criança.Nascimento
    InsertInField bookmarkProcesso, Criança.Processo
    InsertInField bookmarkOfício, Acompanhamento.GetOfícioDisponibilidade(IdCriança)
    InsertInField bookmarkDataDesistência, DataDesistência
    '
    '
    If (SalvarEmDoc) Then
        InsertInField bookmarkNDoc, NúmeroDocumento
        wordDoc.SaveAs2 (dirOfício & OfícioName & ".doc")
    End If
    '
    If (SalvarEmPDF) Then
        wordDoc.ExportAsFixedFormat dirOfício & OfícioName & ".pdf", wdExportFormatPDF, True
    End If
    '
    Set wordDoc = Nothing
    Set wordApp = Nothing
    '
Desistência = True
End Function

Sub Teste()
Desistência 139
'NãoMatriculado 171, True, True
'
Dim Db As DAO.Database
Dim recordsetPendentesOfício As Recordset
'
Set Db = CurrentDb
Set recordsetPendentesOfício = Db.OpenRecordset("tmpOfícios-13-04-2017", dbOpenDynaset)
'
If (Not recordsetPendentesOfício.BOF) Then recordsetPendentesOfício.MoveFirst
'
Do
    NãoMatriculado recordsetPendentesOfício!IdCriança, True, True
    recordsetPendentesOfício.MoveNext
Loop Until recordsetPendentesOfício.EOF
End Sub

Sub InsertInField(bookmarkName As String, Text As String)
    wordDoc.Bookmarks(bookmarkName).Range.Select
    wordApp.Selection.TypeBackspace
    wordApp.Selection.TypeText Text:=Text
End Sub

Function GetNomeOfício() As String
GetNomeOfício = GetNúmeroOfício & "-" & Format(DateTime.Now, "mm") & "-" & Format(DateTime.Now, "yyyy")
End Function

Function GetNúmeroOfício() As String
Dim NúmeroOfício As String
NúmeroOfício = GetÚltimoNúmeroOfício + 1
Debug.Print Format(DateTime.Now, "mm")
Debug.Print (NúmeroOfício)
'
Select Case Len(NúmeroOfício)
    Case 1
        NúmeroOfício = "00" & NúmeroOfício
    Case 2
        NúmeroOfício = "0" & NúmeroOfício
End Select
'
GetNúmeroOfício = NúmeroOfício
'
End Function

Function GetÚltimoNúmeroOfício() As Integer
On Error GoTo Errado
'
Dim file As String
Dim N_Ofício() As String
Dim Número As Integer
Dim NúmeroMáximo As Integer
'
file = Dir(dirOfício)
'
NúmeroMáximo = 0
Número = 0
'
While file <> ""
    N_Ofício = Split(file, "-", 2)
    '
    Número = CInt(N_Ofício(0))
    If (Número > NúmeroMáximo) Then
        NúmeroMáximo = Número
    End If
    '
    file = Dir  'Move para o Próximo arquivo
Wend
'
GetÚltimoNúmeroOfício = NúmeroMáximo
'
Errado:
    If (Err.Number = 13) Then Resume Next   'Tipos Incompatíveis (se não for possível atribuir Número)
End Function