Option Explicit
'
'Atualiza os Ciclos das Liminares
Function DeterminarCicloCrian�as()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCrian�a As Recordset
Dim rgCiclo As Recordset
'
Set Db = CurrentDb
Set rgCrian�a = Db.OpenRecordset("Liminar_DadosCrian�a", dbOpenDynaset)
Set rgCiclo = Db.OpenRecordset("Ciclo", dbOpenDynaset)
'
'Verifica se o foco � a posi��o Inicial na tabela
If (rgCrian�a.BOF = False) Then
    rgCrian�a.MoveFirst
End If
'
'
Dim Crian�aNascimento
Dim CicloNascidoDe
Dim CicloNascidoAt�
Dim CicloEncontrado As Boolean
'
'
Do
    '
    Crian�aNascimento = rgCrian�a!Nascimento
    CicloEncontrado = False
    '
    If (rgCiclo.BOF = False) Then
        rgCiclo.MoveFirst
    End If
    Do
        CicloNascidoAt� = rgCiclo![Nascidos At�]
        CicloNascidoDe = rgCiclo![Nascidos De]
        '
        'Faz a compara��o e verifica se a data de nascimento � a mesma que corresponde ao Ciclo
        If ((Crian�aNascimento <= CicloNascidoAt�) And (Crian�aNascimento >= CicloNascidoDe)) Then
            '
            'Atualiza o ciclo
            With rgCrian�a
                .Edit
                !IdCiclo = rgCiclo!C�digo
                .Update
                .Bookmark = .LastModified
            End With
            CicloEncontrado = True
        End If
        '
        'Move para o pr�ximo registro
        rgCiclo.MoveNext
    Loop Until rgCiclo.EOF
    '
    'Atualiza o Ciclo como nulo caso n�o tenha encontrado o correspondente
    If CicloEncontrado = False Then
        With rgCrian�a
            .Edit
            !IdCiclo = Null
            .Update
            .Bookmark = .LastModified
        End With
    End If
    '
    'Move para o pr�ximo registro
    rgCrian�a.MoveNext
Loop Until rgCrian�a.EOF
'
Errado:
If (Err.Number = 3021) Then Resume Next
End Function

Function DeterminarCiclo(Nascimento, Retorno)
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCiclo As Recordset
'
Set Db = CurrentDb
Set rgCiclo = Db.OpenRecordset("Ciclo", dbOpenDynaset)
'
Dim CicloNascidoDe
Dim CicloNascidoAt�
Dim Encontrado As Boolean
Encontrado = False
'
If (rgCiclo.BOF = False) Then
    rgCiclo.MoveFirst
End If
'
    Do
        CicloNascidoAt� = rgCiclo![Nascidos At�]
        CicloNascidoDe = rgCiclo![Nascidos De]
        '
        'Faz a compara��o e verifica se a data de nascimento � a mesma que corresponde ao Ciclo
        If ((Nascimento <= CicloNascidoAt�) And (Nascimento >= CicloNascidoDe)) Then
            MsgBox ("Nascimento da Crian�a: " & Nascimento & " est� entre " & CicloNascidoDe & " e " & CicloNascidoAt� & " Correspondendo ao Ciclo: " & rgCiclo!Ciclo)
            Encontrado = True
            If (Retorno = "C�digo") Then
                DeterminarCiclo = rgCiclo!C�digo
            ElseIf (Retorno = "NomeCiclo") Then
                DeterminarCiclo = rgCiclo!Ciclo
            End If
        End If
        '
        'Move para o pr�ximo registro
        rgCiclo.MoveNext
    Loop Until rgCiclo.EOF
    If (Encontrado = False) Then
        MsgBox ("A creche n�o atende esta faixa et�ria")
        DeterminarCiclo = ""
    End If
End Function

Function StatusLiminar(IdCrian�a)
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCrian�a As Recordset
'
Set Db = CurrentDb
Set rgCrian�a = Db.OpenRecordset("Liminar_DadosCrian�a", dbOpenDynaset)
'
Dim Status
Dim Ocorr�ncia
'
If (rgCrian�a.BOF = False) Then
    rgCrian�a.MoveFirst
End If
'
    Do
    '
    'Procura na tabela o registro da crian�a
    'Procura o Status da Liminar para essa Crian�a
        Status = rgCrian�a!Status
        If (IdCrian�a = rgCrian�a!C�digo) Then
            MsgBox ("Status Liminar: " & Status)
            '
            If (IsNull(Status)) Or (Status = "") Then
                StatusLiminar = "Pendente"
            Else:
                Select Case Status
                    Case "Aguardo"
                        StatusLiminar = "Aguardo"
                    Case "Atendida"
                        StatusLiminar = "Atendida"
                    Case "Suspensa"
                        StatusLiminar = "Suspensa"
                End Select
            End If
            
        End If
        rgCrian�a.MoveNext
    Loop Until rgCrian�a.EOF
End Function

Function CadastrarStatusLiminar(IdCrian�a, Ocorr�ncia)
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCrian�a As Recordset
'
Set Db = CurrentDb
Set rgCrian�a = Db.OpenRecordset("Liminar_DadosCrian�a", dbOpenDynaset)
'
Dim Status
'
If (rgCrian�a.BOF = False) Then
    rgCrian�a.MoveFirst
End If
'
Select Case Ocorr�ncia
    Case "Disponibilizada"
        Status = "Aguardo"
    Case "Matr�cula Efetuada"
        Status = "Atendida"
    Case "Suspensa"
        Status = "Suspensa"
    Case Else
        Exit Function
End Select
'
    Do
    '
    'Procura na tabela o registro da crian�a

        If (IdCrian�a = rgCrian�a!C�digo) Then
        '
        'Grava na tabela o Status da crian�a
            MsgBox ("Status Liminar: " & Status)
            With rgCrian�a
                .Edit
                !Status = Status
                .Update
                .Bookmark = .LastModified
            End With
        End If
        rgCrian�a.MoveNext
    Loop Until rgCrian�a.EOF
End Function