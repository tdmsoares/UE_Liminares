Option Explicit
'
'Atualiza os Ciclos das Liminares
Function DeterminarCicloCrianças()
On Error GoTo Errado
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCriança As Recordset
Dim rgCiclo As Recordset
'
Set Db = CurrentDb
Set rgCriança = Db.OpenRecordset("Liminar_DadosCriança", dbOpenDynaset)
Set rgCiclo = Db.OpenRecordset("Ciclo", dbOpenDynaset)
'
'Verifica se o foco é a posição Inicial na tabela
If (rgCriança.BOF = False) Then
    rgCriança.MoveFirst
End If
'
'
Dim CriançaNascimento
Dim CicloNascidoDe
Dim CicloNascidoAté
Dim CicloEncontrado As Boolean
'
'
Do
    '
    CriançaNascimento = rgCriança!Nascimento
    CicloEncontrado = False
    '
    If (rgCiclo.BOF = False) Then
        rgCiclo.MoveFirst
    End If
    Do
        CicloNascidoAté = rgCiclo![Nascidos Até]
        CicloNascidoDe = rgCiclo![Nascidos De]
        '
        'Faz a comparação e verifica se a data de nascimento é a mesma que corresponde ao Ciclo
        If ((CriançaNascimento <= CicloNascidoAté) And (CriançaNascimento >= CicloNascidoDe)) Then
            '
            'Atualiza o ciclo
            With rgCriança
                .Edit
                !IdCiclo = rgCiclo!Código
                .Update
                .Bookmark = .LastModified
            End With
            CicloEncontrado = True
        End If
        '
        'Move para o próximo registro
        rgCiclo.MoveNext
    Loop Until rgCiclo.EOF
    '
    'Atualiza o Ciclo como nulo caso não tenha encontrado o correspondente
    If CicloEncontrado = False Then
        With rgCriança
            .Edit
            !IdCiclo = Null
            .Update
            .Bookmark = .LastModified
        End With
    End If
    '
    'Move para o próximo registro
    rgCriança.MoveNext
Loop Until rgCriança.EOF
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
Dim CicloNascidoAté
Dim Encontrado As Boolean
Encontrado = False
'
If (rgCiclo.BOF = False) Then
    rgCiclo.MoveFirst
End If
'
    Do
        CicloNascidoAté = rgCiclo![Nascidos Até]
        CicloNascidoDe = rgCiclo![Nascidos De]
        '
        'Faz a comparação e verifica se a data de nascimento é a mesma que corresponde ao Ciclo
        If ((Nascimento <= CicloNascidoAté) And (Nascimento >= CicloNascidoDe)) Then
            MsgBox ("Nascimento da Criança: " & Nascimento & " está entre " & CicloNascidoDe & " e " & CicloNascidoAté & " Correspondendo ao Ciclo: " & rgCiclo!Ciclo)
            Encontrado = True
            If (Retorno = "Código") Then
                DeterminarCiclo = rgCiclo!Código
            ElseIf (Retorno = "NomeCiclo") Then
                DeterminarCiclo = rgCiclo!Ciclo
            End If
        End If
        '
        'Move para o próximo registro
        rgCiclo.MoveNext
    Loop Until rgCiclo.EOF
    If (Encontrado = False) Then
        MsgBox ("A creche não atende esta faixa etária")
        DeterminarCiclo = ""
    End If
End Function

Function StatusLiminar(IdCriança)
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCriança As Recordset
'
Set Db = CurrentDb
Set rgCriança = Db.OpenRecordset("Liminar_DadosCriança", dbOpenDynaset)
'
Dim Status
Dim Ocorrência
'
If (rgCriança.BOF = False) Then
    rgCriança.MoveFirst
End If
'
    Do
    '
    'Procura na tabela o registro da criança
    'Procura o Status da Liminar para essa Criança
        Status = rgCriança!Status
        If (IdCriança = rgCriança!Código) Then
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
        rgCriança.MoveNext
    Loop Until rgCriança.EOF
End Function

Function CadastrarStatusLiminar(IdCriança, Ocorrência)
'
'Atribui as tabelas
Dim Db As DAO.Database
Dim rgCriança As Recordset
'
Set Db = CurrentDb
Set rgCriança = Db.OpenRecordset("Liminar_DadosCriança", dbOpenDynaset)
'
Dim Status
'
If (rgCriança.BOF = False) Then
    rgCriança.MoveFirst
End If
'
Select Case Ocorrência
    Case "Disponibilizada"
        Status = "Aguardo"
    Case "Matrícula Efetuada"
        Status = "Atendida"
    Case "Suspensa"
        Status = "Suspensa"
    Case Else
        Exit Function
End Select
'
    Do
    '
    'Procura na tabela o registro da criança

        If (IdCriança = rgCriança!Código) Then
        '
        'Grava na tabela o Status da criança
            MsgBox ("Status Liminar: " & Status)
            With rgCriança
                .Edit
                !Status = Status
                .Update
                .Bookmark = .LastModified
            End With
        End If
        rgCriança.MoveNext
    Loop Until rgCriança.EOF
End Function