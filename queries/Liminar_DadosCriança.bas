Operation =1
Option =0
Begin InputTables
    Name ="Ciclo"
    Name ="Criança"
End
Begin OutputColumns
    Expression ="Criança.Código"
    Expression ="Criança.Nome"
    Expression ="Criança.Nascimento"
    Expression ="Criança.Processo"
    Expression ="Criança.Status"
    Expression ="Criança.IdCiclo"
    Expression ="Ciclo.Ciclo"
End
Begin Joins
    LeftTable ="Ciclo"
    RightTable ="Criança"
    Expression ="Ciclo.Código = Criança.IdCiclo"
    Flag =3
End
Begin OrderBy
    Expression ="Criança.Nome"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Criança.Código"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Criança.Nome"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Criança.Nascimento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Criança.Processo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Criança.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Criança.IdCiclo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ciclo.Ciclo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1465
    Bottom =823
    Left =-1
    Top =-1
    Right =1449
    Bottom =517
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Ciclo"
        Name =""
    End
    Begin
        Left =53
        Top =16
        Right =197
        Bottom =160
        Top =0
        Name ="Criança"
        Name =""
    End
End
