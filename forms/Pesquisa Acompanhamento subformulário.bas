Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9317
    DatasheetFontHeight =11
    ItemSuffix =14
    Right =19950
    Bottom =12090
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x026f3338b261e440
    End
    RecordSource ="SELECT Acompanhamento.Código, Acompanhamento.IdCriança, Acompanhamento.Data, Aco"
        "mpanhamento.Ocorrência, Acompanhamento.Ofício, Acompanhamento.[Tipo Ofício], Aco"
        "mpanhamento.Documento, Acompanhamento.OBS FROM Acompanhamento; "
    Caption ="Acompanhamento subformulário"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =186
            FontSize =9
            BorderColor =11050647
            ForeColor =3881787
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
            BorderColor =11050647
        End
        Begin Line
            BorderLineStyle =0
            Width =1701
            BorderColor =11050647
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            Width =1701
            Height =1701
            BorderColor =11050647
        End
        Begin CommandButton
            TextFontCharSet =186
            Width =1701
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =3881787
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =11050647
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =11050647
        End
        Begin BoundObjectFrame
            SizeMode =3
            BorderLineStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
            BorderColor =11050647
        End
        Begin TextBox
            FELineBreak = NotDefault
            TextFontCharSet =204
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =11050647
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ListBox
            TextFontCharSet =204
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            BorderColor =11050647
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            TextFontCharSet =204
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =11050647
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderColor =11050647
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =11050647
        End
        Begin CustomControl
            OldBorderStyle =1
            Width =4536
            Height =2835
            BorderColor =11050647
        End
        Begin ToggleButton
            TextFontCharSet =186
            Width =283
            Height =283
            FontSize =9
            FontWeight =400
            ForeColor =3881787
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            TextFontCharSet =204
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =1701
            Height =1701
            BorderColor =11050647
            LabelX =-1701
        End
        Begin FormHeader
            Height =0
            BackColor =3881787
            Name ="CabeçalhoDoFormulário"
            AutoHeight =1
        End
        Begin Section
            Height =5393
            BackColor =13685460
            Name ="Detalhe"
            AutoHeight =1
            AlternateBackColor =13685460
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =345
                    Width =7260
                    Height =315
                    ColumnWidth =0
                    Name ="Código"
                    ControlSource ="Código"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =345
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =660
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =345
                            Width =1560
                            Height =315
                            Name ="Código_Rótulo"
                            Caption ="Código"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =345
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =660
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =735
                    Width =7260
                    Height =330
                    ColumnWidth =1275
                    TabIndex =1
                    Name ="IdCriança"
                    ControlSource ="IdCriança"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =735
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =1065
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =735
                            Width =1560
                            Height =330
                            Name ="IdCriança_Rótulo"
                            Caption ="IdCriança"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =735
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1065
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =1140
                    Width =7260
                    Height =330
                    ColumnWidth =1200
                    TabIndex =2
                    Name ="Data"
                    ControlSource ="Data"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =1140
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =1470
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1140
                            Width =1560
                            Height =330
                            Name ="Data_Rótulo"
                            Caption ="Data"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1470
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =1545
                    Width =7260
                    Height =600
                    ColumnWidth =1905
                    TabIndex =3
                    Name ="Ocorrência"
                    ControlSource ="Ocorrência"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =1545
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =2145
                    RowStart =3
                    RowEnd =3
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =1545
                            Width =1560
                            Height =600
                            Name ="Ocorrência_Rótulo"
                            Caption ="Ocorrência"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =1545
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =2145
                            RowStart =3
                            RowEnd =3
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =2220
                    Width =7260
                    Height =600
                    ColumnWidth =1350
                    TabIndex =4
                    Name ="OfícioSE"
                    ControlSource ="Ofício"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =2220
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =2820
                    RowStart =4
                    RowEnd =4
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =2220
                            Width =1560
                            Height =600
                            Name ="OfícioSE_Rótulo"
                            Caption ="Ofício"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =2220
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =2820
                            RowStart =4
                            RowEnd =4
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =4215
                    Width =7260
                    Height =1140
                    ColumnWidth =4635
                    TabIndex =7
                    Name ="OBS"
                    ControlSource ="OBS"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =4215
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =5355
                    RowStart =7
                    RowEnd =7
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =4215
                            Width =1560
                            Height =1140
                            Name ="OBS_Rótulo"
                            Caption ="OBS"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =4215
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =5355
                            RowStart =7
                            RowEnd =7
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1965
                    Top =2895
                    Width =7260
                    Height =315
                    ColumnWidth =2325
                    TabIndex =5
                    Name ="Tipo Ofício"
                    ControlSource ="Tipo Ofício"
                    EventProcPrefix ="Tipo_Ofício"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =2895
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =3210
                    RowStart =5
                    RowEnd =5
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =2895
                            Width =1560
                            Height =315
                            Name ="Rótulo12"
                            Caption ="Tipo Ofício:"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =2895
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =3210
                            RowStart =5
                            RowEnd =5
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    BackStyle =1
                    Left =1965
                    Top =3285
                    Width =7260
                    Height =855
                    TabIndex =6
                    Name ="Documento"
                    ControlSource ="Documento"
                    GroupTable =1
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1965
                    LayoutCachedTop =3285
                    LayoutCachedWidth =9225
                    LayoutCachedHeight =4140
                    RowStart =6
                    RowEnd =6
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =345
                            Top =3285
                            Width =1560
                            Height =855
                            Name ="Rótulo13"
                            Caption ="Documento:"
                            GroupTable =1
                            BottomPadding =38
                            LayoutCachedLeft =345
                            LayoutCachedTop =3285
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =4140
                            RowStart =6
                            RowEnd =6
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =13685460
            Name ="RodapéDoFormulário"
            AutoHeight =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
