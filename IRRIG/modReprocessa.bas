Attribute VB_Name = "modReprocessa"
Attribute VB_Name = "modReprocessa"
Option Explicit
DefInt A-Z

' Rotinas para recriar lançamentos contábeis a partir das baixas
Public Sub ReprocessarBaixas(dataAlvo As Date)
    Dim rBaixa As New GRecordSet
    Dim sqlBaixa As String

    ' Seleciona apenas as baixas da data informada
    sqlBaixa = "SELECT * FROM [Baixa Contas] WHERE [Data da Baixa] = #" & Format(dataAlvo, "mm/dd/yyyy") & "#"
    Set rBaixa = vgDb.OpenRecordSet(sqlBaixa)

    Do While Not rBaixa.EOF
        InsereLancamento rBaixa
        rBaixa.MoveNext
    Loop
End Sub

' Grava o lançamento para o registro de baixa recebido
Private Sub InsereLancamento(ByRef reg As GRecordSet)
    Dim Lancamento As New GRecordSet
    Dim Tb As New GRecordSet

    Set Lancamento = vgDb.OpenRecordSet("Lançamentos Contábil")
    Set Tb = vgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq FROM [Lançamentos Contábil]")

    With Lancamento
        .AddNew
        ![Id do Lançamento] = Tb!Seq + 1
        ![Dt do Lançamento] = Mid(reg![Data da Baixa], 1, 5)
        ![Conta Débito] = reg![Código do Débito]
        ![Conta Crédito] = reg![Código do Crédito]
        ![Valor] = reg![Valor Pago]
        ![Código do Histórico] = reg![Código do Histórico]
        ![Complemento do Histórico] = reg![Histórico]
        ![Sequência da Baixa] = reg![Sequência da Baixa]
        ![Data da Baixa] = reg![Data da Baixa]
        .Update
        .BookMark = .LastModified
    End With

    ' Inclui juros pagos na baixa
    If reg![Valor do Juros] > 0 And reg!Conta = "P" Then
        Set Tb = vgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq FROM [Lançamentos Contábil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = Mid(reg![Data da Baixa], 1, 5)
            ![Conta Débito] = 366
            ![Conta Crédito] = reg![Código do Crédito]
            ![Valor] = reg![Valor do Juros]
            ![Código do Histórico] = 181
            ![Complemento do Histórico] = reg![Histórico]
            ![Sequência da Baixa] = reg![Sequência da Baixa]
            ![Data da Baixa] = reg![Data da Baixa]
            .Update
            .BookMark = .LastModified
        End With
    End If

    ' Registra descontos concedidos na baixa
    If reg![Valor do Desconto] > 0 And reg!Conta = "P" Then
        Set Tb = vgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq FROM [Lançamentos Contábil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = Mid(reg![Data da Baixa], 1, 5)
            ![Conta Débito] = reg![Código do Crédito]
            ![Conta Crédito] = 383
            ![Valor] = reg![Valor do Desconto]
            ![Código do Histórico] = 94
            ![Complemento do Histórico] = reg![Histórico]
            ![Sequência da Baixa] = reg![Sequência da Baixa]
            ![Data da Baixa] = reg![Data da Baixa]
            .Update
            .BookMark = .LastModified
        End With
    End If

    ' Juros recebidos
    If reg![Valor do Juros] > 0 And reg!Conta = "R" Then
        Set Tb = vgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq FROM [Lançamentos Contábil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = Mid(reg![Data da Baixa], 1, 5)
            ![Conta Débito] = reg![Código do Débito]
            ![Conta Crédito] = 382
            ![Valor] = reg![Valor do Juros]
            ![Código do Histórico] = 95
            ![Complemento do Histórico] = reg![Histórico]
            ![Sequência da Baixa] = reg![Sequência da Baixa]
            ![Data da Baixa] = reg![Data da Baixa]
            .Update
            .BookMark = .LastModified
        End With
    End If

    ' Descontos recebidos
    If reg![Valor do Desconto] > 0 And reg!Conta = "R" Then
        Set Tb = vgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq FROM [Lançamentos Contábil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = Mid(reg![Data da Baixa], 1, 5)
            ![Conta Débito] = 367
            ![Conta Crédito] = reg![Código do Débito]
            ![Valor] = reg![Valor do Desconto]
            ![Código do Histórico] = 96
            ![Complemento do Histórico] = reg![Histórico]
            ![Sequência da Baixa] = reg![Sequência da Baixa]
            ![Data da Baixa] = reg![Data da Baixa]
            .Update
            .BookMark = .LastModified
        End With
    End If
End Sub

