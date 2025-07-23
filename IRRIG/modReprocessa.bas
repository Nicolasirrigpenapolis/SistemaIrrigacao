Attribute VB_Name = "modReprocessa"

Option Explicit
DefInt A-Z

' Rotinas para recriar lançamentos contábeis a partir das baixas
Public Sub ReprocessarBaixas(dataAlvo As Date)
    Dim rBaixa   As New GRecordSet
    Dim sqlBaixa As String
    Dim sData    As String

    ' Formata data em ISO e usa aspas simples (compatível com SQL Server)
    sData = Format(dataAlvo, "yyyy-mm-dd")

    sqlBaixa = _
      "SELECT * FROM [Baixa Contas] " & _
      "WHERE [Data da Baixa] = '" & sData & "'"
    Set rBaixa = vgDb.OpenRecordSet(sqlBaixa)

    Do While Not rBaixa.EOF
        InsereLancamento rBaixa
        rBaixa.MoveNext
    Loop
End Sub

' Grava o lançamento para o registro de baixa recebido
Private Sub InsereLancamento(ByRef reg As GRecordSet)
    Dim Lancamento As New GRecordSet
    Dim Tb         As New GRecordSet

    ' Abre a tabela correta
    Set Lancamento = vgDb.OpenRecordSet("[Lançamentos Contabil]")
    Set Tb = vgDb.OpenRecordSet( _
      "SELECT MAX([Id do Lançamento]) AS Seq FROM [Lançamentos Contabil]")

    ' --- Lançamento principal ---
    With Lancamento
        .AddNew
        ![Id do Lançamento] = Tb!Seq + 1
        ![Dt do Lançamento] = reg![Data da Baixa]
        ![Conta Debito] = reg![Codigo do Debito]
        ![Conta Credito] = reg![Codigo do Credito]
        ![Valor] = reg![Valor Pago]
        ![Codigo do Historico] = reg![Codigo do Historico]
        ![Complemento do Hist] = reg![Complemento do Hist]
        ![Seqüência da Baixa] = reg![Seqüência da Baixa]
        ![Seqüência da Movimentação CC] = reg![Seqüência da Movimentação CC]
        ![Data da Baixa] = reg![Data da Baixa]
        ![Gerado] = False
        .Update
    End With

    ' --- Juros pagos na baixa (conta P) ---
    If reg![Valor do Juros] > 0 And reg!Conta = "P" Then
        Set Tb = vgDb.OpenRecordSet( _
          "SELECT MAX([Id do Lançamento]) AS Seq FROM [Lançamentos Contabil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = reg![Data da Baixa]
            ![Conta Debito] = 366
            ![Conta Credito] = reg![Codigo do Credito]
            ![Valor] = reg![Valor do Juros]
            ![Codigo do Historico] = 181
            ![Complemento do Hist] = reg![Complemento do Hist]
            ![Seqüência da Baixa] = reg![Seqüência da Baixa]
            ![Seqüência da Movimentação CC] = reg![Seqüência da Movimentação CC]
            ![Data da Baixa] = reg![Data da Baixa]
            ![Gerado] = False
            .Update
        End With
    End If

    ' --- Descontos concedidos na baixa (conta P) ---
    If reg![Valor do Desconto] > 0 And reg!Conta = "P" Then
        Set Tb = vgDb.OpenRecordSet( _
          "SELECT MAX([Id do Lançamento]) AS Seq FROM [Lançamentos Contabil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = reg![Data da Baixa]
            ![Conta Debito] = reg![Codigo do Credito]
            ![Conta Credito] = 383
            ![Valor] = reg![Valor do Desconto]
            ![Codigo do Historico] = 94
            ![Complemento do Hist] = reg![Complemento do Hist]
            ![Seqüência da Baixa] = reg![Seqüência da Baixa]
            ![Seqüência da Movimentação CC] = reg![Seqüência da Movimentação CC]
            ![Data da Baixa] = reg![Data da Baixa]
            ![Gerado] = False
            .Update
        End With
    End If

    ' --- Juros recebidos (conta R) ---
    If reg![Valor do Juros] > 0 And reg!Conta = "R" Then
        Set Tb = vgDb.OpenRecordSet( _
          "SELECT MAX([Id do Lançamento]) AS Seq FROM [Lançamentos Contabil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = reg![Data da Baixa]
            ![Conta Debito] = reg![Codigo do Debito]
            ![Conta Credito] = 382
            ![Valor] = reg![Valor do Juros]
            ![Codigo do Historico] = 95
            ![Complemento do Hist] = reg![Complemento do Hist]
            ![Seqüência da Baixa] = reg![Seqüência da Baixa]
            ![Seqüência da Movimentação CC] = reg![Seqüência da Movimentação CC]
            ![Data da Baixa] = reg![Data da Baixa]
            ![Gerado] = False
            .Update
        End With
    End If

    ' --- Descontos recebidos (conta R) ---
    If reg![Valor do Desconto] > 0 And reg!Conta = "R" Then
        Set Tb = vgDb.OpenRecordSet( _
          "SELECT MAX([Id do Lançamento]) AS Seq FROM [Lançamentos Contabil]")
        With Lancamento
            .AddNew
            ![Id do Lançamento] = Tb!Seq + 1
            ![Dt do Lançamento] = reg![Data da Baixa]
            ![Conta Debito] = 367
            ![Conta Credito] = reg![Codigo do Debito]
            ![Valor] = reg![Valor do Desconto]
            ![Codigo do Historico] = 96
            ![Complemento do Hist] = reg![Complemento do Hist]
            ![Seqüência da Baixa] = reg![Seqüência da Baixa]
            ![Seqüência da Movimentação CC] = reg![Seqüência da Movimentação CC]
            ![Data da Baixa] = reg![Data da Baixa]
            ![Gerado] = False
            .Update
        End With
    End If
End Sub


