Attribute VB_Name = "modOrcamentoLogic"
Option Explicit
DefInt A-Z


Public vgSituacao As Integer                      'situao de edio que do mdulo
Public vgCaracteristica As Integer                'caracteristica do mdulo
Public vgTipo As Integer                          'tipo do mdulo
Public vgFiltroInicial As String                  'filtro inicial do mdulo
Public vgOrdemInicial As String                   'ordem inicial do mdulo
Public vgUltimaOrdem As String                    'ltima ordenao feita no mdulo
Public vgUltimoFiltro As String                   'ltimo filtro definido no mdulo
Public vgUltimoFiltroComTit As String             'titulo do ltimo filtro definido no mdulo
Public vgUltimaOrdemComTit As String              'titulo da ltima ordenao feita no mdulo
Public vgUltimoTabIndex As Integer                'ltimo campo com foco do mdulo
Public vgPriVez As Integer                        'flag de carregamento do mdulo
Public WithEvents vgTb As GRecordSet              'tabela de dados do mdulo
Attribute vgTb.VB_VarHelpID = -1
Public vgSQL As String                            'expresso SQL que define o mdulo
Public vgTemInclusao As Integer                   'flag se tem ou no incluso no mdulo
Public vgTemExclusao As Integer                   'flag se tem ou no excluso no mdulo
Public vgTemProcura As Integer                    'flag se tem ou no procura no mdulo
Public vgTemFiltro As Integer                     'flag se tem ou no filtro no mdulo
Public vgTemAlteracao As Integer                  'flag se tem ou no alterao no mdulo
Public vgTemAlteracaoGrids As Integer              'flag se tem ou no alterao nos grids
Public vgTemBrowse As Integer                     'flag se tem ou no janela em grade no mdulo
Public vgSemVincDados As Integer                  'Flag para definir formulrios sem vinculo com dados
Public vgEmBrowse As Integer                      'flag se o mdulo esta em grade
Public vgRepeticao As Integer                     'flag de repetio do ltimo reg digitado
Public vgAlterar As Integer                       'flag de Alteracao de registros
Public vgUltAlterar As Integer                    'flag de ltima situao de "pode alterar"
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso
Public vgIndDefault As String                     'indice default do mdulo
Public vgFormID As Long                           'identificador nico para o mdulo
Public vgIdentTab As String                       'nome da tabela principal do mdulo
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do mdulo
Public vgFiltroOriginal As String
Private defaultBack55 As Long
Dim txtCampo(161) As New FormataCampos            'classe dos campos tipo texto do mdulo
Dim chkCampo(11) As New FormataCampos             'classe dos campos tipo lgico do mdulo
Dim vgPodeFazerUnLoad As Boolean                  'flag se o mdulo pode ou nao ser removido da memria
Dim opcPainel1(1) As New FormataCampos
Dim opcPainel2(1) As New FormataCampos
Dim Cancelado As Boolean, Data_da_Alteracao As Variant, Hora_da_Alteracao As Variant
Dim Usuario_da_Alteracao As String, Venda_Fechada As Boolean, Valor_Total_IPI_das_Pecas As Double
Dim Valor_Total_das_Pecas As Double, Sequencia_do_Municipio As Long, Sequencia_do_Pais As Long
Dim Sequencia_do_Orcamento As Long, Sequencia_do_Geral As Long, Observacao As String
Dim Fechamento As Integer, Valor_do_Fechamento As Double, Valor_Total_IPI_dos_Produtos As Double
Dim Valor_Total_IPI_dos_Conjuntos As Double, Valor_Total_do_ICMS As Double, Valor_Total_dos_Produtos As Double
Dim Valor_Total_dos_Conjuntos As Double, Valor_Total_de_Produtos_Usados As Double, Valor_Total_Conjuntos_Usados As Double
Dim Valor_Total_dos_Servicos As Double, Valor_Total_do_Orcamento As Double, Nome_Cliente As String
Dim endereco As String, CEP As String, Telefone As String
Dim Fax As String, Email As String, Sequencia_do_Vendedor As Long
Dim Sequencia_do_Pedido As Long, Tipo As Integer, CPF_e_CNPJ As String
Dim RG_e_IE As String, Forma_de_Pagamento As String, Ocultar_Valor_Unitario As Boolean
Dim Sequencia_da_Classificacao As Integer, Bairro As String, Caixa_Postal As String
Dim e_Propriedade As Boolean, Nome_da_Propriedade As String, Numero_do_Endereco As String
Dim Valor_Total_da_Base_de_Calculo As Double, Valor_do_Seguro As Double, Valor_do_Frete As Double
Dim Valor_Total_das_Pecas_Usadas As Double, Sequencia_da_Propriedade As Integer, Complemento As String
Dim Data_de_Emissao As Variant, Data_do_Fechamento As Variant, Codigo_do_Suframa As String
Dim Revenda As Boolean, Valor_Total_do_Tributo As Double, Valor_Total_do_PIS As Double
Dim Valor_Total_do_COFINS As Double, Valor_Total_da_Base_ST As Double, Valor_Total_do_ICMS_ST As Double
Dim Aliquota_do_ISS As Single, Reter_ISS As Boolean, Fatura_Proforma As Boolean
Dim Entrega_Futura As Boolean, Sequencia_da_Transportadora As Long, Orcamento_Avulso As Boolean
Dim Valor_do_Imposto_de_Renda As Double, Local_de_Embarque As String, UF_de_Embarque As String
Dim Numero_da_Proforma As Long, Conjunto_Avulso As Boolean, Descricao_Conjunto_Avulso As String
Dim Vendedor_Intermediario As String, Percentual_do_Vendedor As Double, Rebiut As String
Dim Percentual_Rebiut As Double, Nao_Movimentar_Estoque As Boolean, Gerou_Encargos As Boolean
Dim Peso_Bruto As Double, Peso_Liquido As Double, Volumes As Long
Dim Aviso_de_embarque As String, Hidroturbo As String, Area_irrigada As Double
Dim Precipitacao_bruta As Double, Horas_irrigada As Double, Area_tot_irrigada_em As Double
Dim Aspersor As String, Modelo_do_aspersor As String, Bocal_diametro As Double
Dim Pressao_de_servico As Double, Alcance_do_jato As Double, Espaco_entre_carreadores As Double
Dim Faixa_irrigada As Double, Desnivel_maximo_na_area As Double, Altura_de_succao As Double
Dim Altura_do_aspersor As Single, Tempo_parado_antes_percurso As Single, Com_1 As Double
Dim Com_2 As Double, Com_3 As Double, Modelo_Trecho_A As Long
Dim Modelo_Trecho_B As Long, Modelo_Trecho_C As Long, Qtde_bomba As Integer
Dim Marca_bomba As String, Modelo_bomba As String, Tamanho_bomba As String
Dim N_estagios As Integer, Diametro_bomba As Double, Pressao_bomba As Double
Dim Rendimento_bomba As Double, Rotacao_bomba As Double, Qtde_de_Motor As Double
Dim Marca_do_Motor As String, Modelo_Motor As String, Nivel_de_Protecao As String
Dim Potencia_Nominal As Double, Nro_de_Fases As Integer, Voltagem As Double
Dim Modelo_hidroturbo As String, Eixos As Integer, Rodas As Integer
Dim Pneus As String, Tubos As String, Projetista As Long
Dim Entrega_Tecnica As String, Sequencia_do_Projeto As Long, Outras_Despesas As Double
Dim Refaturamento As Boolean, Data_do_Pedido As Variant, Data_de_Entrega As Variant
Dim Ordem_Interna As Boolean, Orcamento_Vinculado As Long, frete As String
Dim Ajuste As String
Public txtSequencia_do_Orcamento As Object, Aba1 As Object, txtCEP As Object
Public txtCaixaPostal As Object, txtFone As Object, txtObservacao As Object
Public txtMemoAuxiliar As Object, lblRGIE As Object, txtCPFCNPJ_F As Object
Public lblCPFCNPJ_F As Object, txtFax As Object, txtRGIE_F As Object
Public txtCPFCNPJ As Object, txtRGIE As Object, txtEmail As Object
Public txtMunicipio As Object, txtEndereco As Object, txtUF As Object
Public txtBairro As Object, txtNumero As Object, txtComplemento As Object
Public txtVendedor As Object, grdConjuntos As Object, grdPecas As Object
Public Grdparcelamento As Object, txtForma_de_Pagamento As Object, GrdProdutos As Object
Public grdServicos As Object, lblParcelamento As Object, lblCPFCNPJ As Object
Public Veiculo As Object, txtISS As Object, Txtperdas1 As Object
Public Txtperdas2 As Object, Lblvazao As Object, Lblvazaototal As Object
Public Txtvelodesloca As Object, Txtperdas3 As Object, Txtdeslocamento As Object
Public Txtprecipitacaolic As Object, Txtvazaoporturno As Object, Txtalturamanometrica As Object
Public Txtareapordia As Object, Txttempo1 As Object, Txtareafx As Object
Public Txtfaixasirrigadas As Object, Txtturno As Object, Txtdiam1 As Object
Public Txtdiam2 As Object, Txtdiam3 As Object, Txtcoef1 As Object
Public Txtcoef2 As Object, Txtcoef3 As Object, txtHF1 As Object
Public txtHF2 As Object, txtHF3 As Object, Txtvelo1 As Object
Public Txtvelo2 As Object, Txtvelo3 As Object, Txtperdashidro As Object
Public Lblvazaototal2 As Object, Txtpressao As Object, Txtrendimento As Object
Public Txtpotencia As Object, Txtrotacaomotor As Object, Txtdemandamotor As Object
Public Txtamperagem As Object, Txtconsumo As Object, txtFrete As Object
Public txtNF As Object, lblAjuste As Object, lblOrcamento As Object
Public txtPropriedade As Object, txtProjeto As Object, lblVinculo As Object
Dim Orcamento As New GRecordSet, Conjuntos_do_Orcamento As New GRecordSet, Parcelas_Orcamento As New GRecordSet
Dim Pecas_do_Orcamento As New GRecordSet, Produtos_do_Orcamento As New GRecordSet, Servicos_do_Orcamento As New GRecordSet

Private ProdutoAux As New GRecordSet, ConjuntoAux As New GRecordSet, ServicoAux As New GRecordSet, PecaAux As New GRecordSet
Private GeralAux As New GRecordSet, MunicipioAux As New GRecordSet, PropriedadesAux As New GRecordSet, ICMSAux As New GRecordSet
Private ProdutoNCMAux As New GRecordSet, ConjuntoNCMAux As New GRecordSet, PecasNCMAux As New GRecordSet
Private ProdutoUnidadeAux As New GRecordSet, ConjuntoUnidadeAux As New GRecordSet, PecasUnidadeAux As New GRecordSet
Private ClassificaoAux As GRecordSet, TemPropriedade As Boolean, PropriedadesGeralAux As New GRecordSet
Public Tipo2 As Byte, Fechamento2 As Byte

'evento - quando uma opo for selecionada
Private Sub opcPainel1Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel1(Index).Locked Then
      opcPainel1(Val(labopcPainel1.Caption)).Value = True
   Else
      'If Val(labopcPainel1.Caption) <> opcPainel1(Index).BookMark Then 'Manual
      labopcPainel1.Caption = Str$(opcPainel1(Index).BookMark)
      LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
      opcPainel1(Index).Change
      Select Case Index
         Case 0
            MostraFormulas
            Tipo2 = 0
         Case 1
            MostraFormulas
            Tipo2 = 1
      End Select
      'End If 'Manual
   End If
End Sub


'evento - quando uma opo for selecionada
Private Sub opcPainel2Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel2(Index).Locked Then
      opcPainel2(Val(labopcPainel2.Caption)).Value = True
   Else
      'If Val(labopcPainel2.Caption) <> opcPainel2(Index).BookMark Then 'Manual
      labopcPainel2.Caption = Str$(opcPainel2(Index).BookMark)
      LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
      opcPainel2(Index).Change
      Select Case Index
         Case 0
            Fechamento2 = 0
         Case 1
            Fechamento2 = 1
      End Select
      'End If 'Manual
   End If
End Sub


Public Sub RepositionOrcamento()
   Dim Col() As Variant

   On Error Resume Next
   
   TbAuxiliar "Geral", "[Seqncia do Geral] = " & Sequencia_do_Geral, GeralAux
   If Not Vazio(Sequencia_do_Geral) Then
      TbAuxiliar "Propriedades do Geral", "[Seqncia do Geral] = " & Sequencia_do_Geral & " AND [Seqncia da Propriedade Geral] > 0", PropriedadesGeralAux
      If PropriedadesGeralAux.RecordCount > 0 Then
         TbAuxiliar "Propriedades", "[Seqncia da Propriedade] = " & Sequencia_da_Propriedade & " AND [Seqncia da Propriedade] > 0", PropriedadesAux
         If Sequencia_da_Propriedade > 0 Then
            TbAuxiliar "Municpios", "[Seqncia do Municpio] = " & PropriedadesAux![Seqncia Do Municpio], MunicipioAux
         Else
            TbAuxiliar "Municpios", "[Seqncia do Municpio] = " & GeralAux![Seqncia Do Municpio], MunicipioAux
         End If
         TemPropriedade = True
      Else
         TbAuxiliar "Municpios", "[Seqncia do Municpio] = " & GeralAux![Seqncia Do Municpio], MunicipioAux
         TemPropriedade = False
      End If
   Else
      TbAuxiliar "Municpios", "[Seqncia do Municpio] = " & Sequencia_do_Municipio, MunicipioAux
      TemPropriedade = False
   End If
   txtForma_de_Pagamento.ListFields = "Prazo"
   txtPropriedade.PesqSQLExpression = "SELECT Propriedades.[Seqncia da Propriedade], Propriedades.[Nome da Propriedade], Propriedades.CNPJ, " + _
                                      "Propriedades.[Inscrio Estadual], Propriedades.Endereo, Propriedades.[Nmero do Endereo], Propriedades.Complemento, " + _
                                      "Propriedades.[Caixa Postal], Propriedades.Bairro, Municpios.[Seqncia do Municpio], Municpios.Descrio, " + _
                                      "Municpios.UF, Propriedades.CEP FROM Propriedades, Municpios WHERE (Propriedades.[Seqncia da Propriedade] > " + CStr(0) + ") AND " + _
                                      "(Propriedades.[Seqncia do Municpio] = Municpios.[Seqncia do Municpio]) AND (Propriedades.[Seqncia da Propriedade] IN " + _
                                      "(SELECT [Seqncia da Propriedade] FROM [Propriedades do Geral] WHERE [Seqncia do Geral] = " & Sequencia_do_Geral & " AND Inativo = False))"
   txtPropriedade.PesqFieldCapture = "Propriedades.[Seqncia da Propriedade]"
   'MostraFormulas
   'O Reposition ja chama o MostraFormulas
   'Fecha Recordsets ao Incluir
   If vgSituacao = ACAO_INCLUINDO Then
      Set Produtos_do_Orcamento = Nothing
      Set Conjuntos_do_Orcamento = Nothing
      Set Pecas_do_Orcamento = Nothing
      Set Parcelas_Orcamento = Nothing
      Set Servicos_do_Orcamento = Nothing
   End If
   Tipo2 = Tipo
   Fechamento2 = Fechamento
   
   
   'Quando navega pelos registro no respeitava as condies dos grid
   ExecutaGrid 0, Col(), CONDICOES_ESPECIAIS
   ExecutaGrid 1, Col(), CONDICOES_ESPECIAIS
   ExecutaGrid 2, Col(), CONDICOES_ESPECIAIS
   ExecutaGrid 3, Col(), CONDICOES_ESPECIAIS
   ExecutaGrid 4, Col(), CONDICOES_ESPECIAIS
   
   'Alterao
   lblAlteracao.Caption = Orcamento![Usurio da Alterao] & " " & Orcamento![Data da Alterao] & " " & Orcamento![Hora da Alterao]
   CarregaFotos
   
End Sub


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da Parcela]) As Total From [Parcelas Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND " & IIf(Seq > 0, "[Nmero da Parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


'Rotina Manual
'mostra frmulas na janela
Private Sub MostraFormulas()
   Dim MascaraIE As String
   On Error Resume Next                           'previne erros...
         
   Select Case MunicipioAux!UF
      Case "MG" '999.999.999/9999
         MascaraIE = "999.999.999/9999"
      Case "ES", "GO", "MA", "MS", "PA" '999.999.99-9
         MascaraIE = "999.999.99-9"
      Case "RJ" '99.999.99-9
         MascaraIE = "99.999.99-9"
      Case "SC" '999.999.999
         MascaraIE = "999.999.999"
      Case "DF" '99.999999.999-99
         MascaraIE = "99.999999.999-99"
      Case "PR" '99.999.999-99
         MascaraIE = "99.999.999-99"
      Case "PE" '9.999.999-99
         MascaraIE = "9.999.999-99"
      Case "RN", "AM", "PB" '99.999.999-9
         MascaraIE = "99.999.999-9"
      Case "RS" '999/999999-9
         MascaraIE = "999/999999-9"
      Case "RO" '9999999999999-9
         MascaraIE = "9999999999999-9"
      Case "SP" '999.999.999.999
         MascaraIE = "999.999.999.999"
      Case "BA" '99.999.999
         MascaraIE = "99.999.999"
      Case "CE", "AL", "AP", "PI", "SE", "RR" '99999999-9
         MascaraIE = "99999999-9"
      Case "AC" '99.999.999/999-99
         MascaraIE = "99.999.999/999-99"
      Case "MT" '9999999999-9
         MascaraIE = "9999999999-9"
      Case "TO" '99.99.999999-9
         MascaraIE = "99.99.999999-9"
   End Select
   
   If Sequencia_do_Geral = 0 Then
      If Tipo = 0 Then  'CPF e RG
         txtCPFCNPJ.Mask = "999.999.999-99"      'mascara
         lblCPFCNPJ.Caption = "CPF:"             'titulos
         txtRGIE.Mask = "@x"
         lblRGIE.Caption = "RG:"
      Else                    'CNPJ e IE
         txtCPFCNPJ.Mask = "99.999.999/9999-99"  'mascara
         lblCPFCNPJ.Caption = "CNPJ:"            'titulos
         txtRGIE.Mask = MascaraIE
         lblRGIE.Caption = "IE:"
      End If
   Else
      If GeralAux!Tipo = 0 Then  'CPF e RG
         txtCPFCNPJ_F.Mask = "999.999.999-99"      'mascara
         lblCPFCNPJ_F.Caption = "CPF:"             'titulos
         txtRGIE_F.Mask = "@x"
         lblRGIE.Caption = "RG:"
      Else                    'CNPJ e IE
         txtCPFCNPJ_F.Mask = "99.999.999/9999-99"  'mascara
         lblCPFCNPJ_F.Caption = "CNPJ:"            'titulos
         txtRGIE_F.Mask = MascaraIE
         lblRGIE.Caption = "IE:"
      End If
   End If
   
   'Inicio Manual
   If PropriedadesGeralAux.RecordCount > 0 Then
      If PropriedadesAux.RecordCount > 0 Then
         txtRGIE_F.Value = PropriedadesAux![Inscrio Do Produto]
         If Err Then Err = 0: txtRGIE_F.Text = ""
         txtEndereco.Value = PropriedadesAux!Endereo
         If Err Then Err = 0: txtEndereco.Text = ""
         txtCPFCNPJ_F.Value = PropriedadesAux!CNPJ
         If Err Then Err = 0: txtCPFCNPJ_F.Text = ""
         txtCaixaPostal.Value = PropriedadesAux![Caixa Postal]
         If Err Then Err = 0: txtCaixaPostal.Text = ""
         txtCEP.Value = PropriedadesAux!CEP
         If Err Then Err = 0: txtCEP.Text = ""
         txtBairro.Value = PropriedadesAux!Bairro
         If Err Then Err = 0: txtBairro.Text = ""
         txtNumero.Value = PropriedadesAux![Nmero Do Endereo]
         If Err Then Err = 0: txtNumero.Text = ""
         txtComplemento.Value = PropriedadesAux!Complemento
         If Err Then Err = 0: txtComplemento.Text = ""
      Else
         txtRGIE_F.Value = GeralAux![RG e IE]
         If Err Then Err = 0: txtRGIE_F.Text = ""
         txtEndereco.Value = GeralAux!Endereo
         If Err Then Err = 0: txtEndereco.Text = ""
         txtEmail.Value = GeralAux!Email
         If Err Then Err = 0: txtEmail.Text = ""
         txtCPFCNPJ_F.Value = GeralAux![CPF e CNPJ]
         If Err Then Err = 0: txtCPFCNPJ_F.Text = ""
         txtCaixaPostal.Value = GeralAux![Caixa Postal]
         If Err Then Err = 0: txtCaixaPostal.Text = ""
         txtCEP.Value = GeralAux!CEP
         If Err Then Err = 0: txtCEP.Text = ""
         txtBairro.Value = GeralAux!Bairro
         If Err Then Err = 0: txtBairro.Text = ""
         txtNumero.Value = GeralAux![Nmero Do Endereo]
         If Err Then Err = 0: txtNumero.Text = ""
         txtComplemento.Value = GeralAux!Complemento
         If Err Then Err = 0: txtComplemento.Text = ""
      End If
   Else
      txtRGIE_F.Value = GeralAux![RG e IE]
      If Err Then Err = 0: txtRGIE_F.Text = ""
      txtEndereco.Value = GeralAux!Endereo
      If Err Then Err = 0: txtEndereco.Text = ""
      txtEmail.Value = GeralAux!Email
      If Err Then Err = 0: txtEmail.Text = ""
      txtCPFCNPJ_F.Value = GeralAux![CPF e CNPJ]
      If Err Then Err = 0: txtCPFCNPJ_F.Text = ""
      txtCaixaPostal.Value = GeralAux![Caixa Postal]
      If Err Then Err = 0: txtCaixaPostal.Text = ""
      txtCEP.Value = GeralAux!CEP
      If Err Then Err = 0: txtCEP.Text = ""
      txtBairro.Value = GeralAux!Bairro
      If Err Then Err = 0: txtBairro.Text = ""
      txtNumero.Value = GeralAux![Nmero Do Endereo]
      If Err Then Err = 0: txtNumero.Text = ""
      txtComplemento.Value = GeralAux!Complemento
      If Err Then Err = 0: txtComplemento.Text = ""
   End If
   txtEmail.Value = GeralAux!Email
   If Err Then Err = 0: txtEmail.Text = ""
   txtISS.Value = Round((Valor_Total_dos_Servicos * Aliquota_do_ISS / 100), 2)
   If Err Then Err = 0: txtISS.Text = ""
   txtUF.Value = MunicipioAux!UF
   If Err Then Err = 0: txtUF.Text = ""
   txtMunicipio.Value = MunicipioAux!Descrio & "   " & MunicipioAux![Seqncia Do Municpio]
   If Err Then Err = 0: txtMunicipio.Text = ""
   txtFax.Value = GeralAux!Fax
   If Err Then Err = 0: txtFax.Text = ""
   txtFone.Value = GeralAux![Fone 1]
   If Err Then Err = 0: txtFone.Text = ""
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_Total_do_Orcamento) Or (TotalParcelas > Valor_Total_do_Orcamento) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   'label Orcamento
   lblOrcamento.Caption = Me.Caption
   If Cancelado And vgTb.RecordCount > 0 Then
      lblOrcamento.ForeColor = &H8FF
   Else
      lblOrcamento.ForeColor = &H0
   End If
   txtNF.Value = "NF.: " & Format(PegaValor("Nota Fiscal", "Seqncia da Nota Fiscal", "[Seqncia do Pedido] = " & Sequencia_do_Pedido), "000000")
   txtProjeto.Value = Format(Sequencia_do_Projeto, "000000")
   'Fim Manual
   
   lblVinculo.Caption = MostraVinculo
   Lblvazao.Caption = Format(VazaoAux, "##,###,##0.00")
   Lblvazaototal.Text = Format(VazaoAux, "##,###,##0.00")
   Lblvazaototal2.Text = Format(VazaoAux, "##,###,##0.00")
   Txtprecipitacaolic.Text = Precipitacao_bruta * 85 / 100
   Txtvazaoporturno.Text = Area_irrigada * Precipitacao_bruta * 10
   Txtpotencia.Text = Format(VazaoAux * Txtalturamanometrica.Value / 2.7 / Rendimento_bomba, "##,###,##0.00")
   
   If Modelo_Trecho_A > 0 Then
      Txtdiam1.Text = Diam1
      Txtcoef1.Text = Coef1
   Else
      Txtdiam1.Text = 0
      Txtcoef1.Text = 0
   End If
   
   If Modelo_Trecho_B > 0 Then
      Txtdiam2.Text = Diam2
      Txtcoef2.Text = Coef2
   Else
      Txtdiam2.Text = 0
      Txtcoef2.Text = 0
   End If
  
   If Modelo_Trecho_C > 0 Then
      Txtcoef3.Text = Coef3
      Txtdiam3.Text = Diam3
   Else
      Txtcoef3.Text = 0
      Txtcoef3.Text = 0
   End If
   txtHF1.Text = HF1
   txtHF2.Text = HF2
   txtHF3.Text = HF3
   
   Txtvelo1.Text = Velo1
   Txtvelo2.Text = Velo2
   Txtvelo3.Text = Velo3
   
   Txttempo1.Text = TempoFx1
   Txtvelodesloca.Text = VelocidadeDesloca
   Txtareafx.Text = AreaporFx
   Txtfaixasirrigadas.Text = FaixasIrrigadas
   Txtareapordia.Text = areapordia
   Txtdeslocamento.Text = VelocidadeDesloca
   Txtperdas1.Text = HF1
   Txtperdas2.Text = HF2
   Txtperdas3.Text = HF3
   Txtperdashidro.Text = "0.9"
   Txtalturamanometrica.Text = Orcamento![Desnivel maximo na area] + Orcamento![Altura de suco] + Orcamento![Altura do aspersor] + HF1 + HF2 + HF3 + 0.9 + Orcamento![Presso de servio]
   Txtturno.Text = Area_irrigada
   Txtrotacaomotor.Text = Rotacao_bomba
   Txtdemandamotor.Text = Potencia_Nominal * Qtde_de_Motor * 0.735 * 1.05
   'Txtamperagem.Text = VazaoAux * (((Orcamento![Desnivel maximo na area] + Orcamento![Altura de suco] + Orcamento![Altura Do aspersor] + HF1 + HF2 + HF3 + 0.9 + Orcamento![Presso de servio]) - Orcamento![Pressao bomba]) + 1.5) / 2.7 / Orcamento![Rendimento bomba] * 1.05 * 0.736 * 100000 / (1.73 * Voltagem * 0.84 * 94)
   Txtamperagem.Text = ConsumoEstimado 'Orcamento![Rendimento bomba] * 1.05 * 0.736 * 100000 / (1.73 * Voltagem * 0.84 * 94)
   'txtConsumo.Text = VazaoAux * (((Orcamento![Desnivel maximo na area] + Orcamento![Altura de suco] + Orcamento![Altura Do aspersor] + HF1 + HF2 + HF3 + 0.9 + Orcamento![Presso de servio]) - Orcamento![Pressao bomba]) + 1.5) / 2.7 / Orcamento![Rendimento bomba] * 1.05 * 0.736
   Txtconsumo.Text = Orcamento![Potencia Nominal] * 0.736
   Txtpressao.Text = (Orcamento![Desnivel maximo na area] + Orcamento![Altura de suco] + Orcamento![Altura do aspersor] + HF1 + HF2 + HF3 + 0.9 + Orcamento![Presso de servio]) / Qtde_bomba
   If Err Then Err.Clear                          'se houve erro, limpa...
   'If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO  'Se nw estiver editando ou incluindo vamos pegar recuperar as informaes do banco
                                                                                                   'Seno tiver essa linha quando o usuario navegar pelos registro ele nw aplica a mascara
                                                                                                   'como deveria
End Sub


Private Function InfoProdutos(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double, Oq As String) As Variant
 Dim MP As New GRecordSet
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, ProdutoAux
      
   If ProdutoAux.RecordCount = 0 Then Exit Function
   TbAuxiliar "Classificao Fiscal", "[Seqncia da Classificao] = " & ProdutoAux![Seqncia da Classificao] & " AND [Seqncia da Classificao] > 0", ProdutoNCMAux
   TbAuxiliar "Unidades", "[Seqncia da Unidade] = " & ProdutoAux![Seqncia da Unidade] & " AND [Seqncia da Unidade] > 0", ProdutoUnidadeAux
   
   Select Case Oq
      Case "NCM"
         InfoProdutos = ProdutoNCMAux!Ncm
      Case "Sigla"
         InfoProdutos = ProdutoUnidadeAux![Sigla da Unidade]
      Case "Valor Unitrio" 'alterao em 28-11-2024 (estava amarrando antes com o campo Sequencia_do_Produto_Orcamento)
         Set MP = vgDb.OpenRecordSet("SELECT * From [Matria Prima] WHERE [Seqncia do Produto] = " & Sequencia_do_Produto & " And [Seqncia da Matria Prima] = 43602")
         If ProdutoAux![Seqncia do Grupo Produto] = 20 And MP.RecordCount > 0 Then 'peas do pivo galvanizada
            InfoProdutos = ProdutoAux![Valor de Custo] * 3.5
         Else
           InfoProdutos = ProdutoAux![Valor Total]
         End If
         Set MP = Nothing
      Case "Peso"
         InfoProdutos = ProdutoAux!Peso
      Case "Estoque"
         InfoProdutos = ProdutoAux![Quantidade Contbil]
   End Select

End Function


Private Function InfoLista(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, ProdutoAux
      
   If ProdutoAux.RecordCount = 0 Then Exit Function
   
   Select Case Oq
      Case "Valor Unitrio"
      InfoLista = ProdutoAux![Valor de Lista]
   End Select

End Function


Private Function InfoServicos(Sequencia_do_Orcamento As Long, Sequencia_do_Servico_Orcamento As Long, Sequencia_do_Servico As Integer, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_Anterior As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Servios", "[Seqncia do Servio] = " & Sequencia_do_Servico, ServicoAux
      
   If ServicoAux.RecordCount = 0 Then Exit Function
   
   Select Case Oq
      Case "Valor Unitrio"
         InfoServicos = ServicoAux![Valor Do Servio]
   End Select

End Function


Private Function InfoConjuntos(Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Conjuntos", "[Seqncia do Conjunto] = " & Sequencia_do_Conjunto, ConjuntoAux
      
   If ConjuntoAux.RecordCount = 0 Then Exit Function
   TbAuxiliar "Classificao Fiscal", "[Seqncia da Classificao] = " & ConjuntoAux![Seqncia da Classificao] & " AND [Seqncia da Classificao] > 0", ConjuntoNCMAux
   TbAuxiliar "Unidades", "[Seqncia da Unidade] = " & ConjuntoAux![Seqncia da Unidade] & " AND [Seqncia da Unidade] > 0", ConjuntoUnidadeAux
      
   Select Case Oq
      Case "NCM"
         InfoConjuntos = ConjuntoNCMAux!Ncm
      Case "Sigla"
         InfoConjuntos = ConjuntoUnidadeAux![Sigla da Unidade]
      Case "Valor Unitrio"
         InfoConjuntos = ConjuntoAux![Valor Total]
      Case "Estoque"
         InfoConjuntos = ConjuntoAux![Quantidade Contbil]
   End Select

End Function


Private Function InfoPecas(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, PecaAux
      
   If PecaAux.RecordCount = 0 Then Exit Function
   TbAuxiliar "Classificao Fiscal", "[Seqncia da Classificao] = " & PecaAux![Seqncia da Classificao] & " AND [Seqncia da Classificao] > 0", PecasNCMAux
   TbAuxiliar "Unidades", "[Seqncia da Unidade] = " & PecaAux![Seqncia da Unidade] & " AND [Seqncia da Unidade] > 0", PecasUnidadeAux
     
   Select Case Oq
      Case "NCM"
         InfoPecas = PecasNCMAux!Ncm
      Case "Sigla"
         InfoPecas = PecasUnidadeAux![Sigla da Unidade]
      Case "Valor Unitrio"
         InfoPecas = PecaAux![Valor Total]
      Case "Peso"
         InfoPecas = PecaAux!Peso
      Case "Estoque"
         InfoPecas = PecaAux![Quantidade Contbil]
   End Select

End Function



Private Function ProcessaProdutos(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean
   Dim vrAdicional As Double, Tributos As Double, ICMSAuxiliar As Double
   Dim PisRed As Double, CofinsRed As Double
      
   On Error GoTo DeuErro
   
   If Sequencia_do_Produto_Orcamento = 0 Then
      Sequencia_do_Produto_Orcamento = SuperPegaSequencial("Produtos do Oramento", "Seqncia do Produto Oramento") - 1
   End If
   
   PisRed = 0
   CofinsRed = 0
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, ProdutoAux
   TbAuxiliar "Classificao Fiscal", "[Seqncia da Classificao] = " & ProdutoAux![Seqncia da Classificao] & " AND [Seqncia da Classificao] > 0", ProdutoNCMAux
        
   vgDb.BeginTrans
   vgDb.Execute "Update [Produtos do Oramento] Set [Valor Total] = " & Substitui(CCur(Round(Quantidade * Valor_Unitario, 2)), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
   
   If Not Orcamento![Oramento Avulso] Then 'Se no for Oramento Avulso fazer o Calculo Automatico
      vgDb.Execute "Update [Produtos do Oramento] Set CST = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 5, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CST
      If Entrega_Futura Then
         If MunicipioAux!UF = "SP" Then
            vgDb.Execute "Update [Produtos do Oramento] Set CFOP = 5922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CFOP
            vgDb.Execute "Update [Produtos do Oramento] Set CST = 90 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CFOP
            If Sequencia_da_Propriedade = 0 Then
               ICMSAuxiliar = CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            End If ' No  produtor Rural
         Else
            vgDb.Execute "Update [Produtos do Oramento] Set CFOP = 6922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CFOP
            vgDb.Execute "Update [Produtos do Oramento] Set CST = 90 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CFOP
            ICMSAuxiliar = CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
         End If
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor da Base de Clculo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Alquota do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
      Else ' Nw  Entrega Futura
         vgDb.Execute "Update [Produtos do Oramento] Set CFOP = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 1, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'CFOP
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor da Base de Clculo] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 6, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Base de Clculo
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Valor do ICMS
         
         ICMSAuxiliar = CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
         
         Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
         vgDb.Execute "Update [Produtos do Oramento] Set [Alquota do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 3, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Alquota do ICMS
         vgDb.Execute "Update [Produtos do Oramento] Set [Percentual da Reduo] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 2, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Percentual da Reduo
      End If
      vgDb.Execute "Update [Produtos do Oramento] Set [Valor do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 8, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Valor do IPI
      Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 8, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
      vgDb.Execute "Update [Produtos do Oramento] Set [Alquota do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 4, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Alquota do IPI
      vgDb.Execute "Update [Produtos do Oramento] Set Diferido = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 9, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Diferido
      
      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
         PisRed = ((Quantidade * Valor_Unitario) - ICMSAuxiliar) * 48.1 / 100
         vgDb.Execute "Update [Produtos do Oramento] Set [Bc pis] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar - PisRed), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do PIS] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar - PisRed) * 2 / 100, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'PIS
         vgDb.Execute "Update [Produtos do Oramento] Set [Aliq do PIS] = 2 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         Tributos = Tributos + ((Quantidade * Valor_Unitario) - ICMSAuxiliar - PisRed) * 2 / 100
       Else
         vgDb.Execute "Update [Produtos do Oramento] Set [Bc pis] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do PIS] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar) * 1.65 / 100, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'PIS
         vgDb.Execute "Update [Produtos do Oramento] Set [Aliq do PIS] = 1.65 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         Tributos = Tributos + ((Quantidade * Valor_Unitario) - ICMSAuxiliar) * 1.65 / 100
      End If
      
      'CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 10, 1, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
      
      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
         CofinsRed = ((Quantidade * Valor_Unitario) - ICMSAuxiliar) * 48.1 / 100
         vgDb.Execute "Update [Produtos do Oramento] Set [Bc Cofins] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar - CofinsRed), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do Cofins] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 1, (Quantidade * Valor_Unitario) - ICMSAuxiliar - CofinsRed, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Aliq do Cofins] = 9.6 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 1, (Quantidade * Valor_Unitario) - ICMSAuxiliar - CofinsRed, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
      Else
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor do Cofins] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 1, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'COFINS
         vgDb.Execute "Update [Produtos do Oramento] Set [Bc Cofins] = " & Substitui(((Quantidade * Valor_Unitario) - ICMSAuxiliar), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
         vgDb.Execute "Update [Produtos do Oramento] Set [Aliq do Cofins] = 7.6 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento
          Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 1, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
      End If
     
      vgDb.Execute "Update [Produtos do Oramento] Set IVA = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 12, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'IVA
      vgDb.Execute "Update [Produtos do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 13, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Base de Clculo ST
      vgDb.Execute "Update [Produtos do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 14, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Valor do ICMS ST
      vgDb.Execute "Update [Produtos do Oramento] Set [Alquota do ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 15, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Alquota do ICMS ST
      Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 14, 1, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
      vgDb.Execute "Update [Produtos do Oramento] Set [Valor do Tributo] = " & Substitui(CStr(Tributos), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Sequencia_do_Produto_Orcamento 'Tributos
   End If 'Caso Contrario Vamos deixar o Usuario Digitar os Valores
   vgDb.CommitTrans
   
    AtualizaValoresProdutos
    SendK (vbKeyF2)
    
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function



Private Function ProcessaConjuntos(Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single) As Boolean
   Dim vrAdicional As Double, Tributos As Double, ICMSAuxiliar As Double
   On Error GoTo DeuErro
   
   If Sequencia_Conjunto_Orcamento = 0 Then
      Sequencia_Conjunto_Orcamento = SuperPegaSequencial("Conjuntos do Oramento", "Seqncia Conjunto Oramento") - 1
   End If
  
   vgDb.BeginTrans
   vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor Total] = " & Substitui(CCur(Round(Quantidade * Valor_Unitario, 2)), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
   
   If Not Orcamento![Oramento Avulso] Then 'Se no for Oramento Avulso fazer o Calculo Automatico
      vgDb.Execute "Update [Conjuntos do Oramento] Set CST = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 5, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'CST
      If Entrega_Futura Then
         If MunicipioAux!UF = "SP" Then
            vgDb.Execute "Update [Conjuntos do Oramento] Set CFOP = 5922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'CFOP
         Else
            vgDb.Execute "Update [Conjuntos do Oramento] Set CFOP = 6922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'CFOP
            ICMSAuxiliar = CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 7, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         End If
         vgDb.Execute "Update [Conjuntos do Oramento] Set CST = 90 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'CFOP
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor da Base de Clculo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do IPI] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do IPI] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set IVA = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Base de Clculo ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor ICMS ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do ICMS ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento
      Else
         vgDb.Execute "Update [Conjuntos do Oramento] Set CFOP = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 1, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'CFOP
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor da Base de Clculo] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 6, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Base de Clculo
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 7, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Valor do ICMS
         
          ICMSAuxiliar = CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 7, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
          
         Tributos = Tributos + CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 7, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 8, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Valor do IPI
         Tributos = Tributos + CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 8, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 3, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Alquota do ICMS
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 4, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Alquota do IPI
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Percentual da Reduo] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 2, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Percentual da Reduo
         vgDb.Execute "Update [Conjuntos do Oramento] Set IVA = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 12, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'IVA
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 13, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Base de Clculo ST
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 14, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Valor do ICMS ST
         vgDb.Execute "Update [Conjuntos do Oramento] Set [Alquota do ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 15, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Alquota do ICMS ST
      End If
      vgDb.Execute "Update [Conjuntos do Oramento] Set Diferido = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 9, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Diferido
      
      'vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do PIS] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 10, 2, Quantidade * Valor_Unitario, VrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'PIS
      vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do PIS] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 10, 2, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'PIS
      Tributos = Tributos + CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 10, 2, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      
      'vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do Cofins] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 11, 2, Quantidade * Valor_Unitario, VrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'COFINS
      vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do Cofins] = " & Substitui(CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 11, 2, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'COFINS
      Tributos = Tributos + CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 11, 2, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      
      Tributos = Tributos + CalculaImposto(Sequencia_do_Conjunto, Orcamento![Seqncia Do Geral], 14, 2, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor do Tributo] = " & Substitui(CStr(Tributos), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Sequencia_Conjunto_Orcamento 'Tributos
   End If 'Caso Contrario Vamos deixar o Usuario Digitar os Valores
   vgDb.CommitTrans
    
    AtualizaValoresConjuntos
    SendK (vbKeyF2)
      
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Function ProcessaServicos(Sequencia_do_Orcamento As Long, Sequencia_do_Servico_Orcamento As Long, Sequencia_do_Servico As Integer, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_Anterior As Double) As Boolean
      
   On Error GoTo DeuErro
   
   If Sequencia_do_Servico_Orcamento = 0 Then
      Sequencia_do_Servico_Orcamento = SuperPegaSequencial("Servios do Oramento", "Seqncia do Servio Oramento") - 1
   End If
     
   vgDb.BeginTrans
   vgDb.Execute "Update [Servios do Oramento] Set [Valor Total] = " & Substitui(CCur(Round(Quantidade * Valor_Unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Servio Oramento] = " & Sequencia_do_Servico_Orcamento
   vgDb.CommitTrans
     
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Function ProcessaPecas(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double) As Boolean
   Dim vrAdicional As Double, Tributos As Double, ICMSAuxiliar As Double
   On Error GoTo DeuErro
   
   If Sequencia_Pecas_do_Orcamento = 0 Then
      Sequencia_Pecas_do_Orcamento = SuperPegaSequencial("Peas do Oramento", "Seqncia Peas do Oramento") - 1
   End If
 
   vgDb.BeginTrans
   vgDb.Execute "Update [Peas do Oramento] Set [Valor Total] = " & Substitui(CCur(Round(Quantidade * Valor_Unitario, 2)), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
   
   If Not Orcamento![Oramento Avulso] Then 'Se no for Oramento Avulso fazer o Calculo Automatico
      vgDb.Execute "Update [Peas do Oramento] Set CST = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 5, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'CST
      If Entrega_Futura Then
         If MunicipioAux!UF = "SP" Then
            vgDb.Execute "Update [Peas do Oramento] Set CFOP = 5922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'CFOP
         Else
            vgDb.Execute "Update [Peas do Oramento] Set CFOP = 6922 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'CFOP
            ICMSAuxiliar = CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         End If
         vgDb.Execute "Update [Peas do Oramento] Set CST = 90 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'CST
         vgDb.Execute "Update [Peas do Oramento] Set [Valor da Base de Clculo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Valor do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Valor do IPI] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do ICMS] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do IPI] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set IVA = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Base de Clculo ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Valor ICMS ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do ICMS ST] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento
      Else
         vgDb.Execute "Update [Peas do Oramento] Set CFOP = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 1, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'CFOP
         vgDb.Execute "Update [Peas do Oramento] Set [Valor da Base de Clculo] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 6, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Base de Clculo
         vgDb.Execute "Update [Peas do Oramento] Set [Valor do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Valor do ICMS
         
         ICMSAuxiliar = CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         
         Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 7, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         vgDb.Execute "Update [Peas do Oramento] Set [Valor do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 8, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Valor do IPI
         Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 8, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do ICMS] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 3, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Alquota do ICMS
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do IPI] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 4, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Alquota do IPI
         vgDb.Execute "Update [Peas do Oramento] Set [Percentual da Reduo] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 2, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Percentual da Reduo
         vgDb.Execute "Update [Peas do Oramento] Set IVA = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 12, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'IVA
         vgDb.Execute "Update [Peas do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 13, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Base de Clculo ST
         vgDb.Execute "Update [Peas do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 14, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Valor do ICMS ST
         vgDb.Execute "Update [Peas do Oramento] Set [Alquota do ICMS ST] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 15, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Alquota do ICMS ST
      End If
      vgDb.Execute "Update [Peas do Oramento] Set Diferido = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 9, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Diferido
      
      vgDb.Execute "Update [Peas do Oramento] Set [Valor do PIS] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 10, 3, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'PIS
      
      Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 10, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      
      vgDb.Execute "Update [Peas do Oramento] Set [Valor do Cofins] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 3, (Quantidade * Valor_Unitario) - ICMSAuxiliar, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'COFINS
      
      Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 11, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      Tributos = Tributos + CalculaImposto(Sequencia_do_Produto, Orcamento![Seqncia Do Geral], 14, 3, Quantidade * Valor_Unitario, vrAdicional, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
      vgDb.Execute "Update [Peas do Oramento] Set [Valor do Tributo] = " & Substitui(CStr(Tributos), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Sequencia_Pecas_do_Orcamento 'Tributos
   End If 'Caso Contrario Vamos deixar o Usuario Digitar os Valores
   vgDb.CommitTrans
   
   AtualizaValoresPecas
   SendK (vbKeyF2)
       
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Sub CarregaTotalizador()
   On Error Resume Next

   With GrdProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Vr. Total").Index) = True
      .ShowSumCol(.Columns("Peso").Index) = True
      .ShowSumCol(.Columns("Peso Total").Index) = True
      .ShowSumCol(.Columns("Alquota do IPI").Index) = False
      .ShowSumCol(.Columns("Alquota do ICMS").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   With grdServicos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Vr. Total").Index) = True
      .ShowSumCol(.Columns("Porcentagem de ISS").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   With grdConjuntos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Vr. Total").Index) = True
      .ShowSumCol(.Columns("Alquota do IPI").Index) = False
      .ShowSumCol(.Columns("Alquota do ICMS").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   With grdPecas
      .ShowSumBar = True
      .ShowSumCol(.Columns("Vr. Total").Index) = True
      .ShowSumCol(.Columns("Peso").Index) = True
      .ShowSumCol(.Columns("Peso Total").Index) = True
      .ShowSumCol(.Columns("Alquota do IPI").Index) = False
      .ShowSumCol(.Columns("Alquota do ICMS").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   With Grdparcelamento
      .ShowSumBar = True
      .ShowSumCol(.Columns("Vr. Total").Index) = True
      .ShowSumCol(.Columns("Nmero da Parcela").Index) = False
      .ShowSumCol(.Columns("Dias").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
End Sub


Private Sub ComandosProdutos(KeyAscii As Integer, Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double)
 Dim SemiPronto As New GRecordSet
 
   On Error GoTo DeuErro
   
   With GrdProdutos
      Select Case .ColumnField(.Col)
      Case "Quantidade"
      Set SemiPronto = vgDb.OpenRecordSet("SELECT [Seqncia Do Grupo Produto] Grupo, [Tipo do Produto] Tipo From Produtos WHERE [Seqncia do Produto] = " & Sequencia_do_Produto)
       If SemiPronto.RecordCount > 0 And Ordem_Interna Then
          If SemiPronto!Grupo = 18 Or SemiPronto!Tipo = 6 And Ordem_Interna = 0 Then
           If Ordem_Interna = 0 Then
             If MsgBox("ATENO! tem certeza que deseja incluir um Item SEMI-PRONTO no Oramento?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then
                mdiIRRIG.CancelaAlteracoes
             End If
           End If
          End If
       End If
       End Select
   End With
    
   If KeyAscii = vbKeyF12 Then
      With GrdProdutos
         Select Case .ColumnField(.Col)
            Case "Seqncia do Produto"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmProdutos.Show
         End Select
      End With
   ElseIf KeyAscii = vbKeyF2 Then
      SuperAtualizaProdutos
   End If
   
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub


Private Sub ComandosServicos(KeyAscii As Integer, Sequencia_do_Orcamento As Long, Sequencia_do_Servico_Orcamento As Long, Sequencia_do_Servico As Integer, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_Anterior As Double)
   On Error GoTo DeuErro
      
   If KeyAscii = vbKeyF12 Then
      With grdServicos
         Select Case .ColumnField(.Col)
            Case "Seqncia do Servio"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmServicos.Show
         End Select
      End With
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub


Private Sub ComandosConjuntos(KeyAscii As Integer, Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single)
   On Error GoTo DeuErro
      
   If KeyAscii = vbKeyF12 Then
      With grdConjuntos
         Select Case .ColumnField(.Col)
            Case "Seqncia do Conjunto"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmConjunto.Show
         End Select
      End With
   ElseIf KeyAscii = vbKeyF2 Then
      SuperAtualizaConjuntos
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub


Private Sub ComandosPecas(KeyAscii As Integer, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double)
 Dim SemiPronto As New GRecordSet

   On Error GoTo DeuErro
   
   Set SemiPronto = vgDb.OpenRecordSet("SELECT [Seqncia Do Grupo Produto] Grupo, [Tipo do Produto] Tipo From Produtos WHERE [Seqncia do Produto] = " & Sequencia_do_Produto)
    
   With grdPecas
   Select Case .ColumnField(.Col)
   Case "Quantidade"
   If SemiPronto.RecordCount > 0 Then
       If SemiPronto!Grupo = 18 Or SemiPronto!Tipo = 6 Then
          If MsgBox("ATENO! tem certeza que deseja incluir um Item SEMI-PRONTO no Oramento?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then
             mdiIRRIG.CancelaAlteracoes
          End If
       End If
   End If
   End Select
   End With
   
   If KeyAscii = vbKeyF12 Then
      With grdPecas
         Select Case .ColumnField(.Col)
            Case "Seqncia do Produto"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmProdutos.Show
         End Select
      End With
   ElseIf KeyAscii = vbKeyF2 Then
      SuperAtualizaPecas
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub


Private Sub AjustaValores()
   Dim IPIProdutos As New GRecordSet, IpiConjuntos As New GRecordSet, IpiPecas As New GRecordSet
   Dim ICMSProdutos As New GRecordSet, ICMSConjuntos As New GRecordSet, ICMSPecas As New GRecordSet
   Dim ICMSSTProdutos As New GRecordSet, ICMSSTConjuntos As New GRecordSet, ICMSSTPecas As New GRecordSet
   Dim BaseProdutos As New GRecordSet, BaseConjuntos As New GRecordSet, BasePecas As New GRecordSet
   Dim BaseSTProdutos As New GRecordSet, BaseSTConjuntos As New GRecordSet, BaseSTPecas As New GRecordSet
   Dim ValorProdutosUsados As New GRecordSet, ValorConjuntosUsados As New GRecordSet, ValorPecasUsadas As New GRecordSet
   Dim ValorProdutos As New GRecordSet, ValorConjuntos As New GRecordSet, ValorPecas As New GRecordSet, ValorServicos As New GRecordSet
   Dim ValorPIS As New GRecordSet, ValorCOFINS As New GRecordSet, ValorTributos As New GRecordSet
   Dim ValorOrcamento As Currency, BaseServicos As New GRecordSet, ValorISS As New GRecordSet

   On Error GoTo DeuErro
   
   'Campos Optativos
   vgDb.Execute "Update Oramento Set Tipo = " & Tipo2 & ", Fechamento = " & Fechamento2 & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
   
  
   If vgSituacao = ACAO_EXCLUINDO Then Exit Sub
   
  ' If (GeralAux!Revenda Or Orcamento!Revenda) AND Not Orcamento![Oramento Avulso] Then AjustaSubstituicao
   If Not Orcamento![Oramento Avulso] Then
      AtualizaValoresProdutos 'Atualiza Valores Conforme os valores do financeiro
      AtualizaValoresPecas 'Atualiza Valores Conforme os valores do financeiro
      AtualizaValoresConjuntos  'Atualiza Valores Conforme os valores do financeiro
   End If
   
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI dos Produtos
   Set IpiConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI dos Conjuntos
   Set IpiPecas = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI das Peas
   Set ICMSProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Produtos
   Set ICMSConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Conjuntos
   Set ICMSPecas = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Peas
   Set ICMSSTProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Produtos
   Set ICMSSTConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Conjuntos
   Set ICMSSTPecas = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Peas
   Set BaseProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Produtos
   Set BaseConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Conjuntos
   Set BasePecas = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Peas
   Set BaseSTProdutos = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Produtos
   Set BaseSTConjuntos = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Conjuntos
   Set BaseSTPecas = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Peas
   Set ValorProdutosUsados = vgDb.OpenRecordSet("SELECT SUM([Produtos do Oramento].[Valor Total]) Total " & _
                                                "FROM [Produtos do Oramento] INNER JOIN Produtos ON [Produtos do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                                "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Produtos Usados
   Set ValorConjuntosUsados = vgDb.OpenRecordSet("SELECT SUM([Conjuntos do Oramento].[Valor Total]) Total " & _
                                                 "FROM [Conjuntos do Oramento] INNER JOIN Conjuntos ON [Conjuntos do Oramento].[Seqncia Do Conjunto] = Conjuntos.[Seqncia Do Conjunto] " & _
                                                 "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Conjuntos Usados
   Set ValorPecasUsadas = vgDb.OpenRecordSet("SELECT SUM([Peas do Oramento].[Valor Total]) Total " & _
                                             "FROM [Peas do Oramento] INNER JOIN Produtos ON [Peas do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                             "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Peas Usadas
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Produtos do Oramento].[Valor Total]) Total " & _
                                          "FROM [Produtos do Oramento] INNER JOIN Produtos ON [Produtos do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                          "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Produtos Novos
   Set ValorConjuntos = vgDb.OpenRecordSet("SELECT SUM([Conjuntos do Oramento].[Valor Total]) Total " & _
                                           "FROM [Conjuntos do Oramento] INNER JOIN Conjuntos ON [Conjuntos do Oramento].[Seqncia Do Conjunto] = Conjuntos.[Seqncia Do Conjunto] " & _
                                           "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Conjuntos Novos
   Set ValorPecas = vgDb.OpenRecordSet("SELECT SUM([Peas do Oramento].[Valor Total]) Total " & _
                                       "FROM [Peas do Oramento] INNER JOIN Produtos ON [Peas do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                       "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Peas Novas
   Set ValorServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Servios
   Set ValorPIS = vgDb.OpenRecordSet("SELECT SUM([Valor do PIS]) PIS " & _
                                     "FROM(" & _
                                     "SELECT [Produtos do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                     " UNION ALL " & _
                                     "SELECT [Conjuntos do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                     " UNION ALL " & _
                                     "SELECT [Peas do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'PIS
   Set ValorCOFINS = vgDb.OpenRecordSet("SELECT SUM([Valor do Cofins]) COFINS " & _
                                        "FROM(" & _
                                        "SELECT [Produtos do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                        " UNION ALL " & _
                                        "SELECT [Conjuntos do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                        " UNION ALL " & _
                                        "SELECT [Peas do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'COFINS
   Set ValorTributos = vgDb.OpenRecordSet("SELECT SUM([Valor do Tributo]) Tributos " & _
                                          "FROM(" & _
                                          "SELECT [Produtos do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                          " UNION ALL " & _
                                          "SELECT [Conjuntos do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                          " UNION ALL " & _
                                          "SELECT [Peas do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'TRIBUTOS
   Set BaseServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ISS
   Set ValorISS = vgDb.OpenRecordSet("SELECT (SUM([Valor Total]) * [Alquota do ISS] / 100) Total " & _
                                     "FROM Oramento O LEFT JOIN [Servios Do Oramento] SO ON O.[Seqncia do Oramento] = SO.[Seqncia do Oramento] " & _
                                     "WHERE O.[Seqncia Do Oramento] = " & Sequencia_do_Orcamento & _
                                     "GROUP BY [Alquota do ISS]") 'Valor ISS
                                                                                                
   ValorOrcamento = IPIProdutos!Total + IpiConjuntos!Total + IpiPecas!Total + ValorProdutosUsados!Total + ValorConjuntosUsados!Total + ValorPecasUsadas!Total + ValorProdutos!Total + ValorConjuntos!Total + ValorPecas!Total + ValorServicos!Total + Valor_do_Seguro + Valor_do_Frete + Outras_Despesas
   ValorOrcamento = ValorOrcamento + ICMSSTProdutos!Total + ICMSSTConjuntos!Total + ICMSSTPecas!Total
   ValorOrcamento = Format(ValorOrcamento + IIf(Fechamento = 0, CCur(ValorOrcamento) * CCur(Valor_do_Fechamento) / 100, CCur(Valor_do_Fechamento)), "##,###,##0.00")
   If Orcamento![Reter ISS] And ValorServicos.RecordCount > 0 Then ValorOrcamento = ValorOrcamento * (Orcamento![Alquota Do ISS] / 100 + 1) 'Reter ISS
   If ValorServicos.RecordCount > 0 Then ValorOrcamento = ValorOrcamento - Orcamento![Valor Do Imposto de Renda] 'Imposto de Renda Sempre vai Subtrair

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update Oramento Set [Valor Total IPI dos Produtos] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Produtos
   vgDb.Execute "Update Oramento Set [Valor Total IPI dos Conjuntos] = " & Substitui(IpiConjuntos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Conjuntos
   vgDb.Execute "Update Oramento Set [Valor Total IPI das Peas] = " & Substitui(IpiPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Peas
   vgDb.Execute "Update Oramento Set [Valor Total do ICMS] = " & Substitui(ICMSProdutos!Total + ICMSConjuntos!Total + ICMSPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor do ICMS
   vgDb.Execute "Update Oramento Set [Valor Total do ICMS ST] = " & Substitui(ICMSSTProdutos!Total + ICMSSTConjuntos!Total + ICMSSTPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor do ICMS ST
   vgDb.Execute "Update Oramento Set [Valor Total da Base de Clculo] = " & Substitui(BaseProdutos!Total + BaseConjuntos!Total + BasePecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Base de Clculo ICMS
   vgDb.Execute "Update Oramento Set [Valor Total da Base ST] = " & Substitui(BaseSTProdutos!Total + BaseSTConjuntos!Total + BaseSTPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Base de Clculo ICMS ST
   vgDb.Execute "Update Oramento Set [Valor Total de Produtos Usados] = " & Substitui(ValorProdutosUsados!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Produtos Usados
   vgDb.Execute "Update Oramento Set [Valor Total Conjuntos Usados] = " & Substitui(ValorConjuntosUsados!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Conjuntos Usados
   vgDb.Execute "Update Oramento Set [Valor Total das Peas Usadas] = " & Substitui(ValorPecasUsadas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Peas Usadas
   vgDb.Execute "Update Oramento Set [Valor Total dos Produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Produtos Novos
   vgDb.Execute "Update Oramento Set [Valor Total dos Conjuntos] = " & Substitui(ValorConjuntos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Conjuntos Novos
   vgDb.Execute "Update Oramento Set [Valor Total das Peas] = " & Substitui(ValorPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Peas Novas
   vgDb.Execute "Update Oramento Set [Valor Total dos Servios] = " & Substitui(ValorServicos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Servios
   vgDb.Execute "Update Oramento Set [Valor Total do Oramento] = " & Substitui(CStr(ValorOrcamento), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor da Nota
   vgDb.Execute "Update Oramento Set [Valor Total do PIS] = " & Substitui(ValorPIS!pis, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do PIS
   vgDb.Execute "Update Oramento Set [Valor Total do COFINS] = " & Substitui(ValorCOFINS!cofins, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do COFINS
   vgDb.Execute "Update Oramento Set [Valor Total do Tributo] = " & Substitui(ValorTributos!Tributos, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do Tributo
   vgDb.CommitTrans
   
   Produtos_do_Orcamento.Requery
   Conjuntos_do_Orcamento.Requery
   Pecas_do_Orcamento.Requery
   Servicos_do_Orcamento.Requery
   
   Alteracao
   
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Sub GeraPedido()
   Dim Orcamento As New GRecordSet, OProdutos As New GRecordSet, OServicos As New GRecordSet, OConjuntos As New GRecordSet, OPecas As New GRecordSet, OParcelas As New GRecordSet
   Dim Pedido As New GRecordSet, PProdutos As New GRecordSet, PServicos As New GRecordSet, PConjuntos As New GRecordSet, PPecas As New GRecordSet, PParcelas As New GRecordSet
   Dim Geral As New GRecordSet, Propriedade As New GRecordSet, PropriedadesGeral As New GRecordSet
   Dim Filtro As String, BaixarEstoque As Boolean
   
   On Error GoTo DeuErro
   
   'Vamos perguntar para ninguem fazer merda
   If MsgBox("Deseja Gerar o Pedido?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
            
   'Baixa Estoque
   If VerificaEstoque = False Then Exit Sub
   
   BaixarEstoque = True
   vgDb.BeginTrans
   
   If TotalParcelas > 0 And TotalParcelas < Valor_Total_do_Orcamento Then
      MsgBox "Parcelamento Incompleto. Impossvel Gerar o Pedido.", vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
   
   Filtro = "[Seqncia do Oramento] = " & Sequencia_do_Orcamento
   
   Set Orcamento = vgDb.OpenRecordSet("SELECT * FROM Oramento WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set Pedido = vgDb.OpenRecordSet("SELECT * FROM Pedido")
   Set PProdutos = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Pedido]")
   Set PServicos = vgDb.OpenRecordSet("SELECT * FROM [Servios do Pedido]")
   Set PConjuntos = vgDb.OpenRecordSet("SELECT * FROM [Conjuntos do Pedido]")
   Set PPecas = vgDb.OpenRecordSet("SELECT * FROM [Peas do Pedido]")
   Set PParcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Pedido]")
   Set Geral = vgDb.OpenRecordSet("SELECT * FROM Geral")
   Set Propriedade = vgDb.OpenRecordSet("SELECT * FROM Propriedades")
   Set PropriedadesGeral = vgDb.OpenRecordSet("SELECT * FROM [Propriedades do Geral]")
   
   'Opa Cliente Novo
   If Sequencia_do_Geral = 0 Then
      With Orcamento
         Geral.AddNew
         Geral![Razo Social] = ![Nome Cliente]
         Geral![Nome Fantasia] = ![Nome Cliente]
         Geral![Fone 1] = !Telefone
         Geral!Fax = !Fax
         Geral!Tipo = !Tipo
         Geral![Cdigo Do Suframa] = ![Cdigo Do Suframa]
         Geral!Revenda = !Revenda
         Geral!Email = !Email
         Geral![Seqncia do Pas] = ![Seqncia do Pas]
         If e_Propriedade Then
            Propriedade.AddNew
            Propriedade![Nome da Propriedade] = ![Nome da Propriedade]
            Propriedade!Endereo = !Endereo
            Propriedade!Complemento = !Complemento
            Propriedade![Nmero Do Endereo] = ![Nmero Do Endereo]
            Propriedade!Bairro = !Bairro
            Propriedade!CEP = !CEP
            Propriedade![Caixa Postal] = ![Caixa Postal]
            Propriedade![Seqncia Do Municpio] = ![Seqncia Do Municpio]
            Propriedade!CNPJ = ![CPF e CNPJ]
            Propriedade![Inscrio Estadual] = ![RG e IE]
            Propriedade![ Produtor Paulista] = ![ Produtor Paulista]
            Propriedade.Update
            Propriedade.Requery
         Else
            Geral!Endereo = !Endereo
            Geral!Complemento = !Complemento
            Geral![Nmero Do Endereo] = ![Nmero Do Endereo]
            Geral!Bairro = !Bairro
            Geral!CEP = !CEP
            Geral![Caixa Postal] = ![Caixa Postal]
            Geral![Seqncia Do Municpio] = ![Seqncia Do Municpio]
            Geral![CPF e CNPJ] = ![CPF e CNPJ]
            Geral![RG e IE] = ![RG e IE]
            Geral![ Produtor Paulista] = ![ Produtor Paulista]
         End If
         Geral.Update
         Geral.Requery
         Propriedade.MoveNext: Propriedade.MoveLast
         Geral.MoveNext: Geral.MoveLast
         'Propriedades do Geral
         If e_Propriedade Then
            PropriedadesGeral.AddNew
            PropriedadesGeral![Seqncia Do Geral] = Geral![Seqncia Do Geral]
            PropriedadesGeral![Seqncia da Propriedade] = Propriedade![Seqncia da Propriedade]
            PropriedadesGeral.Update
            PropriedadesGeral.Requery
         End If
         'Oramento
         ![Seqncia Do Geral] = Geral![Seqncia Do Geral]
         ![Seqncia da Propriedade] = Propriedade![Seqncia da Propriedade]
         ![Nome Cliente] = ""
         !Telefone = ""
         !Fax = ""
         ![Cdigo Do Suframa] = ""
         !Revenda = False
         !Email = ""
         ![Nome da Propriedade] = ""
         !Endereo = ""
         !Complemento = ""
         ![Nmero Do Endereo] = ""
         !Bairro = ""
         !CEP = ""
         ![Caixa Postal] = ""
         ![Seqncia Do Municpio] = 0
         ![CPF e CNPJ] = ""
         ![RG e IE] = ""
         ![ Produtor Paulista] = False
         Orcamento.Update
         Orcamento.Requery
      End With
   End If
     
   'Processando...
   With Orcamento
      ![Seqncia Do Pedido] = SuperPegaSequencial("Pedido", "Seqncia do Pedido")
      ![Data Do Fechamento] = Date
      'Pedido
      Pedido.AddNew
      Pedido![Seqncia do Oramento] = ![Seqncia do Oramento]
      Pedido![Ocultar Valor Unitrio] = ![Ocultar Valor Unitrio]
      Pedido![Data de Emisso] = Date
      Pedido![Seqncia Do Geral] = ![Seqncia Do Geral]
      Pedido![Seqncia da Propriedade] = ![Seqncia da Propriedade]
      Pedido![Seqncia Do Vendedor] = ![Seqncia Do Vendedor]
      Pedido![Seqncia da Transportadora] = ![Seqncia da Transportadora]
      Pedido!Histrico = !Observao
      Pedido![Seqncia da Classificao] = ![Seqncia da Classificao]
      Pedido![Forma de Pagamento] = ![Forma de Pagamento]
      Pedido!Fechamento = !Fechamento
      Pedido![Valor Do Fechamento] = ![Valor Do Fechamento]
      Pedido![Valor Do Frete] = ![Valor Do Frete]
      Pedido![Valor Do Seguro] = ![Valor Do Seguro]
      Pedido![Valor Total IPI dos Produtos] = ![Valor Total IPI dos Produtos]
      Pedido![Valor Total IPI dos Conjuntos] = ![Valor Total IPI dos Conjuntos]
      Pedido![Valor Total IPI das Peas] = ![Valor Total IPI das Peas]
      'Entrega Futura no tem ICMS
      If Not ![Entrega Futura] Then
         Pedido![Valor Total Do ICMS] = ![Valor Total Do ICMS]
         Pedido![Valor Total da Base de Clculo] = ![Valor Total da Base de Clculo]
      End If
      Pedido![Valor Total de Produtos Usados] = ![Valor Total de Produtos Usados]
      Pedido![Valor Total Conjuntos Usados] = ![Valor Total Conjuntos Usados]
      Pedido![Valor Total das Peas Usadas] = ![Valor Total das Peas Usadas]
      Pedido![Valor Total dos Produtos] = ![Valor Total dos Produtos]
      Pedido![Valor Total dos Conjuntos] = ![Valor Total dos Conjuntos]
      Pedido![Valor Total das Peas] = ![Valor Total das Peas]
      Pedido![Valor Total dos Servios] = ![Valor Total dos Servios]
      Pedido![Valor Total Do Pedido] = ![Valor Total do Oramento]
      Pedido![Valor Total Do PIS] = ![Valor Total Do PIS]
      Pedido![Valor Total Do COFINS] = ![Valor Total Do COFINS]
      Pedido![Valor Total da Base ST] = ![Valor Total da Base ST]
      Pedido![Valor Total Do ICMS ST] = ![Valor Total Do ICMS ST]
      Pedido![Alquota Do ISS] = ![Alquota Do ISS]
      Pedido![Reter ISS] = ![Reter ISS]
      Pedido![Entrega Futura] = ![Entrega Futura]
      Pedido![Valor Do Imposto de Renda] = ![Valor Do Imposto de Renda]
      Pedido.Update
      Pedido.BookMark = .LastModified
      Pedido.MoveLast
      'Produtos
      TbAuxiliar "Produtos do Oramento", "[Seqncia do Oramento] = " & ![Seqncia do Oramento], OProdutos
      If OProdutos.RecordCount > 0 Then
         Do While Not OProdutos.EOF
            PProdutos.AddNew
            PProdutos![Seqncia Do Pedido] = Pedido![Seqncia Do Pedido]
            PProdutos![Seqncia do Produto] = OProdutos![Seqncia do Produto]
            PProdutos!Quantidade = OProdutos!Quantidade
            PProdutos![Valor Unitrio] = OProdutos![Valor Unitrio]
            PProdutos![Valor Total] = OProdutos![Valor Total]
            PProdutos![Valor do IPI] = OProdutos![Valor do IPI]
            PProdutos![Alquota Do IPI] = OProdutos![Alquota Do IPI]
            PProdutos!Diferido = OProdutos!Diferido
            'Entrega Futura no tem ICMS
            If Not Pedido![Entrega Futura] Then
               PProdutos![Valor Do Icms] = OProdutos![Valor Do Icms]
               PProdutos![Valor da Base de Clculo] = OProdutos![Valor da Base de Clculo]
               PProdutos![Alquota Do ICMS] = OProdutos![Alquota Do ICMS]
               PProdutos![Percentual da Reduo] = OProdutos![Percentual da Reduo]
            End If
            PProdutos![Valor Do PIS] = OProdutos![Valor Do PIS]
            PProdutos![Valor Do Cofins] = OProdutos![Valor Do Cofins]
            PProdutos!CST = OProdutos!CST
            PProdutos!CFOP = OProdutos!CFOP
            PProdutos!IVA = OProdutos!IVA
            PProdutos![Base de Clculo ST] = OProdutos![Base de Clculo ST]
            PProdutos![Valor ICMS ST] = OProdutos![Valor ICMS ST]
            PProdutos![Alquota Do ICMS ST] = OProdutos![Alquota Do ICMS ST]
            PProdutos.Update
            PProdutos.BookMark = .LastModified
            UltimaMvto 1, OProdutos![Seqncia do Produto] 'ltimo Movimento
            OProdutos.MoveNext
         Loop
      End If
      'Servios
      TbAuxiliar "Servios do Oramento", "[Seqncia do Oramento] = " & ![Seqncia do Oramento], OServicos
      If OServicos.RecordCount > 0 Then
         Do While Not OServicos.EOF
            PServicos.AddNew
            PServicos![Seqncia Do Pedido] = Pedido![Seqncia Do Pedido]
            PServicos![Seqncia do Servio] = OServicos![Seqncia do Servio]
            PServicos!Quantidade = OServicos!Quantidade
            PServicos![Valor Unitrio] = OServicos![Valor Unitrio]
            PServicos![Valor Total] = OServicos![Valor Total]
            PServicos.Update
            PServicos.BookMark = .LastModified
            OServicos.MoveNext
         Loop
      End If
      'Conjuntos
      TbAuxiliar "Conjuntos do Oramento", "[Seqncia do Oramento] = " & ![Seqncia do Oramento], OConjuntos
      If OConjuntos.RecordCount > 0 Then
         Do While Not OConjuntos.EOF
            PConjuntos.AddNew
            PConjuntos![Seqncia Do Pedido] = Pedido![Seqncia Do Pedido]
            PConjuntos![Seqncia do Conjunto] = OConjuntos![Seqncia do Conjunto]
            PConjuntos!Quantidade = OConjuntos!Quantidade
            PConjuntos![Valor Unitrio] = OConjuntos![Valor Unitrio]
            PConjuntos![Valor Total] = OConjuntos![Valor Total]
            PConjuntos![Valor do IPI] = OConjuntos![Valor do IPI]
            PConjuntos![Alquota Do IPI] = OConjuntos![Alquota Do IPI]
            'Entrega Futura no tem ICMS
            If Not Pedido![Entrega Futura] Then
               PConjuntos![Valor Do Icms] = OConjuntos![Valor Do Icms]
               PConjuntos![Alquota Do ICMS] = OConjuntos![Alquota Do ICMS]
               PConjuntos![Percentual da Reduo] = OConjuntos![Percentual da Reduo]
               PConjuntos![Valor da Base de Clculo] = OConjuntos![Valor da Base de Clculo]
            End If
            PConjuntos!Diferido = OConjuntos!Diferido
            PConjuntos![Valor Do PIS] = OConjuntos![Valor Do PIS]
            PConjuntos![Valor Do Cofins] = OConjuntos![Valor Do Cofins]
            PConjuntos!CST = OConjuntos!CST
            PConjuntos!CFOP = OConjuntos!CFOP
            PConjuntos!IVA = OConjuntos!IVA
            PConjuntos![Base de Clculo ST] = OConjuntos![Base de Clculo ST]
            PConjuntos![Valor ICMS ST] = OConjuntos![Valor ICMS ST]
            PConjuntos![Alquota Do ICMS ST] = OConjuntos![Alquota Do ICMS ST]
            PConjuntos.Update
            PConjuntos.BookMark = .LastModified
            UltimaMvto 2, OConjuntos![Seqncia do Conjunto] 'ltimo Movimento
            OConjuntos.MoveNext
         Loop
      End If
      'Peas
      TbAuxiliar "Peas do Oramento", "[Seqncia do Oramento] = " & ![Seqncia do Oramento], OPecas
      If OPecas.RecordCount > 0 Then
         Do While Not OPecas.EOF
            PPecas.AddNew
            PPecas![Seqncia Do Pedido] = Pedido![Seqncia Do Pedido]
            PPecas![Seqncia do Produto] = OPecas![Seqncia do Produto]
            PPecas!Quantidade = OPecas!Quantidade
            PPecas![Valor Unitrio] = OPecas![Valor Unitrio]
            PPecas![Valor Total] = OPecas![Valor Total]
            PPecas![Valor do IPI] = OPecas![Valor do IPI]
            'Entrega Futura no tem ICMS
            If Not Pedido![Entrega Futura] Then
               PPecas![Valor Do Icms] = OPecas![Valor Do Icms]
               PPecas![Alquota Do ICMS] = OPecas![Alquota Do ICMS]
               PPecas![Percentual da Reduo] = OPecas![Percentual da Reduo]
               PPecas![Valor da Base de Clculo] = OPecas![Valor da Base de Clculo]
            End If
            PPecas![Alquota Do IPI] = OPecas![Alquota Do IPI]
            PPecas!Diferido = OPecas!Diferido
            PPecas![Valor Do PIS] = OPecas![Valor Do PIS]
            PPecas![Valor Do Cofins] = OPecas![Valor Do Cofins]
            PPecas!CST = OPecas!CST
            PPecas!CFOP = OPecas!CFOP
            PPecas!IVA = OPecas!IVA
            PPecas![Base de Clculo ST] = OPecas![Base de Clculo ST]
            PPecas![Valor ICMS ST] = OPecas![Valor ICMS ST]
            PPecas![Alquota Do ICMS ST] = OPecas![Alquota Do ICMS ST]
            PPecas.Update
            PPecas.BookMark = .LastModified
            UltimaMvto 1, OPecas![Seqncia do Produto] 'ltimo Movimento
            OPecas.MoveNext
         Loop
      End If
      'Parcelas
      TbAuxiliar "Parcelas Oramento", "[Seqncia do Oramento] = " & ![Seqncia do Oramento], OParcelas
      If OParcelas.RecordCount > 0 Then
         Do While Not OParcelas.EOF
            PParcelas.AddNew
            PParcelas![Seqncia Do Pedido] = Pedido![Seqncia Do Pedido]
            PParcelas![Nmero da Parcela] = OParcelas![Nmero da Parcela]
            PParcelas!Dias = OParcelas!Dias
            PParcelas![Data de Vencimento] = OParcelas![Data de Vencimento]
            PParcelas![Valor da Parcela] = OParcelas![Valor da Parcela]
            PParcelas.Update
            PParcelas.BookMark = .LastModified
            OParcelas.MoveNext
         Loop
      End If
      'Atualizando...
      Orcamento.Update
   End With
   
   'Atualizando Estoque
   AtualizaEstoque 1, 0 'Produtos
   AtualizaEstoque 2, 0 'Conjuntos
   AtualizaEstoque 3, 0 'Despesas
   
   GeraConta
     
   vgDb.CommitTrans
   Unload Me
   'frmPedido.Show
   'frmPedido.vgFiltroInicial = Filtro
   'InicializaFiltro frmPedido

DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Function UltimaParcela() As Long
   Dim Tb As New GRecordSet

   On Error Resume Next
   
   Set Tb = vgDb.OpenRecordSet("SELECT MAX([Nmero da Parcela]) PC FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!Pc + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Function UltimoDias() As Long
   Dim Tb As New GRecordSet
   
   On Error Resume Next
   
   Set Tb = vgDb.OpenRecordSet("SELECT MAX([Nmero da Parcela]) Pc, Dias FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " GROUP BY Dias")
   
   If Tb.RecordCount > 0 Then
      UltimoDias = Tb!Dias
   Else
      UltimoDias = 0
   End If

End Function


Private Function UltimoVencimento() As Variant
   Dim Tb As New GRecordSet, RetVal As Variant
   'Dim Feriado As New GRecordSet,Venc As Date 'Alterado
   
   On Error Resume Next
   
   'Venc = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, CInt(grdParcelamento.Columns("Data de Vencimento").Index))
   Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de Vencimento]) [Data de Vencimento] FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   'Set Feriado = vgDb.OpenRecordSet("SELECT * FROM Calendario WHERE [Dta do Feriado] = " & D(Venc))
   RetVal = Grdparcelamento.ColumnValue(Grdparcelamento.Row + 1, CInt(Grdparcelamento.Columns("Data de Vencimento").Index))
      
   If Orcamento![Forma de Pagamento] = "Prazo" Then
      If Not IsNull(Tb![Data de Vencimento]) Then
         'Bom Vamos deixar ele colocar qualquer vencimento
         'Soh devemos tomar cuidado para no ser inferior a data de emisso
         'UltimoVencimento = RetVal >= Tb![Data de Vencimento]
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Orcamento![Data de Emisso]
      End If
   Else
      UltimoVencimento = RetVal = Orcamento![Data de Emisso]
   End If
End Function


'Vamos Ajustar a Substituio se modificarmos os valores Adicionais
Private Sub AjustaSubstituicao()
   Dim Produtos As New GRecordSet, Pecas As New GRecordSet, Conjuntos As New GRecordSet
   Dim ValorAdicional As Double
   
   On Error GoTo DeuErro
   
   Set Produtos = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set Pecas = vgDb.OpenRecordSet("SELECT * FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set Conjuntos = vgDb.OpenRecordSet("SELECT * FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
            
   If Produtos.RecordCount > 0 Then
      ValorAdicional = (Valor_do_Frete / Produtos.RecordCount) + (Valor_do_Seguro / Produtos.RecordCount)
      If Fechamento = 1 Then 'Valor
         ValorAdicional = ValorAdicional + (Valor_do_Fechamento / Produtos.RecordCount)
      Else 'Percentual
         ValorAdicional = ValorAdicional + ((Valor_Total_do_Orcamento * Valor_do_Fechamento / 100) / Produtos.RecordCount)
      End If
      ValorAdicional = Round(ValorAdicional, 2)
         
      Do While Not Produtos.EOF
         vgDb.Execute "Update [Produtos do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Produtos![Seqncia do Produto], Orcamento![Seqncia Do Geral], 13, 1, Produtos!Quantidade * Produtos![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Produtos![Seqncia do Produto Oramento] 'Base de Clculo ST
         vgDb.Execute "Update [Produtos do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Produtos![Seqncia do Produto], Orcamento![Seqncia Do Geral], 14, 1, Produtos!Quantidade * Produtos![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia do Produto Oramento] = " & Produtos![Seqncia do Produto Oramento] 'Valor do ICMS ST
         Produtos.MoveNext
      Loop
   ElseIf Conjuntos.RecordCount > 0 Or Pecas.RecordCount > 0 Then
      ValorAdicional = (Valor_do_Frete / (Pecas.RecordCount + Conjuntos.RecordCount)) + (Valor_do_Seguro / (Pecas.RecordCount + Conjuntos.RecordCount))
      If Fechamento = 1 Then 'Valor
         ValorAdicional = ValorAdicional + (Valor_do_Fechamento / (Pecas.RecordCount + Conjuntos.RecordCount))
      Else 'Percentual
         ValorAdicional = ValorAdicional + ((Valor_Total_do_Orcamento * Valor_do_Fechamento / 100) / (Pecas.RecordCount + Conjuntos.RecordCount))
      End If
      ValorAdicional = Round(ValorAdicional, 2)
        
      If Conjuntos.RecordCount > 0 Then
         Do While Not Conjuntos.EOF
            vgDb.Execute "Update [Conjuntos do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Conjuntos![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 13, 1, Conjuntos!Quantidade * Conjuntos![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Conjuntos![Seqncia Conjunto Oramento] 'Base de Clculo ST
            vgDb.Execute "Update [Conjuntos do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Conjuntos![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 14, 1, Conjuntos!Quantidade * Conjuntos![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Conjunto Oramento] = " & Conjuntos![Seqncia Conjunto Oramento] 'Valor do ICMS ST
            Conjuntos.MoveNext
         Loop
      End If
      If Pecas.RecordCount > 0 Then
         Do While Not Pecas.EOF
            vgDb.Execute "Update [Peas do Oramento] Set [Base de Clculo ST] = " & Substitui(CalculaImposto(Pecas![Seqncia do Produto], Orcamento![Seqncia Do Geral], 13, 1, Pecas!Quantidade * Pecas![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Pecas![Seqncia Peas do Oramento] 'Base de Clculo ST
            vgDb.Execute "Update [Peas do Oramento] Set [Valor ICMS ST] = " & Substitui(CalculaImposto(Pecas![Seqncia do Produto], Orcamento![Seqncia Do Geral], 14, 1, Pecas!Quantidade * Pecas![Valor Unitrio], ValorAdicional, Orcamento![Seqncia da Propriedade]), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND [Seqncia Peas do Oramento] = " & Pecas![Seqncia Peas do Oramento] 'Valor do ICMS ST
            Pecas.MoveNext
         Loop
      End If
   End If
      
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
   
End Sub


Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Function PegaNCMPadrao() As Long
   Dim Ncm As New GRecordSet
   
   On Error Resume Next
   
   Set Ncm = vgDb.OpenRecordSet("SELECT [Seqncia da Classificao] FROM [Classificao Fiscal] WHERE NCM = 84248229")
   
   If Ncm.RecordCount > 0 Then
      PegaNCMPadrao = Ncm![Seqncia da Classificao]
   End If

End Function


Private Sub GeraConta()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
   Dim Pedido As New GRecordSet

   On Error GoTo DeuErro
      
   If Vazio(Orcamento![Forma de Pagamento]) Then Exit Sub
   If TotalParcelas() <> Orcamento![Valor Total do Oramento] Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manuteno Contas]")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set Pedido = vgDb.OpenRecordSet("SELECT MAX([Seqncia do Pedido]) Numero FROM Pedido")
   
   vgDb.BeginTrans
   
   'Manuteno
   Do While Not Parcelas.EOF
      With Manutencao
         .AddNew
         !Parcela = Parcelas![Nmero da Parcela]
         ![Seqncia Do Geral] = Orcamento![Seqncia Do Geral]
         ![Seqncia da Propriedade] = Orcamento![Seqncia da Propriedade]
         ![Data de Entrada] = Date
         !Histrico = Orcamento!Observao
         ![Forma de Pagamento] = Orcamento![Forma de Pagamento]
         If ![Forma de Pagamento] = "Vista" Then
            ![Valor Pago] = Parcelas![Valor da Parcela]
            ![Valor Restante] = 0
            ![Data da Baixa] = Orcamento![Data de Emisso]
         Else
            ![Valor Pago] = 0
            ![Valor Restante] = Parcelas![Valor da Parcela]
            ![Data da Baixa] = Null
         End If
         ![Data de Vencimento] = Parcelas![Data de Vencimento]
         ![Valor Total] = Orcamento![Valor Total do Oramento]
         ![Valor da Parcela] = Parcelas![Valor da Parcela]
         ![Tipo da Conta] = IIf(Servicos_do_Orcamento.RecordCount > 0, "NFe", "NFe")
         !Conta = "R"
         ![Seqncia Do Pedido] = Pedido!Numero
         .Update
         .BookMark = .LastModified
      End With
      Parcelas.MoveNext
   Loop
   
   'Baixa
   If Forma_de_Pagamento = "Vista" Then
      Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa Contas]")
      Manutencao.MoveFirst: Manutencao.MoveLast
      With Baixa
         .AddNew
         ![Seqncia da Manuteno] = Manutencao![Seqncia da Manuteno]
         ![Data da Baixa] = Manutencao![Data de Vencimento]
         ![Valor Pago] = Manutencao![Valor da Parcela]
         !Conta = "R"
         .Update
         .BookMark = .LastModified
         Baixa.MoveLast
      End With
      With Manutencao
         .Edit
         ![Seqncia da Baixa] = Baixa![Seqncia da Baixa]
         .Update
         .BookMark = .LastModified
      End With
   End If
   
   vgDb.CommitTrans

DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Function VerificaEstoque() As Boolean
   Dim OProdutos As New GRecordSet, OPecas As New GRecordSet, OConjuntos As New GRecordSet
   Dim Prod As New GRecordSet, Pec As New GRecordSet, Conj As New GRecordSet
   Dim Sequencia() As Long, i As Long, Mensagem As String, Campo As Variant

   On Error GoTo DeuErro
   
   If Orcamento![Entrega Futura] Then VerificaEstoque = True: Exit Function
   If Orcamento![Nao Movimentar Estoque] Then VerificaEstoque = True: Exit Function
   
   Set OProdutos = vgDb.OpenRecordSet("SELECT PO.[Seqncia do Produto], [Quantidade Contbil], Quantidade, [Material Adquirido de Terceiro] FROM [Produtos do Oramento] PO INNER JOIN Produtos P ON PO.[Seqncia do Produto] = P.[Seqncia do Produto] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set OPecas = vgDb.OpenRecordSet("SELECT PO.[Seqncia do Produto], [Quantidade Contbil], Quantidade FROM [Peas do Oramento] PO INNER JOIN Produtos P ON PO.[Seqncia do Produto] = P.[Seqncia do Produto] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   Set OConjuntos = vgDb.OpenRecordSet("SELECT CO.[Seqncia do Conjunto], [Quantidade Contbil], Quantidade FROM [Conjuntos do Oramento] CO INNER JOIN Conjuntos C ON CO.[Seqncia do Conjunto] = C.[Seqncia do Conjunto] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   
   i = 0 'Tamanho do Vetor
   ReDim Preserve Sequencia(0) As Long
   
   'Produtos
   Do While Not OProdutos.EOF
      If (OProdutos![Quantidade Contbil] - OProdutos!Quantidade) < 0 Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = OProdutos![Seqncia do Produto]
      OProdutos.MoveNext
   Loop
   
   'Peas
   Do While Not OPecas.EOF
      If (OPecas![Quantidade Contbil] - OPecas!Quantidade) < 0 Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = OPecas![Seqncia do Produto]
      OPecas.MoveNext
   Loop
   
   'Conjuntos
   Do While Not OConjuntos.EOF
      If (OConjuntos![Quantidade Contbil] - OConjuntos!Quantidade) < 0 Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = OConjuntos![Seqncia do Conjunto]
      OConjuntos.MoveNext
   Loop
      
   'vamos mostrar os Itens que vao estourar estoque
   If UBound(Sequencia) > 0 Then
      Mensagem = "ATENO!!! Alguns tens Faltando no Estoque" & vbCrLf
      For Each Campo In Sequencia
         If Campo > 0 Then Mensagem = Mensagem & vbCrLf & Campo
      Next
      Mensagem = Mensagem & vbCrLf & vbCrLf & "Deseja Realmente Estourar?"
      If MsgBox(Mensagem, vbExclamation + vbYesNo + vbDefaultButton2, vaTitulo) = vbYes Then
         SuperInput
         If Not Vazio(InputArmando) Then
            If InputArmando = SuperSenha Then
               VerificaEstoque = True
            Else
               MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
               VerificaEstoque = False
            End If
            InputArmando = ""
         End If
      Else
         VerificaEstoque = False
      End If
   Else
      VerificaEstoque = True
   End If
   
   
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Function


Private Function VerificaDebitos() As Boolean
   Dim Tb As New GRecordSet
   
   On Error Resume Next
     
   Set Tb = vgDb.OpenRecordSet("SELECT 1 FROM [Manuteno Contas] WHERE [Seqncia do Geral] = " & Sequencia_do_Geral & " AND [Valor Restante] > 0 AND Conta = 'R' AND [Data de Vencimento] < " & D(Date))
        
   If Tb.RecordCount > 0 Then
      VerificaDebitos = True
   End If

End Function


Private Sub Alteracao()
   On Error Resume Next
      
   With vgTb
      .Edit
      ![Data da Alterao] = Date
      ![Hora da Alterao] = Time
      ![Usurio da Alterao] = vgPWUsuario
      .Update
      .BookMark = .LastModified
   End With

End Sub


'ROTINA MANUAL
'para mudar o titulo de quem vez a alterao
Private Sub Form_Deactivate()
   lblAlteracao.Caption = ""
End Sub



Private Sub SuperAtualizaProdutos()
   Dim Tributos As Double
   On Error Resume Next
       
   Screen.MousePointer = vbHourglass
         
   'If Produtos_do_Orcamento.RecordCount > 0 Then
   
    '  Produtos_do_Orcamento.MoveFirst
    '  Do While Not Produtos_do_Orcamento.EOF
    '     With Produtos_do_Orcamento
     '       .Edit
     '       If Entrega_Futura Then
     '          If MunicipioAux!UF = "SP" Then
     '             !CFOP = "5922"
     '             !CST = 90
     '          Else
     '             !CFOP = "6922"
     '             !CST = 90
     '          End If
     '          ![Valor da Base de Clculo] = 0
     '          ![Valor Do ICMS] = 0
     '          ![Alquota Do ICMS] = 0
     '          ![Percentual da Reduo] = 0
     '       Else ' nw  Entrega Futura
     '          !CFOP = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 1, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Valor da Base de Clculo] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 6, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Valor Do ICMS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 7, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          Tributos = Tributos + ![Valor Do ICMS]
     '          ![Alquota Do ICMS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 3, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Percentual da Reduo] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 2, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '       End If
     '          If Not Entrega_Futura Then
     '             !CST = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 5, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade])
     '          End If
     '          ![Valor Do IPI] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 8, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          Tributos = Tributos + ![Valor Do IPI]
     '          ![Alquota Do IPI] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 4, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          !Diferido = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 9, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Valor Do PIS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          Tributos = Tributos + ![Valor Do PIS]
     '          ![Valor Do Cofins] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          Tributos = Tributos + ![Valor Do Cofins]
     '          !IVA = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 12, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Base de Clculo ST] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 13, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          ![Valor ICMS ST] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 14, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
     '          Tributos = Tributos + ![Valor ICMS ST]
     '          ![Alquota Do ICMS ST] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 15, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade])
     '          ![Valor Do Tributo] = Tributos 'Tributos
     '       .Update
     '       .BookMark = .LastModified
     '       .MoveNext
     '    End With
     ' Loop
      
     '  'Atualiza do Grid
      AjustaValores      'Ajusta os Totais
       GrdProdutos.ReBind
      Reposition True    'Atualiza o Formulrio
      
  ' End If
   
   Screen.MousePointer = vbDefault

End Sub


Public Sub SuperAtualizaPecas()
   Dim Tributos As Double
   On Error Resume Next
       
   Screen.MousePointer = vbHourglass
         
   If Pecas_do_Orcamento.RecordCount > 0 Then
   
      Pecas_do_Orcamento.MoveFirst
      Do While Not Pecas_do_Orcamento.EOF
         With Pecas_do_Orcamento
            .Edit
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then

                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
               !IVA = 0
               ![Base de Clculo ST] = 0
               ![Valor ICMS ST] = 0
               ![Alquota Do ICMS ST] = 0
            Else
               !CFOP = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 1, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor da Base de Clculo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 6, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor Do Icms] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 7, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Tributos = Tributos + CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 7, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Percentual da Reduo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 2, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               !IVA = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 12, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Base de Clculo ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 13, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 14, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 15, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            End If
            If Not Entrega_Futura Then
               !CST = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 5, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao])
            End If
            ![Valor IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 8, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 8, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 4, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            !Diferido = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 9, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do PIS] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 10, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 10, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do Cofins] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 11, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 11, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do Tributo] = Tributos 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
         End With
      Loop
      
      grdPecas.ReBind 'Atualiza do Grid
      AjustaValores      'Ajusta os Totais
      Reposition True    'Atualiza o Formulrio
      
   End If
   
   Screen.MousePointer = vbDefault

End Sub


Public Sub SuperAtualizaConjuntos()
   Dim Tributos As Double
   Dim ICMSAuxiliar As Double
   
   On Error Resume Next
         
   Screen.MousePointer = vbHourglass
         
   If Conjuntos_do_Orcamento.RecordCount > 0 Then
   
      Conjuntos_do_Orcamento.MoveFirst
      Do While Not Conjuntos_do_Orcamento.EOF
         With Conjuntos_do_Orcamento
            .Edit
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then
                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
               !IVA = 0
               ![Base de Clculo ST] = 0
               ![Valor ICMS ST] = 0
               ![Alquota Do ICMS ST] = 0
                ICMSAuxiliar = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 7, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Else
               !CFOP = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 1, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor da Base de Clculo] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 6, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor Do Icms] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 7, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Tributos = Tributos + CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 7, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 3, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Percentual da Reduo] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 2, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               !IVA = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 12, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Base de Clculo ST] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 13, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor ICMS ST] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 14, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS ST] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 15, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            End If
            If Not Entrega_Futura Then
               !CST = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 5, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao])
            End If
            ![Valor IPI] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 8, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 8, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 4, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            !Diferido = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 9, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do Cofins] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, (!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do PIS] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 10, 2, (!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 10, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            ![Valor Do Tributo] = Tributos 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
         End With
      Loop
      
      grdConjuntos.ReBind 'Atualiza do Grid
      AjustaValores      'Ajusta os Totais
      Reposition True    'Atualiza o Formulrio
      
   End If
   
   Screen.MousePointer = vbDefault

End Sub


Private Sub VerificaFormulario()
   On Error Resume Next
   
   lblAlteracao.Caption = ""
   
  ' If Me.Caption = "Oramento" Then
  '    FProformaAberto = False
  '    FOrdemInternaAberto = False
  '    Exit Sub
  ' End If
   
   If Me.Caption = "Fatura Proforma" Then
      FProformaAberto = False
   End If
   
   If Me.Caption = "Ordem de Produo Interna" Then
      FOrdemInternaAberto = False
     
   'Else
      'FOrdemInternaAberto = True
   End If
   

End Sub


Private Sub Posiciona()
   On Error Resume Next
   If Me.Caption = "Oramento" Then
      PosicionaRegistro frmOrcament, "Seqncia do Oramento", seqRegistro
   Else
      PosicionaRegistro FProforma, "Seqncia do Oramento", seqRegistro
   End If
   Tipo2 = Tipo
   Fechamento2 = Fechamento
   lblAlteracao.Caption = Orcamento![Usurio da Alterao] & " " & Orcamento![Data da Alterao] & " " & Orcamento![Hora da Alterao]
End Sub


Private Sub AbreRel()
   On Error Resume Next

   If Me.Caption = "Oramento" Then
      Load frmR_Orc2
      frmR_Orc2.Show
   Else 'Proforma
      Load frmR_FProfo
      frmR_FProfo.Show
   End If

End Sub


Private Function RetornaNF() As Long
   Dim SQL As New GRecordSet
   
   Set SQL = vgDb.OpenRecordSet("SELECT [Seqncia da Nota Fiscal] NF FROM [Nota Fiscal] WHERE [Seqncia do Pedido] = " & Orcamento![Seqncia Do Pedido])
   
   If SQL.RecordCount > 0 Then
      RetornaNF = SQL!NF
   End If
   
End Function



Private Sub AbreGerar()
   Dim Geral As New GRecordSet, Propriedade As New GRecordSet, PropriedadesGeral As New GRecordSet, VemOrcamento As Boolean
   Dim DtaHoje As Date, DtaParcela As Date, Parcelas As New GRecordSet
   
   On Error Resume Next
   
   DtaHoje = Date ' Data de Hoje
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   If Parcelas.RecordCount > 0 Then
       Do While Not Parcelas.EOF
       DtaParcela = Parcelas![Data de Vencimento]
       If DtaParcela < DtaHoje Then ' Verifica de o Vencto da Parcela no  menor que a data de hj
       MsgBox "Ateno!! Data de Vencimento da Parcela" & " " & Parcelas![Nmero da Parcela] & " " & " Menor que a data de Hoje. Impossvel Gerar o Pedido.", vbExclamation + vbOKOnly, vaTitulo: Exit Sub
       End If
       Parcelas.MoveNext
       Loop
       Parcelas.MoveFirst
   End If
   
   'Vamos perguntar para ninguem fazer merda
   If MsgBox("Deseja Gerar a Nota Fiscal?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   If VerificaEstoque = False Then Exit Sub 'Baixa Estoque
   If TotalParcelas > 0 And TotalParcelas <> Valor_Total_do_Orcamento Then
      MsgBox "Parcelamento Incompleto. Impossvel Gerar o Pedido.", vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
   
   If Sequencia_do_Geral = 0 Then 'Vamos Incluir o novo Cliente
                            
      vgDb.BeginTrans
           
      Set Orcamento = vgDb.OpenRecordSet("SELECT * FROM Oramento WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
      Set Geral = vgDb.OpenRecordSet("Geral")
      Set Propriedade = vgDb.OpenRecordSet("Propriedades")
      Set PropriedadesGeral = vgDb.OpenRecordSet("[Propriedades do Geral]")
      
      With Orcamento
         Geral.AddNew
         Geral![Razo Social] = ![Nome Cliente]
         Geral![Nome Fantasia] = ![Nome Cliente]
         Geral![Fone 1] = !Telefone
         Geral!Cliente = 1
         Geral!Fax = !Fax
         Geral!Tipo = !Tipo
         Geral![Cdigo Do Suframa] = ![Cdigo Do Suframa]
         Geral!Revenda = !Revenda
         Geral!Email = !Email
         Geral![Seqncia do Pas] = ![Seqncia do Pas]
         If e_Propriedade Then
            Propriedade.AddNew
            Propriedade![Nome da Propriedade] = ![Nome da Propriedade]
            Propriedade!Endereo = !Endereo
            Propriedade!Complemento = !Complemento
            Propriedade![Nmero Do Endereo] = ![Nmero Do Endereo]
            Propriedade!Bairro = !Bairro
            Propriedade!CEP = !CEP
            Propriedade![Caixa Postal] = ![Caixa Postal]
            Propriedade![Seqncia Do Municpio] = ![Seqncia Do Municpio]
            Propriedade!CNPJ = ![CPF e CNPJ]
            Propriedade![Inscrio Estadual] = ![RG e IE]
            Propriedade![ Produtor Paulista] = ![ Produtor Paulista]
            Propriedade.Update
            Propriedade.Requery
         Else
            Geral!Endereo = !Endereo
            Geral!Complemento = !Complemento
            Geral![Nmero Do Endereo] = ![Nmero Do Endereo]
            Geral!Bairro = !Bairro
            Geral!CEP = !CEP
            Geral![Caixa Postal] = ![Caixa Postal]
            Geral![Seqncia Do Municpio] = ![Seqncia Do Municpio]
            Geral![CPF e CNPJ] = ![CPF e CNPJ]
            Geral![RG e IE] = ![RG e IE]
            Geral![ Produtor Paulista] = ![ Produtor Paulista]
         End If
         Geral.Update
         Geral.Requery
         Propriedade.MoveNext: Propriedade.MoveLast
         Geral.MoveNext: Geral.MoveLast
         'Propriedades do Geral
         If e_Propriedade Then
            PropriedadesGeral.AddNew
            PropriedadesGeral![Seqncia Do Geral] = Geral![Seqncia Do Geral]
            PropriedadesGeral![Seqncia da Propriedade] = Propriedade![Seqncia da Propriedade]
            PropriedadesGeral.Update
            PropriedadesGeral.Requery
         End If
         'Oramento
         ![Seqncia Do Geral] = Geral![Seqncia Do Geral]
         ![Seqncia da Propriedade] = Propriedade![Seqncia da Propriedade]
         ![Nome Cliente] = ""
         !Telefone = ""
         !Fax = ""
         ![Cdigo Do Suframa] = ""
         !Revenda = False
         !Email = ""
         ![Nome da Propriedade] = ""
         !Endereo = ""
         !Complemento = ""
         ![Nmero Do Endereo] = ""
         !Bairro = ""
         !CEP = ""
         ![Caixa Postal] = ""
         ![Seqncia Do Municpio] = 0
         ![CPF e CNPJ] = ""
         ![RG e IE] = ""
         ![ Produtor Paulista] = False
         Orcamento.Update
         Orcamento.Requery
      End With
   End If
   
   VemOrcamento = Not vgTb![Fatura Proforma]
   
   Unload Me
   Load frmF_GeraNF
    
   If VemOrcamento Then 'Orcamento
      frmF_GeraNF.Caption = "Oramento => Nota Fiscal"
      frmF_GeraNF.lblTitulo.Caption = "Oramento => Nota Fiscal"
      frmF_GeraNF.GeraNotaFiscal = True
      frmF_GeraNF.txtFrete.Value = txtFrete.Value
      frmF_GeraNF.Executar INICIALIZACOES
      frmF_GeraNF.Executar INI_APELIDOS
      frmF_GeraNF.Show
      SendK vbKeyReturn
      SendK vbKeyReturn
      SendK vbKeyReturn
   Else 'Proforma
      frmF_GeraNF.Caption = "F. Proforma => Nota Fiscal"
      frmF_GeraNF.lblTitulo.Caption = "F. Proforma => Nota Fiscal"
      frmF_GeraNF.GeraNotaFiscal = False
      frmF_GeraNF.Executar INICIALIZACOES
      frmF_GeraNF.Executar INI_APELIDOS
      frmF_GeraNF.Show
      SendK vbKeyReturn
      SendK vbKeyReturn
      SendK vbKeyReturn
   End If
   
   frmF_GeraNF.RepositionOrcamento
   
End Sub






'Rotina para limitar o campo memo
Private Sub LimitaCampo(KeyAscii, MaxLength)
   On Error Resume Next
   
   ' Vamos ignorar alguns comandos
   If Fatura_Proforma = True Then Exit Sub
   If KeyAscii < 32 And KeyAscii <> 22 And KeyAscii <> 13 Then Exit Sub
   
   ' Fazendo uso de campo auxiliar que tem o tamanho do memo do relatorio
     txtMemoAuxiliar.Text = txtObservacao.Text
     If SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) > 5 Or (SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) = 5 And KeyAscii = 13) Then
     'If SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) > 5 Or (SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) = 5 And KeyAscii = 13 Or KeyAscii = 46 Or KeyAscii = 47 Or KeyAscii = 45) Then
      Beep
      KeyAscii = 0
   End If

End Sub


'Tem que usar essa Rotina para podermos mudar o filtro da janela de Nota Fiscal
'Seno dependo da nota fiscal no vai abrir devido a mesma no pertencer ao filtro inicial
Private Sub AbreNotaFiscal()
   Dim Filtro As String, NF As Long
   
   On Error Resume Next
   
   NF = RetornaNF
   
   If NF = 0 Then Exit Sub
   
   Filtro = "[Seqncia da Nota Fiscal] = " & NF
   
   Load frmNotaFisc
   frmNotaFisc.vgFiltroInicial = Filtro
   InicializaFiltro frmNotaFisc
   frmNotaFisc.Show
   frmNotaFisc.SetFocus

End Sub


Public Sub AbreProjeto()

   On Error Resume Next
   
   If Sequencia_do_Projeto = 0 Then Exit Sub
    
   Load frmNovoPrg
   frmNovoPrg.txtPesquisar.Value = Sequencia_do_Projeto
   frmNovoPrg.Show
   frmNovoPrg.txtPesquisar.SetFocus
   SendK (13)

End Sub


'Retorna valor inicial' para 'DATA DE VENCIMENTO
'Private Function Orc_PulaData(Descricao As String, Sequencia_do_Orcamento As Long, Numero_da_Parcela As Integer, _
   Dias As Integer, Data_de_Vencimento As Variant, Valor_da_Parcela As Double, _
   Descricao_da_Cobranca As String) As Date
' Dim Tb As New GRecordSet,Dta As Date,DiaSemana As String
' Dta = DateAdd("d", Dias, Orcamento![Data de Emisso])
' Set Tb = vgDb.OpenRecordSet("SELECT * FROM Calendario WHERE [Dta do Feriado] = " & D(Dta))
' If Tb.RecordCount = 0 Then
' Orc_PulaData = Dta
' End If
' If Tb.RecordCount > 0 Then
' DiaSemana = Tb![Dia da Semana]
' If DiaSemana = "Seg" Then
' Orc_PulaData = Dta + 1
' ElseIf DiaSemana = "Ter" Then
' Orc_PulaData = Dta + 1
' ElseIf DiaSemana = "Qua" Then
' Orc_PulaData = Dta + 1
' ElseIf DiaSemana = "Qui" Then
' Orc_PulaData = Dta + 1
' ElseIf DiaSemana = "Sex" Then
' Orc_PulaData = Dta + 3
' ElseIf DiaSemana = "Sab" Then
' Orc_PulaData = Dta + 2
' ElseIf DiaSemana = "Dom" Then
' Orc_PulaData = Dta + 1
' End If
' End If
'End Function


Private Function ValidaProduto2(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean 'produto
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Produtos Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 
 If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaProduto2 = False
      mdiIRRIG.CancelaAlteracoes: Exit Function
   Else
      ValidaProduto2 = True: Exit Function
   End If
  End If
  
End Function



Private Function ValidaProduto3(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double) As Boolean 'peas
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Produtos Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 
 If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaProduto3 = False
      mdiIRRIG.CancelaAlteracoes: Exit Function
   Else
      ValidaProduto3 = True: Exit Function
   End If
  End If
  
End Function



Private Function FiltroOrc() As String
 Dim CodVendedor As Long
 Dim Tb As New GRecordSet
 
 If vgPWGrupo <> "VENDAS" Then Exit Function
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From [Vendedores Bloqueio]")
   
   Do While Not Tb.EOF
      If Tb![Nome do Vendedor] = vgPWUsuario Then
         CodVendedor = Tb![Codigo do Vendedor]
      End If
      Tb.MoveNext
   Loop
        
 FiltroOrc = "[Seqncia Do Oramento] > 0 AND [Fatura Proforma] = 0 AND [Ordem Interna] = 0 And Cancelado = 0 And [Seqncia do Vendedor] = " & CodVendedor
  
End Function



Private Function Permissao(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean 'Produtos
 Dim Tb As New GRecordSet
 
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Matria Prima] Where [Seqncia do Produto] = " & Sequencia_do_Produto)
  
  If Tb.RecordCount = 0 Then 'ad de terceiro ou Cadastro errado
     Permissao = True
     Exit Function
  End If
  If vgPWUsuario = "YGOR" Or vgPWUsuario = "WAGNER" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "ALEXANDRE" Or vgPWUsuario = "CESAR" Then
     Permissao = True
     Exit Function
  End If
     
End Function


Private Function PermissaoPecas(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double) As Boolean
 Dim Tb As New GRecordSet
  
   Set Tb = vgDb.OpenRecordSet("SELECT * From [Matria Prima] Where [Seqncia do Produto] = " & Sequencia_do_Produto)
    
   If Tb.RecordCount = 0 Then 'ad de terceiro ou Cadastro errado
      PermissaoPecas = True
      Exit Function
   End If
   If vgPWUsuario = "YGOR" Or vgPWUsuario = "WAGNER" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "ALEXANDRE" Or vgPWUsuario = "CESAR" Then
      PermissaoPecas = True
      Exit Function
   End If
    
End Function


Private Function PermissaoConj(Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single) As Boolean
   If vgPWUsuario = "YGOR" Or vgPWUsuario = "WAGNER" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "ALEXANDRE" Or vgPWUsuario = "CESAR" Then
      PermissaoConj = True
      Exit Function
   End If
End Function


Private Function ValidaPecasx(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double) As Boolean
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, ProdutoAux
      
   ValidaPecasx = Valor_Unitario > 0 And Valor_Unitario >= ProdutoAux![Valor Total]
   
End Function


Private Function ValidaProdx(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Sequencia_do_Produto, ProdutoAux
      
   ValidaProdx = Valor_Unitario > 0 And Valor_Unitario >= ProdutoAux![Valor Total]
   
End Function


Private Function ValidaConjx(Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single) As Boolean
   
   On Error Resume Next
   
   TbAuxiliar "Conjuntos", "[Seqncia do Conjunto] = " & Sequencia_do_Conjunto, ConjuntoAux
      
   ValidaConjx = Valor_Unitario > 0 And Valor_Unitario >= ConjuntoAux![Valor Total]
   
End Function



Private Sub EncargosFinanceiros()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Pecas As New GRecordSet
 Dim Servicos As New GRecordSet
 Dim Acrescimo As Currency
 
 If Gerou_Encargos Then
     If MsgBox("ATENO! Os encargos financeiros referente a parcelamento ao Cliente ja foram incluidos nesse Oramento deseja voltar os Valores Originais?", vbYesNo + vbQuestion, vaTitulo) = vbYes Then
        ValorOriginal
        Exit Sub
     Else
        Exit Sub
     End If
 End If
 
 If MsgBox("ATENO! esse processo vai atualizar o valor dos produtos incluindo os encargos financeiros referente a parcelamento ao cliente, caso no for parcelar esse pedido clique em no?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
 
 Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 Set Conjuntos = vgDb.OpenRecordSet("SELECT * From [Conjuntos do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 Set Servicos = vgDb.OpenRecordSet("SELECT * From [Servios do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 Set Pecas = vgDb.OpenRecordSet("SELECT * From [Peas do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)

  Do While Not Produtos.EOF
     With Produtos
       .Edit
         ![Valor Anterior] = ![Valor Unitrio]
          Acrescimo = Round((![Valor Unitrio] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100), 2)
         ![Valor Unitrio] = ![Valor Unitrio] + Acrescimo
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
     Produtos.MoveNext
  Loop
  'Peas
  Do While Not Pecas.EOF
     With Pecas
       .Edit
         ![Valor Anterior] = ![Valor Unitrio]
          Acrescimo = Round((![Valor Unitrio] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100), 2)
         ![Valor Unitrio] = ![Valor Unitrio] + Acrescimo
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
     Pecas.MoveNext
  Loop
  'Conjuntos
   Do While Not Conjuntos.EOF
     With Conjuntos
       .Edit
         ![Valor Anterior] = ![Valor Unitrio]
          Acrescimo = Round((![Valor Unitrio] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100), 2)
         ![Valor Unitrio] = ![Valor Unitrio] + Acrescimo
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
      Conjuntos.MoveNext
  Loop
  'Servicos
   Do While Not Servicos.EOF
     With Servicos
       .Edit
         ![Valor Anterior] = ![Valor Unitrio]
          Acrescimo = Round((![Valor Unitrio] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100), 2)
         ![Valor Unitrio] = ![Valor Unitrio] + Acrescimo
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
     Servicos.MoveNext
  Loop
   vgDb.Execute ("UPDATE Oramento Set [Gerou Encargos] = 1 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
   AjustaValores
   Reposition True
   MsgBox ("Encargos Financeiros Acrescentados com Sucesso!")
End Sub



Private Sub AtualizaValor()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Pecas As New GRecordSet
 Dim Tb As New GRecordSet
 Dim MP As New GRecordSet
  
 If MsgBox("ATENO! esse processo vai atualizar o valor dos produtos?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
 
 Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 Set Conjuntos = vgDb.OpenRecordSet("SELECT * From [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 Set Pecas = vgDb.OpenRecordSet("SELECT * From [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)

  Do While Not Produtos.EOF
   Set Tb = vgDb.OpenRecordSet("SELECT [Valor Total], [Seqncia do Grupo Produto], [Valor de Custo] From Produtos Where [Seqncia do Produto] = " & Produtos![Seqncia do Produto])
    'Set Mp = VgDb.OpenRecordSet("SELECT * From [Matria Prima] WHERE [Seqncia da Matria Prima] = 43602")
    Set MP = vgDb.OpenRecordSet("SELECT * From [Matria Prima] WHERE [Seqncia do Produto] = " & Produtos![Seqncia do Produto] & " And [Seqncia da Matria Prima] = 43602")
    If Tb![Valor Total] > 0 Then
       With Produtos
       .Edit
         If Tb![Seqncia do Grupo Produto] = 20 And MP.RecordCount > 0 Then
            ![Valor Unitrio] = Tb![Valor de Custo] * 3.5
            ![Valor Total] = ![Valor Unitrio] * !Quantidade
         Else
            ![Valor Unitrio] = Tb![Valor Total]
            ![Valor Total] = ![Valor Unitrio] * !Quantidade
         End If
       .Update
       .BookMark = .LastModified
       End With
    End If
     Produtos.MoveNext
  Loop
  'Peas
   Do While Not Pecas.EOF
   Set Tb = vgDb.OpenRecordSet("SELECT [Valor Total] From Produtos Where [Seqncia do Produto] = " & Pecas![Seqncia do Produto])
    If Tb![Valor Total] > 0 Then
       With Pecas
       .Edit
         ![Valor Unitrio] = Tb![Valor Total]
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
       End With
    End If
     Pecas.MoveNext
  Loop
  'Conjuntos
  Do While Not Conjuntos.EOF
     Set Tb = vgDb.OpenRecordSet("SELECT [Valor Total] From Conjuntos Where [Seqncia do Conjunto] = " & Conjuntos![Seqncia do Conjunto])
     If Tb![Valor Total] > 0 Then
        With Conjuntos
         .Edit
          ![Valor Unitrio] = Tb![Valor Total]
          ![Valor Total] = ![Valor Unitrio] * !Quantidade
         .Update
         .BookMark = .LastModified
        End With
     End If
     Conjuntos.MoveNext
  Loop
  
  AjustaValores
  Reposition True
  MsgBox ("Produtos Atualizado com Sucesso!")

End Sub



Private Sub ValorOriginal()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Pecas As New GRecordSet
 Dim Servicos As New GRecordSet
 Dim Acrescimo As Currency
 
  If Gerou_Encargos Then
     If MsgBox("ATENO! Deseja Remover os Encargos Financeiros Referente a Parcelamento ao Cliente?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
  End If
  
  Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
  Set Conjuntos = vgDb.OpenRecordSet("SELECT * From [Conjuntos do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
  Set Servicos = vgDb.OpenRecordSet("SELECT * From [Servios do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
  Set Pecas = vgDb.OpenRecordSet("SELECT * From [Peas do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)

  Do While Not Produtos.EOF
   If Produtos![Valor Anterior] > 0 Then
       With Produtos
       .Edit
         ![Valor Unitrio] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
       End With
    End If
  Produtos.MoveNext
  Loop
  'Conjuntos
   Do While Not Conjuntos.EOF
    If Conjuntos![Valor Anterior] > 0 Then
      With Conjuntos
       .Edit
         ![Valor Unitrio] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
      End With
    End If
  Conjuntos.MoveNext
  Loop
  'Pecas
   Do While Not Pecas.EOF
    If Pecas![Valor Anterior] > 0 Then
      With Pecas
       .Edit
         ![Valor Unitrio] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
      End With
    End If
  Pecas.MoveNext
  Loop
  'Servicos
   Do While Not Servicos.EOF
    If Servicos![Valor Anterior] > 0 Then
      With Servicos
       .Edit
         ![Valor Unitrio] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitrio] * !Quantidade
       .Update
       .BookMark = .LastModified
      End With
    End If
    Servicos.MoveNext
  Loop
  vgDb.Execute ("UPDATE Oramento Set [Gerou Encargos] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
  AjustaValores
  Reposition True
  MsgBox ("Encargos Financeiros Removidos com Sucesso!")
   
End Sub



Private Sub Parcelar()
 Dim Parcelas As New GRecordSet
 Dim Restante As Currency
 Dim i As Integer, DiasAux
 Dim VrPcAux As Currency, PriVez As Boolean
 Dim Tb As New GRecordSet, Abateu As Currency


 Set Parcelas = vgDb.OpenRecordSet("Parcelas Oramento")
 Set Tb = vgDb.OpenRecordSet("SELECT * From [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 
 If Tb.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [Parcelas Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
 End If
 
Caixinha:
 SuperInput4
 
 If QtdParcelaAux > 1 Then
    EncargosFinanceiros
 End If
 
 i = 1
 DiasAux = 0
 PriVez = False
 
 If ValorEntradaAux = 0 Or QtdParcelaAux = 0 Then GoTo Caixinha
 If Valor_Total_do_Orcamento = 0 Then Exit Sub
 
 If ValorEntradaAux > Valor_Total_do_Orcamento Then
    MsgBox ("valor de entrada no pode ser superior ao total do Oramento!")
    Exit Sub
 End If

 Restante = Valor_Total_do_Orcamento - ValorEntradaAux
 VrPcAux = Restante / QtdParcelaAux
 Abateu = Restante
     
    For i = 1 To QtdParcelaAux + 1
       With Parcelas
        .AddNew
           ![Seqncia do Oramento] = Sequencia_do_Orcamento
           ![Nmero da Parcela] = i
           !Dias = DiasAux
            If Not PriVez Then
               ![Data de Vencimento] = Date
               ![Valor da Parcela] = ValorEntradaAux
               PriVez = True
            Else
               ![Data de Vencimento] = DateAdd("d", DiasAux, Date)
                If i = QtdParcelaAux + 1 Then
                   If Restante - Abateu = 0 Then
                      ![Valor da Parcela] = Restante
                   Else
                      ![Valor da Parcela] = Restante + Abateu
                   End If
                Else
                   ![Valor da Parcela] = VrPcAux
                    Restante = Restante - VrPcAux
                    Abateu = Abateu - VrPcAux
                End If
            End If
        .Update
        .BookMark = .LastModified
       End With
       DiasAux = DiasAux + 30
    Next
    vgDb.Execute "DELETE FROM [Parcelas Oramento] WHERE [Valor da Parcela] = 0 And [Seqncia do Oramento] = " & Sequencia_do_Orcamento
    Grdparcelamento.ReBind
    Reposition True
    ValorEntradaAux = 0
    QtdParcelaAux = 0
    
End Sub


Private Sub DistribuiDescontoTotal()
    Dim totalDesc   As Double
    Dim totalItens  As Long
    Dim descUnit    As Double
    Dim descValue   As Double
    Dim assignedSum As Double
    Dim Rs          As GRecordSet
    Dim idx         As Long
    Dim tbl         As Variant

    ' 1) pega o desconto e usa s o mdulo
    totalDesc = Abs(Orcamento![Valor Do Fechamento])
    If totalDesc = 0 Then Exit Sub

    ' 2) conta quantos itens esto preenchidos
    totalItens = ContaRegs("[Produtos do Oramento]") _
               + ContaRegs("[Conjuntos do Oramento]") _
               + ContaRegs("[Peas do Oramento]")
    If totalItens = 0 Then Exit Sub

    ' 3) calcula valor unitrio de desconto e zera acumulador
    descUnit = Round(totalDesc / totalItens, 2)
    assignedSum = 0
    idx = 0

    ' 4) para cada tabela, grava o valor positivo do desconto
    For Each tbl In Array( _
        "[Produtos do Oramento]", _
        "[Conjuntos do Oramento]", _
        "[Peas do Oramento]" _
    )
        ' USAR Sequencia_do_Orcamento (sem acento) aqui:
        Set Rs = vgDb.OpenRecordSet( _
            "SELECT * FROM " & tbl & _
            " WHERE [Seqncia do Oramento]=" & Sequencia_do_Orcamento)

        If Not Rs.EOF Then
            Rs.MoveFirst
            Do While Not Rs.EOF
                idx = idx + 1

                If idx < totalItens Then
                    descValue = descUnit
                Else
                    descValue = Round(totalDesc - assignedSum, 2)
                End If

                Rs.Edit
                  Rs![Valor Do Desconto] = descValue
                Rs.Update

                assignedSum = assignedSum + descValue
                Rs.MoveNext
            Loop
        End If
        Rs.CloseRecordset
    Next tbl

    ' 5) rebind para atualizar a tela
    Grid(0).ReBind   ' Conjuntos
    Grid(1).ReBind   ' Peas
    Grid(3).ReBind   ' Produtos
End Sub



Private Function ContaRegs(tbl As String) As Long
    Dim rsCnt As New GRecordSet
    ' AQUI TAMBM usar Sequencia_do_Orcamento (sem acento):
    Set rsCnt = vgDb.OpenRecordSet( _
        "SELECT COUNT(*) AS Cnt FROM " & tbl & _
        " WHERE [Seqncia do Oramento]=" & Sequencia_do_Orcamento)

    If Not rsCnt.EOF Then
        ContaRegs = rsCnt!Cnt
    Else
        ContaRegs = 0
    End If
    rsCnt.CloseRecordset
End Function





Private Sub LimpaProp()
  If Sequencia_da_Propriedade > 0 Then
     txtPropriedade.Text = ""
  End If
  BuscaVendedor
End Sub


Private Sub BuscaVendedor()
  TbAuxiliar "Geral", "[Seqncia do Geral] = " & Sequencia_do_Geral, GeralAux
  If GeralAux![Seqncia Do Vendedor] > 0 Then
     txtVendedor.Value = GeralAux![Seqncia Do Vendedor]
  End If
End Sub



Private Sub PreValidaImpressao()
   Dim Tb As New GRecordSet, Itens As New GRecordSet
   Dim Vector() As String
   Dim i As Long, vaPrivez As Boolean, Campo As Variant, Mensagem As String
   Dim Servicos As New GRecordSet
   
   On Error GoTo DeuErro
  
   Set Itens = vgDb.OpenRecordSet("SELECT * FROM(" & _
                                  "SELECT PO.[Seqncia do Oramento], 'Produto' Tipo, P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], PO.[Valor Unitrio] " & _
                                  "FROM [Produtos do Oramento] PO LEFT JOIN Produtos P ON PO.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "UNION ALL " & _
                                  "SELECT CO.[Seqncia do Oramento],'Conjuntos' Tipo, C.[Seqncia Do Conjunto], C.Descrio, U.[Sigla da Unidade], CO.[Valor Unitrio] " & _
                                  "FROM [Conjuntos do Oramento] CO LEFT JOIN Conjuntos C ON CO.[Seqncia Do Conjunto] = C.[Seqncia Do Conjunto] " & _
                                  "LEFT JOIN Unidades U ON C.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "LEFT JOIN Oramento O ON CO.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                  "UNION ALL " & _
                                  "SELECT PeOrc.[Seqncia do Oramento],'Peas' Tipo, Pe.[Seqncia Do Produto], Pe.Descrio, U.[Sigla da Unidade], PeOrc.[Valor Unitrio] " & _
                                  "FROM [Peas do Oramento] PeOrc LEFT JOIN Produtos Pe ON PeOrc.[Seqncia Do Produto] = Pe.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON Pe.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "LEFT JOIN Oramento O ON PeOrc.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                  ") A " & _
                                  "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
                                  
   Set Servicos = vgDb.OpenRecordSet("SELECT * FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
                                                               
   If Itens.RecordCount = 0 And Servicos.RecordCount = 0 Then Exit Sub
   
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
 
   Do While Not Itens.EOF
      If Itens![Valor Unitrio] = 0 Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM: " & Itens![Seqncia do Produto] & " - " & Itens!Descrio
      Itens.MoveNext
   Loop
      
   If UBound(Vector) > 0 Then
      Mensagem = "Alguns Itens (sem Valor ou sem Receita):" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
   
   AbreRel
    
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub



'Private Function TemHidroturbo() As Boolean
' Dim Tb As New GRecordSet
 'Dim Tb1 As New GRecordSet
 
' Set Tb = VgDb.OpenRecordSet("SELECT [Seqncia do Conjunto] Seq From [Conjuntos do Oramento] Where [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 
 ' If Tb.RecordCount = 0 Then
 '    TemHidroturbo = False
 '    Exit Function
 ' Else
 '    Set Tb1 = VgDb.OpenRecordSet("SELECT Descrio From Conjuntos WHERE [Seqncia do Conjunto] = " & Tb!Seq & "Descrio Like '% HIDROTURBO %'")
 '        If Tb1.RecordCount > 0 Then
 '           TemHidroturbo = True
 '           Exit Function
 '        End If
 ' End If

'End Function

Private Function VazaoAux() As Double
 VazaoAux = ((Area_irrigada * Precipitacao_bruta * 10) / (Horas_irrigada * Area_tot_irrigada_em))
End Function



Private Function Coef1() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_A)
 Coef1 = Tb!Coeficiente

End Function



Private Function Coef2() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_B)
 Coef2 = Tb!Coeficiente

End Function


Private Function Coef3() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_C)
 Coef3 = Tb!Coeficiente

End Function



Private Function Diam1() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_A)
 Diam1 = Tb![DN mm] - (Tb![E mm] * 2)

End Function


Private Function Diam2() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_B)
 Diam2 = Tb![DN mm] - (Tb![E mm] * 2)

End Function

  
Private Function Diam3() As Double
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Adutoras Where [Sequencia da Adutora] = " & Modelo_Trecho_C)
 Diam3 = Tb![DN mm] - (Tb![E mm] * 2)

End Function



Private Function HF1() As Double
  If Com_1 > 0 Then
     HF1 = (((VazaoAux / Coef1) ^ 1.85) / ((Diam1 / 25.4) / 0.98) ^ 4.87) * 181 * Com_1
  Else
     HF1 = 0
  End If
End Function


Private Function HF2() As Double
  If Com_2 > 0 Then
     HF2 = (((VazaoAux / Coef2) ^ 1.85) / ((Diam2 / 25.4) / 0.98) ^ 4.87) * 181 * Com_2
  Else
     HF2 = 0
  End If
End Function


Private Function HF3() As Double
  If Com_3 > 0 Then
     HF3 = (((VazaoAux / Coef3) ^ 1.85) / ((Diam3 / 25.4) / 0.98) ^ 4.87) * 181 * Com_3
  Else
     HF3 = 0
  End If
End Function



Private Function Velo1() As Double
 Velo1 = (VazaoAux / 3600) / (((Diam1 / 2000) ^ 2) * 3.1416)
End Function


Private Function Velo2() As Double
 Velo2 = (VazaoAux / 3600) / (((Diam2 / 2000) ^ 2) * 3.1416)
End Function


Private Function Velo3() As Double
 Velo3 = (VazaoAux / 3600) / (((Diam3 / 2000) ^ 2) * 3.1416)
End Function


Private Function TempoFx1() As Double
  TempoFx1 = (Orcamento![Espao entre carreadores] * Orcamento![Faixa irrigada] / 10000) * (Orcamento![Precipitao bruta] * 10 / VazaoAux)
End Function


Private Function VelocidadeDesloca() As Double
 VelocidadeDesloca = Faixa_irrigada / TempoFx1
End Function


Private Function AreaporFx() As Double
  AreaporFx = (Orcamento![Espao entre carreadores] * Orcamento![Faixa irrigada]) / 10000
End Function


Private Function FaixasIrrigadas() As Double
  FaixasIrrigadas = Orcamento![Horas irrigada] / TempoFx1
End Function


Private Function areapordia() As Double
 areapordia = FaixasIrrigadas * AreaporFx
End Function


Private Function ConsumoEstimado()
 Dim Bcalc As Double
 
 Bcalc = Orcamento!Voltagem * 1.732 * 0.8
 ConsumoEstimado = (Orcamento![Potencia Nominal] * 0.736 * 1000) / Bcalc
 ConsumoEstimado = ConsumoEstimado + (ConsumoEstimado * 20) / 100

End Function



Public Sub AjustaValoresProforma()
   Dim IPIProdutos As New GRecordSet, IpiConjuntos As New GRecordSet, IpiPecas As New GRecordSet
   Dim ICMSProdutos As New GRecordSet, ICMSConjuntos As New GRecordSet, ICMSPecas As New GRecordSet
   Dim ICMSSTProdutos As New GRecordSet, ICMSSTConjuntos As New GRecordSet, ICMSSTPecas As New GRecordSet
   Dim BaseProdutos As New GRecordSet, BaseConjuntos As New GRecordSet, BasePecas As New GRecordSet
   Dim BaseSTProdutos As New GRecordSet, BaseSTConjuntos As New GRecordSet, BaseSTPecas As New GRecordSet
   Dim ValorProdutosUsados As New GRecordSet, ValorConjuntosUsados As New GRecordSet, ValorPecasUsadas As New GRecordSet
   Dim ValorProdutos As New GRecordSet, ValorConjuntos As New GRecordSet, ValorPecas As New GRecordSet, ValorServicos As New GRecordSet
   Dim ValorPIS As New GRecordSet, ValorCOFINS As New GRecordSet, ValorTributos As New GRecordSet
   Dim ValorOrcamento As Currency, BaseServicos As New GRecordSet, ValorISS As New GRecordSet

   On Error GoTo DeuErro
   
   'Campos Optativos
   vgDb.Execute "Update Oramento Set Tipo = " & Tipo2 & ", Fechamento = " & Fechamento2 & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
   
   If vgSituacao = ACAO_EXCLUINDO Then Exit Sub
   
  ' If (GeralAux!Revenda Or Orcamento!Revenda) AND Not Orcamento![Oramento Avulso] Then AjustaSubstituicao
   If Not Orcamento![Oramento Avulso] Then
      AtualizaValoresProdutos 'Atualiza Valores Conforme os valores do financeiro
      AtualizaValoresPecas 'Atualiza Valores Conforme os valores do financeiro
      AtualizaValoresConjuntos  'Atualiza Valores Conforme os valores do financeiro
   End If
   
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI dos Produtos
   Set IpiConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI dos Conjuntos
   Set IpiPecas = vgDb.OpenRecordSet("SELECT SUM([Valor Do IPI]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'IPI das Peas
   Set ICMSProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Produtos
   Set ICMSConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Conjuntos
   Set ICMSPecas = vgDb.OpenRecordSet("SELECT SUM([Valor Do ICMS]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS Peas
   Set ICMSSTProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Produtos
   Set ICMSSTConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Conjuntos
   Set ICMSSTPecas = vgDb.OpenRecordSet("SELECT SUM([Valor ICMS ST]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'ICMS ST Peas
   Set BaseProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Produtos
   Set BaseConjuntos = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Conjuntos
   Set BasePecas = vgDb.OpenRecordSet("SELECT SUM([Valor da Base de Clculo]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base Peas
   Set BaseSTProdutos = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Produtos
   Set BaseSTConjuntos = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Conjuntos
   Set BaseSTPecas = vgDb.OpenRecordSet("SELECT SUM([Base de Clculo ST]) Total FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ST Peas
   Set ValorProdutosUsados = vgDb.OpenRecordSet("SELECT SUM([Produtos do Oramento].[Valor Total]) Total " & _
                                                "FROM [Produtos do Oramento] INNER JOIN Produtos ON [Produtos do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                                "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Produtos Usados
   Set ValorConjuntosUsados = vgDb.OpenRecordSet("SELECT SUM([Conjuntos do Oramento].[Valor Total]) Total " & _
                                                 "FROM [Conjuntos do Oramento] INNER JOIN Conjuntos ON [Conjuntos do Oramento].[Seqncia Do Conjunto] = Conjuntos.[Seqncia Do Conjunto] " & _
                                                 "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Conjuntos Usados
   Set ValorPecasUsadas = vgDb.OpenRecordSet("SELECT SUM([Peas do Oramento].[Valor Total]) Total " & _
                                             "FROM [Peas do Oramento] INNER JOIN Produtos ON [Peas do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                             "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 1") 'Peas Usadas
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Produtos do Oramento].[Valor Total]) Total " & _
                                          "FROM [Produtos do Oramento] INNER JOIN Produtos ON [Produtos do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                          "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Produtos Novos
   Set ValorConjuntos = vgDb.OpenRecordSet("SELECT SUM([Conjuntos do Oramento].[Valor Total]) Total " & _
                                           "FROM [Conjuntos do Oramento] INNER JOIN Conjuntos ON [Conjuntos do Oramento].[Seqncia Do Conjunto] = Conjuntos.[Seqncia Do Conjunto] " & _
                                           "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Conjuntos Novos
   Set ValorPecas = vgDb.OpenRecordSet("SELECT SUM([Peas do Oramento].[Valor Total]) Total " & _
                                       "FROM [Peas do Oramento] INNER JOIN Produtos ON [Peas do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] " & _
                                       "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento & " AND Usado = 0") 'Peas Novas
   Set ValorServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Servios
   Set ValorPIS = vgDb.OpenRecordSet("SELECT SUM([Valor do PIS]) PIS " & _
                                     "FROM(" & _
                                     "SELECT [Produtos do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                     " UNION ALL " & _
                                     "SELECT [Conjuntos do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                     " UNION ALL " & _
                                     "SELECT [Peas do Oramento].[Valor do PIS] " & _
                                     "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                     "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'PIS
   Set ValorCOFINS = vgDb.OpenRecordSet("SELECT SUM([Valor do Cofins]) COFINS " & _
                                        "FROM(" & _
                                        "SELECT [Produtos do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                        " UNION ALL " & _
                                        "SELECT [Conjuntos do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                        " UNION ALL " & _
                                        "SELECT [Peas do Oramento].[Valor do Cofins] " & _
                                        "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                        "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'COFINS
   Set ValorTributos = vgDb.OpenRecordSet("SELECT SUM([Valor do Tributo]) Tributos " & _
                                          "FROM(" & _
                                          "SELECT [Produtos do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Produtos do Oramento] ON Oramento.[Seqncia do Oramento] = [Produtos do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                          " UNION ALL " & _
                                          "SELECT [Conjuntos do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Conjuntos do Oramento] ON Oramento.[Seqncia do Oramento] = [Conjuntos do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & _
                                          " UNION ALL " & _
                                          "SELECT [Peas do Oramento].[Valor do Tributo] " & _
                                          "FROM Oramento INNER JOIN [Peas do Oramento] ON Oramento.[Seqncia do Oramento] = [Peas do Oramento].[Seqncia do Oramento] " & _
                                          "WHERE Oramento.[Seqncia do Oramento] = " & Sequencia_do_Orcamento & ") A") 'TRIBUTOS
   Set BaseServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento) 'Base ISS
   Set ValorISS = vgDb.OpenRecordSet("SELECT (SUM([Valor Total]) * [Alquota do ISS] / 100) Total " & _
                                     "FROM Oramento O LEFT JOIN [Servios Do Oramento] SO ON O.[Seqncia do Oramento] = SO.[Seqncia do Oramento] " & _
                                     "WHERE O.[Seqncia Do Oramento] = " & Sequencia_do_Orcamento & _
                                     "GROUP BY [Alquota do ISS]") 'Valor ISS
                                                                                                
   ValorOrcamento = IPIProdutos!Total + IpiConjuntos!Total + IpiPecas!Total + ValorProdutosUsados!Total + ValorConjuntosUsados!Total + ValorPecasUsadas!Total + ValorProdutos!Total + ValorConjuntos!Total + ValorPecas!Total + ValorServicos!Total + Valor_do_Seguro + Valor_do_Frete
   ValorOrcamento = ValorOrcamento + ICMSSTProdutos!Total + ICMSSTConjuntos!Total + ICMSSTPecas!Total
   ValorOrcamento = Format(ValorOrcamento + IIf(Fechamento = 0, CCur(ValorOrcamento) * CCur(Valor_do_Fechamento) / 100, CCur(Valor_do_Fechamento)), "##,###,##0.00")
   If Orcamento![Reter ISS] And ValorServicos.RecordCount > 0 Then ValorOrcamento = ValorOrcamento * (Orcamento![Alquota Do ISS] / 100 + 1) 'Reter ISS
   If ValorServicos.RecordCount > 0 Then ValorOrcamento = ValorOrcamento - Orcamento![Valor Do Imposto de Renda] 'Imposto de Renda Sempre vai Subtrair

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update Oramento Set [Valor Total IPI dos Produtos] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Produtos
   vgDb.Execute "Update Oramento Set [Valor Total IPI dos Conjuntos] = " & Substitui(IpiConjuntos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Conjuntos
   vgDb.Execute "Update Oramento Set [Valor Total IPI das Peas] = " & Substitui(IpiPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'IPI Peas
   vgDb.Execute "Update Oramento Set [Valor Total do ICMS] = " & Substitui(ICMSProdutos!Total + ICMSConjuntos!Total + ICMSPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor do ICMS
   vgDb.Execute "Update Oramento Set [Valor Total do ICMS ST] = " & Substitui(ICMSSTProdutos!Total + ICMSSTConjuntos!Total + ICMSSTPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor do ICMS ST
   vgDb.Execute "Update Oramento Set [Valor Total da Base de Clculo] = " & Substitui(BaseProdutos!Total + BaseConjuntos!Total + BasePecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Base de Clculo ICMS
   vgDb.Execute "Update Oramento Set [Valor Total da Base ST] = " & Substitui(BaseSTProdutos!Total + BaseSTConjuntos!Total + BaseSTPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Base de Clculo ICMS ST
   vgDb.Execute "Update Oramento Set [Valor Total de Produtos Usados] = " & Substitui(ValorProdutosUsados!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Produtos Usados
   vgDb.Execute "Update Oramento Set [Valor Total Conjuntos Usados] = " & Substitui(ValorConjuntosUsados!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Conjuntos Usados
   vgDb.Execute "Update Oramento Set [Valor Total das Peas Usadas] = " & Substitui(ValorPecasUsadas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Peas Usadas
   vgDb.Execute "Update Oramento Set [Valor Total dos Produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Produtos Novos
   vgDb.Execute "Update Oramento Set [Valor Total dos Conjuntos] = " & Substitui(ValorConjuntos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Conjuntos Novos
   vgDb.Execute "Update Oramento Set [Valor Total das Peas] = " & Substitui(ValorPecas!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Peas Novas
   vgDb.Execute "Update Oramento Set [Valor Total dos Servios] = " & Substitui(ValorServicos!Total, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Servios
   vgDb.Execute "Update Oramento Set [Valor Total do Oramento] = " & Substitui(CStr(ValorOrcamento), ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor da Nota
   vgDb.Execute "Update Oramento Set [Valor Total do PIS] = " & Substitui(ValorPIS!pis, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do PIS
   vgDb.Execute "Update Oramento Set [Valor Total do COFINS] = " & Substitui(ValorCOFINS!cofins, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do COFINS
   vgDb.Execute "Update Oramento Set [Valor Total do Tributo] = " & Substitui(ValorTributos!Tributos, ",", ".", SO_UM) & " WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento 'Valor Total do Tributo
   vgDb.CommitTrans
   
   Produtos_do_Orcamento.Requery
   Conjuntos_do_Orcamento.Requery
   Pecas_do_Orcamento.Requery
   Servicos_do_Orcamento.Requery
   
   Alteracao
   
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub



Private Function PodeVenderProd(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean 'produto
 Dim Tb As New GRecordSet
 Dim TemReceita As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Produtos Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 Set TemReceita = vgDb.OpenRecordSet("SELECT * From [Matria Prima] Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 
  If Tb.RecordCount > 0 Then
     If Tb![Tipo Do Produto] = 0 And TemReceita.RecordCount = 0 Then
        PodeVenderProd = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Tb![Tipo Do Produto] = 3 Or Tb![Tipo Do Produto] = 5 Or Tb![Tipo Do Produto] = 7 Or Tb![Tipo Do Produto] = 8 Or Tb![Tipo Do Produto] = 9 Then
        PodeVenderProd = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Tb![Material Adquirido de Terceiro] And Tb![ Matria Prima] And Tb![Tipo Do Produto] > 1 Then
         PodeVenderProd = False
         mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Not Tb![Material Adquirido de Terceiro] And Not Tb![ Matria Prima] And Tb![Tipo Do Produto] = 0 And TemReceita.RecordCount = 0 Then
         PodeVenderProd = False
         mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Not Tb![Material Adquirido de Terceiro] And Not Tb![ Matria Prima] And Tb![Tipo Do Produto] = 1 And TemReceita.RecordCount = 0 Then
         PodeVenderProd = False
         mdiIRRIG.CancelaAlteracoes: Exit Function
     Else
         PodeVenderProd = True: Exit Function
     End If
  End If
  
End Function



Private Function PodeVenderPecas(Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, Sequencia_do_Produto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, _
   Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, _
   Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single, Peso As Double) As Boolean 'produto
 Dim Tb As New GRecordSet
 Dim TemReceita As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Produtos Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 Set TemReceita = vgDb.OpenRecordSet("SELECT * From [Matria Prima] Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 
  If Tb.RecordCount > 0 Then
     If Tb![Tipo Do Produto] = 0 And TemReceita.RecordCount = 0 Then
        PodeVenderPecas = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Tb![Tipo Do Produto] = 3 Or Tb![Tipo Do Produto] = 5 Or Tb![Tipo Do Produto] = 7 Or Tb![Tipo Do Produto] = 8 Or Tb![Tipo Do Produto] = 9 Then
        PodeVenderPecas = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Tb![Material Adquirido de Terceiro] And Tb![ Matria Prima] And Tb![Tipo Do Produto] > 1 Then
        PodeVenderPecas = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Not Tb![Material Adquirido de Terceiro] And Not Tb![ Matria Prima] And Tb![Tipo Do Produto] = 0 And TemReceita.RecordCount = 0 Then
        PodeVenderPecas = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     ElseIf Not Tb![Material Adquirido de Terceiro] And Not Tb![ Matria Prima] And Tb![Tipo Do Produto] = 1 And TemReceita.RecordCount = 0 Then
        PodeVenderPecas = False
        mdiIRRIG.CancelaAlteracoes: Exit Function
     Else
        PodeVenderPecas = True: Exit Function
     End If
  End If
  
End Function



'Private Function ValidaNCM(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean
' Dim Tb As New GRecordSet
 
' Set Tb = VgDb.OpenRecordSet("SELECT * From Produtos Where [Seqncia do Produto] = " & Sequencia_do_produto)
 
 'ValidaNCM = Tb![Seqncia da Classificao] > 0
   
 
'End Function

Private Function ValidaNCM(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double) As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Seqncia da Classificao], [Conferido pelo Contabil] From Produtos Where [Seqncia do Produto] = " & Sequencia_do_Produto)
 
 If Ordem_Interna = 0 Then
    ValidaNCM = Tb![Seqncia da Classificao] > 0 And Tb![Conferido pelo Contabil]
 Else
    ValidaNCM = True
 End If
 
End Function



Private Function ValidaConjunto(Sequencia_do_Orcamento As Long, Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, _
   Quantidade As Double, Valor_Unitario As Double, Valor_Total As Double, _
   Valor_do_IPI As Double, Valor_do_ICMS As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean, _
   Valor_da_Base_de_Calculo As Double, Valor_do_Tributo As Double, Valor_do_PIS As Double, _
   Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double, _
   CFOP As Integer, CST As Integer, Valor_ICMS_ST As Double, _
   Aliquota_do_ICMS_ST As Single, Valor_do_Desconto As Double, Valor_do_Frete As Double, _
   Valor_Anterior As Double, Bc_pis As Double, Aliq_do_pis As Single, _
   Bc_cofins As Double, Aliq_do_cofins As Single) As Boolean 'produto
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From Conjuntos Where [Seqncia do Conjunto] = " & Sequencia_do_Conjunto)
 
 If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaConjunto = False
      mdiIRRIG.CancelaAlteracoes: Exit Function
   Else
      ValidaConjunto = True: Exit Function
   End If
  End If
  
End Function



Private Sub CarregaFotos()

    If Vazio(Parametros![Diretorio das Fotos]) Then
       Exit Sub
    End If
    
    If Fatura_Proforma Then
      Exit Sub
    End If
    
    If Ordem_Interna Then
      Exit Sub
    End If
    
    If Existe(Parametros![Diretorio das Fotos] & "Orc_" & Sequencia_do_Orcamento & ".jpg") Then
        Set mmCampo(0).Picture = LoadPicture(Parametros![Diretorio das Fotos] & "Orc_" & Sequencia_do_Orcamento & ".jpg")
        mmCampo(0).ToolTipText = "Mapa da Propriedade"
    Else
        Set mmCampo(0).Picture = Nothing
    End If
   
End Sub



Private Sub AtualizaValoresConjuntos()
   Dim Tributos As Double, ValorFrete As Double, ValorSeguro As Double, ValorAcresDesc As Double
   Dim totalItens As Double, PercentualRateio As Double, VrItem As Double, PercentualFrete As Double
   Dim BC As Double, ICMSAuxiliar As Double, BCAuxiliar As Double
   Dim AliqAuxiliar As Double, Geral As New GRecordSet, Suframa As Boolean
   Dim PisRed As Double, CofinsRed As Double
          
   Screen.MousePointer = vbHourglass
     
   Call DistribuiDescontoTotal
   
   grdConjuntos.ReBind
   
   Set Geral = vgDb.OpenRecordSet("SELECT * From Geral Where [Seqncia Do Geral] = " & Orcamento![Seqncia Do Geral])
   Suframa = (CBool(Not Vazio(Geral![Cdigo Do Suframa])))
   If Suframa Then
      vgDb.Execute "Update [Conjuntos do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
      grdConjuntos.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
           
   If Conjuntos_do_Orcamento.RecordCount > 0 Then
   
      
      If vgSituacao = -ACAO_INCLUINDO Or vgSituacao = -ACAO_EDITANDO Then
         If ValorAcresDesc = 0 And ValorFrete = 0 And ValorSeguro = 0 Then GoTo SaiDaSub
      End If
   
      Conjuntos_do_Orcamento.MoveFirst
      Do While Not Conjuntos_do_Orcamento.EOF
         Tributos = 0
         PisRed = 0
         CofinsRed = 0
         With Conjuntos_do_Orcamento
         .Edit
         
         If Orcamento![Valor Do Fechamento] = 0 Then
            ![Valor Do Desconto] = 0
         End If
         If Orcamento![Valor Do Frete] = 0 Then
            ![Valor Do Frete] = 0
         End If
         
         
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then
                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
               !IVA = 0
               ![Base de Clculo ST] = 0
               ![Valor ICMS ST] = 0
               ![Alquota Do ICMS ST] = 0
            Else
               '![Valor da Base de Clculo] = CalculaImposto(![Seqncia Do Conjunto], Orcamento![Seqncia Do Geral], 6, 2, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               BC = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 6, 2, (!Quantidade * ![Valor Unitrio] + ValorAcresDesc + ValorFrete), 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               'Alterado em 15/07/2024, em conformidade com Maysa e Econet. Frete compondo a BC antes da reduo.
               ![Valor da Base de Clculo] = BC '+ ValorFrete ' Alterado Pq Segundo A MARINA O frete so Acrescenta na Base dpis de reduzir o desconto acrescenta antes de reduzir
               '![Valor Do ICMS] = CalculaImposto(![Seqncia Do Conjunto], Orcamento![Seqncia Do Geral], 7, 2, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 3, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor Do Icms] = Round((![Valor da Base de Clculo] * ![Alquota Do ICMS] / 100), 2)
               Tributos = Tributos + ![Valor Do Icms]
               ![Percentual da Reduo] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 2, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            End If
               If Entrega_Futura Then
                  BC = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 6, 2, (!Quantidade * ![Valor Unitrio] + ValorAcresDesc), 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
                  BCAuxiliar = BC + ValorFrete
                  ' Alterado Pq Segundo A MARINA O frete so Acrescenta na Base dpis de reduzir o desconto acrescenta antes de reduzir
                  AliqAuxiliar = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 3, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
                  ICMSAuxiliar = Round((BCAuxiliar * AliqAuxiliar / 100), 2)
               End If
            ![Valor do IPI] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 8, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + ![Valor do IPI]
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 4, 2, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
             If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
               If Not Entrega_Futura Then 'Oramento Base: 125449 - Budke
                  'PisRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) * 48.1 / 100
                   PisRed = (Conjuntos_do_Orcamento!Quantidade * Conjuntos_do_Orcamento![Valor Unitrio] + ValorFrete - ![Valor Do Desconto] - ![Valor Do Icms]) * 48.1 / 100
                   ![Bc Pis] = Round(((Conjuntos_do_Orcamento!Quantidade * Conjuntos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] + ValorFrete - PisRed), 2)
                   '![Bc Pis] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) - PisRed
                   ![Aliq Do Pis] = 2
                   '![Valor Do //PIS] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 10, 2, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms] - PisRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
                   ![Valor Do PIS] = ![Bc Pis] * (![Aliq Do Pis] / 100)
               Else
                   PisRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Pis] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) - PisRed
                   ![Aliq Do Pis] = 2
                   ![Valor Do PIS] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 10, 2, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar - PisRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
               End If
             Else
               If Not Entrega_Futura Then
                   PisRed = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                   ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) - PisRed
                   ![Aliq Do Pis] = 2
                   ![Valor Do PIS] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms] - PisRed) * 2 / 100 'CalculaImposto(![Seqncia Do Conjunto], Orcamento![Seqncia Do Geral], 10, 2, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Else
                   PisRed = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) - PisRed
                   ![Aliq Do Pis] = 2
                   ![Valor Do PIS] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - PisRed) * 2 / 100
               End If
             End If
            Tributos = Tributos + ![Valor Do PIS]
             If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
               If Not Entrega_Futura Then
                   'CofinsRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) * 48.1 / 100
                   CofinsRed = (Conjuntos_do_Orcamento!Quantidade * Conjuntos_do_Orcamento![Valor Unitrio] + ValorFrete - ![Valor Do Desconto] - ![Valor Do Icms]) * 48.1 / 100
                 ' ![Bc Cofins] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) - CofinsRed
                   ![Bc Cofins] = Round(((Conjuntos_do_Orcamento!Quantidade * Conjuntos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] + ValorFrete - PisRed), 2)
                   ![Aliq Do Cofins] = 9.6
                   '![Valor Do Cofins] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms] - CofinsRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
                   ![Valor Do Cofins] = ![Bc Cofins] * (![Aliq Do Cofins] / 100)
               Else
                   CofinsRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Cofins] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) - CofinsRed
                   ![Aliq Do Cofins] = 9.6
                   ![Valor Do Cofins] = CalculaImposto(![Seqncia do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar - CofinsRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
               End If
             Else ' Nao tem disconto e frete
               If Not Entrega_Futura Then
                  CofinsRed = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                  ![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) - CofinsRed
                  ![Aliq Do Cofins] = 9.6
                  ![Valor Do Cofins] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms] - CofinsRed) * 9.6 / 100 'CalculaImposto(![Seqncia Do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Else
                  CofinsRed = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                  ![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) - CofinsRed
                  ![Aliq Do Cofins] = 9.6
                  ![Valor Do Cofins] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - CofinsRed) * 9.6 / 100 'CalculaImposto(![Seqncia Do Conjunto], Orcamento![Seqncia Do Geral], 11, 2, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               End If
             End If
            Tributos = Tributos + ![Valor Do Cofins]
            Tributos = Tributos + ![Valor ICMS ST]
            ![Valor Do Tributo] = Tributos 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
         End With
      Loop
      
      grdConjuntos.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
      
   End If
   
SaiDaSub:
   Screen.MousePointer = vbDefault

End Sub




Private Sub AtualizaValoresPecas()
   Dim Tributos As Double, ValorFrete As Double, ValorSeguro As Double, ValorAcresDesc As Double
   Dim PercentualFrete As Double, totalItens As Double, PercentualRateio As Double
   Dim VrItem As Double, BC As Double, ICMSAuxiliar As Double, BCAuxiliar As Double
   Dim AliqAuxiliar As Double, Geral As New GRecordSet, Suframa As Boolean
   Dim PisRed As Double, CofinsRed As Double
   Dim descCj As New GRecordSet, descPecas As New GRecordSet
   Dim descontoCj As Double, descontoPecas As Double, divergDesc As Double
          
   Screen.MousePointer = vbHourglass
   
   Call DistribuiDescontoTotal
   Set Geral = vgDb.OpenRecordSet("SELECT * From Geral Where [Seqncia Do Geral] = " & Orcamento![Seqncia Do Geral])
   Suframa = (CBool(Not Vazio(Geral![Cdigo Do Suframa])))
   If Suframa Then
      vgDb.Execute "Update [Peas do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
      grdPecas.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
         
   If Pecas_do_Orcamento.RecordCount > 0 Then
   
      If vgSituacao = -ACAO_INCLUINDO Or vgSituacao = -ACAO_EDITANDO Then
         If ValorAcresDesc = 0 And ValorFrete = 0 And ValorSeguro = 0 Then GoTo SaiDaSub
      End If
      
      'Rotina para corrigir o erro no rateio do desconto
      'Passo 1: Analisa se h tambm a presena conjuntos ou apenas peas
     If Orcamento![Valor Do Fechamento] <> 0 Then 'Se no tiver desconto no tem o que ser corrigido, nem perde tempo
        If Conjuntos_do_Orcamento.RecordCount > 0 Then
     '       descontoCj = verificaDesconto("Conjuntos do Oramento", Sequencia_do_Orcamento)
     '       descontoPecas = verificaDesconto("Peas do Oramento", Sequencia_do_Orcamento)
        Set descCj = vgDb.OpenRecordSet("SELECT SUM([Valor do Desconto]) AS descontoCj " & _
                                 "FROM [Conjuntos do Oramento] " & _
                                 "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)

        descontoCj = descCj!descontoCj

        Set descPecas = vgDb.OpenRecordSet("SELECT SUM([Valor do Desconto]) AS descontoPecas " & _
                                      "FROM [Peas do Oramento] " & _
                                      "WHERE [Seqncia do Oramento]  = " & Sequencia_do_Orcamento)

        descontoPecas = descPecas!descontoPecas
        
        'Para a primeira condio verdadeira o programa vai comparar o Valor do desconto
        'informado pelo vendas (aba Financeiro) com o somatrio do desconto rateado distribuido nos conjuntos e peas
        'a diferena ser armazenada em divergDesc, que ser utilizada para a correo do erro causado por arredondamentos
        If Orcamento![Valor Do Fechamento] <> (descontoCj + descontoPecas) Then
          If (Orcamento![Valor Do Fechamento] - descPecas!descontoPecas) > 0 Then Exit Sub
          divergDesc = Abs(Orcamento![Valor Do Fechamento]) - (descontoCj + descontoPecas)
          If vgPWUsuario = "YGOR" Then MsgBox (divergDesc)
            If divergDesc = 0 Then Exit Sub
        End If
      Else 'Para oramentos apenas com peas faz o mesmo procedimento, mas apenas com o somatrio
           'do desconto rateado entre as peas
        Set descPecas = vgDb.OpenRecordSet("SELECT SUM([Valor do Desconto]) AS descontoPecas " & _
                                      "FROM [Peas do Oramento] " & _
                                      "WHERE [Seqncia do Oramento]  = " & Sequencia_do_Orcamento)

        'descontoPecas = verificaDesconto("Peas do Oramento", Sequencia_do_Orcamento)
        
        
        If Orcamento![Valor Do Fechamento] <> descPecas!descontoPecas Then
         If (Orcamento![Valor Do Fechamento] - descPecas!descontoPecas) > 0 Then Exit Sub
          divergDesc = Abs(Orcamento![Valor Do Fechamento]) - descPecas!descontoPecas
          
          If vgPWUsuario = "YGOR" Then
            If divergDesc <> 0 Then
              MsgBox ("Erro de: " & divergDesc)
            Else
              MsgBox ("Corrigido.")
              Exit Sub
            End If
          End If
        End If
     End If 'sem desconto  pra cair aqui!
    End If
      Pecas_do_Orcamento.MoveFirst
      Do While Not Pecas_do_Orcamento.EOF
      
         
         Tributos = 0
         PisRed = 0
         CofinsRed = 0
         With Pecas_do_Orcamento
         .Edit
         'Valor do Desconto Rateado entre os produtos
         
         If Orcamento![Valor Do Fechamento] = 0 Then
            ![Valor Do Desconto] = 0
         End If
         If Orcamento![Valor Do Frete] = 0 Then
            ![Valor Do Frete] = 0
         End If
         
         If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
         If Orcamento!Fechamento = 1 Then 'Valor
            totalItens = Round((Orcamento![Valor Total dos Produtos] + Orcamento![Valor Total dos Conjuntos] + Orcamento![Valor Total das Peas] + Orcamento![Valor Total das Peas Usadas] + Orcamento![Valor Total de Produtos Usados] + Orcamento![Valor Total Conjuntos Usados]), 2)
            PercentualRateio = (Orcamento![Valor Do Fechamento] / totalItens * 100)
            VrItem = Round((Pecas_do_Orcamento!Quantidade * Pecas_do_Orcamento![Valor Unitrio]), 2)
            ValorAcresDesc = (VrItem * PercentualRateio / 100) ' Valor do Desconto
            
            ' Valor do Frete Rateado entre os Produtos
            PercentualFrete = (Orcamento![Valor Do Frete] / totalItens * 100)
            ValorFrete = (VrItem * PercentualFrete / 100) 'Valor do Frete
            ![Valor Do Desconto] = Abs(ValorAcresDesc)
            ![Valor Do Frete] = ValorFrete
            
               'Faz a correo no somatrio do desconto. Erro gerado por conta do arredondamento
               If Pecas_do_Orcamento.AbsolutePosition = Pecas_do_Orcamento.RecordCount And divergDesc <> 0 Then
                 '   divergDesc = vDesc + Orcamento![Valor do Fechamento]
                 '  If divergDesc <> 0 Then
                 ![Valor Do Desconto] = Abs(ValorAcresDesc) + divergDesc
               Else
                 ![Valor Do Desconto] = Abs(ValorAcresDesc)
                    'End If
               End If
            
        End If
        End If
            
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then
                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
               !IVA = 0
               ![Base de Clculo ST] = 0
               ![Valor ICMS ST] = 0
               ![Alquota Do ICMS ST] = 0
            Else
               '![Valor da Base de Clculo] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 6, 3, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               'Alterado em 15/07/2024, em conformidade com Maysa e Econet. Frete compondo a BC antes da reduo.
               BC = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 6, 3, (!Quantidade * ![Valor Unitrio] - ![Valor Do Desconto] + ValorFrete), 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor da Base de Clculo] = Round(BC, 2) '+ ValorFrete ' Alterado Pq Segundo A MARINA O frete so Acrescenta na Base dpis de reduzir o desconto acrescenta antes de reduzir
               '![Valor Do ICMS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 7, 3, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               ![Valor Do Icms] = Round((![Valor da Base de Clculo] * ![Alquota Do ICMS] / 100), 2)
               Tributos = Tributos + ![Valor Do Icms]
               ![Percentual da Reduo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 2, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            End If
            'Auxiliar ICMS
             If Entrega_Futura Then
                BC = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 6, 3, (!Quantidade * ![Valor Unitrio] + ValorAcresDesc), 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
                BCAuxiliar = BC + ValorFrete ' Alterado Pq Segundo A MARINA O frete so Acrescenta na Base dpis de reduzir o desconto acrescenta antes de reduzir
                AliqAuxiliar = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
                ICMSAuxiliar = Round((BCAuxiliar * AliqAuxiliar / 100), 2)
             End If
            ![Valor do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 8, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            Tributos = Tributos + ![Valor do IPI]
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 4, 3, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
            'Pis Antes do Oziris solicitar mudana 03/2018
             If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
               If Not Entrega_Futura Then
                  'PisRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) * 48.1 / 100
                  PisRed = (Pecas_do_Orcamento!Quantidade * Pecas_do_Orcamento![Valor Unitrio] + ValorFrete - ![Valor Do Desconto] - ![Valor Do Icms]) * 48.1 / 100
                  '![Bc Pis] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) - PisRed
                  '07/10/24 parei aqui - acertar BC PIS para reduzir o erro de arredondamento
                  ![Bc Pis] = Round(((Pecas_do_Orcamento!Quantidade * Pecas_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] - PisRed), 2)
                  ![Aliq Do Pis] = 2
                  ![Valor Do PIS] = Round((![Bc Pis] * ![Aliq Do Pis] / 100), 4) 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 3, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms] - PisRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
               Else
                  PisRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) * 48.1 / 100
                  ![Bc Pis] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) - PisRed
                  ![Aliq Do Pis] = 2
                  ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 3, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar - PisRed), 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, 0)
               End If
             Else ' sem rateio
               If Not Entrega_Futura Then
                  PisRed = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                  ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) - PisRed
                  ![Aliq Do Pis] = 2
                  ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms] - PisRed) * 2 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 3, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Else
                  PisRed = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                  ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) - PisRed
                  ![Aliq Do Pis] = 2
                  ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - PisRed) * 2 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 3, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               End If
                'Pis Novo Conforme contabilidade (Andreza) Subtari icms da base do pis
             End If
            Tributos = Tributos + ![Valor Do PIS]
             If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
               If Not Entrega_Futura Then
                   'CofinsRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) * 48.1 / 100
                   CofinsRed = (Pecas_do_Orcamento!Quantidade * Pecas_do_Orcamento![Valor Unitrio] + ValorFrete - ![Valor Do Desconto] - ![Valor Do Icms]) * 48.1 / 100
                   '![Bc Cofins] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms]) - CofinsRed
                   ![Bc Cofins] = Round(((Pecas_do_Orcamento!Quantidade * Pecas_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] - PisRed), 2)
                   ![Aliq Do Cofins] = 9.6
                   ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 3, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ![Valor Do Icms] - CofinsRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
               Else
                   CofinsRed = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Cofins] = (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar) - CofinsRed
                   ![Aliq Do Cofins] = 9.6
                   ![Valor Do Cofins] = Round((![Bc Cofins] * ![Aliq Do Cofins] / 100), 4) 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 3, (((VrItem + ValorFrete) - Abs(ValorAcresDesc)) - ICMSAuxiliar - CofinsRed), ValorAcresDesc, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF, ValorFrete)
               End If
             Else ' sem rateio
               If Not Entrega_Futura Then
                   CofinsRed = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                   ![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) - CofinsRed
                   ![Aliq Do Cofins] = 9.6
                   ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 '((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms] - CofinsRed) * 9.6 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 3, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               Else
                    DistribuiDescontoTotal
                   CofinsRed = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) - CofinsRed
                   ![Aliq Do Cofins] = 9.6
                   ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - CofinsRed) * 9.6 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 3, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], Orcamento![Seqncia da Classificao], , MunicipioAux!UF)
               End If
             End If
            Tributos = Tributos + ![Valor Do Cofins]
            Tributos = Tributos + ![Valor ICMS ST]
            ![Valor Do Tributo] = Round(Tributos, 2) 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
         End With
      Loop
      
      grdPecas.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
      
   End If
   
SaiDaSub:
   Screen.MousePointer = vbDefault

End Sub


Private Sub AtualizaValoresProdutos()
   Dim Tributos As Double, ValorFrete As Double, ValorSeguro As Double, ValorAcresDesc As Double
   Dim totalItens As Double, PercentualRateio As Double, VrItem As Double, VrReducao As Double
   Dim PercentualFrete As Double, Geral As New GRecordSet, GRevenda As Boolean, ICMSAuxiliar As Double
   Dim GeralS As New GRecordSet, Suframa As Boolean
   Dim ReducaoAuxiliar As Double, RedPis As Double, RedCofins As Double
   Dim VrItem2 As Double
   Dim vDesc As Double, divergDesc As Double
   
         
   Screen.MousePointer = vbHourglass
   
   divergDesc = 0
   vDesc = 0
   VrItem = 0
   RedPis = 0
   RedCofins = 0
   Call DistribuiDescontoTotal
   Set GeralS = vgDb.OpenRecordSet("SELECT * From Geral Where [Seqncia Do Geral] = " & Orcamento![Seqncia Do Geral])
   Suframa = (CBool(Not Vazio(GeralS![Cdigo Do Suframa])))
   If Suframa Then
      vgDb.Execute "Update [Produtos do Oramento] Set [Percentual da Reduo] = 0 WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
      GrdProdutos.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
       
   If Produtos_do_Orcamento.RecordCount > 0 Then
      If vgSituacao = -ACAO_INCLUINDO Or vgSituacao = -ACAO_EDITANDO Then
         If ValorAcresDesc = 0 And ValorFrete = 0 And ValorSeguro = 0 Then GoTo SaiDaSub
      End If
      Produtos_do_Orcamento.MoveFirst
      Do While Not Produtos_do_Orcamento.EOF
         
         
         Tributos = 0
         RedPis = 0
         RedCofins = 0
         TbAuxiliar "Produtos", "[Seqncia do Produto] = " & Produtos_do_Orcamento![Seqncia do Produto], ProdutoAux
         TbAuxiliar "Classificao Fiscal", "[Seqncia da Classificao] = " & ProdutoAux![Seqncia da Classificao] & " AND [Seqncia da Classificao] > 0", ProdutoNCMAux
         With Produtos_do_Orcamento
         .Edit
         
         If Orcamento![Valor Do Fechamento] = 0 Then
            ![Valor Do Desconto] = 0
         End If
         If Orcamento![Valor Do Frete] = 0 Then
            ![Valor Do Frete] = 0
         End If
          
         'Valor do Desconto Rateado entre os produtos
         If Orcamento![Valor Do Fechamento] < 0 Or Orcamento![Valor Do Frete] > 0 Then
         If Orcamento!Fechamento = 1 Then 'Valor
            Set Geral = vgDb.OpenRecordSet("SELECT * From Geral Where [Seqncia Do Geral] = " & Orcamento![Seqncia Do Geral])
            GRevenda = Geral!Revenda
            totalItens = Round((Orcamento![Valor Total dos Produtos] + Orcamento![Valor Total dos Conjuntos] + Orcamento![Valor Total das Peas] + Orcamento![Valor Total das Peas Usadas] + Orcamento![Valor Total de Produtos Usados] + Orcamento![Valor Total Conjuntos Usados]), 2)
            PercentualRateio = (Orcamento![Valor Do Fechamento] / totalItens * 100)
            VrItem = (Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio])
            ValorAcresDesc = (VrItem * PercentualRateio / 100) ' Valor do Desconto
            'vDesc = Round((vDesc + ValorAcresDesc), 2)
            
            'Valor do Frete Rateado entre os Produtos
            PercentualFrete = (Orcamento![Valor Do Frete] / totalItens * 100)
            ValorFrete = (VrItem * PercentualFrete / 100) 'Valor do Frete
            
            VrItem2 = VrItem + ValorFrete + ValorAcresDesc
            'alterado em 16/08/2024 com base no oramento 120949 - Correo no clculo do IPI considerando ValorFrete na base de clculo
            ReducaoAuxiliar = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 2, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            ![Valor do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 8, 1, !Quantidade * ![Valor Unitrio] + ValorAcresDesc + ValorFrete, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
           '![Valor Do IPI] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 8, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            Tributos = Tributos + ![Valor do IPI]
            
            ![Valor Do Desconto] = Abs(ValorAcresDesc)
            ![Valor Do Frete] = ValorFrete
            vDesc = vDesc + ![Valor Do Desconto]
            
         'Faz a correo no somatrio do desconto. Erro gerado por conta do arredondamento
         If Produtos_do_Orcamento.AbsolutePosition = Produtos_do_Orcamento.RecordCount Then
              divergDesc = vDesc + Orcamento![Valor Do Fechamento]
             If divergDesc <> 0 Then
               ![Valor Do Desconto] = Abs(ValorAcresDesc + divergDesc)
             Else
               ![Valor Do Desconto] = Abs(ValorAcresDesc)
             End If
         End If
          ' Verifica se Tem Reduo da BC
         If ReducaoAuxiliar > 0 Then 'Alterado em 15/08/24 - Oramento base.: 121089
            VrItem = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) + Produtos_do_Orcamento![Valor do IPI] + (ValorAcresDesc + divergDesc) + ValorFrete), 2)
            VrReducao = Round((VrItem * ReducaoAuxiliar / 100), 2)
            VrItem = VrItem - VrReducao '+ ValorFrete (alterado em 15/08/24) - Oramento base.: 121089
         Else
          If Not GRevenda Then
             VrItem = ((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) + Produtos_do_Orcamento![Valor do IPI] - ![Valor Do Desconto]) '+ Int((ValorAcresDesc * 100) + 1) / 100)
             VrItem = VrItem + ValorFrete
          Else
            VrItem = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) + ValorAcresDesc), 2)
            VrItem = VrItem + ValorFrete
          End If
         End If
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then
                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ICMSAuxiliar = Round((VrItem * CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF) / 100), 2)
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
            Else ' Nw  Entrega Futura
              !CFOP = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 1, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
              If !CFOP = "5403" Or !CFOP = "6403" Or !Diferido = True Then
                 ![Valor da Base de Clculo] = 0
                 ![Valor Do Icms] = 0
              Else
              ![Valor da Base de Clculo] = VrItem
              ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
              ![Valor Do Icms] = Round((VrItem * Produtos_do_Orcamento![Alquota Do ICMS] / 100), 2)
               '![Valor Do Icms] = Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio] - ![Valor Do Desconto] -
               If ![Alquota Do ICMS] = 0 Then
                  ![Valor da Base de Clculo] = 0
               End If
              End If
               Tributos = Tributos + Produtos_do_Orcamento![Valor Do Icms]
              '![Alquota Do ICMS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 3, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
              ![Percentual da Reduo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 2, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            End If
            '![Valor Do IPI] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 8, 1, !Quantidade * ![Valor Unitrio] + ValorAcresDesc, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            'Tributos = Tributos + ![Valor Do IPI]
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 4, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            'Pis Antes ate 03/2018
             '![Valor Do PIS] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            'Mudana 19/03/2018 a Pedido do Oziris subtrair o icms da base do pis
            'If Not Entrega_Futura Then
            '   ![Valor Do PIS] = ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ![Valor Do ICMS]) * 1.65 / 100'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
            'Else
            '   ![Valor Do PIS] = ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar) * 1.65 / 100'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
            'End If
             If Not Entrega_Futura Then
                If ReducaoAuxiliar = 0 Then 'aqui 2
                 
                   If Not GRevenda Then
                      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedPis = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Pis] = VrItem - ![Valor Do Icms] - RedPis
                         ![Aliq Do Pis] = 2
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100
                      Else
                       ' 19/09/2024 Ajuste na base do clculo com base no ormento 124582 - no estava considerando o frete na BC do PIS e COFINS
                       '![Bc Pis] = VrItem2 - ![Valor Do Icms]
                        ![Bc Pis] = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] + ValorFrete), 2)
                        ![Aliq Do Pis] = 1.65
                        ![Valor Do PIS] = Round((![Bc Pis] * ![Aliq Do Pis] / 100), 4)
                      End If
                   Else
                      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                          RedPis = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Pis] = (VrItem2 - ![Valor Do Icms]) - RedPis '((VrItem) - ![Valor Do ICMS] - RedPis)
                         ![Aliq Do Pis] = 2
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100
                      Else
                         ![Bc Pis] = VrItem2 - ![Valor Do Icms]
                         ![Aliq Do Pis] = 1.65
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100
                      End If
                   End If
                   If !CFOP = "5551" Or !CFOP = "6551" Then
                      ![Valor Do PIS] = 0
                   End If
                Else ' com Reduo
                   If Not GRevenda Then
                      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedPis = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Pis] = (VrItem2 - ![Valor Do Icms]) - RedPis
                         ![Aliq Do Pis] = 2
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100
                      Else
                        ![Bc Pis] = Round((VrItem2 + divergDesc - ![Valor Do Icms]), 2)
                        '![Bc Pis] = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms]), 2) '+ Int((ValorAcresDesc * 100) + 1) / 100) - Essa lgica fora o arredondamento sempre pra cima, mas no foi preciso utilizar, pois o erro que era gerado no (vr unit * qtde) para o vr total foi corrigido
                         ![Aliq Do Pis] = 1.65
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '(((!Quantidade * ![Valor Unitrio]) - Produtos_do_Orcamento![Valor Do IPI] + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS]) * 1.65 / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                      End If
                   Else
                      If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedPis = VrItem2 * 48.1 / 100
                         ![Bc Pis] = VrItem2 - ![Valor Do Icms] - RedPis
                         ![Aliq Do Pis] = 2
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '(((!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS] - RedPis) * 2 / 100
                      Else
                         ![Bc Pis] = VrItem2 - ![Valor Do Icms]
                         ![Aliq Do Pis] = 1.65
                         ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '(((!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS]) * 1.65 / 100
                      End If
                   End If
                   If !CFOP = "5551" Or !CFOP = "6551" Then
                      ![Valor Do PIS] = 0
                   End If
                End If
              Else 'Entrega Futura
                If ReducaoAuxiliar = 0 Then
                   If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                       RedPis = (VrItem2 - ICMSAuxiliar) * 48.1 / 100
                      ![Aliq Do Pis] = 2
                      ![Bc Pis] = VrItem2 - ICMSAuxiliar - RedPis '((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar - RedPis)
                      ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100
                   Else
                     ![Bc Pis] = VrItem2 - ICMSAuxiliar
                     ![Aliq Do Pis] = 1.65
                     ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar) * 1.65 / 100
                   End If
                Else
                If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                    RedPis = (VrItem2 - ICMSAuxiliar) * 48.1 / 100
                   ![Bc Pis] = (VrItem2 - ICMSAuxiliar) - RedPis
                   ![Aliq Do Pis] = 2
                   ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '(((!Quantidade * ![Valor Unitrio] - Produtos_do_Orcamento![Valor Do IPI]) + ValorFrete + ValorAcresDesc) - ICMSAuxiliar - RedPis) * 2 / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                Else
                   ![Aliq Do Pis] = 1.65
                   ![Bc Pis] = VrItem2 - ICMSAuxiliar
                   ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                End If
                End If
             End If
             If !CFOP = "5551" Or !CFOP = "6551" Then
                ![Valor Do PIS] = 0
             End If
            Tributos = Tributos + ![Valor Do PIS]
            'Cofins Antes ate 03/2018
            '![Valor Do Cofins] = CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            'Mudana 19/03/2018 a Pedido do Oziris subtrair o icms da base do Cofins
            If Not Entrega_Futura Then
               If ReducaoAuxiliar = 0 Then
                  If Not GRevenda Then
                     If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedCofins = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                         ![Aliq Do Cofins] = 9.6
                         ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ![Valor Do ICMS] - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     Else
                         'Ajuste BC considerando o frete. Oramento base 124582 em 19/09/2024
                         ![Bc Cofins] = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms] - ![Valor Do Desconto] + ValorFrete), 2)
                        '![Bc Cofins] = Round(VrItem2, 2) - ![Valor Do Icms]
                         ![Aliq Do Cofins] = 7.6
                         ![Valor Do Cofins] = Round((![Bc Cofins] * ![Aliq Do Cofins] / 100), 4) '((!Quantidade * ![Valor Unitrio]) - ![Valor Do ICMS]) * 7.6 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
                         '![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     End If
                  Else
                  If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                      RedCofins = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                      ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                      ![Aliq Do Cofins] = 9.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem) - ![Valor Do ICMS] - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  Else
                      ![Bc Cofins] = VrItem2 - ![Valor Do Icms]
                      ![Aliq Do Cofins] = 7.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  End If
                  End If
               Else 'Reducao
                  If Not GRevenda Then
                     If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedCofins = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                         ![Aliq Do Cofins] = 9.6
                         ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio] - Produtos_do_Orcamento![Valor Do IPI])) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS] - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     Else
                         ![Bc Cofins] = Round((VrItem2 + divergDesc - ![Valor Do Icms]), 2)
                         ![Aliq Do Cofins] = 7.6
                         ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio] - Produtos_do_Orcamento![Valor Do IPI])) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     End If
                  Else
                     'Continua aqui.... 13/12/23....
                     If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                         RedCofins = (VrItem2 - ![Valor Do Icms]) * 48.1 / 100
                         ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                         ![Aliq Do Cofins] = 9.6
                         ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100  'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio])) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS] - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     Else
                         ![Bc Cofins] = VrItem2 - ![Valor Do Icms]
                         ![Aliq Do Cofins] = 7.6
                         ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio])) + ValorFrete + ValorAcresDesc) - ![Valor Do ICMS]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                     End If
                  End If
               End If
            Else 'Entrega Futura
               If ReducaoAuxiliar = 0 Then
                  If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                      RedCofins = (VrItem2 - ICMSAuxiliar) * 48.1 / 100
                      ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                      ![Aliq Do Cofins] = 9.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  Else
                      ![Bc Cofins] = VrItem2 - ICMSAuxiliar
                      ![Aliq Do Cofins] = 7.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((VrItem - Produtos_do_Orcamento![Valor Do IPI]) - ICMSAuxiliar), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  End If
               Else
                  If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                      RedCofins = (VrItem2 - ICMSAuxiliar) * 48.1 / 100
                      ![Bc Cofins] = VrItem2 - ![Valor Do Icms] - RedCofins
                      ![Aliq Do Cofins] = 9.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio]) - Produtos_do_Orcamento![Valor Do IPI]) + ValorFrete + ValorAcresDesc) - ICMSAuxiliar - RedCofins), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  Else
                      ![Bc Cofins] = VrItem2 - ICMSAuxiliar
                      ![Aliq Do Cofins] = 7.6
                      ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, ((((!Quantidade * ![Valor Unitrio]) - Produtos_do_Orcamento![Valor Do IPI]) + ValorFrete + ValorAcresDesc) - ICMSAuxiliar), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, 0)
                  End If
               End If
            End If
            If !CFOP = "5551" Or !CFOP = "6551" Then
                ![Valor Do Cofins] = 0
            End If
            Tributos = Tributos + ![Valor Do Cofins]
            ' Substituio Tributaria Com Frete e Desconto AFF
            !IVA = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 12, 1, (!Quantidade * ![Valor Unitrio]), ValorAcresDesc, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, ValorFrete)
            ![Base de Clculo ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 13, 1, (!Quantidade * ![Valor Unitrio]), ValorAcresDesc, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, ValorFrete)
            ![Alquota Do ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 15, 1, (!Quantidade * ![Valor Unitrio]), ValorAcresDesc, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, ValorFrete)
            ![Valor ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 14, 1, (!Quantidade * ![Valor Unitrio]), ValorAcresDesc, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF, ValorFrete)
            Tributos = Tributos + ![Valor ICMS ST]
            ![Valor Do Tributo] = Round((Tributos), 2) 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
            End If
            End If
            'Qdo No Tem Desconto
            If Orcamento![Valor Do Fechamento] = 0 Then
            If Orcamento![Valor Do Frete] = 0 Then
              
            If Entrega_Futura Then
               If MunicipioAux!UF = "SP" Then
                  !CFOP = "5922"
                  !CST = 90
               Else
                  !CFOP = "6922"
                  !CST = 90
               End If
               ![Valor da Base de Clculo] = 0
               ![Valor Do Icms] = 0
                ICMSAuxiliar = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 7, 1, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               ![Alquota Do ICMS] = 0
               ![Percentual da Reduo] = 0
            Else
               !CFOP = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 1, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               ![Valor da Base de Clculo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 6, 1, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               ![Valor Do Icms] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 7, 1, (!Quantidade * ![Valor Unitrio]) + ValorFrete + ValorAcresDesc + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               Tributos = Tributos + ![Valor Do Icms]
               ![Alquota Do ICMS] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 3, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               ![Percentual da Reduo] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 2, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            End If
            ![Valor do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 8, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            Tributos = Tributos + ![Valor do IPI]
            ![Alquota Do IPI] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 4, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            
            If Entrega_Futura Then
               If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                  RedPis = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                  ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - RedPis)
                  ![Aliq Do Pis] = 2
                  ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - RedPis) * 2 / 100
               Else
                  ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar)
                  ![Aliq Do Pis] = 1.65
                  ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 1.65 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               End If
               If !CFOP = "5551" Or !CFOP = "6551" Then
                  ![Valor Do PIS] = 0
               End If
               Tributos = Tributos + ![Valor Do PIS]
               If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then 'aqui
                  RedCofins = ((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 48.1 / 100
                  ![Bc Cofins] = (!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - RedCofins
                  ![Aliq Do Cofins] = 9.6
                  ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar - RedCofins) * 9.6 / 100
               Else
                  ![Bc Cofins] = (!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar
                  ![Aliq Do Cofins] = 7.6
                  ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 '((!Quantidade * ![Valor Unitrio]) - ICMSAuxiliar) * 7.6 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
               End If
               Tributos = Tributos + ![Valor Do Cofins]
            Else 'nao  entrega futura
             If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                 RedPis = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                 ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms] - RedPis)
                 ![Aliq Do Pis] = 2
                 ![Valor Do PIS] = ![Bc Pis] * ![Aliq Do Pis] / 100 '((!Quantidade * ![Valor Unitrio]) - ![Valor Do ICMS] - RedPis) * 2 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
             Else 'orc 125012
'                ![Bc Pis] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms])
                 ![Bc Pis] = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms]), 2) '+ Int((ValorAcresDesc * 100) + 1) / 100) - Essa lgica fora o arredondamento sempre pra cima, mas no foi preciso utilizar, pois o erro que era gerado no (vr unit * qtde) para o vr total foi corrigido
                 ![Aliq Do Pis] = 1.65
                 ![Valor Do PIS] = Round((![Bc Pis] * ![Aliq Do Pis] / 100), 4) '((!Quantidade * ![Valor Unitrio]) - ![Valor Do ICMS]) * 1.65 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 10, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
             End If
             If !CFOP = "5551" Or !CFOP = "6551" Then
                ![Valor Do PIS] = 0
             End If
               Tributos = Tributos + ![Valor Do PIS]
             If (Mid(ProdutoNCMAux!Ncm, 1, 5) = "84248" Or Mid(ProdutoNCMAux!Ncm, 1, 4) = "7309" Or ProdutoNCMAux!Ncm = "87162000") And Not ProdutoAux![Material Adquirido de Terceiro] Then
                 RedCofins = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) * 48.1 / 100
                 ![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms]) - RedCofins
                 ![Aliq Do Cofins] = 9.6
                 ![Valor Do Cofins] = ![Bc Cofins] * ![Aliq Do Cofins] / 100 '((!Quantidade * ![Valor Unitrio]) - ![Valor Do ICMS] - RedCofins) * 9.6 / 100
             Else
                '![Bc Cofins] = ((!Quantidade * ![Valor Unitrio]) - ![Valor Do Icms])
                 ![Bc Cofins] = Round(((Produtos_do_Orcamento!Quantidade * Produtos_do_Orcamento![Valor Unitrio]) - ![Valor Do Icms]), 2) ' + Int((ValorAcresDesc * 100) + 1) / 100)
                 ![Aliq Do Cofins] = 7.6
                 ![Valor Do Cofins] = Round((![Bc Cofins] * ![Aliq Do Cofins] / 100), 4) '((!Quantidade * ![Valor Unitrio]) - ![Valor Do ICMS]) * 7.6 / 100 'CalculaImposto(![Seqncia Do Produto], Orcamento![Seqncia Do Geral], 11, 1, (!Quantidade * ![Valor Unitrio]) + ValorAcresDesc + ValorFrete + ValorSeguro, 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
             End If
             If !CFOP = "5551" Or !CFOP = "6551" Then
                ![Valor Do Cofins] = 0
             End If
             Tributos = Tributos + ![Valor Do Cofins]
            End If
            If Not Entrega_Futura Then
               !CST = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 5, 1, !Quantidade * ![Valor Unitrio], 0, Orcamento![Seqncia da Propriedade])
            End If
            'Substituio
            !IVA = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 12, 1, (!Quantidade * ![Valor Unitrio]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            ![Base de Clculo ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 13, 1, (!Quantidade * ![Valor Unitrio]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            ![Alquota Do ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 15, 1, (!Quantidade * ![Valor Unitrio]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            ![Valor ICMS ST] = CalculaImposto(![Seqncia do Produto], Orcamento![Seqncia Do Geral], 14, 1, (!Quantidade * ![Valor Unitrio]), 0, Orcamento![Seqncia da Propriedade], , , MunicipioAux!UF)
            Tributos = Tributos + ![Valor ICMS ST]
            ![Valor Do Tributo] = Round(Tributos, 2) 'Tributos
            .Update
            .BookMark = .LastModified
            .MoveNext
            End If
            End If
    End With
      Loop
      GrdProdutos.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulrio
   End If
SaiDaSub:
   Screen.MousePointer = vbDefault
End Sub









Private Function ValidaDesconto() As Boolean
 Dim DescAux As Double
 Dim Tb As New GRecordSet
 Dim Tipo As String
 
 If Valor_do_Fechamento = 0 Then
    ValidaDesconto = True
    Exit Function
 End If
 
 If (Valor_Total_do_Orcamento = 0) Or vgPWUsuario = "WAGNER" Or vgPWUsuario = "DIEGO DELMONACO" Or vgPWUsuario = "YGOR" Or vgPWUsuario = "ISABELA MONTEIRO" Or vgPWUsuario = "JUCELI" Then
    ValidaDesconto = True
    Exit Function
 End If
 
Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
     If Tb.RecordCount > 0 Then
        Tipo = "P"
        Set Tb = Nothing
     End If
     
Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Peas do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
     If Tb.RecordCount > 0 Then
        Tipo = "Pc"
        Set Tb = Nothing
     End If
  
Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Conjuntos do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
     If Tb.RecordCount > 0 Then
        Tipo = "Conj"
        Set Tb = Nothing
     End If
     
Set Tb = Nothing
  
 'If Tipo = "P" Then
    DescAux = Valor_Total_do_Orcamento * 3 / 100
 'ElseIf Tipo = "Pc" Then
 '   DescAux = Valor_Total_Do_Orcamento * 10 / 100
 'ElseIf Tipo = "Conj" Then
 '   DescAux = Valor_Total_Do_Orcamento * 7.5 / 100
' End If
    
 ValidaDesconto = Abs(Valor_do_Fechamento) <= DescAux
  
End Function



Private Function MostraVinculo() As String
 Dim Tb As New GRecordSet
 Dim Tb1 As New GRecordSet
 
 If Orcamento_Vinculado = 0 Then
    MostraVinculo = ""
 End If
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Seqncia do Geral] Seq From Oramento Where [Seqncia do Oramento] = " & Orcamento_Vinculado)
 Set Tb1 = vgDb.OpenRecordSet("SELECT [Razo Social] Nome From Geral WHERE [Seqncia do Geral] = " & Tb!Seq)
 
 MostraVinculo = Tb1!Nome
 
End Function



Private Sub VerificaVinculo()
 Dim Produto As New GRecordSet
 Dim Pecas As New GRecordSet
 Dim Conj As New GRecordSet
 
 Set Produto = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Oramento] Where [Seqncia do Oramento] = " & Orcamento_Vinculado)
 Set Pecas = vgDb.OpenRecordSet("SELECT * FROM [Peas do Oramento] Where [Seqncia do Oramento] = " & Orcamento_Vinculado)
 Set Conj = vgDb.OpenRecordSet("SELECT * FROM [Conjuntos do Oramento] Where [Seqncia do Oramento] = " & Orcamento_Vinculado)
 

End Sub



Private Sub PreValidaVinculo()
   Dim Tb As New GRecordSet, Itens As New GRecordSet
   Dim Vector() As String, TbOrigem As New GRecordSet
   Dim i As Long, vaPrivez As Boolean, Campo As Variant, Mensagem As String
   
   On Error GoTo DeuErro
   
   If Orcamento_Vinculado = 0 Then
      Exit Sub
   End If
     
   Set Itens = vgDb.OpenRecordSet("SELECT * FROM(" & _
                                  "SELECT PO.[Seqncia do Oramento], 'Pecas' Tipo, P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], PO.[Valor Unitrio] " & _
                                  "FROM [Peas do Oramento] PO LEFT JOIN Produtos P ON PO.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "UNION ALL " & _
                                  "SELECT Pro.[Seqncia do Oramento], 'Produto' Tipo, P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], Pro.[Valor Unitrio] " & _
                                  "FROM [Produtos do Oramento] Pro LEFT JOIN Produtos P ON Pro.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "UNION ALL " & _
                                  "SELECT CO.[Seqncia do Oramento], 'Conjuntos' Tipo, C.[Seqncia Do Conjunto], C.Descrio, U.[Sigla da Unidade], CO.[Valor Unitrio] " & _
                                  "FROM [Conjuntos do Oramento] CO LEFT JOIN Conjuntos C ON CO.[Seqncia Do Conjunto] = C.[Seqncia Do Conjunto] " & _
                                  "LEFT JOIN Unidades U ON C.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "LEFT JOIN Oramento O ON CO.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                  ") A " & _
                                  "WHERE [Seqncia do Oramento] = " & Orcamento_Vinculado)
                                                               
   If Itens.RecordCount = 0 Then
      MsgBox ("Oramento Vinculado Invalido!")
      mdiIRRIG.CancelaAlteracoes
      Exit Sub
   End If
   
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
 
   Do While Not Itens.EOF
   Set TbOrigem = vgDb.OpenRecordSet("SELECT * FROM(" & _
                                     "SELECT PO.[Seqncia do Oramento], 'Pecas' Tipo, P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], PO.[Valor Unitrio] " & _
                                     "FROM [Peas do Oramento] PO LEFT JOIN Produtos P ON PO.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                     "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                     "UNION ALL " & _
                                     "SELECT Pro.[Seqncia do Oramento], 'Produto' Tipo, P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], Pro.[Valor Unitrio] " & _
                                     "FROM [Produtos do Oramento] Pro LEFT JOIN Produtos P ON Pro.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                     "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                     "UNION ALL " & _
                                     "SELECT CO.[Seqncia do Oramento], 'Conjuntos' Tipo, C.[Seqncia Do Conjunto], C.Descrio, U.[Sigla da Unidade], CO.[Valor Unitrio] " & _
                                     "FROM [Conjuntos do Oramento] CO LEFT JOIN Conjuntos C ON CO.[Seqncia Do Conjunto] = C.[Seqncia Do Conjunto] " & _
                                     "LEFT JOIN Unidades U ON C.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                     "LEFT JOIN Oramento O ON CO.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                     ") A " & _
                                     "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
                                     
  '  If Itens!Tipo <> TbOrigem!Tipo Then
  '     MsgBox ("Oramento Vinculado Invalido!")
  '     mdiIRRIG.CancelaAlteracoes
  '     Exit Sub
  '  End If
    If Itens!Tipo = "Conjuntos" Then
       If Itens![Seqncia do Produto] <> TbOrigem![Seqncia do Produto] Then
          i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM INVALIDO: " & Itens![Seqncia do Produto] & " - " & Itens!Descrio
       End If
          If Itens!Tipo = "Conjuntos" And TbOrigem!Tipo = "Conjuntos" Then
             If Itens![Seqncia do Produto] = TbOrigem![Seqncia do Produto] Then
                Exit Sub
             End If
          End If
    End If
                                    
   Itens.MoveNext
   Loop
      
   If UBound(Vector) > 0 Then
      Mensagem = "Alguns Itens (Divirgentes):" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      mdiIRRIG.CancelaAlteracoes
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
   
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub



Private Sub PreValidaNFE()
   Dim Tb As New GRecordSet, Itens As New GRecordSet
   Dim Vector() As String
   Dim i As Long, vaPrivez As Boolean, Campo As Variant, Mensagem As String
   Dim Servicos As New GRecordSet
   
   On Error GoTo DeuErro
   'Verifica obrigatoriedade dos campos de Peso Bruto, Peso Lquido e Volumes
   If Trim(txtCampo(143).Text) = "" Or Trim(txtCampo(144).Text) = "" Or _
      Trim(txtCampo(145).Text) = "" Or _
      Val(txtCampo(143).Text) = 0 Or Val(txtCampo(144).Text) = 0 Or Val(txtCampo(145).Text) = 0 Then

      MsgBox "Preencha Peso Bruto, Peso Lquido e Volumes antes de gerar a nota." & vbCrLf & _
             "Os valores no podem ficar em branco nem ser zero.", _
             vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
'Verifica obrigatoriedade dos campos de Placa, UF da Placa e ANTT
If Trim(txtCp(152).Text) = "" Or Trim(txtCp(153).Text) = "" Then
   MsgBox "Preencha Placa e UF da Placa antes de gerar a nota.", _
          vbExclamation + vbOKOnly, vaTitulo
   Exit Sub
End If
  
   Set Itens = vgDb.OpenRecordSet("SELECT * FROM(" & _
                                  "SELECT PO.[Seqncia do Oramento], 'Produto' Tipo, P.[Receita Conferida], P.[Seqncia do Produto], P.Descrio, U.[Sigla da Unidade], PO.[Valor Unitrio] " & _
                                  "FROM [Produtos do Oramento] PO LEFT JOIN Produtos P ON PO.[Seqncia Do Produto] = P.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON P.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "UNION ALL " & _
                                  "SELECT CO.[Seqncia do Oramento],'Conjuntos' Tipo, C.[Receita Conferida], C.[Seqncia Do Conjunto], C.Descrio, U.[Sigla da Unidade], CO.[Valor Unitrio] " & _
                                  "FROM [Conjuntos do Oramento] CO LEFT JOIN Conjuntos C ON CO.[Seqncia Do Conjunto] = C.[Seqncia Do Conjunto] " & _
                                  "LEFT JOIN Unidades U ON C.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "LEFT JOIN Oramento O ON CO.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                  "UNION ALL " & _
                                  "SELECT PeOrc.[Seqncia do Oramento],'Peas' Tipo, Pe.[Receita Conferida], Pe.[Seqncia Do Produto], Pe.Descrio, U.[Sigla da Unidade], PeOrc.[Valor Unitrio] " & _
                                  "FROM [Peas do Oramento] PeOrc LEFT JOIN Produtos Pe ON PeOrc.[Seqncia Do Produto] = Pe.[Seqncia Do Produto] " & _
                                  "LEFT JOIN Unidades U ON Pe.[Seqncia da Unidade] = U.[Seqncia da Unidade] " & _
                                  "LEFT JOIN Oramento O ON PeOrc.[Seqncia do Oramento] = O.[Seqncia do Oramento] " & _
                                  ") A " & _
                                  "WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
                                  
   Set Servicos = vgDb.OpenRecordSet("SELECT * FROM [Servios do Oramento] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
                                                               
   If Itens.RecordCount = 0 And Servicos.RecordCount = 0 Then Exit Sub
   
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
 
   Do While Not Itens.EOF
      'If Itens![Valor Unitrio] = 0 Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM: " & Itens![Seqncia Do Produto] & " - " & Itens!Descrio
      If Itens![Receita Conferida] = 0 Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM: " & Itens![Seqncia do Produto] & " - " & Itens!Descrio
      Itens.MoveNext
   Loop
      
   If UBound(Vector) > 0 And vgPWUsuario <> "YGOR" Then
      Mensagem = "Alguns Itens (sem conferencia da Receita):" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
   
   AbreGerar
    
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub



Private Sub DesmarcaFw()
 Dim Tb As New GRecordSet

 MYSQLInclui
 Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Follow Up Vendas] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento)
 
   If Not Venda_Fechada Then
      vgDb.Execute "DELETE FROM [Follow Up Vendas] WHERE [Seqncia do Oramento] = " & Sequencia_do_Orcamento
      MsgBox ("Pedido excluido do Follow Up!")
   End If
   
End Sub



Private Sub MYSQLInclui()
 Dim SQL As String
 Dim Tb1 As New GRecordSet
 Dim SeqItem As Integer
 Dim cnGas As ADODB.Connection, rsGas As ADODB.RecordSet
 
 Set Tb1 = vgDb.OpenRecordSet("SELECT * From [Itens pendentes] Where [Seqncia Do Oramento] = " & Sequencia_do_Orcamento)
 
 If Tb1.RecordCount > 0 Then
    Exit Sub
 Else
    Set Tb1 = Nothing
 End If
 
 Set cnGas = New ADODB.Connection
 Set rsGas = New ADODB.RecordSet
 
 If EstaEmIDE Then 'MERDA
    'cnGas.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER={SQL Server};Server=DESKTOP-NGAVDP6\SQLEXPRESS2014;Integrated Security=SSPI;Trusted_Connection=YES"'"Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER=[DRIVER];Server=DESKTOP-NGAVDP6\SQLEXPRESS2014;UID=ygor;Pwd=5139249_;Database=IRRIGACAO;Trusted_Connection=YES"
    cnGas.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER=[DRIVER];Server=DESKTOP-CHS14C0\SQLIRRIGACAO;Integrated Security=SSPI;Trusted_Connection=YES;Database=IRRIGACAO"
 Else
    cnGas.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;DRIVER=[DRIVER];Server=SRVSQL\SQLEXPRESS;UID=ygor;Pwd=5139249_;Database=IRRIGACAO;Trusted_Connection=NO"
 End If
 cnGas.Open
                                                                                                                                                                                                          
 Set Tb1 = vgDb.OpenRecordSet("SELECT * From [Itens pendentes]")
 
 SeqItem = 1
 SQL = "SELECT Tipo, [Seqncia Do Produto], Descrio, Quantidade, Total, [Valor Unitrio] " & _
       "FROM( "
       
 SQL = SQL & "SELECT 1 As Tipo, [Peas do Oramento].[Seqncia do Produto], Produtos.Descrio, Sum([Peas do Oramento].Quantidade) Quantidade, " & _
             "Sum([Peas do Oramento].[Valor Total]) As Total, [Peas do Oramento].[Valor Unitrio] " & _
             "FROM [Peas do Oramento] left join Oramento On [Peas do Oramento].[Seqncia do Oramento] = Oramento.[Seqncia do Oramento] " & _
             "Left Join Produtos On [Peas do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] WHERE " & Filtro1()
 SQL = SQL & " Group By [Peas do Oramento].[Seqncia Do Produto], Produtos.Descrio, [Peas do Oramento].[Valor Unitrio] Union "
 SQL = SQL & "Select 2 As Tipo, [Conjuntos do Oramento].[Seqncia Do Conjunto], Conjuntos.Descrio, Sum([Conjuntos do Oramento].Quantidade) Quantidade, Sum([Conjuntos do Oramento].[Valor Total]) Total, [Conjuntos do Oramento].[Valor Unitrio] " & _
             "FROM [Conjuntos do Oramento] Left Join Oramento On [Conjuntos do Oramento].[Seqncia do Oramento] = Oramento.[Seqncia do Oramento] " & _
             "Left Join Conjuntos On [Conjuntos do Oramento].[Seqncia Do Conjunto] = Conjuntos.[Seqncia Do Conjunto] WHERE " & Filtro2()
 SQL = SQL & " Group By [Conjuntos do Oramento].[Seqncia Do Conjunto], Conjuntos.Descrio, [Conjuntos do Oramento].[Valor Unitrio] Union "
 SQL = SQL & "Select 3 As Tipo, [Produtos do Oramento].[Seqncia Do Produto], Produtos.Descrio, Sum([Produtos do Oramento].Quantidade) Quantidade, Sum([Produtos do Oramento].[Valor Total] + [Produtos do Oramento].[Valor Do IPI]) Total, [Produtos do Oramento].[Valor Unitrio] " & _
             "FROM [Produtos do Oramento] left Join Oramento On [Produtos do Oramento].[Seqncia do Oramento] = Oramento.[Seqncia do Oramento] " & _
             "Left Join Produtos On [Produtos do Oramento].[Seqncia Do Produto] = Produtos.[Seqncia Do Produto] WHERE " & Filtro3()
 SQL = SQL & " Group By [Produtos do Oramento].[Seqncia Do Produto], Produtos.Descrio, [Produtos do Oramento].[Valor Unitrio])A Order By Tipo Desc"
 
 rsGas.Open SQL, cnGas, adOpenStatic, adLockPessimistic
 
 
 Do While Not rsGas.EOF
    With Tb1
     .AddNew
       ![Seqncia do Oramento] = Sequencia_do_Orcamento
       ![Sequencia do Item] = SeqItem
        If rsGas!Tipo = 1 Or rsGas!Tipo = 3 Then
           ![Seqncia do Produto] = rsGas![Seqncia do Produto]
        End If
        If rsGas!Tipo = 2 Then
           ![Seqncia do Conjunto] = rsGas![Seqncia do Produto]
        End If
       !Quantidade = rsGas!Quantidade
       ![Valor Total] = rsGas!Total
       ![Valor Unitrio] = rsGas![Valor Unitrio]
       !Tp = rsGas!Tipo
     .Update
     .BookMark = .LastModified
    End With
 rsGas.MoveNext
 SeqItem = SeqItem + 1
 Loop
 rsGas.Close
               
End Sub



Private Function Filtro1() As String
 Filtro1 = "Oramento.[Seqncia Do Oramento] = " & Sequencia_do_Orcamento
End Function


Private Function Filtro2() As String
 Filtro2 = "Oramento.[Seqncia Do Oramento] = " & Sequencia_do_Orcamento
End Function


Private Function Filtro3() As String
 Filtro3 = "Oramento.[Seqncia Do Oramento] = " & Sequencia_do_Orcamento
End Function



'Private Function TemGalv(Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Orcamento As Long, _
   Sequencia_do_Produto_Orcamento As Long, Sequencia_do_Produto As Long, Quantidade As Double, _
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double, Oq As String) As Variant
' Dim Tb As New GRecordSet("")
'End Function

Public Sub LigaDesligaBotoes()
   Botao(1).Enabled = Botao(1).Enabled And Permitido("Geral", ACAO_NAVEGANDO)
   Botao(2).Enabled = Botao(2).Enabled And Permitido("Classificao Fiscal", ACAO_NAVEGANDO)
   Botao(3).Enabled = Botao(3).Enabled And Permitido("Municpios", ACAO_NAVEGANDO)
   Botao(4).Enabled = Botao(4).Enabled And Permitido("Geral", ACAO_NAVEGANDO)
   Botao(5).Enabled = Botao(5).Enabled And Permitido("Pases", ACAO_NAVEGANDO)
   Botao(10).Enabled = Botao(10).Enabled And Permitido("Geral", ACAO_NAVEGANDO)
   Botao(13).Enabled = Botao(13).Enabled And Permitido("Oramento", ACAO_NAVEGANDO)
   Botao(14).Enabled = Botao(14).Enabled And Permitido("Oramento", ACAO_NAVEGANDO)
   Botao(16).Enabled = Botao(16).Enabled And Permitido("Oramento", ACAO_NAVEGANDO)
End Sub


Public Property Get txtCampoTab(Index As Integer) As FormataCampos
   Set txtCampoTab = txtCampo(Index)
End Sub


'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possveis erros
   If (vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Cancelado = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb!Cancelado)
         Data_da_Alteracao = IIf(vgSituacao = ACAO_INCLUINDO, Null, vgTb![Data da Alterao])
         Hora_da_Alteracao = IIf(vgSituacao = ACAO_INCLUINDO, Null, vgTb![Hora da Alterao])
         Usuario_da_Alteracao = IIf(vgSituacao = ACAO_INCLUINDO, "", vgTb![Usurio da Alterao])
         Venda_Fechada = chkCampo(6).Value
         Valor_Total_IPI_das_Pecas = txtCampo(39).Value
         Valor_Total_das_Pecas = txtCampo(44).Value
         Sequencia_do_Municipio = txtCampo(17).Value
         Sequencia_do_Pais = txtCampo(60).Value
         Sequencia_do_Orcamento = txtCampo(0).Value
         Sequencia_do_Geral = txtCampo(156).Value
         Observacao = txtCampo(4).Value
         Fechamento = Val(labopcPainel2.Caption)
         Valor_do_Fechamento = txtCampo(32).Value
         Valor_Total_IPI_dos_Produtos = txtCampo(37).Value
         Valor_Total_IPI_dos_Conjuntos = txtCampo(40).Value
         Valor_Total_do_ICMS = txtCampo(9).Value
         Valor_Total_dos_Produtos = txtCampo(41).Value
         Valor_Total_dos_Conjuntos = txtCampo(42).Value
         Valor_Total_de_Produtos_Usados = txtCampo(31).Value
         Valor_Total_Conjuntos_Usados = txtCampo(38).Value
         Valor_Total_dos_Servicos = txtCampo(43).Value
         Valor_Total_do_Orcamento = txtCampo(36).Value
         Nome_Cliente = txtCampo(18).Value
         endereco = txtCampo(24).Value
         CEP = txtCampo(8).Value
         Telefone = txtCampo(22).Value
         Fax = txtCampo(12).Value
         Email = txtCampo(16).Value
         Sequencia_do_Vendedor = txtCampo(33).Value
         Sequencia_do_Pedido = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Seqncia Do Pedido])
         Tipo = Val(labopcPainel1.Caption)
         CPF_e_CNPJ = txtCampo(13).Value
         RG_e_IE = txtCampo(14).Value
         Forma_de_Pagamento = txtCampo(35).Value
         Ocultar_Valor_Unitario = chkCampo(4).Value
         Sequencia_da_Classificacao = txtCampo(34).Value
         Bairro = txtCampo(23).Value
         Caixa_Postal = txtCampo(21).Value
         e_Propriedade = chkCampo(1).Value
         Nome_da_Propriedade = txtCampo(48).Value
         Numero_do_Endereco = txtCampo(20).Value
         Valor_Total_da_Base_de_Calculo = txtCampo(49).Value
         Valor_do_Seguro = txtCampo(47).Value
         Valor_do_Frete = txtCampo(45).Value
         Valor_Total_das_Pecas_Usadas = txtCampo(46).Value
         Sequencia_da_Propriedade = txtCampo(157).Value
         Complemento = txtCampo(25).Value
         Data_de_Emissao = txtCampo(158).Value
         Data_do_Fechamento = txtCampo(50).Value
         Codigo_do_Suframa = txtCampo(51).Value
         Revenda = chkCampo(0).Value
         Valor_Total_do_Tributo = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Valor Total Do Tributo])
         Valor_Total_do_PIS = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Valor Total Do PIS])
         Valor_Total_do_COFINS = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Valor Total Do COFINS])
         Valor_Total_da_Base_ST = txtCampo(52).Value
         Valor_Total_do_ICMS_ST = txtCampo(53).Value
         Aliquota_do_ISS = txtCampo(54).Value
         Reter_ISS = chkCampo(2).Value
         Fatura_Proforma = chkCampo(8).Value
         Entrega_Futura = chkCampo(5).Value
         Sequencia_da_Transportadora = txtCampo(55).Value
         Orcamento_Avulso = chkCampo(7).Value
         Valor_do_Imposto_de_Renda = txtCampo(57).Value
         Local_de_Embarque = txtCampo(58).Value
         UF_de_Embarque = txtCampo(59).Value
         Numero_da_Proforma = txtCampo(159).Value
         Conjunto_Avulso = chkCampo(3).Value
         Descricao_Conjunto_Avulso = txtCampo(61).Value
         Vendedor_Intermediario = txtCampo(62).Value
         Percentual_do_Vendedor = txtCampo(63).Value
         Rebiut = txtCampo(64).Value
         Percentual_Rebiut = txtCampo(65).Value
         Nao_Movimentar_Estoque = chkCampo(9).Value
         Gerou_Encargos = IIf(vgSituacao = ACAO_INCLUINDO, False, vgTb![Gerou Encargos])
         Peso_Bruto = txtCampo(143).Value
         Peso_Liquido = txtCampo(144).Value
         Volumes = txtCampo(145).Value
         Aviso_de_embarque = txtCampo(146).Value
         Hidroturbo = txtCampo(66).Value
         Area_irrigada = txtCampo(67).Value
         Precipitacao_bruta = txtCampo(68).Value
         Horas_irrigada = txtCampo(69).Value
         Area_tot_irrigada_em = txtCampo(72).Value
         Aspersor = txtCampo(73).Value
         Modelo_do_aspersor = txtCampo(74).Value
         Bocal_diametro = txtCampo(75).Value
         Pressao_de_servico = txtCampo(77).Value
         Alcance_do_jato = txtCampo(78).Value
         Espaco_entre_carreadores = txtCampo(81).Value
         Faixa_irrigada = txtCampo(82).Value
         Desnivel_maximo_na_area = txtCampo(88).Value
         Altura_de_succao = txtCampo(89).Value
         Altura_do_aspersor = txtCampo(90).Value
         Tempo_parado_antes_percurso = txtCampo(95).Value
         Com_1 = txtCampo(96).Value
         Com_2 = txtCampo(97).Value
         Com_3 = txtCampo(98).Value
         Modelo_Trecho_A = txtCampo(99).Value
         Modelo_Trecho_B = txtCampo(100).Value
         Modelo_Trecho_C = txtCampo(101).Value
         Qtde_bomba = txtCampo(115).Value
         Marca_bomba = txtCampo(116).Value
         Modelo_bomba = txtCampo(117).Value
         Tamanho_bomba = txtCampo(118).Value
         N_estagios = txtCampo(119).Value
         Diametro_bomba = txtCampo(120).Value
         Pressao_bomba = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Pressao bomba])
         Rendimento_bomba = txtCampo(123).Value
         Rotacao_bomba = txtCampo(124).Value
         Qtde_de_Motor = txtCampo(126).Value
         Marca_do_Motor = txtCampo(127).Value
         Modelo_Motor = txtCampo(128).Value
         Nivel_de_Protecao = txtCampo(129).Value
         Potencia_Nominal = txtCampo(130).Value
         Nro_de_Fases = txtCampo(131).Value
         Voltagem = txtCampo(134).Value
         Modelo_hidroturbo = txtCampo(138).Value
         Eixos = txtCampo(139).Value
         Rodas = txtCampo(140).Value
         Pneus = txtCampo(141).Value
         Tubos = txtCampo(142).Value
         Projetista = txtCampo(137).Value
         Entrega_Tecnica = txtCampo(147).Value
         Sequencia_do_Projeto = IIf(vgSituacao = ACAO_INCLUINDO, 0, vgTb![Sequencia do Projeto])
         Outras_Despesas = txtCampo(148).Value
         Refaturamento = chkCampo(10).Value
         Data_do_Pedido = txtCampo(149).Value
         Data_de_Entrega = txtCampo(150).Value
         Ordem_Interna = chkCampo(11).Value
         Orcamento_Vinculado = txtCampo(161).Value
         frete = txtCampo(151).Value
      Else
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Data_da_Alteracao = vgTb![Data da Alterao]
         Hora_da_Alteracao = vgTb![Hora da Alterao]
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usurio da Alterao]), "", vgTb![Usurio da Alterao])
         Venda_Fechada = IIf(IsNull(vgTb![Venda Fechada]), 0, vgTb![Venda Fechada])
         Valor_Total_IPI_das_Pecas = IIf(IsNull(vgTb![Valor Total IPI das Peas]), 0, vgTb![Valor Total IPI das Peas])
         Valor_Total_das_Pecas = IIf(IsNull(vgTb![Valor Total das Peas]), 0, vgTb![Valor Total das Peas])
         Sequencia_do_Municipio = IIf(IsNull(vgTb![Seqncia Do Municpio]), 0, vgTb![Seqncia Do Municpio])
         Sequencia_do_Pais = IIf(IsNull(vgTb![Seqncia do Pas]), 0, vgTb![Seqncia do Pas])
         Sequencia_do_Orcamento = IIf(IsNull(vgTb![Seqncia do Oramento]), 0, vgTb![Seqncia do Oramento])
         Sequencia_do_Geral = IIf(IsNull(vgTb![Seqncia Do Geral]), 0, vgTb![Seqncia Do Geral])
         Observacao = IIf(IsNull(vgTb!Observao), "", vgTb!Observao)
         Fechamento = IIf(IsNull(vgTb!Fechamento), 0, vgTb!Fechamento)
         Valor_do_Fechamento = IIf(IsNull(vgTb![Valor Do Fechamento]), 0, vgTb![Valor Do Fechamento])
         Valor_Total_IPI_dos_Produtos = IIf(IsNull(vgTb![Valor Total IPI dos Produtos]), 0, vgTb![Valor Total IPI dos Produtos])
         Valor_Total_IPI_dos_Conjuntos = IIf(IsNull(vgTb![Valor Total IPI dos Conjuntos]), 0, vgTb![Valor Total IPI dos Conjuntos])
         Valor_Total_do_ICMS = IIf(IsNull(vgTb![Valor Total Do ICMS]), 0, vgTb![Valor Total Do ICMS])
         Valor_Total_dos_Produtos = IIf(IsNull(vgTb![Valor Total dos Produtos]), 0, vgTb![Valor Total dos Produtos])
         Valor_Total_dos_Conjuntos = IIf(IsNull(vgTb![Valor Total dos Conjuntos]), 0, vgTb![Valor Total dos Conjuntos])
         Valor_Total_de_Produtos_Usados = IIf(IsNull(vgTb![Valor Total de Produtos Usados]), 0, vgTb![Valor Total de Produtos Usados])
         Valor_Total_Conjuntos_Usados = IIf(IsNull(vgTb![Valor Total Conjuntos Usados]), 0, vgTb![Valor Total Conjuntos Usados])
         Valor_Total_dos_Servicos = IIf(IsNull(vgTb![Valor Total dos Servios]), 0, vgTb![Valor Total dos Servios])
         Valor_Total_do_Orcamento = IIf(IsNull(vgTb![Valor Total do Oramento]), 0, vgTb![Valor Total do Oramento])
         Nome_Cliente = IIf(IsNull(vgTb![Nome Cliente]), "", vgTb![Nome Cliente])
         endereco = IIf(IsNull(vgTb!Endereo), "", vgTb!Endereo)
         CEP = IIf(IsNull(vgTb!CEP), "", vgTb!CEP)
         Telefone = IIf(IsNull(vgTb!Telefone), "", vgTb!Telefone)
         Fax = IIf(IsNull(vgTb!Fax), "", vgTb!Fax)
         Email = IIf(IsNull(vgTb!Email), "", vgTb!Email)
         Sequencia_do_Vendedor = IIf(IsNull(vgTb![Seqncia Do Vendedor]), 0, vgTb![Seqncia Do Vendedor])
         Sequencia_do_Pedido = IIf(IsNull(vgTb![Seqncia Do Pedido]), 0, vgTb![Seqncia Do Pedido])
         Tipo = IIf(IsNull(vgTb!Tipo), 0, vgTb!Tipo)
         CPF_e_CNPJ = IIf(IsNull(vgTb![CPF e CNPJ]), "", vgTb![CPF e CNPJ])
         RG_e_IE = IIf(IsNull(vgTb![RG e IE]), "", vgTb![RG e IE])
         Forma_de_Pagamento = IIf(IsNull(vgTb![Forma de Pagamento]), "", vgTb![Forma de Pagamento])
         Ocultar_Valor_Unitario = IIf(IsNull(vgTb![Ocultar Valor Unitrio]), 0, vgTb![Ocultar Valor Unitrio])
         Sequencia_da_Classificacao = IIf(IsNull(vgTb![Seqncia da Classificao]), 0, vgTb![Seqncia da Classificao])
         Bairro = IIf(IsNull(vgTb!Bairro), "", vgTb!Bairro)
         Caixa_Postal = IIf(IsNull(vgTb![Caixa Postal]), "", vgTb![Caixa Postal])
         e_Propriedade = IIf(IsNull(vgTb![ Propriedade]), 0, vgTb![ Propriedade])
         Nome_da_Propriedade = IIf(IsNull(vgTb![Nome da Propriedade]), "", vgTb![Nome da Propriedade])
         Numero_do_Endereco = IIf(IsNull(vgTb![Nmero Do Endereo]), "", vgTb![Nmero Do Endereo])
         Valor_Total_da_Base_de_Calculo = IIf(IsNull(vgTb![Valor Total da Base de Clculo]), 0, vgTb![Valor Total da Base de Clculo])
         Valor_do_Seguro = IIf(IsNull(vgTb![Valor Do Seguro]), 0, vgTb![Valor Do Seguro])
         Valor_do_Frete = IIf(IsNull(vgTb![Valor Do Frete]), 0, vgTb![Valor Do Frete])
         Valor_Total_das_Pecas_Usadas = IIf(IsNull(vgTb![Valor Total das Peas Usadas]), 0, vgTb![Valor Total das Peas Usadas])
         Sequencia_da_Propriedade = IIf(IsNull(vgTb![Seqncia da Propriedade]), 0, vgTb![Seqncia da Propriedade])
         Complemento = IIf(IsNull(vgTb!Complemento), "", vgTb!Complemento)
         Data_de_Emissao = vgTb![Data de Emisso]
         Data_do_Fechamento = vgTb![Data Do Fechamento]
         Codigo_do_Suframa = IIf(IsNull(vgTb![Cdigo Do Suframa]), "", vgTb![Cdigo Do Suframa])
         Revenda = IIf(IsNull(vgTb!Revenda), 0, vgTb!Revenda)
         Valor_Total_do_Tributo = IIf(IsNull(vgTb![Valor Total Do Tributo]), 0, vgTb![Valor Total Do Tributo])
         Valor_Total_do_PIS = IIf(IsNull(vgTb![Valor Total Do PIS]), 0, vgTb![Valor Total Do PIS])
         Valor_Total_do_COFINS = IIf(IsNull(vgTb![Valor Total Do COFINS]), 0, vgTb![Valor Total Do COFINS])
         Valor_Total_da_Base_ST = IIf(IsNull(vgTb![Valor Total da Base ST]), 0, vgTb![Valor Total da Base ST])
         Valor_Total_do_ICMS_ST = IIf(IsNull(vgTb![Valor Total Do ICMS ST]), 0, vgTb![Valor Total Do ICMS ST])
         Aliquota_do_ISS = IIf(IsNull(vgTb![Alquota Do ISS]), 0, vgTb![Alquota Do ISS])
         Reter_ISS = IIf(IsNull(vgTb![Reter ISS]), 0, vgTb![Reter ISS])
         Fatura_Proforma = IIf(IsNull(vgTb![Fatura Proforma]), 0, vgTb![Fatura Proforma])
         Entrega_Futura = IIf(IsNull(vgTb![Entrega Futura]), 0, vgTb![Entrega Futura])
         Sequencia_da_Transportadora = IIf(IsNull(vgTb![Seqncia da Transportadora]), 0, vgTb![Seqncia da Transportadora])
         Orcamento_Avulso = IIf(IsNull(vgTb![Oramento Avulso]), 0, vgTb![Oramento Avulso])
         Valor_do_Imposto_de_Renda = IIf(IsNull(vgTb![Valor Do Imposto de Renda]), 0, vgTb![Valor Do Imposto de Renda])
         Local_de_Embarque = IIf(IsNull(vgTb![Local de Embarque]), "", vgTb![Local de Embarque])
         UF_de_Embarque = IIf(IsNull(vgTb![UF de Embarque]), "", vgTb![UF de Embarque])
         Numero_da_Proforma = IIf(IsNull(vgTb![Nmero da Proforma]), 0, vgTb![Nmero da Proforma])
         Conjunto_Avulso = IIf(IsNull(vgTb![Conjunto Avulso]), 0, vgTb![Conjunto Avulso])
         Descricao_Conjunto_Avulso = IIf(IsNull(vgTb![Descrio Conjunto Avulso]), "", vgTb![Descrio Conjunto Avulso])
         Vendedor_Intermediario = IIf(IsNull(vgTb![Vendedor Intermediario]), "", vgTb![Vendedor Intermediario])
         Percentual_do_Vendedor = IIf(IsNull(vgTb![Percentual Do Vendedor]), 0, vgTb![Percentual Do Vendedor])
         Rebiut = IIf(IsNull(vgTb!Rebiut), "", vgTb!Rebiut)
         Percentual_Rebiut = IIf(IsNull(vgTb![Percentual Rebiut]), 0, vgTb![Percentual Rebiut])
         Nao_Movimentar_Estoque = IIf(IsNull(vgTb![Nao Movimentar Estoque]), 0, vgTb![Nao Movimentar Estoque])
         Gerou_Encargos = IIf(IsNull(vgTb![Gerou Encargos]), 0, vgTb![Gerou Encargos])
         Peso_Bruto = IIf(IsNull(vgTb![Peso Bruto]), 0, vgTb![Peso Bruto])
         Peso_Liquido = IIf(IsNull(vgTb![Peso Lquido]), 0, vgTb![Peso Lquido])
         Volumes = IIf(IsNull(vgTb!Volumes), 0, vgTb!Volumes)
         Aviso_de_embarque = IIf(IsNull(vgTb![Aviso de embarque]), "", vgTb![Aviso de embarque])
         Hidroturbo = IIf(IsNull(vgTb!Hidroturbo), "", vgTb!Hidroturbo)
         Area_irrigada = IIf(IsNull(vgTb![Area irrigada]), 0, vgTb![Area irrigada])
         Precipitacao_bruta = IIf(IsNull(vgTb![Precipitao bruta]), 0, vgTb![Precipitao bruta])
         Horas_irrigada = IIf(IsNull(vgTb![Horas irrigada]), 0, vgTb![Horas irrigada])
         Area_tot_irrigada_em = IIf(IsNull(vgTb![Area tot irrigada em]), 0, vgTb![Area tot irrigada em])
         Aspersor = IIf(IsNull(vgTb!Aspersor), "", vgTb!Aspersor)
         Modelo_do_aspersor = IIf(IsNull(vgTb![Modelo do aspersor]), "", vgTb![Modelo do aspersor])
         Bocal_diametro = IIf(IsNull(vgTb![Bocal Diametro]), 0, vgTb![Bocal Diametro])
         Pressao_de_servico = IIf(IsNull(vgTb![Presso de servio]), 0, vgTb![Presso de servio])
         Alcance_do_jato = IIf(IsNull(vgTb![Alcance do jato]), 0, vgTb![Alcance do jato])
         Espaco_entre_carreadores = IIf(IsNull(vgTb![Espao entre carreadores]), 0, vgTb![Espao entre carreadores])
         Faixa_irrigada = IIf(IsNull(vgTb![Faixa irrigada]), 0, vgTb![Faixa irrigada])
         Desnivel_maximo_na_area = IIf(IsNull(vgTb![Desnivel maximo na area]), 0, vgTb![Desnivel maximo na area])
         Altura_de_succao = IIf(IsNull(vgTb![Altura de suco]), 0, vgTb![Altura de suco])
         Altura_do_aspersor = IIf(IsNull(vgTb![Altura do aspersor]), 0, vgTb![Altura do aspersor])
         Tempo_parado_antes_percurso = IIf(IsNull(vgTb![Tempo parado antes percurso]), 0, vgTb![Tempo parado antes percurso])
         Com_1 = IIf(IsNull(vgTb![Com 1]), 0, vgTb![Com 1])
         Com_2 = IIf(IsNull(vgTb![Com 2]), 0, vgTb![Com 2])
         Com_3 = IIf(IsNull(vgTb![Com 3]), 0, vgTb![Com 3])
         Modelo_Trecho_A = IIf(IsNull(vgTb![Modelo Trecho A]), 0, vgTb![Modelo Trecho A])
         Modelo_Trecho_B = IIf(IsNull(vgTb![Modelo Trecho B]), 0, vgTb![Modelo Trecho B])
         Modelo_Trecho_C = IIf(IsNull(vgTb![Modelo Trecho C]), 0, vgTb![Modelo Trecho C])
         Qtde_bomba = IIf(IsNull(vgTb![Qtde bomba]), 0, vgTb![Qtde bomba])
         Marca_bomba = IIf(IsNull(vgTb![Marca bomba]), "", vgTb![Marca bomba])
         Modelo_bomba = IIf(IsNull(vgTb![Modelo bomba]), "", vgTb![Modelo bomba])
         Tamanho_bomba = IIf(IsNull(vgTb![Tamanho bomba]), "", vgTb![Tamanho bomba])
         N_estagios = IIf(IsNull(vgTb![N estagios]), 0, vgTb![N estagios])
         Diametro_bomba = IIf(IsNull(vgTb![Diametro bomba]), 0, vgTb![Diametro bomba])
         Pressao_bomba = IIf(IsNull(vgTb![Pressao bomba]), 0, vgTb![Pressao bomba])
         Rendimento_bomba = IIf(IsNull(vgTb![Rendimento bomba]), 0, vgTb![Rendimento bomba])
         Rotacao_bomba = IIf(IsNull(vgTb![Rotao bomba]), 0, vgTb![Rotao bomba])
         Qtde_de_Motor = IIf(IsNull(vgTb![Qtde de Motor]), 0, vgTb![Qtde de Motor])
         Marca_do_Motor = IIf(IsNull(vgTb![Marca Do Motor]), "", vgTb![Marca Do Motor])
         Modelo_Motor = IIf(IsNull(vgTb![Modelo Motor]), "", vgTb![Modelo Motor])
         Nivel_de_Protecao = IIf(IsNull(vgTb![Nivel de Proteo]), "", vgTb![Nivel de Proteo])
         Potencia_Nominal = IIf(IsNull(vgTb![Potencia Nominal]), 0, vgTb![Potencia Nominal])
         Nro_de_Fases = IIf(IsNull(vgTb![Nro de Fases]), 0, vgTb![Nro de Fases])
         Voltagem = IIf(IsNull(vgTb!Voltagem), 0, vgTb!Voltagem)
         Modelo_hidroturbo = IIf(IsNull(vgTb![Modelo hidroturbo]), "", vgTb![Modelo hidroturbo])
         Eixos = IIf(IsNull(vgTb!Eixos), 0, vgTb!Eixos)
         Rodas = IIf(IsNull(vgTb!Rodas), 0, vgTb!Rodas)
         Pneus = IIf(IsNull(vgTb!Pneus), "", vgTb!Pneus)
         Tubos = IIf(IsNull(vgTb!Tubos), "", vgTb!Tubos)
         Projetista = IIf(IsNull(vgTb!Projetista), 0, vgTb!Projetista)
         Entrega_Tecnica = IIf(IsNull(vgTb![Entrega Tecnica]), "", vgTb![Entrega Tecnica])
         Sequencia_do_Projeto = IIf(IsNull(vgTb![Sequencia do Projeto]), 0, vgTb![Sequencia do Projeto])
         Outras_Despesas = IIf(IsNull(vgTb![Outras Despesas]), 0, vgTb![Outras Despesas])
         Refaturamento = IIf(IsNull(vgTb!Refaturamento), 0, vgTb!Refaturamento)
         Data_do_Pedido = vgTb![Data do Pedido]
         Data_de_Entrega = vgTb![Data de Entrega]
         Ordem_Interna = IIf(IsNull(vgTb![Ordem Interna]), 0, vgTb![Ordem Interna])
         Orcamento_Vinculado = IIf(IsNull(vgTb![Oramento Vinculado]), 0, vgTb![Oramento Vinculado])
         frete = IIf(IsNull(vgTb!frete), "", vgTb!frete)
      End If
   End If
   If Err Then Err.Clear                          'se deu algum erro, vamos reset-lo
End Sub


'verifica permisses para as condies especiais
'e muda situao de alguns botes
Private Sub AnalisaCondicoes()
   Dim i As Integer
   On Error Resume Next
   If Not mdiIRRIG.ActiveForm Is Nothing Then
      If mdiIRRIG.ActiveForm.Name <> Me.Name And Me.Visible Then Exit Sub
   End If
   With mdiIRRIG
      i = Permitido(vgIdentTab, ACAO_INCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemInclusao = i
      grdBrowse.AllowInsert = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = (vgPWUsuario = "YGOR" Or (vgPWUsuario = "YGOR BARBOSA" And Ordem_Interna = True)) And Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = ((Sequencia_do_Pedido = 0) And Cancelado = 0 Or vgPWUsuario = "YGOR") And Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i And vgAlterar
      .Menu_Paltera.Enabled = i
      LigaDesligaControles Me, Not i
   End With
End Sub


'executa processos/validacoes nos campos do arquivo
Public Function Executar(vgOq As String, Optional ByRef vgColumn As Integer) As String
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer, vgInd As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validao esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque no fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = ((IIf(vgPWUsuario <> "YGOR", Data_de_Emissao = Date, Not Vazio(Data_de_Emissao))) And (IsDate(Data_de_Emissao) Or Vazio(Data_de_Emissao)))
      vgMsg$ = "Data de Emisso tem que ser a data de hoje!"
      If Not vgOk Then vgColumn = 159
      If vgOk Then
         If Vazio(Nome_Cliente) And vgSituacao = ACAO_INCLUINDO Then
            vgOk = (Sequencia_do_Geral > 0)
            vgMsg$ = "Seqncia do Geral invlido!"
            If Not vgOk Then vgColumn = 157
         End If
      End If
      If vgOk Then
         If Sequencia_do_Geral = 0 Then
            vgOk = (Not Vazio(Nome_Cliente))
            vgMsg$ = "Nome Cliente no pode ser vazio!"
            If Not vgOk Then vgColumn = 19
         End If
      End If
      If vgOk Then
         If (Sequencia_do_Geral = 0) And e_Propriedade = True Then
            vgOk = (Not Vazio(Nome_da_Propriedade))
            vgMsg$ = "Nome da Propriedade no pode ser vazio!"
            If Not vgOk Then vgColumn = 49
         End If
      End If
      If vgOk Then
         If Sequencia_do_Geral = 0 And Fatura_Proforma = False Then
            vgOk = (Not Vazio(Telefone))
            vgMsg$ = "Telefone no pode ser vazio!"
            If Not vgOk Then vgColumn = 23
         End If
      End If
      If vgOk Then
         If Sequencia_do_Geral = 0 Then
            vgOk = (Sequencia_do_Municipio > 0)
            vgMsg$ = "Seqncia do Municpio invlido!"
            If Not vgOk Then vgColumn = 18
         End If
      End If
      If vgOk Then
         If Fatura_Proforma = False Then
            If Tipo2 = 0 Then
               vgOk = (VDV2(CPF_e_CNPJ))
               vgMsg$ = "CPF no pode ser vazio ou CPF Incorreto!"
               If Not vgOk Then vgColumn = 14
            End If
         End If
      End If
      If vgOk Then
         If Fatura_Proforma = False Then
            If Tipo2 = 1 Then
               vgOk = (VCGC(CPF_e_CNPJ))
               vgMsg$ = "CNPJ no pode ser vazio ou CNPJ Incorreto!"
               If Not vgOk Then vgColumn = 14
            End If
         End If
      End If
      If vgOk Then
         If Fatura_Proforma = False Then
            If Tipo2 = 1 Then
               vgOk = ((ValidaIE(MunicipioAux!UF, RG_e_IE, False)) Or Vazio(RG_e_IE))
               vgMsg$ = "" & MsgValIE & ""
               If Not vgOk Then vgColumn = 15
            End If
         End If
      End If
      If vgOk Then
         If Sequencia_do_Geral = 0 Or GeralAux![Seqncia Do Vendedor] = 0 Then
            vgOk = (Sequencia_do_Vendedor > 0)
            vgMsg$ = "Seqncia do Vendedor invlido!"
            If Not vgOk Then vgColumn = 34
         End If
      End If
      If vgOk Then
         If Fatura_Proforma Then
            vgOk = (Sequencia_do_Pais > 0)
            vgMsg$ = "Seqncia do Pas invlido!"
            If Not vgOk Then vgColumn = 61
         End If
      End If
      If vgOk Then
         vgOk = (Valor_do_Fechamento <= 0 And ValidaDesconto())
         vgMsg$ = "Desconto Invalido!"
         If Not vgOk Then vgColumn = 33
      End If
      If vgOk Then
         vgOk = (Outras_Despesas >= 0)
         vgMsg$ = "Outras Despesas invlido!"
         If Not vgOk Then vgColumn = 149
      End If
      If vgOk Then
         If Not Vazio(Vendedor_Intermediario) Then
            vgOk = (IIf((MunicipioAux!UF = "EX" And Vendedor_Intermediario = "Sistemas de Irrigaes G & E Ltda"), (Percentual_do_Vendedor > 0) And Percentual_do_Vendedor <= 20, (Percentual_do_Vendedor > 0) And Percentual_do_Vendedor <= 10))
            vgMsg$ = "Percentual do Vendedor Invalido!"
            If Not vgOk Then vgColumn = 64
         End If
      End If
      If vgOk Then
         vgOk = (Projetista >= 0)
         vgMsg$ = "Projetista invlido!"
         If Not vgOk Then vgColumn = 138
      End If
      If vgOk Then
         If Not Vazio(Rebiut) Then
            vgOk = ((Percentual_Rebiut > 0) And Percentual_Rebiut <= 1)
            vgMsg$ = "Percentual do Vendedor Rebiut no pode ser 0 nem maior que 1%"
            If Not vgOk Then vgColumn = 66
         End If
      End If
      If vgOk Then
         vgOk = (Com_1 >= 0)
         vgMsg$ = "Com 1 invlido!"
         If Not vgOk Then vgColumn = 97
      End If
      If vgOk Then
         vgOk = (Com_2 >= 0)
         vgMsg$ = "Com 2 invlido!"
         If Not vgOk Then vgColumn = 98
      End If
      If vgOk Then
         vgOk = (Com_3 >= 0)
         vgMsg$ = "Com 3 invlido!"
         If Not vgOk Then vgColumn = 99
      End If
      If vgOk Then
         vgOk = (Modelo_Trecho_A >= 0)
         vgMsg$ = "Modelo trecho a invlido!"
         If Not vgOk Then vgColumn = 100
      End If
      If vgOk Then
         vgOk = (Modelo_Trecho_B >= 0)
         vgMsg$ = "Modelo trecho b invlido!"
         If Not vgOk Then vgColumn = 101
      End If
      If vgOk Then
         vgOk = (Modelo_Trecho_C >= 0)
         vgMsg$ = "Modelo trecho c invlido!"
         If Not vgOk Then vgColumn = 102
      End If
      If vgOk Then
         vgOk = (Qtde_bomba >= 0)
         vgMsg$ = "Qtde bomba invlido!"
         If Not vgOk Then vgColumn = 116
      End If
      If vgOk Then
         vgOk = (N_estagios >= 0)
         vgMsg$ = "N estagios invlido!"
         If Not vgOk Then vgColumn = 120
      End If
      If vgOk Then
         vgOk = (Diametro_bomba >= 0)
         vgMsg$ = "Diametro bomba invlido!"
         If Not vgOk Then vgColumn = 121
      End If
      If vgOk Then
         vgOk = (Rendimento_bomba >= 0)
         vgMsg$ = "Rendimento bomba invlido!"
         If Not vgOk Then vgColumn = 124
      End If
      If vgOk Then
         vgOk = (Rotacao_bomba >= 0)
         vgMsg$ = "Rotao bomba invlido!"
         If Not vgOk Then vgColumn = 125
      End If
      If vgOk Then
         vgOk = (Qtde_de_Motor >= 0)
         vgMsg$ = "Qtde de motor invlido!"
         If Not vgOk Then vgColumn = 127
      End If
      If vgOk Then
         vgOk = (Potencia_Nominal >= 0)
         vgMsg$ = "Potencia nominal invlido!"
         If Not vgOk Then vgColumn = 131
      End If
      If vgOk Then
         vgOk = ((Nro_de_Fases <= 3) And Nro_de_Fases >= 0)
         vgMsg$ = "Nro de fases invlido!"
         If Not vgOk Then vgColumn = 132
      End If
      If vgOk Then
         vgOk = (Voltagem >= 0)
         vgMsg$ = "Voltagem invlido!"
         If Not vgOk Then vgColumn = 135
      End If
      If vgOk Then
         vgOk = (Eixos >= 0)
         vgMsg$ = "Eixos invlido!"
         If Not vgOk Then vgColumn = 140
      End If
      If vgOk Then
         vgOk = (Rodas >= 0)
         vgMsg$ = "Rodas invlido!"
         If Not vgOk Then vgColumn = 141
      End If
      If vgOk Then
         vgOk = (Peso_Bruto >= 0)
         vgMsg$ = "Peso bruto invlido!"
         If Not vgOk Then vgColumn = 144
      End If
      If vgOk Then
         vgOk = (Peso_Liquido >= 0)
         vgMsg$ = "Peso lquido invlido!"
         If Not vgOk Then vgColumn = 145
      End If
      If vgOk Then
         vgOk = (Volumes >= 0)
         vgMsg$ = "Volumes invlido!"
         If Not vgOk Then vgColumn = 146
      End If
      If vgOk Then
         If Venda_Fechada Or Ordem_Interna Then
            vgOk = ((Not Vazio(Data_do_Pedido)) And (IsDate(Data_do_Pedido) Or Vazio(Data_do_Pedido)))
            vgMsg$ = "Data do Pedido no pode ser vazio!"
            If Not vgOk Then vgColumn = 150
         End If
      End If
      If vgOk Then
         If Venda_Fechada And Ordem_Interna = False Then
            vgOk = ((Not Vazio(Data_de_Entrega)) And (IsDate(Data_de_Entrega) Or Vazio(Data_de_Entrega)))
            vgMsg$ = "Data de Entrega no pode ser vazio!"
            If Not vgOk Then vgColumn = 151
         End If
      End If
      If vgOk Then
         vgOk = (Orcamento_Vinculado >= 0)
         vgMsg$ = "Orcamento Vinculado invlido!"
         If Not vgOk Then vgColumn = 162
      End If
      If vgOk Then
         If Ordem_Interna = False Then
            vgOk = (Not Vazio(frete))
            vgMsg$ = "Frete no pode ser vazio!"
            If Not vgOk Then vgColumn = 152
         End If
      End If
      If vgOk Then
         vgMsg$ = ""
      ElseIf vgColumn <> -1 And Not vgEmBrowse Then
         txtCampo(vgColumn - 1).SetFocus
      End If
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      If vgPriVez = False Then
         vgPriVez = True
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).Text = ""
            End If
         Next
         InicializaApelidos COM_TEXTBOX
         On Error Resume Next
         chkCampo(0).Value = False
         txtCampo(34).Value = PegaNCMPadrao
         chkCampo(1).Value = False
         opcPainel1(0).Value = True
         opcPainel2(1).Value = True
         txtCampo(54).Value = 3
         chkCampo(2).Value = False
         chkCampo(3).Value = False
         txtCampo(73).Value = "SETORIAL"
         txtCampo(77).Value = 100
                  txtCampo(149).Value = Date
         chkCampo(4).Value = True
         chkCampo(5).Value = False
         chkCampo(6).Value = False
         chkCampo(7).Value = IIf(Fatura_Proforma, True, False)
         chkCampo(8).Locked = False
         chkCampo(8).Value = IIf(Me.Caption = "Fatura Proforma", 1, 0)
         chkCampo(8).Locked = True
         txtCampo(159).Value = IIf(Fatura_Proforma, SuperPegaSequencial("Oramento", "Nmero da Proforma", "[Fatura Proforma] = 1"), 0)
         chkCampo(9).Value = False
         chkCampo(10).Value = False
         chkCampo(11).Value = IIf(Me.Caption = "Ordem de Produo Interna", 1, 0)
                  On Error GoTo DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 17
         PoeRelEFiltroCbo 33
         PoeRelEFiltroCbo 34
         PoeRelEFiltroCbo 55
         PoeRelEFiltroCbo 60
         PoeRelEFiltroCbo 62
         PoeRelEFiltroCbo 64
         PoeRelEFiltroCbo 99
         PoeRelEFiltroCbo 100
         PoeRelEFiltroCbo 101
         PoeRelEFiltroCbo 137
         PoeRelEFiltroCbo 156
         PoeRelEFiltroCbo 157
         PoeRelEFiltroCbo 161
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 17
         PoeRelEFiltroCbo 33
         PoeRelEFiltroCbo 34
         PoeRelEFiltroCbo 55
         PoeRelEFiltroCbo 60
         PoeRelEFiltroCbo 62
         PoeRelEFiltroCbo 64
         PoeRelEFiltroCbo 99
         PoeRelEFiltroCbo 100
         PoeRelEFiltroCbo 101
         PoeRelEFiltroCbo 137
         PoeRelEFiltroCbo 156
         PoeRelEFiltroCbo 157
         PoeRelEFiltroCbo 161
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Revenda
         chkCampo(1).Value = e_Propriedade
         opcPainel1(Tipo).Value = True
         opcPainel2(Fechamento).Value = True
         chkCampo(2).Value = Reter_ISS
         chkCampo(3).Value = Conjunto_Avulso
         chkCampo(4).Value = Ocultar_Valor_Unitario
         chkCampo(5).Value = Entrega_Futura
         chkCampo(6).Value = Venda_Fechada
         chkCampo(7).Value = Orcamento_Avulso
         chkCampo(8).Value = Fatura_Proforma
         chkCampo(9).Value = Nao_Movimentar_Estoque
         chkCampo(10).Value = Refaturamento
         chkCampo(11).Value = Ordem_Interna
         If vgSituacao = ACAO_NAVEGANDO Then
            If Me.Name = mdiIRRIG.ActiveForm.Name Then
               If Not ActiveControl Is Nothing Then
                  If TypeOf ActiveControl Is GListV Then
                     If Not ActiveControl.PreEditing Then DoEvents
                  Else
                     DoEvents
                  End If
               End If
            End If
         End If
      Else
         Executar INICIALIZACOES
      End If
      vgPriVez = False
   ElseIf vgOq = TESTA_VAL_RS Then
      vgTb.Resync 1         'adAffectCurrent
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            If vgTb.Fields(txtCampo(i).DataField).Value <> txtCampo(i).OriginalValue Then
               If Len(vgMsg$) = 0 Then
                  vgMsg$ = Caption + "|" + CStr(3600 + Abs(vgEmBrowse)) + "|" + LoadGasString(122)
               End If
               If vgEmBrowse Then
                  Exit For
               Else
                  vgPriVez = True
                  txtCampo(i).SetOriginalValue = True
                  txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
                  vgPriVez = False
               End If
            End If
         End If
      Next
   ElseIf vgOq = POE_NO_ARQUIVO Then
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            If Not vgTb.Table.Columns(txtCampo(i).DataField).SeqInterno Then
               If (txtCampo(i).Value & "" <> vgTb.Fields(txtCampo(i).DataField).Value & "") Or _
                        (IsNull(txtCampo(i).Value) Xor IsNull(vgTb.Fields(txtCampo(i).DataField).Value)) Then    'se for diferente do contedo atual do RS
                  vgTb.Fields(txtCampo(i).DataField).Value = txtCampo(i).Value
               End If
            End If
         End If
      Next
      Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
      Data_da_Alteracao = vgTb![Data da Alterao]
      Hora_da_Alteracao = vgTb![Hora da Alterao]
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usurio da Alterao]), "", vgTb![Usurio da Alterao])
      InicializaApelidos COM_TEXTBOX
      vgTb![Venda Fechada] = Venda_Fechada
      vgTb!Fechamento = Fechamento
      Sequencia_do_Pedido = IIf(IsNull(vgTb![Seqncia Do Pedido]), 0, vgTb![Seqncia Do Pedido])
      vgTb!Tipo = Tipo
      vgTb![Ocultar Valor Unitrio] = Ocultar_Valor_Unitario
      vgTb![ Propriedade] = e_Propriedade
      vgTb!Revenda = Revenda
      Valor_Total_do_Tributo = IIf(IsNull(vgTb![Valor Total Do Tributo]), 0, vgTb![Valor Total Do Tributo])
      Valor_Total_do_PIS = IIf(IsNull(vgTb![Valor Total Do PIS]), 0, vgTb![Valor Total Do PIS])
      Valor_Total_do_COFINS = IIf(IsNull(vgTb![Valor Total Do COFINS]), 0, vgTb![Valor Total Do COFINS])
      vgTb![Reter ISS] = Reter_ISS
      vgTb![Fatura Proforma] = Fatura_Proforma
      vgTb![Entrega Futura] = Entrega_Futura
      vgTb![Oramento Avulso] = Orcamento_Avulso
      vgTb![Conjunto Avulso] = Conjunto_Avulso
      vgTb![Nao Movimentar Estoque] = Nao_Movimentar_Estoque
      Gerou_Encargos = IIf(IsNull(vgTb![Gerou Encargos]), 0, vgTb![Gerou Encargos])
      Pressao_bomba = IIf(IsNull(vgTb![Pressao bomba]), 0, vgTb![Pressao bomba])
      Sequencia_do_Projeto = IIf(IsNull(vgTb![Sequencia do Projeto]), 0, vgTb![Sequencia do Projeto])
      vgTb!Refaturamento = Refaturamento
      vgTb![Ordem Interna] = Ordem_Interna
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
   ElseIf vgOq = PODE_ALTERAR Then
      vgOk = (vgSituacao = ACAO_INCLUINDO Or vgAlterar)
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            txtCampo(i).Locked = Not (vgOk And txtCampo(i).Editable)
         End If
      Next
      For i = 0 To UBound(chkCampo)
         If Len(chkCampo(i).DataField) > 0 Then
            chkCampo(i).Locked = Not (vgOk And chkCampo(i).Editable)
         End If
      Next
      For i = 0 To UBound(opcPainel1)
         If Len(opcPainel1(i).DataField) > 0 Then
            If Not opcPainel1(i).Value Then    'vamos primeiro desabilitar os no selecionados
               opcPainel1(i).Locked = Not (vgOk And opcPainel1(i).Editable)
            Else
               vgInd = i
            End If
         End If
      Next
      opcPainel1(vgInd).Locked = False
      opcPainel1(vgInd).Value = True
      opcPainel1(vgInd).Locked = Not (vgOk And opcPainel1(vgInd).Editable)
      For i = 0 To UBound(opcPainel2)
         If Len(opcPainel2(i).DataField) > 0 Then
            If Not opcPainel2(i).Value Then    'vamos primeiro desabilitar os no selecionados
               opcPainel2(i).Locked = Not (vgOk And opcPainel2(i).Editable)
            Else
               vgInd = i
            End If
         End If
      Next
      opcPainel2(vgInd).Locked = False
      opcPainel2(vgInd).Value = True
      opcPainel2(vgInd).Locked = Not (vgOk And opcPainel2(vgInd).Editable)
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_INCLUINDO Then
         AjustaValores
      ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
         AjustaValores
      End If
   ElseIf vgOq = PROCESSOS_DIRETOS Then
      InicializaApelidos COM_REGISTRO
      vgTb.Edit
      Set vgRsError = vgTb
      If GeralAux![Seqncia Do Vendedor] > 0 Then
         vgTb![Seqncia Do Vendedor] = (GeralAux![Seqncia Do Vendedor])
         Sequencia_do_Vendedor = vgTb![Seqncia Do Vendedor]
      End If
      vgTb.Update
      Set vgRsError = Nothing
   ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
   End If
   Executar = vgMsg$                              'prepara saida da funo
   vgPriVez = vgPV
   Exit Function                                  'e cai fora...

DeuErro:
   Select Case Err                                'vamos verificar se deu algum erro

      Case -2147467259
         Resume Next

      Case -2147217885                            'registro foi apagado
         vgPriVez = False
         MoveRegistro Me, REG_FORCAVOLTA          'volta um registro
         PrepBotoes Me, vgSituacao                'acerta icones dos botoes

   End Select
   Executar = Err.Source + "|" + Trim$(Str$(Err)) + "|" + Error$ 'no teve jeito o erro no pode ser evitado...
   If Err = 3265 Then Executar = Executar & vbCrLf & vbCrLf & txtCampo(i).DataField
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      Set vgRsError = Nothing
   End If
   vgPriVez = vgPV
End Function


Private Sub grdBrowse_DeleteData(ByVal vgItem As Long, vgColumns() As Variant, vgDataDeleted As Boolean, vgErrorMessage As String)
   vgDataDeleted = mdiIRRIG.ExcluiRegistro()
End Sub

   
Private Sub grdBrowse_InitEdit(CancelEdit As Boolean)
   Reposition
End Sub


Private Sub grdBrowse_ItemSelect(ByVal vgItem As Long, vgColumns() As Variant)
   If vgPriVez Or Not grdBrowse.Visible Then Exit Sub
   If vgSituacao = ACAO_NAVEGANDO Then Executar PEGA_DO_ARQUIVO
End Sub


'evento disparado ao mudar de registro no grid.
Private Sub grdBrowse_SkipRecord(Columns() As Variant, ByVal BookMark As Variant)
   If vgSituacao = ACAO_NAVEGANDO Then Reposition
End Sub


Private Sub grdBrowse_GetColumnFilter(ByVal vgColumn As Integer, vgColumns() As Variant, vgFilter As String)
   If UBound(txtCampo) >= vgColumn - 1 Then
      vgFilter = txtCampo(vgColumn - 1).Filter
   End If
End Sub


   
'executa a pr-validao da coluna do grid do modo grade do formulrio
Private Sub grdBrowse_GetColumnLocked(ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   ExecutaPreValidacao                            'checa as pr-validaes
   vgLocked = Not FormField.Enabled               'aplica as definies de pr-validao que so aplicadas ao campo da tela
End Sub



Private Sub grdBrowse_SaveData(ByVal vgItem As Long, vgColumns() As Variant, vgDataSaved As Boolean, vgColumn As Integer, vgErrorMessage As String)
   mdiIRRIG.SalvaDados vgColumn
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
End Sub

   
Private Sub grdBrowse_StatusChanged(ByVal vgNewStatus As Integer)
   If (vgNewStatus = ACAO_EXCLUINDO And Val(grdBrowse.RecordSet.BookMark) >= 0) Then
      Reposition
   End If
   PrepBotoes Me, vgNewStatus                          'acerta icones dos botoes
   mdiIRRIG.RemontaForm                                'remonta dos os form da tela
End Sub


'apresenta popup menu para trabalhar com o grid
Private Sub grdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal vgCurCol As Integer)
   MostraPopGrid Me, Button
End Sub


'liga/desliga flag de repetio do ltimo reg visualizado
Public Sub LigaDesligaAlterar()
   vgAlterar = Not vgAlterar
   vgUltAlterar = vgAlterar                            'guarda situao de "pode alterar"
   AnalisaCondicoes                                    'vamos atualizar as condies para incluso, excluso, alterao...
   MostraFormulas
   ExecutaVisivel
   PrepBotoes Me, vgSituacao                           'acerta icones dos botoes
End Sub


'evento - quando qq tecla for digitada no formulrio
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim Ok As Boolean
   If Not Me.ActiveControl Is Nothing Then
      Ok = (Not TypeOf Me.ActiveControl Is GListV)         'se no est em um GRID
   Else
      Ok = True
   End If
   If Not Ok Then
      Ok = (Me.ActiveControl.Status = ACAO_NAVEGANDO And Not Me.ActiveControl.PreEditing) 'e se grid no est em pr-edio, edio nem incluso
   End If
   If KeyAscii = vbKeyEscape And Ok Then                                                  'se teclou ESC
      Unload Me                                   'tira este form da memria
   End If
End Sub


'evento - quando o formurio for pintado
Private Sub Form_Paint()
   grdBrowse.Visible = vgEmBrowse                 'AH VB!!...
End Sub


Public Sub CancelGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).CancelEdit
      End If
   Next
End Sub


Public Sub SaveGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).SaveEdit
      End If
   Next
End Sub


'prepara botes e o formulrio para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   Dim i As Integer, x As String, MudouFiltro As Boolean, vgCols() As Variant
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   Set Orcamento = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes
   End If
   On Error Resume Next
   For i = 0 To 4
      Select Case i
         Case 0
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(0).CloseRecordset
            Else
               x$ = ExecutaGrid(0, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(0).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(0).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(0).CloseRecordset
                  End If
                  Grid(0).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(0, vgCols(), CONDICOES_ESPECIAIS)
         Case 1
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(1).CloseRecordset
            Else
               x$ = ExecutaGrid(1, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(1).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(1).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(1).CloseRecordset
                  End If
                  Grid(1).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(1, vgCols(), CONDICOES_ESPECIAIS)
         Case 2
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(2).CloseRecordset
            Else
               x$ = ExecutaGrid(2, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(2).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(2).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(2).CloseRecordset
                  End If
                  Grid(2).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(2, vgCols(), CONDICOES_ESPECIAIS)
         Case 3
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(3).CloseRecordset
            Else
               x$ = ExecutaGrid(3, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(3).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(3).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(3).CloseRecordset
                  End If
                  Grid(3).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(3, vgCols(), CONDICOES_ESPECIAIS)
         Case 4
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
               Grid(4).CloseRecordset
            Else
               x$ = ExecutaGrid(4, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(4).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(4).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).RecordSet.RsSource) > 0 Then
                     Grid(4).CloseRecordset
                  End If
                  Grid(4).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(4, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionOrcamento
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
   vgTemAlteracaoGrids = Not LockGrids
   Executar PODE_ALTERAR
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub


'executa a pr-validao dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (Sequencia_do_Pedido = 0 And vgSituacao = ACAO_NAVEGANDO And Cancelado = False And PermitidoMenu(mnuNF.Tag))
   Botao(0).Enabled = Ok
   Ok = (IIf(Sequencia_do_Geral = 0 And Fatura_Proforma = True, False, True))
   Label(0).Enabled = Ok And vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   txtCampo(8).Locked = Not (vgAlterar And txtCampo(8).Editable)
   txtCampo(8).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   txtCampo(12).Locked = Not (vgAlterar And txtCampo(12).Editable)
   txtCampo(12).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = False)
   txtCampo(13).Locked = Not (vgAlterar And txtCampo(13).Editable)
   txtCampo(13).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = False)
   txtCampo(14).Locked = Not (vgAlterar And txtCampo(14).Editable)
   txtCampo(14).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   chkCampo(0).Locked = Not (vgAlterar And chkCampo(0).Editable)
   chkCampo(0).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(16).Locked = Not (vgAlterar And txtCampo(16).Editable)
   txtCampo(16).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(17).Locked = Not (vgAlterar And txtCampo(17).Editable)
   txtCampo(17).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(18).Locked = Not (vgAlterar And txtCampo(18).Editable)
   txtCampo(18).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(20).Locked = Not (vgAlterar And txtCampo(20).Editable)
   txtCampo(20).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(21).Locked = Not (vgAlterar And txtCampo(21).Editable)
   txtCampo(21).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   txtCampo(22).Locked = Not (vgAlterar And txtCampo(22).Editable)
   txtCampo(22).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(23).Locked = Not (vgAlterar And txtCampo(23).Editable)
   txtCampo(23).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(24).Locked = Not (vgAlterar And txtCampo(24).Editable)
   txtCampo(24).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   txtCampo(25).Locked = Not (vgAlterar And txtCampo(25).Editable)
   txtCampo(25).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 Or GeralAux![Seqncia Do Vendedor] = 0)
   Label(6).Enabled = Ok And vgAlterar
   Ok = (Sequencia_do_Geral = 0 Or GeralAux![Seqncia Do Vendedor] = 0)
   txtCampo(33).Locked = Not (vgAlterar And txtCampo(33).Editable)
   txtCampo(33).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   Label(15).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma = False)
   Label(16).Enabled = Ok And vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   chkCampo(1).Locked = Not (vgAlterar And chkCampo(1).Editable)
   chkCampo(1).Enabled = Ok Or Not vgAlterar
   Ok = ((Sequencia_do_Geral = 0) And e_Propriedade = True)
   Label(31).Enabled = Ok And vgAlterar
   Ok = ((Sequencia_do_Geral = 0) And e_Propriedade = True)
   txtCampo(48).Locked = Not (vgAlterar And txtCampo(48).Editable)
   txtCampo(48).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = False)
   Label(37).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma = False)
   Label(42).Enabled = Ok And vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   opcPainel1(0).Locked = Not (vgAlterar And opcPainel1(0).Editable)
   opcPainel1(0).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0 And Fatura_Proforma = False)
   opcPainel1(1).Locked = Not (vgAlterar And opcPainel1(1).Editable)
   opcPainel1(1).Enabled = Ok Or Not vgAlterar
   Ok = (IIf(Sequencia_do_Geral = 0 And Fatura_Proforma = True, False, True))
   Label(53).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma = False)
   txtCampo(51).Locked = Not (vgAlterar And txtCampo(51).Editable)
   txtCampo(51).Enabled = Ok Or Not vgAlterar
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   Label(56).Enabled = Ok And vgAlterar
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   txtCampo(54).Locked = Not (vgAlterar And txtCampo(54).Editable)
   txtCampo(54).Enabled = Ok Or Not vgAlterar
   Ok = (Aliquota_do_ISS > 0 And Servicos_do_Orcamento.RecordCount > 0)
   chkCampo(2).Locked = Not (vgAlterar And chkCampo(2).Editable)
   chkCampo(2).Enabled = Ok Or Not vgAlterar
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   Label(59).Enabled = Ok And vgAlterar
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   txtCampo(57).Locked = Not (vgAlterar And txtCampo(57).Editable)
   txtCampo(57).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = True)
   Label(60).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma = True)
   txtCampo(58).Locked = Not (vgAlterar And txtCampo(58).Editable)
   txtCampo(58).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = True)
   Label(61).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma = True)
   txtCampo(59).Locked = Not (vgAlterar And txtCampo(59).Editable)
   txtCampo(59).Enabled = Ok Or Not vgAlterar
   Ok = (Fatura_Proforma = False)
   Label(63).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma)
   Label(64).Enabled = Ok And vgAlterar
   Ok = (Fatura_Proforma)
   txtCampo(60).Locked = Not (vgAlterar And txtCampo(60).Editable)
   txtCampo(60).Enabled = Ok Or Not vgAlterar
   Ok = (Sequencia_do_Geral = 0)
   Botao(5).Enabled = Ok
   Ok = (Conjunto_Avulso = True)
   txtCampo(61).Locked = Not (vgAlterar And txtCampo(61).Editable)
   txtCampo(61).Enabled = Ok Or Not vgAlterar
   Ok = (Not Vazio(Vendedor_Intermediario))
   txtCampo(63).Locked = Not (vgAlterar And txtCampo(63).Editable)
   txtCampo(63).Enabled = Ok Or Not vgAlterar
   Ok = (Not Vazio(Rebiut))
   txtCampo(65).Locked = Not (vgAlterar And txtCampo(65).Editable)
   txtCampo(65).Enabled = Ok Or Not vgAlterar
   Ok = (vgSituacao = ACAO_NAVEGANDO And Sequencia_do_Pedido = 0)
   Botao(6).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And Sequencia_do_Pedido = 0)
   Botao(7).Enabled = Ok
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   Label(167).Enabled = Ok And vgAlterar
   Ok = (Servicos_do_Orcamento.RecordCount > 0)
   Label(168).Enabled = Ok And vgAlterar
   Ok = (Venda_Fechada Or Ordem_Interna)
   txtCampo(149).Locked = Not (vgAlterar And txtCampo(149).Editable)
   txtCampo(149).Enabled = Ok Or Not vgAlterar
   Ok = (Venda_Fechada And Ordem_Interna = False)
   txtCampo(150).Locked = Not (vgAlterar And txtCampo(150).Editable)
   txtCampo(150).Enabled = Ok Or Not vgAlterar
   Ok = (Ordem_Interna = False)
   Label(176).Enabled = Ok And vgAlterar
   Ok = (Ordem_Interna = False)
   txtCampo(151).Locked = Not (vgAlterar And txtCampo(151).Editable)
   txtCampo(151).Enabled = Ok Or Not vgAlterar
   Dim modFrete As Integer
   modFrete = CodigoFrete()
   Dim bloquear As Boolean
   bloquear = (modFrete = 3 Or modFrete = 4)
    If bloquear Then Call LimpaTransportadora
   labFdo55.Enabled = Not bloquear
   txtCp(55).Enabled = (Ok And Not bloquear) Or Not vgAlterar



   Ok = (vgSituacao = ACAO_NAVEGANDO)
   Botao(9).Enabled = Ok
   Ok = (Vazio(Nome_Cliente) And vgSituacao = ACAO_INCLUINDO)
   txtCampo(156).Locked = Not (vgAlterar And txtCampo(156).Editable)
   txtCampo(156).Enabled = Ok Or Not vgAlterar
   Ok = (Conjuntos_do_Orcamento.RecordCount > 0 Or Pecas_do_Orcamento.RecordCount > 0)
   chkCampo(4).Locked = Not (vgAlterar And chkCampo(4).Editable)
   chkCampo(4).Enabled = Ok Or Not vgAlterar
   Ok = (TemPropriedade = True And Ordem_Interna = False)
   Label(182).Enabled = Ok And vgAlterar
   Ok = (TemPropriedade = True And Ordem_Interna = False)
   txtCampo(157).Locked = Not (vgAlterar And txtCampo(157).Editable)
   txtCampo(157).Enabled = Ok Or Not vgAlterar
   Ok = (Vazio(Nome_Cliente))
   Label(185).Enabled = Ok And vgAlterar
   Ok = (TemPropriedade = True And Ordem_Interna = False)
   Botao(12).Enabled = Ok
   Ok = (IIf(Me.Caption <> "Ordem de Produo Interna", Venda_Fechada = 0 Or vgPWUsuario = "YGOR", vgPWUsuario = "YGOR" Or vgPWUsuario = "RODRIGO" Or vgPWUsuario = "WAGNER"))
   chkCampo(6).Locked = Not (vgAlterar And chkCampo(6).Editable)
   chkCampo(6).Enabled = Ok Or Not vgAlterar
   Ok = (False)
   chkCampo(8).Locked = Not (vgAlterar And chkCampo(8).Editable)
   chkCampo(8).Enabled = Ok Or Not vgAlterar
   Ok = (vgPWUsuario = "YGOR" Or vgPWUsuario = "MAYSA" Or vgPWUsuario = "VANESSA" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "ALEXANDRA" Or vgPWUsuario = "WAGNER")
   chkCampo(9).Locked = Not (vgAlterar And chkCampo(9).Editable)
   chkCampo(9).Enabled = Ok Or Not vgAlterar
   Ok = (vgSituacao = ACAO_NAVEGANDO)
   Botao(13).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And Existe(Parametros![Diretorio das Fotos] & "Orc_" & Sequencia_do_Orcamento & ".jpg"))
   Botao(14).Enabled = Ok
   Ok = (vgPWUsuario = "YGOR" Or vgPWUsuario = "MAYSA" Or vgPWUsuario = "VANESSA" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "ALEXANDRA" Or vgPWUsuario = "WAGNER")
   chkCampo(10).Locked = Not (vgAlterar And chkCampo(10).Editable)
   chkCampo(10).Enabled = Ok Or Not vgAlterar
   Ok = (False)
   chkCampo(11).Locked = Not (vgAlterar And chkCampo(11).Editable)
   chkCampo(11).Enabled = Ok Or Not vgAlterar
   Ok = (vgSituacao = ACAO_NAVEGANDO)
   Botao(16).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub


'coloca os campos visveis segundo a condio
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(0).Visible = (Fatura_Proforma = False)
   Botao(0).Visible = (Ordem_Interna = False)
   txtCampo(1).Visible = (Sequencia_do_Geral > 0)
   txtCampo(2).Visible = (Sequencia_do_Geral > 0)
   txtCampo(3).Visible = (Sequencia_do_Geral > 0)
   txtCampo(5).Visible = (False)
   txtCampo(6).Visible = (Sequencia_do_Geral > 0)
   Label(1).Visible = (Sequencia_do_Geral > 0)
   Label(2).Visible = (Sequencia_do_Geral > 0)
   txtCampo(7).Visible = (Sequencia_do_Geral > 0)
   txtCampo(8).Visible = (Sequencia_do_Geral = 0)
   Label(3).Visible = (Sequencia_do_Geral > 0)
   txtCampo(10).Visible = (Sequencia_do_Geral > 0)
   txtCampo(11).Visible = (Sequencia_do_Geral <> 0)
   txtCampo(12).Visible = (Sequencia_do_Geral = 0)
   txtCampo(13).Visible = (Sequencia_do_Geral = 0)
   txtCampo(14).Visible = (Sequencia_do_Geral = 0)
   txtCampo(15).Visible = (Sequencia_do_Geral > 0)
   txtCampo(16).Visible = (Sequencia_do_Geral = 0)
   txtCampo(17).Visible = (Sequencia_do_Geral = 0)
   txtCampo(19).Visible = (Sequencia_do_Geral > 0)
   txtCampo(20).Visible = (Sequencia_do_Geral = 0)
   txtCampo(21).Visible = (Sequencia_do_Geral = 0)
   txtCampo(22).Visible = (Sequencia_do_Geral = 0)
   txtCampo(23).Visible = (Sequencia_do_Geral = 0)
   txtCampo(24).Visible = (Sequencia_do_Geral = 0)
   txtCampo(25).Visible = (Sequencia_do_Geral = 0)
   txtCampo(26).Visible = (Sequencia_do_Geral > 0)
   txtCampo(28).Visible = (Sequencia_do_Geral > 0)
   txtCampo(29).Visible = (Sequencia_do_Geral > 0)
   txtCampo(30).Visible = (Sequencia_do_Geral > 0)
   Label(6).Visible = (Ordem_Interna = False)
   Botao(1).Visible = (Ordem_Interna = False)
   txtCampo(33).Visible = (Ordem_Interna = False)
   txtCampo(34).Visible = (Ordem_Interna = False)
   Botao(2).Visible = (Ordem_Interna = False)
   Label(14).Visible = (Sequencia_do_Geral = 0)
   Label(16).Visible = (Sequencia_do_Geral = 0)
   Label(36).Visible = (Sequencia_do_Geral = 0)
   Label(37).Visible = (Sequencia_do_Geral = 0)
   Label(39).Visible = (Ordem_Interna = False)
   Label(41).Visible = (Sequencia_do_Geral > 0)
   Label(42).Visible = (Sequencia_do_Geral = 0)
   Label(43).Visible = (Sequencia_do_Geral = 0)
   Label(44).Visible = (Ordem_Interna = False)
   Label(45).Visible = (Ordem_Interna = False)
   txtCampo(50).Visible = (Ordem_Interna = False)
   Label(51).Visible = (Sequencia_do_Geral > 0)
   Label(52).Visible = (Sequencia_do_Geral > 0)
   txtCampo(51).Visible = (Sequencia_do_Geral = 0)
   Label(56).Visible = (Ordem_Interna = False)
   txtCampo(54).Visible = (Ordem_Interna = False)
   chkCampo(2).Visible = (Ordem_Interna = False)
   Label(57).Visible = (Ordem_Interna = False)
   txtCampo(55).Visible = (Ordem_Interna = False)
   Botao(4).Visible = (Ordem_Interna = False)
   Label(58).Visible = (Servicos_do_Orcamento.RecordCount > 0)
   txtCampo(56).Visible = (Servicos_do_Orcamento.RecordCount > 0)
   Label(60).Visible = (Ordem_Interna = False)
   txtCampo(58).Visible = (Ordem_Interna = False)
   Label(61).Visible = (Ordem_Interna = False)
   txtCampo(59).Visible = (Ordem_Interna = False)
   Label(62).Visible = (Sequencia_do_Geral > 0)
   Label(63).Visible = (Sequencia_do_Geral = 0)
   Label(64).Visible = (Fatura_Proforma)
   txtCampo(60).Visible = (Fatura_Proforma)
   Botao(5).Visible = (Fatura_Proforma)
   chkCampo(3).Visible = (Ordem_Interna = False)
   txtCampo(61).Visible = (Ordem_Interna = False)
   Label(65).Visible = (Ordem_Interna = False)
   txtCampo(62).Visible = (Ordem_Interna = False)
   Label(66).Visible = (Ordem_Interna = False)
   txtCampo(63).Visible = (Ordem_Interna = False)
   Label(67).Visible = (Ordem_Interna = False)
   Label(68).Visible = (Ordem_Interna = False)
   txtCampo(64).Visible = (Ordem_Interna = False)
   txtCampo(65).Visible = (Ordem_Interna = False)
   Label(146).Visible = (Ordem_Interna = False)
   txtCampo(137).Visible = (Ordem_Interna = False)
   Label(167).Visible = (Ordem_Interna = False)
   Label(168).Visible = (Ordem_Interna = False)
   txtCampo(143).Visible = (Ordem_Interna = False)
   txtCampo(144).Visible = (Ordem_Interna = False)
   Label(169).Visible = (Ordem_Interna = False)
   txtCampo(145).Visible = (Ordem_Interna = False)
   Label(170).Visible = (Ordem_Interna = False)
   txtCampo(146).Visible = (Ordem_Interna = False)
   Label(172).Visible = (Ordem_Interna = False)
   txtCampo(147).Visible = (Ordem_Interna = False)
   Label(176).Visible = (Ordem_Interna = False)
   txtCampo(151).Visible = (Ordem_Interna = False)
   Botao(8).Visible = (Ordem_Interna = False)
   txtCampo(155).Visible = (Ordem_Interna = False)
   Botao(9).Visible = (Ordem_Interna = False)
   Label(180).Visible = (False)
   picBox(0).Visible = (Me.Caption <> "Ordem de Produo Interna")
   chkCampo(4).Visible = (Ordem_Interna = False)
   Botao(11).Visible = (False)
   chkCampo(5).Visible = (IIf(Fatura_Proforma Or Ordem_Interna, False, True))
   Label(186).Visible = (VerificaDebitos() = True)
   chkCampo(6).Visible = (vgPWUsuario = "YGOR" Or Ordem_Interna = False)
   chkCampo(7).Visible = (IIf(Fatura_Proforma Or Ordem_Interna, False, True))
   txtCampo(159).Visible = (Fatura_Proforma = True)
   chkCampo(9).Visible = (Ordem_Interna = False)
   Botao(13).Visible = (Ordem_Interna = False)
   Botao(14).Visible = (Ordem_Interna = False)
   txtCampo(160).Visible = (Ordem_Interna = False)
   Botao(15).Visible = (Ordem_Interna = False)
   Label(187).Visible = (Ordem_Interna = False)
   chkCampo(10).Visible = (Ordem_Interna = False)
   mmCampo(1).Visible = (Me.Caption = "Ordem de Produo Interna")
   Botao(16).Visible = (Ordem_Interna = True)
   Label(188).Visible = (Ordem_Interna = True)
   txtCampo(161).Visible = (Ordem_Interna = True)
   Label(189).Visible = (Ordem_Interna = True)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub



'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 17
            PoeRelEFiltroCbo 17
         Case 33
            PoeRelEFiltroCbo 33
            If Len(txtCp(33).Text) = 0 Then
               txtCampo(33).Value = GeralAux![Seqncia Do Vendedor]
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 34
            PoeRelEFiltroCbo 34
            If Len(txtCp(34).Text) = 0 Then
               txtCampo(34).Value = PegaNCMPadrao
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 35
            If Len(txtCp(35).Text) = 0 Then
               txtCampo(35).Value = "Vista"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 54
            If ValBrasil(txtCp(54).Text) = 0 Then
               txtCampo(54).Value = 3
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 55
            PoeRelEFiltroCbo 55
         Case 60
            PoeRelEFiltroCbo 60
         Case 62
            PoeRelEFiltroCbo 62
         Case 64
            PoeRelEFiltroCbo 64
         Case 73
            If Len(txtCp(73).Text) = 0 Then
               txtCampo(73).Value = "SETORIAL"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 77
            If ValBrasil(txtCp(77).Text) = 0 Then
               txtCampo(77).Value = 100
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 99
            PoeRelEFiltroCbo 99
         Case 100
            PoeRelEFiltroCbo 100
         Case 101
            PoeRelEFiltroCbo 101
         Case 137
            PoeRelEFiltroCbo 137
         Case 149
            If Len(txtCp(149).Text) = 0 Then
               txtCampo(149).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 156
            PoeRelEFiltroCbo 156
            If Len(txtCp(156).Text) = 0 Then
               txtCampo(156).Value = IIf(Ordem_Interna, 517, 0)
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 157
            PoeRelEFiltroCbo 157
         Case 158
            If Len(txtCp(158).Text) = 0 Then
               txtCampo(158).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 161
            PoeRelEFiltroCbo 161
      End Select
   End If
   txtCampo(Index).GotFocus
End Sub


Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   ' 1) Se apertou ENTER em modo edio, dispara visibilidade e pr-validao
   If KeyCode = vbKeyReturn And vgSituacao <> ACAO_NAVEGANDO Then
      ExecutaVisivel
      ExecutaPreValidacao
   End If

   ' 2) Se for Delete nos campos 152,153 ou 154, propaga para o GAS control
   If (Index = 152 Or Index = 153 Or Index = 154) And KeyCode = vbKeyDelete Then
      txtCampo(Index).KeyDown KeyCode, Shift
      Exit Sub   ' j propagou, sai para evitar chamada duplicada
   End If

   ' 3) Para todas as outras teclas (inclusive Delete em outros campos), propaga normalmente
   txtCampo(Index).KeyDown KeyCode, Shift
End Sub




'> dispara em cada tecla pressionada
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)

   Select Case Index

      Case 4
         ' seu tratamento atual
         LimitaCampo KeyAscii, 445

      Case 152
         ' PlacaVeiculo: apenas AZ e 09, fora uppercase
         If KeyAscii = vbKeyBack Then
            ' permite apagar
         ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or _
                (KeyAscii >= 65 And KeyAscii <= 90) Or _
                (KeyAscii >= 97 And KeyAscii <= 122) Then
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
         Else
            KeyAscii = 0
         End If

      Case 153
         ' UfPlaca: apenas letras AZ, fora uppercase
         If KeyAscii = vbKeyBack Then
         ElseIf (KeyAscii >= 65 And KeyAscii <= 90) Or _
                (KeyAscii >= 97 And KeyAscii <= 122) Then
            KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
         Else
            KeyAscii = 0
         End If

      Case 154
         ' NumAntt: apenas dgitos
         If KeyAscii = vbKeyBack Then
         ElseIf Not (KeyAscii >= 48 And KeyAscii <= 57) Then
            KeyAscii = 0
         End If

      Case Else
         ' nada para outros ndices

   End Select

   ' Propaga o KeyPress para o GAS control (dispara txtCp_Change)
   txtCampo(Index).KeyPress KeyAscii
End Sub





'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
   txtCampo(Index).LostFocus
   If vgSituacao <> ACAO_NAVEGANDO Then           'se tela esta em edio
      InicializaApelidos COM_TEXTBOX              'pega apelidos dos campos
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
   End If
   Select Case Index
      Case 156
         BuscaVendedor
      Case 161
         PreValidaVinculo
   End Select
End Sub



'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Revenda
      ElseIf Index = 1 Then
         chkCampo(1).Value = e_Propriedade

      ElseIf Index = 2 Then
         chkCampo(2).Value = Reter_ISS

      ElseIf Index = 3 Then
         chkCampo(3).Value = Conjunto_Avulso

      ElseIf Index = 4 Then
         chkCampo(4).Value = Ocultar_Valor_Unitario

      ElseIf Index = 5 Then
         chkCampo(5).Value = Entrega_Futura

      ElseIf Index = 6 Then
         chkCampo(6).Value = Venda_Fechada

      ElseIf Index = 7 Then
         chkCampo(7).Value = Orcamento_Avulso

      ElseIf Index = 8 Then
         chkCampo(8).Value = Fatura_Proforma

      ElseIf Index = 9 Then
         chkCampo(9).Value = Nao_Movimentar_Estoque

      ElseIf Index = 10 Then
         chkCampo(10).Value = Refaturamento

      ElseIf Index = 11 Then
         chkCampo(11).Value = Ordem_Interna
      End If
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
      chkCampo(Index).Change
   End If
   Select Case Index
      Case 6
         DesmarcaFw
   End Select
End Sub


'evento - quando o check receber o foco
Private Sub chkCp_GotFocus(Index As Integer)
   chkCampo(Index).GotFocus
End Sub


'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   chkCampo(Index).KeyDown KeyCode, Shift
End Sub


'evento - quando qq tecla for digitada no check
Private Sub chkCp_KeyPress(Index As Integer, KeyAscii As Integer)
   chkCampo(Index).KeyPress KeyAscii
End Sub


'evento - quando o check perder o foco
Private Sub chkCp_LostFocus(Index As Integer)
   chkCampo(Index).LostFocus
End Sub


'evento - quando qq tecla for digitada no campo
Private Sub opcPainel1Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel1(Index).KeyPress KeyAscii
End Sub


'evento - quando qq tecla for digitada no campo
Private Sub opcPainel1Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel1(Index).KeyDown KeyCode, Shift
End Sub


'evento - quando o campo receber o foco
Private Sub opcPainel1Cp_GotFocus(Index As Integer)
   opcPainel1(Index).GotFocus
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 17
            PoeRelEFiltroCbo 17
         Case 33
            PoeRelEFiltroCbo 33
         Case 34
            PoeRelEFiltroCbo 34
         Case 55
            PoeRelEFiltroCbo 55
         Case 60
            PoeRelEFiltroCbo 60
         Case 62
            PoeRelEFiltroCbo 62
         Case 64
            PoeRelEFiltroCbo 64
         Case 99
            PoeRelEFiltroCbo 99
         Case 100
            PoeRelEFiltroCbo 100
         Case 101
            PoeRelEFiltroCbo 101
         Case 137
            PoeRelEFiltroCbo 137
         Case 156
            PoeRelEFiltroCbo 156
         Case 157
            PoeRelEFiltroCbo 157
         Case 161
            PoeRelEFiltroCbo 161
      End Select
   End If
End Sub


'evento - quando o campo perder o foco
Private Sub opcPainel1Cp_LostFocus(Index As Integer)
   opcPainel1(Index).LostFocus
End Sub


'evento - quando qq tecla for digitada no campo
Private Sub opcPainel2Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel2(Index).KeyPress KeyAscii
End Sub


'evento - quando qq tecla for digitada no campo
Private Sub opcPainel2Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel2(Index).KeyDown KeyCode, Shift
End Sub


'evento - quando o campo receber o foco
Private Sub opcPainel2Cp_GotFocus(Index As Integer)
   opcPainel2(Index).GotFocus
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 17
            PoeRelEFiltroCbo 17
         Case 33
            PoeRelEFiltroCbo 33
         Case 34
            PoeRelEFiltroCbo 34
         Case 55
            PoeRelEFiltroCbo 55
         Case 60
            PoeRelEFiltroCbo 60
         Case 62
            PoeRelEFiltroCbo 62
         Case 64
            PoeRelEFiltroCbo 64
         Case 99
            PoeRelEFiltroCbo 99
         Case 100
            PoeRelEFiltroCbo 100
         Case 101
            PoeRelEFiltroCbo 101
         Case 137
            PoeRelEFiltroCbo 137
         Case 156
            PoeRelEFiltroCbo 156
         Case 157
            PoeRelEFiltroCbo 157
         Case 161
            PoeRelEFiltroCbo 161
      End Select
   End If
End Sub


'evento - quando o campo perder o foco
Private Sub opcPainel2Cp_LostFocus(Index As Integer)
   opcPainel2(Index).LostFocus
End Sub


'evento - quando o formulrio receber o foco
Private Sub Form_Activate()
   If vgPriVez = False Then
      Screen.MousePointer = vbHourglass           'mouse = ampulheta
   Else
      vgPriVez = False
   End If
   Posiciona
   AtivaForm Me
   
   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   If FormEstaAberto("frmEnviaEmail") Then
      If Not frmEnviaEMail.Visible Then
         Unload vgFrmImpCons
         Set vgFrmImpCons = Nothing
         Unload frmEnviaEMail
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub


'evento - inicializao do formulrio
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(73350)
   vgFormID = 714
   vgIdentTab = "Oramento"
   vgFiltroEmUso = -1
   vgIndDefault = "Seqncia do Oramento"
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgRepeticao = -99
   vgAlterar = False
   vgUltAlterar = False
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 200
   vgSituacao = ACAO_NAVEGANDO
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(64855))
   Set Botao(0).PictureDisabled = LoadPicture(LoadGasPicture(64856))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(64857))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(64858))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(64859))
   Set Botao(4).Picture = LoadPicture(LoadGasPicture(64860))
   Set Botao(5).Picture = LoadPicture(LoadGasPicture(64861))
   Set Botao(8).Picture = LoadPicture(LoadGasPicture(64862))
   Set Botao(9).Picture = LoadPicture(LoadGasPicture(64863))
   Set Botao(9).PictureDisabled = LoadPicture(LoadGasPicture(64864))
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(64865))
   Set Botao(10).Picture = LoadPicture(LoadGasPicture(64866))
   Set Botao(12).Picture = LoadPicture(LoadGasPicture(64867))
   Set Botao(13).Picture = LoadPicture(LoadGasPicture(64868))
   Set Botao(13).PictureDisabled = LoadPicture(LoadGasPicture(64869))
   Set Botao(14).Picture = LoadPicture(LoadGasPicture(64870))
   Set Botao(14).PictureDisabled = LoadPicture(LoadGasPicture(64871))
   Set Botao(15).Picture = LoadPicture(LoadGasPicture(64872))
   Set mmCampo(1).Picture = LoadPicture(LoadGasPicture(64873))
   Set Botao(16).Picture = LoadPicture(LoadGasPicture(64874))
   Set Botao(16).PictureDisabled = LoadPicture(LoadGasPicture(64875))
   Set txtSequencia_do_Orcamento = txtCampo(0)
   Set Aba1 = Tabs(0)
   Set txtCEP = txtCampo(1)
   Set txtCaixaPostal = txtCampo(2)
   Set txtFone = txtCampo(3)
   Set txtObservacao = txtCampo(4)
   Set txtMemoAuxiliar = txtCampo(5)
   Set lblRGIE = Label(0)
   Set txtCPFCNPJ_F = txtCampo(6)
   Set lblCPFCNPJ_F = Label(1)
   Set txtFax = txtCampo(7)
   Set txtRGIE_F = txtCampo(10)
   Set txtCPFCNPJ = txtCampo(13)
   Set txtRGIE = txtCampo(14)
   Set txtEmail = txtCampo(15)
   Set txtMunicipio = txtCampo(19)
   Set txtEndereco = txtCampo(26)
   Set txtUF = txtCampo(27)
   Set txtBairro = txtCampo(28)
   Set txtNumero = txtCampo(29)
   Set txtComplemento = txtCampo(30)
   Set txtVendedor = txtCampo(33)
   Set grdConjuntos = Grid(0)
   Set grdPecas = Grid(1)
   Set Grdparcelamento = Grid(2)
   Set txtForma_de_Pagamento = txtCampo(35)
   Set GrdProdutos = Grid(3)
   Set grdServicos = Grid(4)
   Set lblParcelamento = Label(7)
   Set lblCPFCNPJ = Label(16)
   Set Veiculo = Label(40)
   Set txtISS = txtCampo(56)
   Set Txtperdas1 = txtCampo(70)
   Set Txtperdas2 = txtCampo(71)
   Set Lblvazao = Label(80)
   Set Lblvazaototal = txtCampo(76)
   Set Txtvelodesloca = txtCampo(79)
   Set Txtperdas3 = txtCampo(80)
   Set Txtdeslocamento = txtCampo(83)
   Set Txtprecipitacaolic = txtCampo(84)
   Set Txtvazaoporturno = txtCampo(85)
   Set Txtalturamanometrica = txtCampo(86)
   Set Txtareapordia = txtCampo(87)
   Set Txttempo1 = txtCampo(91)
   Set Txtareafx = txtCampo(92)
   Set Txtfaixasirrigadas = txtCampo(93)
   Set Txtturno = txtCampo(94)
   Set Txtdiam1 = txtCampo(102)
   Set Txtdiam2 = txtCampo(103)
   Set Txtdiam3 = txtCampo(104)
   Set Txtcoef1 = txtCampo(105)
   Set Txtcoef2 = txtCampo(106)
   Set Txtcoef3 = txtCampo(107)
   Set txtHF1 = txtCampo(108)
   Set txtHF2 = txtCampo(109)
   Set txtHF3 = txtCampo(110)
   Set Txtvelo1 = txtCampo(111)
   Set Txtvelo2 = txtCampo(112)
   Set Txtvelo3 = txtCampo(113)
   Set Txtperdashidro = txtCampo(114)
   Set Lblvazaototal2 = txtCampo(121)
   Set Txtpressao = txtCampo(122)
   Set Txtrendimento = txtCampo(123)
   Set Txtpotencia = txtCampo(125)
   Set Txtrotacaomotor = txtCampo(132)
   Set Txtdemandamotor = txtCampo(133)
   Set Txtamperagem = txtCampo(135)
   Set Txtconsumo = txtCampo(136)
   Set txtFrete = txtCampo(151)
   Set txtNF = txtCampo(155)
   Set lblAjuste = Label(180)
   Set lblOrcamento = Label(181)
   Set txtPropriedade = txtCampo(157)
   Set txtProjeto = txtCampo(160)
   Set lblVinculo = Label(189)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   If vgPWGrupo = "VENDAS" Then
      vgFiltroInicial$ = vgFiltroInicial$ + FiltroOrc()
   ElseIf vgPWGrupo <> "VENDAS" Then
      vgFiltroInicial$ = vgFiltroInicial$ + "[Seqncia do Oramento] > " & 0 & " AND [Fatura Proforma] = " & 0 & "  And [Ordem Interna] = " & 0 & " AND (([Data do Fechamento] IS NULL AND Cancelado = " & 0 & ") Or ([Data de Emisso] >= CONVERT(VARCHAR, '" & Format$(DateAdd("D", -90, Date), "yyyy-mm-dd hh:mm:ss") & "', 120)))"
   End If
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Botao(0).Caption = LoadGasString(73355)
   Tabs(0).TabCaption(0) = LoadGasString(73356)
   Tabs(0).TabCaption(1) = LoadGasString(73357)
   Tabs(0).TabCaption(2) = LoadGasString(73358)
   Tabs(0).TabCaption(3) = LoadGasString(73359)
   Tabs(0).TabCaption(4) = LoadGasString(73360)
   Tabs(0).TabCaption(5) = LoadGasString(73361)
   Tabs(0).TabCaption(6) = LoadGasString(73362)
   Tabs(0).TabCaption(7) = LoadGasString(73363)
   Tabs(0).TabCaption(8) = LoadGasString(73364)
   Label(0).Caption = LoadGasString(73365)
   Label(1).Caption = LoadGasString(73366)
   Label(2).Caption = LoadGasString(73367)
   Label(3).Caption = LoadGasString(73368)
   chkCampo(0).Caption = LoadGasString(73369)
   Label(4).Caption = LoadGasString(73370)
   vgTooltips.AddTool txtCampo(20).CtPri, 0, LoadGasString(73371)
   vgTooltips.AddTool txtCampo(25).CtPri, 0, LoadGasString(73372)
   Label(5).Caption = LoadGasString(73373)
   vgTooltips.AddTool txtCampo(32).CtPri, 0, LoadGasString(73374)
   Label(6).Caption = LoadGasString(73375)
   vgTooltips.AddTool Botao(1), 0, LoadGasString(73376)
   vgTooltips.AddTool txtCampo(33).CtPri, 0, LoadGasString(73377)
   Label(8).Caption = LoadGasString(73378)
   Label(9).Caption = LoadGasString(73379)
   Label(10).Caption = LoadGasString(73380)
   Label(11).Caption = LoadGasString(73381)
   vgTooltips.AddTool Botao(2), 0, LoadGasString(73382)
   vgTooltips.AddTool Botao(3), 0, LoadGasString(73383)
   Label(12).Caption = LoadGasString(73384)
   Label(13).Caption = LoadGasString(73385)
   Label(14).Caption = LoadGasString(73386)
   Label(15).Caption = LoadGasString(73387)
   Label(16).Caption = LoadGasString(73388)
   Label(17).Caption = LoadGasString(73389)
   Label(18).Caption = LoadGasString(73390)
   Label(19).Caption = LoadGasString(73391)
   Label(20).Caption = LoadGasString(73392)
   Label(21).Caption = LoadGasString(73393)
   Label(22).Caption = LoadGasString(73394)
   Label(23).Caption = LoadGasString(73395)
   Label(24).Caption = LoadGasString(73396)
   Label(25).Caption = LoadGasString(73397)
   Label(26).Caption = LoadGasString(73398)
   Label(27).Caption = LoadGasString(73399)
   Label(28).Caption = LoadGasString(73400)
   Label(29).Caption = LoadGasString(73401)
   Label(30).Caption = LoadGasString(73402)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(73403)
   Label(31).Caption = LoadGasString(73404)
   Label(32).Caption = LoadGasString(73405)
   Label(33).Caption = LoadGasString(73406)
   Label(34).Caption = LoadGasString(73407)
   Label(35).Caption = LoadGasString(73408)
   Label(36).Caption = LoadGasString(73409)
   Label(37).Caption = LoadGasString(73410)
   Label(38).Caption = LoadGasString(73411)
   Label(39).Caption = LoadGasString(73412)
   Label(40).Caption = LoadGasString(73413)
   Label(41).Caption = LoadGasString(73414)
   Label(42).Caption = LoadGasString(73415)
   Label(43).Caption = LoadGasString(73416)
   Label(44).Caption = LoadGasString(73417)
   Label(45).Caption = LoadGasString(73418)
   Label(46).Caption = LoadGasString(73419)
   Label(47).Caption = LoadGasString(73420)
   Label(48).Caption = LoadGasString(73421)
   Label(49).Caption = LoadGasString(73422)
   Label(50).Caption = LoadGasString(73423)
   Label(51).Caption = LoadGasString(73424)
   Label(52).Caption = LoadGasString(73425)
   opcPainel1(0).Caption = LoadGasString(73426)
   opcPainel1(1).Caption = LoadGasString(73427)
   opcPainel2(0).Caption = LoadGasString(73428)
   opcPainel2(1).Caption = LoadGasString(73429)
   Label(53).Caption = LoadGasString(73430)
   Label(54).Caption = LoadGasString(73431)
   Label(55).Caption = LoadGasString(73432)
   Label(56).Caption = LoadGasString(73433)
   vgTooltips.AddTool txtCampo(54).CtPri, 0, LoadGasString(73434)
   vgTooltips.AddTool chkCampo(2).CtPri, 0, LoadGasString(73435)
   Label(57).Caption = LoadGasString(73436)
   vgTooltips.AddTool Botao(4), 0, LoadGasString(73437)
   Label(58).Caption = LoadGasString(73438)
   Label(59).Caption = LoadGasString(73439)
   vgTooltips.AddTool txtCampo(57).CtPri, 0, LoadGasString(73440)
   Label(60).Caption = LoadGasString(73441)
   vgTooltips.AddTool txtCampo(58).CtPri, 0, LoadGasString(73442)
   Label(61).Caption = LoadGasString(73443)
   vgTooltips.AddTool txtCampo(59).CtPri, 0, LoadGasString(73444)
   Label(62).Caption = LoadGasString(73445)
   Label(63).Caption = LoadGasString(73446)
   Label(64).Caption = LoadGasString(73447)
   vgTooltips.AddTool txtCampo(60).CtPri, 0, LoadGasString(73448)
   vgTooltips.AddTool Botao(5), 0, LoadGasString(73449)
   vgTooltips.AddTool chkCampo(3).CtPri, 0, LoadGasString(73450)
   chkCampo(3).Caption = LoadGasString(73451)
   vgTooltips.AddTool txtCampo(61).CtPri, 0, LoadGasString(73452)
   Label(65).Caption = LoadGasString(73453)
   vgTooltips.AddTool txtCampo(62).CtPri, 0, LoadGasString(73454)
   Label(66).Caption = LoadGasString(73455)
   vgTooltips.AddTool txtCampo(63).CtPri, 0, LoadGasString(73456)
   Label(67).Caption = LoadGasString(73457)
   Label(68).Caption = LoadGasString(73458)
   vgTooltips.AddTool txtCampo(64).CtPri, 0, LoadGasString(73459)
   vgTooltips.AddTool txtCampo(65).CtPri, 0, LoadGasString(73460)
   Botao(6).Caption = LoadGasString(73461)
   Botao(7).Caption = LoadGasString(73462)
   Label(69).Caption = LoadGasString(73463)
   vgTooltips.AddTool txtCampo(66).CtPri, 0, LoadGasString(73464)
   Label(70).Caption = LoadGasString(73465)
   Label(71).Caption = LoadGasString(73466)
   vgTooltips.AddTool txtCampo(67).CtPri, 0, LoadGasString(73467)
   Label(72).Caption = LoadGasString(73468)
   vgTooltips.AddTool txtCampo(68).CtPri, 0, LoadGasString(73469)
   Label(73).Caption = LoadGasString(73470)
   Label(74).Caption = LoadGasString(73471)
   vgTooltips.AddTool txtCampo(69).CtPri, 0, LoadGasString(73472)
   Label(75).Caption = LoadGasString(73473)
   vgTooltips.AddTool txtCampo(70).CtPri, 0, LoadGasString(73474)
   Label(76).Caption = LoadGasString(73475)
   vgTooltips.AddTool txtCampo(71).CtPri, 0, LoadGasString(73476)
   Label(77).Caption = LoadGasString(73477)
   vgTooltips.AddTool txtCampo(72).CtPri, 0, LoadGasString(73478)
   Label(78).Caption = LoadGasString(73479)
   Label(79).Caption = LoadGasString(73480)
   vgTooltips.AddTool txtCampo(73).CtPri, 0, LoadGasString(73481)
   Label(81).Caption = LoadGasString(73482)
   vgTooltips.AddTool txtCampo(74).CtPri, 0, LoadGasString(73483)
   vgTooltips.AddTool txtCampo(75).CtPri, 0, LoadGasString(73484)
   vgTooltips.AddTool txtCampo(76).CtPri, 0, LoadGasString(73485)
   Label(83).Caption = LoadGasString(73486)
   Label(84).Caption = LoadGasString(73487)
   vgTooltips.AddTool txtCampo(77).CtPri, 0, LoadGasString(73488)
   vgTooltips.AddTool txtCampo(78).CtPri, 0, LoadGasString(73489)
   Label(85).Caption = LoadGasString(73490)
   vgTooltips.AddTool txtCampo(79).CtPri, 0, LoadGasString(73491)
   vgTooltips.AddTool txtCampo(80).CtPri, 0, LoadGasString(73492)
   vgTooltips.AddTool txtCampo(81).CtPri, 0, LoadGasString(73493)
   Label(86).Caption = LoadGasString(73494)
   Label(87).Caption = LoadGasString(73495)
   vgTooltips.AddTool txtCampo(82).CtPri, 0, LoadGasString(73496)
   Label(88).Caption = LoadGasString(73497)
   vgTooltips.AddTool txtCampo(83).CtPri, 0, LoadGasString(73498)
   Label(89).Caption = LoadGasString(73499)
   vgTooltips.AddTool txtCampo(84).CtPri, 0, LoadGasString(73500)
   Label(90).Caption = LoadGasString(73501)
   vgTooltips.AddTool txtCampo(85).CtPri, 0, LoadGasString(73502)
   vgTooltips.AddTool txtCampo(86).CtPri, 0, LoadGasString(73503)
   vgTooltips.AddTool txtCampo(87).CtPri, 0, LoadGasString(73504)
   Label(91).Caption = LoadGasString(73505)
   Label(92).Caption = LoadGasString(73506)
   vgTooltips.AddTool txtCampo(88).CtPri, 0, LoadGasString(73507)
   Label(93).Caption = LoadGasString(73508)
   vgTooltips.AddTool txtCampo(89).CtPri, 0, LoadGasString(73509)
   Label(94).Caption = LoadGasString(73510)
   Label(95).Caption = LoadGasString(73511)
   vgTooltips.AddTool txtCampo(90).CtPri, 0, LoadGasString(73512)
   vgTooltips.AddTool txtCampo(91).CtPri, 0, LoadGasString(73513)
   vgTooltips.AddTool txtCampo(92).CtPri, 0, LoadGasString(73514)
   vgTooltips.AddTool txtCampo(93).CtPri, 0, LoadGasString(73515)
   vgTooltips.AddTool txtCampo(94).CtPri, 0, LoadGasString(73516)
   Label(96).Caption = LoadGasString(73517)
   Label(97).Caption = LoadGasString(73518)
   Label(98).Caption = LoadGasString(73519)
   vgTooltips.AddTool txtCampo(95).CtPri, 0, LoadGasString(73520)
   Label(99).Caption = LoadGasString(73521)
   Label(100).Caption = LoadGasString(73522)
   Label(101).Caption = LoadGasString(73523)
   Label(102).Caption = LoadGasString(73524)
   Label(103).Caption = LoadGasString(73525)
   Label(104).Caption = LoadGasString(73526)
   Label(105).Caption = LoadGasString(73527)
   Label(106).Caption = LoadGasString(73528)
   vgTooltips.AddTool txtCampo(96).CtPri, 0, LoadGasString(73529)
   Label(107).Caption = LoadGasString(73530)
   vgTooltips.AddTool txtCampo(97).CtPri, 0, LoadGasString(73531)
   vgTooltips.AddTool txtCampo(98).CtPri, 0, LoadGasString(73532)
   Label(108).Caption = LoadGasString(73533)
   vgTooltips.AddTool txtCampo(99).CtPri, 0, LoadGasString(73534)
   vgTooltips.AddTool txtCampo(100).CtPri, 0, LoadGasString(73535)
   vgTooltips.AddTool txtCampo(101).CtPri, 0, LoadGasString(73536)
   Label(109).Caption = LoadGasString(73537)
   Label(110).Caption = LoadGasString(73538)
   Label(111).Caption = LoadGasString(73539)
   Label(112).Caption = LoadGasString(73540)
   Label(113).Caption = LoadGasString(73541)
   vgTooltips.AddTool txtCampo(114).CtPri, 0, LoadGasString(73542)
   Label(114).Caption = LoadGasString(73543)
   Label(115).Caption = LoadGasString(73544)
   vgTooltips.AddTool txtCampo(115).CtPri, 0, LoadGasString(73545)
   Label(116).Caption = LoadGasString(73546)
   vgTooltips.AddTool txtCampo(116).CtPri, 0, LoadGasString(73547)
   Label(117).Caption = LoadGasString(73548)
   Label(118).Caption = LoadGasString(73549)
   vgTooltips.AddTool txtCampo(117).CtPri, 0, LoadGasString(73550)
   vgTooltips.AddTool txtCampo(118).CtPri, 0, LoadGasString(73551)
   Label(119).Caption = LoadGasString(73552)
   Label(120).Caption = LoadGasString(73553)
   vgTooltips.AddTool txtCampo(119).CtPri, 0, LoadGasString(73554)
   Label(121).Caption = LoadGasString(73555)
   vgTooltips.AddTool txtCampo(120).CtPri, 0, LoadGasString(73556)
   vgTooltips.AddTool txtCampo(121).CtPri, 0, LoadGasString(73557)
   Label(122).Caption = LoadGasString(73558)
   vgTooltips.AddTool txtCampo(122).CtPri, 0, LoadGasString(73559)
   Label(123).Caption = LoadGasString(73560)
   Label(124).Caption = LoadGasString(73561)
   vgTooltips.AddTool txtCampo(123).CtPri, 0, LoadGasString(73562)
   Label(125).Caption = LoadGasString(73563)
   vgTooltips.AddTool txtCampo(124).CtPri, 0, LoadGasString(73564)
   Label(126).Caption = LoadGasString(73565)
   Label(127).Caption = LoadGasString(73566)
   Label(128).Caption = LoadGasString(73567)
   Label(129).Caption = LoadGasString(73568)
   Label(130).Caption = LoadGasString(73569)
   Label(131).Caption = LoadGasString(73570)
   Label(132).Caption = LoadGasString(73571)
   Label(133).Caption = LoadGasString(73572)
   vgTooltips.AddTool txtCampo(126).CtPri, 0, LoadGasString(73573)
   Label(134).Caption = LoadGasString(73574)
   Label(135).Caption = LoadGasString(73575)
   vgTooltips.AddTool txtCampo(127).CtPri, 0, LoadGasString(73576)
   vgTooltips.AddTool txtCampo(128).CtPri, 0, LoadGasString(73577)
   Label(136).Caption = LoadGasString(73578)
   Label(137).Caption = LoadGasString(73579)
   vgTooltips.AddTool txtCampo(129).CtPri, 0, LoadGasString(73580)
   vgTooltips.AddTool txtCampo(130).CtPri, 0, LoadGasString(73581)
   Label(138).Caption = LoadGasString(73582)
   Label(139).Caption = LoadGasString(73583)
   vgTooltips.AddTool txtCampo(131).CtPri, 0, LoadGasString(73584)
   vgTooltips.AddTool txtCampo(132).CtPri, 0, LoadGasString(73585)
   Label(140).Caption = LoadGasString(73586)
   Label(141).Caption = LoadGasString(73587)
   Label(142).Caption = LoadGasString(73588)
   vgTooltips.AddTool txtCampo(133).CtPri, 0, LoadGasString(73589)
   Label(143).Caption = LoadGasString(73590)
   vgTooltips.AddTool txtCampo(134).CtPri, 0, LoadGasString(73591)
   Label(144).Caption = LoadGasString(73592)
   vgTooltips.AddTool txtCampo(135).CtPri, 0, LoadGasString(73593)
   Label(145).Caption = LoadGasString(73594)
   vgTooltips.AddTool txtCampo(136).CtPri, 0, LoadGasString(73595)
   Label(146).Caption = LoadGasString(73596)
   vgTooltips.AddTool txtCampo(137).CtPri, 0, LoadGasString(73597)
   Label(147).Caption = LoadGasString(73598)
   Label(148).Caption = LoadGasString(73599)
   Label(149).Caption = LoadGasString(73600)
   Label(150).Caption = LoadGasString(73601)
   Label(151).Caption = LoadGasString(73602)
   Label(152).Caption = LoadGasString(73603)
   vgTooltips.AddTool txtCampo(138).CtPri, 0, LoadGasString(73604)
   vgTooltips.AddTool txtCampo(139).CtPri, 0, LoadGasString(73605)
   vgTooltips.AddTool txtCampo(140).CtPri, 0, LoadGasString(73606)
   vgTooltips.AddTool txtCampo(141).CtPri, 0, LoadGasString(73607)
   vgTooltips.AddTool txtCampo(142).CtPri, 0, LoadGasString(73608)
   Label(153).Caption = LoadGasString(73609)
   Label(154).Caption = LoadGasString(73610)
   Label(155).Caption = LoadGasString(73611)
   Label(156).Caption = LoadGasString(73612)
   Label(157).Caption = LoadGasString(73613)
   Label(158).Caption = LoadGasString(73614)
   Label(159).Caption = LoadGasString(73615)
   Label(160).Caption = LoadGasString(73616)
   Label(161).Caption = LoadGasString(73617)
   Label(162).Caption = LoadGasString(73618)
   Label(163).Caption = LoadGasString(73619)
   Label(164).Caption = LoadGasString(73620)
   Label(165).Caption = LoadGasString(73621)
   Label(166).Caption = LoadGasString(73622)
   Label(167).Caption = LoadGasString(73623)
   Label(168).Caption = LoadGasString(73624)
   vgTooltips.AddTool txtCampo(143).CtPri, 0, LoadGasString(73625)
   vgTooltips.AddTool txtCampo(144).CtPri, 0, LoadGasString(73626)
   Label(169).Caption = LoadGasString(73627)
   vgTooltips.AddTool txtCampo(145).CtPri, 0, LoadGasString(73628)
   Label(170).Caption = LoadGasString(73629)
   vgTooltips.AddTool txtCampo(146).CtPri, 0, LoadGasString(73630)
   Label(171).Caption = LoadGasString(73631)
   Label(172).Caption = LoadGasString(73632)
   vgTooltips.AddTool txtCampo(147).CtPri, 0, LoadGasString(73633)
   Label(173).Caption = LoadGasString(73634)
   vgTooltips.AddTool txtCampo(148).CtPri, 0, LoadGasString(73635)
   Label(174).Caption = LoadGasString(73636)
   vgTooltips.AddTool txtCampo(149).CtPri, 0, LoadGasString(73637)
   Label(175).Caption = LoadGasString(73638)
   vgTooltips.AddTool txtCampo(150).CtPri, 0, LoadGasString(73639)
   Label(176).Caption = LoadGasString(73640)
   vgTooltips.AddTool txtCampo(151).CtPri, 0, LoadGasString(73641)
   Label(177).Caption = LoadGasString(73642)
   Label(178).Caption = LoadGasString(73643)
   Label(179).Caption = LoadGasString(73644)
   vgTooltips.AddTool Botao(8), 0, LoadGasString(73645)
   Botao(9).Caption = LoadGasString(73646)
   Label(180).Caption = LoadGasString(73647)
   Label(181).Caption = LoadGasString(73648)
   vgTooltips.AddTool Botao(10), 0, LoadGasString(73649)
   chkCampo(4).Caption = LoadGasString(73650)
   Label(182).Caption = LoadGasString(73651)
   Label(183).Caption = LoadGasString(73652)
   Label(184).Caption = LoadGasString(73653)
   Label(185).Caption = LoadGasString(73654)
   vgTooltips.AddTool Botao(12), 0, LoadGasString(73655)
   vgTooltips.AddTool chkCampo(5).CtPri, 0, LoadGasString(73656)
   chkCampo(5).Caption = LoadGasString(73657)
   Label(186).Caption = LoadGasString(73658)
   chkCampo(6).Caption = LoadGasString(73659)
   vgTooltips.AddTool chkCampo(7).CtPri, 0, LoadGasString(73660)
   chkCampo(7).Caption = LoadGasString(73661)
   vgTooltips.AddTool chkCampo(8).CtPri, 0, LoadGasString(73662)
   chkCampo(8).Caption = LoadGasString(73663)
   vgTooltips.AddTool txtCampo(159).CtPri, 0, LoadGasString(73664)
   vgTooltips.AddTool chkCampo(9).CtPri, 0, LoadGasString(73665)
   chkCampo(9).Caption = LoadGasString(73666)
   Botao(13).Caption = LoadGasString(73667)
   Botao(14).Caption = LoadGasString(73668)
   vgTooltips.AddTool Botao(15), 0, LoadGasString(73669)
   Label(187).Caption = LoadGasString(73670)
   vgTooltips.AddTool chkCampo(10).CtPri, 0, LoadGasString(73671)
   chkCampo(10).Caption = LoadGasString(73672)
   vgTooltips.AddTool chkCampo(11).CtPri, 0, LoadGasString(73673)
   chkCampo(11).Caption = LoadGasString(73674)
   Botao(16).Caption = LoadGasString(73675)
   Label(188).Caption = LoadGasString(73676)
   vgTooltips.AddTool txtCampo(161).CtPri, 0, LoadGasString(73677)
   With Grid(0)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = Permitido("Conjuntos do Oramento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Conjuntos do Oramento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Conjuntos do Oramento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Conjunto", "Seqncia do Conjunto", TP_NUMERICO, "", , False, , "IRRIGACAO", "Conjuntos", "Seqncia do Conjunto", "Seqncia do Conjunto; Descrio", "Seqncia do Conjunto; Descrio", "Seqncia do Conjunto; Descrio", "", , "1", "Conjuntos.[Seqncia do Conjunto]", "", "IRRIGACAO", "18", 2, "0", 4695
      .AddColumn Nothing, , "CST", "CST", TP_NUMERICO, "999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 435
      .AddColumn Nothing, , "CFOP", "CFOP", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 615
      .AddColumn Nothing, , "Un", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 615
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1095
      .AddColumn Nothing, , "Estoque", , TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1095
      .AddColumn Nothing, , "Vr. Unitrio", "Valor Unitrio", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Desconto", "Valor do Desconto", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Frete", "Valor do Frete", TP_NUMERICO, "9.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "Vr. Total", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "B. Clc ICMS", "Valor da Base de Clculo", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. ICMS", "Valor do ICMS", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. IPI", "Valor do IPI", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "% ICMS", "Alquota do ICMS", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "% IPI", "Alquota do IPI", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "D.", "Diferido", TP_LOGICO, , , False, , , , , , , , , , "0", , , , "0", 1, "0", 330
      .AddColumn Nothing, , "% Reduo", "Percentual da Reduo", TP_NUMERICO, "999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1005
      .AddColumn Nothing, , "IVA Ajustado", "IVA", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "B. Clculo ICMS ST", "Base de Clculo ST", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1575
      .AddColumn Nothing, , "Vr. ICMS ST", "Valor ICMS ST", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1080
      .AddColumn Nothing, , "% ICMS ST", "Alquota do ICMS ST", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 990
      .AddColumn Nothing, , "BC. Pis", "Bc pis", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1110
      .AddColumn Nothing, , "% Pis", "Aliq do pis", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 855
      .AddColumn Nothing, , "Vr. PIS", "Valor do PIS", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "BC. Cofins", "Bc cofins", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1065
      .AddColumn Nothing, , "% Cofins", "Aliq do cofins", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "Vr. COFINS", "Valor do Cofins", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Vr. Tributo", "Valor do Tributo", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
   End With
   With Grid(1)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = Permitido("Peas do Oramento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Peas do Oramento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Peas do Oramento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Pea", "Seqncia do Produto", TP_NUMERICO, "", , False, , "IRRIGACAO", "Produtos", "Seqncia do Produto", "Seqncia do Produto; Descrio", "Seqncia do Produto; Descrio", "Seqncia do Produto; Descrio", "", , "1", "Produtos.[Seqncia do Produto]", "", "IRRIGACAO", "18", 2, "0", 4695
      .AddColumn Nothing, , "CST", "CST", TP_NUMERICO, "999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 435
      .AddColumn Nothing, , "CFOP", "CFOP", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 720
      .AddColumn Nothing, , "Un", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 615
      .AddColumn Nothing, , "Peso", , TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 945
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Estoque", , TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "Peso Total", , TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Vr. Unitrio", "Valor Unitrio", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Desconto", "Valor do Desconto", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1125
      .AddColumn Nothing, , "Valor do Frete", "Valor do Frete", TP_NUMERICO, "9.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1350
      .AddColumn Nothing, , "Vr. Total", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "B. Clc ICMS", "Valor da Base de Clculo", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. ICMS", "Valor do ICMS", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. IPI", "Valor do IPI", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "% ICMS", "Alquota do ICMS", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "% IPI", "Alquota do IPI", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "D.", "Diferido", TP_LOGICO, , , False, , , , , , , , , , "0", , , , "0", 1, "0", 330
      .AddColumn Nothing, , "% Reduo", "Percentual da Reduo", TP_NUMERICO, "999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1110
      .AddColumn Nothing, , "IVA Ajustado", "IVA", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1125
      .AddColumn Nothing, , "B. Clculo ICMS ST", "Base de Clculo ST", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1800
      .AddColumn Nothing, , "Vr. ICMS ST", "Valor ICMS ST", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "% ICMS ST", "Alquota do ICMS ST", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 990
      .AddColumn Nothing, , "BC. Pis", "Bc pis", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "% Pis", "Aliq do pis", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 870
      .AddColumn Nothing, , "Vr. PIS", "Valor do PIS", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1155
      .AddColumn Nothing, , "BC. Cofins", "Bc cofins", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1110
      .AddColumn Nothing, , "% Cofins", "Aliq do cofins", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 900
      .AddColumn Nothing, , "Vr. COFINS", "Valor do Cofins", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1185
      .AddColumn Nothing, , "Vr. Tributo", "Valor do Tributo", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
   End With
   With Grid(2)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = Permitido("Parcelas Oramento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Parcelas Oramento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas Oramento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "N Pc.", "Nmero da Parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 585
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 585
      .AddColumn Nothing, , "Vencimento", "Data de Vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Valor", "Valor da Parcela", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1260
      .AddColumn Nothing, , "Cobrana", "Descrio da Cobrana", TP_CARACTER, "@x", , False, , "IRRIGACAO", "Tipo de Cobrana", "Descrio", "Descrio", "Descrio", "Descrio", "", , "1", , , , "0", 1, "0", 1485
      .AddColumn Nothing, , "Observao", "Descrio", TP_CARACTER, "@x", 120, False, , , , , , , , , , "0", , , , "0", 1, "0", 1860
   End With
   With Grid(3)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = Permitido("Produtos do Oramento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Produtos do Oramento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos do Oramento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Seqncia do Produto", TP_NUMERICO, "", , False, , "IRRIGACAO", "Produtos", "Seqncia do Produto", "Seqncia do Produto; Descrio", "Seqncia do Produto; Descrio", "Seqncia do Produto; Descrio", "", , "1", "Produtos.[Seqncia do Produto]", "", "IRRIGACAO", "18", 2, "0", 4695
      .AddColumn Nothing, , "NCM", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1035
      .AddColumn Nothing, , "CST", "CST", TP_NUMERICO, "999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 435
      .AddColumn Nothing, , "CFOP", "CFOP", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 720
      .AddColumn Nothing, , "Un", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 615
      .AddColumn Nothing, , "Peso", , TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 945
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1095
      .AddColumn Nothing, , "Estoque", , TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1080
      .AddColumn Nothing, , "Peso Total", , TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Vr. Unitrio", "Valor Unitrio", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1185
      .AddColumn Nothing, , "Desconto", "Valor do Desconto", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Frete", "Valor do Frete", TP_NUMERICO, "9.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "Vr. Total", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "B. Clc ICMS", "Valor da Base de Clculo", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. ICMS", "Valor do ICMS", TP_NUMERICO, "9.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1185
      .AddColumn Nothing, , "Vr. IPI", "Valor do IPI", TP_NUMERICO, "9.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "% ICMS", "Alquota do ICMS", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "% IPI", "Alquota do IPI", TP_NUMERICO, "999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "D.", "Diferido", TP_LOGICO, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 330
      .AddColumn Nothing, , "% Reduo", "Percentual da Reduo", TP_NUMERICO, "999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "IVA Ajustado", "IVA", TP_NUMERICO, "999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1185
      .AddColumn Nothing, , "B. Clculo ICMS ST", "Base de Clculo ST", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1545
      .AddColumn Nothing, , "Vr. ICMS ST", "Valor ICMS ST", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "% ICMS ST", "Alquota do ICMS ST", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 990
      .AddColumn Nothing, , "BC. Pis", "Bc pis", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "% Pis", "Aliq do pis", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 945
      .AddColumn Nothing, , "Vr. PIS", "Valor do PIS", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "BC. Cofins", "Bc cofins", TP_NUMERICO, "999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1290
      .AddColumn Nothing, , "% Cofins", "Aliq do cofins", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 855
      .AddColumn Nothing, , "Vr. COFINS", "Valor do Cofins", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. Tributo", "Valor do Tributo", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
   End With
   With Grid(4)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = Permitido("Servios do Oramento", ACAO_INCLUINDO)
      .AllowEdit = Permitido("Servios do Oramento", ACAO_EDITANDO)
      .AllowDelete = Permitido("Servios do Oramento", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Servio", "Seqncia do Servio", TP_NUMERICO, "", , False, , "IRRIGACAO", "Servios", "Seqncia do Servio", "Seqncia do Servio; Descrio", "Seqncia do Servio; Descrio", "Seqncia do Servio; Descrio", "", , "1", "Servios.[Seqncia do Servio]", "", "IRRIGACAO", "18", 2, "0", 4695
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 945
      .AddColumn Nothing, , "Vr. Unitrio", "Valor Unitrio", TP_NUMERICO, "9.999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Vr. Total", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Valor do ISS", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Orcamento = vgTb
   vgPriVez = False
   Reposition
   CarregaTotalizador
   AtualizaCamposFrete
   Screen.MousePointer = vbDefault
   Exit Sub

DeuErro:
   CErr.NumErro = Err
   CErr.FunctionName = "IniciaForm"
   CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
   CErr.Show
End Sub



Private Function CodigoFrete() As Integer
    Dim idx As Long
    idx = txtCampo(151).ListIndex

    Select Case idx
        Case 0: CodigoFrete = 0
        Case 1: CodigoFrete = 1
        Case 2: CodigoFrete = 3
        Case 3: CodigoFrete = 4
        Case Else
            Select Case Trim$(txtCampo(151).Text)
                Case "0", "Emitente":                     CodigoFrete = 0
                Case "1", "Destinatrio":                 CodigoFrete = 1
                Case "3", "Transporte Prprio Remetente": CodigoFrete = 3
                Case "4", "Transporte Prprio Destinatrio": CodigoFrete = 4
                Case Else:                                CodigoFrete = 0
            End Select
    End Select
End Function




'----------------------------------------------------------------------------
' 3) Rotina que bloqueia/libera campos de placa/transportadora
'----------------------------------------------------------------------------
Private Sub LimpaTransportadora()
    txtCp(55).Text = ""
    txtCampo(55).Text = ""
    Sequencia_da_Transportadora = 0
End Sub



'----------------------------------------------------------------------------
' 2) Evento  disparado quando qualquer txtCp(*) muda
'----------------------------------------------------------------------------
'evento  disparado quando o contedo de qualquer txtCp(*) for alterado
Private Sub txtCp_Change(Index As Integer)

   ' Evita disparar na 1 vez ou enquanto o prprio txtCampo ainda est em PriVez
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub

   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX
   txtCampo(Index).Change          ' propaga a mudana ao txtCampo vinculado

   ' Campos que pedem ExecutaVisivel / PreValidacao / MostraFormulas
   If Index = 17 Or Index = 33 Or Index = 34 Or Index = 35 Or _
      Index = 55 Or Index = 59 Or Index = 60 Or Index = 62 Or _
      Index = 64 Or Index = 99 Or Index = 100 Or Index = 101 Or _
      Index = 137 Or Index = 147 Or Index = 152 Or Index = 153 Or _
      Index = 154 Or Index = 156 Or Index = 157 Or Index = 161 Then

      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If

   ' Distribui desconto total (campo 32)
   If Index = 32 Then Call DistribuiDescontoTotal

   ' Se o modo de frete (campo 151) mudou ? bloqueia/libera instantaneamente
   If Index = 151 Then Call AtualizaCamposFrete

   ' Rotinas especficas por ndice
   Select Case Index
      Case 17
         RepositionOrcamento
      Case 156
RepositionOrcamento:          LimpaProp
      Case 157
         RepositionOrcamento
   End Select
End Sub




'----------------------------------------------------------------------------
' 3) Rotina que bloqueia/libera campos de placa/transportadora
'----------------------------------------------------------------------------
Private Sub AtualizaCamposFrete()

    Dim modFrete As Integer
    modFrete = CodigoFrete()

    Dim bloquear As Boolean
    bloquear = (modFrete = 3 Or modFrete = 4)
    If bloquear Then Call LimpaTransportadora

    'Habilita / desabilita
    txtCp(55).Enabled = Not bloquear
    Botao(4).Enabled = Not bloquear
    bottxtCampo55(1).Enabled = Not bloquear
    bottxtCampo55(2).Enabled = Not bloquear

    'Feedback visual
    Const corCinza As Long = &H8000000F   'vbButtonFace
    Dim corNormal As Long: corNormal = vbWindowBackground

    txtCp(55).BackColor = IIf(bloquear, corCinza, corNormal)
    bottxtCampo55(1).BackColor = IIf(bloquear, corCinza, corNormal)
    bottxtCampo55(2).BackColor = IIf(bloquear, corCinza, corNormal)

End Sub







Public Sub DefineControles()
 On Error GoTo DeuErro
 grdBrowse.AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'no deixa o grid tentar gravar automaticamente
 grdBrowse.AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botes
   grdBrowse.AllowDelete = True
   grdBrowse.AllowEdit = vgAlterar
   grdBrowse.SpecialPopupDisabled POP_GRID_BARS

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "999999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set txtCampo(159).CtPri = txtCp(159)
   txtCampo(159).DataType = 1
   txtCampo(159).Mask = "999999"
   txtCampo(159).Editable = False
   txtCampo(159).BoundColumn = ""
   txtCampo(159).ListFields = ""
   txtCampo(159).OrderFields = ""
   txtCampo(159).Relation = ""
   txtCampo(159).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(159).DataField), txtCampo(159)

   Set txtCampo(158).CtPri = txtCp(158)
   Set txtCampo(158).CtFdo = labFdo158
   Set txtCampo(158).CtBot(BOT_ACAO) = bottxtCampo158(BOT_ACAO)
   Set bottxtCampo158(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(158).DataType = 2
   txtCampo(158).Mask = "99/99/9999"
   txtCampo(158).BoundColumn = ""
   txtCampo(158).ListFields = ""
   txtCampo(158).OrderFields = ""
   txtCampo(158).Relation = ""
   txtCampo(158).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(158).DataField), txtCampo(158)

   Set txtCampo(156).CtPri = txtCp(156)
   Set txtCampo(156).CtFdo = labFdo156
   Set txtCampo(156).CtBot(BOT_LISTA) = bottxtCampo156(BOT_LISTA)
   Set txtCampo(156).CtBot(BOT_COMBO) = bottxtCampo156(BOT_COMBO)
   bottxtCampo156(BOT_LISTA).Caption = "P"
   Set bottxtCampo156(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(156).DataType = 1
   txtCampo(156).Mask = "999999"
   txtCampo(156).PesqModoAbertura = 2
   txtCampo(156).PesqFieldCapture = "Geral.[Seqncia do Geral]"
   txtCampo(156).BoundColumn = "Seqncia do Geral"
   txtCampo(156).ListFields = "Razo Social; Seqncia do Geral"
   txtCampo(156).OrderFields = "Razo Social; Seqncia do Geral"
   txtCampo(156).Relation = ""
   txtCampo(156).Source = "Geral"
   txtCampo(156).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(156).DataField), txtCampo(156)

   Set txtCampo(157).CtPri = txtCp(157)
   Set txtCampo(157).CtFdo = labFdo157
   Set txtCampo(157).CtBot(BOT_LISTA) = bottxtCampo157(BOT_LISTA)
   Set txtCampo(157).CtBot(BOT_COMBO) = bottxtCampo157(BOT_COMBO)
   bottxtCampo157(BOT_LISTA).Caption = "P"
   Set bottxtCampo157(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(157).DataType = 1
   txtCampo(157).Mask = "9999"
   txtCampo(157).PesqModoAbertura = 2
   txtCampo(157).BoundColumn = "Seqncia da Propriedade"
   txtCampo(157).ListFields = "Nome da Propriedade; Seqncia da Propriedade"
   txtCampo(157).OrderFields = "Nome da Propriedade; Seqncia da Propriedade"
   txtCampo(157).Relation = ""
   txtCampo(157).Source = "Propriedades"
   txtCampo(157).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(157).DataField), txtCampo(157)

   Set txtCampo(18).CtPri = txtCp(18)
   txtCampo(18).DataType = 0
   txtCampo(18).Mask = "@x"
   txtCampo(18).BoundColumn = ""
   txtCampo(18).ListFields = ""
   txtCampo(18).OrderFields = ""
   txtCampo(18).Relation = ""
   txtCampo(18).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(18).DataField), txtCampo(18)

   Set txtCampo(48).CtPri = txtCp(48)
   txtCampo(48).DataType = 0
   txtCampo(48).Mask = "@x"
   txtCampo(48).BoundColumn = ""
   txtCampo(48).ListFields = ""
   txtCampo(48).OrderFields = ""
   txtCampo(48).Relation = ""
   txtCampo(48).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(48).DataField), txtCampo(48)

   Set chkCampo(1).CtPri = ChkCp(1)
   chkCampo(1).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(1).DataField), chkCampo(1)

   Set txtCampo(24).CtPri = txtCp(24)
   txtCampo(24).DataType = 0
   txtCampo(24).Mask = "@x"
   txtCampo(24).BoundColumn = ""
   txtCampo(24).ListFields = ""
   txtCampo(24).OrderFields = ""
   txtCampo(24).Relation = ""
   txtCampo(24).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(24).DataField), txtCampo(24)

   Set txtCampo(25).CtPri = txtCp(25)
   txtCampo(25).DataType = 0
   txtCampo(25).Mask = "@x"
   txtCampo(25).BoundColumn = ""
   txtCampo(25).ListFields = ""
   txtCampo(25).OrderFields = ""
   txtCampo(25).Relation = ""
   txtCampo(25).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(25).DataField), txtCampo(25)

   Set txtCampo(20).CtPri = txtCp(20)
   txtCampo(20).DataType = 0
   txtCampo(20).Mask = "@x"
   txtCampo(20).BoundColumn = ""
   txtCampo(20).ListFields = ""
   txtCampo(20).OrderFields = ""
   txtCampo(20).Relation = ""
   txtCampo(20).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(20).DataField), txtCampo(20)

   Set txtCampo(23).CtPri = txtCp(23)
   txtCampo(23).DataType = 0
   txtCampo(23).Mask = "@x"
   txtCampo(23).BoundColumn = ""
   txtCampo(23).ListFields = ""
   txtCampo(23).OrderFields = ""
   txtCampo(23).Relation = ""
   txtCampo(23).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(23).DataField), txtCampo(23)

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "99999-999"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(22).CtPri = txtCp(22)
   txtCampo(22).DataType = 0
   txtCampo(22).Mask = "(99)##999-9999"
   txtCampo(22).BoundColumn = ""
   txtCampo(22).ListFields = ""
   txtCampo(22).OrderFields = ""
   txtCampo(22).Relation = ""
   txtCampo(22).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(22).DataField), txtCampo(22)

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 0
   txtCampo(12).Mask = "(99)##999-9999"
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(21).CtPri = txtCp(21)
   txtCampo(21).DataType = 0
   txtCampo(21).Mask = "@x"
   txtCampo(21).BoundColumn = ""
   txtCampo(21).ListFields = ""
   txtCampo(21).OrderFields = ""
   txtCampo(21).Relation = ""
   txtCampo(21).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(21).DataField), txtCampo(21)

   Set txtCampo(17).CtPri = txtCp(17)
   Set txtCampo(17).CtFdo = labFdo17
   Set txtCampo(17).CtBot(BOT_LISTA) = bottxtCampo17(BOT_LISTA)
   Set txtCampo(17).CtBot(BOT_COMBO) = bottxtCampo17(BOT_COMBO)
   bottxtCampo17(BOT_LISTA).Caption = "P"
   Set bottxtCampo17(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(17).DataType = 1
   txtCampo(17).Mask = "99999"
   txtCampo(17).PesqModoAbertura = 2
   txtCampo(17).PesqFieldCapture = "Municpios.[Seqncia do Municpio]"
   txtCampo(17).BoundColumn = "Seqncia do Municpio"
   txtCampo(17).ListFields = "Descrio; Seqncia do Municpio"
   txtCampo(17).OrderFields = "Descrio; Seqncia do Municpio"
   txtCampo(17).Relation = ""
   txtCampo(17).Source = "Municpios"
   txtCampo(17).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(17).DataField), txtCampo(17)

   Set txtCampo(27).CtPri = txtCp(27)
   txtCampo(27).DataType = 0
   txtCampo(27).Mask = ""
   txtCampo(27).Editable = False
   txtCampo(27).BoundColumn = ""
   txtCampo(27).ListFields = ""
   txtCampo(27).OrderFields = ""
   txtCampo(27).Relation = ""
   txtCampo(27).Source = ""

   Set opcPainel1(0).CtPri = opcPainel1Cp(0)
   Set opcPainel1(0).CtFdo = labopcPainel1
   opcPainel1(0).DataType = 6
   opcPainel1(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcPainel1(0).DataField), opcPainel1(0)

   Set opcPainel1(1).CtPri = opcPainel1Cp(1)
   Set opcPainel1(1).CtFdo = labopcPainel1
   opcPainel1(1).DataType = 6
   opcPainel1(1).BookMark = 1

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "@x"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(14).CtPri = txtCp(14)
   txtCampo(14).DataType = 0
   txtCampo(14).Mask = "@x"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(51).CtPri = txtCp(51)
   txtCampo(51).DataType = 0
   txtCampo(51).Mask = "@x"
   txtCampo(51).BoundColumn = ""
   txtCampo(51).ListFields = ""
   txtCampo(51).OrderFields = ""
   txtCampo(51).Relation = ""
   txtCampo(51).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(51).DataField), txtCampo(51)

   Set txtCampo(16).CtPri = txtCp(16)
   Set txtCampo(16).CtFdo = labFdo16
   Set txtCampo(16).CtBot(BOT_ACAO) = bottxtCampo16(BOT_ACAO)
   Set bottxtCampo16(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(16).DataType = 0
   txtCampo(16).Mask = "@a"
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

   Set txtCampo(33).CtPri = txtCp(33)
   Set txtCampo(33).CtFdo = labFdo33
   Set txtCampo(33).CtBot(BOT_LISTA) = bottxtCampo33(BOT_LISTA)
   Set txtCampo(33).CtBot(BOT_COMBO) = bottxtCampo33(BOT_COMBO)
   bottxtCampo33(BOT_LISTA).Caption = "P"
   Set bottxtCampo33(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(33).DataType = 1
   txtCampo(33).Mask = "999999"
   txtCampo(33).PesqModoAbertura = 2
   txtCampo(33).PesqFieldCapture = "Geral.[Seqncia do Geral]"
   txtCampo(33).BoundColumn = "Seqncia do Geral"
   txtCampo(33).ListFields = "Razo Social; Seqncia do Geral"
   txtCampo(33).OrderFields = "Razo Social; Seqncia do Geral"
   txtCampo(33).Relation = ""
   txtCampo(33).Source = "Geral"
   txtCampo(33).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(33).DataField), txtCampo(33)

   Set txtCampo(55).CtPri = txtCp(55)
   Set txtCampo(55).CtFdo = labFdo55
   Set txtCampo(55).CtBot(BOT_LISTA) = bottxtCampo55(BOT_LISTA)
   Set txtCampo(55).CtBot(BOT_COMBO) = bottxtCampo55(BOT_COMBO)
   bottxtCampo55(BOT_LISTA).Caption = "P"
   Set bottxtCampo55(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(55).DataType = 1
   txtCampo(55).Mask = "999999"
   txtCampo(55).PesqModoAbertura = 2
   txtCampo(55).PesqFieldCapture = "Geral.[Seqncia do Geral]"
   txtCampo(55).BoundColumn = "Seqncia do Geral"
   txtCampo(55).ListFields = "Razo Social; Seqncia do Geral"
   txtCampo(55).OrderFields = "Razo Social; Seqncia do Geral"
   txtCampo(55).Relation = ""
   txtCampo(55).Source = "Geral"
   txtCampo(55).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(55).DataField), txtCampo(55)

   Set txtCampo(58).CtPri = txtCp(58)
   txtCampo(58).DataType = 0
   txtCampo(58).Mask = "@x"
   txtCampo(58).BoundColumn = ""
   txtCampo(58).ListFields = ""
   txtCampo(58).OrderFields = ""
   txtCampo(58).Relation = ""
   txtCampo(58).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(58).DataField), txtCampo(58)

   Set txtCampo(59).CtPri = txtCp(59)
   Set txtCampo(59).CtFdo = labFdo59
   Set txtCampo(59).CtBot(BOT_COMBO) = bottxtCampo59(BOT_COMBO)
   Set bottxtCampo59(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(59).DataType = 0
   txtCampo(59).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE|EX"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(59).DataField), txtCampo(59)

   Set txtCampo(54).CtPri = txtCp(54)
   txtCampo(54).DataType = 1
   txtCampo(54).Mask = "99,99"
   txtCampo(54).BoundColumn = ""
   txtCampo(54).ListFields = ""
   txtCampo(54).OrderFields = ""
   txtCampo(54).Relation = ""
   txtCampo(54).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(54).DataField), txtCampo(54)

   Set chkCampo(2).CtPri = ChkCp(2)
   chkCampo(2).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(2).DataField), chkCampo(2)

   Set txtCampo(50).CtPri = txtCp(50)
   txtCampo(50).DataType = 2
   txtCampo(50).Mask = "99/99/9999"
   txtCampo(50).Editable = False
   txtCampo(50).BoundColumn = ""
   txtCampo(50).ListFields = ""
   txtCampo(50).OrderFields = ""
   txtCampo(50).Relation = ""
   txtCampo(50).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(50).DataField), txtCampo(50)

   Set txtCampo(60).CtPri = txtCp(60)
   Set txtCampo(60).CtFdo = labFdo60
   Set txtCampo(60).CtBot(BOT_LISTA) = bottxtCampo60(BOT_LISTA)
   Set txtCampo(60).CtBot(BOT_COMBO) = bottxtCampo60(BOT_COMBO)
   bottxtCampo60(BOT_LISTA).Caption = "P"
   Set bottxtCampo60(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(60).DataType = 1
   txtCampo(60).Mask = "999999"
   txtCampo(60).PesqModoAbertura = 2
   txtCampo(60).PesqFieldCapture = "Pases.[Seqncia do Pas]"
   txtCampo(60).BoundColumn = "Seqncia do Pas"
   txtCampo(60).ListFields = "Descrio; Seqncia do Pas"
   txtCampo(60).OrderFields = "Descrio; Seqncia do Pas"
   txtCampo(60).Relation = ""
   txtCampo(60).Source = "Pases"
   txtCampo(60).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(60).DataField), txtCampo(60)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 4
   txtCampo(4).Mask = ""
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(34).CtPri = txtCp(34)
   Set txtCampo(34).CtFdo = labFdo34
   Set txtCampo(34).CtBot(BOT_LISTA) = bottxtCampo34(BOT_LISTA)
   Set txtCampo(34).CtBot(BOT_COMBO) = bottxtCampo34(BOT_COMBO)
   bottxtCampo34(BOT_LISTA).Caption = "P"
   Set bottxtCampo34(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(34).DataType = 1
   txtCampo(34).Mask = "9999"
   txtCampo(34).PesqModoAbertura = 2
   txtCampo(34).PesqFieldCapture = "[Classificao Fiscal].[Seqncia da Classificao]"
   txtCampo(34).BoundColumn = "Seqncia da Classificao"
   txtCampo(34).ListFields = "NCM"
   txtCampo(34).OrderFields = "NCM"
   txtCampo(34).Relation = ""
   txtCampo(34).Source = "Classificao Fiscal"
   txtCampo(34).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(34).DataField), txtCampo(34)

   Set txtCampo(35).CtPri = txtCp(35)
   Set txtCampo(35).CtFdo = labFdo35
   Set txtCampo(35).CtBot(BOT_COMBO) = bottxtCampo35(BOT_COMBO)
   Set bottxtCampo35(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(35).DataType = 0
   txtCampo(35).ListFields = "Vista|Prazo|Antecipado"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(35).DataField), txtCampo(35)

   Set opcPainel2(0).CtPri = opcPainel2Cp(0)
   Set opcPainel2(0).CtFdo = labopcPainel2
   opcPainel2(0).DataType = 6
   opcPainel2(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcPainel2(0).DataField), opcPainel2(0)

   Set opcPainel2(1).CtPri = opcPainel2Cp(1)
   Set opcPainel2(1).CtFdo = labopcPainel2
   opcPainel2(1).DataType = 6
   opcPainel2(1).BookMark = 1

   Set txtCampo(32).CtPri = txtCp(32)
   Set txtCampo(32).CtFdo = labFdo32
   Set txtCampo(32).CtBot(BOT_ACAO) = bottxtCampo32(BOT_ACAO)
   Set bottxtCampo32(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(32).DataType = 1
   txtCampo(32).Mask = "99.999.999,99"
   txtCampo(32).BoundColumn = ""
   txtCampo(32).ListFields = ""
   txtCampo(32).OrderFields = ""
   txtCampo(32).Relation = ""
   txtCampo(32).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(32).DataField), txtCampo(32)

   Set txtCampo(45).CtPri = txtCp(45)
   Set txtCampo(45).CtFdo = labFdo45
   Set txtCampo(45).CtBot(BOT_ACAO) = bottxtCampo45(BOT_ACAO)
   Set bottxtCampo45(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(45).DataType = 1
   txtCampo(45).Mask = "9.999.999,9999"
   txtCampo(45).BoundColumn = ""
   txtCampo(45).ListFields = ""
   txtCampo(45).OrderFields = ""
   txtCampo(45).Relation = ""
   txtCampo(45).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(45).DataField), txtCampo(45)

   Set txtCampo(47).CtPri = txtCp(47)
   Set txtCampo(47).CtFdo = labFdo47
   Set txtCampo(47).CtBot(BOT_ACAO) = bottxtCampo47(BOT_ACAO)
   Set bottxtCampo47(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(47).DataType = 1
   txtCampo(47).Mask = "99.999.999,99"
   txtCampo(47).BoundColumn = ""
   txtCampo(47).ListFields = ""
   txtCampo(47).OrderFields = ""
   txtCampo(47).Relation = ""
   txtCampo(47).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(47).DataField), txtCampo(47)

 DefineControles1

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub

Public Sub DefineControles1()
 On Error GoTo DeuErro

   Set txtCampo(148).CtPri = txtCp(148)
   txtCampo(148).DataType = 1
   txtCampo(148).Mask = "9.999.999,99"
   txtCampo(148).BoundColumn = ""
   txtCampo(148).ListFields = ""
   txtCampo(148).OrderFields = ""
   txtCampo(148).Relation = ""
   txtCampo(148).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(148).DataField), txtCampo(148)

   Set txtCampo(57).CtPri = txtCp(57)
   Set txtCampo(57).CtFdo = labFdo57
   Set txtCampo(57).CtBot(BOT_ACAO) = bottxtCampo57(BOT_ACAO)
   Set bottxtCampo57(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(57).DataType = 1
   txtCampo(57).Mask = "99.999.999,99"
   txtCampo(57).BoundColumn = ""
   txtCampo(57).ListFields = ""
   txtCampo(57).OrderFields = ""
   txtCampo(57).Relation = ""
   txtCampo(57).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(57).DataField), txtCampo(57)

   Set txtCampo(155).CtPri = txtCp(155)
   txtCampo(155).DataType = 0
   txtCampo(155).Mask = ""
   txtCampo(155).Editable = False
   txtCampo(155).BoundColumn = ""
   txtCampo(155).ListFields = ""
   txtCampo(155).OrderFields = ""
   txtCampo(155).Relation = ""
   txtCampo(155).Source = ""

   Set chkCampo(4).CtPri = ChkCp(4)
   chkCampo(4).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(4).DataField), chkCampo(4)

   Set txtCampo(31).CtPri = txtCp(31)
   txtCampo(31).DataType = 1
   txtCampo(31).Mask = "99.999.999,99"
   txtCampo(31).Editable = False
   txtCampo(31).BoundColumn = ""
   txtCampo(31).ListFields = ""
   txtCampo(31).OrderFields = ""
   txtCampo(31).Relation = ""
   txtCampo(31).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(31).DataField), txtCampo(31)

   Set txtCampo(36).CtPri = txtCp(36)
   txtCampo(36).DataType = 1
   txtCampo(36).Mask = "99.999.999,99"
   txtCampo(36).Editable = False
   txtCampo(36).BoundColumn = ""
   txtCampo(36).ListFields = ""
   txtCampo(36).OrderFields = ""
   txtCampo(36).Relation = ""
   txtCampo(36).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(36).DataField), txtCampo(36)

   Set txtCampo(37).CtPri = txtCp(37)
   txtCampo(37).DataType = 1
   txtCampo(37).Mask = "99.999.999,99"
   txtCampo(37).Editable = False
   txtCampo(37).BoundColumn = ""
   txtCampo(37).ListFields = ""
   txtCampo(37).OrderFields = ""
   txtCampo(37).Relation = ""
   txtCampo(37).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(37).DataField), txtCampo(37)

   Set txtCampo(38).CtPri = txtCp(38)
   txtCampo(38).DataType = 1
   txtCampo(38).Mask = "99.999.999,99"
   txtCampo(38).Editable = False
   txtCampo(38).BoundColumn = ""
   txtCampo(38).ListFields = ""
   txtCampo(38).OrderFields = ""
   txtCampo(38).Relation = ""
   txtCampo(38).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(38).DataField), txtCampo(38)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "99.999.999,99"
   txtCampo(9).Editable = False
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(39).CtPri = txtCp(39)
   txtCampo(39).DataType = 1
   txtCampo(39).Mask = "99.999.999,99"
   txtCampo(39).Editable = False
   txtCampo(39).BoundColumn = ""
   txtCampo(39).ListFields = ""
   txtCampo(39).OrderFields = ""
   txtCampo(39).Relation = ""
   txtCampo(39).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(39).DataField), txtCampo(39)

   Set txtCampo(40).CtPri = txtCp(40)
   txtCampo(40).DataType = 1
   txtCampo(40).Mask = "99.999.999,99"
   txtCampo(40).Editable = False
   txtCampo(40).BoundColumn = ""
   txtCampo(40).ListFields = ""
   txtCampo(40).OrderFields = ""
   txtCampo(40).Relation = ""
   txtCampo(40).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(40).DataField), txtCampo(40)

   Set txtCampo(41).CtPri = txtCp(41)
   txtCampo(41).DataType = 1
   txtCampo(41).Mask = "99.999.999,99"
   txtCampo(41).Editable = False
   txtCampo(41).BoundColumn = ""
   txtCampo(41).ListFields = ""
   txtCampo(41).OrderFields = ""
   txtCampo(41).Relation = ""
   txtCampo(41).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(41).DataField), txtCampo(41)

   Set txtCampo(42).CtPri = txtCp(42)
   txtCampo(42).DataType = 1
   txtCampo(42).Mask = "99.999.999,99"
   txtCampo(42).Editable = False
   txtCampo(42).BoundColumn = ""
   txtCampo(42).ListFields = ""
   txtCampo(42).OrderFields = ""
   txtCampo(42).Relation = ""
   txtCampo(42).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(42).DataField), txtCampo(42)

   Set txtCampo(43).CtPri = txtCp(43)
   txtCampo(43).DataType = 1
   txtCampo(43).Mask = "99.999.999,99"
   txtCampo(43).Editable = False
   txtCampo(43).BoundColumn = ""
   txtCampo(43).ListFields = ""
   txtCampo(43).OrderFields = ""
   txtCampo(43).Relation = ""
   txtCampo(43).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(43).DataField), txtCampo(43)

   Set txtCampo(44).CtPri = txtCp(44)
   txtCampo(44).DataType = 1
   txtCampo(44).Mask = "99.999.999,99"
   txtCampo(44).Editable = False
   txtCampo(44).BoundColumn = ""
   txtCampo(44).ListFields = ""
   txtCampo(44).OrderFields = ""
   txtCampo(44).Relation = ""
   txtCampo(44).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(44).DataField), txtCampo(44)

   Set txtCampo(46).CtPri = txtCp(46)
   txtCampo(46).DataType = 1
   txtCampo(46).Mask = "99.999.999,99"
   txtCampo(46).Editable = False
   txtCampo(46).BoundColumn = ""
   txtCampo(46).ListFields = ""
   txtCampo(46).OrderFields = ""
   txtCampo(46).Relation = ""
   txtCampo(46).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(46).DataField), txtCampo(46)

   Set txtCampo(49).CtPri = txtCp(49)
   txtCampo(49).DataType = 1
   txtCampo(49).Mask = "99.999.999,99"
   txtCampo(49).Editable = False
   txtCampo(49).BoundColumn = ""
   txtCampo(49).ListFields = ""
   txtCampo(49).OrderFields = ""
   txtCampo(49).Relation = ""
   txtCampo(49).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(49).DataField), txtCampo(49)

   Set txtCampo(28).CtPri = txtCp(28)
   txtCampo(28).DataType = 0
   txtCampo(28).Mask = ""
   txtCampo(28).Editable = False
   txtCampo(28).BoundColumn = ""
   txtCampo(28).ListFields = ""
   txtCampo(28).OrderFields = ""
   txtCampo(28).Relation = ""
   txtCampo(28).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = ""
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 0
   txtCampo(2).Mask = ""
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""

   Set txtCampo(19).CtPri = txtCp(19)
   txtCampo(19).DataType = 0
   txtCampo(19).Mask = ""
   txtCampo(19).Editable = False
   txtCampo(19).BoundColumn = ""
   txtCampo(19).ListFields = ""
   txtCampo(19).OrderFields = ""
   txtCampo(19).Relation = ""
   txtCampo(19).Source = ""

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = ""
   txtCampo(6).Editable = False
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = ""
   txtCampo(7).Editable = False
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 0
   txtCampo(10).Mask = ""
   txtCampo(10).Editable = False
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""

   Set txtCampo(29).CtPri = txtCp(29)
   txtCampo(29).DataType = 0
   txtCampo(29).Mask = ""
   txtCampo(29).Editable = False
   txtCampo(29).BoundColumn = ""
   txtCampo(29).ListFields = ""
   txtCampo(29).OrderFields = ""
   txtCampo(29).Relation = ""
   txtCampo(29).Source = ""

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 0
   txtCampo(15).Mask = ""
   txtCampo(15).Editable = False
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""

   Set txtCampo(26).CtPri = txtCp(26)
   txtCampo(26).DataType = 0
   txtCampo(26).Mask = ""
   txtCampo(26).Editable = False
   txtCampo(26).BoundColumn = ""
   txtCampo(26).ListFields = ""
   txtCampo(26).OrderFields = ""
   txtCampo(26).Relation = ""
   txtCampo(26).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = ""
   txtCampo(1).Editable = False
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(30).CtPri = txtCp(30)
   txtCampo(30).DataType = 0
   txtCampo(30).Mask = ""
   txtCampo(30).Editable = False
   txtCampo(30).BoundColumn = ""
   txtCampo(30).ListFields = ""
   txtCampo(30).OrderFields = ""
   txtCampo(30).Relation = ""
   txtCampo(30).Source = ""

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 0
   txtCampo(11).Mask = ""
   txtCampo(11).Editable = False
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""

   Set chkCampo(0).CtPri = ChkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set txtCampo(52).CtPri = txtCp(52)
   txtCampo(52).DataType = 1
   txtCampo(52).Mask = "99.999.999,99"
   txtCampo(52).Editable = False
   txtCampo(52).BoundColumn = ""
   txtCampo(52).ListFields = ""
   txtCampo(52).OrderFields = ""
   txtCampo(52).Relation = ""
   txtCampo(52).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(52).DataField), txtCampo(52)

   Set txtCampo(53).CtPri = txtCp(53)
   txtCampo(53).DataType = 1
   txtCampo(53).Mask = "99.999.999,99"
   txtCampo(53).Editable = False
   txtCampo(53).BoundColumn = ""
   txtCampo(53).ListFields = ""
   txtCampo(53).OrderFields = ""
   txtCampo(53).Relation = ""
   txtCampo(53).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(53).DataField), txtCampo(53)

   Set chkCampo(5).CtPri = ChkCp(5)
   chkCampo(5).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(5).DataField), chkCampo(5)

   Set chkCampo(6).CtPri = ChkCp(6)
   chkCampo(6).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(6).DataField), chkCampo(6)

   Set chkCampo(7).CtPri = ChkCp(7)
   chkCampo(7).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(7).DataField), chkCampo(7)

   Set chkCampo(8).CtPri = ChkCp(8)
   chkCampo(8).DataType = 5
   chkCampo(8).Editable = False
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(8).DataField), chkCampo(8)

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 4
   txtCampo(5).Mask = ""
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set chkCampo(3).CtPri = ChkCp(3)
   chkCampo(3).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(3).DataField), chkCampo(3)

   Set txtCampo(61).CtPri = txtCp(61)
   txtCampo(61).DataType = 0
   txtCampo(61).Mask = "@x"
   txtCampo(61).BoundColumn = ""
   txtCampo(61).ListFields = ""
   txtCampo(61).OrderFields = ""
   txtCampo(61).Relation = ""
   txtCampo(61).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(61).DataField), txtCampo(61)

   Set txtCampo(62).CtPri = txtCp(62)
   Set txtCampo(62).CtFdo = labFdo62
   Set txtCampo(62).CtBot(BOT_COMBO) = bottxtCampo62(BOT_COMBO)
   Set bottxtCampo62(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(62).DataType = 0
   txtCampo(62).Mask = "@x"
   txtCampo(62).BoundColumn = "Titular da Conta"
   txtCampo(62).ListFields = "Titular da Conta"
   txtCampo(62).OrderFields = "Titular da Conta"
   txtCampo(62).Relation = ""
   txtCampo(62).Source = "Conta do Vendedor"
   txtCampo(62).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(62).DataField), txtCampo(62)

   Set txtCampo(63).CtPri = txtCp(63)
   txtCampo(63).DataType = 1
   txtCampo(63).Mask = "999,9999"
   txtCampo(63).BoundColumn = ""
   txtCampo(63).ListFields = ""
   txtCampo(63).OrderFields = ""
   txtCampo(63).Relation = ""
   txtCampo(63).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(63).DataField), txtCampo(63)

   Set txtCampo(137).CtPri = txtCp(137)
   Set txtCampo(137).CtFdo = labFdo137
   Set txtCampo(137).CtBot(BOT_COMBO) = bottxtCampo137(BOT_COMBO)
   Set bottxtCampo137(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(137).DataType = 1
   txtCampo(137).Mask = "999999"
   txtCampo(137).BoundColumn = "Id da Conta"
   txtCampo(137).ListFields = "Titular da Conta"
   txtCampo(137).OrderFields = "Titular da Conta"
   txtCampo(137).Relation = ""
   txtCampo(137).Source = "Conta do Vendedor"
   txtCampo(137).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(137).DataField), txtCampo(137)

   Set txtCampo(64).CtPri = txtCp(64)
   Set txtCampo(64).CtFdo = labFdo64
   Set txtCampo(64).CtBot(BOT_COMBO) = bottxtCampo64(BOT_COMBO)
   Set bottxtCampo64(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(64).DataType = 0
   txtCampo(64).Mask = "@x"
   txtCampo(64).BoundColumn = "Titular da Conta"
   txtCampo(64).ListFields = "Titular da Conta"
   txtCampo(64).OrderFields = "Titular da Conta"
   txtCampo(64).Relation = ""
   txtCampo(64).Source = "Conta do Vendedor"
   txtCampo(64).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(64).DataField), txtCampo(64)

   Set txtCampo(65).CtPri = txtCp(65)
   txtCampo(65).DataType = 1
   txtCampo(65).Mask = "999,9999"
   txtCampo(65).BoundColumn = ""
   txtCampo(65).ListFields = ""
   txtCampo(65).OrderFields = ""
   txtCampo(65).Relation = ""
   txtCampo(65).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(65).DataField), txtCampo(65)

 DefineControles2

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles1"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub

Public Sub DefineControles2()
 On Error GoTo DeuErro

   Set txtCampo(56).CtPri = txtCp(56)
   txtCampo(56).DataType = 1
   txtCampo(56).Mask = "99.999.999,99"
   txtCampo(56).Editable = False
   txtCampo(56).BoundColumn = ""
   txtCampo(56).ListFields = ""
   txtCampo(56).OrderFields = ""
   txtCampo(56).Relation = ""
   txtCampo(56).Source = ""

   Set chkCampo(9).CtPri = ChkCp(9)
   chkCampo(9).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(9).DataField), chkCampo(9)

   Set txtCampo(66).CtPri = txtCp(66)
   txtCampo(66).DataType = 0
   txtCampo(66).Mask = "@!"
   txtCampo(66).BoundColumn = ""
   txtCampo(66).ListFields = ""
   txtCampo(66).OrderFields = ""
   txtCampo(66).Relation = ""
   txtCampo(66).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(66).DataField), txtCampo(66)

   Set txtCampo(67).CtPri = txtCp(67)
   txtCampo(67).DataType = 1
   txtCampo(67).Mask = "99.999,99"
   txtCampo(67).BoundColumn = ""
   txtCampo(67).ListFields = ""
   txtCampo(67).OrderFields = ""
   txtCampo(67).Relation = ""
   txtCampo(67).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(67).DataField), txtCampo(67)

   Set txtCampo(68).CtPri = txtCp(68)
   txtCampo(68).DataType = 1
   txtCampo(68).Mask = "99.999,99"
   txtCampo(68).BoundColumn = ""
   txtCampo(68).ListFields = ""
   txtCampo(68).OrderFields = ""
   txtCampo(68).Relation = ""
   txtCampo(68).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(68).DataField), txtCampo(68)

   Set txtCampo(69).CtPri = txtCp(69)
   txtCampo(69).DataType = 1
   txtCampo(69).Mask = "9.999,99"
   txtCampo(69).BoundColumn = ""
   txtCampo(69).ListFields = ""
   txtCampo(69).OrderFields = ""
   txtCampo(69).Relation = ""
   txtCampo(69).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(69).DataField), txtCampo(69)

   Set txtCampo(72).CtPri = txtCp(72)
   txtCampo(72).DataType = 1
   txtCampo(72).Mask = "9.999,99"
   txtCampo(72).BoundColumn = ""
   txtCampo(72).ListFields = ""
   txtCampo(72).OrderFields = ""
   txtCampo(72).Relation = ""
   txtCampo(72).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(72).DataField), txtCampo(72)

   Set txtCampo(73).CtPri = txtCp(73)
   txtCampo(73).DataType = 0
   txtCampo(73).Mask = "@!"
   txtCampo(73).BoundColumn = ""
   txtCampo(73).ListFields = ""
   txtCampo(73).OrderFields = ""
   txtCampo(73).Relation = ""
   txtCampo(73).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(73).DataField), txtCampo(73)

   Set txtCampo(74).CtPri = txtCp(74)
   txtCampo(74).DataType = 0
   txtCampo(74).Mask = "@!"
   txtCampo(74).BoundColumn = ""
   txtCampo(74).ListFields = ""
   txtCampo(74).OrderFields = ""
   txtCampo(74).Relation = ""
   txtCampo(74).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(74).DataField), txtCampo(74)

   Set txtCampo(75).CtPri = txtCp(75)
   txtCampo(75).DataType = 1
   txtCampo(75).Mask = "9.999,99"
   txtCampo(75).BoundColumn = ""
   txtCampo(75).ListFields = ""
   txtCampo(75).OrderFields = ""
   txtCampo(75).Relation = ""
   txtCampo(75).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(75).DataField), txtCampo(75)

   Set txtCampo(76).CtPri = txtCp(76)
   txtCampo(76).DataType = 1
   txtCampo(76).Mask = "99.999.99,99"
   txtCampo(76).Editable = False
   txtCampo(76).BoundColumn = ""
   txtCampo(76).ListFields = ""
   txtCampo(76).OrderFields = ""
   txtCampo(76).Relation = ""
   txtCampo(76).Source = ""

   Set txtCampo(77).CtPri = txtCp(77)
   txtCampo(77).DataType = 1
   txtCampo(77).Mask = "9.999,99"
   txtCampo(77).BoundColumn = ""
   txtCampo(77).ListFields = ""
   txtCampo(77).OrderFields = ""
   txtCampo(77).Relation = ""
   txtCampo(77).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(77).DataField), txtCampo(77)

   Set txtCampo(78).CtPri = txtCp(78)
   txtCampo(78).DataType = 1
   txtCampo(78).Mask = "9.999,99"
   txtCampo(78).BoundColumn = ""
   txtCampo(78).ListFields = ""
   txtCampo(78).OrderFields = ""
   txtCampo(78).Relation = ""
   txtCampo(78).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(78).DataField), txtCampo(78)

   Set txtCampo(81).CtPri = txtCp(81)
   txtCampo(81).DataType = 1
   txtCampo(81).Mask = "9.999,99"
   txtCampo(81).BoundColumn = ""
   txtCampo(81).ListFields = ""
   txtCampo(81).OrderFields = ""
   txtCampo(81).Relation = ""
   txtCampo(81).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(81).DataField), txtCampo(81)

   Set txtCampo(82).CtPri = txtCp(82)
   txtCampo(82).DataType = 1
   txtCampo(82).Mask = "9.999,99"
   txtCampo(82).BoundColumn = ""
   txtCampo(82).ListFields = ""
   txtCampo(82).OrderFields = ""
   txtCampo(82).Relation = ""
   txtCampo(82).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(82).DataField), txtCampo(82)

   Set txtCampo(79).CtPri = txtCp(79)
   txtCampo(79).DataType = 1
   txtCampo(79).Mask = "99.999.99,99"
   txtCampo(79).Editable = False
   txtCampo(79).BoundColumn = ""
   txtCampo(79).ListFields = ""
   txtCampo(79).OrderFields = ""
   txtCampo(79).Relation = ""
   txtCampo(79).Source = ""

   Set txtCampo(84).CtPri = txtCp(84)
   txtCampo(84).DataType = 1
   txtCampo(84).Mask = "99.999.99,99"
   txtCampo(84).Editable = False
   txtCampo(84).BoundColumn = ""
   txtCampo(84).ListFields = ""
   txtCampo(84).OrderFields = ""
   txtCampo(84).Relation = ""
   txtCampo(84).Source = ""

   Set txtCampo(85).CtPri = txtCp(85)
   txtCampo(85).DataType = 1
   txtCampo(85).Mask = "99.999.99,99"
   txtCampo(85).Editable = False
   txtCampo(85).BoundColumn = ""
   txtCampo(85).ListFields = ""
   txtCampo(85).OrderFields = ""
   txtCampo(85).Relation = ""
   txtCampo(85).Source = ""

   Set txtCampo(88).CtPri = txtCp(88)
   txtCampo(88).DataType = 1
   txtCampo(88).Mask = "9.999,99"
   txtCampo(88).BoundColumn = ""
   txtCampo(88).ListFields = ""
   txtCampo(88).OrderFields = ""
   txtCampo(88).Relation = ""
   txtCampo(88).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(88).DataField), txtCampo(88)

   Set txtCampo(89).CtPri = txtCp(89)
   txtCampo(89).DataType = 1
   txtCampo(89).Mask = "9.999,99"
   txtCampo(89).BoundColumn = ""
   txtCampo(89).ListFields = ""
   txtCampo(89).OrderFields = ""
   txtCampo(89).Relation = ""
   txtCampo(89).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(89).DataField), txtCampo(89)

   Set txtCampo(90).CtPri = txtCp(90)
   txtCampo(90).DataType = 1
   txtCampo(90).Mask = "99,99"
   txtCampo(90).BoundColumn = ""
   txtCampo(90).ListFields = ""
   txtCampo(90).OrderFields = ""
   txtCampo(90).Relation = ""
   txtCampo(90).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(90).DataField), txtCampo(90)

   Set txtCampo(95).CtPri = txtCp(95)
   txtCampo(95).DataType = 1
   txtCampo(95).Mask = "99,99"
   txtCampo(95).BoundColumn = ""
   txtCampo(95).ListFields = ""
   txtCampo(95).OrderFields = ""
   txtCampo(95).Relation = ""
   txtCampo(95).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(95).DataField), txtCampo(95)

   Set txtCampo(70).CtPri = txtCp(70)
   txtCampo(70).DataType = 1
   txtCampo(70).Mask = "99.999.99,99"
   txtCampo(70).Editable = False
   txtCampo(70).BoundColumn = ""
   txtCampo(70).ListFields = ""
   txtCampo(70).OrderFields = ""
   txtCampo(70).Relation = ""
   txtCampo(70).Source = ""

   Set txtCampo(71).CtPri = txtCp(71)
   txtCampo(71).DataType = 1
   txtCampo(71).Mask = "99.999.99,99"
   txtCampo(71).Editable = False
   txtCampo(71).BoundColumn = ""
   txtCampo(71).ListFields = ""
   txtCampo(71).OrderFields = ""
   txtCampo(71).Relation = ""
   txtCampo(71).Source = ""

   Set txtCampo(80).CtPri = txtCp(80)
   txtCampo(80).DataType = 1
   txtCampo(80).Mask = "99.999.99,99"
   txtCampo(80).Editable = False
   txtCampo(80).BoundColumn = ""
   txtCampo(80).ListFields = ""
   txtCampo(80).OrderFields = ""
   txtCampo(80).Relation = ""
   txtCampo(80).Source = ""

   Set txtCampo(86).CtPri = txtCp(86)
   txtCampo(86).DataType = 1
   txtCampo(86).Mask = "99.999.99,99"
   txtCampo(86).Editable = False
   txtCampo(86).BoundColumn = ""
   txtCampo(86).ListFields = ""
   txtCampo(86).OrderFields = ""
   txtCampo(86).Relation = ""
   txtCampo(86).Source = ""

   Set txtCampo(83).CtPri = txtCp(83)
   txtCampo(83).DataType = 1
   txtCampo(83).Mask = "99.999.99,99"
   txtCampo(83).Editable = False
   txtCampo(83).BoundColumn = ""
   txtCampo(83).ListFields = ""
   txtCampo(83).OrderFields = ""
   txtCampo(83).Relation = ""
   txtCampo(83).Source = ""

   Set txtCampo(91).CtPri = txtCp(91)
   txtCampo(91).DataType = 1
   txtCampo(91).Mask = "99.999.99,99"
   txtCampo(91).Editable = False
   txtCampo(91).BoundColumn = ""
   txtCampo(91).ListFields = ""
   txtCampo(91).OrderFields = ""
   txtCampo(91).Relation = ""
   txtCampo(91).Source = ""

   Set txtCampo(92).CtPri = txtCp(92)
   txtCampo(92).DataType = 1
   txtCampo(92).Mask = "99.999.99,99"
   txtCampo(92).Editable = False
   txtCampo(92).BoundColumn = ""
   txtCampo(92).ListFields = ""
   txtCampo(92).OrderFields = ""
   txtCampo(92).Relation = ""
   txtCampo(92).Source = ""

   Set txtCampo(93).CtPri = txtCp(93)
   txtCampo(93).DataType = 1
   txtCampo(93).Mask = "99.999.99,99"
   txtCampo(93).Editable = False
   txtCampo(93).BoundColumn = ""
   txtCampo(93).ListFields = ""
   txtCampo(93).OrderFields = ""
   txtCampo(93).Relation = ""
   txtCampo(93).Source = ""

   Set txtCampo(87).CtPri = txtCp(87)
   txtCampo(87).DataType = 1
   txtCampo(87).Mask = "99.999.99,99"
   txtCampo(87).Editable = False
   txtCampo(87).BoundColumn = ""
   txtCampo(87).ListFields = ""
   txtCampo(87).OrderFields = ""
   txtCampo(87).Relation = ""
   txtCampo(87).Source = ""

   Set txtCampo(94).CtPri = txtCp(94)
   txtCampo(94).DataType = 1
   txtCampo(94).Mask = "99.999.99,99"
   txtCampo(94).Editable = False
   txtCampo(94).BoundColumn = ""
   txtCampo(94).ListFields = ""
   txtCampo(94).OrderFields = ""
   txtCampo(94).Relation = ""
   txtCampo(94).Source = ""

   Set txtCampo(96).CtPri = txtCp(96)
   txtCampo(96).DataType = 1
   txtCampo(96).Mask = "99.999,99"
   txtCampo(96).BoundColumn = ""
   txtCampo(96).ListFields = ""
   txtCampo(96).OrderFields = ""
   txtCampo(96).Relation = ""
   txtCampo(96).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(96).DataField), txtCampo(96)

   Set txtCampo(97).CtPri = txtCp(97)
   txtCampo(97).DataType = 1
   txtCampo(97).Mask = "99.999,99"
   txtCampo(97).BoundColumn = ""
   txtCampo(97).ListFields = ""
   txtCampo(97).OrderFields = ""
   txtCampo(97).Relation = ""
   txtCampo(97).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(97).DataField), txtCampo(97)

   Set txtCampo(98).CtPri = txtCp(98)
   txtCampo(98).DataType = 1
   txtCampo(98).Mask = "99.999,99"
   txtCampo(98).BoundColumn = ""
   txtCampo(98).ListFields = ""
   txtCampo(98).OrderFields = ""
   txtCampo(98).Relation = ""
   txtCampo(98).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(98).DataField), txtCampo(98)

   Set txtCampo(99).CtPri = txtCp(99)
   Set txtCampo(99).CtFdo = labFdo99
   Set txtCampo(99).CtBot(BOT_LISTA) = bottxtCampo99(BOT_LISTA)
   Set txtCampo(99).CtBot(BOT_COMBO) = bottxtCampo99(BOT_COMBO)
   bottxtCampo99(BOT_LISTA).Caption = "P"
   Set bottxtCampo99(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(99).DataType = 1
   txtCampo(99).Mask = "999999"
   txtCampo(99).PesqModoAbertura = 2
   txtCampo(99).PesqFieldCapture = "Adutoras.[Sequencia da Adutora]"
   txtCampo(99).BoundColumn = "Sequencia da Adutora"
   txtCampo(99).ListFields = "Modelo da Adutora"
   txtCampo(99).OrderFields = "Modelo da Adutora"
   txtCampo(99).Relation = ""
   txtCampo(99).Source = "Adutoras"
   txtCampo(99).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(99).DataField), txtCampo(99)

   Set txtCampo(100).CtPri = txtCp(100)
   Set txtCampo(100).CtFdo = labFdo100
   Set txtCampo(100).CtBot(BOT_LISTA) = bottxtCampo100(BOT_LISTA)
   Set txtCampo(100).CtBot(BOT_COMBO) = bottxtCampo100(BOT_COMBO)
   bottxtCampo100(BOT_LISTA).Caption = "P"
   Set bottxtCampo100(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(100).DataType = 1
   txtCampo(100).Mask = "999999"
   txtCampo(100).PesqModoAbertura = 2
   txtCampo(100).PesqFieldCapture = "Adutoras.[Sequencia da Adutora]"
   txtCampo(100).BoundColumn = "Sequencia da Adutora"
   txtCampo(100).ListFields = "Modelo da Adutora"
   txtCampo(100).OrderFields = "Modelo da Adutora"
   txtCampo(100).Relation = ""
   txtCampo(100).Source = "Adutoras"
   txtCampo(100).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(100).DataField), txtCampo(100)

   Set txtCampo(101).CtPri = txtCp(101)
   Set txtCampo(101).CtFdo = labFdo101
   Set txtCampo(101).CtBot(BOT_LISTA) = bottxtCampo101(BOT_LISTA)
   Set txtCampo(101).CtBot(BOT_COMBO) = bottxtCampo101(BOT_COMBO)
   bottxtCampo101(BOT_LISTA).Caption = "P"
   Set bottxtCampo101(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(101).DataType = 1
   txtCampo(101).Mask = "999999"
   txtCampo(101).PesqModoAbertura = 2
   txtCampo(101).PesqFieldCapture = "Adutoras.[Sequencia da Adutora]"
   txtCampo(101).BoundColumn = "Sequencia da Adutora"
   txtCampo(101).ListFields = "Modelo da Adutora"
   txtCampo(101).OrderFields = "Modelo da Adutora"
   txtCampo(101).Relation = ""
   txtCampo(101).Source = "Adutoras"
   txtCampo(101).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(101).DataField), txtCampo(101)

   Set txtCampo(102).CtPri = txtCp(102)
   txtCampo(102).DataType = 1
   txtCampo(102).Mask = "99.999.999,99"
   txtCampo(102).Editable = False
   txtCampo(102).BoundColumn = ""
   txtCampo(102).ListFields = ""
   txtCampo(102).OrderFields = ""
   txtCampo(102).Relation = ""
   txtCampo(102).Source = ""

   Set txtCampo(103).CtPri = txtCp(103)
   txtCampo(103).DataType = 1
   txtCampo(103).Mask = "99.999.999,99"
   txtCampo(103).Editable = False
   txtCampo(103).BoundColumn = ""
   txtCampo(103).ListFields = ""
   txtCampo(103).OrderFields = ""
   txtCampo(103).Relation = ""
   txtCampo(103).Source = ""

   Set txtCampo(104).CtPri = txtCp(104)
   txtCampo(104).DataType = 1
   txtCampo(104).Mask = "99.999.999,99"
   txtCampo(104).Editable = False
   txtCampo(104).BoundColumn = ""
   txtCampo(104).ListFields = ""
   txtCampo(104).OrderFields = ""
   txtCampo(104).Relation = ""
   txtCampo(104).Source = ""

   Set txtCampo(105).CtPri = txtCp(105)
   txtCampo(105).DataType = 1
   txtCampo(105).Mask = "99.999.999,99"
   txtCampo(105).Editable = False
   txtCampo(105).BoundColumn = ""
   txtCampo(105).ListFields = ""
   txtCampo(105).OrderFields = ""
   txtCampo(105).Relation = ""
   txtCampo(105).Source = ""

   Set txtCampo(106).CtPri = txtCp(106)
   txtCampo(106).DataType = 1
   txtCampo(106).Mask = "99.999.999,99"
   txtCampo(106).Editable = False
   txtCampo(106).BoundColumn = ""
   txtCampo(106).ListFields = ""
   txtCampo(106).OrderFields = ""
   txtCampo(106).Relation = ""
   txtCampo(106).Source = ""

   Set txtCampo(107).CtPri = txtCp(107)
   txtCampo(107).DataType = 1
   txtCampo(107).Mask = "99.999.999,99"
   txtCampo(107).Editable = False
   txtCampo(107).BoundColumn = ""
   txtCampo(107).ListFields = ""
   txtCampo(107).OrderFields = ""
   txtCampo(107).Relation = ""
   txtCampo(107).Source = ""

   Set txtCampo(108).CtPri = txtCp(108)
   txtCampo(108).DataType = 1
   txtCampo(108).Mask = "99.999.999,99"
   txtCampo(108).Editable = False
   txtCampo(108).BoundColumn = ""
   txtCampo(108).ListFields = ""
   txtCampo(108).OrderFields = ""
   txtCampo(108).Relation = ""
   txtCampo(108).Source = ""

   Set txtCampo(109).CtPri = txtCp(109)
   txtCampo(109).DataType = 1
   txtCampo(109).Mask = "99.999.999,99"
   txtCampo(109).Editable = False
   txtCampo(109).BoundColumn = ""
   txtCampo(109).ListFields = ""
   txtCampo(109).OrderFields = ""
   txtCampo(109).Relation = ""
   txtCampo(109).Source = ""

   Set txtCampo(110).CtPri = txtCp(110)
   txtCampo(110).DataType = 1
   txtCampo(110).Mask = "99.999.999,99"
   txtCampo(110).Editable = False
   txtCampo(110).BoundColumn = ""
   txtCampo(110).ListFields = ""
   txtCampo(110).OrderFields = ""
   txtCampo(110).Relation = ""
   txtCampo(110).Source = ""

   Set txtCampo(111).CtPri = txtCp(111)
   txtCampo(111).DataType = 1
   txtCampo(111).Mask = "99.999.999,99"
   txtCampo(111).Editable = False
   txtCampo(111).BoundColumn = ""
   txtCampo(111).ListFields = ""
   txtCampo(111).OrderFields = ""
   txtCampo(111).Relation = ""
   txtCampo(111).Source = ""

   Set txtCampo(112).CtPri = txtCp(112)
   txtCampo(112).DataType = 1
   txtCampo(112).Mask = "99.999.999,99"
   txtCampo(112).Editable = False
   txtCampo(112).BoundColumn = ""
   txtCampo(112).ListFields = ""
   txtCampo(112).OrderFields = ""
   txtCampo(112).Relation = ""
   txtCampo(112).Source = ""

 DefineControles3

 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles2"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub


Public Sub DefineControles3()
 On Error GoTo DeuErro

   Set txtCampo(113).CtPri = txtCp(113)
   txtCampo(113).DataType = 1
   txtCampo(113).Mask = "99.999.999,99"
   txtCampo(113).Editable = False
   txtCampo(113).BoundColumn = ""
   txtCampo(113).ListFields = ""
   txtCampo(113).OrderFields = ""
   txtCampo(113).Relation = ""
   txtCampo(113).Source = ""

   Set txtCampo(114).CtPri = txtCp(114)
   txtCampo(114).DataType = 1
   txtCampo(114).Mask = "99.999.99,99"
   txtCampo(114).Editable = False
   txtCampo(114).BoundColumn = ""
   txtCampo(114).ListFields = ""
   txtCampo(114).OrderFields = ""
   txtCampo(114).Relation = ""
   txtCampo(114).Source = ""

   Set txtCampo(115).CtPri = txtCp(115)
   txtCampo(115).DataType = 1
   txtCampo(115).Mask = "99"
   txtCampo(115).BoundColumn = ""
   txtCampo(115).ListFields = ""
   txtCampo(115).OrderFields = ""
   txtCampo(115).Relation = ""
   txtCampo(115).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(115).DataField), txtCampo(115)

   Set txtCampo(116).CtPri = txtCp(116)
   txtCampo(116).DataType = 0
   txtCampo(116).Mask = "@x"
   txtCampo(116).BoundColumn = ""
   txtCampo(116).ListFields = ""
   txtCampo(116).OrderFields = ""
   txtCampo(116).Relation = ""
   txtCampo(116).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(116).DataField), txtCampo(116)

   Set txtCampo(117).CtPri = txtCp(117)
   txtCampo(117).DataType = 0
   txtCampo(117).Mask = "@x"
   txtCampo(117).BoundColumn = ""
   txtCampo(117).ListFields = ""
   txtCampo(117).OrderFields = ""
   txtCampo(117).Relation = ""
   txtCampo(117).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(117).DataField), txtCampo(117)

   Set txtCampo(118).CtPri = txtCp(118)
   txtCampo(118).DataType = 0
   txtCampo(118).Mask = "@x"
   txtCampo(118).BoundColumn = ""
   txtCampo(118).ListFields = ""
   txtCampo(118).OrderFields = ""
   txtCampo(118).Relation = ""
   txtCampo(118).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(118).DataField), txtCampo(118)

   Set txtCampo(119).CtPri = txtCp(119)
   txtCampo(119).DataType = 1
   txtCampo(119).Mask = "99"
   txtCampo(119).BoundColumn = ""
   txtCampo(119).ListFields = ""
   txtCampo(119).OrderFields = ""
   txtCampo(119).Relation = ""
   txtCampo(119).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(119).DataField), txtCampo(119)

   Set txtCampo(120).CtPri = txtCp(120)
   txtCampo(120).DataType = 1
   txtCampo(120).Mask = "99.999,99"
   txtCampo(120).BoundColumn = ""
   txtCampo(120).ListFields = ""
   txtCampo(120).OrderFields = ""
   txtCampo(120).Relation = ""
   txtCampo(120).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(120).DataField), txtCampo(120)

   Set txtCampo(121).CtPri = txtCp(121)
   txtCampo(121).DataType = 1
   txtCampo(121).Mask = "99.999.99,99"
   txtCampo(121).Editable = False
   txtCampo(121).BoundColumn = ""
   txtCampo(121).ListFields = ""
   txtCampo(121).OrderFields = ""
   txtCampo(121).Relation = ""
   txtCampo(121).Source = ""

   Set txtCampo(122).CtPri = txtCp(122)
   txtCampo(122).DataType = 1
   txtCampo(122).Mask = "99.999,99"
   txtCampo(122).Editable = False
   txtCampo(122).BoundColumn = ""
   txtCampo(122).ListFields = ""
   txtCampo(122).OrderFields = ""
   txtCampo(122).Relation = ""
   txtCampo(122).Source = ""

   Set txtCampo(123).CtPri = txtCp(123)
   txtCampo(123).DataType = 1
   txtCampo(123).Mask = "99.999,99"
   txtCampo(123).BoundColumn = ""
   txtCampo(123).ListFields = ""
   txtCampo(123).OrderFields = ""
   txtCampo(123).Relation = ""
   txtCampo(123).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(123).DataField), txtCampo(123)

   Set txtCampo(124).CtPri = txtCp(124)
   txtCampo(124).DataType = 1
   txtCampo(124).Mask = "99.999,99"
   txtCampo(124).BoundColumn = ""
   txtCampo(124).ListFields = ""
   txtCampo(124).OrderFields = ""
   txtCampo(124).Relation = ""
   txtCampo(124).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(124).DataField), txtCampo(124)

   Set txtCampo(125).CtPri = txtCp(125)
   txtCampo(125).DataType = 1
   txtCampo(125).Mask = "99.999,99"
   txtCampo(125).Editable = False
   txtCampo(125).BoundColumn = ""
   txtCampo(125).ListFields = ""
   txtCampo(125).OrderFields = ""
   txtCampo(125).Relation = ""
   txtCampo(125).Source = ""

   Set txtCampo(126).CtPri = txtCp(126)
   txtCampo(126).DataType = 1
   txtCampo(126).Mask = "9.999.999,99"
   txtCampo(126).BoundColumn = ""
   txtCampo(126).ListFields = ""
   txtCampo(126).OrderFields = ""
   txtCampo(126).Relation = ""
   txtCampo(126).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(126).DataField), txtCampo(126)

   Set txtCampo(127).CtPri = txtCp(127)
   txtCampo(127).DataType = 0
   txtCampo(127).Mask = "@x"
   txtCampo(127).BoundColumn = ""
   txtCampo(127).ListFields = ""
   txtCampo(127).OrderFields = ""
   txtCampo(127).Relation = ""
   txtCampo(127).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(127).DataField), txtCampo(127)

   Set txtCampo(128).CtPri = txtCp(128)
   txtCampo(128).DataType = 0
   txtCampo(128).Mask = "@x"
   txtCampo(128).BoundColumn = ""
   txtCampo(128).ListFields = ""
   txtCampo(128).OrderFields = ""
   txtCampo(128).Relation = ""
   txtCampo(128).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(128).DataField), txtCampo(128)

   Set txtCampo(129).CtPri = txtCp(129)
   txtCampo(129).DataType = 0
   txtCampo(129).Mask = "@x"
   txtCampo(129).BoundColumn = ""
   txtCampo(129).ListFields = ""
   txtCampo(129).OrderFields = ""
   txtCampo(129).Relation = ""
   txtCampo(129).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(129).DataField), txtCampo(129)

   Set txtCampo(130).CtPri = txtCp(130)
   txtCampo(130).DataType = 1
   txtCampo(130).Mask = "99.999,99"
   txtCampo(130).BoundColumn = ""
   txtCampo(130).ListFields = ""
   txtCampo(130).OrderFields = ""
   txtCampo(130).Relation = ""
   txtCampo(130).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(130).DataField), txtCampo(130)

   Set txtCampo(131).CtPri = txtCp(131)
   txtCampo(131).DataType = 1
   txtCampo(131).Mask = "999"
   txtCampo(131).BoundColumn = ""
   txtCampo(131).ListFields = ""
   txtCampo(131).OrderFields = ""
   txtCampo(131).Relation = ""
   txtCampo(131).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(131).DataField), txtCampo(131)

   Set txtCampo(132).CtPri = txtCp(132)
   txtCampo(132).DataType = 1
   txtCampo(132).Mask = "99.999.99,99"
   txtCampo(132).Editable = False
   txtCampo(132).BoundColumn = ""
   txtCampo(132).ListFields = ""
   txtCampo(132).OrderFields = ""
   txtCampo(132).Relation = ""
   txtCampo(132).Source = ""

   Set txtCampo(133).CtPri = txtCp(133)
   txtCampo(133).DataType = 1
   txtCampo(133).Mask = "99.999.99,99"
   txtCampo(133).Editable = False
   txtCampo(133).BoundColumn = ""
   txtCampo(133).ListFields = ""
   txtCampo(133).OrderFields = ""
   txtCampo(133).Relation = ""
   txtCampo(133).Source = ""

   Set txtCampo(134).CtPri = txtCp(134)
   txtCampo(134).DataType = 1
   txtCampo(134).Mask = "99.999,99"
   txtCampo(134).BoundColumn = ""
   txtCampo(134).ListFields = ""
   txtCampo(134).OrderFields = ""
   txtCampo(134).Relation = ""
   txtCampo(134).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(134).DataField), txtCampo(134)

   Set txtCampo(135).CtPri = txtCp(135)
   txtCampo(135).DataType = 1
   txtCampo(135).Mask = "99.999.99,99"
   txtCampo(135).Editable = False
   txtCampo(135).BoundColumn = ""
   txtCampo(135).ListFields = ""
   txtCampo(135).OrderFields = ""
   txtCampo(135).Relation = ""
   txtCampo(135).Source = ""

   Set txtCampo(136).CtPri = txtCp(136)
   txtCampo(136).DataType = 1
   txtCampo(136).Mask = "99.999.99,99"
   txtCampo(136).Editable = False
   txtCampo(136).BoundColumn = ""
   txtCampo(136).ListFields = ""
   txtCampo(136).OrderFields = ""
   txtCampo(136).Relation = ""
   txtCampo(136).Source = ""

   Set txtCampo(138).CtPri = txtCp(138)
   txtCampo(138).DataType = 0
   txtCampo(138).Mask = "@!"
   txtCampo(138).BoundColumn = ""
   txtCampo(138).ListFields = ""
   txtCampo(138).OrderFields = ""
   txtCampo(138).Relation = ""
   txtCampo(138).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(138).DataField), txtCampo(138)

   Set txtCampo(139).CtPri = txtCp(139)
   txtCampo(139).DataType = 1
   txtCampo(139).Mask = "99"
   txtCampo(139).BoundColumn = ""
   txtCampo(139).ListFields = ""
   txtCampo(139).OrderFields = ""
   txtCampo(139).Relation = ""
   txtCampo(139).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(139).DataField), txtCampo(139)

   Set txtCampo(140).CtPri = txtCp(140)
   txtCampo(140).DataType = 1
   txtCampo(140).Mask = "99"
   txtCampo(140).BoundColumn = ""
   txtCampo(140).ListFields = ""
   txtCampo(140).OrderFields = ""
   txtCampo(140).Relation = ""
   txtCampo(140).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(140).DataField), txtCampo(140)

   Set txtCampo(141).CtPri = txtCp(141)
   txtCampo(141).DataType = 0
   txtCampo(141).Mask = "@x"
   txtCampo(141).BoundColumn = ""
   txtCampo(141).ListFields = ""
   txtCampo(141).OrderFields = ""
   txtCampo(141).Relation = ""
   txtCampo(141).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(141).DataField), txtCampo(141)

   Set txtCampo(142).CtPri = txtCp(142)
   txtCampo(142).DataType = 0
   txtCampo(142).Mask = "@x"
   txtCampo(142).BoundColumn = ""
   txtCampo(142).ListFields = ""
   txtCampo(142).OrderFields = ""
   txtCampo(142).Relation = ""
   txtCampo(142).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(142).DataField), txtCampo(142)

   Set txtCampo(143).CtPri = txtCp(143)
   txtCampo(143).DataType = 1
   txtCampo(143).Mask = "99.999.999,99"
   txtCampo(143).BoundColumn = ""
   txtCampo(143).ListFields = ""
   txtCampo(143).OrderFields = ""
   txtCampo(143).Relation = ""
   txtCampo(143).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(143).DataField), txtCampo(143)

   Set txtCampo(144).CtPri = txtCp(144)
   txtCampo(144).DataType = 1
   txtCampo(144).Mask = "99.999.999,99"
   txtCampo(144).BoundColumn = ""
   txtCampo(144).ListFields = ""
   txtCampo(144).OrderFields = ""
   txtCampo(144).Relation = ""
   txtCampo(144).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(144).DataField), txtCampo(144)

   Set txtCampo(145).CtPri = txtCp(145)
   txtCampo(145).DataType = 1
   txtCampo(145).Mask = "999999"
   txtCampo(145).BoundColumn = ""
   txtCampo(145).ListFields = ""
   txtCampo(145).OrderFields = ""
   txtCampo(145).Relation = ""
   txtCampo(145).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(145).DataField), txtCampo(145)

   Set txtCampo(146).CtPri = txtCp(146)
   txtCampo(146).DataType = 0
   txtCampo(146).Mask = "@x"
   txtCampo(146).BoundColumn = ""
   txtCampo(146).ListFields = ""
   txtCampo(146).OrderFields = ""
   txtCampo(146).Relation = ""
   txtCampo(146).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(146).DataField), txtCampo(146)

   Set txtCampo(147).CtPri = txtCp(147)
   Set txtCampo(147).CtFdo = labFdo147
   Set txtCampo(147).CtBot(BOT_COMBO) = bottxtCampo147(BOT_COMBO)
   Set bottxtCampo147(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(147).DataType = 0
   txtCampo(147).ListFields = "Irrigao Penapolis|Revenda"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(147).DataField), txtCampo(147)

   Set txtCampo(160).CtPri = txtCp(160)
   txtCampo(160).DataType = 0
   txtCampo(160).Mask = ""
   txtCampo(160).Editable = False
   txtCampo(160).BoundColumn = ""
   txtCampo(160).ListFields = ""
   txtCampo(160).OrderFields = ""
   txtCampo(160).Relation = ""
   txtCampo(160).Source = ""

   Set chkCampo(10).CtPri = ChkCp(10)
   chkCampo(10).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(10).DataField), chkCampo(10)

   Set txtCampo(149).CtPri = txtCp(149)
   Set txtCampo(149).CtFdo = labFdo149
   Set txtCampo(149).CtBot(BOT_ACAO) = bottxtCampo149(BOT_ACAO)
   Set bottxtCampo149(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(149).DataType = 2
   txtCampo(149).Mask = "99/99/9999"
   txtCampo(149).BoundColumn = ""
   txtCampo(149).ListFields = ""
   txtCampo(149).OrderFields = ""
   txtCampo(149).Relation = ""
   txtCampo(149).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(149).DataField), txtCampo(149)

   Set txtCampo(150).CtPri = txtCp(150)
   Set txtCampo(150).CtFdo = labFdo150
   Set txtCampo(150).CtBot(BOT_LISTA) = bottxtCampo150(BOT_LISTA)
   Set bottxtCampo150(BOT_LISTA).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(150).DataType = 2
   txtCampo(150).Mask = "99/99/9999"
   txtCampo(150).PesqModoAbertura = 2
   txtCampo(150).BoundColumn = ""
   txtCampo(150).ListFields = ""
   txtCampo(150).OrderFields = ""
   txtCampo(150).Relation = ""
   txtCampo(150).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(150).DataField), txtCampo(150)

   Set chkCampo(11).CtPri = ChkCp(11)
   chkCampo(11).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(11).DataField), chkCampo(11)

   Set txtCampo(161).CtPri = txtCp(161)
   Set txtCampo(161).CtFdo = labFdo161
   Set txtCampo(161).CtBot(BOT_COMBO) = bottxtCampo161(BOT_COMBO)
   Set bottxtCampo161(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(161).DataType = 1
   txtCampo(161).Mask = "999999"
   txtCampo(161).BoundColumn = "Seqncia do Oramento"
   txtCampo(161).ListFields = "Seqncia do Oramento"
   txtCampo(161).OrderFields = "Seqncia do Oramento-"
   txtCampo(161).Relation = ""
   txtCampo(161).Source = "Oramento"
   txtCampo(161).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(161).DataField), txtCampo(161)

   Set txtCampo(151).CtPri = txtCp(151)
   Set txtCampo(151).CtFdo = labFdo151
   Set txtCampo(151).CtBot(BOT_COMBO) = bottxtCampo151(BOT_COMBO)
   Set bottxtCampo151(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(151).DataType = 0
   txtCampo(151).ListFields = "Emitente|Destinatrio|Transporte Prprio Remetente|Transporte Prprio Destinatrio"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(151).DataField), txtCampo(151)

 
 
 '--- PlacaVeiculo (ndice 152) ---
Set txtCampo(152).CtPri = txtCp(152)
txtCampo(152).DataType = 0
txtCampo(152).Mask = "@x"
txtCp(152).MaxLength = 7                         ' limite antes do AddColumn
txtCampo(152).Source = "Oramento"
txtCampo(152).BoundColumn = "PlacaVeiculo"       ' ou 1, se seu componente exigir ndice
grdBrowse.AddColumn _
    vgDb.Tables(vgIdentTab).Columns("PlacaVeiculo"), txtCampo(152)

'--- UfPlaca (ndice 153) ---
Set txtCampo(153).CtPri = txtCp(153)
txtCampo(153).DataType = 0
txtCampo(153).Mask = "!!"                        ' duas letras
txtCp(153).MaxLength = 2
txtCampo(153).Source = "Oramento"
txtCampo(153).BoundColumn = "UfPlaca"
grdBrowse.AddColumn _
    vgDb.Tables(vgIdentTab).Columns("UfPlaca"), txtCampo(153)

'--- NumAntt (ndice 154) ---
Set txtCampo(154).CtPri = txtCp(154)
txtCampo(154).DataType = 0
txtCampo(154).Mask = ""                  ' oito dgitos
txtCp(154).MaxLength = 8
txtCampo(154).Source = "Oramento"
txtCampo(154).BoundColumn = "NumAntt"
grdBrowse.AddColumn _
    vgDb.Tables(vgIdentTab).Columns("NumAntt"), txtCampo(154)

 
 
 
 Exit Sub

DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles3"
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub



Private Sub txtCp_BeforeUpdate(Index As Integer, Cancel As Integer)
    If Index = 152 Then
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = "^[A-Z]{3}[0-9]{4}$|^[A-Z]{3}[0-9][A-Z][0-9]{2}$"
        re.IgnoreCase = False
        If Not re.Test(txtCp(Index).Text) Then
            MsgBox "Placa invlida! Use LLLNNNN ou LLLNLNN (Mercosul).", _
                   vbExclamation
            Cancel = True   ' cancela sada do controle
        End If
    End If
End Sub



'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 17
         txtCampo(17).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia do Municpio] > 0 AND Inativo = 0", "[Seqncia do Municpio] > 0") & ""
                           txtCampo(17).PesqSQLExpression = "SELECT Municpios.[Seqncia do Municpio], Municpios.Descrio, Municpios.[Cdigo do IBGE], Municpios.CEP FROM Municpios WHERE (Municpios.[Seqncia do Municpio] > " + CStr(0) + ") AND " + _
                                                               "(Municpios.Inativo = False)"
      Case 33
         txtCampo(33).Filter = "([Seqncia do Geral] > " & 0 & ") AND Vendedor = " & 1 & ""
                           txtCampo(33).PesqSQLExpression = "SELECT Geral.[Seqncia do Geral], Geral.[Razo Social], Geral.[Nome Fantasia], Geral.[CPF e CNPJ], Municpios.[Seqncia do Municpio], " + _
                                                               "Municpios.Descrio FROM Geral, Municpios WHERE (Geral.[Seqncia do Geral] > " + CStr(0) + ") AND (Geral.[Seqncia do Municpio] = " + _
                                                               "Municpios.[Seqncia do Municpio]) AND (Geral.Inativo = False) AND (Geral.Vendedor = " + CStr(1) + ")"
      Case 34
         txtCampo(34).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia da Classificao] > 0 AND Inativo = 0", "[Seqncia da Classificao] > 0") & ""
                           txtCampo(34).PesqSQLExpression = "SELECT [Classificao Fiscal].[Seqncia da Classificao], [Classificao Fiscal].NCM, [Classificao Fiscal].[Descrio do NCM] FROM [Classificao Fiscal] WHERE ([Classificao Fiscal].[Seqncia da Classificao] > " + CStr(0) + ") AND " + _
                                                               "([Classificao Fiscal].Inativo = False)"
      Case 55
         txtCampo(55).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "Transportadora = 1 AND [Seqncia do Geral] > 0 AND Inativo = 0", "[Seqncia do Geral] > 0") & ""
                           txtCampo(55).PesqSQLExpression = "SELECT Geral.[Seqncia do Geral], Geral.[Razo Social], Geral.[Nome Fantasia], Geral.Endereo, Geral.[CPF e CNPJ] FROM Geral WHERE (Geral.[Seqncia do Geral] > " + CStr(0) + ") AND " + _
                                                               "(Geral.Inativo = False) AND (Geral.Transportadora = True)"
      Case 60
         txtCampo(60).Filter = "([Seqncia do Pas] > " & 0 & ") AND Inativo = " & 0 & ""
                           txtCampo(60).PesqSQLExpression = "SELECT Pases.[Seqncia do Pas], Pases.Descrio, Pases.[Cdigo do Pas], Pases.Inativo FROM Pases WHERE (Pases.[Seqncia do Pas] > " + CStr(0) + ") AND " + _
                                                               "(Pases.Inativo = False)"
      Case 62
         txtCampo(62).Filter = "Desativado = " & 0 & ""
      Case 64
         txtCampo(64).Filter = "Desativado = " & 0 & ""
      Case 99
         txtCampo(99).Filter = "[Sequencia da Adutora] > " & 0 & ""
                           txtCampo(99).PesqSQLExpression = "SELECT Adutoras.[Sequencia da Adutora], Adutoras.[Modelo da Adutora], Adutoras.DN, Adutoras.[DN mm], " + _
                                                               "Adutoras.Coeficiente, Adutoras.Material, Adutoras.[E mm], Adutoras.[DI mm] FROM Adutoras WHERE (Adutoras.[Sequencia da Adutora] > " + CStr(0) + ")"
      Case 100
         txtCampo(100).Filter = "[Sequencia da Adutora] > " & 0 & ""
                           txtCampo(100).PesqSQLExpression = "SELECT Adutoras.[Sequencia da Adutora], Adutoras.[Modelo da Adutora], Adutoras.DN, Adutoras.[DN mm], " + _
                                                                "Adutoras.Coeficiente, Adutoras.Material, Adutoras.[E mm], Adutoras.[DI mm] FROM Adutoras WHERE (Adutoras.[Sequencia da Adutora] > " + CStr(0) + ")"
      Case 101
         txtCampo(101).Filter = "[Sequencia da Adutora] > " & 0 & ""
                           txtCampo(101).PesqSQLExpression = "SELECT Adutoras.[Sequencia da Adutora], Adutoras.[Modelo da Adutora], Adutoras.DN, Adutoras.[DN mm], " + _
                                                                "Adutoras.Coeficiente, Adutoras.Material, Adutoras.[E mm], Adutoras.[DI mm] FROM Adutoras WHERE (Adutoras.[Sequencia da Adutora] > " + CStr(0) + ")"
      Case 137
         txtCampo(137).Filter = "([Id da Conta] > " & 0 & ") AND [Faz projeto] = " & 1 & ""
      Case 156
         txtCampo(156).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "Cliente = 1 AND [Seqncia do Geral] > 0 AND Inativo = 0", "[Seqncia do Geral] > 0") & ""
                           txtCampo(156).PesqSQLExpression = "SELECT Geral.[Seqncia do Geral], Geral.[Razo Social], Geral.[Nome Fantasia], Geral.[CPF e CNPJ], Municpios.[Seqncia do Municpio], " + _
                                                                "Municpios.Descrio FROM Geral, Municpios WHERE (Geral.[Seqncia do Geral] > " + CStr(0) + ") AND (Geral.[Seqncia do Municpio] = " + _
                                                                "Municpios.[Seqncia do Municpio]) AND (Geral.Inativo = False) AND (Geral.Cliente = " + CStr(1) + ")"
      Case 157
         txtCampo(157).Filter = "[Seqncia da Propriedade] IN (SELECT [Seqncia da Propriedade] FROM [Propriedades do Geral] WHERE [Seqncia do Geral] = " & Sequencia_do_Geral & " AND Inativo = " & 0 & ")"
      Case 161
         txtCampo(161).Filter = "((([Seqncia do Oramento] > " & 0 & ") AND Cancelado = " & 0 & ") AND [Ordem Interna] = " & 0 & ") AND [Venda Fechada] = " & 1 & ""
   End Select
End Sub


'evento - antes de descarregar o formulrio
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If vgPodeFazerUnLoad = False Then
      If vgSituacao <> ACAO_NAVEGANDO And vgBotoesOK Then  'botoeira esta correta?
         AtivaForm Me                                      'entao coloca
      End If
      Cancel = FormPendente(Me)                            've se tem atualizacao pendente
      If Cancel = False Then
         Cancel = True
         timUnLoad.Enabled = True
      End If
   End If
End Sub


'evento - redefinido o tamanho do formulrio
Private Sub Form_Resize()
   AjustaPanFundo Me
   MudaTamCampos Me
End Sub


'evento - descarregando o formulrio da memria
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   VerificaFormulario
   FinalizaForm Me
   Set txtSequencia_do_Orcamento = Nothing
   Set Aba1 = Nothing
   Set txtCEP = Nothing
   Set txtCaixaPostal = Nothing
   Set txtFone = Nothing
   Set txtObservacao = Nothing
   Set txtMemoAuxiliar = Nothing
   Set lblRGIE = Nothing
   Set txtCPFCNPJ_F = Nothing
   Set lblCPFCNPJ_F = Nothing
   Set txtFax = Nothing
   Set txtRGIE_F = Nothing
   Set txtCPFCNPJ = Nothing
   Set txtRGIE = Nothing
   Set txtEmail = Nothing
   Set txtMunicipio = Nothing
   Set txtEndereco = Nothing
   Set txtUF = Nothing
   Set txtBairro = Nothing
   Set txtNumero = Nothing
   Set txtComplemento = Nothing
   Set txtVendedor = Nothing
   Set grdConjuntos = Nothing
   Set grdPecas = Nothing
   Set Grdparcelamento = Nothing
   Set txtForma_de_Pagamento = Nothing
   Set GrdProdutos = Nothing
   Set grdServicos = Nothing
   Set lblParcelamento = Nothing
   Set lblCPFCNPJ = Nothing
   Set Veiculo = Nothing
   Set txtISS = Nothing
   Set Txtperdas1 = Nothing
   Set Txtperdas2 = Nothing
   Set Lblvazao = Nothing
   Set Lblvazaototal = Nothing
   Set Txtvelodesloca = Nothing
   Set Txtperdas3 = Nothing
   Set Txtdeslocamento = Nothing
   Set Txtprecipitacaolic = Nothing
   Set Txtvazaoporturno = Nothing
   Set Txtalturamanometrica = Nothing
   Set Txtareapordia = Nothing
   Set Txttempo1 = Nothing
   Set Txtareafx = Nothing
   Set Txtfaixasirrigadas = Nothing
   Set Txtturno = Nothing
   Set Txtdiam1 = Nothing
   Set Txtdiam2 = Nothing
   Set Txtdiam3 = Nothing
   Set Txtcoef1 = Nothing
   Set Txtcoef2 = Nothing
   Set Txtcoef3 = Nothing
   Set txtHF1 = Nothing
   Set txtHF2 = Nothing
   Set txtHF3 = Nothing
   Set Txtvelo1 = Nothing
   Set Txtvelo2 = Nothing
   Set Txtvelo3 = Nothing
   Set Txtperdashidro = Nothing
   Set Lblvazaototal2 = Nothing
   Set Txtpressao = Nothing
   Set Txtrendimento = Nothing
   Set Txtpotencia = Nothing
   Set Txtrotacaomotor = Nothing
   Set Txtdemandamotor = Nothing
   Set Txtamperagem = Nothing
   Set Txtconsumo = Nothing
   Set txtFrete = Nothing
   Set txtNF = Nothing
   Set lblAjuste = Nothing
   Set lblOrcamento = Nothing
   Set txtPropriedade = Nothing
   Set txtProjeto = Nothing
   Set lblVinculo = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   Set chkCampo(3) = Nothing
   Set chkCampo(4) = Nothing
   Set chkCampo(5) = Nothing
   Set chkCampo(6) = Nothing
   Set chkCampo(7) = Nothing
   Set chkCampo(8) = Nothing
   Set chkCampo(9) = Nothing
   Set chkCampo(10) = Nothing
   Set chkCampo(11) = Nothing
   If Not Orcamento Is Nothing Then
      Set Orcamento = Nothing
   End If
   If Not Conjuntos_do_Orcamento Is Nothing Then
      Conjuntos_do_Orcamento.CloseRecordset
      Set Conjuntos_do_Orcamento = Nothing
   End If
   If Not Parcelas_Orcamento Is Nothing Then
      Parcelas_Orcamento.CloseRecordset
      Set Parcelas_Orcamento = Nothing
   End If
   If Not Pecas_do_Orcamento Is Nothing Then
      Pecas_do_Orcamento.CloseRecordset
      Set Pecas_do_Orcamento = Nothing
   End If
   If Not Produtos_do_Orcamento Is Nothing Then
      Produtos_do_Orcamento.CloseRecordset
      Set Produtos_do_Orcamento = Nothing
   End If
   If Not Servicos_do_Orcamento Is Nothing Then
      Servicos_do_Orcamento.CloseRecordset
      Set Servicos_do_Orcamento = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   Unload vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmOrcament = Nothing                      'libera o segmento de cdigo do form
End Sub


'evento - quando qq tecla for digitada no grid filho
Private Sub Grid_KeyPress(Index As Integer, ByVal KeyAscii As Integer, ByVal Shift As Integer, vgColumns() As Variant)
   Select Case Index
      Case 0
         ExecutaGrid0 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
      Case 1
         ExecutaGrid1 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
      Case 3
         ExecutaGrid3 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
      Case 4
         ExecutaGrid4 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
   End Select
End Sub


'evento - est mudando a linha selecionada do grid
Private Sub Grid_SkipRecord(Index As Integer, vgColumns() As Variant, ByVal vgBookMark As Variant)
   ExecutaGrid Index, vgColumns(), CONDICOES_ESPECIAIS
End Sub


'evento - aps efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 1
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 2
   GeraLog Me, Grid(Index).Status, Index, False
      Case 3
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 4
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
   End Select
End Sub

'evento - antes de efetuar o edit do recordset do grid
Private Sub Grid_BeforeEditRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, True
   'bloqueia edicao de impostos e financeiro para usuarios nao autorizados
   If vgPWUsuario <> "YGOR" And vgPWUsuario <> "JUCELI" Then
      Select Case Index
         Case 0, 1, 3
            If vgColumn >= 11 And vgColumn <= 31 Then
               vgIsValid = False
               Exit Sub
            End If
         Case 2
            If vgColumn >= 2 And vgColumn <= 5 Then
               vgIsValid = False
               Exit Sub
            End If
         Case 4
            If vgColumn = 5 Then
               vgIsValid = False
               Exit Sub
            End If
      End Select
   End If
   Select Case Index
      Case 0
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 1
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 2
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 3
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 4
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub


'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Conjuntos_do_Orcamento![Seqncia do Oramento] = Orcamento![Seqncia do Oramento]
      Case 1
         Pecas_do_Orcamento![Seqncia do Oramento] = Orcamento![Seqncia do Oramento]
      Case 2
         Parcelas_Orcamento![Seqncia do Oramento] = Orcamento![Seqncia do Oramento]
      Case 3
         Produtos_do_Orcamento![Seqncia do Oramento] = Orcamento![Seqncia do Oramento]
      Case 4
         Servicos_do_Orcamento![Seqncia do Oramento] = Orcamento![Seqncia do Oramento]
   End Select
End Sub


'evento - antes de efetuar o delete no recordset do grid
Private Sub Grid_BeforeDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, True
   ExecutaGrid Index, vgColumns(), EXCLUSOES, vgItem, 0, vgIsValid, 0, vgErrorMessage
   Select Case Index
      Case 0
      Case 1
      Case 2
      Case 3
      Case 4
   End Select
End Sub


'evento - quer pegar valores para cada clula
Private Sub Grid_GetColumnValue(Index As Integer, ByVal vgItem As Long, ByVal vgCol As Integer, vgColumns() As Variant, vgNewText As Variant)
   Dim RetVal As Variant, NCol As Integer
   RetVal = ExecutaGrid(Index, vgColumns(), CONTEUDODACOLUNA, vgItem, vgCol, , NCol)
   If NCol = -1 Then
      vgNewText = RetVal
   End If
End Sub


'evento - recordset do grid foi mudado
Private Sub Grid_RecordSetChanged(Index As Integer, ByVal vgNewRecordSet As GRecordSet)
   Select Case Index
      Case 0
         Set Conjuntos_do_Orcamento = vgNewRecordSet
      Case 1
         Set Pecas_do_Orcamento = vgNewRecordSet
      Case 2
         Set Parcelas_Orcamento = vgNewRecordSet
      Case 3
         Set Produtos_do_Orcamento = vgNewRecordSet
      Case 4
         Set Servicos_do_Orcamento = vgNewRecordSet
   End Select
End Sub


'evento - quer validar dados, est gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub


'evento - pega definio de cores segundo condies
Private Sub Grid_GetColor(Index As Integer, ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Select Case Index
      Case 0
         PegaCoresGrid0 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
      Case 1
         PegaCoresGrid1 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
      Case 3
         PegaCoresGrid3 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
   End Select
End Sub


'evento - aps a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 1
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 3
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 4
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
   End Select
   mdiIRRIG.RemontaForm                                   'vamos atualizar os forms de dados
   Grid(Index).Repaint -1                                 'atualiza dados do grid (registro posicionado)
End Sub


Private Sub Grid_ControlButtonClick(Index As Integer)
   Grid(Index).ShowFilterBar = Not Grid(Index).ShowFilterBar
End Sub


Private Sub Grid_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And Grid(Index).Status = ACAO_NAVEGANDO Then                 'o formulrio pai no est em navegao
      mdiIRRIG.SalvaDados                         'salva o resitro em edio
      If vgSituacao <> ACAO_NAVEGANDO And ActiveControl Is Grid(Index) Then 'se no gravou e ainda est com foco no grid
         FocoNoPriControle Me                                               'vamos colocar foco no primeiro controle do pai
      End If
   End If
End Sub


Private Sub Grid_StatusChanged(Index As Integer, ByVal vgNewStatus As Integer)
   Dim vgTemAltGrdOld As Boolean
   If vgNewStatus <> ACAO_NAVEGANDO Then vgNewStatus = -vgNewStatus
   PrepBotoes Me, vgNewStatus                                     'acerta icones dos botoes
   vgTemAltGrdOld = vgTemAlteracaoGrids
   mdiIRRIG.RemontaForm                                           'remonta dos os form da tela
   If vgSituacao = ACAO_NAVEGANDO Then
      Reposition , Not vgTemAltGrdOld
   End If
End Sub


'evento - atualiza valores para os filtros das colunas do grid filho
Private Sub Grid_GetColumnFilter(Index As Integer, ByVal vgColumn As Integer, vgColumns() As Variant, vgFilter As String)
   vgFilter = ExecutaGrid(Index, vgColumns(), PEGAFILTRODASCOLUNAS, , vgColumn)
End Sub


'evento - pega expresso SQL para abertura de pesquisa
Private Sub Grid_GetColumnSQLSearch(Index As Integer, ByVal vgColumn As Integer, vgColumns() As Variant, vgSQLSearch As String)
   vgSQLSearch = ExecutaGrid(Index, vgColumns(), PEGAEXPRESSAOPESQUISA, , vgColumn)
End Sub


'inicializaes, validaes e processos para o grid
Private Function ExecutaGrid(Index As Integer, ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Select Case Index
      Case 0
         ExecutaGrid = ExecutaGrid0(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 1
         ExecutaGrid = ExecutaGrid1(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 2
         ExecutaGrid = ExecutaGrid2(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 3
         ExecutaGrid = ExecutaGrid3(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case 4
         ExecutaGrid = ExecutaGrid4(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function


'inicializaes, validaes e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single
   Dim Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double
   Dim Valor_do_Tributo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, CFOP As Integer
   Dim CST As Integer, Valor_ICMS_ST As Double, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(0).Status <> ACAO_NAVEGANDO And vgItem = Grid(0).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 4
            vgRetVal = (InfoConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, "Sigla"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 6
            vgRetVal = (InfoConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, "Estoque"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 10
            vgRetVal = (Quantidade * Valor_Unitario)
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 3
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 11
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 12
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 13
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 14
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 15
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 16
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 17
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 18
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 19
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 20
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 21
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 24
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 27
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 28
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosConjuntos KeyCodeAscii, Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = (((Sequencia_do_Pedido = 0) And Cancelado = 0 And Produtos_do_Orcamento.RecordCount = 0 And Venda_Fechada = 0) Or Ordem_Interna = 1)
      On Error Resume Next
         Grid(0).AllowEdit = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0 Or vgPWUsuario = "YGOR")
      On Error Resume Next
         Grid(0).AllowDelete = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Conjuntos do Oramento]"

      'definindo a expresso de ligao com o pai
      x$ = "[Seqncia do Oramento] = " & Orcamento![Seqncia do Oramento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 2
            CST = IIf(Fatura_Proforma, 41, 0)
            vgRetVal = CST
         Case 3
            CFOP = IIf(Fatura_Proforma, 7101, 0)
            vgRetVal = CFOP
         Case 7
            Valor_Unitario = InfoConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, "Valor Unitrio")
            If Grid(0).Col = 7 Then
               vgRetVal = Valor_Unitario
            End If
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia do Conjunto] > 0 AND Inativo = 0", "[Seqncia do Conjunto] > 0") & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Conjuntos.[Seqncia do Conjunto], Conjuntos.Descrio, Conjuntos.[Quantidade no Estoque] FROM Conjuntos WHERE (Conjuntos.[Seqncia do Conjunto] > " + CStr(0) + ") AND " + _
                                                  "(Conjuntos.Inativo = False)"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_Conjunto = 0 Then
            vgIsValid = (Sequencia_do_Conjunto > 0)
            If Not vgIsValid Then vgColumn = 1
            vgErrorMessage$ = "Informe o Conjunto"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Conjunto > 0 Then
               vgIsValid = (ValidaConjunto(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Conjunto INATIVO!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn = 5
            vgErrorMessage$ = "Quantidade invlido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (ValidaConjx(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins))
            If Not vgIsValid Then vgColumn = 7
            vgErrorMessage$ = "Valor Unitrio invlido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Conjuntos_do_Orcamento.Edit
         Set vgRsError = Conjuntos_do_Orcamento
         If ProcessaConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins) Then
            Conjuntos_do_Orcamento![Seqncia do Oramento] = (0)
            Sequencia_do_Orcamento = Conjuntos_do_Orcamento![Seqncia do Oramento]
         End If
         Conjuntos_do_Orcamento.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Conjunto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Quantidade = ColumnValue(5)
   Valor_Unitario = ColumnValue(7)
   Valor_do_Desconto = ColumnValue(8)
   Valor_do_Frete = ColumnValue(9)
   Valor_da_Base_de_Calculo = ColumnValue(11)
   Valor_do_ICMS = ColumnValue(12)
   Valor_do_IPI = ColumnValue(13)
   Aliquota_do_ICMS = ColumnValue(14)
   Aliquota_do_IPI = ColumnValue(15)
   Diferido = ColumnValue(16)
   Percentual_da_Reducao = ColumnValue(17)
   IVA = ColumnValue(18)
   Base_de_Calculo_ST = ColumnValue(19)
   Valor_ICMS_ST = ColumnValue(20)
   Aliquota_do_ICMS_ST = ColumnValue(21)
   Bc_pis = ColumnValue(22)
   Aliq_do_pis = ColumnValue(23)
   Valor_do_PIS = ColumnValue(24)
   Bc_cofins = ColumnValue(25)
   Aliq_do_cofins = ColumnValue(26)
   Valor_do_Cofins = ColumnValue(27)
   Valor_do_Tributo = ColumnValue(28)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Conjuntos_do_Orcamento.EOF = False And Conjuntos_do_Orcamento.BOF = False And Conjuntos_do_Orcamento.RecordCount > 0 Then
         Sequencia_Conjunto_Orcamento = Conjuntos_do_Orcamento![Seqncia Conjunto Oramento]
         Valor_Total = Conjuntos_do_Orcamento![Valor Total]
         Valor_Anterior = Conjuntos_do_Orcamento![Valor Anterior]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Conjuntos_do_Orcamento.EOF = False And Conjuntos_do_Orcamento.BOF = False And Conjuntos_do_Orcamento.RecordCount > 0 Then
      Sequencia_Conjunto_Orcamento = Conjuntos_do_Orcamento![Seqncia Conjunto Oramento]
      Sequencia_do_Conjunto = Conjuntos_do_Orcamento![Seqncia do Conjunto]
      Quantidade = Conjuntos_do_Orcamento!Quantidade
      Valor_Unitario = Conjuntos_do_Orcamento![Valor Unitrio]
      Valor_Total = Conjuntos_do_Orcamento![Valor Total]
      Valor_do_IPI = Conjuntos_do_Orcamento![Valor do IPI]
      Valor_do_ICMS = Conjuntos_do_Orcamento![Valor Do Icms]
      Aliquota_do_IPI = Conjuntos_do_Orcamento![Alquota Do IPI]
      Aliquota_do_ICMS = Conjuntos_do_Orcamento![Alquota Do ICMS]
      Percentual_da_Reducao = Conjuntos_do_Orcamento![Percentual da Reduo]
      Diferido = Conjuntos_do_Orcamento!Diferido
      Valor_da_Base_de_Calculo = Conjuntos_do_Orcamento![Valor da Base de Clculo]
      Valor_do_Tributo = Conjuntos_do_Orcamento![Valor Do Tributo]
      Valor_do_PIS = Conjuntos_do_Orcamento![Valor Do PIS]
      Valor_do_Cofins = Conjuntos_do_Orcamento![Valor Do Cofins]
      IVA = Conjuntos_do_Orcamento!IVA
      Base_de_Calculo_ST = Conjuntos_do_Orcamento![Base de Clculo ST]
      CFOP = Conjuntos_do_Orcamento!CFOP
      CST = Conjuntos_do_Orcamento!CST
      Valor_ICMS_ST = Conjuntos_do_Orcamento![Valor ICMS ST]
      Aliquota_do_ICMS_ST = Conjuntos_do_Orcamento![Alquota Do ICMS ST]
      Valor_do_Desconto = Conjuntos_do_Orcamento![Valor Do Desconto]
      Valor_do_Frete = Conjuntos_do_Orcamento![Valor Do Frete]
      Valor_Anterior = Conjuntos_do_Orcamento![Valor Anterior]
      Bc_pis = Conjuntos_do_Orcamento![Bc Pis]
      Aliq_do_pis = Conjuntos_do_Orcamento![Aliq Do Pis]
      Bc_cofins = Conjuntos_do_Orcamento![Bc Cofins]
      Aliq_do_cofins = Conjuntos_do_Orcamento![Aliq Do Cofins]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function



'inicializaes, validaes e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Valor_do_Tributo As Double, Sequencia_do_Produto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, Aliquota_do_IPI As Double
   Dim Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean
   Dim Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double
   Dim CFOP As Integer, CST As Integer, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   Dim Peso As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(1).Status <> ACAO_NAVEGANDO And vgItem = Grid(1).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 4
            vgRetVal = (InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Sigla"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 5
            vgRetVal = Peso
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 7
            vgRetVal = (InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Estoque"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 8
            vgRetVal = (Peso * Quantidade)
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 12
            vgRetVal = (Quantidade * Valor_Unitario)
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 3
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 13
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 14
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 15
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 16
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 17
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 18
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 19
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 20
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 21
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 22
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 23
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 26
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 29
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case 30
            vgRetVal = Not (Orcamento![Oramento Avulso] = True)
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosPecas KeyCodeAscii, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (((Sequencia_do_Pedido = 0) And Cancelado = 0 And Produtos_do_Orcamento.RecordCount = 0 And Venda_Fechada = 0) Or Ordem_Interna = 1)
      On Error Resume Next
         Grid(1).AllowEdit = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0 Or vgPWUsuario = "YGOR")
      On Error Resume Next
         Grid(1).AllowDelete = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Peas do Oramento]"

      'definindo a expresso de ligao com o pai
      x$ = "[Seqncia do Oramento] = " & Orcamento![Seqncia do Oramento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 2
            CST = IIf(Fatura_Proforma, 41, 0)
            vgRetVal = CST
         Case 3
            CFOP = IIf(Fatura_Proforma, 7101, 0)
            vgRetVal = CFOP
         Case 9
            Valor_Unitario = InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Valor Unitrio")
            vgRetVal = Valor_Unitario
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia do Produto] > 0 AND Inativo = 0", "[Seqncia do Produto] > 0") & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Produtos.[Seqncia do Produto], Produtos.Descrio, Produtos.[Quantidade no Estoque], Produtos.[Cdigo de Barras] FROM Produtos WHERE (Produtos.[Seqncia do Produto] > " + CStr(0) + ") AND " + _
                                                  "(Produtos.Inativo = False)"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_Produto = 0 Then
            vgIsValid = (Sequencia_do_Produto > 0)
            If Not vgIsValid Then vgColumn = 1
            vgErrorMessage$ = "Peas no pode ser Vazio!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Produto > 0 Then
               vgIsValid = (ValidaProduto3(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Impossivel Pea Inativa!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Produto > 0 Then
               vgIsValid = (PodeVenderPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Impossivel Orar Cadastro do Item Incompleto!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn = 6
            vgErrorMessage$ = "Quantidade invlido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (ValidaPecasx(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso))
            If Not vgIsValid Then vgColumn = 9
            vgErrorMessage$ = "Valor Unitrio no pode ser menor que o Valor do Sistema!(Valor Unitrio Invalido)"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Pecas_do_Orcamento.Edit
         Set vgRsError = Pecas_do_Orcamento
         If ProcessaPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso) Then
            Pecas_do_Orcamento![Seqncia do Oramento] = (0)
            Sequencia_do_Orcamento = Pecas_do_Orcamento![Seqncia do Oramento]
         End If
         Pecas_do_Orcamento.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Peso = ColumnValue(5)
   Quantidade = ColumnValue(6)
   Valor_Unitario = ColumnValue(9)
   Valor_do_Desconto = ColumnValue(10)
   Valor_do_Frete = ColumnValue(11)
   Valor_da_Base_de_Calculo = ColumnValue(13)
   Valor_do_ICMS = ColumnValue(14)
   Valor_do_IPI = ColumnValue(15)
   Aliquota_do_ICMS = ColumnValue(16)
   Aliquota_do_IPI = ColumnValue(17)
   Diferido = ColumnValue(18)
   Percentual_da_Reducao = ColumnValue(19)
   IVA = ColumnValue(20)
   Base_de_Calculo_ST = ColumnValue(21)
   Valor_ICMS_ST = ColumnValue(22)
   Aliquota_do_ICMS_ST = ColumnValue(23)
   Bc_pis = ColumnValue(24)
   Aliq_do_pis = ColumnValue(25)
   Valor_do_PIS = ColumnValue(26)
   Bc_cofins = ColumnValue(27)
   Aliq_do_cofins = ColumnValue(28)
   Valor_do_Cofins = ColumnValue(29)
   Valor_do_Tributo = ColumnValue(30)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Pecas_do_Orcamento.EOF = False And Pecas_do_Orcamento.BOF = False And Pecas_do_Orcamento.RecordCount > 0 Then
         Valor_Total = Pecas_do_Orcamento![Valor Total]
         Sequencia_Pecas_do_Orcamento = Pecas_do_Orcamento![Seqncia Peas do Oramento]
         Valor_Anterior = Pecas_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Peso")
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Pecas_do_Orcamento.EOF = False And Pecas_do_Orcamento.BOF = False And Pecas_do_Orcamento.RecordCount > 0 Then
      Valor_do_Tributo = Pecas_do_Orcamento![Valor Do Tributo]
      Sequencia_do_Produto = Pecas_do_Orcamento![Seqncia do Produto]
      Quantidade = Pecas_do_Orcamento!Quantidade
      Valor_Unitario = Pecas_do_Orcamento![Valor Unitrio]
      Valor_Total = Pecas_do_Orcamento![Valor Total]
      Valor_do_IPI = Pecas_do_Orcamento![Valor do IPI]
      Valor_do_ICMS = Pecas_do_Orcamento![Valor Do Icms]
      Sequencia_Pecas_do_Orcamento = Pecas_do_Orcamento![Seqncia Peas do Oramento]
      Aliquota_do_IPI = Pecas_do_Orcamento![Alquota Do IPI]
      Aliquota_do_ICMS = Pecas_do_Orcamento![Alquota Do ICMS]
      Percentual_da_Reducao = Pecas_do_Orcamento![Percentual da Reduo]
      Diferido = Pecas_do_Orcamento!Diferido
      Valor_da_Base_de_Calculo = Pecas_do_Orcamento![Valor da Base de Clculo]
      Valor_do_PIS = Pecas_do_Orcamento![Valor Do PIS]
      Valor_do_Cofins = Pecas_do_Orcamento![Valor Do Cofins]
      IVA = Pecas_do_Orcamento!IVA
      Base_de_Calculo_ST = Pecas_do_Orcamento![Base de Clculo ST]
      Valor_ICMS_ST = Pecas_do_Orcamento![Valor ICMS ST]
      CFOP = Pecas_do_Orcamento!CFOP
      CST = Pecas_do_Orcamento!CST
      Aliquota_do_ICMS_ST = Pecas_do_Orcamento![Alquota Do ICMS ST]
      Valor_do_Desconto = Pecas_do_Orcamento![Valor Do Desconto]
      Valor_do_Frete = Pecas_do_Orcamento![Valor Do Frete]
      Valor_Anterior = Pecas_do_Orcamento![Valor Anterior]
      Bc_pis = Pecas_do_Orcamento![Bc Pis]
      Aliq_do_pis = Pecas_do_Orcamento![Aliq Do Pis]
      Bc_cofins = Pecas_do_Orcamento![Bc Cofins]
      Aliq_do_cofins = Pecas_do_Orcamento![Aliq Do Cofins]
   End If
   Peso = InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Peso")
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid1 = vgRetVal
   vgPriVez = False
End Function



'inicializaes, validaes e processos do grid filho
Private Function ExecutaGrid2(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Descricao As String, Numero_da_Parcela As Integer, Dias As Integer
   Dim Data_de_Vencimento As Variant, Valor_da_Parcela As Double, Descricao_da_Cobranca As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(2).Status <> ACAO_NAVEGANDO And vgItem = Grid(2).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(2).AllowInsert = (((Sequencia_do_Pedido = 0) And Cancelado = 0) And Not Vazio(Orcamento![Forma de Pagamento]) And TotalParcelas < Valor_Total_do_Orcamento)
      On Error Resume Next
         Grid(2).AllowEdit = (((Sequencia_do_Pedido = 0) And Cancelado = 0) And Not Vazio(Orcamento![Forma de Pagamento]))
      On Error Resume Next
         Grid(2).AllowDelete = (((Sequencia_do_Pedido = 0) And Cancelado = 0) And Not Vazio(Orcamento![Forma de Pagamento]))
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas Oramento]"

      'definindo a expresso de ligao com o pai
      x$ = "[Seqncia do Oramento] = " & Orcamento![Seqncia do Oramento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

      'vamos definir a ordenao
      x$ = "[Nmero da Parcela]"
      vgRetVal = InsereSQL(vgRetVal, EXP_ORDERBY, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_Parcela = UltimaParcela
            vgRetVal = Numero_da_Parcela
         Case 2
            Dias = DateDiff("D", Orcamento![Data de Emisso], Data_de_Vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_Vencimento = DateAdd("d", Dias, Orcamento![Data de Emisso])
            vgRetVal = Data_de_Vencimento
         Case 4
            Valor_da_Parcela = Orcamento![Valor Total do Oramento] - TotalParcelas()
            vgRetVal = Valor_da_Parcela
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 5
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia da Cobrana] > 0 AND Inativo = 0", "[Seqncia da Cobrana] > 0") & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (1 = 1)
         If Not vgIsValid Then vgColumn = 2
         vgErrorMessage$ = "Dias invlido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = ((UltimoVencimento) And (IsDate(Data_de_Vencimento) Or Vazio(Data_de_Vencimento)))
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Data de Vencimento Invlida!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Orcamento![Forma de Pagamento] = "Prazo", TotalParcelas() + Valor_da_Parcela <= Orcamento![Valor Total do Oramento], Valor_da_Parcela = Orcamento![Valor Total do Oramento]), IIf(Orcamento![Forma de Pagamento] = "Vista", Valor_da_Parcela = Orcamento![Valor Total do Oramento], Valor_da_Parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_Parcela)) + Valor_da_Parcela <= Orcamento![Valor Total do Oramento])))
            If Not vgIsValid Then vgColumn = 4
            vgErrorMessage$ = "A Soma das Parcelas no Correspondem ao Valor Total do Oramento."
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Descricao_da_Cobranca))
            If Not vgIsValid Then vgColumn = 5
            vgErrorMessage$ = "Descrio da Cobrana no pode ser vazio!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = UltimaParcela
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Numero_da_Parcela = ColumnValue(1)
   Dias = ColumnValue(2)
   Data_de_Vencimento = ColumnValue(3)
   Valor_da_Parcela = ColumnValue(4)
   Descricao_da_Cobranca = Parse$(CStr(ColumnValue(5)), Chr$(1), 1)
   Descricao = ColumnValue(6) & ""
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Parcelas_Orcamento.EOF = False And Parcelas_Orcamento.BOF = False And Parcelas_Orcamento.RecordCount > 0 Then
      Descricao = Parcelas_Orcamento!Descrio
      Numero_da_Parcela = Parcelas_Orcamento![Nmero da Parcela]
      Dias = Parcelas_Orcamento!Dias
      Data_de_Vencimento = Parcelas_Orcamento![Data de Vencimento]
      Valor_da_Parcela = Parcelas_Orcamento![Valor da Parcela]
      Descricao_da_Cobranca = Parcelas_Orcamento![Descrio da Cobrana]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid2 = vgRetVal
   vgPriVez = False
End Function



'inicializaes, validaes e processos do grid filho
Private Function ExecutaGrid3(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Produto_Orcamento As Long
   Dim Sequencia_do_Produto As Long, Quantidade As Double, Valor_Unitario As Double
   Dim Valor_Total As Double, Valor_do_IPI As Double, Valor_do_ICMS As Double
   Dim Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single
   Dim Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double
   Dim Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double
   Dim Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer
   Dim Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, Valor_do_Desconto As Double
   Dim Valor_Anterior As Double, Bc_cofins As Double, Aliq_do_pis As Single
   Dim Aliq_do_cofins As Single
   Dim Peso As Double, PesoTotal As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(3).Status <> ACAO_NAVEGANDO And vgItem = Grid(3).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 2
            vgRetVal = (InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "NCM"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 5
            vgRetVal = (InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Sigla"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 6
            vgRetVal = Peso
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 8
            vgRetVal = (InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Estoque"))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 9
            vgRetVal = PesoTotal
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 13
            vgRetVal = (Round((Quantidade * Valor_Unitario), 2))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosProdutos KeyCodeAscii, Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(3).AllowInsert = (((Sequencia_do_Pedido = 0) And Cancelado = 0 And Conjuntos_do_Orcamento.RecordCount = 0 And Pecas_do_Orcamento.RecordCount = 0 And Venda_Fechada = 0) Or Ordem_Interna = 1)
      On Error Resume Next
         Grid(3).AllowEdit = (((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0) Or Ordem_Interna = 1)
      On Error Resume Next
         Grid(3).AllowDelete = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And Venda_Fechada = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Produtos do Oramento]"

      'definindo a expresso de ligao com o pai
      x$ = "[Seqncia do Oramento] = " & Orcamento![Seqncia do Oramento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 3
            CST = IIf(Fatura_Proforma, 41, 0)
            vgRetVal = CST
         Case 4
            CFOP = IIf(Fatura_Proforma, 7101, 0)
            vgRetVal = CFOP
         Case 10
            Valor_Unitario = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Valor Unitrio")
            vgRetVal = Valor_Unitario
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia do Produto] > 0 AND Inativo = 0", "[Seqncia do Produto] > 0") & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Produtos.[Seqncia do Produto], Produtos.Descrio, Produtos.[Quantidade no Estoque], Produtos.[Cdigo de Barras] FROM Produtos WHERE (Produtos.[Seqncia do Produto] > " + CStr(0) + ") AND " + _
                                                  "(Produtos.Inativo = False)"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_Produto = 0 Then
            vgIsValid = (Sequencia_do_Produto > 0)
            If Not vgIsValid Then vgColumn = 1
            vgErrorMessage$ = "Produto no pode ser Vazio!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Produto > 0 Then
               vgIsValid = (ValidaProduto2(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Impossivel Produto Inativo!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Produto > 0 Then
               vgIsValid = (PodeVenderProd(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Impossivel Orar Cadastro do Item Inclompleto!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Produto > 0 Then
               vgIsValid = (ValidaNCM(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal))
               If Not vgIsValid Then vgColumn = 1
               vgErrorMessage$ = "Pedir para Contabilidade Conferir o (NCM)"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn = 7
            vgErrorMessage$ = "Quantidade invlido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (ValidaProdx(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal))
            If Not vgIsValid Then vgColumn = 10
            vgErrorMessage$ = "Valor Unitrio no pode ser menor que o Valor do Sistema!(Valor Unitrio Invalido)"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(3) = IIf(Fatura_Proforma, 41, 0)
         ColumnValue(4) = IIf(Fatura_Proforma, 7101, 0)
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Produtos_do_Orcamento.Edit
         Set vgRsError = Produtos_do_Orcamento
         If ProcessaProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal) Then
            Produtos_do_Orcamento![Seqncia do Oramento] = (0)
            Sequencia_do_Orcamento = Produtos_do_Orcamento![Seqncia do Oramento]
         End If
         Produtos_do_Orcamento.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(3)
   CFOP = ColumnValue(4)
   Peso = ColumnValue(6)
   Quantidade = ColumnValue(7)
   PesoTotal = ColumnValue(9)
   Valor_Unitario = ColumnValue(10)
   Valor_do_Desconto = ColumnValue(11)
   Valor_do_Frete = ColumnValue(12)
   Valor_da_Base_de_Calculo = ColumnValue(14)
   Valor_do_ICMS = ColumnValue(15)
   Valor_do_IPI = ColumnValue(16)
   Aliquota_do_ICMS = ColumnValue(17)
   Aliquota_do_IPI = ColumnValue(18)
   Diferido = ColumnValue(19)
   Percentual_da_Reducao = ColumnValue(20)
   IVA = ColumnValue(21)
   Base_de_Calculo_ST = ColumnValue(22)
   Valor_ICMS_ST = ColumnValue(23)
   Aliquota_do_ICMS_ST = ColumnValue(24)
   Bc_pis = ColumnValue(25)
   Aliq_do_pis = ColumnValue(26)
   Valor_do_PIS = ColumnValue(27)
   Bc_cofins = ColumnValue(28)
   Aliq_do_cofins = ColumnValue(29)
   Valor_do_Cofins = ColumnValue(30)
   Valor_do_Tributo = ColumnValue(31)
   If Grid(3).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqncia do Produto Oramento]
         Valor_Total = Produtos_do_Orcamento![Valor Total]
         Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
      Bc_pis = Produtos_do_Orcamento![Bc Pis]
      Valor_do_Tributo = Produtos_do_Orcamento![Valor Do Tributo]
      Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqncia do Produto Oramento]
      Sequencia_do_Produto = Produtos_do_Orcamento![Seqncia do Produto]
      Quantidade = Produtos_do_Orcamento!Quantidade
      Valor_Unitario = Produtos_do_Orcamento![Valor Unitrio]
      Valor_Total = Produtos_do_Orcamento![Valor Total]
      Valor_do_IPI = Produtos_do_Orcamento![Valor do IPI]
      Valor_do_ICMS = Produtos_do_Orcamento![Valor Do Icms]
      Aliquota_do_IPI = Produtos_do_Orcamento![Alquota Do IPI]
      Aliquota_do_ICMS = Produtos_do_Orcamento![Alquota Do ICMS]
      Percentual_da_Reducao = Produtos_do_Orcamento![Percentual da Reduo]
      Diferido = Produtos_do_Orcamento!Diferido
      Valor_da_Base_de_Calculo = Produtos_do_Orcamento![Valor da Base de Clculo]
      Valor_do_PIS = Produtos_do_Orcamento![Valor Do PIS]
      Valor_do_Cofins = Produtos_do_Orcamento![Valor Do Cofins]
      IVA = Produtos_do_Orcamento!IVA
      Base_de_Calculo_ST = Produtos_do_Orcamento![Base de Clculo ST]
      Valor_ICMS_ST = Produtos_do_Orcamento![Valor ICMS ST]
      CFOP = Produtos_do_Orcamento!CFOP
      CST = Produtos_do_Orcamento!CST
      Aliquota_do_ICMS_ST = Produtos_do_Orcamento![Alquota Do ICMS ST]
      Valor_do_Frete = Produtos_do_Orcamento![Valor Do Frete]
      Valor_do_Desconto = Produtos_do_Orcamento![Valor Do Desconto]
      Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      Bc_cofins = Produtos_do_Orcamento![Bc Cofins]
      Aliq_do_pis = Produtos_do_Orcamento![Aliq Do Pis]
      Aliq_do_cofins = Produtos_do_Orcamento![Aliq Do Cofins]
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid3 = vgRetVal
   vgPriVez = False
End Function



'inicializaes, validaes e processos do grid filho
Private Function ExecutaGrid4(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_do_Servico_Orcamento As Long, Sequencia_do_Servico As Integer, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_Anterior As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(4).Status <> ACAO_NAVEGANDO And vgItem = Grid(4).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 4
            vgRetVal = (Quantidade * Valor_Unitario)
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
         Case 5
            vgRetVal = (((Quantidade * Valor_Unitario) * (Orcamento![Alquota Do ISS]) / 100))
            vgColumn = -1                   'flag para permitir definio desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosServicos KeyCodeAscii, Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(4).AllowInsert = ((Sequencia_do_Pedido = 0) And Cancelado = 0)
      On Error Resume Next
         Grid(4).AllowEdit = ((Sequencia_do_Pedido = 0) And Cancelado = 0)
      On Error Resume Next
         Grid(4).AllowDelete = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And vgPWUsuario <> "RAFAEL")
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Servios do Oramento]"

      'definindo a expresso de ligao com o pai
      x$ = "[Seqncia do Oramento] = " & Orcamento![Seqncia do Oramento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 3
            Valor_Unitario = InfoServicos(Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior, "Valor Unitrio")
            If Grid(4).Col = 3 Then
               vgRetVal = Valor_Unitario
            End If
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqncia do Servio] > 0 AND Inativo = 0", "[Seqncia do Servio] > 0") & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Servios.[Seqncia do Servio], Servios.Descrio FROM Servios WHERE (Servios.[Seqncia do Servio] > " + CStr(0) + ") AND " + _
                                                  "(Servios.Inativo = False)"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_do_Servico > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Seqncia do Servio invlido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Quantidade invlido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_Unitario > 0)
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Valor Unitrio invlido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValores
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Servicos_do_Orcamento.Edit
         Set vgRsError = Servicos_do_Orcamento
         If ProcessaServicos(Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior) Then
            Servicos_do_Orcamento![Seqncia do Oramento] = (0)
            Sequencia_do_Orcamento = Servicos_do_Orcamento![Seqncia do Oramento]
         End If
         Servicos_do_Orcamento.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Servico = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Quantidade = ColumnValue(2)
   Valor_Unitario = ColumnValue(3)
   If Grid(4).Status <> ACAO_INCLUINDO Then
      If Servicos_do_Orcamento.EOF = False And Servicos_do_Orcamento.BOF = False And Servicos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Servico_Orcamento = Servicos_do_Orcamento![Seqncia do Servio Oramento]
         Valor_Total = Servicos_do_Orcamento![Valor Total]
         Valor_Anterior = Servicos_do_Orcamento![Valor Anterior]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Servicos_do_Orcamento.EOF = False And Servicos_do_Orcamento.BOF = False And Servicos_do_Orcamento.RecordCount > 0 Then
      Sequencia_do_Servico_Orcamento = Servicos_do_Orcamento![Seqncia do Servio Oramento]
      Sequencia_do_Servico = Servicos_do_Orcamento![Seqncia do Servio]
      Quantidade = Servicos_do_Orcamento!Quantidade
      Valor_Unitario = Servicos_do_Orcamento![Valor Unitrio]
      Valor_Total = Servicos_do_Orcamento![Valor Total]
      Valor_Anterior = Servicos_do_Orcamento![Valor Anterior]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid4 = vgRetVal
   vgPriVez = False
End Function


'pega definies de cores para o grid
Private Sub PegaCoresGrid0(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single
   Dim Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double
   Dim Valor_do_Tributo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, CFOP As Integer
   Dim CST As Integer, Valor_ICMS_ST As Double, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(0).GetColumnValues(vgItem)

   Sequencia_do_Conjunto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Quantidade = ColumnValue(5)
   Valor_Unitario = ColumnValue(7)
   Valor_do_Desconto = ColumnValue(8)
   Valor_do_Frete = ColumnValue(9)
   Valor_da_Base_de_Calculo = ColumnValue(11)
   Valor_do_ICMS = ColumnValue(12)
   Valor_do_IPI = ColumnValue(13)
   Aliquota_do_ICMS = ColumnValue(14)
   Aliquota_do_IPI = ColumnValue(15)
   Diferido = ColumnValue(16)
   Percentual_da_Reducao = ColumnValue(17)
   IVA = ColumnValue(18)
   Base_de_Calculo_ST = ColumnValue(19)
   Valor_ICMS_ST = ColumnValue(20)
   Aliquota_do_ICMS_ST = ColumnValue(21)
   Bc_pis = ColumnValue(22)
   Aliq_do_pis = ColumnValue(23)
   Valor_do_PIS = ColumnValue(24)
   Bc_cofins = ColumnValue(25)
   Aliq_do_cofins = ColumnValue(26)
   Valor_do_Cofins = ColumnValue(27)
   Valor_do_Tributo = ColumnValue(28)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Conjuntos_do_Orcamento.EOF = False And Conjuntos_do_Orcamento.BOF = False And Conjuntos_do_Orcamento.RecordCount > 0 Then
         Sequencia_Conjunto_Orcamento = Conjuntos_do_Orcamento![Seqncia Conjunto Oramento]
         Valor_Total = Conjuntos_do_Orcamento![Valor Total]
         Valor_Anterior = Conjuntos_do_Orcamento![Valor Anterior]
      End If
   End If

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 6

         'Fundocoluna
         If (Quantidade > InfoConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, "Estoque")) Then
            vgBackColor = &HC0C0FF
         End If

   End Select

   Exit Sub

DeuErro:
End Sub


'pega definies de cores para o grid
Private Sub PegaCoresGrid1(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Valor_do_Tributo As Double, Sequencia_do_Produto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, Aliquota_do_IPI As Double
   Dim Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean
   Dim Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double
   Dim CFOP As Integer, CST As Integer, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   Dim Peso As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(1).GetColumnValues(vgItem)

   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Peso = ColumnValue(5)
   Quantidade = ColumnValue(6)
   Valor_Unitario = ColumnValue(9)
   Valor_do_Desconto = ColumnValue(10)
   Valor_do_Frete = ColumnValue(11)
   Valor_da_Base_de_Calculo = ColumnValue(13)
   Valor_do_ICMS = ColumnValue(14)
   Valor_do_IPI = ColumnValue(15)
   Aliquota_do_ICMS = ColumnValue(16)
   Aliquota_do_IPI = ColumnValue(17)
   Diferido = ColumnValue(18)
   Percentual_da_Reducao = ColumnValue(19)
   IVA = ColumnValue(20)
   Base_de_Calculo_ST = ColumnValue(21)
   Valor_ICMS_ST = ColumnValue(22)
   Aliquota_do_ICMS_ST = ColumnValue(23)
   Bc_pis = ColumnValue(24)
   Aliq_do_pis = ColumnValue(25)
   Valor_do_PIS = ColumnValue(26)
   Bc_cofins = ColumnValue(27)
   Aliq_do_cofins = ColumnValue(28)
   Valor_do_Cofins = ColumnValue(29)
   Valor_do_Tributo = ColumnValue(30)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Pecas_do_Orcamento.EOF = False And Pecas_do_Orcamento.BOF = False And Pecas_do_Orcamento.RecordCount > 0 Then
         Valor_Total = Pecas_do_Orcamento![Valor Total]
         Sequencia_Pecas_do_Orcamento = Pecas_do_Orcamento![Seqncia Peas do Oramento]
         Valor_Anterior = Pecas_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Peso")

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 7

         'Fundocoluna
         If (Quantidade > InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Estoque")) Then
            vgBackColor = &HC0C0FF
         End If

   End Select

   Exit Sub

DeuErro:
End Sub


'pega definies de cores para o grid
Private Sub PegaCoresGrid3(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Produto_Orcamento As Long
   Dim Sequencia_do_Produto As Long, Quantidade As Double, Valor_Unitario As Double
   Dim Valor_Total As Double, Valor_do_IPI As Double, Valor_do_ICMS As Double
   Dim Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single
   Dim Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double
   Dim Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double
   Dim Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer
   Dim Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, Valor_do_Desconto As Double
   Dim Valor_Anterior As Double, Bc_cofins As Double, Aliq_do_pis As Single
   Dim Aliq_do_cofins As Single
   Dim Peso As Double, PesoTotal As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(3).GetColumnValues(vgItem)

   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(3)
   CFOP = ColumnValue(4)
   Peso = ColumnValue(6)
   Quantidade = ColumnValue(7)
   PesoTotal = ColumnValue(9)
   Valor_Unitario = ColumnValue(10)
   Valor_do_Desconto = ColumnValue(11)
   Valor_do_Frete = ColumnValue(12)
   Valor_da_Base_de_Calculo = ColumnValue(14)
   Valor_do_ICMS = ColumnValue(15)
   Valor_do_IPI = ColumnValue(16)
   Aliquota_do_ICMS = ColumnValue(17)
   Aliquota_do_IPI = ColumnValue(18)
   Diferido = ColumnValue(19)
   Percentual_da_Reducao = ColumnValue(20)
   IVA = ColumnValue(21)
   Base_de_Calculo_ST = ColumnValue(22)
   Valor_ICMS_ST = ColumnValue(23)
   Aliquota_do_ICMS_ST = ColumnValue(24)
   Bc_pis = ColumnValue(25)
   Aliq_do_pis = ColumnValue(26)
   Valor_do_PIS = ColumnValue(27)
   Bc_cofins = ColumnValue(28)
   Aliq_do_cofins = ColumnValue(29)
   Valor_do_Cofins = ColumnValue(30)
   Valor_do_Tributo = ColumnValue(31)
   If Grid(3).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqncia do Produto Oramento]
         Valor_Total = Produtos_do_Orcamento![Valor Total]
         Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 8

         'Fundocoluna
         If (Quantidade > InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Estoque")) Then
            vgBackColor = &HC0C0FF
         End If

   End Select

   Exit Sub

DeuErro:
End Sub


'evento - pega o valor inicial das colunas do grid
Private Sub Grid_GetColumnDefaultValue(Index As Integer, ByVal vgCol As Integer, vgColumns() As Variant, ByRef vgDefaultValue As Variant)
   vgDefaultValue = ExecutaGrid(Index, vgColumns(), DEFAULTDASCOLUNAS, , vgCol)
End Sub


'evento - quer pegar valores para cada clula
Private Sub Grid_GetColumnLocked(Index As Integer, ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   Select Case Index
      Case 0
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 1
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 2
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 3
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 4
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
   End Select
End Sub



'evento - quando o tempo esgotar
Private Sub timUnLoad_Timer()
   timUnLoad.Enabled = False
   vgPodeFazerUnLoad = True
   Unload Me
End Sub


'evento - quando o boto for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         PreValidaNFE
      Case 1
         seqRegistro = Sequencia_do_Vendedor
         mdiIRRIG.MGeral
      Case 2
         seqRegistro = Sequencia_da_Classificacao
         mdiIRRIG.MClassifi
      Case 3
         seqRegistro = IIf(Sequencia_do_Geral > 0, MunicipioAux![Seqncia Do Municpio], Sequencia_do_Municipio)
         mdiIRRIG.MMuni
      Case 4
         seqRegistro = Sequencia_da_Transportadora
         mdiIRRIG.MGeral
      Case 5
         seqRegistro = Sequencia_do_Pais
         mdiIRRIG.MPaises
      Case 6
         AtualizaValor
      Case 7
         Parcelar
      Case 8
         AbreNotaFiscal
      Case 9
         PreValidaImpressao
      Case 10
         seqRegistro = Sequencia_do_Geral
         mdiIRRIG.MGeral
      Case 12
         AbrePropriedades Sequencia_do_Geral, Sequencia_da_Propriedade
      Case 13
         mdiIRRIG.MRelfxhid
         AtivaForm Me
      Case 14
         mdiIRRIG.MRelorcfa
         AtivaForm Me
      Case 15
         AbreProjeto
      Case 16
         mdiIRRIG.MEmiAlmox
         AtivaForm Me
   End Select
End Sub


'evento - quando o boto for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 16
         AbreComando "mailto:" + Email
   End Select
End Sub



'evento - quando o boto for apertado
Private Sub bottxtCampo158_Click(Index As Integer)
   txtCampo(158).SetFocus
   DoEvents
   txtCampo(158).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo156_Click(Index As Integer)
   txtCampo(156).SetFocus
   DoEvents
   txtCampo(156).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo157_Click(Index As Integer)
   txtCampo(157).SetFocus
   DoEvents
   txtCampo(157).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo17_Click(Index As Integer)
   txtCampo(17).SetFocus
   DoEvents
   txtCampo(17).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo16_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 16, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(16).BotClick Index
   End If
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo33_Click(Index As Integer)
   txtCampo(33).SetFocus
   DoEvents
   txtCampo(33).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo55_Click(Index As Integer)
   txtCampo(55).SetFocus
   DoEvents
   txtCampo(55).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo59_Click(Index As Integer)
   txtCampo(59).SetFocus
   DoEvents
   txtCampo(59).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo60_Click(Index As Integer)
   txtCampo(60).SetFocus
   DoEvents
   txtCampo(60).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo34_Click(Index As Integer)
   txtCampo(34).SetFocus
   DoEvents
   txtCampo(34).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo35_Click(Index As Integer)
   txtCampo(35).SetFocus
   DoEvents
   txtCampo(35).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo32_Click(Index As Integer)
   txtCampo(32).SetFocus
   DoEvents
   txtCampo(32).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo45_Click(Index As Integer)
   txtCampo(45).SetFocus
   DoEvents
   txtCampo(45).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo47_Click(Index As Integer)
   txtCampo(47).SetFocus
   DoEvents
   txtCampo(47).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo57_Click(Index As Integer)
   txtCampo(57).SetFocus
   DoEvents
   txtCampo(57).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo62_Click(Index As Integer)
   txtCampo(62).SetFocus
   DoEvents
   txtCampo(62).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo137_Click(Index As Integer)
   txtCampo(137).SetFocus
   DoEvents
   txtCampo(137).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo64_Click(Index As Integer)
   txtCampo(64).SetFocus
   DoEvents
   txtCampo(64).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo99_Click(Index As Integer)
   txtCampo(99).SetFocus
   DoEvents
   txtCampo(99).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo100_Click(Index As Integer)
   txtCampo(100).SetFocus
   DoEvents
   txtCampo(100).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo101_Click(Index As Integer)
   txtCampo(101).SetFocus
   DoEvents
   txtCampo(101).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo147_Click(Index As Integer)
   txtCampo(147).SetFocus
   DoEvents
   txtCampo(147).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo149_Click(Index As Integer)
   txtCampo(149).SetFocus
   DoEvents
   txtCampo(149).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo150_Click(Index As Integer)
   txtCampo(150).SetFocus
   DoEvents
   txtCampo(150).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo161_Click(Index As Integer)
   txtCampo(161).SetFocus
   DoEvents
   txtCampo(161).BotClick Index
End Sub


'evento - quando o boto for apertado
Private Sub bottxtCampo151_Click(Index As Integer)
   txtCampo(151).SetFocus
   DoEvents
   txtCampo(151).BotClick Index
End Sub


'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtCampo(Index).MouseDown
End Sub



