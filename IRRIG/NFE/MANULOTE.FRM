VERSION 5.00
Begin VB.Form frmManuLote
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   7590
   Left                     =   3045
   LinkTopic                =   "ManuLote"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7155
   ScaleWidth               =   12435
   Top                      =   1630
   Width                    =   12525
   Begin NFE.GPainel Painel
      BackColor                =   &H8000000F
      BevelOuter               =   0
      BevelInner               =   0
      BorderWidth              =   1
      BevelWidth               =   1
      BorderStyle              =   2
      Stretch                  =   -1
      Stripes                  =   0
      SaveGridBars             =   0
      Height                   =   6990
      Left                     =   15
      Width                    =   12255
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   21605
      _extenty                 =   12323
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1515
         Left                     =   75
         Width                    =   1440
         Top                      =   75
         Index                    =   0 
         TabStop                  =   0
         ScrollBars               =   0
         Locked                   =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   375
         Left                     =   1695
         Width                    =   6420
         Top                      =   615
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   15.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Manutenção Contas á Receber - Lote"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1755
         Y1                       =   1005
         X2                       =   8175
         Y2                       =   1005
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   255
         Width                    =   840
         Top                      =   1860
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Dt. Inicial:"
         Index                    =   1 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1245
         Top                      =   1830
         Width                    =   1515
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   2490
         Top                      =   1845
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1260
         Width                    =   1215
         Top                      =   1845
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2895
         Width                    =   765
         Top                      =   1860
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Dt. Final:"
         Index                    =   2 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3795
         Top                      =   1830
         Width                    =   1515
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   5040
         Top                      =   1845
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   270
         Left                     =   3810
         Width                    =   1215
         Top                      =   1845
         Index                    =   1 
      End
      Begin NFE.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         CellTextColor            =   &H80000008
         EditTextColor            =   &H80000008
         FilterBackColor          =   8454143
         LeftBackColor            =   -2147483633
         MaxBackColor             =   -2147483633
         AvgBackColor             =   -2147483633
         MinBackColor             =   -2147483633
         SumBackColor             =   -2147483633
         StatusBackColor          =   -2147483633
         ColSelectBackColor       =   -2147483633
         CellBackColor            =   -2147483643
         EditBackColor            =   8454143
         SelectBackColor          =   -2147483635
         StripesBackColor         =   13750737
         BackColor                =   -2147483643
         CaptionBackColor         =   -2147483633
         ColHeaderBackColor       =   -2147483633
         GridLinesColor           =   14737632
         FilterTextColor          =   -2147483640
         AvgTextColor             =   16512
         MinTextColor             =   255
         MaxTextColor             =   49152
         StatusTextColor          =   -2147483630
         SumTextColor             =   16711680
         ColTextBackColor         =   12582912
         CellTextColor            =   -2147483640
         SelectTextColor          =   -2147483634
         EditTextColor            =   -2147483640
         StripesTextColor         =   -2147483630
         CaptionTextColor         =   -2147483630
         ColHeaderTextColor       =   -2147483630
         TabIndex                 =   5
         TabStop                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   3630
         Left                     =   195
         Width                    =   11880
         Top                      =   3000
         Caption                  =   ""
         ShowGridCaption          =   0
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5460
         Width                    =   795
         Top                      =   1860
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Especie:"
         Index                    =   3 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6420
         Top                      =   1830
         Width                    =   1620
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7770
         Top                      =   1845
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6435
         Width                    =   1320
         Top                      =   1845
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   435
         Width                    =   660
         Top                      =   2400
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Cliente:"
         Index                    =   4 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1260
         Top                      =   2400
         Width                    =   6795
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   7785
         Top                      =   2415
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7530
         Top                      =   2415
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1275
         Width                    =   6240
         Top                      =   2415
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8160
         Width                    =   375
         Top                      =   1860
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "O.S:"
         Index                    =   5 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8700
         Top                      =   1830
         Width                    =   1080
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   9510
         Top                      =   1845
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   1
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   270
         Left                     =   8715
         Width                    =   780
         Top                      =   1845
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   12060
         Width                    =   45
         Top                      =   6960
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "."
         Index                    =   6 
      End
   End
End
Attribute VB_name = "frmManuLote"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: ManuLote
'* Função....: Manutenção Contas - Lote
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:54
'* * * * * * *

Option Explicit
DefInt A-Z

Public vgSituacao As Integer                      'situação de edição que do módulo
Public vgCaracteristica As Integer                'caracteristica do módulo
Public vgTipo As Integer                          'tipo do módulo
Public vgPriVez As Integer                        'flag de carregamento do módulo  
Public vgFormID As Long                           'identificador único para o módulo 
Public vgIdentTab As String                       'nome da tabela principal do módulo 
Public vgEmBrowse As Integer                      'flag se o módulo esta em grade 
Public vgTemFiltro As Integer                     'flag se tem ou não filtro no módulo
Public vgCancelou As Integer                      'flag cancela ou nao o processo do módulo  
Public vgQueryAtual As String                     'expressão SQL definada para o módulo 

Public vgFP As New frmPreview                     'formuário de preview do relatório 
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFG As frmGauge                           'formulário do gauge 
Public vgQReg As Long                             'quantidade de já feira para o gauge  

Dim txtCampo(4) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim Dt_Inicial As Variant, Dt_Final As Variant, Especie As String
Dim Geral As Double, Ordem As Double
Public grdContas As Object, lblAjuste As Object

Private Sub CarregaTotalizador()
   On Error Resume Next

   With grdContas
      .RowHeight = 315
      .AddControlIgnoreFocus mdiNFe.botCancela.hWnd             'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFe.botSalva.hWnd               'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = False
      .AllowEdit = True
      .AllowDelete = False
      '.AddColumn Nothing, , "Imp.", "Imprimir", TP_LOGICO, , , False, , , , , , , , , , "0", , , , "0", 1, "0", 405 
      .AddColumn Nothing, , "Conta", "Sequencia da manutenção", TP_NUMERICO, "9999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 680     
      .AddColumn Nothing, , "Especie", "Especie de Pagamento", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1100
      .AddColumn Nothing, , "Cliente", "Razão Social", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 4000
      .AddColumn Nothing, , "Dt. Emissão", "Data de Entrada", TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Documento", "Documento", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Pc.", "Parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 390
      .AddColumn Nothing, , "Vr. Parcela", "Valor da Parcela", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Dt. Vencto", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "2", 1345
      .AddColumn Nothing, , "Cobrança", "Sequencia da Cobrança", TP_NUMERICO, "", , False, , "DADOSNFE", "Tipo da Cobrança", "Sequencia da Cobrança", "Descrição da Cobrança; Sequencia da Cobrança", "Descrição da Cobrança; Sequencia da Cobrança", "Descrição da Cobrança", "", , "1", "Tipo da Cobrança.[Sequencia da Cobrança]", "", "DADOSNFE", "18", 1, "0", 1600
      .AddColumn Nothing, , "Vr. Restante", "Valor Restante", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .ShowSumBar = True
      .ShowSumCol(.Columns("Parcela").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With      
End Sub

Private Sub CarregaContas()
   Dim x As String, Tb As New GRecordSet, MsgValida As String
   
   On Error GoTo DeuErro
   
   x$ = "SELECT [Especie de Pagamento], [Sequencia da Cobrança], Geral.[Sequencia do Geral], [Razão Social], Documento, Parcela, [Forma de Pagamento], [Data de Entrada], [Data de Vencimento], [Valor Restante], [Valor da Parcela], [Sequencia da Manutenção] " & _
        "FROM [Manutenção Contas Receber] INNER JOIN Geral ON [Manutenção Contas Receber].[Sequencia do Geral] = Geral.[Sequencia do Geral] WHERE " & Filtro() & " ORDER BY [Data de Vencimento], [Sequencia da Manutenção]"
      
   Set Tb = vgDb.OpenRecordSet(x$)
   
   MsgValida$ = Executar(VALIDACOES)
   If MsgValida$ <> "" Then MsgBox MsgValida$, vbCritical + vbOKOnly, vaTitulo: Exit Sub
   'ExecutaPreValidacao
   
   With grdContas
      .OpenRecordSet x$, CURSOR_TABLE
   End With
   
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub   

Private Function Filtro() As String
   Filtro = "1 = 1"
   Filtro = Filtro & " AND [Valor Restante] > 0"
   If Not Vazio(Dt_Inicial) And Not Vazio(Dt_Final) Then
      Filtro = Filtro & " AND [Data de Vencimento] >= " & D(Dt_Inicial) & " AND [Data de Vencimento] <= " & D(Dt_Final)
   End If   
   If Geral > 0 Then
      Filtro = Filtro & " AND [Manutenção Contas Receber].[Sequencia do Geral] = " & Geral
   End If
   If Ordem > 0 Then
      Filtro = Filtro & " AND [Manutenção Contas Receber].[Sequencia da Ordem] = " & Ordem
   End If
   If Especie <> "" Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Especie de pagamento] = '" & Especie & "'"
   End If
End Function

Private Sub Grid_GotFocus(Index As Integer)
  CarregaContas 
End Sub


Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal Row As Long, Columns() As Variant, IsValid As Boolean, Column As Integer, ErrorMessage As String)
Dim Vencto As Variant, Coluna As Variant
 
 Vencto = Columns(8)'grdContas.ColumnValue(grdContas.Row + 1, grdContas.Col)
 IsValid = ((Not Vazio(Vencto)) And (IsDate(Vencto) Or Vazio(Vencto)))
 
 If Not IsValid Then
    MsgBox "Atenção!! Data de Vencimento Invalida!", vbExclamation + vbOKOnly, vaTitulo
    MDINFE.CancelaAlteracoes
    GrdContas.Rebind  
 End If
 
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Dt_Inicial = txtCampo(0).Value
   Dt_Final = txtCampo(1).Value
   Especie = txtCampo(2).Value
   Geral = txtCampo(3).Value
   Ordem = txtCampo(4).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 3
         txtCampo(3).Filter = "[Sequencia do geral] > " & 0 & ""
                           txtCampo(3).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro, Geral.Cep, Geral.Fone FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ")"
      Case 4
         txtCampo(4).Filter = "(([Sequencia da Ordem] > " & 0 & ") AND Fechado = " & 1 & ") AND Cancelado = " & 0 & ""
   End Select
End Sub

'executa processos/validacoes nos campos do formuário
Public Function Executar(vgOq As String) As String
   Dim vgMsg As String, vgOk As Integer, vgNVez As Integer
   On Error Goto DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
         ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
   End If
   Executar = vgMsg$                              'prepara saida da função
   Exit Function                                  'e cai fora...

DeuErro:
   Select Case Err                                'vamos verificar se deu algum erro
      Case 3197                                   'registro acabou de ser alterado por outro usuário
         Resume
      Case 3046, 3158, 3186, 3187, 3188, 3189, 3218, 3260
         vgNVez = vgNVez + 1                      'outro usuário acessando
         If vgNVez <= 10 Then                     'vamos tentar 10 vezes
            Delay 0.5                             'mais antes espera 1/2 segundo
            Resume                                'antes de tentar
         End If
   End Select
   Executar = Err.Source + "|" + Trim$(str$(Err)) + "|" + Error$ 'não teve jeito o erro não pode ser evitado...
End Function

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(5).Visible = (Empresa![Nome fantasia]  = "Reticom Virabrequim")
   txtCampo(4).Visible = (Empresa![Nome fantasia]  = "Reticom Virabrequim")
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Or Index = 3 Or Index = 4 Then
      ExecutaVisivel
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 3
            PoeRelEFiltroCbo 3
         Case 4
            PoeRelEFiltroCbo 4
      End Select
   txtCampo(Index).GotFocus
End Sub


'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   txtCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)
   txtCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
   txtCampo(Index).LostFocus                      'executa LostFocus na classe
   ExecutaVisivel                                 'torna camos visiveis
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(32880)
   vgFormID = 699
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(24385))
   Set grdContas = Grid(0)
   Set lblAjuste = Label(6)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(32885)
   Label(1).Caption = LoadGasString(32886)
   Label(2).Caption = LoadGasString(32887)
   Label(3).Caption = LoadGasString(32888)
   Label(4).Caption = LoadGasString(32889)
   Label(5).Caption = LoadGasString(32890)
   Label(6).Caption = LoadGasString(32891)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
   CarregaTotalizador  
   Screen.MousePointer = vbDefault
   vgPriVez = False
   Exit Sub

DeuErro:
   CErr.NumErro = Err
   CErr.FunctionName = "Form_Load"
   CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
   CErr.Show
End Sub


'evento - quando o formulário receber o foco
Private Sub Form_Activate()
   Dim i As Integer
   AtivaForm Me
End Sub

'evento - redefinido o tamanho do formulário
Private Sub Form_Resize()
   AjustaPanFundo Me
End Sub

'define as propriedades das classe dos campos do formuário
Public Sub DefineControles()
 On Error GoTo DeuErro

   Set txtCampo(0).CtPri = txtCp(0)
   Set txtCampo(0).CtFdo = labFdo0
   Set txtCampo(0).CtBot(BOT_ACAO) = bottxtCampo0(BOT_ACAO)
   Set bottxtCampo0(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(0).DataType = 2
   txtCampo(0).Mask = "99/99/9999"
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_ACAO) = bottxtCampo1(BOT_ACAO)
   Set bottxtCampo1(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(1).DataType = 2
   txtCampo(1).Mask = "99/99/9999"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 0
   txtCampo(2).ListFields = "Dinheiro|Cartão|Cheque - Pre|Boleto"

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = ""
   txtCampo(4).BoundColumn = "Sequencia da Ordem"
   txtCampo(4).ListFields = "Sequencia da Ordem"
   txtCampo(4).OrderFields = "Sequencia da Ordem-"
   txtCampo(4).Relation = ""
   txtCampo(4).Source = "Ordem de Serviço"
   txtCampo(4).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_LISTA) = bottxtCampo3(BOT_LISTA)
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   bottxtCampo3(BOT_LISTA).Caption = "P"
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = ""
   txtCampo(3).PesqModoAbertura =  2 
   txtCampo(3).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(3).BoundColumn = "Sequencia do geral"
   txtCampo(3).ListFields = "Razão social"
   txtCampo(3).OrderFields = "Razão social"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Geral"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set grdContas = Nothing
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmManuLote = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then                 'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo0_Click(Index As Integer)
   txtCampo(0).SetFocus
   DoEvents
   txtCampo(0).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo4_Click(Index As Integer)
   txtCampo(4).SetFocus
   DoEvents
   txtCampo(4).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

