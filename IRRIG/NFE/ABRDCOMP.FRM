VERSION 5.00
Begin VB.Form frmAbrDComp
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   3480
   Left                     =   4530
   LinkTopic                =   "AbrDComp"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   3045
   ScaleWidth               =   7200
   Top                      =   1645
   Width                    =   7290
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
      Height                   =   2880
      Left                     =   15
      Width                    =   7020
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   12376
      _extenty                 =   5077
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1455
         Left                     =   75
         Width                    =   1410
         Top                      =   75
         Index                    =   0 
         TabStop                  =   0
         ScrollBars               =   0
         Locked                   =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   435
         Left                     =   1695
         Width                    =   2940
         Top                      =   495
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   18
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Gerar Devolução"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1755
         Y1                       =   945
         X2                       =   4650
         Y2                       =   945
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6915
         Width                    =   45
         Top                      =   3015
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
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   300
         Left                     =   195
         Width                    =   4005
         Top                      =   1755
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Informe o Movimento de Entrada:"
         Index                    =   2 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4455
         Top                      =   1755
         Width                    =   2370
         Height                   =   315
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6555
         Top                      =   1770
         Width                    =   255
         Height                   =   285
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6300
         Top                      =   1770
         Width                    =   255
         Height                   =   285
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
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   285
         Left                     =   4470
         Width                    =   1815
         Top                      =   1770
         Index                    =   0 
      End
      Begin NFE.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   1
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   360
         Left                     =   5205
         Width                    =   1620
         Top                      =   2355
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Prossegue"
      End
   End
         Begin VB.Timer timUnLoad
                  Enabled         =   0
                  Interval        =   5
                  left            =   30
                  top             =   950
         End
End
Attribute VB_name = "frmAbrDComp"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: AbrDComp
'* Função....: Devolução
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:40
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

Dim txtCampo(0) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim SeqMovimento As Double
Public lblAjuste As Object

Private Sub SuperDevolucao()
 Dim Produtos As New GRecordSet, Movimento As New GRecordSet
 Dim Devolucao As New GRecordSet, UltimaDevolucao As New GRecordSet
 Dim ItensDevolucao As New GRecordSet, SeqItem As Long
 Dim Totais As New GRecordSet, TotaisDevolucao As New GRecordSet
 Dim TotalMovimentoDev As New GRecordSet
 Dim TotalMovimento As New GRecordSet
 
  On Error GoTo DeuErro
   
  vgDb.BeginTrans
 
  Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & SeqMovimento)
  Set Movimento = vgdb.OpenRecordset("SELECT * FROM [Movimento do estoque] WHERE [Sequencia do movimento] = " & SeqMovimento) 
  Set Devolucao = vgDb.OpenRecordSet("Devolução de compra") 
  Set ItensDevolucao = vgDb.OpenRecordSet("Produtos devolução compra") 
  Set UltimaDevolucao = vgDb.OpenRecordSet("SELECT MAX([Sequencia da devolução]) Seq From [Devolução de compra]")
  Set TotalMovimentoDev = vgdb.OpenRecordset("SELECT SUM([Valor total]) Total FROM [Produtos devolução compra] LEFT JOIN [Devolução de compra] ON [Produtos devolução compra].[Sequencia da devolução] = [Devolução de compra].[Sequencia da devolução] Where [Devolução de compra].Cancelado = 0 And [Produtos devolução compra].[Sequencia do movimento] = " & SeqMovimento)
  Set TotalMovimento = vgdb.OpenRecordset("SELECT SUM([Valor total]) Total FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & SeqMovimento)
  
  If Produtos.RecordCount = 0 Then
     MsgBox("Impossivel Movimento de Entrada sem Itens")
     Exit Sub
  End If
  
  If TotalMovimento!Total = TotalMovimentoDev!Total Then
     MsgBox("Impossivel Já foi feito Devolução Total desse Movimento")
     Exit Sub
  End If
  
  SeqItem = 1
 
  'Processando...
   With Movimento
    'Devolução 
     Devolucao.AddNew
     Devolucao![Sequencia da devolução] = UltimaDevolucao!Seq + 1
     Devolucao![Sequencia Do movimento] = ![Sequencia Do movimento]
     Devolucao![Data da devolução] = Date
     Devolucao![Data da entrada] = ![Data de entrada]
     Devolucao![Sequencia Do geral] = ![Sequencia Do geral]
     Devolucao![Forma de pagamento] = "Prazo"
     Devolucao![Data da alteração] = Date
     Devolucao![Hora da alteração] = Time
     Devolucao![Usuario da alteração] = vgPwUsuario
     Devolucao.Update
     Devolucao.BookMark = .LastModified
     Devolucao.MoveLast
   End With
    'Itens da Devolução
     Do While Not Produtos.EOF
     Set Totais = vgDb.OpenRecordSet("SELECT SUM([Quantidade])Qtd From [Produtos movimento estoque] Where [Sequencia do movimento] = " & SeqMovimento & " And [Sequencia do produto] = " & Produtos![Sequencia Do produto])
     Set TotaisDevolucao = vgDb.OpenRecordSet("SELECT SUM([Produtos devolução compra].[Quantidade])Qtd From " & _
                                              "[Produtos devolução compra] Where [Produtos devolução compra].[Sequencia do produto] = " & Produtos![Sequencia Do produto] & " And [Produtos devolução compra].[Sequencia do movimento] = " & SeqMovimento)     
     ItensDevolucao.AddNew
     ItensDevolucao![Sequencia da devolução] = UltimaDevolucao!Seq + 1
     ItensDevolucao![Sequencia Do movimento] = SeqMovimento
     ItensDevolucao![Sequencia Do Item] = Produtos![Sequencia Do item]
     ItensDevolucao![Sequencia Do produto] = Produtos![Sequencia Do produto]
     ItensDevolucao!Quantidade = 0
     ItensDevolucao![Valor Unitario] = Produtos![Valor Unitario]
     ItensDevolucao![Valor Total] = ItensDevolucao!Quantidade * Produtos![Valor Unitario]
     ItensDevolucao![Valor Do ipi] = Round(Round(ItensDevolucao!Quantidade * Produtos![Valor Unitario], 2) * Produtos![Aliquota Do ipi] / 100, 2)
     'ItensDevolucao![Valor icms st] = Produtos![Valor icms st]
     ItensDevolucao![Aliquota Do ipi] = Produtos![Aliquota Do ipi]
     ItensDevolucao.Update
     ItensDevolucao.BookMark = Produtos.LastModified
     Produtos.MoveNext
     SeqItem = SeqItem + 1
     Loop
     Produtos.MoveFirst
  
     vgDb.CommitTrans
     Unload Me
     frmDevCompr.Show
     frmDevCompr.SetFocus
     'PosicionaRegistro frmDevCompr, "Sequencia da devolução", UltimaDevolucao!Seq + 1
     frmDevCompr.vgFiltroInicial = "[Sequencia da Devolução] = " & UltimaDevolucao!Seq + 1
     InicializaFiltro frmDevCompr
     frmDevCompr.AjustaValores
     frmDevCompr.Reposition
       
 DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
              
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   SeqMovimento = txtCampo(0).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "(([Sequencia do movimento] > " & 0 & ") AND Cancelado = " & 0 & ") AND Fechado = " & 1 & ""
                           txtCampo(0).PesqSQLExpression = "SELECT [Movimento do estoque].[Sequencia do movimento], [Movimento do estoque].Documento, [Movimento do estoque].[Data de entrada], " + _
                                                              "[Movimento do estoque].[Valor total dos produtos], [Movimento do estoque].[Valor total do ipi], [Movimento do estoque].[Valor total do movimento], " + _
                                                              "Geral.[Razão social] FROM [Movimento do estoque], Geral WHERE ([Movimento do estoque].[Sequencia do movimento] > " + CStr(0) + ") AND " + _
                                                              "([Movimento do estoque].[Sequencia do geral] = Geral.[Sequencia do geral])"
   End Select
End Sub

'executa processos/validacoes nos campos do formuário
Public Function Executar(vgOq As String) As String
   Dim vgMsg As String, vgOk As Integer, vgNVez As Integer
   On Error Goto DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = (SeqMovimento > 0)
      vgMsg$ = "Sequencia do Movimento Invalido!"
      If Not vgOk Then txtCampo(0).SetFocus
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
         ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
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

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 0
            PoeRelEFiltroCbo 0
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
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(24000)
   vgFormID = 654
   vgIdentTab = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(15505))
   Set botOk.Picture = LoadPicture(LoadGasPicture(15506))
   Set lblAjuste = Label(1)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(24005)
   Label(1).Caption = LoadGasString(24006)
   Label(2).Caption = LoadGasString(24007)
   vgTooltips.AddTool botOk, 0, LoadGasString(24008)
   botOk.Caption = LoadGasString(24009)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
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
   Set txtCampo(0).CtBot(BOT_LISTA) = bottxtCampo0(BOT_LISTA)
   Set txtCampo(0).CtBot(BOT_COMBO) = bottxtCampo0(BOT_COMBO)
   bottxtCampo0(BOT_LISTA).Caption = "P"
   Set bottxtCampo0(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = ""
   txtCampo(0).PesqModoAbertura =  2 
   txtCampo(0).PesqFieldCapture = "[Movimento do estoque].[Sequencia do movimento]"
   txtCampo(0).BoundColumn = "Sequencia do movimento"
   txtCampo(0).ListFields = "Sequencia do movimento"
   txtCampo(0).OrderFields = "Sequencia do movimento-"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Movimento do estoque"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub

'evento - valida digitações e prossegue a execução
Private Sub botOk_Click()
   Dim k As String
   If vgPriVez Then Exit Sub
   vgPriVez = True
   k$ = Executar(VALIDACOES)                      'faz validaçao dos campos
   If Len(k$) > 0 Then                            'EPA!... tem campo errado
      If Instr(k$, "|") > 0 Then
         Err.Source = Parse$(k$, "|")
         Err.Number = Val(Parse$(k$, "|"))        'deu erro
         Err.Description = Parse$(k$, "|")
         CErr.NumErro = Err.Number
         CErr.Show
      Else
         Beep                                     'manda aviso sonoro
         MsgBox k$, vbCritical, vgAtencao$        'mostra mensagem do erro
      End If
      Screen.MousePointer = vbDefault             'reseta o pointer do mouse
      vgPriVez = False
      Exit Sub                                    'e sai para continuar a ediçao ou inclusao
   End If                                         'OK. os campos estao certos
   SuperDevolucao  
   vgPriVez = False
End Sub

'evento - antes de descarregar o formulário
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If vgPodeFazerUnLoad = False Then
      timUnLoad.Enabled = True
      Cancel = True
   End If
   Cancel = Cancel Or(Not vgFG Is Nothing)
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmAbrDComp = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then                 'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'evento - quando o tempo esgotar  
Private Sub timUnLoad_Timer()
   timUnLoad.Enabled = False
   vgPodeFazerUnLoad = True
   Unload Me
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo0_Click(Index As Integer)
   txtCampo(0).SetFocus
   DoEvents
   txtCampo(0).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

