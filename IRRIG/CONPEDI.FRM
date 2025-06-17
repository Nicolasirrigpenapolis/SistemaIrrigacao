VERSION 5.00
Begin VB.Form frmConpedi
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8775
   Left                     =   4140
   LinkTopic                =   "Conpedi"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8340
   ScaleWidth               =   11850
   Top                      =   1420
   Width                    =   11940
   Begin IRRIG.GPainel Painel
      BackColor                =   &H8000000F
      BevelOuter               =   0
      BevelInner               =   0
      BorderWidth              =   1
      BevelWidth               =   1
      BorderStyle              =   2
      Stretch                  =   -1
      Stripes                  =   0
      SaveGridBars             =   0
      Height                   =   8160
      Left                     =   15
      Width                    =   11670
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   20574
      _extenty                 =   14386
      Index                    =   0 
      Begin IRRIG.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   -1
         Height                   =   1710
         Left                     =   135
         Width                    =   1545
         Top                      =   195
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
         Height                   =   435
         Left                     =   1815
         Width                    =   6330
         Top                      =   195
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   18
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Controle de Pedidos por Cliente"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1860
         Y1                       =   630
         X2                       =   8130
         Y2                       =   630
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1935
         Width                    =   2880
         Top                      =   795
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Digite o Numero do Pedido:"
         Index                    =   1 
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
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   2
         TabStop                  =   0
         Height                   =   300
         Left                     =   4935
         Width                    =   1200
         Top                      =   795
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   435
         Left                     =   1815
         Width                    =   3285
         Top                      =   1095
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   18
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Nome do Cliente"
         Index                    =   2 
      End
      Begin IRRIG.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Arial"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9
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
         EditBackColor            =   16777152
         SelectBackColor          =   -2147483635
         StripesBackColor         =   14737632
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
         TabIndex                 =   3
         TabStop                  =   -1
         Stripes                  =   -1
         SaveGridBars             =   1
         Height                   =   3990
         Left                     =   135
         Width                    =   11325
         Top                      =   3960
         Caption                  =   "Produtos"
         ShowGridCaption          =   -1
         Index                    =   0 
         Tag                      =   "1"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   11580
         Width                    =   45
         Top                      =   8280
         Alignment                =   0
         BeginProperty Font
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "."
         Index                    =   3 
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
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   300
         Left                     =   6195
         Width                    =   705
         Top                      =   795
         TabStop                  =  0
         Index                    =   1 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   5
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   10245
         Width                    =   1170
         Top                      =   180
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Receita"
         Index                    =   0 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1800
         Top                      =   1800
         Width                    =   975
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   2505
         Top                      =   1815
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
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   270
         Left                     =   1815
         Width                    =   675
         Top                      =   1815
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1800
         Width                    =   705
         Top                      =   1560
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
         Caption                  =   "Produto"
         Index                    =   4 
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
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2820
         Width                    =   6060
         Top                      =   1800
         TabStop                  =  0
         Index                    =   3 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   6
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   9015
         Width                    =   1170
         Top                      =   180
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Consultar"
         Index                    =   1 
      End
      Begin IRRIG.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Arial"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9
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
         EditBackColor            =   16777152
         SelectBackColor          =   -2147483635
         StripesBackColor         =   14737632
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
         TabIndex                 =   7
         TabStop                  =   -1
         Stripes                  =   -1
         SaveGridBars             =   1
         Height                   =   1710
         Left                     =   135
         Width                    =   11325
         Top                      =   2160
         Caption                  =   "Conjuntos"
         ShowGridCaption          =   -1
         Index                    =   1 
         Tag                      =   "1"
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   8
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   480
         Left                     =   9015
         Width                    =   2385
         Top                      =   675
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Entregue ao Cliente"
         Index                    =   2 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   9
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   375
         Left                     =   8970
         Width                    =   2130
         Top                      =   1740
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Atualiza Receita"
         Index                    =   3 
      End
   End
End
Attribute VB_name = "frmConpedi"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: Conpedi
'* Função....: Controle de Pedidos (Vendas)
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 15:00:03
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

Dim txtCampo(3) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim Orcamento As Double, NroNFe As Double, Prodstatus As Double
Public Txtorc As Object, grdContas As Object, Lblajuste As Object
Public Txtnfe As Object, txtProdstatus As Object, Grdcontas1 As Object

Private Function Cliente() As String
 Dim Pedido As New GRecordSet, Geral As New GRecordSet
 
   Set Pedido = VgDb.OpenRecordSet("Select [Seqüência do Geral] Seq From Orçamento Where [Seqüência do Orçamento] = " & Orcamento)
   Set Geral = VgDb.OpenRecordSet("Select [Razão Social] From Geral Where [Seqüência do Geral] = " & Pedido!Seq)
   Cliente = Geral![Razão Social]
   
End Function


Private Sub CarregaTotalizador()
   On Error Resume Next

    With grdContas
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = False
      .AllowEdit = True
      .AllowDelete = False
      .AddColumn Nothing, , "Nº", "Sequencia do Item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 700
      .AddColumn Nothing, , "Cod", "Seqüência do Produto", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 700
      .AddColumn Nothing, , "Produto", "Descrição", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 5820
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245 
      .AddColumn Nothing, , "Andamento", "Situação", TP_CARACTER, "@x", , False , , , , , , "Á Produzir|Separado p/ Entregar|Entregue ao Cliente", , , , "1", , , , "0", 1, "0", 2300, , 12648447     
      .ShowSumCol(.Columns("Quantidade").Index) = False
      .ShowFilterBar = True
      .HideStatus = True
   End With
   'Conj
   With grdContas1
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .AllowInsert = False
      .AllowEdit = True
      .AllowDelete = False
      .AddColumn Nothing, , "Nº", "Sequencia do Item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 700
      .AddColumn Nothing, , "Cod", "Seqüência do Conjunto", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 700
      .AddColumn Nothing, , "Conjunto", "Descrição", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 5820
      .AddColumn Nothing, , "Qtde", "Quantidade", TP_NUMERICO, "99.999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245 
      .AddColumn Nothing, , "Andamento", "Situação", TP_CARACTER, "@x", , False , , , , , , "Á Produzir|Á Galnanizar|Na Galvanização|Separado p/ Entregar|Falta MPrima|Entregue ao Cliente|Separado p/ Montar", , , , "1", , , , "0", 1, "0", 2300, , 12648447     
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowFilterBar = True
      .HideStatus = True
   End With
      
End Sub


Private Sub CarregaContas()
   Dim x As String, Tb As New GRecordSet, MsgValida As String
   
   If Orcamento = 0 Then
      Exit Sub
   End If   
   
   Screen.MousePointer = vbHourglass
   
   On Error GoTo DeuErro
   BuscaNota
   MYSQL
   AtuPedido
   
   x$ = "SELECT [Seqüência do Orçamento], [Itens Pendentes].[Seqüência do Produto], [Sequencia do Item], Descrição, Quantidade, [Situação] " & _
        "FROM [Itens Pendentes] LEFT JOIN Produtos ON [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] WHERE " & Filtro()
      
   Set Tb = vgDb.OpenRecordSet(x$)
   
  ' If Tb.RecordCount > 0 Then TemRecordset = True
   MsgValida$ = Executar(VALIDACOES)
   If MsgValida$ <> "" Then MsgBox MsgValida$, vbCritical + vbOKOnly, vaTitulo: Exit Sub
   ExecutaPreValidacao
   
   With grdContas
      .OpenRecordSet x$, CURSOR_TABLE
   End With
   CarregaContas2
   Screen.MousePointer = vbDefault
   
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub 

Private Sub CarregaContas2()
   Dim x As String, Tb As New GRecordSet, MsgValida As String
   
   Screen.MousePointer = vbHourglass
   
   On Error GoTo DeuErro
  
   
   x$ = "SELECT [Seqüência do Orçamento], [Itens Pendentes].[Seqüência do Conjunto], [Sequencia do Item], Descrição, Quantidade, [Situação] " & _
        "FROM [Itens Pendentes] LEFT JOIN Conjuntos ON [Itens Pendentes].[Seqüência do Conjunto] = Conjuntos.[Seqüência do Conjunto] WHERE " & Filtro5()
      
   Set Tb = vgDb.OpenRecordSet(x$)
   
  ' If Tb.RecordCount > 0 Then TemRecordset = True
   MsgValida$ = Executar(VALIDACOES)
   If MsgValida$ <> "" Then MsgBox MsgValida$, vbCritical + vbOKOnly, vaTitulo: Exit Sub
   ExecutaPreValidacao
   
   With grdContas1
      .OpenRecordSet x$, CURSOR_TABLE
   End With
   
   'AtualizaPai
   Screen.MousePointer = vbDefault
   
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub   
  

Private Function Filtro() As String
 Filtro = " [Seqüência do Orçamento] = " & Orcamento
 Filtro = Filtro & " And Tp <> 2 And Produtos.[Seqüência do Produto] > 0" 
End Function

Private Function Filtro5() As String
 Filtro5 = " [Seqüência do Orçamento] = " & Orcamento
 Filtro5 = Filtro5 & " And Tp = 2" 
End Function

   
Private Sub Grid_KeyPress(Index As Integer, ByVal KeyAscii As Integer, ByVal Shift As Integer, Columns() As Variant)
   'ReceitaP KeyAscii
End Sub
   
   
Private Sub BuscaNota()
 Dim Tb As New GRecordSet
 Dim NFe As New GRecordSet
 
  Set Tb = VgDb.OpenRecordSet("SELECT [Seqüência Do Pedido] From Orçamento Where [Seqüência Do Orçamento] = " & Orcamento)
  If Tb![Seqüência Do Pedido] > 0 Then
     Set NFe = VgDb.OpenRecordSet("SELECT [Seqüência Da Nota Fiscal] Seq From [Nota Fiscal] Where [Seqüência Do Pedido] = " & Tb![Seqüência Do Pedido])
     txtNFe.Value = NFe!Seq
 ' Else
     'MsgBox("Pedido Invalido!")
  End If   
 
End Sub


Private Function Filtro1() As String  
 Filtro1 = "Orçamento.[Seqüência Do Orçamento] = " & Orcamento          
End Function

Private Function Filtro2() As String
 Filtro2 = "Orçamento.[Seqüência Do Orçamento] = " & Orcamento         
End Function

Private Function Filtro3() As String
 Filtro3 = "Orçamento.[Seqüência Do Orçamento] = " & Orcamento  
End Function


Private Sub MYSQL()
 Dim SQL As String
 Dim Tb1 As New GRecordSet
 Dim SeqItem As Integer 
 Dim cnGas As ADODB.Connection, rsGas As ADODB.RecordSet
 
 Set Tb1 = VgDb.OpenRecordSet("SELECT * From [Itens pendentes] Where [Seqüência Do Orçamento] = " & Orcamento)
 
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
                                                                                                                                                                                                          
 Set Tb1 = VgDb.OpenRecordSet("SELECT * From [Itens pendentes]")
 
 SeqItem = 1
 SQL = "SELECT Tipo, [Seqüência Do Produto], Descrição, Quantidade, Total, [Valor Unitário] " & _
       "FROM( "
       
 SQL = SQL & "SELECT 1 As Tipo, [Peças do Orçamento].[Seqüência do Produto], Produtos.Descrição, Sum([Peças do Orçamento].Quantidade) Quantidade, " & _
             "Sum([Peças do Orçamento].[Valor Total]) As Total, [Peças do Orçamento].[Valor Unitário] " & _
             "FROM [Peças do Orçamento] left join Orçamento On [Peças do Orçamento].[Seqüência do Orçamento] = Orçamento.[Seqüência do Orçamento] " & _
             "Left Join Produtos On [Peças do Orçamento].[Seqüência Do Produto] = Produtos.[Seqüência Do Produto] WHERE " & Filtro1()            
 SQL = SQL & " Group By [Peças do Orçamento].[Seqüência Do Produto], Produtos.Descrição, [Peças do Orçamento].[Valor Unitário] Union "
 SQL = SQL & "Select 2 As Tipo, [Conjuntos do Orçamento].[Seqüência Do Conjunto], Conjuntos.Descrição, Sum([Conjuntos do Orçamento].Quantidade) Quantidade, Sum([Conjuntos do Orçamento].[Valor Total]) Total, [Conjuntos do Orçamento].[Valor Unitário] " & _
             "FROM [Conjuntos do Orçamento] Left Join Orçamento On [Conjuntos do Orçamento].[Seqüência do Orçamento] = Orçamento.[Seqüência do Orçamento] " & _
             "Left Join Conjuntos On [Conjuntos do Orçamento].[Seqüência Do Conjunto] = Conjuntos.[Seqüência Do Conjunto] WHERE " & Filtro2()          
 SQL = SQL & " Group By [Conjuntos do Orçamento].[Seqüência Do Conjunto], Conjuntos.Descrição, [Conjuntos do Orçamento].[Valor Unitário] Union " 
 SQL = SQL & "Select 3 As Tipo, [Produtos do Orçamento].[Seqüência Do Produto], Produtos.Descrição, Sum([Produtos do Orçamento].Quantidade) Quantidade, Sum([Produtos do Orçamento].[Valor Total] + [Produtos do Orçamento].[Valor Do IPI]) Total, [Produtos do Orçamento].[Valor Unitário] " & _
             "FROM [Produtos do Orçamento] left Join Orçamento On [Produtos do Orçamento].[Seqüência do Orçamento] = Orçamento.[Seqüência do Orçamento] " & _
             "Left Join Produtos On [Produtos do Orçamento].[Seqüência Do Produto] = Produtos.[Seqüência Do Produto] WHERE " & Filtro3()         
 SQL = SQL & " Group By [Produtos do Orçamento].[Seqüência Do Produto], Produtos.Descrição, [Produtos do Orçamento].[Valor Unitário])A Order By Tipo Desc"
 
 rsGas.Open SQL, cnGas, adOpenStatic, adLockPessimistic
 
 
 Do While Not rsGas.EOF
    With Tb1
     .AddNew
       ![Seqüência Do Orçamento] = Orcamento
       ![Sequencia Do Item] = SeqItem
        If rsGas!Tipo = 1 Or rsGas!Tipo = 3 Then
           ![Seqüência Do Produto] = rsGas![Seqüência Do Produto]
        End If   
        If rsGas!Tipo = 2 Then
           ![Seqüência Do Conjunto] = rsGas![Seqüência Do Produto]
        End If   
       !Quantidade = rsGas!Quantidade
       ![Valor Total] = rsGas!Total
       ![Valor Unitário] = rsGas![Valor Unitário]
       !Tp = rsGas!Tipo
     .Update
     .BookMark = .LastModified
    End With
 rsGas.MoveNext
 SeqItem = SeqItem + 1
 Loop
 rsGas.Close 
               
End Sub


'Private Sub ReceitaP()
' Dim Tb As New GRecordSet, Seq As Long
' Dim MP As New GRecordSet, QtdeOrc As Double
' Dim ProdAux As New GRecordSet, SeqItem As New GRecordSet
' Dim ReceitaAux As New GRecordSet, LocalAux As New GRecordSet
' Dim ProdAux2 As New GRecordSet, TemReceita As New GRecordSet
' Dim Lances As New GRecordSet, i As Long, NaoIncluir As New GRecordSet
' Dim ReceitaAuxSec As New GRecordSet, Tb1 As New GRecordSet, TB2 As New GRecordSet
' Dim ConjAux As New GRecordSet, MPrimaSec As New GRecordSet, PodeIncluir As New GRecordSet 
' Dim TerceirosAux As New GRecordSet, TbAux As New GRecordSet, TbAux2 As New GRecordSet  
' Dim SemiPronto As New GRecordSet, TbGravar As New GRecordSet 
' Dim HSemi As New GRecordSet   
       
 'Screen.MousePointer = vbHourglass     
              
  'Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] Left Join Produtos On [Receita Primaria].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Orçamento] = " & Orcamento)
  ' If Tb.RecordCount > 0 Then
  '    'Editar os Registros
  '      Do While Not Tb.EOF
  '         With Tb
  '          .Edit
  '            ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
  '            !Pagto = MostraFinanceiro(![Id Do Pedido])
  '            ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '            ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '            ![Qtde Total] = QtdeTotal(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '             Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
  '            !Localização = LocalAux!Localização
  '          .Update
  '          .BookMark = .LastModified
  '         End With
  '      Tb.MoveNext
  '      Loop
   'Atualiza Os de Terceiros tb pai caso nao tiver
  '  Set TbAux = VgDb.OpenRecordSet("Receita Primaria")
  '  Set TerceirosAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Material Adquirido de Terceiro] = 1 And [Seqüência do Orçamento] = " & Orcamento)
    
  '      Do While Not TerceirosAux.EOF
  '         Set TbAux2 = VgDb.OpenRecordSet("SELECT * From [Itens Pendentes] WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto] & " And [Seqüência do Orçamento] = " & Orcamento)
  '         If TbAux2.RecordCount = 0 Then
  '          With TbAux
  '          .AddNew
  '           ![Seqüência Do Produto] = TerceirosAux![Seqüência Do Produto]
  '           ![Seqüência da Matéria Prima] = TerceirosAux![Seqüência Do Produto]
  '           ![Seqüência Do Orçamento] = Orcamento
  '           !Quantidade = TerceirosAux!Quantidade
  '            Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
  '           ![Sequencia Do Item] = SeqItem!Sequencia + 1
  '           ![Id Do Pedido] = MostraCompra(TerceirosAux![Seqüência Do Produto])
  '           !Pagto = MostraFinanceiro(![Id Do Pedido])
  '           ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
  '           ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
  '            Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto])
  '           !Localização = LocalAux!Localização
  '          .Update
  '          .BookMark = .LastModified 
  '          End With
  '         End If
  '      TerceirosAux.MoveNext
  '      Loop       
        
  '        Set TbGravar = VgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento)
        
  '        Do While Not TbGravar.EOF
  '           If Not Vazio (TbGravar!Situação) Then
  '              If TbGravar!Situação = "Entregue ao Cliente" Or TbGravar!Situação = "Separado p/ Entregar" Then
  '                 VgDb.Execute "Update [Receita Primaria] Set Situação = '" & TbGravar!Situação & "' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & TbGravar![Seqüência Do Produto]
  '              End If    
  '           End If
  '        TbGravar.MoveNext
  '        Loop    
        
  '        AtualizaPai
        
  '    frmRecprima.Show
  '    frmRecprima.vgFiltroInicial = "[Seqüência do Orçamento] = " & Orcamento '& " And [Sequencia do Item] = 1"
  '    InicializaFiltro frmRecprima
  '    frmRecprima.TxtDtaAtual.Value = Date
  '    Exit Sub
      
  ' Else
  '    Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
  '    Set ProdAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where Lance = 0 And [Seqüência do Orçamento] = " & Orcamento)
  '       Do While Not ProdAux.EOF
  '          Set MP = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & ProdAux![Seqüência Do Produto])
  '            If MP.RecordCount > 0 Then
  '               Do While Not MP.EOF
  '               Set PodeIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
  '                If Not PodeIncluir!Y Then
  '                  With Tb
  '                     .AddNew
  '                        ![Seqüência Do Produto] = MP![Seqüência Do Produto]
  '                        ![Seqüência da Matéria Prima] = MP![Seqüência da Matéria Prima]
  '                        ![Seqüência Do Orçamento] = Orcamento
  '                        !Quantidade = MP![Quantidade de Matéria Prima] * ProdAux!Quantidade
  '                         Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
  '                        ![Sequencia Do Item] = SeqItem!Sequencia + 1
  '                        ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
  '                        !Pagto = MostraFinanceiro(![Id Do Pedido])
  '                        ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '                        ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '                         Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
  '                        !Localização = LocalAux!Localização
  '                     .Update
  '                     .BookMark = .LastModified
  '                  End With
  '                End If  
  '                 'Secundaria
  '                Set ReceitaAuxSec = vgDb.OpenRecordSet("SELECT [Mostrar Receita Secundaria] X From Produtos Where [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
  '                    If ReceitaAuxSec!x Then
  '                       Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
  '                        Do While Not Tb1.EOF
  '                         Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
  '                         If Not TB2!Y Then
  '                         With Tb
  '                           .AddNew
  '                             ![Seqüência Do Produto] = ProdAux![Seqüência Do Produto]
  '                             ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
  '                             ![Seqüência Do Orçamento] = Orcamento
  '                             !Quantidade = Tb1![Quantidade de Matéria Prima] * ProdAux!Quantidade
  '                              Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
  '                             ![Sequencia Do Item] = SeqItem!Sequencia + 1
  '                             ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
  '                             !Pagto = MostraFinanceiro(![Id Do Pedido])
  '                             ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '                             ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
  '                              Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
  '                             !Localização = LocalAux!Localização
  '                          .Update
  '                          .BookMark = .LastModified
  '                         End With
  '                         End If
  '                         Tb1.MoveNext
  '                        Loop
  '                    End If
  '                  
  '                  'fim Secundaria
  '               MP.MoveNext
  '               Loop
  '            End If
  '      ProdAux.MoveNext
  '      Loop
  '    End If
      'iNTERMEDIARIOS
 '     Set Lances = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] LEFT JOIN Produtos ON PRODUTOS.[Seqüência do Produto] = [Itens Pendentes].[Seqüência do Produto] WHERE Lance = 1 And [Seqüência do Orçamento] = " & Orcamento)
 '              Do While Not Lances.EOF
 '                  For i = 1 To Lances!Quantidade
 '                    Set TemReceita = vgDb.OpenRecordSet("SELECT * From [Matéria Prima] Where [Seqüência Do Produto] = " & Lances![Seqüência Do Produto])
 '                    Do While Not TemReceita.EOF
 '                    Set PodeIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
 '                      If Not PodeIncluir!Y Then
 '                       With Tb
 '                        .AddNew
 '                         ![Seqüência Do Produto] = Lances![Seqüência Do Produto]
 '                         ![Seqüência da Matéria Prima] = TemReceita![Seqüência da Matéria Prima]
 '                         ![Seqüência Do Orçamento] = Orcamento
 '                         !Quantidade = TemReceita![Quantidade de Matéria Prima] '* ReceitaAux!Quantidade
 '                         Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
 '                         ![Sequencia Do Item] = SeqItem!Sequencia + 1
 '                         ![Id Do Pedido] = MostraCompra(TemReceita![Seqüência da Matéria Prima])
 '                         !Pagto = MostraFinanceiro(![Id Do Pedido])
 '                         ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
 '                         ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
 '                          Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
 '                         !Localização = LocalAux!Localização
 '                       .Update
 '                       .BookMark = .LastModified
 '                       End With
 '                      End If
 '                    'Secundaria
 '                     Set ReceitaAuxSec = vgDb.OpenRecordSet("SELECT [Mostrar Receita Secundaria] X From Produtos Where [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
 '                      If ReceitaAuxSec!x Then
 '                         Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
 '                         Do While Not Tb1.EOF
 '                         Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
 '                          If Not TB2!y Then
 '                           With Tb
 '                            .AddNew
 '                              ![Seqüência Do Produto] = Lances![Seqüência Do Produto]
 '                              ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
 '                              ![Seqüência Do Orçamento] = Orcamento
 '                              !Quantidade = Tb1![Quantidade de Matéria Prima] '* Lances!Quantidade
 '                               Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
 '                              ![Sequencia Do Item] = SeqItem!Sequencia + 1
 '                              ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
 '                              !Pagto = MostraFinanceiro(![Id Do Pedido])
 '                              ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
 '                              ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
 '                               Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
 '                              !Localização = LocalAux!Localização
 '                           .Update
 '                           .BookMark = .LastModified
 '                          End With
 '                          End If
 '                         Tb1.MoveNext
 '                         Loop
 '                      End If
 '                    'Fim Sucundaria
 '                   TemReceita.MoveNext
 '                   Loop
 '                   Next
 '              Lances.MoveNext
 '              Loop       
 '      'AdqTerceiros TB Pai
 '       Set Tb = Nothing
 '       Set Tb = VgDb.OpenRecordSet("Receita Primaria")
 '       Set TerceirosAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Material Adquirido de Terceiro] = 1 And [Seqüência do Orçamento] = " & Orcamento)
 '       Do While Not TerceirosAux.EOF
 '          With Tb
 '          .AddNew
 '            ![Seqüência Do Produto] = TerceirosAux![Seqüência Do Produto]
 '            ![Seqüência da Matéria Prima] = TerceirosAux![Seqüência Do Produto]
 '            ![Seqüência Do Orçamento] = Orcamento
 '            !Quantidade = TerceirosAux!Quantidade
 '             Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
 '            ![Sequencia Do Item] = SeqItem!Sequencia + 1
 '            ![Id Do Pedido] = MostraCompra(TerceirosAux![Seqüência Do Produto])
 '            !Pagto = MostraFinanceiro(![Id Do Pedido])
 '            ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
 '            ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
 '             Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto])
 '            !Localização = LocalAux!Localização
 '          .Update
 '          .BookMark = .LastModified 
 '          End With
 '       TerceirosAux.MoveNext
  '      Loop       
                      
      'Conjuntos
 '      Set Tb = Nothing
 '      Set MP = Nothing
       
 '      Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] Where TP = 2 And [Seqüência do Orçamento] = " & Orcamento)' 2 = conjunto        
 '      Set ConjAux = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
       
 '       Do While Not Tb.EOF
 '          Set MP = VgDb.OpenRecordSet("SELECT * FROM [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] WHERE [Material Adquirido de Terceiro] = 1 And [Seqüência do Conjunto] = " & Tb![Seqüência Do Conjunto])  'so ad terc   
 '          If MP.RecordCount > 0 Then
 '                Do While Not MP.EOF
 '                Set NaoIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y From Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
 '                 If Not NaoIncluir!Y Then 
 '                   With ConjAux 'Validar Nao sair no ckeck list
 '                      .AddNew
 '                         '![Seqüência Do Produto] = 13242
 '                         ![Seqüência Do Conjunto] = MP![Seqüência Do Conjunto]
 '                         Set HSemi = vgDb.OpenRecordSet("SELECT Produtos.[Seqüência do Produto], Descrição From [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Conjunto] = " & MP![Seqüência Do Conjunto] & " and Descrição Like '%Hidroturbo Semi Pronto%'")
 '                         ![Seqüência Do Produto] = HSemi![Seqüência Do Produto]
                          
 '                         ![Seqüência da Matéria Prima] = MP![Seqüência Do Produto]  
 '                         ![Seqüência Do Orçamento] = Orcamento
 '                         !Quantidade = MP![Quantidade Do Produto] * Tb!Quantidade
 '                          Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
 '                         ![Sequencia Do Item] = SeqItem!Sequencia + 1
 '                         ![Id Do Pedido] = MostraCompra(MP![Seqüência Do Produto])
 '                         !Pagto = MostraFinanceiro(![Id Do Pedido])
 '                         ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], MP![Seqüência Do Produto])
 '                         ![Qtde Restante] = QtdeRestante(![Id Do Pedido], MP![Seqüência Do Produto])
 '                          Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
 '                         !Localização = LocalAux!Localização
 '                      .Update
 '                      .BookMark = .LastModified
 '                   End With
 '                 End If  
 '                MP.MoveNext
 '                Loop
 '             End If 
 '       Tb.MoveNext
 '       Loop
       'Secundaria Conjuntos
        
 '      Set Tb = Nothing
 '      Set MP = Nothing
 '      Set ConjAux = Nothing
 '      Set TB1 = Nothing
 '      Set TB2 = Nothing
 '      Set ReceitaAuxSec = Nothing
       
 '      Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] Where TP = 2 And [Seqüência do Orçamento] = " & Orcamento)        
 '      Set ConjAux = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
       
 '       Do While Not Tb.EOF
 '          Set MP = VgDb.OpenRecordSet("SELECT * FROM [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] WHERE [Mostrar Receita Secundaria] = 1 And [Seqüência do Conjunto] = " & Tb![Seqüência Do Conjunto])   
 '          If MP.RecordCount > 0 Then
 '           Do While Not MP.EOF
 '            Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
 '             Do While Not Tb1.EOF
 '             Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima]) 
 '             If Not TB2!Y Then
 '              With ConjAux
 '              .AddNew
 '                ![Seqüência Do Produto] = 13242
 '                ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
 '                ![Sequencia produto principal] = Tb1![Seqüência Do Produto]
 '                ![Seqüência Do Conjunto] = MP![Seqüência Do Conjunto]
 '                ![Seqüência Do Orçamento] = Orcamento
 '                !Quantidade = Tb1![Quantidade de Matéria Prima] * Tb!Quantidade
                 
                   'Set SemiPronto = vgDb.OpenRecordSet("SELECT  Descrição From Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência Do Produto])
                   'If Mid(SemiPronto!Descrição, 1, 22) <> "Hidroturbo Semi Pronto" Then
                   '  !Quantidade = Tb1![Quantidade de Matéria Prima] * Tb!Quantidade
                   'Else
                   '  !Quantidade = Tb1![Quantidade de Matéria Prima]
                   'End If
                 
  '               Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
  '               ![Sequencia Do Item] = SeqItem!Sequencia + 1
  '               ![Id Do Pedido] = MostraCompra(Tb1![Seqüência da Matéria Prima])
  '               !Pagto = MostraFinanceiro(![Id Do Pedido])
  '               ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb1![Seqüência da Matéria Prima])
  '               ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb1![Seqüência da Matéria Prima])
  '                Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb1![Seqüência Do Produto])
  '               !Localização = LocalAux!Localização
  '             .Update
  '             .BookMark = .LastModified
  '             End With
  '            End If
  '            'Dentro da Outra Receita
  '            Set PodeIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y, [Mostrar Receita Secundaria] X FROM Produtos WHERE [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
  '             If PodeIncluir!X Then 
  '                If Not PodeIncluir!Y Then
  '                Set MPrimaSec = VgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
  '                  If MPrimaSec.RecordCount > 0 Then
  '                     Do While Not MPrimaSec.EOF
  '                     Set NaoIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y From Produtos WHERE [Seqüência do Produto] = " & MPrimaSec![Seqüência da Matéria Prima])
  '                     If Not NaoIncluir!Y Then
  '                     With ConjAux
  '                      .AddNew
  '                        ![Seqüência Do Produto] = 13242
  '                        ![Seqüência da Matéria Prima] = MPrimaSec![Seqüência da Matéria Prima]
  '                       ' ![Sequencia produto principal] = MPrimaSec![Seqüência Do Produto]
  '                        ![Sequencia produto principal] = Tb1![Seqüência da Matéria Prima] 'MPrimaSec![Seqüência Do Produto]
  '                        ![Seqüência Do Conjunto] = MP![Seqüência do Conjunto]
  '                        ![Seqüência Do Orçamento] = Orcamento
  '                        !Quantidade = MPrimaSec![Quantidade de Matéria Prima] * Tb1![Quantidade de Matéria Prima]'Tb!Quantidade
  '                         Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
  '                        ![Sequencia Do Item] = SeqItem!Sequencia + 1
  '                        ![Id Do Pedido] = MostraCompra(MPrimaSec![Seqüência da Matéria Prima])
  '                        !Pagto = MostraFinanceiro(![Id Do Pedido])
  '                        ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], MPrimaSec![Seqüência da Matéria Prima])
  '                        ![Qtde Restante] = QtdeRestante(![Id Do Pedido], MPrimaSec![Seqüência da Matéria Prima])
  '                         Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & MPrimaSec![Seqüência da Matéria Prima])
  '                        !Localização = LocalAux!Localização
  '                      .Update
  '                      .BookMark = .LastModified
  '                     End With
  '                     End If
  '                     MPrimaSec.MoveNext
  '                     Loop
  '                  End If
  '                 End If 
  '             End If
  '            Tb1.MoveNext
  '            Loop
  '          MP.MoveNext
  '          Loop
  '          End If 
  '      Tb.MoveNext
  '      Loop
             
  '        Set TbGravar = VgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento)
  '      
  '        Do While Not TbGravar.EOF
  '           If Not Vazio (TbGravar!Situação) Then
  '              If TbGravar!Situação = "Entregue ao Cliente" Or TbGravar!Situação = "Separado p/ Entregar" Then
  '                 VgDb.Execute "Update [Receita Primaria] Set Situação = '" & TbGravar!Situação & "' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & TbGravar![Seqüência Do Produto]   
  '              End If   
  '           End If
  '        TbGravar.MoveNext
  '        Loop       
  '           
  '     AtualizaPai
  '                            
  '     Screen.MousePointer = vbDefault         
  '     frmRecprima.Show
  '     frmRecprima.vgFiltroInicial = "[Seqüência do Orçamento] = " & Orcamento '& " And [Sequencia do Item] = 1"
  '     InicializaFiltro frmRecprima
  '     frmRecprima.TxtDtaAtual.Value = Date
        
'End Sub


'Private Sub AtualizaSituacao()
' If Prodstatus = 0 Then Exit Sub
' VgDb.Execute "UPDATE [Itens Pendentes] Set Situação = '" & Statustela & "'" & " Where [Seqüência do Produto] > 0 And [Seqüência do Produto] = " & Prodstatus & " And [Seqüência do Orçamento] = " & Orcamento
'  If Statustela = "Entregue ao Cliente" Or Statustela = "Separado p/ Entregar" Then
'     VgDb.Execute "UPDATE [Receita Primaria] Set Situação = '" & Statustela & "'" & " Where [Seqüência do Produto] > 0 And [Seqüência do Produto] = " & Prodstatus & " And [Seqüência do Orçamento] = " & Orcamento 
'  End If
' grdContas.Rebind
' Txtsituacao.Text = ""
'End Sub


Private Function MostraItem() As String
  Dim Tb1 As New GRecordSet
  
  'Set Tb = VgDb.OpenRecordSet("SELECT [Seqüência do Produto] Cod FROM [Itens Pendentes] WHERE [Sequencia do Item] = " & Prodstatus & " And [Seqüência do Orçamento] = " & Orcamento)
  Set Tb1 = VgDb.OpenRecordSet("SELECT Descrição Nome FROM Produtos WHERE [Seqüência do Produto] = " & Prodstatus)
  MostraItem = Tb1!Nome
  
End Function


Private Sub AtualizaPai()
 Dim Tb As New GRecordSet
 Dim Tb1 As New GRecordSet
 Dim PodeAtualizar As New GRecordSet
 Dim TbVal As New GRecordSet
 
 Screen.MousePointer = vbHourglass 'Ampulheta
 Set TbVal = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento) 
 Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
 
 If TbVal.RecordCount = 0 Then : Exit Sub
 
' Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
 
 'Entregues ao Cliente
  'Do While Not Tb.EOF
  '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
  '       Do While Not Tb1.EOF
  '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Entregue ao Cliente'") 
  '            If PodeAtualizar.RecordCount = Tb1.RecordCount Then
  '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Entregue ao Cliente' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
  '            End If
  '       Tb1.MoveNext
  '       Loop  
  'Tb.MoveNext
  'Loop

  
  'Na Galvanização
  'Do While Not Tb.EOF
  '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
  '       Do While Not Tb1.EOF
  '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Na Galvanização'") 
  '            If PodeAtualizar.RecordCount > 0 Then
  '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Na Galvanização' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
  '            End If
  '       Tb1.MoveNext
  '       Loop  
  'Tb.MoveNext
  'Loop
  
  'Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
  
  'Á Produzir
  Do While Not Tb.EOF
     Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
         Do While Not Tb1.EOF
            Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Á Produzir'") 
              If PodeAtualizar.RecordCount > 0 Then
                 VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Á Produzir' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
              End If
         Tb1.MoveNext
         Loop  
  Tb.MoveNext
  Loop
  
  
  'Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
  
  'Falta MPrima
  'Do While Not Tb.EOF
  '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
  '       Do While Not Tb1.EOF
  '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Falta MPrima'") 
  '            If PodeAtualizar.RecordCount > 0 Then
  '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Falta MPrima' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
  '            End If
  '       Tb1.MoveNext
  '       Loop  
  'Tb.MoveNext
  'Loop
    
 'Á Galnanizar
   
  'Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
  
  'Do While Not Tb.EOF
  '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
  '       Do While Not Tb1.EOF
  '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Á Galnanizar'") 
  '            If PodeAtualizar.RecordCount > 0 Then
  '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Á Galnanizar' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
  '            End If
  '       Tb1.MoveNext
  '       Loop  
  'Tb.MoveNext
  'Loop
  
  'Separado p/ Montar
  'Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
 
  'Do While Not Tb.EOF
  '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
  '       Do While Not Tb1.EOF
  '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Separado p/ Montar'") 
  '            If PodeAtualizar.RecordCount = Tb1.RecordCount Then
  '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Separado p/ Montar' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
  '            End If
  '       Tb1.MoveNext
  '       Loop  
  'Tb.MoveNext
  'Loop
  
  'Separado p/ Entregar
 
   'Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And Situação <> '' Order By [Seqüência do Produto]")
 
   'Do While Not Tb.EOF
   '   Set Tb1 = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])  
   '       Do While Not Tb1.EOF
   '          Set PodeAtualizar = VgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto] & " And Situação = 'Separado p/ Entregar'") 
   '            If PodeAtualizar.RecordCount = Tb1.RecordCount Then
   '               VgDb.Execute("UPDATE [Itens Pendentes] Set Situação = 'Separado p/ Entregar' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Tb![Seqüência Do Produto])
   '            End If
   '       Tb1.MoveNext
   '      Loop  
  'Tb.MoveNext
  'Loop
  
  grdContas.Rebind
  Screen.MousePointer = vbDefault 'Ampulheta
  
End Sub


Private Function MostraCompra(Item As Long) As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT [Pedido de Compra Novo].[Id do Pedido], [Id do Produto] From [Pedido de Compra Novo] " & _
                             "Left Join [Produtos Do Pedido Compra] On [Pedido de Compra Novo].[Id Do Pedido] = [Produtos Do Pedido Compra].[Id Do Pedido] " & _
                             "Left Join [Orçamentos da compra] On [Pedido de Compra Novo].[Id Do Pedido] = [Orçamentos da compra].[Id Do Pedido] " & _
                             "Where Cancelado = 0 And Prepedido = 0 And [Orçamentos da compra].[Seqüência do Orçamento] = " & Orcamento & " And [Id do Produto] = " & Item) 
 If Tb.RecordCount > 0 Then
    If Tb![Id Do Pedido] > 0 Then                             
       MostraCompra = Tb![Id Do Pedido]
    Else
       MostraCompra = 0        
    End If 
 Else
 MostraCompra = 0 
 End If 

End Function


Private Function QtdeRecebida(Pedido As Long, Produto As Long) As Double
 Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("SELECT [Qtde Recebida] From [Bx Produtos Pedido Compra] Where [Id do Pedido] = " &  Pedido & " And [Id do Produto] = " & Produto)
  
   If Tb.RecordCount > 0 Then
      If Tb![Qtde Recebida] > 0 Then
         QtdeRecebida = Tb![Qtde Recebida]
      Else
         QtdeRecebida = 0
      End If 
   End If    
 
End Function

Private Function QtdeRestante(Pedido As Long, Produto As Long) As Double
 Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("SELECT [Qtde Restante] From [Bx Produtos Pedido Compra] Where [Id do Pedido] = " &  Pedido & " And [Id do Produto] = " & Produto)
  
   If Tb.RecordCount > 0 Then
      If Tb![Qtde Restante] > 0 Then
         QtdeRestante = Tb![Qtde Restante]
      Else
         QtdeRestante = 0
      End If
   End If     
 
End Function


Private Function QtdeTotal(Pedido As Long, Produto As Long) As Double
 Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("SELECT Qtde From [Produtos do Pedido Compra] Where [Id do Pedido] = " &  Pedido & " And [Id do Produto] = " & Produto)
  
   If Tb.RecordCount > 0 Then
      If Tb!Qtde > 0 Then
         QtdeTotal = Tb!Qtde
      Else
         QtdeTotal = 0
      End If
   End If     
 
End Function


Private Function MostraFinanceiro(Pedido As Long) As String
 Dim Tb1 As New GRecordSet
 
 Set Tb1 = vgDb.OpenRecordSet("SELECT SUM([Valor Pago]) Pago, Sum([Valor Restante]) Restante From [Manutenção Contas] WHERE [Sequencia da Compra] = " & Pedido & " And [Sequencia da Compra] > 0")
 
  'If Tb1.RecordCount = 0 Or IsNull(Tb1!Pago) Or IsNull(Tb1!Restante) Then MostraFinanceiro = "" :Exit Function
  
  If Tb1!Pago > 0 And Tb1!Restante > 0 Then
     MostraFinanceiro = "Pago Parcial"
  ElseIf Tb1!Pago > 0 And Tb1!Restante = 0 Then  
     MostraFinanceiro = "Liquidado"
  ElseIf Tb1!Pago = 0 And Tb1!Restante > 0 Then      
    MostraFinanceiro = "A Pagar"
  Else MostraFinanceiro = ""   
  End If
    
End Function


Private Sub AbreMP()

 frmRelfalmp.Show
 frmRelfalmp.SetFocus
 frmRelfalmp.txtOrcamento.Value = Orcamento 
 frmRelfalmp.BtnOk.SetFocus
 Sendk(13)
 
End Sub


Private Sub AtuPedido()
Dim Tb As New GRecordSet
Dim LocalAux As New GRecordSet

 Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] Left Join Produtos On [Receita Primaria].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Orçamento] = " & Orcamento)
   If Tb.RecordCount = 0 Then: Exit Sub
      'Editar os Registros
      Do While Not Tb.EOF
         With Tb
          .Edit
            ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
            !Pagto = MostraFinanceiro(![Id Do Pedido])
            ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
            ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
            ![Qtde Total] = QtdeTotal(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
             Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
            !Localização = LocalAux!Localização
          .Update
          .BookMark = .LastModified
        End With
      Tb.MoveNext
      Loop
End Sub

Private Sub RecPrima(Item As Long)
 Dim TemReceita As New GRecordSet
 Dim PodeIncluir As New GRecordSet, LocalAux As New GRecordSet
 Dim Tb As New GRecordSet, SeqItem As New GRecordSet 
 Dim i As Integer, Lances As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
 'Set Lances = VgDb.OpenRecordSet("SELECT QUANTIDADE FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & Item)

 VgDb.Execute "DELETE FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência Do Produto] = " & Item 

   'For i = 1 To Lances!Quantidade
   Set TemReceita = vgDb.OpenRecordSet("SELECT * From [Matéria Prima] Where [Seqüência Do Produto] = " & Item)
     Do While Not TemReceita.EOF
        Set PodeIncluir = VgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
          If Not PodeIncluir!Y Then
             With Tb
               .AddNew
                  ![Seqüência Do Produto] = Item
                  ![Seqüência da Matéria Prima] = TemReceita![Seqüência da Matéria Prima]
                  ![Seqüência Do Orçamento] = Orcamento
                  !Quantidade = TemReceita![Quantidade de Matéria Prima] '* ReceitaAux!Quantidade
                   Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                  ![Sequencia Do Item] = SeqItem!Sequencia + 1
                  ![Id Do Pedido] = MostraCompra(TemReceita![Seqüência da Matéria Prima])
                  !Pagto = MostraFinanceiro(![Id Do Pedido])
                  ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
                  ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
                   Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
                  !Localização = LocalAux!Localização
                .Update
                .BookMark = .LastModified
             End With
          End If
      TemReceita.MoveNext
      Loop 
      'Next  
      txtProdstatus.Text = ""
End Sub        

Private Sub Extornar()
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Receita primaria] WHERE [Seqüência do Orçamento] = " & Orcamento)
  
  If Tb.RecordCount > 0 Then
     VgDb.OpenRecordSet "DELETE FROM [Receita Primaria] WHERE [Seqüência do Orçamento] = " & Orcamento
  End If
  MsgBox("Concluido com Sucesso!")
 
End Sub


Private Sub Entregou()

 vgDb.Execute "UPDATE [Itens Pendentes] Set Situação = 'Entregue ao Cliente' WHERE [Seqüência do Orçamento] = " & Orcamento
 vgDb.Execute "UPDATE [Receita Primaria] Set Situação = 'Entregue ao Cliente' WHERE [Seqüência do Orçamento] = " & Orcamento
 grdContas.SetFocus
 SendK (vbKeyF5)

End Sub


Private Sub ReceitaP()
 Dim Tb As New GRecordSet, Seq As Long
 Dim MP As New GRecordSet, QtdeOrc As Double
 Dim ProdAux As New GRecordSet, SeqItem As New GRecordSet
 Dim ReceitaAux As New GRecordSet, LocalAux As New GRecordSet
 Dim ProdAux2 As New GRecordSet, TemReceita As New GRecordSet
 Dim Lances As New GRecordSet, i As Long, NaoIncluir As New GRecordSet
 Dim ReceitaAuxSec As New GRecordSet, Tb1 As New GRecordSet, TB2 As New GRecordSet
 Dim ConjAux As New GRecordSet, MPrimaSec As New GRecordSet, PodeIncluir As New GRecordSet
 Dim TerceirosAux As New GRecordSet, TbAux As New GRecordSet, TbAux2 As New GRecordSet
 Dim SemiPronto As New GRecordSet, TbGravar As New GRecordSet
 Dim HSemi As New GRecordSet, TbMat3 As New GRecordSet
       
 Screen.MousePointer = vbHourglass
              
  Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria] Left Join Produtos On [Receita Primaria].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Orçamento] = " & Orcamento)
   If Tb.RecordCount > 0 Then
      'Editar os Registros
        Do While Not Tb.EOF
           With Tb
            .Edit
              ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
              !Pagto = MostraFinanceiro(![Id Do Pedido])
              ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
              ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
              ![Qtde Total] = QtdeTotal(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
               Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
              !Localização = LocalAux!Localização
            .Update
            .BookMark = .LastModified
           End With
        Tb.MoveNext
        Loop
   'Atualiza Os de Terceiros tb pai caso nao tiver
    Set TbAux = vgDb.OpenRecordSet("Receita Primaria")
    Set TerceirosAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Material Adquirido de Terceiro] = 1 And [Seqüência do Orçamento] = " & Orcamento)
    
        Do While Not TerceirosAux.EOF
           Set TbAux2 = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto] & " And [Seqüência do Orçamento] = " & Orcamento)
           If TbAux2.RecordCount = 0 Then
            With TbAux
            .AddNew
             ![Seqüência Do Produto] = TerceirosAux![Seqüência Do Produto]
             ![Seqüência da Matéria Prima] = TerceirosAux![Seqüência Do Produto]
             ![Seqüência Do Orçamento] = Orcamento
             !Quantidade = TerceirosAux!Quantidade
              Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
             ![Sequencia Do Item] = SeqItem!Sequencia + 1
             ![Id Do Pedido] = MostraCompra(TerceirosAux![Seqüência Do Produto])
             !Pagto = MostraFinanceiro(![Id Do Pedido])
             ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
             ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
              Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto])
             !Localização = LocalAux!Localização
            .Update
            .BookMark = .LastModified
            End With
           End If
        TerceirosAux.MoveNext
        Loop
        
          Set TbGravar = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento)
        
          Do While Not TbGravar.EOF
             If Not Vazio(TbGravar!Situação) Then
                If TbGravar!Situação = "Entregue ao Cliente" Or TbGravar!Situação = "Separado p/ Entregar" Then
                   vgDb.Execute "Update [Receita Primaria] Set Situação = '" & TbGravar!Situação & "' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & TbGravar![Seqüência Do Produto]
                End If
             End If
          TbGravar.MoveNext
          Loop
          
          Set TbGravar = Nothing
          'produzir teste
         ' Set TbGravar = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento)
        
          'Do While Not TbGravar.EOF
          '   If Not Vazio(TbGravar!Situação) Then
          '      If TbGravar!Situação = "Á Produzir" Then
          '         Set TbMat3 = VgDb.OpenRecordSet("SELECT * From [Matéria Prima] WHERE [Seqüência do Produto] = " & TbGravar![Seqüência Do Produto]) 
                    'produto em processo
          '           Do While Not TbMat3.EOF
          '              Set SemiPronto = VgDb.OpenRecordSet("SELECT [Tipo do Produto] Tipo From Produtos WHERE [Seqüência do Produto] = " & TbMat3![Seqüência da Matéria Prima])
          '                If SemiPronto!Tipo = 6 then'produto em processo
          '                   vgDb.Execute "Update [Receita Primaria] Set Situação = 'Á Produzir' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência da Matéria Prima] = " & TbMat3![Seqüência da Matéria Prima]  
          '                End If
          '              TbMat3.MoveNext
          '           Loop   
          '      End If
          '   End If
          'TbGravar.MoveNext
          'Loop
          
          'Set TbMat3 = Nothing
          'Set TbGravar = Nothing
          
      'AtualizaPai
      frmRecprima.Show
      frmRecprima.vgFiltroInicial = "[Seqüência do Orçamento] = " & Orcamento '& " And [Sequencia do Item] = 1"
      InicializaFiltro frmRecprima
      frmRecprima.TxtDtaAtual.Value = Date
      Exit Sub
      
   Else
      Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
      Set ProdAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where Lance = 0 And [Seqüência do Orçamento] = " & Orcamento)
         Do While Not ProdAux.EOF
            Set MP = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & ProdAux![Seqüência Do Produto])
              If MP.RecordCount > 0 Then
                 Do While Not MP.EOF
                 Set PodeIncluir = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
                  If Not PodeIncluir!y Then
                    With Tb
                       .AddNew
                          ![Seqüência Do Produto] = MP![Seqüência Do Produto]
                          ![Seqüência da Matéria Prima] = MP![Seqüência da Matéria Prima]
                          ![Seqüência Do Orçamento] = Orcamento
                          !Quantidade = MP![Quantidade de Matéria Prima] * ProdAux!Quantidade
                           Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                          ![Sequencia Do Item] = SeqItem!Sequencia + 1
                          ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
                          !Pagto = MostraFinanceiro(![Id Do Pedido])
                          ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                          ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                           Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
                          !Localização = LocalAux!Localização
                       .Update
                       .BookMark = .LastModified
                    End With
                  End If
                   'Secundaria
                  Set ReceitaAuxSec = vgDb.OpenRecordSet("SELECT [Mostrar Receita Secundaria] X From Produtos Where [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
                      If ReceitaAuxSec!x Then
                         Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & MP![Seqüência da Matéria Prima])
                          Do While Not Tb1.EOF
                           Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
                           If Not TB2!y Then
                           With Tb
                             .AddNew
                               ![Seqüência Do Produto] = ProdAux![Seqüência Do Produto]
                               ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
                               ![Seqüência Do Orçamento] = Orcamento
                               !Quantidade = Tb1![Quantidade de Matéria Prima] * ProdAux!Quantidade
                                Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                               ![Sequencia Do Item] = SeqItem!Sequencia + 1
                               ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
                               !Pagto = MostraFinanceiro(![Id Do Pedido])
                               ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                               ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                                Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
                               !Localização = LocalAux!Localização
                            .Update
                            .BookMark = .LastModified
                           End With
                           End If
                           Tb1.MoveNext
                          Loop
                      End If
                    
                    'fim Secundaria
                 MP.MoveNext
                 Loop
              End If
        ProdAux.MoveNext
        Loop
      End If
      'iNTERMEDIARIOS
      Set Lances = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] LEFT JOIN Produtos ON PRODUTOS.[Seqüência do Produto] = [Itens Pendentes].[Seqüência do Produto] WHERE Lance = 1 And [Seqüência do Orçamento] = " & Orcamento)
               Do While Not Lances.EOF
                   For i = 1 To Lances!Quantidade
                     Set TemReceita = vgDb.OpenRecordSet("SELECT * From [Matéria Prima] Where [Seqüência Do Produto] = " & Lances![Seqüência Do Produto])
                     Do While Not TemReceita.EOF
                     Set PodeIncluir = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
                       If Not PodeIncluir!y Then
                        With Tb
                         .AddNew
                          ![Seqüência Do Produto] = Lances![Seqüência Do Produto]
                          ![Seqüência da Matéria Prima] = TemReceita![Seqüência da Matéria Prima]
                          ![Seqüência Do Orçamento] = Orcamento
                          !Quantidade = TemReceita![Quantidade de Matéria Prima] '* ReceitaAux!Quantidade
                          Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                          ![Sequencia Do Item] = SeqItem!Sequencia + 1
                          ![Id Do Pedido] = MostraCompra(TemReceita![Seqüência da Matéria Prima])
                          !Pagto = MostraFinanceiro(![Id Do Pedido])
                          ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
                          ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TemReceita![Seqüência da Matéria Prima])
                           Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
                          !Localização = LocalAux!Localização
                        .Update
                        .BookMark = .LastModified
                        End With
                       End If
                     'Secundaria
                      Set ReceitaAuxSec = vgDb.OpenRecordSet("SELECT [Mostrar Receita Secundaria] X From Produtos Where [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
                       If ReceitaAuxSec!x Then
                          Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & TemReceita![Seqüência da Matéria Prima])
                          Do While Not Tb1.EOF
                          Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
                           If Not TB2!y Then
                            With Tb
                             .AddNew
                               ![Seqüência Do Produto] = Lances![Seqüência Do Produto]
                               ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
                               ![Seqüência Do Orçamento] = Orcamento
                               !Quantidade = Tb1![Quantidade de Matéria Prima] '* Lances!Quantidade
                                Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                               ![Sequencia Do Item] = SeqItem!Sequencia + 1
                               ![Id Do Pedido] = MostraCompra(Tb![Seqüência da Matéria Prima])
                               !Pagto = MostraFinanceiro(![Id Do Pedido])
                               ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                               ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb![Seqüência da Matéria Prima])
                                Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb![Seqüência da Matéria Prima])
                               !Localização = LocalAux!Localização
                            .Update
                            .BookMark = .LastModified
                           End With
                           End If
                          Tb1.MoveNext
                          Loop
                       End If
                     'Fim Sucundaria
                    TemReceita.MoveNext
                    Loop
                    Next
               Lances.MoveNext
               Loop
       'AdqTerceiros TB Pai
        Set Tb = Nothing
        Set Tb = vgDb.OpenRecordSet("Receita Primaria")
        Set TerceirosAux = vgDb.OpenRecordSet("SELECT * From [Itens Pendentes] Left Join Produtos On [Itens Pendentes].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Material Adquirido de Terceiro] = 1 And [Seqüência do Orçamento] = " & Orcamento)
        Do While Not TerceirosAux.EOF
           With Tb
           .AddNew
             ![Seqüência Do Produto] = TerceirosAux![Seqüência Do Produto]
             ![Seqüência da Matéria Prima] = TerceirosAux![Seqüência Do Produto]
             ![Seqüência Do Orçamento] = Orcamento
             !Quantidade = TerceirosAux!Quantidade
              Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
             ![Sequencia Do Item] = SeqItem!Sequencia + 1
             ![Id Do Pedido] = MostraCompra(TerceirosAux![Seqüência Do Produto])
             !Pagto = MostraFinanceiro(![Id Do Pedido])
             ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
             ![Qtde Restante] = QtdeRestante(![Id Do Pedido], TerceirosAux![Seqüência Do Produto])
              Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & TerceirosAux![Seqüência Do Produto])
             !Localização = LocalAux!Localização
           .Update
           .BookMark = .LastModified
           End With
        TerceirosAux.MoveNext
        Loop
                      
      'Conjuntos
       Set Tb = Nothing
       Set MP = Nothing
       
       Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] Where TP = 2 And [Seqüência do Orçamento] = " & Orcamento) ' 2 = conjunto
       Set ConjAux = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
       
        Do While Not Tb.EOF
           Set MP = vgDb.OpenRecordSet("SELECT * FROM [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] WHERE [Material Adquirido de Terceiro] = 1 And [Seqüência do Conjunto] = " & Tb![Seqüência Do Conjunto])  'so ad terc
           If MP.RecordCount > 0 Then
                 Do While Not MP.EOF
                 Set NaoIncluir = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y From Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
                  If Not NaoIncluir!y Then
                    With ConjAux 'Validar Nao sair no ckeck list
                       .AddNew
                          '![Seqüência Do Produto] = 13242
                          ![Seqüência Do Conjunto] = MP![Seqüência Do Conjunto]
                           Set HSemi = vgDb.OpenRecordSet("SELECT Produtos.[Seqüência do Produto], Descrição From [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Conjunto] = " & MP![Seqüência Do Conjunto] & " and Descrição Like '%Hidroturbo Semi Pronto%'")
                          ![Seqüência Do Produto] = 0 'HSemi![Seqüência do Produto]
                          
                          ![Seqüência da Matéria Prima] = MP![Seqüência Do Produto]
                          ![Seqüência Do Orçamento] = Orcamento
                          !Quantidade = MP![Quantidade Do Produto] * Tb!Quantidade
                           Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                          ![Sequencia Do Item] = SeqItem!Sequencia + 1
                          ![Id Do Pedido] = MostraCompra(MP![Seqüência Do Produto])
                          !Pagto = MostraFinanceiro(![Id Do Pedido])
                          ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], MP![Seqüência Do Produto])
                          ![Qtde Restante] = QtdeRestante(![Id Do Pedido], MP![Seqüência Do Produto])
                           Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
                          !Localização = LocalAux!Localização
                       .Update
                       .BookMark = .LastModified
                    End With
                  End If
                 MP.MoveNext
                 Loop
              End If
        Tb.MoveNext
        Loop
       'Secundaria Conjuntos
        
       Set Tb = Nothing
       Set MP = Nothing
       Set ConjAux = Nothing
       Set Tb1 = Nothing
       Set TB2 = Nothing
       Set ReceitaAuxSec = Nothing
       
       Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] Where TP = 2 And [Seqüência do Orçamento] = " & Orcamento)
       Set ConjAux = vgDb.OpenRecordSet("SELECT * FROM [Receita Primaria]")
       
        Do While Not Tb.EOF
           Set MP = vgDb.OpenRecordSet("SELECT * FROM [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] WHERE [Mostrar Receita Secundaria] = 1 And [Seqüência do Conjunto] = " & Tb![Seqüência Do Conjunto])
           If MP.RecordCount > 0 Then
            Do While Not MP.EOF
             Set Tb1 = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & MP![Seqüência Do Produto])
              Do While Not Tb1.EOF
              Set TB2 = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y FROM Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
              If Not TB2!y Then
               With ConjAux
               .AddNew
                 '![Seqüência do Produto] = 13242
                 ![Seqüência da Matéria Prima] = Tb1![Seqüência da Matéria Prima]
                 ![Sequencia Produto Principal] = Tb1![Seqüência Do Produto]
                 ![Seqüência Do Conjunto] = MP![Seqüência Do Conjunto]
                   Set HSemi = vgDb.OpenRecordSet("SELECT Produtos.[Seqüência do Produto], Descrição From [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Conjunto] = " & MP![Seqüência Do Conjunto] & " and Descrição Like '%Hidroturbo Semi Pronto%'")
                  ![Seqüência Do Produto] = HSemi![Seqüência Do Produto]
                 ![Seqüência Do Orçamento] = Orcamento
                 !Quantidade = Tb1![Quantidade de Matéria Prima] * Tb!Quantidade
                 
                   'Set SemiPronto = vgDb.OpenRecordSet("SELECT  Descrição From Produtos Where [Seqüência do Produto] = " & Tb1![Seqüência Do Produto])
                   'If Mid(SemiPronto!Descrição, 1, 22) <> "Hidroturbo Semi Pronto" Then
                   '  !Quantidade = Tb1![Quantidade de Matéria Prima] * Tb!Quantidade
                   'Else
                   '  !Quantidade = Tb1![Quantidade de Matéria Prima]
                   'End If
                 
                 Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                 ![Sequencia Do Item] = SeqItem!Sequencia + 1
                 ![Id Do Pedido] = MostraCompra(Tb1![Seqüência da Matéria Prima])
                 !Pagto = MostraFinanceiro(![Id Do Pedido])
                 ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], Tb1![Seqüência da Matéria Prima])
                 ![Qtde Restante] = QtdeRestante(![Id Do Pedido], Tb1![Seqüência da Matéria Prima])
                  Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & Tb1![Seqüência Do Produto])
                 !Localização = LocalAux!Localização
               .Update
               .BookMark = .LastModified
               End With
              End If
              'Dentro da Outra Receita
              Set PodeIncluir = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y, [Mostrar Receita Secundaria] X FROM Produtos WHERE [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
               If PodeIncluir!x Then
                  If Not PodeIncluir!y Then
                  Set MPrimaSec = vgDb.OpenRecordSet("SELECT * FROM [Matéria Prima] WHERE [Seqüência do Produto] = " & Tb1![Seqüência da Matéria Prima])
                    If MPrimaSec.RecordCount > 0 Then
                       Do While Not MPrimaSec.EOF
                       Set NaoIncluir = vgDb.OpenRecordSet("SELECT [Nao sair no checklist] Y From Produtos WHERE [Seqüência do Produto] = " & MPrimaSec![Seqüência da Matéria Prima])
                       If Not NaoIncluir!y Then
                       With ConjAux
                        .AddNew
                          
                          ![Seqüência da Matéria Prima] = MPrimaSec![Seqüência da Matéria Prima]
                         ' ![Sequencia produto principal] = MPrimaSec![Seqüência Do Produto]
                          ![Sequencia Produto Principal] = Tb1![Seqüência da Matéria Prima] 'MPrimaSec![Seqüência Do Produto]
                          ![Seqüência Do Conjunto] = MP![Seqüência Do Conjunto]
                          
                           Set HSemi = vgDb.OpenRecordSet("SELECT Produtos.[Seqüência do Produto], Descrição From [Itens do Conjunto] Left Join Produtos On [Itens do Conjunto].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where [Seqüência do Conjunto] = " & MP![Seqüência Do Conjunto] & " and Descrição Like '%Hidroturbo Semi Pronto%'")
                          ![Seqüência Do Produto] = HSemi![Seqüência Do Produto]
                          ![Seqüência Do Orçamento] = Orcamento
                          !Quantidade = MPrimaSec![Quantidade de Matéria Prima] * Tb1![Quantidade de Matéria Prima] 'Tb!Quantidade
                           Set SeqItem = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Sequencia FROM [Receita Primaria] Where [Seqüência do Orçamento] = " & Orcamento)
                          ![Sequencia Do Item] = SeqItem!Sequencia + 1
                          ![Id Do Pedido] = MostraCompra(MPrimaSec![Seqüência da Matéria Prima])
                          !Pagto = MostraFinanceiro(![Id Do Pedido])
                          ![Qtde Recebida] = QtdeRecebida(![Id Do Pedido], MPrimaSec![Seqüência da Matéria Prima])
                          ![Qtde Restante] = QtdeRestante(![Id Do Pedido], MPrimaSec![Seqüência da Matéria Prima])
                           Set LocalAux = vgDb.OpenRecordSet("SELECT Localização From Produtos WHERE [Seqüência do Produto] = " & MPrimaSec![Seqüência da Matéria Prima])
                          !Localização = LocalAux!Localização
                        .Update
                        .BookMark = .LastModified
                       End With
                       End If
                       MPrimaSec.MoveNext
                       Loop
                    End If
                   End If
               End If
              Tb1.MoveNext
              Loop
            MP.MoveNext
            Loop
            End If
        Tb.MoveNext
        Loop
             
          Set TbGravar = vgDb.OpenRecordSet("SELECT * FROM [Itens Pendentes] WHERE [Seqüência do Orçamento] = " & Orcamento)
        
          Do While Not TbGravar.EOF
             If Not Vazio(TbGravar!Situação) Then
                If TbGravar!Situação = "Entregue ao Cliente" Or TbGravar!Situação = "Separado p/ Entregar" Then
                   vgDb.Execute "Update [Receita Primaria] Set Situação = '" & TbGravar!Situação & "' WHERE [Seqüência do Orçamento] = " & Orcamento & " And [Seqüência do Produto] = " & TbGravar![Seqüência Do Produto]
                End If
             End If
          TbGravar.MoveNext
          Loop
              
       'AtualizaPai
                              
       Screen.MousePointer = vbDefault
       frmRecprima.Show
       frmRecprima.vgFiltroInicial = "[Seqüência do Orçamento] = " & Orcamento '& " And [Sequencia do Item] = 1"
       InicializaFiltro frmRecprima
       frmRecprima.TxtDtaAtual.Value = Date
        
End Sub


Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal Row As Long, Columns() As Variant, IsValid As Boolean, Column As Integer, ErrorMessage As String)
 Dim Situacao As String
 Dim Produto As Long
 Dim PAux As New GRecordSet

 
 Produto = Columns(2)
 Situacao = Columns(5)
 
 If Index = 0 Then
    Set PAux = vgDb.OpenRecordSet("SELECT [Receita Conferida] X From Produtos WHERE [Seqüência do Produto] = " & Produto)
 Else
    Set PAux = vgDb.OpenRecordSet("SELECT [Receita Conferida] X From Conjuntos WHERE [Seqüência do Conjunto] = " & Produto)
 End If
 
   If Situacao = "" Then Exit Sub
   If Not PAux!x Then
      MsgBox ("ATENÇÃO Impossivel Prosseguir Receita Não Conferida")
        If Situacao = "Á Produzir" Then
        vgDb.BeginTrans
           vgDb.Execute "Update [Itens Pendentes] set Situação = '' WHERE [Seqüência do Produto] = " & Produto & " And [Seqüência do Orçamento] = " & Orcamento
        vgDb.CommitTrans
           frmConpedi.Txtorc.SetFocus
           grdContas.ReBind
           grdContas.SetFocus
        End If
   End If
 
End Sub

Private Sub PreValidaPedido()
   Dim Tb As New GRecordSet, Itens As New GRecordSet
   Dim Vector() As String
   Dim i As Long, vaPrivez As Boolean, Campo As Variant, Mensagem As String
   Dim Servicos As New GRecordSet
   
   On Error GoTo DeuErro
  
   Set Itens = vgDb.OpenRecordSet("SELECT * FROM(" & _
                                  "SELECT PO.[Seqüência do Orçamento], 'Produto' Tipo, P.[Receita Conferida], P.[Seqüência do Produto], P.Descrição, U.[Sigla da Unidade], PO.[Valor Unitário] " & _
                                  "FROM [Produtos do Orçamento] PO LEFT JOIN Produtos P ON PO.[Seqüência Do Produto] = P.[Seqüência Do Produto] " & _
                                  "LEFT JOIN Unidades U ON P.[Seqüência da Unidade] = U.[Seqüência da Unidade] " & _
                                  "UNION ALL " & _
                                  "SELECT CO.[Seqüência do Orçamento],'Conjuntos' Tipo, C.[Receita Conferida], C.[Seqüência Do Conjunto], C.Descrição, U.[Sigla da Unidade], CO.[Valor Unitário] " & _
                                  "FROM [Conjuntos do Orçamento] CO LEFT JOIN Conjuntos C ON CO.[Seqüência Do Conjunto] = C.[Seqüência Do Conjunto] " & _
                                  "LEFT JOIN Unidades U ON C.[Seqüência da Unidade] = U.[Seqüência da Unidade] " & _
                                  "LEFT JOIN Orçamento O ON CO.[Seqüência do Orçamento] = O.[Seqüência do Orçamento] " & _
                                  "UNION ALL " & _
                                  "SELECT PeOrc.[Seqüência do Orçamento],'Peças' Tipo, Pe.[Receita Conferida], Pe.[Seqüência Do Produto], Pe.Descrição, U.[Sigla da Unidade], PeOrc.[Valor Unitário] " & _
                                  "FROM [Peças do Orçamento] PeOrc LEFT JOIN Produtos Pe ON PeOrc.[Seqüência Do Produto] = Pe.[Seqüência Do Produto] " & _
                                  "LEFT JOIN Unidades U ON Pe.[Seqüência da Unidade] = U.[Seqüência da Unidade] " & _
                                  "LEFT JOIN Orçamento O ON PeOrc.[Seqüência do Orçamento] = O.[Seqüência do Orçamento] " & _
                                  ") A " & _
                                  "WHERE [Seqüência do Orçamento] = " & Orcamento)
                                  
   'Set Servicos = vgDb.OpenRecordSet("SELECT * FROM [Serviços do Orçamento] WHERE [Seqüência do Orçamento] = " & Sequencia_do_Orcamento)                                
                                                               
   If Itens.RecordCount = 0 Then Exit Sub
   
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
 
   Do While Not Itens.EOF

      If Itens![Receita Conferida] = 0 Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM: " & Itens![Seqüência Do Produto] & " - " & Itens!Descrição
      Itens.MoveNext
   Loop
      
   If UBound(Vector) > 0 and vgPWUsuario <> "YGOR" Then
      Mensagem = "Alguns Itens (sem conferencia da Receita):" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
   
   CarregaContas
    
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Orcamento = txtCampo(0).Value
   NroNFe = txtCampo(1).Value
   Prodstatus = txtCampo(2).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "(((([Seqüência do Orçamento] > " & 0 & ") AND [Venda Fechada] = " & 1 & ") AND Cancelado = " & 0 & ") OR [Seqüência do Geral] = " & 517 & ") OR [Seqüência do Geral] = " & 3960 & " Or [Ordem Interna] = " & 1 & ""
      Case 1
         txtCampo(1).Filter = "(Transmitido = " & 1 & ") AND Autorizado = " & 1 & ""
      Case 2
         txtCampo(2).Filter = "([Seqüência do Produto] > " & 0 & ") AND Inativo = " & 0 & ""
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
      If True Then
         vgOk = (Orcamento > 0)
         vgMsg$ = "Informe o Orçamento!"
         If Not vgOk Then txtCampo(0).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
         ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
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

'mostra fórmulas no formulário
Private Sub MostraFormulas()
   On Error Resume Next                           'previne erros...
   Label(2).Caption = Cliente()
   If Err Then Err = 0: Label(2).Caption = ""
   txtCampo(3).Value = MostraItem()
   If Err Then Err = 0: txtCampo(3).Text = ""
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean
   On Error Resume Next                           'prepara para possiveis erros
   Ok = (True)
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (Orcamento > 0)
   Botao(0).Enabled = Ok
   Ok = (Orcamento > 0)
   Botao(1).Enabled = Ok
   Ok = (Prodstatus > 0)
   Botao(3).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(2).Visible = (vgPWUsuario = "YGOR" Or vgPWUsuario = "JERONIMO")
   Label(4).Visible = (vgPWUsuario = "YGOR" Or vgPWUsuario = "JERONIMO")
   txtCampo(3).Visible = (vgPWUsuario = "YGOR" Or vgPWUsuario = "JERONIMO")
   Botao(3).Visible = (vgPWUsuario = "YGOR" Or vgPWUsuario = "JERONIMO")
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o botão for pressionado o Botão do PopupMenu
Public Sub AcaoPopUp0(Index As Integer)
   Select Case Index
      Case 1
         mdiIRRIG.MRelentcl
         AtivaForm Me
   End Select
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 0 Or Index = 1 Or Index = 2 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 0
            PoeRelEFiltroCbo 0
         Case 2
            PoeRelEFiltroCbo 2
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
   MostraFormulas                                 'mostra formulas na janela
   ExecutaVisivel                                 'torna camos visiveis
   ExecutaPreValidacao                            'habilita/desabilita campos
   Select Case Index
      Case 0
         PreValidaPedido
   End Select
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(50420)
   vgFormID = 1343
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(41925))
   Set Txtorc = txtCampo(0)
   Set grdContas = Grid(0)
   Set Lblajuste = Label(3)
   Set Txtnfe = txtCampo(1)
   Set txtProdstatus = txtCampo(2)
   Set Grdcontas1 = Grid(1)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(50425)
   Label(1).Caption = LoadGasString(50426)
   Label(3).Caption = LoadGasString(50427)
   Botao(0).Caption = LoadGasString(50428)
   Label(4).Caption = LoadGasString(50429)
   Botao(1).Caption = LoadGasString(50430)
   Botao(2).Caption = LoadGasString(50431)
   Botao(3).Caption = LoadGasString(50432)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
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

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = ""
   txtCampo(2).BoundColumn = "Seqüência do Produto"
   txtCampo(2).ListFields = "Seqüência do Produto; Descrição"
   txtCampo(2).OrderFields = "Seqüência do Produto; Descrição"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Produtos"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = ""
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "999999"
   txtCampo(0).BoundColumn = "Seqüência do Orçamento"
   txtCampo(0).ListFields = "Seqüência do Orçamento"
   txtCampo(0).OrderFields = "Seqüência do Orçamento-"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Orçamento"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "999999"
   txtCampo(1).Editable = False
   txtCampo(1).BoundColumn = "Seqüência da Nota Fiscal"
   txtCampo(1).ListFields = "Número da NFe"
   txtCampo(1).OrderFields = "Número da NFe-"
   txtCampo(1).Relation = ""
   txtCampo(1).Source = "Nota Fiscal"
   txtCampo(1).vgfrmGMCale.grdListaG.AutoRebind = True

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
   Set Txtorc = Nothing
   Set grdContas = Nothing
   Set Lblajuste = Nothing
   Set Txtnfe = Nothing
   Set txtProdstatus = Nothing
   Set Grdcontas1 = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmConpedi = Nothing                       'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then                 'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         ReceitaP  
      Case 1
         hMenu = CreatePopupMenu()
         AppendMenu hMenu, MF_STRING, 1, "Situação do Pedido"
         GetCursorPos pt
         AcaoPopup0 TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
         DestroyMenu hMenu
      Case 2
         Entregou  
      Case 3
         RecPrima((Prodstatus))
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

