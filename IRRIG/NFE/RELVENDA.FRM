VERSION 5.00
Begin VB.Form frmRelVenda
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   3690
   Left                     =   4395
   LinkTopic                =   "RelVenda"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   3255
   ScaleWidth               =   11970
   Top                      =   2980
   Width                    =   12060
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
      Height                   =   3090
      Left                     =   15
      Width                    =   11790
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   20785
      _extenty                 =   5447
      Index                    =   0 
      Begin NFE.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   0
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   330
         Left                     =   3195
         Width                    =   1485
         Top                      =   1800
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Prossegue"
      End
      Begin VB.PictureBox picBox
         Appearance               =   1
         BackColor                =   &H8000000F
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
         BorderStyle              =   0
         Height                   =   1395
         Left                     =   75
         Width                    =   1455
         Top                      =   75
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3645
         Width                    =   735
         Top                      =   2595
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
         ForeColor                =   &H800000
         Caption                  =   "Destino:"
         Index                    =   0 
      End
      Begin NFE.GPainel Painel
         BackColor                =   &H8000000F
         BevelOuter               =   0
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   240
         Left                     =   4455
         Width                    =   2280
         Top                      =   2595
         TabStop                  =  0
         _extentx                 =   4019
         _extenty                 =   423
         Index                    =   1 
         Begin             VB.Label labopcPainel1
            BorderStyle              =   1
            Height                   =   120
            Left                     =   675
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcPainel1Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   1
            TabStop                  =   0
            Height                   =   240
            Left                     =   0
            Width                    =   915
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Janela"
            Index                    =   0 
         End
         Begin VB.OptionButton opcPainel1Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   2
            TabStop                  =   0
            Height                   =   240
            Left                     =   990
            Width                    =   1320
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Impressora"
            Index                    =   1 
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   135
         Width                    =   810
         Top                      =   2565
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
         ForeColor                =   &H800000
         Caption                  =   "Intervalo:"
         Index                    =   1 
      End
      Begin NFE.GPainel Painel
         BackColor                =   &H8000000F
         BevelOuter               =   0
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   240
         Left                     =   1020
         Width                    =   2010
         Top                      =   2565
         TabStop                  =  0
         _extentx                 =   3543
         _extenty                 =   423
         Index                    =   2 
         Begin             VB.Label labopcPainel2
            BorderStyle              =   1
            Height                   =   120
            Left                     =   665
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcPainel2Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   3
            TabStop                  =   0
            Height                   =   240
            Left                     =   0
            Width                    =   905
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Todas"
            Index                    =   0 
         End
         Begin VB.OptionButton opcPainel2Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   4
            TabStop                  =   0
            Height                   =   240
            Left                     =   975
            Width                    =   1050
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Páginas"
            Index                    =   1 
         End
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
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   315
         Left                     =   3105
         Width                    =   465
         Top                      =   2565
         Index                    =   0 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7845
         Top                      =   2565
         Width                    =   570
         Height                   =   315
      End
      Begin VB.VScrollBar scrVtxtCampo1
         Max                      =   0
         Min                      =   32767
         LargeChange              =   10
         SmallChange              =   1
         Value                    =   0
         TabStop                  =   0
         Left                     =   8205
         Top                      =   2580
         Width                    =   195
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
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   6
         TabStop                  =   0
         Height                   =   285
         Left                     =   7860
         Width                    =   330
         Top                      =   2580
         Index                    =   1 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6810
         Width                    =   960
         Top                      =   2595
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
         Caption                  =   "Nº Copias:"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8490
         Width                    =   705
         Top                      =   2595
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
         Caption                  =   "Impres.:"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   300
         Left                     =   9270
         Width                    =   1920
         Top                      =   2565
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
         ForeColor                =   &HFF0000
         Caption                  =   ""
         Index                    =   4 
      End
      Begin NFE.GBotao botConfImp
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   7
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   360
         Left                     =   11220
         Width                    =   360
         Top                      =   2520
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   ""
      End
      Begin NFE.GBotao botCancela
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   8
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   330
         Left                     =   4680
         Width                    =   1485
         Top                      =   1800
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Cancela"
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   435
         Left                     =   1695
         Width                    =   1440
         Top                      =   555
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
         Caption                  =   "Imprimir"
         Index                    =   5 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   90
         Y1                       =   2385
         X2                       =   11565
         Y2                       =   2385
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   11655
         Width                    =   45
         Top                      =   3240
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
         Caption                  =   "-"
         Index                    =   6 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   9
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   330
         Left                     =   3195
         Width                    =   1485
         Top                      =   1800
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Prossegue"
         Index                    =   0 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1730
         Top                      =   1800
         Width                    =   1330
         Height                   =   300
      End
      Begin VB.Label labtxtCampo2
         Caption                  =   "Venda / O.S: "
         BackStyle                =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Left                     =   315
         Top                      =   1800
         Width                    =   1365
         Height                   =   240
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   2790
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
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   10
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1745
         Width                    =   1030
         Top                      =   1815
         TabStop                  =  0
         Index                    =   2 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1740
         Y1                       =   930
         X2                       =   3165
         Y2                       =   930
         Index                    =   1 
      End
   End
End
Attribute VB_name = "frmRelVenda"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: RelVenda
'* Função....: Imprimir
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:34
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
Public vgUltimoTabIndex As Integer                'último campo com foco do módulo 

Public vgUltimoFiltro As String                   'último filtro definido no módulo
Public vgUltimaOrdem As String                    'última ordenação feita no módulo
Public vgUltimoFiltroComTit As String             'titulo do último filtro definido no módulo
Public vgUltimaOrdemComTit As String              'titulo da última ordenação feita no módulo 
Public vgSQL As String                            'expressão SQL que define o módulo
Public vgFP As New frmPreview                     'formuário de preview do relatório 
Public vgVisivel As Boolean                       'se a tela de parâmetros do relatório esta visível 
Public vgPosLin As Single                         'posição atual de impressão 
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFG As frmGauge                           'formulário do gauge 
Public vgQReg As Long                             'quantidade de já feira para o gauge  

Dim vgBookPg() As Variant                         'marcador do registro do inicio da página 
Dim vgPodeVoltar() As Boolean                     'se pode ou não retroceder para esta página   
Dim vgDataInicial As Variant                      'data que iniciou o relatório  
Dim vgPrimeiroCab As Integer                      'flag se ja foi impresso o primeiro cabeçlho do relatório
Dim vgFiltroInicial As String                     'filtro inicial do relatório
Dim vgOrdemInicial As String                      'ordem incial do relatório 
Dim txtCampo(2) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim vgWidthRel As Single                          'lagura do relatório 
Dim vgHeightRel As Single                         'altura do relatório
Dim vgPagRef() As Integer                         'Ordem da página dentro do intervalo 
Dim opcPainel1(1) As New FormataCampos
Dim opcPainel2(1) As New FormataCampos
Dim vgDestRel As Integer, vgModoInt As Integer, vgIntPag As String
Dim vgCopias As Integer, Ajuste As String, VendaTela As Double
Dim vgMemo0() As Long, vgMemo1() As Long, vgPgvgTot1() As Double, vgMemo2() As Long
Public txtvgDestRel As Object, lblAjuste As Object, txtVenda As Object
Dim Vendas As New GRecordSet

Dim Parcelas As New GRecordSet

Private Sub MySql()
   Dim SQL As String

   SQL = "SELECT TotalVenda, ST, IPI, ISS, [Sequencia da Venda], [Sequencia Do Geral], [Razão Social], [Fone], [Endereço], [Nro], [Bairro], " & _
         "SeqItem, Descricao, Municipio, [Uf], [Rg e ie], [Cpf e cnpj], Cep, " & _
         "Qtde, Tipo, Valor, Emissao, TotalProdutos, TotalServicos, Desconto, Frete, Transporte " & _
         "FROM( " 
          SQL = SQL & "SELECT Vendas.[Valor Total da Venda] As TotalVenda, Vendas.[Valor total icms st] As ST, Vendas.[Valor total do ipi] As IPI, Vendas.[Valor total do ISS] As ISS, Vendas.[Sequencia da Venda], Vendas.[Sequencia Do Geral], Geral.[Razão Social], " & _
                      "Geral.[Fone], Geral.[Endereço], Geral.Nro, Geral.Bairro, [Serviços da Venda].[Sequencia do Serviço] As SeqItem, " & _
                      "Serviços.[Descrição do Serviço] As Descricao, Municipios.[Descrição Do municipio] As Municipio, Municipios.[Uf], " & _
                      "Geral.[Rg e ie], Geral.[Cpf e cnpj], Geral.[Cep] As Cep, [Serviços da Venda].Quantidade As Qtde, 1 As Tipo, " & _
                      "[Serviços da Venda].[Valor Unitario] As Valor, Vendas.[Data da Venda] As Emissao, Vendas.[Valor Total dos Produtos] As TotalProdutos, Vendas.[Valor Total dos Serviços] As TotalServicos, " & _
                      "Vendas.[Valor do Desconto] As Desconto, Vendas.[Valor do Frete] As Frete, Vendas.[Sequencia da Transportadora] As Transporte " & _
                      "FROM Vendas LEFT JOIN Geral ON Vendas.[Sequencia Do Geral] = Geral.[Sequencia Do Geral] " & _
                      "LEFT JOIN Municipios ON Geral.[Sequencia Do Municipio] = Municipios.[Sequencia Do municipio] " & _
                      "LEFT JOIN [Serviços da Venda] ON Vendas.[Sequencia da Venda] = [Serviços da Venda].[Sequencia da Venda] " & _
                      "LEFT JOIN Serviços ON [Serviços da Venda].[Sequencia do Serviço] = Serviços.[Sequencia do Serviço] " & _
                      "WHERE [Serviços da Venda].[Sequencia Do Serviço] > 0 GROUP BY Vendas.[Valor Total da Venda], Vendas.[Valor total icms st], Vendas.[Valor total do ipi], Vendas.[Valor total do ISS], Vendas.[Sequencia da Venda], Vendas.[Sequencia Do Geral], " & _
                      "[Razão Social], [Fone], [Endereço], [Nro], [Bairro], [Serviços da Venda].[Sequencia do Serviço], " & _
                      "Serviços.[Descrição do Serviço], Municipios.[Descrição Do municipio], [Uf], [Rg e ie], [Cpf e cnpj], Geral.Cep, " & _
                      "[Serviços da Venda].Quantidade, Tipo, [Serviços da Venda].[Valor Unitario], Vendas.[Data da Venda], Vendas.[Valor Total dos Produtos], Vendas.[Valor Total dos Serviços], Vendas.[Valor do Desconto], Vendas.[Valor do Frete], Vendas.[Sequencia da Transportadora] Union "          
          SQL = SQL & "SELECT Vendas.[Valor Total da Venda] As TotalVenda, Vendas.[Valor total icms st] As ST, Vendas.[Valor total do ipi] As IPI, Vendas.[Valor total do ISS] As ISS, Vendas.[Sequencia da Venda], Vendas.[Sequencia Do Geral], Geral.[Razão Social], " & _
                      "Geral.[Fone], Geral.[Endereço], Geral.Nro, Geral.Bairro, [Produtos da Venda].[Sequencia do Produto] As SeqItem, " & _                 
                      "Produtos.[Descrição] As Descricao, Municipios.[Descrição Do municipio] As Municipio, Municipios.[Uf], " & _
                      "Geral.[Rg e ie], Geral.[Cpf e cnpj], Geral.[Cep] As Cep, [Produtos da Venda].Quantidade As Qtde, 2 As Tipo, " & _
                      "[Produtos da Venda].[Valor Unitario] As Valor, Vendas.[Data da Venda] As Emissao, Vendas.[Valor Total dos Produtos] As TotalProdutos, Vendas.[Valor Total dos Serviços] As TotalServicos, " & _
                      "Vendas.[Valor do Desconto] As Desconto, Vendas.[Valor do Frete] As Frete, Vendas.[Sequencia da Transportadora] As Transporte " & _
                      "FROM Vendas LEFT JOIN Geral ON Vendas.[Sequencia Do Geral] = Geral.[Sequencia Do Geral] " & _
                      "LEFT JOIN Municipios ON Geral.[Sequencia Do Municipio] = Municipios.[Sequencia Do municipio] " & _
                      "LEFT JOIN [Produtos da Venda] ON Vendas.[Sequencia da Venda] = [Produtos da Venda].[Sequencia da Venda] " & _
                      "LEFT JOIN Produtos ON [Produtos da Venda].[Sequencia do Produto] = Produtos.[Sequencia do Produto] " & _
                      "WHERE [Produtos da Venda].[Sequencia Do Produto] > 0 GROUP BY Vendas.[Valor Total da Venda], Vendas.[Valor total icms st], Vendas.[Valor total do ipi], Vendas.[Valor total do ISS], Vendas.[Sequencia da Venda], Vendas.[Sequencia Do Geral], " & _
                      "[Razão Social], [Fone], [Endereço], [Nro], [Bairro], [Produtos da Venda].[Sequencia do Produto], " & _
                      "Produtos.[Descrição], Municipios.[Descrição Do municipio], [Uf], [Rg e ie], [Cpf e cnpj], Geral.Cep, " & _
                      "[Produtos da Venda].Quantidade, Tipo, [Produtos da Venda].[Valor Unitario], Vendas.[Data da Venda], Vendas.[Valor Total dos Produtos], Vendas.[Valor Total dos Serviços], Vendas.[Valor do Desconto], Vendas.[Valor do Frete], Vendas.[Sequencia da Transportadora]) A " & _
                      "WHERE " & Filtro()
        SQL = SQL & " ORDER BY Tipo Desc, SeqItem" 
        vgQueryAtual = SQL
        CarregaParcelas 
        
        botOk_Click
End Sub

Private Sub CarregaParcelas()
   On Error Resume Next
   
   Set Parcelas = vgDb.OpenRecordSet("SELECT [Numero da parcela], [Data de vencimento], [Valor da parcela], [Especie de Pagamento] " & _
                                     "FROM [Parcelas da Venda] INNER JOIN Vendas ON [Parcelas da Venda].[Sequencia da Venda] = Vendas.[Sequencia da Venda] " & _                                     
                                     "WHERE [Parcelas da Venda].[Sequencia da Venda] = " & VendaTela)
End Sub

Private Function Filtro() As String   
   Filtro = "[Sequencia da Venda] = " & VendaTela
End Function

Private Function ImprimeData() As Variant
   If vgFP.vgContaPg > 1 Then
      Vendas.MovePrevious
      ImprimeData = Vendas!Emissao
      Vendas.MoveNext
   Else
      ImprimeData = Vendas!Emissao
   End If
End Function


Private Function DefineTipo() As String
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Tipo de Venda] As Tp From Vendas Where [Sequencia da Venda] = " & VendaTela)
 
  If Tb!Tp = 0 Then' Venda
     DefineTipo = "Venda - Nº.: " & Format(Vendas![Sequencia da Venda],"0000000")
  Else' Orçamento
     DefineTipo = "Orçamento - Nº.: " & Format(Vendas![Sequencia da Venda],"0000000")   
  End If
  
End Function

Private Function DefineTipo2() As String
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Tipo de Venda] As Tp From Vendas Where [Sequencia da Venda] = " & VendaTela)
 
  If Tb!Tp = 0 Then' Venda
     DefineTipo2 = "Totais da Venda" 
  Else' Orçamento
     DefineTipo2 = "Totais do Orçamento"    
  End If
  
End Function

Private Function MostraObs() As String
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Observações] As Obs From Vendas Where [Sequencia da Venda] = " & VendaTela)
 MostraObs = Tb!Obs
 
End Function

Private Function MostraPclas() As String

 Do While Not Parcelas.EOF
    MostraPclas = MostraPclas & " " & Format(Parcelas![Numero da Parcela], "00") & "  " & Parcelas![Data de Vencimento] & "   R$ " & Format(Parcelas![Valor da Parcela], "##,###,##0.00") & " - " & Parcelas![Especie de Pagamento] & vbCrLf & "---------------------------------------------" & vbCrLf
 Parcelas.MoveNext
 Loop
 Parcelas.MoveFirst

End Function


Private Function Pagto() As String
Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("SELECT [Forma de Pagamento] FPagto FROM Vendas Where [Sequencia da Venda] = " & VendaTela)
 
  If Tb!FPagto = "Vista" Then
     Pagto = " A Vista"
  Else
     Pagto = " A Prazo"    
  End If

End Function


Private Function MostraGeral() As String
  Dim Tb As New GRecordSet, Geral As New GRecordSet
  
  Set Tb = vgDb.OpenRecordSet("SELECT * From Vendas Where [Sequencia da Venda] = " & VendaTela)
  Set Geral = vgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do Geral] = " & Tb![Sequencia Do Geral])

  MostraGeral = Geral![Razão Social]
 
End Function


Private Function MostraVendedor() As String
  Dim Tb As New GRecordSet, Vendedor As New GRecordSet
  
  Set Tb = vgDb.OpenRecordSet("SELECT * From Vendas Where [Sequencia da Venda] = " & VendaTela)
  Set  Vendedor = vgDb.OpenRecordSet("SELECT [Razão Social] From Geral Where [Sequencia do Geral] = " & Tb![Sequencia Do Vendedor])

  MostraVendedor = "Vendedor -> " & Vendedor![Razão Social]
 
End Function

'prepara o intervalo de páginas a ser impresso
Public Property Let vgIntervalPag(ByVal NovoValor As String)
   Dim x As String, z As String, i As Integer, j As Integer, k As Integer
   ReDim vgPagRef(0) As Integer
   vgIntPag$ = NovoValor$
   x$ = Substitui(vgIntPag$, ";", ",", UM_A_UM)
   j = 0
   Do While Len(x$) > 0
      z$ = Parse(x$, ",")
      i = InStr(z$, "-")
      If i > 0 Then
         k = Val(Mid(z$, i + 1))
         i = Val(Left(z$, i - 1))
      Else
         i = Val(z$)
         k = Val(z$)
      End If
      For i = i To k Step ((i > k) ^ (i > k))
         j = j + 1
         ReDim Preserve vgPagRef(j) As Integer
         vgPagRef(j) = i
      Next
   Loop
End Property

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   vgDestRel = Val(labopcPainel1.Caption)
   vgModoInt = Val(labopcPainel2.Caption)
   vgIntPag = txtCampo(0).Value
   vgCopias = txtCampo(1).Value
   VendaTela = txtCampo(2).Value
   vgIntervalPag = vgIntPag
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 2
         txtCampo(2).Filter = "[Sequencia da Venda] > " & 0 & ""
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
      If vgModoInt = 1 Then
         vgOk = (Len(vgIntPag) > 0)
         vgMsg$ = "Esta informação é necessária"
         If Not vgOk Then txtCampo(0).SetFocus
      End If
      If vgOk Then
         If vgDestRel = 1 Then
            vgOk = (vgCopias > 0)
            vgMsg$ = "Número de cópias tem que ser maior que ZERO!"
            If Not vgOk Then txtCampo(1).SetFocus
         End If
      End If
      If vgOk Then
         vgOk = (VendaTela > 0)
         vgMsg$ = "Selecione a Venda / OS"
         If Not vgOk Then txtCampo(2).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
      opcPainel1(0).Value = True
      opcPainel1(0).CtPri.TabStop = False
      opcPainel2(0).Value = True
      opcPainel2(0).CtPri.TabStop = False
      txtCampo(2).Value = frmVendas.vgtb![Sequencia da Venda]
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

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean
   On Error Resume Next                           'prepara para possiveis erros
   Ok = (vgModoInt = 1)
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (vgDestRel = 1)
   txtCampo(1).Enabled = (Ok And txtCampo(1).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   botOk.Visible = (False)
   Label(6).Visible = (False)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Then
      ExecutaVisivel
      ExecutaPreValidacao
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 1
            If ValBrasil(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = 1
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
            End If
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
   ExecutaPreValidacao                            'habilita/desabilita campos
End Sub


'evento - quando uma opção for selecionada
Private Sub opcPainel1Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel1(Index).Locked Then
      opcPainel1(Val(labopcPainel1.Caption)).Value = True
   Else
      If Val(labopcPainel1.Caption) <> opcPainel1(Index).BookMark Then
         labopcPainel1.Caption = Str$(opcPainel1(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaVisivel
         ExecutaPreValidacao
         opcPainel1(Index).Change
      End If
   End If
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
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
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel1Cp_LostFocus(Index As Integer)
   opcPainel1(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcPainel2Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel2(Index).Locked Then
      opcPainel2(Val(labopcPainel2.Caption)).Value = True
   Else
      If Val(labopcPainel2.Caption) <> opcPainel2(Index).BookMark Then
         labopcPainel2.Caption = Str$(opcPainel2(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaVisivel
         ExecutaPreValidacao
         opcPainel2(Index).Change
      End If
   End If
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
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
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel2Cp_LostFocus(Index As Integer)
   opcPainel2(Index).LostFocus
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Dim x As String
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(19860)
   vgFormID = 711
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_RELAT
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   DesligaRepara
   ReDim vgBookPg(0) As Variant, vgPodeVoltar(0) As Boolean
   Load vgFP
   With vgFP
      vgIntervalPag = ""
      Set .vgFormRel = Me
      .vgContaPg = 0
      vgWidthRel = 520 + .ScaleX(196.9841, vbMillimeters, .ScaleMode)
      vgHeightRel = 720 + .ScaleY(286.7901, vbMillimeters, .ScaleMode)
      .picAux.Width = vgWidthRel
      .picAux.Height = vgHeightRel
      .Caption = LoadGasString(7505) + Me.Caption
      .picInterna.Move 0, 0, vgWidthRel, vgHeightRel
      .picAux.ScaleMode = vbMillimeters
      .picAux.Scale (0, 0)-(196.9841, 286.7901)
      .txtFixa(0).Width = 11070
      With .txtFixa(0).Font
         .Name = "Courier New"
         .CharSet = 0
         .Weight = 400
         .Size = 11.25
         .Underline = 0
         .Italic = 0
         .StrikeThrough = 0
      End With
      Load .txtFixa(1)
      .txtFixa(1).Width = 5640
      With .txtFixa(1).Font
         .Name = "Courier New"
         .CharSet = 0
         .Weight = 400
         .Size = 9.75
         .Underline = 0
         .Italic = 0
         .StrikeThrough = 0
      End With
      Load .txtFixa(2)
      .txtFixa(2).Width = 7740
      With .txtFixa(2).Font
         .Name = "Courier New"
         .CharSet = 0
         .Weight = 400
         .Size = 9.75
         .Underline = 0
         .Italic = 0
         .StrikeThrough = 0
      End With
   End With
   vgWidthRel = vgWidthRel - 520
   vgHeightRel = vgHeightRel - 720
   vgVisivel = True
   Label(4).Caption = Printer.DeviceName
   vgSituacao = ACAO_NAVEGANDO
   Set botOk.Picture = LoadPicture(LoadGasPicture(500))
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(11365))
   Set botConfImp.Picture = LoadPicture(LoadGasPicture(540))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(520))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(11366))
   Load vgFP.picFixa(1)
   Set vgFP.picFixa(1).Picture = LoadPicture(LoadGasPicture(11367))
   Set txtvgDestRel = opcPainel1(0)
   Set lblAjuste = Label(6)
   Set txtVenda = txtCampo(2)
   vgFP.picFixa(0).AutoSize = True
   vgFP.picFixa(0).ScaleMode = vbMillimeters
   vgQueryAtual = "SELECT * " + _
                     "FROM VENDAS"
   Set Vendas = vgDb.OpenRecordSet("SELECT TOP 1" + Mid$(vgQueryAtual, 7))
   DefineControles
   vgTooltips.Create
   vgTooltips.AddTool botOk, 0, LoadGasString(19865)
   botOk.Caption = LoadGasString(19866)
   Label(0).Caption = LoadGasString(19867)
   opcPainel1(0).Caption = LoadGasString(19868)
   opcPainel1(1).Caption = LoadGasString(19869)
   Label(1).Caption = LoadGasString(19870)
   opcPainel2(0).Caption = LoadGasString(19871)
   opcPainel2(1).Caption = LoadGasString(19872)
   Label(2).Caption = LoadGasString(19873)
   Label(3).Caption = LoadGasString(19874)
   vgTooltips.AddTool botConfImp, 0, LoadGasString(19875)
   vgTooltips.AddTool botCancela, 0, LoadGasString(19876)
   botCancela.Caption = LoadGasString(19877)
   Label(5).Caption = LoadGasString(19878)
   Label(6).Caption = LoadGasString(19879)
   Botao(0).Caption = LoadGasString(19880)
   labtxtCampo2.Caption = LoadGasString(19881)
   vgSQL$ = Vendas.Source
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
   ExecutaPreValidacao
   txtCampo(1).Value = 1
   Screen.MousePointer = vbDefault
   vgUltimoTabIndex = -1                          'controla lost focus de campos optativos
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

   Set opcPainel1(0).CtPri = opcPainel1Cp(0)
   Set opcPainel1(0).CtFdo = labopcPainel1
   opcPainel1(0).DataType = 6
   opcPainel1(0).BookMark = 0

   Set opcPainel1(1).CtPri = opcPainel1Cp(1)
   Set opcPainel1(1).CtFdo = labopcPainel1
   opcPainel1(1).DataType = 6
   opcPainel1(1).BookMark = 1

   Set opcPainel2(0).CtPri = opcPainel2Cp(0)
   Set opcPainel2(0).CtFdo = labopcPainel2
   opcPainel2(0).DataType = 6
   opcPainel2(0).BookMark = 0

   Set opcPainel2(1).CtPri = opcPainel2Cp(1)
   Set opcPainel2(1).CtFdo = labopcPainel2
   opcPainel2(1).DataType = 6
   opcPainel2(1).BookMark = 1

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 0
   txtCampo(0).Mask = ""
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtScr = ScrVtxtCampo1
   txtCampo(1).LargeChange = 10
   txtCampo(1).SmallChange = 1
   txtCampo(1).DataType = 3
   txtCampo(1).Mask = ""
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   txtCampo(1).Value = 1

   Set txtCampo(2).CtTit = labtxtCampo2
   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999999"
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = "Sequencia da Venda"
   txtCampo(2).ListFields = "Sequencia da Venda"
   txtCampo(2).OrderFields = "Sequencia da Venda"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Vendas"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub

'inicia impressão do relatório/etiqueta
Public Sub ProssegueExecucao(vgObSaida As Object, Optional ByVal vgNumPg As IMPRIME_PAGINA = 0)
   Dim vgPsm As Integer, vgPsz As Integer, vgPfn As String, vgPor As Integer, x As String, vgSemBookMark As Integer
   Dim i As Integer, vgLimiteInf As Single, vgSqlFiltro As String, vgTemFimRel As Boolean
   Dim vgNParc As Integer, vgNVez As Integer, vgNCop As Integer
   Dim vgEmPapel As Boolean, vgPagOk As Boolean, vgPV As Boolean, vgDifProxReg As Single, vgNumPag As Long, vgSemControlePg As Boolean, Ok As Boolean
   Dim j As Integer, m As Integer, vgQt As Integer, vgPagOrig As Integer, vgNovaPg As Integer, k As Integer, vgIniProc As Boolean, vgTestaPriCab As Boolean
   Dim vgPosLinMemo0 As Single, vgQdeLMemo0 As Long, vgLMemo0 As String, vgCL0 As Long, vgPula0 As Boolean, vgTF0 As Single
   Dim vgPriQb20 As Integer, vgQuebra20 As Variant, vgPosLinMemo1 As Single, vgQdeLMemo1 As Long, vgLMemo1 As String, vgCL1 As Long, vgPula1 As Boolean, vgTF1 As Single
   Dim vgTot1 As Double, vgPosLinMemo2 As Single, vgQdeLMemo2 As Long, vgLMemo2 As String, vgCL2 As Long, vgPula2 As Boolean, vgTF2 As Single
   vgCL0 = 0
   vgCL1 = 0
   vgCL2 = 0
   Screen.MousePointer = vbHourglass              'mostra ampulheta
   vgCancelou = False
   vgTemFimRel = False                            'se tem pendência para ser impressa como fim de página
   vgSemBookMark = 0
   If TypeOf vgObSaida Is Printer Then            'estamos em preview ou impressora?
      vgEmPapel = True                            'imprimindo na impressora, confirma
      GeraLogAcao LoadGasString(296) + " Imprimir" 
   Else
      vgEmPapel = False
      vgObSaida.Cls
      vgFP.scrH.Value = 0                         'acerta scroll do preview
      vgFP.scrV.Value = 0
   End If
   vgDataInicial = Now                            'data e hora do inicio do relatório
   vgNovaPg = -1
   vgIniProc = False                              'flag para informar que iniciou procura em navegação (próxima, anterior, última...)

   If vgNumPg = G_IMP_ATUA_PG Then                'solicitou impressão da página atual
      vgNumPg = vgFP.vgContaPg
   End If

   vgNCop = 0                                     'controla num de cop se impressora não tem esta prop
   vgSemControlePg = (False)

OutraCopia:
   On Error GoTo DeuErro                          'vamos cercar erros...

   'abre os recordsets
   vgPV = (vgNumPg = 0 Or (vgEmPapel And vgFP.vgContaPg > 0)) 'veio do botao de prosseguir?
   If vgPV And Not (vgEmPapel And vgFP.vgContaPg > 0)Then
      If TypeOf vgObSaida Is PictureBox Then                  'estamos em preview ou impressora?
         GeraLogAcao LoadGasString(297) + " Imprimir"
      End If
      x$ = vgQueryAtual
      If Len(vgUltimoFiltro) > 0 Then
         x$ = InsereSQL$(x$, EXP_WHERE, vgUltimoFiltro$)
      End If
      vgOrdemInicial = ExtraiSQL(x$, EXP_ORDERBY)
      If Len(vgUltimaOrdem) > 0 Or Len(vgOrdemInicial) > 0 Then
         x$ = InsereSQL(x$, EXP_ORDERBY, vgOrdemInicial + IIf(Len(vgOrdemInicial) > 0 And Len(vgUltimaOrdem) > 0, ", ","") + vgUltimaOrdem)
      End If
      Set Vendas = vgDb.OpenRecordSet(x$)
      If Vendas.RecordCount = 0 Then
       vgPrimeiroCab = 1
       Screen.MousePointer = vbNormal
       Exit Sub
      End If
      ReDim vgBookPg(0) As Variant                            'cria variável para salvar bookmark
      ReDim vgPodeVoltar(0) As Boolean                        'controle de quebra em filhos
      GoSub IniciaImpressao
   Else
      If vgNumPg >= 0 And (vgEmPapel And vgFP.vgContaPg > 0) Then
         GoSub IniciaImpressao
      End If
      If vgModoInt = 1 Then
         vgPagOrig = vgPagRef(vgFP.vgPagRef)
      Else
         vgPagOrig = vgFP.vgPagRef
      End If
      If vgNumPg > 0 Then                                     'se estiver indo para alguma página específica
         If vgModoInt = 1 Then
            j = 0
            For i = 1 To UBound(vgPagRef)                     'vamos procurar a página solicitada no intervalo de páginas permitidas para o relatório...
               If vgPagRef(i) = vgNumPg Then
                  j = i
                  If i >= vgPagOrig Then                      'se a página encontrada estiver após a página atualamente exibida no relatório...
                     Exit For                                 'cai fora
                  End If
               End If
            Next
            If j = 0 Then
               vgFP.vgContaPg = vgFP.vgContaPg - 1            'para evitar problema na SalvaPagina
               GoTo NaoAchouPg
            End If
            vgFP.vgPagRef = j
         End If
         vgFP.vgNovaPag = vgNumPg
         If vgNumPg > UBound(vgBookPg()) Then
            vgNumPg = UBound(vgBookPg())
         End If
         vgNovaPg = vgNumPg - 1
         vgFP.vgProcuraPg = (vgNumPg <> vgFP.vgNovaPag)
      ElseIf vgNumPg = G_IMP_PROX_PG Or vgNumPg = G_IMP_ANT_PG Then
         i = (vgNumPg = G_IMP_ANT_PG) ^ (vgNumPg = G_IMP_ANT_PG) 'i=+1 para G_IMP_PROX_PG e i=-1 para G_IMP_ANT_PG
         If vgModoInt = 1 Then
            vgFP.vgProcuraPg = (vgPagRef(vgFP.vgPagRef + (i * 1)) > UBound(vgBookPg()))
            If vgFP.vgProcuraPg Then
               vgIniProc = True                                  'flag para informar que iniciamos procura
               vgFP.vgNovaPag = vgPagRef(vgFP.vgPagRef + (i * 1))
               vgNovaPg = UBound(vgBookPg) - 1
               vgFP.vgPagRef = vgFP.vgPagRef + (i * 1)
            Else
               vgNovaPg = vgPagRef(vgFP.vgPagRef + (i * 1)) - 1
            End If
         Else
            vgFP.vgProcuraPg = False
            vgNovaPg = vgFP.vgContaPg + (i * 1) - 1
         End If
      ElseIf vgNumPg = G_IMP_PRI_PG Or vgNumPg = 0 Then
         If vgEmPapel And vgFP.vgContaPg > 0 Then                'se for em papel... e tiver vindo do botão IMPRIME do preview...
            vgPagOrig = vgFP.vgPagRef
         End If
         vgFP.vgProcuraPg = False
         If vgModoInt = 1 Then
            vgNovaPg = vgPagRef(1) - 1
         Else
            vgNovaPg = 0
         End If
      ElseIf vgNumPg = G_IMP_ULT_PG Then
         If vgModoInt = 1 Then
            vgFP.vgProcuraPg = False
            If vgFP.vgUltPagVista <= vgPagRef(UBound(vgPagRef())) Then 'se ainda não visualizou a última página...
               vgNovaPg = UBound(vgBookPg()) - 1                       'vamos posicionar em uma página antes da última disponível (vai incrementar depois)
            Else
               vgNovaPg = vgPagRef(UBound(vgPagRef())) - 1             'posicionar uma página antes da última página do relatório
            End If
         Else
            If UBound(vgBookPg()) > 0 Then
               vgNovaPg = UBound(vgBookPg()) - 1
            End If
            vgFP.vgProcuraPg = True
            vgIniProc = True                                           'flag para informar que iniciamos procura
            vgFP.vgNovaPag = 0
         End If
      End If
   End If
   If vgNovaPg >= 0 Then
      If vgPodeVoltar(vgNovaPg + 1) = False And vgNovaPg > 0 Then
         If Not vgFP.vgProcuraPg Then                                  'se nao esta indo para pg ou ultima pg
            vgFP.vgProcuraPg = True
            vgIniProc = True                                           'flag para informar que iniciamos procura
            vgFP.vgNovaPag = vgNovaPg + 1
         End If
         Do While vgPodeVoltar(vgNovaPg + 1) = False                   'acerto para não deixar voltar página com quebra em filho
            vgNovaPg = vgNovaPg - 1
         Loop
         If vgModoInt = 1 Then
            If vgNumPg < 0 Then
               vgFP.vgPagRef = vgFP.vgPagRef + 1
            End If
         Else
            vgFP.vgPagRef = vgFP.vgNovaPag
         End If
      End If
      vgFP.vgContaPg = vgNovaPg
      If Len(vgBookPg(vgNovaPg + 1)) > 0 Then
         Vendas.BookMark = vgBookPg(vgNovaPg + 1)
      Else
         vgSemBookMark = 1
      End If
      vgNumPag = vgNovaPg + 1
      vgCL0 = vgMemo0(vgNumPag)
      vgCL1 = vgMemo1(vgNumPag)
      vgCL2 = vgMemo2(vgNumPag)
   Else
      vgFP.vgPagRef = 1
   End If
   vgFP.botPara.Enabled = (vgFP.vgProcuraPg Or vgNumPg = G_IMP_ULT_PG) 'habilita/desabilita stop
   vgPrimeiroCab = True

Volta:

   'Imprimir
   If Vendas.RecordCount > 0 Then
      If vgNumPg = 0 And Not (vgEmPapel And vgFP.vgContaPg > 0) Then
         Vendas.MoveLast
         Vendas.MoveFirst
      End If
      If vgEmPapel And vgNumPg = 0 And vgPrimeiroCab Or (vgNumPg = vgPagOrig And vgEmPapel) Then
         Set vgFG = New frmGauge
         vgQReg = 0
         Load vgFG
         Set vgFG.vgFrmOrigem = Me
         vgFG.MostraGauge
      End If
      vgLimiteInf = 286
      vgDifProxReg = 0
      On Error Resume Next
      Do While Not Vendas.EOF Or vgSemBookMark = 1
         vgFP.vgFimDoArq = False
         If vgPrimeiroCab Or vgPosLin > vgLimiteInf Then
            vgPosLin = vgPosLin - vgDifProxReg
            vgDifProxReg = 0
            vgPagOk = True
            vgTestaPriCab = True
            GoSub VerificaFimPg
            GoSub OutraFolha
            On Error Resume Next
         End If
         If vgSemBookMark = 0 Then
            If vgCL1 = 0 And vgCL2 = 0 Then
            End If
            vgQuebra20 = Vendas![Sequencia da Venda]
            vgPriQb20 = 0
            vgTot1 = 0
            If vgPrimeiroCab Then
               vgTot1 = vgPgvgTot1(vgNumPag)
            End If
            Do While Not Vendas.EOF
               If vgQuebra20 <> Vendas![Sequencia da Venda] Then Exit Do
               If vgPosLin > vgLimiteInf - 4.23  Or (vgPriQb20 < 2 And vgPosLin > vgLimiteInf  - 14.02) Then
                  vgPagOK = False
                  vgPosLin = vgPosLin - vgDifProxReg
                  vgDifProxReg = 0
                  vgTestaPriCab = False
                  GoSub VerificaFimPg
                  GoSub OutraFolha
                  On Error Resume Next
                  vgPriQb20 = vgPriQb20 - 1
               End If
               If vgPriQb20 < 2 Then
                  GoSub CabQb20
               End If
               If vgCL2 = 0 Then
                  If vgCL1 = 0 Then
                     vgTot1 = vgTot1 + Vendas!Qtde * Vendas!Valor
                     Imprime Format(Vendas!SeqItem,"000000"), 1.32, vgPosLin, "F=Courier New|S=9.75|L=16", vgObSaida, "999999", ,, vgFP.vgProcuraPg
                     Imprime Vendas!Qtde, 135.75, vgPosLin, "F=Courier New|S=9.75|L=22|JD", vgObSaida, "999.999,9999", ,, vgFP.vgProcuraPg
                     Imprime Vendas!Valor, 165.03, vgPosLin, "F=Courier New|S=9.75|L=28|JD", vgObSaida, "999.999,9999", ,, vgFP.vgProcuraPg
                     Imprime Vendas!Qtde * Vendas!Valor, 196.66, vgPosLin, "F=Courier New|S=9.75|L=30|JD", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
                  End If
                  vgPula0 = False
                  vgPosLinMemo0 = -50
                  vgTF0 = 0
                  vgFP.txtFixa(0).Text = MostraObs()
                  vgQdeLMemo0 = vgFP.ContaLinhasMemo(0)
                  vgPula1 = False
                  vgPosLinMemo1 = 0
                  vgTF1 = 0
                  vgFP.txtFixa(1).Text = Vendas!Descricao
                  vgQdeLMemo1 = vgFP.ContaLinhasMemo(1)
                  vgPula2 = False
                  vgPosLinMemo2 = -16.14
                  vgTF2 = 0
                  vgFP.txtFixa(2).Text = MostraPclas()
                  vgQdeLMemo2 = vgFP.ContaLinhasMemo(2)
                  Do While vgCL1 <= vgQdeLMemo1
                     If (vgPosLin + (vgLimiteInf - (vgLimiteInf - (vgPosLinMemo1 + vgTF1))) > vgLimiteInf) Or vgCL1 > vgQdeLMemo1 Then
                        vgPula1 = True
                     End If
                     If vgPula1 Then
                        vgPagOK = False
                        vgPula0 = False
                        vgPosLinMemo0 = 0
                        vgPula1 = False
                        vgPosLinMemo1 = 0
                        vgPula2 = False
                        vgPosLinMemo2 = 0
                        vgTestaPriCab = False
                        GoSub VerificaFimPg
                        GoSub OutraFolha
                        On Error Resume Next
                        vgPriQb20 = vgPriQb20 - 1
                     End If
                     If vgPriQb20 < 2 Then
                        GoSub CabQb20
                     End If
                     If Not vgPula1 And vgCL1 <= vgQdeLMemo1 Then
                        vgFP.txtFixa(1).Text = Vendas!Descricao
                        vgTF1 = Imprime(vgFP.ExtraiLinhaMemo(1, vgCL1), 17.98, vgPosLin + vgPosLinMemo1, "F=Courier New|S=9.75", vgObSaida, , ,, vgFP.vgProcuraPg)
                        vgPosLinMemo1 = vgPosLinMemo1 + vgTF1
                        vgCL1 = vgCL1 + 1
                     End If
                  Loop
                  vgCL1 = 0
                  vgPosLin = vgPosLin + Maior(vgPosLinMemo1, 4.23)
                  vgDifProxReg = 0
               End If
               Vendas.MoveNext
               vgDifProxReg = -1
               vgPosLin = vgPosLin + vgDifProxReg
            Loop
            vgPosLin = vgPosLin - vgDifProxReg
            vgDifProxReg = 0
            vgPosLin = vgPosLin + .26
            If vgPosLin + 4.76 > vgLimiteInf Then
             vgPagOk = False
             vgTestaPriCab = False
             vgTemFimRel = True
             GoSub VerificaFimPg
             GoSub OutraFolha
            End If
            Imprime "", 175.13, vgPosLin, "LIN=B0/" + Trim(Str( 175.13)) + "/" + Trim$(Str$(vgPosLin)) + "/" + Trim(Str(196.82)) + "/" + Trim$(Str$(vgPosLin)), vgObSaida, , ,, vgFP.vgProcuraPg
            Imprime vgTot1, 196.6, vgPosLin +.52, "F=Courier New|W=700|S=9.75|L=31|JD", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
            vgPosLin = vgPosLin + 4.76
            vgDifProxReg = 0
            vgDifProxReg = -1
            vgPosLin = vgPosLin + vgDifProxReg
         Else
            vgSemBookMark = 0
         End If
         If vgEmPapel And vgNumPg = 0 Then
            vgQReg = vgQReg + 1
            If vgModoInt = 1 Then
               If Not vgFp.vgProcuraPg Then
                  vgFG.pgb1.Value = vgFP.vgPagRef * 100 / UBound(vgPagRef())
               End If
            Else
               vgFG.pgb1.Value = vgQReg * 100 / Vendas.RecordCount
            End If
            If vgFG.vgQuerCancelar Then Exit Do
         End If
      Loop
      If vgEmPapel Then
         If vgFG.vgQuerCancelar Then GoTo CancelouRel
      End If
      vgPosLin = vgPosLin - vgDifProxReg
      vgDifProxReg = 0
      If vgPosLin > vgLimiteInf - 18.78 Then
         vgPagOk = False
         vgTemFimRel = True
         vgTestaPriCab = False
         GoSub VerificaFimPg
         GoSub OutraFolha
      End If
      vgPosLin = vgPosLin + 31.21
      Imprime "- de Acordo Com:", 65.87, vgPosLin +.52, "F=Courier New|W=700|S=11.25|L=42", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime MostraVendedor(), 110.31, vgPosLin, "F=Courier New|S=11.25|L=84|JC=84.12", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "", 110.31, vgPosLin +4.76, "LIN=B0/" + Trim(Str( 110.31)) + "/" + Trim$(Str$(vgPosLin +4.76)) + "/" + Trim(Str(194.44)) + "/" + Trim$(Str$(vgPosLin +4.76)), vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "", 108.73, vgPosLin +13.75, "LIN=B0/" + Trim(Str( 108.73)) + "/" + Trim$(Str$(vgPosLin +13.75)) + "/" + Trim(Str(197.61)) + "/" + Trim$(Str$(vgPosLin +13.75)), vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Empresa![Razão social], 107.67, vgPosLin +14.28, "F=Courier New|S=11.25|L=90|JC=89.94", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime MostraGeral(), 10.84, vgPosLin +14.28, "F=Courier New|S=11.25|L=89|JC=88.88", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "", 10.84, vgPosLin +13.75, "LIN=B0/" + Trim(Str( 10.84)) + "/" + Trim$(Str$(vgPosLin +13.75)) + "/" + Trim(Str(99.73)) + "/" + Trim$(Str$(vgPosLin +13.75)), vgObSaida, , ,, vgFP.vgProcuraPg
      vgPosLin = vgPosLin + 18.78
      vgDifProxReg = 0
      If vgFP.vgProcuraPg And UBound(vgBookPg) <= vgFP.vgContaPg Then               'ops... estava buscando por uma página, e ela não existe (chegou no final do arquivo)!
         If vgNumPg = G_IMP_ULT_PG Then
            vgPrimeiroCab = True
            vgFP.vgProcuraPg = True
            vgNumPg = UBound(vgBookPg())
            If Len(vgBookPg(vgNumPg)) = 0 Then     'se a última não tem bookmark, quer dizer que chegou em fim de arquivo mas tem complemento a imprimir em outra folha
               Do While Len(vgBookPg(vgNumPg)) = 0   'vamos pegar a última página com bookmark válido
                  vgNumPg = vgNumPg - 1
               Loop
               vgFP.vgContaPg = vgNumPg              'vamos voltar para última página com bookmark
               vgNumPg = G_IMP_PROX_PG               'e força navegação para próxima (que será a última)
               vgFP.vgProcuraPg = False              'sem procurar...
            End If
            GoTo OutraCopia
         Else
            vgPrimeiroCab = True
            vgFP.vgProcuraPg = False
            If vgModoInt = 1 Then
               vgPagRef(vgFP.vgPagRef) = -vgFP.vgContaPg - 1        'vamos forçar a exclusão de páginas maiores que a última no vetor de referências
               vgFP.vgContaPg = vgPagRef(vgFP.vgPagRef - 1) - 1
               j = 0
               Do While vgPagRef(vgFP.vgPagRef - 1 - j) - 1 <> 0 And vgPagRef(vgFP.vgPagRef - 1 - j) - 1 > UBound(vgBookPg) And vgFP.vgPagRef - j > 0
                  j = j + 1
                  vgFP.vgContaPg = vgPagRef(vgFP.vgPagRef - 1 - j) - 1
               Loop
               If vgFP.vgContaPg > 0 Then
                  Vendas.Bookmark = vgBookPg(vgFP.vgContaPg + 1)
               Else
                  vgFP.vgContaPg = 0
                  Vendas.MoveFirst
               End If
            End If
            If vgModoInt = 1 Then
               GoTo Volta
             ElseIf vgFP.vgNovaPag > vgFP.vgContaPg Then
               GoSub RestauraPag
               GoTo NaoAchouPg
            End If
         End If
      ElseIf vgEmPapel And vgNumPg = 0 And vgFP.vgPagRef < UBound(vgPagRef()) Then 'estava imprimindo intervalo de página e parece ter chegado no fim do arquivo
         If Not vgFG.vgQuerCancelar Then
            vgPosLin = vgLimiteInf + 99                  'força um cabec na próxima vez
            vgFP.vgContaPg = 0                           'zera o contador de página
            GoTo Volta                                   'e vamos voltar para imprimir a próxima página
         End If
      End If
   End If

CancelouRel:
   If vgEmPapel Then
      If vgNumPg > 0 And Not vgPrimeiroCab And Not vgFG.vgQuerCancelar Then
         Unload vgFG
         Set vgFG = Nothing
         vgObSaida.EndDoc
      Else
         If vgFG.vgQuerCancelar Then
            vgCancelou = True
            vgObSaida.KillDoc
         Else
            vgObSaida.EndDoc
         End If
         If vgNumPg = 0 And vgPagOrig > 0 Then                      'solicitou impressão de todo relatório pelo preview
            vgFP.vgPagRef = vgPagOrig
            If vgModoInt = 1 Then
               vgFP.vgContaPg = vgPagRef(vgPagOrig)                     'e arruma o número da página
            Else
               vgFP.vgContaPg = vgPagOrig                               'e arruma o número da página
            End If
         End If
         vgNCop = vgNCop - 1
         If vgFG.vgQuerCancelar = False And vgNCop > 0 Then
            Unload vgFG
            Set vgFG = Nothing
            GoTo OutraCopia
         End If
      End If
      With Printer
         .ScaleMode = vgPsm
         .Orientation = vgPor
         .PaperSize = vgPsz
         .FontName = vgPfn
      End With
   Else
      GoSub SalvaPagina
   End If
   GoTo SaiDaSub

VerificaFimPg:
   If Not vgEmPapel And (Not vgTestaPriCab Or vgPrimeiroCab = False)  Then
      GoSub SalvaPagina
      If Not vgFP.vgProcuraPg Then                      'se não está procurando por uma página
         If vgFP.vgContaPg <> vgFP.vgNovaPag Then                   'opa... parece que o usuário precionou o STOP
            If vgNumPg = G_IMP_ANT_PG Or vgNumPg = G_IMP_PROX_PG Then 'se estava movendo para próxima página ou anterior
               vgFP.vgContaPg = vgPagOrig                               'vamos restaurar a página original
               If vgNumPg = G_IMP_ANT_PG Then                           'restaura o contador
                  vgFP.vgPagRef = vgFP.vgPagRef + 1
               Else
                  vgFP.vgPagRef = vgFP.vgPagRef - 1
               End If
               GoSub SalvaPagina                    'ajusta a botoeira do preview novamente
            End If
         End If
         GoTo SaiDaSub                          'cai fora
      End If
   ElseIf vgEmPapel And vgNumPg > 0 And Not vgPrimeiroCab Then
      If vgFP.vgContaPg = vgFP.vgNovaPag Then GoTo CancelouRel
   ElseIf vgNumPg = 0 And Not vgPrimeiroCab Then
      If Not vgFP.vgProcuraPg Then
         If vgModoInt = 1 And vgFP.vgPagRef >= UBound(vgPagRef()) Then         'se estava imprimindo intervalo de páginas...
            GoTo CancelouRel                                                     'se chegou à última página do intervalo, cai fora
         End If
         GoSub SalvaInicioPg
      End If
   End If
   Return

OutraFolha:
   If vgEmPapel = False Then
      GoSub SalvaInicioPg
      If vgFP.vgProcuraPg And vgNumPg = G_IMP_ULT_PG Then
         vgFP.txtPag.Text = Trim$(Str$(vgFP.vgContaPg + 1))
         vgFP.txtPag.Refresh
      End If
   ElseIf Not vgSemControlePg  Then
      If Not vgPrimeiroCab And Not vgFP.vgProcuraPg Then vgObSaida.NewPage
   End If
   vgFP.vgContaPg = vgFP.vgContaPg + 1
   If vgFP.vgProcuraPg = False Then
    If vgModoInt = 0 Then
     vgFP.vgNovaPag = vgFP.vgContaPg
    Else
     If vgPagRef(vgFP.vgPagRef) < 0 Then
      GoSub ArrumaVetor
     Else
      If vgNumPg = G_IMP_PROX_PG Or vgNumPg > vgFP.vgContaPg Or (vgNumPg = 0 And vgModoInt = 0) Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       Else                                   'usuário cancelou busca pela próxima página
        vgFP.vgPagRef = vgFP.vgPagRef - 1
       End If
      ElseIf vgNumPg = G_IMP_ANT_PG Or (vgNumPg > 0 And vgNumPg < vgFP.vgContaPg) Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = vgFP.vgPagRef - 1
       Else                                   'usuário cancelou busca pela página anterior
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       End If
      ElseIf vgNumPg = G_IMP_PRI_PG Or vgNumPg = 0 Then
       If vgPrimeiroCab Then
        vgFP.vgPagRef = 1
       Else
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       End If
      ElseIf vgNumPg = G_IMP_ULT_PG Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = UBound(vgPagRef())
       Else
        'usuário cancelou busca para última página, vamos procurar o último bookmark existente no intervalo de páginas especificado
        For j = UBound(vgPagRef()) To 1 Step -1           'vamos correr as páginas de referência, de trás para frente
         If vgPagRef(j) <= UBound(vgBookPg) Then         'se tem bookmark para a página correspondente
          vgFP.vgPagRef = j                               'vamos utilizar essa página
          Exit For
         End If
        Next
       End If
      End If
     End If
     vgFP.vgNovaPag = vgPagRef(vgFP.vgPagRef)             'pega o número de página real referente ao número da página no intervalo
     If vgFP.vgNovaPag = 0 Then GoTo CancelouRel          'ops... vamos sair do relatório
     If vgFP.vgNovaPag <> vgFP.vgContaPg Then             'se a página atual não for a página procurada
      If vgFP.vgNovaPag <= UBound(vgBookPg()) Then       'se tem bookmark para a página procurada
       vgFP.vgContaPg = vgFP.vgNovaPag                    'vamos posicionar diretamente na nova página
       Do While vgPodeVoltar(vgFP.vgContaPg) = False      'acerto para não deixar voltar página com quebra em filho
        vgFP.vgContaPg = vgFP.vgContaPg - 1
       Loop
       vgFP.vgProcuraPg = (vgFP.vgContaPg <> vgFP.vgNovaPag)
       Vendas.Bookmark = vgBookPg(vgFP.vgContaPg)      'seta o bookmark correto
       If Not vgEmPapel Then
        vgFP.picAux.Cls                                    'limpa a área
        vgFP.picAux.Picture = LoadPicture("")              'para imprimir a página
       End If
      Else                                                'não tem bookmark...
       vgFP.vgProcuraPg = True                            'então vamos procurar
      End If
     Else
      If Not vgEmPapel Then
       vgFP.picAux.Cls                                     'limpa a área
       vgFP.picAux.Picture = LoadPicture("")               'para imprimir a página
      End If
     End If
    End If
   ElseIf vgFP.vgNovaPag = vgFP.vgContaPg Then            'se a página atual é a página procurada
    vgFP.vgProcuraPg = False                              'vamos cancelar a procura!
   End If
   If vgModoInt = 0 Then                                  'se não tem intervalo de páginas especificado
    vgNumPag = vgFP.vgContaPg
    vgFP.vgPagRef = vgFP.vgContaPg                      'com o mesmo número para página de referência
   Else                                                   'senão... tem intervalo de página
    vgNumPag = vgFP.vgNovaPag                             'então pela o número da nova página
   End If
   vgPosLin = 44.97
   Imprime "Empresa..:", 27.86, 15.87, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "CNPJ.....:", 27.86, 21.69, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "Endereço.:", 27.86, 27.51, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "Município:", 27.86, 33.06, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "E-Mail...:", 27.86, 38.35, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime Empresa!Email & "  Fone: " & Empresa!Fone, 52.53, 38.35, "F=Courier New|W=700|S=11.25|L=144", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime Empresa!Município & " - " & Empresa!Uf & " " & Empresa!Cep, 52.53, 33.06, "F=Courier New|W=700|S=11.25|L=144", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime Empresa!Rua  & " " & Empresa!Nro & " - " & Empresa!Bairro, 52.8, 27.51, "F=Courier New|W=700|S=11.25|L=144", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime Empresa!Cnpj, 53.17, 21.69, "F=Courier New|W=700|S=11.25|L=47", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime Empresa![Nome fantasia], 52.53, 15.87, "F=Courier New|W=700|S=11.25|L=144", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "Dt. Emissão", 166.49, 1.58, "F=Courier New|W=700|S=12|JC=29.89", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime CDate(ImprimeData), 168.07, 6.34, "F=Courier New|W=700|S=11.25|L=27|JC=26.71", vgObSaida, "99/99/9999", ,, vgFP.vgProcuraPg
   Imprime "Pag. Nº.: " & Format(vgFP.vgPagRef, "00"), 169.13, 11.11, "F=Courier New|W=700|S=12|JC=24.6", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime DefineTipo(), 46.11, 6.08, "F=Courier New|W=700|S=15.75|JC=102.64", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime vgFP.picFixa(1),  1.94, 15.07, "FIG=FigFixa1.jpg", vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "", .68, 43.12, "LIN=B0/" + Trim(Str( .68)) + "/" + Trim$(Str$(43.12)) + "/" + Trim(Str(196.98)) + "/" + Trim$(Str$(43.12)), vgObSaida, , ,, vgFP.vgProcuraPg
   Imprime "", 165.23, 1.32, "L=32|BOX=B0/L1/165.23/1.32/31.74/14.28", vgObSaida, , ,, vgFP.vgProcuraPg
   If vgFP.vgContaPg = 1 Then
      Imprime "Cliente....:", 1.94, vgPosLin, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas![Razão Social], 31.83, vgPosLin, "F=Courier New|S=11.25|L=163", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Endereço...:", 1.94, vgPosLin +5.02, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas![Endereço] & " " & Vendas![Nro] & " - " & Vendas![Bairro], 31.83, vgPosLin +5.02, "F=Courier New|S=11.25|L=163", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Município..:", 1.94, vgPosLin +10.84, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas![Municipio] & " - " & Vendas![Uf] & " CEP: " & Vendas!Cep & " Fone " & Vendas![Fone], 31.83, vgPosLin +10.84, "F=Courier New|S=11.25|L=163", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "CPF/CNPJ...:", 1.94, vgPosLin +16.4, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas![Cpf e cnpj], 31.83, vgPosLin +16.4, "F=Courier New|S=11.25|L=44", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "RG/I.E:", 78.13, vgPosLin +16.4, "F=Courier New|W=700|S=11.25", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas![Rg e ie], 96.91, vgPosLin +16.4, "F=Courier New|S=11.25|L=58", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Observações", .52, vgPosLin +46.29, "F=Courier New|W=700|S=12|L=196|JC=196.29|BOX=F14737632/.52/" + Str$(vgPosLin + 46.29) + "/196.29/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime DefineTipo2(), .52, vgPosLin +24.07, "F=Courier New|W=700|S=12|L=196|JC=196.29|BOX=F14737632/.52/" + Str$(vgPosLin + 24.07) + "/196.29/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Total dos Produtos", .79, vgPosLin +29.1, "F=Courier New|W=700|S=11.25|L=53|JC=52.69|BOX=B0/L1/.79/" + Str$(vgPosLin + 29.1) + "/52.69/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!TotalProdutos, .79, vgPosLin +33.33, "F=Courier New|W=700|S=11.25|L=53|JC=52.69|BOX=B0/L1/.79/" + Str$(vgPosLin + 33.33) + "/52.69/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Total dos Serviços", 53.43, vgPosLin +29.1, "F=Courier New|W=700|S=11.25|L=49|JC=48.67|BOX=B0/L1/53.43/" + Str$(vgPosLin + 29.1) + "/48.67/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!TotalServicos, 53.43, vgPosLin +33.33, "F=Courier New|W=700|S=11.25|L=49|JC=48.67|BOX=B0/L1/53.43/" + Str$(vgPosLin + 33.33) + "/48.67/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Total ICMS ST", 102.11, vgPosLin +29.1, "F=Courier New|W=700|S=11.25|L=47|JC=47.08|BOX=B0/L1/102.11/" + Str$(vgPosLin + 29.1) + "/47.08/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!ST, 102.11, vgPosLin +33.33, "F=Courier New|W=700|S=11.25|L=47|JC=47.08|BOX=B0/L1/102.11/" + Str$(vgPosLin + 33.33) + "/47.08/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Total de IPI", .79, vgPosLin +37.56, "F=Courier New|W=700|S=11.25|L=53|JC=52.69|BOX=B0/L1/.79/" + Str$(vgPosLin + 37.56) + "/52.69/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!IPI, .79, vgPosLin +41.79, "F=Courier New|W=700|S=11.25|L=53|JC=52.69|BOX=B0/L1/.79/" + Str$(vgPosLin + 41.79) + "/52.69/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Total do ISS", 53.43, vgPosLin +37.56, "F=Courier New|W=700|S=11.25|L=49|JC=48.67|BOX=B0/L1/53.43/" + Str$(vgPosLin + 37.56) + "/48.67/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!ISS, 53.43, vgPosLin +41.79, "F=Courier New|W=700|S=11.25|L=49|JC=48.67|BOX=B0/L1/53.43/" + Str$(vgPosLin + 41.79) + "/48.67/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Valor do Frete", 102.11, vgPosLin +37.56, "F=Courier New|W=700|S=11.25|L=47|JC=47.08|BOX=B0/L1/102.11/" + Str$(vgPosLin + 37.56) + "/47.08/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!Frete, 102.11, vgPosLin +41.79, "F=Courier New|W=700|S=11.25|L=47|JC=47.08|BOX=B0/L1/102.11/" + Str$(vgPosLin + 41.79) + "/47.08/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Valor do Desconto", 149.2, vgPosLin +29.1, "F=Courier New|W=700|S=11.25|L=48|JC=47.61|BOX=B0/L1/149.2/" + Str$(vgPosLin + 29.1) + "/47.61/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!Desconto, 149.2, vgPosLin +33.33, "F=Courier New|W=700|S=11.25|L=48|JC=47.61|BOX=B0/L1/149.2/" + Str$(vgPosLin + 33.33) + "/47.61/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "Valor Total", 149.2, vgPosLin +37.56, "F=Courier New|W=700|S=11.25|L=48|JC=47.61|BOX=B0/L1/149.2/" + Str$(vgPosLin + 37.56) + "/47.61/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime Vendas!TotalVenda, 149.2, vgPosLin +41.79, "F=Courier New|W=700|S=11.25|L=48|JC=47.61|BOX=B0/L1/149.2/" + Str$(vgPosLin + 41.79) + "/47.61/4.23", vgObSaida, "999.999.999,99", ,, vgFP.vgProcuraPg
      Imprime "", .68, vgPosLin +23.28, "LIN=B0/" + Trim(Str( .68)) + "/" + Trim$(Str$(vgPosLin +23.28)) + "/" + Trim(Str(196.98)) + "/" + Trim$(Str$(vgPosLin +23.28)), vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Financeiro / Parcelamento", 0, vgPosLin +73.54, "F=Courier New|W=700|S=12|L=197|JC=196.82|BOX=F14737632/0/" + Str$(vgPosLin + 73.54) + "/196.82/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      If Parcelas.RecordCount > 0 Then
         Imprime "Pc", 0, vgPosLin +79.62, "F=Courier New|W=700|S=11.25|L=10|JC=9.52", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      If Parcelas.RecordCount > 0 Then
         Imprime "Vencto", 10.45, vgPosLin +79.62, "F=Courier New|W=700|S=11.25|L=26", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      If Parcelas.RecordCount > 0 Then
         Imprime "R$ Valor / Especie", 37.83, vgPosLin +79.62, "F=Courier New|W=700|S=11.25|L=46", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      If Parcelas.RecordCount > 0 Then
         Imprime "Forma de Pagamento", 141, vgPosLin +78.83, "F=Courier New|W=700|S=14.25|L=55", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      If Parcelas.RecordCount > 0 Then
         Imprime Pagto(), 152.38, vgPosLin +84.12, "F=Courier New|W=700|S=14.25|L=31|JC=30.68", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      If Parcelas.RecordCount > 0 Then
         Imprime "", 139.15, vgPosLin +78.83, "L=56|BOX=B0/L1/139.15/" + Str$(vgPosLin + 78.83) + "/56.08/10.58", vgObSaida, , ,, vgFP.vgProcuraPg
      End If
      vgPula0 = False
      vgPosLinMemo0 = 51.32
      vgTF0 = 0
      vgFP.txtFixa(0).Text = MostraObs()
      vgQdeLMemo0 = vgFP.ContaLinhasMemo(0)
      vgPula1 = False
      vgPosLinMemo1 = 101.32
      vgTF1 = 0
      vgFP.txtFixa(1).Text = Vendas!Descricao
      vgQdeLMemo1 = vgFP.ContaLinhasMemo(1)
      vgPula2 = False
      vgPosLinMemo2 = 85.18
      vgTF2 = 0
      vgFP.txtFixa(2).Text = MostraPclas()
      vgQdeLMemo2 = vgFP.ContaLinhasMemo(2)
      Do While (vgCL0 <= vgQdeLMemo0 And (vgPosLinMemo0 + vgTF0) < (69.84 - vgTF0)) Or vgCL2 <= vgQdeLMemo2
         If (vgPosLin + (vgLimiteInf - (vgLimiteInf - (vgPosLinMemo0 + vgTF0))) > vgLimiteInf) Or vgCL0 > vgQdeLMemo0 Then
            vgPula0 = True
         End If
         If (vgPosLin + (vgLimiteInf - (vgLimiteInf - (vgPosLinMemo2 + vgTF2))) > vgLimiteInf) Or vgCL2 > vgQdeLMemo2 Then
            vgPula2 = True
         End If
         If vgPula0 And vgPula2 Then
            vgPula0 = False
            vgPosLinMemo0 = 0
            vgPula1 = False
            vgPosLinMemo1 = 0
            vgPula2 = False
            vgPosLinMemo2 = 0
            vgTestaPriCab = False
            GoSub VerificaFimPg
            GoSub OutraFolha
            On Error Resume Next
         End If
         If Not vgPula0 And vgCL0 <= vgQdeLMemo0 And (vgPosLinMemo0 + vgTF0) < (69.84 - vgTF0) Then
            vgFP.txtFixa(0).Text = MostraObs()
            vgTF0 = Imprime(vgFP.ExtraiLinhaMemo(0, vgCL0), .52, vgPosLin + vgPosLinMemo0, "F=Courier New|S=11.25|L=195", vgObSaida, , ,, vgFP.vgProcuraPg)
            vgPosLinMemo0 = vgPosLinMemo0 + vgTF0
            vgCL0 = vgCL0 + 1
         End If
         If Not vgPula2 And vgCL2 <= vgQdeLMemo2 Then
            vgFP.txtFixa(2).Text = MostraPclas()
            vgTF2 = Imprime(vgFP.ExtraiLinhaMemo(2, vgCL2), 2.11, vgPosLin + vgPosLinMemo2, "F=Courier New|S=9.75", vgObSaida, , ,, vgFP.vgProcuraPg)
            vgPosLinMemo2 = vgPosLinMemo2 + vgTF2
            vgCL2 = vgCL2 + 1
         End If
      Loop
      vgCL0 = 0: vgCL2 = 0
      vgPosLin = vgPosLin + Maior(vgPosLinMemo0, vgPosLinMemo2, 91.53)
      vgDifProxReg = 0
   End If
   vgPrimeiroCab = False
   vgTemFimRel = False
   Return

CabQb20:
   If vgPriQb20 < 1 Then
      Imprime "Produtos / Serviços", .52, vgPosLin, "F=Courier New|W=700|S=12|L=196|JC=196.29|BOX=F14737632/.52/" + Str$(vgPosLin) + "/196.29/4.23", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Codigo", 1.32, vgPosLin +4.49, "F=Courier New|W=700|S=11.25|L=16", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Descrição", 17.98, vgPosLin +4.76, "F=Courier New|W=700|S=11.25|L=22", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Qtde", 136.33, vgPosLin +4.76, "F=Courier New|W=700|S=11.25|L=12|JD", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Vr. Unit", 165.32, vgPosLin +4.76, "F=Courier New|W=700|S=11.25|L=23|JD", vgObSaida, , ,, vgFP.vgProcuraPg
      Imprime "Vr. Total", 196.54, vgPosLin +4.76, "F=Courier New|W=700|S=11.25|L=23|JD", vgObSaida, , ,, vgFP.vgProcuraPg
      vgPosLin = vgPosLin + 9.78
      vgDifProxReg = 0
   End If
   vgPriQb20 = 2
   Return

SalvaInicioPg:
   If (vgPV Or vgFP.vgContaPg + 1 > UBound(vgBookPg)) And (vgTemFimRel Or Not Vendas.EOF) Then 'testa se precisa mais variáveis
      If vgFP.vgContaPg + 1 > UBound(vgBookPg) Then
         ReDim Preserve vgBookPg(vgFP.vgContaPg + 1) As Variant                                'cria variável para salvar bookmark
         ReDim Preserve vgPodeVoltar(vgFP.vgContaPg + 1) As Boolean                            'controle de quebra em filhos
         Redim Preserve vgMemo0(vgFP.vgContaPg + 1) As Long
         Redim Preserve vgMemo1(vgFP.vgContaPg + 1) As Long
         Redim Preserve vgPgvgTot1(vgFP.vgContaPg + 1) As Double
         Redim Preserve vgMemo2(vgFP.vgContaPg + 1) As Long
         vgMemo0(vgFP.vgContaPg + 1) = vgCL0
         vgMemo1(vgFP.vgContaPg + 1) = vgCL1
      vgPgvgTot1(vgFP.vgContaPg + 1) = vgTot1
         vgMemo2(vgFP.vgContaPg + 1) = vgCL2
      End If
      If Not Vendas.EOF Then
         vgBookPg(vgFP.vgContaPg + 1) = Vendas.BookMark                                        'salva inicio desta página
      Else
         vgBookPg(vgFP.vgContaPg + 1) = ""
         vgPodeVoltar(vgFP.vgContaPg + 1) = True                                               'pode voltar preview para esta página?
      End If
      vgPodeVoltar(vgFP.vgContaPg + 1) = vgPagOk Or vgPrimeiroCab                              'pode voltar preview para esta página?
   End If
   If Not vgPriVez Then
      vgPagOk = False
   End If
   Return

SalvaPagina:
   If Not vgEmPapel Or (vgEmPapel And vgNumPg = 0) Then
      With vgFP
         If .vgFimDoArq = False And Vendas.EOF Then
            .vgFimDoArq = True
         End If
         Ok = ((((vgTemFimRel Or .vgFimDoArq = False) Or .vgContaPg < .vgUltPagVista) And vgModoInt = 0) Or ((.vgPagRef < UBound(vgPagRef())) And vgModoInt = 1)) And Not .vgProcuraPg
         If Ok <> .botMais.Enabled Then
            .botMais.Enabled = Ok                 'habilita/desabilita botões do preview
         End If
         Ok = .botMais.Enabled And Not .vgProcuraPg
         If Ok <> .botUltimo.Enabled Then
            .botUltimo.Enabled = Ok
         End If
         Ok = (.vgPagRef > 1 And Not .vgProcuraPg)
         If Ok <> .botVolta.Enabled Then
            .botVolta.Enabled = Ok
         End If
         Ok = (.vgPagRef > 1 And Not .vgProcuraPg)
         If Ok <> .botPrimeiro.Enabled Then
            .botPrimeiro.Enabled = Ok
         End If
         GoSub SalvaInicioPg
         If .vgProcuraPg Or (.vgContaPg <> .vgNovaPag And .vgNovaPag <> 0) Then 'procurando pg adiante...
            If .vgProcuraPg And ((vgModoInt = 0 And (Not vgTemFimRel And Vendas.EOF)) Or .vgContaPg = .vgNovaPag) Then
               .ShowPreview                       'mostra preview
            Else
               .picAux.Cls                        'limpa a área para a próxima pg
               vgPosLin = vgLimiteInf + 99        'força um cabec na próxima vez
            End If
         Else
            .ShowPreview
         End If
      End With
   End If
   Return
'retira as páginas inválidas do intervalo especificado!
ArrumaVetor:
   vgQt = 0
   j = 0
   k = 0
   For i = 1 To UBound(vgPagRef())
      If vgPagRef(i) < 0 Or (Abs(vgPagRef(i)) >= k And k > 0) And i < UBound(vgPagRef()) - vgQt + 1 Then
         If vgPagRef(i) < k And vgPagRef(i) <> 0 Then k = Abs(vgPagRef(i))
         For j = 1 To UBound(vgPagRef())
            If Abs(vgPagRef(j)) >= k And j <= UBound(vgPagRef()) - vgQt + 1 Then
               vgQt = vgQt + 1
               For m = j + 1 To UBound(vgPagRef())
                  vgPagRef(m - 1) = vgPagRef(m)
               Next
               If j <= UBound(vgPagRef()) - vgQt Then
                  j = j - 1
               End If
            End If
         Next
         If i < UBound(vgPagRef()) Then
            i = i - 1
         End If
      End If
   Next
   ReDim Preserve vgPagRef(UBound(vgPagRef()) - vgQt) As Integer
   
   If vgFP.vgPagRef > UBound(vgPagRef()) Then
      vgFP.vgPagRef = UBound(vgPagRef())
   End If
   Return

RestauraPag:
   vgFP.txtPag.Text = CStr(vgPagOrig)
   vgFP.botPara.Enabled = False
   vgFP.vgProcuraPg = False
   If vgPagOrig < vgFP.vgContaPg Then             'se não for a última página, vamos restaurar o bookmark para atualizar os botões do preview
      Vendas.BookMark = vgBookPg(vgPagOrig)
   End If
   vgFP.vgContaPg = vgPagOrig
   vgFP.vgPagRef = vgPagOrig
   Return

IniciaImpressao:
   On Error Resume Next                           'algumas impressoras não tem certas propriedades
   If vgEmPapel Then
      With Printer                                'salva configs da impressora
         vgPsm = .ScaleMode
         vgPor = .Orientation
         vgPsz = .PaperSize
         vgPfn = .FontName
         .PaperSize = vbPRPSA4
         .Orientation = vbPRORPortrait
         If Err Then Err.Clear
         If vgNumPg <> 0 Then                     'procurando pg para imprimir
            .Copies = 1
         Else
            .Copies = vgCopias                    'número de cópias configurada no diálogo
            If Err And vgNCop = 0 Then vgNCop = vgCopias
         End If
      End With
   End If
   With vgObSaida.Font
      'ajusta para nós...
      .Name = "MS Sans Serif"
      .CharSet = 0
      .Weight = 400
      .Size = 8.25
      .Underline = 0
      .Italic = 0
      .StrikeThrough = 0
      vgObSaida.ForeColor = &H0
      vgObSaida.ScaleMode = vbMillimeters         'vamos trabalhar em mm
   End With
   On Error GoTo DeuErro                          'vamos cercar erros...
   Return

NaoAchouPg:
   GoSub SalvaPagina
   vgFP.txtPag.Text = CStr(vgPagOrig)
   MsgBox "A página especificada não foi encontrada no relatório!", vbExclamation
   GoTo FimDaSub

DeuErro:                                          'deu um erro
   Select Case Err                                'vamos verificar se deu algum erro
      Case 3197                                   'registro acabou de ser alterado por outro usuário
         Resume
      Case 3046, 3158, 3186, 3187, 3188, 3189, 3218, 3260
         vgNVez = vgNVez + 1                      'outro usuário acessando
         If vgNVez <= 10 Then                     'vamos tentar 10 vezes
            Delay .5                              'mais antes espera 1/2 segundo
            Resume                                'antes de tentar
         End If
   End Select
   If Err <> 65524 Then                           'se não for um erro de crítica
      CErr.NumErro = Err
      CErr.Show
   End If
   vgPrimeiroCab = 2                              'flag para não dar a msg novamente

SaiDaSub:                                         'saída comum da sub...

FimDaSub:
   If Not vgFG Is Nothing Then
      UnLoad vgFG
      Set vgFG = Nothing
   End If
   Err.Clear
   Screen.MousePointer = vbDefault                'ponteiro do mouse = default
End Sub

'evento - configura impressora
Private Sub botConfImp_Click()
   Dim x As String
   ConfigImpressora                               'chama diáglogo de configuração impressora
   x$ = Printer.DeviceName
   Label(4).Caption = x$
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
   vgFP.vgFimDoArq = False                        'é fim do arquivo
   vgFP.vgContaPg = 0                             'contador de paginas
   vgFP.vgNovaPag = 0                             'número da página real dentro do intervalo (Ex: 3;5;1;2 - se estiver na página real 5, a página atual será 2)
   vgFP.vgPagRef = 0                              'número da página de referência
   vgFP.vgUltPagVista = 0                         'ate onde ja fomos...
   If vgDestRel = 0 Then                          'onde vai ser gerado (janela ou papel)
      vgFP.cboEscala.ListIndex = 3                'escal 100%
      ProssegueExecucao vgFP.picAux               'faz preview
   Else
      vgFP.vgContaPg = 0
      ProssegueExecucao Printer                   'joga direto na impressora
   End If
   If vgPrimeiroCab <> 2 Then                     'se não deu erro
      If vgPrimeiroCab Then                       'nenhuma página foi impressa
         MsgBox LoadGasString(92), vbExclamation, vgAtencao$ 'apresenta a mensagem
      ElseIf vgDestRel = 0 Then
         vgFP.ShowPreview                                    'arruma form preview
         vgFP.Visible = True                                 'mostra form de preview
         Visible = False                                     'esconde o este form
         vgVisivel = False                                   'e seta variavel
      End If
   End If
   vgPriVez = False
End Sub

'evento - cancela formulário
Private Sub botCancela_Click()
   Unload Me
End Sub

'evento - antes de descarregar o formulário
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = Cancel Or(Not vgFG Is Nothing)
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
      Unload vgFP.txtFixa(1)
      Unload vgFP.txtFixa(2)
   Unload vgFP.picFixa(1)
   Set txtvgDestRel = Nothing
   Set lblAjuste = Nothing
   Set txtVenda = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   If Not Vendas Is Nothing Then
      Vendas.CloseRecordSet
      Set Vendas = Nothing
   End If
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   SET vgFP.vgFormRel = Nothing                   'variável associada a este form
   UnLoad vgFP                                    'descarrega o form de preview
   Set vgFP = Nothing                             'tira o form de preview da memória
   Set frmRelVenda = Nothing                      'libera o segmento de código do form
   If vgRelAtivo = vgFormID Then vgRelAtivo = 0
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
         MySql  
   End Select
End Sub


'evento - quando o valor da rolagem modificar
Private Sub scrVtxtCampo1_Change()
   txtCampo(1).scrChange
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

