VERSION 5.00
Begin VB.Form frmConPedid
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   9105
   Left                     =   165
   LinkTopic                =   "ConPedid"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8670
   ScaleWidth               =   18090
   Top                      =   1255
   Width                    =   18180
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
      Height                   =   8490
      Left                     =   15
      Width                    =   17910
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   31575
      _extenty                 =   14967
      Index                    =   0 
      Begin IRRIG.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   -1
         Height                   =   1935
         Left                     =   135
         Width                    =   1650
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
         Left                     =   1875
         Width                    =   5505
         Top                      =   195
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
         Caption                  =   "Controle de Pedidos de Compra"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1890
         Width                    =   510
         Top                      =   1020
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
         Caption                  =   "Inicio:"
         Index                    =   1 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2445
         Top                      =   1020
         Width                    =   1515
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   3690
         Top                      =   1035
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
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   270
         Left                     =   2460
         Width                    =   1215
         Top                      =   1035
         Index                    =   0 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2445
         Top                      =   1620
         Width                    =   1515
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   3690
         Top                      =   1635
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
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   270
         Left                     =   2460
         Width                    =   1215
         Top                      =   1635
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1995
         Width                    =   375
         Top                      =   1620
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
         Caption                  =   "Fim:"
         Index                    =   2 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1920
         Y1                       =   570
         X2                       =   7395
         Y2                       =   570
         Index                    =   0 
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
         TabIndex                 =   10
         TabStop                  =   -1
         Stripes                  =   -1
         SaveGridBars             =   1
         Height                   =   6135
         Left                     =   195
         Width                    =   17355
         Top                      =   2100
         Caption                  =   "Pedidos"
         ShowGridCaption          =   -1
         Index                    =   0 
         Tag                      =   "1"
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   5055
         Top                      =   1020
         Width                    =   1050
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   5835
         Top                      =   1035
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabIndex                 =   3
         TabStop                  =   0
         Height                   =   270
         Left                     =   5070
         Width                    =   750
         Top                      =   1035
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5115
         Width                    =   990
         Top                      =   765
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
         Caption                  =   "Orçamento"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   12135
         Width                    =   1185
         Top                      =   765
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
         Caption                  =   "Fornecedor"
         Index                    =   4 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   12135
         Top                      =   1020
         Width                    =   5325
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   17190
         Top                      =   1035
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabIndex                 =   6
         TabStop                  =   -1
         Height                   =   270
         Left                     =   12150
         Width                    =   5025
         Top                      =   1035
         Index                    =   3 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   4035
         Top                      =   1020
         Width                    =   990
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4755
         Top                      =   1035
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         Left                     =   4050
         Width                    =   690
         Top                      =   1035
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4035
         Width                    =   660
         Top                      =   765
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
         Caption                  =   "Pedido"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6195
         Width                    =   795
         Top                      =   765
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
         Caption                  =   "Cod. Mat"
         Index                    =   6 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   6195
         Top                      =   1020
         Width                    =   1110
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7035
         Top                      =   1035
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6210
         Width                    =   810
         Top                      =   1035
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7455
         Width                    =   930
         Top                      =   765
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
         Caption                  =   "Descrição"
         Index                    =   7 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabStop                  =   -1
         Height                   =   300
         Left                     =   7455
         Width                    =   4590
         Top                      =   1020
         Index                    =   6 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   4050
         Top                      =   1635
         Width                    =   2070
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   5850
         Top                      =   1650
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4065
         Width                    =   1770
         Top                      =   1650
         Index                    =   7 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4035
         Width                    =   750
         Top                      =   1395
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
         Caption                  =   "F. Pagto"
         Index                    =   8 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         MaxLength                =   30
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6195
         Width                    =   1170
         Top                      =   1635
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6195
         Width                    =   1110
         Top                      =   1395
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
         Caption                  =   "Comprador"
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7455
         Width                    =   1110
         Top                      =   1395
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
         Caption                  =   "Cliente"
         Index                    =   10 
      End
      Begin VB.Label labFdo9
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFFF
         Left                     =   7455
         Top                      =   1635
         Width                    =   4590
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo9
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11775
         Top                      =   1650
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFFF
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
         TabIndex                 =   9
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7470
         Width                    =   4290
         Top                      =   1650
         Index                    =   9 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   11
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   16020
         Width                    =   1410
         Top                      =   120
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
         Caption                  =   "Imprimir"
         Index                    =   0 
      End
   End
End
Attribute VB_name = "frmConPedid"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: ConPedid
'* Função....: Controle de Pedidos de Compra
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 14:59:43
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
Public vgFiltraSoUm As String                     'flag para filtrar só um registro 
Public vgCancelou As Integer                      'flag cancela ou nao o processo do módulo  
Public vgQueryAtual As String                     'expressão SQL definada para o módulo 

Public vgFP As New frmPreview                     'formuário de preview do relatório 
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFG As frmGauge                           'formulário do gauge 
Public vgQReg As Long                             'quantidade de já feira para o gauge  

Dim txtCampo(9) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim Dt_Inicial As Variant, Dt_Final As Variant, Orcamento As Double
Dim Codigo_do_Fornecedor As Long, PedidoTela As Double, Produto As Double
Dim Nome As String, F_pagto As String, Comprador As String
Dim Cliente As String
Public grdContas As Object, Txtorc As Object

Dim TemRecordset As Boolean
Dim PriVez As Boolean

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
      .AddColumn Nothing, , "Pedido", "Id do Pedido", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 920
      .AddColumn Nothing, , "Dt. Emissão", "Data do Pedido", TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      '.AddColumn Nothing, , "Comprador", "Comprador", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1500
      .AddColumn Nothing, , "Orçamento", "Seqüência do Orçamento", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 920
      .AddColumn Nothing, , "Cliente", "Cliente", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 3100
      .AddColumn Nothing, , "Cod. Mat", "Seqüência do Produto", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 920
      .AddColumn Nothing, , "Descrição", "Descrição", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 4140
       .AddColumn Nothing, , "Fornecedor", "Razão Social", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 4140
      .AddColumn Nothing, , "Qtde Total", "Qtde_Total", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "R$ Unit. + IPI", "Vr_Unit_Ipi", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Qtde Recebida", "Qtde_Recebida", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Qtde Restante", "Qtde_Restante", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1000
      .AddColumn Nothing, , "Prazo", "Prazo", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1500
      .AddColumn Nothing, , "Financeiro", "Financeiro", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Dt. Baixa", "Data da Baixa", TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 920
      .AddColumn Nothing, , "Previsão", "Prev Entrega", TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      
      .ShowSumCol(.Columns("Qtde").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   'CarregaContas 
   PriVez = False 
End Sub


Private Sub CarregaContas()
 Dim x As String, Tb As New GRecordSet, MsgValida As String
 Dim Pedido As New GRecordSet
 Dim Geral As New GRecordSet
 Dim Bx As New GRecordSet
 Dim DtBx As New GRecordSet
 
 Screen.MousePointer = vbHourglass 'Ampulheta
 
 If Privez = 0 Then
 Set Pedido = vgDb.OpenRecordSet("SELECT [Pedido de Compra Novo].[Id do Pedido], [Codigo do Fornecedor], [Sequencia Do Item], Prazo, Qtde, [Data do Pedido], Comprador, [Seqüência do Orçamento], [Id do Produto], ([Vr Unitario] + ([Vr do IPI] / Qtde)) As [Vr Unitario], ([Vr Total] + [Vr do IPI]) As [Vr Total], [Previsao de Entrega], Dias From [Pedido de Compra Novo] LEFT JOIN [Produtos do Pedido Compra] On [Pedido de Compra Novo].[Id do Pedido] = [Produtos do Pedido Compra].[Id do Pedido] WHERE " & Filtro())
 Set Tb = vgDb.OpenRecordSet("Controle de Compras")
 
 vgDb.Execute ("DELETE FROM [Controle de Compras]")
 
   If Pedido.RecordCount > 0 Then
      Do While Not Pedido.EOF
         With Tb
           .AddNew
              ![Id Do Pedido] = Pedido![Id Do Pedido]
              ![Sequencia Do Item] = Pedido![Sequencia Do Item]
              ![Data Do Pedido] = Pedido![Data Do Pedido]
              !Comprador = Pedido!Comprador
              ![Seqüência Do Orçamento] = Pedido![Seqüência Do Orçamento]
               Set Geral = vgDb.OpenRecordSet("SELECT Geral.[Seqüência Do Geral], [Razão Social] From Orçamento Left Join Geral On Orçamento.[Seqüência Do Geral] = Geral.[Seqüência Do Geral] WHERE [Seqüência Do Orçamento] = " & ![Seqüência Do Orçamento])
              ![Seqüência Do Geral] = Geral![Seqüência Do Geral]
              ![Razão Social] = Geral![Razão Social]
              ![Seqüência Do Produto] = Pedido![Id Do Produto]
              ![Vr_Unit_Ipi] = Pedido![Vr Unitario]
              !Prazo = Pedido!Prazo
              !Financeiro = MostraFinanceiro(Pedido![Id Do Pedido])
               If !Financeiro <> "A Pagar" And Not Vazio(!Financeiro) Then
                   Set DtBx = vgDb.OpenRecordSet("SELECT Max([Data da Baixa]) Baixa From [Manutenção Contas] WHERE [Sequencia da Compra] = " & ![Id Do Pedido])
                  ![Data da Baixa] = DtBx!Baixa
               End If
              !Dias = Pedido!Dias
              ![Prev Entrega] = Pedido![Previsao de Entrega]
              ![Codigo Do Fornecedor] = Pedido![Codigo Do Fornecedor]
              ![Qtde_Total] = Pedido!Qtde
               Set Bx = vgDb.OpenRecordSet("SELECT [Qtde Recebida], [Qtde Restante] FROM [Bx Produtos Pedido Compra] WHERE [Id do Pedido] = " & Pedido![Id Do Pedido] & " And [Sequencia do Item] = " & Pedido![Sequencia Do Item] & " And [Id do Produto] = " & Pedido![Id Do Produto])
                   If Bx.RecordCount > 0 Then
                     ![Qtde_Recebida] = Bx![Qtde Recebida]
                     ![Qtde_Restante] = Bx![Qtde Restante]
                   End If
           .Update
           .BookMark = .LastModified
         End With
      Pedido.MoveNext
      Loop
   End If
   End If
   
   PriVez = True
       
   x = ("SELECT [Id do Pedido], Prazo, Financeiro, [Data do Pedido], [Data da Baixa], Comprador, [Seqüência do Orçamento], [Codigo do Fornecedor], [Controle de Compras].[Razão Social] As Cliente, Geral.[Razão Social], Produtos.[Seqüência do Produto], Produtos.Descrição, [Qtde_Total], [Vr_Unit_Ipi], [Qtde_Recebida], [Qtde_Restante], Dias, [Prev Entrega] From [Controle de Compras] LEFT JOIN Geral On [Controle de Compras].[Codigo do Fornecedor] = Geral.[Seqüência do Geral] LEFT JOIN Produtos On [Controle de Compras].[Seqüência do Produto] = Produtos.[Seqüência do Produto] Where " & Filtro2())
   
   Set Tb = vgDb.OpenRecordSet(x$)
   
   If Tb.RecordCount > 0 Then TemRecordset = True
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

  Screen.MousePointer = vbDefault 'Ampulheta

End Sub


Private Function MostraFinanceiro(Pedido As Long) As String
 Dim Tb1 As New GRecordSet
 
 Set Tb1 = vgDb.OpenRecordSet("SELECT SUM([Valor Pago]) Pago, Sum([Valor Restante]) Restante From [Manutenção Contas] WHERE [Sequencia da Compra] = " & Pedido & " And [Sequencia da Compra] > 0")
 
  If Tb1!Pago > 0 And Tb1!Restante > 0 Then
     MostraFinanceiro = "Pago Parcial"
  ElseIf Tb1!Pago > 0 And Tb1!Restante = 0 Then  
     MostraFinanceiro = "Liquidado"
  ElseIf Tb1!Pago = 0 And Tb1!Restante > 0 Then      
    MostraFinanceiro = "A Pagar"
  Else MostraFinanceiro = ""  
  End If
    
End Function

Private Function Filtro() As String

   Filtro = "[Pedido de Compra Novo].[Id do Pedido] > 0 And [Pedido Fechado] = 0 And [Id do Produto] > 0"
   
   Filtro = Filtro & " AND [Pedido de Compra Novo].Cancelado = 0  And Prepedido = 0"
   If Not Vazio(Dt_Inicial) Then
       Filtro = Filtro & " AND [Pedido de Compra Novo].[Data do Pedido] >= " & D(Dt_Inicial)
   End If
   If Not Vazio(Dt_Final) Then
       Filtro = Filtro & " AND [Pedido de Compra Novo].[Data do Pedido] <= " & D(Dt_Final)
   End If
 
   Filtro = Filtro & " Order By [Pedido de Compra Novo].[Id do Pedido] "   
   
End Function


Private Function Filtro2() As String
  Filtro2 = "[Id do Pedido] > 0"
  If PedidoTela > 0 Then
     Filtro2 = Filtro2  & " And [Id do Pedido] = " & PedidoTela
  End If
  If Orcamento > 0 Then
     Filtro2 = Filtro2 & " AND [Seqüência do Orçamento] = " & Orcamento
  End If
  If Codigo_do_Fornecedor > 0 Then
       Filtro2 = Filtro2 & " AND [Codigo do Fornecedor] = " & Codigo_do_Fornecedor
  End If
  If Produto > 0 Then
      Filtro2 = Filtro2 & " AND [Seqüência do Produto] = " & Produto
  End If
  If Not Vazio(Nome)  Then
     Filtro2 = Filtro2 & " AND Produtos.Descrição Like '%" & Nome & "%'"
  End If
   If Not Vazio(F_pagto) Then
      Filtro2 = Filtro2 & " AND Prazo = '" & F_pagto & "'"
   End If
   If Not Vazio(Comprador) Then
      Filtro2 = Filtro2 & " AND Comprador = '" & Comprador & "'"
   End If
   If Not Vazio(Cliente) Then
      Filtro2 = Filtro2 & " AND [Controle de Compras].[Razão Social] Like '%" & Cliente & "%'"
   End If
   Filtro2 = Filtro2 & "Order By [Id do Pedido]"
End Function   

Private Sub AbreRel()

   frmRelContr.Show
   frmRelContr.txtDt_Inicial.Value = Dt_Inicial 
   frmRelContr.txtDt_Final.Value = Dt_Final
   frmRelContr.txtComprador.Value = Comprador
   frmRelContr.txtPedidoTela.Value = PedidoTela 
   frmRelContr.txtorc.Value = Orcamento
   frmRelContr.txtCliente.Value = Cliente 
   
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Dt_Inicial = txtCampo(0).Value
   Dt_Final = txtCampo(1).Value
   Orcamento = txtCampo(2).Value
   Codigo_do_Fornecedor = txtCampo(3).Value
   PedidoTela = txtCampo(4).Value
   Produto = txtCampo(5).Value
   Nome = txtCampo(6).Value
   F_pagto = txtCampo(7).Value
   Comprador = txtCampo(8).Value
   Cliente = txtCampo(9).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 2
         txtCampo(2).Filter = "(([Seqüência do Orçamento] > " & 0 & ") AND [Venda Fechada] = " & 1 & ") AND Cancelado = " & 0 & ""
      Case 3
         txtCampo(3).Filter = "Fornecedor = " & 1 & ""
      Case 4
         txtCampo(4).Filter = "(([Id do Pedido] > " & 0 & ") AND Cancelado = " & 0 & ") AND Prepedido = " & 0 & ""
      Case 5
         txtCampo(5).Filter = "([Seqüência do Produto] > " & 0 & ") AND Inativo = " & 0 & ""
      Case 9
         txtCampo(9).Filter = "(([Seqüência do Geral] > " & 0 & ") AND Cliente = " & 1 & ") AND Inativo = " & 0 & ""
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
      vgOk = (Not Vazio(Dt_Inicial))
      vgMsg$ = "Data Inicial não pode ser vazia!"
      If Not vgOk Then txtCampo(0).SetFocus
      If vgOk Then
         vgOk = (Not Vazio(Dt_Final))
         vgMsg$ = "Data Final deve ser maior ou igual a Data Inicial!"
         If Not vgOk Then txtCampo(1).SetFocus
      End If
      If vgOk Then
         If True Then
            vgOk = (Orcamento >= 0)
            vgMsg$ = "Informe o Orçamento!"
            If Not vgOk Then txtCampo(2).SetFocus
         End If
      End If
      If vgOk Then
         vgOk = (Codigo_do_Fornecedor >= 0)
         vgMsg$ = "Codigo do Fornecedor inválido!"
         If Not vgOk Then txtCampo(3).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
            txtCampo(0).Value = CDate("01/01/2021")
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
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
   Ok = (True)
   txtCampo(2).Enabled = (Ok And txtCampo(2).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Or Index = 7 Or Index = 9 Then
      ExecutaPreValidacao
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 0
            If Len(txtCp(0).Text) = 0 Then
               txtCampo(0).Value = CDate("01/01/2021")
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 1
            If Len(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), -1))
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 2
            PoeRelEFiltroCbo 2
         Case 3
            PoeRelEFiltroCbo 3
         Case 4
            PoeRelEFiltroCbo 4
         Case 5
            PoeRelEFiltroCbo 5
         Case 9
            PoeRelEFiltroCbo 9
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
   ExecutaPreValidacao                            'habilita/desabilita campos
   Select Case Index
      Case 1
         CarregaContas  
      Case 2
         CarregaContas  
      Case 3
         CarregaContas  
      Case 4
         CarregaContas  
      Case 5
         CarregaContas  
      Case 6
         CarregaContas  
      Case 7
         CarregaContas  
      Case 8
         CarregaContas  
      Case 9
         CarregaContas  
   End Select
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(43910)
   vgFormID = 1379
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(35415))
   Set grdContas = Grid(0)
   Set Txtorc = txtCampo(2)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(43915)
   Label(1).Caption = LoadGasString(43916)
   Label(2).Caption = LoadGasString(43917)
   Label(3).Caption = LoadGasString(43918)
   Label(4).Caption = LoadGasString(43919)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(43920)
   Label(5).Caption = LoadGasString(43921)
   Label(6).Caption = LoadGasString(43922)
   Label(7).Caption = LoadGasString(43923)
   Label(8).Caption = LoadGasString(43924)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(43925)
   Label(9).Caption = LoadGasString(43926)
   Label(10).Caption = LoadGasString(43927)
   Botao(0).Caption = LoadGasString(43928)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaPreValidacao
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

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = ""
   txtCampo(4).BoundColumn = "Id do Pedido"
   txtCampo(4).ListFields = "Id do Pedido"
   txtCampo(4).OrderFields = "Id do Pedido-"
   txtCampo(4).Relation = ""
   txtCampo(4).Source = "Pedido de Compra Novo"
   txtCampo(4).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999999"
   txtCampo(2).BoundColumn = "Seqüência do Orçamento"
   txtCampo(2).ListFields = "Seqüência do Orçamento"
   txtCampo(2).OrderFields = "Seqüência do Orçamento-"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Orçamento"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = ""
   txtCampo(5).BoundColumn = "Seqüência do Produto"
   txtCampo(5).ListFields = "Seqüência do Produto; Descrição"
   txtCampo(5).OrderFields = "Seqüência do Produto; Descrição"
   txtCampo(5).Relation = ""
   txtCampo(5).Source = "Produtos"
   txtCampo(5).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = ""
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "999999"
   txtCampo(3).BoundColumn = "Seqüência do Geral"
   txtCampo(3).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(3).OrderFields = "Razão Social"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Geral"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_COMBO) = bottxtCampo7(BOT_COMBO)
   Set bottxtCampo7(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(7).DataType = 0
   txtCampo(7).ListFields = "Antecipado|Prazo"

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "@x"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""

   Set txtCampo(9).CtPri = txtCp(9)
   Set txtCampo(9).CtFdo = labFdo9
   Set txtCampo(9).CtBot(BOT_COMBO) = bottxtCampo9(BOT_COMBO)
   Set bottxtCampo9(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = ""
   txtCampo(9).BoundColumn = "Razão Social"
   txtCampo(9).ListFields = "Razão Social"
   txtCampo(9).OrderFields = "Razão Social"
   txtCampo(9).Relation = ""
   txtCampo(9).Source = "Geral"
   txtCampo(9).vgfrmGMCale.grdListaG.AutoRebind = True

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
   Set Txtorc = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmConPedid = Nothing                      'libera o segmento de código do form
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
         AbreRel  
   End Select
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
Private Sub bottxtCampo4_Click(Index As Integer)
   txtCampo(4).SetFocus
   DoEvents
   txtCampo(4).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo7_Click(Index As Integer)
   txtCampo(7).SetFocus
   DoEvents
   txtCampo(7).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo9_Click(Index As Integer)
   txtCampo(9).SetFocus
   DoEvents
   txtCampo(9).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

