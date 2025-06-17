Version 5.00
Begin VB.Form frmManuRece
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   4200
   Left                     =   2415
   LinkTopic                =   "ManuRece"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   3765
   ScaleWidth               =   12780
   Top                      =   2185
   Width                    =   12870
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
      Height                   =   3600
      Left                     =   15
      Width                    =   12600
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   22213
      _extenty                 =   6346
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1530
         Left                     =   60
         Width                    =   1545
         Top                      =   60
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
         Left                     =   1680
         Width                    =   5100
         Top                      =   600
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   15.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Manutenção - Contas á Receber"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1680
         Y1                       =   990
         X2                       =   6765
         Y2                       =   990
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   12180
         Width                    =   45
         Top                      =   3780
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
         Height                   =   240
         Left                     =   180
         Width                    =   840
         Top                      =   1800
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
         Caption                  =   "Conta......:"
         Index                    =   2 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   300
         Left                     =   1140
         Width                    =   825
         Top                      =   1800
         DataField                =   "Sequencia da Manutenção"
         TabStop                  =  0
         Index                    =   0 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2160
         Width                    =   1035
         Top                      =   1800
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
         Caption                  =   "Dt. Entrada:"
         Index                    =   3 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   1
         TabStop                  =   0
         Height                   =   300
         Left                     =   3360
         Width                    =   1245
         Top                      =   1800
         DataField                =   "Data de entrada"
         TabStop                  =  0
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4800
         Width                    =   540
         Top                      =   1800
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
         Caption                  =   "Geral:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HE0E0E0
         Left                     =   5580
         Top                      =   1800
         Width                    =   6765
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   12075
         Top                      =   1815
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   2
         TabStop                  =   0
         Height                   =   270
         Left                     =   5595
         Width                    =   6465
         Top                      =   1815
         DataField                =   "Sequencia do geral"
         TabStop                  =  0
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   285
         Width                    =   735
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
         Caption                  =   "Doc.......:"
         Index                    =   5 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         MaxLength                =   25
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   3
         TabStop                  =   0
         Height                   =   300
         Left                     =   1140
         Width                    =   3465
         Top                      =   2400
         DataField                =   "Documento"
         TabStop                  =  0
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4830
         Width                    =   510
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
         Caption                  =   "Pc.....:"
         Index                    =   6 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   300
         Left                     =   5580
         Width                    =   480
         Top                      =   2400
         DataField                =   "Parcela"
         TabStop                  =  0
         Index                    =   4 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6240
         Width                    =   1110
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
         Caption                  =   "Vencimento:"
         Index                    =   7 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7500
         Top                      =   2400
         Width                    =   1425
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   8655
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
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7515
         Width                    =   1125
         Top                      =   2415
         DataField                =   "Data de vencimento"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   9120
         Width                    =   1605
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
         Caption                  =   "Valor da Parcela.:"
         Index                    =   8 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   6
         TabStop                  =   0
         Height                   =   300
         Left                     =   10875
         Width                    =   1470
         Top                      =   2400
         DataField                =   "Valor da parcela"
         TabStop                  =  0
         Index                    =   6 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   90
         Width                    =   930
         Top                      =   3000
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
         Caption                  =   "Cobrança:"
         Index                    =   9 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1140
         Top                      =   3000
         Width                    =   3465
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4335
         Top                      =   3015
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
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1155
         Width                    =   3165
         Top                      =   3015
         DataField                =   "Sequencia da Cobrança"
         Index                    =   7 
      End
   End
         Begin NFE.GListV grdBrowse 
                  Height          =   795
                  Left            =   60
                  TabIndex        =   0
                  Top             =   330
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   1402
                  FullRowSelect   =   0           'False
                  RowHeight       =   225
                  AllowEdit       =   -1          'True
                  AllowInsert     =   -1          'True
                  AllowDelete     =   -1          'True
                  ManualUpdate    =   -1          'True
                  ManualDelete    =   -1          'True
                  NavigationAddMode=   1
                  ShowGridCaption =   0           'False
                  ShowFilterBar   =   -1          'True
                  Caption         =   ""
                  CacheSize       =   100
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0  'False
                           Italic          =   0  'False
                           StrikeThrough   =   0  'False
                  EndProperty
         End
End
Attribute VB_name = "frmManuRece"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: ManuRece
'* Função....: Manutenção - Receber
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:57
'* * * * * * *

Option Explicit
DefInt A-Z

Public vgSituacao As Integer                      'situação de edição que do módulo
Public vgCaracteristica As Integer                'caracteristica do módulo
Public vgTipo As Integer                          'tipo do módulo
Public vgFiltroInicial As String                  'filtro inicial do módulo   
Public vgOrdemInicial As String                   'ordem inicial do módulo
Public vgUltimaOrdem As String                    'última ordenação feita no módulo    
Public vgUltimoFiltro As String                   'último filtro definido no módulo   
Public vgUltimoFiltroComTit As String             'titulo do último filtro definido no módulo
Public vgUltimaOrdemComTit As String              'titulo da última ordenação feita no módulo 
Public vgUltimoTabIndex As Integer                'último campo com foco do módulo
Public vgPriVez As Integer                        'flag de carregamento do módulo  
Public WithEvents vgTb As GRecordSet              'tabela de dados do módulo
Public vgSQL As String                            'expressão SQL que define o módulo
Public vgTemInclusao As Integer                   'flag se tem ou não inclusão no módulo 
Public vgTemExclusao As Integer                   'flag se tem ou não exclusão no módulo 
Public vgTemProcura As Integer                    'flag se tem ou não procura no módulo 
Public vgTemFiltro As Integer                     'flag se tem ou não filtro no módulo
Public vgTemAlteracao As Integer                  'flag se tem ou não alteração no módulo 
Public vgTemAlteracaoGrids As Integer              'flag se tem ou não alteração nos grids
Public vgTemBrowse As Integer                     'flag se tem ou não janela em grade no módulo
Public vgSemVincDados As Integer                  'Flag para definir formulários sem vinculo com dados
Public vgEmBrowse As Integer                      'flag se o módulo esta em grade 
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso 
Public vgIndDefault As String                     'indice default do módulo 
Public vgFormID As Long                           'identificador único para o módulo 
Public vgIdentTab As String                       'nome da tabela principal do módulo 
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFiltroOriginal As String
Dim txtCampo(7) As New FormataCampos              'classe dos campos tipo texto do módulo  
Dim Numero_da_nfe As Long, Sequencia_da_manutencao As Long, Parcela As Integer
Dim Sequencia_do_geral As Long, Data_de_entrada As Variant, Historico As String
Dim Especie_de_pagamento As String, Forma_de_pagamento As String, Data_de_vencimento As Variant
Dim Valor_total As Double, Valor_da_parcela As Double, Valor_pago As Double
Dim Valor_do_juros As Double, Valor_do_desconto As Double, Valor_restante As Double
Dim Data_da_baixa As Variant, Sequencia_da_entrada As Long, Sequencia_da_baixa As Long
Dim Sequencia_da_nota_fiscal As Long, Documento As String, Sequencia_da_devolucao As Long
Dim Sequencia_da_saida As Long, Sequencia_da_Venda As Long, Sequencia_da_Ordem As Long
Dim Sequencia_da_Cobranca As Integer, Tipo_de_conta As String, Teste As String
Public lblAjuste As Object

Private Sub ChamaAdm()
Dim Tb As New GRecordSet

 On Error Resume Next
 
 Set Tb = vgDb.OpenRecordSet("Select * From [Manutenção Contas Receber] Where [Sequencia da Manutenção] = " & Sequencia_da_manutencao)
 
 If Tb![Data de Vencimento] <> Data_de_vencimento Then
 
  InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       MDINFE.SalvaDados  
       MsgBox "Vencimento Alterado com Sucesso!", vbCritical + vbOKOnly, vaTitulo 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       MDINFE.CancelaAlteracoes 
    End If
  End If
 End If 
 
  If Vazio(InputYgorsBom) Then
     MDINFE.CancelaAlteracoes 
  End If
  
InputYgorsBom = ""  
 
End Sub


Private Sub ChamaAdm2()
 On Error Resume Next
  InputSenha
  
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       MDINFE.SalvaDados  
       MsgBox "Cobrança Alterada com Sucesso!", vbCritical + vbOKOnly, vaTitulo 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       MDINFE.CancelaAlteracoes 
    End If
  End If
   
  If Vazio(InputYgorsBom) Then
     MDINFE.CancelaAlteracoes 
  End If
  
InputYgorsBom = ""  
 
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Numero_da_nfe = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Numero da nfe])
         Sequencia_da_manutencao = txtCampo(0).Value
         Parcela = txtCampo(4).Value
         Sequencia_do_geral = txtCampo(2).Value
         Data_de_entrada = txtCampo(1).Value
         Historico = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb!Historico)
         Especie_de_pagamento = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Especie de pagamento])
         Forma_de_pagamento = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Forma de pagamento])
         Data_de_vencimento = txtCampo(5).Value
         Valor_total = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Valor total])
         Valor_da_parcela = txtCampo(6).Value
         Valor_pago = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Valor pago])
         Valor_do_juros = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Valor do juros])
         Valor_do_desconto = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Valor do desconto])
         Valor_restante = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Valor restante])
         Data_da_baixa = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da baixa])
         Sequencia_da_entrada = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da entrada])
         Sequencia_da_baixa = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da baixa])
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Documento = txtCampo(3).Value
         Sequencia_da_devolucao = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da devolução])
         Sequencia_da_saida = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da saida])
         Sequencia_da_Venda = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da Venda])
         Sequencia_da_Ordem = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da Ordem])
         Sequencia_da_Cobranca = txtCampo(7).Value
         Tipo_de_conta = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Tipo de conta])
         Teste = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb!Teste)
      Else
         Numero_da_nfe = IIf(IsNull(vgTb![Numero da nfe]), 0, vgTb![Numero da nfe])
         Sequencia_da_manutencao = IIf(IsNull(vgTb![Sequencia da Manutenção]), 0, vgTb![Sequencia da Manutenção])
         Parcela = IIf(IsNull(vgTb!Parcela), 0, vgTb!Parcela)
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Data_de_entrada = vgTb![Data de entrada]
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Especie_de_pagamento = IIf(IsNull(vgTb![Especie de pagamento]), "", vgTb![Especie de pagamento])
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Data_de_vencimento = vgTb![Data de vencimento]
         Valor_total = IIf(IsNull(vgTb![Valor total]), 0, vgTb![Valor total])
         Valor_da_parcela = IIf(IsNull(vgTb![Valor da parcela]), 0, vgTb![Valor da parcela])
         Valor_pago = IIf(IsNull(vgTb![Valor pago]), 0, vgTb![Valor pago])
         Valor_do_juros = IIf(IsNull(vgTb![Valor do juros]), 0, vgTb![Valor do juros])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Valor_restante = IIf(IsNull(vgTb![Valor restante]), 0, vgTb![Valor restante])
         Data_da_baixa = vgTb![Data da baixa]
         Sequencia_da_entrada = IIf(IsNull(vgTb![Sequencia da entrada]), 0, vgTb![Sequencia da entrada])
         Sequencia_da_baixa = IIf(IsNull(vgTb![Sequencia da baixa]), 0, vgTb![Sequencia da baixa])
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Documento = IIf(IsNull(vgTb!Documento), "", vgTb!Documento)
         Sequencia_da_devolucao = IIf(IsNull(vgTb![Sequencia da devolução]), 0, vgTb![Sequencia da devolução])
         Sequencia_da_saida = IIf(IsNull(vgTb![Sequencia da saida]), 0, vgTb![Sequencia da saida])
         Sequencia_da_Venda = IIf(IsNull(vgTb![Sequencia da Venda]), 0, vgTb![Sequencia da Venda])
         Sequencia_da_Ordem = IIf(IsNull(vgTb![Sequencia da Ordem]), 0, vgTb![Sequencia da Ordem])
         Sequencia_da_Cobranca = IIf(IsNull(vgTb![Sequencia da Cobrança]), 0, vgTb![Sequencia da Cobrança])
         Tipo_de_conta = IIf(IsNull(vgTb![Tipo de conta]), "", vgTb![Tipo de conta])
         Teste = IIf(IsNull(vgTb!Teste), "", vgTb!Teste)
      End If
   End If
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'verifica permissões para as condições especiais
'e muda situação de alguns botões
Private Sub AnalisaCondicoes()
   Dim i As Integer
   On Error Resume Next
   If Not mdiNFE.ActiveForm Is Nothing Then
      If mdiNFE.ActiveForm.Name <> Me.Name And Me.Visible Then Exit Sub
   End If
   With mdiNFE
      i = (False) and Permitido(vgIdentTab, ACAO_INCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemInclusao = i
      grdBrowse.AllowInsert = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = (False) and Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i
      LigaDesligaControles Me, Not i
   End With
End Sub

'executa processos/validacoes nos campos do arquivo
Public Function Executar(vgOq As String, Optional ByRef vgColumn As Integer) As String
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
      vgMsg$ = "Data de vencimento não pode ser vazio!"
      If Not vgOk Then vgColumn = 6
      If vgOk Then
         vgMsg$ = ""
      ElseIf vgColumn <> -1  And Not vgEmBrowse  Then
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
                  On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 7
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 7
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         If vgSituacao = ACAO_NAVEGANDO Then
            If Me.Name = mdiNFE.ActiveForm.Name Then
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
                        (IsNull(txtCampo(i).Value) Xor IsNull(vgTb.Fields(txtCampo(i).DataField).Value)) Then    'se for diferente do conteúdo atual do RS
                  vgTb.Fields(txtCampo(i).DataField).Value = txtCampo(i).Value
               End If
            End If
         End If
      Next
      Numero_da_nfe = IIf(IsNull(vgTb![Numero da nfe]), 0, vgTb![Numero da nfe])
      Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
      Especie_de_pagamento = IIf(IsNull(vgTb![Especie de pagamento]), "", vgTb![Especie de pagamento])
      Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
      Valor_total = IIf(IsNull(vgTb![Valor total]), 0, vgTb![Valor total])
      Valor_pago = IIf(IsNull(vgTb![Valor pago]), 0, vgTb![Valor pago])
      Valor_do_juros = IIf(IsNull(vgTb![Valor do juros]), 0, vgTb![Valor do juros])
      Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
      Valor_restante = IIf(IsNull(vgTb![Valor restante]), 0, vgTb![Valor restante])
      Data_da_baixa = vgTb![Data da baixa]
      Sequencia_da_entrada = IIf(IsNull(vgTb![Sequencia da entrada]), 0, vgTb![Sequencia da entrada])
      Sequencia_da_baixa = IIf(IsNull(vgTb![Sequencia da baixa]), 0, vgTb![Sequencia da baixa])
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      Sequencia_da_devolucao = IIf(IsNull(vgTb![Sequencia da devolução]), 0, vgTb![Sequencia da devolução])
      Sequencia_da_saida = IIf(IsNull(vgTb![Sequencia da saida]), 0, vgTb![Sequencia da saida])
      Sequencia_da_Venda = IIf(IsNull(vgTb![Sequencia da Venda]), 0, vgTb![Sequencia da Venda])
      Sequencia_da_Ordem = IIf(IsNull(vgTb![Sequencia da Ordem]), 0, vgTb![Sequencia da Ordem])
      Tipo_de_conta = IIf(IsNull(vgTb![Tipo de conta]), "", vgTb![Tipo de conta])
      Teste = IIf(IsNull(vgTb!Teste), "", vgTb!Teste)
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
   End If
   Executar = vgMsg$                              'prepara saida da função
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
   Executar = Err.Source + "|" + Trim$(str$(Err)) + "|" + Error$ 'não teve jeito o erro não pode ser evitado...
   If Err = 3265 Then Executar = Executar & vbCrLf & vbCrLf & txtCampo(i).DataField
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      Set vgRsError = Nothing
   End If
   vgPriVez = vgPV
End Function

Private Sub grdBrowse_DeleteData(ByVal vgItem As Long, vgColumns() As Variant, vgDataDeleted As Boolean, vgErrorMessage As String)
   vgDataDeleted = mdiNFE.ExcluiRegistro()
End Sub
   
Private Sub grdBrowse_InitEdit(CancelEdit As Boolean)
   Reposition
End Sub

Private Sub grdBrowse_ItemSelect(ByVal vgItem As Long, vgColumns() As Variant)
   If vgPriVez Or Not grdBrowse.Visible Then Exit Sub
   AjustaRolagem Me
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

   


Private Sub grdBrowse_SaveData(ByVal vgItem As Long, vgColumns() As Variant, vgDataSaved As Boolean, vgColumn As Integer, vgErrorMessage As String)
   mdiNFE.SalvaDados vgColumn
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
End Sub
   
Private Sub grdBrowse_StatusChanged(ByVal vgNewStatus As Integer)
   If (vgNewStatus = ACAO_EXCLUINDO And Val(grdBrowse.Recordset.BookMark) >= 0) Then
      Reposition
   End If
   PrepBotoes Me, vgNewStatus                          'acerta icones dos botoes
   mdiNFE.RemontaForm                                  'remonta dos os form da tela
End Sub

'apresenta popup menu para trabalhar com o grid
Private Sub grdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal vgCurCol As Integer)
   MostraPopGrid Me, Button
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim Ok As Boolean
   If Not Me.ActiveControl Is Nothing Then
      Ok = (Not TypeOf Me.ActiveControl Is GListV)         'se não está em um GRID
   Else
      Ok = True
   End If
   If Not Ok Then
      Ok = (Me.ActiveControl.Status = ACAO_NAVEGANDO And Not Me.ActiveControl.PreEditing) 'e se grid não está em pré-edição, edição nem inclusão
   End If
   If KeyAscii = vbKeyEscape And Ok Then                                                  'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'evento - quando o formuário for pintado
Private Sub Form_Paint()
   grdBrowse.Visible = vgEmBrowse                 'AH VB!!...
End Sub

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   vgTemAlteracaoGrids = Not LockGrids
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   Select Case Index
      Case 7
         ChamaAdm2  
   End Select
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 7
            PoeRelEFiltroCbo 7
      End Select
   End If
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
   txtCampo(Index).LostFocus
   Select Case Index
      Case 5
         ChamaAdm  
   End Select
End Sub

'evento - quando o formulário receber o foco
Private Sub Form_Activate()
   If vgPriVez = False Then
      Screen.MousePointer = vbHourglass           'mouse = ampulheta
   Else
      vgPriVez = False
   End If
   AtivaForm Me
   
   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   If FormEstaAberto("frmEnviaEmail") Then
      If Not frmEnviaEmail.Visible Then
         UnLoad vgFrmImpCons
         Set vgFrmImpCons = Nothing
         UnLoad frmEnviaEmail
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(35560)
   vgFormID = 697
   vgIdentTab = "Manutenção contas receber"
   vgFiltroEmUso = -1
   vgIndDefault = "Seq_ManuReceber"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 7
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(27065))
   Set lblAjuste = Label(1)
   Set vgTb = New GRecordSet
   If Len(vgFiltroOriginal$) = 0 And Len(vgFiltroInicial$) > 0 Then
      vgFiltroOriginal$ = vgFiltroInicial$
   End If
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(35565)
   Label(1).Caption = LoadGasString(35566)
   Label(2).Caption = LoadGasString(35567)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(35568)
   Label(3).Caption = LoadGasString(35569)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(35570)
   Label(4).Caption = LoadGasString(35571)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(35572)
   Label(5).Caption = LoadGasString(35573)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(35574)
   Label(6).Caption = LoadGasString(35575)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(35576)
   Label(7).Caption = LoadGasString(35577)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(35578)
   Label(8).Caption = LoadGasString(35579)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(35580)
   Label(9).Caption = LoadGasString(35581)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(35582)
   AjustaTamanho Me
   IniciaFormDados Me
   vgPriVez = False
   Reposition
   Screen.MousePointer = vbDefault
   Exit Sub

DeuErro:
   CErr.NumErro = Err
   CErr.FunctionName = "IniciaForm"
   CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
   CErr.Show
End Sub

Public Sub DefineControles()
 On Error GoTo DeuErro
 grdBrowse.AddControlIgnoreFocus mdiNFE.botCancela.hwnd           'não deixa o grid tentar gravar automaticamente
 grdBrowse.AddControlIgnoreFocus mdiNFE.botSalva.hwnd             'se estiver perdendo o foco para esses botões
   grdBrowse.AllowDelete = True
   grdBrowse.AllowEdit = True
   grdBrowse.SpecialPopupDisabled POP_GRID_BARS Or POP_GRID_STRIPES

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "9999999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 2
   txtCampo(1).Mask = "99/99/9999"
   txtCampo(1).Editable = False
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999999"
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = "Sequencia do geral"
   txtCampo(2).ListFields = "Razão social"
   txtCampo(2).OrderFields = "Razão social"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Geral"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@!"
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "999"
   txtCampo(4).Editable = False
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_ACAO) = bottxtCampo5(BOT_ACAO)
   Set bottxtCampo5(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(5).DataType = 2
   txtCampo(5).Mask = "99/99/9999"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = "9.999.999,99"
   txtCampo(6).Editable = False
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_COMBO) = bottxtCampo7(BOT_COMBO)
   Set bottxtCampo7(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "999"
   txtCampo(7).BoundColumn = "Sequencia da Cobrança"
   txtCampo(7).ListFields = "Descrição da Cobrança"
   txtCampo(7).OrderFields = "Descrição da Cobrança"
   txtCampo(7).Relation = ""
   txtCampo(7).Source = "Tipo da Cobrança"
   txtCampo(7).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub


'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 2
         txtCampo(2).Filter = "[Sequencia do geral] > " & 0 & ""
      Case 7
         txtCampo(7).Filter = "[Sequencia da Cobrança] > " & 0 & ""
   End Select
End Sub

'evento - antes de descarregar o formulário
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And vgBotoesOK Then  'botoeira esta correta?
      AtivaForm Me                                      'entao coloca
   End If
   Cancel = FormPendente(Me)                            've se tem atualizacao pendente
End Sub

'evento - redefinido o tamanho do formulário
Private Sub Form_Resize()
   AjustaPanFundo Me
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

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   grdBrowse.Finalize                             'descarrega o grid
   Set frmManuRece = Nothing                      'libera o segmento de código do form
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
Private Sub bottxtCampo7_Click(Index As Integer)
   txtCampo(7).SetFocus
   DoEvents
   txtCampo(7).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


