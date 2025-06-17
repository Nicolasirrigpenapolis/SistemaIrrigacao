VERSION 5.00
Begin VB.Form frmF_GNF
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   5610
   Left                     =   2445
   LinkTopic                =   "F_GNF"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   5175
   ScaleWidth               =   9000
   Top                      =   1855
   Width                    =   9090
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
      Height                   =   5010
      Left                     =   15
      Width                    =   8820
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   15549
      _extenty                 =   8832
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1440
         Left                     =   75
         Width                    =   1395
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
         Left                     =   1755
         Width                    =   3315
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
         Caption                  =   "Gerar - Nota Fiscal"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8535
         Width                    =   45
         Top                      =   5175
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
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1755
         Y1                       =   1005
         X2                       =   5070
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
         Left                     =   135
         Width                    =   1020
         Top                      =   1695
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
         Caption                  =   "Sequencia:"
         Index                    =   2 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1335
         Top                      =   1695
         Width                    =   1590
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   2655
         Top                      =   1710
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1350
         Width                    =   1290
         Top                      =   1710
         TabStop                  =  0
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3075
         Width                    =   1455
         Top                      =   1695
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
         Caption                  =   "Transportadora:"
         Index                    =   3 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4695
         Top                      =   1695
         Width                    =   3915
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   8340
         Top                      =   1710
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8085
         Top                      =   1710
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
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4710
         Width                    =   3360
         Top                      =   1710
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4005
         Width                    =   510
         Top                      =   2235
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
         Caption                  =   "Frete:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4695
         Top                      =   2235
         Width                    =   3915
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8340
         Top                      =   2250
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
         Left                     =   4710
         Width                    =   3615
         Top                      =   2250
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   615
         Width                    =   585
         Top                      =   2775
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
         Caption                  =   "CFOP:"
         Index                    =   5 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1320
         Top                      =   2775
         Width                    =   7290
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   8340
         Top                      =   2790
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8085
         Top                      =   2790
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
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1335
         Width                    =   6735
         Top                      =   2790
         Index                    =   3 
      End
      Begin NFE.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   4
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   435
         Left                     =   5400
         Width                    =   1590
         Top                      =   4455
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
      Begin NFE.GBotao botCancela
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   5
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   435
         Left                     =   7035
         Width                    =   1545
         Top                      =   4455
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
         Caption                  =   "Cancela"
      End
      Begin VB.Frame Frame
         Height                   =   960
         Left                     =   120
         Width                    =   8505
         Top                      =   3315
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
         Index                    =   0 
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   120
            Width                    =   540
            Top                      =   360
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
            Index                    =   6 
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
            MaxLength                =   254
            BorderStyle              =   1
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   6
            TabStop                  =   0
            Height                   =   300
            Left                     =   780
            Width                    =   4290
            Top                      =   360
            TabStop                  =  0
            Index                    =   4 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   5220
            Width                    =   1425
            Top                      =   360
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
            Caption                  =   "Tot. Movimento:"
            Index                    =   7 
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
            MaxLength                =   254
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   7
            TabStop                  =   0
            Height                   =   300
            Left                     =   6840
            Width                    =   1530
            Top                      =   360
            TabStop                  =  0
            Index                    =   5 
            MultiLine                =   -1
         End
      End
   End
         Begin VB.Timer timUnLoad
                  Enabled         =   0
                  Interval        =   5
                  left            =   30
                  top             =   950
         End
End
Attribute VB_name = "frmF_GNF"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: F_GNF
'* Função....: Outras Saidas - Nota Fiscal
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:44
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

Dim txtCampo(5) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim MovimentoTela As Double, Transporte As Double, Frete As String
Dim Cfop As Double
Public lblAjuste As Object, txtMovimentoTela As Object, txtTransporte As Object
Public txtFrete As Object, txtCfop As Object

Dim MvtoAux As New GRecordSet

Private Function CarregaOMesmo() As Long
 Dim Tb As New GRecordSet
  
 Set Tb = vgDb.OpenRecordSet("SELECT * From Geral Where [Razão Social] = 'O Mesmo'")
 
 If Tb.RecordCount = 1 Then
    CarregaOMesmo = Tb![Sequencia Do geral]
 End If

End Function

Public Sub RepositionSaida()
   On Error Resume Next
  
   Set MvtoAux = vgDb.OpenRecordSet("SELECT [Data de Emissão], Geral.[Sequencia do geral] SeqGeral, Geral.[Razão social] Geral, [Valor total do movimento] " & _
                                    "FROM [Outras saidas estoque] LEFT JOIN Geral ON [Outras saidas estoque].[Sequencia do geral] = Geral.[Sequencia do geral] " & _
                                    "WHERE [Sequencia da saida] = " & MovimentoTela)
   MostraFormulas
                                                                               
End Sub


Private Function Filtro() As String   
   Dim Tb As New GRecordSet
    
   Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Natureza de Operação] Where [Sequencia da natureza] > 0 And Inativo = 0")
   
   Filtro = "Cfop = 5102" 
                                                           
   If Tb.RecordCount = 0 Then Exit Function
   
   Do While Not Tb.EOF
   If Not Mid(Tb!Cfop,1, 1) = 1 And Not Mid(Tb!Cfop,1, 1) = 2 Then
      Filtro = Filtro & " Or Cfop = " & Tb!Cfop
   End If    
   Tb.MoveNext
   Loop
   
End Function


'O gas não respeita o filtro ENTAO Rotina Manual...
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "[Sequencia da saida] > " & 0 & ""
      Case 1
         txtCampo(1).Filter = "(([Sequencia do geral] > " & 0 & ") AND Inativo = " & 0 & ") AND Transportadora = " & 1 & ""
                           txtCampo(1).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.Transportadora = " + CStr(1) + ")"
      Case 3
         txtCampo(3).Filter = "" + Filtro() + ""
                           txtCampo(3).PesqSQLExpression = "SELECT [Natureza de operação].[Sequencia da natureza], [Natureza de operação].Cfop, [Natureza de operação].[Descrição do cfop] FROM [Natureza de operação] WHERE  " & Filtro() & ""
   End Select
   
End Sub


Private Sub GerarNota()
   Dim NotaFiscal As New GRecordSet, NProdutos As New GRecordSet, NParcelas As New GRecordSet
   Dim Movimento As New GRecordSet, MProdutos As New GRecordSet, MParcelas As New GRecordSet
   Dim Manutencao As New GRecordSet, NaturezaOp As New GRecordSet
   Dim UltimaEletronica As New GRecordset
   Dim GeralAux As New GRecordSet, Contribuinte As Boolean
   
   On Error GoTo DeuErro
   
   If MsgBox("Deseja Realmente Gerar Nota Fiscal desse Movimento de Saída?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   
   Screen.MousePointer = vbHourglass 'Ampulheta
   Set NotaFiscal = vgDb.OpenRecordSet("Nota fiscal")
   Set NProdutos = vgDb.OpenRecordSet("Produtos da nota fiscal")
   Set NParcelas = vgDb.OpenRecordSet("Parcelas Nota Fiscal")
   Set Movimento = vgDb.OpenRecordSet("SELECT * FROM [Outras saidas estoque] WHERE [Sequencia da saida] = " & MovimentoTela)
   Set MProdutos = vgDb.OpenRecordSet("SELECT * FROM [Produtos saidas estoque] WHERE [Sequencia da saida] = " & MovimentoTela)
   Set NaturezaOp = vgDb.OpenRecordSet("SELECT * FROM [Natureza de operação] WHERE [Sequencia da natureza] = " & Cfop)
   Set UltimaEletronica = vgdb.OpenRecordset("SELECT MAX([Numero da nfe]) NFe, MAX([Sequencia da nota fiscal]) SeqNFe FROM [Nota fiscal]")
   Set GeralAux = vgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Movimento![Sequencia Do geral])
   
   If GeralAux!Tipo = 0 Then ' Pessoa Fisica 
      Contribuinte = False
   ElseIf GeralAux!Isento Then
      Contribuinte = False
   Else
      Contribuinte = True    
   End If
   
   'Inicio Produtos
   If MProdutos.RecordCount > 0 Then
      With Movimento
         ![Sequencia da nota fiscal] = UltimaEletronica!SeqNFe + 1 
         NotaFiscal.AddNew
         'Gerando Número da NFe
         NotaFiscal![Sequencia da nota fiscal] = UltimaEletronica!SeqNFe + 1 
         NotaFiscal![Numero da nfe] = UltimaEletronica!NFe + 1         
         NotaFiscal![Tipo de Nota] = 0 'Saida
         NotaFiscal![Finalidade nfe] = 1  
         NotaFiscal!Historico = !Historico  
         NotaFiscal![Sequencia da saida] = ![Sequencia da saida]
         NotaFiscal![Sequencia da natureza] = NaturezaOp![Sequencia da natureza]
         NotaFiscal!Frete = Frete
         NotaFiscal![Sequencia da transportadora] = Transporte
         NotaFiscal![Nota avulsa] = 0
         If NaturezaOp![Cfop] = "6916" Or NaturezaOp![Cfop] = "5916" Then
            NotaFiscal!Conserto = 1
         End If
         NotaFiscal![Data de emissão] = Date
         NotaFiscal![Sequencia Do geral] = ![Sequencia Do geral]
         NotaFiscal![Data de saida] = Date
         NotaFiscal![Hora da saida] = Time
         If Not Vazio(![Forma de pagamento]) Then
            NotaFiscal![Forma de pagamento] = ![Forma de pagamento]
         Else
            NotaFiscal![Forma de pagamento] = "Prazo"    
         End If 
         NotaFiscal![Valor total Do ipi] = ![Valor total Do ipi]
         NotaFiscal![Valor total dos produtos] = ![Valor total dos produtos]
         NotaFiscal![Valor total da nfe] = ![Valor total Do movimento]
         NotaFiscal![Valor Do frete] = ![Valor Do frete]
         NotaFiscal![Valor Do desconto] = ![Valor Do desconto]
         NotaFiscal.Update
         NotaFiscal.BookMark = .LastModified
         NotaFiscal.MoveLast
         'Gerando Número da Nota Eletronica no Contas a Pagar
         Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da saida] = " & ![Sequencia da saida])
         If Manutencao.RecordCount > 0 Then
            Do While Not Manutencao.EOF
               With Manutencao
                  .Edit
                  ![Sequencia da nota fiscal] = NotaFiscal![Sequencia da nota fiscal]
                  .Update
                  .BookMark = .LastModified
               End With
               Manutencao.MoveNext
            Loop
         End If
         .Update
         .BookMark = .LastModified
      End With
      'Produtos
      Do While Not MProdutos.EOF
         With MProdutos
            NProdutos.AddNew
            NProdutos![Sequencia da Nota Fiscal] = NotaFiscal![Sequencia da Nota Fiscal]
            NProdutos![Sequencia Do Produto] = ![Sequencia Do Produto]
            NProdutos!Quantidade = !Quantidade
            NProdutos![Valor Unitario] = ![Valor Unitario]
            NProdutos![Valor Total] = ![Valor Total]
            NProdutos![Aliquota Do ipi] = ![Aliquota Do ipi]
            NProdutos![Valor Do ipi] = ![Valor Do ipi]
            NProdutos!CFOP = NaturezaOp![Cfop]
            If Parametros_da_nfe![Simples nacional] Then
             If Contribuinte Then
                NProdutos!Cst = 101
             Else
                NProdutos!Cst = 102    
             End If
            End If       
            NProdutos.Update
            NProdutos.BookMark = .LastModified
            MProdutos.MoveNext
         End With
      Loop
      NotaFiscal.Update
      NotaFiscal.BookMark = NotaFiscal.LastModified
      NotaFiscal.MoveLast
   End If
   'Fim Produtos      
   'Inicio Parcelas
   Movimento.MoveFirst
   TbAuxiliar "Parcelas outras saidas", "[Sequencia da saida] = " & Movimento![Sequencia da saida], MParcelas
   If MParcelas.RecordCount > 0 Then
      Do While Not MParcelas.EOF
         NParcelas.AddNew
         NParcelas![Sequencia da nota fiscal] = NotaFiscal![Sequencia da nota fiscal]
         NParcelas![Numero da Parcela] = MParcelas![Numero da Parcela]
         NParcelas!Dias = MParcelas!Dias
         NParcelas![Data de Vencimento] = MParcelas![Data de Vencimento]
         NParcelas![Valor da Parcela] = MParcelas![Valor da Parcela]
         NParcelas.Update
         NParcelas.BookMark = NParcelas.LastModified
         MParcelas.MoveNext
      Loop
   End If
        
   vgDb.CommitTrans
   frmNotaFisc.vgFiltroInicial = "[Sequencia da Nota Fiscal] = " & NotaFiscal![Sequencia da nota fiscal] 
   InicializaFiltro frmNotaFisc
   frmNotaFisc.Show
   frmNotaFisc.SetFocus
   If NotaFiscal![Valor Do frete] > 0 Or  NotaFiscal![Valor Do Desconto] > 0 Then
      frmNotaFisc.RateiaProdutos
   Else
      frmNotaFisc.AtualizaImposto
   End If         
  
   Unload Me
     
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   MovimentoTela = txtCampo(0).Value
   Transporte = txtCampo(1).Value
   Frete = txtCampo(2).Value
   Cfop = txtCampo(3).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
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
      vgOk = (Transporte > 0)
      vgMsg$ = "Transportadora não pode ser em Branco!"
      If Not vgOk Then txtCampo(1).SetFocus
      If vgOk Then
         vgOk = (Not Vazio(Frete))
         vgMsg$ = "Frete não pode ser em Branco!!"
         If Not vgOk Then txtCampo(2).SetFocus
      End If
      If vgOk Then
         vgOk = (Cfop > 0)
         vgMsg$ = "CFOP é Obrigatorio!"
         If Not vgOk Then txtCampo(3).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
            txtCampo(0).Value = frmMoSaidas.vgtb![Sequencia da saida]
      txtCampo(1).Value = CarregaOMesmo()
      txtCampo(2).Value = "Destinatario"
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

'mostra fórmulas no formulário
Private Sub MostraFormulas()
   On Error Resume Next                           'previne erros...
   txtCampo(4).Value = MvtoAux!SeqGeral & " - " &   MvtoAux!Geral
   If Err Then Err = 0: txtCampo(4).Text = ""
   txtCampo(5).Value = MvtoAux![Valor total do movimento]
   If Err Then Err = 0: txtCampo(5).Text = ""
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Then
      MostraFormulas
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 1
            PoeRelEFiltroCbo 1
            If Len(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = CarregaOMesmo()
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               MostraFormulas
            End If
         Case 2
            If Len(txtCp(2).Text) = 0 Then
               txtCampo(2).Value = "Destinatario"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               MostraFormulas
            End If
         Case 3
            PoeRelEFiltroCbo 3
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
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(25680)
   vgFormID = 671
   vgIdentTab = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(17185))
   Set botOk.Picture = LoadPicture(LoadGasPicture(17186))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(17187))
   Set lblAjuste = Label(1)
   Set txtMovimentoTela = txtCampo(0)
   Set txtTransporte = txtCampo(1)
   Set txtFrete = txtCampo(2)
   Set txtCfop = txtCampo(3)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(25685)
   Label(1).Caption = LoadGasString(25686)
   Label(2).Caption = LoadGasString(25687)
   Label(3).Caption = LoadGasString(25688)
   Label(4).Caption = LoadGasString(25689)
   Label(5).Caption = LoadGasString(25690)
   vgTooltips.AddTool botOk, 0, LoadGasString(25691)
   botOk.Caption = LoadGasString(25692)
   vgTooltips.AddTool botCancela, 0, LoadGasString(25693)
   botCancela.Caption = LoadGasString(25694)
   Label(6).Caption = LoadGasString(25695)
   Label(7).Caption = LoadGasString(25696)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   MostraFormulas
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
   Set txtCampo(0).CtBot(BOT_COMBO) = bottxtCampo0(BOT_COMBO)
   Set bottxtCampo0(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = ""
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = "Sequencia da saida"
   txtCampo(0).ListFields = "Sequencia da saida"
   txtCampo(0).OrderFields = "Sequencia da saida"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Outras saidas estoque"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_LISTA) = bottxtCampo1(BOT_LISTA)
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   bottxtCampo1(BOT_LISTA).Caption = "P"
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = ""
   txtCampo(1).PesqModoAbertura =  2 
   txtCampo(1).BoundColumn = "Sequencia do geral"
   txtCampo(1).ListFields = "Razão social"
   txtCampo(1).OrderFields = "Razão social"
   txtCampo(1).Relation = ""
   txtCampo(1).Source = "Geral"
   txtCampo(1).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 0
   txtCampo(2).ListFields = "Emitente|Destinatario"

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_LISTA) = bottxtCampo3(BOT_LISTA)
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   bottxtCampo3(BOT_LISTA).Caption = "P"
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = ""
   txtCampo(3).PesqModoAbertura =  2 
   txtCampo(3).PesqFieldCapture = "[Natureza de operação].[Sequencia da natureza]"
   txtCampo(3).BoundColumn = "Sequencia da natureza"
   txtCampo(3).ListFields = "Cfop; Descrição do cfop"
   txtCampo(3).OrderFields = "Cfop; Descrição do cfop"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Natureza de operação"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = ""
   txtCampo(4).Editable = False
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "999.999.999,99"
   txtCampo(5).Editable = False
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

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
   GerarNota  
   vgPriVez = False
End Sub

'evento - cancela formulário
Private Sub botCancela_Click()
   Unload Me
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
   Set txtMovimentoTela = Nothing
   Set txtTransporte = Nothing
   Set txtFrete = Nothing
   Set txtCfop = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmF_GNF = Nothing                         'libera o segmento de código do form
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
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

