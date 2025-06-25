VERSION 5.00
Begin VB.Form frmF_GeraNF 
   ClientHeight    =   5205
   ClientLeft      =   5430
   ClientTop       =   2880
   ClientWidth     =   8865
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "F_GeraNF"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   8865
   Begin VB.PictureBox Painel 
      Height          =   5190
      Index           =   0
      Left            =   15
      ScaleHeight     =   5130
      ScaleWidth      =   8775
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   15
      Width           =   8835
      Begin VB.PictureBox mmCampo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   1440
         Index           =   0
         Left            =   75
         ScaleHeight     =   1440
         ScaleWidth      =   1395
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   1395
      End
      Begin VB.CommandButton bottxtCampo0 
         Height          =   270
         Index           =   2
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1710
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1710
         Width           =   1290
      End
      Begin VB.CommandButton bottxtCampo1 
         Height          =   270
         Index           =   1
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1710
         Width           =   255
      End
      Begin VB.CommandButton bottxtCampo1 
         Height          =   270
         Index           =   2
         Left            =   8085
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1710
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   4710
         TabIndex        =   1
         Top             =   1710
         Width           =   3360
      End
      Begin VB.CommandButton bottxtCampo2 
         Height          =   270
         Index           =   2
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2250
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   4710
         TabIndex        =   2
         Top             =   2250
         Width           =   3615
      End
      Begin VB.CommandButton bottxtCampo3 
         Height          =   270
         Index           =   1
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2790
         Width           =   255
      End
      Begin VB.CommandButton bottxtCampo3 
         Height          =   270
         Index           =   2
         Left            =   8085
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2790
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1335
         TabIndex        =   3
         Top             =   2790
         Width           =   6735
      End
      Begin VB.PictureBox botOk 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5400
         ScaleHeight     =   375
         ScaleWidth      =   1530
         TabIndex        =   4
         Top             =   4455
         Width           =   1590
      End
      Begin VB.PictureBox botCancela 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7035
         ScaleHeight     =   375
         ScaleWidth      =   1485
         TabIndex        =   5
         Top             =   4455
         Width           =   1545
      End
      Begin VB.Frame Frame 
         Height          =   960
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   3315
         Width           =   8505
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   4
            Left            =   780
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   360
            Width           =   4290
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   6840
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Geral:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tot. Devolução:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   5220
            TabIndex        =   28
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Gerar - Nota Fiscal"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   0
         Left            =   1755
         TabIndex        =   10
         Top             =   555
         Width           =   3315
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   8535
         TabIndex        =   11
         Top             =   5175
         Width           =   45
      End
      Begin VB.Line Linha 
         Index           =   0
         X1              =   1755
         X2              =   5070
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Sequencia:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   1695
         Width           =   1020
      End
      Begin VB.Label labFdo0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1335
         TabIndex        =   13
         Top             =   1695
         Width           =   1590
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   3075
         TabIndex        =   15
         Top             =   1695
         Width           =   1455
      End
      Begin VB.Label labFdo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4695
         TabIndex        =   16
         Top             =   1695
         Width           =   3915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Frete:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   4005
         TabIndex        =   19
         Top             =   2235
         Width           =   510
      End
      Begin VB.Label labFdo2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4695
         TabIndex        =   20
         Top             =   2235
         Width           =   3915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "CFOP:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   615
         TabIndex        =   22
         Top             =   2775
         Width           =   585
      End
      Begin VB.Label labFdo3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1320
         TabIndex        =   23
         Top             =   2775
         Width           =   7290
      End
   End
   Begin VB.Timer timUnLoad 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   30
      Top             =   950
   End
End
Attribute VB_Name = "frmF_GeraNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: F_GeraNF
'* Função....: Devolução - Nota Fiscal
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:43
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
Dim DevolucaoTela As Double, Transporte As Double, Frete As String
Dim Cfop As Double
Public lblAjuste As Object, txtDevolucaoTela As Object, txtTransporte As Object
Public txtFrete As Object, txtCfop As Object

Dim DevolucaoAux As New GRecordSet

Public Sub RepositionDevolucao()
   On Error Resume Next
  
   Set DevolucaoAux = vgDb.OpenRecordSet("SELECT [Data da devolução], Geral.[Sequencia do geral] SeqGeral, Geral.[Razão social] Geral, [Valor total da devolução] " & _
                                         "FROM [Devolução de compra] LEFT JOIN Geral ON [Devolução de compra].[Sequencia do geral] = Geral.[Sequencia do geral] " & _
                                         "WHERE [Sequencia da devolução] = " & DevolucaoTela)
   MostraFormulas
                                                                               
End Sub


Private Function CarregaOMesmo() As Long
 Dim Tb As New GRecordSet
  
 Set Tb = vgDb.OpenRecordSet("SELECT * From Geral Where [Razão Social] = 'O Mesmo'")
 
 If Tb.RecordCount = 1 Then
    CarregaOMesmo = Tb![Sequencia Do geral]
 End If

End Function


Private Function DefineCFOP() As Long
 Dim Tb As New GRecordSet, Mun As New GRecordSet
 Dim ForadoEstado As Boolean, Cfop As New GRecordSet
 Dim Devolucao As New GRecordSet
  
  Set Devolucao = vgDb.OpenRecordSet("SELECT * From [Devolução de compra] Where [Sequencia da devolução] = " & DevolucaoTela)
  Set Tb = vgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Devolucao![Sequencia Do geral])
  Set Mun = vgDb.OpenRecordSet("SELECT * From [Municipios] Where [Sequencia do Municipio] = " & Tb![Sequencia Do municipio])
  ForadoEstado = Mun!Uf <> "SP"
  
  If ForadoEstado Then
     Set Cfop = vgDb.OpenRecordSet("SELECT * From [Natureza de operação] Where [Cfop] = 6202")
  Else
     Set Cfop = vgDb.OpenRecordSet("SELECT * From [Natureza de operação] Where [Cfop] = 5202")
  End If
  
  DefineCFOP = Cfop![Sequencia da natureza]
  
End Function


Private Sub GerarNota()
   Dim NotaFiscal As New GRecordSet, NProdutos As New GRecordSet, NParcelas As New GRecordSet
   Dim Devolucao As New GRecordSet, DProdutos As New GRecordSet, DParcelas As New GRecordSet
   Dim Manutencao As New GRecordSet, NaturezaOp As New GRecordSet
   Dim UltimaEletronica As New GRecordSet
   Dim GeralAux As New GRecordSet, Contribuinte As Boolean
   
   On Error GoTo DeuErro
   
   If MsgBox("Deseja Realmente Gerar Nota Fiscal dessa Devolução?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   
   Screen.MousePointer = vbHourglass 'Ampulheta
   Set NotaFiscal = vgDb.OpenRecordSet("Nota fiscal")
   Set NProdutos = vgDb.OpenRecordSet("Produtos da nota fiscal")
   Set NParcelas = vgDb.OpenRecordSet("Parcelas Nota Fiscal")
   Set Devolucao = vgDb.OpenRecordSet("SELECT * FROM [Devolução de compra] WHERE [Sequencia da devolução] = " & DevolucaoTela)
   Set DProdutos = vgDb.OpenRecordSet("SELECT * FROM [Produtos devolução compra] WHERE [Sequencia da devolução] = " & DevolucaoTela)
   Set NaturezaOp = vgDb.OpenRecordSet("SELECT * FROM [Natureza de operação] WHERE [Sequencia da natureza] = " & Cfop)
   Set UltimaEletronica = vgDb.OpenRecordSet("SELECT MAX([Numero da nfe]) NFe, MAX([Sequencia da nota fiscal]) SeqNFe FROM [Nota fiscal]")
   Set GeralAux = vgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Devolucao![Sequencia Do geral])
   
   If GeralAux!Tipo = 0 Then ' Pessoa Fisica
      Contribuinte = False
   ElseIf GeralAux!Isento Then
      Contribuinte = False
   Else
      Contribuinte = True
   End If
   
   'Inicio Produtos
   If DProdutos.RecordCount > 0 Then
      With Devolucao
         ![Sequencia da nota fiscal] = UltimaEletronica!SeqNFe + 1
         NotaFiscal.AddNew
         'Gerando Número da NFe
         NotaFiscal![Sequencia da nota fiscal] = UltimaEletronica!SeqNFe + 1
         NotaFiscal![Numero da nfe] = UltimaEletronica!NFe + 1
         NotaFiscal![Tipo de Nota] = 0 ' Saida
         NotaFiscal![Finalidade nfe] = 4
         NotaFiscal![Sequencia da devolução] = ![Sequencia da devolução]
         NotaFiscal![Sequencia da natureza] = NaturezaOp![Sequencia da natureza]
         NotaFiscal!Frete = Frete
         NotaFiscal![Sequencia da transportadora] = Transporte
         NotaFiscal![Nota avulsa] = 1
         NotaFiscal![Data de emissão] = Date
         NotaFiscal![Sequencia Do geral] = ![Sequencia Do geral]
         NotaFiscal!Historico = "Devolução de Compra"
         NotaFiscal![Data de saida] = Date
         NotaFiscal![Hora da saida] = Time
         If Not Vazio(![Forma de pagamento]) Then
            NotaFiscal![Forma de pagamento] = ![Forma de pagamento]
         Else
            NotaFiscal![Forma de pagamento] = "Prazo"
         End If
         NotaFiscal![Valor total Do ipi] = ![Valor total Do ipi]
         NotaFiscal![Valor total dos produtos] = ![Valor total dos produtos]
         NotaFiscal![Valor total da nfe] = ![Valor total da devolução]
         NotaFiscal![Valor Do frete] = ![Valor Do frete]
         NotaFiscal![Valor Do desconto] = ![Valor Do desconto]
         NotaFiscal.Update
         NotaFiscal.BookMark = .LastModified
         NotaFiscal.MoveLast
         'Gerando Número da Nota Eletronica no Contas aReceber
         Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da devolução] = " & ![Sequencia da devolução])
         If Manutencao.RecordCount > 0 Then
            Do While Not Manutencao.EOF
               With Manutencao
                  .Edit
                  ![Sequencia da nota fiscal] = NotaFiscal![Sequencia da nota fiscal]
                  ![Numero da nfe] = NotaFiscal![Numero da nfe]
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
      Do While Not DProdutos.EOF
         With DProdutos
            NProdutos.AddNew
            NProdutos![Sequencia da nota fiscal] = NotaFiscal![Sequencia da nota fiscal]
            NProdutos![Sequencia Do Produto] = ![Sequencia Do Produto]
            NProdutos!Quantidade = !Quantidade
            NProdutos![Valor Unitario] = ![Valor Unitario]
            NProdutos![Valor Total] = ![Valor Total]
            NProdutos![Aliquota Do ipi] = ![Aliquota Do ipi]
            NProdutos![Valor Do ipi] = ![Valor Do ipi]
            NProdutos!Cfop = NaturezaOp![Cfop]
            If Parametros_da_nfe![Simples nacional] Then
             If Contribuinte Then
                NProdutos!Cst = 101
             Else
                NProdutos!Cst = 102
             End If
            End If
            NProdutos.Update
            NProdutos.BookMark = .LastModified
            DProdutos.MoveNext
         End With
      Loop
      NotaFiscal.Update
      NotaFiscal.BookMark = NotaFiscal.LastModified
      NotaFiscal.MoveLast
   End If
   'Fim Produtos
   'Inicio Parcelas
   Devolucao.MoveFirst
   TbAuxiliar "Parcelas dev compra", "[Sequencia da devolução] = " & Devolucao![Sequencia da devolução], DParcelas
   If DParcelas.RecordCount > 0 Then
      Do While Not DParcelas.EOF
         NParcelas.AddNew
         NParcelas![Sequencia da nota fiscal] = NotaFiscal![Sequencia da nota fiscal]
         NParcelas![Numero da Parcela] = DParcelas![Numero da Parcela]
         NParcelas!Dias = DParcelas!Dias
         NParcelas![Data de Vencimento] = DParcelas![Data de Vencimento]
         NParcelas![Valor da Parcela] = DParcelas![Valor da Parcela]
         NParcelas.Update
         NParcelas.BookMark = NParcelas.LastModified
         DParcelas.MoveNext
      Loop
   End If
        
   vgDb.CommitTrans
   frmNotaFisc.vgFiltroInicial = "[Sequencia da Nota Fiscal] = " & NotaFiscal![Sequencia da nota fiscal]
   InicializaFiltro frmNotaFisc
   frmNotaFisc.Show
   frmNotaFisc.SetFocus
   
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
   DevolucaoTela = txtCampo(0).Value
   Transporte = txtCampo(1).Value
   Frete = txtCampo(2).Value
   Cfop = txtCampo(3).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "(([Sequencia da devolução] > " & 0 & ") AND Cancelado = " & 0 & ") AND Fechado = " & 1 & ""
      Case 1
         txtCampo(1).Filter = "(([Sequencia do geral] > " & 0 & ") AND Inativo = " & 0 & ") AND Transportadora = " & 1 & ""
                           txtCampo(1).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.Transportadora = " + CStr(1) + ")"
      Case 3
         txtCampo(3).Filter = "(((Cfop = " & 5201 & ") OR Cfop = " & 6201 & ") OR Cfop = " & 5202 & ") OR Cfop = " & 6202 & ""
                           txtCampo(3).PesqSQLExpression = "SELECT [Natureza de operação].[Sequencia da natureza], [Natureza de operação].Cfop, [Natureza de operação].[Descrição do cfop] FROM [Natureza de operação] WHERE ([Natureza de operação].Cfop = " + CStr(5201) + " OR " + _
                                                              "[Natureza de operação].Cfop = " + CStr(6201) + " OR [Natureza de operação].Cfop = " + CStr(5202) + " OR " + _
                                                              "[Natureza de operação].Cfop = " + CStr(6202) + ")"
   End Select
End Sub

'executa processos/validacoes nos campos do formuário
Public Function Executar(vgOq As String) As String
   Dim vgMsg As String, vgOk As Integer, vgNVez As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
       Dim freteCod As Integer
      freteCod = Val(txtCampo(2).Value)              'garante codigo numerico (0,1,3,4,9)
      vgOk = (Transporte > 0) Or (freteCod = 3) Or (freteCod = 4) Or (freteCod = 9)
      vgMsg$ = "Transportadora no pode ser em Branco!"
      If Not vgOk Then txtCampo(1).SetFocus
      If vgOk Then
         vgOk = (Not Vazio(Frete))
         vgMsg$ = "Frete no pode ser em Branco!!"
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
            txtCampo(0).Value = frmDevCompr.vgtb![Sequencia da devolução]
      txtCampo(1).Value = CarregaOMesmo()
      txtCampo(2).Value = "Destinatario"
      txtCampo(3).Value = DefineCFOP()
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
   Executar = Err.Source + "|" + Trim$(Str$(Err)) + "|" + Error$ 'não teve jeito o erro não pode ser evitado...
End Function

'mostra fórmulas no formulário
Private Sub MostraFormulas()
   On Error Resume Next                           'previne erros...
   txtCampo(4).Value = DevolucaoAux!SeqGeral & " - " & DevolucaoAux!Geral
   If Err Then Err = 0: txtCampo(4).Text = ""
   txtCampo(5).Value = DevolucaoAux![Valor total da devolução]
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
            If Len(txtCp(3).Text) = 0 Then
               txtCampo(3).Value = DefineCFOP()
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               MostraFormulas
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
   MostraFormulas                                 'mostra formulas na janela
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(25140)
   vgFormID = 664
   vgIdentTab = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(16645))
   Set botOk.Picture = LoadPicture(LoadGasPicture(16646))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(16647))
   Set lblAjuste = Label(1)
   Set txtDevolucaoTela = txtCampo(0)
   Set txtTransporte = txtCampo(1)
   Set txtFrete = txtCampo(2)
   Set txtCfop = txtCampo(3)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(25145)
   Label(1).Caption = LoadGasString(25146)
   Label(2).Caption = LoadGasString(25147)
   Label(3).Caption = LoadGasString(25148)
   Label(4).Caption = LoadGasString(25149)
   Label(5).Caption = LoadGasString(25150)
   vgTooltips.AddTool botOk, 0, LoadGasString(25151)
   botOk.Caption = LoadGasString(25152)
   vgTooltips.AddTool botCancela, 0, LoadGasString(25153)
   botCancela.Caption = LoadGasString(25154)
   Label(6).Caption = LoadGasString(25155)
   Label(7).Caption = LoadGasString(25156)
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
   txtCampo(0).BoundColumn = "Sequencia da devolução"
   txtCampo(0).ListFields = "Sequencia da devolução"
   txtCampo(0).OrderFields = "Sequencia da devolução-"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Devolução de compra"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_LISTA) = bottxtCampo1(BOT_LISTA)
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   bottxtCampo1(BOT_LISTA).Caption = "P"
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = ""
   txtCampo(1).PesqModoAbertura = 2
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
   txtCampo(2).ListFields = "Emitente|Destinatrio|Transporte Prprio Remetente|Transporte Prprio Destinatrio|Sem Transporte"


   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_LISTA) = bottxtCampo3(BOT_LISTA)
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   bottxtCampo3(BOT_LISTA).Caption = "P"
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = ""
   txtCampo(3).PesqModoAbertura = 2
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
  CErr.Origem = CStr(vgFormID) + " - " + Me.Caption
 CErr.Show
End Sub

'evento - valida digitações e prossegue a execução
Private Sub botOk_Click()
   Dim k As String
   If vgPriVez Then Exit Sub
   vgPriVez = True
   k$ = Executar(VALIDACOES)                      'faz validaçao dos campos
   If Len(k$) > 0 Then                            'EPA!... tem campo errado
      If InStr(k$, "|") > 0 Then
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
   Cancel = Cancel Or (Not vgFG Is Nothing)
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set lblAjuste = Nothing
   Set txtDevolucaoTela = Nothing
   Set txtTransporte = Nothing
   Set txtFrete = Nothing
   Set txtCfop = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmF_GeraNF = Nothing                      'libera o segmento de código do form
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
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtCampo(Index).MouseDown
End Sub

