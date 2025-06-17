VERSION 5.00
Begin VB.Form frmtransfer
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   5475
   Left                     =   5070
   LinkTopic                =   "transfer"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   5040
   ScaleWidth               =   7530
   Top                      =   2545
   Width                    =   7620
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
      Height                   =   4860
      Left                     =   15
      Width                    =   7350
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   12958
      _extenty                 =   8568
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   255
         Width                    =   1335
         Top                      =   1425
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
         Caption                  =   "Banco Origem:"
         Index                    =   0 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   270
         Top                      =   1695
         Width                    =   4530
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   4530
         Top                      =   1710
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4275
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
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   270
         Left                     =   285
         Width                    =   3975
         Top                      =   1710
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4875
         Width                    =   480
         Top                      =   1425
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
         Caption                  =   "*C.C.:"
         Index                    =   1 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4950
         Top                      =   1695
         Width                    =   1905
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6585
         Top                      =   1710
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6330
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
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4965
         Width                    =   1350
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
         Left                     =   6870
         Width                    =   45
         Top                      =   4410
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
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   255
         Width                    =   1365
         Top                      =   2115
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
         Caption                  =   "Banco Destino:"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1650
         Width                    =   705
         Top                      =   690
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
         Caption                  =   "*Data.:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1650
         Top                      =   960
         Width                    =   1485
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   2865
         Top                      =   975
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
         MaxLength                =   10
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1665
         Width                    =   1185
         Top                      =   975
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3180
         Width                    =   705
         Top                      =   705
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
         Caption                  =   "*Valor:"
         Index                    =   5 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3240
         Top                      =   960
         Width                    =   1575
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4545
         Top                      =   975
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
         MaxLength                =   12
         BorderStyle              =   0
         Alignment                =   1
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   270
         Left                     =   3255
         Width                    =   1275
         Top                      =   975
         Index                    =   3 
         MultiLine                =   -1
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
         Height                   =   1350
         Left                     =   60
         Width                    =   1335
         Top                      =   0
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   375
         Left                     =   1575
         Width                    =   4335
         Top                      =   180
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   15.75
            Underline                =   -1
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Transferência Entre Contas"
         Index                    =   6 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   255
         Top                      =   2385
         Width                    =   4545
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   4530
         Top                      =   2400
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4275
         Top                      =   2400
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
         Left                     =   270
         Width                    =   3990
         Top                      =   2400
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4875
         Width                    =   480
         Top                      =   2115
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
         Caption                  =   "*C.C.:"
         Index                    =   7 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4950
         Top                      =   2385
         Width                    =   1905
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6585
         Top                      =   2400
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6330
         Top                      =   2400
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
         Left                     =   4965
         Width                    =   1350
         Top                      =   2400
         Index                    =   5 
      End
      Begin IRRIG.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   6
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   360
         Left                     =   6155
         Width                    =   360
         Top                      =   4140
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
      Begin IRRIG.GBotao botCancela
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   7
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   360
         Left                     =   6515
         Width                    =   360
         Top                      =   4140
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
         Alignment                =   0
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   1035
         Left                     =   255
         Width                    =   6660
         Top                      =   2955
         Index                    =   6 
         MultiLine                =   -1
         ScrollBars               =   2
         MaxLength                =   0
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
         Top                      =   2715
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
         Caption                  =   "Histórico:"
         Index                    =   8 
      End
   End
         Begin VB.Timer timUnLoad
                  Enabled         =   0
                  Interval        =   5
                  left            =   30
                  top             =   950
         End
End
Attribute VB_name = "frmtransfer"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: transfer
'* Função....: Transferência Entre Contas
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 15:02:18
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

Dim txtCampo(6) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Sequencia_da_Agencia As Integer, Sequencia_da_CC_da_Agencia As Integer, Data_do_Movimento As Variant
Dim Valor_Total As Double, Sequencia_da_Agencia1 As Integer, Sequencia_da_CC_da_Agencia1 As Integer
Dim Historico As String
Public txtSequencia_da_CC_da_Agencia As Object, lblAjuste As Object, txtSequencia_da_CC_da_Agencia1 As Object
Public txtHistorico As Object

Private Function Filtro() As String   
   Dim Tb As New GRecordSet, TB2 As New GRecordSet
   
   Filtro = "[Seqüência da Agência] > 0"
   
   Set Tb = vgDb.OpenRecordSet("SELECT Agências.[Seqüência da Agência] " & _
                               "FROM Agências INNER JOIN [Conta Corrente da Agência] ON Agências.[Seqüência da Agência] = [Conta Corrente da Agência].[Seqüência da Agência] " & _
                               "WHERE Inativo = 1 AND Agências.[Seqüência da Agência] > 0")
                               
   If Tb.RecordCount = 0 Then Exit Function
   
   Do While Not Tb.EOF
      Set TB2 = vgDb.OpenRecordSet("SELECT Agências.[Seqüência da Agência] " & _
                                   "FROM Agências INNER JOIN [Conta Corrente da Agência] ON Agências.[Seqüência da Agência] = [Conta Corrente da Agência].[Seqüência da Agência] " & _
                                   "WHERE Inativo = 0 AND [Seqüência da CC da Agência] > 0 AND Agências.[Seqüência da Agência] = " & Tb![Seqüência da Agência])
      If TB2.RecordCount = 0 Then
         Filtro = Filtro & " AND [Seqüência da Agência] <> " & Tb![Seqüência da Agência]
      End If
      Tb.MoveNext
   Loop

End Function


Private Sub Lancar()
 Dim Mvto As New GRecordSet
 Dim Manutencao As New GRecordSet
 Dim Sequencia_da_Movimentacao_CC As Long
 Dim Baixa As New GRecordset 
  
 ' If Data_do_Movimento > (Date + 3) Or Data_do_Movimento < (Date + 3) Then
 '     MsgBox("Data do Movimento Invalida!")
 '     Exit Sub
 ' End If
  
  If MsgBox("ATENÇÃO! " & vgPWUsuario & " " & vbCrLf & "Você tem Ciência que a Origem é a Conta de onde sai o Dinheiro?", vbExclamation + vbYesNo, vaTitulo) = vbNo Then Exit Sub
  If MsgBox("ATENÇÃO! " & vgPWUsuario & " " & vbCrLf & "Você tem Ciência que o Destino é a Conta que vai Receber o Dinheiro$?", vbExclamation + vbYesNo, vaTitulo) = vbNo Then Exit Sub
  If MsgBox("Posso Prosseguir", vbExclamation + vbYesNo, "ATENÇÃO!") = vbNo Then Exit Sub
   
  If Sequencia_da_Agencia = Sequencia_da_Agencia1 Then
     If Sequencia_da_CC_da_Agencia = Sequencia_da_CC_da_Agencia1 Then 
        MsgBox("Qui Lokuraaaa... O Banco de Origem e de Destino não podem ser o mesmo Impossivel Prosseguir!")
     End If   
     Exit Sub
  End If
  
  Set Manutencao = vgDb.OpenRecordSet("Manutenção Contas")
  Set Mvto = vgDb.OpenRecordSet("Movimentação da Conta Corrente")
  
  With Mvto
   .AddNew             
         ![Seqüência da Agência] = Sequencia_da_Agencia 'Origem
          If Sequencia_da_Agencia = 1 Then
             ![Codigo Do Historico] = 249 
          Else
             ![Codigo Do Historico] = 250       
          End If
         ![Seqüência da CC da Agência] = Sequencia_da_CC_da_Agencia
         ![Seqüência Grupo Despesa] = 34
         ![Seqüência SubGrupo Despesa] = 246
         ![Data Do Movimento] = Data_do_Movimento 
         ![Valor Total] = Valor_Total 
         ![Data Do Último Dia] = DateAdd("M", 1, Data_do_Movimento) - Day(DateAdd("M", 1, Data_do_Movimento))
         !Conta = "P"
         ![Seqüência Do Histórico] = 105
          If Sequencia_da_Agencia = 341 Then
             ![Codigo Do Debito] = 4
             ![Codigo Do Credito] = 5
          Else
             ![Codigo Do Debito] = 5
             ![Codigo Do Credito] = 4
          End If
          If Sequencia_da_Agencia = 1 Then   
             !Histórico = Historico
          Else
             !Histórico = Historico    
          End If   
         ![Tipo de Movimento da CC] = "Saída"
         ![Data da Alteração] = Date
         ![Hora da Alteração] = Time
         ![Usuário da Alteração] = vgPwUsuario
    .Update
    .BookMark = .LastModified
     Sequencia_da_Movimentacao_CC = ![Seqüência da Movimentação CC]
     LancamentoContabil ![Codigo Do Debito], ![Codigo Do Credito], Sequencia_da_Movimentacao_CC 
  End With
  With Manutencao 'conta Origen
    .AddNew 
       !Conta = "P"       
       !Documento = "Transferencia" 
       ![Sequencia Lan CC] = Sequencia_da_Movimentacao_CC 
       ![Parcela] = 1
        If Sequencia_da_Agencia = 1 Then 'BB
            ![Seqüência Do Geral] = 5017
         ElseIf Sequencia_da_Agencia = 8 Then 'Itau
            ![Seqüência Do Geral] = 5388
         ElseIf Sequencia_da_Agencia = 21 Then 
            ![Seqüência Do Geral] = 29064 'Caixinha  
         ElseIf Sequencia_da_Agencia = 19 Then 
            ![Seqüência Do Geral] = 28013 'sicob      
        End If 
        ![Data de Entrada] = Data_do_Movimento
        !Histórico = Historico 
        ![Forma de Pagamento] = "Prazo" 
        ![Valor Total] = Valor_Total 
        ![Data de Vencimento] = Data_do_Movimento 
        ![Vencimento Original] = Data_do_Movimento 
        ![Valor da Parcela] = Valor_Total 
        ![Valor Pago] = Valor_Total 
        ![Data da Baixa] = Data_do_Movimento
        ![Valor Restante] = 0
        ![Seqüência Grupo Despesa] = 34
        ![Seqüência SubGrupo Despesa] = 246   
    .Update
    .BookMark = .LastModified     
  End With
  'Baixa Origem
  Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa Contas]")
  TbAuxiliar "Baixa Contas", "[Seqüência da Movimentação CC] = " & Sequencia_da_Movimentacao_CC, Baixa
      Manutencao.MoveFirst: Manutencao.MoveLast
      With Baixa
       If Baixa.RecordCount = 0 Then .AddNew 
          ![Seqüência da Manutenção] = Manutencao![Seqüência da Manutenção]
          ![Data da Baixa] = Data_do_Movimento 
          ![Valor Pago] = Valor_Total 
          !Conta = "P"          
          ![Seqüência da Movimentação CC] = Sequencia_da_Movimentacao_CC 
          ![Seqüência da Agência] = Sequencia_da_Agencia
          ![Seqüência da CC da Agência] = Sequencia_da_CC_da_Agencia 
           If Sequencia_da_Agencia = 1 Then
              ![Codigo Do Historico] = 249 
           Else
              ![Codigo Do Historico] = 250       
           End If
           If Sequencia_da_Agencia = 341 Then
              ![Codigo Do Debito] = 4
              ![Codigo Do Credito] = 5
           Else
              ![Codigo Do Debito] = 5
              ![Codigo Do Credito] = 4
           End If 
          !Histórico = Historico
          ![Seqüência Grupo Despesa] = 34
          ![Seqüência SubGrupo Despesa] = 246
      .Update
      .BookMark = .LastModified
      Baixa.MoveLast
      End With
      With Manutencao
         .Edit
         ![Seqüência da Baixa] = Baixa![Seqüência da Baixa]
         .Update
         .BookMark = .LastModified
      End With
      SaldoContas Clng(Sequencia_da_Agencia), Clng(Sequencia_da_CC_da_Agencia)    
'Fim Origem  
  With Mvto 'Inicio Destino 
   .AddNew             
        ![Seqüência da Agência] = Sequencia_da_Agencia1 'Destino
         ![Seqüência da CC da Agência] = Sequencia_da_CC_da_Agencia1
          If Sequencia_da_Agencia1 = 1 Then
             ![Codigo Do Historico] = 249 
          Else
             ![Codigo Do Historico] = 250       
          End If
         ![Seqüência Grupo Despesa] = 34
         ![Seqüência SubGrupo Despesa] = 246
         ![Data Do Movimento] = Data_do_Movimento 
         ![Valor Total] = Valor_Total 
         ![Data Do Último Dia] = DateAdd("M", 1, Data_do_Movimento) - Day(DateAdd("M", 1, Data_do_Movimento))
         !Conta = "R"
         ![Seqüência Do Histórico] = 105
          If Sequencia_da_Agencia1 = 341 Then
             ![Codigo Do Debito] = 5
             ![Codigo Do Credito] = 4
          Else
             ![Codigo Do Debito] = 4
             ![Codigo Do Credito] = 5
          End If   
          If Sequencia_da_Agencia1 = 1 Then   
             !Histórico = Historico
          Else
             !Histórico = Historico    
          End If 
         ![Tipo de Movimento da CC] = "Entrada"
         ![Data da Alteração] = Date
         ![Hora da Alteração] = Time
         ![Usuário da Alteração] = vgPwUsuario
    .Update
    .BookMark = .LastModified
     Sequencia_da_Movimentacao_CC = ![Seqüência da Movimentação CC]
     LancamentoContabil ![Codigo Do Debito], ![Codigo Do Credito], Sequencia_da_Movimentacao_CC
  End With
  With Manutencao 'conta Destino
    .AddNew 
       !Conta = "R"       
       !Documento = "Transferencia" 
       ![Sequencia Lan CC] = Sequencia_da_Movimentacao_CC 
       ![Parcela] = 1
        If Sequencia_da_Agencia1 = 1 Then 'BB
            ![Seqüência Do Geral] = 5017
         ElseIf Sequencia_da_Agencia1 = 8 Then 'Itau
            ![Seqüência Do Geral] = 5388
         ElseIf Sequencia_da_Agencia1 = 21 Then 
            ![Seqüência Do Geral] = 29064 'Caixinha  
         ElseIf Sequencia_da_Agencia1 = 19 Then 
            ![Seqüência Do Geral] = 28013 'sicob      
        End If 
        ![Data de Entrada] = Data_do_Movimento
        !Histórico = Historico 
        ![Forma de Pagamento] = "Prazo" 
        ![Valor Total] = Valor_Total 
        ![Data de Vencimento] = Data_do_Movimento 
        ![Vencimento Original] = Data_do_Movimento 
        ![Valor da Parcela] = Valor_Total 
        ![Valor Pago] = Valor_Total 
        ![Data da Baixa] = Data_do_Movimento
        ![Valor Restante] = 0
        ![Seqüência Grupo Despesa] = 34
        ![Seqüência SubGrupo Despesa] = 246   
    .Update
    .BookMark = .LastModified     
  End With
  Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa Contas]")'Destino
  TbAuxiliar "Baixa Contas", "[Seqüência da Movimentação CC] = " & Sequencia_da_Movimentacao_CC, Baixa
      Manutencao.MoveFirst: Manutencao.MoveLast
      With Baixa
        .AddNew 
          ![Seqüência da Manutenção] = Manutencao![Seqüência da Manutenção]
          ![Data da Baixa] = Data_do_Movimento 
          ![Valor Pago] = Valor_Total 
          !Conta = "R"          
          ![Seqüência da Movimentação CC] = Sequencia_da_Movimentacao_CC 
          ![Seqüência da Agência] = Sequencia_da_Agencia1
          ![Seqüência da CC da Agência] = Sequencia_da_CC_da_Agencia1 
           If Sequencia_da_Agencia1 = 1 Then
              ![Codigo Do Historico] = 249 
           Else
              ![Codigo Do Historico] = 250       
           End If
           If Sequencia_da_Agencia1 = 341 Then
              ![Codigo Do Debito] = 4
              ![Codigo Do Credito] = 5
           Else
              ![Codigo Do Debito] = 5
              ![Codigo Do Credito] = 4
           End If 
          !Histórico = Historico
          ![Seqüência Grupo Despesa] = 34
          ![Seqüência SubGrupo Despesa] = 246
      .Update
      .BookMark = .LastModified
      Baixa.MoveLast
      End With
      With Manutencao
         .Edit
         ![Seqüência da Baixa] = Baixa![Seqüência da Baixa]
         .Update
         .BookMark = .LastModified
      End With  
 SaldoContas Clng(Sequencia_da_Agencia1), Clng(Sequencia_da_CC_da_Agencia1) 
 MsgBox("Transferencia Concluida com Sucesso!")

End Sub


Private Sub LancamentoContabil(Codigo_do_Debito As Long, Codigo_do_Credito As Long, SeqMvto As Long)
 Dim Lancamento As New GRecordSet
 Dim Tb As New GRecordSet

 If Sequencia_da_Agencia = 0 Or Sequencia_da_Agencia = 21 Then: Exit Sub

 Set Lancamento = vgDb.OpenRecordSet("Lançamentos Contabil")
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Id do Lançamento]) Seq From [Lançamentos Contabil]")

    With Lancamento
      .AddNew
      ![Id Do Lançamento] = Tb!Seq + 1
      ![Dt Do Lançamento] = Mid(Data_do_Movimento, 1,5)
      ![Conta Debito] = Codigo_do_Debito
      ![Conta Credito] = Codigo_do_Credito 
      ![Valor] = Valor_Total
       If Sequencia_da_Agencia1 = 1 Then
          ![Codigo Do Historico] = 249 
        Else
          ![Codigo Do Historico] = 250       
       End If
      ![Complemento Do Hist] = "Transferencia"
      ![Data da Baixa] = Data_do_Movimento
      ![Seqüência da Movimentação CC] = SeqMvto 
      .Update
      .BookMark = .LastModified
   End With  
   Unload Me    
End Sub


Private Sub Limpa()
  txtSequencia_da_CC_da_Agencia1.Text = ""
End Sub

Private Sub Limpa2()
  txtSequencia_da_CC_da_Agencia.Text = ""
End Sub

'Private Sub LancaNoContas()
' Dim Manutencao As New GRecordSet
' Dim Baixa As New GRecordset

' Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção Contas]")
 
' TbAuxiliar "Manutenção Contas", "[Sequencia Lan CC] = " & Sequencia_da_Movimentacao_CC, Manutencao
 
'  With Manutencao         
'    If Manutencao.RecordCount = 0 Then .AddNew         
'       If Tipo_do_Movimento_da_CC = "Saída" Then            
'          !Conta = "P"          
'       Else
'          !Conta = "R"                               
'       End If
'         !Documento = "Lancto Avulso"
'         ![Sequencia Lan CC] = Sequencia_da_Movimentacao_CC 
'         ![Parcela] = 1
'         If Sequencia_da_Agencia = 1 Then 'BB
'            ![Seqüência Do Geral] = 5017
'         ElseIf Sequencia_da_Agencia = 8 Then 'Itau
'            ![Seqüência Do Geral] = 5388
''         ElseIf Sequencia_da_Agencia = 18 Then 
'            ![Seqüência Do Geral] = 20822 ' Topazio    
'         ElseIf Sequencia_da_Agencia = 7 Then  
'            ![Seqüência Do Geral] = 4215 ' Bradesco 
'         ElseIf Sequencia_da_Agencia = 21 Then 
'            ![Seqüência Do Geral] = 29064 'Caixinha  
'         ElseIf Sequencia_da_Agencia = 19 Then 
'            ![Seqüência Do Geral] = 28013 'sicob      
'         End If       
'         ![Data de Entrada] = Data_do_Movimento 
'         !Histórico = Historico 
'         ![Forma de Pagamento] = "Prazo"
'         ![Valor Total] = Valor_Total 
'         ![Data de Vencimento] = Data_do_Movimento 
'         ![Vencimento Original] = Data_do_Movimento 
'         ![Valor da Parcela] = Valor_Total 
'         ![Valor Pago] = Valor_Total 
'         ![Data da Baixa] = Data_do_Movimento
'         ![Valor Restante] = 0
'         ![Seqüência Grupo Despesa] = Sequencia_Grupo_Despesa 
'         ![Seqüência SubGrupo Despesa] = Sequencia_SubGrupo_Despesa 
'         .Update
'         .BookMark = .LastModified           
'      End With          
'  Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa Contas]")
'  TbAuxiliar "Baixa Contas", "[Seqüência da Movimentação CC] = " & Sequencia_da_Movimentacao_CC, Baixa
'      Manutencao.MoveFirst: Manutencao.MoveLast
'      With Baixa
'       If Baixa.RecordCount = 0 Then .AddNew 
'          ![Seqüência da Manutenção] = Manutencao![Seqüência da Manutenção]
'          ![Data da Baixa] = Data_do_Movimento 
'          ![Valor Pago] = Valor_Total 
'           If Tipo_do_Movimento_da_CC = "Saída" Then            
'              !Conta = "P"          
'           Else
'              !Conta = "R"                               
'           End If
'          ![Seqüência da Movimentação CC] = Sequencia_da_Movimentacao_CC 
''          ![Seqüência da Agência] = Sequencia_da_Agencia
'          ![Seqüência da CC da Agência] = Sequencia_da_CC_da_Agencia 
'          ![Codigo Do Historico] = Codigo_do_Historico 
'          ![Codigo Do Debito] = Codigo_do_Debito 
'          ![Codigo Do Credito] = Codigo_do_Credito
'          !Histórico = Historico 
'           If !Conta = "P" Then
'             ![Seqüência Grupo Despesa] = Sequencia_Grupo_Despesa 
'             ![Seqüência SubGrupo Despesa] = Sequencia_SubGrupo_Despesa 
'           End If         
'      .Update
'      .BookMark = .LastModified
'       Baixa.MoveLast
'      End With
'      With Manutencao
'         .Edit
'         ![Seqüência da Baixa] = Baixa![Seqüência da Baixa]
'         .Update
'         .BookMark = .LastModified
'      End With    

'End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Sequencia_da_Agencia = txtCampo(0).Value
   Sequencia_da_CC_da_Agencia = txtCampo(1).Value
   Data_do_Movimento = txtCampo(2).Value
   Valor_Total = txtCampo(3).Value
   Sequencia_da_Agencia1 = txtCampo(4).Value
   Sequencia_da_CC_da_Agencia1 = txtCampo(5).Value
   Historico = txtCampo(6).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "([Seqüência da Agência] = " & 1 & ") OR [Seqüência da Agência] = " & 8 & ""
                           txtCampo(0).PesqSQLExpression = "SELECT Agências.[Seqüência da Agência], Agências.[Número do Banco], Agências.[Número da Agência], Agências.[Nome do Banco], " + _
                                                              "Agências.[Nome da Agência], Agências.CEP FROM Agências WHERE (Agências.[Seqüência da Agência] = " + CStr(1) + " OR " + _
                                                              "Agências.[Seqüência da Agência] = " + CStr(8) + ")"
      Case 1
         txtCampo(1).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "(([Seqüência da Agência] =  " & Sequencia_da_Agencia & ") AND [Seqüência da CC da Agência] > 0) AND Inativo = 0", "(([Seqüência da Agência] = " & Sequencia_da_Agencia & ") AND [Seqüência da CC da Agência] > 0)") & ""
                           txtCampo(1).PesqSQLExpression = "SELECT [Conta Corrente da Agência].[Seqüência da CC da Agência], [Conta Corrente da Agência].[Número da Conta Corrente], " + _
                                                              "[Conta Corrente da Agência].Descrição FROM [Conta Corrente da Agência] WHERE ([Conta Corrente da Agência].[Seqüência da CC da Agência] > " + CStr(0) + ") AND " + _
                                                              "([Conta Corrente da Agência].[Seqüência da Agência] = " + CStr([Sequencia_da_Agencia]) + ") AND ([Conta Corrente da Agência].Inativo = " + _
                                                              "False)"
      Case 4
         txtCampo(4).Filter = "([Seqüência da Agência] = " & 1 & ") OR [Seqüência da Agência] = " & 8 & ""
                           txtCampo(4).PesqSQLExpression = "SELECT Agências.[Seqüência da Agência], Agências.[Número do Banco], Agências.[Número da Agência], Agências.[Nome do Banco], " + _
                                                              "Agências.[Nome da Agência], Agências.CEP FROM Agências WHERE (Agências.[Seqüência da Agência] = " + CStr(1) + " OR " + _
                                                              "Agências.[Seqüência da Agência] = " + CStr(8) + ")"
      Case 5
         txtCampo(5).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "(([Seqüência da Agência] =  " & Sequencia_da_Agencia1 & ") AND [Seqüência da CC da Agência] > 0) AND Inativo = 0", "(([Seqüência da Agência] = " & Sequencia_da_Agencia1 & ") AND [Seqüência da CC da Agência] > 0)") & ""
                           txtCampo(5).PesqSQLExpression = "SELECT [Conta Corrente da Agência].[Seqüência da CC da Agência], [Conta Corrente da Agência].[Número da Conta Corrente], " + _
                                                              "[Conta Corrente da Agência].Descrição FROM [Conta Corrente da Agência] WHERE ([Conta Corrente da Agência].[Seqüência da CC da Agência] > " + CStr(0) + ") AND " + _
                                                              "([Conta Corrente da Agência].[Seqüência da Agência] = " + CStr([Sequencia_da_Agencia]) + ") AND ([Conta Corrente da Agência].Inativo = " + _
                                                              "False)"
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
      vgOk = (Not Vazio(Data_do_Movimento))
      vgMsg$ = "Data do Movimento não pode ser vazio!"
      If Not vgOk Then txtCampo(2).SetFocus
      If vgOk Then
         vgOk = (Valor_Total > 0)
         vgMsg$ = "Valor Total inválido!"
         If Not vgOk Then txtCampo(3).SetFocus
      End If
      If vgOk Then
         vgOk = (Sequencia_da_Agencia > 0)
         vgMsg$ = "Banco Origem Invalido"
         If Not vgOk Then txtCampo(0).SetFocus
      End If
      If vgOk Then
         vgOk = (Sequencia_da_CC_da_Agencia > 0)
         vgMsg$ = "Conta Origem Invalida!"
         If Not vgOk Then txtCampo(1).SetFocus
      End If
      If vgOk Then
         vgOk = (Sequencia_da_Agencia1 > 0)
         vgMsg$ = "Banco Destino Invalido"
         If Not vgOk Then txtCampo(4).SetFocus
      End If
      If vgOk Then
         vgOk = (Sequencia_da_CC_da_Agencia1 > 0)
         vgMsg$ = "Conta Destino Invalida!"
         If Not vgOk Then txtCampo(5).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
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
   Select Case Index
      Case 0
         Limpa2  
      Case 4
         Limpa  
   End Select
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 0
            PoeRelEFiltroCbo 0
         Case 1
            PoeRelEFiltroCbo 1
         Case 2
            If Len(txtCp(2).Text) = 0 Then
               txtCampo(2).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
            End If
         Case 4
            PoeRelEFiltroCbo 4
         Case 5
            PoeRelEFiltroCbo 5
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
   Caption = LoadGasString(144490)
   vgFormID = 1438
   vgIdentTab = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(135995))
   Set botOk.Picture = LoadPicture(LoadGasPicture(135996))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(135997))
   Set txtSequencia_da_CC_da_Agencia = txtCampo(1)
   Set lblAjuste = Label(2)
   Set txtSequencia_da_CC_da_Agencia1 = txtCampo(5)
   Set txtHistorico = txtCampo(6)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(144495)
   Label(1).Caption = LoadGasString(144496)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(144497)
   Label(2).Caption = LoadGasString(144498)
   Label(3).Caption = LoadGasString(144499)
   Label(4).Caption = LoadGasString(144500)
   Label(5).Caption = LoadGasString(144501)
   Label(6).Caption = LoadGasString(144502)
   Label(7).Caption = LoadGasString(144503)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(144504)
   vgTooltips.AddTool botOk, 0, LoadGasString(144505)
   vgTooltips.AddTool botCancela, 0, LoadGasString(144506)
   Label(8).Caption = LoadGasString(144507)
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

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_ACAO) = bottxtCampo2(BOT_ACAO)
   Set bottxtCampo2(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(2).DataType = 2
   txtCampo(2).Mask = "99/99/9999"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_ACAO) = bottxtCampo3(BOT_ACAO)
   Set bottxtCampo3(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "9.999.999,9999"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(0).CtPri = txtCp(0)
   Set txtCampo(0).CtFdo = labFdo0
   Set txtCampo(0).CtBot(BOT_LISTA) = bottxtCampo0(BOT_LISTA)
   Set txtCampo(0).CtBot(BOT_COMBO) = bottxtCampo0(BOT_COMBO)
   bottxtCampo0(BOT_LISTA).Caption = "P"
   Set bottxtCampo0(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "9999"
   txtCampo(0).PesqModoAbertura =  2 
   txtCampo(0).PesqFieldCapture = "Agências.[Seqüência da Agência]"
   txtCampo(0).BoundColumn = "Seqüência da Agência"
   txtCampo(0).ListFields = "Nome do Banco; Número do Banco"
   txtCampo(0).OrderFields = "Nome do Banco; Número do Banco"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Agências"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_LISTA) = bottxtCampo1(BOT_LISTA)
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   bottxtCampo1(BOT_LISTA).Caption = "P"
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "999"
   txtCampo(1).PesqModoAbertura =  2 
   txtCampo(1).PesqFieldCapture = "[Conta Corrente da Agência].[Seqüência da CC da Agência]"
   txtCampo(1).BoundColumn = "Seqüência da CC da Agência"
   txtCampo(1).ListFields = "Número da Conta Corrente"
   txtCampo(1).OrderFields = "Número da Conta Corrente"
   txtCampo(1).Relation = ""
   txtCampo(1).Source = "Conta Corrente da Agência"
   txtCampo(1).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_LISTA) = bottxtCampo4(BOT_LISTA)
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   bottxtCampo4(BOT_LISTA).Caption = "P"
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "9999"
   txtCampo(4).PesqModoAbertura =  2 
   txtCampo(4).PesqFieldCapture = "Agências.[Seqüência da Agência]"
   txtCampo(4).BoundColumn = "Seqüência da Agência"
   txtCampo(4).ListFields = "Nome do Banco; Número do Banco"
   txtCampo(4).OrderFields = "Nome do Banco; Número do Banco"
   txtCampo(4).Relation = ""
   txtCampo(4).Source = "Agências"
   txtCampo(4).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_LISTA) = bottxtCampo5(BOT_LISTA)
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   bottxtCampo5(BOT_LISTA).Caption = "P"
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "999"
   txtCampo(5).PesqModoAbertura =  2 
   txtCampo(5).PesqFieldCapture = "[Conta Corrente da Agência].[Seqüência da CC da Agência]"
   txtCampo(5).BoundColumn = "Seqüência da CC da Agência"
   txtCampo(5).ListFields = "Número da Conta Corrente"
   txtCampo(5).OrderFields = "Número da Conta Corrente"
   txtCampo(5).Relation = ""
   txtCampo(5).Source = "Conta Corrente da Agência"
   txtCampo(5).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 4
   txtCampo(6).Mask = ""
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

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
   Lancar  
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
   Set txtSequencia_da_CC_da_Agencia = Nothing
   Set lblAjuste = Nothing
   Set txtSequencia_da_CC_da_Agencia1 = Nothing
   Set txtHistorico = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmtransfer = Nothing                      'libera o segmento de código do form
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
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

