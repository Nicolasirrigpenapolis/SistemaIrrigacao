VERSION 5.00
Begin VB.Form frmGexcel
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   4000
   Left                     =   5760
   LinkTopic                =   "Gexcel"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   3565
   ScaleWidth               =   10110
   Top                      =   4720
   Width                    =   10200
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
      Height                   =   3405
      Left                     =   15
      Width                    =   9930
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   17506
      _extenty                 =   6003
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1455
         Left                     =   75
         Width                    =   1380
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
         Width                    =   960
         Top                      =   315
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
         Caption                  =   "Geral"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1755
         Y1                       =   705
         X2                       =   2640
         Y2                       =   705
         Index                    =   0 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   255
         Left                     =   1755
         Width                    =   1125
         Top                      =   915
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Cliente"
         Index                    =   0 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   0
         Height                   =   255
         Left                     =   2955
         Width                    =   2025
         Top                      =   915
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Transportadora"
         Index                    =   1 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   0
         Height                   =   255
         Left                     =   5115
         Width                    =   1665
         Top                      =   915
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Fornecedor"
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
         Width                    =   1020
         Top                      =   1635
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
         Index                    =   1 
      End
      Begin VB.Label labFdo0
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1395
         Top                      =   1635
         Width                    =   1920
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   3045
         Top                      =   1650
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo0
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   2790
         Top                      =   1650
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
         Left                     =   1410
         Width                    =   1365
         Top                      =   1650
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3495
         Width                    =   2190
         Top                      =   1635
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
         Caption                  =   "Razão Social / Fantasia:"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   9615
         Width                    =   45
         Top                      =   3195
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
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   300
         Left                     =   5835
         Width                    =   3900
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
         Left                     =   375
         Width                    =   900
         Top                      =   2175
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
         Caption                  =   "Municipio:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1395
         Top                      =   2175
         Width                    =   4440
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   5565
         Top                      =   2190
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   5310
         Top                      =   2190
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
         Left                     =   1410
         Width                    =   3885
         Top                      =   2190
         Index                    =   2 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   6
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   375
         Left                     =   6255
         Width                    =   1665
         Top                      =   2415
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
         Index                    =   0 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   7
         TabStop                  =   0
         Height                   =   255
         Left                     =   6855
         Width                    =   1140
         Top                      =   915
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Inativo"
         Index                    =   3 
      End
      Begin NFE.GBotao botCancela
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   8
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   375
         Left                     =   7980
         Width                    =   1665
         Top                      =   2415
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
         Caption                  =   "Cancelar"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   495
         Width                    =   810
         Top                      =   2730
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
         Caption                  =   "Diretorio:"
         Index                    =   5 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1380
         Top                      =   2730
         Width                    =   3690
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4800
         Top                      =   2745
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
         MaxLength                =   255
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   9
         TabStop                  =   0
         Height                   =   270
         Left                     =   1395
         Width                    =   3390
         Top                      =   2745
         Index                    =   3 
      End
   End
End
Attribute VB_name = "frmGexcel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: Gexcel
'* Função....: Geral - Exportar Excel
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:49
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
Dim chkCampo(3) As New FormataCampos              'classe dos compos tipo lógico do módulo 
Dim Cliente As Boolean, Transportadora As Boolean, Fornecedor As Boolean
Dim Sequencia As Double, Nome As String, Municipio As Double
Dim Inativo As Boolean, Diretorio As String
Public lblAjuste As Object, txtDiretorio As Object

Private Sub ExportarGeral()
 Dim Tb As New GRecordSet
 Dim vArquivo As Long, Excel As Object, L As Long
 
 Set Tb = vgDb.OpenRecordSet("SELECT Geral.[Sequencia Do Geral] As Seq, [Razão Social], Geral.[Endereço], Geral.Celular, Geral.[Cpf e cnpj] As CNPJ, " & _
                             "Municipios.[Descrição do Municipio] As Mun, Municipios.UF, Geral.Bairro, Geral.Fone, Geral.[Rg e ie] As Rg, Geral.Inativo, Geral.Nro, Geral.CEP " & _
                             "FROM Geral LEFT JOIN Municipios ON Geral.[Sequencia Do Municipio] = Municipios.[Sequencia Do Municipio] " & _
                             "WHERE " & Filtro() & _
                             " ORDER BY Geral.[Sequencia do Geral]")  
 On Error Resume Next
      
 If Tb.RecordCount > 0 Then
     vArquivo = FreeFile
     If Dir(txtDiretorio.text & "Clientes.xls", vbArchive) <> "" Then Kill txtDiretorio.text & "Clientes.xls"
        Open txtDiretorio.text & "Clientes.xls" For Append As #vArquivo 
        Set Excel = CreateObject("Excel.Application")
      
        Set Excel = Excel.Workbooks.Add
        Set Excel = Excel.Sheets(1)
        Excel.SaveAs vArquivo
        Screen.MousePointer = vbHourglass 'Ampulheta
        'Status
        Me.Caption = "Exportando Clientes para Excel - Exportação " '& " " & Info(Tb)
        Excel.Application.Quit
        Set Excel = CreateObject("Excel.Application")
        With Excel
            .Workbooks.Open FileName:=vArquivo
            .Visible = False
            .Sheets("Plan1").Select
            .Range("A1").Select
            .activecell.Value = "Reticom Virabrequim"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 12
            .activecell.Font.Color = vbGreen
            
            .Range("A2").Select
            .activecell.Value = "Sequencia"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("B2").Select
            .activecell.Value = "Razão Social"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("C2").Select
            .activecell.Value = "Município"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("D2").Select
            .activecell.Value = "UF"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("E2").Select
            .activecell.Value = "Endereço"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("F2").Select
            .activecell.Value = "Nº"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("G2").Select
            .activecell.Value = "CEP"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("H2").Select
            .activecell.Value = "Bairro"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("I2").Select
            .activecell.Value = "Telefone"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("J2").Select
            .activecell.Value = "Celular"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("K2").Select
            .activecell.Value = "CPF/CNPJ"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("L2").Select
            .activecell.Value = "RG/IE"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            '.Range("K2").Select
            '.activecell.Value = "Vendedor"
            '.activecell.Font.Name = "Arial Black"
            '.activecell.Font.Size = 10
            
            '.Range("L2").Select
            '.activecell.Value = "Observações"
            '.activecell.Font.Name = "Arial Black"
            '.activecell.Font.Size = 10
            
          L = 2     
          Do While Not Tb.EOF 
          L = L + 1       
            .Range("A" & Trim(str(L))).Select
            .activecell.Font.Name = "Courier New"
            .activecell.Font.Size = 10
            .activecell.Value = Tb!Seq
            'Sequencia do Geral 
                   
            .Range("B" & Trim(str(L))).Select
            .activecell.Font.Name = "Courier New"
            .activecell.Font.Size = 10
            .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb![Razão Social])), True) 
            'Razão Social
            
            .Range("C" & Trim(str(L))).Select
            .activecell.Font.Name = "Courier New"
            .activecell.Font.Size = 10
            .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb!Mun)), True) 
            'Municipio
            
            .Range("D" & Trim(str(L))).Select
            .activecell.Font.Name = "Courier New"
            .activecell.Font.Size = 10
            .activecell.Value = Tb!UF
            'UF
           
           .Range("E" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb![Endereço])), True) 
           'Endereço
           
           .Range("F" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb!Nro)), True)   
           'Nro
           
           .Range("G" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Format(Tb!CEP,"99999-999"))), True)   
           'cep
           
           .Range("H" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb!Bairro)), True)   
           'Bairro

           .Range("I" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Fone  
           'Fone
           
           .Range("J" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Celular
           'Cel
           
           .Range("K" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!CNPJ
           
           .Range("L" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!RG
                      
           '.Range("K" & Trim(str(L))).Select
           '.activecell.Font.Name = "Courier New"
           '.activecell.Font.Size = 10
           '.activecell.Value = Tb!Vendedor
           
           '.Range("L" & Trim(str(L))).Select
           '.activecell.Font.Name = "Courier New"
           '.activecell.Font.Size = 10
           '.activecell.Value = Tb!Observacao
              
         'VIOS.MoveNext
         'Loop
        ' VIOS.MoveFirst      
         Tb.MoveNext
         Loop
         Tb.MoveFirst       
    GoTo Fim
Fim:  
     MsgBox "Por Favor Renomear Planilha e Salvar Como no diretorio Escolhido!", vbExclamation + vbOKOnly, "ATENÇÃO..."
    .Visible = True
    .Range("A1").ColumnWidth = 10
    .Range("B1").ColumnWidth = 6  ' Sequencia do Geral 
    .Range("B2").ColumnWidth = 70  'Razão Social 
    .Range("C2").ColumnWidth = 35  'Municipio 
    .Range("D2").ColumnWidth = 4   'UF 
    .Range("E2").ColumnWidth = 60  'Endereço 
    .Range("F2").ColumnWidth = 6   'Nro 
    .Range("G2").ColumnWidth = 12  'cep 
    .Range("H2").ColumnWidth = 45  'Bairro 
    .Range("I2").ColumnWidth = 20  'Fone 
    .Range("J2").ColumnWidth = 20 ' Cel 
    .Range("K2").ColumnWidth = 23 ' CNPJ    
    .Range("L2").ColumnWidth = 23 ' RG 
    '.Range("K2").ColumnWidth = 60 ' Vendedor 
    '.Range("L2").ColumnWidth = 90 ' Vendedor 
   
    .Columns("B:B").Select
    .Selection.Style = "Currency"  
    .Columns("C:C").Select
    .Selection.Style = "Currency"
    '.Columns("D:D").Select
    '.Selection.Style = "Currency" 
    .Columns("E:E").Select
    .Selection.Style = "Currency"
    '.Columns("F:F").Select
    '.Selection.Style = "Currency"        
   
     End With  
    Else
    MsgBox "Não consta Registros para estes Filtros!", vbExclamation + vbOKOnly, "Atenção..."
End If  
Finaliza:
      
    Err.Clear
    Screen.MousePointer = vbDefault                'ponteiro do mouse = default
    'Excel.SaveAs vArquivo
    'Excel.save
    Close #vArquivo 
    Me.Caption = "Exportado com Sucesso" 'Status
    Screen.MousePointer = vbDefault 'Normal
       
End Sub


'ROTINA para tirar os caracteres especiais
Public Function RemoveCaracteresExcel(texto As String, Optional excessao As Boolean) As String
   Dim Vector As String, x As Long, y As Integer, Retorno As String, Igual As Boolean
   
   If Not excessao Then
      Vector = "!@#$%¨&*()-_=+'¹²³£¢¬§´`~^;:.,<>/?°][{}ºª\| .-"
   Else 'Usado em Nota Fiscais e Etc...
      Vector = "!@#$%¨&*_=+¹²³£¢¬§´`~^;:<>/?°][{}ºª\|. -,"
   End If
   
   y = 1
      
   For x = 1 To Len(texto)
      y = 1
      Do While y <= Len(Vector)
         If Mid(texto, x, 1) = Mid(Vector, y, 1) Then
            Igual = True
            Exit Do
         End If
         y = y + 1
      Loop
      If Not Igual Then
         Retorno$ = Retorno$ & Mid(texto, x, 1)
      End If
      Igual = False
   Next
   
   RemoveCaracteresExcel = Retorno$
   
End Function                              
                             
                                            
Private Function Filtro() As String
   On Error Resume Next
   
   Filtro = "Geral.[Sequencia do Geral] > 0"
   If Sequencia = 0 Then
      If Not Vazio(Nome) Then 'Geral
         Filtro = Filtro & " AND ([Razão Social] Like '%" & Nome & "%'"
         Filtro = Filtro & " Or [Nome Fantasia] Like '%" & Nome & "%')"
      End If      
      If Municipio > 0 Then 'Município
         Filtro = Filtro & " AND Geral.[Sequencia do Municipio] = " & Municipio
      End If
      If Cliente Then 'Cliente
         Filtro = Filtro & " AND Cliente = 1"
      End If
      If Fornecedor Then 'Fornecedor
         Filtro = Filtro & " AND Fornecedor = 1"
      End If
      If Transportadora Then 'Tranportadora
         Filtro = Filtro & " AND Transportadora = 1"
      End If
      If Inativo Then 'Inativo
         Filtro = Filtro & " AND Geral.Inativo = 1"
      End If
   Else
      Filtro = Filtro & " AND Geral.[Sequencia do Geral] = " & Sequencia
   End If
   
End Function

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Cliente = chkCampo(0).Value
   Transportadora = chkCampo(1).Value
   Fornecedor = chkCampo(2).Value
   Sequencia = txtCampo(0).Value
   Nome = txtCampo(1).Value
   Municipio = txtCampo(2).Value
   Inativo = chkCampo(3).Value
   Diretorio = txtCampo(3).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 0
         txtCampo(0).Filter = "[Sequencia do geral] > " & 0 & ""
                           txtCampo(0).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ")"
      Case 2
         txtCampo(2).Filter = "[Sequencia do municipio] > " & 0 & ""
                           txtCampo(2).PesqSQLExpression = "SELECT Municipios.[Sequencia do municipio], Municipios.[Descrição do municipio], Municipios.Uf, Municipios.Cep FROM Municipios WHERE (Municipios.[Sequencia do municipio] > " + CStr(0) + ")"
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
            chkCampo(0).Value = False
      chkCampo(1).Value = False
      chkCampo(2).Value = False
      chkCampo(3).Value = False
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
   Label(5).Visible = (False)
   txtCampo(3).Visible = (False)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 0 Or Index = 2 Then
      ExecutaVisivel
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
   InicializaApelidos COM_TEXTBOX                 'inicializa apelidos dos campos
   ExecutaVisivel                                 'torna camos visiveis
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel                                 'torna camos visiveis
   chkCampo(Index).Change
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

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(29540)
   vgFormID = 714
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(21045))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(21046))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(21047))
   Set lblAjuste = Label(3)
   Set txtDiretorio = txtCampo(3)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(29545)
   chkCampo(0).Caption = LoadGasString(29546)
   chkCampo(1).Caption = LoadGasString(29547)
   chkCampo(2).Caption = LoadGasString(29548)
   Label(1).Caption = LoadGasString(29549)
   Label(2).Caption = LoadGasString(29550)
   Label(3).Caption = LoadGasString(29551)
   Label(4).Caption = LoadGasString(29552)
   Botao(0).Caption = LoadGasString(29553)
   chkCampo(3).Caption = LoadGasString(29554)
   vgTooltips.AddTool botCancela, 0, LoadGasString(29555)
   botCancela.Caption = LoadGasString(29556)
   Label(5).Caption = LoadGasString(29557)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(29558)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
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

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5

   Set chkCampo(2).CtPri = chkCp(2)
   chkCampo(2).DataType = 5

   Set txtCampo(0).CtPri = txtCp(0)
   Set txtCampo(0).CtFdo = labFdo0
   Set txtCampo(0).CtBot(BOT_LISTA) = bottxtCampo0(BOT_LISTA)
   Set txtCampo(0).CtBot(BOT_COMBO) = bottxtCampo0(BOT_COMBO)
   bottxtCampo0(BOT_LISTA).Caption = "P"
   Set bottxtCampo0(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = ""
   txtCampo(0).PesqModoAbertura =  2 
   txtCampo(0).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(0).BoundColumn = "Sequencia do geral"
   txtCampo(0).ListFields = "Sequencia do geral"
   txtCampo(0).OrderFields = "Sequencia do geral"
   txtCampo(0).Relation = ""
   txtCampo(0).Source = "Geral"
   txtCampo(0).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = "@x"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(2).CtPri = txtCp(2)
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_LISTA) = bottxtCampo2(BOT_LISTA)
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   bottxtCampo2(BOT_LISTA).Caption = "P"
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = ""
   txtCampo(2).PesqModoAbertura =  2 
   txtCampo(2).PesqFieldCapture = "Municipios.[Sequencia do municipio]"
   txtCampo(2).BoundColumn = "Sequencia do municipio"
   txtCampo(2).ListFields = "Descrição do municipio"
   txtCampo(2).OrderFields = "Sequencia do municipio"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Municipios"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True

   Set chkCampo(3).CtPri = chkCp(3)
   chkCampo(3).DataType = 5

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_ACAO) = bottxtCampo3(BOT_ACAO)
   Set bottxtCampo3(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@x"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub

'evento - cancela formulário
Private Sub botCancela_Click()
   Unload Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set lblAjuste = Nothing
   Set txtDiretorio = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   Set chkCampo(3) = Nothing
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmGexcel = Nothing                        'libera o segmento de código do form
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
         ExportarGeral  
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 3
         PegaDir2 Me: txtDiretorio.Text = vaDir
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo0_Click(Index As Integer)
   txtCampo(0).SetFocus
   DoEvents
   txtCampo(0).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo3_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 3, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(3).botClick Index
   End If
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

