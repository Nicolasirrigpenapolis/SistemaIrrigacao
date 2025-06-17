Version 5.00
Begin VB.MDIForm mdiNFE
   Height                   =   3000
   Left                     =   1440
   LinkTopic                =   "NFE"
   LockControls             =   -1
   AutoShowChildren         =   0
   NegotiateToolbars=   0
   Icon                     =   "NFE.frx":0000
   Top                      =   2260
   Width                    =   8670
   Begin NFE.GPainel panTools
      AutoRedraw               =   -1
      BackColor                =   &HE0E0E0
      BevelOuter               =   2
      BevelInner               =   0
      BorderWidth              =   0
      BevelWidth               =   1
      BorderStyle              =   0
      Align                    =   1
      Stretch                  =   -1
      Stripes                  =   0
      SaveGridBars             =   0
      Height                   =   1110
      Left                     =   -75000
      Width                    =   8385
      Top                      =   0
      TabStop                  =  0
      _extentx                 =   14782
      _extenty                 =   1956
      Begin NFE.GPainel Painel
         AutoRedraw               =   -1
         BackColor                =   &H8000000F
         BevelOuter               =   1
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Align                    =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   975
         Left                     =   60
         Width                    =   6000
         Top                      =   75
         TabStop                  =  0
         _extentx                 =   10577
         _extenty                 =   1718
         Index                    =   0 
         Begin NFE.GBotao botInclui
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   1
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   75
            Width                    =   1440
            Top                      =   75
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Incluir"
            TabStop                  =   0
         End
         Begin NFE.GBotao botCancela
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   2
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   1515
            Width                    =   1440
            Top                      =   75
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Cancelar"
            TabStop                  =   0
         End
         Begin NFE.GBotao botBrowse
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   3
            TabStop                  =   -1
            ButtonType               =   1
            Height                   =   360
            Left                     =   2955
            Width                    =   1440
            Top                      =   75
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Grade"
            TabStop                  =   0
         End
         Begin VB.PictureBox picNavega
            BorderStyle              =   0
            Left                     =   4395
            Height                   =   360
            Top                      =   75
            Width                    =   1480
            TabStop                  =   0
            Begin VB.HScrollBar schNavega
               LargeChange              =   10
               SmallChange              =   1
               Max                      =   32767
               Min                      =   0
               Height                   =   360
               Left                     =   0
               Width                    =   1480
               Top                      =   0
               Max                      =  101
               Min                      =  -1
               LargeChange              =  10
               SmallChange              =  1
            End
         End
         Begin NFE.GBotao botPrimeiro
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   4
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   4395
            Width                    =   360
            Top                      =   555
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
            TabStop                  =   0
         End
         Begin NFE.GBotao botAnterior
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   5
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   4770
            Width                    =   360
            Top                      =   555
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
            TabStop                  =   0
         End
         Begin NFE.GBotao botSeguinte
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   6
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   5145
            Width                    =   360
            Top                      =   555
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
            TabStop                  =   0
         End
         Begin NFE.GBotao botUltimo
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   7
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   5520
            Width                    =   360
            Top                      =   555
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
            TabStop                  =   0
         End
         Begin NFE.GBotao botImprime
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   8
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   2955
            Width                    =   1440
            Top                      =   555
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Print"
            TabStop                  =   0
         End
         Begin NFE.GBotao botExclui
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   9
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   1515
            Width                    =   1440
            Top                      =   555
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Excluir"
            TabStop                  =   0
         End
         Begin NFE.GBotao botSalva
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   10
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   360
            Left                     =   75
            Width                    =   1440
            Top                      =   555
            CaptionAlign             =   4
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H0
            Caption                  =   "Gravar"
            TabStop                  =   0
         End
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   0
         TabStop                  =   -1
         ButtonType               =   1
         Height                   =   180
         Left                     =   240
         Width                    =   300
         Top                      =   1740
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
         Index                    =   0 
         TabStop                  =   0
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   495
         Left                     =   420
         Width                    =   5520
         Top                      =   300
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   21.75
            Underline                =   0
            Italic                   =   -1
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   ""
         Index                    =   0 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   11
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   6240
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Geral"
         Index                    =   1 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   12
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   7320
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Produtos"
         Index                    =   2 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   13
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   8400
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "NFe"
         Index                    =   3 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   14
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   13800
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Consultas"
         Index                    =   4 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   15
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   9480
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Estoque"
         Index                    =   5 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   16
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   10560
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Receber"
         Index                    =   6 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   17
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   11640
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Pagar"
         Index                    =   7 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   18
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   12720
         Width                    =   1080
         Top                      =   75
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "O. S"
         Index                    =   8 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   19
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   975
         Left                     =   12720
         Width                    =   1080
         Top                      =   70
         CaptionAlign             =   1
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Vendas"
         Index                    =   9 
      End
   End
   Begin NFE.GPainel Painel
      AutoRedraw               =   -1
      BackColor                =   &HE0E0E0
      BevelOuter               =   1
      BevelInner               =   2
      BorderWidth              =   1
      BevelWidth               =   1
      BorderStyle              =   0
      Align                    =   2
      Stretch                  =   -1
      Stripes                  =   0
      SaveGridBars             =   0
      Height                   =   390
      Left                     =   -75000
      Width                    =   8385
      Top                      =   2010
      TabStop                  =  0
      _extentx                 =   14782
      _extenty                 =   687
      Index                    =   1 
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   255
         Left                     =   60
         Width                    =   2040
         Top                      =   60
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Caption                  =   "Grupo"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   255
         Left                     =   2160
         Width                    =   2040
         Top                      =   60
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Usuario"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   255
         Left                     =   4260
         Width                    =   2040
         Top                      =   60
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Data e Hora"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   255
         Left                     =   6360
         Width                    =   2040
         Top                      =   60
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Pc"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   9420
         Width                    =   60
         Top                      =   60
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
         Caption                  =   ""
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &HE0E0E0
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   14580
         Width                    =   60
         Top                      =   60
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
         Caption                  =   ""
         Index                    =   6 
      End
                  Begin VB.Timer timRefresh
                           Enabled         =   0
                           Interval        =   5000
                           left            =   6120
                           top             =   60
                  End
                  Begin VB.PictureBox picAux 
                           AutoRedraw      =   -1                 'True   
                           AutoSize        =   -1                 'True   
                           BorderStyle     =   0                  'None    
                           Height          =   255         
                           Left            =   6520
                           ScaleHeight     =   255         
                           ScaleWidth      =   285         
                           TabIndex        =   18          
                           Top             =   60
                           Visible         =   0                  'False  
                           Width           =   285         
                  End
         End
   Begin VB.Menu Menu_570
      Caption         =   "&Arquivo"
      Index           =   0
      Tag             =   "570|1"
      Begin VB.Menu Menu_Consulta
         Caption         =   "C&onstrução de consultas"
         Tag             =   "571|2"
      End
      Begin VB.Menu Menu_Repara
         Caption         =   "&Reparação"
         Tag             =   "572|2"
      End
      Begin VB.Menu Menu_Info
         Caption         =   "&Informações sobre o BD"
         Tag             =   "573|2"
      End
      Begin VB.Menu Menu_ConfImp
         Caption         =   "&Configuração impressora"
         Tag             =   "574|2"
      End
      Begin VB.Menu Menu_Senhas
         Caption         =   "&Manutenção de senhas"
         Tag             =   "575|2"
      End
      Begin VB.Menu Menu_N1
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Fim
         Caption         =   "&Finalizar"
         Tag             =   "577|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Editar"
      Index           =   1
      Tag             =   "578|1"
      Begin VB.Menu Menu_Salva
         Caption         =   "&Gravar"
         Shortcut        =   ^S
         Tag             =   "579|2"
      End
      Begin VB.Menu Menu_Cancela
         Caption         =   "&Cancelar"
         Shortcut        =   ^A
         Tag             =   "580|2"
      End
      Begin VB.Menu Menu_Inclui
         Caption         =   "&Incluir"
         Shortcut        =   ^I
         Tag             =   "581|2"
      End
      Begin VB.Menu Menu_Exclui
         Caption         =   "&Excluir"
         Shortcut        =   ^E
         Tag             =   "582|2"
      End
      Begin VB.Menu Menu_N2
         Caption         =   "-"
      End
      Begin VB.Menu Menu_ApagaCol
         Caption         =   "Apagar co&luna"
         Shortcut        =   ^L
         Tag             =   "584|2"
      End
      Begin VB.Menu Menu_Procura
         Caption         =   "&Procurar"
         Shortcut        =   ^P
         Tag             =   "585|2"
      End
      Begin VB.Menu Menu_Filtra
         Caption         =   "&Filtrar"
         Shortcut        =   ^F
         Tag             =   "586|2"
      End
      Begin VB.Menu Menu_Grafa
         Caption         =   "G&rafar"
         Shortcut        =   ^G
         Tag             =   "587|2"
      End
      Begin VB.Menu Menu_Imprime
         Caption         =   "I&mprimir"
         Tag             =   "588|2"
      End
      Begin VB.Menu Menu_ConfGraf
         Caption         =   "Co&nfigurar"
         Tag             =   "589|2"
      End
      Begin VB.Menu Menu_Ortog
         Caption         =   "&Ortografia"
         Tag             =   "590|2"
      End
      Begin VB.Menu Menu_Repeticao
         Caption         =   "&Repetir campos"
         Shortcut        =   {F3}
         Tag             =   "591|2"
      End
      Begin VB.Menu Menu_FRM00621
         Caption         =   "Capitaliza"
         Tag             =   "621|2"
      End
      Begin VB.Menu Menu_N3
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Primeiro
         Caption         =   "Primeir&o"
         Shortcut        =   {F5}
         Tag             =   "594|2"
      End
      Begin VB.Menu Menu_Anterior
         Caption         =   "&Anterior"
         Shortcut        =   {F6}
         Tag             =   "595|2"
      End
      Begin VB.Menu Menu_Seguinte
         Caption         =   "&Seguinte"
         Shortcut        =   {F7}
         Tag             =   "596|2"
      End
      Begin VB.Menu Menu_Ultimo
         Caption         =   "&Ultimo"
         Shortcut        =   {F8}
         Tag             =   "597|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Parametros"
      Index           =   2
      Tag             =   "613|1"
      Begin VB.Menu Menu_614
         Caption         =   "&Parametros do Emitente"
         Index           =   0
         Tag             =   "614|2"
      End
      Begin VB.Menu Menu_614
         Caption         =   "&Parametros da NFe"
         Index           =   1
         Tag             =   "615|2"
      End
      Begin VB.Menu Menu_614
         Caption         =   "&Parametros da O.S"
         Index           =   2
         Tag             =   "677|2"
      End
      Begin VB.Menu Menu_614
         Caption         =   "Registro Zero"
         Index           =   3
         Tag             =   "616|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Cadastro"
      Index           =   3
      Tag             =   "611|1"
      Begin VB.Menu Menu_717
         Caption         =   "&Bancos"
         Index           =   0
         Tag             =   "717|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Icms"
         Index           =   1
         Tag             =   "612|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Classificação Fiscal"
         Index           =   2
         Tag             =   "617|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Natureza de Operação"
         Index           =   3
         Tag             =   "618|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Municípios"
         Index           =   4
         Tag             =   "619|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Geral"
         Index           =   5
         Tag             =   "620|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Unidades"
         Index           =   6
         Tag             =   "622|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Marcas"
         Index           =   7
         Tag             =   "681|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Tipo de Cobrança"
         Index           =   8
         Tag             =   "695|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Centro de Custo"
         Index           =   9
         Tag             =   "708|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Produtos"
         Index           =   10
         Tag             =   "623|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Serviços"
         Index           =   11
         Tag             =   "678|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Virabrequim"
         Index           =   12
         Tag             =   "682|2"
      End
      Begin VB.Menu Menu_717
         Caption         =   "&Modelo do Virabrequim"
         Index           =   13
         Tag             =   "715|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Faturamento"
      Index           =   4
      Tag             =   "702|1"
      Begin VB.Menu Menu_709
         Caption         =   "Orçamento"
         Index           =   0
         Tag             =   "709|2"
      End
      Begin VB.Menu Menu_709
         Caption         =   "&Vendas"
         Index           =   1
         Tag             =   "703|2"
      End
      Begin VB.Menu Menu_709
         Caption         =   "&Relatorios"
         Index           =   2
         Tag             =   "728|2"
         Begin VB.Menu Menu_729
            Caption         =   "Relatório de Vendas"
            Index           =   0
            Tag             =   "729|3"
         End
         Begin VB.Menu Menu_729
            Caption         =   "&Relatório de Vendas por Periodo"
            Index           =   1
            Tag             =   "735|3"
         End
         Begin VB.Menu Menu_729
            Caption         =   "Relatório do Estoque"
            Index           =   2
            Tag             =   "733|3"
         End
      End
      Begin VB.Menu Menu_709
         Caption         =   "&Ocultos"
         Index           =   3
         Tag             =   "706|2"
         Begin VB.Menu Menu_707
            Caption         =   "Vendas - Nota Fiscal"
            Index           =   0
            Tag             =   "707|3"
         End
         Begin VB.Menu Menu_707
            Caption         =   "Imprimir"
            Index           =   1
            Tag             =   "711|3"
         End
         Begin VB.Menu Menu_707
            Caption         =   "Imprimir - Meia Folha"
            Index           =   2
            Tag             =   "734|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Nota Fiscal"
      Index           =   5
      Tag             =   "624|1"
      Begin VB.Menu Menu_625
         Caption         =   "&NFe"
         Index           =   0
         Tag             =   "625|2"
      End
      Begin VB.Menu Menu_625
         Caption         =   "Inutiliza - NFe"
         Index           =   1
         Tag             =   "670|2"
      End
      Begin VB.Menu Menu_625
         Caption         =   "Notas - Emitidas"
         Index           =   2
         Tag             =   "716|2"
      End
      Begin VB.Menu Menu_625
         Caption         =   "&Ocultos"
         Index           =   3
         Tag             =   "629|2"
         Begin VB.Menu Menu_628
            Caption         =   "Cancela Nota Fiscal"
            Index           =   0
            Tag             =   "628|3"
         End
         Begin VB.Menu Menu_628
            Caption         =   "Carta de Correção"
            Index           =   1
            Tag             =   "630|3"
         End
         Begin VB.Menu Menu_628
            Caption         =   "Ativador"
            Index           =   2
            Tag             =   "633|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Movimento do Estoque"
      Index           =   6
      Tag             =   "638|1"
      Begin VB.Menu Menu_639
         Caption         =   "&Entradas - Estoque"
         Index           =   0
         Tag             =   "639|2"
      End
      Begin VB.Menu Menu_639
         Caption         =   "&Saidas - Estoque"
         Index           =   1
         Tag             =   "667|2"
      End
      Begin VB.Menu Menu_639
         Caption         =   "&Devolução de Compra"
         Index           =   2
         Tag             =   "657|2"
         Begin VB.Menu Menu_654
            Caption         =   "Devolução"
            Index           =   0
            Tag             =   "654|3"
         End
         Begin VB.Menu Menu_654
            Caption         =   "Cancela Devolução"
            Index           =   1
            Tag             =   "658|3"
         End
      End
      Begin VB.Menu Menu_639
         Caption         =   "&Cancela Movimento"
         Index           =   3
         Tag             =   "672|2"
         Begin VB.Menu Menu_673
            Caption         =   "Cancela - Entrada"
            Index           =   0
            Tag             =   "673|3"
         End
         Begin VB.Menu Menu_673
            Caption         =   "Cancela - Saida"
            Index           =   1
            Tag             =   "674|3"
         End
      End
      Begin VB.Menu Menu_639
         Caption         =   "&Ocultos"
         Index           =   4
         Tag             =   "641|2"
         Begin VB.Menu Menu_642
            Caption         =   "Importar NFe"
            Index           =   0
            Tag             =   "642|3"
         End
         Begin VB.Menu Menu_642
            Caption         =   "Devolução - Nota Fiscal"
            Index           =   1
            Tag             =   "664|3"
         End
         Begin VB.Menu Menu_642
            Caption         =   "Movimento - Nota Fiscal"
            Index           =   2
            Tag             =   "666|3"
         End
         Begin VB.Menu Menu_642
            Caption         =   "Outras Saidas - Nota Fiscal"
            Index           =   3
            Tag             =   "671|3"
         End
         Begin VB.Menu Menu_642
            Caption         =   "&Devolução de Compra"
            Index           =   4
            Tag             =   "655|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Ordem de Serviço"
      Index           =   7
      Tag             =   "679|1"
      Begin VB.Menu Menu_680
         Caption         =   "&O.S"
         Index           =   0
         Tag             =   "680|2"
      End
      Begin VB.Menu Menu_680
         Caption         =   "&Baixa - Casco"
         Index           =   1
         Tag             =   "726|2"
      End
      Begin VB.Menu Menu_680
         Caption         =   "Dados Adicionais - O.S"
         Index           =   2
         Tag             =   "732|2"
      End
      Begin VB.Menu Menu_680
         Caption         =   "Dados Adicionais - Virabrequim"
         Index           =   3
         Tag             =   "738|2"
      End
      Begin VB.Menu Menu_680
         Caption         =   "&Relatorios"
         Index           =   4
         Tag             =   "691|2"
         Begin VB.Menu Menu_725
            Caption         =   "Virabrequim"
            Index           =   0
            Tag             =   "725|3"
         End
         Begin VB.Menu Menu_725
            Caption         =   "O.S - Emitidas"
            Index           =   1
            Tag             =   "692|3"
         End
         Begin VB.Menu Menu_725
            Caption         =   "Cascos"
            Index           =   2
            Tag             =   "727|3"
         End
         Begin VB.Menu Menu_725
            Caption         =   "O.S - Emitadas (Exportar Excel)"
            Index           =   3
            Tag             =   "713|3"
         End
         Begin VB.Menu Menu_725
            Caption         =   "Geral - Exportar Excel"
            Index           =   4
            Tag             =   "714|3"
         End
      End
      Begin VB.Menu Menu_680
         Caption         =   "&Ocultos"
         Index           =   5
         Tag             =   "687|2"
         Begin VB.Menu Menu_688
            Caption         =   "Imprimir"
            Index           =   0
            Tag             =   "688|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Financeiro"
      Index           =   8
      Tag             =   "646|1"
      Begin VB.Menu Menu_649
         Caption         =   "&Contas a Pagar"
         Index           =   0
         Tag             =   "649|2"
         Begin VB.Menu Menu_694
            Caption         =   "&Entrada Contas"
            Index           =   0
            Tag             =   "694|3"
         End
         Begin VB.Menu Menu_694
            Caption         =   "&Baixa Contas"
            Index           =   1
            Tag             =   "650|3"
         End
         Begin VB.Menu Menu_694
            Caption         =   "&Relatorios"
            Index           =   2
            Tag             =   "647|3"
            Begin VB.Menu Menu_648
               Caption         =   "Contas a Pagar"
               Index           =   0
               Tag             =   "648|4"
            End
            Begin VB.Menu Menu_648
               Caption         =   "Contas Pagas"
               Index           =   1
               Tag             =   "651|4"
            End
            Begin VB.Menu Menu_648
               Caption         =   "Comissão dos Vendedores"
               Index           =   2
               Tag             =   "739|4"
            End
         End
      End
      Begin VB.Menu Menu_649
         Caption         =   "&Contas a Receber"
         Index           =   1
         Tag             =   "660|2"
         Begin VB.Menu Menu_736
            Caption         =   "&Entrada Contas"
            Index           =   0
            Tag             =   "736|3"
         End
         Begin VB.Menu Menu_736
            Caption         =   "&Baixa Contas"
            Index           =   1
            Tag             =   "665|3"
         End
         Begin VB.Menu Menu_736
            Caption         =   "Manutenção - Contas"
            Index           =   2
            Tag             =   "698|3"
         End
         Begin VB.Menu Menu_736
            Caption         =   "Manutenção Contas - Lote"
            Index           =   3
            Tag             =   "699|3"
         End
         Begin VB.Menu Menu_736
            Caption         =   "&Relatorios"
            Index           =   4
            Tag             =   "661|3"
            Begin VB.Menu Menu_662
               Caption         =   "&Contas a Receber"
               Index           =   0
               Tag             =   "662|4"
            End
            Begin VB.Menu Menu_662
               Caption         =   "Contas - Recebidas"
               Index           =   1
               Tag             =   "690|4"
            End
            Begin VB.Menu Menu_662
               Caption         =   "Baixa - Contas á Receber"
               Index           =   2
               Tag             =   "722|4"
            End
         End
      End
      Begin VB.Menu Menu_649
         Caption         =   "&Cheque - Pré"
         Index           =   2
         Tag             =   "718|2"
         Begin VB.Menu Menu_719
            Caption         =   "&Cheques"
            Index           =   0
            Tag             =   "719|3"
         End
         Begin VB.Menu Menu_719
            Caption         =   "&Relatórios"
            Index           =   1
            Tag             =   "720|3"
            Begin VB.Menu Menu_721
               Caption         =   "Relatorio de Cheques de Terceiros"
               Index           =   0
               Tag             =   "721|4"
            End
         End
      End
      Begin VB.Menu Menu_649
         Caption         =   "&Ocultos"
         Index           =   3
         Tag             =   "696|2"
         Begin VB.Menu Menu_697
            Caption         =   "&Manutenção - Receber"
            Index           =   0
            Tag             =   "697|3"
         End
         Begin VB.Menu Menu_697
            Caption         =   "Comissão dos Vendedores"
            Index           =   1
            Tag             =   "740|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Consultas"
      Index           =   9
      Tag             =   "631|1"
      Begin VB.Menu Menu_632
         Caption         =   "Nota Fiscal"
         Index           =   0
         Tag             =   "632|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Geral"
         Index           =   1
         Tag             =   "634|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Entradas - Estoque"
         Index           =   2
         Tag             =   "675|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Saidas - Estoque"
         Index           =   3
         Tag             =   "676|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Devolução de Compra"
         Index           =   4
         Tag             =   "663|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Cheques"
         Index           =   5
         Tag             =   "723|2"
      End
      Begin VB.Menu Menu_632
         Caption         =   "Entrada - Contas"
         Index           =   6
         Tag             =   "724|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Ferramentas"
      Index           =   10
      Tag             =   "636|1"
      Begin VB.Menu Menu_637
         Caption         =   "Copia de Segurança"
         Index           =   0
         Tag             =   "637|2"
      End
      Begin VB.Menu Menu_637
         Caption         =   "&Ocultos"
         Index           =   1
         Tag             =   "644|2"
         Begin VB.Menu Menu_645
            Caption         =   "Senha Administrador"
            Index           =   0
            Tag             =   "645|3"
         End
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "E&xibir"
      Index           =   11
      Tag             =   "598|1"
      Begin VB.Menu Menu_Barra
         Caption         =   "&Barra de ferramentas"
         Tag             =   "599|2"
      End
      Begin VB.Menu Menu_Ajuda
         Caption         =   "&Ajuda ativa"
         Tag             =   "600|2"
      End
      Begin VB.Menu Menu_Browse
         Caption         =   "Em &grade"
         Shortcut        =   {F9}
         Tag             =   "601|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "&Janela"
      Index           =   12
      Windowlist      =   -1
      Tag             =   "602|1"
      Begin VB.Menu Menu_Cascata
         Caption         =   "&Cascata"
         Tag             =   "603|2"
      End
      Begin VB.Menu Menu_Vert
         Caption         =   "Lado a lado &vertical"
         Tag             =   "604|2"
      End
      Begin VB.Menu Menu_Horiz
         Caption         =   "Lado a lado &horizontal"
         Tag             =   "605|2"
      End
      Begin VB.Menu Menu_Organiza
         Caption         =   "&Organizar ícones"
         Tag             =   "606|2"
      End
   End
   Begin VB.Menu Menu_570
      Caption         =   "A&juda"
      Index           =   13
      Tag             =   "607|1"
      Begin VB.Menu Menu_Conteudo
         Caption         =   "&Conteúdo"
         Shortcut        =   {F1}
         Tag             =   "608|2"
      End
      Begin VB.Menu Menu_N4
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Sobre
         Caption         =   "Sobre&..."
         Tag             =   "610|2"
      End
   End
         Begin VB.Menu Menu_PropGrade
                  Caption         =   "PropGrades"
                  Visible         =   0                           'False
                  Begin VB.Menu Menu_Grd_ApgCol
                           Caption         =   "&Apagar colunas"
                  End
                  Begin VB.Menu Menu_Grd_Filtra
                           Caption         =   "&Filtrar"
                  End
                  Begin VB.Menu Menu_Grd_Graf
                           Caption         =   "&Grafar"
                  End
                  Begin VB.Menu Menu_Grd_Imp
                           Caption         =   "&Imprimir"
                  End
                  Begin VB.Menu Menu_Grd_Procura
                           Caption         =   "&Procurar"
                  End
                  Begin VB.Menu Menu_Grd_Constroi
                           Caption         =   "&Remontar a grade"
                  End
         End
         Begin VB.Menu M_Oper 
                  Caption         =   "MOper"
                  Visible         =   0                           'False
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "Like"
                           Index           =   0
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "="
                           Index           =   1
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "<>"
                           Index           =   2
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "Vazio"
                           Index           =   3
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "<"
                           Index           =   4
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "<="
                           Index           =   5
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   ">"
                           Index           =   6
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   ">="
                           Index           =   7
                  End
                  Begin VB.Menu Menu_Operador 
                           Caption         =   "Entre"
                           Index           =   8
                  End
         End
End
Attribute VB_Name = "mdiNFE"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: mdiNFE
'* Função....: Janela principal - MDI
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 14/03/2022 08:09:17
'* * * * * * *

Option Explicit                                   'requer variáveis explicitamente declaradas
DefInt A-Z                                        'inteiras, por default

'declara variáveis públicas - todos os forms possuem essas variáveis
Public vgSituacao As Integer                      'situação de cada form child
Public vgCaracteristica As Integer                'característica dos forms
Public vgTipo As Integer                          'tipo de cada form child
Public vgFormID As Long                           'identificacao do form
Public vgEmBrowse As Integer                      'flag: se true, dados mostrados em browse
Public vgPriVez As Integer
Public vgTb As GRecordSet                         'referência a recordset dos data controls
Public vgTooltips As New cTooltips
Dim vgTile As New cFormBG

'apaga colunas selecionadas na grade
Public Sub ApagaColunas()
   If ActiveForm.grdBrowse.ColumnSelected.Count < 1 Then           'se não tem coluna selecionada,
      Beep                                        'avisa,
      MsgBox LoadGasString(1535), vbExclamation, vgAtencao$ 'mostra mensagem e retorna
   Else                                           'senão,
      With ActiveForm.grdBrowse                   'com a grade...
         .DeleteCols                              'invoca método para apagar colunas selecionadas
      End With
   End If
End Sub

'mostra fórmulas na janela
Private Sub MostraFormulas()
   On Error Resume Next                           'previne erros...
   Label(0).Caption = Empresa![Nome fantasia]
   If Err Then Err = 0: Label(0).Caption = ""
   Label(1).Caption = Ucase(vgPWGrupo)
   If Err Then Err = 0: Label(1).Caption = ""
   Label(2).Caption = UCase(vgPWUsuario)
   If Err Then Err = 0: Label(2).Caption = ""
   Label(4).Caption = UCase(UsuarioPC)
   If Err Then Err = 0: Label(4).Caption = ""
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

' procedure para reparar o banco de dados
Public Sub ReparaBD()
   Dim x As String
   On Error GoTo DeuErro
   Beep                                           'avisa que pode demorar...
   If MsgBox(LoadGasString(36), vbExclamation + vbYesNo + _
      vbDefaultButton2, LoadGasString(35)) = vbYes Then
      MousePointer = vbHourglass                  'ponteiro do mouse = ampulheta

      If Len(x$) > 0 Then
         Beep
         MsgBox x$, vbCritical, vgAtencao$        'mostra mensagem de erro e
         FinalizaAplicacao
         End
      End If

      FechaBancoDeDados                           'fecha tudo e abre novamente

      If Not AbreBancoDeDados Then                'tenta abrir o banco de dados...
         Beep                                     'se não foi possível,
         MsgBox LoadGasString(1130), vbCritical, vgAtencao$ 'mostra mensagem e
         End                                                'sai da aplicação...
      End If

      MousePointer = vbDefault                              'deu tudo certo, acerta ponteiro do mouse
      Beep
      MsgBox LoadGasString(37), vbInformation, LoadGasString(38) 'informa reparado com sucesso!
   End If
   Exit Sub

DeuErro:                                          'erros da reparação caem aqui...
   If Err = 3356 Or Err = 3196 Then               'DB está bloqueado por outros usuários,
      x$ = LoadGasString(130)                     'monta mensagem
   Else                                           'senão,
      x$ = LoadGasString(120) & " (" & Err & ")"  'outro tipo de erro - junta seu número
   End If
   Beep                                           'apita e
   MsgBox x$, vbExclamation, vgAtencao$           'mostra mensagem
   If Not AbreBancoDeDados Then                   'tenta abrir o banco de dados...
      Beep                                        'se não foi possível,
      MsgBox LoadGasString(1130), vbCritical, vgAtencao$ 'mostra mensagem e
      End                                                'sai da aplicação...
   End If
   MousePointer = vbDefault                              'BD foi aberto, acerta o mouse
End Sub

'esta rotina é chamada sempre que o usuário clicar no botão de imprimir ou na
'opção de menu correspondente - esta rotina atua diferentemente para cada tipo
'de form que esteja ativo no momento
Public Sub ImprimeForm()
   Dim vgSz As Integer
   Dim x As Integer
   Dim vgTit As String
   Dim i As Integer
   Dim vgNTbs As String, vgNt As String
   Dim vgBanco As GDataBase, vgTabela As GTable, vgInd As GIndex, vgCp As GColumn

   Select Case ActiveForm.vgTipo                            'verifica o tipo de form
      Case TP_TABELA, TP_PARAMETROS                         'SE FORM DE DADOS...
         If ActiveForm.vgEmBrowse Then                      'se na grade - vamos imprimir colunas
            Set ActiveForm.vgFrmImpCons.vgFrmOrigem = ActiveForm 'cria instância do form ativo
            Load ActiveForm.vgFrmImpCons
            ActiveForm.vgFrmImpCons.Show
            GoTo FimDaSub
         End If
         Beep                                     've se quer imprimir o form (hard copy)
         If MsgBox(LoadGasString(85) + UCase$(ActiveForm.Caption) + "?", _
          vbQuestion + vbYesNo, vgAtencao$) = vbYes Then
            GeraLogAcao LoadGasString(290) + " " + ActiveForm.Caption
            On Error Resume Next
            For i = 0 To ActiveForm.Grid.Count - 1 'vamos correr todos os grids
               ActiveForm.Grid(i).FreezeListView False 'congelar a exibição do grid
            Next
            On Error GoTo ErroImp                      'prepara para pegar um erro
            ActiveForm.PrintForm                       'invoca o método printform do VB
            On Error Resume Next
            For i = 0 To ActiveForm.Grid.Count - 1     'vamos correr todos os grids
	 
               ActiveForm.Grid(i).RestoreListView True 'descongelar a exibição do grid
            Next
         End If
         GoTo FimDaSub                                 'cai fora...

ErroImp:                                               'erro ao imprimir
         Beep                                          'beepa e
         MsgBox LoadGasString(90), vbExclamation, vgAtencao$ 'dá mensagem
         Resume Next                                         'continua a partir do erro...

      Case TP_GRAFICO                                        'SE FORM DE GRÁFICO...
         Beep                                                've se quer imprimir o gráfico
         If MsgBox(LoadGasString(85) + UCase$(ActiveForm.Caption) + "?", _
          vbQuestion + vbYesNo, vgAtencao$) = vbYes Then
            GeraLogAcao LoadGasString(291) + " " + ActiveForm.Caption
            ActiveForm.gra1.PrintStyle = 1                   'impressão colorida...
            ActiveForm.gra1.DrawMode = 5                     'imprime o grafico
         End If

      Case TP_BROWSE                                         'SE FORM DE CONSULTAS...
         Set ActiveForm.vgFrmImpCons.vgFrmOrigem = ActiveForm 'cria instância do form ativo
         Load ActiveForm.vgFrmImpCons
         ActiveForm.vgFrmImpCons.Show

      Case TP_INFODB                                          'SE FORM DE INFORMAÇÔES BD
         GeraLogAcao LoadGasString(290) + " " + ActiveForm.Caption
         Screen.MousePointer = vbHourglass                    'mouse = ampulheta
         vgSz = Printer.FontSize                              'salva tamanho da fonte da impressora
         For Each vgBanco In vgDb                             'corre todos os bancos da aplicação
            Printer.FontBold = True                           'liga negrito
            GoSub TestaSalto                                  'testa fim de página
            GoSub Cabecalho                                   'imprime o cabeçalho,
            GoSub TestaSalto                                  'testa fim de página
            For Each vgTabela In vgBanco.Tables               'corre todas as tabelas do banco atual
               Printer.FontBold = True                        'liga negrito
               GoSub TestaSalto                               'testa fim de página
               Printer.Print LoadGasString(95) + vgTabela.Name 'imprime Tabela: XXXXXXXX
               Printer.FontBold = False                        'desliga negrito
               GoSub TestaSalto                                'testa fim de página
               Printer.Print                                   'uma linha em branco
               GoSub TestaSalto                                'testa fim de página
               Printer.Print LoadGasString(100);               'imprime tít dos índices
               For Each vgInd In vgTabela.Indexes              'corre todos os índices da tabela atual
                  GoSub TestaSalto                             'testa fim de página
                  Printer.Print Tab(12); vgInd.Name            'imprime nome do índice
               Next                                            'vai para o próximo índice da tabela atual
               GoSub TestaSalto                                'testa fim de página
               Printer.Print                                   'uma linha em branco
               GoSub TestaSalto                                'testa fim de página
               Printer.Print Tab(12); LoadGasString(105);      'imprime tít dos campos
               For Each vgCp In vgTabela.Columns               'corre todos os campos da tabela atual
                  GoSub TestaSalto                             'testa fim de página
                  Printer.Print Tab(20); vgCp.Name             'imprime nome do campo
               Next                                            'vai para o próximo campo da tabela atual
               GoSub TestaSalto                                'testa fim de página
               Printer.Print                                   'salta linha
               GoSub TestaSalto                                'testa fim de página
               Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY) 'traça uma linha
               GoSub TestaSalto                   'testa fim de página
               Printer.Print                      'salta linha
            Next                                  'vai para próxima tabela do banco atual
         Next                                     'vai para próximo banco da aplicação
         Printer.FontSize = vgSz                  'restaura tamanho da fonte
         Printer.EndDoc                           'termina o documento
         Screen.MousePointer = vbDefault          'restaura ponteiro do mouse
         GoTo FimDaSub                            'cai fora...

TestaSalto:                                       'trata final de página
         If Printer.CurrentY > (Printer.Height - (Printer.Height * 0.1)) Then 'se 10% faltando
            Printer.NewPage                       'pula página
            GoSub Cabecalho                       'e imprime cabeçalho
         End If
         Return                                   'volta pra quem chamou...

Cabecalho:                                        'imprime cabeçalho
         With Printer                             'com a impressora,
            .FontSize = 16                        'ajusta tamanho e
            .FontBold = True                      'corpo da fonte
            vgTit$ = LoadGasString(110)           'título do relatório
            x = .TextWidth(vgTit$)                'calcula quatos pontos gasta o título
            .CurrentX = (.Width - x) / 2          'e centraliza para imprimir
            Printer.Print vgTit$                  'imprime título
            .FontSize = 14                        'diminui a fonte
            x = .TextWidth(vgBanco.Name)          'calcula quatos pontos gasta o título
            .CurrentX = (.Width - x) / 2          'e centraliza para imprimir
            Printer.Print vgBanco.Name            'imprime nome do banco atual
            Printer.Print                         'salta uma linha
            .FontBold = False                     'desliga negrito
            Printer.Line (0, .CurrentY)-(.Width, .CurrentY + 60), , BF
            Printer.Print                         'salta linha
            .FontSize = 9.5                       'ajusta tamanho da fonte
         End With
         Return                                   'e retorna para quem chamou...

   End Select
   
FimDaSub:

CancelouConfig:                                   'se caiu aqui, cancelou configuração impressora

End Sub

'esta rotina é chamada sempre que o usuário solicitar filtragem de registros,
'atuando diferentemente para cada tipo de form que estiver ativo no momento
Public Sub FiltraRegistros(Optional AbreNaOrdem As Boolean, Optional vgOq As Integer)
   On Error GoTo DeuErro
   vgBotoesOK = False                                       'flag - botões não preparados
   vgFiltroAtual$ = ActiveForm.vgUltimoFiltro               'pega o último filtro feito
   vgFiltroAtualComTit$ = ActiveForm.vgUltimoFiltroComTit
   vgOrdemAtual$ = ActiveForm.vgUltimaOrdem                 'pega a última ordem feito
   vgOrdemAtualComTit$ = ActiveForm.vgUltimaOrdemComTit
   If vgoq = 0 Then
      Load frmFiltra                                        'carrega form de filtragem
      If ActiveForm.vgTipo = TP_TABELA Then
         frmFiltra.IniciaTabDoFiltro ActiveForm, ActiveForm.vgTb 'inicializa a tabela do filtro
      ElseIf ActiveForm.vgTipo = TP_BROWSE Then
         frmFiltra.IniciaTabDoFiltro ActiveForm, ActiveForm.grdBrowse.Recordset 'inicializa a tabela do filtro
      End If
   End If
   If AbreNaOrdem Then
      frmFiltra.tabFiltro.Tab = 1
   End If
   If vgOq = 0 Then frmFiltra.Show vbModal                                      'apresenta o form de filtragem para o usuário
   If ActiveForm.vgUltimoFiltro <> vgFiltroAtual$ Or _
      ActiveForm.vgUltimaOrdem <> vgOrdemAtual$ Or vgOq = 1 Then                'se o novo filtro/ordem feito pelo usuário for diferente
      Screen.MousePointer = vbHourglass                                         'mouse = ampulheta
      With ActiveForm                             'com o form ativo,
         .vgUltimoFiltro = vgFiltroAtual$         'atualiza a variável vgUltimoFiltro 
         .vgUltimoFiltroComTit = vgFiltroAtualComTit$ 'e vgUltimoFiltroComTit
         .vgUltimaOrdem = vgOrdemAtual$                        
         .vgUltimaOrdemComTit = vgOrdemAtualComTit$ 
         InicializaFiltro ActiveForm
         If .vgTipo = TP_TABELA And Not .vgEmBrowse Then 'é form de dados e não está em modo grade?
            .Reposition
         ElseIf .vgTipo = TP_TABELA And .vgEmBrowse Then 'é janeda de dados e está em modo grade?
            .grdBrowse.OpenRecordSet ActiveForm.vgSQL, CURSOR_TABLE, False, ActiveForm.vgUltimoFiltro, ActiveForm.vgUltimaOrdem   
         ElseIf .vgTipo = TP_BROWSE Then                 'é form de consulta?                            
            .grdBrowse.OpenRecordSet ActiveForm.vgSQL, CURSOR_TABLE, False, ActiveForm.vgFiltroAtual, ActiveForm.vgOrdemAtual
         End If
      End With
      Screen.MousePointer = vbDefault                    'restaura o ponteiro do mouse
   End If
   PrepBotoes ActiveForm, ACAO_NAVEGANDO                 'reajusta funcionalidade e aparencia dos botoes
   If ActiveForm.vgTipo = TP_TABELA Then
      mdiNFE.vgPriVez = True
      mdiNFE.vgPriVez = False
   End If 
   
   vgBotoesOK = True                                     'flag - os botões estão OK...
   Exit Sub

DeuErro:
   Beep
   CErr.NumErro = Err
   CErr.Mensagem = LoadGasString(76)
   CErr.Show
   Screen.MousePointer = vbDefault                       'restaura o ponteiro do mouse
End Sub

'esta rotina é chamada para excluir um registro de uma tabela
'qualquer de um form de dados
Public Function ExcluiRegistro() As Boolean
   Dim x As String, i As Integer, vgRegAgora As Variant, vgNVez As Integer   'dimensiona
   If Not ActiveForm.vgTemExclusao Then Exit Function                        'se não pode excluir, cai fora...
      If Not ActiveForm.ActiveControl Is Nothing Then
         If TypeOf ActiveForm.ActiveControl Is GListV Then
            If UCase$(ActiveForm.ActiveControl.Name) = "GRID" Then
               If ActiveForm.ActiveControl.SelectedRowsCount > 0 Then        'se tem linha selecionada
                  Beep
                  ActiveForm.SetFocus                                        'vamos excluir no grid
                  SendK vbKeyDelete
                  Exit Function
               End If
            End If
         End If
      End If
   vgNVez = 0                                     'contador de tentativas
   Beep                                           'fomm, fomm! VAMOS EXCLUIR
   If ActiveForm.vgCaracteristica = F_DADOS And ActiveForm.vgEmBrowse Then
      i = vbYes
   Else
      i = vbYesNo + vbQuestion + vbDefaultButton2 'pede confirmação...
      i = MsgBox(LoadGasString(50), i, vgAtencao$) 'mostrando esta mensagem
   End If
   If i = vbYes Then                               'se confirmou,
      On Error GoTo ErroDele                       'prepara para pegar o erro
      vgBotoesOK = False                           'desliga troca de icones nos botoes
      Screen.MousePointer = vbHourglass            'mouse = ampulheta
      ActiveForm.vgSituacao = ACAO_EXCLUINDO       'ajusta situação atual do form
      vgDb.BeginTrans
      With ActiveForm.vgTb                         'com a tabela do form ativo,
         vgRegAgora = ActiveForm.vgTb.Bookmark     'salva o registro atual
         x$ = ActiveForm.Executar(EXCLUSOES)       'então faz os seus inversos
         If Len(x$) > 0 Then                       'deu erro?
            GoTo ResumeErro
         End If
         GeraLog ActiveForm, ACAO_EXCLUINDO, -1, True
         .Delete                                   'agora sim, apaga o registro corrente
      End With
      GeraLog ActiveForm, ACAO_EXCLUINDO
      vgDb.CommitTrans
      ExcluiRegistro = True                        'ok, exclusão executada com sucesso
      On Error GoTo FimDaSub
      x$ = ActiveForm.Executar(APOS_EDICAO)        'executa ação após exclusão do registro
      If (ActiveForm.vgCaracteristica = F_DADOS) And Not ActiveForm.vgEmBrowse Then
         ActiveForm.vgSituacao = ACAO_NAVEGANDO    'ajusta situação atual do form
         MoveRegistro ActiveForm, REG_FORCAVOLTA   'move para o registro anterior...
      End If
   End If
   GoTo FimDaSub                                   'pula fora...

ErroDele:                                          'se caiu aqui, vamos ver os erros possíveis
   CErr.NumErro = Err
   If CErr.Bloqueado Then                          'se o reg esta bloqueado por outro usuario
      vgNVez = vgNVez + 1                          'vamos tentar
      If vgNVez <= 10 Then                         'por 10 vezes
         Delay 3                                   'a cada 3 segundo
         Resume                                    'a mesma operação
      End If
   End If

   'salva o erro atual
   x$ = Err.Source + "|" + CStr(Err.Number) + "|" + Err.Description
   If Err Then Resume ResumeErro
   
ResumeErro:
   On Error Resume Next
   ActiveForm.vgTb.CancelUpdate                    'cancela alteracao no recordset                 
   Err.Clear
   vgDb.RollBackTrans
   CErr.Mensagem = LoadGasString(123)
   'restaura o erro original
   Err.Source = Parse$(x$, "|")                    'onde o erro foi causado
   Err.Number = Val(Parse$(x$, "|" ))              'o número do erro e a sua descrição
   Err.Description = Parse$(x$, "|")
   If (ActiveForm.vgCaracteristica = F_DADOS) And Not ActiveForm.vgEmBrowse Then
      If CErr.NumErro = 3167 Or CErr.NumErro = -2147217885 Then 'registro foi apagado por outro usuário
         MoveRegistro ActiveForm, REG_FORCAVOLTA                'precisamos retroceder um registro
      Else
         ActiveForm.vgTb.Bookmark = vgRegAgora                  'e volta para o mesmo registro
      End If
   Else
      ActiveForm.vgTb.BookMark = vgRegAgora                     'e volta para o mesmo registro
   End If
   CErr.Show                                      'mostra o erro
   On Error Resume Next

FimDaSub:                                         'finaliza esta rotina...
   If Not ActiveForm.vgEmBrowse Then
      RemontaForm                                 'atualiza os forms de dados
      ActiveForm.SetFocus                         'volta o foco para o form atual
      PrepBotoes ActiveForm, ACAO_NAVEGANDO       'ajusta funcionalidade dos botoes
   End If
   Screen.MousePointer = vbDefault                'restaura ponteiro do mouse
End Function

'inicia inclusão de registro
Public Sub IncluiRegistro()
   On Error GoTo DeuErro
   vgBotoesOK = False                             'botoeira em estado de edição
   PrepBotoes ActiveForm, ACAO_INCLUINDO          'troca icones dos botoes
   ActiveForm.Executar INICIALIZACOES             'coloca valor inicial para os campos
   ActiveForm.Reposition                          'mostra os campos
   FocoNoPriControle ActiveForm                   'coloca foco no 1o. campo do form
   Exit Sub
DeuErro:
   If mdiNFE.ActiveForm Is Nothing Then
      PrepBotoes Me, ACAO_NAVEGANDO               'ajusta barra de farramenta
   End If
End Sub

' rotina para manipular os cancelamentos de alterações nas tabelas
Public Sub CancelaAlteracoes()
   Dim vgIsGrid As Boolean
   On Error GoTo DeuErro
   If ActiveForm.vgSituacao < 0 Or ActiveForm.vgSituacao = ACAO_NAVEGANDO Then  'editando o grid não faz nada
      ActiveForm.Painel(0).Enabled = False
      ActiveForm.Painel(0).Enabled = True
      If ActiveForm.vgSituacao < 0 Then                                         'se estava em edição ou inclusão no GRID
         ActiveForm.vgSituacao = -99                                            'seta vgSituacao para forçar o cancelamento
      End If
      On Error Resume Next                        'pode ser que este form nao tenha esta procedure
      ActiveForm.Grid_LostFocus ActiveForm.ActiveControl.Index 'dispara o LostFocus do grid para cancelar a inclusão/alteração
      ActiveForm.ActiveControl.CancelEdit                      'cancela a inclusão/alteração no grid
      If Err Then Err.Clear                                    'ignora possíveis erros (se não tem grid)
      Exit Sub                                                 'e cai fora
   ElseIf ActiveForm.vgTipo = TP_BROWSE Then
      If ActiveForm.grdBrowse.Status <> ACAO_NAVEGANDO Then
         ActiveForm.grdBrowse.CancelEdit
      End If
      Exit Sub
   End If
   On Error GoTo DeuErro                                       'prepara para captar erros
   vgIsGrid = False
   If (ActiveForm.vgCaracteristica = F_DADOS) Then
      vgIsGrid = ActiveForm.vgEmBrowse
   End If
   If vgIsGrid Then
      ActiveForm.grdBrowse.CancelEdit
   Else
      ActiveForm.vgSituacao = ACAO_NAVEGANDO                   'ajusta situação do form
      PrepBotoes ActiveForm, ACAO_NAVEGANDO                    'acerta icones dos botoes
   End If
   ActiveForm.Reposition                                       'atualiza controles com dados do registro
   ActiveForm.SetFocus                                         'volta o foco para o form
   Exit Sub                                       'e sai...

DeuErro:                                          'se caiu aqui, deu erro
   If mdiNFE.ActiveForm Is Nothing Then
      PrepBotoes Me, ACAO_NAVEGANDO               'ajusta barra de farramenta
   Else
      MoveRegistro ActiveForm, REG_FORCAVOLTA     'vamos somente voltar um registro e
   End If
   Resume Next                                    'continuar
End Sub

'rotina para gravação de dados - invocada quando o botão de gravar é clicado ou
'quando a opção de menu correspondente é selecionada
Public Sub SalvaDados(Optional ByRef vgColumn As Integer)
   Dim vgRegAgora As Variant, x As String, vgIniciouTrans As Integer, vgEtapa As Integer, vgNTent As Integer, vgResume As Boolean, vgAdded As Boolean  'dimensiona
   Dim vgNVez As Integer                          'contador de tentativas
   Dim dlgArquivo As cCommonDialog
   vgResume = True                                'se pode fazer resume no tratamento de erros
   On Error GoTo DeuErro
   If ActiveForm.vgTipo <> TP_GRAFICO And (ActiveForm.vgSituacao < 0 Or ActiveForm.vgSituacao = ACAO_NAVEGANDO) Then 'editando o grid não faz nada
      If ActiveForm.vgSituacao < 0 Then
         ActiveForm.SaveGrids
      End If
      Exit Sub
   ElseIf ActiveForm.vgTipo = TP_BROWSE Then
      If ActiveForm.grdBrowse.Status <> ACAO_NAVEGANDO Then
         ActiveForm.grdBrowse.SaveEdit
      End If
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass              'mouse = ampulheta
   vgNVez = 0                                     'inicia contador de tentativas
   vgIniciouTrans = False                         'flag - iniciou transação
   vgEtapa = 0                                    'etapa do processo de gravação
   vgNTent = 0                                    'número de vezes que tentou executar gravação e ocorreu o erro -2147217864
   On Error GoTo SaiDaSUb                         'prepara para pegar erros
   Select Case ActiveForm.vgTipo                  'verifica o tipo de form
      Case TP_GRAFICO                             'SE FOR TIPO GRÁFICO...
         Set dlgArquivo = New cCommonDialog
         With dlgArquivo                          'vamos apresentar diálogo para salvar
            .CancelError = True                   'vamos testar se o usuário cancelou
            .InitDir = vgDirEXE$                  'inicializa o diretório apresentado
            .DialogTitle = LoadGasString(60)      'e o título do diálogo...
            .Filter = "Bitmaps (*.bmp)|*.bmp|Metafiles (*.wmf)|*.wmf"
            .FilterIndex = 0                      'vamos oferecer *.bmp primeiro
            .Flags = OFN_CREATEPROMPT             'se não existir, poderemos criar?
            .ShowSave                             'mostra o diálogo de gravação
            ActiveForm.gra1.ImageFile = .filename 'ajusta o nome do arquivo
            If InStr(UCase$(.filename), ".WMF") Then 'se tem extensão .WMF
               ActiveForm.gra1.DrawMode = 6          'grava o wmf
            Else                                     'senão
               ActiveForm.gra1.DrawMode = 3          'cria bmp em memoria
               ActiveForm.gra1.DrawMode = 6          'e grava em disco...
            End If
         End With
         Set dlgArquivo = Nothing
      Case Else                                      'SE FORM DIFERENTE DE GRÁFICO...
         vgBotoesOK = False                          'reseta flag
         vgRegAgora = ""
         If ActiveForm.vgTb.RecordCount > 0 Then
            If ActiveForm.vgTb.EOF = False And ActiveForm.vgTb.BOF = False And ActiveForm.vgTb.NoMatch = False Then
               vgRegAgora = ActiveForm.vgTb.Bookmark 'salva o registro atual
            End If
         End If
         x$ = ActiveForm.Executar(VALIDACOES, vgColumn) 'faz validação dos campos
         If Len(x$) > 0 Then                            'EPA!... tem campo errado
            If InStr(x$, "|") > 0 Then
               Err.Source = Parse$(x$, "|")
               Err.Number = Val(Parse$(x$, "|"))        'deu erro
               Err.Description = Parse$(x$, "|")
               CErr.NumErro = Err.Number
               CErr.Show
            Else
               Beep                                     'manda aviso sonoro
               MsgBox x$, vbCritical, vgAtencao$        'mostra mensagem do erro
            End If
            ActiveForm.SetFocus                         'volta o foco para o form
            Screen.MousePointer = vbDefault             'reseta o pointer do mouse
            Exit Sub                                    'e sai para continuar a edição ou inclusão
         End If                                         'OK. os campos estão certos
         On Error GoTo ErroAoGravar                     'vamos cercar erros...
         vgDb.BeginTrans
         vgIniciouTrans = True                          'iniciou transação = true
      VoltaGravacao:
         If ActiveForm.vgSituacao = ACAO_EDITANDO Then  'se estava editando registros
            x$ = ActiveForm.Executar(TESTA_VAL_RS)      'testa se o recordset foi alterado por outro usuário
            If Len(x$) > 0 Then                         'deu erro?
               vgNVez = 100                             'ajusta contador de tentativas
               Err.Source = Parse$(x$, "|")             'força erro que deu...
               Err.Number = Val(Parse$(x$, "|"))
               Err.Description = Parse$(x$, "|")
               vgResume = False
               GoTo ErroAoGravar
            End If
            ActiveForm.vgTb.Edit                        'invoca a edição da tabela
            If vgEtapa < 1 Then
               x$ = ActiveForm.Executar(PROCESSOS_INVERSOS) 'executa os processos inversos
               If Len(x$) > 0 Then                          'deu erro?
                  vgNVez = 100                              'ajusta contador de tentativas
                  Err.Source = Parse$(x$, "|")              'força erro que deu...
                  Err.Number = Val(Parse$(x$, "|"))
                  Err.Description = Parse$(x$, "|")
                  vgResume = False
                  GoTo ErroAoGravar
               Else
                  vgEtapa = vgEtapa + 1
               End If
            End If
         Else
            If vgEtapa < 1 Then
               ActiveForm.vgTb.AddNew                       'cria registro em branco
               vgEtapa = vgEtapa + 1  
            End If
         End If
         GeraLog ActiveForm, ActiveForm.vgSituacao, -1, True
         If vgEtapa < 2 Then
            x$ = ActiveForm.Executar(POE_NO_ARQUIVO)        'atualiza registro com os controles
            If Len(x$) > 0 Then                             'deu erro?
               vgNVez = 100                                 'ajusta contador de tentativas
               Err.Source = Parse$(x$, "|")                 'força erro que deu...
               Err.Number = Val(Parse$(x$, "|"))            'força erro que deu...
               Err.Description = Parse$(x$, "|")
               vgResume = False
               GoTo ErroAoGravar
            End If
            x$ = ActiveForm.Executar(PROCESSOS_ANTES)       'executa os processos diretos definidos para antes do update
            If Len(x$) > 0 Then                             'deu erro?
               vgNVez = 100                                 'ajusta contador de tentativas
               Err.Source = Parse$(x$, "|")                 'força erro que deu...
               Err.Number = Val(Parse$(x$, "|"))            'força erro que deu...
               Err.Description = Parse$(x$, "|")
               vgResume = False
               GoTo ErroAoGravar
            End If
            ActiveForm.vgTb.Update                          'deu certo - atualiza
            vgRegAgora = ActiveForm.vgTb.BookMark           'atualiza para o novo registro
            vgEtapa = vgEtapa + 1
         End If
         If vgEtapa < 3 Then
            x$ = ActiveForm.Executar(PROCESSOS_DIRETOS)     'executa os processos diretos
            If Len(x$) > 0 Then                             'deu erro?
            vgNVez = 100                                    'ajusta contador de tentativas
               Err.Source = Parse$(x$, "|")                 'força erro que deu...
               Err.Number = Val(Parse$(x$, "|"))            'força erro que deu...
               Err.Description = Parse$(x$, "|")
               vgResume = False
               GoTo ErroAoGravar
            Else
               vgEtapa = vgEtapa + 1
            End If
         End If
         GeraLog ActiveForm, ActiveForm.vgSituacao
         vgDb.CommitTrans
         vgIniciouTrans = False                             'reseta a flag
         If ActiveForm.vgTipo = TP_PARAMETROS Then
            Empresa.Requery
            Parametros_da_nfe.Requery
            Parametros_da_OS.Requery
         End If
         On Error GoTo DeuErro                              'vamos cercar erros...
         x$ = ActiveForm.Executar(APOS_EDICAO)              'executa acao após inclusão/Alteração
         RemontaForm                                        'atualiza forms de dados
         ActiveForm.vgSituacao = ACAO_NAVEGANDO             'situacao normal, força método reposition
         If ActiveForm.vgTemBrowse Then
            If ActiveForm.vgEmBrowse Then
               ActiveForm.grdBrowse.CancelEdit
            End If
         End If
         ActiveForm.Reposition True, Not vgAdded            'atualiza controles com dados do registro
         AjustaRolagem ActiveForm
         If Not ActiveForm.vgEmBrowse Then
            PrepBotoes ActiveForm, ACAO_NAVEGANDO
         Else
            ActiveForm.grdBrowse.SetFocus
         End If
         If botInclui.Enabled And panTools.Visible Then     'se toolbar visível e se botão de inclusão habilitado
            If botInclui.ButtonStyle = 0 Then
               botInclui.SetFocus                           'e coloca foco nele
            Else
               ActiveForm.SetFocus
            End If
         End If
   End Select
   GoTo SaiDaSub                                            'fora daqui

DeuErro:                                          'deu um erro
   Resume Next                                    'continuar...

ErroAoGravar:
   CErr.NumErro = Err
   If CErr.Bloqueado Then                         'se o reg esta bloqueado por outro usuario
      vgNVez = vgNVez + 1                         'vamos tentar
      If vgNVez <= 10 Then                        'por 10 vezes
         Delay 3                                  'a cada 3 segundo
         Resume                                   'a mesma operação
      End If
   End If
   If vgIniciouTrans Then                         'se iniciou transação
      If Err = -2147217864 And vgNTent < 10 Then
         Err.Clear
         If vgEtapa < 3 Then
            ActiveForm.vgTb.CancelUpdate          'cancela alteracao no recordset                 'cancela a operacao do recordset
         End If
         ActiveForm.vgTb.CloseRecordset
         InicializaFiltro ActiveForm
         On Error Resume Next
         ActiveForm.vgTb.BookMark = vgRegAgora
         Err.Clear
         On Error GoTo ErroAoGravar
         vgNTent = vgNTent + 1
         GoTo VoltaGravacao
      ElseIf Err <> 20 And Err <> 3600 And Err <> 3601 Then
         x$= Err.Source & "|" & Err.Number & "|" & Err.Description
         If vgResume Then
            Resume ResetaErro
         End If
      End If

ResetaErro:
      On Error Resume Next
      If vgEtapa < 3 Then
         ActiveForm.vgTb.CancelUpdate             'cancela alteracao no recordset                 'cancela a operacao do recordset
      End If
      vgDb.RollBackTrans
      If Len(CStr(vgRegAgora)) > 0 Then
         ActiveForm.vgTb.Bookmark = vgRegAgora    'volta para o registro
      ElseIf ActiveForm.vgTb.RecordCount > 0 Then 'forca inicio de arquivo
         ActiveForm.vgTb.MoveFirst
         If ActiveForm.vgTb.BOF = False Then      'não sabemos o porquê, mas 
            ActiveForm.vgTb.MovePrevious          'algumas vezes já está posicionado em, BOF
         End If
      End If
      Err.Source = Parse$(x$, "|")
      Err.Number = Val(Parse$(x$, "|"))
      Err.Description = Parse$(x$, "|")
   End If
   If Not ActiveForm.vgEmBrowse Then
      FocoNoPriControle ActiveForm
   End If
   CErr.Show
   If CErr.Fatal Then CancelaAlteracoes           'cancela inclusao/altercao
   ActiveForm.SetFocus                            'volta o foco para o form

SaiDaSub:                                         'saída comum da sub...
   If mdiNFE.ActiveForm Is Nothing Then
      PrepBotoes Me, ACAO_NAVEGANDO               'ajusta barra de farramenta
   End If
   Screen.MousePointer = vbDefault                'ponteiro do mouse = default
End Sub

'executa a remontagem de um form
Public Sub RemontaForm()
   Dim F As Form, vgEstadoBotoes As Integer
   
   vgEstadoBotoes = vgBotoesOK                    'salva o estado dos botões
   For Each F In Forms                            'e atualiza os registros
      If F.vgCaracteristica <> F_COMUM Then       'de todos os forms abertos
         vgBotoesOK = False                       'reseta flag
         On Error Resume Next                     'pode ocorrer algum erro...
         If F.vgSituacao = ACAO_NAVEGANDO And F.vgCaracteristica <> F_GRAFICO Then
            If F.vgTipo = TP_BROWSE Then
               F.grdBrowse.Refresh
            Else
               With F.vgTb                        'tabela principal do form
   

   VerificaReg:
                  If .RecordCount > 0 Then
                     If .EOF OR .BOF Then         'tem reg e nao esta visivel
                        .MoveFirst
                     End If
                     .Bookmark = .Bookmark
                     If F.vgCaracteristica = F_DADOS Then
                        F.Reposition
                     End If
                     If Err = 3167 Then           'se o registro já esta deletado
                        .MovePrevious             'move para o registro anterior
                        If .BOF Or Err > 0 Then   'se esta no inicio do arquivo ou deu erro
                           .MoveFirst             'vai para o primeiro do arquivo
                        End If
                        Err.Clear
                        GoTo VerificaReg
                     End If
                  Else
                     F.vgTb.Requery
                     F.Reposition
                  End If
               End With
            End If
         End If
      End If                                      'de inclusao/alteracao
   Next
   vgBotoesOK = vgEstadoBotoes                    'restaura a flag
End Sub

Private Sub botAnterior_Click()
   MoveRegistro ActiveForm, REG_ANTERIOR          'retrocede um registro por estímulo do usuário...
End Sub

Private Sub botBrowse_Click()
   If vgPriVez Then Exit Sub
   ActiveForm.vgEmBrowse = (Not ActiveForm.vgEmBrowse)  'inverte a flag e
   TrocaBrowse ActiveForm                               'troca forma de apresentar os dados
   PrepBotoes ActiveForm, ACAO_NAVEGANDO                'arruma botoes do mdi
   ActiveForm.SetFocus
End Sub

Private Sub botCancela_Click()
   CancelaAlteracoes                              'manda cancelar as alterações...
End Sub

Private Sub botCancela_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If ActiveForm.vgSituacao < 0 Then
      ActiveForm.vgSituacao = -99
   End If
End Sub

Private Sub Menu_Repeticao_Click()
   LigaDesligaRepeticao                           'repeticao de campos na inclusao 
End Sub

'lega/desliga a repeticao de campos na inclusao 
Public Sub LigaDesligaRepeticao()           
   On Error Resume Next
   ActiveForm.LigaDesligaRepeticao
End Sub

Private Sub Menu_Ortog_Click()
   CorrigeOrtografia                              'corrigo ortografia
End Sub

Public Sub CorrigeOrtografia()
   Dim vgDic As Object, vgTexto As String

   On Error Resume Next
   Set vgDic = CreateObject("Word.Basic")         'tenta cria objeto word
   If Err Then                                    'se não consegui
      Beep                                        'avisa e
      MsgBox LoadGasString(155), vbExclamation, vgAtencao$
      Exit Sub                                    'sai
   End If
   On Error GoTo DeuErro                          'espera um erro
   mdiNFE.Enabled = False
   With vgDic
      .FileNew
      .Insert ActiveForm.ActiveControl.Text
      .ToolsSpelling
      .EditSelectAll
      vgTexto$ = .Selection()
      .FileExit 2
   End With
   Set vgDic = Nothing                            'retira o objeto da memória
   mdiNFE.Enabled = True

   If Right$(vgTexto$, 1) = vbCr Then             'se texto termida com Enter
      vgTexto$ = Left$(vgTexto$, Len(vgTexto$) - 1) 'vamos despresar
   End If
   Beep                                             'aviso que passou por aqui
   ActiveForm.ActiveControl.Text =  Substitui$(vgTexto$, vbCr, vbCrLf, SO_UM)
   Exit Sub

DeuErro:
   mdiNFE.Enabled = True
   Set vgDic = Nothing                              'retira o objeto da memória
   Err.Clear                                        'limpa o erro que deu
End Sub

Private Sub botExclui_Click()
   If (ActiveForm.vgCaracteristica = F_DADOS) And ActiveForm.vgEmBrowse Then
      If ActiveForm.grdBrowse.SelectedRowsCount > 0 Then        'se tem registro selecionado
         ActiveForm.grdBrowse.DeleteRecords
      Else
         Beep
         MsgBox LoadGasString(1536), vbExclamation, vgAtencao$  'mostra mensagem e retorna
      End If
   Else
      ExcluiRegistro                                            'manda excluir o registro
   End If
End Sub

Private Sub botSalva_Click()
   SalvaDados                                     'manda gravar os dados
End Sub

Private Sub botImprime_Click()
   ImprimeForm                                    'imprime o form, com tratamento
End Sub

Private Sub botInclui_Click()
   IncluiRegistro                                 'trata inclusão do registro
End Sub

Private Sub botPrimeiro_Click()
   MoveRegistro ActiveForm, REG_PRIMEIRO          'movimenta para o primeiro registro da tabela
End Sub

Private Sub botSeguinte_Click()
   MoveRegistro ActiveForm, REG_SEGUINTE          'invoca movimentação para o registro seguinte
End Sub

Private Sub botUltimo_Click()
   MoveRegistro ActiveForm, REG_ULTIMO            'movimenta para o último registro
End Sub

'liga e desliga a apresentação de ajuda ativa para a aplicação
Private Sub Menu_Ajuda_Click()
   Dim i As Integer
   vgAjudaAtiva = Not vgAjudaAtiva                'inverte a flag
   Menu_Ajuda.Checked = vgAjudaAtiva              'marca/desmarca opção no menu
   On Error Resume Next
   For i = 0 To Forms.Count - 1                   'corre todos os forms abertos
      Forms(i).vgTooltips.Enabled = vgAjudaAtiva  'trocando apresentação da ajuda
   Next 

   'grava status no arquivo .INI
   GravaNoIni vgNomeEstacao$ + "Geral", "Ajuda ativa", Str$(vgAjudaAtiva), vgNomeINI$
End Sub

Private Sub Menu_Cascata_Click()
   mdiNFE.Arrange vbCascade                       'janelas em cascata
End Sub

Private Sub Menu_Horiz_Click()
   mdiNFE.Arrange vbTileHorizontal                'lado a lado horizontal
End Sub

Private Sub Menu_Vert_Click()
   mdiNFE.Arrange vbTileVertical                  'lado a lado vertical
End Sub

Private Sub Menu_Organiza_Click()
   mdiNFE.Arrange vbArrangeIcons                  'organiza icones
End Sub

Private Sub Menu_Anterior_Click()
   MoveRegistro ActiveForm, REG_ANTERIOR          'movimenta para o registro anterior
End Sub

'quer apagar colunas
Private Sub Menu_ApagaCol_Click()
   ApagaColunas                                   'apaga as colunas
End Sub

'liga/desliga a apresentação da barra de ferramentas
Private Sub Menu_Barra_Click()
   vgBarraFerr = Not vgBarraFerr                  'inverte a flag
   PoeTiraBarra

   'grava status no arquivo .INI
   GravaNoIni vgNomeEstacao$ + "Geral", "Barra ferramentas", Str$(vgBarraFerr), vgNomeINI$
End Sub

Public Sub PoeTiraBarra()
   Menu_Barra.Checked = vgBarraFerr               'marca/desmarca no menu
   With panTools
      .Visible = vgBarraFerr                      'mostra/esconde a barra de ferramentas
      If vgBarraFerr Then                         'Ah! VB @#%&! - temos que "sacodir" a barra
         If .Align < 3 Then                       'para que ela não fique parada no meio do form
            .Height = .Height + 15                'em certos casos, causando efeito estético indesejável
            .Height = .Height - 15                'mexemos na altura se o align for topo ou fundo...
         Else
            .Width = .Width + 15                  'se o .align for esquerda ou direita
            .Width = .Width - 15                  'só podemos mexher na largura
         End If
      End If
   End With
End Sub

'controla a forma de apresentação dos dados no form (grade/tela)
Private Sub Menu_Browse_Click()
   ActiveForm.vgEmBrowse = (Not ActiveForm.vgEmBrowse)   'inverte a flag
   TrocaBrowse ActiveForm                                'muda a forma de apresentação
   PrepBotoes ActiveForm, ACAO_NAVEGANDO                 'ajusta a ação no form
End Sub

Private Sub Menu_Cancela_Click()
   CancelaAlteracoes                              'invoca o cancelamento de alterações
End Sub

Private Sub Menu_ConfGraf_Click()
   frmConfGraf.Show vbModal
End Sub

'configura impressora a utilizar
Private Sub Menu_ConfImp_Click()
   ConfigImpressora
End Sub

'apresenta o form de seleção de consultas existentes
Private Sub Menu_Consulta_Click()
   If frmSeleQueries.WindowState = vbMinimized Then   'se minimizado,
      frmSeleQueries.WindowState = vbNormal           'apenas abre
   Else                                               'senão
      frmSeleQueries.Show                             'carrega e mostra
   End If
End Sub

Private Sub Menu_Conteudo_Click()
   AtivaHelp                                      'mostra conteúdo da ajuda on-line
End Sub

'popmenu da grade - apaga colunas
Private Sub Menu_Grd_ApgCol_Click()
   ApagaColunas                                   'chama função para apagar colunas
End Sub

'popmenu da grade - filtra registro
Private Sub Menu_Grd_Filtra_Click()
   FiltraRegistros                                'chama rotina apropriada
End Sub

'popmenu da grade - monta gráfico da consulta
Private Sub Menu_Grd_Graf_Click()
   MontaGraficoConsulta ActiveForm                'chama rotina apropriada
End Sub

'popmenu da grade - impreme consulta
Private Sub Menu_Grd_Imp_Click()
   ImprimeForm                                    'chama rotina apropriada
End Sub

'popmenu da grade - procura registro
Private Sub Menu_Grd_Procura_Click()
   frmProcura.Show vbModal                        'chama rotina apropriada
End Sub

'popmenu da grade - reconstro consulta
Private Sub Menu_Grd_Constroi_Click()
   Dim i As Integer
   With ActiveForm.grdBrowse                      'então, com a grade,
      For i = 0 To .Cols - 1                      'torna todas as colunas
         .Columns(i).Visible = (InStr(.Columns(i).Caption, "~") = 0) 'visivel se não int~lan, cod~lan
      Next
      .SelStartCol = -1                                              'reseta a
      .SelEndCol = -1                                                'seleção de colunas
   End With
End Sub


Private Sub Menu_Exclui_Click()
   If (ActiveForm.vgCaracteristica = F_DADOS) And ActiveForm.vgEmBrowse Then
      If ActiveForm.grdBrowse.SelectedRowsCount > 0 Then        'se tem registro selecionado
         ActiveForm.grdBrowse.DeleteRecords
      Else
         Beep
         MsgBox LoadGasString(1536), vbExclamation, vgAtencao$  'mostra mensagem e retorna
      End If
   Else
      ExcluiRegistro                                            'manda excluir o registro
   End If
End Sub

Private Sub Menu_Filtra_Click()
   FiltraRegistros                                'invoca a filtragem de registros
End Sub

Private Sub Menu_Fim_Click()
   UnLoad Me                                      'descarrega este form que vos fala...
End Sub

Private Sub Menu_Grafa_Click()
   MontaGraficoConsulta ActiveForm                'chama rotina de montagem de gráfico
End Sub

Private Sub Menu_Salva_Click()
   SalvaDados                                     'invoca a gravação de dados
End Sub

Private Sub Menu_Imprime_Click()
   ImprimeForm                                    'invoca a impressão do form
End Sub

Private Sub Menu_Inclui_Click()
   IncluiRegistro                                 'manipula a incluisão do registro
End Sub

Private Sub Menu_Info_Click()
   If frmInfoDB.WindowState = 1 Then              'se minimizado,
      frmInfoDB.WindowState = 0                   'somente abre
   Else                                           'senao
      frmInfoDB.Show                              'carrega e mostra
   End If
End Sub

Private Sub Menu_Primeiro_Click()
   MoveRegistro ActiveForm, REG_PRIMEIRO          '"go top"...
End Sub

Private Sub Menu_Procura_Click()
   frmProcura.Show vbModal                        'apresenta o form para montar pesquisa à tabela
End Sub

Private Sub Menu_Repara_Click()
   ReparaBD                                       'manda reparar o BD
End Sub

Private Sub Menu_Seguinte_Click()
   MoveRegistro ActiveForm, REG_SEGUINTE          'avança um registro
End Sub

Private Sub Menu_Senhas_Click()
   frmSenhas.Show vbModal                         'apresenta módulo de manutenção de senhas...
End Sub

Private Sub Menu_Sobre_Click()
   frmSobre.Show vbModal                          'mostra os créditos...
End Sub

Private Sub Menu_Ultimo_Click()
   MoveRegistro ActiveForm, REG_ULTIMO            'vai para o último registro...
End Sub

Private Sub Menu_Operador_Click(Index As Integer)
   mdiNFE.ActiveForm.ClickOperador Index
End Sub

Private Sub MDIForm_Activate()
   MostraFormulas
   If mdiNFE.ActiveForm Is Nothing Then
      PrepBotoes Me, ACAO_NAVEGANDO               'ajusta barra de farramenta
   End If
End Sub

' carga do form principal da aplicação
Private Sub MDIForm_Load()
   Dim x As String                                'dimensiona
   Dim vgWinInfo As RECT

   If vgCriandoSenha Then Exit Sub

   vgCaracteristica = F_COMUM                     'ajusta característica,
   vgTipo = TP_COMUM                              'seu tipo
   vgFormID = 1                                   'identificacao do form
   vgSituacao = ACAO_NAVEGANDO                    'e a situação atual neste form

   'designa icones para os botoes da barra de ferramenta
   Set botInclui.Picture = LoadPicture(LoadGasPicture(28))
   Set botInclui.PictureDisabled = LoadPicture(LoadGasPicture(52))
   Set botInclui.PictureMouseOver = LoadPicture(LoadGasPicture(12))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(31))
   Set botCancela.PictureDisabled = LoadPicture(LoadGasPicture(51))
   Set botCancela.PictureMouseOver = LoadPicture(LoadGasPicture(11))
   Set botBrowse.Picture = LoadPicture(LoadGasPicture(83))
   Set botBrowse.PictureDisabled = LoadPicture(LoadGasPicture(82))
   Set botBrowse.PictureMouseOver = LoadPicture(LoadGasPicture(81))
   Set botPrimeiro.Picture = LoadPicture(LoadGasPicture(35))
   Set botPrimeiro.PictureDisabled = LoadPicture(LoadGasPicture(54))
   Set botPrimeiro.PictureMouseOver = LoadPicture(LoadGasPicture(14))
   Set botAnterior.Picture = LoadPicture(LoadGasPicture(36))
   Set botAnterior.PictureDisabled = LoadPicture(LoadGasPicture(55))
   Set botAnterior.PictureMouseOver = LoadPicture(LoadGasPicture(15))
   Set botSeguinte.Picture = LoadPicture(LoadGasPicture(37))
   Set botSeguinte.PictureDisabled = LoadPicture(LoadGasPicture(56))
   Set botSeguinte.PictureMouseOver = LoadPicture(LoadGasPicture(16))
   Set botUltimo.Picture = LoadPicture(LoadGasPicture(38))
   Set botUltimo.PictureDisabled = LoadPicture(LoadGasPicture(57))
   Set botUltimo.PictureMouseOver = LoadPicture(LoadGasPicture(17))
   Set botImprime.Picture = LoadPicture(LoadGasPicture(34))
   Set botImprime.PictureDisabled = LoadPicture(LoadGasPicture(61))
   Set botImprime.PictureMouseOver = LoadPicture(LoadGasPicture(21))
   Set botExclui.Picture = LoadPicture(LoadGasPicture(29))
   Set botExclui.PictureDisabled = LoadPicture(LoadGasPicture(53))
   Set botExclui.PictureMouseOver = LoadPicture(LoadGasPicture(13))
   Set botSalva.Picture = LoadPicture(LoadGasPicture(27))
   Set botSalva.PictureDisabled = LoadPicture(LoadGasPicture(50))
   Set botSalva.PictureMouseOver = LoadPicture(LoadGasPicture(10))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(1000))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(1001))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(1002))
   Set Botao(4).Picture = LoadPicture(LoadGasPicture(1003))
   Set Botao(5).Picture = LoadPicture(LoadGasPicture(1004))
   Set Botao(6).Picture = LoadPicture(LoadGasPicture(1005))
   Set Botao(7).Picture = LoadPicture(LoadGasPicture(1006))
   Set Botao(8).Picture = LoadPicture(LoadGasPicture(1007))
   Set Botao(9).Picture = LoadPicture(LoadGasPicture(1008))

   'nome dos menus
   Set MnuCadastro = MDINFE.Menu_570(3)
   Set MnuParametros = MDINFE.Menu_570(2)
   Set mnuGeral = MDINFE.Menu_717(5)
   Set mnuCapitaliza = MDINFE.Menu_FRM00621
   Set mnuProdutos = MDINFE.Menu_717(10)
   Set MnuNotaFiscal = MDINFE.Menu_570(5)
   Set mnuNFe = MDINFE.Menu_625(0)
   Set MnuConsultas = MDINFE.Menu_570(9)
   Set MnuConNFe = MDINFE.Menu_632(0)
   Set MnuConGeral = MDINFE.Menu_632(1)
   Set MnuFerramentas = MDINFE.Menu_570(10)
   Set MnuEstoque = MDINFE.Menu_570(6)
   Set MnuEntradaEstoque = MDINFE.Menu_639(0)
   Set MnuFinanceiro = MDINFE.Menu_570(8)
   Set MnuRelFinanceiro = MDINFE.Menu_694(2)
   Set MnuContasPg = MDINFE.Menu_649(0)
   Set MnuBxPagar = MDINFE.Menu_694(1)
   Set MnuDevCompra = MDINFE.Menu_639(2)
   Set MnuReceber = MDINFE.Menu_649(1)
   Set MnuRelReceber = MDINFE.Menu_736(4)
   Set MnuDevCompra = MDINFE.Menu_632(4)
   Set MnuBxReceber = MDINFE.Menu_736(1)
   Set MnuSaidasEstoque = MDINFE.Menu_639(1)
   Set MnuCancelamento = MDINFE.Menu_639(3)
   Set MnuConEntrada = MDINFE.Menu_632(2)
   Set MnuConSaidas = MDINFE.Menu_632(3)
   Set MnuServicoPai = MDINFE.Menu_570(7)
   Set MnuOSReticom = MDINFE.Menu_680(0)
   Set MnuRelOS = MDINFE.Menu_680(4)
   Set MnuEntradaPg = MDINFE.Menu_694(0)
   Set MnuFaturamento = MDINFE.Menu_570(4)
   Set MnuVendas = MDINFE.Menu_709(1)
   Set MnuOrçamento = MDINFE.Menu_709(0)
   Set MnuCheque = MDINFE.Menu_649(2)
   Set MnuConPg = MDINFE.Menu_632(6)

   'prepara texto de ajuda para os objetos
   vgTooltips.Create
   vgTooltips.AddTool botInclui, 0, LoadGasString(300)
   botInclui.Caption = LoadGasString(301)
   vgTooltips.AddTool botCancela, 0, LoadGasString(302)
   botCancela.Caption = LoadGasString(303)
   vgTooltips.AddTool botBrowse, 0, LoadGasString(304)
   botBrowse.Caption = LoadGasString(305)
   vgTooltips.AddTool botPrimeiro, 0, LoadGasString(306)
   vgTooltips.AddTool botAnterior, 0, LoadGasString(307)
   vgTooltips.AddTool botSeguinte, 0, LoadGasString(308)
   vgTooltips.AddTool botUltimo, 0, LoadGasString(309)
   vgTooltips.AddTool botImprime, 0, LoadGasString(310)
   botImprime.Caption = LoadGasString(311)
   vgTooltips.AddTool botExclui, 0, LoadGasString(312)
   botExclui.Caption = LoadGasString(313)
   vgTooltips.AddTool botSalva, 0, LoadGasString(314)
   botSalva.Caption = LoadGasString(315)
   vgTooltips.AddTool Botao(0), 0, LoadGasString(316)
   Botao(1).Caption = LoadGasString(317)
   Botao(2).Caption = LoadGasString(318)
   Botao(3).Caption = LoadGasString(319)
   Botao(4).Caption = LoadGasString(320)
   Botao(5).Caption = LoadGasString(321)
   Botao(6).Caption = LoadGasString(322)
   Botao(7).Caption = LoadGasString(323)
   Botao(8).Caption = LoadGasString(324)
   Botao(9).Caption = LoadGasString(325)
   Label(3).Caption = LoadGasString(326)
   x$ = PegaStrDoIni("Opcoes", "Tempo refresh", vgNomeINI$) 'pega tempo refresh no INI
   If Val(x$) = 0 Then                                      'não achou,
      x$ = "25"                                             'fixamos 25 segundos
      GravaNoIni "Opcoes", "Tempo refresh", x$, vgNomeINI$  'e gravamos - na próxima vamos achar...
   End If
   timRefresh.Interval = Val(x$) * 1000                     'ajusta em milésimos de segundo

   x$ = PegaStrDoIni("Opcoes", "WindowState", vgNomeINI$)   'vê se o mdi estava maximizado
   If x$ = "2" Then WindowState = Val(x$)
   Set paneControles = Painel(0)
   Set cmdCapitaliza = Botao(0)
   Set lblTitulo = Label(0)
   Set lblGrupo = Label(1)
   Set lblUsuario = Label(2)
   Set lblDataHora = Label(3)
   Set lblUsuario2 = Label(4)
   Set lblAlteracao = Label(5)
   Set lblProgresso = Label(6)

   'ajusta este para ficar um pouco menor do que a tela, desconsiderando o systray
   If SystemParametersInfo(SPI_GETWORKAREA, 0, vgWinInfo, 0) <> 0 Then
      Width = (vgWinInfo.Right * Screen.TwipsPerPixelX) - (40 * Screen.TwipsPerPixelX) 'largura
      Height = (vgWinInfo.Bottom * Screen.TwipsPerPixelY) - (40 * Screen.TwipsPerPixelY) 'altura
   Else
      Width = Screen.Width - 720                  'largura
      Height = Screen.Height - 720                'altura
   End If
   Caption = vgNomeSistema$ + Chr$(160)           'o título será o título da aplicação
   vgTile.Tile.BackGroundType = 1
   vgTile.Tile.Picture = LoadPicture(LoadGasPicture(6))
   vgTile.Tile.TransPicture = LoadPicture(LoadGasPicture(7))
   vgTile.Init Me
   CentraNaTela Me                                'centraliza dentro da área de cliente
   LeParametrosForm Me                            'le parâmetros dest form no .INI e,
   Menu_Ajuda.Checked = vgAjudaAtiva              'marca/desmarca opção no menu
   mdiLoad  
End Sub

'tentou descarrega o form principal - vamos tratar...
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim i As Long                                       'dimensiona
   If vgCriandoSenha Then Exit Sub
   If QtForms > 0 And Visible Then                     'se tem outros forms e se o MDI está visível,
      For i = 1 To Forms.Count - 1                     'enquanto existir forms,
         If FormPendente(Forms(i)) Then                'se tiver pendente de atualização, vamos avisar
            If Forms(i).WindowState = vbMinimized Then 'se form está minimizado,
               Forms(i).WindowState = vbNormal         'ajusta para normal
            End If
            Forms(i).SetFocus                          'e voltar para lá...
            PrepBotoes Forms(i), Forms(i).vgSituacao   'ajeita barra de ferramentas
            Cancel = True                              'cancela a saída
            Exit For                                   'e sai...
         End If
      Next
   End If
   If Not Cancel Then                                  'se realmente vai abandonar a aplicação
      Do While Forms.Count > 1                         'e enquanto existir forms,
         If Forms(1).Name <> "mdiNFE" Then
            UnLoad Forms(1)                            'vamos descarrega-los 
         Else
            UnLoad Forms(0)                            'vamos descarrega-los
         End If
         DoEvents
      Loop
   End If
End Sub

Private Sub MDIForm_Resize()
   Dim vgLargMin As Long, vgAltMin As Long        'vamos limitar o tamanho deste form
   If WindowState = vbNormal Then                 'se não esta minimizado...
      vgAltMin = Screen.Height / 3                'altura e lagura mímina é
      vgLargMin = Screen.Width / 3                'igual a 1/3 do tamanho da tela
      If Height < vgAltMin Then Height = vgAltMin 'limita o rezise vertical e
      If Width < vgLargMin Then Width = vgLargMin 'horizontal
   End If
   vgTile.Refresh
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   If vgCriandoSenha Then Goto UnloadAll
   GravaParametrosForm Me                         'parâmetros genéricos
   timRefresh.Enabled = False                     'desabilita o timer de refresh dos dados
   GravaNoIni "Opcoes", "WindowState", Str$(WindowState), vgNomeINI$ 'estado do form
   Do While Forms.Count > 1                                          'enquanto existir forms,
      Unload Forms(1)                                                'vamos descarrega-los
   Loop
UnloadAll:
   Set paneControles = Nothing
   Set cmdCapitaliza = Nothing
   Set lblTitulo = Nothing
   Set lblGrupo = Nothing
   Set lblUsuario = Nothing
   Set lblDataHora = Nothing
   Set lblUsuario2 = Nothing
   Set lblAlteracao = Nothing
   Set lblProgresso = Nothing
   Set vgTooltips = Nothing                                          'libera memória da variável
   Set vgTile = Nothing
   GeraLogAcao LoadGasString(294)
   If Not vgCriandoSenha Then FinalizaAplicacao                      'Finaliza o programa
End Sub

Private Sub schNavega_Change()
   Dim i As Long
   Static vgValAnt As Integer
   If Not vgNaoAtRolagem Then                     'se não é para atualizar valor do scroll
      Screen.MousePointer = vbHourglass           'mouse = ampulheta
      vgBotoesOK = False                          'reseta flag
      With schNavega
         If .Value = vgValAnt + 1 Or .Value = vgValAnt - 1 Or _
          .Value = vgValAnt + 10 Or .Value = vgValAnt - 10 Then 'testa valor com valor anterior
            ActiveForm.vgTb.Move (.Value - vgValAnt)            'e move o necessário
         Else                                     'senão,
            i = (ActiveForm.vgTb.RecordCount - 1) * .Value / 100
            ActiveForm.vgTb.MoveFirst             'move para o primeiro, calcula posição
            ActiveForm.vgTb.Move i
         End If
      End With
      If ActiveForm.vgTb.EOF Then                 'se final de arquivo,
         ActiveForm.vgTb.MoveLast                 'move para o último registro
      ElseIf ActiveForm.vgTb.BOF Then             'se início,
         ActiveForm.vgTb.MoveFirst                'move para o primeiro
      End If
      vgBotoesOK = True                           'liga flag
      ActiveForm.Reposition                       'remonta a tela
      AjustaRolagem ActiveForm                    'ajusta a rolagem
      Screen.MousePointer = vbDefault
   End If
   vgValAnt = schNavega.Value                     'salva o valor do filtro
End Sub

'esta rotina é chamada a intervalos periódicos de tempo (cerca de 10 seg) para
'remontar os dados nos forms de visualização de dados
Private Sub timRefresh_Timer()
   Dim F As Form, vgMouseAtual As Integer            'dimensiona
   vgMouseAtual = Screen.MousePointer                'salva pointer do mouse
   Screen.MousePointer = vbHourglass                 'passa para ampulheta
   On Error GoTo DeuErro                             'prepara para erros
   If QtForms > 0 Then                               'se tiver algum form aberto
      For Each F In Forms
         If F.vgTipo = TP_TABELA Then                'se é form que contém controles vinculados
            If F.vgSituacao = ACAO_NAVEGANDO And F.vgPriVez = False Then 'se está somente navegando,
               If F.Name <> ActiveForm.Name And F.vgCaracteristica = F_DADOS Then
                  If Err Then
                     Err.Clear
                  Else
                     F.Reposition
                     If F.vgEmBrowse Then                                'se esta em grid
                        F.grdBrowse.Refresh                              'força uma remontagem
                     End If
                  End If
               End If
            End If
         End If
      Next
   End If
   Screen.MousePointer = vgMouseAtual                                    'restabelece o ponteiro do mouse
   On Error Resume Next
   mdiNFE.Botao(0).Visible = (False)
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(0).Visible = False
   End If
   mdiNFE.Botao(1).Enabled = (PermitidoMenu(mnuGeral.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(1).Enabled = False
   End If
   mdiNFE.Botao(2).Enabled = (PermitidoMenu(mnuProdutos.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(2).Enabled = False
   End If
   mdiNFE.Botao(3).Enabled = (PermitidoMenu(mnuNFe.Tag) And Empresa![Emite NFe])
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(3).Enabled = False
   End If
   mdiNFE.Botao(5).Enabled = (PermitidoMenu(MnuEstoque.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(5).Enabled = False
   End If
   mdiNFE.Botao(6).Enabled = (PermitidoMenu(MnuBxReceber.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(6).Enabled = False
   End If
   mdiNFE.Botao(7).Enabled = (PermitidoMenu(MnuBxPagar.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(7).Enabled = False
   End If
   mdiNFE.Botao(8).Visible = (Empresa![Nome fantasia]  = "Reticom Virabrequim")
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(8).Visible = False
   End If
   mdiNFE.Botao(8).Enabled = (PermitidoMenu(MnuOSReticom.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(8).Enabled = False
   End If
   mdiNFE.Botao(9).Visible = (Empresa![Nome fantasia]  <> "Reticom Virabrequim")
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(9).Visible = False
   End If
   mdiNFE.Botao(9).Enabled = (PermitidoMenu(MnuVendas.Tag))
   If Err > 0 Then
      Err.Clear
      mdiNFE.Botao(9).Enabled = False
   End If
   Exit Sub                                       'e cai fora...

DeuErro:                                          'se caiu aqui, deu erro...
   If Err = 3167 Then                             'EPA! este registro está deletado!
      MoveRegistro F, REG_FORCAVOLTA              'retrocede um registro...
   End If
   Resume Next                                    'e continua normalmente
End Sub

'evento - quando o botão for pressionado o Botão do PopupMenu
Public Sub AcaoPopUp0(Index As Integer)
   Select Case Index
      Case 1
         mdiNFE.MConNFe
      Case 2
         mdiNFE.MConGeral
      Case 3
         mdiNFE.MConMvtoE
      Case 4
         mdiNFE.MConMvtoS
      Case 5
         mdiNFE.MConDevCo
      Case 6
         mdiNFE.MConPagar
   End Select
End Sub

'evento - quando o botão for pressionado o Botão do PopupMenu
Public Sub AcaoPopUp1(Index As Integer)
   Select Case Index
      Case 1
         mdiNFE.MEstoque
      Case 2
         mdiNFE.MMoSaidas
   End Select
End Sub

'evento - quando o botão for pressionado o Botão do PopupMenu
Public Sub AcaoPopUp2(Index As Integer)
   Select Case Index
      Case 1
         mdiNFE.MEntConRe
      Case 2
         mdiNFE.MBxConRec
   End Select
End Sub

'evento - quando o botão for pressionado o Botão do PopupMenu
Public Sub AcaoPopUp3(Index As Integer)
   Select Case Index
      Case 1
         mdiNFE.MEntraPag
      Case 2
         mdiNFE.MBxPagar
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 1
         vgPriVez = True
         mdiNFE.MGeral
         vgPriVez = False
      Case 2
         vgPriVez = True
         mdiNFE.MProdutos
         vgPriVez = False
      Case 3
         vgPriVez = True
         mdiNFE.MNotafisc
         vgPriVez = False
      Case 4
         vgPriVez = True
         hMenu = CreatePopupMenu()
         AppendMenu hMenu, MF_STRING, 1, "Nota Fiscal"
         AppendMenu hMenu, MF_STRING, 2, "Geral"
         AppendMenu hMenu, MF_STRING, 3, "Entradas - Estoque"
         AppendMenu hMenu, MF_STRING, 4, "Saidas - Estoque"
         AppendMenu hMenu, MF_STRING, 5, "Devolução de Compra"
         AppendMenu hMenu, MF_STRING, 6, "Entrada Contas Pagar"
         GetCursorPos pt
         AcaoPopup0 TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
         DestroyMenu hMenu
         vgPriVez = False
      Case 5
         vgPriVez = True
         hMenu = CreatePopupMenu()
         AppendMenu hMenu, MF_STRING, 1, "Entradas"
         AppendMenu hMenu, MF_STRING, 2, "Saidas"
         GetCursorPos pt
         AcaoPopup1 TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
         DestroyMenu hMenu
         vgPriVez = False
      Case 6
         vgPriVez = True
         hMenu = CreatePopupMenu()
         AppendMenu hMenu, MF_STRING, 1, "Entrada"
         AppendMenu hMenu, MF_STRING, 2, "Baixa"
         GetCursorPos pt
         AcaoPopup2 TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
         DestroyMenu hMenu
         vgPriVez = False
      Case 7
         vgPriVez = True
         hMenu = CreatePopupMenu()
         AppendMenu hMenu, MF_STRING, 1, "Entrada "
         AppendMenu hMenu, MF_STRING, 2, "Baixa"
         GetCursorPos pt
         AcaoPopup3 TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.HWnd, ByVal 0&)
         DestroyMenu hMenu
         vgPriVez = False
      Case 8
         vgPriVez = True
         mdiNFE.MOServico
         vgPriVez = False
      Case 9
         vgPriVez = True
         mdiNFE.MVendas
         vgPriVez = False
   End Select
End Sub


'Capitaliza
Private Sub MFRM00621
   If mnuCapitaliza.Checked Then
    mnuCapitaliza.Checked = False
    cmdCapitaliza.Value = False
Else
    mnuCapitaliza.Checked = True
    cmdCapitaliza.Value = True
End If
End Sub

'Parametros do Emitente
Private Sub MEmpresa
   On Error Resume Next
   If frmEmpresa.WindowState = vbMinimized Then
      frmEmpresa.WindowState = vbNormal
   Else
      frmEmpresa.Show 
   End If
   frmEmpresa.SetFocus
End Sub

'Parametros da NFe
Private Sub MParNFe
   On Error Resume Next
   If frmParNFe.WindowState = vbMinimized Then
      frmParNFe.WindowState = vbNormal
   Else
      frmParNFe.Show 
   End If
   frmParNFe.SetFocus
End Sub

'Parametros da O.S
Private Sub MParamOS
   On Error Resume Next
   If frmParamOS.WindowState = vbMinimized Then
      frmParamOS.WindowState = vbNormal
   Else
      frmParamOS.Show 
   End If
   frmParamOS.SetFocus
End Sub

'Registro Zero
Private Sub MRegZero
   On Error Resume Next
   If frmRegZero.WindowState = vbMinimized Then
      frmRegZero.WindowState = vbNormal
   Else
      frmRegZero.Show 
   End If
   frmRegZero.SetFocus
End Sub

'Bancos
Private Sub MBanco
   On Error Resume Next
   If frmBanco.WindowState = vbMinimized Then
      frmBanco.WindowState = vbNormal
   Else
      frmBanco.Show 
   End If
   frmBanco.SetFocus
End Sub

'Icms
Private Sub MIcms
   On Error Resume Next
   If frmIcms.WindowState = vbMinimized Then
      frmIcms.WindowState = vbNormal
   Else
      frmIcms.Show 
   End If
   frmIcms.SetFocus
End Sub

'Classificação Fiscal
Public Sub MNcm
   On Error Resume Next
   If frmNcm.WindowState = vbMinimized Then
      frmNcm.WindowState = vbNormal
   Else
      frmNcm.Show 
   End If
   frmNcm.SetFocus
End Sub

'Natureza de Operação
Public Sub MCfop
   On Error Resume Next
   If frmCfop.WindowState = vbMinimized Then
      frmCfop.WindowState = vbNormal
   Else
      frmCfop.Show 
   End If
   frmCfop.SetFocus
End Sub

'Municípios
Public Sub MMunicipi
   On Error Resume Next
   If frmMunicipi.WindowState = vbMinimized Then
      frmMunicipi.WindowState = vbNormal
   Else
      frmMunicipi.Show 
   End If
   frmMunicipi.SetFocus
End Sub

'Geral
Public Sub MGeral
   On Error Resume Next
   If frmGeral.WindowState = vbMinimized Then
      frmGeral.WindowState = vbNormal
   Else
      frmGeral.Show 
   End If
   frmGeral.SetFocus
End Sub

'Unidades
Public Sub MUnidades
   On Error Resume Next
   If frmUnidades.WindowState = vbMinimized Then
      frmUnidades.WindowState = vbNormal
   Else
      frmUnidades.Show 
   End If
   frmUnidades.SetFocus
End Sub

'Marcas
Public Sub MMarcas
   On Error Resume Next
   If frmMarcas.WindowState = vbMinimized Then
      frmMarcas.WindowState = vbNormal
   Else
      frmMarcas.Show 
   End If
   frmMarcas.SetFocus
End Sub

'Tipo de Cobrança
Private Sub MTpCobrar
   On Error Resume Next
   If frmTpCobrar.WindowState = vbMinimized Then
      frmTpCobrar.WindowState = vbNormal
   Else
      frmTpCobrar.Show 
   End If
   frmTpCobrar.SetFocus
End Sub

'Centro de Custo
Public Sub MCentroC
   On Error Resume Next
   If frmCentroC.WindowState = vbMinimized Then
      frmCentroC.WindowState = vbNormal
   Else
      frmCentroC.Show 
   End If
   frmCentroC.SetFocus
End Sub

'Produtos
Public Sub MProdutos
   On Error Resume Next
   If frmProdutos.WindowState = vbMinimized Then
      frmProdutos.WindowState = vbNormal
   Else
      frmProdutos.Show 
   End If
   frmProdutos.SetFocus
End Sub

'Serviços
Private Sub MServicos
   On Error Resume Next
   If frmServicos.WindowState = vbMinimized Then
      frmServicos.WindowState = vbNormal
   Else
      frmServicos.Show 
   End If
   frmServicos.SetFocus
End Sub

'Virabrequim
Private Sub MVirabreq
   On Error Resume Next
   If frmVirabreq.WindowState = vbMinimized Then
      frmVirabreq.WindowState = vbNormal
   Else
      frmVirabreq.Show 
   End If
   frmVirabreq.SetFocus
End Sub

'Modelo do Virabrequim
Private Sub MModelo
   On Error Resume Next
   If frmModelo.WindowState = vbMinimized Then
      frmModelo.WindowState = vbNormal
   Else
      frmModelo.Show 
   End If
   frmModelo.SetFocus
End Sub

'Orçamento
Private Sub MFRM00709
   If OrcamentoAberto Then
    Orcamento.SetFocus
    Exit Sub
 End If
   Orcamento.vgFiltroInicial = ""
   Orcamento.Caption = "Orçamento"
   Orcamento.vgFiltroInicial = "[Sequencia da Venda] > 0 AND [Tipo de Venda] = 1 And Fechado = 0" '0 Venda 1 Orçamento 
   InicializaFiltro Orcamento
   Orcamento.BloqueiaAba
   Orcamento.CarregaTotalizador
   Orcamento.Show   
   Orcamento.Executar INI_APELIDOS, 0
   OrcamentoAberto = True
End Sub

'Vendas
Public Sub MVendas
   On Error Resume Next
   If frmVendas.WindowState = vbMinimized Then
      frmVendas.WindowState = vbNormal
   Else
      frmVendas.Show 
   End If
   frmVendas.SetFocus
End Sub

'Relatório de Vendas
Private Sub MR_Vendas
   On Error Resume Next
   If frmR_Vendas.WindowState = vbMinimized Then
      frmR_Vendas.WindowState = vbNormal
      frmR_Vendas.SetFocus
   Else
      If frmR_Vendas.vgVisivel Then
         frmR_Vendas.Show 
         frmR_Vendas.SetFocus
      Else
         If frmR_Vendas.vgFP.WindowState = vbMinimized Then
            frmR_Vendas.vgFP.WindowState = vbNormal
         Else
            frmR_Vendas.vgFP.Show 
         End If
         frmR_Vendas.vgFP.SetFocus
      End if
   End if
End Sub

'Relatório de Vendas por Periodo
Private Sub MRelVdaPe
   On Error Resume Next
   If frmRelVdaPe.WindowState = vbMinimized Then
      frmRelVdaPe.WindowState = vbNormal
      frmRelVdaPe.SetFocus
   Else
      If frmRelVdaPe.vgVisivel Then
         frmRelVdaPe.Show 
         frmRelVdaPe.SetFocus
      Else
         If frmRelVdaPe.vgFP.WindowState = vbMinimized Then
            frmRelVdaPe.vgFP.WindowState = vbNormal
         Else
            frmRelVdaPe.vgFP.Show 
         End If
         frmRelVdaPe.vgFP.SetFocus
      End if
   End if
End Sub

'Relatório do Estoque
Private Sub MRelProdu
   On Error Resume Next
   If frmRelProdu.WindowState = vbMinimized Then
      frmRelProdu.WindowState = vbNormal
      frmRelProdu.SetFocus
   Else
      If frmRelProdu.vgVisivel Then
         frmRelProdu.Show 
         frmRelProdu.SetFocus
      Else
         If frmRelProdu.vgFP.WindowState = vbMinimized Then
            frmRelProdu.vgFP.WindowState = vbNormal
         Else
            frmRelProdu.vgFP.Show 
         End If
         frmRelProdu.vgFP.SetFocus
      End if
   End if
End Sub

'Vendas - Nota Fiscal
Private Sub MVendaNFe
   On Error Resume Next
   If frmVendaNFe.WindowState = vbMinimized Then
      frmVendaNFe.WindowState = vbNormal
   Else
      frmVendaNFe.Show 
   End If
   frmVendaNFe.SetFocus
End Sub

'Imprimir
Private Sub MRelVenda
   On Error Resume Next
   If frmRelVenda.WindowState = vbMinimized Then
      frmRelVenda.WindowState = vbNormal
      frmRelVenda.SetFocus
   Else
      If frmRelVenda.vgVisivel Then
         frmRelVenda.Show 
         frmRelVenda.SetFocus
      Else
         If frmRelVenda.vgFP.WindowState = vbMinimized Then
            frmRelVenda.vgFP.WindowState = vbNormal
         Else
            frmRelVenda.vgFP.Show 
         End If
         frmRelVenda.vgFP.SetFocus
      End if
   End if
End Sub

'Imprimir - Meia Folha
Private Sub MRelImpV2
   On Error Resume Next
   If frmRelImpV2.WindowState = vbMinimized Then
      frmRelImpV2.WindowState = vbNormal
      frmRelImpV2.SetFocus
   Else
      If frmRelImpV2.vgVisivel Then
         frmRelImpV2.Show 
         frmRelImpV2.SetFocus
      Else
         If frmRelImpV2.vgFP.WindowState = vbMinimized Then
            frmRelImpV2.vgFP.WindowState = vbNormal
         Else
            frmRelImpV2.vgFP.Show 
         End If
         frmRelImpV2.vgFP.SetFocus
      End if
   End if
End Sub

'NFe
Public Sub MNotafisc
   On Error Resume Next
   If frmNotafisc.WindowState = vbMinimized Then
      frmNotafisc.WindowState = vbNormal
   Else
      frmNotafisc.Show 
   End If
   frmNotafisc.SetFocus
End Sub

'Inutiliza - NFe
Private Sub MInutiliz
   On Error Resume Next
   If frmInutiliz.WindowState = vbMinimized Then
      frmInutiliz.WindowState = vbNormal
   Else
      frmInutiliz.Show 
   End If
   frmInutiliz.SetFocus
End Sub

'Notas - Emitidas
Private Sub MRelNFe
   On Error Resume Next
   If frmRelNFe.WindowState = vbMinimized Then
      frmRelNFe.WindowState = vbNormal
      frmRelNFe.SetFocus
   Else
      If frmRelNFe.vgVisivel Then
         frmRelNFe.Show 
         frmRelNFe.SetFocus
      Else
         If frmRelNFe.vgFP.WindowState = vbMinimized Then
            frmRelNFe.vgFP.WindowState = vbNormal
         Else
            frmRelNFe.vgFP.Show 
         End If
         frmRelNFe.vgFP.SetFocus
      End if
   End if
End Sub

'Cancela Nota Fiscal
Private Sub MF_CanNF
   On Error Resume Next
   If frmF_CanNF.WindowState = vbMinimized Then
      frmF_CanNF.WindowState = vbNormal
   Else
      frmF_CanNF.Show 
   End If
   frmF_CanNF.SetFocus
End Sub

'Carta de Correção
Private Sub MCCe
   On Error Resume Next
   If frmCCe.WindowState = vbMinimized Then
      frmCCe.WindowState = vbNormal
   Else
      frmCCe.Show 
   End If
   frmCCe.SetFocus
End Sub

'Ativador
Private Sub MAtivador
   If frmAtivador.WindowState = vbMinimized Then
      frmAtivador.WindowState = vbNormal
   Else
      frmAtivador.Show vbModal
   End If
End Sub

'Entradas - Estoque
Public Sub MEstoque
   On Error Resume Next
   If frmEstoque.WindowState = vbMinimized Then
      frmEstoque.WindowState = vbNormal
   Else
      frmEstoque.Show 
   End If
   frmEstoque.SetFocus
End Sub

'Saidas - Estoque
Public Sub MMoSaidas
   On Error Resume Next
   If frmMoSaidas.WindowState = vbMinimized Then
      frmMoSaidas.WindowState = vbNormal
   Else
      frmMoSaidas.Show 
   End If
   frmMoSaidas.SetFocus
End Sub

'Devolução
Private Sub MAbrDComp
   On Error Resume Next
   If frmAbrDComp.WindowState = vbMinimized Then
      frmAbrDComp.WindowState = vbNormal
   Else
      frmAbrDComp.Show 
   End If
   frmAbrDComp.SetFocus
End Sub

'Cancela Devolução
Private Sub MCanDComp
   On Error Resume Next
   If frmCanDComp.WindowState = vbMinimized Then
      frmCanDComp.WindowState = vbNormal
   Else
      frmCanDComp.Show 
   End If
   frmCanDComp.SetFocus
End Sub

'Cancela - Entrada
Private Sub MCan_Entr
   On Error Resume Next
   If frmCan_Entr.WindowState = vbMinimized Then
      frmCan_Entr.WindowState = vbNormal
   Else
      frmCan_Entr.Show 
   End If
   frmCan_Entr.SetFocus
End Sub

'Cancela - Saida
Private Sub MCan_said
   On Error Resume Next
   If frmCan_said.WindowState = vbMinimized Then
      frmCan_said.WindowState = vbNormal
   Else
      frmCan_said.Show 
   End If
   frmCan_said.SetFocus
End Sub

'Importar NFe
Private Sub MIXML
   On Error Resume Next
   If frmIXML.WindowState = vbMinimized Then
      frmIXML.WindowState = vbNormal
   Else
      frmIXML.Show 
   End If
   frmIXML.SetFocus
End Sub

'Devolução - Nota Fiscal
Private Sub MF_GeraNF
   On Error Resume Next
   If frmF_GeraNF.WindowState = vbMinimized Then
      frmF_GeraNF.WindowState = vbNormal
   Else
      frmF_GeraNF.Show 
   End If
   frmF_GeraNF.SetFocus
End Sub

'Movimento - Nota Fiscal
Private Sub MMvtoNFe
   On Error Resume Next
   If frmMvtoNFe.WindowState = vbMinimized Then
      frmMvtoNFe.WindowState = vbNormal
   Else
      frmMvtoNFe.Show 
   End If
   frmMvtoNFe.SetFocus
End Sub

'Outras Saidas - Nota Fiscal
Private Sub MF_GNF
   On Error Resume Next
   If frmF_GNF.WindowState = vbMinimized Then
      frmF_GNF.WindowState = vbNormal
   Else
      frmF_GNF.Show 
   End If
   frmF_GNF.SetFocus
End Sub

'Devolução de Compra
Private Sub MDevCompr
   On Error Resume Next
   If frmDevCompr.WindowState = vbMinimized Then
      frmDevCompr.WindowState = vbNormal
   Else
      frmDevCompr.Show 
   End If
   frmDevCompr.SetFocus
End Sub

'O.S
Public Sub MOServico
   On Error Resume Next
   If frmOServico.WindowState = vbMinimized Then
      frmOServico.WindowState = vbNormal
   Else
      frmOServico.Show 
   End If
   frmOServico.SetFocus
End Sub

'Baixa - Casco
Private Sub MBxCasco
   On Error Resume Next
   If frmBxCasco.WindowState = vbMinimized Then
      frmBxCasco.WindowState = vbNormal
   Else
      frmBxCasco.Show 
   End If
   frmBxCasco.SetFocus
End Sub

'Dados Adicionais - O.S
Private Sub MOsDadosA
   On Error Resume Next
   If frmOsDadosA.WindowState = vbMinimized Then
      frmOsDadosA.WindowState = vbNormal
   Else
      frmOsDadosA.Show 
   End If
   frmOsDadosA.SetFocus
End Sub

'Dados Adicionais - Virabrequim
Private Sub MDadVirab
   On Error Resume Next
   If frmDadVirab.WindowState = vbMinimized Then
      frmDadVirab.WindowState = vbNormal
   Else
      frmDadVirab.Show 
   End If
   frmDadVirab.SetFocus
End Sub

'Virabrequim
Private Sub MR_Virab
   On Error Resume Next
   If frmR_Virab.WindowState = vbMinimized Then
      frmR_Virab.WindowState = vbNormal
      frmR_Virab.SetFocus
   Else
      If frmR_Virab.vgVisivel Then
         frmR_Virab.Show 
         frmR_Virab.SetFocus
      Else
         If frmR_Virab.vgFP.WindowState = vbMinimized Then
            frmR_Virab.vgFP.WindowState = vbNormal
         Else
            frmR_Virab.vgFP.Show 
         End If
         frmR_Virab.vgFP.SetFocus
      End if
   End if
End Sub

'O.S - Emitidas
Private Sub MR_OSEmit
   On Error Resume Next
   If frmR_OSEmit.WindowState = vbMinimized Then
      frmR_OSEmit.WindowState = vbNormal
      frmR_OSEmit.SetFocus
   Else
      If frmR_OSEmit.vgVisivel Then
         frmR_OSEmit.Show 
         frmR_OSEmit.SetFocus
      Else
         If frmR_OSEmit.vgFP.WindowState = vbMinimized Then
            frmR_OSEmit.vgFP.WindowState = vbNormal
         Else
            frmR_OSEmit.vgFP.Show 
         End If
         frmR_OSEmit.vgFP.SetFocus
      End if
   End if
End Sub

'Cascos
Private Sub MRelCasco
   On Error Resume Next
   If frmRelCasco.WindowState = vbMinimized Then
      frmRelCasco.WindowState = vbNormal
      frmRelCasco.SetFocus
   Else
      If frmRelCasco.vgVisivel Then
         frmRelCasco.Show 
         frmRelCasco.SetFocus
      Else
         If frmRelCasco.vgFP.WindowState = vbMinimized Then
            frmRelCasco.vgFP.WindowState = vbNormal
         Else
            frmRelCasco.vgFP.Show 
         End If
         frmRelCasco.vgFP.SetFocus
      End if
   End if
End Sub

'O.S - Emitadas (Exportar Excel)
Private Sub MOSExcel
   On Error Resume Next
   If frmOSExcel.WindowState = vbMinimized Then
      frmOSExcel.WindowState = vbNormal
   Else
      frmOSExcel.Show 
   End If
   frmOSExcel.SetFocus
End Sub

'Geral - Exportar Excel
Private Sub MGexcel
   On Error Resume Next
   If frmGexcel.WindowState = vbMinimized Then
      frmGexcel.WindowState = vbNormal
   Else
      frmGexcel.Show 
   End If
   frmGexcel.SetFocus
End Sub

'Imprimir
Public Sub MImprimOS
   On Error Resume Next
   If frmImprimOS.WindowState = vbMinimized Then
      frmImprimOS.WindowState = vbNormal
      frmImprimOS.SetFocus
   Else
      If frmImprimOS.vgVisivel Then
         frmImprimOS.Show 
         frmImprimOS.SetFocus
      Else
         If frmImprimOS.vgFP.WindowState = vbMinimized Then
            frmImprimOS.vgFP.WindowState = vbNormal
         Else
            frmImprimOS.vgFP.Show 
         End If
         frmImprimOS.vgFP.SetFocus
      End if
   End if
End Sub

'Entrada Contas
Public Sub MEntraPag
   On Error Resume Next
   If frmEntraPag.WindowState = vbMinimized Then
      frmEntraPag.WindowState = vbNormal
   Else
      frmEntraPag.Show 
   End If
   frmEntraPag.SetFocus
End Sub

'Baixa Contas
Public Sub MBxPagar
   On Error Resume Next
   If frmBxPagar.WindowState = vbMinimized Then
      frmBxPagar.WindowState = vbNormal
   Else
      frmBxPagar.Show 
   End If
   frmBxPagar.SetFocus
End Sub

'Contas a Pagar
Private Sub MRelPagar
   On Error Resume Next
   If frmRelPagar.WindowState = vbMinimized Then
      frmRelPagar.WindowState = vbNormal
      frmRelPagar.SetFocus
   Else
      If frmRelPagar.vgVisivel Then
         frmRelPagar.Show 
         frmRelPagar.SetFocus
      Else
         If frmRelPagar.vgFP.WindowState = vbMinimized Then
            frmRelPagar.vgFP.WindowState = vbNormal
         Else
            frmRelPagar.vgFP.Show 
         End If
         frmRelPagar.vgFP.SetFocus
      End if
   End if
End Sub

'Contas Pagas
Private Sub MPagas
   On Error Resume Next
   If frmPagas.WindowState = vbMinimized Then
      frmPagas.WindowState = vbNormal
      frmPagas.SetFocus
   Else
      If frmPagas.vgVisivel Then
         frmPagas.Show 
         frmPagas.SetFocus
      Else
         If frmPagas.vgFP.WindowState = vbMinimized Then
            frmPagas.vgFP.WindowState = vbNormal
         Else
            frmPagas.vgFP.Show 
         End If
         frmPagas.vgFP.SetFocus
      End if
   End if
End Sub

'Comissão dos Vendedores
Private Sub MRelComis
   On Error Resume Next
   If frmRelComis.WindowState = vbMinimized Then
      frmRelComis.WindowState = vbNormal
   Else
      frmRelComis.Show 
   End If
   frmRelComis.SetFocus
End Sub

'Entrada Contas
Public Sub MEntConRe
   On Error Resume Next
   If frmEntConRe.WindowState = vbMinimized Then
      frmEntConRe.WindowState = vbNormal
   Else
      frmEntConRe.Show 
   End If
   frmEntConRe.SetFocus
End Sub

'Baixa Contas
Public Sub MBxConRec
   On Error Resume Next
   If frmBxConRec.WindowState = vbMinimized Then
      frmBxConRec.WindowState = vbNormal
   Else
      frmBxConRec.Show 
   End If
   frmBxConRec.SetFocus
End Sub

'Manutenção - Contas
Private Sub MConReceb
   On Error Resume Next
   If frmConReceb.WindowState = vbMinimized Then
      frmConReceb.WindowState = vbNormal
   Else
      frmConReceb.Show 
   End If
   frmConReceb.SetFocus
End Sub

'Manutenção Contas - Lote
Private Sub MManuLote
   On Error Resume Next
   If frmManuLote.WindowState = vbMinimized Then
      frmManuLote.WindowState = vbNormal
   Else
      frmManuLote.Show 
   End If
   frmManuLote.SetFocus
End Sub

'Contas a Receber
Private Sub MRelReceb
   On Error Resume Next
   If frmRelReceb.WindowState = vbMinimized Then
      frmRelReceb.WindowState = vbNormal
      frmRelReceb.SetFocus
   Else
      If frmRelReceb.vgVisivel Then
         frmRelReceb.Show 
         frmRelReceb.SetFocus
      Else
         If frmRelReceb.vgFP.WindowState = vbMinimized Then
            frmRelReceb.vgFP.WindowState = vbNormal
         Else
            frmRelReceb.vgFP.Show 
         End If
         frmRelReceb.vgFP.SetFocus
      End if
   End if
End Sub

'Contas - Recebidas
Private Sub MR_Recebi
   On Error Resume Next
   If frmR_Recebi.WindowState = vbMinimized Then
      frmR_Recebi.WindowState = vbNormal
      frmR_Recebi.SetFocus
   Else
      If frmR_Recebi.vgVisivel Then
         frmR_Recebi.Show 
         frmR_Recebi.SetFocus
      Else
         If frmR_Recebi.vgFP.WindowState = vbMinimized Then
            frmR_Recebi.vgFP.WindowState = vbNormal
         Else
            frmR_Recebi.vgFP.Show 
         End If
         frmR_Recebi.vgFP.SetFocus
      End if
   End if
End Sub

'Baixa - Contas á Receber
Private Sub MRelBxRec
   On Error Resume Next
   If frmRelBxRec.WindowState = vbMinimized Then
      frmRelBxRec.WindowState = vbNormal
      frmRelBxRec.SetFocus
   Else
      If frmRelBxRec.vgVisivel Then
         frmRelBxRec.Show 
         frmRelBxRec.SetFocus
      Else
         If frmRelBxRec.vgFP.WindowState = vbMinimized Then
            frmRelBxRec.vgFP.WindowState = vbNormal
         Else
            frmRelBxRec.vgFP.Show 
         End If
         frmRelBxRec.vgFP.SetFocus
      End if
   End if
End Sub

'Cheques
Private Sub MCadChequ
   On Error Resume Next
   If frmCadChequ.WindowState = vbMinimized Then
      frmCadChequ.WindowState = vbNormal
   Else
      frmCadChequ.Show 
   End If
   frmCadChequ.SetFocus
End Sub

'Relatorio de Cheques de Terceiros
Private Sub MRelChequ
   On Error Resume Next
   If frmRelChequ.WindowState = vbMinimized Then
      frmRelChequ.WindowState = vbNormal
      frmRelChequ.SetFocus
   Else
      If frmRelChequ.vgVisivel Then
         frmRelChequ.Show 
         frmRelChequ.SetFocus
      Else
         If frmRelChequ.vgFP.WindowState = vbMinimized Then
            frmRelChequ.vgFP.WindowState = vbNormal
         Else
            frmRelChequ.vgFP.Show 
         End If
         frmRelChequ.vgFP.SetFocus
      End if
   End if
End Sub

'Manutenção - Receber
Private Sub MManuRece
   On Error Resume Next
   If frmManuRece.WindowState = vbMinimized Then
      frmManuRece.WindowState = vbNormal
   Else
      frmManuRece.Show 
   End If
   frmManuRece.SetFocus
End Sub

'Comissão dos Vendedores
Private Sub MRelComVe
   On Error Resume Next
   If frmRelComVe.WindowState = vbMinimized Then
      frmRelComVe.WindowState = vbNormal
      frmRelComVe.SetFocus
   Else
      If frmRelComVe.vgVisivel Then
         frmRelComVe.Show 
         frmRelComVe.SetFocus
      Else
         If frmRelComVe.vgFP.WindowState = vbMinimized Then
            frmRelComVe.vgFP.WindowState = vbNormal
         Else
            frmRelComVe.vgFP.Show 
         End If
         frmRelComVe.vgFP.SetFocus
      End if
   End if
End Sub

'Nota Fiscal
Public Sub MConNFe
   On Error Resume Next
   If frmConNFe.WindowState = vbMinimized Then
      frmConNFe.WindowState = vbNormal
   Else
      frmConNFe.Show 
   End If
   frmConNFe.SetFocus
End Sub

'Geral
Public Sub MConGeral
   On Error Resume Next
   If frmConGeral.WindowState = vbMinimized Then
      frmConGeral.WindowState = vbNormal
   Else
      frmConGeral.Show 
   End If
   frmConGeral.SetFocus
End Sub

'Entradas - Estoque
Public Sub MConMvtoE
   On Error Resume Next
   If frmConMvtoE.WindowState = vbMinimized Then
      frmConMvtoE.WindowState = vbNormal
   Else
      frmConMvtoE.Show 
   End If
   frmConMvtoE.SetFocus
End Sub

'Saidas - Estoque
Public Sub MConMvtoS
   On Error Resume Next
   If frmConMvtoS.WindowState = vbMinimized Then
      frmConMvtoS.WindowState = vbNormal
   Else
      frmConMvtoS.Show 
   End If
   frmConMvtoS.SetFocus
End Sub

'Devolução de Compra
Public Sub MConDevCo
   On Error Resume Next
   If frmConDevCo.WindowState = vbMinimized Then
      frmConDevCo.WindowState = vbNormal
   Else
      frmConDevCo.Show 
   End If
   frmConDevCo.SetFocus
End Sub

'Cheques
Public Sub MConChequ
   On Error Resume Next
   If frmConChequ.WindowState = vbMinimized Then
      frmConChequ.WindowState = vbNormal
   Else
      frmConChequ.Show 
   End If
   frmConChequ.SetFocus
End Sub

'Entrada - Contas
Public Sub MConPagar
   On Error Resume Next
   If frmConPagar.WindowState = vbMinimized Then
      frmConPagar.WindowState = vbNormal
   Else
      frmConPagar.Show 
   End If
   frmConPagar.SetFocus
End Sub

'Copia de Segurança
Private Sub MBackup
   On Error Resume Next
   If frmBackup.WindowState = vbMinimized Then
      frmBackup.WindowState = vbNormal
   Else
      frmBackup.Show 
   End If
   frmBackup.SetFocus
End Sub

'Senha Administrador
Private Sub MSenhaAdm
   If frmSenhaAdm.WindowState = vbMinimized Then
      frmSenhaAdm.WindowState = vbNormal
   Else
      frmSenhaAdm.Show vbModal
   End If
End Sub

'evento - quando o menu de &Parametros do Emitente for pressionado
Private Sub Menu_614_Click(Index As Integer)
   Select Case Index
      Case 0 'Parametros do Emitente
         MEmpresa
      Case 1 'Parametros da NFe
         MParNFe
      Case 2 'Parametros da O.S
         MParamOS
      Case 3 'Registro Zero
         MRegZero
   End Select
End Sub

'evento - quando o menu de &Bancos for pressionado
Private Sub Menu_717_Click(Index As Integer)
   Select Case Index
      Case 0 'Bancos
         MBanco
      Case 1 'Icms
         MIcms
      Case 2 'Classificação Fiscal
         MNcm
      Case 3 'Natureza de Operação
         MCfop
      Case 4 'Municípios
         MMunicipi
      Case 5 'Geral
         MGeral
      Case 6 'Unidades
         MUnidades
      Case 7 'Marcas
         MMarcas
      Case 8 'Tipo de Cobrança
         MTpCobrar
      Case 9 'Centro de Custo
         MCentroC
      Case 10 'Produtos
         MProdutos
      Case 11 'Serviços
         MServicos
      Case 12 'Virabrequim
         MVirabreq
      Case 13 'Modelo do Virabrequim
         MModelo
   End Select
End Sub

'evento - quando o menu de Orçamento for pressionado
Private Sub Menu_709_Click(Index As Integer)
   Select Case Index
      Case 0 'Orçamento
         MFRM00709
      Case 1 'Vendas
         MVendas
   End Select
End Sub

'evento - quando o menu de Relatório de Vendas for pressionado
Private Sub Menu_729_Click(Index As Integer)
   Select Case Index
      Case 0 'Relatório de Vendas
         MR_Vendas
      Case 1 'Relatório de Vendas por Periodo
         MRelVdaPe
      Case 2 'Relatório do Estoque
         MRelProdu
   End Select
End Sub

'evento - quando o menu de Vendas - Nota Fiscal for pressionado
Private Sub Menu_707_Click(Index As Integer)
   Select Case Index
      Case 0 'Vendas - Nota Fiscal
         MVendaNFe
      Case 1 'Imprimir
         MRelVenda
      Case 2 'Imprimir - Meia Folha
         MRelImpV2
   End Select
End Sub

'evento - quando o menu de &NFe for pressionado
Private Sub Menu_625_Click(Index As Integer)
   Select Case Index
      Case 0 'NFe
         MNotafisc
      Case 1 'Inutiliza - NFe
         MInutiliz
      Case 2 'Notas - Emitidas
         MRelNFe
   End Select
End Sub

'evento - quando o menu de Cancela Nota Fiscal for pressionado
Private Sub Menu_628_Click(Index As Integer)
   Select Case Index
      Case 0 'Cancela Nota Fiscal
         MF_CanNF
      Case 1 'Carta de Correção
         MCCe
      Case 2 'Ativador
         MAtivador
   End Select
End Sub

'evento - quando o menu de &Entradas - Estoque for pressionado
Private Sub Menu_639_Click(Index As Integer)
   Select Case Index
      Case 0 'Entradas - Estoque
         MEstoque
      Case 1 'Saidas - Estoque
         MMoSaidas
   End Select
End Sub

'evento - quando o menu de Devolução for pressionado
Private Sub Menu_654_Click(Index As Integer)
   Select Case Index
      Case 0 'Devolução
         MAbrDComp
      Case 1 'Cancela Devolução
         MCanDComp
   End Select
End Sub

'evento - quando o menu de Cancela - Entrada for pressionado
Private Sub Menu_673_Click(Index As Integer)
   Select Case Index
      Case 0 'Cancela - Entrada
         MCan_Entr
      Case 1 'Cancela - Saida
         MCan_said
   End Select
End Sub

'evento - quando o menu de Importar NFe for pressionado
Private Sub Menu_642_Click(Index As Integer)
   Select Case Index
      Case 0 'Importar NFe
         MIXML
      Case 1 'Devolução - Nota Fiscal
         MF_GeraNF
      Case 2 'Movimento - Nota Fiscal
         MMvtoNFe
      Case 3 'Outras Saidas - Nota Fiscal
         MF_GNF
      Case 4 'Devolução de Compra
         MDevCompr
   End Select
End Sub

'evento - quando o menu de &O.S for pressionado
Private Sub Menu_680_Click(Index As Integer)
   Select Case Index
      Case 0 'O.S
         MOServico
      Case 1 'Baixa - Casco
         MBxCasco
      Case 2 'Dados Adicionais - O.S
         MOsDadosA
      Case 3 'Dados Adicionais - Virabrequim
         MDadVirab
   End Select
End Sub

'evento - quando o menu de Virabrequim for pressionado
Private Sub Menu_725_Click(Index As Integer)
   Select Case Index
      Case 0 'Virabrequim
         MR_Virab
      Case 1 'O.S - Emitidas
         MR_OSEmit
      Case 2 'Cascos
         MRelCasco
      Case 3 'O.S - Emitadas (Exportar Excel)
         MOSExcel
      Case 4 'Geral - Exportar Excel
         MGexcel
   End Select
End Sub

'evento - quando o menu de Imprimir for pressionado
Private Sub Menu_688_Click(Index As Integer)
   Select Case Index
      Case 0 'Imprimir
         MImprimOS
   End Select
End Sub

'evento - quando o menu de &Entrada Contas for pressionado
Private Sub Menu_694_Click(Index As Integer)
   Select Case Index
      Case 0 'Entrada Contas
         MEntraPag
      Case 1 'Baixa Contas
         MBxPagar
   End Select
End Sub

'evento - quando o menu de Contas a Pagar for pressionado
Private Sub Menu_648_Click(Index As Integer)
   Select Case Index
      Case 0 'Contas a Pagar
         MRelPagar
      Case 1 'Contas Pagas
         MPagas
      Case 2 'Comissão dos Vendedores
         MRelComis
   End Select
End Sub

'evento - quando o menu de &Entrada Contas for pressionado
Private Sub Menu_736_Click(Index As Integer)
   Select Case Index
      Case 0 'Entrada Contas
         MEntConRe
      Case 1 'Baixa Contas
         MBxConRec
      Case 2 'Manutenção - Contas
         MConReceb
      Case 3 'Manutenção Contas - Lote
         MManuLote
   End Select
End Sub

'evento - quando o menu de &Contas a Receber for pressionado
Private Sub Menu_662_Click(Index As Integer)
   Select Case Index
      Case 0 'Contas a Receber
         MRelReceb
      Case 1 'Contas - Recebidas
         MR_Recebi
      Case 2 'Baixa - Contas á Receber
         MRelBxRec
   End Select
End Sub

'evento - quando o menu de &Cheques for pressionado
Private Sub Menu_719_Click(Index As Integer)
   Select Case Index
      Case 0 'Cheques
         MCadChequ
   End Select
End Sub

'evento - quando o menu de Relatorio de Cheques de Terceiros for pressionado
Private Sub Menu_721_Click(Index As Integer)
   Select Case Index
      Case 0 'Relatorio de Cheques de Terceiros
         MRelChequ
   End Select
End Sub

'evento - quando o menu de &Manutenção - Receber for pressionado
Private Sub Menu_697_Click(Index As Integer)
   Select Case Index
      Case 0 'Manutenção - Receber
         MManuRece
      Case 1 'Comissão dos Vendedores
         MRelComVe
   End Select
End Sub

'evento - quando o menu de Nota Fiscal for pressionado
Private Sub Menu_632_Click(Index As Integer)
   Select Case Index
      Case 0 'Nota Fiscal
         MConNFe
      Case 1 'Geral
         MConGeral
      Case 2 'Entradas - Estoque
         MConMvtoE
      Case 3 'Saidas - Estoque
         MConMvtoS
      Case 4 'Devolução de Compra
         MConDevCo
      Case 5 'Cheques
         MConChequ
      Case 6 'Entrada - Contas
         MConPagar
   End Select
End Sub

'evento - quando o menu de Copia de Segurança for pressionado
Private Sub Menu_637_Click(Index As Integer)
   Select Case Index
      Case 0 'Copia de Segurança
         MBackup
   End Select
End Sub

'evento - quando o menu de Senha Administrador for pressionado
Private Sub Menu_645_Click(Index As Integer)
   Select Case Index
      Case 0 'Senha Administrador
         MSenhaAdm
   End Select
End Sub

'evento - quando o menu de Capitaliza for pressionado
Private Sub Menu_FRM00621_Click()
   MFRM00621
End Sub 

