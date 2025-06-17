VERSION 5.00
Begin VB.Form frmFinanc
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   7545
   Left                     =   5010
   LinkTopic                =   "Financ"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7110
   ScaleWidth               =   10350
   Top                      =   2275
   Width                    =   10440
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
      Height                   =   6930
      Left                     =   15
      Width                    =   10170
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   17929
      _extenty                 =   12217
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   75
         Width                    =   2040
         Top                      =   195
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
         Caption                  =   "Projeto Nº........................:"
         Index                    =   0 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFC0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   330
         Left                     =   2355
         Width                    =   1440
         Top                      =   195
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
         Left                     =   75
         Width                    =   2055
         Top                      =   615
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
         Caption                  =   "Forma de Pagamento.:"
         Index                    =   1 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2340
         Top                      =   615
         Width                    =   1470
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   3540
         Top                      =   630
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
         Left                     =   2355
         Width                    =   1170
         Top                      =   630
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   75
         Width                    =   1995
         Top                      =   975
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
         Caption                  =   "Valor das Peças..........:"
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
         MaxLength                =   11
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   975
         TabStop                  =  0
         Index                    =   2 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   75
         Width                    =   2025
         Top                      =   1350
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
         Caption                  =   "Valor dos Conjuntos....:"
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
         MaxLength                =   11
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   1350
         TabStop                  =  0
         Index                    =   3 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   75
         Width                    =   2025
         Top                      =   1710
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
         Caption                  =   "Valor dos Serviços......:"
         Index                    =   4 
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
         MaxLength                =   10
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   1710
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
         Left                     =   75
         Width                    =   1950
         Top                      =   2055
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
         Caption                  =   "Valor do Frete..............:"
         Index                    =   5 
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
         MaxLength                =   12
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   2055
         Index                    =   5 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   75
         Width                    =   2040
         Top                      =   2430
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
         Caption                  =   "Valor do Desconto.......:"
         Index                    =   6 
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
         MaxLength                =   11
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   6
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   2430
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
         Left                     =   75
         Width                    =   2055
         Top                      =   2850
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
         Caption                  =   "Valor Total do Projeto.:"
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
         MaxLength                =   12
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2340
         Width                    =   1470
         Top                      =   2835
         TabStop                  =  0
         Index                    =   7 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   9960
         Width                    =   45
         Top                      =   6840
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
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3990
         Width                    =   2745
         Top                      =   975
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
         Caption                  =   "Local de Entrega..........................:"
         Index                    =   9 
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
         MaxLength                =   120
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   12
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6855
         Width                    =   2925
         Top                      =   975
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3990
         Width                    =   2745
         Top                      =   1335
         Alignment                =   1
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
         Caption                  =   "Prazo de Entrega Previsto........:"
         Index                    =   10 
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
         MaxLength                =   120
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   13
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6855
         Width                    =   2925
         Top                      =   1335
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3975
         Width                    =   2790
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
         Caption                  =   "Responsavel Entrega Tecnica:"
         Index                    =   11 
      End
      Begin VB.Label labFdo10
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6855
         Top                      =   1680
         Width                    =   2925
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo10
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   9510
         Top                      =   1695
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
         TabIndex                 =   14
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6870
         Width                    =   2625
         Top                      =   1695
         Index                    =   10 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3975
         Width                    =   2805
         Top                      =   195
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
         Caption                  =   "Revenda.........................................:"
         Index                    =   12 
      End
      Begin VB.Label labFdo11
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6855
         Top                      =   195
         Width                    =   1860
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo11
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8445
         Top                      =   210
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
         TabIndex                 =   8
         TabStop                  =   0
         Height                   =   270
         Left                     =   6870
         Width                    =   1560
         Top                      =   210
         Index                    =   11 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3975
         Width                    =   2730
         Top                      =   615
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
         Caption                  =   "Vend Rebiul...................................:"
         Index                    =   13 
      End
      Begin VB.Label labFdo12
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6855
         Top                      =   615
         Width                    =   1860
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo12
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8445
         Top                      =   630
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
         TabIndex                 =   10
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6870
         Width                    =   1560
         Top                      =   630
         Index                    =   12 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8835
         Width                    =   285
         Top                      =   615
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
         Caption                  =   "%:"
         Index                    =   14 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8820
         Width                    =   285
         Top                      =   195
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
         Caption                  =   "%:"
         Index                    =   15 
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
         MaxLength                =   8
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   9
         TabStop                  =   0
         Height                   =   300
         Left                     =   9135
         Width                    =   645
         Top                      =   195
         Index                    =   13 
         MultiLine                =   -1
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
         MaxLength                =   8
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   11
         TabStop                  =   -1
         Height                   =   300
         Left                     =   9135
         Width                    =   645
         Top                      =   615
         Index                    =   14 
         MultiLine                =   -1
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   15
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   405
         Left                     =   180
         Width                    =   3630
         Top                      =   4440
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   -1
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Salvar Dados"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   390
         Left                     =   3975
         Width                    =   5820
         Top                      =   2085
         Alignment                =   2
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   14.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &HFFFFFF
         Caption                  =   "Parcelamento"
         Index                    =   16 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   16
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   405
         Left                     =   180
         Width                    =   3630
         Top                      =   3960
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Caption                  =   "Parcelamento"
         Index                    =   1 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   17
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   405
         Left                     =   180
         Width                    =   3630
         Top                      =   3495
         CaptionAlign             =   0
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Caption                  =   "Atualizar Produtos"
         Index                    =   2 
      End
      Begin IRRIG.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
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
         EditBackColor            =   8454143
         SelectBackColor          =   -2147483635
         StripesBackColor         =   13750737
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
         TabIndex                 =   18
         TabStop                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   2385
         Left                     =   3975
         Width                    =   5820
         Top                      =   2475
         Caption                  =   ""
         ShowGridCaption          =   0
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3960
         Width                    =   600
         Top                      =   4875
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
         Caption                  =   "Status:"
         Index                    =   17 
      End
      Begin VB.Shape Formato
         FillColor                =   255
         BackColor                =   &HFF
         FillStyle                =   1
         Shape                    =   1
         BackStyle                =   1
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         Height                   =   165
         Left                     =   4635
         Width                    =   165
         Top                      =   4920
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4860
         Width                    =   990
         Top                      =   4875
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
         Caption                  =   "Incompleto"
         Index                    =   18 
      End
      Begin VB.Shape Formato
         FillColor                =   255
         BackColor                =   &H80FFFF
         FillStyle                =   1
         Shape                    =   1
         BackStyle                =   1
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         Height                   =   165
         Left                     =   6015
         Width                    =   165
         Top                      =   4920
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6255
         Width                    =   1050
         Top                      =   4875
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
         Caption                  =   "S/ Parcelas"
         Index                    =   19 
      End
      Begin VB.Shape Formato
         FillColor                =   255
         BackColor                =   &H0
         FillStyle                =   1
         Shape                    =   1
         BackStyle                =   1
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         Height                   =   150
         Left                     =   7410
         Width                    =   150
         Top                      =   4920
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7695
         Width                    =   300
         Top                      =   4875
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
         Caption                  =   "Ok."
         Index                    =   20 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   300
         Left                     =   255
         Width                    =   1455
         Top                      =   5055
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Observações:"
         Index                    =   21 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   19
         TabStop                  =   -1
         Height                   =   1230
         Left                     =   240
         Width                    =   9570
         Top                      =   5385
         Index                    =   15 
         MultiLine                =   -1
         ScrollBars               =   2
         MaxLength                =   0
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
         TabIndex                 =   20
         TabStop                  =   -1
         Height                   =   285
         Left                     =   240
         Width                    =   9495
         Top                      =   5505
         Index                    =   16 
         MultiLine                =   -1
         ScrollBars               =   2
         MaxLength                =   0
      End
   End
End
Attribute VB_name = "frmFinanc"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: Financ
'* Função....: Financeiro
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 15:01:03
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

Dim txtCampo(16) As New FormataCampos             'classe dos compos tipo texto do módulo  
Dim Seq_Projeto As Double, Forma_de_Pagamento As String, Valor_Total_das_Pecas As Double
Dim Valor_Total_dos_Conjuntos As Double, Total_dos_Servicos As Double, Valor_do_Frete As Double
Dim Valor_do_Desconto As Double, Valor_Total_do_Projeto As Double, Local_de_Entrega As String
Dim Prazo_de_Entrega_Previsto As String, Entrega_Tecnica As String, Vendedor_Intermediario As String
Dim Rebiut As String, Percentual_do_Vendedor As Double, Percentual_Rebiut As Double
Dim Obs As String
Public txtSeq_Projeto As Object, txtForma_de_Pagamento As Object, txtValor_Total_das_Pecas As Object
Public txtValor_Total_dos_Conjuntos As Object, txtTotal_dos_Servicos As Object, txtValor_do_Frete As Object
Public txtValor_do_Desconto As Object, txtValor_Total_do_Projeto As Object, lblAjuste As Object
Public txtLocal_de_Entrega As Object, txtPrazo_de_Entrega_Previsto As Object, txtEntrega_Tecnica As Object
Public txtVendedor_Intermediario As Object, txtRebiut As Object, txtPercentual_do_Vendedor As Object
Public txtPercentual_Rebiut As Object, lblParcelamento As Object, gridParcelas As Object
Public txtObs As Object, txtMemoAuxiliar As Object

Private Sub AjustaValores()
   Dim TotalPecas As New GRecordSet, TotalConjuntos As New GRecordSet
   Dim ValorMvto As Currency, TotalServicos As New GRecordSet
  
   On Error GoTo DeuErro
   
 
   Set TotalPecas = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Peças do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
   Set TotalConjuntos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Conjuntos do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
   Set TotalServicos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Serviços do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
                                                                                              
   ValorMvto = TotalPecas!Total + TotalConjuntos!Total + TotalServicos!Total + Valor_do_Frete - Abs(Valor_do_Desconto)
   ValorMvto = Format(ValorMvto , "##,###,##0.00")
   
   'Atualizando
   vgDb.BeginTrans
   'vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total das Peças] = " & Substitui(TotalPecas!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   'vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total dos Conjuntos] = " & Substitui(TotalConjuntos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   'vgDb.Execute "Update [Projeto de Irrigação] Set [Total dos Serviços] = " & Substitui(TotalServicos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   'vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total do Projeto] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Local de Entrega] = '" & Local_de_Entrega & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Prazo de Entrega Previsto] = '" & Prazo_de_Entrega_Previsto & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Entrega Tecnica] = '" & Entrega_Tecnica & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Vendedor Intermediario] = '" & Vendedor_Intermediario & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Percentual do Vendedor] = " & Substitui(txtPercentual_do_Vendedor.Value , ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set Rebiut = '" & Rebiut & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Percentual Rebiut] = " & Substitui(txtPercentual_Rebiut.Value, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Valor do Frete] = " & Substitui(txtValor_do_Frete.Value, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Valor do Desconto] = " & Substitui(txtValor_do_Desconto.Value, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Forma de Pagamento] = '" & Forma_de_Pagamento & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set Obs = '" & Obs & "' WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.CommitTrans
     
   AjustaValores2 
      
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Sub CarregaGrids()
  With gridParcelas
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      '.SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = False
      .AllowEdit = False
      .AllowDelete = False
      '.AddColumn Nothing, , "Pedido", "Sequencia do Projeto", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 600
      .AddColumn Nothing, , "Nº. Pc", "Número da Parcela", TP_NUMERICO, "99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 600
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 450
      .AddColumn Nothing, , "Dt. Vencto.", "Data de Vencimento", TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Vr. Parcela", "Valor da Parcela", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Obs", "Descrição", TP_CARACTER, "", , True, , , , , , , , , , "0", , , , "0", 1, "0", 2200
      .ShowSumBar = True
      .ShowSumCol(.Columns("Número da Parcela").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   CarregaContas
  ' MostraFormulas 
End Sub 


Private Sub Parcelar()
 Dim Parcelas As New GRecordSet
 Dim Restante As Currency
 Dim i As Integer, DiasAux
 Dim VrPcAux As Currency, PriVez As Boolean
 Dim Tb As New GRecordSet, Abateu As Currency
 Dim Projeto As New GRecordSet

 Set Parcelas = VgDb.OpenRecordSet("Parcelas do Projeto")
 Set Tb = VgDb.OpenRecordSet("SELECT * From [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
 Set Projeto = VgDb.OpenRecordSet("SELECT [Valor Total Do Projeto] From [Projeto de Irrigação] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
 
 If Tb.RecordCount > 0 Then
    VgDb.Execute "DELETE FROM [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
 End If
 
 Caixinha:
 SuperInput4
 
 'If QtdParcelaAux > 1 Then
 '   EncargosFinanceiros
 'End If   
 
 i = 1
 DiasAux = 0
 PriVez = False
 
 If ValorEntradaAux = 0 Or QtdParcelaAux = 0 Then Goto Caixinha
 If Projeto![Valor Total Do Projeto] = 0 Then Exit Sub
 
 If ValorEntradaAux > Projeto![Valor Total Do Projeto] Then
    MsgBox("valor de entrada não pode ser superior ao total do projeto!")
    Exit Sub
 End If

 Restante = Projeto![Valor Total Do Projeto] - ValorEntradaAux
 VrPcAux = Restante / QtdParcelaAux
 Abateu = Restante
     
    For i = 1 To QtdParcelaAux + 1
       With Parcelas
        .addNew
           ![Sequencia Do Projeto] = txtSeq_Projeto.Value 
           ![Número da Parcela] = i
           !Dias = DiasAux
            If Not PriVez Then
               ![Data de Vencimento] = Date 
               ![Valor da Parcela] = ValorEntradaAux
               PriVez = True
            Else
               ![Data de Vencimento] = DateAdd("d", DiasAux, Date)
                If I = QtdParcelaAux + 1 Then
                   If Restante - Abateu = 0 Then
                      ![Valor da Parcela] = Restante
                   Else
                      ![Valor da Parcela] = Restante + Abateu
                   End If   
                Else
                   ![Valor da Parcela] = VrPcAux
                    Restante = Restante - VrPcAux
                    Abateu = Abateu - VrPcAux
                End If   
            End If 
        .Update
        .BookMark = .LastModified 
       End With
       DiasAux = DiasAux + 30 
    Next
    VgDb.Execute "DELETE FROM [Parcelas do Projeto] WHERE [Valor da Parcela] = 0 And [Sequencia do Projeto] = " & txtSeq_Projeto.Value 
    gridParcelas.Rebind
    ValorEntradaAux = 0 
    QtdParcelaAux = 0
    GridParcelas.Rebind
End Sub

Private Sub EncargosFinanceiros()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Servicos As New GRecordSet
 Dim Acrescimo As Currency
 Dim Projeto As New GRecordSet
 
 Set Projeto = VgDb.OpenRecordSet("SELECT [Gerou Encargos] From [Projeto de Irrigação] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
 
 If Projeto![Gerou Encargos] Then
     If MsgBox("ATENÇÃO! Os encargos financeiros referente a parcelamento ao Cliente ja foram incluidos nesse Projeto deseja voltar os Valores Originais?", vbYesNo + vbQuestion, vaTitulo) = vbYes Then 
        'ValorOriginal
        Exit Sub
     Else
        Exit Sub     
     End If   
 End If
 
 If MsgBox("ATENÇÃO! esse processo vai atualizar o valor dos produtos incluindo os encargos financeiros referente a parcelamento ao cliente, caso não for parcelar esse pedido clique em não?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Peças do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
 Set Conjuntos = VgDb.OpenRecordSet("SELECT * From [Conjuntos do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
 Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value)

  Do While Not Produtos.EOF
     With Produtos
       .Edit
         ![Valor Anterior] = ![Valor Unitário]
          Acrescimo = Round((![Valor Unitário] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100),2)
         ![Valor Unitário] = ![Valor Unitário] + Acrescimo
         ![Valor Total] = ![Valor Unitário] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
     Produtos.MoveNext 
  Loop 
  'Conjuntos
   Do While Not Conjuntos.EOF
     With Conjuntos
       .Edit
         ![Valor Anterior] = ![Valor Unitário]
          Acrescimo = Round((![Valor Unitário] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100),2)
         ![Valor Unitário] = ![Valor Unitário] + Acrescimo
         ![Valor Total] = ![Valor Unitário] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
      Conjuntos.MoveNext
  Loop
  'Servicos
   Do While Not Servicos.EOF
     With Servicos
       .Edit
         ![Valor Anterior] = ![Valor Unitario]
          Acrescimo = Round((![Valor Unitario] * Parametros_do_Produto![Acrescimo Do Parcelamento] / 100),2)
         ![Valor Unitario] = ![Valor Unitario] + Acrescimo
         ![Valor Total] = ![Valor Unitario] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
     Servicos.MoveNext
  Loop 
   VgDb.Execute("UPDATE [Projeto de Irrigação] Set [Gerou Encargos] = 1 WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
   AjustaValores2
   MsgBox("Encargos Financeiros Acrescentados com Sucesso!")
End Sub


Private Sub AtualizaValor()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Tb As New GRecordSet 
  
 If MsgBox("ATENÇÃO! esse processo vai atualizar o valor dos produtos?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Peças do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
 Set Conjuntos = VgDb.OpenRecordSet("SELECT * From [Conjuntos do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 

  Do While Not Produtos.EOF
   Set Tb = vgDb.OpenRecordSet("SELECT [Valor Total] From Produtos Where [Seqüência do Produto] = " & Produtos![Sequencia Do Produto])
    If Tb![Valor Total] > 0 Then
       With Produtos
       .Edit
         ![Valor Unitário] = Tb![Valor Total]
         ![Valor Total] = ![Valor Unitário] * !Quantidade
       .Update
       .BookMark = .LastModified
       End With
    End If   
     Produtos.MoveNext 
  Loop 
  'Conjuntos
  Do While Not Conjuntos.EOF
     Set Tb = vgDb.OpenRecordSet("SELECT [Valor Total] From Conjuntos Where [Seqüência do Conjunto] = " & Conjuntos![Sequencia Do Conjunto])
     If Tb![Valor Total] > 0 Then
        With Conjuntos
         .Edit
          ![Valor Unitário] = Tb![Valor Total]
          ![Valor Total] = ![Valor Unitário] * !Quantidade
         .Update
         .BookMark = .LastModified
        End With
     End If   
     Conjuntos.MoveNext
  Loop
  
  AjustaValores2
  MsgBox("Produtos Atualizado com Sucesso!")

End Sub


Private Sub ValorOriginal()
 Dim Produtos As New GRecordSet
 Dim Conjuntos As New GRecordSet
 Dim Servicos As New GRecordSet
 Dim Acrescimo As Currency
 Dim Projeto As New GRecordSet
 
 Set Projeto = VgDb.OpenRecordSet("SELECT [Gerou Encargos] From [Projeto de Irrigação] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
 
  If Projeto![Gerou Encargos] Then 
     If MsgBox("ATENÇÃO! Deseja Remover os Encargos Financeiros Referente a Parcelamento ao Cliente?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then Exit Sub
  End If
  
  Set Produtos = VgDb.OpenRecordSet("SELECT * From [Peças do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
  Set Conjuntos = VgDb.OpenRecordSet("SELECT * From [Conjuntos do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
  Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços do Projeto] Where [Sequencia do Projeto] = " & txtSeq_Projeto.Value)

  Do While Not Produtos.EOF
   If Produtos![Valor Anterior] > 0 Then
     With Produtos
       .Edit
         ![Valor Unitário] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitário] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
   End If  
     Produtos.MoveNext 
  Loop 
  'Conjuntos
   Do While Not Conjuntos.EOF
    If Conjuntos![Valor Anterior] > 0 Then
     With Conjuntos
       .Edit
         ![Valor Unitário] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitário] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
    End If 
      Conjuntos.MoveNext
  Loop
  'Servicos
   Do While Not Servicos.EOF
    If Servicos![Valor Anterior] > 0 Then
     With Servicos
       .Edit
         ![Valor Unitario] = ![Valor Anterior]
         ![Valor Total] = ![Valor Unitario] * !Quantidade
       .Update
       .BookMark = .LastModified
     End With
    End If  
     Servicos.MoveNext
  Loop 
   VgDb.Execute("UPDATE [Projeto de Irrigação] Set [Gerou Encargos] = 0 WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
   AjustaValores2
   MsgBox("Encargos Financeiros Removidos com Sucesso!")
   
End Sub


Private Sub AjustaValores2()
   Dim TotalPecas As New GRecordSet, TotalConjuntos As New GRecordSet
   Dim ValorMvto As Currency, TotalServicos As New GRecordSet
  
   On Error GoTo DeuErro
   
 
   Set TotalPecas = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Peças do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value)
   Set TotalConjuntos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Conjuntos do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
   Set TotalServicos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Valor Total]), 2) Total FROM [Serviços do Projeto] WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value) 
   
                                                                                              
   ValorMvto = TotalPecas!Total + TotalConjuntos!Total + TotalServicos!Total + Valor_do_Frete - Abs(Valor_do_Desconto)
   ValorMvto = Format(ValorMvto , "##,###,##0.00")
   
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total das Peças] = " & Substitui(TotalPecas!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total dos Conjuntos] = " & Substitui(TotalConjuntos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Total dos Serviços] = " & Substitui(TotalServicos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " & txtSeq_Projeto.Value
   vgDb.Execute "Update [Projeto de Irrigação] Set [Valor Total do Projeto] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Sequencia do Projeto] = " &  txtSeq_Projeto.Value
   vgDb.CommitTrans
       
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
   Financeiro
   MostraFormulas
End Sub


Private Sub CarregaContas()
   Dim x As String, Tb As New GRecordSet, MsgValida As String
   
   'x$ = ""
   On Error GoTo DeuErro
   
      x$ = "SELECT [Número da Parcela], Dias, " & _
           "[Data de Vencimento], [Valor da Parcela], Descrição FROM [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value
  
   Set Tb = vgDb.OpenRecordSet(x$)
   
   With gridParcelas
      .OpenRecordSet x$, CURSOR_TABLE
   End With
   
   'If Tb.RecordCount > 0 Then TemRecordset = True
      MsgValida$ = Executar(VALIDACOES)
   If MsgValida$ <> "" Then MsgBox MsgValida$, vbCritical + vbOKOnly, vaTitulo: Exit Sub
      
DeuErro:
   If Err Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub   


Private Sub Financeiro()
 Dim Tb As New GRecordSet
 
   Set Tb = VgDb.OpenRecordSet("SELECT [Forma de Pagamento], Obs, [Valor do Frete], [Valor do Desconto], [Valor Total das Peças], [Valor Total dos Conjuntos], [Total dos Serviços], [Valor Total do Projeto], [Vendedor Intermediario], [Percentual do Vendedor], [Local de Entrega], Rebiut, [Percentual Rebiut], [Prazo de Entrega Previsto], [Entrega Tecnica] From [Projeto de Irrigação] where [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value) 
     If Tb![Valor Do Frete] > 0 Then  
        txtValor_do_Frete.Value = Tb![Valor Do Frete]
     Else
        frmFinanc.txtValor_do_Frete.Value = 0  
     End If   
     If Tb![Valor Do Desconto] < 0 Then
        txtValor_do_desconto.Value = Tb![Valor Do Desconto]
     Else
        frmFinanc.txtValor_do_desconto.Value = 0  
     End If
     If Tb![Valor Total das Peças] Then
        txtValor_Total_das_Pecas.Value = Tb![Valor Total das Peças]
     End If
     If Tb![Valor Total dos Conjuntos] > 0 Then
        txtValor_Total_dos_Conjuntos.Value = Tb![Valor Total dos Conjuntos] 
     End If
     If Tb![Total dos Serviços] > 0 Then
        txtTotal_dos_Servicos.Value = Tb![Total dos Serviços]
     Else
        txtTotal_dos_Servicos.Value = 0  
     End If
     If Tb![Valor Total Do Projeto] > 0 Then
        txtValor_Total_Do_Projeto.Value = Tb![Valor Total Do Projeto]
     End If
     If Not Vazio(Tb![Vendedor Intermediario]) Then
        txtVendedor_Intermediario.Value = Tb![Vendedor Intermediario]
     End If
     If Tb![Percentual Do Vendedor] > 0 Then
        txtPercentual_Do_Vendedor.Value = Tb![Percentual Do Vendedor]
     End If
     If Not Vazio (Tb!Rebiut) Then
        txtRebiut.Value = Tb!Rebiut 
     End If
     If Tb![Percentual Rebiut] > 0 Then
        txtPercentual_Rebiut.Value = Tb![Percentual Rebiut]
     End If
     If Not Vazio(Tb![Local de Entrega]) Then
        txtLocal_de_Entrega.Value = Tb![Local de Entrega]
     End If
     If Not Vazio(Tb![Prazo de Entrega Previsto]) Then
        txtPrazo_de_Entrega_Previsto.Value = Tb![Prazo de Entrega Previsto]
     End If
     If Not Vazio(Tb![Entrega Tecnica]) Then
        txtEntrega_Tecnica.Value = Tb![Entrega Tecnica]
     End If
     If Not Vazio(Tb!Obs) Then
        txtObs.Value = Tb!Obs   
     End If
         
End Sub


Private Sub DeletaParcela()
 Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("SELECT * From [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value)
 
 If Tb.RecordCount > 0 Then
    VgDb.Execute "DELETE FROM [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value
 End If
 
End Sub

Private Sub MostraFormulas()
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < txtValor_Total_Do_Projeto.Value) Or (TotalParcelas > txtValor_Total_Do_Projeto.Value) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If     
End Sub   


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da Parcela]) As Total From [Parcelas do Projeto] WHERE [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value & " AND " & IIf(Seq > 0, "[Número da Parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function



Private Function ValidaDesconto() 
 Dim DescAux As Double
 Dim Tb As New GRecordSet
 
 If Valor_do_Desconto = 0 Then
    Exit Function
 End If   
 If (Valor_Total_do_Projeto = 0) Or vgPWUsuario = "WAGNER" Or vgPWUsuario = "ALEXANDRA" Or vgPWUsuario = "YGOR" Then 
    Exit Function
 End If
 
 Set Tb = VgDb.OpenRecordSet("SELECT [Valor Total das Peças] + [Valor Total dos Conjuntos] As Total FROM [Projeto de Irrigação] WHERE [Sequencia do Projeto] = " & frmNovoPrg.txtSequencia_do_Projeto.Value) 
 
 DescAux = Tb!Total * 3 / 100

 If ABS(Valor_do_Desconto) > DescAux Then
    MsgBox("Desconto Invalido! maior que 3%")
    txtValor_do_Desconto.Value = 0
    AjustaValores
    Exit Function
 End If
  
End Function



'Rotina para limitar o campo memo
Private Sub LimitaCampo(KeyAscii, MaxLength)
   On Error Resume Next
   
   If KeyAscii < 32 And KeyAscii <> 22 And KeyAscii <> 13 Then Exit Sub
   
   ' Fazendo uso de campo auxiliar que tem o tamanho do memo do relatorio
   txtMemoAuxiliar.Text = txtObs.Text
   If SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) > 5 Or (SendMessage(txtMemoAuxiliar.hWnd, &HBA, 0, 0) = 5 And KeyAscii = 13) Then
      Beep
      KeyAscii = 0
   End If

End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   Seq_Projeto = txtCampo(0).Value
   Forma_de_Pagamento = txtCampo(1).Value
   Valor_Total_das_Pecas = txtCampo(2).Value
   Valor_Total_dos_Conjuntos = txtCampo(3).Value
   Total_dos_Servicos = txtCampo(4).Value
   Valor_do_Frete = txtCampo(5).Value
   Valor_do_Desconto = txtCampo(6).Value
   Valor_Total_do_Projeto = txtCampo(7).Value
   Local_de_Entrega = txtCampo(8).Value
   Prazo_de_Entrega_Previsto = txtCampo(9).Value
   Entrega_Tecnica = txtCampo(10).Value
   Vendedor_Intermediario = txtCampo(11).Value
   Rebiut = txtCampo(12).Value
   Percentual_do_Vendedor = txtCampo(13).Value
   Percentual_Rebiut = txtCampo(14).Value
   Obs = txtCampo(15).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 11
         txtCampo(11).Filter = "Desativado = " & 0 & ""
      Case 12
         txtCampo(12).Filter = "Desativado = " & 0 & ""
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
      vgOk = (Valor_do_Frete >= 0)
      vgMsg$ = "Valor do Frete inválido!"
      If Not vgOk Then txtCampo(5).SetFocus
      If vgOk Then
         vgOk = (Valor_do_Desconto <= 0)
         vgMsg$ = "O Valor do Desconto Tem que Ser Menor que 0 ex: ( -100 )"
         If Not vgOk Then txtCampo(6).SetFocus
      End If
      If vgOk Then
         If Not Vazio(Vendedor_Intermediario) Then
            vgOk = ((Percentual_do_Vendedor > 0) AND Percentual_do_Vendedor <= 12)
            vgMsg$ = "Percentual do Vendedor Invalido!"
            If Not vgOk Then txtCampo(13).SetFocus
         End If
      End If
      If vgOk Then
         If Not Vazio(Rebiut) Then
            vgOk = ((Percentual_Rebiut > 0) AND Percentual_Rebiut <= 1)
            vgMsg$ = "Percentual do Vendedor Rebiut não pode ser 0 nem maior que 1%"
            If Not vgOk Then txtCampo(14).SetFocus
         End If
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
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
   Ok = (Not Vazio(Vendedor_Intermediario))
   txtCampo(13).Enabled = (Ok And txtCampo(13).Editable)
   Ok = (Not Vazio(Rebiut))
   txtCampo(14).Enabled = (Ok And txtCampo(14).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(16).Visible = (False)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 1 Or Index = 10 Or Index = 11 Or Index = 12 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 11
            PoeRelEFiltroCbo 11
         Case 12
            PoeRelEFiltroCbo 12
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
      Case 6
         ValidaDesconto  
   End Select
End Sub


'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(63220)
   vgFormID = 1395
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set txtSeq_Projeto = txtCampo(0)
   Set txtForma_de_Pagamento = txtCampo(1)
   Set txtValor_Total_das_Pecas = txtCampo(2)
   Set txtValor_Total_dos_Conjuntos = txtCampo(3)
   Set txtTotal_dos_Servicos = txtCampo(4)
   Set txtValor_do_Frete = txtCampo(5)
   Set txtValor_do_Desconto = txtCampo(6)
   Set txtValor_Total_do_Projeto = txtCampo(7)
   Set lblAjuste = Label(8)
   Set txtLocal_de_Entrega = txtCampo(8)
   Set txtPrazo_de_Entrega_Previsto = txtCampo(9)
   Set txtEntrega_Tecnica = txtCampo(10)
   Set txtVendedor_Intermediario = txtCampo(11)
   Set txtRebiut = txtCampo(12)
   Set txtPercentual_do_Vendedor = txtCampo(13)
   Set txtPercentual_Rebiut = txtCampo(14)
   Set lblParcelamento = Label(16)
   Set gridParcelas = Grid(0)
   Set txtObs = txtCampo(15)
   Set txtMemoAuxiliar = txtCampo(16)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(63225)
   Label(1).Caption = LoadGasString(63226)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(63227)
   Label(2).Caption = LoadGasString(63228)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(63229)
   Label(3).Caption = LoadGasString(63230)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(63231)
   Label(4).Caption = LoadGasString(63232)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(63233)
   Label(5).Caption = LoadGasString(63234)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(63235)
   Label(6).Caption = LoadGasString(63236)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(63237)
   Label(7).Caption = LoadGasString(63238)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(63239)
   Label(8).Caption = LoadGasString(63240)
   Label(9).Caption = LoadGasString(63241)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(63242)
   Label(10).Caption = LoadGasString(63243)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(63244)
   Label(11).Caption = LoadGasString(63245)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(63246)
   Label(12).Caption = LoadGasString(63247)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(63248)
   Label(13).Caption = LoadGasString(63249)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(63250)
   Label(14).Caption = LoadGasString(63251)
   Label(15).Caption = LoadGasString(63252)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(63253)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(63254)
   Botao(0).Caption = LoadGasString(63255)
   Botao(1).Caption = LoadGasString(63256)
   Botao(2).Caption = LoadGasString(63257)
   Label(17).Caption = LoadGasString(63258)
   Label(18).Caption = LoadGasString(63259)
   Label(19).Caption = LoadGasString(63260)
   Label(20).Caption = LoadGasString(63261)
   Label(21).Caption = LoadGasString(63262)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
   CarregaGrids  
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
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "999999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 0
   txtCampo(1).ListFields = "Vista|Prazo|Antecipado"

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "99.999.999,99"
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "99.999.999,99"
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "9.999.999,99"
   txtCampo(4).Editable = False
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "9.999.999,9999"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = "99.999.999,99"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "999.999.999,99"
   txtCampo(7).Editable = False
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""

   Set txtCampo(11).CtPri = txtCp(11)
   Set txtCampo(11).CtFdo = labFdo11
   Set txtCampo(11).CtBot(BOT_COMBO) = bottxtCampo11(BOT_COMBO)
   Set bottxtCampo11(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(11).DataType = 0
   txtCampo(11).Mask = "@x"
   txtCampo(11).BoundColumn = "Titular da Conta"
   txtCampo(11).ListFields = "Titular da Conta"
   txtCampo(11).OrderFields = "Titular da Conta"
   txtCampo(11).Relation = ""
   txtCampo(11).Source = "Conta do Vendedor"
   txtCampo(11).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 1
   txtCampo(13).Mask = "999,9999"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""

   Set txtCampo(12).CtPri = txtCp(12)
   Set txtCampo(12).CtFdo = labFdo12
   Set txtCampo(12).CtBot(BOT_COMBO) = bottxtCampo12(BOT_COMBO)
   Set bottxtCampo12(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(12).DataType = 0
   txtCampo(12).Mask = "@x"
   txtCampo(12).BoundColumn = "Titular da Conta"
   txtCampo(12).ListFields = "Titular da Conta"
   txtCampo(12).OrderFields = "Titular da Conta"
   txtCampo(12).Relation = ""
   txtCampo(12).Source = "Conta do Vendedor"
   txtCampo(12).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(14).CtPri = txtCp(14)
   txtCampo(14).DataType = 1
   txtCampo(14).Mask = "999,9999"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "@x"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = "@x"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""

   Set txtCampo(10).CtPri = txtCp(10)
   Set txtCampo(10).CtFdo = labFdo10
   Set txtCampo(10).CtBot(BOT_COMBO) = bottxtCampo10(BOT_COMBO)
   Set bottxtCampo10(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(10).DataType = 0
   txtCampo(10).ListFields = "Irrigação Penapolis|Revenda"

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 4
   txtCampo(15).Mask = ""
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""

   Set txtCampo(16).CtPri = txtCp(16)
   txtCampo(16).DataType = 4
   txtCampo(16).Mask = ""
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""

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
   AjustaValores  
   FinalizaForm Me
   Set txtSeq_Projeto = Nothing
   Set txtForma_de_Pagamento = Nothing
   Set txtValor_Total_das_Pecas = Nothing
   Set txtValor_Total_dos_Conjuntos = Nothing
   Set txtTotal_dos_Servicos = Nothing
   Set txtValor_do_Frete = Nothing
   Set txtValor_do_Desconto = Nothing
   Set txtValor_Total_do_Projeto = Nothing
   Set lblAjuste = Nothing
   Set txtLocal_de_Entrega = Nothing
   Set txtPrazo_de_Entrega_Previsto = Nothing
   Set txtEntrega_Tecnica = Nothing
   Set txtVendedor_Intermediario = Nothing
   Set txtRebiut = Nothing
   Set txtPercentual_do_Vendedor = Nothing
   Set txtPercentual_Rebiut = Nothing
   Set lblParcelamento = Nothing
   Set gridParcelas = Nothing
   Set txtObs = Nothing
   Set txtMemoAuxiliar = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmFinanc = Nothing                        'libera o segmento de código do form
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
         AjustaValores  
      Case 1
         Parcelar  
      Case 2
         AtualizaValor  
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo11_Click(Index As Integer)
   txtCampo(11).SetFocus
   DoEvents
   txtCampo(11).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo12_Click(Index As Integer)
   txtCampo(12).SetFocus
   DoEvents
   txtCampo(12).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo10_Click(Index As Integer)
   txtCampo(10).SetFocus
   DoEvents
   txtCampo(10).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

