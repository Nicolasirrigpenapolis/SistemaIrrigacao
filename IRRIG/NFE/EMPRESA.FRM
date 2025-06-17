Version 5.00
Begin VB.Form frmEmpresa
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   6810
   Left                     =   3630
   LinkTopic                =   "Empresa"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   6375
   ScaleWidth               =   8730
   Top                      =   1750
   Width                    =   8820
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
      Height                   =   6210
      Left                     =   15
      Width                    =   8550
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   15073
      _extenty                 =   10948
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   1
         BackColor                =   &H8000000F
         BorderStyle              =   1
         Alignment                =   0
         Height                   =   1335
         Left                     =   120
         Width                    =   1410
         Top                      =   60
         DataField                =   "LogoTipo"
         Index                    =   0 
         TabStop                  =   0
         Stretch                  =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   1335
         Top                      =   1680
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
         ForeColor                =   &H0
         Caption                  =   "*Razão Social:"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   660
         Width                    =   870
         Top                      =   2160
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
         ForeColor                =   &H0
         Caption                  =   "Fantasia:"
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
         ForeColor                =   &H0
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   6570
         Top                      =   2160
         DataField                =   "Nome fantasia"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   960
         Width                    =   570
         Top                      =   2640
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
         ForeColor                =   &H0
         Caption                  =   "*CPF:"
         Index                    =   2 
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
         ForeColor                =   &H0
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   1530
         Top                      =   2640
         DataField                =   "Cpf do emitente"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3360
         Width                    =   630
         Top                      =   2640
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
         ForeColor                =   &H0
         Caption                  =   "*CNPJ:"
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
         ForeColor                =   &H0
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4140
         Width                    =   1800
         Top                      =   2640
         DataField                =   "Cnpj"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6060
         Width                    =   390
         Top                      =   2640
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
         ForeColor                =   &H0
         Caption                  =   "*IE:"
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
         ForeColor                =   &H0
         MaxLength                =   20
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6510
         Width                    =   1740
         Top                      =   2640
         DataField                =   "Inscrição estadual"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   1080
         Width                    =   510
         Top                      =   3120
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
         ForeColor                =   &H0
         Caption                  =   "Rua:"
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
         ForeColor                =   &H0
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   5190
         Top                      =   3120
         DataField                =   "Rua"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6960
         Width                    =   270
         Top                      =   3120
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
         ForeColor                =   &H0
         Caption                  =   "Nº"
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
         ForeColor                =   &H0
         MaxLength                =   8
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   6
         TabStop                  =   -1
         Height                   =   300
         Left                     =   7260
         Width                    =   990
         Top                      =   3120
         DataField                =   "Nro"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   960
         Width                    =   630
         Top                      =   3660
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
         ForeColor                =   &H0
         Caption                  =   "Bairro:"
         Index                    =   7 
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
         ForeColor                =   &H0
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   2190
         Top                      =   3660
         DataField                =   "Bairro"
         Index                    =   6 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3960
         Width                    =   960
         Top                      =   3660
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
         ForeColor                =   &H0
         Caption                  =   "Município:"
         Index                    =   8 
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
         ForeColor                =   &H0
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4980
         Width                    =   2100
         Top                      =   3660
         DataField                =   "Município"
         Index                    =   7 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7200
         Width                    =   330
         Top                      =   3660
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
         ForeColor                =   &H0
         Caption                  =   "UF:"
         Index                    =   9 
      End
      Begin VB.Label labFdo8
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7650
         Top                      =   3660
         Width                    =   600
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7980
         Top                      =   3675
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
         ForeColor                =   &H0
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   9
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7665
         Width                    =   300
         Top                      =   3675
         DataField                =   "Uf"
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   1080
         Width                    =   495
         Top                      =   4140
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
         ForeColor                =   &H0
         Caption                  =   "CEP:"
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
         ForeColor                =   &H0
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   10
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   1140
         Top                      =   4140
         DataField                =   "Cep"
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3000
         Width                    =   1920
         Top                      =   4140
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
         ForeColor                =   &H0
         Caption                  =   "Codigo do Município:"
         Index                    =   11 
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
         ForeColor                =   &H0
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   11
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4980
         Width                    =   885
         Top                      =   4140
         DataField                =   "Codigo do municipio"
         Index                    =   10 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6240
         Width                    =   600
         Top                      =   4140
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
         ForeColor                =   &H0
         Caption                  =   "Fone:"
         Index                    =   12 
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
         ForeColor                =   &H0
         MaxLength                =   13
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   12
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6855
         Width                    =   1380
         Top                      =   4140
         DataField                =   "Fone"
         Index                    =   11 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   13
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   300
         Left                     =   5880
         Width                    =   225
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
         Index                    =   0 
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
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   6570
         Top                      =   1680
         DataField                =   "Razão social"
         Index                    =   12 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   4620
         Width                    =   405
         Top                      =   660
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
         Caption                  =   "Pis:"
         Index                    =   13 
      End
      Begin VB.Label labFdo13
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   5100
         Top                      =   660
         Width                    =   1125
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo13
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   5955
         Top                      =   675
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
         Alignment                =   1
         TabIndex                 =   14
         TabStop                  =   0
         Height                   =   270
         Left                     =   5115
         Width                    =   825
         Top                      =   675
         DataField                =   "Aliquota do pis"
         Index                    =   13 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6360
         Width                    =   705
         Top                      =   660
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
         Caption                  =   "Cofins:"
         Index                    =   14 
      End
      Begin VB.Label labFdo14
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7080
         Top                      =   660
         Width                    =   1125
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   7935
         Top                      =   675
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
         Alignment                =   1
         TabIndex                 =   15
         TabStop                  =   0
         Height                   =   270
         Left                     =   7095
         Width                    =   825
         Top                      =   675
         DataField                =   "Aliquota do cofins"
         Index                    =   14 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   1680
         Width                    =   780
         Top                      =   1200
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
         Caption                  =   "Licença:"
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
         MaxLength                =   130
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   16
         TabStop                  =   -1
         Height                   =   300
         Left                     =   2520
         Width                    =   5699
         Top                      =   1200
         DataField                =   "Chave flexdocs"
         Index                    =   15 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8160
         Width                    =   45
         Top                      =   6360
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
         Index                    =   16 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   480
         Width                    =   1035
         Top                      =   4620
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
         Caption                  =   "Email NFe:"
         Index                    =   17 
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
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   17
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   3735
         Top                      =   4620
         DataField                =   "Email"
         Index                    =   16 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   5625
         Width                    =   1215
         Top                      =   4620
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
         Caption                  =   "Ult Ativação:"
         Index                    =   18 
      End
      Begin VB.Label labFdo17
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6855
         Top                      =   4635
         Width                    =   1380
         Height                   =   285
      End
      Begin VB.CommandButton bottxtCampo17
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   7965
         Top                      =   4650
         Width                    =   255
         Height                   =   255
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
         TabIndex                 =   18
         TabStop                  =   -1
         Height                   =   255
         Left                     =   6870
         Width                    =   1080
         Top                      =   4650
         DataField                =   "Data Inicial da Ativação"
         Index                    =   17 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   420
         Width                    =   1095
         Top                      =   5100
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
         Caption                  =   "Senha Adm:"
         Index                    =   19 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0C0FF
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
         TabIndex                 =   19
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   1020
         Top                      =   5100
         DataField                =   "Senha Admin"
         Index                    =   18 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2820
         Width                    =   1425
         Top                      =   5100
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
         Caption                  =   "Aliquota do ISS:"
         Index                    =   20 
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
         TabIndex                 =   20
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4380
         Width                    =   660
         Top                      =   5100
         DataField                =   "Aliquota do Iss"
         Index                    =   19 
         MultiLine                =   -1
      End
      Begin VB.CheckBox ChkCp
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
         ForeColor                =   &H80000008
         Alignment                =   0
         TabIndex                 =   21
         TabStop                  =   0
         Height                   =   255
         Left                     =   2820
         Width                    =   1980
         Top                      =   5820
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
         Alignment                =   0
         Caption                  =   "Baixa Pagar Auto"
         Index                    =   0 
         DataField                =   "Baixa Pagar Auto"
      End
      Begin VB.CheckBox ChkCp
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
         ForeColor                =   &H80000008
         Alignment                =   0
         TabIndex                 =   22
         TabStop                  =   0
         Height                   =   255
         Left                     =   2820
         Width                    =   2220
         Top                      =   5520
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
         Alignment                =   0
         Caption                  =   "Baixa Receber Auto"
         Index                    =   1 
         DataField                =   "Baixa Receber Auto"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   105
         Width                    =   1410
         Top                      =   5580
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
         Caption                  =   "Inscr. Municipal:"
         Index                    =   21 
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
         MaxLength                =   10
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   23
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1680
         Width                    =   1035
         Top                      =   5580
         DataField                =   "Incrição Municipal"
         Index                    =   20 
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
         ForeColor                =   &H0
         Alignment                =   0
         TabIndex                 =   24
         TabStop                  =   0
         Height                   =   255
         Left                     =   6900
         Width                    =   1395
         Top                      =   120
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Alignment                =   0
         Caption                  =   "Emiti NFe"
         Index                    =   2 
         DataField                =   "Emite NFe"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2160
         Width                    =   1395
         Top                      =   660
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
         Caption                  =   "% da Redução:"
         Index                    =   22 
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
         TabIndex                 =   25
         TabStop                  =   -1
         Height                   =   300
         Left                     =   3645
         Width                    =   795
         Top                      =   660
         DataField                =   "Percentual da redução"
         Index                    =   21 
         MultiLine                =   -1
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
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
         Alignment                =   0
         TabIndex                 =   26
         TabStop                  =   -1
         Height                   =   255
         Left                     =   2160
         Width                    =   3015
         Top                      =   180
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
         Alignment                =   0
         Caption                  =   "Não Controlar Estoque"
         Index                    =   3 
         DataField                =   "Não Controlar Estoque"
      End
      Begin VB.Frame Frame
         Height                   =   705
         Left                     =   5490
         Width                    =   2775
         Top                      =   5160
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
         Caption                  =   "Impressão do Relatorio"
         Index                    =   0 
         Begin             VB.Label labopcFrame0
            BorderStyle              =   1
            DataField                =   "Impressão do Relatorio"
            Height                   =   120
            Left                     =   915
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcFrame0Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "MS Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   8.25
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   2
            TabIndex                 =   27
            TabStop                  =   -1
            Height                   =   255
            Left                     =   180
            Width                    =   975
            Top                      =   300
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "MS Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   8.25
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   0
            Caption                  =   "Padrão"
            Index                    =   0 
         End
         Begin VB.OptionButton opcFrame0Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "MS Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   8.25
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   2
            TabIndex                 =   28
            TabStop                  =   -1
            Height                   =   255
            Left                     =   1260
            Width                    =   1395
            Top                      =   300
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "MS Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   8.25
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   0
            Caption                  =   "Meia Folha"
            Index                    =   1 
         End
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   29
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   300
         Left                     =   120
         Width                    =   1410
         Top                      =   1380
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
         Caption                  =   "LogoTipo"
         Index                    =   1 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
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
         Alignment                =   0
         TabIndex                 =   31
         TabStop                  =   -1
         Height                   =   255
         Left                     =   5520
         Width                    =   2430
         Top                      =   5880
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
         Alignment                =   0
         Caption                  =   "Numeração Auto"
         Index                    =   4 
         DataField                =   "Numeração Auto"
      End
   End
End
Attribute VB_name = "frmEmpresa"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: Empresa
'* Função....: Parametros do Emitente
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:25
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
Public vgTemCondicoesEsp As Integer               'flag se tem expresão de validação para inclusão/alteração ou exclusão
Public vgEmBrowse As Integer                      'flag se o módulo esta em grade 
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso 
Public vgIndDefault As String                     'indice default do módulo 
Public vgFormID As Long                           'identificador único para o módulo 
Public vgIdentTab As String                       'nome da tabela principal do módulo 
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Dim txtCampo(21) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(4) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim imgCampo(0) As New FormataCampos              'classe dos campos tipo imagem do módulo
Dim opcFrame0(1) As New FormataCampos
Dim Razao_social As String, Nome_fantasia As String, Cpf As String
Dim Inscricao_estadual As String, Cnpj As String, Rua As String
Dim Nro As String, Bairro As String, Municipio As String
Dim Codigo_do_municipio As Long, Uf As String, Cep As String
Dim Fone As String, Aliquota_do_cofins As Double, Chave_flexdocs As String
Dim Email As String, Data_Inicial_da_Ativacao As Variant, Aliquota_do_pis As Double
Dim Senha_Admin As Long, Baixa_Pagar_Auto As Boolean, Baixa_Receber_Auto As Boolean
Dim Aliquota_do_Iss As Single, Incricao_Municipal As String, Emite_NFe As Boolean
Dim Percentual_da_reducao As Single, Nao_Controlar_Estoque As Boolean, Impressao_do_Relatorio As Integer
Dim Numeracao_Auto As Boolean
Public txtIE As Object, lblAjuste As Object

Public MsgCMun As String

'ROTINA PARA ABRIR O SITE DO IBGE E JA TRAZER O CODIGO DO MUNICIPIO INFORMADO
Private Sub ConsultaCodMun(municipio As String)
   On Error Resume Next
  
   If Not Vazio(municipio) Then
      ShellExecute 0&, vbNullString, "http://www.ibge.gov.br/home/geociencias/areaterritorial/area.php?nome=" & _
                                     SuperTiraAcentos(municipio) & "&codigo=&submit.x=0&submit.y=0", vbNullString, vbNullString, SW_SHOWNORMAL     
   Else
      ShellExecute 0&, vbNullString, "http://www.ibge.gov.br/home/geociencias/areaterritorial/area.shtm", vbNullString, vbNullString, SW_SHOWNORMAL
   End If

End Sub

'Valida Código Mun de Acordo com o DV e UF
Public Function ValidaCMun(Codigo As Long, UF As String) As Boolean
   Dim A(6), B(6), C(6), D(6) As Integer   'Vetores
   Dim DV, i, Soma, DVTabela, CodEstado As Integer
   Dim RestoDiv As Double
   Dim UFVect(27, 27) As Variant           'Matriz
   
   ValidaCMun = True: MsgCMun = ""
   If Codigo = 0 Then Exit Function        'Caso o Usuário não souber o Código vamos deixar salvar  
   If Codigo = 9999999 Then Exit Function  'Exterior
   
   'Estados com Sigla e Codigo
   UFVect(0, 0) = 11: UFVect(0, 1) = "RO"    'RONDÔNIA
   UFVect(1, 0) = 12: UFVect(1, 1) = "AC"    'ACRE
   UFVect(2, 0) = 13: UFVect(2, 1) = "AM"    'AMAZONAS
   UFVect(3, 0) = 14: UFVect(3, 1) = "RR"    'RORAIMA
   UFVect(4, 0) = 15: UFVect(4, 1) = "PA"    'PARÁ
   UFVect(5, 0) = 16: UFVect(5, 1) = "AP"    'AMAPÁ
   UFVect(6, 0) = 17: UFVect(6, 1) = "TO"    'TOCANTINS
   UFVect(7, 0) = 21: UFVect(7, 1) = "MA"    'MARANHÃO
   UFVect(8, 0) = 22: UFVect(8, 1) = "PI"    'PIAUÍ
   UFVect(9, 0) = 23: UFVect(9, 1) = "CE"    'CEARÁ
   UFVect(10, 0) = 24: UFVect(10, 1) = "RN"  'RIO GRANDE DO NORTE
   UFVect(11, 0) = 25: UFVect(11, 1) = "PB"  'PARAÍBA
   UFVect(12, 0) = 26: UFVect(12, 1) = "PE"  'PERNAMBUCO
   UFVect(13, 0) = 27: UFVect(13, 1) = "AL"  'ALAGOAS
   UFVect(14, 0) = 28: UFVect(14, 1) = "SE"  'SERGIPE
   UFVect(15, 0) = 29: UFVect(15, 1) = "BA"  'BAHIA
   UFVect(16, 0) = 31: UFVect(16, 1) = "MG"  'MINAS GERAIS
   UFVect(17, 0) = 32: UFVect(17, 1) = "ES"  'ESPÍRITO SANTO
   UFVect(18, 0) = 33: UFVect(18, 1) = "RJ"  'RIO DE JANEIRO
   UFVect(19, 0) = 35: UFVect(19, 1) = "SP"  'SÃO PAULO
   UFVect(20, 0) = 41: UFVect(20, 1) = "PR"  'PARANÁ
   UFVect(21, 0) = 42: UFVect(21, 1) = "SC"  'SANTA CATARINA
   UFVect(22, 0) = 43: UFVect(22, 1) = "RS"  'RIO GRANDE DO SUL
   UFVect(23, 0) = 50: UFVect(23, 1) = "MS"  'MATO GROSSO DO SUL
   UFVect(24, 0) = 51: UFVect(24, 1) = "MT"  'MATO GROSS
   UFVect(25, 0) = 52: UFVect(25, 1) = "GO"  'GOIÁS
   UFVect(26, 0) = 53: UFVect(26, 1) = "DF"  'DISTRITO FEDERAL
   
   'Possíveis erros: Digitos faltando ou os dois primeiros digitos nw corresponde a UF informada
   
   'Dígitos
   If Len(CStr(Codigo)) < 7 Then
      MsgCMun = "Código do Município Informado Possuí Valores a Menos!" + vbCrLf + _
                "Eram Esperados 7 Dígitos."
      ValidaCMun = False
      Exit Function
   Else
      If Vazio(UF) Then 'Só por Precausão o GAS ja esta verificando na Validação do campo UF, mas em todo caso neh
         MsgCMun = "Cidade sem UF Cadastrada!" + vbCrLf + _
                   "Favor Informá-la."
         ValidaCMun = False
         Exit Function
      Else              'UF Conrresponde?
         For i = 0 To 26
            If UFVect(i, 1) = UF Then
               If CInt(UFVect(i, 0)) = Mid(Codigo, 1, 2) Then
                  GoTo Continua                 'Parabéns Corresponde
               Else
                  Exit For
               End If
           End If
         Next
         MsgCMun = "Código do Município não Corresponde ao Estado Informado!" + vbCrLf + _
                   "Favor Verificar a UF e o Código do Município."
         ValidaCMun = False       'Vixi deu Merd... Vamos Digitar certo
         Exit Function
      End If
   End If
             
'Vamos Verificar o Digito Verificador agora
Continua:
   DVTabela = Mid(Codigo, 7, 1) 'DV digitado
   
   'Alimentando o Vetor B(Pesos 1 e 2)
   B(0) = 1
   B(1) = 2
   B(2) = 1
   B(3) = 2
   B(4) = 1
   B(5) = 2
     
   For i = 0 To 5
      A(i) = Mid(Codigo, i + 1, 1)  'Vetor Codigo do Municipio
      C(i) = A(i) * B(i)            'Vetor Ponderação(A * B)
      If C(i) <= 9 Then             'Se a Ponderação for maior que 9 vamos somar os algarismos
         D(i) = C(i)
      Else
         D(i) = CInt(Abs((Mid(C(i), 1, 1)) + Abs(Mid(C(i), 2, 1)))) 'Vetor Soma dos Algarismos
      End If
      Soma = Soma + D(i)            'Soma Total do Algarismos
   Next
   
   RestoDiv = Soma Mod 10           'Pegando o Resto da Divição
   DV = 10 - RestoDiv               'Pegando o Dígito Verificador
   DV = IIf(DV = 10, 0, DV)         'Opa se For Resto 0 entao o DV tb é 0
   
   'Ufa acabou vamos somente verificar agora
   If DV <> DVTabela Then            'Verificando se o digito verificador confere
      MsgCMun = "Código do Município Informado está Incorreto!" + vbCrLf + _
                "Dígito Verificador Inválido."
      ValidaCMun = False
   End If
End Function

Private Sub MudaMascara()
   On Error Resume Next
   
      Select Case Uf
         Case "MG" '999.999.999/9999
            txtIE.Mask = "999.999.999/9999"
         Case "ES", "GO", "MA", "PA" '999.999.99-9
            txtIE.Mask = "999.999.99-9"
         Case "RJ" '99.999.99-9
            txtIE.Mask = "99.999.99-9" 
         Case "SC" '999.999.999
            txtIE.Mask = "999.999.999"
         Case "DF" '99.999999.999-99
            txtIE.Mask = "99.999999.999-99"
         Case "PR" '99.999.999-99
            txtIE.Mask = "99.999.999-99"
         Case "PE" '9.999.999-99
            txtIE.Mask = "9.999.999-99"
         Case "RN", "AM", "PB" '99.999.999-9
            txtIE.Mask = "99.999.999-9"
         Case "RS" '999/999999-9
            txtIE.Mask = "999/999999-9"
         Case "RO" '9999999999999-9
            txtIE.Mask = "9999999999999-9"
         Case "SP" '999.999.999.999
            'If Not e_Produtor_Paulista Then
               txtIE.Mask = "999.999.999.999"
            'Else '(Produtor) 99999999.9/999
            '   txtRGIE.Mask = "99999999.9/999"
            'End If
         Case "BA" '#99.999.999
            txtIE.Mask = "#99.999.999"
         Case "CE", "AL", "AP", "PI", "SE", "RR" '99999999-9
            txtIE.Mask = "99999999-9"
         Case "AC" '99.999.999/999-99
            txtIE.Mask = "99.999.999/999-99"
         Case "MT" '9999999999-9
            txtIE.Mask = "9999999999-9"
         Case "MS" '99.999.999-9
            txtIE.Mask = "99.999.999-9"   
         Case "TO" '99.99.999999-9
            txtIE.Mask = "99.99.999999-9"
      End Select      
      
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO   'Se nw estiver editando ou incluindo vamos pegar recuperar as informações do banco
                                                                                                   'Senão tiver essa linha quando o usuario navegar pelos registro ele nw aplica a mascara
                                                                                                   'como deveria      
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Razao_social = txtCampo(12).Value
         Nome_fantasia = txtCampo(0).Value
         Cpf = txtCampo(1).Value
         Inscricao_estadual = txtCampo(3).Value
         Cnpj = txtCampo(2).Value
         Rua = txtCampo(4).Value
         Nro = txtCampo(5).Value
         Bairro = txtCampo(6).Value
         Municipio = txtCampo(7).Value
         Codigo_do_municipio = txtCampo(10).Value
         Uf = txtCampo(8).Value
         Cep = txtCampo(9).Value
         Fone = txtCampo(11).Value
         Aliquota_do_cofins = txtCampo(14).Value
         Chave_flexdocs = txtCampo(15).Value
         Email = txtCampo(16).Value
         Data_Inicial_da_Ativacao = txtCampo(17).Value
         Aliquota_do_pis = txtCampo(13).Value
         Senha_Admin = txtCampo(18).Value
         Baixa_Pagar_Auto = chkCampo(0).Value
         Baixa_Receber_Auto = chkCampo(1).Value
         Aliquota_do_Iss = txtCampo(19).Value
         Incricao_Municipal = txtCampo(20).Value
         Emite_NFe = chkCampo(2).Value
         Percentual_da_reducao = txtCampo(21).Value
         Nao_Controlar_Estoque = chkCampo(3).Value
         Impressao_do_Relatorio = Val(labopcFrame0.Caption)
         Numeracao_Auto = chkCampo(4).Value
      Else
         Razao_social = IIf(IsNull(vgTb![Razão social]), "", vgTb![Razão social])
         Nome_fantasia = IIf(IsNull(vgTb![Nome fantasia]), "", vgTb![Nome fantasia])
         Cpf = IIf(IsNull(vgTb![Cpf do emitente]), "", vgTb![Cpf do emitente])
         Inscricao_estadual = IIf(IsNull(vgTb![Inscrição estadual]), "", vgTb![Inscrição estadual])
         Cnpj = IIf(IsNull(vgTb!Cnpj), "", vgTb!Cnpj)
         Rua = IIf(IsNull(vgTb!Rua), "", vgTb!Rua)
         Nro = IIf(IsNull(vgTb!Nro), "", vgTb!Nro)
         Bairro = IIf(IsNull(vgTb!Bairro), "", vgTb!Bairro)
         Municipio = IIf(IsNull(vgTb!Município), "", vgTb!Município)
         Codigo_do_municipio = IIf(IsNull(vgTb![Codigo do municipio]), 0, vgTb![Codigo do municipio])
         Uf = IIf(IsNull(vgTb!Uf), "", vgTb!Uf)
         Cep = IIf(IsNull(vgTb!Cep), "", vgTb!Cep)
         Fone = IIf(IsNull(vgTb!Fone), "", vgTb!Fone)
         Aliquota_do_cofins = IIf(IsNull(vgTb![Aliquota do cofins]), 0, vgTb![Aliquota do cofins])
         Chave_flexdocs = IIf(IsNull(vgTb![Chave flexdocs]), "", vgTb![Chave flexdocs])
         Email = IIf(IsNull(vgTb!Email), "", vgTb!Email)
         Data_Inicial_da_Ativacao = vgTb![Data Inicial da Ativação]
         Aliquota_do_pis = IIf(IsNull(vgTb![Aliquota do pis]), 0, vgTb![Aliquota do pis])
         Senha_Admin = IIf(IsNull(vgTb![Senha Admin]), 0, vgTb![Senha Admin])
         Baixa_Pagar_Auto = IIf(IsNull(vgTb![Baixa Pagar Auto]), 0, vgTb![Baixa Pagar Auto])
         Baixa_Receber_Auto = IIf(IsNull(vgTb![Baixa Receber Auto]), 0, vgTb![Baixa Receber Auto])
         Aliquota_do_Iss = IIf(IsNull(vgTb![Aliquota do Iss]), 0, vgTb![Aliquota do Iss])
         Incricao_Municipal = IIf(IsNull(vgTb![Incrição Municipal]), "", vgTb![Incrição Municipal])
         Emite_NFe = IIf(IsNull(vgTb![Emite NFe]), 0, vgTb![Emite NFe])
         Percentual_da_reducao = IIf(IsNull(vgTb![Percentual da redução]), 0, vgTb![Percentual da redução])
         Nao_Controlar_Estoque = IIf(IsNull(vgTb![Não Controlar Estoque]), 0, vgTb![Não Controlar Estoque])
         Impressao_do_Relatorio = IIf(IsNull(vgTb![Impressão do Relatorio]), 0, vgTb![Impressão do Relatorio])
         Numeracao_Auto = IIf(IsNull(vgTb![Numeração Auto]), 0, vgTb![Numeração Auto])
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
      i = Permitido(vgIdentTab, ACAO_INCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemInclusao = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      LigaDesligaControles Me, Not i
   End With
End Sub

'executa processos/validacoes nos campos do arquivo
Public Function Executar(vgOq As String, Optional ByRef vgColumn As Integer) As String
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer, vgInd As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = (Not Vazio(Razao_social))
      vgMsg$ = "Razão social não pode ser vazio!"
      If Not vgOk Then txtCampo(12).SetFocus
      If vgOk Then
         vgOk = (Not Vazio(Nome_fantasia))
         vgMsg$ = "Nome fantasia não pode ser vazio!"
         If Not vgOk Then txtCampo(0).SetFocus
      End If
      If vgOk Then
         vgOk = (VDV2(Cpf))
         vgMsg$ = "Cpf inválido!"
         If Not vgOk Then txtCampo(1).SetFocus
      End If
      If vgOk Then
         vgOk = (VCGC(Cnpj))
         vgMsg$ = "Cnpj inválido!"
         If Not vgOk Then txtCampo(2).SetFocus
      End If
      If vgOk Then
         If Uf <> "" Then
            vgOk = (ValidaIE(Uf, Inscricao_estadual, False))
            vgMsg$ = "" & MsgValIE & ""
            If Not vgOk Then txtCampo(3).SetFocus
         End If
      End If
      If vgOk Then
         vgOk = (ValidaCMun(Codigo_do_municipio, UF))
         vgMsg$ = "" & MsgCMun & ""
         If Not vgOk Then txtCampo(10).SetFocus
      End If
      If vgOk Then
         vgOk = ((Not Vazio(Data_Inicial_da_Ativacao)) And (IsDate(Data_Inicial_da_Ativacao) Or Vazio(Data_Inicial_da_Ativacao)))
         vgMsg$ = "Data Inicial da Ativação não pode ser vazio!"
         If Not vgOk Then txtCampo(17).SetFocus
      End If
      If vgOk Then
         vgOk = (Senha_Admin > 0)
         vgMsg$ = "Senha Admin inválido!"
         If Not vgOk Then txtCampo(18).SetFocus
      End If
      If vgOk Then vgMsg$ = ""
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
         mmCampo(0).Clear
         chkCampo(0).Value = False
         chkCampo(1).Value = False
         chkCampo(2).Value = False
         chkCampo(3).Value = False
         opcFrame0(0).Value = True
         chkCampo(4).Value = False
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         InicializaApelidos COM_REGISTRO
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         Set mmCampo(0).DataSource = vgTb
         mmCampo(0).ShowMedia
         chkCampo(0).Value = Baixa_Pagar_Auto
         chkCampo(1).Value = Baixa_Receber_Auto
         chkCampo(2).Value = Emite_NFe
         chkCampo(3).Value = Nao_Controlar_Estoque
         opcFrame0(Impressao_do_Relatorio).Value = True
         chkCampo(4).Value = Numeracao_Auto
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
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            If vgTb.Fields(txtCampo(i).DataField).Value <> txtCampo(i).OriginalValue Then
               If Len(vgMsg$) = 0 Then
                  vgMsg$ = Caption + "|3600|" + LoadGasString(122)
               End If
               vgPriVez = True
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
               vgPriVez = False
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
      InicializaApelidos COM_TEXTBOX
      vgTb![Baixa Pagar Auto] = Baixa_Pagar_Auto
      vgTb![Baixa Receber Auto] = Baixa_Receber_Auto
      vgTb![Emite NFe] = Emite_NFe
      vgTb![Não Controlar Estoque] = Nao_Controlar_Estoque
      vgTb![Impressão do Relatorio] = Impressao_do_Relatorio
      mmCampo(0).SaveMedia
      vgTb![Numeração Auto] = Numeracao_Auto
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

'evento - imagem alterada
Private Sub mmCampo_Change(Index As Integer)
   If vgPriVez Then Exit Sub
   LigaFocos Me
End Sub

'evento - imagem carregada
Private Sub mmCampo_OpenMedia(Index As Integer, Status As Integer)
   mdiNFE.Enabled = (Status <> 1)
End Sub

'evento - perdeu o foco
Private Sub mmCampo_LostFocus(Index As Integer)
   imgCampo(Index).LostFocus
End Sub

'evento - recebeu o foco
Private Sub mmCampo_GotFocus(Index As Integer)
   imgCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then                 'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   vgTemAlteracaoGrids = Not LockGrids
DeuErro:
   
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   Select Case Index
      Case 8
         MudaMascara
   End Select
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 17
            If Len(txtCp(17).Text) = 0 Then
               txtCampo(17).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
            End If
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
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Baixa_Pagar_Auto
      ElseIf Index = 1 Then
         chkCampo(1).Value = Baixa_Receber_Auto

      ElseIf Index = 2 Then
         chkCampo(2).Value = Emite_NFe

      ElseIf Index = 3 Then
         chkCampo(3).Value = Nao_Controlar_Estoque

      ElseIf Index = 4 Then
         chkCampo(4).Value = Numeracao_Auto
      End If
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      chkCampo(Index).Change
   End If
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

'evento - quando uma opção for selecionada
Private Sub opcFrame0Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcFrame0(Index).Locked Then
      opcFrame0(Val(labopcFrame0.Caption)).Value = True
   Else
      If Val(labopcFrame0.Caption) <> opcFrame0(Index).BookMark Then
         labopcFrame0.Caption = Str$(opcFrame0(Index).BookMark)
         LigaFocos Me
         InicializaApelidos COM_TEXTBOX
         opcFrame0(Index).Change
      End If
   End If
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcFrame0Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcFrame0(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcFrame0Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcFrame0(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcFrame0Cp_GotFocus(Index As Integer)
   opcFrame0(Index).GotFocus
End Sub

'evento - quando o campo perder o foco
Private Sub opcFrame0Cp_LostFocus(Index As Integer)
   opcFrame0(Index).LostFocus
End Sub

'evento - quando o formulário receber o foco
Private Sub Form_Activate()
   AtivaForm Me
   
   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   If FormEstaAberto("frmEnviaEmail") Then
      If Not frmEnviaEmail.Visible Then
         UnLoad vgFrmImpCons
         Set vgFrmImpCons = Nothing
         UnLoad frmEnviaEmail
      End If
   End If
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(12000)
   vgFormID = 614
   vgIdentTab = "Empresa"
   vgFiltroEmUso = -1
   vgIndDefault = ""
   vgPriVez = True
   vgTipo = TP_PARAMETROS
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = False
   vgTemFiltro = False
   vgTemBrowse = False
   vgTemCondicoesEsp = False
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 31
   vgSituacao = ACAO_NAVEGANDO
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(3505))
   Set txtIE = txtCampo(3)
   Set lblAjuste = Label(16)
   Set vgTb = New GRecordSet
   DefineControles
   vgTooltips.Create
   vgTooltips.AddTool mmCampo(0), 0, LoadGasString(12005)
   Label(0).Caption = LoadGasString(12006)
   Label(1).Caption = LoadGasString(12007)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(12008)
   Label(2).Caption = LoadGasString(12009)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(12010)
   Label(3).Caption = LoadGasString(12011)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(12012)
   Label(4).Caption = LoadGasString(12013)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(12014)
   Label(5).Caption = LoadGasString(12015)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(12016)
   Label(6).Caption = LoadGasString(12017)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(12018)
   Label(7).Caption = LoadGasString(12019)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(12020)
   Label(8).Caption = LoadGasString(12021)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(12022)
   Label(9).Caption = LoadGasString(12023)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(12024)
   Label(10).Caption = LoadGasString(12025)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(12026)
   Label(11).Caption = LoadGasString(12027)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(12028)
   Label(12).Caption = LoadGasString(12029)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(12030)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(12031)
   Label(13).Caption = LoadGasString(12032)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(12033)
   Label(14).Caption = LoadGasString(12034)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(12035)
   Label(15).Caption = LoadGasString(12036)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(12037)
   Label(16).Caption = LoadGasString(12038)
   Label(17).Caption = LoadGasString(12039)
   vgTooltips.AddTool txtCampo(16).CtPri, 0, LoadGasString(12040)
   Label(18).Caption = LoadGasString(12041)
   vgTooltips.AddTool txtCampo(17).CtPri, 0, LoadGasString(12042)
   Label(19).Caption = LoadGasString(12043)
   vgTooltips.AddTool txtCampo(18).CtPri, 0, LoadGasString(12044)
   Label(20).Caption = LoadGasString(12045)
   vgTooltips.AddTool txtCampo(19).CtPri, 0, LoadGasString(12046)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(12047)
   chkCampo(0).Caption = LoadGasString(12048)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(12049)
   chkCampo(1).Caption = LoadGasString(12050)
   Label(21).Caption = LoadGasString(12051)
   vgTooltips.AddTool txtCampo(20).CtPri, 0, LoadGasString(12052)
   vgTooltips.AddTool chkCampo(2).CtPri, 0, LoadGasString(12053)
   chkCampo(2).Caption = LoadGasString(12054)
   Label(22).Caption = LoadGasString(12055)
   vgTooltips.AddTool txtCampo(21).CtPri, 0, LoadGasString(12056)
   vgTooltips.AddTool chkCampo(3).CtPri, 0, LoadGasString(12057)
   chkCampo(3).Caption = LoadGasString(12058)
   vgTooltips.AddTool Frame(0), 0, LoadGasString(12059)
   Frame(0).Caption = LoadGasString(12060)
   vgTooltips.AddTool opcFrame0(0).CtPri, 0, LoadGasString(12061)
   opcFrame0(0).Caption = LoadGasString(12062)
   vgTooltips.AddTool opcFrame0(1).CtPri, 0, LoadGasString(12063)
   opcFrame0(1).Caption = LoadGasString(12064)
   vgTooltips.AddTool Botao(1), 0, LoadGasString(12065)
   Botao(1).Caption = LoadGasString(12066)
   vgTooltips.AddTool chkCampo(4).CtPri, 0, LoadGasString(12067)
   chkCampo(4).Caption = LoadGasString(12068)
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

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 0
   txtCampo(12).Mask = "@x"
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 0
   txtCampo(0).Mask = "@x"
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = "999.999.999-99"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 0
   txtCampo(2).Mask = "99.999.999/9999-99"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@!"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "@x"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 0
   txtCampo(5).Mask = "@x"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = "@x"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = "@x"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_COMBO) = bottxtCampo8(BOT_COMBO)
   Set bottxtCampo8(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(8).DataType = 0
   txtCampo(8).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE"

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = "99999-999"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = "9999999"
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 0
   txtCampo(11).Mask = "(99)9999-9999"
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""

   Set txtCampo(13).CtPri = txtCp(13)
   Set txtCampo(13).CtFdo = labFdo13
   Set txtCampo(13).CtBot(BOT_ACAO) = bottxtCampo13(BOT_ACAO)
   Set bottxtCampo13(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(13).DataType = 1
   txtCampo(13).Mask = "999,9999"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""

   Set txtCampo(14).CtPri = txtCp(14)
   Set txtCampo(14).CtFdo = labFdo14
   Set txtCampo(14).CtBot(BOT_ACAO) = bottxtCampo14(BOT_ACAO)
   Set bottxtCampo14(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(14).DataType = 1
   txtCampo(14).Mask = "999,9999"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 0
   txtCampo(15).Mask = "@a"
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""

   Set txtCampo(16).CtPri = txtCp(16)
   txtCampo(16).DataType = 0
   txtCampo(16).Mask = "@a"
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""

   Set txtCampo(17).CtPri = txtCp(17)
   Set txtCampo(17).CtFdo = labFdo17
   Set txtCampo(17).CtBot(BOT_LISTA) = bottxtCampo17(BOT_LISTA)
   Set bottxtCampo17(BOT_LISTA).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(17).DataType = 2
   txtCampo(17).Mask = "99/99/9999"
   txtCampo(17).PesqModoAbertura =  2 
   txtCampo(17).BoundColumn = ""
   txtCampo(17).ListFields = ""
   txtCampo(17).OrderFields = ""
   txtCampo(17).Relation = ""
   txtCampo(17).Source = ""

   Set txtCampo(18).CtPri = txtCp(18)
   txtCampo(18).DataType = 1
   txtCampo(18).Mask = "99999999"
   txtCampo(18).BoundColumn = ""
   txtCampo(18).ListFields = ""
   txtCampo(18).OrderFields = ""
   txtCampo(18).Relation = ""
   txtCampo(18).Source = ""

   Set txtCampo(19).CtPri = txtCp(19)
   txtCampo(19).DataType = 1
   txtCampo(19).Mask = "99,99"
   txtCampo(19).BoundColumn = ""
   txtCampo(19).ListFields = ""
   txtCampo(19).OrderFields = ""
   txtCampo(19).Relation = ""
   txtCampo(19).Source = ""

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5

   Set txtCampo(20).CtPri = txtCp(20)
   txtCampo(20).DataType = 0
   txtCampo(20).Mask = "@!"
   txtCampo(20).BoundColumn = ""
   txtCampo(20).ListFields = ""
   txtCampo(20).OrderFields = ""
   txtCampo(20).Relation = ""
   txtCampo(20).Source = ""

   Set chkCampo(2).CtPri = chkCp(2)
   chkCampo(2).DataType = 5

   Set txtCampo(21).CtPri = txtCp(21)
   txtCampo(21).DataType = 1
   txtCampo(21).Mask = "99,99"
   txtCampo(21).BoundColumn = ""
   txtCampo(21).ListFields = ""
   txtCampo(21).OrderFields = ""
   txtCampo(21).Relation = ""
   txtCampo(21).Source = ""

   Set chkCampo(3).CtPri = chkCp(3)
   chkCampo(3).DataType = 5

   Set opcFrame0(0).CtPri = opcFrame0Cp(0)
   Set opcFrame0(0).CtFdo = labopcFrame0
   opcFrame0(0).DataType = 6
   opcFrame0(0).BookMark = 0

   Set opcFrame0(1).CtPri = opcFrame0Cp(1)
   Set opcFrame0(1).CtFdo = labopcFrame0
   opcFrame0(1).DataType = 6
   opcFrame0(1).BookMark = 1

   Set imgCampo(0).CtPri = mmCampo(0)
   imgCampo(0).DataType = 99

   Set chkCampo(4).CtPri = chkCp(4)
   chkCampo(4).DataType = 5

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
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
   Dim i As Integer
   AjustaPanFundo Me
   mmCampo(0).Refresh
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set txtIE = Nothing
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set mmCampo(0) = Nothing
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   Set chkCampo(3) = Nothing
   Set chkCampo(4) = Nothing

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   Set frmEmpresa = Nothing                       'libera o segmento de código do form
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         ConsultaCodMun Municipio 
      Case 1
         mmCampo(0).PopMenu
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo8_Click(Index As Integer)
   txtCampo(8).SetFocus
   DoEvents
   txtCampo(8).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo13_Click(Index As Integer)
   txtCampo(13).SetFocus
   DoEvents
   txtCampo(13).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo14_Click(Index As Integer)
   txtCampo(14).SetFocus
   DoEvents
   txtCampo(14).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo17_Click(Index As Integer)
   txtCampo(17).SetFocus
   DoEvents
   txtCampo(17).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


