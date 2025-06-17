Version 5.00
Begin VB.Form frmPivos
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8790
   Left                     =   1170
   LinkTopic                =   "Pivos"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8355
   ScaleWidth               =   17865
   Top                      =   1150
   Width                    =   17955
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
      Height                   =   8175
      Left                     =   15
      Width                    =   17685
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   31178
      _extenty                 =   14412
      Index                    =   0 
      Begin IRRIG.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   -1
         Height                   =   1545
         Left                     =   120
         Width                    =   1530
         Top                      =   120
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
         Height                   =   360
         Left                     =   1800
         Width                    =   3300
         Top                      =   120
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   14.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Número de Série (Pivo)"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1800
         Y1                       =   510
         X2                       =   5070
         Y2                       =   510
         Index                    =   0 
      End
      Begin IRRIG.GPainel Painel
         BackColor                =   &H8000000F
         BevelOuter               =   1
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   1140
         Left                     =   180
         Width                    =   10350
         Top                      =   1800
         TabStop                  =  0
         _extentx                 =   18247
         _extenty                 =   2009
         Index                    =   1 
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   120
            Width                    =   570
            Top                      =   60
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
            Caption                  =   "Seq..:"
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
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   0
            TabStop                  =   0
            Height                   =   300
            Left                     =   720
            Width                    =   840
            Top                      =   60
            DataField                =   "Seq do Pivo"
            TabStop                  =  0
            Index                    =   0 
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
            Width                    =   810
            Top                      =   60
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
            Caption                  =   "Modelo:"
            Index                    =   2 
         End
         Begin VB.Label labFdo1
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   2520
            Top                      =   60
            Width                    =   1170
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo1
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   3420
            Top                      =   75
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
            Left                     =   2535
            Width                    =   870
            Top                      =   75
            DataField                =   "Modelo do Pivo"
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   3840
            Width                    =   1050
            Top                      =   60
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
            Caption                  =   "Descrição:"
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
            MaxLength                =   30
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   3
            TabStop                  =   -1
            Height                   =   300
            Left                     =   4920
            Width                    =   2010
            Top                      =   60
            DataField                =   "Descri do Pivo"
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   7080
            Width                    =   570
            Top                      =   60
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
            Caption                  =   "Série:"
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
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   4
            TabStop                  =   0
            Height                   =   300
            Left                     =   7740
            Width                    =   600
            Top                      =   60
            DataField                =   "Serie do Pivo"
            TabStop                  =  0
            Index                    =   3 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   8460
            Width                    =   1050
            Top                      =   60
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
            Caption                  =   "Mês e Ano:"
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
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   5
            TabStop                  =   -1
            Height                   =   300
            Left                     =   9600
            Width                    =   600
            Top                      =   60
            DataField                =   "MesAno"
            Index                    =   4 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   120
            Width                    =   570
            Top                      =   480
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
            Caption                  =   "Tipo.:"
            Index                    =   6 
         End
         Begin VB.Label labFdo5
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   720
            Top                      =   480
            Width                    =   870
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo5
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   1320
            Top                      =   495
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
            TabIndex                 =   6
            TabStop                  =   -1
            Height                   =   270
            Left                     =   735
            Width                    =   570
            Top                      =   495
            DataField                =   "Letra do Pivo"
            Index                    =   5 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   1680
            Width                    =   750
            Top                      =   480
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
            Caption                  =   "Data......:"
            Index                    =   7 
         End
         Begin VB.Label labFdo6
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   2520
            Top                      =   480
            Width                    =   1410
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo6
            Style                    =   1
            TabStop                  =   0
            Index                    =   0
            Left                     =   3660
            Top                      =   495
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
            Left                     =   2535
            Width                    =   1110
            Top                      =   495
            DataField                =   "Data de Criação"
            Index                    =   6 
         End
         Begin VB.Frame Frame
            Height                   =   705
            Left                     =   4290
            Width                    =   5910
            Top                      =   375
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
               AutoSize                 =   0
               Height                   =   240
               Left                     =   120
               Width                    =   1290
               Top                      =   240
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
               Caption                  =   "Nº de Série:"
               Index                    =   8 
            End
            Begin VB.TextBox txtCp
               Appearance               =   0
               BackColor                =   &HC0FFC0
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
               Locked                   =   -1
               TabIndex                 =   8
               TabStop                  =   0
               Height                   =   300
               Left                     =   1500
               Width                    =   2610
               Top                      =   240
               DataField                =   "Nro de Serie do Pivo"
               TabStop                  =  0
               Index                    =   7 
            End
            Begin IRRIG.GBotao Botao
               ButtonStyle              =   0
               JoinSeparators           =   0
               LeftSeparator            =   0
               TopSeparator             =   0
               TabIndex                 =   9
               TabStop                  =   -1
               ButtonType               =   0
               Height                   =   420
               Left                     =   4320
               Width                    =   1470
               Top                      =   180
               CaptionAlign             =   0
               CaptionOffset            =   50
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name                     =   "Courier New"
                  CharSet                  =   0
                  Weight                   =   700
                  Size                     =   11.25
                  Underline                =   0
                  Italic                   =   0
                  StrikeThrough            =   0
               EndProperty
               ForeColor                =   &H80000012
               Caption                  =   "Gerar"
               Index                    =   0 
            End
         End
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
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   0
         Height                   =   255
         Left                     =   1800
         Width                    =   2715
         Top                      =   1020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Equipamento Entregue"
         Index                    =   0 
         DataField                =   "Entregue"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   645
         Top                      =   3060
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
         Caption                  =   "Geral.:"
         Index                    =   9 
      End
      Begin VB.Label labFdo8
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HC0FFC0
         Left                     =   840
         Top                      =   3000
         Width                    =   6390
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6960
         Top                      =   3015
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6705
         Top                      =   3015
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HC0FFC0
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
         Left                     =   855
         Width                    =   5835
         Top                      =   3015
         DataField                =   "Codigo do Geral"
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7380
         Width                    =   1605
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
         Caption                  =   "Data da Entrega:"
         Index                    =   10 
      End
      Begin VB.Label labFdo9
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   9000
         Top                      =   3000
         Width                    =   1380
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo9
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10110
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
         TabIndex                 =   11
         TabStop                  =   -1
         Height                   =   270
         Left                     =   9015
         Width                    =   1080
         Top                      =   3015
         DataField                =   "Data da Entrega"
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   600
         Top                      =   3420
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
         Caption                  =   "NFe...:"
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
         ForeColor                =   &H80000008
         MaxLength                =   60
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   12
         TabStop                  =   -1
         Height                   =   300
         Left                     =   840
         Width                    =   4290
         Top                      =   3420
         DataField                =   "NF"
         Index                    =   10 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   615
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
         Caption                  =   "Obs...:"
         Index                    =   12 
      End
      Begin VB.TextBox txtCp
         Appearance               =   1
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
         TabIndex                 =   13
         TabStop                  =   -1
         Height                   =   990
         Left                     =   840
         Width                    =   9675
         Top                      =   3810
         DataField                =   "Obs"
         Index                    =   11 
         MultiLine                =   -1
         ScrollBars               =   2
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   5280
         Width                    =   1050
         Top                      =   3420
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
         Caption                  =   "Vendedor:"
         Index                    =   13 
      End
      Begin VB.Label labFdo12
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HFFFFFF
         Left                     =   6360
         Top                      =   3420
         Width                    =   4050
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo12
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   10140
         Top                      =   3435
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HFFFFFF
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
         Left                     =   6375
         Width                    =   3750
         Top                      =   3435
         DataField                =   "Codigo do Vendedor"
         Index                    =   12 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   5835
         Width                    =   1455
         Top                      =   240
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
         Caption                  =   "Responsavel:"
         Index                    =   14 
      End
      Begin VB.Label labFdo13
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7365
         Top                      =   240
         Width                    =   3060
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo13
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   10155
         Top                      =   255
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
         TabIndex                 =   15
         TabStop                  =   0
         Height                   =   270
         Left                     =   7380
         Width                    =   2760
         Top                      =   255
         DataField                =   "Entrega Tecnica"
         Index                    =   13 
      End
      Begin VB.CheckBox ChkCp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H8000
         Alignment                =   0
         TabIndex                 =   16
         TabStop                  =   0
         Height                   =   255
         Left                     =   1800
         Width                    =   2505
         Top                      =   1260
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H8000
         Alignment                =   0
         Caption                  =   "Fez Entrega Tecnica"
         Index                    =   1 
         DataField                =   "Entrega Tec"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   1800
         Width                    =   1125
         Top                      =   600
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Pedido:"
         Index                    =   15 
      End
      Begin VB.Label labFdo14
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3000
         Top                      =   600
         Width                    =   2025
         Height                   =   330
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   4755
         Top                      =   615
         Width                    =   255
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4500
         Top                      =   615
         Width                    =   255
         Height                   =   300
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   1
         TabIndex                 =   17
         TabStop                  =   0
         Height                   =   300
         Left                     =   3015
         Width                    =   1470
         Top                      =   615
         DataField                =   "Seqüência do Orçamento"
         Index                    =   14 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   17640
         Width                    =   45
         Top                      =   8340
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
         Index                    =   16 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6105
         Width                    =   1185
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
         Caption                  =   "Colaborador:"
         Index                    =   17 
      End
      Begin VB.Label labFdo15
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7365
         Top                      =   600
         Width                    =   3060
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo15
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   10155
         Top                      =   615
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
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   270
         Left                     =   7380
         Width                    =   2760
         Top                      =   615
         DataField                =   "Codigo do solicitante"
         Index                    =   15 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5475
         Width                    =   1815
         Top                      =   1335
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
         Caption                  =   "Dt. Entrega Tecnica:"
         Index                    =   18 
      End
      Begin VB.Label labFdo16
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7365
         Top                      =   1320
         Width                    =   1455
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo16
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   8550
         Top                      =   1335
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
         TabIndex                 =   19
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7380
         Width                    =   1155
         Top                      =   1335
         DataField                =   "Data Tecnica"
         Index                    =   16 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   330
         Left                     =   135
         Width                    =   10380
         Top                      =   5895
         Alignment                =   2
         BeginProperty Font
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Entrega Tecnica / Assistencia  F12 - Carrega Foto"
         Index                    =   19 
      End
      Begin IRRIG.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
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
         TabIndex                 =   20
         TabStop                  =   -1
         Stripes                  =   -1
         SaveGridBars             =   0
         Height                   =   1935
         Left                     =   180
         Width                    =   10335
         Top                      =   6135
         Caption                  =   ""
         ShowGridCaption          =   0
         Index                    =   0 
      End
      Begin IRRIG.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   1
         Stretch                  =   -1
         Height                   =   7785
         Left                     =   10800
         Width                    =   6810
         Top                      =   255
         Index                    =   1 
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
         Height                   =   240
         Left                     =   5220
         Width                    =   2070
         Top                      =   975
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
         Caption                  =   "Prev. de Montagem:"
         Index                    =   20 
      End
      Begin VB.Label labFdo17
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7380
         Top                      =   960
         Width                    =   1440
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo17
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   8550
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
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   21
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7395
         Width                    =   1140
         Top                      =   975
         DataField                =   "Prev Montagem"
         Index                    =   17 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   540
         Top                      =   4815
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
         Caption                  =   "Obs 2"
         Index                    =   21 
      End
      Begin VB.TextBox txtCp
         Appearance               =   1
         BackColor                =   &HC0FFFF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   22
         TabStop                  =   -1
         Height                   =   1080
         Left                     =   840
         Width                    =   9675
         Top                      =   4800
         DataField                =   "DadosAd"
         Index                    =   18 
         MultiLine                =   -1
         ScrollBars               =   2
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   23
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   8940
         Width                    =   1470
         Top                      =   1080
         CaptionAlign             =   0
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
         Caption                  =   "Imprimir"
         Index                    =   1 
      End
   End
         Begin IRRIG.GListV grdBrowse 
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
Attribute VB_name = "frmPivos"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: Pivos
'* Função....: Pivo
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 15:01:55
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
Public vgRepeticao As Integer                     'flag de repetição do último reg digitado
Public vgAlterar As Integer                       'flag de Alteracao de registros
Public vgUltAlterar As Integer                    'flag de última situação de "pode alterar"
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso 
Public vgIndDefault As String                     'indice default do módulo 
Public vgFormID As Long                           'identificador único para o módulo 
Public vgIdentTab As String                       'nome da tabela principal do módulo 
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFiltroOriginal As String
Dim vgPegouSeq As Boolean                         'flag se a sequencia do campo foi incrementada 
Dim txtCampo(18) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(1) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim Seq_do_Pivo As Long, Modelo_do_Pivo As String, Descri_do_Pivo As String
Dim Serie_do_Pivo As Integer, MesAno As String, Letra_do_Pivo As String
Dim Data_de_Criacao As Variant, Nro_de_Serie_do_Pivo As String, Codigo_do_Geral As Long
Dim Entregue As Boolean, Obs As String, NF As String
Dim Data_da_Entrega As Variant, Codigo_do_Vendedor As Long, Entrega_Tecnica As String
Dim Entrega_Tec As Boolean, Sequencia_do_Orcamento As Long, Codigo_do_solicitante As Integer
Dim Data_Tecnica As Variant, Prev_Montagem As Variant, DadosAd As String
Public NroSerie As Object, txtCodigo_do_Geral As Object, txtCodigo_do_Vendedor As Object
Public lblAjuste As Object
Dim Serie_Pivos As New GRecordSet, Manutencao_Pivo As New GRecordSet

Private Function GeraSerie() As Long
Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("Select Max([Serie do Pivo])Seq From [Serie Pivos]")
 
  If Tb.RecordCount > 0 Then
     GeraSerie = Tb!Seq + 1
  Else
     GeraSerie = 1
  End If
        
End Function


Private Sub GeraNroSerie()
Dim NroDeSerie As String, Mes_Ano As String, Seri As String

 Seri = Format(Serie_do_Pivo,"0000")
 Mes_Ano = RemoveCaracteres(MesAno)
 NroDeSerie = Modelo_do_Pivo & "/" & Seri & "/" & Mes_Ano & Letra_do_Pivo    

 NroSerie.Text = NroDeSerie

End Sub



Private Function SeqItem() As Integer
 Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("Select Max([Sequencia do Item])Seq From [Manutenção Pivo] Where [Seq do Pivo] = " & Seq_do_Pivo )
 
  If Tb.RecordCount > 0 Then
     SeqItem = Tb!Seq + 1
  Else
     SeqItem = 1
  End If
        
End Function


Private Sub BuscaFoto(KeyAscii As Integer, Seq_do_Pivo As Long, Sequencia_do_Item As Long, Data_da_Manutencao As Variant, _
   Descricao_da_Manutencao As String) 
   
 On Error GoTo DeuErro
 
 If KeyAscii = vbKeyF12 Then
   If Existe("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & Sequencia_do_Item & ".jpg") Then
      Set mmCampo(1).Picture = LoadPicture("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & Sequencia_do_Item & ".jpg")
      mmCampo(1).ToolTipText = "Foto"
   ElseIf Existe("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & Sequencia_do_Item & ".JPEG") Then
          Set mmCampo(1).Picture = LoadPicture("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & Sequencia_do_Item & ".JPEG")
          mmCampo(1).ToolTipText = "Foto"
   Else
      Set mmCampo(1).Picture = Nothing
 End If
   
   DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
   
 End If
  
End Sub

Private Sub BuscaFotoAux()
 
 If Existe("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & 1 & ".jpg") Then
      Set mmCampo(1).Picture = LoadPicture("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & 1 & ".jpg")
      mmCampo(1).ToolTipText = "Foto"
 ElseIf Existe("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & 1 & ".JPEG") Then
      Set mmCampo(1).Picture = LoadPicture("\\Srvdc\Public$\Entrega _Tecnica _Assistencia_Pivo\" & Seq_do_Pivo & "_" & 1 & ".JPEG")
      mmCampo(1).ToolTipText = "Foto"
 End If
 
End Sub

Private Sub BuscaGeral() 
 Dim Tb As New GRecordSet
 
  If Sequencia_do_Orcamento = 0 Then Exit Sub
  
  Set Tb = VgDb.OpenRecordSet("SELECT [Seqüência Do Geral] Geral, [Seqüência Do Vendedor] Vendedor From Orçamento Where [Seqüência do Orçamento] = " & Sequencia_do_Orcamento)
      
      If Tb!Geral > 0 Then
         txtCodigo_do_Geral.Value = Tb!Geral  
      End If 
      If Tb!Geral > 0 Then
         txtCodigo_do_Vendedor.Value = Tb!Vendedor  
      End If 
 
End Sub

Public Sub LigaDesligaBotoes
   Botao( 1 ).Enabled =    Botao( 1 ).Enabled And Permitido("Serie Pivos", ACAO_NAVEGANDO)
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Seq_do_Pivo = txtCampo(0).Value
         Modelo_do_Pivo = txtCampo(1).Value
         Descri_do_Pivo = txtCampo(2).Value
         Serie_do_Pivo = txtCampo(3).Value
         MesAno = txtCampo(4).Value
         Letra_do_Pivo = txtCampo(5).Value
         Data_de_Criacao = txtCampo(6).Value
         Nro_de_Serie_do_Pivo = txtCampo(7).Value
         Codigo_do_Geral = txtCampo(8).Value
         Entregue = chkCampo(0).Value
         Obs = txtCampo(11).Value
         NF = txtCampo(10).Value
         Data_da_Entrega = txtCampo(9).Value
         Codigo_do_Vendedor = txtCampo(12).Value
         Entrega_Tecnica = txtCampo(13).Value
         Entrega_Tec = chkCampo(1).Value
         Sequencia_do_Orcamento = txtCampo(14).Value
         Codigo_do_solicitante = txtCampo(15).Value
         Data_Tecnica = txtCampo(16).Value
         Prev_Montagem = txtCampo(17).Value
         DadosAd = txtCampo(18).Value
      Else
         Seq_do_Pivo = IIf(IsNull(vgTb![Seq do Pivo]), 0, vgTb![Seq do Pivo])
         Modelo_do_Pivo = IIf(IsNull(vgTb![Modelo do Pivo]), "", vgTb![Modelo do Pivo])
         Descri_do_Pivo = IIf(IsNull(vgTb![Descri do Pivo]), "", vgTb![Descri do Pivo])
         Serie_do_Pivo = IIf(IsNull(vgTb![Serie do Pivo]), 0, vgTb![Serie do Pivo])
         MesAno = IIf(IsNull(vgTb!MesAno), "", vgTb!MesAno)
         Letra_do_Pivo = IIf(IsNull(vgTb![Letra do Pivo]), "", vgTb![Letra do Pivo])
         Data_de_Criacao = vgTb![Data de Criação]
         Nro_de_Serie_do_Pivo = IIf(IsNull(vgTb![Nro de Serie do Pivo]), "", vgTb![Nro de Serie do Pivo])
         Codigo_do_Geral = IIf(IsNull(vgTb![Codigo do Geral]), 0, vgTb![Codigo do Geral])
         Entregue = IIf(IsNull(vgTb!Entregue), 0, vgTb!Entregue)
         Obs = IIf(IsNull(vgTb!Obs), "", vgTb!Obs)
         NF = IIf(IsNull(vgTb!NF), "", vgTb!NF)
         Data_da_Entrega = vgTb![Data da Entrega]
         Codigo_do_Vendedor = IIf(IsNull(vgTb![Codigo do Vendedor]), 0, vgTb![Codigo do Vendedor])
         Entrega_Tecnica = IIf(IsNull(vgTb![Entrega Tecnica]), "", vgTb![Entrega Tecnica])
         Entrega_Tec = IIf(IsNull(vgTb![Entrega Tec]), 0, vgTb![Entrega Tec])
         Sequencia_do_Orcamento = IIf(IsNull(vgTb![Seqüência do Orçamento]), 0, vgTb![Seqüência do Orçamento])
         Codigo_do_solicitante = IIf(IsNull(vgTb![Codigo do solicitante]), 0, vgTb![Codigo do solicitante])
         Data_Tecnica = vgTb![Data Tecnica]
         Prev_Montagem = vgTb![Prev Montagem]
         DadosAd = IIf(IsNull(vgTb!DadosAd), "", vgTb!DadosAd)
      End If
   End If
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'verifica permissões para as condições especiais
'e muda situação de alguns botões
Private Sub AnalisaCondicoes()
   Dim i As Integer
   On Error Resume Next
   If Not mdiIRRIG.ActiveForm Is Nothing Then
      If mdiIRRIG.ActiveForm.Name <> Me.Name And Me.Visible Then Exit Sub
   End If
   With mdiIRRIG
      i = Permitido(vgIdentTab, ACAO_INCLUINDO)
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
      i = ((Entrega_Tec = 0) Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "EDUARDO THEODORO" Or vgPWUsuario = "YGOR") and Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i And vgAlterar
      .Menu_Paltera.Enabled = i
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
      vgOk = (Not Vazio(Modelo_do_Pivo))
      vgMsg$ = "Modelo do Pivo não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = (Not Vazio(MesAno))
         vgMsg$ = "MesAno inválido!"
         If Not vgOk Then vgColumn = 5
      End If
      If vgOk Then
         vgOk = (Not Vazio(Letra_do_Pivo))
         vgMsg$ = "Informar o Tipo do Pivo"
         If Not vgOk Then vgColumn = 6
      End If
      If vgOk Then
         vgOk = ((Not Vazio(Data_de_Criacao)) And (IsDate(Data_de_Criacao) Or Vazio(Data_de_Criacao)))
         vgMsg$ = "Data de Criação não pode ser vazio!"
         If Not vgOk Then vgColumn = 7
      End If
      If vgOk Then
         vgOk = (IsDate(Data_da_Entrega) Or Vazio(Data_da_Entrega))
         vgMsg$ = "Data da Entrega não pode ser vazio!"
         If Not vgOk Then vgColumn = 10
      End If
      If vgOk Then
         vgOk = (Sequencia_do_Orcamento >= 0)
         vgMsg$ = "Seqüência do Orçamento inválido!"
         If Not vgOk Then vgColumn = 15
      End If
      If vgOk Then
         vgOk = (Codigo_do_solicitante >= 0)
         vgMsg$ = "Informe o Colaborador!"
         If Not vgOk Then vgColumn = 16
      End If
      If vgOk Then
         vgOk = (IsDate(Data_Tecnica) Or Vazio(Data_Tecnica))
         vgMsg$ = "Data Tecnica não pode ser vazio!"
         If Not vgOk Then vgColumn = 17
      End If
      If vgOk Then
         If vgPWUsuario = "JERONIMO" Or vgPWUsuario = "YGOR" Then
            vgOk = (IsDate(Prev_Montagem) Or Vazio(Prev_Montagem))
            vgMsg$ = "Prev Montagem não pode ser vazio!"
            If Not vgOk Then vgColumn = 18
         End If
      End If
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
                  txtCampo(3).Value = GeraSerie()
         chkCampo(0).Value = False
         chkCampo(1).Value = False
                  On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 8
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 14
         PoeRelEFiltroCbo 15
      End If
      If vgSituacao = ACAO_INCLUINDO Then
         txtCampo(0).Value = PegaSequencia(vgIdentTab,"Seq do Pivo", 0, 1)
         InicializaApelidos COM_TEXTBOX
         vgPegouSeq = True
      End If
   ElseIf vgOq = VOLTA_SEQUENCIAL Then
      If vgPegouSeq  And vgSituacao = ACAO_INCLUINDO Then
         VoltaSequencia vgIdentTab,"Seq do Pivo"
      End If
      vgPegouSeq = False
   ElseIf vgOq = CONFIRMA_SEQUENCIAL Then
      If vgPegouSeq  And vgSituacao = ACAO_INCLUINDO Then
         ConfirmaSequencia vgIdentTab,"Seq do Pivo"
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 8
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 14
         PoeRelEFiltroCbo 15
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Entregue
         chkCampo(1).Value = Entrega_Tec
         If vgSituacao = ACAO_NAVEGANDO Then
            If Me.Name = mdiIRRIG.ActiveForm.Name Then
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
      InicializaApelidos COM_TEXTBOX
      vgTb!Entregue = Entregue
      vgTb![Entrega Tec] = Entrega_Tec
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaPreValidacao
   ElseIf vgOq = PODE_ALTERAR Then
      vgOk = (vgSituacao = ACAO_INCLUINDO Or vgAlterar)
      For i = 0 To UBound(txtCampo)
         If Len(txtCampo(i).DataField) > 0 Then
            txtCampo(i).Locked = Not (vgOk And txtCampo(i).Editable)
         End If
      Next
      For i = 0 To UBound(chkCampo)
         If Len(chkCampo(i).DataField) > 0 Then
            chkCampo(i).Locked = Not (vgOk And chkCampo(i).Editable)
         End If
      Next
      ExecutaPreValidacao
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
   vgDataDeleted = mdiIRRIG.ExcluiRegistro()
End Sub
   
Private Sub grdBrowse_InitEdit(CancelEdit As Boolean)
   Reposition
End Sub

Private Sub grdBrowse_ItemSelect(ByVal vgItem As Long, vgColumns() As Variant)
   If vgPriVez Or Not grdBrowse.Visible Then Exit Sub
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

   
'executa a pré-validação da coluna do grid do modo grade do formulário
Private Sub grdBrowse_GetColumnLocked(ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   ExecutaPreValidacao                            'checa as pré-validações
   vgLocked = Not FormField.Enabled               'aplica as definições de pré-validação que são aplicadas ao campo da tela
End Sub


Private Sub grdBrowse_SaveData(ByVal vgItem As Long, vgColumns() As Variant, vgDataSaved As Boolean, vgColumn As Integer, vgErrorMessage As String)
   mdiIRRIG.SalvaDados vgColumn
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
End Sub
   
Private Sub grdBrowse_StatusChanged(ByVal vgNewStatus As Integer)
   If (vgNewStatus = ACAO_EXCLUINDO And Val(grdBrowse.Recordset.BookMark) >= 0) Then
      Reposition
   End If
   PrepBotoes Me, vgNewStatus                          'acerta icones dos botoes
   mdiIRRIG.RemontaForm                                'remonta dos os form da tela
End Sub

'apresenta popup menu para trabalhar com o grid
Private Sub grdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal vgCurCol As Integer)
   MostraPopGrid Me, Button
End Sub

'liga/desliga flag de repetição do último reg visualizado
Public Sub LigaDesligaAlterar()
   vgAlterar = Not vgAlterar          
   vgUltAlterar = vgAlterar                            'guarda situação de "pode alterar"
   AnalisaCondicoes                                    'vamos atualizar as condições para inclusão, exclusão, alteração...
   PrepBotoes Me, vgSituacao                           'acerta icones dos botoes
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

Public Sub CancelGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).CancelEdit
      End If
   Next
End Sub

Public Sub SaveGrids()
   Dim i As Integer
   For i = 0 To Grid.Count - 1
      If Grid(i).Status <> ACAO_NAVEGANDO Then
         Grid(i).SaveEdit
      End If
   Next
End Sub

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   Dim i As Integer, x As String, MudouFiltro As Boolean, vgCols() As Variant
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then vgPegouSeq = False
   Set Serie_Pivos = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse) Then
      Grid(0).CloseRecordset
   Else
      x$ = ExecutaGrid(0, vgCols(), ABRE_TABELA_GRID)
      MudouFiltro = (x$ <> Grid(0).SQLSource)
      If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(0).Status = ACAO_NAVEGANDO Then
         If Len(Grid(0).Recordset.RsSource) > 0 Then
            Grid(0).CloseRecordSet
         End If
         Grid(0).OpenRecordSet x$, CURSOR_TABLE
      End If
   End If
   Set mmCampo(1).Picture = Nothing: BuscaFotoAux
   ExecutaPreValidacao
   vgTemAlteracaoGrids = Not LockGrids
   Executar PODE_ALTERAR
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (((((Not Vazio(Modelo_do_Pivo)) AND Not Vazio(Serie_do_Pivo)) AND Not Vazio(MesAno)) AND Not Vazio(Letra_do_Pivo)) AND Entregue = 0)
   Botao(0).Enabled = Ok
   Ok = (Entregue)
   chkCampo(1).Locked = Not (vgAlterar And chkCampo(1).Editable)
   chkCampo(1).Enabled = Ok Or Not vgAlterar
   Ok = (Entrega_Tecnica = "Irrigação Penapolis")
   Label(17).Enabled = Ok And vgAlterar
   Ok = (Entrega_Tecnica = "Irrigação Penapolis")
   Label(18).Enabled = Ok And vgAlterar
   Ok = (vgPWUsuario = "JERONIMO" Or vgPWUsuario = "YGOR")
   txtCampo(17).Locked = Not (vgAlterar And txtCampo(17).Editable)
   txtCampo(17).Enabled = Ok Or Not vgAlterar
   Ok = (vgPWUsuario = "JERONIMO" Or vgPWUsuario = "YGOR")
   txtCampo(18).Locked = Not (vgAlterar And txtCampo(18).Editable)
   txtCampo(18).Enabled = Ok Or Not vgAlterar
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 1 Or Index = 5 Or Index = 8 Or Index = 12 Or Index = 13 Or Index = 14 Or Index = 15 Then
      ExecutaPreValidacao
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 8
            PoeRelEFiltroCbo 8
         Case 12
            PoeRelEFiltroCbo 12
         Case 14
            PoeRelEFiltroCbo 14
         Case 15
            PoeRelEFiltroCbo 15
         Case 16
            If Len(txtCp(16).Text) = 0 Then
               txtCampo(16).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
      End Select
   End If
   txtCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn And vgSituacao <> ACAO_NAVEGANDO Then  'se tela esta em edição e digitou ENTER
      ExecutaPreValidacao                                          'habilita/desabilita campos
   End If
   txtCampo(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyPress(Index As Integer, KeyAscii As Integer)
   txtCampo(Index).KeyPress KeyAscii
End Sub

'evento - quando o campo perder o foco
Private Sub txtCp_LostFocus(Index As Integer)
   txtCampo(Index).LostFocus
   If vgSituacao <> ACAO_NAVEGANDO Then           'se tela esta em edição
      InicializaApelidos COM_TEXTBOX              'pega apelidos dos campos
      ExecutaPreValidacao                         'habilita/desabilita campos
   End If
   Select Case Index
      Case 14
         BuscaGeral  
   End Select
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Entregue
      ElseIf Index = 1 Then
         chkCampo(1).Value = Entrega_Tec
      End If
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      ExecutaPreValidacao                         'habilita/desabilita campos
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
   Caption = LoadGasString(76200)
   vgFormID = 1186
   vgIdentTab = "Serie Pivos"
   vgFiltroEmUso = -1
   vgIndDefault = "Seq do Pivo"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgRepeticao = -99
   vgAlterar = False
   vgUltAlterar = False
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 22
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(67705))
   Set NroSerie = txtCampo(7)
   Set txtCodigo_do_Geral = txtCampo(8)
   Set txtCodigo_do_Vendedor = txtCampo(12)
   Set lblAjuste = Label(16)
   Set vgTb = New GRecordSet
   If Len(vgFiltroOriginal$) = 0 And Len(vgFiltroInicial$) > 0 Then
      vgFiltroOriginal$ = vgFiltroInicial$
   End If
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(76205)
   Label(1).Caption = LoadGasString(76206)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(76207)
   Label(2).Caption = LoadGasString(76208)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(76209)
   Label(3).Caption = LoadGasString(76210)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(76211)
   Label(4).Caption = LoadGasString(76212)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(76213)
   Label(5).Caption = LoadGasString(76214)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(76215)
   Label(6).Caption = LoadGasString(76216)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(76217)
   Label(7).Caption = LoadGasString(76218)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(76219)
   Label(8).Caption = LoadGasString(76220)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(76221)
   Botao(0).Caption = LoadGasString(76222)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(76223)
   chkCampo(0).Caption = LoadGasString(76224)
   Label(9).Caption = LoadGasString(76225)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(76226)
   Label(10).Caption = LoadGasString(76227)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(76228)
   Label(11).Caption = LoadGasString(76229)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(76230)
   Label(12).Caption = LoadGasString(76231)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(76232)
   Label(13).Caption = LoadGasString(76233)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(76234)
   Label(14).Caption = LoadGasString(76235)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(76236)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(76237)
   chkCampo(1).Caption = LoadGasString(76238)
   Label(15).Caption = LoadGasString(76239)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(76240)
   Label(16).Caption = LoadGasString(76241)
   Label(17).Caption = LoadGasString(76242)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(76243)
   Label(18).Caption = LoadGasString(76244)
   vgTooltips.AddTool txtCampo(16).CtPri, 0, LoadGasString(76245)
   Label(19).Caption = LoadGasString(76246)
   Label(20).Caption = LoadGasString(76247)
   vgTooltips.AddTool txtCampo(17).CtPri, 0, LoadGasString(76248)
   Label(21).Caption = LoadGasString(76249)
   vgTooltips.AddTool txtCampo(18).CtPri, 0, LoadGasString(76250)
   Botao(1).Caption = LoadGasString(76251)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS
      .AllowInsert = Permitido("Manutenção Pivo", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Manutenção Pivo", ACAO_EDITANDO)
      .AllowDelete = Permitido("Manutenção Pivo", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Seq", "Sequencia do Item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 975
      .AddColumn Nothing, , "Data", "Data da Manutenção", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1995
      .AddColumn Nothing, , "Descrição da Manutenção", "Descrição da Manutenção", TP_CARACTER, "@x", 120, False, , , , , , , , , , "0", , , , "0", 1, "0", 8220
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Serie_Pivos = vgTb
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
 grdBrowse.AddControlIgnoreFocus mdiIRRIG.botCancela.hwnd           'não deixa o grid tentar gravar automaticamente
 grdBrowse.AddControlIgnoreFocus mdiIRRIG.botSalva.hwnd             'se estiver perdendo o foco para esses botões
   grdBrowse.AllowDelete = True
   grdBrowse.AllowEdit = vgAlterar
   grdBrowse.SpecialPopupDisabled POP_GRID_BARS Or POP_GRID_STRIPES

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "999999"
   txtCampo(0).Editable = False
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 0
   txtCampo(1).ListFields = "IP3000|IP3001|IP2000|IP2001|IP1000|IP1001|IP4000"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 0
   txtCampo(2).Mask = "@x"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "9999"
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "99/99"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 0
   txtCampo(5).ListFields = "A|V"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_ACAO) = bottxtCampo6(BOT_ACAO)
   Set bottxtCampo6(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(6).DataType = 2
   txtCampo(6).Mask = "99/99/9999"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = "@!"
   txtCampo(7).Editable = False
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_LISTA) = bottxtCampo8(BOT_LISTA)
   Set txtCampo(8).CtBot(BOT_COMBO) = bottxtCampo8(BOT_COMBO)
   bottxtCampo8(BOT_LISTA).Caption = "P"
   Set bottxtCampo8(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(8).DataType = 1
   txtCampo(8).Mask = "999999"
   txtCampo(8).PesqModoAbertura =  2 
   txtCampo(8).PesqFieldCapture = "Geral.[Seqüência do Geral]"
   txtCampo(8).BoundColumn = "Seqüência do Geral"
   txtCampo(8).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(8).OrderFields = "Razão Social"
   txtCampo(8).Relation = ""
   txtCampo(8).Source = "Geral"
   txtCampo(8).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   Set txtCampo(9).CtFdo = labFdo9
   Set txtCampo(9).CtBot(BOT_ACAO) = bottxtCampo9(BOT_ACAO)
   Set bottxtCampo9(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(9).DataType = 2
   txtCampo(9).Mask = "99/99/9999"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 0
   txtCampo(10).Mask = "@x"
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 4
   txtCampo(11).Mask = ""
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(12).CtPri = txtCp(12)
   Set txtCampo(12).CtFdo = labFdo12
   Set txtCampo(12).CtBot(BOT_COMBO) = bottxtCampo12(BOT_COMBO)
   Set bottxtCampo12(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(12).DataType = 1
   txtCampo(12).Mask = "999999"
   txtCampo(12).BoundColumn = "Seqüência do Geral"
   txtCampo(12).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(12).OrderFields = "Razão Social"
   txtCampo(12).Relation = ""
   txtCampo(12).Source = "Geral"
   txtCampo(12).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(13).CtPri = txtCp(13)
   Set txtCampo(13).CtFdo = labFdo13
   Set txtCampo(13).CtBot(BOT_COMBO) = bottxtCampo13(BOT_COMBO)
   Set bottxtCampo13(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(13).DataType = 0
   txtCampo(13).ListFields = "Irrigação Penapolis|Revenda"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(1).DataField), chkCampo(1)

   Set txtCampo(14).CtPri = txtCp(14)
   Set txtCampo(14).CtFdo = labFdo14
   Set txtCampo(14).CtBot(BOT_LISTA) = bottxtCampo14(BOT_LISTA)
   Set txtCampo(14).CtBot(BOT_COMBO) = bottxtCampo14(BOT_COMBO)
   bottxtCampo14(BOT_LISTA).Caption = "P"
   Set bottxtCampo14(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(14).DataType = 1
   txtCampo(14).Mask = "999999"
   txtCampo(14).PesqModoAbertura =  2 
   txtCampo(14).PesqFieldCapture = "Orçamento.[Seqüência do Orçamento]"
   txtCampo(14).BoundColumn = "Seqüência do Orçamento"
   txtCampo(14).ListFields = "Seqüência do Orçamento"
   txtCampo(14).OrderFields = "Seqüência do Orçamento-"
   txtCampo(14).Relation = ""
   txtCampo(14).Source = "Orçamento"
   txtCampo(14).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(15).CtPri = txtCp(15)
   Set txtCampo(15).CtFdo = labFdo15
   Set txtCampo(15).CtBot(BOT_COMBO) = bottxtCampo15(BOT_COMBO)
   Set bottxtCampo15(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(15).DataType = 1
   txtCampo(15).Mask = "999"
   txtCampo(15).BoundColumn = "Codigo do solicitante"
   txtCampo(15).ListFields = "Nome do solicitante"
   txtCampo(15).OrderFields = "Nome do solicitante"
   txtCampo(15).Relation = ""
   txtCampo(15).Source = "Solicitantes"
   txtCampo(15).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(15).DataField), txtCampo(15)

   Set txtCampo(16).CtPri = txtCp(16)
   Set txtCampo(16).CtFdo = labFdo16
   Set txtCampo(16).CtBot(BOT_ACAO) = bottxtCampo16(BOT_ACAO)
   Set bottxtCampo16(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(16).DataType = 2
   txtCampo(16).Mask = "99/99/9999"
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

   Set txtCampo(17).CtPri = txtCp(17)
   Set txtCampo(17).CtFdo = labFdo17
   Set txtCampo(17).CtBot(BOT_ACAO) = bottxtCampo17(BOT_ACAO)
   Set bottxtCampo17(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(17).DataType = 2
   txtCampo(17).Mask = "99/99/9999"
   txtCampo(17).BoundColumn = ""
   txtCampo(17).ListFields = ""
   txtCampo(17).OrderFields = ""
   txtCampo(17).Relation = ""
   txtCampo(17).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(17).DataField), txtCampo(17)

   Set txtCampo(18).CtPri = txtCp(18)
   txtCampo(18).DataType = 4
   txtCampo(18).Mask = ""
   txtCampo(18).BoundColumn = ""
   txtCampo(18).ListFields = ""
   txtCampo(18).OrderFields = ""
   txtCampo(18).Relation = ""
   txtCampo(18).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(18).DataField), txtCampo(18)

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
      Case 8
         txtCampo(8).Filter = "(Inativo = " & 0 & ") AND [Seqüência do Geral] > " & 0 & ""
                           txtCampo(8).PesqSQLExpression = "SELECT Geral.[Seqüência do Geral], Geral.[Razão Social] FROM Geral WHERE (Geral.Inativo = " + CStr(0) + ") AND " + _
                                                              "(Geral.[Seqüência do Geral] > " + CStr(0) + ")"
      Case 12
         txtCampo(12).Filter = "(Vendedor = " & 1 & ") AND Inativo = " & 0 & ""
      Case 14
         txtCampo(14).Filter = "[Seqüência do Orçamento] > " & 0 & " AND Cancelado = " & 0 & ""
                           txtCampo(14).PesqSQLExpression = "SELECT Orçamento.[Seqüência do Orçamento], Orçamento.[Data de Emissão], Orçamento.[Vendedor Intermediario], " + _
                                                               "Geral.[Razão Social] FROM Orçamento, Geral WHERE (Orçamento.[Seqüência do Orçamento] > " + CStr(0) + ") AND " + _
                                                               "(Orçamento.[Seqüência do Geral] = Geral.[Seqüência do Geral]) AND (Orçamento.Cancelado = " + CStr(0) + ")"
      Case 15
         txtCampo(15).Filter = "([Codigo do solicitante] > " & 0 & ") AND [Codigo do setor] = " & 13 & ""
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
   Set NroSerie = Nothing
   Set txtCodigo_do_Geral = Nothing
   Set txtCodigo_do_Vendedor = Nothing
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   If Not Serie_Pivos Is Nothing Then
      Set Serie_Pivos = Nothing
   End If
   If Not Manutencao_Pivo Is Nothing Then
      Manutencao_Pivo.CloseRecordSet
      Set Manutencao_Pivo = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmPivos = Nothing                         'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no grid filho
Private Sub Grid_KeyPress(Index As Integer, ByVal KeyAscii As Integer, ByVal Shift As Integer, vgColumns() As Variant)
   ExecutaGrid0 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
End Sub

'evento - após efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, False
End Sub

'evento - antes de efetuar o edit do recordset do grid
Private Sub Grid_BeforeEditRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, True
   ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Manutencao_Pivo![Seq do Pivo] = Serie_Pivos![Seq do Pivo]
End Sub

'evento - antes de efetuar o delete no recordset do grid
Private Sub Grid_BeforeDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, True
   ExecutaGrid Index, vgColumns(), EXCLUSOES, vgItem, 0, vgIsValid, 0, vgErrorMessage
End Sub

'evento - quer pegar valores para cada célula
Private Sub Grid_GetColumnValue(Index As Integer, ByVal vgItem As Long, ByVal vgCol As Integer, vgColumns() As Variant, vgNewText As Variant)
   Dim RetVal As Variant, NCol As Integer
   RetVal = ExecutaGrid(Index, vgColumns(), CONTEUDODACOLUNA, vgItem, vgCol, , NCol)
   If NCol = -1 Then
      vgNewText = RetVal
   End If
End Sub

'evento - recordset do grid foi mudado
Private Sub Grid_RecordSetChanged(Index As Integer, ByVal vgNewRecordSet As GRecordSet)
   Set Manutencao_Pivo = vgNewRecordSet
End Sub

'evento - quer validar dados, está gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - após a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   mdiIRRIG.RemontaForm                                   'vamos atualizar os forms de dados
   Grid(Index).Repaint -1                                 'atualiza dados do grid (registro posicionado) 
End Sub

Private Sub Grid_ControlButtonClick(Index As Integer)
   Grid(Index).ShowFilterBar = Not Grid(Index).ShowFilterBar
End Sub

Private Sub Grid_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And Grid(Index).Status = ACAO_NAVEGANDO Then                 'o formulário pai não está em navegação
      mdiIRRIG.SalvaDados                         'salva o resitro em edição
      If vgSituacao <> ACAO_NAVEGANDO And ActiveControl Is Grid(Index) Then 'se não gravou e ainda está com foco no grid
         FocoNoPriControle Me                                               'vamos colocar foco no primeiro controle do pai
      End If
   End If
End Sub

Private Sub Grid_StatusChanged(Index As Integer, ByVal vgNewStatus As Integer)
   Dim vgTemAltGrdOld As Boolean
   If vgNewStatus <> ACAO_NAVEGANDO Then vgNewStatus = - vgNewStatus
   PrepBotoes Me, vgNewStatus                                     'acerta icones dos botoes
   vgTemAltGrdOld = vgTemAlteracaoGrids
   mdiIRRIG.RemontaForm                                           'remonta dos os form da tela
   If vgSituacao = ACAO_NAVEGANDO Then
      Reposition , Not vgTemAltGrdOld
   End If
End Sub

'evento - atualiza valores para os filtros das colunas do grid filho
Private Sub Grid_GetColumnFilter(Index As Integer, ByVal vgColumn As Integer, vgColumns() As Variant, vgFilter As String)
   vgFilter = ExecutaGrid(Index, vgColumns(), PEGAFILTRODASCOLUNAS, , vgColumn)
End Sub

'evento - pega expressão SQL para abertura de pesquisa
Private Sub Grid_GetColumnSQLSearch(Index As Integer, ByVal vgColumn As Integer, vgColumns() As Variant, vgSQLSearch As String)
   vgSQLSearch = ExecutaGrid(Index, vgColumns(), PEGAEXPRESSAOPESQUISA, , vgColumn)
End Sub

'inicializações, validações e processos para o grid
Private Function ExecutaGrid(Index As Integer, ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Select Case Index
      Case  0 
         ExecutaGrid = ExecutaGrid0(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_Item As Long, Data_da_Manutencao As Variant, Descricao_da_Manutencao As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(0).Status <> ACAO_NAVEGANDO And vgItem = Grid(0).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      BuscaFoto KeyCodeAscii, Seq_do_Pivo, Sequencia_do_Item, Data_da_Manutencao, Descricao_da_Manutencao 
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Manutenção Pivo]"

      'definindo a expressão de ligação com o pai
      x$ = "[Seq do Pivo] = " & Serie_Pivos![Seq do Pivo]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Sequencia_do_Item = SeqItem()
            vgRetVal = Sequencia_do_Item
         Case 2
            Data_da_Manutencao = Date
            If Grid(0).Col = 2 Then
               vgRetVal = Data_da_Manutencao
            End If
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = ((Not Vazio(Data_da_Manutencao)) And (IsDate(Data_da_Manutencao) Or Vazio(Data_da_Manutencao)))
         If Not vgIsValid Then vgColumn =  2 
         vgErrorMessage$ = "Data da Manutenção não pode ser vazio!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Descricao_da_Manutencao))
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Descrição da Manutenção não pode ser vazio!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = SeqItem()
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Item = ColumnValue(1)
   Data_da_Manutencao = ColumnValue(2)
   Descricao_da_Manutencao = ColumnValue(3) & ""
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Manutencao_Pivo.Eof = False And Manutencao_Pivo.BOF = False And Manutencao_Pivo.RecordCount > 0 Then
      Sequencia_do_Item = Manutencao_Pivo![Sequencia do Item]
      Data_da_Manutencao = Manutencao_Pivo![Data da Manutenção]
      Descricao_da_Manutencao = Manutencao_Pivo![Descrição da Manutenção]
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

DeuErro:
   If vgOq = CONTEUDODACOLUNA Or vgOq = DEFAULTDASCOLUNAS Or vgOq < 0 Then
      vgRetVal = Null
   Else
      vgErrorMessage$ = Err.Source + "|" + Trim$(Str$(Err)) + "-" + Error$
      vgIsValid = False
   End If
   If Not vgRsError Is Nothing Then
      vgRsError.CancelUpdate
      vgErrorMessage$ = vgRsError.Table & "=>" & vgErrorMessage$
      Set vgRsError = Nothing
   End If
   Resume ResumeErro

ResumeErro:
   On Error Resume Next

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function

'evento - pega o valor inicial das colunas do grid
Private Sub Grid_GetColumnDefaultValue(Index As Integer, ByVal vgCol As Integer, vgColumns() As Variant, ByRef vgDefaultValue As Variant)
   vgDefaultValue = ExecutaGrid(Index, vgColumns(), DEFAULTDASCOLUNAS, , vgCol)
End Sub

'evento - quer pegar valores para cada célula
Private Sub Grid_GetColumnLocked(Index As Integer, ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         GeraNroSerie  
      Case 1
         mdiIRRIG.MRelSPivo
         AtivaForm Me
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo6_Click(Index As Integer)
   txtCampo(6).SetFocus
   DoEvents
   txtCampo(6).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo8_Click(Index As Integer)
   txtCampo(8).SetFocus
   DoEvents
   txtCampo(8).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo9_Click(Index As Integer)
   txtCampo(9).SetFocus
   DoEvents
   txtCampo(9).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo12_Click(Index As Integer)
   txtCampo(12).SetFocus
   DoEvents
   txtCampo(12).botClick Index
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
Private Sub bottxtCampo15_Click(Index As Integer)
   txtCampo(15).SetFocus
   DoEvents
   txtCampo(15).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo16_Click(Index As Integer)
   txtCampo(16).SetFocus
   DoEvents
   txtCampo(16).botClick Index
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


