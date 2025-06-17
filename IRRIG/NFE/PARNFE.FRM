Version 5.00
Begin VB.Form frmParNFe
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   6285
   Left                     =   345
   LinkTopic                =   "ParNFe"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   5850
   ScaleWidth               =   10755
   Top                      =   1435
   Width                    =   10845
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
      Height                   =   5685
      Left                     =   15
      Width                    =   10575
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   18643
      _extenty                 =   10022
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1290
         Left                     =   0
         Width                    =   1455
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
         Height                   =   435
         Left                     =   1680
         Width                    =   3405
         Top                      =   420
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
         Caption                  =   "Parametros da NFe"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1740
         Y1                       =   870
         X2                       =   5055
         Y2                       =   870
         Index                    =   0 
      End
      Begin VB.Frame Frame
         Height                   =   630
         Left                     =   6600
         Width                    =   3705
         Top                      =   120
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
         Caption                  =   ""
         Index                    =   0 
         Begin             VB.Label labopcFrame0
            BorderStyle              =   1
            DataField                =   "Ambiente"
            Height                   =   120
            Left                     =   1350
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcFrame0Cp
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
            Alignment                =   2
            TabIndex                 =   11
            TabStop                  =   0
            Height                   =   255
            Left                     =   120
            Width                    =   1470
            Top                      =   240
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
            Caption                  =   "Produção"
            Index                    =   0 
         End
         Begin VB.OptionButton opcFrame0Cp
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
            Alignment                =   2
            TabIndex                 =   12
            TabStop                  =   0
            Height                   =   255
            Left                     =   1620
            Width                    =   1950
            Top                      =   240
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
            Caption                  =   "Homologação"
            Index                    =   1 
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
         TabIndex                 =   6
         TabStop                  =   0
         Height                   =   255
         Left                     =   5700
         Width                    =   2235
         Top                      =   960
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
         Caption                  =   "Horario de Verão"
         Index                    =   0 
         DataField                =   "Horario de verao"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   1305
         Top                      =   1740
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
         Caption                  =   "Série da NFe:"
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
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1710
         Width                    =   375
         Top                      =   1740
         DataField                =   "Serie da nfe"
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
         Left                     =   2220
         Width                    =   1365
         Top                      =   1740
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
         Caption                  =   "Codigo da UF:"
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
         ForeColor                =   &H80000008
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   300
         Left                     =   3660
         Width                    =   420
         Top                      =   1740
         DataField                =   "Codigo da uf"
         Index                    =   1 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   4200
         Width                    =   1935
         Top                      =   1740
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
         Caption                  =   "% Simples Nacional:"
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
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6240
         Width                    =   960
         Top                      =   1740
         DataField                =   "Aliquota simples nacional"
         Index                    =   2 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   3825
         Top                      =   2340
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
         Caption                  =   "Diretorio Xml Gerados Homologação.........:"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7320
         Width                    =   1695
         Top                      =   1740
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
         Caption                  =   "% Aprox. Tributos:"
         Index                    =   5 
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
         TabIndex                 =   7
         TabStop                  =   0
         Height                   =   255
         Left                     =   1740
         Width                    =   2295
         Top                      =   960
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
         Caption                  =   "Simples Nacional"
         Index                    =   1 
         DataField                =   "Simples nacional"
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
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   9105
         Width                    =   1170
         Top                      =   1740
         DataField                =   "Percentual tributos"
         Index                    =   3 
         MultiLine                =   -1
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4200
         Top                      =   2340
         Width                    =   6105
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10035
         Top                      =   2355
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
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4215
         Width                    =   5805
         Top                      =   2355
         DataField                =   "Diretorio 2 nfe homologação"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   3885
         Top                      =   2880
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
         Caption                  =   "Diretorio Xml Autorizados Homologação...:"
         Index                    =   6 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4200
         Top                      =   2880
         Width                    =   6105
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10035
         Top                      =   2895
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
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4215
         Width                    =   5805
         Top                      =   2895
         DataField                =   "Diretorio 1 nfe homologação"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   3855
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
         Caption                  =   "Diretorio Xml Gerados Produção..................:"
         Index                    =   7 
      End
      Begin VB.Label labFdo6
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4200
         Top                      =   3420
         Width                    =   6105
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10035
         Top                      =   3435
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
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4215
         Width                    =   5805
         Top                      =   3435
         DataField                =   "Diretorio 2 nfe produção"
         Index                    =   6 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   3810
         Top                      =   3960
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
         Caption                  =   "Diretorio Xml Autorizados Produção............:"
         Index                    =   8 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4200
         Top                      =   3960
         Width                    =   6105
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10035
         Top                      =   3975
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
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4215
         Width                    =   5805
         Top                      =   3975
         DataField                =   "Diretorio 1 nfe produção"
         Index                    =   7 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   3825
         Top                      =   4500
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
         Caption                  =   "Certificado Digital................................................:"
         Index                    =   9 
      End
      Begin VB.Label labFdo8
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   4200
         Top                      =   4500
         Width                    =   6105
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10035
         Top                      =   4515
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
         TabIndex                 =   10
         TabStop                  =   -1
         Height                   =   270
         Left                     =   4215
         Width                    =   5805
         Top                      =   4515
         DataField                =   "Certificado digital"
         Index                    =   8 
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
         TabIndex                 =   13
         TabStop                  =   -1
         Height                   =   255
         Left                     =   4140
         Width                    =   1380
         Top                      =   960
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
         Caption                  =   "ME / EPP"
         Index                    =   2 
         DataField                =   "Micro Empresa"
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
         TabIndex                 =   14
         TabStop                  =   0
         Height                   =   255
         Left                     =   1740
         Width                    =   3270
         Top                      =   1260
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
         Caption                  =   "Usar Responsável Técnico"
         Index                    =   3 
         DataField                =   "Usa Resp Tecnico"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   315
         Width                    =   3825
         Top                      =   4995
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
         Caption                  =   "CNPJ do Contador.............................................:"
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
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   15
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4200
         Width                    =   2115
         Top                      =   4995
         DataField                =   "CNPJ do Contador"
         Index                    =   9 
      End
   End
End
Attribute VB_name = "frmParNFe"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: ParNFe
'* Função....: Parametros da NFe
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
Dim txtCampo(9) As New FormataCampos              'classe dos campos tipo texto do módulo  
Dim chkCampo(3) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim opcFrame0(1) As New FormataCampos
Dim Codigo_da_uf As Integer, Serie_da_nfe As Integer, Ambiente As Integer
Dim Horario_de_verao As Boolean, Aliquota_simples_nacional As Double, Diretorio_1_nfe_homologacao As String
Dim Diretorio_2_nfe_homologacao As String, Diretorio_1_nfe_producao As String, Diretorio_2_nfe_producao As String
Dim Certificado_digital As String, Simples_nacional As Boolean, Percentual_tributos As Double
Dim Micro_Empresa As Boolean, Usa_Resp_Tecnico As Boolean, CNPJ_do_Contador As String
Public txtDir2NFeGeradosHomologacao As Object, txtDir1NFeAutHomologacao As Object, txtDir2NFeGeradosProducao As Object
Public txtDirNFeAutProducao As Object, txtCertificado As Object

Private Sub EscolhaCertificado()
   Dim Resultado As Long, msgResultado As String, certificado As String
   Dim objNFeUtil As Object
   
   On Error Resume Next

   Set objNFeUtil = CreateObject("NFe_util_2G.util")

   ' pega certificado
   Resultado = objNFeUtil.PegaNomeCertificado(certificado, msgResultado)
   If Resultado < 5403 Then
      If InStr(1, certificado, "Associacao", vbTextCompare) > 0 Then
         MsgBox "O certificado digital da Associação não é um certificado válido para consumir os WS da NFe! Procure adquirir um certificado digital válido para prosseguir...", vbCritical + vbOkOnly, vaTitulo
      End If
      txtCertificado.Text = certificado     ' escolhido um certificado digital
   Else
      MsgBox msgResultado, vbCritical + vbOkOnly, vaTitulo
   End If
  
   Set objNFeUtil = Nothing
   
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Codigo_da_uf = txtCampo(1).Value
         Serie_da_nfe = txtCampo(0).Value
         Ambiente = Val(labopcFrame0.Caption)
         Horario_de_verao = chkCampo(0).Value
         Aliquota_simples_nacional = txtCampo(2).Value
         Diretorio_1_nfe_homologacao = txtCampo(5).Value
         Diretorio_2_nfe_homologacao = txtCampo(4).Value
         Diretorio_1_nfe_producao = txtCampo(7).Value
         Diretorio_2_nfe_producao = txtCampo(6).Value
         Certificado_digital = txtCampo(8).Value
         Simples_nacional = chkCampo(1).Value
         Percentual_tributos = txtCampo(3).Value
         Micro_Empresa = chkCampo(2).Value
         Usa_Resp_Tecnico = chkCampo(3).Value
         CNPJ_do_Contador = txtCampo(9).Value
      Else
         Codigo_da_uf = IIf(IsNull(vgTb![Codigo da uf]), 0, vgTb![Codigo da uf])
         Serie_da_nfe = IIf(IsNull(vgTb![Serie da nfe]), 0, vgTb![Serie da nfe])
         Ambiente = IIf(IsNull(vgTb!Ambiente), 0, vgTb!Ambiente)
         Horario_de_verao = IIf(IsNull(vgTb![Horario de verao]), 0, vgTb![Horario de verao])
         Aliquota_simples_nacional = IIf(IsNull(vgTb![Aliquota simples nacional]), 0, vgTb![Aliquota simples nacional])
         Diretorio_1_nfe_homologacao = IIf(IsNull(vgTb![Diretorio 1 nfe homologação]), "", vgTb![Diretorio 1 nfe homologação])
         Diretorio_2_nfe_homologacao = IIf(IsNull(vgTb![Diretorio 2 nfe homologação]), "", vgTb![Diretorio 2 nfe homologação])
         Diretorio_1_nfe_producao = IIf(IsNull(vgTb![Diretorio 1 nfe produção]), "", vgTb![Diretorio 1 nfe produção])
         Diretorio_2_nfe_producao = IIf(IsNull(vgTb![Diretorio 2 nfe produção]), "", vgTb![Diretorio 2 nfe produção])
         Certificado_digital = IIf(IsNull(vgTb![Certificado digital]), "", vgTb![Certificado digital])
         Simples_nacional = IIf(IsNull(vgTb![Simples nacional]), 0, vgTb![Simples nacional])
         Percentual_tributos = IIf(IsNull(vgTb![Percentual tributos]), 0, vgTb![Percentual tributos])
         Micro_Empresa = IIf(IsNull(vgTb![Micro Empresa]), 0, vgTb![Micro Empresa])
         Usa_Resp_Tecnico = IIf(IsNull(vgTb![Usa Resp Tecnico]), 0, vgTb![Usa Resp Tecnico])
         CNPJ_do_Contador = IIf(IsNull(vgTb![CNPJ do Contador]), "", vgTb![CNPJ do Contador])
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
      If vgPWUsuario = "ADMIN" Then
         vgOk = (Codigo_da_uf > 0)
         vgMsg$ = "Codigo da uf inválido!"
         If Not vgOk Then txtCampo(1).SetFocus
      End If
      If vgOk Then
         vgOk = (VCGC(CNPJ_do_Contador))
         vgMsg$ = "CNPJ do Contador inválido!"
         If Not vgOk Then txtCampo(9).SetFocus
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
                  opcFrame0(1).Value = True
         opcFrame0(1).CtPri.TabStop = False
         chkCampo(0).Value = False
         txtCampo(0).Value = 0
         txtCampo(1).Value = 35
         chkCampo(1).Value = False
         chkCampo(2).Value = False
         chkCampo(3).Value = False
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
         opcFrame0(Ambiente).Value = True
         opcFrame0(Ambiente).CtPri.TabStop = False
         chkCampo(0).Value = Horario_de_verao
         chkCampo(1).Value = Simples_nacional
         chkCampo(2).Value = Micro_Empresa
         chkCampo(3).Value = Usa_Resp_Tecnico
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
      vgTb!Ambiente = Ambiente
      vgTb![Horario de verao] = Horario_de_verao
      vgTb![Simples nacional] = Simples_nacional
      vgTb![Micro Empresa] = Micro_Empresa
      vgTb![Usa Resp Tecnico] = Usa_Resp_Tecnico
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
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
   ExecutaPreValidacao
   vgTemAlteracaoGrids = Not LockGrids
DeuErro:
   
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (vgPWUsuario = "ADMIN")
   opcFrame0(0).Enabled = (Ok And opcFrame0(0).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   opcFrame0(1).Enabled = (Ok And opcFrame0(1).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(1).Enabled = (Ok And txtCampo(1).Editable)
   Ok = (Simples_nacional = True)
   txtCampo(2).Enabled = (Ok And txtCampo(2).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   chkCampo(1).Enabled = (Ok And chkCampo(1).Editable)
   Ok = (Simples_nacional = True)
   txtCampo(3).Enabled = (Ok And txtCampo(3).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(4).Enabled = (Ok And txtCampo(4).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(5).Enabled = (Ok And txtCampo(5).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(6).Enabled = (Ok And txtCampo(6).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   txtCampo(7).Enabled = (Ok And txtCampo(7).Editable)
   Ok = (vgPWUsuario = "ADMIN")
   chkCampo(2).Enabled = (Ok And chkCampo(2).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 0
            If ValBrasil(txtCp(0).Text) = 0 Then
               txtCampo(0).Value = 0
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 1
            If ValBrasil(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = 35
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
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Horario_de_verao
      ElseIf Index = 1 Then
         chkCampo(1).Value = Simples_nacional

      ElseIf Index = 2 Then
         chkCampo(2).Value = Micro_Empresa

      ElseIf Index = 3 Then
         chkCampo(3).Value = Usa_Resp_Tecnico
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
         ExecutaPreValidacao
         opcFrame0(Index).Change
      End If
   End If
   opcFrame0(0).CtPri.TabStop = False
   opcFrame0(1).CtPri.TabStop = False
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
   Caption = LoadGasString(12490)
   vgFormID = 615
   vgIdentTab = "Parametros da nfe"
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
   vgUltimoTabIndex = 15
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(3995))
   Set txtDir2NFeGeradosHomologacao = txtCampo(4)
   Set txtDir1NFeAutHomologacao = txtCampo(5)
   Set txtDir2NFeGeradosProducao = txtCampo(6)
   Set txtDirNFeAutProducao = txtCampo(7)
   Set txtCertificado = txtCampo(8)
   Set vgTb = New GRecordSet
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(12495)
   vgTooltips.AddTool Frame(0), 0, LoadGasString(12496)
   vgTooltips.AddTool opcFrame0(0).CtPri, 0, LoadGasString(12497)
   opcFrame0(0).Caption = LoadGasString(12498)
   vgTooltips.AddTool opcFrame0(1).CtPri, 0, LoadGasString(12499)
   opcFrame0(1).Caption = LoadGasString(12500)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(12501)
   chkCampo(0).Caption = LoadGasString(12502)
   Label(1).Caption = LoadGasString(12503)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(12504)
   Label(2).Caption = LoadGasString(12505)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(12506)
   Label(3).Caption = LoadGasString(12507)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(12508)
   Label(4).Caption = LoadGasString(12509)
   Label(5).Caption = LoadGasString(12510)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(12511)
   chkCampo(1).Caption = LoadGasString(12512)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(12513)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(12514)
   Label(6).Caption = LoadGasString(12515)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(12516)
   Label(7).Caption = LoadGasString(12517)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(12518)
   Label(8).Caption = LoadGasString(12519)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(12520)
   Label(9).Caption = LoadGasString(12521)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(12522)
   vgTooltips.AddTool chkCampo(2).CtPri, 0, LoadGasString(12523)
   chkCampo(2).Caption = LoadGasString(12524)
   vgTooltips.AddTool chkCampo(3).CtPri, 0, LoadGasString(12525)
   chkCampo(3).Caption = LoadGasString(12526)
   Label(10).Caption = LoadGasString(12527)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(12528)
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

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 1
   txtCampo(0).Mask = "9"
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "99"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999,9999"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "999,9999"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_ACAO) = bottxtCampo4(BOT_ACAO)
   Set bottxtCampo4(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "@x"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_ACAO) = bottxtCampo5(BOT_ACAO)
   Set bottxtCampo5(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(5).DataType = 0
   txtCampo(5).Mask = "@x"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_ACAO) = bottxtCampo6(BOT_ACAO)
   Set bottxtCampo6(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = "@x"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_ACAO) = bottxtCampo7(BOT_ACAO)
   Set bottxtCampo7(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = "@x"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_ACAO) = bottxtCampo8(BOT_ACAO)
   bottxtCampo8(BOT_ACAO).Caption = "C"
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "@x"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""

   Set opcFrame0(0).CtPri = opcFrame0Cp(0)
   Set opcFrame0(0).CtFdo = labopcFrame0
   opcFrame0(0).DataType = 6
   opcFrame0(0).BookMark = 0

   Set opcFrame0(1).CtPri = opcFrame0Cp(1)
   Set opcFrame0(1).CtFdo = labopcFrame0
   opcFrame0(1).DataType = 6
   opcFrame0(1).BookMark = 1

   Set chkCampo(2).CtPri = chkCp(2)
   chkCampo(2).DataType = 5

   Set chkCampo(3).CtPri = chkCp(3)
   chkCampo(3).DataType = 5

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = "99.999.999/9999-99"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""

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
   AjustaPanFundo Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   FinalizaForm Me
   Set txtDir2NFeGeradosHomologacao = Nothing
   Set txtDir1NFeAutHomologacao = Nothing
   Set txtDir2NFeGeradosProducao = Nothing
   Set txtDirNFeAutProducao = Nothing
   Set txtCertificado = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   Set chkCampo(3) = Nothing

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   Set frmParNFe = Nothing                        'libera o segmento de código do form
End Sub

'evento - quando o botão for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 4
         PegaDir2 Me: txtDir2NFeGeradosHomologacao.Text = vaDir
      Case 5
         PegaDir2 Me: txtDir1NFeAutHomologacao.Text = vaDir
      Case 6
         PegaDir2 Me: txtDir2NFeGeradosProducao.Text = vaDir
      Case 7
         PegaDir2 Me: txtDirNFeAutProducao.Text = vaDir
      Case 8
         EscolhaCertificado
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo4_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 4, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(4).botClick Index
   End If
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 5, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(5).botClick Index
   End If
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo6_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 6, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(6).botClick Index
   End If
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo7_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 7, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(7).botClick Index
   End If
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo8_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 8, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(8).botClick Index
   End If
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


