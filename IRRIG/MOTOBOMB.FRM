Version 5.00
Begin VB.Form frmMotoBomb
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   7290
   Left                     =   3030
   LinkTopic                =   "MotoBomb"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   6855
   ScaleWidth               =   12060
   Top                      =   1690
   Width                    =   12150
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
      Height                   =   6675
      Left                     =   15
      Width                    =   11880
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   20944
      _extenty                 =   11768
      Index                    =   0 
      Begin IRRIG.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   -1
         Height                   =   1215
         Left                     =   120
         Width                    =   1410
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
         Height                   =   375
         Left                     =   1680
         Width                    =   4815
         Top                      =   300
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
         Caption                  =   "Número de Série (Motobomba)"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1740
         Y1                       =   682
         X2                       =   6510
         Y2                       =   682
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   630
         Top                      =   1920
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
         Caption                  =   "Seq...:"
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
         Left                     =   840
         Width                    =   690
         Top                      =   1920
         DataField                =   "Seq Moto Bomba"
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
         Width                    =   1380
         Top                      =   1920
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
         Caption                  =   "Tipo de Motor:"
         Index                    =   2 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3120
         Top                      =   1920
         Width                    =   660
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   3510
         Top                      =   1935
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
         Left                     =   3135
         Width                    =   360
         Top                      =   1935
         DataField                =   "Tp de Motor"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3900
         Width                    =   660
         Top                      =   1920
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
         Locked                   =   -1
         TabIndex                 =   2
         TabStop                  =   0
         Height                   =   300
         Left                     =   4560
         Width                    =   720
         Top                      =   1920
         DataField                =   "Serie da Moto Bomba"
         TabStop                  =  0
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
         Left                     =   5400
         Width                    =   1080
         Top                      =   1920
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
         Alignment                =   0
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6540
         Width                    =   660
         Top                      =   1920
         DataField                =   "MesAno"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7320
         Width                    =   540
         Top                      =   1920
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
         Caption                  =   "Tipo:"
         Index                    =   5 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7980
         Top                      =   1920
         Width                    =   780
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8490
         Top                      =   1935
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
         Left                     =   7995
         Width                    =   480
         Top                      =   1935
         DataField                =   "Função Moto Bomba"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8940
         Width                    =   1140
         Top                      =   1920
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
         Index                    =   6 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   10140
         Top                      =   1920
         Width                    =   1395
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   11265
         Top                      =   1935
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
         Left                     =   10155
         Width                    =   1095
         Top                      =   1935
         DataField                =   "Data de Criação"
         Index                    =   5 
      End
      Begin VB.Frame Frame
         Height                   =   1350
         Left                     =   105
         Width                    =   11430
         Top                      =   2340
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
            Width                    =   1350
            Top                      =   360
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
            Index                    =   7 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H8000000A
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
            TabIndex                 =   6
            TabStop                  =   0
            Height                   =   300
            Left                     =   1560
            Width                    =   1830
            Top                      =   360
            DataField                =   "Nro de Serie Moto Bomba"
            TabStop                  =  0
            Index                    =   6 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   7
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   300
            Left                     =   3480
            Width                    =   870
            Top                      =   360
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Courier New"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Gerar"
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   4500
            Width                    =   1395
            Top                      =   360
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
            Caption                  =   "Nº do Motor:"
            Index                    =   8 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H8000000A
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
            TabIndex                 =   8
            TabStop                  =   -1
            Height                   =   300
            Left                     =   5940
            Width                    =   1695
            Top                      =   360
            DataField                =   "Nro de Serie Motor"
            Index                    =   7 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   7800
            Width                    =   795
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
            Caption                  =   "Modelo:"
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
            MaxLength                =   40
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   9
            TabStop                  =   -1
            Height                   =   300
            Left                     =   8700
            Width                    =   2535
            Top                      =   360
            DataField                =   "Modelo do Motor"
            Index                    =   8 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   120
            Width                    =   1515
            Top                      =   840
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
            Caption                  =   "Nº da Bomba:"
            Index                    =   10 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H8000000A
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
            TabIndex                 =   10
            TabStop                  =   -1
            Height                   =   300
            Left                     =   1680
            Width                    =   2175
            Top                      =   840
            DataField                =   "Nro de Serie Bomba"
            Index                    =   9 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   4020
            Width                    =   795
            Top                      =   840
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
            MaxLength                =   40
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   11
            TabStop                  =   -1
            Height                   =   300
            Left                     =   4920
            Width                    =   2715
            Top                      =   840
            DataField                =   "Modelo da Bomba"
            Index                    =   10 
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   120
         Width                    =   600
         Top                      =   3840
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
         Index                    =   12 
      End
      Begin VB.Label labFdo11
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   780
         Top                      =   3840
         Width                    =   5520
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo11
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6030
         Top                      =   3855
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo11
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   5775
         Top                      =   3855
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
         TabIndex                 =   12
         TabStop                  =   -1
         Height                   =   270
         Left                     =   795
         Width                    =   4965
         Top                      =   3855
         DataField                =   "Codigo do Geral"
         Index                    =   11 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   6420
         Width                    =   1020
         Top                      =   3840
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
         BackColor                =  &H80000005
         Left                     =   7500
         Top                      =   3840
         Width                    =   4020
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo12
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11250
         Top                      =   3855
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
         TabIndex                 =   13
         TabStop                  =   -1
         Height                   =   270
         Left                     =   7515
         Width                    =   3720
         Top                      =   3855
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
         Left                     =   120
         Width                    =   600
         Top                      =   4320
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
         Index                    =   14 
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
         TabIndex                 =   14
         TabStop                  =   -1
         Height                   =   300
         Left                     =   780
         Width                    =   7320
         Top                      =   4320
         DataField                =   "NF"
         Index                    =   13 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8280
         Width                    =   1560
         Top                      =   4320
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
         Caption                  =   "Data de Entrega:"
         Index                    =   15 
      End
      Begin VB.Label labFdo14
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   9960
         Top                      =   4320
         Width                    =   1410
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   11100
         Top                      =   4335
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
         TabStop                  =   -1
         Height                   =   270
         Left                     =   9975
         Width                    =   1110
         Top                      =   4335
         DataField                =   "Dt da Entrega"
         Index                    =   14 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   120
         Width                    =   660
         Top                      =   4800
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
         Index                    =   16 
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
         TabIndex                 =   16
         TabStop                  =   -1
         Height                   =   1410
         Left                     =   780
         Width                    =   10740
         Top                      =   4800
         DataField                =   "Obs"
         Index                    =   15 
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
         Left                     =   7395
         Width                    =   1590
         Top                      =   615
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
         Index                    =   17 
      End
      Begin VB.Label labFdo16
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8955
         Top                      =   615
         Width                    =   2610
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo16
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11295
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
         TabIndex                 =   17
         TabStop                  =   0
         Height                   =   270
         Left                     =   8970
         Width                    =   2310
         Top                      =   630
         DataField                =   "Entrega Tecnica"
         Index                    =   16 
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
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   255
         Left                     =   8955
         Width                    =   2505
         Top                      =   315
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
         Index                    =   0 
         DataField                =   "Entrega Tec"
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
         TabIndex                 =   19
         TabStop                  =   0
         Height                   =   255
         Left                     =   8955
         Width                    =   2715
         Top                      =   75
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
         Index                    =   1 
         DataField                =   "Entregue"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   1695
         Width                    =   1125
         Top                      =   795
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
         Index                    =   18 
      End
      Begin VB.Label labFdo17
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2835
         Top                      =   795
         Width                    =   2025
         Height                   =   330
      End
      Begin VB.CommandButton bottxtCampo17
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   4590
         Top                      =   810
         Width                    =   255
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo17
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4335
         Top                      =   810
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
         TabIndex                 =   20
         TabStop                  =   0
         Height                   =   300
         Left                     =   2850
         Width                    =   1470
         Top                      =   810
         DataField                =   "Seqüência do Orçamento"
         Index                    =   17 
      End
      Begin VB.Label labFdo18
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8955
         Top                      =   975
         Width                    =   2610
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo18
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11295
         Top                      =   990
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
         TabStop                  =   0
         Height                   =   270
         Left                     =   8970
         Width                    =   2310
         Top                      =   990
         DataField                =   "Codigo do solicitante"
         Index                    =   18 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7455
         Width                    =   1515
         Top                      =   990
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
         Caption                  =   "Colaborador:"
         Index                    =   19 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   11520
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
         Index                    =   20 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7035
         Width                    =   1815
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
         Caption                  =   "Dt. Entrega Tecnica:"
         Index                    =   21 
      End
      Begin VB.Label labFdo19
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8955
         Top                      =   1335
         Width                    =   1650
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo19
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   10335
         Top                      =   1350
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
         TabIndex                 =   22
         TabStop                  =   -1
         Height                   =   270
         Left                     =   8970
         Width                    =   1350
         Top                      =   1350
         DataField                =   "Data Tecnica"
         Index                    =   19 
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
Attribute VB_name = "frmMotoBomb"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: MotoBomb
'* Função....: Motobomba
'* CopyRight.: (C)2024 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 16/08/2024 15:01:58
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
Dim txtCampo(19) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(1) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim Seq_Moto_Bomba As Long, Tp_de_Motor As String, Serie_da_Moto_Bomba As Long
Dim MesAno As String, Funcao_Moto_Bomba As String, Nro_de_Serie_Moto_Bomba As String
Dim Nro_de_Serie_Motor As String, Modelo_do_Motor As String, Nro_de_Serie_Bomba As String
Dim Modelo_da_Bomba As String, Data_de_Criacao As Variant, Codigo_do_Geral As Long
Dim Codigo_do_Vendedor As Long, Entregue As Boolean, NF As String
Dim Dt_da_Entrega As Variant, Obs As String, Entrega_Tecnica As String
Dim Entrega_Tec As Boolean, Sequencia_do_Orcamento As Long, Codigo_do_solicitante As Integer
Dim Data_Tecnica As Variant
Public N_Serie As Object, lblAjuste As Object

Private Function GeraSerie() As Long
Dim Tb As New GRecordSet

 Set Tb = VgDb.OpenRecordSet("Select Max([Serie da Moto Bomba])Seq From [Serie Moto Bomba]")
 
  If Tb.RecordCount > 0 Then
     GeraSerie = Tb!Seq + 1
  Else
     GeraSerie = 1
  End If
        
End Function


Private Sub GeraNroSerie()
Dim NroDeSerie As String, Mes_Ano As String, Seri As String

 Seri = Format(Serie_da_Moto_Bomba ,"0000")
 Mes_Ano = RemoveCaracteres(MesAno)
 NroDeSerie = Tp_de_Motor & Seri & "/" & Mes_Ano & Funcao_Moto_Bomba 

 N_Serie.Text = NroDeSerie

End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Seq_Moto_Bomba = txtCampo(0).Value
         Tp_de_Motor = txtCampo(1).Value
         Serie_da_Moto_Bomba = txtCampo(2).Value
         MesAno = txtCampo(3).Value
         Funcao_Moto_Bomba = txtCampo(4).Value
         Nro_de_Serie_Moto_Bomba = txtCampo(6).Value
         Nro_de_Serie_Motor = txtCampo(7).Value
         Modelo_do_Motor = txtCampo(8).Value
         Nro_de_Serie_Bomba = txtCampo(9).Value
         Modelo_da_Bomba = txtCampo(10).Value
         Data_de_Criacao = txtCampo(5).Value
         Codigo_do_Geral = txtCampo(11).Value
         Codigo_do_Vendedor = txtCampo(12).Value
         Entregue = chkCampo(1).Value
         NF = txtCampo(13).Value
         Dt_da_Entrega = txtCampo(14).Value
         Obs = txtCampo(15).Value
         Entrega_Tecnica = txtCampo(16).Value
         Entrega_Tec = chkCampo(0).Value
         Sequencia_do_Orcamento = txtCampo(17).Value
         Codigo_do_solicitante = txtCampo(18).Value
         Data_Tecnica = txtCampo(19).Value
      Else
         Seq_Moto_Bomba = IIf(IsNull(vgTb![Seq Moto Bomba]), 0, vgTb![Seq Moto Bomba])
         Tp_de_Motor = IIf(IsNull(vgTb![Tp de Motor]), "", vgTb![Tp de Motor])
         Serie_da_Moto_Bomba = IIf(IsNull(vgTb![Serie da Moto Bomba]), 0, vgTb![Serie da Moto Bomba])
         MesAno = IIf(IsNull(vgTb!MesAno), "", vgTb!MesAno)
         Funcao_Moto_Bomba = IIf(IsNull(vgTb![Função Moto Bomba]), "", vgTb![Função Moto Bomba])
         Nro_de_Serie_Moto_Bomba = IIf(IsNull(vgTb![Nro de Serie Moto Bomba]), "", vgTb![Nro de Serie Moto Bomba])
         Nro_de_Serie_Motor = IIf(IsNull(vgTb![Nro de Serie Motor]), "", vgTb![Nro de Serie Motor])
         Modelo_do_Motor = IIf(IsNull(vgTb![Modelo do Motor]), "", vgTb![Modelo do Motor])
         Nro_de_Serie_Bomba = IIf(IsNull(vgTb![Nro de Serie Bomba]), "", vgTb![Nro de Serie Bomba])
         Modelo_da_Bomba = IIf(IsNull(vgTb![Modelo da Bomba]), "", vgTb![Modelo da Bomba])
         Data_de_Criacao = vgTb![Data de Criação]
         Codigo_do_Geral = IIf(IsNull(vgTb![Codigo do Geral]), 0, vgTb![Codigo do Geral])
         Codigo_do_Vendedor = IIf(IsNull(vgTb![Codigo do Vendedor]), 0, vgTb![Codigo do Vendedor])
         Entregue = IIf(IsNull(vgTb!Entregue), 0, vgTb!Entregue)
         NF = IIf(IsNull(vgTb!NF), "", vgTb!NF)
         Dt_da_Entrega = vgTb![Dt da Entrega]
         Obs = IIf(IsNull(vgTb!Obs), "", vgTb!Obs)
         Entrega_Tecnica = IIf(IsNull(vgTb![Entrega Tecnica]), "", vgTb![Entrega Tecnica])
         Entrega_Tec = IIf(IsNull(vgTb![Entrega Tec]), 0, vgTb![Entrega Tec])
         Sequencia_do_Orcamento = IIf(IsNull(vgTb![Seqüência do Orçamento]), 0, vgTb![Seqüência do Orçamento])
         Codigo_do_solicitante = IIf(IsNull(vgTb![Codigo do solicitante]), 0, vgTb![Codigo do solicitante])
         Data_Tecnica = vgTb![Data Tecnica]
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
      i = (Entrega_Tec = 0) and Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = (Entrega_Tec = 0 Or vgPWUsuario = "JERONIMO") and Permitido(vgIdentTab, ACAO_EDITANDO)
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
      vgOk = (Not Vazio(Tp_de_Motor))
      vgMsg$ = "Tp de Motor não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = (Not Vazio(MesAno))
         vgMsg$ = "MesAno inválido!"
         If Not vgOk Then vgColumn = 4
      End If
      If vgOk Then
         vgOk = (Not Vazio(Funcao_Moto_Bomba))
         vgMsg$ = "Função Moto Bomba não pode ser vazio!"
         If Not vgOk Then vgColumn = 5
      End If
      If vgOk Then
         vgOk = ((Not Vazio(Data_de_Criacao)) And (IsDate(Data_de_Criacao) Or Vazio(Data_de_Criacao)))
         vgMsg$ = "Data de Criação não pode ser vazio!"
         If Not vgOk Then vgColumn = 6
      End If
      If vgOk Then
         vgOk = (IsDate(Dt_da_Entrega) Or Vazio(Dt_da_Entrega))
         vgMsg$ = "Dt da Entrega não pode ser vazio!"
         If Not vgOk Then vgColumn = 15
      End If
      If vgOk Then
         vgOk = (Sequencia_do_Orcamento >= 0)
         vgMsg$ = "Seqüência do Orçamento inválido!"
         If Not vgOk Then vgColumn = 18
      End If
      If vgOk Then
         vgOk = (Codigo_do_solicitante >= 0)
         vgMsg$ = "Informe o Colaborador!"
         If Not vgOk Then vgColumn = 19
      End If
      If vgOk Then
         vgOk = (IsDate(Data_Tecnica) Or Vazio(Data_Tecnica))
         vgMsg$ = "Data Tecnica não pode ser vazio!"
         If Not vgOk Then vgColumn = 20
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
                  txtCampo(2).Value = GeraSerie()
         chkCampo(0).Value = False
         chkCampo(1).Value = False
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 11
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 17
         PoeRelEFiltroCbo 18
      End If
      If vgSituacao = ACAO_INCLUINDO Then
         txtCampo(0).Value = PegaSequencia(vgIdentTab,"Seq Moto Bomba", 0, 1)
         InicializaApelidos COM_TEXTBOX
         vgPegouSeq = True
      End If
   ElseIf vgOq = VOLTA_SEQUENCIAL Then
      If vgPegouSeq  And vgSituacao = ACAO_INCLUINDO Then
         VoltaSequencia vgIdentTab,"Seq Moto Bomba"
      End If
      vgPegouSeq = False
   ElseIf vgOq = CONFIRMA_SEQUENCIAL Then
      If vgPegouSeq  And vgSituacao = ACAO_INCLUINDO Then
         ConfirmaSequencia vgIdentTab,"Seq Moto Bomba"
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 11
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 17
         PoeRelEFiltroCbo 18
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Entrega_Tec
         chkCampo(1).Value = Entregue
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

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then vgPegouSeq = False
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
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
   Ok = (((((Not Vazio(Tp_de_Motor)) AND Not Vazio(Serie_da_Moto_Bomba)) AND Not Vazio(MesAno)) AND Not Vazio(Funcao_Moto_Bomba)) AND Entregue = 0)
   Botao(0).Enabled = Ok
   Ok = (Entregue)
   chkCampo(0).Locked = Not (vgAlterar And chkCampo(0).Editable)
   chkCampo(0).Enabled = Ok Or Not vgAlterar
   Ok = (Entrega_Tecnica = "Irrigação Penapolis")
   Label(19).Enabled = Ok And vgAlterar
   Ok = (Entrega_Tecnica = "Irrigação Penapolis")
   Label(21).Enabled = Ok And vgAlterar
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 1 Or Index = 4 Or Index = 11 Or Index = 12 Or Index = 16 Or Index = 17 Or Index = 18 Then
      ExecutaPreValidacao
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 5
            If Len(txtCp(5).Text) = 0 Then
               txtCampo(5).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 11
            PoeRelEFiltroCbo 11
         Case 12
            PoeRelEFiltroCbo 12
         Case 14
            If Len(txtCp(14).Text) = 0 Then
               txtCampo(14).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 17
            PoeRelEFiltroCbo 17
         Case 18
            PoeRelEFiltroCbo 18
         Case 19
            If Len(txtCp(19).Text) = 0 Then
               txtCampo(19).Value = Date
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
         chkCampo(0).Value = Entrega_Tec
      ElseIf Index = 1 Then
         chkCampo(1).Value = Entregue
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
   Caption = LoadGasString(77160)
   vgFormID = 1193
   vgIdentTab = "Serie Moto Bomba"
   vgFiltroEmUso = -1
   vgIndDefault = "Seq Moto Bomba"
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
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(68665))
   Set N_Serie = txtCampo(6)
   Set lblAjuste = Label(20)
   Set vgTb = New GRecordSet
   If Len(vgFiltroOriginal$) = 0 And Len(vgFiltroInicial$) > 0 Then
      vgFiltroOriginal$ = vgFiltroInicial$
   End If
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(77165)
   Label(1).Caption = LoadGasString(77166)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(77167)
   Label(2).Caption = LoadGasString(77168)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(77169)
   Label(3).Caption = LoadGasString(77170)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(77171)
   Label(4).Caption = LoadGasString(77172)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(77173)
   Label(5).Caption = LoadGasString(77174)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(77175)
   Label(6).Caption = LoadGasString(77176)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(77177)
   Label(7).Caption = LoadGasString(77178)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(77179)
   Botao(0).Caption = LoadGasString(77180)
   Label(8).Caption = LoadGasString(77181)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(77182)
   Label(9).Caption = LoadGasString(77183)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(77184)
   Label(10).Caption = LoadGasString(77185)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(77186)
   Label(11).Caption = LoadGasString(77187)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(77188)
   Label(12).Caption = LoadGasString(77189)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(77190)
   Label(13).Caption = LoadGasString(77191)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(77192)
   Label(14).Caption = LoadGasString(77193)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(77194)
   Label(15).Caption = LoadGasString(77195)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(77196)
   Label(16).Caption = LoadGasString(77197)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(77198)
   Label(17).Caption = LoadGasString(77199)
   vgTooltips.AddTool txtCampo(16).CtPri, 0, LoadGasString(77200)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(77201)
   chkCampo(0).Caption = LoadGasString(77202)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(77203)
   chkCampo(1).Caption = LoadGasString(77204)
   Label(18).Caption = LoadGasString(77205)
   vgTooltips.AddTool txtCampo(17).CtPri, 0, LoadGasString(77206)
   vgTooltips.AddTool txtCampo(18).CtPri, 0, LoadGasString(77207)
   Label(19).Caption = LoadGasString(77208)
   Label(20).Caption = LoadGasString(77209)
   Label(21).Caption = LoadGasString(77210)
   vgTooltips.AddTool txtCampo(19).CtPri, 0, LoadGasString(77211)
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

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtBot(BOT_COMBO) = bottxtCampo1(BOT_COMBO)
   Set bottxtCampo1(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(1).DataType = 0
   txtCampo(1).ListFields = "D|E"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "99999"
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "99/99"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 0
   txtCampo(4).ListFields = "A|V|A-G"
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
   txtCampo(6).DataType = 0
   txtCampo(6).Mask = "@!"
   txtCampo(6).Editable = False
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 0
   txtCampo(7).Mask = "@!"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "@!"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 0
   txtCampo(9).Mask = "@!"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 0
   txtCampo(10).Mask = "@!"
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   Set txtCampo(11).CtFdo = labFdo11
   Set txtCampo(11).CtBot(BOT_LISTA) = bottxtCampo11(BOT_LISTA)
   Set txtCampo(11).CtBot(BOT_COMBO) = bottxtCampo11(BOT_COMBO)
   bottxtCampo11(BOT_LISTA).Caption = "P"
   Set bottxtCampo11(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "999999"
   txtCampo(11).PesqModoAbertura =  2 
   txtCampo(11).PesqFieldCapture = "Geral.[Seqüência do Geral]"
   txtCampo(11).BoundColumn = "Seqüência do Geral"
   txtCampo(11).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(11).OrderFields = "Razão Social"
   txtCampo(11).Relation = ""
   txtCampo(11).Source = "Geral"
   txtCampo(11).vgfrmGMCale.grdListaG.AutoRebind = True
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
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "@x"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(14).CtPri = txtCp(14)
   Set txtCampo(14).CtFdo = labFdo14
   Set txtCampo(14).CtBot(BOT_ACAO) = bottxtCampo14(BOT_ACAO)
   Set bottxtCampo14(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(14).DataType = 2
   txtCampo(14).Mask = "99/99/9999"
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 4
   txtCampo(15).Mask = ""
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(15).DataField), txtCampo(15)

   Set txtCampo(16).CtPri = txtCp(16)
   Set txtCampo(16).CtFdo = labFdo16
   Set txtCampo(16).CtBot(BOT_COMBO) = bottxtCampo16(BOT_COMBO)
   Set bottxtCampo16(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(16).DataType = 0
   txtCampo(16).ListFields = "Irrigação Penapolis|Revenda"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(1).DataField), chkCampo(1)

   Set txtCampo(17).CtPri = txtCp(17)
   Set txtCampo(17).CtFdo = labFdo17
   Set txtCampo(17).CtBot(BOT_LISTA) = bottxtCampo17(BOT_LISTA)
   Set txtCampo(17).CtBot(BOT_COMBO) = bottxtCampo17(BOT_COMBO)
   bottxtCampo17(BOT_LISTA).Caption = "P"
   Set bottxtCampo17(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(17).DataType = 1
   txtCampo(17).Mask = "999999"
   txtCampo(17).PesqModoAbertura =  2 
   txtCampo(17).PesqFieldCapture = "Orçamento.[Seqüência do Orçamento]"
   txtCampo(17).BoundColumn = "Seqüência do Orçamento"
   txtCampo(17).ListFields = "Seqüência do Orçamento"
   txtCampo(17).OrderFields = "Seqüência do Orçamento-"
   txtCampo(17).Relation = ""
   txtCampo(17).Source = "Orçamento"
   txtCampo(17).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(17).DataField), txtCampo(17)

   Set txtCampo(18).CtPri = txtCp(18)
   Set txtCampo(18).CtFdo = labFdo18
   Set txtCampo(18).CtBot(BOT_COMBO) = bottxtCampo18(BOT_COMBO)
   Set bottxtCampo18(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(18).DataType = 1
   txtCampo(18).Mask = "999"
   txtCampo(18).BoundColumn = "Codigo do solicitante"
   txtCampo(18).ListFields = "Nome do solicitante"
   txtCampo(18).OrderFields = "Nome do solicitante"
   txtCampo(18).Relation = ""
   txtCampo(18).Source = "Solicitantes"
   txtCampo(18).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(18).DataField), txtCampo(18)

   Set txtCampo(19).CtPri = txtCp(19)
   Set txtCampo(19).CtFdo = labFdo19
   Set txtCampo(19).CtBot(BOT_ACAO) = bottxtCampo19(BOT_ACAO)
   Set bottxtCampo19(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(19).DataType = 2
   txtCampo(19).Mask = "99/99/9999"
   txtCampo(19).BoundColumn = ""
   txtCampo(19).ListFields = ""
   txtCampo(19).OrderFields = ""
   txtCampo(19).Relation = ""
   txtCampo(19).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(19).DataField), txtCampo(19)

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
      Case 11
         txtCampo(11).Filter = "((Cliente = " & 1 & ") AND [Seqüência do Geral] > " & 0 & ") AND Inativo = " & 0 & ""
                           txtCampo(11).PesqSQLExpression = "SELECT Geral.[Seqüência do Geral], Geral.[Razão Social] FROM Geral WHERE (Geral.Cliente = " + CStr(1) + ") AND " + _
                                                               "(Geral.[Seqüência do Geral] > " + CStr(0) + ") AND (Geral.Inativo = " + CStr(0) + ")"
      Case 12
         txtCampo(12).Filter = "(Vendedor = " & 1 & ") AND Inativo = " & 0 & ""
      Case 17
         txtCampo(17).Filter = "[Seqüência do Orçamento] > " & 0 & " AND Cancelado = " & 0 & ""
                           txtCampo(17).PesqSQLExpression = "SELECT Orçamento.[Seqüência do Orçamento], Orçamento.[Data de Emissão], Orçamento.[Vendedor Intermediario], " + _
                                                               "Geral.[Razão Social] FROM Orçamento, Geral WHERE (Orçamento.[Seqüência do Orçamento] > " + CStr(0) + ") AND " + _
                                                               "(Orçamento.[Seqüência do Geral] = Geral.[Seqüência do Geral]) AND (Orçamento.Cancelado = " + CStr(0) + ")"
      Case 18
         txtCampo(18).Filter = "([Codigo do solicitante] > " & 0 & ") AND [Codigo do setor] = " & 13 & ""
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
   Set N_Serie = Nothing
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   grdBrowse.Finalize                             'descarrega o grid
   Set frmMotoBomb = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         GeraNroSerie  
   End Select
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
Private Sub bottxtCampo14_Click(Index As Integer)
   txtCampo(14).SetFocus
   DoEvents
   txtCampo(14).botClick Index
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

'evento - quando o botão for apertado
Private Sub bottxtCampo18_Click(Index As Integer)
   txtCampo(18).SetFocus
   DoEvents
   txtCampo(18).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo19_Click(Index As Integer)
   txtCampo(19).SetFocus
   DoEvents
   txtCampo(19).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


