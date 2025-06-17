VERSION 5.00
Begin VB.Form frmRelReceb
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   5370
   Left                     =   2610
   LinkTopic                =   "RelReceb"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   4935
   ScaleWidth               =   12060
   Top                      =   1885
   Width                    =   12150
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
      Height                   =   4755
      Left                     =   15
      Width                    =   11880
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   20944
      _extenty                 =   8383
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   3705
         Width                    =   735
         Top                      =   4080
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
         ForeColor                =   &H800000
         Caption                  =   "Destino:"
         Index                    =   0 
      End
      Begin NFE.GPainel Painel
         BackColor                =   &H8000000F
         BevelOuter               =   0
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   240
         Left                     =   4515
         Width                    =   2280
         Top                      =   4080
         TabStop                  =  0
         _extentx                 =   4019
         _extenty                 =   423
         Index                    =   1 
         Begin             VB.Label labopcPainel1
            BorderStyle              =   1
            Height                   =   120
            Left                     =   690
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcPainel1Cp
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
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   12
            TabStop                  =   0
            Height                   =   240
            Left                     =   0
            Width                    =   930
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Janela"
            Index                    =   0 
         End
         Begin VB.OptionButton opcPainel1Cp
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
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   13
            TabStop                  =   0
            Height                   =   240
            Left                     =   990
            Width                    =   1320
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Impressora"
            Index                    =   1 
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   195
         Width                    =   810
         Top                      =   4050
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
         ForeColor                =   &H800000
         Caption                  =   "Intervalo:"
         Index                    =   1 
      End
      Begin NFE.GPainel Painel
         BackColor                =   &H8000000F
         BevelOuter               =   0
         BevelInner               =   0
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   240
         Left                     =   1080
         Width                    =   2010
         Top                      =   4050
         TabStop                  =  0
         _extentx                 =   3543
         _extenty                 =   423
         Index                    =   2 
         Begin             VB.Label labopcPainel2
            BorderStyle              =   1
            Height                   =   120
            Left                     =   665
            Top                      =   180
            Visible                  =   0
            Width                    =   240
            End
         Begin VB.OptionButton opcPainel2Cp
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
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   14
            TabStop                  =   0
            Height                   =   240
            Left                     =   0
            Width                    =   905
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Todas"
            Index                    =   0 
         End
         Begin VB.OptionButton opcPainel2Cp
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
            ForeColor                =   &H800000
            Alignment                =   2
            TabIndex                 =   15
            TabStop                  =   0
            Height                   =   240
            Left                     =   975
            Width                    =   1050
            Top                      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H800000
            Alignment                =   0
            Caption                  =   "Páginas"
            Index                    =   1 
         End
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
         TabIndex                 =   16
         TabStop                  =   0
         Height                   =   315
         Left                     =   3165
         Width                    =   465
         Top                      =   4050
         Index                    =   0 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7905
         Top                      =   4050
         Width                    =   555
         Height                   =   315
      End
      Begin VB.VScrollBar scrVtxtCampo1
         Max                      =   0
         Min                      =   32767
         LargeChange              =   10
         SmallChange              =   1
         Value                    =   0
         TabStop                  =   0
         Left                     =   8250
         Top                      =   4065
         Width                    =   195
         Height                   =   285
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   11
         TabStop                  =   0
         Height                   =   285
         Left                     =   7920
         Width                    =   315
         Top                      =   4065
         Index                    =   1 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6870
         Width                    =   960
         Top                      =   4080
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
         Caption                  =   "Nº Copias:"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8550
         Width                    =   705
         Top                      =   4080
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
         Caption                  =   "Impres.:"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   300
         Left                     =   9330
         Width                    =   1920
         Top                      =   4050
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
         ForeColor                =   &HFF0000
         Caption                  =   ""
         Index                    =   4 
      End
      Begin NFE.GBotao botConfImp
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   17
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   360
         Left                     =   11280
         Width                    =   360
         Top                      =   4005
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
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   150
         Y1                       =   3870
         X2                       =   11625
         Y2                       =   3870
         Index                    =   0 
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
         Height                   =   1395
         Left                     =   75
         Width                    =   1455
         Top                      =   75
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   465
         Left                     =   1800
         Width                    =   7005
         Top                      =   600
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   20.25
            Underline                =   -1
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Caption                  =   "Contas a Receber - Por Vencimento"
         Index                    =   5 
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
         ForeColor                =   &H80
         Alignment                =   0
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   255
         Left                     =   10710
         Width                    =   900
         Top                      =   375
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80
         Alignment                =   0
         Caption                  =   "Zebrar"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   390
         Width                    =   840
         Top                      =   2385
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
         Caption                  =   "Dt. Inicial:"
         Index                    =   6 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1305
         Top                      =   2355
         Width                    =   1395
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   2430
         Top                      =   2370
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
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1320
         Width                    =   1095
         Top                      =   2370
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2775
         Width                    =   765
         Top                      =   2385
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
         Caption                  =   "Dt. Final:"
         Index                    =   7 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3660
         Top                      =   2355
         Width                    =   1500
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4890
         Top                      =   2370
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
         MaxLength                =   254
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   270
         Left                     =   3675
         Width                    =   1200
         Top                      =   2370
         Index                    =   3 
      End
      Begin NFE.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   9
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   360
         Left                     =   8640
         Width                    =   1485
         Top                      =   3360
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
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
         TabIndex                 =   10
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   360
         Left                     =   10140
         Width                    =   1485
         Top                      =   3360
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Cancela"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   11595
         Width                    =   45
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
         Caption                  =   "."
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
         ForeColor                =   &H80000008
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1305
         Width                    =   3855
         Top                      =   2880
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   435
         Width                    =   765
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
         Caption                  =   "% Geral:"
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5250
         Width                    =   645
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
         Caption                  =   "% Doc:"
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
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6075
         Width                    =   2370
         Top                      =   2880
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5355
         Width                    =   540
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
         Caption                  =   "Geral:"
         Index                    =   11 
      End
      Begin VB.Label labFdo6
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6075
         Top                      =   2340
         Width                    =   5550
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   11355
         Top                      =   2355
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11100
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6090
         Width                    =   4995
         Top                      =   2355
         Index                    =   6 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   405
         Width                    =   795
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
         Caption                  =   "F. Pagto:"
         Index                    =   12 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1320
         Top                      =   3420
         Width                    =   1305
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   2355
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   6
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1335
         Width                    =   1005
         Top                      =   3435
         Index                    =   7 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5460
         Width                    =   435
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
         Caption                  =   "NFe:"
         Index                    =   13 
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
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   270
         Left                     =   6075
         Width                    =   840
         Top                      =   3420
         Index                    =   8 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8595
         Width                    =   945
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
         Caption                  =   "Vendedor:"
         Index                    =   14 
      End
      Begin VB.Label labFdo9
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   9675
         Top                      =   2880
         Width                    =   1950
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo9
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11355
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   270
         Left                     =   9690
         Width                    =   1650
         Top                      =   2895
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2760
         Width                    =   480
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
         Caption                  =   "Cobr:"
         Index                    =   15 
      End
      Begin VB.Label labFdo10
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3360
         Top                      =   3420
         Width                    =   1785
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo10
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   4875
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   270
         Left                     =   3375
         Width                    =   1485
         Top                      =   3435
         Index                    =   10 
      End
      Begin          VB.Label labopcPainel0
         BorderStyle              =   1
         Height                   =   120
         Left                     =   4395
         Top                      =   180
         Visible                  =   0
         Width                    =   240
         End
      Begin VB.OptionButton opcPainel0Cp
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
         TabIndex                 =   19
         TabStop                  =   0
         Height                   =   300
         Left                     =   3315
         Width                    =   1320
         Top                      =   300
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
         Caption                  =   "Emissão"
         Index                    =   0 
      End
      Begin VB.OptionButton opcPainel0Cp
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
         TabIndex                 =   20
         TabStop                  =   0
         Height                   =   255
         Left                     =   4755
         Width                    =   1665
         Top                      =   300
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
         Caption                  =   "Vencimento"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1800
         Width                    =   1320
         Top                      =   300
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
         Caption                  =   "Ordenar Por:"
         Index                    =   16 
      End
      Begin VB.OptionButton opcPainel0Cp
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
         TabIndex                 =   21
         TabStop                  =   0
         Height                   =   255
         Left                     =   6540
         Width                    =   2175
         Top                      =   300
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
         Caption                  =   "Nome do Cliente"
         Index                    =   2 
      End
      Begin VB.OptionButton opcPainel0Cp
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
         TabIndex                 =   22
         TabStop                  =   0
         Height                   =   255
         Left                     =   8820
         Width                    =   2055
         Top                      =   300
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
         Caption                  =   "Documento"
         Index                    =   3 
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
         TabIndex                 =   23
         TabStop                  =   -1
         Height                   =   255
         Left                     =   7020
         Width                    =   1290
         Top                      =   3480
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
         Caption                  =   "Avulsas"
         Index                    =   1 
      End
      Begin VB.Frame Frame
         Height                   =   840
         Left                     =   1815
         Width                    =   4215
         Top                      =   1215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Verdana"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   14.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Empresa"
         Index                    =   0 
         Begin             VB.Label labopcFrame0
            BorderStyle              =   1
            Height                   =   120
            Left                     =   1695
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
               Size                     =   15.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   2
            TabIndex                 =   24
            TabStop                  =   -1
            Height                   =   375
            Left                     =   120
            Width                    =   1815
            Top                      =   420
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   15.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   0
            Caption                  =   "Reticom"
            Index                    =   0 
         End
         Begin VB.OptionButton opcFrame0Cp
            Appearance               =   1
            BackColor                =   &H8000000F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   15.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   2
            TabIndex                 =   25
            TabStop                  =   -1
            Height                   =   375
            Left                     =   2040
            Width                    =   2115
            Top                      =   420
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   15.75
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            Alignment                =   0
            Caption                  =   "Reticom 2"
            Index                    =   1 
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
Attribute VB_name = "frmRelReceb"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: RelReceb
'* Função....: Contas a Receber
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:54
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
Public vgUltimoTabIndex As Integer                'último campo com foco do módulo 

Public vgUltimoFiltro As String                   'último filtro definido no módulo
Public vgUltimaOrdem As String                    'última ordenação feita no módulo
Public vgUltimoFiltroComTit As String             'titulo do último filtro definido no módulo
Public vgUltimaOrdemComTit As String              'titulo da última ordenação feita no módulo 
Public vgSQL As String                            'expressão SQL que define o módulo
Public vgFP As New frmPreview                     'formuário de preview do relatório 
Public vgVisivel As Boolean                       'se a tela de parâmetros do relatório esta visível 
Public vgPosLin As Single                         'posição atual de impressão 
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFG As frmGauge                           'formulário do gauge 
Public vgQReg As Long                             'quantidade de já feira para o gauge  

Dim vgBookPg() As Variant                         'marcador do registro do inicio da página 
Dim vgPodeVoltar() As Boolean                     'se pode ou não retroceder para esta página   
Dim vgDataInicial As Variant                      'data que iniciou o relatório  
Dim vgPrimeiroCab As Integer                      'flag se ja foi impresso o primeiro cabeçlho do relatório
Dim vgFiltroInicial As String                     'filtro inicial do relatório
Dim vgOrdemInicial As String                      'ordem incial do relatório 
Dim txtCampo(10) As New FormataCampos             'classe dos compos tipo texto do módulo  
Dim chkCampo(1) As New FormataCampos              'classe dos compos tipo lógico do módulo 
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim vgWidthRel As Single                          'lagura do relatório 
Dim vgHeightRel As Single                         'altura do relatório
Dim vgPagRef() As Integer                         'Ordem da página dentro do intervalo 
Dim opcFrame0(1) As New FormataCampos
Dim opcPainel0(3) As New FormataCampos
Dim opcPainel1(1) As New FormataCampos
Dim opcPainel2(1) As New FormataCampos
Dim vgDestRel As Integer, vgModoInt As Integer, vgIntPag As String
Dim vgCopias As Integer, Zebrar As Boolean, Dt_Inicial As Variant
Dim Dt_Final As Variant, Nome As String, Documento As String
Dim Geral As Double, F_Pagto As String, NFe As Double
Dim Vendedor As Double, Cobranca As Double, TipoFiltro As Integer
Dim Avulsa As Boolean, Empresa As Integer, vgPgvgTot1() As Double
Dim vgPgvgTot2() As Double, vgPgvgTot3() As Double, vgPgvgTot4() As Double
Dim vgPgvgTot5() As Double, vgPgvgTot6() As Double, vgPgvgTot7() As Double
Dim vgPgvgTot8() As Double, vgPgvgTot9() As Double, vgPgvgTot10() As Double
Dim vgPgvgTot15() As Double, vgPgvgTot13() As Double, vgPgvgTot14() As Double
Dim vgPgvgTot12() As Double, vgPgvgTot11() As Double, vgPgvgTot131() As Double
Dim vgPgvgTot1311() As Double, vgPgvgTot13111() As Double
Public lblAjuste As Object
Dim Manutencao_contas_receber As New GRecordSet

Dim Contagem As Long, vaPrivez As Boolean, vaPg As Integer

Private Function Filtro() As String
   Filtro = "1 = 1"
   If Nome <> "" Then
      Filtro = Filtro & " AND Geral.[Razão social] Like '%" & Nome & "%'"
   End If
   If Not Vazio(Dt_Inicial) Then
    If TipoFiltro = 1 Then
      Filtro = Filtro & " AND [Data de vencimento] >= " & D(Dt_Inicial)
     Else
      Filtro = Filtro & " AND [Data de Entrada] >= " & D(Dt_Inicial)
     End If 
   End If
   If Not Vazio(Dt_Final) Then
     If TipoFiltro = 1 Then
      Filtro = Filtro & " AND [Data de vencimento] <= " & D(Dt_Final)
     Else
      Filtro = Filtro & " AND [Data de Entrada] <= " & D(Dt_Final)
     End If 
   End If
   If Documento <> "" Then
      Filtro = Filtro & " AND [Manutenção contas receber].Documento Like '%" & Documento & "%'"
   End If
   If Not Vazio(F_Pagto) Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Especie de pagamento] = '" & F_Pagto & "'"
   End If
   If Geral > 0 Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Sequencia do geral] = " & Geral
   End If
   
   If NFe > 0 Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Numero da nfe] = " & NFe
   End If
   
   If Cobranca > 0 Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Sequencia da Cobrança] = " & Cobranca
   End If
   
   If Vendedor > 0 Then
      Filtro = Filtro & "  AND [Sequencia da Ordem] IN (Select [Sequencia da Ordem] FROM [Ordem de Serviço] OS INNER JOIN Geral G ON G.[Sequencia Do Geral] = OS.[Sequencia Do Vendedor] WHERE G.[Sequencia Do Geral] = " & Vendedor & ")"
   End If
   
   If Avulsa Then
      Filtro = Filtro & " AND [Manutenção contas receber].[Sequencia da Entrada] > 0" 
   End If
   
   Filtro = Filtro & " AND [Valor Restante] > 0"
   Filtro = Filtro & " AND [Sequencia da Ordem] IN (SELECT [Sequencia da Ordem] FROM [Ordem de Serviço] WHERE [Qual Empresa] = " & Empresa & ")"
   
   If TipoFiltro = 1 Then' Vencimento
      Filtro = Filtro & " ORDER BY [Data de vencimento], [Sequencia da manutenção]"
   ElseIf TipoFiltro = 0 Then ' Emissão 
      Filtro = Filtro & " ORDER BY [Data de Entrada], [Sequencia da manutenção]"
   ElseIf TipoFiltro = 2 Then ' Cliente 
      Filtro = Filtro & " ORDER BY [Razão social], [Sequencia da manutenção]"
   ElseIf TipoFiltro = 3 Then ' Documento
      Filtro = Filtro & " ORDER BY Documento, [Sequencia da manutenção]"    
   End If 
   
     
End Function


'Função para controlar a zebra
Private Function Mudou() As Boolean
   If Not Zebrar Then Exit Function
   If vaPrivez Then Contagem = 0: vaPg = 1
   If vgFP.vgPagRef <> vaPg Then Contagem = 0
   vaPrivez = False
   vaPg = vgFP.vgPagRef
   Contagem = Contagem + 1
   If Contagem Mod 2 > 0 Then Mudou = True
End Function


Private Function MostraVendedor() As String
Dim Tb As New GRecordSet

 Set Tb = vgDb.OpenRecordSet("SELECT [Razão Social] From Geral Where [Sequencia do Geral] = " & Vendedor)
 MostraVendedor = "Vendedor --> " & Tb![Razão Social]
 
End Function

Private Function MostraOrdem() As String
 If TipoFiltro = 0 Then
    MostraOrdem = "Ordenação - Data de Emissão"
 ElseIf TipoFiltro = 1 Then
    MostraOrdem = "Ordenação - Data de Vencimento"
 ElseIf TipoFiltro = 2 Then
    MostraOrdem = "Ordenação - Nome do Cliente" 
 ElseIf TipoFiltro = 3 Then
    MostraOrdem = "Ordenação - Documento"          
 End If     
End Function


Private Function DefineTipo() As String

 If TipoFiltro = 0 Then
    DefineTipo = "Emissão - " & Manutencao_contas_receber![Data de entrada]
 ElseIf TipoFiltro = 1 Then
    DefineTipo = "Vencimento: - " & Manutencao_contas_receber![Data de Vencimento] 
 ElseIf TipoFiltro = 2 Then
    DefineTipo = Manutencao_contas_receber![Razão Social] 
 Else
    DefineTipo = Manutencao_contas_receber!Documento     
 End If
    
End Function


'inicia impressão do relatório/etiqueta
Public Sub ProssegueExecucao(vgObSaida As Object, Optional ByVal vgNumPg As IMPRIME_PAGINA = 0)
   Dim vgPsm As Integer, vgPsz As Integer, vgPfn As String, vgPor As Integer, x As String, vgSemBookMark As Integer
   Dim i As Integer, vgLimiteInf As Single, vgSqlFiltro As String, vgTemFimRel As Boolean
   Dim vgNParc As Integer, vgNVez As Integer, vgNCop As Integer
   Dim vgEmPapel As Boolean, vgPagOk As Boolean, vgPV As Boolean, vgDifProxReg As Single, vgNumPag As Long, vgSemControlePg As Boolean, Ok As Boolean
   Dim j As Integer, m As Integer, vgQt As Integer, vgPagOrig As Integer, vgNovaPg As Integer, k As Integer, vgIniProc As Boolean, vgTestaPriCab As Boolean
   Dim vgQuebra20 As Variant, vgPriQb20 As Integer, vgTot1 As Double, vgTot2 As Double, vgTot3 As Double, vgTot4 As Double
   Dim vgTot5 As Double, vgTot6 As Double, vgTot7 As Double, vgTot8 As Double, vgTot9 As Double, vgTot10 As Double
   Dim vgTot15 As Double, vgTot13 As Double, vgTot14 As Double, vgTot12 As Double, vgTot11 As Double, vgTot131 As Double
   Dim vgTot1311 As Double, vgTot13111 As Double
   vgTot6 = 0
   vgTot7 = 0
   vgTot8 = 0
   vgTot9 = 0
   vgTot10 = 0
   vgTot15 = 0
   vgTot13 = 0
   vgTot14 = 0
   vgTot12 = 0
   vgTot11 = 0
   vgTot131 = 0
   vgTot1311 = 0
   vgTot13111 = 0
   Screen.MousePointer = vbHourglass              'mostra ampulheta
   vgCancelou = False
   vgTemFimRel = False                            'se tem pendência para ser impressa como fim de página
   vgSemBookMark = 0
   If TypeOf vgObSaida Is Printer Then            'estamos em preview ou impressora?
      vgEmPapel = True                            'imprimindo na impressora, confirma
      GeraLogAcao LoadGasString(296) + " Contas a Receber"
   Else
      vgEmPapel = False
      vgObSaida.Cls
      vgFP.scrH.Value = 0                         'acerta scroll do preview
      vgFP.scrV.Value = 0
   End If
   vgDataInicial = Now                            'data e hora do inicio do relatório
   vgNovaPg = -1
   vgIniProc = False                              'flag para informar que iniciou procura em navegação (próxima, anterior, última...)

   If vgNumPg = G_IMP_ATUA_PG Then                'solicitou impressão da página atual
      vgNumPg = vgFP.vgContaPg
   End If

   vgNCop = 0                                     'controla num de cop se impressora não tem esta prop
   vgSemControlePg = (False)

OutraCopia:
   On Error GoTo DeuErro                          'vamos cercar erros...

   'abre os recordsets
   vgPV = (vgNumPg = 0 Or (vgEmPapel And vgFP.vgContaPg > 0)) 'veio do botao de prosseguir?
   If vgPV And Not (vgEmPapel And vgFP.vgContaPg > 0) Then
      If TypeOf vgObSaida Is PictureBox Then                  'estamos em preview ou impressora?
         GeraLogAcao LoadGasString(297) + " Contas a Receber"
      End If
      vgFiltroInicial = "" & Filtro() & ""
      x$ = vgQueryAtual
      If Len(vgUltimoFiltro) > 0 And Len(vgFiltroInicial) > 0 Then
         x$ = InsereSQL(x$, EXP_WHERE, "(" + vgFiltroInicial + ") And (" + vgUltimoFiltro + ")")
      ElseIf Len(vgUltimoFiltro + vgFiltroInicial) > 0 Then
         x$ = InsereSQL(x$, EXP_WHERE, vgFiltroInicial + vgUltimoFiltro)
      End If
      vgOrdemInicial = ExtraiSQL(x$, EXP_ORDERBY)
      If Len(vgUltimaOrdem) > 0 Or Len(vgOrdemInicial) > 0 Then
         x$ = InsereSQL(x$, EXP_ORDERBY, vgOrdemInicial + IIf(Len(vgOrdemInicial) > 0 And Len(vgUltimaOrdem) > 0, ", ", "") + vgUltimaOrdem)
      End If
      Set Manutencao_contas_receber = vgDb.OpenRecordSet(x$)
      If Manutencao_contas_receber.RecordCount = 0 Then
       vgPrimeiroCab = 1
       Screen.MousePointer = vbNormal
       Exit Sub
      End If
      ReDim vgBookPg(0) As Variant                            'cria variável para salvar bookmark
      ReDim vgPodeVoltar(0) As Boolean                        'controle de quebra em filhos
      GoSub IniciaImpressao
   Else
      If vgNumPg >= 0 And (vgEmPapel And vgFP.vgContaPg > 0) Then
         GoSub IniciaImpressao
      End If
      If vgModoInt = 1 Then
         vgPagOrig = vgPagRef(vgFP.vgPagRef)
      Else
         vgPagOrig = vgFP.vgPagRef
      End If
      If vgNumPg > 0 Then                                     'se estiver indo para alguma página específica
         If vgModoInt = 1 Then
            j = 0
            For i = 1 To UBound(vgPagRef)                     'vamos procurar a página solicitada no intervalo de páginas permitidas para o relatório...
               If vgPagRef(i) = vgNumPg Then
                  j = i
                  If i >= vgPagOrig Then                      'se a página encontrada estiver após a página atualamente exibida no relatório...
                     Exit For                                 'cai fora
                  End If
               End If
            Next
            If j = 0 Then
               vgFP.vgContaPg = vgFP.vgContaPg - 1            'para evitar problema na SalvaPagina
               GoTo NaoAchouPg
            End If
            vgFP.vgPagRef = j
         End If
         vgFP.vgNovaPag = vgNumPg
         If vgNumPg > UBound(vgBookPg()) Then
            vgNumPg = UBound(vgBookPg())
         End If
         vgNovaPg = vgNumPg - 1
         vgFP.vgProcuraPg = (vgNumPg <> vgFP.vgNovaPag)
      ElseIf vgNumPg = G_IMP_PROX_PG Or vgNumPg = G_IMP_ANT_PG Then
         i = (vgNumPg = G_IMP_ANT_PG) ^ (vgNumPg = G_IMP_ANT_PG) 'i=+1 para G_IMP_PROX_PG e i=-1 para G_IMP_ANT_PG
         If vgModoInt = 1 Then
            vgFP.vgProcuraPg = (vgPagRef(vgFP.vgPagRef + (i * 1)) > UBound(vgBookPg()))
            If vgFP.vgProcuraPg Then
               vgIniProc = True                                  'flag para informar que iniciamos procura
               vgFP.vgNovaPag = vgPagRef(vgFP.vgPagRef + (i * 1))
               vgNovaPg = UBound(vgBookPg) - 1
               vgFP.vgPagRef = vgFP.vgPagRef + (i * 1)
            Else
               vgNovaPg = vgPagRef(vgFP.vgPagRef + (i * 1)) - 1
            End If
         Else
            vgFP.vgProcuraPg = False
            vgNovaPg = vgFP.vgContaPg + (i * 1) - 1
         End If
      ElseIf vgNumPg = G_IMP_PRI_PG Or vgNumPg = 0 Then
         If vgEmPapel And vgFP.vgContaPg > 0 Then                'se for em papel... e tiver vindo do botão IMPRIME do preview...
            vgPagOrig = vgFP.vgPagRef
         End If
         vgFP.vgProcuraPg = False
         If vgModoInt = 1 Then
            vgNovaPg = vgPagRef(1) - 1
         Else
            vgNovaPg = 0
         End If
      ElseIf vgNumPg = G_IMP_ULT_PG Then
         If vgModoInt = 1 Then
            vgFP.vgProcuraPg = False
            If vgFP.vgUltPagVista <= vgPagRef(UBound(vgPagRef())) Then 'se ainda não visualizou a última página...
               vgNovaPg = UBound(vgBookPg()) - 1                       'vamos posicionar em uma página antes da última disponível (vai incrementar depois)
            Else
               vgNovaPg = vgPagRef(UBound(vgPagRef())) - 1             'posicionar uma página antes da última página do relatório
            End If
         Else
            If UBound(vgBookPg()) > 0 Then
               vgNovaPg = UBound(vgBookPg()) - 1
            End If
            vgFP.vgProcuraPg = True
            vgIniProc = True                                           'flag para informar que iniciamos procura
            vgFP.vgNovaPag = 0
         End If
      End If
   End If
   If vgNovaPg >= 0 Then
      If vgPodeVoltar(vgNovaPg + 1) = False And vgNovaPg > 0 Then
         If Not vgFP.vgProcuraPg Then                                  'se nao esta indo para pg ou ultima pg
            vgFP.vgProcuraPg = True
            vgIniProc = True                                           'flag para informar que iniciamos procura
            vgFP.vgNovaPag = vgNovaPg + 1
         End If
         Do While vgPodeVoltar(vgNovaPg + 1) = False                   'acerto para não deixar voltar página com quebra em filho
            vgNovaPg = vgNovaPg - 1
         Loop
         If vgModoInt = 1 Then
            If vgNumPg < 0 Then
               vgFP.vgPagRef = vgFP.vgPagRef + 1
            End If
         Else
            vgFP.vgPagRef = vgFP.vgNovaPag
         End If
      End If
      vgFP.vgContaPg = vgNovaPg
      If Len(vgBookPg(vgNovaPg + 1)) > 0 Then
         Manutencao_contas_receber.BookMark = vgBookPg(vgNovaPg + 1)
      Else
         vgSemBookMark = 1
      End If
      vgNumPag = vgNovaPg + 1
      vgTot6 = vgPgvgTot6(vgNumPag)
      vgTot7 = vgPgvgTot7(vgNumPag)
      vgTot8 = vgPgvgTot8(vgNumPag)
      vgTot9 = vgPgvgTot9(vgNumPag)
      vgTot10 = vgPgvgTot10(vgNumPag)
      vgTot15 = vgPgvgTot15(vgNumPag)
      vgTot13 = vgPgvgTot13(vgNumPag)
      vgTot14 = vgPgvgTot14(vgNumPag)
      vgTot12 = vgPgvgTot12(vgNumPag)
      vgTot11 = vgPgvgTot11(vgNumPag)
      vgTot131 = vgPgvgTot131(vgNumPag)
      vgTot1311 = vgPgvgTot1311(vgNumPag)
      vgTot13111 = vgPgvgTot13111(vgNumPag)
   Else
      vgFP.vgPagRef = 1
   End If
   vgFP.botPara.Enabled = (vgFP.vgProcuraPg Or vgNumPg = G_IMP_ULT_PG) 'habilita/desabilita stop
   vgPrimeiroCab = True

Volta:

   'Contas a Receber
   If Manutencao_contas_receber.RecordCount > 0 Then
      If vgNumPg = 0 And Not (vgEmPapel And vgFP.vgContaPg > 0) Then
         Manutencao_contas_receber.MoveLast
         Manutencao_contas_receber.MoveFirst
      End If
      If vgEmPapel And vgNumPg = 0 And vgPrimeiroCab Or (vgNumPg = vgPagOrig And vgEmPapel) Then
         Set vgFG = New frmGauge
         vgQReg = 0
         Load vgFG
         Set vgFG.vgFrmOrigem = Me
         vgFG.MostraGauge
      End If
      vgLimiteInf = 196
      vgDifProxReg = 0
      On Error Resume Next
      Do While Not Manutencao_contas_receber.EOF Or vgSemBookMark = 1
         vgFP.vgFimDoArq = False
         If vgPrimeiroCab Or vgPosLin > vgLimiteInf Then
            vgPosLin = vgPosLin - vgDifProxReg
            vgDifProxReg = 0
            vgPagOk = True
            vgTestaPriCab = True
            GoSub VerificaFimPg
            GoSub OutraFolha
            On Error Resume Next
         End If
         If vgSemBookMark = 0 Then
            'vgQuebra20 = Manutencao_contas_receber![Data de Vencimento]
            'Inicio Manual
               If TipoFiltro = 0 Then
                  vgQuebra20 = Manutencao_contas_receber![Data de Entrada]
               End If
               If TipoFiltro = 1 Then
                  vgQuebra20 = Manutencao_contas_receber![Data de Vencimento]
               End If
               If TipoFiltro = 2 Then
                  vgQuebra20 = Manutencao_contas_receber![Razão Social]
               End If
               If TipoFiltro = 3 Then
                  vgQuebra20 = Manutencao_contas_receber!Documento
               End If
          ' Fim Manual
            vgPriQb20 = 0
            vgTot1 = 0
            vgTot2 = 0
            vgTot3 = 0
            vgTot4 = 0
            vgTot5 = 0
            If vgPrimeiroCab Then
               vgTot1 = vgPgvgTot1(vgNumPag)
               vgTot2 = vgPgvgTot2(vgNumPag)
               vgTot3 = vgPgvgTot3(vgNumPag)
               vgTot4 = vgPgvgTot4(vgNumPag)
               vgTot5 = vgPgvgTot5(vgNumPag)
            End If
            Do While Not Manutencao_contas_receber.EOF
               'If vgQuebra20 <> Manutencao_contas_receber![Data de Vencimento] Then Exit Do
               'Inicio Manual
                 If TipoFiltro = 0 Then
                    If vgQuebra20 <> Manutencao_contas_receber![Data de Entrada] Then Exit Do
                 End If
                   If TipoFiltro = 1 Then
                      If vgQuebra20 <> Manutencao_contas_receber![Data de Vencimento] Then Exit Do
                   End If
                   If TipoFiltro = 2 Then
                      If vgQuebra20 <> Manutencao_contas_receber![Razão Social] Then Exit Do
                   End If
                 If TipoFiltro = 3 Then
                    If vgQuebra20 <> Manutencao_contas_receber!Documento Then Exit Do
                 End If
              'Fim Manual
               If vgPosLin > vgLimiteInf - 8.73 Or (vgPriQb20 < 2 And vgPosLin > vgLimiteInf - 12.96) Then
                  vgPagOk = False
                  vgPosLin = vgPosLin - vgDifProxReg
                  vgDifProxReg = 0
                  vgTestaPriCab = False
                  GoSub VerificaFimPg
                  GoSub OutraFolha
                  On Error Resume Next
                  vgPriQb20 = vgPriQb20 - 1
               End If
               If vgPriQb20 < 2 Then
                  GoSub CabQb20
               End If
               vgTot1 = vgTot1 + Manutencao_contas_receber![Valor da Parcela]
               vgTot2 = vgTot2 + Manutencao_contas_receber![Valor Do juros]
               vgTot3 = vgTot3 + Manutencao_contas_receber![Valor Do Desconto]
               vgTot4 = vgTot4 + Manutencao_contas_receber![Valor pago]
               vgTot5 = vgTot5 + Manutencao_contas_receber![Valor restante]
               vgTot6 = vgTot6 + Manutencao_contas_receber![Valor da Parcela]
               vgTot7 = vgTot7 + Manutencao_contas_receber![Valor Do juros]
               vgTot8 = vgTot8 + Manutencao_contas_receber![Valor Do Desconto]
               vgTot9 = vgTot9 + Manutencao_contas_receber![Valor pago]
               vgTot10 = vgTot10 + Manutencao_contas_receber![Valor restante]
               vgTot15 = vgTot15 + IIf(Manutencao_contas_receber![Data de Vencimento] < Now - 1, Manutencao_contas_receber![Valor restante], 0)
               vgTot13 = vgTot13 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Dinheiro", Manutencao_contas_receber![Valor restante], 0)
               vgTot14 = vgTot14 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Boleto", Manutencao_contas_receber![Valor restante], 0)
               vgTot12 = vgTot12 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Cartão", Manutencao_contas_receber![Valor restante], 0)
               vgTot11 = vgTot11 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Cheque - Pre", Manutencao_contas_receber![Valor restante], 0)
               vgTot131 = vgTot131 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Condicional", Manutencao_contas_receber![Valor restante], 0)
               vgTot1311 = vgTot1311 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Garantia", Manutencao_contas_receber![Valor restante], 0)
               vgTot13111 = vgTot13111 + IIf(Manutencao_contas_receber![Especie de Pagamento] = "Deposito", Manutencao_contas_receber![Valor restante], 0)
               If Mudou Then
                  Imprime "", 1.32, vgPosLin + 1.05, "L=285|BOX=F14737632/1.32/" + Str$(vgPosLin + 1.05) + "/285.44/2.64", vgObSaida, , , , vgFP.vgProcuraPg
               End If
               Imprime Manutencao_contas_receber![Razão Social], 1.32, vgPosLin, "F=Courier New|S=9.75|L=61", vgObSaida, , , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber!Documento, 65.34, vgPosLin, "F=Courier New|W=700|S=9.75|L=25", vgObSaida, , , , vgFP.vgProcuraPg
               If TipoFiltro = 1 Or TipoFiltro = 2 Or TipoFiltro = 3 Then
                  Imprime CDate(Manutencao_contas_receber![Data de Entrada]), 94.44, vgPosLin, "F=Courier New|S=9.75|L=22", vgObSaida, "99/99/9999", , , vgFP.vgProcuraPg
               End If
               Imprime Format(Manutencao_contas_receber!Parcela, "000"), 119.84, vgPosLin, "F=Courier New|S=9.75|L=7", vgObSaida, "999", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Valor da Parcela], 191.22, vgPosLin, "F=Courier New|S=9.75|L=28|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Valor Do juros], 215.15, vgPosLin, "F=Courier New|S=9.75|L=26|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Valor Do Desconto], 236.2, vgPosLin, "F=Courier New|S=9.75|L=28|JD", vgObSaida, "99.999.999,99", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Valor pago], 257.95, vgPosLin, "F=Courier New|S=9.75|L=27|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Valor restante], 286.4, vgPosLin, "F=Courier New|S=9.75|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
               Imprime Manutencao_contas_receber![Especie de Pagamento], 143.91, vgPosLin, "F=Courier New|S=9.75|L=24", vgObSaida, , , , vgFP.vgProcuraPg
               Imprime IIf(Manutencao_contas_receber![Data de Vencimento] < Date, DateDiff("D", Manutencao_contas_receber![Data de Vencimento], Now), 0), 130.42, vgPosLin, "F=Courier New|S=9.75|C=&HFF&|L=10|JC=9.52", vgObSaida, "99999", , , vgFP.vgProcuraPg
               If False Then
                  Imprime Manutencao_contas_receber![Sequencia da Cobrança], 8.32, vgPosLin + 0.26, "F=Courier New|S=9.75|L=7|JD", vgObSaida, "999", , , vgFP.vgProcuraPg
               End If
               If Manutencao_contas_receber![Sequencia da Cobrança] > 0 Then
                  Imprime "Cobrança  --> " & Manutencao_contas_receber![Descrição da Cobrança], 1.32, vgPosLin + 3.43, "F=Courier New|W=700|S=9.75|C=&H8000&|L=61", vgObSaida, "@x", , , vgFP.vgProcuraPg
               End If
               If TipoFiltro = 0 Then
                  Imprime CDate(Manutencao_contas_receber![Data de Vencimento]), 94.44, vgPosLin, "F=Courier New|S=9.75|L=23", vgObSaida, "99/99/9999", , , vgFP.vgProcuraPg
               End If
               Imprime Manutencao_contas_receber!Fone, 64.28, vgPosLin + 4.49, "F=Courier New|S=9.75|L=29", vgObSaida, , , , vgFP.vgProcuraPg
               If TipoFiltro = 2 Or TipoFiltro = 3 Then
                  Imprime "Vencimento: - " & Manutencao_contas_receber![Data de Vencimento], 94.44, vgPosLin + 4.49, "F=Courier New|S=9.75|L=63", vgObSaida, , , , vgFP.vgProcuraPg
               End If
               vgPosLin = vgPosLin + 8.73
               vgDifProxReg = 0
               Manutencao_contas_receber.MoveNext
               vgDifProxReg = -1
               vgPosLin = vgPosLin + vgDifProxReg
            Loop
            vgPosLin = vgPosLin - vgDifProxReg
            vgDifProxReg = 0
            vgPosLin = vgPosLin + -3.97
            If vgPosLin + 4.76 > vgLimiteInf Then
             vgPagOk = False
             vgTestaPriCab = False
             vgTemFimRel = True
             GoSub VerificaFimPg
             GoSub OutraFolha
            End If
            Imprime "", 166.93, vgPosLin, "LIN=B0/" + Trim(Str(166.93)) + "/" + Trim$(Str$(vgPosLin)) + "/" + Trim(Str(287.03)) + "/" + Trim$(Str$(vgPosLin)), vgObSaida, , , , vgFP.vgProcuraPg
            Imprime vgTot1, 191.45, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=28|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
            Imprime vgTot2, 214.97, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
            Imprime vgTot3, 236.2, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=28|JD", vgObSaida, "99.999.999,99", , , vgFP.vgProcuraPg
            Imprime vgTot4, 257.95, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=27|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
            Imprime vgTot5, 286.45, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=26|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
            vgPosLin = vgPosLin + 4.76
            vgDifProxReg = 0
            vgDifProxReg = 1
            vgPosLin = vgPosLin + vgDifProxReg
         Else
            vgSemBookMark = 0
         End If
         If vgEmPapel And vgNumPg = 0 Then
            vgQReg = vgQReg + 1
            If vgModoInt = 1 Then
               If Not vgFP.vgProcuraPg Then
                  vgFG.pgb1.Value = vgFP.vgPagRef * 100 / UBound(vgPagRef())
               End If
            Else
               vgFG.pgb1.Value = vgQReg * 100 / Manutencao_contas_receber.RecordCount
            End If
            If vgFG.vgQuerCancelar Then Exit Do
         End If
      Loop
      If vgEmPapel Then
         If vgFG.vgQuerCancelar Then GoTo CancelouRel
      End If
      vgPosLin = vgPosLin - vgDifProxReg
      vgDifProxReg = 0
      If vgPosLin > vgLimiteInf - 26.45 Then
         vgPagOk = False
         vgTemFimRel = True
         vgTestaPriCab = False
         GoSub VerificaFimPg
         GoSub OutraFolha
      End If
      vgPosLin = vgPosLin + 0.52
      Imprime "", 166.93, vgPosLin, "LIN=B0/" + Trim(Str(166.93)) + "/" + Trim$(Str$(vgPosLin)) + "/" + Trim(Str(287.03)) + "/" + Trim$(Str$(vgPosLin)), vgObSaida, , , , vgFP.vgProcuraPg
      Imprime vgTot6, 191.81, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=27|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot7, 215.15, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=26|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot8, 236.2, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=28|JD", vgObSaida, "99.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot9, 257.83, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot10, 286.39, vgPosLin + 0.52, "F=Courier New|W=700|S=9.75|L=28|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime "Cheque - Pre", 6.61, vgPosLin + 16.66, "F=Microsoft Sans Serif|W=700|S=9.75|L=25|BOX=F8454143/6.61/" + Str$(vgPosLin + 16.66) + "/25.39/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Cartão", 50.79, vgPosLin + 16.66, "F=Microsoft Sans Serif|W=700|S=9.75|L=13|BOX=F14737632/50.79/" + Str$(vgPosLin + 16.66) + "/12.69/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Boleto", 83.59, vgPosLin + 16.66, "F=Microsoft Sans Serif|W=700|S=9.75|L=13|BOX=F14737632/83.59/" + Str$(vgPosLin + 16.66) + "/12.69/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Dinheiro", 111.11, vgPosLin + 16.66, "F=Microsoft Sans Serif|W=700|S=9.75|L=16|BOX=F14737632/111.11/" + Str$(vgPosLin + 16.66) + "/15.87/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Vencidas", 244.44, vgPosLin + 16.4, "F=Microsoft Sans Serif|W=700|S=9.75|L=18|BOX=F12632319/244.44/" + Str$(vgPosLin + 16.4) + "/17.98/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime vgTot15, 262.86, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot13, 127.14, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot14, 96.43, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=30|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot12, 64.39, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=30|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot11, 31.58, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=30|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime "Condicional", 138.88, vgPosLin + 16.4, "F=Microsoft Sans Serif|W=700|S=9.75|L=23|BOX=F14737632/138.88/" + Str$(vgPosLin + 16.4) + "/23.01/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Garantia", 176.98, vgPosLin + 16.4, "F=Microsoft Sans Serif|W=700|S=9.75|L=17|BOX=F14737632/176.98/" + Str$(vgPosLin + 16.4) + "/16.66/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime "Deposito", 209.78, vgPosLin + 16.4, "F=Microsoft Sans Serif|W=700|S=9.75|L=18|BOX=F14737632/209.78/" + Str$(vgPosLin + 16.4) + "/17.72/4.23", vgObSaida, , , , vgFP.vgProcuraPg
      Imprime vgTot131, 161.53, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot1311, 194.34, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      Imprime vgTot13111, 227.94, vgPosLin + 21.95, "F=Courier New|W=700|S=11.25|L=29|JD", vgObSaida, "9.999.999,99", , , vgFP.vgProcuraPg
      vgPosLin = vgPosLin + 26.45
      vgDifProxReg = 0
      If vgFP.vgProcuraPg And UBound(vgBookPg) <= vgFP.vgContaPg Then               'ops... estava buscando por uma página, e ela não existe (chegou no final do arquivo)!
         If vgNumPg = G_IMP_ULT_PG Then
            vgPrimeiroCab = True
            vgFP.vgProcuraPg = True
            vgNumPg = UBound(vgBookPg())
            If Len(vgBookPg(vgNumPg)) = 0 Then     'se a última não tem bookmark, quer dizer que chegou em fim de arquivo mas tem complemento a imprimir em outra folha
               Do While Len(vgBookPg(vgNumPg)) = 0   'vamos pegar a última página com bookmark válido
                  vgNumPg = vgNumPg - 1
               Loop
               vgFP.vgContaPg = vgNumPg              'vamos voltar para última página com bookmark
               vgNumPg = G_IMP_PROX_PG               'e força navegação para próxima (que será a última)
               vgFP.vgProcuraPg = False              'sem procurar...
            End If
            GoTo OutraCopia
         Else
            vgPrimeiroCab = True
            vgFP.vgProcuraPg = False
            If vgModoInt = 1 Then
               vgPagRef(vgFP.vgPagRef) = -vgFP.vgContaPg - 1        'vamos forçar a exclusão de páginas maiores que a última no vetor de referências
               vgFP.vgContaPg = vgPagRef(vgFP.vgPagRef - 1) - 1
               j = 0
               Do While vgPagRef(vgFP.vgPagRef - 1 - j) - 1 <> 0 And vgPagRef(vgFP.vgPagRef - 1 - j) - 1 > UBound(vgBookPg) And vgFP.vgPagRef - j > 0
                  j = j + 1
                  vgFP.vgContaPg = vgPagRef(vgFP.vgPagRef - 1 - j) - 1
               Loop
               If vgFP.vgContaPg > 0 Then
                  Manutencao_contas_receber.BookMark = vgBookPg(vgFP.vgContaPg + 1)
               Else
                  vgFP.vgContaPg = 0
                  Manutencao_contas_receber.MoveFirst
               End If
            End If
            If vgModoInt = 1 Then
               GoTo Volta
             ElseIf vgFP.vgNovaPag > vgFP.vgContaPg Then
               GoSub RestauraPag
               GoTo NaoAchouPg
            End If
         End If
      ElseIf vgEmPapel And vgNumPg = 0 And vgFP.vgPagRef < UBound(vgPagRef()) Then 'estava imprimindo intervalo de página e parece ter chegado no fim do arquivo
         If Not vgFG.vgQuerCancelar Then
            vgPosLin = vgLimiteInf + 99                  'força um cabec na próxima vez
            vgFP.vgContaPg = 0                           'zera o contador de página
            GoTo Volta                                   'e vamos voltar para imprimir a próxima página
         End If
      End If
   End If

CancelouRel:
   If vgEmPapel Then
      If vgNumPg > 0 And Not vgPrimeiroCab And Not vgFG.vgQuerCancelar Then
         Unload vgFG
         Set vgFG = Nothing
         vgObSaida.EndDoc
      Else
         If vgFG.vgQuerCancelar Then
            vgCancelou = True
            vgObSaida.KillDoc
         Else
            vgObSaida.EndDoc
         End If
         If vgNumPg = 0 And vgPagOrig > 0 Then                      'solicitou impressão de todo relatório pelo preview
            vgFP.vgPagRef = vgPagOrig
            If vgModoInt = 1 Then
               vgFP.vgContaPg = vgPagRef(vgPagOrig)                     'e arruma o número da página
            Else
               vgFP.vgContaPg = vgPagOrig                               'e arruma o número da página
            End If
         End If
         vgNCop = vgNCop - 1
         If vgFG.vgQuerCancelar = False And vgNCop > 0 Then
            Unload vgFG
            Set vgFG = Nothing
            GoTo OutraCopia
         End If
      End If
      With Printer
         .ScaleMode = vgPsm
         .Orientation = vgPor
         .PaperSize = vgPsz
         .FontName = vgPfn
      End With
   Else
      GoSub SalvaPagina
   End If
   GoTo SaiDaSub

VerificaFimPg:
   If Not vgEmPapel And (Not vgTestaPriCab Or vgPrimeiroCab = False) Then
      GoSub SalvaPagina
      If Not vgFP.vgProcuraPg Then                      'se não está procurando por uma página
         If vgFP.vgContaPg <> vgFP.vgNovaPag Then                   'opa... parece que o usuário precionou o STOP
            If vgNumPg = G_IMP_ANT_PG Or vgNumPg = G_IMP_PROX_PG Then 'se estava movendo para próxima página ou anterior
               vgFP.vgContaPg = vgPagOrig                               'vamos restaurar a página original
               If vgNumPg = G_IMP_ANT_PG Then                           'restaura o contador
                  vgFP.vgPagRef = vgFP.vgPagRef + 1
               Else
                  vgFP.vgPagRef = vgFP.vgPagRef - 1
               End If
               GoSub SalvaPagina                    'ajusta a botoeira do preview novamente
            End If
         End If
         GoTo SaiDaSub                          'cai fora
      End If
   ElseIf vgEmPapel And vgNumPg > 0 And Not vgPrimeiroCab Then
      If vgFP.vgContaPg = vgFP.vgNovaPag Then GoTo CancelouRel
   ElseIf vgNumPg = 0 And Not vgPrimeiroCab Then
      If Not vgFP.vgProcuraPg Then
         If vgModoInt = 1 And vgFP.vgPagRef >= UBound(vgPagRef()) Then         'se estava imprimindo intervalo de páginas...
            GoTo CancelouRel                                                     'se chegou à última página do intervalo, cai fora
         End If
         GoSub SalvaInicioPg
      End If
   End If
   Return

OutraFolha:
   If vgEmPapel = False Then
      GoSub SalvaInicioPg
      If vgFP.vgProcuraPg And vgNumPg = G_IMP_ULT_PG Then
         vgFP.txtPag.Text = Trim$(Str$(vgFP.vgContaPg + 1))
         vgFP.txtPag.Refresh
      End If
   ElseIf Not vgSemControlePg Then
      If Not vgPrimeiroCab And Not vgFP.vgProcuraPg Then vgObSaida.NewPage
   End If
   vgFP.vgContaPg = vgFP.vgContaPg + 1
   If vgFP.vgProcuraPg = False Then
    If vgModoInt = 0 Then
     vgFP.vgNovaPag = vgFP.vgContaPg
    Else
     If vgPagRef(vgFP.vgPagRef) < 0 Then
      GoSub ArrumaVetor
     Else
      If vgNumPg = G_IMP_PROX_PG Or vgNumPg > vgFP.vgContaPg Or (vgNumPg = 0 And vgModoInt = 0) Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       Else                                   'usuário cancelou busca pela próxima página
        vgFP.vgPagRef = vgFP.vgPagRef - 1
       End If
      ElseIf vgNumPg = G_IMP_ANT_PG Or (vgNumPg > 0 And vgNumPg < vgFP.vgContaPg) Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = vgFP.vgPagRef - 1
       Else                                   'usuário cancelou busca pela página anterior
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       End If
      ElseIf vgNumPg = G_IMP_PRI_PG Or vgNumPg = 0 Then
       If vgPrimeiroCab Then
        vgFP.vgPagRef = 1
       Else
        vgFP.vgPagRef = vgFP.vgPagRef + 1
       End If
      ElseIf vgNumPg = G_IMP_ULT_PG Then
       If Not vgIniProc Then                  'se não estava procurando...
        vgIniProc = True                      'agora seta flag de procura
        vgFP.vgPagRef = UBound(vgPagRef())
       Else
        'usuário cancelou busca para última página, vamos procurar o último bookmark existente no intervalo de páginas especificado
        For j = UBound(vgPagRef()) To 1 Step -1           'vamos correr as páginas de referência, de trás para frente
         If vgPagRef(j) <= UBound(vgBookPg) Then         'se tem bookmark para a página correspondente
          vgFP.vgPagRef = j                               'vamos utilizar essa página
          Exit For
         End If
        Next
       End If
      End If
     End If
     vgFP.vgNovaPag = vgPagRef(vgFP.vgPagRef)             'pega o número de página real referente ao número da página no intervalo
     If vgFP.vgNovaPag = 0 Then GoTo CancelouRel          'ops... vamos sair do relatório
     If vgFP.vgNovaPag <> vgFP.vgContaPg Then             'se a página atual não for a página procurada
      If vgFP.vgNovaPag <= UBound(vgBookPg()) Then       'se tem bookmark para a página procurada
       vgFP.vgContaPg = vgFP.vgNovaPag                    'vamos posicionar diretamente na nova página
       Do While vgPodeVoltar(vgFP.vgContaPg) = False      'acerto para não deixar voltar página com quebra em filho
        vgFP.vgContaPg = vgFP.vgContaPg - 1
       Loop
       vgFP.vgProcuraPg = (vgFP.vgContaPg <> vgFP.vgNovaPag)
       Manutencao_contas_receber.BookMark = vgBookPg(vgFP.vgContaPg)      'seta o bookmark correto
       If Not vgEmPapel Then
        vgFP.picAux.Cls                                    'limpa a área
        vgFP.picAux.Picture = LoadPicture("")              'para imprimir a página
       End If
      Else                                                'não tem bookmark...
       vgFP.vgProcuraPg = True                            'então vamos procurar
      End If
     Else
      If Not vgEmPapel Then
       vgFP.picAux.Cls                                     'limpa a área
       vgFP.picAux.Picture = LoadPicture("")               'para imprimir a página
      End If
     End If
    End If
   ElseIf vgFP.vgNovaPag = vgFP.vgContaPg Then            'se a página atual é a página procurada
    vgFP.vgProcuraPg = False                              'vamos cancelar a procura!
   End If
   If vgModoInt = 0 Then                                  'se não tem intervalo de páginas especificado
    vgNumPag = vgFP.vgContaPg
    vgFP.vgPagRef = vgFP.vgContaPg                      'com o mesmo número para página de referência
   Else                                                   'senão... tem intervalo de página
    vgNumPag = vgFP.vgNovaPag                             'então pela o número da nova página
   End If
   vgPosLin = 35.97
   Imprime vgFP.picFixa(1), 0.67, 2.38, "FIG=FigFixa1.jpg", vgObSaida, , , , vgFP.vgProcuraPg
   If Not Vazio(Dt_Inicial) And Not Vazio(Dt_Final) Then
      Imprime "De: " & Dt_Inicial & " a " & Dt_Final, 149.2, 7.93, "F=Courier New|W=700|S=15.75", vgObSaida, , , , vgFP.vgProcuraPg
   End If
   Imprime "Contas a Receber - Por Vencimento", 32.8, 7.67, "F=Courier New|W=700|S=15.75", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime CDate(vgDataInicial), 255.16, 5.02, "F=Courier New|S=9.75|L=23", vgObSaida, "99/99/9999", , , vgFP.vgProcuraPg
   Imprime CDate(vgDataInicial), 275.96, 5.02, "F=Courier New|S=9.75|L=11", vgObSaida, "99:99", , , vgFP.vgProcuraPg
   Imprime vgNumPag, 286.96, 1.32, "F=Courier New|S=9.75|L=11|JD", vgObSaida, "9999", , , vgFP.vgProcuraPg
   Imprime "", 0.12, 35.55, "LIN=B0/" + Trim(Str(0.12)) + "/" + Trim$(Str$(35.55)) + "/" + Trim(Str(286.8)) + "/" + Trim$(Str$(35.55)), vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Página:", 278.88, 1.32, "F=Courier New|S=9.75|L=18|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Geral", 1.32, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=12", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Doc", 65.22, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=10", vgObSaida, , , , vgFP.vgProcuraPg
   If TipoFiltro = 1 Or TipoFiltro = 2 Or TipoFiltro = 3 Then
      Imprime "Dt. Emissão", 93.79, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=21", vgObSaida, , , , vgFP.vgProcuraPg
   End If
   Imprime "Pc.", 119.31, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=8", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Vr. Parcela", 190.98, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=21|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Vr. Pago", 258.53, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=17|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Vr. Juros", 214.49, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=17|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Vr. Desc", 236.24, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=16|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Vr. Restante", 286.26, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=23|JD", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "F. Pagto:", 143.91, 30.68, "F=Microsoft Sans Serif|W=700|S=9|L=16", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Dias", 130.42, 27.77, "F=Microsoft Sans Serif|W=700|S=9|C=&HFF&|L=10", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Atrazo", 129.36, 30.68, "F=Microsoft Sans Serif|W=700|S=9|C=&HFF&|L=12", vgObSaida, , , , vgFP.vgProcuraPg
   If Vendedor > 0 Then
      Imprime MostraVendedor(), 32.8, 20.37, "F=Courier New|W=700|S=15.75|L=189", vgObSaida, , , , vgFP.vgProcuraPg
   End If
   If TipoFiltro = 0 Then
      Imprime "Dt. Vencto", 93.91, 30.95, "F=Microsoft Sans Serif|W=700|S=9|L=24", vgObSaida, , , , vgFP.vgProcuraPg
   End If
   Imprime MostraOrdem() & " - " & F_Pagto, 32.8, 14.02, "F=Courier New|W=700|S=15.75|L=222", vgObSaida, , , , vgFP.vgProcuraPg
   Imprime "Fone", 65.22, 35.97, "F=Microsoft Sans Serif|W=700|S=9|L=12", vgObSaida, , , , vgFP.vgProcuraPg
   vgPrimeiroCab = False
   vgTemFimRel = False
   Return

CabQb20:
   Imprime DefineTipo(), 1.32, vgPosLin, "F=Courier New|W=700|S=11.25|L=61", vgObSaida, , , , vgFP.vgProcuraPg
   vgPosLin = vgPosLin + 4.23
   vgDifProxReg = 0
   vgPriQb20 = 2
   Return

SalvaInicioPg:
   If (vgPV Or vgFP.vgContaPg + 1 > UBound(vgBookPg)) And (vgTemFimRel Or Not Manutencao_contas_receber.EOF) Then 'testa se precisa mais variáveis
      If vgFP.vgContaPg + 1 > UBound(vgBookPg) Then
         ReDim Preserve vgBookPg(vgFP.vgContaPg + 1) As Variant   'cria variável para salvar bookmark
         ReDim Preserve vgPodeVoltar(vgFP.vgContaPg + 1) As Boolean 'controle de quebra em filhos
         ReDim Preserve vgPgvgTot1(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot2(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot3(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot4(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot5(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot6(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot7(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot8(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot9(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot10(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot15(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot13(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot14(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot12(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot11(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot131(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot1311(vgFP.vgContaPg + 1) As Double
         ReDim Preserve vgPgvgTot13111(vgFP.vgContaPg + 1) As Double
      vgPgvgTot1(vgFP.vgContaPg + 1) = vgTot1
      vgPgvgTot2(vgFP.vgContaPg + 1) = vgTot2
      vgPgvgTot3(vgFP.vgContaPg + 1) = vgTot3
      vgPgvgTot4(vgFP.vgContaPg + 1) = vgTot4
      vgPgvgTot5(vgFP.vgContaPg + 1) = vgTot5
      vgPgvgTot6(vgFP.vgContaPg + 1) = vgTot6
      vgPgvgTot7(vgFP.vgContaPg + 1) = vgTot7
      vgPgvgTot8(vgFP.vgContaPg + 1) = vgTot8
      vgPgvgTot9(vgFP.vgContaPg + 1) = vgTot9
      vgPgvgTot10(vgFP.vgContaPg + 1) = vgTot10
      vgPgvgTot15(vgFP.vgContaPg + 1) = vgTot15
      vgPgvgTot13(vgFP.vgContaPg + 1) = vgTot13
      vgPgvgTot14(vgFP.vgContaPg + 1) = vgTot14
      vgPgvgTot12(vgFP.vgContaPg + 1) = vgTot12
      vgPgvgTot11(vgFP.vgContaPg + 1) = vgTot11
      vgPgvgTot131(vgFP.vgContaPg + 1) = vgTot131
      vgPgvgTot1311(vgFP.vgContaPg + 1) = vgTot1311
      vgPgvgTot13111(vgFP.vgContaPg + 1) = vgTot13111
      End If
      If Not Manutencao_contas_receber.EOF Then
         vgBookPg(vgFP.vgContaPg + 1) = Manutencao_contas_receber.BookMark 'salva inicio desta página
      Else
         vgBookPg(vgFP.vgContaPg + 1) = ""
         vgPodeVoltar(vgFP.vgContaPg + 1) = True                           'pode voltar preview para esta página?
      End If
      vgPodeVoltar(vgFP.vgContaPg + 1) = vgPagOk Or vgPrimeiroCab          'pode voltar preview para esta página?
   End If
   If Not vgPriVez Then
      vgPagOk = False
   End If
   Return

SalvaPagina:
   If Not vgEmPapel Or (vgEmPapel And vgNumPg = 0) Then
      With vgFP
         If .vgFimDoArq = False And Manutencao_contas_receber.EOF Then
            .vgFimDoArq = True
         End If
         Ok = ((((vgTemFimRel Or .vgFimDoArq = False) Or .vgContaPg < .vgUltPagVista) And vgModoInt = 0) Or ((.vgPagRef < UBound(vgPagRef())) And vgModoInt = 1)) And Not .vgProcuraPg
         If Ok <> .botMais.Enabled Then
            .botMais.Enabled = Ok                                          'habilita/desabilita botões do preview
         End If
         Ok = .botMais.Enabled And Not .vgProcuraPg
         If Ok <> .botUltimo.Enabled Then
            .botUltimo.Enabled = Ok
         End If
         Ok = (.vgPagRef > 1 And Not .vgProcuraPg)
         If Ok <> .botVolta.Enabled Then
            .botVolta.Enabled = Ok
         End If
         Ok = (.vgPagRef > 1 And Not .vgProcuraPg)
         If Ok <> .botPrimeiro.Enabled Then
            .botPrimeiro.Enabled = Ok
         End If
         GoSub SalvaInicioPg
         If .vgProcuraPg Or (.vgContaPg <> .vgNovaPag And .vgNovaPag <> 0) Then 'procurando pg adiante...
            If .vgProcuraPg And ((vgModoInt = 0 And (Not vgTemFimRel And Manutencao_contas_receber.EOF)) Or .vgContaPg = .vgNovaPag) Then
               .ShowPreview                       'mostra preview
            Else
               .picAux.Cls                        'limpa a área para a próxima pg
               vgPosLin = vgLimiteInf + 99        'força um cabec na próxima vez
            End If
         Else
            .ShowPreview
         End If
      End With
   End If
   Return
'retira as páginas inválidas do intervalo especificado!
ArrumaVetor:
   vgQt = 0
   j = 0
   k = 0
   For i = 1 To UBound(vgPagRef())
      If vgPagRef(i) < 0 Or (Abs(vgPagRef(i)) >= k And k > 0) And i < UBound(vgPagRef()) - vgQt + 1 Then
         If vgPagRef(i) < k And vgPagRef(i) <> 0 Then k = Abs(vgPagRef(i))
         For j = 1 To UBound(vgPagRef())
            If Abs(vgPagRef(j)) >= k And j <= UBound(vgPagRef()) - vgQt + 1 Then
               vgQt = vgQt + 1
               For m = j + 1 To UBound(vgPagRef())
                  vgPagRef(m - 1) = vgPagRef(m)
               Next
               If j <= UBound(vgPagRef()) - vgQt Then
                  j = j - 1
               End If
            End If
         Next
         If i < UBound(vgPagRef()) Then
            i = i - 1
         End If
      End If
   Next
   ReDim Preserve vgPagRef(UBound(vgPagRef()) - vgQt) As Integer
   
   If vgFP.vgPagRef > UBound(vgPagRef()) Then
      vgFP.vgPagRef = UBound(vgPagRef())
   End If
   Return

RestauraPag:
   vgFP.txtPag.Text = CStr(vgPagOrig)
   vgFP.botPara.Enabled = False
   vgFP.vgProcuraPg = False
   If vgPagOrig < vgFP.vgContaPg Then             'se não for a última página, vamos restaurar o bookmark para atualizar os botões do preview
      Manutencao_contas_receber.BookMark = vgBookPg(vgPagOrig)
   End If
   vgFP.vgContaPg = vgPagOrig
   vgFP.vgPagRef = vgPagOrig
   Return

IniciaImpressao:
   On Error Resume Next                           'algumas impressoras não tem certas propriedades
   If vgEmPapel Then
      With Printer                                'salva configs da impressora
         vgPsm = .ScaleMode
         vgPor = .Orientation
         vgPsz = .PaperSize
         vgPfn = .FontName
         .PaperSize = vbPRPSA4
         .Orientation = vbPRORLandscape
         If Err Then Err.Clear
         If vgNumPg <> 0 Then                     'procurando pg para imprimir
            .Copies = 1
         Else
            .Copies = vgCopias                    'número de cópias configurada no diálogo
            If Err And vgNCop = 0 Then vgNCop = vgCopias
         End If
      End With
   End If
   With vgObSaida.Font
      'ajusta para nós...
      .Name = "MS Sans Serif"
      .Charset = 0
      .Weight = 400
      .Size = 8.25
      .Underline = 0
      .Italic = 0
      .Strikethrough = 0
      vgObSaida.ForeColor = &H0
      vgObSaida.ScaleMode = vbMillimeters         'vamos trabalhar em mm
   End With
   On Error GoTo DeuErro                          'vamos cercar erros...
   Return

NaoAchouPg:
   GoSub SalvaPagina
   vgFP.txtPag.Text = CStr(vgPagOrig)
   MsgBox "A página especificada não foi encontrada no relatório!", vbExclamation
   GoTo FimDaSub

DeuErro:                                          'deu um erro
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
   If Err <> 65524 Then                           'se não for um erro de crítica
      CErr.NumErro = Err
      CErr.Show
   End If
   vgPrimeiroCab = 2                              'flag para não dar a msg novamente

SaiDaSub:                                         'saída comum da sub...

FimDaSub:
   If Not vgFG Is Nothing Then
      Unload vgFG
      Set vgFG = Nothing
   End If
   Err.Clear
   Screen.MousePointer = vbDefault                'ponteiro do mouse = default
End Sub

'prepara o intervalo de páginas a ser impresso
Public Property Let vgIntervalPag(ByVal NovoValor As String)
   Dim x As String, z As String, i As Integer, j As Integer, k As Integer
   ReDim vgPagRef(0) As Integer
   vgIntPag$ = NovoValor$
   x$ = Substitui(vgIntPag$, ";", ",", UM_A_UM)
   j = 0
   Do While Len(x$) > 0
      z$ = Parse(x$, ",")
      i = InStr(z$, "-")
      If i > 0 Then
         k = Val(Mid(z$, i + 1))
         i = Val(Left(z$, i - 1))
      Else
         i = Val(z$)
         k = Val(z$)
      End If
      For i = i To k Step ((i > k) ^ (i > k))
         j = j + 1
         ReDim Preserve vgPagRef(j) As Integer
         vgPagRef(j) = i
      Next
   Loop
End Property

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   vgDestRel = Val(labopcPainel1.Caption)
   vgModoInt = Val(labopcPainel2.Caption)
   vgIntPag = txtCampo(0).Value
   vgCopias = txtCampo(1).Value
   Zebrar = chkCampo(0).Value
   Dt_Inicial = txtCampo(2).Value
   Dt_Final = txtCampo(3).Value
   Nome = txtCampo(4).Value
   Documento = txtCampo(5).Value
   Geral = txtCampo(6).Value
   F_Pagto = txtCampo(7).Value
   NFe = txtCampo(8).Value
   Vendedor = txtCampo(9).Value
   Cobranca = txtCampo(10).Value
   TipoFiltro = Val(labopcPainel0.Caption)
   Avulsa = chkCampo(1).Value
   Empresa = Val(labopcFrame0.Caption)
   vgIntervalPag = vgIntPag
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 6
         txtCampo(6).Filter = "[Sequencia do geral] > " & 0 & ""
                           txtCampo(6).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ")"
      Case 9
         txtCampo(9).Filter = "([Sequencia do geral] > " & 0 & ") AND Vendedor = " & 1 & ""
      Case 10
         txtCampo(10).Filter = "[Sequencia da Cobrança] > " & 0 & ""
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
      If vgDestRel = 1 Then
         vgOk = (vgCopias > 0)
         vgMsg$ = "Número de cópias tem que ser maior que ZERO!"
         If Not vgOk Then txtCampo(1).SetFocus
      End If
      If vgOk Then
         If vgModoInt = 1 Then
            vgOk = (Len(vgIntPag) > 0)
            vgMsg$ = "Esta informação é necessária"
            If Not vgOk Then txtCampo(0).SetFocus
         End If
      End If
      If vgOk Then vgMsg$ = ""
      DoEvents
   ElseIf vgOq = INICIALIZACOES Then
      InicializaApelidos COM_TEXTBOX
      On Error Resume Next
      opcPainel1(0).Value = True
      opcPainel1(0).CtPri.TabStop = False
      opcPainel2(0).Value = True
      opcPainel2(0).CtPri.TabStop = False
      chkCampo(0).Value = 1
      opcPainel0(1).Value = True
      opcPainel0(1).CtPri.TabStop = False
      chkCampo(1).Value = False
      opcFrame0(0).Value = True
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
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
   Ok = (vgModoInt = 1)
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (vgDestRel = 1)
   txtCampo(1).Enabled = (Ok And txtCampo(1).Editable)
   Ok = (Geral = 0)
   txtCampo(4).Enabled = (Ok And txtCampo(4).Editable)
   Ok = (NFe = 0)
   txtCampo(5).Enabled = (Ok And txtCampo(5).Editable)
   Ok = (Nome = "")
   txtCampo(6).Enabled = (Ok And txtCampo(6).Editable)
   Ok = (Documento = "")
   txtCampo(8).Enabled = (Ok And txtCampo(8).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 6 Or Index = 7 Or Index = 9 Or Index = 10 Then
      ExecutaPreValidacao
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
         On Error Resume Next
         Select Case Index
         Case 1
            If ValBrasil(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = 1
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 2
            If Len(txtCp(2).Text) = 0 Then
               txtCampo(2).Value = DateSerial(Year(Date), Month(Date), 1)
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 3
            If Len(txtCp(3).Text) = 0 Then
               txtCampo(3).Value = UltD(Dt_Inicial)
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaPreValidacao
            End If
         Case 6
            PoeRelEFiltroCbo 6
         Case 9
            PoeRelEFiltroCbo 9
         Case 10
            PoeRelEFiltroCbo 10
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
   ExecutaPreValidacao                            'habilita/desabilita campos
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX
   ExecutaPreValidacao                            'habilita/desabilita campos
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

'evento - quando uma opção for selecionada
Private Sub opcPainel1Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel1(Index).Locked Then
      opcPainel1(Val(labopcPainel1.Caption)).Value = True
   Else
      If Val(labopcPainel1.Caption) <> opcPainel1(Index).BookMark Then
         labopcPainel1.Caption = Str$(opcPainel1(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaPreValidacao
         opcPainel1(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
   opcPainel0(3).CtPri.TabStop = False
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel1Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel1(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel1Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel1(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcPainel1Cp_GotFocus(Index As Integer)
   opcPainel1(Index).GotFocus
         On Error Resume Next
         Select Case Index
         Case 6
            PoeRelEFiltroCbo 6
         Case 9
            PoeRelEFiltroCbo 9
         Case 10
            PoeRelEFiltroCbo 10
      End Select
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel1Cp_LostFocus(Index As Integer)
   opcPainel1(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcPainel2Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel2(Index).Locked Then
      opcPainel2(Val(labopcPainel2.Caption)).Value = True
   Else
      If Val(labopcPainel2.Caption) <> opcPainel2(Index).BookMark Then
         labopcPainel2.Caption = Str$(opcPainel2(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaPreValidacao
         opcPainel2(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
   opcPainel0(3).CtPri.TabStop = False
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel2Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel2(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel2Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel2(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcPainel2Cp_GotFocus(Index As Integer)
   opcPainel2(Index).GotFocus
         On Error Resume Next
         Select Case Index
         Case 6
            PoeRelEFiltroCbo 6
         Case 9
            PoeRelEFiltroCbo 9
         Case 10
            PoeRelEFiltroCbo 10
      End Select
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel2Cp_LostFocus(Index As Integer)
   opcPainel2(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcPainel0Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcPainel0(Index).Locked Then
      opcPainel0(Val(labopcPainel0.Caption)).Value = True
   Else
      If Val(labopcPainel0.Caption) <> opcPainel0(Index).BookMark Then
         labopcPainel0.Caption = Str$(opcPainel0(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaPreValidacao
         opcPainel0(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
   opcPainel0(3).CtPri.TabStop = False
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel0Cp_KeyPress(Index As Integer, KeyAscii As Integer)
   opcPainel0(Index).KeyPress KeyAscii
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub opcPainel0Cp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   opcPainel0(Index).KeyDown KeyCode, Shift
End Sub

'evento - quando o campo receber o foco
Private Sub opcPainel0Cp_GotFocus(Index As Integer)
   opcPainel0(Index).GotFocus
         On Error Resume Next
         Select Case Index
         Case 6
            PoeRelEFiltroCbo 6
         Case 9
            PoeRelEFiltroCbo 9
         Case 10
            PoeRelEFiltroCbo 10
      End Select
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel0Cp_LostFocus(Index As Integer)
   opcPainel0(Index).LostFocus
End Sub

'evento - quando uma opção for selecionada
Private Sub opcFrame0Cp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If opcFrame0(Index).Locked Then
      opcFrame0(Val(labopcFrame0.Caption)).Value = True
   Else
      If Val(labopcFrame0.Caption) <> opcFrame0(Index).BookMark Then
         labopcFrame0.Caption = Str$(opcFrame0(Index).BookMark)
         InicializaApelidos COM_TEXTBOX
         ExecutaPreValidacao
         opcFrame0(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
   opcPainel0(3).CtPri.TabStop = False
   opcPainel1(0).CtPri.TabStop = False
   opcPainel1(1).CtPri.TabStop = False
   opcPainel2(0).CtPri.TabStop = False
   opcPainel2(1).CtPri.TabStop = False
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
         On Error Resume Next
         Select Case Index
         Case 6
            PoeRelEFiltroCbo 6
         Case 9
            PoeRelEFiltroCbo 9
         Case 10
            PoeRelEFiltroCbo 10
      End Select
End Sub

'evento - quando o campo perder o foco
Private Sub opcFrame0Cp_LostFocus(Index As Integer)
   opcFrame0(Index).LostFocus
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Dim x As String
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(33400)
   vgFormID = 662
   vgIdentTab = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_RELAT
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   DesligaRepara
   ReDim vgBookPg(0) As Variant, vgPodeVoltar(0) As Boolean
   Load vgFP
   With vgFP
      vgIntervalPag = ""
      Set .vgFormRel = Me
      .vgContaPg = 0
      vgWidthRel = 520 + .ScaleX(286.8078, vbMillimeters, .ScaleMode)
      vgHeightRel = 720 + .ScaleY(196.9841, vbMillimeters, .ScaleMode)
      .picAux.Width = vgWidthRel
      .picAux.Height = vgHeightRel
      .Caption = LoadGasString(7505) + Me.Caption
      .picInterna.Move 0, 0, vgWidthRel, vgHeightRel
      .picAux.ScaleMode = vbMillimeters
      .picAux.Scale (0, 0)-(286.8078, 196.9841)
   End With
   vgWidthRel = vgWidthRel - 520
   vgHeightRel = vgHeightRel - 720
   vgVisivel = True
   Label(4).Caption = Printer.DeviceName
   vgSituacao = ACAO_NAVEGANDO
   Set botConfImp.Picture = LoadPicture(LoadGasPicture(24905))
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(24906))
   Set botOk.Picture = LoadPicture(LoadGasPicture(24907))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(24908))
   Load vgFP.picFixa(1)
   Set vgFP.picFixa(1).Picture = LoadPicture(LoadGasPicture(24909))
   Set lblAjuste = Label(8)
   vgFP.picFixa(0).AutoSize = True
   vgFP.picFixa(0).ScaleMode = vbMillimeters
   vgQueryAtual = "SELECT [Sequencia da Manutenção], Geral.[Sequencia do geral], [Razão social], [Fone], [Data de vencimento], " + _
                     "Documento, [Data de entrada], Parcela, [Valor da parcela], [Valor pago], [Valor do juros], [Valor do desconto], " + _
                     "[Valor restante], [Especie de pagamento], [Historico], [Descrição da Cobrança], [Manutenção Contas Receber].[Sequencia da Cobrança] " + _
                     "FROM [Manutenção contas receber] LEFT JOIN Geral ON [Manutenção contas receber].[Sequencia do geral] = " + _
                     "Geral.[Sequencia do geral] LEFT JOIN [Tipo da Cobrança] ON [Manutenção contas receber].[Sequencia da Cobrança] = " + _
                     "[Tipo da Cobrança].[Sequencia da Cobrança]"
   Set Manutencao_contas_receber = vgDb.OpenRecordSet("SELECT TOP 1" + Mid$(vgQueryAtual, 7))
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(33405)
   opcPainel1(0).Caption = LoadGasString(33406)
   opcPainel1(1).Caption = LoadGasString(33407)
   Label(1).Caption = LoadGasString(33408)
   opcPainel2(0).Caption = LoadGasString(33409)
   opcPainel2(1).Caption = LoadGasString(33410)
   Label(2).Caption = LoadGasString(33411)
   Label(3).Caption = LoadGasString(33412)
   vgTooltips.AddTool botConfImp, 0, LoadGasString(33413)
   Label(5).Caption = LoadGasString(33414)
   chkCampo(0).Caption = LoadGasString(33415)
   Label(6).Caption = LoadGasString(33416)
   Label(7).Caption = LoadGasString(33417)
   vgTooltips.AddTool botOk, 0, LoadGasString(33418)
   botOk.Caption = LoadGasString(33419)
   vgTooltips.AddTool botCancela, 0, LoadGasString(33420)
   botCancela.Caption = LoadGasString(33421)
   Label(8).Caption = LoadGasString(33422)
   Label(9).Caption = LoadGasString(33423)
   Label(10).Caption = LoadGasString(33424)
   Label(11).Caption = LoadGasString(33425)
   Label(12).Caption = LoadGasString(33426)
   Label(13).Caption = LoadGasString(33427)
   Label(14).Caption = LoadGasString(33428)
   Label(15).Caption = LoadGasString(33429)
   opcPainel0(0).Caption = LoadGasString(33430)
   opcPainel0(1).Caption = LoadGasString(33431)
   Label(16).Caption = LoadGasString(33432)
   opcPainel0(2).Caption = LoadGasString(33433)
   opcPainel0(3).Caption = LoadGasString(33434)
   chkCampo(1).Caption = LoadGasString(33435)
   vgTooltips.AddTool Frame(0), 0, LoadGasString(33436)
   Frame(0).Caption = LoadGasString(33437)
   vgTooltips.AddTool opcFrame0(0).CtPri, 0, LoadGasString(33438)
   opcFrame0(0).Caption = LoadGasString(33439)
   vgTooltips.AddTool opcFrame0(1).CtPri, 0, LoadGasString(33440)
   opcFrame0(1).Caption = LoadGasString(33441)
   vgSQL$ = Manutencao_contas_receber.Source
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaPreValidacao
   txtCampo(1).Value = 1
   Screen.MousePointer = vbDefault
   vgUltimoTabIndex = -1                          'controla lost focus de campos optativos
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
   txtCampo(3).DataType = 2
   txtCampo(3).Mask = "99/99/9999"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_LISTA) = bottxtCampo6(BOT_LISTA)
   Set txtCampo(6).CtBot(BOT_COMBO) = bottxtCampo6(BOT_COMBO)
   bottxtCampo6(BOT_LISTA).Caption = "P"
   Set bottxtCampo6(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = ""
   txtCampo(6).PesqModoAbertura =  2 
   txtCampo(6).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(6).BoundColumn = "Sequencia do geral"
   txtCampo(6).ListFields = "Razão social"
   txtCampo(6).OrderFields = "Razão social"
   txtCampo(6).Relation = ""
   txtCampo(6).Source = "Geral"
   txtCampo(6).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = ""
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 0
   txtCampo(5).Mask = ""
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set txtCampo(9).CtPri = txtCp(9)
   Set txtCampo(9).CtFdo = labFdo9
   Set txtCampo(9).CtBot(BOT_COMBO) = bottxtCampo9(BOT_COMBO)
   Set bottxtCampo9(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = ""
   txtCampo(9).BoundColumn = "Sequencia do geral"
   txtCampo(9).ListFields = "Razão social"
   txtCampo(9).OrderFields = "Razão social"
   txtCampo(9).Relation = ""
   txtCampo(9).Source = "Geral"
   txtCampo(9).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_COMBO) = bottxtCampo7(BOT_COMBO)
   Set bottxtCampo7(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(7).DataType = 0
   txtCampo(7).ListFields = "Dinheiro|Cartão|Cheque - Pre|Boleto|Condicional|Garantia|Deposito"

   Set txtCampo(10).CtPri = txtCp(10)
   Set txtCampo(10).CtFdo = labFdo10
   Set txtCampo(10).CtBot(BOT_COMBO) = bottxtCampo10(BOT_COMBO)
   Set bottxtCampo10(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = ""
   txtCampo(10).BoundColumn = "Sequencia da Cobrança"
   txtCampo(10).ListFields = "Descrição da Cobrança"
   txtCampo(10).OrderFields = "Descrição da Cobrança"
   txtCampo(10).Relation = ""
   txtCampo(10).Source = "Tipo da Cobrança"
   txtCampo(10).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 1
   txtCampo(8).Mask = "999999"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""

   Set txtCampo(1).CtPri = txtCp(1)
   Set txtCampo(1).CtFdo = labFdo1
   Set txtCampo(1).CtScr = ScrVtxtCampo1
   txtCampo(1).LargeChange = 10
   txtCampo(1).SmallChange = 1
   txtCampo(1).DataType = 3
   txtCampo(1).Mask = ""
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   txtCampo(1).Value = 1

   Set opcPainel1(0).CtPri = opcPainel1Cp(0)
   Set opcPainel1(0).CtFdo = labopcPainel1
   opcPainel1(0).DataType = 6
   opcPainel1(0).BookMark = 0

   Set opcPainel1(1).CtPri = opcPainel1Cp(1)
   Set opcPainel1(1).CtFdo = labopcPainel1
   opcPainel1(1).DataType = 6
   opcPainel1(1).BookMark = 1

   Set opcPainel2(0).CtPri = opcPainel2Cp(0)
   Set opcPainel2(0).CtFdo = labopcPainel2
   opcPainel2(0).DataType = 6
   opcPainel2(0).BookMark = 0

   Set opcPainel2(1).CtPri = opcPainel2Cp(1)
   Set opcPainel2(1).CtFdo = labopcPainel2
   opcPainel2(1).DataType = 6
   opcPainel2(1).BookMark = 1

   Set txtCampo(0).CtPri = txtCp(0)
   txtCampo(0).DataType = 0
   txtCampo(0).Mask = ""
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5

   Set opcPainel0(0).CtPri = opcPainel0Cp(0)
   Set opcPainel0(0).CtFdo = labopcPainel0
   opcPainel0(0).DataType = 6
   opcPainel0(0).BookMark = 0

   Set opcPainel0(1).CtPri = opcPainel0Cp(1)
   Set opcPainel0(1).CtFdo = labopcPainel0
   opcPainel0(1).DataType = 6
   opcPainel0(1).BookMark = 1

   Set opcPainel0(2).CtPri = opcPainel0Cp(2)
   Set opcPainel0(2).CtFdo = labopcPainel0
   opcPainel0(2).DataType = 6
   opcPainel0(2).BookMark = 2

   Set opcPainel0(3).CtPri = opcPainel0Cp(3)
   Set opcPainel0(3).CtFdo = labopcPainel0
   opcPainel0(3).DataType = 6
   opcPainel0(3).BookMark = 3

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5

   Set opcFrame0(0).CtPri = opcFrame0Cp(0)
   Set opcFrame0(0).CtFdo = labopcFrame0
   opcFrame0(0).DataType = 6
   opcFrame0(0).BookMark = 0

   Set opcFrame0(1).CtPri = opcFrame0Cp(1)
   Set opcFrame0(1).CtFdo = labopcFrame0
   opcFrame0(1).DataType = 6
   opcFrame0(1).BookMark = 1

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub

'evento - configura impressora
Private Sub botConfImp_Click()
   Dim x As String
   ConfigImpressora                               'chama diáglogo de configuração impressora
   x$ = Printer.DeviceName
   Label(4).Caption = x$
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
   vgFP.vgFimDoArq = False                        'é fim do arquivo
   vgFP.vgContaPg = 0                             'contador de paginas
   vgFP.vgNovaPag = 0                             'número da página real dentro do intervalo (Ex: 3;5;1;2 - se estiver na página real 5, a página atual será 2)
   vgFP.vgPagRef = 0                              'número da página de referência
   vgFP.vgUltPagVista = 0                         'ate onde ja fomos...
   If vgDestRel = 0 Then                          'onde vai ser gerado (janela ou papel)
      vgFP.cboEscala.ListIndex = 3                'escal 100%
      ProssegueExecucao vgFP.picAux               'faz preview
   Else
      vgFP.vgContaPg = 0
      ProssegueExecucao Printer                   'joga direto na impressora
   End If
   If vgPrimeiroCab <> 2 Then                     'se não deu erro
      If vgPrimeiroCab Then                       'nenhuma página foi impressa
         MsgBox LoadGasString(92), vbExclamation, vgAtencao$ 'apresenta a mensagem
      ElseIf vgDestRel = 0 Then
         vgFP.ShowPreview                                    'arruma form preview
         vgFP.Visible = True                                 'mostra form de preview
         Visible = False                                     'esconde o este form
         vgVisivel = False                                   'e seta variavel
      End If
   End If
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
   Unload vgFP.picFixa(1)
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   If Not Manutencao_contas_receber Is Nothing Then
      Manutencao_contas_receber.CloseRecordSet
      Set Manutencao_contas_receber = Nothing
   End If
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   SET vgFP.vgFormRel = Nothing                   'variável associada a este form
   UnLoad vgFP                                    'descarrega o form de preview
   Set vgFP = Nothing                             'tira o form de preview da memória
   Set frmRelReceb = Nothing                      'libera o segmento de código do form
   If vgRelAtivo = vgFormID Then vgRelAtivo = 0
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
Private Sub bottxtCampo6_Click(Index As Integer)
   txtCampo(6).SetFocus
   DoEvents
   txtCampo(6).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo9_Click(Index As Integer)
   txtCampo(9).SetFocus
   DoEvents
   txtCampo(9).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo7_Click(Index As Integer)
   txtCampo(7).SetFocus
   DoEvents
   txtCampo(7).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo10_Click(Index As Integer)
   txtCampo(10).SetFocus
   DoEvents
   txtCampo(10).botClick Index
End Sub

'evento - quando o valor da rolagem modificar
Private Sub scrVtxtCampo1_Change()
   txtCampo(1).scrChange
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub

