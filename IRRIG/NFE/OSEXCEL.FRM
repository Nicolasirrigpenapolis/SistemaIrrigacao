VERSION 5.00
Begin VB.Form frmOSExcel
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   4515
   Left                     =   -120
   LinkTopic                =   "OSExcel"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   4080
   ScaleWidth               =   12000
   Top                      =   2020
   Width                    =   12090
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
      Height                   =   3915
      Left                     =   15
      Width                    =   11820
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   20838
      _extenty                 =   6902
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
         Top                      =   3300
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
         Top                      =   3300
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
            TabIndex                 =   7
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
            TabIndex                 =   8
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
         Top                      =   3270
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
         Top                      =   3270
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
            TabIndex                 =   9
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
            TabIndex                 =   10
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
         TabIndex                 =   11
         TabStop                  =   0
         Height                   =   315
         Left                     =   3165
         Width                    =   465
         Top                      =   3270
         Index                    =   0 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   7905
         Top                      =   3270
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
         Top                      =   3285
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
         TabIndex                 =   12
         TabStop                  =   0
         Height                   =   285
         Left                     =   7920
         Width                    =   315
         Top                      =   3285
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
         Top                      =   3300
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
         Top                      =   3300
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
         Top                      =   3270
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
         TabIndex                 =   13
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   360
         Left                     =   11280
         Width                    =   360
         Top                      =   3225
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
         Y1                       =   3090
         X2                       =   11625
         Y2                       =   3090
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
         Left                     =   1590
         Width                    =   2820
         Top                      =   540
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
         Caption                  =   "O.S - Emitidas"
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
         TabIndex                 =   14
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
         Top                      =   1605
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
         Top                      =   1575
         Width                    =   1395
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   2430
         Top                      =   1590
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
         Top                      =   1590
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
         Top                      =   1605
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
         Top                      =   1575
         Width                    =   1335
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4725
         Top                      =   1590
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
         Width                    =   1035
         Top                      =   1590
         Index                    =   3 
      End
      Begin NFE.GBotao botOk
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   6
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   360
         Left                     =   8640
         Width                    =   1485
         Top                      =   2580
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
         TabIndex                 =   15
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   360
         Left                     =   10140
         Width                    =   1485
         Top                      =   2580
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
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6075
         Width                    =   5550
         Top                      =   1980
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5130
         Width                    =   825
         Top                      =   1980
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
         Caption                  =   "% Nome:"
         Index                    =   9 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5415
         Width                    =   540
         Top                      =   1560
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
         Index                    =   10 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6075
         Top                      =   1560
         Width                    =   5550
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   11355
         Top                      =   1575
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11100
         Top                      =   1575
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
         Left                     =   6090
         Width                    =   4995
         Top                      =   1575
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   270
         Left                     =   375
         Width                    =   1050
         Top                      =   2655
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Status:"
         Index                    =   11 
      End
      Begin          VB.Label labopcPainel0
         BorderStyle              =   1
         Height                   =   120
         Left                     =   2895
         Top                      =   180
         Visible                  =   0
         Width                    =   240
         End
      Begin VB.OptionButton opcPainel0Cp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   2
         TabIndex                 =   16
         TabStop                  =   0
         Height                   =   330
         Left                     =   1605
         Width                    =   1530
         Top                      =   2655
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   0
         Caption                  =   "Abertas"
         Index                    =   0 
      End
      Begin VB.OptionButton opcPainel0Cp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   2
         TabIndex                 =   17
         TabStop                  =   0
         Height                   =   330
         Left                     =   3225
         Width                    =   1650
         Top                      =   2655
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   0
         Caption                  =   "Fechadas"
         Index                    =   1 
      End
      Begin VB.OptionButton opcPainel0Cp
         Appearance               =   1
         BackColor                =   &H8000000F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   2
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   330
         Left                     =   5025
         Width                    =   1950
         Top                      =   2655
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Alignment                =   0
         Caption                  =   "Canceladas"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5010
         Width                    =   945
         Top                      =   1140
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
         Index                    =   12 
      End
      Begin VB.Label labFdo6
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6075
         Top                      =   1140
         Width                    =   5550
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   11355
         Top                      =   1155
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11100
         Top                      =   1155
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
         Left                     =   6090
         Width                    =   4995
         Top                      =   1155
         Index                    =   6 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4830
         Width                    =   1125
         Top                      =   720
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
         Caption                  =   "Ordenar por:"
         Index                    =   13 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6075
         Top                      =   720
         Width                    =   1620
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7425
         Top                      =   735
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
         TabStop                  =   0
         Height                   =   270
         Left                     =   6090
         Width                    =   1320
         Top                      =   735
         Index                    =   7 
      End
      Begin VB.Label labFdo8
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1305
         Top                      =   2175
         Width                    =   3690
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4725
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
         MaxLength                =   255
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   19
         TabStop                  =   0
         Height                   =   270
         Left                     =   1320
         Width                    =   3390
         Top                      =   2190
         Index                    =   8 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   420
         Width                    =   810
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
         Caption                  =   "Diretorio:"
         Index                    =   14 
      End
   End
End
Attribute VB_name = "frmOSExcel"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: OSExcel
'* Função....: O.S - Emitadas (Exportar Excel)
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:48
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

Public vgFP As New frmPreview                     'formuário de preview do relatório 
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFG As frmGauge                           'formulário do gauge 
Public vgQReg As Long                             'quantidade de já feira para o gauge  

Dim txtCampo(8) As New FormataCampos              'classe dos compos tipo texto do módulo  
Dim chkCampo(0) As New FormataCampos              'classe dos compos tipo lógico do módulo 
Dim opcPainel0(2) As New FormataCampos
Dim opcPainel1(1) As New FormataCampos
Dim opcPainel2(1) As New FormataCampos
Dim vgDestRel As Integer, vgModoInt As Integer, vgIntPag As String
Dim vgCopias As Integer, Zebrar As Boolean, Dt_Inicial As Variant
Dim Dt_Final As Variant, Nome As String, Geral As Double
Dim Status As Integer, Vendedor As Double, Ordenacao As String
Dim Diretorio As String
Public lblAjuste As Object, txtDiretorio As Object

Private Function Filtro() As String

  Filtro = "1 = 1"
   If Nome <> "" Then
      Filtro = Filtro & " AND Geral.[Razão social] Like '%" & Nome & "%'"
   End If
   If Not Vazio(Dt_Inicial) Then
      Filtro = Filtro & " AND [Data da Ordem] >= " & D(Dt_Inicial)
   End If
   If Not Vazio(Dt_Final) Then
      Filtro = Filtro & " AND [Data da Ordem] <= " & D(Dt_Final)
   End If
   If Geral > 0 Then
      Filtro = Filtro & " AND [Ordem de Serviço].[Sequencia do geral] = " & Geral
   End If
   
   If Vendedor > 0 Then
      Filtro = Filtro & " AND [Ordem de Serviço].[Sequencia do Vendedor] = " & Vendedor
   End If
   
   If Status = 0 Then ' Abertas
      Filtro = Filtro & " AND Fechado = 0 And Cancelado = 0"
   ElseIf Status = 1 Then
      Filtro = Filtro & " AND Fechado = 1 And Cancelado = 0" 'Fechadas
   Else
      Filtro = Filtro & " AND Cancelado = 1"    
   End If
    
   Filtro = Filtro & " AND [Ordem de Serviço].[Sequencia da Ordem] > 0 And [Serviços da Os].[Sequencia do Virabrequim] > 0"
   
   If Ordenacao <> "" Then 
    If Ordenacao = "OS" Then
       Filtro = Filtro & " ORDER BY [Ordem de Serviço].[Sequencia da Ordem]"
    ElseIf Ordenacao = "Emissão" Then 
       Filtro = Filtro & " ORDER BY [Ordem de Serviço].[Data da Ordem]"
    Else
       Filtro = Filtro & " ORDER BY Geral.[Razão social]"     
    End If
  End If
   
End Function


Private Sub ExportaExcel()
 Dim Tb As New GRecordSet, VIOS As New GRecordSet
 Dim Servicos As New GRecordSet
 Dim vArquivo As Long, Excel As Object, L As Long
 Dim NomeTrans As New GRecordSet
    
 Set Tb = vgDb.OpenRecordSet("SELECT Virabrequim.[Identificação] As Id, [Ordem de Serviço].[Sequencia da Ordem] As Seq, [Ordem de Serviço].[Sequencia do Transportador] As CodTrans, [Serviços da OS].[Valor Total] As Valor, " & _
                              "[Ordem de Serviço].[Data da Ordem] As Emi, Geral.[Razão Social] As Nome, [Serviços].[Descrição do Serviço] As DescrSer, Municipios.[Descrição do Municipio] As Mun, Municipios.UF, " & _
                              "Virabrequim.Modelo As Modelo_VI, Marcas.[Sequencia da Marca] As SeqMarca, Marcas.[Descrição] As Marca, V.[Razão Social] As Vendedor, [Ordem de Serviço].Obs As Observacao, Geral.[Endereço], Geral.Nro, [Modelo do Virabrequim].[Descrição do Modelo] As ModeloB, [Modelo do Virabrequim].[Linha do Modelo] As Linha " & _
                              "FROM [Virabrequim] Left Join Geral On Virabrequim.[Sequencia do geral] = Geral.[Sequencia do geral] " & _
                              "Left Join Municipios On Geral.[Sequencia do Municipio] = Municipios.[Sequencia do Municipio] " & _
                              "Left Join Marcas On Virabrequim.[Sequencia da Marca] = Marcas.[Sequencia da Marca] " & _
                              "Left Join [Modelo do Virabrequim] On Virabrequim.[Sequencia do Modelo] = [Modelo do Virabrequim].[Sequencia do Modelo] " & _
                              "Left Join [Serviços da OS] On Virabrequim.[Sequencia do Virabrequim] = [Serviços da OS].[Sequencia do Virabrequim] " & _
                              "Left Join [Serviços] On [Serviços da OS].[Sequencia do Serviço] = Serviços.[Sequencia do Serviço] " & _
                              "Left Join [Ordem de Serviço] On [Serviços da OS].[Sequencia da Ordem] = [Ordem de Serviço].[Sequencia da Ordem] " & _ 
                              "Left Join Geral v On [Ordem de Serviço].[Sequencia Do Vendedor] = v.[Sequencia Do geral] WHERE " & Filtro())                                                                  
 On Error Resume Next
      
If Tb.RecordCount > 0 Then
     vArquivo = FreeFile
     If Dir(txtDiretorio.text & "Emitidas.xls", vbArchive) <> "" Then Kill txtDiretorio.text & "Emitidas.xls"
        Open txtDiretorio.text & "Emitidas.xls" For Append As #vArquivo 
        Set Excel = CreateObject("Excel.Application")
      
        Set Excel = Excel.Workbooks.Add
        Set Excel = Excel.Sheets(1)
        Excel.SaveAs vArquivo
        Screen.MousePointer = vbHourglass 'Ampulheta
        'Status
        Me.Caption = "Exportando dados para Excel - Exportação " '& " " & Info(Tb)
        Excel.Application.Quit
        Set Excel = CreateObject("Excel.Application")
        With Excel
            .Workbooks.Open FileName:=vArquivo
            .Visible = False
            .Sheets("Plan1").Select
            .Range("A1").Select
            .activecell.Value = "Reticom Virabrequim"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 15
            .activecell.Font.Color = vbGreen
            
            .Range("A2").Select
            .activecell.Value = "Identificação"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("B2").Select
            .activecell.Value = "Nº O.S"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("C2").Select
            .activecell.Value = "Emissão"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("D2").Select
            .activecell.Value = "Proprietario"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("E2").Select
            .activecell.Value = "Municipio"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("F2").Select
            .activecell.Value = "UF"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("G2").Select
            .activecell.Value = "Endereço"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("H2").Select
            .activecell.Value = "Modelo"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("I2").Select
            .activecell.Value = "ModeloB"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("J2").Select
            .activecell.Value = "Linha"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("K2").Select
            .activecell.Value = "Seq"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("L2").Select
            .activecell.Value = "Marca"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("M2").Select
            .activecell.Value = "Serviço Executado"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("N2").Select
            .activecell.Value = "Valor"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("O2").Select
            .activecell.Value = "Vendedor"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            .Range("P2").Select
            .activecell.Value = "Transportador"
            .activecell.Font.Name = "Arial Black"
            .activecell.Font.Size = 10
            
            '.Range("L2").Select
            '.activecell.Value = "Observações"
            '.activecell.Font.Name = "Arial Black"
            '.activecell.Font.Size = 10
            
          L = 2     
          Do While Not Tb.EOF 
      'Identificação
      ' Set VIOS = VgDb.OpenRecordSet("SELECT * From [Virabrequim da Os] Where [Sequencia da Ordem] = " & OS![Sequencia da Ordem])
      ' Do While Not VIOS.EOF
          L = L + 1 
      ' Set Virabrequim = VgDb.OpenRecordSet("SELECT [Identificação] As ID From Virabrequim Where [Sequencia do Virabrequim] = " & VIOS![Sequencia do Virabrequim])  
          'Virabrequim       
            .Range("A" & Trim(str(L))).Select
            .activecell.Font.Name = "Courier New"
            .activecell.Font.Size = 10
            .activecell.Value = Tb!ID
            'Nº OS        
           .Range("B" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Seq
           'Emissão
           .Range("C" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Emi
           'Proprietario 
           .Range("D" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb!Nome)), True) 
           'Municipio
           .Range("E" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Mun
           'UF
           .Range("F" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!UF
           'Endereço
           .Range("G" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = RemoveCaracteresExcel(SuperTiraAcentos(Trim(Tb![Endereço])), True) & "," & Tb!Nro'Tb![Endereço]
           'Modelo
           .Range("H" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!ModeloB 
           'Modelo b
           .Range("I" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Modelo_VI
           
           'Linha
           .Range("J" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Linha
 
           'Seq Marca
           .Range("k" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!SeqMarca  
           'Marca
           .Range("L" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Marca
           
           .Range("M" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!DescrSer
           
           .Range("N" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Valor
                      
           .Range("O" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
           .activecell.Value = Tb!Vendedor
           
           .Range("P" & Trim(str(L))).Select
           .activecell.Font.Name = "Courier New"
           .activecell.Font.Size = 10
            Set NomeTrans = VgDb.OpenRecordSet("SELECT [Razão Social] Nome From Geral WHERE [Sequencia do Geral] = " & Tb!CodTrans)
           .activecell.Value = NomeTrans!Nome
           
           
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
    .Range("A1").ColumnWidth = 16  
    '.Range("A3").ColumnWidth = 16 
    .Range("B1").ColumnWidth = 12  
    .Range("B2").ColumnWidth = 7    
    .Range("C2").ColumnWidth = 12 
    .Range("D2").ColumnWidth = 72 
    .Range("E2").ColumnWidth = 30 
    .Range("F2").ColumnWidth = 7 
    .Range("G2").ColumnWidth = 40 'Endereço 
    .Range("H2").ColumnWidth = 15 'Modelo 
    .Range("I2").ColumnWidth = 15  'Seq
    .Range("J2").ColumnWidth = 8 
    .Range("K2").ColumnWidth = 10 
    .Range("L2").ColumnWidth = 20 
    .Range("M2").ColumnWidth = 60 
    .Range("N2").ColumnWidth = 15
    .Range("O2").ColumnWidth = 40
    .Range("P2").ColumnWidth = 40
    '.Range("L2").ColumnWidth = 90 ' Vendedor 
   
    .Columns("D:D").Select
    .Selection.Style = "Currency"  
    .Columns("M:M").Select
    .Selection.Style = "Currency"      
   
     End With  
    Else
    MsgBox "Não consta Lançamento para estes Filtros!", vbExclamation + vbOKOnly, "Atenção..."
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
   Geral = txtCampo(5).Value
   Status = Val(labopcPainel0.Caption)
   Vendedor = txtCampo(6).Value
   Ordenacao = txtCampo(7).Value
   Diretorio = txtCampo(8).Value
   If Err Then Err.Clear                          'se deu algum erro, vamos resetá-lo
End Sub

'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 5
         txtCampo(5).Filter = "[Sequencia do geral] > " & 0 & ""
                           txtCampo(5).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Complemento, " + _
                                                              "Geral.Nro, Geral.Bairro, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ")"
      Case 6
         txtCampo(6).Filter = "(([Sequencia do geral] > " & 0 & ") AND Inativo = " & 0 & ") AND Vendedor = " & 1 & ""
                           txtCampo(6).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.Vendedor = " + CStr(1) + ")"
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
      If vgModoInt = 1 Then
         vgOk = (Len(vgIntPag) > 0)
         vgMsg$ = "Esta informação é necessária"
         If Not vgOk Then txtCampo(0).SetFocus
      End If
      If vgOk Then
         If vgDestRel = 1 Then
            vgOk = (vgCopias > 0)
            vgMsg$ = "Número de cópias tem que ser maior que ZERO!"
            If Not vgOk Then txtCampo(1).SetFocus
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
      txtCampo(7).Value = "OS"
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
   Ok = (vgModoInt = 1)
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (vgDestRel = 1)
   txtCampo(1).Enabled = (Ok And txtCampo(1).Editable)
   Ok = (Geral = 0)
   txtCampo(4).Enabled = (Ok And txtCampo(4).Editable)
   Ok = (Nome = "")
   txtCampo(5).Enabled = (Ok And txtCampo(5).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(8).Visible = (False)
   Label(14).Visible = (False)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 5 Or Index = 6 Or Index = 7 Then
      ExecutaVisivel
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
               ExecutaVisivel
               ExecutaPreValidacao
            End If
         Case 2
            If Len(txtCp(2).Text) = 0 Then
               txtCampo(2).Value = DateSerial(Year(Date), Month(Date), 1)
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
            End If
         Case 3
            If Len(txtCp(3).Text) = 0 Then
               txtCampo(3).Value = UltD(Dt_Inicial)
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
            End If
         Case 5
            PoeRelEFiltroCbo 5
         Case 6
            PoeRelEFiltroCbo 6
         Case 7
            If Len(txtCp(7).Text) = 0 Then
               txtCampo(7).Value = "OS"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
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
   InicializaApelidos COM_TEXTBOX                 'inicializa apelidos dos campos
   ExecutaVisivel                                 'torna camos visiveis
   ExecutaPreValidacao                            'habilita/desabilita campos
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel                                 'torna camos visiveis
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
         ExecutaVisivel
         ExecutaPreValidacao
         opcPainel1(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
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
         Case 5
            PoeRelEFiltroCbo 5
         Case 6
            PoeRelEFiltroCbo 6
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
         ExecutaVisivel
         ExecutaPreValidacao
         opcPainel2(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
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
         Case 5
            PoeRelEFiltroCbo 5
         Case 6
            PoeRelEFiltroCbo 6
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
         ExecutaVisivel
         ExecutaPreValidacao
         opcPainel0(Index).Change
      End If
   End If
   opcPainel0(0).CtPri.TabStop = False
   opcPainel0(1).CtPri.TabStop = False
   opcPainel0(2).CtPri.TabStop = False
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
         Case 5
            PoeRelEFiltroCbo 5
         Case 6
            PoeRelEFiltroCbo 6
      End Select
End Sub

'evento - quando o campo perder o foco
Private Sub opcPainel0Cp_LostFocus(Index As Integer)
   opcPainel0(Index).LostFocus
End Sub

'evento - inicialização do formulário
Private Sub Form_Load()
   On Error GoTo DeuErro
   Screen.MousePointer = vbHourglass
   Caption = LoadGasString(29250)
   vgFormID = 713
   vgIdentTab = ""
   vgPriVez = True
   vgTipo = TP_COMUM
   vgCaracteristica = F_COMUM
   vgTemFiltro = False
   vgSituacao = ACAO_NAVEGANDO
   Set botConfImp.Picture = LoadPicture(LoadGasPicture(20755))
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(20756))
   Set botOk.Picture = LoadPicture(LoadGasPicture(20757))
   Set botCancela.Picture = LoadPicture(LoadGasPicture(20758))
   Set lblAjuste = Label(8)
   Set txtDiretorio = txtCampo(8)
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(29255)
   opcPainel1(0).Caption = LoadGasString(29256)
   opcPainel1(1).Caption = LoadGasString(29257)
   Label(1).Caption = LoadGasString(29258)
   opcPainel2(0).Caption = LoadGasString(29259)
   opcPainel2(1).Caption = LoadGasString(29260)
   Label(2).Caption = LoadGasString(29261)
   Label(3).Caption = LoadGasString(29262)
   vgTooltips.AddTool botConfImp, 0, LoadGasString(29263)
   Label(5).Caption = LoadGasString(29264)
   chkCampo(0).Caption = LoadGasString(29265)
   Label(6).Caption = LoadGasString(29266)
   Label(7).Caption = LoadGasString(29267)
   vgTooltips.AddTool botOk, 0, LoadGasString(29268)
   botOk.Caption = LoadGasString(29269)
   vgTooltips.AddTool botCancela, 0, LoadGasString(29270)
   botCancela.Caption = LoadGasString(29271)
   Label(8).Caption = LoadGasString(29272)
   Label(9).Caption = LoadGasString(29273)
   Label(10).Caption = LoadGasString(29274)
   Label(11).Caption = LoadGasString(29275)
   opcPainel0(0).Caption = LoadGasString(29276)
   opcPainel0(1).Caption = LoadGasString(29277)
   opcPainel0(2).Caption = LoadGasString(29278)
   Label(12).Caption = LoadGasString(29279)
   Label(13).Caption = LoadGasString(29280)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(29281)
   Label(14).Caption = LoadGasString(29282)
   AjustaTamanho Me
   Executar INICIALIZACOES
   InicializaApelidos COM_TEXTBOX
   ExecutaVisivel
   ExecutaPreValidacao
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

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_COMBO) = bottxtCampo7(BOT_COMBO)
   Set bottxtCampo7(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(7).DataType = 0
   txtCampo(7).ListFields = "OS|Emissão|Cliente"

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

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_LISTA) = bottxtCampo5(BOT_LISTA)
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   bottxtCampo5(BOT_LISTA).Caption = "P"
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = ""
   txtCampo(5).PesqModoAbertura =  2 
   txtCampo(5).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(5).BoundColumn = "Sequencia do geral"
   txtCampo(5).ListFields = "Razão social"
   txtCampo(5).OrderFields = "Razão social"
   txtCampo(5).Relation = ""
   txtCampo(5).Source = "Geral"
   txtCampo(5).vgfrmGMCale.grdListaG.AutoRebind = True

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = ""
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""

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

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_ACAO) = bottxtCampo8(BOT_ACAO)
   Set bottxtCampo8(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(8).DataType = 0
   txtCampo(8).Mask = "@x"
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""

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
   ExportaExcel  
   vgPriVez = False
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
   Set vgFG = Nothing
   Set vgTooltips = Nothing
   Set frmOSExcel = Nothing                       'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no formulário
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then                 'se teclou ESC
      Unload Me                                   'tira este form da memória
   End If
End Sub

'evento - quando o botão for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 8
         PegaDir2 Me: txtDiretorio.Text = vaDir
   End Select
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
Private Sub bottxtCampo7_Click(Index As Integer)
   txtCampo(7).SetFocus
   DoEvents
   txtCampo(7).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo6_Click(Index As Integer)
   txtCampo(6).SetFocus
   DoEvents
   txtCampo(6).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).botClick Index
End Sub

'evento - quando o valor da rolagem modificar
Private Sub scrVtxtCampo1_Change()
   txtCampo(1).scrChange
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

