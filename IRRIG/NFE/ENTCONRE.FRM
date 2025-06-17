Version 5.00
Begin VB.Form frmEntConRe
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   7500
   Left                     =   2775
   LinkTopic                =   "EntConRe"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7065
   ScaleWidth               =   11325
   Top                      =   2080
   Width                    =   11415
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
      Height                   =   6900
      Left                     =   15
      Width                    =   11145
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   19648
      _extenty                 =   12164
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1605
         Left                     =   60
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
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   375
         Left                     =   1755
         Width                    =   4665
         Top                      =   675
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
         Caption                  =   "Entrada do Contas a Receber"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   1770
         Y1                       =   1005
         X2                       =   6465
         Y2                       =   1005
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   180
         Width                    =   1020
         Top                      =   1800
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
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   -1
         Height                   =   300
         Left                     =   1320
         Width                    =   720
         Top                      =   1800
         DataField                =   "Sequencia da entrada"
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
         Left                     =   2175
         Width                    =   1035
         Top                      =   1800
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
         Index                    =   2 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3300
         Top                      =   1800
         Width                    =   1350
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4380
         Top                      =   1815
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
         Left                     =   3315
         Width                    =   1050
         Top                      =   1815
         DataField                =   "Data de entrada"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   10680
         Width                    =   45
         Top                      =   7200
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
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4815
         Width                    =   420
         Top                      =   1800
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
         Caption                  =   "Doc:"
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
         MaxLength                =   25
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   5340
         Width                    =   1575
         Top                      =   1800
         DataField                =   "Documento"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7035
         Width                    =   1470
         Top                      =   1800
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
         Caption                  =   "Centro de Custo:"
         Index                    =   5 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8640
         Top                      =   1800
         Width                    =   2220
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   10590
         Top                      =   1815
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   10335
         Top                      =   1815
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
         Left                     =   8655
         Width                    =   1665
         Top                      =   1815
         DataField                =   "Tipo de conta"
         Index                    =   3 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HE0E0E0
         Left                     =   1320
         Top                      =   2415
         Width                    =   5340
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   6390
         Top                      =   2430
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6135
         Top                      =   2430
         Width                    =   255
         Height                   =   270
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
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   270
         Left                     =   1335
         Width                    =   4785
         Top                      =   2430
         DataField                =   "Sequencia do geral"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   495
         Width                    =   720
         Top                      =   2415
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
         Caption                  =   "Geral....:"
         Index                    =   6 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   5
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   293
         Left                     =   6675
         Width                    =   240
         Top                      =   2415
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
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7575
         Width                    =   930
         Top                      =   2415
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
         Caption                  =   "F. Pagto...:"
         Index                    =   7 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8640
         Top                      =   2415
         Width                    =   2220
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   10590
         Top                      =   2430
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
         Left                     =   8655
         Width                    =   1920
         Top                      =   2430
         DataField                =   "Forma de pagamento"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   360
         Left                     =   180
         Width                    =   5565
         Top                      =   3015
         Alignment                =   2
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &HFFFFFF
         Caption                  =   "Historico"
         Index                    =   8 
      End
      Begin VB.TextBox txtCp
         Appearance               =   1
         BackColor                =   &H80000005
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
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   7
         TabStop                  =   0
         Height                   =   3285
         Left                     =   180
         Width                    =   5565
         Top                      =   3420
         DataField                =   "Historico"
         Index                    =   6 
         MultiLine                =   -1
         ScrollBars               =   2
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &H0
         BackStyle                =   1
         BorderStyle              =   1
         AutoSize                 =   0
         Height                   =   360
         Left                     =   5775
         Width                    =   5130
         Top                      =   3015
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &HFFFFFF
         Caption                  =   "Total á Pagar  --->"
         Index                    =   9 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80FFFF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         BorderStyle              =   0
         Alignment                =   1
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   330
         Left                     =   9165
         Width                    =   1680
         Top                      =   3025
         DataField                =   "Valor total"
         Index                    =   7 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5955
         Width                    =   1410
         Top                      =   6555
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
         Caption                  =   "Status Parcelas"
         Index                    =   10 
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
         Left                     =   7530
         Width                    =   165
         Top                      =   6600
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   7755
         Width                    =   990
         Top                      =   6555
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
         Index                    =   11 
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
         Left                     =   8850
         Width                    =   165
         Top                      =   6600
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   9150
         Width                    =   1050
         Top                      =   6555
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
         Index                    =   12 
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
         Left                     =   10365
         Width                    =   150
         Top                      =   6600
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   10590
         Width                    =   300
         Top                      =   6555
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
         Index                    =   13 
      End
      Begin NFE.GListV Grid
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
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
         TabIndex                 =   9
         TabStop                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   3045
         Left                     =   5775
         Width                    =   5130
         Top                      =   3420
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
         Height                   =   375
         Left                     =   1755
         Width                    =   105
         Top                      =   1035
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
         Caption                  =   ""
         Index                    =   14 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   10
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   6375
         Width                    =   1470
         Top                      =   135
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
         ForeColor                =   &H80000012
         Caption                  =   "Fechar"
         Index                    =   1 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   11
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   7875
         Width                    =   1470
         Top                      =   135
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
         ForeColor                =   &H80000012
         Caption                  =   "Abrir"
         Index                    =   2 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   12
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   9345
         Width                    =   1470
         Top                      =   135
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
         ForeColor                =   &H80000012
         Caption                  =   "Cancelar"
         Index                    =   3 
      End
   End
         Begin NFE.GListV grdBrowse 
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
         Begin VB.Timer timUnLoad
                  Enabled         =   0
                  Interval        =   5
                  left            =   30
                  top             =   950
         End
End
Attribute VB_name = "frmEntConRe"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: EntConRe
'* Função....: Entrada Contas
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:52
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
Public vgFiltroEmUso As Integer                   'Indice do Filtro atual em uso 
Public vgIndDefault As String                     'indice default do módulo 
Public vgFormID As Long                           'identificador único para o módulo 
Public vgIdentTab As String                       'nome da tabela principal do módulo 
Public vgFrmImpCons As New frmImpCons             'impressao de consutlas
Public vgTooltips As New cTooltips                'classe de ajuda para os controes do módulo
Public vgFiltroOriginal As String
Dim txtCampo(7) As New FormataCampos              'classe dos campos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Sequencia_da_entrada As Long, Data_de_entrada As Variant, Sequencia_do_geral As Long
Dim Documento As String, Historico As String, Forma_de_pagamento As String
Dim Valor_total As Double, Data_da_Alteracao As Variant, Hora_da_Alteracao As Variant
Dim Usuario_da_Alteracao As String, Tipo_de_conta As String, Fechado As Boolean
Dim Cancelado As Boolean
Public lblTitulo As Object, lblAjuste As Object, lblParcelamento As Object
Public grdParcelamento As Object, lblFechamento As Object
Dim Entrada_Contas_Receber As New GRecordSet, Parcelas_Entradas_Receber As New GRecordSet

Private Function PegaSequencial() As Long
 Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
  Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia da Entrada]) Seq From [Entrada Contas Receber]")
  PegaSequencial = Tb!Seq + 1
 End If
    
End Function

Private Function ValidaGeral() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Sequencia_do_geral)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaGeral = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaGeral = True: Exit Function      
   End If
  End If
  
End Function

Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas Entradas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_entrada & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function

Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC FROM [Parcelas Entradas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_entrada)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Sub Alteracao()
   On error Resume Next
      
   With vgTb   
      .Edit
      ![Data da Alteração] = Date
      ![Hora da Alteração] = Time
      ![Usuario da Alteração] = vgPwUsuario
      .Update
      .Bookmark = .lastModified
   End With
   
End Sub

Private Sub MostraFormulas()

   If Fechado Then
      lblFechamento.Caption = "Status -> Conta Fechada"
      lblFechamento.ForeColor = &H0
      lblTitulo.ForeColor = &H0
   Else
      lblFechamento.Caption = "Status -> Conta Aberta"
      lblFechamento.ForeColor = &H8FF
      lblTitulo.ForeColor = &H0
   End If
   If Err Then Err = 0: lblFechamento.Caption = ""   
   If Err Then Err.Clear 
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_total) Or (TotalParcelas > Valor_total) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   
   If Cancelado Then
      lblFechamento.Caption = "Status -> Conta Cancelada"
      lblFechamento.ForeColor = &H8FF
      lblTitulo.ForeColor = &H8FF
   End If
   
End Sub


Private Sub RepositionContas()
 If VgSituacao = ACAO_INCLUINDO Then
    Set Parcelas_Entradas_Receber = Nothing 
 End If
 CarregaTotalizador
 'Alteração
 lblAlteracao.Caption = Entrada_Contas_Receber![Usuario da Alteração] & " " & Entrada_Contas_Receber![Data da Alteração] & " " & Entrada_Contas_Receber![Hora da Alteração]
End Sub


Private Sub CarregaTotalizador()
   On Error Resume Next
    
   With grdParcelamento
      .ShowSumBar = True
      .ShowSumCol(.Columns("Numero da Parcela").Index) = False
      .ShowSumCol(.Columns("Dias").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
End Sub


Private Sub DeletaItens()' Qdo excluir Registro Mae
Dim Parcelas As New GRecordSet

  If vgSITUACAO = ACAO_EXCLUINDO Then
  Set Parcelas = vgDb.OpenRecordSet("SELECT * From [Parcelas Entradas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_Entrada)
  
   If Parcelas.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Parcelas Entradas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_Entrada
      DeletaConta
   End If
  End If 
    
End Sub


Private Sub DeletaConta() 'Qdo exclui registro Mãe
   Dim Manutencao As New GRecordset

   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada)
    
   If Manutencao.RecordCount = 0 Then Exit Sub
   
   vgDb.Execute "DELETE FROM [Manutenção contas Receber] " & _
                "WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada   

DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbokOnly, vaTitulo
      vgdb.RollBackTrans
   End If

End Sub

Private Sub AVISTA()
Dim Parcelas As New GRecordSet

TbAuxiliar "Parcelas Entradas Receber", "[Sequencia da Entrada] = " & Sequencia_da_entrada, Parcelas
 
 If (Entrada_Contas_Receber![Forma de pagamento] <> "Vista" Or Entrada_Contas_Receber![Valor total] = 0) Then Exit Sub 
    With Parcelas   
      If Parcelas.RecordCount = 0 Then .AddNew
      ![Sequencia da Entrada] = Entrada_Contas_Receber![Sequencia da entrada]
      ![Numero da Parcela] = 1
      !Dias = 0
      ![Data de Vencimento] = Entrada_Contas_Receber![Data de entrada]
      ![Valor da Parcela] = Entrada_Contas_Receber![Valor total]
      .Update
      .Bookmark = .lastModified
   End With
   GrdParcelamento.Rebind
   
End Sub


Private Sub AlteraManutencao() 'Qdo altera Registro Mãe
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet
   
   On Error GoTo DeuErro
             
   Set Parcelas = Parcelas_Entradas_Receber
   
   If Parcelas.RecordCount = 0 Then Exit Sub
         
   vgDb.BeginTrans
   
   Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção contas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_entrada)
   If Manutencao.RecordCount = 0 Then Exit Sub
   Do While Not Manutencao.EOF
      With Manutencao
         ![Sequencia Do geral] = Entrada_Contas_Receber![Sequencia Do geral]
         ![Data de entrada] = Entrada_Contas_Receber![Data de entrada]
         !Historico = Historico
         ![Forma de pagamento] = Entrada_Contas_Receber![Forma de pagamento]
         ![Valor total] = Entrada_Contas_Receber![Valor total]
         !Documento = Entrada_Contas_Receber!Documento
         '![Tipo de conta] = Entrada_Contas_Receber![Tipo de conta]
         .Update
         .BookMark = .LastModified
      End With
      Manutencao.MoveNext
   Loop    
   vgDb.CommitTrans
   Alteracao
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
   
End Sub


Private Sub Fechamento()
 Dim Parcelas As New GRecordSet
 Dim Manutencao As New GRecordSet
 
 On Error Resume Next
 
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas Entradas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_Entrada)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada)
  
  If Parcelas.RecordCount > 0 Then
   If TotalParcelas() <> Valor_total Then 
      MsgBox("Impossivel Fechar a Conta!! A Soma das Parcelas não Correspondem ao Valor Total da Conta")
   Exit Sub
   End If
  End If
  
  If Parcelas.RecordCount > 0 And Manutencao.RecordCount = 0 Then
     LancaParcelas 
  Else
     DeletaConta
     LancaParcelas     
  End If
  
  If Parcelas.RecordCount = 0 Then
      MsgBox("Impossivel Fechar a Conta!! Inclua a(s) Parcelas")
      Exit Sub
  End If
  
  vgDb.Execute "UPDATE [Entrada Contas Receber] Set Fechado = 1 WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada
  Reposition True
              
End Sub


Private Sub LancaParcelas()
  Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
  Dim SeqBaixa As Long
   
   On Error GoTo DeuErro
      
   If Vazio(Forma_de_pagamento) Then Exit Sub
   If TotalParcelas() <> Valor_total Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas Receber")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Entradas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_Entrada)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas Receber", "[Sequencia da Entrada] = " & Sequencia_da_Entrada & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew
            !Parcela = Parcelas![Numero da parcela]
            ![Sequencia Do geral] = Sequencia_do_Geral
            !Documento = Documento
            ![Tipo de conta] = Tipo_de_Conta
            ![Data de entrada] = Data_de_entrada
            !Historico = Historico
            ![Forma de Pagamento] = Forma_de_pagamento
             If ![Forma de pagamento] = "Vista" And Empresa![Baixa Receber Auto] Then
                ![Valor pago] = Parcelas![Valor da parcela]
                ![Valor restante] = 0
                ![Data da baixa] = Data_de_entrada
             Else
                ![Valor pago] = 0
                ![Valor restante] = Parcelas![Valor da parcela]
                ![Data da baixa] = Null
             End If
            ![Data de vencimento] = Parcelas![Data de vencimento]
            ![Valor total] = Valor_total
            ![Valor da parcela] = Parcelas![Valor da parcela]
            ![Sequencia da Entrada] = Sequencia_da_Entrada
          .Update
          .BookMark = .LastModified
      End With
      Parcelas.MoveNext
   Loop
   
   'Baixa
   If Forma_de_pagamento = "Vista" And Empresa![Baixa Receber Auto] Then
      Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas Receber]")
      Manutencao.MoveFirst: Manutencao.MoveLast
        With Baixa
           .AddNew
            ![Sequencia da Manutenção] = Manutencao![Sequencia da Manutenção]
            ![Data da Baixa] = Manutencao![Data de Vencimento]
            ![Valor Pago] = Manutencao![Valor da Parcela]       
           .Update
           .BookMark = .LastModified
           Baixa.MoveLast
       End With
       With Manutencao
           .Edit
            ![Sequencia da Baixa] = Baixa![Sequencia da Baixa]
             SeqBaixa = Baixa![Sequencia da Baixa]
           .Update
           .BookMark = .LastModified
       End With
       'MsgBox "PAGAMENTO A VISTA - (Informe o Caixa ou a Conta Corrente!)", vbInformation + vbOKOnly, vaTitulo
        frmBxConRec.Show
        frmBxConRec.SetFocus
        frmBxConRec.vgFiltroInicial = "[Sequencia da Baixa] = " & SeqBaixa
        InicializaFiltro frmBxConRec
        frmBxConRec.Reposition
        'Unload ME            
    End If
   
   vgDb.CommitTrans

DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Sub AbreConta()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean
 Dim Manu As New GRecordSet

 On Error Resume Next
 
 Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas Receber] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada & ")")
 Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas Receber] Where [Sequencia da Entrada] = " & Sequencia_da_Entrada)
 
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If
   
   If TemBaixa Then
      MsgBox("Impossivel abrir a Conta existe Baixa(s) no Contas a Receber para Abri-la Exclua a(s) Baixa(s)")
      Exit Sub
   End If
   

InputSenha
  If Not Vazio(InputYgorsBom) Then
     If InputYgorsBom = Empresa![Senha Admin] Then
        vgDb.Execute "UPDATE [Entrada Contas Receber] Set Fechado = 0 Where [Sequencia da Entrada] = " & Sequencia_da_Entrada
        If Manu.RecordCount > 0 Then
           vgDb.Execute "DELETE FROM [Manutenção contas Receber] WHERE [Sequencia da Entrada] = " & Sequencia_da_Entrada
        End If          
       Reposition True 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  InputYgorsBom = ""
  
End Sub  


Private Sub Cancelamento() 
 Dim Tb As New GRecordSet
   
   On Error GoTo DeuErro
      
   If MsgBox("Deseja Realmente Cancelar essa Conta - As Baixas do Contas a Receber e a Conta Tbm serão Canceladas?" & vbCrLf & "", vbExclamation + vbYesNo, vaTitulo) = vbNo Then Exit Sub
    
   Set Tb = vgDb.OpenRecordSet("SELECT Mvto.[Sequencia da Entrada], " & _
                               "Manu.[Sequencia da manutenção] Conta " & _
                               "FROM [Entrada Contas Receber] Mvto " & _
                               "LEFT JOIN [Manutenção contas Receber] Manu ON Mvto.[Sequencia da Entrada] = Manu.[Sequencia da Entrada] " & _
                               "WHERE Mvto.[Sequencia da Entrada] = " & Sequencia_da_entrada)
                                
   vgDb.BeginTrans 'Inicia a transação
        
   InputSenha
     If Not Vazio(InputYgorsBom) Then
        If InputYgorsBom = Empresa![Senha Admin] Then
        Goto Cancelar
     Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
 
  Cancelar:
  vgDb.Execute "Update [Entrada Contas Receber] Set [Cancelado] = 1 Where [Sequencia da Entrada] = " & Sequencia_da_entrada
    
    If Tb!Conta > 0 AND Not Vazio(Tb!Conta) Then
       Do While Not Tb.EOF 'Excluindo o Contas se houver
          vgDb.Execute "DELETE FROM [Baixa Contas Receber] WHERE [Sequencia da Manutenção] = " & Tb!Conta 'Baixa
          'vgDb.Execute "DELETE FROM [Movimentação da Conta Corrente] WHERE [Seqüência do Lançamento] = " & Tb!Conta 'Conta Corrente
         Tb.MoveNext
       Loop
       Tb.MoveFirst 'Vamos Voltar no Começo
       vgDb.Execute "DELETE FROM [Manutenção Contas Receber] WHERE [Sequencia da Entrada] = " & Tb![Sequencia da Entrada] 'Deleta a Conta gerado pelo Movimento
    End If   
    
    vgDb.CommitTrans 'Vomitando opsss Comitando rsrs
   
    MsgBox "Conta Cancelada com Sucesso!!!", vbInformation + vbOKOnly, vaTitulo 'Suuuuuceeeeesssssoooo
    
    Reposition True
    InputYgorsBom = ""
        
DeuErro:
   If Err.Number <> 0 Then
      vgDb.RollBackTrans
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
   
End Sub

Public Sub LigaDesligaBotoes
   Botao( 0 ).Enabled =    Botao( 0 ).Enabled And Permitido("Geral", ACAO_NAVEGANDO)
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Sequencia_da_entrada = txtCampo(0).Value
         Data_de_entrada = txtCampo(1).Value
         Sequencia_do_geral = txtCampo(4).Value
         Documento = txtCampo(2).Value
         Historico = txtCampo(6).Value
         Forma_de_pagamento = txtCampo(5).Value
         Valor_total = txtCampo(7).Value
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Usuario_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Usuario da Alteração])
         Tipo_de_conta = txtCampo(3).Value
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
      Else
         Sequencia_da_entrada = IIf(IsNull(vgTb![Sequencia da entrada]), 0, vgTb![Sequencia da entrada])
         Data_de_entrada = vgTb![Data de entrada]
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Documento = IIf(IsNull(vgTb!Documento), "", vgTb!Documento)
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Valor_total = IIf(IsNull(vgTb![Valor total]), 0, vgTb![Valor total])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Tipo_de_conta = IIf(IsNull(vgTb![Tipo de conta]), "", vgTb![Tipo de conta])
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
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
      grdBrowse.AllowInsert = i
      .botInclui.Enabled = i
      .Menu_Inclui.Enabled = i
      i = (Fechado = 0 And vgPWUsuario = "ADMIN") and Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = (Fechado = 0 And Cancelado = 0) and Permitido(vgIdentTab, ACAO_EDITANDO)
      If Err Then Err.Clear: i = False
      vgTemAlteracao = i
      grdBrowse.AllowEdit = i
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
      vgOk = ((Not Vazio(Data_de_entrada)) And (IsDate(Data_de_entrada) Or Vazio(Data_de_entrada)))
      vgMsg$ = "Data de entrada não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = (Not Vazio(Documento))
         vgMsg$ = "Documento não pode ser vazio!"
         If Not vgOk Then vgColumn = 3
      End If
      If vgOk Then
         vgOk = (Not Vazio(Tipo_de_conta))
         vgMsg$ = "Tipo de conta não pode ser vazio!"
         If Not vgOk Then vgColumn = 4
      End If
      If vgOk Then
         If Sequencia_do_geral > 0 Then
            vgOk = (ValidaGeral())
            vgMsg$ = "Geral Inativo!"
            If Not vgOk Then vgColumn = 5
         End If
      End If
      If vgOk Then
         If Sequencia_do_geral = 0 Then
            vgOk = (Sequencia_do_geral > 0)
            vgMsg$ = "Geral é Obrigatorio!"
            If Not vgOk Then vgColumn = 5
         End If
      End If
      If vgOk Then
         vgOk = (Not Vazio(Forma_de_pagamento))
         vgMsg$ = "Forma de pagamento não pode ser vazio!"
         If Not vgOk Then vgColumn = 6
      End If
      If vgOk Then
         vgOk = (Valor_total > 0)
         vgMsg$ = "Valor total inválido!"
         If Not vgOk Then vgColumn = 8
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
                  txtCampo(0).Value = PegaSequencial()
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 3
         PoeRelEFiltroCbo 4
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 3
         PoeRelEFiltroCbo 4
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
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
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
      Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
      Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_INCLUINDO Then
         Alteracao:AVISTA
      ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
         AlteraManutencao:AVISTA
      ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
         DeletaItens
      End If
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
   vgDataDeleted = mdiNFE.ExcluiRegistro()
End Sub
   
Private Sub grdBrowse_InitEdit(CancelEdit As Boolean)
   Reposition
End Sub

Private Sub grdBrowse_ItemSelect(ByVal vgItem As Long, vgColumns() As Variant)
   If vgPriVez Or Not grdBrowse.Visible Then Exit Sub
   AjustaRolagem Me
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
   mdiNFE.SalvaDados vgColumn
   vgDataSaved = (vgSituacao = ACAO_NAVEGANDO)
End Sub
   
Private Sub grdBrowse_StatusChanged(ByVal vgNewStatus As Integer)
   If (vgNewStatus = ACAO_EXCLUINDO And Val(grdBrowse.Recordset.BookMark) >= 0) Then
      Reposition
   End If
   PrepBotoes Me, vgNewStatus                          'acerta icones dos botoes
   mdiNFE.RemontaForm                                  'remonta dos os form da tela
End Sub

'apresenta popup menu para trabalhar com o grid
Private Sub grdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single, ByVal vgCurCol As Integer)
   MostraPopGrid Me, Button
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
   Set Entrada_Contas_Receber = vgTb
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
      x$ = ExecutaGrid(0, vgCols(), CONDICOES_ESPECIAIS)
   RepositionContas  
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
   vgTemAlteracaoGrids = Not LockGrids
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (Fechado = False And Cancelado = False And VgSituacao = ACAO_NAVEGANDO)
   Botao(1).Enabled = Ok
   Ok = (Fechado = True And Cancelado = False And VgSituacao = ACAO_NAVEGANDO)
   Botao(2).Enabled = Ok
   Ok = (VgSituacao = ACAO_NAVEGANDO And Not Cancelado)
   Botao(3).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(14).Visible = (Sequencia_da_entrada > 0)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 3 Or Index = 4 Or Index = 5 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 1
            If Len(txtCp(1).Text) = 0 Then
               txtCampo(1).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 3
            PoeRelEFiltroCbo 3
         Case 4
            PoeRelEFiltroCbo 4
      End Select
   End If
   txtCampo(Index).GotFocus
End Sub

'evento - quando qq tecla for digitada no campo
Private Sub txtCp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn And vgSituacao <> ACAO_NAVEGANDO Then  'se tela esta em edição e digitou ENTER
      ExecutaVisivel                                               'torna camos visiveis
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
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
   End If
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
   Caption = LoadGasString(32030)
   vgFormID = 736
   vgIdentTab = "Entrada Contas Receber"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia da entrada"
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 8
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(23535))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(23536))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(23537))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(23538))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(23539))
   Set lblTitulo = Label(0)
   Set lblAjuste = Label(3)
   Set lblParcelamento = Label(9)
   Set grdParcelamento = Grid(0)
   Set lblFechamento = Label(14)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia da entrada] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(32035)
   Label(1).Caption = LoadGasString(32036)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(32037)
   Label(2).Caption = LoadGasString(32038)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(32039)
   Label(3).Caption = LoadGasString(32040)
   Label(4).Caption = LoadGasString(32041)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(32042)
   Label(5).Caption = LoadGasString(32043)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(32044)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(32045)
   Label(6).Caption = LoadGasString(32046)
   Label(7).Caption = LoadGasString(32047)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(32048)
   Label(8).Caption = LoadGasString(32049)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(32050)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(32051)
   Label(10).Caption = LoadGasString(32052)
   Label(11).Caption = LoadGasString(32053)
   Label(12).Caption = LoadGasString(32054)
   Label(13).Caption = LoadGasString(32055)
   Botao(1).Caption = LoadGasString(32056)
   Botao(2).Caption = LoadGasString(32057)
   Botao(3).Caption = LoadGasString(32058)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Parcelas Entradas Receber", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas Entradas Receber", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas Entradas Receber", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 795
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 630
      .AddColumn Nothing, , "Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1650
      .AddColumn Nothing, , "Valor da Parcela", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1320
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Entrada_Contas_Receber = vgTb
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
 grdBrowse.AddControlIgnoreFocus mdiNFE.botCancela.hwnd           'não deixa o grid tentar gravar automaticamente
 grdBrowse.AddControlIgnoreFocus mdiNFE.botSalva.hwnd             'se estiver perdendo o foco para esses botões
   grdBrowse.AllowDelete = True
   grdBrowse.AllowEdit = True
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
   Set txtCampo(1).CtBot(BOT_ACAO) = bottxtCampo1(BOT_ACAO)
   Set bottxtCampo1(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(1).DataType = 2
   txtCampo(1).Mask = "99/99/9999"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 0
   txtCampo(2).Mask = "@!"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_LISTA) = bottxtCampo3(BOT_LISTA)
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   bottxtCampo3(BOT_LISTA).Caption = "P"
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@x"
   txtCampo(3).PesqModoAbertura =  2 
   txtCampo(3).PesqFieldCapture = "[Centro de Custo].[Descrição do Centro]"
   txtCampo(3).BoundColumn = "Descrição do Centro"
   txtCampo(3).ListFields = "Descrição do Centro"
   txtCampo(3).OrderFields = "Descrição do Centro"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Centro de Custo"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_LISTA) = bottxtCampo4(BOT_LISTA)
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   bottxtCampo4(BOT_LISTA).Caption = "P"
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "999999"
   txtCampo(4).PesqModoAbertura =  2 
   txtCampo(4).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(4).BoundColumn = "Sequencia do geral"
   txtCampo(4).ListFields = "Razão social"
   txtCampo(4).OrderFields = "Razão social"
   txtCampo(4).Relation = ""
   txtCampo(4).Source = "Geral"
   txtCampo(4).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 0
   txtCampo(5).ListFields = "Vista|Prazo"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 4
   txtCampo(6).Mask = ""
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "999.999.999,99"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

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
      Case 3
         txtCampo(3).Filter = "[Codigo do Centro] > " & 0 & ""
                           txtCampo(3).PesqSQLExpression = "SELECT [Centro de Custo].[Codigo do Centro], [Centro de Custo].[Descrição do Centro] FROM [Centro de Custo] WHERE ([Centro de Custo].[Codigo do Centro] > " + CStr(0) + ")"
      Case 4
         txtCampo(4).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0", "[Sequencia do geral] > 0") & ""
                           txtCampo(4).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                              "Geral.Bairro, Municipios.[Sequencia do municipio], Municipios.Uf, Geral.Fone FROM Geral, Municipios WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.[Sequencia do municipio] = Municipios.[Sequencia do municipio])"
   End Select
End Sub

'evento - antes de descarregar o formulário
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If vgPodeFazerUnLoad = False Then
      If vgSituacao <> ACAO_NAVEGANDO And vgBotoesOK Then  'botoeira esta correta?
         AtivaForm Me                                      'entao coloca
      End If
      Cancel = FormPendente(Me)                            've se tem atualizacao pendente
      If Cancel = False Then
         Cancel = True
         timUnLoad.Enabled = True
      End If
   End If
End Sub

'evento - redefinido o tamanho do formulário
Private Sub Form_Resize()
   AjustaPanFundo Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   lblAlteracao.Caption = ""
   FinalizaForm Me
   Set lblTitulo = Nothing
   Set lblAjuste = Nothing
   Set lblParcelamento = Nothing
   Set grdParcelamento = Nothing
   Set lblFechamento = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   If Not Entrada_Contas_Receber Is Nothing Then
      Set Entrada_Contas_Receber = Nothing
   End If
   If Not Parcelas_Entradas_Receber Is Nothing Then
      Parcelas_Entradas_Receber.CloseRecordSet
      Set Parcelas_Entradas_Receber = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmEntConRe = Nothing                      'libera o segmento de código do form
End Sub

'evento - está mudando a linha selecionada do grid
Private Sub Grid_SkipRecord(Index As Integer, vgColumns() As Variant, ByVal vgBookMark As Variant)
   ExecutaGrid Index, vgColumns(), CONDICOES_ESPECIAIS
End Sub

'evento - após efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
End Sub

'evento - antes de efetuar o edit do recordset do grid
Private Sub Grid_BeforeEditRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, True
   ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Parcelas_Entradas_Receber![Sequencia da entrada] = Entrada_Contas_Receber![Sequencia da entrada]
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
   Set Parcelas_Entradas_Receber = vgNewRecordSet
End Sub

'evento - quer validar dados, está gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - após a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   mdiNFE.RemontaForm                                   'vamos atualizar os forms de dados
   Grid(Index).Repaint -1                               'atualiza dados do grid (registro posicionado) 
End Sub

Private Sub Grid_ControlButtonClick(Index As Integer)
   Grid(Index).ShowFilterBar = Not Grid(Index).ShowFilterBar
End Sub

Private Sub Grid_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And Grid(Index).Status = ACAO_NAVEGANDO Then                 'o formulário pai não está em navegação
      mdiNFE.SalvaDados                           'salva o resitro em edição
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
   mdiNFE.RemontaForm                                             'remonta dos os form da tela
   If vgSituacao = ACAO_NAVEGANDO Then
      Reposition , Not vgTemAltGrdOld
   End If
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
   Dim Numero_da_parcela As Integer, Dias As Integer, Data_de_vencimento As Variant
   Dim Valor_da_parcela As Double
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
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = Not ((Not Vazio(Entrada_Contas_Receber![Forma de pagamento])) AND Entrada_Contas_Receber![Forma de pagamento] = "Prazo")
         Case 3
            vgRetVal = Not ((Not Vazio(Entrada_Contas_Receber![Forma de pagamento])) AND Entrada_Contas_Receber![Forma de pagamento] = "Prazo")
         Case 4
            vgRetVal = Not (Not Vazio(Entrada_Contas_Receber![Forma de pagamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = (Not Vazio(Entrada_Contas_Receber![Forma de pagamento]) And TotalParcelas < Entrada_Contas_Receber![Valor total] And Fechado = 0 And Cancelado = 0)
      On Error Resume Next
         Grid(0).AllowEdit = (Not Vazio(Entrada_Contas_Receber![Forma de pagamento]) And Fechado = 0 And Cancelado = 0)
      On Error Resume Next
         Grid(0).AllowDelete = (Not Vazio(Entrada_Contas_Receber![Forma de pagamento]) And Fechado = 0 And Cancelado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas Entradas Receber]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da entrada] = " & Entrada_Contas_Receber![Sequencia da entrada]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D", Entrada_Contas_Receber![Data de entrada], Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Entrada_Contas_Receber![Data de entrada])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Entrada_Contas_Receber![Valor total] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If (Not Vazio(Entrada_Contas_Receber![Forma de pagamento])) AND Entrada_Contas_Receber![Forma de pagamento] = "Prazo" Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Entrada_Contas_Receber![Forma de pagamento])) AND Entrada_Contas_Receber![Forma de pagamento] = "Prazo" Then
               vgIsValid = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
               If Not vgIsValid Then vgColumn =  3 
               vgErrorMessage$ = "Data de vencimento não pode ser vazio!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Entrada_Contas_Receber![Forma de pagamento]) Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Entrada_Contas_Receber![Forma de pagamento] = "Prazo" , TotalParcelas() + Valor_da_Parcela <= Entrada_Contas_Receber![Valor total], Valor_da_parcela = Entrada_Contas_Receber![Valor total]), IIf(Entrada_Contas_Receber![Forma de pagamento] = "Vista" , Valor_da_parcela = Entrada_Contas_Receber![Valor total], Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Entrada_Contas_Receber![Valor total])))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total da Conta!"
            End If
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = UltimaParcela
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            Alteracao  
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            Alteracao  
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Numero_da_parcela = ColumnValue(1)
   Dias = ColumnValue(2)
   Data_de_vencimento = ColumnValue(3)
   Valor_da_parcela = ColumnValue(4)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Parcelas_Entradas_Receber.Eof = False And Parcelas_Entradas_Receber.BOF = False And Parcelas_Entradas_Receber.RecordCount > 0 Then
      Numero_da_parcela = Parcelas_Entradas_Receber![Numero da parcela]
      Dias = Parcelas_Entradas_Receber!Dias
      Data_de_vencimento = Parcelas_Entradas_Receber![Data de vencimento]
      Valor_da_parcela = Parcelas_Entradas_Receber![Valor da parcela]
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


'evento - quando o tempo esgotar  
Private Sub timUnLoad_Timer()
   timUnLoad.Enabled = False
   vgPodeFazerUnLoad = True
   Unload Me
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         seqRegistro = Sequencia_do_Geral 
         mdiNFE.MGeral
      Case 1
         Fechamento  
      Case 2
         AbreConta  
      Case 3
         Cancelamento  
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
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


