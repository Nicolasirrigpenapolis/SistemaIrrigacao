Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEstoque
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8790
   Left                     =   825
   LinkTopic                =   "Estoque"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8355
   ScaleWidth               =   13935
   Top                      =   1420
   Width                    =   14025
   Begin NFE.GPainel Painel
      BackColor                =   &H8000000F
      BevelOuter               =   1
      BevelInner               =   2
      BorderWidth              =   1
      BevelWidth               =   1
      BorderStyle              =   2
      Stretch                  =   -1
      Stripes                  =   0
      SaveGridBars             =   0
      Height                   =   8175
      Left                     =   15
      Width                    =   13755
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   24250
      _extenty                 =   14412
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1485
         Left                     =   480
         Width                    =   1380
         Top                      =   180
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
         Height                   =   465
         Left                     =   2160
         Width                    =   3750
         Top                      =   300
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   20.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Entradas - Estoque"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   13320
         Width                    =   45
         Top                      =   7860
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
         Index                    =   1 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   360
         Left                     =   2160
         Width                    =   960
         Top                      =   1260
         DataField                =   "Sequencia do movimento"
         TabStop                  =  0
         Index                    =   0 
         MultiLine                =   -1
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   330
         Left                     =   2160
         Width                    =   960
         Top                      =   1260
         TabStop                  =  0
         Index                    =   1 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   330
         Left                     =   360
         Width                    =   1710
         Top                      =   1800
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
         ForeColor                =   &H80000008
         Caption                  =   "* Doc / NFe...:"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   3960
         Width                    =   1440
         Top                      =   1800
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "* Dt. Entrada:"
         Index                    =   3 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   5460
         Top                      =   1800
         Width                    =   1500
         Height                   =   330
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   6690
         Top                      =   1815
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
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   5475
         Width                    =   1200
         Top                      =   1815
         DataField                =   "Data de entrada"
         Index                    =   2 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         MaxLength                =   25
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   330
         Left                     =   2160
         Width                    =   1665
         Top                      =   1800
         DataField                =   "Documento"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   270
         Left                     =   7080
         Width                    =   1365
         Top                      =   1800
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "* Fornecedor:"
         Index                    =   4 
      End
      Begin VB.Label labFdo4
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8595
         Top                      =   1800
         Width                    =   4530
         Height                   =   330
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   12855
         Top                      =   1815
         Width                    =   255
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo4
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   12600
         Top                      =   1815
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
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   0
         Alignment                =   0
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   8610
         Width                    =   3975
         Top                      =   1815
         DataField                =   "Sequencia do geral"
         Index                    =   4 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   2
         TabHeight                =   519
         BackColor                =   &H8000000F
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   2
         TabIndex                 =   6
         TabStop                  =   0
         Height                   =   3600
         Left                     =   360
         Width                    =   13020
         Top                      =   2640
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
         Caption                  =   "1 - Produtos"
         Index                    =   0 
         _version                 =   65536
         _extentx                 =   22954
         _extenty                 =   6346
         _stockprops              =   15
         TabCaption(0)            =   "1 - Produtos"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "2 - Financeiro"
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Painel(1)"
         Begin NFE.GListV Grid
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "MS Sans Serif"
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
            TabIndex                 =   7
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   3000
            Left                     =   120
            Width                    =   12810
            Top                      =   420
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   0 
         End
         Begin NFE.GPainel Painel
            BackColor                =   &H8000000F
            BevelOuter               =   1
            BevelInner               =   2
            BorderWidth              =   1
            BevelWidth               =   1
            BorderStyle              =   0
            Stretch                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   3045
            Left                     =   -74880
            Width                    =   12810
            Top                      =   420
            TabStop                  =  0
            _extentx                 =   22584
            _extenty                 =   5368
            Index                    =   1 
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   555
               Width                    =   2010
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
               Caption                  =   "Forma de Pagamento:"
               Index                    =   5 
            End
            Begin VB.Label labFdo5
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &H80000005
               Left                     =   2715
               Top                      =   195
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo5
               Style                    =   1
               TabStop                  =   0
               Index                    =   2
               Left                     =   4020
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
               TabIndex                 =   9
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2730
               Width                    =   1275
               Top                      =   210
               DataField                =   "Forma de pagamento"
               Index                    =   5 
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   1245
               Width                    =   1320
               Top                      =   675
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
               Caption                  =   "Valor do Frete:"
               Index                    =   6 
            End
            Begin VB.Label labFdo6
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &H80000005
               Left                     =   2715
               Top                      =   675
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo6
               Style                    =   1
               TabStop                  =   0
               Index                    =   0
               Left                     =   4020
               Top                      =   690
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
               TabIndex                 =   10
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2730
               Width                    =   1275
               Top                      =   690
               DataField                =   "Valor do frete"
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
               Left                     =   855
               Width                    =   1725
               Top                      =   1155
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
               Caption                  =   "Valor do Desconto:"
               Index                    =   7 
            End
            Begin VB.Label labFdo7
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &H80000005
               Left                     =   2715
               Top                      =   1155
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo7
               Style                    =   1
               TabStop                  =   0
               Index                    =   0
               Left                     =   4020
               Top                      =   1170
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
               TabIndex                 =   11
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2730
               Width                    =   1275
               Top                      =   1170
               DataField                =   "Valor do desconto"
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
               Left                     =   825
               Width                    =   1755
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
               Caption                  =   "Total dos Produtos:"
               Index                    =   8 
            End
            Begin VB.Label labFdo8
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &HE0E0E0
               Left                     =   2715
               Top                      =   1635
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo8
               Style                    =   1
               TabStop                  =   0
               Index                    =   0
               Left                     =   4020
               Top                      =   1650
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
               Alignment                =   1
               Locked                   =   -1
               TabIndex                 =   12
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2730
               Width                    =   1275
               Top                      =   1650
               DataField                =   "Valor total dos produtos"
               TabStop                  =  0
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
               Left                     =   1515
               Width                    =   1065
               Top                      =   2115
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
               Caption                  =   "Total do IPI:"
               Index                    =   9 
            End
            Begin VB.Label labFdo9
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &HE0E0E0
               Left                     =   2700
               Top                      =   2115
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo9
               Style                    =   1
               TabStop                  =   0
               Index                    =   1
               Left                     =   4005
               Top                      =   2130
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
               Alignment                =   1
               Locked                   =   -1
               TabIndex                 =   13
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2715
               Width                    =   1275
               Top                      =   2130
               DataField                =   "Valor total do ipi"
               TabStop                  =  0
               Index                    =   9 
               MultiLine                =   -1
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   1230
               Width                    =   1350
               Top                      =   2595
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
               Caption                  =   "Total ICMS ST:"
               Index                    =   10 
            End
            Begin VB.Label labFdo10
               Appearance               =  0
               BorderStyle              =  1
               BackColor                =  &HE0E0E0
               Left                     =   2700
               Top                      =   2595
               Width                    =   1575
               Height                   =   300
            End
            Begin VB.CommandButton bottxtCampo10
               Style                    =   1
               TabStop                  =   0
               Index                    =   0
               Left                     =   4005
               Top                      =   2610
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
               Alignment                =   1
               Locked                   =   -1
               TabIndex                 =   14
               TabStop                  =   -1
               Height                   =   270
               Left                     =   2715
               Width                    =   1275
               Top                      =   2610
               DataField                =   "Valor total icms st"
               TabStop                  =  0
               Index                    =   10 
               MultiLine                =   -1
            End
            Begin VB.Label Label
               Appearance               =   0
               BackColor                =   &H0
               BackStyle                =   1
               BorderStyle              =   1
               AutoSize                 =   0
               Height                   =   255
               Left                     =   7530
               Width                    =   5175
               Top                      =   120
               Alignment                =   2
               BeginProperty Font
                  Name                     =   "Microsoft Sans Serif"
                  CharSet                  =   0
                  Weight                   =   700
                  Size                     =   9.75
                  Underline                =   0
                  Italic                   =   0
                  StrikeThrough            =   0
               EndProperty
               ForeColor                =   &HFFFFFF
               Caption                  =   "Parcelamento do Contas á Pagar"
               Index                    =   11 
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   7575
               Width                    =   600
               Top                      =   2700
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
               Index                    =   12 
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
               Left                     =   8250
               Width                    =   165
               Top                      =   2745
               Index                    =   0 
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   8475
               Width                    =   990
               Top                      =   2700
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
               Index                    =   13 
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
               Left                     =   9570
               Width                    =   165
               Top                      =   2745
               Index                    =   1 
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   9870
               Width                    =   1050
               Top                      =   2700
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
               Index                    =   14 
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
               Left                     =   11085
               Width                    =   150
               Top                      =   2745
               Index                    =   2 
            End
            Begin VB.Label Label
               Appearance               =   1
               BackColor                =   &H8000000F
               BackStyle                =   1
               BorderStyle              =   0
               AutoSize                 =   -1
               Height                   =   240
               Left                     =   11310
               Width                    =   300
               Top                      =   2700
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
               Index                    =   15 
            End
            Begin NFE.GListV Grid
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
                  Name                     =   "MS Sans Serif"
                  CharSet                  =   0
                  Weight                   =   400
                  Size                     =   8.25
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
               Height                   =   2250
               Left                     =   7530
               Width                    =   5190
               Top                      =   420
               Caption                  =   ""
               ShowGridCaption          =   0
               Index                    =   1 
            End
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   300
         Left                     =   8595
         Width                    =   2415
         Top                      =   6375
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
         ForeColor                =   &H80000008
         Caption                  =   "Total do Movimento:"
         Index                    =   16 
      End
      Begin VB.Label labFdo11
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HE0E0E0
         Left                     =   11160
         Top                      =   6375
         Width                    =   2205
         Height                   =   330
      End
      Begin VB.CommandButton bottxtCampo11
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   13095
         Top                      =   6390
         Width                    =   255
         Height                   =   300
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
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
         BorderStyle              =   0
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   15
         TabStop                  =   -1
         Height                   =   300
         Left                     =   11175
         Width                    =   1905
         Top                      =   6390
         DataField                =   "Valor total do movimento"
         TabStop                  =  0
         Index                    =   11 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   360
         Width                    =   600
         Top                      =   6480
         Alignment                =   0
         BeginProperty Font
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   12
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H0
         Caption                  =   "Obs:"
         Index                    =   17 
      End
      Begin VB.TextBox txtCp
         Appearance               =   1
         BackColor                =   &H80000005
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
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
         TabIndex                 =   8
         TabStop                  =   0
         Height                   =   840
         Left                     =   360
         Width                    =   13050
         Top                      =   6840
         DataField                =   "Historico"
         Index                    =   12 
         MultiLine                =   -1
         ScrollBars               =   2
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   16
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   270
         Left                     =   13140
         Width                    =   225
         Top                      =   1800
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
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   17
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   7500
         Width                    =   1170
         Top                      =   120
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "MS Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Importar"
         Index                    =   1 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   19
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   10680
         Width                    =   1290
         Top                      =   180
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
         Caption                  =   "Fechar"
         Index                    =   2 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   20
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   330
         Left                     =   3720
         Width                    =   1290
         Top                      =   1200
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
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   360
         Left                     =   2160
         Width                    =   2865
         Top                      =   780
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
         ForeColor                =   &H8000
         Caption                  =   "Movimento Fechado"
         Index                    =   18 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   21
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   12000
         Width                    =   1290
         Top                      =   180
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
         Caption                  =   "Gerar"
         Index                    =   4 
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
         TabIndex                 =   22
         TabStop                  =   0
         Height                   =   270
         Left                     =   8595
         Width                    =   2865
         Top                      =   1440
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
         Caption                  =   "Entrada Para Conserto"
         Index                    =   0 
         DataField                =   "Entrada Conserto"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   270
         Left                     =   6750
         Width                    =   1695
         Top                      =   2340
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "Centro de Custo:"
         Index                    =   19 
      End
      Begin VB.Label labFdo13
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8580
         Top                      =   2340
         Width                    =   4560
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo13
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   12870
         Top                      =   2355
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo13
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   12615
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
            Size                     =   11.25
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
         Left                     =   8595
         Width                    =   4005
         Top                      =   2355
         DataField                =   "Tipo de conta"
         Index                    =   13 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   23
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   270
         Left                     =   13140
         Width                    =   225
         Top                      =   2340
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
         Index                    =   5 
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
Attribute VB_name = "frmEstoque"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: Estoque
'* Função....: Entradas - Estoque
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:39
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
Dim txtCampo(13) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(0) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Sequencia_do_movimento As Long, Documento As String, Sequencia_do_geral As Long
Dim Data_de_entrada As Variant, Forma_de_pagamento As String, Historico As String
Dim Cancelado As Boolean, Valor_do_frete As Double, Valor_do_desconto As Double
Dim Valor_total_dos_produtos As Double, Valor_total_do_ipi As Double, Valor_total_do_movimento As Double
Dim Valor_total_icms_st As Double, Sequencia_da_nota_fiscal As Long, Data_da_Alteracao As Variant
Dim Hora_da_Alteracao As Variant, Usuario_da_Alteracao As String, Fechado As Boolean
Dim Entrada_Conserto As Boolean, Tipo_de_conta As String
Public lblMovimento As Object, lblAjuste As Object, txtSequencia As Object
Public GrdProdutos As Object, lblParcelamento As Object, grdParcelamento As Object
Public cmdImportar As Object, cmdFechar As Object, cmdAbrir As Object
Public lblFechamento As Object
Dim Movimento_do_estoque As New GRecordSet, Produtos_movimento_estoqu As New GRecordSet, Parcelas_movimento_estoqu As New GRecordSet

Dim ProdutoAux As New GRecordSet, UnidadeAux As New GRecordSet

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
    Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do movimento]) Seq From [Movimento do estoque]")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function 


Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Sub Comandos(KeyAscii As Integer, Sequencia_do_movimento As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double, _
   Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_item As Long)
   On Error GoTo DeuErro
   
   If KeyAscii = vbKeyF12 Then
      With GrdProdutos
         Select Case .ColumnField(.Col)
            Case "Sequencia do produto"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmProdutos.Show        
         End Select
      End With
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub

Private Sub DeletaItens()' Qdo excluir Registro Mae
Dim Tb As New GRecordSet, Parcelas As New GRecordSet

  If vgSITUACAO = ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
  Set Parcelas = vgDb.OpenRecordSet("SELECT * From [Parcelas movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
  
   If Parcelas.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Parcelas movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento
      DeletaConta
   End If
  
   If Tb.RecordCount > 0 Then
      Do While Not Tb.EOF
      vgDb.Execute "DELETE FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " And [Sequencia do produto] = " & Tb![Sequencia Do produto]
      AtualizaEstoque(Tb![Sequencia Do produto])
      UltimoMvto(Tb![Sequencia Do produto])
      UltimoFornecedor(Tb![Sequencia Do produto])
      ValorUnitario(Tb![Sequencia Do produto])
      ValorCustoProdutos(Tb![Sequencia Do produto]) 
      Tb.MoveNext
      Loop
      Tb.MoveFirst
   End If
  End If   
End Sub

Private Sub MostraFormulas()

   If Fechado Then
      lblFechamento.Caption = "Movimento Fechado"
      lblFechamento.ForeColor = &H0
   Else
      lblFechamento.Caption = "Movimento não Finalizado"
      lblFechamento.ForeColor = &H8FF
   End If
   If Err Then Err = 0: lblFechamento.Caption = ""   
   If Cancelado Then
      lblMovimento.ForeColor = &H8FF
      'lblLinha.ForeColor = &H8FF
      lblMovimento.Caption = "Movimento Cancelado"
   Else
     If Not Entrada_conserto Then
       lblMovimento.ForeColor = &H0
      'lblLinha.ForeColor = &H0
       lblMovimento.Caption = "Entradas - Estoque"
     Else
       lblMovimento.ForeColor = &H0
       'lblLinha.ForeColor = &H0
       lblMovimento.Caption = "Entrada Para Conserto" 
     End If 
   End If
   If Err Then Err = 0: lblMovimento.Caption = "" 
   TxtSequencia.Text = Format(Sequencia_do_movimento, "000000")   
   If Err Then Err.Clear 
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_total_do_movimento) Or (TotalParcelas > Valor_total_do_movimento) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   
End Sub


Private Function AtuEstoque(Sequencia_do_movimento As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double, _
   Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_item As Long)
Dim Tb As New GRecordSet

  If vgSITUACAO = -ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
   If Tb.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " And [Sequencia do item] = " & Sequencia_do_item
   End If
  End If 
  
  AtualizaEstoque(Sequencia_do_produto)
  UltimoMvto(Sequencia_do_produto)
  UltimoFornecedor(Sequencia_do_produto)
  ValorUnitario(Sequencia_do_produto)
  ValorCustoProdutos(Sequencia_do_produto)
  AjustaValores
   
End Function


Private Function InfoProdutos(Sequencia_do_movimento As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double, _
   Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_item As Long, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Sequencia do produto] = " & Sequencia_do_Produto, ProdutoAux
      
   If ProdutoAux.RecordCount = 0 Then Exit Function
   
   TbAuxiliar "Unidades", "[Sequencia da unidade] = " & ProdutoAux![Sequencia da unidade], UnidadeAux
   
   Select Case Oq      
      Case "Sigla"
         InfoProdutos = UnidadeAux![Sigla da Unidade]
      Case "Valor Unitário"
         InfoProdutos = Round((ProdutoAux![Valor Total]),4)
      Case "Estoque"
         InfoProdutos = ProdutoAux![Estoque]
   End Select

End Function


Private Function VerificaDocumento() As Boolean
   Dim Tabela As New GRecordSet
   
   If Vazio(Documento) Then VerificaDocumento = False: Exit Function
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then VerificaDocumento = True: Exit Function
   If Sequencia_do_geral = 0 Then VerificaDocumento = True: Exit Function
      
   Set Tabela = vgDb.OpenRecordSet("Select [Sequencia do movimento], documento, [Sequencia do geral] From [Movimento do estoque] " & _
                                   "Where [Sequencia do geral] = " & Sequencia_do_geral & " And Documento = '" & Documento & "'")
                                    
   If Tabela.RecordCount > 1 Then
      If MsgBox("Já existe esse Documento para esse Geral Deseja Prosseguir?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then
         VerificaDocumento = False
         Exit Function
      Else
         VerificaDocumento = True
         Exit Function
      End If
   ElseIf Tabela.RecordCount = 1 Then
   If Tabela![Sequencia Do Geral] = Sequencia_do_geral And Tabela!Documento = Documento And Tabela![Sequencia Do movimento] <> Sequencia_do_movimento Then
      If MsgBox("Já existe esse Documento para esse Geral Deseja Prosseguir?", vbYesNo + vbQuestion, vaTitulo) = vbNo Then
         VerificaDocumento = False
      Exit Function
      Else
         VerificaDocumento = True
      Exit Function
      End If
      End If
   End If
   VerificaDocumento = True
    
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


Private Sub RepositionEstoque()
   On Error Resume Next
   
   If vgSituacao = ACAO_INCLUINDO Then
      Set Produtos_movimento_estoqu = Nothing
      Set Parcelas_movimento_estoqu = Nothing
   End If
   
   CarregaTotalizador
   'Alteração
   lblAlteracao.Caption = Movimento_do_estoque![Usuario da Alteração] & " " & Movimento_do_estoque![Data da Alteração] & " " & Movimento_do_estoque![Hora da Alteração]
  
End Sub


Private Sub CarregaTotalizador()
   On Error Resume Next

   With grdProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Tot. Produtos").Index) = True
      .ShowSumCol(.Columns("Vr. Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowSumCol(.Columns("Vr. ICMS ST").Index) = True
      .ShowSumCol(.Columns("Valor IPI").Index) = True
      .ShowSumCol(.Columns("Aliquota do ipi").Index) = False
      .ColumnWidth(CInt(.Columns("Sequencia do item").Index)) = 0
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
   With grdParcelamento
      .ShowSumBar = True
      .ShowSumCol(.Columns("Numero da Parcela").Index) = False
      .ShowSumCol(.Columns("Dias").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
End Sub

Private Function ProcessaProdutos(Sequencia_do_movimento As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double, _
   Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_item As Long)
         
   On Error GoTo DeuErro
   
   vgDb.BeginTrans      
   vgDb.Execute "Update [Produtos movimento estoque] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " AND [Sequencia do item] = " & Sequencia_do_item
   If Aliquota_do_ipi > 0 Then
      vgDb.Execute "Update [Produtos movimento estoque] Set [Valor do ipi] = " & Substitui(Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2), ",", ".", SO_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " AND [Sequencia do item] = " & Sequencia_do_item 'Valor do IPI
   Else
      vgDb.Execute "Update [Produtos movimento estoque] Set [Valor do ipi] = 0 WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " AND [Sequencia do item] = " & Sequencia_do_item
   End If
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Public Sub AjustaValores()
   Dim IPIProdutos As New GRecordSet, ValorProdutos As New GRecordSet
   Dim ValorICMSST As New GRecordSet, ValorMvto As Currency
   Dim Produtos As New GRecordSet
 
   On Error GoTo DeuErro
   
   Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
   
   'AtualizaLancamentos 'Atualizando os Lançamentos dos Produtos: Valor de Custo, Ultimo Fornecedor, Ultima Compra e etc...
          
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT ROUND(SUM([Valor do ipi]), 2) Total FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento) 'IPI dos Produtos
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento) 'Produtos Novos
   Set ValorICMSST = vgDb.OpenRecordSet("SELECT SUM([Valor icms st]) Total FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento) 'Substituição Produtos                                                                                             
   ValorMvto = IPIProdutos!Total + ValorProdutos!Total + ValorICMSST!Total + Valor_do_Frete - Valor_do_desconto
   ValorMvto = Format(ValorMvto, "##,###,##0.00")
   
   If Produtos.RecordCount = 0 Then
      ValorMvto = 0 
      If Valor_do_frete > 0 Or Valor_do_desconto > 0 Then 
         vgDb.BeginTrans   
         vgDb.Execute "Update [Movimento do estoque] Set [Valor do frete] = 0 WHERE [Sequencia do movimento] = " & Sequencia_do_movimento
         vgDb.Execute "Update [Movimento do estoque] Set [Valor do desconto] = 0 WHERE [Sequencia do movimento] = " & Sequencia_do_movimento
         vgDb.CommitTrans
      End If  
   End If

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Movimento do estoque] Set [Valor total do ipi] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento 'IPI Produtos
   vgDb.Execute "Update [Movimento do estoque] Set [Valor total dos produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento 'Produtos 
   vgDb.Execute "Update [Movimento do estoque] Set [Valor total icms st] = " & Substitui(ValorICMSST!Total, ",", ".", SO_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento 'Substituição
   vgDb.Execute "Update [Movimento do estoque] Set [Valor total do movimento] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Sequencia do movimento] = " & Sequencia_do_movimento 'Valor da Entrada
   vgDb.CommitTrans
     
   Alteracao
   
   If Produtos.RecordCount > 0 And vgSituacao = ACAO_EDITANDO Then
      Do While Not Produtos.EOF
      UltimoMvto(Produtos![Sequencia Do produto])
      UltimoFornecedor(Produtos![Sequencia Do produto])
      Produtos.MoveNext
      Loop
      Produtos.MoveFirst  
   End If
   
   AVISTA
   
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


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


Private Function FiltroProduto() As String

  On Error Resume Next ' Se der Erro que si foda...
  FiltroProduto = "[Sequencia do produto] > 0"
  If vgSituacao <> ACAO_NAVEGANDO Then
  FiltroProduto = FiltroProduto & " And [Sequencia do produto] > 0 AND Inativo = 0"
  End If 
    
End Function


Private Function ValidaProduto(Sequencia_do_movimento As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double, _
   Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_item As Long) As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Produtos Where [Sequencia do Produto] = " & Sequencia_do_produto)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaProduto = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaProduto = True: Exit Function      
   End If
  End If
  
End Function


Private Function SeqItem() As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
 
    If Tb.RecordCount > 0 Then
      SeqItem = Tb!Seq + 1
    Else
      SeqItem = 1
    End If
      
End Function


'Private Sub AbreXML()
'   On Error Resume Next
   
'   Load frmIXML
'   frmIXML.txtNota.Value = Sequencia_do_movimento
'   frmIXML.Show     

'End Sub 


Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC FROM [Parcelas movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Function UltimoVencimento() As Variant
 Dim Tb As New GRecordSet, RetVal As Variant
   
  On Error Resume Next
      
  Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de vencimento]) [Data de Vencimento] FROM [Parcelas movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
      
  RetVal = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, grdParcelamento.Col - 1)
  
   If Movimento_do_estoque![Forma de pagamento] = "Prazo" Then
      If Not IsNull(Tb![Data de Vencimento]) Then
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Movimento_do_estoque![Data de entrada]
      End If
   Else
      UltimoVencimento = RetVal = Movimento_do_estoque![Data de entrada]
   End If
   
End Function


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


Private Sub LancaParcelas()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
   Dim SeqBaixa As Long
   
   On Error GoTo DeuErro
      
   If Vazio(Forma_de_pagamento) Then Exit Sub
   If TotalParcelas() <> Valor_total_do_movimento Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas pagar")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas pagar", "[Sequencia do movimento] = " & Sequencia_do_movimento & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew
         !Parcela = Parcelas![Numero da parcela]
         ![Sequencia Do geral] = Sequencia_do_Geral
         !Documento = Documento
         ![Tipo de conta] = Tipo_de_Conta
         ![Data de entrada] = Data_de_entrada
         !Historico = "Entrada do Estoque Documento: " + Documento
         ![Forma de Pagamento] = Forma_de_pagamento
           If ![Forma de pagamento] = "Vista" And Empresa![Baixa Pagar Auto] Then
              ![Valor pago] = Parcelas![Valor da parcela]
              ![Valor restante] = 0
              ![Data da baixa] = Data_de_entrada
          Else
            ![Valor pago] = 0
            ![Valor restante] = Parcelas![Valor da parcela]
            ![Data da baixa] = Null
          End If
         ![Data de vencimento] = Parcelas![Data de vencimento]
         ![Valor total] = Valor_total_do_movimento
         ![Valor da parcela] = Parcelas![Valor da parcela]
         ![Sequencia Do movimento] = Sequencia_do_movimento
         .Update
         .BookMark = .LastModified
      End With
      Parcelas.MoveNext
   Loop
   
   'Baixa 
    If Forma_de_pagamento = "Vista" And Empresa![Baixa Pagar Auto] Then
       Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas Pagar]")
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
         ' MsgBox "PAGAMENTO A VISTA - (Informe o Caixa ou a Conta Corrente!)", vbInformation + vbOKOnly, vaTitulo
          frmBxPagar.Show
          frmBxPagar.SetFocus
          frmBxPagar.vgFiltroInicial = "[Sequencia da Baixa] = " & SeqBaixa
          InicializaFiltro frmBxPagar
          frmBxPagar.Reposition
          'Unload ME      
    End If
   
   vgDb.CommitTrans

DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Sub AlteraManutencao() 'Qdo altera Registro Mãe
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet
   
   On Error GoTo DeuErro
   
   AjustaValores
            
   Set Parcelas = Parcelas_movimento_estoqu
   
   If Parcelas.RecordCount = 0 Then Exit Sub
         
   vgDb.BeginTrans
   
   Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
   If Manutencao.RecordCount = 0 Then Exit Sub
   Do While Not Manutencao.EOF
      With Manutencao
         ![Sequencia Do geral] = Movimento_do_estoque![Sequencia Do geral]
         ![Data de entrada] = Movimento_do_estoque![Data de entrada]
         !Historico = "Entrada do Estoque Documento: " + Documento
         ![Forma de pagamento] = Movimento_do_estoque![Forma de pagamento]
         ![Valor total] = Movimento_do_estoque![Valor total Do movimento]
         !Documento = Movimento_do_estoque!Documento
         .Update
         .BookMark = .LastModified
      End With
      Manutencao.MoveNext
   Loop    
   vgDb.CommitTrans
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
   
End Sub


Private Sub AlteraParcela(Sequencia_do_movimento As Long, Numero_da_parcela As Integer, Dias As Integer, _
   Data_de_vencimento As Variant, Valor_da_parcela As Double)
 Dim Manutencao As New GRecordset, RetVal As Variant
   
   On Error GoTo DeuErro

   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Movimento_do_estoque![Sequencia Do movimento] & " AND Parcela = " & Numero_da_parcela)

   vgdb.BeginTrans
   
   RetVal = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, CInt(grdParcelamento.Columns("Data de vencimento").Index))
   
   If Manutencao.RecordCount > 0 Then
      With Manutencao
         !Parcela = Numero_da_parcela
         ![Sequencia Do geral] = Movimento_do_estoque![Sequencia Do geral]
         !Documento = Movimento_do_estoque!Documento
         ![Data de Entrada] = Movimento_do_estoque![Data de entrada]
         !Historico = "Entrada do Estoque Documento: " + Documento
         ![Forma de pagamento] = Movimento_do_estoque![Forma de pagamento]
         ![Valor pago] = 0
         ![Valor restante] = Valor_da_parcela 
         ![Data da baixa] = Null
         ![Data de vencimento] = RetVal 
         ![Valor total] = Movimento_do_estoque![Valor total Do movimento] 
         ![Valor da parcela] = Valor_da_parcela          
         ![Sequencia Do movimento] = Movimento_do_estoque![Sequencia Do movimento]
         .Update
         .BookMark = .LastModified         
      End With     
   End If  
   
   vgdb.CommitTrans

DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If 
 
End Sub


Private Sub DeletaParcela(Sequencia_do_movimento As Long, Numero_da_parcela As Integer, Dias As Integer, _
   Data_de_vencimento As Variant, Valor_da_parcela As Double)
   Dim Manutencao As New GRecordset
   
   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Movimento_do_estoque![Sequencia Do movimento] & " AND Parcela = " & Numero_da_parcela)

   vgdb.BeginTrans
   
   If Manutencao.RecordCount > 0 Then
      vgdb.Execute "DELETE FROM [Manutenção Contas pagar] WHERE [Sequencia do movimento] = " & Movimento_do_estoque![Sequencia do movimento] & " AND Parcela = " & Numero_da_parcela 
   End If      
     
   vgdb.CommitTrans
   
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If   

End Sub


Private Sub DeletaConta() 'Qdo exclui registro Mãe
   Dim Manutencao As New GRecordset

   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
    
   If Manutencao.RecordCount = 0 Then Exit Sub
   
   vgDb.Execute "DELETE FROM [Manutenção contas pagar] " & _
                "WHERE [Sequencia do movimento] = " & Sequencia_do_movimento   

DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbokOnly, vaTitulo
      vgdb.RollBackTrans
   End If

End Sub


Private Sub Fechamento()
 Dim Produtos As New GRecordSet, Parcelas As New GRecordSet
 Dim Manutencao As New GRecordSet
 
 On Error Resume Next
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
  
 If Produtos.RecordCount = 0 Then
    MsgBox("Impossivel Fechar o Movimento!! Movimento sem Itens")
    Exit Sub
 End If
 
  If Parcelas.RecordCount > 0 Then
   If TotalParcelas() <> Valor_total_do_movimento Then 
      MsgBox("Impossivel Fechar o Movimento!! A Soma das Parcelas não Correspondem ao Valor Total do Movimento")
   Exit Sub
   End If
  End If
  
  If Parcelas.RecordCount > 0 And Manutencao.RecordCount = 0 Then
     LancaParcelas 
  Else
     DeletaConta
     LancaParcelas     
  End If
  
  If Parcelas.RecordCount = 0 And Not Entrada_conserto Then If MsgBox("Deseja Fechar o Movimento sem Incluir as Parcelas do Contas a Pagar?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
  vgDb.Execute "UPDATE [Movimento do estoque] Set Fechado = 1 WHERE [Sequencia do movimento] = " & Sequencia_do_movimento
  Reposition True
              
End Sub


Private Sub AbreMvto()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean, TemDevolucao As Boolean
 Dim Manu As New GRecordSet, Devolucao As New GRecordSet

 On Error Resume Next
 
 Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas pagar] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & ")")
 Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas pagar] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
 Set Devolucao = vgDb.OpenRecordSet("SELECT * FROM [Devolução de compra] WHERE Cancelado = 0 And [Sequencia do movimento] IN (SELECT [Sequencia do movimento] FROM [Movimento do estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & ")")
  
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If
   
    If Devolucao.RecordCount > 0 Then
       TemDevolucao = True
    Else
       TemDevolucao = False
    End If
   
   If TemBaixa Then
      MsgBox("Impossivel abrir o Movimento existe Baixa(s) no Contas a Pagar para Abri-lo Exclua a(s) Baixa(s)")
      Exit Sub
   End If
   
   If TemDevolucao Then
      MsgBox("Impossivel abrir o Movimento Tem Devoluções para Abri-lo Você deve Cancelar a(s) Devoluções")
      Exit Sub
   End If
 
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "UPDATE [Movimento do estoque] Set Fechado = 0 Where [Sequencia do movimento] = " & Sequencia_do_movimento 
       If Manu.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção contas pagar] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento
       End If          
       Reposition True 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  InputYgorsBom = ""
  
End Sub  


Private Sub AbreGerar()
   Dim Geral As New GRecordSet
   Dim DtaHoje As Date, DtaParcela As Date, Parcelas As New GRecordSet
   
   On Error Resume Next
   
   DtaHoje = Date 'Data de Hoje
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
   If Parcelas.RecordCount > 0 Then
       Do While Not Parcelas.EOF
       DtaParcela = Parcelas![Data de vencimento] 
       If DtaParcela < DtaHoje Then ' Verifica de o Vencto da Parcela não é menor que a data de hj
       MsgBox "Atenção!! Data de Vencimento da Parcela" & " " & Parcelas![Numero da parcela] & " " & "é Menor que a data de Hoje. Impossível Gerar a NFe.", vbExclamation + vbOKOnly, vaTitulo :Exit Sub
       End If    
       Parcelas.MoveNext
       Loop
       Parcelas.MoveFirst
   End If
   
   'Vamos perguntar para ninguem fazer merda
   If MsgBox("Deseja Gerar a Nota Fiscal?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub   
   If TotalParcelas > 0 And TotalParcelas <> Valor_total_do_movimento Then
      MsgBox "Parcelamento Incompleto. Impossível Gerar a NFe.", vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
      
   Unload Me
   Load frmMvtoNFe ' Abre Form p/ Gerar NFe 
   frmMvtoNFe.RepositionMvto
   frmMvtoNFe.Executar INICIALIZACOES
   frmMvtoNFe.Executar INI_APELIDOS
   frmMvtoNFe.Show
   SendK vbKeyReturn
   SendK vbKeyReturn
   SendK vbKeyReturn 
         
End Sub 


Private Sub TrocaPaConserto() 'Qdo Altera de Mvto p/ Conserto Ação Change do Form
   Dim Parcelas As New GRecordSet

   On Error GoTo DeuErro
   
   If Not Entrada_Conserto Then: Exit Sub   

   Set Parcelas = vgdb.OpenRecordset("SELECT * FROM [Parcelas movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento)
   
   If Parcelas.RecordCount = 0 Then: Exit Sub
    
   If Parcelas.RecordCount > 0 Then 
      vgDb.Execute "DELETE FROM [Parcelas movimento estoque] " & _
                   "WHERE [Sequencia do movimento] = " & Sequencia_do_movimento   
   End If     

DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbokOnly, vaTitulo
      vgdb.RollBackTrans
   End If

End Sub 


Private Sub AVISTA()
Dim Parcelas As New GRecordSet
Dim Produtos As New GRecordSet

 Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento)
 TbAuxiliar "Parcelas movimento estoque", "[Sequencia do movimento] = " & Sequencia_do_movimento, Parcelas
 
 If Produtos.RecordCount = 0 And Parcelas.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [Parcelas movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento
    GrdParcelamento.Rebind
    Exit Sub
 End If

 If (Movimento_do_estoque![Forma de pagamento] <> "Vista" Or Movimento_do_estoque![Valor total Do movimento] = 0) Then Exit Sub 
    With Parcelas   
      If Parcelas.RecordCount = 0 Then .AddNew
      ![Sequencia Do Movimento] = Movimento_do_estoque![Sequencia Do movimento]
      ![Numero da Parcela] = 1
      !Dias = 0
      ![Data de Vencimento] = Movimento_do_estoque![Data de entrada]
      ![Valor da Parcela] = Movimento_do_estoque![Valor total Do movimento]
      .Update
      .Bookmark = .lastModified
   End With
   GrdParcelamento.Rebind
   
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
         Sequencia_do_movimento = txtCampo(0).Value
         Documento = txtCampo(3).Value
         Sequencia_do_geral = txtCampo(4).Value
         Data_de_entrada = txtCampo(2).Value
         Forma_de_pagamento = txtCampo(5).Value
         Historico = txtCampo(12).Value
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
         Valor_do_frete = txtCampo(6).Value
         Valor_do_desconto = txtCampo(7).Value
         Valor_total_dos_produtos = txtCampo(8).Value
         Valor_total_do_ipi = txtCampo(9).Value
         Valor_total_do_movimento = txtCampo(11).Value
         Valor_total_icms_st = txtCampo(10).Value
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Usuario_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Usuario da Alteração])
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Entrada_Conserto = chkCampo(0).Value
         Tipo_de_conta = txtCampo(13).Value
      Else
         Sequencia_do_movimento = IIf(IsNull(vgTb![Sequencia do movimento]), 0, vgTb![Sequencia do movimento])
         Documento = IIf(IsNull(vgTb!Documento), "", vgTb!Documento)
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Data_de_entrada = vgTb![Data de entrada]
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Valor_do_frete = IIf(IsNull(vgTb![Valor do frete]), 0, vgTb![Valor do frete])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Valor_total_dos_produtos = IIf(IsNull(vgTb![Valor total dos produtos]), 0, vgTb![Valor total dos produtos])
         Valor_total_do_ipi = IIf(IsNull(vgTb![Valor total do ipi]), 0, vgTb![Valor total do ipi])
         Valor_total_do_movimento = IIf(IsNull(vgTb![Valor total do movimento]), 0, vgTb![Valor total do movimento])
         Valor_total_icms_st = IIf(IsNull(vgTb![Valor total icms st]), 0, vgTb![Valor total icms st])
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Entrada_Conserto = IIf(IsNull(vgTb![Entrada Conserto]), 0, vgTb![Entrada Conserto])
         Tipo_de_conta = IIf(IsNull(vgTb![Tipo de conta]), "", vgTb![Tipo de conta])
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
      i = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0 And vgPWUsuario = "ADMIN") and Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0) and Permitido(vgIdentTab, ACAO_EDITANDO)
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
      vgOk = (VerificaDocumento())
      vgMsg$ = "Documento Invalido!"
      If Not vgOk Then vgColumn = 4
      If vgOk Then
         vgOk = ((Not Vazio(Data_de_entrada)) And (IsDate(Data_de_entrada) Or Vazio(Data_de_entrada)))
         vgMsg$ = "Data de entrada não pode ser vazio!"
         If Not vgOk Then vgColumn = 3
      End If
      If vgOk Then
         If Sequencia_do_geral = 0 Then
            vgOk = (Sequencia_do_geral > 0)
            vgMsg$ = "Fornecedor não pode ser em Branco!"
            If Not vgOk Then vgColumn = 5
         End If
      End If
      If vgOk Then
         If vgSituacao = ACAO_EDITANDO Then
            vgOk = (ValidaGeral())
            vgMsg$ = "Fornecedor Inativo!"
            If Not vgOk Then vgColumn = 5
         End If
      End If
      If vgOk Then
         If False Then
            vgOk = (Not Vazio(Tipo_de_conta))
            vgMsg$ = "Centro de Custo não pode ser vazio!"
            If Not vgOk Then vgColumn = 14
         End If
      End If
      If vgOk Then
         vgOk = (Valor_do_frete >= 0)
         vgMsg$ = "Valor do frete inválido!"
         If Not vgOk Then vgColumn = 7
      End If
      If vgOk Then
         vgOk = (Valor_do_desconto >= 0)
         vgMsg$ = "Valor do desconto inválido!"
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
                  chkCampo(0).Value = False
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 4
         PoeRelEFiltroCbo 13
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 4
         PoeRelEFiltroCbo 13
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Entrada_Conserto
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
      Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
      Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
      InicializaApelidos COM_TEXTBOX
      vgTb![Entrada Conserto] = Entrada_Conserto
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_INCLUINDO Then
         AjustaValores
      ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
         AlteraManutencao
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
   Set Movimento_do_estoque = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   On Error Resume Next
   For i = 0 To 1
      Select Case i
         Case 0
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
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
         Case 1
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(1).CloseRecordset
            Else
               x$ = ExecutaGrid(1, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(1).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(1).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(1).CloseRecordSet
                  End If
                  Grid(1).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(1, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionEstoque  
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
   Ok = (False)
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   Ok = (Sequencia_do_geral > 0 And Parcelas_movimento_estoqu.RecordCount = 0)
   txtCampo(5).Enabled = (Ok And txtCampo(5).Editable)
   Ok = (False)
   txtCampo(8).Enabled = (Ok And txtCampo(8).Editable)
   Ok = (False)
   txtCampo(9).Enabled = (Ok And txtCampo(9).Editable)
   Ok = (False)
   txtCampo(10).Enabled = (Ok And txtCampo(10).Editable)
   Ok = (Produtos_movimento_estoqu.RecordCount = 0)
   Botao(1).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0 And VgSituacao = ACAO_NAVEGANDO)
   Botao(2).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = True)
   Botao(3).Enabled = Ok
   Ok = (Sequencia_da_nota_fiscal = 0 AND vgSituacao = ACAO_NAVEGANDO AND Cancelado = False AND Fechado = True AND PermitidoMenu(MnuNFe.Tag) And Empresa![Emite NFe])
   Botao(4).Enabled = Ok
   Ok = (False)
   chkCampo(0).Enabled = (Ok And chkCampo(0).Editable)
   Ok = (False)
   Label(19).Enabled = Ok
   Ok = (False)
   txtCampo(13).Enabled = (Ok And txtCampo(13).Editable)
   Ok = (False)
   Botao(5).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   txtCampo(0).Visible = (False)
   Botao(3).Visible = (Cancelado = False AND Fechado = True And vgSituacao = ACAO_NAVEGANDO AND Sequencia_da_nota_fiscal = 0)
   Label(18).Visible = (Cancelado = 0)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 4 Or Index = 5 Or Index = 13 Then
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
         Case 2
            If Len(txtCp(2).Text) = 0 Then
               txtCampo(2).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 4
            PoeRelEFiltroCbo 4
         Case 13
            PoeRelEFiltroCbo 13
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


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      chkCampo(0).Value = Entrada_Conserto
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
      ExecutaPreValidacao                         'habilita/desabilita campos
      chkCampo(Index).Change
   End If
   Select Case Index
      Case 0
         TrocaPaConserto  
   End Select
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
   Caption = LoadGasString(23290)
   vgFormID = 639
   vgIdentTab = "Movimento do estoque"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia do movimento"
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
   vgUltimoTabIndex = 11
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(14795))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(14796))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(14797))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(14798))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(14799))
   Set Botao(4).Picture = LoadPicture(LoadGasPicture(14800))
   Set Botao(5).Picture = LoadPicture(LoadGasPicture(14801))
   Set lblMovimento = Label(0)
   Set lblAjuste = Label(1)
   Set txtSequencia = txtCampo(1)
   Set GrdProdutos = Grid(0)
   Set lblParcelamento = Label(11)
   Set grdParcelamento = Grid(1)
   Set cmdImportar = Botao(1)
   Set cmdFechar = Botao(2)
   Set cmdAbrir = Botao(3)
   Set lblFechamento = Label(18)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia do movimento] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(23295)
   Label(1).Caption = LoadGasString(23296)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(23297)
   Label(2).Caption = LoadGasString(23298)
   Label(3).Caption = LoadGasString(23299)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(23300)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(23301)
   Label(4).Caption = LoadGasString(23302)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(23303)
   Tabs(0).TabCaption(0) = LoadGasString(23304)
   Tabs(0).TabCaption(1) = LoadGasString(23305)
   Label(5).Caption = LoadGasString(23306)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(23307)
   Label(6).Caption = LoadGasString(23308)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(23309)
   Label(7).Caption = LoadGasString(23310)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(23311)
   Label(8).Caption = LoadGasString(23312)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(23313)
   Label(9).Caption = LoadGasString(23314)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(23315)
   Label(10).Caption = LoadGasString(23316)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(23317)
   Label(12).Caption = LoadGasString(23318)
   Label(13).Caption = LoadGasString(23319)
   Label(14).Caption = LoadGasString(23320)
   Label(15).Caption = LoadGasString(23321)
   Label(16).Caption = LoadGasString(23322)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(23323)
   Label(17).Caption = LoadGasString(23324)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(23325)
   Botao(1).Caption = LoadGasString(23326)
   vgTooltips.AddTool Botao(2), 0, LoadGasString(23327)
   Botao(2).Caption = LoadGasString(23328)
   vgTooltips.AddTool Botao(3), 0, LoadGasString(23329)
   Botao(3).Caption = LoadGasString(23330)
   Label(18).Caption = LoadGasString(23331)
   Botao(4).Caption = LoadGasString(23332)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(23333)
   chkCampo(0).Caption = LoadGasString(23334)
   Label(19).Caption = LoadGasString(23335)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(23336)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Produtos movimento estoque", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Produtos movimento estoque", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos movimento estoque", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Sequencia do produto", TP_NUMERICO, "", , False, , "DADOSNFE", "Produtos", "Sequencia do produto", "Descrição; Sequencia do produto", "Descrição; Sequencia do produto", "Descrição", "", , "1", "Produtos.[Sequencia do produto]", "", "DADOSNFE", "18", 1, "0", 5145
      .AddColumn Nothing, , "Nº Item:", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 870
      .AddColumn Nothing, , "Un.", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1320
      .AddColumn Nothing, , "Estoque", , TP_NUMERICO, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Vr. Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1425
      .AddColumn Nothing, , "Tot. Produtos", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1500
      .AddColumn Nothing, , "% IPI", "Aliquota do ipi", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Valor IPI", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1350
      .AddColumn Nothing, , "Vr. ICMS ST", "Valor icms st", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1440
   End With
   With Grid(1)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Parcelas movimento estoque", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas movimento estoque", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas movimento estoque", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 660
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 540
      .AddColumn Nothing, , "Dt. de Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1545
      .AddColumn Nothing, , "Vr. da Parcela", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1455
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Movimento_do_estoque = vgTb
   vgPriVez = False
   Reposition
   CarregaTotalizador  
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

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@!"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

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
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

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
   txtCampo(4).ListFields = "Razão social; Sequencia do geral"
   txtCampo(4).OrderFields = "Razão social"
   txtCampo(4).Relation = ""
   txtCampo(4).Source = "Geral"
   txtCampo(4).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(13).CtPri = txtCp(13)
   Set txtCampo(13).CtFdo = labFdo13
   Set txtCampo(13).CtBot(BOT_LISTA) = bottxtCampo13(BOT_LISTA)
   Set txtCampo(13).CtBot(BOT_COMBO) = bottxtCampo13(BOT_COMBO)
   bottxtCampo13(BOT_LISTA).Caption = "P"
   Set bottxtCampo13(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "@x"
   txtCampo(13).PesqModoAbertura =  2 
   txtCampo(13).PesqFieldCapture = "[Centro de Custo].[Descrição do Centro]"
   txtCampo(13).BoundColumn = "Descrição do Centro"
   txtCampo(13).ListFields = "Descrição do Centro"
   txtCampo(13).OrderFields = "Descrição do Centro"
   txtCampo(13).Relation = ""
   txtCampo(13).Source = "Centro de Custo"
   txtCampo(13).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "999999"
   txtCampo(1).Editable = False
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 4
   txtCampo(12).Mask = ""
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(5).CtPri = txtCp(5)
   Set txtCampo(5).CtFdo = labFdo5
   Set txtCampo(5).CtBot(BOT_COMBO) = bottxtCampo5(BOT_COMBO)
   Set bottxtCampo5(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(5).DataType = 0
   txtCampo(5).ListFields = "Vista|Prazo"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_ACAO) = bottxtCampo6(BOT_ACAO)
   Set bottxtCampo6(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = "99.999.999,99"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_ACAO) = bottxtCampo7(BOT_ACAO)
   Set bottxtCampo7(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "99.999.999,99"
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_ACAO) = bottxtCampo8(BOT_ACAO)
   Set bottxtCampo8(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(8).DataType = 1
   txtCampo(8).Mask = "99.999.999,99"
   txtCampo(8).Editable = False
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   Set txtCampo(9).CtFdo = labFdo9
   Set txtCampo(9).CtBot(BOT_LISTA) = bottxtCampo9(BOT_LISTA)
   Set bottxtCampo9(BOT_LISTA).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "99.999.999,99"
   txtCampo(9).Editable = False
   txtCampo(9).PesqModoAbertura =  2 
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   Set txtCampo(10).CtFdo = labFdo10
   Set txtCampo(10).CtBot(BOT_ACAO) = bottxtCampo10(BOT_ACAO)
   Set bottxtCampo10(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = "99.999.999,99"
   txtCampo(10).Editable = False
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   Set txtCampo(11).CtFdo = labFdo11
   Set txtCampo(11).CtBot(BOT_ACAO) = bottxtCampo11(BOT_ACAO)
   Set bottxtCampo11(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "999.999.999,99"
   txtCampo(11).Editable = False
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

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
      Case 4
         txtCampo(4).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0", "[Sequencia do geral] > 0") & ""
                           txtCampo(4).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                              "Geral.Bairro, Geral.Fone, Geral.Celular, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ")"
      Case 13
         txtCampo(13).Filter = "[Codigo do Centro] > " & 0 & ""
                           txtCampo(13).PesqSQLExpression = "SELECT [Centro de Custo].[Codigo do Centro], [Centro de Custo].[Descrição do Centro] FROM [Centro de Custo] WHERE ([Centro de Custo].[Codigo do Centro] > " + CStr(0) + ")"
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
   MudaTamCampos Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   lblAlteracao.Caption = ""
   FinalizaForm Me
   Set lblMovimento = Nothing
   Set lblAjuste = Nothing
   Set txtSequencia = Nothing
   Set GrdProdutos = Nothing
   Set lblParcelamento = Nothing
   Set grdParcelamento = Nothing
   Set cmdImportar = Nothing
   Set cmdFechar = Nothing
   Set cmdAbrir = Nothing
   Set lblFechamento = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   If Not Movimento_do_estoque Is Nothing Then
      Set Movimento_do_estoque = Nothing
   End If
   If Not Produtos_movimento_estoqu Is Nothing Then
      Produtos_movimento_estoqu.CloseRecordSet
      Set Produtos_movimento_estoqu = Nothing
   End If
   If Not Parcelas_movimento_estoqu Is Nothing Then
      Parcelas_movimento_estoqu.CloseRecordSet
      Set Parcelas_movimento_estoqu = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmEstoque = Nothing                       'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no grid filho
Private Sub Grid_KeyPress(Index As Integer, ByVal KeyAscii As Integer, ByVal Shift As Integer, vgColumns() As Variant)
   Select Case Index
      Case 0
         ExecutaGrid0 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
   End Select
End Sub

'evento - está mudando a linha selecionada do grid
Private Sub Grid_SkipRecord(Index As Integer, vgColumns() As Variant, ByVal vgBookMark As Variant)
   ExecutaGrid Index, vgColumns(), CONDICOES_ESPECIAIS
End Sub

'evento - após efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 1
   GeraLog Me, Grid(Index).Status, Index, False
   End Select
End Sub

'evento - antes de efetuar o edit do recordset do grid
Private Sub Grid_BeforeEditRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, Grid(Index).Status, Index, True
   Select Case Index
      Case 0
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 1
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Produtos_movimento_estoqu![Sequencia do movimento] = Movimento_do_estoque![Sequencia do movimento]
      Case 1
         Parcelas_movimento_estoqu![Sequencia do movimento] = Movimento_do_estoque![Sequencia do movimento]
   End Select
End Sub

'evento - antes de efetuar o delete no recordset do grid
Private Sub Grid_BeforeDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, True
   ExecutaGrid Index, vgColumns(), EXCLUSOES, vgItem, 0, vgIsValid, 0, vgErrorMessage
   Select Case Index
      Case 0
      Case 1
   End Select
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
   Select Case Index
      Case 0
         Set Produtos_movimento_estoqu = vgNewRecordSet
      Case 1
         Set Parcelas_movimento_estoqu = vgNewRecordSet
   End Select
End Sub

'evento - quer validar dados, está gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - após a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
   End Select
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
      Case  1 
         ExecutaGrid = ExecutaGrid1(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_produto As Double, Quantidade As Double, Valor_unitario As Double
   Dim Valor_total As Double, Valor_do_ipi As Double, Valor_icms_st As Double
   Dim Aliquota_do_ipi As Single, Sequencia_do_item As Long
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
      Select Case vgCol
         Case 3
            vgRetVal = (InfoProdutos(Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item, "Sigla"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 4
            vgRetVal = (InfoProdutos(Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item, "Estoque"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 9
            vgRetVal = ((Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2)))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      Comandos KeyCodeAscii, Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item 
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(0).AllowEdit = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(0).AllowDelete = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Produtos movimento estoque]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do movimento] = " & Movimento_do_estoque![Sequencia do movimento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

      'vamos definir a ordenação
      x$ = "[Sequencia do item]"
      vgRetVal = InsereSQL(vgRetVal, EXP_ORDERBY, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 2
            Sequencia_do_item = SeqItem()
            vgRetVal = Sequencia_do_item
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" + FiltroProduto() + ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Produtos.[Sequencia do produto], Produtos.Descrição, Produtos.Estoque, Produtos.[Valor total], " + _
                                                  "Produtos.[Material adquirido de terceiro], Produtos.Importado FROM Produtos WHERE (Produtos.[Sequencia do produto] > " + CStr(0) + ") AND " + _
                                                  "(Produtos.Inativo = " + CStr(0) + ")"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_produto = 0 Then
            vgIsValid = (Sequencia_do_produto > 0)
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "Produto não pode ser em Branco"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_produto > 0 Then
               vgIsValid = (ValidaProduto(Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item))
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "Produto Inativo para Inclui - lo Abra o Cadastro do Item e Desmarque a Opção (INATIVO)!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn =  5 
            vgErrorMessage$ = "Quantidade deve ser maior que Zero(0)!!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_unitario > 0)
            If Not vgIsValid Then vgColumn =  6 
            vgErrorMessage$ = "Valor unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(2) = SeqItem()
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AtuEstoque Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item 
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AtuEstoque Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AtuEstoque Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item 
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Movimento_do_estoque.Edit
         Set vgRsError = Movimento_do_estoque
         If ProcessaProdutos(Sequencia_do_movimento, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, Valor_do_ipi, _
   Valor_icms_st, Aliquota_do_ipi, Sequencia_do_item) Then
            Movimento_do_estoque![Sequencia do movimento] = (0)
         End If
         Movimento_do_estoque.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Sequencia_do_item = ColumnValue(2)
   Quantidade = ColumnValue(5)
   Valor_unitario = ColumnValue(6)
   Aliquota_do_ipi = ColumnValue(8)
   Valor_icms_st = ColumnValue(10)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Produtos_movimento_estoqu.Eof = False And Produtos_movimento_estoqu.BOF = False And Produtos_movimento_estoqu.RecordCount > 0 Then
         Valor_total = Produtos_movimento_estoqu![Valor total]
         Valor_do_ipi = Produtos_movimento_estoqu![Valor do ipi]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_movimento_estoqu.Eof = False And Produtos_movimento_estoqu.BOF = False And Produtos_movimento_estoqu.RecordCount > 0 Then
      Sequencia_do_produto = Produtos_movimento_estoqu![Sequencia do produto]
      Quantidade = Produtos_movimento_estoqu!Quantidade
      Valor_unitario = Produtos_movimento_estoqu![Valor unitario]
      Valor_total = Produtos_movimento_estoqu![Valor total]
      Valor_do_ipi = Produtos_movimento_estoqu![Valor do ipi]
      Valor_icms_st = Produtos_movimento_estoqu![Valor icms st]
      Aliquota_do_ipi = Produtos_movimento_estoqu![Aliquota do ipi]
      Sequencia_do_item = Produtos_movimento_estoqu![Sequencia do item]
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
   If Movimento_do_estoque.EOF = False And Movimento_do_estoque.BOF = False Then
      Movimento_do_estoque.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
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
      If Grid(1).Status <> ACAO_NAVEGANDO And vgItem = Grid(1).SelectedItem Then
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
            vgRetVal = Not (((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0) AND Movimento_do_estoque![Forma de pagamento] = "Prazo")
         Case 3
            vgRetVal = Not (((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0) AND Movimento_do_estoque![Forma de pagamento] = "Prazo")
         Case 4
            vgRetVal = Not ((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0)
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (((Movimento_do_estoque!Cancelado = 0) AND Not Vazio(Movimento_do_estoque![Forma de pagamento]) And TotalParcelas < Movimento_do_estoque![Valor total do movimento]) AND Movimento_do_estoque!Fechado = 0)
      On Error Resume Next
         Grid(1).AllowEdit = ((((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Sequencia da nota fiscal] = 0) AND Movimento_do_estoque!Fechado = 0) AND Movimento_do_estoque!Cancelado = 0)
      On Error Resume Next
         Grid(1).AllowDelete = ((((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Sequencia da nota fiscal] = 0) AND Movimento_do_estoque!Fechado = 0) AND Movimento_do_estoque!Cancelado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas movimento estoque]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia do movimento] = " & Movimento_do_estoque![Sequencia do movimento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D", Movimento_do_estoque![Data de entrada] , Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Movimento_do_estoque![Data de entrada])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Movimento_do_estoque![Valor total do movimento] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If ((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0) AND Movimento_do_estoque![Forma de pagamento] = "Prazo" Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If ((Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0) AND Movimento_do_estoque![Forma de pagamento] = "Prazo" Then
               vgIsValid = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
               If Not vgIsValid Then vgColumn =  3 
               vgErrorMessage$ = "Data de Vencimento Invalida!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Movimento_do_estoque![Forma de pagamento])) AND Movimento_do_estoque![Entrada Conserto] = 0 Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Movimento_do_estoque![Forma de pagamento] = "Prazo" , TotalParcelas() + Valor_da_Parcela <= Movimento_do_estoque![Valor total do movimento], Valor_da_parcela = Movimento_do_estoque![Valor total do movimento]), IIf(Movimento_do_estoque![Forma de pagamento] = "Vista" , Valor_da_parcela = Movimento_do_estoque![Valor total do movimento], Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Movimento_do_estoque![Valor total do movimento])))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total do Movimento"
            End If
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = UltimaParcela
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
   If Parcelas_movimento_estoqu.Eof = False And Parcelas_movimento_estoqu.BOF = False And Parcelas_movimento_estoqu.RecordCount > 0 Then
      Numero_da_parcela = Parcelas_movimento_estoqu![Numero da parcela]
      Dias = Parcelas_movimento_estoqu!Dias
      Data_de_vencimento = Parcelas_movimento_estoqu![Data de vencimento]
      Valor_da_parcela = Parcelas_movimento_estoqu![Valor da parcela]
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
   ExecutaGrid1 = vgRetVal
   vgPriVez = False
End Function

'evento - pega o valor inicial das colunas do grid
Private Sub Grid_GetColumnDefaultValue(Index As Integer, ByVal vgCol As Integer, vgColumns() As Variant, ByRef vgDefaultValue As Variant)
   vgDefaultValue = ExecutaGrid(Index, vgColumns(), DEFAULTDASCOLUNAS, , vgCol)
End Sub

'evento - quer pegar valores para cada célula
Private Sub Grid_GetColumnLocked(Index As Integer, ByVal vgRow As Long, ByVal vgCol As Long, vgColumns() As Variant, ByRef FormField As FormataCampos, ByRef vgLocked As Boolean)
   Select Case Index
      Case 0
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 1
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
   End Select
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
      Case 2
         Fechamento  
      Case 3
         AbreMvto  
      Case 4
         AbreGerar  
      Case 5
         frmCentroC.vgFiltroInicial = "[Descrição do Centro] = '" & Tipo_de_conta & "'
         Load frmCentroC
         frmCentroC.vgPriVez = False
         InicializaFiltro frmCentroC
         frmCentroC.Reposition
         mdiNFE.MCentroC
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo4_Click(Index As Integer)
   txtCampo(4).SetFocus
   DoEvents
   txtCampo(4).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo13_Click(Index As Integer)
   txtCampo(13).SetFocus
   DoEvents
   txtCampo(13).botClick Index
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
Private Sub bottxtCampo7_Click(Index As Integer)
   txtCampo(7).SetFocus
   DoEvents
   txtCampo(7).botClick Index
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
Private Sub bottxtCampo10_Click(Index As Integer)
   txtCampo(10).SetFocus
   DoEvents
   txtCampo(10).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo11_Click(Index As Integer)
   txtCampo(11).SetFocus
   DoEvents
   txtCampo(11).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


