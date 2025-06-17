Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMoSaidas
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8145
   Left                     =   2460
   LinkTopic                =   "MoSaidas"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7710
   ScaleWidth               =   13605
   Top                      =   1930
   Width                    =   13695
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
      Height                   =   7605
      Left                     =   15
      Width                    =   13545
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   23879
      _extenty                 =   13407
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   13020
         Width                    =   45
         Top                      =   7680
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
         Index                    =   0 
      End
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1545
         Left                     =   180
         Width                    =   1380
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
         Left                     =   1800
         Width                    =   2985
         Top                      =   480
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
         Caption                  =   "Saidas - Estoque"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   240
         Width                    =   1020
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
         Caption                  =   "Sequencia:"
         Index                    =   2 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000005
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
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   300
         Left                     =   1380
         Width                    =   840
         Top                      =   1740
         DataField                =   "Sequencia da saida"
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
         Left                     =   2400
         Width                    =   960
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
         Caption                  =   "Doc / NFe:"
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
         MaxLength                =   25
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   1
         TabStop                  =   -1
         Height                   =   300
         Left                     =   3540
         Width                    =   1620
         Top                      =   1740
         DataField                =   "Documento"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   5340
         Width                    =   1125
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
         Caption                  =   "Dt. Emissão:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6600
         Top                      =   1740
         Width                    =   1380
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   7710
         Top                      =   1755
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
         Left                     =   6615
         Width                    =   1080
         Top                      =   1755
         DataField                =   "Data de emissão"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8160
         Width                    =   540
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
         Caption                  =   "Geral:"
         Index                    =   5 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8850
         Top                      =   1740
         Width                    =   4230
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   12810
         Top                      =   1755
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   12555
         Top                      =   1755
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
         Left                     =   8865
         Width                    =   3675
         Top                      =   1755
         DataField                =   "Sequencia do geral"
         Index                    =   3 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   2
         TabHeight                =   519
         BackColor                =   &HE0E0E0
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   2
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   3255
         Left                     =   180
         Width                    =   12870
         Top                      =   2460
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
         _extentx                 =   22689
         _extenty                 =   5738
         _stockprops              =   15
         TabCaption(0)            =   "1 - Produtos"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "2 - Financeiro"
         Tab(1).ControlCount=   23
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Label(6)"
         Tab(1).Control(1)=   "labFdo4"
         Tab(1).Control(2)=   "bottxtCampo4(2)"
         Tab(1).Control(3)=   "txtCp(4)"
         Tab(1).Control(4)=   "Label(7)"
         Tab(1).Control(5)=   "txtCp(5)"
         Tab(1).Control(6)=   "Label(8)"
         Tab(1).Control(7)=   "txtCp(6)"
         Tab(1).Control(8)=   "Label(9)"
         Tab(1).Control(9)=   "txtCp(7)"
         Tab(1).Control(10)=   "Label(10)"
         Tab(1).Control(11)=   "txtCp(8)"
         Tab(1).Control(12)=   "Label(11)"
         Tab(1).Control(13)=   "txtCp(9)"
         Tab(1).Control(14)=   "Grid(1)"
         Tab(1).Control(15)=   "Label(12)"
         Tab(1).Control(16)=   "Label(13)"
         Tab(1).Control(17)=   "Formato(0)"
         Tab(1).Control(18)=   "Label(14)"
         Tab(1).Control(19)=   "Formato(1)"
         Tab(1).Control(20)=   "Label(15)"
         Tab(1).Control(21)=   "Formato(2)"
         Tab(1).Control(22)=   "Label(16)"
         Begin NFE.GListV Grid
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
            TabIndex                 =   6
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2655
            Left                     =   120
            Width                    =   12600
            Top                      =   420
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
            Left                     =   -74700
            Width                    =   2010
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
            Caption                  =   "Forma de Pagamento:"
            Index                    =   6 
         End
         Begin VB.Label labFdo4
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -72540
            Top                      =   660
            Width                    =   1470
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo4
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -71340
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
            Alignment                =   0
            TabIndex                 =   8
            TabStop                  =   -1
            Height                   =   270
            Left                     =   -72525
            Width                    =   1185
            Top                      =   675
            DataField                =   "Forma de pagamento"
            Index                    =   4 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -74010
            Width                    =   1320
            Top                      =   1260
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
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   9
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72540
            Width                    =   1470
            Top                      =   1260
            DataField                =   "Valor do frete"
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
            Left                     =   -74415
            Width                    =   1725
            Top                      =   1860
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
            Caption                  =   "Valor do Desconto:"
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
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   10
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72540
            Width                    =   1470
            Top                      =   1860
            DataField                =   "Valor do desconto"
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
            Left                     =   -74460
            Width                    =   1755
            Top                      =   2460
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
            Index                    =   9 
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
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   11
            TabStop                  =   0
            Height                   =   300
            Left                     =   -72540
            Width                    =   1470
            Top                      =   2460
            DataField                =   "Valor total dos produtos"
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
            Height                   =   240
            Left                     =   -70845
            Width                    =   1740
            Top                      =   1260
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
            Caption                  =   "Total Icms S.Trib....:"
            Index                    =   10 
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
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   12
            TabStop                  =   0
            Height                   =   300
            Left                     =   -68940
            Width                    =   1470
            Top                      =   1260
            DataField                =   "Valor total icms st"
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
            Left                     =   -70860
            Width                    =   1740
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
            Caption                  =   "Total de IPI...............:"
            Index                    =   11 
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
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   13
            TabStop                  =   0
            Height                   =   300
            Left                     =   -68940
            Width                    =   1470
            Top                      =   660
            DataField                =   "Valor total do ipi"
            TabStop                  =  0
            Index                    =   9 
            MultiLine                =   -1
         End
         Begin NFE.GListV Grid
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
            TabIndex                 =   14
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2160
            Left                     =   -67200
            Width                    =   4950
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   0
            BackColor                =   &H0
            BackStyle                =   1
            BorderStyle              =   1
            AutoSize                 =   0
            Height                   =   255
            Left                     =   -67185
            Width                    =   4965
            Top                      =   375
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
            Caption                  =   "Parcelamento do Contas á Receber"
            Index                    =   12 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -67200
            Width                    =   600
            Top                      =   2955
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
            Index                    =   13 
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
            Left                     =   -66525
            Width                    =   165
            Top                      =   3000
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -66300
            Width                    =   990
            Top                      =   2955
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
            Index                    =   14 
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
            Left                     =   -65205
            Width                    =   165
            Top                      =   3000
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -64905
            Width                    =   1050
            Top                      =   2955
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
            Index                    =   15 
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
            Left                     =   -63690
            Width                    =   150
            Top                      =   3000
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -63465
            Width                    =   300
            Top                      =   2955
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
            Index                    =   16 
         End
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
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   870
         Left                     =   180
         Width                    =   12900
         Top                      =   6405
         DataField                =   "Historico"
         Index                    =   10 
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
         Left                     =   8640
         Width                    =   2490
         Top                      =   5940
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
         Index                    =   17 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &HE0E0E0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         BorderStyle              =   1
         Alignment                =   1
         Locked                   =   -1
         TabIndex                 =   7
         TabStop                  =   0
         Height                   =   300
         Left                     =   11220
         Width                    =   1830
         Top                      =   5940
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
         AutoSize                 =   -1
         Height                   =   375
         Left                     =   1800
         Width                    =   1005
         Top                      =   960
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
         Caption                  =   "Status"
         Index                    =   18 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   1
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   15
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   300
         Left                     =   5055
         Width                    =   1185
         Top                      =   1020
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
         Index                    =   0 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   16
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   10470
         Width                    =   1290
         Top                      =   135
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
         Index                    =   1 
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
         Left                     =   11790
         Width                    =   1290
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
         Caption                  =   "Gerar"
         Index                    =   2 
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
Attribute VB_name = "frmMoSaidas"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: MoSaidas
'* Função....: Saidas - Estoque
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:40
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
Dim txtCampo(11) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Sequencia_da_saida As Long, Documento As String, Sequencia_do_geral As Long
Dim Data_de_emissao As Variant, Forma_de_pagamento As String, Historico As String
Dim Cancelado As Boolean, Valor_total_dos_produtos As Double, Valor_total_do_ipi As Double
Dim Valor_total_icms_st As Double, Sequencia_da_nota_fiscal As Long, Fechado As Boolean
Dim Valor_do_frete As Double, Valor_do_desconto As Double, Valor_total_do_movimento As Double
Dim Usuario_da_Alteracao As String, Data_da_Alteracao As Variant, Hora_da_Alteracao As Variant
Dim Teste As String
Public lblAjuste As Object, lblMovimento As Object, grdProdutos As Object
Public grdParcelamento As Object, lblParcelamento As Object, lblFechamento As Object
Public cmdAbrir As Object, cmdFechar As Object
Dim Outras_saidas_estoque As New GRecordSet, Produtos_saidas_estoque As New GRecordSet, Parcelas_outras_saidas As New GRecordSet

Dim ProdutoAux As New GRecordSet, UnidadeAux As New GRecordSet

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
    Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia da Saida]) Seq From [Outras Saidas Estoque]")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function 

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Function SeqItem() As Long 'Grid Sequencial do Gas é ruim
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)
 
    If Tb.RecordCount > 0 Then
      SeqItem = Tb!Seq + 1
    Else
      SeqItem = 1
    End If
     
End Function


Private Function FiltroProduto() As String

  On Error Resume Next ' Se der Erro que si foda...
  FiltroProduto = "[Sequencia do produto] > 0"
  If vgSituacao <> ACAO_NAVEGANDO Then
     FiltroProduto = FiltroProduto & " And [Sequencia do produto] > 0 AND Inativo = 0"
  End If 
    
End Function


Private Function ValidaProduto(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single) As Boolean
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


Private Function InfoProdutos(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, Oq As String) As Variant
   
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


Private Sub CarregaTotalizador()
   On Error Resume Next

   With grdProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Vr. Unitário").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowSumCol(.Columns("Vr. ICMS ST").Index) = True
      .ShowSumCol(.Columns("Vr. IPI").Index) = True
      .ShowSumCol(.Columns("Aliquota do ipi").Index) = False
      .ShowSumCol(.Columns("Sequencia do Item").Index) = False
      '.ColumnWidth(CInt(.Columns("Sequencia do item").Index)) = 0
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


Private Sub RepositionEstoque()
   On Error Resume Next
   
   If vgSituacao = ACAO_INCLUINDO Then
      Set Produtos_saidas_estoque = Nothing
      Set Parcelas_outras_saidas = Nothing
   End If
   
   CarregaTotalizador
   'Alteração
   lblAlteracao.Caption = Outras_saidas_estoque![Usuario da Alteração] & " " & Outras_saidas_estoque![Data da Alteração] & " " & Outras_saidas_estoque![Hora da Alteração]
  
End Sub


Private Function ProcessaProdutos(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single)
         
   On Error GoTo DeuErro
   
   vgDb.BeginTrans  
       
   vgDb.Execute "Update [Produtos saidas estoque] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida & " AND [Sequencia do item] = " & Sequencia_do_item
   If Aliquota_do_ipi > 0 Then
      vgDb.Execute "Update [Produtos saidas estoque] Set [Valor do ipi] = " & Substitui(Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2), ",", ".", SO_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida & " AND [Sequencia do item] = " & Sequencia_do_item 'Valor do IPI
   Else
      vgDb.Execute "Update [Produtos saidas estoque] Set [Valor do ipi] = 0 WHERE [Sequencia da saida] = " & Sequencia_da_saida & " AND [Sequencia do item] = " & Sequencia_do_item
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
   
   Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)
          
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT ROUND(SUM([Valor do ipi]), 2) Total FROM [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida) 'IPI dos Produtos
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)'Produtos Novos
   Set ValorICMSST = vgDb.OpenRecordSet("SELECT SUM([Valor icms st]) Total FROM [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida) 'Substituição Produtos                                                                                             
   ValorMvto = IPIProdutos!Total + ValorProdutos!Total + ValorICMSST!Total + Valor_do_Frete - Valor_do_desconto
   ValorMvto = Format(ValorMvto, "##,###,##0.00")
   
   If Produtos.RecordCount = 0 Then
      ValorMvto = 0 
      If Valor_do_frete > 0 Or Valor_do_desconto > 0 Then 
         vgDb.BeginTrans   
         vgDb.Execute "Update [Outras saidas estoque] Set [Valor do frete] = 0 WHERE [Sequencia da saida] = " & Sequencia_da_saida
         vgDb.Execute "Update [Outras saidas estoque] Set [Valor do desconto] = 0 WHERE [Sequencia da saida] = " & Sequencia_da_saida
         vgDb.CommitTrans
      End If  
   End If

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Outras saidas estoque] Set [Valor total do ipi] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida 'IPI Produtos
   vgDb.Execute "Update [Outras saidas estoque] Set [Valor total dos produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida 'Produtos 
   vgDb.Execute "Update [Outras saidas estoque] Set [Valor total icms st] = " & Substitui(ValorICMSST!Total, ",", ".", SO_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida 'Substituição
   vgDb.Execute "Update [Outras saidas estoque] Set [Valor total do movimento] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Sequencia da saida] = " & Sequencia_da_saida 'Valor da Entrada
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


Private Function AtuEstoque(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single)
Dim Tb As New GRecordSet

  If vgSITUACAO = -ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)
   If Tb.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Produtos saidas estoque] WHERE [Sequencia da saida] = " & Sequencia_da_saida & " And [Sequencia do item] = " & Sequencia_do_item
   End If
  End If 
  
   AtualizaEstoque(Sequencia_do_produto)
   UltimoMvto(Sequencia_do_produto)
   UltimoFornecedor(Sequencia_do_produto)
   ValorUnitario(Sequencia_do_produto)
   ValorCustoProdutos(Sequencia_do_produto)
   
   AjustaValores
   
End Function


Private Sub AlteraManutencao() 'Qdo altera Registro Mãe
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet
   
   On Error GoTo DeuErro
   
   AjustaValores
            
   Set Parcelas = Parcelas_outras_saidas
   
   If Parcelas.RecordCount = 0 Then Exit Sub
         
   vgDb.BeginTrans
   
    Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da Saida] = " & Sequencia_da_saida)
    If Manutencao.RecordCount = 0 Then Exit Sub
    Do While Not Manutencao.EOF
       With Manutencao
          ![Sequencia Do geral] = Outras_saidas_estoque![Sequencia Do geral]
          ![Data de entrada] = Outras_saidas_estoque![Data de emissão]
          !Historico = Historico
          ![Forma de pagamento] = Outras_saidas_estoque![Forma de pagamento]
          ![Valor total] = Outras_saidas_estoque![Valor total Do movimento]
          !Documento = Outras_saidas_estoque!Documento
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
      lblMovimento.Caption = "Movimento Cancelado"
   Else
      lblMovimento.ForeColor = &H0
      lblMovimento.Caption = "Saidas - Estoque"
   End If
   If Err Then Err = 0: lblMovimento.Caption = "" 
   If Err Then Err.Clear 
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   
   'Label Do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_total_do_movimento) Or (TotalParcelas > Valor_total_do_movimento) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   
End Sub


Private Sub Comandos(KeyAscii As Integer, Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single)
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


Private Function VerificaDocumento() As Boolean
   Dim Tabela As New GRecordSet
   
   If Vazio(Documento) Then VerificaDocumento = False: Exit Function
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then VerificaDocumento = True: Exit Function
   If Sequencia_do_geral = 0 Then VerificaDocumento = True: Exit Function
      
   Set Tabela = vgDb.OpenRecordSet("Select [Sequencia da saida], documento, [Sequencia do geral] From [Outras saidas estoque] " & _
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
   If Tabela![Sequencia Do Geral] = Sequencia_do_geral And Tabela!Documento = Documento And Tabela![Sequencia da saida] <> Sequencia_da_saida Then
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


Private Sub DeletaItens()' Qdo excluir Registro Mae
Dim Tb As New GRecordSet, Parcelas As New GRecordSet

  If vgSITUACAO = ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)
  Set Parcelas = vgDb.OpenRecordSet("SELECT * From [Parcelas outras saidas] Where [Sequencia da saida] = " & Sequencia_da_saida)
  
   If Parcelas.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Parcelas outras saidas] Where [Sequencia da saida] = " & Sequencia_da_saida
      DeletaConta
   End If
  
   If Tb.RecordCount > 0 Then
      Do While Not Tb.EOF
      vgDb.Execute "DELETE FROM [Produtos saidas estoque] WHERE [Sequencia da saida] = " & Sequencia_da_saida & " And [Sequencia do produto] = " & Tb![Sequencia Do produto]
      AtualizaEstoque(Tb![Sequencia Do produto])
      Tb.MoveNext
      Loop
      Tb.MoveFirst
   End If
  End If   
End Sub


Private Function ValidaEstoque(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single) As Boolean ' Não deixa dar saida em estoque < 0
 
   On Error Resume Next
     
   If vgSituacao = -ACAO_INCLUINDO Then
      ValidaEstoque = (ProdutoAux!Estoque - Quantidade) >= 0
   Else
      ValidaEstoque = (ProdutoAux!Estoque + Produtos_saidas_estoque!Quantidade - Quantidade) >= 0
   End If
   If Not ValidaEstoque Then
      MsgBox "Estoque Insuficiente Digite a Senha do Estoque para Estourar", vbCritical + vbOKOnly, vaTitulo
      InputSenha
      If Not Vazio(InputYgorsBom) Then
         If InputYgorsBom = Empresa![Senha Admin] Then
            ValidaEstoque = True
         Else
            MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
            ValidaEstoque = False
         End If
         InputYgorsBom = ""
      End If
   End If
   Exit Function
  
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


Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC From [Parcelas outras saidas] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function

Private Function UltimoVencimento() As Variant
 Dim Tb As New GRecordSet, RetVal As Variant
   
  On Error Resume Next
      
  Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de vencimento]) [Data de Vencimento] FROM [Parcelas outras saidas] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
      
  RetVal = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, grdParcelamento.Col - 2)
  
   
   If Outras_saidas_estoque![Forma de pagamento] = "Prazo" Then
      If Not IsNull(Tb![Data de Vencimento]) Then
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Outras_saidas_estoque![Data de emissão]
      End If
   Else
      UltimoVencimento = RetVal = Outras_saidas_estoque![Data de emissão]
   End If
    
End Function

Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas outras saidas] Where [Sequencia da saida] = " & Sequencia_da_saida & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


Private Sub DeletaConta() 'Qdo exclui registro Mãe
   Dim Manutencao As New GRecordset

   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
    
   If Manutencao.RecordCount = 0 Then Exit Sub
   
   vgDb.Execute "DELETE FROM [Manutenção contas receber] " & _
                "WHERE [Sequencia da saida] = " & Sequencia_da_saida   

DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbokOnly, vaTitulo
      vgdb.RollBackTrans
   End If

End Sub


Private Sub LancaParcelas()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
   Dim SeqBaixa As Long
   
   On Error GoTo DeuErro
      
   If Vazio(Forma_de_pagamento) Then Exit Sub
   
   If TotalParcelas() <> Valor_total_do_movimento Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas receber")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas outras saidas] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas receber", "[Sequencia da saida] = " & Sequencia_da_saida & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew
         !Parcela = Parcelas![Numero da parcela]
         ![Sequencia Do geral] = Sequencia_do_Geral
         !Documento = Documento
         ![Data de entrada] = Data_de_emissao
         !Historico = Historico
         ![Forma de Pagamento] = Forma_de_pagamento
         ![Especie de pagamento] = Parcelas![Especie de pagamento]
          If ![Forma de pagamento] = "Vista" And Empresa![Baixa Receber Auto] Then
             ![Valor pago] = Parcelas![Valor da parcela]
             ![Valor restante] = 0
             ![Data da baixa] = Data_de_emissao
          Else
             ![Valor pago] = 0
             ![Valor restante] = Parcelas![Valor da parcela]
             ![Data da baixa] = Null
          End If
         ![Data de vencimento] = Parcelas![Data de vencimento]
         ![Valor total] = Valor_total_do_movimento
         ![Valor da parcela] = Parcelas![Valor da parcela]
         ![Sequencia da saida] = Sequencia_da_saida
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


Private Sub Fechamento()
 Dim Produtos As New GRecordSet, Parcelas As New GRecordSet
 Dim Manutencao As New GRecordSet
 
 On Error Resume Next
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos saidas estoque] Where [Sequencia da saida] = " & Sequencia_da_saida)
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas outras saidas] Where [Sequencia da saida] = " & Sequencia_da_saida)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
  
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
  
  If Parcelas.RecordCount = 0 Then If MsgBox("Deseja Fechar o Movimento sem Incluir as Parcelas do Contas a Receber?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
     vgDb.Execute "UPDATE [Outras saidas estoque] Set Fechado = 1 WHERE [Sequencia da saida] = " & Sequencia_da_saida
  Reposition True
              
End Sub


Private Sub AbreMvto()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean, TemDevolucao As Boolean
 Dim Manu As New GRecordSet, Devolucao As New GRecordSet

 On Error Resume Next
 
 Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas receber] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas receber] WHERE [Sequencia da saida] = " & Sequencia_da_saida & ")")
 Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas receber] Where [Sequencia da saida] = " & Sequencia_da_saida)
 'Set Devolucao = vgDb.OpenRecordSet("SELECT * FROM [Devolução de compra] WHERE Cancelado = 0 And [Sequencia do movimento] IN (SELECT [Sequencia do movimento] FROM [Movimento do estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & ")")
  
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If
   
    'If Devolucao.RecordCount > 0 Then
    '   TemDevolucao = True
    'Else
    '   TemDevolucao = False
    'End If
   
   If TemBaixa Then
      MsgBox("Impossivel abrir o Movimento existe Baixa(s) no Contas a Receber para Abri-lo Exclua a(s) Baixa(s)")
      Exit Sub
   End If
   
   'If TemDevolucao Then
   '   MsgBox("Impossivel abrir o Movimento Tem Devoluções para Abri-lo Você deve Cancelar a(s) Devoluções")
   '   Exit Sub
   'End If
 
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "UPDATE [Outras saidas estoque] Set Fechado = 0 Where [Sequencia da saida] = " & Sequencia_da_saida 
       If Manu.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção contas receber] WHERE [Sequencia da saida] = " & Sequencia_da_saida
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
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas outras saidas] WHERE [Sequencia da saida] = " & Sequencia_da_saida)
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
   Load frmF_GNF ' Abre Form p/ Gerar NFe 
   frmF_GNF.RepositionSaida
   frmF_GNF.Executar INICIALIZACOES
   frmF_GNF.Executar INI_APELIDOS
   frmF_GNF.Show
   SendK vbKeyReturn
   SendK vbKeyReturn
   SendK vbKeyReturn 
    
     
End Sub


Private Function ValidaIPI(Sequencia_da_saida As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single) As Boolean
 Dim Produtos As New GRecordSet
 Dim Classificacao As New GRecordSet, TemIPI As Boolean
 Dim AdqTerceiros As Boolean
 
 Set Produtos = vgDb.OpenRecordSet("SELECT * From Produtos Where [Sequencia do produto] = " & Sequencia_do_produto)
 Set Classificacao = vgDb.OpenRecordSet("SELECT * From [Classificação Fiscal] Where [Sequencia da classificação] = " & Produtos![Sequencia da classificação])

 AdqTerceiros = Produtos![Material Adquirido de Terceiro]
 TemIPI = Classificacao![Percentual de ipi] > 0
 
  If AdqTerceiros Then
     ValidaIPI = False
  ElseIf Not AdqTerceiros And TemIPI Then
     ValidaIPI = True
  ElseIf Not AdqTerceiros And Not TemIPI Then 
     ValidaIPI = False     
  End If   
  
End Function


Private Sub AVISTA()
Dim Parcelas As New GRecordSet
Dim Produtos As New GRecordSet

 Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos saidas estoque] Where [Sequencia da Saida] = " & Sequencia_da_saida)
 TbAuxiliar "Parcelas outras saidas", "[Sequencia da Saida] = " & Sequencia_da_saida, Parcelas
 
 If Produtos.RecordCount = 0 And Parcelas.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [Parcelas outras saidas] Where [Sequencia da Saida] = " & Sequencia_da_saida
    GrdParcelamento.Rebind
    Exit Sub
 End If

 If (Outras_saidas_estoque![Forma de pagamento] <> "Vista" Or Outras_saidas_estoque![Valor total Do movimento] = 0) Then Exit Sub 
    With Parcelas   
      If Parcelas.RecordCount = 0 Then .AddNew
      ![Sequencia da Saida] = Outras_saidas_estoque![Sequencia da saida]
      ![Numero da Parcela] = 1
      !Dias = 0
      ![Data de Vencimento] = Outras_saidas_estoque![Data de emissão]
      ![Valor da Parcela] = Outras_saidas_estoque![Valor total Do movimento]
      ![Especie de Pagamento] = "Dinheiro"
      .Update
      .Bookmark = .lastModified
   End With
   GrdParcelamento.Rebind
   
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Sequencia_da_saida = txtCampo(0).Value
         Documento = txtCampo(1).Value
         Sequencia_do_geral = txtCampo(3).Value
         Data_de_emissao = txtCampo(2).Value
         Forma_de_pagamento = txtCampo(4).Value
         Historico = txtCampo(10).Value
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
         Valor_total_dos_produtos = txtCampo(7).Value
         Valor_total_do_ipi = txtCampo(9).Value
         Valor_total_icms_st = txtCampo(8).Value
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Valor_do_frete = txtCampo(5).Value
         Valor_do_desconto = txtCampo(6).Value
         Valor_total_do_movimento = txtCampo(11).Value
         Usuario_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Usuario da Alteração])
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Teste = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb!Teste)
      Else
         Sequencia_da_saida = IIf(IsNull(vgTb![Sequencia da saida]), 0, vgTb![Sequencia da saida])
         Documento = IIf(IsNull(vgTb!Documento), "", vgTb!Documento)
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Data_de_emissao = vgTb![Data de emissão]
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Valor_total_dos_produtos = IIf(IsNull(vgTb![Valor total dos produtos]), 0, vgTb![Valor total dos produtos])
         Valor_total_do_ipi = IIf(IsNull(vgTb![Valor total do ipi]), 0, vgTb![Valor total do ipi])
         Valor_total_icms_st = IIf(IsNull(vgTb![Valor total icms st]), 0, vgTb![Valor total icms st])
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Valor_do_frete = IIf(IsNull(vgTb![Valor do frete]), 0, vgTb![Valor do frete])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Valor_total_do_movimento = IIf(IsNull(vgTb![Valor total do movimento]), 0, vgTb![Valor total do movimento])
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Teste = IIf(IsNull(vgTb!Teste), "", vgTb!Teste)
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
      vgMsg$ = "Documento Invalido ou Vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = ((Not Vazio(Data_de_emissao)) And (IsDate(Data_de_emissao) Or Vazio(Data_de_emissao)))
         vgMsg$ = "Data de emissão não pode ser vazio!"
         If Not vgOk Then vgColumn = 3
      End If
      If vgOk Then
         If Sequencia_do_geral = 0 Then
            vgOk = (Sequencia_do_geral > 0)
            vgMsg$ = "Geral não pode ser em Branco!"
            If Not vgOk Then vgColumn = 4
         End If
      End If
      If vgOk Then
         If Sequencia_do_geral > 0 Then
            vgOk = (ValidaGeral())
            vgMsg$ = "Geral Inativo!"
            If Not vgOk Then vgColumn = 4
         End If
      End If
      If vgOk Then
         vgOk = (Valor_do_frete >= 0)
         vgMsg$ = "Valor do frete inválido!"
         If Not vgOk Then vgColumn = 6
      End If
      If vgOk Then
         vgOk = (Valor_do_desconto >= 0)
         vgMsg$ = "Valor do desconto inválido!"
         If Not vgOk Then vgColumn = 7
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
                  On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 3
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 3
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
      Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Teste = IIf(IsNull(vgTb!Teste), "", vgTb!Teste)
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
   Set Outras_saidas_estoque = vgTb
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
   Ok = (Sequencia_do_geral > 0 And Parcelas_outras_saidas.RecordCount = 0)
   txtCampo(4).Enabled = (Ok And txtCampo(4).Editable)
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = True)
   Botao(0).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0 And VgSituacao = ACAO_NAVEGANDO)
   Botao(1).Enabled = Ok
   Ok = (Sequencia_da_nota_fiscal = 0 AND vgSituacao = ACAO_NAVEGANDO AND Cancelado = False AND Fechado = True AND PermitidoMenu(MnuNFe.Tag) And Empresa![Emite NFe])
   Botao(2).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(18).Visible = (Cancelado = 0)
   Botao(0).Visible = (Cancelado = False AND Fechado = True And vgSituacao = ACAO_NAVEGANDO AND Sequencia_da_nota_fiscal = 0)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 3 Or Index = 4 Then
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
         Case 3
            PoeRelEFiltroCbo 3
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
   Caption = LoadGasString(23800)
   vgFormID = 667
   vgIdentTab = "Outras saidas estoque"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia da saida"
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
   vgUltimoTabIndex = 10
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(15305))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(15306))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(15307))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(15308))
   Set lblAjuste = Label(0)
   Set lblMovimento = Label(1)
   Set grdProdutos = Grid(0)
   Set grdParcelamento = Grid(1)
   Set lblParcelamento = Label(12)
   Set lblFechamento = Label(18)
   Set cmdAbrir = Botao(0)
   Set cmdFechar = Botao(1)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia da saida] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(23805)
   Label(1).Caption = LoadGasString(23806)
   Label(2).Caption = LoadGasString(23807)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(23808)
   Label(3).Caption = LoadGasString(23809)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(23810)
   Label(4).Caption = LoadGasString(23811)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(23812)
   Label(5).Caption = LoadGasString(23813)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(23814)
   Tabs(0).TabCaption(0) = LoadGasString(23815)
   Tabs(0).TabCaption(1) = LoadGasString(23816)
   Label(6).Caption = LoadGasString(23817)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(23818)
   Label(7).Caption = LoadGasString(23819)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(23820)
   Label(8).Caption = LoadGasString(23821)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(23822)
   Label(9).Caption = LoadGasString(23823)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(23824)
   Label(10).Caption = LoadGasString(23825)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(23826)
   Label(11).Caption = LoadGasString(23827)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(23828)
   Label(13).Caption = LoadGasString(23829)
   Label(14).Caption = LoadGasString(23830)
   Label(15).Caption = LoadGasString(23831)
   Label(16).Caption = LoadGasString(23832)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(23833)
   Label(17).Caption = LoadGasString(23834)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(23835)
   vgTooltips.AddTool Botao(0), 0, LoadGasString(23836)
   Botao(0).Caption = LoadGasString(23837)
   vgTooltips.AddTool Botao(1), 0, LoadGasString(23838)
   Botao(1).Caption = LoadGasString(23839)
   Botao(2).Caption = LoadGasString(23840)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Produtos saidas estoque", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Produtos saidas estoque", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos saidas estoque", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Sequencia do produto", TP_CARACTER, "@x", , False, , "DADOSNFE", "Produtos", "Sequencia do produto", "Descrição", "Descrição", "Descrição", "", , "1", "Produtos.[Sequencia do produto]", "", "DADOSNFE", "18", 1, "0", 4995
      .AddColumn Nothing, , "Nº Item", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 780
      .AddColumn Nothing, , "Unid", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1080
      .AddColumn Nothing, , "Estoque", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1110
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1155
      .AddColumn Nothing, , "Vr. Unitário", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "% IPI", "Aliquota do ipi", TP_NUMERICO, "99,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1035
      .AddColumn Nothing, , "Vr. IPI", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Vr. ICMS ST", "Valor icms st", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
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
      .AllowInsert = Permitido("Parcelas outras saidas", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas outras saidas", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas outras saidas", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 705
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 600
      .AddColumn Nothing, , "Dt. Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1440
      .AddColumn Nothing, , "Vr. Parcela", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1380
      .AddColumn Nothing, , "Especie de Pagto", "Especie de pagamento", TP_CARACTER, , , False, , , , , , "Dinheiro|Cartão|Cheque - Pre|Cheque - Vista|Boleto|Condicional|Garantia|Deposito", , , , "1", , , , "0", 1, "0", 1815
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Outras_saidas_estoque = vgTb
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

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = "@!"
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

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

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_LISTA) = bottxtCampo3(BOT_LISTA)
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   bottxtCampo3(BOT_LISTA).Caption = "P"
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "999999"
   txtCampo(3).PesqModoAbertura =  2 
   txtCampo(3).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(3).BoundColumn = "Sequencia do geral"
   txtCampo(3).ListFields = "Razão social"
   txtCampo(3).OrderFields = "Razão social"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Geral"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 4
   txtCampo(10).Mask = ""
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "999.999.999,99"
   txtCampo(11).Editable = False
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(4).CtPri = txtCp(4)
   Set txtCampo(4).CtFdo = labFdo4
   Set txtCampo(4).CtBot(BOT_COMBO) = bottxtCampo4(BOT_COMBO)
   Set bottxtCampo4(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(4).DataType = 0
   txtCampo(4).ListFields = "Vista|Prazo"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "99.999.999,99"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = "99.999.999,99"
   txtCampo(6).BoundColumn = ""
   txtCampo(6).ListFields = ""
   txtCampo(6).OrderFields = ""
   txtCampo(6).Relation = ""
   txtCampo(6).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "99.999.999,99"
   txtCampo(7).Editable = False
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
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
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "99.999.999,99"
   txtCampo(9).Editable = False
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

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
         txtCampo(3).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0", "[Sequencia do geral] > 0") &"
                           txtCampo(3).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                              "Geral.Bairro, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ")"
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
   Set lblAjuste = Nothing
   Set lblMovimento = Nothing
   Set grdProdutos = Nothing
   Set grdParcelamento = Nothing
   Set lblParcelamento = Nothing
   Set lblFechamento = Nothing
   Set cmdAbrir = Nothing
   Set cmdFechar = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   If Not Outras_saidas_estoque Is Nothing Then
      Set Outras_saidas_estoque = Nothing
   End If
   If Not Produtos_saidas_estoque Is Nothing Then
      Produtos_saidas_estoque.CloseRecordSet
      Set Produtos_saidas_estoque = Nothing
   End If
   If Not Parcelas_outras_saidas Is Nothing Then
      Parcelas_outras_saidas.CloseRecordSet
      Set Parcelas_outras_saidas = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmMoSaidas = Nothing                      'libera o segmento de código do form
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
         Produtos_saidas_estoque![Sequencia da saida] = Outras_saidas_estoque![Sequencia da saida]
      Case 1
         Parcelas_outras_saidas![Sequencia da saida] = Outras_saidas_estoque![Sequencia da saida]
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
         Set Produtos_saidas_estoque = vgNewRecordSet
      Case 1
         Set Parcelas_outras_saidas = vgNewRecordSet
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
   Dim Sequencia_do_item As Long, Sequencia_do_produto As Double, Quantidade As Double
   Dim Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double
   Dim Valor_icms_st As Double, Aliquota_do_ipi As Single
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
            vgRetVal = (InfoProdutos(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, "Sigla"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 4
            vgRetVal = (InfoProdutos(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, "Estoque"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 9
            vgRetVal = ((Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2)))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 8
            vgRetVal = Not (ValidaIPI(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      Comandos KeyCodeAscii, Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi 
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
      vgRetVal = "SELECT * FROM [Produtos saidas estoque]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da saida] = " & Outras_saidas_estoque![Sequencia da saida]
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
                                                  "Produtos.[Valor unitario] FROM Produtos WHERE (Produtos.[Sequencia do produto] > " + CStr(0) + ") AND (Produtos.Inativo = " + CStr(0) + ")"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_produto = 0 Then
            vgIsValid = (Sequencia_do_produto > 0)
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "Produto não Pode ser Vazio!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_produto > 0 Then
               vgIsValid = (ValidaProduto(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi))
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "Produto Inativo para Inclui - lo Abra o Cadastro do Item e Desmarque a Opção (INATIVO)!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Quantidade = 0 Then
               vgIsValid = (Quantidade > 0)
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Quantidade deve ser maior que Zero(0)!!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Quantidade > 0 Then
               vgIsValid = (ValidaEstoque(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi))
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Estoque Insuficiente!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_unitario > 0)
            If Not vgIsValid Then vgColumn =  6 
            vgErrorMessage$ = "Valor Unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(2) = SeqItem()
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AtuEstoque Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi 
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AtuEstoque Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AtuEstoque Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi 
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Outras_saidas_estoque.Edit
         Set vgRsError = Outras_saidas_estoque
         If ProcessaProdutos(Sequencia_da_saida, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi) Then
            Outras_saidas_estoque![Sequencia da saida] = (0)
         End If
         Outras_saidas_estoque.Update
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
      If Produtos_saidas_estoque.Eof = False And Produtos_saidas_estoque.BOF = False And Produtos_saidas_estoque.RecordCount > 0 Then
         Valor_total = Produtos_saidas_estoque![Valor total]
         Valor_do_ipi = Produtos_saidas_estoque![Valor do ipi]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_saidas_estoque.Eof = False And Produtos_saidas_estoque.BOF = False And Produtos_saidas_estoque.RecordCount > 0 Then
      Sequencia_do_item = Produtos_saidas_estoque![Sequencia do item]
      Sequencia_do_produto = Produtos_saidas_estoque![Sequencia do produto]
      Quantidade = Produtos_saidas_estoque!Quantidade
      Valor_unitario = Produtos_saidas_estoque![Valor unitario]
      Valor_total = Produtos_saidas_estoque![Valor total]
      Valor_do_ipi = Produtos_saidas_estoque![Valor do ipi]
      Valor_icms_st = Produtos_saidas_estoque![Valor icms st]
      Aliquota_do_ipi = Produtos_saidas_estoque![Aliquota do ipi]
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
   If Outras_saidas_estoque.EOF = False And Outras_saidas_estoque.BOF = False Then
      Outras_saidas_estoque.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Dias As Integer, Data_de_vencimento As Variant, Valor_da_parcela As Double
   Dim Especie_de_pagamento As String, Numero_da_parcela As Integer
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
            vgRetVal = Not ((Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Forma de pagamento] = "Prazo")
         Case 3
            vgRetVal = Not ((Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Forma de pagamento] = "Prazo")
         Case 4
            vgRetVal = Not (Not Vazio(Outras_saidas_estoque![Forma de pagamento]))
         Case 5
            vgRetVal = Not (Not Vazio(Outras_saidas_estoque![Forma de pagamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (((Outras_saidas_estoque!Cancelado = 0) AND Not Vazio(Outras_saidas_estoque![Forma de pagamento]) And TotalParcelas < Outras_saidas_estoque![Valor total do movimento]) AND Outras_saidas_estoque!Fechado = 0)
      On Error Resume Next
         Grid(1).AllowEdit = ((((Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Sequencia da nota fiscal] = 0) AND Outras_saidas_estoque!Fechado = 0) AND outras_saidas_estoque!Cancelado = 0)
      On Error Resume Next
         Grid(1).AllowDelete = ((((Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Sequencia da nota fiscal] = 0) AND Outras_saidas_estoque!Fechado = 0) AND outras_saidas_estoque!Cancelado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas outras saidas]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da saida] = " & Outras_saidas_estoque![Sequencia da saida]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D", Outras_saidas_estoque![Data de emissão] , Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Outras_saidas_estoque![Data de emissão])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Outras_saidas_estoque![Valor total do movimento] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If (Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Forma de pagamento] = "Prazo" Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Outras_saidas_estoque![Forma de pagamento])) AND Outras_saidas_estoque![Forma de pagamento] = "Prazo" Then
               vgIsValid = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
               If Not vgIsValid Then vgColumn =  3 
               vgErrorMessage$ = "Data de Vencimento Invalida!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Outras_saidas_estoque![Forma de pagamento]) Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Outras_saidas_estoque![Forma de pagamento] = "Prazo" , TotalParcelas() + Valor_da_Parcela <= Outras_saidas_estoque![Valor total do movimento], Valor_da_parcela = Outras_saidas_estoque![Valor total do movimento]), IIf(Outras_saidas_estoque![Forma de pagamento]  = "Vista" , Valor_da_parcela = Outras_saidas_estoque![Valor total do movimento], Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Outras_saidas_estoque![Valor total do movimento])))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total do Movimento!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Outras_saidas_estoque![Forma de pagamento]) Then
               vgIsValid = (Not Vazio(Especie_de_pagamento))
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Especie de pagamento não pode ser vazio!"
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
   Especie_de_pagamento = ColumnValue(5) & ""
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Parcelas_outras_saidas.Eof = False And Parcelas_outras_saidas.BOF = False And Parcelas_outras_saidas.RecordCount > 0 Then
      Dias = Parcelas_outras_saidas!Dias
      Data_de_vencimento = Parcelas_outras_saidas![Data de vencimento]
      Valor_da_parcela = Parcelas_outras_saidas![Valor da parcela]
      Especie_de_pagamento = Parcelas_outras_saidas![Especie de pagamento]
      Numero_da_parcela = Parcelas_outras_saidas![Numero da parcela]
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
         AbreMvto  
      Case 1
         Fechamento  
      Case 2
         AbreGerar  
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
Private Sub bottxtCampo4_Click(Index As Integer)
   txtCampo(4).SetFocus
   DoEvents
   txtCampo(4).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


