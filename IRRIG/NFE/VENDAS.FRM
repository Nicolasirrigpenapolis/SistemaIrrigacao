Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVendas
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8325
   Left                     =   3345
   LinkTopic                =   "Vendas"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7890
   ScaleWidth               =   12765
   Top                      =   1930
   Width                    =   12855
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
      Height                   =   7710
      Left                     =   15
      Width                    =   12585
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   22187
      _extenty                 =   13592
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   12420
         Width                    =   45
         Top                      =   7740
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
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   300
         Width                    =   570
         Top                      =   1620
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
         Caption                  =   "Seq:"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   435
         Left                     =   1740
         Width                    =   1305
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
         Caption                  =   "Vendas"
         Index                    =   2 
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
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   0
         TabStop                  =   0
         Height                   =   300
         Left                     =   300
         Width                    =   930
         Top                      =   1920
         DataField                =   "Sequencia da Venda"
         TabStop                  =  0
         Index                    =   0 
         MultiLine                =   -1
      End
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1335
         Left                     =   180
         Width                    =   1470
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
         Height                   =   240
         Left                     =   1320
         Width                    =   480
         Top                      =   1620
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
         Caption                  =   "Data:"
         Index                    =   3 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1320
         Top                      =   1920
         Width                    =   1425
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   2475
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
         Left                     =   1335
         Width                    =   1125
         Top                      =   1935
         DataField                =   "Data da Venda"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   2820
         Width                    =   660
         Top                      =   1620
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
         Caption                  =   "Cliente:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2820
         Top                      =   1920
         Width                    =   5670
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   8220
         Top                      =   1935
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   7965
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
            Weight                   =   700
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
         Left                     =   2835
         Width                    =   5115
         Top                      =   1935
         DataField                =   "Sequencia do geral"
         Index                    =   2 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   3
         TabHeight                =   519
         BackColor                =   &H8000000F
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   3
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   3135
         Left                     =   300
         Width                    =   12105
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
         _extentx                 =   21341
         _extenty                 =   5527
         _stockprops              =   15
         TabCaption(0)            =   "1 - Produtos"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "2 - Serviços"
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Grid(2)"
         TabCaption(2)            =   "3 - Financeiro"
         Tab(2).ControlCount=   29
         Tab(2).ControlEnabled=   0
         Tab(2).Control(0)=   "Label(5)"
         Tab(2).Control(1)=   "labFdo3"
         Tab(2).Control(2)=   "bottxtCampo3(2)"
         Tab(2).Control(3)=   "txtCp(3)"
         Tab(2).Control(4)=   "Label(6)"
         Tab(2).Control(5)=   "txtCp(4)"
         Tab(2).Control(6)=   "Label(7)"
         Tab(2).Control(7)=   "txtCp(5)"
         Tab(2).Control(8)=   "Label(8)"
         Tab(2).Control(9)=   "txtCp(6)"
         Tab(2).Control(10)=   "Label(9)"
         Tab(2).Control(11)=   "txtCp(7)"
         Tab(2).Control(12)=   "Label(10)"
         Tab(2).Control(13)=   "txtCp(8)"
         Tab(2).Control(14)=   "Label(11)"
         Tab(2).Control(15)=   "txtCp(9)"
         Tab(2).Control(16)=   "Label(12)"
         Tab(2).Control(17)=   "txtCp(10)"
         Tab(2).Control(18)=   "Label(13)"
         Tab(2).Control(19)=   "txtCp(11)"
         Tab(2).Control(20)=   "Label(14)"
         Tab(2).Control(21)=   "Label(15)"
         Tab(2).Control(22)=   "Formato(0)"
         Tab(2).Control(23)=   "Label(16)"
         Tab(2).Control(24)=   "Formato(1)"
         Tab(2).Control(25)=   "Label(17)"
         Tab(2).Control(26)=   "Formato(2)"
         Tab(2).Control(27)=   "Label(18)"
         Tab(2).Control(28)=   "Grid(1)"
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
            TabIndex                 =   8
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2430
            Left                     =   180
            Width                    =   11745
            Top                      =   480
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
            Left                     =   -74760
            Width                    =   2295
            Top                      =   600
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
            Caption                  =   "Forma de Pagamento:"
            Index                    =   5 
         End
         Begin VB.Label labFdo3
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -72300
            Top                      =   600
            Width                    =   1380
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo3
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -71190
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
               Weight                   =   700
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
            Left                     =   -72285
            Width                    =   1095
            Top                      =   615
            DataField                =   "Forma de pagamento"
            Index                    =   3 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -73770
            Width                    =   1320
            Top                      =   1080
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
            Caption                  =   "Valor do Frete:"
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
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   10
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72300
            Width                    =   1380
            Top                      =   1080
            DataField                =   "Valor do frete"
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
            Left                     =   -74175
            Width                    =   1725
            Top                      =   1560
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
            TabIndex                 =   11
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72300
            Width                    =   1380
            Top                      =   1560
            DataField                =   "Valor do desconto"
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
            Left                     =   -74205
            Width                    =   1755
            Top                      =   2040
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
            Left                     =   -72300
            Width                    =   1380
            Top                      =   2040
            DataField                =   "Valor total dos produtos"
            TabStop                  =  0
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
            Left                     =   -74190
            Width                    =   1740
            Top                      =   2520
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
            Caption                  =   "Total dos Serviços:"
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
            TabIndex                 =   13
            TabStop                  =   0
            Height                   =   300
            Left                     =   -72300
            Width                    =   1380
            Top                      =   2520
            DataField                =   "Valor Total dos Serviços"
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
            Left                     =   -70050
            Width                    =   270
            Top                      =   1560
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
            Caption                  =   "IPI:"
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
            TabIndex                 =   14
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69600
            Width                    =   1380
            Top                      =   1560
            DataField                =   "Valor total do ipi"
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
            Left                     =   -70620
            Width                    =   840
            Top                      =   1080
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
            Caption                  =   "ICMS ST:"
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
            TabIndex                 =   15
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69600
            Width                    =   1380
            Top                      =   1080
            DataField                =   "Valor total icms st"
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
            Left                     =   -70140
            Width                    =   360
            Top                      =   2040
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
            Caption                  =   "ISS:"
            Index                    =   12 
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
            TabIndex                 =   16
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69600
            Width                    =   1380
            Top                      =   2040
            DataField                =   "Valor Total do ISS"
            TabStop                  =  0
            Index                    =   10 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -70410
            Width                    =   600
            Top                      =   2520
            Alignment                =   1
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
            Caption                  =   "Total:"
            Index                    =   13 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &HC0FFC0
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
            TabIndex                 =   17
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69600
            Width                    =   1395
            Top                      =   2520
            DataField                =   "Valor Total da Venda"
            TabStop                  =  0
            Index                    =   11 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   0
            BackColor                =   &H0
            BackStyle                =   1
            BorderStyle              =   1
            AutoSize                 =   0
            Height                   =   255
            Left                     =   -67935
            Width                    =   5025
            Top                      =   315
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
            Caption                  =   "Parcelamento"
            Index                    =   14 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -67920
            Width                    =   600
            Top                      =   2775
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
            Index                    =   15 
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
            Left                     =   -67245
            Width                    =   165
            Top                      =   2820
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -67020
            Width                    =   990
            Top                      =   2775
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
            Index                    =   16 
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
            Left                     =   -65925
            Width                    =   165
            Top                      =   2820
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65625
            Width                    =   1050
            Top                      =   2775
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
            Index                    =   17 
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
            Left                     =   -64410
            Width                    =   150
            Top                      =   2820
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -64185
            Width                    =   300
            Top                      =   2775
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
            Index                    =   18 
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
            TabIndex                 =   18
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2115
            Left                     =   -67935
            Width                    =   4950
            Top                      =   600
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   1 
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
            TabIndex                 =   24
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2460
            Left                     =   -74820
            Width                    =   11655
            Top                      =   480
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   2 
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   8880
         Width                    =   945
         Top                      =   1620
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
         Index                    =   19 
      End
      Begin VB.Label labFdo12
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   8880
         Top                      =   1920
         Width                    =   3510
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo12
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   12120
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
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   270
         Left                     =   8895
         Width                    =   3210
         Top                      =   1935
         DataField                =   "Sequencia do Vendedor"
         Index                    =   12 
      End
      Begin VB.TextBox txtCp
         Appearance               =   1
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
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   1035
         Left                     =   285
         Width                    =   12135
         Top                      =   6360
         DataField                =   "Observações"
         Index                    =   13 
         MultiLine                =   -1
         ScrollBars               =   2
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   300
         Width                    =   1680
         Top                      =   5940
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
         Caption                  =   "Transportadora:"
         Index                    =   20 
      End
      Begin VB.Label labFdo14
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   2100
         Top                      =   5940
         Width                    =   10290
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   12120
         Top                      =   5955
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   11865
         Top                      =   5955
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
         TabStop                  =   0
         Height                   =   270
         Left                     =   2115
         Width                    =   9735
         Top                      =   5955
         DataField                =   "Sequencia da transportadora"
         Index                    =   14 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   7
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   300
         Left                     =   8520
         Width                    =   270
         Top                      =   1920
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
         TabIndex                 =   19
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   435
         Left                     =   9495
         Width                    =   1410
         Top                      =   120
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
         TabIndex                 =   20
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   435
         Left                     =   10950
         Width                    =   1410
         Top                      =   120
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
         TabIndex                 =   21
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   435
         Left                     =   9495
         Width                    =   1410
         Top                      =   570
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
         Index                    =   3 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   22
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   435
         Left                     =   10950
         Width                    =   1410
         Top                      =   570
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
         Index                    =   4 
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
         TabIndex                 =   23
         TabStop                  =   0
         Height                   =   255
         Left                     =   10680
         Width                    =   1800
         Top                      =   8280
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
         Caption                  =   "Tipo de Venda"
         Index                    =   0 
         DataField                =   "Tipo de Venda"
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   25
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   435
         Left                     =   7980
         Width                    =   1485
         Top                      =   570
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
         Caption                  =   "Pesquisar"
         Index                    =   5 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   26
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   435
         Left                     =   7980
         Width                    =   1485
         Top                      =   120
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
         Caption                  =   "Imprimir"
         Index                    =   6 
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
Attribute VB_name = "frmVendas"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: Vendas
'* Função....: Vendas
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:32
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
Public vgFiltroOriginal As String
Dim txtCampo(14) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(0) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Valor_total_do_ipi As Double, Valor_total_dos_produtos As Double, Valor_Total_dos_Servicos As Double
Dim Valor_Total_da_Venda As Double, Valor_total_icms_st As Double, Sequencia_do_Vendedor As Long
Dim Forma_de_pagamento As String, Valor_Total_do_ISS As Double, Valor_do_frete As Double
Dim Valor_do_desconto As Double, Cancelado As Boolean, Fechado As Boolean
Dim Sequencia_da_transportadora As Long, Data_da_Alteracao As Variant, Hora_da_Alteracao As Variant
Dim Usuario_da_Alteracao As String, Sequencia_da_nota_fiscal As Long, Tipo_de_Venda As Boolean
Dim Sequencia_da_Venda As Long, Data_da_Venda As Variant, Sequencia_do_geral As Long
Dim Observacoes As String, Teste As String
Public lblAjuste As Object, lblFechamento As Object, Aba1 As Object
Public GrdProdutos As Object, lblParcelamento As Object, grdParcelamento As Object
Public grdServicos As Object, cmdAbrir As Object
Dim Vendas As New GRecordSet, Produtos_da_Venda As New GRecordSet, Servicos_da_Venda As New GRecordSet
Dim Parcelas_da_Venda As New GRecordSet

Dim ProdutoAux As New GRecordSet, UnidadeAux As New GRecordSet
Dim ServicoAux As New GRecordSet

Private Function ValidaCliente() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Sequencia_do_geral)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaCliente = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaCliente = True: Exit Function      
   End If
  End If
  
End Function


Private Function ValidaVendedor() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Sequencia_do_Vendedor)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaVendedor = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaVendedor = True: Exit Function      
   End If
  End If
  
End Function

Private Function ValidaTransporte() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Geral Where [Sequencia do geral] = " & Sequencia_da_Transportadora)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaTransporte = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaTransporte = True: Exit Function      
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

Private Function FiltroServico() As String

  On Error Resume Next ' Se der Erro que si foda...
  FiltroServico = "[Sequencia do Serviço] > 0"
  If vgSituacao <> ACAO_NAVEGANDO Then
     FiltroServico = FiltroServico & " And [Sequencia do Serviço] > 0 AND Inativo = 0"
  End If 
    
End Function


Private Function SeqItem() As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 
    If Tb.RecordCount > 0 Then
      SeqItem = Tb!Seq + 1
    Else
      SeqItem = 1
    End If
      
End Function

Private Function SeqItemServico() As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Serviços da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 
    If Tb.RecordCount > 0 Then
      SeqItemServico = Tb!Seq + 1
    Else
      SeqItemServico = 1
    End If
      
End Function


Private Function InfoProdutos(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double, Oq As String) As Variant
   
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
      Case "Ult Venda"
         InfoProdutos = ProdutoAux![Ultima Venda]   
   End Select

End Function

Private Function InfoServicos(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_Servico As Long, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Serviços", "[Sequencia do Serviço] = " & Sequencia_do_Servico, ServicoAux
      
   If ServicoAux.RecordCount = 0 Then Exit Function
   
   Select Case Oq      
      Case "Valor"
      InfoServicos = ServicoAux![Valor do Serviço] 
   End Select

End Function


Private Function ProcessaProdutos(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double) As Boolean  
   On Error GoTo DeuErro
       
   vgDb.BeginTrans   
   vgDb.Execute "Update [Produtos da Venda] Set [Valor Total] = " & Substitui(CCur(Round(Quantidade * Valor_Unitario, 2)), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item
   
    If Not Parametros_da_nfe![Simples nacional] Then   
       vgDb.Execute "Update [Produtos da Venda] Set [Valor icms st] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Vendas![Sequencia Do geral], 12, Quantidade * Valor_Unitario, Vendas![Valor Do frete], Vendas![Valor Do desconto]), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Valor do ICMS ST
       vgDb.Execute "Update [Produtos da Venda] Set [Valor do ipi] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Vendas![Sequencia Do geral],4, Quantidade * Valor_Unitario,0,0), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Valor do IPI
       vgDb.Execute "Update [Produtos da Venda] Set [Aliquota do ipi] = " & Substitui(CalculaImposto(Sequencia_do_Produto, Vendas![Sequencia Do geral], 3, Quantidade * Valor_Unitario,0,0), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Alíquota do IPI   
    Else
      vgDb.Execute "Update [Produtos da Venda] Set [Valor icms st] = " & Substitui(CalculaSimplesNacional(Sequencia_do_Produto, Vendas![Sequencia Do geral], 12, Quantidade * Valor_Unitario, Vendas![Valor Do frete], Vendas![Valor Do desconto]), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Valor do ICMS ST
      vgDb.Execute "Update [Produtos da Venda] Set [Valor do ipi] = " & Substitui(CalculaSimplesNacional(Sequencia_do_Produto, Vendas![Sequencia Do geral],4, Quantidade * Valor_Unitario,0,0), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Valor do IPI
      vgDb.Execute "Update [Produtos da Venda] Set [Aliquota do ipi] = " & Substitui(CalculaSimplesNacional(Sequencia_do_Produto, Vendas![Sequencia Do geral], 3, Quantidade * Valor_Unitario,0,0), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item 'Alíquota do IPI
    End If 
   vgDb.CommitTrans
   Produtos_da_Venda.Requery   
         
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function

Private Function ProcessaServicos(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_Servico As Long, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double)
         
   On Error GoTo DeuErro
   
   vgDb.BeginTrans      
   vgDb.Execute "Update [Serviços da Venda] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " AND [Sequencia do Item] = " & Sequencia_do_Item
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Sub Comandos(KeyAscii As Integer, Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double)
   On Error GoTo DeuErro
   
   If KeyAscii = vbKeyF12 Then
      With GrdProdutos
         Select Case .ColumnField(.Col)
            Case "Sequencia do produto"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmProdutos.Show        
         End Select
      End With
      ElseIf KeyAscii = vbKeyF2 Then
      AtualizaImposto
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub

Private Sub ComandosServicos(KeyAscii As Integer, Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_Servico As Long, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double)
   On Error GoTo DeuErro
   
   If KeyAscii = vbKeyF12 Then
      With grdServicos
         Select Case .ColumnField(.Col)
            Case "Sequencia do Serviço"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               'PosicionaRegistro frmServicos, "Sequencia do Serviço", seqRegistro
               frmServicos.Show        
         End Select
      End With
   End If
   
DeuErro:
   If Err.Number = 438 Then Err.Number = 0
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
End Sub

Private Function ValidaProduto(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double) As Boolean
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

Private Function ValidaServico(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_Servico As Long, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double) As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Serviços Where [Sequencia do Serviço] = " & Sequencia_do_Servico)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaServico = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaServico = True: Exit Function      
   End If
  End If
  
End Function

Private Function AtuEstoque(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double)
Dim Tb As New GRecordSet

  If vgSITUACAO = -ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
   If Tb.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Produtos da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " And [Sequencia do item] = " & Sequencia_do_item
   End If
  End If 
  
  AtualizaEstoque(Sequencia_do_produto)
  UltimaVenda(Sequencia_do_produto)
  AjustaValores
   
End Function


Private Function ValidaEstoque(Sequencia_da_Venda As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Aliquota_do_ipi As Single, Valor_icms_st As Double) As Boolean ' Não deixa dar saida em estoque < 0
 
   On Error Resume Next
   
   If Empresa![Não Controlar Estoque] Then
      ValidaEstoque = True
      Exit Function
   End If
     
   If vgSituacao = -ACAO_INCLUINDO Then
      ValidaEstoque = (ProdutoAux!Estoque - Quantidade) >= 0
   Else
      ValidaEstoque = (ProdutoAux!Estoque + Produtos_da_Venda!Quantidade - Quantidade) >= 0
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

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub

Public Sub AjustaValores()
   Dim IPIProdutos As New GRecordSet, ValorProdutos As New GRecordSet 
   Dim ValorICMSST As New GRecordSet, ValorVenda As Currency, ValorServicos As New GRecordSet
   Dim Produtos As New GRecordSet, Servicos As New GRecordSet
   Dim ValorISS As Double
 
   On Error GoTo DeuErro
   
      If Me.Caption = "Orçamento" Then
         Aba1.TabVisible(2) = True    
      End If
   
   Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
   Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
   
   'AtualizaLancamentos 'Atualizando os Lançamentos dos Produtos: Valor de Custo, Ultimo Fornecedor, Ultima Compra e etc...
          
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT ROUND(SUM([Valor do ipi]), 2) Total FROM [Produtos da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda) 'IPI dos Produtos
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Produtos da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda) 'Produtos Novos
   Set ValorServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Serviços da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda) 'Serviços
   ValorISS = Round((ValorServicos!Total * Empresa![Aliquota Do Iss] / 100),2) 'ISS
   Set ValorICMSST = vgDb.OpenRecordSet("SELECT SUM([Valor icms st]) Total FROM [Produtos da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda) 'Substituição Produtos                                                                                             
   ValorVenda = IPIProdutos!Total + ValorServicos!Total + ValorProdutos!Total + ValorICMSST!Total + Valor_do_Frete - Valor_do_desconto
   ValorVenda = Format(ValorVenda, "##,###,##0.00")
   
   If Produtos.RecordCount = 0 And Servicos.RecordCount = 0 Then
      ValorVenda = 0 
      If Valor_do_frete > 0 Or Valor_do_desconto > 0 Then 
         vgDb.BeginTrans   
         vgDb.Execute "Update Vendas Set [Valor do frete] = 0 WHERE [Sequencia da Venda] = " & Sequencia_da_Venda
         vgDb.Execute "Update Vendas Set [Valor do desconto] = 0 WHERE [Sequencia da Venda] = " & Sequencia_da_Venda
         vgDb.CommitTrans
      End If  
   End If

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update Vendas Set [Valor total do ipi] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'IPI Produtos
   vgDb.Execute "Update Vendas Set [Valor total dos produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'Produtos 
   vgDb.Execute "Update Vendas Set [Valor total dos Serviços] = " & Substitui(ValorServicos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'Serviços
   vgDb.Execute "Update Vendas Set [Valor Total do ISS] = " & Substitui(CStr(ValorISS), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'ISS
   vgDb.Execute "Update Vendas Set [Valor total icms st] = " & Substitui(ValorICMSST!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'Substituição
   vgDb.Execute "Update Vendas Set [Valor total da Venda] = " & Substitui(CStr(ValorVenda), ",", ".", SO_UM) & " WHERE [Sequencia da Venda] = " & Sequencia_da_Venda 'Valor da Entrada
   vgDb.CommitTrans
     
   Alteracao
   
   If Produtos.RecordCount > 0 And vgSituacao = ACAO_EDITANDO Then
      Do While Not Produtos.EOF
         UltimaVenda(Produtos![Sequencia Do produto])
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


Public Sub AtualizaImposto()       
'Screen.MousePointer = vbHourglass

If vgSituacao = ACAO_INCLUINDO Then Exit Sub  
  If Produtos_da_Venda.RecordCount > 0 Then      
   Produtos_da_Venda.MoveFirst
    Do While Not Produtos_da_Venda.EOF
       With Produtos_da_Venda
       .Edit
          If Not Parametros_da_nfe![Simples nacional] Then
             ![Aliquota Do Ipi] = CalculaImposto(![Sequencia Do produto], Sequencia_do_geral, 3, !Quantidade * ![Valor Unitario],0,0)
             ![Valor Do ipi] = CalculaImposto(![Sequencia Do produto], Sequencia_do_geral, 4, !Quantidade * ![Valor Unitario],0,0)
             ![Valor icms st] = CalculaImposto(![Sequencia Do produto], Sequencia_do_geral, 12, !Quantidade * ![Valor Unitario], Valor_do_frete, Valor_do_desconto)
           Else 'Simples Nacional
             ![Aliquota Do Ipi] = CalculaSimplesNacional(![Sequencia Do produto], Sequencia_do_geral, 3, !Quantidade * ![Valor Unitario], 0,0)
             ![Valor Do ipi] = CalculaSimplesNacional(![Sequencia Do produto], Sequencia_do_geral, 4, !Quantidade * ![Valor Unitario], 0,0)
             ![Valor icms st] = CalculaSimplesNacional(![Sequencia Do produto], Sequencia_do_geral, 12, !Quantidade * ![Valor Unitario], Valor_do_frete, Valor_do_desconto) 
           End If
            .Update
            .BookMark = .LastModified
            .MoveNext
         End With
      Loop
      grdProdutos.ReBind 'Atualiza do Grid
      Reposition True    'Atualiza o Formulário   
    End If     
'SaiDaSub:
   'Screen.MousePointer = vbDefault

End Sub


Public Sub CarregaTotalizador()
   On Error Resume Next

   With grdProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowSumCol(.Columns("Valor icms st").Index) = True
      .ShowSumCol(.Columns("Valor do IPI").Index) = True
      .ShowSumCol(.Columns("Aliquota do ipi").Index) = False
      .ColumnWidth(CInt(.Columns("Sequencia do item").Index)) = 0
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
    With grdServicos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
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
   'BloqueiaAba
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


Private Sub RepositionVendas()
   On Error Resume Next
   
   If vgSituacao = ACAO_INCLUINDO Then
      Set Produtos_da_Venda = Nothing
      Set Parcelas_da_Venda = Nothing
      Set Servicos_da_Venda = Nothing
   End If
   
   CarregaTotalizador
   'Alteração
   lblAlteracao.Caption = Vendas![Usuario da Alteração] & " " & Vendas![Data da Alteração] & " " & Vendas![Hora da Alteração]
  
End Sub


Private Sub DeletaItens()' Qdo excluir Registro Mae
Dim Tb As New GRecordSet, Parcelas As New GRecordSet, Servicos As New GRecordSet

  If vgSITUACAO = ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
  Set Servicos = vgDb.OpenRecordSet("SELECT * From [Serviços da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
  Set Parcelas = vgDb.OpenRecordSet("SELECT * From [Parcelas da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
  
   If Parcelas.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Parcelas da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda
   '   DeletaConta
   End If
       If Servicos.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Serviços da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda
       End If
   If Tb.RecordCount > 0 Then
      Do While Not Tb.EOF
         vgDb.Execute "DELETE FROM [Produtos da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & " And [Sequencia do produto] = " & Tb![Sequencia Do produto]
         AtualizaEstoque(Tb![Sequencia Do produto])
         UltimaVenda(Tb![Sequencia Do produto])
         Tb.MoveNext
      Loop
      Tb.MoveFirst
   End If
  End If   
End Sub


Private Sub MostraFormulas()
 If Me.Caption <> "Orçamento" Then    
   If Fechado Then
      lblFechamento.Caption = "Venda Fechada"
      lblFechamento.ForeColor = &H0
   Else
      lblFechamento.Caption = "Venda Aberta"
      lblFechamento.ForeColor = &H8FF
   End If
   If Err Then Err = 0: lblFechamento.Caption = ""   
   If Cancelado Then
      lblFechamento.ForeColor = &H8FF
      lblFechamento.Caption = "Venda Cancelada"
   End If 
 Else ' Orçamento
   lblFechamento.Caption = "Orçamento"
   lblFechamento.ForeColor = &H0
 End If
 If Err Then Err = 0: lblFechamento.Caption = ""  
 If Err Then Err.Clear 
 lblParcelamento.Caption = "Parcelamento"
 If Err Then Err = 0: lblParcelamento.Caption = ""   
    
   'Label do Parcelamento
    lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
    If TotalParcelas = 0 Then
       lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
    ElseIf (TotalParcelas < Valor_Total_da_Venda) Or (TotalParcelas > Valor_Total_da_Venda) Then
       lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
    End If
             
End Sub


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   Total = Tb!Total
   TotalParcelas = Total

End Function

Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC FROM [Parcelas da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Sub AVISTA()
Dim Parcelas As New GRecordSet
Dim Produtos As New GRecordSet

 Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 TbAuxiliar "Parcelas da Venda", "[Sequencia da Venda] = " & Sequencia_da_Venda, Parcelas
 
 If Produtos.RecordCount = 0 And Parcelas.RecordCount > 0 Then
    vgDb.Execute "DELETE FROM [Parcelas da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda
    GrdParcelamento.Rebind
    Exit Sub
 End If

 If (Vendas![Forma de pagamento] <> "Vista" Or Vendas![Valor Total da Venda] = 0) Then Exit Sub 
    With Parcelas   
      If Parcelas.RecordCount = 0 Then .AddNew
         ![Sequencia da Venda] = Vendas![Sequencia da Venda]
         ![Numero da Parcela] = 1
         !Dias = 0
         ![Data de Vencimento] = Vendas![Data da Venda]
         ![Valor da Parcela] = Vendas![Valor Total da Venda]
         ![Especie de Pagamento] = "Dinheiro"
      .Update
      .Bookmark = .lastModified
   End With
   GrdParcelamento.Rebind
   
End Sub


Private Sub LancaParcelas()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
   Dim SeqBaixa As Long
   
   On Error GoTo DeuErro
      
   If Vazio(Forma_de_pagamento) Then Exit Sub
   
   If TotalParcelas() <> Valor_Total_da_Venda Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas receber")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas receber", "[Sequencia da Venda] = " & Sequencia_da_Venda & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew
         !Parcela = Parcelas![Numero da parcela]
         ![Sequencia Do geral] = Sequencia_do_Geral
         !Documento = "Venda Nº " & Sequencia_da_Venda
         ![Data de entrada] = Data_da_Venda
         !Historico = Observacoes
         ![Forma de Pagamento] = Forma_de_pagamento
         ![Especie de pagamento] = Parcelas![Especie de pagamento]
          If ![Forma de pagamento] = "Vista" And Empresa![Baixa Receber Auto] Then
             ![Valor pago] = Parcelas![Valor da parcela]
             ![Valor restante] = 0
             ![Data da baixa] = Data_da_Venda
          Else
             ![Valor pago] = 0
             ![Valor restante] = Parcelas![Valor da parcela]
             ![Data da baixa] = Null
          End If
             ![Data de vencimento] = Parcelas![Data de vencimento]
             ![Valor total] = Valor_Total_da_Venda
             ![Valor da parcela] = Parcelas![Valor da parcela]
             ![Sequencia da Venda] = Sequencia_da_Venda
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
           'frmBxConRec.Show
           'frmBxConRec.SetFocus
           'frmBxConRec.vgFiltroInicial = "[Sequencia da Baixa] = " & SeqBaixa
           'InicializaFiltro frmBxConRec
           'frmBxConRec.Reposition
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
 Dim Manutencao As New GRecordSet, Servicos As New GRecordSet
 
 On Error Resume Next
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas Receber] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda)
  
 If Produtos.RecordCount = 0 And Servicos.RecordCount = 0 Then
    MsgBox("Impossivel Fechar a Venda!! Inclua o(s)Itens / Serviços")
    Exit Sub
 End If
 
  If Parcelas.RecordCount > 0 Then
   If TotalParcelas() <> Valor_Total_da_Venda Then 
      MsgBox("Impossivel Fechar a Venda!! A Soma das Parcelas não Correspondem ao Valor Total da Venda!")
   Exit Sub
   End If
  End If
  
  If Parcelas.RecordCount > 0 And Manutencao.RecordCount = 0 Then
     LancaParcelas 
  'Else
  '   DeletaConta
  '   LancaParcelas     
  End If
  
  If Parcelas.RecordCount = 0 Then 
     MsgBox("Impossivel Fechar a Venda! Inclua a(s) Parcelas do Contas a Receber")
     Exit Sub
  Else   
     vgDb.Execute "UPDATE Vendas Set Fechado = 1 WHERE [Sequencia da Venda] = " & Sequencia_da_Venda
     vgDb.Execute "UPDATE Vendas Set [Tipo de Venda] = 0 WHERE [Sequencia da Venda] = " & Sequencia_da_Venda
  End If   
  Reposition True
  
  If Me.Caption = "Orçamento" Then
     Unload Me
     frmVendas.vgFiltroInicial = "[Sequencia da Venda] = " & Sequencia_da_Venda 
     InicializaFiltro frmVendas
     frmVendas.Show
     frmVendas.SetFocus 
     frmVendas.AtuEstoqueVemOrc
  End If
              
End Sub


Private Sub AbreGerar()
   Dim Geral As New GRecordSet
   Dim DtaHoje As Date, DtaParcela As Date, Parcelas As New GRecordSet
   
   On Error Resume Next
   
   DtaHoje = Date 'Data de Hoje
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas da Venda] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda)
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
   If TotalParcelas > 0 And TotalParcelas <> Valor_Total_da_Venda Then
      MsgBox "Parcelamento Incompleto. Impossível Gerar a NFe.", vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
      
   Unload Me
   Load frmVendaNFe ' Abre Form p/ Gerar NFe 
   'frmVendaNFe.RepositionVendas
   frmVendaNFe.Executar INICIALIZACOES
   frmVendaNFe.Executar INI_APELIDOS
   frmVendaNFe.Show
   SendK vbKeyReturn
   SendK vbKeyReturn
   SendK vbKeyReturn 
   frmVendaNFe.txtVendaTela.Value = Sequencia_da_Venda 
   frmVendaNFe.RepositionVendas      
End Sub 


Private Sub AbreVenda()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean, TemDevolucao As Boolean
 Dim Manu As New GRecordSet

 On Error Resume Next
 
 Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas Receber] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas Receber] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda & ")")
 Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas Receber] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
  
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If
   
   If TemBaixa Then
      MsgBox("Impossivel abrir a Venda existe Baixa(s) no Contas a Receber para Abri-la Exclua a(s) Baixa(s)")
      Exit Sub
   End If
   
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "UPDATE Vendas Set Fechado = 0 Where [Sequencia da Venda] = " & Sequencia_da_Venda 
       If Manu.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção contas Receber] WHERE [Sequencia da Venda] = " & Sequencia_da_Venda
       End If          
       Reposition True 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  InputYgorsBom = ""
  
End Sub  

Public Sub BloqueiaAba()
If Sequencia_da_Venda = 0 Then        
   Aba1.TabVisible(2) = False 
End If      
End Sub

Private Sub Cancelamento() 
 Dim Tb As New GRecordSet
 Dim Produtos As New GRecordSet
    
   On Error GoTo DeuErro
      
   If MsgBox("Deseja Realmente Cancelar essa Venda - As Baixas do Contas a Receber e a Conta Tbm serão Canceladas?" & vbCrLf & "", vbExclamation + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   
   Set Tb = vgDb.OpenRecordSet("SELECT Mvto.[Sequencia da Venda], " & _
                               "Manu.[Sequencia da manutenção] Conta " & _
                               "FROM Vendas Mvto " & _
                               "LEFT JOIN [Manutenção contas receber] Manu ON Mvto.[Sequencia da Venda] = Manu.[Sequencia da Venda] " & _
                               "WHERE Mvto.[Sequencia da Venda] = " & Sequencia_da_Venda)
                               
   Set Produtos = vgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda) 
   
   vgDb.BeginTrans 'Inicia a transação
   
  InputSenha
 If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "Update Vendas Set [Cancelado] = 1 Where [Sequencia da Venda] = " & Sequencia_da_Venda
       If Tb!Conta > 0 AND Not Vazio(Tb!Conta) Then
          Do While Not Tb.EOF 'Excluindo o Contas se houver
             vgDb.Execute "DELETE FROM [Baixa Contas receber] WHERE [Sequencia da Manutenção] = " & Tb!Conta 'Baixa
             'vgDb.Execute "DELETE FROM [Movimentação da Conta Corrente] WHERE [Seqüência do Lançamento] = " & Tb!Conta 'Conta Corrente
             Tb.MoveNext
          Loop
          Tb.MoveFirst 'Vamos Voltar no Começo
          vgDb.Execute "DELETE FROM [Manutenção Contas Receber] WHERE [Sequencia da Venda] = " & Tb![Sequencia da Venda] 
       End If  
    vgDb.CommitTrans 'Vomitando opsss Comitando rsrs
    MsgBox "Venda Cancelada com Sucesso!!!", vbInformation + vbOKOnly, vaTitulo 'Suuuuuceeeeesssssoooo
    If Produtos.RecordCount > 0 Then
      Do While Not Produtos.EOF
         AtualizaEstoque(Produtos![Sequencia Do produto])
      Produtos.MoveNext
      Loop
      Produtos.MoveFirst      
    End If       
       Reposition True 
  Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  InputYgorsBom = ""
         
DeuErro:
   If Err.Number <> 0 Then
      vgDb.RollBackTrans
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
     
End Sub

Private Sub VerificaFormulario
   On error Resume Next
   
   If Me.Caption = "Orçamento" Then
      OrcamentoAberto = False
   End If
   lblAlteracao.caption = ""

End Sub


Public Sub AtuEstoqueVemOrc() 'Chama ele qdo orc vira venda
Dim Tb As New GRecordSet

  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos da Venda] Where [Sequencia da Venda] = " & Sequencia_da_Venda)
   If Tb.RecordCount > 0 Then
      Do While Not Tb.EOF 
      AtualizaEstoque(Tb![Sequencia Do Produto])
      UltimaVenda(Tb![Sequencia Do Produto])
      Tb.MoveNext 
      Loop 
   End If
    
End Sub

Private Sub AbreRelatorio()
 If Empresa![Impressão Do Relatorio] = 0 Then
    frmRelVenda.Show
    frmRelVenda.txtVenda.Value = Sequencia_da_Venda
 Else
    frmRelImpV2.Show
    frmRelImpV2.txtVenda.Value = Sequencia_da_Venda
 End If   
End Sub


Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
    Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia da Venda]) Seq From Vendas")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function

Public Sub LigaDesligaBotoes
   Botao( 0 ).Enabled =    Botao( 0 ).Enabled And Permitido("Geral", ACAO_NAVEGANDO)
   Botao( 5 ).Enabled = mdiNFE.Menu_Procura.Enabled
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Valor_total_do_ipi = txtCampo(8).Value
         Valor_total_dos_produtos = txtCampo(6).Value
         Valor_Total_dos_Servicos = txtCampo(7).Value
         Valor_Total_da_Venda = txtCampo(11).Value
         Valor_total_icms_st = txtCampo(9).Value
         Sequencia_do_Vendedor = txtCampo(12).Value
         Forma_de_pagamento = txtCampo(3).Value
         Valor_Total_do_ISS = txtCampo(10).Value
         Valor_do_frete = txtCampo(4).Value
         Valor_do_desconto = txtCampo(5).Value
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Sequencia_da_transportadora = txtCampo(14).Value
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Usuario_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Usuario da Alteração])
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Tipo_de_Venda = chkCampo(0).Value
         Sequencia_da_Venda = txtCampo(0).Value
         Data_da_Venda = txtCampo(1).Value
         Sequencia_do_geral = txtCampo(2).Value
         Observacoes = txtCampo(13).Value
         Teste = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb!Teste)
      Else
         Valor_total_do_ipi = IIf(IsNull(vgTb![Valor total do ipi]), 0, vgTb![Valor total do ipi])
         Valor_total_dos_produtos = IIf(IsNull(vgTb![Valor total dos produtos]), 0, vgTb![Valor total dos produtos])
         Valor_Total_dos_Servicos = IIf(IsNull(vgTb![Valor Total dos Serviços]), 0, vgTb![Valor Total dos Serviços])
         Valor_Total_da_Venda = IIf(IsNull(vgTb![Valor Total da Venda]), 0, vgTb![Valor Total da Venda])
         Valor_total_icms_st = IIf(IsNull(vgTb![Valor total icms st]), 0, vgTb![Valor total icms st])
         Sequencia_do_Vendedor = IIf(IsNull(vgTb![Sequencia do Vendedor]), 0, vgTb![Sequencia do Vendedor])
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Valor_Total_do_ISS = IIf(IsNull(vgTb![Valor Total do ISS]), 0, vgTb![Valor Total do ISS])
         Valor_do_frete = IIf(IsNull(vgTb![Valor do frete]), 0, vgTb![Valor do frete])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Sequencia_da_transportadora = IIf(IsNull(vgTb![Sequencia da transportadora]), 0, vgTb![Sequencia da transportadora])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Tipo_de_Venda = IIf(IsNull(vgTb![Tipo de Venda]), 0, vgTb![Tipo de Venda])
         Sequencia_da_Venda = IIf(IsNull(vgTb![Sequencia da Venda]), 0, vgTb![Sequencia da Venda])
         Data_da_Venda = vgTb![Data da Venda]
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Observacoes = IIf(IsNull(vgTb!Observações), "", vgTb!Observações)
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
      vgOk = ((Not Vazio(Data_da_Venda)) And (IsDate(Data_da_Venda) Or Vazio(Data_da_Venda)))
      vgMsg$ = "Data da Venda não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         If Sequencia_do_geral = 0 Then
            vgOk = (Sequencia_do_geral > 0)
            vgMsg$ = "Cliente é Obrigatorio caso não tenha Cadastro use Consumidor!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         If Sequencia_do_geral > 0 Then
            vgOk = (ValidaCliente())
            vgMsg$ = "Impossivel - Cliente Inativo!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         If Sequencia_do_Vendedor = 0 Then
            vgOk = (Sequencia_do_Vendedor > 0)
            vgMsg$ = "Vendedor é Obrigatorio! "
            If Not vgOk Then vgColumn = 13
         End If
      End If
      If vgOk Then
         If Sequencia_do_Vendedor > 0 Then
            vgOk = (ValidaVendedor())
            vgMsg$ = "Impossivel - Vendedor Inativo!"
            If Not vgOk Then vgColumn = 13
         End If
      End If
      If vgOk Then
         vgOk = (ValidaTransporte())
         vgMsg$ = "Impossivel - Transportadora Inativa!"
         If Not vgOk Then vgColumn = 15
      End If
      If vgOk Then
         vgOk = (Valor_do_frete >= 0)
         vgMsg$ = "Valor do Frete inválido!"
         If Not vgOk Then vgColumn = 5
      End If
      If vgOk Then
         vgOk = (Valor_do_desconto >= 0)
         vgMsg$ = "Valor do Desconto inválido!"
         If Not vgOk Then vgColumn = 6
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
                  chkCampo(0).Value = IIf(Me.Caption = "Vendas", 0, 1)
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 14
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 12
         PoeRelEFiltroCbo 14
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Tipo_de_Venda
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
      Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      InicializaApelidos COM_TEXTBOX
      vgTb![Tipo de Venda] = Tipo_de_Venda
      Teste = IIf(IsNull(vgTb!Teste), "", vgTb!Teste)
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_INCLUINDO Then
         AjustaValores
      ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
         AjustaValores
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
   Set Vendas = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   On Error Resume Next
   For i = 0 To 2
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
         Case 2
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(2).CloseRecordset
            Else
               x$ = ExecutaGrid(2, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(2).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(2).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(2).CloseRecordSet
                  End If
                  Grid(2).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(2, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionVendas  
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
   Ok = (Parcelas_da_Venda.RecordCount = 0)
   txtCampo(3).Enabled = (Ok And txtCampo(3).Editable)
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0 And VgSituacao = ACAO_NAVEGANDO And Empresa![Nome fantasia] <> "Reticom Virabrequim")
   Botao(1).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = True And VgSituacao = ACAO_NAVEGANDO And Me.Caption = "Vendas" And Empresa![Nome fantasia] <> "Reticom Virabrequim")
   Botao(2).Enabled = Ok
   Ok = (Sequencia_da_nota_fiscal = 0 AND vgSituacao = ACAO_NAVEGANDO AND Cancelado = False AND Fechado = True AND PermitidoMenu(MnuNFe.Tag) And Me.Caption = "Vendas" And Empresa![Emite NFe])
   Botao(3).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = True And VgSituacao = ACAO_NAVEGANDO And Me.Caption = "Vendas")
   Botao(4).Enabled = Ok
   Ok = (False)
   chkCampo(0).Enabled = (Ok And chkCampo(0).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Or Index = 3 Or Index = 12 Or Index = 14 Then
      ExecutaPreValidacao
      MostraFormulas
   End If
   Select Case Index
      Case 2
         AtualizaImposto  
   End Select
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
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 2
            PoeRelEFiltroCbo 2
         Case 12
            PoeRelEFiltroCbo 12
         Case 14
            PoeRelEFiltroCbo 14
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
      MostraFormulas                              'mostra formulas na janela
      ExecutaPreValidacao                         'habilita/desabilita campos
   End If
   Select Case Index
      Case 4
         AtualizaImposto  
      Case 5
         AtualizaImposto  
   End Select
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      chkCampo(0).Value = Tipo_de_Venda
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      MostraFormulas                              'mostra formulas na janela
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
   Caption = LoadGasString(18160)
   vgFormID = 703
   vgIdentTab = "Vendas"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia da Venda"
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
   vgTemCondicoesEsp = True
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 24
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(9665))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(9666))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(9667))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(9668))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(9669))
   Set Botao(4).Picture = LoadPicture(LoadGasPicture(9670))
   Set Botao(5).Picture = LoadPicture(LoadGasPicture(9671))
   Set Botao(6).Picture = LoadPicture(LoadGasPicture(9672))
   Set lblAjuste = Label(0)
   Set lblFechamento = Label(2)
   Set Aba1 = Tabs(0)
   Set GrdProdutos = Grid(0)
   Set lblParcelamento = Label(14)
   Set grdParcelamento = Grid(1)
   Set grdServicos = Grid(2)
   Set cmdAbrir = Botao(2)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "([Sequencia da Venda] > " & 0 & ") AND [Tipo de Venda] = " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(18165)
   Label(1).Caption = LoadGasString(18166)
   Label(2).Caption = LoadGasString(18167)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(18168)
   Label(3).Caption = LoadGasString(18169)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(18170)
   Label(4).Caption = LoadGasString(18171)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(18172)
   Tabs(0).TabCaption(0) = LoadGasString(18173)
   Tabs(0).TabCaption(1) = LoadGasString(18174)
   Tabs(0).TabCaption(2) = LoadGasString(18175)
   Label(5).Caption = LoadGasString(18176)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(18177)
   Label(6).Caption = LoadGasString(18178)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(18179)
   Label(7).Caption = LoadGasString(18180)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(18181)
   Label(8).Caption = LoadGasString(18182)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(18183)
   Label(9).Caption = LoadGasString(18184)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(18185)
   Label(10).Caption = LoadGasString(18186)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(18187)
   Label(11).Caption = LoadGasString(18188)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(18189)
   Label(12).Caption = LoadGasString(18190)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(18191)
   Label(13).Caption = LoadGasString(18192)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(18193)
   Label(15).Caption = LoadGasString(18194)
   Label(16).Caption = LoadGasString(18195)
   Label(17).Caption = LoadGasString(18196)
   Label(18).Caption = LoadGasString(18197)
   Label(19).Caption = LoadGasString(18198)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(18199)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(18200)
   Label(20).Caption = LoadGasString(18201)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(18202)
   Botao(1).Caption = LoadGasString(18203)
   vgTooltips.AddTool Botao(2), 0, LoadGasString(18204)
   Botao(2).Caption = LoadGasString(18205)
   Botao(3).Caption = LoadGasString(18206)
   Botao(4).Caption = LoadGasString(18207)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(18208)
   chkCampo(0).Caption = LoadGasString(18209)
   Botao(5).Caption = LoadGasString(18210)
   Botao(6).Caption = LoadGasString(18211)
   With Grid(0)
      .RowHeight = 330
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Produtos da Venda", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Produtos da Venda", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos da Venda", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Sequencia do produto", TP_NUMERICO, "", , False, , "DADOSNFE", "Produtos", "Sequencia do produto", "Descrição; Sequencia do produto", "Descrição; Sequencia do produto", "Descrição; Sequencia do produto", "", , "1", "Produtos.[Sequencia do produto]", "", "DADOSNFE", "18", 2, "0", 5835
      .AddColumn Nothing, , "Nº Item", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1080
      .AddColumn Nothing, , "Unidade", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1515
      .AddColumn Nothing, , "Estoque", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1440
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1425
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1545
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1515
      .AddColumn Nothing, , "Valor do IPI", "Valor do ipi", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1425
      .AddColumn Nothing, , "% IPI", "Aliquota do ipi", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Valor ICMS ST", "Valor icms st", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1920
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
      .AllowInsert = Permitido("Parcelas da Venda", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas da Venda", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas da Venda", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 825
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 825
      .AddColumn Nothing, , "Dt. Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1530
      .AddColumn Nothing, , "Valor da Parcela", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1920
      .AddColumn Nothing, , "Especie de Pagamento", "Especie de pagamento", TP_CARACTER, , , False, , , , , , "Dinheiro|Cartão|Cheque - Pre|Cheque - Vista|Boleto|Condicional|Garantia|Deposito", , , , "1", , , , "0", 1, "0", 2325
   End With
   With Grid(2)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Serviços da Venda", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Serviços da Venda", ACAO_EDITANDO)
      .AllowDelete = Permitido("Serviços da Venda", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Serviço", "Sequencia do Serviço", TP_CARACTER, "@x", , False, , "DADOSNFE", "Serviços", "Sequencia do Serviço", "Descrição do Serviço", "Descrição do Serviço", "Descrição do Serviço", "", , "1", "Serviços.[Sequencia do Serviço]", "", "DADOSNFE", "18", 2, "0", 7185
      .AddColumn Nothing, , "Nº Item:", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1320
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1425
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Vendas = vgTb
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
   txtCampo(0).Mask = "9999999"
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
   Set txtCampo(2).CtFdo = labFdo2
   Set txtCampo(2).CtBot(BOT_LISTA) = bottxtCampo2(BOT_LISTA)
   Set txtCampo(2).CtBot(BOT_COMBO) = bottxtCampo2(BOT_COMBO)
   bottxtCampo2(BOT_LISTA).Caption = "P"
   Set bottxtCampo2(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(2).DataType = 1
   txtCampo(2).Mask = "999999"
   txtCampo(2).PesqModoAbertura =  2 
   txtCampo(2).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(2).BoundColumn = "Sequencia do geral"
   txtCampo(2).ListFields = "Razão social; Sequencia do geral"
   txtCampo(2).OrderFields = "Razão social; Sequencia do geral"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Geral"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(12).CtPri = txtCp(12)
   Set txtCampo(12).CtFdo = labFdo12
   Set txtCampo(12).CtBot(BOT_COMBO) = bottxtCampo12(BOT_COMBO)
   Set bottxtCampo12(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(12).DataType = 1
   txtCampo(12).Mask = "999999"
   txtCampo(12).BoundColumn = "Sequencia do geral"
   txtCampo(12).ListFields = "Razão social; Sequencia do geral"
   txtCampo(12).OrderFields = "Razão social; Sequencia do geral"
   txtCampo(12).Relation = ""
   txtCampo(12).Source = "Geral"
   txtCampo(12).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 4
   txtCampo(13).Mask = ""
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(14).CtPri = txtCp(14)
   Set txtCampo(14).CtFdo = labFdo14
   Set txtCampo(14).CtBot(BOT_LISTA) = bottxtCampo14(BOT_LISTA)
   Set txtCampo(14).CtBot(BOT_COMBO) = bottxtCampo14(BOT_COMBO)
   bottxtCampo14(BOT_LISTA).Caption = "P"
   Set bottxtCampo14(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(14).DataType = 1
   txtCampo(14).Mask = "999999"
   txtCampo(14).PesqModoAbertura =  2 
   txtCampo(14).PesqFieldCapture = "Geral.[Sequencia do geral]"
   txtCampo(14).BoundColumn = "Sequencia do geral"
   txtCampo(14).ListFields = "Razão social; Endereço; Nro; Bairro; Cpf e cnpj"
   txtCampo(14).OrderFields = "Razão social; Sequencia do geral"
   txtCampo(14).Relation = ""
   txtCampo(14).Source = "Geral"
   txtCampo(14).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(3).CtPri = txtCp(3)
   Set txtCampo(3).CtFdo = labFdo3
   Set txtCampo(3).CtBot(BOT_COMBO) = bottxtCampo3(BOT_COMBO)
   Set bottxtCampo3(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(3).DataType = 0
   txtCampo(3).ListFields = "Vista|Prazo"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "99.999.999,99"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
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
   txtCampo(6).Editable = False
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

   Set txtCampo(10).CtPri = txtCp(10)
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
      Case 2
         txtCampo(2).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0 And Cliente = 1", "[Sequencia do geral] > 0 And Cliente = 1") & ""
                           txtCampo(2).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                              "Geral.Bairro, Municipios.[Descrição do municipio], Municipios.Uf, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], " + _
                                                              "Geral.[Rg e ie] FROM Geral, Municipios WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND (Geral.Cliente = " + CStr(1) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.[Sequencia do municipio] = Municipios.[Sequencia do municipio])"
      Case 12
         txtCampo(12).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0 And Vendedor = 1", "[Sequencia do geral] > 0 And Vendedor = 1") & ""
      Case 14
         txtCampo(14).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0 And Transportadora = 1", "[Sequencia do geral] > 0 And Transportadora = 1") & ""
                           txtCampo(14).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                               "Geral.Bairro, Municipios.[Descrição do municipio], Municipios.Uf, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], " + _
                                                               "Geral.[Rg e ie] FROM Geral, Municipios WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND (Geral.Inativo = " + CStr(0) + ") AND " + _
                                                               "(Geral.Transportadora = " + CStr(1) + ") AND (Geral.[Sequencia do municipio] = Municipios.[Sequencia do municipio])"
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
   VerificaFormulario
   FinalizaForm Me
   Set lblAjuste = Nothing
   Set lblFechamento = Nothing
   Set Aba1 = Nothing
   Set GrdProdutos = Nothing
   Set lblParcelamento = Nothing
   Set grdParcelamento = Nothing
   Set grdServicos = Nothing
   Set cmdAbrir = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   If Not Vendas Is Nothing Then
      Set Vendas = Nothing
   End If
   If Not Produtos_da_Venda Is Nothing Then
      Produtos_da_Venda.CloseRecordSet
      Set Produtos_da_Venda = Nothing
   End If
   If Not Servicos_da_Venda Is Nothing Then
      Servicos_da_Venda.CloseRecordSet
      Set Servicos_da_Venda = Nothing
   End If
   If Not Parcelas_da_Venda Is Nothing Then
      Parcelas_da_Venda.CloseRecordSet
      Set Parcelas_da_Venda = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmVendas = Nothing                        'libera o segmento de código do form
End Sub

'evento - quando qq tecla for digitada no grid filho
Private Sub Grid_KeyPress(Index As Integer, ByVal KeyAscii As Integer, ByVal Shift As Integer, vgColumns() As Variant)
   Select Case Index
      Case 0
         ExecutaGrid0 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
      Case 2
         ExecutaGrid2 vgColumns(), KEYPRESS_NO_GRID, , , , , , KeyAscii
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
      Case 2
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
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
      Case 2
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Produtos_da_Venda![Sequencia da Venda] = Vendas![Sequencia da Venda]
      Case 1
         Parcelas_da_Venda![Sequencia da Venda] = Vendas![Sequencia da Venda]
      Case 2
         Servicos_da_Venda![Sequencia da Venda] = Vendas![Sequencia da Venda]
   End Select
End Sub

'evento - antes de efetuar o delete no recordset do grid
Private Sub Grid_BeforeDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, True
   ExecutaGrid Index, vgColumns(), EXCLUSOES, vgItem, 0, vgIsValid, 0, vgErrorMessage
   Select Case Index
      Case 0
      Case 1
      Case 2
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
         Set Produtos_da_Venda = vgNewRecordSet
      Case 1
         Set Parcelas_da_Venda = vgNewRecordSet
      Case 2
         Set Servicos_da_Venda = vgNewRecordSet
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
      Case 2
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
      Case  2 
         ExecutaGrid = ExecutaGrid2(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_item As Long, Sequencia_do_produto As Double, Quantidade As Double
   Dim Valor_unitario As Double, Valor_total As Double, Valor_do_ipi As Double
   Dim Aliquota_do_ipi As Single, Valor_icms_st As Double
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
            vgRetVal = (InfoProdutos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st, "Sigla"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 4
            vgRetVal = (InfoProdutos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st, "Estoque"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      Comandos KeyCodeAscii, Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st 
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
      vgRetVal = "SELECT * FROM [Produtos da Venda]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Venda] = " & Vendas![Sequencia da Venda]
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
         Case 6
            Valor_unitario = InfoProdutos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st, "Valor Unitário")
            If Grid(0).Col = 6 Then
               vgRetVal = Valor_unitario
            End If
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
                                                  "Produtos.[Ultima nfe], Produtos.[Ultima entrada], Produtos.[Ultimo fornecedor], Produtos.[Ultima venda] FROM Produtos WHERE (Produtos.[Sequencia do produto] > " + CStr(0) + ") AND " + _
                                                  "(Produtos.Inativo = " + CStr(0) + ")"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_produto = 0 Then
            vgIsValid = (Sequencia_do_produto > 0)
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "Produto é Obrigatotio!"
         End If
         If vgIsValid And (vgCol = -1 Or vgCol =  1 ) Then
            If Sequencia_do_produto > 0 Then
               vgIsValid = (ValidaProduto(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st))
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "Impossivel  - (Produto Inativo)"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Quantidade > 0 And Not OrcamentoAberto Then
               vgIsValid = (ValidaEstoque(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st))
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Estoque Insuficiente!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Quantidade = 0 Then
               vgIsValid = (Quantidade > 0)
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Quantidade deve ser maior que zero!"
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
            AtuEstoque Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st 
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AtuEstoque Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AtuEstoque Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st 
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Vendas.Edit
         Set vgRsError = Vendas
         If ProcessaProdutos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Aliquota_do_ipi, Valor_icms_st) Then
            Vendas![Sequencia da Venda] = (0)
         End If
         Vendas.Update
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
   Valor_do_ipi = ColumnValue(8)
   Aliquota_do_ipi = ColumnValue(9)
   Valor_icms_st = ColumnValue(10)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Produtos_da_Venda.Eof = False And Produtos_da_Venda.BOF = False And Produtos_da_Venda.RecordCount > 0 Then
         Valor_total = Produtos_da_Venda![Valor total]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_da_Venda.Eof = False And Produtos_da_Venda.BOF = False And Produtos_da_Venda.RecordCount > 0 Then
      Sequencia_do_item = Produtos_da_Venda![Sequencia do item]
      Sequencia_do_produto = Produtos_da_Venda![Sequencia do produto]
      Quantidade = Produtos_da_Venda!Quantidade
      Valor_unitario = Produtos_da_Venda![Valor unitario]
      Valor_total = Produtos_da_Venda![Valor total]
      Valor_do_ipi = Produtos_da_Venda![Valor do ipi]
      Aliquota_do_ipi = Produtos_da_Venda![Aliquota do ipi]
      Valor_icms_st = Produtos_da_Venda![Valor icms st]
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
   If Vendas.EOF = False And Vendas.BOF = False Then
      Vendas.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Numero_da_parcela As Integer, Dias As Integer, Data_de_vencimento As Variant
   Dim Valor_da_parcela As Double, Especie_de_pagamento As String
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
            vgRetVal = Not ((Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo")
         Case 3
            vgRetVal = Not ((Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo")
         Case 4
            vgRetVal = Not ((Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo")
         Case 5
            vgRetVal = Not (Not Vazio(Vendas![Forma de pagamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (((Vendas!Cancelado = 0) AND Not Vazio(Vendas![Forma de pagamento]) And TotalParcelas < Vendas![Valor Total da Venda]) AND Vendas!Fechado = 0)
      On Error Resume Next
         Grid(1).AllowEdit = ((((Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Sequencia da nota fiscal] = 0) AND Vendas!Fechado = 0) AND Vendas!Cancelado = 0)
      On Error Resume Next
         Grid(1).AllowDelete = ((((Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Sequencia da nota fiscal] = 0) AND Vendas!Fechado = 0) AND Vendas!Cancelado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas da Venda]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Venda] = " & Vendas![Sequencia da Venda]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D", Vendas![Data da Venda] , Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Vendas![Data da Venda])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Vendas![Valor Total da Venda] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If (Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo" Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo" Then
               vgIsValid = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
               If Not vgIsValid Then vgColumn =  3 
               vgErrorMessage$ = "Data de Vencimento não pode ser em Branco!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Vendas![Forma de pagamento])) AND Vendas![Forma de pagamento] = "Prazo" Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Vendas![Forma de pagamento] = "Prazo" , TotalParcelas() + Valor_da_Parcela <= Vendas![Valor Total da Venda], Valor_da_parcela = Vendas![Valor Total da Venda]), IIf(Vendas![Forma de pagamento] = "Vista" , Valor_da_parcela = Vendas![Valor Total da Venda], Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Vendas![Valor Total da Venda])))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total da Venda"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Vendas![Forma de pagamento]) Then
               vgIsValid = (Not Vazio(Especie_de_pagamento))
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Especie de Pagamento não pode ser vazio!"
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
   If Parcelas_da_Venda.Eof = False And Parcelas_da_Venda.BOF = False And Parcelas_da_Venda.RecordCount > 0 Then
      Numero_da_parcela = Parcelas_da_Venda![Numero da parcela]
      Dias = Parcelas_da_Venda!Dias
      Data_de_vencimento = Parcelas_da_Venda![Data de vencimento]
      Valor_da_parcela = Parcelas_da_Venda![Valor da parcela]
      Especie_de_pagamento = Parcelas_da_Venda![Especie de pagamento]
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


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid2(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_item As Long, Sequencia_do_Servico As Long, Quantidade As Double
   Dim Valor_unitario As Double, Valor_total As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(2).Status <> ACAO_NAVEGANDO And vgItem = Grid(2).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 5
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosServicos KeyCodeAscii, Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total 
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Serviços da Venda]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Venda] = " & Vendas![Sequencia da Venda]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

      'vamos definir a ordenação
      x$ = "[Sequencia do item]"
      vgRetVal = InsereSQL(vgRetVal, EXP_ORDERBY, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 2
            Sequencia_do_item = SeqItemServico()
            vgRetVal = Sequencia_do_item
         Case 4
            Valor_unitario = InfoServicos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total, "Valor")
            If Grid(2).Col = 4 Then
               vgRetVal = Valor_unitario
            End If
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" + FiltroServico() + ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Serviços.[Sequencia do Serviço], Serviços.[Descrição do Serviço], Serviços.[Valor do Serviço] FROM Serviços WHERE (Serviços.[Sequencia do Serviço] > " + CStr(0) + ") AND " + _
                                                  "(Serviços.Inativo = " + CStr(0) + ")"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_Servico > 0 Then
            vgIsValid = (ValidaServico(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total) )
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "IMOSSIVEL SERVIÇO INATIVO!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Servico = 0 Then
               vgIsValid = (Sequencia_do_Servico > 0)
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "IMFORME O SERVIÇO!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Quantidade inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_unitario > 0)
            If Not vgIsValid Then vgColumn =  4 
            vgErrorMessage$ = "Valor unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(2) = SeqItemServico()
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValores  
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValores  
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores  
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Vendas.Edit
         Set vgRsError = Vendas
         If ProcessaServicos(Sequencia_da_Venda, Sequencia_do_item, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total) Then
            Vendas![Sequencia da Venda] = (0)
         End If
         Vendas.Update
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
   Sequencia_do_Servico = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Sequencia_do_item = ColumnValue(2)
   Quantidade = ColumnValue(3)
   Valor_unitario = ColumnValue(4)
   If Grid(2).Status <> ACAO_INCLUINDO Then
      If Servicos_da_Venda.Eof = False And Servicos_da_Venda.BOF = False And Servicos_da_Venda.RecordCount > 0 Then
         Valor_total = Servicos_da_Venda![Valor total]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Servicos_da_Venda.Eof = False And Servicos_da_Venda.BOF = False And Servicos_da_Venda.RecordCount > 0 Then
      Sequencia_do_item = Servicos_da_Venda![Sequencia do item]
      Sequencia_do_Servico = Servicos_da_Venda![Sequencia do Serviço]
      Quantidade = Servicos_da_Venda!Quantidade
      Valor_unitario = Servicos_da_Venda![Valor unitario]
      Valor_total = Servicos_da_Venda![Valor total]
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
   If Vendas.EOF = False And Vendas.BOF = False Then
      Vendas.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid2 = vgRetVal
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
      Case 2
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
      Case 1
         Fechamento  
      Case 2
         AbreVenda  
      Case 3
         AbreGerar  
      Case 4
         Cancelamento  
      Case 5
         frmProcura.Show vbModal
      Case 6
         AbreRelatorio  
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo1_Click(Index As Integer)
   txtCampo(1).SetFocus
   DoEvents
   txtCampo(1).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
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
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


