Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOServico
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8445
   Left                     =   2910
   LinkTopic                =   "OServico"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8010
   ScaleWidth               =   14370
   Top                      =   1675
   Width                    =   14460
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
      Height                   =   7830
      Left                     =   15
      Width                    =   14190
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   25016
      _extenty                 =   13804
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1440
         Left                     =   120
         Width                    =   1395
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
         Height                   =   435
         Left                     =   1800
         Width                    =   3180
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
         Caption                  =   "Ordem de Serviço"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   210
         Left                     =   1800
         Width                    =   3255
         Top                      =   585
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Courier New"
            CharSet                  =   0
            Weight                   =   400
            Size                     =   8.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000008
         Caption                  =   "_______________________________"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   240
         Width                    =   1155
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
         ForeColor                =   &H80000008
         Caption                  =   "Sequencia:"
         Index                    =   2 
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
         Left                     =   1440
         Width                    =   795
         Top                      =   1680
         DataField                =   "Sequencia da Ordem"
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
         Width                    =   1230
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
         ForeColor                =   &H80000008
         Caption                  =   "Dt. da Ordem:"
         Index                    =   3 
      End
      Begin VB.Label labFdo1
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   3780
         Top                      =   1680
         Width                    =   1380
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo1
         Style                    =   1
         TabStop                  =   0
         Index                    =   0
         Left                     =   4890
         Top                      =   1695
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
         Left                     =   3795
         Width                    =   1080
         Top                      =   1695
         DataField                =   "Data da Ordem"
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
         Width                    =   540
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
         ForeColor                =   &H80000008
         Caption                  =   "Geral:"
         Index                    =   4 
      End
      Begin VB.Label labFdo2
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6060
         Top                      =   1680
         Width                    =   4200
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   9990
         Top                      =   1695
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo2
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   9735
         Top                      =   1695
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
         Left                     =   6075
         Width                    =   3645
         Top                      =   1695
         DataField                =   "Sequencia do geral"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   13920
         Width                    =   45
         Top                      =   7920
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
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   10440
         Width                    =   1020
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
         ForeColor                =   &H80000008
         Caption                  =   "Vendedor:"
         Index                    =   6 
      End
      Begin VB.Label labFdo3
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   11580
         Top                      =   1680
         Width                    =   2400
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   13710
         Top                      =   1695
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo3
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   13455
         Top                      =   1695
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
         Left                     =   11595
         Width                    =   1845
         Top                      =   1695
         DataField                =   "Sequencia do Vendedor"
         Index                    =   3 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   4
         TabHeight                =   519
         BackColor                =   &H8000000F
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   4
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   3675
         Left                     =   240
         Width                    =   13710
         Top                      =   2580
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
         Caption                  =   "Virabrequim"
         Index                    =   0 
         _version                 =   65536
         _extentx                 =   24170
         _extenty                 =   6479
         _stockprops              =   15
         TabCaption(0)            =   "Virabrequim"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "Itens"
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Grid(1)"
         TabCaption(2)            =   "Serviços"
         Tab(2).ControlCount=   1
         Tab(2).ControlEnabled=   0
         Tab(2).Control(0)=   "Grid(2)"
         TabCaption(3)            =   "Financeiro"
         Tab(3).ControlCount=   29
         Tab(3).ControlEnabled=   0
         Tab(3).Control(0)=   "Label(7)"
         Tab(3).Control(1)=   "labFdo4"
         Tab(3).Control(2)=   "bottxtCampo4(2)"
         Tab(3).Control(3)=   "txtCp(4)"
         Tab(3).Control(4)=   "Label(8)"
         Tab(3).Control(5)=   "txtCp(5)"
         Tab(3).Control(6)=   "Label(9)"
         Tab(3).Control(7)=   "txtCp(6)"
         Tab(3).Control(8)=   "Label(10)"
         Tab(3).Control(9)=   "txtCp(7)"
         Tab(3).Control(10)=   "Label(11)"
         Tab(3).Control(11)=   "Label(12)"
         Tab(3).Control(12)=   "txtCp(8)"
         Tab(3).Control(13)=   "Label(13)"
         Tab(3).Control(14)=   "txtCp(9)"
         Tab(3).Control(15)=   "Label(14)"
         Tab(3).Control(16)=   "txtCp(10)"
         Tab(3).Control(17)=   "txtCp(11)"
         Tab(3).Control(18)=   "Label(15)"
         Tab(3).Control(19)=   "Label(16)"
         Tab(3).Control(20)=   "Formato(0)"
         Tab(3).Control(21)=   "Label(17)"
         Tab(3).Control(22)=   "Formato(1)"
         Tab(3).Control(23)=   "Label(18)"
         Tab(3).Control(24)=   "Formato(2)"
         Tab(3).Control(25)=   "Label(19)"
         Tab(3).Control(26)=   "Grid(3)"
         Tab(3).Control(27)=   "Label(20)"
         Tab(3).Control(28)=   "txtCp(12)"
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
            Height                   =   2880
            Left                     =   180
            Width                    =   13335
            Top                      =   540
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   0 
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
            TabIndex                 =   7
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2880
            Left                     =   -74820
            Width                    =   13335
            Top                      =   540
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
            TabIndex                 =   8
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2880
            Left                     =   -74820
            Width                    =   13335
            Top                      =   540
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -74610
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
            Index                    =   7 
         End
         Begin VB.Label labFdo4
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -72420
            Top                      =   660
            Width                    =   1635
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo4
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -71055
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
            TabIndex                 =   9
            TabStop                  =   -1
            Height                   =   270
            Left                     =   -72405
            Width                    =   1350
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
            Left                     =   -73980
            Width                    =   1380
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
            Caption                  =   "Valor dos Itens:"
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
            TabIndex                 =   10
            TabStop                  =   0
            Height                   =   300
            Left                     =   -72420
            Width                    =   1635
            Top                      =   1200
            DataField                =   "Valor total dos produtos"
            TabStop                  =  0
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
            Left                     =   -74355
            Width                    =   1755
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
            Caption                  =   "Valor dos Serviços:"
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
            Left                     =   -72420
            Width                    =   1635
            Top                      =   1740
            DataField                =   "Valor Total dos Serviços"
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
            Left                     =   -74865
            Width                    =   585
            Top                      =   3240
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
            Caption                  =   "% ISS:"
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   12
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -74160
            Width                    =   660
            Top                      =   3240
            DataField                =   "Aliquota do Iss"
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
            Left                     =   -73320
            Width                    =   1170
            Top                      =   3240
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
            Caption                  =   "Valor do ISS:"
            Index                    =   11 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -70500
            Width                    =   1725
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
            Caption                  =   "Valor do Desconto:"
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
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   13
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -68640
            Width                    =   1395
            Top                      =   660
            DataField                =   "Valor do desconto"
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
            Left                     =   -70095
            Width                    =   1320
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
            Caption                  =   "Valor do Frete:"
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
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   14
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -68625
            Width                    =   1380
            Top                      =   1200
            DataField                =   "Valor do frete"
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
            Left                     =   -69900
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
            Caption                  =   "Total da OS:"
            Index                    =   14 
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
            Left                     =   -68640
            Width                    =   1395
            Top                      =   1740
            DataField                =   "Valor da OS"
            TabStop                  =  0
            Index                    =   10 
            MultiLine                =   -1
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   16
            TabStop                  =   0
            Height                   =   300
            Left                     =   -71970
            Width                    =   1185
            Top                      =   3240
            DataField                =   "Valor do Iss"
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
            Left                     =   -66525
            Width                    =   5175
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
            Index                    =   15 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -66540
            Width                    =   600
            Top                      =   3375
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
            Index                    =   16 
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
            Left                     =   -65865
            Width                    =   165
            Top                      =   3420
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65640
            Width                    =   990
            Top                      =   3375
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
            Index                    =   17 
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
            Left                     =   -64545
            Width                    =   165
            Top                      =   3420
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -64245
            Width                    =   1050
            Top                      =   3375
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
            Index                    =   18 
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
            Left                     =   -63030
            Width                    =   150
            Top                      =   3420
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -62805
            Width                    =   300
            Top                      =   3375
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
            Index                    =   19 
         End
         Begin NFE.GListV Grid
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
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
            TabIndex                 =   17
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2625
            Left                     =   -66525
            Width                    =   5160
            Top                      =   660
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   3 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -70620
            Width                    =   2115
            Top                      =   3240
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
            ForeColor                =   &HFF
            Caption                  =   "Data de Cancelamento:"
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
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   25
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -68400
            Width                    =   1140
            Top                      =   3240
            DataField                =   "Data de Cancelamento"
            TabStop                  =  0
            Index                    =   12 
         End
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   240
         Width                    =   1455
         Top                      =   6300
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
         Caption                  =   "Observações:"
         Index                    =   21 
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
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   990
         Left                     =   240
         Width                    =   13680
         Top                      =   6585
         DataField                =   "Obs"
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
         Height                   =   375
         Left                     =   1800
         Width                    =   1005
         Top                      =   840
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
         Index                    =   22 
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
         Left                     =   11220
         Width                    =   1365
         Top                      =   615
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
         TabIndex                 =   20
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   11220
         Width                    =   1365
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
         TabIndex                 =   21
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   12615
         Width                    =   1365
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
         Index                    =   2 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   22
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   480
         Left                     =   12615
         Width                    =   1365
         Top                      =   620
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
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   4845
         Width                    =   1035
         Top                      =   2280
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
         Caption                  =   "Transporte:"
         Index                    =   23 
      End
      Begin VB.Label labFdo14
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   6060
         Top                      =   2280
         Width                    =   7935
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   1
         Left                     =   13725
         Top                      =   2295
         Width                    =   255
         Height                   =   270
      End
      Begin VB.CommandButton bottxtCampo14
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   13470
         Top                      =   2295
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
         Left                     =   6075
         Width                    =   7380
         Top                      =   2295
         DataField                =   "Sequencia do Transportador"
         Index                    =   14 
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   23
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   255
         Left                     =   6960
         Width                    =   480
         Top                      =   8580
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
         Caption                  =   "OK"
         Index                    =   4 
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
         MaxLength                =   254
         BorderStyle              =   1
         Alignment                =   1
         TabIndex                 =   24
         TabStop                  =   -1
         Height                   =   255
         Left                     =   6060
         Width                    =   840
         Top                      =   8580
         Index                    =   15 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   6060
         Width                    =   1335
         Top                      =   1380
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
         ForeColor                =   &HFF
         Caption                  =   "Inadimplente"
         Index                    =   24 
      End
      Begin VB.Frame Frame
         Height                   =   1020
         Left                     =   6090
         Width                    =   4215
         Top                      =   120
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
            DataField                =   "Qual Empresa"
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
            TabIndex                 =   26
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
            TabIndex                 =   27
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
End
Attribute VB_name = "frmOServico"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: OServico
'* Função....: O.S
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:45
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
Dim txtCampo(15) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim opcFrame0(1) As New FormataCampos
Dim Sequencia_da_Ordem As Long, Data_da_Ordem As Variant, Sequencia_do_geral As Long
Dim Sequencia_do_Vendedor As Long, Forma_de_pagamento As String, Valor_total_dos_produtos As Double
Dim Valor_Total_dos_Servicos As Double, Valor_do_desconto As Double, Valor_do_frete As Double
Dim Valor_da_OS As Double, Aliquota_do_Iss As Single, Valor_do_Iss As Double
Dim Cancelado As Boolean, Fechado As Boolean, Obs As String
Dim Usuario_da_Alteracao As String, Data_da_Alteracao As Variant, Hora_da_Alteracao As Variant
Dim Sequencia_da_nota_fiscal As Long, Sequencia_do_Transportador As Long, Data_de_Cancelamento As Variant
Dim Qual_Empresa As Integer, troca_seq As Double
Public lblOrdem As Object, lblLinha As Object, lblAjuste As Object
Public Aba1 As Object, GrdVirabrequim As Object, grdItens As Object
Public grdServicos As Object, lblParcelamento As Object, grdParcelamento As Object
Public lblFechamento As Object, cmdAbrir As Object, cmdFechar As Object
Dim Ordem_de_Servico As New GRecordSet, Virabrequim_da_OS As New GRecordSet, Itens_da_OS As New GRecordSet
Dim Servicos_da_OS As New GRecordSet, Parcelas_da_os As New GRecordSet

Dim VirabrequimAux As New GRecordSet, MarcaAux As New GRecordSet, VirabrequimAux2 As New GRecordSet
Dim MarcaAux2 As New GRecordSet, ModeloAux As New GRecordSet

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
    Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia da Ordem]) Seq From [Ordem de Serviço]")
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

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Function InfoVirabrequim(Sequencia_da_Ordem As Long, Sequencia_do_Virabrequim As Long, Quantidade As Double, _
   Valor_unitario As Double, Tem_Casco As Boolean, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Virabrequim", "[Sequencia do Virabrequim] = " & Sequencia_do_Virabrequim, VirabrequimAux
      
   If VirabrequimAux.RecordCount = 0 Then Exit Function
   
   TbAuxiliar "Marcas", "[Sequencia da Marca] = " & VirabrequimAux![Sequencia da Marca], MarcaAux
   TbAuxiliar "Modelo do Virabrequim", "[Sequencia do Modelo] = " & VirabrequimAux![Sequencia Do Modelo], ModeloAux
   
   Select Case Oq      
      Case "Modelo"
         InfoVirabrequim = VirabrequimAux![Modelo]
      Case "Marca"
         InfoVirabrequim = MarcaAux![Descrição]
      Case "Valor"
         InfoVirabrequim = VirabrequimAux![Valor Do Virabrequim] 
      Case "ModeloA"
         InfoVirabrequim = ModeloAux![Descrição do Modelo]      
   End Select

End Function


Private Function ProcessaItens(Sequencia_da_Ordem As Long, Sequencia_do_produto As Double, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double)
         
   On Error GoTo DeuErro
   
   vgDb.BeginTrans      
   vgDb.Execute "Update [Itens da OS] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem & " AND [Sequencia do Produto] = " & Sequencia_do_Produto
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Function ProcessaServicos(Sequencia_da_Ordem As Long, Sequencia_do_Servico As Long, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Sequencia_do_Virabrequim As Long, _
   Sequencia_do_item As Long)
         
   On Error GoTo DeuErro
   
   vgDb.BeginTrans      
   vgDb.Execute "Update [Serviços da OS] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem & " AND [Sequencia do Item] = " & Sequencia_do_Item
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
End Function


Private Sub CarregaTotalizador()
   On Error Resume Next

   With grdItens
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowFilterBar = False
      .HideStatus = True
   End With
   
   With grdServicos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
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
   
   BloqueiaAba
   
End Sub


Public Sub AjustaValores()
   Dim ValorItens As New GRecordSet
   Dim ValorServicos As New GRecordSet
   Dim ValorOrdem As Currency, ValorISS As Currency
   Dim Itens As New GRecordSet, Servicos As New GRecordSet
 
   On Error GoTo DeuErro
   
   Set Itens = VgDb.OpenRecordSet("SELECT * From [Itens da OS] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
   Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços da OS] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
   '====================================================================================================================='      
   Set ValorItens = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Itens da OS] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem) 
   Set ValorServicos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Serviços da OS] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem) 
   ValorOrdem = ValorItens!Total + ValorServicos!Total + Valor_do_Frete - Valor_do_desconto
   ValorOrdem = Format(ValorOrdem, "##,###,##0.00")
   
   If Empresa![Aliquota Do Iss] > 0 Then
      ValorISS = ValorServicos!Total * Empresa![Aliquota Do Iss] / 100
      ValorISS = Format(ValorISS, "##,###,##0.00")
   Else
      ValorISS = 0
   End If    
   
   If Itens.RecordCount = 0 And Servicos.RecordCount = 0 Then
      ValorOrdem = 0 
      If Valor_do_frete > 0 Or Valor_do_desconto > 0 Then 
         vgDb.BeginTrans   
         vgDb.Execute "Update [Ordem de Serviço] Set [Valor do frete] = 0 WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
         vgDb.Execute "Update [Ordem de Serviço] Set [Valor do desconto] = 0 WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
         vgDb.CommitTrans
      End If  
   End If

   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Ordem de Serviço] Set [Valor total dos produtos] = " & Substitui(ValorItens!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem 
   vgDb.Execute "Update [Ordem de Serviço] Set [Valor total dos Serviços] = " & Substitui(ValorServicos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem 
   vgDb.Execute "Update [Ordem de Serviço] Set [Valor da OS] = " & Substitui(CStr(ValorOrdem), ",", ".", SO_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem 
   vgDb.Execute "Update [Ordem de Serviço] Set [Valor do ISS] = " & Substitui(CStr(ValorISS), ",", ".", SO_UM) & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem 
   vgDb.CommitTrans
     
   Alteracao
     
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


Private Sub RepositionOS()
   On Error Resume Next
   
   If vgSituacao = ACAO_INCLUINDO Then
      Set Virabrequim_da_OS = Nothing
      Set Itens_da_OS = Nothing
      Set Servicos_da_OS = Nothing
      Set Parcelas_da_OS = Nothing
   End If
   
   CarregaTotalizador
   'Alteração
   lblAlteracao.Caption = Ordem_de_Servico![Usuario da Alteração] & " " & Ordem_de_Servico![Data da Alteração] & " " & Ordem_de_Servico![Hora da Alteração] 
  
End Sub


Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC FROM [Parcelas da os] WHERE [Sequencia da ordem] = " & Sequencia_da_Ordem)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Function UltimoVencimento() As Variant
 Dim Tb As New GRecordSet, RetVal As Variant
   
  On Error Resume Next
      
  Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de vencimento]) [Data de Vencimento] FROM [Parcelas da Os] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem)
      
  RetVal = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, grdParcelamento.Col - 2)
  
   If Ordem_de_Servico![Forma de pagamento] = "Prazo" Then
      If Not IsNull(Tb![Data de Vencimento]) Then
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Ordem_de_Servico![Data da Ordem]
      End If
   Else
      UltimoVencimento = RetVal = Ordem_de_Servico![Data da Ordem]
   End If
  
End Function


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


Private Sub ChangeGeral()
 Dim Tb As New GRecordSet

 Set Tb = vgDb.OpenRecordSet("SELECT * From [Virabrequim da Os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 
 If Tb.RecordCount > 0 Then
     vgDb.Execute "DELETE FROM [Virabrequim da Os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem  
 End If


End Sub


Private Sub Comandos(KeyAscii As Integer, Sequencia_da_Ordem As Long, Sequencia_do_Virabrequim As Long, Quantidade As Double, _
   Valor_unitario As Double, Tem_Casco As Boolean)
   On Error GoTo DeuErro
   
   If KeyAscii = vbKeyF12 Then
      With GrdVirabrequim
         Select Case .ColumnField(.Col)
            Case "Sequencia do Virabrequim"
               seqRegistro = .ColumnValue(.Row + 1, .Col)
               frmVirabreq.Show        
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
Dim Tb1 As New GRecordSet, Tb2 As New GRecordSet

 If vgSITUACAO = ACAO_EXCLUINDO Then
    Set Tb = vgDb.OpenRecordSet("SELECT * From [Virabrequim da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
    Set Tb1 = vgDb.OpenRecordSet("SELECT * From [Itens da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
    Set Tb2 = vgDb.OpenRecordSet("SELECT * From [Serviços da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
    Set Parcelas = vgDb.OpenRecordSet("SELECT * From [Parcelas da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
  
     If Parcelas.RecordCount > 0 Then
        vgDb.Execute "DELETE FROM [Parcelas da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
        'DeletaConta
     End If
       If Tb.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Virabrequim da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
       End If
     If Tb1.RecordCount > 0 Then
         vgDb.Execute "DELETE FROM [Itens da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
     End If
     If Tb2.RecordCount > 0 Then
         vgDb.Execute "DELETE FROM [Serviços da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
     End If 
 
 End If    
      
End Sub


Private Sub MostraFormulas()

   If Fechado And Not Cancelado Then
      lblFechamento.Caption = "O.S Fechada"
      lblFechamento.ForeColor = &H0
      lblOrdem.Caption = "Ordem de Serviço"
      lblOrdem.ForeColor = &H0
      lblLinha.ForeColor = &H0
   Else
    If Not Cancelado Then
      lblOrdem.Caption = "Ordem de Serviço"
      lblFechamento.Caption = "O.S Aberta"
      lblFechamento.ForeColor = &H0
      lblOrdem.ForeColor = &H0
      lblLinha.ForeColor = &H0
    End If  
   End If
   If Err Then Err = 0: lblFechamento.Caption = ""   
   If Cancelado Then
      lblOrdem.ForeColor = &H8FF
      lblLinha.ForeColor = &H8FF
      lblOrdem.Caption = "Ordem de Serviço"
      lblFechamento.Caption = "Cancelada"
      lblFechamento.ForeColor = &H8FF
   End If
   If Err Then Err = 0: lblOrdem.Caption = "" 
   If Err Then Err.Clear 
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_da_os) Or (TotalParcelas > Valor_da_os) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   
End Sub


Private Sub Fechamento()
 Dim Virabrequim As New GRecordSet, Parcelas As New GRecordSet
 Dim Manutencao As New GRecordSet, Servicos As New GRecordSet
 Dim Cascos As New GRecordSet, ManuCasco As New GRecordSet
 
 On Error Resume Next
 
 Set Virabrequim = VgDb.OpenRecordSet("SELECT * From [Virabrequim da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 Set Servicos = VgDb.OpenRecordSet("SELECT * From [Serviços da os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 Set Cascos = VgDb.OpenRecordSet("SELECT * From [Virabrequim da os] Where [Tem Casco] = 1 And [Sequencia da Ordem] = " & Sequencia_da_Ordem)
  
 'If Virabrequim.RecordCount = 0 Then
 '   MsgBox("Impossivel Fechar O.S! (Falta o Virabrequim)")
 '   Exit Sub
 'End If
 
  If Parcelas.RecordCount > 0 Then
   If TotalParcelas() <> Valor_da_os Then 
      MsgBox("Impossivel Fechar O.S!! A Soma das Parcelas não Correspondem ao Valor Total da Ordem")
   Exit Sub
   End If
  End If
  
  If Servicos.RecordCount = 0 Then
    MsgBox("Impossivel Fechar O.S! (Falta o Serviço Executado)")
    Exit Sub
  End If
  
  If Parcelas.RecordCount > 0 And Manutencao.RecordCount = 0 Then
     LancaParcelas 
  Else
     DeletaConta
     LancaParcelas     
  End If
  
  If Parcelas.RecordCount = 0  Then If MsgBox("Deseja Fechar a O.S sem Incluir as Parcelas do Contas a Receber do Cliente?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
     vgDb.Execute "UPDATE [Ordem de Serviço] Set Fechado = 1 WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
  Reposition True
  
  If Virabrequim.RecordCount > 0 Then
      Do While Not Virabrequim.EOF
         vgDb.Execute "UPDATE Virabrequim Set [Ultima OS] = " & D(Ordem_de_Servico![Data da Ordem]) & " WHERE [Sequencia do Virabrequim] = " & Virabrequim![Sequencia Do Virabrequim]
         vgDb.Execute "UPDATE Virabrequim Set [Sequencia da Ordem] = " & Sequencia_da_Ordem  & " WHERE [Sequencia do Virabrequim] = " & Virabrequim![Sequencia Do Virabrequim]
      Virabrequim.MoveNext
      Loop
      Virabrequim.MoveFirst      
  End If
  
  If Cascos.RecordCount > 0 Then
   Set ManuCasco = VgDb.OpenRecordSet("Manutenção Cascos")
    Do While Not Cascos.EOF
       With ManuCasco
         .AddNew
           !Documento = "Os Nº " & Format(Sequencia_da_Ordem,"000000")
           ![Sequencia Do geral] = Ordem_de_Servico![Sequencia Do geral]
           ![Sequencia Do Virabrequim] = Cascos![Sequencia do Virabrequim]
           ![Data de entrada] = Ordem_de_Servico![Data da Ordem] 
           ![Sequencia da Ordem] = Sequencia_da_Ordem 
         .Update
         .BookMark = .LastModified
       End With   
      Cascos.MoveNext
    Loop      
  End If
              
End Sub


Private Sub AbreOS()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean
 Dim Manu As New GRecordSet
 Dim ManuCasco As New GRecordSet

 On Error Resume Next
 
 Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas Receber] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas receber] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem & ")")
 Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas receber] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 Set ManuCasco = vgDb.OpenRecordSet("SELECT * From [Manutenção Cascos] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
  
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If
   
   If TemBaixa Then
      MsgBox("Impossivel abrir a O.S existe Baixa(s) no Contas a Receber para Abri-la Exclua a(s) Baixa(s)")
      Exit Sub
   End If
    
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "UPDATE [Ordem de Serviço] Set Fechado = 0 Where [Sequencia da Ordem] = " & Sequencia_da_Ordem 
       If Manu.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção contas receber] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
       End If 
       If ManuCasco.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção Cascos] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
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

 If Empresa![Nome fantasia] = "Reticom Virabrequim" Then         
    Aba1.TabVisible(1) = False
 End If   
    
End Sub

Private Sub LancaParcelas()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordSet
  
   On Error GoTo DeuErro
      
   If Vazio(Forma_de_pagamento) Then Exit Sub
   
   If TotalParcelas() <> Valor_da_OS Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas receber")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas da os] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas receber", "[Sequencia da Ordem] = " & Sequencia_da_Ordem & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew 
         !Parcela = Parcelas![Numero da parcela]
         ![Sequencia Do geral] = Sequencia_do_Geral
         !Documento = "OS Nº " & Format(Sequencia_da_Ordem,"00000")
         ![Data de entrada] = Data_da_Ordem
         !Historico = Obs
         ![Forma de Pagamento] = Forma_de_pagamento
         ![Especie de pagamento] = Parcelas![Especie de pagamento]
        ' If ![Forma de pagamento] = "Vista" Then
        '    ![Valor pago] = Parcelas![Valor da parcela]
        '    ![Valor restante] = 0
        '    ![Data da baixa] = Data_de_entrada
        ' Else
         ![Valor pago] = 0
         ![Valor restante] = Parcelas![Valor da parcela]
         ![Data da baixa] = Null
        ' End If
         ![Data de vencimento] = Parcelas![Data de vencimento]
         ![Valor total] = Valor_da_Os
         ![Valor da parcela] = Parcelas![Valor da parcela]
         ![Sequencia da Ordem] = Sequencia_da_Ordem
         .Update
         .BookMark = .LastModified
      End With
      Parcelas.MoveNext
   Loop
   
   'Baixa
  ' If Forma_de_pagamento = "Vista" Then
  '    Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas]")
  '    Manutencao.MoveFirst: Manutencao.MoveLast
  '    With Baixa
  '       .AddNew
  '       ![Seqüência da Manutenção] = Manutencao![Seqüência da Manutenção]
  '       ![Data da Baixa] = Manutencao![Data de Vencimento]
  '       ![Valor Pago] = Manutencao![Valor da Parcela]
  '       !Conta = Manutencao!Conta
  '       .Update
  '       .BookMark = .LastModified
  '       Baixa.MoveLast
  '    End With
  '    With Manutencao
  '       .Edit
  '       ![Seqüência da Baixa] = Baixa![Seqüência da Baixa]
  '       .Update
  '       .BookMark = .LastModified
  '    End With      
  ' End If
   
   vgDb.CommitTrans

DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Sub DeletaConta() 'Qdo exclui registro Mãe ou Abre OS
   Dim Manutencao As New GRecordset

   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem)
    
   If Manutencao.RecordCount = 0 Then Exit Sub
   
   vgDb.Execute "DELETE FROM [Manutenção contas Receber] " & _
                "WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem   

DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbokOnly, vaTitulo
      vgdb.RollBackTrans
   End If

End Sub


Private Sub CancelaOS()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean
 Dim Manu As New GRecordSet

 On Error Resume Next
   
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       Cancelamento
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  InputYgorsBom = ""
  Reposition True
  
End Sub  


Private Sub Cancelamento() 
 Dim Tb As New GRecordSet
   
   On Error GoTo DeuErro
      
   If MsgBox("Deseja Realmente Cancelar Essa O.S - As Baixas do Contas a Receber e a Conta Tbm serão Canceladas?" & vbCrLf & "", vbExclamation + vbYesNo, vaTitulo) = vbNo Then Exit Sub
   
   Set Tb = vgDb.OpenRecordSet("SELECT OS.[Sequencia da Ordem], " & _
                               "Manu.[Sequencia da manutenção] Conta " & _
                               "FROM [Ordem de Serviço] OS " & _
                               "LEFT JOIN [Manutenção contas Receber] Manu ON OS.[Sequencia da Ordem] = Manu.[Sequencia da Ordem] " & _
                               "WHERE OS.[Sequencia da Ordem] = " & Sequencia_da_Ordem)                             
   vgDb.BeginTrans 'Inicia a transação
       
   'Marcando Cancelada na Tabela da OS 
    vgDb.Execute "Update [Ordem de Serviço] Set [Cancelado] = 1 Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
    vgDb.Execute "Update [Ordem de Serviço] Set [Data de Cancelamento] = " & D(Date) & " Where [Sequencia da Ordem] = " & Sequencia_da_Ordem
    
    If Tb!Conta > 0 AND Not Vazio(Tb!Conta) Then
      Do While Not Tb.EOF 'Excluindo o Contas se houver
         vgDb.Execute "DELETE FROM [Baixa Contas receber] WHERE [Sequencia da Manutenção] = " & Tb!Conta 'Baixa
         'vgDb.Execute "DELETE FROM [Movimentação da Conta Corrente] WHERE [Seqüência do Lançamento] = " & Tb!Conta 'Conta Corrente
         Tb.MoveNext
      Loop
      Tb.MoveFirst 'Vamos Voltar no Começo
      vgDb.Execute "DELETE FROM [Manutenção Contas Receber] WHERE [Sequencia da Ordem] = " & Tb![Sequencia da Ordem] 'Deleta a Conta gerado pela OS
    End If   
    
    vgDb.CommitTrans 
   
    MsgBox "O.S Cancelada com Sucesso!!!", vbInformation + vbOKOnly, vaTitulo 'Suuuuuceeeeesssssoooo
     
DeuErro:
   If Err.Number <> 0 Then
      vgDb.RollBackTrans
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If
     
End Sub


Private Function Filtro() As String
 Dim Tb As New GRecordSet
 Dim PriVez As Boolean
 
 PriVez = True
 If Virabrequim_da_OS.RecordCount = 0 Then
    Filtro = "[Sequencia do Virabrequim] = -1"
    Exit Function
 End If
 
 Set Tb = vgDb.OpenRecordSet("SELECT * From [Virabrequim da Os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 
 Do While Not Tb.EOF
  If Privez Then
     Filtro = "[Sequencia do Virabrequim] = " & Tb![Sequencia Do Virabrequim]
  End If
     Privez = False
     Filtro = Filtro & " Or [Sequencia do Virabrequim] = " & Tb![Sequencia Do Virabrequim]
  Tb.MoveNext
 Loop

End Function


Private Function InfoVirabrequim2(Sequencia_da_Ordem As Long, Sequencia_do_Servico As Long, Quantidade As Double, _
   Valor_unitario As Double, Valor_total As Double, Sequencia_do_Virabrequim As Long, _
   Sequencia_do_item As Long, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Virabrequim", "[Sequencia do Virabrequim] = " & Sequencia_do_Virabrequim, VirabrequimAux2
      
   If VirabrequimAux2.RecordCount = 0 Then Exit Function
   
   TbAuxiliar "Marcas", "[Sequencia da Marca] = " & VirabrequimAux2![Sequencia da Marca], MarcaAux2
   
   Select Case Oq      
      Case "Modelo"
         InfoVirabrequim2 = VirabrequimAux2![Modelo]
      Case "Marca"
         InfoVirabrequim2 = MarcaAux2![Descrição]
      Case "Valor"
         InfoVirabrequim2 = VirabrequimAux2![Valor Do Virabrequim]   
   End Select

End Function


Private Function SeqItem() As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Serviços da Os] Where [Sequencia da Ordem] = " & Sequencia_da_Ordem)
 
    If Tb.RecordCount > 0 Then
      SeqItem = Tb!Seq + 1
    Else
      SeqItem = 1
    End If
      
End Function


Private Function TemOS(Sequencia_da_Ordem As Long, Sequencia_do_Virabrequim As Long, Quantidade As Double, _
   Valor_unitario As Double, Tem_Casco As Boolean) As Boolean
 Dim ViOs As New GRecordSet, Os As New GRecordSet

 Set ViOs = vgDb.OpenRecordSet("SELECT * FROM [Virabrequim da Os] LEFT JOIN [Ordem de Serviço] On [Virabrequim da Os].[Sequencia da Ordem] = [Ordem de Serviço].[Sequencia da Ordem] And [Ordem de Serviço].Cancelado = 0 " & _  
                               "WHERE [Sequencia do Virabrequim] = " & Sequencia_do_Virabrequim & " And [Ordem de Serviço].[Sequencia da Ordem] <> " & Sequencia_da_Ordem)
 
   If ViOs.RecordCount > 0 Then
      TemOs = False
      MDINFE.CancelaAlteracoes: Exit Function 
   Else
      TemOs = True: Exit Function        
   End If
   
End Function 


Private Function ValidaVencto(Sequencia_da_Ordem As Long, Numero_da_parcela As Integer, Dias As Integer, _
   Data_de_vencimento As Variant, Valor_da_parcela As Double, Especie_de_pagamento As String) As Boolean
 
   If IsDate(Data_de_vencimento) Then
      If Data_de_vencimento < Data_da_Ordem Then
         If MsgBox("ATENÇÃO!! Data de Vencimento Menor que a Data de Emissão Deseja Prosseguir?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then
            ValidaVencto = False
         Else
            ValidaVencto = True: Exit Function
         End If
      Else
            ValidaVencto = True: Exit Function
      End If
   End If
 
End Function


Private Sub TrocaSeqReticom()
 Dim Doc As String
 
 Doc = "OS Nº " & troca_seq
 
 vgDb.Execute "UPDATE [Parcelas da Os] Set [Sequencia da Ordem] = " & troca_seq & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
 vgDb.Execute "UPDATE [Serviços da Os] Set [Sequencia da Ordem] = " & troca_seq & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
 vgDb.Execute "UPDATE [Virabrequim da Os] Set [Sequencia da Ordem] = " & troca_seq & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
 vgDb.Execute "UPDATE [Manutenção contas receber] Set [Sequencia da Ordem] = " & troca_seq & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
 'vgDb.Execute "UPDATE [Manutenção contas receber] Set Documento = '" & Doc & "' WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem                                                                
 vgDb.Execute "UPDATE [Ordem de Serviço] Set [Sequencia da Ordem] = " & troca_seq & " WHERE [Sequencia da Ordem] = " & Sequencia_da_Ordem
 
 'Load frmOServico
 frmOServico.Show
 frmOServico.SetFocus
 frmOServico.vgFiltroInicial = "[Sequencia da Ordem] = " & troca_seq
 InicializaFiltro frmOServico
 frmOServico.Reposition
 
End Sub


Private Function VerificaDebitos() As Boolean
   Dim Tb As New GRecordSet 
   
   On error Resume Next
     
   Set Tb = vgDb.OpenRecordSet("SELECT 1 FROM [Manutenção Contas Receber] WHERE [Sequencia do Geral] = " & Sequencia_do_Geral & " AND [Valor Restante] > 0 AND [Data de Vencimento] < " & D(Date))
        
   If Tb.RecordCount > 0 Then
      VerificaDebitos = True
   End If

End Function

Private Sub DadosAdic(Sequencia_da_Ordem As Long, Sequencia_do_Virabrequim As Long, Quantidade As Double, _
   Valor_unitario As Double, Tem_Casco As Boolean) 
 Dim Id As String 
  
 If Tem_Casco Then
    Dim Tb As New GRecordSet
    Set Tb = VgDb.OpenRecordSet("SELECT * FROM Virabrequim Where [Sequencia do Virabrequim] = " & Sequencia_do_Virabrequim)
    
    
    Id = "Cliente tem que retornar o casco - " & Tb!Identificação & " " & MarcaAux![Descrição] & " " & Tb!Modelo & " " & ModeloAux![Descrição Do Modelo]  
     
    With vgTb   
     .Edit
       If Vazio(Obs) Then
          !Obs = Id
       Else
          !Obs = !Obs & vbCrLf & ID
       End If   
     .Update
     .Bookmark = .lastModified
    End With
      
 End If
  
End Sub

Public Sub LigaDesligaBotoes
   Botao( 3 ).Enabled =    Botao( 3 ).Enabled And Permitido("Ordem de Serviço", ACAO_NAVEGANDO)
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Sequencia_da_Ordem = txtCampo(0).Value
         Data_da_Ordem = txtCampo(1).Value
         Sequencia_do_geral = txtCampo(2).Value
         Sequencia_do_Vendedor = txtCampo(3).Value
         Forma_de_pagamento = txtCampo(4).Value
         Valor_total_dos_produtos = txtCampo(5).Value
         Valor_Total_dos_Servicos = txtCampo(6).Value
         Valor_do_desconto = txtCampo(8).Value
         Valor_do_frete = txtCampo(9).Value
         Valor_da_OS = txtCampo(10).Value
         Aliquota_do_Iss = txtCampo(7).Value
         Valor_do_Iss = txtCampo(11).Value
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Obs = txtCampo(13).Value
         Usuario_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb![Usuario da Alteração])
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Sequencia_do_Transportador = txtCampo(14).Value
         Data_de_Cancelamento = txtCampo(12).Value
         Qual_Empresa = Val(labopcFrame0.Caption)
      Else
         Sequencia_da_Ordem = IIf(IsNull(vgTb![Sequencia da Ordem]), 0, vgTb![Sequencia da Ordem])
         Data_da_Ordem = vgTb![Data da Ordem]
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Sequencia_do_Vendedor = IIf(IsNull(vgTb![Sequencia do Vendedor]), 0, vgTb![Sequencia do Vendedor])
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Valor_total_dos_produtos = IIf(IsNull(vgTb![Valor total dos produtos]), 0, vgTb![Valor total dos produtos])
         Valor_Total_dos_Servicos = IIf(IsNull(vgTb![Valor Total dos Serviços]), 0, vgTb![Valor Total dos Serviços])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Valor_do_frete = IIf(IsNull(vgTb![Valor do frete]), 0, vgTb![Valor do frete])
         Valor_da_OS = IIf(IsNull(vgTb![Valor da OS]), 0, vgTb![Valor da OS])
         Aliquota_do_Iss = IIf(IsNull(vgTb![Aliquota do Iss]), 0, vgTb![Aliquota do Iss])
         Valor_do_Iss = IIf(IsNull(vgTb![Valor do Iss]), 0, vgTb![Valor do Iss])
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Obs = IIf(IsNull(vgTb!Obs), "", vgTb!Obs)
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Sequencia_do_Transportador = IIf(IsNull(vgTb![Sequencia do Transportador]), 0, vgTb![Sequencia do Transportador])
         Data_de_Cancelamento = vgTb![Data de Cancelamento]
         Qual_Empresa = IIf(IsNull(vgTb![Qual Empresa]), 0, vgTb![Qual Empresa])
      End If
      troca_seq = txtCampo(15).Value
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
   Dim i As Integer, vgRsError As GRecordSet, vgMsg As String, vgOk As Integer, vgPV As Boolean, vgNVez As Integer, vgInd As Integer
   On Error GoTo DeuErro                          'fica na espera de um erro...
   vgMsg$ = ""                                    'retorna uma msg dizendo o motivo
   vgOk = True                                    'se a validação esta OK
   vgPV = vgPriVez
   vgColumn = -1
   vgNVez = 0                                     'porque não fez o processo/validacoes
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      vgOk = ((Not Vazio(Data_da_Ordem)) And (IsDate(Data_da_Ordem) Or Vazio(Data_da_Ordem)))
      vgMsg$ = "Data da Ordem não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         If Sequencia_do_geral = 0 Then
            vgOk = (Sequencia_do_geral > 0)
            vgMsg$ = "O Campo Geral é Obrigatorio!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         If Sequencia_do_geral > 0 Then
            vgOk = (ValidaGeral())
            vgMsg$ = "Geral Inativo!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         If Sequencia_do_Vendedor > 0 Then
            vgOk = (ValidaVendedor())
            vgMsg$ = "Vendedor Inativo!"
            If Not vgOk Then vgColumn = 4
         End If
      End If
      If vgOk Then
         vgOk = (Sequencia_do_Transportador >= 0)
         vgMsg$ = "Sequencia do Transportador inválido!"
         If Not vgOk Then vgColumn = 15
      End If
      If vgOk Then
         vgOk = (Valor_do_desconto >= 0)
         vgMsg$ = "Valor do Desconto Inválido!"
         If Not vgOk Then vgColumn = 9
      End If
      If vgOk Then
         vgOk = (Valor_do_frete >= 0)
         vgMsg$ = "Valor do Frete inválido!"
         If Not vgOk Then vgColumn = 10
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
         txtCampo(7).Value = Empresa![Aliquota do Iss]
         txtCampo(12).Value = Date
         opcFrame0(0).Value = True
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 3
         PoeRelEFiltroCbo 14
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 3
         PoeRelEFiltroCbo 14
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         opcFrame0(Qual_Empresa).Value = True
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
      Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      InicializaApelidos COM_TEXTBOX
      vgTb![Qual Empresa] = Qual_Empresa
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
   Set Ordem_de_Servico = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   On Error Resume Next
   For i = 0 To 3
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
         Case 3
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(3).CloseRecordset
            Else
               x$ = ExecutaGrid(3, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(3).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(3).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(3).CloseRecordSet
                  End If
                  Grid(3).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(3, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionOS  
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
   Ok = (Sequencia_do_geral > 0 And Parcelas_da_os.RecordCount = 0)
   txtCampo(4).Enabled = (Ok And txtCampo(4).Editable)
   Ok = (False)
   Label(10).Enabled = Ok
   Ok = (False)
   txtCampo(7).Enabled = (Ok And txtCampo(7).Editable)
   Ok = (False)
   Label(11).Enabled = Ok
   Ok = (False)
   txtCampo(11).Enabled = (Ok And txtCampo(11).Editable)
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = True And vgSituacao = ACAO_NAVEGANDO)
   Botao(0).Enabled = Ok
   Ok = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0 And VgSituacao = ACAO_NAVEGANDO)
   Botao(1).Enabled = Ok
   Ok = ((Sequencia_da_Ordem > 0) AND Cancelado = 0 And vgSituacao = ACAO_NAVEGANDO)
   Botao(2).Enabled = Ok
   Ok = (Cancelado = 0 And VgSituacao = ACAO_NAVEGANDO)
   Botao(3).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(20).Visible = (Cancelado)
   txtCampo(12).Visible = (Cancelado)
   Label(22).Visible = (Sequencia_da_Ordem > 0)
   Botao(4).Visible = (False)
   txtCampo(15).Visible = (False)
   Label(24).Visible = (VerificaDebitos)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Or Index = 3 Or Index = 4 Or Index = 14 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
   Select Case Index
      Case 2
         ChangeGeral  
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
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 2
            PoeRelEFiltroCbo 2
         Case 3
            PoeRelEFiltroCbo 3
         Case 14
            PoeRelEFiltroCbo 14
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
         ExecutaVisivel
         ExecutaPreValidacao
         MostraFormulas
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
   If vgSituacao <> ACAO_NAVEGANDO Or (Len(txtCampo(Index).PesqSQLExpression) > 0) Then
      On Error Resume Next
      Select Case Index
         Case 2
            PoeRelEFiltroCbo 2
         Case 3
            PoeRelEFiltroCbo 3
         Case 14
            PoeRelEFiltroCbo 14
      End Select
   End If
End Sub

'evento - quando o campo perder o foco
Private Sub opcFrame0Cp_LostFocus(Index As Integer)
   opcFrame0(Index).LostFocus
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
   Caption = LoadGasString(26940)
   vgFormID = 680
   vgIdentTab = "Ordem de Serviço"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia da Ordem"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 26
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(18445))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(18446))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(18447))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(18448))
   Set Botao(3).Picture = LoadPicture(LoadGasPicture(18449))
   Set lblOrdem = Label(0)
   Set lblLinha = Label(1)
   Set lblAjuste = Label(5)
   Set Aba1 = Tabs(0)
   Set GrdVirabrequim = Grid(0)
   Set grdItens = Grid(1)
   Set grdServicos = Grid(2)
   Set lblParcelamento = Label(15)
   Set grdParcelamento = Grid(3)
   Set lblFechamento = Label(22)
   Set cmdAbrir = Botao(0)
   Set cmdFechar = Botao(1)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia da Ordem] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(26945)
   Label(1).Caption = LoadGasString(26946)
   Label(2).Caption = LoadGasString(26947)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(26948)
   Label(3).Caption = LoadGasString(26949)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(26950)
   Label(4).Caption = LoadGasString(26951)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(26952)
   Label(5).Caption = LoadGasString(26953)
   Label(6).Caption = LoadGasString(26954)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(26955)
   Tabs(0).TabCaption(0) = LoadGasString(26956)
   Tabs(0).TabCaption(1) = LoadGasString(26957)
   Tabs(0).TabCaption(2) = LoadGasString(26958)
   Tabs(0).TabCaption(3) = LoadGasString(26959)
   Label(7).Caption = LoadGasString(26960)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(26961)
   Label(8).Caption = LoadGasString(26962)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(26963)
   Label(9).Caption = LoadGasString(26964)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(26965)
   Label(10).Caption = LoadGasString(26966)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(26967)
   Label(11).Caption = LoadGasString(26968)
   Label(12).Caption = LoadGasString(26969)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(26970)
   Label(13).Caption = LoadGasString(26971)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(26972)
   Label(14).Caption = LoadGasString(26973)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(26974)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(26975)
   Label(16).Caption = LoadGasString(26976)
   Label(17).Caption = LoadGasString(26977)
   Label(18).Caption = LoadGasString(26978)
   Label(19).Caption = LoadGasString(26979)
   Label(20).Caption = LoadGasString(26980)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(26981)
   Label(21).Caption = LoadGasString(26982)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(26983)
   Label(22).Caption = LoadGasString(26984)
   vgTooltips.AddTool Botao(0), 0, LoadGasString(26985)
   Botao(0).Caption = LoadGasString(26986)
   vgTooltips.AddTool Botao(1), 0, LoadGasString(26987)
   Botao(1).Caption = LoadGasString(26988)
   Botao(2).Caption = LoadGasString(26989)
   Botao(3).Caption = LoadGasString(26990)
   Label(23).Caption = LoadGasString(26991)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(26992)
   Botao(4).Caption = LoadGasString(26993)
   Label(24).Caption = LoadGasString(26994)
   vgTooltips.AddTool Frame(0), 0, LoadGasString(26995)
   Frame(0).Caption = LoadGasString(26996)
   vgTooltips.AddTool opcFrame0(0).CtPri, 0, LoadGasString(26997)
   opcFrame0(0).Caption = LoadGasString(26998)
   vgTooltips.AddTool opcFrame0(1).CtPri, 0, LoadGasString(26999)
   opcFrame0(1).Caption = LoadGasString(27000)
   With Grid(0)
      .RowHeight = 330
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Virabrequim da OS", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Virabrequim da OS", ACAO_EDITANDO)
      .AllowDelete = Permitido("Virabrequim da OS", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Virabrequim", "Sequencia do Virabrequim", TP_CARACTER, "@!", , False, , "DADOSNFE", "Virabrequim", "Sequencia do Virabrequim", "Identificação", "Identificação", "Identificação", "", , "1", , , "DADOSNFE", "18", 1, "0", 1305
      .AddColumn Nothing, , "Marca", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1275
      .AddColumn Nothing, , "Modelo (A)", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1455
      .AddColumn Nothing, , "Modelo (B)", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1440
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1335
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Casco?", "Tem Casco", TP_LOGICO, , , False, , , , , , , , , , "0", , , , "0", 1, "0", 900
   End With
   With Grid(1)
      .RowHeight = 330
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Itens da OS", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Itens da OS", ACAO_EDITANDO)
      .AllowDelete = Permitido("Itens da OS", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Item", "Sequencia do produto", TP_CARACTER, "@x", , False, , "DADOSNFE", "Produtos", "Sequencia do produto", "Descrição", "Descrição", "Descrição", "", , "1", "Produtos.[Sequencia do produto]", "", "DADOSNFE", "18", 1, "0", 6780
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1440
   End With
   With Grid(2)
      .RowHeight = 330
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Serviços da OS", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Serviços da OS", ACAO_EDITANDO)
      .AllowDelete = Permitido("Serviços da OS", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Serviço", "Sequencia do Serviço", TP_CARACTER, "@x", , False, , "DADOSNFE", "Serviços", "Sequencia do Serviço", "Descrição do Serviço", "Descrição do Serviço", "Descrição do Serviço", "", , "1", "Serviços.[Sequencia do Serviço]", "", "DADOSNFE", "18", 2, "0", 6135
      .AddColumn Nothing, , "Nº Item", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 750
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1530
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1275
      .AddColumn Nothing, , "Virabrequim", "Sequencia do Virabrequim", TP_CARACTER, "@!", , False, , "DADOSNFE", "Virabrequim", "Sequencia do Virabrequim", "Identificação", "Identificação", "Identificação", "", , "1", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Marca", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1770
      .AddColumn Nothing, , "Modelo", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 2325
   End With
   With Grid(3)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Parcelas da os", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas da os", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas da os", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 780
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 660
      .AddColumn Nothing, , "Dt. Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1455
      .AddColumn Nothing, , "Vr. Parcela", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Especie de Pagto", "Especie de pagamento", TP_CARACTER, , , False, , , , , , "Dinheiro|Cartão|Cheque - Pre|Cheque - Vista|Boleto|Condicional|Garantia|Deposito", , , , "1", , , , "0", 1, "0", 2355
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Ordem_de_Servico = vgTb
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
   txtCampo(2).ListFields = "Razão social"
   txtCampo(2).OrderFields = "Razão social"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Geral"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True
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
   txtCampo(3).OrderFields = "Razão social-"
   txtCampo(3).Relation = ""
   txtCampo(3).Source = "Geral"
   txtCampo(3).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

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
   txtCampo(14).ListFields = "Razão social; Cpf e cnpj; Rg e ie"
   txtCampo(14).OrderFields = "Razão social"
   txtCampo(14).Relation = ""
   txtCampo(14).Source = "Geral"
   txtCampo(14).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

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
   txtCampo(5).Editable = False
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
   txtCampo(7).Mask = "99,99"
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
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "99.999.999,99"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = "999.999.999,99"
   txtCampo(10).Editable = False
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "9.999.999,99"
   txtCampo(11).Editable = False
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 4
   txtCampo(13).Mask = ""
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 1
   txtCampo(15).Mask = "999999"
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 2
   txtCampo(12).Mask = "99/99/9999"
   txtCampo(12).Editable = False
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set opcFrame0(0).CtPri = opcFrame0Cp(0)
   Set opcFrame0(0).CtFdo = labopcFrame0
   opcFrame0(0).DataType = 6
   opcFrame0(0).BookMark = 0
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(opcFrame0(0).DataField), opcFrame0(0)

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


'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 2
         txtCampo(2).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0", "[Sequencia do geral] > 0") & ""
                           txtCampo(2).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Geral.Nro, " + _
                                                              "Geral.Bairro, Geral.Cep, Geral.Fone, Geral.[Cpf e cnpj], Geral.[Rg e ie] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ")"
      Case 3
         txtCampo(3).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia do geral] > 0 AND Inativo = 0 And Vendedor = 1", "[Sequencia do geral] > 0 And Vendedor = 1") & ""
                           txtCampo(3).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social] FROM Geral WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND " + _
                                                              "(Geral.Inativo = " + CStr(0) + ") AND (Geral.Vendedor = " + CStr(1) + ")"
      Case 14
         txtCampo(14).Filter = "[Sequencia do geral] > " & 0 & ""
                           txtCampo(14).PesqSQLExpression = "SELECT Geral.[Sequencia do geral], Geral.[Razão social], Geral.[Nome fantasia], Geral.Endereço, Municipios.[Descrição do municipio] FROM Geral, " + _
                                                               "Municipios WHERE (Geral.[Sequencia do geral] > " + CStr(0) + ") AND (Geral.[Sequencia do municipio] = " + _
                                                               "Municipios.[Sequencia do municipio])"
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
   MudaTamCampos Me
End Sub

'evento - descarregando o formulário da memória
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   On Error Resume Next
   lblAlteracao.Caption = ""
   FinalizaForm Me
   Set lblOrdem = Nothing
   Set lblLinha = Nothing
   Set lblAjuste = Nothing
   Set Aba1 = Nothing
   Set GrdVirabrequim = Nothing
   Set grdItens = Nothing
   Set grdServicos = Nothing
   Set lblParcelamento = Nothing
   Set grdParcelamento = Nothing
   Set lblFechamento = Nothing
   Set cmdAbrir = Nothing
   Set cmdFechar = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   If Not Ordem_de_Servico Is Nothing Then
      Set Ordem_de_Servico = Nothing
   End If
   If Not Virabrequim_da_OS Is Nothing Then
      Virabrequim_da_OS.CloseRecordSet
      Set Virabrequim_da_OS = Nothing
   End If
   If Not Itens_da_OS Is Nothing Then
      Itens_da_OS.CloseRecordSet
      Set Itens_da_OS = Nothing
   End If
   If Not Servicos_da_OS Is Nothing Then
      Servicos_da_OS.CloseRecordSet
      Set Servicos_da_OS = Nothing
   End If
   If Not Parcelas_da_os Is Nothing Then
      Parcelas_da_os.CloseRecordSet
      Set Parcelas_da_os = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmOServico = Nothing                      'libera o segmento de código do form
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
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 1
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 2
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 3
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
      Case 3
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Virabrequim_da_OS![Sequencia da Ordem] = Ordem_de_Servico![Sequencia da Ordem]
      Case 1
         Itens_da_OS![Sequencia da Ordem] = Ordem_de_Servico![Sequencia da Ordem]
      Case 2
         Servicos_da_OS![Sequencia da Ordem] = Ordem_de_Servico![Sequencia da Ordem]
      Case 3
         Parcelas_da_os![Sequencia da Ordem] = Ordem_de_Servico![Sequencia da Ordem]
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
      Case 3
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
         Set Virabrequim_da_OS = vgNewRecordSet
      Case 1
         Set Itens_da_OS = vgNewRecordSet
      Case 2
         Set Servicos_da_OS = vgNewRecordSet
      Case 3
         Set Parcelas_da_os = vgNewRecordSet
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
      Case 1
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
      Case  3 
         ExecutaGrid = ExecutaGrid3(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_Virabrequim As Long, Quantidade As Double, Valor_unitario As Double
   Dim Tem_Casco As Boolean
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
         Case 2
            vgRetVal = (InfoVirabrequim(Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco, "Marca"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 3
            vgRetVal = (InfoVirabrequim(Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco, "ModeloA"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 4
            vgRetVal = (InfoVirabrequim(Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco, "Modelo"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      Comandos KeyCodeAscii, Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco 
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
      vgRetVal = "SELECT * FROM [Virabrequim da OS]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Ordem] = " & Ordem_de_Servico![Sequencia da Ordem]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 5
            Quantidade = 1
            vgRetVal = Quantidade
         Case 6
            Valor_unitario = InfoVirabrequim(Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco, "Valor")
            If Grid(0).Col = 6 Then
               vgRetVal = Valor_unitario
            End If
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "([Sequencia do Virabrequim] > " & 0 & ") AND [Sequencia do geral] = " & Ordem_de_Servico![Sequencia do geral] & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Sequencia_do_Virabrequim = 0 Then
            vgIsValid = (Sequencia_do_Virabrequim > 0)
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "Informe o Virabrequim!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Sequencia_do_Virabrequim > 0 Then
               vgIsValid = (TemOS(Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco) )
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "Impossivel - Já existe um OS com esse Virabrequim!"
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
         ColumnValue(5) = 1
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_EDITANDO Then
            DadosAdic Sequencia_da_Ordem, Sequencia_do_Virabrequim, Quantidade, Valor_unitario, Tem_Casco 
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Virabrequim = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Quantidade = ColumnValue(5)
   Valor_unitario = ColumnValue(6)
   Tem_Casco = ColumnValue(7)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Virabrequim_da_OS.Eof = False And Virabrequim_da_OS.BOF = False And Virabrequim_da_OS.RecordCount > 0 Then
      Sequencia_do_Virabrequim = Virabrequim_da_OS![Sequencia do Virabrequim]
      Quantidade = Virabrequim_da_OS!Quantidade
      Valor_unitario = Virabrequim_da_OS![Valor unitario]
      Tem_Casco = Virabrequim_da_OS![Tem Casco]
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


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_produto As Double, Quantidade As Double, Valor_unitario As Double
   Dim Valor_total As Double
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
      Select Case vgCol
         Case 4
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(1).AllowEdit = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(1).AllowDelete = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Itens da OS]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Ordem] = " & Ordem_de_Servico![Sequencia da Ordem]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "[Sequencia do produto] > " & 0 & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Produtos.[Sequencia do produto], Produtos.Descrição, Produtos.Estoque, Produtos.[Valor total] FROM Produtos WHERE (Produtos.[Sequencia do produto] > " + CStr(0) + ") AND " + _
                                                  "(Produtos.Inativo = " + CStr(0) + ")"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_do_produto > 0)
         If Not vgIsValid Then vgColumn =  1 
         vgErrorMessage$ = "Sequencia do Produto inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Quantidade inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_unitario > 0)
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Valor Unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
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
         Ordem_de_Servico.Edit
         Set vgRsError = Ordem_de_Servico
         If ProcessaItens(Sequencia_da_Ordem, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total) Then
            Ordem_de_Servico![Sequencia da Ordem] = (0)
         End If
         Ordem_de_Servico.Update
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
   Quantidade = ColumnValue(2)
   Valor_unitario = ColumnValue(3)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Itens_da_OS.Eof = False And Itens_da_OS.BOF = False And Itens_da_OS.RecordCount > 0 Then
         Valor_total = Itens_da_OS![Valor total]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Itens_da_OS.Eof = False And Itens_da_OS.BOF = False And Itens_da_OS.RecordCount > 0 Then
      Sequencia_do_produto = Itens_da_OS![Sequencia do produto]
      Quantidade = Itens_da_OS!Quantidade
      Valor_unitario = Itens_da_OS![Valor unitario]
      Valor_total = Itens_da_OS![Valor total]
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
   If Ordem_de_Servico.EOF = False And Ordem_de_Servico.BOF = False Then
      Ordem_de_Servico.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid1 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid2(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_Servico As Long, Quantidade As Double, Valor_unitario As Double
   Dim Valor_total As Double, Sequencia_do_Virabrequim As Long, Sequencia_do_item As Long
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
         Case 7
            vgRetVal = (InfoVirabrequim2(Sequencia_da_Ordem, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total, Sequencia_do_Virabrequim, _
   Sequencia_do_item, "Marca"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 8
            vgRetVal = (InfoVirabrequim2(Sequencia_da_Ordem, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total, Sequencia_do_Virabrequim, _
   Sequencia_do_item, "Modelo"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 6
            vgRetVal = Not (Virabrequim_da_OS.RecordCount > 0)
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(2).AllowInsert = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(2).AllowEdit = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      On Error Resume Next
         Grid(2).AllowDelete = (((Sequencia_da_nota_fiscal = 0) AND Cancelado = 0) AND Fechado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Serviços da OS]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Ordem] = " & Ordem_de_Servico![Sequencia da Ordem]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

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
            vgRetVal = "[Sequencia do Serviço] > " & 0 & ""
         Case 6
            vgRetVal = "" + Filtro() + ""
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
         vgIsValid = (Sequencia_do_Servico > 0)
         If Not vgIsValid Then vgColumn =  1 
         vgErrorMessage$ = "Sequencia do Serviço inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Quantidade inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_unitario > 0)
            If Not vgIsValid Then vgColumn =  4 
            vgErrorMessage$ = "Valor Unitario inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Virabrequim_da_OS.RecordCount > 0 Then
               vgIsValid = (Sequencia_do_Virabrequim > 0)
               If Not vgIsValid Then vgColumn =  6 
               vgErrorMessage$ = "Informe o Virabrequim!"
            End If
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(2) = SeqItem()
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
         Ordem_de_Servico.Edit
         Set vgRsError = Ordem_de_Servico
         If ProcessaServicos(Sequencia_da_Ordem, Sequencia_do_Servico, Quantidade, Valor_unitario, Valor_total, Sequencia_do_Virabrequim, _
   Sequencia_do_item) Then
            Ordem_de_Servico![Sequencia da Ordem] = (0)
         End If
         Ordem_de_Servico.Update
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
   Sequencia_do_Virabrequim = Val(Parse$(CStr(ColumnValue(6)), Chr$(1), 1))
   If Grid(2).Status <> ACAO_INCLUINDO Then
      If Servicos_da_OS.Eof = False And Servicos_da_OS.BOF = False And Servicos_da_OS.RecordCount > 0 Then
         Valor_total = Servicos_da_OS![Valor total]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Servicos_da_OS.Eof = False And Servicos_da_OS.BOF = False And Servicos_da_OS.RecordCount > 0 Then
      Sequencia_do_Servico = Servicos_da_OS![Sequencia do Serviço]
      Quantidade = Servicos_da_OS!Quantidade
      Valor_unitario = Servicos_da_OS![Valor unitario]
      Valor_total = Servicos_da_OS![Valor total]
      Sequencia_do_Virabrequim = Servicos_da_OS![Sequencia do Virabrequim]
      Sequencia_do_item = Servicos_da_OS![Sequencia do item]
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
   If Ordem_de_Servico.EOF = False And Ordem_de_Servico.BOF = False Then
      Ordem_de_Servico.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid2 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid3(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
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
      If Grid(3).Status <> ACAO_NAVEGANDO And vgItem = Grid(3).SelectedItem Then
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
            vgRetVal = Not (Not Vazio(Ordem_de_Servico![Forma de pagamento]))
         Case 3
            vgRetVal = Not (Not Vazio(Ordem_de_Servico![Forma de pagamento]))
         Case 4
            vgRetVal = Not (Not Vazio(Ordem_de_Servico![Forma de pagamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(3).AllowInsert = (((Ordem_de_Servico!Cancelado = 0) AND Not Vazio(Ordem_de_Servico![Forma de pagamento]) And TotalParcelas < Ordem_de_Servico![Valor da os]) AND Ordem_de_Servico!Fechado = 0)
      On Error Resume Next
         Grid(3).AllowEdit = ((((Not Vazio(Ordem_de_Servico![Forma de pagamento])) AND Ordem_de_Servico![Sequencia da nota fiscal] = 0) AND Ordem_de_Servico!Fechado = 0) AND Ordem_de_Servico!Cancelado = 0)
      On Error Resume Next
         Grid(3).AllowDelete = ((((Not Vazio(Ordem_de_Servico![Forma de pagamento])) AND Ordem_de_Servico![Sequencia da nota fiscal] = 0) AND Ordem_de_Servico!Fechado = 0) AND Ordem_de_Servico!Cancelado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas da os]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da Ordem] = " & Ordem_de_Servico![Sequencia da Ordem]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D", Ordem_de_Servico![Data da Ordem] , Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Ordem_de_Servico![Data da Ordem])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Ordem_de_Servico![Valor da OS] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Not Vazio(Ordem_de_Servico![Forma de pagamento]) Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Ordem_de_Servico![Forma de pagamento]) Then
               If Vazio(Data_de_vencimento) Then
                  vgIsValid = (Not Vazio(Data_de_vencimento))
                  If Not vgIsValid Then vgColumn =  3 
                  vgErrorMessage$ = "Data de Vencimento não pode ser em branco!"
               End If
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Ordem_de_Servico![Forma de pagamento]) Then
               If Not Vazio(Data_de_vencimento) Then
                  vgIsValid = (ValidaVencto(Sequencia_da_Ordem, Numero_da_parcela, Dias, Data_de_vencimento, Valor_da_parcela, Especie_de_pagamento))
                  If Not vgIsValid Then vgColumn =  3 
                  vgErrorMessage$ = "Vencimento Invalido!"
               End If
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Ordem_de_Servico![Forma de pagamento]) Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, IIf(Ordem_de_servico![Forma de pagamento] = "Prazo" , TotalParcelas() + Valor_da_Parcela <= Ordem_de_servico![Valor da Os], Valor_da_parcela = Ordem_de_servico![Valor da Os]), IIf(Ordem_de_servico![Forma de pagamento] = "Vista" , Valor_da_parcela = Ordem_de_servico![Valor da Os], Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Ordem_de_servico![Valor da Os])))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total do Movimento!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Not Vazio(Especie_de_pagamento))
            If Not vgIsValid Then vgColumn =  5 
            vgErrorMessage$ = "Especie de pagamento não pode ser vazio!"
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
   If Parcelas_da_os.Eof = False And Parcelas_da_os.BOF = False And Parcelas_da_os.RecordCount > 0 Then
      Numero_da_parcela = Parcelas_da_os![Numero da parcela]
      Dias = Parcelas_da_os!Dias
      Data_de_vencimento = Parcelas_da_os![Data de vencimento]
      Valor_da_parcela = Parcelas_da_os![Valor da parcela]
      Especie_de_pagamento = Parcelas_da_os![Especie de pagamento]
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
   ExecutaGrid3 = vgRetVal
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
      Case 3
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         AbreOS  
      Case 1
         Fechamento  
      Case 2
         CancelaOS  
      Case 3
         mdiNFE.MImprimOS
         AtivaForm Me
      Case 4
         TrocaSeqReticom  
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
Private Sub bottxtCampo3_Click(Index As Integer)
   txtCampo(3).SetFocus
   DoEvents
   txtCampo(3).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo14_Click(Index As Integer)
   txtCampo(14).SetFocus
   DoEvents
   txtCampo(14).botClick Index
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


