Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDevCompr
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8055
   Left                     =   1425
   LinkTopic                =   "DevCompr"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   7620
   ScaleWidth               =   13335
   Top                      =   1525
   Width                    =   13425
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
      Height                   =   7515
      Left                     =   15
      Width                    =   13275
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   23403
      _extenty                 =   13248
      Index                    =   0 
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   12840
         Width                    =   45
         Top                      =   7560
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
         Height                   =   1560
         Left                     =   120
         Width                    =   1365
         Top                      =   60
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
         Left                     =   1755
         Width                    =   3885
         Top                      =   555
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
         Caption                  =   "Devolução de Compra"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   1815
         Width                    =   3780
         Top                      =   795
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
         Caption                  =   "____________________________________"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   180
         Width                    =   1125
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
         Index                    =   3 
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
         Left                     =   1320
         Width                    =   825
         Top                      =   1680
         DataField                =   "Sequencia da devolução"
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
         Left                     =   2340
         Width                    =   1170
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
         Caption                  =   "Movimento:"
         Index                    =   4 
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
         TabIndex                 =   1
         TabStop                  =   0
         Height                   =   300
         Left                     =   3540
         Width                    =   840
         Top                      =   1680
         DataField                =   "Sequencia do movimento"
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
         Height                   =   240
         Left                     =   4560
         Width                    =   1125
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
         Caption                  =   "Dt. Compra:"
         Index                    =   5 
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
         TabIndex                 =   2
         TabStop                  =   0
         Height                   =   300
         Left                     =   5760
         Width                    =   1365
         Top                      =   1680
         DataField                =   "Data da entrada"
         TabStop                  =  0
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7320
         Width                    =   1425
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
         Caption                  =   "Dt. Devolução:"
         Index                    =   6 
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
         TabIndex                 =   3
         TabStop                  =   0
         Height                   =   300
         Left                     =   8820
         Width                    =   1365
         Top                      =   1680
         DataField                =   "Data da devolução"
         TabStop                  =  0
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   10380
         Width                    =   825
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
         Caption                  =   "Usuario:"
         Index                    =   7 
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
         MaxLength                =   30
         BorderStyle              =   1
         Alignment                =   0
         Locked                   =   -1
         TabIndex                 =   4
         TabStop                  =   0
         Height                   =   300
         Left                     =   11280
         Width                    =   1665
         Top                      =   1680
         DataField                =   "Usuario da Alteração"
         TabStop                  =  0
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   660
         Width                    =   645
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
         Index                    =   8 
      End
      Begin VB.Label labFdo5
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &HE0E0E0
         Left                     =   1320
         Top                      =   2340
         Width                    =   5820
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo5
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   6870
         Top                      =   2355
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
         Locked                   =   -1
         TabIndex                 =   5
         TabStop                  =   0
         Height                   =   270
         Left                     =   1335
         Width                    =   5520
         Top                      =   2355
         DataField                =   "Sequencia do geral"
         TabStop                  =  0
         Index                    =   5 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   2
         TabHeight                =   519
         BackColor                =   &HE0E0E0
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   2
         TabIndex                 =   6
         TabStop                  =   0
         Height                   =   3225
         Left                     =   180
         Width                    =   12780
         Top                      =   3060
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
         _extentx                 =   22531
         _extenty                 =   5685
         _stockprops              =   15
         TabCaption(0)            =   "1 - Produtos"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "2 - Financeiro"
         Tab(1).ControlCount=   25
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Label(9)"
         Tab(1).Control(1)=   "Label(10)"
         Tab(1).Control(2)=   "Formato(0)"
         Tab(1).Control(3)=   "Label(11)"
         Tab(1).Control(4)=   "Formato(1)"
         Tab(1).Control(5)=   "Label(12)"
         Tab(1).Control(6)=   "Formato(2)"
         Tab(1).Control(7)=   "Label(13)"
         Tab(1).Control(8)=   "Label(14)"
         Tab(1).Control(9)=   "txtCp(6)"
         Tab(1).Control(10)=   "Label(15)"
         Tab(1).Control(11)=   "txtCp(7)"
         Tab(1).Control(12)=   "Label(16)"
         Tab(1).Control(13)=   "txtCp(8)"
         Tab(1).Control(14)=   "Grid(1)"
         Tab(1).Control(15)=   "Label(17)"
         Tab(1).Control(16)=   "txtCp(9)"
         Tab(1).Control(17)=   "Label(18)"
         Tab(1).Control(18)=   "labFdo10"
         Tab(1).Control(19)=   "bottxtCampo10(2)"
         Tab(1).Control(20)=   "txtCp(10)"
         Tab(1).Control(21)=   "Label(19)"
         Tab(1).Control(22)=   "txtCp(11)"
         Tab(1).Control(23)=   "Label(20)"
         Tab(1).Control(24)=   "txtCp(12)"
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
            Height                   =   2595
            Left                     =   120
            Width                    =   12525
            Top                      =   420
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   0
            BackColor                =   &H0
            BackStyle                =   1
            BorderStyle              =   1
            AutoSize                 =   0
            Height                   =   255
            Left                     =   -67545
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
            Caption                  =   "Parcelamento"
            Index                    =   9 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -66420
            Width                    =   600
            Top                      =   2895
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
            Left                     =   -65745
            Width                    =   165
            Top                      =   2940
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65520
            Width                    =   990
            Top                      =   2895
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
            Left                     =   -64425
            Width                    =   165
            Top                      =   2940
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -64125
            Width                    =   1050
            Top                      =   2895
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
            Left                     =   -62910
            Width                    =   150
            Top                      =   2940
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -62685
            Width                    =   300
            Top                      =   2895
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
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -71625
            Width                    =   2040
            Top                      =   495
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
            Caption                  =   "Total dos Produtos:"
            Index                    =   14 
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   10
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69420
            Width                    =   1470
            Top                      =   495
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
            Left                     =   -70845
            Width                    =   1260
            Top                      =   1035
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
            Caption                  =   "Total de IPI:"
            Index                    =   15 
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   11
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69420
            Width                    =   1470
            Top                      =   1035
            DataField                =   "Valor total do ipi"
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
            Left                     =   -71700
            Width                    =   2115
            Top                      =   2055
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
            Caption                  =   "Total da Devolução:"
            Index                    =   16 
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   12
            TabStop                  =   0
            Height                   =   300
            Left                     =   -69420
            Width                    =   1500
            Top                      =   2055
            DataField                =   "Valor total da devolução"
            TabStop                  =  0
            Index                    =   8 
            MultiLine                =   -1
         End
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
            TabIndex                 =   9
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2190
            Left                     =   -67545
            Width                    =   5175
            Top                      =   690
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -71145
            Width                    =   1560
            Top                      =   1560
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
            Caption                  =   "Total ICMS ST:"
            Index                    =   17 
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
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   13
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -69420
            Width                    =   1470
            Top                      =   1560
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
            Left                     =   -71310
            Width                    =   1725
            Top                      =   2580
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
            Caption                  =   "Forma de Pagto:"
            Index                    =   18 
         End
         Begin VB.Label labFdo10
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -69420
            Top                      =   2580
            Width                    =   1500
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo10
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -68190
            Top                      =   2595
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
            Alignment                =   1
            TabIndex                 =   14
            TabStop                  =   0
            Height                   =   270
            Left                     =   -69405
            Width                    =   1215
            Top                      =   2595
            DataField                =   "Forma de pagamento"
            Index                    =   10 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -74040
            Width                    =   600
            Top                      =   495
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
            Caption                  =   "Frete:"
            Index                    =   19 
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
            Alignment                =   1
            TabIndex                 =   17
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -73260
            Width                    =   1380
            Top                      =   495
            DataField                =   "Valor do frete"
            Index                    =   11 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -74490
            Width                    =   1050
            Top                      =   1035
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
            Caption                  =   "Desconto:"
            Index                    =   20 
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
            Alignment                =   1
            TabIndex                 =   18
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -73260
            Width                    =   1380
            Top                      =   1035
            DataField                =   "Valor do desconto"
            Index                    =   12 
            MultiLine                =   -1
         End
      End
      Begin NFE.GPainel Painel
         BackColor                =   &HE0E0E0
         BevelOuter               =   2
         BevelInner               =   1
         BorderWidth              =   1
         BevelWidth               =   1
         BorderStyle              =   0
         Stretch                  =   -1
         Stripes                  =   0
         SaveGridBars             =   0
         Height                   =   735
         Left                     =   210
         Width                    =   12780
         Top                      =   6495
         TabStop                  =  0
         _extentx                 =   22531
         _extenty                 =   1295
         Index                    =   1 
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   435
            Left                     =   135
            Width                    =   2430
            Top                      =   135
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
            Caption                  =   "Status: Aberta"
            Index                    =   21 
         End
         Begin NFE.GBotao Botao
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   16
            TabStop                  =   0
            ButtonType               =   0
            Height                   =   480
            Left                     =   3375
            Width                    =   1350
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
            Caption                  =   "Abrir"
            Index                    =   0 
         End
      End
      Begin NFE.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   8
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   10305
         Width                    =   1320
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
         TabIndex                 =   15
         TabStop                  =   -1
         ButtonType               =   0
         Height                   =   480
         Left                     =   11655
         Width                    =   1320
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
Attribute VB_name = "frmDevCompr"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: DevCompr
'* Função....: Devolução de Compra
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:44
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
Dim txtCampo(12) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim vgPodeFazerUnLoad As Boolean                  'flag se o módulo pode ou nao ser removido da memória
Dim Valor_do_frete As Double, Valor_do_desconto As Double, Sequencia_da_devolucao As Long
Dim Sequencia_do_movimento As Long, Sequencia_do_geral As Long, Data_da_entrada As Variant
Dim Data_da_devolucao As Variant, Historico As String, Cancelado As Boolean
Dim Valor_total_dos_produtos As Double, Valor_total_do_ipi As Double, Forma_de_pagamento As String
Dim Valor_total_da_devolucao As Double, Sequencia_da_nota_fiscal As Long, Data_da_Alteracao As Variant
Dim Hora_da_Alteracao As Variant, Usuario_da_Alteracao As String, Fechado As Boolean
Dim Valor_total_icms_st As Double
Public lblAjuste As Object, lblDevolucao As Object, lblLinha As Object
Public GrdProdutos As Object, lblParcelamento As Object, grdParcelamento As Object
Public lblStatus As Object, cmdAbrir As Object, cmdFechar As Object
Dim Devolucao_de_compra As New GRecordSet, Produtos_devolucao_compra As New GRecordSet, Parcelas_dev_compra As New GRecordSet

Public UnidadeAux As New GRecordSet, ProdutoAux As New GRecordSet
Public ItensAux As New GRecordSet

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub

Public Sub AjustaValores()
   Dim IPIProdutos As New GRecordSet, ValorProdutos As New GRecordSet
   Dim ValorICMSST As New GRecordSet, ValorMvto As Currency
   Dim Produtos As New GRecordSet
 
   On Error GoTo DeuErro
   
   Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)
            
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT ROUND(SUM([Valor do ipi]), 2) Total FROM [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao) 'IPI dos Produtos
   Set ValorProdutos = vgDb.OpenRecordSet("SELECT SUM([Valor Total]) Total FROM [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao) 
   Set ValorICMSST = vgDb.OpenRecordSet("SELECT SUM([Valor icms st]) Total FROM [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao) 'Substituição Produtos                                                                                             
   ValorMvto = IPIProdutos!Total + ValorProdutos!Total + ValorICMSST!Total + Valor_do_frete - Valor_do_desconto 
   ValorMvto = Format(ValorMvto, "##,###,##0.00")
   
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Devolução de compra] Set [Valor total do ipi] = " & Substitui(IPIProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao 'IPI Produtos
   vgDb.Execute "Update [Devolução de compra] Set [Valor total dos produtos] = " & Substitui(ValorProdutos!Total, ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao 'Produtos 
   vgDb.Execute "Update [Devolução de compra] Set [Valor total icms st] = " & Substitui(ValorICMSST!Total, ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao 'Substituição
   vgDb.Execute "Update [Devolução de compra] Set [Valor total da devolução] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao 'Valor da Entrada
   vgDb.CommitTrans
     
  ' Alteracao
   
   If Produtos.RecordCount > 0 Then
      Do While Not Produtos.EOF
      UltimoMvto(Produtos![Sequencia Do produto])
      UltimoFornecedor(Produtos![Sequencia Do produto])
      AtualizaEstoque(Produtos![Sequencia Do produto])
      Produtos.MoveNext
      Loop
      Produtos.MoveFirst  
   End If
   
   GrdProdutos.Rebind
   
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If

End Sub


Private Function InfoProdutos(Sequencia_da_devolucao As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, _
   Sequencia_do_movimento As Long, Oq As String) As Variant
   
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
      Case "Descrição"
         InfoProdutos = ProdutoAux![Descrição]   
   End Select

End Function


Private Sub AjustaProdutos(Sequencia_da_devolucao As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, _
   Sequencia_do_movimento As Long) 
 Dim ProdutosMvto As New GRecordSet

  On Error GoTo DeuErro
   
   vgDb.BeginTrans      
   vgDb.Execute "Update [Produtos devolução compra] Set [Valor total] = " & Substitui(CCur(Round(Quantidade * Valor_unitario, 2)), ",", ".", UM_A_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND [Sequencia do item] = " & Sequencia_do_item
   If Aliquota_do_ipi > 0 Then
      vgDb.Execute "Update [Produtos devolução compra] Set [Valor do ipi] = " & Substitui(Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2), ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND [Sequencia do item] = " & Sequencia_do_item 'Valor do IPI
   Else
      vgDb.Execute "Update [Produtos devolução compra] Set [Valor do ipi] = 0 WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND [Sequencia do item] = " & Sequencia_do_item
   End If
   
  
   Set ProdutosMvto = VgDb.OpenRecordSet("SELECT * From [Produtos movimento estoque] Where [Sequencia do movimento] = " & Sequencia_do_movimento & " And [Sequencia do item] = " & Sequencia_do_item & " And [Sequencia do produto] = " & Sequencia_do_produto)
   If ProdutosMvto![Valor icms st] > 0 And Quantidade > 0 Then   
      vgDb.Execute "Update [Produtos devolução compra] Set [Valor icms st] = " & Substitui(Round((ProdutosMvto![Valor icms st] / ProdutosMvto!Quantidade) * Quantidade, 2), ",", ".", SO_UM) & " WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND [Sequencia do item] = " & Sequencia_do_item 'Valor do IPI
   Else
      vgDb.Execute "Update [Produtos devolução compra] Set [Valor icms st] = 0 WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND [Sequencia do item] = " & Sequencia_do_item
   End If
   
   AjustaValores
   
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
   frmDevCompr.vgFiltroInicial = "[Sequencia da Devolução] = " & Sequencia_da_devolucao
   InicializaFiltro frmDevCompr
   frmDevCompr.AjustaValores
   frmDevCompr.Reposition

End Sub


Private Sub CarregaTotalizador()
   On Error Resume Next

   With grdProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor Unitario").Index) = True
      .ShowSumCol(.Columns("Quantidade").Index) = True
      .ShowSumCol(.Columns("Valor ICMS ST").Index) = True
      .ShowSumCol(.Columns("Valor do IPI").Index) = True
      .ShowSumCol(.Columns("% IPI").Index) = False
      .ShowSumCol(.Columns("Sequencia do Item").Index) = False
     ' .ColumnWidth(CInt(.Columns("Sequencia do item").Index)) = 0
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


Private Sub DeletaProdutos(Sequencia_da_devolucao As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, _
   Sequencia_do_movimento As Long) 

  Dim Tb As New GRecordSet

  If vgSITUACAO = -ACAO_EXCLUINDO Then
  Set Tb = vgDb.OpenRecordSet("SELECT * From [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)
   If Tb.RecordCount > 0 Then
      vgDb.Execute "DELETE FROM [Produtos devolução compra] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & " And [Sequencia do item] = " & Sequencia_do_item
   End If
  End If 
  
  AtualizaEstoque(Sequencia_do_produto)
  AjustaValores

End Sub


Private Function InfoMovimento(Sequencia_da_devolucao As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, _
   Sequencia_do_movimento As Long, Oq As String) As Variant
 Dim QtdDevolucao As New GRecordSet
   
   On Error Resume Next
   
   Set QtdDevolucao = vgDb.OpenRecordSet("SELECT SUM([Quantidade]) Total From [Produtos devolução compra] LEFT JOIN [Devolução de compra] ON [Produtos devolução compra].[Sequencia da devolução] = [Devolução de compra].[Sequencia da devolução] Where [Sequencia do Produto] = " & Sequencia_do_produto & " And [Devolução de compra].Cancelado = 0 And [Produtos devolução compra].[Sequencia do movimento] = " & Sequencia_do_movimento)
   TbAuxiliar "Produtos movimento estoque", "[Sequencia do produto] = " & Sequencia_do_produto & " And [Sequencia do item] = " & Sequencia_do_item & " And [Sequencia do movimento] = " & Sequencia_do_movimento, ItensAux
      
   If ItensAux.RecordCount = 0 Then Exit Function
   
   Select Case Oq      
      Case "Qtde"
         InfoMovimento = ItensAux![Quantidade] - QtdDevolucao!Total   
   End Select

End Function

Private Function ValidaQuantidade(Sequencia_da_devolucao As Long, Sequencia_do_item As Long, Sequencia_do_produto As Double, _
   Quantidade As Double, Valor_unitario As Double, Valor_total As Double, _
   Valor_do_ipi As Double, Valor_icms_st As Double, Aliquota_do_ipi As Single, _
   Sequencia_do_movimento As Long) As Boolean
 Dim QtdMvto As New GRecordSet, QtdDevolucao As New GRecordSet
 Dim QtdReg As New GRecordSet
 
  Set QtdDevolucao = vgDb.OpenRecordSet("SELECT SUM([Quantidade]) Total From [Produtos devolução compra] LEFT JOIN [Devolução de compra] ON [Produtos devolução compra].[Sequencia da devolução] = [Devolução de compra].[Sequencia da devolução] Where [Sequencia do Produto] = " & Sequencia_do_produto & " And [Devolução de compra].Cancelado = 0 And [Produtos devolução compra].[Sequencia do movimento] = " & Sequencia_do_movimento)
  Set QtdMvto = vgDb.OpenRecordSet("SELECT SUM([Quantidade]) Total FROM [Produtos movimento estoque] WHERE [Sequencia do movimento] = " & Sequencia_do_movimento & " And [Sequencia do produto] = " & Sequencia_do_produto)
  Set QtdReg = VgDb.OpenRecordSet("Select * From [Devolução de compra] Where [Sequencia do movimento] = " & Sequencia_do_movimento & " And [Devolução de compra].Cancelado = 0")
  
   If QtdReg.RecordCount = 1 Then
      ValidaQuantidade = (QtdMvto!Total - Produtos_devolucao_compra!Quantidade + QtdDevolucao!Total - Quantidade) <= QtdMvto!Total
      If (QtdMvto!Total - Produtos_devolucao_compra!Quantidade + QtdDevolucao!Total - Quantidade) < 0 Then
         ValidaQuantidade = False
      End If
   End If  
 
   If QtdReg.RecordCount > 1 Then
      ValidaQuantidade = (QtdMvto!Total + Produtos_devolucao_compra!Quantidade - Quantidade - QtdDevolucao!Total) <= QtdMvto!Total
      If (QtdMvto!Total + Produtos_devolucao_compra!Quantidade - Quantidade - QtdDevolucao!Total) < 0 Then
         ValidaQuantidade = False
     End If
   End If
    
End Function

Private Sub MostraFormulas()

   If Cancelado Then
      lblDevolucao.ForeColor = &H8FF
      lblLinha.ForeColor = &H8FF
      lblDevolucao.Caption = "Devolução de Compra"
      lblStatus.Caption = "Status: Cancelada" 
      lblStatus.ForeColor = &H8FF
   Else
       lblDevolucao.ForeColor = &H0
       lblLinha.ForeColor = &H0
       lblDevolucao.Caption = "Devolução de Compra"
       If Fechado Then
          lblStatus.Caption = "Status: Fechada"
          lblStatus.ForeColor = &H0  
       Else
          lblStatus.Caption = "Status: Aberta"
          lblStatus.ForeColor = &H0    
       End If 
   End If
   If Err Then Err = 0: lblDevolucao.Caption = "" 
   If Err Then Err = 0: lblStatus.Caption = "" 
   
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Valor_total_da_devolucao) Or (TotalParcelas > Valor_total_da_devolucao) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
      
End Sub

Private Sub Fechamento()
 Dim Produtos As New GRecordSet, Parcelas As New GRecordSet
 Dim Manutencao As New GRecordSet, Sek As Long
 
 On Error Resume Next
 
 Set Produtos = VgDb.OpenRecordSet("SELECT * From [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)
 Set Parcelas = VgDb.OpenRecordSet("SELECT * From [Parcelas dev compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)
 Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
  
 If Produtos.RecordCount = 0 Then
    MsgBox("Impossivel Fechar a Devolução!! Devolução sem Itens")
    Exit Sub
 End If
 
  If Parcelas.RecordCount > 0 Then
   If TotalParcelas() <> Valor_total_da_devolucao Then
      MsgBox("Impossivel Fechar a Devolução!! A Soma das Parcelas não Correspondem ao Valor Total da Devolução")
   Exit Sub
   End If
  End If
  
  If Parcelas.RecordCount > 0 And Manutencao.RecordCount = 0 Then
     LancaParcelas 
  Else
     DeletaConta
     LancaParcelas     
  End If
  
  If Parcelas.RecordCount = 0 Then If MsgBox("Deseja Fechar a Devolução sem Incluir a(s) Parcelas do Contas a Receber?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then Exit Sub
      vgDb.Execute "UPDATE [Devolução de compra] Set Fechado = 1 WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao  
  Sek = Sequencia_da_devolucao
  InputYgorsBom = ""
  frmDevCompr.vgFiltroInicial = "[Sequencia da devolução] = " & Sek
  InicializaFiltro frmDevCompr
  frmDevCompr.Reposition
              
End Sub


Private Sub Reabrir()
 Dim Baixa As New GRecordSet, TemBaixa As Boolean
 Dim Manu As New GRecordSet, Sek As Long

 On Error Resume Next
 
  Set Baixa = vgDb.OpenRecordSet("SELECT * FROM [Baixa contas receber] WHERE [Sequencia da manutenção] IN (SELECT [Sequencia da manutenção] FROM [Manutenção contas receber] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao & ")")
  Set Manu = vgDb.OpenRecordSet("SELECT * From [Manutenção contas receber] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)
 
    If Baixa.RecordCount > 0 Then
       TemBaixa = True
    Else
       TemBaixa = False
    End If
   
    If TemBaixa Then
       MsgBox("Impossivel abrir o Movimento existe Baixa(s) no Contas a Receber para Abri-lo Exclua a(s) Baixa(s)")
       Exit Sub
    End If
     
InputSenha
  If Not Vazio(InputYgorsBom) Then
    If InputYgorsBom = Empresa![Senha Admin] Then
       vgDb.Execute "UPDATE [Devolução de compra] Set Fechado = 0 Where [Sequencia da devolução] = " & Sequencia_da_devolucao 
       If Manu.RecordCount > 0 Then
          vgDb.Execute "DELETE FROM [Manutenção contas receber] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao
       End If          
       Reposition True 
    Else
       MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
       Exit Sub
    End If
  End If
  Sek = Sequencia_da_devolucao
  InputYgorsBom = ""
  frmDevCompr.vgFiltroInicial = "[Sequencia da devolução] = " & Sek
  InicializaFiltro frmDevCompr
  frmDevCompr.Reposition
  
End Sub  


Private Sub PreValidaFechamento()
   Dim Itens As New GRecordSet
   Dim Vector() As String
   Dim i As Long, vaPrivez As Boolean, Campo As Variant, Mensagem As String
 
   On Error GoTo DeuErro
                                  
   Set Itens = vgDb.OpenRecordSet("Select * From [Produtos devolução compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao)                                
                                                              
   
   'Vamos Validar
   i = 0 'Tamanho do Vetor
   ReDim Preserve Vector(0) As String
   
   Do While Not Itens.EOF
      If Itens!Quantidade > 0 Then i = i + 1: ReDim Preserve Vector(i): Vector(i - 1) = "ITEM: " & Itens![Sequencia Do produto] & " Quantidade = 0"
      Itens.MoveNext
   Loop
      
   If UBound(Vector) = 0 Then
      Mensagem = "Pelo Menos 1 dos Produtos a Quantidade deve ser maior que 0:" & vbCrLf
      For Each Campo In Vector
         Mensagem = Mensagem & vbCrLf & Campo
      Next
      If Mensagem <> "" Then
         MsgBox Mensagem, vbCritical + vbOKOnly, vaTitulo
         Exit Sub
      End If
   End If
   
   Fechamento
     
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Sub


Private Function UltimaParcela() As Long
   Dim Tb As New GRecordset

   On Error Resume Next
   
   Set Tb = vgdb.OpenRecordset("SELECT MAX([Numero da parcela]) PC FROM [Parcelas dev compra] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
   
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If

End Function


Private Function UltimoVencimento() As Variant
 Dim Tb As New GRecordSet, RetVal As Variant
   
  On Error Resume Next
      
  Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de vencimento]) [Data de Vencimento] FROM [Parcelas dev compra] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
      
  RetVal = grdParcelamento.ColumnValue(grdParcelamento.Row + 1, grdParcelamento.Col - 2)
  
   If vgSituacao = -ACAO_INCLUINDO Then    
      If Not IsNull(Tb![Data de Vencimento]) Then
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Devolucao_de_compra![Data da devolução]
      End If
  End If
  
  If vgSituacao = -ACAO_EDITANDO Then    
      If Not IsNull(Tb![Data de Vencimento]) Then
         UltimoVencimento = RetVal >= Devolucao_de_compra![Data da devolução]
      Else
         UltimoVencimento = RetVal >= Devolucao_de_compra![Data da devolução]
      End If
  End If
  
End Function

Private Sub RepositionDevolucao()
   On Error Resume Next
   
   If vgSituacao = ACAO_INCLUINDO Then
      Set Produtos_devolucao_compra = Nothing
      Set Parcelas_dev_compra = Nothing
   End If
   
   CarregaTotalizador
   grdProdutos.Rebind
   
End Sub


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da parcela]) As Total From [Parcelas dev compra] Where [Sequencia da devolução] = " & Sequencia_da_devolucao & " AND " & IIf(Seq > 0, "[Numero da parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


Private Sub DeletaConta() 'Qdo exclui registro Mãe
   Dim Manutencao As New GRecordset

   On Error GoTo DeuErro
   
   Set Manutencao = vgdb.OpenRecordset("SELECT * FROM [Manutenção contas receber] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
    
   If Manutencao.RecordCount = 0 Then Exit Sub
   
   vgDb.Execute "DELETE FROM [Manutenção contas receber] " & _
                "WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao

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
   
   If TotalParcelas() <> Valor_total_da_devolucao Then Exit Sub
   
   Set Manutencao = vgDb.OpenRecordSet("Manutenção contas receber")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas dev compra] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
   
   vgDb.BeginTrans
   
   'Manutenção
   Do While Not Parcelas.EOF
      TbAuxiliar "Manutenção contas receber", "[Sequencia da devolução] = " & Sequencia_da_devolucao & " AND Parcela = " & Parcelas![Numero da parcela], Manutencao
      With Manutencao
         If Manutencao.RecordCount = 0 Then .AddNew
         !Parcela = Parcelas![Numero da parcela]
         ![Sequencia Do geral] = Sequencia_do_Geral
         !Documento = "Dev. Compra"
         ![Data de entrada] = Data_da_devolucao
         !Historico = "Devolução de Compra Ref. Mvto de Entrada: " + Format(Sequencia_do_movimento,"000000") 
         ![Forma de Pagamento] = Forma_de_pagamento
         ![Especie de pagamento] = Parcelas![Especie de pagamento]
          If ![Forma de pagamento] = "Vista" And Empresa![Baixa Receber Auto] Then
             ![Valor pago] = Parcelas![Valor da parcela]
             ![Valor restante] = 0
             ![Data da baixa] = Data_da_devolucao
          Else
         ![Valor pago] = 0
         ![Valor restante] = Parcelas![Valor da parcela]
         ![Data da baixa] = Null
          End If
         ![Data de vencimento] = Parcelas![Data de vencimento]
         ![Valor total] = Valor_total_da_devolucao
         ![Valor da parcela] = Parcelas![Valor da parcela]
         ![Sequencia da devolução] = Sequencia_da_devolucao
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


Private Sub AbreGerar()
   Dim Geral As New GRecordSet
   Dim DtaHoje As Date, DtaParcela As Date, Parcelas As New GRecordSet
   
   On Error Resume Next
   
   DtaHoje = Date 'Data de Hoje
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas dev compra] WHERE [Sequencia da devolução] = " & Sequencia_da_devolucao)
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
   If TotalParcelas > 0 And TotalParcelas <> Valor_total_da_devolucao Then
      MsgBox "Parcelamento Incompleto. Impossível Gerar a NFe.", vbExclamation + vbOKOnly, vaTitulo
      Exit Sub
   End If
      
   Unload Me
   Load frmF_GeraNF ' Abre Form p/ Gerar NFe 
   frmF_GeraNF.RepositionDevolucao
   frmF_GeraNF.Executar INICIALIZACOES
   frmF_GeraNF.Executar INI_APELIDOS
   frmF_GeraNF.Show
   SendK vbKeyReturn
   SendK vbKeyReturn
   SendK vbKeyReturn 
    
     
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Valor_do_frete = txtCampo(11).Value
         Valor_do_desconto = txtCampo(12).Value
         Sequencia_da_devolucao = txtCampo(0).Value
         Sequencia_do_movimento = txtCampo(1).Value
         Sequencia_do_geral = txtCampo(5).Value
         Data_da_entrada = txtCampo(2).Value
         Data_da_devolucao = txtCampo(3).Value
         Historico = Iif(vgSituacao = ACAO_INCLUINDO,"", vgTb!Historico)
         Cancelado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Cancelado)
         Valor_total_dos_produtos = txtCampo(6).Value
         Valor_total_do_ipi = txtCampo(7).Value
         Forma_de_pagamento = txtCampo(10).Value
         Valor_total_da_devolucao = txtCampo(8).Value
         Sequencia_da_nota_fiscal = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Sequencia da nota fiscal])
         Data_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Data da Alteração])
         Hora_da_Alteracao = Iif(vgSituacao = ACAO_INCLUINDO,Null, vgTb![Hora da Alteração])
         Usuario_da_Alteracao = txtCampo(4).Value
         Fechado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Fechado)
         Valor_total_icms_st = txtCampo(9).Value
      Else
         Valor_do_frete = IIf(IsNull(vgTb![Valor do frete]), 0, vgTb![Valor do frete])
         Valor_do_desconto = IIf(IsNull(vgTb![Valor do desconto]), 0, vgTb![Valor do desconto])
         Sequencia_da_devolucao = IIf(IsNull(vgTb![Sequencia da devolução]), 0, vgTb![Sequencia da devolução])
         Sequencia_do_movimento = IIf(IsNull(vgTb![Sequencia do movimento]), 0, vgTb![Sequencia do movimento])
         Sequencia_do_geral = IIf(IsNull(vgTb![Sequencia do geral]), 0, vgTb![Sequencia do geral])
         Data_da_entrada = vgTb![Data da entrada]
         Data_da_devolucao = vgTb![Data da devolução]
         Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Valor_total_dos_produtos = IIf(IsNull(vgTb![Valor total dos produtos]), 0, vgTb![Valor total dos produtos])
         Valor_total_do_ipi = IIf(IsNull(vgTb![Valor total do ipi]), 0, vgTb![Valor total do ipi])
         Forma_de_pagamento = IIf(IsNull(vgTb![Forma de pagamento]), "", vgTb![Forma de pagamento])
         Valor_total_da_devolucao = IIf(IsNull(vgTb![Valor total da devolução]), 0, vgTb![Valor total da devolução])
         Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
         Data_da_Alteracao = vgTb![Data da Alteração]
         Hora_da_Alteracao = vgTb![Hora da Alteração]
         Usuario_da_Alteracao = IIf(IsNull(vgTb![Usuario da Alteração]), "", vgTb![Usuario da Alteração])
         Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
         Valor_total_icms_st = IIf(IsNull(vgTb![Valor total icms st]), 0, vgTb![Valor total icms st])
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
      i = (False) and Permitido(vgIdentTab, ACAO_INCLUINDO)
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
      i = Permitido(vgIdentTab, ACAO_EDITANDO)
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
   If vgOq = INICIALIZACOES Then
      If vgPriVez = False Then
         vgPriVez = True
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).Text = ""
            End If
         Next
         InicializaApelidos COM_TEXTBOX
         On Error Resume Next
                  txtCampo(2).Value = Date
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
      Historico = IIf(IsNull(vgTb!Historico), "", vgTb!Historico)
      Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
      Sequencia_da_nota_fiscal = IIf(IsNull(vgTb![Sequencia da nota fiscal]), 0, vgTb![Sequencia da nota fiscal])
      Data_da_Alteracao = vgTb![Data da Alteração]
      Hora_da_Alteracao = vgTb![Hora da Alteração]
      Fechado = IIf(IsNull(vgTb!Fechado), 0, vgTb!Fechado)
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
      ExecutaPreValidacao
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_EDITANDO Then
         AjustaValores
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
   Set Devolucao_de_compra = vgTb
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
   RepositionDevolucao  
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
   txtCampo(2).Enabled = (Ok And txtCampo(2).Editable)
   Ok = (False)
   txtCampo(3).Enabled = (Ok And txtCampo(3).Editable)
   Ok = (False)
   txtCampo(5).Enabled = (Ok And txtCampo(5).Editable)
   Ok = (Cancelado = False AND Fechado = True And vgSituacao = ACAO_NAVEGANDO AND Sequencia_da_nota_fiscal = 0)
   Botao(0).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And Cancelado = False AND Fechado = False)
   Botao(1).Enabled = Ok
   Ok = (Sequencia_da_nota_fiscal = 0 AND vgSituacao = ACAO_NAVEGANDO AND Cancelado = False AND Fechado = True AND PermitidoMenu(MnuNFe.Tag) And Empresa![Emite NFe])
   Botao(2).Enabled = Ok
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Botao(0).Visible = (Fechado = True And Sequencia_da_nota_fiscal = 0)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 5 Or Index = 10 Then
      ExecutaVisivel
      ExecutaPreValidacao
      MostraFormulas
   End If
End Sub

'evento - quando o campo receber o foco
Private Sub txtCp_GotFocus(Index As Integer)
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
   PosicionaRegistro frmDevCompr, "Sequencia da devolução", seqRegistro
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
   Caption = LoadGasString(26230)
   vgFormID = 655
   vgIdentTab = "Devolução de compra"
   vgFiltroEmUso = -1
   vgIndDefault = ""
   vgPriVez = True
   vgPodeFazerUnLoad = False
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = False
   vgTemFiltro = False
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 18
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(17735))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(17736))
   Set Botao(0).PictureDisabled = LoadPicture(LoadGasPicture(17737))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(17738))
   Set Botao(2).Picture = LoadPicture(LoadGasPicture(17739))
   Set lblAjuste = Label(0)
   Set lblDevolucao = Label(1)
   Set lblLinha = Label(2)
   Set GrdProdutos = Grid(0)
   Set lblParcelamento = Label(9)
   Set grdParcelamento = Grid(1)
   Set lblStatus = Label(21)
   Set cmdAbrir = Botao(0)
   Set cmdFechar = Botao(1)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia da devolução] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(26235)
   Label(2).Caption = LoadGasString(26236)
   Label(3).Caption = LoadGasString(26237)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(26238)
   Label(4).Caption = LoadGasString(26239)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(26240)
   Label(5).Caption = LoadGasString(26241)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(26242)
   Label(6).Caption = LoadGasString(26243)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(26244)
   Label(7).Caption = LoadGasString(26245)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(26246)
   Label(8).Caption = LoadGasString(26247)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(26248)
   Tabs(0).TabCaption(0) = LoadGasString(26249)
   Tabs(0).TabCaption(1) = LoadGasString(26250)
   Label(10).Caption = LoadGasString(26251)
   Label(11).Caption = LoadGasString(26252)
   Label(12).Caption = LoadGasString(26253)
   Label(13).Caption = LoadGasString(26254)
   Label(14).Caption = LoadGasString(26255)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(26256)
   Label(15).Caption = LoadGasString(26257)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(26258)
   Label(16).Caption = LoadGasString(26259)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(26260)
   Label(17).Caption = LoadGasString(26261)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(26262)
   Label(18).Caption = LoadGasString(26263)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(26264)
   Label(19).Caption = LoadGasString(26265)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(26266)
   Label(20).Caption = LoadGasString(26267)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(26268)
   Label(21).Caption = LoadGasString(26269)
   vgTooltips.AddTool Botao(0), 0, LoadGasString(26270)
   Botao(0).Caption = LoadGasString(26271)
   vgTooltips.AddTool Botao(1), 0, LoadGasString(26272)
   Botao(1).Caption = LoadGasString(26273)
   Botao(2).Caption = LoadGasString(26274)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiNFE.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiNFE.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Produtos devolução compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Produtos devolução compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos devolução compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Sequencia do produto", TP_NUMERICO, "", , True, , "DADOSNFE", "Produtos", "Sequencia do produto", "Descrição; Sequencia do produto", "Descrição; Sequencia do produto", "Descrição; Sequencia do produto", "", , "1", , , , "0", 1, "0", 4935
      .AddColumn Nothing, , "Nº Item", "Sequencia do item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 900
      .AddColumn Nothing, , "Qtd Restante", , TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1260, , 12648447
      .AddColumn Nothing, , "Quantidade", "Quantidade", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1260
      .AddColumn Nothing, , "Estoque", , TP_NUMERICO, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 1200
      .AddColumn Nothing, , "Valor Unitario", "Valor unitario", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1320
      .AddColumn Nothing, , "Valor Total", , TP_NUMERICO, "999.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "% IPI", "Aliquota do ipi", TP_NUMERICO, "99,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1065
      .AddColumn Nothing, , "Valor do IPI", , TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1305
      .AddColumn Nothing, , "Valor ICMS ST", "Valor icms st", TP_NUMERICO, "99.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1455
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
      .AllowInsert = Permitido("Parcelas dev compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas dev compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas dev compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Numero da parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 660
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 630
      .AddColumn Nothing, , "Dt. Vencimento", "Data de vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 2160
      .AddColumn Nothing, , "Valor", "Valor da parcela", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1440
      .AddColumn Nothing, , "Tipo de Pagto:", "Especie de pagamento", TP_CARACTER, , , False, , , , , , "Dinheiro|Cartão|Cheque - Pre|Cheque - Vista|Boleto|Condicional|Garantia|Deposito", , , , "1", , , , "0", 1, "0", 1905
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Devolucao_de_compra = vgTb
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
   txtCampo(1).DataType = 1
   txtCampo(1).Mask = "999999"
   txtCampo(1).Editable = False
   txtCampo(1).BoundColumn = ""
   txtCampo(1).ListFields = ""
   txtCampo(1).OrderFields = ""
   txtCampo(1).Relation = ""
   txtCampo(1).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(1).DataField), txtCampo(1)

   Set txtCampo(2).CtPri = txtCp(2)
   txtCampo(2).DataType = 2
   txtCampo(2).Mask = "99/99/9999"
   txtCampo(2).Editable = False
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 2
   txtCampo(3).Mask = "99/99/9999"
   txtCampo(3).Editable = False
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "@!"
   txtCampo(4).Editable = False
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
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "999999"
   txtCampo(5).Editable = False
   txtCampo(5).BoundColumn = "Sequencia do geral"
   txtCampo(5).ListFields = "Razão social"
   txtCampo(5).OrderFields = "Sequencia do geral"
   txtCampo(5).Relation = ""
   txtCampo(5).Source = "Geral"
   txtCampo(5).vgfrmGMCale.grdListaG.AutoRebind = True
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
   txtCampo(8).Mask = "999.999.999,99"
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
   Set txtCampo(10).CtFdo = labFdo10
   Set txtCampo(10).CtBot(BOT_COMBO) = bottxtCampo10(BOT_COMBO)
   Set bottxtCampo10(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(10).DataType = 0
   txtCampo(10).ListFields = "Vista|Prazo"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "99.999.999,99"
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 1
   txtCampo(12).Mask = "99.999.999,99"
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

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
      Case 5
         txtCampo(5).Filter = "[Sequencia do geral] > " & 0 & ""
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
   FinalizaForm Me
   Set lblAjuste = Nothing
   Set lblDevolucao = Nothing
   Set lblLinha = Nothing
   Set GrdProdutos = Nothing
   Set lblParcelamento = Nothing
   Set grdParcelamento = Nothing
   Set lblStatus = Nothing
   Set cmdAbrir = Nothing
   Set cmdFechar = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   If Not Devolucao_de_compra Is Nothing Then
      Set Devolucao_de_compra = Nothing
   End If
   If Not Produtos_devolucao_compra Is Nothing Then
      Produtos_devolucao_compra.CloseRecordSet
      Set Produtos_devolucao_compra = Nothing
   End If
   If Not Parcelas_dev_compra Is Nothing Then
      Parcelas_dev_compra.CloseRecordSet
      Set Parcelas_dev_compra = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmDevCompr = Nothing                      'libera o segmento de código do form
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
         Produtos_devolucao_compra![Sequencia da devolução] = Devolucao_de_compra![Sequencia da devolução]
      Case 1
         Parcelas_dev_compra![Sequencia da devolução] = Devolucao_de_compra![Sequencia da devolução]
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
         Set Produtos_devolucao_compra = vgNewRecordSet
      Case 1
         Set Parcelas_dev_compra = vgNewRecordSet
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
   Dim Valor_icms_st As Double, Aliquota_do_ipi As Single, Sequencia_do_movimento As Long
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
            vgRetVal = (InfoMovimento(Sequencia_da_devolucao, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, Sequencia_do_movimento, "Qtde"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 5
            vgRetVal = (InfoProdutos(Sequencia_da_devolucao, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, Sequencia_do_movimento, "Estoque"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = (Quantidade * Valor_unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 9
            vgRetVal = ((Round(Round(Quantidade * Valor_unitario, 2) * Aliquota_do_ipi / 100, 2)))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = (False)
      On Error Resume Next
         Grid(0).AllowEdit = ((Devolucao_de_compra!Cancelado = 0) AND Devolucao_de_compra!Fechado = 0)
      On Error Resume Next
         Grid(0).AllowDelete = (False)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Produtos devolução compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da devolução] = " & Devolucao_de_compra![Sequencia da devolução]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "[Sequencia do produto] > " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (ValidaQuantidade(Sequencia_da_devolucao, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, Sequencia_do_movimento))
         If Not vgIsValid Then vgColumn =  4 
         vgErrorMessage$ = "Quantidade não pode ser maior que a Quantidade Restante do Movimento de Entrada!"
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaProdutos Sequencia_da_devolucao, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, Sequencia_do_movimento 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            DeletaProdutos Sequencia_da_devolucao, Sequencia_do_item, Sequencia_do_produto, Quantidade, Valor_unitario, Valor_total, _
   Valor_do_ipi, Valor_icms_st, Aliquota_do_ipi, Sequencia_do_movimento 
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Sequencia_do_item = ColumnValue(2)
   Quantidade = ColumnValue(4)
   Valor_unitario = ColumnValue(6)
   Aliquota_do_ipi = ColumnValue(8)
   Valor_icms_st = ColumnValue(10)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Produtos_devolucao_compra.Eof = False And Produtos_devolucao_compra.BOF = False And Produtos_devolucao_compra.RecordCount > 0 Then
         Valor_total = Produtos_devolucao_compra![Valor total]
         Valor_do_ipi = Produtos_devolucao_compra![Valor do ipi]
         Sequencia_do_movimento = Produtos_devolucao_compra![Sequencia do movimento]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_devolucao_compra.Eof = False And Produtos_devolucao_compra.BOF = False And Produtos_devolucao_compra.RecordCount > 0 Then
      Sequencia_do_item = Produtos_devolucao_compra![Sequencia do item]
      Sequencia_do_produto = Produtos_devolucao_compra![Sequencia do produto]
      Quantidade = Produtos_devolucao_compra!Quantidade
      Valor_unitario = Produtos_devolucao_compra![Valor unitario]
      Valor_total = Produtos_devolucao_compra![Valor total]
      Valor_do_ipi = Produtos_devolucao_compra![Valor do ipi]
      Valor_icms_st = Produtos_devolucao_compra![Valor icms st]
      Aliquota_do_ipi = Produtos_devolucao_compra![Aliquota do ipi]
      Sequencia_do_movimento = Produtos_devolucao_compra![Sequencia do movimento]
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
   Dim Especie_de_pagamento As String, Numero_da_parcela As Integer, Dias As Integer
   Dim Data_de_vencimento As Variant, Valor_da_parcela As Double
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
            vgRetVal = Not ((Not Vazio(Devolucao_de_compra![Forma de pagamento])) AND Devolucao_de_compra![Forma de pagamento] = "Prazo")
         Case 3
            vgRetVal = Not ((Not Vazio(Devolucao_de_compra![Forma de pagamento])) AND Devolucao_de_compra![Forma de pagamento] = "Prazo")
         Case 4
            vgRetVal = Not (Not Vazio(Devolucao_de_compra![Forma de pagamento]))
         Case 5
            vgRetVal = Not (Not Vazio(Devolucao_de_compra![Forma de pagamento]))
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = ((Devolucao_de_compra!Cancelado = 0) AND Devolucao_de_compra!Fechado = 0 And TotalParcelas < Devolucao_de_compra![Valor total da devolução])
      On Error Resume Next
         Grid(1).AllowEdit = ((Devolucao_de_compra!Cancelado = 0) AND Devolucao_de_compra!Fechado = 0)
      On Error Resume Next
         Grid(1).AllowDelete = ((Devolucao_de_compra!Cancelado = 0) AND Devolucao_de_compra!Fechado = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas dev compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Sequencia da devolução] = " & Devolucao_de_compra![Sequencia da devolução]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_parcela = UltimaParcela
            vgRetVal = Numero_da_parcela
         Case 2
            Dias = DateDiff("D",Devolucao_de_compra![Data da devolução] , Data_de_vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_vencimento = DateAdd("D", Dias, Devolucao_de_compra![Data da devolução])
            vgRetVal = Data_de_vencimento
         Case 4
            Valor_da_parcela = Devolucao_de_compra![Valor total da devolução] - TotalParcelas()
            vgRetVal = Valor_da_parcela
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If (Not Vazio(Devolucao_de_compra![Forma de pagamento])) AND Devolucao_de_compra![Forma de pagamento] = "Prazo" Then
            vgIsValid = (1 = 1)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Dias inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If (Not Vazio(Devolucao_de_compra![Forma de pagamento])) AND Devolucao_de_compra![Forma de pagamento] = "Prazo" Then
               vgIsValid = ((Not Vazio(Data_de_vencimento)) And (IsDate(Data_de_vencimento) Or Vazio(Data_de_vencimento)))
               If Not vgIsValid Then vgColumn =  3 
               vgErrorMessage$ = "Vencimento Invalido!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Devolucao_de_compra![Forma de pagamento]) Then
               vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, TotalParcelas() + Valor_da_Parcela <= Devolucao_de_compra![Valor total da devolução],  Valor_da_parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_parcela)) + Valor_da_Parcela <= Devolucao_de_compra![Valor total da devolução]))
               If Not vgIsValid Then vgColumn =  4 
               vgErrorMessage$ = "A Soma das Parcelas não Correspondem ao Valor Total do Movimento!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            If Not Vazio(Devolucao_de_compra![Forma de pagamento]) Then
               vgIsValid = (Not Vazio(Especie_de_pagamento))
               If Not vgIsValid Then vgColumn =  5 
               vgErrorMessage$ = "Tipo de Pagto Invalido!"
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
   If Parcelas_dev_compra.Eof = False And Parcelas_dev_compra.BOF = False And Parcelas_dev_compra.RecordCount > 0 Then
      Especie_de_pagamento = Parcelas_dev_compra![Especie de pagamento]
      Numero_da_parcela = Parcelas_dev_compra![Numero da parcela]
      Dias = Parcelas_dev_compra!Dias
      Data_de_vencimento = Parcelas_dev_compra![Data de vencimento]
      Valor_da_parcela = Parcelas_dev_compra![Valor da parcela]
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
         Reabrir  
      Case 1
         PreValidaFechamento  
      Case 2
         AbreGerar  
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo5_Click(Index As Integer)
   txtCampo(5).SetFocus
   DoEvents
   txtCampo(5).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo10_Click(Index As Integer)
   txtCampo(10).SetFocus
   DoEvents
   txtCampo(10).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


