Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProdutos
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   6090
   Left                     =   2310
   LinkTopic                =   "Produtos"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   5655
   ScaleWidth               =   10800
   Top                      =   2440
   Width                    =   10890
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
      Height                   =   5490
      Left                     =   15
      Width                    =   10620
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   18723
      _extenty                 =   9678
      Index                    =   0 
      Begin NFE.GCpMM mmCampo
         Appearance               =   0
         BackColor                =   &H8000000F
         BorderStyle              =   0
         Stretch                  =   0
         Height                   =   1470
         Left                     =   240
         Width                    =   1500
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
         Left                     =   1980
         Width                    =   1575
         Top                      =   360
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
         Caption                  =   "Produtos"
         Index                    =   0 
      End
      Begin VB.Line Linha
         BorderColor              =   -2147483640
         BorderStyle              =   1
         BorderWidth              =   1
         X1                       =   2040
         Y1                       =   810
         X2                       =   3540
         Y2                       =   810
         Index                    =   0 
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
         TabIndex                 =   11
         TabStop                  =   0
         Height                   =   255
         Left                     =   2040
         Width                    =   3825
         Top                      =   1080
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
         Caption                  =   "Material Adquirido de Terceiros"
         Index                    =   0 
         DataField                =   "Material adquirido de terceiro"
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
         TabIndex                 =   12
         TabStop                  =   0
         Height                   =   255
         Left                     =   5940
         Width                    =   1545
         Top                      =   1080
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
         Caption                  =   "Importado"
         Index                    =   1 
         DataField                =   "Importado"
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
         TabStop                  =   0
         Height                   =   255
         Left                     =   7560
         Width                    =   1020
         Top                      =   1080
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
         Caption                  =   "Usado"
         Index                    =   2 
         DataField                =   "Usado"
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
         Left                     =   8880
         Width                    =   1185
         Top                      =   120
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
         Caption                  =   "Inativo"
         Index                    =   3 
         DataField                =   "Inativo"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   240
         Left                     =   10320
         Width                    =   45
         Top                      =   5580
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
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   2
         TabHeight                =   519
         BackColor                =   &H8000000F
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   2
         TabIndex                 =   15
         TabStop                  =   0
         Height                   =   3225
         Left                     =   300
         Width                    =   10170
         Top                      =   1800
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
         Caption                  =   "1 - Dados Principais"
         Index                    =   0 
         _version                 =   65536
         _extentx                 =   17929
         _extenty                 =   5685
         _stockprops              =   15
         TabCaption(0)            =   "1 - Dados Principais"
         Tab(0).ControlCount=   32
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Label(2)"
         Tab(0).Control(1)=   "txtCp(0)"
         Tab(0).Control(2)=   "Label(3)"
         Tab(0).Control(3)=   "txtCp(1)"
         Tab(0).Control(4)=   "Label(4)"
         Tab(0).Control(5)=   "labFdo2"
         Tab(0).Control(6)=   "bottxtCampo2(1)"
         Tab(0).Control(7)=   "bottxtCampo2(2)"
         Tab(0).Control(8)=   "txtCp(2)"
         Tab(0).Control(9)=   "Botao(0)"
         Tab(0).Control(10)=   "Label(5)"
         Tab(0).Control(11)=   "txtCp(3)"
         Tab(0).Control(12)=   "Label(6)"
         Tab(0).Control(13)=   "txtCp(4)"
         Tab(0).Control(14)=   "txtCp(5)"
         Tab(0).Control(15)=   "Label(7)"
         Tab(0).Control(16)=   "Botao(1)"
         Tab(0).Control(17)=   "labFdo6"
         Tab(0).Control(18)=   "bottxtCampo6(1)"
         Tab(0).Control(19)=   "bottxtCampo6(2)"
         Tab(0).Control(20)=   "txtCp(6)"
         Tab(0).Control(21)=   "Label(8)"
         Tab(0).Control(22)=   "Label(9)"
         Tab(0).Control(23)=   "txtCp(7)"
         Tab(0).Control(24)=   "Label(10)"
         Tab(0).Control(25)=   "txtCp(8)"
         Tab(0).Control(26)=   "Label(11)"
         Tab(0).Control(27)=   "txtCp(9)"
         Tab(0).Control(28)=   "Label(14)"
         Tab(0).Control(29)=   "txtCp(12)"
         Tab(0).Control(30)=   "Label(15)"
         Tab(0).Control(31)=   "txtCp(13)"
         TabCaption(1)            =   "2 - Produto + Tributos"
         Tab(1).ControlCount=   4
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Label(12)"
         Tab(1).Control(1)=   "txtCp(10)"
         Tab(1).Control(2)=   "Label(13)"
         Tab(1).Control(3)=   "txtCp(11)"
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   375
            Width                    =   705
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
            Caption                  =   "Codigo:"
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
            TabIndex                 =   0
            TabStop                  =   -1
            Height                   =   300
            Left                     =   1215
            Width                    =   840
            Top                      =   675
            DataField                =   "Sequencia do produto"
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
            Left                     =   4815
            Width                    =   1125
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
            Caption                  =   "*Descrição:"
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
            MaxLength                =   80
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   2
            TabStop                  =   -1
            Height                   =   300
            Left                     =   5985
            Width                    =   3945
            Top                      =   675
            DataField                =   "Descrição"
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   180
            Width                    =   945
            Top                      =   1275
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
            Caption                  =   "*Unidade:"
            Index                    =   4 
         End
         Begin VB.Label labFdo2
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   1215
            Top                      =   1275
            Width                    =   3545
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo2
            Style                    =   1
            TabStop                  =   0
            Index                    =   1
            Left                     =   4490
            Top                      =   1290
            Width                    =   255
            Height                   =   270
         End
         Begin VB.CommandButton bottxtCampo2
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   4235
            Top                      =   1290
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
            Left                     =   1230
            Width                    =   2990
            Top                      =   1290
            DataField                =   "Sequencia da unidade"
            Index                    =   2 
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
            Left                     =   4755
            Width                    =   240
            Top                      =   1275
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
            AutoSize                 =   0
            Height                   =   240
            Left                     =   5115
            Width                    =   1095
            Top                      =   1275
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
            Caption                  =   "*Vr. Venda.:"
            Index                    =   5 
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
            TabIndex                 =   4
            TabStop                  =   -1
            Height                   =   300
            Left                     =   6315
            Width                    =   1170
            Top                      =   1275
            DataField                =   "Valor total"
            Index                    =   3 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   7635
            Width                    =   1035
            Top                      =   1275
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
            Caption                  =   "Estoque.....:"
            Index                    =   6 
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
            TabIndex                 =   5
            TabStop                  =   -1
            Height                   =   300
            Left                     =   8790
            Width                    =   1140
            Top                      =   1275
            DataField                =   "Estoque"
            TabStop                  =  0
            Index                    =   4 
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
            MaxLength                =   254
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   7
            TabStop                  =   0
            Height                   =   300
            Left                     =   6315
            Width                    =   1170
            Top                      =   1875
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
            Left                     =   5145
            Width                    =   960
            Top                      =   1875
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
            Caption                  =   "% de IPI....:"
            Index                    =   7 
         End
         Begin NFE.GBotao Botao
            ButtonStyle              =   1
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   17
            TabStop                  =   0
            ButtonType               =   0
            Height                   =   270
            Left                     =   4755
            Width                    =   240
            Top                      =   1875
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
            Index                    =   1 
         End
         Begin VB.Label labFdo6
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   1215
            Top                      =   1875
            Width                    =   3545
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo6
            Style                    =   1
            TabStop                  =   0
            Index                    =   1
            Left                     =   4490
            Top                      =   1890
            Width                    =   255
            Height                   =   270
         End
         Begin VB.CommandButton bottxtCampo6
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   4235
            Top                      =   1890
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
            Left                     =   1230
            Width                    =   2990
            Top                      =   1890
            DataField                =   "Sequencia da classificação"
            Index                    =   6 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   435
            Width                    =   690
            Top                      =   1875
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
            Caption                  =   "*NCM.:"
            Index                    =   8 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   7635
            Width                    =   1080
            Top                      =   1875
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
            Caption                  =   "Ultima NFe:"
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
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   8
            TabStop                  =   -1
            Height                   =   300
            Left                     =   8790
            Width                    =   1140
            Top                      =   1875
            DataField                =   "Ultima nfe"
            TabStop                  =  0
            Index                    =   7 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   2595
            Width                    =   1035
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
            Caption                  =   "Ult.Entrada:"
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
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   9
            TabStop                  =   -1
            Height                   =   300
            Left                     =   3825
            Width                    =   1170
            Top                      =   2460
            DataField                =   "Ultima entrada"
            TabStop                  =  0
            Index                    =   8 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   5145
            Width                    =   1005
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
            Caption                  =   "Ult. Fornec:"
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
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   10
            TabStop                  =   0
            Height                   =   300
            Left                     =   6330
            Width                    =   3600
            Top                      =   2460
            DataField                =   "Ultimo fornecedor"
            TabStop                  =  0
            Index                    =   9 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -74640
            Width                    =   1305
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
            Caption                  =   "Valor Unitário.:"
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
            TabIndex                 =   19
            TabStop                  =   0
            Height                   =   300
            Left                     =   -73260
            Width                    =   1425
            Top                      =   660
            DataField                =   "Valor unitario"
            TabStop                  =  0
            Index                    =   10 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -71640
            Width                    =   2385
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
            Caption                  =   "Valor Unitário + Impostos:"
            Index                    =   13 
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
            TabIndex                 =   20
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -69180
            Width                    =   1425
            Top                      =   660
            DataField                =   "Valor de custo"
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
            Height                   =   240
            Left                     =   180
            Width                    =   930
            Top                      =   2460
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
            Caption                  =   "Ult.Venda:"
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
            Alignment                =   0
            Locked                   =   -1
            TabIndex                 =   21
            TabStop                  =   0
            Height                   =   300
            Left                     =   1200
            Width                    =   1215
            Top                      =   2460
            DataField                =   "Ultima venda"
            TabStop                  =  0
            Index                    =   12 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   2115
            Width                    =   1065
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
            Caption                  =   "Cod Barras:"
            Index                    =   15 
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
            TabIndex                 =   1
            TabStop                  =   -1
            Height                   =   300
            Left                     =   3240
            Width                    =   1500
            Top                      =   675
            DataField                =   "Codigo de Barras"
            Index                    =   13 
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
         TabIndex                 =   18
         TabStop                  =   0
         Height                   =   255
         Left                     =   8880
         Width                    =   1260
         Top                      =   1080
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
         Caption                  =   "Despesa"
         Index                    =   4 
         DataField                =   "Despesa"
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
Attribute VB_name = "frmProdutos"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: Emissor de NFe
'* Empresa...: Ygor Eduardo Dellalio
'* Módulo....: Produtos
'* Função....: Produtos
'* CopyRight.: (C)2022 Ygor Eduardo Dellalio
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 09/03/2022 14:43:30
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
Dim txtCampo(13) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(4) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim Ultima_nfe As Variant, Sequencia_do_produto As Double, Descricao As String
Dim Estoque As Double, Sequencia_da_unidade As Integer, Valor_total As Double
Dim Sequencia_da_classificacao As Long, Usado As Boolean, Inativo As Boolean
Dim Material_adquirido_de_terceiro As Boolean, Importado As Boolean, Ultima_entrada As Variant
Dim Ultimo_fornecedor As Long, Despesa As Boolean, Valor_unitario As Double
Dim Valor_de_custo As Double, Ultima_venda As Variant, Codigo_de_Barras As String
Dim Ajuste As String
Public lblAjuste As Object

Private Classificacao As GRecordSet

Private Sub RepositionProduto()   
   On error Resume Next
   TbAuxiliar "Classificação Fiscal", "[Sequencia da Classificação] = " & Sequencia_da_classificacao, Classificacao      
End Sub

Private Function PegaSequencial() As Long
Dim Tb As New GRecordSet
 
 If VgSituacao = ACAO_INCLUINDO Then
    Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do produto]) Seq From Produtos")
    PegaSequencial = Tb!Seq + 1
 End If
 
End Function 


Private Sub VerificaProduto()
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("Select * From Produtos Where [Descrição] = '" & Descricao & "'")
  
 If Tb.RecordCount > 0 And Descricao <> "" And VgSituacao <> ACAO_NAVEGANDO Then
  If Tb![Sequencia Do produto] <> Sequencia_do_produto Then
   If MsgBox("Ja Existe um Cadastro com esta Descrição do Produto" & " " & Tb![Descrição] & " Sequencia do Produto" & " --> " & " 00" & Tb![Sequencia Do produto] & "   Deseja Alterar os Dados?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then 
      MDINFE.CancelaAlteracoes
   Else
      MDINFE.CancelaAlteracoes
      PosicionaRegistro frmProdutos, "Sequencia do produto", Tb![Sequencia Do produto]
   End If
  End If 
 End If 
   
End Sub 


Private Sub ValidaCodProd()
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("Select * From Produtos Where [Sequencia do produto] = " & Sequencia_do_produto)
  
  If Tb.RecordCount > 0 And Sequencia_do_produto > 0 And VgSituacao = ACAO_INCLUINDO Then
   If MsgBox("Ja Existe um Cadastro com este Codigo" & " Descrição -> " & Tb![Descrição] & "   Deseja Alterar os Dados?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then 
      MDINFE.CancelaAlteracoes
   Else
      MDINFE.CancelaAlteracoes
      PosicionaRegistro frmProdutos, "Sequencia do produto", Tb![Sequencia Do produto]
   End If
  End If 
   
End Sub 

 
Private Function FiltroNCM() As String
 FiltroNCM = "[Sequencia da Classificação] > 0"
 
  If vgSituacao = ACAO_NAVEGANDO Then
      FiltroNCM = FiltroNCM & " And [Sequencia da classificação] > 0"
  End If
  
  If vgSituacao = ACAO_INCLUINDO Or vgSituacao = ACAO_EDITANDO Then
      FiltroNCM = FiltroNCM & " And Inativo = 0"
  End If
 
End Function 


Private Function ValidaNCM() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From [Classificação Fiscal] Where [Sequencia da Classificação] = " & Sequencia_da_classificacao)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaNCM = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaNCM = True: Exit Function      
   End If
  End If
  
End Function


Private Function ValidaUnidade() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Unidades Where [Sequencia da unidade] = " & Sequencia_da_unidade)
 
  If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaUnidade = False
      MDINFE.CancelaAlteracoes: Exit Function
   Else
      ValidaUnidade = True: Exit Function      
   End If
  End If
  
End Function

Private Sub Tabs_Click(Index As Integer, PreviousTab As Integer)
   MudaTamCampos Me
End Sub


Private Function SeqProduto() As Long
 Dim Tb As New GRecordSet
 
 If Not Empresa![Numeração Auto] Then Exit Function
 
 Set Tb = vgDb.OpenRecordSet("SELECT MAX([Sequencia do Produto]) Seq From Produtos")
  
  If Tb.RecordCount = 0 Then
     SeqProduto = 1
  Else
     SeqProduto = Tb!Seq + 1
  End If   

End Function

Public Sub LigaDesligaBotoes
   Botao( 0 ).Enabled =    Botao( 0 ).Enabled And Permitido("Unidades", ACAO_NAVEGANDO)
   Botao( 1 ).Enabled =    Botao( 1 ).Enabled And Permitido("Classificação fiscal", ACAO_NAVEGANDO)
End Sub

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Ultima_nfe = txtCampo(7).Value
         Sequencia_do_produto = txtCampo(0).Value
         Descricao = txtCampo(1).Value
         Estoque = txtCampo(4).Value
         Sequencia_da_unidade = txtCampo(2).Value
         Valor_total = txtCampo(3).Value
         Sequencia_da_classificacao = txtCampo(6).Value
         Usado = chkCampo(2).Value
         Inativo = chkCampo(3).Value
         Material_adquirido_de_terceiro = chkCampo(0).Value
         Importado = chkCampo(1).Value
         Ultima_entrada = txtCampo(8).Value
         Ultimo_fornecedor = txtCampo(9).Value
         Despesa = chkCampo(4).Value
         Valor_unitario = txtCampo(10).Value
         Valor_de_custo = txtCampo(11).Value
         Ultima_venda = txtCampo(12).Value
         Codigo_de_Barras = txtCampo(13).Value
      Else
         Ultima_nfe = vgTb![Ultima nfe]
         Sequencia_do_produto = IIf(IsNull(vgTb![Sequencia do produto]), 0, vgTb![Sequencia do produto])
         Descricao = IIf(IsNull(vgTb!Descrição), "", vgTb!Descrição)
         Estoque = IIf(IsNull(vgTb!Estoque), 0, vgTb!Estoque)
         Sequencia_da_unidade = IIf(IsNull(vgTb![Sequencia da unidade]), 0, vgTb![Sequencia da unidade])
         Valor_total = IIf(IsNull(vgTb![Valor total]), 0, vgTb![Valor total])
         Sequencia_da_classificacao = IIf(IsNull(vgTb![Sequencia da classificação]), 0, vgTb![Sequencia da classificação])
         Usado = IIf(IsNull(vgTb!Usado), 0, vgTb!Usado)
         Inativo = IIf(IsNull(vgTb!Inativo), 0, vgTb!Inativo)
         Material_adquirido_de_terceiro = IIf(IsNull(vgTb![Material adquirido de terceiro]), 0, vgTb![Material adquirido de terceiro])
         Importado = IIf(IsNull(vgTb!Importado), 0, vgTb!Importado)
         Ultima_entrada = vgTb![Ultima entrada]
         Ultimo_fornecedor = IIf(IsNull(vgTb![Ultimo fornecedor]), 0, vgTb![Ultimo fornecedor])
         Despesa = IIf(IsNull(vgTb!Despesa), 0, vgTb!Despesa)
         Valor_unitario = IIf(IsNull(vgTb![Valor unitario]), 0, vgTb![Valor unitario])
         Valor_de_custo = IIf(IsNull(vgTb![Valor de custo]), 0, vgTb![Valor de custo])
         Ultima_venda = vgTb![Ultima venda]
         Codigo_de_Barras = IIf(IsNull(vgTb![Codigo de Barras]), "", vgTb![Codigo de Barras])
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
      i = Permitido(vgIdentTab, ACAO_EXCLUINDO)
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
   If vgOq = VALIDACOES Then
      InicializaApelidos COM_TEXTBOX
      If vgSituacao = ACAO_INCLUINDO And Not Empresa![Numeração Auto] Then
         vgOk = (Sequencia_do_produto > 0)
         vgMsg$ = "Sequencia do produto inválido!"
         If Not vgOk Then vgColumn = 1
      End If
      If vgOk Then
         vgOk = (VDvCB(Codigo_de_Barras))
         vgMsg$ = "Codigo de Barras inválido!"
         If Not vgOk Then vgColumn = 14
      End If
      If vgOk Then
         vgOk = (Not Vazio(Descricao))
         vgMsg$ = "Descrição não pode ser vazio!"
         If Not vgOk Then vgColumn = 2
      End If
      If vgOk Then
         If Sequencia_da_unidade = 0 Then
            vgOk = (Sequencia_da_unidade > 0)
            vgMsg$ = "Unidade não pode ser em Branco!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         If vgSituacao = ACAO_EDITANDO Then
            vgOk = (ValidaUnidade())
            vgMsg$ = "Unidade Inativa!"
            If Not vgOk Then vgColumn = 3
         End If
      End If
      If vgOk Then
         vgOk = (Valor_total > 0)
         vgMsg$ = "Valor Unitario Invalido!"
         If Not vgOk Then vgColumn = 4
      End If
      If vgOk Then
         If Despesa = 0 Then
            vgOk = (Sequencia_da_classificacao > 0)
            vgMsg$ = "NCM não pode ser em Branco!"
            If Not vgOk Then vgColumn = 7
         End If
      End If
      If vgOk Then
         If vgSituacao = ACAO_EDITANDO Then
            vgOk = (ValidaNCM())
            vgMsg$ = "NCM Inativo!"
            If Not vgOk Then vgColumn = 7
         End If
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
                  chkCampo(0).Value = 1
         chkCampo(1).Value = False
         chkCampo(2).Value = False
         chkCampo(3).Value = False
         txtCampo(0).Value = SeqProduto()
         txtCampo(12).Value = Date
         chkCampo(4).Value = False
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 6
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 2
         PoeRelEFiltroCbo 6
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Material_adquirido_de_terceiro
         chkCampo(1).Value = Importado
         chkCampo(2).Value = Usado
         chkCampo(3).Value = Inativo
         chkCampo(4).Value = Despesa
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
      InicializaApelidos COM_TEXTBOX
      vgTb!Usado = Usado
      vgTb!Inativo = Inativo
      vgTb![Material adquirido de terceiro] = Material_adquirido_de_terceiro
      vgTb!Importado = Importado
      vgTb!Despesa = Despesa
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
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

'prepara botões e o formulário para o novo registro
Public Sub Reposition(Optional ForceRebind As Boolean, Optional LockGrids As Boolean = True)
   On Error GoTo DeuErro
   If vgPriVez Then Exit Sub
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   RepositionProduto  
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
   vgTemAlteracaoGrids = Not LockGrids
   If vgEmBrowse And vgSituacao = ACAO_NAVEGANDO And vgFrmImpCons Is Nothing Then grdBrowse.Refresh
DeuErro:
   
End Sub

'mostra fórmulas na janela
Private Sub MostraFormulas()
   On Error Resume Next                           'previne erros...
   txtCampo(5).Value = Classificacao![Percentual de ipi]
   If Err Then Err = 0: txtCampo(5).Text = ""
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'executa a pré-validação dos campos
Private Sub ExecutaPreValidacao()
   Dim Ok As Boolean, vgPV As Integer
   On Error Resume Next                           'prepara para possiveis erros
   vgPV = vgPriVez
   vgPriVez = True
   Ok = (vgSituacao = ACAO_INCLUINDO And Not Empresa![Numeração Auto])
   txtCampo(0).Enabled = (Ok And txtCampo(0).Editable)
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(1).Visible = (False)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 2 Or Index = 6 Or Index = 9 Then
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
         Case 0
            If ValBrasil(txtCp(0).Text) = 0 Then
               txtCampo(0).Value = SeqProduto()
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 2
            PoeRelEFiltroCbo 2
         Case 6
            PoeRelEFiltroCbo 6
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
   Select Case Index
      Case 0
         ValidaCodProd  
      Case 1
         VerificaProduto  
   End Select
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Material_adquirido_de_terceiro
      ElseIf Index = 1 Then
         chkCampo(1).Value = Importado

      ElseIf Index = 2 Then
         chkCampo(2).Value = Usado

      ElseIf Index = 3 Then
         chkCampo(3).Value = Inativo

      ElseIf Index = 4 Then
         chkCampo(4).Value = Despesa
      End If
   Else
   If Len(chkCampo(Index).DataField) > 0 Then LigaFocos Me
      InicializaApelidos COM_TEXTBOX
      MostraFormulas                              'mostra formulas na janela
      ExecutaVisivel                              'torna camos visiveis
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
   PosicionaRegistro frmProdutos, "Sequencia do produto", seqRegistro
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
   Caption = LoadGasString(16640)
   vgFormID = 623
   vgIdentTab = "Produtos"
   vgFiltroEmUso = -1
   vgIndDefault = "Sequencia do produto"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgTemCondicoesEsp = False
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 6
   vgSituacao = ACAO_NAVEGANDO
   Set mmCampo(0).Picture = LoadPicture(LoadGasPicture(8145))
   Set Botao(0).Picture = LoadPicture(LoadGasPicture(8146))
   Set Botao(1).Picture = LoadPicture(LoadGasPicture(8147))
   Set lblAjuste = Label(1)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "[Sequencia do produto] > " & 0
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(16645)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(16646)
   chkCampo(0).Caption = LoadGasString(16647)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(16648)
   chkCampo(1).Caption = LoadGasString(16649)
   vgTooltips.AddTool chkCampo(2).CtPri, 0, LoadGasString(16650)
   chkCampo(2).Caption = LoadGasString(16651)
   vgTooltips.AddTool chkCampo(3).CtPri, 0, LoadGasString(16652)
   chkCampo(3).Caption = LoadGasString(16653)
   Label(1).Caption = LoadGasString(16654)
   Tabs(0).TabCaption(0) = LoadGasString(16655)
   Tabs(0).TabCaption(1) = LoadGasString(16656)
   Label(2).Caption = LoadGasString(16657)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(16658)
   Label(3).Caption = LoadGasString(16659)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(16660)
   Label(4).Caption = LoadGasString(16661)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(16662)
   Label(5).Caption = LoadGasString(16663)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(16664)
   Label(6).Caption = LoadGasString(16665)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(16666)
   Label(7).Caption = LoadGasString(16667)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(16668)
   Label(8).Caption = LoadGasString(16669)
   Label(9).Caption = LoadGasString(16670)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(16671)
   Label(10).Caption = LoadGasString(16672)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(16673)
   Label(11).Caption = LoadGasString(16674)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(16675)
   Label(12).Caption = LoadGasString(16676)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(16677)
   Label(13).Caption = LoadGasString(16678)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(16679)
   Label(14).Caption = LoadGasString(16680)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(16681)
   Label(15).Caption = LoadGasString(16682)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(16683)
   vgTooltips.AddTool chkCampo(4).CtPri, 0, LoadGasString(16684)
   chkCampo(4).Caption = LoadGasString(16685)
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
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
   txtCampo(0).Mask = "9999999999"
   txtCampo(0).BoundColumn = ""
   txtCampo(0).ListFields = ""
   txtCampo(0).OrderFields = ""
   txtCampo(0).Relation = ""
   txtCampo(0).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(0).DataField), txtCampo(0)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 0
   txtCampo(13).Mask = "9999999999999"
   txtCampo(13).StuffChar = "0"
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(1).CtPri = txtCp(1)
   txtCampo(1).DataType = 0
   txtCampo(1).Mask = "@x"
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
   txtCampo(2).Mask = "9999"
   txtCampo(2).PesqModoAbertura =  2 
   txtCampo(2).PesqFieldCapture = "Unidades.[Sequencia da unidade]"
   txtCampo(2).BoundColumn = "Sequencia da unidade"
   txtCampo(2).ListFields = "Descrição; Sigla da unidade"
   txtCampo(2).OrderFields = "Descrição"
   txtCampo(2).Relation = ""
   txtCampo(2).Source = "Unidades"
   txtCampo(2).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 1
   txtCampo(3).Mask = "999.999.999,99"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 1
   txtCampo(4).Mask = "999.999,9999"
   txtCampo(4).Editable = False
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_LISTA) = bottxtCampo6(BOT_LISTA)
   Set txtCampo(6).CtBot(BOT_COMBO) = bottxtCampo6(BOT_COMBO)
   bottxtCampo6(BOT_LISTA).Caption = "P"
   Set bottxtCampo6(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(6).DataType = 1
   txtCampo(6).Mask = "99999"
   txtCampo(6).PesqModoAbertura =  2 
   txtCampo(6).PesqFieldCapture = "[Classificação fiscal].[Sequencia da classificação]"
   txtCampo(6).BoundColumn = "Sequencia da classificação"
   txtCampo(6).ListFields = "Ncm"
   txtCampo(6).OrderFields = "Sequencia da classificação"
   txtCampo(6).Relation = ""
   txtCampo(6).Source = "Classificação fiscal"
   txtCampo(6).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 1
   txtCampo(5).Mask = "999,9999"
   txtCampo(5).Editable = False
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""

   Set txtCampo(7).CtPri = txtCp(7)
   txtCampo(7).DataType = 2
   txtCampo(7).Mask = "99/99/9999"
   txtCampo(7).Editable = False
   txtCampo(7).BoundColumn = ""
   txtCampo(7).ListFields = ""
   txtCampo(7).OrderFields = ""
   txtCampo(7).Relation = ""
   txtCampo(7).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   txtCampo(8).DataType = 2
   txtCampo(8).Mask = "99/99/9999"
   txtCampo(8).Editable = False
   txtCampo(8).BoundColumn = ""
   txtCampo(8).ListFields = ""
   txtCampo(8).OrderFields = ""
   txtCampo(8).Relation = ""
   txtCampo(8).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "999999"
   txtCampo(9).Editable = False
   txtCampo(9).BoundColumn = "Sequencia do geral"
   txtCampo(9).ListFields = "Razão social"
   txtCampo(9).OrderFields = "Razão social"
   txtCampo(9).Relation = ""
   txtCampo(9).Source = "Geral"
   txtCampo(9).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(1).DataField), chkCampo(1)

   Set chkCampo(2).CtPri = chkCp(2)
   chkCampo(2).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(2).DataField), chkCampo(2)

   Set chkCampo(3).CtPri = chkCp(3)
   chkCampo(3).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(3).DataField), chkCampo(3)

   Set chkCampo(4).CtPri = chkCp(4)
   chkCampo(4).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(4).DataField), chkCampo(4)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = "999.999,9999"
   txtCampo(10).Editable = False
   txtCampo(10).BoundColumn = ""
   txtCampo(10).ListFields = ""
   txtCampo(10).OrderFields = ""
   txtCampo(10).Relation = ""
   txtCampo(10).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(10).DataField), txtCampo(10)

   Set txtCampo(11).CtPri = txtCp(11)
   txtCampo(11).DataType = 1
   txtCampo(11).Mask = "999.999,9999"
   txtCampo(11).Editable = False
   txtCampo(11).BoundColumn = ""
   txtCampo(11).ListFields = ""
   txtCampo(11).OrderFields = ""
   txtCampo(11).Relation = ""
   txtCampo(11).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(11).DataField), txtCampo(11)

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
         txtCampo(2).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Sequencia da unidade] > 0 AND Inativo = 0", "[Sequencia da unidade] > 0") & ""
                           txtCampo(2).PesqSQLExpression = "SELECT Unidades.[Sequencia da unidade], Unidades.Descrição, Unidades.[Sigla da unidade] FROM Unidades WHERE (Unidades.[Sequencia da unidade] > " + CStr(0) + ") AND " + _
                                                              "(Unidades.Inativo = " + CStr(0) + ")"
      Case 6
         txtCampo(6).Filter = "" + FiltroNCM() + ""
                           txtCampo(6).PesqSQLExpression = "SELECT [Classificação fiscal].[Sequencia da classificação], [Classificação fiscal].Ncm, [Classificação fiscal].[Descrição do ncm], " + _
                                                              "[Classificação fiscal].[Percentual de ipi], [Classificação fiscal].Iva, [Classificação fiscal].Cest FROM [Classificação fiscal] WHERE ([Classificação fiscal].[Sequencia da classificação] > " + CStr(0) + ") AND " + _
                                                              "([Classificação fiscal].Inativo = " + CStr(0) + ")"
      Case 9
         txtCampo(9).Filter = "[Sequencia do geral] > " & 0 & ""
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
   FinalizaForm Me
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   Set chkCampo(3) = Nothing
   Set chkCampo(4) = Nothing

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   grdBrowse.Finalize                             'descarrega o grid
   Set frmProdutos = Nothing                      'libera o segmento de código do form
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         seqRegistro = Sequencia_da_unidade 
         mdiNFE.MUnidades
      Case 1
         seqRegistro = Sequencia_da_classificacao 
         mdiNFE.MNcm
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo2_Click(Index As Integer)
   txtCampo(2).SetFocus
   DoEvents
   txtCampo(2).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo6_Click(Index As Integer)
   txtCampo(6).SetFocus
   DoEvents
   txtCampo(6).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


