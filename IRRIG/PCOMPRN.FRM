Version 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPComprN
   ForeColor                =   &H80000008
   BorderStyle              =   2
   Height                   =   8610
   Left                     =   2715
   LinkTopic                =   "PComprN"
   LockControls             =   -1
   KeyPreview               =   -1
   MDIChild                 =   -1
   ScaleHeight              =   8175
   ScaleWidth               =   14745
   Top                      =   1090
   Width                    =   14835
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
      Height                   =   7995
      Left                     =   15
      Width                    =   14565
      Top                      =   15
      TabStop                  =  0
      _extentx                 =   25678
      _extenty                 =   14095
      Index                    =   0 
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
         Height                   =   1170
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
         Left                     =   1875
         Width                    =   3600
         Top                      =   360
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
         ForeColor                =   &H0
         Caption                  =   "Pedido de Compra"
         Index                    =   0 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   135
         Width                    =   1170
         Top                      =   1440
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
         Caption                  =   "Sequencia:"
         Index                    =   1 
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000016
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
         Left                     =   1380
         Width                    =   840
         Top                      =   1440
         DataField                =   "Id do Pedido"
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
         Left                     =   2295
         Width                    =   465
         Top                      =   1440
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
         Caption                  =   "Data:"
         Index                    =   2 
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
         Left                     =   2820
         Width                    =   1125
         Top                      =   1440
         DataField                =   "Data do Pedido"
         Index                    =   1 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   4035
         Width                    =   870
         Top                      =   1440
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
         Caption                  =   "Licitação:"
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
         MaxLength                =   9
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   2
         TabStop                  =   -1
         Height                   =   300
         Left                     =   4980
         Width                    =   885
         Top                      =   1440
         DataField                =   "Nro da Licitação"
         Index                    =   2 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   5895
         Width                    =   1110
         Top                      =   1440
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
         Caption                  =   "Comprador:"
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
         MaxLength                =   30
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   3
         TabStop                  =   -1
         Height                   =   300
         Left                     =   7080
         Width                    =   1725
         Top                      =   1440
         DataField                =   "Comprador"
         Index                    =   3 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8895
         Width                    =   990
         Top                      =   1440
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
         Caption                  =   "Vendedor:"
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
         MaxLength                =   30
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   4
         TabStop                  =   -1
         Height                   =   300
         Left                     =   9960
         Width                    =   1665
         Top                      =   1440
         DataField                =   "Vend"
         Index                    =   4 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   11655
         Width                    =   645
         Top                      =   1440
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
         Caption                  =   "Prazo:"
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
         MaxLength                =   40
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   5
         TabStop                  =   -1
         Height                   =   300
         Left                     =   12360
         Width                    =   2025
         Top                      =   1440
         DataField                =   "Prazo"
         Index                    =   5 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   75
         Width                    =   1230
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
         Caption                  =   "*Frete..............:"
         Index                    =   7 
      End
      Begin VB.Label labFdo6
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000005
         Left                     =   1380
         Top                      =   1860
         Width                    =   840
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo6
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   1950
         Top                      =   1875
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
         Left                     =   1395
         Width                    =   540
         Top                      =   1875
         DataField                =   "CifFob"
         Index                    =   6 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   2280
         Width                    =   1185
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
         Caption                  =   "*Fornecedor:"
         Index                    =   8 
      End
      Begin VB.Label labFdo7
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000016
         Left                     =   3555
         Top                      =   1860
         Width                    =   5250
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo7
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   8535
         Top                      =   1875
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000016
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
         TabIndex                 =   7
         TabStop                  =   -1
         Height                   =   270
         Left                     =   3570
         Width                    =   4950
         Top                      =   1875
         DataField                =   "Codigo do Fornecedor"
         Index                    =   7 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8835
         Width                    =   1425
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
         Caption                  =   "*Trasportadora:"
         Index                    =   9 
      End
      Begin VB.Label labFdo8
         Appearance               =  0
         BorderStyle              =  1
         BackColor                =  &H80000016
         Left                     =   10335
         Top                      =   1860
         Width                    =   4050
         Height                   =   300
      End
      Begin VB.CommandButton bottxtCampo8
         Style                    =   1
         TabStop                  =   0
         Index                    =   2
         Left                     =   14115
         Top                      =   1875
         Width                    =   255
         Height                   =   270
      End
      Begin VB.TextBox txtCp
         Appearance               =   0
         BackColor                =   &H80000016
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
         TabIndex                 =   8
         TabStop                  =   -1
         Height                   =   270
         Left                     =   10350
         Width                    =   3750
         Top                      =   1875
         DataField                =   "Codigo da Transportadora"
         Index                    =   8 
      End
      Begin TabDlg.SSTab Tabs
         TabsPerRow               =   8
         TabHeight                =   519
         BackColor                =   &H8000000F
         Style                    =   1
         TabOrientation           =   0
         Tabs                     =   8
         TabIndex                 =   13
         TabStop                  =   -1
         Height                   =   3930
         Left                     =   120
         Width                    =   14265
         Top                      =   2820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H800000
         Caption                  =   "Produtos"
         Index                    =   0 
         _version                 =   65536
         _extentx                 =   25149
         _extenty                 =   6928
         _stockprops              =   15
         TabCaption(0)            =   "Produtos"
         Tab(0).ControlCount=   1
         Tab(0).ControlEnabled=   -1
         Tab(0).Control(0)=   "Grid(0)"
         TabCaption(1)            =   "Despesas"
         Tab(1).ControlCount=   1
         Tab(1).ControlEnabled=   0
         Tab(1).Control(0)=   "Grid(1)"
         TabCaption(2)            =   "Endereço de Entrega"
         Tab(2).ControlCount=   18
         Tab(2).ControlEnabled=   0
         Tab(2).Control(0)=   "Label(17)"
         Tab(2).Control(1)=   "txtCp(16)"
         Tab(2).Control(2)=   "Label(18)"
         Tab(2).Control(3)=   "txtCp(17)"
         Tab(2).Control(4)=   "Label(19)"
         Tab(2).Control(5)=   "txtCp(18)"
         Tab(2).Control(6)=   "Label(20)"
         Tab(2).Control(7)=   "txtCp(19)"
         Tab(2).Control(8)=   "Label(21)"
         Tab(2).Control(9)=   "labFdo20"
         Tab(2).Control(10)=   "bottxtCampo20(2)"
         Tab(2).Control(11)=   "txtCp(20)"
         Tab(2).Control(12)=   "Label(22)"
         Tab(2).Control(13)=   "txtCp(21)"
         Tab(2).Control(14)=   "Label(23)"
         Tab(2).Control(15)=   "txtCp(22)"
         Tab(2).Control(16)=   "Label(24)"
         Tab(2).Control(17)=   "txtCp(23)"
         TabCaption(3)            =   "Financeiro"
         Tab(3).ControlCount=   40
         Tab(3).ControlEnabled=   0
         Tab(3).Control(0)=   "Label(10)"
         Tab(3).Control(1)=   "txtCp(9)"
         Tab(3).Control(2)=   "Label(11)"
         Tab(3).Control(3)=   "txtCp(10)"
         Tab(3).Control(4)=   "Label(12)"
         Tab(3).Control(5)=   "txtCp(11)"
         Tab(3).Control(6)=   "Label(13)"
         Tab(3).Control(7)=   "txtCp(12)"
         Tab(3).Control(8)=   "Label(14)"
         Tab(3).Control(9)=   "txtCp(13)"
         Tab(3).Control(10)=   "Label(15)"
         Tab(3).Control(11)=   "txtCp(14)"
         Tab(3).Control(12)=   "Label(16)"
         Tab(3).Control(13)=   "txtCp(15)"
         Tab(3).Control(14)=   "Label(26)"
         Tab(3).Control(15)=   "Label(27)"
         Tab(3).Control(16)=   "Formato(0)"
         Tab(3).Control(17)=   "Label(28)"
         Tab(3).Control(18)=   "Formato(1)"
         Tab(3).Control(19)=   "Label(29)"
         Tab(3).Control(20)=   "Formato(2)"
         Tab(3).Control(21)=   "Label(30)"
         Tab(3).Control(22)=   "Grid(6)"
         Tab(3).Control(23)=   "Label(31)"
         Tab(3).Control(24)=   "labFdo27"
         Tab(3).Control(25)=   "bottxtCampo27(2)"
         Tab(3).Control(26)=   "txtCp(27)"
         Tab(3).Control(27)=   "Botao(6)"
         Tab(3).Control(28)=   "Grid(7)"
         Tab(3).Control(29)=   "Label(32)"
         Tab(3).Control(30)=   "labFdo28"
         Tab(3).Control(31)=   "bottxtCampo28(1)"
         Tab(3).Control(32)=   "bottxtCampo28(2)"
         Tab(3).Control(33)=   "txtCp(28)"
         Tab(3).Control(34)=   "Label(33)"
         Tab(3).Control(35)=   "labFdo29"
         Tab(3).Control(36)=   "bottxtCampo29(1)"
         Tab(3).Control(37)=   "bottxtCampo29(2)"
         Tab(3).Control(38)=   "txtCp(29)"
         Tab(3).Control(39)=   "Botao(7)"
         TabCaption(4)            =   "Baixa Produtos"
         Tab(4).ControlCount=   7
         Tab(4).ControlEnabled=   0
         Tab(4).Control(0)=   "Grid(2)"
         Tab(4).Control(1)=   "Botao(0)"
         Tab(4).Control(2)=   "Botao(1)"
         Tab(4).Control(3)=   "labFdo24"
         Tab(4).Control(4)=   "labtxtCampo24"
         Tab(4).Control(5)=   "bottxtCampo24(0)"
         Tab(4).Control(6)=   "txtCp(24)"
         TabCaption(5)            =   "Baixa Despesas"
         Tab(5).ControlCount=   7
         Tab(5).ControlEnabled=   0
         Tab(5).Control(0)=   "Grid(3)"
         Tab(5).Control(1)=   "Botao(2)"
         Tab(5).Control(2)=   "Botao(3)"
         Tab(5).Control(3)=   "labFdo25"
         Tab(5).Control(4)=   "labtxtCampo25"
         Tab(5).Control(5)=   "bottxtCampo25(0)"
         Tab(5).Control(6)=   "txtCp(25)"
         TabCaption(6)            =   "Mat. Consumo"
         Tab(6).ControlCount=   1
         Tab(6).ControlEnabled=   0
         Tab(6).Control(0)=   "Grid(4)"
         TabCaption(7)            =   "Baixa Consumo"
         Tab(7).ControlCount=   7
         Tab(7).ControlEnabled=   0
         Tab(7).Control(0)=   "Grid(5)"
         Tab(7).Control(1)=   "Botao(4)"
         Tab(7).Control(2)=   "Botao(5)"
         Tab(7).Control(3)=   "Label(25)"
         Tab(7).Control(4)=   "labFdo26"
         Tab(7).Control(5)=   "bottxtCampo26(0)"
         Tab(7).Control(6)=   "txtCp(26)"
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -74865
            Width                    =   1830
            Top                      =   480
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
            Caption                  =   "Valor do Frete...........:"
            Index                    =   10 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   15
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72900
            Width                    =   1365
            Top                      =   480
            DataField                =   "Vr do Frete"
            Index                    =   9 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -74865
            Width                    =   1845
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
            Caption                  =   "Valor do Desconto..:"
            Index                    =   11 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            TabIndex                 =   16
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -72900
            Width                    =   1365
            Top                      =   840
            DataField                =   "Vr do Desconto"
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
            Left                     =   -74865
            Width                    =   1830
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
            Caption                  =   "Valor Total do IPI.....:"
            Index                    =   12 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
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
            Left                     =   -72900
            Width                    =   1365
            Top                      =   1200
            DataField                =   "Total do IPI"
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
            Left                     =   -74940
            Width                    =   1905
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
            Caption                  =   "Valor Total do ICMS:"
            Index                    =   13 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   18
            TabStop                  =   0
            Height                   =   300
            Left                     =   -72900
            Width                    =   1365
            Top                      =   1560
            DataField                =   "Total do ICMS"
            TabStop                  =  0
            Index                    =   12 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -71220
            Width                    =   2445
            Top                      =   480
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
            Caption                  =   "Valor Total dos Produtos...:"
            Index                    =   14 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
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
            Left                     =   -68640
            Width                    =   1470
            Top                      =   480
            DataField                =   "Total dos Produtos"
            TabStop                  =  0
            Index                    =   13 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -71220
            Width                    =   2445
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
            Caption                  =   "Valor Total das Despesas:"
            Index                    =   15 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   20
            TabStop                  =   0
            Height                   =   300
            Left                     =   -68640
            Width                    =   1470
            Top                      =   840
            DataField                =   "Total das Despesas"
            TabStop                  =  0
            Index                    =   14 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -71220
            Width                    =   2445
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
            Caption                  =   "Valor Total do Pedido.:"
            Index                    =   16 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   1
            Locked                   =   -1
            TabIndex                 =   21
            TabStop                  =   0
            Height                   =   300
            Left                     =   -68640
            Width                    =   1470
            Top                      =   1200
            DataField                =   "Total do Pedido"
            TabStop                  =  0
            Index                    =   15 
            MultiLine                =   -1
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -74880
            Width                    =   960
            Top                      =   960
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
            Caption                  =   "Endereço:"
            Index                    =   17 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   50
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   22
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -73800
            Width                    =   5580
            Top                      =   960
            DataField                =   "Endereco de Entrega"
            Index                    =   16 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -68160
            Width                    =   270
            Top                      =   960
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
            Caption                  =   "Nº"
            Index                    =   18 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   9
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   23
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -67860
            Width                    =   1125
            Top                      =   960
            DataField                =   "Numero do Endereco"
            Index                    =   17 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -66600
            Width                    =   630
            Top                      =   960
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
            Caption                  =   "Bairro:"
            Index                    =   19 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   50
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   24
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -65940
            Width                    =   5025
            Top                      =   960
            DataField                =   "Bairro de Entrega"
            Index                    =   18 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -74880
            Width                    =   990
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
            Caption                  =   "Cidade.....:"
            Index                    =   20 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   50
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   25
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -73800
            Width                    =   3600
            Top                      =   1740
            DataField                =   "Cidade de Entrega"
            Index                    =   19 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -70080
            Width                    =   390
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
            Caption                  =   "UF:"
            Index                    =   21 
         End
         Begin VB.Label labFdo20
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -69660
            Top                      =   1740
            Width                    =   660
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo20
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -69270
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
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   0
            Alignment                =   0
            TabIndex                 =   28
            TabStop                  =   -1
            Height                   =   270
            Left                     =   -69645
            Width                    =   375
            Top                      =   1755
            DataField                =   "UF De Entrega"
            Index                    =   20 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -68880
            Width                    =   945
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
            Caption                  =   "CEP..........:"
            Index                    =   22 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   29
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -67860
            Width                    =   1125
            Top                      =   1740
            DataField                =   "CEP de Entrega"
            Index                    =   21 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -74880
            Width                    =   1065
            Top                      =   2580
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
            Caption                  =   "Telefone..:"
            Index                    =   23 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   20
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   30
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -73800
            Width                    =   1605
            Top                      =   2580
            DataField                =   "Fone de Entrega"
            Index                    =   22 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   0
            Height                   =   240
            Left                     =   -72060
            Width                    =   825
            Top                      =   2580
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
            Caption                  =   "Contato:"
            Index                    =   24 
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000005
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   45
            BorderStyle              =   1
            Alignment                =   0
            TabIndex                 =   31
            TabStop                  =   -1
            Height                   =   300
            Left                     =   -71220
            Width                    =   4485
            Top                      =   2580
            DataField                =   "Contato de Entrega"
            Index                    =   23 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   32
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2985
            Left                     =   120
            Width                    =   14085
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   0 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   33
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2970
            Left                     =   -74880
            Width                    =   13995
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   1 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   36
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2895
            Left                     =   -74865
            Width                    =   14010
            Top                      =   405
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   2 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   37
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -74820
            Width                    =   1605
            Top                      =   3360
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Parcial"
            Index                    =   0 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   38
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -73140
            Width                    =   1605
            Top                      =   3360
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Total"
            Index                    =   1 
         End
         Begin VB.Label labFdo24
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000016
            Left                     =   -70630
            Top                      =   3480
            Width                    =   1270
            Height                   =   270
         End
         Begin VB.Label labtxtCampo24
            Caption                  =   "Nº Nota"
            BackStyle                =   0
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
            Left                     =   -71385
            Top                      =   3480
            Width                    =   705
            Height                   =   240
         End
         Begin VB.CommandButton bottxtCampo24
            Style                    =   1
            TabStop                  =   0
            Index                    =   0
            Left                     =   -69630
            Top                      =   3495
            Width                    =   255
            Height                   =   240
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   254
            BorderStyle              =   0
            Alignment                =   0
            TabIndex                 =   39
            TabStop                  =   -1
            Height                   =   240
            Left                     =   -70615
            Width                    =   985
            Top                      =   3495
            Index                    =   24 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   40
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2490
            Left                     =   -74880
            Width                    =   14010
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   3 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   41
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -74820
            Width                    =   1605
            Top                      =   3300
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Parcial"
            Index                    =   2 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   42
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -73140
            Width                    =   1605
            Top                      =   3300
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Total"
            Index                    =   3 
         End
         Begin VB.Label labFdo25
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000016
            Left                     =   -70630
            Top                      =   3420
            Width                    =   1330
            Height                   =   270
         End
         Begin VB.Label labtxtCampo25
            Caption                  =   "Nº Nota"
            BackStyle                =   0
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
            Left                     =   -71385
            Top                      =   3420
            Width                    =   705
            Height                   =   240
         End
         Begin VB.CommandButton bottxtCampo25
            Style                    =   1
            TabStop                  =   0
            Index                    =   0
            Left                     =   -69570
            Top                      =   3435
            Width                    =   255
            Height                   =   240
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   400
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000008
            MaxLength                =   254
            BorderStyle              =   0
            Alignment                =   0
            TabIndex                 =   43
            TabStop                  =   -1
            Height                   =   240
            Left                     =   -70615
            Width                    =   1045
            Top                      =   3435
            Index                    =   25 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   44
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2760
            Left                     =   -74880
            Width                    =   13980
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   4 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   45
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   2475
            Left                     =   -74880
            Width                    =   13995
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   5 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   46
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -74820
            Width                    =   1620
            Top                      =   3300
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Parcial"
            Index                    =   4 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   47
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   480
            Left                     =   -73080
            Width                    =   1650
            Top                      =   3300
            CaptionAlign             =   0
            CaptionOffset            =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
               Name                     =   "Microsoft Sans Serif"
               CharSet                  =   0
               Weight                   =   700
               Size                     =   9
               Underline                =   0
               Italic                   =   0
               StrikeThrough            =   0
            EndProperty
            ForeColor                =   &H80000012
            Caption                  =   "Baixa Total"
            Index                    =   5 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -71220
            Width                    =   705
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
            Caption                  =   "Nº Nota"
            Index                    =   25 
         End
         Begin VB.Label labFdo26
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000016
            Left                     =   -70380
            Top                      =   3420
            Width                    =   1230
            Height                   =   270
         End
         Begin VB.CommandButton bottxtCampo26
            Style                    =   1
            TabStop                  =   0
            Index                    =   0
            Left                     =   -69420
            Top                      =   3435
            Width                    =   255
            Height                   =   240
         End
         Begin VB.TextBox txtCp
            Appearance               =   0
            BackColor                =   &H80000016
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
            TabIndex                 =   48
            TabStop                  =   -1
            Height                   =   240
            Left                     =   -70365
            Width                    =   945
            Top                      =   3435
            Index                    =   26 
         End
         Begin VB.Label Label
            Appearance               =   0
            BackColor                =   &H0
            BackStyle                =   1
            BorderStyle              =   1
            AutoSize                 =   0
            Height                   =   255
            Left                     =   -65805
            Width                    =   4830
            Top                      =   435
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
            Index                    =   26 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65685
            Width                    =   600
            Top                      =   3600
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
            Index                    =   27 
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
            Left                     =   -65010
            Width                    =   165
            Top                      =   3645
            Index                    =   0 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -64785
            Width                    =   990
            Top                      =   3600
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
            Index                    =   28 
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
            Left                     =   -63630
            Width                    =   165
            Top                      =   3645
            Index                    =   1 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -63390
            Width                    =   1050
            Top                      =   3600
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
            Index                    =   29 
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
            Left                     =   -62175
            Width                    =   150
            Top                      =   3645
            Index                    =   2 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -61950
            Width                    =   300
            Top                      =   3600
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
            Index                    =   30 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   49
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   1560
            Left                     =   -65805
            Width                    =   4800
            Top                      =   720
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   6 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65745
            Width                    =   540
            Top                      =   3135
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
            Caption                  =   "*Tipo:"
            Index                    =   31 
         End
         Begin VB.Label labFdo27
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -65115
            Top                      =   3135
            Width                    =   1620
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo27
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -63765
            Top                      =   3150
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
            TabIndex                 =   50
            TabStop                  =   0
            Height                   =   270
            Left                     =   -65100
            Width                    =   1335
            Top                      =   3150
            Index                    =   27 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   51
            TabStop                  =   -1
            ButtonType               =   0
            Height                   =   300
            Left                     =   -63405
            Width                    =   1350
            Top                      =   3135
            CaptionAlign             =   0
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
            Caption                  =   "Gerar Conta"
            Index                    =   6 
         End
         Begin IRRIG.GListV Grid
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
            TabIndex                 =   53
            TabStop                  =   -1
            Stripes                  =   0
            SaveGridBars             =   0
            Height                   =   1800
            Left                     =   -74880
            Width                    =   8760
            Top                      =   2025
            Caption                  =   ""
            ShowGridCaption          =   0
            Index                    =   7 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65805
            Width                    =   600
            Top                      =   2445
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
            Caption                  =   "Grupo:"
            Index                    =   32 
         End
         Begin VB.Label labFdo28
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -65115
            Top                      =   2415
            Width                    =   4125
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo28
            Style                    =   1
            TabStop                  =   0
            Index                    =   1
            Left                     =   -61260
            Top                      =   2430
            Width                    =   255
            Height                   =   270
         End
         Begin VB.CommandButton bottxtCampo28
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -61515
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
            TabIndex                 =   55
            TabStop                  =   0
            Height                   =   270
            Left                     =   -65100
            Width                    =   3840
            Top                      =   2430
            DataField                =   "Seqüência Grupo Despesa"
            Index                    =   28 
         End
         Begin VB.Label Label
            Appearance               =   1
            BackColor                =   &H8000000F
            BackStyle                =   1
            BorderStyle              =   0
            AutoSize                 =   -1
            Height                   =   240
            Left                     =   -65610
            Width                    =   405
            Top                      =   2805
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
            Caption                  =   "Sub:"
            Index                    =   33 
         End
         Begin VB.Label labFdo29
            Appearance               =  0
            BorderStyle              =  1
            BackColor                =  &H80000005
            Left                     =   -65115
            Top                      =   2775
            Width                    =   4110
            Height                   =   300
         End
         Begin VB.CommandButton bottxtCampo29
            Style                    =   1
            TabStop                  =   0
            Index                    =   1
            Left                     =   -61275
            Top                      =   2790
            Width                    =   255
            Height                   =   270
         End
         Begin VB.CommandButton bottxtCampo29
            Style                    =   1
            TabStop                  =   0
            Index                    =   2
            Left                     =   -61530
            Top                      =   2790
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
            TabIndex                 =   56
            TabStop                  =   0
            Height                   =   270
            Left                     =   -65100
            Width                    =   3825
            Top                      =   2790
            DataField                =   "Seqüência SubGrupo Despesa"
            Index                    =   29 
         End
         Begin IRRIG.GBotao Botao
            ButtonStyle              =   0
            JoinSeparators           =   0
            LeftSeparator            =   0
            TopSeparator             =   0
            TabIndex                 =   59
            TabStop                  =   0
            ButtonType               =   0
            Height                   =   300
            Left                     =   -61995
            Width                    =   990
            Top                      =   3135
            CaptionAlign             =   0
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
            Caption                  =   "Estornar"
            Index                    =   7 
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
         Width                    =   420
         Top                      =   6735
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
         ForeColor                =   &H0
         Caption                  =   "Obs"
         Index                    =   34 
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
         TabIndex                 =   14
         TabStop                  =   -1
         Height                   =   750
         Left                     =   120
         Width                    =   14280
         Top                      =   7035
         DataField                =   "Obs"
         Index                    =   30 
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
         Left                     =   7740
         Width                    =   2745
         Top                      =   60
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
         ForeColor                =   &H80
         Caption                  =   "Previsão de Entrega  --->>"
         Index                    =   35 
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
         TabIndex                 =   27
         TabStop                  =   0
         Height                   =   300
         Left                     =   11865
         Width                    =   1185
         Top                      =   60
         DataField                =   "Previsao de Entrega"
         Index                    =   31 
      End
      Begin IRRIG.GBotao Botao
         ButtonStyle              =   0
         JoinSeparators           =   0
         LeftSeparator            =   0
         TopSeparator             =   0
         TabIndex                 =   34
         TabStop                  =   0
         ButtonType               =   0
         Height                   =   480
         Left                     =   12840
         Width                    =   1515
         Top                      =   780
         CaptionAlign             =   4
         CaptionOffset            =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
            Name                     =   "Microsoft Sans Serif"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   11.25
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &H80000012
         Caption                  =   "Imprimir"
         Index                    =   8 
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
         ForeColor                =   &H0
         Alignment                =   0
         TabIndex                 =   35
         TabStop                  =   -1
         Height                   =   255
         Left                     =   3540
         Width                    =   1470
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
         ForeColor                =   &H0
         Alignment                =   0
         Caption                  =   "Fechado"
         Index                    =   0 
         DataField                =   "Pedido Fechado"
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   270
         Left                     =   1920
         Width                    =   2250
         Top                      =   900
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
         ForeColor                =   &HFF
         Caption                  =   "Pedido Atrasado"
         Index                    =   36 
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
         ForeColor                =   &H0
         Alignment                =   0
         TabIndex                 =   52
         TabStop                  =   -1
         Height                   =   255
         Left                     =   1920
         Width                    =   1440
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
         ForeColor                =   &H0
         Alignment                =   0
         Caption                  =   "Cancelado"
         Index                    =   1 
         DataField                =   "Cancelado"
      End
      Begin VB.Label Label
         Appearance               =   0
         BackColor                =   &HFFFFFF
         BackStyle                =   0
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   600
         Width                    =   13380
         Top                      =   6735
         Alignment                =   0
         BeginProperty Font
            Name                     =   "Arial"
            CharSet                  =   0
            Weight                   =   700
            Size                     =   9.75
            Underline                =   0
            Italic                   =   0
            StrikeThrough            =   0
         EndProperty
         ForeColor                =   &HFF
         Caption                  =   "Título"
         Index                    =   37 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   2295
         Width                    =   1230
         Top                      =   2310
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
         Caption                  =   "*Banco..........:"
         Index                    =   38 
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
         MaxLength                =   20
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   9
         TabStop                  =   -1
         Height                   =   300
         Left                     =   3555
         Width                    =   1500
         Top                      =   2295
         DataField                =   "Nome do Banco 1"
         Index                    =   32 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   5175
         Width                    =   945
         Top                      =   2310
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
         Caption                  =   "*Agência:"
         Index                    =   39 
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
         MaxLength                =   20
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   10
         TabStop                  =   -1
         Height                   =   300
         Left                     =   6180
         Width                    =   780
         Top                      =   2295
         DataField                =   "Agência do Banco 1"
         Index                    =   33 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   7095
         Width                    =   420
         Top                      =   2310
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
         Caption                  =   "*CC:"
         Index                    =   40 
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
         MaxLength                =   15
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   11
         TabStop                  =   -1
         Height                   =   300
         Left                     =   7650
         Width                    =   1155
         Top                      =   2295
         DataField                =   "Conta Corrente do Banco 1"
         Index                    =   34 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   8940
         Width                    =   1320
         Top                      =   2310
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
         Caption                  =   "CNPJ - Titular:"
         Index                    =   41 
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
         TabIndex                 =   12
         TabStop                  =   -1
         Height                   =   300
         Left                     =   10335
         Width                    =   4050
         Top                      =   2295
         DataField                =   "Nome do Correntista do Banco 1"
         Index                    =   35 
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
         TabIndex                 =   54
         TabStop                  =   -1
         Height                   =   255
         Left                     =   5040
         Width                    =   1530
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
         ForeColor                =   &H80000008
         Alignment                =   0
         Caption                  =   "Pré Pedido"
         Index                    =   2 
         DataField                =   "Prepedido"
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
         TabIndex                 =   26
         TabStop                  =   -1
         Height                   =   300
         Left                     =   11220
         Width                    =   570
         Top                      =   60
         DataField                =   "Dias"
         Index                    =   36 
         MultiLine                =   -1
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   0
         Height                   =   240
         Left                     =   10635
         Width                    =   465
         Top                      =   75
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
         Caption                  =   "Dias:"
         Index                    =   42 
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
         MaxLength                =   150
         BorderStyle              =   1
         Alignment                =   0
         TabIndex                 =   57
         TabStop                  =   -1
         Height                   =   300
         Left                     =   7785
         Width                    =   6585
         Top                      =   420
         DataField                =   "Justificar o Atraso"
         Index                    =   37 
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
         TabIndex                 =   58
         TabStop                  =   -1
         Height                   =   300
         Left                     =   13185
         Width                    =   1185
         Top                      =   60
         DataField                =   "Nova Previsao"
         Index                    =   38 
      End
      Begin VB.Label Label
         Appearance               =   1
         BackColor                =   &H8000000F
         BackStyle                =   1
         BorderStyle              =   0
         AutoSize                 =   -1
         Height                   =   195
         Left                     =   14520
         Width                    =   45
         Top                      =   7920
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
         Index                    =   43 
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
Attribute VB_name = "frmPComprN"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'* Sistema...: SISTEMA IRRIGAÇÃO PENÁPOLIS
'* Empresa...: IRRIGAÇÃO PENÁPOLIS
'* Módulo....: PComprN
'* Função....: Pedido
'* CopyRight.: (C)2025 IRRIGAÇÃO PENÁPOLIS
'* Criação...: GAS-2007 - Gerador Automático de Sistemas
'* Data......: 30/05/2025 16:33:25
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
Dim txtCampo(38) As New FormataCampos             'classe dos campos tipo texto do módulo  
Dim chkCampo(2) As New FormataCampos              'classe dos campos tipo lógico do módulo
Dim Id_do_Pedido As Long, Data_do_Pedido As Variant, Nro_da_Licitacao As String
Dim Codigo_do_Fornecedor As Long, Codigo_da_Transportadora As Long, Comprador As String
Dim Vend As String, Prazo As String, CifFob As String
Dim Vr_do_Frete As Double, Vr_do_Desconto As Double, Obs As String
Dim Total_dos_Produtos As Double, Total_das_Despesas As Double, Total_do_IPI As Double
Dim Total_do_ICMS As Double, Total_do_Pedido As Double, Endereco_de_Entrega As String
Dim Numero_do_Endereco As String, Bairro_de_Entrega As String, Cidade_de_Entrega As String
Dim UF_De_Entrega As String, CEP_de_Entrega As String, Fone_de_Entrega As String
Dim Contato_de_Entrega As String, Previsao_de_Entrega As Variant, Pedido_Fechado As Boolean
Dim Validado As Boolean, Cancelado As Boolean, Nome_do_Banco_1 As String
Dim Agencia_do_Banco_1 As String, Conta_Corrente_do_Banco_1 As String, Nome_do_Correntista_do_Banco_1 As String
Dim Prepedido As Boolean, Sequencia_Grupo_Despesa As Integer, Sequencia_SubGrupo_Despesa As Integer
Dim Codigo_da_Licitacao As Long, Justificar_o_Atraso As String, Nova_Previsao As Variant
Dim Dias As Integer, NotaFiscal As String, NotaDespesa As String
Dim NotaConsumo As String, Tipo_da_Licitacao As String
Public lblPedido As Object, Aba1 As Object, GridProdutos As Object
Public GridDespesas As Object, GridBxProduto As Object, Txtnotaprod As Object
Public GridBxDespesa As Object, Txtnotadespesa As Object, GridBxConsumo As Object
Public lblParcelamento As Object, GridParcelas As Object, txtTipo_da_Licitacao As Object
Public btnGerarconta As Object, btnExtornar As Object, Txtprevisao As Object
Public Lblfin As Object, Txtdias As Object, lblAjuste As Object
Dim Pedido_de_Compra_Novo As New GRecordSet, Produtos_do_Pedido_Compra As New GRecordSet, Despesas_do_Pedido_Compra As New GRecordSet
Dim Bx_Produtos_Pedido_Compra As New GRecordSet, Bx_Despesas_Pedido_Compra As New GRecordSet, Consumo_do_Pedido_Compra As New GRecordSet
Dim Bx_Consumo_Pedido_Compra As New GRecordSet, Parcelas_Ped_Compra_Novo As New GRecordSet, Orcamentos_da_compra As New GRecordSet

Dim ProdutoAux As New GRecordSet, DespesaAux As New GRecordSet
Dim ConsumoAux As New GRecordSet, OrcAux As New GRecordSet, ClienteAux As New GRecordSet

Private Sub AjustaValores()
   Dim IPIProdutos As New GRecordSet, IPIConsumo As New GRecordSet, TotalProdutos As New GRecordSet 
   Dim ValorMvto As Currency, ICMSProdutos As New GRecordSet, TotalConsumo As New GRecordSet
   Dim IPIDespesas As New GRecordSet, TotalDespesas As New GRecordSet
   Dim ICMSDespesas As New GRecordSet, ICMSConsumo As New GRecordSet
   Dim ProdAux As New GRecordSet
  
   On Error GoTo DeuErro
   
   Set IPIConsumo =  vgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do IPI]), 2) Total FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'IPI Consumo
   Set IPIProdutos = vgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do IPI]), 2) Total FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'IPI dos Produtos
   Set IPIDespesas = vgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do IPI]), 2) Total FROM [Despesas do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'IPI das Despesas
   Set TotalProdutos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr Total]), 2) Total FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Total dos Produtos
   Set TotalDespesas = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr Total]), 2) Total FROM [Despesas do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Total das Despesas
   Set TotalConsumo = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr Total]), 2) Total FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Total Consumo
   Set ICMSProdutos = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do Icms]), 2) Total FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Icms dos Produtos
   Set ICMSDespesas = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do Icms]), 2) Total FROM [Despesas do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Icms das Despesas
   Set ICMSConsumo = VgDb.OpenRecordSet("SELECT ROUND(SUM([Vr do Icms]), 2) Total FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido) 'Icms Consumo mo Gambiarra
   
   If CifFob = "FOB" Then                                                                                             
      ValorMvto = IPIProdutos!Total + IPIDespesas!Total + IPIConsumo!Total + TotalDespesas!Total + TotalProdutos!Total + TotalConsumo!Total + Vr_do_Frete - Vr_do_Desconto 
      ValorMvto = Format(ValorMvto , "##,###,##0.00")
   End If
   
   If CifFob = "CIF" Then
      ValorMvto = IPIConsumo!Total + TotalConsumo!Total + IPIProdutos!Total + TotalProdutos!Total + IPIDespesas!Total + TotalDespesas!Total - Vr_do_Desconto 
      ValorMvto = Format(ValorMvto , "##,###,##0.00")
   End If
   
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Pedido de Compra Novo] Set [Total do IPI] = " & Substitui(IPIProdutos!Total + IPIDespesas!Total + IPIConsumo!Total, ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido 'IPI Produtos e Despesas
   vgDb.Execute "Update [Pedido de Compra Novo] Set [Total das Despesas] = " & Substitui(TotalDespesas!Total + TotalConsumo!Total, ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido 'Total das Despesas
   vgDb.Execute "Update [Pedido de Compra Novo] Set [Total dos Produtos] = " & Substitui(TotalProdutos!Total, ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido 'Total dos Produtos
   vgDb.Execute "Update [Pedido de Compra Novo] Set [Total do ICMS] = " & Substitui(ICMSProdutos!Total + ICMSDespesas!Total + ICMSConsumo!Total, ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido 'ICMS dos Produtos
   vgDb.Execute "Update [Pedido de Compra Novo] Set [Total do Pedido] = " & Substitui(CStr(ValorMvto), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido 'Total do Pedido
   vgDb.CommitTrans
      
   
       Set ProdAux = VgDb.OpenRecordSet("SELECT * FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
       Do While Not ProdAux.EOF
          ContaEstoque(ProdAux![Id Do Produto]) 
          ProdAux.MoveNext
       Loop    
     
        
DeuErro:
   If Err Then
      MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
   txtTipo_da_Licitacao.Text = ""
End Sub


Private Function AjustaValoresProdutos(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, _
   Sequencia_do_Item As Long, Sequencia_do_Orcamento As Long)
  Dim ValorIPI As Currency, ValorTotalItem As Currency, ValorIcmsItem As Currency
  
  On Error GoTo DeuErro
  
  'If Pedido_de_Compra_Novo![Codigo Do Fornecedor] <> 4957 Then
  '   ValorIPI = Round((((Qtde * Vr_Unitario) *  Aliquota_do_IPI) / 100),2)
  'Else
  '   ValorIPI = Vr_do_IPI   
  'End If 
  
  ValorIPI = Round((((Qtde * Vr_Unitario) *  Aliquota_do_IPI) / 100),2)
  ValorTotalItem = Round((Qtde * Vr_Unitario),2)
  ValorIcmsItem = Round((((Vr_Unitario * Qtde + ((Qtde * Vr_Unitario) * Aliquota_do_IPI) / 100 ) * Aliquota_do_Icms ) / 100),2)
  
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Produtos do Pedido Compra] Set [Vr do IPI] = " &  Substitui(CStr(ValorIPI), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id do Produto] = " & Id_do_Produto & " And [Sequencia do Item] = " & Sequencia_do_Item
   vgDb.Execute "Update [Produtos do Pedido Compra] Set [Vr Total] = " &  Substitui(CStr(ValorTotalItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id do Produto] = " & Id_do_Produto & " And [Sequencia do Item] = " & Sequencia_do_Item 
   vgDb.Execute "Update [Produtos do Pedido Compra] Set [Vr do Icms] = " & Substitui(CStr(ValorIcmsItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id do Produto] = " & Id_do_Produto & " And [Sequencia do Item] = " & Sequencia_do_Item
   vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 0 Where [Id do Pedido] = " & Id_do_Pedido
   vgDb.CommitTrans
   
   AjustaValores ' Atualiza o Pedido
   GridProdutos.Rebind
  
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If
End Function


Private Function AjustaDespesas(Id_do_Pedido As Long, Id_da_Despesa As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double) 
 Dim ValorIPI As Currency, ValorTotalItem As Currency, ValorIcmsItem As Currency
  
  On Error GoTo DeuErro
  
  ValorIPI = Round((((Qtde * Vr_Unitario) *  Aliquota_do_IPI) / 100),2)
  ValorTotalItem = Round((Qtde * Vr_Unitario),2)
  ValorIcmsItem = Round((((Vr_Unitario * Qtde + ((Qtde * Vr_Unitario) * Aliquota_do_IPI) / 100 ) * Aliquota_do_Icms ) / 100),2)
  
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Despesas do Pedido Compra] Set [Vr do IPI] = " &  Substitui(CStr(ValorIPI), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id da Despesa] = " & Id_da_Despesa 
   vgDb.Execute "Update [Despesas do Pedido Compra] Set [Vr Total] = " &  Substitui(CStr(ValorTotalItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id da Despesa] = " & Id_da_Despesa 
   vgDb.Execute "Update [Despesas do Pedido Compra] Set [Vr do Icms] = " & Substitui(CStr(ValorIcmsItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id da Despesa] = " & Id_da_Despesa 
   vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 0 Where [Id do Pedido] = " & Id_do_Pedido
   vgDb.CommitTrans
   
   AjustaValores ' Atualiza o Pedido
  
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If

End Function


Public Sub CarregaTotalizador()
   On Error Resume Next

   With gridProdutos
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor do IPI").Index) = True
      .ShowSumCol(.Columns("Valor do ICMS").Index) = True 
      .ShowSumCol(.Columns("Qtde").Index) = False
      .ShowSumCol(.Columns("Vr Unitario").Index) = False  
      .ShowSumCol(.Columns("Aliquota do IPI").Index) = False 
      .ShowSumCol(.Columns("Aliquota do ICMS").Index) = False
      .ShowFilterBar = False
      .HideStatus = True
   End With
   With gridDespesas
      .ShowSumBar = True
      .ShowSumCol(.Columns("Valor Total").Index) = True
      .ShowSumCol(.Columns("Valor do IPI").Index) = True
      .ShowSumCol(.Columns("Valor do ICMS").Index) = True 
      .ShowSumCol(.Columns("Qtde").Index) = False 
      .ShowSumCol(.Columns("Vr Unitario").Index) = False 
      .ShowSumCol(.Columns("Aliquota do IPI").Index) = False
      .ShowSumCol(.Columns("Aliquota do ICMS").Index) = False        
      .ShowFilterBar = False
      .HideStatus = True
   End With 
   BloqueiaAba
End Sub

Private Sub LancaBxProdutoParcial()
   Dim Produtos As New GRecordSet
   Dim BxProduto As New GRecordSet
   Dim SeqItem As New GRecordSet 
   
   If Produtos_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Produtos_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Produtos_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Produtos = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxProduto = VgDb.OpenRecordSet("Bx Produtos Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Produtos para Realizar Baixa
   Do While Not Produtos.EOF
      With BxProduto
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Do Produto] = Produtos![Id Do Produto]
         
          'Set SeqItem = VgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Seq From [Bx Produtos Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
         '![Sequencia Do Item] = SeqItem!Seq + 1
           If Produtos![Sequencia Do Item] = 0 Then
              Set SeqItem = VgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Seq From [Bx Produtos Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
              ![Sequencia Do Item] = SeqItem!Seq + 1
           Else
              ![Sequencia Do Item] = Produtos![Sequencia Do Item] 
           End If
         ![Qtde Total] = Produtos!Qtde
         ![Vr Unitario] = Produtos![Vr Unitario]+(Produtos![Vr Do Ipi] / Produtos!Qtde)
         ![Vr Total Do Pedido] = Produtos![Vr Total] + Produtos![Vr Do IPI]
         ![Qtde Restante] = Produtos!Qtde
         ![Qtde Recebida] = 0
         ![Notas] = NotaFiscal
         .Update
         .BookMark = .LastModified
      End With
      ContaEstoque(Produtos![Id Do Produto])
      Produtos.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Produtos Incluidos no Grid com Sucesso Informe a Quantidade Recebida de Cada Produto para Efetuar as Baixas") 
   GridBxProduto.Rebind
   TxtnotaProd.Text = ""
End Sub


Private Function AjustaBxProdutos(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde_Total As Double, _
   Vr_Unitario As Double, Vr_Total_do_Pedido As Double, Qtde_Recebida As Double, _
   Qtde_Restante As Double, Total_Restante As Double, Notas As String, _
   Sequencia_do_Item As Long, Teste As String)
 Dim QtdeRestante As Double, DinheiroRestante As Currency
  
  On Error GoTo DeuErro
  
   QtdeRestante = Qtde_Total - Qtde_Recebida
   DinheiroRestante = Round((Qtde_Total * Vr_Unitario ) - (Qtde_Recebida * Vr_Unitario),2)
   
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Bx Produtos Pedido Compra] Set [Qtde Restante] = " &  Substitui(CStr(QtdeRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id do Produto] = " & Id_do_Produto
   vgDb.Execute "Update [Bx Produtos Pedido Compra] Set [Total Restante] = " &  Substitui(CStr(DinheiroRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id do Produto] = " & Id_do_Produto  
   vgDb.CommitTrans
   
   ContaEstoque(Id_do_Produto)
   
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If
End Function

Private Sub LancaBxProdutoTotal()
   Dim Produtos As New GRecordSet
   Dim BxProduto As New GRecordSet
   Dim SeqItem As New GRecordSet
   
   If Produtos_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Produtos_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Produtos_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Produtos = vgDb.OpenRecordSet("SELECT * FROM [Produtos do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxProduto = VgDb.OpenRecordSet("Bx Produtos Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Produtos para Realizar Baixa
   Do While Not Produtos.EOF
      With BxProduto
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Do Produto] = Produtos![Id Do Produto]
           If Produtos![Sequencia Do Item] = 0 Then
              Set SeqItem = VgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Seq From [Bx Produtos Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
              ![Sequencia Do Item] = SeqItem!Seq + 1
           Else
              ![Sequencia Do Item] = Produtos![Sequencia Do Item] 
           End If    
         ![Qtde Total] = Produtos!Qtde
         ![Vr Unitario] = Produtos![Vr Unitario]+(Produtos![Vr Do Ipi] / Produtos!Qtde)
         ![Vr Total Do Pedido] = Produtos![Vr Total] + Produtos![Vr Do IPI]
         ![Qtde Recebida] = Produtos!Qtde
         ![Qtde Restante] = 0
         ![Total Restante] = 0
         ![Notas] = NotaFiscal
         .Update
         .BookMark = .LastModified
      End With
      ContaEstoque(Produtos![Id Do Produto])
      Produtos.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Baixa Realizada Com Sucesso") 
   GridBxProduto.Rebind
   TxtnotaProd.Text = ""
End Sub

Private Sub RepositionPedido()
 
 If vgSituacao = ACAO_INCLUINDO Then
    Set Produtos_do_Pedido_Compra = Nothing
    Set Bx_Produtos_Pedido_Compra = Nothing
    Set Despesas_do_Pedido_Compra = Nothing 
    Set Bx_Despesas_Pedido_Compra = Nothing 
    Set Consumo_do_Pedido_Compra = Nothing
    Set Bx_Consumo_Pedido_Compra = Nothing 
    Set Parcelas_Ped_Compra_Novo = Nothing   
 End If   

End Sub

Private Sub LancaBxDespesaParcial()
   Dim Despesa As New GRecordSet
   Dim BxDespesa As New GRecordSet
   
   If Despesas_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Despesas_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Despesas_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Despesa = vgDb.OpenRecordSet("SELECT * FROM [Despesas do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxDespesa = VgDb.OpenRecordSet("Bx Despesas Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Despesas para Realizar Baixa
   Do While Not Despesa.EOF
      With BxDespesa
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Da Despesa] = Despesa![Id Da Despesa]
         ![Qtde Total] = Despesa!Qtde
         ![Vr Unitario]= Despesa![Vr Unitario]+(Despesa![Vr Do Ipi] / Despesa!Qtde)
         ![Vr Total Do Pedido] = Despesa![Vr Total] + Despesa![Vr Do Ipi]
         ![Qtde Recebida] = 0
         ![Qtde Restante] = Despesa!Qtde
         ![Notas] = NotaDespesa
         .Update
         .BookMark = .LastModified
      End With
      Despesa.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Despesas Incluidas no Grid com Sucesso Informe a Quantidade Recebida de Cada Despesa para Efetuar as Baixas") 
   GridBxDespesa.Rebind
   TxtnotaDespesa.Text = ""
End Sub

Private Sub LancaBxDespesaTotal()
   Dim Despesa As New GRecordSet
   Dim BxDespesa As New GRecordSet
   
   If Despesas_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Despesas_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Despesas_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Despesa = vgDb.OpenRecordSet("SELECT * FROM [Despesas do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxDespesa = VgDb.OpenRecordSet("Bx Despesas Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Despesas para Realizar Baixa
   Do While Not Despesa.EOF
      With BxDespesa
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Da Despesa] = Despesa![Id Da Despesa]
         ![Qtde Total] = Despesa!Qtde
         ![Vr Unitario]= Despesa![Vr Unitario]+(Despesa![Vr Do Ipi] / Despesa!Qtde)
         ![Vr Total Do Pedido] = Despesa![Vr Total] + Despesa![Vr Do Ipi]
         ![Qtde Recebida] = Despesa!Qtde
         ![Qtde Restante] = 0
         ![Total Restante] = 0
         ![Notas] = NotaDespesa
         .Update
         .BookMark = .LastModified
      End With
      Despesa.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Baixa Realizada Com Sucesso") 
   GridBxDespesa.Rebind
   TxtnotaDespesa.Text = ""
End Sub


Private Function AjustaBxDespesas(Id_do_Pedido As Long, Id_da_Despesa As Long, Qtde_Total As Double, _
   Vr_Unitario As Double, Vr_Total_do_Pedido As Double, Qtde_Recebida As Double, _
   Qtde_Restante As Double, Total_Restante As Double, Notas As String)
 Dim QtdeRestante As Double, DinheiroRestante As Currency
  
  On Error GoTo DeuErro
  
   QtdeRestante = Qtde_Total - Qtde_Recebida
   DinheiroRestante = Round((Qtde_Total * Vr_Unitario ) - (Qtde_Recebida * Vr_Unitario),2)
   
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Bx Despesas Pedido Compra] Set [Qtde Restante] = " &  Substitui(CStr(QtdeRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id da Despesa] = " & Id_da_Despesa
   vgDb.Execute "Update [Bx Despesas Pedido Compra] Set [Total Restante] = " &  Substitui(CStr(DinheiroRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id da Despesa] = " & Id_da_Despesa 
   vgDb.CommitTrans
   
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If
End Function


Private Function AjustaConsumo() 
 Dim ValorIPI As Currency, ValorTotalItem As Currency, ValorIcmsItem As Currency
 Dim Consumo As New GRecordSet
  
  On Error GoTo DeuErro
  
  Set Consumo = VgDb.OpenRecordSet("SELECT * FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
  
  Do While Not Consumo.EOF
   ValorIPI = Round((((Consumo!Qtde * Consumo![Vr Unitario]) * Consumo![Aliquota Do IPI]) / 100),2)
   ValorTotalItem = Round((Consumo!Qtde * Consumo![Vr Unitario]),2)
   ValorIcmsItem = Round((((Consumo![Vr Unitario] * Consumo!Qtde + ((Consumo!Qtde * Consumo![Vr Unitario]) * Consumo![Aliquota Do IPI]) / 100 ) * Consumo![Aliquota Do Icms]) / 100),2)
  'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Consumo do Pedido Compra] Set [Vr do IPI] = " &  Substitui(CStr(ValorIPI), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id Despesa] = " & Consumo![Id Despesa]
   vgDb.Execute "Update [Consumo do Pedido Compra] Set [Vr Total] = " &  Substitui(CStr(ValorTotalItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id Despesa] = " & Consumo![Id Despesa]
   vgDb.Execute "Update [Consumo do Pedido Compra] Set [Vr do Icms] = " & Substitui(CStr(ValorIcmsItem), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id Despesa] = " & Consumo![Id Despesa]
   vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 0 Where [Id do Pedido] = " & Id_do_Pedido
   vgDb.CommitTrans
   Consumo.MoveNext
   Loop
   
   AjustaValores ' Atualiza o Pedido
  
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If

End Function


Private Sub LancaBxConsumoParcial()
   Dim Consumo As New GRecordSet
   Dim BxConsumo As New GRecordSet
   
   If Consumo_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Consumo_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Consumo_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Consumo = vgDb.OpenRecordSet("SELECT * FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxConsumo = VgDb.OpenRecordSet("Bx Consumo Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Consumo para Realizar Baixa
   Do While Not Consumo.EOF
      With BxConsumo
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Da Despesa] = Consumo![Id Da Despesa]
         ![Id Despesa] = Consumo![Id Despesa]
         ![Qtde Total] = Consumo!Qtde
         ![Vr Unitario]= Consumo![Vr Unitario]+(Consumo![Vr Do Ipi] / Consumo!Qtde)
         ![Vr Total Do Pedido] = Consumo![Vr Total] + Consumo![Vr Do Ipi]
         ![Qtde Recebida] = 0
         ![Qtde Restante] = Consumo!Qtde
         ![Notas] = NotaConsumo
         .Update
         .BookMark = .LastModified
      End With
      Consumo.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Mat de Consumo Incluidos no Grid com Sucesso Informe a Quantidade Recebida de Cada Item para Efetuar as Baixas") 
   GridBxConsumo.Rebind
End Sub


Private Sub LancaBxConsumoTotal()
   Dim Consumo As New GRecordSet
   Dim BxConsumo As New GRecordSet
   
   If Consumo_do_Pedido_Compra.RecordCount = 0 Then Exit Sub
   If Bx_Consumo_Pedido_Compra.RecordCount > 0 Then Exit Sub
   If Pedido_Fechado = 1 Then Exit Sub
   
   If Consumo_do_Pedido_Compra.RecordCount > 0 Then
   
   Set Consumo = vgDb.OpenRecordSet("SELECT * FROM [Consumo do Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set BxConsumo = VgDb.OpenRecordSet("Bx Consumo Pedido Compra")
   End If
   
   vgDb.BeginTrans
   'Inclui Consumo para Realizar Baixa
   Do While Not Consumo.EOF
      With BxConsumo
         .AddNew
         ![Id Do Pedido] = Id_do_Pedido
         ![Id Da Despesa] = Consumo![Id Da Despesa]
         ![Id Despesa] = Consumo![Id Despesa]
         ![Qtde Total] = Consumo!Qtde
         ![Vr Unitario]= Consumo![Vr Unitario]+(Consumo![Vr Do Ipi] / Consumo!Qtde)
         ![Vr Total Do Pedido] = Consumo![Vr Total] + Consumo![Vr Do Ipi]
         ![Qtde Recebida] = Consumo!Qtde
         ![Qtde Restante] = 0
         ![Total Restante] = 0
         ![Notas] = NotaConsumo
         .Update
         .BookMark = .LastModified
      End With
      Consumo.MoveNext
   Loop
   vgDb.CommitTrans
   MsgBox("Mat de Consumo Incluidos no Grid com Sucesso") 
   GridBxConsumo.Rebind
End Sub


Private Function AjustaBxConsumo()
 Dim QtdeRestante As Double, DinheiroRestante As Currency
 Dim BxConsumo As New GRecordSet
  
  On Error GoTo DeuErro
   
   Set BxConsumo = vgDb.OpenRecordSet("SELECT * FROM [Bx Consumo Pedido Compra] WHERE [Id do Pedido] = " & Id_do_Pedido)
   
   Do While Not BxConsumo.EOF
   QtdeRestante = BxConsumo![Qtde Total] - BxConsumo![Qtde Recebida]
   DinheiroRestante = Round((BxConsumo![Qtde Total] * BxConsumo![Vr Unitario]) - (BxConsumo![Qtde Recebida] * BxConsumo![Vr Unitario]),2)
   'Atualizando
   vgDb.BeginTrans
   vgDb.Execute "Update [Bx Consumo Pedido Compra] Set [Qtde Restante] = " &  Substitui(CStr(QtdeRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id Despesa] = " & BxConsumo![Id Despesa] 
   vgDb.Execute "Update [Bx Consumo Pedido Compra] Set [Total Restante] = " &  Substitui(CStr(DinheiroRestante), ",", ".", SO_UM) & " WHERE [Id do Pedido] = " &  Id_do_Pedido & " And [Id Despesa] = " & BxConsumo![Id Despesa] 
   vgDb.CommitTrans
   BxConsumo.MoveNext
   Loop
   
  DeuErro:
   If Err Then
     MsgBox Err.Descption, vbCritical + vbOKOnly, vaTitulo
     vgDb.RollBackTrans
   End If
End Function


Private Function UltimaParcela() As Long
   Dim Tb As GRecordset
   
   On Error Resume Next
   
   Set Tb = vgDb.OpenRecordSet("SELECT MAX([Número da Parcela]) Pc FROM [Parcelas Ped Compra Novo] WHERE [Id do Pedido] = " & Id_do_Pedido)
         
   If Tb.RecordCount > 0 Then
      UltimaParcela = Tb!PC + 1
   Else
      UltimaParcela = 1
   End If  

End Function


Private Function UltimoVencimento() As Variant
   Dim Tb As New GRecordSet, RetVal As Variant
   
   On Error Resume Next
      
   Set Tb = vgDb.OpenRecordSet("SELECT MAX([Data de Vencimento]) [Data de Vencimento] FROM [Parcelas Ped Compra Novo] WHERE [Id do Pedido] = " & Id_do_Pedido)
      
   RetVal = gridParcelas.ColumnValue(gridParcelas.Row + 1, CInt(gridParcelas.Columns("Data de Vencimento").Index))
      
   'If Vendas![Forma de pagamento] = "PRAZO" Then
      If Not IsNull(Tb![Data de Vencimento]) Then
         'Bom Vamos deixar ele colocar qualquer vencimento
         'Soh devemos tomar cuidado para não ser inferior a data de emissão
         UltimoVencimento = True
      Else
         UltimoVencimento = RetVal >= Pedido_de_Compra_Novo![Data Do Pedido] 
      End If
  ' Else
  '    UltimoVencimento = RetVal = Data_de_emissao
  ' End If

End Function


Private Function TotalParcelas(Optional Seq As Integer) As Currency
   Dim Tb As GRecordSet, Total As Currency
   
   Set Tb = vgDb.OpenRecordSet("SELECT Sum([Valor da Parcela]) As Total From [Parcelas Ped Compra Novo] Where [Id do Pedido] = " & Id_do_Pedido & " AND " & IIf(Seq > 0, "[Número da Parcela] = " & Seq, "1=1"))
   
   Total = Tb!Total
   
   TotalParcelas = Total

End Function


Private Sub MostraFormulas()
   On Error Resume Next     ' se der erro que si foda bora limpa la em baixo...                      
   lblParcelamento.Caption = "Parcelamento"
   If Err Then Err = 0: lblParcelamento.Caption = ""
   'Label do Parcelamento
   lblParcelamento.BackColor = &H0: lblParcelamento.ForeColor = &HFFFFFF
   If TotalParcelas = 0 Then
      lblParcelamento.BackColor = &H80FFFF: lblParcelamento.ForeColor = &H80000012
   ElseIf (TotalParcelas < Pedido_de_Compra_Novo![Total Do Pedido] ) Or (TotalParcelas > Pedido_de_Compra_Novo![Total Do Pedido]) Then
      lblParcelamento.BackColor = &H8FF: lblParcelamento.ForeColor = &HFFFFFF
   End If
   Lblfin.Caption = MostraTotal() & " - " & MostraFinanceiro()
   
   If Me.Caption = "Pedido" Then
      lblPedido.Caption = "Pedido de Compra"
   Else
      lblPedido.Caption = "Pré Pedido"
   End If     
   
   
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub 


Private Function ValidaProduto(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, _
   Sequencia_do_Item As Long, Sequencia_do_Orcamento As Long) As Boolean

 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * From Produtos Where [Seqüência do Produto] = " & Id_do_Produto)
 
 If Tb.RecordCount > 0 Then
   If Tb!Inativo Then
      ValidaProduto = False
      MDIIRRIG.CancelaAlteracoes: Exit Function
   Else
      ValidaProduto = True: Exit Function      
   End If
 End If
  
End Function


Private Function InfoProdutos(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, _
   Sequencia_do_Item As Long, Sequencia_do_Orcamento As Long, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Produtos", "[Seqüência do Produto] = " & Id_do_Produto, ProdutoAux
      
   If ProdutoAux.RecordCount = 0 Then Exit Function  
    
   Select Case Oq
      Case "CUSTO"
         InfoProdutos = ProdutoAux![Valor de Custo]     
   End Select

End Function


Private Function InfoDespesas(Id_do_Pedido As Long, Id_da_Despesa As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Despesas", "[Seqüência da Despesa] = " & Id_da_Despesa, DespesaAux
      
   If DespesaAux.RecordCount = 0 Then Exit Function  
    
   Select Case Oq
      Case "CUSTO"
         InfoDespesas = DespesaAux![Valor de Custo]     
   End Select

End Function

Private Function InfoConsumo(Id_do_Pedido As Long, Id_Despesa As Long, Id_da_Despesa As Long, _
   Qtde As Double, Vr_Unitario As Double, Vr_Total As Double, _
   Aliquota_do_IPI As Double, Aliquota_do_Icms As Double, Vr_do_IPI As Double, _
   Vr_do_Icms As Double, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Despesas", "[Seqüência da Despesa] = " & Id_da_Despesa, ConsumoAux 
      
   If ConsumoAux.RecordCount = 0 Then Exit Function  
    
   Select Case Oq
      Case "CUSTO"
         InfoConsumo = ConsumoAux![Valor de Custo]     
   End Select

End Function


Private Function VerificaUltCusto() As Boolean
   Dim PProdutos As New GRecordSet, PDespesas As New GRecordSet, Consumo As New GRecordSet
   Dim Sequencia() As Long, i As Long, Mensagem As String, Campo As Variant

   On Error GoTo DeuErro
   
   
   Set PProdutos = vgDb.OpenRecordSet("SELECT PP.[Id do Produto], PP.[Vr Unitario], P.[Valor de Custo], P.[Seqüência Do Produto] FROM [Produtos do Pedido Compra] PP INNER JOIN Produtos P ON PP.[Id do Produto] = P.[Seqüência do Produto] WHERE [Id do Pedido] = " & Pedido_de_Compra_Novo![Id Do Pedido])
   Set PDespesas = vgDb.OpenRecordSet("SELECT DP.[Id da Despesa], DP.[Vr Unitario], D.[Valor de Custo], D.[Seqüência Da Despesa] FROM [Despesas do Pedido Compra] DP INNER JOIN Despesas D ON DP.[Id da Despesa] = D.[Seqüência da Despesa] WHERE [Id do Pedido] = " & Pedido_de_Compra_Novo![Id Do Pedido])
   Set Consumo = vgDb.OpenRecordSet("SELECT CP.[Id da Despesa], CP.[Vr Unitario], C.[Valor de Custo], C.[Seqüência Da Despesa] FROM [Consumo do Pedido Compra] CP INNER JOIN Despesas C ON CP.[Id da Despesa] = C.[Seqüência da Despesa] WHERE [Id do Pedido] = " & Pedido_de_Compra_Novo![Id Do Pedido])
   
   i = 0 'Tamanho do Vetor
   ReDim Preserve Sequencia(0) As Long
   
  'Produtos
   Do While Not PProdutos.EOF
       If (PProdutos![Valor de Custo] < PProdutos![Vr Unitario]) Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = PProdutos![Seqüência Do Produto]
       PProdutos.MoveNext
   Loop
   
   'Despesas
   Do While Not PDespesas.EOF
      If (PDespesas![Valor de Custo] < PDespesas![Vr Unitario]) Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = PDespesas![Seqüência Da Despesa]
      PDespesas.MoveNext
   Loop
   
   'Consumo
   Do While Not Consumo.EOF
      If (Consumo![Valor de Custo] < Consumo![Vr Unitario]) Then i = i + 1: ReDim Preserve Sequencia(i): Sequencia(i - 1) = Consumo![Seqüência Da Despesa]
      Consumo.MoveNext
   Loop
       
   'vamos mostrar os Itens que aumentou o Preço des da Ultima Compra
   If UBound(Sequencia) > 0 Then
      Mensagem = "ATENÇÃO!!! Alguns Produtos com Preço Superior a Ultima Compra" & vbCrLf
      For Each Campo In Sequencia
         If Campo > 0 Then Mensagem = Mensagem & vbCrLf & Campo
      Next
      Mensagem = Mensagem & vbCrLf & vbCrLf & "Deseja Realmente Concluir o Pedido com o Preço Superior ao da Ultima Compra?"  
      If MsgBox(Mensagem, vbExclamation + vbYesNo + vbDefaultButton2, vaTitulo) = vbYes Then
         SuperInput2         
         If Not Vazio(InputYgor) Then
            If InputYgor = SuperSenhaPedido Then
               VerificaUltCusto = True
               vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 1 Where [Id do Pedido] = " & Id_do_Pedido 
            Else
               MsgBox "Senha Incorreta", vbCritical + vbOKOnly, vaTitulo
               VerificaUltCusto = False
               vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 0 Where [Id do Pedido] = " & Id_do_Pedido 
            End If
            InputYgor = ""
         End If
      Else
         VerificaUltCusto = False
         vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 0 Where [Id do Pedido] = " & Id_do_Pedido
      End If
   Else
      VerificaUltCusto = True
      vgDb.Execute "UPDATE [Pedido de Compra Novo] Set Validado = 1 Where [Id do Pedido] = " & Id_do_Pedido 
   End If
   
   Reposition True
   
DeuErro:
   If Err <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
   End If

End Function


Private Function TemBaixa() As Boolean
   Dim Baixa As New GRecordset
   
   On Error Resume Next
   
   Set Baixa = vgdb.OpenRecordset("SELECT * FROM [Concilia Conta Antecipada] WHERE [Sequencia da Compra] = " & Id_do_Pedido)
   
   If Baixa.RecordCount > 0 Then
      TemBaixa = True
   Else
      TemBaixa = False
   End If

End Function


Private Sub LancaParcela()
   Dim Manutencao As New GRecordSet, Parcelas As New GRecordSet, Baixa As New GRecordset
   Dim Manu2 As New GRecordSet, Entrada As New GRecordset, SeqEntrada As Long
   Dim ParcelasEntrada As New GRecordSet, Parcelas1 As New GRecordSet
      
   On Error GoTo DeuErro
   
   Screen.MousePointer = vbHourglass 'Ampulheta
   
   If TemConta Then
      MsgBox "Esta conta ja foi gerada por favor confira no contas a pagar!", vbExclamation + vbOKOnly, ""
      Screen.MousePointer = vbDefault
      Exit Sub
   End If 
   
   If Vazio(Tipo_da_Licitacao) Then
      MsgBox "Informe o tipo da licitação!", vbExclamation + vbOKOnly, "" 
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   If Parcelas_Ped_Compra_Novo.RecordCount = 0 Then
      MsgBox "Inclua as parcelas para gerar a conta!", vbExclamation + vbOKOnly, "" 
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
     
   If TotalParcelas() <> Pedido_de_Compra_Novo![Total Do Pedido] Then
      MsgBox "Total das parcelas não confere com o total do pedido impossivel gerar a conta!", vbExclamation + vbOKOnly, "" 
      Screen.MousePointer = vbDefault 
      Exit Sub
   End If
   
   If MsgBox("Deseja gera o contas á pagar referente este pedido no financeiro?", vbQuestion + vbYesNo, vaTitulo) = vbNo Then 
      Screen.MousePointer = vbDefault 
      Exit Sub
   End If
     
   Set Manutencao = vgDb.OpenRecordSet("SELECT * FROM [Manutenção Contas]")
   Set Parcelas = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Ped Compra Novo] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set Parcelas1 = vgDb.OpenRecordSet("SELECT * FROM [Parcelas Ped Compra Novo] WHERE [Id do Pedido] = " & Id_do_Pedido)
   Set Entrada = vgDb.OpenRecordSet("SELECT * FROM [Entrada Contas]") 
   Set ParcelasEntrada = VgDb.OpenRecordSet("SELECT * FROM [Parcelas Entrada Contas]") 
    
   vgDb.BeginTrans
   
   With Entrada
     .AddNew 
       ![Data de Entrada] = Pedido_de_Compra_Novo![Data Do Pedido]
       ![Seqüência Do Geral] = Pedido_de_Compra_Novo![Codigo Do Fornecedor] 
       ![Histórico] = "Referente Pedido de Compra Nº " & Format(Id_do_Pedido,"000000")
       ![Forma de Pagamento] = "Prazo"
       ![Valor Total] = Pedido_de_Compra_Novo![Total Do Pedido]
       ![Seqüência Grupo Despesa] = Sequencia_Grupo_Despesa 
       ![Seqüência SubGrupo Despesa] = Sequencia_SubGrupo_Despesa   
       !Conta = "P"
        If Pedido_de_Compra_Novo!Prazo = "Prazo" Then
           ![Documento] = "Prev. " & Pedido_de_Compra_Novo![Id Do Pedido] 
           !Previsão = 1
        Else
           ![Documento] = Pedido_de_Compra_Novo!Prazo        
        End If   
       ![Sequencia da Compra] = Id_do_Pedido 
       ![Titulo] = "Fornecedores"  
       ![Data da Alteração] = Date
       ![Hora da Alteração] = Time
       ![Usuário da Alteração] = vgPWUsuario 
     .Update
     .BookMark = .LastModified 
      SeqEntrada = ![Seqüência da Entrada]          
   End With
   'Parcelas Entrada
   Do While Not Parcelas1.EOF 
     With ParcelasEntrada
        .AddNew
          ![Seqüência da Entrada] = SeqEntrada
          ![Número da Parcela] = Parcelas1![Número da Parcela]
          !Dias = Parcelas1!Dias 
          ![Data de Vencimento] = Parcelas1![Data de Vencimento] 
          ![Valor da Parcela] = Parcelas1![Valor da Parcela]
          ![Seqüência da Cobrança] = 18  
        .Update
        .BookMark = .LastModified    
     End With
     Parcelas1.MoveNext
    Loop  
   'Manutenção
    Do While Not Parcelas.EOF
       TbAuxiliar "Manutenção Contas", "[Sequencia da Compra] = " & Id_do_Pedido & " AND Parcela = " & Parcelas![Número da Parcela], Manutencao
       TbAuxiliar "Manutenção Contas", "[Sequencia da Compra] = " & Id_do_Pedido & " AND Parcela = " & Parcelas![Número da Parcela], Manu2
        With Manutencao         
          If Manutencao.RecordCount = 0 Then .AddNew              
            !Conta = PAGAR    
            ![Seqüência Grupo Despesa] = Sequencia_Grupo_Despesa 
            ![Seqüência SubGrupo Despesa] = Sequencia_SubGrupo_Despesa
             If Pedido_de_Compra_Novo!Prazo = "Prazo" Then
                !Documento = "Prev. " & Id_do_Pedido
                !Previsão = 1
             Else
                !Documento = Pedido_de_Compra_Novo!Prazo
             End If   
            !Titulo = "Fornecedores"
            ![Seqüência da Cobrança] = Parcelas![Seqüência da Cobrança]
            ![Sequencia da Compra] = Id_do_Pedido
            ![Seqüência da Origem] = SeqEntrada  
            ![Parcela] = Parcelas![Número da Parcela]
            ![Seqüência Do Geral] = Pedido_de_Compra_Novo![Codigo Do Fornecedor]       
            ![Data de Entrada] = Pedido_de_Compra_Novo![Data Do Pedido] 
            !Histórico = "Referente Pedido de Compra Nº " & Format(Id_do_Pedido,"000000")
            ![Forma de Pagamento] = "Prazo"
            ![Valor Total] = Pedido_de_Compra_Novo![Total Do Pedido]
            ![Data de Vencimento] = Parcelas![Data de Vencimento]
             If Manu2.RecordCount > 0 Then 
               ![Vencimento Original] = Manu2![Vencimento Original]
             Else
               ![Vencimento Original] = Parcelas![Data de Vencimento]     
             End If
            ![Valor da Parcela] = Parcelas![Valor da Parcela]
            ![Valor Restante] = Parcelas![Valor da Parcela]
            ![Data da Baixa] = Null
            ![Valor Pago] = 0
             .Update
             .BookMark = .LastModified         
        End With
       Parcelas.MoveNext
    Loop
    vgDb.CommitTrans
    MsgBox "Conta Gerada com Sucesso!", vbExclamation + vbOKOnly, "" 
    Screen.MousePointer = vbDefault
    txtTipo_da_Licitacao.Text = ""
    
DeuErro:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo
      vgDb.RollBackTrans
   End If
  
End Sub

Private Function TemConta() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT * FROM [Manutenção Contas] WHERE [Sequencia da Compra] = " & Id_do_Pedido)
 
  If Tb.RecordCount > 0 Then
     TemConta = True
  Else
     TemConta = False
  End If  
   
End Function


Private Function TemPagto() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Manutenção Contas] WHERE [Valor Pago] > 0 And [Sequencia da Compra] = " & Id_do_Pedido)
 
 If Tb.RecordCount > 0 Then
    TemPagto = False
 Else
    TemPagto = True
 End If      
 
End Function


Private Function ValorTotalConjunto(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, _
   Sequencia_do_Item As Long, Sequencia_do_Orcamento As Long)
           
 Dim Tb As New GRecordset
 Dim TbConj As New GRecordset
 Dim ValorCusto As Currency
 
 Set TbConj = VgDb.OpenRecordSet("SELECT * FROM [Itens do Conjunto] Where [Seqüência do Produto] = " & Id_do_Produto)

  Do While Not TbConj.EOF
     Set Tb = vgdb.OpenRecordset("SELECT SUM([Valor de Custo] * [Quantidade do Produto]) AS Custo FROM Produtos INNER JOIN [Itens do Conjunto] ON Produtos.[Seqüência do Produto] = [Itens do Conjunto].[Seqüência do Produto] WHERE [Seqüência do Conjunto] = " & TbConj![Seqüência Do Conjunto])  
     ValorCusto = Round(Tb!Custo + Tb!Custo * Parametros_do_Produto![Percentual Acréscimo Conjunto] / 100, 2)
       ' If TbConj![Seqüência Do Conjunto] <> 3268 Then
           vgDb.Execute "Update Conjuntos Set [Valor Total] = " & Substitui(ValorCusto * 1.7, ",", ".", SO_UM) & " WHERE [Seqüência do Conjunto] = " & TbConj![Seqüência Do Conjunto] 
        'Else
        '    vgDb.Execute "Update Conjuntos Set [Valor Total] = " & Substitui(ValorCusto * 2, ",", ".", SO_UM) & " WHERE [Seqüência do Conjunto] = " & TbConj![Seqüência Do Conjunto] 
       ' End If     
  TbConj.MoveNext
  Loop
 
End Function


Private Function FiltroOrc() As String
   FiltroOrc = "[Seqüência do Orçamento] > 0 AND [Cancelado] = 0"
End Function


Private Function InfoOrc(Id_do_Pedido As Long, Sequencia_do_Orcamento As Long, Sequencia_do_Item As Long, Oq As String) As Variant
   
   On Error Resume Next
   
   TbAuxiliar "Orçamento", "[Seqüência do Orçamento] = " & Sequencia_do_Orcamento, OrcAux
      
   If OrcAux.RecordCount = 0 Then Exit Function  
   TbAuxiliar "Geral", "[Seqüência do Geral] = " & OrcAux![Seqüência Do Geral] & " AND [Seqüência do Geral] > 0", ClienteAux     
  
   Select Case Oq
      Case "Cliente"
         InfoOrc = ClienteAux![Razão Social]     
      Case "Emissao"
         InfoOrc = OrcAux![Data de Emissão]
   End Select

End Function


Private Function SeqItem() As Integer
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do Item]) Seq From [Orçamentos da Compra] Where [Id do Pedido] = " & Id_do_Pedido)
 
  If Tb.RecordCount > 0 Then
     SeqItem = Tb!Seq + 1  
  Else
     SeqItem = 1
  End If
 
End Function


Private Function MostraFinanceiro() As String
 'Dim Tb As New GRecordSet
 Dim Tb1 As New GRecordSet
 
 'Set Tb = vgDb.OpenRecordSet("SELECT * From [Manutenção Contas] WHERE [Sequencia da Compra] = " & Pedido_de_Compra_Novo![Id Do Pedido])
 Set Tb1 = vgDb.OpenRecordSet("SELECT SUM([Valor Pago]) Pago From [Manutenção Contas] WHERE [Sequencia da Compra] = " & Id_do_Pedido)
 
 'Do While NOT Tb.EOF
    MostraFinanceiro = "Valor Pago R$: " & Format(Tb1!Pago, "##,###,##0.00")
 'Tb.MoveNext
 'Loop
 
End Function

Private Function MostraTotal() As String
 Dim Tb As New GRecordSet
 Set Tb = vgDb.OpenRecordSet("SELECT [Total do Pedido] From [Pedido de Compra Novo] WHERE [Id do Pedido] = " & Id_do_Pedido)
 MostraTotal = "Total do Pedido R$: " & Format(Tb![Total Do Pedido], "##,###,##0.00")
End Function


Private Function SeqItem1() As Long
 Dim Tb As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT MAX([Sequencia do item]) Seq From [Produtos do Pedido Compra] Where [Id do Pedido] = " & Id_do_Pedido)
 
    If Tb.RecordCount > 0 Then
      SeqItem1 = Tb!Seq + 1
    Else
      SeqItem1 = 1
    End If
   
   
End Function


Private Function ValidaOrc() As Boolean
 Dim Tb As New GRecordSet
 
 Set Tb = vgdb.OpenRecordSet("SELECT * FROM [Orçamentos da Compra] Where [Id do Pedido] = " & Id_do_Pedido)
 
 If Tb.RecordCount = 1 Then
    ValidaOrc = False
    Exit Function
 Else
    ValidaOrc = True
    Exit Function
 End If   
 
End Function

Public Sub VerificaFormulario
 Dim Tb As New GRecordSet

   On error Resume Next
   
   If Me.Caption = "Pedido" And Not PrePedido Then
      PedidoAberto = False
      Set Tb = VgDb.OpenRecordSet("SELECT * FROM [Manutenção Contas] Where [Sequencia da Compra] = " & Id_do_Pedido)
          If Tb.RecordCount = 0 And Not Pedido_Fechado And Prazo = "Antecipado" Then
             MsgBox"ATENÇÃO VOCÊ NAO GEROU O FINANCEIRO!" & vbCrLf & "Usuario - "  & vgPWUsuario , vbError 
             Load frmPComprN
             frmPComprN.Show
             frmPComprN.vgFiltroInicial = "[Id do Pedido] = " & Id_do_Pedido
             InicializaFiltro frmPComprN
             frmPComprN.btnGerarconta.SetFocus 
          End If
   Else
      PrePedidoAberto = False
   End If
   
   lblAlteracao.caption = ""

End Sub

Private Sub AtuPrevisao()
 If Dias = 0 Then Exit Sub
    Txtprevisao.Text = DateAdd("d", Dias, Data_Do_Pedido)
End Sub



Private Sub Extornar()
 Dim Tb As New GRecordSet, Entrada As New GRecordSet
 Dim Manu As New GRecordSet, Baixa As New GRecordSet
 Dim Conta As New GRecordSet, SQL As String
 Dim SeqBaixa As Long, Manutencao As New GRecordSet
 Dim MvtoCC As New GRecordSet, PcEntrada As New GRecordSet
 
 Set Tb = VgDb.OpenRecordSet("SELECT [Seqüência da Entrada] From [Entrada Contas] where [Sequencia da Compra] = " & Id_do_Pedido)
 Set Manu = VgDb.OpenRecordSet("SELECT SUM([Valor Pago]) Pago, SUM([Valor Restante]) Restante From [Manutenção Contas] WHERE [Sequencia da Compra] = " & Id_do_Pedido) 
 
 If Manu!Restante = 0 Then
    MsgBox("Esta conta ja esta Paga!") 
    Exit Sub
 End If
  
 If Manu!Pago = 0 Then
    VgDb.Execute "DELETE FROM [Manutenção Contas] WHERE [Sequencia da Compra] = " & Id_do_Pedido 
    VgDb.Execute "DELETE FROM [Entrada Contas] WHERE [Seqüência da Entrada] = " & Tb![Seqüência da Entrada]
    VgDb.Execute "DELETE FROM [Parcelas Entrada Contas] WHERE [Seqüência da Entrada] = " & Tb![Seqüência da Entrada]
 'Else
 '   Set Baixa = VgDb.OpenRecordSet("Baixa Contas")
 '   Set Manutencao = VgDb.OpenRecordSet("Manutenção Contas")
 '   Set Entrada = VgDb.OpenRecordSet("Entrada Contas")
 '   Set MvtoCC = VgDb.OpenRecordSet("Movimentação da Conta Corrente")
 '   Set PcEntrada = VgDb.OpenRecordSet("Parcelas Entrada Contas") 
     
  '       Sql = "SELECT [Seqüência da Cobrança], [Baixa Contas].[Valor do Desconto], [Baixa Contas].[Valor do Juros], Geral.[Seqüência do Geral], [Razão Social], [Notas da Compra], [Baixa Contas].[Seqüência Grupo Despesa], [Baixa Contas].[Seqüência SubGrupo Despesa], [Seqüência da Origem], Documento, Parcela, [Data de Entrada], [Data de Vencimento], [Vencimento Original], [Valor Restante], [Valor da Parcela], [Valor Total], [Baixa Contas].[Valor Pago], [Baixa Contas].[Data da Baixa], [Forma de Pagamento],"
  ' Sql = Sql & "Parcela, [Data de Entrada], [Baixa Contas].[Seqüência da Baixa], [Baixa Contas].[Histórico], [Sequencia da Compra], Titulo, [Seqüência da Agência], [Manutenção Contas].[Seqüência da Manutenção], [Seqüência da CC da Agência], [Seqüência da Movimentação CC], [Codigo do Historico], [Codigo do Debito], [Codigo do Credito] FROM [Manutenção Contas] INNER JOIN Geral ON [Manutenção Contas].[Seqüência do Geral] = Geral.[Seqüência do Geral] INNER JOIN [Baixa Contas] On [Manutenção Contas].[Seqüência da Manutenção] = [Baixa Contas].[Seqüência da Manutenção] WHERE " & Filtro() & " ORDER BY [Baixa Contas].[Data da Baixa]"
           
  ' Set Conta = VgDb.OpenRecordSet(SQL)
  '     Do While Not Conta.EOF
  '        VgDb.Execute "DELETE FROM [Baixa Contas] Where [Seqüência da Baixa] > 0 And [Seqüência da Baixa] = " & Conta![Seqüência da Baixa]
  '        VgDb.Execute "DELETE FROM [Lançamentos Contabil] Where [Seqüência da Baixa] > 0 And [Seqüência da Baixa] = " & Conta![Seqüência da Baixa]
  '        VgDb.Execute "DELETE FROM [Movimentação da Conta Corrente] Where [Seqüência da Movimentação CC] > 0 And [Seqüência da Movimentação CC] = " & Conta![Seqüência da Movimentação CC]
  '        VgDb.Execute "DELETE FROM [Entrada Contas] Where [Sequencia da Compra] > 0 And [Sequencia da Compra] = " & Conta![Sequencia da Compra]
  '        VgDb.Execute "DELETE FROM [Parcelas Entrada Contas] Where [Seqüência da Entrada] > 0 And [Seqüência da Entrada] = " & Conta![Seqüência da Origem]
  '        VgDb.Execute "DELETE FROM [Manutenção Contas] Where [Sequencia da Compra] > 0 And [Sequencia da Compra] = " & Conta![Sequencia da Compra]
            
          'Entrada Contas
 '            With Entrada
 '             .AddNew
 '               ![Data de Entrada] = Conta![Data de Entrada]
 '               ![Seqüência Do Geral] = Conta![Seqüência Do Geral]
 '               !Documento = 
 '             .Update
 '             .BookMark = .LastModified 
 '            End With
 '            
 '          With Manutencao 'Contas
 '            .AddNew 
 '              !Parcela = Conta!Parcela
 '              ![Seqüência Do Geral] = Conta![Seqüência Do Geral]
 '              ![Data de Entrada] = Conta![Data de Entrada]
 '              ![Histórico] = Conta![Histórico]
 '              ![Vencimento Original] = Conta![Vencimento Original]
 '              ![Data de Vencimento] = Conta![Data de Vencimento]
 '              ![Forma de Pagamento] = Conta![Forma de Pagamento]
 '              ![Data de Vencimento] = Conta![Data de Vencimento]
 '              ![Valor Total] = Conta![Valor Total]
 '              ![Valor da Parcela] = Conta![Valor Pago]
 '              ![Valor Pago] = Conta![Valor Pago]
 '              ![Valor Restante] = 0
 '              ![Valor Do Juros] = Conta![Valor Do Juros]
 '              ![Valor Do Desconto] = Conta![Valor Do Desconto]
 '              ![Data da Baixa] = Conta![Data da Baixa]
 '              !Conta = "P"
 '              ![Seqüência da Cobrança] = Conta![Seqüência da Cobrança]
 '              ![Seqüência da Origem] = Conta![Seqüência da Origem]
 '              ![Seqüência Grupo Despesa] = Conta![Seqüência Grupo Despesa]
 '              ![Sequencia da Compra] = Conta![Sequencia da Compra]
 '              ![Notas da Compra] = Conta![Notas da Compra]
 '              !Titulo = Conta!Titulo  
 '            .Update
 '            .BookMark = .LastModified 
 '          End With
          
 '          With Baixa 'Baixa
 '            .AddNew       
 '              ![Seqüência da Manutenção] = Conta![Seqüência da Manutenção]
 '              ![Data da Baixa] = Conta![Data da Baixa]
 '              ![Valor Pago] = Conta![Valor Pago]
 '              ![Valor Do Juros] = Conta![Valor Do Juros]
 '              ![Seqüência da Agência] = Conta![Seqüência da Agência]
 '              ![Seqüência da CC da Agência] = Conta![Seqüência da CC da Agência]
 '              !Conta = "P"
 '              ![Histórico] = Baixa![Histórico]
 '              ![Codigo Do Historico] = Conta![Codigo Do Historico]
 '              ![Codigo Do Debito] = Conta![Codigo Do Debito]
 '              ![Codigo Do Credito] = Conta![Codigo Do Credito]
 '              ![Seqüência Grupo Despesa] = Conta![Seqüência Grupo Despesa]
 '              ![Seqüência SubGrupo Despesa] = Conta![Seqüência SubGrupo Despesa]
 '            .Update
 '            .BookMark = .LastModified 
 '             SeqBaixa = ![Seqüência da Baixa]
 '             VgDb.Execute "UPDATE [Manutenção Contas] Set [Seqüência da Baixa] =  " & SeqBaixa & " WHERE [Seqüência da Manutenção] = " & Conta![Seqüência da Manutenção]
 '           End With
 '          'MvtoCC
 '          'With 
 '          
 '         ' End With 
 '         'Lan Contabil
 '  
 '         'With
 '         
 '         'End With
 '           
 '      Conta.MoveNext
 '      Loop    
 End If
 
 MsgBox("Extornado com Sucesso!")
 Reposition True

End Sub

Public Sub BloqueiaAba()       
 Aba1.TabVisible(6) = False 
 Aba1.TabVisible(7) = False       
End Sub

Private Function Filtro() As String
   Filtro = "[Sequencia da Compra] = " & Id_do_Pedido
   Filtro = Filtro & " AND [Valor Restante] > 0 "'And Parcela = " & Parcela
   Filtro = Filtro & " AND [Manutenção Contas].Conta = 'P'"  
End Function



Public Sub ContaEstoque(Produto As Long)
 Dim Saidas As New GRecordSet
 Dim Saidas2 As New GRecordSet
 Dim Entradas As Double
 '=======================================
 Dim Produtos As New GRecordSet
 Dim Estoque As Double
 Dim Comprados As Double
 
 Estoque = 0
 
 Set Produtos = vgDb.OpenRecordSet("SELECT * FROM Produtos Where Inativo = 0 And Produtos.[Seqüência do Produto] = " & Produto)  
 
   Do While Not Produtos.EOF
      Estoque = Produtos![Quantidade Contábil]
       
      Set Saidas2 = vgDb.OpenRecordSet("SELECT SUM(Quantidade) Qtd, [Seqüência da Matéria Prima] From [Receita Primaria] Left Join Orçamento On [Receita Primaria].[Seqüência do Orçamento] = Orçamento.[Seqüência do Orçamento] Where [Seqüência da Matéria Prima] = " & Produtos![Seqüência Do Produto] & Filtro2() & " Group by [Seqüência da Matéria Prima]") 'Separado p/ Entregar
      Set Saidas = vgDb.OpenRecordSet("SELECT SUM(Quantidade) Qtd, [Seqüência da Matéria Prima] From [Receita Primaria] Left Join Orçamento On [Receita Primaria].[Seqüência do Orçamento] = Orçamento.[Seqüência do Orçamento] Where [Seqüência da Matéria Prima] = " & Produtos![Seqüência Do Produto] & Filtro1() & " Group by [Seqüência da Matéria Prima]") 'Separado p/ Montar
      
      If Saidas2.RecordCount > 0 Then 
         Estoque = Estoque - Saidas2!Qtd
         vgDb.Execute "Update Produtos Set [Qtde Inicial] = " & Substitui(Saidas2!Qtd, ",", ".", SO_UM) & " WHERE [Seqüência do Produto] = " & Produtos![Seqüência Do Produto]
      End If
      If Saidas.RecordCount > 0 Then 
         Estoque = Estoque - Saidas!Qtd
         vgDb.Execute "Update Produtos Set [Separado Montar] = " & Substitui(Saidas!Qtd, ",", ".", SO_UM) & " WHERE [Seqüência do Produto] = " & Produtos![Seqüência Do Produto]
      End If
      If Produtos![Material Adquirido de Terceiro] Then
         Comprados = QtdeRecebida(Produtos![Seqüência Do Produto])
         Estoque = Estoque + Comprados
      End If
      
   vgDb.Execute "Update Produtos Set [Quantidade Fisica] = " & Substitui(CDbl(Estoque), ",", ".", SO_UM) & " WHERE [Seqüência do Produto] = " & Produtos![Seqüência Do Produto]  
   Produtos.MoveNext
   Loop
    
End Sub

Private Function QtdeRecebida(Produto As Long) As Double
 Dim Tb As New GRecordSet
 Dim Tb1 As New GRecordSet 

 Set Tb = vgDb.OpenRecordSet("SELECT SUM([Qtde Restante]) X From [Bx Produtos Pedido Compra] Left Join [Pedido de Compra Novo] On [Bx Produtos Pedido Compra].[Id do Pedido] = [Pedido de Compra Novo].[Id do Pedido] Where [Bx Produtos Pedido Compra].[Id do Produto] = " & Produto & " And Cancelado = 0 And [Pedido Fechado] = 0 And Prepedido = 0")
 Set Tb1 = vgDb.OpenRecordSet("SELECT SUM(Qtde) X From [Produtos do Pedido Compra] Left Join [Pedido de Compra Novo] On [Produtos do Pedido Compra].[Id do Pedido] = [Pedido de Compra Novo].[Id do Pedido] Where [Produtos do Pedido Compra].[Id do Produto] = " & Produto & " And Cancelado = 0 And [Pedido Fechado] = 0 And Prepedido = 0")
 QtdeRecebida = 0
   
   If Tb.RecordCount > 0 Then
   
      If Tb!x > 0 Then
         QtdeRecebida = Tb1!X - Tb!X
      Else
         QtdeRecebida = Tb1!X
      End If   
      
      vgDb.Execute "Update Produtos Set [Comprados Aguardando] = " & Substitui(Cdbl(QtdeRecebida), ",", ".", SO_UM) & " WHERE [Seqüência do Produto] = " & Produto
   End If
 
End Function

Private Function Filtro1() As String
  Filtro1 = " And Situação = 'Separado p/ Montar' And Cancelado = 0 And [Data do Fechamento] <= " & D(Date)
End Function

Private Function Filtro2() As String
  Filtro2 = " And Situação = 'Separado p/ Entregar' And Cancelado = 0 And [Data do Fechamento] <= " & D(Date)
End Function


Private Function MGeral(Id_do_Pedido As Long, Id_do_Produto As Long, Qtde As Double, _
   Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double, _
   Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double, _
   Sequencia_do_Item As Long, Sequencia_do_Orcamento As Long) As String
 Dim Tb As New GRecordSet
 Dim Tb1 As New GRecordSet
 
 Set Tb = vgDb.OpenRecordSet("SELECT [Seqüência Do Geral] Seq From Orçamento Where [Seqüência Do Orçamento] = " & Sequencia_do_Orcamento)
 Set Tb1 = vgDb.OpenRecordSet("SELECT [Razão Social] X From Geral Where [Seqüência Do Geral] = " & Tb!Seq)
 MGeral = Tb1!x
  
End Function

Public Property Get txtCampoTab(Index As Integer) As FormataCampos   
   Set txtCampoTab = txtCampo(Index)
End Sub

'inicializa variaveis (apelidos) coms os campos correspondentes
Private Sub InicializaApelidos(vgComOQue As Integer)
   On Error Resume Next                           'prepara para possíveis erros
   If (vgTb.RecordCount > 0 And vgTb.Eof = False And vgTb.Bof = False) Or vgComOQue = COM_TEXTBOX Then
      If vgComOQue = COM_TEXTBOX Then
         Id_do_Pedido = txtCampo(0).Value
         Data_do_Pedido = txtCampo(1).Value
         Nro_da_Licitacao = txtCampo(2).Value
         Codigo_do_Fornecedor = txtCampo(7).Value
         Codigo_da_Transportadora = txtCampo(8).Value
         Comprador = txtCampo(3).Value
         Vend = txtCampo(4).Value
         Prazo = txtCampo(5).Value
         CifFob = txtCampo(6).Value
         Vr_do_Frete = txtCampo(9).Value
         Vr_do_Desconto = txtCampo(10).Value
         Obs = txtCampo(30).Value
         Total_dos_Produtos = txtCampo(13).Value
         Total_das_Despesas = txtCampo(14).Value
         Total_do_IPI = txtCampo(11).Value
         Total_do_ICMS = txtCampo(12).Value
         Total_do_Pedido = txtCampo(15).Value
         Endereco_de_Entrega = txtCampo(16).Value
         Numero_do_Endereco = txtCampo(17).Value
         Bairro_de_Entrega = txtCampo(18).Value
         Cidade_de_Entrega = txtCampo(19).Value
         UF_De_Entrega = txtCampo(20).Value
         CEP_de_Entrega = txtCampo(21).Value
         Fone_de_Entrega = txtCampo(22).Value
         Contato_de_Entrega = txtCampo(23).Value
         Previsao_de_Entrega = txtCampo(31).Value
         Pedido_Fechado = chkCampo(0).Value
         Validado = Iif(vgSituacao = ACAO_INCLUINDO,False, vgTb!Validado)
         Cancelado = chkCampo(1).Value
         Nome_do_Banco_1 = txtCampo(32).Value
         Agencia_do_Banco_1 = txtCampo(33).Value
         Conta_Corrente_do_Banco_1 = txtCampo(34).Value
         Nome_do_Correntista_do_Banco_1 = txtCampo(35).Value
         Prepedido = chkCampo(2).Value
         Sequencia_Grupo_Despesa = txtCampo(28).Value
         Sequencia_SubGrupo_Despesa = txtCampo(29).Value
         Codigo_da_Licitacao = Iif(vgSituacao = ACAO_INCLUINDO,0, vgTb![Codigo da Licitação])
         Justificar_o_Atraso = txtCampo(37).Value
         Nova_Previsao = txtCampo(38).Value
         Dias = txtCampo(36).Value
      Else
         Id_do_Pedido = IIf(IsNull(vgTb![Id do Pedido]), 0, vgTb![Id do Pedido])
         Data_do_Pedido = vgTb![Data do Pedido]
         Nro_da_Licitacao = IIf(IsNull(vgTb![Nro da Licitação]), "", vgTb![Nro da Licitação])
         Codigo_do_Fornecedor = IIf(IsNull(vgTb![Codigo do Fornecedor]), 0, vgTb![Codigo do Fornecedor])
         Codigo_da_Transportadora = IIf(IsNull(vgTb![Codigo da Transportadora]), 0, vgTb![Codigo da Transportadora])
         Comprador = IIf(IsNull(vgTb!Comprador), "", vgTb!Comprador)
         Vend = IIf(IsNull(vgTb!Vend), "", vgTb!Vend)
         Prazo = IIf(IsNull(vgTb!Prazo), "", vgTb!Prazo)
         CifFob = IIf(IsNull(vgTb!CifFob), "", vgTb!CifFob)
         Vr_do_Frete = IIf(IsNull(vgTb![Vr do Frete]), 0, vgTb![Vr do Frete])
         Vr_do_Desconto = IIf(IsNull(vgTb![Vr do Desconto]), 0, vgTb![Vr do Desconto])
         Obs = IIf(IsNull(vgTb!Obs), "", vgTb!Obs)
         Total_dos_Produtos = IIf(IsNull(vgTb![Total dos Produtos]), 0, vgTb![Total dos Produtos])
         Total_das_Despesas = IIf(IsNull(vgTb![Total das Despesas]), 0, vgTb![Total das Despesas])
         Total_do_IPI = IIf(IsNull(vgTb![Total do IPI]), 0, vgTb![Total do IPI])
         Total_do_ICMS = IIf(IsNull(vgTb![Total do ICMS]), 0, vgTb![Total do ICMS])
         Total_do_Pedido = IIf(IsNull(vgTb![Total do Pedido]), 0, vgTb![Total do Pedido])
         Endereco_de_Entrega = IIf(IsNull(vgTb![Endereco de Entrega]), "", vgTb![Endereco de Entrega])
         Numero_do_Endereco = IIf(IsNull(vgTb![Numero do Endereco]), "", vgTb![Numero do Endereco])
         Bairro_de_Entrega = IIf(IsNull(vgTb![Bairro de Entrega]), "", vgTb![Bairro de Entrega])
         Cidade_de_Entrega = IIf(IsNull(vgTb![Cidade de Entrega]), "", vgTb![Cidade de Entrega])
         UF_De_Entrega = IIf(IsNull(vgTb![UF De Entrega]), "", vgTb![UF De Entrega])
         CEP_de_Entrega = IIf(IsNull(vgTb![CEP de Entrega]), "", vgTb![CEP de Entrega])
         Fone_de_Entrega = IIf(IsNull(vgTb![Fone de Entrega]), "", vgTb![Fone de Entrega])
         Contato_de_Entrega = IIf(IsNull(vgTb![Contato de Entrega]), "", vgTb![Contato de Entrega])
         Previsao_de_Entrega = vgTb![Previsao de Entrega]
         Pedido_Fechado = IIf(IsNull(vgTb![Pedido Fechado]), 0, vgTb![Pedido Fechado])
         Validado = IIf(IsNull(vgTb!Validado), 0, vgTb!Validado)
         Cancelado = IIf(IsNull(vgTb!Cancelado), 0, vgTb!Cancelado)
         Nome_do_Banco_1 = IIf(IsNull(vgTb![Nome do Banco 1]), "", vgTb![Nome do Banco 1])
         Agencia_do_Banco_1 = IIf(IsNull(vgTb![Agência do Banco 1]), "", vgTb![Agência do Banco 1])
         Conta_Corrente_do_Banco_1 = IIf(IsNull(vgTb![Conta Corrente do Banco 1]), "", vgTb![Conta Corrente do Banco 1])
         Nome_do_Correntista_do_Banco_1 = IIf(IsNull(vgTb![Nome do Correntista do Banco 1]), "", vgTb![Nome do Correntista do Banco 1])
         Prepedido = IIf(IsNull(vgTb!Prepedido), 0, vgTb!Prepedido)
         Sequencia_Grupo_Despesa = IIf(IsNull(vgTb![Seqüência Grupo Despesa]), 0, vgTb![Seqüência Grupo Despesa])
         Sequencia_SubGrupo_Despesa = IIf(IsNull(vgTb![Seqüência SubGrupo Despesa]), 0, vgTb![Seqüência SubGrupo Despesa])
         Codigo_da_Licitacao = IIf(IsNull(vgTb![Codigo da Licitação]), 0, vgTb![Codigo da Licitação])
         Justificar_o_Atraso = IIf(IsNull(vgTb![Justificar o Atraso]), "", vgTb![Justificar o Atraso])
         Nova_Previsao = vgTb![Nova Previsao]
         Dias = IIf(IsNull(vgTb!Dias), 0, vgTb!Dias)
      End If
      NotaFiscal = txtCampo(24).Value
      NotaDespesa = txtCampo(25).Value
      NotaConsumo = txtCampo(26).Value
      Tipo_da_Licitacao = txtCampo(27).Value
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
      i = (vgPWUsuario = "YGOR") and Permitido(vgIdentTab, ACAO_EXCLUINDO)
      If Err Then Err.Clear: i = False
      vgTemExclusao = i
      grdBrowse.AllowDelete = i
      .botExclui.Enabled = i
      .Menu_Exclui.Enabled = i
      i = (Pedido_Fechado = 0  Or vgPWUsuario = "YGOR" Or vgPWUsuario = "JERONIMO" Or vgPWUsuario = "CADU" Or vgPWUsuario = "JUCELI") and Permitido(vgIdentTab, ACAO_EDITANDO)
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
      vgOk = ((Not Vazio(Data_do_Pedido)) And (IsDate(Data_do_Pedido) Or Vazio(Data_do_Pedido)))
      vgMsg$ = "Data do Pedido não pode ser vazio!"
      If Not vgOk Then vgColumn = 2
      If vgOk Then
         vgOk = (((Not Vazio(Prazo)) AND Prazo = "Antecipado") OR Prazo = "Prazo")
         vgMsg$ = "Informe Prazo ou Antecipado!"
         If Not vgOk Then vgColumn = 6
      End If
      If vgOk Then
         vgOk = (Not Vazio(CifFob))
         vgMsg$ = "Escolha o Tipo do Frete ( CIF ou Fob )"
         If Not vgOk Then vgColumn = 7
      End If
      If vgOk Then
         If vgPWUsuario = "KALINE" Or vgPWUsuario = "GISELE" Or vgPWUsuario = "KENNEDY" Then
            vgOk = (Codigo_do_Fornecedor > 0)
            vgMsg$ = "Codigo do Fornecedor inválido!"
            If Not vgOk Then vgColumn = 8
         End If
      End If
      If vgOk Then
         vgOk = (Codigo_da_Transportadora > 0)
         vgMsg$ = "Codigo da Transportadora inválido!"
         If Not vgOk Then vgColumn = 9
      End If
      If vgOk Then
         vgOk = (Not Vazio(Nome_do_Banco_1))
         vgMsg$ = "Nome do banco 1 não pode ser vazio!"
         If Not vgOk Then vgColumn = 33
      End If
      If vgOk Then
         vgOk = ((Not Vazio(Previsao_de_Entrega)) And (IsDate(Previsao_de_Entrega) Or Vazio(Previsao_de_Entrega)))
         vgMsg$ = "Previsao de Entrega não pode ser vazio!"
         If Not vgOk Then vgColumn = 32
      End If
      If vgOk Then
         vgOk = (Sequencia_Grupo_Despesa > 0)
         vgMsg$ = "Grupo da Despesa inválido!"
         If Not vgOk Then vgColumn = 29
      End If
      If vgOk Then
         vgOk = (Sequencia_SubGrupo_Despesa > 0)
         vgMsg$ = "SubGrupo Despesa inválido!"
         If Not vgOk Then vgColumn = 30
      End If
      If vgOk Then
         vgOk = (IsDate(Nova_Previsao) Or Vazio(Nova_Previsao))
         vgMsg$ = "Nova Previsao não pode ser vazio!"
         If Not vgOk Then vgColumn = 39
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
               If vgRepeticao = False Or (i <> 28 And i <> 29) Then
                  txtCampo(i).Text = ""
               End if
            End If
         Next
         InicializaApelidos COM_TEXTBOX
         On Error Resume Next
         txtCampo(16).Value = "Rod. Sarg. Luciano Arnaldo Covolan - S/N Km 0,315"
         txtCampo(18).Value = "Zona Rural"
         txtCampo(19).Value = "Penápolis"
         txtCampo(20).Value = "SP"
         txtCampo(21).Value = "16300-000"
         txtCampo(22).Value = "(18)3654-2248"
         txtCampo(23).Value = "Wellington"
         chkCampo(0).Value = False
         chkCampo(1).Value = False
         chkCampo(2).Value = False
         On Error Goto DeuErro
         InicializaApelidos COM_TEXTBOX
         PoeRelEFiltroCbo 7
         PoeRelEFiltroCbo 8
         PoeRelEFiltroCbo 28
         PoeRelEFiltroCbo 29
      End If
   ElseIf vgOq = PEGA_DO_ARQUIVO Then
      If vgTb.RecordCount > 0 And vgTb.EOF = False And vgTb.BOF = False Then
         vgPriVez = True
         vgTb.Resync 1             'adAffectCurrent
         InicializaApelidos COM_REGISTRO
         PoeRelEFiltroCbo 7
         PoeRelEFiltroCbo 8
         PoeRelEFiltroCbo 28
         PoeRelEFiltroCbo 29
         For i = 0 To UBound(txtCampo)
            If Len(txtCampo(i).DataField) > 0 Then
               txtCampo(i).SetOriginalValue = True
               txtCampo(i).Value = vgTb.Fields(txtCampo(i).DataField).Value
            End If
         Next
         chkCampo(0).Value = Pedido_Fechado
         chkCampo(1).Value = Cancelado
         chkCampo(2).Value = Prepedido
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
      vgTb![Pedido Fechado] = Pedido_Fechado
      Validado = IIf(IsNull(vgTb!Validado), 0, vgTb!Validado)
      vgTb!Cancelado = Cancelado
      vgTb!Prepedido = Prepedido
      Codigo_da_Licitacao = IIf(IsNull(vgTb![Codigo da Licitação]), 0, vgTb![Codigo da Licitação])
   ElseIf vgOq = INI_APELIDOS Then
      InicializaApelidos COM_REGISTRO
      ExecutaVisivel
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
   ElseIf vgOq = APOS_EDICAO Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If Abs(vgSituacao) = ACAO_INCLUINDO Then
         AjustaValores
      ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
         AjustaValores
      End If
   ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
      On Error GoTo DeuErro
      InicializaApelidos COM_REGISTRO
      If vgOq = EXCLUSOES Then
         CErr.NumErro = 0
         ExcluiGrid 0
         If CErr.NumErro <> 0 Then GoTo DeuErro
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
Public Sub LigaDesligaRepeticao()
   vgRepeticao = Not vgRepeticao
   MostraFormulas
   ExecutaVisivel
   ExecutaPreValidacao
   PrepBotoes Me, vgSituacao                          'acerta icones dos botoes
End Sub

'liga/desliga flag de repetição do último reg visualizado
Public Sub LigaDesligaAlterar()
   vgAlterar = Not vgAlterar          
   vgUltAlterar = vgAlterar                            'guarda situação de "pode alterar"
   AnalisaCondicoes                                    'vamos atualizar as condições para inclusão, exclusão, alteração...
   MostraFormulas
   ExecutaVisivel
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
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then vgPegouSeq = False
   Set Pedido_de_Compra_Novo = vgTb
   If vgSituacao <> ACAO_INCLUINDO And vgSituacao <> ACAO_EDITANDO Then Executar PEGA_DO_ARQUIVO
   If vgSituacao = ACAO_NAVEGANDO Then
      AnalisaCondicoes 
   End If
   On Error Resume Next
   For i = 0 To 7
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
         Case 4
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(4).CloseRecordset
            Else
               x$ = ExecutaGrid(4, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(4).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(4).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(4).CloseRecordSet
                  End If
                  Grid(4).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(4, vgCols(), CONDICOES_ESPECIAIS)
         Case 5
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(5).CloseRecordset
            Else
               x$ = ExecutaGrid(5, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(5).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(5).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(5).CloseRecordSet
                  End If
                  Grid(5).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(5, vgCols(), CONDICOES_ESPECIAIS)
         Case 6
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(6).CloseRecordset
            Else
               x$ = ExecutaGrid(6, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(6).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(6).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(6).CloseRecordSet
                  End If
                  Grid(6).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(6, vgCols(), CONDICOES_ESPECIAIS)
         Case 7
            If vgSituacao = ACAO_INCLUINDO Or vgTb.EOF Or vgTb.BOF Or (vgSituacao <> ACAO_EXCLUINDO And vgEmBrowse)          Then
               Grid(7).CloseRecordset
            Else
               x$ = ExecutaGrid(7, vgCols(), ABRE_TABELA_GRID)
               MudouFiltro = (x$ <> Grid(7).SQLSource)
               If Err = 0 And (ForceRebind Or MudouFiltro) And Grid(7).Status = ACAO_NAVEGANDO Then
                  If Len(Grid(0).Recordset.RsSource) > 0 Then
                     Grid(7).CloseRecordSet
                  End If
                  Grid(7).OpenRecordSet x$, CURSOR_TABLE
               End If
            End If
               x$ = ExecutaGrid(7, vgCols(), CONDICOES_ESPECIAIS)
      End Select
   Next
   RepositionPedido  
   ExecutaVisivel
   ExecutaPreValidacao
   MostraFormulas
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
   Ok = (vgPWUsuario = "KALINE" Or vgPWUsuario = "GISELE" Or vgPWUsuario = "KENNEDY")
   txtCampo(7).Locked = Not (vgAlterar And txtCampo(7).Editable)
   txtCampo(7).Enabled = Ok Or Not vgAlterar
   Ok = (Pedido_Fechado = False And Produtos_do_Pedido_Compra.RecordCount > 0)
   Botao(0).Enabled = Ok
   Ok = (Pedido_Fechado = False And Produtos_do_Pedido_Compra.RecordCount > 0)
   Botao(1).Enabled = Ok
   Ok = (Pedido_Fechado = False And Despesas_do_Pedido_Compra.RecordCount > 0)
   Botao(2).Enabled = Ok
   Ok = (Pedido_Fechado = False And Despesas_do_Pedido_Compra.RecordCount > 0)
   Botao(3).Enabled = Ok
   Ok = (Pedido_Fechado = False And Consumo_do_Pedido_Compra.RecordCount > 0)
   Botao(4).Enabled = Ok
   Ok = (Pedido_Fechado = False And Consumo_do_Pedido_Compra.RecordCount > 0)
   Botao(5).Enabled = Ok
   Ok = ((Sequencia_Grupo_Despesa > 0) AND Sequencia_SubGrupo_Despesa > 0)
   txtCampo(27).Locked = Not (vgAlterar And txtCampo(27).Editable)
   txtCampo(27).Enabled = Ok Or Not vgAlterar
   Ok = (vgSituacao = ACAO_NAVEGANDO)
   Botao(6).Enabled = Ok
   Ok = (vgSituacao = ACAO_NAVEGANDO And TemConta)
   Botao(7).Enabled = Ok
   Ok = (Prepedido = False)
   Botao(8).Enabled = Ok
   Ok = (VgSituacao <> ACAO_INCLUINDO)
   chkCampo(1).Locked = Not (vgAlterar And chkCampo(1).Editable)
   chkCampo(1).Enabled = Ok Or Not vgAlterar
   Ok = (False)
   txtCampo(37).Locked = Not (vgAlterar And txtCampo(37).Editable)
   txtCampo(37).Enabled = Ok Or Not vgAlterar
   If Err Then Err.Clear                          'se houve erro, limpa...
   vgPriVez = vgPV
End Sub

'coloca os campos visíveis segundo a condição
Private Sub ExecutaVisivel()
   On Error Resume Next                           'prepara para possiveis erros
   Label(36).Visible = ((Previsao_de_Entrega < Date) AND Pedido_Fechado = 0)
   Label(37).Visible = ((TemConta) AND Id_do_Pedido > 0)
   chkCampo(2).Visible = (Prepedido)
   txtCampo(37).Visible = ((Previsao_de_Entrega < Date) AND Pedido_Fechado = 0)
   If Err Then Err.Clear                          'se houve erro, limpa...
End Sub

'evento - quando o conteúdo do campo for alterado
Private Sub txtCp_Change(Index As Integer)
   If vgPriVez Or txtCampo(Index).PriVez Then Exit Sub
   If Len(txtCampo(Index).DataField) > 0 Then LigaFocos Me
   InicializaApelidos COM_TEXTBOX                         'inicializa apelidos com o que esta sendo digitado
   txtCampo(Index).Change
   If Index = 6 Or Index = 7 Or Index = 8 Or Index = 20 Or Index = 27 Or Index = 28 Or Index = 29 Then
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
         Case 7
            PoeRelEFiltroCbo 7
         Case 8
            PoeRelEFiltroCbo 8
         Case 16
            If Len(txtCp(16).Text) = 0 Then
               txtCampo(16).Value = "Rod. Sarg. Luciano Arnaldo Covolan - S/N Km 0,315"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 18
            If Len(txtCp(18).Text) = 0 Then
               txtCampo(18).Value = "Zona Rural"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 19
            If Len(txtCp(19).Text) = 0 Then
               txtCampo(19).Value = "Penápolis"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 20
            If Len(txtCp(20).Text) = 0 Then
               txtCampo(20).Value = "SP"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 21
            If Len(txtCp(21).Text) = 0 Then
               txtCampo(21).Value = "16300-000"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 22
            If Len(txtCp(22).Text) = 0 Then
               txtCampo(22).Value = "(18)3654-2248"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 23
            If Len(txtCp(23).Text) = 0 Then
               txtCampo(23).Value = "Wellington"
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
         Case 28
            PoeRelEFiltroCbo 28
         Case 29
            PoeRelEFiltroCbo 29
         Case 38
            If Len(txtCp(38).Text) = 0 Then
               txtCampo(38).Value = Date
               txtCp_Change Index
               InicializaApelidos COM_TEXTBOX
               ExecutaVisivel
               ExecutaPreValidacao
               MostraFormulas
            End If
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
      Case 36
         AtuPrevisao  
   End Select
End Sub


'evento - quando o check for marcado/desmarcado
Private Sub chkCp_Click(Index As Integer)
   If vgPriVez Then Exit Sub
   If chkCampo(Index).Locked Then
      If Index = 0 Then
         chkCampo(0).Value = Pedido_Fechado
      ElseIf Index = 1 Then
         chkCampo(1).Value = Cancelado

      ElseIf Index = 2 Then
         chkCampo(2).Value = Prepedido
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
   Caption = LoadGasString(45700)
   vgFormID = 1161
   vgIdentTab = "Pedido de Compra Novo"
   vgFiltroEmUso = -1
   vgIndDefault = "Id do Pedido"
   vgPriVez = True
   vgTipo = TP_TABELA
   vgTemInclusao = True
   vgTemExclusao = True
   vgTemAlteracao = True
   vgTemProcura = True
   vgTemFiltro = True
   vgTemBrowse = True
   grdBrowse.Tag = 1
   vgRepeticao = False
   vgAlterar = False
   vgUltAlterar = False
   vgCaracteristica = F_DADOS
   vgUltimoTabIndex = 58
   vgSituacao = ACAO_NAVEGANDO
   Set picBox(0).Picture = LoadPicture(LoadGasPicture(37205))
   Set Botao(8).Picture = LoadPicture(LoadGasPicture(37206))
   Set lblPedido = Label(0)
   Set Aba1 = Tabs(0)
   Set GridProdutos = Grid(0)
   Set GridDespesas = Grid(1)
   Set GridBxProduto = Grid(2)
   Set Txtnotaprod = txtCampo(24)
   Set GridBxDespesa = Grid(3)
   Set Txtnotadespesa = txtCampo(25)
   Set GridBxConsumo = Grid(5)
   Set lblParcelamento = Label(26)
   Set GridParcelas = Grid(6)
   Set txtTipo_da_Licitacao = txtCampo(27)
   Set btnGerarconta = Botao(6)
   Set btnExtornar = Botao(7)
   Set Txtprevisao = txtCampo(31)
   Set Lblfin = Label(37)
   Set Txtdias = txtCampo(36)
   Set lblAjuste = Label(43)
   Set vgTb = New GRecordSet
   If Len(vgFiltroInicial$) > 0 Then
      vgFiltroInicial$ = vgFiltroInicial$ + " And "
   End If
   vgFiltroInicial$ = vgFiltroInicial$ + "([Id do Pedido] > " & 0 & " AND [Data do Pedido] >= CONVERT(VARCHAR, '" & Format$(DateAdd("D", -180, Date), "yyyy-mm-dd hh:mm:ss") & "', 120))"
   vgFiltroOriginal$ = vgFiltroInicial$
   DefineControles
   vgTooltips.Create
   Label(0).Caption = LoadGasString(45705)
   Label(1).Caption = LoadGasString(45706)
   vgTooltips.AddTool txtCampo(0).CtPri, 0, LoadGasString(45707)
   Label(2).Caption = LoadGasString(45708)
   vgTooltips.AddTool txtCampo(1).CtPri, 0, LoadGasString(45709)
   Label(3).Caption = LoadGasString(45710)
   vgTooltips.AddTool txtCampo(2).CtPri, 0, LoadGasString(45711)
   Label(4).Caption = LoadGasString(45712)
   vgTooltips.AddTool txtCampo(3).CtPri, 0, LoadGasString(45713)
   Label(5).Caption = LoadGasString(45714)
   vgTooltips.AddTool txtCampo(4).CtPri, 0, LoadGasString(45715)
   Label(6).Caption = LoadGasString(45716)
   vgTooltips.AddTool txtCampo(5).CtPri, 0, LoadGasString(45717)
   Label(7).Caption = LoadGasString(45718)
   vgTooltips.AddTool txtCampo(6).CtPri, 0, LoadGasString(45719)
   Label(8).Caption = LoadGasString(45720)
   vgTooltips.AddTool txtCampo(7).CtPri, 0, LoadGasString(45721)
   Label(9).Caption = LoadGasString(45722)
   vgTooltips.AddTool txtCampo(8).CtPri, 0, LoadGasString(45723)
   Tabs(0).TabCaption(0) = LoadGasString(45724)
   Tabs(0).TabCaption(1) = LoadGasString(45725)
   Tabs(0).TabCaption(2) = LoadGasString(45726)
   Tabs(0).TabCaption(3) = LoadGasString(45727)
   Tabs(0).TabCaption(4) = LoadGasString(45728)
   Tabs(0).TabCaption(5) = LoadGasString(45729)
   Tabs(0).TabCaption(6) = LoadGasString(45730)
   Tabs(0).TabCaption(7) = LoadGasString(45731)
   Label(10).Caption = LoadGasString(45732)
   vgTooltips.AddTool txtCampo(9).CtPri, 0, LoadGasString(45733)
   Label(11).Caption = LoadGasString(45734)
   vgTooltips.AddTool txtCampo(10).CtPri, 0, LoadGasString(45735)
   Label(12).Caption = LoadGasString(45736)
   vgTooltips.AddTool txtCampo(11).CtPri, 0, LoadGasString(45737)
   Label(13).Caption = LoadGasString(45738)
   vgTooltips.AddTool txtCampo(12).CtPri, 0, LoadGasString(45739)
   Label(14).Caption = LoadGasString(45740)
   vgTooltips.AddTool txtCampo(13).CtPri, 0, LoadGasString(45741)
   Label(15).Caption = LoadGasString(45742)
   vgTooltips.AddTool txtCampo(14).CtPri, 0, LoadGasString(45743)
   Label(16).Caption = LoadGasString(45744)
   vgTooltips.AddTool txtCampo(15).CtPri, 0, LoadGasString(45745)
   Label(17).Caption = LoadGasString(45746)
   vgTooltips.AddTool txtCampo(16).CtPri, 0, LoadGasString(45747)
   Label(18).Caption = LoadGasString(45748)
   vgTooltips.AddTool txtCampo(17).CtPri, 0, LoadGasString(45749)
   Label(19).Caption = LoadGasString(45750)
   vgTooltips.AddTool txtCampo(18).CtPri, 0, LoadGasString(45751)
   Label(20).Caption = LoadGasString(45752)
   vgTooltips.AddTool txtCampo(19).CtPri, 0, LoadGasString(45753)
   Label(21).Caption = LoadGasString(45754)
   vgTooltips.AddTool txtCampo(20).CtPri, 0, LoadGasString(45755)
   Label(22).Caption = LoadGasString(45756)
   vgTooltips.AddTool txtCampo(21).CtPri, 0, LoadGasString(45757)
   Label(23).Caption = LoadGasString(45758)
   vgTooltips.AddTool txtCampo(22).CtPri, 0, LoadGasString(45759)
   Label(24).Caption = LoadGasString(45760)
   vgTooltips.AddTool txtCampo(23).CtPri, 0, LoadGasString(45761)
   Botao(0).Caption = LoadGasString(45762)
   Botao(1).Caption = LoadGasString(45763)
   labtxtCampo24.Caption = LoadGasString(45764)
   Botao(2).Caption = LoadGasString(45765)
   Botao(3).Caption = LoadGasString(45766)
   labtxtCampo25.Caption = LoadGasString(45767)
   Botao(4).Caption = LoadGasString(45768)
   Botao(5).Caption = LoadGasString(45769)
   Label(25).Caption = LoadGasString(45770)
   Label(27).Caption = LoadGasString(45771)
   Label(28).Caption = LoadGasString(45772)
   Label(29).Caption = LoadGasString(45773)
   Label(30).Caption = LoadGasString(45774)
   Label(31).Caption = LoadGasString(45775)
   vgTooltips.AddTool txtCampo(27).CtPri, 0, LoadGasString(45776)
   Botao(6).Caption = LoadGasString(45777)
   Label(32).Caption = LoadGasString(45778)
   Label(33).Caption = LoadGasString(45779)
   Botao(7).Caption = LoadGasString(45780)
   Label(34).Caption = LoadGasString(45781)
   vgTooltips.AddTool txtCampo(30).CtPri, 0, LoadGasString(45782)
   Label(35).Caption = LoadGasString(45783)
   vgTooltips.AddTool txtCampo(31).CtPri, 0, LoadGasString(45784)
   Botao(8).Caption = LoadGasString(45785)
   vgTooltips.AddTool chkCampo(0).CtPri, 0, LoadGasString(45786)
   chkCampo(0).Caption = LoadGasString(45787)
   Label(36).Caption = LoadGasString(45788)
   vgTooltips.AddTool chkCampo(1).CtPri, 0, LoadGasString(45789)
   chkCampo(1).Caption = LoadGasString(45790)
   Label(38).Caption = LoadGasString(45791)
   vgTooltips.AddTool txtCampo(32).CtPri, 0, LoadGasString(45792)
   Label(39).Caption = LoadGasString(45793)
   vgTooltips.AddTool txtCampo(33).CtPri, 0, LoadGasString(45794)
   Label(40).Caption = LoadGasString(45795)
   vgTooltips.AddTool txtCampo(34).CtPri, 0, LoadGasString(45796)
   Label(41).Caption = LoadGasString(45797)
   vgTooltips.AddTool txtCampo(35).CtPri, 0, LoadGasString(45798)
   vgTooltips.AddTool chkCampo(2).CtPri, 0, LoadGasString(45799)
   chkCampo(2).Caption = LoadGasString(45800)
   Label(42).Caption = LoadGasString(45801)
   vgTooltips.AddTool txtCampo(37).CtPri, 0, LoadGasString(45802)
   vgTooltips.AddTool txtCampo(38).CtPri, 0, LoadGasString(45803)
   Label(43).Caption = LoadGasString(45804)
   With Grid(0)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Produtos do Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Produtos do Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Produtos do Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Produto", "Id do Produto", TP_NUMERICO, "", , False, , "IRRIGACAO", "Produtos", "Seqüência do Produto", "Seqüência do Produto; Descrição", "Seqüência do Produto; Descrição", "Seqüência do Produto", "", , "1", , , , "0", 1, "0", 4920
      .AddColumn Nothing, , "Seq", "Sequencia do Item", TP_NUMERICO, "999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 375
      .AddColumn Nothing, , "*Qtde", "Qtde", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Ult. Custo", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1320
      .AddColumn Nothing, , "*Valor Unitario", "Vr Unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "Total", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "% IPI", "Aliquota do IPI", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 765
      .AddColumn Nothing, , "% ICMS", "Aliquota do Icms", TP_NUMERICO, "99,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 960
      .AddColumn Nothing, , "Vr do Ipi", "Vr do IPI", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 930
      .AddColumn Nothing, , "Vr do Icms", "Vr do Icms", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1260
      .AddColumn Nothing, , "Orçamento", "Seqüência do Orçamento", TP_NUMERICO, "999999", , False, , "IRRIGACAO", "Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento", "", , "1", , , , "0", 1, "0", 1230
      .AddColumn Nothing, , "Cliente", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 3270
   End With
   With Grid(1)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Despesas do Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Despesas do Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Despesas do Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Despesa", "Id da Despesa", TP_NUMERICO, "", , False, , "IRRIGACAO", "Despesas", "Seqüência da Despesa", "Seqüência da Despesa; Descrição", "Seqüência da Despesa; Descrição", "Seqüência da Despesa", "", , "1", , , , "0", 1, "0", 5145
      .AddColumn Nothing, , "Qtde", "Qtde", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Ult. Custo", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1320
      .AddColumn Nothing, , "Valor Unitario", "Vr Unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1335
      .AddColumn Nothing, , "Total", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1245
      .AddColumn Nothing, , "% IPI", "Aliquota do IPI", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 705
      .AddColumn Nothing, , "% ICMS", "Aliquota do Icms", TP_NUMERICO, "99,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 825
      .AddColumn Nothing, , "Valor do IPI", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Vr. do ICMS", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1365
   End With
   With Grid(2)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Bx Produtos Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Bx Produtos Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Bx Produtos Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Baixa Produto", "Id do Produto", TP_NUMERICO, "", , True, , "IRRIGACAO", "Produtos", "Seqüência do Produto", "Seqüência do Produto; Descrição", "Seqüência do Produto; Descrição", "Seqüência do Produto", "", , "1", , , , "0", 1, "0", 5010
      .AddColumn Nothing, , "Qtde Total", "Qtde Total", TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Vr.Unitario + IPI", "Vr Unitario", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1350
      .AddColumn Nothing, , "Total", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1020
      .AddColumn Nothing, , "* Qtde Recebida", "Qtde Recebida", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1410
      .AddColumn Nothing, , "Qtde Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "$$ Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1335
      .AddColumn Nothing, , "Notas", "Notas", TP_CARACTER, "@x", 100, False, , , , , , , , , , "0", , , , "0", 1, "0", 5145
   End With
   With Grid(3)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Bx Despesas Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Bx Despesas Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Bx Despesas Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Baixa Despesa", "Id da Despesa", TP_NUMERICO, "", , True, , "IRRIGACAO", "Despesas", "Seqüência da Despesa", "Seqüência da Despesa; Descrição", "Seqüência da Despesa; Descrição", "Seqüência da Despesa", "", , "1", , , , "0", 1, "0", 4890
      .AddColumn Nothing, , "Qtde Total", "Qtde Total", TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Vr Unitario", "Vr Unitario", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Total", "Vr Total do Pedido", TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "Qtde Recebida", "Qtde Recebida", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1560
      .AddColumn Nothing, , "Qtde Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1410
      .AddColumn Nothing, , "$$ Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1380
      .AddColumn Nothing, , "Notas", "Notas", TP_CARACTER, "@x", 100, False, , , , , , , , , , "0", , , , "0", 1, "0", 5025
   End With
   With Grid(4)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Consumo do Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Consumo do Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Consumo do Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Material de Consumo", "Id da Despesa", TP_NUMERICO, "", , False, , "IRRIGACAO", "Despesas", "Seqüência da Despesa", "Seqüência da Despesa; Descrição", "Seqüência da Despesa; Descrição", "Seqüência da Despesa", "", , "1", , , , "0", 1, "0", 5205
      .AddColumn Nothing, , "*Qtde", "Qtde", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Ult. Custo", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "*Vr Unitario", "Vr Unitario", TP_NUMERICO, "999.999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Vr Total", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1275
      .AddColumn Nothing, , "% IPI", "Aliquota do IPI", TP_NUMERICO, "999,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 870
      .AddColumn Nothing, , "% ICMS", "Aliquota do Icms", TP_NUMERICO, "99,9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 810
      .AddColumn Nothing, , "Valor do IPI", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Valor do ICMS", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1365
   End With
   With Grid(5)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Bx Consumo Pedido Compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Bx Consumo Pedido Compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Bx Consumo Pedido Compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Material de Consumo", "Id da Despesa", TP_NUMERICO, "", , True, , "IRRIGACAO", "Despesas", "Seqüência da Despesa", "Seqüência da Despesa; Descrição", "Seqüência da Despesa; Descrição", "Seqüência da Despesa", "", , "1", , , , "0", 1, "0", 5265
      .AddColumn Nothing, , "Qtde Total", "Qtde Total", TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1140
      .AddColumn Nothing, , "Vr Unitario", "Vr Unitario", TP_NUMERICO, "999.999,9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1275
      .AddColumn Nothing, , "Total", "Vr Total do Pedido", TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1395
      .AddColumn Nothing, , "Qtde Recebida", "Qtde Recebida", TP_NUMERICO, "9.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1380
      .AddColumn Nothing, , "Qtde Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1215
      .AddColumn Nothing, , "$$ Restante", , TP_NUMERICO, "9.999.999,99", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1335
      .AddColumn Nothing, , "Notas", "Notas", TP_CARACTER, "@x", 100, False, , , , , , , , , , "0", , , , "0", 1, "0", 3570
   End With
   With Grid(6)
      .RowHeight = 285
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Parcelas Ped Compra Novo", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Parcelas Ped Compra Novo", ACAO_EDITANDO)
      .AllowDelete = Permitido("Parcelas Ped Compra Novo", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº Pc.", "Número da Parcela", TP_NUMERICO, "9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 645
      .AddColumn Nothing, , "Dias", "Dias", TP_NUMERICO, "9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 540
      .AddColumn Nothing, , "Vencimento", "Data de Vencimento", TP_DATA_HORA, "99/99/9999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1425
      .AddColumn Nothing, , "Vr. Parcela", "Valor da Parcela", TP_NUMERICO, "99.999.999,99", , False, , , , , , , , , , "0", , , , "0", 1, "0", 1440
      .AddColumn Nothing, , "Cobrança", "Seqüência da Cobrança", TP_CARACTER, "@x", , False, , "IRRIGACAO", "Tipo de Cobrança", "Seqüência da Cobrança", "Descrição", "Descrição", "Descrição", "", , "1", , , , "0", 1, "0", 1365
      .AddColumn Nothing, , "Nota", "Nota", TP_NUMERICO, "999999999", , False, , , , , , , , , , "0", , , , "0", 1, "0", 960
   End With
   With Grid(7)
      .RowHeight = 315
      .AddControlIgnoreFocus mdiIRRIG.botCancela.hWnd           'não deixa o grid tentar gravar automaticamente
      .AddControlIgnoreFocus mdiIRRIG.botSalva.hWnd             'se estiver perdendo o foco para esses botões
      .FullRowSelect = False
      .BorderStyle = 1
      .NavigationAddMode = 1
      .CacheSize = 100
      .SpecialPopupDisabled POP_GRID_COLS Or POP_GRID_BARS Or POP_GRID_STRIPES
      .AllowInsert = Permitido("Orçamentos da compra", ACAO_INCLUINDO)
      .AllowEdit= Permitido("Orçamentos da compra", ACAO_EDITANDO)
      .AllowDelete = Permitido("Orçamentos da compra", ACAO_EXCLUINDO)
      .AddColumn Nothing, , "Nº", "Sequencia do Item", TP_NUMERICO, "999999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 420
      .AddColumn Nothing, , "Orçamento", "Seqüência do Orçamento", TP_NUMERICO, "999999", , False, , "IRRIGACAO", "Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento", "Seqüência do Orçamento-", "", , "1", , , , "0", 1, "0", 930
      .AddColumn Nothing, , "Emissão", , TP_DATA_HORA, "99/99/9999", , True, , , , , , , , , , "0", , , , "0", 1, "0", 1170
      .AddColumn Nothing, , "Cliente", , TP_CARACTER, , , True, , , , , , , , , , "0", , , , "0", 1, "0", 5625
   End With
   AjustaTamanho Me
   IniciaFormDados Me
   If vgTb.RecordCount > 0 Then vgTb.MoveLast
   Set Pedido_de_Compra_Novo = vgTb
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
   txtCampo(2).Mask = "@x"
   txtCampo(2).BoundColumn = ""
   txtCampo(2).ListFields = ""
   txtCampo(2).OrderFields = ""
   txtCampo(2).Relation = ""
   txtCampo(2).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(2).DataField), txtCampo(2)

   Set txtCampo(3).CtPri = txtCp(3)
   txtCampo(3).DataType = 0
   txtCampo(3).Mask = "@x"
   txtCampo(3).BoundColumn = ""
   txtCampo(3).ListFields = ""
   txtCampo(3).OrderFields = ""
   txtCampo(3).Relation = ""
   txtCampo(3).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(3).DataField), txtCampo(3)

   Set txtCampo(4).CtPri = txtCp(4)
   txtCampo(4).DataType = 0
   txtCampo(4).Mask = "@x"
   txtCampo(4).BoundColumn = ""
   txtCampo(4).ListFields = ""
   txtCampo(4).OrderFields = ""
   txtCampo(4).Relation = ""
   txtCampo(4).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(4).DataField), txtCampo(4)

   Set txtCampo(5).CtPri = txtCp(5)
   txtCampo(5).DataType = 0
   txtCampo(5).Mask = "@x"
   txtCampo(5).BoundColumn = ""
   txtCampo(5).ListFields = ""
   txtCampo(5).OrderFields = ""
   txtCampo(5).Relation = ""
   txtCampo(5).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(5).DataField), txtCampo(5)

   Set txtCampo(6).CtPri = txtCp(6)
   Set txtCampo(6).CtFdo = labFdo6
   Set txtCampo(6).CtBot(BOT_COMBO) = bottxtCampo6(BOT_COMBO)
   Set bottxtCampo6(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(6).DataType = 0
   txtCampo(6).ListFields = "CIF|FOB"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(6).DataField), txtCampo(6)

   Set txtCampo(7).CtPri = txtCp(7)
   Set txtCampo(7).CtFdo = labFdo7
   Set txtCampo(7).CtBot(BOT_COMBO) = bottxtCampo7(BOT_COMBO)
   Set bottxtCampo7(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(7).DataType = 1
   txtCampo(7).Mask = "999999"
   txtCampo(7).BoundColumn = "Seqüência do Geral"
   txtCampo(7).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(7).OrderFields = "Razão Social"
   txtCampo(7).Relation = ""
   txtCampo(7).Source = "Geral"
   txtCampo(7).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(7).DataField), txtCampo(7)

   Set txtCampo(8).CtPri = txtCp(8)
   Set txtCampo(8).CtFdo = labFdo8
   Set txtCampo(8).CtBot(BOT_COMBO) = bottxtCampo8(BOT_COMBO)
   Set bottxtCampo8(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(8).DataType = 1
   txtCampo(8).Mask = "999999"
   txtCampo(8).BoundColumn = "Seqüência do Geral"
   txtCampo(8).ListFields = "Razão Social; Seqüência do Geral"
   txtCampo(8).OrderFields = "Razão Social"
   txtCampo(8).Relation = ""
   txtCampo(8).Source = "Geral"
   txtCampo(8).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(8).DataField), txtCampo(8)

   Set txtCampo(32).CtPri = txtCp(32)
   txtCampo(32).DataType = 0
   txtCampo(32).Mask = "@x"
   txtCampo(32).BoundColumn = ""
   txtCampo(32).ListFields = ""
   txtCampo(32).OrderFields = ""
   txtCampo(32).Relation = ""
   txtCampo(32).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(32).DataField), txtCampo(32)

   Set txtCampo(33).CtPri = txtCp(33)
   txtCampo(33).DataType = 0
   txtCampo(33).Mask = "@x"
   txtCampo(33).BoundColumn = ""
   txtCampo(33).ListFields = ""
   txtCampo(33).OrderFields = ""
   txtCampo(33).Relation = ""
   txtCampo(33).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(33).DataField), txtCampo(33)

   Set txtCampo(34).CtPri = txtCp(34)
   txtCampo(34).DataType = 0
   txtCampo(34).Mask = "@x"
   txtCampo(34).BoundColumn = ""
   txtCampo(34).ListFields = ""
   txtCampo(34).OrderFields = ""
   txtCampo(34).Relation = ""
   txtCampo(34).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(34).DataField), txtCampo(34)

   Set txtCampo(35).CtPri = txtCp(35)
   txtCampo(35).DataType = 0
   txtCampo(35).Mask = "@x"
   txtCampo(35).BoundColumn = ""
   txtCampo(35).ListFields = ""
   txtCampo(35).OrderFields = ""
   txtCampo(35).Relation = ""
   txtCampo(35).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(35).DataField), txtCampo(35)

   Set txtCampo(30).CtPri = txtCp(30)
   txtCampo(30).DataType = 4
   txtCampo(30).Mask = ""
   txtCampo(30).BoundColumn = ""
   txtCampo(30).ListFields = ""
   txtCampo(30).OrderFields = ""
   txtCampo(30).Relation = ""
   txtCampo(30).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(30).DataField), txtCampo(30)

   Set txtCampo(9).CtPri = txtCp(9)
   txtCampo(9).DataType = 1
   txtCampo(9).Mask = "9.999.999,99"
   txtCampo(9).BoundColumn = ""
   txtCampo(9).ListFields = ""
   txtCampo(9).OrderFields = ""
   txtCampo(9).Relation = ""
   txtCampo(9).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(9).DataField), txtCampo(9)

   Set txtCampo(10).CtPri = txtCp(10)
   txtCampo(10).DataType = 1
   txtCampo(10).Mask = "9.999.999,99"
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

   Set txtCampo(12).CtPri = txtCp(12)
   txtCampo(12).DataType = 1
   txtCampo(12).Mask = "9.999.999,99"
   txtCampo(12).Editable = False
   txtCampo(12).BoundColumn = ""
   txtCampo(12).ListFields = ""
   txtCampo(12).OrderFields = ""
   txtCampo(12).Relation = ""
   txtCampo(12).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(12).DataField), txtCampo(12)

   Set txtCampo(13).CtPri = txtCp(13)
   txtCampo(13).DataType = 1
   txtCampo(13).Mask = "9.999.999,99"
   txtCampo(13).Editable = False
   txtCampo(13).BoundColumn = ""
   txtCampo(13).ListFields = ""
   txtCampo(13).OrderFields = ""
   txtCampo(13).Relation = ""
   txtCampo(13).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(13).DataField), txtCampo(13)

   Set txtCampo(14).CtPri = txtCp(14)
   txtCampo(14).DataType = 1
   txtCampo(14).Mask = "9.999.999,99"
   txtCampo(14).Editable = False
   txtCampo(14).BoundColumn = ""
   txtCampo(14).ListFields = ""
   txtCampo(14).OrderFields = ""
   txtCampo(14).Relation = ""
   txtCampo(14).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(14).DataField), txtCampo(14)

   Set txtCampo(15).CtPri = txtCp(15)
   txtCampo(15).DataType = 1
   txtCampo(15).Mask = "9.999.999,99"
   txtCampo(15).Editable = False
   txtCampo(15).BoundColumn = ""
   txtCampo(15).ListFields = ""
   txtCampo(15).OrderFields = ""
   txtCampo(15).Relation = ""
   txtCampo(15).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(15).DataField), txtCampo(15)

   Set txtCampo(16).CtPri = txtCp(16)
   txtCampo(16).DataType = 0
   txtCampo(16).Mask = "@x"
   txtCampo(16).BoundColumn = ""
   txtCampo(16).ListFields = ""
   txtCampo(16).OrderFields = ""
   txtCampo(16).Relation = ""
   txtCampo(16).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(16).DataField), txtCampo(16)

   Set txtCampo(17).CtPri = txtCp(17)
   txtCampo(17).DataType = 0
   txtCampo(17).Mask = "@!"
   txtCampo(17).BoundColumn = ""
   txtCampo(17).ListFields = ""
   txtCampo(17).OrderFields = ""
   txtCampo(17).Relation = ""
   txtCampo(17).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(17).DataField), txtCampo(17)

   Set txtCampo(18).CtPri = txtCp(18)
   txtCampo(18).DataType = 0
   txtCampo(18).Mask = "@x"
   txtCampo(18).BoundColumn = ""
   txtCampo(18).ListFields = ""
   txtCampo(18).OrderFields = ""
   txtCampo(18).Relation = ""
   txtCampo(18).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(18).DataField), txtCampo(18)

   Set txtCampo(19).CtPri = txtCp(19)
   txtCampo(19).DataType = 0
   txtCampo(19).Mask = "@x"
   txtCampo(19).BoundColumn = ""
   txtCampo(19).ListFields = ""
   txtCampo(19).OrderFields = ""
   txtCampo(19).Relation = ""
   txtCampo(19).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(19).DataField), txtCampo(19)

   Set txtCampo(36).CtPri = txtCp(36)
   txtCampo(36).DataType = 1
   txtCampo(36).Mask = "9999"
   txtCampo(36).BoundColumn = ""
   txtCampo(36).ListFields = ""
   txtCampo(36).OrderFields = ""
   txtCampo(36).Relation = ""
   txtCampo(36).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(36).DataField), txtCampo(36)

   Set txtCampo(31).CtPri = txtCp(31)
   txtCampo(31).DataType = 2
   txtCampo(31).Mask = "99/99/9999"
   txtCampo(31).BoundColumn = ""
   txtCampo(31).ListFields = ""
   txtCampo(31).OrderFields = ""
   txtCampo(31).Relation = ""
   txtCampo(31).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(31).DataField), txtCampo(31)

   Set txtCampo(20).CtPri = txtCp(20)
   Set txtCampo(20).CtFdo = labFdo20
   Set txtCampo(20).CtBot(BOT_COMBO) = bottxtCampo20(BOT_COMBO)
   Set bottxtCampo20(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(20).DataType = 0
   txtCampo(20).ListFields = "AC|AL|AP|AM|BA|CE|DF|ES|TO|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RN|RS|RJ|RO|RR|SC|SP|SE|EX"
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(20).DataField), txtCampo(20)

   Set txtCampo(21).CtPri = txtCp(21)
   txtCampo(21).DataType = 0
   txtCampo(21).Mask = "99999-999"
   txtCampo(21).BoundColumn = ""
   txtCampo(21).ListFields = ""
   txtCampo(21).OrderFields = ""
   txtCampo(21).Relation = ""
   txtCampo(21).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(21).DataField), txtCampo(21)

   Set txtCampo(22).CtPri = txtCp(22)
   txtCampo(22).DataType = 0
   txtCampo(22).Mask = "@!"
   txtCampo(22).BoundColumn = ""
   txtCampo(22).ListFields = ""
   txtCampo(22).OrderFields = ""
   txtCampo(22).Relation = ""
   txtCampo(22).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(22).DataField), txtCampo(22)

   Set txtCampo(23).CtPri = txtCp(23)
   txtCampo(23).DataType = 0
   txtCampo(23).Mask = "@x"
   txtCampo(23).BoundColumn = ""
   txtCampo(23).ListFields = ""
   txtCampo(23).OrderFields = ""
   txtCampo(23).Relation = ""
   txtCampo(23).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(23).DataField), txtCampo(23)

   Set chkCampo(0).CtPri = chkCp(0)
   chkCampo(0).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(0).DataField), chkCampo(0)

   Set txtCampo(24).CtTit = labtxtCampo24
   Set txtCampo(24).CtPri = txtCp(24)
   Set txtCampo(24).CtFdo = labFdo24
   Set txtCampo(24).CtBot(BOT_ACAO) = bottxtCampo24(BOT_ACAO)
   Set bottxtCampo24(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(24).DataType = 0
   txtCampo(24).Mask = ""
   txtCampo(24).BoundColumn = ""
   txtCampo(24).ListFields = ""
   txtCampo(24).OrderFields = ""
   txtCampo(24).Relation = ""
   txtCampo(24).Source = ""

   Set txtCampo(25).CtTit = labtxtCampo25
   Set txtCampo(25).CtPri = txtCp(25)
   Set txtCampo(25).CtFdo = labFdo25
   Set txtCampo(25).CtBot(BOT_ACAO) = bottxtCampo25(BOT_ACAO)
   Set bottxtCampo25(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(25).DataType = 0
   txtCampo(25).Mask = ""
   txtCampo(25).BoundColumn = ""
   txtCampo(25).ListFields = ""
   txtCampo(25).OrderFields = ""
   txtCampo(25).Relation = ""
   txtCampo(25).Source = ""

   Set txtCampo(26).CtPri = txtCp(26)
   Set txtCampo(26).CtFdo = labFdo26
   Set txtCampo(26).CtBot(BOT_ACAO) = bottxtCampo26(BOT_ACAO)
   Set bottxtCampo26(BOT_ACAO).Picture = LoadPicture(LoadGasPicture(4))
   txtCampo(26).DataType = 0
   txtCampo(26).Mask = ""
   txtCampo(26).BoundColumn = ""
   txtCampo(26).ListFields = ""
   txtCampo(26).OrderFields = ""
   txtCampo(26).Relation = ""
   txtCampo(26).Source = ""

   Set txtCampo(27).CtPri = txtCp(27)
   Set txtCampo(27).CtFdo = labFdo27
   Set txtCampo(27).CtBot(BOT_COMBO) = bottxtCampo27(BOT_COMBO)
   Set bottxtCampo27(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(27).DataType = 0
   txtCampo(27).ListFields = "MPrima|MRevenda|MConsumo|Despesas|Ativo"

 DefineControles1

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles0"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub
Public Sub DefineControles1()
 On Error GoTo DeuErro

   Set chkCampo(1).CtPri = chkCp(1)
   chkCampo(1).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(1).DataField), chkCampo(1)

   Set chkCampo(2).CtPri = chkCp(2)
   chkCampo(2).DataType = 5
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(chkCampo(2).DataField), chkCampo(2)

   Set txtCampo(28).CtPri = txtCp(28)
   Set txtCampo(28).CtFdo = labFdo28
   Set txtCampo(28).CtBot(BOT_LISTA) = bottxtCampo28(BOT_LISTA)
   Set txtCampo(28).CtBot(BOT_COMBO) = bottxtCampo28(BOT_COMBO)
   bottxtCampo28(BOT_LISTA).Caption = "P"
   Set bottxtCampo28(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(28).DataType = 1
   txtCampo(28).Mask = "9999"
   txtCampo(28).PesqModoAbertura =  2 
   txtCampo(28).PesqFieldCapture = "[Grupo da Despesa].[Seqüência Grupo Despesa]"
   txtCampo(28).BoundColumn = "Seqüência Grupo Despesa"
   txtCampo(28).ListFields = "Descrição; Seqüência Grupo Despesa"
   txtCampo(28).OrderFields = "Descrição; Seqüência Grupo Despesa"
   txtCampo(28).Relation = ""
   txtCampo(28).Source = "Grupo da Despesa"
   txtCampo(28).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(28).DataField), txtCampo(28)

   Set txtCampo(29).CtPri = txtCp(29)
   Set txtCampo(29).CtFdo = labFdo29
   Set txtCampo(29).CtBot(BOT_LISTA) = bottxtCampo29(BOT_LISTA)
   Set txtCampo(29).CtBot(BOT_COMBO) = bottxtCampo29(BOT_COMBO)
   bottxtCampo29(BOT_LISTA).Caption = "P"
   Set bottxtCampo29(BOT_COMBO).Picture = LoadPicture(LoadGasPicture(3))
   txtCampo(29).DataType = 1
   txtCampo(29).Mask = "9999"
   txtCampo(29).PesqModoAbertura =  2 
   txtCampo(29).PesqFieldCapture = "[SubGrupo Despesa].[Seqüência SubGrupo Despesa]"
   txtCampo(29).BoundColumn = "Seqüência SubGrupo Despesa"
   txtCampo(29).ListFields = "Descrição; Seqüência SubGrupo Despesa"
   txtCampo(29).OrderFields = "Descrição; Seqüência SubGrupo Despesa"
   txtCampo(29).Relation = ""
   txtCampo(29).Source = "SubGrupo Despesa"
   txtCampo(29).vgfrmGMCale.grdListaG.AutoRebind = True
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(29).DataField), txtCampo(29)

   Set txtCampo(37).CtPri = txtCp(37)
   txtCampo(37).DataType = 0
   txtCampo(37).Mask = "@x"
   txtCampo(37).BoundColumn = ""
   txtCampo(37).ListFields = ""
   txtCampo(37).OrderFields = ""
   txtCampo(37).Relation = ""
   txtCampo(37).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(37).DataField), txtCampo(37)

   Set txtCampo(38).CtPri = txtCp(38)
   txtCampo(38).DataType = 2
   txtCampo(38).Mask = "99/99/9999"
   txtCampo(38).BoundColumn = ""
   txtCampo(38).ListFields = ""
   txtCampo(38).OrderFields = ""
   txtCampo(38).Relation = ""
   txtCampo(38).Source = ""
   grdBrowse.AddColumn vgDb.Tables(vgIdentTab).Columns(txtCampo(38).DataField), txtCampo(38)

 Exit Sub

 DeuErro:
  CErr.NumErro = Err
  CErr.FunctionName = "DefineControles1"
  CErr.Origem = CStr(vgFormID) +" - "+ Me.Caption
 CErr.Show
End Sub


'poe relacionamento e filtro na lista externa (combobox)
Private Sub PoeRelEFiltroCbo(Index As Integer)
   On Error Resume Next
   Select Case Index
      Case 7
         txtCampo(7).Filter = "Fornecedor = " & 1 & ""
      Case 8
         txtCampo(8).Filter = "Transportadora = " & 1 & ""
      Case 28
         txtCampo(28).Filter = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "Inativo = 0", "[Seqüência Grupo Despesa] > 0") & ""
                           txtCampo(28).PesqSQLExpression = "SELECT [Grupo da Despesa].[Seqüência Grupo Despesa], [Grupo da Despesa].Descrição FROM [Grupo da Despesa] WHERE ([Grupo da Despesa].[Seqüência Grupo Despesa] > " + CStr(0) + ") AND " + _
                                                               "([Grupo da Despesa].Inativo = False)"
      Case 29
         txtCampo(29).Filter = "([Seqüência Grupo Despesa] = " & Sequencia_Grupo_Despesa & ") AND [Seqüência SubGrupo Despesa] > " & 0 & ""
                           txtCampo(29).PesqSQLExpression = "SELECT [SubGrupo Despesa].[Seqüência SubGrupo Despesa], [SubGrupo Despesa].Descrição FROM [SubGrupo Despesa] WHERE ([SubGrupo Despesa].[Seqüência Grupo Despesa] = " + CStr([Sequencia_Grupo_Despesa]) + ") AND " + _
                                                               "([SubGrupo Despesa].[Seqüência SubGrupo Despesa] > " + CStr(0) + ")"
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
   VerificaFormulario  
   FinalizaForm Me
   Set lblPedido = Nothing
   Set Aba1 = Nothing
   Set GridProdutos = Nothing
   Set GridDespesas = Nothing
   Set GridBxProduto = Nothing
   Set Txtnotaprod = Nothing
   Set GridBxDespesa = Nothing
   Set Txtnotadespesa = Nothing
   Set GridBxConsumo = Nothing
   Set lblParcelamento = Nothing
   Set GridParcelas = Nothing
   Set txtTipo_da_Licitacao = Nothing
   Set btnGerarconta = Nothing
   Set btnExtornar = Nothing
   Set Txtprevisao = Nothing
   Set Lblfin = Nothing
   Set Txtdias = Nothing
   Set lblAjuste = Nothing
   For i = 0 To UBound(txtCampo)
      txtCampo(i).Finalize
      Set txtCampo(i) = Nothing
   Next
   Set chkCampo(0) = Nothing
   Set chkCampo(1) = Nothing
   Set chkCampo(2) = Nothing
   If Not Pedido_de_Compra_Novo Is Nothing Then
      Set Pedido_de_Compra_Novo = Nothing
   End If
   If Not Produtos_do_Pedido_Compra Is Nothing Then
      Produtos_do_Pedido_Compra.CloseRecordSet
      Set Produtos_do_Pedido_Compra = Nothing
   End If
   If Not Despesas_do_Pedido_Compra Is Nothing Then
      Despesas_do_Pedido_Compra.CloseRecordSet
      Set Despesas_do_Pedido_Compra = Nothing
   End If
   If Not Bx_Produtos_Pedido_Compra Is Nothing Then
      Bx_Produtos_Pedido_Compra.CloseRecordSet
      Set Bx_Produtos_Pedido_Compra = Nothing
   End If
   If Not Bx_Despesas_Pedido_Compra Is Nothing Then
      Bx_Despesas_Pedido_Compra.CloseRecordSet
      Set Bx_Despesas_Pedido_Compra = Nothing
   End If
   If Not Consumo_do_Pedido_Compra Is Nothing Then
      Consumo_do_Pedido_Compra.CloseRecordSet
      Set Consumo_do_Pedido_Compra = Nothing
   End If
   If Not Bx_Consumo_Pedido_Compra Is Nothing Then
      Bx_Consumo_Pedido_Compra.CloseRecordSet
      Set Bx_Consumo_Pedido_Compra = Nothing
   End If
   If Not Parcelas_Ped_Compra_Novo Is Nothing Then
      Parcelas_Ped_Compra_Novo.CloseRecordSet
      Set Parcelas_Ped_Compra_Novo = Nothing
   End If
   If Not Orcamentos_da_compra Is Nothing Then
      Orcamentos_da_compra.CloseRecordSet
      Set Orcamentos_da_compra = Nothing
   End If

   'se tiver imprimindo registros em grade, fecha form de selecao/preview
   UnLoad vgFrmImpCons
   Set vgFrmImpCons = Nothing

   'vamos descarregar todos os grids
   For i = 0 To Grid.Count - 1
      Grid(i).Finalize
   Next

   grdBrowse.Finalize                             'descarrega o grid
   Set frmPComprN = Nothing                       'libera o segmento de código do form
End Sub

'trata a exclusão dentro dos grid filhos
Private Sub ExcluiGrid(vgNumGrid As Integer)
   Grid(vgNumGrid).DeleteRecords True, False
End Sub

Private Sub vgTb_BeforeBeginTrans()
   Dim i As Integer
   If vgSituacao = ACAO_EXCLUINDO And vgDb(1).TransactionCount = 0 Then
      For i = 0 To Grid.Count - 1
         Grid(i).FreezeListView True                     'vamos fixar a imagem
      Next
      For i = 0 To Grid.Count - 1
         Grid(i).CloseRecordset
      Next
   End If
End Sub

Private Sub vgTb_AfterBeginTrans()
   If vgSituacao = ACAO_EXCLUINDO And vgDb(1).TransactionCount = 1 Then
      Reposition True
   End If
End Sub

Private Sub vgTb_BeforeCommitTrans()
   Dim i As Integer
   If vgSituacao = ACAO_EXCLUINDO And vgDb(1).TransactionCount = 1 Then
      For i = 0 To Grid.Count - 1
         Grid(i).CloseRecordset
      Next
   End If
End Sub

Private Sub vgTb_AfterCommitTrans()
   Dim i As Integer
   If vgDb(1).TransactionCount = 0 Then
      For i = 0 To Grid.Count - 1
         Grid(i).RestoreListView True
      Next
   End If
End Sub

Private Sub vgTb_BeforeRollBackTrans()
   Dim i As Integer
   If vgSituacao = ACAO_EXCLUINDO And vgDb(1).TransactionCount = 1 Then
      For i = 0 To Grid.Count - 1
         Grid(i).CloseRecordset
      Next
   End If
End Sub

Private Sub vgTb_AfterRollBackTrans()
   Dim i As Integer
   If vgDb(1).TransactionCount = 0 Then
      For i = 0 To Grid.Count - 1
         Grid(i).RestoreListView True                    'restaura a imagem
      Next
   End If
End Sub

'evento - está mudando a linha selecionada do grid
Private Sub Grid_SkipRecord(Index As Integer, vgColumns() As Variant, ByVal vgBookMark As Variant)
   Dim vgSqlGrd As String
   ExecutaGrid Index, vgColumns(), CONDICOES_ESPECIAIS
   Select Case Index
      Case 0
         vgSQLGrd = ExecutaGrid0(vgColumns(), REPOSITION_NO_GRID)
      Case 1
         vgSQLGrd = ExecutaGrid1(vgColumns(), REPOSITION_NO_GRID)
      Case 2
         vgSQLGrd = ExecutaGrid2(vgColumns(), REPOSITION_NO_GRID)
      Case 3
         vgSQLGrd = ExecutaGrid3(vgColumns(), REPOSITION_NO_GRID)
      Case 4
         vgSQLGrd = ExecutaGrid4(vgColumns(), REPOSITION_NO_GRID)
      Case 5
         vgSQLGrd = ExecutaGrid5(vgColumns(), REPOSITION_NO_GRID)
      Case 6
         vgSQLGrd = ExecutaGrid6(vgColumns(), REPOSITION_NO_GRID)
      Case 7
         vgSQLGrd = ExecutaGrid7(vgColumns(), REPOSITION_NO_GRID)
   End Select
End Sub

'evento - após efetuar update do recordset do grid
Private Sub Grid_AfterUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), PROCESSOS_DIRETOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 1
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 2
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 3
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 4
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 5
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   GeraLog Me, Grid(Index).Status, Index, False
      Case 6
   GeraLog Me, Grid(Index).Status, Index, False
      Case 7
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
      Case 4
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 5
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 6
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
      Case 7
         ExecutaGrid Index, Grid(Index).GetColumnValues(vgItem), PROCESSOS_INVERSOS, vgItem, 0, vgIsValid, vgColumn, vgErrorMessage
   End Select
End Sub

'evento - antes de efetuar o update do recordset do grid
Private Sub Grid_BeforeUpdateRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   Select Case Index
      Case 0
         Produtos_do_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 1
         Despesas_do_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 2
         Bx_Produtos_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 3
         Bx_Despesas_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 4
         If Grid(Index).Status = ACAO_INCLUINDO Then
            Consumo_do_Pedido_Compra![Id Despesa] = PegaSequencia("Consumo do Pedido Compra", "Id Despesa", 0, 1, Id_do_Pedido)
         End If
         Consumo_do_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 5
         Bx_Consumo_Pedido_Compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 6
         Parcelas_Ped_Compra_Novo![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
      Case 7
         Orcamentos_da_compra![Id do Pedido] = Pedido_de_Compra_Novo![Id do Pedido]
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
      Case 4
      Case 5
      Case 6
      Case 7
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
         Set Produtos_do_Pedido_Compra = vgNewRecordSet
      Case 1
         Set Despesas_do_Pedido_Compra = vgNewRecordSet
      Case 2
         Set Bx_Produtos_Pedido_Compra = vgNewRecordSet
      Case 3
         Set Bx_Despesas_Pedido_Compra = vgNewRecordSet
      Case 4
         Set Consumo_do_Pedido_Compra = vgNewRecordSet
      Case 5
         Set Bx_Consumo_Pedido_Compra = vgNewRecordSet
      Case 6
         Set Parcelas_Ped_Compra_Novo = vgNewRecordSet
      Case 7
         Set Orcamentos_da_compra = vgNewRecordSet
   End Select
End Sub

'evento - quer validar dados, está gravando
Private Sub Grid_ValidateData(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   ExecutaGrid Index, vgColumns(), VALIDACOES, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
End Sub

'evento - pega definição de cores segundo condições
Private Sub Grid_GetColor(Index As Integer, ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Select Case Index
      Case 0
         PegaCoresGrid0 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
      Case 1
         PegaCoresGrid1 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
      Case 4
         PegaCoresGrid4 vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
   End Select
End Sub

'evento - após a exclusao de um registro no grid filho
Private Sub Grid_AfterDeleteRecord(Index As Integer, ByVal vgItem As Long, vgColumns() As Variant, vgIsValid As Boolean, vgColumn As Integer, vgErrorMessage As String)
   GeraLog Me, ACAO_EXCLUINDO, Index, False
   Select Case Index
      Case 0
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 1
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 2
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
      Case 4
         ExecutaGrid Index, vgColumns(), APOS_EDICAO, vgItem, vgColumn, vgIsValid, vgColumn, vgErrorMessage
   End Select
   mdiIRRIG.RemontaForm                                   'vamos atualizar os forms de dados
   Grid(Index).Repaint -1                                 'atualiza dados do grid (registro posicionado) 
End Sub

Private Sub Grid_ControlButtonClick(Index As Integer)
   Grid(Index).ShowFilterBar = Not Grid(Index).ShowFilterBar
End Sub

Private Sub Grid_GotFocus(Index As Integer)
   If vgSituacao <> ACAO_NAVEGANDO And Grid(Index).Status = ACAO_NAVEGANDO Then                 'o formulário pai não está em navegação
      mdiIRRIG.SalvaDados                         'salva o resitro em edição
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
   mdiIRRIG.RemontaForm                                           'remonta dos os form da tela
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
      Case  4 
         ExecutaGrid = ExecutaGrid4(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case  5 
         ExecutaGrid = ExecutaGrid5(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case  6 
         ExecutaGrid = ExecutaGrid6(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
      Case  7 
         ExecutaGrid = ExecutaGrid7(ColumnValue(), vgOq, vgItem, vgCol, vgIsValid, vgColumn, vgErrorMessage, KeyCodeAscii, Shift)
   End Select
End Function

'inicializações, validações e processos do grid filho
Private Function ExecutaGrid0(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Id_do_Produto As Long, Qtde As Double, Vr_Unitario As Double
   Dim Vr_Total As Double, Aliquota_do_IPI As Double, Aliquota_do_Icms As Double
   Dim Vr_do_IPI As Double, Vr_do_Icms As Double, Sequencia_do_Item As Long
   Dim Sequencia_do_Orcamento As Long
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
         Case 4
            vgRetVal = (InfoProdutos(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento, "CUSTO"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 6
            vgRetVal = (Round((Qtde * Vr_Unitario),2))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 12
            vgRetVal = (MGeral(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
      Select Case vgCol
         Case 10
            vgRetVal = Not (False)
         Case Else
            vgRetVal = False
      End Select
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(0).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(0).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(0).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Produtos do Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 2
            Sequencia_do_Item = SeqItem1()
            vgRetVal = Sequencia_do_Item
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqüência do Produto] > 0 AND Inativo = 0", "[Seqüência do Produto] > 0") & ""
         Case 11
            vgRetVal = "(([Seqüência do Orçamento] > " & 0 & ") AND Cancelado = " & 0 & ") AND [Venda Fechada] = " & 1 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         If Id_do_Produto > 0 Then
            vgIsValid = (ValidaProduto(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento))
            If Not vgIsValid Then vgColumn =  1 
            vgErrorMessage$ = "Impossivel Produto Inativo!"
         End If
         If vgIsValid And vgCol = -1 Then
            If Id_do_Produto = 0 Then
               vgIsValid = (Id_do_Produto > 0)
               If Not vgIsValid Then vgColumn =  1 
               vgErrorMessage$ = "Sequencia do Produto deve ser Maior que Zero(0)!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Sequencia_do_Item > 0)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Sequencia do item inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Qtde > 0)
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Qtde inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Vr_Unitario > 0)
            If Not vgIsValid Then vgColumn =  5 
            vgErrorMessage$ = "Vr Unitario inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Vr_do_IPI >= 0)
            If Not vgIsValid Then vgColumn =  9 
            vgErrorMessage$ = "Vr do Ipi inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            If False Then
               vgIsValid = (Vr_do_Icms > 0)
               If Not vgIsValid Then vgColumn =  10 
               vgErrorMessage$ = "Vr do Icms inválido!"
            End If
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Sequencia_do_Orcamento >= 0)
            If Not vgIsValid Then vgColumn =  11 
            vgErrorMessage$ = "Seqüência do Orçamento inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaValoresProdutos Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento 
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaValoresProdutos Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores  
         End If
      ElseIf vgOq = PROCESSOS_DIRETOS Then
         GoSub IniApDaCol
         Pedido_de_Compra_Novo.Edit
         Set vgRsError = Pedido_de_Compra_Novo
         If ValorCusto(Id_do_Produto, True) Then
            Pedido_de_Compra_Novo![Id do Pedido] = (0)
         End If
         If ValorTotalConjunto(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento) Then
            Pedido_de_Compra_Novo![Id do Pedido] = (0)
         End If
         Pedido_de_Compra_Novo.Update
         Set vgRsError = Nothing
      ElseIf vgOq = PROCESSOS_INVERSOS Or vgOq = EXCLUSOES Then
         On Error GoTo DeuErro
         GoSub IniApDaTb
         Pedido_de_Compra_Novo.Edit
         Set vgRsError = Pedido_de_Compra_Novo
         If ValorCusto(Id_do_Produto, True) Then
            Pedido_de_Compra_Novo![Id do Pedido] = (0)
         End If
         If ValorTotalConjunto(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento) Then
            Pedido_de_Compra_Novo![Id do Pedido] = (0)
         End If
         Pedido_de_Compra_Novo.Update
         Set vgRsError = Nothing
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Sequencia_do_Item = ColumnValue(2)
   Qtde = ColumnValue(3)
   Vr_Unitario = ColumnValue(5)
   Aliquota_do_IPI = ColumnValue(7)
   Aliquota_do_Icms = ColumnValue(8)
   Vr_do_IPI = ColumnValue(9)
   Vr_do_Icms = ColumnValue(10)
   Sequencia_do_Orcamento = Val(Parse$(CStr(ColumnValue(11)), Chr$(1), 1))
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Pedido_Compra.Eof = False And Produtos_do_Pedido_Compra.BOF = False And Produtos_do_Pedido_Compra.RecordCount > 0 Then
         Vr_Total = Produtos_do_Pedido_Compra![Vr Total]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_do_Pedido_Compra.Eof = False And Produtos_do_Pedido_Compra.BOF = False And Produtos_do_Pedido_Compra.RecordCount > 0 Then
      Id_do_Produto = Produtos_do_Pedido_Compra![Id do Produto]
      Qtde = Produtos_do_Pedido_Compra!Qtde
      Vr_Unitario = Produtos_do_Pedido_Compra![Vr Unitario]
      Vr_Total = Produtos_do_Pedido_Compra![Vr Total]
      Aliquota_do_IPI = Produtos_do_Pedido_Compra![Aliquota do IPI]
      Aliquota_do_Icms = Produtos_do_Pedido_Compra![Aliquota do Icms]
      Vr_do_IPI = Produtos_do_Pedido_Compra![Vr do IPI]
      Vr_do_Icms = Produtos_do_Pedido_Compra![Vr do Icms]
      Sequencia_do_Item = Produtos_do_Pedido_Compra![Sequencia do Item]
      Sequencia_do_Orcamento = Produtos_do_Pedido_Compra![Seqüência do Orçamento]
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
   If Pedido_de_Compra_Novo.EOF = False And Pedido_de_Compra_Novo.BOF = False Then
      Pedido_de_Compra_Novo.CancelUpdate
   End If

FimDaSub:
   ExecutaGrid0 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid1(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Id_da_Despesa As Long, Qtde As Double, Vr_Unitario As Double
   Dim Vr_Total As Double, Aliquota_do_IPI As Double, Aliquota_do_Icms As Double
   Dim Vr_do_IPI As Double, Vr_do_Icms As Double
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
         Case 3
            vgRetVal = (InfoDespesas(Id_do_Pedido, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 5
            vgRetVal = (Qtde * Vr_Unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 8
            vgRetVal = (((Qtde * Vr_Unitario) *  Aliquota_do_IPI) / 100)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 9
            vgRetVal = (((Vr_Unitario * Qtde + ((Qtde * Vr_Unitario) * Aliquota_do_IPI) / 100 ) * Aliquota_do_Icms ) / 100)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(1).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Despesas_Pedido_Compra.RecordCount = 0 And Not TemBaixa())
      On Error Resume Next
         Grid(1).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Despesas_Pedido_Compra.RecordCount = 0 And Not TemBaixa())
      On Error Resume Next
         Grid(1).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Despesas_Pedido_Compra.RecordCount = 0 And Not TemBaixa())
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Despesas do Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Id_da_Despesa > 0)
         If Not vgIsValid Then vgColumn =  1 
         vgErrorMessage$ = "Id da Despesa inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Qtde > 0)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Qtde inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Vr_Unitario > 0)
            If Not vgIsValid Then vgColumn =  4 
            vgErrorMessage$ = "Vr Unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaDespesas Id_do_Pedido, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms 
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaDespesas Id_do_Pedido, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores  
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde = ColumnValue(2)
   Vr_Unitario = ColumnValue(4)
   Aliquota_do_IPI = ColumnValue(6)
   Aliquota_do_Icms = ColumnValue(7)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Despesas_do_Pedido_Compra.Eof = False And Despesas_do_Pedido_Compra.BOF = False And Despesas_do_Pedido_Compra.RecordCount > 0 Then
         Vr_Total = Despesas_do_Pedido_Compra![Vr Total]
         Vr_do_IPI = Despesas_do_Pedido_Compra![Vr do IPI]
         Vr_do_Icms = Despesas_do_Pedido_Compra![Vr do Icms]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Despesas_do_Pedido_Compra.Eof = False And Despesas_do_Pedido_Compra.BOF = False And Despesas_do_Pedido_Compra.RecordCount > 0 Then
      Id_da_Despesa = Despesas_do_Pedido_Compra![Id da Despesa]
      Qtde = Despesas_do_Pedido_Compra!Qtde
      Vr_Unitario = Despesas_do_Pedido_Compra![Vr Unitario]
      Vr_Total = Despesas_do_Pedido_Compra![Vr Total]
      Aliquota_do_IPI = Despesas_do_Pedido_Compra![Aliquota do IPI]
      Aliquota_do_Icms = Despesas_do_Pedido_Compra![Aliquota do Icms]
      Vr_do_IPI = Despesas_do_Pedido_Compra![Vr do IPI]
      Vr_do_Icms = Despesas_do_Pedido_Compra![Vr do Icms]
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
   Dim Id_do_Produto As Long, Qtde_Total As Double, Vr_Unitario As Double
   Dim Vr_Total_do_Pedido As Double, Qtde_Recebida As Double, Qtde_Restante As Double
   Dim Total_Restante As Double, Notas As String, Sequencia_do_Item As Long
   Dim Teste As String
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
         Case 4
            vgRetVal = (Qtde_Total * Vr_Unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 6
            vgRetVal = (Qtde_Total - Qtde_Recebida)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = ((Qtde_Total * Vr_Unitario ) - (Qtde_Recebida * Vr_Unitario))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(2).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(2).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(2).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Bx Produtos Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Qtde_Recebida <= Qtde_Total)
         If Not vgIsValid Then vgColumn =  5 
         vgErrorMessage$ = "Qtde Recebida Deve Ser Menor ou Igual a Quantidade Total do Pedido"
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaBxProdutos Id_do_Pedido, Id_do_Produto, Qtde_Total, Vr_Unitario, Vr_Total_do_Pedido, Qtde_Recebida, _
   Qtde_Restante, Total_Restante, Notas, Sequencia_do_Item, Teste 
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            ContaEstoque Id_do_Produto 
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde_Total = ColumnValue(2)
   Vr_Unitario = ColumnValue(3)
   Qtde_Recebida = ColumnValue(5)
   Notas = ColumnValue(8) & ""
   If Grid(2).Status <> ACAO_INCLUINDO Then
      If Bx_Produtos_Pedido_Compra.Eof = False And Bx_Produtos_Pedido_Compra.BOF = False And Bx_Produtos_Pedido_Compra.RecordCount > 0 Then
         Vr_Total_do_Pedido = Bx_Produtos_Pedido_Compra![Vr Total do Pedido]
         Qtde_Restante = Bx_Produtos_Pedido_Compra![Qtde Restante]
         Total_Restante = Bx_Produtos_Pedido_Compra![Total Restante]
         Sequencia_do_Item = Bx_Produtos_Pedido_Compra![Sequencia do Item]
         Teste = Bx_Produtos_Pedido_Compra!Teste
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Bx_Produtos_Pedido_Compra.Eof = False And Bx_Produtos_Pedido_Compra.BOF = False And Bx_Produtos_Pedido_Compra.RecordCount > 0 Then
      Id_do_Produto = Bx_Produtos_Pedido_Compra![Id do Produto]
      Qtde_Total = Bx_Produtos_Pedido_Compra![Qtde Total]
      Vr_Unitario = Bx_Produtos_Pedido_Compra![Vr Unitario]
      Vr_Total_do_Pedido = Bx_Produtos_Pedido_Compra![Vr Total do Pedido]
      Qtde_Recebida = Bx_Produtos_Pedido_Compra![Qtde Recebida]
      Qtde_Restante = Bx_Produtos_Pedido_Compra![Qtde Restante]
      Total_Restante = Bx_Produtos_Pedido_Compra![Total Restante]
      Notas = Bx_Produtos_Pedido_Compra!Notas
      Sequencia_do_Item = Bx_Produtos_Pedido_Compra![Sequencia do Item]
      Teste = Bx_Produtos_Pedido_Compra!Teste
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
   ExecutaGrid2 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid3(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Id_da_Despesa As Long, Qtde_Total As Double, Vr_Unitario As Double
   Dim Vr_Total_do_Pedido As Double, Qtde_Recebida As Double, Qtde_Restante As Double
   Dim Total_Restante As Double, Notas As String
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
      Select Case vgCol
         Case 6
            vgRetVal = (Qtde_Total - Qtde_Recebida)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = ((Qtde_Total * Vr_Unitario ) - (Qtde_Recebida * Vr_Unitario))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(3).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(3).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(3).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Bx Despesas Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Qtde_Recebida <= Qtde_Total)
         If Not vgIsValid Then vgColumn =  5 
         vgErrorMessage$ = "Qtde Recebida Deve Ser Menor ou Igual a Quantidade Total do Pedido"
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaBxDespesas Id_do_Pedido, Id_da_Despesa, Qtde_Total, Vr_Unitario, Vr_Total_do_Pedido, Qtde_Recebida, _
   Qtde_Restante, Total_Restante, Notas 
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde_Total = ColumnValue(2)
   Vr_Unitario = ColumnValue(3)
   Vr_Total_do_Pedido = ColumnValue(4)
   Qtde_Recebida = ColumnValue(5)
   Notas = ColumnValue(8) & ""
   If Grid(3).Status <> ACAO_INCLUINDO Then
      If Bx_Despesas_Pedido_Compra.Eof = False And Bx_Despesas_Pedido_Compra.BOF = False And Bx_Despesas_Pedido_Compra.RecordCount > 0 Then
         Qtde_Restante = Bx_Despesas_Pedido_Compra![Qtde Restante]
         Total_Restante = Bx_Despesas_Pedido_Compra![Total Restante]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Bx_Despesas_Pedido_Compra.Eof = False And Bx_Despesas_Pedido_Compra.BOF = False And Bx_Despesas_Pedido_Compra.RecordCount > 0 Then
      Id_da_Despesa = Bx_Despesas_Pedido_Compra![Id da Despesa]
      Qtde_Total = Bx_Despesas_Pedido_Compra![Qtde Total]
      Vr_Unitario = Bx_Despesas_Pedido_Compra![Vr Unitario]
      Vr_Total_do_Pedido = Bx_Despesas_Pedido_Compra![Vr Total do Pedido]
      Qtde_Recebida = Bx_Despesas_Pedido_Compra![Qtde Recebida]
      Qtde_Restante = Bx_Despesas_Pedido_Compra![Qtde Restante]
      Total_Restante = Bx_Despesas_Pedido_Compra![Total Restante]
      Notas = Bx_Despesas_Pedido_Compra!Notas
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


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid4(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Id_Despesa As Long, Id_da_Despesa As Long, Qtde As Double
   Dim Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double
   Dim Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(4).Status <> ACAO_NAVEGANDO And vgItem = Grid(4).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 3
            vgRetVal = (InfoConsumo(Id_do_Pedido, Id_Despesa, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, _
   Aliquota_do_IPI, Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 5
            vgRetVal = (Qtde * Vr_Unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 8
            vgRetVal = (((Qtde * Vr_Unitario) *  Aliquota_do_IPI) / 100)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 9
            vgRetVal = (((Vr_Unitario * Qtde + ((Qtde * Vr_Unitario) * Aliquota_do_IPI) / 100 ) * Aliquota_do_Icms ) / 100)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(4).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Consumo_Pedido_Compra.RecordCount = 0)
      On Error Resume Next
         Grid(4).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Consumo_Pedido_Compra.RecordCount = 0)
      On Error Resume Next
         Grid(4).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False And Bx_Consumo_Pedido_Compra.RecordCount = 0)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Consumo do Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Id_da_Despesa > 0)
         If Not vgIsValid Then vgColumn =  1 
         vgErrorMessage$ = "Id da Despesa inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Qtde > 0)
            If Not vgIsValid Then vgColumn =  2 
            vgErrorMessage$ = "Qtde inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Vr_Unitario > 0)
            If Not vgIsValid Then vgColumn =  4 
            vgErrorMessage$ = "Vr Unitario inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_INCLUINDO Then
            AjustaConsumo  
         ElseIf Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaConsumo  
         ElseIf Abs(vgSituacao) = ACAO_EXCLUINDO Then
            AjustaValores  
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde = ColumnValue(2)
   Vr_Unitario = ColumnValue(4)
   Aliquota_do_IPI = ColumnValue(6)
   Aliquota_do_Icms = ColumnValue(7)
   If Grid(4).Status <> ACAO_INCLUINDO Then
      If Consumo_do_Pedido_Compra.Eof = False And Consumo_do_Pedido_Compra.BOF = False And Consumo_do_Pedido_Compra.RecordCount > 0 Then
         Id_Despesa = Consumo_do_Pedido_Compra![Id Despesa]
         Vr_Total = Consumo_do_Pedido_Compra![Vr Total]
         Vr_do_IPI = Consumo_do_Pedido_Compra![Vr do IPI]
         Vr_do_Icms = Consumo_do_Pedido_Compra![Vr do Icms]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Consumo_do_Pedido_Compra.Eof = False And Consumo_do_Pedido_Compra.BOF = False And Consumo_do_Pedido_Compra.RecordCount > 0 Then
      Id_Despesa = Consumo_do_Pedido_Compra![Id Despesa]
      Id_da_Despesa = Consumo_do_Pedido_Compra![Id da Despesa]
      Qtde = Consumo_do_Pedido_Compra!Qtde
      Vr_Unitario = Consumo_do_Pedido_Compra![Vr Unitario]
      Vr_Total = Consumo_do_Pedido_Compra![Vr Total]
      Aliquota_do_IPI = Consumo_do_Pedido_Compra![Aliquota do IPI]
      Aliquota_do_Icms = Consumo_do_Pedido_Compra![Aliquota do Icms]
      Vr_do_IPI = Consumo_do_Pedido_Compra![Vr do IPI]
      Vr_do_Icms = Consumo_do_Pedido_Compra![Vr do Icms]
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
   ExecutaGrid4 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid5(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Id_Despesa As Long, Id_da_Despesa As Long, Qtde_Total As Double
   Dim Vr_Unitario As Double, Vr_Total_do_Pedido As Double, Qtde_Recebida As Double
   Dim Qtde_Restante As Double, Total_Restante As Double, Notas As String
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(5).Status <> ACAO_NAVEGANDO And vgItem = Grid(5).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 6
            vgRetVal = (Qtde_Total - Qtde_Recebida)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 7
            vgRetVal = ((Qtde_Total * Vr_Unitario ) - (Qtde_Recebida * Vr_Unitario))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = PREVALIDACOES Then
      GoSub IniApDaCol
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(5).AllowInsert = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(5).AllowEdit = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      On Error Resume Next
         Grid(5).AllowDelete = (Pedido_de_Compra_Novo![Pedido Fechado] = False)
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Bx Consumo Pedido Compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Qtde_Recebida <= Qtde_Total)
         If Not vgIsValid Then vgColumn =  5 
         vgErrorMessage$ = "Qtde Recebida inválido!"
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = APOS_EDICAO Then
         On Error GoTo DeuErro
         GoSub IniApDaCol
         If Abs(vgSituacao) = ACAO_EDITANDO Then
            AjustaBxConsumo  
         End If
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde_Total = ColumnValue(2)
   Vr_Unitario = ColumnValue(3)
   Vr_Total_do_Pedido = ColumnValue(4)
   Qtde_Recebida = ColumnValue(5)
   Notas = ColumnValue(8) & ""
   If Grid(5).Status <> ACAO_INCLUINDO Then
      If Bx_Consumo_Pedido_Compra.Eof = False And Bx_Consumo_Pedido_Compra.BOF = False And Bx_Consumo_Pedido_Compra.RecordCount > 0 Then
         Id_Despesa = Bx_Consumo_Pedido_Compra![Id Despesa]
         Qtde_Restante = Bx_Consumo_Pedido_Compra![Qtde Restante]
         Total_Restante = Bx_Consumo_Pedido_Compra![Total Restante]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Bx_Consumo_Pedido_Compra.Eof = False And Bx_Consumo_Pedido_Compra.BOF = False And Bx_Consumo_Pedido_Compra.RecordCount > 0 Then
      Id_Despesa = Bx_Consumo_Pedido_Compra![Id Despesa]
      Id_da_Despesa = Bx_Consumo_Pedido_Compra![Id da Despesa]
      Qtde_Total = Bx_Consumo_Pedido_Compra![Qtde Total]
      Vr_Unitario = Bx_Consumo_Pedido_Compra![Vr Unitario]
      Vr_Total_do_Pedido = Bx_Consumo_Pedido_Compra![Vr Total do Pedido]
      Qtde_Recebida = Bx_Consumo_Pedido_Compra![Qtde Recebida]
      Qtde_Restante = Bx_Consumo_Pedido_Compra![Qtde Restante]
      Total_Restante = Bx_Consumo_Pedido_Compra![Total Restante]
      Notas = Bx_Consumo_Pedido_Compra!Notas
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
   ExecutaGrid5 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid6(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Numero_da_Parcela As Integer, Dias As Integer, Data_de_Vencimento As Variant
   Dim Valor_da_Parcela As Double, Nota_ As Long, Sequencia_da_Cobranca As Integer
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(6).Status <> ACAO_NAVEGANDO And vgItem = Grid(6).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      If Err Then Err.Clear
   ElseIf vgOq = REPOSITION_NO_GRID Then
      GoSub IniApDaTb
      MostraFormulas  
      vgRetVal = ""
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(6).AllowInsert = (TotalParcelas < Pedido_de_Compra_Novo![Total do Pedido])
      On Error Resume Next
         Grid(6).AllowEdit = (Not TemBaixa())
      On Error Resume Next
         Grid(6).AllowDelete = (Not TemBaixa())
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Parcelas Ped Compra Novo]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Numero_da_Parcela = UltimaParcela()
            vgRetVal = Numero_da_Parcela
         Case 2
            Dias = DateDiff("D",Pedido_de_Compra_Novo![Previsao de Entrega], Data_de_Vencimento)
            vgRetVal = Dias
         Case 3
            Data_de_Vencimento = DateAdd("d", Dias, Pedido_de_Compra_Novo![Previsao de Entrega])
            vgRetVal = Data_de_Vencimento
         Case 4
            Valor_da_Parcela = Pedido_de_Compra_Novo![Total do Pedido]  - TotalParcelas()
            vgRetVal = Valor_da_Parcela
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 5
            vgRetVal = "([Seqüência da Cobrança] > " & 0 & ") AND Inativo = " & 0 & ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (1 = 1)
         If Not vgIsValid Then vgColumn =  2 
         vgErrorMessage$ = "Dias inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = ((UltimoVencimento) And (IsDate(Data_de_Vencimento) Or Vazio(Data_de_Vencimento)))
            If Not vgIsValid Then vgColumn =  3 
            vgErrorMessage$ = "Vencimento invalido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (IIf(vgSituacao = -ACAO_INCLUINDO, TotalParcelas() + Valor_da_Parcela <= Pedido_de_Compra_Novo![Total do Pedido], Valor_da_Parcela > 0 And (TotalParcelas() - TotalParcelas(Numero_da_Parcela)) + Valor_da_Parcela <= Pedido_de_Compra_Novo![Total do Pedido]))
            If Not vgIsValid Then vgColumn =  4 
            vgErrorMessage$ = "Valor da Parcela não confere com o valor total!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Sequencia_da_Cobranca > 0)
            If Not vgIsValid Then vgColumn =  5 
            vgErrorMessage$ = "Seqüência da Cobrança inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Nota_ >= 0)
            If Not vgIsValid Then vgColumn =  6 
            vgErrorMessage$ = "Nota inválido!"
         End If
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = UltimaParcela()
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Numero_da_Parcela = ColumnValue(1)
   Dias = ColumnValue(2)
   Data_de_Vencimento = ColumnValue(3)
   Valor_da_Parcela = ColumnValue(4)
   Sequencia_da_Cobranca = Val(Parse$(CStr(ColumnValue(5)), Chr$(1), 1))
   Nota_ = ColumnValue(6)
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Parcelas_Ped_Compra_Novo.Eof = False And Parcelas_Ped_Compra_Novo.BOF = False And Parcelas_Ped_Compra_Novo.RecordCount > 0 Then
      Numero_da_Parcela = Parcelas_Ped_Compra_Novo![Número da Parcela]
      Dias = Parcelas_Ped_Compra_Novo!Dias
      Data_de_Vencimento = Parcelas_Ped_Compra_Novo![Data de Vencimento]
      Valor_da_Parcela = Parcelas_Ped_Compra_Novo![Valor da Parcela]
      Nota_ = Parcelas_Ped_Compra_Novo!Nota
      Sequencia_da_Cobranca = Parcelas_Ped_Compra_Novo![Seqüência da Cobrança]
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
   ExecutaGrid6 = vgRetVal
   vgPriVez = False
End Function


'inicializações, validações e processos do grid filho
Private Function ExecutaGrid7(ColumnValue() As Variant, ByVal vgOq As Integer, Optional ByVal vgItem As Long, Optional ByVal vgCol As Integer, Optional vgIsValid As Boolean, Optional ByRef vgColumn As Integer, Optional vgErrorMessage As String, Optional KeyCodeAscii As Integer, Optional Shift As Integer) As Variant
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x as String, vgNVez As Integer
   Dim Sequencia_do_Orcamento As Long, Sequencia_do_Item As Long
   vgPriVez = True
   If vgOq = PREVALIDACOES Then
      vgRetVal = False
   Else
      vgRetVal = ""
   End If
   vgNVez = 0
   On Error GoTo DeuErro
   If vgOq = CONTEUDODACOLUNA Then
      If Grid(7).Status <> ACAO_NAVEGANDO And vgItem = Grid(7).SelectedItem Then
         GoSub IniApDaCol
      Else
         GoSub IniApDaTb
      End If
      On Error Resume Next
      Select Case vgCol
         Case 3
            vgRetVal = (InfoOrc(Id_do_Pedido, Sequencia_do_Orcamento, Sequencia_do_Item, "Emissao"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 4
            vgRetVal = (InfoOrc(Id_do_Pedido, Sequencia_do_Orcamento, Sequencia_do_Item, "Cliente"))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Orçamentos da compra]"

      'definindo a expressão de ligação com o pai
      x$ = "[Id do Pedido] = " & Pedido_de_Compra_Novo![Id do Pedido]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 1
            Sequencia_do_Item = SeqItem()
            vgRetVal = Sequencia_do_Item
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 2
            vgRetVal = "" + FiltroOrc() + ""
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_do_Orcamento > 0)
         If Not vgIsValid Then vgColumn =  2 
         vgErrorMessage$ = "Seqüência do orçamento inválido!"
         If Not vgIsValid And Len(vgErrorMessage$) = 0 Then vgErrorMessage$ = "Err"
      ElseIf vgOq = INICIALIZACOES Then
         GoSub IniApDaCol
         ColumnValue(1) = SeqItem()
      End If
   End If
   GoTo FimDaSub
   Exit Function

IniApDaCol:
   On Error Resume Next
   Sequencia_do_Item = ColumnValue(1)
   Sequencia_do_Orcamento = Val(Parse$(CStr(ColumnValue(2)), Chr$(1), 1))
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Orcamentos_da_compra.Eof = False And Orcamentos_da_compra.BOF = False And Orcamentos_da_compra.RecordCount > 0 Then
      Sequencia_do_Orcamento = Orcamentos_da_compra![Seqüência do Orçamento]
      Sequencia_do_Item = Orcamentos_da_compra![Sequencia do Item]
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
   ExecutaGrid7 = vgRetVal
   vgPriVez = False
End Function

'pega definições de cores para o grid
Private Sub PegaCoresGrid0(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Id_do_Produto As Long, Qtde As Double, Vr_Unitario As Double
   Dim Vr_Total As Double, Aliquota_do_IPI As Double, Aliquota_do_Icms As Double
   Dim Vr_do_IPI As Double, Vr_do_Icms As Double, Sequencia_do_Item As Long
   Dim Sequencia_do_Orcamento As Long
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(0).GetColumnValues(vgItem)

   Id_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Sequencia_do_Item = ColumnValue(2)
   Qtde = ColumnValue(3)
   Vr_Unitario = ColumnValue(5)
   Aliquota_do_IPI = ColumnValue(7)
   Aliquota_do_Icms = ColumnValue(8)
   Vr_do_IPI = ColumnValue(9)
   Vr_do_Icms = ColumnValue(10)
   Sequencia_do_Orcamento = Val(Parse$(CStr(ColumnValue(11)), Chr$(1), 1))
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Pedido_Compra.Eof = False And Produtos_do_Pedido_Compra.BOF = False And Produtos_do_Pedido_Compra.RecordCount > 0 Then
         Vr_Total = Produtos_do_Pedido_Compra![Vr Total]
      End If
   End If

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case  5 

         'Fundo coluna
         If (Vr_Unitario > InfoProdutos(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento, "CUSTO")) Then
            vgBackColor = &H8080FF&
         ElseIf (Vr_Unitario <= InfoProdutos(Id_do_Pedido, Id_do_Produto, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, Sequencia_do_Item, Sequencia_do_Orcamento, "CUSTO")) Then
            vgBackColor = &HC0FFC0&
         End If

   End Select

   Exit Sub

DeuErro:
End Sub

'pega definições de cores para o grid
Private Sub PegaCoresGrid1(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Id_da_Despesa As Long, Qtde As Double, Vr_Unitario As Double
   Dim Vr_Total As Double, Aliquota_do_IPI As Double, Aliquota_do_Icms As Double
   Dim Vr_do_IPI As Double, Vr_do_Icms As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(1).GetColumnValues(vgItem)

   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde = ColumnValue(2)
   Vr_Unitario = ColumnValue(4)
   Aliquota_do_IPI = ColumnValue(6)
   Aliquota_do_Icms = ColumnValue(7)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Despesas_do_Pedido_Compra.Eof = False And Despesas_do_Pedido_Compra.BOF = False And Despesas_do_Pedido_Compra.RecordCount > 0 Then
         Vr_Total = Despesas_do_Pedido_Compra![Vr Total]
         Vr_do_IPI = Despesas_do_Pedido_Compra![Vr do IPI]
         Vr_do_Icms = Despesas_do_Pedido_Compra![Vr do Icms]
      End If
   End If

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case  4 

         'Fundo coluna
         If (Vr_Unitario > InfoDespesas(Id_do_Pedido, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO")) Then
            vgBackColor = &H8080FF&
         ElseIf (Vr_Unitario <= InfoDespesas(Id_do_Pedido, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, Aliquota_do_IPI, _
   Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO")) Then
            vgBackColor = &HC0FFC0&
         End If

   End Select

   Exit Sub

DeuErro:
End Sub

'pega definições de cores para o grid
Private Sub PegaCoresGrid4(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Id_Despesa As Long, Id_da_Despesa As Long, Qtde As Double
   Dim Vr_Unitario As Double, Vr_Total As Double, Aliquota_do_IPI As Double
   Dim Aliquota_do_Icms As Double, Vr_do_IPI As Double, Vr_do_Icms As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(4).GetColumnValues(vgItem)

   Id_da_Despesa = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   Qtde = ColumnValue(2)
   Vr_Unitario = ColumnValue(4)
   Aliquota_do_IPI = ColumnValue(6)
   Aliquota_do_Icms = ColumnValue(7)
   If Grid(4).Status <> ACAO_INCLUINDO Then
      If Consumo_do_Pedido_Compra.Eof = False And Consumo_do_Pedido_Compra.BOF = False And Consumo_do_Pedido_Compra.RecordCount > 0 Then
         Id_Despesa = Consumo_do_Pedido_Compra![Id Despesa]
         Vr_Total = Consumo_do_Pedido_Compra![Vr Total]
         Vr_do_IPI = Consumo_do_Pedido_Compra![Vr do IPI]
         Vr_do_Icms = Consumo_do_Pedido_Compra![Vr do Icms]
      End If
   End If

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case  4 

         'Fundo coluna
         If (Vr_Unitario > InfoConsumo(Id_do_Pedido, Id_Despesa, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, _
   Aliquota_do_IPI, Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO")) Then
            vgBackColor = &H8080FF&
         ElseIf (Vr_Unitario <= InfoConsumo(Id_do_Pedido, Id_Despesa, Id_da_Despesa, Qtde, Vr_Unitario, Vr_Total, _
   Aliquota_do_IPI, Aliquota_do_Icms, Vr_do_IPI, Vr_do_Icms, "CUSTO")) Then
            vgBackColor = &HC0FFC0&
         End If

   End Select

   Exit Sub

DeuErro:
End Sub

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
      Case 4
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 5
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 6
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
      Case 7
         vgLocked = ExecutaGrid(Index, vgColumns(), PREVALIDACOES, , vgCol)
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub Botao_Click(Index As Integer)
   Dim Cancel As Boolean, hMenu As Long, pt As POINTAPI
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 0
         LancaBxProdutoParcial  
      Case 1
         LancaBxProdutoTotal  
      Case 2
         LancaBxDespesaParcial  
      Case 3
         LancaBxDespesaTotal  
      Case 4
         LancaBxConsumoParcial  
      Case 5
         LancaBxConsumoTotal  
      Case 6
         LancaParcela  
      Case 7
         Extornar  
      Case 8
         mdiIRRIG.MRelPComp
         AtivaForm Me
   End Select
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
Private Sub bottxtCampo20_Click(Index As Integer)
   txtCampo(20).SetFocus
   DoEvents
   txtCampo(20).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo24_Click(Index As Integer)
   txtCampo(24).SetFocus
   DoEvents
   txtCampo(24).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo25_Click(Index As Integer)
   txtCampo(25).SetFocus
   DoEvents
   txtCampo(25).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo26_Click(Index As Integer)
   txtCampo(26).SetFocus
   DoEvents
   txtCampo(26).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo27_Click(Index As Integer)
   txtCampo(27).SetFocus
   DoEvents
   txtCampo(27).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo28_Click(Index As Integer)
   txtCampo(28).SetFocus
   DoEvents
   txtCampo(28).botClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo29_Click(Index As Integer)
   txtCampo(29).SetFocus
   DoEvents
   txtCampo(29).botClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x as Single, y as Single)
   txtCampo(Index).MouseDown
End Sub


