VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrcament 
   ClientHeight    =   9615
   ClientLeft      =   2475
   ClientTop       =   1470
   ClientWidth     =   15690
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Orcament"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15690
   Begin IRRIG.GPainel Painel 
      Height          =   9585
      Index           =   0
      Left            =   15
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   15
      Width           =   15660
      _ExtentX        =   27596
      _ExtentY        =   16907
      BevelOuter      =   0
      Stretch         =   -1  'True
      Begin IRRIG.GBotao Botao 
         Height          =   435
         Index           =   16
         Left            =   13125
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   165
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Imprimir"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   300
         Index           =   15
         Left            =   12810
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   735
         Width           =   240
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   435
         Index           =   14
         Left            =   7410
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   615
         Width           =   2580
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Mapa da Propriedade"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   435
         Index           =   13
         Left            =   10050
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   615
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ficha Técnica"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   300
         Index           =   12
         Left            =   14580
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1650
         Width           =   240
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   30
         Index           =   11
         Left            =   15180
         TabIndex        =   59
         Top             =   8445
         Width           =   30
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   300
         Index           =   10
         Left            =   9810
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1650
         Width           =   240
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   435
         Index           =   9
         Left            =   10050
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   165
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Imprimir"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   300
         Index           =   8
         Left            =   14565
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   750
         Width           =   240
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionOffset   =   50
      End
      Begin IRRIG.GBotao Botao 
         Height          =   435
         Index           =   0
         Left            =   13125
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         _ExtentX        =   0
         _ExtentY        =   0
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Gerar Nota"
         CaptionAlign    =   4
         CaptionOffset   =   50
      End
      Begin VB.TextBox txtCp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Seqüência do Orçamento"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   615
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1650
         Width           =   840
      End
      Begin TabDlg.SSTab Tabs 
         Height          =   6570
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Top             =   2040
         Width           =   15165
         _ExtentX        =   26723
         _ExtentY        =   11589
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         TabsPerRow      =   9
         TabHeight       =   519
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&1 - Dados Principais"
         TabPicture(0)   =   "ORCAMENT.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(179)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label(178)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(177)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "labFdo151"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(176)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "labFdo150"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label(175)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "labFdo149"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label(174)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "labFdo147"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label(172)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label(170)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label(169)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label(168)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label(167)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "labFdo137"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label(146)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "labFdo64"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label(68)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label(67)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Label(66)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "labFdo62"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label(65)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "labFdo60"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label(64)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Label(63)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Label(62)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "labFdo59"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Label(61)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Label(60)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "labFdo55"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Label(57)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Label(56)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Label(53)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "Label(52)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Label(51)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Label(46)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Label(45)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "Linha(1)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "Label(44)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "Label(43)"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "Label(42)"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "Label(41)"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "Label(40)"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "Label(37)"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "Label(36)"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "Label(35)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "Label(34)"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "Label(33)"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "Label(32)"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "Label(31)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "Label(16)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "Label(15)"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "Label(14)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "labFdo33"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "Label(6)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "labFdo17"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "Label(4)"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "labFdo16"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "Label(3)"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "Label(2)"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "Label(1)"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "Label(0)"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "txtCp(154)"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "txtCp(153)"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "txtCp(152)"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "txtCp(151)"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "bottxtCampo151(2)"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "txtCp(150)"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "bottxtCampo150(1)"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "txtCp(149)"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "bottxtCampo149(0)"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "txtCp(147)"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "bottxtCampo147(2)"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "txtCp(146)"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "txtCp(145)"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "txtCp(144)"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "txtCp(143)"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "txtCp(137)"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).Control(79)=   "bottxtCampo137(2)"
         Tab(0).Control(79).Enabled=   0   'False
         Tab(0).Control(80)=   "txtCp(65)"
         Tab(0).Control(80).Enabled=   0   'False
         Tab(0).Control(81)=   "txtCp(64)"
         Tab(0).Control(81).Enabled=   0   'False
         Tab(0).Control(82)=   "bottxtCampo64(2)"
         Tab(0).Control(82).Enabled=   0   'False
         Tab(0).Control(83)=   "txtCp(63)"
         Tab(0).Control(83).Enabled=   0   'False
         Tab(0).Control(84)=   "txtCp(62)"
         Tab(0).Control(84).Enabled=   0   'False
         Tab(0).Control(85)=   "bottxtCampo62(2)"
         Tab(0).Control(85).Enabled=   0   'False
         Tab(0).Control(86)=   "txtCp(60)"
         Tab(0).Control(86).Enabled=   0   'False
         Tab(0).Control(87)=   "bottxtCampo60(2)"
         Tab(0).Control(87).Enabled=   0   'False
         Tab(0).Control(88)=   "bottxtCampo60(1)"
         Tab(0).Control(88).Enabled=   0   'False
         Tab(0).Control(89)=   "txtCp(59)"
         Tab(0).Control(89).Enabled=   0   'False
         Tab(0).Control(90)=   "bottxtCampo59(2)"
         Tab(0).Control(90).Enabled=   0   'False
         Tab(0).Control(91)=   "txtCp(58)"
         Tab(0).Control(91).Enabled=   0   'False
         Tab(0).Control(92)=   "txtCp(55)"
         Tab(0).Control(92).Enabled=   0   'False
         Tab(0).Control(93)=   "bottxtCampo55(2)"
         Tab(0).Control(93).Enabled=   0   'False
         Tab(0).Control(94)=   "bottxtCampo55(1)"
         Tab(0).Control(94).Enabled=   0   'False
         Tab(0).Control(95)=   "ChkCp(2)"
         Tab(0).Control(95).Enabled=   0   'False
         Tab(0).Control(96)=   "txtCp(54)"
         Tab(0).Control(96).Enabled=   0   'False
         Tab(0).Control(97)=   "txtCp(51)"
         Tab(0).Control(97).Enabled=   0   'False
         Tab(0).Control(98)=   "Painel(1)"
         Tab(0).Control(98).Enabled=   0   'False
         Tab(0).Control(99)=   "txtCp(50)"
         Tab(0).Control(99).Enabled=   0   'False
         Tab(0).Control(100)=   "txtCp(48)"
         Tab(0).Control(100).Enabled=   0   'False
         Tab(0).Control(101)=   "ChkCp(1)"
         Tab(0).Control(101).Enabled=   0   'False
         Tab(0).Control(102)=   "txtCp(33)"
         Tab(0).Control(102).Enabled=   0   'False
         Tab(0).Control(103)=   "bottxtCampo33(2)"
         Tab(0).Control(103).Enabled=   0   'False
         Tab(0).Control(104)=   "bottxtCampo33(1)"
         Tab(0).Control(104).Enabled=   0   'False
         Tab(0).Control(105)=   "txtCp(30)"
         Tab(0).Control(105).Enabled=   0   'False
         Tab(0).Control(106)=   "txtCp(29)"
         Tab(0).Control(106).Enabled=   0   'False
         Tab(0).Control(107)=   "txtCp(28)"
         Tab(0).Control(107).Enabled=   0   'False
         Tab(0).Control(108)=   "txtCp(27)"
         Tab(0).Control(108).Enabled=   0   'False
         Tab(0).Control(109)=   "txtCp(26)"
         Tab(0).Control(109).Enabled=   0   'False
         Tab(0).Control(110)=   "txtCp(25)"
         Tab(0).Control(110).Enabled=   0   'False
         Tab(0).Control(111)=   "txtCp(24)"
         Tab(0).Control(111).Enabled=   0   'False
         Tab(0).Control(112)=   "txtCp(23)"
         Tab(0).Control(112).Enabled=   0   'False
         Tab(0).Control(113)=   "txtCp(22)"
         Tab(0).Control(113).Enabled=   0   'False
         Tab(0).Control(114)=   "txtCp(21)"
         Tab(0).Control(114).Enabled=   0   'False
         Tab(0).Control(115)=   "txtCp(20)"
         Tab(0).Control(115).Enabled=   0   'False
         Tab(0).Control(116)=   "txtCp(19)"
         Tab(0).Control(116).Enabled=   0   'False
         Tab(0).Control(117)=   "txtCp(18)"
         Tab(0).Control(117).Enabled=   0   'False
         Tab(0).Control(118)=   "txtCp(17)"
         Tab(0).Control(118).Enabled=   0   'False
         Tab(0).Control(119)=   "bottxtCampo17(2)"
         Tab(0).Control(119).Enabled=   0   'False
         Tab(0).Control(120)=   "bottxtCampo17(1)"
         Tab(0).Control(120).Enabled=   0   'False
         Tab(0).Control(121)=   "txtCp(16)"
         Tab(0).Control(121).Enabled=   0   'False
         Tab(0).Control(122)=   "bottxtCampo16(0)"
         Tab(0).Control(122).Enabled=   0   'False
         Tab(0).Control(123)=   "txtCp(15)"
         Tab(0).Control(123).Enabled=   0   'False
         Tab(0).Control(124)=   "ChkCp(0)"
         Tab(0).Control(124).Enabled=   0   'False
         Tab(0).Control(125)=   "txtCp(14)"
         Tab(0).Control(125).Enabled=   0   'False
         Tab(0).Control(126)=   "txtCp(13)"
         Tab(0).Control(126).Enabled=   0   'False
         Tab(0).Control(127)=   "txtCp(12)"
         Tab(0).Control(127).Enabled=   0   'False
         Tab(0).Control(128)=   "txtCp(11)"
         Tab(0).Control(128).Enabled=   0   'False
         Tab(0).Control(129)=   "txtCp(10)"
         Tab(0).Control(129).Enabled=   0   'False
         Tab(0).Control(130)=   "txtCp(8)"
         Tab(0).Control(130).Enabled=   0   'False
         Tab(0).Control(131)=   "txtCp(7)"
         Tab(0).Control(131).Enabled=   0   'False
         Tab(0).Control(132)=   "txtCp(6)"
         Tab(0).Control(132).Enabled=   0   'False
         Tab(0).Control(133)=   "txtCp(5)"
         Tab(0).Control(133).Enabled=   0   'False
         Tab(0).Control(134)=   "txtCp(4)"
         Tab(0).Control(134).Enabled=   0   'False
         Tab(0).Control(135)=   "txtCp(3)"
         Tab(0).Control(135).Enabled=   0   'False
         Tab(0).Control(136)=   "txtCp(2)"
         Tab(0).Control(136).Enabled=   0   'False
         Tab(0).Control(137)=   "txtCp(1)"
         Tab(0).Control(137).Enabled=   0   'False
         Tab(0).Control(138)=   "Botao(1)"
         Tab(0).Control(138).Enabled=   0   'False
         Tab(0).Control(139)=   "Botao(3)"
         Tab(0).Control(139).Enabled=   0   'False
         Tab(0).Control(140)=   "Botao(4)"
         Tab(0).Control(140).Enabled=   0   'False
         Tab(0).Control(141)=   "Botao(5)"
         Tab(0).Control(141).Enabled=   0   'False
         Tab(0).ControlCount=   142
         TabCaption(1)   =   "&2 - Produtos"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label(50)"
         Tab(1).Control(1)=   "Grid(3)"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "&3 - Serviços"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label(49)"
         Tab(2).Control(1)=   "Grid(4)"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "&4 - Conjuntos"
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtCp(61)"
         Tab(3).Control(1)=   "ChkCp(3)"
         Tab(3).Control(2)=   "Label(47)"
         Tab(3).Control(3)=   "Label(39)"
         Tab(3).Control(4)=   "Botao(2)"
         Tab(3).Control(5)=   "txtCp(34)"
         Tab(3).Control(6)=   "bottxtCampo34(2)"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "bottxtCampo34(1)"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "labFdo34"
         Tab(3).Control(9)=   "Grid(0)"
         Tab(3).ControlCount=   10
         TabCaption(4)   =   "&5 - Peças"
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label(48)"
         Tab(4).Control(1)=   "Grid(1)"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "&6 - Financeiro"
         TabPicture(5)   =   "ORCAMENT.frx":001C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Botao(6)"
         Tab(5).Control(1)=   "Botao(7)"
         Tab(5).Control(2)=   "txtCp(9)"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).Control(3)=   "txtCp(31)"
         Tab(5).Control(3).Enabled=   0   'False
         Tab(5).Control(4)=   "bottxtCampo32(0)"
         Tab(5).Control(4).Enabled=   0   'False
         Tab(5).Control(5)=   "txtCp(32)"
         Tab(5).Control(6)=   "Grid(2)"
         Tab(5).Control(7)=   "bottxtCampo35(2)"
         Tab(5).Control(7).Enabled=   0   'False
         Tab(5).Control(8)=   "txtCp(35)"
         Tab(5).Control(9)=   "txtCp(36)"
         Tab(5).Control(9).Enabled=   0   'False
         Tab(5).Control(10)=   "txtCp(37)"
         Tab(5).Control(10).Enabled=   0   'False
         Tab(5).Control(11)=   "txtCp(38)"
         Tab(5).Control(11).Enabled=   0   'False
         Tab(5).Control(12)=   "txtCp(39)"
         Tab(5).Control(12).Enabled=   0   'False
         Tab(5).Control(13)=   "txtCp(40)"
         Tab(5).Control(13).Enabled=   0   'False
         Tab(5).Control(14)=   "txtCp(41)"
         Tab(5).Control(14).Enabled=   0   'False
         Tab(5).Control(15)=   "txtCp(42)"
         Tab(5).Control(15).Enabled=   0   'False
         Tab(5).Control(16)=   "txtCp(43)"
         Tab(5).Control(16).Enabled=   0   'False
         Tab(5).Control(17)=   "txtCp(44)"
         Tab(5).Control(17).Enabled=   0   'False
         Tab(5).Control(18)=   "bottxtCampo45(0)"
         Tab(5).Control(18).Enabled=   0   'False
         Tab(5).Control(19)=   "txtCp(45)"
         Tab(5).Control(20)=   "txtCp(46)"
         Tab(5).Control(20).Enabled=   0   'False
         Tab(5).Control(21)=   "bottxtCampo47(0)"
         Tab(5).Control(21).Enabled=   0   'False
         Tab(5).Control(22)=   "txtCp(47)"
         Tab(5).Control(23)=   "txtCp(49)"
         Tab(5).Control(23).Enabled=   0   'False
         Tab(5).Control(24)=   "Painel(2)"
         Tab(5).Control(25)=   "txtCp(52)"
         Tab(5).Control(25).Enabled=   0   'False
         Tab(5).Control(26)=   "txtCp(53)"
         Tab(5).Control(26).Enabled=   0   'False
         Tab(5).Control(27)=   "txtCp(56)"
         Tab(5).Control(27).Enabled=   0   'False
         Tab(5).Control(28)=   "bottxtCampo57(0)"
         Tab(5).Control(28).Enabled=   0   'False
         Tab(5).Control(29)=   "txtCp(57)"
         Tab(5).Control(30)=   "txtCp(148)"
         Tab(5).Control(31)=   "Label(5)"
         Tab(5).Control(32)=   "labFdo32"
         Tab(5).Control(33)=   "labFdo35"
         Tab(5).Control(34)=   "Label(7)"
         Tab(5).Control(35)=   "Label(8)"
         Tab(5).Control(36)=   "Formato(0)"
         Tab(5).Control(37)=   "Label(9)"
         Tab(5).Control(38)=   "Formato(1)"
         Tab(5).Control(39)=   "Label(10)"
         Tab(5).Control(40)=   "Formato(2)"
         Tab(5).Control(41)=   "Label(11)"
         Tab(5).Control(42)=   "Label(12)"
         Tab(5).Control(43)=   "Label(13)"
         Tab(5).Control(44)=   "Linha(0)"
         Tab(5).Control(45)=   "Label(17)"
         Tab(5).Control(46)=   "Label(18)"
         Tab(5).Control(47)=   "Label(19)"
         Tab(5).Control(48)=   "Label(20)"
         Tab(5).Control(49)=   "Label(21)"
         Tab(5).Control(50)=   "Label(22)"
         Tab(5).Control(51)=   "Label(23)"
         Tab(5).Control(52)=   "Label(24)"
         Tab(5).Control(53)=   "Label(25)"
         Tab(5).Control(54)=   "Label(26)"
         Tab(5).Control(55)=   "Label(27)"
         Tab(5).Control(56)=   "labFdo45"
         Tab(5).Control(57)=   "Label(28)"
         Tab(5).Control(58)=   "Label(29)"
         Tab(5).Control(59)=   "labFdo47"
         Tab(5).Control(60)=   "Label(30)"
         Tab(5).Control(61)=   "Label(38)"
         Tab(5).Control(62)=   "Label(54)"
         Tab(5).Control(63)=   "Label(55)"
         Tab(5).Control(64)=   "Label(58)"
         Tab(5).Control(65)=   "Label(59)"
         Tab(5).Control(66)=   "labFdo57"
         Tab(5).Control(67)=   "Label(173)"
         Tab(5).ControlCount=   68
         TabCaption(6)   =   "Ficha Técnica (Hidroturbo)"
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Linha(15)"
         Tab(6).Control(1)=   "Label(166)"
         Tab(6).Control(2)=   "Label(165)"
         Tab(6).Control(3)=   "Label(164)"
         Tab(6).Control(4)=   "Label(163)"
         Tab(6).Control(5)=   "Label(162)"
         Tab(6).Control(6)=   "Label(161)"
         Tab(6).Control(7)=   "Label(160)"
         Tab(6).Control(8)=   "Linha(14)"
         Tab(6).Control(9)=   "Label(159)"
         Tab(6).Control(10)=   "Label(158)"
         Tab(6).Control(11)=   "Label(157)"
         Tab(6).Control(12)=   "Label(156)"
         Tab(6).Control(13)=   "Label(155)"
         Tab(6).Control(14)=   "Label(154)"
         Tab(6).Control(15)=   "Label(153)"
         Tab(6).Control(16)=   "txtCp(114)"
         Tab(6).Control(16).Enabled=   0   'False
         Tab(6).Control(17)=   "txtCp(113)"
         Tab(6).Control(17).Enabled=   0   'False
         Tab(6).Control(18)=   "txtCp(112)"
         Tab(6).Control(18).Enabled=   0   'False
         Tab(6).Control(19)=   "txtCp(111)"
         Tab(6).Control(19).Enabled=   0   'False
         Tab(6).Control(20)=   "Label(113)"
         Tab(6).Control(21)=   "txtCp(110)"
         Tab(6).Control(21).Enabled=   0   'False
         Tab(6).Control(22)=   "txtCp(109)"
         Tab(6).Control(22).Enabled=   0   'False
         Tab(6).Control(23)=   "txtCp(108)"
         Tab(6).Control(23).Enabled=   0   'False
         Tab(6).Control(24)=   "Label(112)"
         Tab(6).Control(25)=   "Label(111)"
         Tab(6).Control(26)=   "txtCp(107)"
         Tab(6).Control(26).Enabled=   0   'False
         Tab(6).Control(27)=   "txtCp(106)"
         Tab(6).Control(27).Enabled=   0   'False
         Tab(6).Control(28)=   "txtCp(105)"
         Tab(6).Control(28).Enabled=   0   'False
         Tab(6).Control(29)=   "Label(110)"
         Tab(6).Control(30)=   "Label(109)"
         Tab(6).Control(31)=   "txtCp(104)"
         Tab(6).Control(31).Enabled=   0   'False
         Tab(6).Control(32)=   "txtCp(103)"
         Tab(6).Control(32).Enabled=   0   'False
         Tab(6).Control(33)=   "txtCp(102)"
         Tab(6).Control(33).Enabled=   0   'False
         Tab(6).Control(34)=   "txtCp(101)"
         Tab(6).Control(35)=   "bottxtCampo101(2)"
         Tab(6).Control(35).Enabled=   0   'False
         Tab(6).Control(36)=   "bottxtCampo101(1)"
         Tab(6).Control(36).Enabled=   0   'False
         Tab(6).Control(37)=   "labFdo101"
         Tab(6).Control(38)=   "txtCp(100)"
         Tab(6).Control(39)=   "bottxtCampo100(2)"
         Tab(6).Control(39).Enabled=   0   'False
         Tab(6).Control(40)=   "bottxtCampo100(1)"
         Tab(6).Control(40).Enabled=   0   'False
         Tab(6).Control(41)=   "labFdo100"
         Tab(6).Control(42)=   "Linha(9)"
         Tab(6).Control(43)=   "txtCp(99)"
         Tab(6).Control(44)=   "bottxtCampo99(2)"
         Tab(6).Control(44).Enabled=   0   'False
         Tab(6).Control(45)=   "bottxtCampo99(1)"
         Tab(6).Control(45).Enabled=   0   'False
         Tab(6).Control(46)=   "labFdo99"
         Tab(6).Control(47)=   "Label(108)"
         Tab(6).Control(48)=   "txtCp(98)"
         Tab(6).Control(49)=   "txtCp(97)"
         Tab(6).Control(50)=   "Label(107)"
         Tab(6).Control(51)=   "txtCp(96)"
         Tab(6).Control(52)=   "Label(106)"
         Tab(6).Control(53)=   "Label(105)"
         Tab(6).Control(54)=   "Linha(8)"
         Tab(6).Control(55)=   "Label(104)"
         Tab(6).Control(56)=   "Label(103)"
         Tab(6).Control(57)=   "Label(102)"
         Tab(6).Control(58)=   "Label(101)"
         Tab(6).Control(59)=   "Label(100)"
         Tab(6).Control(60)=   "Label(99)"
         Tab(6).Control(61)=   "txtCp(95)"
         Tab(6).Control(62)=   "Label(98)"
         Tab(6).Control(63)=   "Label(97)"
         Tab(6).Control(64)=   "Linha(7)"
         Tab(6).Control(65)=   "Linha(6)"
         Tab(6).Control(66)=   "Label(96)"
         Tab(6).Control(67)=   "txtCp(94)"
         Tab(6).Control(67).Enabled=   0   'False
         Tab(6).Control(68)=   "txtCp(93)"
         Tab(6).Control(68).Enabled=   0   'False
         Tab(6).Control(69)=   "txtCp(92)"
         Tab(6).Control(69).Enabled=   0   'False
         Tab(6).Control(70)=   "txtCp(91)"
         Tab(6).Control(70).Enabled=   0   'False
         Tab(6).Control(71)=   "txtCp(90)"
         Tab(6).Control(72)=   "Label(95)"
         Tab(6).Control(73)=   "Label(94)"
         Tab(6).Control(74)=   "Linha(5)"
         Tab(6).Control(75)=   "txtCp(89)"
         Tab(6).Control(76)=   "Label(93)"
         Tab(6).Control(77)=   "txtCp(88)"
         Tab(6).Control(78)=   "Label(92)"
         Tab(6).Control(79)=   "Label(91)"
         Tab(6).Control(80)=   "txtCp(87)"
         Tab(6).Control(80).Enabled=   0   'False
         Tab(6).Control(81)=   "txtCp(86)"
         Tab(6).Control(81).Enabled=   0   'False
         Tab(6).Control(82)=   "txtCp(85)"
         Tab(6).Control(82).Enabled=   0   'False
         Tab(6).Control(83)=   "Label(90)"
         Tab(6).Control(84)=   "txtCp(84)"
         Tab(6).Control(84).Enabled=   0   'False
         Tab(6).Control(85)=   "Label(89)"
         Tab(6).Control(86)=   "txtCp(83)"
         Tab(6).Control(86).Enabled=   0   'False
         Tab(6).Control(87)=   "Label(88)"
         Tab(6).Control(88)=   "txtCp(82)"
         Tab(6).Control(89)=   "Label(87)"
         Tab(6).Control(90)=   "Label(86)"
         Tab(6).Control(91)=   "txtCp(81)"
         Tab(6).Control(92)=   "txtCp(80)"
         Tab(6).Control(92).Enabled=   0   'False
         Tab(6).Control(93)=   "txtCp(79)"
         Tab(6).Control(93).Enabled=   0   'False
         Tab(6).Control(94)=   "Label(85)"
         Tab(6).Control(95)=   "txtCp(78)"
         Tab(6).Control(96)=   "txtCp(77)"
         Tab(6).Control(97)=   "Label(84)"
         Tab(6).Control(98)=   "Label(83)"
         Tab(6).Control(99)=   "txtCp(76)"
         Tab(6).Control(99).Enabled=   0   'False
         Tab(6).Control(100)=   "txtCp(75)"
         Tab(6).Control(101)=   "Label(82)"
         Tab(6).Control(102)=   "txtCp(74)"
         Tab(6).Control(103)=   "Label(81)"
         Tab(6).Control(104)=   "txtCp(73)"
         Tab(6).Control(105)=   "Label(80)"
         Tab(6).Control(106)=   "Label(79)"
         Tab(6).Control(107)=   "Label(78)"
         Tab(6).Control(108)=   "Linha(4)"
         Tab(6).Control(109)=   "Linha(3)"
         Tab(6).Control(110)=   "txtCp(72)"
         Tab(6).Control(111)=   "Label(77)"
         Tab(6).Control(112)=   "txtCp(71)"
         Tab(6).Control(112).Enabled=   0   'False
         Tab(6).Control(113)=   "Label(76)"
         Tab(6).Control(114)=   "txtCp(70)"
         Tab(6).Control(114).Enabled=   0   'False
         Tab(6).Control(115)=   "Label(75)"
         Tab(6).Control(116)=   "txtCp(69)"
         Tab(6).Control(117)=   "Label(74)"
         Tab(6).Control(118)=   "Label(73)"
         Tab(6).Control(119)=   "txtCp(68)"
         Tab(6).Control(120)=   "Label(72)"
         Tab(6).Control(121)=   "txtCp(67)"
         Tab(6).Control(122)=   "Label(71)"
         Tab(6).Control(123)=   "Linha(2)"
         Tab(6).Control(124)=   "Label(70)"
         Tab(6).Control(125)=   "txtCp(66)"
         Tab(6).Control(126)=   "Label(69)"
         Tab(6).ControlCount=   127
         TabCaption(7)   =   "Bomba / Motor"
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "txtCp(142)"
         Tab(7).Control(1)=   "txtCp(141)"
         Tab(7).Control(2)=   "txtCp(140)"
         Tab(7).Control(3)=   "txtCp(139)"
         Tab(7).Control(4)=   "txtCp(138)"
         Tab(7).Control(5)=   "Label(152)"
         Tab(7).Control(6)=   "Label(151)"
         Tab(7).Control(7)=   "Label(150)"
         Tab(7).Control(8)=   "Label(149)"
         Tab(7).Control(9)=   "Label(148)"
         Tab(7).Control(10)=   "Label(147)"
         Tab(7).Control(11)=   "txtCp(136)"
         Tab(7).Control(11).Enabled=   0   'False
         Tab(7).Control(12)=   "Label(145)"
         Tab(7).Control(13)=   "txtCp(135)"
         Tab(7).Control(13).Enabled=   0   'False
         Tab(7).Control(14)=   "Label(144)"
         Tab(7).Control(15)=   "txtCp(134)"
         Tab(7).Control(16)=   "Linha(13)"
         Tab(7).Control(17)=   "Label(143)"
         Tab(7).Control(18)=   "txtCp(133)"
         Tab(7).Control(18).Enabled=   0   'False
         Tab(7).Control(19)=   "Label(142)"
         Tab(7).Control(20)=   "Label(141)"
         Tab(7).Control(21)=   "Label(140)"
         Tab(7).Control(22)=   "txtCp(132)"
         Tab(7).Control(22).Enabled=   0   'False
         Tab(7).Control(23)=   "txtCp(131)"
         Tab(7).Control(24)=   "Label(139)"
         Tab(7).Control(25)=   "Label(138)"
         Tab(7).Control(26)=   "txtCp(130)"
         Tab(7).Control(27)=   "txtCp(129)"
         Tab(7).Control(28)=   "Label(137)"
         Tab(7).Control(29)=   "Label(136)"
         Tab(7).Control(30)=   "txtCp(128)"
         Tab(7).Control(31)=   "txtCp(127)"
         Tab(7).Control(32)=   "Label(135)"
         Tab(7).Control(33)=   "Label(134)"
         Tab(7).Control(34)=   "txtCp(126)"
         Tab(7).Control(35)=   "Label(133)"
         Tab(7).Control(36)=   "Label(132)"
         Tab(7).Control(37)=   "Linha(12)"
         Tab(7).Control(38)=   "Linha(11)"
         Tab(7).Control(39)=   "Label(131)"
         Tab(7).Control(40)=   "txtCp(125)"
         Tab(7).Control(40).Enabled=   0   'False
         Tab(7).Control(41)=   "Label(130)"
         Tab(7).Control(42)=   "Label(129)"
         Tab(7).Control(43)=   "Label(128)"
         Tab(7).Control(44)=   "Label(127)"
         Tab(7).Control(45)=   "Label(126)"
         Tab(7).Control(46)=   "txtCp(124)"
         Tab(7).Control(47)=   "Label(125)"
         Tab(7).Control(48)=   "txtCp(123)"
         Tab(7).Control(49)=   "Label(124)"
         Tab(7).Control(50)=   "Label(123)"
         Tab(7).Control(51)=   "txtCp(122)"
         Tab(7).Control(51).Enabled=   0   'False
         Tab(7).Control(52)=   "Label(122)"
         Tab(7).Control(53)=   "txtCp(121)"
         Tab(7).Control(53).Enabled=   0   'False
         Tab(7).Control(54)=   "txtCp(120)"
         Tab(7).Control(55)=   "Label(121)"
         Tab(7).Control(56)=   "txtCp(119)"
         Tab(7).Control(57)=   "Label(120)"
         Tab(7).Control(58)=   "Label(119)"
         Tab(7).Control(59)=   "txtCp(118)"
         Tab(7).Control(60)=   "txtCp(117)"
         Tab(7).Control(61)=   "Label(118)"
         Tab(7).Control(62)=   "Label(117)"
         Tab(7).Control(63)=   "txtCp(116)"
         Tab(7).Control(64)=   "Label(116)"
         Tab(7).Control(65)=   "txtCp(115)"
         Tab(7).Control(66)=   "Label(115)"
         Tab(7).Control(67)=   "Linha(10)"
         Tab(7).Control(68)=   "Label(114)"
         Tab(7).ControlCount=   69
         TabCaption(8)   =   "Mapa da Propriedade"
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "mmCampo(0)"
         Tab(8).Control(1)=   "Label(171)"
         Tab(8).ControlCount=   2
         Begin IRRIG.GBotao Botao 
            Height          =   315
            Index           =   7
            Left            =   -64005
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   495
            Width           =   2055
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Parcelamento"
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   375
            Index           =   6
            Left            =   -65505
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   4695
            Width           =   5070
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Atualizar Valor dos Produtos"
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   300
            Index           =   5
            Left            =   9120
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   4620
            Width           =   240
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   300
            Index           =   4
            Left            =   9600
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2700
            Width           =   240
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   300
            Index           =   3
            Left            =   10125
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1380
            Width           =   240
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   300
            Index           =   2
            Left            =   -71670
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   450
            Width           =   240
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionOffset   =   50
         End
         Begin IRRIG.GBotao Botao 
            Height          =   300
            Index           =   1
            Left            =   3300
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2700
            Width           =   240
            _ExtentX        =   0
            _ExtentY        =   0
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CaptionOffset   =   50
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   11610
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   915
            Width           =   960
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3390
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1380
            Width           =   2235
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   13230
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   915
            Width           =   1755
         End
         Begin VB.TextBox txtCp 
            DataField       =   "Observação"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Index           =   4
            Left            =   660
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Top             =   5220
            Width           =   14145
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   720
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   5520
            Width           =   9495
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   1845
            Width           =   2235
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1440
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "CEP"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   11610
            TabIndex        =   14
            Top             =   915
            Width           =   960
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total do ICMS"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   3615
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   1845
            Width           =   2025
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   7125
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   1845
            Width           =   2010
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Fax"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   12
            Left            =   870
            MaxLength       =   14
            TabIndex        =   16
            Top             =   1380
            Width           =   1440
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "CPF e CNPJ"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   13
            Left            =   870
            MaxLength       =   20
            TabIndex        =   23
            Top             =   1845
            Width           =   2235
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "RG e IE"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   14
            Left            =   3600
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1845
            Width           =   2025
         End
         Begin VB.CheckBox ChkCp 
            Caption         =   "Revenda"
            DataField       =   "Revenda"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   0
            Left            =   9225
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   1950
            Width           =   1140
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   15
            Left            =   11580
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   1860
            Width           =   3135
         End
         Begin VB.CommandButton bottxtCampo16 
            Height          =   270
            Index           =   0
            Left            =   14715
            Style           =   1  'Graphical
            TabIndex        =   207
            TabStop         =   0   'False
            Top             =   1875
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Email"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   16
            Left            =   11595
            MaxLength       =   255
            TabIndex        =   26
            Top             =   1875
            Width           =   3105
         End
         Begin VB.CommandButton bottxtCampo17 
            Height          =   270
            Index           =   1
            Left            =   9855
            Style           =   1  'Graphical
            TabIndex        =   210
            TabStop         =   0   'False
            Top             =   1395
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo17 
            Height          =   270
            Index           =   2
            Left            =   9600
            Style           =   1  'Graphical
            TabIndex        =   211
            TabStop         =   0   'False
            Top             =   1395
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Seqüência do Município"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   17
            Left            =   7140
            TabIndex        =   18
            Top             =   1395
            Width           =   2445
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Nome Cliente"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   18
            Left            =   915
            MaxLength       =   60
            TabIndex        =   7
            Top             =   480
            Width           =   4710
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   19
            Left            =   7125
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   1380
            Width           =   3000
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Número do Endereço"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   20
            Left            =   4950
            MaxLength       =   10
            TabIndex        =   12
            Top             =   915
            Width           =   675
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Caixa Postal"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   21
            Left            =   3390
            MaxLength       =   30
            TabIndex        =   17
            Top             =   1380
            Width           =   2235
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Telefone"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   22
            Left            =   13230
            MaxLength       =   14
            TabIndex        =   15
            Top             =   915
            Width           =   1320
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Bairro"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   23
            Left            =   7140
            MaxLength       =   50
            TabIndex        =   13
            Top             =   900
            Width           =   3240
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Endereço"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   24
            Left            =   11610
            MaxLength       =   100
            TabIndex        =   10
            Top             =   450
            Width           =   3375
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Complemento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   25
            Left            =   870
            MaxLength       =   100
            TabIndex        =   11
            Top             =   915
            Width           =   3615
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   26
            Left            =   11610
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   450
            Width           =   2940
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   27
            Left            =   11610
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1380
            Width           =   405
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   28
            Left            =   7125
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   915
            Width           =   3240
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   29
            Left            =   4950
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   915
            Width           =   675
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   30
            Left            =   870
            Locked          =   -1  'True
            MaxLength       =   254
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   915
            Width           =   3615
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total de Produtos Usados"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   31
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   4545
            Width           =   1605
         End
         Begin VB.CommandButton bottxtCampo32 
            Height          =   270
            Index           =   0
            Left            =   -70275
            Style           =   1  'Graphical
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   885
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Valor do Fechamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   32
            Left            =   -71595
            MultiLine       =   -1  'True
            TabIndex        =   48
            Top             =   885
            Width           =   1320
         End
         Begin VB.CommandButton bottxtCampo33 
            Height          =   270
            Index           =   1
            Left            =   3045
            Style           =   1  'Graphical
            TabIndex        =   217
            TabStop         =   0   'False
            Top             =   2715
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo33 
            Height          =   270
            Index           =   2
            Left            =   2790
            Style           =   1  'Graphical
            TabIndex        =   218
            TabStop         =   0   'False
            Top             =   2715
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Seqüência do Vendedor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   33
            Left            =   750
            TabIndex        =   27
            Top             =   2715
            Width           =   2025
         End
         Begin IRRIG.GListV Grid 
            Height          =   4365
            Index           =   0
            Left            =   -74880
            TabIndex        =   43
            Tag             =   "1"
            Top             =   930
            Width           =   14475
            _extentx        =   0
            _extenty        =   0
            editbackcolor   =   16777152
            edittextcolor   =   -2147483640
            celltextcolor   =   -2147483640
            stripes         =   -1  'True
            stripesbackcolor=   14737632
            rowheight       =   315
            savegridbars    =   -1  'True
            caption         =   "Conjuntos"
            font            =   "ORCAMENT.frx":0038
         End
         Begin IRRIG.GListV Grid 
            Height          =   4845
            Index           =   1
            Left            =   -74880
            TabIndex        =   44
            Tag             =   "1"
            Top             =   450
            Width           =   14475
            _extentx        =   0
            _extenty        =   0
            editbackcolor   =   16777152
            edittextcolor   =   -2147483640
            celltextcolor   =   -2147483640
            stripes         =   -1  'True
            stripesbackcolor=   14737632
            rowheight       =   315
            savegridbars    =   -1  'True
            caption         =   "Peças"
            font            =   "ORCAMENT.frx":005C
         End
         Begin VB.CommandButton bottxtCampo34 
            Height          =   270
            Index           =   1
            Left            =   -71940
            Style           =   1  'Graphical
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   465
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo34 
            Height          =   270
            Index           =   2
            Left            =   -72195
            Style           =   1  'Graphical
            TabIndex        =   221
            TabStop         =   0   'False
            Top             =   465
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Seqüência da Classificação"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   34
            Left            =   -74295
            TabIndex        =   41
            Top             =   465
            Width           =   2355
         End
         Begin IRRIG.GListV Grid 
            Height          =   3750
            Index           =   2
            Left            =   -65475
            TabIndex        =   53
            Tag             =   "1"
            Top             =   900
            Width           =   5070
            _extentx        =   0
            _extenty        =   0
            editbackcolor   =   16777152
            edittextcolor   =   -2147483640
            celltextcolor   =   -2147483640
            stripes         =   -1  'True
            stripesbackcolor=   14737632
            rowheight       =   315
            showgridcaption =   0   'False
            savegridbars    =   -1  'True
            caption         =   ""
            font            =   "ORCAMENT.frx":0080
         End
         Begin VB.CommandButton bottxtCampo35 
            Height          =   270
            Index           =   2
            Left            =   -70275
            Style           =   1  'Graphical
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Forma de Pagamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   35
            Left            =   -71595
            TabIndex        =   45
            Top             =   480
            Width           =   1320
         End
         Begin IRRIG.GListV Grid 
            Height          =   4845
            Index           =   3
            Left            =   -74820
            TabIndex        =   39
            Tag             =   "1"
            Top             =   540
            Width           =   14475
            _extentx        =   0
            _extenty        =   0
            editbackcolor   =   16777152
            edittextcolor   =   -2147483640
            celltextcolor   =   -2147483640
            stripes         =   -1  'True
            stripesbackcolor=   14737632
            rowheight       =   315
            savegridbars    =   -1  'True
            caption         =   "Produtos"
            font            =   "ORCAMENT.frx":00A4
         End
         Begin IRRIG.GListV Grid 
            Height          =   4845
            Index           =   4
            Left            =   -74880
            TabIndex        =   40
            Tag             =   "1"
            Top             =   450
            Width           =   14475
            _extentx        =   0
            _extenty        =   0
            editbackcolor   =   16777152
            edittextcolor   =   -2147483640
            celltextcolor   =   -2147483640
            stripes         =   -1  'True
            stripesbackcolor=   14737632
            rowheight       =   315
            savegridbars    =   -1  'True
            caption         =   "Serviços"
            font            =   "ORCAMENT.frx":00C8
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total do Orçamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   36
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   5010
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total IPI dos Produtos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   37
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2220
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total Conjuntos Usados"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   38
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   5010
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total IPI das Peças"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   39
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   3150
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total IPI dos Conjuntos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   40
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2685
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total dos Produtos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   41
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2220
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total dos Conjuntos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   42
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   2685
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total dos Serviços"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   43
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   3615
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total das Peças"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   44
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   3150
            Width           =   1605
         End
         Begin VB.CommandButton bottxtCampo45 
            Height          =   270
            Index           =   0
            Left            =   -66120
            Style           =   1  'Graphical
            TabIndex        =   246
            TabStop         =   0   'False
            Top             =   885
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Valor do Frete"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   45
            Left            =   -67440
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   885
            Width           =   1320
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total das Peças Usadas"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   46
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   5475
            Width           =   1605
         End
         Begin VB.CommandButton bottxtCampo47 
            Height          =   270
            Index           =   0
            Left            =   -70275
            Style           =   1  'Graphical
            TabIndex        =   250
            TabStop         =   0   'False
            Top             =   1290
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Valor do Seguro"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   47
            Left            =   -71595
            MultiLine       =   -1  'True
            TabIndex        =   50
            Top             =   1290
            Width           =   1320
         End
         Begin VB.CheckBox ChkCp 
            DataField       =   "É Propriedade"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   10170
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   570
            Width           =   200
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Nome da Propriedade"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   48
            Left            =   7140
            MaxLength       =   62
            TabIndex        =   8
            Top             =   480
            Width           =   2970
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total da Base de Cálculo"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   49
            Left            =   -71610
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Data do Fechamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   50
            Left            =   4260
            Locked          =   -1  'True
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1395
         End
         Begin IRRIG.GPainel Painel 
            Height          =   240
            Index           =   1
            Left            =   12630
            TabIndex        =   274
            TabStop         =   0   'False
            Top             =   1395
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   423
            BevelOuter      =   0
            BevelWidth      =   0
            BorderWidth     =   0
            Stretch         =   -1  'True
            Begin VB.OptionButton opcPainel1Cp 
               Caption         =   "Física"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Width           =   840
            End
            Begin VB.OptionButton opcPainel1Cp 
               Caption         =   "Jurídica"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   1
               Left            =   915
               TabIndex        =   22
               Top             =   0
               Width           =   1005
            End
            Begin VB.Label labopcPainel1 
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Tipo"
               Height          =   120
               Left            =   600
               TabIndex        =   275
               Top             =   180
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin IRRIG.GPainel Painel 
            Height          =   240
            Index           =   2
            Left            =   -67455
            TabIndex        =   276
            TabStop         =   0   'False
            Top             =   495
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   423
            BevelOuter      =   0
            BevelWidth      =   0
            BorderWidth     =   0
            Stretch         =   -1  'True
            Begin VB.OptionButton opcPainel2Cp 
               Caption         =   "Perc"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   0
               Left            =   0
               TabIndex        =   46
               Top             =   0
               Width           =   720
            End
            Begin VB.OptionButton opcPainel2Cp 
               Caption         =   "Valor"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   1
               Left            =   795
               TabIndex        =   47
               Top             =   0
               Width           =   780
            End
            Begin VB.Label labopcPainel2 
               BorderStyle     =   1  'Fixed Single
               DataField       =   "Fechamento"
               Height          =   120
               Left            =   480
               TabIndex        =   277
               Top             =   180
               Visible         =   0   'False
               Width           =   240
            End
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Código do Suframa"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   51
            Left            =   7125
            MaxLength       =   9
            TabIndex        =   25
            Top             =   1845
            Width           =   2010
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total da Base ST"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   52
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   4545
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            DataField       =   "Valor Total do ICMS ST"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   53
            Left            =   -67455
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1605
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Alíquota do ISS"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   54
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   3120
            Width           =   1005
         End
         Begin VB.CheckBox ChkCp 
            DataField       =   "Reter ISS"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1800
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   3120
            Width           =   200
         End
         Begin VB.CommandButton bottxtCampo55 
            Height          =   270
            Index           =   1
            Left            =   9315
            Style           =   1  'Graphical
            TabIndex        =   284
            TabStop         =   0   'False
            Top             =   2715
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo55 
            Height          =   270
            Index           =   2
            Left            =   9060
            Style           =   1  'Graphical
            TabIndex        =   285
            TabStop         =   0   'False
            Top             =   2715
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Seqüência da Transportadora"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   55
            Left            =   7395
            TabIndex        =   29
            Top             =   2715
            Width           =   1650
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   56
            Left            =   -67440
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   5475
            Width           =   1605
         End
         Begin VB.CommandButton bottxtCampo57 
            Height          =   270
            Index           =   0
            Left            =   -66120
            Style           =   1  'Graphical
            TabIndex        =   289
            TabStop         =   0   'False
            Top             =   1320
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Valor do Imposto de Renda"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   57
            Left            =   -67440
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   1320
            Width           =   1320
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Local de Embarque"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   58
            Left            =   720
            MaxLength       =   30
            TabIndex        =   31
            Top             =   4080
            Width           =   2085
         End
         Begin VB.CommandButton bottxtCampo59 
            Height          =   270
            Index           =   2
            Left            =   5355
            Style           =   1  'Graphical
            TabIndex        =   293
            TabStop         =   0   'False
            Top             =   4095
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "UF de Embarque"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   59
            Left            =   4275
            TabIndex        =   32
            Top             =   4095
            Width           =   1065
         End
         Begin VB.CommandButton bottxtCampo60 
            Height          =   270
            Index           =   1
            Left            =   8835
            Style           =   1  'Graphical
            TabIndex        =   298
            TabStop         =   0   'False
            Top             =   4635
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo60 
            Height          =   270
            Index           =   2
            Left            =   8580
            Style           =   1  'Graphical
            TabIndex        =   299
            TabStop         =   0   'False
            Top             =   4635
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Seqüência do País"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   60
            Left            =   6195
            TabIndex        =   36
            Top             =   4635
            Width           =   2370
         End
         Begin VB.CheckBox ChkCp 
            Caption         =   "Conjunto Avulso"
            DataField       =   "Conjunto Avulso"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   3
            Left            =   -71070
            TabIndex        =   95
            Top             =   495
            Width           =   1815
         End
         Begin VB.TextBox txtCp 
            DataField       =   "Descrição Conjunto Avulso"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   61
            Left            =   -69270
            MaxLength       =   60
            TabIndex        =   96
            Top             =   480
            Width           =   6810
         End
         Begin VB.CommandButton bottxtCampo62 
            Height          =   270
            Index           =   2
            Left            =   8355
            Style           =   1  'Graphical
            TabIndex        =   302
            TabStop         =   0   'False
            Top             =   4155
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Vendedor Intermediario"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   62
            Left            =   7350
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   4155
            Width           =   990
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Percentual do Vendedor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   63
            Left            =   9180
            MultiLine       =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   4140
            Width           =   765
         End
         Begin VB.CommandButton bottxtCampo64 
            Height          =   270
            Index           =   2
            Left            =   13395
            Style           =   1  'Graphical
            TabIndex        =   307
            TabStop         =   0   'False
            Top             =   4155
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Rebiut"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   64
            Left            =   11490
            TabIndex        =   100
            Top             =   4155
            Width           =   1890
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Percentual Rebiut"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   65
            Left            =   14115
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   4140
            Width           =   870
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Hidroturbo"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   66
            Left            =   -74700
            MaxLength       =   100
            TabIndex        =   106
            Top             =   660
            Width           =   7710
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Area irrigada"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   67
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   107
            Top             =   1380
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Precipitação bruta"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   68
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   108
            Top             =   1755
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Horas irrigada"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   69
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   109
            Top             =   2115
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   70
            Left            =   -67815
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   4275
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   71
            Left            =   -67815
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   4635
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Area tot irrigada em"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   72
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   110
            Top             =   2475
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Aspersor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   73
            Left            =   -68565
            MaxLength       =   20
            TabIndex        =   111
            Top             =   1395
            Width           =   1635
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Modelo do aspersor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   74
            Left            =   -68565
            MaxLength       =   40
            TabIndex        =   112
            Top             =   1755
            Width           =   1635
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Bocal diametro"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   75
            Left            =   -68565
            MultiLine       =   -1  'True
            TabIndex        =   113
            Top             =   2115
            Width           =   1275
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   76
            Left            =   -68565
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   2475
            Width           =   1275
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Pressão de serviço"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   77
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   115
            Top             =   3195
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Alcance do jato"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   78
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   116
            Top             =   3540
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   79
            Left            =   -71820
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   4635
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   80
            Left            =   -67815
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   4995
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Espaço entre carreadores"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   81
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   117
            Top             =   3915
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Faixa irrigada"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   82
            Left            =   -71820
            MultiLine       =   -1  'True
            TabIndex        =   118
            Top             =   4275
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   83
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   1455
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   84
            Left            =   -71820
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   4995
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   85
            Left            =   -71820
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   5355
            Width           =   750
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   86
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   435
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   87
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   2895
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Desnivel maximo na area"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   88
            Left            =   -67815
            MultiLine       =   -1  'True
            TabIndex        =   122
            Top             =   3195
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Altura de sucção"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   89
            Left            =   -67815
            MultiLine       =   -1  'True
            TabIndex        =   123
            Top             =   3540
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Altura do aspersor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   90
            Left            =   -67815
            MultiLine       =   -1  'True
            TabIndex        =   124
            Top             =   3915
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   91
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   1815
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   92
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   2175
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   93
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   2535
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   94
            Left            =   -61305
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   3255
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Tempo parado antes percurso"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   95
            Left            =   -61305
            MultiLine       =   -1  'True
            TabIndex        =   125
            Top             =   1095
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Com 1"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   96
            Left            =   -66660
            MultiLine       =   -1  'True
            TabIndex        =   136
            Top             =   4380
            Width           =   645
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Com 2"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   97
            Left            =   -66660
            MultiLine       =   -1  'True
            TabIndex        =   137
            Top             =   4740
            Width           =   645
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Com 3"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   98
            Left            =   -66660
            MultiLine       =   -1  'True
            TabIndex        =   138
            Top             =   5100
            Width           =   645
         End
         Begin VB.CommandButton bottxtCampo99 
            Height          =   270
            Index           =   1
            Left            =   -64125
            Style           =   1  'Graphical
            TabIndex        =   349
            TabStop         =   0   'False
            Top             =   4395
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo99 
            Height          =   270
            Index           =   2
            Left            =   -64380
            Style           =   1  'Graphical
            TabIndex        =   350
            TabStop         =   0   'False
            Top             =   4395
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            DataField       =   "Modelo Trecho A"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   99
            Left            =   -65865
            TabIndex        =   139
            Top             =   4395
            Width           =   1740
         End
         Begin VB.CommandButton bottxtCampo100 
            Height          =   270
            Index           =   1
            Left            =   -64125
            Style           =   1  'Graphical
            TabIndex        =   352
            TabStop         =   0   'False
            Top             =   4755
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo100 
            Height          =   270
            Index           =   2
            Left            =   -64380
            Style           =   1  'Graphical
            TabIndex        =   353
            TabStop         =   0   'False
            Top             =   4755
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            DataField       =   "Modelo Trecho B"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   100
            Left            =   -65865
            TabIndex        =   140
            Top             =   4755
            Width           =   1740
         End
         Begin VB.CommandButton bottxtCampo101 
            Height          =   270
            Index           =   1
            Left            =   -64125
            Style           =   1  'Graphical
            TabIndex        =   355
            TabStop         =   0   'False
            Top             =   5115
            Width           =   255
         End
         Begin VB.CommandButton bottxtCampo101 
            Height          =   270
            Index           =   2
            Left            =   -64380
            Style           =   1  'Graphical
            TabIndex        =   356
            TabStop         =   0   'False
            Top             =   5115
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            DataField       =   "Modelo Trecho C"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   101
            Left            =   -65865
            TabIndex        =   141
            Top             =   5115
            Width           =   1740
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   102
            Left            =   -63765
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   4380
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   103
            Left            =   -63765
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   4740
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   104
            Left            =   -63765
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   5100
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   105
            Left            =   -62925
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   4380
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   106
            Left            =   -62925
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   4740
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   107
            Left            =   -62925
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   5100
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   108
            Left            =   -62085
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   148
            TabStop         =   0   'False
            Top             =   4380
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   109
            Left            =   -62085
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   4740
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   110
            Left            =   -62085
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   5100
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   111
            Left            =   -61200
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   4380
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   112
            Left            =   -61200
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   4740
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   113
            Left            =   -61200
            Locked          =   -1  'True
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   5100
            Width           =   780
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   114
            Left            =   -67815
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   5355
            Width           =   705
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Qtde bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   115
            Left            =   -73380
            MultiLine       =   -1  'True
            TabIndex        =   155
            Top             =   1140
            Width           =   870
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Marca bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   116
            Left            =   -73380
            MaxLength       =   25
            TabIndex        =   156
            Top             =   1500
            Width           =   2625
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Modelo bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   117
            Left            =   -73380
            MaxLength       =   25
            TabIndex        =   157
            Top             =   1875
            Width           =   2625
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Tamanho bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   118
            Left            =   -73380
            MaxLength       =   20
            TabIndex        =   158
            Top             =   2235
            Width           =   2625
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "N estagios"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   119
            Left            =   -73380
            MultiLine       =   -1  'True
            TabIndex        =   159
            Top             =   2595
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Diametro bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   120
            Left            =   -73380
            MultiLine       =   -1  'True
            TabIndex        =   160
            Top             =   2955
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   121
            Left            =   -73380
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   161
            TabStop         =   0   'False
            Top             =   3315
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   122
            Left            =   -73380
            Locked          =   -1  'True
            MaxLength       =   8
            MultiLine       =   -1  'True
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   3675
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Rendimento bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   123
            Left            =   -73380
            MultiLine       =   -1  'True
            TabIndex        =   163
            Top             =   4035
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Rotação bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   124
            Left            =   -73380
            MultiLine       =   -1  'True
            TabIndex        =   164
            Top             =   4395
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   125
            Left            =   -73380
            Locked          =   -1  'True
            MaxLength       =   8
            MultiLine       =   -1  'True
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   4755
            Width           =   825
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Qtde de Motor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   126
            Left            =   -67380
            MultiLine       =   -1  'True
            TabIndex        =   166
            Top             =   1140
            Width           =   870
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Marca do Motor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   127
            Left            =   -67380
            MaxLength       =   20
            TabIndex        =   167
            Top             =   1500
            Width           =   2460
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Modelo Motor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   128
            Left            =   -67380
            MaxLength       =   25
            TabIndex        =   168
            Top             =   1875
            Width           =   2460
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Nivel de Proteção"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   129
            Left            =   -67380
            MaxLength       =   15
            TabIndex        =   169
            Top             =   2235
            Width           =   2460
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Potencia Nominal"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   130
            Left            =   -67380
            MultiLine       =   -1  'True
            TabIndex        =   170
            Top             =   2595
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Nro de Fases"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   131
            Left            =   -67380
            MultiLine       =   -1  'True
            TabIndex        =   171
            Top             =   2955
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   132
            Left            =   -67380
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   172
            TabStop         =   0   'False
            Top             =   3315
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   133
            Left            =   -67380
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   3675
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Voltagem"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   134
            Left            =   -67380
            MultiLine       =   -1  'True
            TabIndex        =   174
            Top             =   4035
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   135
            Left            =   -67380
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   4395
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   136
            Left            =   -67380
            Locked          =   -1  'True
            MaxLength       =   7
            MultiLine       =   -1  'True
            TabIndex        =   176
            TabStop         =   0   'False
            Top             =   4755
            Width           =   900
         End
         Begin VB.CommandButton bottxtCampo137 
            Height          =   270
            Index           =   2
            Left            =   14715
            Style           =   1  'Graphical
            TabIndex        =   396
            TabStop         =   0   'False
            Top             =   3195
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Projetista"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   137
            Left            =   12675
            TabIndex        =   99
            Top             =   3195
            Width           =   2025
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Modelo hidroturbo"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   138
            Left            =   -63420
            MaxLength       =   30
            TabIndex        =   177
            Top             =   1155
            Width           =   2820
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Eixos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   139
            Left            =   -63420
            MultiLine       =   -1  'True
            TabIndex        =   178
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Rodas"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   140
            Left            =   -63420
            MultiLine       =   -1  'True
            TabIndex        =   179
            Top             =   1935
            Width           =   900
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Pneus"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   141
            Left            =   -63420
            MaxLength       =   20
            TabIndex        =   180
            Top             =   2340
            Width           =   2820
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            DataField       =   "Tubos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   142
            Left            =   -63420
            MaxLength       =   25
            TabIndex        =   181
            Top             =   2760
            Width           =   2820
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Peso Bruto"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   143
            Left            =   735
            MultiLine       =   -1  'True
            TabIndex        =   183
            Top             =   3600
            Width           =   2130
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Peso Líquido"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   144
            Left            =   4260
            MultiLine       =   -1  'True
            TabIndex        =   184
            Top             =   3600
            Width           =   1365
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Volumes"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   145
            Left            =   7380
            MultiLine       =   -1  'True
            TabIndex        =   185
            Top             =   3600
            Width           =   1245
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "Aviso de embarque"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   146
            Left            =   10140
            MaxLength       =   100
            TabIndex        =   186
            Top             =   3660
            Width           =   4845
         End
         Begin IRRIG.GCpMM mmCampo 
            Height          =   4950
            Index           =   0
            Left            =   -74925
            TabIndex        =   422
            TabStop         =   0   'False
            Top             =   1005
            Width           =   14430
            _extentx        =   0
            _extenty        =   0
            captionfont     =   "ORCAMENT.frx":00EC
            locked          =   -1  'True
            scrollbars      =   0   'False
         End
         Begin VB.CommandButton bottxtCampo147 
            Height          =   270
            Index           =   2
            Left            =   4995
            Style           =   1  'Graphical
            TabIndex        =   425
            TabStop         =   0   'False
            Top             =   4635
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Entrega Tecnica"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   147
            Left            =   1995
            TabIndex        =   188
            TabStop         =   0   'False
            Top             =   4635
            Width           =   2985
         End
         Begin VB.TextBox txtCp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "Outras Despesas"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   148
            Left            =   -71610
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   1680
            Width           =   1605
         End
         Begin VB.CommandButton bottxtCampo149 
            Height          =   270
            Index           =   0
            Left            =   8415
            Style           =   1  'Graphical
            TabIndex        =   429
            TabStop         =   0   'False
            Top             =   3195
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            DataField       =   "Data do Pedido"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   149
            Left            =   7395
            TabIndex        =   192
            TabStop         =   0   'False
            Top             =   3195
            Width           =   1005
         End
         Begin VB.CommandButton bottxtCampo150 
            Height          =   270
            Index           =   1
            Left            =   11265
            Style           =   1  'Graphical
            TabIndex        =   432
            TabStop         =   0   'False
            Top             =   3195
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            DataField       =   "Data de Entrega"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   150
            Left            =   10155
            TabIndex        =   193
            TabStop         =   0   'False
            Top             =   3195
            Width           =   1095
         End
         Begin VB.CommandButton bottxtCampo151 
            Height          =   270
            Index           =   2
            Left            =   6015
            Style           =   1  'Graphical
            TabIndex        =   435
            TabStop         =   0   'False
            Top             =   2715
            Width           =   255
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            DataField       =   "Frete"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   151
            Left            =   4275
            TabIndex        =   197
            TabStop         =   0   'False
            Top             =   2715
            Width           =   1725
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "PlacaVeiculo"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   152
            Left            =   10560
            MaxLength       =   254
            TabIndex        =   198
            Top             =   2700
            Width           =   1185
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "UfPlaca"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   153
            Left            =   12240
            MaxLength       =   254
            TabIndex        =   200
            Top             =   2700
            Width           =   645
         End
         Begin VB.TextBox txtCp 
            Appearance      =   0  'Flat
            DataField       =   "NumAntt"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   154
            Left            =   13695
            MaxLength       =   254
            TabIndex        =   199
            Top             =   2700
            Width           =   1290
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "RG:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   3180
            TabIndex        =   202
            Top             =   1875
            Width           =   345
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   203
            Top             =   1875
            Width           =   570
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Município:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   6150
            TabIndex        =   204
            Top             =   1395
            Width           =   900
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   11085
            TabIndex        =   205
            Top             =   945
            Width           =   450
         End
         Begin VB.Label labFdo16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   11580
            TabIndex        =   206
            Top             =   1860
            Width           =   3405
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   10980
            TabIndex        =   208
            Top             =   1875
            Width           =   555
         End
         Begin VB.Label labFdo17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7125
            TabIndex        =   209
            Top             =   1380
            Width           =   3000
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fechamento:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   5
            Left            =   -68700
            TabIndex        =   212
            Top             =   495
            Width           =   1170
         End
         Begin VB.Label labFdo32 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -71610
            TabIndex        =   213
            Top             =   870
            Width           =   1605
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "*Vend.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   60
            TabIndex        =   215
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label labFdo33 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   735
            TabIndex        =   216
            Top             =   2700
            Width           =   2580
         End
         Begin VB.Label labFdo34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -74310
            TabIndex        =   219
            Top             =   450
            Width           =   2640
         End
         Begin VB.Label labFdo35 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -71610
            TabIndex        =   222
            Top             =   465
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Parcelamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   7
            Left            =   -65475
            TabIndex        =   224
            Top             =   465
            Width           =   5070
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   -65055
            TabIndex        =   225
            Top             =   5730
            Width           =   600
         End
         Begin VB.Shape Formato 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            FillColor       =   &H000000FF&
            Height          =   165
            Index           =   0
            Left            =   -64380
            Shape           =   1  'Square
            Top             =   5775
            Width           =   165
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Incompleto"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   -64155
            TabIndex        =   226
            Top             =   5730
            Width           =   990
         End
         Begin VB.Shape Formato 
            BackColor       =   &H0080FFFF&
            BackStyle       =   1  'Opaque
            FillColor       =   &H000000FF&
            Height          =   165
            Index           =   1
            Left            =   -63000
            Shape           =   1  'Square
            Top             =   5775
            Width           =   165
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "S/ Parcelas"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   10
            Left            =   -62760
            TabIndex        =   227
            Top             =   5730
            Width           =   1050
         End
         Begin VB.Shape Formato 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            FillColor       =   &H000000FF&
            Height          =   150
            Index           =   2
            Left            =   -61545
            Shape           =   1  'Square
            Top             =   5775
            Width           =   150
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Ok."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   11
            Left            =   -61320
            TabIndex        =   228
            Top             =   5730
            Width           =   300
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pagamento:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   12
            Left            =   -73695
            TabIndex        =   229
            Top             =   495
            Width           =   2010
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Desconto.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   13
            Left            =   -72645
            TabIndex        =   230
            Top             =   900
            Width           =   960
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "*Município:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   14
            Left            =   6075
            TabIndex        =   231
            Top             =   1395
            Width           =   975
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   15
            Left            =   12090
            TabIndex        =   232
            Top             =   1395
            Width           =   465
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   16
            Left            =   225
            TabIndex        =   233
            Top             =   1875
            Width           =   570
         End
         Begin VB.Line Linha 
            BorderColor     =   &H00C0C0C0&
            Index           =   0
            X1              =   -74295
            X2              =   -65850
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do IPI Produtos:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   17
            Left            =   -73620
            TabIndex        =   234
            Top             =   2250
            Width           =   1935
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do IPI Conjuntos:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   -73695
            TabIndex        =   235
            Top             =   2715
            Width           =   2010
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do IPI Peças:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   19
            Left            =   -73395
            TabIndex        =   236
            Top             =   3180
            Width           =   1710
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   20
            Left            =   -73020
            TabIndex        =   237
            Top             =   3645
            Width           =   1335
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor dos Produtos Usados:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   21
            Left            =   -74220
            TabIndex        =   238
            Top             =   4575
            Width           =   2535
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor dos Conjuntos Usados:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   22
            Left            =   -74295
            TabIndex        =   239
            Top             =   5040
            Width           =   2610
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor dos Produtos:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   23
            Left            =   -69300
            TabIndex        =   240
            Top             =   2250
            Width           =   1770
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor dos Conjuntos:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   24
            Left            =   -69375
            TabIndex        =   241
            Top             =   2715
            Width           =   1845
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor dos Serviços:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   25
            Left            =   -69285
            TabIndex        =   242
            Top             =   3645
            Width           =   1755
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Orçamento:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   26
            Left            =   -69375
            TabIndex        =   243
            Top             =   5040
            Width           =   1845
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do Frete:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   27
            Left            =   -68850
            TabIndex        =   244
            Top             =   900
            Width           =   1320
         End
         Begin VB.Label labFdo45 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -67455
            TabIndex        =   245
            Top             =   870
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor das Peças Usadas:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   28
            Left            =   -73995
            TabIndex        =   247
            Top             =   5505
            Width           =   2310
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor das Peças:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   29
            Left            =   -69075
            TabIndex        =   248
            Top             =   3180
            Width           =   1545
         End
         Begin VB.Label labFdo47 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -71610
            TabIndex        =   249
            Top             =   1275
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do Seguro:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   30
            Left            =   -73200
            TabIndex        =   251
            Top             =   1305
            Width           =   1515
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Propriedade:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   31
            Left            =   5865
            TabIndex        =   252
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Compl.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   32
            Left            =   120
            TabIndex        =   253
            Top             =   945
            Width           =   675
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nº.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   33
            Left            =   4560
            TabIndex        =   254
            Top             =   945
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   34
            Left            =   11220
            TabIndex        =   255
            Top             =   1395
            Width           =   315
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cx. Postal:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   35
            Left            =   2385
            TabIndex        =   256
            Top             =   1410
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   36
            Left            =   6465
            TabIndex        =   257
            Top             =   945
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   37
            Left            =   11085
            TabIndex        =   258
            Top             =   945
            Width           =   450
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo ICMS:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   38
            Left            =   -73740
            TabIndex        =   259
            Top             =   4110
            Width           =   2055
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "NCM:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   39
            Left            =   -74880
            TabIndex        =   260
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   40
            Left            =   120
            TabIndex        =   261
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   41
            Left            =   420
            TabIndex        =   262
            Top             =   1395
            Width           =   375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fone:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   42
            Left            =   12645
            TabIndex        =   263
            Top             =   945
            Width           =   510
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   43
            Left            =   10605
            TabIndex        =   264
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Informações do Orçamento"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   44
            Left            =   120
            TabIndex        =   265
            Top             =   2250
            Width           =   2430
         End
         Begin VB.Line Linha 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   120
            X2              =   14550
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data do Fechamento:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   45
            Left            =   2100
            TabIndex        =   266
            Top             =   3120
            Width           =   1935
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hist.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   46
            Left            =   60
            TabIndex        =   267
            Top             =   5160
            Width           =   435
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "F2 - Recalcula Valores | F12 - Abre Cadastro de Conjuntos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   47
            Left            =   -74880
            TabIndex        =   268
            Top             =   5370
            Width           =   14475
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "F2 - Recalcula Valores | F12 - Abre Cadastro de Produtos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   48
            Left            =   -74880
            TabIndex        =   269
            Top             =   5370
            Width           =   14475
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "F12 - Abre Cadastro de Serviços"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   49
            Left            =   -74880
            TabIndex        =   270
            Top             =   5370
            Width           =   14475
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "F2 - Recalcula Valores | F12 - Abre Cadastro de Produtos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   50
            Left            =   -74880
            TabIndex        =   271
            Top             =   5430
            Width           =   14475
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   51
            Left            =   10605
            TabIndex        =   272
            Top             =   480
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   52
            Left            =   6420
            TabIndex        =   273
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Suframa:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   53
            Left            =   6255
            TabIndex        =   278
            Top             =   1875
            Width           =   795
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo ST:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   54
            Left            =   -69375
            TabIndex        =   279
            Top             =   4575
            Width           =   1845
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS ST:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   55
            Left            =   -69180
            TabIndex        =   280
            Top             =   4110
            Width           =   1650
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "% ISS:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   56
            Left            =   60
            TabIndex        =   281
            Top             =   3120
            Width           =   585
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Transport:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   57
            Left            =   6420
            TabIndex        =   282
            Top             =   2700
            Width           =   915
         End
         Begin VB.Label labFdo55 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7380
            TabIndex        =   283
            Top             =   2700
            Width           =   2205
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Valor do ISS:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   58
            Left            =   -68700
            TabIndex        =   286
            Top             =   5505
            Width           =   1170
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor do I.R.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   59
            Left            =   -68670
            TabIndex        =   287
            Top             =   1305
            Width           =   1140
         End
         Begin VB.Label labFdo57 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -67455
            TabIndex        =   288
            Top             =   1305
            Width           =   1605
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emb.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   60
            Left            =   120
            TabIndex        =   290
            Top             =   4080
            Width           =   510
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "UF.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   61
            Left            =   3660
            TabIndex        =   291
            Top             =   4140
            Width           =   360
         End
         Begin VB.Label labFdo59 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   292
            Top             =   4080
            Width           =   1365
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fone:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   62
            Left            =   12645
            TabIndex        =   294
            Top             =   945
            Width           =   510
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   63
            Left            =   420
            TabIndex        =   295
            Top             =   1395
            Width           =   375
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "*País:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   64
            Left            =   5580
            TabIndex        =   296
            Top             =   4680
            Width           =   525
         End
         Begin VB.Label labFdo60 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6180
            TabIndex        =   297
            Top             =   4620
            Width           =   2925
         End
         Begin VB.Label Label 
            Caption         =   "Revenda:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   65
            Left            =   6300
            TabIndex        =   300
            Top             =   4140
            Width           =   960
         End
         Begin VB.Label labFdo62 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7335
            TabIndex        =   301
            Top             =   4140
            Width           =   1290
         End
         Begin VB.Label Label 
            Caption         =   "%:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   66
            Left            =   8835
            TabIndex        =   303
            Top             =   4140
            Width           =   270
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vend Rebiul:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   67
            Left            =   10155
            TabIndex        =   304
            Top             =   4140
            Width           =   1155
         End
         Begin VB.Label Label 
            Caption         =   "%:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   68
            Left            =   13800
            TabIndex        =   305
            Top             =   4140
            Width           =   285
         End
         Begin VB.Label labFdo64 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   11475
            TabIndex        =   306
            Top             =   4140
            Width           =   2190
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descrição do Hidroturbo"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   69
            Left            =   -74700
            TabIndex        =   308
            Top             =   360
            Width           =   2550
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Dados Técnicos Gerais"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   70
            Left            =   -74685
            TabIndex        =   309
            Top             =   1035
            Width           =   2460
         End
         Begin VB.Line Linha 
            Index           =   2
            X1              =   -74685
            X2              =   -66930
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Área aproximada a ser irrigada:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   71
            Left            =   -74700
            TabIndex        =   310
            Top             =   1380
            Width           =   2835
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Precipitação Bruta.........................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   72
            Left            =   -74670
            TabIndex        =   311
            Top             =   1755
            Width           =   2820
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Horas Irrigada por dia..................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   73
            Left            =   -74685
            TabIndex        =   312
            Top             =   2115
            Width           =   2820
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Perda de Carga (Adutora 2):"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   74
            Left            =   -70545
            TabIndex        =   313
            Top             =   4635
            Width           =   2685
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Perda de Carga no Tubo.....:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   75
            Left            =   -70545
            TabIndex        =   314
            Top             =   4995
            Width           =   2700
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Área Total Irrigada em.................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   76
            Left            =   -74685
            TabIndex        =   315
            Top             =   2475
            Width           =   2820
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vazão Necessária...............:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   77
            Left            =   -74685
            TabIndex        =   316
            Top             =   2835
            Width           =   2880
         End
         Begin VB.Line Linha 
            Index           =   3
            X1              =   -74685
            X2              =   -66870
            Y1              =   3090
            Y2              =   3090
         End
         Begin VB.Line Linha 
            Index           =   4
            X1              =   -74685
            X2              =   -66870
            Y1              =   2805
            Y2              =   2805
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo do Aspersor....:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   78
            Left            =   -70485
            TabIndex        =   317
            Top             =   1395
            Width           =   1800
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Marca/Modelo..........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   79
            Left            =   -70485
            TabIndex        =   318
            Top             =   1755
            Width           =   1815
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vazão:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   80
            Left            =   -71700
            TabIndex        =   319
            Top             =   2835
            Width           =   720
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Diametro do Bocal..:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   81
            Left            =   -70485
            TabIndex        =   320
            Top             =   2115
            Width           =   1815
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vazão Total........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   82
            Left            =   -70485
            TabIndex        =   321
            Top             =   2475
            Width           =   1800
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pressão de Serviço......................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   83
            Left            =   -74685
            TabIndex        =   322
            Top             =   3195
            Width           =   2820
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Alcance do Jato.............................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   84
            Left            =   -74685
            TabIndex        =   323
            Top             =   3555
            Width           =   2805
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Espaçamento Carreadores.......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   85
            Left            =   -74685
            TabIndex        =   324
            Top             =   3915
            Width           =   2805
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Comprimento Faixa Irrigada......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   86
            Left            =   -74685
            TabIndex        =   325
            Top             =   4275
            Width           =   2790
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Perdas no Hidroturbo..........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   87
            Left            =   -70545
            TabIndex        =   326
            Top             =   5355
            Width           =   2715
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Velocidade Deslocamento.......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   88
            Left            =   -74685
            TabIndex        =   327
            Top             =   4635
            Width           =   2775
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Precipitação Liquida....................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   89
            Left            =   -74685
            TabIndex        =   328
            Top             =   4995
            Width           =   2790
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vazão Total por Turno.................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   90
            Left            =   -74685
            TabIndex        =   329
            Top             =   5355
            Width           =   2820
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alturas Manométricas"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   91
            Left            =   -70485
            TabIndex        =   330
            Top             =   2820
            Width           =   2235
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Desnível Maximo na Área.......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   92
            Left            =   -70545
            TabIndex        =   331
            Top             =   3195
            Width           =   2655
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Altura de Sucção........................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   93
            Left            =   -70545
            TabIndex        =   332
            Top             =   3555
            Width           =   2655
         End
         Begin VB.Line Linha 
            Index           =   5
            X1              =   -66795
            X2              =   -66795
            Y1              =   390
            Y2              =   5715
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Altura do Aspersor.....................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   94
            Left            =   -70545
            TabIndex        =   333
            Top             =   3915
            Width           =   2655
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Perda de Carga (Adutora 1):"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   95
            Left            =   -70545
            TabIndex        =   334
            Top             =   4275
            Width           =   2685
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Altura Manometrica Total.........................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   96
            Left            =   -66585
            TabIndex        =   335
            Top             =   435
            Width           =   5115
         End
         Begin VB.Line Linha 
            Index           =   6
            X1              =   -70440
            X2              =   -68205
            Y1              =   3030
            Y2              =   3030
         End
         Begin VB.Line Linha 
            Index           =   7
            X1              =   -66630
            X2              =   -60405
            Y1              =   825
            Y2              =   825
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funcionamento do Sistema"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   97
            Left            =   -66585
            TabIndex        =   336
            Top             =   855
            Width           =   2805
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Parado Antes de Iniciar o Percurso............................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   98
            Left            =   -66585
            TabIndex        =   337
            Top             =   1155
            Width           =   5115
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Velocidade de Deslocamento do Canhão..............................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   99
            Left            =   -66585
            TabIndex        =   338
            Top             =   1455
            Width           =   5130
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tempo para Irrigar 01 Faixa..........................................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   100
            Left            =   -66585
            TabIndex        =   339
            Top             =   1815
            Width           =   5130
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Área Irrigada por Faixa..................................................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   101
            Left            =   -66585
            TabIndex        =   340
            Top             =   2175
            Width           =   5070
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Faixas Irrigada / dia............................................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   102
            Left            =   -66585
            TabIndex        =   341
            Top             =   2535
            Width           =   5070
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Área Irrigada por dia.......................................................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   103
            Left            =   -66585
            TabIndex        =   342
            Top             =   2895
            Width           =   5085
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Área Irrigada por Turno..................................................................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   104
            Left            =   -66585
            TabIndex        =   343
            Top             =   3255
            Width           =   5100
         End
         Begin VB.Line Linha 
            Index           =   8
            X1              =   -66705
            X2              =   -60315
            Y1              =   3585
            Y2              =   3585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Composição da Adutora e Tubos"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   105
            Left            =   -66645
            TabIndex        =   344
            Top             =   3615
            Width           =   3420
         End
         Begin VB.Label Label 
            Caption         =   "Trecho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   106
            Left            =   -66660
            TabIndex        =   345
            Top             =   3900
            Width           =   645
         End
         Begin VB.Label Label 
            Caption         =   "Metros"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   107
            Left            =   -66645
            TabIndex        =   346
            Top             =   4095
            Width           =   645
         End
         Begin VB.Label Label 
            Caption         =   "Adutora"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   108
            Left            =   -65880
            TabIndex        =   347
            Top             =   4095
            Width           =   765
         End
         Begin VB.Label labFdo99 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -65880
            TabIndex        =   348
            Top             =   4380
            Width           =   2025
         End
         Begin VB.Line Linha 
            Index           =   9
            X1              =   -66705
            X2              =   -60315
            Y1              =   3885
            Y2              =   3885
         End
         Begin VB.Label labFdo100 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -65880
            TabIndex        =   351
            Top             =   4740
            Width           =   2025
         End
         Begin VB.Label labFdo101 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   -65880
            TabIndex        =   354
            Top             =   5100
            Width           =   2025
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Diam."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   109
            Left            =   -63630
            TabIndex        =   357
            Top             =   3900
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "(mm)"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   110
            Left            =   -63585
            TabIndex        =   358
            Top             =   4095
            Width           =   450
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Coef."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   111
            Left            =   -62805
            TabIndex        =   359
            Top             =   4095
            Width           =   465
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "HF"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   112
            Left            =   -61845
            TabIndex        =   360
            Top             =   4095
            Width           =   270
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Veloc."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   113
            Left            =   -61125
            TabIndex        =   361
            Top             =   4095
            Width           =   570
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ponto de Funcionamento da Bomba"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   114
            Left            =   -74535
            TabIndex        =   362
            Top             =   495
            Width           =   5040
         End
         Begin VB.Line Linha 
            Index           =   10
            X1              =   -74520
            X2              =   -60540
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Qtde............:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   115
            Left            =   -74535
            TabIndex        =   363
            Top             =   1140
            Width           =   1020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Marca.........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   116
            Left            =   -74535
            TabIndex        =   364
            Top             =   1500
            Width           =   1020
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Modelo.......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   117
            Left            =   -74535
            TabIndex        =   365
            Top             =   1875
            Width           =   1050
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tamanho...:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   118
            Left            =   -74535
            TabIndex        =   366
            Top             =   2235
            Width           =   1050
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nº Rotores:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   119
            Left            =   -74535
            TabIndex        =   367
            Top             =   2595
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Diametro...:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   120
            Left            =   -74535
            TabIndex        =   368
            Top             =   2955
            Width           =   1005
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Vazão........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   121
            Left            =   -74535
            TabIndex        =   369
            Top             =   3315
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pressão.....:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   122
            Left            =   -74535
            TabIndex        =   370
            Top             =   3675
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Rendimento.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   123
            Left            =   -74685
            TabIndex        =   371
            Top             =   4035
            Width           =   1185
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Rotação..:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   124
            Left            =   -74415
            TabIndex        =   372
            Top             =   4395
            Width           =   915
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   125
            Left            =   -72465
            TabIndex        =   373
            Top             =   2955
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m³/h"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   126
            Left            =   -72465
            TabIndex        =   374
            Top             =   3315
            Width           =   390
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mca"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   127
            Left            =   -72465
            TabIndex        =   375
            Top             =   3675
            Width           =   390
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   128
            Left            =   -72465
            TabIndex        =   376
            Top             =   4035
            Width           =   180
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "rpm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   129
            Left            =   -72465
            TabIndex        =   377
            Top             =   4395
            Width           =   345
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Potência.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   130
            Left            =   -74400
            TabIndex        =   378
            Top             =   4755
            Width           =   885
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "cv"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   131
            Left            =   -72465
            TabIndex        =   379
            Top             =   4755
            Width           =   210
         End
         Begin VB.Line Linha 
            Index           =   11
            X1              =   -74520
            X2              =   -60480
            Y1              =   5250
            Y2              =   5250
         End
         Begin VB.Line Linha 
            Index           =   12
            X1              =   -69368
            X2              =   -69368
            Y1              =   900
            Y2              =   5250
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dados do Motor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   132
            Left            =   -69225
            TabIndex        =   380
            Top             =   495
            Width           =   2220
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Qtde............................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   133
            Left            =   -69225
            TabIndex        =   381
            Top             =   1140
            Width           =   1740
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Marca.........................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   134
            Left            =   -69225
            TabIndex        =   382
            Top             =   1500
            Width           =   1740
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Modelo......................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   135
            Left            =   -69225
            TabIndex        =   383
            Top             =   1875
            Width           =   1725
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nivel de Proteção.:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   136
            Left            =   -69225
            TabIndex        =   384
            Top             =   2235
            Width           =   1710
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Potencia Nominal..:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   137
            Left            =   -69225
            TabIndex        =   385
            Top             =   2595
            Width           =   1725
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Fases............:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   138
            Left            =   -69225
            TabIndex        =   386
            Top             =   2955
            Width           =   1710
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Rotação....................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   139
            Left            =   -69225
            TabIndex        =   387
            Top             =   3315
            Width           =   1725
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "rpm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   140
            Left            =   -66345
            TabIndex        =   388
            Top             =   3315
            Width           =   345
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "cv"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   141
            Left            =   -66345
            TabIndex        =   389
            Top             =   2595
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Demanda.................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   142
            Left            =   -69225
            TabIndex        =   390
            Top             =   3675
            Width           =   1710
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Voltagem..................:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   143
            Left            =   -69225
            TabIndex        =   391
            Top             =   4035
            Width           =   1725
         End
         Begin VB.Line Linha 
            Index           =   13
            X1              =   -64748
            X2              =   -64748
            Y1              =   915
            Y2              =   5265
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Amperagem............:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   144
            Left            =   -69225
            TabIndex        =   392
            Top             =   4395
            Width           =   1710
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Consumo Estimado:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   145
            Left            =   -69285
            TabIndex        =   393
            Top             =   4755
            Width           =   1800
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Projetista:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   146
            Left            =   11640
            TabIndex        =   394
            Top             =   3180
            Width           =   885
         End
         Begin VB.Label labFdo137 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   12660
            TabIndex        =   395
            Top             =   3180
            Width           =   2325
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caracteristicas do Carretel"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   147
            Left            =   -64665
            TabIndex        =   397
            Top             =   495
            Width           =   3660
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Modelo.......:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   148
            Left            =   -64605
            TabIndex        =   398
            Top             =   1155
            Width           =   1050
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Eixos............:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   149
            Left            =   -64605
            TabIndex        =   399
            Top             =   1560
            Width           =   1080
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Rodas.........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   150
            Left            =   -64605
            TabIndex        =   400
            Top             =   1935
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Pneus..........:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   151
            Left            =   -64605
            TabIndex        =   401
            Top             =   2340
            Width           =   1065
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tubos PE...:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   152
            Left            =   -64620
            TabIndex        =   402
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "ha"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   153
            Left            =   -71025
            TabIndex        =   403
            Top             =   1395
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m³/h"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   154
            Left            =   -72525
            TabIndex        =   404
            Top             =   2835
            Width           =   375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   155
            Left            =   -71025
            TabIndex        =   405
            Top             =   1755
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m³/h"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   156
            Left            =   -67245
            TabIndex        =   406
            Top             =   2475
            Width           =   375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Hrs"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   157
            Left            =   -71025
            TabIndex        =   407
            Top             =   2175
            Width           =   285
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Dias"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   158
            Left            =   -71025
            TabIndex        =   408
            Top             =   2535
            Width           =   375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   159
            Left            =   -67245
            TabIndex        =   409
            Top             =   2115
            Width           =   330
         End
         Begin VB.Line Linha 
            Index           =   14
            X1              =   -70635
            X2              =   -70635
            Y1              =   1305
            Y2              =   2775
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mca"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   160
            Left            =   -71025
            TabIndex        =   410
            Top             =   3135
            Width           =   360
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   161
            Left            =   -71025
            TabIndex        =   411
            Top             =   3555
            Width           =   165
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   162
            Left            =   -71025
            TabIndex        =   412
            Top             =   3915
            Width           =   165
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   163
            Left            =   -71025
            TabIndex        =   413
            Top             =   4275
            Width           =   165
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m/h"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   164
            Left            =   -71025
            TabIndex        =   414
            Top             =   4635
            Width           =   315
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   165
            Left            =   -71025
            TabIndex        =   415
            Top             =   4995
            Width           =   330
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "m³."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   166
            Left            =   -71025
            TabIndex        =   416
            Top             =   5355
            Width           =   270
         End
         Begin VB.Line Linha 
            Index           =   15
            X1              =   -70620
            X2              =   -70620
            Y1              =   3135
            Y2              =   5775
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "P.Bruto:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   167
            Left            =   0
            TabIndex        =   417
            Top             =   3600
            Width           =   690
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "P.Líquido:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   168
            Left            =   3120
            TabIndex        =   418
            Top             =   3600
            Width           =   885
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Volume:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   169
            Left            =   6480
            TabIndex        =   419
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Embarque:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   170
            Left            =   9000
            TabIndex        =   420
            Top             =   3660
            Width           =   990
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   $"ORCAMENT.frx":011A
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   171
            Left            =   -74850
            TabIndex        =   421
            Top             =   615
            Width           =   14280
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Entrega Tecnica:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   172
            Left            =   120
            TabIndex        =   423
            Top             =   4620
            Width           =   1770
         End
         Begin VB.Label labFdo147 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1980
            TabIndex        =   424
            Top             =   4620
            Width           =   3285
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Outras Despesas / Juros:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   173
            Left            =   -73980
            TabIndex        =   426
            Top             =   1695
            Width           =   2265
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data do Pedido:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   174
            Left            =   5820
            TabIndex        =   427
            Top             =   3180
            Width           =   1470
         End
         Begin VB.Label labFdo149 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7380
            TabIndex        =   428
            Top             =   3180
            Width           =   1305
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Prev. Entrega:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   175
            Left            =   8820
            TabIndex        =   430
            Top             =   3180
            Width           =   1260
         End
         Begin VB.Label labFdo150 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   10140
            TabIndex        =   431
            Top             =   3180
            Width           =   1395
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Frete:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   176
            Left            =   3660
            TabIndex        =   433
            Top             =   2760
            Width           =   510
         End
         Begin VB.Label labFdo151 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   434
            Top             =   2700
            Width           =   2025
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   177
            Left            =   9900
            TabIndex        =   436
            Top             =   2700
            Width           =   570
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   178
            Left            =   11880
            TabIndex        =   437
            Top             =   2700
            Width           =   315
         End
         Begin VB.Label Label 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "ANTT:"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   179
            Left            =   13020
            TabIndex        =   438
            Top             =   2700
            Width           =   600
         End
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   155
         Left            =   13110
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   750
         Width           =   1440
      End
      Begin VB.PictureBox picBox 
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Index           =   0
         Left            =   75
         ScaleHeight     =   1410
         ScaleWidth      =   1425
         TabIndex        =   441
         Top             =   135
         Width           =   1425
      End
      Begin VB.CommandButton bottxtCampo156 
         Height          =   270
         Index           =   1
         Left            =   9540
         Style           =   1  'Graphical
         TabIndex        =   443
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.CommandButton bottxtCampo156 
         Height          =   270
         Index           =   2
         Left            =   9285
         Style           =   1  'Graphical
         TabIndex        =   444
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Seqüência do Geral"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   156
         Left            =   5070
         TabIndex        =   3
         Top             =   1665
         Width           =   4200
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Ocultar Vr. Unitário"
         DataField       =   "Ocultar Valor Unitário"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   4
         Left            =   1665
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1950
      End
      Begin VB.CommandButton bottxtCampo157 
         Height          =   270
         Index           =   1
         Left            =   14310
         Style           =   1  'Graphical
         TabIndex        =   447
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.CommandButton bottxtCampo157 
         Height          =   270
         Index           =   2
         Left            =   14055
         Style           =   1  'Graphical
         TabIndex        =   448
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Seqüência da Propriedade"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   157
         Left            =   11415
         TabIndex        =   5
         Top             =   1665
         Width           =   2625
      End
      Begin VB.CommandButton bottxtCampo158 
         Height          =   270
         Index           =   0
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   450
         TabStop         =   0   'False
         Top             =   1665
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         DataField       =   "Data de Emissão"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   158
         Left            =   2820
         TabIndex        =   2
         Top             =   1665
         Width           =   1185
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Entrega Futura"
         DataField       =   "Entrega Futura"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   5
         Left            =   5520
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1605
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Venda Fechada"
         DataField       =   "Venda Fechada"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   6
         Left            =   3690
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1755
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Avulso"
         DataField       =   "Orçamento Avulso"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   7
         Left            =   7200
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1260
         Width           =   915
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Fatura Proforma"
         DataField       =   "Fatura Proforma"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   8
         Left            =   60
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   9000
         Width           =   1740
      End
      Begin VB.TextBox txtCp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataField       =   "Número da Proforma"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   159
         Left            =   615
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1650
         Width           =   840
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Não Baixar Estoque (Banco)"
         DataField       =   "Nao Movimentar Estoque"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   120
         Width           =   3255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   160
         Left            =   11835
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   735
         Width           =   975
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Refaturamento"
         DataField       =   "Refaturamento"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   10
         Left            =   1680
         TabIndex        =   191
         Top             =   360
         Width           =   2145
      End
      Begin VB.CheckBox ChkCp 
         Caption         =   "Ordem Interna"
         DataField       =   "Ordem Interna"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   194
         Top             =   9240
         Width           =   1440
      End
      Begin IRRIG.GCpMM mmCampo 
         Height          =   1380
         Index           =   1
         Left            =   180
         TabIndex        =   456
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
         _extentx        =   0
         _extenty        =   0
         captionfont     =   "ORCAMENT.frx":01A3
         appearance      =   0
         borderstyle     =   0
         locked          =   -1  'True
         scrollbars      =   0   'False
      End
      Begin VB.CommandButton bottxtCampo161 
         Height          =   270
         Index           =   2
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   459
         TabStop         =   0   'False
         Top             =   135
         Width           =   255
      End
      Begin VB.TextBox txtCp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         DataField       =   "Orçamento Vinculado"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   161
         Left            =   3675
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   135
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "-"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   180
         Left            =   14880
         TabIndex        =   439
         Top             =   8265
         Width           =   45
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Orçamento"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   181
         Left            =   1665
         TabIndex        =   440
         Top             =   540
         Width           =   2145
      End
      Begin VB.Label labFdo156 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5055
         TabIndex        =   442
         Top             =   1650
         Width           =   4755
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Propriedade:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   182
         Left            =   10140
         TabIndex        =   445
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label labFdo157 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   11400
         TabIndex        =   446
         Top             =   1650
         Width           =   3180
      End
      Begin VB.Label labFdo158 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2805
         TabIndex        =   449
         Top             =   1650
         Width           =   1485
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "*Dt. Emissão:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   183
         Left            =   1530
         TabIndex        =   451
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Seq.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   184
         Left            =   120
         TabIndex        =   452
         Top             =   1680
         Width           =   465
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "*Geral:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   185
         Left            =   4365
         TabIndex        =   453
         Top             =   1680
         Width           =   615
      End
      Begin VB.Line Linha 
         Index           =   16
         X1              =   1665
         X2              =   14820
         Y1              =   1095
         Y2              =   1095
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Inadimplente"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   186
         Left            =   8895
         TabIndex        =   454
         Top             =   1335
         Width           =   1155
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Projeto:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   187
         Left            =   11835
         TabIndex        =   455
         Top             =   495
         Width           =   690
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pedido Vinculado"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   188
         Left            =   1680
         TabIndex        =   457
         Top             =   135
         Width           =   1845
      End
      Begin VB.Label labFdo161 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   3660
         TabIndex        =   458
         Top             =   120
         Width           =   1110
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "vinculo"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   189
         Left            =   4995
         TabIndex        =   460
         Top             =   60
         Width           =   1005
      End
   End
   Begin IRRIG.GListV grdBrowse 
      Height          =   795
      Left            =   60
      TabIndex        =   461
      Top             =   330
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1402
      fullrowselect   =   0   'False
      stripesbackcolor=   14737632
      rowheight       =   285
      allowedit       =   -1  'True
      allowinsert     =   -1  'True
      allowdelete     =   -1  'True
      manualupdate    =   -1  'True
      manualdelete    =   -1  'True
      navigationaddmode=   1
      showfilterbar   =   -1  'True
      showgridcaption =   0   'False
      caption         =   ""
      font            =   "ORCAMENT.frx":01D1
      cachesize       =   100
   End
   Begin VB.Timer timUnLoad 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   30
      Top             =   950
   End
End
Attribute VB_Name = "frmOrcament"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    modOrcamentoLogic.opcPainel1Cp_Click Me, Index
    modOrcamentoLogic.opcPainel2Cp_Click Me, Index
    modOrcamentoLogic.RepositionOrcamento Me
    TotalParcelas = modOrcamentoLogic.TotalParcelas(Me, Seq)
    modOrcamentoLogic.MostraFormulas Me
    InfoProdutos = modOrcamentoLogic.InfoProdutos(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal, Oq)
   Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double, _
   Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, _
   Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double, _
   Valor_do_PIS As Double, Valor_do_Cofins As Double, IVA As Double, _
   Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double, CFOP As Integer, _
   CST As Integer, Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, _
   Valor_do_Desconto As Double, Valor_Anterior As Double, Bc_cofins As Double, _
   Aliq_do_pis As Single, Aliq_do_cofins As Single, Peso As Double, PesoTotal As Double, Oq As String) As Variant
    InfoLista = modOrcamentoLogic.InfoLista(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal, Oq)
    InfoServicos = modOrcamentoLogic.InfoServicos(Me, SequenciadoOrcamento, SequenciadoServicoOrcamento, SequenciadoServico, Quantidade, ValorUnitario, ValorTotal, ValorAnterior, Oq)
    InfoConjuntos = modOrcamentoLogic.InfoConjuntos(Me, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Oq)
    InfoPecas = modOrcamentoLogic.InfoPecas(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso, Oq)
    ProcessaProdutos = modOrcamentoLogic.ProcessaProdutos(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)
    ProcessaConjuntos = modOrcamentoLogic.ProcessaConjuntos(Me, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins)
    ProcessaServicos = modOrcamentoLogic.ProcessaServicos(Me, SequenciadoOrcamento, SequenciadoServicoOrcamento, SequenciadoServico, Quantidade, ValorUnitario, ValorTotal, ValorAnterior)
    ProcessaPecas = modOrcamentoLogic.ProcessaPecas(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso)
    modOrcamentoLogic.CarregaTotalizador Me
    modOrcamentoLogic.ComandosProdutos Me, KeyAscii, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal
    modOrcamentoLogic.ComandosServicos Me, KeyAscii, SequenciadoOrcamento, SequenciadoServicoOrcamento, SequenciadoServico, Quantidade, ValorUnitario, ValorTotal, ValorAnterior
    modOrcamentoLogic.ComandosConjuntos Me, KeyAscii, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins
    modOrcamentoLogic.ComandosPecas Me, KeyAscii, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso
    modOrcamentoLogic.AjustaValores Me
    modOrcamentoLogic.GeraPedido Me
    UltimaParcela = modOrcamentoLogic.UltimaParcela(Me)
    UltimoDias = modOrcamentoLogic.UltimoDias(Me)
    UltimoVencimento = modOrcamentoLogic.UltimoVencimento(Me)
    modOrcamentoLogic.AjustaSubstituicao Me
    modOrcamentoLogic.Tabs_Click Me, Index, PreviousTab
    PegaNCMPadrao = modOrcamentoLogic.PegaNCMPadrao(Me)
    modOrcamentoLogic.GeraConta Me
    VerificaEstoque = modOrcamentoLogic.VerificaEstoque(Me)
    VerificaDebitos = modOrcamentoLogic.VerificaDebitos(Me)
    modOrcamentoLogic.Alteracao Me
    modOrcamentoLogic.Form_Deactivate Me
    modOrcamentoLogic.SuperAtualizaProdutos Me
    modOrcamentoLogic.SuperAtualizaPecas Me
    modOrcamentoLogic.SuperAtualizaConjuntos Me
    modOrcamentoLogic.VerificaFormulario Me
    modOrcamentoLogic.Posiciona Me
    modOrcamentoLogic.AbreRel Me
    RetornaNF = modOrcamentoLogic.RetornaNF(Me)
    modOrcamentoLogic.AbreGerar Me
    modOrcamentoLogic.LimitaCampo Me, KeyAscii, MaxLength
    modOrcamentoLogic.AbreNotaFiscal Me
End Sub

    modOrcamentoLogic.AbreProjeto Me

    ValidaProduto2 = modOrcamentoLogic.ValidaProduto2(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)
    ValidaProduto3 = modOrcamentoLogic.ValidaProduto3(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso)
    FiltroOrc = modOrcamentoLogic.FiltroOrc(Me)
    Permissao = modOrcamentoLogic.Permissao(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)
    PermissaoPecas = modOrcamentoLogic.PermissaoPecas(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso)
    PermissaoConj = modOrcamentoLogic.PermissaoConj(Me, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins)
    ValidaPecasx = modOrcamentoLogic.ValidaPecasx(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso)
    ValidaProdx = modOrcamentoLogic.ValidaProdx(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)
    ValidaConjx = modOrcamentoLogic.ValidaConjx(Me, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins)
    modOrcamentoLogic.EncargosFinanceiros Me
    modOrcamentoLogic.AtualizaValor Me
    modOrcamentoLogic.ValorOriginal Me
    modOrcamentoLogic.Parcelar Me
    modOrcamentoLogic.DistribuiDescontoTotal Me
    ContaRegs = modOrcamentoLogic.ContaRegs(Me, tbl)
    modOrcamentoLogic.LimpaProp Me
    modOrcamentoLogic.BuscaVendedor Me
    modOrcamentoLogic.PreValidaImpressao Me
    VazaoAux = modOrcamentoLogic.VazaoAux(Me)
    Coef1 = modOrcamentoLogic.Coef1(Me)
    Coef2 = modOrcamentoLogic.Coef2(Me)
    Coef3 = modOrcamentoLogic.Coef3(Me)
    Diam1 = modOrcamentoLogic.Diam1(Me)

    Diam2 = modOrcamentoLogic.Diam2(Me)
Private Function Diam3() As Double
    Diam3 = modOrcamentoLogic.Diam3(Me)
    HF1 = modOrcamentoLogic.HF1(Me)
    HF2 = modOrcamentoLogic.HF2(Me)
    HF3 = modOrcamentoLogic.HF3(Me)
    Velo1 = modOrcamentoLogic.Velo1(Me)
    Velo2 = modOrcamentoLogic.Velo2(Me)
    Velo3 = modOrcamentoLogic.Velo3(Me)
    TempoFx1 = modOrcamentoLogic.TempoFx1(Me)
    VelocidadeDesloca = modOrcamentoLogic.VelocidadeDesloca(Me)
    AreaporFx = modOrcamentoLogic.AreaporFx(Me)
    FaixasIrrigadas = modOrcamentoLogic.FaixasIrrigadas(Me)
    areapordia = modOrcamentoLogic.areapordia(Me)
    ConsumoEstimado = modOrcamentoLogic.ConsumoEstimado(Me)
    modOrcamentoLogic.AjustaValoresProforma Me
    PodeVenderProd = modOrcamentoLogic.PodeVenderProd(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)

    PodeVenderPecas = modOrcamentoLogic.PodeVenderPecas(Me, ValordoTributo, SequenciadoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, SequenciaPecasdoOrcamento, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins, Peso)
    ValidaNCM = modOrcamentoLogic.ValidaNCM(Me, Bcpis, ValordoTributo, SequenciadoOrcamento, SequenciadoProdutoOrcamento, SequenciadoProduto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, ValorICMSST, CFOP, CST, AliquotadoICMSST, ValordoFrete, ValordoDesconto, ValorAnterior, Bccofins, Aliqdopis, Aliqdocofins, Peso, PesoTotal)
    ValidaConjunto = modOrcamentoLogic.ValidaConjunto(Me, SequenciadoOrcamento, SequenciaConjuntoOrcamento, SequenciadoConjunto, Quantidade, ValorUnitario, ValorTotal, ValordoIPI, ValordoICMS, AliquotadoIPI, AliquotadoICMS, PercentualdaReducao, Diferido, ValordaBasedeCalculo, ValordoTributo, ValordoPIS, ValordoCofins, IVA, BasedeCalculoST, CFOP, CST, ValorICMSST, AliquotadoICMSST, ValordoDesconto, ValordoFrete, ValorAnterior, Bcpis, Aliqdopis, Bccofins, Aliqdocofins)
    modOrcamentoLogic.CarregaFotos Me
    modOrcamentoLogic.AtualizaValoresConjuntos Me
    modOrcamentoLogic.AtualizaValoresPecas Me
    modOrcamentoLogic.AtualizaValoresProdutos Me
    ValidaDesconto = modOrcamentoLogic.ValidaDesconto(Me)
    MostraVinculo = modOrcamentoLogic.MostraVinculo(Me)
    modOrcamentoLogic.VerificaVinculo Me
    modOrcamentoLogic.PreValidaVinculo Me
    modOrcamentoLogic.PreValidaNFE Me
    modOrcamentoLogic.DesmarcaFw Me
    modOrcamentoLogic.MYSQLInclui Me
    Filtro1 = modOrcamentoLogic.Filtro1(Me)
    Filtro2 = modOrcamentoLogic.Filtro2(Me)
    Filtro3 = modOrcamentoLogic.Filtro3(Me)
    modOrcamentoLogic.LigaDesligaBotoes Me
    txtCampoTab = modOrcamentoLogic.txtCampoTab(Me, Index)
End Property
    modOrcamentoLogic.InicializaApelidos Me, vgComOQue
    modOrcamentoLogic.AnalisaCondicoes Me
    Executar = modOrcamentoLogic.Executar(Me, vgOq, ByRef)
    modOrcamentoLogic.grdBrowse_DeleteData Me, vgItem, vgColumns, vgDataDeleted, vgErrorMessage

    modOrcamentoLogic.grdBrowse_InitEdit Me, CancelEdit
    modOrcamentoLogic.grdBrowse_ItemSelect Me, vgItem, vgColumns
    modOrcamentoLogic.grdBrowse_SkipRecord Me, Columns, BookMark
    modOrcamentoLogic.grdBrowse_GetColumnFilter Me, vgColumn, vgColumns, vgFilter
    modOrcamentoLogic.grdBrowse_GetColumnLocked Me, vgRow, vgCol, vgColumns, FormField, vgLocked

    modOrcamentoLogic.grdBrowse_SaveData Me, vgItem, vgColumns, vgDataSaved, vgColumn, vgErrorMessage

    modOrcamentoLogic.grdBrowse_StatusChanged Me, vgNewStatus
    modOrcamentoLogic.grdBrowse_MouseUp Me, Button, Shift, x, y, vgCurCol
    modOrcamentoLogic.LigaDesligaAlterar Me
    modOrcamentoLogic.Form_KeyPress Me, KeyAscii
    modOrcamentoLogic.Form_Paint Me
    modOrcamentoLogic.CancelGrids Me
    modOrcamentoLogic.SaveGrids Me
    modOrcamentoLogic.Reposition Me, ForceRebind, LockGrids
    modOrcamentoLogic.ExecutaPreValidacao Me
    modOrcamentoLogic.ExecutaVisivel Me
    modOrcamentoLogic.txtCp_GotFocus Me, Index
    modOrcamentoLogic.txtCp_KeyDown Me, Index, KeyCode, Shift
End Sub

    modOrcamentoLogic.txtCp_KeyPress Me, Index, KeyAscii
    modOrcamentoLogic.txtCp_LostFocus Me, Index
    modOrcamentoLogic.chkCp_Click Me, Index
    modOrcamentoLogic.chkCp_GotFocus Me, Index
    modOrcamentoLogic.chkCp_KeyDown Me, Index, KeyCode, Shift
    modOrcamentoLogic.chkCp_KeyPress Me, Index, KeyAscii
    modOrcamentoLogic.chkCp_LostFocus Me, Index
    modOrcamentoLogic.opcPainel1Cp_KeyPress Me, Index, KeyAscii
    modOrcamentoLogic.opcPainel1Cp_KeyDown Me, Index, KeyCode, Shift
    modOrcamentoLogic.opcPainel1Cp_GotFocus Me, Index
    modOrcamentoLogic.opcPainel1Cp_LostFocus Me, Index
    modOrcamentoLogic.opcPainel2Cp_KeyPress Me, Index, KeyAscii
    modOrcamentoLogic.opcPainel2Cp_KeyDown Me, Index, KeyCode, Shift
    modOrcamentoLogic.opcPainel2Cp_GotFocus Me, Index
    modOrcamentoLogic.opcPainel2Cp_LostFocus Me, Index
    modOrcamentoLogic.Form_Activate Me
    modOrcamentoLogic.Form_Load Me
    CodigoFrete = modOrcamentoLogic.CodigoFrete(Me)
    modOrcamentoLogic.LimpaTransportadora Me
    modOrcamentoLogic.txtCp_Change Me, Index
    modOrcamentoLogic.AtualizaCamposFrete Me
    modOrcamentoLogic.DefineControles Me
Public Sub DefineControles1()
    modOrcamentoLogic.DefineControles1 Me
Public Sub DefineControles2()
    modOrcamentoLogic.DefineControles2 Me

    modOrcamentoLogic.DefineControles3 Me
    modOrcamentoLogic.txtCp_BeforeUpdate Me, Index, Cancel
    modOrcamentoLogic.PoeRelEFiltroCbo Me, Index
    modOrcamentoLogic.Form_QueryUnload Me, Cancel, UnloadMode
    modOrcamentoLogic.Form_Resize Me
    modOrcamentoLogic.Form_Unload Me, Cancel
    modOrcamentoLogic.Grid_KeyPress Me, Index, KeyAscii, Shift, vgColumns
    modOrcamentoLogic.Grid_SkipRecord Me, Index, vgColumns, vgBookMark
    modOrcamentoLogic.Grid_AfterUpdateRecord Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage

    modOrcamentoLogic.Grid_BeforeEditRecord Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage
    modOrcamentoLogic.Grid_BeforeUpdateRecord Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage
    modOrcamentoLogic.Grid_BeforeDeleteRecord Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage
    modOrcamentoLogic.Grid_GetColumnValue Me, Index, vgItem, vgCol, vgColumns, vgNewText

    modOrcamentoLogic.Grid_RecordSetChanged Me, Index, vgNewRecordSet
    modOrcamentoLogic.Grid_ValidateData Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage
    modOrcamentoLogic.Grid_GetColor Me, Index, vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
    modOrcamentoLogic.Grid_AfterDeleteRecord Me, Index, vgItem, vgColumns, vgIsValid, vgColumn, vgErrorMessage
    modOrcamentoLogic.Grid_ControlButtonClick Me, Index

    modOrcamentoLogic.Grid_GotFocus Me, Index
    modOrcamentoLogic.Grid_StatusChanged Me, Index, vgNewStatus
    modOrcamentoLogic.Grid_GetColumnFilter Me, Index, vgColumn, vgColumns, vgFilter
    modOrcamentoLogic.Grid_GetColumnSQLSearch Me, Index, vgColumn, vgColumns, vgSQLSearch
    ExecutaGrid = modOrcamentoLogic.ExecutaGrid(Me, Index, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    ExecutaGrid0 = modOrcamentoLogic.ExecutaGrid0(Me, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    ExecutaGrid1 = modOrcamentoLogic.ExecutaGrid1(Me, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    ExecutaGrid2 = modOrcamentoLogic.ExecutaGrid2(Me, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    ExecutaGrid3 = modOrcamentoLogic.ExecutaGrid3(Me, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    ExecutaGrid4 = modOrcamentoLogic.ExecutaGrid4(Me, ColumnValue, vgOq, ByVal, ByVal, vgIsValid, ByRef, vgErrorMessage, KeyCodeAscii, Shift)
    modOrcamentoLogic.PegaCoresGrid0 Me, vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
    modOrcamentoLogic.PegaCoresGrid1 Me, vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
    modOrcamentoLogic.PegaCoresGrid3 Me, vgItem, vgSubItem, vgTextColor, vgBackColor, vgSelectTextColor, vgSelectBakColor, vgColumnTextColor, vgColumnBackColor
    modOrcamentoLogic.Grid_GetColumnDefaultValue Me, Index, vgCol, vgColumns, vgDefaultValue
    modOrcamentoLogic.Grid_GetColumnLocked Me, Index, vgRow, vgCol, vgColumns, FormField, vgLocked
    modOrcamentoLogic.timUnLoad_Timer Me
    modOrcamentoLogic.Botao_Click Me, Index
    modOrcamentoLogic.txtCampo_ButtonClick Me, Index, Cancel
    modOrcamentoLogic.bottxtCampo158_Click Me, Index
    modOrcamentoLogic.bottxtCampo156_Click Me, Index
    modOrcamentoLogic.bottxtCampo157_Click Me, Index
    modOrcamentoLogic.bottxtCampo17_Click Me, Index
    modOrcamentoLogic.bottxtCampo16_Click Me, Index
    modOrcamentoLogic.bottxtCampo33_Click Me, Index
    modOrcamentoLogic.bottxtCampo55_Click Me, Index
    modOrcamentoLogic.bottxtCampo59_Click Me, Index
    modOrcamentoLogic.bottxtCampo60_Click Me, Index
    modOrcamentoLogic.bottxtCampo34_Click Me, Index
    modOrcamentoLogic.bottxtCampo35_Click Me, Index
    modOrcamentoLogic.bottxtCampo32_Click Me, Index
    modOrcamentoLogic.bottxtCampo45_Click Me, Index
    modOrcamentoLogic.bottxtCampo47_Click Me, Index
    modOrcamentoLogic.bottxtCampo57_Click Me, Index
    modOrcamentoLogic.bottxtCampo62_Click Me, Index
    modOrcamentoLogic.bottxtCampo137_Click Me, Index
    modOrcamentoLogic.bottxtCampo64_Click Me, Index
    modOrcamentoLogic.bottxtCampo99_Click Me, Index
    modOrcamentoLogic.bottxtCampo100_Click Me, Index
    modOrcamentoLogic.bottxtCampo101_Click Me, Index
    modOrcamentoLogic.bottxtCampo147_Click Me, Index
    modOrcamentoLogic.bottxtCampo149_Click Me, Index
    modOrcamentoLogic.bottxtCampo150_Click Me, Index
    modOrcamentoLogic.bottxtCampo161_Click Me, Index
    modOrcamentoLogic.bottxtCampo151_Click Me, Index
    modOrcamentoLogic.txtCp_MouseDown Me, Index, Button, Shift, x, y
         Set vgRsError = Produtos_do_Orcamento
         If ProcessaProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal) Then
            Produtos_do_Orcamento![Seqüência do Orçamento] = (0)
            Sequencia_do_Orcamento = Produtos_do_Orcamento![Seqüência do Orçamento]
         End If
         Produtos_do_Orcamento.Update
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
   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(3)
   CFOP = ColumnValue(4)
   Peso = ColumnValue(6)
   Quantidade = ColumnValue(7)
   PesoTotal = ColumnValue(9)
   Valor_Unitario = ColumnValue(10)
   Valor_do_Desconto = ColumnValue(11)
   Valor_do_Frete = ColumnValue(12)
   Valor_da_Base_de_Calculo = ColumnValue(14)
   Valor_do_ICMS = ColumnValue(15)
   Valor_do_IPI = ColumnValue(16)
   Aliquota_do_ICMS = ColumnValue(17)
   Aliquota_do_IPI = ColumnValue(18)
   Diferido = ColumnValue(19)
   Percentual_da_Reducao = ColumnValue(20)
   IVA = ColumnValue(21)
   Base_de_Calculo_ST = ColumnValue(22)
   Valor_ICMS_ST = ColumnValue(23)
   Aliquota_do_ICMS_ST = ColumnValue(24)
   Bc_pis = ColumnValue(25)
   Aliq_do_pis = ColumnValue(26)
   Valor_do_PIS = ColumnValue(27)
   Bc_cofins = ColumnValue(28)
   Aliq_do_cofins = ColumnValue(29)
   Valor_do_Cofins = ColumnValue(30)
   Valor_do_Tributo = ColumnValue(31)
   If Grid(3).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqüência do Produto Orçamento]
         Valor_Total = Produtos_do_Orcamento![Valor Total]
         Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
      Bc_pis = Produtos_do_Orcamento![Bc Pis]
      Valor_do_Tributo = Produtos_do_Orcamento![Valor Do Tributo]
      Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqüência do Produto Orçamento]
      Sequencia_do_Produto = Produtos_do_Orcamento![Seqüência do Produto]
      Quantidade = Produtos_do_Orcamento!Quantidade
      Valor_Unitario = Produtos_do_Orcamento![Valor Unitário]
      Valor_Total = Produtos_do_Orcamento![Valor Total]
      Valor_do_IPI = Produtos_do_Orcamento![Valor do IPI]
      Valor_do_ICMS = Produtos_do_Orcamento![Valor Do Icms]
      Aliquota_do_IPI = Produtos_do_Orcamento![Alíquota Do IPI]
      Aliquota_do_ICMS = Produtos_do_Orcamento![Alíquota Do ICMS]
      Percentual_da_Reducao = Produtos_do_Orcamento![Percentual da Redução]
      Diferido = Produtos_do_Orcamento!Diferido
      Valor_da_Base_de_Calculo = Produtos_do_Orcamento![Valor da Base de Cálculo]
      Valor_do_PIS = Produtos_do_Orcamento![Valor Do PIS]
      Valor_do_Cofins = Produtos_do_Orcamento![Valor Do Cofins]
      IVA = Produtos_do_Orcamento!IVA
      Base_de_Calculo_ST = Produtos_do_Orcamento![Base de Cálculo ST]
      Valor_ICMS_ST = Produtos_do_Orcamento![Valor ICMS ST]
      CFOP = Produtos_do_Orcamento!CFOP
      CST = Produtos_do_Orcamento!CST
      Aliquota_do_ICMS_ST = Produtos_do_Orcamento![Alíquota Do ICMS ST]
      Valor_do_Frete = Produtos_do_Orcamento![Valor Do Frete]
      Valor_do_Desconto = Produtos_do_Orcamento![Valor Do Desconto]
      Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      Bc_cofins = Produtos_do_Orcamento![Bc Cofins]
      Aliq_do_pis = Produtos_do_Orcamento![Aliq Do Pis]
      Aliq_do_cofins = Produtos_do_Orcamento![Aliq Do Cofins]
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade
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
   Dim vgRetVal As Variant, vgRsError As GRecordSet, x As String, vgNVez As Integer
   Dim Sequencia_do_Servico_Orcamento As Long, Sequencia_do_Servico As Integer, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_Anterior As Double
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
         Case 4
            vgRetVal = (Quantidade * Valor_Unitario)
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
         Case 5
            vgRetVal = (((Quantidade * Valor_Unitario) * (Orcamento![Alíquota Do ISS]) / 100))
            vgColumn = -1                   'flag para permitir definição desse novo valor para coluna do grid
      End Select
      If Err Then Err.Clear
   ElseIf vgOq = KEYPRESS_NO_GRID Then
      GoSub IniApDaCol
      ComandosServicos KeyCodeAscii, Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior
   ElseIf vgOq = CONDICOES_ESPECIAIS Then
      If vgSituacao <> ACAO_INCLUINDO Then
         GoSub IniApDaTb
      On Error Resume Next
         Grid(4).AllowInsert = ((Sequencia_do_Pedido = 0) And Cancelado = 0)
      On Error Resume Next
         Grid(4).AllowEdit = ((Sequencia_do_Pedido = 0) And Cancelado = 0)
      On Error Resume Next
         Grid(4).AllowDelete = ((Sequencia_do_Pedido = 0) And Cancelado = 0 And vgPWUsuario <> "RAFAEL")
      End If
      vgRetVal = ""
   ElseIf vgOq = ABRE_TABELA_GRID Then
      On Error Resume Next
      vgRetVal = "SELECT * FROM [Serviços do Orçamento]"

      'definindo a expressão de ligação com o pai
      x$ = "[Seqüência do Orçamento] = " & Orcamento![Seqüência do Orçamento]
      vgRetVal = InsereSQL(vgRetVal, EXP_WHERE, x$)

   ElseIf vgOq = DEFAULTDASCOLUNAS Then
      GoSub IniApDaCol
      vgRetVal = Null
      Select Case vgCol
         Case 3
            Valor_Unitario = InfoServicos(Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior, "Valor Unitário")
            If Grid(4).Col = 3 Then
               vgRetVal = Valor_Unitario
            End If
      End Select
   ElseIf vgOq = PEGAFILTRODASCOLUNAS Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
            vgRetVal = "" & IIf(vgSituacao <> ACAO_NAVEGANDO, "[Seqüência do Serviço] > 0 AND Inativo = 0", "[Seqüência do Serviço] > 0") & ""
      End Select
   ElseIf vgOq = PEGAEXPRESSAOPESQUISA Then
      On Error Resume Next
      GoSub IniApDaCol
      Select Case vgCol
         Case 1
                                    vgRetVal = "SELECT Serviços.[Seqüência do Serviço], Serviços.Descrição FROM Serviços WHERE (Serviços.[Seqüência do Serviço] > " + CStr(0) + ") AND " + _
                                                  "(Serviços.Inativo = False)"
      End Select
   Else
      If vgOq = VALIDACOES Then
         GoSub IniApDaCol
         vgIsValid = (Sequencia_do_Servico > 0)
         If Not vgIsValid Then vgColumn = 1
         vgErrorMessage$ = "Seqüência do Serviço inválido!"
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Quantidade > 0)
            If Not vgIsValid Then vgColumn = 2
            vgErrorMessage$ = "Quantidade inválido!"
         End If
         If vgIsValid And vgCol = -1 Then
            vgIsValid = (Valor_Unitario > 0)
            If Not vgIsValid Then vgColumn = 3
            vgErrorMessage$ = "Valor Unitário inválido!"
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
         Servicos_do_Orcamento.Edit
         Set vgRsError = Servicos_do_Orcamento
         If ProcessaServicos(Sequencia_do_Orcamento, Sequencia_do_Servico_Orcamento, Sequencia_do_Servico, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_Anterior) Then
            Servicos_do_Orcamento![Seqüência do Orçamento] = (0)
            Sequencia_do_Orcamento = Servicos_do_Orcamento![Seqüência do Orçamento]
         End If
         Servicos_do_Orcamento.Update
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
   Quantidade = ColumnValue(2)
   Valor_Unitario = ColumnValue(3)
   If Grid(4).Status <> ACAO_INCLUINDO Then
      If Servicos_do_Orcamento.EOF = False And Servicos_do_Orcamento.BOF = False And Servicos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Servico_Orcamento = Servicos_do_Orcamento![Seqüência do Serviço Orçamento]
         Valor_Total = Servicos_do_Orcamento![Valor Total]
         Valor_Anterior = Servicos_do_Orcamento![Valor Anterior]
      End If
   End If
   If Err Then Err.Clear
   On Error GoTo DeuErro
   Return

IniApDaTb:
   On Error Resume Next
   If Servicos_do_Orcamento.EOF = False And Servicos_do_Orcamento.BOF = False And Servicos_do_Orcamento.RecordCount > 0 Then
      Sequencia_do_Servico_Orcamento = Servicos_do_Orcamento![Seqüência do Serviço Orçamento]
      Sequencia_do_Servico = Servicos_do_Orcamento![Seqüência do Serviço]
      Quantidade = Servicos_do_Orcamento!Quantidade
      Valor_Unitario = Servicos_do_Orcamento![Valor Unitário]
      Valor_Total = Servicos_do_Orcamento![Valor Total]
      Valor_Anterior = Servicos_do_Orcamento![Valor Anterior]
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

'pega definições de cores para o grid
Private Sub PegaCoresGrid0(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Sequencia_Conjunto_Orcamento As Long, Sequencia_do_Conjunto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single
   Dim Percentual_da_Reducao As Single, Diferido As Boolean, Valor_da_Base_de_Calculo As Double
   Dim Valor_do_Tributo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, CFOP As Integer
   Dim CST As Integer, Valor_ICMS_ST As Double, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(0).GetColumnValues(vgItem)

   Sequencia_do_Conjunto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Quantidade = ColumnValue(5)
   Valor_Unitario = ColumnValue(7)
   Valor_do_Desconto = ColumnValue(8)
   Valor_do_Frete = ColumnValue(9)
   Valor_da_Base_de_Calculo = ColumnValue(11)
   Valor_do_ICMS = ColumnValue(12)
   Valor_do_IPI = ColumnValue(13)
   Aliquota_do_ICMS = ColumnValue(14)
   Aliquota_do_IPI = ColumnValue(15)
   Diferido = ColumnValue(16)
   Percentual_da_Reducao = ColumnValue(17)
   IVA = ColumnValue(18)
   Base_de_Calculo_ST = ColumnValue(19)
   Valor_ICMS_ST = ColumnValue(20)
   Aliquota_do_ICMS_ST = ColumnValue(21)
   Bc_pis = ColumnValue(22)
   Aliq_do_pis = ColumnValue(23)
   Valor_do_PIS = ColumnValue(24)
   Bc_cofins = ColumnValue(25)
   Aliq_do_cofins = ColumnValue(26)
   Valor_do_Cofins = ColumnValue(27)
   Valor_do_Tributo = ColumnValue(28)
   If Grid(0).Status <> ACAO_INCLUINDO Then
      If Conjuntos_do_Orcamento.EOF = False And Conjuntos_do_Orcamento.BOF = False And Conjuntos_do_Orcamento.RecordCount > 0 Then
         Sequencia_Conjunto_Orcamento = Conjuntos_do_Orcamento![Seqüência Conjunto Orçamento]
         Valor_Total = Conjuntos_do_Orcamento![Valor Total]
         Valor_Anterior = Conjuntos_do_Orcamento![Valor Anterior]
      End If
   End If

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 6

         'Fundo coluna
         If (Quantidade > InfoConjuntos(Sequencia_do_Orcamento, Sequencia_Conjunto_Orcamento, Sequencia_do_Conjunto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, Diferido, _
   Valor_da_Base_de_Calculo, Valor_do_Tributo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   CFOP, CST, Valor_ICMS_ST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, "Estoque")) Then
            vgBackColor = &HC0C0FF
         End If

   End Select

   Exit Sub

DeuErro:
End Sub

'pega definições de cores para o grid
Private Sub PegaCoresGrid1(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Valor_do_Tributo As Double, Sequencia_do_Produto As Long, Quantidade As Double
   Dim Valor_Unitario As Double, Valor_Total As Double, Valor_do_IPI As Double
   Dim Valor_do_ICMS As Double, Sequencia_Pecas_do_Orcamento As Long, Aliquota_do_IPI As Double
   Dim Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single, Diferido As Boolean
   Dim Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double, Valor_do_Cofins As Double
   Dim IVA As Double, Base_de_Calculo_ST As Double, Valor_ICMS_ST As Double
   Dim CFOP As Integer, CST As Integer, Aliquota_do_ICMS_ST As Single
   Dim Valor_do_Desconto As Double, Valor_do_Frete As Double, Valor_Anterior As Double
   Dim Bc_pis As Double, Aliq_do_pis As Single, Bc_cofins As Double
   Dim Aliq_do_cofins As Single
   Dim Peso As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(1).GetColumnValues(vgItem)

   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(2)
   CFOP = ColumnValue(3)
   Peso = ColumnValue(5)
   Quantidade = ColumnValue(6)
   Valor_Unitario = ColumnValue(9)
   Valor_do_Desconto = ColumnValue(10)
   Valor_do_Frete = ColumnValue(11)
   Valor_da_Base_de_Calculo = ColumnValue(13)
   Valor_do_ICMS = ColumnValue(14)
   Valor_do_IPI = ColumnValue(15)
   Aliquota_do_ICMS = ColumnValue(16)
   Aliquota_do_IPI = ColumnValue(17)
   Diferido = ColumnValue(18)
   Percentual_da_Reducao = ColumnValue(19)
   IVA = ColumnValue(20)
   Base_de_Calculo_ST = ColumnValue(21)
   Valor_ICMS_ST = ColumnValue(22)
   Aliquota_do_ICMS_ST = ColumnValue(23)
   Bc_pis = ColumnValue(24)
   Aliq_do_pis = ColumnValue(25)
   Valor_do_PIS = ColumnValue(26)
   Bc_cofins = ColumnValue(27)
   Aliq_do_cofins = ColumnValue(28)
   Valor_do_Cofins = ColumnValue(29)
   Valor_do_Tributo = ColumnValue(30)
   If Grid(1).Status <> ACAO_INCLUINDO Then
      If Pecas_do_Orcamento.EOF = False And Pecas_do_Orcamento.BOF = False And Pecas_do_Orcamento.RecordCount > 0 Then
         Valor_Total = Pecas_do_Orcamento![Valor Total]
         Sequencia_Pecas_do_Orcamento = Pecas_do_Orcamento![Seqüência Peças do Orçamento]
         Valor_Anterior = Pecas_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Peso")

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 7

         'Fundo coluna
         If (Quantidade > InfoPecas(Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto, Quantidade, Valor_Unitario, Valor_Total, _
   Valor_do_IPI, Valor_do_ICMS, Sequencia_Pecas_do_Orcamento, Aliquota_do_IPI, Aliquota_do_ICMS, Percentual_da_Reducao, _
   Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, Base_de_Calculo_ST, _
   Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Desconto, Valor_do_Frete, _
   Valor_Anterior, Bc_pis, Aliq_do_pis, Bc_cofins, Aliq_do_cofins, Peso, "Estoque")) Then
            vgBackColor = &HC0C0FF
         End If

   End Select

   Exit Sub

DeuErro:
End Sub

'pega definições de cores para o grid
Private Sub PegaCoresGrid3(ByVal vgItem As Long, ByVal vgSubItem As Long, vgTextColor As Long, vgBackColor As Long, vgSelectTextColor As Long, vgSelectBakColor As Long, vgColumnTextColor As Long, vgColumnBackColor As Long)
   Dim Bc_pis As Double, Valor_do_Tributo As Double, Sequencia_do_Produto_Orcamento As Long
   Dim Sequencia_do_Produto As Long, Quantidade As Double, Valor_Unitario As Double
   Dim Valor_Total As Double, Valor_do_IPI As Double, Valor_do_ICMS As Double
   Dim Aliquota_do_IPI As Double, Aliquota_do_ICMS As Single, Percentual_da_Reducao As Single
   Dim Diferido As Boolean, Valor_da_Base_de_Calculo As Double, Valor_do_PIS As Double
   Dim Valor_do_Cofins As Double, IVA As Double, Base_de_Calculo_ST As Double
   Dim Valor_ICMS_ST As Double, CFOP As Integer, CST As Integer
   Dim Aliquota_do_ICMS_ST As Single, Valor_do_Frete As Double, Valor_do_Desconto As Double
   Dim Valor_Anterior As Double, Bc_cofins As Double, Aliq_do_pis As Single
   Dim Aliq_do_cofins As Single
   Dim Peso As Double, PesoTotal As Double
   Dim ColumnValue() As Variant

   On Error GoTo DeuErro

   ColumnValue = Grid(3).GetColumnValues(vgItem)

   Sequencia_do_Produto = Val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   CST = ColumnValue(3)
   CFOP = ColumnValue(4)
   Peso = ColumnValue(6)
   Quantidade = ColumnValue(7)
   PesoTotal = ColumnValue(9)
   Valor_Unitario = ColumnValue(10)
   Valor_do_Desconto = ColumnValue(11)
   Valor_do_Frete = ColumnValue(12)
   Valor_da_Base_de_Calculo = ColumnValue(14)
   Valor_do_ICMS = ColumnValue(15)
   Valor_do_IPI = ColumnValue(16)
   Aliquota_do_ICMS = ColumnValue(17)
   Aliquota_do_IPI = ColumnValue(18)
   Diferido = ColumnValue(19)
   Percentual_da_Reducao = ColumnValue(20)
   IVA = ColumnValue(21)
   Base_de_Calculo_ST = ColumnValue(22)
   Valor_ICMS_ST = ColumnValue(23)
   Aliquota_do_ICMS_ST = ColumnValue(24)
   Bc_pis = ColumnValue(25)
   Aliq_do_pis = ColumnValue(26)
   Valor_do_PIS = ColumnValue(27)
   Bc_cofins = ColumnValue(28)
   Aliq_do_cofins = ColumnValue(29)
   Valor_do_Cofins = ColumnValue(30)
   Valor_do_Tributo = ColumnValue(31)
   If Grid(3).Status <> ACAO_INCLUINDO Then
      If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
         Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqüência do Produto Orçamento]
         Valor_Total = Produtos_do_Orcamento![Valor Total]
         Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      End If
   End If
   Peso = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Peso")
   PesoTotal = Peso * Quantidade

   'Vamos definir cores individuais para essas colunas
   Select Case vgSubItem

      Case 8

         'Fundo coluna
         If (Quantidade > InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Estoque")) Then
            vgBackColor = &HC0C0FF
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
         PreValidaNFE
      Case 1
         seqRegistro = Sequencia_do_Vendedor
         mdiIRRIG.MGeral
      Case 2
         seqRegistro = Sequencia_da_Classificacao
         mdiIRRIG.MClassifi
      Case 3
         seqRegistro = IIf(Sequencia_do_Geral > 0, MunicipioAux![Seqüência Do Município], Sequencia_do_Municipio)
         mdiIRRIG.MMuni
      Case 4
         seqRegistro = Sequencia_da_Transportadora
         mdiIRRIG.MGeral
      Case 5
         seqRegistro = Sequencia_do_Pais
         mdiIRRIG.MPaises
      Case 6
         AtualizaValor
      Case 7
         Parcelar
      Case 8
         AbreNotaFiscal
      Case 9
         PreValidaImpressao
      Case 10
         seqRegistro = Sequencia_do_Geral
         mdiIRRIG.MGeral
      Case 12
         AbrePropriedades Sequencia_do_Geral, Sequencia_da_Propriedade
      Case 13
         mdiIRRIG.MRelfxhid
         AtivaForm Me
      Case 14
         mdiIRRIG.MRelorcfa
         AtivaForm Me
      Case 15
         AbreProjeto
      Case 16
         mdiIRRIG.MEmiAlmox
         AtivaForm Me
   End Select
End Sub

'evento - quando o botão for pressionado
Private Sub txtCampo_ButtonClick(Index As Integer, Cancel As Boolean)
   If vgPriVez Then Exit Sub
   Select Case Index
      Case 16
         AbreComando "mailto:" + Email
   End Select
End Sub


'evento - quando o botão for apertado
Private Sub bottxtCampo158_Click(Index As Integer)
   txtCampo(158).SetFocus
   DoEvents
   txtCampo(158).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo156_Click(Index As Integer)
   txtCampo(156).SetFocus
   DoEvents
   txtCampo(156).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo157_Click(Index As Integer)
   txtCampo(157).SetFocus
   DoEvents
   txtCampo(157).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo17_Click(Index As Integer)
   txtCampo(17).SetFocus
   DoEvents
   txtCampo(17).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo16_Click(Index As Integer)
   Dim Cancel As Boolean
   If Index = BOT_ACAO Then
      txtCampo_ButtonClick 16, Cancel
      Exit Sub
   End If
   DoEvents
   If Not Cancel Then
      txtCampo(16).BotClick Index
   End If
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo33_Click(Index As Integer)
   txtCampo(33).SetFocus
   DoEvents
   txtCampo(33).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo55_Click(Index As Integer)
   txtCampo(55).SetFocus
   DoEvents
   txtCampo(55).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo59_Click(Index As Integer)
   txtCampo(59).SetFocus
   DoEvents
   txtCampo(59).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo60_Click(Index As Integer)
   txtCampo(60).SetFocus
   DoEvents
   txtCampo(60).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo34_Click(Index As Integer)
   txtCampo(34).SetFocus
   DoEvents
   txtCampo(34).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo35_Click(Index As Integer)
   txtCampo(35).SetFocus
   DoEvents
   txtCampo(35).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo32_Click(Index As Integer)
   txtCampo(32).SetFocus
   DoEvents
   txtCampo(32).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo45_Click(Index As Integer)
   txtCampo(45).SetFocus
   DoEvents
   txtCampo(45).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo47_Click(Index As Integer)
   txtCampo(47).SetFocus
   DoEvents
   txtCampo(47).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo57_Click(Index As Integer)
   txtCampo(57).SetFocus
   DoEvents
   txtCampo(57).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo62_Click(Index As Integer)
   txtCampo(62).SetFocus
   DoEvents
   txtCampo(62).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo137_Click(Index As Integer)
   txtCampo(137).SetFocus
   DoEvents
   txtCampo(137).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo64_Click(Index As Integer)
   txtCampo(64).SetFocus
   DoEvents
   txtCampo(64).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo99_Click(Index As Integer)
   txtCampo(99).SetFocus
   DoEvents
   txtCampo(99).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo100_Click(Index As Integer)
   txtCampo(100).SetFocus
   DoEvents
   txtCampo(100).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo101_Click(Index As Integer)
   txtCampo(101).SetFocus
   DoEvents
   txtCampo(101).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo147_Click(Index As Integer)
   txtCampo(147).SetFocus
   DoEvents
   txtCampo(147).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo149_Click(Index As Integer)
   txtCampo(149).SetFocus
   DoEvents
   txtCampo(149).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo150_Click(Index As Integer)
   txtCampo(150).SetFocus
   DoEvents
   txtCampo(150).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo161_Click(Index As Integer)
   txtCampo(161).SetFocus
   DoEvents
   txtCampo(161).BotClick Index
End Sub

'evento - quando o botão for apertado
Private Sub bottxtCampo151_Click(Index As Integer)
   txtCampo(151).SetFocus
   DoEvents
   txtCampo(151).BotClick Index
End Sub

'evento - quando o mouse for pressionado sobre o campo
Private Sub txtCp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   txtCampo(Index).MouseDown
End Sub


