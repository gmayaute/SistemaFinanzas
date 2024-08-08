VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmPrograma 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   4815
   ClientTop       =   4230
   ClientWidth     =   10530
   Icon            =   "frmPrograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10530
   Begin TabDlg.SSTab SSTab1 
      Height          =   6030
      Left            =   0
      TabIndex        =   19
      Top             =   -30
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10636
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   1058
      BackColor       =   10442041
      TabCaption(0)   =   "   Traslados y Movilidades"
      TabPicture(0)   =   "frmPrograma.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "txtEmpleado_1"
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(6)=   "lblEmpleado_1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "   Salidas Diversas"
      TabPicture(1)   =   "frmPrograma.frx":0A24
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameSub"
      Tab(1).Control(1)=   "optFrame(3)"
      Tab(1).Control(2)=   "frameVac"
      Tab(1).Control(3)=   "framePer"
      Tab(1).Control(4)=   "optFrame(2)"
      Tab(1).Control(5)=   "optFrame(1)"
      Tab(1).Control(6)=   "optFrame(0)"
      Tab(1).Control(7)=   "frameLic"
      Tab(1).Control(8)=   "txtEmpleado_2"
      Tab(1).Control(9)=   "lblEmpleado_2"
      Tab(1).Control(10)=   "Label17"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   " Refrigerios"
      TabPicture(2)   =   "frmPrograma.frx":0D3E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Label34"
      Tab(2).Control(2)=   "lblEmpleado_3"
      Tab(2).Control(3)=   "txtEmpleado_3"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Bonos"
      TabPicture(3)   =   "frmPrograma.frx":1190
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "txtEmpleado_4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblEmpleado_4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame1 
         BackColor       =   &H009F5539&
         Caption         =   "Bonos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   4875
         Left            =   0
         TabIndex        =   129
         Top             =   1080
         Width           =   10515
         Begin VB.Frame Frame6 
            BackColor       =   &H009F5539&
            Caption         =   "Destino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   1665
            Left            =   120
            TabIndex        =   136
            Top             =   1200
            Width           =   9945
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H009F5539&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Lote"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   2820
               TabIndex        =   144
               Top             =   360
               Width           =   525
            End
            Begin MSForms.ComboBox cmbLoteB 
               Height          =   315
               Left            =   3435
               TabIndex        =   143
               Top             =   360
               Width           =   6300
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               ForeColor       =   255
               DisplayStyle    =   7
               Size            =   "11112;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackColor       =   &H009F5539&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pozo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   90
               TabIndex        =   142
               Top             =   780
               Width           =   780
            End
            Begin MSForms.ComboBox cmbPozoB 
               Height          =   315
               Left            =   930
               TabIndex        =   141
               Top             =   780
               Width           =   8775
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               ForeColor       =   255
               DisplayStyle    =   7
               Size            =   "15478;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackColor       =   &H009F5539&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dpto."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   90
               TabIndex        =   140
               Top             =   375
               Width           =   780
            End
            Begin MSForms.ComboBox cmbDepartB 
               Height          =   315
               Left            =   900
               TabIndex        =   139
               Top             =   360
               Width           =   1755
               VariousPropertyBits=   612386843
               BackColor       =   16777215
               ForeColor       =   255
               DisplayStyle    =   7
               Size            =   "3096;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox CboDivB 
               Height          =   315
               Left            =   900
               TabIndex        =   138
               Top             =   1200
               Width           =   8835
               VariousPropertyBits=   746604571
               BackColor       =   16777215
               ForeColor       =   255
               DisplayStyle    =   7
               Size            =   "15584;556"
               MatchEntry      =   1
               ShowDropButtonWhen=   2
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H009F5539&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "División"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   90
               TabIndex        =   137
               Top             =   1215
               Width           =   780
            End
         End
         Begin MSComCtl2.DTPicker dpIni_bon 
            Height          =   315
            Left            =   960
            TabIndex        =   130
            Top             =   240
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dpFin_bon 
            Height          =   315
            Left            =   990
            TabIndex        =   131
            Top             =   630
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin VB.Label Label44 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   135
            Top             =   630
            Width           =   795
         End
         Begin VB.Label Label43 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   134
            Top             =   270
            Width           =   795
         End
         Begin MSForms.TextBox txtMotivo_bon 
            Height          =   795
            Left            =   3600
            TabIndex        =   133
            Top             =   240
            Width           =   6495
            VariousPropertyBits=   -1400879077
            MaxLength       =   300
            Size            =   "11456;1402"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label42 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Detalle:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2760
            TabIndex        =   132
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frameSub 
         BackColor       =   &H009F5539&
         Caption         =   "Subsidios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1260
         Left            =   -74685
         TabIndex        =   104
         Top             =   4725
         Width           =   10155
         Begin MSComCtl2.DTPicker dtinisub 
            Height          =   315
            Left            =   705
            TabIndex        =   105
            Top             =   270
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dtfinsub 
            Height          =   315
            Left            =   705
            TabIndex        =   106
            Top             =   630
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSForms.ComboBox CboTipo 
            Height          =   315
            Left            =   2655
            TabIndex        =   114
            Top             =   270
            Width           =   5460
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "9631;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label41 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "T.Susp."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1965
            TabIndex        =   113
            Top             =   270
            Width           =   675
         End
         Begin VB.Label Label40 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   112
            Top             =   630
            Width           =   615
         End
         Begin VB.Label Label39 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   111
            Top             =   270
            Width           =   615
         End
         Begin MSForms.TextBox txtmotivosub 
            Height          =   540
            Left            =   2655
            TabIndex        =   110
            Top             =   630
            Width           =   6945
            VariousPropertyBits=   -1400879077
            MaxLength       =   300
            Size            =   "12250;952"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label38 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Motivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1950
            TabIndex        =   109
            Top             =   630
            Width           =   690
         End
         Begin VB.Label Label18 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CIIT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8130
            TabIndex        =   108
            Top             =   270
            Width           =   450
         End
         Begin MSForms.TextBox txtciit 
            Height          =   315
            Left            =   8610
            TabIndex        =   107
            Top             =   270
            Width           =   1485
            VariousPropertyBits=   746604571
            ForeColor       =   0
            MaxLength       =   20
            Size            =   "2619;556"
            Value           =   "F1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.OptionButton optFrame 
         Caption         =   "Option1"
         Height          =   375
         Index           =   3
         Left            =   -74880
         TabIndex        =   103
         Top             =   5190
         Width           =   210
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H009F5539&
         Caption         =   "Refrigerios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   4845
         Left            =   -74940
         TabIndex        =   88
         Top             =   1110
         Width           =   10365
         Begin VB.TextBox txtNroFac 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1410
            MaxLength       =   12
            TabIndex        =   101
            Top             =   1350
            Width           =   1605
         End
         Begin VB.TextBox txtValorRefri 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1410
            MaxLength       =   12
            TabIndex        =   93
            Text            =   "0.00"
            Top             =   1800
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dptFecIniR 
            Height          =   315
            Left            =   1410
            TabIndex        =   89
            Top             =   330
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dptFecFinR 
            Height          =   315
            Left            =   4830
            TabIndex        =   90
            Top             =   330
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin VB.Label Label35 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nro. Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   100
            Top             =   1350
            Width           =   1185
         End
         Begin MSForms.ComboBox cboRest 
            Height          =   315
            Left            =   1410
            TabIndex        =   99
            Top             =   870
            Width           =   4725
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "8334;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label35 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Restaurant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   900
            Width           =   1185
         End
         Begin MSForms.CheckBox chkRefri2 
            Height          =   285
            Index           =   1
            Left            =   5430
            TabIndex        =   96
            Top             =   1800
            Width           =   765
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8454143
            DisplayStyle    =   4
            Size            =   "1349;503"
            Value           =   "0"
            Caption         =   "Cena"
            SpecialEffect   =   0
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.CheckBox chkRefri2 
            Height          =   285
            Index           =   0
            Left            =   3510
            TabIndex        =   95
            Top             =   1800
            Width           =   1215
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8454143
            DisplayStyle    =   4
            Size            =   "2143;503"
            Value           =   "0"
            Caption         =   "Almuerzo"
            SpecialEffect   =   0
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label35 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   94
            Top             =   1800
            Width           =   1185
         End
         Begin VB.Label Label37 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label36 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3480
            TabIndex        =   91
            Top             =   330
            Width           =   1215
         End
      End
      Begin VB.Frame frameVac 
         BackColor       =   &H009F5539&
         Caption         =   "Vacaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1050
         Left            =   -74700
         TabIndex        =   45
         Top             =   1110
         Width           =   10185
         Begin VB.ComboBox cmbPeriodo 
            Height          =   315
            Left            =   3990
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   630
            Width           =   2385
         End
         Begin VB.TextBox txtDias 
            Height          =   315
            Left            =   5640
            MaxLength       =   2
            TabIndex        =   26
            Text            =   "0"
            Top             =   285
            Width           =   435
         End
         Begin VB.TextBox txtMeses 
            Height          =   315
            Left            =   3990
            MaxLength       =   2
            TabIndex        =   25
            Text            =   "0"
            Top             =   285
            Width           =   405
         End
         Begin MSComCtl2.DTPicker dpIni_vac 
            Height          =   315
            Left            =   1020
            TabIndex        =   23
            Top             =   285
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dpFin_vac 
            Height          =   315
            Left            =   1020
            TabIndex        =   24
            Top             =   630
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin Proyecto1.chameleonButton cmdgentareo 
            Height          =   360
            Left            =   8460
            TabIndex        =   102
            ToolTipText     =   "Carga Ingreso Seleccionado"
            Top             =   630
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
            BTYPE           =   14
            TX              =   "Generar Tareo"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8421504
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmPrograma.frx":11AC
            PICN            =   "frmPrograma.frx":11C8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSForms.CheckBox chkIndemnizacion 
            Height          =   315
            Left            =   6420
            TabIndex        =   125
            Top             =   600
            Width           =   2505
            BackColor       =   10442041
            ForeColor       =   8421631
            DisplayStyle    =   4
            Size            =   "4419;556"
            Value           =   "0"
            Caption         =   "Indemnizacion"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.SpinButton sbMeses 
            Height          =   375
            Left            =   4440
            TabIndex        =   61
            Top             =   255
            Width           =   285
            Size            =   "503;661"
            Max             =   12
            Position        =   1
         End
         Begin MSForms.SpinButton sbDias 
            Height          =   375
            Left            =   6090
            TabIndex        =   60
            Top             =   255
            Width           =   285
            Size            =   "503;661"
            Max             =   31
            Position        =   1
         End
         Begin VB.Label Label27 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Meses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3180
            TabIndex        =   59
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label26 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4890
            TabIndex        =   58
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label25 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3180
            TabIndex        =   57
            Top             =   630
            Width           =   765
         End
         Begin VB.Label Label20 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   47
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label19 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   630
            Width           =   795
         End
         Begin MSForms.CheckBox chkGoceHaber 
            Height          =   315
            Left            =   6420
            TabIndex        =   76
            Top             =   270
            Width           =   2505
            BackColor       =   10442041
            ForeColor       =   8421631
            DisplayStyle    =   4
            Size            =   "4419;556"
            Value           =   "1"
            Caption         =   "Compra de Vacaciones"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.Frame framePer 
         BackColor       =   &H009F5539&
         Caption         =   "Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1290
         Left            =   -74700
         TabIndex        =   51
         Top             =   2160
         Width           =   10185
         Begin MSComCtl2.DTPicker dpIni_per 
            Height          =   315
            Left            =   990
            TabIndex        =   28
            Top             =   285
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dpFin_per 
            Height          =   315
            Left            =   990
            TabIndex        =   29
            Top             =   645
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSMask.MaskEdBox meHoraIni 
            Height          =   330
            Left            =   3120
            TabIndex        =   30
            Top             =   270
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm AM/PM"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox meHoraFin 
            Height          =   330
            Left            =   3120
            TabIndex        =   31
            Top             =   630
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm AM/PM"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin MSForms.TextBox txtAutoriz_per 
            Height          =   315
            Left            =   5625
            TabIndex        =   32
            Top             =   285
            Width           =   4485
            VariousPropertyBits=   746604571
            ForeColor       =   0
            MaxLength       =   11
            Size            =   "7911;556"
            Value           =   "F1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label31 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Autorizado Por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4140
            TabIndex        =   65
            Top             =   285
            Width           =   1425
         End
         Begin VB.Label Label30 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2370
            TabIndex        =   64
            Top             =   645
            Width           =   705
         End
         Begin VB.Label Label29 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2370
            TabIndex        =   63
            Top             =   285
            Width           =   705
         End
         Begin VB.Label Label28 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Motivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4140
            TabIndex        =   62
            Top             =   645
            Width           =   735
         End
         Begin MSForms.TextBox txtMotivo_per 
            Height          =   570
            Left            =   4920
            TabIndex        =   33
            Top             =   645
            Width           =   5205
            VariousPropertyBits=   -1400879077
            MaxLength       =   300
            Size            =   "9181;1005"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label22 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   645
            Width           =   765
         End
         Begin VB.Label Label21 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   52
            Top             =   285
            Width           =   795
         End
      End
      Begin VB.OptionButton optFrame 
         Caption         =   "Option1"
         Height          =   375
         Index           =   2
         Left            =   -74888
         TabIndex        =   50
         Top             =   3900
         Width           =   210
      End
      Begin VB.OptionButton optFrame 
         Caption         =   "Option1"
         Height          =   375
         Index           =   1
         Left            =   -74910
         TabIndex        =   49
         Top             =   2625
         Width           =   255
      End
      Begin VB.OptionButton optFrame 
         Caption         =   "Option1"
         Height          =   375
         Index           =   0
         Left            =   -74910
         TabIndex        =   48
         Top             =   1448
         Width           =   255
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H009F5539&
         Caption         =   "Transporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   3240
         Left            =   -69465
         TabIndex        =   38
         Top             =   1065
         Width           =   4965
         Begin VB.OptionButton optIdaVuelta 
            BackColor       =   &H009F5539&
            Caption         =   "Ida y Vuelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1635
            TabIndex        =   80
            Top             =   2280
            Width           =   1365
         End
         Begin VB.OptionButton optIda 
            BackColor       =   &H009F5539&
            Caption         =   "Ida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1035
            TabIndex        =   79
            Top             =   2280
            Width           =   615
         End
         Begin VB.ComboBox cmbMonBoleto 
            Height          =   315
            ItemData        =   "frmPrograma.frx":1BDA
            Left            =   3075
            List            =   "frmPrograma.frx":1BE4
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2235
            Width           =   615
         End
         Begin MSMask.MaskEdBox meBoleto 
            Height          =   315
            Left            =   3705
            TabIndex        =   6
            Top             =   2235
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox meHoraSalida 
            Height          =   315
            Left            =   1035
            TabIndex        =   115
            Top             =   1788
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm AM/PM"
            Mask            =   "##:##"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   116
            Top             =   1788
            Width           =   825
         End
         Begin VB.Label Label16 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   42
            Top             =   2235
            Width           =   825
         End
         Begin MSForms.ComboBox cmbMedio 
            Height          =   315
            Left            =   1035
            TabIndex        =   2
            Top             =   450
            Width           =   3735
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "6588;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label13 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Medio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   41
            Top             =   450
            Width           =   825
         End
         Begin MSForms.ComboBox cmbLinea 
            Height          =   315
            Left            =   1035
            TabIndex        =   4
            Top             =   1320
            Width           =   3735
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "6588;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label12 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Línea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   40
            Top             =   1342
            Width           =   825
         End
         Begin MSForms.ComboBox cmbAgencia 
            Height          =   315
            Left            =   1035
            TabIndex        =   3
            Top             =   896
            Width           =   3735
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "6588;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Agencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   39
            Top             =   896
            Width           =   825
         End
      End
      Begin VB.Frame frameLic 
         BackColor       =   &H009F5539&
         Caption         =   "Licencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1275
         Left            =   -74700
         TabIndex        =   54
         Top             =   3465
         Width           =   10155
         Begin MSComCtl2.DTPicker dpIni_lic 
            Height          =   315
            Left            =   990
            TabIndex        =   34
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSComCtl2.DTPicker dpFin_lic 
            Height          =   315
            Left            =   990
            TabIndex        =   35
            Top             =   630
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin MSForms.TextBox txtAutoriz_lic 
            Height          =   315
            Left            =   4290
            TabIndex        =   36
            Top             =   270
            Width           =   5805
            VariousPropertyBits=   746604571
            ForeColor       =   0
            MaxLength       =   11
            Size            =   "10239;556"
            Value           =   "F1"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label33 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Autorizado Por:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2820
            TabIndex        =   67
            Top             =   270
            Width           =   1425
         End
         Begin VB.Label Label32 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Motivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2820
            TabIndex        =   66
            Top             =   630
            Width           =   735
         End
         Begin MSForms.TextBox txtMotivo_lic 
            Height          =   555
            Left            =   3600
            TabIndex        =   37
            Top             =   630
            Width           =   6495
            VariousPropertyBits=   -1400879077
            MaxLength       =   300
            Size            =   "11456;979"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label24 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inicio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label23 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Fin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   630
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H009F5539&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1830
         Left            =   -74955
         TabIndex        =   117
         Top             =   4155
         Width           =   10425
         Begin MSComCtl2.DTPicker dpsalida 
            Height          =   315
            Left            =   225
            TabIndex        =   120
            Top             =   1260
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Format          =   55640065
            CurrentDate     =   38597
         End
         Begin VB.Label lblcolor 
            Height          =   120
            Left            =   5625
            TabIndex        =   124
            Top             =   1455
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label lbltipo 
            BackColor       =   &H009F5539&
            Height          =   240
            Left            =   8355
            TabIndex        =   121
            Top             =   1290
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   75
            TabIndex        =   119
            Top             =   225
            Width           =   1440
         End
         Begin MSForms.TextBox txtObs 
            Height          =   915
            Left            =   1575
            TabIndex        =   118
            Top             =   225
            Width           =   8610
            VariousPropertyBits=   -1400879077
            MaxLength       =   300
            Size            =   "15187;1614"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H009F5539&
         Caption         =   "Estadía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1560
         Left            =   -74955
         TabIndex        =   72
         Top             =   2670
         Width           =   5505
         Begin VB.OptionButton optNoche 
            BackColor       =   &H009F5539&
            Caption         =   "Noche"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1080
            TabIndex        =   82
            Top             =   1155
            Width           =   945
         End
         Begin VB.OptionButton optEstadia 
            BackColor       =   &H009F5539&
            Caption         =   "Estadía"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2160
            TabIndex        =   81
            Top             =   1155
            Width           =   1065
         End
         Begin MSMask.MaskEdBox meEstancia 
            Height          =   315
            Left            =   3960
            TabIndex        =   9
            Top             =   1155
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.ComboBox cmbMonEstancia 
            Height          =   315
            ItemData        =   "frmPrograma.frx":1BEE
            Left            =   3330
            List            =   "frmPrograma.frx":1BF8
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1155
            Width           =   615
         End
         Begin MSForms.ComboBox cmbNombreEst 
            Height          =   315
            Left            =   1080
            TabIndex        =   78
            Top             =   720
            Width           =   4215
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "7435;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label14 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   75
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label15 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Estancia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   74
            Top             =   285
            Width           =   915
         End
         Begin MSForms.ComboBox cmbEstancia 
            Height          =   315
            Left            =   1080
            TabIndex        =   7
            Top             =   285
            Width           =   2265
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3995;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label11 
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   73
            Top             =   1155
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H009F5539&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1665
         Left            =   -74955
         TabIndex        =   68
         Top             =   1065
         Width           =   5505
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "División"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   123
            Top             =   1215
            Width           =   780
         End
         Begin MSForms.ComboBox CboDiv 
            Height          =   315
            Left            =   900
            TabIndex        =   122
            Top             =   1215
            Width           =   4290
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "7567;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbDepart 
            Height          =   315
            Left            =   900
            TabIndex        =   84
            Top             =   360
            Width           =   1755
            VariousPropertyBits=   612386843
            BackColor       =   16777215
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "3096;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dpto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   71
            Top             =   375
            Width           =   780
         End
         Begin MSForms.ComboBox cmbPozo 
            Height          =   315
            Left            =   930
            TabIndex        =   1
            Top             =   780
            Width           =   4290
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "7567;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Pozo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   90
            TabIndex        =   70
            Top             =   780
            Width           =   780
         End
         Begin MSForms.ComboBox cmbLote 
            Height          =   315
            Left            =   3435
            TabIndex        =   0
            Top             =   360
            Width           =   1755
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "3096;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lote"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   2820
            TabIndex        =   69
            Top             =   360
            Width           =   525
         End
      End
      Begin VB.Label lblEmpleado_4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3840
         TabIndex        =   128
         Top             =   720
         Width           =   6645
      End
      Begin VB.Label Label8 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Empleado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   127
         Top             =   720
         Width           =   1635
      End
      Begin MSForms.TextBox txtEmpleado_4 
         Height          =   315
         Left            =   1800
         TabIndex        =   126
         Top             =   720
         Width           =   2025
         VariousPropertyBits=   746604571
         ForeColor       =   0
         MaxLength       =   11
         Size            =   "3572;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtEmpleado_1 
         Height          =   315
         Left            =   -73290
         TabIndex        =   97
         Top             =   720
         Width           =   2025
         VariousPropertyBits=   746604571
         ForeColor       =   0
         MaxLength       =   11
         Size            =   "3572;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label34 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Empleado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74880
         TabIndex        =   87
         Top             =   750
         Width           =   1545
      End
      Begin VB.Label lblEmpleado_3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -71220
         TabIndex        =   86
         Top             =   750
         Width           =   6645
      End
      Begin MSForms.TextBox txtEmpleado_3 
         Height          =   315
         Left            =   -73290
         TabIndex        =   85
         Top             =   750
         Width           =   2025
         VariousPropertyBits=   746604571
         ForeColor       =   0
         MaxLength       =   11
         Size            =   "3572;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtEmpleado_2 
         Height          =   315
         Left            =   -73260
         TabIndex        =   22
         Top             =   780
         Width           =   2025
         VariousPropertyBits=   746604571
         ForeColor       =   0
         MaxLength       =   11
         Size            =   "3572;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblEmpleado_2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -71190
         TabIndex        =   44
         Top             =   780
         Width           =   6645
      End
      Begin VB.Label Label17 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Empleado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74850
         TabIndex        =   43
         Top             =   780
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cod. Empleado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -75000
         TabIndex        =   21
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label lblEmpleado_1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -71190
         TabIndex        =   20
         Top             =   720
         Width           =   6525
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   555
      Left            =   30
      TabIndex        =   11
      Top             =   6030
      Width           =   10515
      Begin Proyecto1.chameleonButton btnSalir 
         Height          =   345
         Left            =   9960
         TabIndex        =   12
         Top             =   150
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":1C02
         PICN            =   "frmPrograma.frx":1C1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnGrabar 
         Height          =   345
         Left            =   5100
         TabIndex        =   10
         ToolTipText     =   "Guardar"
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Grabar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":1FE4
         PICN            =   "frmPrograma.frx":2000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnmodificar 
         Height          =   345
         Left            =   1200
         TabIndex        =   13
         ToolTipText     =   "Modificar"
         Top             =   150
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Modificar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":2442
         PICN            =   "frmPrograma.frx":245E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnCancelar 
         Height          =   345
         Left            =   4620
         TabIndex        =   14
         ToolTipText     =   "Deshacer"
         Top             =   150
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":288C
         PICN            =   "frmPrograma.frx":28A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnEliminar 
         Height          =   345
         Left            =   2430
         TabIndex        =   16
         ToolTipText     =   "Eliminar"
         Top             =   150
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Eliminar"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":2DEA
         PICN            =   "frmPrograma.frx":2E06
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnNuevo 
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   150
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Nuevo"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":3248
         PICN            =   "frmPrograma.frx":3264
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnReporte 
         Height          =   345
         Left            =   9420
         TabIndex        =   18
         Top             =   150
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrograma.frx":35CE
         PICN            =   "frmPrograma.frx":35EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblcod 
         BackColor       =   &H009F5539&
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   3870
         TabIndex        =   83
         Top             =   210
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCodigoCalen 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGOGRABADO"
         Height          =   255
         Left            =   6960
         TabIndex        =   77
         Top             =   210
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Label lblModo 
      Height          =   195
      Left            =   8250
      TabIndex        =   15
      Top             =   6660
      Width           =   2175
   End
End
Attribute VB_Name = "frmPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As FrmConsultas
Private Tienebono As Boolean
Dim FlgFecha As Boolean

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
End Sub

Private Sub btnGrabar_Click()
    If lblModo = "Nuevo" Then
        If SSTab1.TabVisible(0) Then
            GrabarTab0
            Unload Me
        Else
            If SSTab1.TabVisible(1) Then If ValidarTab1 Then GrabarTab1
            If SSTab1.TabVisible(2) Then If ValidarTab2 Then GrabarTab2
            If SSTab1.TabVisible(3) Then If ValidarTab3 Then GrabarTab3
        End If
    Else
        If lblModo = "Modificar" Then
            If SSTab1.TabVisible(0) Then ActualizaTab0
            If SSTab1.TabVisible(1) Then If ValidarTab1 Then ActualizaTab1
            If SSTab1.TabVisible(2) Then If ValidarTab2 Then ActualizaTab2
            If SSTab1.TabVisible(3) Then If ValidarTab3 Then ActualizaTab3
        End If
    End If
End Sub

Private Function Modificar() As Boolean
    Dim SQL As String
    Dim CodEmp As String
    Dim rsmod As MYSQL_RS
    Modificar = False
    If SSTab1.TabVisible(0) Then CodEmp = Trim(txtEmpleado_1)
    If SSTab1.TabVisible(1) Then CodEmp = Trim(txtEmpleado_2)
    If SSTab1.TabVisible(2) Then CodEmp = Trim(txtEmpleado_3)
    If SSTab1.TabVisible(3) Then CodEmp = Trim(txtEmpleado_4)
    SQL = "Select fec_salida, fec_regreso from calendario where codigo='" & Trim(lblCodigoCalen) & "' and codemp='" & Trim(CodEmp) & "'"
    Set rsmod = oConexion.EjecutaSelectRS(SQL)
    If Not rsmod.EOF Then
        If IsDate(rsmod.Fields("fec_salida")) Then
            If CDate(rsmod.Fields("fec_salida")) >= Date Then Modificar = True
        End If
    End If
    Set rsmod = Nothing
End Function

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    ModoFormulario modNuevo
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cbodiv_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then cmbMedio.SetFocus
End Sub

Private Sub cmbAgencia_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cmbLinea.SetFocus
End Sub

Private Sub cmbEstancia_Change()
  If cmbEstancia.ListCount > 0 Then NomEstancia cmbNombreEst
End Sub

Private Sub cmbEstancia_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cmbNombreEst.SetFocus
End Sub

Private Sub cmbLinea_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then optIda.SetFocus
End Sub

Private Sub cmbLote_Change()
    If cmbLote.ListCount > 0 Then pozo cmbPozo
End Sub

Private Sub cmbLote_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cmbPozo.SetFocus
End Sub

Private Sub cmbMedio_Change()
   If cmbMedio.ListCount > 0 Then Linea cmbLinea
End Sub

Private Sub cmbMedio_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then cmbAgencia.SetFocus
End Sub

Private Sub cmbMonBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meBoleto.SetFocus
End Sub

Private Sub cmbMonEstancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then meEstancia.SetFocus
End Sub

Private Sub cmbNombreEst_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then optNoche.SetFocus
End Sub

Private Sub cmbPeriodo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        chkGoceHaber.SetFocus
    End If
End Sub

Private Sub cmbPozo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then CboDiv.SetFocus
End Sub

Private Sub cmdgentareo_Click()
    GenTareo
End Sub

Sub GenTareo()
On Error GoTo CtrlError
    Dim Mto As Double, SQL As String, AnoMes As String
    Dim RQ As MYSQL_RS, RQ1 As MYSQL_RS
    SQL = "SELECT sbasico,c.codigo,codafp,asigfam,jubilado,sctr,cafp from empleado e left join contrato c on (e.codigo=c.codemp) where " & _
          "e.codigo = '" & Trim(txtEmpleado_2) & "' and c.estado = 'AP'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If Month(dpIni_vac.Value) = Month(dpFin_vac.Value) Then valor = 1 Else valor = 2
        For I = 1 To valor
            If I = 1 Then
                AnoMes = Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2)
            Else
                AnoMes = Year(dpFin_vac.Value) & Right("00" & Month(dpFin_vac.Value), 2)
            End If
            'INGRESOS
            SQL = "Select bono from contrato where codemp='" & Trim(txtEmpleado_2) & "' and estado='AP'"
            Set RQ1 = oConexion.EjecutaSelectRS(SQL)
            If Not RQ1.EOF() Then
                If RQ1.Fields("bono") = "S" Then
                    Mto = FormatNumber(BonosCampo(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N")), 2)
                    SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                          "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                          "'" & Trim(txtEmpleado_2) & "','1001','V'," & _
                          "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                End If
            End If
            If RQ.Fields("asigfam") = "S" Then
                Mto = FormatNumber(AsigFam(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','201','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            'DSCTOS
            Mto = FormatNumber(RtaQta(AnoMes, 4, Trim(txtEmpleado_2), 3700, 8, 14, 17, 20, 30, IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("sbasico")), 2)
            SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                  "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                  "'" & Trim(txtEmpleado_2) & "','605','V'," & _
                  "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            If afp = "06" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(Onp(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("CODAFP"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','607','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            If afp <> "06" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(Afp10(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("CODAFP"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','608','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            If afp <> "06" And RQ.Fields("sctr") = "S" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(Afp2(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("CODAFP"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','611','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            If afp <> "06" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(AfpCom(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("CODAFP"), RQ.Fields("sbasico"), RQ.Fields("cafp")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','601','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            If afp <> "06" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(AfpComSeg(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("CODAFP"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','606','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            Mto = FormatNumber(PrestamosyAdelantos(Right(AnoMes, 2), Trim(txtEmpleado_2), 4, "P"), 2)
            If Mto > 0 Then
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','709','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                ActualizaMontos AnoMes, 4, IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), "P", Trim(txtEmpleado_2), 0
            End If
            Mto = FormatNumber(PrestamosyAdelantos(Right(AnoMes, 2), Trim(txtEmpleado_2), 4, "A"), 2)
            If Mto > 0 Then
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','701','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                ActualizaMontos AnoMes, 4, IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), "A", Trim(txtEmpleado_2), 0
            End If
            Mto = FormatNumber(RetencionJudicial(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N")), 2)
            If Mto > 0 Then
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','703','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            'APORTES
            Mto = FormatNumber(Essalud(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("sbasico")), 2)
            SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                  "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                  "'" & Trim(txtEmpleado_2) & "','804','V'," & _
                  "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            If RQ.Fields("sctr") = "S" Then
                Mto = FormatNumber(Sctr(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','806','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
            If RQ.Fields("codafp") <> "06" And RQ.Fields("sctr") = "S" And RQ.Fields("jubilado") = "N" Then
                Mto = FormatNumber(Afp2E(AnoMes, 4, Trim(txtEmpleado_2), IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N"), RQ.Fields("sbasico")), 2)
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                      "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                      "'" & Trim(txtEmpleado_2) & "','805','V'," & _
                      "" & Mto & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
        Next
    End If
    cmdgentareo.Enabled = False
    Set RQ = Nothing
    Set RQ1 = Nothing
Exit Sub
CtrlError:
    SQL = "delete from pl_tareo where emp = '" & Trim(txtEmpleado_2) & "' and tipo = 4 and anomes = '" & AnoMes & "'"
    oConexionMYSQL.Execute SQL
    MsgBox err.Description, vbCritical, "Error en generación de Tareo. No se grabó el Tareo de Vacaciones"
End Sub

Private Sub dpFin_vac_Change()
    If dpIni_vac <= dpFin_vac Then
        txtMeses = IIf((dpFin_vac - dpIni_vac) >= 30, Left((dpFin_vac - dpIni_vac) / 30, 1), 0)
        txtDias = ((dpFin_vac - dpIni_vac) + 1) Mod 30
    End If
End Sub

Private Sub dpFin_vac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If dpIni_vac <= dpFin_vac Then
            txtMeses.SetFocus
            txtMeses = IIf((dpFin_vac - dpIni_vac) >= 30, Left((dpFin_vac - dpIni_vac) / 30, 1), 0)
            txtDias = ((dpFin_vac - dpIni_vac) + 1) Mod 30
        End If
    End If
End Sub

Private Sub dpIni_vac_Change()
    If FlgFecha = True Then
        Dim SQL As String
        Dim RQ As MYSQL_RS
        SQL = "select fec_salida from calendario where codemp = '" & txtEmpleado_2.Text & "' and movemp = '02' and fec_regreso = '" & Format(dpFin_vac, "yyyy/mm/dd") & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            dpIni_vac.Value = Format(Trim(RQ.Fields("fec_salida")), "dd/mm/yyyy")
        End If
        Set RQ = Nothing
        FlgFecha = False
        MsgBox "No puede modificar la fecha de salida de vacaciones de este Trabajador", vbInformation, "NOVPeru"
    End If
End Sub

Private Sub dpIni_vac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dpFin_vac.SetFocus
    Else
        If VerificaPlanilla(txtEmpleado_2.Text, Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2)) Then
            FlgFecha = True
        Else
            FlgFecha = False
        End If
    End If
End Sub

Function VerificaPlanilla(CodEmp As String, AnoMes As String) As Boolean
Dim SQL As String
Dim RQ As MYSQL_RS
    VerificaPlanilla = False
    SQL = "select anomes from pl_planiproc where anomes = '" & AnoMes & "' and proceso = 1 and mon = '" & IIf(val(CodEmp) = 2 Or val(CodEmp) = 38, "E", "N") & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        VerificaPlanilla = True
    End If
Set RQ = Nothing
End Function

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    If lblModo = "Consulta" Then
        ModoFormulario modConsulta
    Else
        ModoFormulario modAccion
    End If
    RestRefrigerios cboRest
    Lote cmbLoteB
    Depart cmbDepartB
    Divisiones CboDivB
    Set oConsulta = New FrmConsultas
End Sub

Private Sub cmbLoteB_Change()
    If cmbLoteB.ListCount > 0 And cmbLoteB.Value <> "Seleccionar..." Then pozoB cmbPozoB
End Sub

Private Sub Limpiar()
    txtEmpleado_1 = Empty
    txtEmpleado_2 = Empty
    txtEmpleado_3 = Empty
    txtEmpleado_4 = Empty
    lblEmpleado_1 = Empty
    lblEmpleado_2 = Empty
    lblEmpleado_3 = Empty
    lblEmpleado_4 = Empty
    txtValorRefri = "0.00"
    txtObs = Empty
    txtAutoriz_lic = Empty
    txtAutoriz_per = Empty
    txtDias = 0
    txtMeses = 0
    txtMotivo_lic = Empty
    txtMotivo_per = Empty
    txtmotivosub = Empty
    txtciit = Empty
    cmbLote.Clear
    cmbAgencia.Clear
    cmbEstancia.Clear
    cmbMedio.Clear
    cmbLinea.Clear
    cmbPeriodo.Clear
    cmbDepart.Clear
    CboDiv.Clear
    cboRest.Clear
    CboTipo.Clear
    cmbPozo.Clear
    meHoraSalida.Text = Format(Time, "hh:mm")
    meHoraIni.Text = Format(Time, "hh:mm")
    meHoraFin.Text = Format(Time, "hh:mm")
    dptFecFinR.Value = Date
    dptFecIniR.Value = Date
    dpFin_lic.Value = Date
    dpFin_per.Value = Date
    dpFin_vac.Value = Date
    dpIni_lic.Value = Date
    dpIni_per.Value = Date
    dpIni_vac.Value = Date
    dtinisub.Value = Date
    dtfinsub.Value = Date
    dpIni_bon.Value = Date
    dpFin_bon.Value = Date
    optIda.Value = False
    optIdaVuelta.Value = False
    optEstadia.Value = False
    optNoche.Value = False
    chkGoceHaber.Value = False
    chkIndemnizacion.Value = False
End Sub

Private Sub Medio(cbo As MSForms.ComboBox)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = ""
        .AddItem "Terrestre"
        .List(1, 1) = "T"
        .AddItem "Aéreo"
        .List(2, 1) = "A"
        .ListIndex = 0
    End With
End Sub

Private Sub Estancia(cbo As MSForms.ComboBox)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = ""
        .AddItem "Hotel"
        .List(1, 1) = "H"
        .AddItem "Hostal"
        .List(2, 1) = "Hs"
        .AddItem "Pensión"
        .List(3, 1) = "P"
        .AddItem "Departamento"
        .List(4, 1) = "D"
        .AddItem "Otros"
        .List(5, 1) = "O"
        .ListIndex = 0
    End With
End Sub

Private Sub agencia(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsagencia As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo,descrip from agencia order by descrip"
    Set rsagencia = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rsagencia.EOF
            .AddItem rsagencia.Fields("descrip")
            .List(I, 1) = rsagencia.Fields("codigo")
            rsagencia.MoveNext
            I = I + 1
        Loop
        .ListIndex = 0
    End With
    Set rsagencia = Nothing
End Sub

Private Sub TipoSuspension(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsagencia As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo,descrip from pl_tipsuspensionaboral order by descrip"
    Set rsagencia = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rsagencia.EOF
            .AddItem rsagencia.Fields("descrip")
            .List(I, 1) = rsagencia.Fields("codigo")
            rsagencia.MoveNext
            I = I + 1
        Loop
        .ListIndex = 0
    End With
    Set rsagencia = Nothing
End Sub

Private Sub Depart(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsdepat As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo,descrip from departamento order by descrip"
    Set rsdepat = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rsdepat.EOF
            .AddItem rsdepat.Fields("descrip")
            .List(I, 1) = rsdepat.Fields("codigo")
            rsdepat.MoveNext
            I = I + 1
        Loop
        .ListIndex = 0
    End With
    Set rsdepat = Nothing
End Sub

Private Sub Linea(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rslinea As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo, descrip from linea where tipo = '" & cmbMedio.List(cmbMedio.ListIndex, 1) & "'"
    Set rslinea = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rslinea.EOF
            .AddItem rslinea.Fields("descrip")
            .List(I, 1) = rslinea.Fields("codigo")
            rslinea.MoveNext
            I = I + 1
        Loop
       .ListIndex = 0
    End With
    Set rslinea = Nothing
End Sub

Private Sub NomEstancia(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsest As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo, descrip from estancia where tipo = '" & cmbEstancia.List(cmbEstancia.ListIndex, 1) & "'"
    Set rsest = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rsest.EOF
            .AddItem rsest.Fields("descrip")
            .List(I, 1) = rsest.Fields("codigo")
            rsest.MoveNext
            I = I + 1
        Loop
        .ListIndex = 0
    End With
    Set rsest = Nothing
End Sub

Private Sub Lote(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rslote As MYSQL_RS
    Dim I As Integer
    SQL = "cen_lt group by idlote"
    
    Set rslote = oConexion.EjecutaSelect(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = ""
    I = 1
    Do While Not rslote.EOF
        With cbo
            .AddItem rslote.Fields("descrip")
            .List(I, 1) = rslote.Fields("descrip")
            .List(I, 2) = rslote.Fields("idlote")
            I = I + 1
            rslote.MoveNext
        End With
    Loop
    cbo.ListIndex = 0
    Set rslote = Nothing
End Sub

Private Sub pozo(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rspozo As MYSQL_RS
    Dim I As Integer
    SQL = "cen_pz where idlote in (select idlote from novperuvhse.lote where descripcioncorta ='" & cmbLote.List(cmbLote.ListIndex, 0) & "') group by idpozo"
    Set rspozo = oConexion.EjecutaSelect(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = ""
    I = 1
    Do While Not rspozo.EOF
        With cbo
            .AddItem rspozo.Fields("descrip")
            .List(I, 1) = rspozo.Fields("descrip")
            .List(I, 2) = rspozo.Fields("idpozo")
            I = I + 1
            rspozo.MoveNext
        End With
    Loop
    cbo.ListIndex = 0
    Set rspozo = Nothing
End Sub

Private Sub pozoB(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rspozo As MYSQL_RS
    Dim I As Integer
    SQL = "cen_pz where idlote in (select idlote from novperuvhse.lote where descripcioncorta ='" & cmbLoteB.List(cmbLoteB.ListIndex, 0) & "') group by idpozo"
    Set rspozo = oConexion.EjecutaSelect(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = ""
    I = 1
    Do While Not rspozo.EOF
        With cbo
            .AddItem rspozo.Fields("descrip")
            .List(I, 1) = rspozo.Fields("descrip")
            .List(I, 2) = rspozo.Fields("idpozo")
            I = I + 1
            rspozo.MoveNext
        End With
    Loop
    cbo.ListIndex = 0
    Set rspozo = Nothing
End Sub


Private Sub RestRefrigerios(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim rsrestrefri As MYSQL_RS
    Dim I As Integer
    SQL = "Select codigo, descrip from cnauxil where  auxiliar='5' and suspen = 'REFRI'"
    Set rsrestrefri = oConexion.EjecutaSelectRS(SQL)
    With cbo
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "00"
        I = 1
        Do While Not rsrestrefri.EOF
            .AddItem rsrestrefri.Fields("descrip")
            .List(I, 1) = rsrestrefri.Fields("codigo")
            rsrestrefri.MoveNext
            I = I + 1
        Loop
        .ListIndex = 0
    End With
    Set rsrestrefri = Nothing
End Sub

Private Function GeneraCod(emp As String, mov As String) As String
    Dim SQL As String
    Dim rscod As MYSQL_RS
    SQL = " Select Right(max(codigo),2) as codigo from calendario where left(codigo,6)='" & strAnoSistema & strMesSistema & "' and codemp = '" & emp & "' and movemp='" & mov & "'"
    Set rscod = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(rscod.Fields("Codigo")) Then
        GeneraCod = strAnoSistema & strMesSistema & Right("00" & Trim((CDbl(rscod.Fields("Codigo")) + 1)), 2)
    Else
        GeneraCod = strAnoSistema & strMesSistema & "01"
    End If
    Set rscod = Nothing
End Function

Public Function MovEmp(Tipo As String) As String
    Dim SQL As String
    Dim rsmov As MYSQL_RS
    MovEmp = "00"
    SQL = "Select codigo from movi_emp where descrip = '" & UCase(Trim(Tipo)) & "'"
    Set rsmov = oConexion.EjecutaSelectRS(SQL)
    If Not rsmov.EOF Then
        MovEmp = rsmov.Fields("codigo")
    End If
    Set rsmov = Nothing
End Function

Private Function ValidarTab1() As Boolean
    ValidarTab1 = True
    Select Case UCase(frmPrograma.tag)
        Case "VACACIONES"
            If dpIni_vac > dpFin_vac Then
                MsgBox "La Fecha de Inicio de las Vacaciones debe ser menor a la de regreso", vbInformation, gsNomSW
                ValidarTab1 = False
                dpIni_vac.SetFocus
            End If
            If cmbPeriodo.ListCount > 0 And cmbPeriodo.ListIndex < 0 Then
                MsgBox "Selecione el periodo correspondiente a las vacaciones", vbInformation, gsNomSW
                ValidarTab1 = False
                cmbPeriodo.SetFocus
            End If
            If chkGoceHaber.Value = True Then
                If ValidadiasCompra(val(txtDias.Text)) Then
                    MsgBox "No puede registrar Compra de Vacaciones de más de 15 dias", vbInformation, gsNomSW
                    ValidarTab1 = False
                Else
                    If val(txtDias.Text) > 15 Then
                        MsgBox "No puede registrar Compra de Vacaciones de más de 15 dias", vbInformation, gsNomSW
                        ValidarTab1 = False
                    End If
                End If
            End If
            If ValidaTotalDiasVac(val(txtDias.Text)) Then
                MsgBox "No puede registrar Vacaciones de más de 30 días", vbInformation, gsNomSW
                ValidarTab1 = False
            Else
                If val(txtDias.Text) > 30 Then
                    MsgBox "No puede registrar Vacaciones de más de 30 dias", vbInformation, gsNomSW
                    ValidarTab1 = False
                End If
            End If
        Case "PERMISOS"
            If dpIni_per > dpFin_per Then
                MsgBox "La Fecha de Inicio del Permiso debe ser menor a la de regreso", vbInformation, gsNomSW
                ValidarTab1 = False
                dpIni_vac.SetFocus
            End If
        Case "LICENCIA"
            If dpIni_lic > dpFin_lic Then
                MsgBox "La Fecha de Inicio de la Licencia debe ser menor a la de regreso", vbInformation, gsNomSW
                ValidarTab1 = False
                dpIni_vac.SetFocus
            End If
        Case "SUBSIDIOS"
            If dtinisub > dtfinsub Then
                MsgBox "La Fecha de Inicio del Subsidio debe ser menor a la de regreso", vbInformation, gsNomSW
                ValidarTab1 = False
                dtinisub.SetFocus
            End If
    End Select
End Function

Private Function ValidarTab3() As Boolean
    ValidarTab3 = True
        If dpIni_bon > dpFin_bon Then
            MsgBox "La Fecha de Inicio del Bono debe ser menor a la de regreso", vbInformation, gsNomSW
            ValidarTab3 = False
            dpIni_bon.SetFocus
        End If
End Function

Function ValidadiasCompra(dias As Integer) As Boolean
    Dim SQL As String
    Dim RQ As MYSQL_RS
    ValidadiasCompra = False
    SQL = "select IFNULL(sum(monto_viatico),0) as tot from calendario where codemp = '" & Trim(txtEmpleado_2) & "' " & _
          "and gocehaber = 'S' and periodo = '" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If val(RQ.Fields("tot")) + dias > 15 Then
            ValidadiasCompra = True
        End If
    End If
    Set RQ = Nothing
End Function

Function ValidaTotalDiasVac(dias As Integer) As Boolean
    Dim SQL As String
    Dim RQ As MYSQL_RS
    ValidaTotalDiasVac = False
    SQL = "select IFNULL(sum(monto_viatico),0) as tot from calendario where codemp = '" & Trim(txtEmpleado_2) & "' " & _
          "and periodo = '" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If val(RQ.Fields("tot")) + dias > 30 Then
            ValidaTotalDiasVac = True
        End If
    End If
    Set RQ = Nothing
End Function

Private Function ValidarTab2() As Boolean
    ValidarTab2 = True
    If dptFecIniR > dptFecFinR Then
        MsgBox "La Fecha de Inicio de refrigerio debe ser menor a la de último dia", vbInformation, gsNomSW
        ValidarTab2 = False
        dptFecIniR.SetFocus
    End If
    If chkRefri2(0).Value = False And chkRefri2(1).Value = False Then
        MsgBox "Especifique el tipo de refrigerio a registrar", vbInformation, gsNomSW
        ValidarTab2 = False
    End If
End Function

Private Sub GrabarTab1()
    Dim SQL As String, SQL1 As String, valor As Integer
    Dim periodovaciones As String, acantidad As Double
    Dim DiasFaltantes As String
    Dim AnoMes As String, RQ As MYSQL_RS
    lblCodigoCalen = GeneraCod(Trim(txtEmpleado_2), MovEmp(frmPrograma.tag))
    Select Case UCase(frmPrograma.tag)
        Case "VACACIONES"
            DiasFaltantes = VerificaDiasFaltantes
            If DiasFaltantes <> "" Then
                If MsgBox("Existen Períodos Anteriores con días faltantes de Vacaciones: " & Chr(13) & DiasFaltantes & Chr(13) & _
                          "¿Desea Continuar con el registro?", vbQuestion + vbYesNo, gsNomSW) = vbNo Then
                    Exit Sub
                End If
            End If
            periodovaciones = cmbPeriodo.List(cmbPeriodo.ListIndex)
            SQL = " Insert into calendario(dpto,pozo,codagencia,codlinea,tipoboleto,mon_boleto,pagoestancia,mon_estancia,mon_viatico,monto_viatico,observacion,sinbono,codigo,codemp,movemp,fec_salida,fec_regreso,periodo,gocehaber)" & _
                  " values('','','','','','','','',''," & Trim(txtDias) & ",'','','" & lblCodigoCalen & "','" & Trim(txtEmpleado_2) & "' ,'" & MovEmp(frmPrograma.tag) & "','" & Format(dpIni_vac.Value, "yyyy/mm/dd") & "'," & _
                  " '" & Format(dpFin_vac.Value, "yyyy/mm/dd") & "','" & periodovaciones & "'," & _
                  " '" & IIf(chkGoceHaber.Value, "S", IIf(chkIndemnizacion.Value, "I", "N")) & "');"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
            SQL = "SELECT sbasico,c.codigo,codafp from empleado e left join contrato c on (e.codigo=c.codemp) where " & _
                  "e.codigo = '" & Trim(txtEmpleado_2) & "' and c.estado = 'AP'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If Month(dpIni_vac.Value) = Month(dpFin_vac.Value) Then valor = 1 Else valor = 2
                For I = 1 To valor
                    If I = 1 Then
                        AnoMes = Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2)
                    Else
                        AnoMes = Year(dpFin_vac.Value) & Right("00" & Month(dpFin_vac.Value), 2)
                    End If
                    If chkGoceHaber.Value = False Then
                        acantidad = CDbl(diasvaca(Trim(txtEmpleado_2), AnoMes, 4))
                    Else
                        acantidad = CDbl(diasvacaCompra(Trim(txtEmpleado_2), AnoMes, 4))
                    End If
                    SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                          "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                          "'" & Trim(txtEmpleado_2) & "','" & IIf(chkGoceHaber.Value, "117", "121") & "','D'," & _
                          "" & acantidad & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                Next
            End If
        Case "PERMISOS"
            SQL = " Insert into calendario(dpto,lote,pozo,codagencia,codlinea,tipoboleto,mon_boleto,monto_boleto,codestancia,pagoestancia,mon_estancia,mon_viatico,sinbono,periodo,gocehaber,codigo,codemp,movemp,fec_salida,hora_salida,fec_regreso,hora_regreso,autorizado,observacion)" & _
                  " values('','','','','','','',0,'','','','','','','','" & lblCodigoCalen & "', '" & Trim(txtEmpleado_2) & "','" & MovEmp(frmPrograma.tag) & "', '" & Format(dpIni_per.Value, "yyyy/mm/dd") & "', '" & meHoraIni.Text & "'," & _
                  " '" & Format(dpFin_per.Value, "yyyy/mm/dd") & "', '" & meHoraFin.Text & "', '" & Trim(lblcod) & "', '" & Trim(txtMotivo_per) & "');"
        Case "LICENCIA"
            SQL = " Insert into calendario(codigo,codemp,movemp,fec_salida,fec_regreso,autorizado,observacion,sinbono)" & _
                  " values('" & lblCodigoCalen & "', '" & Trim(txtEmpleado_2) & "','" & MovEmp(frmPrograma.tag) & "', '" & Format(dpIni_lic.Value, "yyyy/mm/dd") & "'," & _
                  " '" & Format(dpFin_lic.Value, "yyyy/mm/dd") & "','" & Trim(txtAutoriz_per) & "', '" & Trim(txtMotivo_lic) & "','');"
        Case "SUBSIDIOS"
            SQL = " Insert into calendario(codigo,codemp,movemp,fec_salida,fec_regreso,dpto,observacion,sinbono)" & _
                  " values('" & lblCodigoCalen & "', '" & Trim(txtEmpleado_2) & "','" & MovEmp(frmPrograma.tag) & "', '" & Format(dtinisub.Value, "yyyy/mm/dd") & "'," & _
                  " '" & Format(dtfinsub.Value, "yyyy/mm/dd") & "','" & Trim(CboTipo.List(CboTipo.ListIndex, 1)) & "', '" & Trim(txtmotivosub) & "','" & Trim(txtciit) & "');"
    End Select
    If UCase(frmPrograma.tag) <> "VACACIONES" Then
        If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, True) Then
            MsgBox "Se produjo un error al momento de grabar", vbInformation, gsNomSW
        Else
            ModoFormulario modConsulta
        End If
    Else
        ModoFormulario modConsulta
        cmdgentareo.Enabled = True
    End If
End Sub


Private Sub GrabarTab3()
    Dim SQL As String, SQL1 As String, valor As Integer
    Dim periodovaciones As String, acantidad As Double
    Dim DiasFaltantes As String
    Dim AnoMes As String, RQ As MYSQL_RS
    
    lblCodigoCalen = GeneraCod(Trim(txtEmpleado_4), "01")
    SQL = " Insert into calendario(codigo,codemp,movemp,fec_salida,fec_regreso,dpto,lote,pozo,autorizado,observacion,sinbono,periodo)" & _
                  " values('" & lblCodigoCalen & "', '" & Trim(txtEmpleado_4) & "','01', ''," & _
                  " '" & Format(dpIni_bon.Value, "yyyy/mm/dd") & "','" & cmbDepartB.List(cmbDepartB.ListIndex, 1) & "', '" & cmbLoteB.List(cmbLoteB.ListIndex, 2) & "', '" & cmbPozoB.List(cmbPozoB.ListIndex, 2) & "','', '" & Trim(txtMotivo_bon) & "','14292452', '" & CboDivB.List(CboDivB.ListIndex, 1) & "');"
                  
    If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, True) Then
         MsgBox "Se produjo un error al momento de grabar Periodo 1", vbInformation, gsNomSW
    End If
    
    
    lblCodigoCalen = GeneraCod(Trim(txtEmpleado_4), "01")
    SQL = " Insert into calendario(codigo,codemp,movemp,fec_salida,fec_regreso,dpto,lote,pozo,autorizado,observacion,sinbono,periodo)" & _
                  " values('" & lblCodigoCalen & "', '" & Trim(txtEmpleado_4) & "','01', '" & Format(dpFin_bon.Value, "yyyy/mm/dd") & "'," & _
                  " '','" & cmbDepartB.List(cmbDepartB.ListIndex, 1) & "', '" & cmbLoteB.List(cmbLoteB.ListIndex, 2) & "', '" & cmbPozoB.List(cmbPozoB.ListIndex, 2) & "','', '" & Trim(txtMotivo_bon) & "','14292452', '" & CboDivB.List(CboDivB.ListIndex, 1) & "');"
                  
    If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, True) Then
            MsgBox "Se produjo un error al momento de grabar Periodo 2", vbInformation, gsNomSW
    Else
            ModoFormulario modConsulta
    End If
    
End Sub


Function VerificaDiasFaltantes() As String
    Dim SQL As String, cad As String
    Dim RQ As MYSQL_RS
    
    cad = ""
    SQL = "select sum(monto_viatico) as tot,periodo from calendario c where codemp = '" & Trim(txtEmpleado_2.Text) & "' " & _
          "AND periodo <> '" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "' group by periodo order by periodo,fec_salida"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Do While Not RQ.EOF()
        If RQ.Fields("tot") < 30 Then
            cad = cad & "Período " & Trim(RQ.Fields("periodo")) & ": " & 30 - val(RQ.Fields("tot")) & Chr(13)
        End If
        RQ.MoveNext
    Loop
    VerificaDiasFaltantes = cad
    Set RQ = Nothing
End Function

Private Sub GrabarTab2()
On Error GoTo ErrorGraba
    Dim SQL As String
    lblCodigoCalen = GeneraCod(Trim(txtEmpleado_3), "06")
    If chkRefri2(0).Value = True Then
        SQL = " Insert into calendario (codigo,codemp,movemp,fec_Salida,fec_regreso,sinbono,observacion,mon_estancia,monto_estancia,pozo)" & _
              " values ('" & lblCodigoCalen & "'" & _
              ",'" & Trim(txtEmpleado_3) & "','06','" & Format(dptFecIniR.Value, "yyyy/mm/dd") & "'" & _
              ",'" & Format(dptFecFinR.Value, "yyyy/mm/dd") & "','" & Trim(txtNroFac) & "','ALMUERZO','N'," & txtValorRefri & ",'" & cboRest.List(cboRest.ListIndex, 1) & "')"
        oConexionMYSQL.Execute SQL
    End If
    lblCodigoCalen = GeneraCod(Trim(txtEmpleado_3), "06")
    If chkRefri2(1).Value = True Then
        SQL = " Insert into calendario (codigo,codemp,movemp,fec_Salida,fec_regreso,sinbono,observacion,mon_estancia,monto_estancia,pozo)" & _
              " values ('" & lblCodigoCalen & "'" & _
              ",'" & Trim(txtEmpleado_3) & "','06','" & Format(dptFecIniR.Value, "yyyy/mm/dd") & "'" & _
              ",'" & Format(dptFecFinR.Value, "yyyy/mm/dd") & "','" & Trim(txtNroFac) & "','CENA','N'," & txtValorRefri & ",'" & cboRest.List(cboRest.ListIndex, 1) & "')"
        oConexionMYSQL.Execute SQL
    End If
    ModoFormulario modConsulta
Exit Sub
ErrorGraba:
    MsgBox "Se produjo un error al momento de grabar", vbInformation, gsNomSW
End Sub

Private Sub ActualizaTab1()
    Dim SQL As String, valor As Integer, acantidad As Double, I As Integer
    Dim AnoMes As String, RQ As MYSQL_RS
    Dim FlgPla As Boolean, FlgPla1 As Boolean
    Select Case frmPrograma.tag
        Case "VACACIONES"
            AnoMes = Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2)
            SQL = "SELECT * FROM PL_TAREO WHERE EMP = '" & Trim(txtEmpleado_2) & "' and tipo = 4 and anomes = '" & AnoMes & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                FlgPla = VerificaPlanilla(Trim(txtEmpleado_2.Text), Year(dpFin_vac.Value) & Right("00" & Month(dpFin_vac.Value), 2))
                FlgPla1 = VerificaPlanilla(Trim(txtEmpleado_2.Text), Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2))
                If FlgPla Then
                    MsgBox "No puede modificar las vacaciones para este Trabajador", vbInformation, "NOVPeru"
                    Set RQ = Nothing
                    Exit Sub
                Else
                    If (Month(dpIni_vac.Value) + 1 = Month(dpFin_vac.Value)) And (FlgPla1 = True) Then
                        SQL = " Update calendario set monto_viatico=" & Trim(txtDias) & ", fec_salida='" & Format(dpIni_vac.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_vac.Value, "yyyy/mm/dd") & "'," & _
                              " periodo='" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "', gocehaber= '" & IIf(chkGoceHaber.Value, "S", IIf(chkIndemnizacion.Value, "I", "N")) & "'" & _
                              " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
                        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
                        valor = 3
                        AnoMes = Year(dpFin_vac.Value) & Right("00" & Month(dpFin_vac.Value), 2)
                        GoTo AKI
                    Else
                        If MsgBox("¿Desea modificar el Tareo de Vacaciones para el período " & AnoMes & " de este Trabajador?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                            SQL = " Update calendario set monto_viatico=" & Trim(txtDias) & ", fec_salida='" & Format(dpIni_vac.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_vac.Value, "yyyy/mm/dd") & "'," & _
                                  " periodo='" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "', gocehaber= '" & IIf(chkGoceHaber.Value, "S", IIf(chkIndemnizacion.Value, "I", "N")) & "'" & _
                                  " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
                            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
                            SQL = "DELETE FROM PL_TAREO WHERE EMP = '" & Trim(txtEmpleado_2) & "' and tipo = 4 and anomes = '" & AnoMes & "'"
                            oConexionMYSQL.Execute SQL
                            SQL = "SELECT sbasico,c.codigo,codafp from empleado e left join contrato c on (e.codigo=c.codemp) where " & _
                                  "e.codigo = '" & Trim(txtEmpleado_2) & "' and c.estado = 'AP'"
                            Set RQ = oConexion.EjecutaSelectRS(SQL)
                            If Not RQ.EOF() Then
                                If Month(dpIni_vac.Value) = Month(dpFin_vac.Value) Then valor = 1 Else valor = 2
                                For I = 1 To valor
                                    If I = 1 Then
                                        AnoMes = Year(dpIni_vac.Value) & Right("00" & Month(dpIni_vac.Value), 2)
                                    Else
                                        AnoMes = Year(dpFin_vac.Value) & Right("00" & Month(dpFin_vac.Value), 2)
                                    End If
AKI:
                                    If chkGoceHaber.Value = False Then
                                        acantidad = CDbl(diasvaca(Trim(txtEmpleado_2), AnoMes, 4))
                                    Else
                                        acantidad = CDbl(diasvacaCompra(Trim(txtEmpleado_2), AnoMes, 4))
                                    End If
                                    SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                                          "'" & AnoMes & "','4','" & IIf(val(Trim(txtEmpleado_2)) = 4 Or val(Trim(txtEmpleado_2)) = 38, "E", "N") & "', " & _
                                          "'" & Trim(txtEmpleado_2) & "','" & IIf(chkGoceHaber.Value, "117", "121") & "','D'," & _
                                          "" & acantidad & ",'" & Format(Date, "yyyy/mm/dd") & "'," & RQ.Fields("sbasico") & ",'" & RQ.Fields("codafp") & "','" & RQ.Fields("codigo") & "')"
                                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                                    If valor = 3 Then
                                        ModoFormulario modConsulta
                                        Exit Sub
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            Else
                SQL = " Update calendario set monto_viatico=" & Trim(txtDias) & ", fec_salida='" & Format(dpIni_vac.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_vac.Value, "yyyy/mm/dd") & "'," & _
                      " periodo='" & cmbPeriodo.List(cmbPeriodo.ListIndex) & "', gocehaber= '" & IIf(chkGoceHaber.Value, "S", IIf(chkIndemnizacion.Value, "I", "N")) & "'" & _
                      " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
            End If
            Set RQ = Nothing
        Case "PERMISOS"
            SQL = " Update calendario set fec_salida='" & Format(dpIni_per.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_per.Value, "yyyy/mm/dd") & "'," & _
                  " hora_salida = '" & meHoraIni.Text & "', hora_regreso='" & meHoraFin.Text & "'," & _
                  " autorizado= '" & lblcod & "', observacion = '" & Trim(txtMotivo_per) & "' " & _
                  " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
        Case "LICENCIA"
            SQL = " Update calendario set fec_salida='" & Format(dpIni_lic.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_lic.Value, "yyyy/mm/dd") & "'," & _
                  " autorizado= '" & lblcod & "', observacion = '" & Trim(txtMotivo_lic) & "' " & _
                  " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
        Case "SUBSIDIOS"
            SQL = " Update calendario set fec_salida='" & Format(dtinisub.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dtfinsub.Value, "yyyy/mm/dd") & "'," & _
                  " dpto= '" & Trim(CboTipo.List(CboTipo.ListIndex, 1)) & "', observacion = '" & Trim(txtmotivosub) & "',sinbono='" & Trim(txtciit) & "' " & _
                  " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_2) & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
    End Select
    If frmPrograma.tag <> "VACACIONES" Then
        If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, True) Then
            MsgBox "No se actualizaron los datos correctamente", vbInformation, gsNomSW
        Else
            ModoFormulario modConsulta
        End If
    Else
        ModoFormulario modConsulta
    End If
End Sub

Private Sub ActualizaTab2()
    Dim SQL As String
    SQL = " Update calendario set fec_salida='" & Format(dptFecIniR.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dptFecFinR.Value, "yyyy/mm/dd") & "'," & _
               " monto_viatico = " & CDbl(txtValorRefri) & _
               " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_3) & " and movemp='06'"
    If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, True) Then
        MsgBox "No se actualizaron los datos correctamente", vbInformation, gsNomSW
    Else
        ModoFormulario modConsulta
    End If
End Sub

Private Sub ActualizaTab3()
    Dim SQL As String, valor As Integer, acantidad As Double, I As Integer
    Dim AnoMes As String, RQ As MYSQL_RS
    Dim FlgPla As Boolean, FlgPla1 As Boolean
    SQL = " Update calendario set fec_salida='" & Format(dpIni_bon.Value, "yyyy/mm/dd") & "', fec_regreso= '" & Format(dpFin_bon.Value, "yyyy/mm/dd") & "'," & _
                  " observacion = '" & Trim(txtMotivo_bon) & "' " & _
                  " where codigo='" & lblCodigoCalen & "' and codemp='" & Trim(txtEmpleado_4) & "' and movemp='01'"
                  
End Sub


Public Sub CargaTab1(CodCalen As String, CodEmp As String)
    Dim SQL As String
    Dim rstab As MYSQL_RS
    Dim I As Integer
    SQL = " Select * from calendario where codigo='" & CodCalen & "' and codemp='" & CodEmp & "' and movemp='" & MovEmp(frmPrograma.tag) & "'"
    Set rstab = oConexion.EjecutaSelectRS(SQL)
    If Not rstab.EOF Then
        With rstab
            txtEmpleado_2 = .Fields("codemp")
            lblEmpleado_2 = DescripcionesdeCodigos("EMPLEADO", .Fields("codemp"))
            lblCodigoCalen = .Fields("codigo")
            Select Case frmPrograma.tag
                Case "VACACIONES"
                    dpIni_vac.Value = CDate(.Fields("fec_salida"))
                    dpFin_vac.Value = CDate(.Fields("fec_regreso"))
                    
                    If .Fields("gocehaber") = "S" Then
                        chkGoceHaber.Value = True
                    Else
                        chkGoceHaber.Value = False
                    End If
                    
                    If .Fields("gocehaber") = "I" Then
                        chkIndemnizacion.Value = True
                    Else
                        chkIndemnizacion.Value = False
                    End If
                    
                    txtMeses = IIf(.Fields("monto_viatico") >= 30, Trim(.Fields("monto_viatico") / 30), 0)
                    txtDias = .Fields("monto_viatico") Mod 30
                    Periodo cmbPeriodo
                Case "PERMISOS"
                    dpIni_per = CDate(.Fields("fec_salida"))
                    dpFin_per = CDate(.Fields("fec_regreso"))
                    meHoraIni = .Fields("hora_salida")
                    meHoraFin = .Fields("hora_regreso")
                    lblcod = .Fields("autorizado")
                    txtAutoriz_per = DescripcionesdeCodigos("EMPLEADO", lblcod)
                    txtMotivo_per = .Fields("observacion")
                Case "LICENCIA"
                    dpIni_lic = CDate(.Fields("fec_salida"))
                    dpFin_lic = CDate(.Fields("fec_regreso"))
                    lblcod = .Fields("autorizado")
                    txtAutoriz_lic = DescripcionesdeCodigos("EMPLEADO", lblcod)
                    txtMotivo_lic = .Fields("observacion")
                Case "SUBSIDIOS"
                    dtinisub = CDate(.Fields("fec_salida"))
                    dtfinsub = CDate(.Fields("fec_regreso"))
                    TipoSuspension CboTipo
                    For I = 0 To CboTipo.ListCount - 1
                        If CE(.Fields("dpto")) = Trim(CboTipo.List(I, 1)) Then
                            CboTipo.ListIndex = I
                            Exit For
                        Else
                            CboTipo.ListIndex = 0
                        End If
                    Next
                    txtciit = .Fields("sinbono")
                    txtmotivosub = .Fields("observacion")
            End Select
        End With
    End If
    Set rstab = Nothing
End Sub
'Caso cuando interactua con Calendario
Private Sub CargaTab2(CodCalen As String, CodEmp As String)
    Dim SQL As String
    Dim rstab As MYSQL_RS
    SQL = " Select * from calendario where codigo='" & CodCalen & "' and codemp='" & CodEmp & "' and movemp='06'"
    Set rstab = oConexion.EjecutaSelectRS(SQL)
    If Not rstab.EOF Then
        With rstab
            txtEmpleado_3 = .Fields("codemp")
            lblEmpleado_3 = DescripcionesdeCodigos("EMPLEADO", .Fields("codemp"))
            lblCodigoCalen = .Fields("codigo")
            dptFecIniR.Value = CDate(.Fields("fec_salida"))
            dptFecFinR.Value = CDate(.Fields("fec_regreso"))
        End With
    End If
    Set rstab = Nothing
End Sub

Private Sub GrabarTab0()
    Dim SQL As String
    Dim TipoBoleto As String
    Dim pagoestancia As String
    Dim sinbono As String, aux1 As String, aux2 As String
    Dim I As Integer
    sinbono = ""
    If optIda Then TipoBoleto = "I"
    If optEstadia Then TipoBoleto = "I/V"
    If optNoche Then pagoestancia = "Noche"
    If optEstadia Then pagoestancia = "Estadia"
    If meBoleto = "" Then meBoleto = "0.00"
    If meEstancia = "" Then meEstancia = "0.00"
    lblCodigoCalen = GeneraCod(txtEmpleado_1, MovEmp(frmPrograma.tag))
    SQL = " Insert into calendario(codigo,codemp,movemp,fec_salida,hora_salida,fec_regreso,hora_regreso,dpto,lote,pozo," & _
          " codagencia,codlinea,tipoboleto,mon_boleto,monto_boleto,codestancia,pagoestancia,mon_estancia," & _
          " monto_estancia,mon_viatico,monto_viatico,observacion,sinbono,periodo,gocehaber,autorizado)" & _
          " values('" & lblCodigoCalen & "','" & Right("00000000000" & Trim(txtEmpleado_1), 11) & "', '" & MovEmp(frmPrograma.tag) & "'," & _
          " '" & IIf(lbltipo.Caption = "I", "", Format(dpsalida.Value, "yyyy/mm/dd")) & "', '" & meHoraSalida.Text & "', " & _
          "'" & IIf(lbltipo.Caption = "I", Format(dpsalida.Value, "yyyy/mm/dd"), "") & "'," & _
          " '','" & cmbDepart.List(cmbDepart.ListIndex, 1) & "', '" & cmbLote.List(cmbLote.ListIndex, 2) & "', '" & cmbPozo.List(cmbPozo.ListIndex, 2) & "', " & _
          " '" & cmbAgencia.List(cmbAgencia.ListIndex, 1) & "', '" & cmbLinea.List(cmbLinea.ListIndex, 1) & "', '" & TipoBoleto & "'," & _
          " '" & cmbMonBoleto.List(cmbMonBoleto.ListIndex) & "'," & CDbl(meBoleto.Text) & ",'" & cmbNombreEst.List(cmbNombreEst.ListIndex, 1) & "'," & _
          " '" & pagoestancia & "','" & cmbMonEstancia.List(cmbMonEstancia.ListIndex) & "'," & CDbl(meEstancia.Text) & "," & _
          " '',0, '" & Trim(txtObs) & "', " & _
          " '16777215','" & CboDiv.List(CboDiv.ListIndex, 1) & "','','');"
    If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.insertar, False) Then
        MsgBox "Se produjo un error al momento de grabar", vbInformation, gsNomSW
    Else
        ModoFormulario modConsulta
    End If
End Sub

Private Sub ActualizaTab0()
    Dim SQL As String
    Dim TipoBoleto As String
    Dim pagoestancia As String
    Dim sinbono As String, aux1 As String, aux2 As String
    Dim I As Integer, StrCol As String
    sinbono = ""
    If optIda Then TipoBoleto = "I"
    If optEstadia Then TipoBoleto = "I/V"
    If optNoche Then pagoestancia = "Noche"
    If optEstadia Then pagoestancia = "Estadia"
    If meBoleto = "" Then meBoleto = "0.00"
    If meEstancia = "" Then meEstancia = "0.00"
    StrCol = IIf(lblcolor.Caption <> "16777215", "", 16777215)
    If StrCol <> "" Then
        StrCol = ",sinbono='16777215'"
    End If
    SQL = " Update calendario set fec_salida = '" & IIf(lbltipo.Caption = "I", "", Format(dpsalida.Value, "yyyy/mm/dd")) & "', " & _
          " fec_regreso='" & IIf(lbltipo.Caption = "I", Format(dpsalida.Value, "yyyy/mm/dd"), "") & "'," & _
          " hora_salida='" & meHoraSalida.Text & "', dpto= '" & cmbDepart.List(cmbDepart.ListIndex, 1) & "', lote='" & cmbLote.List(cmbLote.ListIndex, 2) & "'," & _
          " pozo='" & cmbPozo.List(cmbPozo.ListIndex, 2) & "', codagencia = '" & cmbAgencia.List(cmbAgencia.ListIndex, 1) & "'," & _
          " codlinea= '" & cmbLinea.List(cmbLinea.ListIndex, 1) & "', tipoboleto='" & TipoBoleto & "'," & _
          " mon_boleto='" & cmbMonBoleto.List(cmbMonBoleto.ListIndex) & "', monto_boleto=" & CDbl(meBoleto) & "," & _
          " codestancia='" & cmbNombreEst.List(cmbNombreEst.ListIndex, 1) & "', pagoestancia = '" & pagoestancia & "'," & _
          " mon_estancia='" & cmbMonEstancia.List(cmbMonEstancia.ListIndex) & "', monto_estancia=" & CDbl(meEstancia) & "," & _
          " mon_viatico=''" & StrCol & ",periodo='" & CboDiv.List(CboDiv.ListIndex, 1) & "', " & _
          " observacion= '" & Trim(txtObs) & "' where codigo='" & Trim(lblCodigoCalen) & "'" & _
          " and codemp ='" & Trim(txtEmpleado_1) & "' and movemp = '01'"
    If Not oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False) Then
        MsgBox "No se actualizaron los datos correctamente", vbInformation, gsNomSW
    Else
        Unload Me
        ModoFormulario modConsulta
    End If
End Sub

Public Sub CargaTab0(CodCalen As String, CodEmp As String)
    Dim SQL As String
    Dim rstab As MYSQL_RS
    Dim I As Integer
    Limpiar
    SQL = " Select c.codigo,c.codemp,c.fec_salida,c.hora_salida,c.fec_regreso,c.dpto,c.lote,c.pozo," & _
          " c.codagencia,c.codlinea,l.tipo as tipo_linea,c.tipoboleto,c.mon_boleto,c.monto_boleto,c.codestancia,c.pagoestancia," & _
          " e.tipo as tipo_estancia,c.mon_estancia,c.monto_estancia,c.mon_viatico,c.monto_viatico,c.observacion,c.sinbono, t.bono,c.periodo " & _
          " from calendario as c left join linea as l on (l.codigo=c.codlinea)" & _
          " left join estancia as e on(e.codigo=c.codestancia)" & _
          " left join contrato as t on (t.codemp = c.codemp) " & _
          " where c.codigo='" & CodCalen & "' and c.codemp='" & CodEmp & "'"
    Set rstab = oConexion.EjecutaSelectRS(SQL)
    If Not rstab.EOF Then
        Estancia cmbEstancia
        agencia cmbAgencia
        Lote cmbLote
        Depart cmbDepart
        Medio cmbMedio
        Divisiones CboDiv
        With rstab
            lblCodigoCalen = .Fields("codigo")
            txtEmpleado_1 = .Fields("codemp")
            lblEmpleado_1 = DescripcionesdeCodigos("EMPLEADO", Trim(.Fields("codemp")))
            For I = 0 To cmbLote.ListCount - 1
            If cmbLote.List(I, 2) = .Fields("lote") Then
               cmbLote.ListIndex = I
            End If
            Next
            cmbPozo.ListIndex = CargaCboPozo(cmbPozo, .Fields("pozo"))
            For I = 0 To cmbMedio.ListCount - 1
                If cmbMedio.List(I, 1) = .Fields("tipo_linea") Then
                    cmbMedio.ListIndex = I
                End If
            Next
            cmbLinea.ListIndex = CargaCbo(cmbLinea, .Fields("codlinea"))
            cmbAgencia.ListIndex = CargaCbo(cmbAgencia, .Fields("codagencia"))
            cmbDepart.ListIndex = CargaCbo(cmbDepart, .Fields("dpto"))
            CboDiv.ListIndex = CargaCbo(CboDiv, Trim(.Fields("periodo")))
            For I = 0 To cmbEstancia.ListCount - 1
                If cmbEstancia.List(I, 1) = .Fields("tipo_estancia") Then
                    cmbEstancia.ListIndex = I
                End If
            Next
            cmbNombreEst.ListIndex = CargaCbo(cmbNombreEst, .Fields("codestancia"))
            cmbMonBoleto.ListIndex = CargaCboMon(.Fields("mon_boleto"))
            cmbMonEstancia.ListIndex = CargaCboMon(.Fields("mon_estancia"))
            If .Fields("pagoestancia") = "Noche" Then optNoche.Value = True
            If .Fields("pagoestancia") = "Estadia" Then optEstadia.Value = True
            If .Fields("tipoboleto") = "I" Then optIda.Value = True
            If .Fields("tipoboleto") = "I/V" Then optIdaVuelta.Value = True
            meBoleto = FormatNumber(.Fields("monto_boleto"), 2)
            meEstancia = FormatNumber(.Fields("monto_estancia"), 2)
            If IsDate(.Fields("hora_salida")) Then
                meHoraSalida = .Fields("hora_salida")
            Else
                
            End If
            txtObs = .Fields("observacion")
        End With
    End If
    Set rstab = Nothing
End Sub

Private Function CargaCbo(cbo As MSForms.ComboBox, valor As String) As Integer
    Dim I As Integer
    CargaCbo = 0
    For I = 0 To cbo.ListCount - 1
        If cbo.List(I, 1) = valor Then
            CargaCbo = I
        End If
    Next
End Function

Private Function CargaCboPozo(cbo As MSForms.ComboBox, valor As String) As Integer
    Dim I As Integer
    CargaCboPozo = 0
    For I = 0 To cbo.ListCount - 1
        If cbo.List(I, 2) = valor Then
            CargaCboPozo = I
        End If
    Next
End Function

Private Function CargaCboMon(moneda As String) As Integer
    If moneda = "E" Then CargaCboMon = 1
    If moneda = "N" Then CargaCboMon = 0
End Function

Public Sub BloqueoControles(valor As Boolean)
    txtEmpleado_1.Locked = valor
    txtObs.Locked = valor
    cmbLote.Locked = valor
    cmbPozo.Locked = valor
    cmbDepart.Locked = valor
    CboDiv.Locked = valor
    cmbAgencia.Locked = valor
    cmbEstancia.Locked = valor
    cmbLinea.Locked = valor
    cmbMedio.Locked = valor
    cmbNombreEst.Locked = valor
    cmbMonBoleto.Locked = valor
    cmbMonEstancia.Locked = valor
    meHoraSalida.Enabled = Not valor
    meBoleto.Enabled = Not valor
    meEstancia.Enabled = Not valor
    txtEmpleado_2.Locked = valor
    txtAutoriz_lic.Locked = valor
    txtAutoriz_per.Locked = valor
    txtDias.Locked = valor
    txtMeses.Locked = valor
    txtMotivo_lic.Locked = valor
    txtMotivo_per.Locked = valor
    cmbPeriodo.Locked = valor
    chkGoceHaber.Enabled = Not valor
    chkIndemnizacion.Enabled = Not valor
    meHoraIni.Enabled = Not valor
    meHoraFin.Enabled = Not valor
    dpFin_lic.Enabled = Not valor
    dpFin_per.Enabled = Not valor
    dpFin_vac.Enabled = Not valor
    dpIni_lic.Enabled = Not valor
    dpIni_per.Enabled = Not valor
    dpIni_vac.Enabled = Not valor
    optEstadia.Enabled = Not valor
    optIda.Enabled = Not valor
    optIdaVuelta.Enabled = Not valor
    optNoche.Enabled = Not valor
    optFrame(0).Enabled = Not valor
    optFrame(1).Enabled = Not valor
    optFrame(2).Enabled = Not valor
    optFrame(3).Enabled = Not valor
    txtEmpleado_3.Locked = valor
    txtEmpleado_4.Locked = valor
    txtValorRefri.Locked = False
    dtinisub.Enabled = Not valor
    dtfinsub.Enabled = Not valor
    txtmotivosub.Locked = valor
    txtciit.Locked = valor
    CboTipo.Locked = valor
    If valor = True Then
        txtObs.BackColor = ColorDeshabilitado
        txtEmpleado_1.BackColor = ColorDeshabilitado
        txtEmpleado_2.BackColor = ColorDeshabilitado
        txtEmpleado_3.BackColor = ColorDeshabilitado
        txtEmpleado_4.BackColor = ColorDeshabilitado
        txtAutoriz_lic.BackColor = ColorDeshabilitado
        txtAutoriz_per.BackColor = ColorDeshabilitado
        txtDias.BackColor = ColorDeshabilitado
        txtMeses.BackColor = ColorDeshabilitado
        txtMotivo_lic.BackColor = ColorDeshabilitado
        txtMotivo_per.BackColor = ColorDeshabilitado
        txtValorRefri.BackColor = ColorDeshabilitado
        cmbLote.BackColor = ColorDeshabilitado
        cmbPozo.BackColor = ColorDeshabilitado
        cmbDepart.BackColor = ColorDeshabilitado
        CboDiv.BackColor = ColorDeshabilitado
        cmbAgencia.BackColor = ColorDeshabilitado
        cmbEstancia.BackColor = ColorDeshabilitado
        cmbLinea.BackColor = ColorDeshabilitado
        cmbMedio.BackColor = ColorDeshabilitado
        cmbNombreEst.BackColor = ColorDeshabilitado
        cmbPeriodo.BackColor = ColorDeshabilitado
        cmbMonBoleto.BackColor = ColorDeshabilitado
        cmbMonEstancia.BackColor = ColorDeshabilitado
        txtmotivosub.BackColor = ColorDeshabilitado
        txtciit.BackColor = ColorDeshabilitado
        CboTipo.BackColor = ColorDeshabilitado
    Else
        txtEmpleado_1.BackColor = ColorHabilitado
        txtEmpleado_2.BackColor = ColorHabilitado
        txtEmpleado_4.BackColor = ColorHabilitado
        txtAutoriz_lic.BackColor = ColorHabilitado
        txtAutoriz_per.BackColor = ColorHabilitado
        txtDias.BackColor = ColorHabilitado
        txtMeses.BackColor = ColorHabilitado
        txtMotivo_lic.BackColor = ColorHabilitado
        txtMotivo_per.BackColor = ColorHabilitado
        txtValorRefri.BackColor = ColorDeshabilitado
        cmbLote.BackColor = ColorHabilitado
        cmbPozo.BackColor = ColorHabilitado
        cmbDepart.BackColor = ColorHabilitado
        CboDiv.BackColor = ColorHabilitado
        cmbAgencia.BackColor = ColorHabilitado
        cmbEstancia.BackColor = ColorHabilitado
        cmbLinea.BackColor = ColorHabilitado
        cmbMedio.BackColor = ColorHabilitado
        cmbNombreEst.BackColor = ColorHabilitado
        cmbPeriodo.BackColor = ColorHabilitado
        cmbMonBoleto.BackColor = ColorHabilitado
        cmbMonEstancia.BackColor = ColorHabilitado
        txtmotivosub.BackColor = ColorHabilitado
        txtciit.BackColor = ColorHabilitado
        CboTipo.BackColor = ColorHabilitado
        txtObs.BackColor = ColorHabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
            Limpiar
            lblModo = "Acción"
            BloqueoControles True
            btnNuevo.Enabled = True
            btnSalir.Enabled = True
            btnmodificar.Enabled = True
            If SSTab1.Tab = 0 Then
                SSTab1.TabVisible(1) = False
                SSTab1.TabVisible(2) = False
            End If
            If SSTab1.Tab = 1 Then
                SSTab1.TabVisible(0) = False
                SSTab1.TabVisible(2) = False
            End If
            If SSTab1.Tab = 2 Then
                SSTab1.TabVisible(0) = False
                SSTab1.TabVisible(1) = False
            End If
            Exit Sub
        Case ModoForm.modNuevo
            lblModo = "Nuevo"
            Medio cmbMedio
            Estancia cmbEstancia
            agencia cmbAgencia
            Lote cmbLote
            Depart cmbDepart
            Divisiones CboDiv
            BloqueoControles False
            ConfigurarBotones cfgNuevo
            If Modificar Then btnmodificar.Enabled = True Else btnmodificar.Enabled = False
            cmbMonBoleto.ListIndex = 1
            cmbMonEstancia.ListIndex = 1
            optIda.Value = False
            optIdaVuelta.Value = False
            optEstadia.Value = False
            optNoche.Value = False
            TipoSuspension CboTipo
            If SSTab1.Tab = 1 Then
                If frmPrograma.tag = UCase(frameVac.Caption) Then frameVac.Enabled = True: framePer.Enabled = False: frameLic.Enabled = False: frameSub.Enabled = False: _
                    optFrame(0).Value = True: optFrame(1).Enabled = False: optFrame(2).Enabled = False: optFrame(3).Value = False: _
                    txtEmpleado_2.SetFocus: Exit Sub
                If frmPrograma.tag = UCase(framePer.Caption) Then frameVac.Enabled = False: framePer.Enabled = True: frameLic.Enabled = False: frameSub.Enabled = False: _
                    optFrame(1).Value = True: optFrame(0).Enabled = False: optFrame(2).Enabled = False: optFrame(3).Value = False: _
                    txtEmpleado_2.SetFocus: Exit Sub
                If frmPrograma.tag = UCase(frameLic.Caption) Then frameVac.Enabled = False: framePer.Enabled = False: frameLic.Enabled = True: frameSub.Enabled = False: _
                    optFrame(2).Value = True: optFrame(0).Enabled = False: optFrame(1).Enabled = False: optFrame(3).Value = False: _
                    txtEmpleado_2.SetFocus: Exit Sub
                If frmPrograma.tag = UCase(frameSub.Caption) Then frameVac.Enabled = False: framePer.Enabled = False: frameLic.Enabled = False: frameSub.Enabled = True: _
                    optFrame(3).Value = True: optFrame(0).Enabled = False: optFrame(1).Enabled = False: optFrame(2).Enabled = False: _
                    txtEmpleado_2.SetFocus: Exit Sub
            End If
            txtEmpleado_1.SetFocus
            Exit Sub
        Case ModoForm.modConsulta
            lblModo = "Consulta"
            BloqueoControles True
            ConfigurarBotones cfgGrabar
            If Modificar Then
                btnmodificar.Enabled = True
            Else
                btnmodificar.Enabled = False
            End If
            Exit Sub
        Case ModoForm.modEditar
            lblModo = "Modificar"
            BloqueoControles False
            ConfigurarBotones cfgModificar
            If SSTab1.Tab = 1 Then
                If frmPrograma.tag = UCase(frameVac.Caption) Then frameVac.Enabled = True: framePer.Enabled = False: frameLic.Enabled = False: frameSub.Enabled = False: _
                    optFrame(0).Value = True: optFrame(1).Enabled = False: optFrame(2).Enabled = False: optFrame(3).Value = False
                If frmPrograma.tag = UCase(framePer.Caption) Then frameVac.Enabled = False: framePer.Enabled = True: frameLic.Enabled = False: frameSub.Enabled = False: _
                    optFrame(1).Value = True: optFrame(0).Enabled = False: optFrame(2).Enabled = False: optFrame(3).Value = False
                If frmPrograma.tag = UCase(frameLic.Caption) Then frameVac.Enabled = False: framePer.Enabled = False: frameLic.Enabled = True: frameSub.Enabled = False: _
                    optFrame(2).Value = True: optFrame(0).Enabled = False: optFrame(1).Enabled = False: optFrame(3).Value = False
                If frmPrograma.tag = UCase(frameSub.Caption) Then frameVac.Enabled = False: framePer.Enabled = False: frameLic.Enabled = False: frameSub.Enabled = True: _
                    optFrame(3).Value = True: optFrame(0).Enabled = False: optFrame(1).Enabled = False: optFrame(2).Enabled = False
            End If
            Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            btnGrabar.Enabled = True
            btnCancelar.Enabled = True
            btnSalir.Enabled = True
            btnNuevo.Enabled = False
            btnmodificar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgModificar
            btnNuevo.Enabled = False
            btnmodificar.Enabled = False
            btnGrabar.Enabled = True
            btnReporte.Enabled = False
            btnCancelar.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgGrabar
            btnNuevo.Enabled = True
            btnCancelar.Enabled = False
            If Modificar Then
                btnmodificar.Enabled = True
            Else
                btnmodificar.Enabled = False
            End If
            btnGrabar.Enabled = False
            btnSalir.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo", "Consulta"
                    btnNuevo.Enabled = True
                    btnmodificar.Enabled = False
                    btnGrabar.Enabled = False
                    btnReporte.Enabled = False
                    btnCancelar.Enabled = False
                    ModoFormulario modAccion
                Case "Modificar"
                    ModoFormulario modAccion 'modConsulta
            End Select
    End Select
End Sub

Private Sub meBoleto_GotFocus()
    mark1 meBoleto
End Sub

Private Sub meBoleto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If meBoleto <> "" Then
            meBoleto = FormatNumber(meBoleto, 2)
        Else
            meBoleto = "0.00"
        End If
        cmbEstancia.SetFocus
    End If
End Sub

Private Sub meEstancia_GotFocus()
    mark1 meEstancia
End Sub

Private Sub meEstancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If meEstancia <> "" Then
            meEstancia = FormatNumber(meEstancia, 2)
        Else
            meEstancia = "0.00"
        End If
    End If
End Sub

Private Sub meHoraSalida_GotFocus()
    mark1 meHoraSalida
End Sub

Private Sub meHoraSalida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmbDepart.SetFocus
End Sub

Private Sub optEstadia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmbMonEstancia.SetFocus
End Sub

Private Sub optIda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then optIdaVuelta.SetFocus
End Sub

Private Sub optIdaVuelta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmbMonBoleto.SetFocus
End Sub

Private Sub optNoche_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then optEstadia.SetFocus
End Sub

Private Sub sbDias_Change()
    txtDias = CInt(sbDias.Value)
    dpFin_vac = dpIni_vac + (CInt(IIf(txtMeses = "", 0, txtMeses)) * 30 + CInt(IIf(txtDias = "", 0, txtDias))) - 1
End Sub

Private Sub sbMeses_Change()
   txtMeses = CInt(sbMeses.Value)
   dpFin_vac = dpIni_vac + (CInt(IIf(txtMeses = "", 0, txtMeses)) * 30 + CInt(IIf(txtDias = "", 0, txtDias)))
End Sub


Private Sub txtAutoriz_lic_GotFocus()
    mark1 txtAutoriz_lic
End Sub

Private Sub txtAutoriz_lic_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Encargados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_ENCARGADOS
            .Show
        End With
    End If
End Sub

Private Sub txtAutoriz_per_GotFocus()
    mark1 txtAutoriz_per
End Sub

Private Sub txtAutoriz_per_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Encargados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_ENCARGADOS
            .Show
        End With
    End If
End Sub

Private Sub txtciit_GotFocus()
    mark1 txtciit
End Sub

Private Sub txtdias_Change()
    If txtDias = Empty Then txtDias = 0: Exit Sub
    If txtDias <> Empty Or txtDias <> " " Then sbDias.Value = CInt(txtDias) Else txtDias = 0
End Sub

Private Sub txtdias_GotFocus()
    mark1 txtDias
End Sub

Private Sub txtDias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmbPeriodo.SetFocus
    End If
End Sub

Private Sub txtEmpleado_1_GotFocus()
    mark1 txtEmpleado_1
End Sub

Private Sub txtEmpleado_2_LostFocus()
    If frmPrograma.tag = "VACACIONES" Then
        Periodo cmbPeriodo
    End If
End Sub

Private Sub txtEmpleado_3_GotFocus()
    mark1 txtEmpleado_3
End Sub

Private Sub txtEmpleado_4_GotFocus()
    mark1 txtEmpleado_4
End Sub

Private Sub txtEmpleado_1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtEmpleado_1 = Right("00000000000" & Trim(txtEmpleado_1), 11)
        lblEmpleado_1 = DescripcionesdeCodigos("EMPLEADO", Trim(txtEmpleado_1))
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Empleados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_EMP_BONO
            .Show
        End With
    End If
End Sub

Private Function Bono(CodEmp As String) As Boolean
    Dim SQL As String
    Dim rsbono As MYSQL_RS
    Bono = False
    Tienebono = False
    SQL = "Select bono from contrato where codemp= '" & CodEmp & "' and estado='AP'"
    Set rsbono = oConexion.EjecutaSelectRS(SQL)
    If Not rsbono.EOF Then
        If rsbono.Fields("bono") = "S" Then
            Tienebono = True
            Bono = True
        Else
            Tienebono = False
            Bono = False
        End If
    End If
    Set rsbono = Nothing
End Function

Private Sub txtEmpleado_2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtEmpleado_2 = Right("00000000000" & Trim(txtEmpleado_2), 11)
        lblEmpleado_2 = DescripcionesdeCodigos("EMPLEADO", Trim(txtEmpleado_2))
        If frmPrograma.tag = "VACACIONES" Then
            Periodo cmbPeriodo
        End If
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Empleados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_EMP
            .Show
        End With
    End If
End Sub

Private Sub txtEmpleado_3_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtEmpleado_3 = Right("00000000000" & Trim(txtEmpleado_3), 11)
        lblEmpleado_3 = DescripcionesdeCodigos("EMPLEADO", Trim(txtEmpleado_3))
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Empleados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_EMP
            .Show
        End With
    End If
End Sub

Private Sub txtEmpleado_4_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtEmpleado_4 = Right("00000000000" & Trim(txtEmpleado_4), 11)
        lblEmpleado_4 = DescripcionesdeCodigos("EMPLEADO", Trim(txtEmpleado_4))
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1500
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Empleados"
            .pForm = FORM_PROGRAMACION
            .pCaso = LABEL_EMP
            .Show
        End With
    End If
End Sub

Private Sub Periodo(cbo As ComboBox)
    Dim SQL As String
    Dim rsper As MYSQL_RS
    Dim encontro As Integer
    Dim I As Integer, J As Integer
    SQL = " Select periodo from calendario where codemp = '" & Trim(txtEmpleado_2) & "'" & _
          " and MOVEMP='2' AND periodo <> '' AND PERIODO<>'NINGUNO' order by left(periodo,4)"
    Set rsper = oConexion.EjecutaSelectRS(SQL)
    cbo.Clear
    Do While Not rsper.EOF
        If CInt(Left(rsper.Fields("periodo"), 4)) >= CInt(Year(Date) - 5) Then
            For I = 0 To 11
                If Not (rsper.Fields("periodo") = Year(Date) - I & "-" & Year(Date) - I - 1) Then
                    ColocarPeriodo cbo, Year(Date) - I & "-" & Year(Date) - I + 1
                End If
            Next
        End If
        rsper.MoveNext
    Loop
     If rsper.RecordCount = 0 Then
        For I = 0 To 4
            cbo.AddItem Year(Date) - I - 1 & "-" & Year(Date) - I
        Next
    End If
End Sub

Private Sub ColocarPeriodo(cbo As ComboBox, valor As String)
    Dim I As Integer, encontrados As Integer
    encontrados = 0
    If cbo.ListCount = 0 Then cbo.AddItem valor: Exit Sub
    For I = 0 To cbo.ListCount - 1
        If cbo.List(I) = valor Then
            encontrados = encontrados + 1
        End If
    Next
    If encontrados = 0 Then cbo.AddItem valor
End Sub

Private Sub txtMeses_Change()
    If txtMeses = Empty Then txtMeses = 0: Exit Sub
    sbMeses.Value = CInt(IIf(CE(txtMeses) <> Empty, txtMeses, 0))
End Sub

Private Sub txtMeses_GotFocus()
    mark1 txtMeses
End Sub

Private Sub txtMeses_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDias.SetFocus
End Sub

Private Sub txtMotivo_lic_GotFocus()
    mark1 txtMotivo_lic
End Sub

Private Sub txtMotivo_per_GotFocus()
    mark1 txtMotivo_per
End Sub

Private Sub TxtObs_GotFocus()
    mark1 txtObs
End Sub

Private Sub ConfigGrilla(grilla As MSHFlexGrid)
On Error GoTo erroranio
    Dim I As Integer
    With grilla
        .Visible = True
        .Clear
        .Refresh
        .Rows = 2
        .Cols = 12
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 3500
        .ColWidth(1) = 450
        .ColWidth(2) = 450
        .ColWidth(3) = 450
        .ColWidth(4) = 450
        .ColWidth(5) = 450
        .ColWidth(6) = 450
        .ColWidth(7) = 450
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 0
        .row = 0
        .TextMatrix(0, 0) = Space(15) & "Empleado"
        .Col = 0
        .CellFontBold = True
        .TextMatrix(0, 1) = "Lun"
        .Col = 1
        .CellFontBold = True
        .TextMatrix(0, 2) = "Mar"
        .Col = 2
        .CellFontBold = True
        .TextMatrix(0, 3) = "Mie"
        .Col = 3
        .CellFontBold = True
        .TextMatrix(0, 4) = "Jue"
        .Col = 4
        .CellFontBold = True
        .TextMatrix(0, 5) = "Vie"
        .Col = 5
        .CellFontBold = True
        .TextMatrix(0, 6) = "Sab"
        .Col = 6
        .CellFontBold = True
        .TextMatrix(0, 7) = "Dom"
        .Col = 7
        .CellFontBold = True
        .TextMatrix(0, 8) = "Almuerzo"
        .Col = 8
        .CellFontBold = True
        .TextMatrix(0, 9) = "Cena"
        .Col = 9
        .CellFontBold = True
        .TextMatrix(0, 10) = "Valor"
        .Col = 10
        .CellFontBold = True
    End With
Exit Sub
erroranio:
    Exit Sub
End Sub

Private Sub txtValorRefri_GotFocus()
    mark txtValorRefri
End Sub
