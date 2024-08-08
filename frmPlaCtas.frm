VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmplancuentas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Del Plan De Cuentas"
   ClientHeight    =   6825
   ClientLeft      =   885
   ClientTop       =   1905
   ClientWidth     =   9420
   Icon            =   "frmPlaCtas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9420
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1005
      TabIndex        =   73
      ToolTipText     =   "Editar"
      Top             =   6375
      Width           =   885
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   135
      TabIndex        =   72
      ToolTipText     =   "Nuevo"
      Top             =   6375
      Width           =   870
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1890
      TabIndex        =   71
      ToolTipText     =   "Borrar"
      Top             =   6375
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   375
      Left            =   6060
      Picture         =   "frmPlaCtas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Salir"
      Top             =   6390
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdImprimir 
      Height          =   375
      Left            =   8565
      Picture         =   "frmPlaCtas.frx":059C
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Imprimir"
      Top             =   6390
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton CmdVistaPreliminar 
      Height          =   375
      Left            =   8940
      Picture         =   "frmPlaCtas.frx":069E
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Preliminar"
      Top             =   6390
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdGrabar 
      Height          =   375
      Left            =   4665
      Picture         =   "frmPlaCtas.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Grabar"
      Top             =   6390
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   375
      Left            =   4290
      Picture         =   "frmPlaCtas.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Deshacer"
      Top             =   6390
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5865
      Left            =   90
      TabIndex        =   9
      Top             =   105
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   10345
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "Listado Del Plan De Cuentas"
      TabPicture(0)   =   "frmPlaCtas.frx":0E1C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle Del Plan De Cuentas"
      TabPicture(1)   =   "frmPlaCtas.frx":0E38
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameCbos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrmOption"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   500
         Left            =   135
         TabIndex        =   41
         Top             =   360
         Width           =   9015
         Begin VB.TextBox TxtAuxiliar 
            Height          =   285
            Left            =   8190
            MaxLength       =   2
            TabIndex        =   70
            Top             =   165
            Width           =   375
         End
         Begin VB.CommandButton CmdFind 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   8640
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   150
            Width           =   285
         End
         Begin VB.CommandButton CmdFind 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1365
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   150
            Width           =   285
         End
         Begin VB.ComboBox CboTipCta 
            Height          =   315
            ItemData        =   "frmPlaCtas.frx":0E54
            Left            =   2085
            List            =   "frmPlaCtas.frx":0E5E
            TabIndex        =   2
            Top             =   140
            Width           =   870
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   285
            Left            =   3045
            MaxLength       =   40
            TabIndex        =   1
            Top             =   150
            Width           =   4335
         End
         Begin VB.TextBox TxtCuenta 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   0
            Top             =   150
            Width           =   735
         End
         Begin VB.Label LblTipCta 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Auxiliar"
            Height          =   195
            Index           =   1
            Left            =   7605
            TabIndex        =   44
            Top             =   240
            Width           =   495
         End
         Begin VB.Label LblTipCta 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo"
            Height          =   195
            Index           =   0
            Left            =   1725
            TabIndex        =   43
            Top             =   240
            Width           =   315
         End
         Begin VB.Label LblCuenta 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cuenta"
            Height          =   195
            Left            =   60
            TabIndex        =   42
            Top             =   200
            Width           =   510
         End
      End
      Begin VB.Frame FrmOption 
         BackColor       =   &H00E0E0E0&
         Height          =   4065
         Left            =   135
         TabIndex        =   31
         Top             =   800
         Width           =   9015
         Begin VB.CheckBox ChkDivision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Division"
            Height          =   195
            Left            =   60
            TabIndex        =   87
            Top             =   1155
            Width           =   1335
         End
         Begin VB.ComboBox CboMovi 
            Height          =   315
            ItemData        =   "frmPlaCtas.frx":0E73
            Left            =   90
            List            =   "frmPlaCtas.frx":0E80
            TabIndex        =   86
            Text            =   "CboMovi"
            Top             =   3675
            Width           =   1455
         End
         Begin VB.TextBox txtColVen 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   83
            Top             =   3165
            Width           =   375
         End
         Begin VB.TextBox txtColCom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   82
            Top             =   2835
            Width           =   375
         End
         Begin MSMask.MaskEdBox txtPorcal 
            Height          =   285
            Left            =   1170
            TabIndex        =   79
            Top             =   1875
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   2
            Format          =   "##"
            Mask            =   "##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtEstado 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   78
            Top             =   2175
            Width           =   375
         End
         Begin VB.TextBox txtLin_i_e 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   76
            Top             =   2505
            Width           =   375
         End
         Begin VB.ComboBox CboMoneda 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmPlaCtas.frx":0E98
            Left            =   60
            List            =   "frmPlaCtas.frx":0EA5
            TabIndex        =   7
            Text            =   "CboMoneda"
            Top             =   1590
            Width           =   1455
         End
         Begin VB.CheckBox ChkGenerada 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Generada"
            Height          =   195
            Left            =   1680
            TabIndex        =   8
            Top             =   180
            Width           =   1095
         End
         Begin VB.CheckBox ChkCtaCte 
            BackColor       =   &H80000013&
            Caption         =   "Cuenta Corriente"
            Height          =   195
            Left            =   60
            TabIndex        =   5
            Top             =   660
            Width           =   1575
         End
         Begin VB.CheckBox ChkCentroCosto 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Centro Costo"
            Height          =   195
            Left            =   60
            TabIndex        =   4
            Top             =   420
            Width           =   1215
         End
         Begin VB.CheckBox ChkLote 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lote/Pozo"
            Height          =   195
            Left            =   60
            TabIndex        =   6
            Top             =   900
            Width           =   1335
         End
         Begin VB.CheckBox ChkBancos 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Bancos"
            Height          =   195
            Left            =   60
            TabIndex        =   3
            Top             =   180
            Width           =   855
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   3705
            Left            =   1710
            TabIndex        =   32
            Top             =   210
            Width           =   7170
            Begin VB.ComboBox CboPorcentaje1 
               Height          =   315
               Index           =   4
               ItemData        =   "frmPlaCtas.frx":0EC6
               Left            =   6195
               List            =   "frmPlaCtas.frx":0EE8
               TabIndex        =   19
               Top             =   2010
               Width           =   855
            End
            Begin VB.ComboBox CboPorcentaje1 
               Height          =   315
               Index           =   2
               ItemData        =   "frmPlaCtas.frx":0F1F
               Left            =   6225
               List            =   "frmPlaCtas.frx":0F41
               TabIndex        =   15
               Top             =   1290
               Width           =   855
            End
            Begin VB.ComboBox CboPorcentaje1 
               Height          =   315
               Index           =   1
               ItemData        =   "frmPlaCtas.frx":0F78
               Left            =   6210
               List            =   "frmPlaCtas.frx":0F9A
               TabIndex        =   13
               Top             =   930
               Width           =   855
            End
            Begin VB.TextBox TxtCargo1 
               Height          =   315
               Index           =   4
               Left            =   1200
               MaxLength       =   7
               TabIndex        =   18
               Top             =   2010
               Width           =   735
            End
            Begin VB.TextBox TxtCargo1 
               Height          =   315
               Index           =   3
               Left            =   1200
               MaxLength       =   7
               TabIndex        =   16
               Top             =   1650
               Width           =   735
            End
            Begin VB.TextBox TxtCargo1 
               Height          =   315
               Index           =   2
               Left            =   1200
               MaxLength       =   7
               TabIndex        =   14
               Top             =   1290
               Width           =   735
            End
            Begin VB.TextBox TxtCargo1 
               Height          =   300
               Index           =   1
               Left            =   1215
               MaxLength       =   7
               TabIndex        =   12
               Top             =   945
               Width           =   720
            End
            Begin VB.TextBox TxtAbono 
               Height          =   315
               Left            =   1185
               MaxLength       =   7
               TabIndex        =   20
               Top             =   2955
               Width           =   735
            End
            Begin VB.TextBox TxtCargo1 
               Height          =   315
               Index           =   0
               Left            =   1215
               MaxLength       =   7
               TabIndex        =   10
               Top             =   600
               Width           =   720
            End
            Begin VB.ComboBox CboPorcentaje1 
               Height          =   315
               Index           =   0
               ItemData        =   "frmPlaCtas.frx":0FD1
               Left            =   6210
               List            =   "frmPlaCtas.frx":0FF3
               TabIndex        =   11
               Top             =   600
               Width           =   855
            End
            Begin VB.ComboBox CboPorcentaje1 
               Height          =   315
               Index           =   3
               ItemData        =   "frmPlaCtas.frx":102A
               Left            =   6210
               List            =   "frmPlaCtas.frx":104C
               TabIndex        =   17
               Top             =   1650
               Width           =   855
            End
            Begin VB.Label LblDesCargo1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   4
               Left            =   1950
               TabIndex        =   53
               Top             =   2010
               Width           =   4215
            End
            Begin VB.Label LblDesCargo1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   3
               Left            =   1950
               TabIndex        =   52
               Top             =   1650
               Width           =   4215
            End
            Begin VB.Label LblDesCargo1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   2
               Left            =   1950
               TabIndex        =   51
               Top             =   1290
               Width           =   4215
            End
            Begin VB.Label LblDesCargo1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   1
               Left            =   1950
               TabIndex        =   50
               Top             =   930
               Width           =   4215
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargo5"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   60
               TabIndex        =   49
               Top             =   2010
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargo4"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   60
               TabIndex        =   48
               Top             =   1650
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargo3"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   47
               Top             =   1290
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargo2"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   60
               TabIndex        =   46
               Top             =   930
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargo1"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   60
               TabIndex        =   45
               Top             =   570
               Width           =   1095
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Descripción"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1950
               TabIndex        =   40
               Top             =   2520
               Width           =   4215
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Abono"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1185
               TabIndex        =   39
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label LblDesAbono 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1950
               TabIndex        =   38
               Top             =   2940
               Width           =   4215
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cargos"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   37
               Top             =   285
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Porcentaje"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   6195
               TabIndex        =   36
               Top             =   285
               Width           =   855
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Descripción"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   1950
               TabIndex        =   35
               Top             =   285
               Width           =   4215
            End
            Begin VB.Label LblDesCargo1 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   0
               Left            =   1950
               TabIndex        =   34
               Top             =   600
               Width           =   4215
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cuenta"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1185
               TabIndex        =   33
               Top             =   285
               Width           =   735
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Moneda"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   88
            Top             =   1380
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   85
            Top             =   3465
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Col. Ventas"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   81
            Top             =   3195
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Col. Compras"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   80
            Top             =   2895
            Width           =   930
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo Gasto"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   77
            Top             =   2220
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ingreso/Gasto"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   75
            Top             =   2535
            Width           =   1020
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "% Calculo"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   74
            Top             =   1935
            Width           =   690
         End
      End
      Begin VB.Frame FrameCbos 
         BackColor       =   &H00E0E0E0&
         Height          =   975
         Left            =   150
         TabIndex        =   26
         Top             =   4815
         Width           =   9015
         Begin VB.ComboBox CboEstFin 
            Height          =   315
            ItemData        =   "frmPlaCtas.frx":1083
            Left            =   1440
            List            =   "frmPlaCtas.frx":1085
            TabIndex        =   21
            Top             =   165
            Width           =   3015
         End
         Begin VB.ComboBox CboLineasGPN 
            Height          =   315
            Left            =   1440
            TabIndex        =   23
            Top             =   555
            Width           =   3015
         End
         Begin VB.ComboBox CboLineasIE 
            Height          =   315
            Left            =   5880
            TabIndex        =   24
            Top             =   555
            Width           =   3015
         End
         Begin VB.ComboBox CboLineasBGF 
            Height          =   315
            Left            =   5880
            TabIndex        =   22
            Top             =   165
            Width           =   3015
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Estado Financiero"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   30
            Top             =   225
            Width           =   1275
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lineas de BGF"
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   29
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lineas GPN"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   645
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Lineas I.E."
            Height          =   195
            Index           =   3
            Left            =   4680
            TabIndex        =   27
            Top             =   615
            Width           =   750
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Height          =   5400
         Left            =   -74880
         TabIndex        =   25
         Top             =   375
         Width           =   9015
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Buscar"
            Height          =   495
            Left            =   75
            TabIndex        =   61
            Top             =   4845
            Width           =   5775
            Begin VB.TextBox TxtCriterio 
               Height          =   285
               Left            =   2370
               TabIndex        =   64
               Top             =   150
               Width           =   3240
            End
            Begin VB.OptionButton OptDescripcion 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Descripción"
               Height          =   195
               Left            =   1110
               TabIndex        =   63
               Top             =   200
               Width           =   1230
            End
            Begin VB.OptionButton OptCuenta 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Cuenta"
               Height          =   195
               Left            =   60
               TabIndex        =   62
               Top             =   200
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.ComboBox CboOrdenar 
            Height          =   315
            ItemData        =   "frmPlaCtas.frx":1087
            Left            =   7200
            List            =   "frmPlaCtas.frx":1091
            TabIndex        =   59
            Top             =   4980
            Width           =   1575
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshPlanDeCuentas 
            Height          =   4620
            Left            =   45
            TabIndex        =   60
            Top             =   150
            Width           =   8925
            _ExtentX        =   15743
            _ExtentY        =   8149
            _Version        =   393216
            BackColorBkg    =   12632256
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ordenar"
            Height          =   195
            Index           =   1
            Left            =   6480
            TabIndex        =   58
            Top             =   5100
            Width           =   570
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   195
      Index           =   6
      Left            =   300
      TabIndex        =   84
      Top             =   4170
      Width           =   585
   End
   Begin VB.Label LblMensaje 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8010
      TabIndex        =   55
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F1 List De Plan De Cuentas  F2 List De Auxiliares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   90
      TabIndex        =   54
      Top             =   6000
      Width           =   8055
   End
End
Attribute VB_Name = "frmplancuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CantRegistros As Integer
Public MshHabilitado As Boolean
Public FilaSel As Integer
Public tmpCuenta As String
Public tmpFila As Integer
Sub InicarPorcentajes()
    Dim i As Integer
    For i = 0 To 4
        CboPorcentaje1(i).Text = "0%"
    Next
End Sub
Private Function La_Cuenta_Tiene_Mov(cuenta As String) As Boolean
    If sp_La_Cuenta_Tiene_Mov(CodigoEmpresa, cuenta).RecordCount > 0 Then
        La_Cuenta_Tiene_Mov = True
    Else
        La_Cuenta_Tiene_Mov = False
    End If
End Function

Public Function sp_La_Cuenta_Tiene_Mov(Codcia As String, cuenta As String) As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim SQL As String
    
    SQL = " SELECT *"
    SQL = SQL & " From cnmovi "
    SQL = SQL & "              WHERE cuenta = '" & cuenta & "'"
    
    Set sp_La_Cuenta_Tiene_Mov = ADOConexionEmp.Execute(SQL)
End Function

Sub ModoEdicionTituloPorcentajes()
    Dim i As Integer
    For i = 0 To 4
        TxtCargo1(i) = Empty
        CboPorcentaje1(i) = "0%"
        TxtCargo1(i).Locked = True
        TxtCargo1(i).BackColor = ColorDeshabilitado
        CboPorcentaje1(i).Locked = True
        CboPorcentaje1(i).BackColor = ColorDeshabilitado
    Next
End Sub
Sub ModoEdicionTitulo()
    txtPorcal.Enabled = False
    txtPorcal.BackColor = ColorDeshabilitado
    txtPorcal = "__"
    txtLin_i_e.Locked = True
    txtLin_i_e.BackColor = ColorDeshabilitado
    txtLin_i_e = Empty
    txtEstado.Locked = True
    txtEstado.BackColor = ColorDeshabilitado
    txtEstado = Empty
    TxtDescripcion = Empty
    TxtDescripcion.Locked = False
    TxtDescripcion.BackColor = ColorHabilitado
    TxtAuxiliar = Empty
    TxtAuxiliar.Locked = True
    TxtAuxiliar.BackColor = ColorDeshabilitado
    CboMoneda = Empty
    CboMoneda.Locked = True
    CboMoneda.BackColor = ColorDeshabilitado

    Call ModoEdicionTituloPorcentajes
    TxtAbono.Locked = True
    TxtAbono.BackColor = ColorDeshabilitado
    FrmOption.Enabled = False
    
    CboEstFin = Empty
    CboEstFin.Locked = True
    CboEstFin.BackColor = ColorDeshabilitado
    CboLineasBGF = Empty
    CboLineasBGF.Locked = True
    CboLineasBGF.BackColor = ColorDeshabilitado
    CboLineasGPN = Empty
    CboLineasGPN.Locked = True
    CboLineasGPN.BackColor = ColorDeshabilitado
    CboLineasIE = Empty
    CboLineasIE.Locked = True
    CboLineasIE.BackColor = ColorDeshabilitado
    FrameCbos.Enabled = False
End Sub

Sub BorrarRegistro()
On Error GoTo Errdel
    Dim SQL As String
    SQL = "DELETE FROM CNMAYOR WHERE CUENTA ='" + Trim(TxtCuenta) + "'"
    
    ADOConexionEmp.BeginTrans
    ADOConexionEmp.Execute (SQL)
    ADOConexionEmp.CommitTrans
    Exit Sub
Errdel:
    MsgBox "Ha ocurrido un error al momento de eliminar " & Chr(13) & Err.Description, vbCritical, "Error de datos"
    ADOConexionEmp.RollbackTrans
End Sub

Private Function BuscarAuxiliar(Auxiliar As String) As String
    
    Dim RS As ADODB.Recordset
    Dim SQL As String
    
    Set RS = New ADODB.Recordset
    
    
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenStatic
    RS.LockType = adLockBatchOptimistic
    
    
    SQL = "SELECT tip_linea  FROM CNTABLAS WHERE  codtab='1' and tip_linea ='" + Trim(Auxiliar) + "'"

    Set RS = ADOConexionEmp.Execute(SQL)
    If RS.BOF Or RS.EOF Then
        BuscarAuxiliar = Empty
    Else
        BuscarAuxiliar = RS(0)
    End If
    RS.Close
    Set RS = Nothing
    
End Function

Private Function BuscarCuenta(cuenta As String) As Integer
    Dim i As Integer
    With MshPlanDeCuentas
        For i = 1 To CantRegistros
            If Trim(.TextMatrix(i, 1)) = Trim(cuenta) Then
                BuscarCuenta = i
                Exit Function
            End If
        Next
        BuscarCuenta = 0
    End With
End Function

Private Function BuscarDescripcion(cuenta As String) As String
    Dim i As Integer
    With MshPlanDeCuentas
        For i = 1 To CantRegistros
            If Trim(.TextMatrix(i, 1)) = Trim(cuenta) Then
                BuscarDescripcion = .TextMatrix(i, 2)
                Exit Function
            End If
        Next
        BuscarDescripcion = Empty
    End With
End Function

Private Function BuscarEstFin(CodEstFin As String) As String
    Dim i As Integer
    For i = 0 To CboEstFin.ListCount - 1
        If Trim(Left(CboEstFin.List(i), 3)) = Trim(CodEstFin) Then
            BuscarEstFin = CboEstFin.List(i)
            Exit Function
        End If
    Next
    BuscarEstFin = Empty
End Function
Private Function BuscarLineasBGF(CodLineasBGF As String) As String
    Dim i As Integer
    For i = 0 To CboLineasBGF.ListCount - 1
        If Trim(Left(CboLineasBGF.List(i), 3)) = Trim(CodLineasBGF) Then
            BuscarLineasBGF = CboLineasBGF.List(i)
            Exit Function
        End If
    Next
    BuscarLineasBGF = Empty
End Function
Private Function BuscarLineasGPN(CodLineasGPN As String) As String
    Dim i As Integer
    For i = 0 To CboLineasGPN.ListCount - 1
        If Trim(Left(CboLineasGPN.List(i), 3)) = Trim(CodLineasGPN) Then
            BuscarLineasGPN = CboLineasGPN.List(i)
            Exit Function
        End If
    Next
    BuscarLineasGPN = Empty
End Function
Private Function BuscarLineasIE(CodLineasIE As String) As String
    Dim i As Integer
    For i = 0 To CboLineasIE.ListCount - 1
        If Trim(Left(CboLineasIE.List(i), 3)) = Trim(CodLineasIE) Then
            BuscarLineasIE = CboLineasIE.List(i)
            Exit Function
        End If
    Next
    BuscarLineasIE = Empty
End Function
Sub CalcularPorcentaje(ValInicialIndex As Integer)
    Dim SumaCombos, i As Integer
    
    'SumaCombos = Val(CboPorcentaje1(0).Text) + Val(CboPorcentaje1(1).Text) + Val(CboPorcentaje1(2).Text) + Val(CboPorcentaje1(3).Text) + Val(CboPorcentaje1(4).Text)
    SumaCombos = Val(IIf(CboPorcentaje1(0).Locked, 0, CboPorcentaje1(0).Text))
    SumaCombos = SumaCombos + Val(IIf(CboPorcentaje1(1).Locked, 0, CboPorcentaje1(1).Text))
    SumaCombos = SumaCombos + Val(IIf(CboPorcentaje1(2).Locked, 0, CboPorcentaje1(2).Text))
    SumaCombos = SumaCombos + Val(IIf(CboPorcentaje1(3).Locked, 0, CboPorcentaje1(3).Text))
    SumaCombos = SumaCombos + Val(IIf(CboPorcentaje1(4).Locked, 0, CboPorcentaje1(4).Text))
    
    If SumaCombos = 100 Then
        For i = ValInicialIndex + 1 To 4
            TxtCargo1(i).Text = Empty
            TxtCargo1(i).Locked = True
            TxtCargo1(i).BackColor = ColorDeshabilitado
            LblDesCargo1(i) = Empty
            CboPorcentaje1(i).Locked = True
            CboPorcentaje1(i).Text = "0%"
            CboPorcentaje1(i).BackColor = ColorDeshabilitado
        Next
        'CboEstFin.SetFocus
        TxtAbono.SetFocus
    Else
        If SumaCombos > 100 Then
            TxtCargo1(ValInicialIndex) = Empty
            TxtCargo1(ValInicialIndex).Locked = True
            TxtCargo1(ValInicialIndex).BackColor = ColorDeshabilitado
            CboPorcentaje1(ValInicialIndex).Locked = True
            CboPorcentaje1(ValInicialIndex) = "0%"
            CboPorcentaje1(ValInicialIndex).BackColor = ColorDeshabilitado
            Exit Sub
        Else
            If (ValInicialIndex + 1) < 5 Then
                TxtCargo1(ValInicialIndex + 1).Locked = False
                TxtCargo1(ValInicialIndex + 1).BackColor = ColorHabilitado
                CboPorcentaje1(ValInicialIndex + 1).Locked = False
                CboPorcentaje1(ValInicialIndex + 1).BackColor = ColorHabilitado
                Call CboPorcentaje1_KeyPress(ValInicialIndex, 13)
            Else
                Dim R As Integer
                R = Val(Trim(CboPorcentaje1(0).Text)) + Val(Trim(CboPorcentaje1(1).Text)) + Val(Trim(CboPorcentaje1(2).Text)) + Val(Trim(CboPorcentaje1(3).Text)) + Val(Trim(CboPorcentaje1(4).Text))
                R = 100 - R
                MsgBox "Por Favor Distribuya El Porcentaje Restante    " + STR(R) + " %", vbExclamation, "Plan De Cuentas"
                CboPorcentaje1(4).SetFocus
                SendKeys "{HOME}+{END}"
            End If
        End If
    End If
End Sub

Sub ConfigMshPlanDeCuentas()
    With MshPlanDeCuentas
        .Cols = 32
        .FixedCols = 1
        .Rows = 2
        .Clear
        .ColWidth(0) = 0
        .ColWidth(1) = 700
        .ColWidth(2) = 4215
        .ColWidth(3) = 400
        .ColWidth(4) = 400
        .ColWidth(5) = 400
        .ColWidth(6) = 500
        .ColWidth(7) = 700
        .ColWidth(8) = 650
        .ColWidth(9) = 750
        .ColWidth(10) = 500
        .ColWidth(11) = 700
        .ColWidth(12) = 500
        .ColWidth(13) = 700
        .ColWidth(14) = 700
        .ColWidth(15) = 700
        .ColWidth(16) = 700
        .ColWidth(17) = 700
        .ColWidth(18) = 700
        .ColWidth(19) = 700
        .ColWidth(20) = 700
        .ColWidth(21) = 700
        .ColWidth(22) = 700
        .ColWidth(23) = 700
        .ColWidth(24) = 700
        .ColWidth(25) = 700
        .ColWidth(26) = 700
        .ColWidth(27) = 700
        .ColWidth(28) = 700
        .ColWidth(29) = 700
        .ColWidth(30) = 700
        .ColWidth(31) = 700
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Descrip"
        .TextMatrix(0, 3) = "Tipo"
        .TextMatrix(0, 4) = "Aux."
        .TextMatrix(0, 5) = "Mon."
        .TextMatrix(0, 6) = "Bcos"
        .TextMatrix(0, 7) = "Cen.Cos."
        .TextMatrix(0, 8) = "Cta.Cte"
        .TextMatrix(0, 9) = "Agencia" '"Cod.Dep"
        .TextMatrix(0, 10) = "Gnda"
        .TextMatrix(0, 11) = "TGasto" '"Cod.Funcion"
        .TextMatrix(0, 12) = "Est.Fin."
        .TextMatrix(0, 13) = "Lin.Bgf"
        .TextMatrix(0, 14) = "Lin.GpN"
        .TextMatrix(0, 15) = "Lin.I.E"
        .TextMatrix(0, 16) = "Cargo.1"
        .TextMatrix(0, 17) = "Porc.1"
        .TextMatrix(0, 18) = "Cargo.2"
        .TextMatrix(0, 19) = "Porc.2"
        .TextMatrix(0, 20) = "Cargo.3"
        .TextMatrix(0, 21) = "Porc.3"
        .TextMatrix(0, 22) = "Cargo.4"
        .TextMatrix(0, 23) = "Porc.4"
        .TextMatrix(0, 24) = "Cargo.5"
        .TextMatrix(0, 25) = "Porc.5"
        .TextMatrix(0, 26) = "Abono"
        .TextMatrix(0, 27) = "Porcal" '"Dif.Cam"
        .TextMatrix(0, 28) = "ColCom"
        .TextMatrix(0, 29) = "ColVen"
        .TextMatrix(0, 30) = "Movi"
        .TextMatrix(0, 31) = "Divi"
        .FocusRect = flexFocusNone
        '.SelectionMode = flexSelectionByRow
    End With
End Sub
Sub LlenarCboEEFF()
  Dim RsEstFin  As ADODB.Recordset
  Set RsEstFin = New ADODB.Recordset
  Dim SQL As String
  
  RsEstFin.CursorLocation = adUseClient
  RsEstFin.CursorType = adOpenStatic
  RsEstFin.LockType = adLockBatchOptimistic
    
    CboEstFin.Clear
    
    SQL = "SELECT TIP_LINEA,DESCRIP FROM cntablas where codtab ='2' order by 1"
    Set RsEstFin = ADOConexionEmp.Execute(SQL)
    Do While Not RsEstFin.EOF
        CboEstFin.AddItem RsEstFin(0) + "  " + RsEstFin(1)
        RsEstFin.MoveNext
    Loop
    
    RsEstFin.Close
    Set RsEstFin = Nothing
    
End Sub
Sub LlenarCbos()
    Dim RsLineasBGF As ADODB.Recordset
    Dim RsLineasGPN As ADODB.Recordset
    Dim RsLineas_IE As ADODB.Recordset
    Dim SQL As String

    Set RsLineasBGF = New ADODB.Recordset
    Set RsLineasGPN = New ADODB.Recordset
    Set RsLineas_IE = New ADODB.Recordset
    
    RsLineasBGF.CursorLocation = adUseClient
    RsLineasBGF.CursorType = adOpenStatic
    RsLineasBGF.LockType = adLockBatchOptimistic
    
    RsLineasGPN.CursorLocation = adUseClient
    RsLineasGPN.CursorType = adOpenStatic
    RsLineasGPN.LockType = adLockBatchOptimistic
    
    RsLineas_IE.CursorLocation = adUseClient
    RsLineas_IE.CursorType = adOpenStatic
    RsLineas_IE.LockType = adLockBatchOptimistic
    
    CboLineasBGF.Clear
    CboLineasGPN.Clear
    CboLineasIE.Clear
    
    Dim EEFF As String
    EEFF = Left(Trim(CboEstFin), 1)

    If EEFF = "4" Then EEFF = "2"
    
    SQL = "select LINEA,DESCRIP from cnlineas where TIPO ='" & EEFF & "' order by 1"
    Set RsLineasBGF = ADOConexionEmp.Execute(SQL)
    Do While Not RsLineasBGF.EOF
        CboLineasBGF.AddItem RsLineasBGF(0) + "  " + RsLineasBGF(1)
        RsLineasBGF.MoveNext
    Loop
    
    SQL = "select LINEA,DESCRIP from cnlineas where TIPO ='3' order by 1"
    Set RsLineasGPN = ADOConexionEmp.Execute(SQL)
    Do While Not RsLineasGPN.EOF
        CboLineasGPN.AddItem RsLineasGPN(0) + "  " + RsLineasGPN(1)
        RsLineasGPN.MoveNext
    Loop
    
    SQL = "select LINEA,DESCRIP from cnlineas where TIPO ='5' order by 1"
    Set RsLineas_IE = ADOConexionEmp.Execute(SQL)
    Do While Not RsLineas_IE.EOF
        CboLineasIE.AddItem RsLineas_IE(0) + "  " + RsLineas_IE(1)
        RsLineas_IE.MoveNext
    Loop
    
    RsLineasBGF.Close
    RsLineasGPN.Close
    RsLineas_IE.Close
    
    Set RsLineasBGF = Nothing
    Set RsLineasGPN = Nothing
    Set RsLineas_IE = Nothing
    
    
End Sub
Sub BloqueoDeBotones()
  cmdNuevo.Enabled = True
  CmdModificar.Enabled = False
  CmdEliminar.Enabled = False
  cmdGrabar.Enabled = False
  cmdCancelar.Enabled = False
  CmdVistaPreliminar.Enabled = False
  cmdImprimir.Enabled = False
End Sub

Sub LlenarMshPlanDeCuentas()
  Dim SQL As String
  SQL = "SELECT CUENTA,DESCRIP,TIPO,movi,colcom,colven,porcal, estado,AUXILIAR,MONEDA," & _
        " BANCOS,CEN_COS,CTA_CTE, COD_DEP,GENERADA,COD_FUN,EST_FIN,LIN_BGF,LIN_GPN," & _
        " LIN_I_E,CARGO1,PORC1,CARGO2,PORC2,CARGO3,PORC3,CARGO4,PORC4,CARGO5,PORC5," & _
        " ABONO,0 as DIFCAM, cod_dep FROM cnmayor ORDER BY 1"
  
  Dim RS As ADODB.Recordset
  Dim RSCant As ADODB.Recordset
  Set RS = New ADODB.Recordset
  Set RSCant = New ADODB.Recordset
  
  RS.CursorLocation = adUseClient
  RS.CursorType = adOpenStatic
  RS.LockType = adLockBatchOptimistic

  RSCant.CursorLocation = adUseClient
  RSCant.CursorType = adOpenStatic
  RSCant.LockType = adLockBatchOptimistic
  
    If RS.State = ADODB.adOpenStatic Then RS.Close
    Set RS = ADOConexionEmp.Execute(SQL)
    Call ConfigMshPlanDeCuentas
    Dim i As Integer
    With MshPlanDeCuentas
        .Redraw = False
        CantRegistros = RS.RecordCount
      If RS.BOF = False And RS.EOF = False Then
        For i = 1 To CantRegistros
            .TextMatrix(.Rows - 1, 1) = Trim(CE(RS.Fields("cuenta")))
            If .TextMatrix(.Rows - 1, 3) = "T" Then
                .TextMatrix(.Rows - 1, 2) = Trim(CE(RS.Fields("descrip")))
                .CellFontBold = True
            Else
                .TextMatrix(.Rows - 1, 2) = Trim(CE(RS.Fields("descrip")))
            End If
            
            If .TextMatrix(.Rows - 1, 3) = "D" Then
                .TextMatrix(.Rows - 1, 2) = Trim(CE(RS.Fields("descrip")))
                .CellFontBold = False
            End If
            .TextMatrix(.Rows - 1, 3) = Trim(CE(RS.Fields("tipo")))
            .TextMatrix(.Rows - 1, 4) = Trim(CE(RS.Fields("auxiliar")))
            .TextMatrix(.Rows - 1, 5) = Trim(CE(RS.Fields("moneda")))
            .TextMatrix(.Rows - 1, 6) = Trim(CE(RS.Fields("bancos")))
            .TextMatrix(.Rows - 1, 7) = Trim(CE(RS.Fields("cen_cos")))
            .TextMatrix(.Rows - 1, 8) = Trim(CE(RS.Fields("cta_cte")))
            .TextMatrix(.Rows - 1, 9) = Trim(CE(RS.Fields("cod_fun"))) 'lote pozo en el plan de cuentas
            .TextMatrix(.Rows - 1, 10) = Trim(CE(RS.Fields("generada")))
            .TextMatrix(.Rows - 1, 11) = Trim(CE(RS.Fields("estado"))) 'tipo gasto
            .TextMatrix(.Rows - 1, 12) = Trim(CE(RS.Fields("est_fin")))
            .TextMatrix(.Rows - 1, 13) = Trim(CE(RS.Fields("lin_bgf")))
            .TextMatrix(.Rows - 1, 14) = Trim(CE(RS.Fields("lin_gpn")))
            .TextMatrix(.Rows - 1, 15) = Trim(CE(RS.Fields("lin_i_e")))
            .TextMatrix(.Rows - 1, 16) = Trim(CE(RS.Fields("cargo1")))
            .TextMatrix(.Rows - 1, 17) = Trim(CE(RS.Fields("porc1")))
            .TextMatrix(.Rows - 1, 18) = Trim(CE(RS.Fields("cargo2")))
            .TextMatrix(.Rows - 1, 19) = Trim(CE(RS.Fields("porc2")))
            .TextMatrix(.Rows - 1, 20) = Trim(CE(RS.Fields("cargo3")))
            .TextMatrix(.Rows - 1, 21) = Trim(CE(RS.Fields("porc3")))
            .TextMatrix(.Rows - 1, 22) = Trim(CE(RS.Fields("cargo4")))
            .TextMatrix(.Rows - 1, 23) = Trim(CE(RS.Fields("cargo4")))
            .TextMatrix(.Rows - 1, 24) = Trim(CE(RS.Fields("cargo5")))
            .TextMatrix(.Rows - 1, 25) = Trim(CE(RS.Fields("porc5")))
            .TextMatrix(.Rows - 1, 26) = Trim(CE(RS.Fields("abono")))
            .TextMatrix(.Rows - 1, 27) = CE(RS.Fields("porcal"))
            .TextMatrix(.Rows - 1, 28) = CE(RS.Fields("colcom"))
            .TextMatrix(.Rows - 1, 29) = CE(RS.Fields("colven"))
            .TextMatrix(.Rows - 1, 30) = CE(RS.Fields("movi"))
            .TextMatrix(.Rows - 1, 31) = CE(RS.Fields("cod_dep"))
            .Rows = .Rows + 1
            RS.MoveNext
        Next
        .Rows = .Rows - 1
      Else
        Call BloqueoDeBotones
      End If
      .Redraw = True
    End With
    RS.Close
    Set RS = Nothing
End Sub

Sub ModoEdicion()
    TxtCuenta.Locked = False
    TxtCuenta.BackColor = ColorHabilitado
    TxtDescripcion.Locked = False
    TxtDescripcion.BackColor = ColorHabilitado
    TxtAuxiliar.Locked = False
    TxtAuxiliar.BackColor = ColorHabilitado
    CboTipCta.Locked = False
    CboTipCta.BackColor = ColorHabilitado
    FrmOption.Enabled = True
    CboMoneda.Locked = False
    CboMoneda.BackColor = ColorHabilitado
    CboEstFin.Locked = False
    CboEstFin.BackColor = ColorHabilitado
    MshHabilitado = False
    MshPlanDeCuentas.BackColor = ColorDeshabilitado
    CboMovi.Locked = False
    CboMovi.BackColor = ColorHabilitado
    txtColCom.Locked = False
    txtColCom.BackColor = ColorHabilitado
    txtColVen.Locked = False
    txtColVen.BackColor = ColorHabilitado
    txtPorcal.Enabled = True
    txtPorcal.BackColor = ColorHabilitado
    txtLin_i_e.Locked = False
    txtEstado.Locked = False
    txtLin_i_e.BackColor = ColorHabilitado
End Sub

Sub ModoNormal()
    TxtCuenta.Locked = True
    TxtCuenta.BackColor = ColorDeshabilitado
    TxtDescripcion.Locked = True
    TxtDescripcion.BackColor = ColorDeshabilitado
    TxtAuxiliar.Locked = True
    TxtAuxiliar.BackColor = ColorDeshabilitado
    CboTipCta.Locked = True
    CboTipCta.BackColor = ColorDeshabilitado
    ChkBancos.Value = 0
    ChkDivision.Value = 0
    ChkCentroCosto.Value = 0
    ChkCtaCte.Value = 0
    ChkGenerada.Value = 0
    Dim i As Integer
    For i = 0 To 4
        TxtCargo1(i).Locked = True
        TxtCargo1(i).BackColor = ColorDeshabilitado
        CboPorcentaje1(i).Locked = True
        CboPorcentaje1(i).BackColor = ColorDeshabilitado
    Next
    TxtAbono.Locked = True
    TxtAbono.BackColor = ColorDeshabilitado
    FrmOption.Enabled = False
    CboMoneda.Locked = True
    CboMoneda.BackColor = ColorDeshabilitado
    CboEstFin.Locked = True
    CboEstFin.BackColor = ColorDeshabilitado
    CboLineasBGF.Locked = True
    CboLineasBGF.BackColor = ColorDeshabilitado
    CboLineasGPN.Locked = True
    CboLineasGPN.BackColor = ColorDeshabilitado
    CboLineasIE.Locked = True
    CboLineasIE.BackColor = ColorDeshabilitado
    MshHabilitado = True
    MshPlanDeCuentas.BackColor = ColorHabilitado
    txtColCom.Locked = True
    txtColCom.BackColor = ColorDeshabilitado
    txtColVen.Locked = True
    txtColVen.BackColor = ColorDeshabilitado
    CboMovi.Locked = True
    CboMovi.BackColor = ColorDeshabilitado
    txtPorcal.Enabled = False
    txtPorcal.BackColor = ColorDeshabilitado
    txtLin_i_e.Locked = True
    txtEstado.Locked = True
    txtLin_i_e.BackColor = ColorDeshabilitado
End Sub
Sub NavegarPorPlanDeCuentas()
    With MshPlanDeCuentas
        If MshHabilitado = True Then Call ImprimirDatosDeMshPlanDeCuentas(.RowSel, .TextMatrix(.RowSel, 1))
        tmpFila = .row
        .SetFocus
    End With
End Sub
Sub ImprimirDatosDeMshPlanDeCuentas(fila As Integer, cuenta As String)
    With MshPlanDeCuentas
        FilaSel = fila
        TxtCuenta = cuenta
        TxtCuenta.Locked = True
        TxtCuenta.BackColor = ColorDeshabilitado
        TxtDescripcion = BuscarDescripcion(Trim(cuenta))
        If Trim(.TextMatrix(fila, 3)) = "D" Then CboTipCta = CboTipCta.List(0)
        If Trim(.TextMatrix(fila, 3)) = "T" Then CboTipCta = CboTipCta.List(1)
        TxtAuxiliar = Val(.TextMatrix(fila, 4))
        
        If Trim(.TextMatrix(fila, 5)) = "E" Then CboMoneda = CboMoneda.List(1)
        If Trim(.TextMatrix(fila, 5)) = "N" Then CboMoneda = CboMoneda.List(0)
        If Trim(.TextMatrix(fila, 5)) = "A" Then CboMoneda = CboMoneda.List(2)
        
        If Trim(.TextMatrix(fila, 6)) = "S" Then
            ChkBancos.Value = 1
        Else
            ChkBancos.Value = 0
        End If
        
        If Trim(.TextMatrix(fila, 7)) = "S" Then
            ChkCentroCosto.Value = 1
        Else
            ChkCentroCosto.Value = 0
        End If
        If Trim(.TextMatrix(fila, 8)) = "S" Then
            ChkCtaCte.Value = 1
        Else
            ChkCtaCte.Value = 0
        End If
        If Trim(.TextMatrix(fila, 9)) = "S" Then
            ChkLote.Value = 1
        Else
            ChkLote.Value = 0
        End If
        If Trim(.TextMatrix(fila, 10)) = "S" Then
            ChkGenerada.Value = 1
        Else
            ChkGenerada.Value = 0
        End If
        
        If Trim(.TextMatrix(fila, 31)) = "S" Then
            ChkDivision.Value = 1
        Else
            ChkDivision.Value = 0
        End If
        
        TxtCargo1(0) = .TextMatrix(fila, 16)
        LblDesCargo1(0) = BuscarDescripcion(Trim(TxtCargo1(0)))
        CboPorcentaje1(0) = .TextMatrix(fila, 17) + "%"
        
        TxtCargo1(1) = .TextMatrix(fila, 18)
        LblDesCargo1(1) = BuscarDescripcion(Trim(TxtCargo1(1)))
        CboPorcentaje1(1) = Format(.TextMatrix(fila, 19), "0") + "%"
        
        TxtCargo1(2) = .TextMatrix(fila, 20)
        LblDesCargo1(2) = BuscarDescripcion(Trim(TxtCargo1(2)))
        CboPorcentaje1(2) = Format(.TextMatrix(fila, 21), "0") + "%"
        
        TxtCargo1(3) = .TextMatrix(fila, 22)
        LblDesCargo1(3) = BuscarDescripcion(Trim(TxtCargo1(3)))
        CboPorcentaje1(3) = Format(.TextMatrix(fila, 23), "0") + "%"
        
        TxtCargo1(4) = .TextMatrix(fila, 24)
        LblDesCargo1(4) = BuscarDescripcion(Trim(TxtCargo1(4)))
        CboPorcentaje1(4) = Format(.TextMatrix(fila, 25), "0") + "%"
                
        TxtAbono = .TextMatrix(fila, 26)
        txtColCom = .TextMatrix(fila, 28)
        txtColVen = .TextMatrix(fila, 29)
        If Trim(.TextMatrix(fila, 30)) = "D" Then CboMovi.ListIndex = 1
        If Trim(.TextMatrix(fila, 30)) = "H" Then CboMovi.ListIndex = 2
        If Trim(.TextMatrix(fila, 30)) = "A" Then CboMovi.ListIndex = 0
        If Trim(.TextMatrix(fila, 30)) = "" Then CboMovi.ListIndex = 0
        
        txtEstado = .TextMatrix(fila, 11)
        txtLin_i_e = .TextMatrix(fila, 25)
        txtPorcal = Right("__" & CStr(.TextMatrix(fila, 27)), 2)
        
        LblDesAbono = BuscarDescripcion(Trim(TxtAbono))
        
        CboEstFin = BuscarEstFin(.TextMatrix(fila, 12))
        
        Call CboEstFin_Click
        
        CboLineasBGF = BuscarLineasBGF(.TextMatrix(fila, 13))
        CboLineasGPN = BuscarLineasGPN(.TextMatrix(fila, 14))
        
        If Left(cuenta, 1) = "6" Or Left(cuenta, 1) = "7" Or Left(cuenta, 1) = "9" Then
            CboLineasIE = BuscarLineasIE(.TextMatrix(fila, 15))
            CboLineasIE.Locked = False
            CboLineasIE.BackColor = ColorHabilitado
        Else
            CboLineasIE.Locked = True
            CboLineasIE.BackColor = ColorDeshabilitado
        End If
        
        
    End With
End Sub

Private Sub CboEstFin_Click()
    If Len(CboEstFin.Text) > 0 Then
        Dim E As String
        E = CboEstFin.Text
        
        Call LlenarCbos
            If E = CboEstFin.List(0) Or E = CboEstFin.List(1) Or E = CboEstFin.List(3) Then '1 ó 2 ó 4
                CboLineasBGF.Locked = False
                CboLineasBGF.BackColor = ColorHabilitado
                CboLineasGPN = Empty
                CboLineasGPN.Locked = True
                CboLineasGPN.BackColor = ColorDeshabilitado
                'CboLineasIE = Empty
                'CboLineasIE.Locked = True
                'CboLineasIE.BackColor = ColorDeshabilitado
            End If
            If E = CboEstFin.List(2) Then '3
                CboLineasGPN.Locked = False
                CboLineasGPN.BackColor = ColorHabilitado
                CboLineasBGF = Empty
                CboLineasBGF.Locked = True
                CboLineasBGF.BackColor = ColorDeshabilitado
                'CboLineasIE = Empty
                'CboLineasIE.Locked = True
                'CboLineasIE.BackColor = ColorDeshabilitado
            End If
            If E = CboEstFin.List(3) Then  '4
                CboLineasBGF.Locked = False
                CboLineasBGF.BackColor = ColorHabilitado
                CboLineasGPN.Locked = False
                CboLineasGPN.BackColor = ColorHabilitado
                CboLineasIE = Empty
                'CboLineasIE.Locked = True
                'CboLineasIE.BackColor = ColorDeshabilitado
            End If
            If E = CboEstFin.List(4) Then
                'CboLineasIE.Locked = False
                'CboLineasIE.BackColor = ColorHabilitado
                CboLineasBGF = Empty
                CboLineasBGF.Locked = True
                CboLineasBGF.BackColor = ColorDeshabilitado
                CboLineasGPN = Empty
                CboLineasGPN.Locked = True
                CboLineasGPN.BackColor = ColorDeshabilitado
            End If
        
    End If
End Sub

Private Sub CboEstFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim E As String
        E = CboEstFin.Text
            If E = CboEstFin.List(0) Or E = CboEstFin.List(1) Or E = CboEstFin.List(3) Then CboLineasBGF.SetFocus '1 ó 2 ó 4
            If E = CboEstFin.List(2) Then CboLineasGPN.SetFocus '3
            If E = CboEstFin.List(4) Then CboLineasIE.SetFocus '5
    End If
End Sub

Private Sub CboLineasBGF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CboLineasGPN.SetFocus
End Sub

Private Sub CboLineasGPN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CboLineasIE.SetFocus
End Sub

Private Sub CboLineasIE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdGrabar.Enabled = True Then Call cmdGrabar_Click
End Sub

Private Sub cboOrdenar_Click()
    Call OrdenarMshPlanDeCuentas
End Sub
Sub OrdenarMshPlanDeCuentas() 'SE CAMBIO 3/02/01
    If cboOrdenar.List(0) = cboOrdenar.Text Then
        MshPlanDeCuentas.col = 1
        MshPlanDeCuentas.Sort = 3
        MshPlanDeCuentas.Refresh
    End If
    If cboOrdenar.List(1) = cboOrdenar.Text Then
        MshPlanDeCuentas.col = 2
        MshPlanDeCuentas.Sort = 5
        MshPlanDeCuentas.Refresh
    End If
End Sub

Private Sub CboOrdenar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call OrdenarMshPlanDeCuentas
End Sub

Private Sub CboPorcentaje1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim sum As Integer
    Dim cont As Integer
  
  If KeyAscii = 13 Then
        For cont = 0 To Index
            sum = sum + Val(CboPorcentaje1(cont).Text)
        Next cont
        
        If sum < 100 Then
            If Index = 4 Then
              CboEstFin.SetFocus
            Else
              TxtCargo1(Index + 1).Enabled = True
              TxtCargo1(Index + 1).Locked = False
              TxtCargo1(Index + 1).BackColor = ColorHabilitado
              TxtCargo1(Index + 1).SetFocus
              
              CboPorcentaje1(Index + 1).Locked = False
              CboPorcentaje1(Index + 1).BackColor = ColorHabilitado
            End If
          
        Else
          
            If sum = 100 Then
              Call CalcularPorcentaje(Index)
            Else
              MsgBox "Verifique sus porcentajes", vbOKOnly
              CboPorcentaje1(Index).SetFocus
            End If
        End If
        
  End If
End Sub

Private Sub CboTipCta_Click()
    If CboTipCta.Text = CboTipCta.List(0) Then 'si es dato
        Call ModoEdicionDato
    End If
    If CboTipCta.Text = CboTipCta.List(1) Then ' Si Es Título
        Call ModoEdicionTitulo
    End If
End Sub
Private Sub CboTipCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CboTipCta = CboTipCta.List(1) Then ' Si Es Título
            If cmdGrabar.Enabled = True Then Call cmdGrabar_Click
        Else
            TxtDescripcion.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call Limpia_Valores
    Call ModoNormal
    Call BotonNormal
    lblMensaje = Empty
    SSTab1.Tab = 0
    FilaSel = 0
End Sub

Private Sub CmdEliminar_Click()
    If TxtCuenta <> Empty Then
        If MsgBox("Está Seguro De Eliminar La Cuenta " + TxtCuenta + "  (s/n) ", vbInformation + vbYesNo, "Plan De Cuentas") = vbYes Then
            If La_Cuenta_Tiene_Mov(TxtCuenta.Text) = False Then
              Call BorrarRegistro
              MshPlanDeCuentas.Clear
              
              Call LlenarMshPlanDeCuentas
              Call Limpia_Valores
              Call ModoNormal
              Call BotonNormal
              lblMensaje = Empty
              
                If tmpFila > 0 And tmpFila <= MshPlanDeCuentas.Rows Then
                    If tmpFila = MshPlanDeCuentas.Rows Then tmpFila = tmpFila - 1
                    MshPlanDeCuentas.row = tmpFila
                    MshPlanDeCuentas.col = 0
                    MshPlanDeCuentas.SetFocus
                    SendKeys "{HOME}+{END}"
                End If
              
            Else
              MsgBox "No se puede eliminar La Cuenta Nª " & CStr(TxtCuenta) & " Porque Tiene Movimiento ", vbInformation, Caption
            End If
        End If
    Else
        MsgBox "Seleccione El Registro A Eliminar", vbInformation, "Plan De Cuentas"
        MshPlanDeCuentas.SetFocus
    End If
End Sub
Private Sub CmdFind_Click(Index As Integer)
    If Index = 0 Then Call txtCuenta_KeyDown(112, 0)
    If Index = 1 Then Call txtAuxiliar_KeyDown(113, 0)
End Sub

Private Sub cmdGrabar_Click()
    Dim fila As Integer

    If ValidarData = True Then
        Call GrabarData
        MshPlanDeCuentas.Clear
        
        Call LlenarMshPlanDeCuentas
        Call Limpia_Valores
        Call ModoNormal
        Call BotonNormal
        lblMensaje = Empty
        SSTab1.Tab = 0
        
        fila = BuscarCriterio(tmpCuenta)
        If fila > 0 Then
            MshPlanDeCuentas.row = fila
            MshPlanDeCuentas.col = 0
            MshPlanDeCuentas.SetFocus
            SendKeys "{HOME}+{END}"
        End If
    End If
    
End Sub

Sub GrabarData()
On Error GoTo ErrSave
    Dim Cn As ADODB.Connection
    Set Cn = New ADODB.Connection
    Dim movi As String
    Dim SQL As String
    
            Dim TxtLinea  As String
            '****************************************************
            If CboLineasBGF <> Empty Then TxtLinea = Mid(Trim(CboLineasBGF), 1, 2)
            If CboLineasGPN <> Empty Then TxtLinea = Mid(Trim(CboLineasGPN), 1, 2)
            If CboLineasIE <> Empty Then TxtLinea = Mid(Trim(CboLineasIE), 1, 2)
            '****************************************************
    
    If CboMovi.ListIndex = 0 Then movi = ""
    If CboMovi.ListIndex = 1 Then movi = "D"
    If CboMovi.ListIndex = 2 Then movi = "H"
    
    Select Case lblMensaje.Caption
        Case "Nuevo" 'nuevo ok
            
            If Left(Trim(CboTipCta.Text), 1) = "T" Then
                SQL = "INSERT INTO CNMAYOR (CUENTA,DESCRIP,TIPO) "
                SQL = SQL & "VALUES('" & Trim(TxtCuenta) + "','" & Trim(TxtDescripcion) & "','" & Left(Trim(CboTipCta.Text), 1) & "')"
            End If
            
            If Left(Trim(CboTipCta.Text), 1) = "D" Then
                SQL = "INSERT INTO CNMAYOR (movi,colcom,colven,estado, porcal, cod_fun, CUENTA,DESCRIP,TIPO,AUXILIAR,MONEDA,BANCOS,CEN_COS,CTA_CTE,GENERADA,EST_FIN,LIN_BGF,LIN_GPN,LIN_I_E,CARGO1,PORC1,CARGO2,PORC2,CARGO3,PORC3,CARGO4,PORC4,CARGO5,PORC5,ABONO,USUARIO,cod_dep)"
                SQL = SQL + " VALUES('" & Trim(movi) & "','" + Trim(txtColCom) + "','" + Trim(txtColVen.Text) + "','" + Trim(txtEstado) + "','" + Trim(Val(Replace(txtPorcal.Text, "_", ""))) + "','" + IIf(ChkLote.Value = 1, "S", "N") + "','" + Trim(TxtCuenta) + "','"
                SQL = SQL + Trim(TxtDescripcion) + "','" + Left(Trim(CboTipCta.Text), 1) + "','"
                SQL = SQL + IIf((Trim(TxtAuxiliar) = Empty), "0", TxtAuxiliar) + "','"
                SQL = SQL + IIf(CboMoneda.Text = Empty, "0", Left(Trim(CboMoneda.Text), 1)) + "','"
                SQL = SQL + IIf(ChkBancos.Value = 1, "S", "N") + "','"
                SQL = SQL + IIf(ChkCentroCosto.Value = 1, "S", "N") + "','"
                SQL = SQL + IIf(ChkCtaCte.Value = 1, "S", "N") + "','"
                SQL = SQL + IIf(ChkGenerada.Value = 1, "S", "N") + "','"
                SQL = SQL + Trim(Left(CboEstFin.Text, InStr(1, CboEstFin.Text, " "))) + "','"
                
                SQL = SQL + IIf(CboLineasBGF = Empty, Empty, Trim(Left(CboLineasBGF.Text, InStr(1, CboLineasBGF.Text, " ")))) + "','"
                SQL = SQL + IIf(CboLineasGPN = Empty, Empty, Trim(Left(CboLineasGPN.Text, InStr(1, CboLineasGPN.Text, " ")))) + "','"
                SQL = SQL + IIf(CboLineasIE = Empty, Empty, Trim(Left(CboLineasIE.Text, InStr(1, CboLineasIE.Text, " ")))) + "','"
                
                SQL = SQL + IIf(TxtCargo1(0).Text = Empty, Empty, Trim(TxtCargo1(0).Text)) + "',"
                If Trim(CboPorcentaje1(0).Text) <> "0%" And Trim(CboPorcentaje1(0).Text) <> "" Then
                  SQL = SQL & Val(CboPorcentaje1(0).Text) & ",'"
                Else
                  SQL = SQL & 0 & ",'"
                End If
                SQL = SQL & IIf(TxtCargo1(1).Text = Empty, "", Trim(TxtCargo1(1).Text)) + "',"
                If Trim(CboPorcentaje1(1).Text) <> "0%" And Trim(CboPorcentaje1(1).Text) <> "" Then
                  SQL = SQL & Val(CboPorcentaje1(1).Text) & ",'"
                Else
                  SQL = SQL & 0 & ",'"
                End If
                SQL = SQL + IIf(TxtCargo1(2).Text = Empty, "", Trim(TxtCargo1(2).Text)) + "',"
                If Trim(CboPorcentaje1(2).Text) <> "0%" And Trim(CboPorcentaje1(2).Text) <> "" Then
                  SQL = SQL & Val(CboPorcentaje1(2).Text) & ",'"
                Else
                  SQL = SQL & 0 & ",'"
                End If
                SQL = SQL + IIf(TxtCargo1(3).Text = Empty, "", Trim(TxtCargo1(3).Text)) + "',"
                 If Trim(CboPorcentaje1(3).Text) <> "0%" And Trim(CboPorcentaje1(3).Text) <> "" Then
                  SQL = SQL & Val(CboPorcentaje1(3).Text) & ",'"
                Else
                  SQL = SQL & 0 & ",'"
                End If
                 SQL = SQL + IIf(TxtCargo1(4).Text = Empty, "", Trim(TxtCargo1(4).Text)) + "',"
                 If Trim(CboPorcentaje1(4).Text) <> "0%" And Trim(CboPorcentaje1(4).Text) <> "" Then
                  SQL = SQL & Val(CboPorcentaje1(4).Text) & ",'"
                Else
                  SQL = SQL & 0 & ",'"
                End If
                 
                 SQL = SQL & IIf(TxtAbono.Text = Empty, "", Trim(TxtAbono.Text)) + "','"
                 SQL = SQL & Trim(UsuarioActivo) & "', '"
                 SQL = SQL & IIf(ChkDivision.Value = 1, "S", "N") & "')"
            End If
        
        Case "Modificar" 'modificar ok
            
            If Left(Trim(CboTipCta.Text), 1) = "T" Then
                 SQL = "Update CNMAYOR SET DESCRIP='" + Trim(TxtDescripcion) + "'  Where  CUENTA='" + Trim(TxtCuenta) + "' "
            End If
            
            If Left(Trim(CboTipCta.Text), 1) = "D" Then
                SQL = "Update CNMAYOR SET Cuenta ='" + Trim(TxtCuenta) + "',"
                
                SQL = SQL + " movi='" + Trim(movi) + "',"
                SQL = SQL + " colcom='" + Trim(txtColCom) + "',"
                SQL = SQL + " colven='" + Trim(txtColVen) + "',"
                If Trim(txtPorcal.Text) <> "0%" And Trim(txtPorcal.Text) <> "" Then
                  SQL = SQL & "porcal =" & Val(Replace(txtPorcal.Text, "_", "")) & ","
                Else
                  SQL = SQL & "porcal =" & 0 & ","
                End If
                
                SQL = SQL + " estado='" + IIf((Trim(txtEstado) = Empty), "0", txtEstado) + "',"
                SQL = SQL + " cod_fun='" + IIf(Me.ChkLote.Value = 1, "S", "N") + "',"
                SQL = SQL + "DESCRIP ='" + Trim(TxtDescripcion) + "',"
                SQL = SQL + "TIPO ='" + Left(Trim(CboTipCta.Text), 1) + "',"
                SQL = SQL + "Auxiliar ='" + IIf((Trim(TxtAuxiliar) = Empty), "0", TxtAuxiliar) + "',"
                SQL = SQL + "Moneda ='" + IIf(CboMoneda.Text = Empty, "0", Left(Trim(CboMoneda.Text), 1)) + "',"
                SQL = SQL + "BANCOS ='" + IIf(ChkBancos.Value = 1, "S", "N") + "',"
                SQL = SQL + "CEN_COS ='" + IIf(ChkCentroCosto.Value = 1, "S", "N") + "',"
                SQL = SQL + "CTA_CTE ='" + IIf(ChkCtaCte.Value = 1, "S", "N") + "',"
                SQL = SQL + "Generada ='" + IIf(ChkGenerada.Value = 1, "S", "N") + "',"
                SQL = SQL + "EST_FIN ='" + Trim(Left(CboEstFin.Text, InStr(1, CboEstFin.Text, " "))) + "',"
                
                SQL = SQL + "LIN_BGF ='" + IIf(CboLineasBGF = Empty, Empty, Trim(Left(CboLineasBGF.Text, InStr(1, CboLineasBGF.Text, " ")))) + "',"
                SQL = SQL + "LIN_GPN ='" + IIf(CboLineasGPN = Empty, Empty, Trim(Left(CboLineasGPN.Text, InStr(1, CboLineasGPN.Text, " ")))) + "',"
                SQL = SQL + "LIN_I_E ='" + IIf(CboLineasIE = Empty, Empty, Trim(Left(CboLineasIE.Text, InStr(1, CboLineasIE.Text, " ")))) + "',"
                
                SQL = SQL + "CARGO1 ='" + IIf(TxtCargo1(0).Text = Empty, Empty, Trim(TxtCargo1(0).Text)) + "',"
                If Trim(CboPorcentaje1(0).Text) <> "0%" And Trim(CboPorcentaje1(0).Text) <> "" Then
                  SQL = SQL & "PORC1 =" & Val(CboPorcentaje1(0).Text) & ","
                Else
                  SQL = SQL & "PORC1 =" & 0 & ","
                End If
                
                SQL = SQL + "CARGO2 ='" + IIf(TxtCargo1(1).Text = Empty, Empty, Trim(TxtCargo1(1).Text)) + "',"
                If Trim(CboPorcentaje1(1).Text) <> "0%" And Trim(CboPorcentaje1(1).Text) <> "" Then
                  SQL = SQL & "PORC2 =" & Val(CboPorcentaje1(1).Text) & ","
                Else
                  SQL = SQL & "PORC2 =" & 0 & ","
                End If
                
                SQL = SQL + "CARGO3 ='" + IIf(TxtCargo1(2).Text = Empty, Empty, Trim(TxtCargo1(2).Text)) + "',"
                If Trim(CboPorcentaje1(2).Text) <> "0%" And Trim(CboPorcentaje1(2).Text) <> "" Then
                  SQL = SQL & "PORC3 =" & Val(CboPorcentaje1(2).Text) & ","
                Else
                  SQL = SQL & "PORC3 =" & 0 & ","
                End If
                                
                SQL = SQL + "CARGO4 ='" + IIf(TxtCargo1(3).Text = Empty, Empty, Trim(TxtCargo1(3).Text)) + "',"
                If Trim(CboPorcentaje1(3).Text) <> "0%" And Trim(CboPorcentaje1(3).Text) <> "" Then
                  SQL = SQL & "PORC4 =" & Val(CboPorcentaje1(3).Text) & ","
                Else
                  SQL = SQL & "PORC4 =" & 0 & ","
                End If
                
                SQL = SQL + "CARGO5 ='" + IIf(TxtCargo1(4).Text = Empty, Empty, Trim(TxtCargo1(4).Text)) + "',"
                If Trim(CboPorcentaje1(4).Text) <> "0%" And Trim(CboPorcentaje1(4).Text) <> "" Then
                  SQL = SQL & "PORC5 =" & Val(CboPorcentaje1(4).Text) & ","
                Else
                  SQL = SQL & "PORC5 =" & 0 & ","
                End If
                
                SQL = SQL + "Abono ='" + IIf(TxtAbono.Text = Empty, Empty, Trim(TxtAbono.Text)) + "', "
                
                 SQL = SQL + "Cod_dep='" + IIf(ChkDivision.Value = 1, "S", "N") & "' "
                
                SQL = SQL + " Where  CUENTA ='" + Trim(TxtCuenta) + "'"
                
            End If
    End Select
    tmpCuenta = Trim(TxtCuenta)
    ADOConexionEmp.BeginTrans
    ADOConexionEmp.Execute (SQL)
    ADOConexionEmp.CommitTrans
    Exit Sub
ErrSave:
    MsgBox "Ha ocurrido un error al momento de grabar " & Chr(13) & Err.Description, vbCritical, "Error de datos"
    ADOConexionEmp.RollbackTrans
End Sub

Private Function ValidarData() As Boolean
    If Left(Trim(CboTipCta.Text), 1) = "T" Then
        ValidarData = True
        Exit Function
    End If
    
    Dim fila As Integer
    
    If TxtAuxiliar = Empty Then
        MsgBox "El Auxiliar esta en Blanco", vbExclamation, "Plan De Cuenta"
        ValidarData = False
        TxtAuxiliar.SetFocus
        Exit Function
    End If
    
    If txtPorcal = Empty Then
        MsgBox "El Item Porcentaje de Calculo esta en Blanco", vbExclamation, "Plan De Cuenta"
        ValidarData = False
        txtPorcal.SetFocus
        Exit Function
    End If
    
    If Val(txtPorcal) < 0 Or Val(txtPorcal) >= 100 Then
        If txtPorcal <> "__" Then
            MsgBox "El Item Porcentaje debe de ser mayor o igual que 0 y menor que 100 ", vbExclamation, "Plan De Cuenta"
            ValidarData = False
            Exit Function
        Else
            txtPorcal = "00"
        End If
    End If
    
    If Not IsNumeric(txtEstado) Then
        MsgBox "El Item Tipo de gasto esta en Blanco (valor por defecto 0 )", vbExclamation, "Plan De Cuenta"
        txtEstado = Val(txtEstado)
        ValidarData = False
        txtEstado.SetFocus
        Exit Function
        
    End If
    
    If txtEstado = Empty Then
        MsgBox "El Item Tipo de Gasto esta en Blanco", vbExclamation, "Plan De Cuenta"
        ValidarData = False
        txtEstado.SetFocus
        Exit Function
    End If
        
    If TxtCuenta = Empty Then
        MsgBox "El Item Cuenta Está En Blanco", vbExclamation, "Plan De Cuenta"
        ValidarData = False
        TxtCuenta.SetFocus
        Exit Function
    Else
        fila = BuscarCuenta(Trim(TxtCuenta))
        If fila > 0 Then
            If lblMensaje = "Nuevo" Then
                MsgBox "El Registro Ya Existe.Pulse F1 Para Eligir Una Cuenta", vbExclamation, "Plan De Cuenta"
                ValidarData = False
                TxtCuenta.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            End If
        Else
            If lblMensaje = "Modificar" Then
                MsgBox "El Registro No Existe.Pulse F1 Para Eligir Una Cuenta", vbExclamation, "Plan De Cuenta"
                ValidarData = False
                TxtCuenta.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            End If
        End If
    End If
    If TxtDescripcion = Empty Then
        MsgBox "El Item Descripción Está En Blanco", vbExclamation, "Plan De Cuentas"
        ValidarData = False
        TxtDescripcion.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Function
    End If
    If CboTipCta = Empty Then
        MsgBox "El Item Tipo De Cuenta Está En Blanco", vbExclamation, "Plan De Cuenta"
        ValidarData = False
        CboTipCta.SetFocus
        Exit Function
    End If
    If CboTipCta = CboTipCta.List(0) Then
            If CboMoneda.Text = Empty Then
                MsgBox "El Item Moneda Está En Blanco", vbExclamation, "Plan De Cuentas"
                ValidarData = False
                CboMoneda.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            Else
                If ValidarMoneda = False Then
                    MsgBox "El Dato Ingresado En El Item Moneda No Pertenece A un Elemento de la Lista Desplegable", vbExclamation, "Plan De Cuentas"
                    ValidarData = False
                    CboMoneda.SetFocus
                    SendKeys "{HOME}+{END}"
                    Exit Function
                End If
            End If
            If ChkCtaCte.Value = 1 Then
                If TxtAuxiliar = Empty Then
                    MsgBox "Ingrese El Código Auxiliar Que Está En Blanco..!!", vbExclamation, "Plan De Cuenta"
                    ValidarData = False
                    TxtAuxiliar.SetFocus
                    Exit Function
                Else
                    If BuscarAuxiliar(TxtAuxiliar) = Empty Then
                       MsgBox "El Código De Auxiliar No Existe.Pulse F2 Para Elejir Un Auxiliar", vbExclamation, "Plan De Cuentas"
                       TxtAuxiliar.SetFocus
                       SendKeys "{HOME}+{END}"
                       ValidarData = False
                       Exit Function
                    End If
                End If
            End If
            
            If Not (CboLineasBGF.Locked) And CboLineasBGF = Empty And CboLineasBGF.ListCount > 0 Then
                MsgBox "Seleccione un dato de lineas de BGF de la lista desplegable", vbExclamation, "Plan De Cuentas"
                ValidarData = False
                CboEstFin.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            End If
            
            If Not (CboLineasGPN.Locked) And CboLineasGPN = Empty And CboLineasGPN.ListCount > 0 Then
                MsgBox "Seleccione un dato de lineas de GPN de la lista desplegable", vbExclamation, "Plan De Cuentas"
                ValidarData = False
                CboEstFin.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            End If
            
            If Not (CboLineasIE.Locked) And CboLineasIE = Empty And CboLineasIE.ListCount > 0 Then
                MsgBox "Seleccione un dato de lineas de IE de la lista desplegable", vbExclamation, "Plan De Cuentas"
                ValidarData = False
                CboEstFin.SetFocus
                SendKeys "{HOME}+{END}"
                Exit Function
            End If
            
            If ChkGenerada.Value = 0 Then

                    Dim i As Integer
                    For i = 0 To 4
                        If CboPorcentaje1(i).Locked = False And CboPorcentaje1(i).Text = Empty Then
                            MsgBox "Ingrese El Porcentaje Relacionado Al Cargo N° " + STR(i) + "  Que Está En Blanco..!!", vbExclamation, "Plan De Cuentas"
                            ValidarData = False
                            Exit Function
                        End If
                    Next
                    

                        For i = 0 To 4
                        Dim F As Integer
                        If Trim(TxtCargo1(i).Text) <> "" Then 'aquí mofificado
                          F = BuscarCuenta(Trim(TxtCargo1(i).Text))
                            If F <= 0 Then
                                MsgBox "La Cuenta  Nº " & i + 1 & " no Existe.Pulse F1 Para Eligir Una Cuenta", vbExclamation, "Plan De Cuenta"
                                ValidarData = False
                                TxtCargo1(i).SetFocus
                                SendKeys "{HOME}+{END}"
                                Exit Function
                            End If
                        End If 'aquí modificado
                    Next
                    '-------------------------------------------------

                    Dim s As Integer
                    For i = 0 To 4 'aquí modificado
                        If Trim(TxtCargo1(i).Text) <> "" Then 'aquí mofificado
                          If TxtAbono = Empty Then
                              MsgBox "El Item Abono Está En Blanco", vbExclamation, "Plan De Cuentas"
                              ValidarData = False
                              TxtAbono.SetFocus
                              SendKeys "{HOME}+{END}"
                              Exit Function
                          Else
                              fila = BuscarCuenta(Trim(TxtCuenta))
                              If fila <= 0 And lblMensaje <> "Nuevo" Then
                                  MsgBox "La Cuenta No Existe.Pulse F1 Para Eligir Una Cuenta", vbExclamation, "Plan De Cuenta"
                                  ValidarData = False
                                  TxtCuenta.SetFocus
                                  SendKeys "{HOME}+{END}"
                                  Exit Function
                              End If
                          End If
                        End If
                   Next i 'aquí modificado
                    If CboEstFin = Empty Then
                        MsgBox "El Item Estado Financiero Está En Blanco", vbExclamation, "Plan De Cuentas"
                        ValidarData = False
                        CboTipCta.SetFocus
                        SendKeys "{HOME}+{END}"
                        Exit Function
                    Else
                        If ValidarCboEstFin = False Then
                            MsgBox "El Dato Ingresado En El Item Estado Fianaciero No Pertenece A Un Dato De La Lista Desplegable", vbExclamation, "Plan De Cuentas"
                            ValidarData = False
                            CboEstFin.SetFocus
                            SendKeys "{HOME}+{END}"
                            Exit Function
                        End If
                    End If
            End If
    End If
    ValidarData = True
End Function
Private Function ValidarCboEstFin() As Boolean
    Dim i As Integer
    For i = 0 To CboEstFin.ListCount - 1
        If CboEstFin.Text = CboEstFin.List(i) Then
            ValidarCboEstFin = True
            Exit Function
        End If
    ValidarCboEstFin = False
    Next
End Function
Private Function ValidarMoneda() As Boolean
    Dim i As Integer
    For i = 0 To CboMoneda.ListCount - 1
        If CboMoneda.Text = CboMoneda.List(i) Then
            ValidarMoneda = True
            Exit Function
        End If
    Next
    ValidarMoneda = False
End Function

Private Sub CmdModificar_Click()
    If FilaSel > 0 Then
        With MshPlanDeCuentas
            If .TextMatrix(FilaSel, 3) = "T" Then Call ModoEdicionTitulo
            If .TextMatrix(FilaSel, 3) = "D" Then Call ModoEdicionDato
            TxtDescripcion = .TextMatrix(FilaSel, 2)
        End With
        Call BotonEdicion
        
        TxtDescripcion.SetFocus
        SendKeys "{HOME}+{END}"
    Else
        Call BotonEdicion
        Call ModoEdicion
        TxtCuenta.SetFocus
    End If
    lblMensaje = "Modificar"
    SSTab1.Tab = 1
End Sub
Sub ModoEdicionDato()
    Call ModoEdicion
    Call ModoEdicionDatoPorcentaje
    txtPorcal.BackColor = ColorHabilitado
    txtPorcal.Enabled = True
    txtEstado.Locked = False
    txtEstado.BackColor = ColorHabilitado
    txtLin_i_e.Locked = False
    txtLin_i_e.BackColor = ColorHabilitado
    TxtAbono.Locked = False
    TxtAbono.BackColor = ColorHabilitado
    FrameCbos.Enabled = True
    Frame4.Enabled = True
    TxtCargo1(0).Locked = False
    TxtCargo1(0).BackColor = ColorHabilitado
    CboPorcentaje1(0).Locked = False
    CboPorcentaje1(0).BackColor = ColorHabilitado
End Sub
Sub ModoEdicionDatoPorcentaje()
    Dim i As Integer
    For i = 0 To 4
        If TxtCargo1(i).Text <> Empty Then
            TxtCargo1(i).Locked = False
            TxtCargo1(i).BackColor = ColorHabilitado
            CboPorcentaje1(i).Locked = False
            CboPorcentaje1(i).BackColor = ColorHabilitado
        End If
    Next
End Sub
Private Sub cmdNuevo_Click()
    Call Limpia_Valores
    Call ModoEdicion
    Call BotonEdicion
    lblMensaje = "Nuevo"
    TxtCuenta.SetFocus
    Call InicarPorcentajes
    SSTab1.Tab = 1
End Sub
Sub BotonEdicion()
    cmdNuevo.Enabled = False
    CmdModificar.Enabled = False
    cmdCancelar.Enabled = True
    cmdGrabar.Enabled = True
    CmdEliminar.Enabled = False
    cmdSalir.Enabled = False
End Sub

Sub BotonNormal()
    cmdNuevo.Enabled = True
    CmdModificar.Enabled = True
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    CmdEliminar.Enabled = True
    cmdSalir.Enabled = True
End Sub

Sub Limpia_Valores()
    TxtCuenta = Empty
    TxtDescripcion = Empty
    CboTipCta = Empty
    TxtAuxiliar = Empty
    ChkBancos.Value = 0
    ChkCentroCosto.Value = 0
    ChkDivision.Value = 0
    ChkLote.Value = 0
    txtPorcal = "__"
    txtLin_i_e = Empty
    txtEstado = Empty
    ChkCtaCte.Value = 0
    ChkGenerada.Value = 0
    Dim i As Integer
    For i = 0 To 4
        TxtCargo1(i) = Empty
        LblDesCargo1(i) = Empty
        CboPorcentaje1(i) = Empty
    Next
    TxtAbono = Empty
    LblDesAbono = Empty
    CboMoneda = Empty
    CboEstFin = Empty
    CboLineasBGF = Empty
    CboLineasGPN = Empty
    CboLineasIE = Empty
    CboMovi.ListIndex = 2
    txtColCom = Empty
    txtColVen = Empty
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ChkGenerada_Click()
    Dim i As Integer
    If ChkGenerada.Value = 0 Then
        If lblMensaje = "Modificar" Then
            FrmOption.Enabled = True
            Call ModoEdicionDatoPorcentaje
        End If
        If lblMensaje = "Nuevo" Then
            FrmOption.Enabled = True
            Frame4.Enabled = True
            TxtCargo1(0).Locked = False
            TxtCargo1(0).BackColor = ColorHabilitado
            CboPorcentaje1(0).Locked = False
            CboPorcentaje1(0).BackColor = ColorHabilitado
        End If
        TxtAbono.Locked = False
        TxtAbono.BackColor = ColorHabilitado
        If TxtCargo1(0).Locked = False Then TxtCargo1(0).SetFocus
    Else
        Call ModoEdicionTituloPorcentajes
        TxtAbono.Locked = True
        TxtAbono.BackColor = ColorDeshabilitado
        If CboEstFin.Enabled Then CboEstFin.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub CmdVistaPreliminar_Click()
    Dim oReporte As New clsReporte
    MDIPrincipal.MousePointer = vbHourglass
    oReporte.TipoBD = BD_EMPRESA
    oReporte.empresa = NombreEmpresa
    oReporte.Titulo = "MAESTRO DE PLAN DE CUENTAS"
    oReporte.Reporte = "Rep_PlanDeCuentas.rpt"
    oReporte.Query = "select CUENTA, DESCRIP, TIPO, AUXILIAR, MONEDA, BANCOS, CEN_COS, CTA_CTE, ESTADO, GENERADA, EST_FIN, LIN_BGF, LIN_GPN, LIN_I_E, CARGO1, PORC1, CARGO2, PORC2, CARGO3, PORC3, CARGO4, PORC4, CARGO5, PORC5, ABONO, COD_FUN, PORCAL, USUARIO, COD_DEP from cnmayor order by 1"
    oReporte.ImprimeReporte
    Set oReporte = Nothing
    MDIPrincipal.MousePointer = vbNormal
End Sub

Private Sub Form_Activate()
    If cmdNuevo.Enabled = True Then cmdSalir.SetFocus
    Me.Top = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
    End If
End Sub

Private Sub Form_Load()
    Set oConsulta = New clsConsultas
    Left = 0
    Top = 0
    Me.Top = 0
    Call LlenarMshPlanDeCuentas
    Call LlenarCboEEFF
    Call LlenarCbos
    Call ModoNormal
    Call BotonNormal
    lblMensaje = Empty
    SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lblMensaje <> Empty Then
        Dim M As String
        M = MsgBox("¿ Desea Guardar los Cambios Realizados ?", vbInformation + vbYesNoCancel, "Plan De Cuentas")
       Select Case M
            Case vbYes
                If ValidarData = True Then
                    Call GrabarData
                    Cancel = 0
                    MDIPrincipal.Enabled = True
                Else
                    Cancel = 1
                    MDIPrincipal.Enabled = False
                End If
            Case vbNo
                Cancel = 0
                MDIPrincipal.Enabled = True
            Case vbCancel
                Cancel = 1
        End Select
    Else
      MDIPrincipal.Enabled = True
    End If
End Sub

Private Sub MshPlanDeCuentas_Click()
    Call NavegarPorPlanDeCuentas
End Sub

Private Sub TxtctaCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 112 And lblMensaje <> Empty Then
        'Llamamos a la ventana de ayuda de Cuentas De PLan De Cuentas
        oConsulta.Caso = 2
        oConsulta.Formulario = "frmplancuentas"
        oConsulta.Control = "TxtCuenta"
        FrmConsultas.Caption = FrmConsultas.Caption + " de Cuenta Mayor"
        FrmConsultas.Show

    End If
End Sub
   
Private Sub MshPlanDeCuentas_DblClick()
  Call CmdModificar_Click
End Sub

Private Sub MshPlanDeCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call NavegarPorPlanDeCuentas
End Sub

Private Sub OptCuenta_Click()
  OptDescripcion.Value = False
  TxtCriterio.SetFocus
  SendKeys "{HOME}+{END}"
End Sub

Private Sub OptDescripcion_Click()
  OptCuenta.Value = False
  TxtCriterio.SetFocus
  SendKeys "{HOME}+{END}"
End Sub


Private Sub TxtAbono_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And TxtAbono.Locked = False Then
        oConsulta.Caso = 2
        oConsulta.Formulario = "frmplancuentas"
        oConsulta.Control = "TxtAbono"
        FrmConsultas.Caption = FrmConsultas.Caption + " de Plan De Cuentas"
        FrmConsultas.Show
    End If
End Sub

Private Sub TxtAbono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboEstFin.SetFocus
    Else
        KeyAscii = OnlyNumbers(KeyAscii)
        If IsNumeric(KeyAscii) = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtAuxiliar_Change()
    If Len(TxtAuxiliar) > 0 Then
        ChkCtaCte.Value = 1
    Else
        ChkCtaCte.Value = 0
    End If
End Sub

Private Sub txtAuxiliar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 And TxtAuxiliar.Locked = False Then
        oConsulta.Caso = 9
        oConsulta.Formulario = "frmplancuentas"
        oConsulta.Control = "TxtAuxiliar"
        FrmConsultas.Caption = FrmConsultas.Caption + " de Auxiliares"
        FrmConsultas.Show
        
    End If
End Sub

Private Sub TxtAuxiliar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If FrmOption.Enabled = True Then
            ChkBancos.SetFocus
        End If
    Else
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtCargo1_Change(Index As Integer)

If Trim(TxtCargo1(Index).Text) = "" Then
      LblDesCargo1(Index).Caption = ""
      If Index > 0 Then
        LblDesCargo1(Index).BackColor = ColorDeshabilitado
      End If
      
      CboPorcentaje1(Index).Text = ""
      
      If Index > 0 Then
        CboPorcentaje1(Index).Text = "0%"
        CboPorcentaje1(Index).BackColor = ColorDeshabilitado
        TxtCargo1(Index).BackColor = ColorDeshabilitado
        'TxtCargo1(Index - 1).SetFocus
      End If
End If

End Sub

Private Sub TxtCargo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And TxtCargo1.Item(Index).Locked = False Then
        oConsulta.Formulario = "frmplancuentas"
        Select Case Index
            Case 0
                oConsulta.Control = "TxtCargo1(0)"
            Case 1
                oConsulta.Control = "TxtCargo1(1)"
            Case 2
                oConsulta.Control = "TxtCargo1(2)"
            Case 3
                oConsulta.Control = "TxtCargo1(3)"
            Case 4
                oConsulta.Control = "TxtCargo1(4)"
        End Select
        oConsulta.Caso = 2
        FrmConsultas.Caption = FrmConsultas.Caption + " De Plan De Cuentas"
        FrmConsultas.Show
        'Set oFrmConsulta = Nothing
    End If
    If (KeyCode = 8 Or KeyCode = 46) And ChkGenerada.Value = 0 And TxtCargo1(Index).Locked = False Then LblDesCargo1(Index) = Empty
End Sub

Private Sub TxtCargo1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboPorcentaje1(Index).SetFocus
    Else
        KeyAscii = OnlyNumbers(KeyAscii)
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Function BuscarCriterio(CRITERIO As String) As Integer
  Dim i As Integer
  With MshPlanDeCuentas
    If OptCuenta.Value = True Then
        For i = 1 To CantRegistros
          If .TextMatrix(i, 1) Like UCase(CRITERIO) & "*" Then
            BuscarCriterio = i
            Exit Function
          End If
        Next
    End If
    
    If OptDescripcion.Value = True Then
        For i = 1 To CantRegistros
          If .TextMatrix(i, 2) Like UCase(CRITERIO) & "*" Then
            BuscarCriterio = i
            Exit Function
          End If
        Next
    End If
  End With
  BuscarCriterio = 0
End Function

Private Sub txtColCom_Change()
    If IsNumeric(txtColCom) = False Then
        txtColCom = Empty
    Else
        If Val(txtColCom) > 7 Then
            txtColCom = Empty
        End If
    End If
End Sub

Private Sub txtColCom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    RS.Open "SELECT LIB_COM FROM CNPARAM WHERE ANIO='" & AnioSistema & "'", ADOConexionEmp
    
    If KeyCode = 112 Then ' F1
        FlagColComVen = "C"
        oConsulta.Caso = 11
        oConsulta.NumLibro = Trim(RS(0))
        Caso = 11
        oConsulta.Formulario = "frmplancuentas"
        FrmConsultas.Caption = FrmConsultas.Caption + " de columnas C/V"
        FrmConsultas.Show
    End If
    
    RS.Clone
    Set RS = Nothing
End Sub

Private Sub txtColVen_Change()
    If IsNumeric(txtColVen) = False Then
        txtColVen = Empty
    Else
        If Val(txtColVen) > 7 Then
            txtColVen = Empty
        End If
    End If
End Sub

Private Sub txtColVen_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    RS.Open "SELECT LIB_VEN FROM CNPARAM WHERE ANIO='" & AnioSistema & "'", ADOConexionEmp
    
    If KeyCode = 112 Then ' F1
        FlagColComVen = "V"
        oConsulta.Caso = 11
        oConsulta.NumLibro = Trim(RS(0))
        Caso = 11
        oConsulta.Formulario = "frmplancuentas"
        FrmConsultas.Caption = FrmConsultas.Caption + " de columnas C/V"
        FrmConsultas.Show
    End If
    RS.Clone
    Set RS = Nothing
End Sub

Private Sub TxtCriterio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Dim fila As Integer
      fila = BuscarCriterio(TxtCriterio)
      If fila > 0 Then
        MshPlanDeCuentas.row = fila
        MshPlanDeCuentas.col = 0
        MshPlanDeCuentas.SetFocus
      Else
        MsgBox "No existe registro..", vbInformation, Caption
      End If
      SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer) 'SE CAMBIO 3/02/01
    If KeyCode = 112 And TxtAuxiliar.Locked = False Then
      If lblMensaje = "Modificar" Or lblMensaje = "Nuevo" Then
        'Dim oFrmConsulta As clsConsultas
        'Set oFrmConsulta = New clsConsultas
        oConsulta.Caso = 2
        oConsulta.Control = "TxtCuenta"
        oConsulta.Formulario = "frmplancuentas"
        FrmConsultas.Caption = FrmConsultas.Caption + " de Plan De Cuentas(Datos)"
        FrmConsultas.Show
      End If
    End If
End Sub

Private Sub TxtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim FilaEncontrada As Integer
        FilaEncontrada = BuscarCuenta(Trim(TxtCuenta))
        If FilaEncontrada > 0 Then
            If lblMensaje = "Nuevo" Then
                MsgBox "La Cuenta Ya Existe.Por Favor Ingrese Otra Cuenta", vbExclamation, "Plan De Cuentas"
                TxtCuenta.SetFocus
                SendKeys "{HOME}+{END}"
            End If
            If lblMensaje = "Modificar" Then
                Call ImprimirDatosDeMshPlanDeCuentas(FilaEncontrada, TxtCuenta)
                If CboTipCta.Text = CboTipCta.List(0) Then
                    Call ModoEdicionDato
                End If
                If CboTipCta.Text = CboTipCta.List(1) Then
                    Call ModoEdicionTitulo
                    CboTipCta.Locked = True
                    CboTipCta.BackColor = ColorDeshabilitado
                End If
                
                CboTipCta.SetFocus
                SendKeys "{HOME}+{END}"
            End If
            
        Else
            If lblMensaje = "Nuevo" Then
                CboTipCta.SetFocus
            End If
            If lblMensaje = "Modificar" Then
                MsgBox "La Cuenta No Existe.Por Favor Ingrese Otra Cuenta", vbExclamation, "Plan De Cuentas"
                TxtCuenta.SetFocus
                SendKeys "{HOME}+{END}"
            End If
        End If
    Else
        KeyAscii = OnlyNumbers(KeyAscii)
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmdGrabar.Enabled = True Then
        If CboTipCta = CboTipCta.List(1) Then ' SI ES TITULO
            Call cmdGrabar_Click
        Else
            TxtAuxiliar.SetFocus
        End If
    Else
        'KeyAscii = OnlyChar(KeyAscii)
    End If
End Sub


