VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmTareoPlani 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Tareo"
   ClientHeight    =   9735
   ClientLeft      =   1665
   ClientTop       =   6915
   ClientWidth     =   14670
   ForeColor       =   &H00000000&
   Icon            =   "frmTareoPlani.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9735
   ScaleWidth      =   14670
   Begin VB.Frame fraCombos 
      BackColor       =   &H009F5539&
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
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   14670
      Begin MSMask.MaskEdBox DtpFecha 
         Height          =   285
         Left            =   10725
         TabIndex        =   2
         ToolTipText     =   "Fecha_Pago"
         Top             =   180
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   128
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   195
         Width           =   405
      End
      Begin MSForms.ComboBox cboAnio 
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   165
         Width           =   1305
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2302;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboMon 
         Height          =   315
         Left            =   5010
         TabIndex        =   11
         Top             =   165
         Width           =   1545
         VariousPropertyBits=   746604569
         DisplayStyle    =   7
         Size            =   "2725;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboProceso 
         Height          =   315
         Left            =   7545
         TabIndex        =   10
         Top             =   165
         Width           =   2355
         VariousPropertyBits=   746604569
         DisplayStyle    =   7
         Size            =   "4154;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboMes 
         Height          =   315
         Left            =   2355
         TabIndex        =   9
         Top             =   165
         Width           =   1695
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2990;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   4185
         TabIndex        =   8
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   1905
         TabIndex        =   6
         Top             =   195
         Width           =   435
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   9945
         TabIndex        =   3
         Top             =   195
         Width           =   705
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Planilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6765
         TabIndex        =   1
         Top             =   195
         Width           =   735
      End
   End
   Begin VB.Frame fraFondo 
      BackColor       =   &H009F5539&
      Height          =   9420
      Left            =   -30
      TabIndex        =   4
      Top             =   330
      Width           =   14700
      Begin TabDlg.SSTab SSTab1 
         Height          =   7185
         Left            =   45
         TabIndex        =   5
         Top             =   180
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   12674
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   706
         BackColor       =   14737632
         ForeColor       =   10442041
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Greek"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "  Lista de Personal"
         TabPicture(0)   =   "frmTareoPlani.frx":014A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblCodEmp"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "flxTareo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cbMvto"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraMensajes"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin VB.Frame fraMensajes 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   330
            Left            =   9615
            TabIndex        =   41
            Top             =   75
            Width           =   4875
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Sistema Pensiones:"
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
               Height          =   465
               Left            =   2070
               TabIndex        =   45
               Top             =   -45
               Width           =   975
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Base:"
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
               Height          =   255
               Left            =   0
               TabIndex        =   44
               Top             =   45
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label lblPension 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   3000
               TabIndex        =   43
               Top             =   45
               Width           =   1755
            End
            Begin VB.Label lblBase 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Prima"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   255
               Left            =   510
               TabIndex        =   42
               Top             =   45
               Visible         =   0   'False
               Width           =   1545
            End
         End
         Begin VB.ComboBox cbMvto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4110
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1230
            Visible         =   0   'False
            Width           =   2535
         End
         Begin NOVAdmin.flxEdit flxTareo 
            Height          =   6675
            Left            =   120
            TabIndex        =   47
            Top             =   420
            Width           =   14475
            _ExtentX        =   21828
            _ExtentY        =   7805
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CellFontName    =   "MS Sans Serif"
            CellFontSize    =   8.25
            BackColorSel    =   -2147483643
            BackColorFixed  =   9868950
            CellPicture     =   "frmTareoPlani.frx":02A4
            ColAlignment0   =   9
            FixedAlignment0 =   9
            ForeColorSel    =   16711680
            ForeColorFixed  =   14474460
            MouseIcon       =   "frmTareoPlani.frx":02C0
            RowHeight0      =   240
         End
         Begin VB.Label lblCodEmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   60
            TabIndex        =   12
            Top             =   90
            Width           =   4455
         End
      End
      Begin VB.PictureBox picBotones 
         BackColor       =   &H009F5539&
         BorderStyle     =   0  'None
         Height          =   2040
         Left            =   135
         ScaleHeight     =   2040
         ScaleWidth      =   12480
         TabIndex        =   15
         Top             =   7260
         Width           =   12480
         Begin VB.Frame Frame3 
            BackColor       =   &H009F5539&
            Height          =   1305
            Left            =   10020
            TabIndex        =   19
            Top             =   0
            Width           =   2445
            Begin Proyecto1.chameleonButton chBtnSalir 
               CausesValidation=   0   'False
               Height          =   375
               Left            =   1920
               TabIndex        =   20
               ToolTipText     =   "Salir"
               Top             =   180
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   661
               BTYPE           =   14
               TX              =   ""
               ENAB            =   -1  'True
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
               MICON           =   "frmTareoPlani.frx":02DC
               PICN            =   "frmTareoPlani.frx":02F8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin Proyecto1.chameleonButton chBtnReporte 
               Height          =   375
               Left            =   1290
               TabIndex        =   21
               ToolTipText     =   "Ver Reporte"
               Top             =   180
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   661
               BTYPE           =   14
               TX              =   ""
               ENAB            =   -1  'True
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
               MICON           =   "frmTareoPlani.frx":06BE
               PICN            =   "frmTareoPlani.frx":06DA
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
               Left            =   1290
               TabIndex        =   22
               ToolTipText     =   "Guardar"
               Top             =   900
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
               MICON           =   "frmTareoPlani.frx":0C1C
               PICN            =   "frmTareoPlani.frx":0C38
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin Proyecto1.chameleonButton btnEliminaEmp 
               Height          =   345
               Left            =   90
               TabIndex        =   23
               ToolTipText     =   "Modificar"
               Top             =   180
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   609
               BTYPE           =   14
               TX              =   "Quitar"
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
               MICON           =   "frmTareoPlani.frx":107A
               PICN            =   "frmTareoPlani.frx":1096
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
               Left            =   90
               TabIndex        =   24
               Top             =   900
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   609
               BTYPE           =   14
               TX              =   "&Eliminar"
               ENAB            =   -1  'True
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
               MICON           =   "frmTareoPlani.frx":11F0
               PICN            =   "frmTareoPlani.frx":120C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblModo 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Acción"
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
               Height          =   225
               Left            =   3030
               TabIndex        =   25
               Top             =   -90
               Visible         =   0   'False
               Width           =   765
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000004&
               BorderWidth     =   3
               X1              =   30
               X2              =   2430
               Y1              =   690
               Y2              =   690
            End
         End
         Begin Proyecto1.chameleonButton CmdCarga 
            Height          =   345
            Left            =   9480
            TabIndex        =   16
            ToolTipText     =   "Carga Todos los Descuentos y Aportes"
            Top             =   945
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   ""
            ENAB            =   -1  'True
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
            MICON           =   "frmTareoPlani.frx":164E
            PICN            =   "frmTareoPlani.frx":166A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   6660
            Top             =   1500
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "Excel (*.XLSX)|*.xlsx"
         End
         Begin Proyecto1.chameleonButton btnCargarD 
            Height          =   345
            Left            =   8430
            TabIndex        =   17
            ToolTipText     =   "Carga Descuento Seleccionado"
            Top             =   540
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
            MICON           =   "frmTareoPlani.frx":209F
            PICN            =   "frmTareoPlani.frx":20BB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnCargarA 
            Height          =   345
            Left            =   8430
            TabIndex        =   18
            ToolTipText     =   "Carga Aporte Seleccionado"
            Top             =   945
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
            MICON           =   "frmTareoPlani.frx":2ACD
            PICN            =   "frmTareoPlani.frx":2AE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnCargarI 
            Height          =   345
            Left            =   8430
            TabIndex        =   26
            ToolTipText     =   "Carga Ingreso Seleccionado"
            Top             =   120
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
            MICON           =   "frmTareoPlani.frx":34FB
            PICN            =   "frmTareoPlani.frx":3517
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton cmdAnexar 
            Height          =   375
            Left            =   11970
            TabIndex        =   27
            ToolTipText     =   "Nuevo"
            Top             =   1500
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   661
            BTYPE           =   14
            TX              =   ""
            ENAB            =   -1  'True
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
            MICON           =   "frmTareoPlani.frx":3F29
            PICN            =   "frmTareoPlani.frx":3F45
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnDesCargarI 
            Height          =   345
            Left            =   8925
            TabIndex        =   28
            ToolTipText     =   "Elimina Carga de Ingreso Seleccionado"
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
            MICON           =   "frmTareoPlani.frx":409F
            PICN            =   "frmTareoPlani.frx":40BB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnDesCargarD 
            Height          =   345
            Left            =   8925
            TabIndex        =   29
            ToolTipText     =   "Elimina Carga de Descuento Seleccionado"
            Top             =   540
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
            MICON           =   "frmTareoPlani.frx":43D5
            PICN            =   "frmTareoPlani.frx":43F1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnDesCargarA 
            Height          =   345
            Left            =   8925
            TabIndex        =   30
            ToolTipText     =   "Elimina Carga de Aporte Seleccionado"
            Top             =   945
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
            MICON           =   "frmTareoPlani.frx":470B
            PICN            =   "frmTareoPlani.frx":4727
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnImportar 
            Height          =   345
            Left            =   9480
            TabIndex        =   31
            Top             =   540
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   ""
            ENAB            =   -1  'True
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
            MICON           =   "frmTareoPlani.frx":4A41
            PICN            =   "frmTareoPlani.frx":4A5D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Proyecto1.chameleonButton btnActSueldo 
            Height          =   345
            Left            =   9480
            TabIndex        =   32
            ToolTipText     =   "Actualiza Sueldos y AFP"
            Top             =   150
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            BTYPE           =   14
            TX              =   ""
            ENAB            =   -1  'True
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
            MICON           =   "frmTareoPlani.frx":4BB7
            PICN            =   "frmTareoPlani.frx":4BD3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblCadBus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   30
            TabIndex        =   46
            Top             =   1830
            Width           =   75
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   0
            X2              =   12420
            Y1              =   1440
            Y2              =   1440
         End
         Begin MSForms.ComboBox cboIngresos 
            Height          =   315
            Left            =   1320
            TabIndex        =   40
            Top             =   150
            Width           =   7065
            VariousPropertyBits=   746604569
            DisplayStyle    =   7
            Size            =   "12462;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial"
            FontEffects     =   1073750016
            FontHeight      =   135
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CboDescuentos 
            Height          =   315
            Left            =   1320
            TabIndex        =   39
            Top             =   555
            Width           =   7065
            VariousPropertyBits=   746604569
            DisplayStyle    =   7
            Size            =   "12462;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial"
            FontEffects     =   1073750016
            FontHeight      =   135
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboAportes 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   960
            Width           =   7065
            VariousPropertyBits=   746604569
            DisplayStyle    =   7
            Size            =   "12462;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial"
            FontEffects     =   1073750016
            FontHeight      =   135
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "DESCUENTOS"
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
            Height          =   225
            Left            =   0
            TabIndex        =   37
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "APORTES"
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
            Height          =   225
            Left            =   0
            TabIndex        =   36
            Top             =   1005
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "INGRESOS"
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
            Height          =   195
            Left            =   30
            TabIndex        =   35
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Lbl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Palabra de Búsqueda (Esc para borrar)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   34
            Top             =   1530
            Width           =   3495
         End
         Begin MSForms.ComboBox cboAdicionales 
            Height          =   315
            Left            =   7230
            TabIndex        =   33
            Top             =   1530
            Width           =   4725
            VariousPropertyBits=   746604569
            DisplayStyle    =   7
            Size            =   "8334;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Arial"
            FontEffects     =   1073750016
            FontHeight      =   135
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
   End
End
Attribute VB_Name = "frmTareoPlani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim Rs As MYSQL_RS
Dim coloremp As Variant, cargando As Boolean
Dim cadenaemp As String
Dim NroEmpleados As Integer
Dim Periodo As String, plani As String, moneda As String, Rubro As String, empselect As String
Dim FlgCarga As Boolean
Dim FlgAct As Boolean, CodEmpSelect As String
Dim Veces As Integer, FlgRubAl As Boolean, FlgBloq As Boolean
Dim gsRubro As String

Public Sub ConfiguraGrilla()
    Dim i As Integer
    With flxTareo
        .Clear
        .Rows = 1
        .Cols = 10
        .ForeColorFixed = &H404000
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        .ColWidth(1) = 3500
        .TextMatrix(0, 1) = Space(40) + "Empleado"
        .ColWidth(2) = 650
        .TextMatrix(0, 2) = "TipRub"
        .ColWidth(3) = 7000
        .TextMatrix(0, 3) = Space(30) + "Rubro"
        .ColWidth(4) = 500
        .TextMatrix(0, 4) = "Uni"
        .ColType(4) = cadena
        .ColMaxLength(4) = 1
        .CaracteresValidos(4) = "VD"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(2) + "Cantidad"
        .ColType(5) = Numero
        .ColMaxLength(5) = 20
        .CaracteresValidos(5) = "0123456789.,"
        .ColDecimales(5) = 2
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = "EMP"
        .ColType(6) = cadena
        .ColMaxLength(6) = 11
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "SB"
        .ColType(7) = cadena
        .ColMaxLength(7) = 11
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "afp"
        .ColType(8) = cadena
        .ColMaxLength(8) = 11
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = "contrato"
        .ColType(9) = cadena
        .ColMaxLength(9) = 11
    End With
    'If EstadoPlani(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), moneda, plani) = True Then
    '    flxTareo.Enabled = False
    'Else
    '    flxTareo.Enabled = True
    'End If
End Sub
Private Sub btnActSueldo_Click()
On Error GoTo contrato
    If flxTareo.ColSel = 9 And btnGrabar.Enabled = False Then
        If MsgBox("¿Esta seguro que desea actualizar el sueldo actual y/o el afp del empleado?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
            Screen.MousePointer = vbHourglass
            If empselect <> "" Then
                CodEmpSelect = empselect
                SQL = "Update pl_tareo set sbasico=(select sbasico from contrato where estado='AP' and codemp='" & empselect & "'), " & _
                      "codcontrato=(select codigo from contrato where estado='AP' and codemp='" & empselect & "')," & _
                      "afp = (select codafp from empleado where codigo='" & empselect & "') " & _
                      " where anomes='" & cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & _
                      "' and tipo='" & plani & "' and moneda='" & moneda & "' and emp='" & empselect & "'"
                oConexionMYSQL.Execute SQL
                FlgAct = True
                cboMon_Change
                FlgCarga = True
                For i = 1 To CboDescuentos.ListCount - 1
                    CboDescuentos.ListIndex = i
                    If plani = 6 Then
                        If CboDescuentos.List(CboDescuentos.ListIndex, 2) = "401" Or CboDescuentos.List(CboDescuentos.ListIndex, 2) = "201" Then
                            CargarDescuentos CodEmpSelect
                        End If
                    Else
                        CargarDescuentos CodEmpSelect
                    End If
                Next
                CboDescuentos.ListIndex = 0
                For i = 1 To cboAportes.ListCount - 1
                    cboAportes.ListIndex = i
                    If plani = 6 Then
                        If cboAportes.List(cboAportes.ListIndex, 2) = "401" Or cboAportes.List(cboAportes.ListIndex, 2) = "201" Then
                            CargaAportes CodEmpSelect
                        End If
                    Else
                        CargaAportes CodEmpSelect
                    End If
                Next
                cboAportes.ListIndex = 0
                cboProceso_Change
                FlgCarga = False
                flxTareo.Visible = True
                Screen.MousePointer = vbNormal
                FlgAct = False
            End If
         End If
    End If
Exit Sub
contrato:
    MsgBox "Revise los contratos del personal", vbExclamation + vbOKOnly, "NOVPeru"
    Screen.MousePointer = vbNormal
    Exit Sub
End Sub
Private Sub btnCargarA_Click()
    CargaAportes ""
End Sub
Sub CargaAportes(Optional CodEmp As String)
    With cboAportes
        SQL = "Select formula from pl_rubrosremunerativos where codigo='" & .List(.ListIndex, 2) & "'"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If IsNull(Rs.Fields("formula")) = False Then
            CargaRubroAutomatico .List(.ListIndex, 1), .List(.ListIndex, 2), .List(.ListIndex, 0), Rs.Fields("formula"), CodEmp
        End If
        If FlgCarga = False Then
            mensaje = "Se cargó " & .List(.ListIndex, 0)
            BotonRubro .List(.ListIndex, 2), "A"
        End If
    End With
    If FlgAct = False Then cboProceso_Change
    If FlgCarga = False Then MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
    Set Rs = Nothing
End Sub
Private Sub btnCargarD_Click()
    CargarDescuentos ""
End Sub
Sub CargarDescuentos(Optional CodEmp As String)
    With CboDescuentos
        SQL = "Select formula from pl_rubrosremunerativos where codigo='" & .List(.ListIndex, 2) & "'"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If IsNull(Rs.Fields("formula")) = False Then
            CargaRubroAutomatico .List(.ListIndex, 1), .List(.ListIndex, 2), .List(.ListIndex, 0), Rs.Fields("formula"), CodEmp
        End If
        If FlgCarga = False Then
            mensaje = "Se cargó " & .List(.ListIndex, 0)
            BotonRubro .List(.ListIndex, 2), "D"
        End If
    End With
    If FlgAct = False Then cboProceso_Change
    If FlgCarga = False Then MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
    Set Rs = Nothing
End Sub
Private Sub btnCargarI_Click()
    With cboIngresos
        SQL = "Select formula from pl_rubrosremunerativos where codigo='" & .List(.ListIndex, 2) & "'"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If IsNull(Rs.Fields("formula")) = False Then
            CargaRubroAutomatico .List(.ListIndex, 1), .List(.ListIndex, 2), .List(.ListIndex, 0), Rs.Fields("formula")
        End If
        mensaje = "Se cargó " & .List(.ListIndex, 0)
        BotonRubro .List(.ListIndex, 2), "I"
    End With
    cboProceso_Change
    MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
    Set Rs = Nothing
End Sub
Private Sub btnDesCargarA_Click()
    Dim resp As Integer
    With cboAportes
        resp = MsgBox("¿Está seguro de eliminar el tareo del rubro:" & vbNewLine & _
                      .List(.ListIndex, 0), vbYesNo + vbQuestion, "NOVPeru")
        If resp = vbYes Then
            SQL = "Delete from pl_tareo where anomes='" & Periodo & "' and tipo='" & plani & _
                  "' and moneda='" & moneda & "' and rub='" & .List(.ListIndex, 2) & "'"
            oConexionMYSQL.Execute SQL
            mensaje = "Se Descargó " & .List(.ListIndex, 0)
            MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
        End If
        BotonRubro .List(.ListIndex, 2), "A"
    End With
    cboProceso_Change
End Sub
Private Sub btnDesCargarD_Click()
    Dim resp As Integer
    With CboDescuentos
        resp = MsgBox("Está seguro de eliminar el tareo del rubro:" & vbNewLine & _
                      .List(.ListIndex, 0), vbYesNo + vbQuestion, "NOVPeru")
        If resp = vbYes Then
            SQL = "Delete from pl_tareo where anomes='" & Periodo & "' and tipo='" & plani & _
                "' and moneda='" & moneda & "' and rub='" & .List(.ListIndex, 2) & "'"
            oConexionMYSQL.Execute SQL
            If .List(.ListIndex, 2) = "709" Or .List(.ListIndex, 2) = "701" Then
                ActualizaMontos Periodo, plani, moneda, IIf(.List(.ListIndex, 2) = "709", "P", "A"), "", 1
            End If
            mensaje = "Se Descargó " & .List(.ListIndex, 0)
            MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
        End If
        BotonRubro .List(.ListIndex, 2), "D"
    End With
    cboProceso_Change
End Sub
Private Sub btnDesCargarI_Click()
    Dim resp As Integer
    With cboIngresos
        resp = MsgBox("Está seguro de eliminar el tareo del rubro:" & vbNewLine & _
                      .List(.ListIndex, 0), vbYesNo + vbQuestion, "NOVPeru")
        If resp = vbYes Then
            SQL = "Delete from pl_tareo where anomes='" & Periodo & "' and tipo='" & plani & _
                  "' and moneda='" & moneda & "' and rub='" & .List(.ListIndex, 2) & "'"
            oConexionMYSQL.Execute SQL
            mensaje = "Se Descargó " & .List(.ListIndex, 0)
            MsgBox mensaje, vbOKOnly + vbInformation, "NOVPeru"
        End If
        BotonRubro .List(.ListIndex, 2), "I"
    End With
    cboProceso_Change
End Sub
Private Sub btnEliminaEmp_Click()
    If flxTareo.ColSel = 9 And btnGrabar.Enabled = False Then
        If MsgBox("Está seguro de eliminar el tareo de este empleado", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
            SQL = "Delete from pl_tareo where anomes='" & cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "' and emp='" & empselect & "'"
            oConexionMYSQL.Execute SQL
            EliminaMtoPagado cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, empselect
            SQL = "Delete from rh_pagosemp where anomesplani='" & cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "' and tipoplani='" & plani & "' and monplani='" & moneda & "' AND solicita= '" & empselect & "' and liquid = 'N'"
            oConexionMYSQL.Execute SQL
            cboProceso_Change
        End If
    End If
End Sub
Private Sub btnEliminar_Click()
    If MsgBox("Esta seguro de eliminar el tareo de esta planilla", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        If Left(cboMon.Text, 1) = "C" Then
            SQL = "Delete from pl_tareo where anomes='" & cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "' and tipo='" & plani & "' and moneda='" & moneda & "' AND EMP IN(SELECT E.CODIGO AS EMP FROM EMPLEADO AS E WHERE E.TIPOPLANI='S')"
        Else
            SQL = "Delete from pl_tareo where anomes='" & cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2) & "' and tipo='" & plani & "' and moneda='" & moneda & "'"
        End If
        oConexionMYSQL.Execute SQL
'        If Left(cboMon.Text, 1) = "C" Then
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "P", "00000000002", 1
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "A", "00000000002", 1
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "P", "00000000112", 1
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "A", "00000000112", 1
'        Else
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "P", "", 1
'            ActualizaMontos cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), plani, moneda, "A", "", 1
'        End If
        cboProceso_Change
        cboIngresos.ListIndex = 0
        cboAportes.ListIndex = 0
        CboDescuentos.ListIndex = 0
    End If
End Sub
Private Sub btnGrabar_Click()
    Dim i As Integer
    If GrabaTareo(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), IIf(Left(cboMon.Text, 1) = "C", "N", Left(cboMon.Text, 1)), flxTareo.Rows - 1, Format(DtpFecha, "yyyy/mm/dd")) Then
        MsgBox "Se registró correctamente el tareo para la planilla", vbOKOnly + vbInformation
        cboProceso_Change
        CboDescuentos.Enabled = True
        cboAportes.Enabled = True
        cboIngresos.Enabled = True
        btnEliminar.Enabled = True
    Else
        MsgBox "No se pudo grabar la información para el empleado:" & emp & vbNewLine & _
               " revise el tareo o consulte con el administrador del sistema", vbOKOnly + vbExclamation, "NOVPeru"
    End If
End Sub
Public Function GrabaTareo(AnoMes As String, proc As String, mon As String, Cant As Long, fecha As String) As Boolean
    On Error GoTo noGrAba
    Dim i As Integer
    Dim SQL As String, emp As String, rub As String, CANTIDAD As Double, sb As Double, uni As String, afp As String, contrato As String
    ValidaTareo Cant
    If Left(cboMon.Text, 1) = "C" Then
        SQL = "Delete from pl_tareo where anomes='" & AnoMes & "' and tipo='" & proc & "' and moneda='" & mon & "' and emp in('00000000002','00000000112')"
    Else
        SQL = "Delete from pl_tareo where anomes='" & AnoMes & "' and tipo='" & proc & "' and moneda='" & mon & "'"
    End If
    oConexionMYSQL.Execute SQL
    With flxTareo
        For i = 1 To Cant
            emp = .TextMatrix(i, 6)
            rub = Trim(Left(.TextMatrix(i, 3), 4))
            CANTIDAD = .TextMatrix(i, 5)
            sb = .TextMatrix(i, 7)
            afp = .TextMatrix(i, 8)
            uni = .TextMatrix(i, 4)
            contrato = .TextMatrix(i, 9)
            SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                  "'" & AnoMes & "'," & _
                  "'" & proc & "'," & _
                  "'" & mon & "'," & _
                  "'" & emp & "'," & _
                  "'" & rub & "'," & _
                  "'" & uni & "'," & _
                  "" & CANTIDAD & "," & _
                  "'" & fecha & "'," & sb & ",'" & afp & "','" & contrato & "')"
            oConexionMYSQL.Execute SQL
        Next
    End With
    GrabaTareo = True
Exit Function
noGrAba:
    GrabaTareo = False
    Resume
End Function
Private Sub btnImportar_Click()
    Dim UltimaCelda As Variant, fila As Long, i As Long, err As Boolean, mensaje As Integer
    Dim emp As String, rub As String, uni As String, Cant As Double, sb As Double, afp As String, fecha As String
    Dim CantM As Double, rubM As String, Flg As Boolean, contrato As String
    CommonDialog1.ShowOpen
    NombreArchivo = CommonDialog1.Filename
    If NombreArchivo = "" Then
        Excel.Application.Workbooks.Close
        Excel.Application.Quit
        Exit Sub
    End If
    Excel.Application.Workbooks.Open Filename:=NombreArchivo
    Excel.Application.Visible = False
    Hoja = "TAREO"
    UltimaCelda = ActiveCell.SpecialCells(xlCellTypeLastCell).Address
    fila = CDbl(Right(UltimaCelda, Len(UltimaCelda) - InStr(2, UltimaCelda, "$", vbBinaryCompare)))
    i = 1
    Do While i <= fila
        err = False
        emp = Right("00000000000" & Trim(CStr(Worksheets(Hoja).Cells(i, 1).Value)), 11)
        SQL = "Select distinct afp,sbasico,fecha,codcontrato from pl_tareo where anomes='" & Periodo & _
              "' and tipo='" & plani & "' and emp = '" & emp & "'"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If Not (Rs.EOF And Rs.BOF) Then
            sb = Rs.Fields("sbasico")
            afp = Rs.Fields("afp")
            fecha = Rs.Fields("fecha")
            contrato = Rs.Fields("codcontrato")
            rub = Trim(Worksheets(Hoja).Cells(i, 2).Value)
            uni = Trim(Worksheets(Hoja).Cells(i, 3).Value)
            If uni <> "V" And uni <> "D" Then err = True
            Cant = CDbl(Worksheets(Hoja).Cells(i, 4).Value)
            Flg = False
            SQL = "Select * from pl_tareo where anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and emp = '" & emp & "' and rub = '" & rub & "'"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            If Not Rs.EOF() Then
                Cant = Cant + Rs.Fields("cant")
                Flg = True
            End If
            If Cant <= 0 Then err = True
            If err = False Then
                If Flg = True Then
                    SQL = "update pl_tareo set cant = " & Cant & " where anomes='" & Periodo & _
                          "' and tipo='" & plani & "' and emp = '" & emp & "' and rub = '" & rub & "'"
                Else
                    SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                          "'" & Periodo & "'," & _
                          "'" & plani & "'," & _
                          "'" & moneda & "'," & _
                          "'" & emp & "'," & _
                          "'" & rub & "'," & _
                          "'" & uni & "'," & _
                          "" & Cant & "," & _
                          "'" & fecha & "'," & sb & ",'" & afp & "','" & contrato & "')"
                    On Error GoTo ErrGraba
                End If
                oConexionMYSQL.Execute SQL
            Else
                MsgBox "Uno de los datos del empleado " & emp & " no es correcto para el tareo", vbInformation + vbOKOnly, "NOVPeru"
            End If
        Else
            MsgBox "Empleado " & emp & " no está registrado en el tareo de esta planilla", vbInformation + vbOKOnly, "NOVPeru"
        End If
        i = i + 1
    Loop
    Excel.Application.Workbooks.Close
    Excel.Application.Quit
Exit Sub
ErrGraba:
    mensaje = MsgBox("El rubro registrado en la fila " & CStr(i) & " no es correcto o ya está ingresado en el tareo", vbExclamation + vbOKOnly + vbOKCancel, "NOVPeru")
    If mensaje = vbCancel Then
        Excel.Application.Workbooks.Close
        Excel.Application.Quit
        Exit Sub
    Else
        Resume Next
    End If
End Sub
Private Sub btnNuevaFila_Click()
    Dim lastRow As Long
    With flxTareo
        lastRow = .row
        If .row = lastRow Then
            lastRow = lastRow + 1
            .AddItem "", lastRow
            .TextMatrix(lastRow, 1) = .TextMatrix(.row, 1)
            .TextMatrix(lastRow, 6) = .TextMatrix(.row, 6)
            .TextMatrix(lastRow, 7) = .TextMatrix(.row, 7)
            .TextMatrix(lastRow, 8) = .TextMatrix(.row, 8)
            .TextMatrix(lastRow, 9) = .TextMatrix(.row, 9)
        End If
        .row = lastRow
        .Col = 2
    End With
End Sub
Public Sub EnumerarFlex(grilla As flxEdit)
    Dim varaux As String
    Dim cuenta As Integer, Y As Integer, i As Long
    varaux = ""
    cuenta = 1
    Y = 1
    For i = 1 To grilla.Rows - 1
        If i > 1 Then
            If grilla.TextMatrix(i - 1, 1) <> grilla.TextMatrix(i, 1) Then
                Y = Y * -1
            End If
        End If
        If varaux <> grilla.TextMatrix(i, 1) Then
            varaux = grilla.TextMatrix(i, 1)
            grilla.TextMatrix(i, 0) = cuenta
            cuenta = cuenta + 1
            grilla.Col = 0
            grilla.row = i
            PintaFila i, Y
        Else
            grilla.TextMatrix(i, 0) = ""
        End If
    Next
    NroEmpleados = cuenta - 1
End Sub
Private Sub btnBorrar_Click()
    Dim lastRow As Long
    With flxTareo
        lastRow = .Rows - 1
        If lastRow = 1 Then
            For J = 1 To .Cols - 1
                .TextMatrix(lastRow, J) = ""
            Next J
        Else
            .RemoveItem (.row)
            EnumerarItems flxTareo
        End If
    End With
    flxEgresos.Col = 1
End Sub
Private Sub cbMvto_Click()
    With flxTareo
        If cbMvto.ListIndex > 0 Then
            If .Col = 2 Then
                .TextMatrix(.row, .Col) = Trim(Left(cbMvto.List(cbMvto.ListIndex), 4))
            End If
            If .Col = 3 Then
                .TextMatrix(.row, .Col) = Trim(cbMvto.List(cbMvto.ListIndex))
            End If
            cbMvto.Visible = False
        Else
            .TextMatrix(.row, .Col) = ""
        End If
    End With
End Sub
Private Sub cbMvto_LostFocus()
    cbMvto.Visible = False
End Sub
Private Sub cboAnio_Change()
    If cboAnio.ListIndex > 0 Then
        cboMes.Enabled = True
    Else
        If cboMes.ListCount > 0 Then cboMes.ListIndex = 0
        cboMes.Enabled = False
    End If
    cboProceso_Change
    FlgBloq = False
    If Bloqueo Then
        FlgBloq = True
    End If
    If FlgBloq = True Then
        btnCargarA.Enabled = False: btnCargarD.Enabled = False: btnCargarI.Enabled = False
        btnDesCargarA.Enabled = False: btnDesCargarD.Enabled = False: btnDesCargarI.Enabled = False
        btnActSueldo.Enabled = False: btnImportar.Enabled = False: CmdCarga.Enabled = False
        cboIngresos.Enabled = False: CboDescuentos.Enabled = False: cboAportes.Enabled = False
        cmdAnexar.Enabled = False: cboAdicionales.Enabled = False: btnEliminaEmp.Enabled = False
        btnEliminar.Enabled = False: btnGrabar.Enabled = False
    End If
    
    
    'Call Form_Resize
End Sub
Private Sub cboAportes_Click()
    If FlgCarga = False Then
        If cboAportes.ListIndex > 0 Then
            If MsgBox("¿Está seguro que terminó de registrar todos rubros de ingreso?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                btnCargarA.Enabled = True
                BotonRubro cboAportes.List(cboAportes.ListIndex, 2), "A"
            End If
        Else
            btnCargarA.Enabled = False
            btnDesCargarA.Enabled = False
        End If
    End If
End Sub
Private Sub CboDescuentos_Click()
    If FlgCarga = False Then
        If FlgRubAl = False Then
            If CboDescuentos.ListIndex > 0 Then
                If MsgBox("¿Está seguro que termino de registrar todos rubros de ingreso?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    btnCargarD.Enabled = True
                    BotonRubro CboDescuentos.List(CboDescuentos.ListIndex, 2), "D"
                End If
            Else
                btnCargarD.Enabled = False
                btnDesCargarD.Enabled = False
            End If
        End If
    End If
End Sub
Private Sub cboIngresos_Click()
    If cboIngresos.ListIndex > 0 Then
        BotonRubro cboIngresos.List(cboIngresos.ListIndex, 2), "I"
        btnCargarI.Enabled = True
    Else
        btnCargarI.Enabled = False
        btnDesCargarI.Enabled = False
    End If
End Sub
Private Sub cboMes_Change()
    If cboMes.ListIndex > 0 Then
        cboMon.Enabled = True
    Else
        If cboMon.ListCount > 0 Then cboMon.ListIndex = 0
        cboMon.Enabled = False
    End If
    Periodo = cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2)
    cboProceso_Change
    FlgBloq = False
    If Bloqueo Then
        FlgBloq = True
    End If
    If FlgBloq = True Then
        btnCargarA.Enabled = False: btnCargarD.Enabled = False: btnCargarI.Enabled = False
        btnDesCargarA.Enabled = False: btnDesCargarD.Enabled = False: btnDesCargarI.Enabled = False
        btnActSueldo.Enabled = False: btnImportar.Enabled = False: CmdCarga.Enabled = False
        cboIngresos.Enabled = False: CboDescuentos.Enabled = False: cboAportes.Enabled = False
        cmdAnexar.Enabled = False: cboAdicionales.Enabled = False: btnEliminaEmp.Enabled = False
        btnEliminar.Enabled = False: btnGrabar.Enabled = False
    End If
    
    'Call Form_Resize
End Sub
Private Sub cboMon_Change()
    If cboMon.ListIndex > 0 Then
        cboProceso.Enabled = True
        cboProceso_Change
    Else
        If cboProceso.ListCount > 0 Then cboProceso.ListIndex = 0
        cboProceso.Enabled = False
    End If
    moneda = IIf(Left(cboMon.Text, 1) = "C", "N", Left(cboMon.Text, 1))
    If Bloqueo Then
        FlgBloq = True
    Else
        FlgBloq = False
    End If
    If FlgBloq = True Then
        btnCargarA.Enabled = False: btnCargarD.Enabled = False: btnCargarI.Enabled = False
        btnDesCargarA.Enabled = False: btnDesCargarD.Enabled = False: btnDesCargarI.Enabled = False
        btnActSueldo.Enabled = False: btnImportar.Enabled = False: CmdCarga.Enabled = False
        cboIngresos.Enabled = False: CboDescuentos.Enabled = False: cboAportes.Enabled = False
        cmdAnexar.Enabled = False: cboAdicionales.Enabled = False: btnEliminaEmp.Enabled = False
        btnEliminar.Enabled = False: btnGrabar.Enabled = False
    End If
    
    'Call Form_Resize
End Sub
Private Sub cboProceso_Change()
    plani = ""
    cadenaemp = ""
    lblCadBus = ""
    If cboProceso.ListCount > 1 And cboProceso.Enabled = True Then
        cargando = True
        flxTareo.Visible = False
        plani = cboProceso.List(cboProceso.ListIndex, 2)
        ConfiguraGrilla
        If CargarTareo(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), IIf(Left(cboMon.Text, 1) = "C", "N", Left(cboMon.Text, 1))) = False Then
            Select Case plani
                Case "1", "2", "5", "6":
                    LlenarEmpleados IIf(Left(Trim(cboMon.List(cboMon.ListIndex)), 1) = "C", "N", Left(Trim(cboMon.List(cboMon.ListIndex)), 1)), plani
                    CompletarPlantilla cboProceso.List(cboProceso.ListIndex, 2), flxTareo.Rows - 1, cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2)
                Case "4"
                    LlenarEmpVac IIf(Left(Trim(cboMon.List(cboMon.ListIndex)), 1) = "C", "N", Left(Trim(cboMon.List(cboMon.ListIndex)), 1))
                    CompletarPlantilla cboProceso.List(cboProceso.ListIndex, 2), flxTareo.Rows - 1, cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2)
            End Select
            DtpFecha = Format(Date, "dd/mm/yyyy")
            CboDescuentos.Enabled = False
            cboAportes.Enabled = False
            cboIngresos.Enabled = False
            btnGrabar.Enabled = True
            btnEliminar.Enabled = False
        Else
            If EstadoPlani(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), moneda, plani) = False Then
                CboDescuentos.Enabled = True
                cboAportes.Enabled = True
                cboIngresos.Enabled = True
                btnEliminar.Enabled = True
            Else
                CboDescuentos.Enabled = False
                cboAportes.Enabled = False
                cboIngresos.Enabled = False
                btnEliminar.Enabled = False
            End If
            btnGrabar.Enabled = False
            EnumerarFlex flxTareo
        End If
        cargando = False
        If FlgCarga = False Then flxTareo.Visible = True
        If Bloqueo Then
            FlgBloq = True
        Else
            FlgBloq = False
        End If
        If FlgBloq = True Then
            btnCargarA.Enabled = False: btnCargarD.Enabled = False: btnCargarI.Enabled = False
            btnDesCargarA.Enabled = False: btnDesCargarD.Enabled = False: btnDesCargarI.Enabled = False
            btnActSueldo.Enabled = False: btnImportar.Enabled = False: CmdCarga.Enabled = False
            cboIngresos.Enabled = False: CboDescuentos.Enabled = False: cboAportes.Enabled = False
            cmdAnexar.Enabled = False: cboAdicionales.Enabled = False: btnEliminaEmp.Enabled = False
            btnEliminar.Enabled = False: btnGrabar.Enabled = False
        End If
    End If
    
    'Call Form_Resize
End Sub
Public Function CargarTareo(AnoMes As String, proc As String, mon As String) As Boolean
On Error GoTo noCarga
    Dim i As Long, k As Integer
    Dim RQ As MYSQL_RS
    Dim UsuAceptado As Boolean, Strc As String
    i = 0
    Y = 1
    CargarTareo = False
    UsuAceptado = False
    If cboMon.List(cboMon.ListIndex, 1) = "N" Then
        Strc = " AND b.tipoplani = 'N' "
    Else
        SQL = "select * from autorizaciones where codigo = 1"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF
            If Trim(RQ.Fields("usuario")) = strUsuarioId Then
                UsuAceptado = True
                Exit Do
            End If
            RQ.MoveNext
        Loop
        If UsuAceptado = False Then
            SQL = "select autorizado from rh_tempacceso where codigo = 1 and usuario = '" & strUsuarioId & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If Trim(RQ.Fields("autorizado")) = "S" Then
                    Strc = " AND b.tipoplani = 'S' "
                Else
                    flxTareo.Visible = False
                    If Veces = 0 Then
                        MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                        SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If RQ.EOF Then
                            SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                            oConexionMYSQL.Execute SQL
                        End If
                        Veces = 1
                    Else
                        Veces = 0
                    End If
                    Exit Function
                End If
            Else
                flxTareo.Visible = False
                If Veces = 0 Then
                    MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                    SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If RQ.EOF Then
                        SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                        oConexionMYSQL.Execute SQL
                    End If
                    Veces = 1
                Else
                    Veces = 0
                End If
                Exit Function
            End If
        Else
            Strc = " AND b.tipoplani = 'S' "
        End If
    End If
    SQL = "Select a.anomes,a.tipo as trubro,a.emp,a.rub,a.unidad,a.cant,a.fecha,a.afp,a.sbasico," & _
          " concat(b.apepat,' ' ,b.apemat,' ', b.nombre1,' ', b.nombre2) as nombre, " & _
          " c.descrip,c.tipo,c.actsueldafp,a.codcontrato from (pl_tareo as a left join empleado as b on (a.emp=b.codigo))" & _
          " left join pl_rubrosremunerativos as c on a.rub=c.codigo " & _
          " where a.anomes='" & AnoMes & "' and a.tipo='" & proc & "' " & Strc & " and a.moneda='" & mon & "'" & _
          " order by b.apepat,b.apemat, b.nombre1,b.nombre2,a.unidad,a.tipo,c.tipo,a.rub  "
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        Do While Not (RQ.EOF)
            i = i + 1
            flxTareo.Rows = flxTareo.Rows + 1
            .TextMatrix(i, 1) = RQ.Fields("nombre")
            .TextMatrix(i, 2) = RQ.Fields("tipo")
            .TextMatrix(i, 3) = RQ.Fields("rub") & " " & RQ.Fields("descrip")
            .TextMatrix(i, 4) = RQ.Fields("unidad")
            .TextMatrix(i, 5) = FormatNumber(RQ.Fields("cant"), 2)
            .TextMatrix(i, 6) = RQ.Fields("emp")
            .TextMatrix(i, 7) = RQ.Fields("sbasico")
            .TextMatrix(i, 8) = RQ.Fields("afp")
            .TextMatrix(i, 9) = RQ.Fields("codcontrato")
            DtpFecha = Format(RQ.Fields("fecha"), "dd/mm/yyyy")
            If Trim(RQ.Fields("emp")) = Trim(CodEmpSelect) Then
                If FlgAct = True Then
                    If RQ.Fields("actsueldafp") = "S" Then
                        If proc = 6 Then
                            If RQ.Fields("RUB") = "201" Or RQ.Fields("RUB") = "401" Then
                                .Col = 3
                                .row = i
                                .TextMatrix(i, 5) = ""
                                flxtareo_KeyDown 13, 0
                            End If
                        Else
                            .Col = 3
                            .row = i
                            If RQ.Fields("RUB") <> "117" Then
                                .TextMatrix(i, 5) = ""
                            End If
                            flxtareo_KeyDown 13, 0
                        End If
                    End If
                End If
            End If
            RQ.MoveNext
        Loop
    End With
    Set RQ = Nothing
    If i > 0 Then
        CargarTareo = True
        CargarNuevos AnoMes, proc, mon
    End If
Exit Function
noCarga:
    CargarTareo = False
End Function
Public Sub CargarNuevos(AM As String, pr As String, M As String)
    Dim i As Integer
    cboAdicionales.Clear
    Select Case pr
        Case "4":
            SQL = "Select DISTINCT concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre, " & _
                  " b.codemp,b.sbasico,a.codafp,c.gocehaber,b.codigo from (empleado as a left join contrato as b" & _
                  " on a.codigo=b.codemp) left join calendario as c on(c.codemp=a.codigo) " & _
                  " where a.tipo not in(3,4) and c.movemp='02' and b.mon_sueldo='" & M & "' and b.estado='" & APROBADO & "' and b.codtipo not in ('05','06')" & _
                  " AND B.COdEMP NOt IN (SELECT distinct emp from pl_tareo where anomes='" & AM & "' and tipo='" & pr & "' and moneda='" & M & "' )" & _
                  " and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & AM & "'" & _
                  " and concat(left(fec_regreso,4),substring(fec_regreso,6,2))>='" & AM & "'" & _
                  " order by a.apepat,a.apemat, a.nombre1,a.nombre2"
        Case Else
            SQL = "Select concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre," & _
                  " b.codemp,b.sbasico,a.codafp,b.codigo from empleado as a left join contrato as b" & _
                  " on a.codigo=b.codemp where a.tipo not in (3,4) and b.mon_sueldo='" & M & "' and b.estado='" & APROBADO & "' and b.codtipo not in ('05','06')" & _
                  " AND B.COdEMP NOt IN (SELECT distinct emp from pl_tareo where anomes='" & AM & "' and tipo='" & pr & "' and moneda='" & M & "' )" & _
                  " Union" & _
                  " Select concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre," & _
                  " b.codemp,b.sbasico,a.codafp,b.codigo from empleado as a left join contrato as b" & _
                  " on a.codigo=b.codemp where a.tipo not in (3,4) and b.mon_sueldo='" & M & "' and b.estado='" & CANCELADO & "' and b.codtipo not in ('05','06')" & _
                  " and B.CODIGO=(SELECT MAX(CODIGO) FROM CONTRATO WHERE CODEMP=a.codigo)" & _
                  " and a.situacion='0' and left(fec_cese,7)='" & Left(AM, 4) & "/" & Right(AM, 2) & "'" & _
                  " AND  B.COdEMP NOt IN (SELECT distinct emp from pl_tareo where anomes='" & AM & "' and tipo='" & pr & "' and moneda='" & M & "' )" & _
                  " order by NOMBRE"
    End Select
    i = 0
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    Do While Not Rs.EOF
        i = i + 1
        cboAdicionales.AddItem Rs.Fields("NOMBRE")
        cboAdicionales.List(i - 1, 1) = Rs.Fields("codemp")
        cboAdicionales.List(i - 1, 2) = Rs.Fields("sbasico")
        cboAdicionales.List(i - 1, 3) = Rs.Fields("gocehaber")
        cboAdicionales.List(i - 1, 4) = Rs.Fields("codafp")
        cboAdicionales.List(i - 1, 5) = Rs.Fields("codigo")
        Rs.MoveNext
    Loop
    If i > 0 Then
        cboAdicionales.Enabled = True
        cboAdicionales.ListIndex = 0
        cmdAnexar.Enabled = True
    Else
        cboAdicionales.Enabled = False
        cmdAnexar.Enabled = False
    End If
    Set Rs = Nothing
End Sub
Private Sub PintaFila(fila As Long, valor As Integer)
    Dim i As Integer, Color As Variant
    If valor = 1 Then Color = &HC0FFFF
    If valor = -1 Then Color = &HC0FFC0
    flxTareo.row = fila
    For i = 1 To 5
        flxTareo.Col = i
        flxTareo.CellBackColor = Color
    Next
End Sub
Private Sub chBtnReporte_Click()
    If cboMon.ListIndex > 0 Then
        Set oReporte = New clsReporte
        oReporte.Reporte = "Rep_Tareo.rpt"
        oReporte.empresa = "NATIONAL OILWELL VARCO PERU S.R.L."
        Select Case plani
            Case "1": oReporte.Titulo = "REGISTRO DE TAREO - PLANILLA MENSUAL DE REMUNERACIONES " & NombreMes(Right(Periodo, 2), False) & " - " & Left(Periodo, 4)
            Case "2": oReporte.Titulo = "REGISTRO DE TAREO - PLANILLA QUINCENAL DE REMUNERACIONES " & NombreMes(Right(Periodo, 2), False) & " - " & Left(Periodo, 4)
            Case "4": oReporte.Titulo = "REGISTRO DE TAREO - PLANILLA VACACIONAL DE REMUNERACIONES " & NombreMes(Right(Periodo, 2), False) & " - " & Left(Periodo, 4)
        End Select
        If flxTareo.ColSel = 9 Then
            oReporte.sp_Rep_Tareo cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon, 1), Mid(lblCodEmp, 9, 11)
        Else
           oReporte.sp_Rep_Tareo cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon, 1), ""
        End If
    End If
End Sub
Private Sub chBtnSalir_Click()
    Unload Me
End Sub
Private Sub cmdAnexar_Click()
    Dim aemp As String, asbasico As Double, aventa As String, arub As String, acantidad As Double, afecha As String, auni As String, afp As String
    Dim contrato As String, acantidadc As Double, FlgIng As Boolean, FlgIngc As Boolean
    aemp = cboAdicionales.List(cboAdicionales.ListIndex, 1)
    asbasico = cboAdicionales.List(cboAdicionales.ListIndex, 2)
    aventa = cboAdicionales.List(cboAdicionales.ListIndex, 3)
    afecha = Format(DtpFecha, "yyyy/mm/dd")
    Select Case plani
        Case "1", "2": acantidad = CDbl(DiasTrabajos(aemp, Periodo, plani)) - CDbl(diasvaca(aemp, Periodo, plani)) 'rs.Fields("cantidad")
        Case "4":
            acantidad = CDbl(diasvaca(aemp, Periodo, plani))
            acantidadc = CDbl(diasvacaCompra(aemp, Periodo, plani))
    End Select
    afp = cboAdicionales.List(cboAdicionales.ListIndex, 4)
    contrato = cboAdicionales.List(cboAdicionales.ListIndex, 5)
    If aventa = "S" Then
        arub = "117"
        FlgIng = True
    Else
        arub = "121"
        FlgIngc = True
    End If
    auni = "D"
    SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
          "'" & Periodo & "'," & _
          "'" & plani & "'," & _
          "'" & moneda & "'," & _
          "'" & aemp & "'," & _
          "'" & arub & "'," & _
          "'" & auni & "'," & _
          "" & IIf(arub = "117", acantidadc, acantidad) & "," & _
          "'" & afecha & "'," & asbasico & ",'" & afp & "','" & contrato & "')"
    oConexionMYSQL.Execute SQL
    If acantidad > 0 And acantidadc > 0 Then
        SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
              "'" & Periodo & "'," & _
              "'" & plani & "'," & _
              "'" & moneda & "'," & _
              "'" & aemp & "'," & _
              "'" & IIf(FlgIng = True, "121", "117") & "'," & _
              "'" & auni & "'," & _
              "" & IIf(FlgIng = True, acantidad, acantidadc) & "," & _
              "'" & afecha & "'," & asbasico & ",'" & afp & "','" & contrato & "')"
        oConexionMYSQL.Execute SQL
    End If
   
    cboMon_Change
End Sub
Private Sub CmdCarga_Click()
    FlgCarga = True
    flxTareo.Visible = False
    Screen.MousePointer = vbHourglass
    For i = 1 To CboDescuentos.ListCount - 1
        CboDescuentos.ListIndex = i
        btnCargarD_Click
    Next
    CboDescuentos.ListIndex = 0
     For i = 1 To cboAportes.ListCount - 1
        cboAportes.ListIndex = i
        btnCargarA_Click
    Next
    cboAportes.ListIndex = 0
    Screen.MousePointer = vbDefault
    flxTareo.Visible = True
    FlgCarga = False
End Sub
Private Sub flxTareo_Click()
    If FlgBloq = False Then
        If flxTareo.ColSel = 9 And btnEliminar.Enabled = True And empselect <> "" Then
            btnEliminaEmp.Enabled = True
        Else
            btnEliminaEmp.Enabled = False
        End If
    End If
End Sub
Private Sub flxTareo_DblClick()
    If EstadoPlani(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), moneda, plani) = True Then
        Exit Sub
    Else
        Select Case flxTareo.Col
            Case 2
                If flxTareo.TextMatrix(flxTareo.row, 5) = "" Or flxTareo.TextMatrix(flxTareo.row, 5) = "0.00" Then
                    LlenaTRubro
                    cbMvto.ListIndex = EstadoComboTRubro(flxTareo.TextMatrix(flxTareo.row, 2))
                    With cbMvto
                        .Top = flxTareo.CellTop + flxTareo.Top
                        .Left = flxTareo.CellLeft + flxTareo.Left
                        .Width = flxTareo.CellWidth + 3500
                        .Visible = True
                        .SetFocus
                    End With
                End If
            Case 3
                If flxTareo.TextMatrix(flxTareo.row, 2) <> "" And flxTareo.TextMatrix(flxTareo.row, 5) = "" Or flxTareo.TextMatrix(flxTareo.row, 5) = "0.00" Then
                    LlenaRubros flxTareo.TextMatrix(flxTareo.row, 2)
                    cbMvto.ListIndex = EstadoCombo(flxTareo.TextMatrix(flxTareo.row, 3))
                    With cbMvto
                        .Top = flxTareo.CellTop + flxTareo.Top
                        .Left = flxTareo.CellLeft + flxTareo.Left
                        .Width = flxTareo.CellWidth '+ 3500
                        .Visible = True
                        .SetFocus
                    End With
                End If
            Case Else
                cbMvto.Visible = False
        End Select
    End If
End Sub
Private Sub flxtareo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsrub As MYSQL_RS
    Dim rub As String
    Dim MtoPA As Double
    Dim nRegistros As Long
    
    If EstadoPlani(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), moneda, plani) = True Then
        Exit Sub
    Else
        If FlgBloq = False Then
            If KeyCode = 46 Then
            
                If MsgBox("¿ Está usted seguro de eliminar la linea actual?", vbQuestion + vbYesNo + vbDefaultButton2, "Atención") = vbYes Then
            
                SQL = "Delete from pl_tareo where anomes='" & Periodo & "' and tipo='" & plani & _
                      "' and moneda='" & moneda & "' and rub='" & Trim(Left(Rubro, 4)) & _
                      "' and emp='" & Mid(lblCodEmp, 9, 11) & "'"
                oConexionMYSQL.Execute SQL, nRegistros
                
                ActualizaMontos Periodo, plani, moneda, IIf(Trim(Left(Rubro, 4)) = "709", "P", "A"), Mid(lblCodEmp, 9, 11), 1
                
                '--------
                If nRegistros > 0 Then
                    flxTareo.BorrarFilas (False)
                End If
                
                End If
            End If
            
            With flxTareo
                If .Col = 4 Then
                    Publimensaje = "modificar"
                    If .Col = 4 Then TipodeCampo = cadena
                    Exit Sub
                End If
                Dim FlgIng As Boolean
                FlgIng = False
                If .Col = 3 Then
                    rub = .TextMatrix(.row, 3)
                    If rub <> "" And .TextMatrix(.row, 5) = "" And KeyCode = 13 Then
                        SQL = "SELECT * FROM PL_RUBROSREMUNERATIVOS WHERE CODIGO='" & Trim(Left(rub, 4)) & "'"
                        Set rsrub = oConexion.EjecutaSelectRS(SQL)
                        If Not (rsrub.EOF And rsrub.BOF) Then
                            If IsNull(rsrub.Fields("formula")) = False Then
                                If Trim(Left(rub, 4)) <> "117" Then
                                    .TextMatrix(.row, 4) = "V"
                                    .TextMatrix(.row, 5) = FormatNumber(CargaRubro(rsrub.Fields("formula"), .TextMatrix(.row, 6), CDbl(.TextMatrix(.row, 7))), 2)
                                    MtoPA = .TextMatrix(.row, 5)
                                    If Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "709" Or Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "701" Then
                                        ActualizaMontos Periodo, plani, moneda, IIf(Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "709", "P", "A"), .TextMatrix(.row, 6)
                                    End If
                                End If
                                .Col = 5
                            End If
                        End If
                    End If
                End If
                If .Col = 5 Then
                    
                    If .Col = 5 Then TipodeCampo = Numero
                    rub = .TextMatrix(.row, 3)
                    
                    Publimensaje = "modificar"
                    
                    If Left(rub, 3) = "709" Then
                        'Publimensaje = "sin-editar"
                    Else
                        Publimensaje = "modificar"
                    End If
                    
                    
                    If rub <> "" And .TextMatrix(.row, 5) <> "" And KeyCode = 13 Then
                        SQL = "SELECT * FROM PL_TAREO WHERE ANOMES='" & Periodo & "' AND TIPO='" & plani & _
                              "' AND MONEDA='" & moneda & "' AND RUB='" & Trim(Left(rub, 4)) & "' and emp='" & .TextMatrix(.row, 6) & "'"
                        Set rsrub = oConexion.EjecutaSelectRS(SQL)
                        If rsrub.EOF And rsrub.BOF Then
                            SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,unidad,cant,fecha,sbasico,afp,codcontrato) values (" & _
                                  "'" & Periodo & "'," & _
                                  "'" & plani & "'," & _
                                  "'" & moneda & "'," & _
                                  "'" & .TextMatrix(.row, 6) & "'," & _
                                  "'" & Trim(Left(.TextMatrix(.row, 3), 4)) & "'," & _
                                  "'" & .TextMatrix(.row, 4) & "'," & _
                                  "" & CDbl(.TextMatrix(.row, 5)) & "," & _
                                  "'" & Format(DtpFecha, "yyyy/mm/dd") & "'," & CDbl(.TextMatrix(.row, 7)) & ",'" & .TextMatrix(.row, 8) & "','" & .TextMatrix(.row, 9) & "')"
                        Else
                            SQL = " Update pl_tareo set rub='" & Trim(Left(.TextMatrix(.row, 3), 4)) & "', unidad='" & .TextMatrix(.row, 4) & "', cant=" & CDbl(.TextMatrix(.row, 5)) & _
                                  " where anomes='" & Periodo & "' and tipo='" & plani & "' and moneda='" & moneda & "' and emp='" & .TextMatrix(.row, 6) & "' and rub='" & Trim(Left(rub, 4)) & "'"
                        End If
                        oConexionMYSQL.Execute SQL
                        
                        If Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "709" Or Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "701" Then
                            ActualizaMontosG Periodo, plani, moneda, Trim(.TextMatrix(.row, 6)), CDbl(.TextMatrix(.row, 5)), IIf(Mid(Trim(.TextMatrix(.row, 3)), 1, 3) = "709", "P", "A")
                        End If
                    End If
                End If
            End With
        End If
    End If
    Set rsrub = Nothing
End Sub
Sub ActualizaMontosG(AnoMes As String, plani As String, mon As String, emp As String, Mto As Double, Tipo As String)
    Dim SQL As String
    Dim RQ As MYSQL_RS
    SQL = "SELECT MONTOABONADO AS DESCONTADO,DOCUMENTO,(select anomes from rh_divprestadel where anomesdesc= '" & AnoMes & "' " & _
          "and tipoplanidesc='" & plani & "' and tipo= '" & Tipo & "' and solicita = '" & emp & "') AS AMS, " & _
          "(select TIPOPLANI from rh_divprestadel where anomesdesc= '" & AnoMes & "' " & _
          "and tipoplanidesc='" & plani & "' and tipo= '" & Tipo & "' and solicita = '" & emp & "') AS TPL,solicita " & _
          "FROM RH_PAGOSEMP WHERE ANOMESPLANI = '" & AnoMes & "' AND TIPOPLANI = '" & plani & "' AND TIPO = '" & Tipo & "' AND MONPLANI = '" & mon & "' AND SOLICITA = '" & emp & "' " & _
          "order by solicita,documento,anomesplani"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    Do While Not RQ.EOF()
        SQL = "UPDATE documentos_rrhh set montopagado = (montopagado - " & RQ.Fields("descontado") & ") + (" & Mto & ") where solicita = '" & emp & "' " & _
              "and documento = '" & Trim(RQ.Fields("documento")) & "' and coddoc = '" & IIf(Tipo = "P", "PR", "SS") & "'"
        oConexionMYSQL.Execute SQL
        SQL = "UPDATE rh_divprestadel set descontado = (descontado - " & RQ.Fields("descontado") & ") + (" & Mto & ") where solicita = '" & emp & "' " & _
              "and documento = '" & Trim(RQ.Fields("documento")) & "' and tipo = '" & Tipo & "' and anomes = '" & Trim(RQ.Fields("AMS")) & "' and tipoplani = '" & Trim(RQ.Fields("TPL")) & "'"
        oConexionMYSQL.Execute SQL
        SQL = "UPDATE rh_divprestadel set pagado = if(monto=descontado,'S','N'),anomesdesc = if(" & Mto & "= " & RQ.Fields("descontado") & ",'',anomesdesc),tipoplanidesc = if(" & Mto & "= " & RQ.Fields("descontado") & ",'',tipoplanidesc) where solicita = '" & emp & "' " & _
              "and documento = '" & Trim(RQ.Fields("documento")) & "' and tipo = '" & Tipo & "' and anomes = '" & Trim(RQ.Fields("AMS")) & "' and tipoplani = '" & Trim(RQ.Fields("TPL")) & "'"
        oConexionMYSQL.Execute SQL
        SQL = "update rh_pagosemp set montoabonado = (montoabonado - " & RQ.Fields("descontado") & ") + " & Mto & " where solicita = '" & emp & "' and anomesplani = '" & AnoMes & "' and tipoplani = '" & plani & "' " & _
              "and monplani = '" & mon & "' and tipo = '" & Tipo & "' and documento = '" & Trim(RQ.Fields("documento")) & "'"
        oConexionMYSQL.Execute SQL
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub
Private Sub flxtareo_KeyPress(KeyAscii As Integer)
    Dim F As Integer
    If KeyAscii = 13 Then
        If EstadoPlani(cboAnio.List(cboAnio.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), moneda, plani) = True Then
            With flxTareo
                If .Col = 5 Then
                    If .TextMatrix(.row, 5) <> "" Then .TextMatrix(.row, 5) = Trim(.TextMatrix(.row, 5))
                End If
            End With
        End If
    Else
        If flxTareo.Col < 4 Then
            Dim c%, T%, a$, B$
            If KeyAscii = 0 Then
                cadenaemp = "": lblCadBus = ""
                Exit Sub
            End If
            If KeyAscii >= 32 Or KeyAscii = 8 Then
                cadenaemp = cadenaemp & Chr(KeyAscii)
                lblCadBus = cadenaemp
                With flxTareo
                    If KeyAscii <> 8 Then
                        c = Len(cadenaemp)
                        If IsNumeric(cadenaemp) Then
                            a = Right("00000000000" & Trim(cadenaemp), 11)
                            cadenaemp = a
                            lblCadBus = cadenaemp
                            c = 11
                        Else
                            a = cadenaemp
                        End If
                    End If
                    If c >= 1 Then
                        For T = 1 To .Rows - 1
                            If IsNumeric(cadenaemp) Then
                                B = .TextMatrix(T, 6)
                            Else
                                B = .TextMatrix(T, 1)
                            End If
                            If Len(B) >= c Then
                                B = Left(B, c)
                                If Trim(a) = Trim(B) Then
                                    KeyAscii = 0
                                    ItemLista = T
                                    .row = T
                                    .Col = 1
                                    If T >= 6 Then
                                        .TopRow = T - 5
                                    Else
                                        .TopRow = T
                                    End If
                                    flxTareo_RowColChange
                                    Exit For
                                End If
                            End If
                        Next T
                    End If
                End With
            End If
        End If
    End If
End Sub
Private Sub flxTareo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FlgBloq = False Then
        If Button = vbRightButton And flxTareo.TextMatrix(flxTareo.row, 0) <> "" Then Call btnNuevaFila_Click
    End If
End Sub
Private Sub flxTareo_RowColChange()
    empselect = ""
    If cargando = False Then
        If flxTareo.row > 0 And flxTareo.Col > 0 Then
            If flxTareo.TextMatrix(flxTareo.row, 0) <> "" Then coloremp = flxTareo.CellBackColor
            If flxTareo.Col <> 5 Then
                Publimensaje = "sin-editar"
                TipodeCampo = cadena
            End If
            
            If flxTareo.Col = 5 Then
                If Left(flxTareo.TextMatrix(flxTareo.row, 3), 3) = "709" Then
                    'Publimensaje = "sin-editar"
                Else
                    Publimensaje = "modificar"
                End If
            End If
            
            lblCodEmp = "CODIGO: " & Trim(flxTareo.TextMatrix(flxTareo.row, 6)) & "  -  S. Básico: " & flxTareo.TextMatrix(flxTareo.row, 7)
            empselect = Trim(flxTareo.TextMatrix(flxTareo.row, 6))
            lblPension = DescripcionesdeCodigos("AFP", flxTareo.TextMatrix(flxTareo.row, 8), "Descrip")
            Rubro = flxTareo.TextMatrix(flxTareo.row, 3)
            gsRubro = Rubro
            DoEvents
            
            
            
            If flxTareo.Col = 6 Then
                flxTareo.Col = 5
                flxtareo_KeyDown 13, 0
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    'Me.WindowState = vbMaximized
    DoEvents
    Call WheelHook(frmTareoPlani)
    Me.Top = 0
    Me.Left = 0
    cadenaemp = ""
    lblCadBus = ""
    LlenarMesP cboMes
    LlenarMonedaP cboMon
    LlenarAnio
    cboMon.AddItem "CONFIDENCIAL"
    cboMon.List(3, 1) = "C"
    LlenarProcesos cboProceso
    TareosAutomaticos
    TareosAutomaticosD
    TareosAutomaticosA
    DtpFecha.Text = Format(CStr(Date), "dd/mm/yyyy")
    Publimensaje = "sin-editar"
    SSTab1.Tab = 0
    FlgRubAl = False
    FlgBloq = False
    If Bloqueo Then
        FlgBloq = True
    End If
    If FlgBloq = True Then
        btnCargarA.Enabled = False: btnCargarD.Enabled = False: btnCargarI.Enabled = False
        btnDesCargarA.Enabled = False: btnDesCargarD.Enabled = False: btnDesCargarI.Enabled = False
        btnActSueldo.Enabled = False: btnImportar.Enabled = False: CmdCarga.Enabled = False
        cboIngresos.Enabled = False: CboDescuentos.Enabled = False: cboAportes.Enabled = False
        cmdAnexar.Enabled = False: cboAdicionales.Enabled = False: btnEliminaEmp.Enabled = False
        btnEliminar.Enabled = False: btnGrabar.Enabled = False
    End If
    
    
    Me.flxTareo.SelectionMode = flexSelectionFree
End Sub
Function Bloqueo() As Boolean
    Bloqueo = False
    If cboAnio.List(cboAnio.ListIndex, 1) < Year(Date) Then Bloqueo = True
    If cboAnio.List(cboAnio.ListIndex, 1) = Year(Date) Then
        If val(cboMes.List(cboMes.ListIndex, 2)) = Month(Date) - 1 Then
            If Day(Date) > 5 Then Bloqueo = True
        Else
            If val(cboMes.List(cboMes.ListIndex, 2)) < Month(Date) - 1 Then
                Bloqueo = True
            End If
        End If
    End If
    If Mid(cboProceso.List(cboProceso.ListIndex, 0), 1, 1) = 6 Then
        Bloqueo = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set oReporte = Nothing
    BorrarSolicitud
End Sub
Sub BorrarSolicitud()
    SQL = "delete from rh_tempacceso where usuario = '" & strUsuarioId & "' and codigo = 1"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        btnNuevaFila.Enabled = True
        btnBorrar.Enabled = True
        btnSubir.Enabled = True
        btnBajar.Enabled = True
        cbMvto.Enabled = True
    End If
End Sub
Public Sub TareosAutomaticos()
    Dim i As Integer
    i = 1
    SQL = "SELECT A.CODIGO,A.TIPO,A.DESCRIP FROM pl_rubrosremunerativos AS A" & _
          " LEFT JOIN pl_tiporubros AS B ON (A.TIPO=B.CODIGO)" & _
          " WHERE B.TIPO='I' AND A.CALCULO='A' ORDER BY A.DESCRIP"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With cboIngresos
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "0000"
        Do While Not (Rs.EOF)
            .AddItem Rs.Fields("CODIGO") & " - " & Rs.Fields("DESCRIP")
            .List(i, 1) = Rs.Fields("TIPO")
            .List(i, 2) = Rs.Fields("CODIGO")
            i = i + 1
            Rs.MoveNext
        Loop
        .ListIndex = 0
    End With
    Set Rs = Nothing
End Sub
Public Sub TareosAutomaticosD()
    Dim i As Integer
    i = 1
    SQL = "SELECT A.CODIGO,A.TIPO,A.DESCRIP FROM pl_rubrosremunerativos AS A" & _
          " LEFT JOIN pl_tiporubros AS B ON (A.TIPO=B.CODIGO)" & _
          " WHERE B.TIPO='D' AND A.CALCULO='A' ORDER BY A.DESCRIP"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With CboDescuentos
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "0000"
        Do While Not (Rs.EOF)
            .AddItem Rs.Fields("CODIGO") & " - " & Rs.Fields("DESCRIP")
            .List(i, 1) = Rs.Fields("TIPO")
            .List(i, 2) = Rs.Fields("CODIGO")
            i = i + 1
            Rs.MoveNext
        Loop
        .ListIndex = 0
    End With
    Set Rs = Nothing
End Sub
Public Sub TareosAutomaticosA()
    Dim i As Integer
    i = 1
    SQL = "SELECT A.CODIGO,A.TIPO,A.DESCRIP FROM pl_rubrosremunerativos AS A" & _
          " LEFT JOIN pl_tiporubros AS B ON (A.TIPO=B.CODIGO)" & _
          " WHERE B.TIPO='A' AND A.CALCULO='A' ORDER BY A.DESCRIP"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With cboAportes
        .Clear
        .AddItem "Seleccionar..."
        .List(0, 1) = "0000"
        Do While Not (Rs.EOF)
            .AddItem Rs.Fields("CODIGO") & " - " & Rs.Fields("DESCRIP")
            .List(i, 1) = Rs.Fields("TIPO")
            .List(i, 2) = Rs.Fields("CODIGO")
            i = i + 1
            Rs.MoveNext
        Loop
        .ListIndex = 0
    End With
    Set Rs = Nothing
End Sub
Public Sub LlenarEmpleados(mon As String, tplani As String)
    Dim i As Integer
    Dim str As String
    Dim UsuAceptado As Boolean, Strc As String
    i = 0
    cbMvto.Clear
    UsuAceptado = False
    If cboMon.List(cboMon.ListIndex, 1) = "N" Then
        Strc = " AND a.tipoplani = 'N' "
    Else
        SQL = "select * from autorizaciones where codigo = 1"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF
            If Trim(RQ.Fields("usuario")) = strUsuarioId Then
                UsuAceptado = True
                Exit Do
            End If
            RQ.MoveNext
        Loop
        If UsuAceptado = False Then
            SQL = "select autorizado from rh_tempacceso where codigo = 1 and usuario = '" & strUsuarioId & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If Trim(RQ.Fields("autorizado")) = "S" Then
                    Strc = " AND a.tipoplani = 'S' "
                Else
                    flxTareo.Visible = False
                    If Veces = 0 Then
                        MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                        SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If RQ.EOF Then
                            SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                            oConexionMYSQL.Execute SQL
                        End If
                        Veces = 1
                    Else
                        Veces = 0
                    End If
                    Exit Sub
                End If
            Else
                flxTareo.Visible = False
                If Veces = 0 Then
                    MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                    SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If RQ.EOF Then
                        SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                        oConexionMYSQL.Execute SQL
                    End If
                    Veces = 1
                Else
                    Veces = 0
                End If
                Exit Sub
            End If
        Else
            Strc = " AND a.tipoplani = 'S' "
        End If
    End If
    If tplani = "5" Then
        SQL = "Select concat(a.apepat,' ',a.apemat,' ',a.nombre1,' ',a.nombre2) as nombre," & _
              " b.codemp,b.sbasico,a.codafp,b.codigo from empleado as a left join contrato as b " & _
              " on a.codigo=b.codemp where a.tipo not in (3,4) and b.mon_sueldo='" & mon & "' and a.situacion='1' " & Strc & " and b.estado='" & APROBADO & "' and b.codtipo not in ('05','06')" & _
              " order by a.apepat,a.apemat, a.nombre1,a.nombre2"
    Else
        SQL = "Select concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre," & _
              " b.codemp,b.sbasico,a.codafp,b.codigo from empleado as a left join contrato as b" & _
              " on a.codigo=b.codemp where a.tipo not in (3,4) and b.mon_sueldo='" & mon & "' " & Strc & " and b.estado='" & APROBADO & "' and b.codtipo not in ('05','06')" & _
              " Union" & _
              " Select concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre," & _
              " b.codemp,b.sbasico,a.codafp,b.codigo from empleado as a left join contrato as b" & _
              " on a.codigo=b.codemp where a.tipo not in (3,4) and b.mon_sueldo='" & mon & "' " & Strc & " and b.estado='" & CANCELADO & "' and b.codtipo not in ('05','06')" & _
              " and B.CODIGO=(SELECT MAX(CODIGO) FROM CONTRATO WHERE CODEMP=a.codigo)" & _
              " and a.situacion='0' and left(fec_cese,7)='" & Left(Periodo, 4) & "/" & Right(Periodo, 2) & "'" & _
              " order by NOMBRE"
    End If
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        .Visible = False
        Do While Not (Rs.EOF)
            cbMvto.AddItem Rs.Fields("nombre")
            .Rows = .Rows + 1
            For k = 1 To 5
                .Col = k
                .row = i + 1
                If i Mod 2 = 0 Then
                    .CellBackColor = &HC0FFFF
                Else
                    .CellBackColor = &HC0FFC0
                End If
            Next
            .TextMatrix(i + 1, 0) = i + 1
            .TextMatrix(i + 1, 1) = Rs.Fields("nombre")
            .TextMatrix(i + 1, 6) = Rs.Fields("codemp")
            .TextMatrix(i + 1, 7) = Rs.Fields("sbasico")
            .TextMatrix(i + 1, 8) = Rs.Fields("codafp")
            .TextMatrix(i + 1, 9) = Rs.Fields("codigo")
            Rs.MoveNext
            i = i + 1
        Loop
        .Visible = True
    End With
    Set Rs = Nothing
End Sub
Public Sub LlenarEmpVac(mon As String)
    Dim i As Integer
    Dim UsuAceptado As Boolean, Strc As String
    i = 0
    cbMvto.Clear
    UsuAceptado = False
    If cboMon.List(cboMon.ListIndex, 1) = "N" Then
        Strc = " AND a.tipoplani = 'N' "
    Else
        SQL = "select * from autorizaciones where codigo = 1"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF
            If Trim(RQ.Fields("usuario")) = strUsuarioId Then
                UsuAceptado = True
                Exit Do
            End If
            RQ.MoveNext
        Loop
        If UsuAceptado = False Then
            SQL = "select autorizado from rh_tempacceso where codigo = 1 and usuario = '" & strUsuarioId & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                If Trim(RQ.Fields("autorizado")) = "S" Then
                    Strc = " AND a.tipoplani = 'S' "
                Else
                    flxTareo.Visible = False
                    If Veces = 0 Then
                        MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                        SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                        Set RQ = oConexion.EjecutaSelectRS(SQL)
                        If RQ.EOF Then
                            SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                            oConexionMYSQL.Execute SQL
                        End If
                        Veces = 1
                    Else
                        Veces = 0
                    End If
                    Exit Sub
                End If
            Else
                flxTareo.Visible = False
                If Veces = 0 Then
                    MsgBox "Usted no se encuentra autorizado a visualizar esta planilla. Solicite una autorización", vbInformation, "NOVPeru"
                    SQL = "SELECT * FROM rh_tempacceso WHERE USUARIO = '" & strUsuarioId & "' AND CODIGO = 1"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If RQ.EOF Then
                        SQL = "insert into rh_tempacceso(codigo,usuario,autorizado) values(1,'" & strUsuarioId & "','N')"
                        oConexionMYSQL.Execute SQL
                    End If
                    Veces = 1
                Else
                    Veces = 0
                End If
                Exit Sub
            End If
        Else
            Strc = " AND a.tipoplani = 'S' "
        End If
    End If
    SQL = "Select DISTINCT concat(a.apepat,' ' ,a.apemat,' ', a.nombre1,' ', a.nombre2) as nombre," & _
          " b.codemp,b.sbasico,c.gocehaber,a.codafp,b.codigo from (empleado as a left join contrato as b " & _
          " on a.codigo=b.codemp) left join calendario as c on(c.codemp=a.codigo) " & _
          " where a.tipo not in (3,4) and c.movemp='02' and b.mon_sueldo='" & mon & "' " & Strc & " and b.estado='" & APROBADO & "'" & _
          " and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & Periodo & "'" & _
          " and concat(left(fec_regreso,4),substring(fec_regreso,6,2))>='" & Periodo & "'" & _
          " order by a.apepat,a.apemat, a.nombre1,a.nombre2"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        .Visible = False
        Do While Not (Rs.EOF)
            cbMvto.AddItem Rs.Fields("nombre")
            .Rows = .Rows + 1
            For k = 1 To 5
                .Col = k
                .row = i + 1
                If i Mod 2 = 0 Then
                    .CellBackColor = &HC0FFFF
                Else
                    .CellBackColor = &HC0FFC0
                End If
            Next
            .TextMatrix(i + 1, 0) = i + 1
            .TextMatrix(i + 1, 1) = Rs.Fields("nombre")
            .TextMatrix(i + 1, 6) = Rs.Fields("codemp")
            .TextMatrix(i + 1, 7) = Rs.Fields("sbasico")
            .TextMatrix(i + 1, 8) = Rs.Fields("codafp")
            .TextMatrix(i + 1, 9) = Rs.Fields("codigo")
            Rs.MoveNext
            i = i + 1
        Loop
        .Visible = True
    End With
    Set Rs = Nothing
End Sub
Public Sub LlenaTRubro()
    Dim i As Integer
    cbMvto.Clear
    cbMvto.AddItem "Seleccionar..."
    SQL = "Select * from pl_tiporubros order by codigo"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        Do While Not (Rs.EOF)
            cbMvto.AddItem Rs.Fields("codigo") & " " & Rs.Fields("descrip")
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Sub
Public Sub LlenaRubros(Tipo As String)
    Dim i As Integer
    cbMvto.Clear
    cbMvto.AddItem "Seleccionar..."
    SQL = "Select * from pl_rubrosremunerativos where principal='S' and  tipo='" & Tipo & "'  order by descrip"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        Do While Not (Rs.EOF)
            cbMvto.AddItem Rs.Fields("codigo") & " " & Rs.Fields("descrip")
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Sub
Public Function EstadoComboTRubro(Texto As String) As Integer
    Dim i As Integer
    For i = 0 To cbMvto.ListCount - 1
        If Left(Trim(cbMvto.List(i)), 4) = Left(Trim(Texto) & "    ", 4) Then
            EstadoComboTRubro = i
            Exit Function
        End If
    Next
    EstadoComboTRubro = 0
End Function
Public Function EstadoCombo(Texto As String) As Integer
    Dim i As Integer
    For i = 0 To cbMvto.ListCount - 1
        If Trim(cbMvto.List(i)) = Trim(Texto) Then
            EstadoCombo = i
            Exit Function
        End If
    Next
    EstadoCombo = 0
End Function
Private Sub CompletarPlantilla(Tipo As String, nemp As Long, AnoMes As String)
    Dim i As Integer, k As Integer
    Select Case Tipo
        Case 1, 2, 5, 6:
            SQL = "Select a.codigo,a.tiprub,a.rub,a.unidad,a.cantidad,b.descrip from pl_plantilla_plani as a" & _
                  " left join pl_rubrosremunerativos as b on (a.rub=b.codigo) where a.codigo='" & Tipo & "' order by tiprub,rub"
        Case 4
            SQL = "Select a.codigo,a.tiprub,a.rub,a.unidad,a.cantidad,b.descrip from pl_plantilla_plani as a" & _
                  " left join pl_rubrosremunerativos as b on (a.rub=b.codigo) where a.codigo='" & Tipo & "' order by tiprub,rub"
    End Select
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    With flxTareo
        i = 1
        Do While Not (Rs.EOF)
            If i = 1 Then
                For k = 1 To nemp
                    .TextMatrix(k, 2) = Rs.Fields("tiprub")
                    .TextMatrix(k, 3) = Rs.Fields("rub") & " " & Rs.Fields("descrip")
                    .TextMatrix(k, 4) = Rs.Fields("unidad")
                    Select Case plani
                        Case "1", "2": .TextMatrix(k, 5) = FormatNumber(DiasTrabajos(.TextMatrix(k, 6), AnoMes, plani) - diasvaca(.TextMatrix(k, 6), Periodo, Tipo), 2) 'rs.Fields("cantidad")
                        Case "4": .TextMatrix(k, 5) = FormatNumber(diasvaca(.TextMatrix(k, 6), Periodo, Tipo), 2)
                        Case "5": .TextMatrix(k, 5) = FormatNumber(DiasTrabajosGRATI(.TextMatrix(k, 6), AnoMes, plani), 2)  'rs.Fields("cantidad")
                        Case "6": .TextMatrix(k, 5) = FormatNumber(DiasTrabajosCTS(.TextMatrix(k, 6), AnoMes, plani), 2)  'rs.Fields("cantidad")
                    End Select
                Next
            End If
            If i > 1 Then
            End If
            i = i + 1
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Sub
Public Function DiasTrabajos(emp As String, AnoMes As String, plani As String) As Integer
On Error GoTo edias
    Dim rsdt As MYSQL_RS
    Dim PDia As String, UDia As String, fechainiactual As String, flag As Boolean
    DiasTrabajos = 0
    If plani = "2" Then
        UDia = Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/16"
    Else
        UDia = Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/31"
    End If
    PDia = Left(AnoMes, 4) & "/" & Right(AnoMes, 2) & "/01"
    SQL = "Select codigo,estado,codemp,f_termino,f_inicio,fechacese, (select fec_cese from empleado where codigo='" & emp & "') as fcese " & _
          " from contrato where codemp='" & emp & "'" & _
          " ORDER BY CODIGO DESC"
    empant = ""
    Set rsdt = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsdt.EOF
        If empant <> rsdt.Fields("codemp") Then
            empant = rsdt.Fields("codemp")
            flag = False
        End If
        If Trim(rsdt.Fields("estado")) = APROBADO Then
            If Trim(rsdt.Fields("f_termino")) < UDia And Trim(rsdt.Fields("f_termino")) <> "" Then
                If Trim(rsdt.Fields("f_inicio")) <= PDia Then
                    If Right(AnoMes, 2) = "02" And DiasDelMes(AnoMes) = val(Right(Trim(rsdt.Fields("f_termino")), 2)) Then
                        DiasTrabajos = 30
                    Else
                        DiasTrabajos = CDate(rsdt.Fields("f_termino")) - CDate(PDia) + 1
                    End If
                Else
                    DiasTrabajos = CDate(rsdt.Fields("f_termino")) - CDate(Trim(rsdt.Fields("f_inicio"))) + 1
                End If
            Else
                If Trim(rsdt.Fields("f_inicio")) >= PDia Then
                    If plani = "2" Then
                        DiasTrabajos = 15 - (CDate(rsdt.Fields("f_inicio")) - CDate(PDia))
                    Else
                        DiasTrabajos = 30 - (CDate(rsdt.Fields("f_inicio")) - CDate(PDia))
                    End If
                Else
                    If plani = "2" Then
                        DiasTrabajos = 15
                    Else
                        DiasTrabajos = 30
                    End If
                End If
            End If
            flag = True
        Else
            If Trim(rsdt.Fields("estado")) = CANCELADO Then
                If Trim(rsdt.Fields("f_termino")) > PDia And Trim(rsdt.Fields("f_termino")) <> "" Then
                    If Trim(rsdt.Fields("f_INICIO")) >= PDia Then
                        DiasTrabajos = DiasTrabajos + (CDate(IIf(rsdt.Fields("fechacese") <> "", rsdt.Fields("fechacese"), rsdt.Fields("f_termino"))) - CDate(rsdt.Fields("f_inicio")))
                    Else
                        DiasTrabajos = DiasTrabajos + (CDate(IIf(rsdt.Fields("fechacese") <> "", CStr(rsdt.Fields("fechacese")), CStr(rsdt.Fields("f_termino")))) - CDate(PDia)) + 1
                    End If
                End If
            End If
        End If
        rsdt.MoveNext
     Loop
     If DiasTrabajos < 0 Then DiasTrabajos = 0
     Set rsdt = Nothing
Exit Function
edias:
    DiasTrabajos = 0
End Function
Public Function DiasTrabajosCTS(emp As String, AnoMes As String, plani As String) As Integer
On Error GoTo edias
    Dim rsdt As MYSQL_RS
    Dim FechaIni As String, FechaFin As String, flag As Boolean
    DiasTrabajosCTS = 0
    
    SQL = "Select fec_ingreso " & _
          " from empleado where codigo='" & emp & "'"
    empant = ""
    Set rsdt = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsdt.EOF
        Select Case Right(AnoMes, 2)
            Case "04":
                FechaIni = CStr(val(Left(AnoMes, 4)) - 1) & "/10/30"
                FechaFin = Left(AnoMes, 4) & "/04/30"
                If rsdt.Fields("fec_ingreso") >= FechaIni Then
                    If rsdt.Fields("fec_ingreso") >= Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/11/01" Then
                        DiasTrabajosCTS = CDate(Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/11/30") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 150
                    End If
                    If rsdt.Fields("fec_ingreso") > Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/11/30" And rsdt.Fields("fec_ingreso") <= Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/12/31" Then
                        DiasTrabajosCTS = CDate(Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/12/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 120
                    End If
                    If rsdt.Fields("fec_ingreso") > Trim(cboAnio.List(cboAnio.ListIndex, 1) - 1) & "/12/30" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/01/31" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/01/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 90
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/01/31" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/02/" & CStr(DiasDelMes(cboAnio.List(cboAnio.ListIndex, 1) & "02")) Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/02/" & CStr(DiasDelMes(cboAnio.List(cboAnio.ListIndex, 1) & "02"))) - CDate(rsdt.Fields("fec_ingreso")) + 1 + 60
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/02/" & CStr(DiasDelMes(cboAnio.List(cboAnio.ListIndex, 1) & "02")) And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/03/31" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/03/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 30
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/03/31" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/04/30" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/04/30") - CDate(rsdt.Fields("fec_ingreso")) + 1
                    End If
                Else
                    DiasTrabajosCTS = 180
                End If
            Case "10":
                FechaIni = Left(AnoMes, 4) & "/04/30"
                FechaFin = Left(AnoMes, 4) & "/10/30"
                If rsdt.Fields("fec_ingreso") >= FechaIni Then
                    If rsdt.Fields("fec_ingreso") >= cboAnio.List(cboAnio.ListIndex, 1) & "/05/01" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/05/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 150
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/05/31" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/06/30" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/06/30") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 120
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/06/30" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/07/31" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/07/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 90
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/07/31" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/08/31" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/08/31") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 60
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/08/31" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/09/30" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/09/30") - CDate(rsdt.Fields("fec_ingreso")) + 1 + 30
                    End If
                    If rsdt.Fields("fec_ingreso") > cboAnio.List(cboAnio.ListIndex, 1) & "/09/30" And rsdt.Fields("fec_ingreso") <= cboAnio.List(cboAnio.ListIndex, 1) & "/10/30" Then
                        DiasTrabajosCTS = CDate(cboAnio.List(cboAnio.ListIndex, 1) & "/10/30") - CDate(rsdt.Fields("fec_ingreso")) + 1
                    End If
                Else
                    DiasTrabajosCTS = 180
                End If
        End Select
        rsdt.MoveNext
    Loop
    If DiasTrabajosCTS < 0 Then DiasTrabajosCTS = 0
    If DiasTrabajosCTS > 180 Then DiasTrabajosCTS = 180
    Set rsdt = Nothing
Exit Function
edias:
    DiasTrabajosCTS = 0
End Function
Public Function DiasTrabajosGRATI(emp As String, AnoMes As String, plani As String) As Integer
On Error GoTo edias
    Dim rsdt As MYSQL_RS
    Dim FechaIni As String, FechaFin As String, flag As Boolean, Mes As String, Fingreso
    DiasTrabajosGRATI = 0
    Mes = Right(AnoMes, 2)
    Select Case Mes
        Case "07":
            FechaIni = CStr(val(Left(AnoMes, 4)) - 1) & "/12/30"
            FechaFin = Left(AnoMes, 4) & "/07/30"
        Case "12":
            FechaIni = Left(AnoMes, 4) & "/06/30"
            FechaFin = Left(AnoMes, 4) & "/12/30"
    End Select
    SQL = "Select fec_ingreso " & _
          " from empleado where codigo='" & emp & "'"
    Fingreso = FechaPersonal(emp)
    If Fingreso >= FechaIni Then
        If Fingreso > Left(FechaIni, 4) & IIf(Mes = "12", "/06/30", "/12/31") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/07/31", "/01/31")) - CDate(Fingreso) + 150
        End If
        If Fingreso > Left(AnoMes, 4) & IIf(Mes = "12", "/07/31", "/01/31") And _
            Fingreso <= Left(AnoMes, 4) & IIf(Mes = "12", "/08/31", "/02/29") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/08/30", "/02/" & str(DiasMes(2, val(Left(AnoMes, 4)))))) - CDate(Fingreso) + 120
        End If
        If Fingreso > Left(AnoMes, 4) & IIf(Mes = "12", "/08/31", "/02/" & str(DiasMes(2, val(Left(AnoMes, 4))))) And _
            Fingreso <= Left(AnoMes, 4) & IIf(Mes = "12", "/09/30", "/03/31") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/09/30", "/03/31")) - CDate(Fingreso) + 90 + IIf(Mes = "12", 1, 0)
        End If
        If Fingreso > Left(AnoMes, 4) & IIf(Mes = "12", "/09/30", "/03/31") And _
            Fingreso <= Left(AnoMes, 4) & IIf(Mes = "12", "/10/31", "/04/30") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/10/31", "/04/30")) - CDate(Fingreso) + 60 + IIf(Mes = "12", 0, 1)
        End If
        If Fingreso > Left(AnoMes, 4) & IIf(Mes = "12", "/10/31", "/04/30") And _
            Fingreso <= Left(AnoMes, 4) & IIf(Mes = "12", "/11/30", "/05/31") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/11/30", "/05/31")) - CDate(Fingreso) + 30 + IIf(Mes = "12", 1, 0)
        End If
        If Fingreso > Left(AnoMes, 4) & IIf(Mes = "12", "/11/31", "/05/31") And Fingreso <= Left(AnoMes, 4) & IIf(Mes = "12", "/12/31", "/06/30") Then
            DiasTrabajosGRATI = CDate(Left(AnoMes, 4) & IIf(Mes = "12", "/12/31", "/06/30")) - CDate(Fingreso) + IIf(Mes = "12", 0, 1)
        End If
    Else
        DiasTrabajosGRATI = 180
    End If
    If DiasTrabajosGRATI < 0 Then DiasTrabajosGRATI = 0
    If DiasTrabajosGRATI > 180 Then DiasTrabajosGRATI = 180
    Set rsdt = Nothing
Exit Function
edias:
    DiasTrabajosGRATI = 0
End Function
Public Function ValidaTareo(tareos As Long) As Long
    Dim i As Long
    ValidaTareo = 0
    With flxTareo
        For i = 1 To tareos
            If .TextMatrix(i, 1) = "" Then
            End If
        Next
        For i = 1 To tareos
            If .TextMatrix(i, 2) = "" Then
                MsgBox "El tipo de rubro en la linea " & str(i) & " esta en blanco", vbExclamation + vbOKOnly, "NOVPeru"
                ValidaTareo = i
            Else
                If ExisteTRubro(.TextMatrix(i, 2)) = False Then
                    MsgBox "El tipo de rubro en la línea " & str(i) & " no es correcto", vbExclamation + vbOKOnly, "NOVPeru"
                    ValidaTareo = i
                End If
            End If
        Next
        For i = 1 To tareos
            If .TextMatrix(i, 3) = "" Then
                MsgBox "El rubro en la linea " & str(i) & " esta en blanco", vbExclamation + vbOKOnly, "NOVPeru"
                ValidaTareo = i
            Else
                If ExisteRubro(Trim(Left(.TextMatrix(i, 3), 4))) = False Then
                    MsgBox "El rubro en la línea " & str(i) & " no es correcto", vbExclamation + vbOKOnly, "NOVPeru"
                    ValidaTareo = i
                End If
            End If
        Next
    End With
End Function
Public Function ExisteTRubro(Tipo As String) As Boolean
    SQL = "Select * from pl_tiporubros where codigo='" & Tipo & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    ExisteTRubro = False
    Do While Not (Rs.EOF)
        ExisteTRubro = True
        Rs.MoveNext
    Loop
    Set Rs = Nothing
End Function
Public Function ExisteRubro(rub As String) As Boolean
    SQL = "Select * from pl_rubrosremunerativos where codigo='" & rub & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    ExisteRubro = False
    Do While Not (Rs.EOF)
        ExisteRubro = True
        Rs.MoveNext
    Loop
    Set Rs = Nothing
End Function

Private Function GenerarPrestamo(pCodEmpleado As String, cTipoPlanilla As String) As Boolean
    On Error GoTo SERROR
    Dim pNumDoc As String
    Dim SQL As String
    Dim cAnoMes As String
    cAnoMes = CE(cboAnio.Text) & Right("00" & CE(cboMes.List(cboMes.ListIndex, 2)), 2)
    
    SQL = "select distinct documento from rh_divprestadel where solicita='" & pCodEmpleado & "' and anomes='" & cAnoMes & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    
    
    Do While Not (Rs.EOF)
        SQL = "call RH_SP_Prestamos_Detalle('PROCESA_PRESTAMO_PLAN',0,'','" & cTipoPlanilla & "','',0,'" & pCodEmpleado & "','P','" & CE(Rs.Fields("documento")) & "','" & cAnoMes & "',0,'','');"
        If ADO_EjecutaQry(SQL) = False Then
            Mensajes "Error al general el prestamo para el codigo de empleado " & pCodEmpleado & ", con documento " & pNumDoc
        End If
            
        Rs.MoveNext
    Loop
    GenerarPrestamo = True
    Exit Function
SERROR:
    GenerarPrestamo = False
End Function



Private Function ActualizaPlanilla() As Boolean
    On Error GoTo SERROR
    Dim pNumDoc As String
    Dim SQL As String
    Dim cAnoMes As String
    cAnoMes = CE(cboAnio.Text) & Right("00" & CE(cboMes.List(cboMes.ListIndex, 2)), 2)
    
        SQL = "call RH_SP_Prestamos_Detalle('ACTUALIZA_PRESTAMO_PLAN',0,'','','',0,'','','','" & cAnoMes & "',0,'','');"
        If ADO_EjecutaQry(SQL) = False Then
            Mensajes "Error al actualizar los descuentos de los prestamos del mes de los empleado"
        End If
            
    ActualizaPlanilla = True
    Exit Function
SERROR:
    ActualizaPlanilla = False
End Function


Private Sub CargaRubroAutomatico(trub As String, rub As String, rubdescrip As String, funcion As String, Optional CodEmp As String)
    Dim rsEmp As MYSQL_RS
    Dim emp As String, sbasico As Double, i As Integer
    Dim calculo As Double, f1 As String, f2 As String, afp As String
    Dim Mto As Double, SEmp As String, contrato As String
    If CodEmp <> "" Then SEmp = " and emp = '" & CodEmp & "'"
    Select Case plani
        Case "1":
            SQL = "Select emp,sbasico from pl_tareo where rub='121' and anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "'" & SEmp
            If funcion = "BC" Then
                f1 = InputBox("Ingrese la fecha inicial de bonos", "BONOS DESDE...", Format(Date, "dd/mm/yyyy"))
                f2 = InputBox("Ingrese la fecha final de bonos", "BONOS HASTA...", Format(Date, "dd/mm/yyyy"))
            End If
            If funcion = "AE" Then
                f1 = InputBox("Ingrese la fecha inicial de almuerzos", "ALMUERZOS DESDE...", Format(Date, "dd/mm/yyyy"))
                f2 = InputBox("Ingrese la fecha final de almuerzos", "ALMUERZOS HASTA...", Format(Date, "dd/mm/yyyy"))
            End If
        Case "2"
            SQL = "Select emp,sbasico from pl_tareo where rub='121' and anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "'" & SEmp
             If funcion = "BC" Then
                f1 = InputBox("Ingrese la fecha inicial de bonos", "BONOS DESDE...", Format(Date, "dd/mm/yyyy"))
                f2 = InputBox("Ingrese la fecha final de bonos", "BONOS HASTA...", Format(Date, "dd/mm/yyyy"))
            End If
        Case "3"
        Case "4"
            SQL = "Select emp,sbasico from pl_tareo where (rub='121' or rub='117') and anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "'" & SEmp
        Case "5":
            SQL = "Select emp,sbasico from pl_tareo where rub='121' and anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "'" & SEmp
        Case "6":
            SQL = "Select emp,sbasico from pl_tareo where rub='121' and anomes='" & Periodo & _
                  "' and tipo='" & plani & "' and moneda='" & moneda & "'" & SEmp
    End Select
    
    If rub = "709" Then
        Call ActualizaPlanilla
        DoEvents
    End If
    
    
    Set rsEmp = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsEmp.EOF
        emp = rsEmp.Fields("emp")
        
        sbasico = rsEmp.Fields("sbasico")
        calculo = 0
        Mto = 0
        SQL = "Select emp from pl_tareo where rub='" & rub & "' and anomes='" & Periodo & _
              "' and tipo='" & plani & "' and moneda='" & moneda & "' and emp='" & emp & "'"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If Rs.RecordCount = 0 Then
            SQL = "Select codafp,jubilado,asigfam,sctr,c.codigo,cafp from empleado e inner join contrato c on (e.codigo=c.codemp) " & _
                  "where e.codigo='" & emp & "' and estado = 'AP'"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            afp = Rs.Fields("codafp")
            contrato = Rs.Fields("codigo")
            Select Case funcion
                ' I N G R E S O S
                Case "ASIGFAM"
                    If Rs.Fields("asigfam") = "S" Then ' SI EMPLEADO TIENE ASIGNACION FAMILIAR
                        calculo = FormatNumber(AsigFam(Periodo, plani, emp, moneda), 2)
                    End If
                Case "GRATI"
                    calculo = FormatNumber(Gratificacion(Periodo, plani, emp, moneda), 2)
                Case "BC"
                    SQL = "Select bono from contrato where codemp='" & emp & "' and estado='AP'"
                    Set Rs = oConexion.EjecutaSelectRS(SQL)
                    If plani = "6" Then
                        If Rs.Fields("bono") = "S" Then ' SI EMPLEADO TIENE BONO DE CAMPO
                            calculo = FormatNumber(BonosCampo(Left(Periodo, 4) & Right("00" & Trim(str(val(Right(Periodo, 2)) + 1)), 2), plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
                        End If
                    Else
                        If Rs.Fields("bono") = "S" Then ' SI EMPLEADO TIENE BONO DE CAMPO
                            calculo = FormatNumber(BonosCampo(Periodo, plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
                        End If
                    End If
                Case "AV"
                    calculo = FormatNumber(AdelantoVacacional(Periodo, plani, emp, moneda, "I", , "V"), 2)
                Case "AVC"
                    calculo = FormatNumber(AdelantoVacacional(Periodo, plani, emp, moneda, "I", , "C"), 2)
                Case "AE"
                    calculo = FormatNumber(AlimentacionEspecie(Periodo, plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
                Case "MOVILI"
                     SQL = "Select Monto from movilidades where Codemp='" & emp & "' "
                    Set Rs = oConexion.EjecutaSelectRS(SQL)
                        If Rs.EOF = False Then ' SI EMPLEADO TIENE MOVILIDAD
                            calculo = FormatNumber(Rs.Fields("Monto"), 2)
                        End If
                Case "TS"
                    SQL = "Select bono from contrato where codemp='" & emp & "' and estado='AP'"
                    Set Rs = oConexion.EjecutaSelectRS(SQL)
                    If plani = "6" Then
                        If Rs.Fields("bono") = "S" Then ' SI EMPLEADO TIENE BONO DE CAMPO
                            calculo = FormatNumber(HorasExtra(Left(Periodo, 4) & Right("00" & Trim(str(val(Right(Periodo, 2)) + 1)), 2), plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
                        End If
                    Else
                        If Rs.Fields("bono") = "S" Then ' SI EMPLEADO TIENE BONO DE CAMPO
                            calculo = FormatNumber(HorasExtra(Periodo, plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
                        End If
                    End If
                    
                    
                    ' D E S C U E N T O S
                Case "AED"
                    calculo = FormatNumber(AlimentacionEspecieDesc(Periodo, plani, emp, moneda), 2)
                Case "AQ"
                    calculo = FormatNumber(AdelantoQuincenal(Periodo, plani, emp, moneda), 2)
                Case "AVD"
                    calculo = FormatNumber(AdelantoVacacional(Periodo, plani, emp, moneda, "D"), 2)
                Case "RTAQTA"
                    calculo = FormatNumber(RtaQta(Periodo, plani, emp, 4200, 8, 14, 17, 20, 30, moneda, sbasico), 2)
                Case "ONP"
                    If afp = "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO ESTA EN ONP
                        calculo = FormatNumber(Onp(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
                    End If
                Case "AFP10"
                    If afp <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                        calculo = FormatNumber(Afp10(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
                    End If
                Case "AFP2"
                    If afp <> "06" And Rs.Fields("sctr") = "S" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                        calculo = FormatNumber(Afp2(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
                    End If
                Case "AFPCOM"
                    If afp <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                        calculo = FormatNumber(AfpCom(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico, Rs.Fields("cafp")), 2)
                    End If
                Case "AFPCOMSEG"
                    If afp <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                        calculo = FormatNumber(AfpComSeg(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
                    End If
                Case "PRESTAMO"
                
                    Call GenerarPrestamo(emp, plani)
'                    calculo = FormatNumber(PrestamosyAdelantos(cboMes.List(cboMes.ListIndex, 2), emp, plani, "P"), 2)
'                    If plani = 1 Then
'                        calculo = calculo + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 4, "PR"), 2)
'                        calculo = calculo + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 2, "PR"), 2)
'                    End If
                    
                    calculo = 0
                    
                Case "ADEL"
                    calculo = FormatNumber(PrestamosyAdelantos(cboMes.List(cboMes.ListIndex, 2), emp, plani, "A"), 2)
                    If plani = 1 Then
                        calculo = calculo + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 4, "SS"), 2)
                        calculo = calculo + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 2, "SS"), 2)
                    End If
                Case "RETJUDI"
                    calculo = FormatNumber(RetencionJudicial(Periodo, plani, emp, moneda), 2)
                    'A P O R T E S
                Case "ESSALUD"
                    calculo = FormatNumber(Essalud(Periodo, plani, emp, moneda, sbasico), 2)
                Case "SENATI"
                    If Rs.Fields("sctr") = "S" Then ' SI EMPLEADO ESTA EN SENATI
                        'calculo = FormatNumber(Sctr(Periodo, plani, emp, moneda, sbasico), 2)
                        calculo = FormatNumber(SENATI(Periodo, plani, emp, moneda, sbasico), 2)
                    End If
                Case "AFPSCTR"
                    If Rs.Fields("codafp") <> "06" And Rs.Fields("sctr") = "S" And Rs.Fields("jubilado") = "N" Then  ' SI EMPLEADO NO ESTA EN ONP
                        calculo = FormatNumber(Afp2E(Periodo, plani, emp, moneda, sbasico), 2)
                    End If
            End Select
            If (calculo <> 0) Then
                SQL = "Insert into pl_tareo (anomes,tipo,moneda,emp,rub,cant,fecha,afp,sbasico,codcontrato) values " & _
                      " ('" & Periodo & "','" & plani & "','" & moneda & "','" & emp & "','" & rub & "'," & calculo & _
                      " ,'" & Format(DtpFecha.Text, "yyyy/mm/dd") & "','" & afp & "'," & sbasico & ",'" & contrato & "')"
                oConexionMYSQL.Execute SQL
            End If
            If funcion = "PRESTAMO" Or funcion = "ADEL" Then
                If calculo > 0 Then
                    ActualizaMontos Periodo, plani, moneda, IIf(rub = "709", "P", "A"), emp, 0
                End If
            End If
        End If
        rsEmp.MoveNext
    Loop
    If funcion = "AE" Then
        FlgRubAl = True
        For i = 1 To CboDescuentos.ListCount - 1
            CboDescuentos.ListIndex = i
            If CboDescuentos.List(CboDescuentos.ListIndex, 2) = "708" Then
                btnCargarD_Click
                BotonRubro CboDescuentos.List(CboDescuentos.ListIndex, 2), "D"
                FlgRubAl = False
                Exit For
            End If
        Next
    End If
    Set rsEmp = Nothing
    Set Rs = Nothing
End Sub
Private Function CargaRubro(funcion As String, emp As String, sbasico As Double) As Double
    Dim f1 As String, f2 As String
    Dim Mto As Double
    SQL = "Select asigfam,jubilado,sctr,codafp,cafp from empleado where codigo='" & emp & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    Select Case funcion
        ' I N G R E S O S
        Case "ASIGFAM"
            If Rs.Fields("asigfam") = "S" Then ' SI EMPLEADO TIENE ASIGNACION FAMILIAR
                CargaRubro = Round(AsigFam(Periodo, plani, emp, moneda), 2)
            End If
        Case "BC"
            f1 = InputBox("Ingrese la fecha inicial de bonos", "BONOS DESDE...", Format(Date, "dd/mm/yyyy"))
            f2 = InputBox("Ingrese la fecha final de bonos", "BONOS HASTA...", Format(Date, "dd/mm/yyyy"))
            SQL = "Select bono from contrato where codemp='" & emp & "' and estado='AP'"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            If Rs.Fields("bono") = "S" Then ' SI EMPLEADO TIENE BONO DE CAMPO
                CargaRubro = Round(BonosCampo(Periodo, plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
            End If
        Case "AV"
            CargaRubro = Round(AdelantoVacacional(Periodo, plani, emp, moneda, "I", , "V"), 2)
        Case "AVC"
            CargaRubro = Round(AdelantoVacacional(Periodo, plani, emp, moneda, "I", , "C"), 2)
        Case "AE"
            f1 = InputBox("Ingrese la fecha inicial de almuerzos", "ALMUERZOS DESDE...", Format(Date, "dd/mm/yyyy"))
            f2 = InputBox("Ingrese la fecha final de almuerzos", "ALMUERZOS HASTA...", Format(Date, "dd/mm/yyyy"))
            CargaRubro = Round(AlimentacionEspecie(Periodo, plani, emp, moneda, Format(f1, "yyyy/mm/dd"), Format(f2, "yyyy/mm/dd")), 2)
            ' D E S C U E N T O S
        Case "AED"
            CargaRubro = Round(AlimentacionEspecie(Periodo, plani, emp, moneda), 2)
        Case "AQ"
            CargaRubro = Round(AdelantoQuincenal(Periodo, plani, emp, moneda), 2)
        Case "AVD"
            CargaRubro = Round(AdelantoVacacional(Periodo, plani, emp, moneda, "D"), 2)
        Case "RTAQTA"
            CargaRubro = Round(RtaQta(Periodo, plani, emp, 4200, 8, 14, 17, 20, 30, moneda, sbasico), 2)
        Case "ONP"
            If Rs.Fields("codafp") = "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO ESTA EN ONP
                CargaRubro = Round(Onp(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
            End If
        Case "AFP10"
            If Rs.Fields("codafp") <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                CargaRubro = Round(Afp10(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
            End If
        Case "AFP2"
            If Rs.Fields("codafp") <> "06" And Rs.Fields("sctr") = "S" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                CargaRubro = Round(Afp2(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
            End If
        Case "AFPCOM"
            If Rs.Fields("codafp") <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                CargaRubro = Round(AfpCom(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico, Rs.Fields("cafp")), 2)
            End If
        Case "AFPCOMSEG"
            If Rs.Fields("codafp") <> "06" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                CargaRubro = Round(AfpComSeg(Periodo, plani, emp, moneda, Rs.Fields("CODAFP"), sbasico), 2)
            End If
        Case "PRESTAMO"
            CargaRubro = FormatNumber(PrestamosyAdelantos(cboMes.List(cboMes.ListIndex, 2), emp, plani, "P"), 2)
            If plani = 1 Then
                CargaRubro = CargaRubro + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 4, "PR"), 2)
                CargaRubro = CargaRubro + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 2, "PR"), 2)
            End If
        Case "ADEL"
            CargaRubro = FormatNumber(PrestamosyAdelantos(cboMes.List(cboMes.ListIndex, 2), emp, plani, "A"), 2)
            If plani = 1 Then
                CargaRubro = CargaRubro + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 4, "SS"), 2)
                CargaRubro = CargaRubro + FormatNumber(TareoQuincenal(cboMes.List(cboMes.ListIndex, 2), emp, 2, "SS"), 2)
            End If
        Case "RETJUDI"
            CargaRubro = FormatNumber(RetencionJudicial(Periodo, plani, emp, moneda), 2)
            'A P O R T E S
        Case "ESSALUD"
            CargaRubro = FormatNumber(Essalud(Periodo, plani, emp, moneda, sbasico), 2)
        Case "SENATI"
            If Rs.Fields("sctr") = "S" Then ' SI EMPLEADO TIENE SCTR
                'CargaRubro = Round(Sctr(Periodo, plani, emp, moneda, sbasico), 2)
                CargaRubro = Round(SENATI(Periodo, plani, emp, moneda, sbasico), 2)
            End If
        Case "AFPSCTR"
            If Rs.Fields("codafp") <> "06" And Rs.Fields("sctr") = "S" And Rs.Fields("jubilado") = "N" Then ' SI EMPLEADO NO ESTA EN ONP
                CargaRubro = Round(Afp2E(Periodo, plani, emp, moneda, sbasico), 2)
            End If
    End Select
End Function
Private Sub BotonRubro(rub As String, Tip As String)
    Select Case Tip
        Case "I": btnDesCargarI.Enabled = False
        Case "D": btnDesCargarD.Enabled = False
        Case "A": btnDesCargarA.Enabled = False
    End Select
    SQL = "Select distinct rub from pl_tareo where anomes='" & Periodo & _
          "' and tipo='" & plani & "' and moneda='" & moneda & "' and rub='" & rub & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Not (Rs.EOF And Rs.BOF) Then
        Select Case Tip
            Case "I": btnDesCargarI.Enabled = True
            Case "D": btnDesCargarD.Enabled = True
            Case "A": btnDesCargarA.Enabled = True
        End Select
    End If
    Set Rs = Nothing
End Sub
Private Function EstadoPlani(AnoMes As String, moneda As String, plani As String) As Boolean
    EstadoPlani = False
    SQL = "Select * from pl_planiproc where anomes='" & AnoMes & _
          "' and proceso='" & plani & "' and mon='" & moneda & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Not (Rs.EOF And Rs.BOF) Then
        EstadoPlani = True
    End If
    Set Rs = Nothing
End Function
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flxTareo
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 100 Then
            Lstep = 1
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 0
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        If NewValue = 0 Then
            .TopRow = 1
        Else
            .TopRow = NewValue
        End If
    End With
End Sub
Sub ActualizaPagos(Periodo As String, plani As String, moneda As String, emp As String, Mto As Double, Tipo As String)
    Dim Tot As Integer
    If plani = 1 Then
        SQL = "select * from rh_pagosemp where solicita = '" & emp & "' and tipoplani ='2' and anomesplani = '" & Periodo & "' and monplani = '" & moneda & "' and liquid = 'N' order by fecha"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If Not Rs.EOF() Then
            Mto = Mto - CDbl(Rs.Fields("montoabonado"))
        End If
        SQL = "select * from rh_pagosemp where solicita = '" & emp & "' and tipoplani ='4' and anomesplani = '" & Periodo & "' and monplani = '" & moneda & "' and liquid = 'N' order by fecha"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If Not Rs.EOF() Then
            Mto = Mto - CDbl(Rs.Fields("montoabonado"))
        End If
        Set Rs = Nothing
    End If
    SQL = "select * from rh_pagosemp where solicita = '" & emp & "' and tipoplani = '" & plani & "' " & _
          "and anomesplani = '" & Periodo & "' and monplani = '" & moneda & "' and tipo = '" & Tipo & "' and liquid = 'N' order by fecha"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    Tot = 1
    Do While Not Rs.EOF
        If Mto >= CDbl(Rs.Fields("montoabonado")) Then
            If Rs.RecordCount = Tot Then
                SQL = "update RH_PAGOSEMP set montoabonado = " & Mto & " WHERE ANOMESPLANI = '" & Periodo & "' and tipoplani='" & plani & "' " & _
                      "AND MONPLANI = '" & moneda & "' and solicita = '" & Mid(lblCodEmp, 9, 11) & "' " & _
                      "and documento ='" & Trim(Rs.Fields("documento")) & "' and tipo = '" & Trim(Rs.Fields("tipo")) & "' and liquid = 'N'"
                oConexionMYSQL.Execute SQL
                ActualizaMtoPagado CDbl(Mto), Trim(Rs.Fields("anomes")), Trim(Rs.Fields("documento")), Trim(Rs.Fields("fecha")), Trim(Rs.Fields("solicita")), 2, CDbl(Rs.Fields("montoabonado"))
            Else
                Mto = Mto - CDbl(Rs.Fields("montoabonado"))
            End If
        Else
            If Mto = 0 Then
                SQL = "DELETE FROM RH_PAGOSEMP WHERE ANOMESPLANI = '" & Periodo & "' and tipoplani='" & plani & "' " & _
                      "AND MONPLANI = '" & moneda & "' and solicita = '" & Mid(lblCodEmp, 9, 11) & "' " & _
                      "and documento ='" & Trim(Rs.Fields("documento")) & "' and tipo = '" & Trim(Rs.Fields("tipo")) & "' and liquid = 'N'"
                oConexionMYSQL.Execute SQL
                ActualizaMtoPagado CDbl(Mto), Trim(Rs.Fields("anomes")), Trim(Rs.Fields("documento")), Trim(Rs.Fields("fecha")), Trim(Rs.Fields("solicita")), 2, CDbl(Rs.Fields("montoabonado"))
            Else
                SQL = "update RH_PAGOSEMP set montoabonado = " & Mto & " WHERE ANOMESPLANI = '" & Periodo & "' and tipoplani='" & plani & "' " & _
                      "AND MONPLANI = '" & moneda & "' and solicita = '" & Mid(lblCodEmp, 9, 11) & "' " & _
                      "and documento ='" & Trim(Rs.Fields("documento")) & "' and tipo = '" & Trim(Rs.Fields("tipo")) & "' and liquid = 'N'"
                oConexionMYSQL.Execute SQL
                ActualizaMtoPagado CDbl(Mto), Trim(Rs.Fields("anomes")), Trim(Rs.Fields("documento")), Trim(Rs.Fields("fecha")), Trim(Rs.Fields("solicita")), 2, CDbl(Rs.Fields("montoabonado"))
                If CDbl(Rs.Fields("montoabonado")) > Mto Then Mto = 0
            End If
        End If
        Rs.MoveNext
        Tot = Tot + 1
    Loop
    Set Rs = Nothing
End Sub
Sub EliminaMtoPagado(AnoMes As String, plani As String, mon As String, Optional Tipo As String, Optional CodEmp As String)
    Dim SQ As String
    SQ = ""
    If CodEmp <> "" Then
        SQ = " and solicita = '" & CodEmp & "'"
    End If
    SQL = "select * from rh_divprestadel where anomes='" & AnoMes & "' and tipo = '" & Tipo & "' and tipoplani='" & plani & "' " & SQ & _
          " union select * from rh_divprestadel where anomes='" & IIf(Right(Mes, 2) = "01", Left(AnoMes, 4) - 1, Left(AnoMes, 4)) & IIf(Right(AnoMes, 2) = "01", "12", Right("00" & val(Right(AnoMes, 2)) - 1, 2)) & "' " & _
          "and tipo = '" & Tipo & "' and descontado < monto AND tipoplani='" & plani & "' " & SQ & _
          " order by solicita,documento"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    Do While Not Rs.EOF()
        ActualizaMtoPagado CDbl(Rs.Fields("monto")), Tipo, Trim(Rs.Fields("documento")), "", Trim(Rs.Fields("solicita")), 0
        Rs.MoveNext
    Loop
End Sub
Sub ActualizaMtoPagado(Mto As Double, AnoMes As String, Doc As String, fec As String, soli As String, Tip As Integer, Optional Monto As Double)
    If Tip = 0 Then
        SQL = "UPDATE documentos_rrhh set montopagado = montopagado + " & CDbl(Mto) & " " & _
              "where documento = '" & Trim(Doc) & "' and coddoc = '" & IIf(AnoMes = "P", "PR", "SS") & "' " & _
              "and solicita = '" & Trim(soli) & "'"
    ElseIf Tip = 1 Then
        SQL = "UPDATE documentos_rrhh set montopagado = montopagado - " & CDbl(Mto) & " " & _
              "where documento = '" & Trim(Doc) & "' and coddoc = '" & IIf(AnoMes = "P", "PR", "SS") & "' " & _
              "and solicita = '" & Trim(soli) & "'"
    Else
        SQL = "UPDATE documentos_rrhh set montopagado = montopagado - " & CDbl(Monto) & " + " & CDbl(Mto) & " " & _
              "where documento = '" & Trim(Doc) & "' and coddoc = '" & IIf(AnoMes = "P", "PR", "SS") & "' " & _
              "and solicita = '" & Trim(soli) & "'"
    End If
    oConexionMYSQL.Execute SQL
End Sub
Public Sub LlenarAnio()
    Dim i As Integer, anioctual As Integer
    cboAnio.Clear
    cboAnio.AddItem "Seleccionar..."
    cboAnio.List(0, 2) = "00"
    For i = 1 To 3
        cboAnio.AddItem Trim(str(CDbl(strAnoSistema) - (i - 1)))
        cboAnio.List(i, 1) = Trim(str(CDbl(strAnoSistema) - (i - 1)))
        If Trim(Year(Date)) = Trim(cboAnio.List(i, 1)) Then
            anioctual = i
        End If
    Next
    cboAnio.ListIndex = anioctual
End Sub
