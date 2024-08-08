VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmRegBonos 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Bonos - Salidas a Campo"
   ClientHeight    =   7740
   ClientLeft      =   1110
   ClientTop       =   5025
   ClientWidth     =   15120
   Icon            =   "frmRegBonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   15120
   Begin Proyecto1.chameleonButton cmdRefrescar 
      Height          =   330
      Left            =   12345
      TabIndex        =   39
      ToolTipText     =   "Refrescar"
      Top             =   45
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   582
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
      MICON           =   "frmRegBonos.frx":0442
      PICN            =   "frmRegBonos.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frlote 
      BackColor       =   &H009F5539&
      Caption         =   "Lotes"
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
      Left            =   12390
      TabIndex        =   21
      Top             =   360
      Width           =   2415
      Begin MSForms.ComboBox cbolote 
         Height          =   300
         Left            =   45
         TabIndex        =   22
         Top             =   165
         Width           =   2310
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "4075;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Caption         =   "Reporte"
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
      Height          =   645
      Left            =   3495
      TabIndex        =   14
      Top             =   7020
      Width           =   5445
      Begin VB.TextBox txtValorRefri 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4020
         MaxLength       =   12
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dptFecIniR 
         Height          =   315
         Left            =   765
         TabIndex        =   16
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105644033
         CurrentDate     =   38597
      End
      Begin MSComCtl2.DTPicker dptFecFinR 
         Height          =   315
         Left            =   2100
         TabIndex        =   17
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105644033
         CurrentDate     =   38597
      End
      Begin Proyecto1.chameleonButton btnReporte 
         Height          =   345
         Left            =   4920
         TabIndex        =   18
         Top             =   210
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
         MICON           =   "frmRegBonos.frx":05B8
         PICN            =   "frmRegBonos.frx":05D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechas"
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
         Height          =   270
         Left            =   60
         TabIndex        =   20
         Top             =   255
         Width           =   690
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Height          =   285
         Index           =   0
         Left            =   3465
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   15
      TabIndex        =   1
      Top             =   -75
      Width           =   12300
      Begin VB.TextBox TxtImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5415
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   150
         Width           =   585
      End
      Begin VB.TextBox TxtNom 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6870
         TabIndex        =   2
         Top             =   150
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker DtFecIni 
         Height          =   315
         Left            =   9585
         TabIndex        =   36
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105644033
         CurrentDate     =   38597
      End
      Begin MSComCtl2.DTPicker DtFecFin 
         Height          =   315
         Left            =   10920
         TabIndex        =   37
         Top             =   150
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   105644033
         CurrentDate     =   38597
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechas"
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
         Height          =   285
         Index           =   5
         Left            =   8865
         TabIndex        =   38
         Top             =   180
         Width           =   705
      End
      Begin MSForms.ComboBox CboDiv 
         Height          =   315
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   2085
         VariousPropertyBits=   746604569
         DisplayStyle    =   7
         Size            =   "3678;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   2070
         TabIndex        =   8
         Top             =   180
         Width           =   735
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   420
      End
      Begin MSForms.ComboBox CboMes 
         Height          =   315
         Left            =   495
         TabIndex        =   6
         Top             =   150
         Width           =   1560
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2752;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Height          =   270
         Left            =   4905
         TabIndex        =   5
         Top             =   180
         Width           =   510
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apellidos"
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
         Height          =   285
         Index           =   4
         Left            =   6030
         TabIndex        =   4
         Top             =   180
         Width           =   825
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshEmpO 
      Height          =   5790
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   10213
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   16777215
      Rows            =   4
      FixedRows       =   3
      FixedCols       =   0
      ForeColorSel    =   16777215
      GridColor       =   12498349
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Msf 
      Height          =   6135
      Left            =   14550
      TabIndex        =   10
      Top             =   870
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   14737632
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   14640
      TabIndex        =   11
      Top             =   7200
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
      MICON           =   "frmRegBonos.frx":0B16
      PICN            =   "frmRegBonos.frx":0B32
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
      Left            =   12000
      TabIndex        =   12
      ToolTipText     =   "Guardar"
      Top             =   8850
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Grabar"
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
      MICON           =   "frmRegBonos.frx":0EF8
      PICN            =   "frmRegBonos.frx":0F14
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TabStrip TabS 
      Height          =   6120
      Left            =   30
      TabIndex        =   13
      Top             =   885
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   10795
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   13
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  A - B  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  C - D  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  E - F  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  G - H  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  I - J  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  K - L  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  M - Ñ  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  O - P  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Q - R  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  S - T  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  U - V  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  W - X  "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Y - Z  "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblcodemp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   855
      TabIndex        =   34
      Top             =   600
      Width           =   2640
   End
   Begin VB.Label lbldiv 
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
      Height          =   210
      Left            =   9870
      TabIndex        =   33
      Top             =   585
      Width           =   2610
   End
   Begin VB.Label Lbldato 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   240
      Left            =   13470
      TabIndex        =   32
      Top             =   9000
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSForms.ListBox LstSede 
      Height          =   6150
      Left            =   12390
      TabIndex        =   31
      Top             =   855
      Width           =   2145
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3784;10065"
      MatchEntry      =   0
      ListStyle       =   1
      FontName        =   "Arial"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   90
      TabIndex        =   30
      Top             =   7425
      Visible         =   0   'False
      Width           =   75
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
      Index           =   1
      Left            =   60
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Label lblTotGen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   12705
      TabIndex        =   28
      Top             =   7215
      Width           =   1905
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total General"
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
      Height          =   240
      Left            =   11220
      TabIndex        =   27
      Top             =   7245
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "División :"
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
      Height          =   195
      Left            =   8955
      TabIndex        =   26
      Top             =   600
      Width           =   810
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Código :"
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
      Height          =   195
      Left            =   75
      TabIndex        =   25
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   9000
      TabIndex        =   24
      Top             =   7275
      Width           =   465
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   9465
      TabIndex        =   23
      Top             =   7215
      Width           =   1665
   End
   Begin MSForms.Label lblo 
      Height          =   360
      Left            =   30
      TabIndex        =   35
      Top             =   510
      Width           =   12390
      ForeColor       =   128
      Caption         =   "Listado  de  Empleados"
      PicturePosition =   393216
      Size            =   "21855;635"
      BorderColor     =   -2147483639
      SpecialEffect   =   3
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRegBonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColIniO As Integer, FilaIniO As Integer
Dim emp As String
Dim Datos As String, CantTrab As Integer
Public GridSel As Integer
Dim cadenaemp As String, dias As Integer
Private Trab(1 To 500) As Integer
Private TempTrab(1 To 500) As Integer
Dim EntroBusqueda As Boolean, H As Integer
Dim CReaTabla As Boolean
Dim ColIniMesAnt As Integer, ColIniMesAct As Integer, ColIniMesPos As Integer
Dim MesIni As String, MesFin As String
Dim AnoIni As String, AnoFin As String, AnoMed As String

Sub GrabaBonos()
On Error GoTo CtrlError
    Dim I As Integer, J As Integer, k As Integer
    Dim SQ As String, sede As String
    Dim Tot As Integer, TotC As Integer
    Dim Anio As String, Mes As String
    Dim Cant As Integer
    Screen.MousePointer = vbHourglass
    With MshEmpO
        For k = 1 To CantTrab
            I = Trab(k) ' - 1
            Tot = 0
            SQ = "delete from rh_bonosempleados where DATE_FORMAT(fecha,'%Y%m') between '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "' " & _
                 "and '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "' and codemp = '" & .TextMatrix(I, .Cols - 2) & "'"
            oConexionMYSQL.Execute SQ
            For J = 4 To .Cols - 5
                If Trim(.TextMatrix(I, J)) = "T" Then
                    .row = I: .Col = J
                    If ColIniMesAnt <= J Then
                        Anio = AnoIni
                        Mes = MesIni
                    End If
                    If ColIniMesAct <= J Then
                        Anio = strAnoSistema
                        Mes = cboMes.List(cboMes.ListIndex, 2)
                    End If
                    If ColIniMesPos <= J Then
                        Anio = AnoFin
                        Mes = MesFin
                    End If
                    If .CellForeColor <> &HFFFFFF Then
                        sede = DevCodPozo(I, J)
                        SQ = "call Insert_Bonos('" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                             "'" & Format(CDate(Anio & "/" & Mes & "/" & Right("00" & .TextMatrix(2, J), 2)), "yyyy/mm/dd") & "', " & _
                             "'" & sede & "','" & CboLote.List(CboLote.ListIndex, 2) & "','N'," & Trim(CDbl(IIf(val(.TextMatrix(I, 3)) = 0, 0, .TextMatrix(I, 3)))) & ")"
                        oConexionMYSQL.Execute SQ
                    End If
                    If Format(CDate(Anio & "/" & Mes & "/" & Right("00" & .TextMatrix(2, J), 2)), "dd/mm/yyyy") >= CDate(dtFecIni.Value) Then
                        If Format(CDate(Anio & "/" & Mes & "/" & Right("00" & .TextMatrix(2, J), 2)), "dd/mm/yyyy") <= CDate(dtFecFin) Then
                            If Trim(.TextMatrix(I, J)) = "T" Then Tot = Tot + 1
                        End If
                    End If
                End If
            Next
        Next
        .TextMatrix(I, 2) = Tot
        .TextMatrix(I, .Cols - 3) = TotalBono(Trim(.TextMatrix(I, .Cols - 2))) 'CDbl(IIf(val(.TextMatrix(i, 3)) = 0, 0, .TextMatrix(i, 3))) * Tot
    End With
    Screen.MousePointer = vbDefault
Exit Sub
CtrlError:
    MsgBox "No se pudo grabar la información para el empleado:" & emp & vbNewLine & _
           "revise los refrigerios o consulte con el administrador del sistema", vbOKOnly + vbExclamation, "NOVPeru"
End Sub

Function TotalBono(emp As String) As Double
Dim SumaTot As Double
Dim SQL As String
Dim RQ As MYSQL_RS
    SQL = "select sum(bono) as mto from rh_bonosempleados where fecha between '" & Format(dtFecIni.Value, "yyyy/mm/dd") & "' and '" & Format(dtFecFin, "yyyy/mm/dd") & "' and codemp = '" & emp & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        TotalBono = FormatNumber(RQ.Fields("mto"), 2)
    End If
Set RQ = Nothing
End Function

Function DevCodPozo(fila As Integer, columna As Integer) As String
Dim k As Integer
    For k = 0 To LstSede.ListCount - 1
        With MshEmpO
            .Col = columna: .row = fila
            Msf.Col = 0: Msf.row = k
            If .CellForeColor = Msf.CellBackColor Then
                DevCodPozo = LstSede.List(k, 1)
                Exit For
            End If
        End With
    Next
End Function


Private Sub btnReporte_Click()
    Me.MousePointer = vbHourglass
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.fecha = cboMes.Text
    oReporte.Nombre1 = "S/."
    oReporte.Titulo = "REPORTE  RESUMEN  DE  BONOS  AL " & dptFecFinR.Value
    oReporte.Reporte = "Rep_BonosResumenN.rpt"
    oReporte.sp_Rep_Bonos CDbl(IIf(val(txtValorRefri) = 0, 0, txtValorRefri)), Format(dptFecIniR.Value, "yyyy/mm/dd"), Format(dptFecFinR.Value, "yyyy/mm/dd"), Right("00" & Month(Format(dptFecIniR.Value, "dd/mm/yyyy")), 2), cboDiv.List(cboDiv.ListIndex, 1)
    Me.MousePointer = vbNormal
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cbodiv_Change()
    If cboDiv.ListIndex > -1 Then
        Select Case TabS.SelectedItem.Index
            Case 1: Filtrado "A", "B", cboDiv.List(cboDiv.ListIndex, 1)
            Case 2: Filtrado "C", "D", cboDiv.List(cboDiv.ListIndex, 1)
            Case 3: Filtrado "E", "F", cboDiv.List(cboDiv.ListIndex, 1)
            Case 4: Filtrado "G", "H", cboDiv.List(cboDiv.ListIndex, 1)
            Case 5: Filtrado "I", "J", cboDiv.List(cboDiv.ListIndex, 1)
            Case 6: Filtrado "K", "L", cboDiv.List(cboDiv.ListIndex, 1)
            Case 7: Filtrado "M", "Ñ", cboDiv.List(cboDiv.ListIndex, 1)
            Case 8: Filtrado "O", "P", cboDiv.List(cboDiv.ListIndex, 1)
            Case 9: Filtrado "Q", "R", cboDiv.List(cboDiv.ListIndex, 1)
            Case 10: Filtrado "S", "T", cboDiv.List(cboDiv.ListIndex, 1)
            Case 11: Filtrado "U", "V", cboDiv.List(cboDiv.ListIndex, 1)
            Case 12: Filtrado "W", "X", cboDiv.List(cboDiv.ListIndex, 1)
            Case 13: Filtrado "Y", "Z", cboDiv.List(cboDiv.ListIndex, 1)
        End Select
    End If
End Sub

Private Sub Filtrado(LetraIni As String, LetraFin As String, Division As String)
    Dim I As Integer
    Dim Col As Integer
    With MshEmpO
        For I = 3 To .Rows - 1
            Col = 1
            .RowHeight(I) = 245
            If Division = "0006" Or Division = "00" Then
                If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) Then
                Else
                    .RowHeight(I) = 0
                End If
            Else
                If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) And (.TextMatrix(I, .Cols - 1) = Division) Then
                Else
                    .RowHeight(I) = 0
                End If
            End If
        Next
    End With
End Sub

Private Sub CboLote_Change()
    Pozos
End Sub

Private Sub cboMes_Change()
    If cboMes.ListIndex > 0 Then
        If cboMes.List(cboMes.ListIndex, 2) = "12" Then
            MesIni = "11"
            MesFin = "01"
            AnoIni = strAnoSistema
            AnoFin = CDbl(strAnoSistema) + 1
        Else
            MesIni = IIf(cboMes.List(cboMes.ListIndex, 2) = "01", "12", cboMes.List(cboMes.ListIndex - 1, 2))
            MesFin = IIf(cboMes.List(cboMes.ListIndex, 2) = "12", "01", cboMes.List(cboMes.ListIndex + 1, 2))
            AnoIni = IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema)
            AnoFin = IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema)
        End If
        cboDiv.Enabled = True
        H = 1
        TabS_Click
        MshEmpO.LeftCol = ColIniMesAct
    Else
        If cboDiv.ListCount > 0 Then cboDiv.ListIndex = 0
        cboDiv.Enabled = False
    End If
End Sub

Sub CalculoTotal()
Dim SumaTot As Double
    With MshEmpO
        For I = 3 To .Rows - 1
            SumaTot = SumaTot + CDbl(IIf(val(.TextMatrix(I, .Cols - 3)) = 0, 0, .TextMatrix(I, .Cols - 3)))
        Next
        lblTotal = "S/.  " & FormatNumber(SumaTot, 2)
    End With
End Sub

Sub TotalGeneral()
Dim SumaTot As Double
Dim SQL As String
Dim RQ As MYSQL_RS
    SQL = "select sum(bono) as mto from rh_bonosempleados where fecha between '" & Format(dtFecIni.Value, "yyyy/mm/dd") & "' and '" & Format(dtFecFin, "yyyy/mm/dd") & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        lblTotGen = "S/.  " & FormatNumber(RQ.Fields("mto"), 2)
    End If
Set RQ = Nothing
End Sub

Private Sub cmdRefrescar_Click()
    cboMes_Change
End Sub

Private Sub Form_Load()
    Call WheelHook(frmRegBonos)
    CReaTabla = True
    Me.Left = 0
    Me.Top = 0
    SCol = -1
    dptFecIniR.Value = Date
    dptFecFinR.Value = Date
    dtFecIni.Value = DateSerial(Year(Date), Month(Date) + 0, 1)
    dtFecFin.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
    cadenaemp = ""
    lblCadBus = ""
    lblTotGen = "S/. 0.00"
    lblTotal = "S/. 0.00"
    ConfiguraGrilla
    Divisiones cboDiv
    ArrColores
    CargaLotes CboLote
    CboLote.ListIndex = 1
    Pozos
    LlenarMesP cboMes
End Sub

Sub ConfiguraGrilla()
Dim fecha As String
Dim Fin As Integer
    With MshEmpO
        fecha = "01/" & cboMes.List(cboMes.ListIndex, 2) & "/" & strAnoSistema
        If IsDate(fecha) Then
            .Rows = 4
            .Cols = 4
            .FixedRows = 3
            .Clear
            .TextMatrix(2, 0) = "Item"
            .ColWidth(0) = 0 '330

            .ColWidth(1) = 2900
            .TextMatrix(2, 1) = Space(25) & "Empleado"
            
            .ColWidth(2) = 360
            .TextMatrix(2, 2) = "Días"
            
            .ColWidth(3) = 750
            .TextMatrix(2, 3) = "Bono"
                        
            dias = Day(DateSerial(IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema), val(MesIni) + 1, 0))
            .Cols = .Cols + dias
            ColIniMesAnt = 4
            For I = 1 To dias
                .ColWidth(I + 3) = 250
                .TextMatrix(0, I + 3) = NombreMes(MesIni, False)
                .TextMatrix(1, I + 3) = Left(Format(CDate(Right("00" & I, 2) & "/" & MesIni & "/" & strAnoSistema), "dddd"), 1)
                .TextMatrix(2, I + 3) = CStr(I)
                If .TextMatrix(1, I + 3) = "D" Then
                    .row = 1: .Col = I + 3: .CellForeColor = vbRed
                    .row = 2: .Col = I + 3: .CellForeColor = vbRed
                Else
                    .row = 1: .Col = I + 3: .CellForeColor = vbBlack
                    .row = 2: .Col = I + 3: .CellForeColor = vbBlack
                End If
                .ColAlignment(I + 3) = flexAlignCenterCenter
            Next
            Fin = I + 3
            
            dias = Day(DateSerial(Year(CDate(fecha)), Month(CDate(fecha)) + 1, 0))
            .Cols = .Cols + dias
            ColIniMesAct = Fin
            For I = 1 To dias
                .ColWidth(Fin) = 250
                .TextMatrix(0, Fin) = NombreMes(cboMes.List(cboMes.ListIndex, 2), False)
                .TextMatrix(1, Fin) = Left(Format(CDate(Right("00" & I, 2) & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & strAnoSistema), "dddd"), 1)
                .TextMatrix(2, Fin) = CStr(I)
                If .TextMatrix(1, Fin) = "D" Then
                    .row = 1: .Col = Fin: .CellForeColor = vbRed
                    .row = 2: .Col = Fin: .CellForeColor = vbRed
                Else
                    .row = 1: .Col = Fin: .CellForeColor = vbBlack
                    .row = 2: .Col = Fin: .CellForeColor = vbBlack
                End If
                .ColAlignment(Fin) = flexAlignCenterCenter
                Fin = Fin + 1
            Next
            
            dias = Day(DateSerial(IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema), val(MesFin) + 1, 0))
            .Cols = .Cols + dias
            ColIniMesPos = Fin
            For I = 1 To dias
                .ColWidth(Fin) = 250
                .TextMatrix(0, Fin) = NombreMes(MesFin, False)
                .TextMatrix(1, Fin) = Left(Format(CDate(Right("00" & I, 2) & "/" & MesFin & "/" & strAnoSistema), "dddd"), 1)
                .TextMatrix(2, Fin) = CStr(I)
                If .TextMatrix(1, Fin) = "D" Then
                    .row = 1: .Col = Fin: .CellForeColor = vbRed
                    .row = 2: .Col = Fin: .CellForeColor = vbRed
                Else
                    .row = 1: .Col = Fin: .CellForeColor = vbBlack
                    .row = 2: .Col = Fin: .CellForeColor = vbBlack
                End If
                .ColAlignment(Fin) = flexAlignCenterCenter
                Fin = Fin + 1
            Next
                        
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(2, .Cols - 1) = "fechaingreso"
            
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 1000
            .TextMatrix(2, .Cols - 1) = "SubTotal"

            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(2, .Cols - 1) = "codigo"

            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(2, .Cols - 1) = "ccHFM"

            .FixedCols = 4
            .GridColorFixed = &H35312F
            .CellForeColor = &H80FFFF
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictColumns
        End If
    End With
End Sub

Sub ConfiguraGrid()
    With Msf
        .Cols = 2
        .Rows = 0
        .ColWidth(0) = 265
        .ColWidth(1) = 0
    End With
End Sub

Sub Pozos()
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer

    SQ = "cencos_pozo_lote where lote = '" & CboLote.List(CboLote.ListIndex, 2) & "' group by pozo"
    Set RQ = oConexion.EjecutaSelect(SQ)
    LstSede.Clear
    
    With Msf
        I = 0
        ConfiguraGrid
        If Not RQ.EOF() Then
            Do While Not RQ.EOF()
                LstSede.AddItem Trim(RQ.Fields("dpozo"))
                LstSede.List(I, 1) = Trim(RQ.Fields("dpozo"))
                .Rows = .Rows + 1
                .TextMatrix(I, 1) = I
                .Col = 0: .row = I
                .RowHeight(I) = 277
                .CellBackColor = ArrColor(I + 1)
                I = I + 1
                RQ.MoveNext
            Loop
            LstSede.ListIndex = 0
        End If
    End With
    Set RQ = Nothing
End Sub

Sub CargaEmpleados(divis As String, Nombre As String, LI As String, LF As String)
Dim SQ As String
Dim RQ As MYSQL_RS, RQASIS As MYSQL_RS
Dim I As Integer, FechaDia As String, k As Integer
Dim DiaIniV As Integer, DiaFinV As Integer, T As Integer, MesIniV As Integer
Dim MesFinV As Integer, TotC As Integer, TotO As Integer, FlgInsert As Boolean
Dim NomEmp As String, FecIngreso As String, FecTermino As String
Dim aux As String, codAux As String, AnioIniV As String, AnioFinV As String
Dim SqLDiasAnt As String, SqLDiasDes As String, ColIni As Integer, SqlDiasAct As String
Dim AnioMes As String

    lblCadBus = ""
    cadenaemp = ""
    MshEmpO.Redraw = False
    Msf.Redraw = False
    ConfiguraGrilla
    Screen.MousePointer = vbHourglass
    NomEmp = ""
    EntroBusqueda = False
    If Trim(TxtNom) <> "" Then
        EntroBusqueda = True
        NomEmp = " and e.apepat = '" & Trim(TxtNom) & "'"
    End If
    CrearTablaTemporal
    
    SQ = "insert into rh_tmpempasis(nombres,codigo,divi,situacion,mon,bono) " & _
         "SELECT DISTINCT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo as cod, " & _
         "c.Division as divI,e.situacion,mon_bono,monto_bono From empleado e LEFT OUTER JOIN (select fecha,codemp from rh_bonosempleados) as t " & _
         "ON (e.codigo=t.codemp) and ifnull(DATE_FORMAT(ifnull(t.fecha,''),'%Y%m'),'') BETWEEN '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "' " & _
         " AND '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "' " & _
         "RIGHT OUTER JOIN (select if(ifnull(f_termino,'')<>'',f_termino,sysdate()) as f_termino,codemp,estado,mon_bono,monto_bono,codigo,division from contrato) as c " & _
         "ON (e.codigo=c.codemp) and c.monto_bono > 0 AND ((c.codigo = (select max(codigo) from contrato tt where tt.codemp = c.codemp) and ifnull(date_format(ifnull(c.f_termino,''),'%Y%m'),'') BETWEEN '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "'  AND '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "') " & _
         "or (c.estado='AP')) where (ifnull(e.fec_cese,'') = '' or date_format(ifnull(e.fec_cese,''),'%Y%m') >= '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "') and (left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) >= '" & LI & "' and left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) <= '" & LF & "') " & NomEmp & " group by e.codigo order by nombres"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
                
    AnioMes = "year(t.fecha)='" & AnoIni & "' and right(concat('00',month(t.fecha)),2)='" & MesIni & "'"
    SqLDiasAnt = DevDiasSql(AnioMes, "A")
        
    AnioMes = "year(t.fecha)='" & strAnoSistema & "' and right(concat('00',month(t.fecha)),2)='" & cboMes.List(cboMes.ListIndex, 2) & "'"
    SqlDiasAct = DevDiasSql(AnioMes, "D")
    
    AnioMes = "year(t.fecha)='" & AnoFin & "' and right(concat('00',month(t.fecha)),2)='" & MesFin & "'"
    SqLDiasDes = DevDiasSql(AnioMes, "P")
    
    SQ = "SELECT DISTINCT e.nombres, " & SqLDiasAnt & "," & SqlDiasAct & "," & SqLDiasDes & ",e.item,e.codigo as cod,e.divI,ifnull(c.fec_Salida,'') as fec_salida,IFNULL(c.fec_Regreso,'') as fec_regreso,e.situacion,ifnull(tot3.subtotal,0) as subtotal, " & _
         "ifnull(T.FECHA,'') as fecha,ifnull(tot1.totdias,0) as dias,ifnull((select descripcioncorta from novperuvhse.pozo where idpozo=t.pozo),'') as pozo,ifnull((select descripcioncorta from novperuvhse.lote where idlote=t.lote),'') as lote,if(ifnull(t.mon,'')='',e.mon,ifnull(t.mon,'')) as mon,if(ifnull(e.bono,'0')='0',t.bono,e.bono) as monto " & _
         "From rh_tmpempasis e LEFT OUTER JOIN (select fecha,codemp,pozo,lote,mon,bono from rh_bonosempleados) as t " & _
         "ON (e.codigo=t.codemp) and DATE_FORMAT(t.fecha,'%Y%m') BETWEEN '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "' " & _
         " AND '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "' " & _
         "left join (select codemp,fec_salida,fec_regreso from calendario where movemp = '02' and gocehaber = 'N') as c " & _
         "on (c.codemp=e.codigo) and ((concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "' " & _
         "and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2)) >= '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) & MesIni & "') OR " & _
         " (concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2)) >= '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "') OR " & _
         " (concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "' " & _
         "and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2)) >= '" & IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) & MesFin & "')) " & _
         "LEFT OUTER JOIN (select count(*) as totdias,codemp,fecha from rh_bonosempleados where fecha BETWEEN '" & Format(dtFecIni, "yyyy/mm/dd") & "' " & _
         " AND '" & Format(dtFecFin, "yyyy/mm/dd") & "' group by codemp) as tot1 ON (e.codigo=tot1.codemp) " & _
         "LEFT OUTER JOIN (select sum(bono) as subtotal,codemp,fecha from rh_bonosempleados where fecha BETWEEN '" & Format(dtFecIni, "yyyy/mm/dd") & "' " & _
         " AND '" & Format(dtFecFin, "yyyy/mm/dd") & "' group by codemp) as tot3 ON (e.codigo=tot3.codemp) Where e.codigo Is Not Null group by nombres,fecha,fec_salida"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    I = 1: k = 0
    If Not RQ.EOF() Then
        Do While Not RQ.EOF
            With MshEmpO
                If RQ.Fields("ITEM") = I Then
                    If k = 0 Then
                        FecIngreso = FechaPersonal(RQ.Fields("cod"))
                        FecTermino = FechaPersonal(RQ.Fields("cod"), 2)
                        .Rows = .Rows + 1
                        .TextMatrix(I + 2, 0) = H
                        .Col = 1: .row = I + 2
                        .CellBackColor = &H2B2826
                        .Col = 2: .row = I + 2
                        .CellBackColor = &H2B2826: .CellForeColor = vbWhite: .CellFontBold = True: .CellFontSize = 8.5
                        .Col = 3: .row = I + 2
                        .CellBackColor = &H2B2826: .CellForeColor = vbWhite: .CellFontBold = True: .CellFontSize = 8.5
                        .TextMatrix(I + 2, 1) = Trim(RQ.Fields("nombres"))
                        .TextMatrix(I + 2, 2) = RQ.Fields("dias")
                        .TextMatrix(I + 2, 3) = FormatNumber(RQ.Fields("monto"), 2)
                        .TextMatrix(I + 2, .Cols - 2) = RQ.Fields("cod")
                        .TextMatrix(I + 2, .Cols - 1) = RQ.Fields("divi")
                        .TextMatrix(I + 2, .Cols - 3) = FormatNumber(RQ.Fields("subtotal"), 2)
                        If val(.TextMatrix(I + 2, .Cols - 3)) = "0.00" Then .TextMatrix(I + 2, .Cols - 3) = "0"
                        If val(.TextMatrix(I + 2, 3)) = "0.00" Then .TextMatrix(I + 2, 3) = "0"
                        .TextMatrix(I + 2, .Cols - 4) = FecIngreso
                        If IsDate(RQ.Fields("fec_salida")) Then
                            DiaIniV = Day(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy"))
                            DiaFinV = Day(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            MesIniV = Month(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy"))
                            MesFinV = Month(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            AnioIniV = Trim(Year(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy")))
                            AnioFinV = Year(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            ColIni = DevPosCol(str(MesIniV), str(AnioIniV))
                            If AnioIniV = AnioFinV Then
                                If Trim(AnioIniV) & Right("00" & Trim(MesIniV), 2) = Trim(AnioFinV) & Right("00" & Trim(MesFinV), 2) Then
                                    For T = (ColIni + DiaIniV) - 1 To ((ColIni + DiaIniV) - 1) + (DiaFinV - DiaIniV)
                                        .TextMatrix(I + 2, T) = "V"
                                    Next
                                Else
                                    If Trim(AnioIniV & Right("00" & Trim(MesIniV), 2)) < Trim(AnoIni & Right("00" & Trim(MesIni), 2)) Then
                                        For T = ColIni To ((ColIni + DiaFinV) - 1)
                                            .TextMatrix(I + 2, T) = "V"
                                        Next
                                    Else
                                        dias = Day(DateSerial(AnioIniV, MesIniV + 1, 0)) - DiaIniV
                                        For T = (ColIni + DiaIniV) - 1 To ((ColIni + DiaIniV) - 1) + dias
                                            .TextMatrix(I + 2, T) = "V"
                                        Next
                                        For T = ((ColIni + DiaIniV) - 1) + dias + 1 To ((ColIni + DiaIniV) - 1) + dias + DiaFinV
                                            If T <= .Cols - 5 Then
                                                .TextMatrix(I + 2, T) = "V"
                                            End If
                                        Next
                                    End If
                                End If
                            Else
                                dias = Day(DateSerial(AnioIniV, MesIniV + 1, 0)) - DiaIniV
                                For T = (ColIni + DiaIniV) - 1 To ((ColIni + DiaIniV) - 1) + dias
                                    .TextMatrix(I + 2, T) = "V"
                                Next
                                For T = ((ColIni + DiaIniV) - 1) + dias + 1 To ((ColIni + DiaIniV) - 1) + dias + DiaFinV
                                    .TextMatrix(I + 2, T) = "V"
                                Next
                            End If
                        End If
                        .Col = 1: .row = I + 2: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        .Col = 0: .row = I + 2: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        If RQ.Fields("situacion") = 0 Then
                            ColIni = DevPosCol(str(Month(Format(FecTermino, "dd/mm/yyyy"))), str(Year(Format(FecTermino, "dd/mm/yyyy"))))
                            If ColIni > 0 Then
                                .Col = 0: .row = I + 2: .CellForeColor = vbRed
                                .Col = 1: .row = I + 2: .CellForeColor = vbRed
                                MarcarDias ColIni + Day(FecTermino), .Cols - 5, I + 2
                            End If
                        End If
                        ColIni = DevPosCol(str(Month(Format(FecIngreso, "dd/mm/yyyy"))), str(Year(Format(FecIngreso, "dd/mm/yyyy"))), 1)
                        If ColIni > 0 Then
                            Dim FecTermAnt As String
                            MarcarDias IIf(ColIni = ColIniMesAct Or ColIni = ColIniMesPos, 4, ColIni), (Day(FecIngreso) + ColIni) - 2, I + 2
                            FecTermAnt = VerificaFec(Trim(RQ.Fields("cod")), FecIngreso)
                            If FecTermAnt <> "" Then
                                DesMarcarDias IIf(ColIni = ColIniMesAct Or ColIni = ColIniMesPos, 4, ColIni), Trim(Day(FecTermAnt) + IIf(ColIni = ColIniMesAct Or ColIni = ColIniMesPos, 4, ColIni)) - 1, I + 2
                            End If
                        End If
                        k = k + 1
                    End If
                    .row = I + 2
                    If IsDate(RQ.Fields("FECHA")) Then
                        Dim dia As String, campo As String, letra As String, NumMes As String
                        Dim Cant1, Cant2, Cant3 As Integer
                        Cant1 = 0
                        Cant2 = 0
                        Cant3 = 0
                        For Y = 4 To .Cols - 5
                            If ColIniMesAnt <= Y Then
                                letra = "A": If Cant1 = 0 Then dia = 1: Cant1 = Cant1 + 1
                                NumMes = MesIni
                                AnioMes = AnoIni
                            End If
                            If ColIniMesAct <= Y Then
                                letra = "D": If Cant2 = 0 Then dia = 1: Cant2 = Cant2 + 1
                                NumMes = cboMes.List(cboMes.ListIndex, 2)
                                AnioMes = strAnoSistema
                            End If
                            If ColIniMesPos <= Y Then
                                letra = "P": If Cant3 = 0 Then dia = 1: Cant3 = Cant3 + 1
                                NumMes = MesFin
                                AnioMes = AnoFin
                            End If
                            campo = letra & dia
                            .TextMatrix(I + 2, Y) = IIf(.TextMatrix(I + 2, Y) = "", Trim(RQ.Fields(campo)), .TextMatrix(I + 2, Y))
                            .Col = Y: If .TextMatrix(I + 2, Y) <> "V" And (Trim(RQ.Fields("FECHA")) = AnioMes & "/" & NumMes & "/" & Right("00" & Trim(.TextMatrix(2, Y)), 2)) Then .CellForeColor = DevColor(RQ.Fields("pozo"), RQ.Fields("lote"))
                            dia = dia + 1
                        Next
                    End If
                    RQ.MoveNext
                Else
                    k = 0
                    I = I + 1
                    H = H + 1
                End If
            End With
        Loop
    End If
    H = H + 1
    If I > 1 And Trim(MshEmpO.TextMatrix(MshEmpO.Rows - 1, 1)) = "" Then
        MshEmpO.Rows = MshEmpO.Rows - 1
    End If
    If MshEmpO.Visible = True Then MshEmpO.SetFocus
    MshEmpO.Redraw = True
    Msf.Redraw = True
    CalculoTotal
    TotalGeneral
    Screen.MousePointer = vbDefault
    Set RQ = Nothing
End Sub

Function VerificaFec(CodEmp As String, FecIng As String) As String
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer

    SQ = "select IFNULL(max(f_termino),'') as fecha from contrato t where codigo < (SELECT codigo from contrato c " & _
         "where c.codemp = t.codemp and f_inicio = '" & FecIng & "') and t.codemp = '" & CodEmp & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    If Not RQ.EOF() Then
        VerificaFec = Trim(RQ.Fields("fecha"))
    End If
    Set RQ = Nothing
End Function

Function DevDiasSql(AnioMes As String, letra As String)
    DevDiasSql = "IF(" & AnioMes & " and DAY(T.FECHA) = '01','T','') AS " & letra & "1,IF(" & AnioMes & " and DAY(T.FECHA) = '02','T','') AS " & letra & "2, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '03','T','') AS " & letra & "3,IF(" & AnioMes & " and DAY(T.FECHA) = '04','T','') AS " & letra & "4, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '05','T','') AS " & letra & "5,IF(" & AnioMes & " and DAY(T.FECHA) = '06','T','') AS " & letra & "6, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '07','T','') AS " & letra & "7,IF(" & AnioMes & " and DAY(T.FECHA) = '08','T','') AS " & letra & "8, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '09','T','') AS " & letra & "9,IF(" & AnioMes & " and DAY(T.FECHA) = '10','T','') AS " & letra & "10, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '11','T','') AS " & letra & "11,IF(" & AnioMes & " and DAY(T.FECHA) = '12','T','') AS " & letra & "12, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '13','T','') AS " & letra & "13,IF(" & AnioMes & " and DAY(T.FECHA) = '14','T','') AS " & letra & "14, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '15','T','') AS " & letra & "15,IF(" & AnioMes & " and DAY(T.FECHA) = '16','T','') AS " & letra & "16, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '17','T','') AS " & letra & "17,IF(" & AnioMes & " and DAY(T.FECHA) = '18','T','') AS " & letra & "18, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '19','T','') AS " & letra & "19,IF(" & AnioMes & " and DAY(T.FECHA) = '20','T','') AS " & letra & "20, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '21','T','') AS " & letra & "21,IF(" & AnioMes & " and DAY(T.FECHA) = '22','T','') AS " & letra & "22, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '23','T','') AS " & letra & "23,IF(" & AnioMes & " and DAY(T.FECHA) = '24','T','') AS " & letra & "24, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '25','T','') AS " & letra & "25,IF(" & AnioMes & " and DAY(T.FECHA) = '26','T','') AS " & letra & "26, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '27','T','') AS " & letra & "27,IF(" & AnioMes & " and DAY(T.FECHA) = '28','T','') AS " & letra & "28, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '29','T','') AS " & letra & "29,IF(" & AnioMes & " and DAY(T.FECHA) = '30','T','') AS " & letra & "30, " & _
                 "IF(" & AnioMes & " and DAY(T.FECHA) = '31','T','') AS " & letra & "31"
End Function

Function DevPosCol(fec As String, Anio As String, Optional Tipo As Integer) As Integer
    Select Case Right("00" & Trim(fec), 2)
        Case MesIni
            If IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) = Trim(Anio) Then
                DevPosCol = ColIniMesAnt
            End If
        Case cboMes.List(cboMes.ListIndex, 2)
            If strAnoSistema = Trim(Anio) Then
                DevPosCol = ColIniMesAct
            End If
        Case MesFin
            If IIf(cboMes.List(cboMes.ListIndex, 2) = "12", strAnoSistema + 1, strAnoSistema) = Trim(Anio) Then
                DevPosCol = ColIniMesPos
            End If
        Case MesIni - 1
            If Tipo = 1 Then
                DevPosCol = 0
            Else
                If IIf(cboMes.List(cboMes.ListIndex, 2) = "01", strAnoSistema - 1, strAnoSistema) = Trim(Anio) Then
                    DevPosCol = 4
                End If
            End If
        Case Else
            DevPosCol = 0
    End Select
End Function

Sub MarcarDias(DiaIni As Integer, DiaFin As Integer, fila As Integer)
    With MshEmpO
        For J = DiaIni To DiaFin
            .Col = J: .row = fila: .CellBackColor = &H7E7B72
        Next
    End With
End Sub

Sub DesMarcarDias(DiaIni As Integer, DiaFin As Integer, fila As Integer)
    With MshEmpO
        For J = DiaIni To DiaFin
            .Col = J: .row = fila: .CellBackColor = &H2B2826
        Next
    End With
End Sub
' Colorea al Empleado ubicado en un determinado lote
Function DevColor(pozo As String, Lote As String) As Variant
Dim k As Integer
    For k = 0 To CboLote.ListCount - 1
        If Lote = Trim(CboLote.List(k, 1)) Then
            CboLote.ListIndex = k
            Exit For
        Else
            CboLote.ListIndex = 0
        End If
    Next
    For k = 0 To LstSede.ListCount - 1
        Msf.Col = 0: Msf.row = k
        If LstSede.List(k, 1) = pozo Then
            DevColor = Msf.CellBackColor
            Exit For
        End If
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("¿Seguro que desea salir del Control de Refrigerios?", vbQuestion + vbYesNo) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub MshEmpO_Click()
    With MshEmpO
        If .ColSel = 1 Then
            LstSede.ListIndex = 0
            If .MouseRow <> 0 Then
                Exit Sub
            End If
            Ordenar_Columna_FlexGrid MshEmpO, .MouseCol
        End If
        If .ColSel >= 2 And .ColSel <= .Cols - 9 Then
            If ColIniO = .ColSel Then
                .Redraw = False
                Msf.Redraw = False
                .row = .Rowsel: .Col = .ColSel
                If .CellForeColor <> &HFFFFFF Then
                    If .TextMatrix(.row, .Col) <> "" Then
                        MarcaSede .row, .Col
                    End If
                End If
                Msf.Redraw = True
                .Redraw = True
            End If
        Else
            LstSede.ListIndex = 0
        End If
    End With
End Sub

' Marcar Sede
Sub MarcaSede(fila As Integer, columna As Integer)
Dim k As Integer
Dim SQ As String
Dim RQ As MYSQL_RS
Dim fec As String, Numdia As Integer, I As Integer
Dim Anio As String, Mes As String
    If ColIniMesAnt <= columna Then
        Anio = AnoIni
        Mes = MesIni
    End If
    If ColIniMesAct <= columna Then
        Anio = strAnoSistema
        Mes = cboMes.List(cboMes.ListIndex, 2)
    End If
    If ColIniMesPos <= columna Then
        Anio = AnoFin
        Mes = MesFin
    End If
    With MshEmpO
        SQ = "select * from rh_bonosempleados where codemp = '" & .TextMatrix(fila, .Cols - 2) & "' and " & _
             "fecha = '" & Anio & "/" & Mes & "/" & Right("00" & Trim(.TextMatrix(2, columna)), 2) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        If Not RQ.EOF() Then
            For I = 0 To CboLote.ListCount - 1
                If CE(RQ.Fields("lote")) = Trim(CboLote.List(I, 2)) Then
                    CboLote.ListIndex = I
                    Exit For
                Else
                    CboLote.ListIndex = 0
                End If
            Next
            For k = 0 To Msf.Rows - 1
                .Col = columna: .row = fila
                Msf.Col = 0: Msf.row = k
                If .CellForeColor = Msf.CellBackColor Then
                    LstSede.ListIndex = k
                    Exit For
                End If
            Next
        End If
    End With
    Set RQ = Nothing
End Sub

Private Sub MshEmpO_KeyPress(KeyAscii As Integer)
Dim TKeyascii As Integer
    If KeyAscii <> 13 Then
        TKeyascii = KeyAscii
        CadenaBusqueda MshEmpO, KeyAscii, cadenaemp, "1"
    End If
End Sub

Sub CadenaBusqueda(Msh As MSHFlexGrid, KeyAscii As Integer, cad As String, Tipo As String)
    If Msh.ColSel < 2 Then
        Dim c%, T%, a$, B$
        If KeyAscii = 27 Then
            cad = "": lblCadBus = ""
            Exit Sub
        End If
        If KeyAscii >= 32 Or KeyAscii = 8 Then
            cad = cad & Chr(KeyAscii)
            lblCadBus = cad
            With Msh
                If KeyAscii <> 8 Then
                    c = Len(cad)
                    If IsNumeric(cad) Then
                        a = Right("00000000000" & Trim(cad), 11)
                    Else
                        a = cad
                    End If
                End If
                If c >= 1 Then
                    For T = 1 To .Rows - 1
                        If IsNumeric(cad) Then
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
                                If Tipo = "1" Then MshEmpO_RowColChange
                                Exit For
                            End If
                        End If
                    Next T
                End If
            End With
        End If
    End If
End Sub

Private Sub MshEmpO_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim flgdias As Boolean, FlgNoDias As Boolean
    With MshEmpO
        .Redraw = False
        If Button = 1 Then
            GridSel = 1
            If .ColSel > 1 Then
                .MergeCells = flexMergeNever
                ColIniO = .ColSel
                FilaIniO = .Rowsel
            Else
                .MergeCells = flexMergeRestrictColumns
            End If
        ElseIf Button = 2 Then
            Dim I As Integer, Flg As Boolean
            Dim k As Integer
            GridSel = 1
            If .Rows > 1 And Trim(.TextMatrix(3, 1)) <> "" And (.ColSel > 1) And ColIniO > 0 Then
                k = FilaIniO
                If val(TempTrab(FilaIniO)) = 0 Then
                    TempTrab(FilaIniO) = FilaIniO
                    CantTrab = CantTrab + 1
                    Trab(CantTrab) = FilaIniO
                End If
                flgdias = False
                If Validadias(k, ColIniO, .ColSel) = False Then
                    Flg = False
                    For I = ColIniO To .ColSel
                        If Trim(.TextMatrix(k, I)) = "" Then
                            Flg = True
                            Exit For
                        End If
                    Next
                    If Flg = True Then
                        If MsgBox("Existen algunos dias que el trabajador no tiene registrada asistencia." & Chr(13) & _
                                  "¿Desea registrar el refrigerio de esos días?", vbQuestion + vbYesNo, "NOVPeru") = vbNo Then
                            flgdias = True
                        End If
                    End If
                End If
                For I = ColIniO To .ColSel
                    If (LstSede.List(LstSede.ListIndex, 1) <> "00000000000") And (.TextMatrix(k, I) <> "V") Then
                        FlgNoDias = False
                        If flgdias = True Then
                            If Valida(k, I) = True Then
                                FlgNoDias = True
                            End If
                        Else
                            FlgNoDias = True
                        End If
                        If FlgNoDias = True Then
                            .row = k: .Col = I
                            If .CellBackColor <> &H7E7B72 Then
                                If Trim(.TextMatrix(k, I)) <> "" Then
                                    If I <= .Cols - 5 Then
                                        .TextMatrix(k, I) = ""
                                    End If
                                Else
                                    .row = k: .Col = I
                                    .CellForeColor = &HFFFFFF
                                    If Left(Trim(.TextMatrix(k, I)), 1) <> "0" Then
                                        .row = k: .Col = I
                                        If LstSede.ListIndex > -1 Then
                                            Msf.row = LstSede.ListIndex
                                            Msf.Col = 0
                                            .CellForeColor = Msf.CellBackColor
                                            .TextMatrix(k, I) = "T"
                                        Else
                                            .TextMatrix(k, I) = ""
                                        End If
                                    End If
                                End If
                                .Col = I: .row = k
                                .CellFontSize = 8.5
                                .CellFontBold = True
                            End If
                        End If
                    End If
                Next
                GrabaBonos
AQUI:
                CalculoTotal
                TotalGeneral
                LimpiaArreglos
                CantTrab = 0
                FilaIniO = 0
                ColIniO = 0
                .MergeCells = flexMergeRestrictColumns
                .Redraw = True
            End If
        End If
        .Redraw = True
    End With
End Sub

Function Validadias(fila As Integer, ColIni As Integer, ColFin As Integer) As Boolean
Dim SQ As String
Dim RQ As MYSQL_RS
Dim fec As String, Numdia As Integer
    With MshEmpO
        SQ = "select * from rh_entsalempleado where emp = '" & .TextMatrix(fila, .Cols - 2) & "' and tipo = 'E' and " & _
             "fecha between '" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Trim(.TextMatrix(2, ColIni)), 2) & "' and " & _
             "'" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Trim(.TextMatrix(2, ColFin)), 2) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        Numdia = Trim(.TextMatrix(2, ColIni))
        If Not RQ.EOF() Then
            Do While Not RQ.EOF()
                fec = strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Numdia, 2)
                If fec = Trim(RQ.Fields("fecha")) Then
                    Numdia = Numdia + 1
                Else
                    Validadias = False
                    Set RQ = Nothing
                    Exit Function
                End If
                RQ.MoveNext
            Loop
        Else
            Validadias = False
            Set RQ = Nothing
            Exit Function
        End If
        Validadias = True
        Set RQ = Nothing
    End With
End Function

Function Valida(fila As Integer, Col As Integer) As Boolean
Dim SQ As String
Dim RQ As MYSQL_RS
Dim Anio As String, Mes As String
    If ColIniMesAnt <= Col Then
        Anio = AnoIni
        Mes = MesIni
    End If
    If ColIniMesAct <= Col Then
        Anio = strAnoSistema
        Mes = cboMes.List(cboMes.ListIndex, 2)
    End If
    If ColIniMesPos <= Col Then
        Anio = AnoFin
        Mes = MesFin
    End If
    Valida = False
    With MshEmpO
        SQ = "select * from rh_entsalempleado where emp = '" & .TextMatrix(fila, .Cols - 2) & "' and " & _
             "fecha = '" & Anio & "/" & Mes & "/" & Right("00" & Trim(.TextMatrix(2, Col)), 2) & "' and tipo = 'E'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        If Not RQ.EOF() Then
            Valida = True
        End If
        Set RQ = Nothing
    End With
End Function

Sub LimpiaArreglos()
Dim I As Integer
    For I = 1 To CantTrab
        TempTrab(Trab(I)) = 0
        Trab(I) = 0
    Next
End Sub

Private Sub MshEmpO_RowColChange()
    DesplazarporGrid MshEmpO.Rowsel
    MshEmpO.Refresh
End Sub

Sub DesplazarporGrid(fila As Integer)
    With MshEmpO
        lblCodEmp = Trim(.TextMatrix(fila, .Cols - 2))
        lbldiv = DescripcionesdeCodigos("DES_DIVISION", .TextMatrix(fila, .Cols - 1), "Descrip")
    End With
End Sub

Private Sub MshEmpO_Scroll()
    MshEmpO.Refresh
End Sub

Private Sub TabS_Click()
Dim LI As String, LF As String
    If cboMes.ListIndex > -1 Then
        Select Case TabS.SelectedItem.Index
            Case 1: LI = "A": LF = "B"
            Case 2: LI = "C": LF = "D"
            Case 3: LI = "E": LF = "F"
            Case 4: LI = "G": LF = "H"
            Case 5: LI = "I": LF = "J"
            Case 6: LI = "K": LF = "L"
            Case 7: LI = "M": LF = "Ñ"
            Case 8: LI = "O": LF = "P"
            Case 9: LI = "Q": LF = "R"
            Case 10: LI = "S": LF = "T"
            Case 11: LI = "U": LF = "V"
            Case 12: LI = "W": LF = "X"
            Case 13: LI = "Y": LF = "Z"
        End Select
        CargaEmpleados "", UCase(Trim(TxtNom)), LI, LF
        MshEmpO.LeftCol = ColIniMesAct
        If cboDiv.ListIndex > -1 Then
            cbodiv_Change
        End If
    End If
End Sub

Sub CrearTablaTemporal()
On Error GoTo CtrlError
Dim SQL As String
    If CReaTabla = True Then
        SQL = "drop table rh_tmpempasis"
        oConexion.EjecutaSelectRS (SQL)
    End If
    SQL = "create table rh_tmpempasis( Item INT NOT NULL AUTO_INCREMENT, nombres char(100),codigo char(11),divi char(4), " & _
          "situacion char(1),mon char(1),bono double(15,3),PRIMARY KEY  (ITEM),UNIQUE KEY ITEM (ITEM))"
    oConexion.EjecutaSelectRS (SQL)
    CReaTabla = True
Exit Sub
CtrlError:
    CReaTabla = False
    MsgBox err.Description, "Error Creando Tabla temporal", "NOVPeru"
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next

    With MshEmpO
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
