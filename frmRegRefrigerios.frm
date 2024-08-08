VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Object = "{495FAA50-CD31-4123-AF30-347C5038B9CD}#1.0#0"; "menu_list.ocx"
Begin VB.Form frmRegRefrigerios 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Refrigerios"
   ClientHeight    =   7875
   ClientLeft      =   2310
   ClientTop       =   4845
   ClientWidth     =   15165
   Icon            =   "frmRegRefrigerios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   15165
   Begin Menu_List.Menu_ListView Menu 
      Height          =   315
      Left            =   420
      TabIndex        =   47
      Top             =   7530
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
   End
   Begin VB.CheckBox ChkCena 
      BackColor       =   &H009F5539&
      Caption         =   "[C]ena"
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
      Left            =   12660
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CheckBox Chkalmuerzo 
      BackColor       =   &H009F5539&
      Caption         =   "[A]lmuerzo"
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
      Left            =   9900
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Frame frmce 
      BackColor       =   &H009F5539&
      Caption         =   "[C]ena"
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
      Height          =   600
      Left            =   11925
      TabIndex        =   34
      Top             =   30
      Width           =   3105
      Begin VB.TextBox TxtValC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2355
         MaxLength       =   12
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox TxtFactC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   735
         MaxLength       =   12
         TabIndex        =   35
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Factura"
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
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   38
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   1860
         TabIndex        =   37
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame frmal 
      BackColor       =   &H009F5539&
      Caption         =   "[A]lmuerzo"
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
      Height          =   600
      Left            =   8850
      TabIndex        =   30
      Top             =   30
      Width           =   3090
      Begin VB.TextBox TxtFactA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   750
         MaxLength       =   12
         TabIndex        =   32
         Top             =   240
         Width           =   1110
      End
      Begin VB.TextBox TxtValA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2355
         MaxLength       =   12
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Factura"
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
         Height          =   195
         Left            =   75
         TabIndex        =   39
         Top             =   285
         Width           =   675
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1890
         TabIndex        =   33
         Top             =   285
         Width           =   450
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshEmpO 
      Height          =   5790
      Left            =   -15
      TabIndex        =   18
      Top             =   1020
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   10213
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   16777215
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      ForeColorSel    =   16777215
      GridColor       =   3485999
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
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
      Left            =   0
      TabIndex        =   4
      Top             =   68
      Width           =   8835
      Begin VB.TextBox TxtNom 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6855
         TabIndex        =   41
         Top             =   150
         Width           =   1920
      End
      Begin VB.TextBox TxtImp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5415
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   150
         Width           =   585
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
         TabIndex        =   43
         Top             =   165
         Width           =   825
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
         TabIndex        =   42
         Top             =   165
         Width           =   510
      End
      Begin MSForms.ComboBox CboMes 
         Height          =   315
         Left            =   495
         TabIndex        =   8
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
         Top             =   165
         Width           =   420
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
         TabIndex        =   6
         Top             =   165
         Width           =   735
      End
      Begin MSForms.ComboBox CboDiv 
         Height          =   315
         Left            =   2790
         TabIndex        =   5
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
   End
   Begin MSFlexGridLib.MSFlexGrid Msf 
      Height          =   6135
      Left            =   14460
      TabIndex        =   1
      Top             =   1020
      Width           =   690
      _ExtentX        =   1217
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   14625
      TabIndex        =   9
      Top             =   7350
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
      MICON           =   "frmRegRefrigerios.frx":030A
      PICN            =   "frmRegRefrigerios.frx":0326
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
      Left            =   11985
      TabIndex        =   10
      ToolTipText     =   "Guardar"
      Top             =   8880
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
      MICON           =   "frmRegRefrigerios.frx":06EC
      PICN            =   "frmRegRefrigerios.frx":0708
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
      Left            =   15
      TabIndex        =   0
      Top             =   1035
      Width           =   12525
      _ExtentX        =   22093
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
      Left            =   3480
      TabIndex        =   23
      Top             =   7170
      Width           =   5445
      Begin VB.TextBox txtValorRefri 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4020
         MaxLength       =   12
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dptFecIniR 
         Height          =   315
         Left            =   765
         TabIndex        =   26
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
         TabIndex        =   27
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
         TabIndex        =   29
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
         MICON           =   "frmRegRefrigerios.frx":0B4A
         PICN            =   "frmRegRefrigerios.frx":0B66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         TabIndex        =   25
         Top             =   240
         Width           =   540
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
         TabIndex        =   24
         Top             =   255
         Width           =   690
      End
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
      Left            =   9630
      TabIndex        =   46
      Top             =   7365
      Width           =   1290
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
      Left            =   9060
      TabIndex        =   45
      Top             =   7425
      Width           =   465
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   " A | C "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   14655
      TabIndex        =   44
      Top             =   795
      Width           =   495
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
      Left            =   60
      TabIndex        =   22
      Top             =   750
      Width           =   720
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
      Left            =   8940
      TabIndex        =   21
      Top             =   750
      Width           =   810
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
      Left            =   11205
      TabIndex        =   20
      Top             =   7402
      Width           =   1455
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
      Left            =   12690
      TabIndex        =   19
      Top             =   7365
      Width           =   1905
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
      Left            =   45
      TabIndex        =   17
      Top             =   7230
      Width           =   3390
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
      Left            =   75
      TabIndex        =   16
      Top             =   7575
      Width           =   75
   End
   Begin MSForms.ListBox LstSede 
      Height          =   6150
      Left            =   12510
      TabIndex        =   15
      Top             =   1005
      Width           =   1950
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3440;10433"
      MatchEntry      =   0
      ListStyle       =   1
      FontName        =   "Arial"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbldato 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   240
      Left            =   13455
      TabIndex        =   14
      Top             =   9030
      Visible         =   0   'False
      Width           =   225
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
      Left            =   9855
      TabIndex        =   12
      Top             =   735
      Width           =   2610
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
      Left            =   840
      TabIndex        =   11
      Top             =   750
      Width           =   2640
   End
   Begin MSForms.Label lblo 
      Height          =   360
      Left            =   15
      TabIndex        =   13
      Top             =   660
      Width           =   12495
      ForeColor       =   128
      Caption         =   "Listado  de  Empleados"
      PicturePosition =   393216
      Size            =   "22040;635"
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
Attribute VB_Name = "frmRegRefrigerios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColIniO As Integer, FilaIniO As Integer
Dim emp As String, OpMov As Boolean
Dim Datos As String, CantTrab As Integer
Public GridSel As Integer
Dim cadenaemp As String, dias As Integer
Dim Entro As Boolean, FilaSel As Integer
Private Trab(1 To 500) As Integer
Private TempTrab(1 To 500) As Integer
Dim EntroBusqueda As Boolean, H As Integer
Dim CReaTabla As Boolean

Sub GrabaRefrigerios(ColIni As Integer, ColFin As Integer)
On Error GoTo CtrlError
    Dim I As Integer, J As Integer, k As Integer
    Dim SQ As String, rest As String, aux As String, codAux As String
    Dim TotA As Integer, TotC As Integer, SubTot As Double
    Dim ValorA As Double, ValorC As Double, Tipo As String
    Dim FlgIngDel As Boolean, auxa As String, codauxa As String, resta As String
    Dim AL As Integer, CE As Integer
    Screen.MousePointer = vbHourglass
    With MshEmpO
        For k = 1 To CantTrab
            I = Trab(k) ' - 1
            SubTot = CDbl(IIf(val(.TextMatrix(I, .Cols - 3)) = 0, 0, .TextMatrix(I, .Cols - 3)))
            FlgIngDel = False
            AL = 0
            CE = 0
            resta = ""
            codauxa = ""
            auxa = ""
            rest = ""
            codAux = ""
            aux = ""
            For J = ColIni To ColFin
                TotA = 0
                TotC = 0
                If Trim(.TextMatrix(I, J)) = "C" Or Trim(.TextMatrix(I, J)) = "A" Or Trim(.TextMatrix(I, J)) = "S" Then
                    .row = I: .Col = J
                    If .CellForeColor <> &HFFFFFF Then
                        If Trim(.TextMatrix(I, J)) = "S" Then
                            resta = DevCodAuxil(2)
                            rest = DevCodAuxil(3)
                            If resta = "" And rest = "" Then
                                rest = DevCodAux(I, J)
                                codAux = Mid(rest, 1, InStr(1, rest, "-") - 1)
                                aux = Mid(rest, InStr(1, rest, "-") + 1, Len(rest) - 1)
                                resta = rest
                                codauxa = codAux
                                auxa = aux
                            Else
                                If resta <> "" Then
                                    codauxa = Mid(resta, 1, InStr(1, resta, "-") - 1)
                                    auxa = Mid(resta, InStr(1, resta, "-") + 1, Len(resta) - 1)
                                End If
                                If rest <> "" Then
                                    codAux = Mid(rest, 1, InStr(1, rest, "-") - 1)
                                    aux = Mid(rest, InStr(1, rest, "-") + 1, Len(rest) - 1)
                                End If
                            End If
                        Else
                            If Trim(.TextMatrix(I, J)) = "A" Then
                                resta = DevCodAux(I, J)
                                codauxa = Mid(resta, 1, InStr(1, resta, "-") - 1)
                                auxa = Mid(resta, InStr(1, resta, "-") + 1, Len(resta) - 1)
                            End If
                            If Trim(.TextMatrix(I, J)) = "C" Then
                                rest = DevCodAux(I, J)
                                codAux = Mid(rest, 1, InStr(1, rest, "-") - 1)
                                aux = Mid(rest, InStr(1, rest, "-") + 1, Len(rest) - 1)
                            End If
                        End If
                        ValorA = 0: ValorC = 0
                        If Trim(.TextMatrix(I, J)) = "A" Or Trim(.TextMatrix(I, J)) = "S" Then
                            ValorA = FormatNumber(CDbl(IIf(val(TxtValA) = 0, 0, TxtValA)), 2)
                        End If
                        If Trim(.TextMatrix(I, J)) = "C" Or Trim(.TextMatrix(I, J)) = "S" Then
                            ValorC = FormatNumber(CDbl(IIf(val(TxtValC) = 0, 0, TxtValC)), 2)
                        End If
                        If Trim(.TextMatrix(I, J)) = "S" Then
                            SQ = "call Insert_Refrigerios('" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                                 "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                                 "'A','" & auxa & "','" & codauxa & "', " & _
                                 "'" & Trim(TxtFactA) & "'," & CDbl(ValorA) & ",'','','',0)"
                            oConexionMYSQL.Execute SQ
                            SQ = "call Insert_Refrigerios('" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                                 "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                                 "'C','" & aux & "','" & codAux & "', " & _
                                 "'" & Trim(TxtFactC) & "'," & CDbl(ValorC) & ",'','','',0)"
                            oConexionMYSQL.Execute SQ
                        Else
                            SQ = "call Insert_Refrigerios('" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                                 "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                                 "'" & Trim(.TextMatrix(I, J)) & "','" & IIf(Trim(.TextMatrix(I, J)) = "C", aux, auxa) & "','" & IIf(Trim(.TextMatrix(I, J)) = "C", codAux, codauxa) & "', " & _
                                 "'" & IIf(Trim(.TextMatrix(I, J)) = "C", Trim(TxtFactC), Trim(TxtFactA)) & "'," & IIf(Trim(.TextMatrix(I, J)) = "C", CDbl(ValorC), CDbl(ValorA)) & ",'','','',0)"
                            oConexionMYSQL.Execute SQ
                        End If
                    End If
                    If Trim(.TextMatrix(I, J)) = "A" Or Trim(.TextMatrix(I, J)) = "S" Then TotA = TotA + 1: AL = AL + 1
                    If Trim(.TextMatrix(I, J)) = "C" Or Trim(.TextMatrix(I, J)) = "S" Then TotC = TotC + 1: CE = CE + 1
                    SubTot = SubTot + (TotA * ValorA) + (TotC * ValorC)
                Else
                    'Tipo = DevTipoDia("tipo", Trim(.TextMatrix(I, .Cols - 2)), Format(CDate(strAnoSistema & "/" & CboMes.List(CboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd"))
                    Dim jj As Integer
                    For jj = 1 To 2
                        Tipo = IIf(jj = 1, "A", "C")
                        ValorC = DevTipoDia("C", Trim(.TextMatrix(I, .Cols - 2)), Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd"))
                        ValorA = DevTipoDia("A", Trim(.TextMatrix(I, .Cols - 2)), Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd"))
                        SQ = "delete from rh_refriempleados where fecha = '" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "' " & _
                             "and codemp = '" & .TextMatrix(I, .Cols - 2) & "' and tipo = '" & Tipo & "'"
                        oConexionMYSQL.Execute SQ
                        If Tipo = "A" Or Tipo = "S" Then TotA = TotA + 1: AL = AL + 1
                        If Tipo = "C" Or Tipo = "S" Then TotC = TotC + 1: CE = CE + 1
                        If Tipo <> "" Then
                            SubTot = CDbl(Trim(SubTot) - ((TotA * ValorA) + (TotC * ValorC)))
                            FlgIngDel = True
                        End If
                    Next
                End If
            Next
        Next
        .TextMatrix(I, .Cols - 5) = CDbl(IIf(val(.TextMatrix(I, .Cols - 5)) = 0, 0, .TextMatrix(I, .Cols - 5))) + IIf(FlgIngDel = True, CE * (-1), CE)
        .TextMatrix(I, .Cols - 6) = CDbl(IIf(val(.TextMatrix(I, .Cols - 6)) = 0, 0, .TextMatrix(I, .Cols - 6))) + IIf(FlgIngDel = True, AL * (-1), AL)
        .TextMatrix(I, .Cols - 3) = SubTot
    End With
    Screen.MousePointer = vbDefault
Exit Sub
CtrlError:
    MsgBox "No se pudo grabar la información para el empleado:" & emp & vbNewLine & _
           "revise los refrigerios o consulte con el administrador del sistema", vbOKOnly + vbExclamation, "NOVPeru"
End Sub
Function DevTipoDia(campo As String, Cod As String, fecha As String) As Variant
Dim SQL As String
Dim RQ As MYSQL_RS
    DevTipoDia = 0
    SQL = "select montoa as cmpo from rh_refriempleados where fecha = '" & fecha & "' and codemp = '" & Cod & "' and tipo = '" & campo & "' order by tipo"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        DevTipoDia = Trim(RQ.Fields("cmpo"))
    End If
Set RQ = Nothing
End Function
Function DevCodAuxil(columna As Integer) As String
Dim a As Integer, c As Integer
Dim fLA As Boolean, FLC As Boolean
    With Msf
        For a = 0 To .Rows - 1
            If .TextMatrix(a, IIf(columna = 2, columna, columna - 1)) = strChecked Then
                fLA = True
                Exit For
            End If
        Next
        For c = 0 To .Rows - 1
            If .TextMatrix(c, IIf(columna = 2, columna + 1, columna)) = strChecked Then
                FLC = True
                Exit For
            End If
        Next
        If fLA = True And FLC = True Then
            DevCodAuxil = IIf(columna = 2, LstSede.List(a, 1) & "-" & LstSede.List(a, 2), LstSede.List(c, 1) & "-" & LstSede.List(c, 2))
        End If
    End With
End Function

Function Valida(Tipo As String) As Boolean
    If Tipo = "A" Then
        If Trim(TxtFactA) = "" Then
            MsgBox "Ingrese el Número de Factura del Almuerzo", vbInformation, "NOVPeru"
            TxtFactA.SetFocus
            Valida = False
            Exit Function
        End If
        If val(TxtValA) = 0 Then
            MsgBox "Ingrese el Valor de la Estancia para el Almuerzo", vbInformation, "NOVPeru"
            TxtValA.SetFocus
            Valida = False
            Exit Function
        End If
    End If
    If Tipo = "C" Then
        If Trim(TxtFactC) = "" Then
            MsgBox "Ingrese el Número de Factura la Cena", vbInformation, "NOVPeru"
            TxtFactC.SetFocus
            Valida = False
            Exit Function
        End If
        If val(TxtValC) = 0 Then
            MsgBox "Ingrese el Valor de la Estancia para la Cena", vbInformation, "NOVPeru"
            TxtValC.SetFocus
            Valida = False
            Exit Function
        End If
    End If
    
    If Tipo = "S" Then
        If Trim(TxtFactA) = "" Then
            MsgBox "Ingrese el Número de Factura del Almuerzo", vbInformation, "NOVPeru"
            TxtFactA.SetFocus
            Valida = False
            Exit Function
        End If
        If val(TxtValA) = 0 Then
            MsgBox "Ingrese el Valor de la Estancia para el Almuerzo", vbInformation, "NOVPeru"
            TxtValA.SetFocus
            Valida = False
            Exit Function
        End If
        If Trim(TxtFactC) = "" Then
            MsgBox "Ingrese el Número de Factura la Cena", vbInformation, "NOVPeru"
            TxtFactC.SetFocus
            Valida = False
            Exit Function
        End If
        If val(TxtValC) = 0 Then
            MsgBox "Ingrese el Valor de la Estancia para la Cena", vbInformation, "NOVPeru"
            TxtValC.SetFocus
            Valida = False
            Exit Function
        End If
    End If
    Valida = True
End Function

Function DevCodAux(fila As Integer, columna As Integer) As String
Dim k As Integer
    For k = 0 To LstSede.ListCount - 1
        With MshEmpO
            .Col = columna: .row = fila
            Msf.Col = 0: Msf.row = k
            If .CellForeColor = Msf.CellBackColor Then
                DevCodAux = LstSede.List(k, 1) & "-" & LstSede.List(k, 2)
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
    oReporte.Titulo = "REPORTE RESUMEN DE CONTROL DE REFRIGERIOS  DE " & dptFecIniR.Value & " HASTA " & dptFecFinR.Value
    oReporte.Reporte = "Rep_BonosResumen.rpt"
    oReporte.sp_Rep_Refrigerios CDbl(IIf(val(txtValorRefri) = 0, 0, txtValorRefri)), Format(dptFecIniR.Value, "yyyy/mm/dd"), Format(dptFecFinR.Value, "yyyy/mm/dd"), cboMes.List(cboMes.ListIndex, 2), cboDiv.List(cboDiv.ListIndex, 1)
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
        For I = 2 To .Rows - 1
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

Private Sub cboMes_Change()
    If cboMes.ListIndex > 0 Then
        cboDiv.Enabled = True
        H = 1
        TabS_Click
    Else
        If cboDiv.ListCount > 0 Then cboDiv.ListIndex = 0
        cboDiv.Enabled = False
    End If
End Sub

Sub CalculoTotal()
Dim SumaTot As Double
    With MshEmpO
        For I = 2 To .Rows - 1
            SumaTot = SumaTot + CDbl(IIf(val(.TextMatrix(I, .Cols - 3)) = 0, 0, .TextMatrix(I, .Cols - 3)))
        Next
        lblTotal = "S/.  " & FormatNumber(SumaTot, 2)
    End With
End Sub

Sub TotalGeneral()
Dim SumaTot As Double
Dim SQL As String
Dim RQ As MYSQL_RS
    SQL = "select sum(montoa+montoc) as mto from rh_refriempleados where date_format(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        lblTotGen = "S/.  " & FormatNumber(RQ.Fields("mto"), 2)
    End If
Set RQ = Nothing
End Sub

Private Sub Form_Load()
    Call WheelHook(frmRegRefrigerios)
    CReaTabla = True
    Me.Left = 0
    Me.Top = 0
    SCol = -1
    dptFecIniR.Value = Date
    dptFecFinR.Value = Date
    cadenaemp = ""
    lblCadBus = ""
    lblTotGen = "S/. 0.00"
    lblTotal = "S/. 0.00"
    ConfiguraGrilla
    Divisiones cboDiv
    ArrColores
    Restaurantes
    LlenarMesP cboMes
    CargaMenu
End Sub

Sub CargaMenu()
    With Menu
        .addItem_Menu "Almuerzo", "A"
        .addItem_Menu "Cena", "C"
        .addItem_Menu "Ambos", "S"
        .addItem_Menu "Borrar", "B"
    End With
End Sub

Sub ConfiguraGrilla()
Dim fecha As String
    With MshEmpO
        fecha = "01/" & cboMes.List(cboMes.ListIndex, 2) & "/" & strAnoSistema
        If IsDate(fecha) Then
            .Rows = 3
            .Cols = 2
            .Clear
            .TextMatrix(1, 0) = "Item"
            .ColWidth(0) = 0 '330
            .ColWidth(1) = 2900
            .TextMatrix(1, 1) = Space(30) & "Empleado"
            dias = Day(DateSerial(Year(CDate(fecha)), Month(CDate(fecha)) + 1, 0))
            .Cols = .Cols + dias
            For I = 1 To dias
                .ColWidth(I + 1) = 250
                .TextMatrix(0, I + 1) = Left(Format(CDate(Right("00" & I, 2) & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & strAnoSistema), "dddd"), 1)
                .TextMatrix(1, I + 1) = CStr(I)
                If .TextMatrix(0, I + 1) = "D" Then
                    .row = 0: .Col = I + 1: .CellForeColor = vbRed
                    .row = 1: .Col = I + 1: .CellForeColor = vbRed
                Else
                    .row = 0: .Col = I + 1: .CellForeColor = vbBlack
                    .row = 1: .Col = I + 1: .CellForeColor = vbBlack
                End If
                .ColAlignment(I + 1) = flexAlignCenterCenter
            Next
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "factura"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "fechaingreso"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 270
            .TextMatrix(1, .Cols - 1) = "A"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 270
            .TextMatrix(1, .Cols - 1) = "C"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "Bono"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 670
            .TextMatrix(1, .Cols - 1) = "SubTotal"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "codigo"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "ccHFM"
            .CellForeColor = &H80FFFF
        End If
    End With
End Sub

Sub ConfiguraGrid()
    With Msf
        .Cols = 4
        .Rows = 0
        .ColWidth(0) = 230
        .ColWidth(1) = 0
        .ColWidth(2) = 230
        .ColWidth(3) = 230
    End With
End Sub

Sub Restaurantes()
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer
    SQ = "SELECT auxiliar,codigo, descrip from cnauxil where suspen = 'REFRI' " & _
         "UNION ALL select '0' as auxiliar,'00000000000' as codigo,'NINGUNO' AS DESCRIP order by codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    LstSede.Clear
    With Msf
        I = 0
        ConfiguraGrid
        If Not RQ.EOF() Then
            Do While Not RQ.EOF()
                LstSede.AddItem Trim(RQ.Fields("descrip"))
                LstSede.List(I, 1) = Trim(RQ.Fields("codigo"))
                LstSede.List(I, 2) = Trim(RQ.Fields("auxiliar"))
                .Rows = .Rows + 1
                .TextMatrix(I, 1) = I
                .Col = 0: .row = I
                .RowHeight(I) = 260
                .CellBackColor = ArrColor(I + 1)
                .Col = 2
                .CellFontName = "Wingdings"
                .CellFontSize = 9
                .ColAlignment(2) = flexAlignCenterBottom
                .TextMatrix(I, 2) = strUnChecked
                .Col = 3
                .CellFontName = "Wingdings"
                .CellFontSize = 9
                .ColAlignment(3) = flexAlignCenterBottom
                .TextMatrix(I, 3) = strUnChecked
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
Dim aux As String, codAux As String
Dim AnioIniV As String, AnioFinV As String
    lblCadBus = ""
    cadenaemp = ""
    MshEmpO.Redraw = False
    ConfiguraGrilla
    Screen.MousePointer = vbHourglass
    NomEmp = ""
    EntroBusqueda = False
    If Trim(TxtNom) <> "" Then
        EntroBusqueda = True
        NomEmp = " and e.apepat = '" & Trim(TxtNom) & "'"
    End If
    CrearTablaTemporal
    SQ = "insert into rh_tmpempasis(nombres,codigo,divi,situacion,modalidad) " & _
         "SELECT DISTINCT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo as cod, " & _
         "c.Division as divI,e.situacion,e.modalidad From empleado e LEFT OUTER JOIN (select fecha,codemp from rh_refriempleados) as t " & _
         "ON (e.codigo=t.codemp) and ifnull(DATE_FORMAT(ifnull(t.fecha,''),'%Y%m'),'') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "RIGHT OUTER JOIN (select if(ifnull(f_termino,'')<>'',f_termino,sysdate()) as f_termino,codemp,estado,division,fechacese from contrato) as c " & _
         "ON (e.codigo=c.codemp) and (if(c.fechacese='',ifnull(date_format(ifnull(c.f_termino,''),'%Y%m'),''),ifnull(date_format(ifnull(c.fechacese,''),'%Y%m'),'')) >= '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "or c.estado='AP') where (left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) >= '" & LI & "' and left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) <= '" & LF & "') " & NomEmp & " order by nombres"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    SQ = "SELECT DISTINCT e.nombres, IF(DAY(T.FECHA) = '01',ifnull(t.tipo,''),'') AS D1,IF(DAY(T.FECHA) = '02',ifnull(t.tipo,''),'') AS D2, " & _
         "IF(DAY(T.FECHA) = '03',ifnull(t.tipo,''),'') AS D3,IF(DAY(T.FECHA) = '04',ifnull(t.tipo,''),'') AS D4, " & _
         "IF(DAY(T.FECHA) = '05',ifnull(t.tipo,''),'') AS D5,IF(DAY(T.FECHA) = '06',ifnull(t.tipo,''),'') AS D6, " & _
         "IF(DAY(T.FECHA) = '07',ifnull(t.tipo,''),'') AS D7,IF(DAY(T.FECHA) = '08',ifnull(t.tipo,''),'') AS D8, " & _
         "IF(DAY(T.FECHA) = '09',ifnull(t.tipo,''),'') AS D9,IF(DAY(T.FECHA) = '10',ifnull(t.tipo,''),'') AS D10, " & _
         "IF(DAY(T.FECHA) = '11',ifnull(t.tipo,''),'') AS D11,IF(DAY(T.FECHA) = '12',ifnull(t.tipo,''),'') AS D12, " & _
         "IF(DAY(T.FECHA) = '13',ifnull(t.tipo,''),'') AS D13,IF(DAY(T.FECHA) = '14',ifnull(t.tipo,''),'') AS D14, " & _
         "IF(DAY(T.FECHA) = '15',ifnull(t.tipo,''),'') AS D15,IF(DAY(T.FECHA) = '16',ifnull(t.tipo,''),'') AS D16, " & _
         "IF(DAY(T.FECHA) = '17',ifnull(t.tipo,''),'') AS D17,IF(DAY(T.FECHA) = '18',ifnull(t.tipo,''),'') AS D18, " & _
         "IF(DAY(T.FECHA) = '19',ifnull(t.tipo,''),'') AS D19,IF(DAY(T.FECHA) = '20',ifnull(t.tipo,''),'') AS D20, " & _
         "IF(DAY(T.FECHA) = '21',ifnull(t.tipo,''),'') AS D21,IF(DAY(T.FECHA) = '22',ifnull(t.tipo,''),'') AS D22, " & _
         "IF(DAY(T.FECHA) = '23',ifnull(t.tipo,''),'') AS D23,IF(DAY(T.FECHA) = '24',ifnull(t.tipo,''),'') AS D24, " & _
         "IF(DAY(T.FECHA) = '25',ifnull(t.tipo,''),'') AS D25,IF(DAY(T.FECHA) = '26',ifnull(t.tipo,''),'') AS D26, " & _
         "IF(DAY(T.FECHA) = '27',ifnull(t.tipo,''),'') AS D27,IF(DAY(T.FECHA) = '28',ifnull(t.tipo,''),'') AS D28, " & _
         "IF(DAY(T.FECHA) = '29',ifnull(t.tipo,''),'') AS D29,IF(DAY(T.FECHA) = '30',ifnull(t.tipo,''),'') AS D30, " & _
         "IF(DAY(T.FECHA) = '31',ifnull(t.tipo,''),'') AS D31,e.item,e.codigo as cod,e.divI,ifnull(c.fec_Salida,'') as fec_salida,IFNULL(c.fec_Regreso,'') as fec_regreso,e.situacion,ifnull(tot3.subtotal,0) as subtotal, " & _
         "ifnull(T.FECHA,'') as fecha,ifnull(tot1.canta,0) as totala,ifnull(tot2.cantC,0) as totalc,ifnull(t.auxiliarA,'') as auxiliara,ifnull(t.codauxA,'') as codauxa,ifnull(t.auxiliarc,'') as auxiliarc,ifnull(t.codauxc,'') as codauxc,e.modalidad,COUNT(T.FECHA) AS totxfecha " & _
         "From rh_tmpempasis e LEFT OUTER JOIN (select fecha,codemp,tipo,auxiliarA,codauxA,facturaA,montoA,auxiliara as auxiliarc,codauxa as codauxc,facturaa as FACTURAC,montoa as MONTOC from rh_refriempleados) as t " & _
         "ON (e.codigo=t.codemp) and DATE_FORMAT(t.fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "left join (select codemp,fec_salida,fec_regreso from calendario where movemp = '02' and gocehaber = 'N') as c " & _
         "on (c.codemp=e.codigo) and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "LEFT OUTER JOIN (select count(*) as canta,codemp,fecha from rh_refriempleados where tipo IN ('A','S') and DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by codemp) as tot1 ON (e.codigo=tot1.codemp) " & _
         "LEFT OUTER JOIN (select count(*) as cantC,codemp,fecha from rh_refriempleados where tipo IN ('C','S') and DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by codemp) as tot2 ON (e.codigo=tot2.codemp) " & _
         "LEFT OUTER JOIN (select sum(montoa) as subtotal,codemp,fecha from rh_refriempleados where DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by codemp) as tot3 ON (e.codigo=tot3.codemp) Where e.codigo Is Not Null group by nombres,fecha,fec_salida"
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
                        .TextMatrix(I + 1, 0) = H
                        .TextMatrix(I + 1, 1) = Trim(RQ.Fields("nombres"))
                        .TextMatrix(I + 1, .Cols - 2) = RQ.Fields("cod")
                        .TextMatrix(I + 1, .Cols - 1) = RQ.Fields("divi")
                        .TextMatrix(I + 1, .Cols - 4) = FormatNumber(RQ.Fields("monto"), 2)
                        .TextMatrix(I + 1, .Cols - 5) = RQ.Fields("totalC")
                        .TextMatrix(I + 1, .Cols - 6) = RQ.Fields("totalA")
                        .TextMatrix(I + 1, .Cols - 3) = FormatNumber(RQ.Fields("subtotal"), 2)
                        If val(.TextMatrix(I + 1, .Cols - 3)) = "0.00" Then .TextMatrix(I + 1, .Cols - 3) = "0"
                        If val(.TextMatrix(I + 1, .Cols - 4)) = "0.00" Then .TextMatrix(I + 1, .Cols - 4) = "0"
                        .TextMatrix(I + 1, .Cols - 7) = RQ.Fields("modalidad") 'FecIngreso
                        If IsDate(RQ.Fields("fec_salida")) Then
                            DiaIniV = Day(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy"))
                            DiaFinV = Day(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            MesIniV = Month(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy"))
                            MesFinV = Month(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            AnioIniV = Year(Format(RQ.Fields("fec_salida"), "dd/mm/yyyy"))
                            AnioFinV = Year(Format(RQ.Fields("fec_regreso"), "dd/mm/yyyy"))
                            If AnioIniV < strAnoSistema Then DiaIniV = "01"
                            If MesIniV < val(cboMes.List(cboMes.ListIndex, 2)) Then DiaIniV = 1
                            If MesFinV > val(cboMes.List(cboMes.ListIndex, 2)) Then
                                DiaFinV = dias
                            Else
                                If AnioIniV = strAnoSistema And AnioFinV <> strAnoSistema Then
                                    DiaFinV = IIf(MesIniV = "12", "31", Day(DateSerial(CInt(AnioIniV), CInt(MesIniV) + 1, 0)))
                                End If
                            End If
                            For T = DiaIniV To DiaFinV
                                .TextMatrix(I + 1, T + 1) = "V"
                            Next
                        End If
                        .Col = 1: .row = I + 1: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        .Col = 0: .row = I + 1: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        If RQ.Fields("situacion") = 0 Then
                            If Format(FecTermino, "yyyymm") = strAnoSistema & cboMes.List(cboMes.ListIndex, 2) Then
                                .Col = 0: .row = I + 1: .CellForeColor = vbRed
                                .Col = 1: .row = I + 1: .CellForeColor = vbRed
                                MarcarDias Day(FecTermino), .Cols - 9, I + 1
                            End If
                        End If
                        If Format(FecIngreso, "yyyymm") = strAnoSistema & cboMes.List(cboMes.ListIndex, 2) Then
                            MarcarDias 1, Day(FecIngreso), I + 1
                        End If
                        .Col = .Cols - 5: .row = I + 1: .CellForeColor = &H80FFFF
                        .Col = .Cols - 6: .row = I + 1: .CellForeColor = &H80FFFF
                        k = k + 1
                    End If
                    .row = I + 1
                    If Trim(RQ.Fields("codauxa")) <> "" Then
                        aux = Trim(RQ.Fields("auxiliara"))
                        codAux = Trim(RQ.Fields("codauxa"))
                    Else
                        If Trim(RQ.Fields("codauxc")) <> "" Then
                            aux = Trim(RQ.Fields("auxiliarc"))
                            codAux = Trim(RQ.Fields("codauxc"))
                        End If
                    End If
                    If IsDate(RQ.Fields("FECHA")) Then
                        .TextMatrix(I + 1, 2) = IIf(.TextMatrix(I + 1, 2) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D1")) <> "", "S", Trim(RQ.Fields("D1"))), .TextMatrix(I + 1, 2))
                        .Col = 2: If .TextMatrix(I + 1, 2) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 3) = IIf(.TextMatrix(I + 1, 3) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D2")) <> "", "S", Trim(RQ.Fields("D2"))), .TextMatrix(I + 1, 3))
                        .Col = 3: If .TextMatrix(I + 1, 3) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 4) = IIf(.TextMatrix(I + 1, 4) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D3")) <> "", "S", Trim(RQ.Fields("D3"))), .TextMatrix(I + 1, 4))
                        .Col = 4: If .TextMatrix(I + 1, 4) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 5) = IIf(.TextMatrix(I + 1, 5) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D4")) <> "", "S", Trim(RQ.Fields("D4"))), .TextMatrix(I + 1, 5))
                        .Col = 5: If .TextMatrix(I + 1, 5) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 6) = IIf(.TextMatrix(I + 1, 6) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D5")) <> "", "S", Trim(RQ.Fields("D5"))), .TextMatrix(I + 1, 6))
                        .Col = 6: If .TextMatrix(I + 1, 6) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 7) = IIf(.TextMatrix(I + 1, 7) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D6")) <> "", "S", Trim(RQ.Fields("D6"))), .TextMatrix(I + 1, 7))
                        .Col = 7: If .TextMatrix(I + 1, 7) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 8) = IIf(.TextMatrix(I + 1, 8) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D7")) <> "", "S", Trim(RQ.Fields("D7"))), .TextMatrix(I + 1, 8))
                        .Col = 8: If .TextMatrix(I + 1, 8) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 9) = IIf(.TextMatrix(I + 1, 9) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D8")) <> "", "S", Trim(RQ.Fields("D8"))), .TextMatrix(I + 1, 9))
                        .Col = 9: If .TextMatrix(I + 1, 9) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 10) = IIf(.TextMatrix(I + 1, 10) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D9")) <> "", "S", Trim(RQ.Fields("D9"))), .TextMatrix(I + 1, 10))
                        .Col = 10: If .TextMatrix(I + 1, 10) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 11) = IIf(.TextMatrix(I + 1, 11) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D10")) <> "", "S", Trim(RQ.Fields("D10"))), .TextMatrix(I + 1, 11))
                        .Col = 11: If .TextMatrix(I + 1, 11) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 12) = IIf(.TextMatrix(I + 1, 12) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D11")) <> "", "S", Trim(RQ.Fields("D11"))), .TextMatrix(I + 1, 12))
                        .Col = 12: If .TextMatrix(I + 1, 12) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 13) = IIf(.TextMatrix(I + 1, 13) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D12")) <> "", "S", Trim(RQ.Fields("D12"))), .TextMatrix(I + 1, 13))
                        .Col = 13: If .TextMatrix(I + 1, 13) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 14) = IIf(.TextMatrix(I + 1, 14) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D13")) <> "", "S", Trim(RQ.Fields("D13"))), .TextMatrix(I + 1, 14))
                        .Col = 14: If .TextMatrix(I + 1, 14) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 15) = IIf(.TextMatrix(I + 1, 15) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D14")) <> "", "S", Trim(RQ.Fields("D14"))), .TextMatrix(I + 1, 15))
                        .Col = 15: If .TextMatrix(I + 1, 15) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 16) = IIf(.TextMatrix(I + 1, 16) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D15")) <> "", "S", Trim(RQ.Fields("D15"))), .TextMatrix(I + 1, 16))
                        .Col = 16: If .TextMatrix(I + 1, 16) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 17) = IIf(.TextMatrix(I + 1, 17) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D16")) <> "", "S", Trim(RQ.Fields("D16"))), .TextMatrix(I + 1, 17))
                        .Col = 17: If .TextMatrix(I + 1, 17) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 18) = IIf(.TextMatrix(I + 1, 18) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D17")) <> "", "S", Trim(RQ.Fields("D17"))), .TextMatrix(I + 1, 18))
                        .Col = 18: If .TextMatrix(I + 1, 18) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 19) = IIf(.TextMatrix(I + 1, 19) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D18")) <> "", "S", Trim(RQ.Fields("D18"))), .TextMatrix(I + 1, 19))
                        .Col = 19: If .TextMatrix(I + 1, 19) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 20) = IIf(.TextMatrix(I + 1, 20) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D19")) <> "", "S", Trim(RQ.Fields("D19"))), .TextMatrix(I + 1, 20))
                        .Col = 20: If .TextMatrix(I + 1, 20) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 21) = IIf(.TextMatrix(I + 1, 21) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D20")) <> "", "S", Trim(RQ.Fields("D20"))), .TextMatrix(I + 1, 21))
                        .Col = 21: If .TextMatrix(I + 1, 21) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 22) = IIf(.TextMatrix(I + 1, 22) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D21")) <> "", "S", Trim(RQ.Fields("D21"))), .TextMatrix(I + 1, 22))
                        .Col = 22: If .TextMatrix(I + 1, 22) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 23) = IIf(.TextMatrix(I + 1, 23) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D22")) <> "", "S", Trim(RQ.Fields("D22"))), .TextMatrix(I + 1, 23))
                        .Col = 23: If .TextMatrix(I + 1, 23) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 24) = IIf(.TextMatrix(I + 1, 24) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D23")) <> "", "S", Trim(RQ.Fields("D23"))), .TextMatrix(I + 1, 24))
                        .Col = 24: If .TextMatrix(I + 1, 24) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 25) = IIf(.TextMatrix(I + 1, 25) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D24")) <> "", "S", Trim(RQ.Fields("D24"))), .TextMatrix(I + 1, 25))
                        .Col = 25: If .TextMatrix(I + 1, 25) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 26) = IIf(.TextMatrix(I + 1, 26) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D25")) <> "", "S", Trim(RQ.Fields("D25"))), .TextMatrix(I + 1, 26))
                        .Col = 26: If .TextMatrix(I + 1, 26) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 27) = IIf(.TextMatrix(I + 1, 27) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D26")) <> "", "S", Trim(RQ.Fields("D26"))), .TextMatrix(I + 1, 27))
                        .Col = 27: If .TextMatrix(I + 1, 27) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 28) = IIf(.TextMatrix(I + 1, 28) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D27")) <> "", "S", Trim(RQ.Fields("D27"))), .TextMatrix(I + 1, 28))
                        .Col = 28: If .TextMatrix(I + 1, 28) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 29) = IIf(.TextMatrix(I + 1, 29) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D28")) <> "", "S", Trim(RQ.Fields("D28"))), .TextMatrix(I + 1, 29))
                        .Col = 29: If .TextMatrix(I + 1, 29) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 30) = IIf(.TextMatrix(I + 1, 30) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D29")) <> "", "S", Trim(RQ.Fields("D29"))), .TextMatrix(I + 1, 30))
                        .Col = 30: If .TextMatrix(I + 1, 30) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 31) = IIf(.TextMatrix(I + 1, 31) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D30")) <> "", "S", Trim(RQ.Fields("D30"))), .TextMatrix(I + 1, 31))
                        .Col = 31: If .TextMatrix(I + 1, 31) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
                        .TextMatrix(I + 1, 32) = IIf(.TextMatrix(I + 1, 32) = "", IIf(Trim(RQ.Fields("totxfecha")) > 1 And Trim(RQ.Fields("D31")) <> "", "S", Trim(RQ.Fields("D31"))), .TextMatrix(I + 1, 32))
                        .Col = 32: If .TextMatrix(I + 1, 32) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(aux, codAux)
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
    CalculoTotal
    TotalGeneral
    Screen.MousePointer = vbDefault
    Set RQ = Nothing
End Sub

Sub MarcarDias(DiaIni As Integer, DiaFin As Integer, fila As Integer)
    With MshEmpO
        For J = DiaIni + 2 To DiaFin
            .Col = J: .row = fila: .CellBackColor = &H7E7B72
        Next
    End With
End Sub

Function DevColor(aux As String, codAux As String) As Variant
Dim k As Integer
    For k = 0 To LstSede.ListCount - 1
        Msf.Col = 0: Msf.row = k
        If LstSede.List(k, 1) = codAux And LstSede.List(k, 2) = aux Then
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

Private Sub REFRIGERIOS(Tipo As String)
Dim I As Integer, Flg As Boolean
Dim k As Integer, flgdias As Boolean
Dim SubA As Double, SubC As Double
Dim FlgNoDias As Boolean
    With MshEmpO
        .Redraw = False
        GridSel = 1
        Flg = False
        If .Rows > 1 And Trim(.TextMatrix(3, 1)) <> "" And (.ColSel > 1) And ColIniO > 0 Then
            k = FilaIniO
            If val(TempTrab(FilaIniO)) = 0 Then
                TempTrab(FilaIniO) = FilaIniO
                CantTrab = CantTrab + 1
                Trab(CantTrab) = FilaIniO
            End If
            flgdias = False
            If Valida(Tipo) Then
                If Validadias(k, ColIniO, .ColSel) = False Then
                    If MsgBox("Existen algunos dias que el trabajador no tiene registrada asistencia." & Chr(13) & _
                              "¿Desea registrar el refrigerio de esos días?", vbQuestion + vbYesNo, "NOVPeru") = vbNo Then
                        flgdias = True
                    End If
                End If
                For I = ColIniO To .ColSel
                    If (LstSede.List(LstSede.ListIndex, 1) <> "00000000000") And (.TextMatrix(k, I) <> "V") Then
                        FlgNoDias = False
                        If flgdias = True Then
                            If ValidadiasAsistencia(k, I) = True Then
                                FlgNoDias = True
                            End If
                        Else
                            FlgNoDias = True
                        End If
                        If FlgNoDias = True Then
                            .row = k: .Col = I
                            If .CellBackColor <> &H7E7B72 Then
                                If Trim(.TextMatrix(k, I)) <> "" Then
                                    If Trim(.TextMatrix(k, I)) = "A" Or Trim(.TextMatrix(k, I)) = "S" Then
                                        .TextMatrix(k, .Cols - 6) = CDbl(IIf(val(.TextMatrix(k, .Cols - 6)) = 0, 0, .TextMatrix(k, .Cols - 6))) - 1
                                        SubA = DevTipoDia("A", Trim(.TextMatrix(k, .Cols - 2)), Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, I), 2)), "yyyy/mm/dd"))
                                    End If
                                    If Trim(.TextMatrix(k, I)) = "C" Or Trim(.TextMatrix(k, I)) = "S" Then
                                        .TextMatrix(k, .Cols - 5) = CDbl(IIf(val(.TextMatrix(k, .Cols - 5)) = 0, 0, .TextMatrix(k, .Cols - 5))) - 1
                                        SubC = DevTipoDia("C", Trim(.TextMatrix(k, .Cols - 2)), Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, I), 2)), "yyyy/mm/dd"))
                                    End If
                                    .TextMatrix(k, .Cols - 3) = CDbl(IIf(val(.TextMatrix(k, .Cols - 3)) = 0, 0, .TextMatrix(k, .Cols - 3))) - (SubA + SubC)
                                    SQL = "delete from rh_refriempleados where codemp = '" & .TextMatrix(k, .Cols - 2) & "' and " & _
                                          "fecha = '" & Format(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & .TextMatrix(1, I), "yyyy/mm/dd") & "'"
                                    oConexion.EjecutaSelectRS (SQL)
                                End If
                                .row = k: .Col = I
                                .CellForeColor = &HFFFFFF
                                If I <= 32 Then
                                    If LstSede.ListIndex > -1 Then
                                        Msf.row = LstSede.ListIndex
                                        Msf.Col = 0
                                        .CellForeColor = Msf.CellBackColor
                                        .TextMatrix(k, I) = Tipo
                                    Else
                                        .TextMatrix(k, I) = ""
                                    End If
                                    Flg = True
                                End If
                                .Col = I: .row = k
                                .CellFontSize = 8.5
                                .CellFontBold = True
                            End If
                        End If
                    End If
                Next
            End If
            If Flg = True Then
                GrabaRefrigerios ColIniO, .ColSel
                CalculoTotal
                TotalGeneral
            End If
            LimpiaArreglos
            CantTrab = 0
            FilaIniO = 0
            ColIniO = 0
            .Redraw = True
        End If
    End With
End Sub

Function Validadias(fila As Integer, ColIni As Integer, ColFin As Integer) As Boolean
Dim SQ As String
Dim RQ As MYSQL_RS
Dim fec As String, Numdia As Integer
    With MshEmpO
        SQ = "select * from rh_entsalempleado where emp = '" & .TextMatrix(fila, .Cols - 2) & "' and tipo = 'E' and " & _
             "fecha between '" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Trim(.TextMatrix(1, ColIni)), 2) & "' and " & _
             "'" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Trim(.TextMatrix(1, ColFin)), 2) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        Numdia = Trim(.TextMatrix(1, ColIni))
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

Function ValidadiasAsistencia(fila As Integer, Col As Integer) As Boolean
Dim SQ As String
Dim RQ As MYSQL_RS
    ValidadiasAsistencia = False
    With MshEmpO
        SQ = "select * from rh_entsalempleado where emp = '" & .TextMatrix(fila, .Cols - 2) & "' and tipo = 'E' and " & _
             "fecha = '" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & Trim(.TextMatrix(1, Col)), 2) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        If Not RQ.EOF() Then
            ValidadiasAsistencia = True
        End If
        Set RQ = Nothing
    End With
End Function

Private Sub BORRADO()
Dim I As Integer, Flg As Boolean
Dim k As Integer
    With MshEmpO
        GridSel = 1
        Flg = False
        If .Rows > 1 And Trim(.TextMatrix(3, 1)) <> "" And (.ColSel > 1) And ColIniO > 0 Then
            k = FilaIniO
            If val(TempTrab(FilaIniO)) = 0 Then
                TempTrab(FilaIniO) = FilaIniO
                CantTrab = CantTrab + 1
                Trab(CantTrab) = FilaIniO
            End If
            For I = ColIniO To .ColSel
                If (LstSede.List(LstSede.ListIndex, 1) <> "00000000000") And (.TextMatrix(k, I) <> "V") Then
                    .row = k: .Col = I
                    If .CellBackColor <> &H7E7B72 Then
                        If Trim(.TextMatrix(k, I)) <> "" Then
                            If I <= 32 Then
                                .TextMatrix(k, I) = ""
                            End If
                            Flg = True
                        End If
                    End If
                End If
            Next
            If Flg = True Then
                GrabaRefrigerios ColIniO, .ColSel
                CalculoTotal
                TotalGeneral
            End If
            LimpiaArreglos
            CantTrab = 0
            FilaIniO = 0
            ColIniO = 0
            .Redraw = True
        End If
    End With
End Sub

Private Sub Menu_Click()
    Select Case Menu.getSelectedItem_Index
        Case 1: REFRIGERIOS ("A") 'ALMUERZOS
        Case 2: REFRIGERIOS ("C") 'CENAS
        Case 3: REFRIGERIOS ("S") 'AMBOS
        Case 4: BORRADO
    End Select
End Sub

Private Sub Msf_Click()
    With Msf
        If .row > 0 Then
            If .Col = 2 Or .Col = 3 Then
                If Trim(.TextMatrix(.row, .Col)) = strUnChecked Then
                    If Not OtrosChecks(.Col) Then
                        .TextMatrix(.row, .Col) = strChecked
                        If .Col = 2 Then Chkalmuerzo.Value = 1
                        If .Col = 3 Then ChkCena.Value = 1
                    End If
                Else
                    .TextMatrix(.row, .Col) = strUnChecked
                    If .Col = 2 Then Chkalmuerzo.Value = 0
                    If .Col = 3 Then ChkCena.Value = 0
                End If
            End If
        End If
    End With
End Sub

Function OtrosChecks(columna As Integer) As Boolean
    OtrosChecks = False
    With Msf
        For I = 0 To .Rows - 1
            If .TextMatrix(I, columna) = strChecked Then
                OtrosChecks = True
                Exit For
            End If
        Next
    End With
End Function

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
                .row = .Rowsel: .Col = .ColSel
                If .CellForeColor <> &HFFFFFF Then
                    If .TextMatrix(.row, .Col) <> "" Then
                        MarcaSede .row, .Col
                    End If
                End If
                .Redraw = True
            End If
        Else
            LstSede.ListIndex = 0
        End If
    End With
End Sub

Sub MarcaSede(fila As Integer, columna As Integer)
Dim k As Integer
    For k = 0 To Msf.Rows - 1
        With MshEmpO
            .Col = columna: .row = fila
            Msf.Col = 0: Msf.row = k
            If .CellForeColor = Msf.CellBackColor Then
                LstSede.ListIndex = k
                Exit For
            End If
        End With
    Next
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
            Menu.Show X, Y
        End If
        .Redraw = True
    End With
End Sub

Sub LimpiaArreglos()
Dim I As Integer
    For I = 1 To CantTrab
        TempTrab(Trab(I)) = 0
        Trab(I) = 0
    Next
End Sub

Private Sub MshEmpO_RowColChange()
    DesplazarporGrid MshEmpO.Rowsel
End Sub

Sub DesplazarporGrid(fila As Integer)
    With MshEmpO
        lblCodEmp = Trim(.TextMatrix(fila, .Cols - 2))
        lbldiv = DescripcionesdeCodigos("DES_DIVISION", .TextMatrix(fila, .Cols - 1), "Descrip")
    End With
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
        If cboDiv.ListIndex > -1 Then
            cbodiv_Change
        End If
        If Trim(TxtImp) <> "" Then
            TxtImp_KeyPress 13
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
    SQL = "CREate table rh_tmpempasis( Item INT NOT NULL AUTO_INCREMENT, nombres char(100),codigo char(11),divi char(4), " & _
          "situacion char(1),modalidad char(2),PRIMARY KEY  (ITEM),UNIQUE KEY ITEM (ITEM))"
    oConexion.EjecutaSelectRS (SQL)
    CReaTabla = True
Exit Sub
CtrlError:
    CReaTabla = False
    MsgBox err.Description, "Error Creando Tabla temporal", "NOVPeru"
End Sub

Private Sub TxtFactA_GotFocus()
    mark TxtFactA
End Sub

Private Sub TxtFactA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtValA.SetFocus
End Sub

Private Sub TxtFactC_GotFocus()
    mark TxtFactC
End Sub

Private Sub TxtFactC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtValC.SetFocus
End Sub

Private Sub TxtImp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case TabS.SelectedItem.Index
            Case 1: FiltradoTxt "A", "B", Trim(TxtImp)
            Case 2: FiltradoTxt "C", "D", Trim(TxtImp)
            Case 3: FiltradoTxt "E", "F", Trim(TxtImp)
            Case 4: FiltradoTxt "G", "H", Trim(TxtImp)
            Case 5: FiltradoTxt "I", "J", Trim(TxtImp)
            Case 6: FiltradoTxt "K", "L", Trim(TxtImp)
            Case 7: FiltradoTxt "M", "Ñ", Trim(TxtImp)
            Case 8: FiltradoTxt "O", "P", Trim(TxtImp)
            Case 9: FiltradoTxt "Q", "R", Trim(TxtImp)
            Case 10: FiltradoTxt "S", "T", Trim(TxtImp)
            Case 11: FiltradoTxt "U", "V", Trim(TxtImp)
            Case 12: FiltradoTxt "W", "X", Trim(TxtImp)
            Case 13: FiltradoTxt "Y", "Z", Trim(TxtImp)
        End Select
    End If
End Sub

Private Sub FiltradoTxt(LetraIni As String, LetraFin As String, valor As String)
    Dim I As Integer
    Dim Col As Integer
    With MshEmpO
        For I = 2 To .Rows - 1
            Col = 1
            .RowHeight(I) = 245
            If valor = "" Or val(valor) = 0 Then
                If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) Then
                Else
                    .RowHeight(I) = 0
                End If
            Else
                If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) And (Trim(.TextMatrix(I, .Cols - 4)) = valor) Then
                Else
                    .RowHeight(I) = 0
                End If
            End If
        Next
    End With
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabS_Click
    End If
End Sub

Private Sub TxtValA_GotFocus()
    mark TxtValA
End Sub

Private Sub TxtValA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtFactA.SetFocus
End Sub

Private Sub TxtValC_GotFocus()
    mark TxtValC
End Sub

Private Sub TxtValC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TxtFactC.SetFocus
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
