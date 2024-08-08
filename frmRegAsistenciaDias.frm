VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRegAsistenciaDias 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Entrada y Salida de Personal"
   ClientHeight    =   8130
   ClientLeft      =   2865
   ClientTop       =   5070
   ClientWidth     =   15120
   Icon            =   "frmRegAsistenciaDias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8130
   ScaleWidth      =   15120
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshEmpO 
      Height          =   6390
      Left            =   0
      TabIndex        =   27
      Top             =   780
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   11271
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
   Begin MSComctlLib.TabStrip TabS 
      Height          =   6720
      Left            =   0
      TabIndex        =   26
      Top             =   795
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   11853
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
   Begin VB.Frame Frame2 
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
      Left            =   7845
      TabIndex        =   18
      Top             =   -75
      Width           =   3930
      Begin VB.TextBox TxtNom 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   945
         TabIndex        =   20
         Top             =   150
         Width           =   2895
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
         Left            =   60
         TabIndex        =   19
         Top             =   180
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Msf 
      Height          =   7005
      Left            =   14850
      TabIndex        =   17
      Top             =   495
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   12356
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      GridColor       =   14737632
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.CheckBox ChkD 
      BackColor       =   &H009F5539&
      Caption         =   "Domingos"
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
      Left            =   13350
      TabIndex        =   13
      Top             =   135
      Value           =   1  'Checked
      Width           =   1200
   End
   Begin VB.CheckBox ChkS 
      BackColor       =   &H009F5539&
      Caption         =   "Sábados"
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
      Left            =   12105
      TabIndex        =   12
      Top             =   135
      Value           =   1  'Checked
      Width           =   1125
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
      Left            =   -15
      TabIndex        =   7
      Top             =   -75
      Width           =   7815
      Begin MSMask.MaskEdBox DtpFecha 
         Height          =   285
         Left            =   6270
         TabIndex        =   8
         ToolTipText     =   "Fecha_Pago"
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   128
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSForms.ComboBox CboDiv 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   150
         Width           =   2535
         VariousPropertyBits=   746604569
         DisplayStyle    =   7
         Size            =   "4471;556"
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
         Left            =   2220
         TabIndex        =   14
         Top             =   180
         Width           =   765
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         Left            =   5580
         TabIndex        =   9
         Top             =   165
         Width           =   675
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
         TabIndex        =   0
         Top             =   180
         Width           =   420
      End
      Begin MSForms.ComboBox CboMes 
         Height          =   315
         Left            =   495
         TabIndex        =   1
         Top             =   150
         Width           =   1650
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2910;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin Proyecto1.chameleonButton btnReporte 
      Height          =   345
      Left            =   14055
      TabIndex        =   3
      Top             =   7635
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
      MICON           =   "frmRegAsistenciaDias.frx":030A
      PICN            =   "frmRegAsistenciaDias.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   14550
      TabIndex        =   4
      Top             =   7635
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
      MICON           =   "frmRegAsistenciaDias.frx":0868
      PICN            =   "frmRegAsistenciaDias.frx":0884
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
      Left            =   12420
      TabIndex        =   5
      ToolTipText     =   "Guardar"
      Top             =   7995
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmRegAsistenciaDias.frx":0C4A
      PICN            =   "frmRegAsistenciaDias.frx":0C66
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
      Left            =   11115
      TabIndex        =   6
      ToolTipText     =   "Eliminar"
      Top             =   7995
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
      MICON           =   "frmRegAsistenciaDias.frx":10A8
      PICN            =   "frmRegAsistenciaDias.frx":10C4
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   5
      Left            =   6960
      TabIndex        =   31
      Top             =   7620
      Width           =   495
   End
   Begin MSForms.ComboBox CboPozo 
      Height          =   315
      Left            =   7485
      TabIndex        =   30
      Top             =   7605
      Width           =   3075
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "5424;556"
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   29
      Top             =   7620
      Width           =   495
   End
   Begin MSForms.ComboBox CboLote 
      Height          =   315
      Left            =   5085
      TabIndex        =   28
      Top             =   7605
      Width           =   1680
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "2963;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   825
      TabIndex        =   25
      Top             =   480
      Width           =   75
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
      TabIndex        =   24
      Top             =   480
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
      Left            =   60
      TabIndex        =   23
      Top             =   480
      Width           =   720
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
      Left            =   9840
      TabIndex        =   22
      Top             =   465
      Width           =   2760
   End
   Begin MSForms.Label lblo 
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   420
      Width           =   12690
      ForeColor       =   128
      Caption         =   "Listado  de  Empleados"
      PicturePosition =   393216
      Size            =   "22384;635"
      BorderColor     =   -2147483639
      SpecialEffect   =   3
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Lbldato 
      BorderStyle     =   1  'Fixed Single
      DragMode        =   1  'Automatic
      Height          =   240
      Left            =   13305
      TabIndex        =   16
      Top             =   7650
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSForms.ListBox LstSede 
      Height          =   7050
      Left            =   12690
      TabIndex        =   15
      Top             =   465
      Width           =   2145
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "3784;12435"
      MatchEntry      =   0
      ListStyle       =   1
      FontHeight      =   165
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
      Left            =   45
      TabIndex        =   11
      Top             =   7875
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
      Left            =   30
      TabIndex        =   10
      Top             =   7575
      Width           =   3495
   End
End
Attribute VB_Name = "frmRegAsistenciaDias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ColIniO As Integer, FilaIniO As Integer, ColFin As Integer
Dim emp As String, OpMov As Boolean
Dim Datos As String, CantTrab As Integer
Public GridSel As Integer
Dim cadenaemp As String, dias As Integer
Dim Entro As Boolean, FilaSel As Integer
Private Trab(1 To 500) As Integer
Private TempTrab(1 To 500) As Integer
Dim EntroBusqueda As Boolean, H As Integer
Dim CReaTabla As Boolean

Private Sub btnGrabar_Click()
    If GrabaAsistencia Then
    Else
        MsgBox "No se pudo grabar la información para el empleado:" & emp & vbNewLine & _
               "revise la asistencia o consulte con el administrador del sistema", vbOKOnly + vbExclamation, "NOVPeru"
    End If
End Sub

Function GrabaAsistencia() As Boolean
On Error GoTo CtrlError
    Dim I As Integer, J As Integer, k As Integer
    Dim SQ As String, sede As String
    Dim TotO As Integer, TotC As Integer
    GrabaAsistencia = False
    Screen.MousePointer = vbHourglass
    With MshEmpO
        For k = 1 To CantTrab
            I = Trab(k) ' - 1
            TotO = 0
            TotC = 0
            SQ = "delete from rh_entsalempleado where DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
                 "and emp = '" & .TextMatrix(I, .Cols - 2) & "' and envio = 'X'"
            oConexionMYSQL.Execute SQ
            For J = 2 To .Cols - 7
                If Trim(.TextMatrix(I, J)) = "C" Or Trim(.TextMatrix(I, J)) = "O" Then
                    .row = I: .Col = J
                    If .CellForeColor <> &HFFFFFF Then
                        sede = DevSede(I, J)
                        SQ = "call Insert_EntSal('" & sede & "','" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                             "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                             "'" & Format(IIf(Trim(.TextMatrix(I, J)) = "O", "08:30:00", "06:00:00"), "HH:MM:SS") & "','E','X','" & Trim(.TextMatrix(I, J)) & "')"
                        oConexionMYSQL.Execute SQ
                        SQ = "call Insert_EntSal('" & sede & "','" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                             "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                             "'" & Format(IIf(Trim(.TextMatrix(I, J)) = "O", "17:30:00", "18:00:00"), "HH:MM:SS") & "','S','X','" & Trim(.TextMatrix(I, J)) & "')"
                        oConexionMYSQL.Execute SQ
                        If Trim(.TextMatrix(I, .Cols - 1)) = "0001" Then
                            If J >= ColIniO And J <= ColFin Then
                                If VerificaRegBono(Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd"), Trim(.TextMatrix(I, .Cols - 2))) = False Then
                                    SQ = "delete from rh_bonosempleados where fecha = '" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "' " & _
                                         "and codemp = '" & .TextMatrix(I, .Cols - 2) & "'"
                                    oConexionMYSQL.Execute SQ
                                    SQ = "call Insert_Bonos('" & Trim(.TextMatrix(I, .Cols - 2)) & "'," & _
                                         "'" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "', " & _
                                         "'" & CboPozo.List(CboPozo.ListIndex, 1) & "','" & CboLote.List(CboLote.ListIndex, 1) & "','N'," & Trim(CDbl(IIf(val(.TextMatrix(I, .Cols - 6)) = 0, 0, .TextMatrix(I, .Cols - 6)))) & ")"
                                    oConexionMYSQL.Execute SQ
                                End If
                            End If
                        End If
                    End If
                    If Trim(.TextMatrix(I, J)) = "O" Then TotO = TotO + 1
                    If Trim(.TextMatrix(I, J)) = "C" Then TotC = TotC + 1
                Else
                    If Trim(.TextMatrix(I, .Cols - 1)) = "0001" Then
                        If J >= ColIniO And J <= ColFin Then
                            If VerificaRegBono(Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd"), Trim(.TextMatrix(I, .Cols - 2))) = False Then
                                SQ = "delete from rh_bonosempleados where fecha = '" & Format(CDate(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, J), 2)), "yyyy/mm/dd") & "' " & _
                                     "and codemp = '" & .TextMatrix(I, .Cols - 2) & "'"
                                oConexionMYSQL.Execute SQ
                            End If
                        End If
                    End If
                End If
            Next
        Next
        .TextMatrix(I, .Cols - 4) = TotC
        .TextMatrix(I, .Cols - 5) = TotO
    End With
    GrabaAsistencia = True
    Screen.MousePointer = vbDefault
Exit Function
CtrlError:
    GrabaAsistencia = False
End Function

Function VerificaRegBono(fec As String, CodEmp As String) As Boolean
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer
    VerificaRegBono = False
    SQ = "SELECT * from rh_bonosempleados where fecha = '" & fec & "' and codemp = '" & CodEmp & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    If Not RQ.EOF() Then
        VerificaRegBono = True
    End If
    Set RQ = Nothing
End Function

Sub LimpiaArreglos()
Dim I As Integer
    For I = 1 To CantTrab
        TempTrab(Trab(I)) = 0
        Trab(I) = 0
    Next
End Sub

Function DevSede(fila As Integer, columna As Integer) As String
Dim k As Integer
    For k = 0 To LstSede.ListCount - 1
        With MshEmpO
            .Col = columna: .row = fila
            Msf.Col = 0: Msf.row = k
            If .CellForeColor = Msf.CellBackColor Then
                DevSede = LstSede.List(k, 1)
                Exit For
            End If
        End With
    Next
End Function

Function DevColor(sede As String) As Variant
Dim k As Integer
    For k = 0 To LstSede.ListCount - 1
        Msf.Col = 0: Msf.row = k
        If LstSede.List(k, 1) = sede Then
            DevColor = Msf.CellBackColor
            Exit For
        End If
    Next
End Function

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

Private Sub btnReporte_Click()
    Screen.MousePointer = vbHourglass
    RegistraAsistencia
    Screen.MousePointer = vbDefault
    Set oReporte = New clsReporte
        oReporte.empresa = strNombreEmpresa
        oReporte.Titulo = "Registro de Asistencia de Personal - " & NombreMes(cboMes.List(cboMes.ListIndex, 2), False) & " " & strAnoSistema
        oReporte.Reporte = "Rep_Asistencia.rpt"
        oReporte.sp_Rep_Asistencias
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

Private Sub CboLote_Change()
    CargaPozos CboLote.List(CboLote.ListIndex, 1), CboPozo
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

Private Sub Form_Load()
    Call WheelHook(frmRegAsistenciaDias)
    CReaTabla = True
    Me.Left = 0
    Me.Top = 0
    SCol = -1
    DtpFecha.Text = Format(CStr(Date), "dd/mm/yyyy")
    cadenaemp = ""
    lblCadBus = ""
    ConfiguraGrilla
    Divisiones cboDiv
    ArrColores
    Sedes
    LlenarMesP cboMes
    CargaLotes CboLote
    CargaPozos CboLote.List(CboLote.ListIndex, 1), CboPozo
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
            .ColWidth(0) = 0
            .ColWidth(1) = 3630
            .TextMatrix(1, 1) = Space(37) & "Empleado"
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
            .TextMatrix(1, .Cols - 1) = "diaini"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 270
            .TextMatrix(1, .Cols - 1) = "O"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 270
            .TextMatrix(1, .Cols - 1) = "C"
            .Cols = .Cols + 1
            .ColWidth(.Cols - 1) = 0
            .TextMatrix(1, .Cols - 1) = "sedeorigen"
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

Sub CargaEmpleados(divis As String, Nombre As String, LI As String, LF As String)
Dim SQ As String
Dim RQ As MYSQL_RS, RQASIS As MYSQL_RS, RQ1 As MYSQL_RS
Dim I As Integer, FechaDia As String, k As Integer
Dim STRDiv As String
Dim DiaIniV As Integer, DiaFinV As Integer, T As Integer, MesIniV As Integer
Dim MesFinV As Integer, TotC As Integer, TotO As Integer, FlgInsert As Boolean
Dim NomEmp As String, FecIngreso As String, FecTermino As String
Dim AnioIniV As String, AnioFinV As String
    lblCadBus = ""
    cadenaemp = ""
    MshEmpO.Redraw = False
    ConfiguraGrilla
    Screen.MousePointer = vbHourglass
    STRDiv = ""
    NomEmp = ""
    EntroBusqueda = False
    If Trim(TxtNom) <> "" Then
        EntroBusqueda = True
        NomEmp = " and e.apepat = '" & Trim(TxtNom) & "'"
    End If
    CrearTablaTemporal
    SQ = "insert into rh_tmpempasis(nombres,codigo,divi,situacion,mon,bono) " & _
         "SELECT DISTINCT concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2) as nombres,e.codigo as cod, " & _
         "c.Division as divI,e.situacion,mon_bono,monto_bono From empleado e LEFT OUTER JOIN (select fecha,emp from rh_entsalempleado  where left(fecha,4)='" & strAnoSistema & "') as t " & _
         "ON (e.codigo=t.emp) and ifnull(DATE_FORMAT(ifnull(t.fecha,''),'%Y%m'),'') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "RIGHT OUTER JOIN (select if(ifnull(f_termino,'')<>'',f_termino,sysdate()) as f_termino,codemp,estado,mon_bono,monto_bono,division from contrato) as c " & _
         "ON (e.codigo=c.codemp) and (ifnull(date_format(ifnull(c.f_termino,''),'%Y%m'),'') >= '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "or c.estado='AP') where (ifnull(e.fec_cese,'') = '' or date_format(ifnull(e.fec_cese,''),'%Y%m') >= '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "') and (left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) >= '" & LI & "' and left(concat(e.apepat,' ',e.apemat,' ',e.nombre1,' ',e.nombre2),1) <= '" & LF & "') " & NomEmp & " group by e.codigo order by nombres"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    SQ = "SELECT DISTINCT e.nombres, IF(DAY(T.FECHA) = '01',ifnull(t.tiposede,''),'') AS D1,IF(DAY(T.FECHA) = '02',ifnull(t.tiposede,''),'') AS D2, " & _
         "IF(DAY(T.FECHA) = '03',ifnull(t.tiposede,''),'') AS D3,IF(DAY(T.FECHA) = '04',ifnull(t.tiposede,''),'') AS D4, " & _
         "IF(DAY(T.FECHA) = '05',ifnull(t.tiposede,''),'') AS D5,IF(DAY(T.FECHA) = '06',ifnull(t.tiposede,''),'') AS D6, " & _
         "IF(DAY(T.FECHA) = '07',ifnull(t.tiposede,''),'') AS D7,IF(DAY(T.FECHA) = '08',ifnull(t.tiposede,''),'') AS D8, " & _
         "IF(DAY(T.FECHA) = '09',ifnull(t.tiposede,''),'') AS D9,IF(DAY(T.FECHA) = '10',ifnull(t.tiposede,''),'') AS D10, " & _
         "IF(DAY(T.FECHA) = '11',ifnull(t.tiposede,''),'') AS D11,IF(DAY(T.FECHA) = '12',ifnull(t.tiposede,''),'') AS D12, " & _
         "IF(DAY(T.FECHA) = '13',ifnull(t.tiposede,''),'') AS D13,IF(DAY(T.FECHA) = '14',ifnull(t.tiposede,''),'') AS D14, " & _
         "IF(DAY(T.FECHA) = '15',ifnull(t.tiposede,''),'') AS D15,IF(DAY(T.FECHA) = '16',ifnull(t.tiposede,''),'') AS D16, " & _
         "IF(DAY(T.FECHA) = '17',ifnull(t.tiposede,''),'') AS D17,IF(DAY(T.FECHA) = '18',ifnull(t.tiposede,''),'') AS D18, " & _
         "IF(DAY(T.FECHA) = '19',ifnull(t.tiposede,''),'') AS D19,IF(DAY(T.FECHA) = '20',ifnull(t.tiposede,''),'') AS D20, " & _
         "IF(DAY(T.FECHA) = '21',ifnull(t.tiposede,''),'') AS D21,IF(DAY(T.FECHA) = '22',ifnull(t.tiposede,''),'') AS D22, " & _
         "IF(DAY(T.FECHA) = '23',ifnull(t.tiposede,''),'') AS D23,IF(DAY(T.FECHA) = '24',ifnull(t.tiposede,''),'') AS D24, " & _
         "IF(DAY(T.FECHA) = '25',ifnull(t.tiposede,''),'') AS D25,IF(DAY(T.FECHA) = '26',ifnull(t.tiposede,''),'') AS D26, " & _
         "IF(DAY(T.FECHA) = '27',ifnull(t.tiposede,''),'') AS D27,IF(DAY(T.FECHA) = '28',ifnull(t.tiposede,''),'') AS D28, " & _
         "IF(DAY(T.FECHA) = '29',ifnull(t.tiposede,''),'') AS D29,IF(DAY(T.FECHA) = '30',ifnull(t.tiposede,''),'') AS D30, " & _
         "IF(DAY(T.FECHA) = '31',ifnull(t.tiposede,''),'') AS D31,e.item,e.codigo as cod,e.divI,ifnull(c.fec_Salida,'') as fec_salida,IFNULL(c.fec_Regreso,'') as fec_regreso,e.situacion,ifnull(t.envio,'') as envio, " & _
         "ifnull(t.sede,'') as sede,ifnull(T.FECHA,'') as fecha,ifnull(tot1.canto,0) as totalO,ifnull(tot2.cantC,0) as totalc,e.bono " & _
         "From rh_tmpempasis e LEFT OUTER JOIN (select fecha,emp,tiposede,envio,sede from rh_entsalempleado where tipo='E') as t " & _
         "ON (e.codigo=t.emp) and DATE_FORMAT(t.fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' left join (select codemp,min(fec_salida) as fec_salida,max(fec_regreso) as fec_regreso from calendario where movemp = '02' and gocehaber = 'N' and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "and concat(left(fec_regreso,4),substring(fec_regreso,6,2))>='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by codemp) as c " & _
         "on (c.codemp=e.codigo) and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "and concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
         "LEFT OUTER JOIN (select count(*) as cantO,emp,fecha from rh_entsalempleado where tipo='E' and tiposede = 'O' and DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by emp) as tot1 ON (e.codigo=tot1.emp) " & _
         "LEFT OUTER JOIN (select count(*) as cantC,emp,fecha from rh_entsalempleado where tipo='E' and tiposede = 'C' and DATE_FORMAT(fecha,'%Y%m') = '" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' group by emp) as tot2 ON (e.codigo=tot2.emp) " & _
         "Where e.codigo Is Not Null group by nombres,fecha,fec_salida"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    I = 1: k = 0
    If Not RQ.EOF() Then
        Do While Not RQ.EOF
            With MshEmpO
                If RQ.Fields("ITEM") = I Then
                    If k = 0 Then
                        FecIngreso = FechaPersonal(RQ.Fields("cod"))
                        SQ = "SELECT FEC_CESE from empleado where codigo = '" & RQ.Fields("cod") & "'"
                        Set RQ1 = oConexion.EjecutaSelectRS(SQ)
                        If Not RQ1.EOF() Then
                            If RQ1.Fields("fec_cese") <> "" Then
                                FecTermino = RQ1.Fields("fec_cese")
                            Else
                                FecTermino = FechaPersonal(RQ.Fields("cod"), 2)
                            End If
                        End If
                        Set RQ1 = Nothing
                        .Rows = .Rows + 1
                        .TextMatrix(I + 2, 0) = H
                        .TextMatrix(I + 2, 1) = Trim(RQ.Fields("nombres"))
                        .TextMatrix(I + 2, .Cols - 2) = RQ.Fields("cod")
                        .TextMatrix(I + 2, .Cols - 1) = RQ.Fields("divi")
                        .TextMatrix(I + 2, .Cols - 4) = RQ.Fields("totalc")
                        .TextMatrix(I + 2, .Cols - 5) = RQ.Fields("totalo")
                        .TextMatrix(I + 2, .Cols - 6) = RQ.Fields("bono") 'FecIngreso
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
                                .TextMatrix(I + 2, T + 1) = "V"
                            Next
                            
                            Dim RQ_Vac As MYSQL_RS, contador As Integer
                            Dim DIni As Integer, DFin As Integer
                            contador = 0
                            SQ = "select codemp,fec_salida,fec_regreso from calendario where movemp = '02' and gocehaber = 'N' " & _
                                 "and concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
                                 "and concat(left(fec_regreso,4),substring(fec_regreso,6,2))>='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' " & _
                                 "and codemp = '" & RQ.Fields("cod") & "' order by fec_salida"
                            Set RQ_Vac = oConexion.EjecutaSelectRS(SQ)
                            Do While Not RQ_Vac.EOF()
                                If contador = 0 Then
                                    DIni = Day(RQ_Vac.Fields("fec_regreso"))
                                    contador = contador + 1
                                ElseIf contador > 0 Then
                                    DFin = Day(RQ_Vac.Fields("fec_salida"))
                                    contador = contador + 1
                                End If
                                RQ_Vac.MoveNext
                            Loop
                            Set RQ_Vac = Nothing
                            For T = DIni + 1 To DFin - 1
                                .TextMatrix(I + 2, T + 1) = ""
                            Next
                        End If
                        .Col = 1: .row = I + 2: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        .Col = 0: .row = I + 2: .CellFontBold = False: .CellFontSize = 7: .CellForeColor = &H80FFFF
                        If RQ.Fields("situacion") = 0 Then
                            If Format(FecTermino, "yyyymm") = strAnoSistema & cboMes.List(cboMes.ListIndex, 2) Then
                                .Col = 0: .row = I + 2: .CellForeColor = vbRed
                                .Col = 1: .row = I + 2: .CellForeColor = vbRed
                                MarcarDias Day(FecTermino), .Cols - 7, I + 2
                            End If
                        End If
                        If Format(FecIngreso, "yyyymm") = strAnoSistema & cboMes.List(cboMes.ListIndex, 2) Then
                            MarcarDias 1, Day(FecIngreso), I + 2
                        End If
                        .Col = .Cols - 4: .row = I + 2: .CellForeColor = &H80FFFF
                        .Col = .Cols - 5: .row = I + 2: .CellForeColor = &H80FFFF
                        k = k + 1
                    End If
                    .row = I + 2
                    If IsDate(RQ.Fields("FECHA")) Then
                        .TextMatrix(I + 2, 2) = IIf(.TextMatrix(I + 2, 2) = "", Trim(RQ.Fields("D1")), .TextMatrix(I + 2, 2))
                        .Col = 2: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 2) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 3) = IIf(.TextMatrix(I + 2, 3) = "", Trim(RQ.Fields("D2")), .TextMatrix(I + 2, 3))
                        .Col = 3: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 3) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 4) = IIf(.TextMatrix(I + 2, 4) = "", Trim(RQ.Fields("D3")), .TextMatrix(I + 2, 4))
                        .Col = 4: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 4) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 5) = IIf(.TextMatrix(I + 2, 5) = "", Trim(RQ.Fields("D4")), .TextMatrix(I + 2, 5))
                        .Col = 5: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 5) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 6) = IIf(.TextMatrix(I + 2, 6) = "", Trim(RQ.Fields("D5")), .TextMatrix(I + 2, 6))
                        .Col = 6: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 6) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 7) = IIf(.TextMatrix(I + 2, 7) = "", Trim(RQ.Fields("D6")), .TextMatrix(I + 2, 7))
                        .Col = 7: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 7) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 8) = IIf(.TextMatrix(I + 2, 8) = "", Trim(RQ.Fields("D7")), .TextMatrix(I + 2, 8))
                        .Col = 8: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 8) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 9) = IIf(.TextMatrix(I + 2, 9) = "", Trim(RQ.Fields("D8")), .TextMatrix(I + 2, 9))
                        .Col = 9: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 9) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 10) = IIf(.TextMatrix(I + 2, 10) = "", Trim(RQ.Fields("D9")), .TextMatrix(I + 2, 10))
                        .Col = 10: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 10) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 11) = IIf(.TextMatrix(I + 2, 11) = "", Trim(RQ.Fields("D10")), .TextMatrix(I + 2, 11))
                        .Col = 11: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 11) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 12) = IIf(.TextMatrix(I + 2, 12) = "", Trim(RQ.Fields("D11")), .TextMatrix(I + 2, 12))
                        .Col = 12: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 12) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 13) = IIf(.TextMatrix(I + 2, 13) = "", Trim(RQ.Fields("D12")), .TextMatrix(I + 2, 13))
                        .Col = 13: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 13) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 14) = IIf(.TextMatrix(I + 2, 14) = "", Trim(RQ.Fields("D13")), .TextMatrix(I + 2, 14))
                        .Col = 14: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 14) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 15) = IIf(.TextMatrix(I + 2, 15) = "", Trim(RQ.Fields("D14")), .TextMatrix(I + 2, 15))
                        .Col = 15: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 15) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 16) = IIf(.TextMatrix(I + 2, 16) = "", Trim(RQ.Fields("D15")), .TextMatrix(I + 2, 16))
                        .Col = 16: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 16) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 17) = IIf(.TextMatrix(I + 2, 17) = "", Trim(RQ.Fields("D16")), .TextMatrix(I + 2, 17))
                        .Col = 17: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 17) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 18) = IIf(.TextMatrix(I + 2, 18) = "", Trim(RQ.Fields("D17")), .TextMatrix(I + 2, 18))
                        .Col = 18: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 18) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 19) = IIf(.TextMatrix(I + 2, 19) = "", Trim(RQ.Fields("D18")), .TextMatrix(I + 2, 19))
                        .Col = 19: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 19) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 20) = IIf(.TextMatrix(I + 2, 20) = "", Trim(RQ.Fields("D19")), .TextMatrix(I + 2, 20))
                        .Col = 20: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 20) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 21) = IIf(.TextMatrix(I + 2, 21) = "", Trim(RQ.Fields("D20")), .TextMatrix(I + 2, 21))
                        .Col = 21: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 21) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 22) = IIf(.TextMatrix(I + 2, 22) = "", Trim(RQ.Fields("D21")), .TextMatrix(I + 2, 22))
                        .Col = 22: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 22) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 23) = IIf(.TextMatrix(I + 2, 23) = "", Trim(RQ.Fields("D22")), .TextMatrix(I + 2, 23))
                        .Col = 23: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 23) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 24) = IIf(.TextMatrix(I + 2, 24) = "", Trim(RQ.Fields("D23")), .TextMatrix(I + 2, 24))
                        .Col = 24: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 24) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 25) = IIf(.TextMatrix(I + 2, 25) = "", Trim(RQ.Fields("D24")), .TextMatrix(I + 2, 25))
                        .Col = 25: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 25) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 26) = IIf(.TextMatrix(I + 2, 26) = "", Trim(RQ.Fields("D25")), .TextMatrix(I + 2, 26))
                        .Col = 26: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 26) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 27) = IIf(.TextMatrix(I + 2, 27) = "", Trim(RQ.Fields("D26")), .TextMatrix(I + 2, 27))
                        .Col = 27: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 27) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 28) = IIf(.TextMatrix(I + 2, 28) = "", Trim(RQ.Fields("D27")), .TextMatrix(I + 2, 28))
                        .Col = 28: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 28) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 29) = IIf(.TextMatrix(I + 2, 29) = "", Trim(RQ.Fields("D28")), .TextMatrix(I + 2, 29))
                        .Col = 29: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 29) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 30) = IIf(.TextMatrix(I + 2, 30) = "", Trim(RQ.Fields("D29")), .TextMatrix(I + 2, 30))
                        .Col = 30: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 30) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 31) = IIf(.TextMatrix(I + 2, 31) = "", Trim(RQ.Fields("D30")), .TextMatrix(I + 2, 31))
                        .Col = 31: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 31) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                        .TextMatrix(I + 2, 32) = IIf(.TextMatrix(I + 2, 32) = "", Trim(RQ.Fields("D31")), .TextMatrix(I + 2, 32))
                        .Col = 32: If RQ.Fields("envio") = "X" And .TextMatrix(I + 2, 32) <> "V" And Trim(Day(RQ.Fields("FECHA"))) = .Col - 1 Then .CellForeColor = DevColor(RQ.Fields("sede")) 'Else .CellForeColor = &HFFFFFF
                    End If
                    Dim NumDias As Integer, DiaIniSal As Integer
                    Dim TipoSal As String
                    SQ = "Select c.fec_Salida,c.fec_Regreso,(select descrip from movi_emp m where m.codigo=c.movemp) as movemp from empleado as a left join calendario as c " & _
                         "on(c.codemp=a.codigo) where c.movemp in ('03','05','07') and " & _
                         "concat(left(fec_salida,4),substring(fec_salida,6,2))<='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' and " & _
                         "concat(left(c.fec_regreso,4),substring(c.fec_regreso,6,2))>='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' and " & _
                         "c.codemp = '" & RQ.Fields("cod") & "'"
                    Set RQ1 = oConexion.EjecutaSelectRS(SQ)
                    Dim ps As String
                    ps = strAnoSistema & Right("00" & Trim(str(cboMes.List(cboMes.ListIndex, 2))), 2)
                    Do While Not RQ1.EOF()
                        'NumDias = IIf(Replace(Left(RQ1.Fields("fec_regreso"), 7), "/", "") < per, Day(CDate(RQ1.Fields("fec_regreso"))), Day(DateSerial(strAnoSistema, val(cboMes.List(cboMes.ListIndex, 2)) + 1, 0))) - IIf(Replace(Left(RQ1.Fields("fec_regreso"), 7), "/", "") >= per, Day(CDate(RQ1.Fields("fec_salida"))), 1) + 1
                        'DiaIniSal = IIf(Replace(Left(RQ1.Fields("fec_salida"), 7), "/", "") >= per, Day(CDate(RQ1.Fields("fec_salida"))), Day(DateSerial(strAnoSistema, val(cboMes.List(cboMes.ListIndex, 2)) + 1, 1)))
                           
                        Dim PI As String
                        Dim PF As String
                        
                        PI = Left(Replace(RQ1.Fields("fec_salida"), "/", ""), 6)
                        PF = Left(Replace(RQ1.Fields("fec_regreso"), "/", ""), 6)
                           
                        DiaIniSal = 0
                        NumDias = 0
                        '--------------------------------------------------------
                        If PI < ps And PF = ps Then
                            DiaIniSal = 1
                            NumDias = val(Right(RQ1.Fields("fec_regreso"), 2))
                        End If
                        
                        If PI = ps And PF = ps Then
                            DiaIniSal = val(Right(RQ1.Fields("fec_salida"), 2))
                            NumDias = val(Right(RQ1.Fields("fec_regreso"), 2)) - DiaIniSal + 1
                        End If
                        
                        If PI = ps And PF > ps Then
                            DiaIniSal = val(Right(RQ1.Fields("fec_salida"), 2))
                            NumDias = DiasDelMes(Left(Replace(RQ1.Fields("fec_salida"), "/", ""), 6)) - DiaIniSal + 1
                        End If
                        
                        If (PI <> ps And PF <> ps) Then
                            DiaIniSal = 1
                            NumDias = DiasDelMes(ps)
                        End If
                        
                        '--------------------------------------------------------
                        If NumDias <> 0 Then
                            For k = 1 To NumDias
                                .TextMatrix(I + 2, DiaIniSal + k) = Mid(Trim(RQ1.Fields("movemp")), 1, 1)
                            Next
                        End If
                        '--------------------------------------------------------
                        RQ1.MoveNext
                    Loop
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

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("¿Seguro que desea salir del Registro de Asistencia?", vbQuestion + vbYesNo) = vbNo Then
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
        If .ColSel >= 2 And .ColSel <= .Cols - 7 Then
            If ColIniO = .ColSel Then
                .Redraw = False
                .row = .Rowsel: .Col = .ColSel
                If .CellForeColor <> &HFFFFFF Then
                    If .TextMatrix(.row, .Col) <> "" Then
                        MarcaSede .row, .Col
                    End If
                Else
                    If .TextMatrix(.row, .Col) <> "" Then
                        MostrarSede .row, .Col
                    End If
                End If
                .Redraw = True
            End If
        Else
            LstSede.ListIndex = 0
        End If
    End With
End Sub

Sub MostrarSede(fila As Integer, Colu As Integer)
    Dim RQ As MYSQL_RS, I As Integer
    Dim SQL As String, FLIn As Boolean
    With MshEmpO
        SQL = "select e.nombre,E.CODIGO from rh_entsalempleado r left join rh_estacionestrabajo e on (r.sede=e.codigo) " & _
              "where emp = '" & .TextMatrix(fila, .Cols - 2) & "' and fecha = '" & strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, Colu), 2) & "' and r.tipo = 'E'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            FLIn = False
            For I = 0 To LstSede.ListCount - 1
                If LstSede.List(I, 1) = Trim(RQ.Fields("CODIGO")) Then
                    FLIn = True
                    LstSede.ListIndex = I
                    Exit For
                End If
            Next
            If FLIn = False Then LstSede.ListIndex = 0
        Else
            LstSede.ListIndex = 0
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
            Dim I As Integer, Flg As Boolean
            Dim k As Integer
            GridSel = 1
            TotO = 0
            TotC = 0
            If .Rows > 1 And Trim(.TextMatrix(3, 1)) <> "" And (.ColSel > 1) And ColIniO > 0 Then
                k = FilaIniO
                If val(TempTrab(FilaIniO)) = 0 Then
                    TempTrab(FilaIniO) = FilaIniO
                    CantTrab = CantTrab + 1
                    Trab(CantTrab) = FilaIniO
                End If
                ColFin = .ColSel
                For I = ColIniO To .ColSel
                    If (LstSede.List(LstSede.ListIndex, 1) <> "00") And (.TextMatrix(k, I) <> "V") Then
                        .row = k: .Col = I
                        If .CellBackColor <> &H7E7B72 Then
                            If .TextMatrix(k, I) <> "" Then
                                If Not BuscaFirma(k, I, MshEmpO) And I <= 32 Then
                                    .TextMatrix(k, I) = ""
                                End If
                            Else
                                If Trim(.TextMatrix(k, .Cols - 1)) = "0001" Then
                                    If Not Valida Then
                                        If MsgBox("Si desea que se registre automáticamente los Bonos, debe seleccionar un Lote y un Pozo." & Chr(13) & _
                                                  "¿Desea registrar la Asistencia sin seleccionar Lote y/o Pozo?", vbQuestion + vbYesNo, "NOVPeru") = vbNo Then
                                            CboLote.SetFocus
                                            GoTo AQUI
                                        End If
                                    End If
                                End If
                                .row = k: .Col = I
                                .CellForeColor = &HFFFFFF
                                If ChkS.Value = 1 And ChkD.Value = 1 And I <= 32 Then
                                    If Left(Trim(.TextMatrix(k, I)), 1) <> "0" Then
                                        .row = k: .Col = I
                                        If LstSede.ListIndex > -1 Then
                                            Msf.row = LstSede.ListIndex
                                            Msf.Col = 0
                                            .CellForeColor = Msf.CellBackColor
                                            .TextMatrix(k, I) = IIf(LstSede.List(LstSede.ListIndex, 2) = "1", "O", "C")
                                        Else
                                            .TextMatrix(k, I) = "X"
                                        End If
                                    End If
                                Else
                                    If ChkS.Value = 0 And I <= 32 Then
                                        If IsDate(.TextMatrix(1, I) & "/" & cboMes.List(cboMes.ListIndex) & "/" & strAnoSistema) Then
                                            If Format(CDate(.TextMatrix(1, I) & "/" & cboMes.List(cboMes.ListIndex) & "/" & strAnoSistema), "dddd") = "Sábado" Then
                                                .TextMatrix(k, I) = ""
                                                Flg = True
                                            Else
                                                If LstSede.ListIndex > -1 Then
                                                    Msf.row = LstSede.ListIndex
                                                    Msf.Col = 0
                                                    .CellForeColor = Msf.CellBackColor
                                                    .TextMatrix(k, I) = IIf(LstSede.List(LstSede.ListIndex, 2) = "1", "O", "C")
                                                Else
                                                    .TextMatrix(k, I) = "X"
                                                End If
                                                Flg = False
                                            End If
                                        End If
                                    End If
                                    If ChkD.Value = 0 And Flg = False And I <= 32 Then
                                        If IsDate(.TextMatrix(1, I) & "/" & cboMes.List(cboMes.ListIndex) & "/" & strAnoSistema) Then
                                            If Format(.TextMatrix(1, I) & "/" & cboMes.List(cboMes.ListIndex) & "/" & strAnoSistema, "dddd") = "Domingo" Then
                                                .TextMatrix(k, I) = ""
                                            Else
                                                If LstSede.ListIndex > -1 Then
                                                    Msf.row = LstSede.ListIndex
                                                    Msf.Col = 0
                                                    .CellForeColor = Msf.CellBackColor
                                                    .TextMatrix(k, I) = IIf(LstSede.List(LstSede.ListIndex, 2) = "1", "O", "C")
                                                Else
                                                    .TextMatrix(k, I) = "X"
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            .Col = I: .row = k
                            .CellFontSize = 8.5
                            .CellFontBold = True
                        End If
                    End If
                Next
                btnGrabar_Click
AQUI:
                LimpiaArreglos
                CantTrab = 0
                FilaIniO = 0
                ColIniO = 0
                ColFin = 0
                .Redraw = True
            End If
        End If
        .Redraw = True
    End With
End Sub

Function Valida() As Boolean
    If CboLote.ListIndex <= 0 Then
        CboLote.SetFocus
        Valida = False
        Exit Function
    End If
    If CboPozo.ListIndex <= 0 Then
        CboPozo.SetFocus
        Valida = False
        Exit Function
    End If
    Valida = True
End Function

Private Sub MshEmpO_RowColChange()
    DesplazarporGrid MshEmpO.Rowsel
End Sub

Sub DesplazarporGrid(fila As Integer)
    With MshEmpO
        lblCodEmp = Trim(.TextMatrix(fila, .Cols - 2))
        lbldiv = DescripcionesdeCodigos("DES_DIVISION", .TextMatrix(fila, .Cols - 1), "Descrip")
    End With
End Sub

Sub ConfiguraGrid()
    With Msf
        .Cols = 2
        .Rows = 0
        .ColWidth(0) = 230
        .ColWidth(1) = 0
    End With
End Sub

Sub Sedes()
    Dim RQ As MYSQL_RS
    Dim SQ As String, I As Integer
    SQ = "SELECT * from rh_estacionestrabajo order by codigo"
    Set RQ = oConexion.EjecutaSelectRS(SQ)
    LstSede.Clear
    I = 0
    ConfiguraGrid
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            LstSede.AddItem Trim(RQ.Fields("nombre"))
            LstSede.List(I, 1) = Trim(RQ.Fields("codigo"))
            LstSede.List(I, 2) = Trim(RQ.Fields("tipo"))
            Msf.Rows = Msf.Rows + 1
            Msf.TextMatrix(I, 1) = I
            Msf.Col = 0: Msf.row = I
            Msf.RowHeight(I) = 260
            Msf.CellBackColor = ArrColor(I + 1)
            I = I + 1
            RQ.MoveNext
        Loop
        LstSede.ListIndex = 0
    End If
    Set RQ = Nothing
End Sub

Function ExtraeDatos(Dato As String) As String
Dim I As Integer
Dim cad As String
    For I = 1 To Len(Dato)
        If Mid(Dato, I, 1) = ";" Then
            cad = Mid(Dato, 1, I - 1)
            Exit For
        End If
    Next
    Dato = Mid(Dato, I + 1, Len(Dato))
    ExtraeDatos = cad
End Function

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

Sub RegistraAsistencia()
    Dim I As Integer, J As Integer
    Dim RQ As MYSQL_RS
    Dim SQ As String
    SQ = "delete from rh_tmpasistencias where usuario = '" & strUsuarioId & "'"
    oConexionMYSQL.Execute SQ
    With MshEmpO
        For I = 3 To .Rows - 1
            sede = DevSede(I, J)
            SQ = "select tipo from empleado where codigo = '" & Trim(.TextMatrix(I, .Cols - 2)) & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQ)
            If Not RQ.EOF() Then
                sede = Trim(RQ.Fields("tipo"))
            End If
            SQ = "insert into rh_tmpasistencias values('" & I - 2 & "','" & Trim(.TextMatrix(I, 1)) & "', " & _
                 "'" & Trim(.TextMatrix(I, .Cols - 2)) & "','" & IIf(sede = 3, "PRACTICANTE", "") & "','','','','','','','','','',''," & _
                 "'','','','','','','','','','','','','','','','','','','','','','','" & .TextMatrix(I, .Cols - 1) & "','','','" & strUsuarioId & "')"
            oConexionMYSQL.Execute SQ
            For J = 2 To .Cols - 7
                If .TextMatrix(I, J) <> "" Then
                    SQ = "update rh_tmpasistencias set D" & .TextMatrix(1, J) & "= '" & .TextMatrix(I, J) & "'" & " where codigo = '" & Trim(.TextMatrix(I, .Cols - 2)) & "' " & _
                         "and usuario = '" & strUsuarioId & "'"
                    oConexionMYSQL.Execute SQ
                End If
            Next
        Next
    End With
End Sub

Function BuscaFirma(fila As Integer, Colum As Integer, Msh As MSHFlexGrid) As Boolean
    Dim RQ As MYSQL_RS
    Dim SQ As String
    With Msh
        SQ = "select * from rh_entsalempleado where emp = '" & .TextMatrix(fila, .Cols - 2) & "' " & _
             "and fecha = '" & Format(strAnoSistema & "/" & cboMes.List(cboMes.ListIndex, 2) & "/" & Right("00" & .TextMatrix(1, Colum), 2), "yyyy/mm/dd") & "' " & _
             "and envio <> 'X'"
        Set RQ = oConexion.EjecutaSelectRS(SQ)
        If Not RQ.EOF() Then
            BuscaFirma = True
        End If
        Set RQ = Nothing
    End With
End Function

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
    End If
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabS_Click
    End If
End Sub

Private Sub Filtrado(LetraIni As String, LetraFin As String, Division As String)
    Dim I As Integer
    Dim CriterioIni As String, CriterioFin As String
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

Private Sub Busqueda(FiltroNom As String, Optional Division As String, Optional LetraIni As String, Optional LetraFin As String)
    Dim I As Integer
    Dim criterio As String
    Dim Col As Integer
    With MshEmpO
        criterio = "" & UCase(FiltroNom) & "*"
        For I = 2 To .Rows - 1
            Col = 1
            If criterio = "*" Then
                .RowHeight(I) = 245
                If Division = "0006" Or Division = "00" Then
                    If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) Then
                        .RowHeight(I) = 245
                    Else
                        .RowHeight(I) = 0
                    End If
                Else
                    If (UCase(Left(.TextMatrix(I, Col), 1)) >= LetraIni) And (UCase(Left(.TextMatrix(I, Col), 1)) <= LetraFin) And (.TextMatrix(I, .Cols - 1) = Division) Then
                    Else
                        .RowHeight(I) = 0
                    End If
                End If
            Else
                If UCase(.TextMatrix(I, Col)) Like criterio Then
                    .RowHeight(I) = 245
                Else
                    .RowHeight(I) = 0
                End If
            End If
        Next
    End With
End Sub

Sub CrearTablaTemporal()
On Error GoTo CtrlError
Dim SQL As String
    If CReaTabla = True Then
        SQL = "drop table rh_tmpempasis"
        oConexion.EjecutaSelectRS (SQL)
    End If
    SQL = "CREate table rh_tmpempasis( Item INT NOT NULL AUTO_INCREMENT, nombres char(100),codigo char(11),divi char(12), " & _
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
