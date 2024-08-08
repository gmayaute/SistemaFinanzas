VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepPlaniAFP 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Descuentos para el sistema privado de pensiones por AFP"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "frmRepPlaniAFP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H009F5539&
      Caption         =   "4.- Información sobre el Pago"
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
      Height          =   4035
      Left            =   3930
      TabIndex        =   19
      Top             =   720
      Width           =   4485
      Begin VB.Frame Frame5 
         BackColor       =   &H009F5539&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1890
         TabIndex        =   36
         Top             =   600
         Width           =   2505
         Begin VB.OptionButton OptFPagoF 
            BackColor       =   &H009F5539&
            Caption         =   "Efectivo"
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
            Index           =   0
            Left            =   60
            TabIndex        =   38
            Top             =   120
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton OptFPagoF 
            BackColor       =   &H009F5539&
            Caption         =   "Cheque"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.TextBox txtNroChequeA 
         Height          =   315
         Left            =   1920
         TabIndex        =   30
         Top             =   3120
         Width           =   2355
      End
      Begin VB.OptionButton OptFPagoA 
         BackColor       =   &H009F5539&
         Caption         =   "Efectivo"
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
         Index           =   3
         Left            =   1950
         TabIndex        =   29
         Top             =   2700
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton OptFPagoA 
         BackColor       =   &H009F5539&
         Caption         =   "Cheque"
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
         Index           =   2
         Left            =   3150
         TabIndex        =   28
         Top             =   2730
         Width           =   1095
      End
      Begin VB.TextBox txtNroChequeF 
         Height          =   315
         Left            =   1950
         TabIndex        =   24
         Top             =   1200
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO A LA AFP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Left            =   1620
         TabIndex        =   40
         Top             =   2280
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO AL FONDO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   225
         Left            =   1620
         TabIndex        =   39
         Top             =   330
         Width           =   2055
      End
      Begin MSForms.ComboBox cboBancoA 
         Height          =   315
         Left            =   1920
         TabIndex        =   31
         Top             =   3570
         Width           =   2355
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "4154;556"
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
         Caption         =   "Forma de Pago"
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
         Index           =   9
         Left            =   180
         TabIndex        =   27
         Top             =   2640
         Width           =   1635
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Index           =   8
         Left            =   180
         TabIndex        =   26
         Top             =   3600
         Width           =   1635
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. de Cheque"
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
         Index           =   7
         Left            =   180
         TabIndex        =   25
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma de Pago"
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
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   690
         Width           =   1635
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   1635
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. de Cheque"
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
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1665
      End
      Begin MSForms.ComboBox cboBancoF 
         Height          =   315
         Left            =   1950
         TabIndex        =   20
         Top             =   1680
         Width           =   2355
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "4154;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Caption         =   "3.- Tipo de Pago"
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
      Height          =   1785
      Left            =   150
      TabIndex        =   14
      Top             =   3540
      Width           =   3705
      Begin VB.OptionButton optTPago 
         BackColor       =   &H009F5539&
         Caption         =   "Extemporáneo"
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
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   1380
         Width           =   2745
      End
      Begin VB.OptionButton optTPago 
         BackColor       =   &H009F5539&
         Caption         =   "Normal"
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   330
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton optTPago 
         BackColor       =   &H009F5539&
         Caption         =   "Regularización de N° de Planilla"
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
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   660
         Width           =   3285
      End
      Begin VB.OptionButton optTPago 
         BackColor       =   &H009F5539&
         Caption         =   "Liquidación de cobranza N°"
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
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   1020
         Width           =   2745
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Caption         =   "2.- Método de Pago"
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
      Height          =   1425
      Left            =   150
      TabIndex        =   10
      Top             =   2010
      Width           =   3705
      Begin VB.OptionButton optPago 
         BackColor       =   &H009F5539&
         Caption         =   "Declaración sin pago"
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
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   1020
         Width           =   2745
      End
      Begin VB.OptionButton optPago 
         BackColor       =   &H009F5539&
         Caption         =   "Pago medio magnético"
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
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   660
         Width           =   2355
      End
      Begin VB.OptionButton optPago 
         BackColor       =   &H009F5539&
         Caption         =   "Pago con listado"
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   330
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "1.- Información General"
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
      Height          =   1815
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3705
      Begin VB.OptionButton OptSCTR 
         BackColor       =   &H009F5539&
         Caption         =   "No"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2460
         TabIndex        =   8
         Top             =   1110
         Width           =   885
      End
      Begin VB.OptionButton OptSCTR 
         BackColor       =   &H009F5539&
         Caption         =   "Si"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1470
         TabIndex        =   7
         Top             =   1110
         Width           =   885
      End
      Begin MSMask.MaskEdBox DtpFecha 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Fecha_Pago"
         Top             =   720
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   128
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   315
         Left            =   1410
         TabIndex        =   9
         Top             =   1410
         Width           =   2115
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "3731;556"
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
         Caption         =   "Fecha Pago"
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S.C.T.R"
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1125
      End
      Begin MSForms.ComboBox cboProceso 
         Height          =   315
         Left            =   1410
         TabIndex        =   3
         Top             =   300
         Width           =   2055
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "3625;556"
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
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AFP"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1125
      End
   End
   Begin Proyecto1.chameleonButton chBtnSalir 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   32
      ToolTipText     =   "Salir"
      Top             =   4830
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
      MICON           =   "frmRepPlaniAFP.frx":014A
      PICN            =   "frmRepPlaniAFP.frx":0166
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
      Left            =   5580
      TabIndex        =   33
      ToolTipText     =   "Ver Reporte"
      Top             =   4830
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
      MICON           =   "frmRepPlaniAFP.frx":052C
      PICN            =   "frmRepPlaniAFP.frx":0548
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox ComboBox2 
      Height          =   315
      Left            =   4800
      TabIndex        =   35
      Top             =   270
      Width           =   3615
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "6376;556"
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
      Caption         =   "Usuario"
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
      Index           =   10
      Left            =   3930
      TabIndex        =   34
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "frmRepPlaniAFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
