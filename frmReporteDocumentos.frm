VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmReporteDocumentos 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Documentos"
   ClientHeight    =   8085
   ClientLeft      =   3750
   ClientTop       =   4935
   ClientWidth     =   16185
   Icon            =   "frmReporteDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   16185
   ShowInTaskbar   =   0   'False
   Begin Proyecto1.chameleonButton chBtnSalir 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9870
      TabIndex        =   19
      Top             =   7620
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
      BCOL            =   13160660
      BCOLO           =   15309923
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReporteDocumentos.frx":2AFA
      PICN            =   "frmReporteDocumentos.frx":2B16
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
      Left            =   9420
      TabIndex        =   20
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
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
      BCOL            =   13160660
      BCOLO           =   15309923
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReporteDocumentos.frx":2EDC
      PICN            =   "frmReporteDocumentos.frx":2EF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "Documento"
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
      Height          =   2670
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   13035
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   5220
         MaxLength       =   14
         TabIndex        =   44
         Top             =   1905
         Width           =   1335
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H009F5539&
         Caption         =   "Anual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   8610
         TabIndex        =   36
         Top             =   720
         Width           =   1035
      End
      Begin VB.OptionButton optMov 
         BackColor       =   &H009F5539&
         Caption         =   "Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   8580
         TabIndex        =   35
         Top             =   330
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox txtDiviIni 
         Height          =   285
         Left            =   60
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1905
         Width           =   1065
      End
      Begin VB.TextBox txtDiviFin 
         Height          =   285
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1905
         Width           =   1245
      End
      Begin MSMask.MaskEdBox meFolIni 
         Height          =   285
         Left            =   6630
         TabIndex        =   23
         Top             =   555
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox meImporte 
         Height          =   285
         Left            =   4680
         TabIndex        =   22
         Top             =   555
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meFecEini 
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   1215
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   " ##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCorrFin 
         Height          =   285
         Left            =   3120
         MaxLength       =   9
         TabIndex        =   8
         Top             =   555
         Width           =   1365
      End
      Begin VB.TextBox txtCorrIni 
         Height          =   285
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   7
         Top             =   555
         Width           =   1365
      End
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   60
         MaxLength       =   5
         TabIndex        =   6
         Top             =   555
         Width           =   1215
      End
      Begin MSMask.MaskEdBox meFecEfin 
         Height          =   315
         Left            =   1590
         TabIndex        =   14
         Top             =   1215
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   " ##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meFecRini 
         Height          =   315
         Left            =   3150
         TabIndex        =   15
         Top             =   1215
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   " ##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox meFecRfin 
         Height          =   315
         Left            =   5220
         TabIndex        =   16
         Top             =   1215
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   " ##/##/####"
         PromptChar      =   "_"
      End
      Begin Proyecto1.chameleonButton cmdBuscar 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   21
         Top             =   570
         Width           =   405
         _ExtentX        =   714
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReporteDocumentos.frx":343A
         PICN            =   "frmReporteDocumentos.frx":3456
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H009F5539&
         Caption         =   "Orden"
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
         Height          =   1485
         Left            =   6690
         TabIndex        =   24
         Top             =   975
         Width           =   3255
         Begin VB.Frame Frame4 
            BackColor       =   &H009F5539&
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   900
            TabIndex        =   40
            Top             =   150
            Width           =   1875
            Begin VB.OptionButton OptSentido 
               BackColor       =   &H009F5539&
               Caption         =   "Desc"
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
               Index           =   1
               Left            =   1020
               TabIndex        =   42
               Top             =   60
               Width           =   825
            End
            Begin VB.OptionButton OptSentido 
               BackColor       =   &H009F5539&
               Caption         =   "Asc"
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
               Index           =   0
               Left            =   30
               TabIndex        =   41
               Top             =   60
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   7
            Left            =   1710
            TabIndex        =   39
            Top             =   1200
            Width           =   1275
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Importe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   6
            Left            =   1710
            TabIndex        =   38
            Top             =   930
            Width           =   1275
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Fec. Emisión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   3
            Left            =   60
            TabIndex        =   37
            Top             =   1200
            Width           =   1485
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Auxiliar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   29
            Top             =   930
            Width           =   1275
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Fec. Vcto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   5
            Left            =   1710
            TabIndex        =   28
            Top             =   690
            Width           =   1275
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "T. Doc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   450
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Nro. Doc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   26
            Top             =   690
            Width           =   1275
         End
         Begin VB.OptionButton optOrden 
            BackColor       =   &H009F5539&
            Caption         =   "Fec. Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   225
            Index           =   4
            Left            =   1710
            TabIndex        =   25
            Top             =   450
            Width           =   1515
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° de Correlativo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   50
         Top             =   255
         Width           =   2985
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   60
         TabIndex        =   49
         Top             =   1605
         Width           =   2805
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auxiliares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3150
         TabIndex        =   48
         Top             =   1605
         Width           =   3405
      End
      Begin MSForms.ComboBox CboTipo 
         Height          =   315
         Left            =   1065
         TabIndex        =   47
         Top             =   2265
         Width           =   4110
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "7250;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Doc:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   46
         Top             =   2288
         Width           =   915
      End
      Begin MSForms.ComboBox cboAux 
         Height          =   285
         Left            =   3150
         TabIndex        =   45
         Top             =   1905
         Width           =   2055
         DisplayStyle    =   7
         Size            =   "3625;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H009F5539&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   1290
         TabIndex        =   32
         Top             =   1890
         Width           =   165
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H009F5539&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   4740
         TabIndex        =   18
         Top             =   1215
         Width           =   165
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H009F5539&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   1380
         TabIndex        =   17
         Top             =   1215
         Width           =   165
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3150
         TabIndex        =   13
         Top             =   930
         Width           =   3405
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Emisión:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   930
         Width           =   2805
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H009F5539&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   2910
         TabIndex        =   5
         Top             =   540
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H009F5539&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   360
         Left            =   1350
         TabIndex        =   4
         Top             =   510
         Width           =   120
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Folio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6630
         TabIndex        =   2
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° de Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Caption         =   "Detalle"
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
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   2670
      Width           =   13035
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexDetalle 
         Height          =   4605
         Left            =   60
         TabIndex        =   10
         Top             =   210
         Width           =   12825
         _ExtentX        =   22622
         _ExtentY        =   8123
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblMensaje 
      BackColor       =   &H80000007&
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
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   3600
      TabIndex        =   43
      Top             =   7650
      Width           =   5775
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T. Doc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   34
      Top             =   7665
      Width           =   915
   End
   Begin MSForms.ComboBox cmbTDocs 
      Height          =   345
      Left            =   990
      TabIndex        =   33
      Top             =   7650
      Width           =   2565
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4524;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmReporteDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oConsulta As FrmConsultas
Dim orden As String
Dim filtro As Boolean
Dim ContFiltros As Integer
Dim VSel(14) As Seleccion

Private Sub cboAux_Change()
    strTipoAuxiliar = cboAux.List(cboAux.ListIndex, 1)
    If cboAux.ListIndex > 0 Then
        If VSel(12).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(12).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(12).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub CboTipo_Change()
    If CboTipo.ListIndex > 0 Then
        If VSel(14).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(14).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(14).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub chBtnReporte_Click()
    Me.MousePointer = vbHourglass
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    If cmbTDocs.ListCount = 2 And UCase(cmbTDocs.List(1, 1)) = "O" And cmbTDocs.ListIndex = 0 And cmbTDocs.BackColor = ColorHabilitado Then
        If optMov(0) Then oReporte.Titulo = "REPORTE DE ORDEN DE COMPRA/SERVICIO DEL MES " & NombreMes(strMesSistema, False)
        If optMov(1) Then oReporte.Titulo = "REPORTE DE ORDEN DE COMPRA/SERVICIO DEL AÑO " & strAnoSistema
        oReporte.Reporte = "Rep_OrdenCompra.rpt"
        oReporte.sp_Rep_OrdenCompra Trim(txtCorrIni.Text), Trim(txtCorrFin.Text)
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    If cmbTDocs.ListCount > 0 And cmbTDocs.ListIndex = 0 And cmbTDocs.BackColor = ColorHabilitado Then
        If optMov(0) Then oReporte.Titulo = "REPORTE DE DOCUMENTOS VARIOS DEL MES " & NombreMes(strMesSistema, False)
        If IsDate(meFecRini.Text) And IsDate(meFecRfin) Then
            oReporte.Titulo = "REPORTE DE DOCUMENTOS VARIOS DEL " & meFecRini.Text & " AL " & meFecRfin
        Else
            If optMov(1) Then oReporte.Titulo = "REPORTE DE DOCUMENTOS VARIOS DEL AÑO " & strAnoSistema
        End If
        oReporte.Reporte = "Rep_Varios.rpt"
        oReporte.sp_Rep_Varios
    End If
    If cmbTDocs.ListCount > 0 And cmbTDocs.ListIndex > 0 And cmbTDocs.BackColor = ColorHabilitado Then
        If optMov(0) Then oReporte.Titulo = "REPORTE DE " & cmbTDocs.List(cmbTDocs.ListIndex, 0) & " DEL MES " & NombreMes(strMesSistema, False)
        If optMov(1) Then oReporte.Titulo = "REPORTE DE " & cmbTDocs.List(cmbTDocs.ListIndex, 0) & " DEL AÑO " & strAnoSistema
        Select Case cmbTDocs.List(cmbTDocs.ListIndex, 2)
            Case FAMILIA_DOC.CONTABLES
                If cmbTDocs.List(cmbTDocs.ListIndex, 1) = "01" Then
                    oReporte.Reporte = "Rep_Busqueda_Fac2.rpt"
                Else
                    oReporte.Reporte = "Rep_Contables2.rpt"
                End If
                oReporte.sp_Rep_Doc cmbTDocs.List(cmbTDocs.ListIndex, 1)
            Case FAMILIA_DOC.ORDENES
                oReporte.Reporte = "Rep_Ordenes2.rpt"
                oReporte.sp_Rep_Doc cmbTDocs.List(cmbTDocs.ListIndex, 1)
            Case FAMILIA_DOC.ENTIDADES
                oReporte.Reporte = "Rep_Contables2.rpt"
                oReporte.sp_Rep_Doc cmbTDocs.List(cmbTDocs.ListIndex, 1)
            Case FAMILIA_DOC.GENERALES
                oReporte.Reporte = "Rep_AdmPersonal2.rpt"
                oReporte.sp_Rep_Doc cmbTDocs.List(cmbTDocs.ListIndex, 1)
        End Select
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub chBtnSalir_Click()
    Unload Me
End Sub

Private Sub cmbTDocs_Change()
    Dim rsFiltro As MYSQL_RS
    Dim SQL As String
    If filtro = True Then
        If cmbTDocs.ListIndex > 0 And cmbTDocs.BackColor = ColorHabilitado Then
            SQL = "Select Identificador,tipo,documento,fec_registro,fec_emision,auxiliar," & _
                  " codigo,TRIM(cenco) AS CENCO,mon,total,desestado  from temprptdocumento where usu='" & strUsuarioId & "' and tipo='" & cmbTDocs.List(cmbTDocs.ListIndex, 1) & "'" & orden
        Else
            SQL = "Select Identificador,tipo,documento,fec_registro,fec_emision,auxiliar," & _
                  " codigo,TRIM(cenco) AS CENCO,mon,total,desestado  from temprptdocumento where usu='" & strUsuarioId & "'" & orden
        End If
            Llenargrilladetalle
            Set rsFiltro = oConexion.EjecutaSelectRS(SQL)
            LLenarDatos rsFiltro
    End If
    Set rsFiltro = Nothing
End Sub

Private Sub cmbTDocs_Click()
    filtro = True
End Sub

Private Sub cmdBuscar_Click()
    filtro = False
    Llenargrilladetalle
    Buscar
End Sub

Private Sub flexDetalle_RowColChange()
If flexDetalle.row > 0 And flexDetalle.Col > 0 Then DesplazarenFlex LblMensaje, flexDetalle
End Sub

Private Sub Form_Activate()
    ContFiltros = 1
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call WheelHook(frmReporteDocumentos)
    Set oDoc = New clsDocumento
    fitro = False
    Llenargrilladetalle
    LlenarAux
    LlenarTipoDocs
    LlenaArregloSel
    cmbTDocs.Enabled = False
    cmbTDocs.BackColor = ColorDeshabilitado
    Call optOrden_Click(0)
    Set oConsulta = New FrmConsultas
End Sub

Public Function Buscar()
    Dim rsdetalle As MYSQL_RS
    Dim SQL As String
    oConexion.EjecutaInsertUpdateDelete "Delete from temprptdocumento where usu='" & strUsuarioId & "'", TIPO_QUERY.Eliminar, False
    Dim I As Integer
    For I = 1 To 4
        Select Case I
            Case 1
                SQL = " Insert into temprptdocumento select a.Identificador,1 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " concat(doc.serie,'-',doc.correl) AS Documento,doc.orden,doc.guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                      " TRIM(doc.cenco) AS cenco,c.descrip as descenco,doc.mon AS mon,doc.dias_vcto,doc.fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,doc.total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,documento_contables AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo))left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea )"
            Case 2
                SQL = " Insert into temprptdocumento select a.Identificador,4 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " concat(doc.serie,'-',doc.correl) AS Documento,'' as orden,'' as guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                      " TRIM(doc.cenco) AS cenco,c.descrip as descenco,'' as mon,'' as dias_vcto,'' as fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,0 as total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,concat_ws('-',doc.obs,doc.asunto) as obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,documento_entidades AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea )"
            Case 3
                SQL = " Insert into temprptdocumento select a.Identificador,6 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " doc.correl AS Documento,'' as orden,'' AS guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                      " TRIM(doc.cenco) AS cenco,c.descrip as descenco,doc.mon,'' as dias_vcto,'' AS fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,doc.total,'' AS estado,m.Cod_Estado as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,doc.obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,orden_compra AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo)) left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea )"
            Case 4
                SQL = " Insert into temprptdocumento select a.Identificador,1 as familia,a.Cod_Tipo_Doc AS Tipo,cdoc.Descrip as Destipo," & _
                      " concat(doc.serie,'-',doc.correl) AS Documento,doc.orden,doc.guia," & _
                      " doc.auxiliar AS auxiliar, t.descrip as desaux,doc.codigo as Codigo,aux.descrip as descod," & _
                      " TRIM(doc.cenco) AS cenco,c.descrip as descenco,doc.mon AS mon,doc.dias_vcto,doc.fec_vcto," & _
                      " doc.fec_emision,a.fecha_registro as fec_registro,doc.total,m.Cod_Estado AS estado,e.descripcion as desestado," & _
                      " (SELECT U.AREA FROM 3cnuser AS U WHERE USUARIO_ID='" & strUsuarioId & "') AS area, m.prioridad AS prioridad, doc.division,m.usuario,'" & strUsuarioId & "' as usu,concat_ws('-',doc.obs,doc.asunto) as obs" & _
                      " from (((((amarre_documento as a, movi_documento as m ,documento_generales AS doc)" & _
                      " left join cndocum as cdoc on (a.Cod_Tipo_Doc=cdoc.Coddoc)) left join cnauxil as aux" & _
                      " on (doc.auxiliar=aux.auxiliar and doc.codigo=aux.codigo))left join cncosto as c on (doc.cenco=c.cenco))" & _
                      " left join doc_estado as e on (e.cod_estado=m.cod_estado)) left join cntablas as t" & _
                      " on (doc.auxiliar=t.tip_linea )"
        End Select
        SQL = SQL & " where T.codtab = 1 and (a.identificador=m.identificador) and (a.identificador=doc.identificador) " & _
                    " and (cdoc.protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=Cdoc.coddoc AND usuario = '" & strUsuarioId & "')=1)"
        If (Trim(meFolIni) <> Empty) And optMov(0) Then
           SQL = SQL & " and (a.Identificador='" & strAnoSistema & strMesSistema & meFolIni.Text & "')"
        End If
        If (Trim(meFolIni) <> Empty) And optMov(1) Then
           SQL = SQL & " and (right(a.Identificador,4)='" & meFolIni.Text & "')"
        End If
           
        If txtSerie <> Empty Then
            SQL = SQL & " and doc.serie= '" & txtSerie.Text & "'"
        End If
        If (txtCorrIni <> Empty Or txtCorrFin <> Empty) Then
            If txtCorrIni <> Empty And txtCorrFin = Empty Then
                SQL = SQL & " and (doc.correl >= '" & txtCorrIni.Text & "' and doc.correl <='" & txtCorrIni.Text & "') "
            End If
            If txtCorrFin <> Empty And txtCorrIni = Empty Then
                SQL = SQL & " and (doc.correl >= '" & txtCorrFin.Text & "' and doc.correl <='" & txtCorrFin.Text & "') "
            End If
            If txtCorrIni <> Empty And txtCorrFin <> Empty Then
                SQL = SQL & " and (doc.correl >= '" & txtCorrIni.Text & "' and doc.correl <='" & txtCorrFin.Text & "') "
            End If
        End If
        If (txtDiviIni <> Empty Or txtDiviFin <> Empty) Then
            If txtDiviIni <> Empty And txtDiviFin = Empty Then
                SQL = SQL & " and (doc.division >= '" & txtDiviIni.Text & "' and doc.division <='" & txtDiviIni.Text & "') "
            End If
            If txtDiviFin <> Empty And txtDiviIni = Empty Then
                SQL = SQL & " and (doc.division >= '" & txtDiviFin.Text & "' and doc.division <='" & txtDiviFin.Text & "') "
            End If
            If txtDiviIni <> Empty And txtDiviFin <> Empty Then
                SQL = SQL & " and (doc.division >= '" & txtDiviFin.Text & "' and doc.division <='" & txtDiviFin.Text & "') "
            End If
        End If
        If optMov(0) Then
            SQL = SQL & " and (a.anomes= '" & strAnoSistema + strMesSistema & "') "
        End If
        If IsDate(Trim(meFecEini)) Or IsDate(Trim(meFecEfin)) Then
            If IsDate(Trim(meFecEini)) And Not IsDate(Trim(meFecEfin)) Then
                SQL = SQL & " and (doc.Fec_Emision >= '" & Format(meFecEini, "yyyy/mm/dd") & "' and Fec_Emision <= '" & Format(meFecEini, "yyyy/mm/dd") & "')"
            End If
            If Not IsDate(Trim(meFecEini)) And IsDate(Trim(meFecEfin)) Then
                SQL = SQL & " and (doc.Fec_Emision >= '" & Format(meFecEfin, "yyyy/mm/dd") & "' and Fec_Emision <= '" & Format(meFecEfin, "yyyy/mm/dd") & "')"
            End If
            If IsDate(Trim(meFecEini)) And IsDate(Trim(meFecEfin)) Then
                SQL = SQL & " and (doc.Fec_Emision >= '" & Format(meFecEini, "yyyy/mm/dd") & "' and Fec_Emision <= '" & Format(meFecEfin, "yyyy/mm/dd") & "')"
            End If
        End If
        If IsDate(Trim(meFecRini)) Or IsDate(Trim(meFecRfin)) Then
            If IsDate(Trim(meFecRini)) And Not IsDate(Trim(meFecRfin)) Then
                SQL = SQL & " and (a.Fecha_Registro>= '" & Format(meFecRini, "yyyy/mm/dd") & "' and a.Fecha_Registro <= '" & Format(meFecRini, "yyyy/mm/dd") & "')"
            End If
            If Not IsDate(Trim(meFecRini)) And IsDate(Trim(meFecRfin)) Then
                SQL = SQL & " and (a.Fecha_Registro >= '" & Format(meFecRfin, "yyyy/mm/dd") & "' and a.Fecha_Registro <= '" & Format(meFecRfin, "yyyy/mm/dd") & "')"
            End If
            If IsDate(Trim(meFecRini)) And IsDate(Trim(meFecRfin)) Then
                SQL = SQL & " and (a.Fecha_Registro >= '" & Format(meFecRini.Text, "yyyy/mm/dd") & "' and a.Fecha_Registro <= '" & Format(meFecRfin, "yyyy/mm/dd") & "')"
            End If
        End If
        If CboTipo.ListIndex > 0 Then
            SQL = SQL & " and (a.cod_tipo_doc='" & CboTipo.List(CboTipo.ListIndex, 1) & "')"
        End If
        If cboAux.ListIndex > 0 And txtCodigo <> "" Then
            SQL = SQL & " and (doc.Auxiliar='" & cboAux.List(cboAux.ListIndex, 1) & _
                        "' and doc.codigo='" & txtCodigo & "')"
        End If
        If cboAux.ListIndex > 0 And txtCodigo = "" Then
            SQL = SQL & " and (doc.Auxiliar='" & cboAux.List(cboAux.ListIndex, 1) & "')"
        End If
        If meImporte <> Empty Then
            SQL = SQL & " and doc.total=" & CDbl(meImporte.Text) & " "
        End If
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    Next
    SQL = "Select Identificador,tipo,documento,fec_registro,fec_emision,auxiliar," & _
          " codigo,TRIM(cenco) AS CENCO,mon,total,desestado  from temprptdocumento where usu='" & strUsuarioId & "' " & orden
    LlenarTDocs
    Set rsdetalle = oConexion.EjecutaSelectRS(SQL)
    LLenarDatos rsdetalle
    Set rsdetalle = Nothing
End Function

Public Sub LlenarTDocs()
    Dim rsTdocs As MYSQL_RS
    Dim SQL As String
    Dim I As Integer
    I = 1
    cmbTDocs.Clear
    SQL = "select distinct destipo,tipo,familia  From temprptdocumento where usu='" & strUsuarioId & "' order by destipo "
    Set rsTdocs = oConexion.EjecutaSelectRS(SQL)
    cmbTDocs.AddItem "Selecionar..."
    cmbTDocs.List(0, 1) = "00"
    cmbTDocs.List(0, 2) = 0
    If rsTdocs.RecordCount > 0 Then
        Do While Not rsTdocs.EOF
            cmbTDocs.AddItem CE(rsTdocs.Fields("destipo"))
            cmbTDocs.List(I, 1) = CE(rsTdocs.Fields("tipo"))
            cmbTDocs.List(I, 2) = CE(rsTdocs.Fields("familia"))
            I = I + 1
            rsTdocs.MoveNext
            filtro = True
        Loop
        cmbTDocs.Enabled = True
        cmbTDocs.BackColor = ColorHabilitado
    End If
    cmbTDocs.ListIndex = 0
    Set rsTdocs = Nothing
End Sub

Public Sub LlenarAux()
    Dim rsaux As MYSQL_RS
    Dim SQL As String
    Dim I As Integer
    I = 0
    cboAux.Clear
    SQL = "select aux,descripcion From auxiliares order by aux "
    Set rsaux = oConexion.EjecutaSelectRS(SQL)
    If rsaux.RecordCount > 0 Then
        Do While Not rsaux.EOF
            cboAux.AddItem CE(rsaux.Fields("descripcion"))
            cboAux.List(I, 1) = CE(rsaux.Fields("aux"))
            I = I + 1
            rsaux.MoveNext
        Loop
    End If
    cboAux.ListIndex = 0
    Set rsaux = Nothing
End Sub

Private Sub Llenargrilladetalle()
    Dim I As Integer
    With flexDetalle
        .Clear
        .Rows = 1
        .Cols = 12
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        .FixedRows = 0
        .ColWidth(1) = 450
        .TextMatrix(0, 1) = "Folio"
        .ColWidth(2) = 500
        .TextMatrix(0, 2) = "Tipo"
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = Space(6) + "N° Documento"
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Fec. Registro"
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = "Fec. Emision"
        .ColWidth(6) = 400
        .TextMatrix(0, 6) = "Aux"
        .ColWidth(7) = 1200
        .TextMatrix(0, 7) = Space(6) + "Cod.Aux"
        .ColWidth(8) = 1200
        .TextMatrix(0, 8) = Space(6) + "Cenco"
        .ColWidth(9) = 400
        .TextMatrix(0, 9) = Space(0) + "Mon"
        .ColWidth(10) = 1200
        .TextMatrix(0, 10) = Space(6) + "Importe"
        .ColWidth(11) = 1200
        .TextMatrix(0, 11) = Space(6) + "Estado"
        For I = 0 To .Cols - 1
            .row = 0
            .Col = I
            .CellForeColor = &H80000002
            .CellBackColor = &H8000000F
        Next I
    End With
End Sub

Public Sub LLenarDatos(r As MYSQL_RS)
    Dim J As Integer
    With flexDetalle
        If r.RecordCount = 0 Then
            MsgBox "No se encuentra Datos", vbInformation, gsNomSW
            cmbTDocs.Enabled = False
            cmbTDocs.BackColor = ColorDeshabilitado
            filtro = False
        Else
            .Rows = r.RecordCount + 1
            .FixedRows = 1
            .BackColor = vbWhite
            .row = 0
             Do While Not (r.EOF)
                .row = .row + 1
                For J = 1 To .Cols - 1
                    .TextMatrix(.row, J) = CE(r.Fields(J - 1))
                    If r.Fields(J - 1).name Like "total" Then
                        .TextMatrix(.row, J) = FormatNumber(CEN(r.Fields(J - 1)), 2)
                    End If
                    If r.Fields(J - 1).name Like "*fec*" Then
                        .TextMatrix(.row, J) = Format(CE(r.Fields(J - 1)), "dd/mm/yyyy")
                    End If
                    If r.Fields(J - 1).name Like "*esta*" Then
                        .Col = J
                        .CellForeColor = vbRed
                    End If
                Next
                r.MoveNext
            Loop
            EnumerarItems1 flexDetalle
        End If
     End With
     r.CloseRecordset
     Set r = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Set oConsulta = Nothing
End Sub

Private Sub meFecEfin_Change()
    If meFecEfin <> "" Then
        If VSel(7).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(7).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(7).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meFecEfin_GotFocus()
    mark1 meFecEfin
End Sub

Private Sub meFecEfin_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        meFecRini.SetFocus
    End If
End Sub

Private Sub meFecEini_Change()
    If meFecEini <> "" Then
        If VSel(6).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(6).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(6).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meFecEini_GotFocus()
   mark1 meFecEini
End Sub

Private Sub meFecEini_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        meFecEfin.SetFocus
    End If
End Sub

Private Sub meFecRfin_Change()
    If meFecRfin <> "" Then
        If VSel(9).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(9).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(9).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meFecRfin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        meFolIni.SetFocus
    End If
End Sub

Private Sub meFecRini_Change()
    If meFecRini <> "" Then
        If VSel(8).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(8).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(8).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meFecRini_GotFocus()
    mark1 meFecRini
End Sub

Private Sub meFecRini_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        meFecRfin.SetFocus
    End If
End Sub

Private Sub meFolIni_Change()
    If meFolIni <> "" Then
        If VSel(5).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(5).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(5).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meFolIni_GotFocus()
    mark1 meFolIni
End Sub

Private Sub meFolIni_LostFocus()
    If Trim(meFolIni) <> Empty Then
        meFolIni = Right("0000" + Trim(meFolIni), 4)
    Else
        meFolIni = "    "
    End If
End Sub

Private Sub meImporte_Change()
    If meImporte <> "" Then
        If VSel(4).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(4).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(4).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub meImporte_GotFocus()
    mark1 meImporte
End Sub

Private Sub meImporte_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        meFecEini.SetFocus
    End If
End Sub

Private Sub optOrden_Click(Index As Integer)
    Select Case Index
        Case 0: orden = IIf(OptSentido(0), " Order by tipo asc ", " Order by tipo desc")
        Case 1: orden = IIf(OptSentido(0), " Order by Documento asc ", "Order by Documento desc")
        Case 2: orden = IIf(OptSentido(0), " Order by auxiliar ,descod asc ", "Order by auxiliar ,descod desc")
        Case 3: orden = IIf(OptSentido(0), " Order by fec_emision asc ", "Order by fec_emision as desc")
        Case 4: orden = IIf(OptSentido(0), " Order by fec_registro asc ", "Order by fec_registro desc")
        Case 5: orden = IIf(OptSentido(0), " Order by fec_vcto asc ", "Order by fec_vcto desc")
        Case 6: orden = IIf(OptSentido(0), " Order by Total asc ", "Order by Total desc")
        Case 7: orden = IIf(OptSentido(0), " Order by Estado asc ", "Order by Estado desc")
    End Select
End Sub

Private Sub OptSentido_Click(Index As Integer)
    If Index = 0 Then
        orden = Left(Trim(orden), Len(Trim(orden)) - 4) & " asc "
    Else
        orden = Left(Trim(orden), Len(Trim(orden)) - 4) & " desc"
    End If
End Sub

Private Sub txtCodigo_Change()
    If txtCodigo <> "" Then
        If VSel(13).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(13).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(13).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Lista de Auxiliares"
            .pForm = FORM_BUSQ_DOC
            .pCaso = Label_Descrip_Auxil
            .Show
        End With
    End If
End Sub

Private Sub txtCorrFin_Change()
    If txtCorrFin <> "" Then
        If VSel(3).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(3).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(3).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtCorrFin_GotFocus()
    'mark txtCorrFin
End Sub

Private Sub txtCorrFin_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        meImporte.SetFocus
    End If
End Sub

Private Sub txtCorrFin_LostFocus()
    'If txtCorrFin <> Empty Then txtCorrFin = Right("000000000" + Trim(txtCorrFin), 9)
End Sub

Private Sub txtCorrIni_Change()
    If txtCorrIni <> "" Then
        If VSel(2).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(2).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(2).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtCorrIni_GotFocus()
    'mark txtCorrIni
End Sub

Private Sub txtCorrIni_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtCorrFin.SetFocus
    End If
End Sub

Private Sub txtCorrIni_LostFocus()
     'If txtCorrIni <> Empty Then txtCorrIni = Right("000000000" + Trim(txtCorrIni), 9)
End Sub

Private Sub txtDiviFin_Change()
    If txtDiviFin <> "" Then
        If VSel(11).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(11).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(11).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtDiviFin_GotFocus()
        mark txtDiviFin
End Sub

Private Sub txtDiviFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 600
            .pCol = 1: .pAnchoCol = 2500
            .pTitulo = "ccHFM"
            .pForm = FORM_RPT_DOCS
            .pCaso = LABEL_DIVISIONES
            .Show
        End With
    End If
    PressF1 = False
End Sub

Private Sub txtDiviFin_LostFocus()
    If txtDiviFin <> Empty Then txtDiviFin = Right("0000" + Trim(txtDiviFin), 4)
End Sub

Private Sub txtDiviIni_Change()
    If txtDiviIni <> "" Then
        If VSel(10).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(10).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(10).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtDiviIni_GotFocus()
    mark txtDiviIni
End Sub

Private Sub txtDiviIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 600
            .pCol = 1: .pAnchoCol = 2500
            .pTitulo = "ccHFM"
            .pForm = FORM_RPT_DOCS
            .pCaso = LABEL_DIVISIONES
            .Show
        End With
    End If
    PressF1 = True
End Sub

Private Sub txtDiviIni_KeyPress(KeyAscii As Integer)
    If Keyscii = 13 Then
        mark txtDiviFin
    End If
End Sub

Private Sub txtDiviIni_LostFocus()
    If txtDiviIni <> Empty Then txtDiviIni = Right("0000" + Trim(txtDiviIni), 4)
End Sub

Private Sub txtSerie_Change()
    If txtSerie <> "" Then
        If VSel(1).Flg = False Then
            ContFiltros = ContFiltros + 1
            VSel(1).Flg = True
        End If
    Else
        ContFiltros = ContFiltros - 1
        VSel(1).Flg = False
    End If
    If ContFiltros >= 3 Then
        optMov(1).Enabled = True
    Else
        optMov(1).Enabled = False
    End If
End Sub

Private Sub txtSerie_GotFocus()
    'mark txtSerie
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCorrIni.SetFocus
    End If
End Sub

Private Sub txtSerie_LostFocus()
    'If txtSerie <> Empty Then txtSerie = Right("00000" + Trim(txtSerie), 5)
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flexDetalle
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 1
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        .TopRow = NewValue
    End With
End Sub

Public Sub LlenarTipoDocs()
    Dim RQ As MYSQL_RS
    Dim SQL As String
    Dim I As Integer
    I = 1
    CboTipo.Clear
    SQL = "select coddoc,descrip from cndocum C where (PROTEGIDO = 'N' OR " & _
          "(SELECT PERMISO FROM docsusuario D WHERE D.CODDOC=C.CODDOC AND USUARIO = '" & strUsuarioId & "')=1) " & _
          "order by descrip"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    CboTipo.AddItem "Seleccionar..."
    If RQ.RecordCount > 0 Then
        Do While Not RQ.EOF
            CboTipo.AddItem CE(RQ.Fields("descrip"))
            CboTipo.List(I, 1) = CE(RQ.Fields("coddoc"))
            I = I + 1
            RQ.MoveNext
        Loop
    End If
    CboTipo.ListIndex = 0
    Set RQ = Nothing
End Sub

Sub LlenaArregloSel()
    For I = 1 To 14
        VSel(I).Ctrl = I
        VSel(I).Flg = False
    Next
End Sub
