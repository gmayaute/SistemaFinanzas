VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmPagoDctos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Descuentos"
   ClientHeight    =   6435
   ClientLeft      =   3150
   ClientTop       =   3570
   ClientWidth     =   12600
   Icon            =   "frmPagoDctos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12600
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Height          =   525
      Left            =   90
      TabIndex        =   12
      Top             =   3960
      Width           =   11205
      Begin VB.Label lblNumDctos 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Height          =   285
         Left            =   9630
         TabIndex        =   24
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Num Dsctos:"
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
         Left            =   8370
         TabIndex        =   23
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Nro:"
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
         Left            =   4680
         TabIndex        =   22
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Num. Registros:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lblFac 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "000001-000000001"
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
         Height          =   315
         Left            =   6090
         TabIndex        =   20
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lblNumReg 
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   315
         Left            =   1590
         TabIndex        =   13
         Top             =   180
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Height          =   285
         Left            =   60
         TabIndex        =   19
         Top             =   150
         Width           =   11025
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   5115
      Left            =   0
      TabIndex        =   11
      Top             =   1350
      Width           =   12555
      Begin NOVAdmin.flxEdit flxDetDctos 
         Height          =   1905
         Left            =   120
         TabIndex        =   25
         Top             =   3180
         Width           =   11925
         _extentx        =   21034
         _extenty        =   3360
         font            =   "frmPagoDctos.frx":014A
         cellfontname    =   "MS Sans Serif"
         cellfontsize    =   8.25
         backcolorsel    =   -2147483643
         backcolorfixed  =   9868950
         cellpicture     =   "frmPagoDctos.frx":0176
         colalignment0   =   9
         fixedalignment0 =   9
         forecolorsel    =   16711680
         forecolorfixed  =   14474460
         rowheight0      =   240
         mouseicon       =   "frmPagoDctos.frx":0194
      End
      Begin MSFlexGridLib.MSFlexGrid flxPagoDctos 
         Height          =   2475
         Left            =   90
         TabIndex        =   15
         Top             =   180
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   4366
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin Proyecto1.chameleonButton btnInsertar 
         Height          =   345
         Left            =   12030
         TabIndex        =   16
         Top             =   3300
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
         MICON           =   "frmPagoDctos.frx":01B2
         PICN            =   "frmPagoDctos.frx":01CE
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
         Left            =   12030
         TabIndex        =   17
         ToolTipText     =   "Reporte"
         Top             =   4170
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
         MICON           =   "frmPagoDctos.frx":0328
         PICN            =   "frmPagoDctos.frx":0344
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
         Left            =   12030
         TabIndex        =   18
         ToolTipText     =   "Salir"
         Top             =   4620
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
         MICON           =   "frmPagoDctos.frx":0886
         PICN            =   "frmPagoDctos.frx":08A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "FACTURA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1305
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   12555
      Begin VB.OptionButton optBusq 
         BackColor       =   &H009F5539&
         Caption         =   "Documentos con Dcto."
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
         Height          =   315
         Index           =   2
         Left            =   7050
         TabIndex        =   14
         Top             =   900
         Width           =   2445
      End
      Begin VB.OptionButton optBusq 
         BackColor       =   &H009F5539&
         Caption         =   "Otros Documentos"
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
         Height          =   315
         Index           =   1
         Left            =   7050
         TabIndex        =   4
         Top             =   570
         Width           =   1965
      End
      Begin VB.OptionButton optBusq 
         BackColor       =   &H009F5539&
         Caption         =   "Afecto a Detracción"
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
         Height          =   285
         Index           =   0
         Left            =   7050
         TabIndex        =   3
         Top             =   270
         Width           =   2175
      End
      Begin Proyecto1.chameleonButton btnBuscar 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   9270
         TabIndex        =   5
         ToolTipText     =   "Realizar Búsqueda"
         Top             =   570
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Buscar"
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
         MICON           =   "frmPagoDctos.frx":0C68
         PICN            =   "frmPagoDctos.frx":0C84
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdBuscar01 
         Height          =   345
         Left            =   11280
         TabIndex        =   26
         Top             =   810
         Width           =   405
         _ExtentX        =   2143
         _ExtentY        =   714
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
         MICON           =   "frmPagoDctos.frx":3006
         PICN            =   "frmPagoDctos.frx":3022
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdVer01 
         Height          =   375
         Left            =   11880
         TabIndex        =   27
         Top             =   840
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
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPagoDctos.frx":35BC
         PICN            =   "frmPagoDctos.frx":35D8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComDlg.CommonDialog CmD 
         Left            =   11880
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Documentos Excel"
      End
      Begin Proyecto1.chameleonButton btnInterfazAsiento 
         Height          =   345
         Left            =   10830
         TabIndex        =   30
         ToolTipText     =   "Genera Asientos Mochona"
         Top             =   180
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
         MICON           =   "frmPagoDctos.frx":3B1A
         PICN            =   "frmPagoDctos.frx":3B36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblDistrito 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Archivos:"
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
         Index           =   15
         Left            =   11280
         TabIndex        =   29
         Top             =   600
         Width           =   810
      End
      Begin MSForms.TextBox txtFolio 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   690
         Width           =   2475
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   10
         Size            =   "4366;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Folio Ref.:"
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
         Height          =   315
         Left            =   4200
         TabIndex        =   10
         Top             =   330
         Width           =   2475
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   390
         Left            =   1350
         TabIndex        =   9
         Top             =   660
         Width           =   210
      End
      Begin MSForms.TextBox txtCorrelativo 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   690
         Width           =   2295
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   9
         Size            =   "4048;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   1590
         TabIndex        =   8
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   1215
      End
      Begin MSForms.TextBox txtSerie 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   690
         Width           =   1215
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   5
         Size            =   "2143;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   12840
      TabIndex        =   28
      Top             =   -240
      Width           =   2295
   End
End
Attribute VB_Name = "frmPagoDctos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As New FrmConsultas
Private TmpDcto As Double
Private FilaSel As Integer

Private Sub ConfigGrilla()
    With flxPagoDctos
        .Clear
        .Rows = 1
        .Cols = 20
        .ForeColorFixed = &H404000
        
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = Space(1) + "Identificador"
        
        .ColWidth(2) = 3500
        .TextMatrix(0, 2) = Space(30) + "Cliente"
            
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = Space(8) + "Nro Fact"
        
        .ColWidth(4) = 1350
        .TextMatrix(0, 4) = Space(8) + "Importe"
    
        .ColWidth(5) = 1350
        .TextMatrix(0, 5) = Space(8) + "Total Dcto."
        
        .ColWidth(6) = 1350
        .TextMatrix(0, 6) = Space(8) + "Saldo"
        
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = Space(1) + "Asiento"
    End With
End Sub

Private Sub ConfigGrillaDet()
    With flxDetDctos
        .Clear
        .Rows = 1
        .Cols = 10
        .ForeColorFixed = &H404000
        
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .ColType(0) = cadena
        .FixedCols = 1
        
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = Space(2) + "Código"
        .CaracteresValidos(1) = "0123456789"
        .ColType(1) = cadena
        .ColMaxLength(1) = 2
        
        .ColWidth(2) = 3500
        .TextMatrix(0, 2) = Space(17) + "Descripción"
        .CaracteresValidos(2) = "ab cd" & Chr(13) & "efghijklmnopqrstuvwxyz" & UCase("abcdefghijklmnopqrstuvwxyz") & ""
        .ColType(2) = cadena
        .ColMaxLength(2) = 25
        
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = Space(6) + "Ser. Doc."
        .CaracteresValidos(3) = "1234567890-"
        .ColType(3) = cadena
        .ColMaxLength(3) = 15
        
        .ColWidth(4) = 1500
        .TextMatrix(0, 4) = Space(6) + "Nro. Doc."
        .CaracteresValidos(4) = "1234567890-"
        .ColType(4) = cadena
        .ColMaxLength(4) = 15
        
        .ColWidth(5) = 1350
        .TextMatrix(0, 5) = Space(5) + "Importe. Dcto"
        .CaracteresValidos(5) = "1234567890."
        .ColType(5) = cadena
        .ColMaxLength(5) = 20
        
        .ColWidth(6) = 1000
        .TextMatrix(0, 6) = Space(1) + "Fecha"
        .CaracteresValidos(6) = "0123456789//"
        .ColType(6) = fecha
        .ColMaxLength(6) = 10
        .ColAlignment(6) = 6
                
        .ColWidth(7) = 1000
        .TextMatrix(0, 7) = Space(1) + "Cancelado"
        .CaracteresValidos(7) = "1234567890."
        .ColType(7) = cadena
        .ColMaxLength(7) = 2
                
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
End Sub

Private Sub btnBuscar_Click()
    If Trim(txtFolio) <> Empty Then BusqxCod "", Trim(txtFolio):  Exit Sub
    If txtCorrelativo <> Empty Then BusqxCod Right("00000" & Trim(txtSerie), 5) & "-" & Right("000000000" & Trim(txtCorrelativo), 9), "": txtCorrelativo.SetFocus:  Exit Sub
    If txtCorrelativo = Empty Then Busqueda: Exit Sub
End Sub

Private Sub btnInsertar_Click()
    With flxDetDctos
        .Rows = .Rows + 1
        .row = .Rows - 1
        .Col = 7
        .CellFontName = "Wingdings"
        .CellFontSize = 11
        .CellBackColor = ColorDeshabilitado
        .ColAlignment(7) = flexAlignCenterBottom
        .TextMatrix(.row, 7) = strUnChecked
        EnumerarItems flxDetDctos
        Publimensaje = "modifica"
        .Col = 1
    End With
End Sub

Private Sub btnReporte_Click()
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "REGISTRO DE DOCUMENTOS CON DESCUENTO"
    oReporte.Reporte = "Rep_Docs_Descuentos.rpt"
    oReporte.sp_Pagos_Descuentos
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cmdBuscar01_Click()
            Dim RutaOrigen As String
            Dim RutaDestinoCliente As String
            Dim NomArchivo As String, SQL As String
            Dim TxtCodigoID As String
                CmD.CancelError = True
                CmD.Flags = cdlOFNFileMustExist Or _
                        cdlOFNHideReadOnly Or _
                        cdlOFNExplorer Or _
                        cdlOFNLongNames
            On Error Resume Next
                CmD.ShowOpen
                If err.Number = cdlCancel Or err.Number <> 0 Then Exit Sub
                RutaOrigen = CmD.Filename
                If RutaOrigen = "" Then Exit Sub
            On Error GoTo ErrorAbrir
                RutaDestinoCliente = "\\172.26.35.12\ddsi$\Cobranzas"
                NomArchivo = Dir(RutaOrigen)
                TxtCodigoID = "I_" & Trim(flxPagoDctos.TextMatrix(FilaSel, 1))
                
                If NomArchivo <> "" Then
                    If EncuentraArchivo(NomArchivo, TxtCodigoID) = False Then
                        CopiarArchivos RutaOrigen, RutaDestinoCliente, TxtCodigoID
                        SQL = "call Insert_ArchAdjuntos('11','" & TxtCodigoID & "', " & _
                              "'" & Replace(RutaDestinoCliente & "\" & TxtCodigoID & "\", "\", "*") & "','" & NomArchivo & "')"
                        ADOConexion.Execute (SQL)
                        MsgBox "Archivo: " & NomArchivo & " adjuntado", vbInformation, "NOVPeru"
                    Else
                        MsgBox "Archivo: " & NomArchivo & " ya se encuentra adjuntado a este documento", vbInformation, "NOVPeru"
                    End If
                End If
            Exit Sub
ErrorAbrir:
                MsgBox err.Description
                MsgBox "Ocurrió un error al momento de copiar el archivo, " & vbNewLine & _
                       "consulte con el administrador del sistema", vbExclamation + vbOKOnly, "NOVPeru"
End Sub
 
Private Sub cmdVer01_Click()
 Dim cDestino As String
 Dim SQL As String
 Dim RutaDestinoCliente As String
 Dim NombreArchivoVer As String
 
 frmArchivosAdjuntos.AnioSel = strAnoSistema
 frmArchivosAdjuntos.MesSel = strMesSistema
 frmArchivosAdjuntos.IdentificadorAr = "I_" & Trim(flxPagoDctos.TextMatrix(FilaSel, 1))
 frmArchivosAdjuntos.Show
 
End Sub

Private Sub flxDetDctos_Click()
    Dim Flg As Boolean, nfd As Integer
    Flg = False
    
     With flxDetDctos
        nfd = .row
        If .Col = 7 Then
            
            .TextMatrix(nfd, 7) = IIf(.TextMatrix(nfd, 7) = strChecked, strUnChecked, strChecked)
            
            If .TextMatrix(nfd, 7) = strUnChecked Then
               If MsgBox("¿Estas Seguro que quieres eliminar la linea con importe " & .TextMatrix(nfd, 5) & " ?", vbYesNo + vbQuestion, "NOV") = vbNo Then
                 Exit Sub
               End If
               If MsgBox("¿Seguro que desea eliminar la linea con importe " & .TextMatrix(nfd, 5) & ", porque podria estar eliminando data importante ?", vbYesNo + vbQuestion, "NOV") = vbNo Then
                 Exit Sub
               End If
            End If
            
            If .TextMatrix(nfd, 7) = strChecked Then Flg = True
            InsertarDetalle
            If Flg = True Then
                flxPagoDctos.row = FilaSel
                'GenerarAsientos nfd
                'If VerificaCancelacion(Trim(flxPagoDctos.TextMatrix(flxPagoDctos.row, 1))) Then
                '    ActualizaEstadoDoc (Trim(flxPagoDctos.TextMatrix(flxPagoDctos.row, 1)))
                'End If
            End If
            flxPagoDctos.SetFocus
            Call keybd_event(vbKeyHome, 0, 0, 0)
        End If
    End With
End Sub



Private Function VerificaCancelacion(identificador As String) As Boolean
    Dim SQL As String
    Dim rsdctoC As MYSQL_RS
    Dim TotFact As Double
    Dim PartFact As Double
    VerificaCancelacion = False
    
    SQL = " SELECT a.Importe, b.Total_Ref," & _
          " a.fecha as fec from detalledctos as a left JOIN documento_contables as b " & _
          " on a.identificador = b.identificador  where a.identificador = '" & identificador & "'"
    Set rsdctoC = oConexion.EjecutaSelectRS(SQL)
    lblNumDctos = rsdctoC.RecordCount
    If rsdctoC.RecordCount > 0 Then
        Do While Not rsdctoC.EOF
         TotFact = rsdctoC.Fields("Total_Ref")
         PartFact = PartFact + rsdctoC.Fields("Importe")
          If TotFact = PartFact Then
            VerificaCancelacion = True
          End If
        rsdctoC.MoveNext
        Loop
    End If
    
    Set rsdctoC = Nothing
    
End Function

Private Sub ActualizaEstadoDoc(identificador As String)
  Dim SQLACT As String
  Dim FechaCanc As String
  SQLACT = "Update movi_documento set Cod_Estado='CA' where Identificador='" & identificador & "'"
  oConexion.EjecutaInsertUpdateDelete SQLACT, TIPO_QUERY.Modificar, False
  SQLACT = "Update documento_contables set Cancelado='0.00' where Identificador='" & identificador & "'"
  oConexion.EjecutaInsertUpdateDelete SQLACT, TIPO_QUERY.Modificar, False
  FechaCanc = InputBox("Ingrese la Fecha de Cancelación del Documento formato:DD/MM/YYYY", "Cancelación de Documentos", "R")
  GuardarFecCancelacion FechaCanc, identificador
End Sub


Public Sub GuardarFecCancelacion(fecha As String, Ident As String)
    Dim SQL As String
    If Trim(fecha) <> "" Then
        SQL = " Update documento_contables set Fec_Pago=" & _
              " '" & Format(CDate(fecha), "yyyy/mm/dd") & "'" & _
              " where Identificador='" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    Else
        SQL = " Update documento_contables set Fec_Pago = ''" & _
              " where Identificador='" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    End If
End Sub


Private Sub flxDetDctos_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxDetDctos
        If KeyCode = vbKeyF1 Then
            If .Col = 1 Then
                With oConsulta
                    .pCols = 3
                    .pCol = 0: .pAnchoCol = 1200
                    .pCol = 1: .pAnchoCol = 3000
                    .pCol = 2: .pAnchoCol = 1000
                    .pTitulo = "Consulta de Docs Lquidacion"
                    .pForm = FORM_PAGODCTOS
                    .pCaso = LABEL_TIP_DOC
                    .Show
                End With
                .SiguienteCelda
            End If
        End If
        If .Col = 6 Then
            TipodeCampo = fecha
        Else
            TipodeCampo = cadena
        End If
    End With
End Sub

Private Sub flxDetDctos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With flxDetDctos
            If .Col = 5 Then
                If Trim(.TextMatrix(.row, 5)) <> Empty Then
                    .TextMatrix(.row, 5) = FormatNumber(.TextMatrix(.row, 5), 2)
                End If
            End If
            If .Col = 1 Then
                If Trim(.TextMatrix(.row, 1)) <> Empty Then
                    .TextMatrix(.row, 1) = Space(2) & .TextMatrix(.row, 1)
                    .TextMatrix(.row, 2) = Space(2) & DescripcionesdeCodigos("CNDOCUM", Trim(.TextMatrix(.row, 1)), "1")
                End If
            End If
            If .Col = 2 And Trim(.TextMatrix(.row, 2)) <> Empty Then
                .TextMatrix(.row, 2) = Space(2) & .TextMatrix(.row, 2)
            End If
            If .Col = 3 Then
                .TextMatrix(.row, 3) = Space(2) & .TextMatrix(.row, 3)
            End If
            If .Col = 4 And Trim(.TextMatrix(.row, 4)) <> Empty Then
                .TextMatrix(.row, 4) = Space(2) & .TextMatrix(.row, 4)
            End If
        End With
    End If
End Sub

Private Sub flxDetDctos_RowColChange()
    With flxDetDctos
        If Trim(.TextMatrix(.row, 7)) = strUnChecked Then
            Publimensaje = "modificar"
        Else
            Publimensaje = "no-modificar"
        End If
    End With
End Sub

Private Sub InsertarDetalle()
    Dim SQL As String
    Dim I As Integer
    SQL = "Delete from detalledctos where identificador = '" & Trim(flxPagoDctos.TextMatrix(FilaSel, 1)) & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    With flxDetDctos
        For I = 1 To .Rows - 1
            If Trim(.TextMatrix(I, 7)) = strChecked And (Trim(.TextMatrix(I, 5)) <> Empty And Trim(.TextMatrix(I, 5)) <> "0.00") Then
                SQL = "Call Insert_DetalleDctos ('" & Trim(flxPagoDctos.TextMatrix(FilaSel, 1)) & "'," & _
                      " '" & Right("00" & Trim(.TextMatrix(I, 1)), 2) & "', '" & Trim(.TextMatrix(I, 2)) & "'," & _
                      " '" & Trim(.TextMatrix(I, 3)) & "','" & Trim(.TextMatrix(I, 4)) & "', " & CDbl(CEN(.TextMatrix(I, 5))) & ", '" & Format(.TextMatrix(I, 6), "yyyy/mm/dd") & "' );"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
            End If
        Next
    End With
    SQL = "Call Update_PagoDctos ('" & Trim(flxPagoDctos.TextMatrix(FilaSel, 1)) & "', " & CDbl(CalcImpRef) & ")"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    
    GenerarAsientos I - 1
    Busqueda
End Sub

Private Function CalcImpRef() As Double
    Dim I As Integer
    CalcImpRef = 0
    With flxDetDctos
        For I = 1 To .Rows - 1
            If Trim(.TextMatrix(I, 7)) = strChecked Then
                CalcImpRef = CalcImpRef + CDbl(CEN(.TextMatrix(I, 5)))
            End If
        Next
    End With
End Function

Private Sub flxPagoDctos_Click()
    DesplazarxGrilla
End Sub


        
Private Sub DesplazarxGrilla()
    With flxPagoDctos
        lblFac = Trim(.TextMatrix(.row, 3))
        FilaSel = .row
        If Not CargarDetalle(Trim(.TextMatrix(.row, 1))) Then
            ConfigGrillaDet
        End If
        .TextMatrix(.row, 11) = IIf(.TextMatrix(.row, 11) = strChecked, strUnChecked, strChecked)
    End With
End Sub

Private Sub flxPagoDctos_RowColChange()
   DesplazarxGrilla
End Sub

'Carga el detalle de la Grilla de item
Private Function CargarDetalle(identificador As String) As Boolean
    Dim SQL As String
    Dim rsdcto As MYSQL_RS
    CargarDetalle = False
    SQL = " SELECT a.Identificador , a.cod, a.descrip, a.SERDoc, a.NumDoc, B.total, a.Importe, b.Total_Ref," & _
          " (CASE A.COD WHEN '17' THEN 'N' Else 'E' END) as Moneda," & _
          " (CASE A.COD WHEN '17' THEN '00000000004' ELSE '00000000002' END) AS auxil," & _
          " a.fecha as fec from detalledctos as a left JOIN documento_contables as b " & _
          " on a.identificador = b.identificador  where a.identificador = '" & identificador & "'"
    Set rsdcto = oConexion.EjecutaSelectRS(SQL)
    lblNumDctos = rsdcto.RecordCount
    If rsdcto.RecordCount > 0 Then
        CargarDetalle = True
        CargarGrillaDet rsdcto
    End If
    Set rsdcto = Nothing
End Function

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Publimensaje = "modificar"
    TipodeCampo = cadena
    optBusq(0).Value = True
    FilaSel = 1
    txtSerie = serie
    Busqueda
    DesplazarxGrilla
    Call WheelHook(frmPagoDctos)
    Set oConsulta = New FrmConsultas
End Sub

Private Sub Busqueda()
    Dim SQL As String
    Dim rs2 As MYSQL_RS
    SQL = " SELECT a.Identificador," & _
          " (Select descrip from cnauxil as e where e.codigo = a.codigo and e.auxiliar ='2') as cliente," & _
          " CONCAT(a.serie,'-',a.correl) as Numero,a.Total,a.Total_Ref,a.Total-a.Total_Ref as saldo," & _
          " a.fec_emision as fechaemi,a.codigo as Cod,a.division as divi,a.mon as Moneda" & _
          " from (((documento_contables as a left join amarre_documento as b " & _
          " on a.identificador= b.identificador) right join detallefact as c " & _
          " on a.identificador = c.identificador)right join movi_documento as d " & _
          " on d.identificador = a.identificador) where left(a.Fec_Emision,4)='" & strAnoSistema & "' and (b.Cod_tipo_doc='01' OR b.Cod_tipo_doc='03') AND A.TOTAL>0  and d.cod_estado = '" & IMPRESO & "'"
    If optBusq(0).Value Then
       SQL = SQL & " and c.afecto = '1' Group By  c.identificador order by Numero"
    End If
    If optBusq(1).Value Then
       SQL = SQL & " and c.afecto = '0' Group By c.identificador order by Numero"
    End If
    If optBusq(2).Value Then
        SQL = SQL & " and a.Total_Ref <> 0 Group By c.identificador  order by Numero"
    End If
    Set rs2 = oConexion.EjecutaSelectRS(SQL)
    lblNumReg = CE(rs2.RecordCount)
    If rs2.RecordCount > 0 Then
        CargarGrilla rs2
        DesplazarxGrilla
    End If
    Set rs2 = Nothing
End Sub

Private Sub CargarGrillaDet(rsdatos As MYSQL_RS)
    Dim I As Integer
    ConfigGrillaDet
    I = 1
    With flxDetDctos
        .Visible = False
        Do While Not rsdatos.EOF
        .Rows = .Rows + 1
        .TextMatrix(I, 1) = Space(2) & Trim(CE(rsdatos.Fields("cod")))
        .TextMatrix(I, 2) = Space(2) & Trim(CE(rsdatos.Fields("descrip")))
        .TextMatrix(I, 3) = Space(2) & Trim(CE(rsdatos.Fields("SERdoc")))
        .TextMatrix(I, 4) = Space(2) & Trim(CE(rsdatos.Fields("NUMdoc")))
        .TextMatrix(I, 5) = FormatNumber(CEN(rsdatos.Fields("Importe")), 2)
        .TextMatrix(I, 6) = Format(CE(rsdatos.Fields("fec")), "dd/mm/yyyy")
        .Col = 7
        .row = I
        .CellFontName = "Wingdings"
        .CellFontSize = 11
        .CellBackColor = ColorDeshabilitado
        .ColAlignment(7) = flexAlignCenterBottom
        If CE(rsdatos.Fields("Total_Ref")) = 0 Then
            .TextMatrix(I, 7) = strUnChecked
        Else
            .TextMatrix(I, 7) = strChecked
        End If
        .TextMatrix(I, 8) = Trim(CE(rsdatos.Fields("auxil")))
        .TextMatrix(I, 9) = Trim(CE(rsdatos.Fields("moneda")))
        I = I + 1
        EnumerarItems flxDetDctos
        rsdatos.MoveNext
        Loop
        .Visible = True
    End With
    Set rsdatos = Nothing
End Sub

Private Sub CargarGrilla(rsdatos As MYSQL_RS)
    Dim I As Integer
    ConfigGrilla
    I = 1
    With flxPagoDctos
        .Visible = False
        Do While Not rsdatos.EOF
            .Rows = .Rows + 1
            .TextMatrix(I, 0) = I
            .TextMatrix(I, 1) = Space(2) & Trim(rsdatos.Fields("Identificador"))
            .TextMatrix(I, 2) = Space(2) & Trim(rsdatos.Fields("cliente"))
            .TextMatrix(I, 3) = Space(2) & Trim(rsdatos.Fields("numero"))
            .TextMatrix(I, 4) = FormatNumber(rsdatos.Fields("Total"), 2)
            .TextMatrix(I, 5) = FormatNumber(rsdatos.Fields("Total_Ref"), 2)
            .TextMatrix(I, 6) = FormatNumber(rsdatos.Fields("saldo"), 2)
            .TextMatrix(I, 7) = Format(rsdatos.Fields("fechaemi"), "dd/mm/yyyy")
            .TextMatrix(I, 8) = Trim(rsdatos.Fields("cod"))
            .TextMatrix(I, 9) = Trim(rsdatos.Fields("divi"))
            .TextMatrix(I, 10) = Trim(rsdatos.Fields("moneda"))
            
            'Checked

            .Col = 11
            .row = I
            .CellFontName = "Wingdings"
            .CellFontSize = 11
            .CellBackColor = ColorHabilitado
            .ColAlignment(11) = flexAlignCenterBottom
            .TextMatrix(I, 11) = strUnChecked
            
            I = I + 1
            rsdatos.MoveNext
        Loop
        If .Rows > 1 Then
            .row = FilaSel
            .ColSel = 6
        End If
        .Visible = True
    End With
    Set rsdatos = Nothing
End Sub

Private Sub BusqxCod(NumFact As String, folio As String)
    Dim SQL As String
    Dim sqlwhere As String
    Dim rsbusq As MYSQL_RS
    
    SQL = " SELECT a.Identificador," & _
          " (Select descrip from cnauxil as e where e.codigo = a.codigo and e.auxiliar ='2') as cliente," & _
          " CONCAT(a.serie,'-',a.correl) as Numero,a.Total,a.Total_Ref,a.Total-a.Total_Ref as saldo, c.afecto," & _
          " a.fec_emision as fechaemi,a.codigo as Cod,a.division as divi,a.mon as Moneda" & _
          " from (((documento_contables as a left join amarre_documento as b " & _
          " on a.identificador= b.identificador) right join detallefact as c " & _
          " on a.identificador = c.identificador)right join movi_documento as d " & _
          " on d.identificador = a.identificador) where (b.Cod_tipo_doc='01' OR b.Cod_tipo_doc='03') AND A.TOTAL>0 and d.cod_estado = '" & IMPRESO & "'"
    If folio <> "" Then
        SQL = SQL & " and a.Identificador = '" & folio & "'  Group By c.identificador order by Numero"
    End If
    If NumFact <> "" Then
        SQL = SQL & "  and CONCAT(a.serie,'-',a.correl)='" & NumFact & "' Group By c.identificador order by Numero "
    End If
    Set rsbusq = oConexion.EjecutaSelectRS(SQL)
    lblNumReg = CE(rsbusq.RecordCount)
    If rsbusq.RecordCount > 0 Then
        If rsbusq.Fields("afecto") = 1 Then
           optBusq(0).Value = True
        Else
           optBusq(1).Value = True
        End If
        CargarGrilla rsbusq
    End If
    Set rsbusq = Nothing
End Sub

Private Function serie() As String
    Dim SQL As String
    Dim rs1 As MYSQL_RS
    SQL = "Select Serie from opcfact where codigo = '01'"
    Set rs1 = oConexion.EjecutaSelectRS(SQL)
    If rs1.RecordCount = 1 Then
        serie = rs1.Fields("Serie")
    Else
        serie = ""
    End If
    Set rs1 = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set oConsulta = Nothing
    Set oReporte = Nothing
    WheelUnHook
End Sub

Private Sub optBusq_Click(Index As Integer)
    FilaSel = 1
    ConfigGrillaDet
    btnBuscar_Click
End Sub

Private Sub txtCorrelativo_Change()
    If txtCorrelativo = Empty Then Busqueda
End Sub

Private Sub txtCorrelativo_GotFocus()
    mark1 txtCorrelativo
End Sub

Private Sub txtCorrelativo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtCorrelativo = Right("000000000" & Trim(txtCorrelativo), 9)
        If txtCorrelativo = "000000000" Then txtCorrelativo = Empty
        BusqxCod Right("00000" & Trim(txtSerie), 5) & "-" & Right("000000000" & Trim(txtCorrelativo), 9), ""
    End If
End Sub

Private Sub txtCorrelativo_LostFocus()
    txtCorrelativo = Right("000000000" & Trim(txtCorrelativo), 9)
    If txtCorrelativo = "000000000" Then txtCorrelativo = Empty
End Sub

Private Sub txtFolio_Change()
    If txtFolio = Empty Then Busqueda
End Sub

Private Sub txtFolio_GotFocus()
    mark1 txtFolio
End Sub

Private Sub txtFolio_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        If Len(Trim(txtFolio)) = 4 Then txtFolio = strAnoSistema & strMesSistema & txtFolio
        BusqxCod "", Trim(txtFolio)
    End If
End Sub

Private Sub txtSerie_GotFocus()
    mark1 txtSerie
End Sub

Private Sub txtSerie_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtSerie = Right("00000" & Trim(txtSerie), 5)
        txtCorrelativo.SetFocus
    End If
End Sub

Private Sub Limpiar()
    txtSerie = Empty
    txtCorrelativo = Empty
    txtFolio = Empty
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
On Error Resume Next
    With flxPagoDctos
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

Private Sub txtSerie_LostFocus()
    txtSerie = Right("00000" & Trim(txtSerie), 5)
    If txtSerie = "00000" Then txtSerie = Empty
End Sub

Sub GenerarAsientos(fila As Integer)
On Error GoTo CtrlError
    Dim AnoMes As String
    Dim lib As String
    Dim v As String, Cta As String, mon As String
    Dim Serdoc As String, Numdoc As String, Div As String, correl As String
    Dim glo As String, aux As String, caux As String
    Dim tc As Double
    Dim fecha As String, Identif As String, TipoPago As String
    Dim SQL As String
    Dim RServ As MYSQL_RS

'    If (Right(Trim(flxDetDctos.TextMatrix(fila, 6)), 4)) <> strAnoSistema Then
'     If MsgBox("¿Desea Continuar, a pesar de que acaba de ingresar un año distinto al del sistema?", vbYesNo + vbQuestion, "Asientos Automaticos") = vbYes Then
'
'     Else
'       MsgBox "No se Inserto nada, proceda a seleccionarlo para eliminar la referencia.", vbCritical, "Asientos Automaticos"
'       Exit Sub
'     End If
'    End If


    lib = "01"
    TipoPago = Trim(flxDetDctos.TextMatrix(fila, 1))
    fecha = Trim(flxDetDctos.TextMatrix(fila, 6))
    Identif = Trim(flxPagoDctos.TextMatrix(FilaSel, 1))
    AnoMes = Right(Trim(flxDetDctos.TextMatrix(fila, 6)), 4) & Mid(Trim(flxDetDctos.TextMatrix(fila, 6)), 4, 2)
    v = MaxVoucher(AnoMes, lib)
    glo = IIf(TipoPago = "17", "DETRACCION ", "COBRANZA ") & Trim(flxPagoDctos.TextMatrix(FilaSel, 2))
    TipoCambio (Trim(flxDetDctos.TextMatrix(fila, 6)))
    tc = dblTipoCmbV
    'mon = Trim(flxDetDctos.TextMatrix(fila, 9))
    mon = "N"
    
    SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & fecha & _
          " ','" & fecha & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
          " ','CUADRADO','','','','','N','','')"
    oConexionMYSQL.Execute (SQL)

    
        '**********************************************
        
        'Dame Tipo de cambio en el registro de ventas para la Detracc
        Serdoc = "F" & Mid(Trim(flxPagoDctos.TextMatrix(FilaSel, 3)), 3, 3)
        Numdoc = Right(Trim(flxPagoDctos.TextMatrix(FilaSel, 3)), 8)
        fecha = vDameFechaFactura(Serdoc, Numdoc)
        TipoCambio (fecha)
        'Fin Dame Tipo de cambio en el registro de ventas para la Detracc
        
        tc = dblTipoCmbV
        Cta = IIf(TipoPago = "17", "104201", "104102")
        dh = "D"
        sol = Round(CDbl(flxDetDctos.TextMatrix(fila, 5)) * tc, 2)
        dol = Round(CDbl(flxDetDctos.TextMatrix(fila, 5)), 2)
        correl = Right("0000" & Trim(CStr(1)), 4)
        Serdoc = Trim(flxDetDctos.TextMatrix(fila, 3))
        Numdoc = Trim(flxDetDctos.TextMatrix(fila, 4)) ' Replace(Format(fecha, "dd/mm/yy"), "/", ".")
        td = "9" 'IIf(TipoPago = "17", "09", "9")
        'caux = Trim(flxDetDctos.TextMatrix(fila, 8))
        aux = 1
        If Cta = "104201" Then
          caux = "00000000004"
        Else
          caux = "00000000002"
        End If
        
        colv = ""
        cto = glo
        Div = "013100003836" '0001 IIf(Trim(flxPagoDctos.TextMatrix(FilaSel, 2)) = "0003", "0003", "0001")
        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Identif & "','" & _
              v & "','" & Trim(Serdoc) & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
              aux & "','" & caux & "','0000','00000000000','N','" & cto & "'," & _
              IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
              IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
              fecha & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
              colv & "','000','')"
        oConexionMYSQL.Execute (SQL)
        '**********************************************

        mon = Trim(flxPagoDctos.TextMatrix(FilaSel, 10))
        Cta = IIf(mon = "N", "121302", "121301")   '"12102" '"12101"
        dh = "H"
        sol = Round(CDbl(flxDetDctos.TextMatrix(fila, 5)) * tc, 2)
        dol = Round(CDbl(flxDetDctos.TextMatrix(fila, 5)), 2)
        correl = Right("0000" & Trim(CStr(2)), 4)
        Serdoc = "F" & Mid(Trim(flxPagoDctos.TextMatrix(FilaSel, 3)), 3, 3)
        Numdoc = Right(Trim(flxPagoDctos.TextMatrix(FilaSel, 3)), 8)
        
        td = IIf(TipoPago = "17", "01", Trim(flxDetDctos.TextMatrix(fila, 1)))
        aux = 2
        caux = Trim(flxPagoDctos.TextMatrix(FilaSel, 8))
        colv = ""
        cto = Trim(flxPagoDctos.TextMatrix(FilaSel, 2))
        Div = ValidaDameCCHFMTubulares(Trim(flxPagoDctos.TextMatrix(FilaSel, 9)))
        
        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Identif & "','" & _
              v & "','" & Trim(Serdoc) & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
              aux & "','" & caux & "','0000','00000000000','N','" & cto & "'," & _
              IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
              IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
              fecha & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
              colv & "','000','')"
        oConexionMYSQL.Execute (SQL)
        '**********************************************
    
Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error Generando Asientos"
    Resume
    Resume Next
End Sub


Function ValidaDameCCHFMTubulares(ByVal vdivi As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    ValidaDameCCHFMTubulares = ""
    SQL = " Select atipo from cnmdepar where coddep= '" & vdivi & "' and atipo='TUBULAR SERVICES' "
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        ValidaDameCCHFMTubulares = "013100003841"
    Else
        ValidaDameCCHFMTubulares = vdivi
    End If
    Set RQ = Nothing
End Function


Function vDameFechaFactura(ByVal vSerdoc As String, ByVal vNumdoc As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    vDameFechaFactura = Date
    
    SQL = "Select c.fecha from cnmovi as d left join cnvouc as c on (d.anomes=c.anomes) and (d.voucher=c.voucher) where d.serdoc='" & vSerdoc & "' and d.numdoc='" & vNumdoc & "' and c.codlib='04'  limit 1"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        vDameFechaFactura = RQ.Fields("fecha")
    End If
    Set RQ = Nothing
End Function

Function EncuentraArchivo(Nom As String, TxtID As String) As Boolean
EncuentraArchivo = False
    Dim SQL As String
    Dim RT As MYSQL_RS
    Dim RutaDestinoCliente As String
  
    RutaDestinoCliente = "\\172.26.35.12\ddsi$\Cobranzas\" & TxtID & ""
    SQL = "select * from archivosadjuntos where modulo = '11' and identificador = '" & TxtID & "' " & _
          "and ruta = '" & Replace(RutaDestinoCliente & "\", "\", "*") & "' and nombre = '" & Nom & "'"
    Set RT = oConexion.EjecutaSelectRS(SQL)
    
    If Not RT.EOF() Then
        EncuentraArchivo = True
    End If
    
    Set RT = Nothing
End Function
