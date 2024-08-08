VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmIngresarDocumento 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Documentos"
   ClientHeight    =   6570
   ClientLeft      =   3975
   ClientTop       =   5190
   ClientWidth     =   10410
   Icon            =   "frmIngresarDocuemnto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   1635
      Left            =   0
      TabIndex        =   7
      Top             =   -60
      Width           =   10395
      Begin VB.TextBox txtFolio 
         Alignment       =   2  'Center
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
         Left            =   6285
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "0"
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtDesDocIden 
         BackColor       =   &H00E0E0E0&
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
         Left            =   5895
         Locked          =   -1  'True
         TabIndex        =   29
         Tag             =   "90"
         Top             =   900
         Width           =   1965
      End
      Begin VB.TextBox TxtTipo 
         Height          =   315
         Left            =   840
         MaxLength       =   3
         TabIndex        =   1
         ToolTipText     =   "F1 para Busqueda"
         Top             =   180
         Width           =   645
      End
      Begin MSMask.MaskEdBox mskhora 
         Height          =   315
         Left            =   8670
         TabIndex        =   24
         Top             =   540
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm AM/PM"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpfecha 
         Height          =   315
         Left            =   8730
         TabIndex        =   23
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   122683393
         CurrentDate     =   38477
      End
      Begin VB.TextBox TxTipo_Ide_Doc 
         Height          =   315
         Left            =   5475
         MaxLength       =   2
         TabIndex        =   4
         ToolTipText     =   "F1 para Busqueda"
         Top             =   900
         Width           =   435
      End
      Begin VB.TextBox TxtObs 
         Height          =   315
         Left            =   810
         TabIndex        =   6
         Top             =   1260
         Width           =   3255
      End
      Begin VB.TextBox TxtNomEmisor 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   810
         TabIndex        =   3
         Top             =   900
         Width           =   3585
      End
      Begin VB.TextBox TxtNombreEmpresa 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   510
         Width           =   7065
      End
      Begin VB.TextBox TxtIdemisor 
         Height          =   315
         Left            =   8700
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "6"
         Top             =   900
         Width           =   1575
      End
      Begin VB.OptionButton OptRecepcion 
         BackColor       =   &H009F5539&
         Caption         =   "Recepción"
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
         Left            =   7980
         TabIndex        =   26
         Top             =   1305
         Width           =   1245
      End
      Begin VB.OptionButton OptRemitir 
         BackColor       =   &H009F5539&
         Caption         =   "Emisión"
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
         Left            =   9210
         TabIndex        =   27
         Top             =   1320
         Width           =   1035
      End
      Begin Proyecto1.chameleonButton BtnFoliosAut 
         Height          =   315
         Left            =   4080
         TabIndex        =   44
         ToolTipText     =   "Modificar"
         Top             =   1320
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "Folios Aut."
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
         MICON           =   "frmIngresarDocuemnto.frx":0442
         PICN            =   "frmIngresarDocuemnto.frx":045E
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
         Left            =   6270
         TabIndex        =   46
         Top             =   1230
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
         MICON           =   "frmIngresarDocuemnto.frx":088C
         PICN            =   "frmIngresarDocuemnto.frx":08A8
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
         Left            =   6660
         TabIndex        =   47
         Top             =   1200
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
         MICON           =   "frmIngresarDocuemnto.frx":0E42
         PICN            =   "frmIngresarDocuemnto.frx":0E5E
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
         Left            =   7860
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Documentos Excel"
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
         Left            =   5430
         TabIndex        =   48
         Top             =   1320
         Width           =   810
      End
      Begin MSForms.CheckBox chkcierre 
         Height          =   255
         Left            =   7110
         TabIndex        =   39
         Top             =   1290
         Width           =   825
         VariousPropertyBits=   746588179
         BackColor       =   10442041
         ForeColor       =   16777215
         DisplayStyle    =   4
         Size            =   "1455;450"
         Value           =   "0"
         Caption         =   "Cierre"
         PicturePosition =   131072
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblEstadoDoc 
         Alignment       =   2  'Center
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
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   5610
         TabIndex        =   38
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label9 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Folio:"
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
         Left            =   5475
         TabIndex        =   36
         Top             =   180
         Width           =   765
      End
      Begin VB.Label txtnombredoc 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         TabIndex        =   34
         Top             =   180
         Width           =   3915
      End
      Begin VB.Label Label8 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.:"
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
         Left            =   7980
         TabIndex        =   28
         Top             =   900
         Width           =   705
      End
      Begin VB.Label LBL 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7950
         TabIndex        =   25
         Top             =   1260
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Obs.:"
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
         Left            =   60
         TabIndex        =   22
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doc. Iden.:"
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
         Left            =   4395
         TabIndex        =   13
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emisor:"
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
         Left            =   60
         TabIndex        =   12
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora:"
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
         Left            =   7980
         TabIndex        =   11
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label4 
         BackColor       =   &H009F5539&
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
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   7980
         TabIndex        =   10
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label3 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sr (es):"
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
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo:"
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
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   5115
      Left            =   0
      TabIndex        =   14
      Top             =   1485
      Width           =   10395
      Begin Proyecto1.chameleonButton chameleonButton2 
         Height          =   345
         Left            =   5880
         TabIndex        =   43
         ToolTipText     =   "Modificar"
         Top             =   -240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Modificar"
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
         MICON           =   "frmIngresarDocuemnto.frx":13A0
         PICN            =   "frmIngresarDocuemnto.frx":13BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin NOVAdmin.flxEdit flxDetalles 
         Height          =   4455
         Left            =   0
         TabIndex        =   41
         Top             =   120
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   7858
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
         BackColorBkg    =   12632256
         CellPicture     =   "frmIngresarDocuemnto.frx":17EA
         ColAlignment0   =   9
         FixedAlignment0 =   9
         ForeColorSel    =   16711680
         ForeColorFixed  =   4194304
         GridColorFixed  =   8421504
         MouseIcon       =   "frmIngresarDocuemnto.frx":1806
         RowHeight0      =   240
      End
      Begin Proyecto1.chameleonButton btnReportes 
         Height          =   345
         Left            =   9375
         TabIndex        =   19
         ToolTipText     =   "Generar Reporte"
         Top             =   4620
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
         MICON           =   "frmIngresarDocuemnto.frx":1822
         PICN            =   "frmIngresarDocuemnto.frx":183E
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
         Left            =   9855
         TabIndex        =   20
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
         MICON           =   "frmIngresarDocuemnto.frx":1D80
         PICN            =   "frmIngresarDocuemnto.frx":1D9C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnCancelar 
         Height          =   345
         Left            =   3990
         TabIndex        =   17
         ToolTipText     =   "Deshacer"
         Top             =   4620
         Width           =   405
         _ExtentX        =   714
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
         MICON           =   "frmIngresarDocuemnto.frx":2162
         PICN            =   "frmIngresarDocuemnto.frx":217E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnNuevo 
         Height          =   345
         Left            =   75
         TabIndex        =   15
         ToolTipText     =   "Nuevo"
         Top             =   4620
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Nuevo"
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
         MICON           =   "frmIngresarDocuemnto.frx":26C0
         PICN            =   "frmIngresarDocuemnto.frx":26DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnEliminar 
         Height          =   345
         Left            =   2535
         TabIndex        =   16
         ToolTipText     =   "Eliminar"
         Top             =   4620
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Eliminar"
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
         MICON           =   "frmIngresarDocuemnto.frx":2A46
         PICN            =   "frmIngresarDocuemnto.frx":2A62
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnModificar 
         Height          =   345
         Left            =   1245
         TabIndex        =   21
         ToolTipText     =   "Modificar"
         Top             =   4620
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Modificar"
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
         MICON           =   "frmIngresarDocuemnto.frx":2EA4
         PICN            =   "frmIngresarDocuemnto.frx":2EC0
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
         Left            =   4470
         TabIndex        =   18
         ToolTipText     =   "Guardar"
         Top             =   4620
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
         MICON           =   "frmIngresarDocuemnto.frx":32EE
         PICN            =   "frmIngresarDocuemnto.frx":330A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdProg 
         Height          =   345
         Left            =   5580
         TabIndex        =   40
         ToolTipText     =   "Generar Reporte"
         Top             =   4620
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
         MICON           =   "frmIngresarDocuemnto.frx":374C
         PICN            =   "frmIngresarDocuemnto.frx":3768
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblGuardado 
         BackColor       =   &H009F5539&
         Caption         =   "Guardado con No Folio: xxxx"
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
         Left            =   6270
         TabIndex        =   30
         Top             =   4650
         Visible         =   0   'False
         Width           =   2925
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8220
      Top             =   4350
   End
   Begin Proyecto1.chameleonButton chameleonButton1 
      Height          =   345
      Left            =   9600
      TabIndex        =   42
      ToolTipText     =   "Modificar"
      Top             =   16080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Modificar"
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
      MICON           =   "frmIngresarDocuemnto.frx":3A82
      PICN            =   "frmIngresarDocuemnto.frx":3A9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblOrden 
      Caption         =   "Local"
      Height          =   165
      Left            =   360
      TabIndex        =   45
      Top             =   7200
      Width           =   1635
   End
   Begin VB.Label lblBusqueda 
      Caption         =   "Local"
      Height          =   165
      Left            =   7230
      TabIndex        =   37
      Top             =   6810
      Width           =   1635
   End
   Begin VB.Label lblEstado 
      Caption         =   "lblEstado"
      Height          =   225
      Left            =   5490
      TabIndex        =   35
      Top             =   6780
      Width           =   1395
   End
   Begin VB.Label lblModo 
      Caption         =   "Nuevo"
      Height          =   195
      Left            =   390
      TabIndex        =   33
      Top             =   6750
      Width           =   1425
   End
   Begin VB.Label lblTipoctamn 
      Height          =   225
      Left            =   3810
      TabIndex        =   32
      Top             =   6735
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblTipoctame 
      Height          =   225
      Left            =   2010
      TabIndex        =   31
      Top             =   6735
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmIngresarDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Col As Integer
Private m_NomCol(15) As String
Private m_AnchoCol(15) As Integer
Private SQL As String
Public familia As FAMILIA_DOC
Public Descuento As String
Public total As Double
Dim strvariable As String
Dim Rs As New MYSQL_RS
Public ContaFilas As Integer
Private oConsulta As FrmConsultas
Public Ident_OrdComp As String
Public dblImpEquiv As Double
Public ImpEquiv As Double
Dim Monto As Double
Dim MtoOtros As Double
Dim Veces As Integer

Dim Orden_fol As String
Dim Auxil_fol As String
Dim Codigo_fol As String
Dim Cen_fol As String
Dim Mon_fol As String
Dim dvcto_fol As String
Dim fecvcto_fol As String
Dim fecemis_fol As String
Dim tot_fol As Double
Dim divi_fol As String
 
Public Property Let pCol(valor As Integer)
    m_Col = valor
End Property

Public Property Let pCols(valor As Integer)
    m_Cols = valor
End Property

Public Property Let pNomCol(valor As String)
    m_NomCol(m_Col) = valor
End Property

Public Property Let pAnchoCol(valor As Integer)
    m_AnchoCol(m_Col) = valor
End Property

Sub LlenarDetalles()
    With flxDetalles
        .Clear
        .Cols = 3
        .Rows = 1
        .row = 0
        .Col = 0
        .CellForeColor = &H80000002
        .ColWidth(0) = 2000
        .ColType(0) = cadena
        .ColMaxLength(0) = 30
        .TextMatrix(0, 0) = Space(14) + "Detalle"
        
        .row = 0
        .Col = 1
        .CellForeColor = &H80000002
        .ColWidth(1) = 1999
        .ColType(1) = cadena
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.IZQUIERDA
        .CaracteresValidos(1) = "0123456789abcdefghijklmnñopqrstvwxyzABCDEFGHIJKLMNÑOPQRSTVWXYZ./"
        .TextMatrix(0, 1) = Space(15) + "Código"
        
        .row = 0
        .Col = 2
        .CellForeColor = &H80000002
        .ColWidth(2) = 5980
        .ColType(2) = cadena
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.IZQUIERDA
        .CaracteresValidos(2) = "0123456789abcdefghijklmnñopqrstvwxyzABCDEFGHIJKLMNÑOPQRSTVWXYZ./"
        .TextMatrix(0, 2) = Space(58) & "Descripción"
    End With
End Sub

Private Function GenerarCodigo() As String
    Dim Rs As MYSQL_RS
    Dim AnoMes As String
    Dim SQL As String
    AnoMes = strAnoSistema & strMesSistema
    SQL = "max_identificador where anomes = '" & AnoMes & "'"
    Set Rs = oConexion.EjecutaSelect(SQL)
    If Not Rs.EOF Then
        GenerarCodigo = Rs.Fields("anomes") & Right("0000" & Trim(str(val(Rs.Fields("maximo")) + 1)), 4)
    End If
    If Rs.RecordCount = 0 Then
        GenerarCodigo = AnoMes & "0001"
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Function

Private Function GenerarAuxiliarLegal(ByVal rucLegal As String) As String
    Dim Rs As MYSQL_RS
    Dim SQL As String
    
    SQL = "Call cn_carga_auxiliar_Legal ('" & rucLegal & "');"
    Set Rs = oConexion.EjecutaSelect(SQL)
    
    If Not Rs.EOF Then
        If Trim(flxDetalles.TextMatrix(8, 1)) = "N" Then
         GenerarAuxiliar = Rs.Fields("TipoCtaMonNac") & "-" & Rs.Fields("numcta_mn")
        Else
         GenerarAuxiliar = Rs.Fields("TipoCtaMonExt") & "-" & Rs.Fields("numcta_me")
        End If
    End If
    
    If Rs.RecordCount = 0 Then
        GenerarAuxiliar = "CHEQUE GERENCIA" & "-" & "00000000"
    End If
    
    Rs.CloseRecordset
    Set Rs = Nothing
End Function



Private Sub btnCancelar_Click()
    If lblBusqueda = "Local" Then
        ConfigurarBotones cfgCancelar
        chkcierre.Value = False
        txtFolio.SetFocus
    End If
    If lblBusqueda = "Foranea" Then
        ModoFormulario modEditar
    End If
End Sub

Private Sub btnEliminar_Click()
    Dim RES As Integer
    Dim ix As Integer
    RES = MsgBox("Esta seguro que desea Eliminar" & vbNewLine & vbNewLine & "El Documento con Folio Nro" & Right(strIdentificador, 4), vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        CambioEstado strIdentificador, ELIMINADO
        ModoFormulario modNuevo
        MsgBox "El Documento Ha sido Eliminado", vbInformation, "Aviso"
        Me.SetFocus
    End If
End Sub

Private Sub BtnFoliosAut_Click()
Dim FechaIni As String
Dim FechaFin As String

 If MsgBox("¿Desea generar los folios automaticos de todo el año?", vbQuestion + vbYesNo, gsNomSW) = vbYes Then
    FechaIni = InputBox("Ingrese la Fecha de Emisión de los folios  p.e:01/01/2017 ", "Folios Automaticos", "01/01/2017")
    If FechaIni = "" Then Exit Sub
    FechaFin = InputBox("Ingrese la Fecha de Vencimiento de los folios  p.e:01/12/2017 ", "Folios Automaticos", "01/01/2017")
    If FechaFin = "" Then Exit Sub
      GenerafoliosAutomaticos Orden_fol, FechaIni, FechaFin, Auxil_fol, Codigo_fol, Cen_fol, Mon_fol, dvcto_fol, fecvcto_fol, fecemis_fol, tot_fol, divi_fol
      
      MsgBox "Se generaron los folios correctamente, por favor proceda a verificar.", vbInformation, gsNomSW
 Else
   MsgBox "No se Ejecuto", vbInformation, gsNomSW
 End If

End Sub

Private Sub GenerafoliosAutomaticos(vOrden As String, FechaIni As String, FechaFin As String, Auxil_fol As String, Codigo_fol As String, Cen_fol As String, Mon_fol As String, dvcto_fol As String, fecvcto_fol, fecemis_fol As String, tot_fol As Double, divi_fol As String)
Dim strIdentificador As String
Dim vAnio As String, vMes As String, strConcepto As String, strFecha As String, strFechaE As String
Dim meImporte As Double
    strConcepto = "Folio Automatico generado"
    meImporte = CDbl(tot_fol) / 12
    
    'Folio 1
    vAnio = Right(FechaIni, 4)
    vMes = Mid(FechaIni, 4, 2)
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = FechaFin
    strFechaE = FechaIni
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
    
    
    'Folio 2
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol

    'Folio 3
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 4
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 5
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
      
    'Folio 6
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
      
    'Folio 7
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
      
    'Folio 8
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 9
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 10
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 11
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol
      
    'Folio 12
    If vMes = "12" Then
     vAnio = CStr(CInt(vAnio) + 1)
     
     vMes = "01"
    Else
     vMes = Right("0" & CStr(CInt(vMes) + 1), 2)
    End If
    strIdentificador = GenerarFolioAutom(vAnio, vMes)
    strFecha = Left(FechaFin, 2) & "/" & vMes & "/" & vAnio
    strFechaE = Left(FechaIni, 2) & "/" & vMes & "/" & vAnio
    EjecutaFolioAutomatico strIdentificador, vOrden, Auxil_fol, Codigo_fol, Mon_fol, strFecha, meImporte, strConcepto, divi_fol, strFechaE, Cen_fol

End Sub


Public Function GenerarFolioAutom(ByVal vAnio As String, ByVal vMes As String) As String
    Dim rsfolio As MYSQL_RS
    Dim AnoMes As String
    Dim SQL As String
    AnoMes = vAnio & vMes
    SQL = "max_identificador where anomes = '" & AnoMes & "'"
    Set rsfolio = oConexion.EjecutaSelect(SQL)
    If Not rsfolio.EOF Then
        GenerarFolioAutom = rsfolio.Fields("anomes") & Right("0000" & Trim(str(val(rsfolio.Fields("maximo")) + 1)), 4)
    End If
    If rsfolio.RecordCount = 0 Then
        GenerarFolioAutom = AnoMes & "0001"
    End If
    rsfolio.CloseRecordset
    Set rsfolio = Nothing
End Function


Private Sub EjecutaFolioAutomatico(strIdentificador As String, strOrden As String, strAuxiliar As String, strCodigo As String, strMoneda As String, strFecha As String, meImporte As Double, strConcepto As String, strDivi As String, strFechaE As String, strCenco As String)
Dim SQL As String
       SQL = "Call Insert_AmarreDoc ('" & strIdentificador & "','14'," & _
                " '" & Format(strFechaE, "yyyy/mm/dd") & "', '" & Format(Time, "HH:MM:SS") & "'," & _
                " '00','', 'NINGUNO', ''," & _
                " '','1', '" & strUsuarioId & "'," & _
                " '" & Left(strIdentificador, 6) & "','1'); "
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        
        SQL = "Call Insert_Contables ( '" & strIdentificador & "', " & _
              " '00000', " & _
              " 'PENDIENTE', " & _
              " '" & strOrden & "', " & _
              " '' ,  '" & strAuxiliar & "', " & _
              " '" & Trim(strCodigo) & "', '" & Trim(strCenco) & "', " & _
              " 'ENCARGA',  '" & strMoneda & "', " & _
              " 0,  '" & Trim(Format(strFecha, "yyyy/mm/dd")) & "', " & _
              " '" & Trim(Format(strFechaE, "yyyy/mm/dd")) & "', " & _
              " '" & Trim(Format(strFecha, "yyyy/mm/dd")) & "', " & _
              " 0.00,0.00,0.00," & CDbl(meImporte) & ",'', " & _
              " '" & strConcepto & "', " & _
              " '', " & _
              " '', " & _
              " '', " & _
              " '" & strDivi & "', " & _
              " 0.00, " & _
              " '',0.00,'',0.00,0)"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                 
        SQL = "Call Insert_Movi_Doc('" & strIdentificador & "', '" & Format(strFechaE, "yyyy/mm/dd") & "'," & _
                  " '" & EMITIDO & "','1','" & strUsuarioId & "'); "
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
         
        SQL = "Call Insert_HistorialDoc ( '" & strIdentificador & "', '" & EMITIDO & "', '" & DescripcionesdeCodigos("CNUSER", strUsuarioId, "area") & "'," & _
                  "'" & Format(Date, "yyyy/mm/dd") & "', '" & strUsuarioId & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        
End Sub

Private Sub btnGrabar_Click()
    Dim ORD As String
    Dim montofact As Double
    Dim ix As Integer
    Dim valorchk As Integer
    Dim FlgIng As Boolean
    Dim intCodigoOC As String
    Dim NewOrdenLegal As String
    Dim DatosBancLegal As String
    Dim TmpDatosBancLegal
    Dim FechaVoucher As String
    Dim AnoMes As String
    Dim lib As String
    Dim v As String
    Dim RY As New MYSQL_RS
    
    If TxtTipo.Text = "O" Then
       MsgBox "Usted no esta autorizado para dar ingreso a esta orden, solicitelo con Gerencia Administrativa", vbInformation, gsNomSW
    End If
    
    If ValidarData Then
        If lblModo = "Nuevo" Then
            AsignarValores
            If VerificaOrden(oDocumento.CodFamDoc) Then
                FlgIng = False
                If strAnoSistema > Year(Date) Then
                    FlgIng = True
                ElseIf strAnoSistema = Year(Date) Then
                    If val(strMesSistema) > val(Month(Date)) Then
                        'FlgIng = True
                    End If
                End If
                If FlgIng = False Then
                    If oDocumento.Guardar Then
'                        UpdateOrden TxtTipo.Text, 1
                    End If
                    If chkcierre.Value = True Then valochk = 1 Else valorchk = 0
                    oDocumento.ActCierre 1, valorchk
                    lblGuardado.Visible = True
                    If ConfirGuarda(oDocumento.Item(1).valor) Then
                        lblGuardado.ForeColor = &H8080FF
                        lblGuardado.Caption = "Guardado con N° Folio " & Right(oDocumento.Item(1).valor, 4)
                        strIdentificador = oDocumento.Item(1).valor
                        For ix = 1 To flxDetalles.Rows - 1
                            If Trim(flxDetalles.TextMatrix(ix, 0)) = "Total" Then Monto = CDbl(flxDetalles.TextMatrix(ix, 2))
                            If Trim(flxDetalles.TextMatrix(ix, 0)) = "Otros" Then MtoOtros = CDbl(flxDetalles.TextMatrix(ix, 2))
                        Next
                        
'                            If Trim(flxDetalles.TextMatrix(4, 1)) = "B" Then
'                                ' Proveedor Legal
'                                NewOrdenLegal = "B" & Right(Trim(flxDetalles.TextMatrix(5, 1)), 4) & Right(Trim(flxDetalles.TextMatrix(3, 2)), 4)
'                                intCodigoOC = GenerarCodigo
'                                DatosBancLegal = GenerarAuxiliarLegal(Trim(flxDetalles.TextMatrix(5, 1)))
'                                TmpDatosBancLegal = Split(Trim(DatosBancLegal), "-")
'
'                                SQL = "INSERT INTO `orden_compra` " & _
'                                      "VALUES ('" & intCodigoOC & "','00000','" & NewOrdenLegal & "','S','6','" & Trim(flxDetalles.TextMatrix(5, 1)) & "','2007/12/03','" & Trim(flxDetalles.TextMatrix(6, 1)) & "','FS','GVB','ORDENES IMPORTADAS DEL SISTEMA LOGISTICO'," & _
'                                      "'N','Autorizado'," & Trim(flxDetalles.TextMatrix(16, 2)) & "," & Trim(flxDetalles.TextMatrix(17, 2)) & ",0.000,0.000," & Trim(flxDetalles.TextMatrix(19, 2)) & "," & Trim(flxDetalles.TextMatrix(19, 2)) & ",7.000,0.000,'N'," & _
'                                      "'" & TmpDatosBancLegal(0) & "','BANCO PROV LEGAL','" & TmpDatosBancLegal(1) & "','0007',30,'N','N','0','0',NULL,NULL,NULL,NULL,NULL);"
'                                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
'
'
'                                SQL = "UPDATE DOCUMENTO_CONTABLES  SET ORDEN='" & NewOrdenLegal & "' WHERE IDENTIFICADOR='" & oDocumento.Item(1).valor & "' "
'                                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
'                                ' Fin Proveedor Legal
'                            End If
                            
                            
                            'Verificamos las ordenes que llegan las facturas
                            If TxtTipo.Text = "01" Then
                              'MsgBox "Entro a la Factura", vbInformation, gsNomSW
                              ' Trim (flxDetalles.TextMatrix(1, 1)) -- orden
                              VANOMESVOUCHERORDEN = ""
                              'SQL = "Call VerificaOrdenparaGeneraExtorno('" & Trim(flxDetalles.TextMatrix(1, 1)) & "'); "
                              SQL = "select concat(c.anomes,c.voucher) as validador from cnvouc as c left join cnmovi as d on (c.anomes=d.anomes) and (c.voucher=d.voucher) " & _
                                    "where c.glosa like '%PROVISION NRO ORDEN%' and d.numdoc='" & Trim(flxDetalles.TextMatrix(1, 1)) & "' limit 1"
                              'TxtNombreEmpresa.Text = SQL
                              Set RY = oConexion.EjecutaSelectRS(SQL)
                              
                              If Not RY.EOF() Then
                                   VANOMESVOUCHERORDEN = RY.Fields("validador")
                              End If
                                 
                              If VANOMESVOUCHERORDEN <> "" Then
                                'MsgBox "Entro a CREAR EL EXTORNO", vbInformation, gsNomSW
                                lib = "05"
                                FechaVoucher = Date
                                AnoMes = strAnoSistema & strMesSistema
                                v = MaxVoucher(AnoMes, lib)
                                glo = "EXTORNO ORDEN DE COMPRA: " & Trim(flxDetalles.TextMatrix(1, 1))
                                mon = "N"
                                TipoCambio (Trim(FechaVoucher))
                                tc = dblTipoCmbV
                                
                                
                                SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & FechaVoucher & _
                                      " ','" & FechaVoucher & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
                                      " ','CUADRADO','','','','','N','','')"
                                'TxtNombreEmpresa.Text = SQL
                                oConexionMYSQL.Execute (SQL)
                                
'                                SQL = "Call insert into `cnmovi`(`col_cv`,`cuenta`,`codlib`,`voucher`,`anomes`,`coddoc`,`serdoc`,`numdoc`,`correl`,`moneda`,`auxiliar`,`codaux`,`cencos`,`generada`,`concepto`,`cargos`,`abonos`,`cargod`,`abonod`,`lote`,`fecha`,`usuario`,`ruc`,`coddep`,`codfun`,`codmovi`,`cenco`,`codmor`) " & _
'                                      "select `col_cv`,`cuenta`,'" & lib & "','" & v & "','" & AnoMes & "',`coddoc`,`serdoc`,`numdoc`,`correl`,`moneda`,`auxiliar`,`codaux`,`cencos`,`generada`,`concepto`,`abonos`,`cargos`,`abonod`,`cargod`,`lote`,`fecha`,`usuario`,`ruc`,`coddep`,`codfun`,'H',`cenco`,`codmor` from cnmovi " & _
'                                      "where anomes='" & Left(VANOMESVOUCHERORDEN, 6) & "' and voucher='" & Right(VANOMESVOUCHERORDEN, 6) & "' and codmovi='D';"

                                 SQL = "Call GeneraVoucheExtornoNOVCONT ('" & lib & "','" & v & "','" & AnoMes & "','" & VANOMESVOUCHERORDEN & _
                                      " ','H','D'," & tc & ")"
                                oConexionMYSQL.Execute (SQL)
                                
'                                SQL = "Call insert into `cnmovi`(`col_cv`,`cuenta`,`codlib`,`voucher`,`anomes`,,`coddoc`,`serdoc`,`numdoc`,`correl`,`moneda`,`auxiliar`,`codaux`,`cencos`,`generada`,`concepto`,`cargos`,`abonos`,`cargod`,`abonod`,`lote`,`fecha`,`usuario`,`ruc`,`coddep`,`codfun`,`codmovi`,`cenco`,`codmor`) " & _
'                                      "select `col_cv`,`cuenta`,'" & lib & "','" & v & "','" & AnoMes & "',`coddoc`,`serdoc`,`numdoc`,`correl`,`moneda`,`auxiliar`,`codaux`,`cencos`,`generada`,`concepto`,`abonos`,`cargos`,`abonod`,`cargod`,`lote`,`fecha`,`usuario`,`ruc`,`coddep`,`codfun`,'D',`cenco`,`codmor` from cnmovi " & _
'                                      "where anomes='" & Left(VANOMESVOUCHERORDEN, 6) & "' and voucher='" & Right(VANOMESVOUCHERORDEN, 6) & "' and codmovi='H';"

                                SQL = "Call GeneraVoucheExtornoNOVCONT ('" & lib & "','" & v & "','" & AnoMes & "','" & VANOMESVOUCHERORDEN & _
                                      " ','D','H'," & tc & ")"
                                oConexionMYSQL.Execute (SQL)
                                
                                MsgBox "La Orden ya tiene provisión, ha sido extornada,revise NOVCONT", vbInformation, gsNomSW
                              End If
                              
                            
                            End If
                            'Fin Verificamos las ordenes que llegan las facturas
                            
                    
                    Else
                        oConexion.EjecutaInsertUpdateDelete "Delete from documento_contables where identificador='" & oDocumento.Item(1).valor & "'", TIPO_QUERY.Eliminar, False
                        oConexion.EjecutaInsertUpdateDelete "Delete from historial_docs where identificador='" & oDocumento.Item(1).valor & "'", TIPO_QUERY.Eliminar, False
                        oConexion.EjecutaInsertUpdateDelete "Delete from movi_documento where identificador='" & oDocumento.Item(1).valor & "'", TIPO_QUERY.Eliminar, False
                        oConexion.EjecutaInsertUpdateDelete "Delete from amarre_documento where identificador='" & oDocumento.Item(1).valor & "'", TIPO_QUERY.Eliminar, False
                        lblGuardado.ForeColor = &H80&
                        lblGuardado.Caption = "Error al momento de Guardar"
                    End If
                    ModoFormulario modConsulta
                Else
                    MsgBox "La fecha del Sistema es mayor a la fecha actual", vbInformation, gsNomSW
                    ModoFormulario modConsulta
                    ConfigurarBotones cfgEliminar
                End If
            Else
                MsgBox "La Orden ya se encuentra Registrada", vbInformation, gsNomSW
                'ModoFormulario modConsulta
                ModoFormulario modNuevo
                ConfigurarBotones cfgEliminar
            End If
        End If
        If lblModo = "Modificar" Then
            AsignarValores
            If oDocumento.Actualizar Then
'                UpdateOrden TxtTipo.Text, 2
            End If
            If chkcierre.Value = True Then valorchk = 1 Else valorchk = 0
            oDocumento.ActCierre 2, valorchk
            For ix = 1 To flxDetalles.Rows - 1
                If Trim(flxDetalles.TextMatrix(ix, 0)) = "Total" Then Monto = CDbl(flxDetalles.TextMatrix(ix, 2))
            Next
            ModoFormulario modConsulta
            If lblBusqueda = "Foranea" Then
                frmBusquedaDocumentaria.EjecutaBusqueda
            End If
            Me.SetFocus
        End If
        Publimensaje = "sineditar"
    End If
End Sub

Private Function VerificaOrden(familia As Integer) As Boolean
    Dim I As Integer
    If familia <> ORDENES Then
        VerificaOrden = True
        Exit Function
    End If
    If familia = ORDENES Then
        VerificaOrden = False
        For I = 1 To oDocumento.Count
            If oDocumento.Item(I).Nombre = "Correl" Then
                If (existeOrden(oDocumento.Item(I).valor)) Then
                    VerificaOrden = False
                    Exit Function
                Else
                    VerificaOrden = True
                End If
             End If
        Next
    End If
End Function

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    Monto = 0
    MtoOtros = 0
    ModoFormulario modNuevo
    TxtTipo.SetFocus
End Sub

Private Sub btnSalir_Click()
    Unload frmIngresarDocumento
End Sub

Private Sub btnReportes_Click()
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "REPORTE DE " & txtnombredoc
    Select Case familia
        Case FAMILIA_DOC.CONTABLES
            oReporte.Reporte = "Rep_Documentos.rpt"
            oReporte.sp_Rep_DocumentoCont strIdentificador, "documento_contables", familia
        Case FAMILIA_DOC.ORDENES
            oReporte.Reporte = "Rep_Documentos.rpt"
            oReporte.sp_Rep_DocumentoCont strIdentificador, "orden_compra", familia
        Case FAMILIA_DOC.ENTIDADES
            oReporte.Reporte = "Rep_Documentos.rpt"
            oReporte.sp_Rep_DocumentoCont strIdentificador, "documento_entidades", familia
        Case FAMILIA_DOC.GENERALES
            oReporte.Reporte = "Rep_Documentos.rpt"
            oReporte.sp_Rep_DocumentoCont strIdentificador, "documento_generales", familia
    End Select
End Sub

Private Sub cmdBuscar01_Click()
        Dim RutaOrigen As String
            Dim RutaDestinoCliente As String
            Dim NomArchivo As String, SQL As String
            Dim TxtCodigoID As String
                Cmd.CancelError = True
                Cmd.Flags = cdlOFNFileMustExist Or _
                        cdlOFNHideReadOnly Or _
                        cdlOFNExplorer Or _
                        cdlOFNLongNames
            On Error Resume Next
                Cmd.ShowOpen
                If err.Number = cdlCancel Or err.Number <> 0 Then Exit Sub
                RutaOrigen = Cmd.Filename
                If RutaOrigen = "" Then Exit Sub
            On Error GoTo ErrorAbrir
                RutaDestinoCliente = "\\172.26.35.12\ddsi$\Cobranzas"
                NomArchivo = Dir(RutaOrigen)
                TxtCodigoID = "I_" & Trim(strAnoSistema & strMesSistema & txtFolio)
                
                If NomArchivo <> "" Then
                    If EncuentraArchivo(NomArchivo, TxtCodigoID) = False Then
                        CopiarArchivos RutaOrigen, RutaDestinoCliente, TxtCodigoID
                        SQL = "call Insert_ArchAdjuntos('1','" & TxtCodigoID & "', " & _
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

Private Sub cmdProg_Click()
    frmProg.Show
End Sub

Private Sub cmdVer01_Click()
 Dim cDestino As String
 Dim SQL As String
 Dim RutaDestinoCliente As String
 Dim NombreArchivoVer As String
 
 txtObs.Text = strAnoSistema & strMesSistema & txtFolio.Text
 frmArchivosAdjuntos.AnioSel = strAnoSistema
 frmArchivosAdjuntos.MesSel = strMesSistema
 frmArchivosAdjuntos.IdentificadorAr = "I_" & Trim(strAnoSistema & strMesSistema & txtFolio.Text)
 frmArchivosAdjuntos.Show
End Sub

Private Sub flxDetalles_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NADA:
    With flxDetalles
        If KeyCode = 13 Then
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Descripcion = .TextMatrix(.row, 0) Then
                    If oDocumento.Item(I + 1).Descripcion = "Dias de Vencimiento" Then
                        'SendKeys "{F2}"
                        Call keybd_event(vbKeyF2, 0, 0, 0)
                    End If
                    If oDocumento.Item(I + 1).Descripcion = "Fecha de Emision" Then
                        .TextMatrix(.row + 1, 2) = Date
                        'SendKeys "{F2}"
                        Call keybd_event(vbKeyF2, 0, 0, 0)
                    End If
                End If
            Next
        End If
        If KeyCode = 112 And Publimensaje = "modificar" And flxDetalles.CellBackColor <> ColorCelda.Desabilitado Then
            Me.MousePointer = vbHourglass
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Descripcion = .TextMatrix(.row, 0) Then
                    If oDocumento.Item(I).Validacion = 1 Then
                        Select Case oDocumento.Item(I).TabladelCampo
                            Case "CENCOS"
                                With oConsulta
                                    .pCols = 3
                                    .pCol = 0: .pAnchoCol = 1500
                                    .pCol = 1: .pAnchoCol = 5000
                                    .pCol = 2: .pAnchoCol = 0
                                    .pTitulo = "Centros de Costo"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_CENCOS
                                    .Show
                                End With
                            Case "AUXILIARES"
                                With oConsulta
                                    .pCols = 2
                                    .pCol = 0: .pAnchoCol = 1500
                                    .pCol = 1: .pAnchoCol = 5000
                                    .pTitulo = "Tipos de Auxiliar"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_AUXILIARES
                                    .Show
                                End With
                               
                                
                            Case "TIPOS_AUXILIARES"
                                If .TextMatrix(.row, 0) = "Codigo del Empleado" Then
                                    strTipoAuxiliar = 3
                                End If
                                With oConsulta
                                    .pCols = 8
                                    .pCol = 0: .pAnchoCol = 1200
                                    .pCol = 1: .pAnchoCol = 4500
                                    .pTitulo = "Códigos de " & DescripcionesdeCodigos("CNAUXIL", strTipoAuxiliar)
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = Label_Descrip_Auxil
                                    .Show
                                End With
                               
                                 
                            Case "ORDENCOMPRA"
                                With oConsulta
                                    .pCols = 7
                                    .pCol = 0: .pAnchoCol = 0
                                    .pCol = 1: .pAnchoCol = 1100
                                    .pCol = 2: .pAnchoCol = 500
                                    .pCol = 3: .pAnchoCol = 1200
                                    .pCol = 4: .pAnchoCol = 1200
                                    .pCol = 5: .pAnchoCol = 800
                                    .pCol = 6: .pAnchoCol = 1000
                                    .pTitulo = "Consulta de Ordenes Compra/Servicio"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_ORDENCS
                                    .Show
                                End With
                            Case "DESCUENTOS"
                                With oConsulta
                                    .pCols = 3
                                    .pCol = 0: .pAnchoCol = 1000
                                    .pCol = 1: .pAnchoCol = 4000
                                    .pCol = 2: .pAnchoCol = 1000
                                    .pTitulo = "TIPOS DE DESCUENTOS"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_DESCUENTOS
                                    .Show
                                End With
                            
                            Case "TIPOPAGO"
                                With oConsulta
                                    .pCols = 2
                                    .pCol = 0: .pAnchoCol = 1200
                                    .pCol = 1: .pAnchoCol = 3000
                                    .pTitulo = "Consulta Tipos de Pago"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_TIPOPAG
                                    .Show
                                End With
                                
                            Case "ENCARGADOS"
                                With oConsulta
                                    .pCols = 2
                                    .pCol = 0: .pAnchoCol = 1200
                                    .pCol = 1: .pAnchoCol = 3000
                                    .pTitulo = "Consulta de Empleados"
                                    .pForm = FORM_ING_DOCUMENTO
                                    .pCaso = LABEL_ENCARGADOS
                                    .Show
                                End With
                            Case Else
                        End Select
                    End If
                End If
            Next
        End If
    End With
    Me.MousePointer = vbNormal
Exit Sub
NADA:
    Exit Sub
End Sub


Public Sub flxDetalles_KeyPress(KeyAscii As Integer)
    On Error GoTo NADA
    Dim I As Integer, J As Integer
    Dim rowaux As Integer
    Dim aux As String
    Dim SQL As String
    Dim rsaux As MYSQL_RS
    Dim k As Integer, orden As Boolean
    With flxDetalles
        If KeyAscii = 13 Then
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Descripcion = .TextMatrix(.row, 0) Then
                    If I <> oDocumento.Count Then
                        If oDocumento.Item(I + 1).Descripcion = "Dias de Vencimiento" Then
                            'SendKeys "{F2}"
                            Call keybd_event(vbKeyF2, 0, 0, 0)
                        End If
                    End If
                    If oDocumento.Item(I).Validacion = 1 Then
                        If oDocumento.Item(I).CompletaCero = 1 Then
                            .TextMatrix(.row, 1) = Space(2) & Right(Replace(Space(oDocumento.Item(I).Tamanio), " ", "0") & Trim(.TextMatrix(.row, 1)), oDocumento.Item(I).Tamanio)
                            Exit For
                        End If
                    Else
                        If oDocumento.Item(I).CompletaCero = 1 Then
                            .TextMatrix(.row, 2) = Space(10) & Right(Replace(Space(oDocumento.Item(I).Tamanio), " ", "0") & Trim(.TextMatrix(.row, 2)), oDocumento.Item(I).Tamanio)
                            Exit For
                        End If
                    End If
                End If
            Next
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Descripcion = .TextMatrix(.row, 0) Then
                    Select Case oDocumento.Item(I).Descripcion
                        Case "No. de Orden"
                            PressF1 = True
                            .Col = 2
                            .CellForeColor = &H80&
                            If CDbl(Trim((.TextMatrix(.row, 1)))) <> 0 Then LLenarOrdenCompra Trim((.TextMatrix(.row, 1))), , TxtTipo
                        Case "Codigo"
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(.row, 1)), strTipoAuxiliar, "Descrip")
                            aux = DescripcionesdeCodigos("CREDITO", Trim(.TextMatrix(.row, 1)), strTipoAuxiliar)
                            If (strTipoAuxiliar = 5) And (Trim(.TextMatrix(1, 0)) = "No. de Orden") Then
                                lblTipoctamn = ""
                                lblTipoctame = ""
                                SQL = "ordencompradatos where correl = '" & Trim(.TextMatrix(1, 1)) & "'"
                                Set rsaux = oConexion.EjecutaSelect(SQL)
                                If Not rsaux.EOF Then
                                    If Trim(rsaux.Fields("mon")) = "N" Then
                                        lblTipoctamn = TipoPago(Trim(rsaux.Fields("mpago")))
                                    Else
                                        lblTipoctame = TipoPago(Trim(rsaux.Fields("mpago")))
                                    End If
                                End If
                            Else
                                SQL = "Auxil where auxiliar = '" & strTipoAuxiliar & "' and codigo = '" & Trim(.TextMatrix(.row, 1)) & "'"
                                Set rsaux = oConexion.EjecutaSelect(SQL)
                                If Not rsaux.EOF Then
                                    lblTipoctame = rsaux.Fields("tipcta_me")
                                    lblTipoctamn = rsaux.Fields("tipcta_mn")
                                End If
                                If strTipoAuxiliar = 3 Then
                                    Dim CentroCosto As String
                                    SQL = "select TRIM(cencos) AS CENCOS from contrato where codemp = '" & Trim(.TextMatrix(.row, 1)) & "' and estado = 'AP'"
                                    Set rsaux = oConexion.EjecutaSelectRS(SQL)
                                    If Not rsaux.EOF() Then
                                        CentroCosto = Trim(rsaux.Fields("cencos"))
                                    End If
                                End If
                            End If
                            Set rsaux = Nothing
                            rowaux = .row
                            For k = 1 To .Rows - 1
                                If .TextMatrix(k, 0) = "Dias de Vencimiento" Then
                                    .TextMatrix(k, 1) = Space(2) & DescripcionesdeCodigos("FORMPAG1", aux, "2")
                                    .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("FORMPAG1", aux, "1")
                                    .row = k
                                    .Col = 2
                                    .CellForeColor = &H80&
                                End If
                                If .TextMatrix(k, 0) = "Centro de Costo" Then
                                    .TextMatrix(k, 1) = Space(2) & CentroCosto
                                    .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("CENCO", CentroCosto, "1")
                                    .row = k
                                    .Col = 2
                                    .CellForeColor = &H80&
                                End If
                            Next
                            .row = rowaux
                            .Col = 2
                            .CellForeColor = &H80&
                            Exit For
                        Case "Dias de Vencimiento"
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("FORMPAG", .TextMatrix(.row, 1), "2")
                            For k = 1 To .Rows - 1
                                Select Case .TextMatrix(k, 0)
                                    Case "Fecha de Vencimiento"
                                        .TextMatrix(k, 2) = Space(10) & CalcularFechaVcto(val(Trim(.TextMatrix(.row, 1))))
                                    Case "Fecha de Pago"
                                        If Trim(.TextMatrix(.row, 1)) = "0" Then
                                            .TextMatrix(k, 2) = Space(10) & CalcularFechaVcto(val(Trim(.TextMatrix(.row, 1))))
                                        Else
                                            .TextMatrix(k, 2) = Space(10) & Format(CalcularFechaPago(CalcularFechaVcto(val(Trim(.TextMatrix(.row, 1))))), "dd/mm/yyyy")
                                        End If
                                End Select
                            Next
                            Exit For
                        Case "Centro de Costo"
                            .TextMatrix(.row, 1) = Space(2) & UCase(Trim(.TextMatrix(.row, 1)))
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("CENCO", Trim(.TextMatrix(.row, 1)), "1")
                            Encargado Trim(.TextMatrix(.row, 1)), .row
                            .Col = 2
                            .CellForeColor = &H80&
                            Exit For
                        Case "Encargado"
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(.row, 1)), 3, "Descrip")
                            .Col = 2
                            .CellForeColor = &H80&
                            Exit For
                        Case "Moneda"
                            .TextMatrix(.row, 1) = Space(2) & UCase(Trim(.TextMatrix(.row, 1)))
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("MONEDA", Trim(.TextMatrix(.row, 1)))
                            rowaux = .row
                            If Trim(.TextMatrix(.row, 1)) = "N" Then
                                For k = 1 To .Rows - 1
                                    If lblTipoctamn <> Empty Then
                                        If .TextMatrix(k, 0) = "Tipo de Pago" Then
                                            .TextMatrix(k, 1) = Space(2) & lblTipoctamn
                                            .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(k, 1)))
                                            .row = k
                                            .Col = 2
                                            .CellForeColor = &H80&
                                        End If
                                    End If
                                    If Trim(.TextMatrix(k, 0)) = "Fecha de Pago" Then
                                        .TextMatrix(k, 2) = Space(10) & Format(CalcularFechaVcto(2), "dd/mm/yyyy")
                                    End If
                                Next
                                If dblTipoCmbV <> 0 Then
                                    dblImpEquiv = val(1 / dblTipoCmbV)
                                Else
                                    dblImpEquiv = 1
                                End If
                            Else
                                For k = 1 To .Rows - 1
                                    If lblTipoctame <> Empty Then
                                        If .TextMatrix(k, 0) = "Tipo de Pago" Then
                                            .TextMatrix(k, 1) = Space(2) & lblTipoctame
                                            .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(k, 1)))
                                            .row = k
                                            .Col = 2
                                            .CellForeColor = &H80&
                                        End If
                                    End If
                                     If Trim(.TextMatrix(k, 0)) = "Fecha de Pago" Then
                                        .TextMatrix(k, 2) = Space(10) & Format(CalcularFechaVcto(2), "dd/mm/yyyy")
                                     End If
                                Next
                                If dblTipoCmbV <> 0 Then
                                    dblImpEquiv = dblTipoCmbV
                                Else
                                    dblImpEquiv = 1
                                End If
                            End If
                            .row = rowaux
                            .Col = 2
                            .CellForeColor = &H80&
                            Exit For
                        Case "Tipo de Pago"
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(.row, 1)))
                            .Col = 2
                            .CellForeColor = &H80&
                        Case "Tipo-Orden"
                            If Trim(.TextMatrix(.row, 1)) = "C" Or Trim(.TextMatrix(.row, 1)) = "c" Then
                                .TextMatrix(.row, 2) = Space(10) & "Compra"
                            End If
                            If Trim(.TextMatrix(.row, 1)) = "S" Or Trim(.TextMatrix(.row, 1)) = "s" Then
                                .TextMatrix(.row, 2) = Space(10) & "Servicio"
                            End If
                                .Col = 2
                                .CellForeColor = &H80&
                            Exit For
                        Case "Auxiliar"
                            strTipoAuxiliar = Trim(.TextMatrix(.row, 1))
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("CNAUXIL", strTipoAuxiliar)
                            .Col = 2
                            .CellForeColor = &H80&
                            Exit For
                        Case "Sub-Total"
                            If Trim(.TextMatrix(.row, 2)) <> "" Then
                                orden = False
                                For k = 1 To .Rows - 1
                                    If Trim(.TextMatrix(k, 0)) = "No. de Orden" Then
                                        If Trim(.TextMatrix(k, 1)) <> "" Then
                                            orden = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If orden = True Then
                                    Veces = Veces + 1
                                    If Veces = 2 Then
                                        If VerificaMtoOrden(CDbl(IIf(val(.TextMatrix(.row, 2)) = 0, 0, .TextMatrix(.row, 2))), .row, "Sub-Total") Then
                                            CalcularTotal CDbl(.TextMatrix(.row, 2)), .row
                                            If Veces = 2 Then Veces = 0
                                        Else
                                            Veces = 0
                                            'MsgBox "No puede ingresar un Monto mayor al de la Orden", vbInformation, "NOVPeru"
                                        End If
                                    End If
                                Else
                                    CalcularTotal CDbl(.TextMatrix(.row, 2)), .row
                                End If
                            End If
                            Exit For
                        Case "Descuentos"
                            .TextMatrix(.row, 1) = Space(2) & Right("000" & Trim(.TextMatrix(.row, 1)), 3)
                            .Col = 2
                            .CellForeColor = &H80&
                            .TextMatrix(.row, 2) = Space(10) & DescripcionesdeCodigos("Otros_Descuentos", Trim(.TextMatrix(.row, 1)), "1") & _
                                                   " : " & DescripcionesdeCodigos("Otros_Descuentos", Trim(.TextMatrix(.row, 1)), "2") & "%"
                            rowaux = .row
                            If Trim(.TextMatrix(.row, 1)) <> "" Then
                                For J = 1 To .Rows - 1
                                    If .TextMatrix(J, 0) = "Sub-Total" Then
                                        CalcularTotal CDbl(Trim(.TextMatrix(J, 2))), rowaux
                                        Exit For
                                    End If
                                Next
                            End If
                            Exit For
                        Case "Otros"
                            rowaux = .row
                            orden = False
                            For k = 1 To .Rows - 1
                                If Trim(.TextMatrix(k, 0)) = "No. de Orden" Then
                                    If Trim(.TextMatrix(k, 1)) <> "" Then
                                        orden = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If Trim(.TextMatrix(.row, 2)) <> "" Then
                                If orden = True Then
                                    Veces = Veces + 1
                                    If Veces = 2 Then
                                    For J = 1 To .Rows - 1
                                        If .TextMatrix(J, 0) = "Sub-Total" Then
                                            If VerificaMtoOrden(CDbl(IIf(val(.TextMatrix(.row, 2)) = 0, 0, .TextMatrix(.row, 2))), .row, "Otros") Then
                                                CalcularTotal CDbl(Trim(.TextMatrix(J, 2))), rowaux
                                            Else
                                                MsgBox "No puede ingresar un Monto en OTROS que hacen que el Total sea mayor al de la Orden", vbInformation, "NOVPeru"
                                            End If
                                            Veces = 0
                                            Exit For
                                        End If
                                    Next
                                    End If
                                Else
                                    For J = 1 To .Rows - 1
                                        If .TextMatrix(J, 0) = "Sub-Total" Then
                                            CalcularTotal CDbl(Trim(.TextMatrix(J, 2))), rowaux
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                            Exit For
                        Case "Total"
                            orden = False
                            For k = 1 To .Rows - 1
                                If Trim(.TextMatrix(k, 0)) = "No. de Orden" Then
                                    If Trim(.TextMatrix(k, 1)) <> "" Then
                                        orden = True
                                        Exit For
                                    End If
                                End If

                                
                            Next
                            
                            
                            If orden = True Then
                                Veces = Veces + 1
                                If Veces = 2 Then
                                If VerificaMtoOrden(CDbl(IIf(val(.TextMatrix(.row, 2)) = 0, 0, .TextMatrix(.row, 2))), .row, "Total") Then
                                    If Veces = 2 Then
                                        For k = 1 To .Rows - 1
                                            If .TextMatrix(k, 0) = "Sub-Total" Then
                                                Exit For
                                            End If
                                        Next
                                        .TextMatrix(k, 2) = Space(10) & FormatNumber(CDbl(Trim(.TextMatrix(.row, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                                        CalcularTotal CDbl(Trim(.TextMatrix(.row, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), rowaux
                                        Veces = 0
                                    End If
                                Else
                                    Veces = 0
                                    MsgBox "No puede ingresar un Monto mayor al de la Orden", vbInformation, "NOVPeru"
                                End If
                                End If
                            End If
                            Exit For
                    End Select
                End If
            Next
        End If
    End With
Exit Sub
NADA:
    Veces = 0
    Exit Sub
End Sub

Function VerificaMtoOrden(Mto As Double, fila As Integer, Opc As String) As Boolean
    Dim SQL As String
    Dim RQ As MYSQL_RS
    Dim I As Integer
    Dim J As Integer
    Dim total As Double
    VerificaMtoOrden = False
    
    With flxDetalles
    
        SQL = "Select oc.total,((select sum(t.total) + sum(t.otros_montos) from (Select if(am.Cod_Tipo_Doc='07',a.total*(-1),a.total) as total, " & _
              " a.otros_montos from (documento_contables as a left join movi_documento as b  ON(a.identificador=b.identificador)) left join amarre_documento as am ON am.Identificador = a.Identificador " & _
              " where a.orden=oc.correl and  b.cod_estado<>'EL' and b.cod_estado<>'AN') as T) ) as" & _
              " monto_factura from orden_compra as oc where oc.correl='" & Trim(.TextMatrix(1, 1)) & "'"
        
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            total = CDbl(RQ.Fields("total") - RQ.Fields("monto_factura"))
            Select Case Opc
                Case "Sub-Total"
                    For I = 1 To .Rows - 1
                        If Trim(.TextMatrix(I, 0)) = "Otros" Then
                            Exit For
                        End If
                    Next
                    If CDbl(FormatNumber(CDbl(RQ.Fields("total") / DevIgv(strAnoSistema & strMesSistema, "1")), 2)) >= CDbl(FormatNumber(Abs(CDbl(RQ.Fields("monto_factura") / DevIgv(strAnoSistema & strMesSistema, "1")) - ((Monto - CDbl(.TextMatrix(I, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"))) + CDbl(FormatNumber(Mto, 2)), 2)) Then
                        VerificaMtoOrden = True
                    Else
                        If lblModo = "Nuevo" Then
                            .TextMatrix(fila, 2) = Space(10) & FormatNumber(CDbl(RQ.Fields("total") - RQ.Fields("monto_factura") - CDbl(.TextMatrix(I, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (CDbl(RQ.Fields("total") - RQ.Fields("monto_factura") - CDbl(.TextMatrix(I, 2)))) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        Else
                            .TextMatrix(fila, 2) = Space(10) & FormatNumber((Monto - CDbl(.TextMatrix(I, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (Monto - CDbl(.TextMatrix(I, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        End If
                    End If
                Case "Otros"
                    For I = 1 To .Rows - 1
                        If Trim(.TextMatrix(I, 0)) = "Total" Then
                            Exit For
                        End If
                    Next
                    If CDbl(RQ.Fields("total")) >= Abs(CDbl(RQ.Fields("monto_factura")) - Monto) + CDbl(FormatNumber(.TextMatrix(I, 2), 2)) + CDbl(FormatNumber(Mto, 2)) Then
                        VerificaMtoOrden = True
                    Else
                        .TextMatrix(fila, 2) = Space(10) & FormatNumber(MtoOtros, 2)
                        For I = 1 To .Rows - 1
                            If Trim(.TextMatrix(I, 0)) = "Sub-Total" Then
                                Exit For
                            End If
                        Next
                        If lblModo = "Nuevo" Then
                            .TextMatrix(I, 2) = Space(10) & FormatNumber((CDbl(RQ.Fields("total")) - CDbl(RQ.Fields("monto_factura")) - MtoOtros) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (CDbl(RQ.Fields("total")) - CDbl(RQ.Fields("monto_factura")) - MtoOtros) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        Else
                            .TextMatrix(I, 2) = Space(10) & FormatNumber((Monto - MtoOtros) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (Monto - MtoOtros) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        End If
                    End If
                Case "Total"
                    If CDbl(RQ.Fields("total")) >= Abs(CDbl(RQ.Fields("monto_factura")) - Monto) + CDbl(FormatNumber(Mto, 2)) Then
                        VerificaMtoOrden = True
                    Else
                        For I = 1 To .Rows - 1
                            If Trim(.TextMatrix(I, 0)) = "Sub-Total" Then
                                Exit For
                            End If
                        Next
                        For J = 1 To .Rows - 1
                            If Trim(.TextMatrix(J, 0)) = "Otros" Then
                                Exit For
                            End If
                        Next
                        If lblModo = "Nuevo" Then
                            .TextMatrix(I, 2) = Space(10) & FormatNumber((CDbl(RQ.Fields("total")) - CDbl(RQ.Fields("monto_factura")) - CDbl(.TextMatrix(J, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (CDbl(RQ.Fields("total")) - CDbl(RQ.Fields("monto_factura")) - CDbl(.TextMatrix(J, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        Else
                            .TextMatrix(I, 2) = Space(10) & FormatNumber((Monto - CDbl(.TextMatrix(J, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), 2)
                            CalcularTotal (Monto - CDbl(.TextMatrix(J, 2))) / DevIgv(strAnoSistema & strMesSistema, "1"), fila
                        End If
                    End If
            End Select
        End If
    End With
    Set RQ = Nothing
End Function

Public Sub CalcularTotal(SubTotal As Double, rowanterior As Integer, Optional Dcto As Double, Optional Valor2 As Double)
    Dim I As Integer, J As Integer, row As Integer
    Dim Igv As Double, Otros As Double, total As Double
    Igv = 0
    Otros = 0
    With flxDetalles
        For I = 1 To .Rows - 1
              If .TextMatrix(I, 0) = "I.G.V." Then
                Igv = SubTotal * DevIgv(strAnoSistema & strMesSistema, "0")
                .TextMatrix(I, 2) = Space(10) & FormatNumber(Igv, 2)
            End If
        Next
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 0) = "Descuentos" Then
                For J = 1 To .Rows - 1
                    If .TextMatrix(J, 0) = "Otros" Then
                        If val(Trim(.TextMatrix(I, 1))) = 0 Then
                            Otros = CDbl(IIf(Trim(.TextMatrix(J, 2)) = "", 0, Trim(.TextMatrix(J, 2))))
                            .TextMatrix(J, 2) = Space(10) & FormatNumber((CDbl(Otros)), 2)
                        Else
                            .TextMatrix(J, 2) = Space(10) & FormatNumber(str(-0.01 * (SubTotal + Igv) * CDbl(DescripcionesdeCodigos("Otros_Descuentos", Trim(.TextMatrix(I, 1)), "2"))), 2)
                            Otros = CDbl(Trim(.TextMatrix(J, 2)))
                        End If
                    End If
                Next
            End If
        Next
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 0) = "Total" Then
                total = SubTotal + Igv + Otros
                .TextMatrix(I, 2) = Space(10) & FormatNumber(total, 2)
            End If
        Next
        ImpEquiv = FormatNumber(total, 2) * dblImpEquiv
        .row = rowanterior
    End With
End Sub

Public Sub CalcularSubTotal(valor As Double, rowanterior As Integer)
      Dim I As Integer
    Dim val As Double
    val = 0
    With flxDetalles
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 0) = "Sub-Total" Then
                val = valor / (1 + DevIgv(strAnoSistema & strMesSistema, "0"))
                .TextMatrix(I, 2) = Space(10) & FormatNumber(val, 2)
            End If
            If .TextMatrix(I, 0) = "I.G.V." Then
                val = val * DevIgv(strAnoSistema & strMesSistema, "0")
                .TextMatrix(I, 2) = Space(10) & FormatNumber(val, 2)
            End If
        Next I
        .row = rowanterior
    End With
End Sub

Public Function CalcularFechaVcto(valor As Integer) As String
    Dim I As Integer
    Dim val As Date
    val = CDate(DtpFecha.Value)
    val = val + valor
    CalcularFechaVcto = Format(CStr(val), "dd/mm/yyyy")
End Function

Public Sub Encargado(valor As String, rowanterior As Integer)
    Dim I As Integer
    Dim rsEncargado As MYSQL_RS
    Dim SQL As String
    SQL = "Select cod_auxil from cncosto where TRIM(cenco)='" & UCase(Trim(valor)) & "'"
    Set rsEncargado = oConexion.EjecutaSelectRS(SQL)
    If rsEncargado.RecordCount > 0 Then
        With flxDetalles
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) = "Encargado" Then
                    .TextMatrix(I, 1) = Space(2) & CE(rsEncargado.Fields("Cod_Auxil"))
                    .TextMatrix(I, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", CE(rsEncargado.Fields("Cod_auxil")), 3, "Descrip")
                    .row = I
                    .Col = 2
                    .CellForeColor = &H80&
                    Exit For
                End If
            Next I
            .row = rowanterior
        End With
    Else
        Set rsEncargado = Nothing
        Exit Sub
    End If
End Sub

Private Sub flxDetalles_RowColChange()
    Dim I As Integer
    With flxDetalles
        If .row > 0 Then
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Descripcion = .TextMatrix(.row, 0) Then
                    If .ControlVisible = False Then
                        TipodeCampo = oDocumento.Item(I).Tipo
                        .ColType(.Col) = TipodeCampo
                        .ColMaxLength(.Col) = oDocumento.Item(I).Tamanio
                        .CaracteresValidos(.Col) = oDocumento.Item(I).CaractValidos
                        .ColDecimales(.Col) = oDocumento.Item(I).Presicion
                    End If
                    Exit For
                End If
            Next
        Else
            TipodeCampo = cadena
            .ColType(.Col) = cadena
        End If
    End With
End Sub

Private Sub Form_Activate()
    Me.MousePointer = vbNormal
    Select Case oConsulta.pCaso
        Case LABEL_TIP_DOC
            TxtNombreEmpresa.SetFocus
        Case LABEL_DOC_IDEM
            TxtIdemisor.SetFocus
        Case LABEL_CENCOS, LABEL_PROVEEDORES, Label_Descrip_Auxil, LABEL_ORDENCS, LABEL_ENCARGADOS, LABEL_AUXILIARES
            flxDetalles.SiguienteCelda
    End Select
    If txtFolio <> Empty And Len(Trim(txtFolio)) = 4 Then
        strIdentificador = strAnoSistema & strMesSistema & txtFolio
    End If
End Sub

Private Sub Form_Load()
    Call WheelHook(frmIngresarDocumento)
    Set oConsulta = New FrmConsultas
    Me.Top = 0
    Me.Left = 0
    total = 0
    DtpFecha.Value = Date
    mskhora.Text = Format(Time, "HH:MM")
    flxDetalles.Clear
    OptRecepcion.Value = True
    ModoFormulario modAccion
    PressF1 = False
    'Si es usuario Supervisor
    DtpFecha.Enabled = False
    'Fin Si
    
End Sub

Public Sub LlenarGrilla(grilla As flxEdit)
    Dim I As Integer
    Dim J As Integer
    J = 1
    With oDocumento
        For I = 1 To .Count
            If .Item(I).Visible = 1 Then
                grilla.Rows = grilla.Rows + 1
                grilla.TextMatrix(J, 0) = .Item(I).Descripcion
                grilla.row = J
                Select Case .Item(I).Validacion
                    Case 0
                        grilla.Col = 1
                        grilla.CellBackColor = ColorCelda.Desabilitado
                    Case 1
                        grilla.Col = 2
                        grilla.CellBackColor = ColorCelda.Desabilitado
                    Case 2
                        grilla.Col = 1
                        grilla.CellBackColor = ColorCelda.Desabilitado
                        grilla.Col = 2
                        grilla.CellBackColor = ColorCelda.Desabilitado
                End Select
                J = J + 1
            End If
        Next I
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    btnSalir_Click
    Set oDocumento = Nothing
    Set oConsulta = Nothing
    total = 0
    WheelUnHook
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And lblModo = "Acción" Then
        If Len(Trim(txtFolio)) > 5 Then
            strIdentificador = txtFolio
        Else
            strIdentificador = strAnoSistema & strMesSistema & Right("0000" & Trim(txtFolio), 4)
            txtFolio = Right("0000" & txtFolio, 4)
        End If
        lblBusqueda = "Local"
        If BuscaLocal(strIdentificador) = True Then
            ModoFormulario modConsulta
            If lblEstadoDoc = "ANULADO" Or lblEstadoDoc = "ELIMINADO" Or lblEstadoDoc = "CANCELADO" Then
                btnGrabar.Enabled = False: BtnModificar.Enabled = False: BtnEliminar.Enabled = False: 'cmdProg.Enabled = False
            Else
                BtnModificar.Enabled = True: BtnEliminar.Enabled = True: cmdProg.Enabled = True
            End If
        Else
            mark1 txtFolio
        End If
    End If
End Sub

Private Function BuscaLocal(vIdentificador As String) As Boolean
    Dim SQL As String
    Dim Rs As MYSQL_RS
    SQL = "Select * from amarre_documento where identificador = '" & vIdentificador & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    BuscaLocal = False
    If Rs.EOF And Rs.BOF Then
        MsgBox "No existe documento para el folio", vbInformation, "NOVPeru"
        Exit Function
    Else
        Dim RQ As MYSQL_RS
        SQL = "select DESCRIP, Cod_Fam from cndocum C where CODDOC='" & Rs.Fields("cod_tipo_doc") & "' " & _
              "AND (PROTEGIDO = 'N' OR (SELECT PERMISO FROM `docsusuario` D WHERE D.CODDOC=C.CODDOC " & _
              "AND USUARIO = '" & strUsuarioId & "')=1)"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If RQ.EOF() Then
            MsgBox "No se encuentra autorizado para visualizar este folio", vbInformation, "NOVPeru"
            Exit Function
        End If
        Set RQ = Nothing
    End If
   
    BuscaLocal = True
    DtpFecha.Value = Rs.Fields("Fecha_registro")
    mskhora.Text = Format(IIf(Right(Rs.Fields("Hora_registro"), 1) = ":", Left(Trim(Rs.Fields("Hora_registro")), 1), Trim(Rs.Fields("Hora_registro"))), "hh:mm")
    TxTipo_Ide_Doc.Text = Rs.Fields("Tipo_Doc_Ide")
    TxtIdemisor.Text = Rs.Fields("Ide_Mensajero")
    TxtNomEmisor.Text = Rs.Fields("Nombre_Mensajero")
    TxtNombreEmpresa.Text = Rs.Fields("Empresa")
    txtObs.Text = Rs.Fields("Obs")
    FolioAuto = Trim(Rs.Fields("Obs"))
    If Rs.Fields("Flag") = 1 Then OptRecepcion.Value = True
    If Rs.Fields("Flag") = 0 Then OptRemitir.Value = True
    TxtTipo = Rs.Fields("Cod_Tipo_Doc")
    familia = Rs.Fields("Cod_Fam")
    txtnombredoc = DescripcionesdeCodigos("CNDOCUM", Trim(TxtTipo), "1")
    Set Rs = Nothing
    SQL = "Select DESCRIPCION FROM doc_identificacion where TIPO_DOC_IDE = '" & TxTipo_Ide_Doc.Text & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    txtDesDocIden.Text = Rs.Fields("DESCRIPCION")
    Set Rs = Nothing
    Select Case familia
        Case FAMILIA_DOC.CONTABLES
            SQL = "Select * from Documento_contables where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.ORDENES
            SQL = " Select * from ORDEN_Compra where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.GENERALES
            SQL = " Select * from documento_generales where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.ENTIDADES
            SQL = " Select * from documento_entidades where Identificador = '" & vIdentificador & "'"
    End Select
    Set rsfolio = Nothing
    lblEstadoDoc = ""
    lblEstadoDoc = DescripcionesdeCodigos("DOC_ESTADO", EstadoDoc(vIdentificador))
    
    ConfGridDoc
    CargarInfoRegistrada SQL
End Function

Private Sub TxtIdemisor_GotFocus()
    mark TxtIdemisor
End Sub

Private Sub TxtIdemisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtObs.SetFocus
    End If
End Sub

Private Sub TxTipo_Ide_Doc_GotFocus()
    mark TxTipo_Ide_Doc
End Sub

Private Sub TxTipo_Ide_Doc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And TxTipo_Ide_Doc.BackColor <> ColorDeshabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 1000
            .pCol = 1: .pAnchoCol = 3500
            .pTitulo = "Tipos de Doc. de Identidad"
            .pForm = FORM_ING_DOCUMENTO
            .pCaso = LABEL_DOC_IDEM
            .Show
        End With
    End If
End Sub

Private Sub TxTipo_Ide_Doc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim SQL As String
        Dim encontro As Boolean
        encontro = False
        Dim Rs As New MYSQL_RS
        TxTipo_Ide_Doc = Right("00" & Trim(TxTipo_Ide_Doc), 2)
        Set Rs = oConexion.EjecutaSelect("DOC_IDEM")
        Do While Not (Rs.EOF)
            If Trim(TxTipo_Ide_Doc) = Rs.Fields(0) Then
                txtDesDocIden = Rs.Fields(1)
                TxtIdemisor.SetFocus
                encontro = True
            End If
            Rs.MoveNext
        Loop
        If encontro = False Then
            MsgBox "No se encuentra el código ingresado", vbInformation, gsNomSW
            TxTipo_Ide_Doc = Empty
            TxTipo_Ide_Doc.SetFocus
            Exit Sub
        End If
    End If
    Set Rs = Nothing
End Sub

Private Sub TxtNombreEmpresa_GotFocus()
    mark TxtNombreEmpresa
End Sub

Private Sub TxtNombreEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtNomEmisor.SetFocus
    End If
End Sub

Private Sub TxtNomEmisor_GotFocus()
    mark TxtNomEmisor
End Sub

Private Sub TxtNomEmisor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxTipo_Ide_Doc.SetFocus
    End If
End Sub

Private Sub TxtObs_GotFocus()
    mark txtObs
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxDetalles.SetFocus
    End If
End Sub

Private Sub txtTipo_Change()
    If TxtTipo = Empty Then
        txtnombredoc = Empty
    End If
End Sub

Private Sub txtTipo_GotFocus()
    mark TxtTipo
End Sub

Private Sub TxtTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    Set oDocumento = New clsDocumento
    If KeyCode = vbKeyF1 And TxtTipo.BackColor <> ColorDeshabilitado Then
        With oConsulta
            .pCols = 4
            .pCol = 0: .pAnchoCol = 800
            .pCol = 1: .pAnchoCol = 800
            .pCol = 2: .pAnchoCol = 3000
            .pCol = 3: .pAnchoCol = 0
            .pTitulo = "Tipos de Documento"
            .pForm = FORM_ING_DOCUMENTO
            .pCaso = LABEL_TIP_DOC
            .Show
        End With
    End If
End Sub

Public Sub TxtTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtTipo = UCase$(TxtTipo)
        Set oDocumento = New clsDocumento
        lblGuardado.Visible = False
        If oDocumento.ObtenerNombreDocumento(TxtTipo) <> Empty Then
            familia = oDocumento.CodFamDoc
            If familia = ORDENES Then
                OptRemitir.Value = True
                OptRecepcion.Value = False
            Else
                OptRecepcion.Value = True
                OptRemitir.Value = False
            End If
            txtnombredoc = oDocumento.ObtenerNombreDocumento(TxtTipo)
            If lblModo = "Nuevo" Or lblModo = "Modificar" Then
                Dim SQL As String
                Dim Rs As New MYSQL_RS
                SQL = "Configura_grillaV2 where tipo_doc = '" & TxtTipo & "' "
                Set Rs = oConexion.EjecutaSelect(SQL)
                If Rs.RecordCount <> 0 Then
                    oDocumento.TipoDoc = TxtTipo
                    oDocumento.Rscampos = Rs
                    oDocumento.ConfigCampos
                    LlenarDetalles
                    LlenarGrilla flxDetalles
                    PosicionarCelda
                    TxtNombreEmpresa.SetFocus
                Else
                    Set Rs = Nothing
                    If MsgBox("Debe configurar los campos del Documento " & txtnombredoc & " Desea Configurarlo (S/N)", vbInformation + vbYesNo, m_Titulo) = vbYes Then
                        With frmConfigurarDoc
                            .txtTipoDoc = TxtTipo.Text
                            .txtTipoDoc_KeyPress (13)
                            .Show
                            Exit Sub
                        End With
                    Else
                        Exit Sub
                    End If
                End If
                Set Rs = Nothing
            End If
            If Trim(TxtTipo) = "01" Then chkcierre.Locked = False Else chkcierre.Locked = True
            If lblModo = "Acción" Then txtFolio.SetFocus
            If lblModo = "Nuevo" Then TxtNombreEmpresa.SetFocus
        Else
            MsgBox "No se encuentra el codigo del documento", vbInformation, gsNomSW
            TxtTipo = Empty
            TxtTipo.SetFocus
        End If
    End If
    Publimensaje = "modificar"
End Sub

Private Sub ConfGridDoc()
    Dim SQL As String
    Dim Rs As New MYSQL_RS
    Set oDocumento = New clsDocumento
    SQL = "Configura_grillaV2 where tipo_doc = '" & TxtTipo & "' "
    Set Rs = oConexion.EjecutaSelect(SQL)
    If Rs.RecordCount <> 0 Then
        txtnombredoc = oDocumento.ObtenerNombreDocumento(TxtTipo)
        oDocumento.TipoDoc = TxtTipo
        oDocumento.Rscampos = Rs
        oDocumento.ConfigCampos
        LlenarDetalles
        LlenarGrilla flxDetalles
        PosicionarCelda
        TxtNombreEmpresa.SetFocus
    Else
        Set Rs = Nothing
        If MsgBox("Debe configurar los campos del Documento " & txtnombredoc & " Desea Configurarlo (S/N)", vbInformation + vbYesNo, m_Titulo) = vbYes Then
            With frmConfigurarDoc
                .txtTipoDoc = TxtTipo.Text
                .txtTipoDoc_KeyPress (13)
                .Show
                Exit Sub
            End With
        Else
            Exit Sub
        End If
    End If
    Set Rs = Nothing
End Sub

Public Sub PosicionarCelda()
    flxDetalles.row = 1
    flxDetalles.Col = 1
    flxDetalles.Col = IIf(flxDetalles.CellBackColor = ColorCelda.Desabilitado, flxDetalles.Col + 1, flxDetalles.Col)
End Sub

Public Sub AsignarValores()
    Dim J As Integer
    Dim valor As String
    Dim Division As String
    
    With oDocumento
        J = 1
        .Item(1).valor = GenerarFolio
        For I = 2 To .Count
            If .Item(I).Visible = 1 Then
                If .Item(I).Descripcion = flxDetalles.TextMatrix(J, 0) Then
                    flxDetalles.row = J
                    Select Case .Item(I).Validacion
                        Case 0: flxDetalles.Col = 2
                        Case 1: flxDetalles.Col = 1
                        Case 2: flxDetalles.Col = 2
                    End Select
                    valor = Trim(flxDetalles.TextMatrix(J, flxDetalles.Col))
                    Select Case .Item(I).Tipo
                        Case flextype.cadena: .Item(I).valor = IIf(valor = "", " ", valor)
                        Case flextype.Entero: .Item(I).valor = IIf(valor = "", 0, CDbl(valor))
                        Case flextype.fecha: .Item(I).valor = IIf(valor = "", "", Format(valor, "yyyy/mm/dd"))
                        Case flextype.Numero
                            If valor = "" Then valor = "0"
                            .Item(I).valor = CDbl(valor)
                    End Select
                    If .Item(I).Nombre = "Cenco" Then
                        Division = DescripcionesdeCodigos("CENCO", valor, "2")
                    End If
                    J = J + 1
                End If
            Else
                Select Case .Item(I).Tipo
                    Case flextype.cadena: .Item(I).valor = " "
                    Case flextype.Entero: .Item(I).valor = 0
                    Case flextype.fecha: .Item(I).valor = Format(Date, "yyyy/mm/dd")
                    Case flextype.Numero: .Item(I).valor = 0
                End Select
                If .Item(I).Nombre = "Division" Then 'CAMBIO CON LA NUEVA ESTRUCTURA HFM
                    '.Item(I).valor = CE(Division)
                    .Item(I).valor = GenerarDivHFM(CE(Division), .Item(8).valor)
                End If
                If .Item(I).Nombre = "ImpEqui" Then
                    .Item(I).valor = ImpEquiv
                End If
                If .Item(I).Nombre = "Serie" Then
                    .Item(I).valor = "00000"
                End If
            End If
        Next I
        .cl_strNomEmpresa = TxtNombreEmpresa.Text
        .cl_intTipoIdentMsjro = TxTipo_Ide_Doc.Text
        .cl_strNumIdentMsjro = TxtIdemisor.Text
        .cl_strNombreMsjro = TxtNomEmisor.Text
        .cl_strObs = txtObs.Text
        .HoraReg = mskhora.Text
        .FechaReg = Format(DtpFecha, "yyyy/mm/dd")
        .cl_strAnomes = strAnoSistema & strMesSistema
        If OptRecepcion.Value = True Then
            .cl_strFlag = 1
            .EstadoDoc = REGISTRADO
        Else
            .cl_strFlag = 0
            .EstadoDoc = EMITIDO
        End If
    End With
    
    
    
End Sub

Public Sub BuscayPone(ByVal vIdentificador As String)
    Dim Rs As New MYSQL_RS
    Dim TipoDoc As String
    SQL = "Select * from amarre_documento where identificador = '" & vIdentificador & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    DtpFecha.Value = Rs.Fields("Fecha_registro")
    mskhora.Text = Format(Rs.Fields("Hora_registro"), "hh:mm")
    TxTipo_Ide_Doc.Text = Rs.Fields("Tipo_Doc_Ide")
    TxtIdemisor.Text = Rs.Fields("Ide_Mensajero")
    TxtNomEmisor.Text = Rs.Fields("Nombre_Mensajero")
    TxtNombreEmpresa.Text = Rs.Fields("Empresa")
    txtObs.Text = Rs.Fields("Obs")
    If Rs.Fields("Flag") = 1 Then OptRecepcion.Value = True
    If Rs.Fields("Flag") = 0 Then OptRemitir.Value = True
    familia = Rs.Fields("Cod_Fam")
    TipoDoc = CE(Rs.Fields("Cod_Tipo_Doc"))
    Set Rs = Nothing
    SQL = "Select DESCRIPCION FROM doc_identificacion where TIPO_DOC_IDE = '" & TxTipo_Ide_Doc.Text & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    txtDesDocIden.Text = Rs.Fields("DESCRIPCION")
    Set Rs = Nothing
    Select Case familia
        Case FAMILIA_DOC.CONTABLES
        SQL = "Select * from Documento_contables where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.ORDENES
        SQL = " Select * from ORDEN_Compra where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.GENERALES
        SQL = " Select * from documento_generales where Identificador = '" & vIdentificador & "'"
        Case FAMILIA_DOC.ENTIDADES
        SQL = " Select * from documento_entidades where Identificador = '" & vIdentificador & "'"
    End Select
    If lblBusqueda = "Foranea" Then
        CargarInfoRegistrada (SQL)
    End If
    If lblBusqueda = "Local" Then
        ConfGridDoc
        CargarInfoRegistrada (SQL)
    End If
    Publimensaje = "sineditar"
End Sub

Private Function CargarInfoRegistrada(ByVal SQL As String)
    Dim Rs As New MYSQL_RS
    Dim I, J As Integer
    J = 1
    On Error GoTo errorcarga
    With flxDetalles
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        For I = 1 To oDocumento.Count
            If oDocumento.Item(I).Visible = 1 Then
                If oDocumento.Item(I).Validacion = 1 Then
                    .Col = 2
                    .row = J
                    .CellForeColor = &H80&
                    .TextMatrix(J, 1) = Space(2) & Rs.Fields(oDocumento.Item(I).Nombre)
                    Select Case oDocumento.Item(I).Descripcion
                        Case "Auxiliar": .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("CNAUXIL", Trim(.TextMatrix(J, 1))): strTipoAuxiliar = Trim(.TextMatrix(J, 1))
                        Case "Centro de Costo": .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("CENCO", Trim(.TextMatrix(J, 1)), "1")
                        Case "Codigo"
                            .TextMatrix(J, 1) = Space(2) & Right("00000000000" & Trim(.TextMatrix(J, 1)), 11)
                            .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(J, 1)), strTipoAuxiliar, "Descrip")
                        Case "Encargado": .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(J, 1)), 3, "Descrip")
                        Case "Tipo de Pago": .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(J, 1)))
                        Case "Moneda"
                            If Rs.Fields(oDocumento.Item(I).Nombre) = "N" Then
                                .TextMatrix(J, 1) = Space(2) & "N"
                                .TextMatrix(J, 2) = Space(10) & "Nacional"
                                If dblTipoCmbV <> 0 Then
                                    dblImpEquiv = val(1 / dblTipoCmbV)
                                Else
                                    dblImpEquiv = 1
                                End If
                            Else
                                .TextMatrix(J, 1) = Space(2) & "E"
                                .TextMatrix(J, 2) = Space(10) & "Extranjera"
                                If dblTipoCmbV <> 0 Then
                                    dblImpEquiv = dblTipoCmbV
                                Else
                                    dblImpEquiv = 1
                                End If
                            End If
                        Case "Tipo-Orden"
                            If Rs.Fields(oDocumento.Item(I).Nombre) = "C" Or Rs.Fields(oDocumento.Item(I).Nombre) = "c" Then
                                .TextMatrix(J, 1) = Space(2) & "C"
                                .TextMatrix(J, 2) = Space(10) & "Compra"
                            End If
                            If Rs.Fields(oDocumento.Item(I).Nombre) = "S" Or Rs.Fields(oDocumento.Item(I).Nombre) = "s" Then
                                .TextMatrix(J, 1) = Space(2) & "S"
                                .TextMatrix(J, 2) = Space(10) & "Servicio"
                            End If
                        Case "Solicitado por"
                            .TextMatrix(J, 1) = Space(2) & "00000000000"
                            .TextMatrix(J, 2) = Space(10) & "Ninguno - Código de Sistema Losgistico Nro. " & Rs.Fields(oDocumento.Item(I).Nombre)
                        Case "Descuentos": .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("Otros_Descuentos", Trim(.TextMatrix(.row, 1)), "1") & _
                                                               " : " & DescripcionesdeCodigos("Otros_Descuentos", Trim(.TextMatrix(.row, 1)), "2") & "%"
                    End Select
                Else
                    .Col = 1
                    .row = J
                    .CellForeColor = vbBlack
                    If UCase(oDocumento.Item(I).Nombre) Like "*FEC*" Then
                        .TextMatrix(J, 2) = Space(10) & Format(Rs.Fields(oDocumento.Item(I).Nombre), "dd/mm/yyyy")
                    Else
                        If oDocumento.Item(I).Presicion = 2 Then
                            .TextMatrix(J, 2) = Space(10) & FormatNumber(Rs.Fields(oDocumento.Item(I).Nombre), 2)
                            If UCase(oDocumento.Item(I).Nombre) = "TOTAL" Then Monto = FormatNumber(Rs.Fields(oDocumento.Item(I).Nombre), 2)
                            If UCase(oDocumento.Item(I).Nombre) = "OTROS_MONTOS" Then MtoOtros = FormatNumber(Rs.Fields(oDocumento.Item(I).Nombre), 2)
                        Else
                            .TextMatrix(J, 2) = Space(10) & Rs.Fields(oDocumento.Item(I).Nombre)
                        End If
                    End If
                End If
                J = J + 1
            Else
                If oDocumento.Item(I).Nombre = "ImpEqui" Then
                    ImpEquiv = Rs.Fields("ImpEqui")
                End If
            End If
        Next I
    End With
    Set Rs = Nothing
    Exit Function
errorcarga:
        MsgBox "El folio no está correctamente registrado", vbCritical + vbOKOnly, "NOVPeru"
End Function

Public Function LLenarOrdenCompra(ByVal numorden As String, Optional roworden As Integer, Optional td As String)
    Dim Rs As New MYSQL_RS
    Dim aux1 As String
    Dim I, J, k As Integer
    Dim aux As Integer
    Dim fpagoaux As String
    Dim CodAux As String
    J = 1
    
    'Validacion si la fecha de la orden es mayor que la fecha de la factura
    SQL = "select fec_emision from orden_compra where correl = " & numorden
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Rs.RecordCount > 0 Then
    LblOrden = Rs.Fields("fec_emision")
    End If
    
    Set Rs = Nothing
    'Fin Validacion si la fecha de la orden es mayor que la fecha de la factura
    
    SQL = "Select factor from cndocum where coddoc='" & td & "'"
    
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    
    If Rs.RecordCount > 0 Then
        If Rs.Fields("FACTOR") > 0 Then
            SQL = "SELECT * FROM orden_compra as oc left join (Select a.orden,SUM(a.total*d.FACTOR) as monto_factura" & _
            " From ((`documento_contables` as a Left Join `amarre_documento` as b" & _
            " on (a.identificador=b.identificador)) left join `movi_documento` as c" & _
            " on (a.identificador=c.identificador)) left join `cndocum` as d" & _
            " on (d.`CODDOC`=b.cod_tipo_doc)" & _
            " where c.cod_estado<>'AN' and c.cod_estado<>'EL' and a.orden='" & numorden & "'" & _
            " group by a.orden)  as z on z.orden=oc.correl" & _
            " where oc.correl = '" & numorden & "'  and oc.flag<>'S' "
            'and (oc.otros_montos+oc.total>z.monto_factura OR z.monto_factura is null)
        Else
            SQL = "SELECT * FROM orden_compra as oc left join (Select a.orden,SUM(a.total*d.FACTOR) as monto_factura" & _
            " From ((`documento_contables` as a Left Join `amarre_documento` as b" & _
            " on (a.identificador=b.identificador)) left join `movi_documento` as c" & _
            " on (a.identificador=c.identificador)) left join `cndocum` as d" & _
            " on (d.`CODDOC`=b.cod_tipo_doc)" & _
            " where c.cod_estado<>'AN' and c.cod_estado<>'EL' and a.orden='" & numorden & "'" & _
            " group by a.orden)  as z on z.orden=oc.correl" & _
            " where oc.correl = '" & numorden & "'  and oc.flag<>'S'"
        End If
    Else
        Set Rs = Nothing
        Exit Function
    End If
    
    With flxDetalles
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        
        If Rs.RecordCount > 0 Then
            If Trim(Rs.Fields("AutoAct")) <> "2" Then  'Validación de orden con firmas incompletas
                If MsgBox("Esta orden no tiene las firmas completas!!!,Desea Continuar de todos modos?", vbQuestion + vbYesNo, "Autorización de Firmas") = vbNo Then
                 Exit Function
                End If
            End If
            'Parametros Folios Automaticos
            Orden_fol = Rs.Fields("Correl")
            Auxil_fol = Rs.Fields("Auxiliar")
            Codigo_fol = Rs.Fields("codigo")
            Cen_fol = Rs.Fields("cenco")
            Mon_fol = Rs.Fields("mon")
            dvcto_fol = "0"
            fecvcto_fol = ""
            fecemis_fol = Rs.Fields("Fec_Emision")
            tot_fol = Rs.Fields("Total")
            divi_fol = Rs.Fields("Division")
            
            For I = 1 To oDocumento.Count
                If oDocumento.Item(I).Visible = 1 Then
                    If oDocumento.Item(I).Validacion = 1 Then
                        If Trim(.TextMatrix(J, 1)) = "" Or PressF1 = True Then
                            .Col = 2
                            .row = J
                            .CellForeColor = &H80&
                            Select Case oDocumento.Item(I).Descripcion
                                Case "Auxiliar":
                                    .TextMatrix(J, 1) = Space(2) & Rs.Fields(oDocumento.Item(I).Nombre)
                                    .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("CNAUXIL", Trim(.TextMatrix(J, 1)))
                                    aux1 = Trim(.TextMatrix(J, 1))
                                Case "Codigo"
                                    .TextMatrix(J, 1) = Space(2) & Rs.Fields(oDocumento.Item(I).Nombre)
                                    .TextMatrix(J, 1) = Space(2) & Right("00000000000" & Trim(.TextMatrix(J, 1)), 11)
                                    .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(J, 1)), aux1, "Descrip")
                                    CodAux = Trim(.TextMatrix(J, 1))
                                Case "Dias de Vencimiento"
                                    .TextMatrix(J, 1) = Space(2) & Rs.Fields("CREDITO")
                                    .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("FORMPAG", Trim(.TextMatrix(J, 1)), "2")
                                    For k = 1 To .Rows - 1
                                        Select Case .TextMatrix(k, 0)
                                            Case "Fecha de Vencimiento"
                                                .TextMatrix(k, 2) = Space(10) & CalcularFechaVcto(val(Trim(.TextMatrix(J, 1))))
                                                fpagoaux = Trim(.TextMatrix(k, 2))
                                            Case "Fecha de Pago"
                                                If Trim(.TextMatrix(k, 1)) = "0" Then
                                                    .TextMatrix(k, 2) = Space(10) & fpagoaux
                                                Else
                                                    .TextMatrix(k, 2) = Space(10) & CalcularFechaPago(fpagoaux)
                                                End If
                                        End Select
                                    Next
                                Case "Centro de Costo":
                                    .TextMatrix(J, 1) = Space(2) & Rs.Fields(oDocumento.Item(I).Nombre)
                                    .TextMatrix(J, 2) = Space(10) & DescripcionesdeCodigos("CENCO", Trim(.TextMatrix(J, 1)), "1")
                                Case "Moneda"
                                    lblTipoctamn = ""
                                    lblTipoctame = ""
                                    If Rs.Fields("mon") = "N" Then
                                        lblTipoctamn = TipoPago(Trim(Rs.Fields("mpago")))
                                    Else
                                        lblTipoctame = TipoPago(Trim(Rs.Fields("mpago")))
                                    End If
                                    .TextMatrix(J, 1) = Space(2) & Rs.Fields(oDocumento.Item(I).Nombre)
                                    Select Case Rs.Fields(oDocumento.Item(I).Nombre)
                                        Case "N"
                                            .TextMatrix(J, 2) = Space(10) & "Nacional"
                                            For k = 1 To .Rows - 1
                                                If lblTipoctamn <> Empty Then
                                                    If .TextMatrix(k, 0) = "Tipo de Pago" Then
                                                        .TextMatrix(k, 1) = Space(2) & lblTipoctamn
                                                        .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(k, 1)))
                                                        .row = k
                                                        .Col = 2
                                                        .CellForeColor = &H80&
                                                    End If
                                                End If
                                            Next
                                        Case "E"
                                            .TextMatrix(J, 2) = Space(10) & "Extranjera"
                                            For k = 1 To .Rows - 1
                                                If lblTipoctame <> Empty Then
                                                    If .TextMatrix(k, 0) = "Tipo de Pago" Then
                                                        .TextMatrix(k, 1) = Space(2) & lblTipoctame
                                                        .TextMatrix(k, 2) = Space(10) & DescripcionesdeCodigos("TIPOPAGO", Trim(.TextMatrix(k, 1)))
                                                        .row = k
                                                        .Col = 2
                                                        .CellForeColor = &H80&
                                                    End If
                                                End If
                                            Next
                                        Case Else
                                            .TextMatrix(J, 2) = Space(10) & "Ninguna"
                                    End Select
                                Case "Encargado"
                                    .TextMatrix(J, 1) = Space(2) & "00000000000"
                                    .TextMatrix(J, 2) = Space(10) & "Ninguno - Código de Sistema Logístico Nro. " & Rs.Fields(oDocumento.Item(I).Nombre)
                                Case "No. de Orden"
                                    .TextMatrix(J, 1) = Space(2) & numorden
                                    Select Case UCase(Trim(Rs.Fields(3)))
                                        Case "C"
                                            .TextMatrix(J, 2) = Space(10) & "Compra"
                                        Case "S"
                                            .TextMatrix(J, 2) = Space(10) & "Servicio"
                                        Case "I"
                                            .TextMatrix(J, 2) = Space(10) & "Importación"
                                        Case Else
                                            .TextMatrix(J, 2) = Space(10) & "Ninguna"
                                    End Select
                            End Select
                        End If
                    Else
                        If Trim(.TextMatrix(J, 2)) = "" Or PressF1 = True Then
                            .Col = 1
                            .row = J
                            .CellForeColor = vbBlack
                            Select Case oDocumento.Item(I).Descripcion
                                Case "Serie": .TextMatrix(J, 2) = Empty
                                Case "Correlativo": .TextMatrix(J, 2) = Empty
                                Case Else
                                    If UCase(oDocumento.Item(I).Nombre) Like "*FEC*" Then
                                    Else
                                        .TextMatrix(J, 2) = Space(10) & Rs.Fields(oDocumento.Item(I).Nombre)
                                    End If
                                    If oDocumento.Item(I).Descripcion = "Total" Then
                                        ImporteOrden = Abs(Rs.Fields("Total") - Rs.Fields("monto_factura"))
                                        .TextMatrix(J, 2) = Space(10) & FormatNumber(ImporteOrden, 2)
                                        dblImpEquiv = FormatNumber(ImporteOrden, 2) * dblImpEquiv
                                        CalcularSubTotal ImporteOrden, .row
                                    End If
                            End Select
                        End If
                    End If
                    J = J + 1
                End If
            Next I
        Else
            MsgBox "No existe orden o ya ha sido facturada en su totalidad", vbOKOnly + vbInformation, gsNomSW
            flxDetalles.TextMatrix(1, 1) = ""
            Set Rs = Nothing
            Exit Function
        End If
    End With
    Set Rs = Nothing
    flxDetalles.row = roworden
    flxDetalles.SiguienteCelda
End Function
'Public Function ValidarEstadoOrden(firma As String) As Boolean
'    On Error GoTo novalida
'    Dim I As Integer, r As Integer
'    ValidarData = True
'    With flxDetalles
'        For I = 1 To .Rows - 1
'            Select Case .TextMatrix(I, 0)
'                Case "Serie"
'                        If Trim(.TextMatrix(I, 2)) = Empty Then
'                            MsgBox "El Campo Serie no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
'                            .TextMatrix(I, 2) = "00000"
'                            ValidarData = False
'                            .row = I
'                            Exit Function
'                        End If
'                Case "Correlativo"
'                    If Trim(.TextMatrix(I, 2)) = Empty Then
'                        MsgBox "El Campo Correlativo no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
'                        .TextMatrix(I, 2) = "000000000"
'                        ValidarData = False
'                        .row = I
'                        Exit Function
'                    End If
'                Case "Moneda"
'                    If Trim(.TextMatrix(I, 1)) = Empty Then
'                        MsgBox "El Campo Moneda no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
'                        .TextMatrix(I, 1) = ""
'                        ValidarData = False
'                        .row = I
'                        .Col = 1
'                        Exit Function
'                    End If
'                Case "Fecha de Emision"
'                    If Trim(.TextMatrix(I, 2)) <> Empty Then
'                        If Year(Trim(.TextMatrix(I, 2))) > Year(Date) And Month(Trim(.TextMatrix(I, 2))) > Month(Date) Then
'                            MsgBox "La Fecha de emisión no puede ser mayor a la fecha actual", vbOKOnly + vbCritical, "Aviso"
'                            ValidarData = False
'                            .row = I
'                            Exit Function
'                        Else
'                            If Year(Trim(.TextMatrix(I, 2))) < Year(Date) Then
'                               r = MsgBox("Fecha de emisión es de año anterior... Está Seguro de continuar?", vbYesNo + vbCritical, "Aviso")
'                               If r = vbNo Then
'                                    ValidarData = False
'                                    .row = I
'                                    Exit Function
'                               End If
'                            End If
'                        End If
'                    Else
'                        MsgBox "La Fecha de emisión no puede quedar en blanco", vbOKOnly + vbCritical, "Aviso"
'                        ValidarData = False
'                        .row = I
'                        Exit Function
'                    End If
'            End Select
'        Next
'    End With
'    Exit Function
'novalida:
'    MsgBox "Error desconocido... verifique los datos e intentelo denuevo", vbOKOnly + vbCritical, "Aviso"
'    ValidarData = False
'    Exit Function
'End Function




Public Function ValidarData() As Boolean
    On Error GoTo novalida
    Dim I As Integer, r As Integer
    Dim pivot As String
    Dim valnumdoc As String, valauxiliar As String, valcodaux As String, valcoddoc As String
    Dim flgFecDoc As String
    Dim difFecDoc As Long
    
    ValidarData = True
    valcoddoc = TxtTipo.Text
    
    'Validación de Documentos Repetidos a nivel de folios sistema administrativo.
'    flgFecDoc = Trim(flxDetalles.TextMatrix(11, 2))
'
'    difFecDoc = DateDiff("d", Right(Trim(flxDetalles.TextMatrix(11, 2)), 4) & "/" & Mid(Trim(flxDetalles.TextMatrix(11, 2)), 4, 2) & "/" & Left(Trim(flxDetalles.TextMatrix(11, 2)), 2), LblOrden)
'
    If (valcoddoc = "01") Or (valcoddoc = "02") Or (valcoddoc = "03") Or (valcoddoc = "07") Or (valcoddoc = "08") Or (valcoddoc = "10") Or (valcoddoc = "11") Or (valcoddoc = "13") Or (valcoddoc = "14") Or (valcoddoc = "S") Then
        valnumdoc = Trim(flxDetalles.TextMatrix(3, 2))  'numero de documento
        valauxiliar = Trim(flxDetalles.TextMatrix(4, 1))  'auxiliar
        valcodaux = Trim(flxDetalles.TextMatrix(5, 1))   'codaux
        
        If ValidaSiExisteDocumento(valcoddoc, valnumdoc, valauxiliar, valcodaux) Then
         If lblModo = "Nuevo" Then
          MsgBox "Revisar los posibles inconvenientes:El proveedor esta de Baja o El Documento esta DUPLICADO, por favor verificar", vbOKOnly + vbCritical, "Aviso"
          ValidarData = False
          Exit Function
         Else
          MsgBox "El Documento corre riesgo de estar DUPLICADO, cuidado mira bien a todos lados", vbOKOnly + vbCritical, "Aviso"
         End If
        End If
    End If
    
    
    If (valcoddoc = "9") Or (valcoddoc = "SG") Or (valcoddoc = "RG") Or (valcoddoc = "LS") Or (valcoddoc = "P") Or (valcoddoc = "PL") Or (valcoddoc = "TR") Then
        valnumdoc = Trim(flxDetalles.TextMatrix(1, 2))  'numero de documento
        valauxiliar = Trim(flxDetalles.TextMatrix(2, 1))  'auxiliar
        valcodaux = Trim(flxDetalles.TextMatrix(3, 1))   'codaux
        
        If ValidaSiExisteDocumento(valcoddoc, valnumdoc, valauxiliar, valcodaux) Then
         If lblModo = "Nuevo" Then
          MsgBox "El Documento esta DUPLICADO, por favor verificar", vbOKOnly + vbCritical, "Aviso"
          ValidarData = False
          Exit Function
         Else
          MsgBox "El Documento corre riesgo de estar DUPLICADO, cuidado", vbOKOnly + vbCritical, "Aviso"
         End If
        End If
    End If
    
    
    If (valcoddoc = "1") Then
        valnumdoc = Trim(flxDetalles.TextMatrix(2, 2))  'numero de documento
        valauxiliar = Trim(flxDetalles.TextMatrix(4, 1))  'auxiliar
        valcodaux = Trim(flxDetalles.TextMatrix(5, 1))   'codaux
        
        If ValidaSiExisteDocumento(valcoddoc, valnumdoc, valauxiliar, valcodaux) Then
         If lblModo = "Nuevo" Then
          MsgBox "El Documento esta DUPLICADO, por favor verificar", vbOKOnly + vbCritical, "Aviso"
          ValidarData = False
          Exit Function
         Else
          MsgBox "El Documento corre riesgo de estar DUPLICADO, cuidado mira bien a todos lados", vbOKOnly + vbCritical, "Aviso"
         End If
        End If
    End If
    

    'Fin Validación de Documentos Repetidos a nivel de folios sistema administrativo.
    
    With flxDetalles
        For I = 1 To .Rows - 1
            Select Case .TextMatrix(I, 0)
                Case "Serie"
                        If Trim(.TextMatrix(I, 2)) = Empty Then
                            MsgBox "El Campo Serie no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
                            .TextMatrix(I, 2) = "00000"
                            ValidarData = False
                            .row = I
                            Exit Function
                        End If
                Case "Correlativo"
                    If Trim(.TextMatrix(I, 2)) = Empty Then
                        MsgBox "El Campo Correlativo no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
                        .TextMatrix(I, 2) = "000000000"
                        ValidarData = False
                        .row = I
                        Exit Function
                    End If
                Case "Moneda"
                    If Trim(.TextMatrix(I, 1)) = Empty Then
                        MsgBox "El Campo Moneda no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
                        .TextMatrix(I, 1) = ""
                        ValidarData = False
                        .row = I
                        .Col = 1
                        Exit Function
                    End If
                Case "Centro de Costo"
                    If Trim(.TextMatrix(I, 1)) = Empty Then
                        MsgBox "El Centro de Costo no puede quedar vacio", vbOKOnly + vbCritical, "Aviso"
                        .TextMatrix(I, 1) = ""
                        ValidarData = False
                        .row = I
                        .Col = 1
                        Exit Function
                    End If
                Case "Fecha de Emision"
                    If Trim(.TextMatrix(I, 2)) <> Empty Then
                        If Year(Trim(.TextMatrix(I, 2))) > Year(Date) And Month(Trim(.TextMatrix(I, 2))) > Month(Date) Then
                            MsgBox "La Fecha de emisión no puede ser mayor a la fecha actual", vbOKOnly + vbCritical, "Aviso"
                             txtObs = "La Fecha de Emisión es mayor a la fecha actual"
'                            ValidarData = False
'                            .row = I
'                            Exit Function
                        Else
                            If Year(Trim(.TextMatrix(I, 2))) < Year(Date) Then
                               r = MsgBox("Fecha de emisión es de año anterior... Está Seguro de continuar?", vbYesNo + vbCritical, "Aviso")
                               If r = vbNo Then
                                    ValidarData = False
                                    .row = I
                                    Exit Function
                               End If
                            End If
                        End If
                    Else
                        MsgBox "La Fecha de emisión no puede quedar en blanco", vbOKOnly + vbCritical, "Aviso"
                        ValidarData = False
                        .row = I
                        Exit Function
                    End If
            End Select
        Next
    End With
    Exit Function
novalida:
    MsgBox "Error desconocido... verifique los datos e intentelo denuevo", vbOKOnly + vbCritical, "Aviso"
    ValidarData = False
    Exit Function
End Function

Private Function ConfirGuarda(folio As String) As Boolean
    Dim SQL As String
    Dim rsconfirm As MYSQL_RS
    Dim rsfamilias As MYSQL_RS
    ConfirGuarda = False
    SQL = "Select Identificador,Cod_Fam from amarre_documento where identificador = '" & folio & "'"
    Set rsconfirm = oConexion.EjecutaSelectRS(SQL)
    If Not rsconfirm.EOF Then
        Select Case rsconfirm.Fields("Cod_Fam")
               Case FAMILIA_DOC.CONTABLES
                    SQL = "Select Identificador from documento_contables where identificador = '" & folio & "'"
                    Set rsfamilias = oConexion.EjecutaSelectRS(SQL)
                    If Not rsfamilias.EOF Then: ConfirGuarda = True: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    If rsfamilias.RecordCount = 0 Then ConfirGuarda = False: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    
               Case FAMILIA_DOC.ENTIDADES
                    SQL = "Select Identificador from documento_entidades where identificador = '" & folio & "'"
                    Set rsfamilias = oConexion.EjecutaSelectRS(SQL)
                    If Not rsfamilias.EOF Then ConfirGuarda = True: Set rsfamilias = Nothing:   Set rsconfirm = Nothing: Exit Function
                    If rsfamilias.RecordCount = 0 Then ConfirGuarda = False: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    
               Case FAMILIA_DOC.GENERALES
                    SQL = "Select Identificador from documento_generales where identificador = '" & folio & "'"
                    Set rsfamilias = oConexion.EjecutaSelectRS(SQL)
                    If Not rsfamilias.EOF Then ConfirGuarda = True: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    If rsfamilias.RecordCount = 0 Then ConfirGuarda = False: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    
               Case FAMILIA_DOC.ORDENES
                    SQL = "Select Identificador from orden_compra where identificador = '" & folio & "'"
                    Set rsfamilias = oConexion.EjecutaSelectRS(SQL)
                    If Not rsfamilias.EOF Then ConfirGuarda = True: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
                    If rsfamilias.RecordCount = 0 Then ConfirGuarda = False: Set rsfamilias = Nothing: Set rsconfirm = Nothing: Exit Function
        End Select
    End If
    If rsconfirm.RecordCount = 0 Then ConfirGuarda = False: Set rsconfirm = Nothing: Exit Function
End Function

Private Sub LimpiarDatos()
    Limpiar frmIngresarDocumento
    txtnombredoc = Empty
    LlenarDetalles
    DtpFecha = Date
    mskhora = Format(Time, "HH:MM")
End Sub

Public Sub BloqueoControles(valor As Boolean)
    TxtTipo.Locked = valor
    TxtNombreEmpresa.Locked = valor
    txtDesDocIden.Locked = valor
    TxtIdemisor.Locked = valor
    txtObs.Locked = valor
    TxTipo_Ide_Doc.Locked = valor
    TxtNomEmisor.Locked = valor
    txtFolio.Locked = Not valor
    mskhora.Enabled = Not valor
    OptRecepcion.Enabled = Not valor
    OptRemitir.Enabled = Not valor
    OptRecepcion.Value = Not valor
    chkcierre.Locked = valor
    If valor = True Then
        TxtTipo.BackColor = ColorDeshabilitado
        TxtNombreEmpresa.BackColor = ColorDeshabilitado
        TxtIdemisor.BackColor = ColorDeshabilitado
        txtObs.BackColor = ColorDeshabilitado
        TxTipo_Ide_Doc.BackColor = ColorDeshabilitado
        TxtNomEmisor.BackColor = ColorDeshabilitado
        mskhora.BackColor = ColorDeshabilitado
    Else
        TxtTipo.BackColor = ColorHabilitado
        TxtNombreEmpresa.BackColor = ColorHabilitado
        TxtIdemisor.BackColor = ColorHabilitado
        txtObs.BackColor = ColorHabilitado
        TxTipo_Ide_Doc.BackColor = ColorHabilitado
        TxtNomEmisor.BackColor = ColorHabilitado
        mskhora.BackColor = ColorHabilitado
        txtFolio.BackColor = ColorDeshabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             lblModo = "Acción"
             BloqueoControles True
             TxtTipo.Locked = True
             txtFolio.Locked = False
             TxtTipo.BackColor = ColorDeshabilitado
             txtFolio.BackColor = ColorHabilitado
             BtnNuevo.Enabled = True
             BtnSalir.Enabled = True
             BtnModificar.Enabled = False
             cmdProg.Enabled = False
             btnGrabar.Enabled = False
             chkcierre.Value = False
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             Publimensaje = "modificar"
             lblModo = "Nuevo"
             BloqueoControles False
             ConfigurarBotones cfgNuevo
             lblGuardado.Visible = False
             'mskFecha = Date
             BtnNuevo.Enabled = True
             lblEstadoDoc = ""
             chkcierre.Value = False
             Exit Sub
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             txtFolio.Locked = True
             txtFolio.BackColor = ColorDeshabilitado
             ConfigurarBotones cfgGrabar
             BtnNuevo.Enabled = True
         Case ModoForm.modEditar
             lblModo = "Modificar"
             BloqueoControles False
             ConfigurarBotones cfgModificar
             Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            cmdProg.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnReportes.Enabled = False
            BtnCancelar.Enabled = True
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnReportes.Enabled = False
            BtnCancelar.Enabled = True
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = False
            cmdProg.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnReportes.Enabled = False
            BtnCancelar.Enabled = False
            BtnSalir.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            If PermiUsu(strUsuarioId) Then
               BtnModificar.Enabled = True
            Else: BtnModificar.Enabled = False
            End If
            cmdProg.Enabled = True
            btnGrabar.Enabled = False
            BtnEliminar.Enabled = True
            btnReportes.Enabled = True
            BtnCancelar.Enabled = True
            Publimensaje = ""
            Exit Sub
            
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo", "Acción"
                     lblModo = "Acción"
                     Publimensaje = ""
                     ModoFormulario modAccion
                     lblGuardado.Visible = False
                Case "Consulta"
                     lblModo = "Acción"
                     Publimensaje = ""
                     ModoFormulario modAccion
                     BtnEliminar.Enabled = False
                     lblGuardado.Visible = False
                     lblEstadoDoc = ""
                Case "Modificar"
                     Publimensaje = "modificar"
                     ModoFormulario modConsulta

            End Select
    End Select
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
    On Error Resume Next
    
    With flxDetalles
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



Private Function GenerarDivHFM(ByVal vDivision As String, ByVal vCenco As String) As String
    Dim RsDIVHFM As MYSQL_RS
    Dim MovAux As String
    Dim SQL As String
    
    If vDivision <> "0006" Then
        SQL = "select Cod_Auxil FROM CNCOSTO WHERE division='" & vDivision & "' and CENCO='" & vCenco & "' limit 1"
        Set RsDIVHFM = oConexion.EjecutaSelectRS(SQL)
        
        If Not RsDIVHFM.EOF Then
            MovAux = "0" & RsDIVHFM.Fields("Cod_Auxil")
        End If
        
        If RsDIVHFM.RecordCount = 0 Then
            MovAux = "013100003836"
        End If
        
            GenerarDivHFM = MovAux
    Else
            GenerarDivHFM = "0006"
    End If
    
    RsDIVHFM.CloseRecordset
    Set RsDIVHFM = Nothing
End Function



Function NoExisteAuxiliar(ByVal strAuxiliar As String, ByVal strCodigo As String) As Boolean
    Dim SQL As String
    Dim rsR As MYSQL_RS
    NoExisteAuxiliar = True
    
    SQL = " SELECT CODIGO FROM CNAUXIL WHERE AUXILIAR='" & strAuxiliar & "'  AND CODIGO='" & strCodigo & "'"
    Set rsR = oConexion.EjecutaSelectRS(SQL)
    
   Do While Not rsR.EOF
     NoExisteAuxiliar = False
   Loop
    
    Set rsR = Nothing
End Function

Function ValidaSiExisteDocumento(ByVal strcoddoc As String, ByVal strdoc As String, ByVal strAuxiliar As String, ByVal strCodigo As String) As Boolean
    Dim SQL As String
    Dim rsV As MYSQL_RS
    ValidaSiExisteDocumento = False
    
    SQL = "select d.identificador,(SELECT COD_ESTADO FROM MOVI_DOCUMENTO WHERE IDENTIFICADOR=d.Identificador) AS estado,(select FechaBaja from cnauxil where auxiliar='5' and codigo='" & strCodigo & "') AS fecbaja from documento_contables as d " & _
          "left join amarre_documento as f on d.Identificador = f.Identificador " & _
          "where f.Cod_Tipo_Doc='" & strcoddoc & "' and d.correl='" & strdoc & "' and d.AUXILIAR='" & strAuxiliar & "'  AND d.CODIGO='" & strCodigo & "'  "
    Set rsV = oConexion.EjecutaSelectRS(SQL)
    
    Do While Not rsV.EOF
        If rsV.Fields("estado") <> "EL" Then
         ValidaSiExisteDocumento = True
         Exit Do
        End If
        
        'Valida si esta dado de baja el proveedor
        If rsV.Fields("fecbaja") <> "" Then
            If CDate(rsV.Fields("fecbaja")) <= Date Then
             ValidaSiExisteDocumento = True
             Exit Do
            End If
        End If
        
        rsV.MoveNext
    Loop
    
    Set rsV = Nothing
End Function

Function EncuentraArchivo(Nom As String, TxtID As String) As Boolean
EncuentraArchivo = False
    Dim SQL As String
    Dim RT As MYSQL_RS
    Dim RutaDestinoCliente As String
  
    RutaDestinoCliente = "\\172.26.35.12\ddsi$\Cobranzas\" & TxtID & ""
    SQL = "select * from archivosadjuntos where modulo = '1' and identificador = '" & TxtID & "' " & _
          "and ruta = '" & Replace(RutaDestinoCliente & "\", "\", "*") & "' and nombre = '" & Nom & "'"
    Set RT = oConexion.EjecutaSelectRS(SQL)
    
    If Not RT.EOF() Then
        EncuentraArchivo = True
    End If
    
    Set RT = Nothing
End Function


