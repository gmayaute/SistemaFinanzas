VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmMovTelewiese 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orden de Transferencia Telebanking"
   ClientHeight    =   6495
   ClientLeft      =   2310
   ClientTop       =   6495
   ClientWidth     =   12150
   Icon            =   "frmMovTelewiese.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   12150
   Begin NOVAdmin.flxEdit flexDocumentos 
      Height          =   2265
      Left            =   0
      TabIndex        =   62
      Top             =   2010
      Width           =   9735
      _ExtentX        =   15690
      _ExtentY        =   3995
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
      CellPicture     =   "frmMovTelewiese.frx":08CA
      ColAlignment0   =   9
      FixedAlignment0 =   9
      ForeColorSel    =   16711680
      ForeColorFixed  =   14474460
      MouseIcon       =   "frmMovTelewiese.frx":08E6
      RowHeight0      =   240
   End
   Begin VB.PictureBox picDerecho 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   8970
      ScaleHeight     =   2310
      ScaleWidth      =   3150
      TabIndex        =   52
      Top             =   1980
      Width           =   3150
      Begin Proyecto1.chameleonButton cmdAnexar 
         Height          =   375
         Left            =   840
         TabIndex        =   53
         ToolTipText     =   "Nuevo"
         Top             =   30
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Documentos"
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
         BCOL            =   14737632
         BCOLO           =   15309923
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMovTelewiese.frx":0902
         PICN            =   "frmMovTelewiese.frx":091E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.Label lblTCambio 
         Height          =   315
         Left            =   1560
         TabIndex        =   61
         Top             =   1530
         Width           =   645
         ForeColor       =   128
         Caption         =   "3.15"
         Size            =   "1138;556"
         SpecialEffect   =   2
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/C"
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
         Left            =   810
         TabIndex        =   60
         Top             =   1530
         Width           =   525
      End
      Begin MSForms.Label lblImpEqu 
         Height          =   315
         Left            =   1380
         TabIndex        =   59
         Top             =   1890
         Width           =   1695
         ForeColor       =   8388608
         Caption         =   "0.00"
         Size            =   "2990;556"
         SpecialEffect   =   2
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin VB.Label lblMonEqu 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   810
         TabIndex        =   58
         Top             =   1920
         Width           =   525
      End
      Begin MSForms.Label lblDocAnexados 
         Height          =   225
         Left            =   840
         TabIndex        =   57
         Top             =   450
         Visible         =   0   'False
         Width           =   2220
         ForeColor       =   65280
         BackColor       =   8421504
         Size            =   "3916;397"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe "
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
         Left            =   840
         TabIndex        =   56
         Top             =   750
         Width           =   2220
      End
      Begin VB.Label lblMoneda 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   840
         TabIndex        =   55
         Top             =   1140
         Width           =   495
      End
      Begin MSForms.TextBox meImporte 
         Height          =   315
         Left            =   1350
         TabIndex        =   54
         Top             =   1140
         Width           =   1695
         VariousPropertyBits=   746604569
         BackColor       =   14737632
         ForeColor       =   8388608
         Size            =   "2990;556"
         Value           =   "0.00"
         BorderColor     =   12632256
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
   End
   Begin VB.PictureBox picInferior 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   45
      ScaleHeight     =   1995
      ScaleWidth      =   12075
      TabIndex        =   25
      Top             =   4275
      Width           =   12075
      Begin Proyecto1.chameleonButton BtnModificar 
         Height          =   345
         Left            =   1290
         TabIndex        =   35
         ToolTipText     =   "Modificar"
         Top             =   1530
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "M&odificar"
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
         MICON           =   "frmMovTelewiese.frx":0A78
         PICN            =   "frmMovTelewiese.frx":0A94
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton ChBtnSalir 
         Height          =   345
         Left            =   10350
         TabIndex        =   36
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":0EC2
         PICN            =   "frmMovTelewiese.frx":0EDE
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
         Left            =   9780
         TabIndex        =   37
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":12A4
         PICN            =   "frmMovTelewiese.frx":12C0
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
         Left            =   2640
         TabIndex        =   38
         ToolTipText     =   "Eliminar"
         Top             =   1530
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "frmMovTelewiese.frx":1802
         PICN            =   "frmMovTelewiese.frx":181E
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
         Left            =   4680
         TabIndex        =   39
         ToolTipText     =   "Guardar"
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":1C60
         PICN            =   "frmMovTelewiese.frx":1C7C
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
         Left            =   30
         TabIndex        =   40
         ToolTipText     =   "Nuevo"
         Top             =   1530
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Nuevo"
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
         MICON           =   "frmMovTelewiese.frx":20BE
         PICN            =   "frmMovTelewiese.frx":20DA
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
         Left            =   4230
         TabIndex        =   41
         ToolTipText     =   "Deshacer"
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":2444
         PICN            =   "frmMovTelewiese.frx":2460
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnInterfaz 
         Height          =   345
         Left            =   8640
         TabIndex        =   42
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":29A2
         PICN            =   "frmMovTelewiese.frx":29BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnAnulaTw 
         Height          =   345
         Left            =   9210
         TabIndex        =   43
         ToolTipText     =   "Anular Transeferencia"
         Top             =   1530
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
         MICON           =   "frmMovTelewiese.frx":4D40
         PICN            =   "frmMovTelewiese.frx":4D5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdenviar 
         Height          =   360
         Left            =   270
         TabIndex        =   44
         ToolTipText     =   "Enviar Documento"
         Top             =   510
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   635
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
         MICON           =   "frmMovTelewiese.frx":7B1E
         PICN            =   "frmMovTelewiese.frx":7B3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H009F5539&
         Caption         =   "Opciones de Orden"
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
         Height          =   1035
         Left            =   5400
         TabIndex        =   26
         Top             =   0
         Width           =   6615
         Begin VB.TextBox txtCodOpcional 
            Height          =   315
            Left            =   1980
            MaxLength       =   11
            TabIndex        =   32
            Top             =   180
            Width           =   1425
         End
         Begin VB.TextBox txtDocOpcional 
            Height          =   315
            Left            =   1980
            MaxLength       =   15
            TabIndex        =   29
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkCodOpcional 
            BackColor       =   &H009F5539&
            Caption         =   "Orden con Código:"
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
            Left            =   30
            TabIndex        =   31
            Top             =   240
            Width           =   1965
         End
         Begin VB.CheckBox chkDocOpcional 
            BackColor       =   &H009F5539&
            Caption         =   "Orden con Doc.:"
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
            Left            =   30
            TabIndex        =   30
            Top             =   660
            Width           =   1965
         End
         Begin VB.CheckBox chkPreparar 
            BackColor       =   &H009F5539&
            Caption         =   "Preparar Orden"
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
            Left            =   3510
            TabIndex        =   28
            Top             =   750
            Width           =   2145
         End
         Begin VB.CheckBox chkPagoUnico 
            BackColor       =   &H009F5539&
            Caption         =   "Pago Unicio"
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
            Left            =   3510
            TabIndex        =   27
            Top             =   510
            Width           =   2145
         End
         Begin Proyecto1.chameleonButton cmdArchivos 
            Height          =   345
            Left            =   6090
            TabIndex        =   67
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
            MICON           =   "frmMovTelewiese.frx":84B0
            PICN            =   "frmMovTelewiese.frx":84CC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSForms.ComboBox cmbTipoOrden 
            Height          =   315
            Left            =   4050
            TabIndex        =   34
            Top             =   120
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
         Begin VB.Label Label15 
            BackColor       =   &H009F5539&
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
            Height          =   195
            Left            =   3450
            TabIndex        =   33
            Top             =   240
            Width           =   375
         End
      End
      Begin MSForms.ComboBox cmbLiqPagos 
         Height          =   315
         Left            =   11160
         TabIndex        =   63
         Top             =   1560
         Width           =   855
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "1508;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtConcepto 
         Height          =   975
         Left            =   1020
         TabIndex        =   51
         Top             =   30
         Width           =   4305
         VariousPropertyBits=   746604571
         Size            =   "7594;1720"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Obs:"
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
         Left            =   0
         TabIndex        =   50
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   0
         TabIndex        =   49
         Top             =   1080
         Width           =   12015
      End
      Begin VB.Label lblModo 
         BackColor       =   &H00000000&
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
         Left            =   8460
         TabIndex        =   48
         Top             =   1140
         Width           =   2655
      End
      Begin VB.Label lblMsjgrilla 
         BackColor       =   &H80000008&
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
         TabIndex        =   47
         Top             =   1110
         Width           =   8205
      End
      Begin MSForms.Label lblV 
         Height          =   345
         Left            =   5880
         TabIndex        =   46
         Top             =   1560
         Width           =   1935
         ForeColor       =   8421631
         BackColor       =   10442041
         Caption         =   "Voucher Cancelación:"
         Size            =   "3413;609"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblvoucher 
         Height          =   345
         Left            =   7920
         TabIndex        =   45
         Top             =   1560
         Width           =   765
         ForeColor       =   65280
         BackColor       =   10442041
         Caption         =   "010001"
         Size            =   "1349;609"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   1965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12075
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   6720
         MaxLength       =   11
         TabIndex        =   5
         Top             =   480
         Width           =   1905
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   10110
         TabIndex        =   3
         Top             =   150
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Proyecto1.chameleonButton cmdSaldosporPagar 
         Height          =   300
         Left            =   11580
         TabIndex        =   66
         ToolTipText     =   "Enviar Documento"
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
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
         MICON           =   "frmMovTelewiese.frx":A84E
         PICN            =   "frmMovTelewiese.frx":A86A
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
         Height          =   525
         Left            =   11430
         TabIndex        =   68
         Top             =   1320
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   926
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
         MICON           =   "frmMovTelewiese.frx":B1E0
         PICN            =   "frmMovTelewiese.frx":B1FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Liq:"
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
         Height          =   435
         Left            =   11400
         TabIndex        =   65
         Top             =   480
         Width           =   825
      End
      Begin MSForms.TextBox TxtLiq 
         Height          =   345
         Left            =   11400
         TabIndex        =   64
         Top             =   840
         Width           =   675
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   10
         Size            =   "1191;609"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtCuentaAux 
         Height          =   315
         Left            =   9810
         TabIndex        =   10
         Top             =   1200
         Width           =   1485
         VariousPropertyBits=   746604575
         MaxLength       =   20
         Size            =   "2619;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta:"
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
         Left            =   8760
         TabIndex        =   24
         Top             =   1200
         Width           =   915
      End
      Begin MSForms.Label lblEstado 
         Height          =   255
         Left            =   3180
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   1665
         ForeColor       =   8421631
         BackColor       =   -2147483634
         VariousPropertyBits=   8388627
         Caption         =   "TRANSFERIDO"
         Size            =   "2937;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTipoPago 
         Height          =   315
         Left            =   8760
         TabIndex        =   9
         Top             =   840
         Width           =   2505
         VariousPropertyBits=   746588191
         DisplayStyle    =   7
         Size            =   "4419;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Pago:"
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
         Left            =   8760
         TabIndex        =   22
         Top             =   510
         Width           =   1305
      End
      Begin MSForms.TextBox txtOficina 
         Height          =   315
         Left            =   9810
         TabIndex        =   11
         Top             =   1530
         Width           =   1485
         VariousPropertyBits=   746604575
         MaxLength       =   11
         Size            =   "2619;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Oficina:"
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
         Left            =   8760
         TabIndex        =   21
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sr(es):"
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
         TabIndex        =   20
         Top             =   870
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código/RUC:"
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
         Left            =   5370
         TabIndex        =   19
         Top             =   510
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cta. Cte:"
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
         Left            =   3060
         TabIndex        =   18
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auxiliar:"
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
         TabIndex        =   17
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda:"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1275
      End
      Begin MSForms.ComboBox cmbMoneda 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   1410
         Width           =   1635
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "2884;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCtaCte 
         Height          =   345
         Left            =   4020
         TabIndex        =   8
         Top             =   1380
         Width           =   4275
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "7541;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro. Orden:"
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
         TabIndex        =   15
         Top             =   120
         Width           =   1275
      End
      Begin MSForms.ComboBox cmbAuxiliares 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   510
         Width           =   3435
         DisplayStyle    =   7
         Size            =   "6059;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ListBox lstBeneficiario 
         Height          =   1035
         Left            =   1380
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   870
         Visible         =   0   'False
         Width           =   6975
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "12303;1826"
         MatchEntry      =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtBeneficiario 
         Height          =   345
         Left            =   1380
         TabIndex        =   6
         Top             =   870
         Width           =   6885
         VariousPropertyBits=   603998235
         Size            =   "12144;609"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
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
         Left            =   5370
         TabIndex        =   13
         Top             =   150
         Width           =   1365
      End
      Begin MSForms.TextBox meOrden 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   120
         Width           =   1755
         VariousPropertyBits=   746604571
         MaxLength       =   4
         Size            =   "3096;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox lblFolio 
         Height          =   345
         Left            =   6810
         TabIndex        =   2
         Top             =   120
         Width           =   1875
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   10
         Size            =   "3307;609"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   8760
         TabIndex        =   12
         Top             =   120
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmMovTelewiese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tipcta_mn As String
Public numcta_mn As String
Public tipcta_me As String
Public numcta_me As String
Private Cambiaaux As Boolean
Private oConsulta As FrmConsultas
Private Conv As New clsNumToLet
Dim FlgTw As Boolean

Public Function GrabaOrden() As Boolean
    Dim I As Integer
    Dim sqlorden As String
    Dim sqlHist As String
    Dim sqlMovi As String
    Dim orden As String
    Dim Cta As String
    GrabaOrden = False
    
    orden = strAnoSistema + strMesSistema + Right("0000" + Trim(meOrden), 4)
    If ValidarDatos = True Then
        With flexDocumentos
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) <> "" Then
                    If val(.TextMatrix(I, 11)) = 9 Then Cta = Trim(.TextMatrix(I, 13)) Else Cta = Right(.TextMatrix(I, 13), 7)
                    
                    sqlorden = "Call Insert_MovTw (" & I & ", " & _
                               " '" & orden & "', " & _
                               " '" & Trim(Format(mskFecha.Text, "yyyy/mm/dd")) & "', " & _
                               " '" & .TextMatrix(I, 5) & "', " & _
                               " '" & .TextMatrix(I, 10) & "', " & _
                               " '" & .TextMatrix(I, 2) & "', " & _
                               " '" & cmbMoneda.List(cmbMoneda.ListIndex, 1) & "', " & _
                               " '" & Trim(cmbCtaCte.List(cmbCtaCte.ListIndex, 1)) & "', " & _
                               " '" & .TextMatrix(I, 3) & "', " & _
                               " '" & .TextMatrix(I, 1) & "', " & _
                               " '" & .TextMatrix(I, 6) & "', " & _
                               " '" & .TextMatrix(I, 7) & "', " & _
                               "  " & CDbl(CEN(.TextMatrix(I, 8))) & "," & _
                               "  " & CDbl(CEN(.TextMatrix(I, 9))) & ", " & _
                               " '" & Cta & "'," & _
                               " '" & .TextMatrix(I, 12) & "'," & _
                               " '" & .TextMatrix(I, 11) & "', " & _
                               " '" & txtConcepto & "', " & _
                               " '" & EMITIDO & "','" & .TextMatrix(I, 14) & "')"
                    oConexion.EjecutaInsertUpdateDelete sqlorden, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        lblEstado.Visible = True
        lblEstado.tag = EMITIDO
        lblEstado = DescripcionesdeCodigos("DOC_ESTADO", EMITIDO)
        GrabaOrden = True
    End If
End Function

Private Sub btnAnulaTw_Click()
    Dim RES As Integer
    RES = MsgBox("¿Esta Seguro que desea ANULAR la transferencia de la orden Nro. " & meOrden, vbQuestion + vbYesNo, gsNomSW)
    If RES = vbYes Then
        InvalidaOrden strAnoSistema + strMesSistema + meOrden, EMITIDO
        AnulaOrden strAnoSistema + strMesSistema + meOrden
        ActualizaAnulaEstadoReporte strAnoSistema & strMesSistema & meOrden
        If lblvoucher.Visible = True Then
            If MsgBox("¿Se va a eliminar el voucher " & lblvoucher.Caption & " de cancelación", vbQuestion + vbYesNo, gsNomSW) = vbYes Then
                oConexionMYSQL.Execute "delete from cnvouc where anomes='" & strAnoSistema + strMesSistema & "' and voucher='" & lblvoucher.Caption & "'"
            End If
        End If
        lblEstado = "EMITIDO"
        lblEstado.tag = EMITIDO
        ModoFormulario modConsulta
    End If
End Sub

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
End Sub
Private Sub btnEliminar_Click()
    Dim RES As Integer
    RES = MsgBox("¿Esta Seguro que desea ELIMINAR el movimiento de la orden Nro. " & meOrden, vbQuestion + vbYesNo, gsNomSW)
    If RES = vbYes Then
        InvalidaOrden strAnoSistema + strMesSistema + meOrden, ELIMINADO
        AnulaOrden strAnoSistema + strMesSistema + meOrden
        lblEstado.tag = ELIMINADO
        lblEstado = "ELIMINADO"
        ModoFormulario modConsulta
    End If
End Sub
Private Sub InvalidaOrden(Ident As String, Estado As String)
    Dim sqlMovi As String
    Dim sqlHist As String
    sqlMovi = "Call Update_MovTwEstado('" & Ident & "','" & Estado & "')"
    oConexion.EjecutaInsertUpdateDelete sqlMovi, TIPO_QUERY.Modificar, False
End Sub
Public Sub AnulaOrden(orden As String)
    Dim I As Integer
    Dim SQL As String
    Dim rsCuenta As MYSQL_RS
    SQL = "Select identificador from documento_contables where ref='WIESE" & orden & "'"
    Set rsCuenta = oConexion.EjecutaSelectRS(SQL)
    If rsCuenta.RecordCount > 0 Then
        rsCuenta.MoveFirst
        Do While Not rsCuenta.EOF
            SQL = "Update documento_contables set ref='',cancelado=0 where identificador='" & rsCuenta.Fields("Identificador") & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
            CambioEstado rsCuenta.Fields("Identificador"), REGISTRADO
            rsCuenta.MoveNext
        Loop
    End If
    rsCuenta.CloseRecordset
    Set rsCuenta = Nothing
End Sub
Private Sub btnGrabar_Click()
    If lblModo = "Nuevo Movimiento" Then
        If GrabaOrden Then
            ModoFormulario modConsulta
        End If
    End If
    If lblModo = "Modificar Movimiento" Then
        If ActualizaOrden Then
            If lblDocAnexados.Visible = True Then
                AnulaOrden strAnoSistema + strMesSistema + Right("0000" & Trim(meOrden), 4)
                MovReferencia strAnoSistema + strMesSistema + Right("0000" & Trim(meOrden), 4)
            End If
            ModoFormulario modConsulta
        End If
    End If
End Sub

Public Sub MovReferencia(orden As String)
    Dim I As Integer
    Dim SQL As String
    If CDbl(lblDocAnexados.tag) > 0 Then
        For I = 1 To flexDocumentos.Rows - 1
             SQL = "Update documento_contables set cancelado=" & CDbl(flexDocumentos.TextMatrix(I, 8)) & ", ref='WIESE" & orden & "' where identificador='" & flexDocumentos.TextMatrix(I, 1) & "'"
             oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
             CambioEstado flexDocumentos.TextMatrix(I, 1), CANCELADO
        Next
    End If
End Sub

Private Sub btnInterfaz_Click()
    Dim TipOrden As String
    Dim aux As String
    aux = cmbAuxiliares.List(cmbAuxiliares.ListIndex, 1)
    If meOrden <> Empty And txtCodigo <> Empty Then
        If cmbTipoOrden.ListCount > 1 Then
            TipOrden = cmbTipoOrden.List(cmbTipoOrden.ListIndex, 1)
        Else
            TipOrden = "0"
        End If
        If GeneraTxt(meOrden, aux, txtCodOpcional, txtDocOpcional, chkPreparar.Value, TipOrden, chkPagoUnico.Value) Then
            MovReferencia strAnoSistema & strMesSistema & meOrden
            ActualizaEstadoReporte strAnoSistema & strMesSistema & meOrden
            If aux = "5" Or aux = "6" Then
                GeneraAsisentoCancelacion strAnoSistema & strMesSistema & meOrden
            End If
            'If aux = "3" And (flexDocumentos.TextMatrix(1, 3)) <> "SG" And (flexDocumentos.TextMatrix(1, 3)) <> "RG" And (flexDocumentos.TextMatrix(1, 6)) <> "CAJ" Then
            If aux = "3" And (flexDocumentos.TextMatrix(1, 3)) <> "SG" Then
                GeneraAsisentoPago meOrden
            End If
            
            'cmdenviarNotificacionPago
            
            lblEstado = "TRANSFERIDO"
            lblEstado.tag = TRANSFERIDO
            ModoFormulario modConsulta
            meOrden.SetFocus
        Else
            MsgBox "Error al Generar el archivo Telewiese", vbInformation, gsNomSW
        End If
    End If
End Sub


Public Sub ActualizaEstadoReporte(vOrden As String)
       SQL = "Call cn_Update_liquidPagosRep('" & vOrden & "');"
           oConexionMYSQL.Execute (SQL)
End Sub

Public Sub ActualizaEstadoReporteEl(vOrden As String)
       SQL = "Call cn_Delete_liquidPagosRep('');"
           oConexionMYSQL.Execute (SQL)
End Sub

Public Sub ActualizaAnulaEstadoReporte(vOrden As String)
    SQL = "Call cn_Update_Anula_liquidPagosRep('" & vOrden & "');"
           oConexionMYSQL.Execute (SQL)
End Sub

Public Sub GeneraAsisentoPago(orden As String)
On Error GoTo FallaAsiento
    Dim I As Integer
    Set Rs = New MYSQL_RS
    Dim SerDocu As String, NumDocu As String, v As String, AnoMes As String
    Dim glo As String, fec As String, SQL As String, vc As String, cor As String
    Dim mon As String, td As String, Div As String, Cta As String, dh As String
    Dim aux As String, caux As String, cto As String, Col As String
    Dim sol As Double, dol As Double, imp10s As Double, imp10d As Double
    Dim GlosaAct As String, Concepto As String
    imp10s = 0
    imp10d = 0
    fec = Format(CStr(mskFecha), "dd/mm/yyyy")
    TipoCambio fec
    tc = dblTipoCmbV
    AnoMes = strAnoSistema & strMesSistema
    mon = cmbMoneda.List(cmbMoneda.ListIndex, 1)
    glo = "POR ACTUALIZAR"
    v = MaxVoucher(AnoMes, "01")
    SQL = "Call cn_Insert_Voucher('" & Left(Trim(v), 2) & "','" & v & "','" & glo & "','" & Trim(fec) & _
            "','" & Trim(fec) & "','V'," & tc & ",'" & Trim(mon) & "','" & Trim(AnoMes) & "','" & strUsuarioId & _
             " ','CUADRADO','','','N','','N','','')"
            oConexionMYSQL.Execute (SQL)
    lib = "01"
    cencos = "0000"
    cenco = "00000000000"
    gen = "N"
    
    With flexDocumentos
        For I = 1 To .Rows - 1
            aux = Trim(.TextMatrix(I, 10))
            caux = Trim(.TextMatrix(I, 2))
            tdoc = IIf(Trim(.TextMatrix(I, 3)) = "RG", "SG", "7")
            Divi = DescripcionesdeCodigos("CONTRATO", caux)
            Divi = IIf(Divi = "", "0000", Divi)
            cor = MaxCorrela(AnoMes, v)
            SerDocu = Trim(.TextMatrix(I, 6))
            NumDocu = Trim(.TextMatrix(I, 7))
            Select Case Trim(.TextMatrix(I, 3))
                Case "RG", "SG"
                        If aux = "3" Then
                            Cta = IIf(mon = "E", "141302", "141301")
                        Else
                            If aux = "6" Then
                                Cta = IIf(mon = "E", "168102", "168101")
                            Else
                                Cta = IIf(mon = "E", "141302", "141301")
                            End If
                        End If
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "REEMBOLSO DE GASTOS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "REEMBOLSO DE GASTOS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "REEMBOLSO DE GASTOS "
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If

                Case "SS"
                    If InStr(1, Trim(.TextMatrix(I, 6)), "VAC") > 0 Then
                        Cta = "421202"
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "VACACIONES"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If

                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "ADEL") > 0 Or InStr(1, Trim(.TextMatrix(I, 6)), "ADL") > 0 Then
                        Cta = IIf(mon = "E", "141202", "141201")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "ADELANTO DE SUELDO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "PRE") > 0 Then
                        Cta = IIf(mon = "E", "141102", "141101")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "PRESTAMO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    Else
                        Cta = IIf(mon = "E", "141102", "141101")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "OTROS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "OTROS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "OTROS"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    End If
                Case "PL"
                    If InStr(1, Trim(.TextMatrix(I, 6)), "VAC") > 0 Then
                        Cta = "411501"
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "VACACIONES"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "ADEL") > 0 Or InStr(1, Trim(.TextMatrix(I, 6)), "ADL") > 0 Then
                        Cta = IIf(mon = "E", "141202", "141201")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "ADELANTO DE SUELDO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "PRE") > 0 Then
                        Cta = IIf(mon = "E", "141102", "141101")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "PRESTAMO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    Else
                        Cta = "411501"
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "OTROS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "OTROS " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "OTROS"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    End If
                Case "AS"
                    If InStr(1, Trim(.TextMatrix(I, 6)), "VAC") > 0 Then
                        Cta = "411501"
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "VACACIONES " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "VACACIONES"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "ADEL") > 0 Or InStr(1, Trim(.TextMatrix(I, 6)), "ADL") > 0 Then
                        Cta = IIf(mon = "E", "141202", "141201")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "ADELANTO DE SUELDO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    ElseIf InStr(1, Trim(.TextMatrix(I, 6)), "PRE") > 0 Then
                        Cta = IIf(mon = "E", "141102", "141101")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "PRESTAMO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "PRESTAMO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    Else
                        Cta = IIf(mon = "E", "141202", "141201")
                        If I = 1 Then
                            If I = .Rows - 1 Then
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Nom")
                                Concepto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
                            Else
                                GlosaAct = "ADELANTO DE SUELDO " & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                                Concepto = "ADELANTO DE SUELDO"
                            End If
                        Else
                            GlosaAct = GlosaAct & "/" & DescripcionesdeCodigos("EMPLEADOABREV", caux, "Descrip")
                        End If
                    End If
            End Select
            cto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
            dh = "D"
            colcv = "00"
            If mon = "N" Then
                sol = Abs(Round(CDbl(.TextMatrix(I, 8)), 2)) '* 0.06
                dol = Abs(Round(CDbl(.TextMatrix(I, 8)) / tc, 2)) '* 0.06
            Else
                sol = Abs(Round(CDbl(.TextMatrix(I, 8)) * tc, 2)) '* 0.06
                dol = Abs(Round(CDbl(.TextMatrix(I, 8)), 2)) '* 0.06
            End If
            imp10s = imp10s + sol
            imp10d = imp10d + dol
            
            SQL = "call cn_Insert_Movi ('" & lib & "','" & tdoc & "','" & Divi & "','0000000000','" & _
                  v & "','" & Serdoc & "','" & NumDocu & "','" & cor & "','" & mon & "','" & Trim(Cta) & "','" & _
                  aux & "','" & caux & "','" & cencos & "','" & cenco & "','" & gen & "','" & _
                  cto & "'," & _
                  IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                  IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                  fec & "','" & strAnoSistema & strMesSistema & "','" & strUsuarioId & "','" & dh & "','" & _
                  colcv & "','000','')"
             oConexionMYSQL.Execute (SQL)
               
        Next
    End With
    aux = "1"
    caux = IIf(mon = "E", "00000000002", "00000000001")
    Cta = IIf(mon = "E", "104102", "104101")
    sol = imp10s
    dol = imp10d
    dh = "H"
    Divi = "013100003836" '0001 IIf(Trim(DIVI) = "0003", "0003", "0001")
    cto = Concepto
    tdoc = "9"
    Serdoc = "TW"
    NumDocu = orden
    cor = MaxCorrela(AnoMes, v)
    SQL = "call cn_Insert_Movi ('" & lib & "','" & tdoc & "','" & Divi & "','0000000000','" & _
          v & "','" & Serdoc & "','" & NumDocu & "','" & cor & "','" & mon & "','" & Trim(Cta) & "','" & _
          "1" & "','" & caux & "','" & cencos & "','" & cenco & "','" & gen & "','" & _
          cto & "'," & _
          IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
          IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
          fec & "','" & strAnoSistema & strMesSistema & "','" & strUsuarioId & "','" & dh & "','" & _
          colcv & "','000','')"
    oConexionMYSQL.Execute (SQL)
    oConexionMYSQL.Execute "Update mov_telewiese set vou='" & v & "' where identificador='" & AnoMes & orden & "'"
    If Len(GlosaAct) > 150 Then
        GlosaAct = Mid(GlosaAct, 1, 8)
    End If
    oConexionMYSQL.Execute "Update cnvouc set glosa='" & GlosaAct & "' where voucher='" & v & "' and anomes = '" & AnoMes & "'"
    Set Rs = Nothing
    MsgBox "El asiento fue generado en el voucher N° " & v & " del mes " & AnoMes, vbInformation, "NOVPeru"
    lblV.Visible = True
    lblvoucher.Visible = True
    lblvoucher.Caption = v
Exit Sub
FallaAsiento:
    Resume
    Resume Next
    SQL = "Delete from cnvouc where anomes='" & AnoMes & "' and voucher='" & v & "'"
    oConexionMYSQL.Execute SQL
    
    lblV.Visible = False
    lblvoucher.Visible = False
    MsgBox "Hubo un error al generar el voucher N° " & v & vbNewLine & "Consulte con su administrador del sistema ", vbOKOnly, "NOVPeru"
End Sub
Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub
Private Sub btnNuevo_Click()
    Dim consultaLiq As String
    ModoFormulario modNuevo
    ItemLista = -1
    ActualizaEstadoReporteEl ""
    oConexion.EjecutaInsertUpdateDelete consultaLiq, TIPO_QUERY.insertar, False
End Sub

Private Sub btnReporte_Click()
    Dim orden As String
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "PROGRAMACION DE MOVIMIENTOS DE PAGOS TELEWISSE POR N ° DE ORDEN " & meOrden.Text & " "
    oReporte.Reporte = "Rep_OrdenTeleWiesse.rpt"
    oReporte.sp_PagosTelewiesse strAnoSistema & strMesSistema & meOrden.Text
End Sub

Private Sub chameleonButton1_Click()

End Sub

Private Sub chBtnSalir_Click()
    Unload Me
End Sub

Private Sub chkCodOpcional_Click()
    If chkCodOpcional.Value = 1 Then
        txtCodOpcional.Enabled = True
        txtCodOpcional.BackColor = ColorHabilitado
        txtCodOpcional.SetFocus
    Else
        txtCodOpcional = Empty
        txtCodOpcional.Enabled = False
        txtCodOpcional.BackColor = ColorDeshabilitado
    End If
End Sub
Private Sub chkDocOpcional_Click()
    If chkDocOpcional.Value = 1 Then
        txtDocOpcional.Enabled = True
        txtDocOpcional.BackColor = ColorHabilitado
        txtDocOpcional.SetFocus
    Else
        txtDocOpcional = Empty
        txtDocOpcional.Enabled = False
        txtDocOpcional.BackColor = ColorDeshabilitado
    End If
End Sub


Private Sub cmbAuxiliares_Change()
    Me.MousePointer = vbHourglass
    If Cambiaaux And cmbAuxiliares.ListIndex >= 0 And lblModo.Caption <> "Consulta Movimiento" And lblModo.Caption <> "Acción" Then Beneficiarios lstBeneficiario
    Me.MousePointer = vbNormal
End Sub
Private Sub cmbAuxiliares_GotFocus()
    Cambiaaux = True
End Sub
Private Sub cmbAuxiliares_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtCodigo.SetFocus
    End If
End Sub

Private Sub cmbTipoOrden_Change()
    If cmbTipoOrden.ListCount <> 0 Then
        If cmbTipoOrden.List(cmbTipoOrden.ListIndex, 1) = 1 Then
            chkPagoUnico.Caption = "No Remuneraciones"
        Else
            chkPagoUnico.Caption = "Pago Unico"
        End If
    End If
End Sub
Private Sub cmbTipoPago_Change()
    If cmbTipoPago.ListCount > 0 Then
        If cmbTipoPago.ListIndex > 1 Then
        Else
            txtOficina = "000"
            txtCuentaAux = "0000000"
        End If
    End If
End Sub
Private Sub cmbTipoPago_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtCuentaAux.SetFocus
    End If
End Sub
Private Sub cmdAnexar_Click()
    If cmbTipoPago.ListIndex = 0 Then
        MsgBox "Selecione el tipo de pago del " & cmbAuxiliares.Text & vbNewLine & txtBeneficiario.Text, vbCritical + vbOKOnly, gsNomSW
        cmbTipoPago.SetFocus
        Exit Sub
    End If
    frmAnexosTelewiese.Show vbModal
End Sub


Private Sub cmdArchivos_Click()
   Set oReporte = New clsReporte
   Dim Compra As Double
   Dim Venta As Double
   oReporte.Reporte = "Rep_ControlSOX_AR_Reporte.rpt"
   oReporte.Titulo = "Orden de Pago Nro:" & meOrden.Text
   VerificaItemsGrillaPagos strAnoSistema & strMesSistema & meOrden.Text
   oReporte.sp_Rep_CtasPagarPreview strAnoSistema & strMesSistema & meOrden.Text, lblFolio.Text
End Sub


Private Sub VerificaItemsGrillaPagos(ByVal vOrden As String)
    Dim sqlorden As String
        With flexDocumentos
            For I = 1 To .Rows - 1
                If .TextMatrix(I, 0) <> "" Then
                    sqlorden = "Call Actualiza_ReporteLiqPagos (" & _
                               " '" & vOrden & "', " & _
                               " '" & .TextMatrix(I, 2) & "', " & _
                               " '" & .TextMatrix(I, 6) & "', " & _
                               " '" & .TextMatrix(I, 7) & "')"
                    oConexion.EjecutaInsertUpdateDelete sqlorden, TIPO_QUERY.insertar, False
                End If
            Next
        End With
        
        sqlorden = " Delete from liquidpagos_rep " & _
                   " where idliq = '" & vOrden & "' and VERIFICA=0; "
        oConexion.EjecutaInsertUpdateDelete sqlorden, TIPO_QUERY.insertar, False
        
        sqlorden = " Update liquidpagos_rep" & _
                   " Set VERIFICA=0" & _
                   " where idliq = '" & vOrden & "' and VERIFICA=1; "
        oConexion.EjecutaInsertUpdateDelete sqlorden, TIPO_QUERY.insertar, False
                          
End Sub

Private Sub cmdenviar_Click()
    On Error GoTo SERROR
    Dim SQL As String, ic As String, Idi As String, msj As String, Email As String, Asunto As String
    Dim RQ1 As MYSQL_RS
    Dim NomArchivo As String
    Dim FlagEnvio As String
    Dim pivotruc As String
    Dim msjRuc As String
    Dim msj1 As String
    Dim msj2 As String
    Dim orden As String
   
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "CUENTA CORRIENTE POR AUXILIAR Y DOCUMENTO """
    oReporte.Reporte = "Rep_Det_Cta_CteXdoc.rpt"

    Dim I As Integer
    Dim frmRep As frmReportPreview
    
    FlagEnvio = ""
    Asunto = "Pago de NATIONAL OILWELL VARCO PERU SRL  a  su empresa"
    msj1 = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>Para informale que el dia " & LCase(Day(CDate(mskFecha))) & " de " & LCase(NombreMes(Month(CDate(mskFecha)), False)) & _
           " La compañia National Oilwell Varco Peru SRL realizó una transferencia a su compañia </font>"
    msj2 = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'></br> Por favor comunicar una vez recibido los fondos.</br></br>Muchas Gracias." & _
           "</br></br>Saludos Cordiales,</br></br>Rosa Gallegos y Julia Ignacio </br> Dpto de Contabilidad</font>"
    
    pivotruc = flexDocumentos.TextMatrix(1, 2)
    
    msj = "<table border='1' cellpadding='0' cellspacing='0' style='color:Black; font-weight:bold; width:800px; text-align:center;font-family:Verdana; font-size:small;'>"
    msj = msj & "<tr style='background-color:#FFF6FA;><td width='100px'>Monto</td><td width='200px'>Nro Documento</td><td width='300px'>Pago Realizado</td><td width='50px'>Oficina</td><td width='150px'>Nro Cuenta</td></tr>"
                 
    With flexDocumentos
        For I = 1 To .Rows
            If .TextMatrix(I, 0) <> "" Then
                 Email = EmailContactoAuxiliar(pivotruc)
                 msjRuc = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>con  RUC " & pivotruc & ":</br></font>"
                                    
                 If (Email <> "") And (pivotruc = .TextMatrix(I, 2)) Then
                    msj = msj & "<tr style='background-color:#E6E6FA;'><td> " & IIf(.TextMatrix(I, 4) = "N", "S/.", "US$") & " " & .TextMatrix(I, 8) & "</td><td>" & .TextMatrix(I, 7) & "</td><td>" & DameTipoPago(.TextMatrix(I, 11)) & "</td><td>" & .TextMatrix(I, 12) & "</td><td>" & .TextMatrix(I, 13) & "</td></tr>  "
                 Else
                    If Email <> "" Then
                       If EnviarEmail(Email, "", "", "", msj1 & msjRuc & msj & "</table>" & msj2, Asunto, "Rosa.GallegosGallegos@nov.com;julia.ignacio@nov.com;", "") Then
                         FlagEnvio = "Enviado"
                       End If
                    End If
                    
                    pivotruc = .TextMatrix(I, 2)
                    msj = "<table border='1' cellpadding='0' cellspacing='0' style='color:Blue; font-weight:bold; width:800px; text-align:center;font-family:Verdana; font-size:small;'>"
                    msj = msj & "<tr style='background-color:Gray;><td width='100px'>Monto</td><td width='200px'>Nro Documento</td><td width='300px'>Pago Realizado</td><td width='50px'>Oficina</td><td width='150px'>Nro Cuenta</td></tr>"
                    msj = msj & "<tr style='background-color:Lime;'><td> " & IIf(.TextMatrix(I, 4) = "N", "S/.", "US$") & " " & .TextMatrix(I, 8) & "</td><td>" & .TextMatrix(I, 7) & "</td><td>" & DameTipoPago(.TextMatrix(I, 11)) & "</td><td>" & .TextMatrix(I, 12) & "</td><td>" & .TextMatrix(I, 13) & "</td></tr>  "
                 End If
            End If
        Next
        
     If Email <> "" And msj <> "" Then
        Email = EmailContactoAuxiliar(pivotruc)
        msjRuc = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>con  RUC " & pivotruc & ":</br></font>"
        If EnviarEmail(Email, "", "", "", msj1 & msjRuc & msj & "</table>" & msj2, Asunto, "Rosa.GallegosGallegos@nov.com;", "") Then
          FlagEnvio = "Enviado"
        End If
     End If
        
    End With
  
    If FlagEnvio = "Enviado" Then
       MsgBox "El Proveedor/Beneficiario ha recibido el Email de Contacto", vbInformation, "NOVPeru"
    Else
       MsgBox "El Proveedor/Beneficiario no posee ningún Email de Contacto", vbInformation, "NOVPeru"
    End If
    
    '    btnCopia_Click
    'FlgEmail = False
    Exit Sub
SERROR:
    Mensajes err.Description

End Sub
Function EmailContactoP(tcod As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    EmailContacto = ""
    SQL = "select * from contactoscliente where auxiliar = '5' and codigo = '" & tcod & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            If InStr(1, RQ.Fields("email"), "@") > 0 Then
                EmailContactoP = EmailContactoP & Trim(RQ.Fields("email")) & ";"
                'Exit Do
            End If
            RQ.MoveNext
        Loop
    End If
    Set RQ = Nothing
    
End Function

Function EmailContactoAuxiliar(tcod As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    EmailContacto = ""
    SQL = "select * from cnauxil where auxiliar = '5' and codigo = '" & tcod & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            If InStr(1, RQ.Fields("email"), "@") > 0 Then
                EmailContactoAuxiliar = EmailContactoAuxiliar & Trim(RQ.Fields("email")) & ";"
                'Exit Do
            End If
            RQ.MoveNext
        Loop
    End If
    Set RQ = Nothing
End Function


Function NombreContactoP(tcod As String) As String
    Dim SQL As String, I As Integer
    Dim RQ As MYSQL_RS
    EmailContacto = ""
    SQL = "select * from contactoscliente where auxiliar = '5' and codigo = '" & tcod & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        Do While Not RQ.EOF()
            If UCase(RQ.Fields("contacto")) <> "" Then
                NombreContactoP = NombreContactoP & IIf(I > 0, "/", "") & Trim(RQ.Fields("contacto"))
                I = I + 1
            End If
            RQ.MoveNext
        Loop
    End If
    Set RQ = Nothing
    
End Function

Private Sub cmdImpSal_Click()
    'FrmSaldosPorPagar.Show
 
     FrmSaldosPorPagar.Show vbModal
End Sub

Private Sub cmdSaldosporPagar_Click()
 FrmSaldosPorPagar.Show vbModal
End Sub

Private Sub cmdVer01_Click()
 Dim cDestino As String
 Dim SQL As String
 Dim RutaDestinoCliente As String
 Dim NombreArchivoVer As String
 
 frmArchivosAdjuntos.AnioSel = strAnoSistema
 frmArchivosAdjuntos.MesSel = strMesSistema
 
 frmArchivosAdjuntos.IdentificadorPagos = "1"
 frmArchivosAdjuntos.IdentificadorAr = strAnoSistema & strMesSistema & meOrden
 
 frmArchivosAdjuntos.Show
End Sub

Private Sub flexDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If flexDocumentos.Col = 8 And flexDocumentos.TextMatrix(flexDocumentos.row, 10) <> "5" And flexDocumentos.TextMatrix(flexDocumentos.row, 10) <> "6" Then
        Publimensaje = "modificar"
    Else
        Publimensaje = ""
    End If
End Sub



Private Sub flexDocumentos_RowColChange()
    Dim I As Integer
    Dim valor As Double
    Dim num As Integer
    Dim valorequ As Double
    Cambiaaux = False
    If flexDocumentos.Rows > 1 Then
        lblFolio = flexDocumentos.TextMatrix(flexDocumentos.row, 1)
        txtCodigo = flexDocumentos.TextMatrix(flexDocumentos.row, 2)
        txtBeneficiario = DescripcionesdeCodigos("AUXILIARES", flexDocumentos.TextMatrix(flexDocumentos.row, 2), flexDocumentos.TextMatrix(flexDocumentos.row, 10), "Descrip")
        Auxiliares cmbAuxiliares
        num = DocAnexadosxCodigo(txtCodigo)
        If num = 1 Then lblDocAnexados = str(num) & " Documento Anexado": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
        If num > 1 Then lblDocAnexados = str(num) & " Documentos Anexados": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
        
        For I = 0 To cmbAuxiliares.ListCount - 1
            If cmbAuxiliares.List(I, 1) = flexDocumentos.TextMatrix(flexDocumentos.row, 10) Then
                cmbAuxiliares.ListIndex = I
            End If
        Next
        txtOficina = flexDocumentos.TextMatrix(flexDocumentos.row, 12)
        txtCuentaAux = flexDocumentos.TextMatrix(flexDocumentos.row, 13)
        For I = 0 To cmbTipoPago.ListCount - 1
            If cmbTipoPago.List(I, 1) = CE(flexDocumentos.TextMatrix(flexDocumentos.row, 11)) Then
                cmbTipoPago.ListIndex = I
            End If
        Next
        For I = 1 To flexDocumentos.Rows - 1
            valor = valor + CDbl(CEN(flexDocumentos.TextMatrix(I, 8)))
            valorequ = valorequ + CDbl(CEN(flexDocumentos.TextMatrix(I, 9)))
        Next
        meImporte = FormatNumber(valor, 2)
        lblImpEqu = FormatNumber(valorequ, 2)
    End If
End Sub
Private Function DocAnexadosxCodigo(codigo As String) As Integer
    Dim I As Integer
    DocAnexadosxCodigo = 0
    For I = 1 To flexDocumentos.Rows - 1
        If flexDocumentos.TextMatrix(I, 2) = codigo Then
            DocAnexadosxCodigo = DocAnexadosxCodigo + 1
        End If
    Next
End Function
Private Sub Form_Activate()
    If flexDocumentos.Rows > 1 Then
        Call flexDocumentos_RowColChange
    End If
End Sub
Private Sub Form_Load()
    'Me.WindowState = vbMaximized
    DoEvents
    Publimensaje = ""
    Call WheelHook(frmMovTelewiese)
    LlenaLiqPagos cmbLiqPagos 'combo Liquid Pagos
    ModoFormulario modAccion
    Set oConsulta = New FrmConsultas
    Set oReporte = New clsReporte
End Sub
Private Sub CargarTipoOrden()
    With cmbTipoOrden
        .Clear
        .AddItem "Selecionar"
        .List(0, 1) = "0"
        .AddItem "Planilla"
        .List(1, 1) = "1"
        .AddItem "Factura"
        .List(2, 1) = "2"
        .AddItem "Varios"
        .List(3, 1) = "3"
        .ListIndex = 0
    End With
End Sub
Private Sub LlenarGrilla()
    Dim I As Integer
    With flexDocumentos
        .Clear
        .Rows = 1
        .Cols = 15
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(0) + "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 1100
        .TextMatrix(0, 1) = Space(2) + "Folio Ref"
        .ColType(1) = cadena
        .ColMaxLength(1) = 1
        .CaracteresValidos(1) = "*"
        
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = Space(2) + "Código"
        .ColType(2) = cadena
        .ColMaxLength(2) = 11
        .CaracteresValidos(2) = "*"
        
        .ColWidth(3) = 400
        .TextMatrix(0, 3) = Space(0) + "TD"
        .ColType(3) = cadena
        .ColMaxLength(3) = 2
        .CaracteresValidos(3) = "*"
        
        .ColWidth(4) = 500
        .TextMatrix(0, 4) = Space(0) + "Mon."
        .ColType(4) = cadena
        .ColMaxLength(4) = 1
        .CaracteresValidos(4) = "*"
        
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(0) + "Fec. Emisión"
        .ColType(5) = cadena
        .ColMaxLength(5) = 1
        .CaracteresValidos(5) = "*"
        
        .ColWidth(6) = 1000
        .TextMatrix(0, 6) = Space(4) + "Serie"
        .ColType(6) = cadena
        .ColMaxLength(6) = 10
        .CaracteresValidos(6) = "*"
        
        .ColWidth(7) = 1550
        .TextMatrix(0, 7) = Space(4) + "Documento"
        .ColType(7) = cadena
        .ColMaxLength(7) = 20
        .CaracteresValidos(7) = "*"
        
        .ColWidth(8) = 1000
        .TextMatrix(0, 8) = Space(4) + "Importe"
        .ColType(8) = Numero
        .ColMaxLength(8) = 15
        .ColDecimales(8) = 2
        .CaracteresValidos(8) = "0123456789."
        
        .ColWidth(9) = 1000
        .TextMatrix(0, 9) = Space(3) + "Importe Equ."
        .ColType(9) = cadena
        .ColMaxLength(9) = 1
        .CaracteresValidos(9) = "*"
        
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = Space(0) + "Aux"
        .ColType(10) = cadena
        .ColMaxLength(10) = 1
        .CaracteresValidos(10) = "*"
        
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = Space(0) + "tipopago"
        .ColType(11) = cadena
        .ColMaxLength(11) = 1
        .CaracteresValidos(11) = "*"
        
        .ColWidth(12) = 0
        .TextMatrix(0, 12) = Space(0) + "oficina"
        .ColType(12) = cadena
        .ColMaxLength(12) = 1
        .CaracteresValidos(12) = "*"
        
        .ColWidth(13) = 0
        .TextMatrix(0, 13) = Space(0) + "cuenta"
        .ColType(13) = cadena
        .ColMaxLength(13) = 1
        .CaracteresValidos(13) = "*"
            
        .ColWidth(14) = 0
        .ColType(14) = cadena
        .ColMaxLength(14) = 4
        .CaracteresValidos(14) = "1234567890"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim resp As Integer
    resp = MsgBox("¿Desea salir del formulario?", vbYesNo + vbQuestion, gsNomSW)
    If resp = vbNo Then
        Cancel = 1
    Else
        Set oConsulta = Nothing
        Set oReporte = Nothing
    End If
    WheelUnHook
End Sub

Private Sub lblFolio_GotFocus()
    mark1 lblFolio
End Sub

Private Sub lblFolio_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 And (lblModo = "Nuevo Movimiento" Or lblModo = "Modificar Movimiento") Then
        If Trim(lblFolio) <> "" Then
            If Len(lblFolio) < 5 Then
                If BuscarFolioRef(strAnoSistema & strMesSistema & Right("0000" & Trim(lblFolio), 4)) Then
                    Call cmbMoneda_Change
                Else
                    MsgBox "No se encuentra el folio de referencia", vbOKOnly + vbInformation, gsNomSW
                End If
            Else
                If BuscarFolioRef(Trim(lblFolio)) Then
                    Call cmbMoneda_Change
                Else
                    MsgBox "No se encuentra el folio de referencia", vbOKOnly + vbInformation, gsNomSW
                End If
            End If
       Else
            lblFolio.SetFocus
       End If
    End If
    If KeyCode = 13 And lblModo = "Nuevo Movimiento" Then
       lblFolio.SetFocus
    End If
End Sub

Public Function BuscarFolioRef(folio As String) As Boolean
    Dim I As Integer
    Dim SQL As String
    Dim rsFolRef As MYSQL_RS
    BuscarFolioRef = False
    SQL = "Select auxiliar,codigo,mon,cod_tipo_doc from doc_prog where identificador='" & folio & "'"
    Set rsFolRef = oConexion.EjecutaSelectRS(SQL)
    If rsFolRef.RecordCount = 1 Then
        Dim RQ As MYSQL_RS
        SQL = "select DESCRIP,Cod_Fam from cndocum C where CODDOC='" & rsFolRef.Fields("cod_tipo_doc") & "' " & _
              "AND (PROTEGIDO = 'N' OR (SELECT PERMISO FROM docsusuario D WHERE D.CODDOC=C.CODDOC " & _
              "AND USUARIO = '" & strUsuarioId & "')=1)"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If RQ.EOF() Then
            MsgBox "No se encuentra autorizado para visualizar este folio", vbInformation, "NOVPeru"
            BuscarFolioRef = False
            Exit Function
        End If
        
        Set RQ = Nothing
        For I = 0 To cmbAuxiliares.ListCount - 1
            If CE(rsFolRef.Fields("auxiliar")) = cmbAuxiliares.List(I, 1) Then
                cmbAuxiliares.ListIndex = I
            End If
        Next
        txtCodigo = CE(rsFolRef.Fields("codigo"))
        txtBeneficiario = DescripcionesdeCodigos("AUXILIARES", CE(rsFolRef.Fields("codigo")), CE(rsFolRef.Fields("auxiliar")), "Descrip")
        If lblModo <> "Modificar Movimiento" Then lblModo = "Nuevo Movimiento"
         For I = 0 To cmbMoneda.ListCount - 1
            If CE(rsFolRef.Fields("mon")) = cmbMoneda.List(I, 1) Then
                cmbMoneda.ListIndex = I
            End If
        Next
        BuscarFolioRef = True
    End If
    rsFolRef.CloseRecordset
    Set rsFolRef = Nothing
End Function
Private Sub meImporte_Change()
    If lblTCambio <> Empty Then
        If cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
            If CDbl(lblTCambio) > 0 Then
                lblImpEqu = FormatNumber(CDbl(meImporte) / CDbl(lblTCambio))
            Else
                lblImpEqu = "0.00"
            End If
        Else
            lblImpEqu = FormatNumber(CDbl(meImporte) * CDbl(lblTCambio))
        End If
    End If
End Sub
Private Sub meOrden_GotFocus()
    mark1 meOrden
End Sub
Private Sub meOrden_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 And lblModo = "Acción" Then
       If Trim(meOrden) <> "" Then
            If BuscarOrden(strAnoSistema & strMesSistema & Right("0000" & Trim(meOrden), 4)) Then
                ModoFormulario modConsulta
                flexDocumentos.SetFocus
                Call flexDocumentos_RowColChange
            Else
                If FlgTw = False Then
                    MsgBox "No se encuentra el movimiento", vbOKOnly + vbInformation, gsNomSW
                Else
                    FlgTw = False
                End If
                ModoFormulario modAccion
            End If
       Else
            lblFolio.SetFocus
       End If
    End If
    If KeyCode = 13 And lblModo = "Nuevo Movimiento" Then
       lblFolio.SetFocus
    End If
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 5
            .pCol = 0: .pAnchoCol = 450
            .pCol = 1: .pAnchoCol = 800
            .pCol = 2: .pAnchoCol = 1200
            .pCol = 3: .pAnchoCol = 1200
            .pCol = 4: .pAnchoCol = 800
            .pTitulo = "Ordenes Telewiese del Mes de " & strMesSistema
            .pForm = FORM_TELEWIESE
            .pCaso = LABEL_ORDENTW
            .Show
        End With
    End If
End Sub
Private Sub mskFecha_GotFocus()
    mark1 mskFecha
End Sub
Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbAuxiliares.SetFocus
    End If
End Sub

Private Sub txtBeneficiario_Change()
    If Trim(txtBeneficiario) = "" Then txtCodigo = ""
End Sub
Private Sub txtBeneficiario_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim BENEFICIARIO As String
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 4
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3500
            .pCol = 2: .pAnchoCol = 500
            .pCol = 3: .pAnchoCol = 1200
            .pTitulo = "Lista de " & cmbAuxiliares.Text
            .pForm = FORM_TELEWIESE
            .pCaso = Label_Descrip_Auxil
            .Show
        End With
    End If
    If KeyCode.Value = vbKeyDown Then
        txtBeneficiario.SelStart = 0
        txtBeneficiario.SelLength = 0
        Call keybd_event(vbKeyLeft, 0, 0, 0)
        lstBeneficiario.Visible = True
        lstBeneficiario.SetFocus
        lstBeneficiario.ListIndex = ItemLista
    End If
    If KeyCode.Value = 13 Then
        If ItemLista <> -1 Then
            txtCodigo.Text = lstBeneficiario.List(ItemLista, 1)
        End If
        cmbMoneda.SetFocus
    End If
End Sub
Private Sub txtBeneficiario_KeyPress(KeyAscii As MSForms.ReturnInteger)
    AutoComplete txtBeneficiario, KeyAscii, lstBeneficiario
End Sub
Private Sub Beneficiarios(a As MSForms.ListBox)
    Dim SQL As String
    Dim I As Integer
    Dim rsBen As MYSQL_RS
    SQL = "auxil where auxiliar='" & cmbAuxiliares.List(cmbAuxiliares.ListIndex, 1) & "' order by descrip"
    Set rsBen = oConexion.EjecutaSelect(SQL)
    If rsBen.RecordCount = 0 Then Exit Sub
    a.Clear
    Do While Not rsBen.EOF
        a.AddItem CE(rsBen.Fields("descrip"))
        a.List(I, 1) = CE(rsBen.Fields("codigo"))
        a.List(I, 2) = CE(rsBen.Fields("ruc"))
        I = I + 1
        rsBen.MoveNext
    Loop
    Set rsBen = Nothing
End Sub
Private Sub cmbCtaCte_Change()
    If cmbCtaCte.ListIndex = 0 And lblModo.Caption <> "Consulta Movimiento" Then
        meImporte.Enabled = False
        meImporte.BackColor = ColorDeshabilitado
        cmdAnexar.Enabled = False
    End If
    If cmbCtaCte.ListIndex > 0 And lblModo.Caption <> "Consulta Movimiento" Then
       If lblModo = "Nuevo Movimiento" Or lblModo = "Modificar Movimiento" Then cmdAnexar.Enabled = True
       meImporte.Enabled = True
       meImporte.BackColor = ColorHabilitado
    End If
End Sub
Private Sub cmbCtaCte_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        If cmbCtaCte.ListIndex > 0 Then
            cmbTipoPago.SetFocus
        End If
    End If
End Sub
Private Sub cmbMoneda_Change()
    Dim I As Integer
    If cmbMoneda.ListIndex = 0 Then
        If lblModo.Caption <> "Consulta Movimiento" Then
            cmbCtaCte.Enabled = False
            cmbCtaCte.BackColor = ColorDeshabilitado
        End If
        lblMoneda = ""
        lblMonEqu = ""
        txtCuentaAux = "0000000"
        txtOficina = "000"
        cmbCtaCte.Clear
        Exit Sub
    End If
    If cmbMoneda.ListIndex > 0 Then
        If lblModo.Caption <> "Consulta Movimiento" Then
            cmbCtaCte.Enabled = True
            cmbCtaCte.BackColor = ColorHabilitado
        End If
        CtasCtes cmbCtaCte
        If cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
            lblMoneda = "S/."
            lblMonEqu = "US$"
            For I = 0 To cmbTipoPago.ListCount - 1
                If cmbTipoPago.List(I, 1) = tipcta_mn Then
                    cmbTipoPago.ListIndex = I
                End If
            Next
            txtCuentaAux = IIf(numcta_mn = "", "0000000", numcta_mn)
            txtOficina = IIf(Left(numcta_mn, 3) = "", "000", Left(numcta_mn, 3))
        Else
            lblMoneda = "US$"
            lblMonEqu = "S/."
            For I = 0 To cmbTipoPago.ListCount - 1
                If cmbTipoPago.List(I, 1) = tipcta_me Then
                    cmbTipoPago.ListIndex = I
                End If
            Next
            txtCuentaAux = IIf(numcta_me = "", "0000000", numcta_me)
            txtOficina = IIf(Left(numcta_me, 3) = "", "000", Left(numcta_me, 3))
        End If
    End If
End Sub
Private Sub cmbMoneda_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        If cmbMoneda.ListIndex > 0 Then
            cmbCtaCte.SetFocus
        End If
    End If
End Sub
Private Sub Auxiliares(a As MSForms.ComboBox)
    Dim I As Integer
    Dim SQL As String
    Dim rsaux As MYSQL_RS
    SQL = "aux_pagos order by descripcion"
    Set rsaux = oConexion.EjecutaSelect(SQL)
    If rsaux.RecordCount = 0 Then Exit Sub
    I = 0
    a.Clear
    Do While Not rsaux.EOF
        a.AddItem CE(rsaux.Fields("descripcion"))
        a.List(I, 1) = CE(rsaux.Fields("aux"))
        I = I + 1
        rsaux.MoveNext
    Loop
    Set rsaux = Nothing
    If a.ListCount > 0 Then a.ListIndex = 0
End Sub
Private Sub CtasCtes(a As MSForms.ComboBox)
    Dim I As Integer
    Dim SQL As String
    Dim rsCtasCtes As MYSQL_RS
    SQL = "auxil where auxiliar='1' and moneda='" & cmbMoneda.List(cmbMoneda.ListIndex, 1) & "' and afiliada='S' order by descrip"
    Set rsCtasCtes = oConexion.EjecutaSelect(SQL)
    If rsCtasCtes.RecordCount = 0 Then Exit Sub
    a.Clear
    a.AddItem "Seleccionar..."
    I = 1
    Do While Not rsCtasCtes.EOF
        a.AddItem CE(rsCtasCtes.Fields("descrip"))
        a.List(I, 1) = CE(rsCtasCtes.Fields("codigo"))
        I = I + 1
        rsCtasCtes.MoveNext
    Loop
    Set rsCtasCtes = Nothing
    If a.ListCount > 0 Then a.ListIndex = 1
End Sub
Private Sub TipPago(a As MSForms.ComboBox)
    Dim I As Integer
    Dim SQL As String
    Dim rsTipPago As MYSQL_RS
    SQL = "TipoPagos where codpago<>0 order by codpago "
    Set rsTipPago = oConexion.EjecutaSelect(SQL)
    If rsTipPago.RecordCount = 0 Then Exit Sub
    a.Clear
    a.AddItem "Seleccionar..."
    I = 1
    Do While Not rsTipPago.EOF
        a.AddItem CE(rsTipPago.Fields("descrip"))
        a.List(I, 1) = CE(rsTipPago.Fields("codpago"))
        I = I + 1
        rsTipPago.MoveNext
    Loop
    Set rsTipPago = Nothing
    If a.ListCount > 0 Then a.ListIndex = 0
End Sub
Private Sub moneda(a As MSForms.ComboBox)
    a.Clear
    a.AddItem "Seleccionar..."
    a.List(0, 1) = "0"
    a.AddItem "Nacional"
    a.List(1, 1) = "N"
    a.AddItem "Extranjera"
    a.List(2, 1) = "E"
    a.ListIndex = 0
    If a.ListCount > 0 Then a.ListIndex = 0
End Sub
Private Sub txtCodigo_Change()
    If Len(txtCodigo) = 11 Then
        BuscarCuentaOficina cmbAuxiliares.List(cmbAuxiliares.ListIndex, 1), txtCodigo
    End If
    ItemLista = -1
End Sub
Private Sub txtCodigo_GotFocus()
    mark txtCodigo
End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 4
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3500
            .pCol = 2: .pAnchoCol = 500
            .pCol = 3: .pAnchoCol = 1200
            .pTitulo = "Lista de " & cmbAuxiliares.Text
            .pForm = FORM_TELEWIESE
            .pCaso = Label_Descrip_Auxil
            .Show
        End With
    End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        txtBeneficiario = DescripcionesdeCodigos("AUXILIARES", txtCodigo, cmbAuxiliares.List(cmbAuxiliares.ListIndex, 1), "Descrip")
        txtBeneficiario.SelStart = 0
        txtBeneficiario.SelLength = Len(txtBeneficiario)
        txtBeneficiario.SetFocus
    End If
End Sub
Private Sub lblFolio_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii = 8) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 47 Then
        KeyAscii = KeyAscii
    Else
       KeyAscii = 0
    End If
End Sub
Private Sub lstBeneficiario_Change()
    ItemLista = lstBeneficiario.ListIndex
End Sub
Private Sub lstBeneficiario_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     If KeyCode.Value = 13 Then
        txtBeneficiario.Text = lstBeneficiario.List(lstBeneficiario.ListIndex, 0)
        If cmbAuxiliares.List(cmbAuxiliares.ListIndex, 1) = "5" Then
            txtCodigo.Text = lstBeneficiario.List(lstBeneficiario.ListIndex, 2)
        Else
            txtCodigo.Text = lstBeneficiario.List(lstBeneficiario.ListIndex, 1)
        End If
        mark1 txtBeneficiario
        lstBeneficiario.Visible = False
        txtBeneficiario.SetFocus
    End If
    If KeyCode.Value = 27 Then
        If lstBeneficiario.Visible = True Then
            lstBeneficiario.Visible = False
            mark1 txtBeneficiario
            txtBeneficiario.SetFocus
        End If
    End If
End Sub
Private Sub meImporte_GotFocus()
    mark1 meImporte
End Sub
Public Sub LimpiarDatos()
    meOrden = Empty
    lblFolio = Empty
    mskFecha = "__/__/____"
    txtBeneficiario = Empty
    txtCodigo = Empty
    meImporte = "0.00"
    lblImpEqu = "0.00"
    txtConcepto = Empty
    cmbAuxiliares.Clear
    cmbMoneda.Clear
    cmbCtaCte.Clear
    lblMoneda = Empty
    lblMonEqu = Empty
    lblTCambio = Empty
    lblEstado.Visible = False
    lblDocAnexados = Empty
    txtCuentaAux = Empty
    txtDocOpcional = Empty
    txtCodOpcional = Empty
End Sub
Public Sub BloqueoControles(valor As Boolean)
    meOrden.Locked = valor
    lblFolio.Locked = valor
    mskFecha.Enabled = Not (valor)
    txtBeneficiario.Locked = valor
    txtCodigo.Locked = valor
    meImporte.Locked = valor
    txtConcepto.Locked = valor
    cmbAuxiliares.Locked = valor
    cmbMoneda.Locked = valor
    cmbCtaCte.Locked = valor
    
    cmbTipoPago.Locked = True
   
    txtOficina.Locked = True
    txtCuentaAux.Locked = True
    
    If valor = True Then
        meOrden.BackColor = ColorDeshabilitado
        lblFolio.BackColor = ColorDeshabilitado
        mskFecha.BackColor = ColorDeshabilitado
        txtBeneficiario.BackColor = ColorDeshabilitado
        txtCodigo.BackColor = ColorDeshabilitado
        meImporte.BackColor = ColorDeshabilitado
        txtConcepto.BackColor = ColorDeshabilitado
        cmbAuxiliares.BackColor = ColorDeshabilitado
        cmbMoneda.BackColor = ColorDeshabilitado
        cmbCtaCte.BackColor = ColorDeshabilitado
        cmbTipoPago.BackColor = ColorDeshabilitado
        txtOficina.BackColor = ColorDeshabilitado
        txtCuentaAux.BackColor = ColorDeshabilitado
        cmbLiqPagos.BackColor = ColorDeshabilitado
    Else
        meOrden.BackColor = ColorHabilitado
        lblFolio.BackColor = ColorHabilitado
        mskFecha.BackColor = ColorHabilitado
        txtBeneficiario.BackColor = ColorHabilitado
        txtCodigo.BackColor = ColorHabilitado
        meImporte.BackColor = ColorHabilitado
        txtConcepto.BackColor = ColorHabilitado
        cmbAuxiliares.BackColor = ColorHabilitado
        cmbMoneda.BackColor = ColorHabilitado
        cmbCtaCte.BackColor = ColorHabilitado
        cmbTipoPago.BackColor = ColorHabilitado
        txtOficina.BackColor = ColorHabilitado
        txtCuentaAux.BackColor = ColorHabilitado
        cmbLiqPagos.BackColor = ColorHabilitado
    End If
    
End Sub
Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
            LimpiarDatos
            lblModo = "Acción"
            BloqueoControles True
            meOrden.Locked = False
            meOrden.BackColor = ColorHabilitado
            BtnNuevo.Enabled = True
            lblDocAnexados.Visible = False
            lblV.Visible = False
            lblvoucher.Visible = False
            lblvoucher.Caption = ""
            cmdAnexar.Enabled = False
            chkCodOpcional.Enabled = False
            chkPagoUnico.Enabled = False
            txtCodOpcional.BackColor = ColorDeshabilitado
            txtCodOpcional.Locked = True
            chkDocOpcional.Enabled = False
            txtDocOpcional.BackColor = ColorDeshabilitado
            txtDocOpcional.Locked = True
            chkPreparar.Enabled = False
            cmbTipoOrden.Clear
            cmbTipoOrden.BackColor = ColorDeshabilitado
            cmbTipoOrden.Locked = True
            LlenarGrilla
        Case ModoForm.modNuevo
            LimpiarDatos
            lblTCambio.Caption = FormatNumber(str(TCambio), 2)
            lblModo = "Nuevo Movimiento"
            lblV.Visible = False
            lblvoucher.Visible = False
            lblvoucher.Caption = ""
            BloqueoControles False
            If (Len(CStr(Date)) < 10) Then
             mskFecha.Text = "0" & CStr(Date)
            Else
             mskFecha.Text = CStr(Date)
            End If
            Auxiliares cmbAuxiliares
            moneda cmbMoneda
            TipPago cmbTipoPago
            meOrden = Right(GenerarOrden(strAnoSistema & strMesSistema), 4)
            strIdentificador = GenerarOrden(strAnoSistema & strMesSistema)
            lblDocAnexados.Visible = False
            meOrden.Locked = True
            meOrden.BackColor = ColorDeshabilitado
            ConfigurarBotones cfgNuevo
            meOrden.SetFocus
            LlenarGrilla
        Case ModoForm.modConsulta
            lblModo = "Consulta Movimiento"
            BloqueoControles True
            ConfigurarBotones cfgGrabar
        Case ModoForm.modEditar
            lblModo = "Modificar Movimiento"
            BloqueoControles False
            ConfigurarBotones cfgModificar
    End Select
    Call chkDocOpcional_Click
    Call chkCodOpcional_Click
End Sub
Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Call chkDocOpcional_Click
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnInterfaz.Enabled = False
            btnAnulaTw.Enabled = False
            btnReporte.Enabled = False
            chkCodOpcional.Enabled = False
            chkCodOpcional.Value = 0
            chkPagoUnico.Enabled = False
            chkPagoUnico.Value = 0
            txtCodOpcional.BackColor = ColorDeshabilitado
            txtCodOpcional.Locked = True
            chkDocOpcional.Enabled = False
            chkDocOpcional.Value = 0
            txtDocOpcional.BackColor = ColorDeshabilitado
            txtDocOpcional.Locked = True
            chkPreparar.Enabled = False
            chkPreparar.Value = 0
            cmbTipoOrden.Clear
            cmbTipoOrden.BackColor = ColorDeshabilitado
            cmbTipoOrden.Locked = True
            BtnCancelar.Enabled = True
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnInterfaz.Enabled = False
            btnAnulaTw.Enabled = False
            btnReporte.Enabled = False
            chkCodOpcional.Enabled = False
            chkCodOpcional.Value = 0
            chkPagoUnico.Enabled = False
            chkPagoUnico.Value = 0
            txtCodOpcional.BackColor = ColorDeshabilitado
            txtCodOpcional.Locked = True
            chkDocOpcional.Enabled = False
            chkDocOpcional.Value = 0
            txtDocOpcional.BackColor = ColorDeshabilitado
            txtDocOpcional.Locked = True
            chkPreparar.Enabled = False
            chkPreparar.Value = 0
            cmbTipoOrden.Clear
            cmbTipoOrden.BackColor = ColorDeshabilitado
            cmbTipoOrden.Locked = True
            BtnCancelar.Enabled = True
            cmdAnexar.Enabled = True
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnInterfaz.Enabled = False
            btnAnulaTw.Enabled = False
            btnReporte.Enabled = False
            chkCodOpcional.Enabled = False
            chkCodOpcional.Value = 0
            chkPagoUnico.Enabled = False
            chkPagoUnico.Value = 0
            txtCodOpcional.BackColor = ColorDeshabilitado
            txtCodOpcional.Locked = True
            chkDocOpcional.Enabled = False
            chkDocOpcional.Value = 0
            txtDocOpcional.BackColor = ColorDeshabilitado
            txtDocOpcional.Locked = True
            chkPreparar.Enabled = False
            chkPreparar.Value = 0
            cmbTipoOrden.Clear
            cmbTipoOrden.BackColor = ColorDeshabilitado
            cmbTipoOrden.Locked = True
            BtnCancelar.Enabled = False
        Case ConfigBotones.cfgAnular
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = False
            BtnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnInterfaz.Enabled = False
            btnAnulaTw.Enabled = False
            chkCodOpcional.Enabled = False
            chkCodOpcional.Value = 0
            chkPagoUnico.Enabled = False
            chkPagoUnico.Value = 0
            txtCodOpcional.BackColor = ColorDeshabilitado
            txtCodOpcional.Locked = True
            chkDocOpcional.Enabled = False
            chkDocOpcional.Value = 0
            txtDocOpcional.BackColor = ColorDeshabilitado
            txtDocOpcional.Locked = True
            chkPreparar.Enabled = False
            chkPreparar.Value = 0
            cmbTipoOrden.Clear
            cmbTipoOrden.BackColor = ColorDeshabilitado
            cmbTipoOrden.Locked = True
            btnReporte.Enabled = False
            BtnCancelar.Enabled = False
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            If lblEstado.tag = EMITIDO Then
                BtnModificar.Enabled = True
                BtnEliminar.Enabled = True
            Else
                BtnModificar.Enabled = False
                BtnEliminar.Enabled = False
            End If
            If lblEstado.tag = ANULADO Then
                btnAnulaTw.Enabled = False
                btnInterfaz.Enabled = False
                chkCodOpcional.Enabled = False
                chkPagoUnico.Enabled = False
                txtCodOpcional.BackColor = ColorDeshabilitado
                txtCodOpcional.Locked = True
                chkDocOpcional.Enabled = False
                txtDocOpcional.BackColor = ColorDeshabilitado
                txtDocOpcional.Locked = True
                chkPreparar.Enabled = False
                cmbTipoOrden.Clear
                cmbTipoOrden.BackColor = ColorDeshabilitado
                cmbTipoOrden.Locked = True
            End If
            If lblEstado.tag = ELIMINADO Then
                btnAnulaTw.Enabled = False
                btnInterfaz.Enabled = False
                chkCodOpcional.Enabled = False
                chkPagoUnico.Enabled = False
                txtCodOpcional.BackColor = ColorDeshabilitado
                txtCodOpcional.Locked = True
                chkDocOpcional.Enabled = False
                txtDocOpcional.BackColor = ColorDeshabilitado
                txtDocOpcional.Locked = True
                chkPreparar.Enabled = False
                cmbTipoOrden.Clear
                cmbTipoOrden.BackColor = ColorDeshabilitado
                cmbTipoOrden.Locked = True
            End If
            If lblEstado.tag = TRANSFERIDO Then
                btnAnulaTw.Enabled = True
                btnInterfaz.Enabled = False
                chkCodOpcional.Enabled = False
                chkCodOpcional.Value = 0
                chkPagoUnico.Enabled = False
                chkPagoUnico.Value = 0
                txtCodOpcional.BackColor = ColorDeshabilitado
                txtCodOpcional.Locked = True
                chkDocOpcional.Enabled = False
                chkDocOpcional.Value = 0
                txtDocOpcional.BackColor = ColorDeshabilitado
                txtDocOpcional.Locked = True
                chkPreparar.Enabled = False
                chkPreparar.Value = 0
                cmbTipoOrden.Clear
                cmbTipoOrden.BackColor = ColorDeshabilitado
                cmbTipoOrden.Locked = True
            Else
                If lblEstado.tag <> ELIMINADO And lblEstado.tag <> ANULADO Then
                    btnAnulaTw.Enabled = False
                    btnInterfaz.Enabled = True
                    chkCodOpcional.Enabled = True
                    chkPagoUnico.Enabled = True
                    txtCodOpcional.BackColor = ColorHabilitado
                    txtCodOpcional.Locked = False
                    chkDocOpcional.Enabled = True
                    txtDocOpcional.BackColor = ColorHabilitado
                    txtDocOpcional.Locked = False
                    chkPreparar.Enabled = True
                    CargarTipoOrden
                    cmbTipoOrden.Locked = False
                End If
            End If
            btnGrabar.Enabled = False
            btnReporte.Enabled = True
            BtnCancelar.Enabled = True
            cmdAnexar.Enabled = False
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo Movimiento", "Consulta Movimiento"
                     ModoFormulario modAccion
                     BtnNuevo.Enabled = True
                     BtnModificar.Enabled = False
                     BtnEliminar.Enabled = False
                     btnGrabar.Enabled = False
                     btnInterfaz.Enabled = False
                     btnAnulaTw.Enabled = False
                     btnReporte.Enabled = False
                     BtnCancelar.Enabled = False
                     chkCodOpcional.Enabled = False
                     chkCodOpcional.Value = 0
                     chkPagoUnico.Enabled = False
                     chkPagoUnico.Value = 0
                     txtCodOpcional.BackColor = ColorDeshabilitado
                     txtCodOpcional.Locked = True
                     chkDocOpcional.Enabled = False
                     chkDocOpcional.Value = 0
                     txtDocOpcional.BackColor = ColorDeshabilitado
                     txtDocOpcional.Locked = True
                     chkPreparar.Enabled = False
                     chkPreparar.Value = 0
                     cmbTipoOrden.Clear
                     cmbTipoOrden.BackColor = ColorDeshabilitado
                     cmbTipoOrden.Locked = True
                     meOrden.SetFocus
                     lblV.Visible = False
                     lblvoucher.Visible = False
                Case "Modificar Movimiento"
                    ModoFormulario modConsulta
            End Select
    End Select
End Sub
Public Function BuscarOrden(orden As String) As Boolean
    Dim sqlorden As String
    Dim rsorden As MYSQL_RS
    BuscarOrden = False
    If orden <> "" Then
        sqlorden = "Select * from mov_telewiese" & _
                    " where identificador='" & orden & "' order by documento"
        Set rsorden = oConexion.EjecutaSelectRS(sqlorden)
        If rsorden.RecordCount > 0 Then
            If Trim(rsorden.Fields("tipdoc")) <> "" Then
                Dim RQ As MYSQL_RS
                sqlorden = "select descrip,Cod_Fam from cndocum C where CODDOC='" & rsorden.Fields("tipdoc") & "' " & _
                      "AND (protegido = 'N' OR (SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc " & _
                      "AND usuario = '" & strUsuarioId & "')=1)"
                Set RQ = oConexion.EjecutaSelectRS(sqlorden)
                If RQ.EOF() Then
                    FlgTw = True
                    MsgBox "No se encuentra autorizado para visualizar esta Orden", vbInformation, "NOVPeru"
                    BuscarOrden = False
                    Exit Function
                End If
                Set RQ = Nothing
            End If
            CargarDatos rsorden
            BuscarOrden = True
            Exit Function
        End If
    End If
End Function

Public Sub CargarDatos(rsorden As MYSQL_RS)
    Dim I As Integer, num As Integer
    With rsorden
        .MoveFirst
        Do While Not .EOF
            If val(.Fields("item")) = 1 Then
                meOrden = Right(CE(.Fields("identificador")), 4)
                strIdentificador = CE(.Fields("identificador"))
                lblEstado = DescripcionesdeCodigos("DOC_ESTADO", CE(.Fields("estado")))
                lblEstado.tag = CE(.Fields("estado"))
                If CE(.Fields("voucher")) <> "" Then
                    lblV.Visible = True
                    lblvoucher.Visible = True
                    lblvoucher.Caption = CE(.Fields("voucher"))
                Else
                    lblV.Visible = False
                    lblvoucher.Visible = False
                End If
                If .RecordCount = 1 Then
                    lblFolio = CE(.Fields("folioref"))
                End If
                If CE(.Fields("fecha") <> "") Then mskFecha = Format(CE(.Fields("fecha")), "dd/mm/yyyy")
                Auxiliares cmbAuxiliares
                For I = 0 To cmbAuxiliares.ListCount - 1
                    If cmbAuxiliares.List(I, 1) = CE(.Fields("Auxiliar")) Then
                        cmbAuxiliares.ListIndex = I
                    End If
                Next
                txtCodigo = CE(.Fields("Codigo"))
                txtBeneficiario = DescripcionesdeCodigos("AUXILIARES", CE(.Fields("Codigo")), CE(.Fields("Auxiliar")), "Descrip")
                txtOficina = CE(.Fields("Oficina"))
                txtCuentaAux = CE(.Fields("ctaaux"))
                moneda cmbMoneda
                For I = 0 To cmbMoneda.ListCount - 1
                    If cmbMoneda.List(I, 1) = CE(.Fields("Moneda")) Then
                        cmbMoneda.ListIndex = I
                    End If
                Next
                For I = 0 To cmbCtaCte.ListCount - 1
                    If cmbCtaCte.List(I, 1) = CE(.Fields("ctacte")) Then
                        cmbCtaCte.ListIndex = I
                    End If
                Next
                TipPago cmbTipoPago
                For I = 0 To cmbTipoPago.ListCount - 1
                    If cmbTipoPago.List(I, 1) = CE(.Fields("tipopago")) Then
                        cmbTipoPago.ListIndex = I
                    End If
                Next
                meImporte = str(CDbl(meImporte) + CDbl(FormatNumber(CEN(.Fields("importe")), 2)))
                txtConcepto = CE(.Fields("OBS"))
            End If
            flexDocumentos.Rows = flexDocumentos.Rows + 1
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 0) = Left(CE(rsorden.Fields("item")), 6)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 1) = CE(.Fields("folioref"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 2) = CE(.Fields("codigo"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 3) = CE(.Fields("tipdoc"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 4) = CE(.Fields("moneda"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 5) = CE(.Fields("fec_emi"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 6) = CE(.Fields("serie"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 7) = CE(.Fields("documento"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 8) = FormatNumber(CEN(.Fields("importe")), 2)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 9) = FormatNumber(CEN(.Fields("importeequ")), 2)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 10) = CE(.Fields("auxiliar"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 11) = CE(.Fields("tipopago"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 12) = CE(.Fields("oficina"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 13) = CE(.Fields("ctaaux"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 14) = CE(.Fields("DIVI"))
            lblimporte = str(CDbl(lblimporte) + CEN(.Fields("importe")))
            lblImpEqu = str(CDbl(lblImpEqu) + CEN(rsorden.Fields("importeequ")))
            num = num + 1
            rsorden.MoveNext
        Loop
    End With
    lblTCambio = TCambio(mskFecha)
    If num = 1 Then lblDocAnexados = str(num) & " Documento Anexado": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
    If num > 1 Then lblDocAnexados = str(num) & " Documentos Anexados": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
    lblEstado.Visible = True
End Sub


Public Sub BuscarCuentaOficina(aux As String, Cod As String)
    Dim SQL As String
    Dim rsBusca As MYSQL_RS
    Dim RQ As MYSQL_RS
    Dim Flg As Boolean
    Dim NumFolio As String, StrAnio As String, strMes As String
    If aux = 3 Then
        SQL = "Select tipcta_mn,numcta_mn,tipcta_me,numcta_me from empleado" & _
              " where codigo='" & Cod & "'"
    ElseIf aux = 5 Then
        If Len(lblFolio) < 5 Then
            StrAnio = strAnoSistema
            strMes = strMesSistema
        Else
            StrAnio = Mid(lblFolio, 1, 4)
            strMes = Mid(lblFolio, 5, 2)
        End If
        SQL = "select orden from documento_contables where identificador = '" & StrAnio & strMes & Right("0000" & Trim(lblFolio), 4) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            If Trim(RQ.Fields("orden")) <> "" Then Flg = True
        End If
        If Flg = True Then
            SQL = "SELECT o.MPago,o.CtaBco,d.mon From orden_compra O INNER JOIN documento_contables D " & _
                  "ON (o.Correl=d.Orden) Where (d.Identificador = '" & StrAnio & strMes & Right("0000" & Trim(lblFolio), 4) & "')"
        Else
            SQL = "Select tipcta_mn,numcta_mn,tipcta_me,numcta_me from cnauxil" & _
                  " where auxiliar ='" & aux & "' and codigo='" & Cod & "'"
        End If
    Else
        SQL = "Select tipcta_mn,numcta_mn,tipcta_me,numcta_me from cnauxil" & _
              " where auxiliar ='" & aux & "' and codigo='" & Cod & "'"
    End If
    
    Set rsBusca = oConexion.EjecutaSelectRS(SQL)
    If rsBusca.RecordCount > 0 Then
        If aux = 5 Then
            If Flg = True Then
                tipcta_mn = "": numcta_mn = ""
                tipcta_me = "": numcta_me = ""
                If rsBusca.Fields("mon") = "N" Then
                    tipcta_mn = TipoPago(Trim(CE(rsBusca.Fields("mpago"))))
                    numcta_mn = TipoCuenta(Trim(CE(rsBusca.Fields("ctabco"))))
                Else
                    tipcta_me = TipoPago(Trim(CE(rsBusca.Fields("mpago"))))
                    numcta_me = TipoCuenta(Trim(CE(rsBusca.Fields("ctabco"))))
                End If
            Else
                tipcta_mn = Trim(CE(rsBusca.Fields("tipcta_mn")))
                numcta_mn = Trim(CE(rsBusca.Fields("numcta_mn")))
                tipcta_me = Trim(CE(rsBusca.Fields("tipcta_me")))
                numcta_me = Trim(CE(rsBusca.Fields("numcta_me")))
            End If
        Else
            tipcta_mn = Trim(CE(rsBusca.Fields("tipcta_mn")))
            numcta_mn = Trim(CE(rsBusca.Fields("numcta_mn")))
            tipcta_me = Trim(CE(rsBusca.Fields("tipcta_me")))
            numcta_me = Trim(CE(rsBusca.Fields("numcta_me")))
        End If
    Else
        If aux = 3 Then
            MsgBox "No existe el código del Empleado", vbOKOnly + vbInformation, gsNomSW
        Else
            MsgBox "No existe el código de Auxiliar", vbOKOnly + vbInformation, gsNomSW
        End If
    End If
    rsBusca.CloseRecordset
    Set rsBusca = Nothing
    Set RQ = Nothing
End Sub
Private Sub txtCodOpcional_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 4
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 3500
            .pCol = 2: .pAnchoCol = 500
            .pCol = 3: .pAnchoCol = 1200
            .pTitulo = "Lista de " & cmbAuxiliares.Text
            .pForm = FORM_TELEWIESE
            .pCaso = Label_Descrip_Auxil
            .Show
        End With
    End If
End Sub
Private Sub txtCuentaAux_GotFocus()
    mark1 txtCuentaAux
End Sub
Private Sub txtCuentaAux_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If KeyCode = 13 Then
       txtOficina.SetFocus
    End If
End Sub
Private Sub txtOficina_GotFocus()
    mark1 txtOficina
End Sub
Private Sub txtOficina_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
       cmdAnexar.SetFocus
    End If
End Sub
Private Function ValidarDatos() As Boolean
On Error GoTo NADA
    Dim I As Integer
    Dim TotalOrden As Double
    Dim SQL As String, documento As String
    Dim rsorden As MYSQL_RS
    ValidarDatos = False
    If Trim(meOrden) = "" Then
        MsgBox "Ingrese un número de orden", vbCritical + vbOKOnly, gsNomSW
        meOrden.SetFocus
        Exit Function
    Else
        If lblModo = "Nuevo Movimiento" Then
            orden = strAnoSistema + strMesSistema + Right("0000" & Trim(meOrden), 4)
            SQL = "Select identificador from mov_telewiese where" & _
                  " identificador='" & orden & "' and (estado='" & EMITIDO & "' or estado='" & TRANSFERIDO & "')"
            Set rsorden = oConexion.EjecutaSelectRS(SQL)
            If rsorden.RecordCount > 0 Then
                MsgBox "La orden número:" & Trim(meOrden) & " en :" & NombreMes(strMesSistema, False) & " del " & strAnoSistema, vbOKOnly + vbExclamation, gsNomSW
                meOrden = Empty
                meOrden.SetFocus
                Exit Function
            End If
        End If
    End If
    If CDbl(meImporte) = 0 Then
        If MsgBox("Desea guardar la orden sin documentos", vbYesNo + vbInformation, gsNomSW) = vbNo Then
            For I = 1 To flexDocumentos.Rows - 1
                TotalOrden = TotalOrden + CDbl(Trim(flexDocumentos.TextMatrix(I, 7)))
            Next
            If Trim(TotalOrden) <> CDbl(Trim(meImporte)) Then
                MsgBox "El importe es diferente al Total de los documentos anexados ", vbCritical + vbOKOnly, gsNomSW
                meImporte = FormatNumber(TotalOrden, 2)
                meImporte.SetFocus
                Exit Function
            End If
            cmdAnexar.SetFocus
            Exit Function
        Else
            flexDocumentos.Rows = flexDocumentos.Rows + 1
            flexDocumentos.TextMatrix(1, 0) = "1"
        End If
    End If
    If Trim(mskFecha) = "" Then
        MsgBox "Ingrese fecha de emisión", vbCritical + vbOKOnly, gsNomSW
        mskFecha.SetFocus
        Exit Function
    End If
    If Trim(txtCodigo) = "" Then
        MsgBox "Ingrese código de " & cmbAuxiliares.Text, vbCritical + vbOKOnly, gsNomSW
        txtCodigo.SetFocus
        Exit Function
    End If
    If Trim(txtBeneficiario) = "" Then
        MsgBox "Ingrese descripción de " & cmbAuxiliares.Text, vbCritical + vbOKOnly, gsNomSW
        txtBeneficiario.SetFocus
        Exit Function
    End If
    If cmbMoneda.ListIndex = 0 Then
        MsgBox "Selecione la moneda del movimiento", vbCritical + vbOKOnly, gsNomSW
        cmbMoneda.SetFocus
        Exit Function
    End If
    If cmbTipoPago.ListIndex = 0 Then
        MsgBox "Selecione el tipo de pago del movimiento", vbCritical + vbOKOnly, gsNomSW
        cmbMoneda.SetFocus
        Exit Function
    End If
    If cmbCtaCte.ListIndex = 0 Then
        MsgBox "Selecione la cuenta corriente", vbCritical + vbOKOnly, gsNomSW
        cmbCtaCte.SetFocus
        Exit Function
    End If
    ValidarDatos = True
Exit Function
NADA:
    Exit Function
End Function
Public Function ActualizaOrden() As Boolean
    Dim I As Integer
    Dim sqlorden As String
    Dim orden As String
    ActualizaOrden = False
    orden = Right("00000000000000" + Trim(meOrden), 14)
    sqlorden = "Call Delete_MovTw ( '" & strIdentificador & "')"
    oConexion.EjecutaInsertUpdateDelete sqlorden, TIPO_QUERY.Eliminar, False
    If GrabaOrden Then
        MsgBox "Se actualizó la orden número:" & meOrden, vbOKOnly + vbInformation, gsNomSW
    End If
End Function


Public Sub GeneraAsisentoCancelacion(orden As String)
On Error GoTo FallaAsiento
    Dim I As Integer
    Dim k As Integer
    Set Rs = New MYSQL_RS
    Dim SerDocu As String, NumDocu As String, v As String, AnoMes As String
    Dim glo As String, fec As String, SQL As String, vc As String, cor As String
    Dim mon As String, td As String, Div As String, Cta As String, dh As String
    Dim aux As String, caux As String, cto As String, Col As String
    Dim sol As Double, dol As Double, imp10s As Double, imp10d As Double
    Dim auxpivot As String
    Dim AuxSeg As String
    imp10s = 0
    imp10d = 0
    AuxSeg = ""
    fec = Format(CStr(Date), "dd/mm/yyyy")
    TipoCambio fec
    tc = dblTipoCmbV
    AnoMes = strAnoSistema & strMesSistema
    mon = cmbMoneda.List(cmbMoneda.ListIndex, 1)
    
    auxpivot = Trim(flexDocumentos.TextMatrix(1, 2))
    With flexDocumentos
        For k = 1 To .Rows - 1
          'Valida que no se genere asiento de cancelación para el caso de Seguros
           If ValidaSiesSeguro_NoAsientoCancel(Trim(.TextMatrix(k, 1))) = "S" Then
             AuxSeg = "1"
             Exit For
           End If
                    
           If auxpivot <> Trim(.TextMatrix(k, 2)) Then
             auxpivot = "varios"
             Exit For
           End If
        Next
    End With
    
    'Si el documento es un seguro el que se esta cancelando, no genere Asiento
    If AuxSeg = "1" Then
     Exit Sub
    End If
    
    If (flexDocumentos.Rows > 2) And (auxpivot = "varios") Then
     glo = "PAGO PROVEEDORES VARIOS"
    Else
     glo = "PAGO PROVEEDOR " & DescripcionesdeCodigos("AUXILIARES", Trim(flexDocumentos.TextMatrix(1, 2)), Trim(flexDocumentos.TextMatrix(1, 10)), "Descrip")
    End If
    
    v = MaxVoucher(AnoMes, "01")
    SQL = "Call cn_Insert_Voucher('" & Left(Trim(v), 2) & "','" & v & "','" & glo & "','" & Trim(fec) & _
            "','" & Trim(fec) & "','V'," & tc & ",'" & Trim(mon) & "','" & Trim(AnoMes) & "','" & strUsuarioId & _
             " ','CUADRADO','','','','','N','','')"
            oConexionMYSQL.Execute (SQL)
    lib = "01"
    cencos = "0000"
    cenco = "00000000000"
    gen = "N"
    With flexDocumentos
        For I = 1 To .Rows - 1
'            ' CUENTA 42
            aux = Trim(.TextMatrix(I, 10))
            caux = Trim(.TextMatrix(I, 2))
            tdoc = .TextMatrix(I, 3) '.TextMatrix(i, 15)
            Divi = IIf(Trim(.TextMatrix(I, 14)) = "", "000000000000", Trim(.TextMatrix(I, 14)))
            Divi = ValidaDameCCHFMTubulares(Divi)
                    
            cor = MaxCorrela(AnoMes, v)
            tdoc = Trim(.TextMatrix(I, 3)) 'Trim(rs.Fields("CODDOC"))   '.TextMatrix(i, 15)
            SerDocu = Trim(.TextMatrix(I, 6))
            NumDocu = Trim(.TextMatrix(I, 7))
            
            Cta = IIf(aux = "6", IIf(mon = "E", "469912", "469911"), IIf(aux = "3", IIf(mon = "E", "424002", "424001"), IIf(mon = "E", "421202", "421201")))
            
            If tdoc = "02" Then  'Caso Recibo x Honorarios
             Cta = IIf(mon = "E", "424002", "424001")
            End If
            
            cto = DescripcionesdeCodigos("AUXILIARES", caux, aux, "Descrip")
            dh = "D"
            colcv = "00"
            
            If mon = "N" Then
                sol = Abs(Round(CDbl(.TextMatrix(I, 8)), 2)) '* 0.06
                dol = Abs(Round(CDbl(.TextMatrix(I, 8)) / tc, 2)) '* 0.06
            Else
                sol = Abs(Round(CDbl(.TextMatrix(I, 8)) * tc, 2)) '* 0.06
                dol = Abs(Round(CDbl(.TextMatrix(I, 8)), 2)) '* 0.06
            End If
            
            imp10s = imp10s + sol
            imp10d = imp10d + dol
            
            SQL = "call cn_Insert_Movi ('" & lib & "','" & tdoc & "','" & Divi & "','0000000000','" & _
                v & "','" & SerDocu & "','" & NumDocu & "','" & cor & "','" & mon & "','" & Trim(Cta) & "','" & _
                aux & "','" & caux & "','" & cencos & "','" & cenco & "','" & gen & "','" & _
                cto & "'," & _
                IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                fec & "','" & strAnoSistema & strMesSistema & "','" & strUsuarioId & "','" & dh & "','" & _
                colcv & "','000','')"
            oConexionMYSQL.Execute (SQL)
             
            'Si es pago Parcial, Actualizamos Status
             SQL = "update cnmovi_paypart  set estado='P' where serdoc='" & SerDocu & "'  and numdoc='" & NumDocu & "' and codaux='" & caux & "' and estado is null"
             ADOConexion.Execute (SQL)
             
            
        Next
    End With
    aux = "1"
    caux = IIf(mon = "E", "00000000002", "00000000001")
    Cta = IIf(mon = "E", "104102", "104101")
    sol = imp10s
    dol = imp10d
    dh = "H"
    Divi = "013100003836" '0001 IIf(Trim(DIVI) = "0003", "0003", "0001")
    cto = "PAGO PROVEEDORES"
    tdoc = "9"
    SerDocu = "TW"
    NumDocu = orden
    cor = MaxCorrela(AnoMes, v)
    SQL = "call cn_Insert_Movi ('" & lib & "','" & tdoc & "','" & Divi & "','0000000000','" & _
          v & "','" & SerDocu & "','" & NumDocu & "','" & cor & "','" & mon & "','" & Trim(Cta) & "','" & _
          "1" & "','" & caux & "','" & cencos & "','" & cenco & "','" & gen & "','" & _
          cto & "'," & _
          IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
          IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
          fec & "','" & strAnoSistema & strMesSistema & "','" & strUsuarioId & "','" & dh & "','" & _
          colcv & "','000','')"
    oConexionMYSQL.Execute SQL
    oConexionMYSQL.Execute "Update mov_telewiese set vou='" & v & "' where identificador='" & orden & "'"
    Set Rs = Nothing
    MsgBox "El asiento fue generado en el voucher N° " & v & " del mes " & AnoMes, vbInformation, "NOVPeru"
    lblV.Visible = True
    lblvoucher.Visible = True
    lblvoucher.Caption = v
Exit Sub
FallaAsiento:
    SQL = "Delete from cnvouc where anomes='" & AnoMes & "' and voucher='" & v & "'"
    oConexionMYSQL.Execute SQL
    lblV.Visible = False
    
    MsgBox "Hubo un error al generar el voucher N° " & v & vbNewLine & "Consulte con su administrador del sistema ", vbOKOnly, "NOVPeru"
    Resume
    
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flexDocumentos
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
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
        .TopRow = NewValue
    End With
End Sub


Public Sub LlenaLiqPagos(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I As Integer
    SQL = "select distinct(idliq) as idliquidacion from liquidpagos where procesado=0 order by idliq desc"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "0"
    I = 1
    Do While Not Rs.EOF
        cbo.AddItem CE(Rs.Fields("idliquidacion"))
        cbo.List(I, 1) = CE(Rs.Fields("idliquidacion"))
        I = I + 1
        Rs.MoveNext
    Loop
    cbo.ListIndex = 0
    Set Rs = Nothing
End Sub

Private Sub cmbLiqPagos_Change()
'Cambia
End Sub

Private Sub cmbLiqPagos_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'Key Down
End Sub


Function ValidaDameCCHFMTubulares(ByVal vdivi As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    ValidaDameCCHFMTubulares = "013100003836"
    SQL = " Select atipo from cnmdepar where coddep= '" & vdivi & "' and atipo='TUBULAR SERVICES' "
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        ValidaDameCCHFMTubulares = "013100003841"
    Else
        ValidaDameCCHFMTubulares = vdivi
    End If
    Set RQ = Nothing
End Function


Private Sub TxtLiq_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
  If BuscarLiquidacion Then
     MsgBox "Liquidación visualizada", vbInformation, "NOVPeru"
  End If
End Sub

Public Function BuscarLiquidacion() As Boolean
    Dim sqlorden As String
    Dim rsorden As New ADODB.Recordset
    BuscarLiquidacion = False
    If TxtLiq <> "" Then
        sqlorden = "Call cn_buscar_liquidacion('" & CInt(TxtLiq.Text) & "','" & strAnoSistema & strMesSistema & Right("0000" & Trim(meOrden), 4) & "');"
        Set rsorden = ADO_LlenaRs(sqlorden)
    
    If Not rsorden Is Nothing Then
            CargarDatosLiq rsorden
            BuscarLiquidacion = True
            Exit Function
        End If
    End If
End Function


Public Sub CargarDatosLiq(rsorden As ADODB.Recordset)
    Dim I As Integer, num As Integer
    Dim AcumImporte As Double
    Dim AcumImporteEQ As Double
    
    lblimporte = 0
    lblImpEqu = 0

    With rsorden
        .MoveFirst
        Do While Not .EOF
            If val(.Fields("item")) = 1 Then
                meOrden = Right(CE(.Fields("identificador")), 4)
                strIdentificador = CE(.Fields("identificador"))
                lblEstado = DescripcionesdeCodigos("DOC_ESTADO", CE(.Fields("estado")))
                lblEstado.tag = CE(.Fields("estado"))
                If CE(.Fields("voucher")) <> "" Then
                    lblV.Visible = True
                    lblvoucher.Visible = True
                    lblvoucher.Caption = CE(.Fields("voucher"))
                Else
                    lblV.Visible = False
                    lblvoucher.Visible = False
                End If
                If .RecordCount = 1 Then
                    lblFolio = CE(.Fields("folioref"))
                End If
                If CE(.Fields("fecha") <> "") Then mskFecha = Format(CE(.Fields("fecha")), "dd/mm/yyyy")
                Auxiliares cmbAuxiliares
                For I = 0 To cmbAuxiliares.ListCount - 1
                    If cmbAuxiliares.List(I, 1) = CE(.Fields("Auxiliar")) Then
                        cmbAuxiliares.ListIndex = I
                    End If
                Next
                txtCodigo = CE(.Fields("Codigo"))
                txtBeneficiario = DescripcionesdeCodigos("AUXILIARES", CE(.Fields("Codigo")), CE(.Fields("Auxiliar")), "Descrip")
                txtOficina = CE(.Fields("Oficina"))
                txtCuentaAux = CE(.Fields("ctaaux"))
                moneda cmbMoneda
                For I = 0 To cmbMoneda.ListCount - 1
                    If cmbMoneda.List(I, 1) = CE(.Fields("Moneda")) Then
                        cmbMoneda.ListIndex = I
                    End If
                Next
                For I = 0 To cmbCtaCte.ListCount - 1
                    If cmbCtaCte.List(I, 1) = CE(.Fields("ctacte")) Then
                        cmbCtaCte.ListIndex = I
                    End If
                Next
                TipPago cmbTipoPago
                For I = 0 To cmbTipoPago.ListCount - 1
                    If cmbTipoPago.List(I, 1) = CE(.Fields("tipopago")) Then
                        cmbTipoPago.ListIndex = I
                    End If
                Next
                meImporte = str(CDbl(meImporte) + CDbl(FormatNumber(CEN(.Fields("importe")), 2)))
                txtConcepto = CE(.Fields("OBS"))
            End If
            flexDocumentos.Rows = flexDocumentos.Rows + 1
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 0) = Left(CE(rsorden.Fields("item")), 6)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 1) = CE(.Fields("folioref"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 2) = CE(.Fields("codigo"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 3) = CE(.Fields("tipdoc"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 4) = CE(.Fields("moneda"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 5) = CE(.Fields("fec_emi"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 6) = CE(.Fields("serie"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 7) = CE(.Fields("documento"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 8) = FormatNumber(CEN(.Fields("importe")), 2)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 9) = FormatNumber(CEN(.Fields("importeequ")), 2)
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 10) = CE(.Fields("auxiliar"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 11) = CE(.Fields("tipopago"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 12) = CE(.Fields("oficina"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 13) = CE(.Fields("ctaaux"))
            flexDocumentos.TextMatrix(flexDocumentos.Rows - 1, 14) = CE(.Fields("DIVI"))
            AcumImporte = CDbl(AcumImporte) + CEN(CDbl(.Fields("importe")) * (-1))
            AcumImporteEQ = CDbl(AcumImporteEQ) + CEN(CDbl(.Fields("importeequ")) * (-1))
            num = num + 1
            rsorden.MoveNext
        Loop
    End With
    
    meImporte.Text = CStr(AcumImporte)
    lblImpEqu = CStr(AcumImporteEQ)
    
    lblTCambio = TCambio(mskFecha)
    If num = 1 Then lblDocAnexados = str(num) & " Documento Anexado": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
    If num > 1 Then lblDocAnexados = str(num) & " Documentos Anexados": lblDocAnexados.Visible = True: lblDocAnexados.tag = num
    lblEstado.Visible = True
    
End Sub


Function ValidaSiesSeguro_NoAsientoCancel(ByVal vfolio As String) As String
    Dim SQL As String
    Dim RQS As MYSQL_RS
    
    ValidaSiesSeguro = "-"
    SQL = "select d.identificador from documento_contables as d left join amarre_documento as f on d.Identificador = f.Identificador where f.Cod_Tipo_Doc='S' and d.identificador='" & vfolio & "' "
    Set RQS = oConexion.EjecutaSelectRS(SQL)
    
    If Not RQS.EOF() Then
        ValidaSiesSeguro = "S"
    End If
    
    Set RQS = Nothing
End Function




Public Sub cmdenviarNotificacionPago()
    On Error GoTo SERROR
    Dim SQL As String, ic As String, Idi As String, msj As String, Email As String, Asunto As String
    Dim RQ1 As MYSQL_RS
    Dim NomArchivo As String
    Dim FlagEnvio As String
    Dim pivotruc As String
    Dim msjRuc As String
    Dim msj1 As String
    Dim msj2 As String
    Dim orden As String
   
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "CUENTA CORRIENTE POR AUXILIAR Y DOCUMENTO """
    oReporte.Reporte = "Rep_Det_Cta_CteXdoc.rpt"

    Dim I As Integer
    Dim frmRep As frmReportPreview
    
    FlagEnvio = ""
    Asunto = "Pago de NATIONAL OILWELL VARCO PERU SRL  a  su empresa"
    msj1 = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>Para informale que el dia " & LCase(Day(CDate(mskFecha))) & " de " & LCase(NombreMes(Month(CDate(mskFecha)), False)) & _
           " La compañia National Oilwell Varco Peru SRL realizó una transferencia a su compañia </font>"
    msj2 = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'></br> Por favor comunicar una vez recibido los fondos.</br></br>Muchas Gracias." & _
           "</br></br>Saludos Cordiales,</br></br>Rosa Gallegos y Julia Ignacio </br> Dpto de Contabilidad</font>"
    
    pivotruc = flexDocumentos.TextMatrix(1, 2)
    
    msj = "<table border='1' cellpadding='0' cellspacing='0' style='color:Black; font-weight:bold; width:800px; text-align:center;font-family:Verdana; font-size:small;'>"
    msj = msj & "<tr style='background-color:#FFF6FA;><td width='100px'>Monto</td><td width='200px'>Nro Documento</td><td width='300px'>Pago Realizado</td><td width='50px'>Oficina</td><td width='150px'>Nro Cuenta</td></tr>"
                 
    With flexDocumentos
        For I = 1 To .Rows
            If .TextMatrix(I, 0) <> "" Then
                 Email = EmailContactoAuxiliar(pivotruc)
                 msjRuc = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>con  RUC " & pivotruc & ":</br></font>"
                                    
                 If (Email <> "") And (pivotruc = .TextMatrix(I, 2)) Then
                    msj = msj & "<tr style='background-color:#E6E6FA;'><td> " & IIf(.TextMatrix(I, 4) = "N", "S/.", "US$") & " " & .TextMatrix(I, 8) & "</td><td>" & .TextMatrix(I, 7) & "</td><td>" & DameTipoPago(.TextMatrix(I, 11)) & "</td><td>" & .TextMatrix(I, 12) & "</td><td>" & .TextMatrix(I, 13) & "</td></tr>  "
                 Else
                    If Email <> "" Then
                       If EnviarEmail(Email, "", "", "", msj1 & msjRuc & msj & "</table>" & msj2, Asunto, "Rosa.GallegosGallegos@nov.com;julia.ignacio@nov.com;", "") Then
                         FlagEnvio = "Enviado"
                       End If
                    End If
                    
                    pivotruc = .TextMatrix(I, 2)
                    msj = "<table border='1' cellpadding='0' cellspacing='0' style='color:Blue; font-weight:bold; width:800px; text-align:center;font-family:Verdana; font-size:small;'>"
                    msj = msj & "<tr style='background-color:Gray;><td width='100px'>Monto</td><td width='200px'>Nro Documento</td><td width='300px'>Pago Realizado</td><td width='50px'>Oficina</td><td width='150px'>Nro Cuenta</td></tr>"
                    msj = msj & "<tr style='background-color:Lime;'><td> " & IIf(.TextMatrix(I, 4) = "N", "S/.", "US$") & " " & .TextMatrix(I, 8) & "</td><td>" & .TextMatrix(I, 7) & "</td><td>" & DameTipoPago(.TextMatrix(I, 11)) & "</td><td>" & .TextMatrix(I, 12) & "</td><td>" & .TextMatrix(I, 13) & "</td></tr>  "
                 End If
            End If
        Next
        
     If Email <> "" And msj <> "" Then
        Email = EmailContactoAuxiliar(pivotruc)
        msjRuc = "<font style='font-weight:bold;font-family:Verdana; font-size:small;'>con  RUC " & pivotruc & ":</br></font>"
        If EnviarEmail(Email, "", "", "", msj1 & msjRuc & msj & "</table>" & msj2, Asunto, "Rosa.GallegosGallegos@nov.com;julia.ignacio@nov.com;gustavo.mayaute@nov.com", "") Then
          FlagEnvio = "Enviado"
        End If
     End If
        
    End With
  
    If FlagEnvio = "Enviado" Then
       MsgBox "El Proveedor/Beneficiario ha recibido el Email de Contacto", vbInformation, "NOVPeru"
    Else
       MsgBox "El Proveedor/Beneficiario no posee ningún Email de Contacto", vbInformation, "NOVPeru"
    End If
    
    '    btnCopia_Click
    'FlgEmail = False
    Exit Sub
SERROR:
    Mensajes err.Description

End Sub


