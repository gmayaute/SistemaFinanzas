VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmLetras 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administración de Letras"
   ClientHeight    =   7305
   ClientLeft      =   1740
   ClientTop       =   4260
   ClientWidth     =   13785
   Icon            =   "frmLetras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   13785
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Txt 
      Height          =   285
      Left            =   3210
      TabIndex        =   36
      Top             =   6540
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshLetrasRen 
      Height          =   1980
      Left            =   60
      TabIndex        =   16
      Top             =   5295
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   3493
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshFacturas 
      Height          =   3375
      Left            =   8910
      TabIndex        =   15
      Top             =   1530
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshLetras 
      Height          =   2265
      Left            =   60
      TabIndex        =   7
      Top             =   1530
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   3995
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSMask.MaskEdBox meFecGiro 
      Height          =   315
      Left            =   4290
      TabIndex        =   5
      Top             =   180
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meFecvcto 
      Height          =   315
      Left            =   7455
      TabIndex        =   6
      Top             =   180
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton BtnNuevo 
      Height          =   345
      Left            =   45
      TabIndex        =   10
      ToolTipText     =   "Nuevo"
      Top             =   1140
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Nuevo"
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
      MICON           =   "frmLetras.frx":030A
      PICN            =   "frmLetras.frx":0326
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
      Left            =   2460
      TabIndex        =   11
      ToolTipText     =   "Eliminar"
      Top             =   1140
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Eliminar"
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
      MICON           =   "frmLetras.frx":0690
      PICN            =   "frmLetras.frx":06AC
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
      Left            =   1263
      TabIndex        =   12
      ToolTipText     =   "Modificar"
      Top             =   1140
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Modificar"
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
      MICON           =   "frmLetras.frx":0AEE
      PICN            =   "frmLetras.frx":0B0A
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
      Left            =   3690
      TabIndex        =   13
      ToolTipText     =   "Guardar"
      Top             =   1140
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
      MICON           =   "frmLetras.frx":0F38
      PICN            =   "frmLetras.frx":0F54
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdRenovar 
      Height          =   345
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   3810
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Renovar"
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
      MICON           =   "frmLetras.frx":1396
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdFacs 
      Height          =   345
      Left            =   9960
      TabIndex        =   9
      ToolTipText     =   "Guardar"
      Top             =   1140
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "..."
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
      MICON           =   "frmLetras.frx":13B2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox mePago 
      Height          =   315
      Left            =   10125
      TabIndex        =   20
      Top             =   180
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meInteres 
      Height          =   315
      Left            =   1770
      TabIndex        =   23
      Top             =   4620
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox meIntBco 
      Height          =   315
      Left            =   3480
      TabIndex        =   24
      Top             =   4620
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   15
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton BtnSalir 
      Height          =   345
      Left            =   8280
      TabIndex        =   28
      ToolTipText     =   "Salir"
      Top             =   1140
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
      MICON           =   "frmLetras.frx":13CE
      PICN            =   "frmLetras.frx":13EA
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
      Left            =   6780
      TabIndex        =   29
      ToolTipText     =   "Deshacer"
      Top             =   1140
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
      MICON           =   "frmLetras.frx":17B0
      PICN            =   "frmLetras.frx":17CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdAnular 
      Height          =   345
      Left            =   4815
      TabIndex        =   37
      ToolTipText     =   "Anular"
      Top             =   1140
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Anular"
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
      MICON           =   "frmLetras.frx":1D0E
      PICN            =   "frmLetras.frx":1D2A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdCancel 
      Height          =   345
      Left            =   1320
      TabIndex        =   38
      ToolTipText     =   "Ingresar Deducciones"
      Top             =   3810
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Cancelar Letras"
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
      MICON           =   "frmLetras.frx":3D2C
      PICN            =   "frmLetras.frx":3D48
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.Label lblV 
      Height          =   345
      Left            =   9480
      TabIndex        =   35
      Top             =   735
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
      Height          =   330
      Left            =   11520
      TabIndex        =   34
      Top             =   735
      Width           =   2145
      ForeColor       =   65280
      BackColor       =   10442041
      Caption         =   "010001"
      Size            =   "3784;582"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CheckBox ChkCancel 
      Height          =   285
      Left            =   10575
      TabIndex        =   33
      Top             =   1170
      Width           =   2010
      BackColor       =   10442041
      ForeColor       =   8421631
      DisplayStyle    =   4
      Size            =   "3545;503"
      Value           =   "0"
      Caption         =   "Cancelar Facturas"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label lblren 
      Height          =   165
      Left            =   13245
      TabIndex        =   32
      Top             =   705
      Visible         =   0   'False
      Width           =   375
      BackColor       =   10442041
      Size            =   "661;291"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblimporte 
      BackColor       =   &H009F5539&
      Caption         =   "Label10"
      Height          =   210
      Left            =   13185
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblModo 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   12735
      TabIndex        =   30
      Top             =   1215
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSForms.CheckBox ChkND 
      Height          =   285
      Left            =   5220
      TabIndex        =   27
      Top             =   4485
      Width           =   3405
      BackColor       =   10442041
      ForeColor       =   8421631
      DisplayStyle    =   4
      Size            =   "6006;503"
      Value           =   "0"
      Caption         =   "Generar Nota de Débito por interés"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label8 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interés Bco."
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
      Height          =   315
      Left            =   3480
      TabIndex        =   26
      Top             =   4230
      Width           =   1605
   End
   Begin VB.Label Label7 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interés"
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
      Height          =   315
      Left            =   1800
      TabIndex        =   25
      Top             =   4230
      Width           =   1545
   End
   Begin MSForms.TextBox txtCodBco 
      Height          =   345
      Left            =   90
      TabIndex        =   22
      Top             =   4590
      Width           =   1575
      VariousPropertyBits=   746604571
      MaxLength       =   10
      Size            =   "2778;609"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label6 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CobBanco"
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
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   1545
   End
   Begin VB.Label Label5 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pago:"
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
      Height          =   315
      Left            =   8955
      TabIndex        =   19
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RENOVACIONES"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURAS"
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
      Left            =   8925
      TabIndex        =   17
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Vcto.:"
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
      Height          =   345
      Left            =   5790
      TabIndex        =   4
      Top             =   165
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Giro:"
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
      Height          =   315
      Left            =   2610
      TabIndex        =   3
      Top             =   180
      Width           =   1605
   End
   Begin MSForms.ComboBox cboCliente 
      Height          =   375
      Left            =   1290
      TabIndex        =   8
      Top             =   600
      Width           =   7515
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "13256;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtNumLetra 
      Height          =   345
      Left            =   1290
      TabIndex        =   2
      Top             =   165
      Width           =   1155
      VariousPropertyBits=   746604571
      MaxLength       =   6
      Size            =   "2037;609"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      BackColor       =   &H009F5539&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   30
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Letra:"
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
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SerieND As String
Dim GridSel As Integer
Dim HabModRen As Boolean

Public Sub Clientes()
    Dim SQL As String
    Dim I As Integer
    Dim rsClientes As MYSQL_RS
    SQL = " Select distinct dc.Auxiliar,dc.codigo,aux.descrip " & _
         " from ((amarre_documento as a LEFT Join documento_contables as dc  on (a.Identificador=dc.Identificador))" & _
         " LEFT join movi_documento as mov  on (dc.Identificador=mov.Identificador))LEFT join cnauxil as aux on (dc.codigo=aux.codigo )" & _
         " where  a.Flag='0' " & _
         " and aux.auxiliar='2' and (mov.Cod_estado='IM' or  mov.Cod_estado='EM') AND mov.Cod_estado<>'AN' and mov.Cod_estado <>'EL' AND" & _
         " dc.total<>0 and (a.cod_tipo_doc='01' or a.cod_tipo_doc='03' or a.cod_tipo_doc='P')  ORDER by aux.descrip"
    Set rsClientes = oConexion.EjecutaSelectRS(SQL)
    I = 0
    cboCliente.Clear
    Do While Not rsClientes.EOF
        cboCliente.AddItem CE(rsClientes.Fields("Descrip"))
        cboCliente.List(I, 1) = CE(rsClientes.Fields("Codigo"))
        I = I + 1
        rsClientes.MoveNext
    Loop
    cboCliente.ListIndex = -1
    Set rsClientes = Nothing
End Sub

Public Sub CargaLetras(Cod As String)
    Dim SQL As String
    Dim rsletras As MYSQL_RS
    Dim RQ As MYSQL_RS
    With mshLetras
        .Clear
        .Rows = 1
        .Cols = 12
        .ForeColorFixed = &H404000
        
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 900
        .TextMatrix(0, 1) = "Letra"
        
        .ColWidth(2) = 1100
        .TextMatrix(0, 2) = "Fec. Giro"
            
        .ColWidth(3) = 1100
        .TextMatrix(0, 3) = "Fec. Vcto."
        
        .ColWidth(4) = 1300
        .TextMatrix(0, 4) = "Importe"
    
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = "Abono"
        
        .ColWidth(6) = 1300
        .TextMatrix(0, 6) = Space(8) + "Saldo"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "cli"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "codbanco"
        
        .ColWidth(9) = 1400
        .TextMatrix(0, 9) = "NDébito"
        
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = "voucher"
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = "estado"
        SQL = " Select l.* from letra l where codaux='" & Cod & "' and tipo = 'P' and codestado <> 'EL' order by numero desc"
        Set rsletras = oConexion.EjecutaSelectRS(SQL)
        Do While Not rsletras.EOF
            .Rows = .Rows + 1
            If .Rows = 2 Then
                .FixedRows = 1
            End If
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = rsletras.Fields("numero")
            .TextMatrix(.Rows - 1, 2) = rsletras.Fields("fecgiro")
            .TextMatrix(.Rows - 1, 3) = rsletras.Fields("fecvcto")
            .TextMatrix(.Rows - 1, 4) = FormatNumber(rsletras.Fields("importe"), 2)
            .TextMatrix(.Rows - 1, 5) = FormatNumber(rsletras.Fields("abono"), 2)
            .TextMatrix(.Rows - 1, 6) = FormatNumber(rsletras.Fields("importe") - rsletras.Fields("abono"), 2)
            .TextMatrix(.Rows - 1, 7) = rsletras.Fields("codaux")
            .TextMatrix(.Rows - 1, 8) = rsletras.Fields("codbco")
            .TextMatrix(.Rows - 1, 9) = rsletras.Fields("ndebito")
            .TextMatrix(.Rows - 1, 10) = rsletras.Fields("voucher")
            .TextMatrix(.Rows - 1, 11) = rsletras.Fields("codestado")
            If Trim(rsletras.Fields("codestado")) = "AN" Then
                Dim I As Integer
                For I = 1 To .Cols - 1
                    .Col = I
                    .row = .Rows - 1
                    .CellBackColor = vbGreen
                Next
            End If
            If rsletras.Fields("codestado") = "CA" Then
                For I = 1 To .Cols - 1
                    .Col = I
                    .row = .Rows - 1
                    .CellBackColor = vbRed
                Next
            Else
                SQL = "select * from letra where ref = '" & rsletras.Fields("numero") & "' and codestado = 'CA'"
                Set RQ = oConexion.EjecutaSelectRS(SQL)
                If Not RQ.EOF() Then
                    .TextMatrix(.Rows - 1, 11) = "CA"
                    For I = 1 To .Cols - 1
                        .Col = I
                        .row = .Rows - 1
                        .CellBackColor = vbRed
                    Next
                End If
                Set RQ = Nothing
            End If
            rsletras.MoveNext
        Loop
        Set rsletras = Nothing
    End With
End Sub

Sub CargarFacturasPorLetra(num As String)
    Dim RQ As MYSQL_RS, I As Integer
    Dim SQL As String
    FormatoFacturas
    I = 1
    SQL = "select f.identificador,CONCAT(d.Serie,'-',d.Correl) as Documento,d.fec_emision,f.monto " & _
          "from factura_letra f inner join documento_contables d on (f.identificador=d.identificador) " & _
          "where letra = '" & num & "' order by documento"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    With MshFacturas
        Do While Not RQ.EOF()
            .TextMatrix(I, 0) = Trim(RQ.Fields("identificador"))
            .TextMatrix(I, 1) = Trim(RQ.Fields("documento"))
            .TextMatrix(I, 2) = Format(RQ.Fields("fec_emision"), "dd/mm/yyyy")
            .TextMatrix(I, 3) = FormatNumber(RQ.Fields("monto"), 2)
            I = I + 1
            .Rows = .Rows + 1
            RQ.MoveNext
        Loop
        If Trim(MshFacturas.TextMatrix(I, 0)) = "" And MshFacturas.Rows > 2 Then
            MshFacturas.Rows = MshFacturas.Rows - 1
        End If
    End With
    Set RQ = Nothing
End Sub

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    ModoFormulario modConsulta
    lblModo = ""
    txtNumLetra = ""
    txtNumLetra.Locked = False
    txtNumLetra.BackColor = ColorHabilitado
    cmdRenovar.Enabled = False
    lblvoucher.Visible = False
    lblV.Visible = False
End Sub

Private Sub btnEliminar_Click()
Dim SQL As String, FlgRen As Boolean
    If MsgBox("¿Seguro que desea eliminar la Letra " & mshLetras.TextMatrix(mshLetras.row, 1) & "?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        FlgRen = VerificaRenovaciones(mshLetras.TextMatrix(mshLetras.row, 1))
        If FlgRen = True Then
            If MsgBox("Esta Letra tiene renovaciones. Si la eliimna, se eliminarán sus renovaciones. ¿Seguro que desea continuar?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                SQL = "delete from letra where ref = '" & mshLetras.TextMatrix(mshLetras.row, 1) & "'"
                oConexionMYSQL.Execute SQL
                GoTo AQUI
            Else
                Exit Sub
            End If
        Else
            GoTo AQUI
        End If
AQUI:
        SQL = "delete from factura_letra where letra = '" & mshLetras.TextMatrix(mshLetras.row, 1) & "'"
        oConexionMYSQL.Execute SQL
        SQL = "delete from letra where numero = '" & mshLetras.TextMatrix(mshLetras.row, 1) & "'"
        oConexionMYSQL.Execute SQL
        If cboCliente.ListIndex > -1 Then
            CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
        End If
        lblvoucher.Visible = False
        lblV.Visible = False
    End If
End Sub

Function VerificaRenovaciones(NumLetra As String) As Boolean
    Dim SQL As String, RQ As MYSQL_RS
    VerificaRenovaciones = False
    SQL = "select * from letra where ref = '" & NumLetra & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        VerificaRenovaciones = True
    End If
    Set RQ = Nothing
End Function

Private Sub btnGrabar_Click()
     If lblModo = "nuevo" Then
        If GrabarDatosDoc Then
           ModoFormulario modConsulta
        End If
    End If
    If lblModo = "modificar" Then
        If ActualizaData Then
           ModoFormulario modConsulta
        End If
    End If
End Sub

Function ActualizaData() As Boolean
On Error GoTo errorgrabar
    Dim SQL As String, I As Integer
    If ValidarData Then
        SQL = "delete from factura_letra where letra = '" & Trim(txtNumLetra) & "'"
        oConexionMYSQL.Execute SQL
        With MshFacturas
            For I = 1 To .Rows - 1
                SQL = "insert into factura_letra(letra,identificador,monto) " & _
                      "values ('" & Trim(txtNumLetra.Text) & "','" & Trim(.TextMatrix(I, 0)) & "', " & _
                      "" & CDbl(.TextMatrix(I, 3)) & ")"
                oConexionMYSQL.Execute SQL
            Next
        End With
        SQL = "update letra set importe=" & CDbl(lblimporte) & ", fecgiro= '" & meFecGiro & "',fecvcto='" & meFecvcto & "',abono=" & mePago.Text & ", " & _
              "codbco='" & txtCodBco & "',codestado = 'MO' where numero = '" & Trim(txtNumLetra) & "'"
        oConexionMYSQL.Execute SQL
        ActualizaData = True
        CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
    End If
Exit Function
errorgrabar:
    ActualizaData = False
   MsgBox err.Description, vbOKOnly, "ERROR"
   MsgBox "Error al momento de grabar, revise los datos y vuelva a intentarlo", vbOKOnly + vbExclamation, "NOVBRANDT"
   Resume Next
End Function

Function GrabarDatosDoc() As Boolean
On Error GoTo errorgrabar
    Dim SQL As String, I As Integer
    If ValidarData Then
        SQL = "Insert into letra (numero,fecgiro,fecvcto,importe,abono,codaux,codbco,tipo,Interes," & _
              " InteresBco,codestado,ref,usuario) values ('" & Right("000000" & Trim(txtNumLetra), 6) & "','" & meFecGiro & "','" & _
              meFecvcto & "'," & CDbl(lblimporte.Caption) & ",0,'" & cboCliente.List(cboCliente.ListIndex, 1) & "','" & txtCodBco & "','P'," & _
              "0,0,'EM','','" & strUsuarioId & "')"
        oConexionMYSQL.Execute SQL
        SQL = "delete from factura_letra where letra = '" & Trim(txtNumLetra.Text) & "'"
        oConexionMYSQL.Execute SQL
        With MshFacturas
            For I = 1 To .Rows - 1
                SQL = "insert into factura_letra(letra,identificador,monto) " & _
                      "values ('" & Trim(txtNumLetra.Text) & "','" & Trim(.TextMatrix(I, 0)) & "', " & _
                      "" & CDbl(.TextMatrix(I, 3)) & ")"
                oConexionMYSQL.Execute SQL
            Next
        End With
        If ChkCancel.Value = True Then
            CancelarFacturas meFecGiro
            GenerarAsientoCancelacion cboCliente.List(cboCliente.ListIndex, 1), Trim(txtNumLetra.Text)
        End If
        CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
        GrabarDatosDoc = True
    End If
Exit Function
errorgrabar:
    GrabarDatosDoc = False
   MsgBox err.Description, vbOKOnly, "ERROR"
   MsgBox "Error al momento de grabar, revise los datos y vuelva a intentarlo", vbOKOnly + vbExclamation, "NOVBRANDT"
End Function

Sub CancelarFacturas(fecha As String)
Dim SQL As String, I As Integer
Dim RQ As MYSQL_RS
    With MshFacturas
        For I = 1 To .Rows - 1
            SQL = "SELECT (total-total_ref- " & rptIgv("Fec_emision", "1") & " * ifnull((Select sum(valor) from factura_proforma where " & _
                  "prof=CONCAT(serie,'-',correl)),0)) as saldo from documento_contables where identificador = '" & .TextMatrix(I, 0) & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF Then
                If CDbl(RQ.Fields("saldo")) - CDbl(.TextMatrix(I, 3)) = 0 Then
                    SQL = "Update movi_documento set Cod_Estado='" & CANCELADO & "' where Identificador='" & .TextMatrix(I, 0) & "'"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                    SQL = "Update documento_contables set Cancelado=" & CDbl(.TextMatrix(I, 3)) & ",fec_pago='" & Format(fecha, "yyyy/mm/dd") & "' where Identificador='" & .TextMatrix(I, 0) & "'"
                    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                    .TextMatrix(I, 4) = "x"
                End If
            End If
        Next
        Set RQ = Nothing
    End With
End Sub

Sub GenerarAsientoCancelacion(CodCli As String, NumLetra As String)
On Error GoTo CtrlError
    Dim AnoMes As String, lib As String
    Dim v As String, Cta As String, mon As String, cencos As String, tipovoc As String
    Dim Numdoc As String, Div As String, cenco As String, correl As String
    Dim glo As String, aux As String, caux As String, prof As String, td As String
    Dim tc As Double, SQL As String, RQ As MYSQL_RS, cto As String, dh As String
    Dim fec As String, colv As String, I As Integer, sol As Double, dol As Double
    lib = "05"
    AnoMes = strAnoSistema & strMesSistema
    v = MaxVoucher(AnoMes, lib)
    glo = "CANJE DE LETRAS " & DescripcionesdeCodigos("AUXILIARES", Trim(CodCli), 2, "Descrip")
    fec = Date
    TipoCambio (fec)
    tc = dblTipoCmbV
    mon = "E"
    colv = ""
    SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & fec & _
          " ','" & fec & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
          " ','CUADRADO','','','','','N','','')"
    oConexionMYSQL.Execute (SQL)
    With MshFacturas
        For I = 1 To .Rows - 1
            td = "01"
            Div = Trim(.TextMatrix(I, 5))
            Numdoc = Trim(.TextMatrix(I, 1))
            correl = Right("0000" & Trim(CStr(I)), 4)
            mon = Trim(.TextMatrix(I, 6))
            Cta = IIf(mon = "N", "121302", "121301")
            aux = 2
            caux = CodCli
            cto = DescripcionesdeCodigos("AUXILIARES", Trim(CodCli), 2, "Descrip")
            sol = Round(CDbl(.TextMatrix(I, 3)) * tc, 2)
            dol = Round(CDbl(.TextMatrix(I, 3)), 2)
            dh = "H"
            colv = ""
            SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','','" & _
                  v & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                  aux & "','" & caux & "','0000','0000000000','N','" & cto & "'," & _
                  IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                  IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                  fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                  colv & "','000','')"
            oConexionMYSQL.Execute (SQL)
        Next
        SQL = "Select distinct dc.division as divi,SUM(monto) as mto from factura_letra f left join " & _
              "documento_contables dc on (f.identificador=dc.identificador) where letra = '" & NumLetra & "' group by dc.division"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        Do While Not RQ.EOF()
            td = "L"
            Div = Trim(RQ.Fields("divi"))
            Numdoc = NumLetra
            correl = Right("0000" & Trim(CStr(I)), 4)
            mon = "E"
            Cta = "12301"
            aux = "2"
            caux = CodCli
            cto = "LETRAS N° " & NumLetra
            dh = "D"
            sol = Round(CDbl(RQ.Fields("mto")) * tc, 2)
            dol = Round(CDbl(RQ.Fields("mto")), 2)
            colv = ""
            SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','','" & _
                  v & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                  aux & "','" & caux & "','0000','0000000000','N','" & cto & "'," & _
                  IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                  IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                  fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                  colv & "','000','')"
            oConexionMYSQL.Execute (SQL)
            I = I + 1
            RQ.MoveNext
        Loop
        lblvoucher.Visible = True
        lblV.Visible = True
        lblvoucher.Caption = AnoMes & "-" & v
        SQL = "update letra set voucher = '" & AnoMes & v & "' where NUMERO = '" & NumLetra & "'"
        oConexionMYSQL.Execute (SQL)
    End With
Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error Generando Asientos"
End Sub

Private Function ValidarData() As Boolean
    Dim I As Integer
    ValidarData = True
    If txtNumLetra = "" Then
        MsgBox "Debe ingresar el Número de Letra", vbOKOnly + vbCritical, "NOVPeru"
        txtNumLetra.SetFocus
        ValidarData = False
        Exit Function
    End If
    If meFecGiro = "__/__/____" Or Len(meFecGiro) > 10 Or Len(meFecGiro) < 10 Then
        MsgBox "Ingrese una Fecha de Giro Válida", vbOKOnly + vbCritical, "NOVPeru"
        meFecGiro.SetFocus
        ValidarData = False
        Exit Function
    End If
    If meFecvcto = "__/__/____" Or Len(meFecvcto) > 10 Or Len(meFecvcto) < 10 Then
        MsgBox "Ingrese una Fecha de Vencimiento Válida", vbOKOnly + vbCritical, "NOVPeru"
        meFecvcto.SetFocus
        ValidarData = False
        Exit Function
    End If
    If cboCliente.ListIndex = -1 Then
        MsgBox "Debe Seleccionar un Cliente", vbOKOnly + vbCritical, "NOVPeru"
        cboCliente.SetFocus
        ValidarData = False
        Exit Function
    End If
    If MshFacturas.TextMatrix(1, 0) = "" Then
        MsgBox "No se puede registrar una Letra" & vbNewLine & "sin relacionarlo con alfuna factura", vbOKOnly + vbCritical, "NOVPeru"
        MshFacturas.SetFocus
        ValidarData = False
        Exit Function
    End If
End Function

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    txtNumLetra.Locked = False
    meFecGiro.Enabled = True
    meFecvcto.Enabled = True
    cboCliente.Enabled = True
    btnGrabar.Enabled = True
    cmdFacs.Enabled = True
    ModoFormulario modNuevo
    lblvoucher.Visible = False
    lblV.Visible = False
    meFecGiro.SetFocus
End Sub

Private Sub LimpiarDatos()
    txtNumLetra.Text = ""
    meFecGiro = "__/__/____"
    meFecvcto = "__/__/____"
    mePago = "0.00"
    txtCodBco.Text = ""
    meInteres = "0.00"
    meIntBco = "0.00"
    cboCliente.ListIndex = -1
    ChkND.Value = False
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
            LimpiarDatos
            BloqueoControles True
            lblModo = "Acción"
            cboCliente.BackColor = ColorHabilitado
            cboCliente.Locked = False
            mshLetras.BackColor = ColorHabilitado
            mshLetras.Enabled = True
        Case ModoForm.modNuevo
            LimpiarDatos
            BloqueoControles False
            ConfigurarBotones cfgNuevo
            CargaLetras "0"
            txtNumLetra = MaxCorrelativo
            lblModo = "nuevo"
            txtNumLetra.Locked = True
            txtNumLetra.BackColor = ColorDeshabilitado
            mshLetras.BackColor = ColorDeshabilitado
            mshLetras.Enabled = False
            CargarLetrasRenovadas "0"
        Case ModoForm.modConsulta
            BloqueoControles True
            ConfigurarBotones cfgGrabar
            cmdFacs.Enabled = False
            btnCancelar.Enabled = False
            cboCliente.BackColor = ColorHabilitado
            cboCliente.Locked = False
            mshLetras.BackColor = ColorHabilitado
            mshLetras.Enabled = True
            lblModo = ""
        Case ModoForm.modEditar
            BloqueoControles False
            ConfigurarBotones cfgModificar
            lblModo = "modificar"
            txtNumLetra.Locked = True
            txtNumLetra.BackColor = ColorDeshabilitado
            mshLetras.BackColor = ColorDeshabilitado
            mshLetras.Enabled = False
            cmdFacs.Enabled = True
            meFecGiro.SetFocus
    End Select
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtNumLetra.Locked = valor
    txtCodBco.Locked = valor
    meFecGiro.Enabled = Not (valor)
    meFecvcto.Enabled = Not (valor)
    mePago.Enabled = Not (valor)
    meInteres.Enabled = Not (valor)
    meIntBco.Enabled = Not (valor)
    cboCliente.Locked = valor
    ChkND.Locked = valor
    If valor = True Then
        txtNumLetra.BackColor = ColorDeshabilitado
        txtCodBco.BackColor = ColorDeshabilitado
        meFecGiro.BackColor = ColorDeshabilitado
        meFecvcto.BackColor = ColorDeshabilitado
        mePago.BackColor = ColorDeshabilitado
        meInteres.BackColor = ColorDeshabilitado
        meIntBco.BackColor = ColorDeshabilitado
        cboCliente.BackColor = ColorDeshabilitado
        MshFacturas.BackColor = ColorDeshabilitado
    Else
        txtNumLetra.BackColor = ColorHabilitado
        txtCodBco.BackColor = ColorHabilitado
        meFecGiro.BackColor = ColorHabilitado
        meFecvcto.BackColor = ColorHabilitado
        mePago.BackColor = ColorHabilitado
        meInteres.BackColor = ColorHabilitado
        meIntBco.BackColor = ColorHabilitado
        cboCliente.BackColor = ColorHabilitado
        MshFacturas.BackColor = ColorHabilitado
    End If
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            cmdAnular.Enabled = False
            btnGrabar.Enabled = True
            btnCancelar.Enabled = True
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            cmdAnular.Enabled = False
            btnGrabar.Enabled = True
            btnCancelar.Enabled = True
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            cmdAnular.Enabled = False
            btnGrabar.Enabled = False
            btnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgAnular
            BtnNuevo.Enabled = True
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            cmdAnular.Enabled = False
            btnGrabar.Enabled = False
            btnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
            cmdAnular.Enabled = True
            btnGrabar.Enabled = False
            btnCancelar.Enabled = True
            Publimensaje = ""
        Case ConfigBotones.cfgCancelar
            Publimensaje = ""
            BtnNuevo.Enabled = True
            btnModificar.Enabled = True
            btnEliminar.Enabled = False
            cmdAnular.Enabled = False
            btnGrabar.Enabled = False
            btnCancelar.Enabled = False
    End Select
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cboCliente_Change()
    FormatoFacturas
    lblvoucher.Visible = False
    lblV.Visible = False
    If cboCliente.ListIndex > -1 Then
        CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
    End If
End Sub

Private Sub cmdAnular_Click()
Dim SQL As String
Dim NumLetra As String
    If GridSel = 1 Then NumLetra = mshLetras.TextMatrix(mshLetras.row, 1)
    If GridSel = 2 Then NumLetra = MshLetrasRen.TextMatrix(MshLetrasRen.row, 1)
    If MsgBox("¿Seguro que desea anular la Letra " & NumLetra & "?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
        SQL = "update letra set codestado = 'AN' where numero = '" & NumLetra & "'"
        oConexionMYSQL.Execute SQL
        SQL = "delete from factura_letra where letra = '" & NumLetra & "'"
        oConexionMYSQL.Execute SQL
        If cboCliente.ListIndex > -1 Then
            CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
        End If
        CargarLetrasRenovadas NumLetra
        lblvoucher.Visible = False
        lblV.Visible = False
    End If
End Sub
Private Sub cmdCancel_Click()
    Dim NumLetra As String, Abono As Double, Interes As Double
    Dim SQL As String, rpta As Variant
    If GridSel = 1 Then NumLetra = Trim(mshLetras.TextMatrix(mshLetras.row, 1))
    If GridSel = 2 Then NumLetra = Trim(MshLetrasRen.TextMatrix(MshLetrasRen.row, 1))
    rpta = InputBox("Ingrese el Monto de Abono de la Letra", "Abono", 0)
    If rpta = "" Then Exit Sub
    Abono = CDbl(rpta)
    rpta = InputBox("Ingrese el Monto de Interés del Banco", "Interés Banco", 0)
    If rpta = "" Then Exit Sub
    Interes = CDbl(rpta)
    SQL = "update letra set abono = " & CDbl(Abono) & ",codestado ='CA' where numero = '" & NumLetra & "'"
    If oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False) Then
        SQL = "DELETE FROM interesbcoletra WHERE LETRA = '" & NumLetra & "'"
        oConexionMYSQL.Execute SQL
        SQL = "insert into interesbcoletra(interes,letra) values (" & CDbl(Interes) & ",'" & NumLetra & "')"
        oConexionMYSQL.Execute SQL
        MsgBox "Se canceló la Letra", vbInformation, "NOVPeru"
        If cboCliente.ListIndex > -1 Then CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
        CargarLetrasRenovadas NumLetra
    End If
End Sub
Private Sub cmdFacs_Click()
    frmFacturasCliente.Show
End Sub
Private Sub cmdRenovar_Click()
Dim SQL As String, NumLetra As String
Dim FecGiro As String, AbonoRen As Double, ndebito As String
Dim FecVcto As String, Interes As Double, Interesbanco As Double
On Error GoTo CtrlError
    If lblren = "P" Then
        With mshLetras
            If ValidarRenovacion Then
                FecGiro = InputBox("Ingrese la fecha de Giro de la nueva Letra", "FECHA", Date)
                FecGiro = Format(FecGiro, "dd/mm/yyyy")
                If IsDate(FecGiro) Then
                    FecVcto = InputBox("Ingrese la fecha de Vencimiento de la nueva Letra", "FECHA", Date)
                    FecVcto = Format(FecVcto, "dd/mm/yyyy")
                    If IsDate(FecVcto) Then
                        If MsgBox("¿Seguro que desea ingresar los siguientes datos: ?" & Chr(13) & _
                                  "Fecha Giro: " & FecGiro & Chr(13) & "Fecha Vencimiento: " & FecVcto, vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                            NumLetra = MaxCorrelativo
                            SQL = "Insert into letra (numero,fecgiro,fecvcto,importe,abono,codaux,codbco,tipo,Interes," & _
                                 " InteresBco,codestado,ref,usuario) values ('" & NumLetra & "','" & Format(FecGiro, "dd/mm/yyyy") & "','" & _
                                 Format(FecVcto, "dd/mm/yyyy") & "' " & _
                                 "," & CDbl(.TextMatrix(.row, 4)) - CDbl(mePago.Text) & ",0,'" & cboCliente.List(cboCliente.ListIndex, 1) & "', " & _
                                 "'" & txtCodBco & "','R'," & CDbl(meInteres) & "," & CDbl(meIntBco) & ",'EM','" & .TextMatrix(.row, 1) & "','" & strUsuarioId & "')"
                            oConexionMYSQL.Execute SQL
                            
                            SQL = "update letra set abono = " & CDbl(mePago.Text) & " where numero = '" & Trim(.TextMatrix(.row, 1)) & "'"
                            oConexionMYSQL.Execute SQL
                            If ChkND.Value = True Then
                                ndebito = GenNotaDebito(CDbl(meInteres), NumLetra)
                                SQL = "update letra set ndebito = '" & ndebito & "' where numero = '" & Trim(.TextMatrix(.row, 1)) & "'"
                                oConexionMYSQL.Execute SQL
                                ChkND.Value = False
                            End If
                            If cboCliente.ListIndex > -1 Then
                                CargaLetras cboCliente.List(cboCliente.ListIndex, 1)
                            End If
                            CargarLetrasRenovadas .TextMatrix(.row, 1)
                        End If
                    End If
                End If
            End If
        End With
        ModoFormulario modConsulta
    Else
        With MshLetrasRen
            If .TextMatrix(.row, 1) = "" Then
                MsgBox "No se puede renovar una Letra" & vbNewLine & "sin seleccionar una Letra Origen", vbOKOnly + vbCritical, "NOVPeru"
                .SetFocus
                Exit Sub
            End If
            FecGiro = InputBox("Ingrese la fecha de Giro de la nueva Letra", "FECHA", Date)
            FecGiro = Format(FecGiro, "dd/mm/yyyy")
            If IsDate(FecGiro) Then
                FecVcto = InputBox("Ingrese la fecha de Vencimiento de la nueva Letra", "FECHA", Date)
                FecVcto = Format(FecVcto, "dd/mm/yyyy")
                If IsDate(FecVcto) Then
                    AbonoRen = CDbl(InputBox("Ingrese el Monto a Abonar de la Letra", "Monto a Abonar", 0))
                    Interes = CDbl(InputBox("Ingrese el Monto de Interés para la nueva Letra", "Interés", 0))
                    Interesbanco = CDbl(InputBox("Ingrese el Monto de Interés del Banco para la nueva Letra", "Interés Banco", 0))
                    If MsgBox("¿Seguro que desea ingresar estos datos?" & Chr(13) & _
                              "Fecha Giro Letra: " & FecGiro & Chr(13) & "Fecha Vencimiento Letra: " & FecVcto & Chr(13) & _
                              "Monto Abono Letra: " & FormatNumber(AbonoRen, 2) & Chr(13) & _
                              "Monto Interés para nueva Letra: " & FormatNumber(Interes, 2) & Chr(13) & _
                              "Monto Interés Banco para nueva Letra: " & FormatNumber(Interesbanco, 2), vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                        NumLetra = MaxCorrelativo
                        SQL = "Insert into letra (numero,fecgiro,fecvcto,importe,abono,codaux,codbco,tipo,Interes," & _
                              " InteresBco,codestado,ref,usuario) values ('" & NumLetra & "','" & Format(FecGiro, "dd/mm/yyyy") & "','" & _
                              Format(FecVcto, "dd/mm/yyyy") & "' " & _
                              "," & CDbl(.TextMatrix(.row, 4)) - CDbl(AbonoRen) & ",0,'" & .TextMatrix(.row, 7) & "', " & _
                              "'" & .TextMatrix(.row, 8) & "','R'," & CDbl(Interes) & "," & CDbl(Interesbanco) & ",'EM','" & .TextMatrix(.row, 11) & "','" & strUsuarioId & "')"
                        oConexionMYSQL.Execute SQL
                            
                        SQL = "update letra set abono = " & CDbl(AbonoRen) & " where numero = '" & Trim(.TextMatrix(.row, 1)) & "'"
                        oConexionMYSQL.Execute SQL
                        If ChkND.Value = True Then
                            ndebito = GenNotaDebito(CDbl(Interes), NumLetra)
                            SQL = "update letra set ndebito = '" & ndebito & "' where numero = '" & Trim(.TextMatrix(.row, 1)) & "'"
                            oConexionMYSQL.Execute SQL
                            ChkND.Value = False
                        End If
                        CargarLetrasRenovadas .TextMatrix(.row, 11)
                    End If
                End If
            End If
        End With
    End If
Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Function GenNotaDebito(Interes As Double, NumLetra As String) As String
On Error GoTo CtrlError
    Dim Ident As String, Numdoc As String
    Dim SQL As String, sqlAmarre As String, sqlMovi As String, sqlHist As String
    Dim sqldetalle As String
    Ident = GenerarFolio
    Numdoc = CorrelativoND
    SQL = "Call Insert_ContablesFact('" & Ident & "','" & Left(Numdoc, 5) & "','" & Right(Numdoc, 9) & "', " & _
          " '','N','2','" & Trim(cboCliente.List(cboCliente.ListIndex, 1)) & "', '', " & _
          " '00000000000','E','0','','" & Format(Date, "yyyy/mm/dd") & "','', " & _
          " " & CDbl(Interes / 1.18) & "," & CDbl(Interes) - CDbl(Interes / 1.18) & ",0, " & _
          " " & CDbl(Interes) & ",'','DOCUMENTOS CONTABLES EMITIDOS','','','', " & _
          " '0000',0,'','0','',0,'00');"
    sqlAmarre = "Call Insert_AmarreDoc ('" & Ident & "','08'," & _
                " '" & Format(Date, "yyyy/mm/dd") & "','" & Format(Time, "HH:MM") & "'," & _
                " '00','00000000','NINGUNO', '" & strNombreEmpresa & "'," & _
                " '','1', '" & strUsuarioId & "'," & _
                " '" & strAnoSistema & strMesSistema & "','0'); "
    sqlMovi = "Call Insert_Movi_Doc('" & Ident & "', '" & Format(Date, "yyyy/mm/dd") & "'," & _
              " '" & EMITIDO & "','1','" & strUsuarioId & "'); "
    sqlHist = "Call Insert_HistorialDoc ( '" & Ident & "', '" & EMITIDO & _
              "', '" & DescripcionesdeCodigos("CNUSER", strUsuarioId, "area") & "'," & _
              "'" & Format(Date, "yyyy/mm/dd") & "', '" & strUsuarioId & "');"
    oConexion.EjecutaInsertUpdateDelete sqlAmarre, TIPO_QUERY.insertar, False
    oConexion.EjecutaInsertUpdateDelete sqlMovi, TIPO_QUERY.insertar, False
    oConexion.EjecutaInsertUpdateDelete sqlHist, TIPO_QUERY.insertar, False
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    
    sqldetalle = "Call Insert_DetFact('1','" & Ident & "'," & _
                 "'POR INTERESES DE LETRA VENCIDAS N° " & NumLetra & "',1," & _
                 " " & CDbl(Interes / 1.18) & ",0," & CDbl(Interes / 1.18) & " ,'1','E'," & _
                 "'634','0024','LIMA','LIMA');"
    If (oConexion.EjecutaInsertUpdateDelete(sqldetalle, TIPO_QUERY.insertar, False)) = False Then
        SQL = "Delete from historial_docs where identificador = '" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        SQL = "Delete from movi_documento where identificador = '" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        SQL = "Delete from documento_contables where identificador = '" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        SQL = "Delete from amarre_documento where identificador = '" & Ident & "'"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        GenNotaDebito = ""
        MsgBox " Se generó un error al momento de guardar los datos " & vbNewLine & _
               " revise los datos y vuelva a intentarlo.", vbCritical, gsNomSW
        Exit Function
    Else
        UpdateSerTar "0024", Trim(cboCliente.List(cboCliente.ListIndex, 1)), 0, True
        GenNotaDebito = Numdoc
    End If
Exit Function
CtrlError:
    MsgBox err.Description, vbCritical, "Error generando Nota de Débito"
End Function

Private Function UpdateSerTar(codigo As String, codAux As String, Monto As Double, filtro As Boolean)
    Dim SQL As String
    Dim sqlupdate As String
    Dim cuota As Integer
    Dim rsservicio As MYSQL_RS
     
    SQL = "Select * from serv_tarif where " & _
          "IDSerTar = '" & codigo & "' and codaux = '" & codAux & "'"
    Set rsservicio = oConexion.EjecutaSelectRS(SQL)
    If Not rsservicio.EOF Then
        If rsservicio.Fields("Num_Cuota") > 1 Then
            If filtro = True Then
                cuota = CEN(rsservicio.Fields("Cuota")) + 1
                sqlupdate = "Call Update_MontosSerTar ('" & codigo & "', '" & Monto & "','" & cuota & "' );"
                oConexion.EjecutaInsertUpdateDelete sqlupdate, TIPO_QUERY.Modificar, False
            Else
                sqlupdate = "Call Update_MontosSerTar ('" & codigo & "', '" & Monto & "','" & CEN(rsservicio.Fields("Cuota")) & "' );"
                oConexion.EjecutaInsertUpdateDelete sqlupdate, TIPO_QUERY.Modificar, False
            End If
        End If
    End If
    Set rsservicio = Nothing
End Function

Private Function CorrelativoND() As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    
    SQL = "Select CORRELATIVO,SERIE from opcfact where CODIGO = '08'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF Then
        CorrelativoND = RQ.Fields("SERIE") & "-" & RQ.Fields("CORRELATIVO")
        SerieND = RQ.Fields("SERIE")
    End If
    Set RQ = Nothing
End Function

Sub CargarLetrasRenovadas(NumLetra As String)
    Dim SQL As String
    Dim rsletras As MYSQL_RS
    With MshLetrasRen
        .Clear
        .Rows = 1
        .Cols = 14
        .ForeColorFixed = &H404000
        
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = "Letra"
        
        .ColWidth(2) = 1500
        .TextMatrix(0, 2) = "Fec. Giro"
            
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Fec. Vcto."
        
        .ColWidth(4) = 1350
        .TextMatrix(0, 4) = "Importe"
    
        .ColWidth(5) = 1350
        .TextMatrix(0, 5) = "Abono"
        
        .ColWidth(6) = 1350
        .TextMatrix(0, 6) = Space(8) + "Saldo"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "cli"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = "codbanco"
        
        .ColWidth(9) = 1350
        .TextMatrix(0, 9) = "Interes"
    
        .ColWidth(10) = 1350
        .TextMatrix(0, 10) = "Interes Bco"
        
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = "REF"
        
        .ColWidth(12) = 1400
        .TextMatrix(0, 12) = "NDébito"
        
        .ColWidth(13) = 0
        .TextMatrix(0, 13) = "estado"
        SQL = " Select l.* from letra l where l.ref='" & NumLetra & "' and l.codestado <> 'EL' order by numero"
        Set rsletras = oConexion.EjecutaSelectRS(SQL)
        Do While Not rsletras.EOF
            .Rows = .Rows + 1
            If .Rows = 2 Then
                .FixedRows = 1
            End If
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = rsletras.Fields("numero")
            .TextMatrix(.Rows - 1, 2) = rsletras.Fields("fecgiro")
            .TextMatrix(.Rows - 1, 3) = rsletras.Fields("fecvcto")
            .TextMatrix(.Rows - 1, 4) = FormatNumber(rsletras.Fields("importe"), 2)
            .TextMatrix(.Rows - 1, 5) = FormatNumber(rsletras.Fields("abono"), 2)
            .TextMatrix(.Rows - 1, 6) = FormatNumber(rsletras.Fields("importe") - rsletras.Fields("abono"), 2)
            .TextMatrix(.Rows - 1, 7) = rsletras.Fields("codaux")
            .TextMatrix(.Rows - 1, 8) = rsletras.Fields("codbco")
            .TextMatrix(.Rows - 1, 9) = FormatNumber(rsletras.Fields("interes"), 2)
            .TextMatrix(.Rows - 1, 10) = FormatNumber(rsletras.Fields("interesbco"), 2)
            .TextMatrix(.Rows - 1, 11) = NumLetra
            .TextMatrix(.Rows - 1, 12) = rsletras.Fields("ndebito")
            .TextMatrix(.Rows - 1, 13) = rsletras.Fields("codestado")
            If Trim(rsletras.Fields("codestado")) = "AN" Then
                Dim I As Integer
                For I = 1 To .Cols - 1
                    .Col = I
                    .row = .Rows - 1
                    .CellBackColor = vbGreen
                Next
            End If
            Dim RQ As MYSQL_RS
            SQL = "SELECT * FROM interesbcoletra WHERE LETRA = '" & rsletras.Fields("numero") & "'"
            Set RQ = oConexion.EjecutaSelectRS(SQL)
            If Not RQ.EOF() Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = rsletras.Fields("numero")
                .TextMatrix(.Rows - 1, 2) = rsletras.Fields("fecgiro")
                .TextMatrix(.Rows - 1, 3) = rsletras.Fields("fecvcto")
                .TextMatrix(.Rows - 1, 4) = FormatNumber(rsletras.Fields("importe"), 2)
                .TextMatrix(.Rows - 1, 5) = FormatNumber(rsletras.Fields("abono"), 2)
                .TextMatrix(.Rows - 1, 6) = FormatNumber(rsletras.Fields("importe") - rsletras.Fields("abono"), 2)
                .TextMatrix(.Rows - 1, 7) = rsletras.Fields("codaux")
                .TextMatrix(.Rows - 1, 8) = rsletras.Fields("codbco")
                .TextMatrix(.Rows - 1, 9) = "0.00"
                .TextMatrix(.Rows - 1, 10) = FormatNumber(RQ.Fields("interes"), 2)
                .TextMatrix(.Rows - 1, 11) = NumLetra
                .TextMatrix(.Rows - 1, 12) = ""
                .TextMatrix(.Rows - 1, 13) = rsletras.Fields("codestado")
            End If
            rsletras.MoveNext
        Loop
        Set rsletras = Nothing
        Set RQ = Nothing
    End With
End Sub

Function ValidarRenovacion()
    Dim I As Integer
    ValidarRenovacion = True
    If val(mePago) = 0 Then
        MsgBox "Debe ingresar un Monto de Abono para la Letra", vbOKOnly + vbCritical, "NOVPeru"
        mePago.Enabled = True
        mePago.BackColor = ColorHabilitado
        mePago.SetFocus
        ValidarRenovacion = False
        Exit Function
    End If
    If val(meInteres) = 0 Then
        MsgBox "Debe ingresar un Monto de Interés", vbOKOnly + vbCritical, "NOVPeru"
        meInteres.Enabled = True
        meInteres.BackColor = ColorHabilitado
        meInteres.SetFocus
        ValidarRenovacion = False
        Exit Function
    End If
    If mshLetras.TextMatrix(mshLetras.row, 1) = "" Then
        MsgBox "No se puede renovar una Letra" & vbNewLine & "sin seleccionar una Letra Origen", vbOKOnly + vbCritical, "NOVPeru"
        mshLetras.SetFocus
        ValidarRenovacion = False
        Exit Function
    End If
End Function

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    txtNumLetra.Locked = True
    meFecGiro.Enabled = False
    meFecvcto.Enabled = False
    cboCliente.Enabled = True
    ModoFormulario modAccion
    Clientes
    CargaLetras "0"
    CargarLetrasRenovadas "0"
    FormatoFacturas
End Sub

Sub FormatoFacturas()
    With MshFacturas
        .Clear
        .Rows = 2
        .Cols = 7
        .ColWidth(0) = 0
        .TextMatrix(0, 0) = "Folio"
        .ColWidth(1) = 1600
        .TextMatrix(0, 1) = "Número"
        .ColAlignment(1) = 4
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "F.Emisión"
        .ColAlignment(2) = 4
        .ColWidth(3) = 1300
        .TextMatrix(0, 3) = "Monto"
        .ColAlignment(3) = 6
        .ColWidth(4) = 0
        .TextMatrix(0, 4) = "marca"
        .ColWidth(5) = 0
        .TextMatrix(0, 5) = "ccHFM"
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = "mon"
    End With
End Sub

Private Sub meFecGiro_GotFocus()
    meFecGiro.SelStart = 0
    meFecGiro.SelLength = 10
End Sub

Private Sub meFecGiro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        meFecvcto.SetFocus
    End If
End Sub

Private Sub meFecvcto_GotFocus()
    meFecvcto.SelStart = 0
    meFecvcto.SelLength = 10
End Sub

Private Sub meFecvcto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboCliente.SetFocus
    End If
End Sub

Private Sub MshFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQL As String, I As Integer
    If lblModo = "modificar" And MshFacturas.BackColor = ColorHabilitado Then
        With MshFacturas
            If KeyCode = 46 Then
                SQL = "delete from factura_letra where letra = '" & Trim(txtNumLetra) & "' and identificador = '" & .TextMatrix(.row, 0) & "'"
                oConexionMYSQL.Execute SQL
                If .Rows = 2 Then
                    For I = 1 To .Cols - 1
                        .TextMatrix(1, I) = ""
                    Next
                End If
                If .Rows > 2 Then
                    .RemoveItem .row
                End If
            End If
        End With
    End If
End Sub

Private Sub mshLetras_RowColChange()
    With mshLetras
        If .Rows > 1 Then
            GridSel = 1
            Dim I As Integer
            lblren = "P"
            If lblModo = "modificar" Or lblModo = "" Or lblModo = "Acción" Then txtNumLetra = Trim(.TextMatrix(.row, 1))
            meFecGiro = Format(.TextMatrix(.row, 2), "dd/mm/yyyy")
            meFecvcto = Format(.TextMatrix(.row, 3), "dd/mm/yyyy")
            mePago = FormatNumber(.TextMatrix(.row, 5), 2)
            lblvoucher = Left(.TextMatrix(.row, 10), 4) & "-" & Right(.TextMatrix(.row, 10), 6)
            txtCodBco = Trim(.TextMatrix(.row, 8))
            For I = 0 To cboCliente.ListCount - 1
                If .TextMatrix(.row, 7) = cboCliente.List(I, 1) Then
                    cboCliente.ListIndex = I
                    Exit For
                End If
            Next
            CargarFacturasPorLetra Trim(.TextMatrix(.row, 1))
            CargarLetrasRenovadas Trim(.TextMatrix(.row, 1))
            If VerificaRenovacion(txtNumLetra, "P") Then
                cmdRenovar.Enabled = False
                cmdCancel.Enabled = False
                btnModificar.Enabled = False
                btnGrabar.Enabled = False
                btnCancelar.Enabled = False
                BtnNuevo.Enabled = True
                btnEliminar.Enabled = False
                cmdAnular.Enabled = False
            Else
                btnModificar.Enabled = True
                BtnNuevo.Enabled = True
                cmdRenovar.Enabled = True
                cmdCancel.Enabled = True
                If lblModo <> "" Then
                    btnEliminar.Enabled = True
                    cmdAnular.Enabled = True
                End If
            End If
            If Trim(.TextMatrix(.row, 10)) <> "" Then
                btnModificar.Enabled = False
                btnEliminar.Enabled = False
                cmdAnular.Enabled = False
            Else
                btnModificar.Enabled = True
                btnEliminar.Enabled = True
                cmdAnular.Enabled = True
            End If
            If Trim(.TextMatrix(.row, 11)) = "AN" Or Trim(.TextMatrix(.row, 11)) = "CA" Then
                cmdRenovar.Enabled = False
                cmdCancel.Enabled = False
                btnModificar.Enabled = False
                btnGrabar.Enabled = False
                btnCancelar.Enabled = False
                BtnNuevo.Enabled = True
                btnEliminar.Enabled = False
                cmdAnular.Enabled = False
            End If
        Else
            cmdRenovar.Enabled = False
            cmdCancel.Enabled = False
        End If
    End With
End Sub

Function VerificaRenovacion(letra As String, Tipo As String) As Boolean
Dim SQL As String
Dim RQ As MYSQL_RS
    VerificaRenovacion = False
    If Tipo = "P" Then
        SQL = "select count(*) as cantidad from letra where ref = '" & letra & "'"
    Else
        SQL = "select abono as cantidad from letra where numero = '" & letra & "'"
    End If
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If RQ.Fields("cantidad") > 0 Then
            VerificaRenovacion = True
        End If
    End If
    Set RQ = Nothing
End Function

Private Sub MshLetrasRen_Click()
    If txt.Visible = True Then txt.Visible = False
End Sub

Private Sub MshLetrasRen_DblClick()
    If HabModRen = True Then
        With MshLetrasRen
            frmModLetraRen.letra = .TextMatrix(.row, 1)
            frmModLetraRen.meFecGiro = .TextMatrix(.row, 2)
            frmModLetraRen.txtCodBco = .TextMatrix(.row, 8)
            frmModLetraRen.meFecvcto = .TextMatrix(.row, 3)
            frmModLetraRen.txtinteres = .TextMatrix(.row, 9)
            frmModLetraRen.txtinteresbco = .TextMatrix(.row, 10)
            frmModLetraRen.Show
        End With
    End If
End Sub

Private Sub MshLetrasRen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        With MshLetrasRen
            If VerificaLetra(.TextMatrix(.Rowsel, 11), .TextMatrix(.Rowsel, 1)) = True Then
                If .Rows > 1 And Trim(.TextMatrix(.Rowsel, 1)) <> "" Then
                    .RemoveItem .Rowsel
                End If
            End If
        End With
    End If
End Sub

Function VerificaLetra(letra As String, Numero As String) As Boolean
Dim SQL As String
Dim RQ As MYSQL_RS
    VerificaLetra = False
    SQL = "select numero from letra where tipo = 'R' and ref = '" & letra & "' order by numero desc limit 1"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If Numero = RQ.Fields("numero") Then
            VerificaLetra = True
        End If
    End If
    Set RQ = Nothing
End Function

Private Sub MshLetrasRen_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        With MshLetrasRen
            If .ColSel = 12 And .Rows > 1 And (.TextMatrix(.row, 13) <> "CA" Or .TextMatrix(.row, 13) = "AN") Then
                txt.Left = .CellLeft + .Left
                txt.Top = .CellTop + .Top
                txt.Width = .CellWidth
                txt.Height = .CellHeight
                txt.Text = ""
                txt.Visible = True
                txt.SetFocus
                txt.Text = Chr$(KeyAscii)
            End If
        End With
    End If
End Sub

Private Sub MshLetrasRen_RowColChange()
    meInteres = "0.00"
    meIntBco = "0.00"
    mePago = "0.00"
    GridSel = 2
    If VerificaRenovacion(MshLetrasRen.TextMatrix(MshLetrasRen.row, 1), "R") Then
        cmdRenovar.Enabled = False
        cmdCancel.Enabled = False
        ChkND.Locked = True
        meInteres.BackColor = ColorDeshabilitado
        meIntBco.BackColor = ColorDeshabilitado
        mePago.BackColor = ColorDeshabilitado
        meInteres.Enabled = False
        meIntBco.Enabled = False
        mePago.Enabled = False
        HabModRen = False
    Else
        If MshLetrasRen.TextMatrix(MshLetrasRen.row, 13) = "CA" Or MshLetrasRen.TextMatrix(MshLetrasRen.row, 13) = "AN" Then
            cmdRenovar.Enabled = False
            cmdCancel.Enabled = False
            ChkND.Locked = True
            meInteres.BackColor = ColorDeshabilitado
            meIntBco.BackColor = ColorDeshabilitado
            mePago.BackColor = ColorDeshabilitado
            meInteres.Enabled = False
            meIntBco.Enabled = False
            mePago.Enabled = False
            HabModRen = False
        Else
            lblren = "R"
            HabModRen = True
            cmdRenovar.Enabled = True
            cmdCancel.Enabled = True
            ChkND.Locked = False
            meInteres.BackColor = ColorHabilitado
            meIntBco.BackColor = ColorHabilitado
            mePago.BackColor = ColorHabilitado
            meInteres.Enabled = True
            meIntBco.Enabled = True
            mePago.Enabled = True
        End If
    End If
End Sub

Private Sub txt_GotFocus()
    mark txt
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        txt.Visible = False
        MshLetrasRen.Col = MshLetrasRen.ColSel
        MshLetrasRen.SetFocus
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With MshLetrasRen
            .TextMatrix(.Rowsel, .ColSel) = UCase(Trim(txt.Text))
            txt.Visible = False
            GrabarNDebito .TextMatrix(.Rowsel, 1), .TextMatrix(.Rowsel, 12)
            .Col = 12: .row = .Rowsel
            .SetFocus
        End With
    End If
End Sub

Sub GrabarNDebito(letra As String, ndebito As String)
    Dim SQL As String
    SQL = "update letra set ndebito='" & ndebito & "' where numero = '" & letra & "'"
    oConexionMYSQL.Execute SQL
End Sub

Private Sub txtNumLetra_GotFocus()
    txtNumLetra.SelStart = 0
    txtNumLetra.SelLength = Len(txtNumLetra)
End Sub

Private Sub txtNumLetra_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 And lblModo = "nuevo" Then
        meFecGiro.SetFocus
    End If
End Sub

Function MaxCorrelativo() As String
    Dim RQ As MYSQL_RS
    Dim SQL As String
    SQL = "Select max(numero) as corr from letra"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If IsNull(RQ.Fields("corr")) = False Then
        MaxCorrelativo = Right("000000" & val(Trim(RQ.Fields("corr"))) + 1, 6)
    Else
        MaxCorrelativo = "000001"
    End If
    Set RQ = Nothing
End Function

Public Sub TotFacturas()
Dim SumFact As Double, I As Integer
    With MshFacturas
        For I = 1 To .Rows - 1
            SumFact = SumFact + CDbl(.TextMatrix(I, 3))
        Next
        lblimporte = FormatNumber(SumFact, 2)
    End With
End Sub
