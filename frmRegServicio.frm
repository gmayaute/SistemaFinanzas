VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRegServicio 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Servicios y Tarifas"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12645
   Icon            =   "frmRegServicio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   12645
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   7245
      Left            =   -30
      TabIndex        =   10
      Top             =   -90
      Width           =   12645
      Begin VB.TextBox TxtCriterio 
         Height          =   285
         Left            =   2700
         TabIndex        =   35
         Top             =   6345
         Width           =   4650
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxServicios 
         Height          =   4365
         Left            =   180
         TabIndex        =   33
         Top             =   1860
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   7699
         _Version        =   393216
         BackColorBkg    =   8421504
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSMask.MaskEdBox meFecha 
         Height          =   315
         Left            =   10530
         TabIndex        =   2
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Proyecto1.chameleonButton btnModificar 
         Height          =   345
         Left            =   1380
         TabIndex        =   11
         ToolTipText     =   "Modificar"
         Top             =   6780
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
         MICON           =   "frmRegServicio.frx":030A
         PICN            =   "frmRegServicio.frx":0326
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
         Left            =   12090
         TabIndex        =   12
         Top             =   6690
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmRegServicio.frx":0754
         PICN            =   "frmRegServicio.frx":0770
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
         Left            =   5760
         TabIndex        =   13
         ToolTipText     =   "Guardar"
         Top             =   6780
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
         MICON           =   "frmRegServicio.frx":0B36
         PICN            =   "frmRegServicio.frx":0B52
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnNuevo 
         Height          =   345
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "Nuevo"
         Top             =   6750
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmRegServicio.frx":0F94
         PICN            =   "frmRegServicio.frx":0FB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnCancelar 
         Height          =   345
         Left            =   5280
         TabIndex        =   15
         ToolTipText     =   "Deshacer"
         Top             =   6780
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
         MICON           =   "frmRegServicio.frx":131A
         PICN            =   "frmRegServicio.frx":1336
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
         Left            =   2730
         TabIndex        =   16
         Top             =   6780
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Eliminar"
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
         MICON           =   "frmRegServicio.frx":1878
         PICN            =   "frmRegServicio.frx":1894
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
         Left            =   11550
         TabIndex        =   17
         Top             =   6690
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmRegServicio.frx":1CD6
         PICN            =   "frmRegServicio.frx":1CF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.ComboBox cboCampos 
         Height          =   315
         Left            =   180
         TabIndex        =   36
         Top             =   6330
         Width           =   2415
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4260;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CheckBox chkPorcentaje 
         Height          =   285
         Left            =   3240
         TabIndex        =   34
         Top             =   1500
         Width           =   1905
         VariousPropertyBits=   746588179
         BackColor       =   -2147483644
         ForeColor       =   16777215
         DisplayStyle    =   4
         Size            =   "3360;503"
         Value           =   "0"
         Caption         =   "Porcentaje"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtContrato 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   210
         Width           =   2685
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   50
         Size            =   "4736;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtTotal 
         Height          =   315
         Left            =   11040
         TabIndex        =   9
         Top             =   1470
         Width           =   1395
         VariousPropertyBits=   746604571
         MaxLength       =   11
         Size            =   "2461;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtMonto 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   1470
         Width           =   1785
         VariousPropertyBits=   746604571
         MaxLength       =   11
         Size            =   "3149;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtCuotas 
         Height          =   315
         Left            =   6420
         TabIndex        =   8
         Top             =   1470
         Width           =   525
         VariousPropertyBits=   746604571
         MaxLength       =   3
         Size            =   "926;556"
         Value           =   "1"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuotas:"
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
         TabIndex        =   32
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Label lblTarifa 
         BackColor       =   &H00C0C0C0&
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
         Left            =   11040
         TabIndex        =   31
         Top             =   960
         Width           =   1395
      End
      Begin MSForms.TextBox txtTarifa 
         Height          =   315
         Left            =   10530
         TabIndex        =   6
         Top             =   960
         Width           =   465
         VariousPropertyBits=   746604571
         MaxLength       =   2
         Size            =   "820;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarifa:"
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
         Left            =   9330
         TabIndex        =   30
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label lblServicio 
         BackColor       =   &H00C0C0C0&
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
         Height          =   555
         Left            =   3240
         TabIndex        =   29
         Top             =   900
         Width           =   6045
      End
      Begin MSForms.TextBox txtServicio 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   930
         Width           =   1785
         VariousPropertyBits=   746604571
         MaxLength       =   6
         Size            =   "3149;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicio:"
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
         Left            =   180
         TabIndex        =   28
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total:"
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
         Left            =   9330
         TabIndex        =   27
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto:"
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
         Left            =   180
         TabIndex        =   26
         Top             =   1470
         Width           =   1155
      End
      Begin VB.Label Label4 
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
         Left            =   9330
         TabIndex        =   25
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contrato:"
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
         Left            =   180
         TabIndex        =   24
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Señores:"
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
         Left            =   180
         TabIndex        =   23
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Código/Ruc:"
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
         Left            =   9330
         TabIndex        =   22
         Top             =   600
         Width           =   1155
      End
      Begin MSForms.TextBox txtCodigo 
         Height          =   315
         Left            =   10530
         TabIndex        =   4
         Top             =   600
         Width           =   1905
         VariousPropertyBits=   746604571
         MaxLength       =   14
         Size            =   "3360;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cboMoneda 
         Height          =   315
         Left            =   5130
         TabIndex        =   1
         Top             =   210
         Width           =   1815
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "3201;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
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
         Left            =   4170
         TabIndex        =   20
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lblModo 
         Caption         =   "Acción"
         Height          =   315
         Left            =   7080
         TabIndex        =   19
         Top             =   5550
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00C0C0C0&
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
         Left            =   10530
         TabIndex        =   18
         Top             =   1470
         Width           =   465
      End
      Begin MSForms.ListBox lstClientes 
         Height          =   675
         Left            =   1380
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   7905
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "13944;1191"
         MatchEntry      =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCliente 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   570
         Width           =   7905
         VariousPropertyBits=   746604571
         MaxLength       =   11
         Size            =   "13944;556"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmRegServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As New FrmConsultas
Dim fila As Integer
Private IdeSerTar As String

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    
    DesplazarxGrilla fila
    flxServicios.Col = 1
    flxServicios.ColSel = 11
    
End Sub

Private Sub btnEliminar_Click()
    Dim SQL As String
    Dim RES As Integer
    RES = MsgBox("Esta seguro de Eliminar el Servicio?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        SQL = "Call Delete_ServTar('" & IdeSerTar & "','" & txtCodigo & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
        fila = 1
        ModoFormulario modAccion
    End If
End Sub

Private Sub btnGrabar_Click()
    If lblModo = "Modificar" Then
        If ValidarData Then
            fila = flxServicios.row
            Actualizar
            ModoFormulario modAccion
            flxServicios.SetFocus
            'SendKeys "{LEFT}"
            Call keybd_event(vbKeyLeft, 0, 0, 0)
        End If
    End If
    If lblModo = "Nuevo" Then
        If ValidarData Then
            fila = flxServicios.row
            Grabar
            ModoFormulario modAccion
            flxServicios.SetFocus
            'SendKeys "{LEFT}"
            Call keybd_event(vbKeyLeft, 0, 0, 0)
        End If
    End If
End Sub

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    ModoFormulario modNuevo
    MaxId
End Sub

Private Sub btnReporte_Click()
Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    oReporte.Titulo = "REPORTE DE SERVICIOS Y TARIFAS"
    oReporte.Reporte = "Rep_Registros_Servicios.rpt"
    oReporte.sp_Reporte_Registros_Servicios
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cboCampos_Change()
    TxtCriterio = ""
    Call TxtCriterio_Change
End Sub

Private Sub cboMoneda_Change()
    If cboMoneda.ListIndex = 1 Then lblMoneda = " S/. "
    If cboMoneda.ListIndex = 2 Then lblMoneda = " $$ "
    If cboMoneda.ListIndex = 0 Then lblMoneda = Empty
End Sub

Private Sub flxResultados_Click()
    DesplazarxGrilla fila
End Sub

Private Sub cboMoneda_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        meFecha.SetFocus
    End If
End Sub

Private Sub chkPorcentaje_Click()
    If chkPorcentaje.Value = True Then
        txtCuotas.Enabled = False
        txtTotal.Enabled = False
    Else
        txtCuotas.Enabled = True
        txtTotal.Enabled = True
    End If
End Sub

Private Sub flxServicios_Click()
    If flxServicios.Rows > 0 Then
        fila = flxServicios.row
        DesplazarxGrilla fila
    End If
End Sub

Private Sub flxServicios_DblClick()
    btnModificar_Click
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call WheelHook(frmRegServicio)
    Clientes lstClientes
    moneda cboMoneda
    LlenaCboCampos
    strTipoAuxiliar = "2"
    meFecha = Date
    fila = 1
    ModoFormulario modAccion
    Set oConsulta = New FrmConsultas
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Set oConsulta = Nothing
    Set oReporte = Nothing
End Sub

Private Sub lstClientes_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtCliente = lstClientes.List(lstClientes.ListIndex, 0)
        txtCodigo = lstClientes.List(lstClientes.ListIndex, 1)
        lstClientes.Visible = False
        txtCodigo.SetFocus
    End If
    If KeyCode.Value = 27 Then
        If lstClientes.Visible = True Then
            lstClientes.Visible = False
            txtCliente.SetFocus
        End If
    End If
End Sub

Private Sub LimpiarDatos()
    txtContrato = Empty
    txtCodigo = Empty
    txtCliente = Empty
    txtServicio = Empty
    txtTarifa = Empty
    txtMonto = Empty
    txtCuotas = Empty
    txtTotal = Empty
    lblServicio = Empty
    lblTarifa = Empty
    lstClientes.Clear 'x Ver
    chkPorcentaje.Value = False
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtContrato.Locked = valor
    txtCodigo.Locked = valor
    txtServicio.Locked = valor
    txtTarifa.Locked = valor
    txtMonto.Locked = valor
    txtCuotas.Locked = valor
    txtTotal.Locked = valor
    txtCliente.Locked = valor
    meFecha.Enabled = Not valor
    cboMoneda.Locked = valor
    flxServicios.Enabled = valor
    chkPorcentaje.Enabled = Not valor
    If valor = True Then
        txtContrato.BackColor = ColorDeshabilitado
        txtCodigo.BackColor = ColorDeshabilitado
        txtServicio.BackColor = ColorDeshabilitado
        txtTarifa.BackColor = ColorDeshabilitado
        txtMonto.BackColor = ColorDeshabilitado
        txtCuotas.BackColor = ColorDeshabilitado
        txtTotal.BackColor = ColorDeshabilitado
        txtCliente.BackColor = ColorDeshabilitado
        meFecha.BackColor = ColorDeshabilitado
        cboMoneda.BackColor = ColorDeshabilitado
        flxServicios.BackColor = ColorHabilitado
    Else
        txtContrato.BackColor = ColorHabilitado
        txtCodigo.BackColor = ColorHabilitado
        txtServicio.BackColor = ColorHabilitado
        txtTarifa.BackColor = ColorHabilitado
        txtMonto.BackColor = ColorHabilitado
        txtCuotas.BackColor = ColorHabilitado
        txtTotal.BackColor = ColorHabilitado
        txtCliente.BackColor = ColorHabilitado
        meFecha.BackColor = ColorHabilitado
        cboMoneda.BackColor = ColorHabilitado
        flxServicios.BackColor = ColorDeshabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             ConfigGrilla
             If TxtCriterio.Text = "" Then
                LlenarGrilla 1, 1
             Else
                LlenarGrilla cboCampos.List(cboCampos.ListIndex, 1), TxtCriterio.Text
             End If
             flxServicios.row = fila
             DesplazarxGrilla fila
             'flxServicios.ColSel = 11
             lblModo = "Acción"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
             Clientes lstClientes
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             meFecha = Date
             lblModo = "Nuevo"
             BloqueoControles False
             Clientes lstClientes
             moneda cboMoneda
             ConfigurarBotones cfgNuevo
             txtContrato.SetFocus
             Exit Sub
        Case ModoForm.modConsulta
             ConfigGrilla
             LlenarGrilla 1, 1
             lblModo = "Consulta"
             BloqueoControles True
             txtCliente.Locked = True
             txtCodigo.Locked = True
             txtCliente.BackColor = ColorDeshabilitado
             txtCodigo.BackColor = ColorDeshabilitado
             ConfigurarBotones cfgGrabar
         Case ModoForm.modEditar
              lblModo = "Modificar"
              BloqueoControles False
              Clientes lstClientes
              txtCodigo.Locked = True
              txtCliente.Locked = True
              txtCliente.BackColor = ColorDeshabilitado
              txtCodigo.BackColor = ColorDeshabilitado
              ConfigurarBotones cfgModificar
              txtContrato.SetFocus
         Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnGrabar.Enabled = True
            btnReporte.Enabled = False
            btnEliminar.Enabled = False
            btnCancelar.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnGrabar.Enabled = True
            btnReporte.Enabled = False
            btnEliminar.Enabled = False
            btnCancelar.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnReporte.Enabled = False
            btnEliminar.Enabled = False
            btnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgGrabar
            btnEliminar.Enabled = True
            BtnNuevo.Enabled = True
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnReporte.Enabled = True
            btnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo"
                     ModoFormulario modAccion
                     BtnNuevo.Enabled = True
                     btnModificar.Enabled = True
                     btnEliminar.Enabled = True
                     btnGrabar.Enabled = False
                     btnReporte.Enabled = False
                     btnCancelar.Enabled = False
                     DesplazarxGrilla fila
                Case "Modificar"
                    ConfigurarBotones cfgGrabar
                    lblModo = "Acción"
                    BloqueoControles True
                    btnGrabar.Enabled = False
            End Select
    End Select
End Sub

Private Sub meFecha_GotFocus()
    mark1 meFecha
End Sub

Private Sub meFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCliente.SetFocus
    End If
End Sub

Private Sub txtCliente_GotFocus()
    mark1 txtCliente
End Sub

Private Sub txtCliente_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
Dim I As Integer
    If KeyCode.Value = vbKeyDown Then
        txtCliente.SelStart = 0
        txtCliente.SelLength = 0
        'SendKeys "{LEFT}"
        Call keybd_event(vbKeyLeft, 0, 0, 0)
        lstClientes.Visible = True
        lstClientes.SetFocus
        'SendKeys "{UP}"
        Call keybd_event(vbKeyUp, 0, 0, 0)
        lstClientes.ListIndex = ItemLista
    End If
    If KeyCode.Value = 13 Then
        If txtCliente <> Empty Then
            txtCodigo = lstClientes.List(ItemLista, 1)
        Else
            txtCliente = "VARIOS"
            txtCodigo = BuscaenLista("", txtCliente)
            txtServicio.SetFocus
        End If
    End If
End Sub

Private Function BuscaenLista(codigo As String, Cliente As String) As String
    Dim I As Integer
    If Cliente <> "" Then
        For I = 1 To lstClientes.ListCount - 1
            If Cliente = lstClientes.List(I, 0) Then
                BuscaenLista = lstClientes.List(I, 1) 'Devuelve el Nombre del Cliente
                Exit For
            End If
        Next
    End If
    If codigo <> "" Then
        For I = 1 To lstClientes.ListCount - 1
            If codigo = lstClientes.List(I, 1) Then
                BuscaenLista = lstClientes.List(I, 0) 'Devuelve el Codigo del Cliente
                Exit For
            End If
        Next
    End If
End Function

Private Sub txtCliente_KeyPress(KeyAscii As MSForms.ReturnInteger)
  AutoComplete txtCliente, KeyAscii, lstClientes
End Sub

Private Sub txtCodigo_GotFocus()
    mark1 txtCodigo
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 5
            .pCol = 0: .pAnchoCol = 1200
            .pCol = 1: .pAnchoCol = 4500
            .pTitulo = "Códigos de Clientes"
            .pForm = FORM_REGSERVICIO
            .pCaso = Label_Descrip_Auxil
            .Show
        End With
    End If
    If KeyCode = 13 Then
        
        txtCodigo = Trim(txtCodigo)
        
        txtCliente = DescripcionesdeCodigos("AUXILIARES", Trim(txtCodigo), strTipoAuxiliar, "Descrip")
    End If
End Sub

Private Sub Clientes(Lista As MSForms.ListBox)
    Dim SQL As String
    Dim I As Integer
    Dim rscli As MYSQL_RS
    SQL = "clientes order by descrip"
    Set rscli = oConexion.EjecutaSelect(SQL)
    If rscli.RecordCount = 0 Then Exit Sub
    Lista.Clear
    I = 0
    Do While Not rscli.EOF
        Lista.AddItem CE(rscli.Fields("descrip"))
        Lista.List(I, 1) = CE(rscli.Fields("codigo"))
        Lista.List(I, 2) = CE(rscli.Fields("ruc"))
        I = I + 1
        rscli.MoveNext
    Loop
    Set rscli = Nothing
End Sub

Private Sub txtContrato_GotFocus()
    mark1 txtContrato
End Sub

Private Sub TxtCriterio_Change()
    Dim filtro As String
    filtro = TxtCriterio.Text
    Me.MousePointer = vbHourglass
    LlenarGrilla cboCampos.List(cboCampos.ListIndex, 1), filtro
    Me.MousePointer = vbNormal
    DesplazarxGrilla 1
End Sub

Private Sub txtCuotas_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        If txtCuotas <> Empty And txtMonto <> Empty Then
            If IsNumeric(txtCuotas) And IsNumeric(txtMonto) Then
                txtTotal = CalculaTotal(txtMonto, txtCuotas)
            Else
                txtTotal = 0
            End If
        End If
    End If
End Sub

Private Function CalculaTotal(Monto As Double, NumCuotas As Double) As Double
    CalculaTotal = Monto * NumCuotas
End Function

Private Sub txtCuotas_LostFocus()
    If txtCuotas <> Empty And txtMonto <> Empty Then
        If IsNumeric(txtCuotas) And IsNumeric(txtMonto) Then
            txtTotal = FormatNumber(CalculaTotal(txtMonto, txtCuotas), 2)
        Else
            txtTotal = 0
        End If
    Else
        txtTotal = 0
    End If
    If txtCuotas = Empty Then
        txtCuotas = 1
        txtTotal = FormatNumber(CalculaTotal(IIf(txtMonto = "", 0, txtMonto), txtCuotas), 2)
    End If
End Sub

Private Sub txtMonto_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        If txtMonto <> Empty Then
              txtMonto = FormatNumber(txtMonto, 2)
            If txtCuotas <> Empty Then
                If IsNumeric(txtCuotas) And IsNumeric(str(txtMonto)) Then
                    txtTotal = FormatNumber(CalculaTotal(CDbl(txtMonto), str(txtCuotas)), 2)
                Else
                    txtTotal = 0
                End If
            End If
        End If
    Else
        txtTotal = 0
    End If
End Sub

Private Sub txtMonto_LostFocus()
    If txtCuotas <> Empty And txtMonto <> Empty Then
        If IsNumeric(txtCuotas) And IsNumeric(txtMonto) Then
          txtTotal = FormatNumber(CalculaTotal(txtMonto, txtCuotas), 2)
        Else
          txtTotal = 0
        End If
    Else
        txtTotal = 0
    End If
End Sub

Private Sub txtServicio_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 5
            .pCol = 0: .pAnchoCol = 800
            .pCol = 1: .pAnchoCol = 4000
            .pCol = 2: .pAnchoCol = 500
            .pTitulo = "Servicios"
            .pForm = FORM_REGSERVICIO
            .pCaso = LABEL_SERVICIOS
            .Show
        End With
    End If
    If KeyCode = 13 Then
        txtServicio = Right("000000" & Trim(txtServicio), 6)
        lblServicio = DescripcionesdeCodigos("SERVICIO", Trim(txtServicio))
    End If
End Sub

Private Sub txtTarifa_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        With oConsulta
            .pCols = 5
            .pCol = 0: .pAnchoCol = 800
            .pCol = 1: .pAnchoCol = 1200
            .pTitulo = "Tarifas"
            .pForm = FORM_REGSERVICIO
            .pCaso = LABEL_TARIFAS
            .Show
        End With
    End If
    If KeyCode = 13 Then
        txtTarifa = Right("00" & Trim(txtTarifa), 2)
        lblTarifa = Space(2) & DescripcionesdeCodigos("TARIFA", Trim(txtTarifa))
    End If
End Sub

Private Sub moneda(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "0"
    cbo.AddItem "Nacional"
    cbo.List(1, 1) = "N"
    cbo.AddItem "Extranjera"
    cbo.List(2, 1) = "E"
    cbo.ListIndex = intTipoMoneda
End Sub

Private Sub Grabar()
    Dim SQL As String
    SQL = "Call Insert_ServTar( '" & MaxId & "', '" & txtServicio & "', '" & txtTarifa & "'," & _
          " '" & txtCodigo & "', '" & txtContrato & "','" & Format(meFecha, "yyyy/mm/dd") & "', " & _
          " " & CDbl(Trim(txtMonto)) & ",'" & cboMoneda.List(cboMoneda.ListIndex, 1) & "'," & CE(Trim(txtCuotas)) & ",'0'," & _
          " '0'," & CDbl(Trim(txtTotal)) & ",'" & IIf(chkPorcentaje.Value = True, "S", "N") & "' );"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
End Sub

Private Sub Actualizar()
    Dim SQL As String
    SQL = "Call Update_ServTar( '" & IdeSerTar & "', '" & txtServicio & "', '" & txtTarifa & "'," & _
          " '" & txtCodigo & "', '" & txtContrato & "','" & Format(meFecha, "yyyy/mm/dd") & "', " & _
          " " & CDbl(Trim(txtMonto)) & ",'" & cboMoneda.List(cboMoneda.ListIndex, 1) & "'," & CE(Trim(txtCuotas)) & ",'0'," & _
          " '0'," & CDbl(Trim(txtTotal)) & ",'" & IIf(chkPorcentaje.Value = True, "S", "N") & "' );"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
End Sub

Private Function ValidarData() As Boolean
    Dim I As Integer
    ValidarData = True
    If txtCodigo = Empty Then ValidarData = False: MsgBox "Es Necesario el Código del Solicitante del Servicio", vbInformation, gsNomSW: txtCodigo.SetFocus: Exit Function
    If txtServicio = Empty Then ValidarData = False: MsgBox "Ingrese Algun Tipo de Servicio", vbInformation, gsNomSW: txtServicio.SetFocus: Exit Function
    If txtTarifa = Empty Then ValidarData = False: MsgBox "Ingrese la Tarifa del Servicio", vbInformation, gsNomSW: txtTarifa.SetFocus: Exit Function
    If txtMonto = Empty Then ValidarData = False: MsgBox "Ingrese el Monto del Servicio", vbInformation, gsNomSW: txtMonto.SetFocus: Exit Function
    If txtCuotas = Empty Then ValidarData = False: MsgBox "Ingrese el Número de Cuotas", vbInformation, gsNomSW: txtCuotas.SetFocus: Exit Function
    If meFecha.Text = "__/__/____" Or Len(meFecha) < 10 Then ValidarData = False: MsgBox "Ingrese la Fecha de inicio del Servicio", vbInformation, gsNomSW: meFecha.SetFocus: Exit Function
    If cboMoneda.ListIndex = 0 Then ValidarData = False: MsgBox "Elegir el Tipo de Moneda", vbInformation, gsNomSW: cboMoneda.SetFocus: Exit Function
    With flxServicios
        For I = 1 To .Rows - 1
            If txtServicio = .TextMatrix(I, 4) And txtCodigo = .TextMatrix(I, 1) And _
               txtTarifa = .TextMatrix(I, 5) And lblModo <> "Modificar" Then
                ValidarData = False
                MsgBox "No puede Ingresar un servicio duplicado para el mismo cliente"
                txtServicio.SetFocus
                Exit For
                Exit Function
            End If
        Next
    End With
End Function
Private Sub LlenaCboCampos()
    With cboCampos
        .Clear
        .AddItem "CLIENTE"
        .List(0, 1) = "CodAux"
        
        .AddItem "CONTRATO"
        .List(1, 1) = "Contrato"
        
        .AddItem "SERVICIO"
        .List(2, 1) = "CodServ"
        
        '.AddItem "SERVICIO"
        '.List(3, 1) = "CodServ"
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub
Private Function MaxId() As String
    Dim SQL As String
    Dim Rs As MYSQL_RS
    SQL = "Select max(IDSerTar) from serv_tarif where codaux = '" & txtCodigo & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(Rs.Fields(IdSerTar)) Then
        MaxId = Trim(Right("0000" & Trim(CDbl(Rs.Fields(IdSerTar) + 1)), 4))
        IdeSerTar = MaxId
    Else
        MaxId = "0001"
        IdeSerTar = MaxId
    End If
    Set Rs = Nothing
End Function

Private Sub DesplazarxGrilla(fila As Integer)
    With flxServicios
        If .Rows > 1 Then
            txtCliente = BuscaenLista(CE(.TextMatrix(fila, 1)), "")
        Else
            Exit Sub
        End If
        txtCodigo = CE(.TextMatrix(fila, 1))
        txtContrato = CE(.TextMatrix(fila, 2))
        If Trim(.TextMatrix(fila, 3)) <> Empty Then
            meFecha = Format(Trim(.TextMatrix(fila, 3)), "dd/mm/yyyy")
        End If
        txtServicio = CE(.TextMatrix(fila, 4))
        lblServicio = Space(0) & DescripcionesdeCodigos("SERVICIO", Trim(CE(txtServicio)))
        txtTarifa = CE(.TextMatrix(fila, 5))
        lblTarifa = Space(0) & DescripcionesdeCodigos("TARIFA", Trim(CE(txtTarifa)))
        txtMonto = CE(.TextMatrix(fila, 8))
        txtCuotas = CE(.TextMatrix(fila, 7))
        txtTotal = CE(.TextMatrix(fila, 10))
        chkPorcentaje.Value = IIf(CE(.TextMatrix(fila, 12)) = "N", False, True)
        If .TextMatrix(fila, 6) = "N" Then
            cboMoneda.ListIndex = 1
        Else
            cboMoneda.ListIndex = 2
        End If
        IdeSerTar = CE(.TextMatrix(fila, 11))
    End With
End Sub

Private Sub ConfigGrilla()
    With flxServicios
        .Clear
        .Rows = 2
        .Cols = 13
        .ColWidth(0) = 550
        .TextMatrix(0, 0) = Space(1) + "Item"
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = Space(4) + "CodAux"
        .ColWidth(2) = 1400
        .TextMatrix(0, 2) = Space(8) + "Contrato"
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = Space(3) + "Fecha"
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = Space(1) + "Serv"
        .ColWidth(5) = 500
        .TextMatrix(0, 5) = Space(2) + "Tar"
        .ColWidth(6) = 400
        .TextMatrix(0, 6) = Space(0) + "Mon"
        .ColWidth(7) = 500
        .TextMatrix(0, 7) = Space(0) + "Cuot."
        .ColWidth(8) = 1300
        .TextMatrix(0, 8) = Space(8) + "Monto"
        .ColWidth(9) = 1300
        .TextMatrix(0, 9) = Space(5) + "Facturado"
        .ColWidth(10) = 1400
        .TextMatrix(0, 10) = Space(9) + "Total"
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = Space(8) + "IdCodServ"
        .ColWidth(12) = 0
        .TextMatrix(0, 12) = Space(8) + "PORC"
    End With
End Sub

Private Sub LlenarGrilla(Optional criterio As String, Optional filtro As String)
    Dim SQL As String
    Dim I As Integer
    Dim rsgrid As MYSQL_RS
    SQL = "Select * from serv_tarif WHERE " & criterio & " like '%" & filtro & "%' ORDER BY " & criterio
    Set rsgrid = oConexion.EjecutaSelectRS(SQL)
    I = 1
    If rsgrid.RecordCount > 0 Then ConfigGrilla
    flxServicios.Visible = False
    Do While Not rsgrid.EOF
        With flxServicios
            .TextMatrix(I, 0) = CE(rsgrid.Fields("IDSerTar"))
            .TextMatrix(I, 1) = CE(rsgrid.Fields("CodAux"))
            .TextMatrix(I, 2) = CE(rsgrid.Fields("Contrato"))
            .TextMatrix(I, 3) = CE(Format(rsgrid.Fields("Fec_Ini"), "dd/mm/yyyy"))
            .TextMatrix(I, 4) = CE(rsgrid.Fields("CodServ"))
            .TextMatrix(I, 5) = CE(rsgrid.Fields("CodTar"))
            .TextMatrix(I, 6) = CE(rsgrid.Fields("Moneda"))
            .TextMatrix(I, 7) = CE(rsgrid.Fields("Num_Cuota"))
            .TextMatrix(I, 8) = FormatNumber(CE(rsgrid.Fields("Monto")), 2)
            .TextMatrix(I, 9) = FormatNumber(CE(rsgrid.Fields("Facturado")), 2)
            .TextMatrix(I, 10) = FormatNumber(CE(rsgrid.Fields("Total")), 2)
            .TextMatrix(I, 11) = CE(rsgrid.Fields("IDSerTar"))
            .TextMatrix(I, 12) = CE(rsgrid.Fields("PORCENTAJE"))
            .Rows = .Rows + 1
            I = I + 1
            rsgrid.MoveNext
        End With
    Loop
    flxServicios.Rows = flxServicios.Rows - 1
    flxServicios.Visible = True
    Set rsgrid = Nothing
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flxServicios
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
