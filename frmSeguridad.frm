VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmSeguridad 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de usuarios"
   ClientHeight    =   8055
   ClientLeft      =   4065
   ClientTop       =   8715
   ClientWidth     =   8895
   Icon            =   "frmSeguridad.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8895
   Begin Proyecto1.chameleonButton cmdNuevo 
      Height          =   405
      Left            =   60
      TabIndex        =   20
      Top             =   7620
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
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
      MICON           =   "frmSeguridad.frx":57E2
      PICN            =   "frmSeguridad.frx":57FE
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
      Caption         =   "Opciones de consulta para los usuarios del sistema"
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
      Height          =   780
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   8775
      Begin VB.TextBox txtBuscar 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Top             =   300
         Width           =   5055
      End
      Begin VB.ComboBox cboBuscar 
         Height          =   315
         ItemData        =   "frmSeguridad.frx":5B68
         Left            =   105
         List            =   "frmSeguridad.frx":5B72
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2865
      End
      Begin Proyecto1.chameleonButton cmdBuscar 
         Height          =   405
         Left            =   8220
         TabIndex        =   19
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   714
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
         MICON           =   "frmSeguridad.frx":5B8E
         PICN            =   "frmSeguridad.frx":5BAA
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
   Begin TabDlg.SSTab TabSeguridad 
      Height          =   6225
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   10980
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      BackColor       =   14737632
      TabCaption(0)   =   "  &Compañias"
      TabPicture(0)   =   "frmSeguridad.frx":6144
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "   &Datos"
      TabPicture(1)   =   "frmSeguridad.frx":7AD6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   2640
         Left            =   -74910
         TabIndex        =   6
         Top             =   660
         Width           =   8460
         Begin VB.TextBox txtNombre 
            Height          =   300
            Left            =   2700
            MaxLength       =   50
            TabIndex        =   10
            Top             =   810
            Width           =   4710
         End
         Begin VB.TextBox txtClave 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2700
            MaxLength       =   12
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   1125
            Width           =   1395
         End
         Begin MSForms.ComboBox cboArea 
            Height          =   315
            Left            =   4770
            TabIndex        =   35
            Top             =   1110
            Width           =   2655
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "4683;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Area:"
            Height          =   195
            Index           =   1
            Left            =   4290
            TabIndex        =   34
            Top             =   1170
            Width           =   375
         End
         Begin MSForms.ComboBox txtUsuario 
            Height          =   315
            Left            =   2700
            TabIndex        =   33
            Top             =   150
            Width           =   1695
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "2990;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cboAnio 
            Height          =   285
            Left            =   3450
            TabIndex        =   32
            Top             =   2100
            Width           =   1215
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2143;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CboMesDeSistema 
            Height          =   315
            Left            =   5010
            TabIndex        =   31
            Top             =   2070
            Width           =   1785
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "3149;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox txtMesbloq 
            Height          =   315
            Left            =   1590
            TabIndex        =   30
            Top             =   2070
            Width           =   1515
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2672;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CboPerfil 
            Height          =   345
            Left            =   2700
            TabIndex        =   29
            Top             =   1470
            Width           =   4755
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "8387;609"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox CboEmpresa 
            Height          =   315
            Left            =   2700
            TabIndex        =   28
            Top             =   480
            Width           =   4725
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "8334;556"
            ColumnCount     =   2
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mes de Sistema"
            Height          =   195
            Index           =   5
            Left            =   5010
            TabIndex        =   18
            Top             =   1860
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Año De Sistema"
            Height          =   195
            Index           =   4
            Left            =   3450
            TabIndex        =   17
            Top             =   1860
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mes Bloqueo"
            Height          =   195
            Index           =   3
            Left            =   1605
            TabIndex        =   16
            Top             =   1860
            Width           =   930
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Index           =   1
            Left            =   1605
            TabIndex        =   8
            Top             =   525
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Perfil"
            Height          =   195
            Index           =   0
            Left            =   1605
            TabIndex        =   13
            Top             =   1545
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Contraseña"
            Height          =   165
            Index           =   0
            Left            =   1605
            TabIndex        =   11
            Top             =   1200
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   1605
            TabIndex        =   9
            Top             =   855
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Id. del Usuario"
            Height          =   195
            Left            =   1605
            TabIndex        =   7
            Top             =   225
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5610
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   8535
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHCompañias 
            Height          =   5355
            Left            =   90
            TabIndex        =   2
            Top             =   180
            Width           =   8325
            _ExtentX        =   14684
            _ExtentY        =   9446
            _Version        =   393216
            BackColorBkg    =   12632256
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin Proyecto1.chameleonButton cmdEditar 
      Height          =   405
      Left            =   1140
      TabIndex        =   21
      Top             =   7620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Modificar"
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
      MICON           =   "frmSeguridad.frx":C8D8
      PICN            =   "frmSeguridad.frx":C8F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdBorrar 
      Height          =   405
      Left            =   2400
      TabIndex        =   22
      Top             =   7620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Eliminar"
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
      MICON           =   "frmSeguridad.frx":CD22
      PICN            =   "frmSeguridad.frx":CD3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdCancelar 
      Height          =   405
      Left            =   4110
      TabIndex        =   23
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
      MICON           =   "frmSeguridad.frx":D180
      PICN            =   "frmSeguridad.frx":D19C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdGrabar 
      Height          =   405
      Left            =   4560
      TabIndex        =   24
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
      MICON           =   "frmSeguridad.frx":D6DE
      PICN            =   "frmSeguridad.frx":D6FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdSalir 
      Height          =   405
      Left            =   8340
      TabIndex        =   25
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
      MICON           =   "frmSeguridad.frx":DB3C
      PICN            =   "frmSeguridad.frx":DB58
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton BtnReporte 
      Height          =   405
      Left            =   7920
      TabIndex        =   26
      ToolTipText     =   "Mostrar Reporte"
      Top             =   7620
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
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
      MICON           =   "frmSeguridad.frx":DF1E
      PICN            =   "frmSeguridad.frx":DF3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdMantNiveles 
      Height          =   405
      Left            =   5310
      TabIndex        =   27
      Top             =   7620
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Mant. de Perfiles"
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
      MICON           =   "frmSeguridad.frx":E47C
      PICN            =   "frmSeguridad.frx":E498
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   300
      Left            =   45
      TabIndex        =   14
      Top             =   7155
      Width           =   8700
   End
End
Attribute VB_Name = "frmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public CantRegistros As Integer
Public MSHabilitado  As Boolean
Private m_perfil_id As String

Private Function ImprimirPerfil(perfil As String) As Integer
Dim i As Integer
    For i = 0 To CboPerfil.ListCount - 1
        If Trim(CboPerfil.List(i, 1)) = perfil Then
            ImprimirPerfil = i
            Exit Function
        End If
    Next
    ImprimirPerfil = 0
End Function

Sub LlenarMSHCompañias()
    Dim RsCias As MYSQL_RS
    Set RsCias = New MYSQL_RS
    Dim SqlUser As String
    Dim SqlCias  As String
    
    SqlCias = "SELECT b.usuario_id,a.CODCIA, c.nombre , c.clave, c.perfil_id," & _
              " b.ANOmes_bloq, b.anio_actual, b.mes_actual, c.area " & _
              " FROM (7cia_user AS b LEFT JOIN 6cncias AS a " & _
              " ON a.codcia=b.codcia) left join 3cnuser as c  on b.usuario_id = c.usuario_id where c.estado='1'"
    Set RsCias = oConexion.EjecutaSelectRS(SqlCias)
    If RsCias.EOF And RsCias.BOF Then
        BloqueoDeBotones
    Else
        ConfigMSHCompañias
        Dim i As Integer
        With MSHCompañias
            Do While Not RsCias.EOF
                .TextMatrix(.Rows - 1, 1) = CE(RsCias.Fields(0))
                .TextMatrix(.Rows - 1, 2) = CE(RsCias.Fields(1))
                .TextMatrix(.Rows - 1, 3) = CE(RsCias.Fields(2))
                .TextMatrix(.Rows - 1, 4) = CE(RsCias.Fields(3))
                .TextMatrix(.Rows - 1, 5) = CE(RsCias.Fields(4))
                .TextMatrix(.Rows - 1, 6) = CE(RsCias.Fields(5))
                .TextMatrix(.Rows - 1, 7) = CE(RsCias.Fields(6))
                .TextMatrix(.Rows - 1, 8) = CE(RsCias.Fields(7))
                .TextMatrix(.Rows - 1, 9) = CE(RsCias.Fields(8))
                RsCias.MoveNext
                .Rows = .Rows + 1
            Loop
            .Rows = .Rows - 1
        End With
    End If
    Set RsCias = Nothing
End Sub

Sub BloqueoDeBotones()
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = False
    cmdBorrar.Enabled = False
    cmdCancelar.Enabled = False
    cmdGrabar.Enabled = False
    cmdSalir.Enabled = False
End Sub

Public Property Let perfil_id(valor As String)
    m_perfil_id = Trim(valor)
End Property

Sub ModoNormal()
    txtUsuario.BackColor = ColorDeshabilitado
    txtNombre.BackColor = ColorDeshabilitado
    txtClave.BackColor = ColorDeshabilitado
    cboEmpresa.BackColor = ColorDeshabilitado
    CboPerfil.BackColor = ColorDeshabilitado
    cboArea.BackColor = ColorDeshabilitado
    txtMesbloq.BackColor = ColorDeshabilitado
    CboMesDeSistema.BackColor = ColorDeshabilitado
    cboAnio.BackColor = ColorDeshabilitado
    MSHCompañias.BackColor = ColorHabilitado
    txtUsuario.Locked = True
    txtNombre.Locked = True
    txtClave.Locked = True
    cboEmpresa.Locked = True
    cboArea.Locked = True
    CboPerfil.Locked = True
    txtMesbloq.Locked = True
    CboMesDeSistema.Locked = True
    cboAnio.Locked = True
    MSHabilitado = True
End Sub

Sub ModoEdicion()
    txtUsuario.BackColor = ColorHabilitado
    txtNombre.BackColor = ColorHabilitado
    txtClave.BackColor = ColorHabilitado
    cboEmpresa.BackColor = ColorHabilitado
    CboPerfil.BackColor = ColorHabilitado
    cboArea.BackColor = ColorHabilitado
    txtMesbloq.BackColor = ColorHabilitado
    CboMesDeSistema.BackColor = ColorHabilitado
    cboAnio.BackColor = ColorHabilitado
    MSHCompañias.BackColor = ColorDeshabilitado
    txtUsuario.Locked = False
    txtNombre.Locked = False
    txtClave.Locked = False
    cboEmpresa.Locked = False
    cboArea.Locked = False
    CboPerfil.Locked = False
    txtMesbloq.Locked = False
    CboMesDeSistema.Locked = False
    cboAnio.Locked = False
    MSHabilitado = False
End Sub

Sub BotonNormal()
    cmdNuevo.Enabled = True
    cmdEditar.Enabled = True
    cmdBorrar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    cmdSalir.Enabled = True
End Sub

Sub BotonEdicion()
    cmdNuevo.Enabled = False
    cmdEditar.Enabled = False
    cmdBorrar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
End Sub

Sub LimpiarDatos()
    txtUsuario = Empty
    txtNombre = Empty
    txtClave = Empty
    m_perfil_id = Empty
End Sub

Sub ConfigMSHCompañias()
    With MSHCompañias
        .Cols = 10
        .Rows = 2
        .Clear
        .ColWidth(0) = 0
        .ColWidth(1) = 1200
        .ColWidth(2) = 800
        .ColWidth(3) = 2500
        .ColWidth(4) = 1200
        .ColWidth(5) = 1000
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        .ColWidth(8) = 800
        .ColWidth(9) = 800
        .TextMatrix(0, 1) = "Id.Usuario"
        .TextMatrix(0, 2) = "Cod.Cia"
        .TextMatrix(0, 3) = "Nombre del usuario"
        .TextMatrix(0, 4) = "Clave"
        .TextMatrix(0, 5) = "Perfil ID"
        .TextMatrix(0, 6) = "Mes Bloq"
        .TextMatrix(0, 7) = "Año Sist"
        .TextMatrix(0, 8) = "Mes Sist"
        .TextMatrix(0, 9) = "Area"
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub btnReporte_Click()
    Set oReporte = New clsReporte
    oReporte.Reporte = "Rep_Usuarios.rpt"
    oReporte.sp_Rep_Usuarios
End Sub

Private Sub cboAnio_Change()
    carga_mesistema Trim(cboAnio.Text)
End Sub

Private Sub cboBuscar_Click()
    If Trim(cboBuscar.Text) <> Empty Then
        txtBuscar.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
    End If
End Sub

Private Sub cmdBorrar_Click()
  If txtUsuario <> Empty Then
        If MsgBox("Está Seguro De Eliminar El Usuario Con El ID. User ==> " + CStr(txtUsuario) + " (S/N)", vbExclamation + vbYesNo, Caption) = vbYes Then
            BorrarRegistro
            MSHCompañias.Clear
            LlenarMSHCompañias
            LimpiarDatos
            ModoNormal
            BotonNormal
            LblMensaje = Empty
        End If
    Else
        MsgBox "Seleccione El Registro A Eliminar", vbInformation, Caption
        MSHCompañias.SetFocus
    End If
End Sub

Sub BorrarRegistro()
    Dim SQL As String
    SQL = " Call Delete_CiaUser ('" & cboEmpresa.List(cboEmpresa.ListIndex, 1) & "','" & Trim(txtUsuario.Text) & "');"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
End Sub

Private Sub cmdEditar_Click()
    ModoEdicion
    BotonEdicion
    TabSeguridad.Tab = 1
    LblMensaje = "Modificar"
End Sub

Private Sub cmdBuscar_Click()
    Call txtBuscar_KeyPress(13)
End Sub

Private Function BuscarUsuario(criterio As String) As Integer
    Dim i As Integer
    With MSHCompañias
        Select Case cboBuscar.ListIndex
            Case 0 ' -- Nombre
                For i = 1 To .Rows - 1
                    If UCase(Trim(.TextMatrix(i, 3))) Like "*" & UCase(criterio) & "*" Then
                        BuscarUsuario = i
                        Exit Function
                    End If
                Next
            Case 1 ' -- Identificación
                For i = 1 To .Rows - 1
                    If UCase(Trim(.TextMatrix(i, 1))) Like "*" & UCase(criterio) & "*" Then
                        BuscarUsuario = i
                        Exit Function
                    End If
                Next
        End Select
    End With
    BuscarUsuario = 0
End Function

Private Sub cmdCancelar_Click()
    LimpiarDatos
    ModoNormal
    BotonNormal
    LblMensaje = Empty
    TabSeguridad.Tab = 0
    Call MSHCompañias_Click
End Sub

Private Sub cmdGrabar_Click()
    If VerificarDatos Then
        mdiInicio.MousePointer = vbHourglass
        GrabarDatos
        ModoNormal
        BotonNormal
        LblMensaje = Empty
        MSHCompañias.Clear
        LlenarMSHCompañias
        TabSeguridad.Tab = 0
        mdiInicio.MousePointer = vbNormal
    End If
End Sub

Sub GrabarDatos()
    Dim ClaveDeBusqueda As String, TextoCodificado As String
    Dim i As Integer
    Dim Rs As MYSQL_RS
    Dim SQL As String
    
    For i = 0 To 255
        ClaveDeBusqueda = ClaveDeBusqueda + Chr$(i)
    Next
    Select Case LblMensaje.Caption
        Case "Nuevo"
            SQL = "SELECT usuario_id FROM usuarios WHERE usuario_id='" & Trim(txtUsuario) & "'"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            If Rs.RecordCount = 0 Then
                TextoCodificado = ChrTran(txtClave, ClaveDeBusqueda, ClaveAleatoria)
                SQL = "Call Insert_Usuarios ('" & Trim(txtUsuario) & "','" & Trim(Me.txtNombre) & "','" & _
                       Trim(TextoCodificado) & "','" & Trim(CboPerfil.List(CboPerfil.ListIndex, 1)) & "','" & Trim(cboArea.List(cboArea.ListIndex, 1)) & "')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
                SQL = "Call Insert_CiaUser ('" & Trim(cboEmpresa.List(cboEmpresa.ListIndex, 1)) & "','" & _
                                                 Trim(txtUsuario.Text) & "','" & _
                                                 Right(Trim(CboMesDeSistema.List(CboMesDeSistema.ListIndex, 1)), 2) & "','" & _
                                                 Trim(cboAnio.List(cboAnio.ListIndex, 0)) & "','" & _
                                                 Trim(txtMesbloq.List(txtMesbloq.ListIndex, 1)) & "','','','','','',0)"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
                GrabaMenudeUsuario Trim(txtUsuario.Text), CboPerfil.List(CboPerfil.ListIndex, 1)
                GrabaEstadosUsuario Trim(txtUsuario.Text)
                GrabaDocsUsuario Trim(txtUsuario.Text)
                Rs.CloseRecordset
                Set Rs = Nothing
            Else
                SQL = "SELECT codcia,usuario_id FROM 7cia_user WHERE usuario_id='" & Trim(txtUsuario) & "' and codcia='" & Trim(cboEmpresa.List(cboEmpresa.ListIndex, 1)) & "'"
                Set Rs = oConexion.EjecutaSelectRS(SQL)
                If Rs.RecordCount = 0 Then
                    SQL = "Call Insert_CiaUser ('" & Trim(cboEmpresa.List(cboEmpresa.ListIndex, 1)) & "','" & _
                                                 Trim(txtUsuario.Text) & "','" & _
                                                 Right(Trim(CboMesDeSistema.List(CboMesDeSistema.ListIndex, 1)), 2) & "','" & _
                                                 Trim(cboAnio.List(cboAnio.ListIndex, 0)) & "','" & _
                                                 Trim(txtMesbloq.List(txtMesbloq.ListIndex, 1)) & "','','','','','',0)"
                Else
                    MsgBox "El usuario ya existe... debe asociarlo a otra empresa ", vbInformation, "Error de datos"
                End If
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
                GrabaMenudeUsuario Trim(txtUsuario.Text), CboPerfil.List(CboPerfil.ListIndex, 1)
                GrabaEstadosUsuario Trim(txtUsuario.Text)
                GrabaDocsUsuario Trim(txtUsuario.Text)
            End If
        Case "Modificar"
            TextoCodificado = ChrTran(txtClave, ClaveDeBusqueda, ClaveAleatoria)
            SQL = "CALL Update_Usuarios ('" & Trim(txtUsuario.Text) & "','" & MSHCompañias.TextMatrix(MSHCompañias.row, 1) & "','" & txtNombre & "','" & TextoCodificado & "','" & CboPerfil.List(CboPerfil.ListIndex, 1) & "','" & Trim(cboArea.List(cboArea.ListIndex, 1)) & "')"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
            SQL = "Call Update_CiaUser ('" & Trim(cboEmpresa.List(cboEmpresa.ListIndex, 1)) & "','" & _
                                                 Trim(txtUsuario.Text) & "','" & _
                                                 Right(Trim(CboMesDeSistema.List(CboMesDeSistema.ListIndex, 1)), 2) & "','" & _
                                                 Trim(cboAnio.List(cboAnio.ListIndex, 0)) & "','" & _
                                                 Trim(txtMesbloq.List(txtMesbloq.ListIndex, 1)) & "','','','','','',0)"
              
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
    End Select
    Set Rs = Nothing
    Exit Sub
End Sub

Function VerificarDatos() As Boolean
    If Trim(txtUsuario) = Empty Then
        MsgBox "Ingrese identificación del usuario", vbExclamation, "Error de datos"
        txtUsuario.SetFocus
        VerificarDatos = False
        Exit Function
    End If
    If cboEmpresa.Text = Empty Then
      MsgBox "Escoja una empresa", vbExclamation, "Error de datos"
      cboEmpresa.SetFocus
      VerificarDatos = False
      Exit Function
    End If
    If Trim(txtNombre) = Empty Then
        MsgBox "Ingrese nombre del usuario", vbExclamation, "Error de datos"
        txtNombre.SetFocus
        VerificarDatos = False
        Exit Function
    End If
    If Trim(txtClave) = Empty Then
        MsgBox "Ingrese contraseña del usuario", vbExclamation, "Error de datos"
        txtClave.SetFocus
        VerificarDatos = False
        Exit Function
    End If
    If CboPerfil.Text = Empty Then
        MsgBox "Falta asignar los perfiles para el usuario", vbExclamation, "Error de datos"
        VerificarDatos = False
        CboPerfil.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
        Exit Function
    End If
    VerificarDatos = True
End Function

Private Sub CmdMantNiveles_Click()
    ConfigurarFormulario FORM_MAESTRO_PERFILES
End Sub

Private Sub cmdNuevo_Click()
    LimpiarDatos
    ModoEdicion
    BotonEdicion
    LblMensaje = "Nuevo"
    TabSeguridad.Tab = 1
    txtUsuario.SetFocus
    MSHabilitado = False
    cboAnio = strAnoSistema
    CboMesDeSistema.ListIndex = 0
    txtMesbloq.ListIndex = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub carga_anomes_bloq()
    Dim SQL As String, i As Integer
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    SQL = "select distinct anomes from documentos order by anomes desc"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    txtMesbloq.Clear
    txtMesbloq.AddItem "Ninguno"
    txtMesbloq.List(0, 1) = "000000"
    i = 1
    If Rs.RecordCount = 0 Then
        txtMesbloq.AddItem Year(Date)
    Else
        Do While Not (Rs.EOF)
            txtMesbloq.AddItem NombreMes(Right(Rs.Fields("anomes"), 2), True) & Space(2) & Left(Rs.Fields("anomes"), 4)
            txtMesbloq.List(i, 1) = Rs.Fields("anomes")
            i = i + 1
            Rs.MoveNext
        Loop
    End If
    Rs.CloseRecordset
    SQL = "select distinct left(anomes,4) as anio from documentos order by left(anomes,4) desc"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cboAnio.Clear
    If Rs.RecordCount = 0 Then
        cboAnio.AddItem Year(Date)
    Else
        Do While Not (Rs.EOF)
            cboAnio.AddItem Rs.Fields("anio")
            Rs.MoveNext
        Loop
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Sub

Private Sub carga_mesistema(Anio As String)
    Dim SQL As String, i As Integer
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    SQL = "select distinct anomes from documentos where left(anomes,4)='" & Anio & "' order by anomes desc"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    CboMesDeSistema.Clear
    If Rs.RecordCount = 0 Then
        CboMesDeSistema.AddItem NombreMes(strMesSistema, False)
        CboMesDeSistema.List(i, 1) = Year(Date) & strMesSistema
    Else
        Do While Not (Rs.EOF)
            CboMesDeSistema.AddItem NombreMes(Right(Trim(Rs.Fields("anomes")), 2), False)
            CboMesDeSistema.List(i, 1) = Rs.Fields("anomes")
            i = i + 1
            Rs.MoveNext
        Loop
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
      Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    ModoNormal
    BotonNormal
    LimpiarDatos
    carga_anomes_bloq
    LlenarMSHCompañias
    LlenarCbos
    LblMensaje.Caption = Empty
    TabSeguridad.Tab = 0
    Call MSHCompañias_Click
End Sub

Sub LlenarCbos()
    CboPerfil.Clear
    cboEmpresa.Clear
    txtUsuario.Clear
    cboArea.Clear
    Dim i As Integer
    Dim RsPerfil As MYSQL_RS
    Set RsPerfil = New MYSQL_RS
    Dim RsEmpresas As MYSQL_RS
    Set RsEmpresas = New MYSQL_RS
    Dim RsUser As MYSQL_RS
    Set RsUser = New MYSQL_RS
    Dim RsArea As MYSQL_RS
    Set RsArea = New MYSQL_RS
    Dim SQL As String
    i = 0
    Set RsPerfil = oConexion.EjecutaSelect("perfiles")
    If RsPerfil.BOF Or RsPerfil.EOF Then
        MsgBox "No se encontraron perfiles registrados", vbInformation, Caption
    Else
        Do While Not RsPerfil.EOF
            CboPerfil.AddItem Trim(RsPerfil.Fields("descripcion"))
            CboPerfil.List(i, 1) = Trim(RsPerfil.Fields("perfil_id"))
            i = i + 1
            RsPerfil.MoveNext
        Loop
    End If
    RsPerfil.CloseRecordset
    Set RsPerfil = Nothing
    i = 0
    Set RsEmpresas = oConexion.EjecutaSelect("empresas")
    If RsEmpresas.BOF = False Or RsEmpresas.EOF = False Then
        Do While Not RsEmpresas.EOF
            cboEmpresa.AddItem RsEmpresas.Fields("descrip")
            cboEmpresa.List(i, 1) = RsEmpresas.Fields("codcia")
            i = i + 1
            RsEmpresas.MoveNext
        Loop
    Else
        MsgBox "No hay empresas registradas", vbInformation, Caption
    End If
    RsEmpresas.CloseRecordset
    Set RsEmpresas = Nothing
  
    i = 0
    Set RsArea = oConexion.EjecutaSelect("areas")
    If RsArea.BOF = False Or RsArea.EOF = False Then
        Do While Not RsArea.EOF
            cboArea.AddItem RsArea.Fields("descrip")
            cboArea.List(i, 1) = RsArea.Fields("idarea")
            i = i + 1
            RsArea.MoveNext
        Loop
    Else
        MsgBox "No hay empresas registradas", vbInformation, Caption
    End If
    RsArea.CloseRecordset
    Set RsArea = Nothing
  
    i = 0
    Set RsUser = oConexion.EjecutaSelect("usuarios")
    If RsUser.BOF = False Or RsUser.EOF = False Then
        Do While Not RsUser.EOF
            txtUsuario.AddItem RsUser.Fields("usuario_id")
            txtUsuario.List(i, 1) = RsUser.Fields("nombre")
            txtUsuario.List(i, 2) = RsUser.Fields("clave")
            txtUsuario.List(i, 3) = RsUser.Fields("perfil_id")
            i = i + 1
            RsUser.MoveNext
        Loop
    Else
        MsgBox "No hay usuarios registrados", vbInformation, Caption
    End If
    RsUser.CloseRecordset
    Set RsUser = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiInicio.Enabled = True
End Sub

Private Sub MSHCompañias_Click()
    NavegarPorMSHCompañias
End Sub

Sub NavegarPorMSHCompañias()
    Dim ClaveDeBusqueda As String, TextoCodificado As String, TextoOriginal As String
    Dim i As Integer
    If MSHabilitado = True Then
        With MSHCompañias
            txtUsuario.Text = .TextMatrix(.Rowsel, 1)
            cboEmpresa.ListIndex = ImprimirEmpresa(.TextMatrix(.Rowsel, 2))
            txtNombre.Text = .TextMatrix(.Rowsel, 3)
            For i = 0 To 255
                ClaveDeBusqueda = ClaveDeBusqueda + Chr$(i)
            Next
            TextoOriginal = ChrTran(.TextMatrix(.Rowsel, 4), ClaveAleatoria, ClaveDeBusqueda)
            txtClave.Text = TextoOriginal
            CboPerfil.ListIndex = ImprimirPerfil(.TextMatrix(.Rowsel, 5))
            txtMesbloq.ListIndex = ImprimirMesBloqueo(.TextMatrix(.Rowsel, 6))
            cboAnio.ListIndex = ImprimirAnio(.TextMatrix(.Rowsel, 7))
            CboMesDeSistema.ListIndex = ImprimirMes(.TextMatrix(.Rowsel, 8))
            cboArea.ListIndex = ImprimirArea(.TextMatrix(.Rowsel, 9))
        End With
    End If
End Sub

Private Function ImprimirMes(Mes As String) As Integer
    Dim i As Integer
    For i = 0 To CboMesDeSistema.ListCount - 1
        If Right(CboMesDeSistema.List(i, 1), 2) = Mes Then
            ImprimirMes = i
            Exit Function
        End If
    Next
    ImprimirMes = 0
End Function

Private Function ImprimirMesBloqueo(anomesblo As String) As Integer
    Dim i As Integer
    For i = 0 To txtMesbloq.ListCount - 1
        If Trim(anomesblo) = Trim(txtMesbloq.List(i, 1)) Then
          ImprimirMesBloqueo = i
          Exit Function
        End If
    Next
    ImprimirMesBloqueo = 0
End Function

Private Function ImprimirAnio(Anio As String) As Integer
    Dim i As Integer
    For i = 0 To cboAnio.ListCount - 1
        If Trim(Anio) = Trim(cboAnio.List(i, 1)) Then
          ImprimirAnio = i
          Exit Function
        End If
    Next
    ImprimirAnio = 0
End Function

Private Function ImprimirEmpresa(XCodcia As String) As Integer
    Dim i As Integer
    For i = 0 To cboEmpresa.ListCount - 1
        If Trim(cboEmpresa.List(i, 1)) = Trim(XCodcia) Then
            ImprimirEmpresa = i
            Exit Function
        End If
    Next
    ImprimirEmpresa = 0
End Function

Private Function ImprimirArea(XArea As String) As Integer
    Dim i As Integer
    For i = 0 To cboArea.ListCount - 1
        If Trim(cboArea.List(i, 1)) = Trim(XArea) Then
            ImprimirArea = i
            Exit Function
        End If
    Next
    ImprimirArea = 0
End Function

Private Sub MSHCompañias_DblClick()
    Call cmdEditar_Click
End Sub

Private Sub MSHCompañias_KeyDown(KeyCode As Integer, Shift As Integer)
    NavegarPorMSHCompañias
End Sub

Private Sub TabSeguridad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
    End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Screen.MousePointer = vbHourglass
        If cboBuscar.Text = Empty Then
            MsgBox "Seleccione la opción de búsqueda", vbExclamation, "Consulta"
            Screen.MousePointer = vbNormal
            Exit Sub
        Else
            Dim fila As Integer
            fila = BuscarUsuario(txtBuscar)
            If fila > 0 Then
                TabSeguridad.Tab = 0
                MSHCompañias.row = fila
                MSHCompañias.Col = 0
                MSHCompañias.SetFocus
                Call keybd_event(vbKeyHome, 0, 0, 0)
            Else
                MsgBox "Usuario no existe", vbInformation, Caption
                txtBuscar.SetFocus
                Call keybd_event(vbKeyHome, 0, 0, 0)
            End If
        End If
        Screen.MousePointer = vbNormal
    Else
        KeyAscii = OnlyChar(KeyAscii)
    End If
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdGrabar_Click
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CboPerfil.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
    Else
        KeyAscii = OnlyChar(KeyAscii)
    End If
End Sub

Public Sub GrabaMenudeUsuario(usuario As String, perfil As String)
    Dim SQL As String
    Dim i As Integer
    SQL = "delete from 5USUARIO_MENU where USUARIO_ID='" & usuario & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    For i = 0 To 4
        SQL = "Insert into 5USUARIO_MENU select '" & usuario & "' as usuario_id ,modulo,item,visible from 4perfil_menu where perfil_id='" & perfil & "' and modulo=" & i
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
    Next
End Sub

Public Sub GrabaEstadosUsuario(usuario As String)
    Dim SQL As String
    Dim i As Integer
    Dim rsestados As MYSQL_RS
    SQL = "Select * from doc_estados "
    Set rsestados = oConexion.EjecutaSelectRS(SQL)
    SQL = "delete from ESTADO_USU where USUARIO_ID='" & usuario & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    Do While Not rsestados.EOF
        SQL = "Insert into ESTADO_USU (Cod_estado,Usuario_id,permiso) " & _
              " values('" & rsestados.Fields("COD_ESTADO") & "','" & usuario & "',0)"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        rsestados.MoveNext
    Loop
    Set rsestados = Nothing
End Sub

Public Sub GrabaDocsUsuario(usuario As String)
    Dim SQL As String
    Dim i As Integer
    Dim RQ As MYSQL_RS
    SQL = "Select * from cndocum"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    SQL = "delete from docsusuario where USUARIO_ID='" & usuario & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    Do While Not RQ.EOF
        SQL = "Insert into docsusuario (coddoc,Usuario,permiso) " & _
              " values('" & RQ.Fields("coddoc") & "','" & usuario & "',0)"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub

Private Sub txtUsuario_Change()
    Dim i As Integer
    If txtUsuario.ListIndex <> -1 Then
        txtNombre.Text = txtUsuario.List(txtUsuario.ListIndex, 1)
        txtClave.Text = txtUsuario.List(txtUsuario.ListIndex, 2)
        For i = 0 To CboPerfil.ListCount - 1
            If CboPerfil.List(i, 1) = txtUsuario.List(txtUsuario.ListIndex, 3) Then
                CboPerfil.ListIndex = i
                Exit For
            End If
        Next
    End If
End Sub
