VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmLoginUSer 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Integrado Administrativo - Inicio de Sesión"
   ClientHeight    =   4050
   ClientLeft      =   6015
   ClientTop       =   5325
   ClientWidth     =   6060
   Icon            =   "frmLoginUSer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "frmLoginUSer.frx":1CCA
   ScaleHeight     =   4050
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   -30
      TabIndex        =   13
      Top             =   -120
      Width           =   6315
      Begin VB.Image Image2 
         Height          =   1035
         Left            =   60
         Picture         =   "frmLoginUSer.frx":C183
         Stretch         =   -1  'True
         Top             =   150
         Width           =   6000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   3540
      Left            =   -60
      TabIndex        =   8
      Top             =   900
      Width           =   6165
      Begin VB.TextBox txtConfirmaPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1530
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Proyecto1.chameleonButton cmdAceptar 
         Height          =   345
         Left            =   1830
         TabIndex        =   5
         Top             =   2400
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Aceptar"
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLoginUSer.frx":140DD
         PICN            =   "frmLoginUSer.frx":140F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture1 
         Height          =   2625
         Left            =   90
         ScaleHeight     =   2565
         ScaleWidth      =   1455
         TabIndex        =   11
         Top             =   3375
         Visible         =   0   'False
         Width           =   1515
         Begin VB.Image Image1 
            Height          =   3255
            Left            =   30
            Picture         =   "frmLoginUSer.frx":14B0B
            Stretch         =   -1  'True
            Top             =   90
            Visible         =   0   'False
            Width           =   1650
         End
      End
      Begin VB.TextBox txtClave 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1860
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1170
         Width           =   1815
      End
      Begin Proyecto1.chameleonButton cmdExit 
         Height          =   345
         Left            =   3210
         TabIndex        =   6
         Top             =   2400
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Salir"
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLoginUSer.frx":17166
         PICN            =   "frmLoginUSer.frx":17182
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdCPass 
         Height          =   345
         Left            =   4170
         TabIndex        =   7
         Top             =   1140
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Cambiar"
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLoginUSer.frx":17B94
         PICN            =   "frmLoginUSer.frx":17BB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblServidor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "172.26.35.1"
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
         Left            =   4995
         TabIndex        =   16
         Top             =   2835
         Width           =   1035
      End
      Begin MSForms.ComboBox cmbAniosAnt 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   1920
         Width           =   2625
         VariousPropertyBits=   746588185
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4630;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Histórico"
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
         Left            =   810
         TabIndex        =   15
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label lblConfirmar 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmar"
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
         Left            =   810
         TabIndex        =   14
         Top             =   1590
         Visible         =   0   'False
         Width           =   810
      End
      Begin MSForms.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   390
         Width           =   3645
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6429;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   810
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin MSForms.ComboBox cboUsuario 
         Height          =   285
         Left            =   1860
         TabIndex        =   1
         Top             =   810
         Width           =   1815
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3201;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   810
         TabIndex        =   10
         Top             =   870
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
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
         Left            =   810
         TabIndex        =   9
         Top             =   1230
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmLoginUSer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PassAnterior As String

Private Sub cboEmpresa_Change()
    If cboEmpresa.ListIndex >= 0 Then
        With cboEmpresa
            UserEmpresas cboUsuario, .List(.ListIndex, 1)     'Actualiza Usuarios según la Empresa
            strCodigoEmpresa = .List(.ListIndex, 1)
            strNombreEmpresa = .List(.ListIndex, 0)
            strRucEmpresa = .List(.ListIndex, 5)
        End With
    End If
End Sub

Private Sub cboUsuario_Change()
    If cboUsuario.ListIndex >= 0 Then
        With cboUsuario
            strUsuarioId = .List(.ListIndex, 0)
            strClaveUsuario = .List(.ListIndex, 1)
            strNombreUsuario = .List(.ListIndex, 2)
            strPerfilUsuario = .List(.ListIndex, 3)
            strAnoSistema = .List(.ListIndex, 4)
            strMesSistema = .List(.ListIndex, 5)
            strAreaUsuario = .List(.ListIndex, 7)
            strUsuarioProv = .List(.ListIndex, 8)
            strCondUsuario = ""
        End With
    End If
End Sub

Private Sub cmbAniosAnt_Click()
    strODBCEmp = cmbAniosAnt.List(cmbAniosAnt.ListIndex, 1)
    
    If cmbAniosAnt.ListIndex > 0 Then
        strAnioConex = cmbAniosAnt.List(cmbAniosAnt.ListIndex, 1)
        strAnoSistema = Right(cmbAniosAnt.List(cmbAniosAnt.ListIndex, 1), 4)
    End If
    
    If oConexionSQL.ConectarEmpresa = True Then
        Empresas cboEmpresa
        cboEmpresa_Change
    End If
        
    
End Sub

Private Sub cmdAceptar_Click()
    Dim ClaveDeBusqueda As String, TextoCodificado As String
    Dim I As Integer
    For I = 0 To 255
         ClaveDeBusqueda = ClaveDeBusqueda + Chr$(I)
    Next
    TextoCodificado = ChrTran(txtClave, ClaveDeBusqueda, ClaveAleatoria)
    If PassAnterior <> Empty Then
        If Len(txtClave) < 4 Then
            MsgBox "Debe ingresar como mínimo 4 caracteres", vbInformation, "NOVPeru"
        Else
            SQL = "Update 3cnuser set clave='" & TextoCodificado & "' where usuario_id='" & strUsuarioId & "'"
            oConexionMYSQL.Execute SQL
            ValidaUser PassAnterior
            
        End If
    Else
        ValidaUser txtClave.Text
    End If
End Sub

Private Sub cmdCPass_Click()
    Dim ClaveDeBusqueda As String, TextoCodificado As String, TextoOriginal As String
    Dim I As Integer
    For I = 0 To 255
        ClaveDeBusqueda = ClaveDeBusqueda + Chr$(I)
    Next
    TextoOriginal = ChrTran(strClaveUsuario, ClaveAleatoria, ClaveDeBusqueda)
    PassAnterior = TextoOriginal
    txtClave = Empty
    lblConfirmar.Visible = True
    txtConfirmaPass.Visible = True
    cmdCPass.Enabled = False
    txtClave.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub Form_Activate()
    txtClave.SetFocus
End Sub

Private Sub Form_Load()
    Empresas cboEmpresa
    AniosAnteriores
    
    
    lblServidor.Caption = "SERVER: " & gsServidor
    

    
End Sub

Private Sub AniosAnteriores()
    Dim SQL As String, I As Integer
    Dim rsant As MYSQL_RS
   
    SQL = "Select BD from cnanio where bd<>'BRANDTMYSQLHFM' ORDER BY BD"
    Set rsant = oConexion.EjecutaSelectRS(SQL)
    cmbAniosAnt.Clear
    I = 1
    cmbAniosAnt.AddItem "Selecionar...", 0
    cmbAniosAnt.List(0, 1) = "BRANDTMYSQLHFM"
    Do While Not (rsant.EOF)
        cmbAniosAnt.AddItem Right(rsant.Fields("bd"), 4), I
        cmbAniosAnt.List(I, 1) = rsant.Fields("bd")       'Almacena la clave del Usuario
        I = I + 1
        rsant.MoveNext
    Loop
    If I >= 2 Then
        cmbAniosAnt.Enabled = True
        cmbAniosAnt.ListIndex = 0
    Else
        cmbAniosAnt.Enabled = False
    End If
    Set rsant = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If PassAnterior <> "" Then
        If ValidaUser(PassAnterior) = False Then
            End
        End If
    Else
        If ValidaUser(txtClave) = False Then
            End
        End If
    End If
End Sub

Private Sub txtClave_Change()
    Dim ClaveDeBusqueda As String, TextoCodificado As String, TextoOriginal As String
    Dim I As Integer
    For I = 0 To 255
        ClaveDeBusqueda = ClaveDeBusqueda + Chr$(I)
    Next
    TextoOriginal = ChrTran(strClaveUsuario, ClaveAleatoria, ClaveDeBusqueda)
    cmdAceptar.Enabled = False
    If UCase(TextoOriginal) = UCase(txtClave) And txtConfirmaPass.Visible = False Then
        cmdCPass.Enabled = True
        If Len(UCase(TextoOriginal)) >= 4 Then
            cmdAceptar.Enabled = True
        End If
    Else
        If txtClave.Text = Empty Then txtConfirmaPass = ""
        cmdCPass.Enabled = False
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
         If cmdAceptar.Enabled = True Then
            Call cmdAceptar_Click
        End If
     End If
End Sub

Private Sub txtConfirmaPass_Change()
   If UCase(txtClave) = UCase(txtConfirmaPass) And txtClave <> Empty Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub

