VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLeyenda 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Colores en Leyenda"
   ClientHeight    =   3420
   ClientLeft      =   7440
   ClientTop       =   7935
   ClientWidth     =   7020
   Icon            =   "frmLeyenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   7020
   Begin VB.PictureBox pic1 
      Height          =   195
      Left            =   4215
      ScaleHeight     =   135
      ScaleWidth      =   405
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Piccolor 
      Height          =   330
      Left            =   705
      ScaleHeight     =   270
      ScaleWidth      =   405
      TabIndex        =   2
      Top             =   67
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid msfleyenda 
      Height          =   1755
      Left            =   30
      TabIndex        =   0
      Top             =   1170
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColor       =   14737632
   End
   Begin Proyecto1.chameleonButton CmdSelColor 
      Height          =   360
      Left            =   1230
      TabIndex        =   5
      ToolTipText     =   "Ver Datos"
      Top             =   60
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   635
      BTYPE           =   14
      TX              =   "&Seleccionar"
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
      MICON           =   "frmLeyenda.frx":030A
      PICN            =   "frmLeyenda.frx":0326
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
      Left            =   6510
      TabIndex        =   6
      Top             =   3000
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
      MICON           =   "frmLeyenda.frx":0480
      PICN            =   "frmLeyenda.frx":049C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton btnModificar 
      Height          =   345
      Left            =   1170
      TabIndex        =   7
      ToolTipText     =   "Modificar"
      Top             =   3000
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmLeyenda.frx":0862
      PICN            =   "frmLeyenda.frx":087E
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
      Left            =   4305
      TabIndex        =   8
      ToolTipText     =   "Deshacer"
      Top             =   3000
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
      MICON           =   "frmLeyenda.frx":0CAC
      PICN            =   "frmLeyenda.frx":0CC8
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
      Left            =   4725
      TabIndex        =   9
      ToolTipText     =   "Guardar"
      Top             =   3000
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmLeyenda.frx":120A
      PICN            =   "frmLeyenda.frx":1226
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
      Left            =   30
      TabIndex        =   10
      ToolTipText     =   "Nuevo"
      Top             =   3000
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmLeyenda.frx":1668
      PICN            =   "frmLeyenda.frx":1684
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
      Left            =   2340
      TabIndex        =   11
      ToolTipText     =   "Eliminar"
      Top             =   3000
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmLeyenda.frx":19EE
      PICN            =   "frmLeyenda.frx":1A0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblmodo 
      Height          =   135
      Left            =   3645
      TabIndex        =   12
      Top             =   3210
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSForms.TextBox txtdescrip 
      Height          =   600
      Left            =   1215
      TabIndex        =   4
      Top             =   480
      Width           =   5745
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "10134;1058"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
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
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   450
   End
End
Attribute VB_Name = "frmLeyenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    txtdescrip.BackColor = ColorDeshabilitado
    txtdescrip.Locked = True
End Sub

Private Sub btnEliminar_Click()
Dim res As Integer
    res = MsgBox("¿Esta Seguro que desea ELIMINAR el Color Seleccionado", vbQuestion + vbYesNo, gsNomSW)
    If res = 6 Then
        If Not VerificaReg(msfleyenda.TextMatrix(msfleyenda.Rowsel, 2)) Then
            EliminarColor msfleyenda.TextMatrix(msfleyenda.Rowsel, 2)
            CargarDatos
            ModoFormulario modConsulta
        Else
            MsgBox "No se puede Eliminar el color seleccionado porque existen movimientos registrados", vbInformation, gsNomSW
        End If
    End If
End Sub

Sub EliminarColor(Color)
    SQL = "delete from leyendatraslados where color = '" & Color & "'"
    oConexionMYSQL.Execute SQL
End Sub

Function VerificaReg(Color As String) As Boolean
    Dim SQL As String
    Dim RQ As MYSQL_RS
    VerificaReg = False
    SQL = "select count(*) as cant from calendario where sinbono ='" & Color & "'"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        If RQ.Fields("cant") > 0 Then
            VerificaReg = True
        End If
    End If
    Set RQ = Nothing
End Function

Private Sub btnGrabar_Click()
     If lblModo = "Nuevo" Then
        If Grabar Then
            CargarDatos
            ModoFormulario modConsulta
        End If
    End If
    If lblModo = "Modificar" Then
        If Actualizar Then
            CargarDatos
            ModoFormulario modConsulta
        End If
    End If
End Sub

Function Grabar() As Boolean
    Dim SQL As String
    Grabar = False
    SQL = "insert into leyendatraslados(color,descripcion) values('" & Piccolor.BackColor & "','" & Trim(txtdescrip.Text) & "')"
    oConexionMYSQL.Execute SQL
    Grabar = True
End Function

Function Actualizar() As Boolean
    Dim SQL As String
    Actualizar = False
    SQL = "update leyendatraslados set descripcion = '" & Trim(txtdescrip.Text) & "',color='" & Piccolor.BackColor & "' where color = '" & pic1.BackColor & "'"
    oConexionMYSQL.Execute SQL
    SQL = "update calendario set sinbono = '" & Piccolor.BackColor & "' where sinbono = '" & pic1.BackColor & "'"
    oConexionMYSQL.Execute SQL
    Actualizar = True
End Function

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    LimpiarDatos
    ModoFormulario modNuevo
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub CmdSelColor_Click()
    Piccolor.BackColor = frmColor.ShowColor(1, vbWhite, "Seleccionar color")
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    ModoFormulario modAccion
    CargarDatos
End Sub

Sub ConfiguraGrilla()
    With msfleyenda
        .Cols = 3
        .Rows = 1
        .Clear
        .FormatString = "Color|Descripción|"
        .ColWidth(0) = 800
        .ColWidth(1) = 6000
        .ColWidth(2) = 0
    End With
End Sub

Sub CargarDatos()
Dim SQL As String, I As Integer
Dim RQ As MYSQL_RS
    ConfiguraGrilla
    SQL = "select * from leyendatraslados"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    I = 1
    Do While Not RQ.EOF()
        msfleyenda.AddItem "" & vbTab & Trim(RQ.Fields("descripcion")) & vbTab & Trim(RQ.Fields("color"))
        msfleyenda.row = I
        msfleyenda.Col = 0
        msfleyenda.CellBackColor = Trim(RQ.Fields("color"))
        I = I + 1
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub

Private Sub LimpiarDatos()
    txtdescrip.Text = Empty
    Piccolor.BackColor = &H8000000F
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtdescrip.Locked = valor
    If valor = True Then
        txtdescrip.BackColor = ColorDeshabilitado
    Else
        txtdescrip.BackColor = ColorHabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             lblModo = "Acción"
             ConfigurarBotones cfgCancelar
             BloqueoControles True
             CmdSelColor.Enabled = False
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             lblModo = "Nuevo"
             BloqueoControles False
             ConfigurarBotones cfgNuevo
             CmdSelColor.Enabled = True
             Exit Sub
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
             CmdSelColor.Enabled = False
         Case ModoForm.modEditar
             lblModo = "Modificar"
             BloqueoControles False
             ConfigurarBotones cfgModificar
             CmdSelColor.Enabled = True
         Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = True
            BtnCancelar.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            BtnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = True
            BtnCancelar.Enabled = True
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = True
            btnEliminar.Enabled = False
            btnGrabar.Enabled = False
            BtnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = True
            btnEliminar.Enabled = True
            btnGrabar.Enabled = False
            BtnCancelar.Enabled = True
        Case ConfigBotones.cfgCancelar
            BtnNuevo.Enabled = True
            BtnModificar.Enabled = True
            btnEliminar.Enabled = True
            btnGrabar.Enabled = False
            BtnCancelar.Enabled = False
    End Select
End Sub

Private Sub msfleyenda_RowColChange()
    With msfleyenda
        Piccolor.BackColor = .TextMatrix(.Rowsel, 2)
        pic1.BackColor = .TextMatrix(.Rowsel, 2)
        txtdescrip.Text = .TextMatrix(.Rowsel, 1)
    End With
End Sub
