VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmPolizas 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polizas Seguro"
   ClientHeight    =   3270
   ClientLeft      =   8010
   ClientTop       =   7095
   ClientWidth     =   7620
   Icon            =   "frmPolizas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7620
   Begin NOVAdmin.flxEdit msfpoliza 
      Height          =   2025
      Left            =   30
      TabIndex        =   15
      Top             =   780
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   3572
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   7110
      TabIndex        =   5
      Top             =   2865
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
      MICON           =   "frmPolizas.frx":014A
      PICN            =   "frmPolizas.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdModificar 
      Height          =   345
      Left            =   1275
      TabIndex        =   6
      ToolTipText     =   "Modificar"
      Top             =   2865
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
      MICON           =   "frmPolizas.frx":052C
      PICN            =   "frmPolizas.frx":0548
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
      Height          =   345
      Left            =   4455
      TabIndex        =   7
      ToolTipText     =   "Deshacer"
      Top             =   2865
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
      MICON           =   "frmPolizas.frx":0976
      PICN            =   "frmPolizas.frx":0992
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
      Height          =   345
      Left            =   4875
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   2865
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
      MICON           =   "frmPolizas.frx":0ED4
      PICN            =   "frmPolizas.frx":0EF0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdNuevo 
      Height          =   345
      Left            =   45
      TabIndex        =   8
      ToolTipText     =   "Nuevo"
      Top             =   2865
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
      MICON           =   "frmPolizas.frx":1332
      PICN            =   "frmPolizas.frx":134E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtFecIni 
      Height          =   285
      Left            =   3870
      TabIndex        =   2
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   101187585
      CurrentDate     =   38637
   End
   Begin MSComCtl2.DTPicker dtFecFin 
      Height          =   285
      Left            =   6150
      TabIndex        =   3
      Top             =   450
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   503
      _Version        =   393216
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   101187585
      CurrentDate     =   38637
   End
   Begin Proyecto1.chameleonButton cmdEliminar 
      Height          =   345
      Left            =   2505
      TabIndex        =   14
      Top             =   2865
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
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
      MICON           =   "frmPolizas.frx":16B8
      PICN            =   "frmPolizas.frx":16D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblcod 
      Caption         =   "Label3"
      Height          =   180
      Left            =   7950
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin Vig."
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
      Left            =   5430
      TabIndex        =   12
      Top             =   495
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio Vig."
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
      Left            =   2940
      TabIndex        =   11
      Top             =   495
      Width           =   870
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° Póliza"
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
      Left            =   30
      TabIndex        =   10
      Top             =   495
      Width           =   795
   End
   Begin MSForms.ComboBox cboSeguro 
      Height          =   315
      Left            =   855
      TabIndex        =   0
      Top             =   90
      Width           =   4440
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "7832;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
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
      Left            =   75
      TabIndex        =   9
      Top             =   105
      Width           =   615
   End
   Begin MSForms.TextBox txtPoliza 
      Height          =   315
      Left            =   855
      TabIndex        =   1
      Top             =   435
      Width           =   1995
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "3519;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoPoliza As String
Dim lblModo As String

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    ModoFormulario modAccion
    ConfigurarBotones cfgCancelar
End Sub

Private Sub cmdEliminar_Click()
    Dim res As Integer
    If msfpoliza.Rows > 1 Then
        res = MsgBox("¿Seguro que desea Eliminar el Registro de la póliza N° " & txtPoliza.Text & "?", vbQuestion + vbYesNo, gsNomSW)
        If res = 6 Then
            Eliminar
            ModoFormulario modAccion
        End If
    End If
End Sub

Private Sub Eliminar()
    Dim SQL As String
    SQL = "delete from polizas where codigo = '" & Trim(lblcod.Caption) & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
End Sub

Private Sub cmdGrabar_Click()
    If lblModo = "Nuevo" Then
        If Grabar Then
            MostrarPolizas
            ModoFormulario modConsulta
        End If
    End If
    If lblModo = "Modificar" Then
        If Actualizar Then
            MostrarPolizas
            ModoFormulario modConsulta
        End If
    End If
End Sub

Private Function Actualizar() As Boolean
    Dim SQL As String
    Dim FecIni As String
    Dim FecFin As String
    Actualizar = False
    If Validar Then
        FecIni = IIf(Not IsDate(dtFecIni.Value), Empty, dtFecIni.Value)
        FecFin = IIf(Not IsDate(dtFecFin.Value), Empty, dtFecFin.Value)
        SQL = "Call Update_Polizas ('" & Trim(lblcod) & "','" & cboSeguro.List(cboSeguro.ListIndex, 1) & "','" & Trim(txtPoliza) & "'," & _
              "'" & FecIni & "','" & FecFin & "','" & TipoPoliza & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
        Actualizar = True
    End If
End Function

Private Function Grabar() As Boolean
    Dim SQL As String
    Dim FecIni As String
    Dim FecFin As String
    Dim Cod As Integer
    Grabar = False
    If Validar Then
        FecIni = IIf(Not IsDate(dtFecIni.Value), Empty, dtFecIni.Value)
        FecFin = IIf(Not IsDate(dtFecFin.Value), Empty, dtFecFin.Value)
        Cod = GeneraCod
        SQL = "Call Insert_Polizas (" & Cod & ",'" & cboSeguro.List(cboSeguro.ListIndex, 1) & "','" & Trim(txtPoliza) & "'," & _
              "'" & FecIni & "','" & FecFin & "','" & TipoPoliza & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        Grabar = True
    End If
End Function

Private Function Validar() As Boolean
    Dim I As Integer, J As Integer, numhij As Integer
    Validar = True
    If Trim(txtPoliza.Text) = "" Then Validar = False: MsgBox "Ingrese un Número de Póliza", vbInformation, gsNomSW: txtPoliza.SetFocus: Exit Function
    If cboSeguro.ListIndex <= 0 Then Validar = False: MsgBox "Debe escoger un tipo de Seguro", vbInformation, gsNomSW: cboSeguro.SetFocus: Exit Function
End Function

Private Sub cmdModificar_Click()
    If msfpoliza.Rows > 1 Then
        ModoFormulario modEditar
    End If
End Sub

Private Sub cmdNuevo_Click()
    ModoFormulario modNuevo
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    dtFecIni = Date
    dtFecFin = Date
    ModoFormulario modAccion
    msfpoliza.BackColor = ColorDeshabilitado
    Seguros cboSeguro
End Sub

Private Sub LimpiarDatos()
    txtPoliza = ""
    lblcod = ""
    dtFecIni.Value = Date
    dtFecFin.Value = Date
    Seguros cboSeguro
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtPoliza.Locked = valor
    cboSeguro.Locked = valor
    dtFecIni.Enabled = Not valor
    dtFecFin.Enabled = Not valor
    If valor = True Then
        txtPoliza.BackColor = ColorDeshabilitado
        cboSeguro.BackColor = ColorDeshabilitado
    Else
        txtPoliza.BackColor = ColorHabilitado
        cboSeguro.BackColor = ColorHabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             lblModo = "Acción"
             BloqueoControles True
             Seguros cboSeguro
             ConfigurarBotones cfgCancelar
             MostrarPolizas
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             lblModo = "Nuevo"
             ConfigurarBotones cfgNuevo
             BloqueoControles False
             Seguros cboSeguro
             For I = 1 To msfpoliza.Rows - 1
                msfpoliza.row = I: msfpoliza.Col = 2
                msfpoliza.CellBackColor = ColorDeshabilitado
             Next
             Exit Sub
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
             Exit Sub
        Case ModoForm.modEditar
             lblModo = "Modificar"
             BloqueoControles False
             ConfigurarBotones cfgModificar
             For I = 1 To msfpoliza.Rows - 1
                msfpoliza.row = I: msfpoliza.Col = 2
                msfpoliza.CellBackColor = ColorDeshabilitado
             Next
             Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Dim res As Integer
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
            cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
        Case ConfigBotones.cfgModificar
            cmdNuevo.Enabled = False
            cmdModificar.Enabled = False
            cmdEliminar.Enabled = False
            cmdGrabar.Enabled = True
            cmdCancelar.Enabled = True
        Case ConfigBotones.cfgEliminar
            cmdNuevo.Enabled = True
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = False
            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = False
        Case ConfigBotones.cfgGrabar
            cmdNuevo.Enabled = True
            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = False
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
        Case ConfigBotones.cfgCancelar
            cmdNuevo.Enabled = True
            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = False
            cmdModificar.Enabled = True
            cmdEliminar.Enabled = True
    End Select
End Sub

Private Sub Seguros(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I As Integer
    SQL = "Select * from seguro order by descrip"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "00"
    I = 1
    Do While Not Rs.EOF
        cbo.AddItem CE(Rs.Fields("DESCRIP"))
        cbo.List(I, 1) = CE(Rs.Fields("CODIGO"))
        I = I + 1
        Rs.MoveNext
    Loop
    cbo.ListIndex = 0
    Set Rs = Nothing
End Sub

Sub MostrarPolizas()
    Dim SQL As String, I As Integer
    Dim RQ As MYSQL_RS
    ConfiguraGrilla
    SQL = "select * from polizas where tipo = '" & TipoPoliza & "' order by fecini desc"
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    I = 1
    With msfpoliza
        Do While Not RQ.EOF()
            .Rows = .Rows + 1
            .TextMatrix(I, 1) = Trim(RQ.Fields("codseguro"))
            .TextMatrix(I, 2) = DescripcionesdeCodigos("SEGURO", Trim(RQ.Fields("codseguro")))
            .TextMatrix(I, 3) = Trim(RQ.Fields("numpoliza"))
            .TextMatrix(I, 4) = Format(Trim(RQ.Fields("fecini")), "dd/mm/yyyy")
            .TextMatrix(I, 5) = Format(Trim(RQ.Fields("fecfin")), "dd/mm/yyyy")
            .TextMatrix(I, 6) = RQ.Fields("codigo")
            RQ.MoveNext
            I = I + 1
        Loop
    End With
    Set RQ = Nothing
End Sub

Sub ConfiguraGrilla()
    With msfpoliza
        .Clear
        .Cols = 7
        .Rows = 1
        .FixedCols = 1
        .ColWidth(0) = 200
        .ColWidth(1) = 400
        .TextMatrix(0, 1) = "Cod"
        .ColType(1) = cadena
        .ColMaxLength(1) = 2
        .CaracteresValidos(1) = "1234567890"
        .ColWidth(2) = 1500
        .TextMatrix(0, 2) = Space(4) & "Nombre"
        .ColType(2) = cadena
        .ColWidth(3) = 1900
        .TextMatrix(0, 3) = Space(10) & "Número"
        .ColType(3) = cadena
        .ColMaxLength(3) = 15
        .CaracteresValidos(3) = "ABCDEFGHIJKLMÑNOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz1234567890"
        .ColWidth(4) = 1300
        .TextMatrix(0, 4) = "Inicio Vig."
        .ColType(4) = fecha
        .ColMaxLength(4) = 10
        .CaracteresValidos(4) = "0123456789/"
        .ColWidth(5) = 1300
        .TextMatrix(0, 5) = Space(3) & "Fin Vig."
        .ColType(5) = fecha
        .ColMaxLength(5) = 10
        .CaracteresValidos(5) = "0123456789/"
        .ColWidth(6) = 0
    End With
End Sub

Private Function GeneraCod() As Long
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    Set Rs = oConexion.EjecutaSelectRS("select max(codigo) as maximo from polizas")
    If Not Rs.EOF Then
        GeneraCod = CEN(Rs.Fields("maximo")) + 1
    End If
    If Rs.EOF Then
        GeneraCod = 1
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Function

Private Sub msfpoliza_RowColChange()
    With msfpoliza
        txtPoliza = Trim(.TextMatrix(.row, 3))
        lblcod = Trim(.TextMatrix(.row, 6))
        For I = 1 To cboSeguro.ListCount - 1
            If CE(.TextMatrix(.row, 1)) = Trim(cboSeguro.List(I, 1)) Then
                cboSeguro.ListIndex = I
                Exit For
            Else
                cboSeguro.ListIndex = 0
            End If
        Next
        If IsDate(CE(.TextMatrix(.row, 4))) Then
            dtFecIni.Value = CE(.TextMatrix(.row, 4))
        End If
        If IsDate(CE(.TextMatrix(.row, 5))) Then
            dtFecFin.Value = CE(.TextMatrix(.row, 5))
        End If
    End With
End Sub
