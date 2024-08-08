VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmRubrosRem 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rubros de Remuneración"
   ClientHeight    =   5985
   ClientLeft      =   3210
   ClientTop       =   4320
   ClientWidth     =   10320
   Icon            =   "frmRubrosRem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10320
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      BackColor       =   10442041
      TabCaption(0)   =   "Rubro"
      TabPicture(0)   =   "frmRubrosRem.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Afectación"
      TabPicture(1)   =   "frmRubrosRem.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H009F5539&
         Height          =   4215
         Left            =   30
         TabIndex        =   16
         Top             =   435
         Width           =   10170
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRubros 
            Height          =   4050
            Left            =   30
            TabIndex        =   18
            Top             =   135
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7144
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H009F5539&
         Height          =   4230
         Left            =   -74970
         TabIndex        =   15
         Top             =   435
         Width           =   10170
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAfectos 
            Height          =   4050
            Left            =   30
            TabIndex        =   19
            Top             =   135
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7144
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.ComboBox cboGenFlx 
         Height          =   315
         Left            =   -65520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin Proyecto1.chameleonButton btnSalir 
      Height          =   345
      Left            =   9780
      TabIndex        =   6
      Top             =   5580
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
      MICON           =   "frmRubrosRem.frx":0182
      PICN            =   "frmRubrosRem.frx":019E
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
      Left            =   10290
      TabIndex        =   7
      Top             =   6360
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
      MICON           =   "frmRubrosRem.frx":0564
      PICN            =   "frmRubrosRem.frx":0580
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
      Left            =   2670
      TabIndex        =   8
      ToolTipText     =   "Eliminar"
      Top             =   5580
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Eliminar"
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
      MICON           =   "frmRubrosRem.frx":0AC2
      PICN            =   "frmRubrosRem.frx":0ADE
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
      Left            =   1380
      TabIndex        =   9
      ToolTipText     =   "Modificar"
      Top             =   5580
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Modificar"
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
      MICON           =   "frmRubrosRem.frx":0C38
      PICN            =   "frmRubrosRem.frx":0C54
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
      Left            =   4350
      TabIndex        =   10
      ToolTipText     =   "Deshacer"
      Top             =   5580
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
      MICON           =   "frmRubrosRem.frx":1082
      PICN            =   "frmRubrosRem.frx":109E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdArchivo 
      Left            =   8820
      Top             =   9030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Proyecto1.chameleonButton btnGrabar 
      Height          =   345
      Left            =   6780
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   5580
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
      MICON           =   "frmRubrosRem.frx":15E0
      PICN            =   "frmRubrosRem.frx":15FC
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
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Nuevo"
      Top             =   5580
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Nuevo"
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
      MICON           =   "frmRubrosRem.frx":1A3E
      PICN            =   "frmRubrosRem.frx":1A5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      TabIndex        =   12
      Top             =   540
      Width           =   795
   End
   Begin MSForms.TextBox txtCod 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   480
      Width           =   675
      VariousPropertyBits=   746604571
      MaxLength       =   20
      Size            =   "1191;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   1470
      TabIndex        =   20
      Top             =   540
      Width           =   1095
   End
   Begin MSForms.TextBox txtDescrip 
      Height          =   315
      Left            =   2610
      TabIndex        =   3
      Top             =   480
      Width           =   7635
      VariousPropertyBits=   746604571
      Size            =   "13467;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
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
      TabIndex        =   17
      Top             =   150
      Width           =   495
   End
   Begin MSForms.ComboBox cboTipoRubro 
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   90
      Width           =   6150
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "10848;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblSituacEmp 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
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
      Left            =   8010
      TabIndex        =   13
      Top             =   5700
      Width           =   1725
   End
   Begin VB.Label lblModo 
      BackStyle       =   0  'Transparent
      Caption         =   "Acción"
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
      Height          =   225
      Left            =   8490
      TabIndex        =   11
      Top             =   5730
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "frmRubrosRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As New FrmConsultas
Private generocod As Boolean
Dim SQL As String
Dim Rs As MYSQL_RS

Private Sub TipoRubro(cbo As MSForms.ComboBox)
    SQL = "Select * from pl_tiporubros order by tipo,descrip"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    cbo.Clear
    i = 0
    Do While Not Rs.EOF
        cbo.AddItem CE(Rs.Fields("TIPO")) & " - " & CE(Rs.Fields("DESCRIP"))
        cbo.List(i, 1) = CE(Rs.Fields("CODIGO"))
        i = i + 1
        Rs.MoveNext
    Loop
    cbo.ListIndex = 0
    Rubros 0
    Set Rs = Nothing
End Sub

Private Sub Rubros(Tipo As String)
    Dim i As Integer
    i = 0
    With mshRubros
        .Clear
        ConfiguraGrillaRubros
        SQL = "Select * from pl_Rubrosremunerativos where tipo='" & Tipo & "' order by codigo"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        Do While Not Rs.EOF
            i = i + 1
            .Rows = i + 1
            .Col = 3
            .row = i
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .Col = 4
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .Col = 6
            .row = i
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .Col = 7
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .FixedCols = 1
            .FixedRows = 1
            .TextMatrix(i, 0) = str(i)
            .TextMatrix(i, 1) = CE(Rs.Fields("codigo"))
            .TextMatrix(i, 2) = CE(Rs.Fields("descrip"))
            .TextMatrix(i, 3) = IIf(CE(Rs.Fields("calculo")) = "S", strChecked, strUnChecked)
            .TextMatrix(i, 4) = IIf(CE(Rs.Fields("principal")) = "S", strChecked, strUnChecked)
            .TextMatrix(i, 6) = IIf(CE(Rs.Fields("archivo")) = "S", strChecked, strUnChecked)
            .TextMatrix(i, 7) = IIf(CE(Rs.Fields("actsueldafp")) = "S", strChecked, strUnChecked)
            Rs.MoveNext
        Loop
        Set Rs = Nothing
        If i > 0 Then
            btnEliminar.Enabled = True
            btnModificar.Enabled = True
        Else
            btnEliminar.Enabled = False
            btnModificar.Enabled = False
        End If
    End With
End Sub

Private Sub Rubros_Afectos(rub As String)
    Dim i As Integer
    i = 0
    With mshAfectos
        .Clear
        ConfiguraGrillaAfectos
        SQL = "SELECT a.codigo,a.descrip," & _
              " IFNULL((select descrip FROM pl_rubros_afectos WHERE CODAFEC=A.CODIGO AND CODRUB='" & rub & "'),'N')" & _
              " AS AFECTO " & _
              " from pl_rubrosafectacion as a " & _
              " ORDER BY A.CODIGO"
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        Do While Not Rs.EOF
            i = i + 1
            .Rows = i + 1
            .Col = 3
            .row = i
            .CellFontName = "Wingdings"
            .CellFontSize = 14
            .FixedCols = 1
            .FixedRows = 1
            .TextMatrix(i, 0) = str(i)
            .TextMatrix(i, 1) = CE(Rs.Fields("codigo"))
            .TextMatrix(i, 2) = CE(Rs.Fields("descrip"))
            .TextMatrix(i, 3) = IIf(Trim(CE(Rs.Fields("afecto"))) = "", strChecked, strUnChecked)
            Rs.MoveNext
        Loop
        Set Rs = Nothing
    End With
End Sub

Private Sub ConfiguraGrillaRubros()
    With mshRubros
        .Clear
        .Refresh
        .Rows = 1
        .Cols = 8
        .RowHeight(0) = 315
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 500
        .TextMatrix(0, 1) = "Código"
        .ColWidth(2) = 6800
        .TextMatrix(0, 2) = Space(30) + "Descripción"
        .ColWidth(3) = 0
        .TextMatrix(0, 3) = "Cálculo"
        .ColWidth(4) = 700
        .TextMatrix(0, 4) = "Principal"
        .ColWidth(5) = 0
        .TextMatrix(0, 5) = ""
        .ColWidth(6) = 800
        .TextMatrix(0, 6) = "Afec.Reten.Jud."
        .ColWidth(7) = 800
        .TextMatrix(0, 7) = "Afec.Sueldo/AFP"
        For i = 1 To .Cols - 1
            .Col = i
            .row = 0
            .CellBackColor = &H8000000F
        Next
    End With
End Sub

Private Sub ConfiguraGrillaAfectos()
    With mshAfectos
        .Clear
        .Refresh
        .Rows = 1
        .Cols = 4
        .RowHeight(0) = 315
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 800
        .TextMatrix(0, 1) = "Código"
        .ColWidth(2) = 5000
        .TextMatrix(0, 2) = Space(20) + "Descripción"
        .ColWidth(3) = 800
        .TextMatrix(0, 3) = "Afecto"
        For i = 1 To .Cols - 1
            .Col = i
            .row = 0
            .CellBackColor = &H8000000F
        Next
    End With
End Sub

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub cboTipoRubro_Change()
    If cboTipoRubro.ListIndex >= 0 Then
        txtCod = ""
        txtdescrip = ""
        Rubros cboTipoRubro.List(cboTipoRubro.ListIndex, 1)
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    ModoFormulario modAccion
    ConfiguraGrillaRubros
    Call WheelHook(frmRubrosRem)
    Set oConsulta = New FrmConsultas
    SSTab1.Tab = 0
    TipoRubro cboTipoRubro
End Sub

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    If mshRubros.Rows > 1 Then
        mshRubros.row = 1
    Else
        cboTipoRubro.SetFocus
    End If
    ModoFormulario modConsulta
End Sub

Private Sub Eliminar(Cod As String)
    SQL = "Delete from pl_rubrosremunerativos where codigo = '" & Cod & "' "
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
End Sub

Private Sub btnEliminar_Click()
    Dim RES As Integer
    RES = MsgBox("¿Esta seguro que desea Eliminar el rubro, " & vbNewLine & " con código Nro. " & Trim(txtCod) & " ?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        Eliminar Trim(txtCod)
        Rubros cboTipoRubro.List(cboTipoRubro.ListIndex, 1)
        ModoFormulario modAccion
    End If
End Sub

Private Sub btnGrabar_Click()
    Dim RES As Integer
    If lblModo = "Nuevo" Then
        If Grabar Then
            Rubros cboTipoRubro.List(cboTipoRubro.ListIndex, 1)
            ModoFormulario modConsulta
        End If
    End If
    If lblModo = "Modificar" Then
        If Actualizar Then
            Rubros cboTipoRubro.List(cboTipoRubro.ListIndex, 1)
            ModoFormulario modConsulta
        End If
    End If
End Sub

Private Function Grabar() As Boolean
On Error GoTo problema
    Grabar = False
    SQL = "Insert into pl_rubrosremunerativos (codigo,tipo,descrip,principal,archivo,actsueldafp)  values " & _
           "('" & Trim(txtCod) & "','" & Trim(cboTipoRubro.List(cboTipoRubro.ListIndex, 1)) & "', " & _
           "'" & Trim(txtdescrip) & "','N','N','N')"
    oConexionMYSQL.Execute SQL
Grabar = True
Exit Function
problema:
    MsgBox "Ocurrió un error al momento de grabar el registro", vbOKOnly + vbExclamation, "NOVPeru"
    Grabar = False
    Exit Function
End Function

Private Function Actualizar() As Boolean
On Error GoTo problema
    Dim i As Integer
    Actualizar = False
    SQL = "Update pl_rubrosremunerativos set descrip='" & Trim(txtdescrip) & "', " & _
          "principal='" & IIf(Trim(mshRubros.TextMatrix(mshRubros.Rowsel, 4)) = strChecked, "S", "N") & "'," & _
          "archivo='" & IIf(Trim(mshRubros.TextMatrix(mshRubros.Rowsel, 6)) = strChecked, "S", "N") & "'," & _
          "actsueldafp='" & IIf(Trim(mshRubros.TextMatrix(mshRubros.Rowsel, 7)) = strChecked, "S", "N") & "'" & _
          " where codigo='" & Trim(txtCod) & "'"
    oConexionMYSQL.Execute SQL
    SQL = "delete from pl_rubros_afectos where codrub = '" & Trim(txtCod) & "'"
    oConexionMYSQL.Execute SQL
    With mshAfectos
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) = strChecked Then
                SQL = "insert into pl_rubros_afectos(codrub,codafec,descrip) values ('" & Trim(txtCod) & "','" & Trim(.TextMatrix(i, 1)) & "','')"
                oConexionMYSQL.Execute SQL
            End If
        Next
    End With
    Actualizar = True
Exit Function
problema:
    MsgBox "Ocurrió un error al momento de grabar el registro", vbOKOnly + vbExclamation, "NOVPeru"
    Actualizar = False
    Exit Function
End Function

Private Sub btnNuevo_Click()
    Dim RES As Integer
    ModoFormulario modNuevo
    RES = MsgBox("Desea Generar Código?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        generocod = True
        txtCod = GenCod(cboTipoRubro.List(cboTipoRubro.ListIndex, 1))
        txtCod.SetFocus
        txtCod.SelStart = 3
    Else
        txtCod.SetFocus
        generocod = False
    End If
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim RES As Integer
    RES = MsgBox("¿Desea salir del formulario?", vbYesNo + vbQuestion, "Aviso")
    If RES = vbNo Then
        Cancel = 1
    Else
        WheelUnHook
        Set oConsulta = Nothing
        Set Rs = Nothing
    End If
End Sub

Private Function GenCod(Tipo As String) As String
    Dim Cont As Integer
    SQL = "SELECT MAX(CODIGO) AS MAXIMO FROM pl_rubrosremunerativos where tipo='" & Tipo & "'"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Rs.RecordCount > 0 Then
        GenCod = Right("000" + RTrim(CStr(CDbl(Rs.Fields("MAXIMO")) + 1)), 3)
    Else
        GenCod = "000"
    End If
    Set Rs = Nothing
End Function

Private Sub LimpiarDatos()
    txtCod = ""
    txtdescrip = ""
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtCod.Locked = valor
    txtdescrip.Locked = valor
    cboTipoRubro.Locked = Not valor
    If valor = True Then
        txtCod.BackColor = ColorDeshabilitado
        txtdescrip.BackColor = ColorDeshabilitado
        mshRubros.BackColor = ColorDeshabilitado
        mshAfectos.BackColor = ColorDeshabilitado
        cboTipoRubro.BackColor = ColorHabilitado
    Else
        txtCod.BackColor = ColorHabilitado
        txtdescrip.BackColor = ColorHabilitado
        mshRubros.BackColor = ColorHabilitado
        mshAfectos.BackColor = ColorHabilitado
        cboTipoRubro.BackColor = ColorDeshabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             lblModo = "Acción"
             BloqueoControles True
             ConfigurarBotones cfgCancelar
             BtnNuevo.Enabled = True
             SSTab1.Tab = 0
             generocod = False
             txtCod.SelStart = 0
             frmRubrosRem.Caption = "Rubros Remunerativos"
             btnSalir.Enabled = True
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             lblModo = "Nuevo"
             BloqueoControles False
             ConfigurarBotones cfgNuevo
             txtCod.SetFocus
             SSTab1.Tab = 0
             frmRubrosRem.Caption = "Rubros Remunetarivos - [ Registrar Nuevo ]"
             Exit Sub
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
             generocod = False
             btnCancelar.Enabled = False
             frmRubrosRem.Caption = "Rubros Remunerativos"
             Exit Sub
        Case ModoForm.modEditar
             lblModo = "Modificar"
             BloqueoControles False
             ConfigurarBotones cfgModificar
             txtCod.Locked = True
             txtCod.BackColor = ColorDeshabilitado
             generocod = False
             txtdescrip.SetFocus
             Exit Sub
    End Select
End Sub

Public Sub ConfigurarBotones(cfg As ConfigBotones)
    Dim RES As Integer
    Select Case cfg
        Case ConfigBotones.cfgNuevo
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnCancelar.Enabled = True
            btnReporte.Enabled = False
            btnSalir.Enabled = False
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgModificar
            BtnNuevo.Enabled = False
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = True
            btnCancelar.Enabled = True
            Publimensaje = "modificar"
            Exit Sub
        Case ConfigBotones.cfgEliminar
            BtnNuevo.Enabled = True
            btnModificar.Enabled = False
            btnEliminar.Enabled = False
            btnGrabar.Enabled = False
            btnReporte.Enabled = False
            btnCancelar.Enabled = False
            btnSalir.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgGrabar
            BtnNuevo.Enabled = True
            btnGrabar.Enabled = False
            btnCancelar.Enabled = True
            btnModificar.Enabled = True
            Publimensaje = ""
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo"
                    RES = MsgBox("Desea Cancelar el registro?", vbQuestion + vbYesNo, gsNomSW)
                    If RES = 6 Then
                        Publimensaje = ""
                        ModoFormulario modAccion
                    End If
                Case "Consulta"
                     Publimensaje = ""
                     ModoFormulario modAccion
                     BtnNuevo.Enabled = True
                     btnGrabar.Enabled = False
                     btnReporte.Enabled = False
                     btnCancelar.Enabled = False
                 Case "Modificar"
                    If btnEliminar.tag <> "" Then btnEliminar.Enabled = btnEliminar.tag Else: btnEliminar.Enabled = False
            End Select
    End Select
End Sub

Private Sub mshAfectos_Click()
Dim SCol As Integer
    With mshAfectos
        SCol = .Col
        If .BackColor = ColorHabilitado Then
            If SCol = 3 Then
                If .TextMatrix(.row, 3) = strChecked Then
                    .TextMatrix(.row, 3) = strUnChecked
                Else
                    .TextMatrix(.row, 3) = strChecked
                End If
            End If
        End If
    End With
End Sub

Private Sub mshRubros_Click()
Dim SCol As Integer
    With mshRubros
        SCol = .Col
        If .BackColor = ColorHabilitado Then
            If SCol = 4 Then
                If .TextMatrix(.row, 4) = strChecked Then
                    .TextMatrix(.row, 4) = strUnChecked
                Else
                    .TextMatrix(.row, 4) = strChecked
                End If
            End If
        End If
    End With
End Sub

Private Sub mshRubros_DblClick()
    SSTab1.Tab = 1
End Sub

Private Sub txtdescrip_GotFocus()
    mark1 txtdescrip
End Sub

Private Sub txtNombre1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        btnGrabar.SetFocus
        Exit Sub
    End If
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flxDependientes
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

Private Sub mshRubros_RowColChange()
    With mshRubros
        If .row > 0 Then
            Rubros_Afectos .TextMatrix(.row, 1)
            txtCod = .TextMatrix(.row, 1)
            txtdescrip = .TextMatrix(.row, 2)
        Else
            txtCod = ""
            txtdescrip = ""
        End If
    End With
End Sub

Private Sub TxtDescrip_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then btnGrabar_Click
End Sub
