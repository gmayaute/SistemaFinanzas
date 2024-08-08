VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Servicios"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14775
   Icon            =   "frmServicios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14775
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   2235
      Left            =   0
      TabIndex        =   7
      Top             =   -60
      Width           =   14745
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Detalle:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "U.M.:"
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
         Left            =   4530
         TabIndex        =   13
         Top             =   570
         Width           =   855
      End
      Begin VB.Label lblDivision 
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
         Left            =   6300
         TabIndex        =   12
         Top             =   210
         Width           =   4965
      End
      Begin MSForms.ComboBox cboClasificacion 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   570
         Width           =   2895
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "5106;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblUMedida 
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
         Left            =   6300
         TabIndex        =   11
         Top             =   570
         Width           =   4965
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clasificación:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "División:"
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
         Left            =   4530
         TabIndex        =   8
         Top             =   210
         Width           =   855
      End
      Begin MSForms.TextBox txtDivision 
         Height          =   345
         Left            =   5460
         TabIndex        =   2
         Top             =   195
         Width           =   765
         VariousPropertyBits=   746604571
         MaxLength       =   4
         Size            =   "1349;609"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtNombre 
         Height          =   345
         Left            =   1470
         TabIndex        =   5
         Top             =   930
         Width           =   9795
         VariousPropertyBits=   746604571
         Size            =   "17277;609"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDetalle 
         Height          =   795
         Left            =   1470
         TabIndex        =   6
         Top             =   1290
         Width           =   9795
         VariousPropertyBits=   -1400879077
         Size            =   "17277;1402"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtCodigo 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   210
         Width           =   1485
         VariousPropertyBits=   746604571
         ForeColor       =   128
         MaxLength       =   6
         Size            =   "2619;556"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtUM 
         Height          =   345
         Left            =   5460
         TabIndex        =   4
         Top             =   570
         Width           =   765
         VariousPropertyBits=   746604571
         MaxLength       =   2
         Size            =   "1349;609"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   6465
      Left            =   0
      TabIndex        =   16
      Top             =   2160
      Width           =   14745
      Begin VB.TextBox TxtCriterio 
         Height          =   285
         Left            =   2640
         TabIndex        =   25
         Top             =   5655
         Width           =   4650
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxServicios 
         Height          =   5325
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   14595
         _ExtentX        =   25744
         _ExtentY        =   9393
         _Version        =   393216
         BackColorBkg    =   8421504
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Proyecto1.chameleonButton btnModificar 
         Height          =   345
         Left            =   1290
         TabIndex        =   17
         ToolTipText     =   "Modificar"
         Top             =   6030
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmServicios.frx":0442
         PICN            =   "frmServicios.frx":045E
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
         Left            =   13110
         TabIndex        =   18
         Top             =   6030
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
         MICON           =   "frmServicios.frx":088C
         PICN            =   "frmServicios.frx":08A8
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
         Left            =   6180
         TabIndex        =   19
         ToolTipText     =   "Guardar"
         Top             =   6030
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
         MICON           =   "frmServicios.frx":0C6E
         PICN            =   "frmServicios.frx":0C8A
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
         Left            =   90
         TabIndex        =   20
         ToolTipText     =   "Nuevo"
         Top             =   6030
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
         MICON           =   "frmServicios.frx":10CC
         PICN            =   "frmServicios.frx":10E8
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
         Left            =   5670
         TabIndex        =   21
         ToolTipText     =   "Deshacer"
         Top             =   6030
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
         MICON           =   "frmServicios.frx":1452
         PICN            =   "frmServicios.frx":146E
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
         Left            =   2610
         TabIndex        =   22
         Top             =   6030
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
         MICON           =   "frmServicios.frx":19B0
         PICN            =   "frmServicios.frx":19CC
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
         Left            =   12570
         TabIndex        =   23
         Top             =   6030
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
         MICON           =   "frmServicios.frx":1E0E
         PICN            =   "frmServicios.frx":1E2A
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
         Left            =   90
         TabIndex        =   26
         Top             =   5640
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
      Begin VB.Label lblModo 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   6090
         Visible         =   0   'False
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConsulta As New FrmConsultas
Dim fila As Integer

Private Sub Clasificacion(cbo As MSForms.ComboBox)
    Dim SQL As String
    Dim i As Integer
    Dim rsclasf As MYSQL_RS
    SQL = "Select * from clasif_serv"
    Set rsclasf = oConexion.EjecutaSelectRS(SQL)
    i = 0
    cbo.Clear
    Do While Not rsclasf.EOF
        cbo.AddItem CE(rsclasf.Fields("Descrip"))
        cbo.List(i, 1) = CE(rsclasf.Fields("Codigo"))
        i = i + 1
        rsclasf.MoveNext
    Loop
    cbo.ListIndex = 0
    Set rsclasf = Nothing
End Sub

Private Sub btnCancelar_Click()
    ConfigurarBotones cfgCancelar
    flxServicios.row = 1
    flxServicios.ColSel = 5
    DesplazarxGrilla
End Sub

Private Sub btnEliminar_Click()
    Dim RES As Integer
    Dim SQL As String
    RES = MsgBox("¿Esta seguro que desea Eliminar el Servicio," _
                 & vbNewLine & vbNewLine & " con Código No. " _
                 & flxServicios.TextMatrix(flxServicios.row, 0) & " ?", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        If Eliminar(Trim(CE(flxServicios.TextMatrix(flxServicios.row, 0)))) Then
            SQL = "Call Delete_Servicio ( '" & flxServicios.TextMatrix(flxServicios.row, 0) & "' ); "
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
            fila = 1
            ModoFormulario modAccion
        Else
            fila = flxServicios.row
            ModoFormulario modAccion
            flxServicios.SetFocus
            Call keybd_event(vbKeyLeft, 0, 0, 0)
        End If
    End If
End Sub

Private Function Eliminar(Servicio As String) As Boolean
    Dim SQL As String
    Dim rselimina As MYSQL_RS
    Eliminar = False
    SQL = "Select codserv from serv_tarif where codserv = '" & Servicio & "'"
    Set rselimina = oConexion.EjecutaSelectRS(SQL)
    If rselimina.RecordCount = 0 Then Eliminar = True: Exit Function
    If rselimina.RecordCount > 0 Then Eliminar = False: _
                                 MsgBox "No se puede eliminar el Servicio " & vbNewLine & vbNewLine & _
                                        Servicio & ", por estar relacionado", vbInformation, gsNomSW: _
                                        Exit Function
End Function

Private Sub btnGrabar_Click()
    If lblModo = "Modificar" Then
        If Actualizar Then
            fila = flxServicios.row
            ModoFormulario modAccion
            flxServicios.SetFocus
            Call keybd_event(vbKeyLeft, 0, 0, 0)
        End If
    End If
    If lblModo = "Nuevo" Then
        If Grabar Then
            fila = flxServicios.Rows
            ModoFormulario modAccion
            flxServicios.SetFocus
            Call keybd_event(vbKeyLeft, 0, 0, 0)
        End If
    End If
End Sub

Private Sub btnModificar_Click()
    ModoFormulario modEditar
End Sub

Private Sub btnNuevo_Click()
    ModoFormulario modNuevo
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cboCampos_Click()
    TxtCriterio = ""
    Call TxtCriterio_Change
End Sub

Private Sub cboClasificacion_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtUM.SetFocus
    End If
End Sub

Private Sub flxServicios_DblClick()
    btnModificar_Click
End Sub

Private Sub flxServicios_RowColChange()
    DesplazarxGrilla
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call WheelHook(frmServicios)
    fila = 1
    ModoFormulario modAccion
    LlenaCboCampos
    Set oConsulta = New FrmConsultas
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
End Sub

Private Sub TxtCriterio_Change()
    Dim filtro As String
    filtro = TxtCriterio.Text
    LlenarGrilla cboCampos.List(cboCampos.ListIndex, 1), filtro
End Sub

Private Sub txtDivision_Change()
    If TxtDivision = Empty Then
        lblDivision = Empty
    End If
End Sub

Private Sub txtDivision_GotFocus()
    mark1 TxtDivision
End Sub

Private Sub txtDivision_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        TxtDivision = Right("0000" & TxtDivision, 4)
        lblDivision = DescripcionesdeCodigos("DES_DIVISION", Trim(CE(TxtDivision)))
    End If
    If KeyCode = vbKeyF1 And TxtDivision.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 2
            .pCol = 0: .pAnchoCol = 800
            .pCol = 1: .pAnchoCol = 3000
            .pTitulo = "ccHFM"
            .pForm = FORM_SERVICIOS
            .pCaso = LABEL_DIVISIONES
            .Show
        End With
        cboClasificacion.SetFocus
    End If
End Sub

Private Sub ConfigGrilla()
    With flxServicios
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 6
        
        .ColWidth(0) = 1000
        .TextMatrix(0, 0) = Space(1) + "Código"
        
        .ColWidth(1) = 500
        .TextMatrix(0, 1) = Space(0) + "Clasf."
        
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = Space(2) + "ccHFM."
        
        .ColWidth(3) = 5000
        .TextMatrix(0, 3) = Space(7) + "Nombre"
        
        .ColWidth(4) = 500
        .TextMatrix(0, 4) = Space(1) + "U.M."
        
        .ColWidth(5) = 6500
        .TextMatrix(0, 5) = Space(2) + "Descripción"
    End With
End Sub
Private Sub LlenaCboCampos()
    With cboCampos
        .Clear
        .AddItem "CODIGO"
        .List(0, 1) = "codigo"
        
        .AddItem "CODCLAF"
        .List(1, 1) = "codclasf"
        
        .AddItem "CODDIV"
        .List(2, 1) = "coddiv"
        
        .AddItem "NOMBRE"
        .List(3, 1) = "nombre"
        
        .AddItem "UM"
        .List(4, 1) = "cod_um"
        
        .AddItem "DESCRIPCION"
        .List(5, 1) = "descrip"
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub


Private Sub LlenarGrilla(Optional criterio As String, Optional filtro As String)
    Dim SQL As String
    Dim i As Integer
    Dim rsgrid As MYSQL_RS
    
    SQL = "Select * from servicio where " & criterio & " like '%" & filtro & "%' ORDER BY " & criterio

    Set rsgrid = oConexion.EjecutaSelectRS(SQL)
    i = 1
    If rsgrid.RecordCount > 0 Then ConfigGrilla
    Do While Not rsgrid.EOF
        With flxServicios
            .TextMatrix(i, 0) = CE(rsgrid.Fields("Codigo"))
            .TextMatrix(i, 1) = CE(rsgrid.Fields("CodClasf"))
            .TextMatrix(i, 2) = CE(rsgrid.Fields("CodDiv"))
            .TextMatrix(i, 3) = Space(2) & CE(rsgrid.Fields("Nombre"))
            .TextMatrix(i, 4) = CE(rsgrid.Fields("Cod_UM"))
            .TextMatrix(i, 5) = Space(2) & Trim(CE(rsgrid.Fields("Descrip")))
            rsgrid.MoveNext
            .Rows = .Rows + 1
            i = i + 1
        End With
    Loop
     If rsgrid.RecordCount > 0 Then flxServicios.Rows = flxServicios.Rows - 1
    Set rsgrid = Nothing
End Sub

Private Sub DesplazarxGrilla()
    Dim i As Integer
    With flxServicios
        txtCodigo = CE(.TextMatrix(.row, 0))
        TxtDivision = CE(.TextMatrix(.row, 2))
        lblDivision = DescripcionesdeCodigos("DES_DIVISION", Trim(CE(.TextMatrix(.row, 2))))
        txtNombre = CE(Trim(.TextMatrix(.row, 3)))
        txtUM = CE(.TextMatrix(.row, 4))
        lblUMedida = DescripcionesdeCodigos("UM", CE(.TextMatrix(.row, 4)))
        txtDetalle = CE(Trim(.TextMatrix(.row, 5)))
        For i = 0 To cboClasificacion.ListCount - 1
            If .TextMatrix(.row, 1) = cboClasificacion.List(i, 1) Then
                cboClasificacion.ListIndex = i
                Exit For
            End If
        Next
    End With
End Sub

Private Function GeneraCod() As String
    Dim SQL As String
    Dim rscod As MYSQL_RS
    SQL = "Select Max(Codigo) as codigo from servicio "
    Set rscod = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(rscod.Fields("Codigo")) Then
        GeneraCod = Right("000000" & (CDbl(rscod.Fields("codigo")) + 1), 6)
    End If
    If IsNull(rscod.Fields("Codigo")) Then
        GeneraCod = "000001"
    End If
    Set rscod = Nothing
End Function

Private Function Grabar() As Boolean
    Dim SQL As String
    Grabar = False
    If Validar Then
        SQL = " Call Insert_Servicio ( '" & Trim(CE(txtCodigo)) & "', '" & cboClasificacion.List(cboClasificacion.ListIndex, 1) & "'," & _
              " '" & Trim(CE(TxtDivision)) & "', '" & Trim(CE(txtNombre)) & "', '" & Trim(CE(txtUM)) & "','" & Trim(CE(txtDetalle)) & "' );"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
        Grabar = True
    End If
End Function

Private Function Actualizar() As Boolean
    Dim SQL As String
    Actualizar = False
    If Validar Then
        SQL = " Call Update_Servicio ('" & Trim(CE(txtCodigo)) & "', '" & cboClasificacion.List(cboClasificacion.ListIndex, 1) & "'," & _
              " '" & Trim(CE(TxtDivision)) & "', '" & Trim(CE(txtNombre)) & "', '" & Trim(CE(txtUM)) & "', '" & Trim(CE(txtDetalle)) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Actualizar = True
    End If
End Function

Private Function Validar() As Boolean
    Validar = True
    If TxtDivision = Empty Then MsgBox "Debe Ingresar un ccHFM", vbInformation, gsNomSW: Validar = False: TxtDivision.SetFocus: Exit Function
    If txtUM = Empty Then MsgBox "Debe Ingresar la Unidad de Medida", vbInformation, gsNomSW: Validar = False: txtUM.SetFocus: Exit Function
    If txtNombre = Empty Then MsgBox "Debe Ingresar el Nombre del Servicio", vbInformation, gsNomSW: Validar = False: txtNombre.SetFocus: Exit Function
    If txtDetalle = Empty Then MsgBox "Debe Ingresar el Detalle del Servicio", vbInformation, gsNomSW: Validar = False: TxtDivision.SetFocus: Exit Function
End Function

Private Sub LimpiarDatos()
    txtCodigo = Empty
    txtDetalle = Empty
    TxtDivision = Empty
    txtNombre = Empty
    txtUM = Empty
    lblDivision = Empty
    lblUMedida = Empty
    cboClasificacion.Clear
End Sub

Public Sub BloqueoControles(valor As Boolean)
    txtCodigo.Locked = valor
    txtDetalle.Locked = valor
    TxtDivision.Locked = valor
    txtNombre.Locked = valor
    txtUM.Locked = valor
    cboClasificacion.Locked = valor
    flxServicios.Enabled = valor
    If valor = True Then
        txtCodigo.BackColor = ColorDeshabilitado
        txtDetalle.BackColor = ColorDeshabilitado
        TxtDivision.BackColor = ColorDeshabilitado
        txtNombre.BackColor = ColorDeshabilitado
        txtUM.BackColor = ColorDeshabilitado
        lblDivision.BackColor = ColorDeshabilitado
        lblUMedida.BackColor = ColorDeshabilitado
        cboClasificacion.BackColor = ColorDeshabilitado
        flxServicios.BackColor = ColorHabilitado
    Else
        txtCodigo.BackColor = ColorHabilitado
        txtDetalle.BackColor = ColorHabilitado
        TxtDivision.BackColor = ColorHabilitado
        txtNombre.BackColor = ColorHabilitado
        txtUM.BackColor = ColorHabilitado
        cboClasificacion.BackColor = ColorHabilitado
        flxServicios.BackColor = ColorDeshabilitado
    End If
End Sub

Public Sub ModoFormulario(modo As ModoForm)
    Select Case modo
        Case ModoForm.modAccion
             LimpiarDatos
             
             LlenarGrilla 1, 1
             Clasificacion cboClasificacion
             BloqueoControles True
             flxServicios.row = fila
             DesplazarxGrilla
             flxServicios.ColSel = 5
             lblModo = "Acción"
             ConfigurarBotones cfgGrabar
             Exit Sub
        Case ModoForm.modNuevo
             LimpiarDatos
             lblModo = "Nuevo"
             BloqueoControles False
             ConfigurarBotones cfgNuevo
             Clasificacion cboClasificacion
             txtCodigo = GeneraCod
             txtCodigo.Locked = True
             TxtDivision.SetFocus
             Exit Sub
        Case ModoForm.modConsulta
             lblModo = "Consulta"
             BloqueoControles True
             ConfigurarBotones cfgGrabar
         Case ModoForm.modEditar
              lblModo = "Modificar"
              Clasificacion cboClasificacion
              DesplazarxGrilla
              BloqueoControles False
              ConfigurarBotones cfgModificar
              txtCodigo.Locked = True
              txtCodigo.BackColor = ColorDeshabilitado
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
            btnSalir.Enabled = False
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
             BtnNuevo.Enabled = True
             btnModificar.Enabled = True
             btnEliminar.Enabled = True
             btnSalir.Enabled = True
             btnGrabar.Enabled = False
             btnCancelar.Enabled = False
            Exit Sub
        Case ConfigBotones.cfgCancelar
            Select Case lblModo.Caption
                Case "Nuevo"
                     ModoFormulario modConsulta
                     BtnNuevo.Enabled = True
                     btnModificar.Enabled = True
                     btnEliminar.Enabled = True
                     btnGrabar.Enabled = False
                     btnReporte.Enabled = False
                     btnCancelar.Enabled = False
                Case "Modificar"
                     ModoFormulario modConsulta
                     btnGrabar.Enabled = False
            End Select
    End Select
End Sub

Private Sub txtNombre_GotFocus()
    mark1 txtNombre
End Sub

Private Sub txtUM_Change()
    If txtUM = Empty Then
        lblUMedida = Empty
    End If
End Sub

Private Sub txtUM_GotFocus()
    mark1 txtUM
End Sub

Private Sub txtUM_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode = 13 Then
        txtUM = Right("00" & txtUM, 2)
        lblUMedida = DescripcionesdeCodigos("UM", Trim(CE(txtUM)))
    End If
    If KeyCode = vbKeyF1 And txtUM.BackColor = ColorHabilitado Then
        With oConsulta
            .pCols = 3
            .pCol = 0: .pAnchoCol = 500
            .pCol = 1: .pAnchoCol = 600
            .pCol = 2: .pAnchoCol = 1500
            .pTitulo = "Unidad de Medida"
            .pForm = FORM_SERVICIOS
            .pCaso = LABEL_UM
            .Show
        End With
        txtNombre.SetFocus
    End If
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
