VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmSedesTrabajo 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sedes de Trabajo"
   ClientHeight    =   4410
   ClientLeft      =   5370
   ClientTop       =   8775
   ClientWidth     =   7575
   Icon            =   "frmSedesTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDirec 
      Height          =   285
      Left            =   1215
      MaxLength       =   70
      TabIndex        =   1
      Top             =   480
      Width           =   6075
   End
   Begin VB.TextBox txtTelef 
      Height          =   285
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   2
      Top             =   900
      Width           =   2895
   End
   Begin VB.TextBox txtDes 
      Height          =   285
      Left            =   2805
      TabIndex        =   0
      Top             =   90
      Width           =   4485
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   1215
      MaxLength       =   2
      TabIndex        =   16
      Top             =   90
      Width           =   375
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H009F5539&
      Height          =   2535
      Left            =   60
      TabIndex        =   12
      Top             =   1320
      Width           =   7440
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSedes 
         Height          =   2190
         Left            =   90
         TabIndex        =   4
         Top             =   240
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   3863
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   12632256
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin Proyecto1.chameleonButton cmdSalir 
      Height          =   405
      Left            =   7020
      TabIndex        =   5
      Top             =   3900
      Width           =   435
      _ExtentX        =   767
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
      MICON           =   "frmSedesTrabajo.frx":0442
      PICN            =   "frmSedesTrabajo.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdVistaPreliminar 
      Height          =   405
      Left            =   6570
      TabIndex        =   6
      Top             =   3900
      Width           =   435
      _ExtentX        =   767
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
      MICON           =   "frmSedesTrabajo.frx":0824
      PICN            =   "frmSedesTrabajo.frx":0840
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
      Left            =   4590
      TabIndex        =   7
      Top             =   3900
      Width           =   435
      _ExtentX        =   767
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
      MICON           =   "frmSedesTrabajo.frx":0D82
      PICN            =   "frmSedesTrabajo.frx":0D9E
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
      Left            =   4140
      TabIndex        =   8
      Top             =   3900
      Width           =   435
      _ExtentX        =   767
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
      MICON           =   "frmSedesTrabajo.frx":11E0
      PICN            =   "frmSedesTrabajo.frx":11FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdEliminar 
      Height          =   405
      Left            =   2565
      TabIndex        =   9
      Top             =   3900
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmSedesTrabajo.frx":173E
      PICN            =   "frmSedesTrabajo.frx":175A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdModificar 
      Height          =   405
      Left            =   1305
      TabIndex        =   10
      Top             =   3900
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmSedesTrabajo.frx":1B9C
      PICN            =   "frmSedesTrabajo.frx":1BB8
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
      Height          =   405
      Left            =   45
      TabIndex        =   11
      Top             =   3900
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmSedesTrabajo.frx":1FE6
      PICN            =   "frmSedesTrabajo.frx":2002
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox cmbSede 
      Height          =   315
      Left            =   5325
      TabIndex        =   3
      Top             =   885
      Width           =   1965
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3466;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Left            =   4575
      TabIndex        =   18
      Top             =   945
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección"
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
      Left            =   135
      TabIndex        =   17
      Top             =   525
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono(s)"
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
      Left            =   135
      TabIndex        =   15
      Top             =   945
      Width           =   975
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1695
      TabIndex        =   14
      Top             =   135
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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
      Left            =   135
      TabIndex        =   13
      Top             =   135
      Width           =   600
   End
End
Attribute VB_Name = "frmSedesTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MshHabilitado As Boolean
Dim FilEli As Long
Dim strAccion As String

Sub BotonEdicion()
    CmdNuevo.Enabled = False
    cmdModificar.Enabled = False
    cmdGrabar.Enabled = True
    CmdCancelar.Enabled = True
    CmdEliminar.Enabled = False
    cmdsalir.Enabled = False
End Sub

Sub BotonNormal()
    Dim I As Integer
    CmdNuevo.Enabled = True
    cmdModificar.Enabled = True
    cmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    CmdEliminar.Enabled = True
    cmdsalir.Enabled = True
    txtCod.Enabled = False
    txtDes.Enabled = False
    txtDirec.Enabled = False
    txtTelef.Enabled = False
    cmbSede.Enabled = False
End Sub

Sub ConfigMshMaestros()
    Dim I As Integer
    With mshSedes
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Cod"
        .ColWidth(0) = 400
        .TextMatrix(0, 1) = "Descrip"
        .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "Direcc"
        .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "Telef"
        .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Tipo"
        .ColWidth(4) = 1000
        .FixedCols = 1
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeBoth
    End With
End Sub
 
Sub BloqueoDeBotones()
    CmdNuevo.Enabled = True
    cmdModificar.Enabled = False
    CmdEliminar.Enabled = False
    cmdGrabar.Enabled = False
    CmdCancelar.Enabled = False
    CmdVistaPreliminar.Enabled = False
End Sub

Sub ModoEdicion()
    txtDes.Enabled = True
    txtDes.Locked = False
    txtDes.BackColor = ColorHabilitado
    txtDirec.Enabled = True
    txtDirec.Locked = False
    txtDirec.BackColor = ColorHabilitado
    txtTelef.Enabled = True
    txtTelef.Locked = False
    txtTelef.BackColor = ColorHabilitado
    cmbSede.Enabled = True
    cmbSede.Locked = False
    cmbSede.BackColor = ColorHabilitado
    MshHabilitado = False
    mshSedes.BackColor = ColorHabilitado
End Sub

Sub ModoNormal()
    Dim I, J As Integer
    txtCod.Locked = True
    txtCod.BackColor = ColorDeshabilitado
    txtDes.Locked = True
    txtDes.BackColor = ColorDeshabilitado
    txtDirec.Locked = True
    txtDirec.BackColor = ColorDeshabilitado
    txtTelef.Locked = True
    txtTelef.BackColor = ColorDeshabilitado
    cmbSede.Locked = True
    cmbSede.BackColor = ColorDeshabilitado
    MshHabilitado = True
    mshSedes.BackColor = ColorHabilitado
End Sub

Private Sub cmbSede_KeyPress(KeyAscii As MSForms.ReturnInteger)
    cmdGrabar.SetFocus
End Sub

Private Sub cmdCancelar_Click()
    strAccion = "Cancelar"
    ModoNormal
    BotonNormal
    Limpia_Valores
    DesplazarPorLaGrilla
End Sub

Private Sub Limpia_Valores()
    Dim I As Integer
    txtCod.Text = Empty
    txtDes.Text = Empty
    txtDirec.Text = Empty
    txtTelef.Text = Empty
    cmbSede.ListIndex = 0
End Sub

Private Sub cmdEliminar_Click()
    strAccion = "Eliminar"
    If txtCod.Text <> Empty Then
        If MsgBox("Está seguro de eliminar el registro con código N° " & txtCod & "  (S/N)", vbInformation + vbYesNo, m_Titulo) = vbYes Then
            FilEli = mshSedes.row
            BorrarRegistro
            ConfigMshMaestros
            LlenarMshMaestros
            Limpia_Valores
            DesplazarPorLaGrilla
            ModoNormal
            BotonNormal
            If FilEli > 1 Then
                mshSedes.row = FilEli - 1
                mshSedes.SetFocus
                Call keybd_event(vbKeyHome, 0, 0, 0)
            Else
                mshSedes.row = 1
            End If
            lblMensaje = Empty
        End If
    Else
        MsgBox "Seleccione El Registro A Eliminar", vbInformation, "NOVPeru"
        mshSedes.SetFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    If ValidarData = True Then
        Me.MousePointer = vbHourglass
        GrabarData strAccion
        mshSedes.Clear
        ConfigMshMaestros
        LlenarMshMaestros
        DoEvents
        Limpia_Valores
        ModoNormal
        BotonNormal
        mshSedes.row = FilSel
        DesplazarPorLaGrilla
        Me.MousePointer = vbNormal
        mshSedes.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
    Else
       MsgBox "No se puede grabar la sede, confirme los datos e inténtelo denuevo", vbOKOnly + vbInformation, "NOVPeru"
    End If
End Sub

Private Sub cmdModificar_Click()
    ModoEdicion
    strAccion = "Modificar"
    BotonEdicion
End Sub

Private Sub cmdNuevo_Click()
    strAccion = "Nuevo"
    ModoEdicion
    Limpia_Valores
    BotonEdicion
    txtCod = GenerarCodigo
    txtDes.SetFocus
    Call keybd_event(vbKeyHome, 0, 0, 0)
End Sub

Private Function GenerarCodigo() As String
    Dim SQL As String
    Dim sql2 As String
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    SQL = "select MAX(codigo) from rh_estacionestrabajo "
    sql2 = "select codigo from rh_estacionestrabajo LIMIT 1"
    Dim I As Integer, aux As String, Longitud As Integer
    Set Rs = oConexion.EjecutaSelectRS(sql2)
    Longitud = Rs.Fields(0).DefinedSize
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    aux = ""
    For I = 1 To Longitud
        aux = aux & "0"
    Next I
    If IsNull(Rs.Fields(0)) Then
        GenerarCodigo = Right(aux & "1", Longitud)
    Else
        GenerarCodigo = Right(aux & CStr(Rs.Fields(0) + 1), Longitud)
    End If
    Set Rs = Nothing
End Function

Private Sub cmdSalir_Click()
    Unload Me
    mdiInicio.Enabled = True
End Sub

Private Sub Form_Activate()
    mshSedes.SetFocus
    Call keybd_event(vbKeyHome, 0, 0, 0)
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    ConfigMshMaestros
    LlenarMshMaestros
    cmbSede.AddItem "1-OFICINA"
    cmbSede.List(0, 2) = "1"
    cmbSede.AddItem "2-CAMPO"
    cmbSede.List(1, 2) = "2"
    DoEvents
    ModoNormal
    BotonNormal
    DesplazarPorLaGrilla
End Sub

Sub LlenarMshMaestros()
    Dim Pos As Integer, pos2 As Integer
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    Set Rs = oConexion.EjecutaSelectRS("select codigo,nombre,direccion,telefonos ,tipo from rh_estacionestrabajo order by codigo")
    Dim I As Integer, J As Integer
    With mshSedes
        .Redraw = False
        If Not (Rs.BOF And Rs.EOF) Then
            For I = 0 To Rs.RecordCount - 1
                For J = 0 To 4
                    .TextMatrix(.Rows - 1, J) = IIf(IsNull(Rs.Fields(J)), "", " " & Rs.Fields(J))
                Next J
                .Rows = .Rows + 1
                Rs.MoveNext
            Next
            .Rows = .Rows - 1
        Else
            BloqueoDeBotones
        End If
        .Redraw = True
    End With
    If Rs.State = adStateOpen Then Rs.CloseRecordset
    Set Rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lblMensaje <> Empty Then
        Dim M As String
        M = MsgBox("¿ Desea Guardar los Cambios Realizados ?", vbInformation + vbYesNoCancel, Caption)
        Select Case M
            Case vbYes
                If ValidarData = True Then
                    GrabarData strAccion
                    Cancel = False
                Else
                    Cancel = True
                End If
            Case vbNo
                Cancel = False
                mdiInicio.Enabled = True
            Case vbCancel
                Cancel = True
        End Select
    Else
        Cancel = False
    End If
End Sub

Sub GrabarData(accion As String)
On Error GoTo ErrSave
    Dim SQL As String
    Dim I As Integer
    Select Case accion
        Case "Nuevo"
            SQL = "Insert into rh_estacionestrabajo (codigo,nombre,direccion,telefonos,tipo) values (" & _
                  "'" & txtCod & "','" & txtDes & "','" & txtDirec & "','" & txtTelef & "','" & cmbSede.List(cmbSede.ListIndex, 2) & "')"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
        Case "Modificar"
            SQL = "Update rh_estacionestrabajo set nombre='" & txtDes & "',direccion='" & _
                   txtDirec & "',telefonos='" & txtTelef & "',tipo='" & cmbSede.List(cmbSede.ListIndex, 2) & _
                   "' where codigo='" & txtCod & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Case "Eliminar"
            SQL = "Delete from rh_estacionestrabajo where codigo='" & txtCod & "'"
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
    End Select
Exit Sub
ErrSave:
    MsgBox "Ha ocurrido un error al momento de grabar" & Chr(13) & err.Description, vbCritical, "Error de datos"
    oConexion.DeshacerTransaccion
End Sub

Sub BorrarRegistro()
On Error GoTo Errdel
    oConexion.EjecutaInsertUpdateDelete "Delete from rh_estacionestrabajo where codigo='" & txtCod & "'", TIPO_QUERY.Eliminar, True
Exit Sub
Errdel:
    MsgBox "Ha ocurrido un error al momento de eliminar" & Chr(13) & err.Description, vbCritical, "Error de datos"
End Sub

Public Function ValidarData() As Boolean
    If txtDes = Empty Then
        MsgBox "El Item Descripción está en blanco", vbInformation, "NOVPeru"
        ValidarData = False
        txtDes.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
        Exit Function
    End If
    If txtDirec = Empty Then
        MsgBox "El Item Dirección está en blanco", vbInformation, "NOVPeru"
        ValidarData = False
        txtDirec.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
        Exit Function
    End If
    If txtTelef = Empty Then
        MsgBox "El Item Telefonos está en blanco", vbInformation, "NOVPeru"
        ValidarData = False
        txtTelef.SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
        Exit Function
    End If
    ValidarData = True
End Function

Private Sub mshSedes_Click()
    DesplazarPorLaGrilla
End Sub

Private Sub mshSedes_DblClick()
    cmdModificar_Click
End Sub

Private Sub mshSedes_KeyDown(KeyCode As Integer, Shift As Integer)
    DesplazarPorLaGrilla
End Sub

Sub DesplazarPorLaGrilla()
    Dim I As Integer
    If MshHabilitado = True Then
        With mshSedes
            txtCod.Locked = True
            txtCod.BackColor = ColorDeshabilitado
            txtCod.Text = Trim(.TextMatrix(.Rowsel, 0))
            txtDes.Locked = True
            txtDes.BackColor = ColorDeshabilitado
            txtDes.Text = Trim(.TextMatrix(.Rowsel, 1))
            txtDirec.Locked = True
            txtDirec.BackColor = ColorDeshabilitado
            txtDirec.Text = Trim(.TextMatrix(.Rowsel, 2))
            txtTelef.Locked = True
            txtTelef.BackColor = ColorDeshabilitado
            txtTelef.Text = Trim(.TextMatrix(.Rowsel, 3))
            For I = 0 To cmbSede.ListCount - 1
                If cmbSede.List(I, 2) = Trim(.TextMatrix(.Rowsel, 4)) Then
                    cmbSede.ListIndex = I
                    Exit For
                End If
            Next I
        End With
    End If
End Sub
    
Private Sub txtDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDes.Text = Empty) Then
            MsgBox "El nombre no puede quedar en blanco", vbInformation + vbOKOnly, "NOVPeru"
            txtDes.SetFocus
        Else
            txtDirec.SetFocus
        End If
    End If
End Sub

Private Sub txtDirec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDirec.Text = Empty) Then
            MsgBox "El nombre no puede quedar en blanco", vbInformation + vbOKOnly, "NOVPeru"
            txtDirec.SetFocus
        Else
            txtTelef.SetFocus
        End If
    End If
End Sub

Private Sub txtTelef_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        If Trim(txtTelef.Text = Empty) Then
            MsgBox "El nombre no puede quedar en blanco", vbInformation + vbOKOnly, "NOVPeru"
            txtTelef.SetFocus
        Else
            cmbSede.SetFocus
        End If
    End If
End Sub
