VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmNivel 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Perfiles"
   ClientHeight    =   3315
   ClientLeft      =   2820
   ClientTop       =   3180
   ClientWidth     =   5160
   Icon            =   "frmNivel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5160
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   870
      TabIndex        =   15
      ToolTipText     =   "Editar"
      Top             =   2895
      Width           =   885
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   -15
      TabIndex        =   14
      ToolTipText     =   "Nuevo"
      Top             =   2895
      Width           =   870
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   1770
      TabIndex        =   13
      ToolTipText     =   "Borrar"
      Top             =   2895
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   945
      Left            =   0
      TabIndex        =   7
      Top             =   -90
      Width           =   5130
      Begin VB.TextBox TxtNivel 
         Height          =   285
         Left            =   1215
         TabIndex        =   9
         Top             =   195
         Width           =   720
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   1215
         TabIndex        =   8
         Top             =   510
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Nivel"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   795
      Width           =   5130
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHNivel 
         Height          =   1545
         Left            =   30
         TabIndex        =   6
         Top             =   135
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   2725
         _Version        =   393216
         FixedCols       =   0
         BackColorBkg    =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label LblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMensaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   1710
         Width           =   5010
      End
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   375
      Left            =   4695
      Picture         =   "frmNivel.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   2895
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton cmdImprimir 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3690
      Picture         =   "frmNivel.frx":059C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir"
      Top             =   2895
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton CmdVistaPreliminar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4125
      Picture         =   "frmNivel.frx":069E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Preliminar"
      Top             =   2895
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton cmdGrabar 
      Height          =   375
      Left            =   3180
      Picture         =   "frmNivel.frx":0BD0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Grabar"
      Top             =   2895
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   375
      Left            =   2805
      Picture         =   "frmNivel.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Deshacer"
      Top             =   2895
      UseMaskColor    =   -1  'True
      Width           =   375
   End
End
Attribute VB_Name = "FrmNivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CantRegistros As Integer
Public MshHabilitado As Boolean

Sub BloqueoDeBotones()
  cmdNuevo.Enabled = True
  CmdModificar.Enabled = False
  CmdEliminar.Enabled = False
  cmdGrabar.Enabled = False
  cmdCancelar.Enabled = False
  cmdImprimir.Enabled = False
  CmdVistaPreliminar.Enabled = False
  cmdSalir.Enabled = True
End Sub

Sub BorrarRegistro()
On Error GoTo Errdel

  Dim SQL As String
  
  SQL = "DELETE FROM 1cnperfil WHERE perfil_id ='" & Trim(TxtNivel) & "'"
  
  
  ADOConexion.BeginTrans
  ADOConexion.Execute (SQL)
  ADOConexion.CommitTrans
  Exit Sub

Errdel:
    MsgBox "Ha ocurrido un error al momento de eliminar" & Chr(13) & Err.Description, vbCritical, "Error de datos"
    ADOConexion.RollbackTrans
End Sub

Private Sub cmdCancelar_Click()
  Call Limpia_Valores
  Call ModoNormal
  Call BotonNormal
  LblMensaje.Caption = Empty
End Sub

Private Sub CmdEliminar_Click()
  If TxtNivel <> Empty Then
      If MsgBox("Está Seguro De Eliminar El Nivel Con En Código N° " + CStr(TxtNivel.Text) + " (S/N)", vbExclamation + vbYesNo, Caption) = vbYes Then
          Call BorrarRegistro
          MSHNivel.Clear
          Call LlenarMSHNivel
          Call Limpia_Valores
          Call ModoNormal
          Call BotonNormal
          LblMensaje.Caption = Empty
      End If
  Else
      MsgBox "Seleccione El Registro A Eliminar", vbExclamation, Caption
      MSHNivel.SetFocus
  End If
End Sub
Private Sub cmdGrabar_Click()
    If ValidarData = True Then
        MDIPrincipal.MousePointer = vbHourglass
        Call GrabarData
        MSHNivel.Clear
        Call LlenarMSHNivel
        Call Limpia_Valores
        Call ModoNormal
        Call BotonNormal
        LblMensaje.Caption = Empty
        MDIPrincipal.MousePointer = vbNormal
    End If
End Sub
Sub GrabarData()
On Error GoTo ErrSave
  
  Dim SQL As String
  
  Select Case LblMensaje.Caption
    
    Case "Nuevo"
      SQL = "INSERT INTO 1cnperfil (perfil_id,descripcion) VALUES ('" & UCase(Trim(TxtNivel)) & "','" & UCase(Trim(TxtDescripcion)) & "')"
    
    Case "Edicion"
      SQL = "UPDATE 1cnperfil SET descripcion = '" & UCase(Trim(TxtDescripcion)) & "' WHERE perfil_id ='" & UCase(Trim(TxtNivel)) & "'"
  
  End Select
  

  ADOConexion.BeginTrans
  ADOConexion.Execute (SQL)
  ADOConexion.CommitTrans
  Exit Sub

ErrSave:
    MsgBox "Ha ocurrido un error al momento de grabar" & Chr(13) & Err.Description, vbCritical, "Error de datos"
    ADOConexion.RollbackTrans
End Sub
Private Function ValidarData() As Boolean
    If TxtNivel = Empty Then
        MsgBox "El Item Línea Está En Blanco ", vbExclamation, "Líneas de G.P. x Función"
        TxtLinea.SetFocus
        ValidarData = False
        Exit Function
    End If
    If TxtDescripcion = Empty Then
        MsgBox "El Item Descripción Está En Blanco", vbExclamation, Caption
        ValidarData = False
        TxtDescripcion.SetFocus
        SendKeys "{HOME}+{END}"
        Exit Function
    End If
    ValidarData = True
End Function

Private Sub CmdModificar_Click()
  Call ModoEdicion
  Call BotonEdicion
  LblMensaje.Caption = "Edicion"
  TxtDescripcion.SetFocus
  SendKeys "{HOME}+{END}"
End Sub

Private Sub cmdNuevo_Click()
  Call ModoEdicion
  Call BotonEdicion
  Call Limpia_Valores
  LblMensaje.Caption = "Nuevo"
  TxtNivel.SetFocus
End Sub

Private Sub cmdSalir_Click()
  Unload Me
  Call frmSeguridad.LlenarCbos
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Unload Me
End If
End Sub

Private Sub Form_Load()
  Me.Left = 0
  Me.Top = 0
  'MDIPrincipal.Enabled = False
  Call ModoNormal
  Call BotonNormal
  Call LlenarMSHNivel
  LblMensaje = Empty
End Sub

Sub BotonNormal()
  cmdNuevo.Enabled = True
  CmdModificar.Enabled = True
  CmdEliminar.Enabled = True
  cmdGrabar.Enabled = False
  cmdCancelar.Enabled = False
  cmdSalir.Enabled = True
  cmdImprimir.Enabled = True
  CmdVistaPreliminar.Enabled = True
End Sub
Sub BotonEdicion()
    cmdNuevo.Enabled = False
    CmdModificar.Enabled = False
    CmdEliminar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdSalir.Enabled = False
    cmdImprimir.Enabled = False
    CmdVistaPreliminar.Enabled = False
End Sub
Sub ModoNormal()
    TxtNivel.Locked = True
    TxtNivel.BackColor = ColorDeshabilitado
    TxtDescripcion.Locked = True
    TxtDescripcion.BackColor = ColorDeshabilitado
    MshHabilitado = True
    MSHNivel.BackColor = ColorHabilitado
End Sub
Sub ModoEdicion()
    TxtNivel.Locked = False
    TxtNivel.BackColor = ColorHabilitado
    TxtDescripcion.Locked = False
    TxtDescripcion.BackColor = ColorHabilitado
    MshHabilitado = False
    MSHNivel.BackColor = ColorDeshabilitado
End Sub
Sub Limpia_Valores()
    TxtDescripcion = Empty
    TxtNivel = Empty
End Sub
Sub LlenarMSHNivel()

    Dim RS As ADODB.Recordset

    Set RS = New ADODB.Recordset


    Dim SQL, Cant As String
    SQL = "SELECT perfil_id,Descripcion FROM 1cnperfil ORDER BY 1"

    Set RS = ADOConexion.Execute(SQL)

    Call ConfigMSHNivel

    Dim i As Integer
    With MSHNivel
      If Not (RS.BOF And RS.EOF) Then
        Do While Not RS.EOF
            .TextMatrix(.Rows - 1, 1) = RS(0)
            .TextMatrix(.Rows - 1, 2) = RS(1)
            .Rows = .Rows + 1
            RS.MoveNext
        Loop
        .Rows = .Rows - 1
      Else
        Call BloqueoDeBotones
      End If
    End With
    Set RS = Nothing
End Sub
Sub ConfigMSHNivel()
    With MSHNivel
        .Cols = 3
        .Rows = 2
        .ColWidth(0) = 0
        .ColWidth(1) = 600
        .ColWidth(2) = 3000
        .TextMatrix(0, 1) = "Nivel"
        .TextMatrix(0, 2) = "Descripción"
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If LblMensaje <> Empty Then
    Dim M As String
    M = MsgBox("¿ Desea Guardar los Cambios Realizados ?", vbExclamation + vbYesNoCancel, Caption)
    Select Case M
      Case vbYes
        If ValidarData = True Then
          Call GrabarData
          Cancel = 0
          MDIPrincipal.Enabled = True
        Else
          Cancel = 1
        End If
      Case vbNo
        Cancel = 0
        MDIPrincipal.Enabled = True
      Case vbCancel
        Cancel = 1
    End Select
  Else
      MDIPrincipal.Enabled = True
  End If
End Sub

Sub NavegarGrilla()
    'If MshHabilitado = True Then
            With MSHNivel
                TxtNivel.Text = .TextMatrix(.RowSel, 1)
                TxtDescripcion.Text = .TextMatrix(.RowSel, 2)
            End With
    'End If
End Sub

Private Sub MSHNivel_Click()
  Call NavegarGrilla
End Sub

Private Sub MSHNivel_KeyDown(KeyCode As Integer, Shift As Integer)
  Call NavegarGrilla
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And cmdGrabar.Enabled = True Then Call cmdGrabar_Click
End Sub


Private Sub TxtNivel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    TxtDescripcion.SetFocus
  Else
    KeyAscii = (KeyAscii)
  End If
End Sub

