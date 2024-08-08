VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form FrmPrac 
   Caption         =   "Mantenimiento Terceros"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7530
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHArea 
         Height          =   2445
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4313
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin Proyecto1.chameleonButton cmdNuevo 
      Height          =   405
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1050
      _ExtentX        =   1852
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
      MICON           =   "FrmPrac.frx":0000
      PICN            =   "FrmPrac.frx":001C
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
      Left            =   1320
      TabIndex        =   7
      Top             =   4320
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "FrmPrac.frx":0386
      PICN            =   "FrmPrac.frx":03A2
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
      Left            =   2520
      TabIndex        =   8
      Top             =   4320
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "FrmPrac.frx":07D0
      PICN            =   "FrmPrac.frx":07EC
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
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      Top             =   4320
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "FrmPrac.frx":0C2E
      PICN            =   "FrmPrac.frx":0C4A
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
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   4320
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "FrmPrac.frx":118C
      PICN            =   "FrmPrac.frx":11A8
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
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   4320
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "FrmPrac.frx":15EA
      PICN            =   "FrmPrac.frx":1606
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
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4320
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   ""
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
      MICON           =   "FrmPrac.frx":19CC
      PICN            =   "FrmPrac.frx":19E8
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
      Height          =   1230
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7650
      Begin VB.TextBox TxtCodAux 
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtDivision 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtCodArea 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   3
         Top             =   360
         Width           =   4755
      End
      Begin MSForms.ComboBox cboDoc 
         Height          =   285
         Left            =   6240
         TabIndex        =   23
         Top             =   720
         Width           =   1035
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "1826;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   5760
         TabIndex        =   22
         Top             =   720
         Width           =   390
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "División"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   720
         Width           =   675
      End
      Begin VB.Label LblEstFin 
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   660
      End
      Begin VB.Label LblDescripcion 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6675
      Top             =   885
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   7410
      Begin MSForms.ComboBox cboDocBsq 
         Height          =   285
         Left            =   6000
         TabIndex        =   24
         Top             =   600
         Width           =   1155
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2037;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label LblDescripcion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   540
   End
   Begin VB.Label LblMensaje 
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
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   4905
      TabIndex        =   15
      Top             =   3375
      Width           =   1305
   End
   Begin VB.Label Lbl1 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   270
      TabIndex        =   14
      Top             =   3375
      Width           =   5490
   End
End
Attribute VB_Name = "FrmPrac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CantRegistros As String
Public MshHabilitado As Boolean
Public FilaSel As Integer
Public FilaEli As Integer

Private Sub cboDocBsq_Change()
 LlenarMshArea
 ModoNormal
 BotonNormal
End Sub

Private Sub cmdCancelar_Click()
    ModoNormal
    BotonNormal
    Limpia_Valores
    LblMensaje = Empty
    FilaSel = 0
End Sub

Private Sub CmdEliminar_Click()
  If MsgBox("Está Seguro De Eliminar El Registro ?", vbExclamation + vbYesNo, "Eliminar Registro") = vbYes Then
        FilaEli = FilaSel
        EliminarDatos MSHArea.TextMatrix(FilaEli, 2)
        MSHArea.Clear
        LlenarMshArea
        ModoNormal
        BotonNormal
  End If
End Sub


Sub EliminarDatos(CodAux As String)
On Error GoTo CtrlError
Dim SQL As String

    SQL = "call cn_Delete_Terceros('" & CodAux & "')"
    oConexionMYSQL.BeginTrans
    oConexionMYSQL.Execute (SQL)
    oConexionMYSQL.CommitTrans
Exit Sub
CtrlError:
    ADOConexion.RollbackTrans
    MsgBox err.Description, vbCritical, "Error al Eliminar un Registro"
End Sub

Private Sub cmdGrabar_Click()
  If ValidarData(MSHArea) = True Then
        GrabarData
        MSHArea.Clear
        LlenarMshArea
        ModoNormal
        BotonNormal
        LblMensaje = Empty
        Limpia_Valores
    End If
End Sub

Private Sub CmdModificar_Click()
    ModoEdicion
    BotonEdicion
    LblMensaje = "Modificar"
    TxtCodArea.SetFocus
End Sub

Sub BloqueoEspecial2()
    If FilaSel > 0 Then
        TxtCodArea.Locked = True
        TxtCodArea.BackColor = ColorDeshabilitado
        TxtDescripcion.SetFocus
    Else
        TxtCodArea.SetFocus
    End If
End Sub

Private Sub cmdNuevo_Click()
    Limpia_Valores
    ModoEdicion
    BotonEdicion
    LblMensaje = "Nuevo"
    TxtCodArea.SetFocus
    Call keybd_event(vbKeyHome, 0, 0, 0)
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
LlenarCombo
 LlenarComboBsq
 LlenarMshArea
 ModoNormal
 BotonNormal
End Sub

Sub BloqueoDeBotones()
    cmdNuevo.Enabled = True
    CmdModificar.Enabled = False
    CmdEliminar.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    CmdVistaPreliminar.Enabled = False
End Sub


Sub LlenarMshArea()
Dim SQL As String
Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

    Dim i As Integer
    SQL = "Call CONT_Rep_Proc_Genericos('cntpract','" & cboDocBsq.Value & "','','','','','','','','','');"
    Set Rs = ADO_LlenaRs(SQL)
    ConfigMshArea
    With MSHArea
        If Rs Is Nothing = False Then
            .Redraw = False
            Do While Not Rs.EOF
                .TextMatrix(.Rows - 1, 1) = Rs.Fields("auxiliar")
                .TextMatrix(.Rows - 1, 2) = Rs.Fields("codaux")
                .TextMatrix(.Rows - 1, 3) = Rs.Fields("concepto")
                .TextMatrix(.Rows - 1, 4) = Rs.Fields("cargos")
                .TextMatrix(.Rows - 1, 5) = Rs.Fields("divi")
                .TextMatrix(.Rows - 1, 6) = Rs.Fields("funcion")
                .Rows = .Rows + 1
                Rs.MoveNext
            Loop
            .Rows = .Rows - 1
            Rs.Close
        Else
            BloqueoDeBotones
        End If
        .Redraw = True
    End With
    
    Set Rs = Nothing
End Sub


Sub ConfigMshArea()
    With MSHArea
        .Cols = 7
        .Rows = 1
        .TextMatrix(0, 1) = "Auxiliar"
        .TextMatrix(0, 2) = "CodAux"
        .TextMatrix(0, 3) = "Concepto"
        .TextMatrix(0, 4) = "Monto"
        .TextMatrix(0, 5) = "División"
        .TextMatrix(0, 6) = "Func"
        .ColWidth(0) = 250
        .ColWidth(1) = 700
        .ColWidth(2) = 1200
        .ColWidth(3) = 2500
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 600
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    End With
End Sub

Sub Limpia_Valores()
    TxtCodArea = Empty
    TxtDescripcion = Empty
End Sub

Sub ModoNormal()
    TxtCodArea.Locked = True
    TxtCodArea.BackColor = ColorDeshabilitado
    TxtDescripcion.Locked = True
    TxtDescripcion.BackColor = ColorDeshabilitado
    TxtDivision.Locked = True
    TxtDivision.BackColor = ColorDeshabilitado
    MshHabilitado = True
    MSHArea.BackColor = ColorHabilitado
End Sub

Sub ModoEdicion()
    TxtCodArea.Locked = False
    TxtCodArea.BackColor = ColorHabilitado
    TxtDescripcion.Locked = False
    TxtDescripcion.BackColor = ColorHabilitado
    TxtDivision.Locked = False
    TxtDivision.BackColor = ColorHabilitado
    MshHabilitado = False
    MSHArea.BackColor = ColorDeshabilitado
End Sub

Sub BotonNormal()
    cmdNuevo.Enabled = True
    CmdModificar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    CmdEliminar.Enabled = True
    cmdSalir.Enabled = True
End Sub

Sub BotonEdicion()
    cmdNuevo.Enabled = False
    CmdModificar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    CmdEliminar.Enabled = False
    cmdSalir.Enabled = False
End Sub


Sub GrabarData()
On Error GoTo ErrSave
Dim SQL As String
    Select Case LblMensaje.Caption
        Case "Nuevo"
            SQL = "call cn_Insert_terceros ('" & Trim(TxtCodAux) & "','" & Trim(TxtCodArea) & "','" & Trim(TxtDescripcion) & "','" & Trim(TxtDivision) & "','" & cboDoc.Value & "')"
        Case "Modificar"
            SQL = "call cn_Update_terceros ('" & Trim(TxtCodAux) & "','" & Trim(TxtCodArea) & "','" & Trim(TxtDescripcion) & "','" & Trim(TxtDivision) & "','" & cboDoc.Value & "')"
    End Select
    oConexionMYSQL.BeginTrans
    oConexionMYSQL.Execute (SQL)
    
    oConexionMYSQL.CommitTrans
Exit Sub
ErrSave:
    MsgBox "Ha ocurrido un error al momento de grabar" & Chr(13) & err.Description, vbCritical, "Error de datos"
    ADOConexion.RollbackTrans
End Sub

Public Function ValidarData(ByRef MSHArea As MSHFlexGrid) As Boolean
    If Me.TxtCodArea = Empty Then
        MsgBox "El Item Año Mes Está en Blanco", vbExclamation, Caption
        ValidarData = False
        Exit Function
    Else
        Dim fila As Integer
        fila = BuscarRegistro(MSHArea)
        If fila > 0 Then
           If Me.LblMensaje = "Nuevo" Then
                MsgBox "El Registro ya existe", vbExclamation, Caption
                
                Call keybd_event(vbKeyHome, 0, 0, 0)
                ValidarData = False
                Exit Function
           End If
        End If
    End If
    If Me.TxtDescripcion = Empty Then
        MsgBox "El Item Tipo de Cambio OAnda en Blanco", vbExclamation, Caption
        ValidarData = False
        Exit Function
    End If
    ValidarData = True
End Function


Public Function BuscarRegistro(ByRef MSHArea As MSHFlexGrid) As Integer
    Dim i As Integer
    For i = 1 To MSHArea.Rows - 1
        If MSHArea.TextMatrix(i, 1) = Trim(Me.TxtCodArea) Then
            BuscarRegistro = i
            Exit Function
        End If
    Next
    BuscarRegistro = 0
End Function

Private Sub MSHArea_Click()
  NavegarPorGrilla
End Sub


Private Sub MSHArea_DblClick()
 Call CmdModificar_Click
End Sub

Private Sub MSHArea_KeyDown(KeyCode As Integer, Shift As Integer)
 NavegarPorGrilla
End Sub

Sub NavegarPorGrilla()
    If MshHabilitado = True Then
        With MSHArea
            FilaSel = .Rowsel
            TxtCodAux = .TextMatrix(.Rowsel, 2)
            TxtCodArea = .TextMatrix(.Rowsel, 3)
            TxtDescripcion = .TextMatrix(.Rowsel, 4)
            TxtDivision = .TextMatrix(.Rowsel, 5)
            cboDoc.Value = .TextMatrix(.Rowsel, 6)
        End With
    End If
End Sub

Private Sub TxtCodArea_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 112 Then ' F1
'     oConsulta.Caso = 3
'     oConsulta.CodigoAuxil = 6
'
'     If oConsulta.Caso <> 100 Then
'        oConsulta.Formulario = "FrmPRpt"
'
'        'FrmConsultas.Parametro = 0
'        'FrmConsultas.Caption = Me.Caption
'        FrmConsultas.Show
'     End If
'
'  End If
End Sub


Sub LlenarCombo()
 cboDoc.AddItem "-"
 cboDoc.List(0, 2) = "-"
 cboDoc.AddItem "P"
 cboDoc.List(1, 2) = "P"
 cboDoc.AddItem "R"
 cboDoc.List(2, 2) = "R"
 cboDoc.AddItem "S"
 cboDoc.List(3, 2) = "S"
 
 cboDoc.ListIndex = 0
End Sub

Sub LlenarComboBsq()
 cboDocBsq.AddItem "-"
 cboDocBsq.List(0, 2) = "-"
 cboDocBsq.AddItem "P"
 cboDocBsq.List(1, 2) = "P"
 cboDocBsq.AddItem "R"
 cboDocBsq.List(2, 2) = "R"
 cboDocBsq.AddItem "S"
 cboDocBsq.List(3, 2) = "S"
 
 cboDocBsq.ListIndex = 0
End Sub


