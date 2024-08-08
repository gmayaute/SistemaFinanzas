VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCambioEstadoenMasa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Estados en Bloque"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   Icon            =   "frmCambioEstadoenMasa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   11265
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6345
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   11265
      Begin VB.CheckBox chkSeleccionaTodos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccionar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9120
         TabIndex        =   4
         Top             =   5370
         Width           =   1965
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCambiosenMasa 
         Height          =   4605
         Left            =   90
         TabIndex        =   1
         Top             =   660
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   8123
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Proyecto1.chameleonButton btnGrabar 
         Height          =   345
         Left            =   9060
         TabIndex        =   2
         ToolTipText     =   "Guardar"
         Top             =   5850
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Aceptar"
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
         BCOL            =   14737632
         BCOLO           =   15309923
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCambioEstadoenMasa.frx":0442
         PICN            =   "frmCambioEstadoenMasa.frx":045E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton ChBtnSalir 
         Height          =   345
         Left            =   10530
         TabIndex        =   3
         Top             =   5850
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
         BCOL            =   14737632
         BCOLO           =   15309923
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCambioEstadoenMasa.frx":0CF0
         PICN            =   "frmCambioEstadoenMasa.frx":0D0C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnModificar 
         Height          =   345
         Left            =   7740
         TabIndex        =   5
         ToolTipText     =   "Modificar"
         Top             =   5850
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "Consultar"
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
         BCOL            =   14737632
         BCOLO           =   15309923
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCambioEstadoenMasa.frx":10D2
         PICN            =   "frmCambioEstadoenMasa.frx":10EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblMsjDocs 
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
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   11025
      End
      Begin VB.Label lblMsjgrilla 
         BackColor       =   &H80000008&
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
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   5370
         Width           =   8115
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   5340
         Width           =   8895
      End
   End
End
Attribute VB_Name = "frmCambioEstadoenMasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function LlenarGrilla()
    With flxCambiosenMasa
         .Clear
        .Rows = 1
        .Cols = 18
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        .FixedRows = 0
        
        .ColWidth(1) = 450
        .TextMatrix(0, 1) = Space(0) + "Folio"
        
        .ColWidth(2) = 1400
        .TextMatrix(0, 2) = Space(2) + "N° Documento"
    
        .ColWidth(3) = 400
        .TextMatrix(0, 3) = Space(0) + "Tipo"
    
        .ColWidth(4) = 400
        .TextMatrix(0, 4) = Space(0) + "Prio"
    
        .ColWidth(5) = 500
        .TextMatrix(0, 5) = Space(0) + "Area"
    
        .ColWidth(6) = 1200
        .TextMatrix(0, 6) = Space(8) + "Cenco"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = Space(8) + "Div"
    
        .ColWidth(8) = 400
        .TextMatrix(0, 8) = Space(0) + "Aux"
        
        .ColWidth(9) = 1200
        .TextMatrix(0, 9) = Space(3) + "Cod.Aux"
    
        .ColWidth(10) = 400
        .TextMatrix(0, 10) = Space(0) + "Mon"
        
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = Space(4) + "Importe"
        
        .ColWidth(12) = 1000
        .TextMatrix(0, 12) = Space(4) + "Fec. Emi"
        
        .ColWidth(13) = 1000
        .TextMatrix(0, 13) = Space(4) + "Fec. Pago"
        
        .ColWidth(14) = 0
        .TextMatrix(0, 14) = Space(8) + "Estado"
        
        .ColWidth(15) = 1000
        .TextMatrix(0, 15) = Space(3) + frmBusquedaDocumentaria.btnEstado.Caption
        
        .ColWidth(16) = 0
        .TextMatrix(0, 16) = "Familia"
         
        .ColWidth(17) = 0
        .TextMatrix(0, 17) = "anomes"
         
        For i = 0 To 17
            .row = 0
            .col = i
            .CellForeColor = &H80000002
            .CellBackColor = &H8000000F
        Next i
        
    End With
End Function
Private Function LLenarDatos()
    Dim j, i As Integer
    With frmBusquedaDocumentaria.flxResultados
        flxCambiosenMasa.Rows = flxCambiosenMasa.Rows + 1
        flxCambiosenMasa.FixedRows = 1
        flxCambiosenMasa.BackColor = vbWhite
        flxCambiosenMasa.Rows = .Rows
        For i = 1 To .Rows - 1
            For j = 0 To .Cols - 1
             If j = 15 Then
                flxCambiosenMasa.col = j
                flxCambiosenMasa.row = i
                flxCambiosenMasa.CellFontName = "Wingdings"
                flxCambiosenMasa.CellFontSize = 14
                flxCambiosenMasa.CellAlignment = flexAlignCenterCenter
                flxCambiosenMasa.Text = strUnChecked
                Else
                flxCambiosenMasa.TextMatrix(i, j) = .TextMatrix(i, j)
             End If
            Next j
        Next i
        lblMsjDocs.Caption = .Rows - 1 & " " & "DOCUMENTOS" & " " & .TextMatrix(1, 14) & "S"
    End With
End Function

Private Sub btnGrabar_Click()
    Dim filas As Integer
    filas = flxCambiosenMasa.Rows - 1
    For i = 1 To filas
        If flxCambiosenMasa.TextMatrix(i, 15) = strChecked Then
            CambioEstado flxCambiosenMasa.TextMatrix(i, 17) & flxCambiosenMasa.TextMatrix(i, 1), DescripcionesdeCodigos("ESTADOenCODIGO", Trim(UCase(flxCambiosenMasa.TextMatrix(0, 15))))
        End If
    Next i
    btnGrabar.Enabled = False
    frmBusquedaDocumentaria.cmbEstado.ListIndex = -1
    frmBusquedaDocumentaria.EjecutaBusqueda
    Unload Me
    frmBusquedaDocumentaria.SetFocus
End Sub

Private Sub btnModificar_Click()
    Dim fila As Integer
    With flxCambiosenMasa
        If .row > 0 Then
            fila = .row
            Aceptar .TextMatrix(fila, 17), .TextMatrix(fila, 1), .TextMatrix(fila, 3)
        End If
    End With
End Sub

Private Sub ChBtnSalir_Click()
    Unload Me
End Sub

Private Sub chkSeleccionaTodos_Click()
    Dim i As Integer
    If chkSeleccionaTodos.Value = Checked Then
            For i = 1 To flxCambiosenMasa.Rows - 1
                flxCambiosenMasa.TextMatrix(i, 15) = strChecked
            Next i
        Else
            For i = 1 To flxCambiosenMasa.Rows - 1
                flxCambiosenMasa.TextMatrix(i, 15) = strUnChecked
            Next i
    End If
End Sub

Private Sub flxCambiosenMasa_Click()
    Dim i As Integer
    With flxCambiosenMasa
        If .ColSel = 15 Then
                .TextMatrix(.row, 15) = IIf(.TextMatrix(.row, 15) = strChecked, strUnChecked, strChecked)
        End If
    End With
End Sub

Private Sub flxCambiosenMasa_RowColChange()
    If flxCambiosenMasa.row > 0 And flxCambiosenMasa.col > 0 Then DesplazarenFlex lblMsjgrilla, flxCambiosenMasa
End Sub

Private Sub Form_Load()
    LlenarGrilla
    LLenarDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

