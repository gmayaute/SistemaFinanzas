VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmModificarEstado 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar de Estado"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   Icon            =   "frmModificarEstado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11325
   Begin VB.TextBox txtNumIdent 
      Height          =   315
      Left            =   8220
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   540
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   3705
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   11325
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDetalleMovDoc 
         Height          =   2835
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   5001
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Proyecto1.chameleonButton ChBtnSalir 
         Height          =   345
         Left            =   10710
         TabIndex        =   9
         Top             =   3150
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
         MICON           =   "frmModificarEstado.frx":0442
         PICN            =   "frmModificarEstado.frx":045E
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
         Height          =   345
         Left            =   10110
         TabIndex        =   10
         Top             =   3150
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
         MICON           =   "frmModificarEstado.frx":0824
         PICN            =   "frmModificarEstado.frx":0840
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnEstado 
         Height          =   375
         Left            =   8400
         TabIndex        =   8
         ToolTipText     =   "Modificar"
         Top             =   3150
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
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
         MICON           =   "frmModificarEstado.frx":0D82
         PICN            =   "frmModificarEstado.frx":0D9E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   3180
         Width           =   4155
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   60
         TabIndex        =   20
         Top             =   3120
         Width           =   5085
      End
      Begin MSForms.ComboBox cmbEstado 
         Height          =   315
         Left            =   5310
         TabIndex        =   18
         Top             =   3180
         Width           =   2985
         VariousPropertyBits=   746604569
         BackColor       =   14737632
         ForeColor       =   128
         DisplayStyle    =   7
         Size            =   "5265;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11325
      Begin VB.TextBox txtNumDoc 
         Height          =   315
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   3015
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox txtDescTipoDoc 
         BackColor       =   &H00C0C0C0&
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Descripcion del Tipo de Documento"
         Top             =   180
         Width           =   4995
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   90
         TabIndex        =   22
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label lblDescDivis 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion del la Division"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   1260
         Width           =   9555
      End
      Begin VB.Label lblDescCenco 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion del Centro de Costo"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   1650
         Width           =   9555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centro de Costo:"
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
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Folio:"
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
         Height          =   315
         Left            =   6750
         TabIndex        =   14
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label lblDescAux 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion de Auxiliar"
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
         Left            =   90
         TabIndex        =   13
         Top             =   900
         Width           =   11145
      End
      Begin VB.Label lblAuxiliar 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auxiliar:"
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
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   6585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Documento:"
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
         Height          =   315
         Left            =   6720
         TabIndex        =   5
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo:"
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
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Label lblIdentificador 
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   6210
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "frmModificarEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New MYSQL_RS
Dim SQL As String
Dim flagcombo As Boolean
Dim familia As FAMILIA_DOC
Dim bandera As Boolean

Private Sub LLenarCabecerasGrilla()
With flxDetalleMovDoc
    .Clear
    .Rows = 1
    .Cols = 10
    .BackColor = &H8000000F
    
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(0) + "N°Mov."
        
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(4) + "Folio"
          
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = Space(4) + "CodDestino"
        
        .ColWidth(3) = 1800
        .TextMatrix(0, 3) = Space(4) + "Area"
        
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = Space(3) + "Usuario"
        
        .ColWidth(5) = 1500
        .TextMatrix(0, 5) = Space(1) + "Fecha Modific."
        
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = Space(3) + "CodEstado"
               
        .ColWidth(7) = 1300
        .TextMatrix(0, 7) = Space(3) + "Estado"
                
        .ColWidth(8) = 2400
        .TextMatrix(0, 8) = Space(4) + "Observación"
        
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = Space(4) + "anomes"
       
 End With
End Sub

Public Sub LLenarDatosGrilla(ByVal identificador As String)
    Dim I As Integer
        SQL = "asignar_estados where identificador = '" & identificador & "' "
        Set Rs = oConexion.EjecutaSelect(SQL)
            With flxDetalleMovDoc
                Do While Not (Rs.EOF)
                .Rows = .Rows + 1
                .FixedRows = 1
                .BackColor = vbWhite
                For I = 1 To .Cols - 1
                        .TextMatrix(.Rows - 1, I) = Rs.Fields(I)
                Next I
                Rs.MoveNext
                Loop
                EnumerarItems1 flxDetalleMovDoc
          End With
        Set Rs = Nothing
        flagcombo = True
        CargaCbo cmbEstado
        bandera = True
End Sub

Private Sub btnEstado_Click()
    Dim RES As Integer
    RES = MsgBox("¿Esta Seguro que desea poner en " & btnEstado.Caption & " el Documento", vbQuestion + vbYesNo, gsNomSW)
    If RES = 6 Then
        With flxDetalleMovDoc
            CambioEstado .TextMatrix(.Rows - 1, 1), strEstado
        End With
        Else
            Exit Sub
    End If
    ModoUpdate
End Sub

Private Sub chBtnSalir_Click()
    Unload Me
End Sub

Private Sub cmbEstado_Change()
    Dim folio As String
    If Len(txtNumIdent) < 5 Then
        folio = strAnoSistema & strMesSistema & txtNumIdent
    Else
        folio = txtNumIdent
    End If
    If bandera Then
      strEstado = cmbEstado.List(cmbEstado.ListIndex, 1)
      With flxDetalleMovDoc
        CambioEstado folio, strEstado '.TextMatrix(.Rows - 1, 1), strEstado
      End With
      ModoUpdate
      flagcombo = False
    End If
End Sub

Private Sub CmdVistaPreliminar_Click()
Set oReporte = New clsReporte
oReporte.empresa = strNombreEmpresa
oReporte.Titulo = "REPORTE DE MODIFICACION DE ESTADOS DE " & txtDescTipoDoc.Text
familia = DescripcionesdeCodigos("CNDOCUM", txtTipoDoc.Text, "2")
    Select Case familia
           Case FAMILIA_DOC.CONTABLES
                oReporte.Reporte = "Rep_Modificar_Doc.rpt"
                oReporte.sp_Rep_ModificarDoc "documento_contables", strIdentificador
           Case FAMILIA_DOC.ENTIDADES
                oReporte.Reporte = "Rep_Modificar_Doc.rpt"
                oReporte.sp_Rep_ModificarDoc "documento_entidades", strIdentificador
           Case FAMILIA_DOC.GENERALES
                oReporte.Reporte = "Rep_Modificar_Doc.rpt"
                oReporte.sp_Rep_ModificarDoc "documento_generales", strIdentificador
           Case FAMILIA_DOC.ORDENES
                oReporte.Reporte = "Rep_Modificar_Doc.rpt"
                oReporte.sp_Rep_ModificarDoc "orden_compra", strIdentificador
    End Select
End Sub

Private Sub flxDetalleMovDoc_RowColChange()
 If flxDetalleMovDoc.row > 0 And flxDetalleMovDoc.Col > 0 Then DesplazarenFlex lblMsjgrilla, flxDetalleMovDoc
End Sub

Private Sub Form_Load()
    Call WheelHook(frmModificarEstado)
    LLenarCabecerasGrilla
    Me.Top = 0
    Me.Left = 0
    flagcombo = False
    CargaCbo cmbEstado
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Unload Me
End Sub

Private Sub txtNumDoc_GotFocus()
    mark txtNumDoc
End Sub

Private Sub txtTipoDoc_GotFocus()
    mark txtTipoDoc
End Sub

Public Sub EstadoBoton(ByVal TipoDoc As String)
    Dim Xciclo As String
    CiclodeVidaDoc (TipoDoc)
    If ciclo(flxDetalleMovDoc.Rows, 1) <> " " Then
        If flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.Rows - 1, 6) = REGISTRADO Or _
           flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.Rows - 1, 6) = PROGRAMADO Or _
           flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.Rows - 1, 6) = REVISADO Or _
           flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.Rows - 1, 6) = APROBADO Or _
           flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.Rows - 1, 6) = CANCELADO Then
            Xciclo = flxDetalleMovDoc.row + 1
            Else
            Xciclo = BuscaAnt
        End If
        CaptionBoton Xciclo
        Else
    End If

End Sub

Public Function BuscaAnt() As String
    Dim I As Integer
    With flxDetalleMovDoc
        For I = .Rows - 1 To 1 Step -1
            If .TextMatrix(I, 6) Like REGISTRADO Or _
               .TextMatrix(I, 6) Like PROGRAMADO Or _
               .TextMatrix(I, 6) Like REVISADO Or _
               .TextMatrix(I, 6) Like APROBADO Or _
               .TextMatrix(I, 6) Like CANCELADO Then
                BuscaAnt = I + 1
                Exit Function
            End If
        Next
    End With
End Function

Public Function CaptionBoton(ByVal v As String)
    On Error GoTo ciclo
    Dim J As Integer
    J = Len(DescripcionesdeCodigos("DOC_ESTADO", ciclo(v, 1)))
    btnEstado.Caption = Left(DescripcionesdeCodigos("DOC_ESTADO", ciclo(v, 1)), 1) & LCase(Right(DescripcionesdeCodigos("DOC_ESTADO", ciclo(v, 1)), J - 1))
    btnEstado.ToolTipText = btnEstado.Caption
    strEstado = ciclo(v, 1)
    If ciclo(v, 2) = "1" Then
        btnEstado.Enabled = True
    Else
        btnEstado.Enabled = False
    End If
    If ciclo(v, 1) = CANCELADO Then
        btnEstado.Enabled = False
    Else
    End If
    Exit Function
ciclo:
   Exit Function
End Function

Private Function CargaCbo(combo As Control)
    Dim rscbo As New MYSQL_RS
    bandera = False
    combo.Clear
    combo.AddItem "Seleccionar..."
    combo.List(0, 1) = "00"
    If flagcombo = True Then
        SQL = "carga_estausu where Usuario_id = '" & strUsuarioId & "' and Permiso = '1'" & _
              " and cod_estado <> '" & PROGRAMADO & "'" & _
              " and cod_estado <> '" & REVISADO & "' and cod_estado <> '" & APROBADO & "' and cod_estado <> '" & TRANSFERIDO & "'" & _
              " and cod_estado <> '" & EMITIDO & "' and cod_estado <> '" & MODIFICADO & "'" & _
              " and cod_estado <> '" & flxDetalleMovDoc.TextMatrix(flxDetalleMovDoc.row, 6) & "'"
        Set rscbo = oConexion.EjecutaSelect(SQL)
        If Not rscbo.EOF And rscbo.RecordCount <> 0 Then
            combo.Enabled = True
            combo.BackColor = &H80000005
            combo.Value = "Seleccionar..."
            For I = 1 To rscbo.RecordCount
                combo.AddItem rscbo.Fields("descripcion")
                combo.List(I, 1) = rscbo.Fields("cod_estado")
                rscbo.MoveNext
            Next
            Else
            combo.Enabled = False
        End If
    End If
End Function

Private Sub ModoUpdate()
    flxDetalleMovDoc.Clear
    LLenarCabecerasGrilla
    frmBusquedaDocumentaria.flxResultados_DblClick
    frmBusquedaDocumentaria.TxtTipo = txtTipoDoc
    frmBusquedaDocumentaria.EjecutaBusqueda
    frmBusquedaDocumentaria.TxtTipo_KeyDown 116, 0
    Unload Me
    frmBusquedaDocumentaria.SetFocus
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
    On Error Resume Next
    
    With flxDetalleMovDoc
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 0
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
