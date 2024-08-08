VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmAnexosTelewiese 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anexar Documentos al Pago"
   ClientHeight    =   5235
   ClientLeft      =   3405
   ClientTop       =   3705
   ClientWidth     =   12705
   Icon            =   "frmTelewise.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flexAnexos 
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   6747
      _Version        =   393216
   End
   Begin Proyecto1.chameleonButton btnAceptar 
      Height          =   345
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Guardar"
      Top             =   4770
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Aceptar"
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
      MICON           =   "frmTelewise.frx":014A
      PICN            =   "frmTelewise.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro. Orden:"
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
      Left            =   30
      TabIndex        =   16
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sr(es):"
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
      Left            =   3090
      TabIndex        =   15
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda:"
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
      Left            =   30
      TabIndex        =   14
      Top             =   450
      Width           =   1275
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total a Pagar:"
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
      Left            =   4350
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha de Pago"
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
      Left            =   3090
      TabIndex        =   12
      Top             =   450
      Width           =   1365
   End
   Begin MSForms.ComboBox cmbMoneda 
      Height          =   315
      Left            =   1380
      TabIndex        =   11
      Top             =   450
      Width           =   1635
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2884;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblCheque 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
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
      Left            =   1380
      TabIndex        =   10
      Top             =   0
      Width           =   1635
   End
   Begin VB.Label lblSrs 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
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
      Left            =   4530
      TabIndex        =   9
      Top             =   0
      Width           =   6615
   End
   Begin MSForms.ComboBox cmbFecProg 
      Height          =   315
      Left            =   8280
      TabIndex        =   8
      Top             =   450
      Width           =   1695
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2990;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Programada"
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
      Left            =   6420
      TabIndex        =   7
      Top             =   450
      Width           =   1695
   End
   Begin VB.Label lblImporte 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Label lblImporteEqu 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8850
      TabIndex        =   5
      Top             =   4800
      Width           =   1725
   End
   Begin VB.Label lblMImp 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
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
      Left            =   5910
      TabIndex        =   4
      Top             =   4800
      Width           =   435
   End
   Begin VB.Label lblMImpEqu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$"
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
      Left            =   8310
      TabIndex        =   3
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblFecPago 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
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
      Left            =   4530
      TabIndex        =   2
      Top             =   450
      Width           =   1635
   End
End
Attribute VB_Name = "frmAnexosTelewiese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LlenarGrilla()
    Dim I As Integer
    With flexAnexos
        .Clear
        .Rows = 1
        .Cols = 15
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .FixedCols = 1
        .ColWidth(1) = 800
        .ColAlignment(1) = 3
        .TextMatrix(0, 1) = Space(1) & "Anomes"
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = Space(0) + "Folio"
        .ColWidth(3) = 400
        .TextMatrix(0, 3) = Space(0) + "TD"
        .ColWidth(4) = 500
        .TextMatrix(0, 4) = Space(0) + "Mon."
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = Space(2) + "Fec. Emi"
        .ColWidth(6) = 1400
        .TextMatrix(0, 6) = Space(4) + "Serie"
        .ColWidth(7) = 1400
        .TextMatrix(0, 7) = Space(4) + "Documento"
        .ColWidth(8) = 1500
        .TextMatrix(0, 8) = Space(4) + "Importe"
        .ColWidth(9) = 1500
        .TextMatrix(0, 9) = Space(3) + "Importe Equ."
        .ColWidth(10) = 1000
        .TextMatrix(0, 10) = Space(2) + "Fec. Vcto"
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = Space(2) + "Fec. Pago"
        .ColWidth(12) = 700
        .TextMatrix(0, 12) = Space(2) + "Anexar"
        .ColWidth(13) = 0
        .TextMatrix(0, 13) = Space(0) + "Ident"
        .ColWidth(14) = 0
        .TextMatrix(0, 14) = Space(0) + "ccHFM"
    End With
End Sub


Private Sub moneda(a As MSForms.ComboBox)
    a.Clear
    a.AddItem "Seleccionar..."
    a.List(0, 1) = "0"
    a.AddItem "Nacional"
    a.List(1, 1) = "N"
    a.AddItem "Extranjera"
    a.List(2, 1) = "E"
    a.ListIndex = 0
    If a.ListCount > 0 Then a.ListIndex = 0
End Sub


Private Sub CargarFechasProg(aux As String, Cod As String)
    Dim rsfechas As MYSQL_RS
    Dim SQL As String
    Dim I As Integer
    SQL = "Select distinct fec_pago from DOC_PROG WHERE  " & _
            " MON ='" & frmMovTelewiese.cmbMoneda.List(frmMovTelewiese.cmbMoneda.ListIndex, 1) & "' AND " & _
            " AUXILIAR ='" & frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1) & "' AND " & _
            " CODIGO='" & frmMovTelewiese.txtCodigo & "' ORDER BY FEC_PAGO"
    Set rsfechas = oConexion.EjecutaSelectRS(SQL)
    cmbFecProg.Clear
    cmbFecProg.AddItem "Seleccionar..."
    If rsfechas.RecordCount > 0 Then
        Do While Not rsfechas.EOF
            cmbFecProg.AddItem Format(CE(rsfechas.Fields("fec_pago")), "dd/mm/yyyy")
            rsfechas.MoveNext
        Loop
        For I = 0 To cmbFecProg.ListCount - 1
            If cmbFecProg.List(I, 0) = frmMovTelewiese.mskFecha Then
                cmbFecProg.ListIndex = I
            Else
                cmbFecProg.ListIndex = 0
            End If
        Next
        cmbFecProg.BackColor = ColorHabilitado
        cmbFecProg.Enabled = True
    Else
        cmbFecProg.ListIndex = 0
    End If
End Sub
 
Private Sub CargarGrilla(aux As String, Cod As String, Optional fechaprg As String, Optional mon As String)
    Dim rsdocs As MYSQL_RS
    Dim SQL As String, Str1 As String, str2 As String
    Dim I As Integer
   '***** revisar
    If fechaprg <> "" Then Str1 = " and fec_pago='" & Format(fechaprg, "yyyy/mm/dd") & "' "
    If mon <> "" Then str2 = " and mon='" & mon & "' "
    SQL = "DOC_PROG WHERE " & _
          " COD_ESTADO <>'" & ELIMINADO & "' AND " & _
          " COD_ESTADO <>'" & ANULADO & "' AND " & _
          " (COD_ESTADO <>'" & CANCELADO & "' OR CANCELADO< TOTAL )  AND" & _
          " AUXILIAR ='" & frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1) & "' AND " & _
          " CODIGO='" & frmMovTelewiese.txtCodigo & "'" & Str1 & str2 & " ORDER BY cod_tipo_doc,documento"
    Set rsdocs = oConexion.EjecutaSelect(SQL)
    With flexAnexos
        If rsdocs.RecordCount > 0 Then
            LlenarGrilla
            Do While Not rsdocs.EOF
                .Rows = .Rows + 1
                .Col = 12
                .row = .Rows - 1
                .CellFontName = "Wingdings"
                .CellFontSize = 14
                .CellAlignment = flexAlignCenterCenter
                .Col = 6 'Documento
                .CellBackColor = &HFFFFC0
                .Col = 7 'Documento
                .CellBackColor = &HFFFFC0
                .Col = 8 'Importe
                .CellBackColor = &HC0E0FF
                .Col = 11 'Fecha de Pago
                .CellBackColor = &HC0FFFF
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Left(CE(rsdocs.Fields("identificador")), 6)
                .TextMatrix(.Rows - 1, 2) = Right(CE(rsdocs.Fields("identificador")), 4)
                .TextMatrix(.Rows - 1, 3) = CE(rsdocs.Fields("cod_tipo_doc"))
                .TextMatrix(.Rows - 1, 4) = CE(rsdocs.Fields("mon"))
                .TextMatrix(.Rows - 1, 5) = Format(CE(rsdocs.Fields("Fec_emision")), "dd/mm/yyyy")
                    If (aux = "5" Or aux = "6") And .TextMatrix(.Rows - 1, 3) <> "TR" Then
                        .TextMatrix(.Rows - 1, 8) = FormatNumber(SaldoDoc(CE(rsdocs.Fields("cod_tipo_doc")), Trim(rsdocs.Fields("serie")), Trim(rsdocs.Fields("documento")), aux, Cod, mon), 2) 'FormatNumber(CEN(rsdocs.Fields("total")), 2)
                        .TextMatrix(.Rows - 1, 9) = FormatNumber(SaldoDoc(CE(rsdocs.Fields("cod_tipo_doc")), Trim(rsdocs.Fields("serie")), Trim(rsdocs.Fields("documento")), aux, Cod, IIf(mon = "N", "E", "N")), 2)  'FormatNumber(CEN(rsdocs.Fields("impequi")), 2)
                        .TextMatrix(.Rows - 1, 6) = Trim(rsdocs.Fields("serie"))
                        .TextMatrix(.Rows - 1, 7) = Trim(rsdocs.Fields("documento"))
                        .TextMatrix(.Rows - 1, 14) = CE(NDOCDivi)
                    Else
                        .TextMatrix(.Rows - 1, 8) = FormatNumber(CEN(rsdocs.Fields("total")), 2)
                        .TextMatrix(.Rows - 1, 9) = FormatNumber(CEN(rsdocs.Fields("impequi")), 2)
                        .TextMatrix(.Rows - 1, 6) = Trim(rsdocs.Fields("serie"))
                        .TextMatrix(.Rows - 1, 7) = Trim(rsdocs.Fields("documento"))
                        If aux = 3 Then
                            .TextMatrix(.Rows - 1, 14) = CE(rsdocs.Fields("division"))
                        Else
                            .TextMatrix(.Rows - 1, 14) = "000000000000"
                        End If
                    End If
                .TextMatrix(.Rows - 1, 10) = Format(CE(rsdocs.Fields("fec_vcto")), "dd/mm/yyyy")
                .TextMatrix(.Rows - 1, 11) = Format(CE(rsdocs.Fields("fec_pago")), "dd/mm/yyyy")
                .TextMatrix(.Rows - 1, 12) = strUnChecked
                .TextMatrix(.Rows - 1, 13) = CE(rsdocs.Fields("identificador"))
                rsdocs.MoveNext
            Loop
        End If
    End With
End Sub

Private Sub btnAceptar_Click()
    On Error GoTo NADA
    Dim I As Integer
    Dim num As Integer
    Dim Item As Integer
    If frmMovTelewiese.flexDocumentos.Rows > 1 Then
        Item = frmMovTelewiese.flexDocumentos.Rows
    Else
        Item = 1
    End If
    num = 1
    For I = 1 To flexAnexos.Rows - 1
        If flexAnexos.TextMatrix(I, 12) = strChecked Then
            With frmMovTelewiese.flexDocumentos
                If BuscaFolioAnexado(flexAnexos.TextMatrix(I, 13)) = False Then
                    .Rows = .Rows + 1
                    .TextMatrix(Item, 1) = flexAnexos.TextMatrix(I, 13)
                    .TextMatrix(Item, 2) = frmMovTelewiese.txtCodigo
                    .TextMatrix(Item, 3) = flexAnexos.TextMatrix(I, 3)
                    .TextMatrix(Item, 4) = flexAnexos.TextMatrix(I, 4)
                    .TextMatrix(Item, 5) = flexAnexos.TextMatrix(I, 5)
                    .TextMatrix(Item, 6) = flexAnexos.TextMatrix(I, 6)
                    .TextMatrix(Item, 7) = flexAnexos.TextMatrix(I, 7)
                    .TextMatrix(Item, 8) = flexAnexos.TextMatrix(I, 8)
                    .TextMatrix(Item, 9) = flexAnexos.TextMatrix(I, 9)
                    .TextMatrix(Item, 10) = frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1)
                    .TextMatrix(Item, 11) = frmMovTelewiese.cmbTipoPago.List(frmMovTelewiese.cmbTipoPago.ListIndex, 1)
                    .TextMatrix(Item, 12) = frmMovTelewiese.txtOficina
                    .TextMatrix(Item, 13) = frmMovTelewiese.txtCuentaAux
                    .TextMatrix(Item, 14) = flexAnexos.TextMatrix(I, 14)
                    
                    Item = Item + 1
                End If
            End With
            num = num + 1
        Else '
            With frmMovTelewiese.flexDocumentos
                If BuscaFolioAnexado(flexAnexos.TextMatrix(I, 13)) = True Then
                    EliminaItems flexAnexos.TextMatrix(I, 13)
                    Item = Item - 1
                End If
            End With
        End If
    Next
    If num = 1 Then frmMovTelewiese.lblDocAnexados = "": frmMovTelewiese.lblDocAnexados.Visible = False: frmMovTelewiese.lblDocAnexados.tag = 0
    If num = 2 Then frmMovTelewiese.lblDocAnexados = str(num - 1) & " Documento Anexado": frmMovTelewiese.lblDocAnexados.Visible = True: frmMovTelewiese.lblDocAnexados.tag = num - 1
    If num > 2 Then frmMovTelewiese.lblDocAnexados = str(num - 1) & " Documentos Anexados": frmMovTelewiese.lblDocAnexados.Visible = True: frmMovTelewiese.lblDocAnexados.tag = num - 1
    frmMovTelewiese.meImporte = lblimporte
    Unload Me
    EnumerarItems frmMovTelewiese.flexDocumentos
    If frmMovTelewiese.flexDocumentos.Rows > 1 Then frmMovTelewiese.flexDocumentos.row = frmMovTelewiese.flexDocumentos.Rows - 1 '1
    Call keybd_event(vbKeyEnd, 0, 0, 0)
    Exit Sub
NADA:
    Exit Sub
End Sub
Private Function EliminaItems(folio As String)
    Dim J As Integer
    If frmMovTelewiese.flexDocumentos.Rows > 2 Then
        For J = 1 To frmMovTelewiese.flexDocumentos.Rows - 1
            If frmMovTelewiese.flexDocumentos.TextMatrix(J, 1) = folio Then
                frmMovTelewiese.flexDocumentos.RemoveItem J
            End If
        Next
    End If
    If frmMovTelewiese.flexDocumentos.Rows = 2 Then
        For J = 1 To frmMovTelewiese.flexDocumentos.Rows - 1
            If frmMovTelewiese.flexDocumentos.TextMatrix(J, 1) = folio Then
                frmMovTelewiese.flexDocumentos.Rows = 1
            End If
        Next
    End If
End Function
Private Function BuscaFolioAnexado(folio) As Boolean
    Dim I As Integer
    BuscaFolioAnexado = False
    With frmMovTelewiese.flexDocumentos
        For I = 1 To .Rows - 1
            If Trim(folio) = .TextMatrix(I, 1) Then
                BuscaFolioAnexado = True
                Exit Function
            End If
        Next
    End With
End Function
Private Sub cmbFecProg_Change()
    If cmbFecProg.ListCount > 0 Then
        CargarGrilla frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1), frmMovTelewiese.txtCodigo, IIf(cmbFecProg.ListIndex = 0, "", cmbFecProg.List(cmbFecProg.ListIndex, 0)), IIf(cmbMoneda.ListIndex = 0, "", cmbMoneda.List(cmbMoneda.ListIndex, 1))
        If frmMovTelewiese.lblDocAnexados.Visible = True Then CargaAnexos frmMovTelewiese.flexDocumentos.Rows - 1
    End If
End Sub
Private Sub cmbMoneda_Change()
    With frmMovTelewiese
        If .cmbMoneda.ListIndex = 0 Then
            lblMImp.Caption = .lblMoneda.Caption
            lblMImpEqu.Caption = .lblMonEqu.Caption
        Else
            If .cmbMoneda.ListIndex = 1 Then
                lblMImp.Caption = "S/."
                lblMImpEqu.Caption = "US$"
            Else
                lblMImp.Caption = "US$"
                lblMImpEqu.Caption = "S/."
            End If
        End If
        If cmbFecProg.ListCount > 0 Then
            CargarGrilla .cmbAuxiliares.List(.cmbAuxiliares.ListIndex, 1), .txtCodigo, IIf(cmbFecProg.ListIndex = 0, "", cmbFecProg.List(cmbFecProg.ListIndex, 0)), IIf(cmbMoneda.ListIndex = 0, "", cmbMoneda.List(cmbMoneda.ListIndex, 1))
        End If
    End With
End Sub
Private Sub flexAnexos_Click()
    With flexAnexos
    On Error GoTo Importe
        If CDbl(Trim(.TextMatrix(.row, 8))) > 0 Then
            If .Col = 12 And (frmMovTelewiese.lblModo = "Nuevo Movimiento" Or frmMovTelewiese.lblModo = "Modificar Movimiento") Then
                If .TextMatrix(.row, 12) = strUnChecked Then
                    .TextMatrix(.row, 12) = strChecked
                    If .TextMatrix(.row, 4) = "N" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
                        lblimporte = FormatNumber(str(CDbl(lblimporte) + CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblImporteEqu) + CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(flexAnexos.row, 4) = "E" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
                        lblimporte = FormatNumber(str(CDbl(lblImporteEqu) + CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblimporte) + CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(.row, 4) = "N" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "E" Then
                        lblimporte = FormatNumber(str(CDbl(lblImporteEqu) + CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblimporte) + CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(.row, 4) = "E" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "E" Then
                        lblimporte = FormatNumber(str(CDbl(lblimporte) + CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblImporteEqu) + CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                Else
                    .TextMatrix(.row, 12) = strUnChecked
                    If .TextMatrix(.row, 4) = "N" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
                        lblimporte = FormatNumber(str(CDbl(lblimporte) - CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblImporteEqu) - CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(.row, 4) = "N" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "E" Then
                        lblimporte = FormatNumber(str(CDbl(lblImporteEqu) - CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblimporte) - CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(.row, 4) = "E" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "N" Then
                        lblimporte = FormatNumber(str(CDbl(lblImporteEqu) - CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblimporte) - CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                    If .TextMatrix(.row, 4) = "E" And frmMovTelewiese.cmbMoneda.List(cmbMoneda.ListIndex, 1) = "E" Then
                        lblimporte = FormatNumber(str(CDbl(lblimporte) - CDbl(.TextMatrix(.row, 8))), 2)
                        lblImporteEqu = FormatNumber(str(CDbl(lblImporteEqu) - CDbl(.TextMatrix(.row, 9))), 2)
                    End If
                End If
            End If
        Else
            MsgBox "Documento cancelado o no registrado en contabilidad", vbInformation + vbOKOnly, "NOVPeru"
            Exit Sub
        End If
    End With
    Exit Sub
Importe:
    MsgBox "Los datos no son correcto, revise o consulte con el administrador del sistema", vbInformation + vbOKOnly, "NOVPeru"
    Exit Sub
End Sub
Private Sub Form_Load()
    Call WheelHook(frmAnexosTelewiese)
    LlenarGrilla
    moneda cmbMoneda
    lblCheque.Caption = frmMovTelewiese.meOrden
    lblSrs.Caption = frmMovTelewiese.txtBeneficiario
    cmbMoneda.ListIndex = frmMovTelewiese.cmbMoneda.ListIndex
    lblFecPago.Caption = Format(Date, "dd/mm/yyyy")
    cmbFecProg.BackColor = ColorDeshabilitado
    cmbFecProg.Enabled = False
    CargarFechasProg frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1), frmMovTelewiese.txtCodigo
    If frmMovTelewiese.lblModo = "Nuevo Movimiento" And frmMovTelewiese.lblDocAnexados = "" Then
        CargarGrilla frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1), frmMovTelewiese.txtCodigo, IIf(cmbFecProg.ListIndex = 0, "", cmbFecProg.List(cmbFecProg.ListIndex, 0)), IIf(cmbMoneda.ListIndex = 0, "", cmbMoneda.List(cmbMoneda.ListIndex, 1))
    End If
    If frmMovTelewiese.lblModo = "Modificar Movimiento" Then
        CargarAnexosGrabados strAnoSistema + strMesSistema + frmMovTelewiese.meOrden, frmMovTelewiese.cmbAuxiliares.List(frmMovTelewiese.cmbAuxiliares.ListIndex, 1), frmMovTelewiese.txtCodigo, frmMovTelewiese.cmbMoneda.List(frmMovTelewiese.cmbMoneda.ListIndex, 1)
    End If
End Sub
Public Sub CargaAnexos(Anexos As Integer)
    Dim I As Integer, J As Integer
    For I = 1 To flexAnexos.Rows - 1
        For J = 1 To Anexos
            If flexAnexos.TextMatrix(I, 13) = frmMovTelewiese.flexDocumentos.TextMatrix(J, 1) Then
                flexAnexos.row = I
                flexAnexos.Col = 12
                flexAnexos_Click
            End If
        Next
    Next
End Sub
Public Sub CargarAnexosGrabados(orden As String, aux As String, Cod As String, mon As String)
    Dim I As Integer
    Dim SQL As String
    Dim rsdocs As MYSQL_RS
    SQL = "Select a.identificador,b.cod_tipo_doc,a.mon,a.fec_emision,a.auxiliar,a.codigo," & _
          " concat(a.serie,'-',a.correl) as documento,a.serie as SERDOC,a.correl as NUMDOC,a.total,a.impequi,a.fec_vcto,a.division" & _
          " from  documento_contables as a left join amarre_documento as b " & _
          " on (a.identificador=b.identificador) where a.ref='WIESE" & orden & "' order by b.cod_tipo_doc,a.serie,a.correl"
    Set rsdocs = oConexion.EjecutaSelectRS(SQL)
    If rsdocs.RecordCount > 0 Then
        rsdocs.MoveFirst
        With flexAnexos
            Do While Not rsdocs.EOF
                .Rows = .Rows + 1
                .Col = 12
                .row = .Rows - 1
                .CellFontName = "Wingdings"
                .CellFontSize = 14
                .CellAlignment = flexAlignCenterCenter
                .Col = 6 'Documento
                .CellBackColor = &HFFFFC0
                .Col = 7 'Importe
                .CellBackColor = &HC0E0FF
                .Col = 10 'Fecha de Pago
                .CellBackColor = &HC0FFFF
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Left(CE(rsdocs.Fields("identificador")), 6)
                .TextMatrix(.Rows - 1, 2) = Right(CE(rsdocs.Fields("identificador")), 4)
                .TextMatrix(.Rows - 1, 3) = CE(rsdocs.Fields("cod_tipo_doc"))
                .TextMatrix(.Rows - 1, 4) = CE(rsdocs.Fields("mon"))
                .TextMatrix(.Rows - 1, 5) = Format(CE(rsdocs.Fields("Fec_emision")), "dd/mm/yyyy")
                If aux = "5" Or aux = "6" Then
                    .TextMatrix(.Rows - 1, 7) = FormatNumber(SaldoDoc(CE(rsdocs.Fields("cod_tipo_doc")), Trim(rsdocs.Fields("SERDOC")), Trim(rsdocs.Fields("NUMDOC")), aux, Cod, mon), 2) 'FormatNumber(CEN(rsdocs.Fields("total")), 2)
                    .TextMatrix(.Rows - 1, 8) = FormatNumber(SaldoDoc(CE(rsdocs.Fields("cod_tipo_doc")), Trim(rsdocs.Fields("SERDOC")), Trim(rsdocs.Fields("NUMDOC")), aux, Cod, IIf(mon = "N", "E", "N")), 2) 'FormatNumber(CEN(rsdocs.Fields("impequi")), 2)
                    .TextMatrix(.Rows - 1, 6) = CE(NDOCCont)
                    .TextMatrix(.Rows - 1, 13) = CE(NDOCDivi)
                Else
                    .TextMatrix(.Rows - 1, 7) = FormatNumber(CEN(rsdocs.Fields("total")), 2)
                    .TextMatrix(.Rows - 1, 8) = FormatNumber(CEN(rsdocs.Fields("impequi")), 2)
                    .TextMatrix(.Rows - 1, 6) = CE(rsdocs.Fields("documento"))
                    If aux = 3 Then
                        .TextMatrix(.Rows - 1, 13) = CE(rsdocs.Fields("division"))
                    Else
                        .TextMatrix(.Rows - 1, 13) = "0000"
                    End If
                End If
                .TextMatrix(.Rows - 1, 9) = Format(CE(rsdocs.Fields("fec_vcto")), "dd/mm/yyyy")
                .TextMatrix(.Rows - 1, 10) = Format(CE(rsdocs.Fields("fec_pago")), "dd/mm/yyyy")
                .TextMatrix(.Rows - 1, 11) = strChecked
                .TextMatrix(.Rows - 1, 12) = CE(rsdocs.Fields("identificador"))
                lblimporte = str(CDbl(lblimporte) + CDbl(CEN(.TextMatrix(.Rows - 1, 7))))
                lblImporteEqu = str(CDbl(lblImporteEqu) + CDbl(CEN(.TextMatrix(.Rows - 1, 8))))
                rsdocs.MoveNext
            Loop
        End With
    End If
    rsdocs.CloseRecordset
    Set rsdocs = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If CDbl(lblimporte) = 0 And (frmMovTelewiese.lblModo = "Nuevo Movimiento" Or frmMovTelewiese.lblModo = "Modificar Movimiento") Then
        resp = MsgBox("No ha selecionado ningún documento... ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso")
        If resp = vbYes Then
            frmMovTelewiese.lblDocAnexados = Empty
            frmMovTelewiese.lblDocAnexados.Visible = False
            Call btnAceptar_Click
        Else
            Cancel = 1
        End If
    End If
    WheelUnHook
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flexAnexos
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
