VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmTramiteGerencial 
   BackColor       =   &H00A49FB5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trámite de Documentos"
   ClientHeight    =   6120
   ClientLeft      =   3405
   ClientTop       =   5340
   ClientWidth     =   11805
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmTramiteGerencial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11805
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   5295
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   11805
      Begin VB.ComboBox cboFiltro 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2670
         TabIndex        =   4
         Text            =   "cboFiltro"
         Top             =   420
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTramite 
         Height          =   4095
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   16777215
         GridColor       =   8421504
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblCodigo 
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
         Left            =   6360
         TabIndex        =   15
         Top             =   4890
         Width           =   2715
      End
      Begin VB.Label lblImporte 
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
         Left            =   9120
         TabIndex        =   13
         Top             =   4890
         Width           =   2535
      End
      Begin VB.Label lblFecPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000008&
         Caption         =   "  "
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
         Left            =   9120
         TabIndex        =   12
         Top             =   4350
         Width           =   2535
      End
      Begin VB.Label lblDivision 
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
         Left            =   120
         TabIndex        =   11
         Top             =   4620
         Width           =   3285
      End
      Begin VB.Label lblCenco 
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
         Left            =   3450
         TabIndex        =   10
         Top             =   4620
         Width           =   8205
      End
      Begin VB.Label lblAuxiliar 
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
         Left            =   120
         TabIndex        =   9
         Top             =   4890
         Width           =   6255
      End
      Begin VB.Label lblDocNum 
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
         Left            =   120
         TabIndex        =   7
         Top             =   4350
         Width           =   8955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   915
         Left            =   60
         TabIndex        =   6
         Top             =   4290
         Width           =   11685
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   11805
      Begin Proyecto1.chameleonButton btnSalir 
         Height          =   345
         Left            =   11250
         TabIndex        =   1
         Top             =   510
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
         MICON           =   "frmTramiteGerencial.frx":014A
         PICN            =   "frmTramiteGerencial.frx":0166
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
         Left            =   10710
         TabIndex        =   2
         Top             =   510
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
         MICON           =   "frmTramiteGerencial.frx":052C
         PICN            =   "frmTramiteGerencial.frx":0548
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flxColores 
         Height          =   795
         Left            =   90
         TabIndex        =   8
         Top             =   120
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1402
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin Proyecto1.chameleonButton btnRefrescar 
         Height          =   345
         Left            =   10140
         TabIndex        =   14
         ToolTipText     =   "Refrescar Búsqueda - F5"
         Top             =   510
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
         MICON           =   "frmTramiteGerencial.frx":0A8A
         PICN            =   "frmTramiteGerencial.frx":0AA6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmTramiteGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Color(1 To 15) As String
Private DocColor(1 To 70, 1 To 2) As String
Private filtro(1 To 17, 1 To 2) As String
Private columna As Integer
Private Sub ConfigGrilla()
    With flxTramite
        .Clear
        .Rows = 1
        .Cols = 18
        .RowHeight(0) = 315
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(0) + "Item"
        .FixedCols = 1
        .FixedRows = 0
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Folio"
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = Space(0) + "Tipo"
        .ColWidth(3) = 1700
        .TextMatrix(0, 3) = Space(3) + "N° Documento"
        .ColWidth(4) = 3200
        .TextMatrix(0, 4) = Space(8) + "Auxiliar"
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = Space(2) + "Fec. Pago"
        .ColWidth(6) = 400
        .TextMatrix(0, 6) = Space(0) + "Mon"
        .ColWidth(7) = 1000
        .TextMatrix(0, 7) = Space(4) + "Importe"
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = Space(8) + "CodEst1"
        .ColWidth(9) = 1500
        .TextMatrix(0, 9) = Space(8) + "Est_Act"
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = Space(8) + "CodEst2"
        .ColWidth(11) = 400
        .TextMatrix(0, 11) = ""
        .ColWidth(12) = 1500
        .TextMatrix(0, 12) = Space(8) + "Est_Sgt"
        .ColWidth(13) = 700
        .TextMatrix(0, 13) = Space(2) + "ccHFM."
        .ColWidth(14) = 1800
        .TextMatrix(0, 14) = Space(8) + "Cenco"
        .ColWidth(15) = 0
        .TextMatrix(0, 15) = Space(0) + "Aux"
        .ColWidth(16) = 0
        .TextMatrix(0, 16) = Space(7) + "CodAux"
        .ColWidth(17) = 0
        .TextMatrix(0, 17) = "Familia"
        For I = 0 To 17
            .row = 0
            .Col = I
            .CellForeColor = &H80000002
            .CellBackColor = &H8000000F
        Next I
    End With
End Sub
Private Sub LlenaGrilla(Query As String)
    Dim I%, J%, k%, p%
    Dim Doc As String
    Dim ultciclo As String
    Dim rsgrid As MYSQL_RS
    Me.MousePointer = vbHourglass
    If Query = "" Then
        Query = "select a.Identificador AS Identificador," & _
                " concat(doc.Serie,'-',doc.Correl) AS NumDoc," & _
                " a.Cod_Tipo_Doc AS TipoDoc,TRIM(doc.Cenco) AS cenco, ax.descrip as descrip," & _
                " doc.Auxiliar AS auxiliar, doc.Division as division, doc.Fec_Pago as FPago, doc.Codigo AS Codaux,doc.Mon AS moneda," & _
                " doc.Total AS importe,m.Cod_Estado AS estado,a.Cod_Fam AS familia " & _
                " from (documento_contables as doc left join movi_documento As m " & _
                " on (m.Identificador = doc.Identificador)) left join  amarre_documento As a " & _
                " on (m.Identificador = a.Identificador) left join cnauxil as ax " & _
                " on (doc.Codigo = ax.codigo and doc.auxiliar = ax.auxiliar)" & _
                " right join (SELECT coddoc FROM cndocum C WHERE (c.protegido = 'N' OR " & _
                " (SELECT permiso FROM docsusuario D WHERE D.coddoc=C.coddoc AND usuario = '" & strUsuarioId & "')=1)) K " & _
                " on (a.Cod_Tipo_Doc=k.coddoc) " & _
                " where (a.Flag = '1')" & _
                " and (m.Cod_Estado <> '" & CANCELADO & "' and m.Cod_estado <> '" & APROBADO & "'" & _
                " and m.Cod_Estado <> '" & ANULADO & "' and  m.Cod_estado <> '" & NING & "' and m.Cod_estado <> '" & ELIMINADO & "')  " & _
                " order by TipoDoc, Auxiliar, Codaux, Numdoc;  "
    End If
    Set rsgrid = oConexion.EjecutaSelectRS(Query)
    p = 1
    With flxTramite
        flxColores.Clear
        flxColores.Rows = 1
        I = 1
        Do While Not rsgrid.EOF
            If Doc = "" Then
                Doc = UCase(rsgrid.Fields("TipoDoc"))
                ultciclo = CiclodeVidaDoc(UCase(CE(rsgrid.Fields("TipoDoc"))))
            Else
                If Doc <> UCase(rsgrid.Fields("TipoDoc")) Then
                    Doc = UCase(rsgrid.Fields("TipoDoc"))
                    ultciclo = CiclodeVidaDoc(UCase(CE(rsgrid.Fields("TipoDoc"))))
                End If
            End If
            If Not CE(rsgrid.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                .Rows = .Rows + 1
                .TextMatrix(I, 1) = rsgrid.Fields("Identificador")
                .TextMatrix(I, 2) = Space(1) & UCase(CE(rsgrid.Fields("TipoDoc")))
                .TextMatrix(I, 3) = Space(0) & CE(rsgrid.Fields("NumDoc"))
                .TextMatrix(I, 4) = Space(1) & DescripcionesdeCodigos("AUXILIARES", Trim(CE(rsgrid.Fields("codaux"))), CE(rsgrid.Fields("auxiliar")), "Descrip")
                .TextMatrix(I, 5) = Format(CE(rsgrid.Fields("FPago")), "dd/mm/yyyy")
                .TextMatrix(I, 6) = Space(1) & CE(rsgrid.Fields("moneda"))
                .TextMatrix(I, 7) = FormatNumber(CE(rsgrid.Fields("Importe")), 2)
                .TextMatrix(I, 8) = CE(rsgrid.Fields("estado"))
                .TextMatrix(I, 9) = Space(1) & DescripcionesdeCodigos("DOC_ESTADO", Trim(CE(rsgrid.Fields("estado"))))
                .TextMatrix(I, 13) = CE(rsgrid.Fields("Division"))
                .TextMatrix(I, 14) = Space(1) & CE(rsgrid.Fields("Cenco"))
                .TextMatrix(I, 15) = Space(1) & CE(rsgrid.Fields("auxiliar"))
                .TextMatrix(I, 16) = Space(1) & CE(rsgrid.Fields("codaux"))
                .Col = 11
                .row = I
                .CellFontName = "Wingdings"
                .CellFontSize = 11
                .Text = strUnChecked
                For J = 1 To 10
                    If ciclo(J, 1) = rsgrid.Fields("estado") Then
                        .TextMatrix(I, 10) = ciclo(J + 1, 1)
                        .TextMatrix(I, 12) = Space(1) & Replace(DescripcionesdeCodigos("DOC_ESTADO", Trim(CE(.TextMatrix(I, 10)))), "ADO", "AR")
                        If ciclo(J + 1, 2) = CStr(1) Then
                            .Col = 12
                            .CellForeColor = &H80&
                            .CellFontBold = True
                            .Col = 11
                            .row = I
                            .CellForeColor = vbBlack
                        Else
                            .Col = 12
                            .CellForeColor = &H808080
                            .Col = 11
                            .row = I
                            .CellForeColor = &H808080
                        End If
                        Exit For
                    End If
                Next
                .TextMatrix(I, 17) = CE(rsgrid.Fields("familia"))
                For k = 1 To 70
                    If DocColor(k, 1) = Doc Then
                        For p = 1 To 17
                            .Col = p
                            .CellBackColor = DocColor(k, 2)
                        Next
                        GridColor Doc, DocColor(k, 2)
                        Exit For
                    End If
                Next
                I = I + 1
            Else
                rsgrid.Fields("TipoDoc") = "*"
            End If
            rsgrid.MoveNext
        Loop
    End With
    EnumerarItems1 flxTramite
    Set rsgrid = Nothing
    Me.MousePointer = vbNormal
End Sub
Private Sub Command1_Click()
    flxTramite.Visible = False
    flxTramite.Visible = True
End Sub
Private Sub btnRefrescar_Click()
    Me.MousePointer = vbHourglass
    flxTramite.Visible = False
    ConfigGrilla
    LlenaGrilla ""
    Me.MousePointer = vbNormal
    DesplazarFlx
    flxTramite.Visible = True
    If flxTramite.Rows > 1 Then
        flxTramite.row = 1
        flxTramite.ColSel = 17
    End If
    Limpiar
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub cboFiltro_Click()
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I%, J%
    SQL = "select a.Identificador AS Identificador," & _
          " concat(doc.Serie,'-',doc.Correl) AS NumDoc," & _
          " a.Cod_Tipo_Doc AS TipoDoc,TRIM(doc.Cenco) AS cenco, ax.descrip as descrip," & _
          " doc.Auxiliar AS auxiliar, doc.Division as division, doc.Fec_Pago as FPago, doc.Codigo AS Codaux,doc.Mon AS moneda," & _
          " doc.Total AS importe,m.Cod_Estado AS estado,a.Cod_Fam AS familia " & _
          " from (documento_contables as doc left join movi_documento As m " & _
          " on (m.Identificador = doc.Identificador)) left join  amarre_documento As a " & _
          " on (m.Identificador = a.Identificador) left join cnauxil as ax " & _
          " on (doc.Codigo = ax.codigo and doc.auxiliar = ax.auxiliar)" & _
          " where (a.Flag = '1')" & _
          " and (m.Cod_Estado <> '" & CANCELADO & "' and m.Cod_estado <> '" & APROBADO & "'" & _
          " and m.Cod_Estado <> '" & ANULADO & "' AND  m.Cod_Estado <> '" & NING & "' and m.Cod_estado <> '" & ELIMINADO & "')  "
    If flxTramite.Col = 7 And cboFiltro.Text <> "Seleccionar..." Then
        filtro(flxTramite.Col, 2) = CDbl(cboFiltro.Text)
        flxTramite.row = 0
        flxTramite.CellBackColor = vbRed
    Else
        If flxTramite.Col = 9 Then
            filtro(flxTramite.Col, 2) = DescripcionesdeCodigos("ESTADOenCODIGO", Trim(cboFiltro.Text))
        Else
                filtro(flxTramite.Col, 2) = cboFiltro.Text
        End If
    End If
    If cboFiltro.Text = "Seleccionar..." Then
        filtro(flxTramite.Col, 1) = Empty
        filtro(flxTramite.Col, 2) = Empty
    End If
    For I = 1 To 17
        If filtro(I, 1) <> Empty Then
            If I = 5 Then
                SQL = SQL & " and " & filtro(I, 1) & " = '" & Format(filtro(I, 2), "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " and " & filtro(I, 1) & " = '" & filtro(I, 2) & "'"
            End If
            
        End If
    Next
    SQL = SQL & " order by TipoDoc, Auxiliar, Codaux, Numdoc; "
    flxTramite.Visible = False
    ConfigGrilla
    LlenaGrilla SQL
    flxTramite.Visible = True
    flxTramite.row = 0
    cboFiltro.Visible = False
    DesplazarFlx
    For I = 1 To 17
        If filtro(I, 1) <> Empty Then
            flxTramite.row = 0
            flxTramite.Col = I
            flxTramite.CellForeColor = vbRed
        End If
    Next
    If flxTramite.Rows > 1 Then
        flxTramite.row = 1
        flxTramite.ColSel = 17
    End If
End Sub
Private Sub cboFiltro_DropDown()
    Dim sqlcombo As String
    Dim ultciclo As String
    Dim Doc As String
    Dim rscombo As MYSQL_RS
    Dim J%
    sqlcombo = "select a.Identificador AS Identificador," & _
          " concat(doc.Serie,'-',doc.Correl) AS NumDoc," & _
          " a.Cod_Tipo_Doc AS TipoDoc,TRIM(doc.Cenco) AS cenco, ax.descrip as descrip," & _
          " doc.Auxiliar AS auxiliar, doc.Division as division, doc.Fec_Pago as FPago," & _
          " doc.Codigo AS Codaux,doc.Mon AS moneda," & _
          " doc.Total AS importe,m.Cod_Estado AS estado,a.Cod_Fam AS familia " & _
          " from (documento_contables as doc left join movi_documento As m " & _
          " on (m.Identificador = doc.Identificador)) left join  amarre_documento As a " & _
          " on (m.Identificador = a.Identificador) left join cnauxil as ax " & _
          " on (doc.Codigo = ax.codigo and doc.auxiliar = ax.auxiliar )" & _
          " where (a.Flag = '1')" & _
          " and (m.Cod_Estado <> '" & CANCELADO & "' and m.Cod_estado <> '" & APROBADO & "'" & _
          " and m.Cod_Estado <> '" & ANULADO & "' and m.Cod_estado <> '" & ELIMINADO & "')  "
    If cboFiltro.Text = "Auxiliar" Then
        sqlcombo = sqlcombo & "group by Codaux order by Descrip;"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If Doc = "" Then
                Doc = rscombo.Fields("TipoDoc")
                ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
            Else
                If Doc <> rscombo.Fields("TipoDoc") Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                End If
            End If
            If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                cboFiltro.AddItem CE(rscombo.Fields("descrip"))
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "ax.descrip"
        cboFiltro.Text = "Auxiliar"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "Tipo" Then
        sqlcombo = sqlcombo & "group by TipoDoc order by TipoDoc;"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If Doc = "" Then
                Doc = rscombo.Fields("TipoDoc")
                ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
            Else
                If Doc <> rscombo.Fields("TipoDoc") Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                End If
            End If
            If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                cboFiltro.AddItem CE(rscombo.Fields("TipoDoc"))
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "a.Cod_Tipo_Doc"
        cboFiltro.Text = "Tipo"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "Mon" Then
        sqlcombo = sqlcombo & "group by moneda order by moneda;"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            cboFiltro.AddItem CE(rscombo.Fields("moneda"))
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "doc.Mon"
        cboFiltro.Text = "Mon"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "N° Documento" Then
        sqlcombo = sqlcombo & "group by NumDoc order by NumDoc;"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If Doc = "" Then
                Doc = CE(rscombo.Fields("TipoDoc"))
                ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
            Else
                If Doc <> rscombo.Fields("TipoDoc") Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                End If
            End If
            If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                cboFiltro.AddItem CE(rscombo.Fields("NumDoc"))
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "concat(doc.Serie,'-',doc.Correl)"
        cboFiltro.Text = "N° Documento"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "Fec. Pago" Then
        sqlcombo = sqlcombo & "group by FPago order by FPago"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If rscombo.Fields("FPago") <> Empty Then
                If Doc = "" Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                Else
                    If Doc <> rscombo.Fields("TipoDoc") Then
                        Doc = rscombo.Fields("TipoDoc")
                        ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                    End If
                End If
                If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                    cboFiltro.AddItem Format(rscombo.Fields("FPago"), "dd/mm/YY")
                End If
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "doc.Fec_Pago"
        cboFiltro.Text = "Fec. Pago"
        Doc = ""
    End If
    If cboFiltro.Text = "Importe" Then
        sqlcombo = sqlcombo & "group by importe order by estado"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If rscombo.Fields("Importe") <> Empty Then
                If Doc = "" Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                Else
                    If Doc <> rscombo.Fields("TipoDoc") Then
                        Doc = rscombo.Fields("TipoDoc")
                        ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                    End If
                End If
                If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                    cboFiltro.AddItem FormatNumber(CEN(rscombo.Fields("Importe")), 2)
                End If
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "doc.Total"
        cboFiltro.Text = "Importe"
        Doc = ""
    End If
    If cboFiltro.Text = "Est_Act" Then
        sqlcombo = sqlcombo & "group by Identificador order by estado"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If rscombo.Fields("Estado") <> Empty Then
                If Doc = "" Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                Else
                    If Doc <> rscombo.Fields("TipoDoc") Then
                        Doc = rscombo.Fields("TipoDoc")
                        ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                    End If
                End If
                If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                    Dim aux As String
                    aux = DescripcionesdeCodigos("DOC_ESTADO", Trim(CE(rscombo.Fields("estado"))))
                        If cboFiltro.List(cboFiltro.ListCount - 1) <> aux Then
                            cboFiltro.AddItem aux
                        End If
                End If
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "m.Cod_estado"
        cboFiltro.Text = "Est_Act"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "Est_Sgt" Then
        sqlcombo = sqlcombo & "group by Identificador order by estado"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If rscombo.Fields("Estado") <> Empty Then
                If Doc = "" Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                Else
                    If Doc <> rscombo.Fields("TipoDoc") Then
                        Doc = rscombo.Fields("TipoDoc")
                        ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                    End If
                End If
                If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                    For J = 1 To 10
                        If ciclo(J, 1) = rscombo.Fields("estado") Then
                            aux = Replace(DescripcionesdeCodigos("DOC_ESTADO", Trim(CE(ciclo(J + 1, 1)))), "ADO", "AR")
                            If cboFiltro.List(cboFiltro.ListCount - 1) <> aux Then
                                cboFiltro.AddItem Replace(aux, "ADO", "AR")
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            rscombo.MoveNext
        Wend
        Exit Sub
    End If
    If cboFiltro.Text = "Cenco" Then
        sqlcombo = sqlcombo & "group by Identificador order by cenco"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If Doc = "" Then
                Doc = rscombo.Fields("TipoDoc")
                ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
            Else
                If Doc <> rscombo.Fields("TipoDoc") Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                End If
            End If
            If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                If cboFiltro.List(cboFiltro.ListCount - 1) <> rscombo.Fields("cenco") Then
                    cboFiltro.AddItem rscombo.Fields("cenco")
                End If
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "doc.cenco"
        cboFiltro.Text = "Cenco"
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "Div." Then
        sqlcombo = sqlcombo & "group by Identificador order by division"
        Set rscombo = oConexion.EjecutaSelectRS(sqlcombo)
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
        While Not rscombo.EOF
            If Doc = "" Then
                Doc = rscombo.Fields("TipoDoc")
                ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
            Else
                If Doc <> rscombo.Fields("TipoDoc") Then
                    Doc = rscombo.Fields("TipoDoc")
                    ultciclo = CiclodeVidaDoc(CE(rscombo.Fields("TipoDoc")))
                End If
            End If
            If Not CE(rscombo.Fields("estado")) = ultciclo And ultciclo = CANCELADO Then
                aux = rscombo.Fields("division")
                If cboFiltro.List(cboFiltro.ListCount - 1) <> aux Then
                    cboFiltro.AddItem rscombo.Fields("division")
                End If
            End If
            rscombo.MoveNext
        Wend
        filtro(flxTramite.Col, 1) = "doc.division"
        cboFiltro.Text = "Div."
        Doc = ""
        Exit Sub
    End If
    If cboFiltro.Text = "" Then
        cboFiltro.Clear
        cboFiltro.AddItem "Seleccionar..."
    End If
    Set rscombo = Nothing
End Sub
Private Sub cboFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cboFiltro.Visible = False
End Sub
Private Sub flxColores_Click()
    With flxColores
        If .Col = 1 Then
         .TextMatrix(.row, 1) = IIf(.TextMatrix(.row, 1) = strChecked, strUnChecked, strChecked)
        End If
    End With
    If Trim(flxColores.TextMatrix(flxColores.row, 1)) = strChecked And flxColores.Col = 1 Then
        With flxTramite
            For I = 1 To .Rows - 1
                .Col = 11
                .row = I
                If .CellForeColor <> &H808080 Then
                    If Trim(.TextMatrix(I, 2)) = Trim(flxColores.TextMatrix(flxColores.row, 2)) Then
                        .TextMatrix(I, 11) = strChecked
                        CambioEstado Trim(.TextMatrix(I, 1)), Trim(.TextMatrix(I, 10))
                    End If
                End If
            Next
        End With
    End If
    If Trim(flxColores.TextMatrix(flxColores.row, 1)) = strUnChecked And flxColores.Col = 1 Then
        With flxTramite
            For I = 1 To .Rows - 1
                .Col = 11
                .row = I
                If .CellForeColor <> &H808080 Then
                    If Trim(.TextMatrix(I, 2)) = Trim(flxColores.TextMatrix(flxColores.row, 2)) Then
                       .TextMatrix(I, 11) = strUnChecked
                       DelHist Trim(.TextMatrix(I, 1)), Trim(.TextMatrix(I, 8))
                    End If
                End If
            Next
        End With
    End If
End Sub
Private Sub flxTramite_Click()
    With flxTramite
        If .ColSel = 11 And .CellForeColor <> &H808080 Then
           .TextMatrix(.row, 11) = IIf(.TextMatrix(.row, 11) = strChecked, strUnChecked, strChecked)
           If .TextMatrix(.row, 11) = strChecked Then
                CambioEstado Trim(.TextMatrix(.row, 1)), Trim(.TextMatrix(.row, 10))
           End If
           If .TextMatrix(.row, 11) = strUnChecked Then
                DelHist Trim(.TextMatrix(.row, 1)), Trim(.TextMatrix(.row, 8))
           End If
        End If
    End With
    If flxTramite.row = 0 And flxTramite.Col <> 11 _
       And flxTramite.Col <> 12 And flxTramite.Col <> 1 Then
        flxTramite.row = 0
        With cboFiltro
            .Top = flxTramite.CellTop + flxTramite.Top
            .Left = flxTramite.CellLeft + flxTramite.Left
            .Width = flxTramite.CellWidth
            .Text = Trim(flxTramite.Text)
            .Visible = True
            .ZOrder
            .SetFocus
            .SelStart = Len(.Text)
        End With
    Else
        cboFiltro.Visible = False
    End If
End Sub
Private Sub DesplazarFlx()
    With flxTramite
        lblDocNum = Space(2) & DescripcionesdeCodigos("CNDOCUM", Trim(.TextMatrix(.row, 2)), "1") & " N°: " & Trim(.TextMatrix(.row, 3))
        lblCenco = Space(2) & "Centro de Costo: " & DescripcionesdeCodigos("CENCO", Trim(.TextMatrix(.row, 14)), "1")
        lblDivision = Space(2) & "ccHFM: " & DescripcionesdeCodigos("DES_DIVISION", Trim(.TextMatrix(.row, 13)))
        lblAuxiliar = Space(2) & DescripcionesdeCodigos("CNAUXIL", Trim(.TextMatrix(.row, 15))) & ": " & _
                      DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(.row, 16)), Trim(.TextMatrix(.row, 15)), "Descrip")
        lblFecPago = "Fecha de Pago: " & Trim(CE(.TextMatrix(.row, 5)))
        If Trim(.TextMatrix(.row, 6)) = "N" Then
            lblimporte = Space(2) & "Importe= " & "S/. " & Trim(CEN(.TextMatrix(.row, 7)))
        End If
        If Trim(.TextMatrix(.row, 6)) = "E" Then
            lblimporte = Space(2) & "Importe= " & "US$ " & Trim(CEN(.TextMatrix(.row, 7)))
        End If
        If Trim(.TextMatrix(.row, 15)) = "5" Then
            lblCodigo = Space(1) & "R.U.C.: " & Trim(CE(.TextMatrix(.row, 16)))
        Else
            lblCodigo = Space(1) & "Código: " & Trim(CE(.TextMatrix(.row, 16)))
        End If
    End With
End Sub
Private Sub DelHist(Ident As String, Estado As String)
    Dim sqlMovi As String
    Dim sqlHist As String
    sqlMovi = "Call Update_Movidoc ( '" & Ident & "','" & Estado & "'," & _
          " '" & Format(Date, "yyyy/mm/dd") & "','" & strUsuarioId & "');  "
    sqlHist = "Call Delete_HistDoc ('" & Ident & "','" & Estado & "');"
    oConexion.EjecutaInsertUpdateDelete sqlHist, TIPO_QUERY.Eliminar, False
    oConexion.EjecutaInsertUpdateDelete sqlMovi, TIPO_QUERY.Modificar, False
End Sub
Private Sub flxTramite_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        btnRefrescar_Click
    End If
End Sub
Private Sub flxTramite_RowColChange()
    Dim I%, fila%
    DesplazarFlx
End Sub
Private Sub Form_Load()
    Call WheelHook(frmTramiteGerencial)
    Me.Top = 0
    Me.Left = 0
    ConfigColores
    ConfigGrilla
    LlenaGrilla ""
    If flxTramite.Rows > 1 Then
        DesplazarFlx
        flxTramite.row = 1
        flxTramite.ColSel = 17
    End If
    Limpiar
End Sub
Private Sub Limpiar()
    Dim I%, J%
    For I = 1 To 17
        For J = 1 To 2
            filtro(I, J) = Empty
        Next
    Next
End Sub
Private Sub ConfigColores()
    Dim SQL As String
    Dim rscolores As MYSQL_RS
    Dim I%, J%
    colores
    SQL = "Select coddoc from cndocum"
    Set rscolores = oConexion.EjecutaSelectRS(SQL)
    J = 0
    If Not rscolores.EOF Then
        For I = 1 To rscolores.RecordCount
            DocColor(I, 1) = rscolores.Fields("Coddoc")
            If J > 7 Then
               J = 1
            Else
               J = J + 1
            End If
            DocColor(I, 2) = Color(J)
            rscolores.MoveNext
        Next
    End If
    Set rscolores = Nothing
End Sub
Private Sub colores()
    Color(1) = &HFFFFC0 'Celeste
    Color(2) = &HC0FFFF 'Amarillo
    Color(3) = &HC0C0FF 'rosado
    Color(4) = &HC0FFC0 'verde
    Color(5) = &HC0E0FF 'naranja
    Color(6) = &HFFC0C0 'moradito
    Color(7) = &HFFFFFF 'blanco
    Color(8) = &HE0E0E0 'gris
    Color(9) = &HFFC0FF 'lila
End Sub
Private Sub GridColor(TipoDoc As String, Color As String)
    With flxColores
        If (.Rows = 1 Or .Rows > 1) Then
            If .TextMatrix(.Rows - 1, 0) <> Empty Then
                If Trim(.TextMatrix(.Rows - 1, 0)) <> DescripcionesdeCodigos("CNDOCUM", TipoDoc, "1") Then
                    .Rows = .Rows + 1
                    .Col = 0
                    .row = .Rows - 1
                    .CellBackColor = Color
                    .ColWidth(0) = 3000
                    .TextMatrix(.Rows - 1, 0) = DescripcionesdeCodigos("CNDOCUM", TipoDoc, "1")
                    .Col = 1
                    .ColWidth(1) = 300
                    .row = .Rows - 1
                    .CellBackColor = Color
                    .CellFontName = "Wingdings"
                    .CellFontSize = 11
                    .Text = strUnChecked
                    .Col = 2
                    .ColWidth(2) = 0
                    .row = .Rows - 1
                    .TextMatrix(.Rows - 1, 2) = TipoDoc
                End If
            Else
                .Col = 0
                .row = 0
                .CellBackColor = Color
                .ColWidth(0) = 3000
                .TextMatrix(.Rows - 1, 0) = DescripcionesdeCodigos("CNDOCUM", TipoDoc, "1")
                .Col = 1
                .ColWidth(1) = 300
                .row = 0
                .CellBackColor = Color
                .CellFontName = "Wingdings"
                .CellFontSize = 11
                .Text = strUnChecked
                .Col = 2
                .ColWidth(2) = 0
                .row = 0
                .TextMatrix(.Rows - 1, 2) = TipoDoc
            End If
        End If
    End With
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flxTramite
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
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
End Sub
