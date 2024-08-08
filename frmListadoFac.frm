VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmListadoFac 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Facturación"
   ClientHeight    =   6540
   ClientLeft      =   2835
   ClientTop       =   1920
   ClientWidth     =   11385
   Icon            =   "frmListadoFac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   11355
      Begin Proyecto1.chameleonButton btnGenerar 
         Height          =   345
         Left            =   8550
         TabIndex        =   12
         ToolTipText     =   "Guardar"
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         BTYPE           =   14
         TX              =   "&Generar Asiento"
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
         MICON           =   "frmListadoFac.frx":014A
         PICN            =   "frmListadoFac.frx":0166
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
         Left            =   10800
         TabIndex        =   13
         Top             =   240
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
         MICON           =   "frmListadoFac.frx":09F8
         PICN            =   "frmListadoFac.frx":0A14
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDocs 
         BackStyle       =   0  'Transparent
         Caption         =   "20"
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
         Height          =   225
         Left            =   6960
         TabIndex        =   14
         Top             =   300
         Width           =   645
      End
      Begin MSForms.ComboBox cboEstado 
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2990;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         Height          =   285
         Left            =   3990
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Documento"
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
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   1995
      End
      Begin MSForms.ComboBox cboTipDoc 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2990;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   11355
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexListado 
         Height          =   5085
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   8969
         _Version        =   393216
         BackColor       =   -2147483624
         BackColorSel    =   8421631
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Height          =   765
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   11355
      Begin VB.Label lblEstado 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
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
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   450
         Width           =   5565
      End
      Begin VB.Line Line2 
         X1              =   30
         X2              =   11340
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Line Line1 
         X1              =   5790
         X2              =   5790
         Y1              =   150
         Y2              =   720
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   5565
      End
      Begin VB.Label lblDivi 
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
         ForeColor       =   &H008080FF&
         Height          =   225
         Left            =   5910
         TabIndex        =   9
         Top             =   450
         Width           =   5415
      End
      Begin VB.Label lblCencos 
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Costo"
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
         Height          =   225
         Left            =   5910
         TabIndex        =   8
         Top             =   150
         Width           =   5385
      End
   End
End
Attribute VB_Name = "frmListadoFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagFiltro As Boolean

Public Sub ConfigGrilla()
    Dim I As Integer
    With flexListado
        .Clear
        .Rows = 1
        .Cols = 19
        .ColWidth(0) = 400
        .TextMatrix(0, 0) = Space(1) + "It."
        .FixedCols = 1
        
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = "Id"
        
        .ColWidth(2) = 300
        .TextMatrix(0, 2) = "TD"
        
        .ColWidth(3) = 1400
        .TextMatrix(0, 3) = Space(7) + "N°Doc."
        
        .ColWidth(4) = 0
        .TextMatrix(0, 4) = "ccHFM."
    
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = Space(1) + "Cencos"
        
        .ColWidth(6) = 0
        .TextMatrix(0, 6) = "A"
        
        .ColWidth(7) = 1100
        .TextMatrix(0, 7) = Space(8) + "C.Aux"
        
        .ColWidth(8) = 1000
        .TextMatrix(0, 8) = Space(3) + "Fec_Emi"
        '.ColAlignment(8) = vbCenter
        
        .ColWidth(9) = 300
        .TextMatrix(0, 9) = "M"
        
        .ColWidth(10) = 1000
        .TextMatrix(0, 10) = Space(5) + "Sub-Total"
        
        .ColWidth(11) = 800
        .TextMatrix(0, 11) = "I.G.V."
        
        .ColWidth(12) = 1000
        .TextMatrix(0, 12) = "Total"
        
        .ColWidth(13) = 300
        .TextMatrix(0, 13) = "D"
        
        .ColWidth(14) = 500
        .TextMatrix(0, 14) = "Cont"
    
        .ColWidth(15) = 1000
        .TextMatrix(0, 15) = "Fec. Vouc"
    
        .ColWidth(16) = 800
        .TextMatrix(0, 16) = "Voucher"
        
        .ColWidth(17) = 0
        .TextMatrix(0, 17) = "Estado"
        
        .ColWidth(18) = 200
        .TextMatrix(0, 18) = "Ref"
    End With
End Sub

Sub Contabiliza(Iden As String, vou As String)
    Dim SQL As String
    SQL = "Update documento_contables set fec_contabilizada='" & CStr(Format(Date, "yyyy/mm/dd")) & "'," & _
          " voucher='" & CStr(Trim(vou)) & "' where identificador='" & Iden & "'"
    oConexionMYSQL.Execute (SQL)
End Sub

Sub LlenarTipDoc()
    cboTipDoc.Clear
    cboTipDoc.AddItem "Todos...", 0
    cboTipDoc.List(0, 2) = ""
    cboTipDoc.AddItem "Factura", 1
    cboTipDoc.List(1, 2) = "01"
    cboTipDoc.AddItem "Boleta", 2
    cboTipDoc.List(2, 2) = "03"
    cboTipDoc.AddItem "Nota de Crédito", 3
    cboTipDoc.List(3, 2) = "07"
    cboTipDoc.AddItem "Nota de Débito", 4
    cboTipDoc.List(4, 2) = "08"
    cboTipDoc.AddItem "Proforma", 5
    cboTipDoc.List(5, 2) = "P"
End Sub

Sub CargarDocs()
    Dim SQL As String
    Dim rsDocumentos As MYSQL_RS
    Dim I As Integer
    lblDocs = "0"
        SQL = "Select distinct b.identificador,b.cod_tipo_doc,a.cod_estado,c.division,c.mon,concat(c.serie,'-',c.correl) as numdoc,TRIM(c.cenco) AS CENCO,c.auxiliar,c.codigo,c.fec_emision," & _
            " c.subtotal,c.igv,c.total,c.fec_contabilizada,c.voucher,d.afecto,c.ref " & _
            " from ((movi_documento as a left join amarre_documento as b " & _
            " on (a.identificador=b.identificador)) left join documento_contables as c on (b.identificador=c.identificador)) left join detallefact as d on (b.identificador=d.identificador) " & _
            " where a.cod_estado<>'" & ELIMINADO & "' and  c.auxiliar='2' and  b.flag='0' and b.cod_tipo_doc<>'O' and left(c.fec_emision,7)='" & strAnoSistema & "/" & strMesSistema & "'"
        If cboTipDoc.ListIndex > 0 Then
            SQL = SQL & " and b.cod_tipo_doc='" & cboTipDoc.List(cboTipDoc.ListIndex, 2) & "' "
        End If
        If cboEstado.ListIndex > 0 Then
            SQL = SQL & " and a.cod_estado='" & cboEstado.List(cboEstado.ListIndex, 2) & "' "
        End If
         
        If strUsuarioId = "JHERRERA" Then ' Caso de Downhole
            SQL = SQL & " and c.division in ('0003','0010','0011','0014','0016','0017','2001','9162') "
        End If
         'SQL = SQL & " order by c.serie,c.correl"
        SQL = SQL & " order by c.fec_emision,c.serie,c.correl"
    Set rsDocumentos = oConexion.EjecutaSelectRS(SQL)
    With flexListado
        ConfigGrilla
        .FixedRows = 0
        .ForeColorFixed = vbRed
        .Visible = False
        Do While Not (rsDocumentos.EOF)
            I = I + 1
            .Rows = .Rows + 1
            If I = 1 Then flexListado.FixedRows = 1
            .TextMatrix(I, 0) = CStr(I)
            .TextMatrix(I, 1) = rsDocumentos.Fields("IDENTIFICADOR")
            .TextMatrix(I, 2) = rsDocumentos.Fields("COD_TIPO_DOC")
            .TextMatrix(I, 3) = rsDocumentos.Fields("NUMDOC")
            .TextMatrix(I, 4) = rsDocumentos.Fields("DIVISION")
            .TextMatrix(I, 5) = rsDocumentos.Fields("CENCO")
            .TextMatrix(I, 6) = rsDocumentos.Fields("AUXILIAR")
            .TextMatrix(I, 7) = rsDocumentos.Fields("CODIGO")
            .TextMatrix(I, 8) = Format(CDate(rsDocumentos.Fields("FEC_EMISION")), "dd/mm/yyyy")
            .TextMatrix(I, 9) = rsDocumentos.Fields("MON")
            .TextMatrix(I, 10) = FormatNumber(rsDocumentos.Fields("SUBTOTAL"), 2)
            .TextMatrix(I, 11) = FormatNumber(rsDocumentos.Fields("IGV"), 2)
            .TextMatrix(I, 12) = FormatNumber(rsDocumentos.Fields("TOTAL"), 2)
            
            .row = I
            .Col = 13
            .CellFontSize = 8
            If rsDocumentos.Fields("AFECTO") = "0" Then
                .TextMatrix(I, 13) = "N"
            Else
                .CellFontBold = True
                .CellForeColor = vbRed
                .TextMatrix(I, 13) = "S"
            End If
            .row = I
            .Col = 14
            .CellFontBold = True
            .CellFontSize = 8
            If rsDocumentos.Fields("VOUCHER") = "" Then
                .CellForeColor = vbBlue
                .CellFontName = "Wingdings"
                .CellFontSize = 13
                .TextMatrix(I, 14) = strUnChecked
            Else
                .TextMatrix(I, 14) = "S"
            End If
            If rsDocumentos.Fields("FEC_CONTABILIZADA") <> "" Then
                .TextMatrix(I, 15) = Format(CDate(rsDocumentos.Fields("FEC_CONTABILIZADA")), "dd/mm/yyyy")
            Else
                .TextMatrix(I, 15) = ""
            End If
            .TextMatrix(I, 16) = rsDocumentos.Fields("VOUCHER")
            .TextMatrix(I, 17) = rsDocumentos.Fields("COD_ESTADO")
            .TextMatrix(I, 18) = rsDocumentos.Fields("REF")
            rsDocumentos.MoveNext
        Loop
        If I > 0 Then
            .row = 1
            .Col = 1
            .Rowsel = 1
            .ColSel = 16
        End If
        .Visible = True
    End With
    lblDocs = str(I)
    Set rsDocumentos = Nothing
End Sub

Private Sub BtnGenerar_Click()
    Dim I As Integer, J As Integer, SQL As String
    Dim v As String, lib As String, glo As String, Serdoc As String, Numdoc As String, fec As String, AnoMes As String, det As String
    Dim tc As Double, td As String, Div As String, DivServ As String, mon As String, aux As String, caux As String, cenco As String, Cta As String, cto As String, dh As String
    Dim sol As Double, dol As Double, colv As String, correl As String, Clasf As String, At As String
    Dim rsservicio As MYSQL_RS
    Dim rsDocumentos As MYSQL_RS
    Dim VoucRec As String
    Dim AnoMesRec As String
    Dim VsolRef As Double, VdolRef As Double
     
    With flexListado
        For I = 1 To .Rows - 1
            If .TextMatrix(I, 14) = strChecked Then
                .row = I
                If .row > 0 Then
                    .Col = 1
                    .ColSel = 17
                End If
                'SendKeys "{HOME}+{END}"
                If .TextMatrix(I, 2) = "P" Then
                    lib = "08"
                Else
                    lib = "04"
                End If
                AnoMes = Right(Trim(.TextMatrix(I, 8)), 4) & Mid(Trim(.TextMatrix(I, 8)), 4, 2)
                v = MaxVoucher(AnoMes, lib)
                td = Trim(.TextMatrix(I, 2))
                
                'Aqui se inserta referencia
                If td = "07" Then
                 SQL = "select anomes,voucher,sum(cargos) as Vsol,sum(cargod) as Vdol from cnmovi where codlib='04' and serdoc=left('" & .TextMatrix(I, 18) & "',4) and numdoc=right('" & .TextMatrix(I, 18) & "',8)"
                 Set rsDocumentos = oConexion.EjecutaSelectRS(SQL)
                 Do While Not (rsDocumentos.EOF)
                  VoucRec = rsDocumentos.Fields("VOUCHER")
                  AnoMesRec = rsDocumentos.Fields("ANOMES")
                  VsolRef = rsDocumentos.Fields("Vsol")
                  VdolRef = rsDocumentos.Fields("Vdol")
                  rsDocumentos.MoveNext
                 Loop
                 Set rsDocumentos = Nothing
                End If
                'Fin inserta referencia
                
'                Serdoc = Mid(Trim(.TextMatrix(i, 3)), 3, 3)
'                Numdoc = Right(Trim(.TextMatrix(i, 3)), 6)
'
                If (lib = "04") Then
                 Serdoc = "F" & Mid(Trim(.TextMatrix(I, 3)), 3, 3)
                 Numdoc = Right(Trim(.TextMatrix(I, 3)), 8)
                Else
                 Serdoc = Mid(Trim(.TextMatrix(I, 3)), 3, 3)
                 Numdoc = Right(Trim(.TextMatrix(I, 3)), 6)
                End If
                
                If .TextMatrix(I, 3) = ANULADO Then
                    Select Case .TextMatrix(I, 2)
                        Case "01": glo = "F/." & Serdoc & "-" & Numdoc & " ANULADO"
                        Case "03": glo = "B/." & Serdoc & "-" & Numdoc & " ANULADO"
                        Case "07": glo = "NC/." & Serdoc & "-" & Numdoc & " ANULADO"
                        Case "P": glo = "P/." & Serdoc & "-" & Numdoc & " ANULADO"
                    End Select
                Else
                    Select Case .TextMatrix(I, 2)
                        Case "01": glo = "F/." & Serdoc & "-" & Numdoc & " " & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        Case "03": glo = "B/." & Serdoc & "-" & Numdoc & " " & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        Case "07": glo = "NC/." & Serdoc & "-" & Numdoc & " " & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        Case "P": glo = "P/." & Serdoc & "-" & Numdoc & " " & DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                    End Select
                End If
                
                fec = .TextMatrix(I, 8)
                TipoCambio (fec)
                tc = dblTipoCmbV
                det = Trim(.TextMatrix(I, 13))
                td = Trim(.TextMatrix(I, 2))
                Div = Trim(.TextMatrix(I, 4))
                mon = Trim(.TextMatrix(I, 9))
                aux = Trim(.TextMatrix(I, 6))
                caux = Trim(.TextMatrix(I, 7))
                cenco = Trim(.TextMatrix(I, 5))
                cencos = "0000"
                
                'Insertar referencia'
                If td = "07" Then
                   tc = Round(VsolRef / VdolRef, 3)
                   SQL = "Call cn_Insert_DocRef('" & AnoMes & "','" & v & "','" & AnoMesRec & "','" & VoucRec & _
                         " ','0001')"
                   oConexionMYSQL.Execute (SQL)
                End If
                'Fin Insertar referencia'
                 
                SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & fec & _
                      " ','" & fec & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
                      " ','CUADRADO','','','','','" & det & "','','')"
                oConexionMYSQL.Execute (SQL)
               
                '******************************************
                If .TextMatrix(I, 17) = ANULADO Then
                    Cta = "121301"
                    cto = "ANULADA"
                    colv = "01"
                    correl = "0001"
                    SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','0001','" & Trim(.TextMatrix(I, 1)) & "','" & _
                           v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                           aux & "','00000000006','" & cencos & "','" & cenco & "','N','" & _
                           cto & "',0.00,0.00,0.00,0.00,'" & _
                           fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','07','000','')"
                    oConexionMYSQL.Execute (SQL)
                    Contabiliza Trim(.TextMatrix(I, 1)), v
                    'Next i
                Else
                    If .TextMatrix(I, 2) = "07" Then
                        '*****************************************************
                        SQL = " SELECT a.item,a.identificador,a.codser,b.codclasf,(select descripcioncorta from novperuvhse.lote where idlote=a.lote) as lote,(select descripcioncorta from novperuvhse.pozo where idpozo=a.pozo) as pozo,a.descripcion,a.cantidad,a.monto,a.total,(select CODDIV from servicio where codigo=a.codser) as divix" & _
                               " FROM detallefact as a left join servicio as b on(a.codser=b.codigo) " & _
                               " WHERE a.identificador = '" & .TextMatrix(I, 1) & "' ORDER BY b.codclasf "
                        Set rsservicio = oConexion.EjecutaSelectRS(SQL)
                        Clasf = ""
                        Div = Trim(.TextMatrix(I, 4))
                        J = 1
                        Do While Not (rsservicio.EOF)
                            If Clasf <> rsservicio.Fields("CODCLASF") Then
                                Clasf = rsservicio.Fields("CODCLASF")
                                If rsservicio.Fields("CODCLASF") = "10" Then
                                    Cta = "759901"
                                Else
                                    Cta = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "CUENTA")
                                    If Cta = "0000000" Then Cta = "70101"
                                End If
                                cto = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "DESCRIP")
                                aux = Left(DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "AUXCONT"), 1)
                                caux = Right(DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "AUXCONT"), 11)
                                At = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "ATCONT")
                                dh = "D"
                                sol = Round(CDbl(rsservicio.Fields("TOTAL")) * tc, 2)
                                dol = Round(CDbl(rsservicio.Fields("TOTAL")), 2)
                                colv = "01"
                                correl = Right("0000" & Trim(CStr(J)), 4)
                                DivServ = rsservicio.Fields("DIVIX")
                                
                                SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & DivServ & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                                       v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & _
                                       "','" & aux & "','" & caux & "','" & At & "','" & cenco & "','N','" & _
                                       cto & "'," & _
                                       IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                                       IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                                       fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                                       colv & "','000','')"
                                J = J + 1
                                oConexionMYSQL.Execute (SQL)
                            Else
                                sol = Round(CDbl(rsservicio.Fields("TOTAL")) * tc, 2)
                                dol = Round(CDbl(rsservicio.Fields("TOTAL")), 2)
                                SQL = "Update cnmovi set abonos=abonos+" & sol & ",abonod=abonod+" & dol & _
                                    " where anomes='" & AnoMes & "' and voucher='" & v & "' and correl='" & correl & "'"
                                oConexionMYSQL.Execute (SQL)
                            End If
                            rsservicio.MoveNext
                        Loop
                        '***************************************
                        Cta = "401111"
                        cto = DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        Div = "013100003836" '0001 IIf(Trim(.TextMatrix(i, 4)) = "0003", "0003", "0001")
                        dh = "D"
                        sol = Round(CDbl(.TextMatrix(I, 11)) * tc, 2)
                        dol = Round(CDbl(.TextMatrix(I, 11)), 2)
                        colv = "02"
                        correl = Right("0000" & Trim(CStr(J)), 4)
                        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                               v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','0" & _
                               "','00000000000','0000','" & cenco & "','N','" & _
                               cto & "'," & _
                               IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                               IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                               fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                               colv & "','000','')"
                        oConexionMYSQL.Execute (SQL)
                        
                        '**********************************************
                        Cta = IIf(mon = "N", "121302", IIf(td = "P", "121101", "121301"))
                        correl = Right("0000" & Trim(CStr(J + 1)), 4)
                        sol = Round(CDbl(.TextMatrix(I, 12)), 2) * tc
                        dol = Round(CDbl(.TextMatrix(I, 12)), 2)
                        dh = "H"
                        cto = DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        colv = "07"
                        aux = Trim(.TextMatrix(I, 6))
                        caux = Trim(.TextMatrix(I, 7))
                        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                           v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                           aux & "','" & caux & "','" & cencos & "','" & cenco & "','N','" & _
                           cto & "'," & _
                           IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                           IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                           fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                           colv & "','000','')"
                        oConexionMYSQL.Execute (SQL)
                    
                    Else
                        '**********************************************
                        Div = ValidaDameCCHFMTubulares(Trim(.TextMatrix(I, 4)))
                        Cta = IIf(mon = "N", "121302", IIf(td = "P", "121101", "121301"))   '"12102" '"12101"
                        correl = "0001"
                        aux = Trim(.TextMatrix(I, 6))
                        caux = Trim(.TextMatrix(I, 7))
                        sol = Round(CDbl(.TextMatrix(I, 12)) * tc, 2)
                        dol = Round(CDbl(.TextMatrix(I, 12)), 2)
                        dh = "D"
                        cto = DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        colv = "07"
                        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                           v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                           aux & "','" & caux & "','" & cencos & "','" & cenco & "','N','" & _
                           cto & "'," & _
                           IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                           IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                           fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                           colv & "','000','')"
                        oConexionMYSQL.Execute (SQL)
                        
                        '**********************************************
                        Cta = "401111"
                        Div = "013100003836" '0001 IIf(Trim(.TextMatrix(i, 4)) = "0003", "0003", "0001")
                        cto = DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(I, 7)), Trim(.TextMatrix(I, 6)), "Descrip")
                        dh = "H"
                        sol = Round(CDbl(.TextMatrix(I, 11)) * tc, 2)
                        dol = Round(CDbl(.TextMatrix(I, 11)), 2)
                        colv = "02"
                        correl = "0002"
                        SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                               v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','0" & _
                               "','00000000000','0000','" & cenco & "','N','" & _
                               cto & "'," & _
                               IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                               IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                               fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                               colv & "','000','')"
                        oConexionMYSQL.Execute (SQL)
                        
                        Div = Trim(.TextMatrix(I, 4))
                         
                        '*****************************************************
                        SQL = " SELECT a.item,a.identificador,a.codser,b.codclasf,(select descripcioncorta from novperuvhse.lote where idlote=a.lote) as lote,(select descripcioncorta from novperuvhse.pozo where idpozo=a.pozo) as pozo,a.descripcion,a.cantidad,a.monto,a.total,(select CODDIV from servicio where codigo=a.codser) as divix" & _
                               " FROM detallefact as a left join servicio as b on(a.codser=b.codigo) " & _
                               " WHERE a.identificador = '" & .TextMatrix(I, 1) & "' and b.CODCLASF<>'00' ORDER BY b.codclasf "
                        Set rsservicio = oConexion.EjecutaSelectRS(SQL)
                        Clasf = ""
                        J = 3
                        Do While Not (rsservicio.EOF)
                            If Clasf <> rsservicio.Fields("CODCLASF") Then
                                Clasf = rsservicio.Fields("CODCLASF")
                                
                                If rsservicio.Fields("CODCLASF") = "10" Then
                                    Cta = "759901"
                                Else
                                    Cta = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "CUENTA")
                                    If Cta = "0000000" Then Cta = "70101"
                                End If
                                cto = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "DESCRIP")
                                aux = Left(DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "AUXCONT"), 1)
                                caux = Right(DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "AUXCONT"), 11)
                                At = DescripcionesdeCodigos("CLASF_SERVICIO", rsservicio.Fields("CODCLASF"), "", "ATCONT")
                                dh = "H"
                                sol = Round(CDbl(rsservicio.Fields("TOTAL")) * tc, 2)
                                dol = Round(CDbl(rsservicio.Fields("TOTAL")), 2)
                                colv = "01"
                                correl = Right("0000" & Trim(CStr(J)), 4)
                                DivServ = rsservicio.Fields("DIVIX")
                                
                                SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & DivServ & "','" & Trim(.TextMatrix(I, 1)) & "','" & _
                                       v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & _
                                       "','" & aux & "','" & caux & "','" & At & "','" & cenco & "','N','" & _
                                       cto & "'," & _
                                       IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                                       IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                                       fec & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                                       colv & "','000','')"
                                J = J + 1
                                oConexionMYSQL.Execute (SQL)
                            Else
                                sol = Round(CDbl(rsservicio.Fields("TOTAL")) * tc, 2)
                                dol = Round(CDbl(rsservicio.Fields("TOTAL")), 2)
                                SQL = "Update cnmovi set abonos=abonos+" & sol & ",abonod=abonod+" & dol & _
                                    " where anomes='" & AnoMes & "' and voucher='" & v & "' and correl='" & correl & "'"
                                oConexionMYSQL.Execute (SQL)
                            End If
                            rsservicio.MoveNext
                        Loop
                    End If
                    Contabiliza Trim(.TextMatrix(I, 1)), v
                End If
                If Trim(.TextMatrix(I, 2)) = "01" Then
                    GenerarExtornos I
                End If
            End If
        Next
    End With
    MsgBox "Asientos Terminados", vbOKOnly, "NOVADMIN"
    cboTipDoc_Change
End Sub

Private Sub cboEstado_Change()
    If FlagFiltro = True Then CargarDocs
End Sub

Private Sub cboTipDoc_Change()
    LlenarEstados
End Sub

Private Sub chBtnSalir_Click()
     Unload Me
End Sub

Private Sub flexListado_Click()
    Dim SCol As Integer
    Dim rsDetfac As MYSQL_RS
    With flexListado
        SCol = .Col
        If .row > 0 Then
            .Col = 1
            .ColSel = 17
        End If
        If SCol = 14 And .TextMatrix(.row, 14) <> "S" Then
            If .TextMatrix(.row, 14) = strChecked Then
                .TextMatrix(.row, 14) = strUnChecked
            Else
                .TextMatrix(.row, 14) = strChecked
            End If
        End If
    End With
End Sub

Private Sub flexListado_RowColChange()
    If flexListado.Rows > 1 Then
        NavegarGrilla flexListado.row
    End If
End Sub

Private Sub Form_Load()
    Call WheelHook(frmListadoFac)
    Me.Left = 0
    Me.Top = 0
    LlenarTipDoc
    cboTipDoc.ListIndex = 0
End Sub

Sub LlenarEstados()
    Dim SQL As String
    Dim rsdocs As MYSQL_RS
    Dim I As Integer
    I = 0
    
    If cboTipDoc.ListIndex > 0 Then
        SQL = "Select c.descripcion,a.cod_estado " & _
            " from (movi_documento as a left join amarre_documento as b " & _
            " on (a.identificador=b.identificador)) left join doc_estado as c on (a.cod_estado=c.cod_estado) " & _
            " where b.flag='0' and b.cod_tipo_doc<>'O' and b.anomes='" & strAnoSistema & strMesSistema & "' and " & _
            " b.cod_tipo_doc='" & cboTipDoc.List(cboTipDoc.ListIndex, 2) & "' " & _
            " group by c.descripcion,a.cod_estado order by c.descripcion,a.cod_estado"
    Else
        SQL = "Select c.descripcion,a.cod_estado " & _
            " from (movi_documento as a left join amarre_documento as b " & _
            " on (a.identificador=b.identificador)) left join doc_estado as c on (a.cod_estado=c.cod_estado) " & _
            " where b.flag='0' and b.cod_tipo_doc<>'O' and b.anomes='" & strAnoSistema & strMesSistema & "' " & _
            " group by c.descripcion,a.cod_estado order by c.descripcion,a.cod_estado"
    End If
    Set rsdocs = oConexion.EjecutaSelectRS(SQL)
    FlagFiltro = False
    cboEstado.Clear
    cboEstado.AddItem "Todos...", I
    cboEstado.List(I, 2) = ""
    I = I + 1
    Do While Not (rsdocs.EOF)
        cboEstado.AddItem rsdocs.Fields("descripcion"), I
        cboEstado.List(I, 2) = rsdocs.Fields("cod_estado")
        I = I + 1
        rsdocs.MoveNext
    Loop
    FlagFiltro = True
    If I > 1 Then
        cboEstado.ListIndex = 0
        cboEstado.Enabled = True
        cboEstado.BackColor = ColorHabilitado
    Else
        cboEstado.Enabled = False
        cboEstado.BackColor = ColorDeshabilitado
        flexListado.Rows = 1
    End If
    Set rsdocs = Nothing
End Sub

Sub NavegarGrilla(fila As Integer)
    With flexListado
        lblCliente = DescripcionesdeCodigos("AUXILIARES", .TextMatrix(fila, 7), "2", "Descrip")
        lblEstado = DescripcionesdeCodigos("DOC_ESTADO", .TextMatrix(fila, 17), "", "")
        lblcencos = DescripcionesdeCodigos("CENCO", .TextMatrix(fila, 5), "1", "")
        lblDivi = DescripcionesdeCodigos("DES_DIVISION", .TextMatrix(fila, 4), "", "")
    End With
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
    On Error Resume Next
    
    With flexListado
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
        If NewValue = 0 Then .TopRow = 1
    End With
End Sub

Sub GenerarExtornos(fila As Integer)
On Error GoTo CtrlError
    Dim AnoMes As String
    Dim lib As String
    Dim v As String, Cta As String, mon As String, cencos As String, tipovoc As String
    Dim Serdoc As String, Numdoc As String, Div As String, DivServ As String, cenco As String, correl As String
    Dim glo As String, aux As String, caux As String, prof As String
    Dim tc As Double
    Dim I As Integer, k As Integer
    Dim SumaD As Double, SumaS As Double, SIgvS As Double, SigvD As Double
    Dim SQL As String
    Dim RServ As MYSQL_RS
    Dim RQ As MYSQL_RS
    
    With flexListado
        SQL = "Select * from factura_proforma where identificador= '" & Trim(.TextMatrix(fila, 1)) & "'"
        Set RQ = oConexion.EjecutaSelectRS(SQL)
        If Not RQ.EOF() Then
            lib = "08" 'Libro donde se van a generar las reversiones
            AnoMes = Right(Trim(.TextMatrix(fila, 8)), 4) & Mid(Trim(.TextMatrix(fila, 8)), 4, 2)
            v = MaxVoucher(AnoMes, lib)
            Serdoc = Mid(Trim(.TextMatrix(fila, 3)), 3, 3)
            Numdoc = Right(Trim(.TextMatrix(fila, 3)), 6)
            glo = "REVERSION DE PROFORMAS " & Mid(RQ.Fields("prof"), 3, 3) & " / FAC. " & Serdoc & "-" & Numdoc
            TipoCambio (.TextMatrix(fila, 8))
            tc = dblTipoCmbV
            td = "P"
            mon = .TextMatrix(fila, 9)
            cenco = Trim(.TextMatrix(fila, 5))
            cencos = "0000"
            Div = Trim(.TextMatrix(fila, 4))
            colv = ""
            
            SQL = "Call cn_Insert_Voucher('" & lib & "','" & v & "','" & glo & "','" & .TextMatrix(fila, 8) & _
                  " ','" & .TextMatrix(fila, 8) & "','V'," & tc & ",'" & mon & "','" & AnoMes & "','" & strUsuarioId & _
                  " ','CUADRADO','','','','" & Mid(RQ.Fields("prof"), 3, 4) & Right(RQ.Fields("prof"), 6) & "','N','','P')"
            oConexionMYSQL.Execute (SQL)
                
            k = 1
            prof = RQ.Fields("prof")
            Do While Not RQ.EOF()
                SQL = "SELECT codclasf,coddiv  From servicio Where CODIGO = '" & Trim(RQ.Fields("codserv")) & "'"
                Set RServ = oConexion.EjecutaSelectRS(SQL)
                If Not RServ.EOF() Then
                    Cta = DescripcionesdeCodigos("CLASF_SERVICIO", RServ.Fields("CODCLASF"), "", "CUENTA")
                    If Cta = "0000000" Then Cta = "70101"
                    cto = DescripcionesdeCodigos("CLASF_SERVICIO", RServ.Fields("CODCLASF"), "", "DESCRIP")
                    aux = Left(DescripcionesdeCodigos("CLASF_SERVICIO", RServ.Fields("CODCLASF"), "", "AUXCONT"), 1)
                    caux = Right(DescripcionesdeCodigos("CLASF_SERVICIO", RServ.Fields("CODCLASF"), "", "AUXCONT"), 11)
                    At = DescripcionesdeCodigos("CLASF_SERVICIO", RServ.Fields("CODCLASF"), "", "ATCONT")
                    dh = "D"
                    sol = Round(CDbl(FormatNumber(RQ.Fields("valor"), 2)) * tc, 2)
                    dol = Round(CDbl(FormatNumber(RQ.Fields("valor"), 2)), 2)
                    correl = Right("0000" & Trim(CStr(k)), 4)
                    Serdoc = Mid(Trim(RQ.Fields("prof")), 3, 3)
                    Numdoc = Right(Trim(RQ.Fields("prof")), 6)
                    DivServ = RServ.Fields("CODDIV")
                    SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & DivServ & "','" & Trim(.TextMatrix(fila, 1)) & "','" & _
                          v & "','" & Serdoc & "','" & Numdoc & "','" & correl & "','" & mon & "','" & Trim(Cta) & _
                          "','" & aux & "','" & caux & "','" & At & "','" & cenco & "','N','" & _
                          cto & "'," & _
                          IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                          IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                          Trim(.TextMatrix(fila, 8)) & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                          colv & "','000','')"
                    oConexionMYSQL.Execute (SQL)
                End If
                SumaD = SumaD + dol
                SumaS = SumaS + sol
                RQ.MoveNext
                If prof <> RQ.Fields("prof") Then
                    '**********************************************
                    k = k + 1
                    Cta = "401111"
                    dh = "D"
                    SIgvS = SumaS * 0.18
                    SigvD = SumaD * 0.18
                    Div = "013100003836" '0001 IIf(Trim(.TextMatrix(fila, 4)) = "0003", "0003", "0001")
                    sol = Round(CDbl(SIgvS), 2)
                    dol = Round(CDbl(SigvD), 2)
                    correl = Right("0000" & Trim(CStr(k)), 4)
                    
                    SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(fila, 1)) & "','" & _
                          v & "','" & Trim(Serdoc) & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','0" & _
                          "','00000000000','0000','" & cenco & "','N','" & _
                          cto & "'," & _
                          IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                          IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                          Trim(.TextMatrix(fila, 8)) & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                          colv & "','000','')"
                    oConexionMYSQL.Execute (SQL)
                    '**********************************************
                    k = k + 1
                    Cta = IIf(mon = "N", "121302", IIf(td = "P", "121101", "121301"))   '"12102" '"12101"
                    correl = Right("0000" & Trim(CStr(k)), 4)
                    sol = Round(CDbl(SumaS + SIgvS), 2)
                    dol = Round(CDbl(SumaD + SigvD), 2)
                    dh = "H"
                    aux = 2
                    
                    'Div = ValidaDameCCHFMTubulares(Trim(.TextMatrix(fila, 4)))
                    Div = ValidaDameCCHFMTubulares(Trim(DivServ))
                    
                    caux = Trim(.TextMatrix(fila, 7))
                    cto = DescripcionesdeCodigos("AUXILIARES", Trim(.TextMatrix(fila, 7)), 2, "Descrip")
                    SQL = "call cn_Insert_Movi ('" & lib & "','" & td & "','" & Div & "','" & Trim(.TextMatrix(fila, 1)) & "','" & _
                          v & "','" & Trim(Serdoc) & "','" & Trim(Numdoc) & "','" & correl & "','" & mon & "','" & Trim(Cta) & "','" & _
                          aux & "','" & caux & "','" & cencos & "','" & cenco & "','N','" & _
                          cto & "'," & _
                          IIf(dh = "D", sol, 0) & "," & IIf(dh = "H", sol, 0) & "," & _
                          IIf(dh = "D", dol, 0) & "," & IIf(dh = "H", dol, 0) & ",'" & _
                          Trim(.TextMatrix(fila, 8)) & "','" & AnoMes & "','" & strUsuarioId & "','" & dh & "','" & _
                          colv & "','000','')"
                    oConexionMYSQL.Execute (SQL)
                    '**********************************************
                    SumaS = 0
                    SumaD = 0
                    SIgvS = 0
                    SigvD = 0
                    prof = RQ.Fields("prof")
                End If
                k = k + 1
            Loop
        End If
    End With
Exit Sub
CtrlError:
    MsgBox err.Description, vbCritical, "Error Generando Extornos"
End Sub


Function ValidaDameCCHFMTubulares(ByVal vdivi As String) As String
    Dim SQL As String
    Dim RQ As MYSQL_RS
    ValidaDameCCHFMTubulares = ""
    SQL = " Select atipo from cnmdepar where coddep= '" & vdivi & "' and atipo='TUBULAR SERVICES' "
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        ValidaDameCCHFMTubulares = "013100003841"
    Else
        ValidaDameCCHFMTubulares = vdivi
    End If
    Set RQ = Nothing
End Function

