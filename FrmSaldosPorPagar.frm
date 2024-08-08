VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form FrmSaldosPorPagar 
   BackColor       =   &H009F5539&
   Caption         =   "Listado Saldos por Pagar"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   20760
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboNacExt 
      Height          =   315
      ItemData        =   "FrmSaldosPorPagar.frx":0000
      Left            =   18330
      List            =   "FrmSaldosPorPagar.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4200
      Width           =   1575
   End
   Begin VB.ComboBox cboauxiliar 
      Height          =   315
      Left            =   18300
      TabIndex        =   13
      Top             =   3420
      Width           =   2145
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexListado 
      Height          =   6255
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   18075
      _ExtentX        =   31882
      _ExtentY        =   11033
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
   Begin Proyecto1.chameleonButton BtnGenerar 
      Height          =   615
      Left            =   18390
      TabIndex        =   3
      ToolTipText     =   "Eliminar"
      Top             =   330
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Generar Pagos"
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
      MICON           =   "FrmSaldosPorPagar.frx":0024
      PICN            =   "FrmSaldosPorPagar.frx":0040
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox mskFecha1 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   6480
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton chameleonButton1 
      Height          =   525
      Left            =   18360
      TabIndex        =   7
      ToolTipText     =   "Ver Programacion Pagos"
      Top             =   6330
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   926
      BTYPE           =   14
      TX              =   "&Actualizar"
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
      MICON           =   "FrmSaldosPorPagar.frx":23C2
      PICN            =   "FrmSaldosPorPagar.frx":23DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox mskFecha2 
      Height          =   315
      Left            =   4140
      TabIndex        =   8
      Top             =   6450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton btnInterfaz 
      Height          =   345
      Left            =   15690
      TabIndex        =   11
      Top             =   6450
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
      MICON           =   "FrmSaldosPorPagar.frx":4760
      PICN            =   "FrmSaldosPorPagar.frx":477C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   8370
      Width           =   20445
      VariousPropertyBits=   746604571
      MaxLength       =   800
      Size            =   "36063;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda"
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
      Left            =   18330
      TabIndex        =   16
      Top             =   3930
      Width           =   1995
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Auxiliar"
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
      Left            =   18330
      TabIndex        =   14
      Top             =   3150
      Width           =   1995
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar Todo"
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
      Left            =   16230
      TabIndex        =   12
      Top             =   6570
      Width           =   1995
   End
   Begin MSForms.TextBox meProveedor 
      Height          =   315
      Left            =   8310
      TabIndex        =   10
      Top             =   6450
      Width           =   1755
      VariousPropertyBits=   746604571
      MaxLength       =   30
      Size            =   "3096;556"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
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
      Left            =   6150
      TabIndex        =   9
      Top             =   6480
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "hasta"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   6570
      Width           =   525
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rango de Fechas"
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
      TabIndex        =   5
      Top             =   6510
      Width           =   1995
   End
   Begin MSForms.ComboBox cboTipDoc 
      Height          =   285
      Left            =   2190
      TabIndex        =   2
      Top             =   6990
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
      Left            =   120
      TabIndex        =   1
      Top             =   6990
      Width           =   1995
   End
End
Attribute VB_Name = "FrmSaldosPorPagar"
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
        .Cols = 21
        .ColWidth(0) = 500
        .CellFontSize = 9
        
        .TextMatrix(0, 0) = Space(1) + "It."
        .FixedCols = 1
        
        .ColWidth(1) = 1500
        .CellFontSize = 9
        .TextMatrix(0, 1) = "Identificador"
        
        .ColWidth(2) = 900
        .CellFontSize = 9
        .TextMatrix(0, 2) = "S°Doc."
        
        .ColWidth(3) = 1600
        .CellFontSize = 9
        .TextMatrix(0, 3) = "N°Doc."
    
        .ColWidth(4) = 300
        .CellFontSize = 9
        .TextMatrix(0, 4) = Space(1) + "Moneda"
        
        .ColWidth(5) = 1200
        .CellFontSize = 9
        .TextMatrix(0, 5) = "Total"
        
        .ColWidth(6) = 1200
        .CellFontSize = 9
        .TextMatrix(0, 6) = Space(8) + "Pendiente"
        
        .ColWidth(7) = 500
        .CellFontSize = 9
        .TextMatrix(0, 7) = Space(3) + "Auxiliar"
        
        .ColWidth(8) = 1000
        .CellFontSize = 9
        .TextMatrix(0, 8) = "Fec Pago"
        
        .ColWidth(9) = 1000
        .CellFontSize = 9
        .TextMatrix(0, 9) = Space(5) + "Equivalente"
        
        .ColWidth(10) = 1600
        .CellFontSize = 9
        .TextMatrix(0, 10) = "Ruc"
        
        .ColWidth(11) = 1500
        .CellFontSize = 9
        .TextMatrix(0, 11) = "CtaBco"
        
        .ColWidth(12) = 1200
        .CellFontSize = 9
        .TextMatrix(0, 12) = "C.Costo"
        
        .ColWidth(13) = 600
        .CellFontSize = 9
        .TextMatrix(0, 13) = "Orden"
    
        .ColWidth(14) = 300
        .CellFontSize = 9
        .TextMatrix(0, 14) = "Selec."
    
        .ColWidth(15) = 1000
        .CellFontSize = 9
        .TextMatrix(0, 15) = "FecVcto"
        
        .ColWidth(16) = 1300
        .CellFontSize = 9
        .TextMatrix(0, 16) = "Banco"
        
        .ColWidth(17) = 500
        .CellFontSize = 9
        .TextMatrix(0, 17) = "TD"
        
        .ColWidth(18) = 1500
        .CellFontSize = 9
        .TextMatrix(0, 18) = "MPago"
        
        .ColWidth(19) = 1500
        .CellFontSize = 9
        .TextMatrix(0, 19) = "Nombre"
        
        .ColWidth(20) = 1500
        .CellFontSize = 9
        .TextMatrix(0, 20) = "Detracc"
    End With
End Sub

Private Sub BtnGenerar_Click()
    On Error GoTo NADA
    Dim I As Integer
    Dim num As Integer
    Dim Item As Integer
    Dim consultaLiq As String
    Dim consultaLiqRep As String
    
   'Tabla Temporal Liquidación/Aqui agregamos el historial de pagos asociados
    Item = 1
    num = 1
    consultaLiq = "delete from liquidpagos;"
    oConexion.EjecutaInsertUpdateDelete consultaLiq, TIPO_QUERY.insertar, False

    consultaLiq = ""
    consultaLiqRep = ""
    For I = 1 To flexListado.Rows - 1
        If flexListado.TextMatrix(I, 14) = strChecked Then
    
        consultaLiq = "Insert into liquidpagos(IDLIQ,codigo,auxiliar,nombre,fecha,importe,importeeqv,serdoc,numdoc,cuenta,usuario,orden,td,mon,folio,mpago,banco,ctabanco,divi,FAC_O) " & _
        "SELECT " & CInt("1") & " , '" & flexListado.TextMatrix(I, 10) & "' ,  '" & flexListado.TextMatrix(I, 7) & "' ,  '" & flexListado.TextMatrix(I, 19) & "' , '" & flexListado.TextMatrix(I, 8) & "' , " & _
        "" & CDbl(flexListado.TextMatrix(I, 6)) - CDbl(flexListado.TextMatrix(I, 20)) & " , " & flexListado.TextMatrix(I, 9) & " ,'" & flexListado.TextMatrix(I, 2) & "' ,'" & flexListado.TextMatrix(I, 3) & "', " & _
        "'421201' ," & _
        "'ADM','" & flexListado.TextMatrix(I, 13) & "','" & flexListado.TextMatrix(I, 17) & "','" & flexListado.TextMatrix(I, 4) & "','" & flexListado.TextMatrix(I, 1) & "','" & flexListado.TextMatrix(I, 18) & "','" & flexListado.TextMatrix(I, 16) & "','" & flexListado.TextMatrix(I, 11) & "' , " & _
        "'" & flexListado.TextMatrix(I, 12) & "'," & flexListado.TextMatrix(I, 6) & "; "
        oConexion.EjecutaInsertUpdateDelete consultaLiq, TIPO_QUERY.insertar, False
        
        
        consultaLiqRep = "Insert into liquidpagos_rep(IDLIQ,codigo,auxiliar,nombre,fecha,importe,importeeqv,serdoc,numdoc,cuenta,usuario,orden,td,mon,folio,mpago,banco,ctabanco,divi,FAC_O) " & _
        "SELECT '" & strAnoSistema & strMesSistema & frmMovTelewiese.meOrden.Text & "'  , '" & flexListado.TextMatrix(I, 10) & "' ,  '" & flexListado.TextMatrix(I, 7) & "' ,  '" & flexListado.TextMatrix(I, 19) & "' , '" & flexListado.TextMatrix(I, 8) & "' , " & _
        "" & CDbl(flexListado.TextMatrix(I, 6)) - CDbl(flexListado.TextMatrix(I, 20)) & " , " & flexListado.TextMatrix(I, 9) & " ,'" & flexListado.TextMatrix(I, 2) & "' ,'" & flexListado.TextMatrix(I, 3) & "', " & _
        "'421201' ," & _
        "'ADM','" & flexListado.TextMatrix(I, 13) & "','" & flexListado.TextMatrix(I, 17) & "','" & flexListado.TextMatrix(I, 4) & "','" & flexListado.TextMatrix(I, 1) & "','" & flexListado.TextMatrix(I, 18) & "','" & flexListado.TextMatrix(I, 16) & "','" & flexListado.TextMatrix(I, 11) & "' , " & _
        "'" & flexListado.TextMatrix(I, 12) & "'," & flexListado.TextMatrix(I, 6) & "; "
        oConexion.EjecutaInsertUpdateDelete consultaLiqRep, TIPO_QUERY.insertar, False
        
        Item = Item + 1
        End If
    Next
    
    MsgBox "Proceso OK, importe en la orden TBK correspondiente(Sistema Administrativo). ", vbOKOnly + vbExclamation, "NOV"
    Unload FrmSaldosPorPagar
    
    Call keybd_event(vbKeyEnd, 0, 0, 0)
    
    Exit Sub
NADA:
    Exit Sub
End Sub

Private Sub btnInterfaz_Click()
  
  With flexListado
   For I = 1 To flexListado.Rows - 1
      .TextMatrix(I, 14) = strChecked
   Next
  End With
  
End Sub

Private Sub chameleonButton1_Click()
Dim SQL As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String
Dim sql5 As String

Dim rsDocumentos As MYSQL_RS
Dim I As Integer
    lblDocs = "0"
    SQL = ""
    sql2 = ""
    sql3 = ""
    sql4 = ""
    sql5 = ""
        
        SQL = "Select d.identificador,d.serie, d.documento, d.mon,d.Total,d.pendiente,d.auxiliar,d.codigo,d.cenco,d.division,d.fec_pago , d.fec_emision, d.fec_vcto, d.impEqui,d.Cod_Tipo_Doc, d.cod_estado, c.orden,C.obs,C.Division,ifnull(O.MPago,ifnull((select descrip from tipopago where codpago=e.tipcta_mn),'')) as MPago,ifnull(left(O.CtaBco,3),ifnull(left(e.numcta_mn,3),'000')) as TxtOficina,ifnull(O.CtaBco,ifnull(e.numcta_mn,'0000000')) as CtaBco," & _
            "ifnull(O.Banco,ifnull((select descrip from pl_entidadfinanciera where codigo=e.codbanco),'')) as Banco,(select descrip from cnauxil where auxiliar=d.auxiliar and codigo=d.codigo) as nombre, Ifnull((select If(d.mon='N',Sum(c.cargos),Sum(c.cargod)) as Valor from cnmovi as c left join cnvouc as x on (c.anomes=x.anomes and c.voucher=x.voucher) where c.codaux=trim(d.codigo) and c.serdoc=trim(d.serie) and c.numdoc= trim(d.documento) and ((x.glosa like 'detracc%') or (x.glosa like 'reten%')or (x.glosa like 'antic%') or(x.glosa like 'aplic%'))),'0') as detracc  " & _
            "from ((DOC_PROG as d  left join documento_contables as c on (d.identificador=c.identificador)) left join  orden_compra as O on c.orden=O.correl) left join empleado as e on (e.codigo=d.codigo) where 1=1 "
'            left(d.identificador,4)='" & strAnoSistema & "'
        
'        If (mskFecha1.Text <> "") And (mskFecha2.Text <> "") And ((mskFecha1.Text <> "__/__/____") And (mskFecha2.Text <> "__/__/____")) Then
'          sql2 = " and concat(right('" & mskFecha1.Text & "',4),mid('" & mskFecha1.Text & "',3,4),left('" & mskFecha1.Text & "',2))<= d.fec_Pago and " & _
'                  "concat(right('" & mskFecha2.Text & "',4),mid('" & mskFecha2.Text & "',3,4),left('" & mskFecha2.Text & "',2))>= d.fec_Pago "
'        End If
        
        If mskFecha1.Text <> "__/__/____" And mskFecha2.Text <> "__/__/____" Then
                 sql2 = " and d.Fec_Pago>=concat(right('" & mskFecha1.Text & "',4),mid('" & mskFecha1.Text & "',3,4),left('" & mskFecha1.Text & "',2)) and d.Fec_Pago<=concat(right('" & mskFecha2.Text & "',4),mid('" & mskFecha2.Text & "',3,4),left('" & mskFecha2.Text & "',2))  "
                 
        Else
                If mskFecha1.Text <> "__/__/____" And mskFecha2.Text = "__/__/____" Then
                    sql2 = " and d.Fec_Pago>=concat(right('" & mskFecha1.Text & "',4),mid('" & mskFecha1.Text & "',3,4),left('" & mskFecha1.Text & "',2)) and d.Fec_Pago<>''  "
                    
                   
                Else
                    If mskFecha1.Text = "__/__/____" And mskFecha2.Text <> "__/__/____" Then
                        sql2 = " and d.Fec_Pago<=concat(right('" & mskFecha2.Text & "',4),mid('" & mskFecha2.Text & "',3,4),left('" & mskFecha2.Text & "',2)) and d.Fec_Pago<>''  "
                        
                    End If
                End If
        End If
        
        
        
        If (meProveedor.Text <> "") Then
           If Len(meProveedor.Text) > 10 Then
              sql3 = "and d.codigo='" & Trim(meProveedor.Text) & "' "
           Else
              sql3 = "and d.codigo='" & Trim(Right("00000000000" & meProveedor.Text, 11)) & "' "
           End If
        End If
        
       If Trim(Left(cboauxiliar.Text, InStr(1, cboauxiliar.Text, " "))) <> "" Then
              sql4 = " and left(d.auxiliar,1)='" & Trim(Left(cboauxiliar.Text, InStr(1, cboauxiliar.Text, " "))) & "' "
       End If
       
       If CboNacExt.Text <> "" Then
             sql5 = " and d.mon='" & Left(Trim(CboNacExt.Text), 1) & "' "
       End If
       
    SQL = SQL & sql2 & sql3 & sql4 & sql5 & " order by d.fec_Pago,d.serie, d.documento"
    
    TextBox1.Text = SQL
    
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
            .row = I
            .Col = 0
            .CellFontSize = 9
            .TextMatrix(I, 0) = CStr(I)
            .row = I
            .Col = 1
            .CellFontSize = 9
            .TextMatrix(I, 1) = rsDocumentos.Fields("IDENTIFICADOR")
            .row = I
            .Col = 2
            .CellFontSize = 9
            .TextMatrix(I, 2) = rsDocumentos.Fields("SERIE")
            .row = I
            .Col = 3
            .CellFontSize = 9
            .TextMatrix(I, 3) = rsDocumentos.Fields("DOCUMENTO")
            .row = I
            .Col = 4
            .CellFontSize = 9
            .TextMatrix(I, 4) = rsDocumentos.Fields("MON")
            .row = I
            .Col = 5
            .CellFontSize = 9
            .TextMatrix(I, 5) = rsDocumentos.Fields("TOTAL")
            .row = I
            .Col = 6
            .CellFontSize = 9
            .TextMatrix(I, 6) = rsDocumentos.Fields("PENDIENTE")
            .row = I
            .Col = 7
            .CellFontSize = 9
            .TextMatrix(I, 7) = rsDocumentos.Fields("AUXILIAR")
            .row = I
            .Col = 8
            .CellFontSize = 9
            .TextMatrix(I, 8) = Format(CDate(rsDocumentos.Fields("FEC_PAGO")), "dd/mm/yyyy")
            .row = I
            .Col = 9
            .CellFontSize = 9
            .TextMatrix(I, 9) = rsDocumentos.Fields("IMPEQUI")
            .row = I
            .Col = 10
            .CellFontSize = 9
            .TextMatrix(I, 10) = rsDocumentos.Fields("CODIGO")
            .row = I
            .Col = 11
            .CellFontSize = 9
            .TextMatrix(I, 11) = rsDocumentos.Fields("CtaBco")
            .row = I
            .Col = 12
            .CellFontSize = 9
            .TextMatrix(I, 12) = rsDocumentos.Fields("Division")
            .row = I
            .Col = 13
            .CellFontSize = 9
            .TextMatrix(I, 13) = rsDocumentos.Fields("ORDEN")
            .row = I
            .Col = 14
            .CellFontBold = True
            .CellFontSize = 8
            .CellForeColor = vbBlue
            .CellFontName = "Wingdings"
            .CellFontSize = 13
            .TextMatrix(I, 14) = strUnChecked
            
            If Trim(rsDocumentos.Fields("fec_vcto")) <> "" Then
                .TextMatrix(I, 15) = Format(CDate(rsDocumentos.Fields("fec_vcto")), "dd/mm/yyyy")
            Else
                .TextMatrix(I, 15) = ""
            End If
            .row = I
            .Col = 16
            .CellFontSize = 9
            .TextMatrix(I, 16) = rsDocumentos.Fields("Banco")
            .row = I
            .Col = 17
            .CellFontSize = 9
            .TextMatrix(I, 17) = rsDocumentos.Fields("Cod_Tipo_Doc")
            .row = I
            .Col = 18
            .CellFontSize = 9
            .TextMatrix(I, 18) = rsDocumentos.Fields("MPago")
            .row = I
            .Col = 19
            .CellFontSize = 9
            .TextMatrix(I, 19) = rsDocumentos.Fields("nombre")
            .row = I
            .Col = 20
            .CellFontSize = 9
            .TextMatrix(I, 20) = rsDocumentos.Fields("detracc")
            
            rsDocumentos.MoveNext
        Loop
        If I > 0 Then
            .row = 1
            .Col = 0
            .Rowsel = 0
            .ColSel = 16
        End If
        .Visible = True
    End With
    lblDocs = str(I)
    Set rsDocumentos = Nothing
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    LlenarTipDoc
    cboTipDoc.ListIndex = 0
    CargarDocs
    LLenarCbo "F"
    
End Sub

Private Sub flexListado_Click()
    Dim SCol As Integer
    Dim rsDetfac As MYSQL_RS
    With flexListado
        SCol = .Col
        If .row > 0 Then
            .Col = 0
            .ColSel = 20
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
    SQL = "Select d.identificador,d.serie, d.documento, d.mon,d.Total,d.pendiente,d.auxiliar,d.codigo,d.cenco,d.division," & _
            "d.fec_pago , d.fec_emision, d.fec_vcto, d.impEqui, d.Cod_Tipo_Doc, d.cod_estado, c.orden,C.obs,C.Division,O.auxiliar,ifnull(O.MPago,'-') as MPago,left (O.CtaBco,3) as TxtOficina,ifnull(O.CtaBco,'-') as CtaBco,ifnull(O.Banco,'-') as Banco,(select descrip from cnauxil where auxiliar=d.auxiliar and codigo=d.codigo) as nombre, Ifnull((select If(d.mon='N',Sum(c.cargos),Sum(c.cargod)) as Valor from cnmovi as c left join cnvouc as x on (c.anomes=x.anomes and c.voucher=x.voucher) where c.codaux=trim(d.codigo) and c.serdoc=trim(d.serie) and c.numdoc= trim(d.documento) and ((x.glosa like 'detracc%') or (x.glosa like 'reten%')or (x.glosa like 'antic%') or(x.glosa like 'aplic%'))),'0') as detracc  " & _
            "from (DOC_PROG as d  left join documento_contables as c on (d.identificador=c.identificador)) left join  orden_compra as O on c.orden=O.correl where left(d.identificador,4)='" & strAnoSistema & "'  order by d.fec_pago,d.serie, d.documento"
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
            .row = I
            .Col = 0
            .CellFontSize = 9
            .TextMatrix(I, 0) = CStr(I)
            .row = I
            .Col = 1
            .CellFontSize = 9
            .TextMatrix(I, 1) = rsDocumentos.Fields("IDENTIFICADOR")
            .row = I
            .Col = 2
            .CellFontSize = 9
            .TextMatrix(I, 2) = rsDocumentos.Fields("SERIE")
            .row = I
            .Col = 3
            .CellFontSize = 9
            .TextMatrix(I, 3) = rsDocumentos.Fields("DOCUMENTO")
            .row = I
            .Col = 4
            .CellFontSize = 9
            .TextMatrix(I, 4) = rsDocumentos.Fields("MON")
            .row = I
            .Col = 5
            .CellFontSize = 9
            .TextMatrix(I, 5) = rsDocumentos.Fields("TOTAL")
            .row = I
            .Col = 6
            .CellFontSize = 9
            .TextMatrix(I, 6) = rsDocumentos.Fields("PENDIENTE")
            .row = I
            .Col = 7
            .CellFontSize = 9
            .TextMatrix(I, 7) = rsDocumentos.Fields("AUXILIAR")
            .row = I
            .Col = 8
            .CellFontSize = 9
            .TextMatrix(I, 8) = Format(CDate(rsDocumentos.Fields("FEC_PAGO")), "dd/mm/yyyy")
            .row = I
            .Col = 9
            .CellFontSize = 9
            .TextMatrix(I, 9) = rsDocumentos.Fields("IMPEQUI")
            .row = I
            .Col = 10
            .CellFontSize = 9
            .TextMatrix(I, 10) = rsDocumentos.Fields("CODIGO")
            .row = I
            .Col = 11
            .CellFontSize = 9
            .TextMatrix(I, 11) = rsDocumentos.Fields("CtaBco")
            .row = I
            .Col = 12
            .CellFontSize = 9
            .TextMatrix(I, 12) = rsDocumentos.Fields("Division")
            .row = I
            .Col = 13
            .CellFontSize = 9
            .TextMatrix(I, 13) = rsDocumentos.Fields("ORDEN")
            .row = I
            .Col = 14
            .CellFontBold = True
            .CellFontSize = 8
            .CellForeColor = vbBlue
            .CellFontName = "Wingdings"
            .CellFontSize = 13
            .TextMatrix(I, 14) = strUnChecked
            
            If Trim(rsDocumentos.Fields("fec_vcto")) <> "" Then
                .TextMatrix(I, 15) = Format(CDate(rsDocumentos.Fields("fec_vcto")), "dd/mm/yyyy")
            Else
                .TextMatrix(I, 15) = ""
            End If
            .row = I
            .Col = 16
            .CellFontSize = 9
            .TextMatrix(I, 16) = rsDocumentos.Fields("Banco")
            .row = I
            .Col = 17
            .CellFontSize = 9
            .TextMatrix(I, 17) = rsDocumentos.Fields("Cod_Tipo_Doc")
            .row = I
            .Col = 18
            .CellFontSize = 9
            .TextMatrix(I, 18) = rsDocumentos.Fields("MPago")
            .row = I
            .Col = 19
            .CellFontSize = 9
            .TextMatrix(I, 19) = rsDocumentos.Fields("nombre")
            .row = I
            .Col = 20
            .CellFontSize = 9
            .TextMatrix(I, 20) = rsDocumentos.Fields("detracc")
            rsDocumentos.MoveNext
        Loop
        If I > 0 Then
            .row = 1
            .Col = 0
            .Rowsel = 0
            .ColSel = 16
        End If
        .Visible = True
    End With
    lblDocs = str(I)
    Set rsDocumentos = Nothing
End Sub


Private Sub meProveedor_Change()
  Dim SQL As String
  Dim sql2 As String
  Dim sql3 As String
  Dim rsDocumentos As MYSQL_RS
  Dim I As Integer
  lblDocs = "0"
  SQL = ""
  sql2 = ""
  sql3 = ""
    
        
        SQL = "Select d.identificador,d.serie, d.documento, d.mon,d.Total,d.pendiente,d.auxiliar,d.codigo,d.cenco,d.division,d.fec_pago , d.fec_emision, d.fec_vcto, d.impEqui,d.Cod_Tipo_Doc, d.cod_estado, c.orden,C.obs,C.Division,ifnull(O.MPago,ifnull((select descrip from tipopago where codpago=e.tipcta_mn),'')) as MPago,ifnull(left(O.CtaBco,3),ifnull(left(e.numcta_mn,3),'000')) as TxtOficina,ifnull(O.CtaBco,ifnull(e.numcta_mn,'0000000')) as CtaBco," & _
            "ifnull(O.Banco,ifnull((select descrip from pl_entidadfinanciera where codigo=e.codbanco),'')) as Banco,(select descrip from cnauxil where auxiliar=d.auxiliar and codigo=d.codigo) as nombre, Ifnull((select If(d.mon='N',Sum(c.cargos),Sum(c.cargod)) as Valor from cnmovi as c left join cnvouc as x on (c.anomes=x.anomes and c.voucher=x.voucher) where c.codaux=trim(d.codigo) and c.serdoc=trim(d.serie) and c.numdoc= trim(d.documento) and ((x.glosa like 'detracc%') or (x.glosa like 'reten%')or (x.glosa like 'antic%') or(x.glosa like 'aplic%'))),'0') as detracc  " & _
            "from ((DOC_PROG as d  left join documento_contables as c on (d.identificador=c.identificador)) left join  orden_compra as O on c.orden=O.correl) left join empleado as e on (e.codigo=d.codigo) where 1=1 "
        
'        left(d.identificador,4)='" & strAnoSistema & "'
        
        If (mskFecha1.Text <> "") And (mskFecha2.Text <> "") And ((mskFecha1.Text <> "__/__/____") And (mskFecha2.Text <> "__/__/____")) Then
           sql2 = " and concat(right('" & mskFecha1.Text & "',4),mid('" & mskFecha1.Text & "',3,4),left('" & mskFecha1.Text & "',2))<= d.fec_emision and " & _
                  "concat(right('" & mskFecha2.Text & "',4),mid('" & mskFecha2.Text & "',3,4),left('" & mskFecha2.Text & "',2))>= d.fec_emision "
        End If
        
        If (meProveedor.Text <> "") Then
           If Len(meProveedor.Text) > 10 Then
              sql3 = "and d.codigo='" & Trim(meProveedor.Text) & "' "
           Else
              sql3 = "and d.codigo='" & Trim(Right("00000000000" & meProveedor.Text, 11)) & "' "
           End If
        End If
          
    SQL = SQL & sql2 & sql3 & " order by d.serie, d.documento"
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
            .row = I
            .Col = 0
            .CellFontSize = 9
            .TextMatrix(I, 0) = CStr(I)
            .row = I
            .Col = 1
            .CellFontSize = 9
            .TextMatrix(I, 1) = rsDocumentos.Fields("IDENTIFICADOR")
            .row = I
            .Col = 2
            .CellFontSize = 9
            .TextMatrix(I, 2) = rsDocumentos.Fields("SERIE")
            .row = I
            .Col = 3
            .CellFontSize = 9
            .TextMatrix(I, 3) = rsDocumentos.Fields("DOCUMENTO")
            .row = I
            .Col = 4
            .CellFontSize = 9
            .TextMatrix(I, 4) = rsDocumentos.Fields("MON")
            .row = I
            .Col = 5
            .CellFontSize = 9
            .TextMatrix(I, 5) = rsDocumentos.Fields("TOTAL")
            .row = I
            .Col = 6
            .CellFontSize = 9
            .TextMatrix(I, 6) = rsDocumentos.Fields("PENDIENTE")
            .row = I
            .Col = 7
            .CellFontSize = 9
            .TextMatrix(I, 7) = rsDocumentos.Fields("AUXILIAR")
            .row = I
            .Col = 8
            .CellFontSize = 9
            .TextMatrix(I, 8) = Format(CDate(rsDocumentos.Fields("FEC_PAGO")), "dd/mm/yyyy")
            .row = I
            .Col = 9
            .CellFontSize = 9
            .TextMatrix(I, 9) = rsDocumentos.Fields("IMPEQUI")
            .row = I
            .Col = 10
            .CellFontSize = 9
            .TextMatrix(I, 10) = rsDocumentos.Fields("CODIGO")
            .row = I
            .Col = 11
            .CellFontSize = 9
            .TextMatrix(I, 11) = rsDocumentos.Fields("CtaBco")
            .row = I
            .Col = 12
            .CellFontSize = 9
            .TextMatrix(I, 12) = rsDocumentos.Fields("Division")
            .row = I
            .Col = 13
            .CellFontSize = 9
            .TextMatrix(I, 13) = rsDocumentos.Fields("ORDEN")
            .row = I
            .Col = 14
            .CellFontBold = True
            .CellFontSize = 8
            .CellForeColor = vbBlue
            .CellFontName = "Wingdings"
            .CellFontSize = 13
            .TextMatrix(I, 14) = strUnChecked
            
            If Trim(rsDocumentos.Fields("fec_vcto")) <> "" Then
                .TextMatrix(I, 15) = Format(CDate(rsDocumentos.Fields("fec_vcto")), "dd/mm/yyyy")
            Else
                .TextMatrix(I, 15) = ""
            End If
            .row = I
            .Col = 16
            .CellFontSize = 9
            .TextMatrix(I, 16) = rsDocumentos.Fields("Banco")
            .row = I
            .Col = 17
            .CellFontSize = 9
            .TextMatrix(I, 17) = rsDocumentos.Fields("Cod_Tipo_Doc")
            .row = I
            .Col = 18
            .CellFontSize = 9
            .TextMatrix(I, 18) = rsDocumentos.Fields("MPago")
            .row = I
            .Col = 19
            .CellFontSize = 9
            .TextMatrix(I, 19) = rsDocumentos.Fields("nombre")
            .row = I
            .Col = 20
            .CellFontSize = 9
            .TextMatrix(I, 20) = rsDocumentos.Fields("detracc")
            rsDocumentos.MoveNext
        Loop
        If I > 0 Then
            .row = 1
            .Col = 0
            .Rowsel = 0
            .ColSel = 16
        End If
        .Visible = True
    End With
    lblDocs = str(I)
    Set rsDocumentos = Nothing
  
 
  
End Sub


Sub LLenarCbo(F As String)
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    Dim SQL As String
    If F = "Cuentasxpagar" Then
        SQL = "Select tip_linea,descrip From CNTABLAS where codtab='1' and " & _
              "(tip_linea='0' or tip_linea>='3' and tip_linea<='5') Order By 1"
    Else
        If F = "Clientes" Then
            SQL = "Select tip_linea,descrip From CNTABLAS where codtab='1' and " & _
                  "(tip_linea='0' or tip_linea='2') Order By 1"
        Else
            SQL = "Select tip_linea,descrip From CNTABLAS where codtab='1' Order By 1"
        End If
    End If
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    Do While Not Rs.EOF
        cboauxiliar.AddItem Rs.Fields(0) + "  " + Rs.Fields(1)
        Rs.MoveNext
    Loop
    Rs.CloseRecordset
    Set Rs = Nothing
End Sub



