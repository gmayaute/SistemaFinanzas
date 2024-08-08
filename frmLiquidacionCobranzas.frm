VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmLiquidacionCobranzas 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidacion de Cobranza"
   ClientHeight    =   6780
   ClientLeft      =   2445
   ClientTop       =   4365
   ClientWidth     =   11205
   Icon            =   "frmLiquidacionCobranzas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11205
   Begin NOVAdmin.flxEdit flxLiquidacion 
      Height          =   5445
      Left            =   60
      TabIndex        =   10
      Top             =   750
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   9604
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   585
      Left            =   30
      TabIndex        =   5
      Top             =   6180
      Width           =   11115
      Begin Proyecto1.chameleonButton chBtnSalir 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   10590
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   150
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         MICON           =   "frmLiquidacionCobranzas.frx":068A
         PICN            =   "frmLiquidacionCobranzas.frx":06A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton chBtnReporte 
         Height          =   375
         Left            =   10110
         TabIndex        =   7
         ToolTipText     =   "Ver Reporte"
         Top             =   150
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
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
         MICON           =   "frmLiquidacionCobranzas.frx":0A6C
         PICN            =   "frmLiquidacionCobranzas.frx":0A88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton BtnGrabar 
         Height          =   375
         Left            =   4890
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Grabar"
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
         MICON           =   "frmLiquidacionCobranzas.frx":0FCA
         PICN            =   "frmLiquidacionCobranzas.frx":0FE6
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
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Caption         =   "Semana"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11115
      Begin MSMask.MaskEdBox DtpFecha 
         Height          =   285
         Index           =   0
         Left            =   900
         TabIndex        =   1
         ToolTipText     =   "Fecha_Pago"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   128
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DtpFecha 
         Height          =   285
         Index           =   1
         Left            =   2970
         TabIndex        =   3
         ToolTipText     =   "Fecha_Pago"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   128
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Proyecto1.chameleonButton cmdBuscar 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   4650
         TabIndex        =   9
         ToolTipText     =   "Busqueda"
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "&Buscar"
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
         MICON           =   "frmLiquidacionCobranzas.frx":1428
         PICN            =   "frmLiquidacionCobranzas.frx":1444
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Al:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   2430
         TabIndex        =   4
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Del:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmLiquidacionCobranzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Consulta As FrmConsultas

Private Sub btnGrabar_Click()
    Dim I As Integer
    Dim aviso As Integer
    Dim SQL As String
    aviso = 0
    With flxLiquidacion
        For I = 1 To .Rows - 1
            If ActualizarDatosLiquid(I) Then aviso = aviso + 1
        Next
    End With
    If aviso > 0 Then
        MsgBox "Se Registraron los Datos de Liquidacion de Cobranzas ", vbInformation, "Liquidacion de Cobranzas"
    Else
        MsgBox "Los Datos ya fueron Actualizados o Registrados", vbInformation, "Liquidacion de Cobranzas"
    End If
End Sub

Private Sub chBtnReporte_Click()
'    llenartempLiq
    Set oReporte = New clsReporte
        oReporte.empresa = strNombreEmpresa
        If DtpFecha(0) <> "__/__/____" And DtpFecha(1) <> "__/__/____" Then oReporte.Titulo = "LIQUIDACION DE COBRANZAS DEL " & DtpFecha(0) & " AL " & DtpFecha(1)
        If DtpFecha(0) <> "__/__/____" And DtpFecha(1) = "__/__/____" Then oReporte.Titulo = "LIQUIDACION DE COBRANZAS DEL " & DtpFecha(0) & " AL " & DtpFecha(0)
        If DtpFecha(0) = "__/__/____" And DtpFecha(1) <> "__/__/____" Then oReporte.Titulo = "LIQUIDACION DE COBRANZAS DEL " & DtpFecha(1) & " AL " & DtpFecha(1)
        If DtpFecha(0) = "__/__/____" And DtpFecha(1) = "__/__/____" Then oReporte.Titulo = "LIQUIDACION DE COBRANZAS DEL MES DE " & NombreMes(strMesSistema, False)
        oReporte.Reporte = "Rep_LiquidCobranza.rpt"
        oReporte.sp_LiquidCobranza DtpFecha(0), DtpFecha(1)
End Sub

Private Sub chBtnSalir_Click()
    Set Consulta = Nothing
    Unload Me
End Sub

Public Sub FormatFlexLiq()
    Dim I As Integer
    With flxLiquidacion
        .Clear
        .Rows = 1
        .Cols = 19
        .ForeColorFixed = &H404000
        
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(5) + "Identificador"
        
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = Space(1) + "Fec_Cobro"
        
        .ColWidth(3) = 0
        .TextMatrix(0, 3) = Space(5) + "TDoc"
        
        .ColWidth(4) = 900
        .TextMatrix(0, 4) = Space(5) + "Doc"
        
        .ColWidth(5) = 1500
        .TextMatrix(0, 5) = Space(9) + "N° Doc"

        .ColWidth(6) = 0
        .TextMatrix(0, 6) = Space(1) + "Aux"
        
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = Space(6) + "Codigo"
        
        .ColWidth(8) = 0
        .TextMatrix(0, 8) = Space(5) + "Proveedor"
        
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = Space(1) + "Estado"
        
        .ColWidth(10) = 1200
        .TextMatrix(0, 10) = Space(6) + "Importe"
        .ColType(10) = cadena
        .ColDecimales(10) = 2
        
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = Space(5) + "TDoc_Ref"
        
        .ColWidth(12) = 0
        .TextMatrix(0, 12) = Space(11) + "Doc_Desc"
        .ColType(12) = cadena
        .CaracteresValidos(12) = "0123456789_abcdefghijklmnñopqrstuvwxyz ABECDEFGHIJKLMNÑOPQRSTUVWXYZ"
        .ColMaxLength(12) = 30
        
        .ColWidth(13) = 0
        .TextMatrix(0, 13) = Space(6) + "N° Doc_Ref"
        .ColType(13) = cadena
        .CaracteresValidos(13) = "0123456789-"
        .ColMaxLength(13) = 15
        
        .ColWidth(14) = 1200
        .TextMatrix(0, 14) = Space(3) + "Tot. Dctos."
        .ColType(14) = cadena
        .ColDecimales(14) = 2
        .CaracteresValidos(14) = "0123456789-."
        .ColMaxLength(14) = 15
        
        .ColWidth(15) = 1200
        .TextMatrix(0, 15) = Space(9) + "Saldo"
        .ColType(15) = cadena
        .ColDecimales(15) = 2

        .ColWidth(16) = 0
        .TextMatrix(0, 16) = Space(10) + "Afecto"
        .ColType(16) = cadena
        
        .ColWidth(17) = 0
        .TextMatrix(0, 17) = Space(10) + "%Detracccion"
        .ColType(17) = cadena
        
        .ColWidth(18) = 600
        .TextMatrix(0, 18) = "Mostrar"
        .ColType(18) = cadena
        .CaracteresValidos(18) = "SN"
        .ColMaxLength(18) = 1
     End With
End Sub

Private Sub cmdBuscar_Click()
    FormatFlexLiq
    flxLiquidacion.Visible = False
    BuscarLiquidacion DtpFecha(0), DtpFecha(1)
    flxLiquidacion.Visible = True
End Sub

Private Sub DtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And KeyAscii = 13 Then DtpFecha(1).SetFocus
    If Index = 1 And KeyAscii = 13 Then cmdBuscar.SetFocus
End Sub

Private Sub flxLiquidacion_Click()
    With flxLiquidacion
        If .Col = 18 Then
            If .TextMatrix(.row, 18) = strChecked Then
                .TextMatrix(.row, 18) = strUnChecked
            Else
                .TextMatrix(.row, 18) = strChecked
            End If
            ActualizaSeleccion IIf(.TextMatrix(.row, 18) = strChecked, "S", "N"), .TextMatrix(.row, 1)
        End If
    End With
End Sub

Private Sub flxLiquidacion_KeyDown(KeyCode As Integer, Shift As Integer)
    With flxLiquidacion
        If .Col = 12 Or .Col = 13 Or .Col = 14 Then
           Publimensaje = "modificar"
        Else
           Publimensaje = "sin-editar"
        End If
        If .Col = 14 Then
            TipodeCampo = Numero
        Else
            TipodeCampo = cadena
        End If
        If .Col = 12 Then
            If KeyCode = 112 And Publimensaje = "modificar" Then
                With Consulta
                    .pCols = 3
                    .pCol = 0: .pAnchoCol = 1200
                    .pCol = 1: .pAnchoCol = 1000
                    .pCol = 2: .pAnchoCol = 3000
                    .pTitulo = "Consulta de Docs Lquidacion"
                    .pForm = FORM_LIQUIDCOBRANZA
                    .pCaso = LABEL_TIP_DOC
                    .Show
                End With
            End If
        End If
    End With
End Sub

Private Sub flxLiquidacion_KeyPress(KeyAscii As Integer)
    With flxLiquidacion
        If .TextMatrix(.row, 14) <> "" Then
            If KeyAscii = 13 Then
                .TextMatrix(.row, 15) = FormatNumber(.TextMatrix(.row, 10) - .TextMatrix(.row, 14), 2)
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call WheelHook(frmLiquidacionCobranzas)
    Set Consulta = New FrmConsultas
    FormatFlexLiq
    TipodeCampo = cadena
    Publimensaje = "sin-editar"
End Sub

Public Sub llenartempLiq()
    Dim SQL As String
    SQL = "Delete from tmpliquidaciondecobranza "
          oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    SQL = "Insert into tmpliquidaciondecobranza " & _
          "select distinct dc.Identificador,dc.Fec_Pago as FechaCobro,a.cod_tipo_doc,'FACTURA'," & _
          "CONCAT(dc.serie,'-',dc.Correl) as Documento,dc.auxiliar,dc.codigo," & _
          "aux.descrip as Proveedor,mov.Cod_Estado,dc.Total," & _
          "dc.Cod_Tipo_Ref,cn.descrip as Tipo_Ref,dc.Ref as NDocto_Ref,dc.Total_Ref," & _
          " (dc.total-dc.total_Ref) as Saldo, dt.Afecto,des.Porcentaje, " & _
          " IF(dc.Total_Ref > 0,'S','N') as sel " & _
          " from (((((amarre_documento as a LEFT Join documento_contables as dc " & _
          " on (a.Identificador=dc.Identificador))Left Join movi_documento as mov " & _
          " on (mov.Identificador=dc.Identificador))Left Join detallefact as dt " & _
          " on (dc.Identificador=dt.Identificador))LEFT Join descuentos as des " & _
          " on (des.descripcion LIKE '%Detraccion'))LEFT Join cndocum as cn " & _
          " on (dc.Cod_Tipo_Ref=cn.CodDoc))left join cnauxil as aux" & _
          " on (dc.auxiliar=aux.auxiliar and dc.codigo=aux.codigo) " & _
          " where a.flag='0' and dc.auxiliar='2' and a.cod_tipo_doc='01' and (mov.Cod_Estado='" & CANCELADO & "' or dc.TOTAL_REF>0.00) "
          oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
End Sub

Public Sub BuscarLiquidacion(ByVal FechaIni As String, ByVal FechaFin As String)
    Dim SQL As String
    Dim rsLiq As MYSQL_RS

    llenartempLiq
    SQL = "select * from tmpliquidaciondecobranza where"
    If FechaIni <> "__/__/____" And FechaFin <> "__/__/____" Then
        SQL = SQL & " FechaCobro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and FechaCobro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
    Else
        If FechaIni <> "__/__/____" And FechaFin = "__/__/____" Then
            SQL = SQL & " FechaCobro>='" & Format(FechaIni, "yyyy/mm/dd") & "' and FechaCobro<='" & Format(FechaIni, "yyyy/mm/dd") & "'"
        Else
            If FechaIni = "__/__/____" And FechaFin <> "__/__/____" Then
                SQL = SQL & " FechaCobro>='" & Format(FechaFin, "yyyy/mm/dd") & "' and FechaCobro<='" & Format(FechaFin, "yyyy/mm/dd") & "'"
            Else
                SQL = SQL & " substring(FechaCobro,6,2)>='" & strMesSistema & "' and substring(FechaCobro,6,2)<='" & strMesSistema & "'"
            End If
        End If
    End If
    
    SQL = SQL & " ORDER BY Ndocto"
    Set rsLiq = oConexion.EjecutaSelectRS(SQL)
    DatosLiquidacion rsLiq
    Set rsLiq = Nothing
End Sub

Public Sub DatosLiquidacion(Rs As MYSQL_RS)
    Dim I As Integer
    With flxLiquidacion
        Do While Not (Rs.EOF)
            .Rows = .Rows + 1
            .row = .Rows - 1
            If Rs.Fields("FechaCobro") <> "" Then Rs.Fields("FechaCobro") = Format(Rs.Fields("FechaCobro"), "DD/MM/YYYY")
            For I = 1 To .Cols - 1
                If Rs.Fields("Afecto") = 1 Then
                    Select Case I
                    Case 10: .TextMatrix(.row, I) = FormatNumber(CDbl(Rs.Fields("Total")), 2)
                    Case 11: .TextMatrix(.row, I) = "17"
                    Case 12: .TextMatrix(.row, I) = "PAGO DETRACCION"
                    Case 18:
                        .Col = 18
                        .CellFontName = "Wingdings"
                        .CellFontSize = 11
                        .ColAlignment(18) = flexAlignCenterBottom
                        .TextMatrix(.row, I) = IIf(Rs.Fields("SELECCION") = "S", strChecked, strUnChecked)
                    Case 14:
                        If Rs.Fields("Total_Ref") = 0 Then
                            .TextMatrix(.row, 14) = FormatNumber((CDbl(Rs.Fields("Total")) * CDbl(Rs.Fields("Descripcion"))) / 100, 2)
                        Else
                            .TextMatrix(.row, 14) = CEN(Rs.Fields("Total_Ref"))
                        End If
                    Case 15
                         .TextMatrix(.row, 15) = FormatNumber(CEN(Rs.Fields("Total")) - CEN(Rs.Fields("Total_ref")), 2)
                    Case Else
                        .TextMatrix(.row, I) = Rs.Fields(I - 1)
                End Select
            Else
                If I = 10 Then
                    .TextMatrix(.row, I) = FormatNumber(CDbl(Rs.Fields("Total")), 2)
                Else
                    If I = 15 Then
                        .TextMatrix(.row, 15) = FormatNumber(CEN(Rs.Fields("Total")) - CEN(Rs.Fields("Total_ref")), 2)
                    Else
                        If I = 18 Then
                            .Col = 18
                            .CellFontName = "Wingdings"
                            .CellFontSize = 11
                            .ColAlignment(18) = flexAlignCenterBottom
                            .TextMatrix(.row, I) = IIf(Rs.Fields("SELECCION") = "S", strChecked, strUnChecked)
                        Else
                            .TextMatrix(.row, I) = CE(Rs.Fields(I - 1))
                        End If
                    End If
                End If
            End If
       Next I
       Rs.MoveNext
    Loop
  End With
Set Rs = Nothing
EnumerarItems flxLiquidacion
End Sub
Private Function ActualizarDatosLiquid(fila As Integer) As Boolean
ActualizarDatosLiquid = False
Dim SQL As String
    With flxLiquidacion
        If .TextMatrix(fila, 14) <> "" Then
            SQL = "Update documento_contables set Cod_Tipo_Ref='" & .TextMatrix(fila, 11) & "'," & _
                  "Ref='" & .TextMatrix(fila, 13) & "',TOTAL_Ref='" & CDbl(.TextMatrix(fila, 14)) & "'," & _
                  "Cancelado='" & CDbl(.TextMatrix(fila, 15)) & "' where Identificador='" & .TextMatrix(fila, 1) & "'"
            ActualizarDatosLiquid = oConexion.EjecutaInsertUpdateDelete(SQL, TIPO_QUERY.Modificar, False)
        End If
    End With
End Function
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    
    On Error Resume Next
    
    With flxLiquidacion
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

Public Function ExisteDocumento(Id As String) As Boolean
Dim SQL As String
Dim rsDoc As MYSQL_RS
SQL = " Select Identificador from liquidaciondecobranza " & _
      " where Identificador='" & Id & "'"
Set rsDoc = oConexion.EjecutaSelectRS(SQL)
If rsDoc.RecordCount = 0 Then ExisteDocumento = False
If rsDoc.RecordCount >= 1 Then ExisteDocumento = True
End Function

Sub ActualizaSeleccion(Sel As String, Ident As String)
Dim SQL As String
    SQL = "Update tmpliquidaciondecobranza SET seleccion = '" & Sel & "' " & _
          "where identificador = '" & Ident & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
End Sub
