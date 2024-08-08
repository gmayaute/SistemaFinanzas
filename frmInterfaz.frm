VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterfaz 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación de Orden de Compra/Servicio"
   ClientHeight    =   2535
   ClientLeft      =   4545
   ClientTop       =   3930
   ClientWidth     =   6855
   Icon            =   "frmInterfaz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6855
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   510
      Top             =   2820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   """Text (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txtpatharchivo 
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   1020
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H009F5539&
      DrawMode        =   6  'Mask Pen Not
      Height          =   555
      Left            =   210
      Picture         =   "frmInterfaz.frx":113A
      ScaleHeight     =   495
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   150
      Width           =   525
   End
   Begin Proyecto1.chameleonButton CmdExaminar 
      Height          =   405
      Left            =   4080
      TabIndex        =   5
      Top             =   1470
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Examinar"
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
      MICON           =   "frmInterfaz.frx":157C
      PICN            =   "frmInterfaz.frx":1598
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdProcesar 
      Height          =   405
      Left            =   5460
      TabIndex        =   6
      Top             =   1500
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Importar"
      ENAB            =   0   'False
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
      MICON           =   "frmInterfaz.frx":391A
      PICN            =   "frmInterfaz.frx":3936
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar pbProgreso 
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   2250
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblActualizadas 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizadas..."
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
      Left            =   60
      TabIndex        =   7
      Top             =   1980
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label lblInsertadas 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Insertadas..."
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
      Left            =   60
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label lblMensaje 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo Archivo..."
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
      Left            =   60
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Archivo:"
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
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   930
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el Archivo de Texto, utilizando el botón ""Examinar"". Luego presione el botón ""Procesar"""
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
      Height          =   495
      Left            =   900
      TabIndex        =   1
      Top             =   180
      Width           =   5745
   End
   Begin VB.Label Label1 
      BackColor       =   &H009F5539&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6765
   End
End
Attribute VB_Name = "frmInterfaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents a As NOVAdmin.cFile
Attribute a.VB_VarHelpID = -1
Dim separador As String
Dim id_ordencompra As String
Dim cuenta As Long
Dim ctaupd As Long
Dim total As Long

Private Sub a_ParseComplete()
    pbProgreso.Value = 0
End Sub

Private Sub a_PercentParsing(bPercent As Byte)
    If bPercent = 0 Then
        Debug.Print
    Else
        pbProgreso.Value = bPercent
    End If
End Sub

Private Sub CmdExaminar_Click()
 On Error GoTo err
    separador = ","
    Set a = Nothing
    Set a = New NOVAdmin.cFile
    With a
        .CloseFile
        CommonDialog1.ShowOpen
        .Filename = CommonDialog1.Filename
        NombreArchivo = CommonDialog1.Filename
        txtpatharchivo.Text = NombreArchivo
        If separador <> "" Then
              .FieldSeparator = separador
        End If
        Me.MousePointer = vbHourglass
        LblMensaje.Visible = True
        lblInsertadas.Visible = False
        lblActualizadas.Visible = False
        LblMensaje.Caption = "Leyendo archivo..."
        .Parse
        .CloseFile
        Me.MousePointer = vbNormal
        LblMensaje.Caption = str(.Lines.Count) & " Ordenes de Compra por procesar."
    End With
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdProcesar_Click()
    If LblMensaje.Visible = False Then LblMensaje.Visible = True
    Me.MousePointer = vbHourglass
    If InStr(1, txtpatharchivo, "OCREC20", vbTextCompare) > 0 Then
        LLenarEntregas
    Else
        LLenarTablaenBD
    End If
    Me.MousePointer = vbNormal
    MsgBox "PROCESO TERMINADO", vbOKOnly + vbInformation, "AVISO"
    LblMensaje.Caption = total & " Ordenes de Compra procesadas."
    lblInsertadas.Visible = True
    lblInsertadas.Caption = cuenta & " Insertadas"
    lblActualizadas.Visible = True
    lblActualizadas.Caption = ctaupd & " Actualizadas"
    txtpatharchivo.Text = Empty
    cmdProcesar.Enabled = False
    cmdExaminar.SetFocus
    pbProgreso.Value = 0
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set a = Nothing
End Sub

Private Sub txtpatharchivo_Change()
    If txtpatharchivo.Text <> Empty Then
        cmdProcesar.Enabled = True
    End If
End Sub

Private Sub LLenarTablaenBD()

' On Error GoTo mensaje
    Dim Query As String
    Dim QueryInsOrden As String
    Dim QueryInsMovi As String
    Dim QueryUpdOrden As String
    Dim QueryInsAmarre As String
    Dim QueryInsHist As String
    Dim QueryUpdMovi As String
    Dim i As Long
    Dim intCodigo As String
    Dim SQL As String
    Dim TmpDireccion
    
    cuenta = 0
    ctaupd = 0
    total = a.Lines.Count
    
    For i = 1 To a.Lines.Count
             intCodigo = GenerarCodigo
             QueryInsOrden = "Call Insert_Orden( '" & intCodigo & "',"
             QueryInsMovi = "Call Insert_Movi_Doc('" & intCodigo & "'," & _
                            "'" & Format(Date, "yyyy/mm/dd") & "'," & _
                            "'" & Left(a.Lines(i).Fields(ORDENCS.ESTADOA).Value, 2) & "'," & _
                            "'" & 1 & "',  '" & strUsuarioId & "'); "
             QueryInsAmarre = "Call Insert_AmarreDoc( '" & intCodigo & "'," & _
                             "'O', '" & Format(Date, "yyyy/mm/dd") & "'," & _
                             "'" & Format(Time, "HH:MM") & "','00', '00000000'," & _
                             "'NINGUNO', '" & strNombreEmpresa & "', 'NINGUNO', " & _
                             " '" & FAMILIA_DOC.ORDENES & "', '" & strUsuarioId & "'," & _
                             "'" & strAnoSistema & strMesSistema & "','" & 0 & "');"
             QueryInsHist = "Call Insert_HistorialDoc ( '" & intCodigo & "', '" & Right(a.Lines(i).Fields(ORDENCS.ESTADOA).Value, 2) & "'," & _
                            "'" & DescripcionesdeCodigos("CNUSER", strUsuarioId, "area") & "'," & _
                            "'" & Format(a.Lines(i).Fields(ORDENCS.FEC_EMI).Value, "yyyy/mm/dd") & "'," & _
                            "'" & strUsuarioId & "');"
             QueryUpdOrden = "Call Update_OrdenCompraImportadas('" & Right("000000000" & Trim(a.Lines(i).Fields(ORDENCS.corr).Value), 9) & "','" & _
                             a.Lines(i).Fields(ORDENCS.Proveedor).Value & "','" & Left(a.Lines(i).Fields(ORDENCS.FEC_EMI).Value, 10) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.CENCOSTO).Value) & "','" & Left(a.Lines(i).Fields(ORDENCS.VISADO).Value, 10) & "','" & _
                             a.Lines(i).Fields(ORDENCS.USUARIOO).Value & "','" & a.Lines(i).Fields(ORDENCS.MONORDEN).Value & "','" & _
                             Left(a.Lines(i).Fields(ORDENCS.ESTADOA).Value, 10) & "'," & a.Lines(i).Fields(ORDENCS.VCOMPRA).Value & "," & _
                             a.Lines(i).Fields(ORDENCS.IGVO).Value & "," & a.Lines(i).Fields(ORDENCS.EXONERADOO).Value & "," & _
                             CEN(a.Lines(i).Fields(ORDENCS.TOTORDEN).Value) & "," & a.Lines(i).Fields(ORDENCS.TUNIDADES).Value & "," & _
                             a.Lines(i).Fields(ORDENCS.UNIDADESR).Value & ",'" & a.Lines(i).Fields(ORDENCS.FLAGO).Value & "','" & _
                             a.Lines(i).Fields(ORDENCS.MEDIOPAGO).Value & "'," & CEN(Trim(a.Lines(i).Fields(ORDENCS.DIASCRE).Value)) & ",'" & _
                             Trim(a.Lines(i).Fields(ORDENCS.BANCOO).Value) & "','" & Trim(a.Lines(i).Fields(ORDENCS.CTABCO).Value) & "','" & _
                             IIf(a.Lines(i).Fields(ORDENCS.FANULADO).Value = " ", "N", a.Lines(i).Fields(ORDENCS.FANULADO).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.ORDENABIERTA).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.NIVELAUT).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.AUTACTUAL).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.ORDENCTA).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.CUENTAAT).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.DivLog).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.DivCont).Value) & "');"
 
             Query = " '00000'," & _
                     " '" & Right("000000000" & Trim(a.Lines(i).Fields(ORDENCS.corr).Value), 9) & "'," & _
                     " '" & a.Lines(i).Fields(ORDENCS.Tipo).Value & "'," & _
                     " '" & 5 & "'," & " '" & a.Lines(i).Fields(ORDENCS.Proveedor).Value & "'," & _
                     " '" & Left(a.Lines(i).Fields(ORDENCS.FEC_EMI).Value, 10) & "'," & " '" & Left(a.Lines(i).Fields(ORDENCS.CENCOSTO).Value, 11) & "'," & _
                     " '" & Trim(Left(a.Lines(i).Fields(ORDENCS.VISADO).Value, 10)) & "'," & _
                     " '" & a.Lines(i).Fields(ORDENCS.USUARIOO).Value & "'," & " 'ORDENES IMPORTADAS DEL SISTEMA LOGISTICO'," & _
                     " '" & UCase(a.Lines(i).Fields(ORDENCS.MONORDEN).Value) & "'," & _
                     " '" & Trim(Left(a.Lines(i).Fields(ORDENCS.ESTADOA).Value, 10)) & "', " & _
                     " " & a.Lines(i).Fields(ORDENCS.VCOMPRA).Value & ", " & " " & a.Lines(i).Fields(ORDENCS.IGVO).Value & "," & _
                     " " & a.Lines(i).Fields(ORDENCS.EXONERADOO).Value & "," & " " & CEN(a.Lines(i).Fields(ORDENCS.TOTORDEN).Value) & "," & _
                     " " & a.Lines(i).Fields(ORDENCS.TUNIDADES).Value & "," & " " & a.Lines(i).Fields(ORDENCS.UNIDADESR).Value & "," & _
                     " '" & a.Lines(i).Fields(ORDENCS.FLAGO).Value & "'," & " '" & Trim(a.Lines(i).Fields(ORDENCS.MEDIOPAGO).Value) & "'," & _
                     " '" & DescripcionesdeCodigos("CENCO", a.Lines(i).Fields(ORDENCS.CENCOSTO).Value, "2") & "'," & _
                     " " & CEN(Trim(a.Lines(i).Fields(ORDENCS.DIASCRE).Value)) & "," & _
                     " '" & a.Lines(i).Fields(ORDENCS.BANCOO).Value & "'," & _
                     " '" & a.Lines(i).Fields(ORDENCS.CTABCO).Value & "'," & _
                     " '" & IIf(a.Lines(i).Fields(ORDENCS.FANULADO).Value = " ", "N", a.Lines(i).Fields(ORDENCS.FANULADO).Value) & "','" & _
                     Trim(a.Lines(i).Fields(ORDENCS.ORDENABIERTA).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.NIVELAUT).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.AUTACTUAL).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.ORDENCTA).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.CUENTAAT).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.DivLog).Value) & "','" & _
                             Trim(a.Lines(i).Fields(ORDENCS.DivCont).Value) & "');"
                     
             If Not existeOrden(Right("000000000" & Trim(a.Lines(i).Fields(ORDENCS.corr).Value), 9)) Then
                 oConexion.EjecutaInsertUpdateDelete QueryInsAmarre, TIPO_QUERY.insertar, False
                 oConexion.EjecutaInsertUpdateDelete QueryInsOrden & Query, TIPO_QUERY.insertar, False
                 oConexion.EjecutaInsertUpdateDelete QueryInsMovi, TIPO_QUERY.insertar, False
                 oConexion.EjecutaInsertUpdateDelete QueryInsHist, TIPO_QUERY.insertar, False
                 cuenta = cuenta + 1
             Else
                 oConexion.EjecutaInsertUpdateDelete QueryUpdOrden, TIPO_QUERY.Modificar, False
                 ctaupd = ctaupd + 1
             End If
             
             TmpDireccion = Split(Trim(a.Lines(i).Fields(30).Value), "|")
             
             
             'AQUI VALIDAMOS LA EXISTENCIA DEL PROVEEDOR
             'If (NoExisteAuxiliarEnCont("5", Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value)) = True) And (Left(Trim(a.Lines(i).Fields(4).Value), 4) = strAnoSistema) And (Len(Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value)) >= 11) Then
              If (NoExisteAuxiliarEnCont("5", Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value)) = True) And (Left(Trim(a.Lines(i).Fields(4).Value), 4) = strAnoSistema) Then
                SQL = "Call Insert_Auxiliar ('5'" & _
                ",'" & Replace(Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value), ",", " ") & "','" & Replace(Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value), ",", " ") & "','" & Replace(Trim(a.Lines(i).Fields(29).Value), ",", " ") & _
                "','" & Replace(Trim(a.Lines(i).Fields(30).Value), "|", "-") & "','" & Replace(IIf(UBound(TmpDireccion) > 0, IIf(UBound(TmpDireccion) < 3, "", Trim(TmpDireccion(UBound(TmpDireccion)))), ""), ",", " ") & "','" & _
                "','N','','','" & _
                "','','','N',0" & _
                ",'" & Trim(a.Lines(i).Fields(44).Value) & "','','" & _
                "','','','" & _
                "','06" & _
                "','','','', " & _
                "'','','','','L')"
                
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
                
              Else
                'MODIFICACION DE LA ORDEN
                SQL = "Call Update_Auxiliar_masivo ('5'" & _
                ",'" & Replace(Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value), ",", " ") & "','" & Replace(Trim(a.Lines(i).Fields(ORDENCS.Proveedor).Value), ",", " ") & "','" & Replace(Trim(a.Lines(i).Fields(29).Value), ",", " ") & _
                "','" & Replace(Trim(a.Lines(i).Fields(30).Value), "|", "-") & "','" & Replace(IIf(UBound(TmpDireccion) > 0, IIf(UBound(TmpDireccion) < 3, "", Trim(TmpDireccion(UBound(TmpDireccion)))), ""), ",", " ") & "','" & _
                "',0,'','','','" & _
                "','','','N',0" & _
                ",'" & Trim(a.Lines(i).Fields(44).Value) & "','','" & _
                "','','','" & _
                "','06" & _
                "','','','', " & _
                "'','','','','" & Trim(a.Lines(i).Fields(31).Value) & "','" & Trim(a.Lines(i).Fields(32).Value) & "','L')"
                
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
                
              End If
             'FIN DE VALIDACION DE EXISTENCIA DE PROVEEDOR
             
            pbProgreso.Value = i * 100 / a.Lines.Count
       Next i
'mensaje:
'    MsgBox "No se grabó correctamente la linea " & I & " modificada Uno de los datos es incorrecto ", vbOKOnly + vbInformation, "Aviso"
'    Resume
'    Resume Next
'    Exit Sub
End Sub
Private Sub LLenarEntregas()
    Dim Query As String
    Dim QueryInsOrden As String
    Dim QueryInsMovi As String
    Dim QueryUpdOrden As String
    Dim QueryInsAmarre As String
    Dim QueryInsHist As String
    Dim QueryUpdMovi As String
    Dim i As Long
    Dim intCodigo As Integer
    
    cuenta = 0
    ctaupd = 0
    total = a.Lines.Count
    
    For i = 1 To a.Lines.Count
             intCodigo = GenerarItemOrRec(Trim(a.Lines(i).Fields(1).Value))
             
             'QueryDelOrden = "Call Delete_CantRecOrd( '" & Right("000000000" & Trim(a.Lines(I).Fields(1).Value), 9) & "')"
                          
             QueryInsOrden = "Call Insert_CantRecOrd( '" & Right("000000000" & Trim(a.Lines(i).Fields(1).Value), 9) & "',"
             
             QueryUpdOrden = "Call Update_EntregaOrden('" & Right("000000000" & Trim(a.Lines(i).Fields(1).Value), 9) & "','" & _
                              a.Lines(i).Fields(2).Value & "','" & _
                              a.Lines(i).Fields(3).Value & "','" & _
                              a.Lines(i).Fields(4).Value & "','" & _
                              Left(a.Lines(i).Fields(5).Value, 10) & "'," & _
                              a.Lines(i).Fields(6).Value & ",'" & _
                              IIf(a.Lines(i).Fields(7).Value = " ", "N", a.Lines(i).Fields(7).Value) & "','" & Trim(a.Lines(i).Fields(8).Value) & "')"

             Query = intCodigo & ",'" & _
                     a.Lines(i).Fields(2).Value & "','" & _
                     a.Lines(i).Fields(3).Value & "','" & _
                     a.Lines(i).Fields(4).Value & "','" & _
                     Left(a.Lines(i).Fields(5).Value, 10) & "'," & _
                     a.Lines(i).Fields(6).Value & ",'" & _
                     IIf(a.Lines(i).Fields(7).Value = " ", "N", a.Lines(i).Fields(7).Value) & "','" & _
                     a.Lines(i).Fields(8).Value & "')"
                                         
             'If Not existenrecibidas(Trim(a.Lines(I).Fields(1)), Trim(a.Lines(I).Fields(8).Value)) Then
'                 oConexion.EjecutaInsertUpdateDelete QueryInsOrden & Query, TIPO_QUERY.insertar, False
                cuenta = cuenta + oConexion.EjecutaInsertUpdateDelete2(QueryInsOrden & Query, TIPO_QUERY.insertar, False)
             'Else
                'oConexion.EjecutaInsertUpdateDelete QueryUpdOrden, TIPO_QUERY.Modificar, False
                ctaupd = ctaupd + oConexion.EjecutaInsertUpdateDelete2(QueryUpdOrden, TIPO_QUERY.Modificar, False)
             'End If
             
             
            pbProgreso.Value = i * 100 / a.Lines.Count
       Next i
       
End Sub

Private Function GenerarCodigo() As String
    Dim Rs As MYSQL_RS
    Dim AnoMes As String
    Dim SQL As String
    AnoMes = strAnoSistema & strMesSistema
    SQL = "max_identificador where anomes = '" & AnoMes & "'"
    Set Rs = oConexion.EjecutaSelect(SQL)
    If Not Rs.EOF Then
        GenerarCodigo = Rs.Fields("anomes") & Right("0000" & Trim(str(val(Rs.Fields("maximo")) + 1)), 4)
    End If
    If Rs.EOF Then
        GenerarCodigo = AnoMes & "0001"
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Function

Private Function GenerarItemOrRec(ORD As String) As Integer
    Dim rsA As MYSQL_RS
    Dim AnoMes As String
    Dim SQL As String
    SQL = "Select max(item) as m from orden_rec where  orden='" & Right("000000000" & Trim(ORD), 9) & "'"
    Set rsA = oConexion.EjecutaSelectRS(SQL)
    If Not IsNull(rsA.Fields("m")) Then
        GenerarItemOrRec = rsA.Fields("m") + 1
    Else
        GenerarItemOrRec = 1
    End If
    rsA.CloseRecordset
    Set rsA = Nothing
End Function

Private Function Estado(ByVal vEstado As String) As String
    Dim rsestado As New MYSQL_RS
    Dim SQL As String
    SQL = "doc_estados where descripcion = '" & vEstado & "' "
    Set rsestado = oConexion.EjecutaSelect(SQL)
    If Not rsestado.EOF Then
        Estado = rsestado.Fields("cod_estado")
        Else
            If rsestado.EOF Then
                Estado = "00"
            End If
    End If
    Set rsestado = Nothing
End Function

Private Function InsHistorial(ByVal i As Integer)
    Dim strEstado As String
    Dim SQL As String
    Dim sQuery As String
    Dim Rs As New MYSQL_RS
    Dim idarea As String
    Dim fecha As String
    Dim itemHist As Long
    Dim usuario As String
        strEstado = a.Lines(i).Fields(ORDENCS.ESTADOA).Value
        SQL = "Select * from historial_docs where identificador = '" & id_ordencompra & "'"
        Set Rs = New MYSQL_RS
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        If Rs.Fields("Cod_Estado") <> strEstado Then
        idarea = Rs.Fields("Id_Area")
        strEstado = Rs.Fields("Cod_Estado")
        fecha = Rs.Fields("Fecha_Movi")
        usuario = Rs.Fields("Usuario")
        itemHist = GeneracodigoHist
        sQuery = "Call Insert_HistorialDoc ('" & id_ordencompra & "', '" & strEstado & "','" & idarea & "', '" & fecha & "', '" & usuario & "' );"
        oConexion.EjecutaInsertUpdateDelete sQuery, TIPO_QUERY.insertar, True
            Else: Exit Function
        End If
        Set Rs = Nothing
End Function

Private Function MonedaOC(ByVal nomMoneda As String) As String
    If nomMoneda = "Soles" Then MonedaOC = "N": Exit Function
    If nomMoneda = "Dólares" Or nomMoneda = "Dolares" Then MonedaOC = "E": Exit Function
End Function



Function NoExisteAuxiliarEnCont(ByVal strAuxiliar As String, ByVal strCodigo As String) As Boolean
    Dim SQL As String
    Dim rsR As MYSQL_RS
    NoExisteAuxiliarEnCont = True
    
    SQL = " SELECT CODIGO FROM CNAUXIL WHERE AUXILIAR='" & strAuxiliar & "'  AND CODIGO='" & strCodigo & "'"
    Set rsR = oConexion.EjecutaSelectRS(SQL)
    
   Do While Not rsR.EOF
     NoExisteAuxiliarEnCont = False
     Exit Do
   Loop
    
    Set rsR = Nothing
End Function
