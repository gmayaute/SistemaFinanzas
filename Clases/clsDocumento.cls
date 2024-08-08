VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mCampo As Collection
Private cl_rs As New MYSQL_RS
Private cl_intTipoDoc As Variant
Private cl_intCodFamDoc As Integer
Private cl_strFechaReg As String
Private cl_strHoraReg As String
Public cl_strNomEmpresa As String
Public cl_strNombreMsjro As String
Public cl_strObs As String
Public cl_intTipoIdentMsjro As String
Public cl_strNumIdentMsjro As String
Public cl_strAnomes As String
Public cl_strFlag As Integer
Private cl_strEstado As String

Public Property Let EstadoDoc(ByVal vEstadodoc As String)
    cl_strEstado = vEstadodoc
End Property
Public Property Get EstadoDoc() As String
    EstadoDoc = cl_strEstado
End Property
Public Property Let Rscampos(ByVal vRscampos As MYSQL_RS)
    Set cl_rs = vRscampos
End Property
Public Property Let TipoDoc(ByVal vTipoDoc As Variant)
    cl_intTipoDoc = vTipoDoc
End Property
Public Property Get TipoDoc() As Variant
    TipoDoc = cl_intTipoDoc
End Property
Public Property Let CodFamDoc(ByVal vCodFamDoc As Integer)
    cl_intCodFamDoc = vCodFamDoc
End Property
Public Property Get CodFamDoc() As Integer
    CodFamDoc = cl_intCodFamDoc
End Property
Public Property Let FechaReg(ByVal vFechareg As String)
    cl_strFechaReg = vFechareg
End Property
Public Property Get FechaReg() As String
    FechaReg = cl_strFechaReg
End Property
Public Property Let HoraReg(ByVal vHoraReg As String)
    cl_strHoraReg = vHoraReg
End Property
Public Property Get HoraReg() As String
    HoraReg = cl_strHoraReg
End Property
Public Function Add(intTipo As flextype, strNombre As String, strdescripcion As String, strValor As String, intTamanio As Integer, intValidacion As Integer, intVisible As Integer, bolActualizable As Boolean, strCaractValid As String, intPresicion As Integer, strTablaCampo As String, intCompletaCero As Integer, Optional sKey As String) As clsCampo
    Dim oCampo As clsCampo
    Set oCampo = New clsCampo
    oCampo.Tipo = intTipo
    oCampo.Nombre = strNombre
    oCampo.Descripcion = strdescripcion
    oCampo.valor = strValor
    oCampo.Tamanio = intTamanio
    oCampo.Validacion = intValidacion
    oCampo.Visible = intVisible
    oCampo.Actualizable = bolActualizable
    oCampo.CaractValidos = strCaractValid
    oCampo.Presicion = intPresicion
    oCampo.CompletaCero = intCompletaCero
    oCampo.TabladelCampo = strTablaCampo
    If Len(sKey) = 0 Then
        mCampo.Add oCampo
    Else
        mCampo.Add oCampo, sKey
    End If
    Set Add = oCampo
    Set oCampo = Nothing
End Function
Public Property Get Item(vntIndexKey As Variant) As clsCampo
    Set Item = mCampo(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mCampo.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mCampo.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
    Set NewEnum = mCampo.[_NewEnum]
End Property
Private Sub Class_Initialize()
   Set mCampo = New Collection
   Set cl_rs = New MYSQL_RS
End Sub
Private Sub Class_Terminate()
   Set mCampo = Nothing
   Set cl_rs = Nothing
End Sub
Public Function ObtenerNombreDocumento(vTipo As Variant) As String
    Dim SQL As String
    SQL = "select DESCRIP, Cod_Fam from cndocum C where CODDOC='" & vTipo & "' " & _
          "AND (PROTEGIDO = 'N' OR (SELECT PERMISO FROM `docsusuario` D WHERE D.CODDOC=C.CODDOC " & _
          "AND USUARIO = '" & strUsuarioId & "')=1)"
    Dim Rs As New MYSQL_RS
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    If Rs.RecordCount <> 0 Then
        cl_intCodFamDoc = Rs.Fields("Cod_Fam")
        ObtenerNombreDocumento = Rs.Fields("descrip")
    Else: ObtenerNombreDocumento = Empty
    End If
    Rs.CloseRecordset
    Set Rs = Nothing
End Function

Public Function ConfigCampos()
    Dim Tipo As flextype
    With cl_rs
        Do While Not .EOF
            Select Case .Fields("tipo")
                Case 0
                    Tipo = Entero
                Case 1
                    Tipo = cadena
                Case 2
                    Tipo = fecha
                Case 3
                    Tipo = Numero
                Case Else
                    Tipo = cadena
            End Select
            Add Tipo, CE(.Fields("nom_campo")), CE(.Fields("descripcion")), "", CE(val(.Fields("Tamaño_campo"))), CE(val(.Fields("Validacion"))), CE(val(.Fields("Visible"))), True, CE(.Fields("Caracteres_Validos")), CE(.Fields("Presicion")), CE(.Fields("TablaProcedencia")), CE(.Fields("Completarcero"))
            .MoveNext
        Loop
    End With
    cl_rs.CloseRecordset
    Set cl_rs = Nothing
    End Function
    
Public Function Guardar() As Boolean
On Error GoTo CtrlError
    Dim Query_Insert As String
    Dim Query_Update As String
    Dim est As String
    Dim Fecha_pivot As String
    Dim Fecha_reg As String

    Guardar = False

    Fecha_pivot = Date
    Fecha_reg = Right(cl_strFechaReg, 2) & "/" & Mid(cl_strFechaReg, 6, 2) & "/" & Left(cl_strFechaReg, 4)
    
    If Fecha_pivot <> Fecha_reg Then
     cl_strObs = "Se Registra con Diferente Fecha de la Actual."
    End If
     
    Query_Insert = "Call Insert_AmarreDoc ('" & Item(1).valor & "','" & cl_intTipoDoc & "'," & _
                   " '" & cl_strFechaReg & "', '" & cl_strHoraReg & "'," & _
                   "'" & cl_intTipoIdentMsjro & "','" & cl_strNumIdentMsjro & "'," & _
                   "'" & cl_strNombreMsjro & "', '" & cl_strNomEmpresa & "'," & _
                   "'" & cl_strObs & "','" & cl_intCodFamDoc & "'," & _
                   "'" & strUsuarioId & "', '" & cl_strAnomes & "', '" & cl_strFlag & "');"
    oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, False
    If cl_strFlag = 0 Then
        est = EMITIDO
    Else
        est = REGISTRADO
    End If
    Query_Insert = "Call Insert_Movi_Doc('" & Item(1).valor & "', " & _
                   " '" & Format(Date, "yyyy/mm/dd") & "', '" & est & "' , '" & 1 & "'," & _
                   "'" & strUsuarioId & "'); "
    oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, False
    Query_Insert = "Call Insert_HistorialDoc ('" & Item(1).valor & "', '" & est & "'," & _
                   "'" & DescripcionesdeCodigos("CNUSER", strUsuarioId, "area") & "'," & _
                   "'" & Format(Date, "yyyy/mm/dd") & "', '" & strUsuarioId & "');"
    oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, False
    Select Case cl_intCodFamDoc
        Case FAMILIA_DOC.CONTABLES
            Query_Insert = InsertarDatos("documento_contables")
            oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, True
            
        Case FAMILIA_DOC.ORDENES
            Query_Insert = InsertarDatos("Orden_Compra")
            oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, True
        Case FAMILIA_DOC.ENTIDADES
            Query_Insert = InsertarDatos("documento_entidades")
            oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, True
        Case FAMILIA_DOC.GENERALES
            Query_Insert = InsertarDatos("documento_generales")
            oConexion.EjecutaInsertUpdateDelete Query_Insert, TIPO_QUERY.insertar, True
        Case FAMILIA_DOC.NINGUNO
    End Select
    
    
    
    
    Guardar = True
    
Exit Function
CtrlError:
    Guardar = False
    MsgBox "Error al Guardar Documento. Error: " & err.Description, vbCritical, "NOVPeru"
End Function
Public Function Actualizar() As Boolean
On Error GoTo CtrlError
    Dim SQL As String
    Dim sqlAmarre As String
    Dim Fecha_pivot As String
    Dim Fecha_reg As String
    
    Fecha_pivot = Date
    Fecha_reg = Right(cl_strFechaReg, 2) & "/" & Mid(cl_strFechaReg, 6, 2) & "/" & Left(cl_strFechaReg, 4)
    
    If Fecha_pivot <> Fecha_reg Then
     cl_strObs = "Se Registra con Diferente Fecha de la Actual."
    End If
    
    Actualizar = False
    sqlAmarre = "Call Insert_HistorialDoc ( '" & strIdentificador & "', '" & MODIFICADO & "'," & _
                "'" & DescripcionesdeCodigos("CNUSER", strUsuarioId, "area") & "'," & _
                "'" & Format(Date, "yyyy/mm/dd") & "', '" & strUsuarioId & "');"
    oConexion.EjecutaInsertUpdateDelete sqlAmarre, TIPO_QUERY.insertar, False
    sqlAmarre = "Call Update_AmarreDoc ('" & strIdentificador & "','" & cl_intTipoDoc & "'," & _
                " '" & cl_strFechaReg & "', '" & cl_strHoraReg & "'," & _
                "'" & cl_intTipoIdentMsjro & "','" & cl_strNumIdentMsjro & "'," & _
                "'" & cl_strNombreMsjro & "', '" & cl_strNomEmpresa & "'," & _
                "'" & IIf(FolioAuto = "F.A.", "F.A.MOD", cl_strObs) & "','" & cl_intCodFamDoc & "'," & _
                "'" & strUsuarioId & "', '" & Left(strIdentificador, 6) & "', '" & cl_strFlag & "');"
    oConexion.EjecutaInsertUpdateDelete sqlAmarre, TIPO_QUERY.Modificar, False
    Select Case cl_intCodFamDoc
        Case FAMILIA_DOC.CONTABLES
            SQL = ActualizarDatos("documento_contables")
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Case FAMILIA_DOC.ORDENES
            SQL = ActualizarDatos("Orden_Compra")
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Case FAMILIA_DOC.ENTIDADES
            SQL = ActualizarDatos("documento_entidades")
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Case FAMILIA_DOC.GENERALES
            SQL = ActualizarDatos("documento_generales")
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
        Case FAMILIA_DOC.NINGUNO
    End Select
    Actualizar = True
Exit Function
CtrlError:
    Actualizar = False
    MsgBox "Error al actualizar Documento. Error: " & err.Description, vbCritical, "NOVPeru"
End Function

Public Sub ActCierre(Tipo As String, valor As Integer)
    Dim SQL As String, Ident As String
    If Tipo = 1 Then Ident = Item(1).valor Else Ident = strIdentificador
    SQL = "update documento_contables set contrato = '" & IIf(valor = 1, "FC", "") & "' where identificador = '" & Ident & "'"
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
End Sub
