VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cl_intTipo As flextype 'si es caracter, numero...
Private cl_strNombre As String
Private cl_strDescripcion As String
Private cl_strValor As Variant
Private cl_intTamanio As Integer
Private cl_intValidacion As Integer
Private cl_intVisible As Integer
Private cl_bolActualizable As Boolean
Private cl_intPresicion As Integer
Private cl_CompletaCero As Integer
Private cl_strnomTabla As String
Private cl_strCaracteresValidos As String
Public Property Let CaractValidos(ByVal vCaracteres As String)
    cl_strCaracteresValidos = vCaracteres
End Property
Public Property Get CaractValidos() As String
    CaractValidos = cl_strCaracteresValidos
End Property
Public Property Let Presicion(ByVal vPresicion As Integer)
    cl_intPresicion = vPresicion
End Property
Public Property Get Presicion() As Integer
    Presicion = cl_intPresicion
End Property
Public Property Let CompletaCero(ByVal vCompletaCero As Integer)
    cl_CompletaCero = vCompletaCero
End Property
Public Property Get CompletaCero() As Integer
    CompletaCero = cl_CompletaCero
End Property
Public Property Let Tipo(ByVal vTipo As flextype)
    cl_intTipo = vTipo
End Property
Public Property Get Tipo() As flextype
    Tipo = cl_intTipo
End Property
Public Property Let Nombre(ByVal vNombre As String)
    cl_strNombre = vNombre
End Property
Public Property Get Nombre() As String
    Nombre = cl_strNombre
End Property
Public Property Let Descripcion(ByVal vDescripcion As String)
    cl_strDescripcion = vDescripcion
End Property
Public Property Get Descripcion() As String
    Descripcion = cl_strDescripcion
End Property
Public Property Let Validacion(ByVal vValidacion As Integer)
    cl_intValidacion = vValidacion
End Property
Public Property Get Validacion() As Integer
    Validacion = cl_intValidacion
End Property
Public Property Let Visible(ByVal vVisible As Integer)
    cl_intVisible = vVisible
End Property
Public Property Get Visible() As Integer
    Visible = cl_intVisible
End Property
Public Property Let Actualizable(ByVal vActualizable As Boolean)
    cl_bolActualizable = vActualizable
End Property
Public Property Get Actualizable() As Boolean
    Actualizable = cl_bolActualizable
End Property
Public Property Let Valor(ByVal vValor As String)
  cl_strValor = vValor
End Property
Public Property Get Valor() As String
    Valor = cl_strValor
End Property
Public Property Let Tamanio(ByVal vTamanio As Integer)
  cl_intTamanio = vTamanio
End Property
Public Property Get Tamanio() As Integer
    Tamanio = cl_intTamanio
End Property
Public Property Let TabladelCampo(ByVal vnomTabla As String)
     cl_strnomTabla = vnomTabla
End Property
Public Property Get TabladelCampo() As String
    TabladelCampo = cl_strnomTabla
End Property
