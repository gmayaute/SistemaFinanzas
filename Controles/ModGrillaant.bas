Attribute VB_Name = "ModGrilla"
Option Explicit

Public Sub CagarGrillaEdit()
    Dim lsCadena As String
    Dim i As Long
    
    For i = 32 To 255
        lsCadena = lsCadena & Chr(i)
    Next
    
    With frmKardex.flxeditDetalle
'        .Rows = 7
'        .Cols = 0
'        .FormatString = "<Número |<Introduzca una fecha|<String      |<Solo ABCDEFyG"
'        .ColType(0) = Entero
'        .ColMaxLength(0) = 5
'        .ColType(1) = fecha
'        .ColType(2) = Cadena
'        .ColMaxLength(2) = 15
'        .CaracteresValidos(2) = lsCadena
'        .ColType(3) = Cadena
'        .CaracteresValidos(3) = "ABCDEFyG"
       .Rows = 1
       .Cols = 7
       .ColWidth(0) = 500
       .ColAlignment(0) = flexAlignCenterCenter
       .TextMatrix(0, 0) = "ITEM"
       
       .ColWidth(1) = 1300
       .ColAlignment(1) = flexAlignCenterCenter
       .TextMatrix(0, 1) = "CODIGO"
       
       .ColWidth(2) = 3200
       .ColAlignment(2) = flexAlignCenterCenter
       .TextMatrix(0, 2) = "DESCRIPCION"
       
       .ColWidth(3) = 500
       .ColAlignment(3) = flexAlignCenterCenter
       .TextMatrix(0, 3) = "MON"
       
       .ColWidth(4) = 1300
       .ColAlignment(4) = flexAlignCenterCenter
       .TextMatrix(0, 4) = "VALOR"
       
       .ColWidth(5) = 1200
       .ColAlignment(5) = flexAlignCenterCenter
       .TextMatrix(0, 5) = "CANTIDAD"
       
       .ColWidth(6) = 3000
       .ColAlignment(6) = flexAlignCenterCenter
       .TextMatrix(0, 6) = "OSERVACION"
    End With

End Sub

