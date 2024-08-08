VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumToLet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private astrConversor(0 To 2, 1 To 9) As String
Private mvarNumero As Variant ' copia local
Private mstrMoneda As String  ' copia local

Public Property Let Numero(ByVal vData As Variant)
    If IsNumeric(vData) Then
        mvarNumero = vData
    End If
End Property
Public Property Get Numero() As Variant
   Numero = mvarNumero
End Property
Public Function ALetra(Optional Nvalor As Boolean) As String
    Dim I As Integer
    Dim intProceder As Integer
    Dim intPosNumero As Integer
    Dim intLongNumero As Integer
    Dim strNumero As String
    Dim strLetraNumero As String
    mvarNumero = Abs(mvarNumero)             'Deberá ser positivo
    strNumero = CStr(Fix(mvarNumero))        'A Cadena
    
    intLongNumero = Len(strNumero)
    intPosNumero = intLongNumero
    For I = 1 To intLongNumero
        intProceder = True
        If (intPosNumero Mod 3) = 1 Then
            If intLongNumero > intPosNumero Then
                Select Case Mid(strNumero, I - 1, 2)
                    Case "00"
                        If UCase(Right(strLetraNumero, 7)) = "CIENTO " Then
                            strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 7)
                            strLetraNumero = strLetraNumero & "CIEN "
                        End If
                    Case "11"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 5)
                        strLetraNumero = strLetraNumero & "ONCE "
                        intProceder = False
                    Case "12"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 5)
                        strLetraNumero = strLetraNumero & "DOCE "
                        intProceder = False
                    Case "13"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 5)
                        strLetraNumero = strLetraNumero & "TRECE "
                        intProceder = False
                    Case "14"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 5)
                        strLetraNumero = strLetraNumero & "CATORCE "
                        intProceder = False
                    Case "15"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 5)
                        strLetraNumero = strLetraNumero & "QUINCE "
                        intProceder = False
                    Case "16", "17", "18", "19"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 2)
                        strLetraNumero = strLetraNumero & "CI"
                    Case "21" To "29"
                        strLetraNumero = Left(strLetraNumero, Len(strLetraNumero) - 2)
                        strLetraNumero = strLetraNumero & "i"
                    Case Else
                        If val(Mid(strNumero, I, 1)) > 0 And val(Mid(strNumero, I - 1, 1)) > 0 Then
                            strLetraNumero = strLetraNumero & "Y "
                        End If
                End Select
            End If
        End If
        If val(Mid(strNumero, I, 1)) > 0 And intProceder Then
            strLetraNumero = strLetraNumero & astrConversor((intPosNumero Mod 3), val(Mid(strNumero, I, 1))) & " "
        End If
            If Nvalor = False Then
                Select Case intPosNumero
                    Case 4
                        If Right(strLetraNumero, 9) <> "MILLONES " And Right(strLetraNumero, 13) <> "MIL MILLONES " And Right(strLetraNumero, 9) <> "BILLONES " Then
                            strLetraNumero = strLetraNumero & "mil "
                        End If
                    Case 7
                        If Right(strLetraNumero, 13) <> "MIL MILLONES " And Right(strLetraNumero, 9) <> "BILLONES " Then
                            strLetraNumero = strLetraNumero & "MILLONES "
                        End If
                    Case 10
                        If Right(strLetraNumero, 9) <> "BILLONES " Then
                            strLetraNumero = strLetraNumero & "MILLARDOS "
                        End If
                    Case 13
                        strLetraNumero = strLetraNumero & "BILLONES "
                    Case Else
                End Select
            End If
            intPosNumero = intPosNumero - 1
        Next I
        strLetraNumero = Left(strLetraNumero, 1) & UCase(Mid(strLetraNumero, 2))
        If Nvalor = False Then
            If Numero <> Fix(Numero) Then
                I = (Numero - Fix(Numero)) * 100
                If Len(str(I)) > 2 Then
                 strLetraNumero = strLetraNumero & " Y " & str(I) & "/100 " & UCase(mstrMoneda)
                Else
                 strLetraNumero = strLetraNumero & " Y 0" & Trim(str(I)) & "/100 " & UCase(mstrMoneda)
                End If
            Else
                strLetraNumero = strLetraNumero & " Y 00/100 " & UCase(mstrMoneda)
            End If
            If UCase(mstrMoneda) = "PESOS" Then
                strLetraNumero = strLetraNumero & " M. N."
            End If
        End If
    ALetra = strLetraNumero
End Function
Private Sub Asigna()
    astrConversor(0, 1) = "CIENTO"
    astrConversor(0, 2) = "DOSCIENTOS"
    astrConversor(0, 3) = "TRESCIENTOS"
    astrConversor(0, 4) = "CUATROCIENTOS"
    astrConversor(0, 5) = "QUINIENTOS"
    astrConversor(0, 6) = "SEISCIENTOS"
    astrConversor(0, 7) = "SETECIENTOS"
    astrConversor(0, 8) = "OCHOCIENTOS"
    astrConversor(0, 9) = "NOVECIENTOS"
    astrConversor(1, 1) = "UN"
    astrConversor(1, 2) = "DOS"
    astrConversor(1, 3) = "TRES"
    astrConversor(1, 4) = "CUATRO"
    astrConversor(1, 5) = "CINCO"
    astrConversor(1, 6) = "SEIS"
    astrConversor(1, 7) = "SIETE"
    astrConversor(1, 8) = "OCHO"
    astrConversor(1, 9) = "NUEVE"
    astrConversor(2, 1) = "DIEZ"
    astrConversor(2, 2) = "VEINTE"
    astrConversor(2, 3) = "TREINTA"
    astrConversor(2, 4) = "CUARENTA"
    astrConversor(2, 5) = "CINCUENTA"
    astrConversor(2, 6) = "SESENTA"
    astrConversor(2, 7) = "SETENTA"
    astrConversor(2, 8) = "OCHENTA"
    astrConversor(2, 9) = "NOVENTA"
End Sub
Public Property Get moneda() As String
    moneda = mstrMoneda
End Property
Public Property Let moneda(ByVal vNewValue As String)
    mstrMoneda = vNewValue
End Property
Private Sub Class_Initialize()
    Asigna
    mstrMoneda = "DOLARES AMERICANOS"
End Sub
Private Sub Class_Terminate()
    Erase astrConversor
End Sub


