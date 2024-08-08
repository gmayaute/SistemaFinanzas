VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmVerContratos 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de "
   ClientHeight    =   7755
   ClientLeft      =   6495
   ClientTop       =   6015
   ClientWidth     =   11445
   Icon            =   "frmVerContratos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11445
   Begin VB.Frame Frame3 
      BackColor       =   &H009F5539&
      Caption         =   "Detalle de movimientos"
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
      Height          =   3330
      Left            =   0
      TabIndex        =   5
      Top             =   3915
      Width           =   11415
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexContratos 
         Height          =   3030
         Left            =   75
         TabIndex        =   6
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5345
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   12632256
         BackColorBkg    =   8421504
         HighLight       =   2
         SelectionMode   =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Proyecto1.chameleonButton btnEliminar 
         Height          =   345
         Left            =   10785
         TabIndex        =   10
         ToolTipText     =   "Eliminar"
         Top             =   270
         Width           =   465
         _ExtentX        =   820
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
         MICON           =   "frmVerContratos.frx":014A
         PICN            =   "frmVerContratos.frx":0166
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
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   555
      Left            =   0
      TabIndex        =   2
      Top             =   7170
      Width           =   11415
      Begin Proyecto1.chameleonButton btnSalir 
         Height          =   345
         Left            =   10740
         TabIndex        =   3
         Top             =   165
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
         MICON           =   "frmVerContratos.frx":05A8
         PICN            =   "frmVerContratos.frx":05C4
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
         Left            =   10140
         TabIndex        =   4
         Top             =   165
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
         MICON           =   "frmVerContratos.frx":098A
         PICN            =   "frmVerContratos.frx":09A6
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
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   11415
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
         Left            =   8550
         TabIndex        =   7
         Text            =   "cboFiltro"
         Top             =   690
         Visible         =   0   'False
         Width           =   1845
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flexPersonal 
         Height          =   3285
         Left            =   45
         TabIndex        =   1
         Top             =   150
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   5794
         _Version        =   393216
         BackColor       =   -2147483634
         BackColorFixed  =   12632256
         BackColorBkg    =   8421504
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblCadBus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   9120
      TabIndex        =   12
      Top             =   120
      Width           =   75
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Palabra de Búsqueda (Esc para borrar)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5460
      TabIndex        =   11
      Top             =   90
      Width           =   3495
   End
   Begin MSForms.ComboBox cboPrograma 
      Height          =   315
      Left            =   1590
      TabIndex        =   8
      Top             =   60
      Width           =   3525
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "6218;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   105
      Width           =   1515
      ForeColor       =   -2147483634
      BackColor       =   10442041
      Caption         =   "Programa de:"
      Size            =   "2672;397"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmVerContratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private filtro(1 To 6, 1 To 2) As String
Dim SQL As String
Dim cadenaemp As String
Private Sub ConfigGrillaCont()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 21
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .FixedCols = 1
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Tipo"
        .ColWidth(2) = 2500
        .TextMatrix(0, 2) = Space(17) + "Decripción"
        .ColWidth(3) = 500
        .TextMatrix(0, 3) = Space(0) + "Bono"
        .ColWidth(4) = 1300
        .TextMatrix(0, 4) = Space(3) + "Monto Bono"
        .ColWidth(5) = 1500
        .TextMatrix(0, 5) = Space(4) + "Sueldo Básico"
        .ColWidth(6) = 1500
        .TextMatrix(0, 6) = Space(5) + "Sueldo Bruto"
        .ColWidth(7) = 1200
        .TextMatrix(0, 7) = Space(6) + "Fec Ini"
        .ColWidth(8) = 1200
        .TextMatrix(0, 8) = Space(6) + "Fec Ter"
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = "COD"
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = "Estado"
        .ColWidth(11) = 0
        .TextMatrix(0, 11) = "anomes"
        .ColWidth(12) = 0
        .TextMatrix(0, 12) = "monsueldo"
        .ColWidth(13) = 0
        .TextMatrix(0, 13) = "monbono"
        .ColWidth(14) = 0
        .TextMatrix(0, 14) = "esttrabajo"
        .ColWidth(15) = 0
        .TextMatrix(0, 15) = "horlab"
        .ColWidth(16) = 0
        .TextMatrix(0, 16) = "cencos"
        .ColWidth(17) = 0
        .TextMatrix(0, 17) = "divgas"
        .ColWidth(18) = 0
        .TextMatrix(0, 18) = "ccHFM"
        .ColWidth(19) = 0
        .TextMatrix(0, 19) = "tcese"
        .ColWidth(20) = 0
        .TextMatrix(0, 20) = "fcese"
         For I = 0 To 20
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
         Next I
    End With
End Sub
Private Sub ConfigVaca()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 7
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = Space(0) + "Periodo"
        .ColWidth(3) = 2200
        .TextMatrix(0, 3) = Space(17) + "Decripción"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = Space(6) + "Fec. Sal"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(6) + "Fec. Reg"
        .ColWidth(6) = 1200
        .TextMatrix(0, 6) = Space(6) + "Dias"
         For I = 0 To 6
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
         Next I
    End With
End Sub
Private Sub ConfigPermiso()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 8
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = Space(6) + "Fec. Sal"
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = Space(6) + "Hora Sal."
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = Space(6) + "Fec. Reg"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(6) + "Hora Reg."
        .ColWidth(6) = 5000
        .TextMatrix(0, 6) = Space(45) + "Motivo"
        .ColWidth(7) = 0
        .TextMatrix(0, 7) = "Autorizado"
        For I = 0 To 7
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
        Next I
    End With
End Sub
Private Sub ConfigMovilidades()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 8
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 460
        .TextMatrix(0, 2) = Space(0) + "Tipo"
        .ColWidth(3) = 1080
        .TextMatrix(0, 3) = Space(3) + "Fec.Salida"
        .ColWidth(4) = 1080
        .TextMatrix(0, 4) = Space(3) + "Fec.Ingreso"
        .ColWidth(5) = 1000
        .TextMatrix(0, 5) = Space(6) + "Lote"
        .ColAlignment(5) = 6
        .ColWidth(6) = 1500
        .TextMatrix(0, 6) = Space(6) + "Pozo"
        .ColAlignment(6) = 6
        .ColWidth(7) = 4580
        .TextMatrix(0, 7) = Space(6) + "Observaciones"
        For I = 0 To 7
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
        Next I
    End With
End Sub
Private Sub ConfigLicencia()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = Space(6) + "Fec. Sal"
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = Space(6) + "Fec. Reg"
        .ColWidth(4) = 5000
        .TextMatrix(0, 4) = Space(45) + "Motivo"
        .ColWidth(5) = 0
        .TextMatrix(0, 5) = "Autorizado"
        For I = 0 To 5
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
        Next I
    End With
End Sub
Private Sub ConfigSubSidios()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 7
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = Space(6) + "Fec. Sal"
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = Space(6) + "Fec. Reg"
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = Space(6) + "CIIT"
        .ColWidth(5) = 3400
        .TextMatrix(0, 5) = Space(6) + "Tipo Suspensión"
        .ColWidth(6) = 3600
        .TextMatrix(0, 6) = Space(45) + "Motivo"
        For I = 0 To 6
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
        Next I
    End With
End Sub
Private Sub ConfigCampo()
    With flexContratos
        .Clear
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 300
        .ColWidth(0) = 450
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 0
        .TextMatrix(0, 1) = Space(0) + "Codigo"
        .ColWidth(2) = 800
        .TextMatrix(0, 2) = Space(0) + "Lote"
        .ColWidth(3) = 2000
        .TextMatrix(0, 3) = Space(17) + "Pozo"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = Space(6) + "Fec. Sal"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(6) + "Fec. Reg"
         For I = 0 To 5
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
         Next I
    End With
End Sub
Private Sub ConfigGrillaEmp()
    With flexPersonal
        .Clear
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 350
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(0) + "Item"
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = Space(5) + "Codigo"
        .ColWidth(2) = 4000
        .TextMatrix(0, 2) = Space(20) + "Nombre Completo"
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = Space(8) + "Situacion"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = Space(5) + "Fec_Ingreso"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = Space(6) + "Fec_Cese"
         For I = 0 To 5
            .row = 0
            .Col = I
            .CellBackColor = &HC0C0C0
            .CellForeColor = &H800000
         Next I
    End With
End Sub
Private Sub CargarDatosEmp(SQL As String)
    Dim Rs As MYSQL_RS
    Dim I As Integer, J As Integer
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    I = 1
    With flexPersonal
        Do While Not Rs.EOF
            .TextMatrix(I, 0) = I
            .TextMatrix(I, 1) = Rs.Fields("Codigo")
            .TextMatrix(I, 2) = Space(2) & UCase(Rs.Fields("Nombres"))
            .TextMatrix(I, 3) = Space(8) & IIf(val(Rs.Fields("Situacion")) = 0, "CESADO", "ACTIVO")
            .TextMatrix(I, 4) = IIf(Rs.Fields("Fec_Ingreso") <> "", Format(Rs.Fields("Fec_Ingreso"), "dd/mm/yyyy"), Empty)
            .TextMatrix(I, 5) = IIf(Rs.Fields("Fec_Cese") <> "", Format(Rs.Fields("Fec_Cese"), "dd/mm/yyyy"), Empty)
            If Rs.Fields("Situacion") = "0" Then
                For J = 1 To .Cols - 1
                    .row = I
                    .Col = J
                    .CellBackColor = &H8080FF
                Next
            End If
            .Rows = .Rows + 1
            I = I + 1
            Rs.MoveNext
        Loop
        .Rows = .Rows - 1
    End With
    Set Rs = Nothing
End Sub
Private Sub btnEliminar_Click()
Dim Flg As Boolean
Dim lastRow As Integer
    Flg = False
    With flexContratos
        lastRow = .row
        Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
            Case "02"
                If MsgBox("¿Seguro que desea eliminar las " & IIf(.TextMatrix(.row, 3) = "SISTEMA", "VACACIONES x SISTEMA", .TextMatrix(.row, 3)) & " - " & .TextMatrix(.row, 2) & "?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    SQL = "delete from calendario where codigo='" & .TextMatrix(.row, 1) & "' and codemp='" & flexPersonal.TextMatrix(flexPersonal.row, 1) & "' and movemp='" & frmPrograma.MovEmp("VACACIONES") & "'"
                    oConexionMYSQL.Execute SQL
                    Dim RQ As MYSQL_RS
                    SQL = "SELECT * from pl_tareo where emp = '" & Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)) & "' and tipo = 4 and anomes = '" & Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2) & "'"
                    Set RQ = oConexion.EjecutaSelectRS(SQL)
                    If Not RQ.EOF() Then
                        If val(Month(.TextMatrix(.row, 4))) = val(Month(.TextMatrix(.row, 5))) Then
                            MsgBox "Se eliminará el Tareo de Vacaciones del Año/Mes: " & Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2), vbInformation, "NOVPeru"
                            SQL = "delete from pl_tareo where emp = '" & Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)) & "' and tipo = 4 and anomes = '" & Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2) & "'"
                            oConexionMYSQL.Execute SQL
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "P", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "A", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                        Else
                            MsgBox "Se eliminará el Tareo de Vacaciones del Año/Mes: " & Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2) & " y del Año/Mes: " & Year(CDate(.TextMatrix(.row, 5))) & Right("00" & Month(CDate(.TextMatrix(.row, 5))), 2), vbInformation, "NOVPeru"
                            SQL = "delete from pl_tareo where emp = '" & Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)) & "' and tipo = 4 and anomes = '" & Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2) & "'"
                            oConexionMYSQL.Execute SQL
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "P", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 4))) & Right("00" & Month(CDate(.TextMatrix(.row, 4))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "A", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                            SQL = "delete from pl_tareo where emp = '" & Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)) & "' and tipo = 4 and anomes = '" & Year(CDate(.TextMatrix(.row, 5))) & Right("00" & Month(CDate(.TextMatrix(.row, 5))), 2) & "'"
                            oConexionMYSQL.Execute SQL
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 5))) & Right("00" & Month(CDate(.TextMatrix(.row, 5))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "P", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                            ActualizaMontos Year(CDate(.TextMatrix(.row, 5))) & Right("00" & Month(CDate(.TextMatrix(.row, 5))), 2), 4, IIf(val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 4 Or val(flexPersonal.TextMatrix(flexPersonal.row, 1)) = 38, "E", "N"), "A", Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), 1
                        End If
                    End If
                    Set RQ = Nothing
                    Flg = True
                End If
            Case "03"
                If MsgBox("¿Seguro que desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    SQL = "delete from calendario where codigo='" & .TextMatrix(.row, 1) & "' and codemp='" & flexPersonal.TextMatrix(flexPersonal.row, 1) & "' and movemp='" & frmPrograma.MovEmp("PERMISOS") & "'"
                    oConexionMYSQL.Execute SQL
                    Flg = True
                End If
            Case "04"
                If MsgBox("¿Seguro que desea eliminar el contrato de Fecha Inicial " & .TextMatrix(.row, 7) & " y Fecha Final " & flexContratos.TextMatrix(flexContratos.row, 8) & "?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    ActualizarContratos Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)), Trim(.TextMatrix(.row, 9))
                    oConexionMYSQL.Execute SQL
                    Flg = True
                End If
            Case "05"
                If MsgBox("¿Seguro que desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    SQL = "delete from calendario where codigo='" & .TextMatrix(.row, 1) & "' and codemp='" & flexPersonal.TextMatrix(flexPersonal.row, 1) & "' and movemp='" & frmPrograma.MovEmp("LICENCIA") & "'"
                    oConexionMYSQL.Execute SQL
                    Flg = True
                End If
            Case "07"
                If MsgBox("¿Seguro que desea eliminar la fila seleccionada?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    SQL = "delete from calendario where codigo='" & .TextMatrix(.row, 1) & "' and codemp='" & flexPersonal.TextMatrix(flexPersonal.row, 1) & "' and movemp='" & frmPrograma.MovEmp("SUBSIDIOS") & "'"
                    oConexionMYSQL.Execute SQL
                    Flg = True
                End If
        End Select
        If Flg = True Then
            flexPersonal_RowColChange
            .row = lastRow
            .Col = 1
            .SetFocus
            
            Call keybd_event(vbKeyHome, 0, 0, 0)
        End If
    End With
End Sub
Private Sub btnReporte_Click()
    Dim Tit As String, fec As Variant
    Dim emp As String, Divi As String
    If flexPersonal.ColSel = 5 Then emp = Trim(flexPersonal.TextMatrix(flexPersonal.row, 1)) Else emp = ""
    Set oReporte = New clsReporte
    oReporte.empresa = strNombreEmpresa
    Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
        Case "01":
            Tit = "SALIDAS  A  CAMPO"
        
        Case "02":
            Tit = "VACACIONES"
            oReporte.Titulo = "HISTORIAL  DE  " & Tit
            oReporte.Reporte = "Rep_VacacionesTrab.rpt"
            fec = InputBox("Ingrese fecha de cálculo de vacaciones", "FECHA", Date)
            If fec = "" Then Exit Sub
            Divi = ""
            If emp = "" Then
                If MsgBox("¿Todas las ccHFM?", vbQuestion + vbYesNo, "NOVPeru") = vbNo Then
                    Divi = InputBox("Ingrese Código de la ccHFM", "ccHFM", "0006")
                    Divi = Right("0000" & Divi, 4)
                End If
            End If
            oReporte.sp_Rep_VacacionesTrab emp, IIf(fec = "", Format(Date, "yyyy/mm/dd"), Format(fec, "yyyy/mm/dd")), Divi
        Case "03":
            Tit = "PERMISOS"
            oReporte.Titulo = "HISTORIAL  DE  " & Tit
            oReporte.Reporte = "Rep_PermisosTrab.rpt"
            oReporte.sp_Rep_PermisosTrab emp, "03"
        Case "04":
            Tit = "CONTRATOS"
            oReporte.Titulo = "HISTORIAL  DE  " & Tit
            oReporte.Reporte = "Rep_ContratosTrab.rpt"
            oReporte.sp_Rep_ContratosTrab emp
        Case "05":
            Tit = "LICENCIAS"
            oReporte.Titulo = "HISTORIAL  DE  " & Tit
            oReporte.Reporte = "Rep_PermisosTrab.rpt"
            oReporte.sp_Rep_PermisosTrab emp, "05"
        Case "06": Tit = "REFRIGERIOS"
        Case "07":
            Tit = "SUBSIDIOS"
            oReporte.Titulo = "HISTORIAL  DE  " & Tit
            oReporte.Reporte = "Rep_SubsidiosTrab.rpt"
            oReporte.sp_Rep_SubsidiosTrab emp, "07"
    End Select
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub LimpiarVector()
    Dim I As Integer, J As Integer
    For I = 1 To 6
        For J = 1 To 2
            filtro(I, J) = ""
        Next
    Next
End Sub
Private Sub cboFiltro_Click()
    Dim strcampo As String
    Dim SQL As String, sqlwhere As String
    Dim I As Integer
    strcampo = Trim(flexPersonal.TextMatrix(0, flexPersonal.Col))
    If strcampo = "Nombre Completo" Then strcampo = "Concat(ApePat,' ',ApeMat,' ',Nombre1,' ',Nombre2)"
    If cboFiltro.Text <> "( Todos )" Then
        filtro(flexPersonal.Col + 1, 1) = Trim(strcampo)
        filtro(flexPersonal.Col + 1, 2) = Trim(cboFiltro.Text)
    Else
        filtro(flexPersonal.Col + 1, 1) = ""
        filtro(flexPersonal.Col + 1, 2) = ""
    End If
    SQL = " Select Codigo, Concat(ApePat,' ',ApeMat,' ',Nombre1,' ',Nombre2) as Nombres, Situacion, " & _
          " Fec_Ingreso, Fec_Cese from empleado "
    sqlwhere = " where "
    For I = 1 To 5
        If filtro(I, 2) <> "" Then sqlwhere = sqlwhere & filtro(I, 1) & "='" & filtro(I, 2) & "' and "
    Next
    If sqlwhere <> " where " Then SQL = SQL & Left(Trim(sqlwhere), Len(Trim(sqlwhere)) - 3)
    ConfigGrillaEmp
    ConfigGrillaCont
    CargarDatosEmp SQL & " ORDER BY APEPAT,APEMAT,NOMBRE1,NOMBRE2"
    For I = 1 To flexPersonal.Cols
        If filtro(I, 2) <> Empty Then
            With flexPersonal
                .Col = I - 1
                .row = 0
                .CellForeColor = vbBlue
            End With
        Else
            With flexPersonal
                .Col = I - 1
                .row = 0
                .CellForeColor = &H800000
            End With
        End If
    Next
    If flexPersonal.Rows > 1 Then VerContrato flexPersonal.TextMatrix(1, 1), 1
    cboFiltro.Visible = False
End Sub
Private Sub cboFiltro_DropDown()
    Dim strcampo As String
    Dim rsFiltro As MYSQL_RS
    strcampo = cboFiltro.Text
    cboFiltro.Clear
    cboFiltro.AddItem "( Todos )"
    cboFiltro.Text = strcampo
    If strcampo = "Nombre Completo" Then strcampo = "Concat(ApePat,' ',ApeMat,' ',Nombre1,' ',Nombre2)"
    SQL = "Select " & strcampo & " from empleado group by 1 order by 1"
    Set rsFiltro = oConexion.EjecutaSelectRS(SQL)
    Do While Not rsFiltro.EOF
        cboFiltro.AddItem UCase(rsFiltro.Fields(0))
        rsFiltro.MoveNext
    Loop
    If cboFiltro.Width < 1200 Then cboFiltro.Width = 1400
    Set rsFiltro = Nothing
End Sub
Private Sub cboFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then cboFiltro.Visible = False
End Sub
Private Sub cboPrograma_Change()
    LimpiarVector
    flexPersonal.Redraw = False
    SQL = " Select Codigo, Concat(ApePat,' ',ApeMat,' ',Nombre1,' ',NOMBRE2) as Nombres, Situacion, " & _
          " Fec_Ingreso, Fec_Cese from empleado where tipo not in (3,4) Order by APEPAT,APEMAT,NOMBRE1,NOMBRE2"
    ConfigGrillaEmp
    CargarDatosEmp SQL
    Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
        Case "01":
            Me.Caption = "Historial de salidas a campo"
            If flexPersonal.Rows > 1 Then VerCampo Trim(flexPersonal.TextMatrix(1, 1)), 1, "01"
         Case "02":
            Me.Caption = "Historial de Vacaciones"
            If flexPersonal.Rows > 1 Then VerVaca Trim(flexPersonal.TextMatrix(1, 1)), 1, "02"
        Case "04":
            Me.Caption = "Historial de Contratos"
            If flexPersonal.Rows > 1 Then VerContrato Trim(flexPersonal.TextMatrix(1, 1)), 1
    End Select
    flexPersonal.Redraw = True
End Sub
Private Sub flexContratos_DblClick()
    Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
        Case "02", "03", "05", "07"
            With frmPrograma
                If flexContratos.TextMatrix(0, 1) <> Empty Then
                    If flexContratos.row > 0 Then
                        If Trim(flexContratos.TextMatrix(flexContratos.row, 1)) = Empty Then
                            Exit Sub
                        Else
                            .SSTab1.TabVisible(0) = False
                            .SSTab1.TabVisible(1) = True
                            .SSTab1.TabVisible(2) = False
                            .SSTab1.Tab = 1
                            Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
                                Case "02":
                                    .tag = "VACACIONES"
                                    .optFrame(0).Value = True
                                Case "03":
                                    .tag = "PERMISOS"
                                    .optFrame(1).Value = True
                                Case "05":
                                    .tag = "LICENCIA"
                                    .optFrame(2).Value = True
                                Case "07":
                                    .tag = "SUBSIDIOS"
                                    .optFrame(3).Value = True
                            End Select
                            .btnmodificar.Visible = True
                            .CargaTab1 flexContratos.TextMatrix(flexContratos.row, 1), flexPersonal.TextMatrix(flexPersonal.row, 1)
                            .lblModo = "Modificar"
                            If cboPrograma.List(cboPrograma.ListIndex, 1) = "02" Then
                                For I = 0 To .cmbPeriodo.ListCount - 1
                                    If .cmbPeriodo.List(I) = Trim(flexContratos.TextMatrix(flexContratos.row, 2)) Then .cmbPeriodo.ListIndex = I
                                Next
                            End If
                        End If
                    End If
                End If
            End With
        Case "04"
            If Trim(flexPersonal.TextMatrix(flexPersonal.row, 3)) <> "CESADO" Then
                With frmContrato
                    If flexContratos.TextMatrix(0, 1) <> Empty Then
                        If flexContratos.row > 0 Then
                            If Trim(flexContratos.TextMatrix(flexContratos.row, 1)) = Empty Then
                                Exit Sub
                            Else
                                .lblModo = "Consulta"
                                .lblfecfin = flexContratos.TextMatrix(flexContratos.row, 8)
                                .DatosContrato flexPersonal.TextMatrix(flexPersonal.row, 1), _
                                               flexContratos.TextMatrix(flexContratos.row, 9), Format(flexContratos.TextMatrix(flexContratos.row, 7), "yyyy/mm/dd"), "I"
                                .Show
                            End If
                        End If
                    End If
                End With
            End If
    End Select
End Sub
Private Sub flexContratos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If flexContratos.row > 0 Then btnEliminar_Click
    End If
End Sub
Private Sub flexContratos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
        Case "04"
            If Button = vbRightButton And flexContratos.TextMatrix(flexContratos.row, 0) <> "" Then
                If MsgBox("¿Desea agregar un Contrato?", vbQuestion + vbYesNo, "NOVPeru") = vbYes Then
                    NuevaFila Trim(flexPersonal.TextMatrix(flexPersonal.row, 1))
                End If
            End If
    End Select
End Sub
Private Sub flexPersonal_Click()
     If flexPersonal.row = 0 Then
        With cboFiltro
            .Top = flexPersonal.CellTop + flexPersonal.Top
            .Left = flexPersonal.CellLeft + flexPersonal.Left
            .Width = flexPersonal.CellWidth
            .Text = Trim(flexPersonal.Text)
            .Visible = True
            .ZOrder
            .SetFocus
            .SelStart = Len(.Text)
        End With
    Else
        cboFiltro.Visible = False
    End If
End Sub
Private Sub VerMovilidades(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = " Select c.codigo,c.codemp,c.movemp,c.fec_salida,c.hora_salida,c.fec_regreso,c.hora_regreso," & _
          " c.depto,(select descripcioncorta from novperuvhse.lote where idlote=c.lote) as lote," & _
          "(select descripcioncorta from novperuvhse.pozo where idpozo=c.pozo) as pozo," & _
          "c.codagencia,c.codlinea,c.tipoboleto,c.mon_boleto,c.codestancia,c.pagoestancia," & _
          "c.mon_estancia,c.monto_estancia,c.mon_viatico,c.monto_viatico,c.observacion," & _
          "c.sinbono , c.Periodo, c.gocehaber, c.autorizado, c.fec_autorizacion" & _
          " from calendario as c " & _
          "where codemp = '" & CodEmp & "' and movemp='" & Tipo & "' ORDER BY CODIGO"
    
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    flexContratos.Redraw = False
    ConfigMovilidades
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = IIf(rsver.Fields("fec_salida") = "", "IN", "SA")
                .TextMatrix(I, 3) = Format(rsver.Fields("fec_salida"), "dd/mm/yyyy")
                .TextMatrix(I, 4) = Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy")
                .TextMatrix(I, 5) = rsver.Fields("lote")
                .TextMatrix(I, 6) = rsver.Fields("pozo")
                .TextMatrix(I, 7) = rsver.Fields("observacion")
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
            .SetFocus
        End With
    Loop
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    flexPersonal.row = fila
    flexPersonal.Col = 1
    If flexPersonal.Visible = True Then flexPersonal.SetFocus
    Set rsver = Nothing
    flexContratos.Redraw = True
End Sub
Private Sub VerContrato(CodEmp As String, fila As Integer)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = " Select codtipo, (select descrip from cncontrato where codigo=codtipo) as descrip," & _
          " bono,monto_bono, sbasico,0 as sbruto,f_inicio,f_termino,codigo,estado,anomes, " & _
          "mon_sueldo,mon_bono,esttrabajo,horlab,TRIM(cencos) AS CENCOS,divgas,division,codtipocese,fechacese from contrato where codemp = '" & CodEmp & "'"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    ConfigGrillaCont
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codtipo")
                .TextMatrix(I, 2) = rsver.Fields("descrip")
                .TextMatrix(I, 3) = rsver.Fields("bono")
                .TextMatrix(I, 4) = FormatNumber(rsver.Fields("monto_bono"), 2)
                .TextMatrix(I, 5) = FormatNumber(CEN(rsver.Fields("sbasico")), 2)
                .TextMatrix(I, 6) = FormatNumber(CEN(rsver.Fields("sbruto")), 2)
                .TextMatrix(I, 7) = IIf(rsver.Fields("f_inicio") <> Empty, Format(rsver.Fields("f_inicio"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 8) = IIf(rsver.Fields("f_termino") <> Empty, Format(rsver.Fields("f_termino"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 9) = rsver.Fields("codigo")
                .TextMatrix(I, 10) = rsver.Fields("estado")
                .TextMatrix(I, 11) = rsver.Fields("anomes")
                .TextMatrix(I, 12) = rsver.Fields("mon_sueldo")
                .TextMatrix(I, 13) = rsver.Fields("mon_bono")
                .TextMatrix(I, 14) = rsver.Fields("esttrabajo")
                .TextMatrix(I, 15) = rsver.Fields("horlab")
                .TextMatrix(I, 16) = rsver.Fields("cencos")
                .TextMatrix(I, 17) = rsver.Fields("divgas")
                .TextMatrix(I, 18) = rsver.Fields("division")
                .TextMatrix(I, 19) = rsver.Fields("codtipocese")
                .TextMatrix(I, 20) = rsver.Fields("fechacese")
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
        End With
    Loop
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    Set rsver = Nothing
End Sub
Private Sub VerVaca(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer, J As Integer
    SQL = " Select * from calendario where codemp = '" & CodEmp & "' and movemp='" & Tipo & "' ORDER BY periodo,fec_salida"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    flexContratos.Redraw = False
    ConfigVaca
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = rsver.Fields("periodo")
                .TextMatrix(I, 3) = IIf(rsver.Fields("gocehaber") = "S", "Compra de vacaciones", IIf(rsver.Fields("gocehaber") = "N", "Vacaciones físicas", "Sistema"))
                .TextMatrix(I, 4) = IIf(rsver.Fields("fec_salida") <> Empty, Format(rsver.Fields("fec_salida"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 5) = IIf(rsver.Fields("fec_regreso") <> Empty, Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 6) = IIf(rsver.Fields("monto_viatico") <> Empty, rsver.Fields("monto_viatico"), Empty)
                If CDate(flexPersonal.TextMatrix(flexPersonal.row, 4)) >= CDate(.TextMatrix(I, 4)) Then
                    For J = 1 To .Cols - 1
                        .row = I
                        .Col = J
                        .CellBackColor = &H8080FF
                    Next
                End If
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
        End With
    Loop
    flexContratos.Redraw = True
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    Set rsver = Nothing
End Sub
Private Sub VerPermiso(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = " Select * from calendario where codemp = '" & CodEmp & "' and movemp='" & Tipo & "' ORDER BY CODIGO"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    flexContratos.Redraw = False
    ConfigPermiso
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = IIf(rsver.Fields("fec_salida") <> Empty, Format(rsver.Fields("fec_salida"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 3) = rsver.Fields("hora_salida")
                .TextMatrix(I, 4) = IIf(rsver.Fields("fec_regreso") <> Empty, Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 5) = rsver.Fields("hora_regreso")
                .TextMatrix(I, 6) = rsver.Fields("observacion")
                .TextMatrix(I, 7) = rsver.Fields("autorizado")
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
            .SetFocus
        End With
    Loop
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    flexPersonal.row = fila
    flexPersonal.Col = 1
    If flexPersonal.Visible = True Then flexPersonal.SetFocus
    Set rsver = Nothing
    flexContratos.Redraw = True
End Sub
Private Sub VerLicencia(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = " Select * from calendario where codemp = '" & CodEmp & "' and movemp='" & Tipo & "' ORDER BY CODIGO"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    flexContratos.Redraw = False
    ConfigLicencia
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = IIf(rsver.Fields("fec_salida") <> Empty, Format(rsver.Fields("fec_salida"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 3) = IIf(rsver.Fields("fec_regreso") <> Empty, Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 4) = rsver.Fields("observacion")
                .TextMatrix(I, 5) = rsver.Fields("autorizado")
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
            .SetFocus
        End With
    Loop
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    flexPersonal.row = fila
    flexPersonal.Col = 1
    If flexPersonal.Visible = True Then flexPersonal.SetFocus
    Set rsver = Nothing
    flexContratos.Redraw = True
End Sub
Private Sub VerSubSidios(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = " Select * from calendario where codemp = '" & CodEmp & "' and movemp='" & Tipo & "' ORDER BY CODIGO"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    flexContratos.Redraw = False
    ConfigSubSidios
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = IIf(rsver.Fields("fec_salida") <> Empty, Format(rsver.Fields("fec_salida"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 3) = IIf(rsver.Fields("fec_regreso") <> Empty, Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 4) = rsver.Fields("sinbono")
                .TextMatrix(I, 5) = DescripcionesdeCodigos("TIPOSUSPENSION", rsver.Fields("dpto"))
                .TextMatrix(I, 6) = rsver.Fields("observacion")
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
            .SetFocus
        End With
    Loop
    If I > 0 Then btnEliminar.Enabled = True Else btnEliminar.Enabled = False
    flexPersonal.row = fila
    flexPersonal.Col = 1
    If flexPersonal.Visible = True Then flexPersonal.SetFocus
    Set rsver = Nothing
    flexContratos.Redraw = True
End Sub
Private Sub VerCampo(CodEmp As String, fila As Integer, Tipo As String)
    Dim rsver As MYSQL_RS
    Dim I As Integer
    SQL = "Select c.codigo,c.codemp,c.movemp,c.fec_salida,c.hora_salida,c.fec_regreso,c.hora_regreso," & _
          "c.depto,(select descripcioncorta from novperuvhse.lote where idlote=c.lote) as lote," & _
          "(select descripcioncorta from novperuvhse.pozo where idpozo=c.pozo) as pozo," & _
          "c.codagencia,c.codlinea,c.tipoboleto,c.mon_boleto,c.codestancia,c.pagoestancia," & _
          "c.mon_estancia,c.monto_estancia,c.mon_viatico,c.monto_viatico,c.observacion," & _
          "c.sinbono , c.Periodo, c.gocehaber, c.autorizado, c.fec_autorizacion " & _
          "from calendario as c " & _
          "where codemp = '" & CodEmp & "' and movemp='" & Tipo & "'"
    Set rsver = oConexion.EjecutaSelectRS(SQL)
    ConfigCampo
    Do While Not rsver.EOF
        With flexContratos
            For I = 1 To rsver.RecordCount
                .Rows = .Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsver.Fields("Codigo")
                .TextMatrix(I, 2) = rsver.Fields("lote")
                .TextMatrix(I, 3) = rsver.Fields("pozo")
                .TextMatrix(I, 4) = IIf(rsver.Fields("fec_salida") <> Empty, Format(rsver.Fields("fec_salida"), "dd/mm/yyyy"), Empty)
                .TextMatrix(I, 5) = IIf(rsver.Fields("fec_regreso") <> Empty, Format(rsver.Fields("fec_regreso"), "dd/mm/yyyy"), Empty)
                rsver.MoveNext
            Next
            .row = 1
            .Col = 1
            .SetFocus
        End With
    Loop
    flexPersonal.row = fila
    flexPersonal.Col = 1
    If flexPersonal.Visible = True Then flexPersonal.SetFocus
    Set rsver = Nothing
End Sub
Private Sub flexPersonal_KeyPress(KeyAscii As Integer)
    Dim F As Integer
    If flexPersonal.Col < 3 Then
        Dim c%, T%, a$, B$
        If KeyAscii = 0 Or KeyAscii = 27 Then
            cadenaemp = "": lblCadBus = ""
            Exit Sub
        End If
        If KeyAscii >= 32 Or KeyAscii = 8 Then
            cadenaemp = cadenaemp & Chr(KeyAscii)
            lblCadBus = cadenaemp
            With flexPersonal
                If KeyAscii <> 8 Then
                    c = Len(cadenaemp)
                    If IsNumeric(cadenaemp) Then a = Right("00000000000" & Trim(cadenaemp), 11) Else a = cadenaemp
                End If
                If c >= 1 Then
                    For T = 1 To .Rows - 1
                        If IsNumeric(cadenaemp) Then B = Trim(.TextMatrix(T, 1)) Else B = Trim(.TextMatrix(T, 2))
                        If Len(B) >= c Then
                            B = Left(B, c)
                            If Trim(a) = Trim(B) Then
                                KeyAscii = 0
                                ItemLista = T
                                .row = T
                                .Col = 1
                                .ColSel = 5
                                If T >= 6 Then .TopRow = T - 5 Else .TopRow = T
                                flexPersonal_RowColChange
                                Exit For
                            End If
                        End If
                    Next T
                End If
            End With
        End If
    End If
End Sub
Private Sub flexPersonal_RowColChange()
    With flexPersonal
        If .row > 0 Then
            .SelectionMode = flexSelectionFree
            Select Case cboPrograma.List(cboPrograma.ListIndex, 1)
                Case "01":
                    If flexPersonal.Rows > 1 Then VerMovilidades Trim(.TextMatrix(.row, 1)), .row, "01"
                Case "02":
                    If flexPersonal.Rows > 1 Then VerVaca Trim(.TextMatrix(.row, 1)), .row, "02"
                Case "03":
                    If flexPersonal.Rows > 1 Then VerPermiso Trim(.TextMatrix(.row, 1)), .row, "03"
                Case "04":
                    If flexPersonal.Rows > 1 Then VerContrato Trim(.TextMatrix(.row, 1)), .row
                Case "05":
                    If flexPersonal.Rows > 1 Then VerLicencia Trim(.TextMatrix(.row, 1)), .row, "05"
                Case "07":
                    If flexPersonal.Rows > 1 Then VerSubSidios Trim(.TextMatrix(.row, 1)), .row, "07"
            End Select
        End If
    End With
End Sub
Private Sub flexPersonal_SelChange()
    flexPersonal.Rowsel = flexPersonal.row
End Sub
Private Sub Form_Load()
    Dim SQL As String
    Me.Top = 0
    Me.Left = 0
    cadenaemp = ""
    lblCadBus = ""
    MovEmp cboPrograma
    Call WheelHook(frmVerContratos)
End Sub
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
On Error Resume Next
    With flexPersonal
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then Lstep = 1
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then NewValue = 0
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then NewValue = .Rows - 1
        End If
        .TopRow = NewValue
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call WheelUnHook
End Sub
Sub ActualizarContratos(CodEmp As String, Cod As String)
    SQL = "delete from contrato where codemp = '" & CodEmp & "' and codigo = '" & Cod & "'"
    oConexionMYSQL.Execute SQL
    ActualizarCodContrato CodEmp, Cod, "asc"
End Sub
Private Sub NuevaFila(CodEmp As String)
    Dim lastRow As Long
    With flexContratos
        lastRow = .row
        If .row = lastRow Then
            lastRow = lastRow + 1
            ActualizarCodContrato CodEmp, .TextMatrix(.row, 9), "desc", "CN" & Right("0000" & val(Right(.TextMatrix(.Rows - 2, 9), 2)) + 1, 2)
            InsertarContrato .TextMatrix(.row, 9), CodEmp, .row
            flexPersonal_RowColChange
        End If
        .row = lastRow
        .Col = 1
        .SetFocus
        Call keybd_event(vbKeyHome, 0, 0, 0)
    End With
End Sub
Sub ActualizarCodContrato(CodEmp As String, Cod As String, orden As String, Optional UltCod As String)
    Dim RQ As MYSQL_RS
    Dim SCond As String
    SQL = "Select * from contrato where codemp = '" & CodEmp & "' and codigo > '" & Cod & "' order by codigo " & orden
    Set RQ = oConexion.EjecutaSelectRS(SQL)
    If UltCod <> "" Then Cod = UltCod
    Do While Not RQ.EOF
        SQL = "update contrato set codigo = '" & Cod & "' where codemp = '" & CodEmp & "' and codigo = '" & Trim(RQ.Fields("codigo")) & "'"
        oConexionMYSQL.Execute SQL
        If orden = "asc" Then Cod = "CN" & Right("0000" & val(Right(Cod, 2)) + 1, 2) Else Cod = "CN" & Right("0000" & val(Right(Cod, 2)) - 1, 2)
        RQ.MoveNext
    Loop
    Set RQ = Nothing
End Sub
Sub InsertarContrato(CodCon As String, CodEmp As String, fila As Integer)
    CodCon = "CN" & Right("0000" & val(Right(CodCon, 2)) + 1, 2)
    With flexContratos
        SQL = " Call Insert_Cont ('" & CodCon & "', '" & Trim(.TextMatrix(fila, 11)) & "' , " & _
              "'" & Right("00" & Trim(.TextMatrix(fila, 1)), 2) & "','" & Trim(CodEmp) & "', " & _
              "'" & Format(.TextMatrix(fila, 7), "yyyy/mm/dd") & "', '" & Format(.TextMatrix(fila, 8), "yyyy/mm/dd") & "'," & _
              " " & CDbl(.TextMatrix(fila, 5)) & " ," & CDbl(.TextMatrix(fila, 6)) & ", " & _
              "'" & Trim(.TextMatrix(fila, 3)) & "' , " & CDbl(.TextMatrix(fila, 4)) & ", 'CA', " & _
              "'" & Trim(.TextMatrix(fila, 12)) & "','" & Trim(.TextMatrix(fila, 13)) & "', " & _
              "'" & Trim(.TextMatrix(fila, 14)) & "','" & Trim(.TextMatrix(fila, 15)) & "', " & _
              "'" & Trim(.TextMatrix(fila, 16)) & "','" & Trim(.TextMatrix(fila, 17)) & "', " & _
              "'" & Trim(.TextMatrix(fila, 18)) & "','" & Trim(.TextMatrix(fila, 19)) & "', " & _
              "'" & Trim(.TextMatrix(fila, 20)) & "');"
        oConexionMYSQL.Execute SQL
    End With
End Sub
