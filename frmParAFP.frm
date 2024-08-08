VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmParAFP 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros de AFP"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12120
   Icon            =   "frmParAFP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin NOVAdmin.flxEdit flxTasas 
      Height          =   2535
      Left            =   30
      TabIndex        =   3
      Top             =   540
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   4471
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CellFontName    =   "MS Sans Serif"
      CellFontSize    =   8.25
      BackColorSel    =   -2147483635
      BackColorFixed  =   -2147483633
      CellPicture     =   "frmParAFP.frx":014A
      ConfirmarBorradoLinea=   0   'False
      ColWidth0       =   960
      ColAlignment0   =   9
      FixedAlignment0 =   9
      ColWidth1       =   960
      ColAlignment1   =   9
      FixedAlignment1 =   9
      ForeColorSel    =   -2147483634
      ForeColorFixed  =   -2147483630
      GridColorFixed  =   12632256
      MouseIcon       =   "frmParAFP.frx":0166
      RowHeight0      =   240
      RowHeight1      =   240
   End
   Begin Proyecto1.chameleonButton btnInsertar 
      Height          =   345
      Left            =   8730
      TabIndex        =   2
      ToolTipText     =   "Copiar del mes anterior"
      Top             =   60
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
      MICON           =   "frmParAFP.frx":0182
      PICN            =   "frmParAFP.frx":019E
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
      BackStyle       =   0  'Transparent
      Caption         =   "MES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   645
   End
   Begin MSForms.ComboBox cboMes 
      Height          =   315
      Left            =   750
      TabIndex        =   0
      Top             =   90
      Width           =   2145
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "3784;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmParAFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Public Sub ConfiguraGrilla()
    Dim I As Integer
    With flxTasas
        .Clear
        .Rows = 1
        .Cols = 10
        .ForeColorFixed = &H404000
        .ColWidth(0) = 500
        .TextMatrix(0, 0) = Space(1) + "Item"
        .FixedCols = 1
        
        .ColWidth(1) = 2500
        .TextMatrix(0, 1) = Space(20) + "AFP"
        
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = "Cod"

        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "%  Obli."
        .ColType(3) = 3
        .ColMaxLength(3) = 6
        .CaracteresValidos(3) = "0123456789."
        .ColDecimales(3) = 2
        
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "% Com. Flujo"
        .ColType(4) = 3
        .ColMaxLength(4) = 6
        .CaracteresValidos(4) = "0123456789."
        .ColDecimales(4) = 2
        
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = "% Com. Mixta"
        .ColType(5) = 3
        .ColMaxLength(5) = 6
        .CaracteresValidos(5) = "0123456789."
        .ColDecimales(5) = 2

              
        .ColWidth(6) = 1000
        .TextMatrix(0, 6) = "%   Seg."
        .ColType(6) = 3
        .ColMaxLength(6) = 6
        .CaracteresValidos(6) = "0123456789."
        .ColDecimales(6) = 2
        
        .ColWidth(7) = 1000
        .TextMatrix(0, 7) = " Tope Oblig."
        .ColType(7) = 3
        .ColMaxLength(7) = 11
        .CaracteresValidos(7) = "0123456789."
        .ColDecimales(7) = 2
              
        .ColWidth(8) = 1000
        .TextMatrix(0, 8) = " Tope Com."
        .ColType(8) = 3
        .ColMaxLength(8) = 11
        .CaracteresValidos(8) = "0123456789."
        .ColDecimales(8) = 2
        
        .ColWidth(9) = 1200
        .TextMatrix(0, 9) = " Tope Seg."
        .ColType(9) = 3
        .ColMaxLength(9) = 11
        .CaracteresValidos(9) = "0123456789."
        .ColDecimales(9) = 2
        
    End With
End Sub
Sub CargaTasas(AnoMes As String)
    Dim rsTasas As MYSQL_RS, I As Integer
    ConfiguraGrilla
    I = 1
    SQL = " Select a.*,b.*" & _
          " from afp as a inner join `pl_afppar` as b " & _
          " on (a.codigo=b.codafp)" & _
          " where a.dcto>0 and b.anomes='" & AnoMes & "'" & _
          " order by a.codigo"
    Set rsTasas = oConexion.EjecutaSelectRS(SQL)
    Do While Not (rsTasas.EOF)
       With flxTasas
            flxTasas.Rows = flxTasas.Rows + 1
            .TextMatrix(I, 0) = I
            .TextMatrix(I, 1) = rsTasas.Fields("Nombre")
            .TextMatrix(I, 2) = rsTasas.Fields("Codigo")
            .TextMatrix(I, 3) = Space(10) & FormatNumber(rsTasas.Fields("PorcDscto"), 2)
            .TextMatrix(I, 4) = Space(10) & FormatNumber(rsTasas.Fields("PorcCom"), 2)
            .TextMatrix(I, 5) = Space(10) & FormatNumber(rsTasas.Fields("PorcMix"), 2)
            .TextMatrix(I, 6) = Space(10) & FormatNumber(rsTasas.Fields("PorcSeg"), 2)
            .TextMatrix(I, 7) = Space(10) & FormatNumber(rsTasas.Fields("TopDscto"), 2)
            .TextMatrix(I, 8) = Space(10) & FormatNumber(rsTasas.Fields("TopCom"), 2)
            .TextMatrix(I, 9) = Space(10) & FormatNumber(rsTasas.Fields("TopSeg"), 2)
            I = I + 1
            rsTasas.MoveNext
       End With
    Loop
    btnInsertar.Enabled = False
    rsTasas.CloseRecordset
    If I = 1 Then
        SQL = "Select * from afp where dcto>0 order by codigo"
        Set rsTasas = oConexion.EjecutaSelectRS(SQL)
        Do While Not (rsTasas.EOF)
           With flxTasas
                flxTasas.Rows = flxTasas.Rows + 1
                .TextMatrix(I, 0) = I
                .TextMatrix(I, 1) = rsTasas.Fields("Nombre")
                .TextMatrix(I, 2) = rsTasas.Fields("Codigo")
                .TextMatrix(I, 3) = "0.00"
                .TextMatrix(I, 4) = "0.00"
                .TextMatrix(I, 5) = "0.00"
                .TextMatrix(I, 6) = "0.00"
                .TextMatrix(I, 7) = "0.00"
                .TextMatrix(I, 8) = "0.00"
                .TextMatrix(I, 9) = "0.00"
                I = I + 1
                rsTasas.MoveNext
           End With
        Loop
        btnInsertar.Enabled = True
    End If
    Set rsTasas = Nothing
End Sub

Private Sub btnInsertar_Click()
    Dim anomesant As String, v As Boolean
    Dim rspar As MYSQL_RS
    v = False
    If Trim(cboMes.List(cboMes.ListIndex, 2)) = "01" Then
        anomesant = Trim(CStr(val(strAnoSistema) - 1)) & "12"
    Else
        anomesant = strAnoSistema & Right("00" & Trim(CStr(val(cboMes.List(cboMes.ListIndex, 2)) - 1)), 2)
    End If
    SQL = "Select * from  pl_afppar where anomes='" & anomesant & "'"
    Set rspar = oConexion.EjecutaSelectRS(SQL)
    Do While Not (rspar.EOF)
        v = True
        SQL = "Insert into pl_afppar (anomes,codafp,porcdscto,porccom,porcmix,porcseg,topdscto,topcom,topseg) values (" & _
            "'" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "','" & rspar.Fields(1) & "'," & _
             rspar.Fields(2) & "," & rspar.Fields(3) & "," & rspar.Fields(4) & "," & rspar.Fields(5) & "," & rspar.Fields(6) & "," & rspar.Fields(7) & "," & rspar.Fields(8) & ")"
        oConexionMYSQL.Execute SQL
        rspar.MoveNext
    Loop
    cboMes_Change
    Set rspar = Nothing
    If v = True Then
        Exit Sub
    Else
        MsgBox "No hay parametros en el mes anterior", vbOKOnly + vbInformation, "NOVPeru"
    End If
    
End Sub

Private Sub cboMes_Change()
    CargaTasas strAnoSistema & cboMes.List(cboMes.ListIndex, 2)
End Sub

Private Sub flxTasas_KeyPress(KeyAscii As Integer)
        Dim campo As String
    If flxTasas.Col > 2 Then
        If KeyAscii = 13 Then
            Select Case flxTasas.Col
                Case 3: campo = "porcdscto"
                Case 4: campo = "porccom"
                Case 5: campo = "porcmix"
                Case 6: campo = "porcseg"
                Case 7: campo = "topdscto"
                Case 8: campo = "topcom"
                Case 9: campo = "topseg"
                
            End Select
            flxTasas.AsignarCelda
            SQL = "Update pl_afppar set  " & campo & "='" & CDbl(Trim(flxTasas.TextMatrix(flxTasas.row, flxTasas.Col))) & _
            "' where anomes='" & strAnoSistema & cboMes.List(cboMes.ListIndex, 2) & "' and codafp='" & flxTasas.TextMatrix(flxTasas.row, 2) & "'"
            oConexionMYSQL.Execute SQL
            cboMes_Change
        End If
    End If

End Sub

Private Sub flxTasas_RowColChange()
    If flxTasas.Col > 1 Then
        Publimensaje = "modificar"
        TipodeCampo = Numero
        
    Else
        Publimensaje = "sin-editar"
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LlenarMesP cboMes
    ConfiguraGrilla
    cboMes_Change
End Sub



