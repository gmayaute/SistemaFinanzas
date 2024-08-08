VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Facturación"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4800
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   4845
      Left            =   0
      TabIndex        =   1
      Top             =   -30
      Width           =   4785
      Begin NOVAdmin.flxEdit flxOpciones 
         Height          =   2265
         Left            =   60
         TabIndex        =   7
         Top             =   150
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   3995
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
         BackColorSel    =   -2147483643
         BackColorFixed  =   9868950
         CellPicture     =   "frmOpciones.frx":0442
         ColAlignment0   =   9
         FixedAlignment0 =   9
         ForeColorSel    =   16711680
         ForeColorFixed  =   14474460
         MouseIcon       =   "frmOpciones.frx":045E
         RowHeight0      =   240
      End
      Begin Proyecto1.chameleonButton btnGrabar 
         Height          =   345
         Left            =   3540
         TabIndex        =   2
         ToolTipText     =   "Guardar"
         Top             =   4350
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
         MICON           =   "frmOpciones.frx":047A
         PICN            =   "frmOpciones.frx":0496
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnSalir 
         Height          =   345
         Left            =   4170
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   4350
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
         MICON           =   "frmOpciones.frx":08D8
         PICN            =   "frmOpciones.frx":08F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.ComboBox cboMoneda 
         Height          =   315
         Left            =   1050
         TabIndex        =   6
         Top             =   2460
         Width           =   3585
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "6324;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda:"
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
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H009F5539&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa de Detracción:"
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
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   2820
         Width           =   4575
      End
      Begin MSForms.TextBox txtGlosa 
         Height          =   1035
         Left            =   90
         TabIndex        =   0
         Top             =   3210
         Width           =   4575
         VariousPropertyBits=   -1400879077
         Size            =   "8070;1826"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean

Private Sub InicializaGrid()
    With flxOpciones
        .Clear
        .Cols = 5
        .Rows = 1
        
        .row = 0
        .Col = 0
        .CellForeColor = &H80000002
        .ColWidth(0) = 2000
        .ColType(0) = cadena
        .ColMaxLength(0) = 30
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.IZQUIERDA
        .TextMatrix(0, 0) = Space(14) + "Documento"
        
        .row = 0
        .Col = 1
        .CellForeColor = &H80000002
        .ColWidth(1) = 800
        .ColType(1) = cadena
        .ColMaxLength(0) = 30
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.IZQUIERDA
        .CaracteresValidos(1) = "0123456789"
        .TextMatrix(0, 1) = Space(2) + "Serie"
        
        .row = 0
        .Col = 2
        .CellForeColor = &H80000002
        .ColWidth(2) = 1500
        .ColType(2) = cadena
        .ColMaxLength(0) = 30
        .ColAlignment(0) = MSHFLEXGRID_ALINEACION.IZQUIERDA
        .CaracteresValidos(2) = "0123456789"
        .TextMatrix(0, 2) = Space(6) & "Correlativo"
        
        .row = 0
        .Col = 3
        .CellForeColor = &H80000002
        .ColWidth(3) = 0
        .ColType(3) = cadena
        .ColMaxLength(0) = 30
        .TextMatrix(0, 3) = Space(14) + "Codigo"
        
        .row = 0
        .Col = 4
        .CellForeColor = &H80000002
        .ColWidth(4) = 0
        .ColType(4) = cadena
        .ColMaxLength(0) = 30
        .TextMatrix(0, 4) = Space(14) + "Moneda"
        
    End With
End Sub

Private Sub btnGrabar_Click()
    ValidarCorrel
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub cboMoneda_Change()
    If cboMoneda.ListIndex = 1 Then
        intTipoMoneda = 1
    End If
    If cboMoneda.ListIndex = 2 Then
       intTipoMoneda = 2
    End If
    If cboMoneda.ListIndex = 0 Then
        intTipoMoneda = 0
    End If
End Sub

Private Sub moneda(cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem "Seleccionar..."
    cbo.List(0, 1) = "0"
    cbo.AddItem "Nacional"
    cbo.List(1, 1) = "N"
    cbo.AddItem "Extranjera"
    cbo.List(2, 1) = "E"
    cbo.ListIndex = intTipoMoneda
End Sub

Private Sub flxOpciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With flxOpciones
            If .Col = 1 Then
                .TextMatrix(.row, .Col) = Right("00000" & .TextMatrix(.row, .Col), 5)
            End If
            If .Col = 2 Then
                .TextMatrix(.row, .Col) = Right("00000000" & .TextMatrix(.row, .Col), 9)
            End If
        End With
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    InicializaGrid
    moneda cboMoneda
    ConfigObjet
    LlenarGrid
    Publimensaje = "modificar"
End Sub

Private Sub ConfigObjet()
    Select Case ModActivo
        Case 3
            frmOpciones.Height = 5130
            frmOpciones.Caption = "Parámetros Facturación"
            btnGrabar.Top = 4320
            btnSalir.Top = 4320
            cboMoneda.Visible = True
            txtGlosa.Visible = True
            Label18.Visible = True
            Label1.Visible = True
        Case 5
            frmOpciones.Caption = "Correlativo del Cheque"
            frmOpciones.Height = 3255
            btnGrabar.Top = 2460
            btnSalir.Top = 2460
            cboMoneda.Visible = False
            txtGlosa.Visible = False
            Label18.Visible = False
            Label1.Visible = False
            
    End Select
End Sub

Private Sub LlenarGrid()
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim I As Integer
    Select Case ModActivo
        Case 3
            SQL = "opcfact where codigo = '01' or codigo = '03' or codigo = '07' or codigo = '08' or codigo = 'P'"
        Case 5
            SQL = "opcfact where codigo = '1'"
    End Select
    Set Rs = oConexion.EjecutaSelect(SQL)
    I = 0
    With flxOpciones
        Do While Not Rs.EOF
            .Rows = .Rows + 1
            I = I + 1
            
            .TextMatrix(I, 0) = CE(Rs.Fields("DOCUMENTO"))
            .ColType(0) = cadena
            .ColMaxLength(0) = 25
             TipodeCampo = cadena
             
            .TextMatrix(I, 1) = CE(Rs.Fields("SERIE"))
            .ColType(1) = cadena
            .ColMaxLength(1) = 5
            TipodeCampo = cadena
            
            .TextMatrix(I, 2) = CE(Rs.Fields("CORRELATIVO"))
            .ColType(2) = cadena
            .ColMaxLength(2) = 9
            TipodeCampo = cadena
            
            .TextMatrix(I, 3) = CE(Rs.Fields("CODIGO"))
            .ColType(3) = cadena
            .ColMaxLength(3) = 3
            .Col = 3
            .CellForeColor = ColorDeshabilitado
             TipodeCampo = cadena
            
            .TextMatrix(I, 4) = CE(Rs.Fields("MONEDA"))
            .ColType(4) = cadena
            .ColMaxLength(4) = 1
            .Col = 4
            .CellForeColor = ColorDeshabilitado
             TipodeCampo = cadena
            
            txtGlosa = CE(Rs.Fields("GLOSA"))
            Rs.MoveNext
        Loop
    End With
    Set Rs = Nothing
End Sub

Private Sub Grabar(codigo As String, serie As String, correl As String, moneda As String)
    Dim SQL As String
    Dim I As Integer
    If txtGlosa <> Empty And cboMoneda.ListIndex <> 0 Then
        Select Case ModActivo
            Case 3
                SQL = "Call Update_OpcFact ('" & codigo & "', '" & serie & "', '" & correl & "'," & _
                       " '" & txtGlosa & "','" & cboMoneda.List(cboMoneda.ListIndex, 1) & "' ) ;"
            Case 5
                SQL = "Call Update_OpcFact ('" & codigo & "', '" & serie & "', '" & correl & "'," & _
                       " '" & txtGlosa & "','" & moneda & "' ) ;"
        End Select
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    Else
        txtGlosa.SetFocus
    End If
End Sub

Private Sub ValidarCorrel()
    Dim I As Integer
    Dim RES As Integer
    Dim CorrelSgt As String
    bandera = False
    With flxOpciones
        For I = 1 To .Rows - 1
            CorrelSgt = GenerarCorrel(.TextMatrix(I, 3), .TextMatrix(I, 4))
            Grabar .TextMatrix(I, 3), .TextMatrix(I, 1), .TextMatrix(I, 2), .TextMatrix(I, 4)
            
            'If CDbl(.TextMatrix(I, 2)) = CDbl(CorrelSgt) Then Grabar .TextMatrix(I, 3), .TextMatrix(I, 1), .TextMatrix(I, 2), .TextMatrix(I, 4)
            'If CDbl(.TextMatrix(I, 2)) > CDbl(CorrelSgt) Then
            '    RES = MsgBox("Está quedando" & str(CDbl(.TextMatrix(I, 2)) - CDbl(CorrelSgt)) & " correlativo(s) en blanco" & vbNewLine & vbNewLine & "¿Desea Continuar?", vbYesNo + vbInformation, gsNomSW)
            '    If RES = vbYes Then Grabar .TextMatrix(I, 3), .TextMatrix(I, 1), .TextMatrix(I, 2), .TextMatrix(I, 4): CorrelSgt = "0"
            '    If RES = vbNo Then .row = I: .Col = 2: Call keybd_event(vbKeyF2, 0, 0, 0): Exit For
            'End If
            'If CDbl(.TextMatrix(I, 2)) < CDbl(CorrelSgt) Then MsgBox "Correlativo no valido para documento " & .TextMatrix(I, 0), vbOKOnly + vbExclamation, gsNomSW: .row = I: .Col = 2: Call keybd_event(vbKeyF2, 0, 0, 0): Exit For
            
        Next
    End With
    
    MsgBox "Los Correlativos del Cheque han sido modificados, por favor verifique su estado, BAJO SU RESPONSABILIDAD."
End Sub

Private Function GenerarCorrel(TipoDoc As String, Optional moneda As String) As String
    Dim SQL As String
    Dim rscorrel As MYSQL_RS
    Select Case ModActivo
        Case 3
            SQL = " Select Max(correl) as doc from (documento_contables as a left join amarre_documento as b" & _
                   " on a.identificador = b.identificador) LEFT JOIN movi_documento AS C  ON C.Identificador = b.Identificador" & _
                   " where b.Cod_Tipo_Doc = '" & TipoDoc & "' and flag = '0' AND C.Cod_Estado<>'EL'"
        Case 5
            SQL = "Select Max(REF) as doc from documento_contables as a  left join amarre_documento as b " & _
                  " on a.identificador = b.identificador where b.Cod_Tipo_Doc = '" & TipoDoc & "'" & _
                  " and flag = '0' and a.mon = '" & moneda & "'"
    End Select
    
    Set rscorrel = oConexion.EjecutaSelectRS(SQL)
    
    If Not IsNull(rscorrel.Fields("doc")) Then
        GenerarCorrel = Right("000000000" & (CDbl(rscorrel.Fields("doc")) + 1), 9)
    End If
    If IsNull(rscorrel.Fields("doc")) Then
        GenerarCorrel = "000000001"
    End If
    Set rscorrel = Nothing
End Function
