VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmParametrosGrles 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Parámetros Generales"
   ClientHeight    =   2895
   ClientLeft      =   3945
   ClientTop       =   6075
   ClientWidth     =   7590
   Icon            =   "frmParametrosGrles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7590
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   3015
      Left            =   0
      TabIndex        =   4
      Top             =   -90
      Width           =   7695
      Begin VB.Frame Frame5 
         BackColor       =   &H009F5539&
         Height          =   2205
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   7485
         Begin VB.TextBox txtRutaTw 
            Height          =   315
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1440
            Width           =   5175
         End
         Begin Proyecto1.chameleonButton btnExaminar 
            Height          =   375
            Left            =   6870
            TabIndex        =   3
            Top             =   1410
            Width           =   405
            _ExtentX        =   714
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
            MICON           =   "frmParametrosGrles.frx":0442
            PICN            =   "frmParametrosGrles.frx":045E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox meIgv 
            Height          =   315
            Left            =   1650
            TabIndex        =   0
            Top             =   270
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##.##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox meImpDet 
            Height          =   315
            Left            =   1650
            TabIndex        =   1
            Top             =   660
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##.##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox meImpRet 
            Height          =   315
            Left            =   1650
            TabIndex        =   2
            Top             =   1050
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##.##"
            PromptChar      =   " "
         End
         Begin MSForms.ComboBox cboSedes 
            Height          =   315
            Left            =   1650
            TabIndex        =   20
            Top             =   1830
            Width           =   5205
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "9181;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label9 
            BackColor       =   &H009F5539&
            Caption         =   "Sede de trabajo:"
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
            Height          =   195
            Left            =   90
            TabIndex        =   19
            Top             =   1890
            Width           =   1485
         End
         Begin MSForms.ComboBox cmbAnioSistema 
            Height          =   315
            Left            =   6240
            TabIndex        =   18
            Top             =   240
            Width           =   1185
            VariousPropertyBits=   746604571
            DisplayStyle    =   7
            Size            =   "2090;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            Caption         =   "Año de Sistema:"
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
            Height          =   195
            Left            =   4800
            TabIndex        =   17
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            Caption         =   "Imp. Retención:"
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
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   1110
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   195
            Left            =   2310
            TabIndex        =   12
            Top             =   1110
            Width           =   150
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            Caption         =   "Imp. Detracción:"
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
            Height          =   195
            Left            =   90
            TabIndex        =   11
            Top             =   720
            Width           =   1425
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   195
            Left            =   2310
            TabIndex        =   10
            Top             =   720
            Width           =   150
         End
         Begin VB.Label LblLibVentas 
            AutoSize        =   -1  'True
            BackColor       =   &H009F5539&
            Caption         =   "IGV:"
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
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Height          =   195
            Left            =   2310
            TabIndex        =   8
            Top             =   330
            Width           =   150
         End
         Begin VB.Label Label3 
            BackColor       =   &H009F5539&
            Caption         =   "Ruta Telewiese:"
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
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   1500
            Width           =   1395
         End
      End
      Begin Proyecto1.chameleonButton cmdSalir 
         Height          =   405
         Left            =   7050
         TabIndex        =   14
         Top             =   2430
         Width           =   465
         _ExtentX        =   767
         _ExtentY        =   714
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
         MICON           =   "frmParametrosGrles.frx":27E0
         PICN            =   "frmParametrosGrles.frx":27FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdGrabar 
         Height          =   405
         Left            =   3900
         TabIndex        =   15
         Top             =   2460
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
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
         MICON           =   "frmParametrosGrles.frx":2BC2
         PICN            =   "frmParametrosGrles.frx":2BDE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdCancelar 
         Height          =   405
         Left            =   3330
         TabIndex        =   16
         Top             =   2460
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
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
         MICON           =   "frmParametrosGrles.frx":3020
         PICN            =   "frmParametrosGrles.frx":303C
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2460
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   """Text (*.txt)|*.txt|All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "FrmParametrosGrles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExaminar_Click()
    CommonDialog1.Filter = ".txt(*.txt)"
    CommonDialog1.ShowOpen
    If CommonDialog1.Filename <> "" Then
        txtRutaTw = CommonDialog1.Filename
        txtRutaTw = Replace(txtRutaTw, CommonDialog1.FileTitle, Empty)
    End If
End Sub

Private Function ValidaData() As Boolean
    ValidaData = True
    If meIgv.Text = Empty Then MsgBox "El campo IGV no puede quedar vacío", vbInformation, gsNomSW: ValidaData = False: meIgv.SetFocus: Exit Function
    If meImpDet.Text = Empty Then MsgBox "El campo de Detracción no puede quedar vacío", vbInformation, gsNomSW: ValidaData = False: meImpDet.SetFocus: Exit Function
    If meImpRet.Text = Empty Then MsgBox "El campo de Retención no puede quedar vacío", vbInformation, gsNomSW: ValidaData = False: meImpRet.SetFocus: Exit Function
    If txtRutaTw = Empty Then MsgBox "La ruta del Telewiese no puede quedar vacía", vbInformation, gsNomSW: ValidaData = False: txtRutaTw.SetFocus: Exit Function
End Function

Private Sub Grabar()
    Dim SQL As String
    If ValidaData Then
        SQL = " Call Update_Param ( '" & strAnoSistema & "'," & _
              " " & CDbl(CEN(meIgv)) & ", " & CDbl(CEN(meImpRet)) & "," & _
              " " & CDbl(CEN(meImpDet)) & ", '" & CE(Replace(txtRutaTw, "\", "*")) & "','" & cboSedes.List(cboSedes.ListIndex, 1) & "');"
        oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
    End If
End Sub

Private Sub cmbAnioSistema_Change()
    strAnoSistema = Trim(cmbAnioSistema.Text)
    mdiInicio.sbPrincipal.Panels(4).Text = Trim(cmbAnioSistema.Text)
    CargarDatos
End Sub

Private Sub cmdCancelar_Click()
    CargarDatos
    meIgv.SetFocus
End Sub

Private Sub cmdGrabar_Click()
    Grabar
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Dim SQL As String, sede As String, I As Integer
    SQL = "Select anio,sede from cnparam order by anio"
    Set rsdato = oConexion.EjecutaSelectRS(SQL)
    rsdato.MoveFirst
    I = 0
    Do While Not rsdato.EOF
        cmbAnioSistema.AddItem rsdato.Fields("ANIO")
        cmbAnioSistema.List(I, 1) = rsdato.Fields("sede")
        I = I + 1
        rsdato.MoveNext
    Loop
    For I = 0 To cmbAnioSistema.ListCount - 1
        If cmbAnioSistema.List(I, 0) = strAnoSistema Then
            cmbAnioSistema.ListIndex = I
            sede = cmbAnioSistema.List(I, 1)
        End If
    Next
    SQL = "Select nombre,codigo from rh_estacionestrabajo order by nombre"
    Set rsdato = oConexion.EjecutaSelectRS(SQL)
    rsdato.MoveFirst
    I = 0
    Do While Not rsdato.EOF
        cboSedes.AddItem rsdato.Fields("nombre")
        cboSedes.List(I, 1) = rsdato.Fields("codigo")
        I = I + 1
        rsdato.MoveNext
    Loop
    For I = 0 To cboSedes.ListCount - 1
        If cboSedes.List(I, 1) = sede Then
            cboSedes.ListIndex = I
        End If
    Next
    Set rsdato = Nothing
    CargarDatos
End Sub

Private Sub meIgv_GotFocus()
    mark1 meIgv
End Sub

Private Sub meIgv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        meImpDet.SetFocus
    End If
End Sub

Private Sub meImpDet_GotFocus()
    mark1 meImpDet
End Sub

Private Sub meImpDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        meImpRet.SetFocus
    End If
End Sub

Private Sub meImpRet_GotFocus()
    mark1 meImpRet
End Sub

Private Sub meImpRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        btnExaminar.SetFocus
    End If
End Sub

Private Sub CargarDatos()
On Error GoTo problema
    Dim SQL As String
    Dim rsdato As MYSQL_RS
    SQL = "Select * from cnparam where anio = '" & strAnoSistema & "'"
    Set rsdato = oConexion.EjecutaSelectRS(SQL)
    If Not rsdato.EOF Then
        txtRutaTw = CE(Replace(rsdato.Fields("rutatw"), "*", "\"))
        meIgv.Text = Trim(CE(FormatNumber(rsdato.Fields("igv"), 2)))
        meImpDet.Text = CE(FormatNumber(rsdato.Fields("detraccion"), 2))
        If CE(FormatNumber(rsdato.Fields("retencion"), 2)) > 0 Then
            meImpRet.Text = CE(FormatNumber(rsdato.Fields("retencion"), 2))
        End If
    End If
    Set rsdato = Nothing
Exit Sub
problema:
    MsgBox "Los Datos para este año no se encuntran completos" & vbNewLine & "Por favor completelos y grabe los cambios", vbOKOnly + vbInformation, gsNomSW
    Exit Sub
End Sub

Private Sub txtRutaTw_GotFocus()
    mark1 txtRutaTw
End Sub
