VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmMaestro 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento - Maestros"
   ClientHeight    =   6600
   ClientLeft      =   10440
   ClientTop       =   3990
   ClientWidth     =   10560
   Icon            =   "frmMaestro.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10560
   Begin Proyecto1.chameleonButton cmdSalir 
      Height          =   405
      Left            =   7440
      TabIndex        =   31
      Top             =   6000
      Width           =   435
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
      MICON           =   "frmMaestro.frx":0442
      PICN            =   "frmMaestro.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdVistaPreliminar 
      Height          =   405
      Left            =   6960
      TabIndex        =   30
      Top             =   6000
      Width           =   435
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
      MICON           =   "frmMaestro.frx":0824
      PICN            =   "frmMaestro.frx":0840
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
      Left            =   4380
      TabIndex        =   29
      Top             =   5970
      Width           =   435
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
      MICON           =   "frmMaestro.frx":0D82
      PICN            =   "frmMaestro.frx":0D9E
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
      Left            =   3900
      TabIndex        =   28
      Top             =   5970
      Width           =   435
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
      MICON           =   "frmMaestro.frx":11E0
      PICN            =   "frmMaestro.frx":11FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdEliminar 
      Height          =   405
      Left            =   2460
      TabIndex        =   27
      Top             =   5955
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Eliminar"
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
      MICON           =   "frmMaestro.frx":173E
      PICN            =   "frmMaestro.frx":175A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton CmdModificar 
      Height          =   405
      Left            =   1140
      TabIndex        =   26
      Top             =   5970
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Modificar"
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
      MICON           =   "frmMaestro.frx":1B9C
      PICN            =   "frmMaestro.frx":1BB8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton cmdNuevo 
      Height          =   405
      Left            =   30
      TabIndex        =   25
      Top             =   5970
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "&Nuevo"
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
      MICON           =   "frmMaestro.frx":1FE6
      PICN            =   "frmMaestro.frx":2002
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtObs 
      Height          =   1380
      Left            =   7890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.ComboBox cboActa 
      Height          =   315
      ItemData        =   "frmMaestro.frx":236C
      Left            =   1050
      List            =   "frmMaestro.frx":236E
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   1590
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   10470
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Index           =   4
         Left            =   1170
         TabIndex        =   34
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118030337
         CurrentDate     =   38464
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   1
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         Top             =   570
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   7
         Left            =   6780
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1215
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   6
         Left            =   6780
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   870
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   5
         Left            =   6780
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   10
         Top             =   525
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   4
         Left            =   6780
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Top             =   200
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   3
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1200
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   2
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   870
         Width           =   3690
      End
      Begin VB.TextBox T 
         Height          =   285
         Index           =   0
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   3690
      End
      Begin MSForms.Label lblL 
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   33
         Top             =   570
         Width           =   1095
         ForeColor       =   16777215
         BackColor       =   10442041
         Size            =   "1931;503"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   195
         Index           =   7
         Left            =   2610
         TabIndex        =   22
         Top             =   1230
         Width           =   3450
         ForeColor       =   8421631
         BackColor       =   10442041
         VariousPropertyBits=   276824091
         Caption         =   "Division"
         Size            =   "6085;344"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   225
         Index           =   5
         Left            =   2520
         TabIndex        =   20
         Top             =   570
         Width           =   2775
         ForeColor       =   8421631
         BackColor       =   10442041
         Size            =   "4895;397"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   255
         Index           =   4
         Left            =   5790
         TabIndex        =   19
         Top             =   210
         Width           =   885
         ForeColor       =   16777215
         BackColor       =   10442041
         Size            =   "1561;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   18
         Top             =   1230
         Width           =   1065
         ForeColor       =   16777215
         BackColor       =   10442041
         Size            =   "1879;503"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   17
         Top             =   900
         Width           =   1095
         ForeColor       =   16777215
         BackColor       =   10442041
         Size            =   "1931;503"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Width           =   1095
         ForeColor       =   16777215
         BackColor       =   10442041
         Size            =   "1931;503"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblL 
         Height          =   195
         Index           =   6
         Left            =   2730
         TabIndex        =   21
         Top             =   930
         Width           =   3390
         ForeColor       =   8421631
         BackColor       =   10442041
         VariousPropertyBits=   276824091
         Caption         =   "Prueba dsfadsfasdfasdfadsfsadfsadfdsa"
         Size            =   "5980;344"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame FrameP 
      BackColor       =   &H009F5539&
      BorderStyle     =   0  'None
      Height          =   4260
      Left            =   0
      TabIndex        =   4
      Top             =   1740
      Width           =   7920
      Begin VB.ComboBox cboCampos 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3795
         Width           =   2220
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H009F5539&
         Height          =   600
         Left            =   0
         TabIndex        =   6
         Top             =   3585
         Width           =   7920
         Begin VB.TextBox TxtCriterio 
            Height          =   285
            Left            =   2400
            TabIndex        =   7
            Top             =   210
            Width           =   5370
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H009F5539&
         Height          =   3315
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   7860
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMaestros 
            Height          =   3090
            Left            =   75
            TabIndex        =   0
            Top             =   150
            Width           =   7710
            _ExtentX        =   13600
            _ExtentY        =   5450
            _Version        =   393216
            FixedCols       =   0
            BackColorBkg    =   12632256
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSForms.Label lblMensaje 
         Height          =   225
         Left            =   5220
         TabIndex        =   35
         Top             =   2580
         Width           =   1515
         ForeColor       =   65280
         BackColor       =   10442041
         Size            =   "2672;397"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label lblMensaje2 
         Height          =   225
         Left            =   60
         TabIndex        =   24
         Top             =   3390
         Width           =   4995
         ForeColor       =   8421631
         BackColor       =   10442041
         Caption         =   "F2  Código   - F3  Descripción  -  F4 Captura Fecha/Hora"
         Size            =   "8811;397"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin MSForms.Label lblL 
      Height          =   315
      Index           =   8
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   885
      ForeColor       =   16777215
      BackColor       =   10442041
      Size            =   "1561;556"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmMaestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MshHabilitado As Boolean
Public CantRegistros As String
Private m_QueryRep As String
Private m_Reporte As String
Private m_Titulo As String
Private m_Nombre As String 'titulo principal del from
Private m_Cols As Integer
Private m_ColsHabilitada As Integer
Private m_Col As Integer
Private m_NomCol(10) As String
Private m_AnchoCol(10) As Integer
Private m_tipo As TIPO_FORMULARIO
Dim FilSel As Integer
Dim FilEli As Integer

Public Property Let pTipo(valor As Integer)
    m_tipo = valor
End Property

Public Property Let pQueryRep(valor As String)
    m_QueryRep = valor
End Property

Public Property Let pReporte(valor As String)
    m_Reporte = valor
End Property

Public Property Let pTitulo(valor As String)
    m_Titulo = valor
End Property

Public Property Let pNombre(valor As String)
    m_Nombre = valor
End Property

Public Property Let pCols(valor As Integer)
    m_Cols = valor
End Property
Public Property Let pColsHabilitada(valor As Integer)
    m_ColsHabilitada = valor
End Property
Public Property Let pCol(valor As Integer)
    m_Col = valor
End Property
Public Property Let pNomCol(valor As String)
    m_NomCol(m_Col) = valor
End Property
Public Property Let pAnchoCol(valor As Integer)
    m_AnchoCol(m_Col) = valor
End Property
Sub BloqueoEspecial()
    MshHabilitado = False
    mshMaestros.BackColor = ColorDeshabilitado
    T(0).Locked = True
    T(0).BackColor = ColorDeshabilitado
End Sub
Sub BotonEdicion()
    cmdNuevo.Enabled = False
    CmdModificar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    CmdEliminar.Enabled = False
    cmdSalir.Enabled = False
End Sub
Sub BotonNormal()
    Dim I As Integer
    cmdNuevo.Enabled = True
    CmdModificar.Enabled = True
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    CmdEliminar.Enabled = True
    cmdSalir.Enabled = True
    For I = 0 To m_Cols - m_ColsHabilitada - 1
        T(I).Enabled = False
    Next I
    txtObs.Locked = True
    txtObs.BackColor = ColorDeshabilitado
End Sub
Sub ConfigMshMaestros()
    Dim I As Integer
    With mshMaestros
        .Clear
        .Cols = m_Cols
        .Rows = 2
        cboCampos.Clear
        For I = 0 To m_Cols - 1
            .TextMatrix(0, I) = m_NomCol(I)
            .ColWidth(I) = m_AnchoCol(I)
            cboCampos.AddItem m_NomCol(I)
        Next I
        If cboCampos.ListCount > 0 Then cboCampos.ListIndex = 0
        .FixedCols = m_Cols - m_ColsHabilitada
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
    End With
End Sub
Sub BloqueoDeBotones()
    cmdNuevo.Enabled = True
    CmdModificar.Enabled = False
    CmdEliminar.Enabled = False
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    CmdVistaPreliminar.Enabled = False
End Sub
Sub ModoEdicion()
    Dim I As Integer
    Select Case m_tipo
           Case FORM_MAESTRO_COMPAÑIAS, FORM_MAESTRO_AREAS, FORM_MAESTRO_PERFILES, _
                FORM_MAESTROS_CENCO, FORM_MAESTRO_TIP_DOC, FORM_MAESTROS_FAM, _
                FORM_MAESTROS_CARGOS, FORM_MAESTRO_TARIFAS, FORM_MAESTRO_TIPOSDESC, _
                FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_UM, _
                FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, FORM_MAESTRO_BANCOS, _
                FORM_MAESTRO_AFP, FORM_MAESTRO_SEGUROS, FORM_MAESTRO_ALMACEN, _
                FORM_MAESTRO_CONTRATO, FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_LINEAS, _
                FORM_MAESTRO_AGENCIAS, FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_CUENTAS, _
                FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, FORM_MAESTRO_GRADOS, _
                FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                For I = 0 To m_Cols - m_ColsHabilitada - 1
                    T(I).Enabled = False
                    T(I).Locked = True
                    T(I).BackColor = ColorDeshabilitado
                Next I
                For I = m_Cols - m_ColsHabilitada To m_Cols - 1
                    T(I).Enabled = True
                    T(I).Locked = False
                    T(I).BackColor = ColorHabilitado
                Next I
           Case FORM_MAESTRO_ESTADO
                For I = 0 To m_Cols - m_ColsHabilitada - 1
                    T(I).Enabled = False
                    T(I).Locked = True
                    T(I).BackColor = ColorDeshabilitado
                Next I
                If T(0) <> ELIMINADO Or T(0) <> EMITIDO Or T(0) <> MODIFICADO Then
                    For I = m_Cols - m_ColsHabilitada To m_Cols - 1
                        T(I).Enabled = True
                        T(I).Locked = False
                        T(I).BackColor = ColorHabilitado
                    Next I
                Else
                    For I = m_Cols - m_ColsHabilitada To m_Cols - 1
                        T(I).Enabled = False
                        T(I).Locked = True
                        T(I).BackColor = ColorDeshabilitado
                    Next I
                End If
           Case FORM_MAESTRO_TIPO_CAMBIO
                For I = m_Cols - m_ColsHabilitada To m_Cols - 1
                    T(I).Enabled = True
                    T(I).Locked = False
                    T(I).BackColor = ColorHabilitado
                Next I
                dtpFecha(4).Enabled = False
    End Select
    txtObs.Locked = False
    txtObs.BackColor = ColorHabilitado
    MshHabilitado = False
    mshMaestros.BackColor = ColorDeshabilitado
    If T(m_Cols - m_ColsHabilitada).Enabled = True Then T(m_Cols - m_ColsHabilitada).SetFocus
End Sub

Sub ModoNormal()
    Dim I, J As Integer
    For I = 0 To m_Cols - 1
        T(I).Locked = True
        T(I).BackColor = ColorDeshabilitado
    Next I
    MshHabilitado = True
    mshMaestros.BackColor = ColorHabilitado
End Sub

Sub Ayuda(KeyCode As Integer)
    If KeyCode = vbKeyF2 Then
        Select Case m_tipo
               Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_COMPAÑIAS, FORM_MAESTRO_AREAS, _
                    FORM_MAESTRO_PERFILES, FORM_MAESTROS_CENCO, FORM_MAESTROS_FAM, _
                    FORM_MAESTROS_CARGOS, FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_UM, _
                    FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_TIPOSDESC, FORM_MAESTRO_AFP, _
                    FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, FORM_MAESTRO_SEGUROS, _
                    FORM_MAESTRO_ALMACEN, FORM_MAESTRO_BANCOS, FORM_MAESTRO_CONTRATO, _
                    FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_LINEAS, FORM_MAESTRO_AGENCIAS, _
                    FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, _
                    FORM_MAESTRO_GRADOS, FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                    TxtCriterio.SetFocus
                    cboCampos.Text = lblL(0)
        End Select
    End If
    If KeyCode = vbKeyF3 Then
        Select Case m_tipo
               Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_COMPAÑIAS, FORM_MAESTRO_AREAS, _
                    FORM_MAESTRO_PERFILES, FORM_MAESTROS_CENCO, FORM_MAESTROS_CARGOS, _
                    FORM_MAESTRO_UM, FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_TIPOSDESC, _
                    FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, FORM_MAESTRO_ALMACEN, _
                    FORM_MAESTRO_BANCOS, FORM_MAESTRO_CONTRATO, FORM_MAESTRO_LINEAS, _
                    FORM_MAESTRO_AGENCIAS, FORM_MAESTRO_ESTANCIAS
                    TxtCriterio.SetFocus
                    cboCampos.Text = lblL(1)
        End Select
    End If
End Sub

Private Sub cboActa_Click()
    ConfigMshMaestros
    LlenarMshMaestros
End Sub

Private Sub cboCampos_Click()
    mshMaestros.Col = cboCampos.ListIndex
    mshMaestros.Sort = flexSortStringAscending
    mshMaestros.Refresh
End Sub

Private Sub cmdCancelar_Click()
    ModoNormal
    BotonNormal
    Limpia_Valores
    LblMensaje = Empty
    DesplazarPorLaGrilla
End Sub

Private Sub Asigna_Valores()
    Dim I As Integer
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_TIP_DOC, FORM_MAESTRO_COMPAÑIAS, _
                FORM_MAESTRO_AREAS, FORM_MAESTRO_PERFILES, FORM_MAESTROS_CENCO, _
                FORM_MAESTROS_FAM, FORM_MAESTROS_CARGOS, FORM_MAESTRO_FORMPAGO, _
                FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_TIPOSDESC, FORM_MAESTRO_AFP, _
                FORM_MAESTRO_UM, FORM_MAESTRO_PARENT, FORM_MAESTRO_SEGUROS, _
                FORM_MAESTRO_ALMACEN, FORM_MAESTRO_BANCOS, FORM_MAESTRO_CONTRATO, _
                FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_LINEAS, FORM_MAESTRO_AGENCIAS, _
                FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, _
                FORM_MAESTRO_GRADOS, FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                For I = 0 To m_Cols - 1
                    mshMaestros.TextMatrix(mshMaestros.row, I) = Trim(T(I).Text)
                Next I
           Case FORM_MAESTRO_TIPO_CAMBIO
                For I = 1 To m_Cols - 1
                    mshMaestros.TextMatrix(mshMaestros.row, I) = Trim(T(I).Text)
                Next I
                mshMaestros.TextMatrix(mshMaestros.row, 0) = dtpFecha(4).Value
    End Select
End Sub

Private Sub Limpia_Valores()
    Dim I As Integer
    For I = 0 To m_Cols - 1
        T(I).Text = Empty
    Next I
    For I = m_Cols To 7
        T(I).Text = Empty
        lblL(I).Caption = Empty
    Next I
    txtObs = Empty
End Sub

Private Sub cmdEliminar_Click()
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_TIP_DOC, FORM_MAESTRO_COMPAÑIAS, _
                FORM_MAESTRO_AREAS, FORM_MAESTRO_PERFILES, FORM_MAESTROS_CENCO, _
                FORM_MAESTROS_FAM, FORM_MAESTROS_CARGOS, FORM_MAESTRO_TARIFAS, _
                FORM_MAESTRO_UM, FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_TIPOPAGO, _
                FORM_MAESTRO_TIPOSDESC, FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, _
                FORM_MAESTRO_AFP, FORM_MAESTRO_SEGUROS, FORM_MAESTRO_ALMACEN, _
                FORM_MAESTRO_BANCOS, FORM_MAESTRO_CONTRATO, FORM_MAESTRO_OPERACIONES, _
                FORM_MAESTRO_LINEAS, FORM_MAESTRO_AGENCIAS, FORM_MAESTRO_ESTANCIAS, _
                FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, FORM_MAESTRO_GRADOS, _
                FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                If T(0) <> Empty Then
                    If MsgBox("Está Seguro De Eliminar El Item Con El Código N° " + T(0) + "  (S/N)", vbInformation + vbYesNo, m_Titulo) = vbYes Then
                        FilEli = mshMaestros.row
                        BorrarRegistro
                        mshMaestros.Clear
                        ConfigMshMaestros
                        LlenarMshMaestros
                        Limpia_Valores
                        DesplazarPorLaGrilla
                        ModoNormal
                        BotonNormal
                        If FilEli > 1 Then
                            mshMaestros.row = FilEli - 1
                            mshMaestros.SetFocus
                            'SendKeys "{HOME}"
                            Call keybd_event(vbKeyHome, 0, 0, 0)
                        Else
                            mshMaestros.row = 1
                        End If
                        LblMensaje = Empty
                    End If
                Else
                    MsgBox "Seleccione El Registro A Eliminar", vbInformation, "STC Centro De Costo"
                    mshMaestros.SetFocus
                End If
           Case FORM_MAESTRO_TIPO_CAMBIO
                If dtpFecha(4).Value <> Empty Then
                    If MsgBox("Está Seguro De Eliminar El Tipo de Cambio Con Fecha = " & dtpFecha(4).Value & "  (S/N)", vbInformation + vbYesNo, m_Titulo) = vbYes Then
                        FilEli = mshMaestros.row
                        BorrarRegistro
                        mshMaestros.Clear
                        ConfigMshMaestros
                        LlenarMshMaestros
                        Limpia_Valores
                        DesplazarPorLaGrilla
                        ModoNormal
                        BotonNormal
                        If FilEli > 1 Then
                            mshMaestros.row = FilEli - 1
                            mshMaestros.SetFocus
                            'SendKeys "{HOME}"
                            Call keybd_event(vbKeyHome, 0, 0, 0)
                        Else
                            mshMaestros.row = 1
                        End If
                        LblMensaje = Empty
                    End If
                Else
                    MsgBox "Seleccione El Registro A Eliminar", vbInformation, "STC Centro De Costo"
                    mshMaestros.SetFocus
                End If
    End Select
End Sub

Private Sub cmdGrabar_Click()
    If ValidarData = True Then
        Me.MousePointer = vbHourglass
        If dtpFecha(4).Visible = True Then
            TxtCriterio = dtpFecha(4).Value
        Else
            TxtCriterio = T(0).Text
        End If
        GrabarData
        ActualizarCicloDoc
        mshMaestros.Clear
        ConfigMshMaestros
        LlenarMshMaestros
        DoEvents
        Limpia_Valores
        ModoNormal
        BotonNormal
        FilSel = BuscarCriterio
        mshMaestros.row = FilSel
        DesplazarPorLaGrilla
        LblMensaje = Empty
        Me.MousePointer = vbNormal
        Call cboCampos_Click
        mshMaestros.SetFocus
        'SendKeys "{HOME}"
        Call keybd_event(vbKeyHome, 0, 0, 0)
    Else
        Call cboCampos_Click
    End If
End Sub

Private Sub ActualizarCicloDoc()
Dim SQL As String
    Select Case T(3).Text
           Case "1"
                SQL = "call Update_CicloDoc ('" & T(0) & "','" & REGISTRADO & "', " & _
                      "'" & CANCELADO & "','00','00','00','00','00','00','00','00')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
           Case "4"
                SQL = "call Update_CicloDoc ('" & T(0) & "','" & REGISTRADO & "', " & _
                      "'" & REVISADO & "','00','00','00','00','00','00','00','00')"
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, False
    End Select
End Sub

Private Sub cmdModificar_Click()
    ModoEdicion
    LblMensaje = "Modificar"
    BotonEdicion
End Sub

Private Sub cmdNuevo_Click()
    ModoEdicion
    Limpia_Valores
    BotonEdicion
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_AREAS, FORM_MAESTRO_PERFILES, _
                FORM_MAESTROS_FAM, FORM_MAESTROS_CARGOS, FORM_MAESTRO_TARIFAS, FORM_MAESTRO_SERVxTAR, FORM_MAESTRO_FORMPAGO, _
                FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_TIPOSDESC, FORM_MAESTRO_UM, _
                FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, FORM_MAESTRO_AFP, _
                FORM_MAESTRO_SEGUROS, FORM_MAESTRO_ALMACEN, FORM_MAESTRO_BANCOS, _
                FORM_MAESTRO_CONTRATO, FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_LINEAS, _
                FORM_MAESTRO_AGENCIAS, FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_DEPARTAMENTO, _
                FORM_MAESTRO_GRADOS, FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                T(0) = GenerarCodigo(m_tipo)
           Case FORM_MAESTRO_TIP_DOC, FORM_MAESTROS_CENCO, FORM_MAESTRO_DISTRITO
                T(0).Enabled = True
                T(0).Locked = False
                T(0).BackColor = ColorHabilitado
           Case FORM_MAESTRO_TIPO_CAMBIO
                dtpFecha(4).Enabled = True
                dtpFecha(4).Value = Date
    End Select
    If T(0).Enabled Then T(0).SetFocus
    'SendKeys "{HOME}+{END}"
    Call keybd_event(vbKeyHome, 0, 0, 0)
    LblMensaje = "Nuevo"
End Sub

Private Function GenerarCodigo(Tipo As TIPO_FORMULARIO) As String
    Dim SQL As String
    Dim sql2 As String
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    
    Select Case Tipo
           Case FORM_MAESTRO_ESTADO
                SQL = "select MAX(cod_estado) from doc_estado where cod_estado < '55' "
                sql2 = "select cod_estado from doc_estado LIMIT 1"
           Case FORM_MAESTRO_COMPAÑIAS
                SQL = "select MAX(codcia) from empresas"
                sql2 = "select codcia from empresas LIMIT 1"
           Case FORM_MAESTRO_AREAS
                SQL = "select MAX(idarea) from areas"
                sql2 = "select idarea from areas LIMIT 1"
           Case FORM_MAESTRO_PERFILES
                SQL = "select MAX(perfil_id) from perfiles"
                sql2 = "select perfil_id from perfiles LIMIT 1"
           Case FORM_MAESTROS_CARGOS
                SQL = "select MAX(Codigo) from  cnCargos"
                sql2 = "select Codigo from cnCargos LIMIT 1"
           Case FORM_MAESTROS_FAM
                SQL = "Select MAX(Cod_Fam) from familia_documento"
                sql2 = "Select Cod_Fam from familia_documento LIMIT 1"
           Case FORM_MAESTRO_TARIFAS
                SQL = "Select max(CODTAR) from tarifa"
                sql2 = "Select CODTAR from tarifa LIMIT 1"
           Case FORM_MAESTRO_FORMPAGO
                SQL = "Select max(CODIGO) from forma_pago"
                sql2 = "Select CODIGO from forma_pago LIMIT 1"
           Case FORM_MAESTRO_TIPOPAGO
                SQL = "Select max(codpago) from tipopago"
                sql2 = "Select codpago from tipopago LIMIT 1"
           Case FORM_MAESTRO_TIPOSDESC
                SQL = "Select max(codigo) from descuentos"
                sql2 = "Select codigo from descuentos LIMIT 1"
           Case FORM_MAESTRO_UM
                SQL = "Select max(codigo) from und_med"
                sql2 = "Select codigo from und_med LIMIT 1"
           Case FORM_MAESTRO_CLASF
                SQL = "Select max(codigo) from clasif_serv"
                sql2 = "Select codigo from clasif_serv LIMIT 1"
           Case FORM_MAESTRO_PARENT
                SQL = "Select max(codigo) from pl_vinculofam"
                sql2 = "Select codigo from pl_vinculofam LIMIT 1"
           Case FORM_MAESTRO_AFP
                SQL = "Select max(codigo) from AFP"
                sql2 = "Select codigo from AFP LIMIT 1"
           Case FORM_MAESTRO_SEGUROS
                SQL = "Select max(codigo) from seguro"
                sql2 = "Select codigo from seguro LIMIT 1"
           Case FORM_MAESTRO_ALMACEN
                SQL = "Select max(codigo) from almacen"
                sql2 = "Select codigo from almacen LIMIT 1"
           Case FORM_MAESTRO_BANCOS
                SQL = "Select max(codigo) from pl_entidadfinanciera"
                sql2 = "Select codigo from pl_entidadfinanciera LIMIT 1"
           Case FORM_MAESTRO_CONTRATO
                SQL = "Select max(codigo) from cncontrato"
                sql2 = "Select codigo from cncontrato LIMIT 1"
           Case FORM_MAESTRO_OPERACIONES
                SQL = "Select max(CODOPE) from tipo_OperacionPago"
                sql2 = "Select CODOPE from tipo_OperacionPago LIMIT 1"
           Case FORM_MAESTRO_AGENCIAS
                SQL = "Select max(codigo) from agencia"
                sql2 = "Select codigo from agencia LIMIT 1"
           Case FORM_MAESTRO_LINEAS
                SQL = "Select max(codigo) from linea"
                sql2 = "Select codigo from linea LIMIT 1"
           Case FORM_MAESTRO_ESTANCIAS
                SQL = "Select max(codigo) from estancia"
                sql2 = "Select codigo from estancia LIMIT 1"
           Case FORM_MAESTRO_DEPARTAMENTO
                SQL = "Select max(codigo) from departamento"
                sql2 = "Select codigo from departamento LIMIT 1"
           Case FORM_MAESTRO_GRADOS
                SQL = "Select max(convert(coduni,signed)) from pl_universia"
                sql2 = "Select coduni from pl_universia LIMIT 1"
           Case FORM_MAESTRO_TITULOS
                SQL = "Select max(codcarr) from pl_carrprof"
                sql2 = "Select codcarr from pl_carrprof LIMIT 1"
           Case FORM_MAESTRO_CESES
                SQL = "Select max(codigo) from pl_tipocese"
                sql2 = "Select codigo from pl_tipocese LIMIT 1"
    End Select
    Dim I As Integer, aux As String, Longitud As Integer
    Set Rs = oConexion.EjecutaSelectRS(sql2)
    Longitud = Rs.Fields(0).DefinedSize
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    aux = ""
    For I = 1 To Longitud
        aux = aux & "0"
    Next I
    If IsNull(Rs.Fields(0)) Then
        GenerarCodigo = Right(aux & "1", Longitud)
    Else
        GenerarCodigo = Right(aux & CStr(Rs.Fields(0) + 1), Longitud)
    End If
    Set Rs = Nothing
End Function

Private Sub cmdSalir_Click()
    Unload Me
    mdiInicio.Enabled = True
End Sub

Private Sub ConfiguraObjetos()
    Dim I As Integer
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    
    Set Rs = oConexion.EjecutaSelect(m_QueryRep)
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO, FORM_MAESTRO_AREAS, FORM_MAESTRO_PERFILES, _
                FORM_MAESTROS_FAM, FORM_MAESTRO_TARIFAS, _
                FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_PARENT, _
                FORM_MAESTRO_ALMACEN, FORM_MAESTRO_BANCOS, FORM_MAESTRO_CONTRATO, _
                FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, _
                FORM_MAESTRO_GRADOS, FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                'DOS COLUMNAS
                For I = 0 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                If m_Cols < 6 And m_tipo <> FORM_MAESTROS_CARGOS Then
                    dtpFecha(m_Cols + 2).Visible = False
                End If
           Case FORM_MAESTRO_TIPO_CAMBIO
                For I = 1 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                dtpFecha(m_Cols).Visible = True
                lblL(0).Visible = True
                lblL(0).Caption = m_NomCol(0)
           Case FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_UM, FORM_MAESTRO_TIPOSDESC, _
                FORM_MAESTRO_AFP, FORM_MAESTRO_SEGUROS    'TRES COLUMNAS
                For I = 0 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                dtpFecha(m_Cols + 1).Visible = False
           Case FORM_MAESTRO_COMPAÑIAS, FORM_MAESTRO_TIP_DOC, FORM_MAESTRO_AGENCIAS, _
                FORM_MAESTRO_LINEAS, FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_CUENTAS 'CUATRO COLUMNAS
                For I = 0 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                If m_tipo = FORM_MAESTRO_TIP_DOC Then
                    lblL(7).Visible = True
                    lblL(7).Left = 2100
                    lblL(7).Width = 4360
                    lblL(7).AutoSize = True
                End If
                If (m_tipo = FORM_MAESTRO_LINEAS Or m_tipo = FORM_MAESTRO_ESTANCIAS Or m_tipo = FORM_MAESTRO_TIP_DOC Or FORM_MAESTRO_CUENTAS) Then
                    dtpFecha(4).Visible = False
                End If
           Case FORM_MAESTROS_CENCO, FORM_MAESTRO_CLASF  'CUATRO COLUMNAS
                For I = 0 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                If (m_tipo = FORM_MAESTROS_CENCO) Then
                    lblL(5).Left = T(5).Left - 980
                    lblL(6).Visible = True
                    lblL(6).Left = 2600
                    lblL(6).AutoSize = True
                    lblL(7).Visible = True
                    lblL(7).Left = 2600
                    lblL(7).AutoSize = True
                End If
                If (m_tipo = FORM_MAESTRO_CLASF) Then
                    lblL(5).Left = T(5).Left - 980
                    lblL(6).Left = T(6).Left - 980
                End If
                dtpFecha(4).Visible = False
            Case FORM_MAESTROS_CARGOS
                For I = 0 To m_Cols - 1
                    T(I).MaxLength = Rs.Fields(I).DefinedSize
                    lblL(I).Caption = m_NomCol(I)
                    lblL(I).Visible = True
                    T(I).Width = m_AnchoCol(I)
                    T(I).Text = Empty
                    T(I).Visible = True
                Next I
                
                'T(4).Left = T(4).Left + 5000
                lblL(5).Left = lblL(5).Left + 5000
                T(5).Left = lblL(5).Left + 1200
                lblL(4).Left = lblL(5).Left
                T(4).Left = T(5).Left
                dtpFecha(4).Visible = False
    End Select
    If Rs.State = adStateOpen Then Rs.CloseRecordset
    Set Rs = Nothing
End Sub

Private Sub IniciaObjetos()
    Dim I As Integer
    For I = 0 To CANT_OBJETOS - 1
        lblL(I).ForeColor = &H80000005
        lblL(I).Visible = False
        T(I).Visible = False
        T(I).Locked = True
        T(I).BackColor = ColorDeshabilitado
    Next I
End Sub

Private Sub ConfiguraFormulario()
    If m_Cols > 4 And (m_tipo = FORM_MAESTROS_CENCO Or m_tipo = FORM_MAESTROS_CARGOS) Then  ' FORM_MAESTRO_CUENTAS Or m_tipo <> FORM_MAESTRO_SERVxTAR Or m_tipo = FORM_MAESTRO_FORMPAGO Then
        Me.Width = 15000
        Me.mshMaestros.Width = 14550
        Me.Frame1.Width = 14775
        Me.Frame4.Width = 14775
        Me.FrameP.Width = 14775
    End If
End Sub

Private Sub CmdVistaPreliminar_Click()
    Set oReporte = New clsReporte
    Select Case m_tipo
           Case FORM_MAESTROS_CENCO
                oReporte.empresa = strNombreEmpresa
                oReporte.Titulo = "REPORTE DE CENTRO DE COSTOS"
                oReporte.Reporte = "Rep_CCosto.rpt"
                oReporte.sp_Rep_CentroCostos ("3")
           Case FORM_MAESTRO_TIP_DOC
                oReporte.empresa = strNombreEmpresa
                oReporte.Reporte = "Rep_Tipos de Documentos.rpt"
                oReporte.sp_Rep_TiposDoc
    End Select
End Sub

Private Sub Form_Activate()
    mshMaestros.TopRow = 1
End Sub

Private Sub Form_Load()
    Activo(m_tipo) = True
    Me.Left = 0
    Me.Top = 0
    Me.Caption = m_Titulo
    ConfiguraFormulario
    IniciaObjetos
    ConfiguraObjetos
    ' Parar saber si se configura el combo
    cboActa.Visible = False
    lblL(8).Visible = False
    LblMensaje = Empty
    ConfigMshMaestros
    LlenarMshMaestros
    DoEvents
    ModoNormal
    BotonNormal
    DesplazarPorLaGrilla
End Sub

Sub LlenarMshMaestros()
    Dim Pos As Integer, pos2 As Integer
    Dim Rs As MYSQL_RS
    Set Rs = New MYSQL_RS
    
    'Vista Generica
    'm_QueryRep = m_QueryRep & " where " & criterio & " like '%" & FILTRO & "%' ORDER BY " & criterio
   
    Set Rs = oConexion.EjecutaSelect(m_QueryRep)
    
    Dim I As Integer, J As Integer
    With mshMaestros
        .Redraw = False
        If Not (Rs.BOF And Rs.EOF) Then
            For I = 0 To Rs.RecordCount - 1
                For J = 0 To m_Cols - 1
                    .TextMatrix(.Rows - 1, J) = IIf(IsNull(Rs.Fields(J)), "", " " & Rs.Fields(J))
                Next J
                .Rows = .Rows + 1
                Rs.MoveNext
            Next
            .Rows = .Rows - 1
        Else
            BloqueoDeBotones
        End If
        .Redraw = True
    End With
    
    If Rs.State = adStateOpen Then Rs.CloseRecordset
    Set Rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If LblMensaje <> Empty Then
        Dim M As String
        M = MsgBox("¿ Desea Guardar los Cambios Realizados ?", vbInformation + vbYesNoCancel, Caption)
        Select Case M
               Case vbYes
                    If ValidarData = True Then
                        Call GrabarData
                        Cancel = False
                        mdiInicio.Enabled = True
                        Activo(m_tipo) = False
                    Else
                        Cancel = True
                    End If
               Case vbNo
                    Cancel = False
                    mdiInicio.Enabled = True
                    Activo(m_tipo) = False
               Case vbCancel
                    Cancel = True
        End Select
    Else
        If mdiInicio.Picture1.Visible = False And ModActivo = 0 Then mdiInicio.Picture1.Visible = True
        Cancel = False
        mdiInicio.Enabled = True
        Activo(m_tipo) = False
    End If
End Sub

Private Function Query_Insert() As String
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO
                Query_Insert = "Call Insert_Doc_Estado('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTROS_CENCO
                Query_Insert = "Call Insert_CenCos ('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "','" & T(4) & "', '" & T(5) & "');"
           Case FORM_MAESTRO_TIP_DOC
                If T(3).Text <> "" Then
                    Query_Insert = "Call Insert_Tipo_Doc ('" & T(0) & "','" & Trim(T(1)) & "', '" & Trim(T(2)) & "',1," & Trim(T(3)) & ",'" & Trim(T(4)) & "'); "
                Else
                    Query_Insert = "Call Insert_Tipo_Doc ('" & T(0) & "','" & Trim(T(1)) & "', '" & Trim(T(2)) & "',1,'1','" & Trim(T(4)) & "'); "
                End If
           Case FORM_MAESTRO_TIPO_CAMBIO
                Query_Insert = "Call Insert_Tipo_Cambio ('" & dtpFecha(4) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "') ;"
           Case FORM_MAESTRO_COMPAÑIAS
                Query_Insert = "Call Insert_Empresas('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "');"
           Case FORM_MAESTRO_AREAS
                Query_Insert = "Call Insert_Areas('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTRO_PERFILES
                Query_Insert = "Call Insert_Perfiles('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTROS_CARGOS
                Query_Insert = "Call Insert_Cargos('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "','" & T(4) & "','" & T(5) & "');"
           Case FORM_MAESTROS_FAM
                Query_Insert = "Call Insert_FamDoc ('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTRO_TARIFAS
                Query_Insert = "call Insert_Tarifa  ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_SERVxTAR
                Query_Insert = "call Insert_ServTar  ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "');"
           Case FORM_MAESTRO_FORMPAGO
                Query_Insert = "Call Insert_FormaPago  ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_TIPOPAGO
                Query_Insert = "Call Insert_TipoPago('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_TIPOSDESC
                Query_Insert = "Call Insert_Descuentos ('" & T(0) & "','" & Trim(T(1)) & "', '" & Trim(T(2)) & "'); "
           Case FORM_MAESTRO_UM
                Query_Insert = "Call Insert_UMed ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_CLASF
                Query_Insert = "Call Insert_Clasf ('" & T(0) & "', '" & T(1) & "','" & T(2) & T(3) & "','" & T(4) & "','" & T(5) & "','" & T(6) & "');"
           Case FORM_MAESTRO_PARENT
                Query_Insert = "Call Insert_Parentesco ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_AFP
                Query_Insert = "Call Insert_AFP ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_SEGUROS
                Query_Insert = "Call Insert_Seguro ('" & T(0) & "', '" & T(1) & "', " & T(2) & ");"
           Case FORM_MAESTRO_ALMACEN
                Query_Insert = "Call Insert_Almacen ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_BANCOS
                Query_Insert = "Call Insert_Banco ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_CONTRATO
                Query_Insert = "Call Insert_cncont ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_OPERACIONES
                Query_Insert = "Call Insert_OperacionPago ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_AGENCIAS
                Query_Insert = "Call Insert_Agencia ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "');"
           Case FORM_MAESTRO_LINEAS
                Query_Insert = "Call Insert_Linea ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "');"
           Case FORM_MAESTRO_ESTANCIAS
                Query_Insert = "Call Insert_Estancia ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "');"
           Case FORM_MAESTRO_DISTRITO
                Query_Insert = "Call Insert_Distrito ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_DEPARTAMENTO
                Query_Insert = "Call Insert_Departamento ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_GRADOS
                Query_Insert = "Call Insert_Universia ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_TITULOS
                Query_Insert = "Call Insert_CarrProf ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_CESES
                Query_Insert = "Call Insert_TiposCese ('" & T(0) & "', '" & T(1) & "');"
    End Select
End Function

Private Function Query_Update() As String
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO
                Query_Update = "Call Update_Doc_Estado ('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTRO_TIP_DOC
                Query_Update = "Call Update_Tipo_Doc ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "',0," & T(3) & ",'" & T(4) & "');"
           Case FORM_MAESTROS_CENCO
                Query_Update = "Call Update_CenCos ('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "', '" & T(4) & "', '" & T(5) & "');"
           Case FORM_MAESTRO_TIPO_CAMBIO
                Query_Update = "Call Update_Tipo_Cambio ('" & dtpFecha(4) & "', " & T(1) & ", " & T(2) & ", " & T(3) & ");  "
           Case FORM_MAESTRO_COMPAÑIAS
                Query_Update = "Call Update_Empresas('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "');"
           Case FORM_MAESTRO_AREAS
                Query_Update = "Call Update_Areas ('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTRO_PERFILES
                Query_Update = "Call Update_Perfiles ('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTROS_CARGOS
                Query_Update = "Call Update_Cargos ('" & T(0) & "','" & T(1) & "','" & T(2) & "','" & T(3) & "','" & T(4) & "','" & T(5) & "');"
           Case FORM_MAESTROS_FAM
                Query_Update = "Call Update_FamDoc ('" & T(0) & "','" & T(1) & "');"
           Case FORM_MAESTRO_TARIFAS
                Query_Update = "call Update_Tarifa  ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_SERVxTAR
                Query_Update = "call Update_ServTar  ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "');"
           Case FORM_MAESTRO_FORMPAGO
                Query_Update = "call Update_FormaPago  ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_TIPOPAGO
                Query_Update = "Call Update_TipoPago('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_TIPOSDESC
                Query_Update = "Call Update_Descuentos ('" & T(0) & "','" & Trim(T(1)) & "', '" & Trim(T(2)) & "'); "
           Case FORM_MAESTRO_UM
                Query_Update = "Call Update_UMed ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_CLASF
                Query_Update = "Call Update_Clasf ('" & T(0) & "', '" & T(1) & "','" & T(2) & T(3) & "','" & T(4) & "','" & T(5) & "','" & T(6) & "');"
           Case FORM_MAESTRO_PARENT
                Query_Update = "Call Update_Parentesco ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_AFP
                Query_Update = "Call Update_AFP ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_SEGUROS
                Query_Update = "Call Update_Seguro ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "');"
           Case FORM_MAESTRO_ALMACEN
                Query_Update = "Call Update_Almacen ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_BANCOS
                Query_Update = "Call Update_Banco ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_CONTRATO
                Query_Update = "Call Update_cncont ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_OPERACIONES
                Query_Update = "Call Update_OperacionPago ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_AGENCIAS
                Query_Update = "Call Update_Agencia ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "');"
           Case FORM_MAESTRO_LINEAS
                Query_Update = "Call Update_Linea ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "');"
           Case FORM_MAESTRO_ESTANCIAS
                Query_Update = "Call Update_Estancia ('" & T(0) & "', '" & T(1) & "', '" & T(2) & "', '" & T(3) & "','" & T(4) & "' );"
           Case FORM_MAESTRO_DISTRITO
                Query_Update = "Call Update_Distrito ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_DEPARTAMENTO
                Query_Update = "Call Update_Departamento ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_GRADOS
                Query_Update = "Call Update_Universia ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_TITULOS
                Query_Update = "Call Update_CarrProf ('" & T(0) & "', '" & T(1) & "');"
           Case FORM_MAESTRO_CESES
                Query_Update = "Call Update_TiposCese ('" & T(0) & "', '" & T(1) & "');"
    End Select
End Function

Private Function Query_Delete() As String
    Select Case m_tipo
           Case FORM_MAESTRO_ESTADO
                Query_Delete = " Call Delete_Doc_Estado ('" & T(0) & "');"
           Case FORM_MAESTROS_CENCO
                Query_Delete = " Call Delete_CenCos ('" & T(0) & "');"
           Case FORM_MAESTRO_TIP_DOC
                Query_Delete = " Call Delete_Tipo_Doc ('" & T(0) & "');"
           Case FORM_MAESTRO_TIPO_CAMBIO
                Query_Delete = "Call Delete_Tipo_Cambio('" & dtpFecha(4).Value & "');  "
           Case FORM_MAESTRO_COMPAÑIAS
                Query_Delete = "Call Delete_Empresas ('" & T(0) & "'); "
           Case FORM_MAESTRO_AREAS
                Query_Delete = " Call Delete_Areas ('" & T(0) & "');"
           Case FORM_MAESTRO_PERFILES
                Query_Delete = " Call Delete_Perfiles ('" & T(0) & "');"
           Case FORM_MAESTROS_CARGOS
                Query_Delete = " Call Delete_Cargos ('" & T(0) & "');"
           Case FORM_MAESTROS_FAM
                Query_Delete = "Call Delete_FamDoc ('" & T(0) & "');"
           Case FORM_MAESTRO_TARIFAS
                Query_Delete = "call Delete_Tarifa  ('" & T(0) & "');"
           Case FORM_MAESTRO_SERVxTAR
                Query_Delete = "call Delete_ServTar  ('" & T(0) & "');"
           Case FORM_MAESTRO_FORMPAGO
                Query_Delete = "call Delete_FormaPago ('" & T(0) & "');"
           Case FORM_MAESTRO_TIPOPAGO
                Query_Delete = "Call Delete_TipoPago('" & T(0) & "');"
           Case FORM_MAESTRO_TIPOSDESC
                Query_Delete = "Call Delete_Descuentos('" & T(0) & "');"
           Case FORM_MAESTRO_UM
                Query_Delete = "Call Delete_UMed('" & T(0) & "');"
           Case FORM_MAESTRO_CLASF
                Query_Delete = "Call Delete_Clasf('" & T(0) & "');"
           Case FORM_MAESTRO_PARENT
                Query_Delete = "Call Delete_Parentesco('" & T(0) & "');"
           Case FORM_MAESTRO_AFP
                Query_Delete = "Call Delete_AFP ('" & T(0) & "');"
           Case FORM_MAESTRO_SEGUROS
                Query_Delete = "Call Delete_Seguro ('" & T(0) & "');"
           Case FORM_MAESTRO_ALMACEN
                Query_Delete = "Call Delete_Almacen ('" & T(0) & "');"
           Case FORM_MAESTRO_BANCOS
                Query_Delete = "Call Delete_Banco ('" & T(0) & "');"
           Case FORM_MAESTRO_CONTRATO
                Query_Delete = "Call Delete_cncont ('" & T(0) & "');"
           Case FORM_MAESTRO_OPERACIONES
                Query_Delete = "Call Delete_OperacionPago ('" & T(0) & "');"
           Case FORM_MAESTRO_AGENCIAS
                Query_Delete = "Call Delete_Agencia ('" & T(0) & "');"
           Case FORM_MAESTRO_LINEAS
                Query_Delete = "Call Delete_Linea ('" & T(0) & "');"
           Case FORM_MAESTRO_ESTANCIAS
                Query_Delete = "Call Delete_Estancia ('" & T(0) & "');"
           Case FORM_MAESTRO_DISTRITO
                Query_Delete = "Call Delete_Distrito ('" & T(0) & "');"
           Case FORM_MAESTRO_DEPARTAMENTO
                Query_Delete = "Call Delete_Departamento ('" & T(0) & "');"
           Case FORM_MAESTRO_GRADOS
                Query_Delete = "Call Delete_Universia ('" & T(0) & "');"
           Case FORM_MAESTRO_TITULOS
                Query_Delete = "Call Delete_CarrProf ('" & T(0) & "');"
           Case FORM_MAESTRO_CESES
                Query_Delete = "Call Delete_TiposCese ('" & T(0) & "');"
    End Select
End Function

Sub GrabarData()
On Error GoTo ErrSave
    Dim SQL As String
    Dim I As Integer
    
    Select Case LblMensaje.Caption
           Case "Nuevo"
                SQL = Query_Insert
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, True
                If m_tipo = FORM_MAESTRO_ESTADO Then GrabaEstado T(0).Text
           Case "Modificar"
                SQL = Query_Update
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Modificar, True
           Case "Eliminar"
                SQL = Query_Delete
                oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, True
    End Select
Exit Sub
ErrSave:
    MsgBox "Ha ocurrido un error al momento de grabar" & Chr(13) & err.Description, vbCritical, "Error de datos"
    oConexion.DeshacerTransaccion
End Sub

Sub BorrarRegistro()
On Error GoTo Errdel
    oConexion.EjecutaInsertUpdateDelete Query_Delete, TIPO_QUERY.Eliminar, True
Exit Sub
Errdel:
    MsgBox "Ha ocurrido un error al momento de eliminar" & Chr(13) & err.Description, vbCritical, "Error de datos"
End Sub

Public Function ValidarData() As Boolean
    Dim I As Integer
    For I = 0 To m_Cols - 1
        If lblL(I) Like "*Just*" And (T(I) <> "S" And T(I) <> "N") Then
            MsgBox "El Item " & lblL(I) & " debe ser S o N", vbInformation, m_Titulo
            ValidarData = False
            T(I) = UCase(T(I))
            T(I).SetFocus
            'SendKeys "{HOME}+{END}"
            Call keybd_event(vbKeyHome, 0, 0, 0)
            Exit Function
        End If
        If T(6) = "N" Then T(7) = Empty
        If Trim(T(I)) = Empty And T(6) = "S" Then
            MsgBox "El Item " & lblL(I) & " Está En Blanco", vbInformation, m_Titulo
            ValidarData = False
            T(I).SetFocus
            'SendKeys "{HOME}+{END}"
            Call keybd_event(vbKeyHome, 0, 0, 0)
            Exit Function
        End If
    Next I
    ValidarData = True
End Function

Private Sub Frame1_Click()
    txtObs_LostFocus
End Sub

Private Sub mshMaestros_Click()
    DesplazarPorLaGrilla
End Sub

Private Sub mshMaestros_DblClick()
    cmdModificar_Click
End Sub

Private Sub mshMaestros_KeyDown(KeyCode As Integer, Shift As Integer)
    DesplazarPorLaGrilla
    Call Ayuda(KeyCode)
End Sub

Sub DesplazarPorLaGrilla()
    Dim I As Integer
    If MshHabilitado = True Then
        Select Case m_tipo
               Case FORM_MAESTRO_TIP_DOC, FORM_MAESTRO_COMPAÑIAS, FORM_MAESTRO_AREAS, _
                    FORM_MAESTROS_CENCO, FORM_MAESTRO_PERFILES, FORM_MAESTROS_FAM, _
                    FORM_MAESTROS_CARGOS, FORM_MAESTRO_TARIFAS, FORM_MAESTRO_UM, _
                    FORM_MAESTRO_FORMPAGO, FORM_MAESTRO_TIPOPAGO, FORM_MAESTRO_TIPOSDESC, _
                    FORM_MAESTRO_CLASF, FORM_MAESTRO_PARENT, FORM_MAESTRO_AFP, _
                    FORM_MAESTRO_SEGUROS, FORM_MAESTRO_ALMACEN, FORM_MAESTRO_BANCOS, _
                    FORM_MAESTRO_CONTRATO, FORM_MAESTRO_OPERACIONES, FORM_MAESTRO_AGENCIAS, _
                    FORM_MAESTRO_LINEAS, FORM_MAESTRO_ESTANCIAS, FORM_MAESTRO_CUENTAS, _
                    FORM_MAESTRO_DISTRITO, FORM_MAESTRO_DEPARTAMENTO, FORM_MAESTRO_GRADOS, _
                    FORM_MAESTRO_TITULOS, FORM_MAESTRO_CESES
                    With mshMaestros
                        For I = 0 To m_Cols - 1
                            T(I).Locked = True
                            T(I).BackColor = ColorDeshabilitado
                            T(I).Text = Trim(.TextMatrix(.Rowsel, I))
                        Next I
                        If m_tipo = FORM_MAESTROS_CENCO Then
                            lblL(6).ForeColor = &H8080FF
                            lblL(6).Caption = DescripcionesdeCodigos("ENCARGADOS", Trim(.TextMatrix(.Rowsel, 2)))
                            lblL(7).ForeColor = &H8080FF
                            lblL(7).Caption = DescripcionesdeCodigos("DES_DIVISION", Trim(.TextMatrix(.Rowsel, 3)))
                            lblL(7).AutoSize = True
                        End If
                        If m_tipo = FORM_MAESTRO_TIP_DOC Then
                            lblL(7).ForeColor = &H8080FF
                            lblL(7).Caption = DescripcionesdeCodigos("DES_FAMILIA", Trim(.TextMatrix(.Rowsel, 3)))
                        End If
                    End With
               Case FORM_MAESTRO_TIPO_CAMBIO
                    With mshMaestros
                        For I = 1 To m_Cols - 1
                            T(I).Locked = True
                            T(I).BackColor = ColorDeshabilitado
                            T(I).Text = Trim(.TextMatrix(.Rowsel, I))
                            If Trim(.TextMatrix(.Rowsel, 0)) <> Empty Then
                                dtpFecha(m_Cols).Value = Trim(.TextMatrix(.Rowsel, 0))
                            Else
                                dtpFecha(m_Cols).Value = Date
                            End If
                        Next I
                    End With
               Case FORM_MAESTRO_ESTADO
                    With mshMaestros
                        For I = 0 To m_Cols - 1
                            T(I).Locked = True
                            T(I).BackColor = ColorDeshabilitado
                            T(I).Text = Trim(.TextMatrix(.Rowsel, I))
                            If T(0) Like "*0*" Then
                                CmdModificar.Enabled = True
                                Else
                                CmdModificar.Enabled = False
                            End If
                        Next I
                    End With
        End Select
    End If
End Sub

Private Sub OptCodigo_Click()
    TxtCriterio.SetFocus
End Sub

Private Sub OptDescripcion_Click()
    TxtCriterio.SetFocus
End Sub

Private Sub T_Change(Index As Integer)
    T(Index) = UCase(T(Index))
    T(Index).SelStart = Len(T(Index))
End Sub

Private Sub T_GotFocus(Index As Integer)
    If Index = 7 Then
       txtObs.Visible = True
       txtObs.Text = T(7)
       txtObs.SetFocus
       txtObs.SelLength = Len(txtObs)
    End If
End Sub

Private Sub TxtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call Ayuda(vbKeyF2): Exit Sub
    If KeyCode = vbKeyF3 Then Call Ayuda(vbKeyF3): Exit Sub
End Sub

Private Sub txtObs_LostFocus()
    txtObs.Visible = False
    T(7) = txtObs.Text
End Sub

Private Sub TxtCriterio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mshMaestros.Col = 0
        Dim fila As Integer
        fila = BuscarCriterio
        If fila > 0 Then
            mshMaestros.row = fila
            mshMaestros.Col = 0
            mshMaestros.SetFocus
            'SendKeys "{HOME}+{END}"
            Call keybd_event(vbKeyHome, 0, 0, 0)
        Else
            MsgBox "Registro No Encontrado", vbInformation, "STC Centro De Costo"
            mshMaestros.row = 1
            mshMaestros.SetFocus
            'SendKeys "{HOME}+{END}"
            Call keybd_event(vbKeyHome, 0, 0, 0)
        End If
        DoEvents
        TxtCriterio.SetFocus
    End If
End Sub

Private Function BuscarCriterio() As Integer
    Dim I As Integer
    Dim criterio As String
    With mshMaestros
        criterio = " *" & UCase(Trim(TxtCriterio)) & "*"
        For I = mshMaestros.row + 1 To mshMaestros.Rows - 1
            If UCase(.TextMatrix(I, cboCampos.ListIndex)) Like criterio Then
                BuscarCriterio = I
                Exit Function
            End If
        Next
    End With
    BuscarCriterio = 0
End Function

Private Sub T_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If T(Index).BackColor = ColorDeshabilitado And T(Index).BackColor = Empty Then Exit Sub
    Dim Texto As String
    Dim oConsulta As FrmConsultas
    Set oConsulta = New FrmConsultas
    Select Case KeyCode
           Case vbKeyF1
                Texto = lblL(Index).Caption
                If Texto Like "Servicio" And m_tipo = FORM_MAESTRO_SERVxTAR Then
                    With oConsulta
                        .pCols = 3
                        .pCol = 0: .pAnchoCol = 1200
                        .pCol = 1: .pAnchoCol = 1200
                        .pCol = 2: .pAnchoCol = 2000
                        .pTitulo = "Servicio"
                        .pForm = FORM_MAESTRO_SERVxTAR
                        .pCaso = LABEL_SERVICIOS
                        .Show
                    End With
                End If
                If Texto Like "Tarifa" And m_tipo = FORM_MAESTRO_SERVxTAR Then
                    With oConsulta
                        .pCols = 2
                        .pCol = 0: .pAnchoCol = 1200
                        .pCol = 1: .pAnchoCol = 3500
                        .pTitulo = "Tarifa"
                        .pForm = FORM_MAESTRO_SERVxTAR
                        .pCaso = LABEL_TARIFAS
                        .Show
                    End With
                End If
                If Texto Like "Encargado" And m_tipo = FORM_MAESTROS_CENCO And T(Index).BackColor = ColorHabilitado Then
                    With oConsulta
                        .pCols = 2
                        .pCol = 0: .pAnchoCol = 1200
                        .pCol = 1: .pAnchoCol = 3500
                        .pTitulo = "Empleados"
                        .pForm = FORM_MAESTROS_CENCO
                        .pCaso = Label_Descrip_Auxil
                        .Show
                    End With
                End If
                If Texto Like "ccHFM" And m_tipo = FORM_MAESTROS_CENCO And T(Index).BackColor = ColorHabilitado Then
                    With oConsulta
                        .pCols = 2
                        .pCol = 0: .pAnchoCol = 600
                        .pCol = 1: .pAnchoCol = 2500
                        .pTitulo = "ccHFM"
                        .pForm = FORM_MAESTROS_CENCO
                        .pCaso = LABEL_DIVISIONES
                        .Show
                    End With
                End If
                If Texto Like "Familia" And m_tipo = FORM_MAESTRO_TIP_DOC And T(Index).BackColor = ColorHabilitado Then
                    With oConsulta
                        .pCols = 2
                        .pCol = 0: .pAnchoCol = 900
                        .pCol = 1: .pAnchoCol = 4000
                        .pTitulo = "Tipos de Familia"
                        .pForm = FORM_MAESTRO_TIP_DOC
                        .pCaso = LABEL_FAMILIA
                        .Show
                    End With
                End If
           Case vbKeyF4
                If lblL(Index) Like "Hora*" Then T(Index) = Format(Time, "hh:mm:ss")
                If lblL(Index) Like "Fecha*" Then T(Index) = Format(Date, "dd/mm/yyyy")
                'SendKeys "{HOME}+{END}"
                Call keybd_event(vbKeyHome, 0, 0, 0)
    End Select
End Sub

Private Sub T_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 27 And LblMensaje <> "Modificar" Then
        Unload Me
        Exit Sub
    End If
    If Index = 0 Then
        If KeyAscii = 13 Then
            Dim fila As Integer
            fila = BuscarMaestro(Trim(T(0)))
            If fila > 0 Then
                If LblMensaje = "Modificar" Or LblMensaje = "Eliminar" Then
                    With mshMaestros
                        T(0) = .TextMatrix(fila, 0)
                        T(1) = .TextMatrix(fila, 1)
                    End With
                    Call BloqueoEspecial
                End If
                If LblMensaje = "Nuevo" Then
                    MsgBox "El Registro Ya existe", vbInformation, m_Titulo
                    'SendKeys "{HOME}+{END}"
                    Call keybd_event(vbKeyHome, 0, 0, 0)
                End If
            Else
                If LblMensaje = "Modificar" Or LblMensaje = "Eliminar" Then
                    MsgBox "El Registro no existe", vbInformation, m_Titulo
                    'SendKeys "{HOME}+{END}"
                    Call keybd_event(vbKeyHome, 0, 0, 0)
                End If
                If LblMensaje = "Nuevo" Then
                
                End If
            End If
        Else
            If KeyAscii = 8 Then
                If T(0).Locked = False Then
                    T(1) = Empty
                End If
            End If
        End If
    End If
    If Index < m_Cols - 1 And KeyAscii = 13 Then
        If T(Index + 1).Visible Then T(Index + 1).SetFocus
    End If
    If Index = m_Cols - 1 Then
        If KeyAscii = 13 And (lblL(Index) = "Sustento" Or lblL(Index) Like "*Obser*") Then Exit Sub
        If KeyAscii = 13 And T(Index).Locked = False Then cmdGrabar_Click
    End If
End Sub

Private Function BuscarMaestro(codigo As String) As Integer
    Dim I As Integer
    With mshMaestros
        For I = 0 To mshMaestros.Rows - 1
            If .TextMatrix(I, 0) = Trim(codigo) Then
                BuscarMaestro = I
                Exit Function
            End If
        Next
        BuscarMaestro = 0
    End With
End Function
