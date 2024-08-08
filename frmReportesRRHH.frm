VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmReportesRRHH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cofiguración de Reportes RRHH"
   ClientHeight    =   6735
   ClientLeft      =   3120
   ClientTop       =   3390
   ClientWidth     =   10725
   Icon            =   "frmReportesRRHH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10725
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4095
      Left            =   0
      TabIndex        =   11
      Top             =   2640
      Width           =   10725
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxReporte 
         Height          =   3795
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   6694
         _Version        =   393216
         BackColor       =   16777215
         BackColorBkg    =   8421504
         GridColor       =   8421504
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblNumRegistros 
         Alignment       =   2  'Center
         BackColor       =   &H009F5539&
         Height          =   285
         Left            =   6960
         TabIndex        =   14
         Top             =   3630
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   2730
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   10725
      Begin Proyecto1.chameleonButton btnBajar 
         Height          =   375
         Left            =   5895
         TabIndex        =   1
         ToolTipText     =   "Desplazar hacia Abajo"
         Top             =   1620
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":014A
         PICN            =   "frmReportesRRHH.frx":0166
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnSubir 
         Height          =   375
         Left            =   5895
         TabIndex        =   2
         ToolTipText     =   "Desplazar Hacia Arriba"
         Top             =   1140
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":01EC
         PICN            =   "frmReportesRRHH.frx":0208
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnTodos 
         Height          =   375
         Left            =   2745
         TabIndex        =   7
         Top             =   720
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":028E
         PICN            =   "frmReportesRRHH.frx":02AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnNinguno 
         Height          =   375
         Left            =   2745
         TabIndex        =   8
         Top             =   2010
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":0343
         PICN            =   "frmReportesRRHH.frx":035F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnAgregar 
         Height          =   375
         Left            =   2745
         TabIndex        =   9
         Top             =   1140
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":03FB
         PICN            =   "frmReportesRRHH.frx":0417
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnQuitar 
         Height          =   375
         Left            =   2745
         TabIndex        =   10
         Top             =   1575
         Width           =   435
         _ExtentX        =   767
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
         MICON           =   "frmReportesRRHH.frx":04A0
         PICN            =   "frmReportesRRHH.frx":04BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton btnVerGrilla 
         Height          =   465
         Left            =   7410
         TabIndex        =   12
         ToolTipText     =   "Ver Datos"
         Top             =   1320
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   820
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
         MICON           =   "frmReportesRRHH.frx":0545
         PICN            =   "frmReportesRRHH.frx":0561
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
         Height          =   465
         Left            =   9090
         TabIndex        =   15
         Top             =   1320
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   820
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
         MICON           =   "frmReportesRRHH.frx":06BB
         PICN            =   "frmReportesRRHH.frx":06D7
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
         Height          =   465
         Left            =   8280
         TabIndex        =   16
         Top             =   1320
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   820
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
         MICON           =   "frmReportesRRHH.frx":0831
         PICN            =   "frmReportesRRHH.frx":084D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.CheckBox chkActivos 
         Height          =   315
         Left            =   5895
         TabIndex        =   17
         Top             =   2280
         Width           =   1935
         VariousPropertyBits=   746588179
         BackColor       =   10442041
         ForeColor       =   8421631
         DisplayStyle    =   4
         Size            =   "3413;556"
         Value           =   "0"
         Caption         =   "Mostrar Inactivos"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   315
         Left            =   3330
         TabIndex        =   6
         Top             =   165
         Width           =   2565
         ForeColor       =   16777215
         BackColor       =   10442041
         Caption         =   "Campos Seleccionados:"
         Size            =   "4524;556"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   150
         Width           =   2355
         ForeColor       =   16777215
         BackColor       =   10442041
         Caption         =   "Campos Disponibles:"
         Size            =   "4154;556"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ListBox lstCampos 
         Height          =   2175
         Left            =   150
         TabIndex        =   4
         Top             =   450
         Width           =   2415
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "4260;3836"
         MatchEntry      =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ListBox lstCamposSelec 
         Height          =   2175
         Left            =   3330
         TabIndex        =   3
         Top             =   465
         Width           =   2415
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "4260;3836"
         MatchEntry      =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmReportesRRHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vCampos(0 To 80) As String
Private filtro(1 To 17, 1 To 2) As String
Private tblcampo(1 To 30, 1 To 2) As String
Public cabeceraGrid As String
Private Sub BtnAgregar_Click()
    Dim I As Integer
    If lstCampos.ListIndex >= 0 Then
        With lstCamposSelec
            I = lstCampos.ListIndex
            If I <= lstCampos.ListCount - 1 And lstCamposSelec.ListCount < 25 Then
                If lstCampos.Text <> "SBASICO" Then
                    .AddItem lstCampos.List(I, 0)
                    .List(.ListCount - 1, 1) = lstCampos.List(I, 1)
                    lstCampos.RemoveItem I
                    ConfigSelect .ListCount - 1, lstCamposSelec
                    lstCampos.SetFocus
                Else
                    If VerContratoySueldos Then
                        .AddItem lstCampos.List(I, 0)
                        .List(.ListCount - 1, 1) = lstCampos.List(I, 1)
                        lstCampos.RemoveItem I
                        ConfigSelect .ListCount - 1, lstCamposSelec
                        lstCampos.SetFocus
                    Else
                        MsgBox "Usted no se encuentra autorizado a visualizar este Campo", vbOKOnly + vbExclamation, "NOVPeru"
                    End If
                End If
            End If
        End With
    End If
    lblCamposSel = lstCamposSelec.ListCount & "/15"
    ConfiguraBtns lstCampos.ListCount, lstCamposSelec.ListCount
End Sub
Private Sub ConfigSelect(I As Integer, lst As MSForms.ListBox)
    With lst
        If I = 0 Then .Selected(I) = True: Exit Sub
        If I > 0 Then
            .Selected(I - 1) = False
            .Selected(I) = True
            .ListIndex = I
        End If
    End With
End Sub
Private Sub btnBajar_Click()
    BajarItem lstCamposSelec
End Sub
Private Sub BajarItem(lst As MSForms.ListBox)
    Dim I As Integer
    Dim tmp As String
    Dim tmp1 As String
    With lst
        I = .ListIndex
        If I < .ListCount - 1 Then
            tmp = .List(I + 1, 0)
            tmp1 = .List(I + 1, 1)
            .List(I + 1, 0) = .List(I, 0)
            .List(I + 1, 1) = .List(I, 1)
            .List(I, 0) = tmp
            .List(I, 1) = tmp1
            .ListIndex = I + 1
        End If
    End With
End Sub
Private Sub TablasCampos()
    tblcampo(1, 1) = "codigo": tblcampo(1, 2) = "emp_reg"
    tblcampo(2, 1) = "codcargo": tblcampo(2, 2) = "cncargos"
    tblcampo(3, 1) = "cododocide": tblcampo(3, 2) = "doc_identificacion"
    tblcampo(4, 1) = "distrito": tblcampo(4, 2) = "distrito"
    tblcampo(5, 1) = "codbanco": tblcampo(5, 2) = "pl_entidadfinanciera"
    tblcampo(6, 1) = "tipcta_me": tblcampo(6, 2) = "tipopago"
    tblcampo(7, 1) = "tipcta_mn": tblcampo(7, 2) = "tipopago"
    tblcampo(8, 1) = "ctsbanco": tblcampo(8, 2) = "pl_entidadfinanciera"
    tblcampo(9, 1) = "tipo": tblcampo(9, 2) = "tipoemp"
    tblcampo(10, 1) = "personal": tblcampo(10, 2) = "tippersonal"
    tblcampo(11, 1) = "situacion": tblcampo(11, 2) = "situacionemp"
    tblcampo(12, 1) = "categoria": tblcampo(12, 2) = "rh_categoria"
    tblcampo(13, 1) = "TipBrevete": tblcampo(13, 2) = "tiposbrevete"
    tblcampo(14, 1) = "GSangre": tblcampo(14, 2) = "GSangre"
    tblcampo(15, 1) = "est_civil": tblcampo(15, 2) = "estadocivil"
    tblcampo(16, 1) = "nacionalidad": tblcampo(16, 2) = "nacionalidad"
    tblcampo(17, 1) = "ccHFM": tblcampo(17, 2) = "cnmdepar"
    tblcampo(18, 1) = "codafp": tblcampo(18, 2) = "afp"
    tblcampo(19, 1) = "codtipo": tblcampo(19, 2) = "cncontrato"
    tblcampo(20, 1) = "codgrado": tblcampo(20, 2) = "gradoinstruc"
    tblcampo(21, 1) = "coddocide": tblcampo(21, 2) = "doc_identificacion"
    tblcampo(22, 1) = "esttrabajo": tblcampo(22, 2) = "rh_estacionestrabajo"
    tblcampo(23, 1) = "cencos": tblcampo(23, 2) = "cncosto"
    tblcampo(24, 1) = "divgas": tblcampo(24, 2) = "cnmdepar"
    tblcampo(25, 1) = "codtit": tblcampo(25, 2) = "titulo"
    tblcampo(26, 1) = "departamento": tblcampo(26, 2) = "departamento"
    tblcampo(27, 1) = "codseg": tblcampo(27, 2) = "seguro"
End Sub
Private Sub btnQuitar_Click()
    Dim posRemove As Integer, I As Integer
    On Error GoTo err
    For I = 0 To 40
        If vCampos(I) = lstCamposSelec.List(lstCamposSelec.ListIndex, 0) Then posRemove = I: Exit For
    Next
    PasarDato posRemove, lstCampos.ListCount - 1, lstCamposSelec.ListIndex, lstCampos, lstCamposSelec
    lblCamposSel = lstCamposSelec.ListCount & "/15"
    ConfigSelect lstCamposSelec.ListIndex, lstCamposSelec
    ConfiguraBtns lstCampos.ListCount, lstCamposSelec.ListCount
err:
    Resume Next
    Exit Sub
End Sub
Private Sub PasarDato(posRemove As Integer, contador As Integer, posItem As Integer, lst1 As MSForms.ListBox, lst2 As MSForms.ListBox)
    Dim posInsert As Integer
    Dim J As Integer
    If contador > 0 Then
        For J = 0 To 40
         If lst1.List(contador, 0) = vCampos(J) Then posInsert = J: Exit For
        Next
    End If
    If (posInsert > posRemove) Then
        contador = contador - 1
        PasarDato posRemove, contador, posItem, lst1, lst2
    Else
        If posRemove = 0 Then contador = -1
        lst1.AddItem lst2.List(posItem, 0), contador + 1
        lst1.List(contador + 1, 1) = lst2.List(posItem, 1)
        lst2.RemoveItem posItem
        lst2.SetFocus
        Exit Sub
    End If
End Sub
Private Sub ConfiguraBtns(CountList1 As Integer, CountList2 As Integer)
    If CountList2 = 25 Then btnAgregar.Enabled = False: btnTodos.Enabled = False:   btnQuitar.Enabled = True: _
                            btnNinguno.Enabled = True: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                           Exit Sub
    If CountList2 = 0 And CountList1 > 0 Then btnAgregar.Enabled = True: btnTodos.Enabled = True: _
                                            btnQuitar.Enabled = False: btnNinguno.Enabled = False: _
                                            btnSubir.Enabled = False: btnBajar.Enabled = False: _
                                            Exit Sub
    If CountList2 > 0 And CountList2 <= 25 Then btnAgregar.Enabled = True: btnTodos.Enabled = True: _
                                                btnQuitar.Enabled = True: btnNinguno.Enabled = True: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                                               Exit Sub
    If CountList2 > 0 And CountList2 <= 25 Then btnAgregar.Enabled = True: btnTodos.Enabled = True: _
                                                btnQuitar.Enabled = True: btnNinguno.Enabled = True: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                                                Exit Sub
    If val(CountList2 + CountList3) = 25 Then btnAgregar.Enabled = False: btnTodos.Enabled = False: _
                                                btnQuitar.Enabled = True: btnNinguno.Enabled = False: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                                               Exit Sub
    If CountList2 > 0 And CountList2 <= 25 Then btnAgregar.Enabled = True: btnTodos.Enabled = True: _
                                                btnQuitar.Enabled = True: btnNinguno.Enabled = False: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                                                Exit Sub
    If CountList2 = 1 Then btnAgregar.Enabled = True: btnTodos.Enabled = True: _
                                                btnQuitar.Enabled = False: btnNinguno.Enabled = False: btnSubir.Enabled = True: btnBajar.Enabled = True: _
                                                Exit Sub
End Sub
Private Sub btnNinguno_Click()
    Dim I As Integer, J As Integer, posRemove As Integer
    For J = lstCamposSelec.ListCount - 1 To 0 Step -1
        If Not (J = 1) Then
            For I = 0 To 40
                If vCampos(I) = lstCamposSelec.List(J, 0) Then posRemove = I: Exit For
            Next
            PasarDato posRemove, lstCampos.ListCount - 1, lstCamposSelec.ListIndex, lstCampos, lstCamposSelec
        End If
    Next
    lstCamposSelec.Clear
    lblCamposSel = lstCamposSelec.ListCount & "/15"
    ConfiguraBtns lstCampos.ListCount, lstCamposSelec.ListCount
End Sub
Private Sub btnReporte_Click()
    Me.MousePointer = vbHourglass
    If Exportar_Excel(Rep_Documents & "\REPORTE.XLS", flxReporte) Then
    End If
    Me.MousePointer = vbNormal
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Sub btnSubir_Click()
    SubirItem lstCamposSelec
End Sub
Private Sub SubirItem(lst As MSForms.ListBox)
    Dim I As Integer
    Dim tmp As String
    Dim tmp1 As String
    With lst
        I = .ListIndex
        If I > 0 Then
            tmp = .List(I - 1, 0)
            tmp1 = .List(I - 1, 1)
            .List(I - 1, 0) = .List(I, 0)
            .List(I - 1, 1) = .List(I, 1)
            .List(I, 0) = tmp
            .List(I, 1) = tmp1
            .ListIndex = I - 1
        End If
    End With
End Sub
Private Sub btnTodos_Click()
    Dim I As Integer, J As Integer
    If Not lstCamposSelec.ListCount = 25 Then
        For I = lstCamposSelec.ListCount To 14
            With lstCamposSelec
                .AddItem lstCampos.List(0, 0)
                .List(I, 0) = lstCampos.List(0, 0)
                .List(I, 1) = lstCampos.List(0, 1)
                lstCampos.RemoveItem 0
            End With
        Next
    End If
    lblCamposSel = lstCamposSelec.ListCount & "/15"
    ConfigSelect lstCamposSelec.ListCount - 1, lstCamposSelec
    ConfiguraBtns lstCampos.ListCount, lstCamposSelec.ListCount
End Sub
Private Sub BloquearBotones(valor As Boolean)
    btnTodos.Enabled = valor
    btnAgregar.Enabled = valor
    btnQuitar.Enabled = valor
    btnNinguno.Enabled = valor
    btnSubir.Enabled = valor
    btnBajar.Enabled = valor
End Sub
Private Sub btnVerGrilla_Click()
    Dim grupos As Integer
    If lstCamposSelec.ListCount > 0 Then
        ConfigGrilla
        LlenarGrilla ConsultaGral
        LimpiarVector
    End If
End Sub
Private Sub LimpiarVector()
    Dim I As Integer
    Dim J As Integer
    For I = 1 To 17
        For J = 1 To 2
            filtro(I, J) = ""
        Next
    Next
End Sub

Private Sub flxReporte_Click()
    Dim strcampo As String
    Dim strTbl As String
    Dim I As Integer, columna As Integer
    Dim Colum As String
    With flxReporte
        If .row = 0 Then
            If Trim(.TextMatrix(0, .Col)) = strChecked Then
                strTbl = tabla(.Col)
                If strTbl <> "" Then
                    Colum = Columnas(strTbl)
                    If Colum <> "" Then
                    CambiarDatos .Col, strTbl, Mid(Colum, InStr(1, Colum, ",") + 1, Len(Colum)), Mid(Colum, 1, InStr(1, Colum, ",") - 1)
                    End If
                    .TextMatrix(0, .Col) = strUnChecked
                End If
            Else
                strTbl = tabla(.Col)
                If strTbl <> "" Then
                    Colum = Columnas(strTbl)
                    If Colum <> "" Then
                        CambiarDatos .Col, strTbl, Mid(Colum, 1, InStr(1, Colum, ",") - 1), Mid(Colum, InStr(1, Colum, ",") + 1, Len(Colum))
                    End If
                    .TextMatrix(0, .Col) = strChecked
                End If
           End If
        End If
    End With
End Sub

Private Function tabla(flxcol As Integer) As String
    Dim strcampo As String
    Dim Col As Integer
    Col = flxcol
    strcampo = lstCamposSelec.List(Col, 0)
    For I = 1 To 30
        If strcampo = UCase(tblcampo(I, 1)) Then tabla = tblcampo(I, 2): Exit For
    Next
End Function

Private Sub CambiarDatos(columna As Integer, tabla As String, campo1 As String, campo2 As String)
    Dim SQL As String
    Dim Rs As MYSQL_RS
    Dim rslen As MYSQL_RS
    With flxReporte
        For I = 2 To .Rows - 1
            SQL = "Select " & campo2 & " from " & tabla & " where " & campo1 & " ='" & Trim(.TextMatrix(I, columna)) & "' "
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            If Rs.RecordCount > 0 Then
                .TextMatrix(I, columna) = Trim(Rs.Fields("" & campo2 & ""))
            End If
        Next
    End With
    Set Rs = Nothing
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = 0
    CargarLstCampos
    ConfiguraBtns lstCampos.ListCount, lstCamposSelec.ListCount
    LimpiarVector
    TablasCampos
    cabeceraGrid = ""
    Call WheelHook(frmReportesRRHH)
End Sub
Private Sub CargarLstCampos()
    Dim SQL As String
    Dim sql2 As String
    Dim Rs As MYSQL_RS
    Dim rs2 As MYSQL_RS
    Dim I As Integer
    lstCampos.Clear
    I = 0
    SQL = " SELECT * from empleado"
    SQL = "select a.codigo,a.codcargo,a.codTit,a.CodGrado, a.situacion,a.tipo," & _
          " a.modalidad,a.categoria,a.coddocide,a.personal,a.numdocide," & _
          " a.CarnetExt,a.pasaporte,a.TipBrevete,a.Brevete,a.nombre1," & _
          " a.nombre2,a.apepat,a.apemat,a.fec_nac,a.edad,a.sexo," & _
          " a.Gsangre,a.est_civil,a.num_hijos,a.direccion,a.distrito," & _
          " a.departamento,a.nacionalidad,a.fonofijo,a.fonomovil,a.mail," & _
          " a.Estatura,a.Peso,a.Calzado,a.Mameluco,a.foto,a.AsigFam," & _
          " a.jubilado,a.sctr,a.svl,a.codseg,a.numseg,a.codafp,a.numafp," & _
          " a.codbanco,a.tipcta_mn,a.numcta_mn,a.tipcta_me,a.numcta_me,  a.obs," & _
          " a.ctsbanco,a.ctsmon,a.ctsnumcta, b.division,a.fec_ingreso," & _
          " a.fec_cese,a.NomApo,a.DirecApo,a.FonoApo,a.MovilApo,a.codigohcm,B.DIVGAS,TRIM(B.CENCOS) AS CENCOS, " & _
          " b.codigo,b.anomes,b.codtipo,b.f_inicio,b.f_termino,b.mon_sueldo,b.sbasico," & _
          " B.bono,B.monto_bono,B.horlab,B.esttrabajo" & _
          " from empleado as a left join contrato as b" & _
          " on (a.codigo=b.codemp) where a.tipo not in (3,4) and b.codigo=(select max(codigo) from contrato where codemp=a.codigo)"
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    For I = 0 To Rs.FieldCount - 1
        sql2 = "select max(length(" & Rs.Fields(I).name & ")) AS TAMANIO FROM EMPLEADO"
        Set rs2 = oConexion.EjecutaSelectRS(sql2)
        With lstCampos
            .AddItem UCase(Rs.Fields(I).name) 'nombre del campo
            .List(I, 1) = rs2.Fields("TAMANIO") 'tamanio max del campo
            vCampos(I) = UCase(Rs.Fields(I).name)
        End With
    Next
    lstCampos.ListIndex = 0
    Set Rs = Nothing
    Set rs2 = Nothing
End Sub

Public Sub ConfigGrilla()
    Dim I As Integer, J As Integer
    Dim contador As Integer
    Dim grupos As Integer
    Dim SQL As String, Alias As String
    With flxReporte
        .Clear
        .Cols = lstCamposSelec.ListCount
        .Rows = 2
        .FixedCols = 0
        .FixedRows = 0
        .RowHeight(0) = 300
        .RowHeight(1) = 300
        J = 0
        For I = 0 To .Cols - 1
            Select Case LCase(lstCamposSelec.List(J, 0))
                Case "codigo": Alias = "a."
                Case "codcargo": Alias = "b."
                Case Else: Alias = ""
            End Select
            .TextMatrix(1, I) = Alias & LCase(lstCamposSelec.List(J, 0))
            .ColWidth(I) = IIf(Len(Trim(lstCamposSelec.List(J, 0))) * 100 > CDbl(IIf(lstCamposSelec.List(J, 1) = "", 0, lstCamposSelec.List(J, 1)) * 150), Len(Trim(lstCamposSelec.List(J, 0))) * 100, IIf(lstCamposSelec.List(J, 1) = "", 0, lstCamposSelec.List(J, 1)) * 150)
            J = J + 1
        Next
        .ScrollBars = flexScrollBarBoth
    End With
    For I = 0 To flxReporte.Cols - 1
        For J = 0 To 1
            flxReporte.row = J
            flxReporte.Col = I
            flxReporte.CellForeColor = &H80000002
            flxReporte.CellBackColor = &H8000000F
            If J = 0 Then
                    flxReporte.CellFontName = "Wingdings"
                    flxReporte.CellFontSize = 11
                    flxReporte.row = 0
                If cabeceraGrid = "" Then
                    flxReporte.Text = strUnChecked
                Else
                    flxReporte.Text = Left(Trim(cabeceraGrid), 1)
                    cabeceraGrid = Right(Trim(cabeceraGrid), Len(cabeceraGrid) - 1)
                End If
            End If
        Next
    Next I
End Sub

Private Function ConsultaGral() As String
    Dim SQL As String, str As String
    Dim I As Integer
    SQL = "Select "
    If chkActivos.Value = True Then
        str = " "
    Else
        str = " and a.situacion='1' "
    End If
    With flxReporte
        For I = 0 To .Cols - 1
            If I = .Cols - 1 Then SQL = SQL & .TextMatrix(1, I): Exit For
            SQL = SQL & Trim(.TextMatrix(1, I)) & ","
        Next
    End With
    ConsultaGral = SQL & " from empleado as a left join contrato as b on (a.codigo=b.codemp) where a.tipo not in (3,4) and b.codigo=(select max(codigo) from contrato where codemp=a.codigo) " & str & " order by a.apepat,a.apemat,a.nombre1,a.nombre2"
End Function

Public Sub LlenarGrilla(SQL As String)
    Dim I As Integer, J As Integer, M As Integer
    Dim Rs As MYSQL_RS
    Set Rs = oConexion.EjecutaSelectRS(SQL)
    J = 2
    Do While Not Rs.EOF
        flxReporte.Rows = flxReporte.Rows + 1
        For I = 0 To flxReporte.Cols - 1
            With flxReporte
                .TextMatrix(J, I) = IIf(IsNull(Rs.Fields(I)), "", Rs.Fields(I)) 'Left(rs.Fields(i) & Space(espacio), espacio)
            End With
        Next
        J = J + 1
        Rs.MoveNext
    Loop
    For I = 0 To flxReporte.Cols - 1
        If flxReporte.TextMatrix(0, I) = strChecked Then
            CambiarDatos I, tabla(I), Mid(Colum, 1, InStr(1, Colum, ",") - 1), Mid(Colum, InStr(1, Colum, ",") + 1, Len(Colum))
        End If
    Next
    lblNumRegistros = Rs.RecordCount
    Set Rs = Nothing
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single
    On Error Resume Next
    With flxReporte
        Lstep = .Height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 0
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        .TopRow = NewValue
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
End Sub

Function Columnas(Tbl As String) As String
    Dim SQL As String, cad As String
    'Dim RQ As MYSQL_RS, I As Integer
    
    Dim RQ As New ADODB.Recordset
    cad = ""
    SQL = "SHOW COLUMNS FROM " & Tbl
    Set RQ = ADO_LlenaRs(SQL)
    
    'Set RQ = oConexion.EjecutaSelectRS(SQL)
    If Not RQ.EOF() Then
        I = 1
        Do While Not RQ.EOF
            If I <= 2 Then
                cad = IIf(cad <> "", cad & ",", cad) & UCase(RQ.Fields("field"))
                I = I + 1
            Else
                Exit Do
            End If
            RQ.MoveNext
        Loop
    End If
    Columnas = cad
    Set RQ = Nothing
End Function

