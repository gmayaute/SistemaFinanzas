VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAsignarEstados 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de usuarios a estado del documento"
   ClientHeight    =   5820
   ClientLeft      =   5595
   ClientTop       =   7080
   ClientWidth     =   8610
   Icon            =   "frmPermisos.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8610
   Begin VB.Frame Frame2 
      BackColor       =   &H009F5539&
      Height          =   5820
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   8565
      Begin Proyecto1.chameleonButton cmdGrabar 
         Height          =   375
         Left            =   7980
         TabIndex        =   6
         Top             =   300
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
         MICON           =   "frmPermisos.frx":0A02
         PICN            =   "frmPermisos.frx":0A1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5535
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   9763
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   1058
         BackColor       =   10442041
         TabCaption(0)   =   "Estados"
         TabPicture(0)   =   "frmPermisos.frx":0E60
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstPadre"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lstHijo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lstNieto"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TreeEstados"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin MSComctlLib.TreeView TreeEstados 
            Height          =   4515
            Left            =   120
            TabIndex        =   2
            Top             =   780
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   7964
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin MSForms.ListBox lstNieto 
            Height          =   1095
            Left            =   4800
            TabIndex        =   5
            Top             =   3480
            Visible         =   0   'False
            Width           =   2295
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "4048;1931"
            MatchEntry      =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox lstHijo 
            Height          =   1215
            Left            =   4800
            TabIndex        =   4
            Top             =   2160
            Visible         =   0   'False
            Width           =   2295
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "4048;2143"
            MatchEntry      =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ListBox lstPadre 
            Height          =   1215
            Left            =   4830
            TabIndex        =   3
            Top             =   1500
            Visible         =   0   'False
            Width           =   2295
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "4048;2143"
            MatchEntry      =   0
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin Proyecto1.chameleonButton CmdVistaPreliminar 
         Height          =   375
         Left            =   7980
         TabIndex        =   7
         Top             =   750
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
         MICON           =   "frmPermisos.frx":115E
         PICN            =   "frmPermisos.frx":117A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdSalir 
         Height          =   375
         Left            =   7980
         TabIndex        =   8
         Top             =   5280
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
         MICON           =   "frmPermisos.frx":16BC
         PICN            =   "frmPermisos.frx":16D8
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
End
Attribute VB_Name = "frmAsignarEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Formulario As TIPO_FORMULARIO
Public Property Let pFormulario(valor As TIPO_FORMULARIO)
    m_Formulario = valor
End Property
Private Sub cmdGrabar_Click()
    mdiInicio.MousePointer = vbHourglass
    GrabarAsigancionEstados
    mdiInicio.MousePointer = vbNormal
End Sub



Private Sub GrabarAsigancionEstados()
    On Error GoTo Errdel
    Dim I As Integer
    Dim Estado As String
    Dim SQL As String
    Select Case m_Formulario
        Case 58: SQL = "delete from docsusuario"
        Case 59: SQL = "delete from estado_usu "
    End Select
    oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.Eliminar, False
    For I = 2 To TreeEstados.Nodes.Count
        If Left(TreeEstados.Nodes.Item(I).Key, 2) = "NE" Then
            Estado = Right(TreeEstados.Nodes.Item(I).Key, 2)
        Else
            Select Case m_Formulario
                Case 58
                    If Estado = "LS" Or Estado = "AS" Or Estado = "PL" Or Estado = "SS" Then
                        SQL = "INSERT INTO docsusuario (coddoc,USUARIO,PERMISO) " & _
                              " VALUES('" & Estado & "' ,'" & _
                              TreeEstados.Nodes.Item(I).Text & "'," & _
                              IIf(TreeEstados.Nodes(I).Checked = True, 1, 0) & ")"
                    End If
                Case 59
                        SQL = "INSERT INTO ESTADO_USU (COD_ESTADO,USUARIO_ID,PERMISO) " & _
                              " VALUES('" & Estado & "' ,'" & _
                              TreeEstados.Nodes.Item(I).Text & "'," & _
                              IIf(TreeEstados.Nodes(I).Checked = True, 1, 0) & ")"
            End Select
            oConexion.EjecutaInsertUpdateDelete SQL, TIPO_QUERY.insertar, False
        End If
    Next I
    MsgBox "Permisos Asignados", vbInformation, "NOVPeru"
Exit Sub
Errdel:
    MsgBox "Ha ocurrido un error al momento de Grabar" & Chr(13) & err.Description, vbCritical, "Error de datos"
    ADOConexion.RollbackTrans
End Sub
Private Sub LimpiaArbol(Tipo As TipoArbol)
    Dim I As Integer
    For I = 1 To Me.TreeEstados.Nodes.Count
        Me.TreeEstados.Nodes(I).Checked = False
    Next I
End Sub
Private Sub EliminarNodos(ByRef Arbol As TreeView)
    Dim I As Integer
    If Arbol.Nodes.Count > 0 Then
        For I = Arbol.Nodes.Count To 1 Step -1
            Arbol.Nodes.Remove (I)
        Next I
    End If
End Sub
Private Sub LlenaArbol(ByRef Arbol As TreeView)
    Dim Rs As MYSQL_RS
    Dim SQL As String
    Dim Cant As Integer
    Dim I As Integer
    Dim nNode  As Node
    Dim Estado As String, usuario As String
    Dim Nivel2 As Integer
    Dim Nivel3 As Integer
    Dim Nivel4 As Integer
    EliminarNodos Arbol
    Select Case m_Formulario
        Case 58
            SSTab1.Caption = "Documentos"
            SQL = "select b.descrip as descripcion,a.coddoc as cod_estado,a.usuario as usuario_id,a.permiso " & _
                  "from cndocum as b left join docsusuario as a" & _
                  " on a.coddoc=b.coddoc where protegido = 'S' order by b.descrip,a.usuario"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            Set nNode = Arbol.Nodes.Add(Key:="a", Text:="Documentos")
        Case 59
            SSTab1.Caption = "Estados"
            SQL = "select b.descripcion,a.cod_estado,a.usuario_id,a.permiso from estado_usu as a left join doc_estado as b" & _
                  " on a.cod_estado=b.cod_estado order by b.descripcion,a.usuario_id"
            Set Rs = oConexion.EjecutaSelectRS(SQL)
            Set nNode = Arbol.Nodes.Add(Key:="a", Text:="Estados")
    End Select
    I = 1
    Estado = ""
    Do While Not (Rs.EOF)
        If Estado <> Rs.Fields("cod_estado") Then
            Estado = Rs.Fields("cod_estado")
            Set nNode = Arbol.Nodes.Add("a", tvwChild, Key:="NEstados" & Estado, Text:=CE(Rs.Fields("descripcion")))
            I = I + 1
            Set nNode = Arbol.Nodes.Add("NEstados" & Estado, tvwChild, Key:="NUsuarios" & I, Text:=CE(Rs.Fields("usuario_id")))
        Else
            Set nNode = Arbol.Nodes.Add("NEstados" & Estado, tvwChild, Key:="NUsuarios" & I, Text:=CE(IIf(IsNull(Rs.Fields("usuario_id")), "", Rs.Fields("usuario_id"))))
        End If
        I = I + 1
        If Left(Arbol.Nodes(I).Key, 2) = "NU" Then
            Arbol.Nodes(I).Checked = IIf(CE(Rs.Fields("permiso")) = 0, False, True)
        End If
        Rs.MoveNext
    Loop
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LlenaArbol TreeEstados
End Sub
Private Sub Treeestados_NodeCheck_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim oTree As clsTreeView
    Set oTree = New clsTreeView
    If Node.Children Then oTree.NodeChildrenCheck Node
    If Not Node.Parent Is Nothing And Node.Checked Then
        oTree.NodeParentsCheck Node
    Else
        If Not Node.Parent Is Nothing And Not Node.Checked Then oTree.NodeSelectedCheck Node.Parent
    End If
    Set oTree = Nothing
End Sub
