VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmInterfazAsistencia 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cargar resgistro de otras Sedes"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmInterfazAsistencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtpatharchivo 
      Height          =   315
      Left            =   1110
      TabIndex        =   0
      Top             =   1050
      Width           =   4215
   End
   Begin Proyecto1.chameleonButton CmdExaminar 
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      Top             =   1050
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "..."
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
      MICON           =   "frmInterfazAsistencia.frx":2372
      PICN            =   "frmInterfazAsistencia.frx":238E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6420
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   """Text (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin Proyecto1.chameleonButton cmdProcesar 
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   1530
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmInterfazAsistencia.frx":4710
      PICN            =   "frmInterfazAsistencia.frx":472C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar pbProgreso 
      Height          =   285
      Left            =   30
      TabIndex        =   7
      Top             =   2220
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin Proyecto1.chameleonButton cmdGenerarA 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   "Generar archivos de actualización de BD para asistencia remota"
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
      MICON           =   "frmInterfazAsistencia.frx":5876
      PICN            =   "frmInterfazAsistencia.frx":5892
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblSede 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "AQUI NOMBRE DE SEDE A CARGAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   30
      TabIndex        =   6
      Top             =   1650
      Width           =   5265
   End
   Begin VB.Label lblMensaje 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Leyendo Archivo..."
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
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label lblInsertadas 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Insertadas..."
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
      Height          =   255
      Left            =   30
      TabIndex        =   3
      Top             =   1950
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Archivo:"
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
      Height          =   435
      Left            =   30
      TabIndex        =   2
      Top             =   960
      Width           =   1035
   End
End
Attribute VB_Name = "frmInterfazAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents a As cFile
Attribute a.VB_VarHelpID = -1
Dim separador As String
Dim cuenta As Long
Dim total As Long
Dim NArchivo As String

Private Sub CmdExaminar_Click()
    Dim FSO As Object
    Dim objFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
   
 On Error GoTo err
    separador = "|"
    Set a = Nothing
    Set a = New cFile
    With a
        .CloseFile
        CommonDialog1.ShowOpen
        .Filename = CommonDialog1.Filename
        NArchivo = CommonDialog1.Filename
        If InStr(1, NArchivo, "Cargado", vbBinaryCompare) <> 0 Then
           MsgBox "Este archivo ya ha sido cargado", vbOKOnly + vbInformation, "NOVADMIN"
           txtpatharchivo.Text = Empty
           Exit Sub
        End If
        Dim SQL As String
        Dim rsSede As MYSQL_RS
        
        SQL = "Select nombre,codigo from rh_estacionestrabajo where codigo='" & Mid(Right(NArchivo, 16), 2, 2) & "'"
        Set rsSede = oConexion.EjecutaSelectRS(SQL)
        If rsSede.RecordCount > 0 Then
            lblSede.Caption = rsSede.Fields("nombre")
            lblSede.tag = rsSede.Fields("codigo")
        End If
        rsSede.CloseRecordset
        Set rsSede = Nothing
        
        Set objFile = FSO.GetFile(NArchivo)

        FechaVoucher = Left(Trim(objFile.DateCreated), 10)
        
        Set objFile = Nothing
        Set FSO = Nothing
        txtpatharchivo.Text = NArchivo
        If separador <> "" Then
              .FieldSeparator = separador
        End If
        Me.MousePointer = vbHourglass
        lblMensaje.Visible = True
        lblInsertadas.Visible = False
        lblMensaje.Caption = "Leyendo archivo..."
        .Parse
        .CloseFile
        Me.MousePointer = vbNormal
        lblMensaje.Caption = str(.Lines.Count) & " líneas por procesar."
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdGenerarA_Click()
    Dim RQ As MYSQL_RS, SQL As String
    Dim NomArchivo As String, Cont As Integer
    Dim filetemp As Integer
    filetemp = FreeFile()

    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNFileMustExist Or _
            cdlOFNHideReadOnly Or _
            cdlOFNExplorer Or _
            cdlOFNLongNames
    
    On Error Resume Next
    CommonDialog1.Filename = "EMPLEADO_" & Replace(Format(Date, "yyyy/mm/dd"), "/", "")
    CommonDialog1.ShowSave
    If err.Number = cdlCancel Or err.Number <> 0 Then Exit Sub
    NomArchivo = CommonDialog1.Filename
    Cont = 0
    SQL = "select codigo,codcargo,situacion,categoria,coddocide,numdocide," & _
          " LEFT(concat(nombre1,repeat(' ',30 )),30) as nombre1," & _
          " LEFT(concat(nombre2,repeat(' ',30 )),30) as nombre2," & _
          " LEFT(concat(apepat,repeat(' ',30 )),30) as apepat, " & _
          " LEFT(concat(apemat,repeat(' ',30 )),30) as apemat" & _
          " From empleado where situacion='1' and tipo not in (3,4) and nacionalidad='01' order by codigo"
    
     Set RQ = oConexion.EjecutaSelectRS(SQL)
    
    If Not RQ.EOF Then
        On Error GoTo ErrorAbrir
        If NomArchivo = "" Then
            Set RQ = Nothing
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Open NomArchivo For Output As #filetemp
        
        Print #filetemp, "EMPLEADO"; d
        Do While Not RQ.EOF
        
           Print #filetemp, RQ.Fields("codigo") & "|" & _
                             RQ.Fields("codcargo") & "|" & _
                             RQ.Fields("situacion") & "|" & _
                             RQ.Fields("categoria") & "|" & _
                             RQ.Fields("coddocide") & "|" & _
                             RQ.Fields("numdocide") & "|" & _
                             RQ.Fields("nombre1") & "|" & _
                             RQ.Fields("nombre2") & "|" & _
                             RQ.Fields("apepat") & "|" & _
                             RQ.Fields("apemat") & "|"; d
            Cont = Cont + 1
            RQ.MoveNext
        Loop
    End If
    Close filetemp
    Set RQ = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Se generó el archivo satisfactoriamente." & vbNewLine & "Ruta: " & NomArchivo & ".txt" & vbNewLine & Cont & " Registro(s) generado(s)", vbInformation, "NOVADMIN"
    
Exit Sub
ErrorAbrir:
    MsgBox "Ocurrió un error al momento de generar el archivo de Interfaz, " & vbNewLine & _
           "consulte con el administrador del sistema", vbExclamation + vbOKOnly, "NOVPeru"
End Sub

Private Sub cmdProcesar_Click()
    If lblMensaje.Visible = False Then lblMensaje.Visible = True
    Me.MousePointer = vbHourglass
    CargarAsistencia
    Me.MousePointer = vbNormal
    MsgBox "PROCESO TERMINADO", vbOKOnly + vbInformation, "AVISO"
    lblMensaje.Caption = total & " Registros procesadas."
    lblInsertadas.Visible = True
    lblInsertadas.Caption = cuenta & " Insertadas"
    Name txtpatharchivo.Text As Left(txtpatharchivo.Text, Len(txtpatharchivo.Text) - 4) & "Cargado.txt"
    txtpatharchivo.Text = Empty
    cmdProcesar.Enabled = False
    cmdExaminar.SetFocus
    pbProgreso.Value = 0
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub txtpatharchivo_Change()
    If txtpatharchivo.Text <> Empty Then
        cmdProcesar.Enabled = True
    End If
End Sub

Private Sub CargarAsistencia()
    Dim Query As String
    Dim QueryInsMovi As String
    Dim I As Long
    Dim intCodigo As Integer
    Dim RQ As MYSQL_RS
    cuenta = 0
    total = a.Lines.Count
    For I = 1 To a.Lines.Count
        QueryInsMovi = "SELECT * from rh_entsalempleado where emp = '" & a.Lines(I).Fields(2).Value & "' " & _
                       " and fecha = '" & a.Lines(I).Fields(3).Value & "' and tipo = '" & a.Lines(I).Fields(5).Value & "'"
        Set RQ = oConexion.EjecutaSelectRS(QueryInsMovi)
        If RQ.EOF() Then
            Query = "Insert into rh_entsalempleado (sede,emp,fecha,hor,tipo,envio,tiposede) values ('" & _
                    a.Lines(I).Fields(1).Value & "','" & _
                    a.Lines(I).Fields(2).Value & "','" & _
                    a.Lines(I).Fields(3).Value & "','" & _
                    a.Lines(I).Fields(4).Value & "','" & _
                    a.Lines(I).Fields(5).Value & "','" & _
                    a.Lines(I).Fields(6).Value & "','" & _
                    a.Lines(I).Fields(7).Value & "')"
            oConexionMYSQL.Execute Query
        Else
            Query = "update rh_entsalempleado set sede = '" & a.Lines(I).Fields(1).Value & "',hor='" & a.Lines(I).Fields(4).Value & "',envio = '" & a.Lines(I).Fields(6).Value & "',tiposede='" & a.Lines(I).Fields(7).Value & "' " & _
                    "where emp = '" & a.Lines(I).Fields(2).Value & "' and fecha = '" & a.Lines(I).Fields(3).Value & "' and tipo ='" & a.Lines(I).Fields(5).Value & "'"
            oConexionMYSQL.Execute Query
        End If
        cuenta = cuenta + 1
        pbProgreso.Value = I * 100 / a.Lines.Count
    Next I
    Set RQ = Nothing
End Sub
