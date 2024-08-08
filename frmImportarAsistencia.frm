VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "Botom.ocx"
Begin VB.Form frmImportAsistencia 
   BackColor       =   &H009F5539&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Asistencia"
   ClientHeight    =   7455
   ClientLeft      =   3510
   ClientTop       =   3765
   ClientWidth     =   10650
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10650
   Begin VB.Frame Frame1 
      BackColor       =   &H009F5539&
      Height          =   7395
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   10305
      Begin Proyecto1.chameleonButton cmdExaminar 
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   660
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
         BTYPE           =   14
         TX              =   "&Examinar"
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
         MICON           =   "frmImportarAsistencia.frx":0000
         PICN            =   "frmImportarAsistencia.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar pbMapeo 
         Height          =   225
         Left            =   1590
         TabIndex        =   2
         Top             =   1410
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshAsistencia 
         Height          =   4800
         Left            =   150
         TabIndex        =   5
         Top             =   1770
         Width           =   10035
         _ExtentX        =   17701
         _ExtentY        =   8467
         _Version        =   393216
         BackColorBkg    =   12632256
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Proyecto1.chameleonButton cmdMapear 
         Height          =   405
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   714
         BTYPE           =   14
         TX              =   "&Importar"
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
         MICON           =   "frmImportarAsistencia.frx":05B6
         PICN            =   "frmImportarAsistencia.frx":05D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Proyecto1.chameleonButton cmdProc 
         Height          =   465
         Left            =   3420
         TabIndex        =   7
         Top             =   6750
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   820
         BTYPE           =   14
         TX              =   "&Procesar Asistencia del Mes"
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
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmImportarAsistencia.frx":0B6C
         PICN            =   "frmImportarAsistencia.frx":0B88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSForms.Label lblarchivo 
         Height          =   315
         Left            =   1620
         TabIndex        =   4
         Top             =   720
         Width           =   8235
         ForeColor       =   8421631
         BackColor       =   10442041
         Caption         =   "Nombre Archivo"
         Size            =   "14526;556"
         BorderStyle     =   1
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label13 
         BackColor       =   &H009F5539&
         Caption         =   "Seleccione archivo Excel con el formato adecuado para la asistencia del presente mes"
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
         Left            =   270
         TabIndex        =   3
         Top             =   270
         Width           =   8895
      End
   End
End
Attribute VB_Name = "frmImportAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMapear_Click()
  On Error GoTo mal
    Dim badgeno As String, badgename As String, empno As String, dia As String
    Dim ingreso As String, Salida As String, weekday As String
    Dim QueryOrden As String
    
    
    Excel.Application.Workbooks.Open Filename:=lblarchivo
    Excel.Application.Visible = False
    Hoja = "ASISTENCIA"
    UltimaCelda = ActiveCell.SpecialCells(xlCellTypeLastCell).Address
    fila = CDbl(Right(UltimaCelda, Len(UltimaCelda) - InStr(2, UltimaCelda, "$", vbBinaryCompare)))
    I = 1
    
    pbMapeo.Min = 1
    pbMapeo.Max = fila + 1
    
    'Limpiamos la tabla temporal "ASISG2"
    QueryOrden = "delete from asisg2"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Eliminar, False
    
    'Importación del excel
    Do While I <= fila
        err = False
        badgeno = Trim(CStr(Worksheets(Hoja).Cells(I, 1).Value))
        badgename = Trim(CStr(Worksheets(Hoja).Cells(I, 2).Value))
        empno = Trim(CStr(Worksheets(Hoja).Cells(I, 3).Value))
        dia = Trim(CStr(Worksheets(Hoja).Cells(I, 4).Value))
        ingreso = Trim(CStr(Worksheets(Hoja).Cells(I, 5).Value))
        Salida = Trim(CStr(Worksheets(Hoja).Cells(I, 8).Value))
        weekday = Trim(CStr(Worksheets(Hoja).Cells(I, 9).Value))
        
        QueryOrden = "insert into `asisg2`(`badgeno`,`badgename`,`empno`,`date`,`ing`,`out`,`weekdate`)" & _
              "values ('" & badgeno & "','" & badgename & "','" & empno & "','" & dia & "','" & ingreso & "','" & Salida & "','" & weekday & "');"
        
        oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.insertar, False
        I = I + 1
        pbMapeo.Value = I
        pbMapeo.Refresh
    Loop
    Excel.Application.Workbooks.Close
    Excel.Application.Quit
    MsgBox "Proceso terminado con éxito", vbOKOnly, "MAPEO"
    pbMapeo.Value = 1
    pbMapeo.Refresh
    
    'Primer paso borrar lo de las otras categorias
    QueryOrden = "delete from asisg2 Where empno in (select codigo from empleado where categoria in ('00','01') )"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Eliminar, False
    
    QueryOrden = "delete from asisg2 Where weekdate in ('Sábado','Domingo')"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Eliminar, False
    
    'Preparamos la tabla asisg2 para los valores nulos
    QueryOrden = "Update asisg2 SET ing = concat(ing,':00') Where Length(Ing) = 5"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Modificar, False

    QueryOrden = "Update asisg2 SET ing=ifnull(ADDTIME(ing,concat('00:','00',':',round(MOD((rand() * 3600),55)))),'08:34:43')Where Length(Ing) = 8"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Modificar, False
    
    QueryOrden = "Update asisg2 SET ing=ifnull(ADDTIME('08:30:00',concat('00:', round(MOD((rand() * 10),60)),':',round(MOD((rand() * 3600),55)))),'08:31:43') Where Ing Is Null"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Modificar, False
    
    
    'Actualizamos con el formato de fecha respecto a la tabla entsalempleado
    QueryOrden = "update asisg2 SET date=concat(right(date,4),'/',substring(date,4,2),'/',left(date,2))"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Modificar, False
    
    'Eliminamos Personal Externo
    QueryOrden = "delete  from asisg2 where badgeno in ('8816259','42024772','15738428')"
    oConexion.EjecutaInsertUpdateDelete QueryOrden, TIPO_QUERY.Eliminar, False
    
    
    LlenarMshAsistencia
    
    
mal:
    MsgBox "Hubo un error en la importación, revise el excel de asistencia", vbOKOnly, "ERROR"
    pbMapeo.Value = 0
    pbMapeo.Refresh
    Excel.Application.Workbooks.Close
    Excel.Application.Quit
    
End Sub


Sub LlenarMshAsistencia()
    Dim RsCias As MYSQL_RS
    Set RsCias = New MYSQL_RS
    Dim SqlCias  As String
    
    SqlCias = "select badgeno,badgename,empno,date,ing,`out`,weekdate from asisg2"
    Set RsCias = oConexion.EjecutaSelectRS(SqlCias)
    If RsCias.EOF And RsCias.BOF Then
        MsgBox "No se ha cargado correctamente la Asistencia, Consulte al Administrador", vbInformation + vbOKOnly, "Asistencia del Mes"
    Else
        ConfigMshAsistencia
        Dim I As Integer
        With MshAsistencia
            Do While Not RsCias.EOF
                .TextMatrix(.Rows - 1, 1) = CE(RsCias.Fields(0))
                .TextMatrix(.Rows - 1, 2) = CE(RsCias.Fields(1))
                .TextMatrix(.Rows - 1, 3) = CE(RsCias.Fields(2))
                
                If (CE(RsCias.Fields(2)) <> "") = False Then
                        .Col = 3
                        .CellForeColor = vbRed
                End If
                
                .TextMatrix(.Rows - 1, 4) = CE(RsCias.Fields(3))
                .TextMatrix(.Rows - 1, 5) = CE(RsCias.Fields(4))
                .TextMatrix(.Rows - 1, 6) = CE(RsCias.Fields(5))
                .TextMatrix(.Rows - 1, 7) = CE(RsCias.Fields(6))
                
                RsCias.MoveNext
                .Rows = .Rows + 1
            Loop
            .Rows = .Rows - 1
        End With
    End If
    Set RsCias = Nothing
End Sub

Sub ConfigMshAsistencia()
    With MshAsistencia
        .Cols = 7
        .FixedCols = 1
        .Rows = 2
        .Clear
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 700
        .ColWidth(4) = 700
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        
        
        .TextMatrix(0, 1) = "Codigo"
        .TextMatrix(0, 2) = "Nombre"
        .TextMatrix(0, 3) = "Codigo Empleado"
        .TextMatrix(0, 4) = "Fecha"
        .TextMatrix(0, 5) = "Hora Ingreso"
        .TextMatrix(0, 6) = "Hora Salida"
        .TextMatrix(0, 7) = "Dia"
        
    End With
End Sub


Private Sub cmdProc_Click()
  Dim QueryAsis As String
   QueryAsis = "Insert into rh_entsalempleado (sede,emp,fecha,hor,tipo,envio,TipoSede) select (select EstTrabajo from contrato where codemp=a.empno order by codigo desc limit 1),a.empno,a.date,a.ing,'E','PRO','O' from asisg2 as a where empno not in (select emp from rh_entsalempleado  as r where r.emp=a.empno and r.fecha=a.date and r.hor=a.ing and r.tipo='E')"
   oConexion.EjecutaInsertUpdateDelete QueryAsis, TIPO_QUERY.insertar, False
 
   QueryAsis = "Insert into rh_entsalempleado (sede,emp,fecha,hor,tipo,envio,TipoSede) select sede,emp,fecha,ifnull(ADDTIME('17:30:00',concat('00:', round(MOD((rand() * 10),60)),':',round(MOD((rand() * 3600),55)))),'17:34:43'),'S','I3','O' from rh_entsalempleado where envio='PRO'"
   oConexion.EjecutaInsertUpdateDelete QueryAsis, TIPO_QUERY.insertar, False
   
   MsgBox "Se proceso correctamente la Asistencia del Mes", vbOKOnly, "AVISO"
 
End Sub
