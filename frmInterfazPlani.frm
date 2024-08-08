VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{15A4AECE-7618-4F12-AD87-DA1E11EABB34}#1.0#0"; "botom.ocx"
Begin VB.Form frmInterfazPlani 
   BackColor       =   &H009F5539&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interfaz Telebanking"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   Icon            =   "frmInterfazPlani.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTc 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   1035
   End
   Begin MSMask.MaskEdBox DtpFecha 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      ToolTipText     =   "Fecha_Pago"
      Top             =   1470
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      _Version        =   393216
      ForeColor       =   128
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin Proyecto1.chameleonButton chBtnSalir 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3300
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   1920
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
      MICON           =   "frmInterfazPlani.frx":030A
      PICN            =   "frmInterfazPlani.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton chBtnReporte 
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      ToolTipText     =   "Ver Reporte"
      Top             =   1920
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
      MICON           =   "frmInterfazPlani.frx":06EC
      PICN            =   "frmInterfazPlani.frx":0708
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Proyecto1.chameleonButton BtnGenerar 
      Height          =   375
      Left            =   90
      TabIndex        =   8
      ToolTipText     =   "Eliminar"
      Top             =   1920
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Generar Tw"
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
      MICON           =   "frmInterfazPlani.frx":0C4A
      PICN            =   "frmInterfazPlani.frx":0C66
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T/C"
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
      Height          =   285
      Left            =   2130
      TabIndex        =   11
      Top             =   1470
      Width           =   465
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
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
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   10
      Top             =   510
      Width           =   765
   End
   Begin MSForms.ComboBox cboMon 
      Height          =   315
      Left            =   900
      TabIndex        =   9
      Top             =   480
      Width           =   2775
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4895;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planilla"
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
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
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
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   1470
      Width           =   765
   End
   Begin VB.Label Lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes"
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
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   90
      Width           =   765
   End
   Begin MSForms.ComboBox cboMes 
      Height          =   315
      Left            =   900
      TabIndex        =   2
      Top             =   90
      Width           =   2775
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4895;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cboProceso 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "4895;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmInterfazPlani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim SQL As String
Dim res As Integer
Dim Rs As MYSQL_RS

Private Sub BtnGenerar_Click()
    Dim tc As Double
    If cboMon.ListIndex > 0 Then
        If cboMon.ListIndex = 2 Then
            If Not (IsNumeric(txtTc)) Then
                MsgBox "Ingrese correctamente el tipo de cambio para esta planilla", vbOKOnly + vbExclamation, "NOVPeru"
                txtTc.SetFocus
                Exit Sub
            End If
            tc = CDbl(txtTc)
        Else
            tc = 1
        End If
        GeneraTxt cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon, 1), Format(DtpFecha.Text, "yyyy/mm/dd"), tc
    End If
End Sub

Private Sub cboMon_Click()
    If cboMon.ListCount > 1 Then
        If cboMon.ListIndex = 2 Then
            txtTc.Enabled = True
        Else
            txtTc.Enabled = False
        End If
    End If
End Sub

Private Sub chBtnReporte_Click()
    Dim tc As String
    If cboMon.ListIndex > 0 Then
        If cboMon.ListIndex = 2 Then
            If Not (IsNumeric(txtTc)) Then
                MsgBox "Ingrese correctamente el tipo de cambio para esta planilla", vbOKOnly + vbExclamation, "NOVPeru"
                txtTc.SetFocus
                Exit Sub
            End If
            tc = FormatNumber(txtTc, 3)
        Else
            tc = "1"
        End If
         Set oReporte = New clsReporte
        oReporte.empresa = strNombreEmpresa
        oReporte.TipoCambio = tc
        oReporte.Reporte = "Rep_PagoTw.rpt"
        oReporte.sp_Rep_PlanillaTW cboMes.List(cboMes.ListIndex, 1) & cboMes.List(cboMes.ListIndex, 2), cboProceso.List(cboProceso.ListIndex, 2), Left(cboMon, 1)
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    LlenarMesP cboMes
    LlenarMonedaP cboMon
    LlenarProcesos cboProceso
    DtpFecha.Text = Format(CStr(Date), "dd/mm/yyyy")
End Sub

Public Function GeneraTxt(AnoMes As String, plani As String, mon As String, fecha As String, TIPC As Double) As Boolean
        Dim Cont As Integer, str As String
        Dim NomArchivo As String, NomEmp As String, ido As String, cci As String
        Dim neto As Double, total As Double, tc As Double
        Dim oficina As String, cuenta As String, TipCta As String, codigo As String
        Dim filetemp As Integer
        filetemp = FreeFile()
        Cont = 0
        GeneraTxt = False
        ido = GenerarOrden(AnoMes)
        
        Select Case plani
            Case "1", "4":
                str = " if(a.unidad='D',a.cant*(a.sbasico/30),a.cant) , 0 )"
                Concepto = "PAGO PLANILLA"
            Case "2":
                str = "a.cant*.80* (a.sbasico/30) , '' )"
                Concepto = "PAGO PLANILLA"
            Case "5":
                str = " if(a.unidad='D',a.cant*(a.sbasico/180),a.cant) , 0 )"
                Concepto = "PAGO GRATIF."
            Case "6":
                str = " if(a.unidad='D',360*(a.sbasico/360) , a.cant ),0)"
                Concepto = "DEPOSITO CTS"
        End Select
        
        If mon = "N" Then
            tc = 1
            NomArchivo = strRutaTelewiese & "PagosPL\" & Right(AnoMes, 2) & "\" & Right(ido, 4) & AnoMes & "PLMN"
        Else
            tc = TIPC
            NomArchivo = strRutaTelewiese & "PagosPL\" & Right(AnoMes, 2) & "\" & Right(ido, 4) & AnoMes & "PLME"
        End If
        
        If plani = "6" Then
            SQL = "SELECT a.anomes,a.emp,'3' as tipcta_mn,c.ctsnumcta as numcta_mn,'3' as tipcta_me,c.ctsnumcta as numcta_me,CONCAT_WS(' ',C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2) as nombre," & _
                  " a.tipo  as proceso,c.ctsmon as moneda," & _
                  "(select cant from pl_tareo where ANOMES='" & AnoMes & "' AND TIPO='" & plani & "'" & _
                  " AND rub='121' and unidad='D' and emp=a.emp) *" & _
                  " (sum(IF( a.rub ='121', " & str & ") + " & _
                  " sum( IF( a.rub ='115', a.cant, 0 )) + sum( IF( a.rub ='1001', a.cant, 0 )) + sum( IF( a.rub ='1002', a.cant, 0 )) + sum( IF( a.rub ='1007', a.cant, 0 )) + sum( IF( a.rub ='1005', a.cant, 0 )) + sum( IF( a.rub ='1008', a.cant, 0 )) + sum( IF( a.rub ='201', a.cant, 0 ))  + sum( IF( a.rub ='105', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='401', a.cant, 0 )))/360 AS ingresos, sum( IF( a.rub ='703', a.cant, 0 )) as descuentos,a.fecha" & _
                  " FROM pl_tareo as a LEFT JOIN" & _
                  " empleado  AS c ON (a.EMP=c.CODIGO)" & _
                  " WHERE c.ctsbanco='09' and a.ANOMES='" & AnoMes & "' AND a.TIPO='" & plani & "'" & _
                  " and c.ctsmon='" & mon & "' " & _
                  " GROUP BY a.anomes,a.emp,a.tipo,c.ctsmon,a.fecha ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
        Else
            SQL = "SELECT a.anomes,a.emp,c.tipcta_mn,c.numcta_mn,c.tipcta_me,c.numcta_me,CONCAT(C.APEPAT,' ',C.APEMAT,' ',C.NOMBRE1,' ',C.NOMBRE2) AS NOMBRE," & _
                  " a.tipo as proceso,a.moneda,sum( IF( a.rub ='121', " & str & ") + " & _
                  " sum( IF( a.rub ='1001', a.cant, 0 )) + sum( IF( a.rub ='1005', a.cant, 0 )) + sum( IF( a.rub ='1008', a.cant, 0 )) +  sum( IF( a.rub ='1006', a.cant, 0 ))  + sum( IF( a.rub ='1007', a.cant, 0 ))  + sum( IF( a.rub ='1002', a.cant, 0 )) + sum( IF( a.rub ='1003', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='102', a.cant, 0 )) + sum( IF( a.rub ='103', a.cant, 0 )) +  sum( IF( a.rub ='105', a.cant, 0 )) +  sum( IF( a.rub ='107', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='115', a.cant, 0 )) + sum( IF( a.rub ='116', a.cant, 0 )) + sum( IF( a.rub ='118', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='117', " & str & " ) + " & _
                  " sum( IF( a.rub ='504', a.cant, 0 )) + sum( IF( a.rub ='201', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='202', a.cant, 0 )) + sum( IF( a.rub ='313', a.cant, 0 )) + sum( IF( a.rub ='407', a.cant, 0 )) + sum( IF( a.rub ='910', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='911', a.cant, 0 )) + sum( IF( a.rub ='915', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='904', a.cant, 0 )) + sum( IF( a.rub ='909', a.cant, 0 )) + sum( IF( a.rub ='916', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='401', a.cant, 0 )) + sum( IF( a.rub ='403', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='404', a.cant, 0 )) + sum( IF( a.rub ='902', a.cant, 0 )) AS ingresos," & _
                  " sum( IF( a.rub ='601', a.cant, 0 )) + sum( IF( a.rub ='604', a.cant, 0 )) + sum( IF( a.rub ='605', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='607', a.cant, 0 )) + sum( IF( a.rub ='608', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='606', a.cant, 0 )) + sum( IF( a.rub ='611', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='701', a.cant, 0 )) + sum( IF( a.rub ='703', a.cant, 0 )) + sum( IF( a.rub ='704', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='705', a.cant, 0 )) + sum( IF( a.rub ='706', a.cant, 0 )) + sum( IF( a.rub ='707', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='708', a.cant, 0 )) + sum( IF( a.rub ='709', a.cant, 0 )) + " & _
                  " sum( IF( a.rub ='710', a.cant, 0 )) + sum( IF( a.rub ='711', a.cant, 0 )) + sum( IF( a.rub ='712', a.cant, 0 )) + sum( IF( a.rub ='713', a.cant, 0 )) + sum( IF( a.rub ='714', a.cant, 0 )) as Descuentos, " & _
                  " a.fecha FROM pl_tareo as a left join empleado  AS c on (a.emp=c.codigo)" & _
                  " where A.ANOMES='" & AnoMes & "' and a.tipo='" & plani & "'" & _
                  " AND a.moneda='" & mon & "' " & _
                  " GROUP BY a.anomes,a.emp,a.tipo,a.moneda,a.fecha ORDER BY C.APEPAT,C.APEMAT,C.NOMBRE1,C.NOMBRE2"
        End If
        
        Set Rs = oConexion.EjecutaSelectRS(SQL)
        
        If Not Rs.EOF Then
                Rs.MoveFirst
                With Rs
                    On Error GoTo ErrorAbrir
                    Open NomArchivo & ".txt" For Output As #filetemp
                                    
                    Do While Not .EOF
                        
                        TipCta = IIf(plani = "6", "3", .Fields("tipcta_mn"))
                        
                        NomEmp = Left(Trim(.Fields("nombre")), 46)
                        Dni = DevNumDoc(Trim(.Fields("emp")), "3")
                        codigo = Right(Trim(.Fields("emp")), 8)
                        neto = Round((CDbl(.Fields("INGRESOS")) - CDbl(.Fields("DESCUENTOS"))) / tc, 2)
                        total = total + neto
                        If mon = "N" Then
                            If TipCta = "9" Then
                                TipCta = " "
                                cci = Right(Trim(.Fields("numcta_mn")), 20)
                                oficina = ""
                                cuenta = "4"
                            Else
                                cci = ""
                                oficina = Left(Trim(.Fields("numcta_mn")), 3)
                                cuenta = Right(Trim(.Fields("numcta_mn")), 7)
                            End If
                        Else
                            cci = ""
                            oficina = Left(Trim(.Fields("numcta_me")), 3)
                            cuenta = Right(Trim(.Fields("numcta_me")), 7)
                        End If
                        If neto > 0 Then
                            If cuenta <> "4" Then
                             Print #filetemp, Space(15) & _
                             Left(Trim(NomEmp) & Space(46), 46) & _
                             Left(Concepto & Space(14), 14) & _
                             Replace(fecha, "/", "") & _
                             Space(22) & Right(Space(11) & Trim(Replace(Replace(FormatNumber(neto, 2), ".", ""), ",", "")), 11) & _
                             Space(28) & TipCta & _
                             Right(Space(3) & Trim(oficina), 3) & Right(Space(7) & Trim(cuenta), 7) & Space(16) & Left(Trim(Dni), 8) & Space(1) & Right(Trim(codigo), 8) & Space(11) & Right(Trim(cci), 29)
                            Else
'                             Print #filetemp, Space(15) & _
'                             Left(Trim(NomEmp) & Space(46), 46) & _
'                             Left(Concepto & Space(14), 14) & _
'                             Replace(fecha, "/", "") & _
'                             Space(22) & Right(Space(11) & Trim(Replace(Replace(FormatNumber(neto, 2), ".", ""), ",", "")), 11) & _
'                             Space(28) & TipCta & _
'                             Trim(cuenta) & Space(16) & Left(Trim(Dni), 8) & Space(1) & Right(Trim(codigo), 8) & Space(11) & Right(Trim(cci), 29)

                              Print #filetemp, Space(15) & _
                              Left(Trim(NomEmp) & Space(46), 46) & _
                              Left(Concepto & Space(14), 14) & _
                              Replace(fecha, "/", "") & _
                              Space(22) & Right(Space(11) & Trim(Replace(Replace(FormatNumber(neto, 2), ".", ""), ",", "")), 11) & _
                              Space(27) & TipCta & _
                              Trim(cuenta) & Space(26) & Left(Trim(Dni), 8) & Space(1) & Right(Trim(codigo), 8) & Space(11) & Right(Trim(cci), 29)
                            End If
                            
                            
                        End If
                        Cont = Cont + 1
                        .MoveNext
                    Loop
                    .CloseRecordset
                    GeneraTxt = True
                    MsgBox "La Transferencia Telewiese se realizó satisfactoriamente." & vbNewLine & Cont & " Registro(s) generado(s)", vbInformation, gsNomSW
                End With
                Close filetemp
                Set Rs = Nothing
        End If
      SQL = "INSERT INTO mov_telewiese (ITEM, IDENTIFICADOR, FECHA, FEC_EMI, AUXILIAR, CODIGO, MONEDA, CTACTE, TIPDOC, FOLIOREF,DOCUMENTO,IMPORTE,IMPORTEEQU, CTAAUX,OFICINA,TIPOPAGO,OBS,ESTADO) VALUES" & _
            " (1,'" & ido & "','" & fecha & "','" & fecha & "','6','00000000068','" & mon & "','00000000001','PL','" & ido & "','PLANI-0" & Replace(fecha, "/", "") & "'," & Round(total, 2) & ",0,'0000000','000','2','','TR')"
      oConexionMYSQL.Execute SQL
      
      SQL = " update pl_planiproc SET tw='" & ido & "',fecha ='" & fecha & "', tipcam = " & tc & " where " & _
            " anomes='" & AnoMes & "' and proceso='" & plani & "' and mon='" & mon & "'"
      oConexionMYSQL.Execute SQL
      
      SQL = "INSERT INTO pl_planiproc (anomes, proceso, mon, contab,fecha,tw,tipcam) VALUES" & _
            " ('" & AnoMes & "','" & plani & "','" & mon & "','N','" & fecha & "','" & ido & "'," & tc & ")"
      oConexionMYSQL.Execute SQL
      
      Exit Function
ErrorAbrir:
    If err.Number = 76 Then
        res = MsgBox("No se encuentra la ruta: " & strRutaTelewiese & NombreMes(strMesSistema, False) & vbNewLine & vbNewLine & ",Desea crearla?", vbQuestion + vbYesNo, gsNomSW)
        If res = 6 Then
            MkDir (strRutaTelewiese & "PagosPL\" & Right(AnoMes, 2) & "\")
            Open NomArchivo & ".txt" For Output As #filetemp
            Resume Next
        Else
            GeneraTxt = False
        End If
    End If
End Function
