VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRViewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportPreview 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   7515
   ClientTop       =   10500
   ClientWidth     =   10665
   Icon            =   "frmReportePreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   10665
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      Picture         =   "frmReportePreview.frx":0442
      ScaleHeight     =   285
      ScaleWidth      =   330
      TabIndex        =   1
      ToolTipText     =   " Configurar Impresora "
      Top             =   45
      Width           =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5565
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   1965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10650
      lastProp        =   500
      _cx             =   5080
      _cy             =   5080
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReportPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oReporte As New CRAXDRT.Report
Dim ArchivoXLS  As String
Private m_Rs As CRAXDRT.Report

Public Property Let Recordset(ByRef valor As CRAXDRT.Report)
  Set m_Rs = valor
End Property

'Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'    CommonDialog.Flags = cdlPDPrintSetup Or cdlPDReturnIC
'    CommonDialog.ShowPrinter
'End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbDefault
    CRViewer1.EnableGroupTree = True
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Public Sub SetReporte(rptReporteCrystal As CRAXDRT.Report) 'CRPEAUTo.Report)
    Screen.MousePointer = vbHourglass
    Set oReporte = rptReporteCrystal
    Set RCrystal = rptReporteCrystal
    CRViewer1.ReportSource = oReporte
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault
End Sub

Public Sub ExportaExcel(NomArchivo)
    On Error GoTo SERROR
    RCrystal.ExportOptions.DestinationType = crEDTDiskFile
    RCrystal.ExportOptions.DiskFileName = NomArchivo
    RCrystal.ExportOptions.FormatType = crEFTExcel50
    RCrystal.ExportOptions.ExchangeFolderPath = Rep_Documents
    RCrystal.Export False
    Exit Sub
SERROR:
    Mensajes err.Description

End Sub

Public Sub ExportaPdf(NomArchivo)
    On Error GoTo SERROR
    RCrystal.ExportOptions.DestinationType = 1 'crEDTDiskFile
    RCrystal.ExportOptions.DiskFileName = NomArchivo
    RCrystal.ExportOptions.FormatType = 31 'crEFTPortableDocFormat
    RCrystal.ExportOptions.ExchangeFolderPath = Rep_Documents
    RCrystal.ExportOptions.PDFExportAllPages = True
    RCrystal.Export False
    Exit Sub
SERROR:
    Mensajes err.Description
End Sub

Private Sub Picture1_Click()
    oReporte.PrinterSetup Me.hwnd
    CRViewer1.ReportSource = oReporte
    CRViewer1.ViewReport
    DoEvents
    CRViewer1.PrintReport
End Sub
