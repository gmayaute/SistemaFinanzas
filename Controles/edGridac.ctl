VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl flxEditfac 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   MaskColor       =   &H00E0E0E0&
   ScaleHeight     =   3600
   ScaleWidth      =   4605
   ToolboxBitmap   =   "edGridac.ctx":0000
   Begin VB.TextBox txtTemp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "edGridac.ctx":0312
      Top             =   2370
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid flxControl 
      Height          =   2085
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   3678
      _Version        =   393216
      BackColor       =   16777215
      BackColorBkg    =   12632256
      GridColorFixed  =   12632256
      WordWrap        =   -1  'True
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtedit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2340
      Visible         =   0   'False
      Width           =   2115
   End
End
Attribute VB_Name = "flxEditfac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit
Const cNuevaFila As String = ">>*"  ' Para indicar que es una nueva fila
Private PresionaF2 As Boolean         ' Flag para saber si se presiono la tecla F2
Public ControlVisible As Boolean     ' Si el control est� o no visible (edit�ndose)
Private lastRow As Long               ' La �ltima fila en que se edit�
Private LastCol As Long               ' La �ltima columna en que se edit�
Private Tamanio(1 To 13, 1 To 2) As Long
Private C_ColType() As flextype
Private c_MaxLenght() As Integer
Private c_Decimales() As Integer
Private c_CaracteresValidos() As String
Private confirmar As Boolean
Public TipodeCampo As flextype

'ANTIGUO
'Const ColorHabilitado = &HFFFFFF
'Const ColorDeshabilitado = &HFFC0C0

Const ColorHabilitado = &HFFFFFF '&H80000009
Const ColorDeshabilitado = &HE0E0E0



Private cNombre As String


Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Compare(ByVal Row1 As Long, ByVal Row2 As Long, CMP As Integer)
Event EnterCell()
Event HitTest(X As Single, Y As Single, HitResult As Integer)
Event LeaveCell()
Event Paint()
Event Resize()
Event RowColChange()
Event SelChange()
Event Scroll()
Event Show()
Event FilasBorradas()

Private bNoEditar As Boolean


Public Property Get Nombre() As String
    Nombre = cNombre
End Property

Public Property Let Nombre(ByVal New_Name As String)
     cNombre = New_Name
End Property

Public Property Get NoEditar() As Boolean
    NoEditar = bNoEditar
End Property

Public Property Let NoEditar(ByVal new_NoEditar As Boolean)
    bNoEditar = new_NoEditar
End Property

Public Property Get Visible() As Boolean
    Visible = flxControl.Visible
End Property
Public Property Let Visible(ByVal New_Visible As Boolean)
    flxControl.Visible = New_Visible
    
End Property

Public Property Get Rowsel() As AppearanceSettings
    Rowsel = flxControl.Rowsel
End Property

Public Property Let Rowsel(ByVal New_Rowsel As Long)
    flxControl.Rowsel = New_Rowsel
    PropertyChanged "RowSel"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = flxControl.BackColor
End Property

Public Property Get FormColor() As OLE_COLOR
    FormColor = UserControl.BackColor
End Property

Public Property Let FormColor(ByVal New_FormBackColor As OLE_COLOR)
    UserControl.BackColor = New_FormBackColor
    PropertyChanged "FormColor"
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    flxControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get CellBackColor() As OLE_COLOR
    CellBackColor = flxControl.CellBackColor
End Property
Public Property Let CellBackColor(ByVal New_CellBackColor As OLE_COLOR)
    flxControl.CellBackColor() = New_CellBackColor
    PropertyChanged "CellBackColor"
End Property
Public Property Get CellForeColor() As OLE_COLOR
    CellForeColor = flxControl.CellForeColor
End Property
Public Property Let CellForeColor(ByVal New_CellForeColor As OLE_COLOR)
    flxControl.CellForeColor() = New_CellForeColor
    PropertyChanged "CellForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = flxControl.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    flxControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
    Set Font = flxControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set flxControl.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Appearance() As AppearanceSettings
    Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As AppearanceSettings)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property
Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property
Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = flxControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    flxControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Sub Refresh()
    flxControl.Refresh
End Sub
Private Sub flxControl_Click()
    Grid2_Click
    RaiseEvent Click
End Sub
Private Sub flxControl_DblClick()
    Grid2_DblClick
    RaiseEvent DblClick
End Sub
Private Sub flxControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Grid2_KeyDown KeyCode, Shift
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub flxControl_KeyPress(KeyAscii As Integer)
    Grid2_KeyPress KeyAscii
    RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub flxControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub flxControl_LostFocus()
    txtTemp.Visible = False
End Sub

Private Sub flxControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub flxControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub flxControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Public Property Get AllowUserResizing() As AllowUserResizeSettings
    AllowUserResizing = flxControl.AllowUserResizing
End Property
Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
    flxControl.AllowUserResizing() = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"
End Property
Public Property Get AllowBigSelection() As Boolean
    AllowBigSelection = flxControl.AllowBigSelection
End Property
Public Property Let AllowBigSelection(ByVal New_AllowBigSelection As Boolean)
    flxControl.AllowBigSelection() = New_AllowBigSelection
    PropertyChanged "AllowBigSelection"
End Property
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
    flxControl.AddItem Item, Index
End Sub
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property
Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = flxControl.BackColorSel
End Property
Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
    flxControl.BackColorSel() = New_BackColorSel
    PropertyChanged "BackColorSel"
End Property
Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = flxControl.BackColorFixed
End Property
Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    flxControl.BackColorFixed() = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property
Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = flxControl.BackColorBkg
End Property
Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    flxControl.BackColorBkg() = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property
Public Property Get CellWidth() As Long
    CellWidth = flxControl.CellWidth
End Property
Public Property Get CellTop() As Long
    CellTop = flxControl.CellTop
End Property

Public Property Get Top() As Long
    Top = flxControl.Top
End Property

Public Property Get Left() As Long
    Left = flxControl.Left
End Property

Public Property Get CellPicture() As Picture
    Set CellPicture = flxControl.CellPicture
End Property
Public Property Set CellPicture(ByVal New_CellPicture As Picture)
    Set flxControl.CellPicture = New_CellPicture
    PropertyChanged "CellPicture"
End Property
Public Property Get CellLeft() As Long
    CellLeft = flxControl.CellLeft
End Property
Public Property Get CellHeight() As Long
    CellHeight = flxControl.CellHeight
End Property
Public Sub Clear()
    flxControl.Clear
End Sub
Public Property Get ColWidth(ByVal Index As Long) As Long
    ColWidth = flxControl.ColWidth(Index)
End Property
Public Property Let ColWidth(ByVal Index As Long, ByVal New_ColWidth As Long)
    flxControl.ColWidth(Index) = New_ColWidth
    PropertyChanged "ColWidth"
End Property
Public Property Get Cols() As Long
    Cols = flxControl.Cols
End Property
Public Property Let Cols(ByVal New_Cols As Long)
    flxControl.Cols() = New_Cols
    ReDim C_ColType(0 To New_Cols)
    ReDim c_Decimales(0 To New_Cols)
    ReDim c_MaxLenght(0 To New_Cols)
    ReDim c_CaracteresValidos(0 To New_Cols)
    PropertyChanged "Cols"
End Property
Public Property Let ColPosition(ByVal Index As Long, ByVal New_ColPosition As Long)
    flxControl.ColPosition(Index) = New_ColPosition
    PropertyChanged "ColPosition"
End Property
Public Property Get ColPos(ByVal Index As Long) As Long
    ColPos = flxControl.ColPos(Index)
End Property
Public Property Get ColIsVisible(ByVal Index As Long) As Boolean
    ColIsVisible = flxControl.ColIsVisible(Index)
End Property
Public Property Get ColData(ByVal Index As Long) As Long
    ColData = flxControl.ColData(Index)
End Property
Public Property Let ColData(ByVal Index As Long, ByVal New_ColData As Long)
    flxControl.ColData(Index) = New_ColData
    PropertyChanged "ColData"
End Property
Public Property Get ColDecimales(ByVal Index As Long) As Integer
 On Error Resume Next
    ColDecimales = c_Decimales(Index)
End Property
Public Property Let ColDecimales(ByVal Index As Long, ByVal New_ColAlignment As Integer)
On Error Resume Next
    c_Decimales(Index) = New_ColAlignment
    PropertyChanged "ColDecimales"
End Property
Public Property Get ColMaxLength(ByVal Index As Long) As Integer
 On Error Resume Next
    ColMaxLength = c_MaxLenght(Index)
End Property
Public Property Let ColMaxLength(ByVal Index As Long, ByVal New_ColAlignment As Integer)
On Error Resume Next
    c_MaxLenght(Index) = New_ColAlignment
    PropertyChanged "ColMaxLength"
End Property
Public Property Get CaracteresValidos(ByVal Index As Long) As String
On Error Resume Next
    CaracteresValidos = c_CaracteresValidos(Index)
End Property
Public Property Let CaracteresValidos(ByVal Index As Long, ByVal New_Caracteres As String)
    c_CaracteresValidos(Index) = New_Caracteres
End Property
Public Property Get ColType(ByVal Index As Long) As flextype
    ColType = C_ColType(Index)
End Property
Public Property Let ColType(ByVal Index As Long, ByVal New_ColAlignment As flextype)
    C_ColType(Index) = New_ColAlignment
    PropertyChanged "ColType"
End Property
Public Property Get confirmarborradolinea() As Boolean
     confirmarborradolinea = confirmar
End Property
Public Property Let confirmarborradolinea(ByVal new_borrado As Boolean)
    confirmar = new_borrado
    PropertyChanged "ConfirmarBorradoLinea"
End Property
Public Property Get ColAlignment(ByVal Index As Long) As Integer
    ColAlignment = flxControl.ColAlignment(Index)
End Property
Public Property Let ColAlignment(ByVal Index As Long, ByVal New_ColAlignment As Integer)
    flxControl.ColAlignment(Index) = New_ColAlignment
    PropertyChanged "ColAlignment"
End Property
Private Sub flxControl_Compare(ByVal Row1 As Long, ByVal Row2 As Long, CMP As Integer)
    RaiseEvent Compare(Row1, Row2, CMP)
End Sub
Public Property Get CurrentY() As Single
    CurrentY = UserControl.CurrentY
End Property
Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property
Public Property Get CurrentX() As Single
    CurrentX = UserControl.CurrentX
End Property
Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property
Private Sub flxControl_EnterCell()
    RaiseEvent EnterCell
End Sub
Public Property Get DrawWidth() As Integer
    DrawWidth = UserControl.DrawWidth
End Property
Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    UserControl.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property
Public Property Get DrawStyle() As Integer
    DrawStyle = UserControl.DrawStyle
End Property
Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property
Public Property Get DrawMode() As Integer
    DrawMode = UserControl.DrawMode
End Property
Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    UserControl.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property
Public Property Get FocusRect() As FocusRectSettings
    FocusRect = flxControl.FocusRect
End Property
Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
    flxControl.FocusRect() = New_FocusRect
    PropertyChanged "FocusRect"
End Property
Public Property Get FixedRows() As Long
    FixedRows = flxControl.FixedRows
End Property
Public Property Let FixedRows(ByVal New_FixedRows As Long)
    flxControl.FixedRows() = New_FixedRows
    PropertyChanged "FixedRows"
End Property
Public Property Get FixedCols() As Long
    FixedCols = flxControl.FixedCols
End Property
Public Property Let FixedCols(ByVal New_FixedCols As Long)
    flxControl.FixedCols() = New_FixedCols
    PropertyChanged "FixedCols"
End Property
Public Property Get FixedAlignment(ByVal Index As Long) As Integer
    FixedAlignment = flxControl.FixedAlignment(Index)
End Property
Public Property Let FixedAlignment(ByVal Index As Long, ByVal New_FixedAlignment As Integer)
    flxControl.FixedAlignment(Index) = New_FixedAlignment
    PropertyChanged "FixedAlignment"
End Property
Public Property Get FillStyle() As FillStyleSettings
    FillStyle = flxControl.FillStyle
End Property
Public Property Let FillStyle(ByVal New_FillStyle As FillStyleSettings)
    flxControl.FillStyle() = New_FillStyle
    PropertyChanged "FillStyle"
End Property
Public Property Get FillColor() As OLE_COLOR
    FillColor = UserControl.FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property
Public Property Get FormatString() As String
    FormatString = flxControl.FormatString
End Property
Public Property Let FormatString(ByVal New_FormatString As String)
    flxControl.FormatString() = New_FormatString
    PropertyChanged "FormatString"
    PropertyChanged "cols"
End Property
Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = flxControl.ForeColorSel
End Property
Public Property Let ForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
    flxControl.ForeColorSel() = New_ForeColorSel
    PropertyChanged "ForeColorSel"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = flxControl.ForeColorFixed
End Property
Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    flxControl.ForeColorFixed() = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property
Public Property Get Col() As Integer
    Col = flxControl.Col
End Property
Public Property Let Col(ByVal New_col As Integer)
    flxControl.Col = New_col
    PropertyChanged "col"
End Property
Public Property Get row() As Integer
    row = flxControl.row
End Property
Public Property Let row(ByVal New_row As Integer)
    flxControl.row = New_row
    PropertyChanged "row"
End Property
Public Property Get GridLineWidth() As Integer
    GridLineWidth = flxControl.GridLineWidth
End Property
Public Property Let GridLineWidth(ByVal New_GridLineWidth As Integer)
    flxControl.GridLineWidth() = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property
Public Property Get GridLinesFixed() As GridLineSettings
    GridLinesFixed = flxControl.GridLinesFixed
End Property
Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
    flxControl.GridLinesFixed() = New_GridLinesFixed
    PropertyChanged "GridLinesFixed"
End Property
Public Property Get GridLines() As GridLineSettings
    GridLines = flxControl.GridLines
End Property
Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
    flxControl.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property
Public Property Get GridColorFixed() As OLE_COLOR
    GridColorFixed = flxControl.GridColorFixed
End Property
Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
    flxControl.GridColorFixed() = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property
Public Property Get GridColor() As OLE_COLOR
    GridColor = flxControl.GridColor
End Property
Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    flxControl.GridColor() = New_GridColor
    PropertyChanged "GridColor"
End Property
Private Sub txtedit_Change()
    Dim I As Integer, LINEAS As Long
    Dim cadenatemp As String
    LINEAS = LineCount(txtedit)
    cadenatemp = flxControl.TextMatrix(flxControl.row, flxControl.Col)
    flxControl.TextMatrix(flxControl.row, flxControl.Col) = ""
    For I = 1 To flxControl.Cols - 1
        txtTemp.Height = Tamanio(I, 1)
        txtTemp.Width = Tamanio(I, 2) + 20
        txtTemp.Text = flxControl.TextMatrix(flxControl.row, I)
        If LineCount(txtTemp) > LINEAS Then
            LINEAS = LineCount(txtTemp)
        End If
    Next
    flxControl.TextMatrix(flxControl.row, flxControl.Col) = cadenatemp
    If LINEAS = 1 Then
        txtedit.Height = 330
        flxControl.RowHeight(flxControl.row) = 350
    End If
    If LINEAS = 2 Then
        txtedit.Height = 550
        flxControl.RowHeight(flxControl.row) = 570
    End If
    If LINEAS = 3 Then
        txtedit.Height = 760
        flxControl.RowHeight(flxControl.row) = 780
    End If
    If LINEAS = 4 Then
        txtedit.Height = 990
        flxControl.RowHeight(flxControl.row) = 1010
    End If
    If LINEAS = 5 Then
        txtedit.Height = 1220
        flxControl.RowHeight(flxControl.row) = 1240
    End If
    If LINEAS = 6 Then
        txtedit.Height = 1450
        flxControl.RowHeight(flxControl.row) = 1470
    End If
End Sub
Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    RaiseEvent HitTest(X, Y, HitResult)
End Sub
Public Property Get valor(columna As Integer, fila As Integer) As Variant
    If Not (columna > flxControl.Cols Or fila > flxControl.Rows) Then
        Select Case C_ColType(columna)
            Case flextype.cadena
                valor = flxControl.TextMatrix(columna, fila)
            Case flextype.Entero
                On Error Resume Next
                    valor = CInt(flxControl.TextMatrix(columna, fila))
                On Error GoTo 0
            Case flextype.Numero
                On Error Resume Next
                    valor = CDbl(flxControl.TextMatrix(columna, fila))
                On Error GoTo 0
        End Select
    End If
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Private Sub flxControl_LeaveCell()
    RaiseEvent LeaveCell
End Sub
Public Property Get Image() As Picture
    Set Image = UserControl.Image
End Property
Public Property Get MergeRow(ByVal Index As Long) As Boolean
    MergeRow = flxControl.MergeRow(Index)
End Property
Public Property Let MergeRow(ByVal Index As Long, ByVal New_MergeRow As Boolean)
    flxControl.MergeRow(Index) = New_MergeRow
    PropertyChanged "MergeRow"
End Property
Public Property Get MergeCol(ByVal Index As Long) As Boolean
    MergeCol = flxControl.MergeCol(Index)
End Property
Public Property Let MergeCol(ByVal Index As Long, ByVal New_MergeCol As Boolean)
    flxControl.MergeCol(Index) = New_MergeCol
    PropertyChanged "MergeCol"
End Property
Public Property Get MergeCells() As MergeCellsSettings
    MergeCells = flxControl.MergeCells
End Property
Public Property Let MergeCells(ByVal New_MergeCells As MergeCellsSettings)
    flxControl.MergeCells() = New_MergeCells
    PropertyChanged "MergeCells"
End Property
Public Property Get MaskPicture() As Picture
    Set MaskPicture = UserControl.MaskPicture
End Property
Public Property Set MaskPicture(ByVal New_MaskPicture As Picture)
    Set UserControl.MaskPicture = New_MaskPicture
    PropertyChanged "MaskPicture"
End Property
Public Property Get MaskColor() As Long
    MaskColor = UserControl.MaskColor
End Property
Public Property Let MaskColor(ByVal New_MaskColor As Long)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property
Public Property Get MouseCol() As Long
    MouseCol = flxControl.MouseCol
End Property
Public Property Get MouseIcon() As Picture
    Set MouseIcon = flxControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set flxControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As MousePointerSettings
    MousePointer = flxControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerSettings)
    flxControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
Public Property Get MouseRow() As Long
    MouseRow = flxControl.MouseRow
End Property
Private Sub UserControl_Initialize()
    TipodeCampo = cadena
    flxControl.WordWrap = True
End Sub
Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub
Public Sub RemoveItem(ByVal Index As Long)
    flxControl.RemoveItem Index
End Sub
Private Sub UserControl_Resize()
    flxControl.Width = UserControl.Width
    flxControl.Height = UserControl.Height
    RaiseEvent Resize
End Sub
Public Property Get Rows() As Long
    Rows = flxControl.Rows
End Property
Public Property Let Rows(ByVal New_Rows As Long)
    flxControl.Rows() = New_Rows
    PropertyChanged "Rows"
End Property
Public Property Let RowPosition(ByVal Index As Long, ByVal New_RowPosition As Long)
    flxControl.RowPosition(Index) = New_RowPosition
    PropertyChanged "RowPosition"
End Property
Public Property Get RowPos(ByVal Index As Long) As Long
    RowPos = flxControl.RowPos(Index)
End Property
Public Property Get RowIsVisible(ByVal Index As Long) As Boolean
    RowIsVisible = flxControl.RowIsVisible(Index)
End Property
Public Property Get RowHeightMin() As Long
    RowHeightMin = flxControl.RowHeightMin
End Property
Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    flxControl.RowHeightMin() = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property
Public Property Get RowHeight(ByVal Index As Long) As Long
    RowHeight = flxControl.RowHeight(Index)
End Property
Public Property Let RowHeight(ByVal Index As Long, ByVal New_RowHeight As Long)
    flxControl.RowHeight(Index) = New_RowHeight
    PropertyChanged "RowHeight"
End Property
Public Property Get RowData(ByVal Index As Long) As Long
    RowData = flxControl.RowData(Index)
End Property
Public Property Let RowData(ByVal Index As Long, ByVal New_RowData As Long)
    flxControl.RowData(Index) = New_RowData
    PropertyChanged "RowData"
End Property
Private Sub flxControl_RowColChange()
    RaiseEvent RowColChange
    On Error Resume Next
    flxControl.SetFocus
    txtTemp.Visible = False
    On Error GoTo 0
End Sub
Private Sub flxControl_SelChange()
    RaiseEvent SelChange
End Sub
Public Property Get ScrollTrack() As Boolean
    ScrollTrack = flxControl.ScrollTrack
End Property
Public Property Let ScrollTrack(ByVal New_ScrollTrack As Boolean)
    flxControl.ScrollTrack() = New_ScrollTrack
    PropertyChanged "ScrollTrack"
End Property
Public Property Get ScrollBars() As ScrollBarsSettings
    ScrollBars = flxControl.ScrollBars
End Property
Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    flxControl.ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property
Private Sub flxControl_Scroll()
    Grid2_Scroll
    RaiseEvent Scroll
End Sub
Public Property Get SelectionMode() As SelectionModeSettings
    SelectionMode = flxControl.SelectionMode
End Property
Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    flxControl.SelectionMode() = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property
Private Sub UserControl_Show()
    RaiseEvent Show
End Sub
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
    UserControl.Size Width, Height
End Sub
Public Property Get ToolTipText() As String
    ToolTipText = flxControl.ToolTipText
End Property
Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    flxControl.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Public Property Get TextStyleFixed() As TextStyleSettings
    TextStyleFixed = flxControl.TextStyleFixed
End Property
Public Property Let TextStyleFixed(ByVal New_TextStyleFixed As TextStyleSettings)
    flxControl.TextStyleFixed() = New_TextStyleFixed
    PropertyChanged "TextStyleFixed"
End Property
Public Property Get TextStyle() As TextStyleSettings
    TextStyle = flxControl.TextStyle
End Property
Public Property Let TextStyle(ByVal New_TextStyle As TextStyleSettings)
    flxControl.TextStyle() = New_TextStyle
    PropertyChanged "TextStyle"
End Property
Public Property Get TextMatrix(ByVal row As Long, ByVal Col As Long) As String
On Error Resume Next
    TextMatrix = flxControl.TextMatrix(row, Col)
End Property
Public Property Let TextMatrix(ByVal row As Long, ByVal Col As Long, ByVal New_TextMatrix As String)
    flxControl.TextMatrix(row, Col) = New_TextMatrix
    PropertyChanged "TextMatrix"
End Property
Public Property Get TextArray(ByVal Index As Long) As String
    TextArray = flxControl.TextArray(Index)
End Property
Public Property Let TextArray(ByVal Index As Long, ByVal New_TextArray As String)
    flxControl.TextArray(Index) = New_TextArray
    PropertyChanged "TextArray"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    flxControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    flxControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set flxControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    flxControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    flxControl.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 0)
    flxControl.AllowBigSelection = PropBag.ReadProperty("AllowBigSelection", True)
    flxControl.BackColorSel = PropBag.ReadProperty("BackColorSel", &H80000005)
    flxControl.BackColorFixed = PropBag.ReadProperty("BackColorFixed", RGB(150, 150, 150))
    flxControl.BackColorBkg = PropBag.ReadProperty("BackColorBkg", 8421504)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    Set CellPicture = PropBag.ReadProperty("CellPicture", Nothing)
    ReDim C_ColType(0 To PropBag.ReadProperty("Cols", 1))
    ReDim c_Decimales(0 To PropBag.ReadProperty("Cols", 1))
    ReDim c_MaxLenght(0 To PropBag.ReadProperty("Cols", 1))
    ReDim c_CaracteresValidos(0 To PropBag.ReadProperty("Cols", 1))
    For Index = 0 To flxControl.Cols - 1
        flxControl.ColWidth(Index) = PropBag.ReadProperty("ColWidth" & Index, 0)
        flxControl.ColPosition(Index) = PropBag.ReadProperty("ColPosition" & Index, 0)
        flxControl.ColData(Index) = PropBag.ReadProperty("ColData" & Index, 0)
        C_ColType(Index) = PropBag.ReadProperty("ColType" & Index, Entero)
        flxControl.ColAlignment(Index) = PropBag.ReadProperty("ColAlignment" & Index, 0)
        flxControl.FixedAlignment(Index) = PropBag.ReadProperty("FixedAlignment" & Index, 0)
        flxControl.MergeCol(Index) = PropBag.ReadProperty("MergeCol" & Index, 0)
    Next
    confirmar = PropBag.ReadProperty("ConfirmarBorradoLinea", True)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    UserControl.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    flxControl.FocusRect = PropBag.ReadProperty("FocusRect", 1)
    flxControl.FixedRows = PropBag.ReadProperty("FixedRows", 1)
    flxControl.FixedCols = PropBag.ReadProperty("FixedCols", 1)
    flxControl.FillStyle = PropBag.ReadProperty("FillStyle", 0)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    flxControl.ForeColorSel = PropBag.ReadProperty("ForeColorSel", RGB(0, 0, 255))
    flxControl.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", RGB(220, 220, 220))
    flxControl.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    flxControl.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", 2)
    flxControl.GridLines = PropBag.ReadProperty("GridLines", 1)
    flxControl.GridColorFixed = PropBag.ReadProperty("GridColorFixed", 0)
    flxControl.GridColor = PropBag.ReadProperty("GridColor", 12632256)
    For Index = 0 To flxControl.Rows - 1
        flxControl.MergeRow(Index) = PropBag.ReadProperty("MergeRow" & Index, 0)
        flxControl.RowPosition(Index) = PropBag.ReadProperty("RowPosition" & Index, 0)
        flxControl.RowHeight(Index) = PropBag.ReadProperty("RowHeight" & Index, 0)
        flxControl.RowData(Index) = PropBag.ReadProperty("RowData" & Index, 0)
    Next
    flxControl.MergeCells = PropBag.ReadProperty("MergeCells", 0)
    Set MaskPicture = PropBag.ReadProperty("MaskPicture", Nothing)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    flxControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    flxControl.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
    flxControl.ScrollTrack = PropBag.ReadProperty("ScrollTrack", False)
    flxControl.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
    flxControl.SelectionMode = PropBag.ReadProperty("SelectionMode", 0)
    flxControl.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    flxControl.TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", 0)
    flxControl.TextStyle = PropBag.ReadProperty("TextStyle", 0)
    flxControl.Rows = PropBag.ReadProperty("Rows", 2)
    flxControl.Cols = PropBag.ReadProperty("Cols", 2)
    flxControl.FormatString = PropBag.ReadProperty("FormatString", "")
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    Call PropBag.WriteProperty("BackColor", flxControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", flxControl.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", flxControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", flxControl.BorderStyle, 1)
    Call PropBag.WriteProperty("AllowUserResizing", flxControl.AllowUserResizing, 0)
    Call PropBag.WriteProperty("AllowBigSelection", flxControl.AllowBigSelection, True)
    Call PropBag.WriteProperty("BackColorSel", flxControl.BackColorSel, 2147483661#)
    Call PropBag.WriteProperty("BackColorFixed", flxControl.BackColorFixed, 2147483663#)
    Call PropBag.WriteProperty("BackColorBkg", flxControl.BackColorBkg, 8421504)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("CellPicture", CellPicture, Nothing)
    Call PropBag.WriteProperty("ConfirmarBorradoLinea", confirmar, True)
    Call PropBag.WriteProperty("Cols", flxControl.Cols, 2)
    On Error Resume Next
    For Index = 0 To flxControl.Cols - 1
        Call PropBag.WriteProperty("ColWidth" & Index, flxControl.ColWidth(Index), 0)
        Call PropBag.WriteProperty("ColData" & Index, flxControl.ColData(Index), 0)
        Call PropBag.WriteProperty("ColAlignment" & Index, flxControl.ColAlignment(Index), 0)
        Call PropBag.WriteProperty("FixedAlignment" & Index, flxControl.FixedAlignment(Index), 0)
        Call PropBag.WriteProperty("MergeCol" & Index, flxControl.MergeCol(Index), 0)
        Call PropBag.WriteProperty("ColType" & Index, C_ColType(Index), Entero)
        If err <> 0 Then ReDim C_ColType(0 To flxControl.Cols - 1)
        Call PropBag.WriteProperty("ColMaxLenght" & Index, c_MaxLenght(Index), 0)
        Call PropBag.WriteProperty("ColDecimales" & Index, c_Decimales(Index), 0)
    Next
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("FocusRect", flxControl.FocusRect, 1)
    Call PropBag.WriteProperty("FixedRows", flxControl.FixedRows, 1)
    Call PropBag.WriteProperty("FixedCols", flxControl.FixedCols, 1)
    Call PropBag.WriteProperty("FillStyle", flxControl.FillStyle, 0)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("FormatString", flxControl.FormatString, "")
    Call PropBag.WriteProperty("ForeColorSel", flxControl.ForeColorSel, 2147483662#)
    Call PropBag.WriteProperty("ForeColorFixed", flxControl.ForeColorFixed, 2147483666#)
    Call PropBag.WriteProperty("GridLineWidth", flxControl.GridLineWidth, 1)
    Call PropBag.WriteProperty("GridLinesFixed", flxControl.GridLinesFixed, 2)
    Call PropBag.WriteProperty("GridLines", flxControl.GridLines, 1)
    Call PropBag.WriteProperty("GridColorFixed", flxControl.GridColorFixed, 0)
    Call PropBag.WriteProperty("GridColor", flxControl.GridColor, 12632256)
    Call PropBag.WriteProperty("MergeCells", flxControl.MergeCells, 0)
    Call PropBag.WriteProperty("MaskPicture", MaskPicture, Nothing)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", flxControl.MousePointer, 0)
    Call PropBag.WriteProperty("Rows", flxControl.Rows, 2)
    Call PropBag.WriteProperty("RowHeightMin", flxControl.RowHeightMin, 0)
    For Index = 0 To flxControl.Rows - 1
        Call PropBag.WriteProperty("MergeRow" & Index, flxControl.MergeRow(Index), 0)
        Call PropBag.WriteProperty("RowHeight" & Index, flxControl.RowHeight(Index), 0)
        Call PropBag.WriteProperty("RowData" & Index, flxControl.RowData(Index), 0)
    Next
    Call PropBag.WriteProperty("ScrollTrack", flxControl.ScrollTrack, False)
    Call PropBag.WriteProperty("ScrollBars", flxControl.ScrollBars, 3)
    Call PropBag.WriteProperty("SelectionMode", flxControl.SelectionMode, 0)
    Call PropBag.WriteProperty("ToolTipText", flxControl.ToolTipText, "")
    Call PropBag.WriteProperty("TextStyleFixed", flxControl.TextStyleFixed, 0)
    Call PropBag.WriteProperty("TextStyle", flxControl.TextStyle, 0)
End Sub
Private Sub txtEdit_GotFocus()
     With txtedit
        If PresionaF2 = True Then
            .SelStart = 0
            .SelLength = Len(.Text)
            PresionaF2 = False
            Exit Sub
        End If
        mark txtedit
    End With
End Sub
Private Sub Txtedit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    On Error Resume Next ' evitamos errores desagradables, aunque podamos producir
                         ' mal funcionamientos ......
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
            AsignarCelda
            RaiseEvent KeyDown(13, 1)
            SiguienteCelda
            RaiseEvent KeyPress(KeyAscii)
        Case vbKeyEscape
            KeyAscii = 0
            txtedit.Visible = False
            ControlVisible = False
        Case Else
            Select Case C_ColType(flxControl.Col)
                Case flextype.Entero
                    If InStr("01234567890" & Mid(FormatNumber(1500), 2, 1), Chr(KeyAscii)) = 0 And KeyAscii > 31 Then
                        KeyAscii = 0
                    End If
                Case flextype.Numero
                    'If InStr(c_CaracteresValidos(flxControl.Col), Chr(KeyAscii)) = 0 And KeyAscii > 31 Then
                    If InStr(1, "-01234567890.", Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                    flxControl.WordWrap = True
                Case flextype.cadena
                    If c_CaracteresValidos(flxControl.Col) = "" Then
                    Else
                        If InStr(c_CaracteresValidos(flxControl.Col), Chr(KeyAscii)) = 0 And KeyAscii > 31 Then
                            KeyAscii = 0
                        End If
                    End If
            End Select
    End Select
End Sub
Private Sub AsignarCelda()
    Dim s As String
    ocultarControles
    ControlVisible = False
    s = txtedit.Text
    On Error Resume Next
    Select Case TipodeCampo 'C_ColType(flxControl.col)
      Case flextype.Entero
         flxControl.TextMatrix(lastRow, LastCol) = Space(10) & FormatNumber(CDbl(s), 0)
      Case flextype.Numero
          flxControl.TextMatrix(lastRow, LastCol) = Space(10) & FormatNumber(CDbl(s), c_Decimales(flxControl.Col))
      Case flextype.cadena
         flxControl.TextMatrix(lastRow, LastCol) = s  'UCase(s)
      Case Else
        flxControl.TextMatrix(lastRow, LastCol) = s 'UCase(s)
    End Select
End Sub
Private Sub ocultarControles()
   txtedit.Visible = False
End Sub
Private Sub Grid2_Click()
    If txtedit.Visible Then
       AsignarCelda
    End If
End Sub
Private Sub Grid2_DblClick()
    lastRow = flxControl.row
    LastCol = flxControl.Col
    If LastCol = 1 Then Exit Sub
    ocultarControles
    
    If bNoEditar = False Then
        MostrarCelda TipodeCampo
    End If
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        PresionaF2 = True
        MostrarCelda TipodeCampo
    ElseIf KeyCode = vbKeyDelete Then
    
        If bNoEditar = False Then
            BorrarFilas
        End If
    End If
End Sub
Private Sub Grid2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        KeyAscii = 0
        SiguienteCelda
    Case vbKeyEscape
        KeyAscii = 0
        AsignarCelda
    Case 32 To 255
        MostrarCelda TipodeCampo
        With txtedit
            If .Visible Then
             ' As� compruebo si es un caracte v�lido..
             DoEvents
               Txtedit_KeyPress KeyAscii
               
               If KeyAscii <> 0 Then
                    .Text = Chr$(KeyAscii)
                    .SelStart = Len(.Text) + 1
               End If
                
            End If
        End With
    End Select
End Sub
Private Sub Grid2_Scroll()
    On Error GoTo nose
    If flxControl.ColIsVisible(LastCol) = False Then
        ocultarControles
        Exit Sub
    End If
    If flxControl.RowIsVisible(lastRow) = False Then
        ocultarControles
        Exit Sub
    End If
    If ControlVisible Then
        MostrarCelda TipodeCampo
    End If
nose:
    Exit Sub
End Sub
Public Sub SiguienteCelda()
    Dim Col As Integer
    If flxControl.Col < flxControl.Cols - 1 Then
        flxControl.Col = flxControl.Col + 1
        If flxControl.CellBackColor = ColorDeshabilitado Then SiguienteCelda
    Else
        flxControl.Col = 1
        If flxControl.row < flxControl.Rows - 1 Then
            flxControl.row = flxControl.row + 1
        End If
        If flxControl.CellBackColor = ColorDeshabilitado Then SiguienteCelda
    End If
End Sub
Public Sub MostrarCelda(TipCam As flextype)
    Dim I As Integer
    Static YaEstoy As Boolean
    If (flxControl.Col <= flxControl.FixedCols - 1 Or flxControl.row <= flxControl.FixedRows - 1) Or Publimensaje <> "modificar" Then
        Exit Sub
    End If
    If YaEstoy Then Exit Sub
    YaEstoy = True
    ocultarControles
    lastRow = flxControl.row
    LastCol = flxControl.Col
    With flxControl
        If .TextMatrix(lastRow, 0) = cNuevaFila Then
            .Rows = .Rows + 1
            .TextMatrix(lastRow, 0) = lastRow
            .TextMatrix(.Rows - 1, 0) = cNuevaFila
        End If
        
        .row = lastRow
        
        If .CellBackColor = ColorDeshabilitado Then
            YaEstoy = False
            Exit Sub
        End If
        
        txtedit.Move Abs(flxControl.CellLeft - Screen.TwipsPerPixelX + 20), Abs(flxControl.CellTop - Screen.TwipsPerPixelY + 50), Abs((flxControl.CellWidth + Screen.TwipsPerPixelX * 2) - 80), Abs((flxControl.CellHeight + Screen.TwipsPerPixelY * 2) - 100)
        Tamanio(.Col, 1) = txtedit.Height
        Tamanio(.Col, 2) = txtedit.Width
    End With
    Select Case TipCam
        Case flextype.cadena
            txtedit.MaxLength = c_MaxLenght(flxControl.Col)
            txtedit.Text = Trim(flxControl.Text)
            txtedit.Alignment = vbLeftJustify
            
        Case flextype.Entero
            txtedit.MaxLength = c_MaxLenght(flxControl.Col)
            txtedit.Text = Trim(flxControl.Text)
            
            txtedit.Alignment = vbRightJustify
        Case flextype.Numero
            txtedit.MaxLength = c_MaxLenght(flxControl.Col)
            txtedit.Text = Trim(flxControl.Text)
            
            txtedit.Alignment = vbRightJustify
        Case Else
        End Select
        If Len(Trim(flxControl.Text)) = 0 Then
            If lastRow > 1 Then
                Select Case TipCam
                    Case flextype.cadena
                        txtedit.Text = Trim(flxControl.Text)
                    Case flextype.Entero
                        txtedit.Text = Trim(flxControl.Text)
                    Case flextype.Numero
                        txtedit.Text = Trim(flxControl.Text)
                    Case Else
                End Select
            End If
        End If
        txtedit.Visible = True
        If txtedit.Visible Then
            txtedit.ZOrder
            txtedit.SetFocus
        End If
    ControlVisible = True
    YaEstoy = False
End Sub
Private Sub BorrarFilas()
    Dim I As Long
    Dim J As Long
    Dim k As Long
    Dim n As Long
    Dim fila As Long
    If flxControl.Rowsel = flxControl.Rows - 1 Then
        Beep
        Exit Sub
    End If
    If flxControl.row = flxControl.Rows - 1 Then
        Beep
        Exit Sub
    End If
    If confirmar Then
        If MsgBox("� Est� usted seguro de eliminar la linea actual?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenci�n") = vbNo Then Exit Sub
    End If
    fila = flxControl.row
    I = flxControl.row
    J = flxControl.Rowsel
    If I < J Then
        k = I
        I = J
        J = k
    End If
    For n = I To J Step -1
        flxControl.RemoveItem n
    Next
    lastRow = flxControl.Rows - 1
    LastCol = 1
    On Error Resume Next
    flxControl.Col = LastCol
    flxControl.row = fila
    RaiseEvent FilasBorradas
End Sub
