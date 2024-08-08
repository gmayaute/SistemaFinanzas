VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Menu_ListView 
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   ScaleHeight     =   720
   ScaleWidth      =   2550
   Begin VB.PictureBox picMenu 
      BorderStyle     =   0  'None
      Height          =   620
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   2525
      Begin MSComctlLib.ListView lvwMenu 
         Height          =   525
         Left            =   30
         TabIndex        =   1
         Top             =   45
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   926
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
         EndProperty
      End
      Begin VB.Line line_sup 
         BorderColor     =   &H80000005&
         X1              =   10
         X2              =   2480
         Y1              =   10
         Y2              =   10
      End
      Begin VB.Line line_izq 
         BorderColor     =   &H80000005&
         X1              =   10
         X2              =   10
         Y1              =   10
         Y2              =   580
      End
      Begin VB.Line line_der_2 
         X1              =   2505
         X2              =   2505
         Y1              =   20
         Y2              =   610
      End
      Begin VB.Line line_inf_2 
         X1              =   2505
         X2              =   0
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line line_der_1 
         BorderColor     =   &H8000000C&
         X1              =   2490
         X2              =   2490
         Y1              =   30
         Y2              =   580
      End
      Begin VB.Line line_inf_1 
         BorderColor     =   &H8000000C&
         X1              =   2490
         X2              =   15
         Y1              =   585
         Y2              =   585
      End
   End
End
Attribute VB_Name = "Menu_ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const WM_SETREDRAW As Long = &HB&
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Event Click()
Attribute Click.VB_Description = "Ocurre cuando se hace click sobre algún item del Menu"
Event MouseLeave()
Attribute MouseLeave.VB_Description = "Ocurre cuando el puntero del mouse abandonó el Menu"
Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Sub setImgList(ByVal img As ImageList)
Attribute setImgList.VB_Description = "Establece la Lista de la cual se pueden extraer imagenes para asignar a los items del Menu"
    Set lvwMenu.SmallIcons = img
End Sub
Public Sub ClearItems()
Attribute ClearItems.VB_Description = "Limpia los items del Menu"
    lvwMenu.ListItems.Clear
End Sub
Public Sub addItem_Menu(ByVal strCaption As String, Optional ByVal strTag As String, Optional ByVal icon As Integer, Optional ByVal ghost As Boolean = False)
Attribute addItem_Menu.VB_Description = "Agrega un item al Menu"
    Dim lItem As ListItem
    Set lItem = lvwMenu.ListItems.Add(, , strCaption, , icon)
    lItem.tag = strTag
    lItem.Ghosted = ghost
    If ghost Then
        lItem.ForeColor = &H80000011
    End If
    Set lItem = Nothing
End Sub
Private Sub lvwMenu_Click()
    If Not lvwMenu.SelectedItem Is Nothing Then
        If lvwMenu.SelectedItem.Ghosted = False Then
            UserControl.Extender.Visible = False
        End If
    End If
    RaiseEvent Click
End Sub
Private Sub lvwMenu_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lvwMenu_Click
    End If
End Sub
Private Sub lvwMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lItem As ListItem
    If X <= lvwMenu.Width And X >= 0 And Y <= lvwMenu.Height And Y >= 0 Then
        SetCapture lvwMenu.hWnd
        If lvwMenu.Visible = True Then lvwMenu.SetFocus
        Set lItem = lvwMenu.HitTest(X, Y)
        If Not lItem Is Nothing Then
            lItem.Selected = True
            lvwMenu.ToolTipText = lItem.Text
        End If
    Else
        ReleaseCapture
        UserControl.Extender.Visible = False
        RaiseEvent MouseLeave
    End If
    Set lItem = Nothing
End Sub
Public Function getSelectedItem_Index() As Integer
Attribute getSelectedItem_Index.VB_Description = "Obtiene el indice del item seleccionado"
    getSelectedItem_Index = lvwMenu.SelectedItem.Index
End Function
Public Function getSelectedItem_Tag() As String
Attribute getSelectedItem_Tag.VB_Description = "Obtiene el tag del item seleccionado"
    getSelectedItem_Tag = lvwMenu.SelectedItem.tag
End Function
Public Function SelectedItem_Ghosted() As Boolean
Attribute SelectedItem_Ghosted.VB_Description = "Indica si el item seleccionado está bloqueado"
    SelectedItem_Ghosted = lvwMenu.SelectedItem.Ghosted
End Function
Private Function AnchoCol() As Long
    Dim lIt As ListItem
    For Each lIt In lvwMenu.ListItems
       If lIt.Width > AnchoCol Then
            AnchoCol = lIt.Width
       End If
    Next
End Function
Private Sub SetSize()
    lvwMenu.Width = 0
    picMenu.Width = 0
    Ajustar_Cols lvwMenu, 1, -2
    lvwMenu.ColumnHeaders(1).Width = lvwMenu.ColumnHeaders(1).Width + 400
    Dim lIt As ListItem
    Dim acum As Long
    For Each lIt In lvwMenu.ListItems
        acum = acum + lIt.Height
    Next
    lvwMenu.Height = acum
    lvwMenu.Width = AnchoCol
    picMenu.Width = lvwMenu.Width + 80
    picMenu.Height = lvwMenu.Height + 100
    line_der_2.X1 = picMenu.Width - 20
    line_der_2.X2 = picMenu.Width - 20
    line_der_1.X1 = line_der_2.X1 - 15
    line_der_1.X2 = line_der_2.X2 - 15
    line_der_1.Y2 = picMenu.Height - 10
    line_der_2.Y2 = picMenu.Height - 10
    line_izq.Y2 = picMenu.Height - 50
    line_sup.X2 = picMenu.Width - 20
    line_inf_2.Y1 = picMenu.Height - 30
    line_inf_2.Y2 = picMenu.Height - 30
    line_inf_1.Y1 = line_inf_2.Y1 - 15
    line_inf_1.Y2 = line_inf_2.Y2 - 15
    line_inf_1.X1 = 0
    line_inf_2.X1 = 0
    line_inf_1.X2 = line_sup.X2
    line_inf_2.X2 = line_sup.X2
    UserControl.Width = picMenu.Width
    UserControl.Height = picMenu.Height
End Sub
Private Sub Ajustar_Cols(El_ListView As ListView, ByVal La_Columna As Long, ByVal Modo_De_Ajuste As Long)
    With El_ListView
        Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, _
                             La_Columna - 1, ByVal Modo_De_Ajuste)
    End With
End Sub
Private Sub SetCoordenadas(ByVal X As Single, Y As Single)
    Dim dif_Top As Long, dif_Left As Long
    SetSize
    dif_Top = Y + UserControl.Height - getAltoContainer + 100
    dif_Left = X + UserControl.Width - getAnchoContainer + 100
    If dif_Top >= 0 Then
        UserControl.Extender.Top = Y - dif_Top
    Else
        UserControl.Extender.Top = Y
    End If
    If dif_Left >= 0 Then
        UserControl.Extender.Left = X - dif_Left
    Else
        UserControl.Extender.Left = X
    End If
End Sub
Private Function getAltoContainer() As Long
    Dim oRect As RECT
    GetClientRect UserControl.ContainerHwnd, oRect
    getAltoContainer = (oRect.Top + oRect.Bottom) * Screen.TwipsPerPixelY
End Function
Private Function getAnchoContainer() As Long
    Dim oRect As RECT
    GetClientRect UserControl.ContainerHwnd, oRect
    getAnchoContainer = (oRect.Left + oRect.Right) * Screen.TwipsPerPixelX
End Function
Public Sub Show(ByVal X As Single, Y As Single)
Attribute Show.VB_Description = "Muestra el menu"
    SetCoordenadas X, Y
    UserControl.Extender.ZOrder 0
    UserControl.Extender.Visible = True
End Sub
