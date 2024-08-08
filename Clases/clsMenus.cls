VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private msRs As MYSQL_RS
Private mMenus As Collection

Public Property Get RsMenu() As MYSQL_RS
    Set RsMenu = msRs
End Property
Public Property Let RsMenu(ByVal dMenu As MYSQL_RS)
    Set msRs = dMenu
End Property
Private Function Add(vPos As String, vNivel As String, vIndex As Integer, vMenu As Menu, vNombre As String, vVisible As Boolean, vGrupo As String, vSubGrupo As String, Optional sKey As String) As clsMenu
    Dim objNewMember As clsMenu
    Set objNewMember = New clsMenu
    objNewMember.Pos = vPos
    objNewMember.Nivel = vNivel
    objNewMember.Indice = vIndex
    objNewMember.objMenu = vMenu
    objNewMember.objMenu.Caption = vNombre
    objNewMember.Grupo = vGrupo
    objNewMember.SubGrupo = vSubGrupo
    objNewMember.objMenu.Visible = True
    If vNombre = "-" Then
        objNewMember.objMenu.Enabled = True
    Else
        objNewMember.objMenu.Enabled = vVisible
    End If
    If Len(sKey) = 0 Then
        mMenus.Add objNewMember
    Else
        mMenus.Add objNewMember, sKey
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function
Public Property Get Item(vntIndexKey As Variant) As clsMenu
    Set Item = mMenus(vntIndexKey)
End Property
Public Property Get Count() As Long
    Count = mMenus.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
    mMenus.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
    Set NewEnum = mMenus.[_NewEnum]
End Property
Private Sub Class_Initialize()
   Set mMenus = New Collection
   Set msRs = New MYSQL_RS
End Sub
Private Sub Class_Terminate()
   Set mMenus = Nothing
   Set msRs = Nothing
End Sub
Public Function LlenarMenu()
    Dim PosMenu As String
    Dim NivelMenu As String
    Dim GrupoMenu As String
    Dim SubGrupoMenu As String
    Dim IndexMenu01 As Integer
    Dim IndexMenu02 As Integer
    Dim IndexMenu03 As Integer
    Dim IndexMenu04 As Integer
    IndexMenu01 = 0
    IndexMenu02 = 0
    IndexMenu03 = 0
    IndexMenu04 = 0
    With msRs
        Do While Not .EOF
            PosMenu = .Fields("POS")
            NivelMenu = .Fields("NIVEL")
            GrupoMenu = .Fields("GRUPO")
            SubGrupoMenu = .Fields("SUBGRUPO")
            If .Fields("SUBNIVEL") = "S" Then
                Select Case NivelMenu
                    Case "01": IndexMenu01 = 0
                    Case "02": IndexMenu03 = 0: IndexMenu02 = IndexMenu02 + IIf(IndexMenu02 Mod 2 = 0, 0, 1)
                    Case "03": IndexMenu03 = IndexMenu03 + IIf(IndexMenu03 Mod 2 = 0, 0, 1)
                    Case "04":
                End Select
            Else
                Select Case NivelMenu
                    Case "01": IndexMenu01 = 0
                    Case "02": IndexMenu03 = 0: IndexMenu02 = IndexMenu02 + IIf(IndexMenu02 Mod 2 <> 0, 0, 1)
                    Case "03": IndexMenu04 = 0: If (PosMenu = "03" Or PosMenu = "04" Or PosMenu = "05") Then IndexMenu03 = IndexMenu03 + IIf(IndexMenu03 Mod 2 <> 0, 0, 1)
                    Case "04":
                End Select
            End If
            Select Case PosMenu
                Case "01"
                    Select Case NivelMenu
                        Case "01"
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_01nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_01nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "06"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_06(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "07"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_07(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "08"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_08(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "09"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_09(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "10"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_01nivel03_10(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                    End Select
                Case "02"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_02nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_02nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "06"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_06(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "07"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_07(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "08"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_08(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "09"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_09(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "10"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_02nivel03_10(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                    End Select
                Case "03"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_03nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_03nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "06"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_06(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "07"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_07(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "08"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_08(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "09"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_09(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "10"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_03nivel03_10(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                        Case "04"
                            Select Case GrupoMenu
                                Case "01"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "02"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "03"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "04"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "05"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "06"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "07"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "08"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "09"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "10"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_01_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_02_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_03_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_04_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_03nivel04_05_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                            End Select
                    End Select
                Case "04"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_04nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_04nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "06"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_06(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "07"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_07(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "08"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_08(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "09"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_09(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "10"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_04nivel03_10(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                        Case "04"
                            Select Case GrupoMenu
                                Case "01"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "02"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "03"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "04"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "05"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "06"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "07"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "08"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "09"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "10"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_01_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_02_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_03_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_04_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_04nivel04_05_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                            End Select
                    End Select
                Case "05"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_05nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_05nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "06"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_06(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "07"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_07(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "08"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_08(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "09"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_09(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "10"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_05nivel03_10(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                        Case "04"
                            Select Case GrupoMenu
                                Case "01"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_01(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "02"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_02(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "03"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_03(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "04"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_04(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "05"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_05(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "06"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_06(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "07"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_07(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "08"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_08(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "09"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_09(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                               Case "10"
                                    Select Case SubGrupoMenu
                                         Case "01"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_01_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "02"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_02_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "03"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_03_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "04"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_04_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                         Case "05"
                                            Add PosMenu, NivelMenu, IndexMenu04, mdiInicio.m_05nivel04_05_10(IndexMenu04), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                            IndexMenu04 = IndexMenu04 + 1
                                    End Select
                            End Select
                    End Select
                Case "06"
                    Select Case NivelMenu
                        Case "01"
                             IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_06nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_06nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_06nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_06nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_06nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_06nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_06nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                    End Select
                Case "07"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_07nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_07nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_07nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_07nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_07nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_07nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_07nivel03_05(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                            End Select
                    End Select
                Case "08"
                    Select Case NivelMenu
                        Case "01"
                            IndexMenu02 = 0
                            Add PosMenu, NivelMenu, IndexMenu01, mdiInicio.m_08nivel01(0), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                        Case "02"
                            Add PosMenu, NivelMenu, IndexMenu02, mdiInicio.m_08nivel02(IndexMenu02), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                            IndexMenu02 = IndexMenu02 + 1
                        Case "03"
                            Select Case GrupoMenu
                                Case "01"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_08nivel03_01(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "02"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_08nivel03_02(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "03"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_08nivel03_03(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "04"
                                    Add PosMenu, NivelMenu, IndexMenu03, mdiInicio.m_08nivel03_04(IndexMenu03), .Fields("DESCRIP"), IIf(.Fields("Visible") = 0, False, True), GrupoMenu, SubGrupoMenu
                                    IndexMenu03 = IndexMenu03 + 1
                                Case "05"
                            End Select
                    End Select
            End Select
            .MoveNext
        Loop
    End With
End Function
