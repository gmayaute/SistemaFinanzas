VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cLine"
Attribute VB_Ext_KEY = "Member0" ,"cLine"
Option Explicit
Private mCol As Collection

Public Function Add(Lenght As Long, Value As String, Fields As Fields, Optional sKey As String) As cLine
On Error GoTo controlError
    Dim objNewMember As cLine
    Set objNewMember = New cLine
    objNewMember.Lenght = Lenght
    objNewMember.Value = Left(Right(Value, Len(Value) - 1), Len(Value) - 2)
    Set objNewMember.Fields = Fields
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing
controlError:
    If err.Number <> Empty Then
       Resume Next
    End If
End Function
Public Property Get Item(vntIndexKey As Variant) As cLine
Attribute Item.VB_UserMemId = 0
   Set Item = mCol(vntIndexKey)
End Property
Public Property Get Count() As Long
   Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property
Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub
