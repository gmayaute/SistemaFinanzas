VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub NodeParentsCheck(Node As Node)
Dim sNode As Node
    Set sNode = Node.Parent
    sNode.Checked = Node.Checked
    If Not sNode.Parent Is Nothing Then NodeParentsCheck sNode
End Sub
Public Sub NodeChildrenCheck(Node As Node)
Dim sNode As Node
Dim iLoop As Integer
    Set sNode = Node.Child
    For iLoop = 1 To Node.Children
        sNode.Checked = Node.Checked
        If Node.Children Then NodeChildrenCheck sNode
        Set sNode = sNode.Next
    Next iLoop
End Sub

Public Sub NodeSelectedCheck(Node As Node)
Dim sNode       As Node
Dim iLoop       As Integer
Dim bFound      As Boolean
    Set sNode = Node.Child
    For iLoop = 1 To Node.Children
        If sNode.Checked Then bFound = True: Exit For
        Set sNode = sNode.Next
    Next iLoop
    
    If bFound = False Then
        Node.Checked = False
        If Not Node.Parent Is Nothing Then NodeSelectedCheck Node.Parent
    End If
    Set sNode = Nothing
End Sub
