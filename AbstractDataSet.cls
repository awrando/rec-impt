VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AbstractNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class AbstractNode
'Interface for all Enode classes

Public Property Get propertyName() As String
End Property

Public Property Let propertyName(nm As String)
End Property

Public Function isLeaf() As Boolean
End Function

Public Sub add(nm As String, vl As String) ' ref to this node.
End Sub

Public Function getLeaves() As ENodeCollection
End Function

Public Sub addChild(en As ENode)
End Sub

Public Function getChild(nm As String) As ENode
End Function




