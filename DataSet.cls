VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class DataSet

'Implements AbstractNode


Private name As String
Private value As String

Property Let AbstractNode_propertyName(nm As String)
    name = nm
End Property

Property Get AbstractNode_propertyName() As String
    propertyName = name
End Property

Property Let propertyValue(vl As String)
    value = vl
End Property

Property Get propertyValue() As String
    propertyValue = value
End Property

Public Function isLeaf() As Boolean
    isLeaf = True
End Function

Public Sub add(nm As String, vl As String) ' ref to this node.
End Sub

Public Function getLeaves() As DataSets
    Err.Raise vbObjectError + 513, , "No leaves in base DataSet class"
End Function

Public Sub addDataSet(ds As DataSet)
    Err.Raise vbObjectError + 513, , "Not a DataSets collention class to add to"
End Sub

Public Function getChild(nm As String) As DataSets
    Err.Raise vbObjectError + 513, , "Not a collection. No Children here."
End Function
