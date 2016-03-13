VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xml As ENode
Private xmlRoot As ENode
Private xmlCollection As ENodeCollection
Private thisNode As ENode
Private position As Integer
Private name As String

Private Function hasNext() As Boolean 'uses position var from outside
    Dim max As Integer
    Dim min As Integer
    Dim Count As Integer
    
    min = 0
    max = xml.Children.Count
    
    If position < max Then
        hasNext = True
    Else
        hasNext = False
    End If
End Function

Private Function nextObject() As ENode
    Set nextObject = xml.Children.GetNodeByIndex(position)
    position = position + 1
End Function

'Public Property Let propertyName(nm As String)
'    xml.name = nm
'End Property
'
'Public Property Get propertyName() As String
'    propertyName = xml.name
'End Property

Private Function isLeaf() As Boolean
    If xmlRoot.name = "" Then
        isLeaf = True
    Else
        isLeaf = False
    End If
End Function

Public Sub setRoot(nm As String) ' ref to this node
    xmlRoot.name = nm
End Sub

Private Function getLeaves() As Object
    getLeaves = xmlRoot.Children
End Function

Public Sub addChild(en As ENode)
    xmlRoot.Children.add en
End Sub


Public Function printNodes() As String
    xmlRoot.FormatOutput = True
    printNodes = xmlRoot.XMLSaveToString
End Function

Private Sub Class_Initialize()
    Set xml = New ENode
    Set xmlRoot = New ENode
    position = 0
End Sub
