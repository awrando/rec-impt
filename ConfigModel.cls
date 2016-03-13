VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class ConfigModel

Private configEnode As ENode
Private cScaff As ConfigScaffolds

Public Sub SaveFile()
End Sub

Private Sub Class_Initialize()
    Set cScaff = New ConfigScaffolds
    Set configEnode = cScaff.getScaffold(2)
End Sub

Public Function getEnode() As ENode
    Set getEnode = configEnode
End Function
