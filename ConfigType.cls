VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CLASS CONFIGTYPE
Private t As IConfigVersion

Sub newType(typ As Integer, x As ENode)
    Select Case typ
        Case VERSION9
            Set t = New VERSION9
        Case VERSION93
            Set t = New VERSION93
        Case Else
            Set t = New VERSION93
            MsgBox "Unable to find version, defaulting to Config Version 9.3"
    End Select
    t.InitParameters x
End Sub

Public Function getTypeCode() As Integer
    getTypeCode = t.getTypeCode
End Function

Public Sub SaveFile(x As ENode)
    t.SaveFile x
End Sub
