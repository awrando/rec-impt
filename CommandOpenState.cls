VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenStateCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Command
'Class OpenStateCommand

Private med As Mediator

Public Sub init(md As Mediator)
Set med = md
med.registerOpenState OpenState
End Sub

Private Sub Command_Execute()
med.OpenState
End Sub

