VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigScaffolds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class ConfigScaffolds

Private enode9 As IParameters
Private enode93 As IParameters
Private supportedEnodes As Collection
Private cfg As Integer

Public Function getScaffold(v As Integer) As ENode
    Select Case v
    Case 1
        cfg = v
        Set getScaffold = enode9.getENode
    Case 2
        cfg = v
        Set getScaffold = enode93.getENode
    Case Else
        cfg = v
        Set getScaffold = enode93.getENode
    End Select
End Function

Public Function getScaffoldHash() As HASHTABLE
    If cfg = 1 Then
        Set getScaffoldHash = enode9.getHash
    Else
        Set getScaffoldHash = New HASHTABLE
    End If
End Function

Public Function getConfigVersion(unknownenode As ENode) As Integer
    Dim unknowName As String
    Dim knownName As String
    Dim Count As Integer
    Dim supportedInt As Integer
    Dim unknownInt As Integer
    Dim currentE As ENode
    
    For supportedInt = 1 To supportedEnodes.Count '1 based else out of range
        Dim supportedEnodeName As String
        Dim unknownEnodeName As String
        
        Count = 0
        Set currentE = supportedEnodes(supportedInt)
    
        For unknownInt = 1 To currentE.Children.Count '1 based else out of range
            supportedEnodeName = currentE.Children.GetNodeByIndex(unknownInt).Name
            
            On Error GoTo fxUnkwnEName 'stops app from crashing when getnodebyname can't find a var
            unknownEnodeName = unknownenode.Children.GetNodeByName(supportedEnodeName).Name
            
            'strcomp()=0 will chk if a string is true or false in a traditional boolean format
            If (StrComp(supportedEnodeName, unknownEnodeName, vbTextCompare)) = 0 Then
                Count = Count + 1
            End If
        Next
                
        If Count = 26 Then
            getConfigVersion = VERSION9
            Exit For
        ElseIf Count = 45 Then
            getConfigVersion = VERSION93
            Exit For
        End If
    Next
    
fxUnkwnEName:
    unknownEnodeName = ""
Resume Next
End Function

Private Sub Class_Initialize()
    Set enode9 = New V9Params
    Set enode93 = New v93Params

    Set supportedEnodes = New Collection

    supportedEnodes.Add enode9.getENode
    supportedEnodes.Add enode93.getENode
End Sub
