VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigModelA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CLASS CONFIGMODELA

Private x As ENode
Private t As ConfigType

Public Function getType() As Integer
    getType = t.getTypeCode
End Function

Public Sub SaveFile()
    x.Children.GetNodeByName("DMSUserName").encoding = EnodeEncoding_Base64
    x.Children.GetNodeByName("DMSPassword").encoding = EnodeEncoding_Base64
    
    t.SaveFile x
    
    Me.setSaveFileMenu
End Sub

Public Sub SaveFileAs()
    x.Children.GetNodeByName("DMSUserName").encoding = EnodeEncoding_Base64
    x.Children.GetNodeByName("DMSPassword").encoding = EnodeEncoding_Base64
    
    t.SaveFile x
    
    Me.setSaveFileMenu
End Sub

Public Sub newType(nt As Integer)
    t.newType nt, x
End Sub

Public Sub OpenFile(path As String)
    x.Children.Clear
    x.XMLLoadFromFile path
    Me.newType (Me.getConfigVersion(x)) 'blanks out gui
    
    x.XMLLoadFromFile path 'bug fix, load gui again after config
End Sub

Public Function getConfigNode() As ENode
    Set getConfigNode = x
End Function

Private Sub Class_Initialize()
    Set x = New ENode
    Set t = New ConfigType
    
    x.Name = "Config"
    x.FormatOutput = True
    
    t.newType VERSION93, x
    
    Me.setSaveFileMenu
End Sub

Public Function getConfigVersion(unknownenode As ENode) As Integer
    Dim unknowName As String
    Dim knownName As String
    Dim count As Integer
    Dim supportedInt As Integer
    Dim unknownInt As Integer
    Dim currentE As ENode
    Dim enode9 As IParameters
    Dim enode93 As IParameters
    Dim supportedEnodes As Collection
    
    Set enode9 = New V9Params
    Set enode93 = New v93Params

    Set supportedEnodes = New Collection

    supportedEnodes.Add enode9.getENode
    supportedEnodes.Add enode93.getENode
    
    For supportedInt = 1 To supportedEnodes.count '1 based else out of range
        Dim supportedEnodeName As String
        Dim unknownEnodeName As String
        
        count = 0
        Set currentE = supportedEnodes(supportedInt)
    
        For unknownInt = 1 To currentE.Children.count '1 based else out of range
            supportedEnodeName = currentE.Children.GetNodeByIndex(unknownInt).Name
            
            On Error GoTo fxUnkwnEName 'stops app from crashing when getnodebyname can't find a var
            unknownEnodeName = unknownenode.Children.GetNodeByName(supportedEnodeName).Name
            
            'strcomp()=0 will chk if a string is true or false in a traditional boolean format
            If (StrComp(supportedEnodeName, unknownEnodeName, vbTextCompare)) = 0 Then
                count = count + 1
            End If
        Next
                
        If count = 25 Then
            getConfigVersion = VERSION9
            Exit For
        ElseIf count = 45 Then
            getConfigVersion = VERSION93
            Exit For
        Else
            getConfigVersion = VERSION93
            Exit For
        End If
    Next
    
fxUnkwnEName:
    unknownEnodeName = ""
Resume Next
End Function

Public Sub setSaveFileMenu()
    If Len(RecordImport.CommonDialog1.FileName) <> 0 Then
        RecordImport.SaveFile.Enabled = True
    Else
        RecordImport.SaveFile.Enabled = False
    End If
End Sub
