VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "V9Params"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParameters
'Class V9Params

Private enode9 As ENode
Private cnvrtblHsh As HASHTABLE

Public Sub IParameters_InitParameters()
    With enode9.Children
        .AddNode "RunHidden"
        .AddNode "DMSServerURL"
        .AddNode "DMSSiteName"
        .AddNode "DMSUserName"
        .AddNode "DMSPassword"
        .AddNode "ColumnHeaderFile"
        .AddNode "DataFile"
        .AddNode "TableName"
        .AddNode "RecordFilter"
        .AddNode "DeleteDataFile"
        .AddNode "ImageRoot"
        .AddNode "Margin"
        .AddNode "BCHeight"
        .AddNode "BCWidth"
        .AddNode "BCCaption"
        .AddNode "LabelPrinter"
        .AddNode "DefaultDeleteImage"
        .AddNode "DefaultFolderName"
        .AddNode "DefaultMakeBarCode"
        .AddNode "DefaultReprint"
        .AddNode "DefaultLaunchWorkFlow"
        .AddNode "DefaultImageExtension"
        .AddNode "Printers"
        '.AddNode "Printer"
        .AddNode "MappedFields"
        .AddNode "StaticFields"
    End With
    
    cnvrtblHsh.Add "DocumentRoot", "ImageRoot"
    cnvrtblHsh.Add "BarCodeWidth", "BCWidth"
    cnvrtblHsh.Add "BarCodeHeight", "BCHeight"
    cnvrtblHsh.Add "CaptionText", "BCCaption"
    cnvrtblHsh.Add "PrintBarCode", "DefaultMakeBarCode"
    cnvrtblHsh.Add "DeleteDocument", "DefaultDeleteImage"
    cnvrtblHsh.Add "DMSFolderName", "DefaultFolderName"
    cnvrtblHsh.Add "LaunchWorkFlow", "DefaultLaunchWorkFlow"
    cnvrtblHsh.Add "DocumentExtension", "DefaultImageExtension"
    
End Sub

Public Function IParameters_getENode() As ENode
    Set IParameters_getENode = enode9
End Function

Public Function IParameters_getHash() As HASHTABLE
    Set IParameters_getHash = cnvrtblHsh
End Function

Private Sub Class_Initialize()
    Set enode9 = New ENode
    Set cnvrtblHsh = New HASHTABLE
    cnvrtblHsh.SetSize (100)
    cnvrtblHsh.IgnoreCase = True
    Me.IParameters_InitParameters
End Sub

