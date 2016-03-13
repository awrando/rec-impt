VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "v93Params"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParameters
'Class V93Params

Private enode93 As ENode
Private cnvrtblHsh As HASHTABLE

Private Function IParameters_getHash() As HASHTABLE
    Set cnvrtblHsh = New HASHTABLE
    cnvrtblHsh.SetSize (1000)
    IParameters_getHash = cnvrtblHsh
End Function

Public Sub IParameters_InitParameters()
    With enode93.Children
        .AddNode "ConnectionString"
        .AddNode "TableName"
        .AddNode "RecordFilter"
        .AddNode "SelectString"
        .AddNode "ColumnHeaderFile"
        .AddNode "DocumentRoot"
        .AddNode "DeleteDocument"
        .AddNode "DocumentExtension"
        .AddNode "DMSServerURL"
        .AddNode "DMSSiteName"
        .AddNode "DMSUserName"
        .AddNode "DMSPassword"
        .AddNode "DMSFolderName"
        .AddNode "LaunchWorkFlow"
        .AddNode "PrintBarCode"
        .AddNode "PrinterName"
        .AddNode "LabelPrinter"
        .AddNode "PaperBin"
        .AddNode "PaperSize"
        .AddNode "Orientation"
        .AddNode "LabelSheet"
        .AddNode "LabelSheetRows"
        .AddNode "LabelSheetMarginTop"
        .AddNode "LabelSheetMarginLeft"
        .AddNode "LabelSheetHorizontalSpacing"
        .AddNode "LabelSheetVerticalSpacing"
        .AddNode "LabelSheetLabelWidth"
        .AddNode "LabelSheetLabelHeight"
        .AddNode "BarCodeTop"
        .AddNode "BarCodeLeft"
        .AddNode "BarCodeWidth"
        .AddNode "BarCodeHeight"
        .AddNode "BarCodeType"
        .AddNode "BarCodeValue"
        .AddNode "CaptionTop"
        .AddNode "CaptionLeft"
        .AddNode "CaptionWidth"
        .AddNode "CaptionHeight"
        .AddNode "CaptionText"
        .AddNode "CaptionFontName"
        .AddNode "CaptionFontSize"
        .AddNode "KeepLogs"
        .AddNode "EndOnError"
        .AddNode "MappedFields"
        .AddNode "StaticFields"
    End With
End Sub

Public Function IParameters_getENode() As ENode
    Set IParameters_getENode = enode93
End Function

Private Sub Class_Initialize()
    Set enode93 = New ENode
    Me.IParameters_InitParameters
End Sub
