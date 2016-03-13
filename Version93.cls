VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version93"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class Version93
Implements IConfigVersion


Public Function IConfigVersion_getTypeCode() As Integer
    IConfigVersion_getTypeCode = VERSION93
End Function

Public Sub IConfigVersion_InitParameters(x As ESCTree4.ENode)
    With x.Children
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

Public Sub IConfigVersion_SaveFile(x As ENode)
    'can save if dialog is canceled. fix later
    'RecordImport.CommonDialog1.ShowSave
    Me.saveData (x)
End Sub

Sub saveData(x As ENode)
    Dim c As Control
    Dim e As ENode
    Dim count As Long

    On Error GoTo saveErrHdlr
    For count = 1 To x.Children.count
        Set e = x.Children.GetNodeByIndex(count)
        
        If TypeName(RecordImport.Controls(e.Name)) = "Object" Then
            Set c = RecordImport.Controls(e.Name)(VERSION93)
        Else
            Set c = RecordImport.Controls(e.Name)
        End If
        
        Select Case TypeName(c)
           Case "TextBox"
              x.Children.GetNodeByName(e.Name).value = c.Text
           Case "CheckBox"
              x.Children.GetNodeByName(e.Name).value = c.value
           Case "ComboBox"
              x.Children.GetNodeByName(e.Name).value = c.Text
        End Select
    Next

    Dim index As Integer
    Dim sfCount As Integer
    Dim mfCount As Integer
    Dim currentCtrlName As String
    Dim currentCtrlVal As String
    Dim LI As ListItem
    Const fSource = 1
    Const fValue = 2
    Const fDms = 3
    Const fPurpose = 4


    Dim mf As ENode
    Dim sf As ENode

    Set mf = x.Children.GetNodeByName("MappedFields")
    Set sf = x.Children.GetNodeByName("StaticFields")

    mf.Children.Clear
    sf.Children.Clear
    mf.value = "" ' Bug fix. Select Case TypeName(c) keeps giving mf & sf "continue on error" values
    sf.value = ""

    sfCount = 1 'sf index
    mfCount = 1 'mf index

    For index = 1 To RecordImport.ListViewData93.ListItems.count
        If RecordImport.ListViewData93.ListItems.Item(index).Text = "Mapped Field" Then
            mf.Children.AddNode "MappedField"
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "ADOFieldName", RecordImport.ListViewData93.ListItems.Item(index).SubItems(fSource)
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "DMSFieldName", RecordImport.ListViewData93.ListItems.Item(index).SubItems(fDms)
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "FieldPurpose", RecordImport.ListViewData93.ListItems.Item(index).SubItems(fPurpose)
            mfCount = mfCount + 1
        End If
        If RecordImport.ListViewData93.ListItems.Item(index).Text = "Static Field" Then
            sf.Children.AddNode "StaticField"
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "DMSFieldName", RecordImport.ListViewData93.ListItems.Item(index).SubItems(2)
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "FieldValue", RecordImport.ListViewData93.ListItems.Item(index).SubItems(4)
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "FieldPurpose", RecordImport.ListViewData93.ListItems.Item(index).SubItems(3)
            sfCount = sfCount + 1
        End If
    Next index
     
    
    
    x.XMLSaveToFile RecordImport.CommonDialog1.FileName
    
    Exit Sub
    
saveErrHdlr:
    Select Case Err.Number
        Case 730
            ' Do nothing. 730 is hit when we can't find mapped or static fields.
            ' we've got functions to handle that
        Case Default
            MsgBox Err.Number & " " & Err.Description
    End Select
    
    Resume Next
End Sub
