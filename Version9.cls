VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'CLASS VERSION9
Implements IConfigVersion

Private cnvrtblHsh As HASHTABLE

Public Sub IConfigVersion_InitParameters(x As ENode)
    x.Children.Clear
    Set cnvrtblHsh = New HASHTABLE
    cnvrtblHsh.SetSize (100)
    cnvrtblHsh.IgnoreCase = True
    
    With x.Children
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

Private Function IConfigVersion_getTypeCode() As Integer
    IConfigVersion_getTypeCode = VERSION9
End Function

Public Sub IConfigVersion_SaveFile(x As ENode)
    Dim c As Control
    Dim e As ENode
    Dim count As Long

    On Error GoTo saveErrHdlr
    For count = 1 To x.Children.count
        Set e = x.Children.GetNodeByIndex(count)
        
        If TypeName(RecordImport.Controls(e.Name)) = "Object" Then
            Set c = RecordImport.Controls(e.Name)(VERSION9)
        Else
            Set c = RecordImport.Controls(e.Name)
        End If
        
        Select Case TypeName(c) ' Bug fix. Clear c values so it doesn't cary to e's that don't have values.
           Case "TextBox"
              x.Children.GetNodeByName(e.Name).value = c.Text
              c.Text = ""
           Case "CheckBox"
              x.Children.GetNodeByName(e.Name).value = c.value
              c.value = ""
           Case "ComboBox"
              x.Children.GetNodeByName(e.Name).value = c.Text
              c.Text = ""
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

    For index = 1 To RecordImport.ListViewData9.ListItems.count
        If RecordImport.ListViewData9.ListItems.Item(index).Text = "Mapped Field" Then
            mf.Children.AddNode "MappedField"
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "ADOFieldName", RecordImport.ListViewData9.ListItems.Item(index).SubItems(fSource)
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "DMSFieldName", RecordImport.ListViewData9.ListItems.Item(index).SubItems(fDms)
            mf.Children.GetNodeByIndex(mfCount).Children.AddNode "FieldPurpose", RecordImport.ListViewData9.ListItems.Item(index).SubItems(fPurpose)
            mfCount = mfCount + 1
        End If
        If RecordImport.ListViewData9.ListItems.Item(index).Text = "Static Field" Then
            sf.Children.AddNode "StaticField"
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "DMSFieldName", RecordImport.ListViewData9.ListItems.Item(index).SubItems(2)
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "FieldValue", RecordImport.ListViewData9.ListItems.Item(index).SubItems(4)
            sf.Children.GetNodeByIndex(sfCount).Children.AddNode "FieldPurpose", RecordImport.ListViewData9.ListItems.Item(index).SubItems(3)
            sfCount = sfCount + 1
        End If
    Next index
    
    Dim ep As ENode
    Set ep = x.Children.GetNodeByName("Printers")
    
    ep.Children.Clear ' Fix bug. Duplicate printers would list if not cleared first.
    
    For index = 1 To RecordImport.ListViewPrint.ListItems.count
        ep.Children.AddNode ("Printer")
        ep.Children.GetNodeByIndex(index).Children.AddNode "PrinterName", RecordImport.ListViewPrint.ListItems(index).Text
        ep.Children.GetNodeByIndex(index).Children.AddNode "FolderName", RecordImport.ListViewPrint.ListItems(index).SubItems(1)
    Next index
    
    'can save if dialog is canceled. fix later
    RecordImport.CommonDialog1.ShowSave
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


