Attribute VB_Name = "modUDT"
Option Explicit

Public Type GlobalVariables
  
  ConfigFile                      As String
  LogPath                         As String
  
  DMSServerURL                    As String
  DMSSiteName                     As String
  DMSUserName                     As String
  DMSPassword                     As String
  DMSFolderName                   As String
  DMSQuery                        As String
  DMSServerConfig                 As LVDMSCore9.ServerConfig
  DirectConnect                   As Boolean
  
  
  
  HostSQLServerName               As String
  HostSQLUserName                 As String
  HostSQLPassword                 As String
  HostSQLWinAuth                  As Boolean
  HostSQLProvider                 As String
  HostForceTranslate              As Boolean
  HostConnectionString            As String
  HostTable                       As String
  HostQuery                       As String
  HostPreprocessingQuery          As String
  
  AddLinkedRecords                As Boolean
  CopyDMSFieldsToLinkedRecords    As Boolean
  CheckForExistingLinkedRecord    As Boolean
  QueryForExistingLinkedRecord    As String
    
  
  ContinueAfterError              As Boolean
 ' KeepLogs                        As Long
  
  Fields                          As UpdateFieldList
  
  RunBatch                        As Boolean
  RunHidden                       As Boolean
  
  CancelFlag                      As Boolean
  

  CaptionText                     As String
  Connection                      As String
  File                            As String
  FolderName                      As String
  KeepLogs                        As Integer
  Password                        As String
  'Recordset
  ServerURL                       As String
  SiteName                        As String
  SqlSelect                       As String
  TableName                       As String
  UserName                        As String
  
End Type

