VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REGADP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Private Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Private Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)
Private Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, lpDataBuff&, nSize&)

 Const HKEY_CLASSES_ROOT = &H80000000
 Const HKEY_CURRENT_USER = &H80000001
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const HKEY_USERS = &H80000003

 Const ERROR_SUCCESS = 0&
 Const REG_SZ = 1&                          ' Unicode nul terminated string
 Const REG_DWORD = 4&                       ' 32-bit number

 Const KEY_QUERY_VALUE = &H1&
 Const KEY_SET_VALUE = &H2&
 Const KEY_CREATE_SUB_KEY = &H4&
 Const KEY_ENUMERATE_SUB_KEYS = &H8&
 Const KEY_NOTIFY = &H10&
 Const KEY_CREATE_LINK = &H20&
 Const READ_CONTROL = &H20000
 Const WRITE_DAC = &H40000
 Const WRITE_OWNER = &H80000
 Const SYNCHRONIZE = &H100000
 Const STANDARD_RIGHTS_REQUIRED = &HF0000
 Const STANDARD_RIGHTS_READ = READ_CONTROL
 Const STANDARD_RIGHTS_WRITE = READ_CONTROL
 Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
 Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
 Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
 Const KEY_EXECUTE = KEY_READ

Function RegGetValue$(MainKey&, SubKey$, value$)
   ' MainKey must be one of the Publicly declared HKEY constants.
   Dim sKeyType&       'to return the key type.  This function expects REG_SZ or REG_DWORD
   Dim ret&            'returned by registry functions, should be 0&
   Dim lpHKey&         'return handle to opened key
   Dim lpcbData&       'length of data in returned string
   Dim ReturnedString$ 'returned string value
   Dim ReturnedLong&   'returned long value
   If MainKey >= &H80000000 And MainKey <= &H80000006 Then
      ' Open key
      ret = RegOpenKeyExA(MainKey, SubKey, 0&, KEY_READ, lpHKey)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""
         Exit Function     'No key open, so leave
      End If
      
      ' Set up buffer for data to be returned in.
      ' Adjust next value for larger buffers.
      lpcbData = 255
      ReturnedString = Space$(lpcbData)

      ' Read key
      ret& = RegQueryValueExA(lpHKey, value, ByVal 0&, sKeyType, ReturnedString, lpcbData)
      If ret <> ERROR_SUCCESS Then
         RegGetValue = ""   'Value probably doesn't exist
      Else
        If sKeyType = REG_DWORD Then
            ret = RegQueryValueEx(lpHKey, value, ByVal 0&, sKeyType, ReturnedLong, 4)
            If ret = ERROR_SUCCESS Then RegGetValue = CStr(ReturnedLong)
        Else
            RegGetValue = Left$(ReturnedString, lpcbData - 1)
        End If
    End If
      ' Always close opened keys.
      ret = RegCloseKey(lpHKey)
   End If
End Function
