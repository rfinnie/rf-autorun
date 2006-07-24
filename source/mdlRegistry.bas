Attribute VB_Name = "mdlRegistry"
   Private Const REG_SZ As Long = 1 'REG_SZ represents a fixed-length text string.
   Private Const REG_DWORD As Long = 4 'REG_DWORD represents data by a number that is 4 bytes long.

   Private Const HKEY_CLASSES_ROOT = &H80000000 'The information stored here ensures that the correct program opens when you open a file by using Windows Explorer.
   Private Const HKEY_CURRENT_USER = &H80000001 'Contains the root of the configuration information for the user who is currently logged on.
   Private Const HKEY_LOCAL_MACHINE = &H80000002 'Contains configuration information particular to the computer (for any user).
   Private Const HKEY_USERS = &H80000003 'Contains the root of all user profiles on the computer.

   'Return values for all registry functions
   Private Const ERROR_SUCCESS = 0
   Private Const ERROR_NONE = 0

   Private Const KEY_QUERY_VALUE = &H1 'Required to query the values of a registry key.
   Private Const KEY_ALL_ACCESS = &H3F 'Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.


'API Calls for writing to Registry
  'Close Registry Key
   Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  'Create Registry Key
   Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  'Open Registry Key
   Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  'Query a String Value
   Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
  'Query a Long Value
   Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
  'Query a NULL Value
   Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
  'Enumerate Sub Keys
   Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
  'Store a Value
   Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
  'Delete Key
   Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Sub SaveValue(hKey As Long, strPath As String, strvalue As String, strData As String)
    
   Dim ret
  'Create a new key
   RegCreateKey hKey, strPath, ret
  'Save a string to the key
   RegSetValueEx ret, strvalue, 0, REG_SZ, ByVal strData, Len(strData)
  'close the key
   RegCloseKey ret
    
End Sub

Function GetURLCommand(ByVal url As String)
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\HTTP\Shell\Open\command", 0, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, "", vValue)
    
    urlopener = vValue
    RegCloseKey (hKey) 'Close the Key
        
    If urlopener = "" Then
        urlcommand = "explorer " & Chr(34) & url & Chr(34)
    Else
        strpos = InStr(urlopener, "%1")
        If strpos > 0 Then
            urlcommand = Left(urlopener, strpos - 1) & url & Right(urlopener, Len(urlopener) - strpos - 1)
        Else
            urlcommand = urlopener & " " & Chr(34) & url & Chr(34)
        End If
    End If
    GetURLCommand = urlcommand
End Function

Private Sub QueryValue(sKeyName As String, sValueName As String)
       
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long         'handle of opened key
  Dim vValue As Variant      'setting of queried value

  lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_QUERY_VALUE, hKey) 'Open Key to Query a value
  lRetVal = QueryValueEx(hKey, sValueName, vValue) 'Query (determine) the value stored

  frmRegistry.Caption = vValue 'Set the Form's Caption to whatever text was stored
  RegCloseKey (hKey) 'Close the Key
       
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       
       Dim Data As Long
       Dim retval As Long 'Return value of RegQuery functions
       Dim lType As Long 'Determine data type of present data
       Dim lValue As Long 'Long value
       Dim sValue As String 'String value

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       retval = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, Data)
       
       If retval <> ERROR_NONE Then Error 5

       Select Case lType
           ' Determine strings
           Case REG_SZ:
               sValue = String(Data, 0)

               retval = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, Data)
               
               If retval = ERROR_NONE Then
                   vValue = Left$(sValue, Data - 1)
               Else
                   vValue = Empty
               End If
               
           ' Determine DWORDS
           Case REG_DWORD:
               retval = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, Data)
               
               If retval = ERROR_NONE Then vValue = lValue
           
           Case Else
               'all other data types not supported
               retval = -1
       End Select
    
QueryValueExError:
       QueryValueEx = retval
       Exit Function

   End Function

