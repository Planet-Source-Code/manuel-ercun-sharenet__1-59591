Attribute VB_Name = "Module1"
Option Explicit

  ' DWORD WNetOpenEnum(

   ' DWORD dwScope,  // scope of enumeration
   ' DWORD dwType,   // resource types to list
  '  DWORD dwUsage,  // resource usage to list
  '  LPNETRESOURCE lpNetResource,    // pointer to resource structure
  '  LPHANDLE lphEnum    // pointer to enumeration handle buffer
  ' );



   'DWORD WNetEnumResource(

   ' HANDLE hEnum,   // handle to enumeration
   ' LPDWORD lpcCount,   // pointer to entries to list
   ' LPVOID lpBuffer,    // pointer to buffer for results
   ' LPDWORD lpBufferSize    // pointer to buffer size variable
  ' );



 ' typedef struct _NETRESOURCE {  // nr
 '   DWORD  dwScope;
 '   DWORD  dwType;
 '   DWORD  dwDisplayType;
 '   DWORD  dwUsage;
 '   LPTSTR lpLocalName;
 '   LPTSTR lpRemoteName;
 '   LPTSTR lpComment;
 '   LPTSTR lpProvider;
'} NETRESOURCE;


'DWORD WNetGetLastError(

'    LPDWORD lpError,    // pointer to error code
'    LPTSTR lpErrorBuf,  // pointer to string describing error
'    DWORD nErrorBufSize,    // size of description buffer, in characters
'    LPTSTR lpNameBuf,   // pointer to buffer for provider name
'   DWORD nNameBufSize  // size of provider name buffer
'   );




Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, ByRef lpNetResource As NETRESOURCE, ByRef lphEnum As Long) As Long


Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, ByRef lpcCount As Long, ByRef lpbuffer As Any, ByRef lpBufferSize As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (ByRef lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long
'Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)




Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long


Public Type NETRESOURCE
    dwScope As Long      '
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Type NETRESOURCELONG
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As Long
        lpRemoteName As Long
        lpComment As Long
        lpProvider As Long
End Type



Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpbuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const RESOURCE_CONNECTED As Long = &H1
Public Const RESOURCETYPE_ANY As Long = &H0
Public Const RESOURCEUSAGE_CONTAINER As Long = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE As Long = &H3
Public Const RESOURCE_GLOBALNET As Long = &H2
Public Const RESOURCETYPE_DISK As Long = &H1
Public Const RESOURCEUSAGE_CONNECTABLE As Long = &H1
















Public Function LastErrorApi(errordll As Long) As String
Dim res As Long
Dim s As String
s = String(255, vbNullChar)
res = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, errordll, 0, s, Len(s), 255)
If res <> 0 Then LastErrorApi = Left(s, res)

End Function

Public Function WNetLastErrorApi(errordll As Long) As String
Dim res As Long
Dim s As String

s = String(255, vbNullChar)
res = WNetGetLastError(errordll, s, Len(s), vbNullString, 0&)
If res = 0 Then WNetLastErrorApi = Left(s, InStr(s, vbNullChar) - 1)
End Function

Public Function Asciipoiter(resorce As Long) As String
Dim carpetas As String

carpetas = Space(lstrlen(resorce))
lstrcpy carpetas, resorce
If carpetas <> vbNullChar Then Asciipoiter = Left(carpetas, (Len(carpetas)))

End Function


Public Function ShareEnum(server As String) As Boolean
Dim res As Long
Dim nethandle As Long
Dim netres As NETRESOURCE
Dim lpbuffer() As NETRESOURCELONG
Dim str As String
Dim str1 As String




netres.dwScope = RESOURCE_GLOBALNET
netres.dwType = RESOURCETYPE_ANY
netres.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
netres.dwUsage = RESOURCEUSAGE_CONTAINER
'netres.lpLocalName = "x:"
netres.lpRemoteName = server
netres.lpComment = "AAAA"

res = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, 0, netres, nethandle)
If res <> 0 Then ShareEnum = False: Form2.List1.AddItem server & ":Err WnetopenEnum: " & Err.LastDllError & " " & LastErrorApi(res) & " " & WNetLastErrorApi(res): Exit Function
ReDim lpbuffer(1000)

res = WNetEnumResource(nethandle, &HFFFFFFFF, lpbuffer(0), 10000)
If res <> 0 Then ShareEnum = False: Form2.List1.AddItem server & ":Err WNetEnumResource: " & Err.LastDllError & " " & LastErrorApi(res) & " " & WNetLastErrorApi(Err.LastDllError): Exit Function

Set a = Form1.TreeView1.Nodes.Add(, , , server, 2)

For i = 0 To 100

If lpbuffer(i).lpRemoteName <> 0 Then
 str = Asciipoiter(lpbuffer(i).lpRemoteName)

End If

 If lpbuffer(i).lpComment <> 0 Then
 str1 = Asciipoiter(lpbuffer(i).lpComment)
 Form1.TreeView1.Nodes.Add a, tvwChild, , Mid(Mid(str, 3) & "(" & str1 & ")", InStr(Mid(str, 3) & "(" & str1, "\") + 1), 1
 x1 = x1 + 1
 End If
 
 
 

Next i

ShareEnum = True

WNetCloseEnum nethandle

End Function

