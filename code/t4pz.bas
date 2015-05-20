Attribute VB_Name = "t4pz"
 Option Explicit
 Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
 Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
 Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
 Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
 Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
 Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
 Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
 Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
 Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
 Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
     Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long

 
 Private Type LUID
 LowPart As Long
 HighPart As Long
 End Type

 Private Type LUID_AND_ATTRIBUTES
 pLuid As LUID
 Attributes As Long
 End Type

 Private Type TOKEN_PRIVILEGES
 PrivilegeCount As Long
 TheLuid As LUID
 Attributes As Long
 End Type


 Public Const MAX_PATH As Integer = 260
 Public Const TH32CS_SNAPPROCESS As Long = 2&

 Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szexeFile As String * MAX_PATH
 End Type


 Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

 Function ProcessTerminate(Optional lProcessID As Long, Optional lHwndWindow As Long) As Boolean
 Dim lhwndProcess As Long
 Dim lExitCode As Long
 Dim lRetVal As Long
 Dim lhThisProc As Long
 Dim lhTokenHandle As Long
 Dim tLuid As LUID
 Dim tTokenPriv As TOKEN_PRIVILEGES, tTokenPrivNew As TOKEN_PRIVILEGES
 Dim lBufferNeeded As Long

 Const PROCESS_ALL_ACCESS = &H1F0FFF, PROCESS_TERMINAT = &H1
 Const ANYSIZE_ARRAY = 1, TOKEN_ADJUST_PRIVILEGES = &H20
 Const TOKEN_QUERY = &H8, SE_DEBUG_NAME As String = "SeDebugPrivilege"
Const SE_PRIVILEGE_ENABLED = &H2

 On Error Resume Next
If lHwndWindow Then
 'Get the process ID from the window handle
 lRetVal = GetWindowThreadProcessId(lHwndWindow, lProcessID)
End If

If lProcessID Then
 'Give Kill permissions to this process
 lhThisProc = GetCurrentProcess

 OpenProcessToken lhThisProc, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lhTokenHandle
 LookupPrivilegeValue "", SE_DEBUG_NAME, tLuid
 'Set the number of privileges to be change
 tTokenPriv.PrivilegeCount = 1
 tTokenPriv.TheLuid = tLuid
 tTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
 'Enable the kill privilege in the access token of this process
 AdjustTokenPrivileges lhTokenHandle, False, tTokenPriv, Len(tTokenPrivNew), tTokenPrivNew, lBufferNeeded

 'Open the process to kill
 lhwndProcess = OpenProcess(PROCESS_TERMINAT, 0, lProcessID)

If lhwndProcess Then
 'Obtained process handle, kill the process
 ProcessTerminate = CBool(TerminateProcess(lhwndProcess, lExitCode))
 Call CloseHandle(lhwndProcess)
End If
End If
 On Error GoTo 0
 End Function

 Public Function KillProcessus(nom_process) As String
 Dim i As Integer
 Dim hSnapshot As Long
 Dim uProcess As PROCESSENTRY32
 Dim r As Long
 Dim nom(1 To 100)
 Dim num(1 To 100)
 Dim nr As Integer
 nr = 0
 hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
 If hSnapshot = 0 Then Exit Function
 uProcess.dwSize = Len(uProcess)
 r = ProcessFirst(hSnapshot, uProcess)
 Do While r
 nr = nr + 1
 nom(nr) = uProcess.szexeFile
 num(nr) = uProcess.th32ProcessID
 r = ProcessNext(hSnapshot, uProcess)
 Loop
 For i = 1 To nr
If InStr(UCase(nom(i)), UCase(nom_process)) <> 0 Then
 ProcessTerminate (num(i))
 Exit For
End If
 Next i
 End Function
 
 
            
    ' Main routine to Dimension variables, retrieve user name
     ' and display answer.
     Sub Get_User_Name()

     ' Dimension variables
     Dim lpBuff As String * 25
     Dim ret As Long, UserName As String

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUserName(lpBuff, 25)
     UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

     ' Display the User Name
Form1.userweak.Caption = UserName
     End Sub
