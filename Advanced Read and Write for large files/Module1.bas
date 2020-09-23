Attribute VB_Name = "Module1"
Option Explicit

Const MOVEFILE_REPLACE_EXISTING = &H1
Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const OPEN_EXISTING = 3
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Type FILETIME ' 8 Bytes
  dwLowDateTime  As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA ' 318 Bytes
  dwFileAttributes     As Long
  ftCreationTime   As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime  As FILETIME
  nFileSizeHigh        As Long
  nFileSizeLow         As Long
  dwReserved_          As Long
  dwReserved1          As Long
  cFileName    As String * 260
  cAlternate    As String * 14
End Type

Public TotalBytes   As Variant
Public CurrentBytes As Variant
Public CurrentRead  As Variant
Public bBytes()     As Byte

Private Function vFileSize(strFileName As String) As Variant
Dim lpFindFileData As WIN32_FIND_DATA
Dim hFindFirst                As Long
  'get the file
  hFindFirst = FindFirstFile(strFileName, lpFindFileData)
    If hFindFirst > 0 Then
      FindClose hFindFirst
        'get the files size
        vFileSize = FileLength(lpFindFileData.nFileSizeHigh, lpFindFileData.nFileSizeLow)
    Else
      vFileSize = 0
    End If
End Function

Private Function FileLength(ByRef lFileSizeHigh As Long, ByRef lFileSizeLow As Long) As Variant
Static cBit  As Variant
Dim lFile    As Variant
  'If IsEmpty(cBit) Then cBit = CDec(2 ^ 32)
    cBit = CDec(2 ^ 32)
    lFile = CDec(0)
  If lFileSizeHigh < 0 Then
    lFile = (cBit + CDec(lFileSizeHigh)) * cBit
  Else
    lFile = CDec(lFileSizeHigh) * cBit
  End If
  If lFileSizeLow < 0 Then
    FileLength = lFile + cBit + CDec(lFileSizeLow)
  Else
    FileLength = lFile + CDec(lFileSizeLow)
  End If
End Function

Public Sub FileCopy(sFile As String, nsFile As String, BufferSize As Long)
Dim sSave  As String
Dim hOrgFile As Long
Dim hNewFile As Long
Dim ret      As Long
Dim sTemp  As String
Dim nSize    As Long
Dim lFile As Variant
Dim Lpos  As Variant
Dim Calc  As Variant
Dim RFile As Boolean
  'deafult out some vars we use
  RFile = True
  Lpos = 0
  CurrentRead = 0
  CurrentBytes = 0
  TotalBytes = 0
  sTemp = String(260, 0)
  'Get a temporary filename
  lFile = vFileSize(sFile)
  TotalBytes = lFile
    GetTempFileName App.Path & "\", "QQ-", 0, sTemp
  'Remove all the unnecessary chr$(0)'s
  sTemp = Left(sTemp, InStr(1, sTemp, Chr$(0)) - 1)
    'Set the file attributes
    SetFileAttributes sTemp, FILE_ATTRIBUTE_TEMPORARY
  'Open the files
  hNewFile = CreateFile(sTemp, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
  hOrgFile = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
     'sets buffer size by setting how big of a byte array to use when reading a file
      ReDim bBytes(BufferSize) As Byte
        While RFile = True
DoEvents
          Lpos = Lpos + UBound(bBytes)
            If Lpos >= lFile Then 'if whats left of the file is equal to or less than our buffer size
              Calc = (Lpos - lFile)
              Calc = (UBound(bBytes) - Calc)
              CurrentBytes = Calc
                ReadFile hOrgFile, bBytes(1), Calc, ret, ByVal 0&
                WriteFile hNewFile, bBytes(1), Calc, ret, ByVal 0&
              RFile = False
            Else 'if whats left of the file is greater than the size of our buffer
              CurrentBytes = Lpos
                ReadFile hOrgFile, bBytes(1), UBound(bBytes), ret, ByVal 0&
                WriteFile hNewFile, bBytes(1), UBound(bBytes), ret, ByVal 0&
            End If
        CurrentRead = CurrentRead + BufferSize
      Wend
    'Close the files now that we our done reading/writing
    CloseHandle hOrgFile
    CloseHandle hNewFile
    'Move the file
    MoveFileEx sTemp, nsFile, MOVEFILE_REPLACE_EXISTING
  CurrentRead = 0
End Sub

Public Function GetFileString(Cd As CommonDialog, Title As String, OpenDialog As Boolean, Optional FileStr As String = "") As String
On Error GoTo err
  With Cd
    .FileName = FileStr$
    .DialogTitle = Title$
    .Filter = "All Supported Types|*.*"
      Select Case OpenDialog
        Case True
          .ShowOpen
        Case False
          .ShowSave
      End Select
    GetFileString$ = .FileName
  End With
err:
End Function
