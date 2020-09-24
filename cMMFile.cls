VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMMFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************************
' Copyright © 2003 UMGEDV GmbH
'
' Data handling and storage is a sensitive area of programming and there are people out
' there who sue for the craziest reasons; so I have to make the following statements and
' you, by using this code, do agree to abide by them.
'
' ******************************************************************************
' *  This software is provided "AS IS" and any express or implied warranties,  *
' *  including, but not limited to, the implied warranties of merchantability  *
' *  or fitness for a particular purpose, are disclaimed. In no event shall    *
' *  the author be liable for any direct, indirect, incidental, special,       *
' *  exemplary, or consequential damages (including, but not limited to,       *
' *  procurement of substitute goods or services; loss of use, data, or        *
' *  profits; or business interruption) however caused and on any theory of    *
' *  liability, whether in contract, strict liability, or tort (including      *
' *  negligence or otherwise) arising in any way out of the use of, or the     *
' *  impossibility to use, this software, even if advised of the possibility   *
' *  of such damage.                                                           *
' ******************************************************************************
'
' The following restrictions apply to this program code:
'
' 1 You may use this code in non-commercial applications for your own personal use.
'   If you want to use this code in a commercial application, prior permission
'   in writing by the author is required.
'   You can contact me via email at umgedv@aol.com
'
' 2 These notes and restrictions must not be removed from this file.
'********************************************************************************************

'A Memory Mapped File resides in virtual memory from the moment it is opened until it is closed.
'This makes access to it very fast, in fact read and write merely consist of data moves in memory.
'All parts of the file are accessible by means of a zero based offset and the length of
'the data chunks you can transfer is only limited by the file size itself. The system will
'swap out and in pages of virtual memory as necessary.
'
'Read operations take about 2.5 µsecs and write operations about 25 µsecs on my AMD Athlon XP1800+
'
'Only when you close the file the data are "lazily" written to disc, and the Physical
'Write Operations are limited to the altered memory pages.
'
'When you open an MMF you are required to supply an estimate of how large the file is gonna
'be. Any reasonable estimate will be fine, however, you cannot exceed this estimate, ie you
'can't extend a file while it is open, but you can truncate the file size to the actual
'size used when you close it. You can also extend the file during closing but the contents
'of the extension will be unpredictable. You can extend an existing file by giving the appropriate
'estimate when you open it, write data into the extension and then truncate it to the larger
'size when closing. Btw, your estimate is rounded up to the next memory allocation granularity
'multiple.
'
'MMFs are byte oriented and that requires a bit of understanding the data types which VB uses
'and how they are represented physically. The class contains a tool however to allocate
'Strings of bytes rather than unicode characters.

'It is up to you whether you want to use and react to all the error returns. A zero return
'indicates success in all cases.
'********************************************************************************************
'
'Development History
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'17Mar2003   V1.0.0        UMG
'
'Prototype
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
    dwOemID                             As Long
    dwPageSize                          As Long
    lpMinimumApplicationAddress         As Long
    lpMaximumApplicationAddress         As Long
    dwActiveProcessorMask               As Long
    dwNumberOrfProcessors               As Long
    dwProcessorType                     As Long
    dwAllocationGranularity             As Long
    dwReserved                          As Long
End Type
Private SysInfo                         As SYSTEM_INFO

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function FlushViewOfFile Lib "kernel32" (lpBaseAddress As Any, ByVal dwNumberOfBytesToFlush As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ClearMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private Declare Function AllocMem Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal OleStr As Long, ByVal NumBytes As Long) As Long

Private Const SHARE_READ                As Long = 1
Private Const SHARE_WRITE               As Long = 2
Private Const RANDOM_ACCESS             As Long = &H10000000
Private Const GENERIC_WRITE             As Long = &H40000000
Private Const GENERIC_READ              As Long = &H80000000

Private Const NIL                       As Long = 0

Private Const CREATE_NEW                As Long = 1
Private Const CREATE_ALWAYS             As Long = 2
Private Const OPEN_EXISTING             As Long = 3

Private Const PAGE_READWRITE            As Long = 4

Private Const FILE_MAP_READ             As Long = 2
Private Const FILE_MAP_WRITE            As Long = 4

Private Const INVALID_HANDLE            As Long = -1

Private Const FROM_BOF                  As Long = 0

Public Enum OpenMode
    OpenNonExisting = CREATE_NEW 'file must not exist
    OpenExisting = OPEN_EXISTING 'file must exist
    OpenAlways = CREATE_ALWAYS   'if file exists it will be overwritten else it will be created
End Enum
#If False Then
Private OpenNonExisting, OpenExisting, OpenAlways
#End If

Public Enum OpenErr
    NoOpenFail = 0
    MapViewFailed = 1
    CreateMapFailed = 2
    OpenFileFailed = 3
    OpenParamErr = 9
End Enum
#If False Then
Private NoOpenFail, MapViewFailed, CreateMapFailed, OpenFailed, OpenParamError
#End If

Public Enum CloseErr
    NoCloseFail = 0
    UnmapFailed = 1
    CloseMapFailed = 2
    TruncFileFailed = 3
    MovePointerFailed = 4
    CloseFileFailed = 5
    CloseParamErr = 9
End Enum
#If False Then
Private NoCloseFail, UnmapFailed, CloseMapFailed, TruncFileFailed, MovePointerFailed, CloseFileFailed, CloseParamError
#End If

Public Enum ReadErr
    NoReadFail = 0
    ReadOffsetFailed = 1
    ReadParamErr = 9
End Enum
#If False Then
Private NoReadFail, ReadOffsetFailed, ReadParamErr
#End If

Public Enum WriteErr
    NoWriteFail = 0
    WriteOffsetFailed = 1
    WriteParamErr = 9
End Enum
#If False Then
Private NoWriteFail, WriteOffsetFailed, WriteParamErr
#End If

Private myFileName                      As String
Private myCurrentFileSize               As Long
Private myFileHandle                    As Long
Private myMapHandle                     As Long
Private myBaseAddress                   As Long
Private myPreliminarySize               As Long

Private Sub Class_Initialize()

    myCurrentFileSize = -1

End Sub

Public Function CloseMMFile(ByVal ActualSize As Long) As CloseErr

  'Release memory space and close both the mapped and the physical file.
  'The physical file will be truncated or extended to ActualSize. If you extend the file
  'the contents of the extension will be unpredictable.

  'Returns Zero if successful

    If ActualSize >= 0 Then
        If UnmapViewOfFile(myBaseAddress) Then 'flushing to disk okay
            If CloseHandle(myMapHandle) Then 'don't need the file view no more
                If SetFilePointer(myFileHandle, ActualSize, NIL, FROM_BOF) = ActualSize Then  'sets the file pointer for subsequent truncation
                    If SetEndOfFile(myFileHandle) Then 'trunc okay
                        If CloseHandle(myFileHandle) Then
                            CloseMMFile = NoCloseFail
                          Else 'CLOSEHANDLE(MYFILEHANDLE) = FALSE/0
                            CloseMMFile = CloseFileFailed
                        End If
                      Else 'SETENDOFFILE(MYFILEHANDLE) = FALSE/0
                        CloseMMFile = TruncFileFailed
                    End If
                  Else 'NOT SETFILEPOINTER(MYFILEHANDLE,...
                    CloseMMFile = MovePointerFailed
                End If
              Else 'CLOSEHANDLE(MYMAPHANDLE) = FALSE/0
                CloseMMFile = CloseMapFailed
            End If
          Else 'UNMAPVIEWOFFILE(MYBASEADDRESS) = FALSE/0
            CloseMMFile = UnmapFailed
        End If
      Else 'NOT ACTUALSIZE...
        CloseMMFile = CloseParamErr
    End If

End Function

Public Function CreateString(ByVal SizeInBytes As Long) As String

  'Creates a string with size in bytes, rounded up to the next even number.
  'Found this at PSC some time ago, tnx goes to the author.

    If SizeInBytes > 0 Then
        CopyMemory ByVal VarPtr(CreateString), AllocMem(NIL, SizeInBytes + (SizeInBytes And 1)), 4
      Else 'NOT SIZEINBYTES...
        CreateString = ""
    End If

End Function

Public Property Get CurrentFileSize() As Long

    If myCurrentFileSize = -1 Then
        CurrentFileSize = myPreliminarySize
      Else 'NOT MYCURRENTFILESIZE...
        CurrentFileSize = myCurrentFileSize
    End If

End Property

Public Property Get FileName() As String

    FileName = myFileName

End Property

Public Function OpenMMFile(FileName As String, ByVal Mode As OpenMode, ByVal PreliminarySize As Long) As OpenErr

  'Creates or opens a file of preliminary size and maps the whole file into my memory space

  'Returns Zero if successful

  Dim Rmndr As Long

    myFileName = FileName
    If Len(FileName) And PreliminarySize > 0 And (Mode = OpenAlways Or Mode = OpenExisting Or Mode = OpenNonExisting) Then 'we got a filename and a file size
        GetSystemInfo SysInfo
        With SysInfo
            Rmndr = PreliminarySize Mod .dwAllocationGranularity
            myPreliminarySize = PreliminarySize + .dwAllocationGranularity * Sgn(Rmndr) - Rmndr 'round up to granularity
        End With 'SYSINFO
        myFileHandle = CreateFile(FileName, GENERIC_WRITE Or GENERIC_READ, SHARE_READ Or SHARE_WRITE, ByVal NIL, Mode, RANDOM_ACCESS, NIL)
        If myFileHandle <> INVALID_HANDLE Then 'create file mapping, requesting rounded preliminary size as maximum...
            myMapHandle = CreateFileMapping(myFileHandle, ByVal NIL, PAGE_READWRITE, NIL, myPreliminarySize, vbNullString)
            If myMapHandle Then 'map the whole file into this process's address space...
                myBaseAddress = MapViewOfFile(myMapHandle, FILE_MAP_READ Or FILE_MAP_WRITE, NIL, NIL, myPreliminarySize)
                If myBaseAddress Then 'we have an address
                    OpenMMFile = NoOpenFail
                    If Mode = OpenNonExisting Then 'clear memory
                        ClearMemory ByVal myBaseAddress, myPreliminarySize
                    End If
                  Else 'MYBASEADDRESS = FALSE/0
                    OpenMMFile = MapViewFailed
                End If
              Else 'MYMAPHANDLE = FALSE/0
                OpenMMFile = CreateMapFailed
            End If
          Else 'NOT MYFILEHANDLE...
            OpenMMFile = OpenFileFailed
        End If
        myPreliminarySize = PreliminarySize 'reset to requested size
      Else 'NOT LEN(FILENAME)...
        OpenMMFile = OpenParamErr
    End If

End Function

Public Function ReadMMFile(ByVal Offset As Long, ByVal Into As Long, ByVal Length As Long) As ReadErr

  'Read chunk from a memory-mapped file.
  'Offset ist the chunk address within the file (zero based)
  '   If offset is positive or zero then the chunk is obtained by calculating it's position fron BOF
  '   else the chunk is obtained by calculating it's position fron EOF (which is the preliminary size)
  'Length is the length of the chunk in bytes.

  'Returns Zero if successful

    If Length > 0 Then
        If Offset < 0 Then
            If Offset + Length <= 0 Then
                CopyMemory ByVal Into, ByVal myBaseAddress + myPreliminarySize + Offset, Length
                ReadMMFile = NoReadFail
              Else 'NOT OFFSET...
                ReadMMFile = ReadOffsetFailed
            End If
          Else 'NOT OFFSET...
            If Offset + Length < myPreliminarySize Then
                CopyMemory ByVal Into, ByVal myBaseAddress + Offset, Length
                ReadMMFile = NoReadFail
              Else 'NOT OFFSET...
                ReadMMFile = ReadOffsetFailed
            End If
        End If
      Else 'NOT LENGTH...
        ReadMMFile = ReadParamErr
    End If

End Function

Public Sub SaveMMFile()

  'Since updating the underlying file is defered until the file mapping object is closed,
  'a hardware or power failure may make it impossible to write changed data to disk.
  'Use this Sub at regular intervals if you application is critical. However, calling
  'this Sub too often may result in heavy disk traffic.

    If myCurrentFileSize > 0 Then
        FlushViewOfFile myBaseAddress, myCurrentFileSize
    End If

End Sub

Public Function WriteMMFile(ByVal Offset As Long, ByVal From As Long, ByVal Length As Long) As WriteErr

  'Write chunk to a memory-mapped file.
  'Offset ist the chunk address within the file (zero based)
  '   If offset is positive or zero then the chunk is obtained by calculating it's position fron BOF
  '   else the chunk is obtained by calculating it's position fron EOF (which is the preliminary size)
  'Length is the length of the chunk in bytes.
  'The caller must use StrPtr (for Srings) and VarPtr (for all ohers) to supply the From parameter.

  'Returns Zero if successful

    If Length > 0 Then
        If Offset < 0 Then
            If Offset + Length <= 0 Then
                CopyMemory ByVal myBaseAddress + myPreliminarySize + Offset, ByVal From, Length
                WriteMMFile = NoWriteFail
                If myPreliminarySize + Offset + Length > myCurrentFileSize Then
                    myCurrentFileSize = myPreliminarySize + Offset + Length
                End If
              Else 'NOT OFFSET...
                WriteMMFile = WriteOffsetFailed
            End If
          Else 'NOT OFFSET...
            If Offset + Length < myPreliminarySize Then
                CopyMemory ByVal myBaseAddress + Offset, ByVal From, Length
                WriteMMFile = NoWriteFail
                If Offset + Length > myCurrentFileSize Then
                    myCurrentFileSize = Offset + Length
                End If
              Else 'NOT OFFSET...
                WriteMMFile = WriteOffsetFailed
            End If
        End If
      Else 'NOT LENGTH...
        WriteMMFile = WriteParamErr
    End If

End Function

':) Ulli's VB Code Formatter V2.16.11 (2003-Mrz-17 19:54) 175 + 198 = 373 Lines
