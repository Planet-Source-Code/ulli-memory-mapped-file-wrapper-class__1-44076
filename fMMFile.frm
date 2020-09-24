VERSION 5.00
Begin VB.Form fMMFile 
   Caption         =   "Test MMF Class"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "fMMFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Full featured example of using a Memory Mapped File class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'A Memory Mapped File resides in virtual memory from the moment it is opened until it is closed.
'This makes access to it very fast, in fact read and write merely consist of memory moves.
'All parts of the file are accessible by means of a zero based offset and the length of
'the data chunks you can transfer is only limited by the file size itself. The system will
'swap out and in pages of virtual memory as necessary.
'
'Only when you close the file the data are "lazily" written to disc, and the Physical
'Write Operations are limited to the altered memory pages.
'
'When you open an MMF you are required to supply an estimate of how large the file is gonna
'be. You cannot exceed this estimate, ie you can't extend a file while it's is open, but you
'can truncate the file size to the actual size used when you close it. You can however extend
'an existing file by giving the appropriate estimate when you open it and truncating it to
'the larger size when closing.
'
'MMFs are byte oriented and that requires understanding the data types which VB uses
'and how they are represented physically. The class contains tools however to handle the
'different aspects.

'It is up to you whether you want to use and react to all the error returns. A zero return
'indicates success in all cases.

Private Const FileName        As String = "\ULLIMMTEXT.TXT"

Private Sub btGo_Click()

  Dim MMFile            As New cMMFile 'memory mapped file class

  Dim TestWriteString   As String
  Dim TestReadString    As String

  Dim TestWriteLong     As Long
  Dim TestReadLong      As Long

  Dim LenChunk          As Long
  Const Estimate        As Long = 512

  Dim OpenErr           As OpenErr
  Dim CloseErr          As CloseErr
  Dim ReadErr           As ReadErr
  Dim WriteErr          As WriteErr

    Cls

    TestWriteString = "ABCDEF"
    LenChunk = LenB(TestWriteString)  'LenB returns the length in bytes, not in characters (remember VB uses it's unicode)

    With MMFile

        'open
        OpenErr = .OpenMMFile(App.Path & FileName, OpenAlways, Estimate)
        Select Case OpenErr
          Case CreateMapFailed
            MsgBox "Create Map Failed"
          Case MapViewFailed
            MsgBox "Map View Failed"
          Case OpenFileFailed
            MsgBox "Open File Failed"
          Case OpenParamErr
            MsgBox "Open Param Error"

          Case NoOpenFail
            'write chunk
            WriteErr = .WriteMMFile(0, StrPtr(TestWriteString), LenChunk)
            Select Case WriteErr
              Case WriteOffsetFailed
                MsgBox "Write Offset Failed"
              Case WriteParamErr
                MsgBox "Write Param Error"
            End Select

            'write chunk again in lower case
            WriteErr = .WriteMMFile(.CurrentFileSize, StrPtr(LCase$(TestWriteString)), LenChunk)
            Select Case WriteErr
              Case WriteOffsetFailed
                MsgBox "Write Offset Failed"
              Case WriteParamErr
                MsgBox "Write Param Error"
            End Select

            'create a string to receive two chunks
            TestReadString = .CreateString(LenChunk * 2)

            'read the two chunks
            ReadErr = .ReadMMFile(0, StrPtr(TestReadString), LenChunk * 2)
            Select Case ReadErr
              Case ReadOffsetFailed
                MsgBox "Read Offset Failed"
              Case ReadParamErr
                MsgBox "Read Param Error"
            End Select

            'display the two chunks
            Print TestReadString

            'write a long
            TestWriteLong = &H410042 'that will be displayed as BA (unicode again and little endian!)
            WriteErr = .WriteMMFile(.CurrentFileSize, VarPtr(TestWriteLong), LenB(TestWriteLong))
            Select Case WriteErr
              Case WriteOffsetFailed
                MsgBox "Write Offset Failed"
              Case WriteParamErr
                MsgBox "Write Param Error"
            End Select

            'create a string to receive the whole file (two chunks and a long)
            TestReadString = .CreateString(.CurrentFileSize)

            'read the two chunks and the long
            ReadErr = .ReadMMFile(0, StrPtr(TestReadString), .CurrentFileSize)
            Select Case ReadErr
              Case ReadOffsetFailed
                MsgBox "Read Offset Failed"
              Case ReadParamErr
                MsgBox "Read Param Error"
            End Select

            'display
            Print TestReadString

            'read the long alone(offset from end of file)
            ReadErr = .ReadMMFile(-(Estimate - .CurrentFileSize + LenB(TestReadLong)), VarPtr(TestReadLong), LenB(TestReadLong))
            Select Case ReadErr
              Case ReadOffsetFailed
                MsgBox "Read Offset Failed"
              Case ReadParamErr
                MsgBox "Read Param Error"
            End Select

            'display
            Print Hex$(TestReadLong)

            'close
            CloseErr = .CloseMMFile(.CurrentFileSize) 'limit file length
            Select Case CloseErr
              Case CloseFileFailed
                MsgBox "Close File Failed"
              Case CloseMapFailed
                MsgBox "Close Map Failed"
              Case CloseParamErr
                MsgBox "Close Param Error"
              Case MovePointerFailed
                MsgBox "Move Pointer Failed"
              Case TruncFileFailed
                MsgBox "Trunc File Failed"
              Case UnmapFailed
                MsgBox "Unmap Failed"
            End Select
        End Select

    End With 'MMFILE

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Shell "Notepad.exe " & App.Path & FileName, vbNormalFocus

End Sub

':) Ulli's VB Code Formatter V2.16.11 (2003-Mrz-17 19:54) 28 + 138 = 166 Lines
