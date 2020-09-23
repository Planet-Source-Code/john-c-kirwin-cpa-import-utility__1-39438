VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//****************************************************************************
'// cCPA_Files - File manipulation class
'//
'// Version 6.20.0
'// Created:  03/03/2002 John C. Kirwin (JCK)
'// Modified: 03/03/2002 JCK - Format & Comment
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************

Private Declare Function CopyFile& Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long)
Private Declare Function DeleteFile& Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'local variable(s) to hold property value(s)
Private sWindowsPath As String
Private intFileID As Integer
Private strFilePath As String
Private strFileName As String
Private lngCurrentRow As Long
Private lngRowCount As Long

Public Enum enmWriteType
    AppendData = 0
    InsertData = 1
    ReplaceData = 2
End Enum

Public Sub WriteData(intWriteType As enmWriteType, strData As String, lRowID)

  '//**************************************************************************
  '// WriteData -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

  Dim intWriteFileID As Integer
  Dim lngCounter As Long
  Dim lngWriteToRowID As Long
  Dim strTemp As String

    '//****
    If (lRowID > lngRowCount) Or (intFileID = 0) Then Exit Sub

    '//****
    lngCounter = 0

    '//****
    intWriteFileID = FreeFile

    '//****
    Open strFilePath & "tmp" & strFileName For Output As intWriteFileID

    '//****
    If intWriteType = AppendData Then

        '//****
        lngWriteToRowID = lngRowCount

        '//****
      Else

        '//****
        lngWriteToRowID = lRowID + Abs((lngCurrentRow * (lRowID = 0))) - 1

        '//****
    End If

    '//****
    MoveFirst

    '//****
    Do While lngCounter < lngWriteToRowID

        '//****
        Line Input #intFileID, strTemp

        '//****
        Print #intWriteFileID, strTemp

        '//****
        lngCounter = lngCounter + 1

        '//****
    Loop

    '//****
    If Len(Trim$(strData)) > 0 Then Print #intWriteFileID, strData

    '//****
    If intWriteType = ReplaceData Then Line Input #intFileID, strTemp

    '//****
    Do While Not EOF(intFileID)

        '//****
        Line Input #intFileID, strTemp

        '//****
        Print #intWriteFileID, strTemp

        '//****
    Loop

    Close

    '//****
    Kill strFilePath & strFileName

    '//****
    Name strFilePath & "tmp" & strFileName As strFilePath & strFileName

    '//****
    OpenFile strFilePath, strFileName

    '//****
    lngCurrentRow = lngWriteToRowID + 1

    '//****
    MoveToRow lngWriteToRowID + 1

    '//****
    Close intFileID

End Sub

Public Sub OpenFile(strFileLocation As String, pstrFileName As String)

  '//**************************************************************************
  '// OpenFile -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    intFileID = FreeFile

    '//****
    If Right$(strFileLocation, 1) = "\" Then

        '//****
        strFilePath = strFileLocation

        '//****
      Else

        '//****
        strFilePath = strFileLocation & "\"

        '//****
    End If

    '//****
    strFileName = pstrFileName

    '//****
    If Len(Dir$(strFilePath & strFileName)) > 0 Then SetRowCount

    '//****

End Sub

Public Sub MoveFirst()

  '//**************************************************************************
  '// MoveFirst -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    Close intFileID

    '//****
    If intFileID > 0 Then Open strFilePath & strFileName For Input As intFileID

    '//****
    lngCurrentRow = 0

    '//****

End Sub

Public Sub MoveLast()

  '//**************************************************************************
  '// MoveLast -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    MoveToRow lngRowCount

    '//****

End Sub

Private Sub SetRowCount()

  '//**************************************************************************
  '// SetRowCount -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

  Dim strTemp As String

    '//****
    Close intFileID

    '//****
    lngRowCount = 0

    '//****
    Open strFilePath & strFileName For Input As intFileID

    '//****
    Do While Not EOF(intFileID)

        '//****
        Line Input #1, strTemp

        '//****
        lngRowCount = lngRowCount + 1

        '//****
    Loop

    '//****
    MoveFirst

    '//****

End Sub

Public Function RowCount() As Long

  '//**************************************************************************
  '// RowCount -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    RowCount = lngRowCount

    '//****

End Function

Public Function CurrentRow() As Long

  '//**************************************************************************
  '// CurrentRow -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    CurrentRow = lngCurrentRow

End Function

Public Sub MoveToRow(lRowID As Long)

  '//**************************************************************************
  '// MoveToRow -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

  Dim lngTemp As Long
  Dim strTemp As String

    If (lRowID > lngRowCount) Or (intFileID = 0) Then Exit Sub

    lngTemp = 0

    MoveFirst
    Do While (Not EOF(intFileID)) And (lngTemp < lRowID)
        Line Input #intFileID, strTemp
        lngTemp = lngTemp + 1
    Loop

    lngCurrentRow = lngTemp

End Sub

Public Function GetDataFromRow(lRowID As Long) As String

  '//**************************************************************************
  '// GetDataFromRow -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

  Dim strTemp As String

    '//****
    If (lRowID > lngRowCount) Or (intFileID = 0) Then Exit Function

    '//****
    lngCurrentRow = lRowID + Abs((lngCurrentRow * (lRowID = 0)))

    '//****
    MoveToRow lngCurrentRow - 1

    '//****
    Line Input #intFileID, strTemp

    '//****
    MoveToRow lngCurrentRow + 1

    '//****
    GetDataFromRow = strTemp

    '//****

End Function

Public Sub DeleteRow(lRowID As Long)

  '//**************************************************************************
  '// DeleteRow -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    If lngRowCount = 0 Then Exit Sub

    '//****
    WriteData ReplaceData, "", lRowID

    '//****
    MoveToRow (Abs((lngCurrentRow - 1) = 0) + (lngCurrentRow - 1))

    '//****

End Sub

Public Function GetFilePath(sFileNamePath As String) As String

  '//**************************************************************************
  '// GetFilePath - funtion to get the file's path from the file as String
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sFileNamePath As String - Path and file
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************
  '//**** Variable Declarations

  Dim iPos As Integer
  Dim sChar As String

    '//**** Move back one character at a time by Step -1
    For iPos = Len(sFileNamePath) To 0 Step -1

        '//**** Extract character from iPos
        sChar = Mid$(sFileNamePath, iPos, 1)

        '//****
        If sChar = "\" Then

            '//****
            GetFilePath = Left$(sFileNamePath, iPos)

            '//****
            Exit Function

            '//****
        End If

        '//****
    Next iPos

    '//****

End Function

Public Function FileExists(sFileNamePath As String) As Boolean

  '//**************************************************************************
  '// FileExists - funtion to verify if the file sFileNamePath exists as Boolean
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sFileNamePath As String - Path and file
  '//
  '// Returns:
  '//     True if it exists, False if it doesn't
  '//
  '//**************************************************************************

  '//****

    sFileNamePath = Trim$(sFileNamePath)

    '//****
    If sFileNamePath = "" Then

        '//****
        FileExists = False

        '//****
        Exit Function

        '//****
    End If

    '//****
    FileExists = Len(Dir$(sFileNamePath)) <> 0

    '//****

End Function

Private Function FileDelete(sFileNamePath As String) As Boolean

  '//**************************************************************************
  '// FileDelete - funtion to delete a file as Boolean
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sFileNamePath As String - Path and file
  '//
  '// Returns:
  '//     True if the file was deleted, False if it was not
  '//
  '//**************************************************************************

  '//****

    DeleteFile (sFileNamePath)

    '//****
    FileDelete = True

    '//****

End Function

Private Function FileCopy(sFileNamePath As String, sNewFile As String, lOverwrite As Long)

  '//**************************************************************************
  '// FileCopy - function to copy a file
  '//            If lngOverwrite is set to a nonzero value (1), the function
  '//            will fail if the destination file already exists
  '//            If lngOverwrite is set to 0, the destination file will
  '//            be overwritten if it exists
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sFileNamePath As String - Path and file
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    FileCopy = CopyFile(sFileNamePath, sNewFile, lOverwrite)

    '//****
    If FileCopy = 0 Then

        '//**** Error

        '//****
        Exit Function

        '//****
    End If

    '//****

End Function

Public Function GetFileNameFromPath(strPath As String) As String

  '//**************************************************************************
  '// Class_Initialize -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

  Dim intX As Integer
  Dim intPlace As Integer
  Dim intLastPlace As Integer

    '//****
    intLastPlace = 0

    '//****
    For intX = 1 To Len(strPath)

        '//****
        intPlace = InStr(intLastPlace + 1, strPath, "\")

        '//****
        If intPlace = 0 Then

            '//****
            GetFileNameFromPath = Right$(strPath, Len(strPath) - intLastPlace)

            '//****
            Exit Function

            '//****
          Else

            '//****
            intLastPlace = intPlace

            '//****
        End If

        '//****
    Next intX

    '//****

End Function

Private Sub Class_Terminate()

  '//**************************************************************************
  '// Class_Terminate -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    Close

    '//****

End Sub

Private Sub Class_Initialize()

  '//**************************************************************************
  '// Class_Initialize -
  '// Created:  11/03/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     none
  '//
  '// Returns:
  '//     none
  '//
  '//**************************************************************************

  '//****

    intFileID = 0

    '//****

End Sub

Private Property Let WinPath(ByVal strWindowsPath As String)

  '//**************************************************************************
  '// Property Let WinPath - used when assigning Windows Path value property.
  '// Version 6.20.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Syntax:        X.WinPath = 5
  '//
  '//**************************************************************************

  '//****

    sWindowsPath = strWindowsPath

End Property

Public Property Get WinPath() As String

  '//**************************************************************************
  '// Property Get WinPath - used when retrieving Windows Path value property.
  '// Version 6.20.0
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Syntax:        Debug.Print X.WinPath
  '//
  '//**************************************************************************

  '//****

    WinPath = sWindowsPath

End Property

Public Sub GetWinPath()

  '//**** Variable Declarations

  Dim sWinPath As String

    '//**** Handle Errors
    On Error GoTo Error

    '//**** Set buffer
    sWindowsPath = Space$(255)

    '//**** Get length of windows directory
    sWinPath = GetWindowsDirectory(sWindowsPath, 255)

    '//****
    sWindowsPath = Left$(sWindowsPath, sWinPath)

    '//****
    WinPath = sWindowsPath

    '//****

Exit Sub

    '//****
Error:

    '//****
    Err.Raise 10001, "cCPA_Files.cls", "Unable To read the windows path."

End Sub


