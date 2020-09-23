VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Tracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//****************************************************************************
'// cCPA_Tracker - object to log/track messages to file, screen, & debug.print
'//
'// Version 6.20.0
'// Created:  12/26/2001 John C. Kirwin (JCK)
'// Modified: 03/03/2002 JCK
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************

Public Function Tracker(sMessage As String, sFile As String, _
                        bMsg As Boolean, bLog As Boolean)

  '//**************************************************************************
  '// Tracker - function to Log/Track user experience information or errors by
  '// indicating whether to pop infomative messages and/or log messages for
  '// procedure flow. Helps development with debug.print of sMessage to immediate window.
  '//
  '// Version 6.20.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sMessage as String - The information to be tracked/logged
  '//     sFile As String    - Indicate the file
  '//     bMsg as Boolean    - Pop Msg to UI?
  '//     bLog as Log        - Log entry to file?
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Variable Declarations

  Dim LogFile As Long

    '//**** Evaluate bLog Boolean to determine whether to log sMessage
    If bLog Then

        '//**** Assign the variable nFile to a Free File Num.
        LogFile = FreeFile

        '//**** Open the file for editing
        Open sFile For Append Shared As #LogFile

        '//**** Write to file and timestamped with the date and time formated
        '//     ex. 12/26/01 23:59:59 - text of sMessage
        Print #LogFile, Format$(Now, "mm/dd/yy") & " " & _
              Format$(Now, "hh:mm:ss") & " - " & sMessage

        '//**** Close the file
        Close #LogFile

        '//**** End if bLog to determine whether to log sMessage
    End If

    '//**** Pop MsgBox sMessage
    If bMsg Then
        MsgBox sMessage, vbInformation, "Warning"
    End If

    '//**** Debug.Print sMessage
    Debug.Print sMessage

End Function

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


End Sub




