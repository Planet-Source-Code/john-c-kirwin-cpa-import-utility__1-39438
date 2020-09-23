VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//****************************************************************************
'// cCPA_Strings - object to use for working with strings
'// Version 6.20.0
'// Created:  12/26/2001 John C. Kirwin (JCK)
'// Modified: 03/03/2002 JCK
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************

Dim sMessage As String
Dim CPATracker As cCPA_Tracker                                                        ' CPATracker. Object

Public Function FillAlign(strIn As String, bytLen As Byte, strFillChar As String, _
                          Optional blnAlignRight As Boolean) As String

  '//**************************************************************************
  '// FillAlign - Fill and align string character to length indicated
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     strIn As String - Inputstring
  '//     strFillChar As String - Character
  '//     bytLen As Byte - Length
  '//     Optional blnAlignRight As Boolean - Align the String right if true
  '//
  '// Returns:
  '//     FillAlign As String
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//**** Variable Declarations
  Dim iDiff As Integer

    '//**** Evaluate if strIn is shorter than bytLen
    If Len(strIn) < bytLen Then

        '//****
        For iDiff = 1 To (bytLen - Len(strIn))

            '//****
            If blnAlignRight Then

                '//****
                strIn = strFillChar & strIn

                '//****
              Else

                '//****
                strIn = strIn & strFillChar

                '//****
            End If

            '//****
        Next iDiff

        '//****
    End If

    '//****
    FillAlign = Left$(strIn, bytLen)

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the PendingInvoices Sub"
    '//**** Error Handling
    Select Case Err.Number
      Case 6
        '//**** Handle VB error 6 Overflow
        '//**** Resume Next
      Case 13
        '//**** Type mismatch occurred
        '//**** Resume Next
      Case 62
        '//**** Input past end of file occurred
        '//**** Resume Next
      Case 91
        '//**** Handle VB error 91 Object variable or With block variable not set
        '//**** Resume Next
      Case 440
        '//**** Handle VB error 440 Automation
        '//**** Resume Next
      Case Else
        '//**** Unhandled errors
        '//**** Resume Next
    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

End Function

Public Function sCrop(sBuf As String, iLength As Integer) As String

  '//**************************************************************************
  '// sCrop - Truncates a string by a specified number of characters.
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sBuf As String - The string to truncate
  '//     iLength As Integer - The number of characters to truncate
  '//
  '// Returns:
  '//     sCrop as string truncated as specified
  '//
  '//**************************************************************************

  '//****

    If Len(sBuf) <= iLength Then

        '//****
        sCrop = ""

        '//****
      Else

        '//****
        sCrop = Left$(sBuf, Len(sBuf) - iLength)

        '//****
    End If

    '//****

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
  'Set CPATracker = New cCPA_Tracker

End Sub

