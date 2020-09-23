VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//****************************************************************************
'// cCPAError Handler - object to properly handle distributed application
'// errors.  Specifically errors that occur in a class module by raising
'// custom error codes to be handled by the application that using
'// using the class that raised the error.
'//
'// Version 6.20.0
'// Created:  12/26/2001 John C. Kirwin (JCK)
'// Modified: 03/03/2002 JCK - Format & Comment
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************
Dim CPATracker As cCPA_Tracker                                                ' CPATracker. Object

Public Function ErrorHandler(sObject As String, sSource As String, _
                             ErrNumber As Integer, ErrDesc As String, _
                             sFile As String, bMsg As Boolean, bLog As Boolean)

  '//**************************************************************************
  '// ErrorHandler -
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

  Dim sMessage As String

    '//**** Error Handling
    Select Case Err.Number

        '//**** Handle VB error 6 Overflow
      Case 6

        '//****

        '//**** Type mismatch occurred
      Case 13

        '//****

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//****

        '//**** Handle File error 52 Bad file name or number
      Case 52

        '//****

        '//**** Handle File error 52 Bad file name or number
      Case 53

        '//****

        '//**** Handle File error 75 Path/File access error
      Case 75

        '//****

        '//**** Handle VB error 440 Automation
      Case 440

        '//****

        '//**** Unhandled errors
      Case Else

        '//****
        sMessage = ErrNumber & ": " & ErrDesc & _
                   " occurred during " & sSource & _
                   " in " & sObject

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, sFile, bMsg, bLog

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

    Set CPATracker = New cCPA_Tracker

    '//****

End Sub


