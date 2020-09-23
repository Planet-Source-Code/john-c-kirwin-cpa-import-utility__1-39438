VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_Layouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'//****************************************************************************
'// cCPA_Layouts - Class Interface to Load Import definition into
'// multi-dimensional array allocating elements based on
'// Import Layout definition from Layouts.ini file
'//
'// Version 6.20.0
'// Created:  07/28/2002 John C. Kirwin (JCK)
'// Modified: 07/28/2002 JCK - Format & Comment
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************

Private sMessage As String

'//**** Object Declarations
Private CPAFiles As cCPA_Files                                                ' CPAFiles Object
Private CPAINI As cCPA_INI                                                    ' CPAINI Object
Private CPAStrings As cCPA_Strings                                            ' CPAStrings Object
Private CPATracker As cCPA_Tracker                                            ' CPATracker Object
Private CPAErrorHandler As cCPA_ErrorHandler                                  ' CPAErrorHandler Object
Private CPAParseDelimited As cCPA_ParseDelimited                              ' CPAParseDelimited Object
Private CPALayouts As cCPA_Layouts                                            ' CPALayouts Object

Public Function GLAccounts(aGLARL() As String, iRows As Integer, iColumns As Integer) As Boolean

  '//**************************************************************************
  '// GLAccounts - Loads Import definition into multi-dimensional array
  '//              allocating elements based on Import Layout definition
  '//              from Layouts.ini file
  '//
  '// Created:  07/26/2002 John C. Kirwin (JCK)
  '// Modified: 07/26/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     aGLARL() As String
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     GLAccounts as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** GL Account Record Layout array (8 Fields)

    '//**** Define the dynamic array size variables
    iRows = 8
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLARL(iRows, iColumns)

    '//**** GLARL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLARL(iPos, 1) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLARL(iPos, 2) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Description"), sLayoutINI)
        aGLARL(iPos, 3) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GLAccounts Sub"
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

Public Function GLSegments(aGLSRL() As String, iRows As Integer, iColumns As Integer) As Boolean

  '//**************************************************************************
  '// GLSegments - Loads Import definition into multi-dimensional array
  '//              allocating elements based on Import Layout definition
  '//              from Layouts.ini file.
  '//
  '// Created:  07/26/2002 John C. Kirwin (JCK)
  '// Modified: 07/26/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     PendingInvoices as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** GL Account Record Layout array (8 Fields)

    '//**** Define the dynamic array size variables
    iRows = 7
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLSRL(iRows, iColumns)

    '//**** GLSRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLSRL(iPos, 1) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLSRL(iPos, 2) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aGLSRL(iPos, 3) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GLSegments Sub"
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

Public Function GLNaturalAccts(aGLNRL() As String, iRows As Integer, _
                               iColumns As Integer) As Boolean

  '//**************************************************************************
  '// GLNaturalAccts - Loads Import definition into multi-dimensional array
  '//                  allocating elements based on Import Layout definition
  '//                  from Layouts.ini file.
  '//
  '// Created:  07/26/2002 John C. Kirwin (JCK)
  '// Modified: 07/26/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     GLNaturalAccts as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** GL Account Record Layout array (8 Fields)

    '//**** Define the dynamic array size variables
    iRows = 10
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLNRL(iRows, iColumns)

    '//**** GLNRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLNRL(iPos, 1) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLNRL(iPos, 2) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aGLNRL(iPos, 3) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GLNaturalAccts Sub"
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

Public Function GLTransactions(aGLBRL() As String, aGLTRL() As String, _
                               aGLFRL() As String, iRows As Integer, _
                               iColumns As Integer) As Boolean

  '//**************************************************************************
  '// GLNaturalAccts - Loads Import definition into multi-dimensional array
  '//                  allocating elements based on Import Layout definition
  '//                  from Layouts.ini file.
  '//
  '// Created:  07/26/2002 John C. Kirwin (JCK)
  '// Modified: 07/26/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     GLTransactions as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** Define the dynamic array size variables
    iRows = 4
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLBRL(iRows, iColumns)

    '//**** GLBRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLBRL(iPos, 1) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLBRL(iPos, 2) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aGLBRL(iPos, 3) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Define the dynamic array size variables
    iRows = 19
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLTRL(iRows, iColumns)

    '//**** GLTRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLTRL(iPos, 1) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLTRL(iPos, 2) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aGLTRL(iPos, 3) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Define the dynamic array size variables
    iRows = 2
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aGLFRL(iRows, iColumns)

    '//**** GLFRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aGLFRL(iPos, 1) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aGLFRL(iPos, 2) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aGLFRL(iPos, 3) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GLTransactions Sub"
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

Public Function PostedInvoices(aRIBRL() As String, aRIRRL() As String, aRIDRL() As String _
                               , aRIARL() As String, aRIERL() As String, aRITRL() As String _
                               , iRows As Integer, iColumns As Integer) As Boolean

  '//**************************************************************************
  '// PostedInvoices - Loads Import definition into multi-dimensional array
  '//                  allocating elements based on Import Layout definition
  '//                  from Layouts.ini file.
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     PostedInvoices as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"         ' Layouts.ini

    '//**** Posted Invoice Batch Record Layout array (17 Fields)

    '//**** Define the dynamic array size variables
    iRows = 9
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRIBRL(iRows, iColumns)

    '//**** RIBRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRIBRL(iPos, 1) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRIBRL(iPos, 2) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRIBRL(iPos, 3) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Invoice Header Record Layout array (84 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 40
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRIRRL(iRows, iColumns)

    '//**** RIRRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRIRRL(iPos, 1) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRIRRL(iPos, 2) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRIRRL(iPos, 3) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)

    Loop

    '//**** Posted Invoice Detail Record Layout array (44 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 28
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRIDRL(iRows, iColumns)

    '//**** RIDRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRIDRL(iPos, 1) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRIDRL(iPos, 2) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRIDRL(iPos, 3) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Invoice Application Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 12
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRIARL(iRows, iColumns)

    '//**** RIARL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRIARL(iPos, 1) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRIARL(iPos, 2) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRIARL(iPos, 3) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Invoice Tax Header Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 10
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRIERL(iRows, iColumns)

    '//**** RIERL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRIERL(iPos, 1) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRIERL(iPos, 2) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRIERL(iPos, 3) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Invoice Tax Detail Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 12
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aRITRL(iRows, iColumns)

    '//**** RITRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aRITRL(iPos, 1) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aRITRL(iPos, 2) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aRITRL(iPos, 3) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the PostedInvoices Sub"
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

Public Function PendingInvoices(aPIBRL() As String, aPIHRL() As String, _
                                aPIDRL() As String, iRows As Integer, _
                                iColumns As Integer) As Boolean

  '//**************************************************************************
  '// PendingInvoices - Loads Import definition into multi-dimensional array
  '//                   allocating elements based on Import Layout definition
  '//                   from Layouts.ini file.
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     PendingInvoices as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** Pending Invoice Batch Record Layout array (17 Fields)

    '//**** Define the dynamic array size variables
    iRows = 17
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPIBRL(iRows, iColumns)

    '//**** PIBRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPIBRL(iPos, 1) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPIBRL(iPos, 2) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPIBRL(iPos, 3) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Pending Invoice Header Record Layout array (84 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 84
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPIHRL(iRows, iColumns)

    '//**** PIHRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPIHRL(iPos, 1) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPIHRL(iPos, 2) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPIHRL(iPos, 3) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)

    Loop

    '//**** Pending Invoice Detail Record Layout array (44 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 44
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPIDRL(iRows, iColumns)

    '//**** PIDRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPIDRL(iPos, 1) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPIDRL(iPos, 2) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPIDRL(iPos, 3) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

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
Public Function PostedVouchers(aVPBRL() As String, aVPHRL() As String, aVPDRL() As String _
                               , aVPARL() As String, aVPXRL() As String, aVPTRL() As String _
                               , iRows As Integer, iColumns As Integer) As Boolean

  '//**************************************************************************
  '// PostedVouchers - Loads Import definition into multi-dimensional array
  '//                  allocating elements based on Import Layout definition
  '//                  from Layouts.ini file.
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     PostedVouchers as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** Posted Voucher Batch Record Layout array

    '//**** Define the dynamic array size variables
    iRows = 8
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPBRL(iRows, iColumns)

    '//**** VPBRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPBRL(iPos, 1) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPBRL(iPos, 2) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPBRL(iPos, 3) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Voucher Header Record Layout array (84 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 39
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPHRL(iRows, iColumns)

    '//**** VPHRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPHRL(iPos, 1) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPHRL(iPos, 2) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPHRL(iPos, 3) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)

    Loop

    '//**** Posted Voucher Detail Record Layout array (44 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 25
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPDRL(iRows, iColumns)

    '//**** VPDRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPDRL(iPos, 1) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPDRL(iPos, 2) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPDRL(iPos, 3) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Voucher Application Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 10
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPARL(iRows, iColumns)

    '//**** VPARL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPARL(iPos, 1) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPARL(iPos, 2) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPARL(iPos, 3) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Voucher Tax Header Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 15
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPXRL(iRows, iColumns)

    '//**** VPXRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPXRL(iPos, 1) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPXRL(iPos, 2) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPXRL(iPos, 3) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Posted Voucher Tax Detail Record Layout array

    '//**** Define the size of the dynamic array MyArray()
    iRows = 16
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aVPTRL(iRows, iColumns)

    '//**** VPTRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aVPTRL(iPos, 1) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aVPTRL(iPos, 2) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aVPTRL(iPos, 3) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the PostedVouchers Sub"
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
Public Function PendingVouchers(aPVBRL() As String, aPVHRL() As String, _
                                aPVDRL() As String, iRows As Integer, _
                                iColumns As Integer) As Boolean

  '//**************************************************************************
  '// PendingVouchers - Loads Import definition into multi-dimensional array
  '//                   allocating elements based on Import Layout definition
  '//                   from Layouts.ini file.
  '//
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     PendingInvoices as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** Pending Voucher Batch Record Layout array (17 Fields)

    '//**** Define the dynamic array size variables
    iRows = 17
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPVBRL(iRows, iColumns)

    '//**** PVBRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPVBRL(iPos, 1) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPVBRL(iPos, 2) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPVBRL(iPos, 3) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** Pending Voucher Header Record Layout array (84 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 75
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPVHRL(iRows, iColumns)

    '//**** PIHRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPVHRL(iPos, 1) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPVHRL(iPos, 2) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPVHRL(iPos, 3) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)

    Loop

    '//**** Pending Voucher Detail Record Layout array (44 Fields)

    '//**** Define the size of the dynamic array MyArray()
    iRows = 38
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aPVDRL(iRows, iColumns)

    '//**** PIDRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aPVDRL(iPos, 1) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aPVDRL(iPos, 2) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aPVDRL(iPos, 3) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)

    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the PendingVouchers Sub"
    '//**** Error Handling
    Select Case Err.Number
      Case 6
        '//**** Handle VB error 6 Overflow
        '//**** Resume Next
      Case 13
        '//**** Type mismatch occurred
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

Public Function IMItems(aIMIRL() As String, aIMURL() As String, _
                        aIMLRL() As String, iRows As Integer, _
                        iColumns As Integer) As Boolean

  '//**************************************************************************
  '// IMItems - Loads Import definition into multi-dimensional array
  '//           allocating elements based on Import Layout definition
  '//           from Layouts.ini file.
  '//
  '// Created:  07/26/2002 John C. Kirwin (JCK)
  '// Modified: 07/26/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     iRows As Integer - Total Rows in the Array
  '//     iColumns As Integer - Total Columns in the Array
  '//
  '// Returns:
  '//     IMItems as Boolean to indicate success for failure
  '//
  '//**************************************************************************

    On Error GoTo EH

  Dim iPos As Integer
  Dim sLayoutINI As String

    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"

    '//**** IM Items Record Layout array (61 Fields)

    '//**** Define the dynamic array size variables
    iRows = 61
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aIMIRL(iRows, iColumns)

    '//**** IMIRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aIMIRL(iPos, 1) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aIMIRL(iPos, 2) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aIMIRL(iPos, 3) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** IM Items Unit of Measure Record Layout array (61 Fields)

    '//**** Define the dynamic array size variables
    iRows = 8
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aIMURL(iRows, iColumns)

    '//**** IMURL
    Do Until iPos = iRows
        iPos = iPos + 1
        aIMURL(iPos, 1) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aIMURL(iPos, 2) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aIMURL(iPos, 3) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    '//**** IM Landed Cost Factor Record Layout array (3 Fields)

    '//**** Define the dynamic array size variables
    iRows = 3
    iColumns = 3
    iPos = 0

    '//**** Re Dimension the multi-dimensional array
    '//     allocating elements based on iRows
    '//     and iColumns
    ReDim aIMLRL(iRows, iColumns)

    '//**** IMLRL
    Do Until iPos = iRows
        iPos = iPos + 1
        aIMLRL(iPos, 1) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_FieldName"), sLayoutINI)
        aIMLRL(iPos, 2) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                                             CPAStrings.FillAlign(CStr(iPos), _
                                             3, "0", True) & _
                                            "_Description"), sLayoutINI)
        aIMLRL(iPos, 3) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                                            CPAStrings.FillAlign(CStr(iPos), _
                                            3, "0", True) & _
                                             "_Required"), sLayoutINI)
    Loop

    Set CPAFiles = Nothing
    Set CPAINI = Nothing
    Set CPAStrings = Nothing
    Set CPATracker = Nothing
    Set CPAErrorHandler = Nothing
    Set CPAParseDelimited = Nothing
    Set CPALayouts = Nothing

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the IMItems Sub"
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


