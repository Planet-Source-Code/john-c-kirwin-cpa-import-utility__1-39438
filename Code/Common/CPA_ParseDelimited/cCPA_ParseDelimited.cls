VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_ParseDelimited"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'//****************************************************************************
'// cCPA_ParseDelimited - Class Interface to delimited string parcing utilities
'//
'// Version:       6.20.0
'// Created:       07/28/2002 John C. Kirwin (JCK)
'// Modified:      07/28/2002 JCK - Format & Comment
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************

'//**** Local variables to hold property values.
Private mstrINIPath As String
Private mstrFileName As String
Private sMessage As String

Private CPAFiles As cCPA_Files                                                            ' CPAFiles Object
Private CPATracker As cCPA_Tracker                                                        ' CPATracker Object

Public Function ParseString(ByVal sString As String _
                            , ByVal iNumber As Integer _
                            , sDelimiter As String) As String

  '//**********************************************************************
  '// ParseString - Parses out a field as a string from a delimited string
  '// sString ByVal so changes are not propagated back to the caller
  '//
  '// Version:  6.20.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '// sString As String - delimited string
  '// iNumber As Integer - number of the field to be parsed out of delimited string
  '// sDelimiter As String - Delimiter
  '//
  '// Returns:
  '//     ParseString - string parsed out of a delimited string
  '//
  '//**********************************************************************

    On Error GoTo EH

    '//**** Variable Declarations
  Dim iFieldPos As Integer
  Dim iFieldPosold As Integer
  Dim iCurrentPosition As Integer
  Dim sText As String

    '//**** Wrap sString with delimeters
    sString = sDelimiter & sString & sDelimiter

    '//**** Verify there is a string to parce sString and that it conatains the
    '//     delimeter as string sDelimeter and the section to be parsed is
    '//     indicated by an integer greater than 0
    If Len(sString) > 0 And InStr(sString, sDelimiter) > 0 And iNumber > 0 Then

        '//****
        iFieldPos = 1

        '//****
        iFieldPosold = 1

        '//**** Validate the field to parse out of the string indicated
        '//     by the iNumber is not greater than the total number of
        '//     the fields separated by the sDelimiter
        Do While InStr(iFieldPosold + 1, sString, sDelimiter) > 0

            '//**** Track the old field position
            iFieldPosold = iFieldPos

            '//**** Return the position of the first occurrence of the string
            iFieldPos = InStr(iFieldPos + 1, sString, sDelimiter)

            '//**** Increment Current Position by 1
            iCurrentPosition = iCurrentPosition + 1

            '//**** Continue Do While...Loop through fields of sString
            '//     separated by the sDelimiter
        Loop

        '//**** Compare the the field to parse out is greater than
        '//     the total parsible fields
        If iNumber > iCurrentPosition Then

            '//**** Exit function if it is greater than the total
            '//     parsible fields
            Exit Function

        End If

        '//**** Reset field position integer iFieldPos to 1
        iFieldPos = 1

        '//**** Reset current position integer iCurrentPosition to 0
        iCurrentPosition = 0

        '//**** Begin the procedure to parse out the iNumber field
        '//     from sString
        Do

            '//**** Track field position
            iFieldPosold = iFieldPos

            '//****
            iFieldPos = InStr(iFieldPos + 1, sString, sDelimiter)

            '//****
            If iFieldPos > 0 Then

                '//****
                sText = Mid$(sString, iFieldPosold + 1, (iFieldPos - 1) - iFieldPosold)

              Else
                '//****
                sText = Mid$(sString, iFieldPosold + 1)

            End If

            '//**** Increment Current Position by 1
            iCurrentPosition = iCurrentPosition + 1

            '//**** Compare the current position integer iCurrentPosition
            '//     to the field integer indicated as the field to parse
            '//     out of the string sString
            If iCurrentPosition = iNumber Then

                '//**** Set ParseString string function
                ParseString = sText

                Exit Do

            End If

            '//**** Continue Do...Loop
        Loop

    End If

    '//****
    ParseString = sText

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ParseString function"
    '//**** Error Handling
    Select Case Err.Number

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next
    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function/procedure
    Resume Next

    '//****

End Function

Public Function lStringToArray(vData() As Variant _
                               , ByVal sSource As String _
                               , ByVal lDelimiter As Long) As Long

  '//**********************************************************************
  '// lStringToArray - Return the total number of members of the array as Long
  '//
  '// Version:  6.20.0
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '// vData() As Variant
  '// ByVal sSource As String
  '// ByVal lDelimiter As Long
  '//
  '// Returns:
  '//     None
  '//
  '//**********************************************************************

    On Error GoTo EH

    '//**** Declare variables
  Dim lngCurPos As Long
  Dim lngNextPos As Long
  Dim lDelimiterLength As Long
  Dim lArrayMemberCount As Long
  Dim varValue As Variant

    '//**** Initialze the array
    ReDim vData(1) As Variant

    '//**** Add delimiters to start and end of sSource string to make loop simpler
    sSource = Chr$(lDelimiter) & sSource & Chr$(lDelimiter)

    '//**** Calculate the delimiter length only once
    lDelimiterLength = Len(Chr$(lDelimiter))

    '//**** Initialize the count and position
    lArrayMemberCount = 0
    lngCurPos = lDelimiterLength
    lngNextPos = InStr(lngCurPos, sSource, Chr$(lDelimiter))

    '//**** Loop searching for delimiters
    Do Until lngNextPos = 0

        '//**** Extract a new member to store in the vData array
        varValue = Mid$(sSource, lngCurPos, lngNextPos - lngCurPos)

        '//**** Add room for the new member in the array:
10                      ReDim Preserve vData(lArrayMemberCount) As Variant

        '//**** Put the member data in the array
        vData(lArrayMemberCount) = varValue

        Debug.Print "vData(" & lArrayMemberCount & ") : " & vData(CInt(lArrayMemberCount)) & " varValue: " & varValue & " - " & lngCurPos

        '//**** Increment the data counter
        lArrayMemberCount = lArrayMemberCount + 1

        If lngNextPos = 0 Then
            Exit Do
          Else

            '//**** Position to the last found delimiter
            lngCurPos = lngNextPos + lDelimiterLength

            '//**** Find the next delimiter
            lngNextPos = InStr(lngCurPos, sSource, Chr$(lDelimiter))

            If lngCurPos >= Len(sSource) Then
                varValue = ""
                lngNextPos = 0
                GoTo 10
            End If
        End If

    Loop
    '//**** Return the total number of members of the array
    lStringToArray = (lArrayMemberCount - 2)

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the lStringToArray function"
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

Public Function sParse(sBuf As String, sDelim As String) As String

  '//**************************************************************************
  '// sParse - Returns a section of a string from the beginning to the
  '//          first occurance of a delimitter.
  '// Version:  6.20.0
  '// Created:  03/03/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sBuf As String - (in/out) - The string to parse
  '//     sDelim As String - The delimitter used for the parse
  '//
  '// Returns:
  '//     sParse as string parsed from the beginning a sBuf string
  '//     to the first occurance of the sDelim delimitter indicated
  '//
  '//**************************************************************************

  '//****

  Dim iPosition As Integer

    '//****
    iPosition = InStr(sBuf, sDelim)

    '//**** If delimitter does not exist in string, return whole string
    If iPosition = 0 Then

        '//****
        sParse = Trim$(sBuf)

        '//****
        sBuf = ""

        '//****
      Else

        '//****
        sParse = Trim$(Left$(sBuf, iPosition - 1))

        '//****
        sBuf = Mid$(sBuf, iPosition + Len(sDelim))

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


