VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCPA_INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//****************************************************************************
'// cCPA_INI - Class Interface to Layouts.ini file read & write to ini
'// files using the Windows API
'//
'// Version:  6.20.0
'// Created:  12/26/2001 John C. Kirwin (JCK)
'// Modified: 03/03/2002 JCK - Format & Comment
'//
'// Dependencies:  None
'// Implements:    None
'//
'//****************************************************************************
'//**** GetPrivateProfileString
Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Private Declare Function GetPrivateProfileInt& Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String)
Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String)
Private Declare Function WritePrivateProfileSection& Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String)

'//**** Local variables to hold property values.
Private mstrINIPath As String
Private mstrFileName As String

'Private CPAFiles As cCPA_Files                                                ' CPAFiles. Object
Public Function GetFromINI(sSection As String, sKey As String, sPath As String) As String

  '//**************************************************************************
  '// GetFromINI - function returns the string value of a key in the
  '//     Layouts.ini file or an empty string if the key is not found
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     Directory As String - the path and name of the Layouts.ini file
  '//     sSection As String  - the section of the Layouts.ini file the key is located under
  '//     Key As String -  the key of the value to be returned
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Variable Declarations

  Dim strBuffer As String

    '//**** Set buffer string
    strBuffer = String$(750, Chr$(0))

    '//**** GetPrivateProfileString API call to get KeyValue string from INI file
    GetFromINI = Left$(strBuffer, GetPrivateProfileString(sSection, _
                 ByVal sKey, "", strBuffer, Len(strBuffer), sPath))

    '//**** Exit Sub/Function before error handler

End Function

Public Function DeleteKey(sSection As String, sKey As String, sPath As String)

  '//**************************************************************************
  '// DeleteKey - method that deletes a key in the specified section
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sSection As String - the name of the section to delete key from
  '//     sKey As String - the key to delete
  '//     sPath As String - the path and name of the Layouts.ini file
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** API Call WritePrivateProfileString

    DeleteKey = WritePrivateProfileString(sSection, sKey, vbNullString, sPath)

    '//**** Exit Sub/Function before error handler

End Function

Public Function DeleteSection(sSection As String, sPath As String)

  '//**************************************************************************
  '// DeleteSection - function that deletes a section and all its keys from ini file
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sSection As String - the name of the section to delete
  '//     sPath As String - the path and name of the Layouts.ini file
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** API Call WritePrivateProfileString

    DeleteSection = WritePrivateProfileString(sSection, vbNullString, vbNullString, sPath)

    '//**** Exit Sub/Function before error handler

End Function

Public Function WriteToINI(sSection As String, sKey As String, _
                           sKeyValue As String, sPath As String)

  '//**************************************************************************
  '// WriteToINI
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     sSection As String  - the section of the Layouts.ini file the key is located under
  '//     Key As String -  the key of the value to be returned
  '//     sKeyValue As String - the new value to be written to the Layouts.ini file
  '//     Path As String - the path and name of the Layouts.ini file
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** API Call WritePrivateProfileString

    WriteToINI = WritePrivateProfileString(sSection, sKey, vbNullString, sPath)
    WriteToINI = WritePrivateProfileString(sSection, sKey, sKeyValue, sPath)

    '//**** Exit Sub/Function before error handler

End Function

Private Sub Class_Initialize()

  '//**************************************************************************
  '// Class_Initialize
  '// Created:  12/26/2001 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Format & Comment
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

End Sub


