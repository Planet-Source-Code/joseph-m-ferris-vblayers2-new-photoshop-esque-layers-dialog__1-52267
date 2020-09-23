Attribute VB_Name = "modGUID"
'**********************************************************************************************************************'
'*'
'*' Module    : modGUID
'*'
'*'
'*' Author    : Joseph M. Ferris <jferris@desertdocs.com>
'*'
'*' Date      : 02.26.2004
'*'
'*' Depends   : None
'*'
'*' Purpose   : Provides a wrapped function to assist in the creation of GUIDs (Globally Unique Identifers)
'*'
'*' Notes     :
'*'
'**********************************************************************************************************************'
Option Explicit

'**********************************************************************************************************************'
'*'
'*' Private API Declarations - ole32.dll
'*'
'**********************************************************************************************************************'
Private Declare Function CoCreateGuid Lib "ole32" ( _
        id As Any) As Long

'**********************************************************************************************************************'
'*'
'*' Procedure : CreateGUID
'*'
'*'
'*' Date      : 02.26.2004
'*'
'*' Purpose   : Generate a GUID
'*'
'*' Input     : None.
'*'
'*' Output    : CreateGUID (String)
'*'
'**********************************************************************************************************************'
Public Function CreateGUID() As String

'*' Fail through on local errors.
'*'
On Error Resume Next

Dim bytID(0 To 15)                      As Byte
Dim lngCounter                          As Long
    
    '*' Make sure that CoCreateGuid can properly seed the GUID.
    '*'
    If CoCreateGuid(bytID(0)) = 0 Then
    
        '*' Iterate through each byte of the GUID.
        '*'
        For lngCounter = 0 To 15
        
            '*' Append it with the byte value from the function call.
            '*'
            CreateGUID = CreateGUID + IIf(bytID(lngCounter) < 16, "0", vbNullString) + Hex$(bytID(lngCounter))
            
        Next lngCounter
        
        '*' Format it to match the dashed presentation format.
        '*'
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + _
                     Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
                     
    End If
    
End Function
