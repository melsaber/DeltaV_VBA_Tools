Attribute VB_Name = "Module1"
'/*
' * Copyright (c) 2019-2020, Mohamed Elsaber
' * 
' * use my code as per the GNU GPL V3.0
' */
Option Explicit
Public Type DV_Value
    strResult As String
    fltResult As Double
    'iValueType As Integer '0: not readable, 1: string, 2 float, 3: float and string
    sReadErrorMsg As String
End Type
Public Type textMatch
    matchtext As String
    matchStart As Integer
    matchLength As Integer
    matchType As String
End Type

Public matches() As textMatch
Public Const sTitle = "Code Helper"
Private Function KeyWords() As String
    KeyWords = "WHILE"
    KeyWords = KeyWords & "|END_WHILE"
    KeyWords = KeyWords & "|DO"
    KeyWords = KeyWords & "|IF"
    KeyWords = KeyWords & "|ELSE"
    KeyWords = KeyWords & "|THEN"
    KeyWords = KeyWords & "|ENDIF"
    KeyWords = KeyWords & "|END_IF"
    KeyWords = KeyWords & "|AND"
    KeyWords = KeyWords & "|OR"
    KeyWords = KeyWords & "|NOT"
    KeyWords = KeyWords & "|TRUE"
    KeyWords = KeyWords & "|FALSE"
    KeyWords = KeyWords & "|LO"
    KeyWords = KeyWords & "|MAN"
    KeyWords = KeyWords & "|IMAN"
    KeyWords = KeyWords & "|AUTO"
    
    
End Function

Private Sub FindMatches(ByRef txt As String, sPattern As String, matchType As String)
    Dim rgxMatches As Object, Item As Object, regEx As Object
    Dim i As Integer
    Set regEx = CreateObject("vbscript.regexp")
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = sPattern '
    End With
    Set rgxMatches = regEx.Execute(txt)
    i = UBound(matches) + 1
    If rgxMatches.Count > 0 Then
        ReDim Preserve matches(i + rgxMatches.Count - 1)
        For Each Item In rgxMatches
            matches(i).matchtext = Item.Value
            matches(i).matchStart = Item.FirstIndex
            matches(i).matchLength = Item.Length
            matches(i).matchType = matchType
            txt = Left(txt, Item.FirstIndex) & Space(Item.Length) & Right(txt, Len(txt) - Item.FirstIndex - Item.Length)
            i = i + 1
        Next Item
    End If
    
    Set regEx = Nothing
End Sub

Public Sub GetMatches(codeText As String)
    
    Dim tmpmatch As textMatch
    Dim i As Integer, j As Integer
    ReDim matches(0)


    FindMatches codeText, "\(\*.+?\*\)", "Comment"
    FindMatches codeText, """.+?""", "String"
    FindMatches codeText, "'((?![:\s]).)+?'", "DV Reference"
    FindMatches codeText, "'\S+:.+?'", "Named Set"
    FindMatches codeText, KeyWords, "Keyword"

    'sorting results
    For i = 1 To UBound(matches)
        For j = i To UBound(matches)
            If matches(i).matchStart > matches(j).matchStart Then
                tmpmatch = matches(i)
                matches(i) = matches(j)
                matches(j) = tmpmatch
            End If
        Next j
    Next i
    
End Sub

Public Function getRTF(txtCode As String) As String
    Dim txt As String
    Dim CodeRTF As String
    Dim xxx As String
    Dim lastPos As Integer
    Dim i As Integer
    txt = Replace(txtCode, vbCrLf, vbCr)
    
    Call GetMatches(txt)
    
    CodeRTF = ""
    lastPos = 1
    For i = 1 To UBound(matches)
        If (matches(i).matchStart > lastPos) Then
            xxx = "\cf0" & Mid$(txt, lastPos, matches(i).matchStart - lastPos + 1)
            CodeRTF = CodeRTF & vbLf & xxx
            
        End If
        
        lastPos = matches(i).matchStart + matches(i).matchLength
        
        Select Case matches(i).matchType
    
        Case "Keyword"
            CodeRTF = CodeRTF & vbLf & "\cf2" & matches(i).matchtext
        Case "DV Reference"
            CodeRTF = CodeRTF & vbLf & "\cf3" & matches(i).matchtext
        Case "Comment"
            CodeRTF = CodeRTF & vbLf & "\cf1" & matches(i).matchtext
        Case "String"
            CodeRTF = CodeRTF & vbLf & "\cf5" & matches(i).matchtext
        Case "Named Set"
            CodeRTF = CodeRTF & vbLf & "\cf4" & matches(i).matchtext
        End Select
        
    Next i
    
    If Len(txt) >= lastPos Then
        xxx = "\cf0" & Mid$(txt, lastPos, Len(txt) - lastPos + 1)
        CodeRTF = CodeRTF & vbLf & "\cf0" & xxx
    End If
    
    
    CodeRTF = Replace(CodeRTF, vbCr, "\par ")
    
    getRTF = "{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1033{\fonttbl{\f0\fnil\fcharset0 Courier New;}}"
    getRTF = getRTF & vbCrLf & "{\colortbl ;"          ' cf0 black
    getRTF = getRTF & vbCrLf & "\red0\green128\blue0;" 'cf1 comments
    getRTF = getRTF & vbCrLf & "\red0\green0\blue255;" 'cf2 keywords
    getRTF = getRTF & vbCrLf & "\red150\green0\blue0;" 'cf3ref
    getRTF = getRTF & vbCrLf & "\red0\green0\blue150;" 'cf4 namedset
    getRTF = getRTF & vbCrLf & "\red255\green0\blue0;" 'cf5 string
    getRTF = getRTF & vbCrLf & "}"
    getRTF = getRTF & vbCrLf & "\pard\f0\fs20"
    
    getRTF = getRTF & vbCrLf & CodeRTF
    
    getRTF = getRTF & vbCrLf & "\par"
    getRTF = getRTF & vbCrLf & "}"

End Function

Public Function ClearMatches()
    ReDim matches(0)
End Function

Public Function GetCursorMatch(cursorPos As Integer) As textMatch
    Dim i As Integer
    
    If UBound(matches) = 0 Then
        Exit Function
    End If
    
    For i = 1 To UBound(matches)
        If cursorPos >= matches(i).matchStart And cursorPos <= matches(i).matchStart + matches(i).matchLength Then
            GetCursorMatch = matches(i)
            Exit Function
        ElseIf (matches(i).matchStart + matches(i).matchLength) > cursorPos Then
            Exit Function
        End If
    Next i
    
End Function

Public Function ClipPaste() As String

    Dim DataObj As New MSForms.DataObject
    On Error GoTo Err_handler
    DataObj.GetFromClipboard

    ClipPaste = Replace(DataObj.GetText(1), vbCr & vbCr, vbCr)
    DataObj.SetText ""
    DataObj.PutInClipboard
Err_handler:
    Set DataObj = Nothing
End Function

Public Sub adjustClip()

    Dim DataObj As New MSForms.DataObject
    Dim clip_text As String
    On Error GoTo Err_handler
    DataObj.GetFromClipboard
    If InStr(1, DataObj.GetText(1), vbCr & vbCr) > 0 Then
        clip_text = Replace(DataObj.GetText(1), vbCr & vbCr, vbCr)
        DataObj.SetText clip_text
        DataObj.PutInClipboard
    End If
Err_handler:
    Set DataObj = Nothing
End Sub


Sub CopyTextToClipboard(txt As String)

    Dim DataObj As New MSForms.DataObject
    On Error GoTo Err_handler
    'Make object's text equal above string variable
    DataObj.SetText txt

    'Place DataObject's text into the Clipboard
    DataObj.PutInClipboard
Err_handler:
    Set DataObj = Nothing

End Sub
Public Function ReadPathValue(ByVal sPath As String) As DV_Value
    Dim lngErrNumberStr As Long
    Dim strErrDescriptionStr As String
    Dim lngErrNumberFlt As Long
    Dim strErrDescriptionFlt As String
    Dim sPathParam As String
    Dim sPathField As String
    
    If InStr(1, sPath, ".") > 0 Then
        sPathParam = Left(sPath, InStr(1, sPath, "."))
        sPathField = Right(sPath, Len(sPath) - Len(sPathParam))
    Else
        sPathParam = sPath & "."
        sPathField = "CV"
    End If
    
    'read string value
    ReadPathValue.strResult = frsReadValue("DVSYS." & sPathParam & "A_" & sPathField, lngErrNumberStr, strErrDescriptionStr, False)
    ReadPathValue.fltResult = CDbl(frsReadValue("DVSYS." & sPathParam & "F_" & sPathField, lngErrNumberFlt, strErrDescriptionFlt, False))
    If lngErrNumberStr <> 0 And lngErrNumberFlt <> 0 Then
        ReadPathValue.sReadErrorMsg = "Error:" & strErrDescriptionStr
    End If
End Function

Public Function WritePathValue(ByVal sPath As String, sValue As String, Optional Float As Boolean = False) As Boolean
    Dim lngErrNumber As Long
    Dim strErrDescription As String
    Dim sPathParam As String
    Dim sPathField As String
    
    If InStr(1, sPath, ".") > 0 Then
        sPathParam = Left(sPath, InStr(1, sPath, "."))
        sPathField = Right(sPath, Len(sPath) - Len(sPathParam))
    Else
        sPathParam = sPath & "."
        sPathField = "CV"
    End If
    
    'read string value
    If Float Then
        'frsWriteValue sValue, "DVSYS." & sPathParam & "F_" & sPathField, lngErrNumber, strErrDescription, False
        sPathParam = sPathParam & "F_" & sPathField
    Else
        'frsWriteValue sValue, "DVSYS." & sPathParam & "A_" & sPathField, lngErrNumber, strErrDescription, False
        sPathParam = sPathParam & "A_" & sPathField
    End If
    
    frsWriteValue sValue, "DVSYS." & sPathParam, lngErrNumber, strErrDescription, False
    
    If lngErrNumber <> 0 Then
        MsgBox "Error: " & strErrDescription & vbCrLf & "Failed writing value: " & sValue & vbCrLf & "To : " & sPathParam, vbCritical, sTitle
        WritePathValue = False
    Else
        WritePathValue = True
    End If
    'If lngErrNumberStr = 0 Then
    'Else
    'End If
End Function


