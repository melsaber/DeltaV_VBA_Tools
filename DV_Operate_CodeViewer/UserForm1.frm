VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "CodeTestHelper"
   ClientHeight    =   10680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13260
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'/*
' * Copyright (c) 2019-2020, Mohamed Elsaber
' * 
' * use my code as per the GNU GPL V3.0
' */
Option Explicit
Private selectionProcessing As Boolean
Private frmRightx As Single
Private frmButtomY As Single, activeResize As Boolean
Private Sub CheckBox1_Click()

End Sub

Private Sub ckReadValues_Change()
    frmToolTip.Visible = False
End Sub

Private Sub ckReadValues_Click()

End Sub

Private Sub CommandButton2_Click()
    Call WritePathValue(selPath.Text, txtValueFlt.Text, True)
End Sub

Private Sub CommandButton3_Click()
    Call WritePathValue(selPath.Text, txtValueString.Text)
End Sub

Private Sub fmtRight_Click()

End Sub

Private Sub fmtRight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    frmRightx = X
    activeResize = True
End Sub

Private Sub fmtRight_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If activeResize And (txtCode.Width > 400 Or X > frmRightx) Then
        fmtRight.Left = fmtRight.Left + X - frmRightx
        Me.Width = Me.Width + X - frmRightx
        txtCode.Width = txtCode.Width + X - frmRightx
        ftmBottom.Width = ftmBottom.Width + X - frmRightx
    End If
End Sub

Private Sub fmtRight_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    activeResize = False
End Sub

Private Sub ftmBottom_Click()

End Sub

Private Sub ftmBottom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    frmButtomY = Y
    activeResize = True
End Sub

Private Sub ftmBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If activeResize And (txtCode.Height > 200 Or Y > frmButtomY) Then
        ftmBottom.Top = ftmBottom.Top + Y - frmButtomY
        Me.Height = Me.Height + Y - frmButtomY
        txtCode.Height = txtCode.Height + Y - frmButtomY
        fmtRight.Height = fmtRight.Height + Y - frmButtomY
    End If
End Sub

Private Sub ftmBottom_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    activeResize = False
End Sub

Private Sub txtCode_Change()
    Dim Y As Integer
    frmToolTip.Visible = False
    'If CheckBox1.Value = False Then txtRTF = txtCode.TextRTF
    If txtCode.Tag = "" Then
        txtCode.Tag = 1

        Y = txtCode.SelStart
   
        txtCode.TextRTF = getRTF(txtCode.Text)
    
        txtCode.SelStart = Y
        
        txtCode.Tag = ""
    End If
End Sub

Private Sub txtCode_KeyDown(pKey As Long, ByVal ShiftKey As Integer)
    If pKey = 17 Then
        Call adjustClip  ' to reformat clipboard - before user pontentially ctrl+v
    End If
End Sub

Private Sub txtCode_MouseUp(ByVal Button As Integer, ByVal ShiftKey As Integer, ByVal xMouse As Long, ByVal yMouse As Long)
        frmToolTip.Left = xMouse * 0.75 + txtCode.Left
        frmToolTip.Top = yMouse * 0.75 - frmToolTip.Height * 1.3 + txtCode.Top
        selectionProcessing = False
End Sub

Private Sub txtCode_SelChange()
    If selectionProcessing Or Me.ckReadValues = False Then
        Exit Sub
    End If
    
    
    Dim X As textMatch, Alias As String, AliasResolv As String
    Dim fullpath As String
    Dim pointRes As DV_Value
    selectionProcessing = True
    X = GetCursorMatch(txtCode.SelStart)
    If X.matchType = "DV Reference" Then
        fullpath = Replace(X.matchtext, "'", "")
        If Left(fullpath, 3) = "//#" Then
            Alias = Mid(fullpath, 4, InStr(5, fullpath, "#/") - 4)
            pointRes = ReadPathValue(txtUnit.Text & "/" & Alias)
            If pointRes.sReadErrorMsg = "" Then
                If pointRes.strResult <> "" Then
                    fullpath = Replace(fullpath, "#" & Alias & "#", pointRes.strResult)
                Else
                    lblToolTip.Caption = "#" & Alias & "#" & " : ignored"
                    frmToolTip.Visible = True
                    frmToolTip.Width = lblToolTip.Width * 1.5 + 4
                    selPath.Text = fullpath
                    Exit Sub
                End If
                
            Else
                lblToolTip.Caption = "#" & Alias & "#" & " : " & pointRes.sReadErrorMsg
                frmToolTip.Visible = True
                frmToolTip.Width = lblToolTip.Width * 1.5 + 4
                selPath.Text = fullpath
                Exit Sub
            End If
            fullpath = Right(fullpath, Len(fullpath) - 2)
            
        ElseIf Left(fullpath, 2) = "///" Then
            fullpath = Right(fullpath, Len(fullpath) - 2)
        ElseIf Left(fullpath, 2) = "^/" Then
            fullpath = Left(txtPath.Text, InStrRev(txtPath.Text, "/") - 1) & Right(fullpath, Len(fullpath) - 1)
        ElseIf Left(fullpath, 1) = "/" Then
            fullpath = Left(txtPath.Text, InStr(1, txtPath.Text, "/")) & Right(fullpath, Len(fullpath) - 1)
            
        Else
            fullpath = txtPath.Text & "/" & fullpath
        End If
        txtCode.SelStart = X.matchStart
        txtCode.SelLength = X.matchLength
        selPath.Text = fullpath
        pointRes = ReadPathValue(fullpath)
        If pointRes.sReadErrorMsg = "" Then
            txtValueString.Text = pointRes.strResult
            txtValueFlt.Text = pointRes.fltResult
            If (pointRes.strResult = Trim$(Str$(pointRes.fltResult))) Then
                lblToolTip.Caption = pointRes.strResult
                
            Else
                lblToolTip.Caption = pointRes.strResult & " (" & pointRes.fltResult & ")"
                
            End If
        Else
            lblToolTip.Caption = fullpath & " : " & pointRes.sReadErrorMsg
            txtValueString.Text = pointRes.sReadErrorMsg
        End If
            
        frmToolTip.Width = lblToolTip.Width * 1.5 + 4
        frmToolTip.Visible = True
    Else
        frmToolTip.Visible = False
    End If
End Sub

Private Sub txtRTF_Change()
    'If CheckBox1.Value = True Then
    '    txtCode.TextRTF = txtRTF.Text
    'End If
End Sub

Private Sub UserForm_Activate()
    On Error GoTo errHandler
    If UBound(matches) > 0 Then
        Exit Sub
    End If
errHandler:
    ReDim matches(0)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    Me.Hide
End Sub
