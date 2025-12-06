Attribute VB_Name = "ClearCheckboxes"
Option Explicit

Sub ClearCheckboxes()
    Dim r As Range
    Dim allCells As Range
    Dim cb As CheckBox
    Dim ws As Worksheet
    Dim cbName As String
    Dim answer As VbMsgBoxResult
    Dim lastRow As Long
    
    answer = MsgBox("Ви дійсно хочете зняти позначки?", _
                    vbYesNo + vbQuestion, _
                    "Підтвердження")
    
    If answer = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    
    ws.Unprotect
    
    lastRow = ws.usedRange.Rows(ws.usedRange.Rows.Count).Row
    'lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    Set allCells = ws.Range("D4:O" & lastRow)
    
    For Each r In allCells 'Selection
        cbName = "CB_" & r.Address(False, False)
        
        On Error Resume Next
            Set cb = ws.CheckBoxes(cbName)
        On Error GoTo 0
        
        If Not cb Is Nothing Then
            cb.Value = xlOff
        End If
        
        r.Value = ""
    Next r
    
    ws.Protect AllowFiltering:=True
    
    Application.ScreenUpdating = True
End Sub
