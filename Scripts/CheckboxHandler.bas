Attribute VB_Name = "CheckboxHandler"
Option Explicit

Sub CheckCheckboxHandler()
    Dim cb As CheckBox
    Dim r As Range
    Dim ws As Worksheet
    Dim cbName As String
    
    Set ws = ActiveSheet
    
    ws.Unprotect
    
    cbName = CStr(Application.Caller)
    
    On Error Resume Next
        Set cb = ws.CheckBoxes(cbName)
    On Error GoTo 0
    
    If Not cb Is Nothing Then
        Set r = cb.TopLeftCell
        
        If cb.Value = xlOn Then
            r.Value = "Показати"
        Else
            r.Value = "Приховати"
        End If
    End If
    
    ws.Protect AllowFiltering:=True
End Sub

Sub DateCheckboxHandler()
    Dim cb As CheckBox
    Dim r As Range
    Dim ws As Worksheet
    Dim cbName As String
    
    Set ws = ActiveSheet
    
    ws.Unprotect
    
    cbName = CStr(Application.Caller)

    On Error Resume Next
        Set cb = ws.CheckBoxes(cbName)
    On Error GoTo 0
    
    If Not cb Is Nothing Then
        Set r = cb.TopLeftCell
    
        If cb.Value = xlOn Then
            r.Value = Date
            r.NumberFormat = "dd.mm"
        Else
            r.Value = ""
        End If
    End If
    
    ws.Protect AllowFiltering:=True
End Sub
