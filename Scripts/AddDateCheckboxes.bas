Attribute VB_Name = "AddDateCheckboxes"
Option Explicit

Sub AddDateCheckboxes()
    Dim r As Range
    Dim monthRange As Range
    Dim cb As CheckBox
    Dim ws As Worksheet
    Dim cbName As String
    Dim tempWsName As String
    Dim colWidth As Long
    Dim tmp As Range
    Dim tempSheet As Worksheet
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    
    ws.Unprotect
    
    colWidth = 9

    For Each r In Selection
        cbName = "CB_" & r.Address(False, False)
        r.columnWidth = colWidth

        Set cb = ws.CheckBoxes.Add(r.Left, r.Top, r.Width, r.Height)
        
        With cb
            .name = cbName
            .Caption = ""
            .Placement = xlMoveAndSize
            .OnAction = "DateCheckboxHandler"
            .Value = xlOff
        End With
        
        With cb.ShapeRange
            .Top = r.Top
            .Left = r.Left
            .Width = r.Width
            .Height = r.Height
        End With
    Next r
    
    Set monthRange = Selection
    
    tempWsName = "TempSheet"
    
    On Error Resume Next
        Set tempSheet = ThisWorkbook.Sheets(tempWsName)
    On Error GoTo 0
    
    If tempSheet Is Nothing Then
        Set tempSheet = ThisWorkbook.Sheets.Add
        
        tempSheet.name = tempWsName
    End If
    
    tempSheet.Range("A1").formula = "=LEN(TRIM(" & monthRange.Range("A1").Address(False, False) & ")) > 0"
    
    'ws.Activate
    
    With monthRange
        .FormatConditions.Delete
        .FormatConditions.Add _
            Type:=xlExpression, _
            Formula1:=tempSheet.Range("A1").formulaLocal
        .FormatConditions(1).Interior.Color = RGB(128, 128, 128)
    End With
    
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    
    
    ws.Protect AllowFiltering:=True
    Application.ScreenUpdating = True
End Sub
