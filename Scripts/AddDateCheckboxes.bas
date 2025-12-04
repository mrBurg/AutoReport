Attribute VB_Name = "AddDateCheckboxes"
Option Explicit

Sub AddDateCheckboxes()
    Dim r As Range
    Dim cb As CheckBox
    Dim ws As Worksheet
    Dim cbName As String
    Dim colWidth As Long
    
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
        
    Selection.FormatConditions.Delete
    Selection.FormatConditions.Add _
        Type:=xlExpression, _
        Formula1:="=LEN(Trim(" & Selection.Cells(1, 1).Address(False, False) & ")) > 0"
    Selection.FormatConditions(1).Interior.Color = RGB(128, 128, 128)
    
    
    ws.Protect AllowFiltering:=True
    Application.ScreenUpdating = True
End Sub
