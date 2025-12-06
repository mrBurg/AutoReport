Attribute VB_Name = "AddCheckboxes"
Option Explicit

Sub AddCheckCheckboxes()
    Dim r As Range
    Dim cb As CheckBox
    Dim cbName As String
    Dim ws As Worksheet
    Dim colWidth As Long
    
    Application.ScreenUpdating = False
    
    Set ws = ActiveSheet
    
    ws.Unprotect
    
    colWidth = 15

    For Each r In Selection
        cbName = "CB_" & r.Address(False, False)
        r.columnWidth = colWidth

        Set cb = ws.CheckBoxes.Add(r.Left, r.Top, r.Width, r.Height)
        
        With cb
            .name = cbName
            .Caption = ""
            .Placement = xlMoveAndSize
            .OnAction = "CheckCheckboxHandler"
            .Value = xlOn
        End With
        
        With cb.ShapeRange
            .Top = r.Top
            .Left = r.Left
            .Width = r.Width
            .Height = r.Height
        End With
        
        With r
            .Value = "Показати"
            .IndentLevel = 2
        End With
        
    Next r
    
    ws.Protect AllowFiltering:=True
    
    Application.ScreenUpdating = True
End Sub
