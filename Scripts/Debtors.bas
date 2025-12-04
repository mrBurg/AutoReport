Attribute VB_Name = "Debtors"
Option Explicit

Sub MoveRowsWithNotAllTrue()
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim r As Long
    Dim allTrue As Boolean
    Dim sheetName As String
    Dim destName As String
    Dim startRow As Long
    Dim destRow As Long
    Dim legendStart As Long
    Dim headerRange As Range
    Dim allCells As Range
    Dim cb As CheckBox
    Dim cbName As String
    Dim colLetter As String

    Application.CutCopyMode = False
    Application.ReferenceStyle = xlA1
    Application.ScreenUpdating = False
    
    sheetName = ActiveSheet.name
    
    Set wsSrc = ThisWorkbook.Sheets(sheetName)
    
    destName = "Боржники " & sheetName
    
    On Error Resume Next
        Set wsDest = ThisWorkbook.Sheets(destName)
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add
        
        wsDest.name = destName
        
        With wsDest.Cells
            .Font.name = "Times New Roman"
            .Font.Size = 12
        End With
    End If
    
    wsDest.Cells.Clear
    
    Set headerRange = wsSrc.Range("A1:P3")
    
    headerRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
        
    For c = 1 To headerRange.Columns.Count
        wsDest.Columns(c).columnWidth = wsSrc.Columns(c).columnWidth
    Next c
    
    For r = 1 To headerRange.Rows.Count
        wsDest.Rows(r).RowHeight = wsSrc.Rows(r).RowHeight
    Next r
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column - 1
    
    startRow = 4
    destRow = 4
    
    For i = startRow To lastRow
        allTrue = True
        cbName = "CB_Q" & i
        
        Set cb = Nothing
        
        On Error Resume Next
            Set cb = wsSrc.CheckBoxes(cbName)
        On Error GoTo 0
        
        If cb Is Nothing Then
            GoTo SkipRow
        ElseIf cb.Value <> xlOn Then
            GoTo SkipRow
        End If
        
        For j = 4 To lastCol
            colLetter = getColLetter(j)
            cbName = "CB_" & colLetter & i
            
            Set cb = Nothing
            
            On Error Resume Next
                Set cb = wsSrc.CheckBoxes(cbName)
            On Error GoTo 0
            
            If Not cb Is Nothing Then
                If cb.Value <> xlOn Then
                    allTrue = False

                    Exit For
                End If
            Else
                allTrue = False
                
                Exit For
            End If
        Next j
        
        If Not allTrue Then
            With wsSrc.Rows(i).Range("A1").Resize(1, lastCol + 1)
                wsDest.Rows(destRow).Range("A1").Resize(1, lastCol + 1).Value = .Value
                .Copy
                wsDest.Rows(destRow).Range("A1").PasteSpecial xlPasteFormats
            End With
            
            With wsDest
                .Cells(destRow, getCol("A")).Formula = "=ROW() - ROW(A" & startRow & ") + 1"
                .Cells(destRow, getCol("A")).Borders.LineStyle = xlContinuous
            End With
            
            destRow = destRow + 1
        End If
    
    Set allCells = wsDest.Range("D4:O" & lastRow)
    
    allCells.FormatConditions.Delete
    allCells.FormatConditions.Add _
        Type:=xlExpression, _
        Formula1:="=LEN(Trim(" & allCells.Cells(1, 1).Address(False, False) & ")) > 0"
    allCells.FormatConditions(1).Interior.Color = RGB(128, 128, 128)
SkipRow:
    Next i
    
    legendStart = destRow + 1
    
    With wsDest
        .Cells(legendStart, "C").Value = "Здано"
        .Cells(legendStart, "B").Interior.Color = RGB(128, 128, 128)
        ' .Cells(legendStart, "B").Font.Color = RGB(128, 128, 128)
        
        .Cells(legendStart + 1, "C").Value = "Не здано"
        ' .Cells(legendStart + 1, "B").Interior.Color = RGB(255, 255, 255)
        ' .Cells(legendStart + 1, "B").Font.Color = RGB(255, 255, 255)
        ' .Cells(legendStart + 1, "C").Font.Bold = True
        
        .Range(.Cells(legendStart, "B"), .Cells(legendStart + 1, "C")).Borders.LineStyle = xlContinuous
        .Columns("B:C").AutoFit
    End With
    
    wsDest.Activate
    
    MsgBox "Знайдено " & destRow - startRow & " боржників.", vbInformation
    
    Application.ScreenUpdating = True
End Sub

