Attribute VB_Name = "Debtors"
Option Explicit

Sub MoveRowsWithNotAllTrue()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim startRow As Long, lastRow As Long
    Dim startCol As Long, lastCol As Long
    Dim i As Long, j As Long, c As Long, r As Long
    Dim allTrue As Boolean
    Dim sheetName As String, destName As String
    Dim destRow As Long
    Dim legendStart As Long
    Dim headerRange As Range
    Dim rowRange As Range, rowDestRange As Range
    Dim cb As CheckBox
    Dim cbName As String
    Dim colLetter As String

    Application.CutCopyMode = False
    Application.ReferenceStyle = xlA1
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
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
    
    startCol = 4
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column - 1
    startRow = 4
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    
    destRow = startRow
    
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
        
        For j = startCol To lastCol
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
        
        Set rowRange = wsSrc.Rows(i).Range("A1").Resize(1, lastCol + 1)
        Set rowDestRange = wsDest.Rows(destRow).Range("A1")
        
        If Not allTrue Then
            With rowRange
                rowDestRange.Resize(1, lastCol + 1).Value = .Value
                '.Copy
                'rowDestRange.PasteSpecial xlPasteFormats
            End With
            
            With wsDest.Cells(destRow, getCol("A"))
                .Formula = "=ROW() - ROW(A" & startRow & ") + 1"
                '.Borders.LineStyle = xlContinuous
            End With
            
            destRow = destRow + 1
        End If
SkipRow:
    Next i
    
    lastRow = wsDest.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    
    wsSrc.Range("A4:P" & lastRow).Copy
    wsDest.Range("A4").PasteSpecial xlPasteFormats
    
    legendStart = destRow + 1
    
    With wsDest
        .Cells(legendStart, "C").Value = "Здано"
        .Cells(legendStart, "B").Interior.Color = RGB(128, 128, 128)
        '.Cells(legendStart, "B").Font.Color = RGB(128, 128, 128)
        
        .Cells(legendStart + 1, "C").Value = "Не здано"
        .Cells(legendStart + 1, "B").Interior.Color = RGB(255, 255, 255)
        '.Cells(legendStart + 1, "B").Font.Color = RGB(255, 255, 255)
        '.Cells(legendStart + 1, "C").Font.Bold = True
        
        .Range(.Cells(legendStart, "B"), .Cells(legendStart + 1, "C")).Borders.LineStyle = xlContinuous
        .Columns("B:C").AutoFit
    End With
    
    With wsDest.PageSetup
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    With wsDest
        .PageSetup.PrintArea = ""
        .PageSetup.PrintArea = .UsedRange.Address(True, True)
        .Activate
    End With
    
    ActiveWindow.View = xlPageBreakPreview
    
    MsgBox "Знайдено " & destRow - startRow & " боржників.", vbInformation
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



