Attribute VB_Name = "Debtors"
Sub MoveRowsWithNotAllTrue()

    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim allTrue As Boolean
    Dim destRow As Long
    Dim sheetName As String
    Dim destName As String
    
    Application.CutCopyMode = False
    Application.ReferenceStyle = xlA1
    Application.ScreenUpdating = False
    
    ' Получаем имя текущего листа
    sheetName = ActiveSheet.Name
    Set wsSrc = ThisWorkbook.Sheets(sheetName)
    
    ' Имя нового листа
    destName = "Боржники " & sheetName
    
    ' Проверяем, есть ли лист назначения
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destName)
    On Error GoTo 0
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add
        wsDest.Name = destName
        With wsDest.Cells
            .Font.Name = "Times New Roman"
            .Font.Size = 12
        End With
    End If
    
    ' Очищаем старые данные
    wsDest.Cells.ClearContents
    
    ' Копируем заголовки (только значения)
    Dim headerRange As Range
    Set headerRange = wsSrc.Range("A1:P3")
    
    headerRange.Copy
    wsDest.Range("A1").PasteSpecial xlPasteAll
        
    ' Ширина
    For c = 1 To headerRange.Columns.Count
        wsDest.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
    Next c
    
    ' Высота
    For R = 1 To headerRange.Rows.Count
        wsDest.Rows(R).RowHeight = wsSrc.Rows(R).RowHeight
    Next R
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column - 1
    
    startRow = 4
    destRow = 4
    
    '=== Основной цикл ===
    For i = 4 To lastRow
        If wsSrc.Cells(i, "Q").Value = False Or wsSrc.Cells(i, "Q").Value = "FALSE" Then
            GoTo SkipRow
        End If
    
        allTrue = True
        
        ' Проверяем только месяцы (начиная с 3-го столбца)
        For j = 4 To lastCol
            If wsSrc.Cells(i, j).Value <> True And wsSrc.Cells(i, j).Value <> "TRUE" Then
                allTrue = False
                Exit For
            End If
        Next j
        
        ' Если не все TRUE — переносим строку (только значения)
        If Not allTrue Then
            ' === Копируем всю строку с исходного листа ===
            ' wsSrc.Rows(i).Copy
        
            ' Вставляем всё (значения, форматы, ширину и т.д.) на целевой лист
            ' wsDest.Rows(destRow).PasteSpecial xlPasteAll
    
             With wsSrc.Rows(i).Range("A1").Resize(1, lastCol + 1)
                wsDest.Rows(destRow).Range("A1").Resize(1, lastCol + 1).Value = .Value
                .Copy
                wsDest.Rows(destRow).Range("A1").PasteSpecial xlPasteFormats
            End With
            
            ' === Добавляем номер по порядку в столбец A ===
            wsDest.Cells(destRow, Col("A")).Formula = "=ROW() - ROW(A" & startRow & ") + 1"
            wsDest.Cells(destRow, Col("A")).Borders.LineStyle = xlContinuous
            
            destRow = destRow + 1
        End If
        
SkipRow:
    Next i
    
    ' === Добавляем легенду TRUE / FALSE ===
    Dim legendStart As Long
    legendStart = destRow + 1
    
    With wsDest
        .Cells(legendStart, "C").Value = "Здано"
        .Cells(legendStart, "B").Interior.Color = RGB(128, 128, 128)
        ' .Cells(legendStart, "B").Font.Color = RGB(128, 128, 128)
        
        .Cells(legendStart + 1, "C").Value = "Не здано"
        ' .Cells(legendStart + 1, "B").Interior.Color = RGB(255, 255, 255)
        ' .Cells(legendStart + 1, "B").Font.Color = RGB(255, 255, 255)
        ' .Cells(legendStart + 1, "C").Font.Bold = True
        
        ' Общий стиль
        .Range(.Cells(legendStart, "B"), .Cells(legendStart + 1, "C")).Borders.LineStyle = xlContinuous
        .Columns("B:C").AutoFit
    End With
    
    MsgBox "Перенос завершён. Перемещено " & destRow - 4 & " строк.", vbInformation

End Sub

