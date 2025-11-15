Attribute VB_Name = "ModuleRows"
Option Explicit

Public Sub WriteRows(wsTarget As Worksheet, cbDates As Object)
    Dim insertRow As Long
    Dim i As Long
    Dim rng As Range
    Dim M As Long, Y As Long
    Dim parts() As String
    Dim daysInMonth As Long
    Dim weekdaysUkr() As Variant
    Dim currentDate As Date
    Dim cols() As Variant
    Dim col As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    weekdaysUkr = Sheets("Params").Range("B2:B8").Value
    
    parts = Split(cbDates.Value, ".")
    M = CLng(parts(0))
    Y = CLng(parts(1))
    daysInMonth = Day(DateSerial(Y, M + 1, 0))
    
    insertRow = 15
    
    cols = Array(3, 4, 7, 8)
    
    For i = daysInMonth To 1 Step -1
        wsTarget.Rows(insertRow + 1 & ":" & insertRow + 1).Insert Shift:=xlDown
        
        currentDate = DateSerial(Y, M, i)
        
        wsTarget.Cells(insertRow + 1, 1).Value = i
        wsTarget.Cells(insertRow + 1, 2).Value = weekdaysUkr(Weekday(currentDate, vbMonday), 1)
        wsTarget.Rows(insertRow + 1).Font.Size = 12
        
        If Weekday(currentDate, vbMonday) <> 7 Then
            With wsTarget.Cells(insertRow + 1, 5)
                .Formula = "=IF(OR(C" & insertRow + 1 & "="""",D" & insertRow + 1 & "=""""),"""",IF(D" & insertRow + 1 & "-C" & insertRow + 1 & "<0,""Ошибка"",D" & insertRow + 1 & "-C" & insertRow + 1 & "))"
                .NumberFormat = "hh:mm"
                .Locked = True
                
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, Formula1:="=E" & insertRow + 1 & "=""Ошибка"""
                .FormatConditions(1).Interior.Color = RGB(255, 0, 0)  ' Красный фон
                .FormatConditions(1).Font.Color = RGB(255, 255, 255)   ' Белый текст для контраста
            End With
            
            With wsTarget.Cells(insertRow + 1, 9)
                .Formula = "=IF(OR(G" & insertRow + 1 & "="""",H" & insertRow + 1 & "=""""),"""",IF(H" & insertRow + 1 & "-G" & insertRow + 1 & "<0,""Ошибка"",H" & insertRow + 1 & "-G" & insertRow + 1 & "))"
                .NumberFormat = "hh:mm"
                .Locked = True
                
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlExpression, Formula1:="=I" & insertRow + 1 & "=""Ошибка"""
                .FormatConditions(1).Interior.Color = RGB(255, 0, 0)
                .FormatConditions(1).Font.Color = RGB(255, 255, 255)
            End With
            
            For Each col In cols
                With wsTarget.Cells(insertRow + 1, col).Validation
                    .Delete
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertInformation, _
                         Operator:=xlBetween, _
                         Formula1:="=Params!$A$2:$A$58"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                    .ErrorTitle = "Увага"
                    .ErrorMessage = "Значення відсутне"
                End With
                
                wsTarget.Cells(insertRow + 1, col).NumberFormat = "hh:mm"
            Next col
            'wsTarget.Cells(insertRow + 1, 8).Validation.InputTitle = ""
            'wsTarget.Cells(insertRow + 1, 8).Validation.InputMessage = ""
        Else
            With wsTarget.Range(wsTarget.Cells(insertRow + 1, 3), wsTarget.Cells(insertRow + 1, 9))
                .Merge
                .Value = "Вихідний"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = False
            End With
        End If
        
        Set rng = wsTarget.Range(wsTarget.Cells(insertRow + 1, 1), wsTarget.Cells(insertRow + 1, 9))
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    Next i
    
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    lastCol = 9
    
    wsTarget.PageSetup.PrintArea = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lastRow, lastCol)).Address
    
    With wsTarget.PageSetup
        '.LeftMargin = Application.InchesToPoints(0.5)
        '.RightMargin = Application.InchesToPoints(0.5)
        '.TopMargin = Application.InchesToPoints(0.75)
        '.BottomMargin = Application.InchesToPoints(0.75)
        '.HeaderMargin = Application.InchesToPoints(0.3)
        '.FooterMargin = Application.InchesToPoints(0.3)
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
        
    'Debug.Print daysInMonth
End Sub
