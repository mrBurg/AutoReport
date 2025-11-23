Attribute VB_Name = "ModuleCreateReport"
Option Explicit

Public Sub CreateReport(ws As Worksheet)
    Dim cbNames As Object
    Dim cbDates As Object
    Dim wsTarget As Worksheet
    Dim wsSchedule As Worksheet
    Dim sheetName As String
    Dim weekDataStart As Long
    Dim weekDataLen As Long
    Dim weekDataCount As Long
    Dim lastRow As Long
    Dim cellValue As String
    Dim startDateStr As String, endDateStr As String
    Dim startDate As Date, endDate As Date
    
    Set cbNames = ws.OLEObjects("AcademicStaff").Object
    Set cbDates = ws.OLEObjects("UserDate").Object
        
    sheetName = cbDates.Value
    
    If sheetName = "" Then
        MsgBox "Оберіть дату (Ім’я листа)!"
        
        Exit Sub
    End If
    
    Set wsSchedule = ThisWorkbook.Sheets("Schedule")
    
    weekDataStart = 2
    weekDataLen = 43
    weekDataCount = weekDataStart
    
    lastRow = wsSchedule.UsedRange.Rows(wsSchedule.UsedRange.Rows.Count).Row
    
    Do While weekDataCount <= lastRow
        cellValue = wsSchedule.Range("C" & weekDataCount).Value
        
        If cellValue <> "" Then
            If InStr(cellValue, "з ") > 0 And InStr(cellValue, " по ") > 0 Then
                startDateStr = Trim(Split(Split(cellValue, "з ")(1), " по ")(0))
                endDateStr = Trim(Split(Split(Split(cellValue, " по ")(1), " з")(0), " р.")(0))
                                
                startDate = DateValue(startDateStr)
                endDate = DateValue(endDateStr)
                
                If Format(startDate, "mm.yyyy") = cbDates.Value Or Format(endDate, "mm.yyyy") = cbDates.Value Then
                    Debug.Print cellValue
                End If
            End If
        End If
        
        weekDataCount = weekDataCount + weekDataLen
    Loop
    
    'GoTo Finish
    
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not wsTarget Is Nothing Then
        wsTarget.Unprotect
        wsTarget.Cells.Clear
        wsTarget.Cells.ClearFormats
        
        'MsgBox "Лист " & wsTarget.Name & " оновлено"
    Else
        Set wsTarget = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsTarget.Name = sheetName
        
        'MsgBox "Створено нового листа: " & wsTarget.Name
    End If
    
    CopyTemplate ws, wsTarget
    SortSheetsByDate
    
    ModuleRows.WriteRows wsTarget, cbDates
    
    wsTarget.Activate
    
Finish:
End Sub
