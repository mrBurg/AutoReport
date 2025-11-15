Attribute VB_Name = "ModuleCreateReport"
Option Explicit

Public Sub CreateReport(ws As Worksheet)
    Dim cbNames As Object
    Dim cbDates As Object
    Dim wsTarget As Worksheet
    Dim sheetName As String
    
    Set cbNames = ws.OLEObjects("AcademicStaff").Object
    Set cbDates = ws.OLEObjects("UserDate").Object
        
    sheetName = cbDates.Value
    
    If sheetName = "" Then
        MsgBox "Оберіть дату (Ім’я листа)!"
        
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not wsTarget Is Nothing Then
        wsTarget.Cells.Clear
        
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
End Sub
