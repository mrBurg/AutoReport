Attribute VB_Name = "ModuleControls"
Option Explicit

Public Sub FillAcademicStaff(ws As Worksheet, cbNames As Object)
    Dim wsData As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    Set wsData = ThisWorkbook.Sheets("моо")
    Set rng = wsData.Range("B2:B47")
    
    cbNames.AddItem
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            cbNames.AddItem ShortName(cell.Value)
        End If
    Next cell
End Sub

Public Sub FillDates(cbDates As Object)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    
    startDate = DateSerial(2024, 9, 1)
    endDate = DateSerial(2027, 8, 1)
    currentDate = startDate
    
    cbDates.AddItem
    Do While currentDate <= endDate
        cbDates.AddItem Format(currentDate, "mm.yyyy")
        currentDate = DateAdd("m", 1, currentDate)
    Loop
End Sub

Public Sub FillAll(ws As Worksheet)
    Dim cbNames As Object
    Dim cbDates As Object
    
    Set cbNames = ws.OLEObjects("AcademicStaff").Object
    Set cbDates = ws.OLEObjects("UserDate").Object
    
    cbNames.Clear
    cbDates.Clear
    
    FillAcademicStaff ws, cbNames
    FillDates cbDates
    CheckFields ws
End Sub

Public Sub CheckFields(ws As Worksheet)
    Dim cbNames As Object
    Dim cbDates As Object
    Dim btn As OLEObject
    
    Set cbNames = ws.OLEObjects("AcademicStaff").Object
    Set cbDates = ws.OLEObjects("UserDate").Object
    Set btn = ws.OLEObjects("SubmitBtn")
    
    If Trim(cbNames.Value) = "" Or Trim(cbDates.Value) = "" Then
        btn.Enabled = False
    Else
        btn.Enabled = True
    End If
    
    FillTemplate ws, cbNames.Value, cbDates.Value
End Sub
