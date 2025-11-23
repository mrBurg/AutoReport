Attribute VB_Name = "ModuleControls"
Option Explicit

Public Sub FillAcademicStaff(cbNames As Object)
    Dim wsData As Worksheet
    Dim rng As Range
    Dim Cell As Range
    
    Set wsData = ThisWorkbook.Sheets("SPW")
    Set rng = wsData.Range("B2:B47")
    
    cbNames.AddItem
    
    For Each Cell In rng
        If Not IsEmpty(Cell.Value) Then
            cbNames.AddItem ShortName(Cell.Value)
        End If
    Next Cell
End Sub

Public Sub FillDates(cbDates As Object)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    
    ' startDate = DateSerial(2024, 9, 1)
    startDate = DateAdd("m", -6, Date)
    endDate = DateAdd("m", 6, Date)
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
    
    FillAcademicStaff cbNames
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
