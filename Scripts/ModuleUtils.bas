Attribute VB_Name = "ModuleUtils"
Option Explicit

Public Function ShortName(fullName As String) As String
    Dim parts() As String
    Dim result As String
    Dim i As Integer

    parts = Split(fullName, " ")
    
    If UBound(parts) < 1 Then
        ShortName = fullName
        
        Exit Function
    End If

    result = parts(0)

    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            result = result & " " & Left(parts(i), 1) & "."
        End If
    Next i

    ShortName = result
End Function

Public Sub FillTemplate(ws As Worksheet, fullName As String, currentDate As String)
    Dim rng As Range
    Dim fullDate As Date
    Dim parts() As String
    Dim monthsUkr As Variant
    Dim wsData As Worksheet
    Dim searchRange As Range
    Dim i As Long
    Dim percentRate As Double
    Dim foundRate As Boolean
    
    Set rng = Union(ws.Range("D9"), ws.Range("H18"))
    monthsUkr = Sheets("Params").Range("C2:C13").Value
    
    ws.Unprotect
    If fullName <> "" Then
        rng.Value = fullName
    Else
        rng.Value = "Ï.².Á"
    End If
    
    If currentDate <> "" Then
        parts = Split(currentDate, ".")
        fullDate = DateSerial(CInt(parts(1)), CInt(parts(0)), 1)
        
        ws.Range("E11").Value = monthsUkr(Month(fullDate), 1)
        ws.Range("F11").Value = Year(fullDate)
    Else
        ws.Range("E11").Value = "Ì³ñÿöü"
        ws.Range("F11").Value = "Ð³ê"
    End If
    
    Set wsData = ThisWorkbook.Sheets("ÍÏÏ")
    Set searchRange = wsData.Range("B2:B47")
    
    foundRate = False
    For i = 1 To searchRange.Rows.Count
        If Trim(ShortName(searchRange.Cells(i, 1).Value)) = fullName Then
            percentRate = wsData.Cells(searchRange.Cells(i, 1).Row, "D").Value
            foundRate = True
            
            Exit For
        End If
    Next i
    
    If foundRate Then
        ws.Range("F9").Value = percentRate
    Else
        ws.Range("F9").Value = "Ñòàâêà"
    End If
    ws.Protect
End Sub

Public Sub CopyTemplate(ws As Worksheet, wsTarget As Worksheet)
    ws.Range("A1:I18").Copy
    wsTarget.Range("A1").PasteSpecial xlPasteAll
    wsTarget.Range("A1").PasteSpecial xlPasteColumnWidths
        
    Application.CutCopyMode = False
End Sub

Private Function SheetNameToDate(s As String) As Date
    Dim p1 As String, p2 As String
    Dim Y As Long, M As Long
    
    If s = "AutoReport" Or s = "ÍÏÏ" Or s = "Params" Then
        SheetNameToDate = DateSerial(1900, 1, 1)
        Exit Function
    End If
        
    p1 = Split(s, ".")(0)
    p2 = Split(s, ".")(1)
    
    If Len(p1) = 4 Then
        Y = CLng(p1)
        M = CLng(p2)
    Else
        Y = CLng(p2)
        M = CLng(p1)
    End If
        
    SheetNameToDate = DateSerial(Y, M, 1)
End Function

Public Sub SortSheetsByDate()
    Dim i As Long, j As Long
    Dim d1 As Date, d2 As Date
    Dim s1 As String, s2 As String

    For i = 1 To ThisWorkbook.Sheets.Count - 1
        s1 = ThisWorkbook.Sheets(i).Name
        If s1 = "AutoReport" Or s1 = "ÍÏÏ" Or s1 = "Params" Then GoTo ContinueI

        For j = i + 1 To ThisWorkbook.Sheets.Count

            s2 = ThisWorkbook.Sheets(j).Name
            If s2 = "AutoReport" Or s2 = "ÍÏÏ" Or s2 = "Params" Then GoTo ContinueJ

            d1 = SheetNameToDate(s1)
            d2 = SheetNameToDate(s2)

            If d2 > d1 Then
                ThisWorkbook.Sheets(j).Move Before:=ThisWorkbook.Sheets(i)
            End If

ContinueJ:
        Next j

ContinueI:
    Next i
End Sub

'Private Sub MoveSpecialSheetsToTop()
    'Dim ws As Worksheet
    
    'For Each ws In ThisWorkbook.Sheets
        'If ws.Name = "ÍÏÏ" Then
            'ws.Move Before:=ThisWorkbook.Sheets(1)
        'End If
    'Next ws

    'For Each ws In ThisWorkbook.Sheets
        'If ws.Name = "AutoReport" Then
            'ws.Move Before:=ThisWorkbook.Sheets(1)
        'End If
    'Next ws
'End Sub

'Public Sub SortSheetsWithPriority()
    'Debug.Print "RUN: SortSheetsWithPriority"
    'SortSheetsByDate
    'MoveSpecialSheetsToTop
'End Sub
