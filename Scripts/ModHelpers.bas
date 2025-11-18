Attribute VB_Name = "modHelpers"
Public Function Col(colLetter As String) As Long
    Col = Range(colLetter & "1").Column
End Function

Public Function Cell(rowNum As Long, colLetter As String) As Range
    Set Cell = ActiveSheet.Cells(rowNum, Col(colLetter))
End Function
