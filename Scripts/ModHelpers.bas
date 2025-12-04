Attribute VB_Name = "modHelpers"
Option Explicit

Public Function getCol(ByVal colLetter As String) As Long
    getCol = Range(colLetter & "1").Column
End Function

Public Function getColLetter(colNum As Long) As String
    getColLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Public Function getCell(ByVal rowNum As Long, colLetter As String) As Range
    Set getCell = ActiveSheet.Cells(rowNum, getCol(colLetter))
End Function

