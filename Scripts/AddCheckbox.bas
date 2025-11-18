Attribute VB_Name = "AddCheckbox"
Sub InsertCheckboxesRight()
    Dim Cell As Range
    Dim cb As Object
    Dim cbWidth As Double
    Dim leftPos As Double, topPos As Double

    cbWidth = 14 ' ширина чекбокса

    For Each Cell In Selection
        ' вычисляем координаты
        leftPos = Cell.Left + Cell.Width - cbWidth - 2
        topPos = Cell.Top + (Cell.Height - cbWidth) / 2

        ' создаём чекбокс
        Set cb = ActiveSheet.CheckBoxes.Add(leftPos, topPos, cbWidth, cbWidth)

        ' задаём свойства
        With cb
            .Caption = ""
            .LinkedCell = Cell.Address
            .Placement = xlMoveAndSize
            .Value = xlOn  ' ставим флажок в состояние "снят"
        End With
    Next Cell
End Sub
