Sub colour_and_bold_rows_from_rock_type()

Sheet1.Select

'Format for transcribing the last row and column. Useful for when the sheet has many of rows/columns
LastRow = Range("A1").End(xlDown).Row
lastcol = Range("A1").End(xlToRight).Column

'If a row has either Granite, Clay or Limestone in col15,
'the entire row becomes bold and coloured
For i = 1 To LastRow
    If Cells(i, 15).Value = "Granite" Then
        Cells(i, 15).Interior.ColorIndex = 37
        Range(Cells(i, 1), Cells(i, lastcol)).Interior.ColorIndex = 37
        Range(Cells(i, 1), Cells(i, lastcol)).Font.Bold = True
    ElseIf Cells(i, 15).Value = "Clay" Then
        Cells(i, 15).Interior.ColorIndex = 40
        Range(Cells(i, 1), Cells(i, lastcol)).Interior.ColorIndex = 40
        Range(Cells(i, 1), Cells(i, lastcol)).Font.Bold = True
    ElseIf Cells(i, 15).Value = "Phosphate" Then
        Cells(i, 15).Interior.ColorIndex = 6
        Range(Cells(i, 1), Cells(i, lastcol)).Interior.ColorIndex = 6
        Range(Cells(i, 1), Cells(i, lastcol)).Font.Bold = True
    ElseIf Cells(i, 15).Value = "Limestone" Then
        Cells(i, 15).Interior.ColorIndex = 43
        Range(Cells(i, 1), Cells(i, lastcol)).Interior.ColorIndex = 43
        Range(Cells(i, 1), Cells(i, lastcol)).Font.Bold = True
    End If
Next i
analyze_year_of_column_34_references

End Sub

' ---------------------------------------------------------------------------------

Sub analyze_year_of_column_34_references()

'Loop over all rows in sheet
For i = 1 To Range("a1").End(xlDown).Row
    cellValue = Cells(i, 34).Value
    openingParen = InStr(cellValue, "(")
    If InStr(cellValue, ")") = 0 Then
        closingParen = 1
    Else: closingParen = InStr(cellValue, ")")
    End If
    'Retrieve substring position of open and closed brackets within cell text
    'Assign string within brackets to enclosedValue
    enclosedValue = Mid(cellValue, openingParen + 1, closingParen - openingParen - 1)
    If IsNumeric(enclosedValue) Then
        'Check if enclosedValue (the year) is between 1980 and 2000
        If CInt(enclosedValue) >= 1980 And CInt(enclosedValue) <= 2000 Then
            'If year is within range, change interior cell color to bright pink and bold the cell
            Range(Cells(i, 34), Cells(i, 34)).Interior.ColorIndex = 7
            Range(Cells(i, 34), Cells(i, 34)).Font.Bold = True
        End If
    End If
Next i

End Sub
