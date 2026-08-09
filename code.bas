' 1. DELETE USELESS COLUMNS
Sub Remove_useless_columns()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("soil_pollution_diseases")

    'Region: it does not match the real geography
    'Crop_Type, Farming_Practice, Health_Symptoms, Soil_Texture:
    'irrelevant to our analysis

    ws.Columns("K").EntireColumn.Delete

    For i = 1 To 19

        If ws.Cells(1, i).Value = "Region" Or _
           ws.Cells(1, i).Value = "Crop_Type" Or _
           ws.Cells(1, i).Value = "Farming_Practice" Or _
           ws.Cells(1, i).Value = "Health_Symptoms" Or _
           ws.Cells(1, i).Value = "Soil_Texture" Then

            ws.Columns(i).EntireColumn.Delete

        End If

    Next i

End Sub


' 2. REMOVE DUPLICATES
Sub Remove_duplicates()

    Dim ws As Worksheet
    Dim i As Long
    Dim last_row As Long

    Set ws = ThisWorkbook.Sheets("soil_pollution_diseases")

    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 3 To last_row

        If ws.Cells(i, "C") = ws.Cells(i - 1, "C") And _
           ws.Cells(i, "D") = ws.Cells(i - 1, "D") And _
           ws.Cells(i, "E") = ws.Cells(i - 1, "E") Then

            ws.Rows(i).EntireRow.Delete

        End If

    Next i

End Sub


' 3. HANDLE MISSING VALUES

Sub Handle_missing_values()

    Dim ws As Worksheet
    Dim last_row As Long
    Dim last_column As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long

    Set ws = ThisWorkbook.Sheets("soil_pollution_diseases")

    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    last_column = ws.Cells(1, Columns.Count).End(xlUp).Column

    'since most of the values are not related to each other on the space and time
    '(different dates/countries), using the mean of the previous and next values
    'to fill missing ones will...

    For i = 3 To last_row

        For j = 1 To last_column

            If IsEmpty(ws.Cells(i, j)) Then

                If IsNumeric(ws.Cells(1, j).Value) Then

                    n = 0

                    For k = 2 To last_row

                        If ws.Cells(i, "C").Value = ws.Cells(k, "C").Value And _
                           ws.Cells(i, "D").Value = ws.Cells(k, "D").Value Then

                            ws.Cells(i, j).Value = ws.Cells(i, j).Value + k
                            n = n + 1

                        End If

                    Next k

                    If n = 0 Then
                        ws.Rows(i).EntireRow.Delete
                    Else
                        ws.Cells(i, j).Value = ws.Cells(i, j).Value / n
                    End If

                Else

                    ws.Rows(i).EntireRow.Delete

                End If

            End If

        Next j

    Next i

End Sub



' 4. SORTING DATES

Sub Sorting_Dates() ' (Using Record Macro)

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "9/12/2024"

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "9/12/2024"

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "9/12/2024"

    Range("B2:B3001").Select

    ActiveWorkbook.Worksheets("soil_pollution_diseases").Sort.SortFields.Clear

    ActiveWorkbook.Worksheets("soil_pollution_diseases").Sort.SortFields.Add Key _
        :=Range("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal

    With ActiveWorkbook.Worksheets("soil_pollution_diseases").Sort

        .SetRange Range("A2:S3001")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply

    End With

End Sub



' 5. CALCULATE STATISTICS


Private Sub CommandButton1_Click()

    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("soil_pollution_diseases")

    m = WorksheetFunction.Average(Range("E2:E3001"))

    MsgBox "Mean of pollutant Concentration = " & m

End Sub


'  
' 6. FORECAST FUTURE VALUES
'  

Sub Bouton10_Cliquer()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nextT As Double
    Dim nextValue As Double

    Set ws = ThisWorkbook.Worksheets("soil_pollution_diseases")

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    nextT = lastRow + 1

    nextValue = WorksheetFunction.Forecast( _
        nextT, _
        ws.Range("F2:F" & lastRow), _
        ws.Range("B2:B" & lastRow))

    MsgBox "Forecast Pollutant Concentration for next time step (" & _
        nextT & ") = " & Format(nextValue, "0.00") & " mg/kg"

End Sub



' 7. HIGHLIGHT ECOLOGICAL-RISK THRESHOLDS


Sub Bouton2_Cliquer()

    Dim i As Long
    Dim last_row As Long
    Dim n As Long
    Dim percentage As Double
    Dim x As Long
    Dim j As Long

    last_row = Cells(Rows.Count, 1).End(xlUp).Row

    n = 0

    For i = 2 To last_row

        If Cells(i, "E").Value > 151.5 Then

            n = n + 1
            Cells(i, "E").Interior.Color = vbRed

        End If

        For j = 1 To 20

            If Cells(1, j).Value = "Disease_Severity" Then
                x = j
            End If

        Next j

        If Cells(i, x).Value = "Mild" Then

            Cells(i, x).Interior.Color = vbGreen

        ElseIf Cells(i, x).Value = "Moderate" Then

            Cells(i, x).Interior.Color = vbYellow

        Else

            Cells(i, x).Interior.Color = vbRed

        End If

    Next i

    total_rows = last_row - 1
    percentage = (n / total_rows) * 100

    MsgBox "The upper quartile of Pollutant Concentration is about 151.5 mg/kg. " & _
           n & " cases (" & percentage & "% of the dataset) are above this level and " & _
           "will be flagged as HIGH RISK (highlighted in red)"

End Sub
