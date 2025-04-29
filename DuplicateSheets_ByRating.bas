Sub DuplicateTemp_ForAllRows_WithRatingNames()

    Dim wsData As Worksheet
    Dim wsTemplate As Worksheet
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim rating As String
    Dim sheetName As String

    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsTemplate = ThisWorkbook.Sheets("Temp")

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        wsTemplate.Copy After:=Sheets(Sheets.Count)
        Set newSheet = ActiveSheet

        For Each cell In newSheet.Range("B2:B8")
            If cell.HasFormula Then
                cell.Formula = UpdateRowInFormula(cell.Formula, i)
            End If
        Next cell

        DoEvents

        rating = newSheet.Range("B5").Value
        sheetName = rating

        On Error Resume Next
        newSheet.Name = sheetName
        If Err.Number <> 0 Then
            sheetName = rating & "_" & i
            newSheet.Name = sheetName
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    MsgBox "All sheets created and named by Rating!", vbInformation
End Sub

Function UpdateRowInFormula(formula As String, newRow As Long) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "(Data!\$?[A-Z]+\$?)(\d+)"
    regex.Global = True

    UpdateRowInFormula = regex.Replace(formula, "$1" & newRow)
End Function