Sub DeleteSheetsWithSameDataFromB11()

    Dim i As Long, j As Long
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim keepSheets As Object
    Dim compareRange1 As Range, compareRange2 As Range
    Dim v1 As Variant, v2 As Variant
    Dim areSame As Boolean

    Set keepSheets = CreateObject("Scripting.Dictionary")
    Application.DisplayAlerts = False

    For i = ThisWorkbook.Sheets.Count To 1 Step -1
        Set ws1 = ThisWorkbook.Sheets(i)
        areSame = False

        With ws1
            Set compareRange1 = .Range("B11", .Cells(.Rows.Count, "C").End(xlUp))
        End With

        For Each key In keepSheets.Keys
            Set ws2 = ThisWorkbook.Sheets(key)
            With ws2
                Set compareRange2 = .Range("B11", .Cells(.Rows.Count, "C").End(xlUp))
            End With

            If compareRange1.Rows.Count = compareRange2.Rows.Count Then
                v1 = compareRange1.Value
                v2 = compareRange2.Value

                areSame = True
                For j = 1 To UBound(v1, 1)
                    If v1(j, 1) <> v2(j, 1) Or v1(j, 2) <> v2(j, 2) Then
                        areSame = False
                        Exit For
                    End If
                Next j

                If areSame Then Exit For
            End If
        Next key

        If areSame Then
            ws1.Delete
        Else
            keepSheets.Add ws1.Name, True
        End If
    Next i

    Application.DisplayAlerts = True
    MsgBox "Duplicate sheets (based on B11:C data) removed!", vbInformation

End Sub