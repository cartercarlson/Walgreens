'Test for characters in ingredient string that would prevent ingredients from saving the Walgreens website
Sub AmbiguousCharacters(testString As String)
    testString = LCase(testString)
    For i = 0 To UBound(badCharacters)
        If InStr(testString, badCharacters(i)) Then
            If badCharacters(i) <> "*" Or badCharacters(i) <> ";" Then
                MainSheet.Range("T" & sku.Row) = "Item ingredients should not include the word '" & badCharacters(i) & "'"
            Else
                MainSheet.Range("T" & sku.Row) = "Ambiguous characters detected - '" & badCharacters(i) & "'"
            End If
            issue = True: Exit For
        End If
    Next i
    If InStr(testString, ",") = 0 And InStr(testString, vbLf) = 0 And Len(testString) > 80 Then
        MainSheet.Range("T" & sku.Row) = "Item ingredients are not separated. Separation should be done by commas"
        issue = True
    End If
End Sub
