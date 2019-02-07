'Purpose:
'   a. Analyze ingredient information provided by supplier 
'   b. Ensure the ingredients will be able to save on the Walgreens website.
Sub moveSKUs()

    Dim MainSheet as Worksheet, MatchSheet as Worksheet, ActiveIngredientSheet As Worksheet
    Dim activeIngredientString As String, inactiveIngredientString As String, issue As Boolean

    Set MainSheet = Sheet1
    Set MatchSheet = Sheet41
    Set ActiveIngredientSheet = Sheet6
    MainSheet.Range("T2:T" & lastRow(MainSheet) + 1).ClearContents
    MatchSheet.Range("A2:B" & lastRow(MatchSheet) + 1).ClearContents

    For Each sku In MainSheet.Range("A2:A" & lastRow(MainSheet))
        Dim separatedString As New Collection
        Dim x, y As Long
        y = 1
        issue = False

        'See if item has more than one set of ingredients
        For x = 5 To 17 Step 3
            If MainSheet.Cells(sku.Row, x) <> "" Or MainSheet.Cells(sku.Row, x + 1) <> "" Or MainSheet.Cells(sku.Row, x + 2) <> "" Then
                y = y + 1
            End If
        Next x

        'Loop through all ingredient sets for the item
        For x = 1 To y
            activeIngredientString = Trim(MainSheet.Cells(sku.Row, x * 3))
            inactiveIngredientString = Trim(MainSheet.Cells(sku.Row, x * 3 + 1))
            inactiveIngredientString = cleanString(inactiveIngredientString)
            Call AmbiguousCharacters(inactiveIngredientString)
            If Not issue Then
                Set separatedString = stringToCollection(inactiveIngredientString)
                Call DuplicateIngredients(separatedString)
            End If
            'Skip to next item if a problem was detected with the ingredients
            If issue Then Exit For
            If Len(activeIngredientString) > 5 Then
                ActiveIngredientSheet.Range("A" & lastRow(ActiveIngredientSheet) + 1) = sku
                ActiveIngredientSheet.Range("A" & lastRow(ActiveIngredientSheet)).Offset(, x) = activeIngredientString
                ActiveIngredientSheet.Range("G" & lastRow(ActiveIngredientSheet)) = Now()
            End If
        Next x
        MatchSheet.Range("A" & lastRow(MatchSheet) + 1) = sku
    Next sku

    If Application.WorksheetFunction.CountA(MainSheet.Range("T:T")) > 1 Then
        'Issue detected with at least one item
        MatchSheet.Visible = xlVeryHidden
        MainSheet.AutoFilter.Sort.SortFields.Clear
        MainSheet.AutoFilter.Sort.SortFields.Add(Range( _
            "A1"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB( _
            255, 80, 80)
        With MainSheet.AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        MsgBox "Issues detected with SKU ingredient entry.  Please re-enter ingredient information for those items. "
    Else
        'No issue with items- proceed to update
        MatchSheet.Visible = True
        Application.GoTo MatchSheet.Range("A1"), True
        MsgBox "No issues separating ingredients.  Proceed to update items in MAP."
    End If
    ActiveIngredientSheet.Range("$A$1:$I$" & lastRow(ActiveIngredientSheet)).RemoveDuplicates Columns:=1, Header:=xlYes
End Sub


'Look for an indredient listed twice - Walgreens website does not allow duplicate ingredients
'(algorithm - "Triangular number")
Sub DuplicateIngredients(ingredientString As Collection)
    Dim a as Long, b As Long
    For a = 1 To ingredientString.Count
        For b = a + 1 To ingredientString.Count
            'check for too long of a string or duplicate ingredients
            If Len(ingredientString(a)) > 80 Then
                MainSheet.Range("T" & sku.Row) = "Item ingredients are missing commas for proper separation"
                issue = True
            ElseIf ingredientString(a) = ingredientString(b) Then
                MainSheet.Range("T" & sku.Row) = "Item has the ingredient '" & ingredientString(a) & "' listed twice"
                issue = True
            End If
            If issue Then Exit Sub
        Next b
    Next a

End Sub


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


'List of strings that should not be in the ingredient description
Function badCharacters() As Variant

    badCharacters = Array("*", ";", "ingredient", "pack", "inactive", "active")

End Function


'Return the last used row in a worksheet
Function lastRow(wksht As Worksheet) As Long

    lastRow = wksht.Cells(Rows.Count, 1).End(xlUp).Row
    
End Function
