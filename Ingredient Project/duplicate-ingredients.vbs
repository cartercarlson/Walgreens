Sub DuplicateIngredients(ingredientString As Collection)
'**************************************************************************************************************************
'Purpose: ensure that the same ingredient is not listed twice within the set of ingredients - Walgreens website does not
'         allow duplicate ingredients
'Note:    "nth triangular number" algorithm used to check for duplicates
'**************************************************************************************************************************
    Dim a, b As Long
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
