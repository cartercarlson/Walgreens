Const URL = "http://c0merchsvc01p.dmzwalgreenco.net/map2/index.html#/search/product/1/1/"
Const searchResultsElement = "//attribute-grid/div/div[1]/div[5]/span[@class='k-pager-info k-label']"
Const deleteInactiveIngElement = "//ingredients-editor[2]/section/div/div/div/div/div/span[@title='Delete this ingredient']"
Const deleteActiveIngElement = "//ingredients-editor[1]/section/div/div/div/div/div/span[@title='Delete this ingredient']"
Const deleteAllIngElement = "//form/div[2]/div[2]/div[1]/div[2]/div/div[1]/div[2]/button[@data-bind='click: $parent.deleteIngredientNutrients.bind($parent, $data)']"
Const addAllIngElement = "//form/div[2]/div[2]/div[1]/div[1]/div/button[@data-bind='click: $component.addNewIngredientNutrients.bind($component)']"
Const confirmElement = "//button[@class='btn-confirm k-button']"
Const expandIngElement = "//form/div[2]/div[2]/div[1]/div[2]/div/div[2]/ul/li/input[@checked='checked']"
Const addInactive = "//ingredients-editor[2]/section/div/div[3]/div[2]/div[@class='map-select single']"
Const ingDisplayedElement = "//ingredients-editor[1]/section/div/div[3]/div[2]/div/div[1]/div[@class='select-placeholder']"
Const saveElement = "//div[1]/div/button[@data-bind='click: save, enable: canSave']"
Const undoElement = "//div[1]/div/button[@data-bind='click: undo, enable: canUndo']"
Const dropdownElement = "//ingredients-editor[2]/section/div/div[3]/div[2]/div/div[2]/div[2]/div"
Const inactiveIngElement = "//ingredients-editor[2]/section/div/div/div/div/div[@data-bind='text: IngredientName']"
Const activeIngElement = "//ingredients-editor[1]/section/div/div/div/div/div[@data-bind='text: IngredientName']"
Const searchboxElement = "//ingredients-editor[2]/section/div/div[3]/div[2]/div/div[2]/div[1]"
Const packCountElement = "//form/div[2]/div[2]/div[1]/div[2]/div[@class='well']"
Const packNameElement = "//input[@placeholder='Enter Product Name...']"
const errorElement = "//div[2]/div[1]/ul/li[@data-bind='text: message']"

Sub Main()
    Dim MatchSheet, MainSheet, DataSheet, ActiveIngredientSheet, SKUSheet As Worksheet
    Set MatchSheet = Sheet41 '  "Items Updated"
    Set MainSheet = Sheet1 ' "SKU Entry"
    Set DataSheet = Sheet11 'Data
    Set ActiveIngredientSheet = Sheet6 ' "Update Active Ingredients"
    Set SKUSheet = Worksheets("SKU's checked")

    If MainSheet.Range("A2") = "" Then
        MsgBox "You can't update ingredients if there aren't any SKU's.  Doh!"
        Exit Sub
    ElseIf MainSheet.Range("C2") = "" And MainSheet.Range("D2") = "" Then
        MsgBox "You did not enter any items to separate ingredients.  Please try again."
        Exit Sub
    ElseIf MsgBox("Are you sure you want to update these items in MAP?  Excel will be unavailable while this is running.", vbOKCancel) <> vbOK Then
        Exit Sub
    End If

    'Get cell locations used to record important stats
    Dim dateRow as Long
    Dim itemsChecked, itemsMatched, itemsUpdated, ingredientsAnalyzed, ingredientsMatched, matchedRunTime, updatedRunTime as Range
    With DataSheet
      dateRow = WorksheetFunction.Match(.Range("A1:A" & lastrow(DataSheet)))
      Set itemsChecked = .Range("B" & dateRow)
      Set itemsMatched = .Range("C" & dateRow)
      Set itemsUpdated = .Range("D" & dateRow)
      Set ingredientsAnalyzed = .Range("E" & dateRow)
      Set ingredientsMatched = .Range("F" & dateRow)
      Set matchedRunTime = .Range("H" & dateRow)
      Set updatedRunTime = .Range("I" & dateRow)
    End With

    'if some items have been updated, double-check the most recent 3.  Otherwise, start from the top.
    Dim checkpoint As Long
    If MatchSheet.Cells(Rows.Count, 2).End(xlUp).Row <= 3 Then
        checkpoint = 2
    Else
        checkpoint = MatchSheet.Cells(Rows.Count, 2).End(xlUp).Row - 1
    End If

    Set S = New Selenium.ChromeDriver
    Application.GoTo Range("A" & checkpoint), True

    Dim searchRange As Range, numSearchItems As Long, searchString As String
    Set searchRange = MainSheet.Range("A" & checkpoint & ":A" & lastRow(MainSheet))
    numSearchItems = Application.WorksheetFunction.CountA(searchRange)
    searchString = ""

    'Create search string of all SKU's
    For Each sku In searchRange
        searchString = searchString & " " & Trim(sku)
    Next sku

    S.get URL

    'input SKUs
    Set searchbox = S.FindElementsByClass("form-control")(1)
    searchbox = searchbox.ExecuteScript("arguments[0].value='" & searchString & "';$(arguments[0]).trigger('change');", searchbox)

    'click to search results
    S.FindElementByClass("js-button-search").Click

    'Ensure the search results return every SKU searched
    If InStr(S.FindElementByXPath(searchResultsElement).Text, numSearchItems & " items") = 0 Then
        MsgBox "ERROR: number of items returned is different than number of items searched.  Make sure each item can be found in MAP and try again."
        End
    End If

    'right click on first result item
    S.Actions.ClickContext(S.FindElementByXPath("//div[1]/div[4]/table/tbody/tr[1]/td[2]")).Perform

    'click to expand ingredient information
    S.Mouse.moveTo(S.FindElementByXPath("//li[@data-name='EditItemDetails']")).Click

    For Each sku In searchRange
        Dim IngredientCollection As New Collection, packNameCollection As New Collection, separatedString As New Collection
        Set IngredientCollection = New Collection
        Set packNameCollection = New Collection
        totalIngredientCount = 0
        y = 1
        multipack = False

        'check for multiple sets of ingredients for the item
        For x = 5 To 17 Step 3
            If MainSheet.Cells(sku.Row, x) <> "" Or MainSheet.Cells(sku.Row, x + 1) <> "" Or MainSheet.Cells(sku.Row, x + 2) <> "" Then
                multipack = True
                y = y + 1
            End If
        Next x

        'add each set of ingredients as a unique collection
        For x = 1 To y
            'clean up ingredient string
            inactiveIngredientString = Trim(MainSheet.Cells(sku.Row, x * 3 + 1))
            inactiveIngredientString = cleanString(inactiveIngredientString)

            'convert cleaned string to collection of ingredients
            Set separatedString = New Collection
            Set separatedString = stringToCollection(inactiveIngredientString)
            totalIngredientCount = totalIngredientCount + separatedString.Count
            IngredientCollection.Add separatedString

            'save each pack name if they exist
            If multipack Then packNameCollection.Add Trim(MainSheet.Cells(sku.Row, x * 3 - 1))
        Next x

        ingredientsAnalyzed = ingredientsAnalyzed + totalIngredientCount
        packCount = IngredientCollection.Count
        If packCount < 1 Then packCount = 1

        StartTime = Now()

        'expand all ingredients to check
        For x = 1 To S.FindElementsByXPath(expandIngElement).Count
            S.FindElementsByXPath(expandIngElement)(x).Click
        Next x

        'count the # of inactive ingredients to compare to ingredients for item in excel
        MAPinactiveCount = S.FindElementsByXPath(deleteInactiveIngElement).Count
        MAPactiveCount = S.FindElementsByXPath(deleteActiveIngElement).Count
        If MAPactiveCount = 0 And MAPinactiveCount = 0 Then GoTo AddIngredients

        Dim enteredWrong as Boolean
        If MAPinactiveCount <> totalIngredientCount Or packCount > S.FindElementsByXPath(expandIngElement).Count Then _
            enteredWrong = True

        'check if active ingredient is the same as the inactive ingredient
        Dim enteredActiveWrong as Boolean
        For x = 1 To IngredientCollection.Count
            For y = 1 To IngredientCollection(x).Count
                ingredient = IngredientCollection(x)(y)
                For z = 1 To MAPactiveCount
                    MAPingredient = S.FindElementsByXPath(activeIngElement)(z).Text
                    If LCase(MAPingredient) = LCase(ingredient) Then
                        enteredActiveWrong = True
                    End If
                Next z
            Next y
        Next x

        If enteredWrong or enteredActiveWrong Then GoTo DeleteIngredients

        'check if inactive ingredients match ingredients in Excel doc
        i = 1
        For x = 1 To IngredientCollection.Count
            For y = 1 To IngredientCollection(x).Count
                ingredient = IngredientCollection(x)(y)
                MAPingredient = S.FindElementsByXPath(inactiveIngElement)(i).Text
                If InStr(LCase(MAPingredient), LCase(ingredient)) = 0 And InStr(LCase(ingredient), LCase(MAPingredient)) = 0 Then
                    enteredWrong = True
                    GoTo DeleteIngredients
                End If
                ingredientsMatched = ingredientsMatched + 1
                i = i + 1
            Next y
        Next x

        'All ingredients matched - confirm and move to next SKU
        itemsMatched = itemsMatched + 1
        matchedRunTime = matchedRunTime + Now() - StartTime
        MatchSheet.Range("B" & sku.Row) = "no change needed"
        GoTo NextSKU

DeleteIngredients:
        If MAPactiveCount = 0 Or enteredActiveWrong Then
            'delete all ingredient packs
            deleteCount = S.FindElementsByXPath(deleteAllIngElement).Count
            For x = 1 To deleteCount
                S.FindElementByXPath(deleteAllIngElement).Click
                S.FindElementByXPath(confirmElement).Click
            Next x
            S.FindElementByXPath(saveElement).Click
        Else
            'only delete inactive ingredients
            For x = 1 To MAPinactiveCount
                S.FindElementsByXPath(deleteInactiveIngElement)(1).Click
            Next x
        End If

AddIngredients:
        For x = 1 To packCount
            For y = 1 To IngredientCollection(x).Count
                ingredient = IngredientCollection(x)(y)

                'Add ingredient set if we need to
                If x > S.FindElementsByXPath(ingDisplayedElement).Count Then
                    S.FindElementByXPath(addAllIngElement).Click
                    S.FindElementsByXPath(expandIngElement)(x).Click
                    S.FindElementsByXPath(addInactive)(x).Click
                End if
                'click to add inactive ingredient
                Set searchbox = S.FindElementsByXPath(searchboxElement & "/input")(x)
                searchbox = searchbox.ExecuteScript("arguments[0].value='" & ingredient & "';$(arguments[0]).trigger('change');", searchbox)

                'Click on input box again to active dropdown listener
                S.FindElementsByXPath(searchboxElement)(x).Click

                'Count dropdown options
                siblingCount = S.FindElementsByXPath(dropdownElement).Count

                'Evaluate first 3 available dropdown options
                For i = 2 To 4
                    If i > siblingCount Then Exit For
                    If ingredient = S.FindElementsByXPath(dropdownElement)(i).Text Then
                        'Click on ingredient that matches the ingredient entered
                        S.FindElementsByXPath(dropdownElement)(i).Click
                        GoTo addNextIngredient
                    End If
                Next i

                'if there are no dropdown ingredients, or no dropdown value the same as the ingredient, add it as a new ingredient
                If siblingCount > 1 Then
                    S.FindElementByXPath("//div[@class='option add']").Click
                Else
                    S.FindElementByXPath("//div[@class='option add selected']").Click
                End If

                'confirm new ingredient addition
                S.FindElementByXPath(confirmElement).Click
addNextIngredient:
          Next y
      Next x

      'save changes after last ingredient is entered
        S.FindElementByXPath(saveElement).Click

        'Check if there was an issue saving the ingredient update
        If S.FindElementsByXPath(errorElement).Count > 0 Then
            Dim errorMessage as String
            errorMessage = S.FindElementByXPath(errorElement).Text
            If MAPactiveCount = 0 And Not secondAttempt Then
                secondAttempt = True

                'Delete entire ingredient list
                S.FindElementByXPath(deleteAllIngElement).Click
                S.FindElementByXPath(confirmElement).Click

                'Add new ingredient list
                S.FindElementByXPath(addAllIngElement).Click
                S.FindElementByXPath(saveElement).Click

                'Attempt to re-add ingredients
                GoTo AddIngredients

            ElseIf InStr(errorMessage, "An item with the same key has already been added.") > 0 Then
                MatchSheet.Range("B" & sku.Row) = "ERROR: Duplicate ingredients detected"
            ElseIf InStr(errorMessage, "Invalid Amount Per entry:") Then
                MatchSheet.Range("B" & sku.Row) = "ERROR: Nutrients entered incorrectly"
            ElseIf InStr(errorMessage, "is not a recognized unit measurement") Then
                MatchSheet.Range("B" & sku.Row) = "ERROR: Nutrient units entered incorrectly"
            End If

            'Undo changes
            S.FindElementByXPath(undoElement).Click
            S.FindElementByXPath(confirmElement).Click
        End If

NextSKU:
    'Click to the next SKU
    S.FindElementByClass("btn-next").Click
    Next sku

    'Close browser
    Set S = Nothing

    SKUSheet.Range("$A$1:$B$" & lastRow(SKUSheet)).RemoveDuplicates Columns:=1, Header:=xlYes

    i = 0
    For Each cell In MatchSheet.Range("B2:B" & lastRow(MatchSheet))
        If cell = "item updated" Or cell = "no change needed" Then i = i + 1
    Next cell

    MsgBox "Ingredient information up to date for " & i & " items!"

End Sub


Function lastRow(wksht As Worksheet) As Long
    lastRow = wksht.Cells(Rows.Count, 1).End(xlUp).Row + 1
End Function
