'Send emails to suppliers that need to update item information
Sub EmailLoop()

    If MsgBox("Are you sure you want to email these vendors?  This cannot be undone.", vbYesNoCancel) <> vbYes Then Exit Sub
    Call General("start")

    Dim combinedDigitalCMs, emailCount, cols, subjects, templates As Variant
    combinedDigitalCMs = Array("Maholm", "Sharp", "Mintzias")
    emailCount = Array(0, 0, 0, 0)
    cols = Array("A", "B", "C", "D")
    subjects = Array("CEP ACTION REQUIRED: ", "CEP ACTION NEEDED: Submit Assets", "CEP ACTION NEEDED: ", "CEP ACTION NEEDED: ")
    templates = Array("Correction Required", "In Progress", "Not Submitted")

    Dim HTML, ResetSheet, RefSheet, DataSheet, VendorSheet As Worksheet
    Set HTML = Worksheets("HTML")
    Set ResetSheet = Worksheets("Reset Names")
    Set RefSheet = Worksheets("Ref")
    Set DataSheet = Worksheets("Data")
    Set VendorSheet = Worksheets("Vendors")

    If (VendorSheet.AutoFilterMode And VendorSheet.FilterMode) Or VendorSheet.FilterMode Then VendorSheet.ShowAllData
    If (DataSheet.AutoFilterMode And DataSheet.FilterMode) Or DataSheet.FilterMode Then DataSheet.ShowAllData
    VendorSheet.Columns("F:G").Hidden = False
    DataSheet.Columns("E:E").Hidden = True

    Dim vendor, vendors, items As Range, bottom As Long
    Set vendors = VendorSheet.Range("C2:C" & lastRow(VendorSheet))
    Set items = DataSheet.Range("A1:N" & lastRow(DataSheet))
    bottom = lastRow(ResetSheet)

    'Find the file attachments
    Dim MyObj, MySource as Object, file as Variant
    Dim file1, file2, file3 As String
    Set MySource = MyObj.GetFolder("I:\EcomNEW\DigitalCommerce\Merchandising Operations\Digital Reset Scorecard\CEP Supplier Outreach\Attachments\")
    For Each file In MySource.Files
        If InStr(file.name, "Image Guide") Then
            file1 = file.name
        ElseIf InStr(file.name, "CEP") Then
            file2 = file.name
        ElseIf InStr(file.name, "Training Schedule") Then
            file3 = file.name
        End If
    Next file

    For Each vendor In VendorSheet.Range("C2:C" & lastRow(VendorSheet))
        Dim resetName, dearName As String, alreadyContacted As Boolean

        'Skip to next vendor if they have 0 items or no email documented
        If VendorSheet.Range("G" & vendor.Row) = 0 Or VendorSheet.Range("E") & vendor.Row = "" Then GoTo nextVendor

        resetName = VendorSheet.Range("B" & vendor.Row)
        dearName = VendorSheet.Range("D" & vendor.Row)
        If dearName = "" Then dearName = Application.WorksheetFunction.Proper(vendor)
        If VendorSheet.Range("F" & vendor.Row) <> "" Then alreadyContacted = True

        For Each template In templates
            Dim vendorItems As Range, resetRow, index, itemCount As Integer
            Dim emailSubject, emailCC, emailBody, col As String

            'Skip to the next email template if the vendor has 0 items with the current template
            With DataSheet
                itemCount = Application.WorksheetFunction.CountIfs(.Range("A:A"), resetName, .Range("C:C"), template, .Range("N:N"), vendor)
            End With
            If itemCount = 0 Then GoTo nextTemplate

            With items
                .AutoFilter Field:=3, Criteria1:=template
                .AutoFilter Field:=1, Criteria1:=resetName
                .AutoFilter Field:=14, Criteria1:=vendor
            End With

            'Decide which email template to use
            If template = templates(0) Then
                index = 0
            ElseIf template = templates(1) Then
                index = 1
            Else
                If Not alreadyContacted Then
                    index = 2
                    VendorSheet.Range("F" & vendor.Row) = Now()
                Else
                    index = 3
                End If
            End If

            Set vendorItems = DataSheet.Range("D1:G" & lastRow(DataSheet)).SpecialCells(xlCellTypeVisible)
            resetRow = Application.WorksheetFunction.Match(resetName, ResetSheet.Range("B2:B" & bottom))

            'Create subject, CC, body for email
            If index <> 1 Then
                emailSubject = subjects(index) & vendor
            Else
                emailSubject = subjects(index)
            End If
            If resetRow > 0 Then
                emailCC = ResetSheet.Range("E" & resetRow)
            Else
                emailCC = ""
            End If
            col = cols(index)
            With HTML
                emailBody = .Range(col & "1") & dearName & .Range(col & "3") & ConvertRangetoHTML(vendor_items) & .Range(col & "5") & Application.UserName & .Range(col & "7")
            End With

            'Create Outlook instance and send email
            Set OutApp = CreateObject("Outlook.Application")
            Set outMail = OutApp.createitem(olMailItem)
            With outMail
                .SentOnBehalfOfName = "CEPSupport@walgreens.com"
                .To = VendorSheet.Range("E" & vendor.Row)
                .CC = emailCC
                .subject = emailSubject
                .HTMLBody = emailBody
                If template = "Not Submitted" Then
                    .Attachments.Add FolderLocation & file1
                    .Attachments.Add FolderLocation & file2
                    .Attachments.Add FolderLocation & file3
                End If
                .Send
            End With
            emailCount(index) = emailCount(index) + 1
nextTemplate:
        Next
nextVendor:
    Next vendor

    VendorSheet.Columns("F:G").Hidden = True
    DataSheet.Columns("E:E").Hidden = False
    If (DataSheet.AutoFilterMode And DataSheet.FilterMode) Or DataSheet.FilterMode Then DataSheet.ShowAllData

    'Display how many emails were sent with each template
    ReDim Preserve templates(3)
    templates(3) = "Not Started (follow-up)"
    message = "Emails Sent:"
    For i = 0 To 3
        message = message & vbLf & templates(i) & "  -  " & emailCount(i)
    Next i
    MsgBox message
    Call General("end")
End Sub


'Speed up code runtime and reduce lines of code in main sub
Sub General(startOrEnd As Variant)
    If startOrEnd = "start" Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculate
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub


'Find the last used cell row in a worksheet
Function lastRow(wksht As Worksheet) As Long
    lastRow = wksht.Cells(Rows.Count, 1).End(xlUp).Row
End Function


'Convert vendor items into the appropriate HTML code for email body
Function ConvertRangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
            SourceType:=xlSourceRange, _
            Filename:=TempFile, _
            Sheet:=TempWB.Sheets(1).Name, _
            Source:=TempWB.Sheets(1).UsedRange.Address, _
            HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into ConvertRangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    ConvertRangetoHTML = ts.readall
    ts.Close
    ConvertRangetoHTML = Replace(ConvertRangetoHTML, "align=center x:publishsource=", _
                            "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
