Public Sub generateDoc(Template, inKey, fileName As String)
    Dim WordApp, WordDoc As Object
    Dim ws As Worksheet
    Dim columnTitle As String
    Dim lastCol, currentCol, selectedRow As Long
    Dim cellValue As Variant
    Dim foundCell As Range
    
    'Set the worksheet and specified row
    Set ws = ThisWorkbook.Sheets("Data")
    Set foundCell = ws.Cells.Find(What:=inKey, LookIn:=xlValues, LookAt:=xlWhole)
    selectedRow = foundCell.Row

    'Open Word template
    Set WordApp = Nothing
    On Error Resume Next
    Set WordApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("Word.Application")
    End If
    
    WordApp.Visible = True
    Set WordDoc = WordApp.Documents.Open(Template)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    'Insert current date
    With WordApp.Selection.Find
        .Text = "<<TODAY>>"
        .Replacement.Text = Format(Now(), "dd MMMM yyyy")
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    'Insert username
    userN = Replace(Environ("Username"), ".", " ") 'customize
    With WordApp.Selection.Find
        .Text = "<<USERNAME>>"
        .Replacement.Text = StrConv(userN, vbProperCase)
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    ' Loop through each column
    For currentCol = 1 To lastCol
        columnTitle = ws.Cells(1, currentCol).Value
        columnTitle = Replace(columnTitle, "?", "") 'remove wildcard chars
        cellValue = ws.Cells(selectedRow, currentCol).Value
        
        'format datetime values
        Select Case columnTitle
            Case "Created", "Closed" 'add datetime col names here
                cellValue = Format(cellValue, "d/MM/yyyy")
        End Select
        
        'replace fields with long strings
        With WordDoc.Range
            .Find.Text = "<<" & columnTitle & ">>"
            Do While .Find.Execute
                .Text = cellValue
                .Collapse wdCollapseEnd
            Loop
        End With
    Next currentCol
    
    'Save the document
    WordDoc.SaveAs2 (Application.ActiveWorkbook.Path & "\" & fileName)
    WordDoc.Close
    WordApp.Quit
    
    'cleanup
    ThisWorkbook.Sheets("Notify").Range("C4").Value = "" 'clear input
    ThisWorkbook.Sheets("Notify").Range("C7").Value = "" 'clear input
    ThisWorkbook.Sheets("Notify").Range("C4").Select 'reset selection
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set foundCell = Nothing

End Sub

Public Sub GenButton()
    Dim fileName As String
    Dim Template As String
    
    inKey = ThisWorkbook.Sheets("Notify").Range("C4").Value
    notifyType = ThisWorkbook.Sheets("Notify").Range("C7").Value
    templatePath = ThisWorkbook.Path & "\Templates\"
    
    'missing inputs
    If inKey = "" Or notifyType = "" Then
        MsgBox "Please select a Key and Type"
        End
    End If
    
    'template trigger
    Select Case notifyType
        Case Is = "Template1"
            Template = templatePath & ThisWorkbook.Sheets("Notify").Range("C13").Value
            fileName = inKey & " - SampleDocument1.docx"
            generateDoc Template, inKey, fileName 'generate document
        Case Is = "Template2"
            Template = templatePath & ThisWorkbook.Sheets("Notify").Range("C14").Value
            fileName = inKey & " - SampleDocument2.docx"
            generateDoc Template, inKey, fileName 'generate document
    End Select
    Application.CutCopyMode = False
End Sub
