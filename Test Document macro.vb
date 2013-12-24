Option Explicit

Sub Something()
    Dim myListString$
    Dim myHeading$
    Dim myParagraph$
    Dim DocPara As Paragraph
    Dim tableColumns As Integer
    Dim tableRows As Integer
    Dim rowCellCount As Integer
    Dim myListStrings() As String
    Dim myListStringsCount As Integer
    Dim doingTable As Boolean
    Dim tableHeadingCount As Integer
    Dim tempHeightOfEquation As Integer
    Dim tempWidthOfEquation As Integer
    
    Dim oExcel As Excel.Application
    Dim oWB As Workbook
    Set oExcel = New Excel.Application
    Set oWB = oExcel.Workbooks.Add
    
    On Error GoTo errH
    
    Dim errorOnSaveAs As Boolean
    errorOnSaveAs = True
    
    oWB.SaveAs FileName:="C:\Users\J.Smith\Desktop\Test.xlsx"
    
    errorOnSaveAs = False
    
    oExcel.Visible = True
    
    Dim paragraphCount As Integer
    Dim headingCount As Integer
    
    With oWB.ActiveSheet
        .Cells(1, 1) = "No."
        .Cells(1, 2) = "Title"
        .Cells(1, 3) = "Paragraph"
        .Rows(1).Font.Bold = True
    End With
    
    headingCount = headingCount + 1
    rowCellCount = 1 'Initialize rowCellCount
    myListStringsCount = 0 'Initialize number of non-main outline characters
    tableHeadingCount = 0 'Initialize location of tables
    
    ' ---------------------------------------------
    '| Loop through the paragraphs of the document |
    ' ---------------------------------------------
    For Each DocPara In ActiveDocument.Paragraphs
        paragraphCount = paragraphCount + 1
        myListString$ = DocPara.Range.ListFormat.ListString
        
        ' ------------------------------------
        '| The paragraph is not in an outline |
        ' ------------------------------------
        If myListString$ = "" Then 'The paragraph is not in an outline
            ' ------------------------------
            '| This paragraph is a picture! |
            ' ------------------------------
            If DocPara.Range.InlineShapes.Count > 0 Then 'This paragraph is a picture!
                doingTable = False
                
                If Not IsEmpty(oWB.ActiveSheet.Cells(headingCount, 3)) Then 'We have already been here!
                    headingCount = headingCount + 1 'Go to the next row
                End If
                
                myParagraph$ = DocPara.Range.InlineShapes(1).Range
                
                oWB.ActiveSheet.Rows(CStr(headingCount)).RowHeight = DocPara.Range.InlineShapes(1).Height 'Make this row big enough for the picture
'                oWB.ActiveSheet.Columns("C").ColumnWidth = DocPara.Range.InlineShapes(1).Width 'Make the column big enough for the picture
                                
                'Insert the picture
                With oWB.ActiveSheet.Pictures.Insert(DocPara.Range.InlineShapes(1).AlternativeText)
                    With .ShapeRange
                        .LockAspectRatio = True
                        .Width = DocPara.Range.InlineShapes(1).Width
                        .Height = DocPara.Range.InlineShapes(1).Height
                    End With
                    .Left = oWB.ActiveSheet.Cells(headingCount, 3).Left
                    .Top = oWB.ActiveSheet.Cells(headingCount, 3).Top
                    .Placement = 1
                    .PrintObject = True
                End With
                
                rowCellCount = 1 'Re-initialize rowCellCount
                oWB.ActiveSheet.Cells(headingCount, 3) = myParagraph$ 'Import to Excel
            ' -------------------------------
            '| This paragraph is an Equation |
            ' -------------------------------
            ElseIf DocPara.Range.OMaths.Count > 0 Then 'This paragraph is an Equation
                doingTable = False
                
                If Not IsEmpty(oWB.ActiveSheet.Cells(headingCount, 3)) Then 'We have already been here!
                    headingCount = headingCount + 1 'Go to the next row
                End If
                
                myParagraph$ = DocPara.Range.OMaths(1).Range
                DocPara.Range.OMaths(1).Range.Select
                With Selection
                    .CopyAsPicture
                    With oWB.ActiveSheet
                        .Paste Destination:=.Cells(headingCount, 3)
                        tempHeightOfEquation = .Shapes(.Shapes.Count).Height
                        tempWidthOfEquation = .Shapes(.Shapes.Count).Width
                        .Shapes(.Shapes.Count).Delete
                        .Rows(CStr(headingCount)).RowHeight = tempHeightOfEquation
                        .Paste Destination:=.Cells(headingCount, 3)
                        .Shapes(.Shapes.Count).Width = tempWidthOfEquation
                    End With
                End With
                
                rowCellCount = 1 'Re-initialize rowCellCount
            ' -----------------------------------
            '| This paragraph is part of a table |
            ' -----------------------------------
            ElseIf DocPara.Range.Tables.Count > 0 Then 'This paragraph is part of a table
                tableColumns = DocPara.Range.Tables(1).Columns.Count 'Get number of columns of table
                tableRows = DocPara.Range.Tables(1).Rows.Count 'Get number of rows of table
                
                If Not doingTable Then 'This is the first table cell
                    tableHeadingCount = tableHeadingCount + 2
                    With oWB.ActiveSheet.Cells(headingCount, 3)
                        .FormulaR1C1 = "=HYPERLINK(""" & "[" & oWB.Name & "]Sheet2!A" & tableHeadingCount & """, ""Click for table on Sheet2"")"
                        .WrapText = True
                    End With
                    doingTable = True
                End If
                
                If rowCellCount > tableColumns + 1 Then 'We need to move to the next row
                    tableHeadingCount = tableHeadingCount + 1 'Go to the next row on Sheet2
                    rowCellCount = 1 'Re-initialize rowCellCount
                End If
                
                myParagraph$ = Left(DocPara.Range.Text, Len(DocPara.Range.Text) - 1) 'Save table cell contents
                
                Dim isAtEndOfTableRow As Boolean
                isAtEndOfTableRow = oWB.Sheets(2).Cells(tableHeadingCount, 1 + rowCellCount - 1).Next.column < tableColumns + 2
                
                If isAtEndOfTableRow Then 'We have the end of the row indicator on our hands
                    oWB.Sheets(2).Cells(tableHeadingCount, 1 + rowCellCount - 1) = myParagraph$ 'Import to Excel on Sheet2
                    With oWB.Sheets(2).Cells(tableHeadingCount, 1 + rowCellCount - 1).Borders
                        .LineStyle = xlContinuous
                        .Color = vbBlack
                        .Weight = xlThin
                    End With
                End If
                rowCellCount = rowCellCount + 1 'Move to next cell in row
                
                'TODO: This is where you would make the table look pretty
            ' --------------------------------------------------
            '| This paragraph is the body of an outline section |
            ' --------------------------------------------------
            Else 'This paragraph is the body of an outline section
                doingTable = False
                If Not IsEmpty(oWB.ActiveSheet.Cells(headingCount, 3)) Then 'We have already been here!
                    headingCount = headingCount + 1 'Go to the next row
                End If
                
                myParagraph$ = DocPara.Range.Text 'Save body of outline section
                rowCellCount = 1 'Re-initialize rowCellCount
                oWB.ActiveSheet.Cells(headingCount, 3) = myParagraph$ 'Import to Excel
                headingCount = headingCount + 1 'go to the next row
            End If
        ' ----------------------------------------
        '| This paragraph is in an outline header |
        ' ----------------------------------------
        Else 'This paragraph is in an outline header
            headingCount = headingCount + 1 'Move to the next row
            myHeading$ = DocPara.Range.Text 'Get the outline section header
            ' ------------------------------------------
            '| This paragraph is in the overall outline |
            ' ------------------------------------------
            If IsNumeric(Left(myListString$, 1)) Then 'This paragraph is in the overall outline
                oWB.ActiveSheet.Cells(headingCount, 1) = myListString$ 'Import to Excel
                oWB.ActiveSheet.Cells(headingCount, 2) = myHeading$ 'Import to Excel
            ' -----------------------------------------------
            '| This paragraph is in a section body's outline |
            ' -----------------------------------------------
            Else 'This paragraph is in a section body's outline
                Dim column As Integer
                column = 3 + DocPara.OutlineLevel - 10 'Initialize column to place outline item
                
                If myListStringsCount = 0 Then 'Initialize myListStrings array
                    ReDim myListStrings(0 To 0) As String
                    myListStrings(UBound(myListStrings)) = myListString$
                    myListStringsCount = 1
                ElseIf IsInArray(myListString$, myListStrings) Then 'myListString$ is not a new outline character
                    Dim i As Integer
                    For i = LBound(myListStrings) To UBound(myListStrings)
                        If myListString$ = myListStrings(i) Then
                            Exit For
                        End If
                    Next i
                    column = column + i
                Else 'myListString$ is a new outline character
                    ReDim Preserve myListStrings(0 To UBound(myListStrings) + 1) As String
                    myListStrings(UBound(myListStrings)) = myListString$
                    myListStringsCount = myListStringsCount + 1
                    column = column + UBound(myListStrings)
                End If
                oWB.ActiveSheet.Cells(headingCount, column) = myListString$ & myHeading$  'Import to Excel
            End If
        End If
    Next DocPara
    
    With oWB.ActiveSheet
        .Columns("A:B").AutoFit
        .Columns("A:C").HorizontalAlignment = xlHAlignLeft
    End With
errH:
    If errorOnSaveAs Then
        MsgBox "Saving file was likely cancelled:" & vbCr & vbCr & Err.Description
    End If
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
