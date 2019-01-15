Sub do_2()

Dim wrdApp As Word.Application
Dim wrdDoc As Word.Document
Dim i As Integer
Dim ws As Excel.Worksheet
Dim table_inv_obj As Object
Dim oCell As Object
Dim pg As Paragraph
Dim sh As Shape

Set ws = ThisWorkbook.Sheets("Foglio1")

On Error Resume Next
    Set wrdApp = GetObject(, "Word.Application")
If Err Then
    Set wrdApp = CreateObject("Word.Application")
End If
On Error GoTo 0

wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add(Template:="Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\paper1.dotx", Newtemplate:=False, DocumentType:=0)
'Set wrdRange = wrdDoc.Range


With wrdDoc
'    .Shapes(1).TextFrame.TextRange.Text = "Fondo Feri PIR"

    Set wrdRange = .Paragraphs.Last.Range
    
'===========================
'create table1
'===========================
    .Tables.Add Range:=wrdRange, NumRows:=13, NumColumns:=2
    .Tables(1).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(1).Rows.HorizontalPosition = Application.CentimetersToPoints(0)
    .Tables(1).Rows.VerticalPosition = Application.CentimetersToPoints(3.6)
    .Tables(1).Columns(1).Width = Application.CentimetersToPoints(2.4)
    .Tables(1).Columns(2).Width = Application.CentimetersToPoints(2.5)
    .Tables(1).Rows.HeightRule = 2
    .Tables(1).Rows.Height = Application.CentimetersToPoints(0.4)
    
'    .Tables(1).Borders.InsideLineStyle = wdLineStyleDashDot
    With .Tables(1).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
    wrdRange.Collapse Direction:=wdCollapseEnd
    With wrdRange
        .Collapse Direction:=wdCollapseEnd
        .InsertParagraphAfter
        .Collapse Direction:=wdCollapseEnd
    End With
    
'===========================
'create table2
'===========================
    .Tables.Add Range:=wrdRange, NumRows:=12, NumColumns:=2
    .Tables(2).Rows.HorizontalPosition = Application.CentimetersToPoints(0)
    .Tables(2).Rows.VerticalPosition = Application.CentimetersToPoints(6.2)
    .Tables(2).Columns(1).Width = Application.CentimetersToPoints(2.4)
    .Tables(2).Columns(2).Width = Application.CentimetersToPoints(2.5)
    .Tables(2).Rows.HeightRule = 2
    .Tables(2).Rows.Height = Application.CentimetersToPoints(0.4)
'    .Tables(1).Borders.InsideLineStyle = wdLineStyleDashDot
    With .Tables(2).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
    Set wrdRange = .Paragraphs.Last.Range
'===========================
'create table3
'===========================
    .Tables.Add Range:=wrdRange, NumRows:=10, NumColumns:=2
    .Tables(3).Rows.HorizontalPosition = Application.CentimetersToPoints(0)
    .Tables(3).Rows.VerticalPosition = Application.CentimetersToPoints(11.8)
    .Tables(3).Columns(1).Width = Application.CentimetersToPoints(3.4)
    .Tables(3).Columns(2).Width = Application.CentimetersToPoints(1.5)
    .Tables(3).Rows.HeightRule = 2
    .Tables(3).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(3).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'===========================
'create table4
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=2, NumColumns:=1
    .Tables(4).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(4).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(4).Rows.VerticalPosition = Application.CentimetersToPoints(3.6)
    .Tables(4).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(4).Rows.HeightRule = 1
    .Tables(4).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(4).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'===========================
'create table5
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(5).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(5).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(5).Rows.VerticalPosition = Application.CentimetersToPoints(6.6)
    .Tables(5).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(5).Rows.HeightRule = 2
    .Tables(5).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(5).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'===========================
'create table6
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=7
    .Tables(6).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(6).Rows.HorizontalPosition = Application.CentimetersToPoints(8)
    .Tables(6).Rows.VerticalPosition = Application.CentimetersToPoints(7.3)
    .Tables(6).Columns.Width = Application.CentimetersToPoints(1)
    .Tables(6).Rows.HeightRule = 1
    .Tables(6).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(6)
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.InsideLineWidth = wdLineWidth150pt
        .Borders.InsideColor = wdColorDarkBlue
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineWidth = wdLineWidth150pt
        .Borders(wdBorderTop).Color = wdColorDarkBlue
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
        .Borders(wdBorderBottom).Color = wdColorDarkBlue
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).LineWidth = wdLineWidth150pt
        .Borders(wdBorderLeft).Color = wdColorDarkBlue
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineWidth = wdLineWidth150pt
        .Borders(wdBorderRight).Color = wdColorDarkBlue
    End With
    
'===========================
'create table7
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=3
    .Tables(7).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(7).Rows.HorizontalPosition = Application.CentimetersToPoints(8)
    .Tables(7).Rows.VerticalPosition = Application.CentimetersToPoints(7.6)
    .Tables(7).Columns(1).Width = Application.CentimetersToPoints(2.2)
    .Tables(7).Columns(2).Width = Application.CentimetersToPoints(2.6)
    .Tables(7).Columns(3).Width = Application.CentimetersToPoints(2.2)
    .Tables(7).Rows.HeightRule = 1
    .Tables(7).Rows.Height = Application.CentimetersToPoints(0.4)
    
    With .Tables(7).Cell(1, 1).Range
        .Text = "Lower risk" & Chr(10) & "Lower return"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 8
        .Font.ColorIndex = wdDarkBlue
    End With
    With .Tables(7).Cell(1, 3).Range
        .Text = "Higher risk" & Chr(10) & "Higher return"
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        .Font.Size = 8
        .Font.ColorIndex = wdDarkBlue
    End With
    
'===========================
'create table8
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    Debug.Print wrdRange
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(8).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(8).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(8).Rows.VerticalPosition = Application.CentimetersToPoints(8.5)
    .Tables(8).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(8).Rows.HeightRule = 1
    .Tables(8).Rows.Height = Application.CentimetersToPoints(0.4)
    With .Tables(8).Cell(1, 1).Range
        .Text = "Featured on the Key Information Document (KID), " _
            & "the SRRI is a measure of the overall risk and reward profile of a fund. The SRRI is derived from the volatility of past returns over a 5-year period. The lowest category does not mean risk free."
        .ParagraphFormat.Alignment = wdAlignParagraphJustifyMed
        .Font.Size = 6
        .Font.ColorIndex = wdDarkBlue
    End With
    
'===========================
'create table9
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(9).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(9).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(9).Rows.VerticalPosition = Application.CentimetersToPoints(9.4)
    .Tables(9).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(9).Rows.HeightRule = 2
    .Tables(9).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(9).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'===========================
'create table10
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    Debug.Print wrdRange
    .Tables.Add Range:=wrdRange, NumRows:=3, NumColumns:=9
    .Tables(10).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(10).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(10).Rows.VerticalPosition = Application.CentimetersToPoints(19.5)
'    .Tables(10).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(10).Rows.HeightRule = 2
    .Tables(10).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(10).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
    
'===========================
'create table11
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    Debug.Print wrdRange
    .Tables.Add Range:=wrdRange, NumRows:=4, NumColumns:=6
    .Tables(11).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(11).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(11).Rows.VerticalPosition = Application.CentimetersToPoints(21)
'    .Tables(11).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(11).Rows.HeightRule = 2
    .Tables(11).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(11).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
    
'===========================
'create table12
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    Debug.Print wrdRange
    .Tables.Add Range:=wrdRange, NumRows:=4, NumColumns:=13
    .Tables(12).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(12).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(12).Rows.VerticalPosition = Application.CentimetersToPoints(23)
'    .Tables(12).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(12).Rows.HeightRule = 2
    .Tables(12).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(12).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'===========================
'create table15
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(13).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(13).Rows.HorizontalPosition = Application.CentimetersToPoints(5.9)
    .Tables(13).Rows.VerticalPosition = Application.CentimetersToPoints(18.3)
    .Tables(13).Columns(1).Width = Application.CentimetersToPoints(13)
    .Tables(13).Rows.HeightRule = 1
    .Tables(13).Rows.Height = Application.CentimetersToPoints(0.4)

'======================================
'FILL TABLE 15
'======================================

    With .Tables(13).Cell(1, 1).Range
        .Text = "Fund performance displayed represents past performance which is not guarantee of future results. " _
             & "Investment returns and principal values may fluctuate so that an investor's shares, when redeemed," _
             & " may be worth more or less than their original cost."
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 6
        .Font.Bold = True
        .Font.ColorIndex = wdDarkBlue

    End With
 
   
'======================================
'FILL TABLE 1
'======================================
    
    For i = 1 To .Tables(1).Rows.Count
        For j = 1 To .Tables(1).Columns.Count
            .Tables(1).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            With .Tables(1).Cell(i, j).Range
                
                .Text = ws.Cells(i + 5, j + 2)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                If j = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphLeft
                ElseIf j = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
                
                .Font.ColorIndex = wdDarkBlue
            End With
        Next j
    Next i
    
'======================================
'FILL TABLE 2
'======================================


    For i = 1 To .Tables(2).Rows.Count
        For j = 1 To .Tables(2).Columns.Count
        .Tables(2).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            With .Tables(2).Cell(i, j).Range
                .Text = ws.Cells(i + 20, j + 2)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                   
                If j = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphLeft
                ElseIf j = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
                
                .Font.ColorIndex = wdDarkBlue
            End With
            
            If (i Mod 4) = 0 And i < 10 And j = 2 Then
                Set Rng = .Tables(2).Cell(i, j - 1).Range
                Rng.End = .Tables(2).Cell(i, j).Range.End
                Rng.Cells.Merge
                .Tables(2).Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            End If
        Next j
    Next i
    
'======================================
'FILL TABLE 3
'======================================
    
    For i = 1 To .Tables(3).Rows.Count
        For j = 1 To .Tables(3).Columns.Count
        .Tables(3).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            With .Tables(3).Cell(i, j).Range
                .Text = ws.Cells(i + 33, j + 2)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                If j = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphLeft
                ElseIf j = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
                
                .Font.ColorIndex = wdDarkBlue
            End With
        Next j
    Next i
    
    
'======================================
'FILL TABLE 4
'======================================
    .Tables(4).Cell(1, 1).Range.Font.Bold = True
    For i = 1 To .Tables(4).Rows.Count
        For j = 1 To .Tables(4).Columns.Count
        .Tables(4).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            With .Tables(4).Cell(i, j).Range
                .Text = ws.Cells(i + 1, j + 2)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                .ParagraphFormat.Alignment = wdAlignParagraphJustifyMed
                .Font.ColorIndex = wdDarkBlue
            End With
        Next j
    Next i
    
'======================================
'FILL TABLE 5
'======================================
    .Tables(5).Cell(1, 1).VerticalAlignment = wdCellAlignVerticalBottom
    With .Tables(5).Cell(1, 1).Range
        .Text = "Synthetic Risk & Reward Indicator (SRRI)"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 8
        .Font.Bold = True
        .Font.ColorIndex = wdDarkBlue

    End With
    
    
'======================================
'FILL TABLE 6
'======================================
    
        For j = 1 To .Tables(6).Columns.Count
        .Tables(6).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            With .Tables(6).Cell(1, j).Range
                .Text = j
                .ParagraphFormat.Alignment = wdAlignParagraphDistribute
                .Font.Size = 8
                .Font.Bold = True

                If j = 4 Then
                    .Font.ColorIndex = wdWhite
                    wrdDoc.Tables(6).Cell(1, j).Shading.BackgroundPatternColor = wdColorDarkBlue
                Else
                    .Font.ColorIndex = wdDarkBlue
                End If
                
            End With
        Next j
        
'======================================
'FILL TABLE 9
'======================================
    .Tables(9).Cell(1, 1).VerticalAlignment = wdCellAlignVerticalBottom
    With .Tables(9).Cell(1, 1).Range
        .Text = "Growth of 100 since Inception"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 8
        .Font.Bold = True
        .Font.ColorIndex = wdDarkBlue

    End With
    
'======================================
'put image1 and 2
'======================================
    l_par = .Paragraphs.Count
    .Paragraphs(l_par).Range.InlineShapes.AddPicture ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\perf_plt.jpg")
    .Content.InsertParagraphAfter
    
    .Paragraphs(l_par + 1).Range.InlineShapes.AddPicture ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\pie_asset_allocation.jpg")
    .Content.InsertParagraphAfter
    
    With ActiveDocument.InlineShapes(1)
        'border
        'conversion to Shape
        .ConvertToShape
    End With
    
    With ActiveDocument.Shapes(1)
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Top = Application.CentimetersToPoints(10.07)
        .Left = Application.CentimetersToPoints(7.44)
        .Width = Application.CentimetersToPoints(13)
        .Height = Application.CentimetersToPoints(8)
    End With
    
    With ActiveDocument.InlineShapes(1)
        'border
        'conversion to Shape
        .ConvertToShape
    End With
    
    With ActiveDocument.Shapes(2)
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Top = Application.CentimetersToPoints(20.44)
        .Left = Application.CentimetersToPoints(1.6)
        .Width = Application.CentimetersToPoints(5.01)
        .Height = Application.CentimetersToPoints(4.76)
    End With

  

'======================================
'FILL TABLE 10
'======================================
    
    For i = 1 To .Tables(10).Rows.Count
        For j = 1 To .Tables(10).Columns.Count
        .Tables(10).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            .Tables(10).Columns.Width = Application.CentimetersToPoints(1.44)
            With .Tables(10).Cell(i, j).Range
                .Text = ws.Cells(i + 5, j + 6)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                If i = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                ElseIf i = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Bold = True
                    .Font.Size = 8
                Else
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    
                End If
                If j = 1 Then .ParagraphFormat.Alignment = wdAlignParagraphLeft
                
                .Font.ColorIndex = wdDarkBlue
            End With
'            If i = 1 And j = .Tables(10).Columns.Count Then
'                Set Rng = .Tables(10).Cell(i, 1).Range
'                Rng.End = .Tables(10).Cell(i, .Tables(10).Columns.Count).Range.End
'                Rng.Cells.Merge
'                .Tables(10).Cell(i, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
'            End If

        Next j
    Next i
    Set Rng = .Tables(10).Cell(1, 1).Range
    Rng.End = .Tables(10).Cell(1, .Tables(10).Columns.Count).Range.End
    Rng.Cells.Merge
    .Tables(10).Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
'======================================
'FILL TABLE 11
'======================================
    
    For i = 1 To .Tables(11).Rows.Count
        For j = 1 To .Tables(11).Columns.Count
        .Tables(11).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
        
            .Tables(11).Columns.Width = Application.CentimetersToPoints(2.166)
            With .Tables(11).Cell(i, j).Range
                .Text = ws.Cells(i + 11, j + 6)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                If i = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                ElseIf i = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Bold = True
                    .Font.Size = 8
                Else
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    
                End If
                If j = 1 Then .ParagraphFormat.Alignment = wdAlignParagraphLeft
                
                .Font.ColorIndex = wdDarkBlue
            End With

        Next j
    Next i
    Set Rng = .Tables(11).Cell(1, 1).Range
    Rng.End = .Tables(11).Cell(1, .Tables(11).Columns.Count).Range.End
    Rng.Cells.Merge
    .Tables(11).Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
    
    
'======================================
'FILL TABLE 12
'======================================
    
    For i = 1 To .Tables(12).Rows.Count
        For j = 1 To .Tables(12).Columns.Count
        .Tables(12).Cell(i, j).VerticalAlignment = wdCellAlignVerticalBottom
            .Tables(12).Columns.Width = Application.CentimetersToPoints(1)
            With .Tables(12).Cell(i, j).Range
                .Text = ws.Cells(i + 17, j + 6)
                .ParagraphFormat.Alignment = wdAlignParagraphRight
                .Font.Size = 8
                If i = 1 Then
                    .Font.Bold = True
                    .ParagraphFormat.Alignment = wdAlignParagraphRight
                ElseIf i = 2 Then
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Bold = True
                    .Font.Size = 8
                    
                Else
                    .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    .Font.Size = 7
                End If
                If j = 1 Then .ParagraphFormat.Alignment = wdAlignParagraphLeft
                
                
                .Font.ColorIndex = wdDarkBlue
            End With

        Next j
    Next i
    Set Rng = .Tables(12).Cell(1, 1).Range
    Rng.End = .Tables(12).Cell(1, .Tables(12).Columns.Count).Range.End
    Rng.Cells.Merge
    .Tables(12).Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    l_par = .Paragraphs.Count
    
    
    .Paragraphs(l_par).Range.InsertBreak Type:=wdPageBreak
   
   
 lineOftext = .Sections(1).Headers(wdHeaderFooterFirstPage)
    With lineOftext.Find
        .Text = "<Mese>"
        .Replacement.Text = "Fondo FERI PIR" & vbCrLf & "Report as of " & Format(Now(), "dd/mm/yy")
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    
      
 '===========================
'create table14
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(14).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(14).Rows.HorizontalPosition = Application.CentimetersToPoints(0)
    .Tables(14).Rows.VerticalPosition = Application.CentimetersToPoints(3.6)
    .Tables(14).Columns(1).Width = Application.CentimetersToPoints(8.5)
    .Tables(14).Rows.HeightRule = 1
    .Tables(14).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(14).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
    
'======================================
'FILL TABLE 14
'======================================
    
    With .Tables(14).Cell(1, 1).Range
        .Text = "TOP 10 Holding"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 8
        .Font.Bold = True
        .Font.ColorIndex = wdDarkBlue

    End With
    
 '===========================
'create table15
'===========================
    Set wrdRange = .Paragraphs.Last.Range
    .Tables.Add Range:=wrdRange, NumRows:=1, NumColumns:=1
    .Tables(15).Rows.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Tables(15).Rows.HorizontalPosition = Application.CentimetersToPoints(10)
    .Tables(15).Rows.VerticalPosition = Application.CentimetersToPoints(3.6)
    .Tables(15).Columns(1).Width = Application.CentimetersToPoints(8.5)
    .Tables(15).Rows.HeightRule = 1
    .Tables(15).Rows.Height = Application.CentimetersToPoints(0.4)

    With .Tables(15).Rows(1).Cells.Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth100pt
        .Color = wdColorDarkBlue
    End With
           
 '======================================
'FILL TABLE 15
'======================================
    .Tables(15).Cell(1, 1).VerticalAlignment = wdCellAlignVerticalBottom
    With .Tables(15).Cell(1, 1).Range
        .Text = "Industry Breakdown"
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.Size = 8
        .Font.Bold = True
        .Font.ColorIndex = wdDarkBlue

    End With
    
''======================================
''put images 3 and 4
''======================================
    l_par = .Paragraphs.Count
    .Paragraphs(l_par).Range.InlineShapes.AddPicture ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\top_10_holding.jpg")
    .Content.InsertParagraphAfter
'
    With ActiveDocument.InlineShapes(1)
        'border
        'conversion to Shape
        .ConvertToShape
    End With

    With ActiveDocument.Shapes(3)
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Top = Application.CentimetersToPoints(4.29)
        .Left = Application.CentimetersToPoints(1.51)
        .Width = Application.CentimetersToPoints(8.5)
        .Height = Application.CentimetersToPoints(6.77)
    End With
    
    
    .Paragraphs(l_par + 1).Range.InlineShapes.AddPicture ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\industry_allocation.jpg")
    .Content.InsertParagraphAfter
    
    With ActiveDocument.InlineShapes(1)
        'border
        'conversion to Shape
        .ConvertToShape
    End With
    
    With ActiveDocument.Shapes(4)
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .Top = Application.CentimetersToPoints(4.29)
        .Left = Application.CentimetersToPoints(11.51)
        .Width = Application.CentimetersToPoints(8.5)
        .Height = Application.CentimetersToPoints(6.77)
    End With

    
    


    
End With

End Sub
