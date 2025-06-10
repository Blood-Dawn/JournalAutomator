Attribute VB_Name = "WordJournalMacros"
Option Explicit

' Apply two-column layout starting at the given section index (1-based)
Sub ApplyTwoColumnLayout(startPage As Integer)
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        If sec.Index >= startPage Then
            With sec.PageSetup.TextColumns
                .SetCount NumColumns:=2
                .EvenlySpaced = True
                .LineBetween = False
            End With
        End If
    Next sec
End Sub

' Add simple left and right page borders beginning with the given section
Sub AddPageBorders(startSection As Integer)
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        If sec.Index >= startSection Then
            With sec.PageSetup
                .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                .Borders(wdBorderLeft).LineWidth = wdLineWidth025pt
                .Borders(wdBorderLeft).Color = wdColorBlack
                .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                .Borders(wdBorderRight).LineWidth = wdLineWidth025pt
                .Borders(wdBorderRight).Color = wdColorBlack
            End With
        End If
    Next sec
End Sub

' Bold and center the first paragraph of the document
Sub FormatFrontCover()
    If ActiveDocument.Paragraphs.Count = 0 Then Exit Sub
    With ActiveDocument.Paragraphs(1)
        .Alignment = wdAlignParagraphCenter
        .Range.Font.Bold = True
    End With
End Sub

' Center footer text across all sections
Sub LayoutFooter()
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        With sec.Footers(wdHeaderFooterPrimary).Range.ParagraphFormat
            .Alignment = wdAlignParagraphCenter
        End With
    Next sec
End Sub

' Apply front cover and footer formatting with optional font size and spacing
Sub FormatFrontAndFooter(Optional fontSize As Integer = 0, Optional lineSpacing As Single = 0)
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If InStr(1, p.Range.Text, "Volume") > 0 Then
            If lineSpacing > 0 Then
                p.LineSpacingRule = wdLineSpaceExactly
                p.LineSpacing = lineSpacing * 12
            End If
            If fontSize > 0 Then
                p.Range.Font.Size = fontSize
            End If
            Exit For
        End If
    Next p

    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        For Each p In sec.Footers(wdHeaderFooterPrimary).Range.Paragraphs
            If lineSpacing > 0 Then
                p.LineSpacingRule = wdLineSpaceExactly
                p.LineSpacing = lineSpacing * 12
            End If
            If fontSize > 0 Then
                p.Range.Font.Size = fontSize
            End If
        Next p
    Next sec
End Sub

' Update the front cover information block
Sub UpdateFrontCover(volume As String, issue As String, monthYear As String)
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If InStr(1, p.Range.Text, "Volume") > 0 Then
            p.Range.Text = "Volume " & volume & ", Issue " & issue & vbCr & monthYear
            p.Range.Font.Bold = True
            p.Alignment = wdAlignParagraphCenter
            Exit Sub
        End If
    Next p
End Sub

' Replace the header text on page 2
Sub UpdatePage2Header(newHeader As String, pageNum As Integer)
    Dim sec As Section
    Dim headerText As String
    headerText = newHeader & vbCr & "Page " & pageNum
    For Each sec In ActiveDocument.Sections
        With sec.Headers(wdHeaderFooterPrimary).Range.Paragraphs(1)
            If Len(.Range.Text) > 1 Then
                .Range.Text = headerText
                Exit Sub
            End If
        End With
    Next sec
End Sub

