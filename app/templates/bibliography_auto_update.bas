Attribute VB_Name = "BibliographyAutoUpdate"
Option Explicit

Public Sub AutoOpen()
    UpdateThesisFields
End Sub

Public Sub Document_Open()
    UpdateThesisFields
End Sub

Public Sub UpdateThesisFields()
    Dim toc As TableOfContents
    Dim storyRange As Range

    On Error Resume Next

    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next toc

    For Each storyRange In ActiveDocument.StoryRanges
        storyRange.Fields.Update
    Next storyRange

    ActiveDocument.Fields.Update
End Sub
