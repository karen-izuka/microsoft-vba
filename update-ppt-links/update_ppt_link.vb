Option Explicit

Sub HyperLinkSearchReplace()
    Dim slide As slide
    Dim link As Hyperlink
    Dim search_string As String
    Dim replace_string As String
    Dim shape As shape
    Dim i As Integer

    search_string = InputBox("What text should I search for?", "Search Box")
    If search_string = "" Then
        Exit Sub
    End If

    replace_string = InputBox("What text should I replace" & vbCrLf & search_string _ 
        & vbCrLf & "with?", "Replace Box")
    If replace_string = "" Then
        Exit Sub
    End If
    On Error Resume Next

    For Each slide In ActivePresentation.Slides
        For Each link In slide.Hyperlinks
            link.Address = Replace(link.Address, search_string, replace_string)
            link.SubAddress = Replace(link.SubAddress, search_string, replace_string)
        Next

        i = 1
        For Each shape In slide.Shapes
            If shape.Type = msoLinkedOLEObject Or shape.Type = msoMedia Then
                shape.LinkFormat.SourceFullName = Replace(shape.LinkFormat.SourceFullName, search_string, replace_string)
                i = i + 1
            End If
        Next
    Next
    MsgBox "All Done!"
End Sub
