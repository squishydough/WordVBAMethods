'----------------------------------------------------------------------------
'   Description     :   Opens a userform and centers it on the user's screen
'----------------------------------------------------------------------------
Public Sub CenterUserForm(frm As Object)
    With frm
        .StartUpPosition = 0
        .Left = Application.ActiveWindow.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show False
    End With
End Sub

'----------------------------------------------------------------------------
'   Description     :   Clears all DocVariables
'----------------------------------------------------------------------------
Public Sub ClearDocVariables()
    Dim i As Long
    
    For i = ThisDocument.Variables.Count To 1 Step -1
        ThisDocument.Variables(i).Delete
    Next i
    ThisDocument.Fields.Update
End Sub

'----------------------------------------------------------------------------
'   Description     :   Since Word lacks a way to get a table height, this
'   method does just that.  Returns the Points value as a single rounded to 
'   two decimal places.
'----------------------------------------------------------------------------
Public Function GetTableHeight(tbl As Table) As Single
    Dim i As Long
    Dim TableHeight As Single
    TableHeight = 0

    '**
    '*  Add a dummy row at the bottom of the table
    '*  This ensures the last table row with data also gets resized
    '**
    tbl.Rows.Add
    With tbl.Rows(tbl.Rows.Count)
        .Height = 1
        .HeightRule = wdRowHeightExactly
    End With

    '**
    '*  Calculate child table height
    '**
    For i = 1 To tbl.Rows.Count - 1 Step 1
        With tbl.Rows(i)
            TableHeight = TableHeight + tbl.Rows(i + 1).Range.Information(wdVerticalPositionRelativeToPage) - tbl.Rows(i).Range.Information(wdVerticalPositionRelativeToPage)
        End With
    Next i
    
    '**
    '*  Delete the dummy row
    '**
    tbl.Rows(tbl.Rows.Count).Delete
    
    GetTableHeight = Format(TableHeight, "#.##")
End Function

'---------------------------------------------------------------------------
'   Description :   Calls the GetTableHeight method, but converts the 
'   Points to Inches
'---------------------------------------------------------------------------
Public Function GetTableHeightInInches(tbl As Table) As Single
    GetTableHeightInches = Format(PointsToInches(GetTableHeight(tbl)), "#.##")
End Function

'---------------------------------------------------------------------------
'   Description :   Converts text to a hyperlink.  Allows you to optionally
'   pass a WdColor (wdWhite) to change the hyperlink color
'---------------------------------------------------------------------------
Public Sub MakeTextHyperlink(TargetRange As Range, SearchText As String, LinkAddress As String, Optional LinkColor As WdColor)
    With TargetRange
        With .Find
            .Text = SearchText
            .MatchWholeWord = True
            .Forward = True
            .Execute
        End With
        
        If .Find.Found Then
            .Select
            ThisDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=LinkAddress
            If Not IsMissing(LinkColor) Then .Font.ColorIndex = LinkColor
        End If
    End With
End Sub