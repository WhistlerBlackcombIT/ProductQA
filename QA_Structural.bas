Attribute VB_Name = "QA_Structural"
'
'
' TODO: Popup at run with select box for season (this season summer/winter plus one either way +/-)
'



' =====================================================
' QA_Structure - Change some sheet names, strip some crap, hide empty things yadda yadda
' =====================================================
Sub QA_Structure()
Attribute QA_Structure.VB_ProcData.VB_Invoke_Func = "Q\n14"
    ' Present season selection form
    'frmSeason.Show (1)

    ' Load formatting definitions
    QA_Formatting

    Application.DisplayAlerts = False
    For Each ws In Worksheets
        ws.Activate

        ' Delete first row of 21+ to make compatible
        If ws.Index > 20 Then
            Rows("1:1").EntireRow.Delete
        End If

        ' Move name to sheet tab
        ws.Name = Range("$A$1").Value
        
        'Delete first row (has the name and don't need it)
        Rows("1:1").EntireRow.Delete
        ' Hide sheet if contains no data
        If Range("B2").Value = vbNullString Then ws.Visible = False

        'Format All sheets same with Auto filters and freeze panes
        Cells.Select
        With Selection
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .MergeCells = False
            .Font.Name = fntFace
            .Font.Size = fntSize
        End With
        Rows("1:1").Select
        With Selection
                .HorizontalAlignment = xlCenter
            .WrapText = False
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .MergeCells = False
            .Font.Name = fntFace
            .Font.Size = fntSize
            .Font.Bold = True
        End With
        Cells.Select
        Cells.EntireColumn.AutoFit
        Rows("2:2").Select
        ActiveWindow.FreezePanes = True
        Rows("1:1").Select
        Selection.AutoFilter

        ' Identify last column in data set
        LastColumn = Application.WorksheetFunction.CountA(Range("1:1"))
        ' Change color of table headers
        Range(Cells(1, 1), Cells(1, LastColumn)).Interior.Color = clrColumnHeader
        ' Highlight last column header yellow (to clearly identify last one)
        Cells(1, LastColumn).Interior.Color = clrColumnHeaderLast
        ' Select first cell
        Cells(1, 2).Select
    Next

    Application.DisplayAlerts = True
    ActiveWindow.TabRatio = 0.8
    Sheets(1).Select
End Sub
