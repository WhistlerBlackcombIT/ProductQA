Attribute VB_Name = "QA_Declarations"
' Dates
Public intNextYear
Public intNextYearFull
Public intNextYear2Digit
Public intThisYear
Public intThisYearFull
Public intThisYear2Digit
Public intLastYear
Public intLastYearFull
Public intLastYear2Digit
Public intThisSeason
Public intNextSeason
Public intLastSeason

' Colours
Public clrGreen As Long, clrOrange As Long, clrRed As Long _
    , clrBlack As Long, clrYellow As Long, clrMagenta As Long
Public clrNone As Long, clrColumnHeader As Long, clrColumnHeaderLast As Long _
    , clrHighlight1 As Long, clrHighlight2 As Long

' Typeface/Font
Public fntFace As String, fntSize As Integer


' =====================================================
' QA_Colors - Default color declarations
' =====================================================
Sub QA_Formatting()
Attribute QA_Formatting.VB_ProcData.VB_Invoke_Func = " \n14"
    ' DO NOT CHANGE THESE! Re-define them in QA_Formatting_Custom

    ' Colors
    clrGreen = RGB(152, 204, 0) ' CI 43
    clrYellow = RGB(255, 255, 0) ' CI 6
    clrOrange = RGB(255, 152, 0) ' CI 45
    clrBlue = RGB(173, 216, 230)
    clrRed = RGB(255, 0, 0) ' CI 3
    clrBlack = RGB(0, 0, 0) ' CI 1
    clrMagenta = RGB(255, 0, 255) ' CI 7
    clrGray = RGB(192, 192, 192) ' CI 15
    clrWhite = RGB(255, 255, 255) ' CI 2
    clrNone = xlColorIndexNone ' NOTHING!
    clrColumnHeader = clrBlue
    clrColumnHeaderLast = clrYellow
    clrHighlight1 = RGB(192, 192, 192) ' or CI 35 green 204 255 204
    clrHighlight2 = RGB(150, 150, 150) ' or CI 37 blue 153 204 255

    ' Typeface/font
    fntFace = "Arial"
    fntSize = 9

    ' Pull in the user-defined color overrides if any
    QA_Formatting_Custom
End Sub


' =====================================================
' ApplyCF - Apply conditional formatting to a range of cells
' =====================================================
Sub ApplyCF(rng As Range, fml As String, clr As Long)
    rng.Select
    With Selection.FormatConditions
        .Add Type:=xlExpression, Formula1:=fml
    End With
    With Selection
        With .FormatConditions(.FormatConditions.Count)
            .Interior.Color = clr
            .StopIfTrue = False
        End With
    End With
End Sub


' =====================================================
' ApplyFilter - Apply filtering to a range of cells
' =====================================================
Sub ApplyFilter(fld As String, crt As String)
    ActiveSheet.AutoFilterMode = False
    [A:AZ].AutoFilter Field:=L2N(fld), Criteria1:=crt
End Sub


' =====================================================
' ApplyFilterArray - Apply an array of filters to a range of cells
' =====================================================
Sub ApplyFilterArray(fld As String, crt As Variant)
    ' You CANNOT send more than 2 string-based criteria (begins with: "text*") in your array else you will get NOTHING
    ActiveSheet.AutoFilterMode = False
    [A:AZ].AutoFilter Field:=L2N(fld), Criteria1:=crt, Operator:=xlFilterValues
End Sub


' =====================================================
' ApplyFilterColor - Filter column by cell colour (usually set by ApplyCF)
' =====================================================
Sub ApplyFilterColor(fld As String, clr As Long)
    [A:AZ].AutoFilter Field:=L2N(fld), Criteria1:=clr, Operator:=xlFilterCellColor
End Sub


' =====================================================
' ApplySort - Apply sorting to a spreadsheet
' =====================================================
Sub ApplySort(ParamArray args() As Variant)
    With ActiveSheet.Sort
        With .SortFields
            .Clear

            For i = LBound(args) To UBound(args)
                .Add Key:=Range(args(i) & ":" & args(i)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            Next
        End With

        .SetRange Range("A:AZ")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


' =====================================================
' ApplyStripZ - Remove z-prefixed rows from a range
' =====================================================
Sub ApplyStripZ(fld As String)
    [A:AZ].AutoFilter Field:=L2N(fld), Criteria1:="<>z*"
End Sub


' =====================================================
' L2N - Converts a letter to an equivalent column number
' =====================================================
Function L2N(col As String) As Integer
    For i = 1 To Len(col)
        num = (Asc(UCase(Mid(col, i, 1))) - 64) + num * 26
    Next i
    L2N = num
End Function


' =====================================================
' DataRange - Return the range of cells actually containing data
' @param Optional integer column-offset (value 1 starts selection from column B instead of default A)
' =====================================================
Function DataRange(Optional colOffset As Integer, Optional rowOffset As Integer) As Range
    Set DataRange = Range(Cells(2, 1 + colOffset), Cells(ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row, ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column))
End Function


' =====================================================
' IsInArray - Return bool if value exists in array
' =====================================================
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


Public Function ChangeTabColorBasedOnValidationColorInSheet(intSheet As Integer)
    ' Pull in the colours we've configured elsewhere
    QA_Formatting

    ' Var declarations
    Dim X, Y As Integer
    Dim Z As Long

    ' If there aren't any conditions, don't bother looking any further
    If Worksheets(intSheet).UsedRange.FormatConditions.Count < 1 Then Exit Function

    ' Step through each cell in the sheet and determine if conditional colours and change tab-color to match
    With Worksheets(intSheet)
        For X = 2 To .UsedRange.Rows.Count
            For Y = 1 To .UsedRange.Columns.Count
                Z = getCellColorForReals(Range(Cells(X, Y), Cells(X, Y)))
    
                Select Case Z
                    Case 16777215
                    Case xlColorIndexNone
                    Case 5296274 ' Green
                    Case clrGreen ' Green
                    Case clrWhite ' White
                    Case clrHighlight1
                    Case clrHighlight2
                    Case Else
                        Exit For
                End Select
            Next
    
            Select Case Z
                Case 16777215
                Case xlColorIndexNone
                Case 5296274 ' Green
                Case clrGreen ' Green
                Case clrWhite ' White
                Case clrHighlight1
                Case clrHighlight2
                Case Else
                    'MsgBox "Row: " & X & ", Col: " & Y & ", Clr: " & Z & ", Val: " & .Range(Cells(X, Y), Cells(X, Y)).Value
                    .Tab.Color = Z
                    Exit For
            End Select
        Next
    
    End With
End Function


Function getCellColorForReals(r As Range) As Long
    getCellColorForReals = r.DisplayFormat.Interior.Color
End Function

