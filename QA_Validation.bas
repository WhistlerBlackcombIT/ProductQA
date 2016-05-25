Attribute VB_Name = "QA_Validation"

' =====================================================
' QA_Structure - MAEK IT LOOK PRATTY!
' =====================================================
Sub QA_Validate()
Attribute QA_Validate.VB_ProcData.VB_Invoke_Func = "W\n14"
    ' Date definitions
    intNextYear = Year(Date) + 1
    intThisYear = Year(Date)
    intLastYear = Year(Date) - 1
    intNextYearFull = DateSerial(intNextYear, 1, 1)
    intNextYear2Digit = Format(intNextYearFull, "YY")
    intThisYearFull = DateSerial(intThisYear, 1, 1)
    intThisYear2Digit = Format(intThisYearFull, "YY")
    intLastYearFull = DateSerial(intLastYear, 1, 1)
    intLastYear2Digit = Format(intLastYearFull, "YY")

    ' Season dates
    intNextSeason = intNextYear
    intThisSeason = intThisYear
    intLastSeason = intLastYear
    'If (1 <= Month(Date) <= 5) Then
    '    intNextSeason = intNextYear - 1
    '    intThisSeason = intThisYear - 1
    '    intLastSeason = intLastYear - 1
    'End If
    intNextSeasonFull = DateSerial(intNextSeason, 1, 1)
    intNextSeason2Digit = Format(intNextSeasonFull, "YY")
    intThisSeasonFull = DateSerial(intThisSeason, 1, 1)
    intThisSeason2Digit = Format(intThisSeasonFull, "YY")
    intLastSeasonFull = DateSerial(intLastSeason, 1, 1)
    intLastSeason2Digit = Format(intLastSeasonFull, "YY")
'MsgBox "Last: " & intLastYear & " This: " & intThisYear & " Next: " & intNextYear
'End
    ' Load formatting definitions
    QA_Formatting

    For Each ws In Worksheets
        ' Activate current worksheet
        ws.Activate

        ' Strip existing conditional formatting from sheets
        [A:ZZ].FormatConditions.Delete

        Select Case ws.Index
            Case 1 'Product Header
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True

                ' Highlight where Blackout is NOT "Days=2" OR "Always Available"
                ApplyCF [Z:AA], "=AND( ROW()>1, LEN($Z1)>0, AND( $Z1<>""Always Available"", OR( $Z1<>""Days"", AND( $Z1=""Days"", $AA1<>2 ) ) ) )", clrRed
                ' Highlight Security Levels
                ApplyCF [J:J], "=AND(LEN($J1)>0,OR($J1=""0"",$J1=0))", clrNone
                ApplyCF [J:J], "=OR($J1=""1"",$J1=1)", clrGreen
                ApplyCF [J:J], "=OR($J1=""2"",$J1=2)", clrYellow
                ApplyCF [J:J], "=OR($J1=""3"",$J1=3)", clrOrange
                ApplyCF [J:J], "=OR($J1=""4"",$J1=4)", clrMagenta
                ApplyCF [J:J], "=OR($J1=""5"",$J1=5)", clrRed


            ' ==================================================================================
            Case 2 'Accounting
                ' Hide colummns we don't much care about
                [A:C].EntireColumn.Hidden = True
                [F:G].EntireColumn.Hidden = True
                [Q:U].EntireColumn.Hidden = True
                'Application.GoTo ws.Range("V1"), True
                '[AY1].Select

                ' Match default against actual
                ' Earned Segment 3
                ApplyCF [X:X], "=AND( ROW()>1, $X1<>$AD1 )", clrRed
                ApplyCF [AD:AD], "=AND( ROW()>1, $X1<>$AD1 )", clrRed
                ' Earned Segment 4
                ApplyCF [Y:Y], "=AND( ROW()>1, $Y1<>$AE1 )", clrRed
                ApplyCF [AE:AE], "=AND( ROW()>1, $Y1<>$AE1 )", clrRed
                ' Unearned Segment 3
                ApplyCF [Z:Z], "=AND( ROW()>1, $Z1<>$AH1 )", clrRed
                ApplyCF [AH:AH], "=AND( ROW()>1, $Z1<>$AH1 )", clrRed
                ' Unearned Segment 4
                ApplyCF [AA:AA], "=AND( ROW()>1, $AA1<>$AI1 )", clrRed
                ApplyCF [AI:AI], "=AND( ROW()>1, $AA1<>$AI1 )", clrRed
                ' Earned Segment 2 (revenue location)
        ' Disabled until revenue names are updated with SAP codes
                'ApplyCF [W:W], "=AND( ROW()>1, MID($W1,FIND(""("",$W1)+1,FIND("")"",$W1)-FIND(""("",$W1)-1)<>$AC1 )", clrRed
                'ApplyCF [AC:AC], "=AND( ROW()>1, MID($W1,FIND(""("",$W1)+1,FIND("")"",$W1)-FIND(""("",$W1)-1)<>$AC1 )", clrRed


            ' ==================================================================================
            Case 3 'Component
                ' Hide yo kids
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [K:L].EntireColumn.Hidden = True

                ' Highlight Access Rule Code if not 4-chars
                ApplyCF [W:W], "=AND(NOT(ROW()=1),LEN($W1)>0,LEN($W1)<>4)", clrRed
                ' Highlight non-zero starting Access Rule Code
                ApplyCF [W:W], "=AND(NOT(ROW()=1),$T1=""Access Product"",LEN($W1)>0,NOT(LEFT($W1,1)=""0""))", clrOrange


            ' ==================================================================================
            Case 4 'Pricing PS
                ' Hide yo wife
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [L:M].EntireColumn.Hidden = True

                ' Filter only pricing seasons from this year and previous year
                'ApplyFilterArray "O", Array(intThisSeason & "*", intLastSeason & "*")
                ApplyFilterArray "O", Array("??" & intThisSeason2Digit & "*", "??" & intLastSeason2Digit & "*")
                ', intThisYear2Digit & "*", intLastYear2Digit & "*"
                ' You CANNOT send more than 2 "begins with" criteria in your array else you will get NOTHING

                ' Sort pricing (PHeader, Season, Product)
                ApplySort "E", "O", "K"
                ' Sort pricing (PHOrder, Product, Season, Channel)
                'ApplySort "H", "K", "O", "I"

                ' Conditional formatting for this year's pricing
                ApplyCF [K:P], "=AND(LEN($O1)>0,OR(LEFT(SUBSTITUTE($O1,""" + CStr(intThisSeason) + """,""" + CStr(intThisSeason2Digit) + """),2)=""" + CStr(intThisSeason2Digit) + """,LEFT(SUBSTITUTE($O1,""" + CStr(intNextSeason) + """,""" + CStr(intNextSeason2Digit) + """),2)=""" + CStr(intNextSeason2Digit) + """))", clrHighlight1

                ' Compare dates in PS name to Min/Max
                ' YEAR()+1 versions work for winter
                ' Other one works in summer
' @todo This needs to be switched depending on choices from pop-up box
                ApplyCF [T:T], "=AND(MID($O1,FIND(1,$O1),2)=""" + CStr(intThisYear2Digit) + """,DATEVALUE(CONCATENATE(SUBSTITUTE(MID($O1,FIND(""("",$O1)+4,2),""-"",""""),""-"",MID($O1,FIND(""("",$O1)+1,3),""-"",IF(OR(MID($O1,FIND(""("",$O1)+1,3)=""Nov"",MID($O1,FIND(""("",$O1)+1,3)=""Dec""),YEAR(NOW()),YEAR(NOW())+1)))<>$T1)", clrRed
                'ApplyCF [T:T], "=AND(MID($O1,FIND(1,$O1),2)=""" + CStr(intThisYear2Digit) + """,DATEVALUE(CONCATENATE(SUBSTITUTE(MID($O1,FIND(""("",$O1)+4,2),""-"",""""),""-"",MID($O1,FIND(""("",$O1)+1,3),""-"",IF(OR(MID($O1,FIND(""("",$O1)+1,3)=""Nov"",MID($O1,FIND(""("",$O1)+1,3)=""Dec""),YEAR(NOW()),YEAR(NOW()))))<>$T1)", clrRed
                ApplyCF [U:U], "=AND(MID($O1,FIND(1,$O1),2)=""" + CStr(intThisYear2Digit) + """,DATEVALUE(CONCATENATE(SUBSTITUTE(MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+4,2),"")"",""""),""-"",MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+1,3),""-"",IF(OR(MID($O1,FIND("")"",$O1)-5,3)=""Nov"",MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+1,3)=""Dec""),YEAR(NOW()),YEAR(NOW())+1)))<>$U1)", clrRed
                'ApplyCF [U:U], "=AND(MID($O1,FIND(1,$O1),2)=""" + CStr(intThisYear2Digit) + """,DATEVALUE(CONCATENATE(SUBSTITUTE(MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+4,2),"")"",""""),""-"",MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+1,3),""-"",IF(OR(MID($O1,FIND("")"",$O1)-5,3)=""Nov"",MID($O1,FIND(""-"",$O1,FIND(""("",$O1))+1,3)=""Dec""),YEAR(NOW()),YEAR(NOW()))))<>$U1)", clrRed


            ' ==================================================================================
            Case 5: ' Pricing DR
                ' Hide some more stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [L:M].EntireColumn.Hidden = True

                ' Filter only pricing seasons from this year
                ApplyFilter "O", ">" & intThisSeasonFull

                ' Sort pricing (PHeader, Season, Product)
                ApplySort "E", "O", "K"
                ' Sort pricing (PHOrder, Product, Season, Channel)
                'ApplySort "H", "K", "O", "I"

                ' Conditional formatting for this year's pricing
                ApplyCF [O:P], "=AND(LEN($O1)>0,YEAR($O1)=" + CStr(intNextSeason) + ")", clrHighlight1
                ' CHANGE: Make it check if date > start of appropriate season (5/1/XX summer, 11/1/XX winter)

                ' Highlight taxes
                ApplyCF [V:AE], "=AND(ROW()>1,LEN(V1)>0)", clrHighlight1


            ' ==================================================================================
            Case 6: ' Access Product
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [K:L].EntireColumn.Hidden = True

                ' Highlight Expiration Date neither this year nor next year where Effective Type is FIXED
                ApplyCF [Q:Q], "=AND($P1=""FIXEDDATE"",NOT(OR(YEAR($Q1)=" + CStr(intThisSeason) + ",YEAR($Q1)=" + CStr(intNextSeason) + ")))", clrRed
                ApplyCF [T:T], "=AND($P1=""FIXEDDATE"",NOT(OR(YEAR($T1)=" + CStr(intThisSeason) + ",YEAR($T1)=" + CStr(intNextSeason) + ")))", clrRed
                ' Highlight Access Rule Code if not 4-chars
                ApplyCF [M:M], "=AND(NOT(ROW()=1),LEN($M1)>0,LEN($M1)<>4)", clrRed
                ' Highlight non-zero starting Access Rule Code
                ' DEFUNCT: not using since spreading scope of access codes in to non-zero prefix territory
                ' ApplyCF [M:M], "=AND(NOT(ROW()=1),LEN($M1)>0,NOT(LEFT($M1,1)=""0""))", clrOrange

                ' Expiration Days that don't match product header
                ApplyCF [S:S], "=AND(ROW()>1,FIND(""Lift"",E1),NOT(OR(AND($S1=1,INT(MID(J1,FIND(""D Lift"",J1)-2,2))=1),AND(INT(MID(J1,FIND(""D Lift"",J1)-2,2))>1,INT(MID(J1,FIND(""D Lift"",J1)-2,2))<=9,$S1=INT(MID(J1,FIND(""D Lift"",J1)-2,2))+1),AND(INT(MID(J1,FIND(""D Lift"",J1)-2,2))>=10,INT(MID(J1,FIND(""D Lift"",J1)-2,2))<=14,$S1=INT(MID(J1,FIND(""D Lift"",J1)-2,2))+2),AND(INT(MID(J1,FIND(""D Lift"",J1)-2,2))>=15,INT(MID(J1,FIND(""D Lift"",J1)-2,2))<=21,$S1=INT(MID(J1,FIND(""D Lift"",J1)-2,2))+3))))", clrRed

                ' Highlight difference between Never Preload and Always Preload
                ApplyCF [U:U], "=AND(ROW()>1,FIND(""Always"",U1))", clrHighlight1
                ApplyCF [U:U], "=AND(ROW()>1,FIND(""Never"",U1))", clrHighlight2

            ' ==================================================================================
            Case 7: ' Access Rule
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [K:L].EntireColumn.Hidden = True

                ' Filter out Access Rule Location Group: Peak 2 Peak
                ApplyFilter "M", "<>Peak 2 Peak"
                ' Highlight non-2025 expiry cells
                ApplyCF Range("U:U"), "=AND(LEN($U1)>0,NOT(YEAR($U1)=2025))", clrRed
                ' Highlight non-zero starting Access Rule Code
                ApplyCF [N:N], "=AND(NOT(ROW()=1),LEN($N1)>0,LEN($N1)<>4)", clrRed
                ' Highlight non-zero starting Access Rule Code
                ApplyCF [N:N], "=AND(NOT(ROW()=1),LEN($N1)>0,NOT(LEFT($N1,1)=""0""))", clrOrange
                ' Max Days/Total Days validity check (2 of 3, 10 of 12, 15 of 18)
                ' RFID
                ApplyCF [Y:Z], "=AND(NOT(ROW()=1),FIND(""RFID"",$J1),FIND(""Lift"",$R1),NOT(OR(AND($Y1=1,$Z1=1),AND($Y1>1,$Y1<=9,$Z1=$Y1+2),AND($Y1>=10,$Y1<=14,$Z1=$Y1+3),AND($Y1>=15,$Y1<=21,$Z1=$Y1+4))))", clrRed
                ' Everything else
                ApplyCF [Y:Z], "=AND(NOT(ROW()=1),NOT(IFERROR(FIND(""RFID"",$J1),FALSE)),FIND(""Lift"",$R1),NOT(OR(AND($Y1=1,$Z1=1),AND($Y1>1,$Y1<=9,$Z1=$Y1+1),AND($Y1>=10,$Y1<=14,$Z1=$Y1+2),AND($Y1>=15,$Y1<=21,$Z1=$Y1+3))))", clrRed
                ' Usage products
                ApplyCF [AT:AT], "=AND(NOT(ROW()=1),LEN($AT1)>0)", clrOrange

            ' ==================================================================================
            Case 8: ' Output
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True

                ' Sort media templates (PHeader, Product, Channel)
                ApplySort "H", "K", "O", "I"

                ' Highlight if print label is different for same component on different sales channels
                ApplyCF [R:Z], "=IF(AND(INDIRECT(ADDRESS(ROW()-1,11))=INDIRECT(ADDRESS(ROW(),11)),LEN(ADDRESS(ROW()-1,COLUMN()))>0,INDIRECT(ADDRESS(ROW()-1,COLUMN()))<>INDIRECT(ADDRESS(ROW(),COLUMN()))),TRUE,FALSE)", clrOrange

                ' Highlight if stock type isn't RFID when product header or media template is
                'ApplyCF [AI:AI], "=AND(ROW()>1,OR(IFERROR(FIND(""RFID"",O1),FALSE),IFERROR(FIND(""RFID"",AI1),FALSE)))", clrOrange

                ' Highlight if last year's date found in print label
                ApplyCF [R:R], "=FIND(""" & intLastYear & """, $R1)", clrRed
                ApplyCF [S:S], "=FIND(""" & intLastYear & """, $S1)", clrRed
                ApplyCF [T:T], "=FIND(""" & intLastYear & """, $T1)", clrRed
                ApplyCF [U:U], "=FIND(""" & intLastYear & """, $U1)", clrRed
                ApplyCF [V:V], "=FIND(""" & intLastYear & """, $V1)", clrRed
                ApplyCF [W:W], "=FIND(""" & intLastYear & """, $W1)", clrRed
                ApplyCF [X:X], "=FIND(""" & intLastYear & """, $X1)", clrRed
                ApplyCF [Y:Y], "=FIND(""" & intLastYear & """, $Y1)", clrRed
                ApplyCF [Z:Z], "=FIND(""" & intLastYear & """, $Z1)", clrRed

            ' ==================================================================================
            Case 9: ' Inventory Pools
                ' Hide some stuff
                [A:B].EntireColumn.Hidden = True
                [D:D].EntireColumn.Hidden = True


            ' ==================================================================================
            Case 10: ' Private
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                [K:L].EntireColumn.Hidden = True


            ' ==================================================================================
            Case 11 ' Tax
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True


            ' ==================================================================================
            Case 12 'Rezo
                ' Change colour of tab to show RRW availability
                ws.Tab.Color = clrMagenta

                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True

                ' Filter only Direct Booking lines
                ApplyFilter "I", "Direct Booking"

                
            ' ==================================================================================
            Case 13: ' Rest
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True
                
                
            ' ==================================================================================
            Case 14: ' Cal Def
                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True


            ' ==================================================================================
            Case 15 ' Voucher
                ' Change tab color to stand out
                ws.Tab.Color = clrYellow

                ' Hide some stuff
                [A:C].EntireColumn.Hidden = True
                [F:H].EntireColumn.Hidden = True

                ' Highlight if expiration isn't this year
                ApplyCF [W:W], "=AND($V1=""Date"",0<LEN($W1),YEAR($W1)<" & intThisSeason & ")", clrRed
                ApplyCF [W:W], "=AND($V1=""Date"",0<LEN($W1),YEAR($W1)>" & intThisSeason & ")", clrOrange


            ' ==================================================================================
            Case 16 ' Rule
                ApplySort "J"
                ApplyCF [N:O], "=IF(AND(INDIRECT(ADDRESS(ROW()-1,10))=INDIRECT(ADDRESS(ROW(),10)),LEN(ADDRESS(ROW()-1,COLUMN()))>0,INDIRECT(ADDRESS(ROW()-1,COLUMN()))<>INDIRECT(ADDRESS(ROW(),COLUMN()))),TRUE,FALSE)", clrOrange
                ApplyCF [O:O], "=IF(AND($J1=""Price Change Level"",AND($O1<>""5"",$O1<>5)),1,0)", clrRed


            ' ==================================================================================
            Case 17 ' Auth


            ' ==================================================================================
            Case 18 ' Discounts-PH
            ' Hide some stuff
            [A:C].EntireColumn.Hidden = True
            [F:H].EntireColumn.Hidden = True
                
            ' Sort discounts (PHOrder, Discount, Date)
            ApplySort "H", "J", "P"


            ' ==================================================================================
            Case 19 ' Discounts-LOB
            ' Hide some stuff
            [A:C].EntireColumn.Hidden = True
            [F:H].EntireColumn.Hidden = True
                
            ' Sort discounts (PHOrder, Discount, Date)
            ApplySort "H", "K", "Q"


            ' ==================================================================================
            Case 20 ' RuleAltView
                ' Highlight any Price Change Levels that aren't set to 5
                ApplyCF DataRange, "=IF(AND(ROW()>1,COLUMN()>1,INDIRECT(ADDRESS(1,COLUMN()))=""Price Change Level"",INDIRECT(ADDRESS(ROW(),COLUMN()))<>""5""),TRUE,FALSE)", clrRed
                ' Highlight any values that don't match the value above (Rule inconsitencies)
                ApplyCF DataRange, "=IF(AND(ROW()>2,COLUMN()>1,INDIRECT(ADDRESS(ROW()-1,COLUMN()))<>INDIRECT(ADDRESS(ROW(),COLUMN()))),TRUE,FALSE)", clrOrange

                ' Move Customer Min Age left of Max age, for easier comparison
                Dim celMin As Range
                Set celMin = [A1:AZ1].Find("Customer Minimum Age")
                Set celMax = [A1:AZ1].Find("Customer Maximum Age")
                If Not (celMax Is Nothing) Then
                    Columns(celMin.Column).Select
                    Selection.Cut
                    Columns(celMax.Column).Select
                    Selection.Insert Shift:=xlToLeft

                    ' Check Min Age < Max Age (duh)
                    ApplyCF DataRange, "=IF(AND(ROW()>1,COLUMN()>1,INDIRECT(ADDRESS(1,COLUMN()))=""Customer Minimum Age"",INDIRECT(ADDRESS(1,COLUMN()+1))=""Customer Maximum Age"",INT(INDIRECT(ADDRESS(ROW(),COLUMN())))>=INT(INDIRECT(ADDRESS(ROW(),COLUMN()+1)))),TRUE,FALSE)", clrOrange
                End If


            ' ==================================================================================
            Case 21 ' Sale Locations
                ApplyCF DataRange, "=IF(A2=""Y"",TRUE,FALSE)", clrGreen
                ApplyCF DataRange, "=IF(A2=""N"",TRUE,FALSE)", clrRed

                ApplyCF [A1:AZ1], "=A1=""RRW Interface""", clrOrange
                ApplyCF [A1:AZ1], "=A1=""All Inv Pools / bStore Master""", clrMagenta

            ' ==================================================================================
            Case 22 ' Sales Channels
                ApplyCF DataRange(1), "=IF(B2=""Y"",TRUE,FALSE)", clrGreen
                ApplyCF DataRange(1), "=IF(B2=""Y"",FALSE,TRUE)", clrRed


            ' ==================================================================================
            Case 23 ' PH Other


            ' ==================================================================================
            Case Else
                'MsgBox "Sheet " & ws.Name & " is Index " & ws.Index
        End Select

        ' Back to the start.  If you pass go, do not collect $200
        [A1].Select

        ' Change the tab color based on if there's any conditional highlighting
        ChangeTabColorBasedOnValidationColorInSheet (ws.Index)
        ' Run sub-function for per-user custom validation
        QA_Validate_Custom ws.Index

        ' Back to the start again
        [A1].Select
    ' Iterate next sheet
    Next

    ' Reset back to Product sheet after formatting
    Worksheets(1).Activate
End Sub


