Attribute VB_Name = "QA_Validation_Custom"

' =====================================================
' QA_Colors - Fine! Declare your own damn colors!
' =====================================================
Sub QA_Formatting_Custom(Optional intNull As Integer = 0)
    ' Colours
    ' clrColumnHeader = RGB(173, 216, 130)

    ' Font
    fntSize = 8
End Sub


' =====================================================
' QA_Validation_Custom - Add your own personal touches to your QA
' =====================================================
Sub QA_Validate_Custom(intSheet As Integer)
    ' Put things here to apply them to EVERY sheet
    ' CODE HERE

    ' Put things in their respective areas to apply to SPECIFIC sheets
    Select Case intSheet
        ' ==================================================================================
        Case 1 'Product Header
            ' HERE FOLLOWS AN EXAMPLE OF HOW TO REFERENCE THE WORKSHEET:
                'ActiveSheet.Tab.Color = clrRed

        ' ==================================================================================
        Case 2 'Accounting
            ' Filter only non-z prefix components
            ApplyStripZ "K"

        ' ==================================================================================
        Case 3 'Component
            ' Filter only non-z prefix components
            ApplyStripZ "J"
            ApplySort "E", "J"
            
        ' ==================================================================================
        Case 4 'Pricing PS

        ' ==================================================================================
        Case 5: ' Pricing DR

        ' ==================================================================================
        Case 6: ' Access Product
            ' Filter only non-z prefix components
            ApplyStripZ "J"
            ' Sort columns
            ApplySort "J"

        ' ==================================================================================
        Case 7: ' Access Rule
            ' Filter only non-z prefix components
            ApplyStripZ "J"

        ' ==================================================================================
        Case 8: ' Output
            ' Filter only non-z prefix components
            ApplyStripZ "K"

        ' ==================================================================================
        Case 9: ' Inventory Pools

        ' ==================================================================================
        Case 10: ' Private

        ' ==================================================================================
        Case 11 ' Tax
            ' Filter only non-z prefix components
            ApplyStripZ "J"

        ' ==================================================================================
        Case 12 'Rezo

        ' ==================================================================================
        Case 13: ' Rest

        ' ==================================================================================
        Case 14: ' Cal Def

        ' ==================================================================================
        Case 15: ' Voucher

        ' ==================================================================================
        Case 16 ' Rule

        ' ==================================================================================
        Case 17 ' Auth
            ' Change colour of tab to (mostly) ignore
            ActiveSheet.Tab.Color = clrBlack

        ' ==================================================================================
        Case 18: ' Discounts PH

        ' ==================================================================================
        Case 19 ' Discounts LOB

        ' ==================================================================================
        Case 20 ' Rules (Alt View)

        ' ==================================================================================
        Case 21: ' Sales Locations

        ' ==================================================================================
        Case 22 ' Sales Channels

        ' ==================================================================================
        Case 23: ' Other PH Details

        ' ==================================================================================
        Case Else
            ' Community chest!
    End Select
End Sub




