Attribute VB_Name = "Module1"
Sub Stock_Tracker():

For Each ws In Worksheets
    'variable declarations
    Dim ticker As String
    Dim Volume As Double
    Dim Opening As Double
    Dim Closing As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table As Integer
    Summary_Table = 2
    Volume = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'x l(this is a lowercase L) Up

    'Summary Table construction
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Volume"
        
    'Loop through rows to LastRow
    For i = 2 To LastRow

        'Capture Opening value for each new Ticker
        If Volume = 0 Then
            Opening = ws.Cells(i, 3).Value
        
        End If
        'Check that values are different in order to continue adding to the total
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
            'Set Ticker value
            ticker = ws.Cells(i, 1).Value
        
            'Set Opening and Closing value
        
            Closing = ws.Cells(i, 6).Value

            yearly_change = Closing - Opening
            percent_change = (Closing - Opening) / Opening
        
            'Add volume to existing volume
            Volume = Volume + ws.Cells(i, 7)
        
            'Summary_Table fullfillment
            ws.Range("J" & Summary_Table).Value = ticker
            ws.Range("K" & Summary_Table).Value = yearly_change
            ws.Range("L" & Summary_Table).Value = percent_change
            ws.Range("M" & Summary_Table).Value = Volume
        
            'Iterate on Summary_Table
            Summary_Table = Summary_Table + 1
        
            'Reset Volume
            Volume = 0
        
        Else
        'If values are the same, add to the volume
            Volume = Volume + ws.Cells(i, 7).Value
        
        
        End If
    
    'iterate For loop
    Next i
    
    'formatting Summary_Table correctly
    For a = 2 To LastRow
        ws.Cells(a, 11).Style = "Currency"
            
            'Format for color, green for positive change
            If ws.Cells(a, 11).Value > 0 Then
                ws.Cells(a, 11).Interior.ColorIndex = 4
            'red for negative change
            Else
                ws.Cells(a, 11).Interior.ColorIndex = 3
            End If
        
    Next a
    
    For a = 2 To LastRow
        ws.Cells(a, 12).NumberFormat = "0.00%"
    Next a
'next worksheet in file
Next ws

End Sub
