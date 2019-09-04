Attribute VB_Name = "Module1"
Sub vba_homework()


'define our variables for tracking the ticker, total, and placement of the summary data
Dim ticker As String
Dim total As Double
Dim summary_row As Integer
Dim summary_row2 As Integer
Dim openval As Double
Dim closeval As Double
Dim yearcng As Double
Dim pctcng As Double
Dim cnt As Integer
Dim maxup As Integer
Dim maxdn As Integer
Dim maxvol As Integer



'loop thru each worksheet in the workbook
For Each ws In Worksheets

    'just a sanity check
    'MsgBox (ws.Name)
    
    
    total = 0
    summary_row = 2
    cnt = 1
    
    'let's add our new columns here, once for each sheet iteration
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'find the range of data to work with
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'inner loop to perform on each worksheet (we start on row 2, 1 is header data)
    For i = 2 To lrow
        
        'check to see if the next row is the same ticker or a new one
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'this must be our closing value
            closeval = ws.Cells(i, 6).Value
            
            'assign the ticker variable to the current ticker value
            ticker = ws.Cells(i, 1).Value
            'add to the running total for this ticker
            total = total + ws.Cells(i, 7).Value
            'calculate our pctcng and yearcng
            yearcng = closeval - openval
            'need to check for divide by 0 to avoid overflow (there is probably a way to do this easily in the forumula but i'm lazy
            If openval <> 0 Then
                pctcng = ((closeval - openval) / openval) '* 100 don't need to do this anymore since formatting handle is
            Else
                pctcng = 0
            End If
            
            'write the ticker and running total to the summary area
            ws.Range("I" & summary_row).Value = ticker                                  'Ticker
            ws.Range("J" & summary_row).Value = yearcng                                 'Yearly Change
            ws.Range("K" & summary_row).Value = pctcng                                  'Percent Change
            ws.Range("L" & summary_row).Value = total                                   'Total Stock Volume
            'handle formatting
            If yearcng > 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
            End If
            
            'format as %
            ws.Range("k" & summary_row).NumberFormat = "0.00%"
    
            'increment our summary_row for each new ticker symbol
            summary_row = summary_row + 1
            
            'reset the total for the next ticker
            total = 0
            cnt = 1
            
        Else
            'set the openval (we're assuming they are sorted in order of date, otherwise we'll need to improve this to do sanity checks on the date val
            If cnt = 1 Then
                openval = ws.Cells(i, 3).Value
            Else
            End If
            
            cnt = cnt + 1
            'we are on the same ticker so let's just keep adding up the totals
            total = total + ws.Cells(i, 7).Value
    
        End If
    
    'move on to the next row
    Next i
    
    

    'identify the index of the max and min values we're looking for
    maxup = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0)
    maxdn = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0)
    maxvol = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)
    
    ws.Range("O1").Value = "Greatest % Increase"
    ws.Range("P1").Value = ws.Cells(maxup, 9).Value
    ws.Range("Q1").Value = ws.Cells(maxup, 11).Value
    
    ws.Range("O2").Value = "Greatest % Decrease"
    ws.Range("P2").Value = ws.Cells(maxdn, 9).Value
    ws.Range("Q2").Value = ws.Cells(maxdn, 11).Value
    
    ws.Range("O3").Value = "Greatest Total Volume"
    ws.Range("P3").Value = ws.Cells(maxvol, 9).Value
    ws.Range("Q3").Value = ws.Cells(maxvol, 12).Value
    
    'format as %
    ws.Range("Q1").NumberFormat = "0.00%"
    ws.Range("Q2").NumberFormat = "0.00%"
'move on to the next worksheet
Next ws

'notify us that we're done processing
MsgBox ("Done with the updates")

End Sub




