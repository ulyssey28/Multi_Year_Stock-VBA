Attribute VB_Name = "Module1"
Sub VBAassignCode()
'Define a variable that will assume the value of each worksheet in turn
Dim ws As Worksheet
'Create a For Each statement to act upon each worksheet
For Each ws In Worksheets
'Begin by sorting the dataset based on the first column (in ascending order)
'To do this we first define the last row of the reference sheet
Dim lastrow As Long
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
ws.Range("A2:G" & lastrow).Sort key1:=ws.Range("A2:A" & lastrow), _
   order1:=xlAscending, Header:=xlNo
'create the columns that will hold the ticker type and the Total stock volume
  ws.Range("I1").Value = "Ticker"
  ws.Range("L1").Value = "Total Stock Volume"
'Next filter through rows to identify when there is a difference in ticker

'If condition of "difference in next cell" is met add new ticker element under Ticker column in order
'--------------------------------------------------------------------------------------------------
'Define a counter that will increase by one integer value every time a new ticker element has been found by our loop
'This counter will be added to the row number element Range("I1").row within a cells funtion, where (Range("I1")+ Counter) will act as our row number
'This will allow us to shift down one row in our Ticker column to put the newest found ticker (from our loop) under the previously found ticker
     Dim Counter As Long
     Counter = 0
     For x = 2 To lastrow
       If ws.Cells(x, 1).Value <> ws.Cells(x + 1, 1).Value Then
         Counter = Counter + 1
         ws.Cells(ws.Range("I1").Row + Counter, 9).Value = ws.Cells(x, 1).Value
       End If
     Next x
     
'---------------------------------------------------------------------------------------------------
'Now set the appropriate summation of volumes next to the appropriate ticker type
'First we use a for loop and conditions to run through each ticker row and assign appropriate volume summation in the adjecent cell
      Dim Totalvol, Counter2 As Long
      Counter2 = 0
      Totalvol = Cells(2, 7).Value
      For i = 2 To lastrow
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            Totalvol = Totalvol + ws.Cells(i + 1, 7).Value
        Else
         Counter2 = Counter2 + 1
         ws.Cells(ws.Range("J1").Row + Counter2, 12) = Totalvol
         Totalvol = ws.Cells(i + 1, 7)
       End If
      Next i
'-----------------------------------------------------------------------------------------------------
'Now we create a yearly change column
ws.Range("J1").Value = "Yearly Change"
'Yearly change will be the difference between the final close value of a ticker type and the first open value of that samr ticker type
'Define two counter which will help us keep track of changes in rows and help place the value in the correct spot
    Dim Yearly_chg As Double
    Dim Counter3, Counter4 As Long
      Counter3 = 0
      Counter4 = 0
      For j = 2 To lastrow
'in the ticker column if consecutive cells have the value then call 1 to the first counter
        
        If ws.Cells(j, 1).Value = ws.Cells(j + 1, 1).Value Then
            Counter3 = Counter3 + 1
            
'Once a difference in ticker value is found add 1 to the second ticker which will help us place the final values
        
        Else
         Counter4 = Counter4 + 1
         
'Define the yearly change variable as the difference between the final closed value of a ticker type and the first open value
'The (j - Counter3) in this case allows us to reference the first row of a newly found ticker

         Yearly_chg = ws.Cells(j, 6).Value - ws.Cells(j - Counter3, 3)
         ws.Cells(ws.Range("K1").Row + Counter4, 10).Value = Yearly_chg
         Counter3 = 0
       End If
       
      Next j
'--------------------------------------------------------------------------------------------------------
'Now we create conditional formatting
'First define last row of yearly change column
      Yearly_chglastrow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
      For y = 2 To Yearly_chglastrow
        If ws.Cells(y, 10).Value > 0 Then
          ws.Cells(y, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(y, 10).Value < 0 Then
          ws.Cells(y, 10).Interior.ColorIndex = 3
        Else
        End If
      Next y
      
'--------------------------------------------------------------------------------------------------------
'First create column header for Percent change
'Define the percent change for the case of starting from zero as a string
'Define the normal percent change as a double
'Define two additonal counters
     ws.Range("K1").Value = "Percent Change"
     Dim Percent_chgzer As String
     Dim Percent_chg As Double
     Dim Counter5, Counter6 As Long
      Counter5 = 0
      Counter6 = 0
     For k = 2 To lastrow
       If ws.Cells(k, 1).Value = ws.Cells(k + 1, 1).Value Then
        Counter5 = Counter5 + 1
       ElseIf (ws.Cells(k, 1).Value <> ws.Cells(k + 1, 1).Value) And ws.Cells(k - Counter5, 3).Value <> 0 Then
         Counter6 = Counter6 + 1
'Percent change is defined as the yearly change value for a ticker divided by the initial open value of a ticker
         Percent_chg = ws.Cells(ws.Range("J1").Row + Counter6, 10).Value / ws.Cells(k - Counter5, 3).Value
         ws.Cells(ws.Range("L1").Row + Counter6, 11).Value = Percent_chg
             ws.Cells(ws.Range("L1").Row + Counter6, 11).NumberFormat = "0.00%"
         Counter5 = 0
'Additional else condition is to avoid error cause by division by zero
       Else
        Counter6 = Counter6 + 1
        Percent_chgzer = "Undefined"
        ws.Cells(ws.Range("L1").Row + Counter6, 11).Value = Percent_chgzer
        Counter5 = 0
       End If
     Next k
'----------------------------------------------------------------------------------------------------------
'Create labels for the value cells
'One for greatest percent increase,one for greatest percent decrease, and one for greatest total volume
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     ws.Range(ws.Range("O2"), ws.Range("O4")).Columns.AutoFit
'Create ticker and value labels
     ws.Range("P1") = "Ticker"
     ws.Range("Q1") = "Value"
'Find the Maximum value within the percent change column
    ws.Range("Q2").Value = Application.WorksheetFunction.Max _
        (Range(ws.Cells(1, 11), ws.Cells(Yearly_chglastrow, 11)))
'Change to percent format
    ws.Range("Q2").NumberFormat = "0.00%"
'Find the Minimum value within the percent change column
    ws.Range("Q3").Value = Application.WorksheetFunction.Min _
        (Range(ws.Cells(1, 11), ws.Cells(Yearly_chglastrow, 11)))
'Change to percent format
    ws.Range("Q3").NumberFormat = "0.00%"
'Find the Maximum value within the total volume column
    ws.Range("Q4").Value = Application.WorksheetFunction.Max _
        (Range(ws.Cells(1, 12), ws.Cells(Yearly_chglastrow, 12)))
'----------------------------------------------------------------------------------------------------------
'Set ticker values
   For t = 2 To Yearly_chglastrow
     If ws.Cells(t, 11).Value = ws.Range("Q2").Value Then
       ws.Range("P2").Value = ws.Cells(t, 9).Value
    Else
    End If
   Next t
   
   For r = 2 To Yearly_chglastrow
     If ws.Cells(r, 11).Value = ws.Range("Q3").Value Then
       ws.Range("P3").Value = ws.Cells(r, 9).Value
    Else
    End If
   Next r
   
    For s = 2 To Yearly_chglastrow
     If ws.Cells(s, 12).Value = ws.Range("Q4").Value Then
       ws.Range("P4").Value = ws.Cells(s, 9).Value
    Else
    End If
   Next s
   ws.Range("Q4").NumberFormat = "General"

'Define last column in order to autofit headers
LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
ws.Range("H1", ws.Cells(1, 12)).Columns.AutoFit
Next ws
End Sub


