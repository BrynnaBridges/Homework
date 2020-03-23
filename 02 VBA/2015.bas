Attribute VB_Name = "Module2"
Sub Multiple_year_stock_data()

' Setting the variables
    'Setting the variable for Ticker Symbol
        Dim ticker As String
        
    'Setting Initial Value of each ticker
        Dim Initial As Double
        Initial = 0
        
    'Setting Closing Value of each ticker
        Dim Closing As Double
        Closing = 0
 
   'Setting the variable for Total Stock Volume
        Dim Total_Stock As Double
        Total_Stock = 0
        
    'Setting the variable for Yearly Change
        Dim Yearly_Change As Double

        
    'Settting the Variable for Percent Change
        Dim Percent_Change As Double
        Percent_Change = 0

' Location for Ticker symbol in Excel Sheet
        Dim Answer_Row As Integer
        Answer_Row = 2
        

' Loop through all Ticker Symbols
   For i = 2 To 705714

    ' Check first row for different ticker symbols
             If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
      ' Set the Ticker Symbols
            ticker = Cells(i, 1).Value
    
      ' Add to the Total Stock
            Total_Stock = Total_Stock + Cells(i, 7).Value
            

      ' Print the Ticker Symbol in Column I
            Range("I" & Answer_Row).Value = ticker

      ' Print the Total Stock to Column J
            Range("J" & Answer_Row).Value = Total_Stock
                      
            
      ' Moving Down one Row
            Answer_Row = Answer_Row + 1
      
      ' Total Stock needs to start at 0
            Total_Stock = 0
             

    ' If next Row is same Ticker Symbol....
    
     Else

      ' Add to the Total Stock Volume Column
      Total_Stock = Total_Stock + Cells(i, 7).Value
        
 ' Check Second Row for first date of ticker
            If Cells(i, 2).Value = "20150101" Then
            
     ' Set the initial Value
            Initial = Cells(i, 3).Value
            
 'Print the Initial Value
            Range("M" & Answer_Row).Value = Initial
            
  End If
  End If
  
'Check Second Row for last date of ticker
        If Cells(i, 2).Value = "20151231" Then
        
        'Set the Closing Value
            Closing = Cells(i, 6).Value
            
         'Print the Closing Value
            Range("N" & "2").Value = Closing

            
'Finding the change in the year
    Yearly_Change = (Closing - Initial)
    
'Print the change in the year
    Range("K" & "2").Value = Yearly_Change
        

 End If
 
             If Initial <> 0 Then
            Percent_Change = Yearly_Change / Initial
    
'Print the Percent Change
    Range("L" & "2").Value = Percent_Change
    
    End If
    
  'Conditional Formatting
 If Cells(i, 12) < 0 Then
    Cells(i, 12).Interior.ColorIndex = 3
Else
    Cells(i, 12).Interior.ColorIndex = 10
 
 End If
    
 Next i
 
 End Sub
