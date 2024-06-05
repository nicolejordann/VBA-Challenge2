# VBA-Challenge2
This is the README for my VBA-Challenge2 repository 

Sub please_work():

Dim lastrow As Long
Dim sumticker As Double
Dim Ticker As Double
Dim qchange As Double
Dim closeprice As Double
Dim openprice As Double
Dim pchange As Double
Dim totalvol As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestticker As String
Dim lowestticker As String
Dim greatestvol As Double
Dim greatvolticker As String
Dim ws As Worksheet


'To loop through all worksheets in this woorkbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate


totalvol = 0
Ticker = 1
sumticker = 1
qchange = 2
pchange = 2



' Headers for summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"


' Calling the last row in column A
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


'Loop To find the next ticker
  For Ticker = 2 To lastrow + 1
        
        
      If Cells(Ticker, 1).Value <> Cells(sumticker, 9).Value Then
                
                      'To allow us to collect first ticker name before running the calculations
                     If Ticker <> 2 Then
                    
                          'quartly change for previous ticker
                          Cells(qchange, 10).Value = closeprice - openprice
                             qchange = qchange + 1
             
                          'percent change for previous ticker
                           Cells(pchange, 11).Value = (closeprice - openprice) / openprice
                    
                          'display total volume for previous ticker
                          Cells(pchange, 12).Value = totalvol
                           pchange = pchange + 1
                           totalvol = 0
                           
                        
                    End If
                 
             
                     'find openprice (these are outside the loop so that they dont change when the ticker changes)
                    openprice = Cells(Ticker, 3).Value
            
            
            
                    ' Summary ticker name
                 Cells(sumticker + 1, 9).Value = Cells(Ticker, 1).Value
                 sumticker = sumticker + 1
            
         End If
         
         
      'find close price (this will change each row until ticker is changed)
       closeprice = Cells(Ticker, 6).Value
       
      'collect volume from each row and add it to itself
       totalvol = Cells(Ticker, 7).Value + totalvol
    
        
         
    Next Ticker
    
    
   'reset variables
   pchange = 2
   greatestincrease = 0
   greatestdecrease = 0
   greatestvol = 0
   
   
   ' Calling the last row in column I
lastrow = Cells(Rows.Count, 9).End(xlUp).Row


For pchange = 2 To lastrow


'loop to calculate highlights table
         If greatestincrease < Cells(pchange, 11).Value Then
         
         'obtain greatest % increase info
         greatestincrease = Cells(pchange, 11).Value
         greatestticker = Cells(pchange, 9).Value
            
        End If
        
        'obtain greatest % decrease info
        If greatestdecrease > Cells(pchange, 11).Value Then
            
            greatestdecrease = Cells(pchange, 11).Value
            lowestticker = Cells(pchange, 9).Value
            
        End If
        
        'obtain greatest volume info
        If greatestvol < Cells(pchange, 12).Value Then
        
            greatestvol = Cells(pchange, 12).Value
            greatvolticker = Cells(pchange, 9).Value
            
        End If
        
    Next pchange
    
    'inserting  highlights info into correct cells
    Cells(2, 17).Value = greatestincrease
    Cells(2, 16).Value = greatestticker
    
    Cells(3, 17).Value = greatestdecrease
    Cells(3, 16).Value = lowestticker
    
    Cells(4, 17).Value = greatestvol
    Cells(4, 16).Value = greatvolticker


'reset variables
qchange = 2

'Conditional formatting to Yearly Change column
          For qchange = 2 To lastrow
          
                If Cells(qchange, 10).Value > 0 Then
                        Cells(qchange, 10).Interior.Color = RGB(0, 255, 0)
                    
                    ElseIf Cells(qchange, 10).Value < 0 Then
                        Cells(qchange, 10).Interior.Color = RGB(255, 0, 0)
                
                End If
                
            Next qchange
            
    Next ws
    
End Sub


