# Stock- Analysis

## Purpose ##
  The main Purpose of this Project is to Analyze the entire dataset for the stock market over the last few years using VBA and Excel, I will edit  or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, I’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, a written analysis will be include that explains your findings.

  > Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.
         
## Background ##

  >  Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.


 ## Deliverable 1: Refactor VBA Code and Measure Performance ##
 Looking at the Perfomance for 2017 on the screenshot below, the overall return were better compare with the second screenshot which belong to 2018
 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/112348240/196322020-6e76b843-f604-4d32-acaf-28cdadf4658c.png)  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/112348240/196322288-7719b20c-c1aa-4980-92fe-7941a71e3366.png)

  As you notice above The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module

**Step 1a:**

Create a `tickerIndex` variable and set it equal to zero before iterating over all the rows. You will use this `tickerIndex` to access the correct `index` across the four different arrays you’ll be using: the `tickers` array and the three output `arrays` you’ll create in Step 1b.

'1a) Create a ticker Index
    
    tickerIndex = 0
    
**Step 1b:** Create three output arrays

    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
 **Step 2a:** Create a for loop to initialize the tickerVolumes to zero.

    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
        tickerStartingPrices(i) = 0
        
        tickerEndingPrices(i) = 0
    
    Next i
   
 
 **Step 2b:** Loop over all the rows in the spreadsheet.
 
    For i = 2 To RowCount
    
 **Step 3a:** Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
  **Step 3b:** Check if the current row is the first row with the selected tickerIndex.
        
        'If  Then
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
 **Step 3c:** check if the current row is the last row with the selected ticker
 
        'If  Then
        
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
         End If
         
 **Step 3d:** Increase the tickerIndex.
            
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
             
                tickerIndex = tickerIndex + 1
                
            End If
    
    Next i
 
 **Step 4:** Loop through your arrays to output the (`Ticker`, `tickerVolumes`,`tickerEndingPrices` and `tickerEndingPrices`)
   
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    **'Formatting**
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

  Below, screenshot shows how long did it take to run the code after we edit the original code. Also, you could see that a button was create to `Run Analysis for Assignment` or `Clear Worksheet`
   
 ***Example for 2017***
 
![Refactor Code 2017](https://user-images.githubusercontent.com/112348240/196331915-bfc3725e-8394-49b5-b91e-71d15263197d.png)

 ***Example for 2018***

![Refactor Code 2018](https://user-images.githubusercontent.com/112348240/196331741-0fff8d6a-2ee5-4c5e-9040-cb4c2a6dca4d.png)

  
-Check the total daily volume in the **"DQ Analysis"** worksheet. You should see that DQ traded 107,873,900 shares in 2018. Also, we were able to see that the return for 2018 was 63% negative which it will give Steve a better idea whether expand the Porfolio or not for his parents

![DQ Analysis Screenshot](https://user-images.githubusercontent.com/112348240/196323101-dc92adfd-354b-462a-9b70-0abedff74e65.png)

## Formatting

  The magic of VBA is that we were able to code the formatting to our results, place color base on the numbers as well as set **Bold**, ***cursive*** and so on formatting styles available. I added few multi screenshot to see the changes using those code on the formatting 
  
![image](https://user-images.githubusercontent.com/112348240/196016167-35daa478-3baf-4d3b-bf9c-77b170050668.png)

![image](https://user-images.githubusercontent.com/112348240/196016508-344a0b9e-2a50-49d2-8c15-3d90b3ce571e.png)

![image](https://user-images.githubusercontent.com/112348240/196016817-04c7df62-4701-4c60-853d-fc82c5081448.png)

# Creating a Clear Worksheet Button as well as Run Analysis one

1. Create a button to run the ClearWorksheet() macro. (Hint: Try selecting the button we’ve already created, copying it, and then pasting it.)
Create a button on the DQ worksheet to run the DQAnalysis() macro. ***This part was complete all the information wasa clear then another button was added to bring back the analysis for each sheet***
2. Answer this question: What happens if you create a button to run ClearWorksheet() on the DQ worksheet? ***If you create ClearWorksheet you also will need to create another button to assign the macro relate to the analysis to display the information again***
3. Create a button to run ClearWorksheet() on the DQ worksheet. Does its behavior match what you thought would happen? ***it just clear the worksheet so I was expecting that to happen. I create ClearWorksheet for DQ Analysis***

![image](https://user-images.githubusercontent.com/112348240/196331363-b2515792-cc60-4b03-9e4c-1c21850efc0f.png)

 # Summary
 
 
 *Advantages of refactoring code in general*
                                                  
  >According to Martin Fowler's definition - "Refactoring is a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior."

  >VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source
  
  >Refactoring improves the design of software, makes software easier to understand, helps us find bugs and also helps in executing the program faster. There is an additional benefit of refactoring. It changes the way a developer thinks about the implementation when not refactoring.
disadvantages of refactoring code in general 
  
  *Disadvantages  of refactoring code in general*
  
  >Imprecise refactoring could introduce new bugs and errors into the code
  
  >it's a lack of proper tools. Refactoring nearly always includes renaming variables and methods, changing method signatures, and moving things around. Trying to make all these changes by hand can easily lead to disaster

  >Refactoring process can affect the testing outcomes

 *Advantages of the original and refactored VBA script*
 
 >A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier. 

 *Disadvantages of the original and refactored VBA script*
  
 >Sometimes writing, testing and debugging the script will take longer than using worksheet features. VBA does not adjust in the way that formulae do when you move data from one worksheet to another, insert a column, delete rows, etc.
