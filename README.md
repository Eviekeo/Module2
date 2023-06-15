# Module2
As part of my submission for the Module 2 - VBA Challenge, I have uploaded the following files
  - screenshot of the 2018 tab in Excel  
  - screenshot of the 2019 tab in Excel  
  - screenshot of the 2020 tab in Excel
  - VBA script of code  

# Code Source
The following code was provided as guidance from the bootcamp tutor, for the purpose of locating and storing the opening price of each stock, which is the open price at the start of each year. This was used to calculate the yearly change of each stock.
       
        'set a variable to hold the first row of each stock
        Dim openprice_row As Double
        openprice_row = 2
        
        'below code was used in the for loop, to find the opening value of each stock (located at the open price row, and in column 3).
        openvalue = ws.Cells(openprice_row, 3).Value
        
        'as the code goes through each iteration, the first price of each subsequent stock after row 2 will be equal to 'i=2'
        openprice_row = i + 1
