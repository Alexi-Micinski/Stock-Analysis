# VBA of Wall Street

## Overview of Project:
Steve’s parents are passionate about green energy, especially as fossil fuels become more scarce and energy demands shift towards renewable energy sources such as hydroelectricity, geothermal energy, wind energy, and bioenergy. They have only invested in one renewable energy company, DQ. Steve is concerned that their investments aren’t more diversified. Steve has started researching other companies in addition to DQ and put together the stock data.

Steve is interested in the total daily volume, or total number of shares traded in a day, and yearly return, or percentage difference from the beginning to year end, for the stocks he’s investigating.

First, Steve wanted to investigate the performance of DQ in 2018. Performance was measured by calculating the yearly return: percentage increase or decrease throughout the year. We found that Daqo’s stocks fell 63% in 2018, so they may not be the best option for Steve’s parents.

<img width="246" alt="Screen Shot 2022-07-26 at 8 08 40 PM" src="https://user-images.githubusercontent.com/106785377/181145387-7ec55b65-7314-4512-9e56-f0f3d8debfd6.png">

Other stocks were analyzed since Daqo didn’t seem like the best fit for Steve’s parents. The work that was used to create the DQ analysis was repurposed to analyze the other stocks for any year. Formatting was added to create an easier to read output, and a script was added to see how fast VBA could execute the code.

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

#### 2017 Results
<img width="225" alt="2017" src="https://user-images.githubusercontent.com/106785377/181146528-f98cf48d-9950-4dfd-920f-aaac6fc06b15.png">

#### 2018 Results
<img width="225" alt="2018" src="https://user-images.githubusercontent.com/106785377/181146540-eeeebc44-f9b8-454c-bf6c-2e14c86ea1b2.png">

The green energy stocks had better rates of return in 2017 compared to 2018 for all stocks except for ENPH and RUN. This is easy to pick out with the color indications. These indications were created using the following code.

```
dataRowStart = 4
    dataRowEnd = 15

    For x = dataRowStart To dataRowEnd
        
        If Cells(x, 3) > 0 Then
            
            Cells(x, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(x, 3).Interior.Color = vbRed
            
        End If
        
    Next x
```

Other colors can be created using [this](https://analysistabs.com/excel-vba/colorindex/) chart.

The total daily volume went up from 2017 to 2018 for some stocks but down for others. Total daily value is not the best way determine stock performance over time.

## Summary: In a summary statement, address the following questions.


### What are the advantages or disadvantages of refactoring code?


### How do these pros and cons apply to refactoring the original VBA script?

