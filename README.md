# Stock-Analysis

### Overview
  In this analysis, we took a look at the overall stock market dataset to find the stocks that have the best return rate for safe investing. We also refactored old code to make the script run even better than before. 
  
### Results


#### Charts
![Chart2017](https://user-images.githubusercontent.com/119345840/207919495-afa28e17-56de-472c-829f-043efbf9c092.PNG)

![Chart2018](https://user-images.githubusercontent.com/119345840/207919575-3e39a3d2-b7ab-4108-b703-626270568243.PNG)

  For the years of 2017 and 2018, the stock market had very extreme differences. In 2017, the market florished with huge returns up to a staggering 199.4% (DQ). 2018, however, was the exact opposite. The market seemingly crashed, with only the companies of RUN and ENPH wielding postive results. Similarly, our code also had differences as well. Our origin code used to loop through and scan 11 different times to output the results, while our refactored code only loops through and scans once. This still outputs the same data, but requires significantly less processing. Because of this change, our codes had very different run times. For our origin code, the run time was 0.718 seconds and 0.734 seconds for 2017 and 2018 respectively. Our updated code ran significantly faster; having times of 0.308 seconds and 0.207 seconds for 2017 and 2018. 

#### Code Examples 

##### Origin

loops on the ticker index, scanning the file each time

For i = 0 To 11
    ...
       For j = 2 To RowCount
	...
       Next 



##### Repurposed

loops only once, tracking the ticker index by comparing previous and current ticker value.

For v = 2 To RowCount
    ...
    If Cells(v + 1, 1).Value <> Cells(v, 1).Value Then
       tickerIndex = tickerIndex + 1
    End If
Next

## Summary
