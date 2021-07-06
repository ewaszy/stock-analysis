# stock-analysis
Module 2 Files

## Overview of Project

This projects consisted of two parts. First, I created an original subroutine called "All Stocks Analysis" that performed an analysis on the volumes and returns of all stocks for a chosen year (either 2017 or 2018). This original script used a nested for loop that iterated over the dataset multiple times to perform the analysis. For the second part, the original subroutine "All Stocks Analysis" was refactored without a nested for loop to iterate over all rows in the selected year's worksheet only once to perform the same analysis with the same output. 

### Purpose
The purpose of this challenge was to determine if refactoring the original code had an impact on the script run time.  

## Results with Images and Examples of Code

### Original Code for "All Stocks Analysis" Using a Nested For Loop
![Nested_ForLoop_Original_Script](https://user-images.githubusercontent.com/84869167/124545930-7e405680-ddef-11eb-9ab8-a333001b3854.PNG)

### Refactored Code for "All Stocks Analysis" Using a Single For Loop
![Single_Loop_Refactored_Script](https://user-images.githubusercontent.com/84869167/124545947-85676480-ddef-11eb-959c-99ca80e00e7a.PNG)

### Stock Market Performance 2017 vs 2018

In 2017, we can see that all stocks except TERP had a positive return for that year. DQ displayed the highest return percentage amongst all other stocks. 

![VBA_Challenge_2017_Stocks Output](https://user-images.githubusercontent.com/84869167/124538502-5d710480-dde1-11eb-976d-4a932f4474d2.PNG)

In 2018, all stocks except two, ENPH and RUN, experienced a loss in return. ENPH and RUN would be considered good investments in 2018. Both stocks Total daily volumes increased by at least $200,000,000 by the end of 2018.   

![VBA_Challenge_2018_Stocks_Output](https://user-images.githubusercontent.com/84869167/124538675-a88b1780-dde1-11eb-995e-9232f37875f4.PNG)


### Execution Times of Original Script vs Refactored Script

The original script ran the analysis for 2017 in seconds.

![VBA_Challenge_2017_Original_Script_Runtime](https://user-images.githubusercontent.com/84869167/124541300-bb541b00-dde6-11eb-85b5-35cb3840f436.PNG)

The original script ran the analysis for 2018 in seconds. 

![VBA_Challenge_2018_Original_Script_Runtime](https://user-images.githubusercontent.com/84869167/124541313-bee7a200-dde6-11eb-8b87-cb2ee16a34fd.PNG)

The refectored script ran the analysis for 2017 in seconds.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/84869167/124541332-cad36400-dde6-11eb-9243-2d8a81e410a8.PNG)

The refactored script ran the analysis for 2018 in seconds. 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/84869167/124541352-d32b9f00-dde6-11eb-9712-926520d8c573.PNG)


## Summary
### 1.What are the advantages or disadvantages of refactoring code?
**Advantages**

Refactored code eliminates redundant or duplicate loops and subroutines to the same results in less time. Refactored code eliminates using the same line of code in several locations and rearranges the process in a more logical and efficient manner, allowing another programmer to easily work with the code. Some lines of code can be directly copied, others can be reused if modified. Last, refactoring can reveal bugs or errors found in the original script. 
**Disadvantages**

Disadvantages of refactoring code include the potential to introduce logical errors in the structure when copying and modifying code into the new script, affecting the desired outcomes of an analysis. Refactoring may also take longer than simply creating a new script.

### 2.How do these pros and cons apply to refactoring the original VBA script?
Eliminating the nested For Loop reduces the amount of times the script has to loop through the data; thus producing a more logical code that runs faster and more efficiently using code that was copied and/or modified from the original subroutine. 


