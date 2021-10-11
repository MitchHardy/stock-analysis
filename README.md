# Stock-Analysis with VBA

# Overview of Project

### In this project, we edited and refactored the Stock Market dataset by using VBA solution code. The purpose of this was to loop through the data one time in order to collect an entire dataset. We then determined whether refactoring the code made the VBA script run faster and more efficient. 

## Results: 

Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1a. The tickerIndex is set equal to zero before iterating over all the rows. This will be used to access the index in the arrays. 

1b. Output arrays are created for tickerVolumes, tickerStartingPrices, and tickerEndingPrices. We then put the tickerVolumes array to Long data type. For the tickerStartingPrices and tickerEndingPrices arrays we put them as a Single data type.

2a. Created a for loop to initialize the tickerVolumes to zero. Then made it so if the next row’s ticker doesn’t match then it will increase the tickerIndex.

2b. A loop is created for all the rows in the spreadsheet. 

3a. A script was created inside the loop that increases the current tickerVolumes variable and adds the ticker volume for the current stock ticker.

3b. To check if the current row is the first row with the selected tickerIndex we created an if-then statement. 

3c. We then write a script if it the statement above is true, and then we assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.

3d. After this we'll write a script that will increas the tickerIndex if the following row's ticker doesn't match. 

4. The last step we use a for loop to run through all of the arrays so we can then return "Ticker", "Total Daily Volume" and "Return" in the spreadsheet columns. To more easily determine which stocks did well, we will make positive return green and negative return red.

## Summary:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

From what I gathered doing this challenge is that it's important to focus on the all of the small steps in your code. The more specific you are better your output will be in the end. Small

Disadvantages:


Advantages:


2. How do these pros and cons apply to refactoring the original VBA script?

