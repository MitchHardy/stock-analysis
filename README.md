# Stock-Analysis with VBA

# Overview of Project

### In this project, we edited and refactored the Stock Market dataset by using VBA solution code. The purpose of this was to loop through the data one time in order to collect an entire dataset. We then determined whether refactoring the code made the VBA script run faster and more efficient. 

## Results: 

Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1a. The tickerIndex is set equal to zero before iterating over all the rows. This will be used to access the index in the arrays. 

name-of-you-image

1b. Output arrays are created for tickerVolumes, tickerStartingPrices, and tickerEndingPrices. We then put the tickerVolumes array to Long data type. For the tickerStartingPrices and tickerEndingPrices arrays we put them as a Single data type.

name-of-you-image

2a. Created a for loop to initialize the tickerVolumes to zero. Then made it so if the next row’s ticker doesn’t match then it will increase the tickerIndex.

name-of-you-image

2b. A loop is created for all the rows in the spreadsheet. 

name of you image

3a. A script was created inside the loop that increases the current tickerVolumes variable and adds the ticker volume for the current stock ticker.

3b. To check if the current row is the first row with the selected tickerIndex we created an if-then statement. 

3c. We then write a script if it the statement above is true, and then we assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.

3d. After this we'll write a script that will increas the tickerIndex if the following row's ticker doesn't match. 

name-of-you-image

4. The last step we use a for loop to run through all of the arrays so we can then return "Ticker", "Total Daily Volume" and "Return" in the spreadsheet columns. To more easily determine which stocks did well, we will make positive return green and negative return red.

name-of-you-image

Dataset Examples Provided

name-of-you-image

## Summary:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

From what I gathered doing this challenge is that it's important to focus on the all of the small steps in your code. The more specific you are better your output will be in the end. Small

Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
A complex unstructured code is usually best to split in several functions.
Refactoring process can affect the testing outcomes.

Advantages:

Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

2. How do these pros and cons apply to refactoring the original VBA script?

Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code.

We need to consider the code refactoring process as cleaning up the orderly house. Unnecessary clutter in a home can create a chaotic and stressful environment. - The same goes for written code.

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.
