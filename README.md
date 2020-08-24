# VBA-challenge
Wall Street Data Challenge

WallStreet.vbs contains the GetTicker Macro, which goes through all the rows on one sheet and then displays the data for yearly change, percent change, and total stock volume for each ticker value.

Additionally, it will also check to see if the yearly change is positive or negative, and color code the background depending on the value. Red for negative and green for positive.

Greatest.vbs contains the code to get and display the greatest % increase, greastest % decrease, and greatest volume of all the ticker values.

MultiSheet.vbs combines the code from the previous two files and calculates the data for every worksheet in a file.

There were some modifications to MultiSheet.vbs from the single sheet versions:
1. When calculating the percent change, a condition was added to check if the value at the beginning of the year was 0 so it wouldn't result in an error in the calculation and the program ending.
2. Some values had to be reset when calculating the Greatest values so it wouldn't print out data from an earlier worksheet. This reset wasn't necessary when working on a single worksheet.
