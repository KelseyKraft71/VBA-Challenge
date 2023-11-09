# VBA-Challenge
In this challenge, we were tasked with creating VBA code that loops through provided stock data and outputs the ticker, yearly change, percent change, and total stock volume for each unique ticker. 
  After this is done, those values gathered are looped through again to return the greatest percent increase, the greatest percent decrease, and the greatest total stock volume.
  Finally, all of the data is formatted to both look nicer/more legible. There is also conditional formatting for the yearly change and percent change columns.
  
I have in this repository some screenshots of my scripts results applied to the workbook provided as well as a text file containing the VBA code itself.

For some portions of the code, I had to do some research to find a solution for a particular section.


  I found the code for looping through all worksheets at https://stackoverflow.com/questions/21918166/excel-vba-for-each-worksheet-loop<br>
  I could not figure out how to get a For Each loop to work in this context so I had to get some help.

  I found the percentage formatting code at https://stackoverflow.com/questions/45510730/vba-how-to-convert-a-column-to-percentages<br>
  I kept trying to use the round function and I could not get it to work.

  I asked ChatGPT for help with the corresponding ticker line in each of the loops done after the initial loop. <br>
  I initially had the code without this and it just kept returning the last ticker it encountered so I needed some assistance.

  To make everything easier to read, I used an autofit function I found here https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
