# # Stock Market Analysis with Excel Macros in VBA
## Overview of Project
Using excel and it's macros function to gain an understanding of stock performances by year based on total trade volume and overall positive or negative performance reported by individual stocks.  This project also is a teaching methodology to understand code (VBA) refactoring and how it can improve overall appearance and performance.

## Results, Analysis and Challenges
The stock funds dataset included two years of stock performance for only 12 publicly traded stocks.  The two methods for determining *best* stock performance were total number of trades explained here as 'Total Volume', and overall growth in by percentage from the beginning of the year to the end of the year.  I wrote the code first for a specific stock "DQ" to understand macros performace and syntax.  Then took it a step further and wrote the code to 

### Refactored Final Code

### Refactored Time Stamp of 2017 data points
![VBA_Challenge_2017](https://user-images.githubusercontent.com/102183530/163653581-17fb0e66-259c-4837-ae28-24e572cfd0fd.png)

- *Figure 1. 2017 stock performance by total volume reported and gains or losses recorded.*

### Refactored Time Stamp of 2018 data points
![VBA_Challenge_2018](https://user-images.githubusercontent.com/102183530/163653509-392493c0-5f2a-437e-b81d-fd23bc10f9eb.png)

- *Figure 2. 2018 stock performance by total volume reported and gains or losses recorded.*

## Summary

- What are the advantages and disadvantages of refactoring code in general?

These are the main advantages for refactoring in general, having a more efficient code, a code that is more useful in a broader sense, and making code easier to read.  You are in essence making the code more accurate, while still usable.  For example, when this code was first written there was no defining or labelling of some of the values, which left VBA to give its "best guess" or just assume all the values were *Long* data types taking up more memory than these values needed; thus slowing down the code performance.  Added 'comments' makes the code easier to read, as well as explaining intent of that portion of the code.  Should another coder take a look after you, they can easily understand where you were going with your code.  

Some disadvantages of refactoring could be with more comments the code itself gets longer creating more opportunity for mis-steps.  Another possible disadvantage is that the orginal coder may not know a way to refactor a part of their code in a more efficient manor.  This is a key in working together with others, and seeking out different ways of syntax and answers, but still might remail ellusive for that particular code.  A final potential downfall of refactoring is simply time.  Depending on what the budget for the code looks like or the time parameters, this might be prohibitive, even after all its' advantages are so pronounced.

- What are the advantages and disadvantages of the original VBA Stock Analysis and the Refactored script?

These happen to be similar to some of the general advantages and disadvantages listed above.  Namely, the code is easier to read, the comments help the code look neat, explain what each line is expecting and producing, and it definitely decreases overall run time, proving the refactoring is more efficient.

The disadvantages here are also mirroring the ones above, it took a longer amount of time to refactor some of this code based simply on my original lack of knowledge.  Group sourcing solutions and 'plug-and-play' were key to figuring out the correct formulas to run this efficient code, but I still needed to expend the time to figure out which sources would work to the best advantage of this code.  As I was not under budget constraints, this seems to not apply as much for this type of situation.

