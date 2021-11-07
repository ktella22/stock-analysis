# Overview of the Project
## Purpose 

The purpose of this project is to refactor the Microsoft VBA existing script to gather bigger informations based on 2017 and 2018 stockmarket database to forseen whether or not it is worth to investing. In this project, refactoring is done on the existing VBA script and by doing it to run the script faster and more effiect with fewer steps.

# Results
## Analysis

Prior to refactoring my code, I started copying the code that was provided in Challenge starter code file. Using the code in the starter file, I selected the correct workspreadsheet to create the required headers that included years of the analysis, ticker array, Total daily volumn and the calculated Retrun values. After that, I followed the steps to add the codes in the required area. Below is the screenshot of the my refactored script:

<img width="840" alt="Screenshot Challenge" src="https://user-images.githubusercontent.com/92502292/140599650-c745077d-9558-4447-8d17-b8a558c9155a.PNG">


# Summary

### The advantages and disadvantages of refactoring code 

- In the modified code, my code seems to be cleaner and more organized. When I run the code, it is more efficient and the return values came out a lot faster. There is a significant improvent in getting the clean data, less bugging in the code and software improvement. The downside of the refactoring code is that we don't have the benefit of having the shorter code in the modified version. Some of the codes are still the same and long as original script.

### The advantages and disadvantages of the original and refactored VBA script

- Refactoring definitely shorten the time of the analysis run and it is a major advantage. The original VBA scipt took over one second to run the code, but the refactored script took around 0.23 seconds to run the code for each stockmarket year. The refactored version reduced the running time by 75% from the orginal script. However, there are three for loops that was used to run the code in refactored VBA script instead of using nested loop in original script. With three for loops by using same "i" iterators can be quite confusing and it can easily be misplaced the closing loop in the wrong line.