# stock-analysis
DU Data Analytics Bootcamp Module 2

## Overview of Project
This VBA project serves users by automating stocks analysis for a given data set included in the VBA_Challenge file. The automation enhances user ability to make investment decisions by giving users a view of historical stock performance during years provided.

### Purpose
The purpose of the macro programmed into the spreadsheet is to make better use of an analyst's time be creating an efficient table to view summarized information on the population of stocks in the dataset. Historical information was provided for years 2017 and 2018, with each year comprising approximately 3,000 lines of data. In order to analyze the data, a summary table based on some chosen key metrics enables the user to get much more clear and concise information and much more rapidly.

## Analysis and Challenges
The challenge of VBA development in this module was improving on some of the more basic Excel knowledge to automate routine table building. The specific focus of much of the module was to create efficient code that would summarize the data while not consuming unnecessary memory. In order to accomplish this, the challenge really became to go through the logic of the intial code developed in the module and try to refine it in the challenge in order to output the same results, but with better performance. A start and end timer was implemented into the code to measure results of performance.

### Analysis of Outcomes Based on Coding Improvements
<TABLE align="center" CELLSPACING="20">
<TR>
<TD><p align="center">
    <img src="https://github.com/cb19weber/stock-analysis/blob/main/resources/Module_Green_Stocks_2017.png" />
    </p></TD>
<TD><p align="center">
    <img src="https://github.com/cb19weber/stock-analysis/blob/main/resources/VBA_Challenge_2017.png" />
    </p></TD>
</TR>
</TABLE>
The left image above displays the time consumed to run the code from the initial module activity, while the code on the right displays the time after code refactoring for performance improvements. Looping the FOR statements and creation of the tickerIndex to reduce the amount of iterations improved code performance by just over 84% for the 2017 data!
<p></p>
<TABLE align="center" CELLSPACING="20">
<TR>
<TD><p align="center">
    <img src="https://github.com/cb19weber/stock-analysis/blob/main/resources/Module_Green_Stocks_2018.png" />
    </p></TD>
<TD><p align="center">
    <img src="https://github.com/cb19weber/stock-analysis/blob/main/resources/VBA_Challenge_2018.png" />
    </p></TD>
</TR>
</TABLE>
The second set of images display the time variance in the coding for iterating through the 2018 dataset. The improvement in code performance demonstrated a similar improvement of just under 82%.

### Challenges and Difficulties Encountered
One of the challenges I had in developing the VBA code was an additional desire to refactor the challenge code (and actually the module code as well) to make the VBA array of tickers a bit more dynamic. My desire was to create a code that would grab all <i>unique</i> tickers from the total dataset and then populate a variable sized array out of those values. I was able to grab unique values by copying the ticker column and then removing duplicates. I was additionally able to populate the array using the unique values. Where I ran into trouble is trying to declare a variable sized array. I attempted to create a variable that would serve as an integer to replace the constant in the declaration, but was given a compile error. The only way I was able to get around this was to simply declare the VBA array for a larger potential number of tickers than I knew I needed.
There are also some limitations in the challenge code that I wanted to overcome. One of those limitations was that the initial challenge code assumes and requires that the dataset be sorted a certain way and also must iterate through the dataset to gather the starting and ending price for each stock. I wanted to make the code a little more robust so that how the dataset was sorted was irrelevant. In order to accomplish this I had to conver the date column to something that could be used as a reference and then develop and excel formula to search through the data.

## Results
I am quite please with the coding performance improvements acheived through better logic, and that is definitely something I want to use in future coding projects. I am also pleased that I was at least able to build in some functionality improvements into the challenge that overcome some of the limitations discussed above. By implementing the XLOOKUP formula with an array based search using minimum and maximum date values, I was able to find the oldest and newest starting price and I was no longer dependent on how the data was sorted.