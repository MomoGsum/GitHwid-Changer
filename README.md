A small automation project that I develop to automate data cleaning for Taiwan textile company.

There will be raw data consisting of inventory release coming every year and I will save it into JOIN directory which then the script on SCRIPT directory will run and clean (read an excel file from specified directory, renaming column and value for consistencies, fixing typo, filling NaN with necessary value, dropping duplicates, grouping value into categories, pivoting dataset and transforming from wide into long data for easy analysis) and finally merging multiple yearly data into one for easy analysis and exporting it into .csv long data format.

This project takes few hours to build and could save me numerous time in the future so I could be more productive doing other things.

Hope my code could be of any help.
