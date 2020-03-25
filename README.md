# ExcelFindAndReplace

Description: automatically search excel files in directory and subdirectory which contain the specific text, and automatically replace the specific text in all excel files with another text which is specified by user.

Background: our company mainly uses excel file to record staff information and report to Airport Authority in daily basis. However, due to the division of works, there are many spelling mistakes in staff name across multiple documents. Because opening every excel file to check whether the spelling is correct and rectify it is time consuming, I then wrote this Java program to improve work efficiency.

Technologies used:
Java Swing, multi-threading

Library: Apache POI, Apache Common IO

Data Structure: Tree

How it works:
The program searches all excel files in the specified location with the specified text. User can also specify whether the search includes subdirectories or not, and also specify whether the operation is case sensitive, exact cell match, trim find text and trim replace text.

The program can simply use Apache POI function XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook) to update all existing formula cells in excel workbook after it replaced text in certain cells. However, the performance is very low if the program handles a very large excel file (e.g. an excel workbook including thousands of excel formula). Therefore, the program provides another way: update all dependent formula cells.

The program uses data structure: Tree to record dependencies for every matched cell it finds. The dependent cells are stored in Tree which is suitable to represent the nature of dependencies (dependent cells and sub-dependent cells). The program then traverses the tree and updates the formula cells recorded in every tree node.

The search and update process runs on the separated thread to avoid freeze of user interface.
