# Grocery List Generator

This is a program that generates a grocery list in Chinese and allows the list to be printed.

It was written in C# using Visual Studio 2017, and tested in Windows 8.1 platform.

# Purpose

Manually writing a grocery list several times a week while trying to maintain some variety in one's diet is a tedious task.
This program seeks to alleviate the memory work required to complete the task.

# Design Choices

- The generated list is in Chinese because the person in charge of grocery shopping can only read Chinese.
- The program is made to allow changes to the item pool. For this reason, an Excel spreadsheet was chosen because it is an 
accessible file type for non-technical users.

# How to Use

No installation is required. 

1. Run `GroceryListGenerator2.exe` in the folder `GroceryListGenerator2\bin\Release`.
2. Tick the items to add, and specify the quantity of each. There are different tabs for different categories of items.
3. Click the `Generate` button. The list of items will be shown in English and Chinese.
4. If the list is correct, click the `Save to PDF` button. A Print Preview will shown at the right side of the program.
5. Click `Print` to print the PDF file. 
6. The `Save Print Data` button saves the number of lines printed, to figure out how many grocery lists one set of ink cartridge
can print.

# Versions

**[Current]**<br/>
Version 2.0: This is the current version of the program.

**[Previous]**<br/>
Version 1.0: This program was based on a different idea. The user chooses the number of items to include in the list 
for each category of items. 
The program generates a list by randomly selecting items from a local database for each category.

# License

Copyright Ⓒ Edward Chong 2019. All Rights Reserved.
