### Google Script App for Sheets to Count # of Cells in a Range with specific Background Colour
version 0.1: 25th August, 2021

Google Sheets offers hundreds of built-in functions like [AVERAGE](https://support.google.com/drive/answer/3093615 "AVERAGE documentation page"), [SUM](https://support.google.com/drive/answer/3093669 "SUM documentation page"), and [VLOOKUP](https://support.google.com/drive/answer/3093318 "VLOOKUP documentation page"). But there isn't a built in function to analyse cells by their formatting - specifially the cell background colour.

This Google App Script is designed to fill this gap, by creating two ["custom functions"](https://developers.google.com/apps-script/guides/sheets/functions "") called countColours() and sumColours().

## countColours()
### Sample usage
E.g. =countColours("A1:A10","Red")

### Syntax
The function takes two arguments:
1. The range that you want to analyse. This can be any range in the current spreadsheet, written using the standard syntax. Examples include dataSheet1!A1:A10 to reference cells A1 through A10 on the dataSheet1 sheet, or just A1:A10 if the function is on the dataSheet1 sheet, to refer to the same region.
2. The second argument is flexible, and can take one of three forms:
  - a cell reference: as with the first argument, this is given in the standard syntax and can point to a cell on the current sheet (B1) or on a different sheet in the same spreadsheet (dataSheet1!B1). The function will use the background colour of this cell to perform the calculations.
  - a colour name: you can enter any of the standard HTML colour names available [here](https://www.w3schools.com/colors/colors_names.asp "Colour names from w3school.com"). It is not case sensitive, so examples include: "firebrick", "Plum", "SIENNA".
  - a HEX colour value: you can enter any of the standard HTML colour HEX values [here](https://www.w3schools.com/colors/colors_hex.asp "Colour HEX values from w3school.com"). Again this is case insensitive, so examples include: "#b22222", "#dda0dd", "#a0522d" (these match the three colour names above)

### Function Result
The function gives one of three results:
1. the number of cells in the given range with the background colour given as the second argument
2. "Check the 1st value, it doesn't seem to be a valid range." - if the first argument is not a valid range reference, this error message is displayed
3. "Check the 2nd value, it doesn't seem to be a valid colour or cell reference." - if the script cannot determine a background colour from the second argument, then it returns this error message.

### Examples
See end of this readme page

## sumColours()
### Sample usage
E.g. =sumColours("A1:A10","Red")

### Syntax
The function takes two arguments:
1. The range that you want to analyse. This can be any range in the current spreadsheet, written using the standard syntax. Examples include dataSheet1!A1:A10 to reference cells A1 through A10 on the dataSheet1 sheet, or just A1:A10 if the function is on the dataSheet1 sheet, to refer to the same region.
2. The second argument is flexible, and can take one of three forms:
  - a cell reference: as with the first argument, this is given in the standard syntax and can point to a cell on the current sheet (B1) or on a different sheet in the same spreadsheet (dataSheet1!B1). The function will use the background colour of this cell to perform the calculations.
  - a colour name: you can enter any of the standard HTML colour names available [here](https://www.w3schools.com/colors/colors_names.asp "Colour names from w3school.com"). It is not case sensitive, so examples include: "firebrick", "Plum", "SIENNA".
  - a HEX colour value: you can enter any of the standard HTML colour HEX values [here](https://www.w3schools.com/colors/colors_hex.asp "Colour HEX values from w3school.com"). Again this is case insensitive, so examples include: "#b22222", "#dda0dd", "#a0522d" (these match the three colour names above)

### Function Result
The function gives one of three results:
1. the sum of all the numeric cells in the given range with the background colour given as the second argument
2. "Check the 1st value, it doesn't seem to be a valid range." - if the first argument is not a valid range reference, this error message is displayed
3. "Check the 2nd value, it doesn't seem to be a valid colour or cell reference." - if the script cannot determine a background colour from the second argument, then it returns this error message.

## Examples
![Image of Examples](https://lh5.googleusercontent.com/-QUoFAWwC5xEXkdABu1CP9mgTIlkjPgBq-sXfIvUBLs-nqnvLYWFpNAoy2WHapnog57fGj7G7XD7BxZE9SDf=w1029-h682)


## Installation
- Create or open a spreadsheet in Google Sheets
- Select the menu item Tools > Script editor
- Copy the code in the gsCode.gs file and paste into the Script editor
- Authorize the code as normal
- Enter the function in a cell, ensuring your include 2 valid parameters

