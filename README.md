PlsFix Google Sheets Add-on
============================================

This add-on is a simple tool to improve the functionality of your Google Sheets. With the click of a button, you can perform various actions such as error wrapping, flipping sign, commenting cells, cleaning cells, anchoring formulas, pasting exact and insert, and flattening cells.

Installation
------------

1.  Open the Google Sheet where you want to use the utilities.
2.  Click on the `Tools` menu and select `Script editor`.
3.  Copy and paste the code into the script editor.
4.  Save the script by clicking on `File` and then `Save`.
5.  Refresh the Google Sheet.

Features
--------

### Formulas & Data

-   Error Wrap: Wraps the selected formula with the `IFERROR` function to show the error message "NA" instead of an error.
-   Flip Sign (!): Flips the sign of numbers in the selected cells, negates positive numbers, and makes negative numbers positive.
-   Comment Cells (!): Comments the selected cells if they are numbers or formulas. If the cell is already commented, it will uncomment it.
-   Clean Cells: Cleans up the selected cells by removing all spaces, line breaks, and non-printable characters.
-   Anchor Formula (!): Anchors the formula in the selected cells to their relative position, so that they remain unchanged when inserted or deleted.
-   Paste Exact (!!): Pastes the exact value of the copied cells, ignoring any formulas, formatting, and data validation rules.
-   Paste Insert (!!): Inserts the copied cells as if they were typed directly into the selected cells.
-   Flatten Cells: Flattens the multi-dimensional arrays in the selected cells to a single dimension by concatenating the arrays with a comma.

### Sheets

-   Unhide Sheets: Unhides all hidden sheets in the current workbook.

### Formatting

-   AutoColor Selection: Automatically colors the background of the selected cells based on their values.

Usage
-----

The utilities are available in the Google Sheet through the menu bar: `Formulas & Data`, `Sheets`, and `Formatting`. To use the utility, select the cells you want to perform the action on and then click on the corresponding utility in the menu.

Contributions
-------------

If you have any suggestions or bug reports, please feel free to create an issue on this repository. Contributions are always welcome!
