PlsFix Google Sheets Add-on
============================================

This add-on is a simple tool to improve the functionality of your Google Sheets. With the click of a button, you can wrap your formulas in an `IFERROR` statement or flip the sign of numeric inputs and formulas in the selected range.

Usage
-----

After installation, the add-on will add a "Formulas" menu to your Google Sheets. From there, you can access the two functions:

### Error Wrap

The `Error Wrap` function will add an `IFERROR` statement to the selected cell's formula. If the formula already starts with an `IFERROR` statement, it will remove it and keep the original formula.

### Flip Sign

The `Flip Sign` function will invert the sign of numeric inputs and formulas in the selected range. It is useful when changing the sign convention in financial statements, for example.

Code
----

The code for this add-on consists of two functions: `errorWrap` and `flipSign`.

`errorWrap` takes the selected cell's formula and wraps it in an `IFERROR` statement, or removes the `IFERROR` statement if it already exists.

`flipSign` takes the selected range and inverts the sign of all numeric inputs and formulas. It also handles array formulas.

Contributions
-------------

If you have any suggestions or bug reports, please feel free to create an issue on this repository. Contributions are always welcome!
