# Excel-finder-macro
This macro can replace Excel's native Find capability.  It prompts the user for a word/string/number/character to search the active worksheet for, and when it finds the first occurrence, it does the following:
- highlights the cell in yellow (or turqoise, if it's already yellow)
- bolds the text (if it's not already bold)
- selects that cell, to move focus to it
- displays a message box with the number of occurrences found so far and the current cell's location (i.e. "#3: A5")

If there are multiple matches, the message box will give you the option to stop searching after each match or continue (except for the last match -- there will be nothing to search for after that).  Once you make your choice, the macro returns the current cell to its previous formatting.

The macro currently searches all cells within a range from the farthest row that has a value in column A, all the way over to column ZZ.  It can be made to search farther over, or search rows based on a different column, by changing the `cellRange` variable.  The macro is case-insensitive and will find a character or partial string inside other text.

To use this macro, load it into a macro-enabled Excel worksheet and assign it a shortcut key.  You can use `<CTRL>` `<F>` to replace the built-in Find function, or a different combination (i.e. `<CTRL>` `<SHIFT>` `<F>`).

### Excel versions tested in:
- Excel 2010 / Windows 7
- Excel 2016 / Windows 10
