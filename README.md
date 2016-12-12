# Excel-finder-macro
This macro can replace (or supplement) Excel's native find capability.  It prompts the user for a word/string/number/character to search for, and when it finds the first occurrence, it does the following:
- highlights the cell in yellow (or turqoise, if it's already yellow)
- bolds the text (if not already bold)
- selects that cell

A message box pops up with the cell's location (i.e. "A5").  When you close the message box, the macro will continue to search to see if there are more occurrences.  If there are, it moves on to the next one; if not, it ends.  After finding each occurrence and highlighting it, the macro puts it back to its previous formatting.

The macro currently searches all cells within a range from the farthest row that has a value in column A, all the way over to column ZZ.  You can make it search farther over, or make it search rows based on a different column, by changing the `cellRange` variable.

To use this macro, load it into a macro-enabled Excel worksheet and give it a shortcut key.  You can use `<CTRL>` `<F>` to replace the native find function if you want, or something different (i.e. `<CTRL>` `<SHIFT>` `<F>`).
