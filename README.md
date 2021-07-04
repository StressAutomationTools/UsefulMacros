# UsefulMacros
A variety of macros

TakeMeThere
Contains a Macro that activates the cell from which the return value of the following functions comes from:
 -Max
 -Min
 -V/Hlookup
 -Index(Match)
 
 A1All
 Activates cell A1 on all worksheets
 
 RemoveFilters
 Removes all active filters on all sheets to make all data visible
 
 SpellcheckAll
 Goes through all sheets and starts the spellchecker

# PatranListMacros

ExpandList
Provide a starting cell reference and a patran list.
The program will output a column starting from the starting cell with the expanded list (exapnding n:m, n: m :s type lists)

SelectionToClipboard
Turns the content of the selected cells into a space separated list and safes it to the clipboard, so it can be pasted into patran.
(Ensure the first cell in the selection is an identifier or otherwise provide an identifier in patran (Elm, Node etc.) before pasting the list). No compression of the list (n:m) is done, so the lists will be long.

SelectionToClipboardPatranCompress
Same as above but with compression (n:m and n:m:s)

# Envelope
A macro to create envelopes for 2D Data. Splits data into quadrants and creates envelopes for each quadrant including the necessary intermediate points to square off and connect the plot segments.
