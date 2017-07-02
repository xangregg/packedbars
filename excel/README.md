Excel-related files

The VBA file is the script built-in to the MakePackedBars template.

Notes from the VBA file:

This script makes a Packed Bars chart from a two-column Excel data in columns A and B.
For general info, see my [blog post](https://community.jmp.com/t5/JMP-Blog/Introducing-packed-bars-a-new-chart-form/ba-p/39972).

This is my first VBA script, so don't assume I know what I'm doing but otherwise use
this code as you see fit.

The data must be in the first two columns of sheet one with one row of header labels. A few
tunable parameters are in columns D and E of the sheet:
 *   the number of primary categories (and the number of rows in the chart)
 *   value threshold for showing secondary labels (% or smallest primary bar value). Use 1.0 for none.
 *   graph width
 *   graph Height
 *   gap adjustment, postive for bigger gaps, negative for smaller gaps

What's working:
 *  Primary bars ordered with prominent color and labels left aligned
 *  Secondary bars randomish grays, some labeled
 *  Hover labels on all bars
Not working:
 *  No UI to choose data range
 *  No clipping/wrapping of labels that get too Long
 *  The gap between bars is relative to the bar width, instead of being a fixed small size
 *  The gap can be irregular if the bars don't fit into the graph height evenly
 *  We not too careful about avoiding Excel's limit of 256 series in one chart
 *  Some chart featurs are sensitive to current cell selection, requiring code to reset selection.

Xan Gregg July 2017, developed on Excel for Mac 15.35

The QuickSortArray function is from
   [stackoverflow](https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba/5104206#5104206)

