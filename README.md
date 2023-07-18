# Excel Chart Utilities
Some Excel VBA scripts to format charts faster

Purpose:
1. coloring.bas formats all charts or cells in a sheet depending on usage. Use with painter_table.xlsx to manage and input chart formats. To use the script, you need to first input your chart formats in painter_table.xlsx. The Excel file has the following columns:
- C - Legend - This corresponds to the legends on your chart
- D - Color - Your preferred color for a certain series. This usually pertains to foreground color
- E - Bar Pattern - An Excel constant which can be picked from a drop-down list. If not specified, defaults to no pattern
- F - Bar Pattern Color - The background color for the bar pattern, if it was chosen. If not specified, defaults to a light blue color commonly recognized as "no color" for Excel
- G - Weight - For lines, specifically the thickness. If not specified, defaults to Excel's default line thickness
- H - Dashed? - Whether to make dashed lines
- I - Dash Type - If dashed, an Excel constant which can be picked from a drop-down list. If not specified, defaults to solid
- J - Marker? - Whether to put markers on lines
- K - Marker Type - If to put markers, an Excel constant which can be picked from a drop-down list. If not specified, defaults to no marker
- L - Marker Size - If to put markers, specify the size of the marker. If not specified, defaults to Excel's default marker size
- M - Marker Forecolor
- N - Marker Backcolor
- O - Opacity

Upon entering your formats, highlight the area from R to AD, along the rows where you entered the formats, and then press Ctrl+Shift+W to run the script. This will record all of the data under the specified area. After that, you can go to your sheets containing charts and press Ctrl+Shift+E to color your charts. There are several modes for this script:
- When the cell cursor is on a blank cell, color all charts
- When the cell cursor spans a cell with text, go to Cell Mode and color all cells that contain the legends specified in the painter_table
- When a specific chart is selected, color only that chart (does not work on multiple selected charts at the moment)

2. chartlabeler.bas changes the data labels in a chart by linking it to a specific cell. To use, first select the area where you want the values to be displayed in the chart. Press Ctrl+Shift+N to make the script remember that area. After that, select your chart and press Ctrl+Shift+X to run the script. This will change all the data labels into formula-based ones that link to the area that you selected with Ctrl+Shift+N.

If you encounter any issues or confusions using the script, feel free to open an issue.
