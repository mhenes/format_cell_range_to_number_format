# format_cell_range_to_number_format
Format range of cells to number format in excel using openpyxl.

Using openpyxl this script converts the specified range to no decimal number format.

USING THE SCRIPT
  Place the script in a folder of your choice.
  Place excel sheets to having formatting applied in the same folder as the script.
  Run the script in terminal or shell script: python3 format_cell_range_to_number_format.py
  The script will update the ranges and save the sheets.

EDITING THE CELL RANGE
  Edit line 30 "for row in ws.iter_rows('B2:C200'):"
    Change the values "B2:C200" to your desired cell range.
