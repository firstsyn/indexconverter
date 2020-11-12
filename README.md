Convert Excel .csv file to Word .docx file.

Convert an index for a course (i.e., SANS) from Excel to Word.  The first
step is to save the main Excel index spreadsheet from the .xlsx file to
a UTF-8 comma-delimited .csv file; this step is done manually in Excel.
This script then converts the .csv file to a Word .docx file in a
layout that can be easily printed and used during the open book cert exam.

Parameters
----------
csvfile : str
  Name of index CSV file to convert to .docx.

Output
------
<csvfile>.docx
  Word document created from CSV index file; created in directory script is run from.

Requires
--------
python-docx
  https://python-docx.readthedocs.io/en/latest/
  https://buildmedia.readthedocs.org/media/pdf/python-docx/latest/python-docx.pdf

Format
------
csvfile
  Comma-separated data columns: TOPIC, BK#, PG#, COMMENTS
  
