# PDF-to-Excel
Simple script converts PDF to text to XLS

This was a basic proof of concept, to prove my wife it could be done.

1 - The script lists all files in a given directory (be sure to have only PDF files in there)

2 - Converts PDF to text string using Tika library https://pypi.org/project/tika/

3 - Searches for fixed pices of strings, and then by reference extracts the values and adds them to a list (sorry but no time to learn regullar expressions)

4 - It appends the list to a xls file, one line per pdf

Note: I'm not adding any pdf because they contain private information, but you easily get the point, the xls file (list of values.xlsx) has the headers of the information I wanted to extract from the pdf
