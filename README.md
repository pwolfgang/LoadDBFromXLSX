# Program to load a database table from an Excel xlsx file.

Program to load a database table from a specified spreadsheet within an Excel
workbook contained in an xlsx file.  The class DoUpload can also be called from 
a web server application. The first row of the worksheet must contain column 
labels corresponding to the database columns. The database table must already 
exist and the data types are determined from the database metadata. The cells 
must not contain formulae.  Cells that contain formula will be treated as NULL. 
If the sheet contain formulae copy the sheet to a blank sheet using paste-values.
This program has only been tested with MySQL database targets, but earlier 
versions of the code were used to load a Microsoft SQLServer database.

The command line arguments are as follows:
<dl>
<dt>args[0]</dt><dd>Text file containing DataSource parameters.</dd>
<dt>args[1]</dt><dd>The name of the table.</dd>
<dt>args[2]</dt><dd>The name of the xlsx file.</dd>
<dt>args[3]</dt><dd>The name of the sheet in the workbook containing the data</dd>
</dl>

