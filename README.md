# cla_Odbc
Demo of using ODBC 13 with Clarion win32.

Demo shows how to use the ODBC 13.1 driver with clarion.  

The ODBC 13 driver does not expose ANSI strings and clarion does not support wide strings,
at least not yet. some day, soon (tm), maybe, it could happen, ...

Anyway. the code uses the svcom.* files to convert the ansi string from clarion to wide srings  when required by the driver.

there are examples of calling a query, 
calling stored procedures, with and without parameters,
calling a stored procedure with multiple result sets, the demo use two result sets.
caling a stroed proceudre with out parameters,
calling a scalar function,
calling a stored procedure with a table valued parameter, the table input can be used to insert, update and delete rows.

