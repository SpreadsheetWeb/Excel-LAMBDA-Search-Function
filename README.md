CustomSearch Excel LAMBDA Function
Overview
CustomSearch is a custom LAMBDA function for Excel that enables advanced search operations within a table. This function searches for text within a specified column and returns specified columns of the matching rows.
Function Definition
=LAMBDA(SearchText;DataTable;SearchColumn;ReturnColumns;NotFoundText;ErrorText;
  LET(
    ColumnExists; ISNUMBER(MATCH(SearchColumn; SEQUENCE(1; COLUMNS(DataTable)); 0));
    ReturnColumnsValid; AND(ISNUMBER(MATCH(ReturnColumns; SEQUENCE(1; COLUMNS(DataTable)); 0)));
    IF(
      ISBLANK(SearchText);
      NotFoundText;
      IF(
        NOT(ColumnExists);
        ErrorText;
        IF(
          NOT(ReturnColumnsValid);
          ErrorText;
          LET(
            FilteredData; IFERROR(FILTER(DataTable; ISNUMBER(SEARCH(UPPER(SearchText); UPPER(INDEX(DataTable;;SearchColumn))))); "");
            IF(
              COUNTA(FilteredData) = 0;
              NotFoundText;
              LET(
                Result; IFERROR(CHOOSECOLS(FilteredData; ReturnColumns); NotFoundText);
                IF(
                  ISERROR(Result);
                  ErrorText;
                  Result
                )
              )
            )
          )
        )
      )
    )
  )
)
How to Create and Use the CustomSearch Function
Step 1: Define the CustomSearch Function
1.	Open Excel and go to the Formulas tab.
2.	Click on Name Manager and then click New.
3.	In the Name field, enter CustomSearch.
4.	In the Refers to field, paste the LAMBDA function definition provided above.
5.	Click OK to save the new named formula.
Step 2: Using the CustomSearch Function
To use the CustomSearch function in your Excel workbook, follow these steps:
1.	Prepare your data table:
o	Ensure your data table is defined and named (e.g., Table1).
2.	Specify the search text:
o	Enter the search text in a cell, for example, cell A1.
3.	Enter the formula:
o	Use the following formula to perform the search and return the results:
=CustomSearch(A1, Table1, 2, {1, 2, 3, 4}, "No data found", "Error")
Parameters
•	SearchText: The text to search for. In the example, this is the value in cell A1.
•	DataTable: The table to search within. This should be a named range or a structured table (e.g., Table1).
•	SearchColumn: The column number within DataTable to search in. For example, 2 indicates the second column.
•	ReturnColumns: An array of column numbers to return from DataTable. For example, {1, 2, 3, 4} indicates that columns 1, 2, 3, and 4 should be returned.
•	NotFoundText: The text to display if no matching data is found (e.g., "No data found").
•	ErrorText: The text to display if there is an error in the function parameters (e.g., "Error").
Example
Assume you have the following data in Table1:
ID	Name	Department	Salary
1	Alice	HR	50000
2	Bob	IT	60000
3	Charlie	Finance	70000
4	David	IT	65000
You want to search for the term in cell A1 within the Name column (second column) and return the ID, Name, Department, and Salary columns.
1.	Enter Bob in cell A1.
2.	Use the formula:
=CustomSearch(A1, Table1, 2, {1, 2, 3, 4}, "No data found", "Error")
3.	The result will be:
ID	Name	Department	Salary
2	Bob	IT	60000
Notes
•	Ensure that SearchColumn and ReturnColumns are within the bounds of the DataTable columns.
•	SearchText is case-insensitive.
•	If no matching data is found, the function returns NotFoundText.
•	If there is an error in the parameters, the function returns ErrorText.
Conclusion
The CustomSearch LAMBDA function provides a powerful and flexible way to search within a data table and retrieve specific columns. By following the steps outlined in this guide, you can easily implement and use this function in your Excel workbooks.
Contributing
If you find any issues or have suggestions for improvements, please open an issue or submit a pull request on the GitHub repository.
License
This project is licensed under the MIT License.

