# ExeclTOSQL
When you enter some data in Excel it will create sql table rows

Assume the table structure is 
CREATE TABLE [dbo].[PersonalDetails](
	[Id] [int] NULL,
	[Name] [nchar](100) NULL,
	[Address] [nchar](100) NULL
) ON [PRIMARY]

```sql
CREATE TABLE [dbo].[PersonalDetails](
    [Id] [int] NULL,
    [Name] [nchar](100) NULL,
    [Address] [nchar](100) NULL
) ON [PRIMARY]



![image](https://github.com/prasad1508/ExeclTOSQL/assets/7384960/457da7c0-e00e-45e8-bb05-b00bb462db67)
--------------------------------------
|  Id  |    Name     |   Address    |
--------------------------------------
|  1   |   prasad    |    Galle     |
|  2   |  madushan   |   Mathara    |
|  3   | samarasekara|   Colombo    |
--------------------------------------


Open excel and alt+F8 will open macro
copy and paste the below code to macro and make the necessary changes in the code

Sub InsertDataIntoSQLTable()
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    ' Define your SQL Server connection string
    connStr = "Provider=SQLOLEDB;Data Source=COMPUTER-NAME;Initial Catalog=prasadtest;Integrated Security=SSPI;"
    
    ' Open the database connection
    conn.Open connStr
    
    ' Create a record set
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Define the Excel sheet and range where your data is located
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your sheet name
    
    ' Get the column names from the first row of the Excel sheet
    Dim colNames As String
    colNames = "Id, Name, Address"
    
    ' Loop through the rows in the Excel sheet and insert values dynamically
    Dim row As Range
    For Each row In ws.UsedRange.Rows
        ' Extract the "Id" value and convert it to an integer
        Dim idValue As Integer
        idValue = CInt(row.Cells(1, 1).Value)
        
        ' Build the VALUES part of the SQL statement dynamically
        Dim values As String
        values = idValue & ", '" & row.Cells(1, 2).Value & "', '" & row.Cells(1, 3).Value & "'"
        
        ' Construct the complete SQL statement
        strSQL = "INSERT INTO PersonalDetails (" & colNames & ") VALUES (" & values & ")"
        
        ' Execute the SQL statement
        conn.Execute strSQL
    Next row
    
    ' Close the database connection
    conn.Close
    
    MsgBox "Data inserted successfully."
End Sub

