# vba
Helpful Classes and Modules

'Declare the following values as global variables in a separate module plus include the GetRange Function by Johan Kreszner below
'Or just copy everything below this line.

Global thisDBCon As New clsDatabaseConnection
Global InnerValuesArray() As Variant
Global OuterValuesArray() As Variant
Global HeaderValuesArray() As Variant

'SAMPLE USAGE
'If you want to query a table from an MS ACCESS data source:

'thisDBCon.QueryDatabase "SELECT * FROM TABLE"

'to get the values from the array which queries the above statement
'FOR counter = 0 to UBOUND(OuterValuesArray())
'	Debug.Print OuterValuesArray(counter)(0)
'NEXT counter

'If you want to query a table from an MS EXCEL table data source:

'thisDBCon.QueryDatabase "SELECT * FROM " & GetRange(Table_Name_Here_With_Quotes"), -1

'the only difference is the -1 at the end
'you need to declare the paths to your MS ACCESS repositories in the clsDatabaseConnection_version2.cls

'it justs basically performs your query without the need to declare the nitty gritty details of db connections
'you can perform queries with a single line of code

Public Function GetRange(ByVal sListName As String) As String
'Regards to Johan Kreszner
'https://stackoverflow.com/questions/19755396/performing-sql-queries-on-an-excel-table-within-a-workbook-with-vba-macro

	Dim oListObject As ListObject
	Dim wb As Workbook
	Dim ws As Worksheet

	Set wb = ThisWorkbook

	For Each ws In wb.Sheets
		For Each oListObject In ws.ListObjects
			If UCase(oListObject.Name) = UCase(sListName) Then
				GetRange = "[" & ws.Name & "$" & Replace(oListObject.Range.Address, "$", "") & "]"
			Exit Function
			End If
		Next oListObject
	Next ws
End Function
