# Excel
Microsoft Excel functions using [ClosedXML](https://github.com/ClosedXML/ClosedXML). Maybe change to [OpenXML SDK](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk) because its maintained by Microsoft.

```Javascript
function open() 
{

	var wb = xlsx.load("C:\\Users\\manuel.zarat\\Desktop\\Gold.xlsx"); // load an excel file	
	
	var allsheets = xlsx.list_ws(wb); // list all sheets in an array
	println(allsheets);
	
	var ws = xlsx.get_ws(wb, "Tabelle1"); // get a sheet by its name

	println(xlsx.columns(ws)); // no of total columns
	println(xlsx.rows(ws)); // no of total rows
	
	println(xlsx.usedcolumns(ws)); // no of used columns
	println(xlsx.usedrows(ws)); // no of used rows

	xlsx.set(ws, 1, 3, "Hello "); // write a value into the cell at row 1, column 3
	println(xlsx.get(ws, 1, 3)); // get the value from the cell at row 1, column 3
	
	wb.SaveAs("modified.xlsx"); // save the excel file to disk

}

function create() 
{	
	
	var wb = xlsx.new(); // create a new excel workbook	

	var sheet1 = xlsx.add_ws(wb, "User"); // add a worksheet named "User"
	var sheet2 = xlsx.add_ws(wb, "Gruppen"); // add a worksheet named "Gruppen"
	var sheet3 = xlsx.add_ws(wb, "Unknown"); // add a worksheet named "Unknown"
	
	xlsx.remove_ws(wb, "Unknown"); // remove a sheet by its name
	
	var sheets = xlsx.list_ws(wb); // list all sheets as an array
	var sheetname;
	foreach(sheetname in sheets)
		println(sheetname);
	
	xlsx.set(sheet1, 1, 1, "Hello world"); // write "Hello world" into cell at row 1, column 1
	xlsx.set(sheet1, 2, 1, 3.14159); // write 3.14159 into cell at row 2, column 1
	xlsx.set_formula(sheet1, 3, 1, "=SUM(A1:A2)"); // write a formula into cell at row 3, column 1
	print(xlsx.get_formula(sheet1, 3, 1)); // get a formula from cell at row 3, column 1
	
	//xlsx_close(wb);
	wb.SaveAs("HelloWorld.xlsx");
	
	
}
```
