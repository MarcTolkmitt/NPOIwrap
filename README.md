
THIS DLL IS AVAILABLE AS NUGET-PACKAGE 'helper-**NPOIwrap**-use-Excel-xlsx.1.0.2.nupkg'   
## 1. my motivation  
Seeing **NPOI** and using it is great. Usage for my own software is reading and writing 
data from/to real Excel files.  
To wrap **NPOI** there are just a few use cases that need to be cared for if you listen to **NPOI's** creator and your own view about the data I/O.  
There are just 3 kinds of cells used: 'string', 'numeric' and 'function' - while i don't need 'function'.  
Concerning that i analysed the lines: there are special ones ( 'mixed' ) and the general list types ( 'string', 'double' ). For both cases you can use the given classes:  

- **ExcelDataRow**: the mixed special version only as example. You would create your own data mixture in that way.  
- **ExcelDataRowListString/-Double**: the list type having only one celltype in the row  

Value is not in empty cells on the program side - but you buffer in whole blocks of data. 
You can use header lines to mark your data in your own logic.  
## 2. Procedure:
I added functions to 
get and to give data to the wrapper ( **<u>DataList...As...(), Array...ToDataList...()</u>** ). 
In the program you take an instance of the '**<u>NPOIexcel</u>**'-class and everything is wrapped.  

To **read** from an Excel-file:
- NPOIexcel myData = new NPOIexcel();	*// the wrapper for **NPOI***  
- myData.ReadWorkbook();	*// this will give you the file dialog*  
- myData.ReadSheets();	*// first overview of the given file for the workbook*
- myData.ReadSheetAsListDouble( 0 );	*// no header used, filled into dataListDouble*  
- double[][] doubles = myData.DataListDoubleAsArrayRagged();	*// there you have your Excel's file data to your convenience*
- string[] headers = myData.GetHeaderNo( sheetNumber );


You should add data from the program side into the lists and then **write** the file:
- myData.CreateWorkbook();	*// start empty*
- myData.CreateSheet( 0 );	*// new sheet to be filled*
- myData.ArrayRaggedToDataListDouble( doubles );	*// you give him your data*
- myData.CreateSheetFromListDouble( 0 );	*// this adds the data now to the workbook*
- myData.SaveWorkbook( fileName );	*// this will save the file in real excel format thanks to **NPOI***

I use lists to handle the workbook's possible complexity. They will be instanciated with standard values and later with the special operation of reading sheet# you can get your real headers, too.
### 3. Demoprogram

Demoprogram for the DLL is: WPFwithNPOI. It shows how easy you can read and write Excel-xlsx-files. Every menuitem uses its local version of the NPOIexcel-class and thus works as complete example about how-to-use the NPOIwrap on your own.


This helper class uses **NPOI** by  
**<u>Author: Tony Qu,NPOI Contributors</u>**  

