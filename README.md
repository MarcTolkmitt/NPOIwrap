Seeing NPOI and using it is great. Usage for my own software is reading and writing 
data from/to real Excel files.  
To wrap NPOI there are just a few use cases that need to be cared for if you listen to NPOI's creator and your own view about the data I/O.  
There are just 3 kinds of cells used: 'string', 'numeric' and 'function' - while i don't need 'function'.  
Concerning that i analysed the rows: there are special ones ( 'mixed' ) and the generell list types ( 'string', 'double' ). For both cases you can use the given classes:  
-> ExcelDataRow: the mixed special version only as example. You would create your own data mixture in that way.  
-> ExcelDataRowList: the list type having only one celltype in the row   
Value is not in empty cells on the program side - but you buffer in whole blocks of data. 
You can use header lines to mark your data in your own logic.  
Procedure:  
\t- NPOIexcel myData = new NPOIexcel();  
\t- myData.ReadWorkbook();	// this will give you the file dialog  
\t- myData.ReadSheets();	// instanciates all sheets into the wrapper class  
\t- mydata.ReadSheetAsListDouble( 0 );	// no header, filled into dataListDouble  
-> there you have your Excel's file data to your convenience, you can now get the data 
with "double[][] doubles = myData.DataListDoubleAsArrayRagged();".  
I added functions to 
get and to give data to the wrapper ( DataList*As*(), Array*ToDataList*() ). 
In the program you take an instance of the 'NPOIexcel'-class and everything is wrapped. 
You should add data from the program side into the lists and then write the file.  
Demoprogram for the DLL is: WPFwithNPOI  
\tThis helper class uses NPOI by  
\tAuthor: Tony Qu,NPOI Contributors  
