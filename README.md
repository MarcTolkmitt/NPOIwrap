NPOI is installed with a bunch of other packets with NUGET and i don't like to see that in my programs - it is too confusing, anyone likes to know whats used.  
Thats why i created this DLL to wrap NPOI. There are just a few use cases that need to be done if you listen to NPOI's creator and your own view towards it.  
There are just 3 kinds of cells used: 'string', 'numeric' and 'function' - where i don't need 'function'.  
Concerning that i analysed the rows: there are special ones ( 'mixed' ) and the generell list types ( 'string', 'double' ). For both cases you can use the given classes:  
-> ExcelDataRow: the mixed special version only as example. You would create your own data mixture in that way.  
-> ExcelDataRowList: the list type having only one celltype in the row   
Value is not in empty cells on the program side. To form new mixed rows programmaticly will lead to new mixed rows and that too has no value for me.  
In the program you take an instance of the 'NPOIexcel'-class and everything is wrapped. 
You should add data from the program side into the lists and then write the file.  
Demoprogram for the DLL is: WPFwithNPOI
