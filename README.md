THIS DLL IS AVAILABLE AS NUGET-PACKAGE 'helper-NPOIwrap-use-Excel-xlsx.1.0.0.nupkg'   

Demoprogram for the DLL is: WPFwithNPOI. It shows how easy you can read and write Excel-xlsx-files. Every menuitem uses its local version of the NPOIexcel-class and thus works as complete example about how-to-use the NPOIwrap on your own.

  Thoughts about my motivation:  
To wrap NPOI there are just a few use cases that need to be done if you listen to NPOI's creator and your own view towards it.  
There are just 3 kinds of cells used: 'string', 'numeric' and 'function' - where i don't need 'function'.  
Concerning that i analysed the rows: there are special ones ( 'mixed' ) and the generell list types ( 'string', 'double' ). For both cases you can use the given classes:  
-> ExcelDataRow: the mixed special version only as example. You would create your own data mixture in that way.  
-> ExcelDataRowList: the list type having only one celltype in the row   
Value is not in empty cells on the program side. To form new mixed rows programmaticly will lead to new mixed rows and that too has no value for me.  
In the program you take an instance of the 'NPOIexcel'-class and everything is wrapped. You still need to include the NPOI-package in your program to make things work.   
You should add data from the program side into the lists and then write the file.  

