/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for Additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Xml;

namespace NPOIwrap
{
    /// <summary>
    /// Wrappper for the work with Excel using NPOI.
    /// <para></para>
    /// </summary>
    public class NPOIexcel
    {
        // Erstellt ab: 08.02.2024
        // letzte Änderung: 13.09.24
        public Version version = new Version("1.0.3");
        // local variables
        /// <summary>
        /// Excel file ending
        /// </summary>
        public string fileEnding = ".xlsx";
        /// <summary>
        /// example filename
        /// </summary>
        public string fileName = @"DemoExcelFile.xlsx";
        /// <summary>
        /// the read excel file is handled by NPOI as a 'IWorkbook'
        /// </summary>
        IWorkbook workbook;
        /// <summary>
        /// usable globally in the workbook
        /// </summary>
        List<ISheet> sheets = new List<ISheet>();
        /// <summary>
        /// usable globally in the workbook
        /// </summary>
        public List<string> sheetsNames = new List<string>();
        /// <summary>
        /// usable globally in the workbook
        /// </summary>
        public List<ExcelDataRowListString> sheetsHeaders = new List<ExcelDataRowListString>();
        /// <summary>
        /// usable globally in the workbook
        /// </summary>
        public List<bool> sheetsHeadersBool = new List<bool>();
        /// <summary>
        /// general data list will be filled with ReadSheetAsListString
        /// </summary>
        public List<ExcelDataRowListString> dataListString = 
            new List<ExcelDataRowListString>();
        /// <summary>
        /// general data list will be filled with ReadSheetAsListDouble
        /// </summary>
        public List<ExcelDataRowListDouble> dataListDouble = 
            new List<ExcelDataRowListDouble>();
        /// <summary>
        /// general data list will be filled with ReadSheetAsListMixed
        /// </summary>
        public List<ExcelDataRow> dataListMixed = new List<ExcelDataRow>();

        /// <summary>
        /// Construktor: the interface classes ( 'IWorkbook, Ixyz, ... )
        /// need to be nullable or not, but i use lists mostly.
        /// </summary>
        public NPOIexcel( )
        {
            workbook = new XSSFWorkbook();
            sheetsNames.Add( " table 0 " );
            
        }   // end: public NPOIexcel

        /// <summary>
        /// Attempts to read the '.xlsx'-Excel file.
        /// </summary>
        /// <param name="path">given file should be in real '.xlsx'-format</param>
        public void ReadWorkbook( string name = "", bool silent = false )
        {
            if ( !string.IsNullOrEmpty( name ) )
                fileName = GetCurrentDir() + name;
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "choose your Excel-File";
            dialog.DefaultExt = fileEnding;
            dialog.Filter = "Excel-File (.xlsx) | *.xlsx"; // Filter files by extension
            dialog.FileName = fileName;
            
            if ( !silent )
            {
                var result = dialog.ShowDialog();
                if ( result == true )
                    fileName = dialog.FileName;

            }

            try
            {
                FileStream fs = new FileStream ( fileName,
                    FileMode.Open, FileAccess.Read,
                    FileShare.ReadWrite );
                workbook = new XSSFWorkbook( fs );
                fs.Close();

            }
            catch ( Exception ex )
            {
                MessageBox.Show( ex.Message,
                    "Excel read error", MessageBoxButton.OK,
                    MessageBoxImage.Error );

            }

        }   // end: public void ReadWorkbook

        /// <summary>
        /// writes the workbook to a new created file
        /// </summary>
        /// <param name="path">the new filename if given</param>
        public void SaveWorkbook( string name = "", bool silent = false )
        {
            if ( !string.IsNullOrEmpty( name ) )
                fileName = GetCurrentDir() + name;
            var dialog = new Microsoft.Win32.SaveFileDialog();
            dialog.Title = "choose your Excel-File";
            dialog.DefaultExt = fileEnding;
            dialog.Filter = "Excel-File (.xlsx) | *.xlsx"; // Filter files by extension
            dialog.FileName = fileName;

            if ( !silent )
            {
                var result = dialog.ShowDialog();
                if ( result == true )
                    fileName = dialog.FileName;

            }

            try
            {
                FileStream fs = new FileStream( fileName,
                    FileMode.Create, FileAccess.Write );
                workbook.Write( fs );
                fs.Close();

            }
            catch ( Exception ex )
            {
                MessageBox.Show( ex.Message,
                    "write to file error -> is it open in Excel?",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error );

            }

        }   // end: public void SaveWorkbook

        /// <summary>
        /// Creates a new empty workbook
        /// </summary>
        public void CreateWorkbook( )
        {
            workbook = new XSSFWorkbook();
            sheets.Clear();
            sheetsHeaders.Clear();
            sheetsHeadersBool.Clear();
            sheetsNames.Clear();
            dataListString.Clear();
            dataListDouble.Clear();
            dataListMixed.Clear();

        }   // end: public void CreateWorkbook

        /// <summary>
        /// Without any sheet Excel won't be able to read the file!
        /// <para/>
        /// Every new sheet will be at the last position or an existing
        /// one will be overwritten. Just use an existing number for that.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="name">new name or one will be made</param>
        /// <param name="withHeader">using a header row ?</param>
        /// <returns>the number of the created sheet</returns>
        public int CreateSheet( int number, string name = "", bool withHeader = false )
        {
            // first a new sheet by creation or deletion
            if ( number > ( workbook.NumberOfSheets - 1 ) )
            {   // no sheet with such a number
                sheets.Add( workbook.CreateSheet( name ) );
                number = workbook.NumberOfSheets - 1;
                sheetsHeadersBool.Add( withHeader );
                sheetsHeaders.Add( new ExcelDataRowListString() );
                sheetsHeaders[ number ].cellData.Add( "empty header" );
                sheetsNames.Add( "no name yet" );

            }
            else if ( sheets[ number ].LastRowNum >= 0 )
            {   // sheet is not empty
                for ( int i = sheets[ number ].LastRowNum; i == 0; i++ )
                    sheets[ number ].RemoveRow( sheets[ number ].GetRow( i ) );
                sheetsHeadersBool[ number ] = withHeader;
                sheetsHeaders[ number ] = new ExcelDataRowListString();
                sheetsHeaders[ number ].cellData.Add( "empty header" );

            }
            // then the name or a made one
            string sheetName;
            if ( !string.IsNullOrEmpty( name ) )
                sheetName = name;
            else
                sheetName = $"table {number}";
            sheetsNames[ number ] = sheetName;
            // update the local data
            //ReadSheets();

            return ( number );

        }   // end: public void CreateSheet

        /// <summary>
        /// For the first demo.
        /// </summary>
        public void CreateHelloWorld( )
        {
            IRow row;
            ICell cell;
            // first a new workbook
            CreateWorkbook();
            CreateSheet( 0, "Hello World" );
            // the data
            row = sheets[ 0 ].CreateRow( 0 );
            cell = row.CreateCell( 0, CellType.String );
            cell.SetCellValue( "Hello" );
            cell = row.CreateCell( 1, CellType.String );
            cell.SetCellValue( "World" );
            cell = row.CreateCell( 2, CellType.String );
            cell.SetCellValue( ".. greets Marc from germany." );
            row = sheets[ 0 ].CreateRow( 1 );
            cell = row.CreateCell( 0, CellType.String );
            cell.SetCellValue( "0.815" );
            cell = row.CreateCell( 1, CellType.String );
            cell.SetCellValue( "13" );
            for ( int i = 0; i < 3; i++ )
                sheets[ 0 ].AutoSizeColumn( i );
            // now into a file
            SaveWorkbook( "HelloWorld.xlsx", true );

        }   // end: public void CreateHelloWorld

        /// <summary>
        /// Reads the sheets out of the workbook, but
        /// the data has to be read somewhere else. This includes
        /// the headers -> ReadSheetAsListDouble/-String with 
        /// 'useHeader' = true.
        /// </summary>
        public void ReadSheets( )
        {
            sheets.Clear();
            sheetsHeaders.Clear();
            sheetsHeadersBool.Clear();
            sheetsNames.Clear();
            for ( int i = 0; i < workbook.NumberOfSheets; i++ )
            {
                sheets.Add( workbook.GetSheetAt( i ) );
                sheetsHeaders.Add( new ExcelDataRowListString() );
                sheetsHeaders[ i ].cellData.Add( "init sheets" );
                sheetsHeadersBool.Add( false );
                sheetsNames.Add( "empty header" );

            }

        }   // end: public void ReadSheets

        /// <summary>
        /// Reads all rows of the sheet into the local
        /// version.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool ReadSheetAsListString( int number, bool useHeader = false,
            bool verbose = false )
        {
            if ( number >= sheets.Count )
                return ( false );

            int firstRow = 0;
            int deltaRow = 0;
            dataListString.Clear();
            if ( useHeader )
            {   // read the first row as the header strings
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].FromRow( sheets[ number ].GetRow( 0 ) );
                firstRow++;
                deltaRow--;

            }
            for ( int i = firstRow; i <= sheets[ number ].LastRowNum; i++ )
            {   // read the sheet's rows
                dataListString.Add( new ExcelDataRowListString( verbose, verbose ) );
                dataListString[ i + deltaRow ].FromRow( sheets[ number ].GetRow( i ) );

            }

            return ( true );
        }   // end: public void ReadSheetAsListString

        /// <summary>
        /// Put all local rows into a new sheet.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="name">name of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool CreateSheetFromListString( int number, string name = "", bool useHeader = false )
        {
            number = CreateSheet( number, name, useHeader );
            IRow row;

            int deltaRow = 0;
            if ( useHeader )
            {   // put the header strings as the first row
                row = sheets[ number ].CreateRow( 0 );
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].AsRow( ref row );

                deltaRow++;

            }
            for ( int i = 0; i < dataListString.Count; i++ )
            {   // put the sheet's rows
                row = sheets[ number ].CreateRow( i + deltaRow );
                dataListString[ i ].AsRow( ref row );

            }

            return ( true );
        }   // end: public void CreateSheetFromListString

        /// <summary>
        /// Reads all rows of the sheet into the local
        /// version.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool ReadSheetAsListDouble( int number, bool useHeader = false, 
            bool verbose = false )
        {
            if ( number >= sheets.Count )
                return ( false );

            int firstRow = 0;
            int deltaRow = 0;
            dataListDouble.Clear();
            if ( useHeader )
            {   // read the first row as the header strings
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].FromRow( sheets[ number ].GetRow( 0 ) );
                firstRow++;
                deltaRow--;

            }
            for ( int i = firstRow; i <= sheets[ number ].LastRowNum; i++ )
            {   // read the sheet's rows
                dataListDouble.Add( new ExcelDataRowListDouble( verbose, verbose ) );
                dataListDouble[ i + deltaRow ].FromRow( sheets[ number ].GetRow( i ) );

            }

            return ( true );
        }   // end: public void ReadSheetAsListDouble

        /// <summary>
        /// Put all local rows into a new sheet.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="name">name of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool CreateSheetFromListDouble( int number, string name = "", bool useHeader = false )
        {
            number = CreateSheet( number, name, useHeader );
            IRow row;

            int deltaRow = 0;
            if ( useHeader )
            {   // put the header strings as the first row
                row = sheets[ number ].CreateRow( 0 );
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].AsRow( ref row );

                deltaRow++;

            }
            for ( int i = 0; i < dataListDouble.Count; i++ )
            {   // put the sheet's rows
                row = sheets[ number ].CreateRow( i + deltaRow );
                dataListDouble[ i ].AsRow( ref row );

            }

            return ( true );
        }   // end: public void CreateSheetFromListDouble

        /// <summary>
        /// Reads all rows of the sheet into the local
        /// version.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool ReadSheetAsListMixed( int number, bool useHeader = false )
        {
            if ( number >= sheets.Count )
                return ( false );

            int firstRow = 0;
            int deltaRow = 0;
            dataListMixed.Clear();
            if ( useHeader )
            {   // read the first row as the header strings
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].FromRow( sheets[ number ].GetRow( 0 ) );
                firstRow++;
                deltaRow--;

            }
            for ( int i = firstRow; i <= sheets[ number ].LastRowNum; i++ )
            {   // read the sheet's rows
                dataListMixed.Add( new ExcelDataRow() );
                dataListMixed[ i + deltaRow ].FromRow( sheets[ number ].GetRow( i ) );

            }

            return ( true );
        }   // end: public void ReadSheetAsListMixed

        /// <summary>
        /// Put all local rows into a new sheet.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="name">name of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool CreateSheetFromListMixed( int number, string name = "", bool useHeader = false )
        {
            number = CreateSheet( number, name, useHeader );
            IRow row;

            int deltaRow = 0;
            if ( useHeader )
            {   // put the header strings as the first row
                row = sheets[ number ].CreateRow( 0 );
                sheetsHeadersBool[ number ] = true;
                sheetsHeaders[ number ].AsRow( ref row );

                deltaRow++;

            }
            for ( int i = 0; i < dataListMixed.Count; i++ )
            {   // put the sheet's rows
                row = sheets[ number ].CreateRow( i + deltaRow );
                dataListMixed[ i ].AsRow( ref row );

            }

            return ( true );
        }   // end: public void CreateSheetFromListMixed

        /// <summary>
        /// Changes the header of a sheet - only if it exists.
        /// </summary>
        /// <param name="number">sheet number</param>
        /// <param name="heads">a field of strings</param>
        /// <returns>truth of success</returns>
        public bool ChangeHeader( int number, string[] heads )
        {
            if ( number > ( workbook.NumberOfSheets - 1 ) )
                return ( false );
            sheetsHeadersBool[ number ] = true;
            sheetsHeaders[ number ].cellData.Clear();
            foreach( string head in heads )
                sheetsHeaders[ number ].cellData.Add( head );
            var row = sheets[ number ].GetRow( 0 );
            sheetsHeaders[ number ].AsRow( ref row );
            SaveWorkbook( "", true );
            
            return ( true );
        }   // end: public void AddHeaders

        /// <summary>
        /// For the first demo.
        /// </summary>
        public void ReadHelloWorld( )
        {
            ReadWorkbook( "HelloWorld.xlsx", true );
            ReadSheets();
            ReadSheetAsListString( 0 );

        }   // end: public void ReadHelloWorld

        /// <summary>
        /// ToString(): for the row list having string cells.
        /// </summary>
        /// <returns>the message</returns>
        public string DataListString_ToString( int sheetNo = 0, bool useHeader = false )
        {
            string text = "";
            if ( useHeader )
                text += sheetsHeaders[ sheetNo ].ToString() + "\n";
            if ( dataListString.Count > 0 )
                for ( int i = 0; i < dataListString.Count; i++ )
                    text += dataListString[ i ].ToString() + "\n";
            else
                text += "empty list... \n";

            return ( text );

        }   // end:public void DataListString_ToString

        /// <summary>
        /// ToString(): for the row list having double cells.
        /// </summary>
        /// <returns>the message</returns>
        public string DataListDouble_ToString( int sheetNo = 0, bool useHeader = false )
        {
            string text = "";
            if ( useHeader )
                text += sheetsHeaders[ sheetNo ].ToString() + "\n";
            if ( dataListDouble.Count > 0 )
                for ( int i = 0; i < dataListDouble.Count; i++ )
                    text += dataListDouble[ i ].ToString() + "\n";
            else
                text += "empty list... \n";

            return ( text );

        }   // end:public void DataListDouble_ToString

        /// <summary>
        /// ToString(): for the mixed list
        /// </summary>
        /// <returns>the message</returns>
        public string DataListMixed_ToString( int sheetNo = 0, bool useHeader = false )
        {
            string text = "";
            if ( useHeader )
                text += sheetsHeaders[ sheetNo ].ToString() + "\n";
            if ( dataListMixed.Count > 0 )
                for ( int i = 0; i < dataListMixed.Count; i++ )
                    text += dataListMixed[ i ].ToString() + "\n";
            else
                text += "empty list... \n";

            return ( text );

        }   // end:public void DataListMixed_ToString

        /// <summary>
        /// delivers the current directory
        /// </summary>
        /// <returns>the path</returns>
        public string GetCurrentDir( )
        {
            string text =
                Directory.GetCurrentDirectory() +
                System.IO.Path.DirectorySeparatorChar;

            return ( text );

        }   // end: public string GetCurrentDir

        // ----------------------------------- QoL input/output
        
        /// <summary>
        /// Gives you the data as 'string[,]'.
        /// </summary>
        /// <returns>two dimensional array ( string[,] )</returns>
        public string[,] DataListStringAsArray()
        {
            string[,] temp = new string[ dataListString.Count,
                dataListString[ 0 ].cellData.Count];
            for ( int line = 0; line < dataListString.Count; line++ )
                for ( int column = 0; column < dataListString[ 0 ].cellData.Count; column++ )
                    temp[ line, column ] = dataListString[ line ].cellData[ column ];

            return( temp );

        }   // end: DataListStringAsArray

        /// <summary>
        /// Gives you the data as 'string[][]'.
        /// </summary>
        /// <returns>ragged array ( string[][] )</returns>
        public string[][] DataListStringAsArrayRagged( )
        {
            string[][] temp = new string[ dataListString.Count ][];
                
            for ( int line = 0; line < dataListString.Count; line++ )
                for ( int column = 0; column < dataListString[ 0 ].cellData.Count; column++ )
                {
                    temp[ line ] = new string[ dataListString[ 0 ].cellData.Count ];
                    temp[ line ][ column ] = dataListString[ line ].cellData[ column ];
                
                }

            return ( temp );

        }   // end: DataListStringAsArrayRagged

        /// <summary>
        /// Stores your data into the handler.
        /// </summary>
        /// <param name="data">a two dimensinal array ( string[,] )</param>
        public void ArrayToDataListString( string[,] data )
        {
            dataListString.Clear();
            for ( int line = 0; line < data.Length; line++ )
            {
                string[] dataLine = new string[ data.Length ];
                for ( int index = 0; index < dataLine.Length; index++ )
                    dataLine[ index ] = data[ line, index ];
                ExcelDataRowListString newLine = new ExcelDataRowListString();
                newLine.ArrayToCellData( dataLine );
                dataListString.Add( newLine );

            }

        }   // end: ArrayToDataListString

        /// <summary>
        /// Stores your data into the handler.
        /// </summary>
        /// <param name="data">a two dimensinal array ( string[,] )</param>
        public void ArrayRaggedToDataListString( string[][] data )
        {
            dataListString.Clear();
            for ( int line = 0; line < data.Length; line++ )
            {
                ExcelDataRowListString newLine = new ExcelDataRowListString();
                newLine.ArrayToCellData( data[ line ] );
                dataListString.Add( newLine );

            }

        }   // end: ArrayToDataListString

        /// <summary>
        /// Gives you the data as 'double[,]'.
        /// </summary>
        /// <returns>two dimensional array ( double[,] )</returns>
        public double[,] DataListDoubleAsArray( )
        {
            double[,] temp = new double[ dataListDouble.Count,
                dataListDouble[ 0 ].cellData.Count];
            for ( int line = 0; line < dataListDouble.Count; line++ )
                for ( int column = 0; column < dataListDouble[ 0 ].cellData.Count; column++ )
                    temp[ line, column ] = dataListDouble[ line ].cellData[ column ];

            return ( temp );

        }   // end: DataListDoubleAsArray

        /// <summary>
        /// Gives you the data as 'double[][]'.
        /// </summary>
        /// <returns>ragged array ( double[][] )</returns>
        public double[][] DataListDoubleAsArrayRagged( )
        {
            double[][] temp = new double[ dataListDouble.Count ][];

            for ( int line = 0; line < dataListDouble.Count; line++ )
                for ( int column = 0; column < dataListDouble[ 0 ].cellData.Count; column++ )
                {
                    temp[ line ] = new double[ dataListDouble[ 0 ].cellData.Count ];
                    temp[ line ][ column ] = dataListDouble[ line ].cellData[ column ];

                }

            return ( temp );

        }   // end: DataListDoubleAsArrayRagged

        /// <summary>
        /// Stores your data into the handler.
        /// </summary>
        /// <param name="data">a two dimensinal array ( double[,] )</param>
        public void ArrayToDataListDouble( double[,] data )
        {
            dataListDouble.Clear();
            for ( int line = 0; line < data.Length; line++ )
            {
                double[] dataLine = new double[ data.Length ];
                for ( int index = 0; index < dataLine.Length; index++ )
                    dataLine[ index ] = data[ line, index ];
                ExcelDataRowListDouble newLine = new ExcelDataRowListDouble();
                newLine.ArrayToCellData( dataLine );
                dataListDouble.Add( newLine );

            }

        }   // end: ArrayToDataListDouble

        /// <summary>
        /// Stores your data into the handler.
        /// </summary>
        /// <param name="data">a two dimensinal array ( double[,] )</param>
        public void ArrayRaggedToDataListDouble( double[][] data )
        {
            dataListDouble.Clear();
            for ( int line = 0; line < data.Length; line++ )
            {
                ExcelDataRowListDouble newLine = new ExcelDataRowListDouble();
                newLine.ArrayToCellData( data[ line ] );
                dataListDouble.Add( newLine );

            }

        }   // end: ArrayRaggedToDataListDouble

    }   // end: internal class NPOIexcel

}
