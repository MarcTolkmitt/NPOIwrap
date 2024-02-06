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
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace NPOIwrap
{
    /// <summary>
    /// Wrappper for the work with Excel using NPOI.
    /// </summary>
    public class NPOIexcel
    {
        // local variables
        public string fileEnding = ".xlsx";
        public string fileName = @"DemoExcelFile.xlsx";
        IWorkbook workbook;
        List<ISheet> sheets = new List<ISheet>();       // usable globally in the workbook
        List<string> sheetsNames = new List<string>();  // usable globally in the workbook
        List<ExcelDataRowList> sheetsHeaders = new List<ExcelDataRowList>();// usable globally in the workbook
        List<bool> sheetsHeadersBool = new List<bool>();    // usable globally in the workbook
        // the data lists should be public
        public List<ExcelDataRowList> dataListString = new List<ExcelDataRowList>();
        public List<ExcelDataRowList> dataListDouble = new List<ExcelDataRowList>();
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
                sheetsHeaders.Add( new ExcelDataRowList() );

            }
            else if ( sheets[ number ].LastRowNum >= 0 )
            {   // sheet is not empty
                for ( int i = sheets[ number ].LastRowNum; i == 0; i++ )
                    sheets[ number ].RemoveRow( sheets[ number ].GetRow( i ) );
                sheetsHeadersBool[ number ] = withHeader;
                sheetsHeaders[ number ] = new ExcelDataRowList();

            }
            // then the name or a made one
            if ( !string.IsNullOrEmpty( name ) )
                sheetsNames[ number ] = name;
            else
                sheetsNames[ number ] = $"table {number}";
            // update the local data
            ReadSheets();

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
        /// the data has to be read somewhere else.
        /// </summary>
        public void ReadSheets( )
        {
            sheets.Clear();
            sheetsHeaders.Clear();
            sheetsHeadersBool.Clear();
            for ( int i = 0; i < workbook.NumberOfSheets; i++ )
            {
                sheets.Add( workbook.GetSheetAt( i ) );
                sheetsHeaders.Add( new ExcelDataRowList() );
                sheetsHeadersBool.Add( false );

            }

        }   // end: public void ReadSheets

        /// <summary>
        /// Reads all rows of the sheet into the local
        /// version.
        /// </summary>
        /// <param name="number">number of the sheet</param>
        /// <param name="useHeader">use a header row</param>
        /// <returns>the truth of success ( wrong number? )</returns>
        public bool ReadSheetAsListString( int number, bool useHeader = false )
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
                dataListString.Add( new ExcelDataRowList( CellType.String ) );
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
                if ( sheetsHeadersBool[ number ] )
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
        public bool ReadSheetAsListDouble( int number, bool useHeader = false )
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
                dataListDouble.Add( new ExcelDataRowList( CellType.Numeric ) );
                dataListDouble[ i - deltaRow ].FromRow( sheets[ number ].GetRow( i ) );

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
                if ( sheetsHeadersBool[ number ] )
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
                dataListMixed[ i - deltaRow ].FromRow( sheets[ number ].GetRow( i ) );

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
                if ( sheetsHeadersBool[ number ] )
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
        public string DataListString_ToString( )
        {
            string text = "";
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
        public string DataListDouble_ToString( )
        {
            string text = "";
            if ( dataListString.Count > 0 )
                for ( int i = 0; i < dataListString.Count; i++ )
                    text += dataListString[ i ].ToString() + "\n";
            else
                text += "empty list... \n";

            return ( text );

        }   // end:public void DataListDouble_ToString

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

    }   // end: internal class NPOIexcel

}
