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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace NPOIwrap
{
    /// <summary>
    /// A class for a list of row data, to be
    /// adapted to the special cases in other software.
    /// <para/>
    /// Meaning: the data type of the list is
    /// a parameter for the constructor.
    /// <para/>
    /// Internally i use 'string' and 'double' in the list. I see no other
    /// celltype to be usable without mix as a row.
    /// </summary>
    public class ExcelDataRowListString
    {
        // Erstellt ab: 12.02.2024
        // letzte Änderung: 12.09.24
        Version version = new Version("1.0.2");

        /// <summary>
        /// list of the cells in the given row
        /// </summary>
        public List<string> cellData;
        /// <summary>
        /// use one of them [ CellType.Numeric, CellType.String ]
        /// </summary>
        public CellType cellDataType = CellType.String;
        /// <summary>
        ///  debug information
        /// </summary>
        public bool debugTextOn = false;
        /// <summary>
        /// message boxes to be seen
        /// </summary>
        public bool showMessageBox = false;
        
        // private var's
        int indexFirstCell;
        int indexLastCell;
        int numCells;
        string debugText ="Debuginfo: ";

        /// <summary>
        /// Constructor.
        /// <para/>At the moment only 'double' is the alternative to 
        /// 'string'.
        /// <para/>Debug information was added to see information about empty cells,
        /// as Excel is storing a list of special nodes per row that have their
        /// position in '.ColumnsIndex'.
        /// </summary>
        /// <param name="turnDebugTextOn">add debug information to 'ToString'</param>
        /// <param name="showMessageBoxOn">even show the debug information as 'MessageBox'</param>
        public ExcelDataRowListString( bool turnDebugTextOn = false, 
            bool showMessageBoxOn = false )
        {
            cellData = new List<string>();
            if ( turnDebugTextOn )
                debugTextOn = true;
            if ( showMessageBoxOn )
                showMessageBox = true;

        }   // end: public ExcelDataRowListString

        /// <summary>
        /// Looks like a common save/load-routine:
        /// first_in will be first_out.
        /// </summary>
        /// <param name="row">the row to be used</param>
        /// <returns>bool: the success</returns>
        public bool AsRow( ref IRow row )
        {
            ICell cell;
            // clear all cells first
            if ( row.LastCellNum > 0 )
                for ( int i = ( row.LastCellNum - 1 ); i >= 0; i-- )
                    row.RemoveCell( row.GetCell( i ) );
            // cellDataType is used to cast the data from object-list
            if ( cellData.Count > 0 )
            {
                for ( int i = 0; i < cellData.Count; i++ )
                {
                    cell = row.CreateCell( i, CellType.String );
                    cell.SetCellValue( cellData[ i ] );

                }   // end: for

            }   // end: if

            return ( true );

        }   // end: public bool AsRow

        /// <summary>
        /// Looks like a common save/load-routine:
        /// first_in will be first_out.
        /// </summary>
        /// <param name="row">the row to be used</param>
        /// <returns>bool: the success</returns>
        public bool FromRow( IRow row )
        {
            // read the example data or result is false
            if ( row == null )
                return ( false );
            if ( row.Cells.All( d => d.CellType == CellType.Blank ) )
                return ( false );

            cellData.Clear();
            indexFirstCell = row.GetCell( row.FirstCellNum ).ColumnIndex;
            indexLastCell = row.GetCell( row.LastCellNum - 1 ).ColumnIndex;
            numCells = 0;
            debugText += $" [ (first -/last cell): {indexFirstCell}, {indexLastCell} ] " +
                $"-> (# of cells):{row.Cells.Count} Cells.";

            for ( int j = 0; j <= indexLastCell; j++ )
            {
                var thisCell = row.GetCell( j );
                if ( thisCell != null )
                {
                    numCells++;
                    if ( thisCell.CellType != CellType.String )
                    {
                        MessageBox.Show( $"This is not a string-type cell ! Index: {thisCell.ColumnIndex}" );
                        return ( false );
                    
                    }
                    cellData.Add( thisCell.StringCellValue );

                }
                else
                {
                    var newCell = row.CreateCell( j, cellDataType );
                    newCell.SetCellValue( "" );
                    cellData.Add( newCell.StringCellValue );

                }   // end: null-test

            }   // for-loop
            debugText += $" original cells = {numCells} ";
            if ( showMessageBox )
                MessageBox.Show( debugText );

            return ( true );

        }   // end: public bool FromRow

        /// <summary>
        /// Delivers a representation of the list as string.
        /// </summary>
        /// <returns>the message</returns>
        override
        public string ToString( )
        {

            if ( cellData.Count == 0 )
                return ( "Datarow: empty" );
            else
            {
                string text = "Datarow: [ ";
                for ( int i = 0; i < cellData.Count; i++ )
                {
                    text += $"'{cellData[ i ]}'";
                    if ( i < ( cellData.Count - 1 ) )
                        text += ", ";

                }
                text += " ] ";
                if ( debugTextOn )
                    text += debugText;
                return ( text );

            }

        }   // end: public string ToString

        /// <summary>
        /// Returns the list of data as 'string[]'
        /// </summary>
        /// <returns>the string-array</returns>
        public string[] CellDataAsArray( )
        {
            string[] temp = new string[ cellData.Count ];
            for ( int i = 0; i < cellData.Count; ++i )
                temp[ i ] = cellData[ i ];

            return ( temp );

        }   // end: CellDataAsArray

        /// <summary>
        /// Stores the 'string's in the cellData-list
        /// </summary>
        /// <param name="strings">array of 'string's</param>
        public void ArrayToCellData( string[] strings )
        {
            cellData.Clear();
            foreach ( string s in strings )
                cellData.Add( s );

        }   // end: ArrayToCellData

    }   // end: public class ExcelDataRowListString

}   // end: namespace NPOIwrap
