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
    public class ExcelDataRowList
    {
        // the data
        public List<object> cellData;
        public CellType cellDataType;
        public bool debugTextOn = false;
        public bool showMessageBox = false;
        // private var's
        int indexFirstCell;
        int indexLastCell;
        int numCells;
        string debugText ="Debuginfo: ";

        /// <summary>
        /// Constructor that can alter the list-type.
        /// <para/>At the moment only 'double' is the alternative to 
        /// 'string'.
        /// <para/>Debug information was added to see information about empty cells,
        /// as Excel is storing a list of special nodes per row that have their
        /// position in '.ColumnsIndex'.
        /// </summary>
        /// <param name="cellTypeToUse">Standard is 'string' or choose 'double'.</param>
        /// <param name="turnDebugTextOn">add debug information to 'ToString'</param>
        /// <param name="showMessageBoxOn">even show the debug information as 'MessageBox'</param>
        public ExcelDataRowList( CellType cellTypeToUse = CellType.String,
            bool turnDebugTextOn = false, bool showMessageBoxOn = false )
        {
            if ( cellTypeToUse == CellType.Numeric )
                cellDataType = cellTypeToUse;
            else
                cellDataType = CellType.String;
            cellData = new List<object>();
            if ( turnDebugTextOn )
                debugTextOn = true;
            if ( showMessageBoxOn )
                showMessageBox = true;

        }   // end: public ExcelDataRowList

        /// <summary>
        /// Look like a common save/load-routine:
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
                    switch ( cellDataType )
                    {
                        case CellType.String:
                            cell = row.CreateCell( i, CellType.String );
                            cell.SetCellValue( (string)cellData[ i ] );
                            break;
                        case CellType.Numeric:
                            cell = row.CreateCell( i, CellType.Numeric );
                            cell.SetCellValue( (double)cellData[ i ] );
                            break;
                        default:
                            break;

                    }   // end: switch

                }   // end: for

            }   // end; if

            return ( true );

        }   // end: public bool AsRow

        /// <summary>
        /// Look like a common save/load-routine:
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
                    switch ( cellDataType )
                    {
                        case CellType.String:
                            if ( thisCell.CellType != CellType.String )
                            {
                                MessageBox.Show( $"This is not a string-type cell ! Index: {thisCell.ColumnIndex}" );
                                return ( false );
                            }
                            cellData.Add( (object)thisCell.StringCellValue );
                            break;
                        case CellType.Numeric:
                            if ( thisCell.CellType != CellType.Numeric )
                            {
                                MessageBox.Show( $"This is not a numeric-type cell ! Index: {thisCell.ColumnIndex}" );
                                return ( false );
                            }
                            cellData.Add( (object)thisCell.NumericCellValue );
                            break;
                        default:
                            break;

                    }   // end: switch

                }
                else
                {
                    var newCell = row.CreateCell( j, cellDataType );
                    switch ( cellDataType )
                    {
                        case CellType.String:
                            newCell.SetCellValue( "" );
                            cellData.Add( (object)newCell.StringCellValue );
                            break;
                        case CellType.Numeric:
                            newCell.SetCellValue( double.NaN );
                            cellData.Add( (object)newCell.NumericCellValue );
                            break;
                        default:
                            break;

                    }   // end: switch

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

    }   // end: public class ExcelDataRowList

}
