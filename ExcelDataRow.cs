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

using Microsoft.VisualBasic;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIwrap
{
    /// <summary>
    /// A class for mixed data rows.
    /// An opportunity for filling data in from inside Excel.
    /// </summary>
    public class ExcelDataRow
    {
        // Erstellt ab: 08.02.2024
        // letzte Änderung: 12.09.24
        Version version = new Version("1.0.1");
        // the data - should be public
        public int exampleIntNumber;
        public double exampleDoubleNumber;
        public string exampleText;

        /// <summary>
        /// Example for a mixed row of data cells. No debug information
        /// needed for that usage, as there is no sense in empty cells.
        /// </summary>
        public ExcelDataRow( )
        {
            exampleIntNumber = 0;
            exampleDoubleNumber = 0.13;
            exampleText = "this is example's data";

        }   // end: public ExcelDataRow

        /// <summary>
        /// Look like a common save/load-routine:
        /// first_in will be first_out.
        /// <para>This is just an example and you can see the its
        /// taking care of the example's data only.</para>
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
            // in example i use the given vars -> to be changed for real usage
            cell = row.CreateCell( 0, CellType.Numeric );
            cell.SetCellValue( exampleIntNumber );
            cell = row.CreateCell( 1, CellType.Numeric );
            cell.SetCellValue( exampleDoubleNumber );
            cell = row.CreateCell( 2, CellType.String );
            cell.SetCellValue( exampleText );

            return ( true );

        }   // end: public bool AsRow

        /// <summary>
        /// Look like a common save/load-routine:
        /// first_in will be first_out.
        /// <para>This is just an example and you can see the 'switch'
        /// taking care of the example's data only.</para>
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

            for ( int j = 0; j < row.LastCellNum; j++ )
            {
                if ( row.GetCell( j ) != null )
                {
                    switch ( j )
                    {
                        case 0:
                            exampleIntNumber = (int)row.GetCell( j ).NumericCellValue;
                            break;
                        case 1:
                            exampleDoubleNumber = row.GetCell( j ).NumericCellValue;
                            break;
                        case 2:
                            exampleText = row.GetCell( j ).StringCellValue;
                            break;
                        default:
                            break;

                    }   // end: switch

                }
                else
                    return ( false );

            }   // end: for


            return ( true );

        }   // end: public bool FromRow

        /// <summary>
        /// Delivers a representation of the vars as string.
        /// </summary>
        /// <returns>the message</returns>
        override
        public string ToString( )
        {

            string text = "Datarow: [ " +
                $"exampleIntNumber = {exampleIntNumber}, " +
                $"exampleDoubleNumber = {exampleDoubleNumber}, " +
                $"exampleText = {exampleText} ]";
            return ( text );

        }   // end: public string ToString

    }   // end: public class ExcelDataRow

}
