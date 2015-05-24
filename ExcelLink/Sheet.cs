/*
The MIT License (MIT)

Copyright (c) 2015 Sander Cox - Parallel Dimension

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLink
{
    public class Sheet : IDisposable
    {
        private Excel.Worksheet xlWorksheet;

        public String Name
        {
            get
            {
                return xlWorksheet.Name;
            }
        }

        public class RowIndexer
        {
            private Sheet sheet;
            private int row;

            internal RowIndexer(Sheet sheet, int row)
            {
                this.sheet = sheet;
                this.row = row;
            }

            public Cell this[int column]
            {
                get { return sheet.Cell(row, column); }
            }
        }

        public class CellIndexer
        {
            private Sheet sheet;

            public Cell this[int row, int column]
            {
                get
                {
                    return sheet.Cell(row, column);
                }
            }

            public RowIndexer this[int row]
            {
                get { return new RowIndexer(sheet, row); }
            }

            public CellIndexer(Sheet sheet)
            {
                this.sheet = sheet;
            }
        }

        private CellIndexer _cellIndexer;
        public CellIndexer Cells { get { return _cellIndexer; } }
        
        private Dictionary<Tuple<int, int>, Cell> CellsDict;

        public Sheet(Excel.Worksheet worksheet)
        {
            _cellIndexer = new CellIndexer(this);
            CellsDict = new Dictionary<Tuple<int, int>, Cell>();
            xlWorksheet = worksheet;
        }

        public Cell Cell(int row, int column)
        {
            var rcTuple = new Tuple<int, int>(row, column);
            if (!CellsDict.Keys.Contains(rcTuple))
            {
                CellsDict[rcTuple] = new Cell(this, row, column);
            }
            return CellsDict[rcTuple];
        }

        public void SetValue(int row, int column, String value)
        {
            xlWorksheet.Cells.set_Item(row, column, value);
        }

        public String GetValue(int row, int column)
        {
            return xlWorksheet.Cells.get_Item(row, column).Value2.ToString();
        }

        internal void UpdateCells(Excel.Range updateRange)
        {
            // row / columns of the range are offset
            Console.WriteLine("Updating cells for " + this.Name);

            for (int rowOffset = 0; rowOffset < updateRange.Rows.Count; rowOffset++)
            {
                for (int colOffset = 0; colOffset < updateRange.Columns.Count; colOffset++)
                {
                    UpdateCell(updateRange.Row + rowOffset, updateRange.Column + colOffset);
                }
            }
        }

        private void UpdateCell(int row, int col)
        {
            List<Tuple<int, int>> cellsToUpdate = new List<Tuple<int, int>>();
            cellsToUpdate.Add(new Tuple<int, int>(row, col));

            //Console.WriteLine("Start updating cell " + row + ", " + col);
            try
            {
                // throws exception is no dependents are available so we need to try catch this :(
                Excel.Range xlRange = (xlWorksheet.Cells[row, col] as Excel.Range);
                Excel.Range xlDeps = xlRange.Dependents;
                for (int i = 1; i <= xlDeps.Areas.Count; i++)
                {
                    Excel.Range xlArea = xlDeps.Areas[i];

                    for (int rowOffset = 0; rowOffset < xlArea.Rows.Count; rowOffset++)
                    {
                        for (int colOffset = 0; colOffset < xlArea.Columns.Count; colOffset++)
                        {
                            //Console.WriteLine(" adding dependent cell: " + (xlArea.Row + rowOffset) + ", " + (xlArea.Column + colOffset));
                            cellsToUpdate.Add(new Tuple<int, int>(xlArea.Row + rowOffset, xlArea.Column + colOffset));
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            foreach(var cell in cellsToUpdate)
            {
                //Console.WriteLine("Updating cell " + cell.Item1 + ", " + cell.Item2);

                if (CellsDict.Keys.Contains(cell))
                {
                    Cell c = CellsDict[cell];
                    c.Reload();
                }
            }
        }

        // Flag: Has Dispose already been called? 
        bool disposed = false;

        // Public implementation of Dispose pattern callable by consumers. 
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // Protected implementation of Dispose pattern. 
        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
            }

            // Free any unmanaged objects here. 
            //
            disposed = true;
        }
    }
}
