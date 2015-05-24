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
    public class Workbook : IDisposable
    {
        private Excel.Workbook xlWorkbook;
        public Dictionary<String, Sheet> SheetsDict { get; set; }

        public class SheetIndexer
        {
            private Workbook wb;

            public Sheet this[int index]
            {
                get
                {
                    return wb.Sheet(index);
                }
            }

            public Sheet this[String name]
            {
                get
                {
                    return wb.Sheet(name);
                }         
            }

            public SheetIndexer(Workbook wb)
            {
                this.wb = wb;
            }
        }

        private SheetIndexer _sheets;
        public SheetIndexer Sheets { get { return _sheets; } }

        protected Workbook()
        {
            SheetsDict = new Dictionary<String, Sheet>();
            _sheets = new SheetIndexer(this);
        }

        public Workbook(String filepath, bool makeVisible = true) : this()
        {

            if (!System.IO.File.Exists(filepath))
            {
                throw new System.IO.FileNotFoundException("Excel document not found", filepath);
            }

            xlWorkbook = System.Runtime.InteropServices.Marshal.BindToMoniker(filepath) as Excel.Workbook;

            if (xlWorkbook == null)
            {
                throw new System.Exception("Could not open Excel with the file!");
            }

            //Excel.Worksheet xlSheet = xlWorkBook.Sheets[1];
            //Excel.Range xlRange = xlSheet.Cells[1, 2];
            //xlSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            //Binding b = new Binding("Value2");
            //b.Source = xlRange;
            //b.Mode = BindingMode.TwoWay;
            //excelFirstname.SetBinding(TextBox.TextProperty, b);
            //xlApp.ScreenUpdating = true;
            //xlApp.Visible = true;

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            ShowWorkbook(makeVisible);
        }

        public Workbook(Excel.Workbook workbook, bool makeVisible = false) : this()
        {
            xlWorkbook = workbook;

            ShowWorkbook(makeVisible);
        }

        private void ShowWorkbook(bool makeVisible)
        {
            Excel.Application xlApp = xlWorkbook.Application;
            if (makeVisible)
            {
                xlApp.Visible = true;
            }

            if (xlWorkbook.Windows.Count == 0)
            {
                xlWorkbook.NewWindow();
            }
            if (xlWorkbook.Windows.Count > 0)
            {
                xlWorkbook.Windows[1].Visible = true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            xlWorkbook.SheetChange += xlWorkbook_SheetChange;
        }

        void xlWorkbook_SheetChange(object Sh, Excel.Range Target)
        {
            if (Sh is Excel.Worksheet)
            {
                Excel.Worksheet xlSheet = Sh as Excel.Worksheet;
                if (SheetsDict.Keys.Contains(xlSheet.Name))
                {
                    Sheet sheet = SheetsDict[xlSheet.Name];
                    sheet.UpdateCells(Target);
                }
            }
        }

        public Sheet Sheet(String sheetName)
        {
            if (!SheetsDict.Keys.Contains(sheetName))
            {
                bool found = false;
                for (int i = 1; i <= xlWorkbook.Worksheets.Count && !found; i++)
                {
                    Excel.Worksheet xlSheet = xlWorkbook.Worksheets[i];

                    if (xlSheet != null)
                    {
                        if (xlSheet.Name == sheetName)
                        {
                            SheetsDict[sheetName] = new Sheet(xlSheet);
                            found = true;
                        }
                        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                    }
                }

                if (!found)
                {
                    return null;
                }
            }

            return SheetsDict[sheetName];
        }

        public Sheet Sheet(int sheetIndex)
        {
            if (sheetIndex < 1 || sheetIndex > xlWorkbook.Worksheets.Count)
                return null;

            Excel.Worksheet xlSheet = xlWorkbook.Worksheets[sheetIndex];
            if (!SheetsDict.Keys.Contains(xlSheet.Name))
            {
                SheetsDict[xlSheet.Name] = new Sheet(xlSheet);
            }
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);

            return SheetsDict[xlSheet.Name];
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
                foreach (var sheet in SheetsDict)
	            {
                    sheet.Value.Dispose();
	            }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            }

            // Free any unmanaged objects here. 
            //
            disposed = true;
        }
    }
}
