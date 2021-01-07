using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace workExcel
{
    class Program
    {
        static void Main(string[] args)
        {
           try
           {
                using ExcelHelper helper = new ExcelHelper();
                if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "TestLT.xlsx")))
                {
                    helper.Set(column: "A", row: 1, data: "ldaslfd");
                    helper.Set(column: "B", row: 1, data: DateTime.Now);

                    helper.Save();
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }

    class ExcelHelper : IDisposable
    {
        private Application _excel;
        private Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal bool Set(string column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close(); 
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
