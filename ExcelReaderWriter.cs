using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;           //Library for Marshal
using excel = Microsoft.Office.Interop.Excel;   //Library for Excel

namespace ConsoleApp1
{
    internal class ExcelReaderWriter
    {
        excel.Application application;
        excel.Workbook workbook;
        excel.Sheets sheets;
        excel.Worksheet worksheet;
        bool opened;
        string path;
        public ExcelReaderWriter(string filepath)
        {
            if (File.Exists(filepath))
            {
                try
                {
                    application = new excel.Application();
                    workbook = application.Workbooks.Open(filepath);
                    sheets = workbook.Sheets;
                    worksheet = sheets[0];
                    opened = true;
                    path = filepath;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// read from input cell.
        /// returns value in cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public string ReadCell(string cell)
        {
            string ret = string.Empty;
            if (opened) { 
                try
                {
                    ret = worksheet.Cells[cell].Value2.ToString();
                }
                catch (Exception ex)
                {
                 Console.WriteLine(ex.Message);
                }
            }
            return ret;
        }

        /// <summary>
        /// write input value to input cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public void WriteCell(string cell, string value)
        {
            if (opened)
            {
                try
                {
                    worksheet.Cells[cell].Value2 = value;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// save changes and dispose objects to kill all processes
        /// </summary>
        public void SaveAndDispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(sheets);

            object misValue = System.Reflection.Missing.Value;

            workbook.AcceptAllChanges();
            workbook.SaveAs2(path, excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            workbook.Close(0);
            Marshal.FinalReleaseComObject(workbook);

            application.Quit();
            Marshal.FinalReleaseComObject(application);
        }
    }
}
