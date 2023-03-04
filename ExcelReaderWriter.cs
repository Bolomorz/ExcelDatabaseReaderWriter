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
            else
            {
                Console.WriteLine("File does not exist!");
                opened = false;
            }
        }

        /// <summary>
        /// read from input cell.
        /// returns value in cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public Tuple<string, string> ReadCell(string cell)
        {
            string value = string.Empty;

            string message = string.Empty;

            if (opened) { 
                try
                {
                    value = worksheet.Cells[cell].Value2.ToString();
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            }
            else
            {
                message = "App closed!";
            }

            Tuple<string, string> ret = new Tuple<string, string>(message, value);

            return ret;
        }

        /// <summary>
        /// write input value to input cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public string WriteCell(string cell, string value)
        {
            string ret = string.Empty;

            if (opened)
            {
                try
                {
                    worksheet.Cells[cell].Value2 = value;
                    SaveChanges();
                }
                catch (Exception ex)
                {
                    ret = ex.Message;
                }
            }
            else
            {
                ret = "App closed!";
            }

            return ret;
        }

        public void SaveChanges()
        {
            object misValue = System.Reflection.Missing.Value;

            workbook.AcceptAllChanges();
            workbook.SaveAs2(path, excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

        }

        /// <summary>
        /// save changes and dispose objects to kill all processes
        /// </summary>
        public void QuitAndDispose()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(sheets);

            workbook.Close(0);
            Marshal.FinalReleaseComObject(workbook);

            application.Quit();
            Marshal.FinalReleaseComObject(application);

            opened = false;
        }
    }
}
