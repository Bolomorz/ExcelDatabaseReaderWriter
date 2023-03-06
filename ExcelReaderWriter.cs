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
                    worksheet = sheets[1];
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
                try
                {
                    application = new excel.Application();
                    workbook = application.Workbooks.Add();
                    sheets = workbook.Sheets;
                    worksheet = workbook.ActiveSheet;
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
        public Tuple<string, string> ReadCell(int row, int column)
        {
            string value = string.Empty;

            string message = string.Empty;

            if (opened) { 
                try
                {
                    value = worksheet.Cells[row, column].Value2.ToString();
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
        public string WriteCell(int row, int column, string value)
        {
            string ret = string.Empty;

            if (opened)
            {
                try
                {
                    worksheet.Cells[row, column].Value2 = value;
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

        /// <summary>
        /// Save file.
        /// </summary>
        public void SaveChanges()
        {
            try
            {
                workbook.SaveAs(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// save changes and dispose objects to kill all processes
        /// </summary>
        public void Quit()
        {
            workbook.Close(0);
            application.Quit();
            opened = false;
        }
    }
}
