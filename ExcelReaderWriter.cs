using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;           //Library for Marshal
using excel = Microsoft.Office.Interop.Excel;   //Library for Excel

namespace ConsoleApp1
{
    public class ExcelReaderWriter
    {
        excel.Application? application;
        excel.Workbook? workbook;
        excel.Sheets? sheets;
        excel.Worksheet? worksheet;
        bool opened;
        string path;
        public ExcelReaderWriter(string filepath)
        {
            path = filepath;
            opened = false;
        }

        /// <summary>
        /// read from input cell.
        /// returns value in cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public Tuple<string?, object?> ReadCell(int row, int column)
        {
            EstablishConnection();
            
            object? value = null;

            string? message = null;

            if (opened) { 
                try
                {
                    value = worksheet.Cells[row, column].Value2;
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

            Tuple<string?, object?> ret = new Tuple<string?, object?>(message, value);

            SaveChanges();
            Quit();

            return ret;
        }

        /// <summary>
        /// write input value to input cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public string? WriteCell(int row, int column, object value)
        {
            EstablishConnection();
            
            string? ret = null;

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

            SaveChanges();
            Quit();

            return ret;
        }

        /// <summary>
        /// Establish connection to excel application with file in path.
        /// </summary>
        private void EstablishConnection()
        {
            if (File.Exists(path))
            {
                try
                {
                    application = new excel.Application();
                    workbook = application.Workbooks.Open(path);
                    sheets = workbook.Sheets;
                    worksheet = sheets[1];
                    opened = true;
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
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// Save file.
        /// </summary>
        private void SaveChanges()
        {
            try
            {
                if(File.Exists(path))
                {
                    workbook.Save();
                }
                else
                {
                    workbook.SaveAs2(path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        /// <summary>
        /// Quit Process.
        /// </summary>
        private void Quit()
        {
            workbook.Close(0);
            application.Quit();
            opened = false;
        }
    }
}
