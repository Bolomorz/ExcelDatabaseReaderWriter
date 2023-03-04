using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;       //Library for SQL

namespace ConsoleApp1
{
    internal class DatabaseReaderWriter
    {
        private string connection;

        public DatabaseReaderWriter(string server, string database, string userid, string password)
        {
            connection = "";
            connection += "server=" + server;
            connection += ";database=" + database;
            connection += ";userid=" + userid;
            connection += ";password=" + password;
        }

        public DatabaseReaderWriter(string server, string database, string userid) 
        {
            connection = "";
            connection += "server=" + server;
            connection += ";database=" + database;
            connection += ";userid=" + userid;
        }

        /// <summary>
        /// UPDATE, INSERT INTO, DELETE commands.
        /// returns string.Empty if successfull, Exception ex if thrown.
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public string Command(string command)
        {
            string ret = string.Empty;

            var con = new MySqlConnection();
            try
            {
                con = new MySqlConnection(connection);
                con.Open();
                var cmd = new MySqlCommand();
                cmd.Connection = con;
                cmd.CommandText = command;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ret = ex.Message;
            }
            finally
            {
                con.Close();
            }
            return ret;
        }

        /// <summary>
        /// SELECT command.
        /// returns string.Empty if successfull, Exception ex if thrown.
        /// returns List of rows. Each row as a List of strings.
        /// </summary>
        /// <param name="command"></param>
        /// <param name="columns"></param>
        /// <returns></returns>
        public Tuple<string, List<List<string>>> Select(string command, int columns)
        {
            string message = string.Empty;
            List<List<string>> rows = new List<List<string>>();

            var con = new MySqlConnection();
            try
            {
                con = new MySqlConnection(connection);
                con.Open();
                var cmd = new MySqlCommand(command, con);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while(rdr.Read())
                {
                    List<string> row = new List<string>();
                    for(int i = 0; i < columns; i++)
                    {
                        row.Add(rdr[i].ToString());
                    }
                    rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            finally
            {
                con.Close();
            }

            Tuple<string, List<List<string>>> ret = new Tuple<string, List<List<string>>>(message, rows);
            return ret;
        }
    }
}
