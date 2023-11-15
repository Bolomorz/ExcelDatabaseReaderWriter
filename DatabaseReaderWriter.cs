using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;       //Library for SQL

namespace ConsoleApp1
{
    public class DatabaseReaderWriter
    {
        public struct Query
        {
            public string? errormessage;
            public List<List<object>>? rows;
        }
        
        private string connection;

        public DatabaseReaderWriter(string _connection) 
        {
            connection = _connection;
        }

        /// <summary>
        /// Non Query commands.
        /// returns string.Empty if successfull, Exception ex if thrown.
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public string? CommandNonQuery(string command)
        {
            string? ret = null;

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
        /// QUERY command.
        /// returns string.Empty if successfull, Exception ex if thrown.
        /// returns List of rows. Each row as a List of strings.
        /// </summary>
        /// <param name="command"></param>
        /// <returns></returns>
        public Query CommandQuery(string command)
        {
            string? message = null;
            List<List<object>> rows = new List<List<object>>();

            var con = new MySqlConnection();
            try
            {
                con = new MySqlConnection(connection);
                con.Open();
                var cmd = new MySqlCommand(command, con);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while(rdr.Read())
                {
                    List<object> row = new List<object>();
                    for(int i = 0; i < rdr.FieldCount; i++)
                    {
                        row.Add(rdr[i]);
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

            Query ret = new Query() { errormessage = message, rows = rows};
            return ret;
        }
    }
}
