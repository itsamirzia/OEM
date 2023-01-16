using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace OnBase_Export_Management
{
    
    public static class db
    {
        static string dbCon = ConfigurationManager.ConnectionStrings["dbCon"].ToString();
        /// <summary>
        /// Execute Non Query
        /// </summary>
        /// <param name="query"></param>
        /// <returns>Boolean</returns>
        public static bool ExecuteNonQuery(string query)
        {
            bool success = false;
            SqlConnection con = null;
            try
            {
                con = new SqlConnection(dbCon);
                SqlCommand cmd = new SqlCommand(query, con);  
                con.Open();
                if (con.State != ConnectionState.Open)
                {
                    return false;
                }
                cmd.ExecuteNonQuery();
                success = true;
                
            }
            catch (Exception e)
            {
                Console.WriteLine("OOPs, something went wrong." + e);
                success = false;
            }
            // Closing the connection  
            finally
            {
                con.Close();
            }
        
            return success;
        }
        /// <summary>
        /// Check if data has rows
        /// </summary>
        /// <param name="query"></param>
        /// <returns>Boolean</returns>
        public static bool HasDataRows(string query)
        {
            DataTable dt = new DataTable();
            SqlConnection con = null;
            try
            {
                con = new SqlConnection(dbCon);
                con.Open();
                if (con.State != ConnectionState.Open)
                {
                    return false;
                }
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = query;
                cmd.CommandTimeout = 60;
                cmd.Connection = con;
                SqlDataReader sdr = cmd.ExecuteReader();
                if (sdr.HasRows)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// Execute SQL Query
        /// </summary>
        /// <param name="query"></param>
        /// <returns>DataTable</returns>
        public static DataTable ExecuteSQLQuery(string query)
        {
            DataTable dt = new DataTable();
            SqlConnection con = null;
            try
            {
                con = new SqlConnection(dbCon);
                con.Open();
                if (con.State != ConnectionState.Open)
                {
                    return null;
                }
                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = query;
                cmd.CommandTimeout = 60;
                cmd.Connection = con;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(dt);
                return dt;
            }
            catch
            {
                return dt;
            }
        }
    }
}
