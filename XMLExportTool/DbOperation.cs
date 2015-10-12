using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace XMLExportTool
{
    class DbOperation
    {
        protected SqlConnection Connection;
        protected SqlDataReader reader;
        public bool Connect(string serverIP,string dbName,string userName,string password)
        {
            string strDataBase = "Server="+serverIP+";DataBase="+dbName+";Uid="+userName+";pwd="+password+";";
            Connection = new SqlConnection(strDataBase);
            try
            {
                Connection.Open();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
        public SqlDataReader GetDataReader(string sqlStatement)
        {
            try
            {
                SqlCommand cmd = new SqlCommand(sqlStatement, Connection);
                reader = cmd.ExecuteReader();
            }
            catch (Exception)
            {
                return null;
            }
            return reader;
        }
        public int ExecuteNonQuery(string sqlStatement)
        {
            SqlCommand cmd = new SqlCommand(sqlStatement, Connection);
            return cmd.ExecuteNonQuery();
        }
        public void Close()
        {
            try
            {
                reader.Close();
                Connection.Close();
            }
            catch (Exception e)
            {
            }
        }
    }
}
