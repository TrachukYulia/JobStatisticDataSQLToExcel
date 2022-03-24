using System;
using System.Text;
using System.Data.Common;
using System.Data.SqlClient;


namespace TransferToExcel
{
    public class DBConnection
    {
        public static SqlConnection GetDBConnection()
        {
            
            //Data Source=JULIYA\MYMSSQL;Initial Catalog=TrainTicketsDB;Integrated Security=True
            string datasource = @"JULIYA\MYMSSQL";
            string database = "TutorialDB";

            return GetDBConnection(datasource, database);
        }
        public static SqlConnection GetDBConnection(string datasource, string database)
        {

            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Integrated Security=True;";
            SqlConnection conn = new SqlConnection(connString);
            return conn;
        }
    }
}
