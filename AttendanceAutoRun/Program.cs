using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;

namespace AttendanceAutoRun
{
    internal class Program
    {
        static OleDbConnection con;
        static OleDbCommand cmd;
        static OleDbDataReader reader;

        static void Main(string[] args)
        {
            try
            {
                con = new OleDbConnection();
                //con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:/Private/Coba-coba/attBackup.mdb";
                con.ConnectionString = ConfigurationManager.AppSettings.Get("AccessConnectionString");
                cmd = new OleDbCommand();
                cmd.Connection = con;
                var whereDate = DateTime.Today.Date;

                var query = "SELECT * FROM CHECKINOUT";
                //var query = "SELECT * FROM CHECKINOUT where CHECKTIME >= #" + whereDate.ToString("dd/MM/yyyy HH:mm:ss") + "#";
                var finalQuery = string.Empty;

                using (var reader = new StringReader(query))
                {
                    finalQuery = reader.ReadLine();
                    if (string.IsNullOrEmpty(finalQuery))
                    {
                        finalQuery = "SELECT * FROM CHECKINOUT";
                    }
                }

                //cmd.CommandText = query;
                cmd.CommandText = finalQuery;
                con.Open();
                reader = cmd.ExecuteReader();
                var datas = new DataTable();
                datas.Load(reader);

                //var conn = ConfigurationManager.ConnectionStrings["SQLServerConnectionString"].ConnectionString;
                //using (SqlConnection connection = new SqlConnection("Data Source=localhost;Initial Catalog=DB_ATTD;Integrated Security=True"))
                var conn = ConfigurationManager.AppSettings.Get("SQLServerConnectionString");
                using (SqlConnection connection = new SqlConnection(conn))
                {
                    connection.Open();

                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "DELETE FROM dbo.CHECKINOUT WHERE CHECKTIME >= @word";
                        cmd.Parameters.AddWithValue("@word", whereDate);
                        cmd.ExecuteNonQuery();
                    }

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        bulkCopy.DestinationTableName = "dbo.CHECKINOUT";

                        try
                        {
                            // Write unchanged rows from the source to the destination.
                            bulkCopy.WriteToServer(datas, DataRowState.Unchanged);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("==================================== Filed Execute ====================================");
                            Console.WriteLine("Exception : " + ex.Message);
                        }
                    }
                }

                con.Close();
                Console.WriteLine("==================================== Success Execute ====================================");
            }
            catch (Exception e)
            {
                Console.WriteLine("==================================== Filed Execute ====================================");
                Console.WriteLine("Exception : " + e.Message);
            }
        }
    }
}
