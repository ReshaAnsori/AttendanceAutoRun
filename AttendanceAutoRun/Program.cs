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
                var conn = ConfigurationManager.AppSettings.Get("SQLServerConnectionString");
                var item = 0;

                using (SqlConnection connection = new SqlConnection(conn))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "Select count(*) from CHECKINOUT";
                        item = (Int32)cmd.ExecuteScalar();
                    }
                    connection.Close();
                }
                var query = "SELECT * FROM CHECKINOUT";
                query += item < 1 ? "" : " where CHECKTIME >= #" + whereDate.ToString("dd/MM/yyyy HH:mm:ss") + "#";
                
                cmd.CommandText = query;
                con.Open();
                reader = cmd.ExecuteReader();
                var datas = new DataTable();
                datas.Load(reader);

                //using (SqlConnection connection = new SqlConnection("Data Source=localhost;Initial Catalog=DB_ATTD;Integrated Security=True"))
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

                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "DELETE FROM dbo.AttendanceMachinePolling"; //WHERE AttendanceDate = @word
                        //cmd.Parameters.AddWithValue("@word", whereDate);
                        cmd.ExecuteNonQuery();
                    }

                    //var queryAttd = ConfigurationManager.ConnectionStrings["QueryToAttendanceMachinePolling"].ConnectionString;
                    using (var cmd = connection.CreateCommand())
                    {
                        cmd.CommandText = "insert into AttendanceMachinePolling select null as Id, SENSORID as Barcode, concat(cast(CHECKTIME as date), ' 00:00:00.000') as AttendanceDate, Cast(CHECKTIME as time(7)) as AttendanceTime, case when CHECKTYPE = 'o' then 0 else 1 end as AttendanceType, GETDATE() as CreatedDate from CHECKINOUT";
                        //cmd.Parameters.AddWithValue("@word", whereDate);
                        cmd.ExecuteNonQuery();
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

            Console.ReadKey();
        }
    }
}
