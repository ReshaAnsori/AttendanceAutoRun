using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;

namespace AttendanceAutoRun
{
    internal class Program
    {
        static OleDbConnection con;
        static OleDbCommand cmd;
        static OleDbDataReader reader;

        static void Main(string[] args)
        {
            var ErrorStts = false;
            var timeout = 120;
            var cekTimeout = Int32.TryParse(ConfigurationManager.AppSettings["TimeOutInSeconds"], out timeout);
            timeout = cekTimeout ? timeout : 120;

            if (ConfigurationManager.AppSettings["DebugConsole"].ToLower() == "true")
                ErrorStts = true;
            try
            {
                con = new OleDbConnection();
                //con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:/Private/Coba-coba/attBackup.mdb";
                con.ConnectionString = ConfigurationManager.AppSettings.Get("AccessConnectionString");
                cmd = new OleDbCommand();
                cmd.Connection = con;

                int backDay = 1;
                var cek = Int32.TryParse(ConfigurationManager.AppSettings["BackDay"], out backDay);
                backDay = !cek ? 1 : backDay;

                // Set parameter
                //var backDay = DateTime.Today.DayOfYear;
                var whereDate = DateTime.Today.AddDays(-backDay).Date;
                var conn = ConfigurationManager.AppSettings.Get("SQLServerConnectionString");
                //var item = 0;
                var item = false;

                // Cek data existing pada stagging SQL server
                using (SqlConnection connection = new SqlConnection(conn))
                {
                    connection.Open();
                    using (var cmd = connection.CreateCommand())
                    {
                        try
                        {
                            cmd.CommandTimeout = timeout;
                            //cmd.CommandText = "Select count(*) from CHECKINOUT";
                            cmd.CommandText = "select case when count(*) > 0 then 1 else 0 end from CHECKINOUT";
                            item = (Int32)cmd.ExecuteScalar() > 0;
                        }
                        catch (Exception ex)
                        {
                            ErrorStts = true;
                            Console.WriteLine("==================================== Failed Execute ====================================");
                            Console.WriteLine("Exception 1 : " + ex.Message);
                        }

                    }
                    connection.Close();
                }

                // Query select untuk mdb, jika data stagging kosong, parameter where > tanggal gak ditambah
                var query = "SELECT * FROM CHECKINOUT";
                query += !item ? "" : " where CHECKTIME >= #" + whereDate.ToString("dd/MM/yyyy HH:mm:ss") + "#";
                
                cmd.CommandText = query;
                con.Open();
                reader = cmd.ExecuteReader();
                var datas = new DataTable();
                datas.Load(reader);

                List<DataRow> list = datas.AsEnumerable().ToList();
                Console.WriteLine("Data dari access per-tanggal " + whereDate.ToString("dd/MM/yyyy HH:mm:ss") + " : " + list.Count());

                //using (SqlConnection connection = new SqlConnection("Data Source=localhost;Initial Catalog=DB_ATTD;Integrated Security=True"))
                using (SqlConnection connection = new SqlConnection(conn))
                {
                    connection.Open();
                    // jika data stagging ada, ini di jalankan dengan menghapus data existing >= parameter date
                    if(item && list.Count > 0)
                    {
                        using (var cmd = connection.CreateCommand())
                        {
                            try
                            {
                                cmd.CommandTimeout = timeout;
                                cmd.CommandText = "DELETE FROM dbo.CHECKINOUT WHERE CHECKTIME >= @word";
                                cmd.Parameters.AddWithValue("@word", whereDate);
                                cmd.ExecuteNonQuery();

                            }
                            catch (Exception ex)
                            {
                                ErrorStts = true;
                                Console.WriteLine("==================================== Failed Execute ====================================");
                                Console.WriteLine("Exception 2 : " + ex.Message);
                            }
                        }
                    }

                    // insert data mdb ke stagging
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                    {
                        try
                        {
                            bulkCopy.DestinationTableName = "dbo.CHECKINOUT";
                            bulkCopy.BulkCopyTimeout = timeout;
                            // Write unchanged rows from the source to the destination.
                            bulkCopy.WriteToServer(datas, DataRowState.Unchanged);
                        }
                        catch (Exception ex)
                        {
                            ErrorStts = true;
                            Console.WriteLine("==================================== Failed Execute ====================================");
                            Console.WriteLine("Exception 3 : " + ex.Message);
                        }
                    }

                    // hapus semua data pada tabel ready yang date nya >= parameter tanggal
                    using (var cmd = connection.CreateCommand())
                    {
                        try
                        {
                            cmd.CommandTimeout = timeout;
                            cmd.CommandText = "DELETE FROM dbo.AttendanceMachinePolling WHERE AttendanceDate >= @word"; //WHERE AttendanceDate = @word
                            cmd.Parameters.AddWithValue("@word", whereDate);
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            ErrorStts = true;
                            Console.WriteLine("==================================== Failed Execute ====================================");
                            Console.WriteLine("Exception 4 : " + ex.Message);
                        }
                    }

                    var queryAttd = ConfigurationManager.AppSettings["MigrationQuery"].ToLower();
                    queryAttd += !item ? "" : " WHERE CHECKTIME >= @word";

                    using (var cmd = connection.CreateCommand())
                    {
                        try
                        {
                            cmd.CommandTimeout = timeout;
                            cmd.CommandText = queryAttd;
                            if (item)
                            {
                                cmd.Parameters.AddWithValue("@word", whereDate);
                            }
                            cmd.ExecuteNonQuery();

                        }
                        catch (Exception ex)
                        {
                            ErrorStts = true;
                            Console.WriteLine("==================================== Failed Execute ====================================");
                            Console.WriteLine("Exception 5 : " + ex.Message);
                        }
                    }
                }

                con.Close();
                Console.WriteLine("==================================== Success Execute ====================================");
            }
            catch (Exception e)
            {
                ErrorStts = true;
                Console.WriteLine("==================================== Failed Execute ====================================");
                Console.WriteLine("Exception 6 : " + e.Message);
            }
            if(ErrorStts)
                Console.ReadKey();
        }
    }
}
