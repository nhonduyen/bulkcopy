using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace BulkCopyDemo
{
    public class mgrDataSQL
    {
        public static string connStr = ConfigurationManager.ConnectionStrings["cnnString"].ConnectionString;

        public static bool ExecuteBulkCopy(DataTable data, string destinationTable)
        {
            var finished = false;
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                connect.Open();
                using (var bulk = new SqlBulkCopy(connect))
                {
                    bulk.DestinationTableName = destinationTable;
                    bulk.WriteToServer(data);
                    finished = true;
                }
            }
            return finished;
        }
        public static async Task ExecuteBulkCopyAsync(DataTable data, string destinationTable)
        {
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                connect.Open();
                using (var bulk = new SqlBulkCopy(connect))
                {
                    bulk.BulkCopyTimeout = 60*5;
                    bulk.DestinationTableName = destinationTable;
                    await bulk.WriteToServerAsync(data);
                }
            }
        }

        public static async Task ExecuteNonQueryAsync(string sql, Dictionary<string, object> param = null)
        {
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                try
                {
                    connect.Open();
                    using (SqlCommand command = new SqlCommand(sql, connect))
                    {
                    
                        if (param != null)
                        {
                            foreach (var item in param)
                            {
                                command.Parameters.AddWithValue(item.Key.ToString(), item.Value);
                            }
                        }

                       await command.ExecuteNonQueryAsync();
                   
                    }
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    connect.Dispose();
                }
            }
        }

        public static int ExecuteNonQuery(string sql, Dictionary<string, object> param = null)
        {
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                try
                {
                    connect.Open();
                    using (SqlCommand command = new SqlCommand(sql, connect))
                    {
                        SqlTransaction transaction;
                        transaction = connect.BeginTransaction();
                        command.Transaction = transaction;
                        command.Notification = null;
                        if (param != null)
                        {
                            foreach (var item in param)
                            {
                                command.Parameters.AddWithValue(item.Key.ToString(), item.Value);
                            }
                        }

                        int count = command.ExecuteNonQuery();
                        transaction.Commit();

                        return count;
                    }
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    connect.Dispose();
                }
            }
        }

        public static DataTable ExecuteReader(string sql, Dictionary<string, object> param = null)
        {
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                DataTable dtb = new DataTable();

                try
                {
                    connect.Open();
                    using (SqlCommand command = new SqlCommand(sql, connect))
                    {
                        if (param != null)
                        {
                            foreach (var item in param)
                            {
                                command.Parameters.AddWithValue(item.Key.ToString(), item.Value);
                            }
                        }
                        command.CommandTimeout = 120;
                        SqlDataReader reader = command.ExecuteReader();
                        dtb.Load(reader);
                    }
                    return dtb;
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    connect.Dispose();
                }
            }

        }
        public static object ExecuteScalar(string sql, Dictionary<string, object> param = null)
        {
            using (SqlConnection connect = new SqlConnection(connStr))
            {
                try
                {
                    using (SqlCommand command = new SqlCommand(sql, connect))
                    {
                        command.CommandTimeout = 120;
                        if (param != null)
                        {
                            foreach (var item in param)
                            {
                                command.Parameters.AddWithValue(item.Key.ToString(), item.Value);
                            }
                        }
                        connect.Open();
                        object result = command.ExecuteScalar();
                        return result;
                    }
                }
                catch (SqlException ex)
                {
                    throw ex;
                }
                finally
                {
                    connect.Dispose();
                }
            }
        }

    }
}
