using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using OfficeOpenXml;

namespace BulkCopyDemo
{
    public class Bulk
    {
        public DataTable GenerateData(int numRecord = 500)
        {
            DataTable tbResult = new DataTable();
            tbResult.Columns.Add(new DataColumn("NAME", typeof(string)));
            tbResult.Columns.Add(new DataColumn("PRICE", typeof(int)));
            Random rand = new Random();
            for (int i = 0; i < numRecord; i++)
            {
                var r = tbResult.NewRow();
                r["NAME"] = Path.GetRandomFileName().Replace(".", "");
                r["PRICE"] = rand.Next(1, 10000);
                tbResult.Rows.Add(r);
            }
            return tbResult;
        }
        public bool InsertBulk(DataTable data, string destinationTable)
        {
            return mgrDataSQL.ExecuteBulkCopy(data, destinationTable);
        }
        public int Delete()
        {
            var sql = "TRUNCATE TABLE BULKCOPY;";
            return mgrDataSQL.ExecuteNonQuery(sql);
        }
        public DataTable SelectPaging(int start = 1, int end = 10000)
        {
            var sql = "SELECT NAME,PRICE FROM(SELECT ROW_NUMBER() OVER (order by name) AS ROWNUM, * FROM BULKCOPY) as u  WHERE   RowNum BETWEEN @start  AND @end ORDER BY RowNum;";
            var param = new Dictionary<string, object>();
            param.Add("@start", start);
            param.Add("@end", end);
            return mgrDataSQL.ExecuteReader(sql, param);
        }
        public bool WriteToExcel(string path)
        {
            var result = false;
            var filename = DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            var count = (int)mgrDataSQL.ExecuteScalar("SELECT COUNT(1) FROM BULKCOPY");

            using (var package = new ExcelPackage(new FileInfo(path + "//" + filename)))
            {
               
                if (count <= 1000000)
                {
                    var ws = package.Workbook.Worksheets.Add("Export");
                    var data = mgrDataSQL.ExecuteReader("SELECT * FROM BULKCOPY;");
                    ws.Cells["A1"].LoadFromDataTable(data, true);
                    package.Save();
                    result = true;
                }
                else
                {
                    var n = count / 1000000;
                    for (int i = 0; i <= n; i++)
                    {
                        var ws = package.Workbook.Worksheets.Add("Export"+i);
                        package.Workbook.Worksheets.MoveToStart("Export" + i);
                        for (int j = 0; j < i * 1000000 + 1000000; j++)
                        {
                            var cell = j == 0 ? 1 : j;
                            var end = j + 10000;
                            var dtb = SelectPaging(j, end);
                            ws.Cells["A" + cell.ToString()].LoadFromDataTable(dtb, true);
                            package.Save();
                            j += end;
                        }
                    }
                   
                    result = true;
                }
            }
            return result;
        }
        public Task InsertBulkAsync(DataTable data, string destinationTable)
        {
            return mgrDataSQL.ExecuteBulkCopyAsync(data, destinationTable);
        }
        public DataTable SelectTop100()
        {
            var sql = "SELECT TOP 100 * FROM BULKCOPY ORDER BY NAME;";
            return mgrDataSQL.ExecuteReader(sql);
        }

        public async Task MergeData(string tempTableName, string destinationTable, DataTable data, int timeout = 120)
        {
            using (var connect = new SqlConnection(mgrDataSQL.connStr))
            {
                connect.Open();
                var command = new SqlCommand(string.Format(@"CREATE TABLE {0}(NAME NCHAR(50), PRICE INT);", tempTableName), connect);
                command.CommandTimeout = timeout; // 120 = 2 minute
                command.ExecuteNonQuery();

                using (var bulk = new SqlBulkCopy(connect))
                {
                    bulk.DestinationTableName = tempTableName;
                    await bulk.WriteToServerAsync(data);
                }
                var sql = string.Format(@"
MERGE INTO {0} AS TARGET USING {1} AS SOURCE ON TARGET.NAME=SOURCE.NAME
WHEN MATCHED THEN UPDATE SET TARGET.NAME=SOURCE.NAME, TARGET.PRICE=SOURCE.PRICE
WHEN NOT MATCHED THEN INSERT(NAME, PRICE) VALUES(SOURCE.NAME, SOURCE.PRICE);
", destinationTable, tempTableName);
                command.CommandText = sql;
                await command.ExecuteNonQueryAsync();
                command.CommandText = string.Format(@"DROP TABLE {0};", tempTableName);
                var t1 = command.ExecuteNonQuery();
            }

        }

        private List<DataTable> CloneTable(DataTable tableToClone, int countLimit)
        {
            List<DataTable> tables = new List<DataTable>();
            int count = 0;
            DataTable copyTable = null;
            foreach (DataRow dr in tableToClone.Rows)
            {
                if ((count++ % countLimit) == 0)
                {
                    copyTable = new DataTable();
                    // Clone the structure of the table.
                    copyTable = tableToClone.Clone();
                    // Add the new DataTable to the list.
                    tables.Add(copyTable);
                }
                // Import the current row.
                copyTable.ImportRow(dr);
                dr.Delete();
            }
            return tables;
        }
    }
}
