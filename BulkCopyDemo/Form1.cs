using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BulkCopyDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            Bulk bulk = new Bulk();
            var number = string.IsNullOrEmpty(txtNum.Text) ? 500 : Convert.ToInt32(txtNum.Text);
            var data = bulk.GenerateData(number);
            var tableName = "BULKCOPY";
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var result = bulk.InsertBulk(data, tableName);
            watch.Stop();
            TimeSpan t = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            string time = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        t.Hours,
                        t.Minutes,
                        t.Seconds,
                        t.Milliseconds);
            if (result)
            {
                lblResult.Text = string.Format("Insert {0} record complete in {1} miliseconds", data.Rows.Count, time);
            }
            else
            {
                lblResult.Text = "Insert Fail.";
            }
        }

        private void btnCount_Click(object sender, EventArgs e)
        {
            int numRows = (int)mgrDataSQL.ExecuteScalar("SELECT COUNT(1) FROM BULKCOPY;");
            lblCount.Text = numRows.ToString();
        }

        private async void btnAsync_Click(object sender, EventArgs e)
        {
            Bulk bulk = new Bulk();
            var number = string.IsNullOrEmpty(txtNum.Text) ? 500 : Convert.ToInt32(txtNum.Text);
            var data = bulk.GenerateData(number);
            var tableName = "BULKCOPY";
            var watch = System.Diagnostics.Stopwatch.StartNew();
            await mgrDataSQL.ExecuteBulkCopyAsync(data, tableName);
            watch.Stop();
            TimeSpan t = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            string time = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        t.Hours,
                        t.Minutes,
                        t.Seconds,
                        t.Milliseconds);

            lblResult.Text = string.Format("Insert {0} record complete in {1}", data.Rows.Count, time);
        }

        private async void btnMerge_Click(object sender, EventArgs e)
        {
            Random rand = new Random();
            Bulk BULK = new Bulk();
            var number = string.IsNullOrEmpty(txtNum.Text) ? 500 : Convert.ToInt32(txtNum.Text);
            var data = BULK.GenerateData(number);
            var top100 = BULK.SelectTop100();
            foreach (DataRow row in top100.Rows)
            {
                var newRow = data.NewRow();
                newRow["NAME"] = row["NAME"];
                newRow["PRICE"] = rand.Next(1, 1000000);
                data.Rows.Add(newRow);
            }

            var watch = System.Diagnostics.Stopwatch.StartNew();
            await BULK.MergeData("#TEMPBULKCOPY", "BULKCOPY", data);
            watch.Stop();
            TimeSpan t = TimeSpan.FromMilliseconds(watch.ElapsedMilliseconds);
            string time = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        t.Hours,
                        t.Minutes,
                        t.Seconds,
                        t.Milliseconds);

            lblResult.Text = string.Format("Merge {0} record complete in {1}", data.Rows.Count, time);

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Are you sure?", "Confirm delete", MessageBoxButtons.YesNo);
            if (confirm.Equals(DialogResult .Yes ))
            {
                Bulk b = new Bulk();
                b.Delete();
                MessageBox.Show("Deleted");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Bulk bulk = new Bulk();
            var result = bulk.WriteToExcel("../../Output/");
            if (result)
            {
                MessageBox.Show("Export done");
            }
            else
            {
                MessageBox.Show("Export Fail");
            }
        }
    }
}
