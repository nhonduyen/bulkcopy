namespace BulkCopyDemo
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnInsert = new System.Windows.Forms.Button();
            this.btnMerge = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblResult = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNum = new System.Windows.Forms.TextBox();
            this.btnCount = new System.Windows.Forms.Button();
            this.lblCount = new System.Windows.Forms.Label();
            this.btnAsync = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnInsert
            // 
            this.btnInsert.Location = new System.Drawing.Point(264, 22);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(75, 23);
            this.btnInsert.TabIndex = 0;
            this.btnInsert.Text = "Bulk Insert";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // btnMerge
            // 
            this.btnMerge.Location = new System.Drawing.Point(13, 64);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new System.Drawing.Size(75, 23);
            this.btnMerge.TabIndex = 1;
            this.btnMerge.Text = "Merge";
            this.btnMerge.UseVisualStyleBackColor = true;
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 223);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Result";
            // 
            // lblResult
            // 
            this.lblResult.AutoSize = true;
            this.lblResult.Location = new System.Drawing.Point(56, 223);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(35, 13);
            this.lblResult.TabIndex = 3;
            this.lblResult.Text = "label2";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Number";
            // 
            // txtNum
            // 
            this.txtNum.Location = new System.Drawing.Point(76, 24);
            this.txtNum.Name = "txtNum";
            this.txtNum.Size = new System.Drawing.Size(168, 20);
            this.txtNum.TabIndex = 5;
            // 
            // btnCount
            // 
            this.btnCount.Location = new System.Drawing.Point(117, 64);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(75, 23);
            this.btnCount.TabIndex = 6;
            this.btnCount.Text = "Count";
            this.btnCount.UseVisualStyleBackColor = true;
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // lblCount
            // 
            this.lblCount.AutoSize = true;
            this.lblCount.Location = new System.Drawing.Point(16, 190);
            this.lblCount.Name = "lblCount";
            this.lblCount.Size = new System.Drawing.Size(35, 13);
            this.lblCount.TabIndex = 7;
            this.lblCount.Text = "label3";
            // 
            // btnAsync
            // 
            this.btnAsync.Location = new System.Drawing.Point(353, 22);
            this.btnAsync.Name = "btnAsync";
            this.btnAsync.Size = new System.Drawing.Size(75, 23);
            this.btnAsync.TabIndex = 8;
            this.btnAsync.Text = "Bulk Async";
            this.btnAsync.UseVisualStyleBackColor = true;
            this.btnAsync.Click += new System.EventHandler(this.btnAsync_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(208, 64);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 9;
            this.btnDelete.Text = "Delete All";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(12, 104);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 10;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 261);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnAsync);
            this.Controls.Add(this.lblCount);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.txtNum);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnMerge);
            this.Controls.Add(this.btnInsert);
            this.Name = "Form1";
            this.Text = "Bulk Copy";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNum;
        private System.Windows.Forms.Button btnCount;
        private System.Windows.Forms.Label lblCount;
        private System.Windows.Forms.Button btnAsync;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnExport;
    }
}

