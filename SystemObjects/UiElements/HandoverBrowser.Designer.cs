using System;

namespace MyntraExcelAddin.SystemObjects.UiElements
{
    partial class HandoverBrowser
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.next = new System.Windows.Forms.Button();
            this.previous = new System.Windows.Forms.Button();
            this.download = new System.Windows.Forms.Button();
            this.comboBox_pagesize = new System.Windows.Forms.ComboBox();
            this.pagesize = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.ControlText;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(776, 426);
            this.dataGridView1.TabIndex = 0;
            // 
            // next
            // 
            this.next.Location = new System.Drawing.Point(696, 444);
            this.next.Name = "next";
            this.next.Size = new System.Drawing.Size(92, 43);
            this.next.TabIndex = 1;
            this.next.Text = "Next";
            this.next.UseVisualStyleBackColor = true;
            this.next.Click += new System.EventHandler(this.next_Click);
            // 
            // previous
            // 
            this.previous.Location = new System.Drawing.Point(598, 444);
            this.previous.Name = "previous";
            this.previous.Size = new System.Drawing.Size(92, 43);
            this.previous.TabIndex = 2;
            this.previous.Text = "Previous";
            this.previous.UseVisualStyleBackColor = true;
            this.previous.Click += new System.EventHandler(this.previous_Click);
            // 
            // download
            // 
            this.download.Location = new System.Drawing.Point(598, 502);
            this.download.Name = "download";
            this.download.Size = new System.Drawing.Size(190, 49);
            this.download.TabIndex = 3;
            this.download.Text = "Download";
            this.download.UseVisualStyleBackColor = true;
            this.download.Click += new System.EventHandler(this.download_Click);
            // 
            // comboBox_pagesize
            // 
            this.comboBox_pagesize.FormattingEnabled = true;
            this.comboBox_pagesize.Items.AddRange(new object[] {
            "10",
            "20",
            "50",
            "100"});
            this.comboBox_pagesize.Location = new System.Drawing.Point(518, 444);
            this.comboBox_pagesize.Name = "comboBox_pagesize";
            this.comboBox_pagesize.Size = new System.Drawing.Size(74, 28);
            this.comboBox_pagesize.TabIndex = 4;
            this.comboBox_pagesize.Text = "10";
            this.comboBox_pagesize.SelectedIndexChanged += new System.EventHandler(this.comboBox_pagesize_SelectedIndexChanged);            
            // 
            // pagesize
            // 
            this.pagesize.AutoSize = true;
            this.pagesize.Location = new System.Drawing.Point(431, 447);
            this.pagesize.Name = "pagesize";
            this.pagesize.Size = new System.Drawing.Size(81, 20);
            this.pagesize.TabIndex = 5;
            this.pagesize.Text = "Page Size";
            // 
            // HandoverBrowser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 564);
            this.Controls.Add(this.pagesize);
            this.Controls.Add(this.comboBox_pagesize);
            this.Controls.Add(this.download);
            this.Controls.Add(this.previous);
            this.Controls.Add(this.next);
            this.Controls.Add(this.dataGridView1);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "HandoverBrowser";
            this.Text = "Handover Browser";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button next;
        private System.Windows.Forms.Button previous;
        private System.Windows.Forms.Button download;
        private System.Windows.Forms.ComboBox comboBox_pagesize;
        private System.Windows.Forms.Label pagesize;
    }
}