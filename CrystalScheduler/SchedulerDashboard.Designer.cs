namespace CrystalScheduler
{
    partial class SchedulerDashboard
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SchedulerDashboard));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dgvScheduledReports = new System.Windows.Forms.DataGridView();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnScheduleReport = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.tspView = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tspDuplicate = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.tspDelete = new System.Windows.Forms.ToolStripButton();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvScheduledReports)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dgvScheduledReports, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.toolStrip1, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 35F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(784, 562);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // dgvScheduledReports
            // 
            this.dgvScheduledReports.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvScheduledReports.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvScheduledReports.Location = new System.Drawing.Point(3, 38);
            this.dgvScheduledReports.Name = "dgvScheduledReports";
            this.dgvScheduledReports.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvScheduledReports.Size = new System.Drawing.Size(778, 521);
            this.dgvScheduledReports.TabIndex = 0;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnScheduleReport,
            this.toolStripSeparator4,
            this.tspView,
            this.toolStripSeparator1,
            this.tspDuplicate,
            this.toolStripSeparator2,
            this.tspDelete});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(784, 35);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // btnScheduleReport
            // 
            this.btnScheduleReport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnScheduleReport.Image = ((System.Drawing.Image)(resources.GetObject("btnScheduleReport.Image")));
            this.btnScheduleReport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnScheduleReport.Name = "btnScheduleReport";
            this.btnScheduleReport.Size = new System.Drawing.Size(124, 32);
            this.btnScheduleReport.Text = "Schedule New Report";
            this.btnScheduleReport.Click += new System.EventHandler(this.btnScheduleNewReport_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 35);
            // 
            // tspView
            // 
            this.tspView.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tspView.Image = ((System.Drawing.Image)(resources.GetObject("tspView.Image")));
            this.tspView.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tspView.Name = "tspView";
            this.tspView.Size = new System.Drawing.Size(36, 32);
            this.tspView.Text = "View";
            this.tspView.Click += new System.EventHandler(this.tspView_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 35);
            // 
            // tspDuplicate
            // 
            this.tspDuplicate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tspDuplicate.Image = ((System.Drawing.Image)(resources.GetObject("tspDuplicate.Image")));
            this.tspDuplicate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tspDuplicate.Name = "tspDuplicate";
            this.tspDuplicate.Size = new System.Drawing.Size(61, 32);
            this.tspDuplicate.Text = "Duplicate";
            this.tspDuplicate.Click += new System.EventHandler(this.tspDuplicate_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 35);
            // 
            // tspDelete
            // 
            this.tspDelete.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tspDelete.Image = ((System.Drawing.Image)(resources.GetObject("tspDelete.Image")));
            this.tspDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tspDelete.Name = "tspDelete";
            this.tspDelete.Size = new System.Drawing.Size(44, 32);
            this.tspDelete.Text = "Delete";
            this.tspDelete.Click += new System.EventHandler(this.tspDelete_Click);
            // 
            // SchedulerDashboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.tableLayoutPanel1);
            this.MinimumSize = new System.Drawing.Size(800, 600);
            this.Name = "SchedulerDashboard";
            this.Text = "";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvScheduledReports)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dgvScheduledReports;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton btnScheduleReport;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton tspDuplicate;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton tspDelete;
        private System.Windows.Forms.ToolStripButton tspView;

    }
}

