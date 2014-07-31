using System;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.IO;

namespace CrystalScheduler
{
    public partial class ReportViewer : SchedulerBase
    {
        private string _reportFile;

        public ReportViewer(string reportFile)
        {
            InitializeComponent();

            _reportFile = reportFile;

            this.Text = base.Title + " - Report Viewer";
        }

        private void ReportViewer_Load(object sender, EventArgs e)
        {
            FileInfo fi = new FileInfo(_reportFile);
         
            ReportDocument reportDocument = new ReportDocument();
            reportDocument.Load(fi.FullName);
            this.crystalReportViewer.ReportSource = reportDocument;
        }
    }
}
