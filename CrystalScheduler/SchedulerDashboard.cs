using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace CrystalScheduler
{
    public partial class SchedulerDashboard : SchedulerBase
    {
        private List<ScheduledReport> _scheduledReportList;
        private DataTable _dataTable;

        public SchedulerDashboard()
        {
            InitializeComponent();
            this.Text = base.Title + " - Dashboard";
            InitializeReportGrid();
            LoadReportGrid();
        }

        private void InitializeReportGrid()
        {
            dgvScheduledReports.AllowUserToAddRows = false;
            dgvScheduledReports.AllowUserToDeleteRows = false;
            dgvScheduledReports.AllowDrop = false;

            _dataTable = new DataTable();
            _dataTable.Columns.Add(CrystalScheduler._ID).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._REPORT).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._NAME).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._DESCRIPTION).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._FILE_PATH).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._FILE_NAME).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._PARAMETERS).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._SCHEDULE).ReadOnly = true;
            _dataTable.Columns.Add(CrystalScheduler._OUTPUT).ReadOnly = true;
        }

        private void LoadReportGrid()
        {
            _scheduledReportList = DataAccess.GetScheduledReports();

            _dataTable.Clear();

            if (_scheduledReportList.Count > 0)
            {
                DataRow dataRow;
                foreach (ScheduledReport scheduledReport in _scheduledReportList)
                {
                    dataRow = _dataTable.NewRow();
                    dataRow[CrystalScheduler._ID] = scheduledReport.ScheduledReportID;
                    dataRow[CrystalScheduler._REPORT] = scheduledReport.Report;
                    dataRow[CrystalScheduler._FILE_NAME] = scheduledReport.FileName;
                    dataRow[CrystalScheduler._FILE_PATH] = scheduledReport.FilePath;
                    dataRow[CrystalScheduler._NAME] = scheduledReport.Name;
                    dataRow[CrystalScheduler._DESCRIPTION] = scheduledReport.Description;
                    dataRow[CrystalScheduler._PARAMETERS] = SerializeParameter(scheduledReport.Parameters);
                    dataRow[CrystalScheduler._SCHEDULE] = SerializeSchedule(scheduledReport.Schedule);
                    dataRow[CrystalScheduler._OUTPUT] = SerializeOutput(scheduledReport.Output);
                    _dataTable.Rows.Add(dataRow);
                }

                dgvScheduledReports.DataSource = _dataTable;
                dgvScheduledReports.Columns[CrystalScheduler._ID].Visible = false;
                dgvScheduledReports.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dgvScheduledReports.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dgvScheduledReports.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dgvScheduledReports.AutoResizeColumns();
                dgvScheduledReports.AutoResizeRows();
                EnableToolbar(true);
            }
            else
            {
                EnableToolbar(false);
            }
        }

        private void EnableToolbar(bool p)
        {
            tspView.Enabled = p;
            tspDuplicate.Enabled = p;
            tspDelete.Enabled = p;
        }
        
        private string SerializeParameter(List<ScheduledReportParameter> scheduledReportParameterList)
        {
            string result = string.Empty;

            foreach (ScheduledReportParameter p in scheduledReportParameterList)
            {

            }

            return result;
        }

        private string SerializeSchedule(ScheduledReportSchedule scheduledReportSchedule)
        {
            StringBuilder result = new StringBuilder();

            if (!scheduledReportSchedule.IsEnabled)
                result.AppendLine("DISABLED");

            switch (scheduledReportSchedule.Frequency)
            {
                case "OneTime":
                    //Occurs once on {mm/dd/yyyy} {hh:mm ap}
                    result.AppendLine(
                        string.Format("Occurs once on {0} {1};",
                        scheduledReportSchedule.StartDate.ToShortDateString(),
                        scheduledReportSchedule.StartTime.ToShortTimeString())
                        );
                    break;
                case "Daily":
                    //Recurs every {n} days; starts {mm/dd/yyyy} {hh:mm ap}; expires {mm/dd/yyyy} {hh:mm ap} 
                    //Recurs every {n} days; starts {mm/dd/yyyy} {hh:mm ap}; does not expire 
                    result.AppendLine(
                        string.Format(
                            "Recurs every {0} days;",
                            scheduledReportSchedule.Daily_RecursEveryNDays
                            )
                        );
                    result.AppendLine(
                        string.Format(
                            "starts {0} {1};",
                            scheduledReportSchedule.StartDate.ToShortDateString(),
                            scheduledReportSchedule.StartTime.ToShortTimeString()
                            )
                        );
                    if (scheduledReportSchedule.Expires)
                    {
                        result.AppendLine(
                            string.Format(
                                " expires {0} {1}",
                                scheduledReportSchedule.ExpireDate.ToShortDateString(),
                                scheduledReportSchedule.ExpireTime.ToShortTimeString()
                                )
                            );
                    }
                    else
                    {
                        result.Append(" does not expire");
                    }
                    break;
                case "Weekly":
                    //Recurs every {n} weeks on {days}; starts {mm/dd/yyyy} {hh:mm ap}; expires {mm/dd/yyyy} {hh:mm ap} 
                    //Recurs every {n} weeks on {days}; starts {mm/dd/yyyy} {hh:mm ap}; does not expire 
                    result.AppendLine(
                        string.Format(
                            "Recurs every {0} weeks on {1};",
                            scheduledReportSchedule.Weekly_RecursEveryNWeeks,
                            scheduledReportSchedule.Weekly_WhichDaysOfWeek
                            )
                        );
                    result.AppendLine(
                        string.Format(
                            "starts {0} {1};",
                            scheduledReportSchedule.StartDate.ToShortDateString(),
                            scheduledReportSchedule.StartTime.ToShortTimeString()
                            )
                        );
                    if (scheduledReportSchedule.Expires)
                    {
                        result.AppendLine(
                            string.Format(
                                " expires {0} {1}",
                                scheduledReportSchedule.ExpireDate.ToShortDateString(),
                                scheduledReportSchedule.ExpireTime.ToShortTimeString()
                                )
                            );
                    }
                    else
                    {
                        result.Append(" does not expire");
                    }
                    break;
                case "Monthly":
                    //Recurs every {months} on {days}; starts {mm/dd/yyyy} {hh:mm ap}; expires {mm/dd/yyyy} {hh:mm ap} 
                    //Recurs every {months} on {days}; starts {mm/dd/yyyy} {hh:mm ap}; does not expire 
                    //Recurs every {months} on the {day freq} {days}; starts {mm/dd/yyyy} {hh:mm ap}; expires {mm/dd/yyyy} {hh:mm ap}
                    //Recurs every {months} on the {day freq} days}; starts {mm/dd/yyyy} {hh:mm ap}; does not expire 
                    if (scheduledReportSchedule.Monthly_DaysOrFrequency == "Days")
                    {
                        result.AppendLine(
                            string.Format(
                                "Recurs every {0} on {1};",
                                scheduledReportSchedule.Monthly_WhichMonths,
                                scheduledReportSchedule.Monthly_Days_WhichDaysOfMonth
                                )
                            );
                    }
                    else
                    {
                        result.AppendLine(
                            string.Format(
                                "Recurs every {0} on the {1} {2};",
                                scheduledReportSchedule.Monthly_WhichMonths,
                                scheduledReportSchedule.Monthly_Frequency_WhichFrequency,
                                scheduledReportSchedule.Monthly_Frequency_WhichDaysOfWeek
                                )
                            );
                    }
                    result.AppendLine(
                        string.Format(
                            "starts {0} {1};",
                            scheduledReportSchedule.StartDate.ToShortDateString(),
                            scheduledReportSchedule.StartTime.ToShortTimeString()
                            )
                        );
                    if (scheduledReportSchedule.Expires)
                    {
                        result.AppendLine(
                            string.Format(
                                "expires {0} {1}",
                                scheduledReportSchedule.ExpireDate.ToShortDateString(),
                                scheduledReportSchedule.ExpireTime.ToShortTimeString()
                                )
                            );
                    }
                    else
                    {
                        result.AppendLine("does not expire");
                    }
                    break;
            }

            return result.ToString();
        }

        private string SerializeOutput(ScheduledReportOutput scheduledReportOutput)
        {
            StringBuilder result = new StringBuilder();

            if (scheduledReportOutput.SendToPrinter)
            {
                result.AppendLine(
                    string.Format(
                        "Print to {0};",
                        scheduledReportOutput.Printer
                        )
                    );
            }
            if (scheduledReportOutput.SendToFolder)
            {
                result.AppendLine(
                    string.Format(
                        "Save to folder {0};",
                        scheduledReportOutput.Folder
                        )
                    );
            }
            if (scheduledReportOutput.SendToEmail)
            {
                result.AppendLine(
                    string.Format(
                        "Email To: {0}, From: {1}, CC: {2}, Subject: {3}, Message: {4}",
                        scheduledReportOutput.EmailTo,
                        scheduledReportOutput.EmailFrom,
                        scheduledReportOutput.EmailCC,
                        scheduledReportOutput.EmailSubject,
                        scheduledReportOutput.EmailMessage
                        )
                    );
            }
            return result.ToString();
        }

        private void btnScheduleNewReport_Click(object sender, EventArgs e)
        {
            ReportScheduler reportScheduler = null;
            try
            {
                this.UseWaitCursor = true;

                reportScheduler = new ReportScheduler();
                reportScheduler.ShowDialog(this);
                if (reportScheduler.SaveClicked)
                {
                    DataAccess.SaveScheduledReport(reportScheduler.ScheduledReport);
                    LoadReportGrid();
                }
            }
            finally
            {
                if (reportScheduler != null)
                    reportScheduler.Dispose();
                this.UseWaitCursor = false;
            }
        }

        private void tspView_Click(object sender, EventArgs e)
        {
            if (dgvScheduledReports.CurrentRow != null)
            {
                int id = Convert.ToInt32(dgvScheduledReports.SelectedRows[0].Cells[CrystalScheduler._ID].Value);
                ScheduledReport scheduledReport = DataAccess.GetScheduledReportByID(id);
                ReportScheduler reportScheduler = null;
                try
                {
                    this.UseWaitCursor = true;

                    reportScheduler = new ReportScheduler(scheduledReport);
                    reportScheduler.ShowDialog(this);
                    if (reportScheduler.SaveClicked)
                    {
                        DataAccess.SaveScheduledReport(reportScheduler.ScheduledReport);
                        LoadReportGrid();
                    }
                }
                finally
                {
                    if (reportScheduler != null)
                        reportScheduler.Dispose();
                    this.UseWaitCursor = false;
                }
            }   
        }

        private void tspDuplicate_Click(object sender, EventArgs e)
        {
            if (dgvScheduledReports.CurrentRow != null)
            {
                try
                {
                    this.UseWaitCursor = true;
                    int id = Convert.ToInt32(dgvScheduledReports.SelectedRows[0].Cells[CrystalScheduler._ID].Value);
                    DataAccess.DuplicateScheduledReport(id);
                    LoadReportGrid();
                }
                finally
                {
                    this.UseWaitCursor = false;
                }
            }
        }
        
        private void tspDelete_Click(object sender, EventArgs e)
        {
            if (dgvScheduledReports.CurrentRow != null)
            {
                DialogResult result = MessageBox.Show(this, 
                    "Are you sure you want to delete the selected Scheduled Report?", 
                    "Are you sure?", 
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        this.UseWaitCursor = true;
                        int id = Convert.ToInt32(dgvScheduledReports.SelectedRows[0].Cells[CrystalScheduler._ID].Value);
                        DataAccess.DeleteScheduledReport(id);
                        LoadReportGrid();
                    }
                    finally
                    {
                        this.UseWaitCursor = false;
                    }
                }
            }
        }
    }
}
