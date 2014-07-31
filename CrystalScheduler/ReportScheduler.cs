using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CrystalScheduler
{
    public partial class ReportScheduler : SchedulerBase
    {
        private bool _saveNeeded = false;
        private string _reportFile = string.Empty;
        private ReportDocument _reportDocument = new ReportDocument();
        private ScheduledReport _scheduledReport;
        private bool _saveClicked = false;

        public bool SaveClicked { get { return _saveClicked; } }
        public ScheduledReport ScheduledReport { get { return _scheduledReport; } }

        public ReportScheduler()
        {
            InitializeComponent();
            _scheduledReport = new ScheduledReport();
            Initialize();           
        }

        public ReportScheduler(ScheduledReport scheduledReport)
        {
            InitializeComponent();
            _scheduledReport = scheduledReport;
            Initialize();
        }

        private void Initialize()
        {
            this.Text = base.Title + " - Schedule Report";
            this.Width = 600;
            this.Height = 365;
            PopulateTabs(); 
        }

        private void PopulateTabs()
        {
            EnablePreview(false);
            EnableSave(false);

            PopulateReportTab();
            PopulateScheduleTab();
            PopulateOutputTab();
        }

        private void EnablePreview(bool p)
        {
            btnPreview.Enabled = p;
        }

        private void EnableSave(bool p)
        {
            btnSave.Enabled = p;
        }

        private void EnableTabs(bool p)
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Text != "Report")
                {
                    tabPage.Enabled = p;
                }
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            ReportViewer reportViewer = null;
            try
            {
                this.UseWaitCursor = true;
                reportViewer = new ReportViewer(_reportFile);
                reportViewer.ShowDialog(this);
                reportViewer.Dispose();
            }
            finally
            {
                if (reportViewer != null)
                    reportViewer.Dispose();
                this.UseWaitCursor = false;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveScreenDataToObject();
            _saveClicked = true;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _saveClicked = false;
            Close();
        }

        private void SaveScreenDataToObject()
        {
            _scheduledReport.Report = txtReport.Text;
            _scheduledReport.FileName = txtFileName.Text;
            _scheduledReport.FilePath = txtFilePath.Text;
            _scheduledReport.Name = txtName.Text;
            _scheduledReport.Description = txtDescription.Text;
            
            _scheduledReport.Schedule.IsEnabled = chkEnabled.Checked;
            _scheduledReport.Schedule.Frequency = GetFrequency();
            _scheduledReport.Schedule.StartDate = Convert.ToDateTime(dtpStartDate.Value);
            _scheduledReport.Schedule.StartTime = Convert.ToDateTime(dtpStartTime.Value);
            _scheduledReport.Schedule.Expires = chkExpire.Checked;
            _scheduledReport.Schedule.ExpireDate = Convert.ToDateTime(dtpExpireDate.Value);
            _scheduledReport.Schedule.ExpireTime = Convert.ToDateTime(dtpExpireTime.Value);
            _scheduledReport.Schedule.Daily_RecursEveryNDays = Convert.ToInt32(txtDailyNDays.Text);
            _scheduledReport.Schedule.Weekly_RecursEveryNWeeks = Convert.ToInt32(txtWeeklyNWeeks.Text);
            _scheduledReport.Schedule.Weekly_WhichDaysOfWeek = GetWeeklyWhichDaysOfWeek();
            _scheduledReport.Schedule.Monthly_WhichMonths = cboMonthlyMonths.Text;
            _scheduledReport.Schedule.Monthly_DaysOrFrequency = GetMonthlyDaysOrFrequency();
            _scheduledReport.Schedule.Monthly_Days_WhichDaysOfMonth = cboMonthlyDaysOfMonth.Text;
            _scheduledReport.Schedule.Monthly_Frequency_WhichFrequency = cboMonthlyFreq.Text;
            _scheduledReport.Schedule.Monthly_Frequency_WhichDaysOfWeek = cboMonthlyDaysOfWeek.Text;

            _scheduledReport.Output.SendToPrinter = chkOutputPrinter.Checked;
            _scheduledReport.Output.Printer = cboOutputPrinters.Text;
            _scheduledReport.Output.SendToFolder = chkOutputFolder.Checked;
            _scheduledReport.Output.Folder = txtOutputFolder.Text;
            _scheduledReport.Output.SendToEmail = chkOutputEmail.Checked;
            _scheduledReport.Output.EmailTo = txtOutputEmailTo.Text;
            _scheduledReport.Output.EmailFrom = txtOutputEmailFrom.Text;
            _scheduledReport.Output.EmailCC = txtOutputEmailCC.Text;
            _scheduledReport.Output.EmailSubject = txtOutputEmailSubject.Text;
            _scheduledReport.Output.EmailMessage = txtOutputEmailMessage.Text;
        }

        private string GetFrequency()
        {
            string result = string.Empty;

            if (optOneTime.Checked)
                result = "OneTime";
            else if (optDaily.Checked)
                result = "Daily";
            else if (optWeekly.Checked)
                result = "Weekly";
            else if (optMonthly.Checked)
                result = "Monthly";

            return result;
        }
        
        private string GetWeeklyWhichDaysOfWeek()
        {
            StringBuilder result = new StringBuilder();

            if (chkWeeklySunday.Checked)
                result.Append("Sunday,");
            if (chkWeeklyMonday.Checked)
                result.Append("Monday,");
            if (chkWeeklyTuesday.Checked)
                result.Append("Tuesday,");
            if (chkWeeklyWednesday.Checked)
                result.Append("Wednesday,");
            if (chkWeeklyThursday.Checked)
                result.Append("Thursday,");
            if (chkWeeklyFriday.Checked)
                result.Append("Friday,");
            if (chkWeeklySaturday.Checked)
                result.Append("Saturday,");

            if (result.Length > 0)
                result.Length--;

            return result.ToString();
        }

        private string GetMonthlyDaysOrFrequency()
        {
            string result = string.Empty;

            if (optMonthlyDays.Checked)
                result = "Days";
            else if (optMonthlyFrequency.Checked)
                result = "Frequency";

            return result;
        }

        #region Report Tab
        private void PopulateReportTab()
        {
            if (_scheduledReport.Report != string.Empty)
            {
                txtReport.Text = _scheduledReport.Report;
                txtFileName.Text = _scheduledReport.FileName;
                txtFilePath.Text = _scheduledReport.FilePath;
                txtName.Text = _scheduledReport.Name;
                txtDescription.Text = _scheduledReport.Description;
            }
            else
                EnableTabs(false);
        }

        private void btnSelectReport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = null;
            try
            {
                this.UseWaitCursor = true;

                EnablePreview(false);
                EnableSave(false);

                ofd = new OpenFileDialog();
                ofd.Filter = "Crystal Report files (*.rpt)|*.rpt";
                ofd.Title = "Select a Crystal Report file";
                DialogResult dr = ofd.ShowDialog(this);
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    _reportFile = ofd.FileName;
                    RefreshReportFileInfo();
                }
            }
            finally
            {
                if (ofd != null)
                    ofd.Dispose();
                this.UseWaitCursor = false;
            }
        }

        private void txtReport_TextChanged(object sender, EventArgs e)
        {
            bool reportSelected = (txtReport.Text != string.Empty);

            EnablePreview(reportSelected);
            EnableTabs(reportSelected);
            EnableSave(reportSelected);
        }

        private void RefreshReportFileInfo()
        {
            Cursor.Current = Cursors.WaitCursor;

            FileInfo fi = new FileInfo(_reportFile);
            
            _reportDocument.Load(fi.FullName);

            txtReport.Text = string.IsNullOrEmpty(_reportDocument.Name) ? fi.Name.Replace(".rpt", "") : _reportDocument.Name;
            txtFileName.Text = fi.Name;
            txtFilePath.Text = fi.DirectoryName;

            _parameterFields = _reportDocument.ParameterFields;
            PopulateParameterTab();

            Cursor.Current = Cursors.Default;
        }
        #endregion

        #region Parameter Tab
        private bool _loadingParameters = false;
        private ParameterFields _parameterFields;

        private void PopulateParameterTab()
        {
            LoadParameters();
            DisableParameterControls();
        }

        private void LoadParameters()
        {
            _loadingParameters = true;

            string value = string.Empty;

            foreach (ParameterField parameterField in _parameterFields)
            {
                if (parameterField.ParameterValueType == ParameterValueKind.DateParameter)
                {
                    if (parameterField.PromptingType == DiscreteOrRangeKind.RangeValue)
                    {
                        DateTime today = DateTime.Today;
                        value = today.ToShortDateString() + " - " + today.ToShortDateString();
                        dtpParameterFromDate.Value = today;
                        dtpParameterToDate.Value = today;
                    }
                    else if (parameterField.PromptingType == DiscreteOrRangeKind.DiscreteValue)
                    {
                        DateTime today = DateTime.Today;
                        value = today.ToShortDateString();
                        dtpParameterFromDate.Value = today;
                    }
                }
                else if (parameterField.ParameterValueType == ParameterValueKind.DateTimeParameter)
                {
                    if (parameterField.PromptingType == DiscreteOrRangeKind.RangeValue)
                    {
                        DateTime today = DateTime.Today;
                        value = string.Format("{0} {1} - {2} {3}", today.ToShortDateString(), today.ToShortTimeString(), today.ToShortDateString(), today.ToShortTimeString());
                        dtpParameterFromDate.Value = today;
                        dtpParameterFromTime.Value = today;
                        dtpParameterToDate.Value = today;
                        dtpParameterToTime.Value = today;
                    }
                    else if (parameterField.PromptingType == DiscreteOrRangeKind.DiscreteValue)
                    {
                        DateTime today = DateTime.Today;
                        value = string.Format("{0} {1}", today.ToShortDateString(), today.ToShortTimeString());
                        dtpParameterFromDate.Value = today;
                        dtpParameterFromTime.Value = today;
                    }
                }
                else
                {
                    if (parameterField.DefaultValues.Count > 0)
                    {
                        CrystalDecisions.Shared.ParameterValue parameterValue = parameterField.DefaultValues[0];
                        value = parameterValue.ToString();
                    }
                    else
                    {
                        value = string.Empty;
                    }
                }

                ListViewItem listViewItem = new ListViewItem();
                listViewItem.Text = parameterField.Name;
                listViewItem.SubItems.Add(value);
                lvwParameters.Items.Add(listViewItem);
            }

            _loadingParameters = false;
            AdjustParameterColumnWidth("Parameters");
        }

        private void DisableParameterControls()
        {
            dtpParameterFromDate.Visible = false;
            dtpParameterFromTime.Visible = false;
            dtpParameterToDate.Visible = false;
            dtpParameterToTime.Visible = false;
            cboParameterValue.Visible = false;
            txtParameterValue.Visible = false;
        }

        private void btnParameterApplyChange_Click(object sender, EventArgs e)
        {
            SaveParameterChanges();
            lvwParameters.Focus();
        }

        private void SaveParameterChanges()
        {
            int selectedIndex = -1;

            for (int i = 0; i < lvwParameters.Items.Count; i++)
            {
                if (lvwParameters.Items[i].SubItems[0].Text == txtParameter.Text)
                    selectedIndex = i;
            }

            if (cboParameterValue.Visible && !lvwParameterValues.Visible)
            {
                lvwParameters.Items[selectedIndex].SubItems[1].Text = cboParameterValue.Text;
            }
            else
            {
                if (dtpParameterFromDate.Visible)
                {
                    if (dtpParameterToDate.Visible)
                    {
                        if (dtpParameterToTime.Visible)
                        {
                            lvwParameters.Items[selectedIndex].SubItems[1].Text =
                                (dtpParameterFromDate.Checked ? dtpParameterFromDate.Value.ToShortDateString() + " " + dtpParameterFromTime.Value.ToShortDateString() : "up to") + " - " +
                                (dtpParameterToDate.Checked ? dtpParameterToDate.Value.ToShortDateString() + " " + dtpParameterToTime.Value.ToShortDateString() : "and up");
                        }
                        else
                        {
                            lvwParameters.Items[selectedIndex].SubItems[1].Text =
                                (dtpParameterFromDate.Checked ? dtpParameterFromDate.Value.ToShortDateString() : "up to") + " - " +
                                (dtpParameterToDate.Checked ? dtpParameterToDate.Value.ToShortDateString() : "and up");
                        }
                    }
                    else
                    {
                        if (dtpParameterFromTime.Visible)
                        {
                            lvwParameters.Items[selectedIndex].SubItems[1].Text = dtpParameterFromDate.Checked ? dtpParameterFromDate.Value.ToShortDateString() + " " + dtpParameterFromTime.Value.ToShortDateString() : "up to";
                        }
                        else
                        {
                            lvwParameters.Items[selectedIndex].SubItems[1].Text = dtpParameterFromDate.Checked ? dtpParameterFromDate.Value.ToShortDateString() : "up to";
                        }
                    }
                }
                else
                {
                    string value = string.Empty;

                    if (lvwParameterValues.Items.Count > 0)
                    {
                        for (int i = 0; i < lvwParameterValues.Items.Count; i++)
                        {
                            if (i > 0)
                            {
                                value = value + ";";
                            }
                            value = value + lvwParameterValues.Items[i].Text.Trim();
                        }
                    }
                    else
                    {
                        value = "All";
                    }

                    lvwParameters.Items[selectedIndex].SubItems[1].Text = value;
                }
            }

            _saveNeeded = false;
        }

        private void btnParameterRemove_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lvwParameterValues.SelectedItems.Count - 1; i++)
            {
                lvwParameterValues.SelectedItems[0].Remove();
            }

            if (lvwParameterValues.Items.Count == 0)
            {
                ListViewItem listViewItem = new ListViewItem();
                listViewItem.Text = "All";
                lvwParameterValues.Items.Add(listViewItem);
            }

            txtParameterValue.Focus();
        }

        private void LoadParameterCombo(string p)
        {
            _loadingParameters = true;

            cboParameterValue.Enabled = true;
            cboParameterValue.Items.Clear();
            cboParameterValue.DisplayMember = "text";
            cboParameterValue.ValueMember = "tag";

            for (int i = 0; i < _parameterFields[p].DefaultValues.Count; i++)
            {
                ParameterDiscreteValue parameterDiscreteValue = new ParameterDiscreteValue();
                parameterDiscreteValue = (ParameterDiscreteValue)_parameterFields[p].DefaultValues[i];
                ListViewItem listViewItem = new ListViewItem();
                listViewItem.Text = parameterDiscreteValue.Value.ToString();
                cboParameterValue.Items.Add(listViewItem);
            }

            cboParameterValue.Text = string.Empty;

            _loadingParameters = false;
        }

        private void cboParameterValue_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_loadingParameters)
                return;

            _saveNeeded = true;

            if (cboParameterValue.SelectedIndex > -1)
            {
                //btnOK.Enabled = true;
                if (lvwParameterValues.Visible)
                    AddParameterValueListItem();
            }
        }

        private void dtpParameterFromDate_ValueChanged(object sender, EventArgs e)
        {
            _saveNeeded = true;
            if (!dtpParameterFromDate.Checked)
                if (!dtpParameterToDate.Checked)
                    dtpParameterToDate.Checked = true;
        }

        private void dtpParameterToDate_ValueChanged(object sender, EventArgs e)
        {
            _saveNeeded = true;
            if (!dtpParameterToDate.Checked)
                if (!dtpParameterFromDate.Checked)
                    dtpParameterFromDate.Checked = true;
        }

        private void lvwParameters_DoubleClick(object sender, EventArgs e)
        {
            if (cboParameterValue.Enabled)
                cboParameterValue.Focus();
            else
                dtpParameterFromDate.Focus();
        }

        private void lvwParameters_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvwParameters.SelectedItems.Count > 0)
            {
                btnParameterApplyChange.Enabled = true;
                btnParameterApplyChange.Visible = true;
                txtParameter.Text = lvwParameters.SelectedItems[0].Text;
                lvwParameterValues.Items.Clear();
            }

            lvwParameterValues.Visible = _parameterFields[lvwParameters.SelectedItems[0].Text].EnableAllowMultipleValue;
            txtDescription.Text = _parameterFields[lvwParameters.SelectedItems[0].Text].PromptText;

            if (_parameterFields[lvwParameters.SelectedItems[0].Text].ParameterValueType == ParameterValueKind.DateParameter)
            {
                DisableParameterControls();

                if (_parameterFields[lvwParameters.SelectedItems[0].Text].DiscreteOrRangeKind == DiscreteOrRangeKind.RangeValue)
                {
                    dtpParameterToDate.Visible = true;
                }

                dtpParameterFromDate.Visible = true;
                dtpParameterFromDate.Focus();
            }
            else if (_parameterFields[lvwParameters.SelectedItems[0].Text].ParameterValueType == ParameterValueKind.DateTimeParameter)
            {
                DisableParameterControls();

                if (_parameterFields[lvwParameters.SelectedItems[0].Text].DiscreteOrRangeKind == DiscreteOrRangeKind.RangeValue)
                {
                    dtpParameterToDate.Visible = true;
                    dtpParameterToTime.Visible = true;
                }

                dtpParameterFromDate.Visible = true;
                dtpParameterFromTime.Visible = true;
                dtpParameterFromDate.Focus();
            }
            else if (_parameterFields[lvwParameters.SelectedItems[0].Text].EnableAllowMultipleValue)
            {
                LoadParameterValueListView();
                DisableParameterControls();

                if (_parameterFields[lvwParameters.SelectedItems[0].Text].DefaultValues.Count > 1)
                {
                    LoadParameterCombo(lvwParameters.SelectedItems[0].Text);
                    cboParameterValue.Visible = true;
                    cboParameterValue.Focus();
                }
                else
                {
                    txtParameterValue.Visible = true;
                    txtParameterValue.Focus();
                }
            }
            else
            {
                DisableParameterControls();
                LoadParameterCombo(lvwParameters.SelectedItems[0].Text);
                cboParameterValue.Visible = true;
                cboParameterValue.Focus();
            }
        }

        private void lvwValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvwParameterValues.SelectedItems.Count > 0)
                btnParameterRemove.Enabled = (lvwParameterValues.SelectedItems[0].Text == "All" ? false : true);
        }

        private void txtParameterValue_Leave(object sender, EventArgs e)
        {
            if (lvwParameterValues.Visible && txtParameterValue.TextLength > 0)
                AddParameterValueListItem();
        }

        private void AddParameterValueListItem()
        {
            bool proceed = true;

            for (int i = 0; i < lvwParameterValues.Items.Count; i++)
            {
                if (lvwParameterValues.Items[i].Text.ToUpper() == "ALL")
                {
                    lvwParameterValues.Items[i].Remove();
                }
                else if (lvwParameterValues.Items[i].Text.ToUpper() == (txtParameterValue.Visible ? txtParameterValue.Text.ToUpper() : cboParameterValue.Text.ToUpper()))
                {
                    proceed = false;
                    break;
                }
            }

            if (proceed)
            {
                if (txtParameterValue.Visible && txtParameterValue.Text.ToUpper() == "All")
                    lvwParameterValues.Items.Clear();

                for (int i = 0; i < lvwParameterValues.Items.Count; i++)
                {
                    if (lvwParameterValues.Items[i].Text.ToUpper() == (txtParameterValue.Visible ? txtParameterValue.Text.ToUpper() : cboParameterValue.Text.ToUpper()))
                    {
                        proceed = false;
                        break;
                    }
                }
            }

            if (proceed)
            {
                ListViewItem listViewItem = new ListViewItem();
                listViewItem.Text = (txtParameterValue.Visible ? txtParameterValue.Text : cboParameterValue.Text);
                if (listViewItem.Text == string.Empty)
                    listViewItem.Text = cboParameterValue.Text;

                lvwParameterValues.Items.Add(listViewItem);

                txtParameterValue.Text = string.Empty;
                cboParameterValue.Text = string.Empty;
                cboParameterValue.SelectedIndex = -1;
            }

            btnParameterApplyChange.Enabled = true;
            AdjustParameterColumnWidth("Values");
            txtParameterValue.Focus();
        }

        private void AdjustParameterColumnWidth(string whichColumn)
        {
            switch (whichColumn)
            {
                case "Values":
                    lvwParameters.Columns[0].Width = (lvwParameters.Items.Count > 5 ? 185 : 200);
                    break;
                case "Parameters":
                    lvwParameters.Columns[1].Width = (lvwParameters.Items.Count > 12 ? 175 : 190);
                    break;
                default:
                    break;
            }
        }

        private void txtValue_TextChanged(object sender, EventArgs e)
        {
            _saveNeeded = true;
            if (lvwParameterValues.Visible)
                btnParameterRemove.Enabled = lvwParameterValues.SelectedItems.Count > 0 ? true : false;
        }

        private void LoadParameterValueListView()
        {
            btnParameterRemove.Enabled = false;
            btnParameterRemove.Visible = true;

            int position = 0;
            int startPosition = 0;
            lvwParameterValues.Items.Clear();

            while (true)
            {
                position = lvwParameters.SelectedItems[0].SubItems[1].Text.IndexOf(";", startPosition);
                if (position == 0)
                {
                    ListViewItem listViewItem = new ListViewItem();
                    listViewItem.Text = lvwParameters.SelectedItems[0].SubItems[1].Text.Substring(startPosition, lvwParameters.SelectedItems[0].SubItems[1].Text.Length - startPosition);
                    lvwParameterValues.Items.Add(listViewItem);
                    break;
                }
                else
                {
                    ListViewItem listViewItem = new ListViewItem();
                    listViewItem.Text = lvwParameters.SelectedItems[0].SubItems[1].Text.Substring(startPosition, position - startPosition);
                    lvwParameterValues.Items.Add(listViewItem);
                    startPosition = position + 1;
                }
            }
            AdjustParameterColumnWidth("Values");
            txtParameterValue.Focus();
        }
        #endregion

        #region Schedule Tab
        private enum WhichFreq { OneTime, Daily, Weekly, Monthly }
        private enum MonthlyOptions { Days, On }

        private void PopulateScheduleTab()
        {
            dtpStartTime.ShowUpDown = true;
            dtpExpireTime.ShowUpDown = true;
            chkExpire.Checked = false;
            dtpExpireDate.Enabled = false;
            dtpExpireTime.Enabled = false;

            ShowWhichFreq(WhichFreq.OneTime);
            ShowMonthlyOption(MonthlyOptions.Days);

            string[] months = new string[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            foreach (string month in months)
            {
                cboMonthlyMonths.Items.Add(month);
            }

            string[] days = new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "Last" };
            foreach (string day in days)
            {
                cboMonthlyDaysOfMonth.Items.Add(day);
            }

            string[] frequencies = new string[] { "First", "Second", "Third", "Fourth", "Last" };
            foreach (string frequency in frequencies)
            {
                cboMonthlyFreq.Items.Add(frequency);
            }

            string[] daysOfWeek = new string[] { "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday" };
            foreach (string dayOfWeek in daysOfWeek)
            {
                cboMonthlyDaysOfWeek.Items.Add(dayOfWeek);
            }
        }

        private void optOneTime_CheckedChanged(object sender, EventArgs e)
        {
            if (optOneTime.Checked)
            {
                ShowWhichFreq(WhichFreq.OneTime);
            }
        }

        private void optDaily_CheckedChanged(object sender, EventArgs e)
        {
            if (optDaily.Checked)
            {
                ShowWhichFreq(WhichFreq.Daily);
            }
        }

        private void optWeekly_CheckedChanged(object sender, EventArgs e)
        {
            if (optWeekly.Checked)
            {
                ShowWhichFreq(WhichFreq.Weekly);
            }
        }

        private void optMonthly_CheckedChanged(object sender, EventArgs e)
        {
            if (optMonthly.Checked)
            {
                ShowWhichFreq(WhichFreq.Monthly);
            }
        }

        private void ShowWhichFreq(WhichFreq which)
        {
            switch (which)
            {
                case WhichFreq.OneTime:
                    chkExpire.Visible = false;
                    dtpExpireDate.Visible = false;
                    dtpExpireTime.Visible = false;
                    grpDaily.Visible = false;
                    grpWeekly.Visible = false;
                    grpMonthly.Visible = false;
                    break;
                case WhichFreq.Daily:
                    chkExpire.Visible = true;
                    dtpExpireDate.Visible = true;
                    dtpExpireTime.Visible = true;
                    grpDaily.Visible = true;
                    grpWeekly.Visible = false;
                    grpMonthly.Visible = false;
                    grpDaily.Location = new Point(121, 74);
                    break;
                case WhichFreq.Weekly:
                    chkExpire.Visible = true;
                    dtpExpireDate.Visible = true;
                    dtpExpireTime.Visible = true;
                    grpDaily.Visible = false;
                    grpWeekly.Visible = true;
                    grpMonthly.Visible = false;
                    grpWeekly.Location = new Point(121, 74);
                    break;
                case WhichFreq.Monthly:
                    chkExpire.Visible = true;
                    dtpExpireDate.Visible = true;
                    dtpExpireTime.Visible = true;
                    grpDaily.Visible = false;
                    grpWeekly.Visible = false;
                    grpMonthly.Visible = true;
                    grpMonthly.Location = new Point(121, 74);
                    break;
            }
        }

        private void chkExpire_CheckedChanged(object sender, EventArgs e)
        {
            dtpExpireDate.Enabled = chkExpire.Checked;
            dtpExpireTime.Enabled = chkExpire.Checked;
        }

        private void optMonthlyDays_CheckedChanged(object sender, EventArgs e)
        {
            if (optMonthlyDays.Checked)
            {
                ShowMonthlyOption(MonthlyOptions.Days);
            }
        }

        private void optMonthlyOn_CheckedChanged(object sender, EventArgs e)
        {
            if (optMonthlyFrequency.Checked)
            {
                ShowMonthlyOption(MonthlyOptions.On);
            }
        }

        private void ShowMonthlyOption(MonthlyOptions monthlyOptions)
        {
            switch (monthlyOptions)
            {
                case (MonthlyOptions.Days):
                    cboMonthlyDaysOfMonth.Enabled = true;
                    cboMonthlyFreq.Enabled = false;
                    cboMonthlyDaysOfWeek.Enabled = false;
                    break;
                case (MonthlyOptions.On):
                    cboMonthlyDaysOfMonth.Enabled = false;
                    cboMonthlyFreq.Enabled = true;
                    cboMonthlyDaysOfWeek.Enabled = true;
                    break;
            }
        }
        #endregion

        #region Output Tab
        private void PopulateOutputTab()
        {
            EnableOutputPrinters(false);
            EnableOutputFolder(false);
            EnableOutputEmail(false);
            PopulatePrinters();
            PopulateEmailDefaults();
        }

        private void EnableOutputPrinters(bool p)
        {
            chkOutputPrinter.Checked = p;
            cboOutputPrinters.Enabled = p;
        }

        private void EnableOutputFolder(bool p)
        {
            chkOutputFolder.Checked = p;
            txtOutputFolder.Enabled = p;
            btnSelectFolder.Enabled = p;
        }

        private void EnableOutputEmail(bool p)
        {
            chkOutputEmail.Checked = p;
            txtOutputEmailTo.Enabled = p;
            txtOutputEmailFrom.Enabled = p;
            txtOutputEmailCC.Enabled = p;
            txtOutputEmailSubject.Enabled = p;
            txtOutputEmailMessage.Enabled = p;
        }

        private void PopulatePrinters()
        {
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                cboOutputPrinters.Items.Add(printer);
            }
        }

        private void PopulateEmailDefaults()
        {
            txtOutputEmailTo.Text = Properties.Settings.Default.EmailDefault_To;
            txtOutputEmailFrom.Text = Properties.Settings.Default.EmailDefault_From;
            txtOutputEmailCC.Text = Properties.Settings.Default.EmailDefault_Cc;
            txtOutputEmailSubject.Text = Properties.Settings.Default.EmailDefault_Subject;
            txtOutputEmailMessage.Text = Properties.Settings.Default.EmailDefault_Message;
        }

        private void chkOutputPrinter_CheckedChanged(object sender, EventArgs e)
        {
            EnableOutputPrinters(chkOutputPrinter.Checked);
        }

        private void chkOutputFolder_CheckedChanged(object sender, EventArgs e)
        {
            EnableOutputFolder(chkOutputFolder.Checked);
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = null;
            try
            {
                this.UseWaitCursor = true;
                fbd = new FolderBrowserDialog();
                fbd.ShowNewFolderButton = true;
                fbd.Description = "Select an Output folder";
                DialogResult dr = fbd.ShowDialog(this);
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    txtOutputFolder.Text = fbd.SelectedPath;
                }
            }
            finally
            {
                if (fbd != null)
                    fbd.Dispose();
                this.UseWaitCursor = false;
            }
        }

        private void chkOutputEmail_CheckedChanged(object sender, EventArgs e)
        {
            EnableOutputEmail(chkOutputEmail.Checked);
        }
        #endregion
    }
}
