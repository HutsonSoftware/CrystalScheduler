using System;
using System.Collections.Generic;

namespace CrystalScheduler
{
    public class ScheduledReport
    {
        public ScheduledReport()
        {
            ScheduledReportID = 0;
            Report = string.Empty;
            FileName = string.Empty;
            FilePath = string.Empty;
            Name = string.Empty;
            Description = string.Empty;
            TaskName = string.Empty;
            Parameters = new List<ScheduledReportParameter>();
            Schedule = new ScheduledReportSchedule();
            Output = new ScheduledReportOutput();
        }
        public int ScheduledReportID { get; set; }
        private string _report = string.Empty;
        public string Report 
        {
            get
            {
                return _report;
            }
            set
            {
                _report = value + "_" + DateTime.Now.ToShortDateString() + DateTime.Now.ToShortTimeString();
            }
        }
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        private string _taskName = string.Empty;
        public string TaskName 
        { 
            get 
            { 
                if (_taskName == string.Empty)
                    _taskName = Report; 
                return _taskName; 
            }
            set
            {
                _taskName = value;
            }
        }

        public List<ScheduledReportParameter> Parameters { get; set; }
        public ScheduledReportSchedule Schedule { get; set; }
        public ScheduledReportOutput Output { get; set; }
    }

    public class ScheduledReportParameter
    {
        public int ScheduledReportParameterID { get; set; }
        public string ParameterName { get; set; }
        public string ParameterValue { get; set; }
        public int ParameterOrder { get; set; }
    }

    public class ScheduledReportSchedule
    {
        public int ScheduledReportScheduleID { get; set; }
        public bool IsEnabled { get; set; }
        public string Frequency { get; set; } //{ OneTime, Daily, Weekly, Monthly }
        public DateTime StartDate { get; set; }
        public DateTime StartTime { get; set; }
        public bool Expires { get; set; }
        public DateTime ExpireDate { get; set; }
        public DateTime ExpireTime { get; set; }
        
        public int Daily_RecursEveryNDays { get; set; }
        
        public int Weekly_RecursEveryNWeeks { get; set; }
        public string Weekly_WhichDaysOfWeek { get; set; } //{ Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday }
        
        public string Monthly_WhichMonths { get; set; } //{ January, February, March, April, May, June, July, August, September, October, November, December }
        public string Monthly_DaysOrFrequency { get; set; } //{ Days, Frequency }
        public string Monthly_Days_WhichDaysOfMonth { get; set; } //{ 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, Last }
        public string Monthly_Frequency_WhichFrequency { get; set; } //{ First, Second, Third, Fourth, Last }
        public string Monthly_Frequency_WhichDaysOfWeek { get; set; } //{ Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday }
    }

    public class ScheduledReportOutput
    {
        public int ScheduledReportOutputID { get; set; }
        public bool SendToPrinter { get; set; }
        public string Printer { get; set; }
        public bool SendToFolder { get; set; }
        public string Folder { get; set; }
        public bool SendToEmail { get; set; }
        public string EmailTo { get; set; }
        public string EmailFrom { get; set; }
        public string EmailCC { get; set; }
        public string EmailSubject { get; set; }
        public string EmailMessage { get; set; }
    }
}
