using System;
using System.Collections.Generic;
using System.Data.SqlServerCe;
using TaskScheduler;

namespace CrystalScheduler
{
    internal static class DataAccess
    {
        internal static List<ScheduledReport> GetScheduledReports()
        {
            List<ScheduledReport> scheduledReportList = new List<ScheduledReport>();

            string conString = Properties.Settings.Default.CrystalSchedulerConnectionString;
        
            using (SqlCeConnection con = new SqlCeConnection(conString))
            {
                con.Open();
                string sql = "SELECT ScheduledReportID, Report, FileName, FilePath, Name, Description FROM ScheduledReport";
                using (SqlCeCommand com = new SqlCeCommand(sql, con))
                {
                    SqlCeDataReader reader = com.ExecuteReader();
	                while (reader.Read())
	                {
                        ScheduledReport scheduledReport = new ScheduledReport();

                        scheduledReport.ScheduledReportID = reader.GetInt32(0);
                        scheduledReport.Report = reader.GetString(1);
                        scheduledReport.FileName = reader.GetString(2);
                        scheduledReport.FilePath = reader.GetString(3);
                        scheduledReport.Name = reader.GetString(4);
                        scheduledReport.Description = reader.GetString(5);

                        string sql2 = string.Format("SELECT ScheduledReportParameterID, ParameterName, ParameterValue, ParameterOrder FROM ScheduledReportParameter WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com2 = new SqlCeCommand(sql2, con))
                        {
                            SqlCeDataReader reader2 = com2.ExecuteReader();
                            while (reader2.Read())
                            {
                                ScheduledReportParameter param = new ScheduledReportParameter();
                                param.ScheduledReportParameterID = reader2.GetInt32(0);
                                param.ParameterName = reader2.GetString(1);
                                param.ParameterValue = reader2.GetString(2);
                                param.ParameterOrder = reader2.GetInt32(3);

                                scheduledReport.Parameters.Add(param);
                            }
                        }

                        string sql3 = string.Format("SELECT ScheduledReportScheduleID, IsEnabled, Frequency, StartDate, StartTime, Expires, ExpireDate, ExpireTime, Daily_RecursEveryNDays, Weekly_RecursEveryNWeeks, Weekly_WhichDaysOfWeek, Monthly_WhichMonths, Monthly_DaysOrFrequency, Monthly_Days_WhichDaysOfMonth, Monthly_Frequency_WhichFrequency, Monthly_Frequency_WhichDaysOfWeek FROM ScheduledReportSchedule WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com3 = new SqlCeCommand(sql3, con))
                        {
                            SqlCeDataReader reader3 = com3.ExecuteReader();
                            if (reader3.Read())
                            {
                                ScheduledReportSchedule sched = new ScheduledReportSchedule();
                                sched.ScheduledReportScheduleID = reader3.GetInt32(0);
                                sched.IsEnabled = reader3.GetBoolean(1);
                                sched.Frequency = reader3.GetString(2);
                                sched.StartDate = reader3.GetDateTime(3);
                                sched.StartTime = reader3.GetDateTime(4);
                                sched.Expires = reader3.GetBoolean(5);
                                sched.ExpireDate = reader3.GetDateTime(6);
                                sched.ExpireTime = reader3.GetDateTime(7);
                                sched.Daily_RecursEveryNDays = reader3.GetInt32(8);
                                sched.Weekly_RecursEveryNWeeks = reader3.GetInt32(9);
                                sched.Weekly_WhichDaysOfWeek = reader3.GetString(10);
                                sched.Monthly_WhichMonths = reader3.GetString(11);
                                sched.Monthly_DaysOrFrequency = reader3.GetString(12);
                                sched.Monthly_Days_WhichDaysOfMonth = reader3.GetString(13);
                                sched.Monthly_Frequency_WhichFrequency = reader3.GetString(14);
                                sched.Monthly_Frequency_WhichDaysOfWeek = reader3.GetString(15);

                                scheduledReport.Schedule = sched;
                            }
                        }

                        string sql4 = string.Format("SELECT ScheduledReportOutputID, SendToPrinter, Printer, SendToFolder, Folder, SendToEmail, EmailTo, EmailFrom, EmailCC, EmailSubject, EmailMessage FROM ScheduledReportOutput WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com4 = new SqlCeCommand(sql4, con))
                        {
                            SqlCeDataReader reader4 = com4.ExecuteReader();
                            if (reader4.Read())
                            {
                                ScheduledReportOutput output = new ScheduledReportOutput();
                                output.ScheduledReportOutputID = reader4.GetInt32(0);
                                output.SendToPrinter = reader4.GetBoolean(1);
                                output.Printer = reader4.GetString(2);
                                output.SendToFolder = reader4.GetBoolean(3);
                                output.Folder = reader4.GetString(4);
                                output.SendToEmail = reader4.GetBoolean(5);
                                output.EmailTo = reader4.GetString(6);
                                output.EmailFrom = reader4.GetString(7);
                                output.EmailCC = reader4.GetString(8);
                                output.EmailSubject = reader4.GetString(9);
                                output.EmailMessage = reader4.GetString(10);

                                scheduledReport.Output = output;
                            }
                        }

                        scheduledReportList.Add(scheduledReport);
	                }
                }
            }

            return scheduledReportList;
        }

        internal static ScheduledReport GetScheduledReportByID(int id)
        {
            ScheduledReport scheduledReport = new ScheduledReport();
            string conString = Properties.Settings.Default.CrystalSchedulerConnectionString;

            using (SqlCeConnection con = new SqlCeConnection(conString))
            {
                con.Open();
                string sql = "SELECT ScheduledReportID, Report, FileName, FilePath, Name, Description FROM ScheduledReport WHERE ScheduledReportID = @ID";
                using (SqlCeCommand com = new SqlCeCommand(sql, con))
                {
                    com.Parameters.AddWithValue("@ID", id);
                    SqlCeDataReader reader = com.ExecuteReader();
                    if (reader.Read())
                    {
                        scheduledReport.ScheduledReportID = reader.GetInt32(0);
                        scheduledReport.Report = reader.GetString(1);
                        scheduledReport.FileName = reader.GetString(2);
                        scheduledReport.FilePath = reader.GetString(3);
                        scheduledReport.Name = reader.GetString(4);
                        scheduledReport.Description = reader.GetString(5);

                        string sql2 = string.Format("SELECT ScheduledReportParameterID, ParameterName, ParameterValue, ParameterOrder FROM ScheduledReportParameter WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com2 = new SqlCeCommand(sql2, con))
                        {
                            SqlCeDataReader reader2 = com2.ExecuteReader();
                            while (reader2.Read())
                            {
                                ScheduledReportParameter param = new ScheduledReportParameter();
                                param.ScheduledReportParameterID = reader2.GetInt32(0);
                                param.ParameterName = reader2.GetString(1);
                                param.ParameterValue = reader2.GetString(2);
                                param.ParameterOrder = reader2.GetInt32(3);

                                scheduledReport.Parameters.Add(param);
                            }
                        }

                        string sql3 = string.Format("SELECT ScheduledReportScheduleID, IsEnabled, Frequency, StartDate, StartTime, Expires, ExpireDate, ExpireTime, Daily_RecursEveryNDays, Weekly_RecursEveryNWeeks, Weekly_WhichDaysOfWeek, Monthly_WhichMonths, Monthly_DaysOrFrequency, Monthly_Days_WhichDaysOfMonth, Monthly_Frequency_WhichFrequency, Monthly_Frequency_WhichDaysOfWeek FROM ScheduledReportSchedule WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com3 = new SqlCeCommand(sql3, con))
                        {
                            SqlCeDataReader reader3 = com3.ExecuteReader();
                            if (reader3.Read())
                            {
                                ScheduledReportSchedule sched = new ScheduledReportSchedule();
                                sched.ScheduledReportScheduleID = reader3.GetInt32(0);
                                sched.IsEnabled = reader3.GetBoolean(1);
                                sched.Frequency = reader3.GetString(2);
                                sched.StartDate = reader3.GetDateTime(3);
                                sched.StartTime = reader3.GetDateTime(4);
                                sched.Expires = reader3.GetBoolean(5);
                                sched.ExpireDate = reader3.GetDateTime(6);
                                sched.ExpireTime = reader3.GetDateTime(7);
                                sched.Daily_RecursEveryNDays = reader3.GetInt32(8);
                                sched.Weekly_RecursEveryNWeeks = reader3.GetInt32(9);
                                sched.Weekly_WhichDaysOfWeek = reader3.GetString(10);
                                sched.Monthly_WhichMonths = reader3.GetString(11);
                                sched.Monthly_DaysOrFrequency = reader3.GetString(12);
                                sched.Monthly_Days_WhichDaysOfMonth = reader3.GetString(13);
                                sched.Monthly_Frequency_WhichFrequency = reader3.GetString(14);
                                sched.Monthly_Frequency_WhichDaysOfWeek = reader3.GetString(15);

                                scheduledReport.Schedule = sched;
                            }
                        }

                        string sql4 = string.Format("SELECT ScheduledReportOutputID, SendToPrinter, Printer, SendToFolder, Folder, SendToEmail, EmailTo, EmailFrom, EmailCC, EmailSubject, EmailMessage FROM ScheduledReportOutput WHERE ScheduledReportID = {0}", scheduledReport.ScheduledReportID);
                        using (SqlCeCommand com4 = new SqlCeCommand(sql4, con))
                        {
                            SqlCeDataReader reader4 = com4.ExecuteReader();
                            if (reader4.Read())
                            {
                                ScheduledReportOutput output = new ScheduledReportOutput();
                                output.ScheduledReportOutputID = reader4.GetInt32(0);
                                output.SendToPrinter = reader4.GetBoolean(1);
                                output.Printer = reader4.GetString(2);
                                output.SendToFolder = reader4.GetBoolean(3);
                                output.Folder = reader4.GetString(4);
                                output.SendToEmail = reader4.GetBoolean(5);
                                output.EmailTo = reader4.GetString(6);
                                output.EmailFrom = reader4.GetString(7);
                                output.EmailCC = reader4.GetString(8);
                                output.EmailSubject = reader4.GetString(9);
                                output.EmailMessage = reader4.GetString(10);

                                scheduledReport.Output = output;
                            }
                        }
                    }
                }
            }

            return scheduledReport;
        }

        internal static void SaveScheduledReport(ScheduledReport scheduledReport)
        {
            if (scheduledReport.ScheduledReportID == 0)
            {
                CreateWindowsTask(scheduledReport);
                InsertScheduledReport(scheduledReport);
            }
            else
            {
                UpdateWindowsTask(scheduledReport);
                UpdateScheduledReport(scheduledReport);
            }
        }

        private static void CreateWindowsTask(ScheduledReport scheduledReport)
        {
            ITaskService taskService = new TaskScheduler.TaskScheduler();
            taskService.Connect();

            ITaskDefinition taskDefinition = taskService.NewTask(0);
            taskDefinition.RegistrationInfo.Description = scheduledReport.Description;
            taskDefinition.Settings.Enabled = true;
            taskDefinition.Settings.Hidden = true;
            taskDefinition.Settings.Compatibility = _TASK_COMPATIBILITY.TASK_COMPATIBILITY_V2_1;
            taskDefinition.Settings.WakeToRun = true;

            ITriggerCollection triggers = taskDefinition.Triggers;
            ITrigger trigger;
            switch (scheduledReport.Schedule.Frequency)
            {
                case "Daily":
                    trigger = triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_DAILY);
                    //trigger.Repetition = scheduledReport.Schedule.Daily_RecursEveryNDays;
                    break;
                case "Weekly":
                    break;
                case "Monthly":
                    break;
            }
            

            IActionCollection actions = taskDefinition.Actions;
            _TASK_ACTION_TYPE actionType = _TASK_ACTION_TYPE.TASK_ACTION_EXEC;
            IAction action = actions.Create(actionType);
            IExecAction execAction = action as IExecAction;
            execAction.Path = @"C:\HutsonSoftware\CrystalScheduler\Reporter.exe";
            execAction.Arguments = scheduledReport.TaskName;

            ITaskFolder crystalSchedulerFolder = GetCrystalSchedulerTaskFolder(taskService);
            crystalSchedulerFolder.RegisterTaskDefinition(
                scheduledReport.TaskName, 
                taskDefinition, 
                6, 
                null, 
                null, 
                _TASK_LOGON_TYPE.TASK_LOGON_NONE, 
                null);
        }

        private static ITaskFolder GetCrystalSchedulerTaskFolder(ITaskService taskService)
        {
            ITaskFolder crystalSchedulerFolder, hutSoftFolder;

            try
            {
                hutSoftFolder = taskService.GetFolder(@"\HutsonSoftware");
            }
            catch
            {
                taskService.GetFolder(@"\").CreateFolder(@"HutsonSoftware");
            }

            try
            {
                crystalSchedulerFolder = taskService.GetFolder(@"\HutsonSoftware\CrystalScheduler");
            }
            catch
            {
                taskService.GetFolder(@"\").CreateFolder(@"HutsonSoftware\CrystalScheduler");
                crystalSchedulerFolder = taskService.GetFolder(@"\HutsonSoftware\CrystalScheduler");
            }

            return crystalSchedulerFolder;
        }

        private static void UpdateWindowsTask(ScheduledReport scheduledReport)
        {
            
        }

        private static void InsertScheduledReport(ScheduledReport scheduledReport)
        {
            string conString = Properties.Settings.Default.CrystalSchedulerConnectionString;

            int scheduledReportID = 0;
            using (SqlCeConnection con = new SqlCeConnection(conString))
            {
                con.Open();

                using (SqlCeCommand com = new SqlCeCommand("INSERT ScheduledReport (Report, FileName, FilePath, Name, Description) VALUES (@Report, @FileName, @FilePath, @Name, @Description)", con))
                {
                    com.Parameters.AddWithValue("@Report", scheduledReport.Report);
                    com.Parameters.AddWithValue("@FileName", scheduledReport.FileName);
                    com.Parameters.AddWithValue("@FilePath", scheduledReport.FilePath);
                    com.Parameters.AddWithValue("@Name", scheduledReport.Name);
                    com.Parameters.AddWithValue("@Description", scheduledReport.Description);
                    com.ExecuteNonQuery();

                    com.CommandText = "SELECT @@IDENTITY";
                    scheduledReportID = Convert.ToInt32((decimal)com.ExecuteScalar());
                }

                foreach (ScheduledReportParameter scheduledReportParameter in scheduledReport.Parameters)
                {
                    using (SqlCeCommand com = new SqlCeCommand("INSERT ScheduledReportParameter (ScheduledReportID, ParameterName, ParameterValue, ParameterOrder) VALUES (@ScheduledReportID, @ParameterName, @ParameterValue, @ParameterOrder)", con))
                    {
                        com.Parameters.AddWithValue("@ScheduledReportID", scheduledReportID);
                        com.Parameters.AddWithValue("@ParameterName", scheduledReportParameter.ParameterName);
                        com.Parameters.AddWithValue("@ParameterValue", scheduledReportParameter.ParameterValue);
                        com.Parameters.AddWithValue("@ParameterOrder", scheduledReportParameter.ParameterOrder);
                        com.ExecuteNonQuery();
                    }
                }

                using (SqlCeCommand com = new SqlCeCommand("INSERT ScheduledReportSchedule (ScheduledReportID, IsEnabled, Frequency, StartDate, StartTime, Expires, ExpireDate, ExpireTime, Daily_RecursEveryNDays, Weekly_RecursEveryNWeeks, Weekly_WhichDaysOfWeek, Monthly_WhichMonths, Monthly_DaysOrFrequency, Monthly_Days_WhichDaysOfMonth, Monthly_Frequency_WhichFrequency, Monthly_Frequency_WhichDaysOfWeek) VALUES (@ScheduledReportID, @IsEnabled, @Frequency, @StartDate, @StartTime, @Expires, @ExpireDate, @ExpireTime, @Daily_RecursEveryNDays, @Weekly_RecursEveryNWeeks, @Weekly_WhichDaysOfWeek, @Monthly_WhichMonths, @Monthly_DaysOrFrequency, @Monthly_Days_WhichDaysOfMonth, @Monthly_Frequency_WhichFrequency, @Monthly_Frequency_WhichDaysOfWeek)", con))
                {
                    com.Parameters.AddWithValue("@ScheduledReportID", scheduledReportID);
                    com.Parameters.AddWithValue("@IsEnabled", scheduledReport.Schedule.IsEnabled);
                    com.Parameters.AddWithValue("@Frequency", scheduledReport.Schedule.Frequency);
                    com.Parameters.AddWithValue("@StartDate", scheduledReport.Schedule.StartDate);
                    com.Parameters.AddWithValue("@StartTime", scheduledReport.Schedule.StartTime);
                    com.Parameters.AddWithValue("@Expires", scheduledReport.Schedule.Expires);
                    com.Parameters.AddWithValue("@ExpireDate", scheduledReport.Schedule.ExpireDate);
                    com.Parameters.AddWithValue("@ExpireTime", scheduledReport.Schedule.ExpireTime);
                    com.Parameters.AddWithValue("@Daily_RecursEveryNDays", scheduledReport.Schedule.Daily_RecursEveryNDays);
                    com.Parameters.AddWithValue("@Weekly_RecursEveryNWeeks", scheduledReport.Schedule.Weekly_RecursEveryNWeeks);
                    com.Parameters.AddWithValue("@Weekly_WhichDaysOfWeek", scheduledReport.Schedule.Weekly_WhichDaysOfWeek);
                    com.Parameters.AddWithValue("@Monthly_WhichMonths", scheduledReport.Schedule.Monthly_WhichMonths);
                    com.Parameters.AddWithValue("@Monthly_DaysOrFrequency", scheduledReport.Schedule.Monthly_DaysOrFrequency);
                    com.Parameters.AddWithValue("@Monthly_Days_WhichDaysOfMonth", scheduledReport.Schedule.Monthly_Days_WhichDaysOfMonth);
                    com.Parameters.AddWithValue("@Monthly_Frequency_WhichFrequency", scheduledReport.Schedule.Monthly_Frequency_WhichFrequency);
                    com.Parameters.AddWithValue("@Monthly_Frequency_WhichDaysOfWeek", scheduledReport.Schedule.Monthly_Frequency_WhichDaysOfWeek);
                    com.ExecuteNonQuery();
                }

                using (SqlCeCommand com = new SqlCeCommand("INSERT ScheduledReportOutput (ScheduledReportID, SendToPrinter, Printer, SendToFolder, Folder, SendToEmail, EmailTo, EmailFrom, EmailCC, EmailSubject, EmailMessage) VALUES (@ScheduledReportID, @SendToPrinter, @Printer, @SendToFolder, @Folder, @SendToEmail, @EmailTo, @EmailFrom, @EmailCC, @EmailSubject, @EmailMessage)", con))
                {
                    com.Parameters.AddWithValue("@ScheduledReportID", scheduledReportID);
                    com.Parameters.AddWithValue("@SendToPrinter", scheduledReport.Output.SendToPrinter);
                    com.Parameters.AddWithValue("@Printer", scheduledReport.Output.Printer);
                    com.Parameters.AddWithValue("@SendToFolder", scheduledReport.Output.SendToFolder);
                    com.Parameters.AddWithValue("@Folder", scheduledReport.Output.Folder);
                    com.Parameters.AddWithValue("@SendToEmail", scheduledReport.Output.SendToEmail);
                    com.Parameters.AddWithValue("@EmailTo", scheduledReport.Output.EmailTo);
                    com.Parameters.AddWithValue("@EmailFrom", scheduledReport.Output.EmailFrom);
                    com.Parameters.AddWithValue("@EmailCC", scheduledReport.Output.EmailCC);
                    com.Parameters.AddWithValue("@EmailSubject", scheduledReport.Output.EmailSubject);
                    com.Parameters.AddWithValue("@EmailMessage", scheduledReport.Output.EmailMessage);
                    com.ExecuteNonQuery();
                }

            }
        }

        private static void UpdateScheduledReport(ScheduledReport scheduledReport)
        {
            string conString = Properties.Settings.Default.CrystalSchedulerConnectionString;

            using (SqlCeConnection con = new SqlCeConnection(conString))
            {
                con.Open();

                using (SqlCeCommand com = new SqlCeCommand("UPDATE ScheduledReport SET Report = @Report, FileName = @FileName, FilePath = @FilePath, Name = @Name, Description = @Description WHERE ScheduledReportID = @ScheduledReportID", con))
                {
                    com.Parameters.AddWithValue("@Report", scheduledReport.Report);
                    com.Parameters.AddWithValue("@FileName", scheduledReport.FileName);
                    com.Parameters.AddWithValue("@FilePath", scheduledReport.FilePath);
                    com.Parameters.AddWithValue("@Name", scheduledReport.Name);
                    com.Parameters.AddWithValue("@Description", scheduledReport.Description);
                    com.Parameters.AddWithValue("@ScheduledReportID", scheduledReport.ScheduledReportID);
                    com.ExecuteNonQuery();
                }

                foreach (ScheduledReportParameter scheduledReportParameter in scheduledReport.Parameters)
                {
                    if (scheduledReportParameter.ScheduledReportParameterID == 0)
                        using (SqlCeCommand com = new SqlCeCommand("INSERT ScheduledReportParameter (ScheduledReportID, ParameterName, ParameterValue, ParameterOrder) VALUES (@ScheduledReportID, @ParameterName, @ParameterValue, @ParameterOrder)", con))
                        {
                            com.Parameters.AddWithValue("@ScheduledReportID", scheduledReport.ScheduledReportID);
                            com.Parameters.AddWithValue("@ParameterName", scheduledReportParameter.ParameterName);
                            com.Parameters.AddWithValue("@ParameterValue", scheduledReportParameter.ParameterValue);
                            com.Parameters.AddWithValue("@ParameterOrder", scheduledReportParameter.ParameterOrder);
                            com.ExecuteNonQuery();
                        }
                    else
                        using (SqlCeCommand com = new SqlCeCommand("UPDATE ScheduledReportParameter SET ParameterName = @ParameterName, ParameterValue = @ParameterValue, ParameterOrder = @ParameterOrder WHERE ScheduledReportParameterID = @ScheduledReportParameterID", con))
                        {
                            com.Parameters.AddWithValue("@ParameterName", scheduledReportParameter.ParameterName);
                            com.Parameters.AddWithValue("@ParameterValue", scheduledReportParameter.ParameterValue);
                            com.Parameters.AddWithValue("@ParameterOrder", scheduledReportParameter.ParameterOrder);
                            com.Parameters.AddWithValue("@ScheduledReportParameterID", scheduledReportParameter.ScheduledReportParameterID);
                            com.ExecuteNonQuery();
                        }
                }

                using (SqlCeCommand com = new SqlCeCommand("UPDATE ScheduledReportSchedule SET IsEnabled = @IsEnabled, Frequency = @Frequency, StartDate = @StartDate, StartTime = @StartTime, Expires = @Expires, ExpireDate = @ExpireDate, ExpireTime = @ExpireTime, Daily_RecursEveryNDays = @Daily_RecursEveryNDays, Weekly_RecursEveryNWeeks = @Weekly_RecursEveryNWeeks, Weekly_WhichDaysOfWeek = @Weekly_WhichDaysOfWeek, Monthly_WhichMonths = @Monthly_WhichMonths, Monthly_DaysOrFrequency = @Monthly_DaysOrFrequency, Monthly_Days_WhichDaysOfMonth = @Monthly_Days_WhichDaysOfMonth, Monthly_Frequency_WhichFrequency = @Monthly_Frequency_WhichFrequency, Monthly_Frequency_WhichDaysOfWeek = @Monthly_Frequency_WhichDaysOfWeek WHERE ScheduledReportScheduleID = @ScheduledReportScheduleID)", con))
                {
                    com.Parameters.AddWithValue("@IsEnabled", scheduledReport.Schedule.IsEnabled);
                    com.Parameters.AddWithValue("@Frequency", scheduledReport.Schedule.Frequency);
                    com.Parameters.AddWithValue("@StartDate", scheduledReport.Schedule.StartDate);
                    com.Parameters.AddWithValue("@StartTime", scheduledReport.Schedule.StartTime);
                    com.Parameters.AddWithValue("@Expires", scheduledReport.Schedule.Expires);
                    com.Parameters.AddWithValue("@ExpireDate", scheduledReport.Schedule.ExpireDate);
                    com.Parameters.AddWithValue("@ExpireTime", scheduledReport.Schedule.ExpireTime);
                    com.Parameters.AddWithValue("@Daily_RecursEveryNDays", scheduledReport.Schedule.Daily_RecursEveryNDays);
                    com.Parameters.AddWithValue("@Weekly_RecursEveryNWeeks", scheduledReport.Schedule.Weekly_RecursEveryNWeeks);
                    com.Parameters.AddWithValue("@Weekly_WhichDaysOfWeek", scheduledReport.Schedule.Weekly_WhichDaysOfWeek);
                    com.Parameters.AddWithValue("@Monthly_WhichMonths", scheduledReport.Schedule.Monthly_WhichMonths);
                    com.Parameters.AddWithValue("@Monthly_DaysOrFrequency", scheduledReport.Schedule.Monthly_DaysOrFrequency);
                    com.Parameters.AddWithValue("@Monthly_Days_WhichDaysOfMonth", scheduledReport.Schedule.Monthly_Days_WhichDaysOfMonth);
                    com.Parameters.AddWithValue("@Monthly_Frequency_WhichFrequency", scheduledReport.Schedule.Monthly_Frequency_WhichFrequency);
                    com.Parameters.AddWithValue("@Monthly_Frequency_WhichDaysOfWeek", scheduledReport.Schedule.Monthly_Frequency_WhichDaysOfWeek);
                    com.Parameters.AddWithValue("@ScheduledReportScheduleID", scheduledReport.Schedule.ScheduledReportScheduleID);
                    com.ExecuteNonQuery();
                }

                using (SqlCeCommand com = new SqlCeCommand("UPDATE ScheduledReportOutput SET SendToPrinter = @SendToPrinter, Printer = @Printer, SendToFolder = @SendToFolder, Folder = @Folder, SendToEmail = @SendToEmail, EmailTo = @EmailTo, EmailFrom = @EmailFrom, EmailCC = @EmailCC, EmailSubject = @EmailSubject, EmailMessage = @EmailMessage WHERE ScheduledReportOutputID = @ScheduledReportOutputID", con))
                {
                    com.Parameters.AddWithValue("@SendToPrinter", scheduledReport.Output.SendToPrinter);
                    com.Parameters.AddWithValue("@Printer", scheduledReport.Output.Printer);
                    com.Parameters.AddWithValue("@SendToFolder", scheduledReport.Output.SendToFolder);
                    com.Parameters.AddWithValue("@Folder", scheduledReport.Output.Folder);
                    com.Parameters.AddWithValue("@SendToEmail", scheduledReport.Output.SendToEmail);
                    com.Parameters.AddWithValue("@EmailTo", scheduledReport.Output.EmailTo);
                    com.Parameters.AddWithValue("@EmailFrom", scheduledReport.Output.EmailFrom);
                    com.Parameters.AddWithValue("@EmailCC", scheduledReport.Output.EmailCC);
                    com.Parameters.AddWithValue("@EmailSubject", scheduledReport.Output.EmailSubject);
                    com.Parameters.AddWithValue("@EmailMessage", scheduledReport.Output.EmailMessage);
                    com.Parameters.AddWithValue("@ScheduledReportID", scheduledReport.Output.ScheduledReportOutputID);
                    com.ExecuteNonQuery();
                }

            }
        }

        internal static void DuplicateScheduledReport(int id)
        {
            SaveScheduledReport(GetScheduledReportByID(id));
        }

        internal static void DeleteScheduledReport(int id)
        {
            string conString = Properties.Settings.Default.CrystalSchedulerConnectionString;
            using (SqlCeConnection con = new SqlCeConnection(conString))
            {
                con.Open();

                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM ScheduledReportParameter WHERE ScheduledReportID = @ID", con))
                {
                    com.Parameters.AddWithValue("@ID", id);
                    com.ExecuteNonQuery();
                }
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM ScheduledReportSchedule WHERE ScheduledReportID = @ID", con))
                {
                    com.Parameters.AddWithValue("@ID", id);
                    com.ExecuteNonQuery();
                }
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM ScheduledReportOutput WHERE ScheduledReportID = @ID", con))
                {
                    com.Parameters.AddWithValue("@ID", id);
                    com.ExecuteNonQuery();
                }
                using (SqlCeCommand com = new SqlCeCommand("DELETE FROM ScheduledReport WHERE ScheduledReportID = @ID", con))
                {
                    com.Parameters.AddWithValue("@ID", id);
                    com.ExecuteNonQuery();
                }
            }
        }
    }

}
