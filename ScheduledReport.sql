--ScheduledReport
ScheduledReportID
Report
FileName
FilePath
Name
Description

--ScheduledReportParameter
ScheduledReportParameterID
ScheduledReportID
ParameterName
ParameterValue
ParameterOrder

--ScheduledReportSchedule
ScheduledReportScheduleID
ScheduledReportID
IsEnabled
Frequency [OneTime,Daily,Weekly,Monthly]
StartDate
StartTime
Expires
ExpireDate
ExpireTime
RecursEveryNDays
RecursEveryNWeeks
DaysOfWeek [Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday]
Months [January,February,March,April,May,June,July,August,September,October,November,December]
Days [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,Last]
OnFrequency [First,Second,Third,Fourth,Last]
OnDay [Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday]

--ScheduledReportOutput
ScheduledReportOutputID
ScheduledReportID
SendToPrinter
Printer
SendToFolder
Folder
SendToEmail
EmailTo
EmailFrom
EmailCC
EmailSubject
EmailMessage
