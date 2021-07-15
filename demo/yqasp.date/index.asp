<!--#include file="../../code/yqasp.asp" --><%
Dim datetime
datetime = "2013-12-22 23:34:45"
YQasp.print("给定日期是：")
YQasp.println YQasp.Date.Format(datetime, "y-mm-dd hh:ii:ss 星期w")
YQasp.print("从现在算起是：")
YQasp.println YQasp.Date.Format(datetime, Now)
YQasp.print("这个月的第一天是：")
YQasp.println YQasp.Date.FirstDayOfMonth(datetime)
YQasp.print("这个月的最后一天是：")
YQasp.println YQasp.Date.LastDayOfMonth(datetime)
YQasp.print("如果每周从星期一开始，这周的第一天是：")
YQasp.println YQasp.Date.FirstDayOfWeek(datetime)
'从星期日开始一周
YQasp.Date.WeekStarting = 1
YQasp.print("如果每周从星期日开始，这周的最后一天是：")
YQasp.println YQasp.Date.LastDayOfWeek(datetime)
YQasp.print("这个时间转成Unix时间戳是：")
YQasp.println YQasp.Date.ToUnixTimeCn(datetime)
%>