<!--#include file="../../code/yqasp.asp" --><%
'YQasp.Debug = False
'YQasp.Error.Redirect = False
YQasp.Error.OnErrorContinue = True
YQasp.Error.ConsoleDetail = False
YQasp.Println Link("", YQasp.GetUrl(-3) & "?" & YQasp.NewID, "")
Dim s
'YQasp.Console "[Error]数据库读取错误。"
'YQasp.Console YQasp.Error.Debug

'For Each s In Request.ServerVariables
'  YQasp.Println s & " : " & Request.ServerVariables(s)
'Next
On Error Resume Next
'YQasp.Db.SetConn 0, "YQasp", "sa:pass@(local))"
YQasp.Ext("check").Meinv
Dim conn
'Set conn = YQasp.Db.GetConn()
'Err.Raise 45, "my error"
'YQasp.Console YQasp.Error.Redirect
'YQasp.Error.Detail = "(sa:pass@(local))"
'YQasp.Error.Raise 12

Function Link(ByVal string, ByVal url, ByVal attr)
  Dim a
  a = YQasp.IfHas(string, url)
  Link = "<a href=""" & url & """" & YQasp.IfThen(YQasp.Has(attr), " " & attr) & ">" & a & "</a>"
End Function

YQasp.Println "============================"
YQasp.Println YQasp.GetScriptTime & "s"

%>
