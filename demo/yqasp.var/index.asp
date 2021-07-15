<!--#include file="../../code/yqasp.asp" --><%
YQasp.NoCache()
YQasp.Var("myname") = "Lin"
If YQasp.Has(Request.Form) Then
  YQasp.println "YQasp.Var(""yqasp.newid"") : " & YQasp.Var("yqasp.newid")
  YQasp.println "YQasp.Var(""url"") : " & YQasp.Var("url")
  YQasp.println "YQasp.Var(""myname"") : " & YQasp.Var("myname")
  YQasp.println "YQasp.Var(""get.username"") : " & YQasp.Var("get.username")
  YQasp.println "YQasp.Var(""post.username"") : " & YQasp.Var("post.username")
  YQasp.println "YQasp.Var(""username"") : " & YQasp.Var("username")
  YQasp.println "YQasp.Var(""msg"") : " & YQasp.Var("msg")
  YQasp.println "YQasp.Var(""action"") : " & YQasp.Var("action")
  If YQasp.Var.Has("action_array") Then
    'YQasp.print "如果同一名称URL参数有多个值：Request.QueryString(""action"").Count : "
    'YQasp.println Request.QueryString("action").Count
    YQasp.println "YQasp.Var(""action_array"") : " & YQasp.Str.ToString(YQasp.Var("action_array"))
  End If
  YQasp.println "YQasp.Var(""type"") : " & YQasp.Var("type")
  If YQasp.Var.Has("type_array") Then
    'YQasp.print "如果同一名称表单有多个值：Request.Form(""type"").Count : "
    'YQasp.println Request.Form("type").Count
    YQasp.println "YQasp.Var(""type_array"") : " & YQasp.Str.ToString(YQasp.Var("type_array"))
  End If
  YQasp.Println "YQasp.Var(""server.remote_addr"") : " & YQasp.Var("server.remote_addr")
  YQasp.Println "YQasp.Var(""server.http_user_agent"") : " & YQasp.Var("server.http_user_agent")
  '显示所有的变量
  YQasp.println "=============================="
  YQasp.println "遍历所有的EasyAsp变量："
  Dim vars, key 
  Set vars = YQasp.Var.GetObject()
  For Each key In vars
    YQasp.print "YQasp.Var(""" & key & """) : "
    YQasp.println YQasp.Str.ToString(vars(key))
  Next
  Set vars = Nothing
End If
%>
<form action="?action=save&username=coldstone&action=update" method="post">
  username: <input type="text" size="60" name="username" value="ray" /><br />
  msg: <input type="text" size="60" name="msg" value="I'm here" /><br />
  <input type="checkbox" name="type" value="1" checked="checked" />type1
  <input type="checkbox" name="type" value="2" checked="checked" />type2<br />
  <button type="submit">Submit to "?action=save&username=coldstone&action=update"</button>
</form>