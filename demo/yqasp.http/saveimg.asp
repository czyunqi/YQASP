<!--#include file="../../code/yqasp.asp" --><%
Dim tmp

''=========================
''Demo 7 - 保存远程图片：
YQasp.Http.Get "http://www.cnbeta.com/articles/280317.htm"
tmp = YQasp.Http.SaveImgTo("imgatlocal/")
YQasp.Println YQasp.Str.HtmlEncode(tmp)
''=========================
  
YQasp.Println ""
YQasp.Println "------------------------------------"
YQasp.Print "页面执行时间： " & YQasp.GetScriptTime & " 秒"
%>