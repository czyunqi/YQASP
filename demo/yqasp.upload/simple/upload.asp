<!--#include file="../../../code/yqasp.asp" -->
<%
dim File
YQasp.Var("test1") = "test1"
YQasp.Upload.AllowFileTypes = "*.jpg"
YQasp.Upload.AllowMaxFileSize = "10MB"
YQasp.Upload.AllowMaxSize = "20mb"
YQasp.Upload.CharSet = "utf-8"
YQasp.Println "YQasp.Var(""form1"") => " & YQasp.Var("form1")  '在调用 YQasp.Upload.GetData() 之前是取不到表单数据的
YQasp.Println "YQasp.Var(""act"") => " & YQasp.Var("act") 'querystring的值则可以随时调用
YQasp.Println "YQasp.Var(""test1"") => " & YQasp.Var("test1")
if not YQasp.Upload.GetData() then 
	YQasp.Println YQasp.Upload.Description
else
  YQasp.Var("test2") = "test2"
	YQasp.Upload.SavePath = "/_upload"
	YQasp.Println "YQasp.Var(""test2"") => " & YQasp.Var("test2")
	YQasp.Println "YQasp.Post(""form1"") => " & YQasp.Post("form1")
	YQasp.Println "YQasp.Db.ToSql(""delete from T where Tname in ({(form1)})"") =>" & _
	             YQasp.Db.ToSql("delete from T where Tname in ({(form1)})")
	Set File = YQasp.Upload.Save("file1",0,true)
	if File.Succeed then
		YQasp.Println "文件'" & File.LocalName & "'上传成功，保存位置'" & File.Path & File.FileName & "',文件大小" & File.Size & "字节"
	else
		YQasp.Println File.Exception & "<br />"
	end if
end if
%>