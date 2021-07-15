<!--#include file="../../../code/yqasp.asp" -->
<%
Dim File,F
YQasp.Upload.AllowFileTypes = "jpg|png|gif"
YQasp.Upload.AllowMaxFileSize = "1MB"
YQasp.Upload.AllowMaxSize = "20mb"
YQasp.Upload.CharSet = "utf-8"
if not YQasp.Upload.GetData() then 
	YQasp.PrintEnd YQasp.Upload.Description
else
	YQasp.Upload.SavePath = "/_upload"
	YQasp.Println "<b>保存所有文件： </b>"
	for each file in YQasp.Upload.Files("-1")
		Set F = YQasp.Upload.Save(file,0,true)
		if F.Succeed then
			YQasp.Println "文件'" & F.LocalName & "'上传成功，保存位置'" & F.Path & F.filename & "',文件大小" & F.size & "字节"
		else
			YQasp.Println F.Exception
		end if		
	next
end if
%>