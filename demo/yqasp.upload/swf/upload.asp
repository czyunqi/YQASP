<!--#include file="../../../code/yqasp.asp" -->
<%
Dim File
YQasp.Upload.AllowMaxSize="200mb"
YQasp.Upload.AllowMaxFileSize="200mb"
YQasp.Upload.AllowFileTypes="*.*" 
YQasp.Upload.Charset="utf-8"
if not YQasp.Upload.GetData() then
	YQasp.Echo "{err:true,msg:'" & YQasp.Upload.Description & "'}"
else
	YQasp.Upload.SavePath = "/_upload"
	set File=YQasp.Upload.files("filedata") 
	if YQasp.Upload.Save(File,0,true).Succeed then
		YQasp.Echo "{err:false,msg:'upload',name:'" & File.filename & "',src:'" & File.LocalName & "',name2:'" & YQasp.Upload.Post("name") & "'}"
	else
		YQasp.Echo "{err:true,msg:'" & File.Exception & "'}"
	end if
	set File=nothing
end if
%>
