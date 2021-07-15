<!--#include file="../../code/yqasp.asp" --><%
'所有日志文件默认保存在站点根目录同一父目录下，站点文件夹名称后加_log的文件夹内
'如果保存失败，请确认是否有写入权限
YQasp.Log.Enable = True
YQasp.Log.Style("info") = YQasp.Log.Style("info") & ", {note}, {param}"
YQasp.Log.Set "note", "所有的信息中都有"
YQasp.Log.SetOne "param", "只替换一次"
YQasp.Log.Info "来测试一条信息吧"
YQasp.Log.Info "再来测试一条，和上面不同哦"
YQasp.Log.Warn "使用默认的模板输出警告信息"
YQasp.Log.Error "这里出错啦", "问题出在这个文件(定位):index.asp:9"

%>