<!--#include file="../../code/yqasp.asp" --><%

'注意：部分代码需数据库支持，请设置正常的测试数据库连接并修改下面例子中55行的获取记录集的代码

'清除所有缓存
'YQasp.Cache.Clear
'YQasp.PrintEnd "Cache Cleared."

'不统计缓存总数(YQasp.Cache.Count将无法取到YQasp缓存总数)
'YQasp.Cache.CountEnabled = False

'文件缓存的保存路径(默认为/_cache)
'YQasp.Cache.SavePath = "/_cache"
'文件缓存的保存文件扩展名
YQasp.Cache.FileType = ".cache"

'统一缓存的保存时间，单位为分钟，不设置默认为5分钟
'YQasp.Cache.Expires = 5
'或设置为具体的过期时间：
'YQasp.Cache.Expires = "2014/10/01 08:00:00"

'缓存字符串到文件缓存
YQasp.Println "字符串缓存到文件缓存示例："
Dim tmp
'单独设置某一缓存的过期时间
YQasp.Cache("test").Expires = 3
YQasp.Println "此缓存的过期时间被设置为: " & YQasp.Cache("test").Expires & " 分钟后"
If YQasp.Cache("test").Ready Then
'如果缓存存在且没有过期
	tmp = YQasp.Cache("test")
	YQasp.Println "已读取缓存(test):"
Else
'如果缓存不存在或已过期
	'赋值给缓存对象
	tmp = "<i>测试字符串</i> 保存时间为 (" & Now() & ")"
	YQasp.Cache("test") = tmp
	'保存缓存到文件缓存（注意：保存为文件缓存还是内存缓存，区别只有一点，就是使用Save还是SaveApp，它们的获取方式是一样的）
	YQasp.Cache("test").Save
	YQasp.Println "已保存缓存(test)."
End If
YQasp.Println "---------"
YQasp.Println tmp
YQasp.Println "======================================"

''缓存记录集到文件缓存
'YQasp.Println "记录集缓存到文件缓存示例："
'Dim rs
''缓存过期时间为1分钟
'YQasp.Cache("test/rs").Expires = 1
'If YQasp.Cache("test/rs").Ready Then
'	'读取文件缓存中的记录集(不需要Set)
'	rs = YQasp.Cache("test/rs")
'	YQasp.Println "已读取缓存(test/rs):"
'Else
'	Set rs = YQasp.Db.Query("Select * From ShopList Order By ID Desc") '这里要换成你自己的数据库相关内容
'	YQasp.Cache("test/rs") = rs
'	'保存记录集到文件缓存，如果将记录集保存到内存缓存（.SaveApp）的话，会自动存为二维数组(GetRows)
'	YQasp.Cache("test/rs").Save
'	YQasp.Println "已保存缓存(test/rs)."
'End If
'YQasp.Println "---------"
'If YQasp.Has(rs) Then
'	While Not rs.EOF
'		YQasp.Println "【" & rs(0) & "】" & rs(1) & " ( "&rs(2)&" )"
'		rs.MoveNext
'	Wend
'Else
'	YQasp.Println "记录集为空"
'End If
'YQasp.Db.Close(rs)

YQasp.Println "======================================"

'缓存Dictionary对象到内存缓存
YQasp.Println "字典对象缓存示例："
Dim dict, key
YQasp.Cache("test/dict").Expires = 1
If YQasp.Cache("test/dict").Ready Then
	dict = YQasp.Cache("test/dict")
	YQasp.Println "已读取缓存(test/dict):"
Else
	Set dict = Server.CreateObject("Scripting.Dictionary")
	dict.add "a", "aaaaa"
	dict.add "b", "bbbbb"
	YQasp.Cache("test/dict") = dict
	'保存到内存缓存用SaveApp
	YQasp.Cache("test/dict").SaveApp
	YQasp.Println "已保存缓存(test/dict)."
End If
YQasp.Println "---------"
'列出字典内容
For Each key In dict
	YQasp.Println key & ":" & dict(key)
Next
Set dict = Nothing

YQasp.Println "======================================"
'缓存数量
YQasp.Println "总共有缓存：" & YQasp.Cache.Count & "个"
'遍历缓存名称
Dim caches,ckey : Set caches = YQasp.GetApplication("YQasp_Cache_Count")(1)
For Each ckey In caches
	YQasp.Println ckey
Next
Set caches = Nothing

YQasp.Println "------------------------------------"
YQasp.Print "页面执行时间： " & YQasp.GetScriptTime & " 秒, 共查询数据库： " & YQasp.Db.QueryTimes & " 次"
%>