<!--#include file="../../code/yqasp.asp" --><%
Dim http, tmp, rule, arr, i
''=========================
''Demo 1 - 最简单的应用(Get)：
'直接获取页面源码
tmp = YQasp.Http.Get("http://www.yqasp.cn")
YQasp.Println YQasp.Str.HtmlEncode(tmp)
''=========================

''=========================
''Demo 2 - 最简单的Post：
'YQasp.Http.Data = Array("search_type:0","keyword:月关")
'tmp = YQasp.Http.Post("http://book.2345.com/search.html")
'YQasp.Println YQasp.Str.HtmlEncode(tmp)
''=========================

''=========================
''Demo 3 - 通过属性配置：
'Set http = YQasp.Http.New()
''http.ResolveTimeout = 20000	'服务器解析超时时间，毫秒，默认20秒
''http.ConnectTimeout = 20000	'服务器连接超时时间，毫秒，默认20秒
''http.SendTimeout = 300000		'发送数据超时时间，毫秒，默认5分钟
''http.ReceiveTimeout = 60000	'接受数据超时时间，毫秒，默认1分钟
'http.Url = "http://book.2345.com/search.html"	'目标URL地址
'http.Method = "GET"  'GET 或者 POST, 默认GET
''目标文件编码，一般不用设置此属性，YQasp会自动判断目标地址的编码
''http.CharSet = "gb2312"
'http.Async = False	'异步，默认False，建议不要修改
''数据提交方式一，如果是GET则会附在URL后以参数形式提交：
'http.Data = "search_type=0&keyword=月关"
''数据提交方式二，可以用Array参数的方式提交：
''http.Data = Array("search_type:0","keyword:月关")
''http.User = ""	'如果访问目标URL需要用户名
''http.Password = ""	'如果访问目标URL需要密码
'http.Open
'YQasp.PrintEnd YQasp.Str.HtmlEncode(http.Html)
'Set http = Nothing
'''=========================

'=========================
'Demo 4 - 获取文件头：
'YQasp.Http.SetHeader "Referer:http://www.baidu.com"
'YQasp.Http.Get "http://www.yqasp.cn"
'tmp = YQasp.Http.Headers
'YQasp.Println YQasp.Str.HtmlEncode(tmp)
'=========================

''=========================
''Demo 5 - 获取文件指定部分内容：
'Dim bookid,bookname,bookdesc,uptime,readlink
'bookid = 1639199
'YQasp.Http.Get("http://www.qidian.com/Book/"&bookid&".aspx")
''用SubStr按字符截取部分文本
'bookname = YQasp.Http.Cut("<div class=""title"">"&vbCrLf&" <h1>","</h1>",0)
'bookdesc = YQasp.Http.Cut("</div>"&vbCrLf&" <div class=""txt"">","</div>",0)
''用Find可按正则获取一段文本
'uptime = YQasp.Http.Find("更新时间：[\d- :]+")
''用Select可按正则编组选择匹配的部分文本,$0是获取正则匹配的字符串本身
'readlink = YQasp.Http.Select("(<a href="")(/BookReader/\d+.aspx)(.+</a>)","$1http://www.qidian.com$2$3")
'YQasp.Println "<b>书名：</b>《" & bookname & "》  " & uptime
'YQasp.Println "<b>阅读地址：</b>" & readlink
'YQasp.Println "<b>内容简介：</b>"
'YQasp.Println bookdesc
''=========================

''=========================
''Demo 6 - 获取文件循环部分：
'YQasp.Http.Get "http://code.google.com/p/yqasp/updates/list"
'rule = "<span class=""date below-more"" title=""(.+?)""[\s\S]+?>(.+?)</span>[\s\S]+?<span class=""title""><a class=""ot-revision-link"" href=""/p/yqasp/source/detail\?r=(?:\d+?)"">(r\d+?)</a>\n \(([\s\S]+?)\).+>(\w+?)</a></span>"
'arr = YQasp.Http.Search(rule)
'YQasp.Println "====前5个匹配===="
'For i = 0 To 4
'	YQasp.Println "<b>第" & i + 1 & "个匹配项：</b>"
'	YQasp.Println YQasp.Str.HtmlEncode(arr(i))
'Next
'YQasp.Println ""
''还可以用正则来进行更复杂的应用
'Dim Matches, Match
'Set Matches = YQasp.Str.Match(YQasp.Http.Html, rule)
'YQasp.Println "====YunqASP更新日志摘要===="
'For Each Match In Matches
'	If Match.SubMatches(3)<>"[No log message]" Then YQasp.Println YQasp.Str.Format("<li>{3}, {4} ({5} @ {2})</li>",Match)
'Next
'Set Matches = Nothing
''=========================

''=========================
''Demo 7 - 保存远程图片：
'YQasp.Http.Get "http://www.baidu.com"
'tmp = YQasp.Http.SaveStringImgTo(YQasp.Http.Html, "imgatlocal/")
'YQasp.Println YQasp.Str.HtmlEncode(tmp)
''=========================

''=========================
''Demo 8 - WebService(SOAP1.1)示例：
''获得腾讯QQ在线状态
'Dim QQ,xml : QQ = 800010000
'tmp = "<?xml version=""1.0"" encoding=""utf-8""?>"
'tmp = tmp & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
'tmp = tmp & "  <soap:Body>"
'tmp = tmp & "    <qqCheckOnline xmlns=""http://WebXml.com.cn/"">"
'tmp = tmp & "      <qqCode>" & QQ & "</qqCode>"
'tmp = tmp & "    </qqCheckOnline>"
'tmp = tmp & "  </soap:Body>"
'tmp = tmp & "</soap:Envelope>"
'Set http = YQasp.Http.New
''设置请求头信息的三种方式，其一：
'http.RequestHeader("Host") = "www.webxml.com.cn"
''其二：
'http.SetHeader "Content-Type:text/xml; charset=utf-8"
''其三：
'http.SetHeader Array("Content-Length:" & Len(tmp), "SOAPAction:http://WebXml.com.cn/qqCheckOnline")
'http.Data = tmp
'tmp = http.Post("http://www.webxml.com.cn/webservices/qqOnlineWebService.asmx?wsdl")
'Set http = Nothing
''解析返回数据
'Set xml = YQasp.Xml.New
'xml.Load tmp
'tmp = xml("qqCheckOnlineResult").Value
'Set xml = Nothing
'Select Case tmp
'	Case "Y" tmp = "在线"
'	Case "N" tmp = "离线"
'	Case "E" tmp = "号码错误"
'	Case "A" tmp = "商业用户验证失败"
'	CAse "V" tmp = "免费用户超过数量"
'End Select
'YQasp.Println "QQ:" & QQ & " (" & tmp & ")"

'=========================

YQasp.Println ""
YQasp.Println "------------------------------------"
YQasp.Print "页面执行时间： " & YQasp.GetScriptTime & " 秒"
%>