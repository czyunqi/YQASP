<!--#include file="../../code/yqasp.asp" --><%
Dim str,n,i
str = 			"<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
str = str & "<microblog>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Tencent"">腾讯微博</name>" & vbCrLf
str = str & "		<url>http://t.qq.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""me""><name>@lengshi</name><nick>Ray</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[今天我们这里下<em>大雨</em>啦！]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Sina"">新浪微博</name>" & vbCrLf
str = str & "		<url>http://t.sina.com.cn</url>" & vbCrLf
str = str & "		<account nick=""email"" for=""me""><name>@tainray</name><nick>tainray@sina.com</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[是不是<font color=""red"">这样</font>的噢，我也不知道哈。<img src=""http://img.t.sinajs.cn/t4/appstyle/expression/ext/normal/af/cry.gif"" />]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Twitter"">推特</name>" & vbCrLf
str = str & "		<url>http://twitter.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""notme""><name haha=""1"">@ccav</name><nick>CCAV</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[I don't need this feature <strong>(>_<)</strong> any more.]]></last></site>" & vbCrLf
str = str & "</microblog>"

'载入Xml数据
'YQasp.Xml.Load "http://yqasp.lengshi.cn/data/xml/microblog_catalog.xml"
YQasp.Xml.Load str
''选择所有标签为name的节点，并输出找到的节点个数
'YQasp.PrintlnHtml YQasp.Xml("name").Length
'YQasp.Println "--------"
''选择所有包含属性alias的标签为name的节点
'YQasp.PrintlnHtml YQasp.Xml("name[alias]").Length
'YQasp.Println "--------"
''选择所有属性for等于me，nick属性不等于email的标签为account的节点，并输出其Xml代码
'YQasp.PrintlnHtml YQasp.Xml("account[for='me'][nick!='email']").Xml
'YQasp.Println "--------"
''选择site节点的子节点中标签为name的节点
'YQasp.PrintlnHtml YQasp.Xml("site>name").Xml
'YQasp.Println "--------"
''选择account节点的后代节点中标签为name的节点
'YQasp.PrintlnHtml YQasp.Xml("account name").Xml
'YQasp.Println "--------"
''选择所有的url和last节点
'YQasp.PrintlnHtml YQasp.Xml("url,last").Xml
YQasp.Println "--------"
YQasp.PrintlnHtml YQasp.Xml("url")(2).Xml
YQasp.Xml("url")(2).Text = "<test>sss</test>"
YQasp.PrintlnHtml YQasp.Xml("url")(2).Xml

'YQasp.Xml.XSLT = "xsl/microblog.xsl"
'YQasp.PrintlnHtml YQasp.Xml.Dom.Xml

'YQasp.Println YQasp.Xml.SaveAs("news.xml>gbk")
'YQasp.Println YQasp.Xml.SaveAs("microblog.xml>utf-8")

'Set n = YQasp.Xml("title")
'For i = 0 To n.Length-1
'	YQasp.Println n(i).Value
'Next
'Set n = Nothing

'YQasp.Println YQasp.Xml("last")(2).Value
Set n = YQasp.Xml("last")
For i = 0 To n.Length-1
	YQasp.Println n(i).Type
	YQasp.Println n(i).Value
Next
'YQasp.Println n.Text
'YQasp.Println n(1).Root.Type
'YQasp.Println n(2).Parent.Name
'YQasp.Println n(0).Clone(1).Text
'Set n = Nothing
'YQasp.Xml("name")(0).RemoveAttr("alias")
'YQasp.PrintlnHtml YQasp.Xml("name")(0).Xml
'YQasp.Xml("site")(1).Clear
'YQasp.PrintlnHtml YQasp.Xml("site")(1).Xml

'YQasp.PrintlnHtml TypeName(YQasp.Xml("site")(0).Parent.Parent.Dom)
'YQasp.Xml("url").Remove
'YQasp.Xml("name").Attr("alias") = Null
'YQasp.Xml("microblog").Remove
'YQasp.Println YQasp.Xml.Sel("//site").Length
'YQasp.Println YQasp.Xml.Select("//site").Length
'YQasp.Println YQasp.Xml("site").Length
'YQasp.Println YQasp.Xml("site").Type
'YQasp.Xml("url")(2).Value = "http://sss.com"
'YQasp.Println TypeName(n)
'替换节点
'Set n = YQasp.Xml("name")(1).ReplaceWith(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("name").ReplaceWith(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("name")(1).ReplaceWith(YQasp.Xml("url")(2))
'YQasp.PrintlnHtml n.Xml
'清空
'YQasp.Xml("url").Empty
'YQasp.Xml("name").Clear
'从前面加入节点
'Set n = YQasp.Xml("account")(1).Before(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("account")(1).Before(YQasp.Xml("url")(2))
'Set n = YQasp.Xml("account").Before(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("account").Before(YQasp.Xml("url")(2))
'从后面加入节点
'Set n = YQasp.Xml("account")(2).After(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("last")(1).After(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("account")(1).After(YQasp.Xml("url")(2))
'Set n = YQasp.Xml("account").After(YQasp.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = YQasp.Xml("account").After(YQasp.Xml("url")(2))


'YQasp.PrintlnHtml n.Xml
'YQasp.PrintlnHtml YQasp.Xml.Dom.Xml

'YQasp.PrintlnHtml YQasp.Xml("name").Length
'YQasp.PrintlnHtml YQasp.Xml("site name").Length
'YQasp.PrintlnHtml YQasp.Xml("site>name").Length
'YQasp.PrintlnHtml YQasp.Xml("name[alias='Tencent'],url").Length
'YQasp.PrintlnHtml YQasp.Xml("name[alias='Tencent'],url").Text
'YQasp.PrintlnHtml YQasp.Xml.Select("//account[@nick='user' and position()<2]").Length
'YQasp.PrintlnHtml YQasp.Xml.Select("//account[@nick='user' and position()<2]").Xml
'YQasp.PrintlnHtml YQasp.Xml("account[nick='user'][for!='me'],account[nick!='user']").Xml

'YQasp.PrintlnHtml YQasp.Xml("site")(1).Find("account").Root.TypeString
'YQasp.PrintlnHtml YQasp.Xml.Root.TypeString

'Set n = Nothing
YQasp.Println ""
YQasp.Println "------------------------------------"
YQasp.Print "页面执行时间： " & YQasp.GetScriptTime & " 秒"
%>