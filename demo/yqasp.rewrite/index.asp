<!--#include file="../../code/yqasp.asp" --><%

'转到带地址栏的链接
YQasp.Print "<a href=""?type=YQasp.coldstone&id=1983-09-23&page=3&lang=%E4%B8%AD%E6%96%87"">set address querystring</a>"
YQasp.Print "&nbsp;&nbsp;&nbsp;"
YQasp.Print "<a href=""./index.asp?photo-203HTKJI9B-6.html"">set address rewrite</a>"
YQasp.Print "&nbsp;&nbsp;&nbsp;"
YQasp.Println "<a href=""./?photo-203HTKJI9B-6.html"">set address rewrite without 'index.asp'</a>"

YQasp.Println "YQasp.DefaultPageName : " & YQasp.DefaultPageName
YQasp.Println "[All] YQasp.GetUrl("""") : " & YQasp.GetUrl("")
YQasp.Println "[Url] YQasp.GetUrl(1) : " & YQasp.GetUrl(1)
YQasp.Println "[Url] YQasp.GetUrl(0) : " & YQasp.GetUrl(0)
YQasp.Println "[Host] YQasp.GetUrl(-1) : " & YQasp.GetUrl(-1)
YQasp.Println "[Dir] YQasp.GetUrl(-2) : " & YQasp.GetUrl(-2)
YQasp.Println "[File] YQasp.GetUrl(-3) : " & YQasp.GetUrl(-3)
YQasp.Println "[White] YQasp.GetUrl(""type,id"") : " & YQasp.GetUrl("type,id")
YQasp.Println "[Black] YQasp.GetUrl(""-type,-id"") : " & YQasp.GetUrl("-type,-id")
YQasp.Println "[Remove all param] YQasp.GetUrl(""-:all"") : " & YQasp.GetUrl("-:all")
YQasp.Println "[New param] YQasp.GetUrlWith(""-page,-lang"", ""page=4&lang=english"") : " & YQasp.GetUrlWith("-page,-lang", "page=4&lang=english")
YQasp.Println "[New page & param] YQasp.GetUrlWith(""./newpage.asp?-type,-id,-lang"", ""lang=english"") : " & YQasp.GetUrlWith("./newpage.asp?-type,-id,-lang", "lang=english")
YQasp.Println ""
'设置伪静态规则
'YQasp.RewriteRule "/testcase/rewrite/\?(\w+)-(\w+)-(\d+).html", "/testcase/rewrite/?type=$1&id=$2&page=$3"
'另一种方式设置伪静态规则
'YQasp.Rewrite "/testcase/rewrite/index.asp", "(\w+)-(\w+)-(\d+).html", "type=$1&id=$2&page=$3"
'设置本页面伪静态规则
YQasp.Println "YQasp.Rewrite """", ""(\w+)-(\w+)-(\d+).html"", ""type=$1&id=$2&page=$3"""
YQasp.Rewrite "", "(\w+)-(\w+)-(\d+).html", "type=$1&id=$2&page=$3"

YQasp.Println "当前页是否符合伪静态规则：" & YQasp.IsRewrite()

YQasp.Println "输出参数值："
YQasp.Println "YQasp.Get(""type"") : " & YQasp.Get("type")
YQasp.Println "YQasp.Get(""id"") : " & YQasp.Get("id")
YQasp.Println "YQasp.Get(""page"") : " & YQasp.Get("page")
YQasp.Println "替换URL参数值："
YQasp.Println "YQasp.ReplaceUrl(""page"", 2) : " & YQasp.ReplaceUrl("page", 2)
YQasp.Println "YQasp.ReplaceUrl(""class"", 2) : " & YQasp.ReplaceUrl("class", 2)

YQasp.Println "YQasp.Str.ToString(YQasp.Var.GetObject) : " & YQasp.Str.ToString(YQasp.Var.GetObject)

%>