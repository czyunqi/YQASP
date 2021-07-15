<!--#include file="../../code/yqasp.asp" --><!--#include file="../../yqasp/plugin/yqasp.hanzi.asp" --><%
YQasp.BasePath = "/source/yqasp"
'测试汉字转拼音插件
Dim Hanzi, cn
cn = "二哥，传说嘉陵江边的重庆人都是重口味，跟《麻辣烫》中一样一样的。"
'Set Hanzi = YQasp.Ext("Hanzi")
Set Hanzi = New YunqASP_Hanzi
YQasp.Println "TEXT : " & cn
YQasp.Println "TitleCase : " & Hanzi.TitleCase
''设置为首字母不大写
'Hanzi.FirstLetterUpcase = False
YQasp.Println "GetPinYin : " & Hanzi.GetPinYin(cn)
YQasp.Println "GetPY : " & Hanzi.GetPY(cn)
YQasp.Println "GetPinYinRead : " & Hanzi.GetPinYinRead(cn)
YQasp.Println "GetPinyin1234 : " & Hanzi.GetPinyin1234(cn)
'GetPinYinWith("中文字符串", 拼音韵母转为字母, 拼音后标识声调, 拼音间加空格, 仅取首字母, 首字母大写)
YQasp.Println "GetPinYinWith : " & Hanzi.GetPinYinWith(cn, True, False, True, False, True)
YQasp.Println "GetEnglish : " & Hanzi.GetEnglish(cn)
YQasp.Println "GetEnglishDash : " & Hanzi.GetEnglishDash(cn)
YQasp.Println "GetKeyWord : " & Hanzi.GetKeyWord(cn)
YQasp.Println "GetKeyWordArray : " & YQasp.Encode(Hanzi.GetKeyWordArray(cn))
YQasp.Println "============================"
YQasp.Println YQasp.GetScriptTime & "s"

%>