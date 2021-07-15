<!--#include file="../../code/yqasp.asp" --><%

Session("StringTest") = "This is a test string for YQasp.Str.ToString"

YQasp.SetCookie "app_name", "easyAsp_Name", ""
YQasp.SetCookie "site>my_name", "coldstone_yqasp", ""
YQasp.SetCookie "site>mytype", "very&diaosi", ""

YQasp.Println "YQasp.Str.IsSame(""ABCD"", ""abcd"") : " & YQasp.Str.IsSame("ABCD", "abcd")
YQasp.Println "YQasp.Str.IsEqual(""ABCD"", ""abcd"") : " & YQasp.Str.IsEqual("ABCD", "abcd")
YQasp.Println "YQasp.Str.IsEqual(""ABCD"", ""<"", ""abcd"") : " & YQasp.Str.Compare("ABCD", "<", "abcd")
YQasp.Println ""
YQasp.Println "YQasp.Str.GetColonName(""username"") : " & YQasp.Str.GetColonName("username")
YQasp.Println "YQasp.Str.GetColonName(""username:value"") : " & YQasp.Str.GetColonName("username:value")
YQasp.Println "YQasp.Str.GetColonValue(""username"") : " & YQasp.Str.GetColonValue("username")
YQasp.Println "YQasp.Str.GetColonValue(""username:value"") : " & YQasp.Str.GetColonValue("username:value")
YQasp.Println "YQasp.Str.GetColonValue(""username:"") : " & YQasp.Str.GetColonValue("username:")
YQasp.Println ""
YQasp.Println "YQasp.Str.ToString(""testing"") : " & YQasp.Str.ToString("testing")
YQasp.Println "YQasp.Str.ToString(Array(""yes"",""no"",""unknown"")) : " & YQasp.Str.ToString(Array("yes","no","unknown"))
YQasp.Println "YQasp.Str.ToString(Array(12,34,111,98,0)) : " & YQasp.Str.ToString(Array(12,34,111,98,0))
YQasp.Println "YQasp.Str.ToString(Array()) : " & YQasp.Str.ToString(Array())
YQasp.Println "YQasp.Str.ToString(Empty) : " & YQasp.Str.ToString(Empty)
YQasp.Println "YQasp.Str.ToString(Null) : " & YQasp.Str.ToString(Null)
YQasp.Println "YQasp.Str.ToString(Nothing) : " & YQasp.Str.ToString(Nothing)
YQasp.Println "YQasp.Str.ToString(Err) : " & YQasp.Str.ToString(Err)
YQasp.Println ""
Dim dic : Set dic = Server.CreateObject("Scripting.Dictionary")
YQasp.SetDictionaryKey dic, "my-time", Now
YQasp.SetDictionaryKey dic, "my", "Yes,it's me."
YQasp.Println "YQasp.Str.ToString(dic) : " & YQasp.Str.ToString(dic)
YQasp.Println ""
YQasp.Var("YQaspVar") = "Easyasp variable"
YQasp.Var("TestArray") = Array(2323,490,108,"我是中文",992,83,920)
YQasp.Println "YQasp.Str.ToString(YQasp.Var.GetObject) : " & YQasp.Str.ToString(YQasp.Var.GetObject)
'YQasp.Console YQasp.Var.GetObject
YQasp.Println ""
YQasp.Println "YQasp.Str.ToString(12.1256) : " & YQasp.Str.ToString(12.1256)
YQasp.Println "YQasp.Str.ToString(12.100) : " & YQasp.Str.ToString(12.100)
YQasp.Println "YQasp.Str.ToString(12.005) : " & YQasp.Str.ToString(12.005)
YQasp.Println "YQasp.Str.ToString(12.00) : " & YQasp.Str.ToString(12.00)
YQasp.Println ""
YQasp.Println "YQasp.Str.ToString(Session) : " & YQasp.Str.ToString(Session)
YQasp.Println "YQasp.Str.ToString(Request.Cookies) : " & YQasp.Str.ToString(Request.Cookies)
YQasp.Println "YQasp.Str.ToString(Request.QueryString) : " & YQasp.Str.ToString(Request.QueryString)
YQasp.Println "YQasp.Str.ToString(Request.Form) : " & YQasp.Str.ToString(Request.Form)
YQasp.Println ""
YQasp.Println "YQasp.Str.Cut(""This我"",3) : " & YQasp.Str.Cut("This我",3)
YQasp.Println "YQasp.Str.Cut(""我是一个人"",4) : " & YQasp.Str.Cut("我是一个人",4)
YQasp.Println ""
YQasp.Println "YQasp.Str.RepPart(""photo-3.html"", ""^(\w+)-(\d+)\.html$"", ""$2"", ""4"") : " & YQasp.Str.ReplacePart("photo-3.html", "^(\w+)-(\d+)\.html$", "$2", "4")
YQasp.Println ""
YQasp.Println "YQasp.Str.RandomNumber(1000,9999) : " & YQasp.Str.RandomNumber(1000,9999)
YQasp.Println "YQasp.Str.RandomStr(10) : " & YQasp.Str.RandomStr(10)
YQasp.Println "YQasp.Str.RandomStr(""12:0123456789abcdefghijklmnopqrstuvwxyz~!@#$%^&*_-+="") : " & YQasp.Str.RandomStr("12:0123456789abcdefghijklmnopqrstuvwxyz~!@#$%^&*_-+=")
YQasp.Println "YQasp.Str.RandomStr(""10000-99999"") : " & YQasp.Str.RandomStr("10000-99999")
Dim color : color = YQasp.Str.RandomStr("#<3>:0123456789ABCDEF")
YQasp.Println "YQasp.Str.RandomStr(""Random Color \: #<3>:0123456789ABCDEF"") : <span style=""background-color:" & color & """>Random Color : " & color & "</span>"
YQasp.Println "YQasp.Str.RandomStr(""{<8>-<4>-<4>-<4>-<12>}:0123456789ABCDEF"") : " & YQasp.Str.RandomStr("{<8>-<4>-<4>-<4>-<12>}:0123456789ABCDEF")
YQasp.Println "YQasp.Str.RandomStr(""CN-\<86\>-<6>-<10000-99999>"") : " & YQasp.Str.RandomStr("CN-\<86\>-<6>-<10000-99999>")
YQasp.Println ""
YQasp.Println "YQasp.Str.ToNumber(number, decimalType) 方法："
YQasp.Println "如果第二个参数为N，则保留N位小数，小数位数不足的补0"
YQasp.Println "YQasp.Str.ToNumber(0.345678, 3) : " & YQasp.Str.ToNumber(0.345678, 3)
YQasp.Println "YQasp.Str.ToNumber(0.34, 3) : " & YQasp.Str.ToNumber(0.34, 3)
YQasp.Println "YQasp.Str.ToNumber(0, 3) : " & YQasp.Str.ToNumber(0, 3)
YQasp.Println "如果第二个参数为0，则保留所有小数位数"
YQasp.Println "YQasp.Str.ToNumber(0.345678, 0) : " & YQasp.Str.ToNumber(0.345678, 0)
YQasp.Println "YQasp.Str.ToNumber(0.34, 0) : " & YQasp.Str.ToNumber(0.34, 0)
YQasp.Println "YQasp.Str.ToNumber(0, 0) : " & YQasp.Str.ToNumber(0, 0)
YQasp.Println "如果第二个参数为-N，则保留N位小数，但小数位数不足的不补0"
YQasp.Println "YQasp.Str.ToNumber(0.345678, -3) : " & YQasp.Str.ToNumber(0.345678, -3)
YQasp.Println "YQasp.Str.ToNumber(0.34, -3) : " & YQasp.Str.ToNumber(0.34, -3)
YQasp.Println "YQasp.Str.ToNumber(0, -3) : " & YQasp.Str.ToNumber(0, -3)
YQasp.Println ""
YQasp.Println "YQasp.Str.ToPrice(12.3456) : " & YQasp.Str.ToPrice(12.3456)
YQasp.Println "YQasp.Str.ToPercent(0.3456) : " & YQasp.Str.ToPercent(0.3456)
YQasp.Println ""
YQasp.Println "YQasp.Str.Half2Full(""半角To全角"") : " & YQasp.Str.Half2Full("半角To全角")
YQasp.Println "YQasp.Str.Full2Half(""全角Ｔｏ半角"") : " & YQasp.Str.Full2Half("全角Ｔｏ半角")
%>