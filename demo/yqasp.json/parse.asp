<!--#include file="../../code/yqasp.asp" --><%
'测试解析JSON
Dim s_json, obj
s_json = YQasp.Fso.Read("sample.json")
Set obj = YQasp.Decode(s_json)
Dim Devices,i
'如果是数组需要用.GetArray方法取出后方可循环
Devices = obj("Circuit[0].Devices").GetArray
'Devices = Devices
For i = 0 To Ubound(Devices)
  YQasp.Print "Name:"
  YQasp.Println Devices(i)("Name")
  YQasp.Print "dSID:"
  YQasp.Println Devices(i)("dSID")
  YQasp.Print "ZoneID:"
  YQasp.Println Devices(i)("ZoneID")
  YQasp.Println "======="
Next
Set obj = Nothing
YQasp.Println "=============================="
s_json = YQasp.Fso.Read("samplewithcomment.json")
Set obj = YQasp.Decode(s_json)
'有两种方式访问解析后的Json对象或数组
YQasp.Println obj(0)("alert")("message")(1)("set")("name")
YQasp.Println obj(0)("alert.message[2].switch.case.input.title")
YQasp.Println "=============================="
YQasp.PrintlnString YQasp.Encode(obj)
%>