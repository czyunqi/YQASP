<!--#include file="../../code/yqasp.asp" --><%
'构建JSON，EasyAsp中有四种可以使用的方式
Dim Json
'======================================================
'第一种方式是所有的赋值操作都在主对象上完成：
'======================================================
'先建立一个NewObject或者NewArray的主对象
Set Json = YQasp.Json.NewObject
Json("Image") = YQasp.Json.NewObject
'可以用下面这种方式直接设置 key/value
Json("Image")("position") = YQasp.Json.NewArray
'或者用Put方法设置 key/value
'Json("Image").Put "position", YQasp.Json.NewArray
'数组可以直接向下标添加value
Json("Image")("position")(0) = YQasp.Json.NewObject
'也可以用下面的方式添加
'Json("Image")("position").Add YQasp.Json.NewObject
Json("Image")("position")(0)("x") = 200
Json("Image")("position")(0)("y") = 131.5
'Json("Image")("position")(1) = YQasp.Json.NewObject
Json("Image")("position").Add YQasp.Json.NewObject
Json("Image")("position")(1)("x") = 240
Json("Image")("position")(1)("y") = -100.5
Json("Image")("position")(4) = Empty
'可以随时查看对象生成的Json：
YQasp.Console Json("Image")("position").ToString
Json("Image").Put "Width", 800
Json("Image").Put "Height", 600
Json("Image")("Title") = "View from 15th Floor"
Json("Image")("Thumbnail") = YQasp.Json.NewObject
Json("Image")("Thumbnail")("Url") = "http://www.example.com/image/481989943"
Json("Image")("Thumbnail")("Width") = 125
Json("Image")("Thumbnail")("Height") = 100
Json("Image")("Thumbnail")("Border") = False
Json("Image")("IDs") = Array(116, 943, ",-1,23,453,", 234, 3879365862)
Json("Text") = "Photo by 冷石"
Json("Alt") = Null
YQasp.Println "第一种方式，直接输出："
'输出Json字符串
YQasp.Println Json.ToString()
'也可以直接将Json对象转为字符串：
YQasp.Println "第一种方式，把Json对象输出为字符串："
'可用YQasp.Str.EncodeJsonUnicode设置是否编码Unicode字符
YQasp.Println YQasp.Str.ToString(Json)
YQasp.Println "第一种方式，不编码Unicode字符："
'设置不编码Unicode字符
'用YQasp.Str.ToString(Json)不受此属性影响，可用YQasp.Str.EncodeJsonUnicode设置
YQasp.Json.EncodeUnicode = False
YQasp.Println Json.ToString()
Set Json = Nothing
YQasp.Println "=========================="
'======================================================
'第二种方式则看上去非常直观：
'======================================================
Set Json = YQasp.Json.NewObject
Json("Image") = YQasp.Json.NewObject
'可以用下面这种方式直接设置 key/value
Json("Image.position") = YQasp.Json.NewArray
'数组可以直接向下标添加value
Json("Image.position[0]") = YQasp.Json.NewObject
'也可以用下面的方式添加
Json("Image.position[0].x") = 200
Json("Image.position[0].y") = 131.5
Json("Image.position[1]") = YQasp.Json.NewObject
Json("Image.position[1].x") = 240
Json("Image.position[1].y") = -100.5
'YQasp.Println Json("Image.position[1].y")
Json("Image.position[4]") = Empty
YQasp.Println Json.ToString()
Set Json = Nothing
YQasp.Println "=========================="
'======================================================
'第三种方式是逐步建立对象，然后再一级级把对象组装起来：
'======================================================
Set Json = YQasp.Json.NewObject
Dim img, pos, xy
Set xy = YQasp.Json.NewObject
Set pos = YQasp.Json.NewArray
Set img = YQasp.Json.NewObject
xy("x") = 200
xy("y") = 131.5
pos(0) = xy
Set xy = YQasp.Json.NewObject
xy("x") = 240
xy("y") = -100.5
pos(1) = xy
pos(4) = Empty
img("position") = pos
Json("Image") = img
YQasp.Println Json.ToString()
Set xy = Nothing
Set pos = Nothing
Set img = Nothing
Set Json = Nothing
YQasp.Println "=========================="
'======================================================
'第四种方式则是用原生的字典对象和数组来构建：
'======================================================
Set Json = Server.CreateObject("Scripting.Dictionary")
Dim img1, pos1(4), xy1
Set xy1 = Server.CreateObject("Scripting.Dictionary")
xy1.Add "x", 200
xy1.Add "y", 131.5
Set pos1(0) = xy1
Set xy1 = Server.CreateObject("Scripting.Dictionary")
xy1.Add "x", 240
xy1.Add "y", -100.5
Set pos1(1) = xy1
pos1(2) = Null
pos1(3) = Null
pos1(4) = Empty
Set img1 = Server.CreateObject("Scripting.Dictionary")
img1.Add "position", pos1
Json.Add "Image", img1
YQasp.Println YQasp.Str.ToString(Json)
Set xy1 = Nothing
Set img1 = Nothing
Set Json = Nothing
'======================================================
'当然，前面的几种方式是可以混用的！你可以在任何时候选择任何一种方式。
'
'最后需要说明一下，在YunqASP中，所有的集合、N维数组、记录集和绝大
'多数ASP对象以及YQasp的List对象都可以直接用 YQasp.Decode(Object)
'方法转为 Json 格式字符串，你可以尝试一下。
'======================================================

%>