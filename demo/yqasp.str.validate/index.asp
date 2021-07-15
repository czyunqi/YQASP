<!--#include file="../../code/yqasp.asp" --><%

If YQasp.Has(YQasp.Get("action")) Then
  Dim act
  '验证url参数，必须等于 save
  act = YQasp.GetVal("action").Name("action").Same("save").Alert
  YQasp.Var("username") =  YQasp.PostVal("username").Name("用户名").Required.Test("username").Alert
  YQasp.Var("email") =  YQasp.PostVal("email").Name("Email").Test("email").Alert
  '验证两次输入的密码一致
  YQasp.VarVal("post.password").Name("密码").Required.Trim().MinLength(6).SamePost("passwordrepeate").Alert
  YQasp.Var("password") = YQasp("md5")(YQasp.Post("password"))
  '验证序列
  YQasp.Var("type") = YQasp.VarVal("post.type").Name("类型").Split(", ").IsNumber.Join("|").Alert()
  '验证验证码
  Session("verifycode") = "E92A"
  YQasp.Println "Verify Code:" & YQasp.PostVal("verify").SameSession("verifycode").Msg("wrong code").PrintEndJson()
  YQasp.Println "UserName:" & YQasp.Var("username")
  YQasp.Println "Password:" & YQasp.Var("password")
  YQasp.Println "Type:" & YQasp.Var("type")
End If
%>
<form action="?action=save&username=coldstone" method="post">
  username: <input type="text" size="60" name="username" value="" /><br />
  email: <input type="text" size="60" name="email" value="" /><br />
  password: <input type="password" size="60" name="password" value="" /><br />
  repeat: <input type="password" size="60" name="passwordrepeate" value="" /><br />
  verify: <input type="text" size="20" name="verify" value="" />E92A<br />
  <input type="checkbox" name="type" value="1" checked="checked" />type1
  <input type="checkbox" name="type" value="2" />type2
  <input type="checkbox" name="type" value="3" />type3
  <input type="checkbox" name="type" value="4" />type4<br />
  <button type="submit">Submit to "?action=save&username=coldstone"</button>
</form>