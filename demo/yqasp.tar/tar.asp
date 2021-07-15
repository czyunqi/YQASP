<!--#include file="../../code/yqasp.asp" --><%
Dim act : act = YQasp.Var("act")
Select Case act
  Case "pack", "down"
    Dim i
    '在压缩包内创建新文件
    YQasp.Tar.CreateFile YQasp.Var("newfile"), YQasp.Var("newfilecontent")
    '在压缩包内创建新文件夹
    YQasp.Tar.CreateFolder YQasp.Var("newfolder")
    For i = 0 To Ubound(YQasp.Var("files_array"))
      '添加文件夹或文件到压缩包
      If Left(YQasp.Var("files_array_" & i), 1) <> "|" Then
        YQasp.Tar.Add YQasp.Var("files_array_" & i)
      End If
    Next
    If act = "pack" Then
      '打包保存到硬盘
      If YQasp.Tar.PackTo(YQasp.Var("tarfile")) Then
        YQasp.Str.JsAlertUrl "打包成功！文件存放到站点的下面位置：" & vbCrLf & YQasp.Tar.SavePath, "."
      Else
        YQasp.Str.JsAlertUrl "打包失败", "."
      End If
    Else
      '打包后直接输出到浏览器让用户下载
      YQasp.Tar.DownLoad YQasp.Var("tarfile")
      YQasp.RR "."
    End If
  Case "unpack"
    YQasp.Tar.HasSelf = YQasp.Var("hasself") = "1"
    '方法一：
    'YQasp.Tar.SavePath = YQasp.Var("savepath")
    'YQasp.Tar.LoadTar YQasp.Var("untarfile")
    'YQasp.Tar.Unpack()
    '方法二：
    YQasp.Tar.UnPackTo YQasp.Var("untarfile"), YQasp.Var("savepath")
    YQasp.Str.JsAlertUrl "解压成功！文件解压到站点的下面位置：" & vbCrLf & YQasp.Tar.SavePath, "."
  Case "sitetree" '取站点目录树
    Dim arr1, arr2, obj1, fileList, Root
    YQasp.Var("root") = UnEscape(YQasp.Var("root"))
    arr1 = YQasp.Fso.Dir(YQasp.Var("root"))
    Set arr2 = YQasp.Json.NewArray
    For i = 0 To Ubound(arr1,2)
      Set obj1 = YQasp.Json.NewObject
      obj1.Put "name", YQasp.Var("root") & arr1(0,i)
      obj1.Put "type", YQasp.IIF(Right(arr1(0,i), 1)="/", "folder", "file")
      arr2.Add obj1
    Next
    fileList = YQasp.Encode(arr2)
    Set obj1 = Nothing
    Set arr2 = Nothing
    YQasp.PrintEnd fileList
End Select 
%>