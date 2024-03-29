<%
'#########################################################################
'## YQasp.json.asp
'## ----------------------------------------------------------------------
'## Feature     :   YQAsp Json Class
'## Version     :   1.0
'## Author      :   云奇(114066164@qq.com)
'## Update Date :   2021-7-15
'## Description :   Create a json string or Parse a json object/array.
'##                 Based on VBJSON by Michael Glaser (vbjson@ediy.co.nz).
'#########################################################################

Class YQAsp_Json

  Private b_encode, b_quickMode

  Private Sub Class_Initialize()
    YQasp.Error("error-json-invalid-json") = YQasp.Lang("error-json-invalid-json")
    YQasp.Error("error-json-missing-brace") = YQasp.Lang("error-json-missing-brace")
    YQasp.Error("error-json-missing-bracket") = YQasp.Lang("error-json-missing-bracket")
    YQasp.Error("error-json-wrong-key") = YQasp.Lang("error-json-wrong-key")
    YQasp.Error("error-json-wrong-array") = YQasp.Lang("error-json-wrong-array")
    YQasp.Error("error-json-invalid-boolean") = YQasp.Lang("error-json-invalid-boolean")
    YQasp.Error("error-json-invalid-null") = YQasp.Lang("error-json-invalid-null")
    YQasp.Error("error-json-invalid-key") = YQasp.Lang("error-json-invalid-key")
    YQasp.Error("error-json-create-json") = YQasp.Lang("error-json-create-json")
    b_encode = True
    b_quickMode = True
  End Sub
  
  '设置和读取生成Json字符串是是否编码 Unicode 字符
  Public Property Get EncodeUnicode
    EncodeUnicode = b_encode
  End Property
  Public Property Let EncodeUnicode(ByRef bool)
    b_encode = bool
  End Property
  '设置和读取操作Json时是否可以使用快速模式
  '快速模式即使用 Json("aaa.bbb[2].ccc") 的方式
  Public Property Get QuickMode
    QuickMode = b_quickMode
  End Property
  Public Property Let QuickMode(ByRef bool)
    b_quickMode = bool
  End Property
  '新建一个Object对象
  Public Function NewObject()
    Set NewObject = New YQAsp_Json_Object
  End Function
  '新建一个Array对象
  Public Function NewArray()
    Set NewArray = New YQAsp_Json_Array
  End Function
  
  '解析Json字符串并返回 YQaspJson 对象
  Public Function Parse(ByRef str)
    Dim index
    index = 1
    On Error Resume Next
    Call skipChar(str, index)
    Select Case Mid(str, index, 1)
      Case "{"
        Set Parse = ParseObject(str, index)
      Case "["
        Set Parse = ParseArray(str, index)
      Case Else
        YQasp.Error.FunctionName = "YQasp.Json.Parse"
        YQasp.Error.Raise "error-json-invalid-json"
    End Select
  End Function

  '解析 key/value 键值对
  Private Function ParseObject(ByRef str, ByRef index)
    'Set ParseObject = Server.CreateObject("Scripting.Dictionary")
    Set ParseObject = New YQAsp_Json_Object
    Dim sKey
    ' "{"
    Call skipChar(str, index)
    index = index + 1
    Do
      Call skipChar(str, index)
      If "}" = Mid(str, index, 1) Then
        index = index + 1
        Exit Do
      ElseIf "," = Mid(str, index, 1) Then
        index = index + 1
        Call skipChar(str, index)
      ElseIf index > Len(str) Then
        YQasp.Error.FunctionName = "Json.ParseObject"
        YQasp.Error.Detail = Right(str, 20)
        YQasp.Error.Raise "error-json-missing-brace"
        Exit Do
      End If
      '把键值对存入 Dictinary 对象
      sKey = ParseKey(str, index)
      On Error Resume Next
      ParseObject.Put sKey, ParseValue(str, index)
      If Err.Number <> 0 Then
        YQasp.Error.FunctionName = "Json.ParseObject"
        YQasp.Error.Detail = sKey
        YQasp.Error.Raise "error-json-wrong-key"
        Exit Do
      End If
    Loop
  End Function

  '解析数组
  Private Function ParseArray(ByRef str, ByRef index)
    Dim o_dic, i_tmp, s_tmp
    'Set o_dic = Server.CreateObject("Scripting.Dictionary")
    Set ParseArray = New YQAsp_Json_Array
    ' "["
    Call skipChar(str, index)
    index = index + 1
    Do
      Call skipChar(str, index)
      If "]" = Mid(str, index, 1) Then
        index = index + 1
        Exit Do
      ElseIf "," = Mid(str, index, 1) Then
        index = index + 1
        Call skipChar(str, index)
      ElseIf index > Len(str) Then
        YQasp.Error.FunctionName = "Json.ParseArray"
        YQasp.Error.Detail = Right(str, 20)
        YQasp.Error.Raise "error-json-missing-bracket"
        Exit Do
      End If
      '把值加入到数组中
      On Error Resume Next
      'o_dic.Add o_dic.Count, ParseValue(str, index)
      ParseArray.Add ParseValue(str, index)
      If Err.Number <> 0 Then
        YQasp.Error.FunctionName = "Json.ParseArray"
        YQasp.Error.Detail = Mid(str, index, 20)
        YQasp.Error.Raise "error-json-wrong-array"
        Exit Do
      End If
    Loop
    '取得数组
    'ParseArray = o_dic.Items
    'Set o_dic = Nothing
  End Function

  '解析json值：string / number / object / array / true / false / null
  Private Function ParseValue(ByRef str, ByRef index)
    Call skipChar(str, index)
    Select Case Mid(str, index, 1)
      Case "{"
        Set ParseValue = ParseObject(str, index)
      Case "["
        Set ParseValue = ParseArray(str, index)
      Case """", "'"
        ParseValue = ParseString(str, index)
      Case "t", "f"
        ParseValue = ParseBoolean(str, index)
      Case "n"
        ParseValue = ParseNull(str, index)
      Case Else
        ParseValue = ParseNumber(str, index)
    End Select
  End Function

  '解析字符串
  Private Function ParseString(ByRef str, ByRef index)
    Dim quote, Char, Code, SB
    Set SB = YQasp.Str.StringBuilder()
    Call skipChar(str, index)
    quote = Mid(str, index, 1)
    index = index + 1
    Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
        Case "\"
          index = index + 1
            Char = Mid(str, index, 1)
            Select Case (Char)
              Case """", "\", "/", "'"
                SB.Append Char
                index = index + 1
              Case "b"
                SB.Append vbBack
                index = index + 1
              Case "f"
                SB.Append vbFormFeed
                index = index + 1
              Case "n"
                SB.Append vbLf
                index = index + 1
              Case "r"
                SB.Append vbCr
                index = index + 1
              Case "t"
                SB.Append vbTab
                index = index + 1
              Case "u"
                index = index + 1
                Code = Mid(str, index, 4)
                SB.Append ChrW(Eval("&h" + Code))
                index = index + 4
            End Select
         Case quote
            index = index + 1
            ParseString = SB.toString
            Set SB = Nothing
            Exit Function
         Case Else
            SB.Append Char
            index = index + 1
      End Select
    Loop
    ParseString = SB.toString
    Set SB = Nothing
  End Function

  '解析数字
  Private Function ParseNumber(ByRef str, ByRef index)
    Dim Value, Char
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      If InStr("+-0123456789.eE", Char) Then
        Value = Value & Char
        index = index + 1
      Else
        ParseNumber = YQasp.Str.ToNumber(Value,0)
        Exit Function
      End If
    Loop
  End Function

  '解析 true / false
  Private Function ParseBoolean(ByRef str, ByRef index)
     Call skipChar(str, index)
     If Mid(str, index, 4) = "true" Then
        ParseBoolean = True
        index = index + 4
     ElseIf Mid(str, index, 5) = "false" Then
        ParseBoolean = False
        index = index + 5
     Else
        YQasp.Error.FunctionName = "Json.ParseBoolean"
        YQasp.Error.Detail = Array(index, Mid(str, index))
        YQasp.Error.Raise "error-json-invalid-boolean"
     End If
  End Function

  '解析 null
  Private Function ParseNull(ByRef str, ByRef index)
     Call skipChar(str, index)
     If Mid(str, index, 4) = "null" Then
        ParseNull = Null
        index = index + 4
     Else
        YQasp.Error.FunctionName = "Json.ParseNull"
        YQasp.Error.Detail = Array(index, Mid(str, index))
        YQasp.Error.Raise "error-json-invalid-null"
     End If
  End Function
  '解析键值
  Private Function ParseKey(ByRef str, ByRef index)
    Dim dquote, squote, Char
    Call skipChar(str, index)
    Do While index > 0 And index <= Len(str)
      Char = Mid(str, index, 1)
      Select Case (Char)
        Case """"
          dquote = Not dquote
          index = index + 1
          If Not dquote Then
            Call skipChar(str, index)
            If Mid(str, index, 1) <> ":" Then
              YQasp.Error.FunctionName = "Json.ParseKey"
              YQasp.Error.Detail = Array(index, ParseKey)
              YQasp.Error.Raise "error-json-invalid-key"
              Exit Do
            End If
          End If
        Case "'"
          squote = Not squote
          index = index + 1
          If Not squote Then
            Call skipChar(str, index)
            If Mid(str, index, 1) <> ":" Then
              YQasp.Error.FunctionName = "Json.ParseKey"
              YQasp.Error.Detail = Array(index, ParseKey)
              YQasp.Error.Raise "error-json-invalid-key"
              Exit Do
            End If
          End If
        Case ":"
          index = index + 1
          If Not dquote And Not squote Then
            Exit Do
          Else
            ParseKey = ParseKey & Char
          End If
        Case Else
          If InStr(vbCrLf & vbCr & vbLf & vbTab & " ", Char) Then
          Else
            ParseKey = ParseKey & Char
          End If
          index = index + 1
      End Select
    Loop
  End Function

  '过滤特殊字符
  Private Sub skipChar(ByRef str, ByRef index)
    Dim bComment, bStartComment, bLongComment
    Do While index > 0 And index <= Len(str)
      Select Case Mid(str, index, 1)
      Case vbCr, vbLf
        If Not bLongComment Then
          bStartComment = False
          bComment = False
        End If
      Case vbTab, " ", "(", ")"
      Case "/"
        If Not bLongComment Then
          If bStartComment Then
            bStartComment = False
            bComment = True
          Else
            bStartComment = True
            bComment = False
            bLongComment = False
          End If
        Else
          If bStartComment Then
            bLongComment = False
            bStartComment = False
            bComment = False
          End If
        End If
      Case "*"
        If bStartComment Then
          bStartComment = False
          bComment = True
          bLongComment = True
        Else
          bStartComment = True
        End If
      Case Else
         If Not bComment Then
            Exit Do
         End If
      End Select
      index = index + 1
    Loop
  End Sub

  '把对象输出为Json字符串
  Public Function toString(ByRef obj)
    Dim b_encodeJson
    b_encodeJson = YQasp.Str.EncodeJsonUnicode
    YQasp.Str.EncodeJsonUnicode = b_encode
    ToString = YQasp.Str.ToString(obj)
    YQasp.Str.EncodeJsonUnicode = b_encodeJson
  End Function

  Public Function ToEvalKey(ByVal key)
    key = Replace(key, """", "")
    key = YQasp.IIF(YQasp.Str.StartsWith(key, "["), "(" & Mid(key,2), "(""" & key)
    key = key & YQasp.IIF(YQasp.Str.EndsWith(key, "]"), ")", """)")
    key = Replace(key, "][", ")(")
    key = Replace(key, "[", """)(")
    key = Replace(key, "].", ")(""")
    key = Replace(key, ".", """)(""")
    key = Replace(key, "]", "")
    ToEvalKey = key
  End Function

End Class

'JsonObject构建类
Class YQAsp_Json_Object
  Private o_dic
  Private Sub Class_Initialize()
    Set o_dic = Server.CreateObject("Scripting.Dictionary")
    o_dic.CompareMode = 1
  End Sub
  Private Sub Class_Terminate()
    Set o_dic = Nothing
  End Sub
  '设置或读取key/value值
  Public Default Property Get [Get](ByVal key)
    If YQasp.Json.QuickMode And (Instr(key, ".") Or Instr(key, "[")) Then
      On Error Resume Next
      Dim evalKey
      evalKey = "Me.Get" & YQasp.Json.ToEvalKey(key)
      If IsObject(Eval(evalKey)) Then
        Execute "Set [Get] = " & evalKey
      Else
        Execute "[Get] = " & evalKey
      End If
      If Err.Number<>0 Then
        If YQasp.Debug Then
          YQasp.Error.FunctionName = "JsonObject.Get"
          YQasp.Error.Detail = key
          YQasp.Error.Raise "error-json-wrong-key"
        End If
      End If
      Exit Property
    End If
    If IsObject(o_dic(key)) Then
      Set [Get] = o_dic(key)
    Else
      [Get] = o_dic(key)
    End If
  End Property
  Public Property Let [Get](ByVal key, ByRef value)
    Put key, value
  End Property
  '取对象的长度
  Public Property Get Count
    Count = o_dic.Count
  End Property
  '取得Dictionary对象
  Public Property Get GetDictionary
    Set GetDictionary = o_dic
  End Property
  '以对象初始化
  Public Property Let Source(ByVal obj)
    If TypeName(obj) = "Dictionary" Then
      Set o_dic = obj
    ElseIf TypeName(obj) = "YQAsp_Json_Object" Then
      Set o_dic = obj.GetDictionary
    Else
      Set o_dic = Nothing
    End If
  End Property
  '设置key/value值
  '参数： @key   - 可以是本对象下的键名，也可以是本对象下的对象字符串，如：
  '               "key" 或者 "key.key1[0].key2"
  '      @value - 要设置的键值 
  Public Sub Put(ByVal key, ByRef value)
    On Error Resume Next
    '如果是字符串方式
    If YQasp.Json.QuickMode And (Instr(key, ".") Or Instr(key, "[")) Then
      Execute "Me.Get" & YQasp.Json.ToEvalKey(key) & " = value"
      Exit Sub
    Else
      YQasp.SetDictionaryKey o_dic, key, value
    End If
    If Err.Number<>0 Then
      If YQasp.Debug Then
        YQasp.Error.FunctionName = "JsonObject.Put"
        YQasp.Error.Detail = Array(key, "(" & TypeName(value) & ") " & YQasp.Str.ToString(value))
        YQasp.Error.Raise "error-json-create-json"
      End If
    End If
  End Sub
  '检测键值是否存在
  Public Function Exists(ByVal key)
    Exists = o_dic.Exists(key)
  End Function
  '检测键值是否存在有效值
  Public Function Has(ByVal key)
    Has  = YQasp.Has(o_dic(key))
  End Function
  '移除某一元素
  Public Sub Remove(ByVal key)
    If o_dic.Exists(key) Then
      If IsObject(o_dic(key)) Then Set o_dic(key) = Nothing
      o_dic.Remove key
    End If
  End Sub
  '全部清空
  Public Sub Clear()
    o_dic.RemoveAll()
    Set o_dic = Nothing
    Set o_dic = Server.CreateObject("Scripting.Dictionary")
    o_dic.CompareMode = 1
  End Sub
  '把Json Object对象输出为字符串
  Public Function ToString()
    ToString = YQasp.Json.ToString(o_dic)
  End Function
  '复制为新对象
  Public Function Clone()
    Dim o_dicN, key
    Set o_dicN = Server.CreateObject("Scripting.Dictionary")
    o_dicN.CompareMode = 1
    Set Clone = New YQAsp_Json_Object
    For Each key In o_dic
      If IsObject(o_dic(key)) Then
        Set o_dicN(key) = o_dic(key)
      Else
        o_dicN(key) = o_dic(key)
      End If
    Next
    Clone.Source = o_dicN
  End Function
End Class
'JsonArray构建类
Class YQAsp_Json_Array
  Private o_dic
  Private Sub Class_Initialize()
    Set o_dic = Server.CreateObject("Scripting.Dictionary")
  End Sub
  Private Sub Class_Terminate()
    Set o_dic = Nothing
  End Sub
  '读取或设置数组元素的值
  Public Default Property Get [Get](ByVal index)
    If IsObject(o_dic(index)) Then
      Set [Get] = o_dic(index)
    Else
      [Get] = o_dic(index)
    End If
  End Property
  Public Property Let [Get](ByVal index, ByRef value)
    Dim i
    If index > 0 Then
      For i = 0 To index - 1
        If Not o_dic.Exists(i) Then o_dic.Add i, Null
      Next
    End If
    YQasp.SetDictionaryKey o_dic, index, value
  End Property
  '取数组的长度
  Public Property Get Length
    Length = o_dic.Count
  End Property
  '取得数组对象
  Public Property Get GetArray
    GetArray = o_dic.Items
  End Property
  '添加一个值
  Public Sub Add(ByRef value)
    o_dic.Add o_dic.Count, value
  End Sub
  '全部清空
  Public Sub Clear()
    o_dic.RemoveAll()
    Set o_dic = Nothing
    Set o_dic = Server.CreateObject("Scripting.Dictionary")
  End Sub
  '将数组值赋给JsonArray对象
  Public Sub SetArray(ByVal arr)
    If IsArray(arr) Then
      Dim i
      Clear()
      For i = 0 To Ubound(arr)
        Add arr(i)
      Next
    End If
  End Sub
  '移除某一元素
  Public Sub Remove(ByVal index)
    If (index = (o_dic.Count-1)) Then
      o_dic.Remove(index)
    ElseIf o_dic.Exists(index) Then
      YQasp.SetDictionaryKey o_dic, index, Null
    End If
  End Sub
  
  '将Json Array对象输出为字符串
  Public Function ToString()
    ToString = YQasp.Json.ToString(o_dic.Items)
  End Function
End Class
%>