(window.webpackJsonp=window.webpackJsonp||[]).push([[62],{480:function(s,t,a){"use strict";a.r(t);var n=a(15),e=Object(n.a)({},(function(){var s=this,t=s.$createElement,a=s._self._c||t;return a("ContentSlotsDistributor",{attrs:{"slot-key":s.$parent.slotKey}},[a("p",[s._v("方法，将复杂的各类集合对象格式化为字符串")]),s._v(" "),a("h3",{attrs:{id:"语法"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#语法"}},[s._v("#")]),s._v(" 🔸 语法")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format stirng"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" obj\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("h3",{attrs:{id:"参数"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#参数"}},[s._v("#")]),s._v(" 🔸 参数")]),s._v(" "),a("table",[a("thead",[a("tr",[a("th",{staticStyle:{"text-align":"left"}},[s._v("参数")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("类型")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("说明")])])]),s._v(" "),a("tbody",[a("tr",[a("td",{staticStyle:{"text-align":"left"}},[s._v("string")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("String (字符串)")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("包含占位符的字符串")])]),s._v(" "),a("tr",[a("td",{staticStyle:{"text-align":"left"}},[s._v("obj")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("Array (数组) 或 Object (ASP对象) 或 Recordset (记录集对象) 或 String (字符串)")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("用于格式化占位符的数据源")])])])]),s._v(" "),a("h3",{attrs:{id:"返回值"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#返回值"}},[s._v("#")]),s._v(" 🔸 返回值")]),s._v(" "),a("table",[a("thead",[a("tr",[a("th",{staticStyle:{"text-align":"left"}},[s._v("类型")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("说明")])])]),s._v(" "),a("tbody",[a("tr",[a("td",{staticStyle:{"text-align":"left"}},[s._v("String (字符串)")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("经过字符替换后的字符串")])])])]),s._v(" "),a("h3",{attrs:{id:"说明"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#说明"}},[s._v("#")]),s._v(" 🔸 说明")]),s._v(" "),a("p",[s._v("调用此方法可以把各种带有集合特征的对象按字符串内占位符的形式把值替换到字符串中，能够很方便的用于各种字符串内变量的拼接。同时可在占位符中对要替换的值进行一些简单的格式化操作。")]),s._v(" "),a("p",[s._v("其中字符串参数为包含占位符的字符串，占位符用"),a("code",[s._v("{}")]),s._v("符号包含，视参数的不同，占位符中间可以是数字或者名称。如果占位符是数字，则编号从0开始表示第一个元素。如果字符串中本身就包含 "),a("code",[s._v("{")]),s._v(" 字符，则需要用 "),a("code",[s._v("\\{")]),s._v(" 进行转义。")]),s._v(" "),a("p",[s._v("用于格式化占位符的数据源，可以是以下类型：\n"),a("code",[s._v("字符串")]),s._v(" - 只替换 {0} 占位符；\n"),a("code",[s._v("数组")]),s._v(" - 依次替换占位符中的 {数字} 为对应的数组元素；\n"),a("code",[s._v("记录集(Recordset)")]),s._v(" - 依次替换占位符中的 {数字} 或 {列名} 为本条记录的相应值；\n"),a("code",[s._v("字典(Dictinary)")]),s._v(" - 依次替换占位符中的 {键名} 为字典中对应的值；\n"),a("code",[s._v("Match集合")]),s._v(" - 替换 {0} 为匹配本身，然后从 {1} 开始替换占位符中的数字为对应的子集合(SubMatches)的值；\n"),a("code",[s._v("SubMatches集合")]),s._v(" - 依次替换占位符中的 {数字} 为本条集合中对应的值；\n"),a("code",[s._v("EasyASP List对象")]),s._v(" - 依次替换占位符中的 {数字} 或 {Hash列名} 为数组中对应的值。")]),s._v(" "),a("h3",{attrs:{id:"示例"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#示例"}},[s._v("#")]),s._v(" 🔸 示例")]),s._v(" "),a("p",[s._v("字符串")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("W Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"This is a {0}."')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Text"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("This "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("is")]),s._v(" a Text"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("数组")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("W Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"name:{0} / sex:{1} / age:{2}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" Array"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Ertan"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Male"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("38")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("name"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(":")]),s._v("Ertan "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" sex"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(":")]),s._v("Male "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" age"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(":")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("38")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("数据集")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Dim")]),s._v(" Rs\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" Rs "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Db"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Sel"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"select * from easp_test"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("While")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Not")]),s._v(" Rs"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("eof\n    Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0} / {1} / {age} / {sex}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" Rs"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\nRs"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("movenext\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Wend")]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Db"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Close Rs\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br"),a("span",{staticClass:"line-number"},[s._v("6")]),a("br"),a("span",{staticClass:"line-number"},[s._v("7")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("22")]),s._v("J3398N1Q "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 张三 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("23")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 男\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("22")]),s._v("J339IZL7 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 李四 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("45")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 男\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("22")]),s._v("J339SSU7 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 王五 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("36")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 女\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("22")]),s._v("J33A25TA "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 赵六 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("54")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 女\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("22")]),s._v("J33ACL2E "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 李七 "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("57")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" 女\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br")])]),a("p",[s._v("Dictionary字典对象")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Dim")]),s._v(" d\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" d "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" CreateObject"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Scripting.Dictionary"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\nd"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Add "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"a"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Athens"')]),s._v("\nd"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Add "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"b"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Belgrade"')]),s._v("\nd"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Add "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"c"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Cairo"')]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{a} / {c}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" d"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" d "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token boolean"}},[s._v("Nothing")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br"),a("span",{staticClass:"line-number"},[s._v("6")]),a("br"),a("span",{staticClass:"line-number"},[s._v("7")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Athens "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" Cairo\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("List数组对象")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Dim")]),s._v(" list\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" list "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("List"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("New")]),s._v("\nlist"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Hash "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"name:Scott pwd:tiger age:39 job:CFO"')]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0} / {pwd} / {2} / {job}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" list"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\nlist"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Data "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"Scott tiger 39 CFO"')]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0} / {1} / {2} / {3}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" list"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" list "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token boolean"}},[s._v("Nothing")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br"),a("span",{staticClass:"line-number"},[s._v("6")]),a("br"),a("span",{staticClass:"line-number"},[s._v("7")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Scott "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" tiger "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("39")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" CFO\nScott "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" tiger "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("39")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" CFO\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br")])]),a("p",[s._v("Match对象（正则搜索子集合）")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Dim")]),s._v(" s"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v("m"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v("match\ns "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"<em>Scott/Tiger</em><em>Smith/Cat</em>"')]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" m "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Match"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),s._v("s"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"<em>(\\w+)/(\\w+)</em>"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("For")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Each")]),s._v(" match "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("In")]),s._v(" m\n    Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0} | {1}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" match"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("SubMatches"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Next")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("For")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Each")]),s._v(" match "),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("In")]),s._v(" m\n    Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0} > {1} | {2}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" match"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Next")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token keyword"}},[s._v("Set")]),s._v(" m "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token boolean"}},[s._v("Nothing")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br"),a("span",{staticClass:"line-number"},[s._v("6")]),a("br"),a("span",{staticClass:"line-number"},[s._v("7")]),a("br"),a("span",{staticClass:"line-number"},[s._v("8")]),a("br"),a("span",{staticClass:"line-number"},[s._v("9")]),a("br"),a("span",{staticClass:"line-number"},[s._v("10")]),a("br")])]),a("p",[s._v("浏览器返回（HTML源代码）：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Scott | Tiger"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("br "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\nSmith | Cat"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("br "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("em"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("Scott"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v("Tiger"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v("em"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v(" Scott | Tiger"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("br "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("em"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("Smith"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v("Cat"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v("em"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v(" Smith | Cat"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("br "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br")])]),a("p",[a("strong",[s._v("占位符参数格式化")]),s._v(" "),a("strong",[s._v("参数N")]),s._v("\n表示格式化数字，语法为 "),a("code",[s._v("N[,或%或(][数字]")]),s._v("，后面两项均为可选:\n"),a("code",[s._v(",")]),s._v(" 用千位分隔符\n"),a("code",[s._v("%")]),s._v(" 转换为百分比\n"),a("code",[s._v("(")]),s._v(" 负数用括号包含\n"),a("code",[s._v("数字")]),s._v(" 小数位数")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{0:N} / {0:N3} / {0:N%2} / {1:N,} / {2:N(1}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" a"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("0.3453")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("0.345")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("34.53")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("2")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("372")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("291.88")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("34.3")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[a("strong",[s._v("参数D")]),s._v("\n表示格式化日期，用法和 "),a("code",[s._v("Easp.Date.Format")]),s._v(" 一样")]),s._v(" "),a("p",[a("strong",[s._v("参数U")]),s._v("\n表示转化为大写字母")]),s._v(" "),a("p",[a("strong",[s._v("参数L")]),s._v("\n表示转化为小写字母")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{3:Dy-mm-dd} / {4:U} / {4:L}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" a"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("2016")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("-")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("05")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("-")]),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("23")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v(" I"),a("span",{pre:!0,attrs:{class:"token comment"}},[s._v("'M COLDSTONE / i'm coldstone")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[a("strong",[s._v("参数E")]),s._v("\n表示使用ASP语句表达式，用"),a("code",[s._v("%s")]),s._v("代替值在语句中的位置，"),a("code",[s._v("Easp")]),s._v("的方法可省略 "),a("code",[s._v("Easp.")]),s._v(" 的前缀")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{4:EReplace(%s,"" "","" / "")} / {3:EReplace(%s,"" "","" - "")}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v("　a"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("I"),a("span",{pre:!0,attrs:{class:"token comment"}},[s._v("'m / Coldstone / 2016/5/23 - 9:57:43")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[a("strong",[s._v("参数为数字")]),s._v("\n如果直接在冒号后跟数字，或者跟数字+任意字符，则是同　"),a("code",[s._v("Easp.Str.Cut")]),s._v("　方法一样可截取字符串")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Str"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Format"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"{4:6} / {4:6…}"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" a"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("I"),a("span",{pre:!0,attrs:{class:"token comment"}},[s._v("'m Coldston / I'm Coldst…")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("h3",{attrs:{id:"应用场景-使用心得"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#应用场景-使用心得"}},[s._v("#")]),s._v(" 🔸 应用场景 & 使用心得")]),s._v(" "),a("p",[s._v("非常给力的一个方法，前提是你已经领会如何使用。\n当你处理你的数据，想从一种格式变成另一种格式的时候，希望你第一个想到这个方法，无论是单独一个数据，还是一个循环的列表里的数据，"),a("code",[s._v("Easp.Str.Format")]),s._v(" 都能非常方便的进行格式变换。")])])}),[],!1,null,null,null);t.default=e.exports}}]);