(window.webpackJsonp=window.webpackJsonp||[]).push([[54],{471:function(s,t,a){"use strict";a.r(t);var e=a(15),n=Object(e.a)({},(function(){var s=this,t=s.$createElement,a=s._self._c||t;return a("ContentSlotsDistributor",{attrs:{"slot-key":s.$parent.slotKey}},[a("p",[s._v("方法，得到服务器端包含文件的结果")]),s._v(" "),a("h3",{attrs:{id:"语法"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#语法"}},[s._v("#")]),s._v(" 🔸 语法")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("GetInclude filePath\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("h3",{attrs:{id:"参数"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#参数"}},[s._v("#")]),s._v(" 🔸 参数")]),s._v(" "),a("table",[a("thead",[a("tr",[a("th",{staticStyle:{"text-align":"left"}},[s._v("参数")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("类型")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("说明")])])]),s._v(" "),a("tbody",[a("tr",[a("td",{staticStyle:{"text-align":"left"}},[s._v("filePath")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("String(字符串)")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("要包含的文件路径，可以是相对路径、以 "),a("code",[s._v("/")]),s._v(" 开头的站点绝对路径和以 "),a("code",[s._v("盘符:\\")]),s._v(" 开头的硬盘绝对路径")])])])]),s._v(" "),a("h3",{attrs:{id:"返回值"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#返回值"}},[s._v("#")]),s._v(" 🔸 返回值")]),s._v(" "),a("table",[a("thead",[a("tr",[a("th",{staticStyle:{"text-align":"left"}},[s._v("类型")]),s._v(" "),a("th",{staticStyle:{"text-align":"left"}},[s._v("说明")])])]),s._v(" "),a("tbody",[a("tr",[a("td",{staticStyle:{"text-align":"left"}},[s._v("String(字符串)")]),s._v(" "),a("td",{staticStyle:{"text-align":"left"}},[s._v("服务器端包含文件运行的结果")])])])]),s._v(" "),a("h3",{attrs:{id:"说明"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#说明"}},[s._v("#")]),s._v(" 🔸 说明")]),s._v(" "),a("p",[s._v("此方法和 "),a("code",[s._v("Easp.Include")]),s._v(" 很相似，不同的地方是 "),a("code",[s._v("Easp.Include")]),s._v(" 方法如果包含有HTML内容时会直接输出，而此方法会将包含文件输出的所有HTML内容返回为一个字符串变量。")]),s._v(" "),a("p",[s._v("其实看源码就能看出来，"),a("code",[s._v("Easp.GetInclude")]),s._v(" 和 "),a("code",[s._v("Easp.Include")]),s._v(" 唯一的区别就是 "),a("code",[s._v("Easp.GetInclude")]),s._v(" 有返回值了。可以赋值给变量进行处理了。")]),s._v(" "),a("h3",{attrs:{id:"示例"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#示例"}},[s._v("#")]),s._v(" 🔸 示例")]),s._v(" "),a("p",[s._v("文件 "),a("code",[s._v("1.asp")]),s._v(" 内容：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"1.asp"')]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("b"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("加粗"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("/")]),s._v("b"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),s._v("br"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("Include "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"/3.html"')]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br"),a("span",{staticClass:"line-number"},[s._v("4")]),a("br"),a("span",{staticClass:"line-number"},[s._v("5")]),a("br"),a("span",{staticClass:"line-number"},[s._v("6")]),a("br"),a("span",{staticClass:"line-number"},[s._v("7")]),a("br")])]),a("p",[s._v("文件 "),a("code",[s._v("3.html")]),s._v(" 内容，注意，是html哟：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("<")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("WN "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"3.html"')]),s._v("\n"),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("%")]),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v(">")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br")])]),a("p",[s._v("测试页 "),a("code",[s._v("2.asp")]),s._v(" 内容：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("GetInclude "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"/1.asp"')]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("1")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("asp\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("3")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("html\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br")])]),a("p",[s._v("加粗的两个字没有了，也就是Html代码没有了，但ASP代码还是被执行了。")]),s._v(" "),a("p",[a("code",[s._v("2.asp")]),s._v(" 里还可以写成：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("W Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("GetInclude"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"/1.asp"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("1")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("asp\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("3")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("html\n加粗\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br")])]),a("p",[s._v("这回加粗两个字又出现了，说明 "),a("code",[s._v("Easp.GetInclude")]),s._v(" 只是一个值，不负责输出内容。")]),s._v(" "),a("p",[s._v("再改一下 "),a("code",[s._v("2.asp")]),s._v(" 里的代码：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[s._v("a "),a("span",{pre:!0,attrs:{class:"token operator"}},[s._v("=")]),s._v(" Easp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("GetInclude"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"/1.asp"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\nEasp"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("W Replace"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v("(")]),s._v("a"),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"加粗"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(",")]),s._v(" "),a("span",{pre:!0,attrs:{class:"token string"}},[s._v('"加得非常粗"')]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(")")]),s._v("\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br")])]),a("p",[s._v("浏览器返回：")]),s._v(" "),a("div",{staticClass:"language-vb line-numbers-mode"},[a("pre",{pre:!0,attrs:{class:"language-vb"}},[a("code",[a("span",{pre:!0,attrs:{class:"token number"}},[s._v("1")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("asp\n"),a("span",{pre:!0,attrs:{class:"token number"}},[s._v("3")]),a("span",{pre:!0,attrs:{class:"token punctuation"}},[s._v(".")]),s._v("html\n加得非常粗\n")])]),s._v(" "),a("div",{staticClass:"line-numbers-wrapper"},[a("span",{staticClass:"line-number"},[s._v("1")]),a("br"),a("span",{staticClass:"line-number"},[s._v("2")]),a("br"),a("span",{staticClass:"line-number"},[s._v("3")]),a("br")])]),a("p",[s._v("这回更清楚地理解这个方法的作用了吧。")]),s._v(" "),a("h3",{attrs:{id:"应用场景-使用心得"}},[a("a",{staticClass:"header-anchor",attrs:{href:"#应用场景-使用心得"}},[s._v("#")]),s._v(" 🔸 应用场景 & 使用心得")])])}),[],!1,null,null,null);t.default=n.exports}}]);