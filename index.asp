<!--#include file="Conn.asp" -->
<!--#include file="Kin_Db_Pager.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="Style/css.css" media="all" />
<style>
body {
	background-color:#FFFFFF;
}
</style>
<title>Kin_Db_Pager 分页类 - KinJAVA日志</title>
</head>
<body>
<p>Kin_Db_Pager 分页类(<a href="http://jorkin.reallydo.com" target="_blank">KinJAVA日志</a> <a href="http://jorkin.reallydo.com/article.asp?id=534" target="_blank">下载</a>) <a href="help.asp">使用帮助</a></p>
<%
'//-----------------------------------------------------------------------------
'// 定义Eg()样例相关变量 如果未使用Option Explicit可省略
'//-----------------------------------------------------------------------------
'Dim i, iCols, iColsPercent, iPageSize
'Dim iCurrPage, iRecordCount, iPageCount
'Dim sPageInfo, sPager, sJumpPage
Dim oDbPager
'//-----------------------------------------------------------------------------
'// Kin_Db_Pager分页类开始
'//-----------------------------------------------------------------------------
iPageSize = 15
Set oDbPager = New Kin_db_Pager
'//-----------------------------------------------------------------------------
'// 进行数据库查询前的相关参数设置
'//-----------------------------------------------------------------------------
'//指定数据库连接
oDbPager.Connect(oConn) '//方法一(推荐)
'Set oDbPager.ActiveConnection = oConn '//方法二
'oDbPager.ConnectionString = oConn.ConnectionString '//方法三
'//指定表示页数的URL变量 默认值:"page"
'oDbPager.PageParam = "page"
'//指定数据库类型.默认值:"MSSQL"
oDbPager.DbType = "ACCESS"
'//指定目标表 可用临时表"(Select * From [Table])"
oDbPager.TableName = "Kin_Article"
'//选择列 用逗号分隔 默认值:"*"
oDbPager.Fields = "*"
'//指定该表的主键
oDbPager.PKey = "Article_ID"
'//指定排序条件
oDbPager.OrderBy = "Article_ID DESC"
'//添加条件 可多次使用.如果用Or条件需要(条件1 Or 条件2 Or ...)
oDbPager.AddCondition "Article_Status > 0"
oDbPager.AddCondition "(Article_ID < 104 Or Article_ID > 222)"
'//指定每页记录集数量
oDbPager.PageSize = iPageSize
'//指定当前页数
oDbPager.Page = Request.QueryString("page")
'//也可以直接使用自定义的SQL语句选取记录集
'oDbPager.Sql = "Select * From Kin_Article Where Article_ID < 222 Order By Article_ID Desc"
'//输出SQL语句 方便调试
'Response.Write(oDbPager.GetSql())
'Response.End()
Set oRs = oDbPager.Recordset
'//-----------------------------------------------------------------------------
'// 对该记录集的分页样式及模板进行设置(不设置则使用默认样式)
'//-----------------------------------------------------------------------------
'//选择 分页链接 输出的样式
'//为0: 可以使用样式表对分页链接进行美化(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//为1: 可使用<font>等HTML代码进行颜色设置
'oDbPager.Style = 0
'//定义 首页/上一页/下一页/末页 链接样式(支持HTML)
'oDbPager.FirstPage = "&lt;&lt;"
'oDbPager.PreviewPage = "&lt;"
'oDbPager.NextPage = "&gt;"
'oDbPager.LastPage = "&gt;&gt;"
'//定义 当前页/列表页 链接样式 {$CurrentPage}{$ListPage}将被替换成 当前页/列表页 的数字
'oDbPager.CurrentPage = "{$CurrentPage}"
'oDbPager.ListPage = "{$ListPage}"
'//定义分页列表前后要显示几个链接 如12...456...78 默认为0
'oDbPager.PagerTop = 2
'//定义分页列表最大数量 默认为7
'oDbPager.PagerSize = 5
'//定义记录集综合信息
'oDbPager.PageInfo = "共有 {$Kin_RecordCount} 记录 页次:{$Kin_Page}/{$Kin_PageCount}"
'//自定义ISAPI_REWRITE路径 * 号 将被替换为当前页数
'oDbPager.RewritePath = "Article/*.html"
'//定义跳转列表为<INPUT>文本框 默认为"SELECT"
'oDbPager.JumpPageType = "INPUT"
'//定义页面跳的SELECT/INPUT的样式(HTML代码)
'oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// 获取所需要变量以便进行输出
'//-----------------------------------------------------------------------------
'//获取当前页码
'iCurrPage = oDbPager.Page
'//获取记录集数量
'iRecordCount = oDbPager.RecordCount
'//获取页面总计数量
'iPageCount = oDbPager.PageCount
'//获取记录集信息
sPageInfo = oDbPager.PageInfo
'//获取分页信息
sPager = oDbPager.Pager
'//获取跳转列表
sJumpPage = oDbPager.JumpPage
%>
多行多列 方法一(填满整个表格)<br />
<%
'//-----------------------------------------------------------------------------
'// 例子: 动态输出M行N列, 多行多列, 循环行列, 循环TABLE
'//-----------------------------------------------------------------------------
'//初始化i准备循环
i = 0
'//定义一行最多有几列(正整数)
iCols = 3
iColsPercent = FormatPercent(1 / iCols, 0)
'//输出TABLE表头
Response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""2"" bgcolor=""#000000""><tr>")
'//循环开始 (i < iPageSize)条件为ACCESS分页时必要
Do While Not oRs.EOF
   '//每行例满了就加一个新行
   If i > 0 And i Mod iCols = 0 Then Response.Write("</tr><tr>")
   i = i + 1
   Response.Write("<td width=""" & iColsPercent & """ bgcolor=""#CCE8CF""><font color=""#000000"">" & Server.HTMLEncode(oRs(2)) & "</font></td>")
   oRs.MoveNext
Loop
'//循环结束 开始补空缺的列
Do While i < iPageSize
   '//以下两个条件二选一
   If i Mod iCols = 0 Then
       Response.Write("</tr><tr>") '//如果要补满整个表格就继续输出<tr></tr>
       'Exit Do '//如果只补满最后一行就直接结束
   End If
   i = i + 1
   Response.Write("<td width="""&FormatPercent(1 / iCols, 0)&""" bgcolor=""#CCCCCC""><font color=""red"">填满整个表格</font></td>")
Loop
'//输出分页信息/样式/TABLE尾
Response.Write("</tr><tr><td colspan=""" & iCols & """ bgcolor=""#CCE8CF""><div style=""float:right"">" & sPager & " 下拉跳转: " & sJumpPage & " 页</div>" & sPageInfo & "</td></tr></table>")
%>
多行多列 方法二(只填满最后一行)<br />
<%
oDbPager.RewritePath = "http://down.chinaz.com/soft/24767.htm#*.html"" target=""_blank"
sPager = oDbPager.Pager
oDbPager.JumpPageType = "INPUT"
sJumpPage = oDbPager.JumpPage
oRs.AbsolutePage = oDbPager.Page()
'//-----------------------------------------------------------------------------
'// 例子: 动态输出M行N列, 多行多列, 循环行列, 循环TABLE
'//-----------------------------------------------------------------------------
'//初始化i准备循环
i = 0
'//定义一行最多有几列(正整数)
iCols = 5
iColsPercent = FormatPercent(1 / iCols, 0)
'//输出TABLE表头
Response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""2"" bgcolor=""#000000""><tr>")
'//循环开始 (i < iPageSize)条件为ACCESS分页时必要
Do While Not oRs.EOF
   '//每行例满了就加一个新行
   If i > 0 And i Mod iCols = 0 Then Response.Write("</tr><tr>")
   i = i + 1
   Response.Write("<td width=""" & iColsPercent & """ bgcolor=""#CCE8CF""><font color=""#000000"">" & Server.HTMLEncode(oRs(2)) & "</font></td>")
   oRs.MoveNext
Loop
'//循环结束 开始补空缺的列
Do While i < iPageSize
   '//以下两个条件二选一
   If i Mod iCols = 0 Then
       'Response.Write("</tr><tr>") '//如果要补满整个表格就继续输出<tr></tr>
       Exit Do '//如果只补满最后一行就直接结束
   End If
   i = i + 1
   Response.Write("<td width="""&FormatPercent(1 / iCols, 0)&""" bgcolor=""#CCCCCC""><font color=""red"">只填满最后一行</font></td>")
Loop
'//输出分页信息/样式/TABLE尾
Response.Write("</tr><tr><td colspan=""" & iCols & """ bgcolor=""#CCE8CF"">" & sPageInfo & "<br />" & sPager & "<br /><font color=""blue"">&nbsp; &nbsp; 翻页链接自定义为: http://down.chinaz.com/soft/24767.htm#页数.html</font></td></tr></table>")
oRs.Close
%>
<%
'//-----------------------------------------------------------------------------
'// 对该记录集的分页样式及模板进行设置(不设置则使用默认样式)
'//-----------------------------------------------------------------------------
'//选择 分页链接 输出的样式
'//为0: 可以使用样式表对分页链接进行美化(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//为1: 可使用<font>等HTML代码进行颜色设置
oDbPager.Style = 0
'//定义 首页/上一页/下一页/末页 链接样式(支持HTML)
oDbPager.FirstPage = ""
oDbPager.PreviewPage = "<span id=""np""></span><font color=""#0000CC""><strong>Previous</strong></font>"
oDbPager.NextPage = "<span id=""nn""></span><font color=""#0000CC""><strong>Next</strong></font>"
oDbPager.LastPage = ""
'//定义 当前页/列表页 链接样式 {$CurrentPage}{$ListPage}将被替换成 当前页/列表页 的数字
oDbPager.CurrentPage = "<span id=""nc""></span>{$CurrentPage}"
oDbPager.ListPage = "<span class=""nr""></span>{$ListPage}"
'//定义分页列表前后要显示几个链接 如12...456...78 默认为0
oDbPager.PagerTop = 0
'//定义分页列表最大数量 默认为7
oDbPager.PagerSize = 9
'//定义记录集综合信息
'oDbPager.PageInfo = "共有 {$Kin_RecordCount} 记录 页次:{$Kin_Page}/{$Kin_PageCount}"
'//自定义ISAPI_REWRITE路径 * 号 将被替换为当前页数
oDbPager.RewritePath = "./?page=*"" onclick=""window.open('http://www.google.cn/search?hl=zh-CN&newwindow=1&q=kin_db_pager&start='+eval(2-1)+'0&sa=N&filter=0');"
'//定义跳转列表为<INPUT>文本框 默认为"SELECT"
oDbPager.JumpPageType = "INPUT"
'//定义页面跳的SELECT/INPUT的样式(HTML代码)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
oDbPager.SpaceMark = "</td><td>"
'//-----------------------------------------------------------------------------
'// 获取所需要变量以便进行输出
'//-----------------------------------------------------------------------------
'//获取当前页码
'iCurrPage = oDbPager.Page
'//获取记录集数量
'iRecordCount = oDbPager.RecordCount
'//获取页面总计数量
'iPageCount = oDbPager.PageCount
'//获取记录集信息
sPageInfo = oDbPager.PageInfo
'//获取分页信息
sPager = oDbPager.Pager
'//获取跳转列表
sJumpPage = oDbPager.JumpPage
%>
<style>
.google{
margin:20px;
padding:10px;
}
.google a,.google .current,.google .disabled{
display: block;
float:left;
padding:26px 2px 0px 2px;
text-align:center;
COLOR: #000;
}
.google td{font-size:12px;}
.google .current{	FONT-WEIGHT: bold;
	COLOR: #a90a08;}
.google {
	text-align:center;
	background-color:#FFFFFF;
	height:44px;

}

#np {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	MARGIN-left: auto;
	WIDTH: 44px;
	CURSOR: pointer;
}
#nf {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	BACKGROUND-POSITION: -26px 0px;
	WIDTH: 18px
}
#nc {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	BACKGROUND-POSITION: -44px 0px;
	WIDTH: 16px
}
#nn {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	BACKGROUND-POSITION: -76px 0px;
	WIDTH: 66px;
	MARGIN-RIGHT: 34px;
		CURSOR: pointer;
}
#nl {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	BACKGROUND-POSITION: -76px 0px;
	WIDTH: 46px
}
.nr {
	DISPLAY: block;
	BACKGROUND: url(Style/nav_logo3.png) no-repeat;
	HEIGHT: 26px;
	BACKGROUND-POSITION: -60px 0px;
	WIDTH: 16px;
	CURSOR: pointer;
}
</style>
<div class="google">
  <table>
    <tr>
      <td><%=sPager%></td>
    </tr>
  </table>
</div>
<%
'//-----------------------------------------------------------------------------
'// 对该记录集的分页样式及模板进行设置(不设置则使用默认样式)
'//-----------------------------------------------------------------------------
'//选择 分页链接 输出的样式
'//为0: 可以使用样式表对分页链接进行美化(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//为1: 可使用<font>等HTML代码进行颜色设置
oDbPager.Style = 0
'//定义 首页/上一页/下一页/末页 链接样式(支持HTML)
oDbPager.FirstPage = "&lt;&lt;"
oDbPager.PreviewPage = "&lt;"
oDbPager.NextPage = "&gt;"
oDbPager.LastPage = "&gt;&gt;"
'//定义 当前页/列表页 链接样式 {$CurrentPage}{$ListPage}将被替换成 当前页/列表页 的数字
oDbPager.CurrentPage = "{$CurrentPage}"
oDbPager.ListPage = "{$ListPage}"
'//定义分页列表前后要显示几个链接 如12...456...78 默认为0
oDbPager.PagerTop = 0
'//定义分页列表最大数量 默认为7
oDbPager.PagerSize = 9
oDbPager.SpaceMark = " "
'//定义记录集综合信息
'oDbPager.PageInfo = "共有 {$Kin_RecordCount} 记录 页次:{$Kin_Page}/{$Kin_PageCount}"
'//自定义ISAPI_REWRITE路径 * 号 将被替换为当前页数
oDbPager.RewritePath = "javascript:ajaxLoad(*);"
'//定义跳转列表为<INPUT>文本框 默认为"SELECT"
oDbPager.JumpPageType = "INPUT"
'//定义页面跳的SELECT/INPUT的样式(HTML代码)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// 获取所需要变量以便进行输出
'//-----------------------------------------------------------------------------
'//获取当前页码
'iCurrPage = oDbPager.Page
'//获取记录集数量
'iRecordCount = oDbPager.RecordCount
'//获取页面总计数量
'iPageCount = oDbPager.PageCount
'//获取记录集信息
sPageInfo = oDbPager.PageInfo
'//获取分页信息
sPager = oDbPager.Pager
'//获取跳转列表
sJumpPage = oDbPager.JumpPage
%>
<script>
function ajaxLoad(i){
	alert('ajax读取第 '+i+' 页.');
//	location.href='?<%=oDbPager.PageParam%>='+i;
}
</script>
<style>
.ajax { background-color:#ccFFee; margin:5px; padding:5px;}
</style>
<div class="ajax"><%="ajax 分页 " & sPageInfo & " " & sPager & " ajax数字分页:" & sJumpPage%></div>
<%
'//-----------------------------------------------------------------------------
'// 对该记录集的分页样式及模板进行设置(不设置则使用默认样式)
'//-----------------------------------------------------------------------------
'//选择 分页链接 输出的样式
'//为0: 可以使用样式表对分页链接进行美化(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//为1: 可使用<font>等HTML代码进行颜色设置
oDbPager.Style = 0
'//定义 首页/上一页/下一页/末页 链接样式(支持HTML)
oDbPager.FirstPage = "&lt;&lt;"
oDbPager.PreviewPage = "&lt;"
oDbPager.NextPage = "&gt;"
oDbPager.LastPage = "&gt;&gt;"
'//定义 当前页/列表页 链接样式 {$CurrentPage}{$ListPage}将被替换成 当前页/列表页 的数字
oDbPager.CurrentPage = "{$CurrentPage}"
oDbPager.ListPage = "{$ListPage}"
'//定义分页列表前后要显示几个链接 如12...456...78 默认为0
oDbPager.PagerTop = 2
'//定义分页列表最大数量 默认为7
oDbPager.PagerSize = 9
'//定义记录集综合信息
'oDbPager.PageInfo = "共有 {$Kin_RecordCount} 记录 页次:{$Kin_Page}/{$Kin_PageCount}"
'//自定义ISAPI_REWRITE路径 * 号 将被替换为当前页数
oDbPager.RewritePath = ""
'//定义跳转列表为<INPUT>文本框 默认为"SELECT"
oDbPager.JumpPageType = "INPUT"
'//定义页面跳的SELECT/INPUT的样式(HTML代码)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// 获取所需要变量以便进行输出
'//-----------------------------------------------------------------------------
'//获取当前页码
'iCurrPage = oDbPager.Page
'//获取记录集数量
'iRecordCount = oDbPager.RecordCount
'//获取页面总计数量
'iPageCount = oDbPager.PageCount
'//获取记录集信息
sPageInfo = oDbPager.PageInfo
'//获取分页信息
sPager = oDbPager.Pager
'//获取跳转列表
sJumpPage = oDbPager.JumpPage
%>
<style>
div.kindbpager0 {
	border:1px solid #0066FF;
	margin:2px;
	padding:2px;
	text-align:center;
	height:30px;
	font-size:14px;
	color:red;
	background-color:#FFFFFF;
}
div.kindbpager0 input {
	border:1px solid #0033FF;
}
div.kindbpager0 a, div.kindbpager0 .current, div.kindbpager0 .disabled {
	text-decoration:none;
	color:#FF0000;
	border:1px solid #0033FF;
	padding:1px 5px;
	background-color:#FFFFFF;
}
div.kindbpager0 .current {
	background-color:#00CCFF;
}
div.kindbpager0 .disabled {
	background-color:#DDDDDD;
}
div.kindbpager0 a:hover, div.kindbpager0 a:actived {
 text-decoration: underline;
 background-color:#66FFFF;
}
</style>
<div class="kindbpager0"> <%= sPageInfo %> <%= sPager %> 输入数字跳转: <%= sJumpPage %> 页</div>
<p style="background-color:red;padding:10px;">从这里往下oDbPager.PagerTop = 2</p>
<p><a href="http://www.digg.com">Digg</a>Style</p>
<div class="digg"><%= sPager %></div>
<p><a href="http://www.yahoo.com">Yahoo</a>Style</p>
<div class="yahoo"><%= sPager %></div>
<p><a href="http://www.yahoo.com">New Yahoo!</a>Style</p>
<div class="yahoo2"><%= sPager %></div>
<p><a href="http://www.meneame.net">Meneame</a>Style</p>
<div class="meneame"><%= sPager %></div>
<p><a href="http://www.flickr.com">Flickr</a>Style</p>
<div class="flickr"><%= sPager %></div>
<p><a href="http://sabros.us">Sabros.us</a>Style</p>
<div class="sabrosus"><%= sPager %></div>
<p>Green Style</p>
<div class="scott"><%= sPager %></div>
<p>Gray Style</p>
<div class="quotes"><%= sPager %></div>
<p>Black Style</p>
<div class="black"><%= sPager %></div>
<p><a href="http://www.mis-algoritmos.com">Mis Algoritmos</a>Style</p>
<div class="black2"><%= sPager %></div>
<p>Black-Red Style</p>
<div style="padding-top:10px;padding-bottom:10px;background-color:#313131;">
  <div class="black-red"><%= sPager %></div>
</div>
<%
oDbPager.PagerTop = 0
sPager = oDbPager.Pager
%>
<p style="background-color:red;padding:10px;">从这里往下oDbPager.PagerTop = 0</p>
<p>Gray Style 2</p>
<div class="grayr"><%= sPager %></div>
<p>Yellow Style</p>
<div class="yellow"><%= sPager %></div>
<p><a href="http://jogger.pl/">jogger</a>Style</p>
<div class="jogger"><%= sPager %></div>
<p><a href="http://eu.starcraft2.com/screenshots.xml">starcraft 2</a>Style</p>
<div class="starcraft2"><%= sPager %></div>
<p>Tres Style</p>
<div class="tres"><%= sPager %></div>
<p><a href="http://www.512megas.com">512megas</a>Style</p>
<div class="megas512"><%= sPager %></div>
<p><a href="http://www.technorati.com/">Technorati</a>Style</p>
<div class="technorati"><%= sPager %></div>
<p><a href="http://www.youtube.com/">YouTube</a>Style</p>
<div class="youtube"><%= sPager %></div>
<p><a href="http://search.msdn.microsoft.com/">MSDN Search</a>Style</p>
<div class="msdn"><%= sPager %></div>
<p><a href="http://badoo.com/">Badoo</a>
<div class="badoo"><%= sPager %></div>
<p>Blue Style</p>
<div class="manu"><%= sPager %></div>
<p>Green-Black Style</p>
<div class="green-black"><%= sPager %></div>
<p>viciao Style</p>
<div class="viciao"><%= sPager %></div>
<div style="display:none">
  <script language="javascript" type="text/javascript" src="http://js.users.51.la/1269467.js"></script>
  <noscript>
  <a href="http://www.51.la/?1269467" target="_blank"><img alt="&#x6211;&#x8981;&#x5566;&#x514D;&#x8D39;&#x7EDF;&#x8BA1;" src="http://img.users.51.la/1269467.asp" style="border:none" /></a>
  </noscript>
</div>
</body>
</html>
