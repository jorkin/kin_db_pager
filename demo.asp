<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True
Response.Charset = "utf-8"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Kin_Db_Pager 实例 -- Jorkin.ME</title>
</head>
<body>
<!--#include file="Conn.asp" -->
<!--#include file="Jorkin_Function.asp" -->
<!--#include file="Kin_Db_Pager.asp" -->
<%
'//说明:部分函数可能需要与Jorkin_Function.asp配合使用。
OpenConn()
Dim oPager, sKeyword, oRs, aRS, sJSON, sJSArray, sSql, i, j
Dim sPager, sJumpPager1, sJumpPager2, sPagerInfo
Set oPager = New Kin_Db_Pager
With oPager
    '.IsDebug = True '//调试模式，输出错误信息，正式发布前请注释掉。
    .CacheType = 0 '//缓存类型Application(0),Session(1),Cookies(2),其它不缓存
    .CacheTimeOut = 30 '//缓存时间（单位：分钟）
    .Connect(oConn) '//oConn是打开的数据库连接[推荐]。
    '.DbType = "MSSQLPRODUCE" '//设为存储过程分页
    .TableName = "Kin_Article"
    .PKey = "Article_ID"
    '.PKeyOrder = "DESC"
    .Fields = "*"
    '.Distinct = True
    '.Speed = 0
    '.AddCondition "Article_ID <= 50050"
    '.AddCondition "(Article_ID <= 10000 Or Article_ID > 30000)"
    sKeyword = to_Str(Request.QueryString("keyword"))
    If Not IsBlank(sKeyword) Then
        .AddCondition "Article_Content LIKE '%" & Str4Like(sKeyword) & "%'" '//注意过滤SQL注入。参考防注入函数。
    End If
    '.Condition = Null '//可以用此方法清空查询条件。
    .OrderBy = "Article_ID"
    '.AddOrderBy "ID desc"
    '.PageParam = "page"
	'.Page1Size = 15
    .PageSize = 20

    '.MaxRecords = 1000
    'If Not IsBlank(sKeyword) Then
    '    .CacheType = 0 '//有关键字时不缓存 
    '    .ReWritePath = "/Article/List/" & Server.URLEncode(sKeyword) & "/*.html"'//有关键字时，用关键字。
    'Else
    '    .ReWritePath = "/Article/List/*.html" '//无关键字时，直接用列表。
    'End If
    '.ReWritePath = "javascript:ajaxPager(*)" '//使用脚本翻页，需自行编写ajaxPager函数，可实现AJAX翻页。
    .PagerStyle = 3 '//尝试不同的内置风格，体会强大的自定义导航功能吧。（1,2,3....）
    '.PagerTop = 2
    '.PagerSize = 10
    '.PagerGroup = True
    sPager = .Pager
    sJumpPager1 = .JumpPager("BUTTON", "style=""border:3px double #CCC;margin:auto 2px;""") '//参数的顺序和数量都可自定义。
    sJumpPager2 = .JumpPager("SELECT", "style=""background:#CCC;""") '//参数的顺序和数量都可自定义。
    sPagerInfo = .PagerInfo
    Set oRs = .Recordset() '//执行结果的记录集。
    'aRS = .GetRows() '//将多个记录检索到数组中。
    sSql = .GetSQL() '//输出调试生成的SQL语句。
	'sJSON = .GetJSON()
	'sJSArray = .GetJSArray()
	'trace sSql
End With
echo("<table width=""100%"" border=""1""><tr><td><font color=""red""><nobr>rownum<nobr></font></td>")
For i = 0 To oRs.Fields.Count - 1
    echo("<td>" & oRs(i).Name & "</td>")
Next
Do Until oRs.EOF
    j = j + 1
    echo("<tr><td>" & j & "</td>")
    For i = 0 To oRs.Fields.Count - 1
        If IsNull(oRs(i)) Then
            echo("<td><font color=""red"">&lt;NULL&gt;</font></td>")
        Else
            echo("<td>" & to_HTML(oRs(i)) & "</td>")
        End If
    Next
    echo("</tr>")
    oRs.MoveNext
Loop
echo("</table>")
%>
<script>function ajaxPager(n){location.href="?page="+n;}</script>
<style type="text/css"></style>
<style type="text/css">
<!--
div.kpager{background-color:#000;color:#a0a0a0;margin:3px;padding:10px 3px;text-align:center;font-size:14px;}
div.kpager a,div.kpager span.ellipsis{border:#909090 1px solid;color:#c0c0c0;margin-right:3px;padding:2px 5px;text-decoration:none;}
div.kpager a:hover,div.kpager a:active{background-color:#404040 border:#f0f0f0 1px solid;color:#ffffff;}
div.kpager span.current{background-color:#606060;border:#ffffff 1px solid;color:#ffffff;font-weight:bold;margin-right:3px;padding:2px 5px;}
div.kpager span.disabled{border:#606060 1px solid;color:#808080;margin-right:3px;padding:2px 5px;}
/* div.kpager span.disabled {display:none;} /*取消左侧注释看看有什么不同*/
-->
</style>
<div class="kpager">
<p><%=sPager & sJumpPager2%></p>
<p><%=sPagerInfo & sJumpPager1%></p>
<p><%=sSql%></p>
</div>
<%Set oPager = Nothing%>
</body>
</html>
