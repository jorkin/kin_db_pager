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
<title>Kin_Db_Pager ��ҳ�� - KinJAVA��־</title>
</head>
<body>
<p>Kin_Db_Pager ��ҳ��(<a href="http://jorkin.reallydo.com" target="_blank">KinJAVA��־</a> <a href="http://jorkin.reallydo.com/article.asp?id=534" target="_blank">����</a>) <a href="help.asp">ʹ�ð���</a></p>
<%
'//-----------------------------------------------------------------------------
'// ����Eg()������ر��� ���δʹ��Option Explicit��ʡ��
'//-----------------------------------------------------------------------------
'Dim i, iCols, iColsPercent, iPageSize
'Dim iCurrPage, iRecordCount, iPageCount
'Dim sPageInfo, sPager, sJumpPage
Dim oDbPager
'//-----------------------------------------------------------------------------
'// Kin_Db_Pager��ҳ�࿪ʼ
'//-----------------------------------------------------------------------------
iPageSize = 15
Set oDbPager = New Kin_db_Pager
'//-----------------------------------------------------------------------------
'// �������ݿ��ѯǰ����ز�������
'//-----------------------------------------------------------------------------
'//ָ�����ݿ�����
oDbPager.Connect(oConn) '//����һ(�Ƽ�)
'Set oDbPager.ActiveConnection = oConn '//������
'oDbPager.ConnectionString = oConn.ConnectionString '//������
'//ָ����ʾҳ����URL���� Ĭ��ֵ:"page"
'oDbPager.PageParam = "page"
'//ָ�����ݿ�����.Ĭ��ֵ:"MSSQL"
oDbPager.DbType = "ACCESS"
'//ָ��Ŀ��� ������ʱ��"(Select * From [Table])"
oDbPager.TableName = "Kin_Article"
'//ѡ���� �ö��ŷָ� Ĭ��ֵ:"*"
oDbPager.Fields = "*"
'//ָ���ñ������
oDbPager.PKey = "Article_ID"
'//ָ����������
oDbPager.OrderBy = "Article_ID DESC"
'//������� �ɶ��ʹ��.�����Or������Ҫ(����1 Or ����2 Or ...)
oDbPager.AddCondition "Article_Status > 0"
oDbPager.AddCondition "(Article_ID < 104 Or Article_ID > 222)"
'//ָ��ÿҳ��¼������
oDbPager.PageSize = iPageSize
'//ָ����ǰҳ��
oDbPager.Page = Request.QueryString("page")
'//Ҳ����ֱ��ʹ���Զ����SQL���ѡȡ��¼��
'oDbPager.Sql = "Select * From Kin_Article Where Article_ID < 222 Order By Article_ID Desc"
'//���SQL��� �������
'Response.Write(oDbPager.GetSql())
'Response.End()
Set oRs = oDbPager.Recordset
'//-----------------------------------------------------------------------------
'// �Ըü�¼���ķ�ҳ��ʽ��ģ���������(��������ʹ��Ĭ����ʽ)
'//-----------------------------------------------------------------------------
'//ѡ�� ��ҳ���� �������ʽ
'//Ϊ0: ����ʹ����ʽ��Է�ҳ���ӽ�������(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//Ϊ1: ��ʹ��<font>��HTML���������ɫ����
'oDbPager.Style = 0
'//���� ��ҳ/��һҳ/��һҳ/ĩҳ ������ʽ(֧��HTML)
'oDbPager.FirstPage = "&lt;&lt;"
'oDbPager.PreviewPage = "&lt;"
'oDbPager.NextPage = "&gt;"
'oDbPager.LastPage = "&gt;&gt;"
'//���� ��ǰҳ/�б�ҳ ������ʽ {$CurrentPage}{$ListPage}�����滻�� ��ǰҳ/�б�ҳ ������
'oDbPager.CurrentPage = "{$CurrentPage}"
'oDbPager.ListPage = "{$ListPage}"
'//�����ҳ�б�ǰ��Ҫ��ʾ�������� ��12...456...78 Ĭ��Ϊ0
'oDbPager.PagerTop = 2
'//�����ҳ�б�������� Ĭ��Ϊ7
'oDbPager.PagerSize = 5
'//�����¼���ۺ���Ϣ
'oDbPager.PageInfo = "���� {$Kin_RecordCount} ��¼ ҳ��:{$Kin_Page}/{$Kin_PageCount}"
'//�Զ���ISAPI_REWRITE·�� * �� �����滻Ϊ��ǰҳ��
'oDbPager.RewritePath = "Article/*.html"
'//������ת�б�Ϊ<INPUT>�ı��� Ĭ��Ϊ"SELECT"
'oDbPager.JumpPageType = "INPUT"
'//����ҳ������SELECT/INPUT����ʽ(HTML����)
'oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// ��ȡ����Ҫ�����Ա�������
'//-----------------------------------------------------------------------------
'//��ȡ��ǰҳ��
'iCurrPage = oDbPager.Page
'//��ȡ��¼������
'iRecordCount = oDbPager.RecordCount
'//��ȡҳ���ܼ�����
'iPageCount = oDbPager.PageCount
'//��ȡ��¼����Ϣ
sPageInfo = oDbPager.PageInfo
'//��ȡ��ҳ��Ϣ
sPager = oDbPager.Pager
'//��ȡ��ת�б�
sJumpPage = oDbPager.JumpPage
%>
���ж��� ����һ(�����������)<br />
<%
'//-----------------------------------------------------------------------------
'// ����: ��̬���M��N��, ���ж���, ѭ������, ѭ��TABLE
'//-----------------------------------------------------------------------------
'//��ʼ��i׼��ѭ��
i = 0
'//����һ������м���(������)
iCols = 3
iColsPercent = FormatPercent(1 / iCols, 0)
'//���TABLE��ͷ
Response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""2"" bgcolor=""#000000""><tr>")
'//ѭ����ʼ (i < iPageSize)����ΪACCESS��ҳʱ��Ҫ
Do While Not oRs.EOF
   '//ÿ�������˾ͼ�һ������
   If i > 0 And i Mod iCols = 0 Then Response.Write("</tr><tr>")
   i = i + 1
   Response.Write("<td width=""" & iColsPercent & """ bgcolor=""#CCE8CF""><font color=""#000000"">" & Server.HTMLEncode(oRs(2)) & "</font></td>")
   oRs.MoveNext
Loop
'//ѭ������ ��ʼ����ȱ����
Do While i < iPageSize
   '//��������������ѡһ
   If i Mod iCols = 0 Then
       Response.Write("</tr><tr>") '//���Ҫ�����������ͼ������<tr></tr>
       'Exit Do '//���ֻ�������һ�о�ֱ�ӽ���
   End If
   i = i + 1
   Response.Write("<td width="""&FormatPercent(1 / iCols, 0)&""" bgcolor=""#CCCCCC""><font color=""red"">�����������</font></td>")
Loop
'//�����ҳ��Ϣ/��ʽ/TABLEβ
Response.Write("</tr><tr><td colspan=""" & iCols & """ bgcolor=""#CCE8CF""><div style=""float:right"">" & sPager & " ������ת: " & sJumpPage & " ҳ</div>" & sPageInfo & "</td></tr></table>")
%>
���ж��� ������(ֻ�������һ��)<br />
<%
oDbPager.RewritePath = "http://down.chinaz.com/soft/24767.htm#*.html"" target=""_blank"
sPager = oDbPager.Pager
oDbPager.JumpPageType = "INPUT"
sJumpPage = oDbPager.JumpPage
oRs.AbsolutePage = oDbPager.Page()
'//-----------------------------------------------------------------------------
'// ����: ��̬���M��N��, ���ж���, ѭ������, ѭ��TABLE
'//-----------------------------------------------------------------------------
'//��ʼ��i׼��ѭ��
i = 0
'//����һ������м���(������)
iCols = 5
iColsPercent = FormatPercent(1 / iCols, 0)
'//���TABLE��ͷ
Response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""2"" bgcolor=""#000000""><tr>")
'//ѭ����ʼ (i < iPageSize)����ΪACCESS��ҳʱ��Ҫ
Do While Not oRs.EOF
   '//ÿ�������˾ͼ�һ������
   If i > 0 And i Mod iCols = 0 Then Response.Write("</tr><tr>")
   i = i + 1
   Response.Write("<td width=""" & iColsPercent & """ bgcolor=""#CCE8CF""><font color=""#000000"">" & Server.HTMLEncode(oRs(2)) & "</font></td>")
   oRs.MoveNext
Loop
'//ѭ������ ��ʼ����ȱ����
Do While i < iPageSize
   '//��������������ѡһ
   If i Mod iCols = 0 Then
       'Response.Write("</tr><tr>") '//���Ҫ�����������ͼ������<tr></tr>
       Exit Do '//���ֻ�������һ�о�ֱ�ӽ���
   End If
   i = i + 1
   Response.Write("<td width="""&FormatPercent(1 / iCols, 0)&""" bgcolor=""#CCCCCC""><font color=""red"">ֻ�������һ��</font></td>")
Loop
'//�����ҳ��Ϣ/��ʽ/TABLEβ
Response.Write("</tr><tr><td colspan=""" & iCols & """ bgcolor=""#CCE8CF"">" & sPageInfo & "<br />" & sPager & "<br /><font color=""blue"">&nbsp; &nbsp; ��ҳ�����Զ���Ϊ: http://down.chinaz.com/soft/24767.htm#ҳ��.html</font></td></tr></table>")
oRs.Close
%>
<%
'//-----------------------------------------------------------------------------
'// �Ըü�¼���ķ�ҳ��ʽ��ģ���������(��������ʹ��Ĭ����ʽ)
'//-----------------------------------------------------------------------------
'//ѡ�� ��ҳ���� �������ʽ
'//Ϊ0: ����ʹ����ʽ��Է�ҳ���ӽ�������(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//Ϊ1: ��ʹ��<font>��HTML���������ɫ����
oDbPager.Style = 0
'//���� ��ҳ/��һҳ/��һҳ/ĩҳ ������ʽ(֧��HTML)
oDbPager.FirstPage = ""
oDbPager.PreviewPage = "<span id=""np""></span><font color=""#0000CC""><strong>Previous</strong></font>"
oDbPager.NextPage = "<span id=""nn""></span><font color=""#0000CC""><strong>Next</strong></font>"
oDbPager.LastPage = ""
'//���� ��ǰҳ/�б�ҳ ������ʽ {$CurrentPage}{$ListPage}�����滻�� ��ǰҳ/�б�ҳ ������
oDbPager.CurrentPage = "<span id=""nc""></span>{$CurrentPage}"
oDbPager.ListPage = "<span class=""nr""></span>{$ListPage}"
'//�����ҳ�б�ǰ��Ҫ��ʾ�������� ��12...456...78 Ĭ��Ϊ0
oDbPager.PagerTop = 0
'//�����ҳ�б�������� Ĭ��Ϊ7
oDbPager.PagerSize = 9
'//�����¼���ۺ���Ϣ
'oDbPager.PageInfo = "���� {$Kin_RecordCount} ��¼ ҳ��:{$Kin_Page}/{$Kin_PageCount}"
'//�Զ���ISAPI_REWRITE·�� * �� �����滻Ϊ��ǰҳ��
oDbPager.RewritePath = "./?page=*"" onclick=""window.open('http://www.google.cn/search?hl=zh-CN&newwindow=1&q=kin_db_pager&start='+eval(2-1)+'0&sa=N&filter=0');"
'//������ת�б�Ϊ<INPUT>�ı��� Ĭ��Ϊ"SELECT"
oDbPager.JumpPageType = "INPUT"
'//����ҳ������SELECT/INPUT����ʽ(HTML����)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
oDbPager.SpaceMark = "</td><td>"
'//-----------------------------------------------------------------------------
'// ��ȡ����Ҫ�����Ա�������
'//-----------------------------------------------------------------------------
'//��ȡ��ǰҳ��
'iCurrPage = oDbPager.Page
'//��ȡ��¼������
'iRecordCount = oDbPager.RecordCount
'//��ȡҳ���ܼ�����
'iPageCount = oDbPager.PageCount
'//��ȡ��¼����Ϣ
sPageInfo = oDbPager.PageInfo
'//��ȡ��ҳ��Ϣ
sPager = oDbPager.Pager
'//��ȡ��ת�б�
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
'// �Ըü�¼���ķ�ҳ��ʽ��ģ���������(��������ʹ��Ĭ����ʽ)
'//-----------------------------------------------------------------------------
'//ѡ�� ��ҳ���� �������ʽ
'//Ϊ0: ����ʹ����ʽ��Է�ҳ���ӽ�������(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//Ϊ1: ��ʹ��<font>��HTML���������ɫ����
oDbPager.Style = 0
'//���� ��ҳ/��һҳ/��һҳ/ĩҳ ������ʽ(֧��HTML)
oDbPager.FirstPage = "&lt;&lt;"
oDbPager.PreviewPage = "&lt;"
oDbPager.NextPage = "&gt;"
oDbPager.LastPage = "&gt;&gt;"
'//���� ��ǰҳ/�б�ҳ ������ʽ {$CurrentPage}{$ListPage}�����滻�� ��ǰҳ/�б�ҳ ������
oDbPager.CurrentPage = "{$CurrentPage}"
oDbPager.ListPage = "{$ListPage}"
'//�����ҳ�б�ǰ��Ҫ��ʾ�������� ��12...456...78 Ĭ��Ϊ0
oDbPager.PagerTop = 0
'//�����ҳ�б�������� Ĭ��Ϊ7
oDbPager.PagerSize = 9
oDbPager.SpaceMark = " "
'//�����¼���ۺ���Ϣ
'oDbPager.PageInfo = "���� {$Kin_RecordCount} ��¼ ҳ��:{$Kin_Page}/{$Kin_PageCount}"
'//�Զ���ISAPI_REWRITE·�� * �� �����滻Ϊ��ǰҳ��
oDbPager.RewritePath = "javascript:ajaxLoad(*);"
'//������ת�б�Ϊ<INPUT>�ı��� Ĭ��Ϊ"SELECT"
oDbPager.JumpPageType = "INPUT"
'//����ҳ������SELECT/INPUT����ʽ(HTML����)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// ��ȡ����Ҫ�����Ա�������
'//-----------------------------------------------------------------------------
'//��ȡ��ǰҳ��
'iCurrPage = oDbPager.Page
'//��ȡ��¼������
'iRecordCount = oDbPager.RecordCount
'//��ȡҳ���ܼ�����
'iPageCount = oDbPager.PageCount
'//��ȡ��¼����Ϣ
sPageInfo = oDbPager.PageInfo
'//��ȡ��ҳ��Ϣ
sPager = oDbPager.Pager
'//��ȡ��ת�б�
sJumpPage = oDbPager.JumpPage
%>
<script>
function ajaxLoad(i){
	alert('ajax��ȡ�� '+i+' ҳ.');
//	location.href='?<%=oDbPager.PageParam%>='+i;
}
</script>
<style>
.ajax { background-color:#ccFFee; margin:5px; padding:5px;}
</style>
<div class="ajax"><%="ajax ��ҳ " & sPageInfo & " " & sPager & " ajax���ַ�ҳ:" & sJumpPage%></div>
<%
'//-----------------------------------------------------------------------------
'// �Ըü�¼���ķ�ҳ��ʽ��ģ���������(��������ʹ��Ĭ����ʽ)
'//-----------------------------------------------------------------------------
'//ѡ�� ��ҳ���� �������ʽ
'//Ϊ0: ����ʹ����ʽ��Է�ҳ���ӽ�������(http://jorkin.reallydo.com/kin_db_pager/?page=10)
'//Ϊ1: ��ʹ��<font>��HTML���������ɫ����
oDbPager.Style = 0
'//���� ��ҳ/��һҳ/��һҳ/ĩҳ ������ʽ(֧��HTML)
oDbPager.FirstPage = "&lt;&lt;"
oDbPager.PreviewPage = "&lt;"
oDbPager.NextPage = "&gt;"
oDbPager.LastPage = "&gt;&gt;"
'//���� ��ǰҳ/�б�ҳ ������ʽ {$CurrentPage}{$ListPage}�����滻�� ��ǰҳ/�б�ҳ ������
oDbPager.CurrentPage = "{$CurrentPage}"
oDbPager.ListPage = "{$ListPage}"
'//�����ҳ�б�ǰ��Ҫ��ʾ�������� ��12...456...78 Ĭ��Ϊ0
oDbPager.PagerTop = 2
'//�����ҳ�б�������� Ĭ��Ϊ7
oDbPager.PagerSize = 9
'//�����¼���ۺ���Ϣ
'oDbPager.PageInfo = "���� {$Kin_RecordCount} ��¼ ҳ��:{$Kin_Page}/{$Kin_PageCount}"
'//�Զ���ISAPI_REWRITE·�� * �� �����滻Ϊ��ǰҳ��
oDbPager.RewritePath = ""
'//������ת�б�Ϊ<INPUT>�ı��� Ĭ��Ϊ"SELECT"
oDbPager.JumpPageType = "INPUT"
'//����ҳ������SELECT/INPUT����ʽ(HTML����)
oDbPager.JumpPageAttr = "class=""reallydo"" style=""color:#820222"""
'//-----------------------------------------------------------------------------
'// ��ȡ����Ҫ�����Ա�������
'//-----------------------------------------------------------------------------
'//��ȡ��ǰҳ��
'iCurrPage = oDbPager.Page
'//��ȡ��¼������
'iRecordCount = oDbPager.RecordCount
'//��ȡҳ���ܼ�����
'iPageCount = oDbPager.PageCount
'//��ȡ��¼����Ϣ
sPageInfo = oDbPager.PageInfo
'//��ȡ��ҳ��Ϣ
sPager = oDbPager.Pager
'//��ȡ��ת�б�
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
<div class="kindbpager0"> <%= sPageInfo %> <%= sPager %> ����������ת: <%= sJumpPage %> ҳ</div>
<p style="background-color:red;padding:10px;">����������oDbPager.PagerTop = 2</p>
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
<p style="background-color:red;padding:10px;">����������oDbPager.PagerTop = 0</p>
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
