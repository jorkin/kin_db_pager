<%
'/*------------------------------------------------ Jorkin 自定义类 翻页优化代码
' ********************************************************************* 说    明
' * 来源: KinJAVA日志 (http://jorkin.reallydo.com/article.asp?id=534)
' * 最后更新: 2009-03-22
' * 当前版本: Ver: 1.09
' ********************************************************************* 特色功能
' * 有方便的 Eg() 实例,不需要记住每个变量的名称,为开发者提供方便.
' * ISAPI_REWRITE 功能可以轻松实现静态(伪静态)翻页以及ajax翻页.
' ********************************************************************* 更新历史
' * 2009-03-22
' *   修正qsliuliu反馈自定义PageParam后分页错误BUG。
' * 2009-03-19
' *   重要更新，优化，提高速度。
' * 2009-02-14
' *   优化ReWrite()函数，大大提高了效率。
' *   加入新函数GetCondition()，用来做多条件搜索，不过需要Jorkin_Function.asp库。
' *   方法为：GetCondition("表单字段名", "表单比较运算符", "表单关键字")。
' *   比较运算符可选(<, =, >, <=, >=, <>, !=, !<, !>, like, not like)，关键字可用?表示单字符，用*表示零个或更多字符。
' *   警告！！！前台页面慎用，会显露数据库字段名称，可参考GetCondition()自行修改。
' * 2009-02-11
' *   小修小改，加强了几个判断，更新一些说明，想用存储过程的请下载叶子分页类的sp_Util_Page.sql，本类完全兼容。
' * 2008-11-26
' *   根据数据库连接自动判断数据库类型。
' * 2008-09-09
' *   修正 Eg() 实例BUG。
' *   继续完善一直烂尾的Select Where In排序功能，还是未全部完成。(方法想命名为Kin_Db_Pager.OrderIn(字段名,排序))
' *   修正了几个马虎导致的拼写错误。 (-_-#)
' *   加入了自定义翻页样式时设定空值的判断。
' * 2008-08-28
' *   增加 Connect() 方法进行数据库连接，比 ActiveConnection 和 ConnectionString 更安全有效。
' *   修正Bug: 如果数字跳转INPUT框在一个FORM里，回车时将会进行提交表单操作。
' *   修正Bug: 数字跳转INPUT框不支持自定义ISAPI_REWRITE路径。
' *   修正Bug: 使用自定义SQL语句时分页出错。
' *   重写 Eg() 样例,使其更容易被理解。
' *   删除大量无用代码.
' * Ver: 1.03之前
' *   一行代码即可实现帮助,不需要记住所有的属性设定。
' *   请先使用 Eg() 查看生成的代码,将其全选复制放入ASP代码块内即为本分页类的操作模板。
' ********************************************************************* 鸣    谢
' * 感谢以下大大的分页类思想及代码:
' * Sunrise_Chen (http://www.ccopus.com)
' * 才子 (http://www.54caizi.org)
' * 风声 (http://www.fonshen.com)
' * 叶子 (http://www.yeeh.org)
'*/-----------------------------------------------------------------------------

Class Kin_Db_Pager

    '//-------------------------------------------------------------------------
    '// 定义变量 开始

    Private oConn '//连接对象
    Private sDbType '//数据库类型
    Private sTableName '//表名
    Private sPKey '//主键
    Private sFields '//输出的字段名
    Private sOrderBy '//排序字符串
    Private sSql '//当前的查询语句
    Private sSqlString '//自定义Sql语句
    Private aCondition() '//查询条件(数组)
    Private sCondition '//查询条件(字符串)
    Private iPage '//当前页码
    Private iPageSize '//每页记录数
    Private iPageCount '//总页数
    Private iRecordCount '//当前查询条件下的记录数
    Private sPage '//当前页 替换字符串
    Private sPageCount '//总页数 替换字符串
    Private sRecordCount '//当前查询条件下的记录数 替换字符串
    Private sProjectName '//项目名
    Private sVersion '//版本号
    Private bShowError '//是否显示错误信息
    Private bDistinct '//是否显示唯一记录
    Private sPageInfo '//记录数、页码等信息
    Private sPageParam '//page参数名称
    Private iStyle '//翻页的样式
    Private iPagerSize '//翻页按钮的数值
    Private iCurrentPageSize '//当前页面记录数量
    Private sReWrite '//用ISAP REWRITE做的路径,可用Javascript函数实现AJAX翻页
    Private iTableKind '//表的类型, 是否需要强制加 [ ]
    Private sFirstPage '//首页链接 样式
    Private sPreviewPage '//上一页链接 样式
    Private sCurrentPage '//当前页链接 样式
    Private sListPage '//分页列表链接 样式
    Private sNextPage '//下一页链接 样式
    Private sLastPage '//末页链接 样式
    Private iPagerTop '//分页列表头尾数量
    Private iPagerGroup '//多少页做为一组
    Private sJumpPage '//分页跳转功能
    Private sJumpPageType '//分页跳转类型(可选SELECT或INPUT)
    Private sJumpPageAttr '//分页跳转其他HTML属性
    Private sUrl, sQueryString, x, y
    Private sSpaceMark '//链接之前间隔符

    '//定义变量 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//事件、方法: 类初始化事件 开始

    Private Sub Class_Initialize()
        ReDim aCondition( -1)
        sProjectName = "Jorkin &#25968;&#25454;&#24211;&#20998;&#39029;&#31867;  Kin_Db_Pager"
        sDbType = "MSSQL"
        sVersion = "Ver: 1.09 Build: 090322"
        sPKey = "ID"
        sFields = "*"
        sCondition = ""
        sOrderBy = ""
        sSqlString = ""
        iPageSize = 20
        iPage = 1
        iRecordCount = Null
        iPageCount = Null
        bShowError = True
        bDistinct = False
        iPagerTop = 0
        sPage = "{$Kin_Page}"
        sPageCount = "{$Kin_PageCount}"
        sRecordCount = "{$Kin_RecordCount}"
        sPageInfo = "&#20849;&#26377;  {$Kin_RecordCount} &#26465;&#35760;&#24405;  &#39029;&#27425; : {$Kin_Page}/{$Kin_PageCount}"
        sPageParam = "page"
        setPageParam(sPageParam)
        iStyle = 29252888
        iTableKind = 0
        iPagerSize = 7
        sFirstPage = "[&#39318;&#39029;]"
        sPreviewPage = "[&#19978;&#19968;&#39029;]"
        sCurrentPage = "[{$CurrentPage}]"
        sListPage = "[{$ListPage}]"
        sNextPage = "[&#19979;&#19968;&#39029;]"
        sLastPage = "[&#26411;&#39029;]"
        sJumpPage = ""
        sJumpPageType = "SELECT"
        sSpaceMark = " "
    End Sub

    '//类结束事件

    Private Sub Class_Terminate()
        Set oConn = Nothing
    End Sub

    '//事件、方法: 类初始化事件 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//函数、方法 开始

    '功能:ASP里的IIF
    '来源:http://jorkin.reallydo.com/article.asp?id=26

    Private Function IIf(bExp1, sVal1, sVal2)
        If (bExp1) Then
            IIf = sVal1
        Else
            IIf = sVal2
        End If
    End Function

    '功能:只取数字
    '来源:http://jorkin.reallydo.com/article.asp?id=395

    Private Function Bint(sValue)
        On Error Resume Next
        Bint = 0
        Bint = Fix(CDbl(sValue))
    End Function

    '功能:判断是否是空值
    '来源:http://jorkin.reallydo.com/article.asp?id=386

    Private Function IsBlank(byref TempVar)
        IsBlank = False
        Select Case VarType(TempVar)
            Case 0, 1
                IsBlank = True
            Case 8
                If Len(TempVar) = 0 Then
                    IsBlank = True
                End If
            Case 9
                tmpType = TypeName(TempVar)
                If (tmpType = "Nothing") Or (tmpType = "Empty") Then
                    IsBlank = True
                End If
            Case 8192, 8204, 8209
                If UBound(TempVar) = -1 Then
                    IsBlank = True
                End If
        End Select
    End Function

    '//检查数据库连接是否可用

    Public Function Connect(o)
        If TypeName(o) <> "Connection" Then
            doError "无效的数据库连接。"
        Else
            If o.State = 1 Then
                Set oConn = o
                sDbType = GetDbType(oConn)
            Else
                doError "数据库连接已关闭。"
            End If
        End If
    End Function

    '//处理错误信息

    Public Sub doError(s)
        On Error Resume Next
		If Not bShowError Then Exit Sub
        Dim nRnd
        Randomize()
        nRnd = CLng(Rnd() * 29252888)
        With Response
            .Clear
            .Expires = 0
            .Write "<br />"
            .Write "<div style=""width:100%; font-size:12px; cursor:pointer;line-height:150%"">"
            .Write "<label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">"
            .Write "<span style=""background-color:820222;color:#FFFFFF;height:23px;font-size:14px;"">〖 Kin_Db_Pager &#25552;&#31034;&#20449;&#24687;  ERROR 〗</span><br />"
            .Write "</label>"
            .Write "<div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">"
            .Write "<span style=""color:#FF0000;"">Description</span> " & Server.HTMLEncode(s) & "<br />"
            .Write "<span style=""color:#FF0000;"">Provider</span> " & sProjectName & "<br />"
            .Write "<span style=""color:#FF0000;"">Version</span> " & sVersion & "<br />"
            .Write "<span style=""color:#FF0000;"">Information</span> Coding By <a href=""http://jorkin.reallydo.com"">Jorkin</a>.<br />"
            .Write "<img width=""0"" height=""0"" src=""http://img.users.51.la/2782986.asp"" style=""display:none"" /></div>"
            .Write "</div>"
            .Write "<br />"
            .End()
        End With
    End Sub

    '//产生分页的SQL语句

    Public Function getSql()
        If Not IsBlank(sSqlString) Then
            getSql = sSqlString
            Exit Function
        End If
        Dim iStart, iEnd
        Call makeCondition()
        iStart = ( iPage - 1 ) * iPageSize
        iEnd = iStart + iPageSize
        Select Case sDbType
            Case "MSSQL"
                getSql = " SELECT " & IIf(bDistinct, "DISTINCT", "") & " " & sFields & " FROM " & TableFormat(sTableName) & " " _
                         & " WHERE [" & sPKey & "] IN ( " _
                         & "   SELECT TOP " & iEnd & " [" & sPKey & "] FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy & " " _
                         & " )"
                If iPage>1 Then
                    getSql = getSql & " AND [" & sPKey & "] NOT IN ( " _
                             & "   SELECT TOP " & iStart & " [" & sPKey & "] FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy & " " _
                             & " )"
                End If
                getSql = getSql & " " & sOrderBy
            Case "MYSQL"
                getSql = "SELECT " & sFields & " FROM " & TableFormat(sTableName)& " " & sCondition & " " & sOrderBy & " LIMIT "&(iPage -1) * iPageSize&"," & iPageSize
            Case "MSSQLPRODUCE"
            Case "ACCESS"
                getSql = "SELECT " & IIf(bDistinct, "DISTINCT ", " ") & " Top " & iPage * iPageSize & " " & sFields & " FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy
            Case Else
                getSql = "SELECT " & sFields & " FROM " & TableFormat(sTableName) & " " & sCondition & " " & sOrderBy
        End Select
    End Function

    '//产生条件字符串

    Private Sub makeCondition()
        If Not IsBlank(sCondition) Then Exit Sub
        If UBound(aCondition)>= 0 Then
            sCondition = " WHERE " & Join(aCondition, " AND ")
        End If
    End Sub

    '//计算记录数

    Private Sub CaculateRecordCount()
        On Error Resume Next
        Dim oRs
        If Not IsBlank(sSqlString) Then
            sSql = "SELECT COUNT(0) FROM (" & sSqlString & ")"
        Else
            Call makeCondition()
            sSql = "SELECT COUNT(0) FROM " & TableFormat(sTableName) & " " & IIf(IsBlank(sCondition), "", sCondition)
        End If
        Set oRs = oConn.Execute( sSql )
        If Err Then
            doError Err.Description
        End If
        iRecordCount = oRs.Fields.Item(0).Value
        Set oRs = Nothing
    End Sub

    '//计算页数

    Private Sub CaculatePageCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If iRecordCount = 0 Then
            iPageCount = 0
            Exit Sub
        End If
        iPageCount = Abs( Int( 0 - (iRecordCount / iPageSize) ) )
    End Sub

    '//设置页码

    Private Function setPage(n)
        iPage = Bint(n)
        If iPage < 1 Then iPage = 1
    End Function

    '//增加条件

    Public Sub AddCondition(s)
        If IsBlank(s) Then Exit Sub
        ReDim Preserve aCondition(UBound(aCondition) + 1)
        aCondition(UBound(aCondition)) = s
    End Sub

    '//判断页面连接

    Private Function ReWrite(n)
        n = Bint(n)
        If Not IsBlank(sRewrite) Then
            ReWrite = Replace(sReWrite, "*", n)
        Else
            ReWrite = sUrl & IIf(n>0, n, "")
        End If
    End Function

    '//数据库表加 []

    Private Function TableFormat(s)
        Select Case iTableKind
            Case 0
                TableFormat = "[" & s & "]"
            Case 1
                TableFormat = " " & s & " "
        End Select
    End Function

    '//按Where In顺序进行排序

    Public Function OrderIn(s, sOrderIn)
        OrderIn = " "
        If Not IsBlank(s) And Not IsBlank(sOrderIn) Then
            sOrderIn = Replace(sOrderIn, " ", "")
            sOrderIn = Replace(sOrderIn, "'", "")
            sOrderIn = "'" & sOrderIn & "'"
            Select Case sDbType
                Case "MYSQL"
                    OrderIn = "FIND_IN_SET(" & s & ", " & sOrderIn & ")"
                Case "ACCESS"
                    OrderIn = "INSTR(','+CStr(" & sOrderIn & ")+',',','+CStr(" & s & ")+',')"
                Case Else
                    OrderIn = "PATINDEX('% ' + CONVERT(nvarchar(820222), " & s & ") + ' %',' ' + CONVERT(nvarchar(820222), Replace(" & sOrderIn & ", ',', ' , ')) + ' ')"
            End Select
        End If
        OrderIn = OrderIn & " "
    End Function

    '//根据数据库连接判断数据库类型

    Private Function GetDbType(o)
        Select Case (o.Provider)
            Case "MSDASQL.1", "SQLOLEDB.1", "SQLOLEDB"
                GetDbType = "MSSQL"
            Case "MSDAORA.1", "OraOLEDB.Oracle"
                GetDbType = "ORACLE"
            Case "Microsoft.Jet.OLEDB.4.0"
                GetDbType = "ACCESS"
        End Select
    End Function

    '//设定分页变量的名称

    Private Function setPageParam(s)
        sQueryString = ""
        For Each x In Request.QueryString
            If x <> sPageParam Then
                For Each y In Request.QueryString(x)
                    sQueryString = "&" & x & "=" & Server.URLEncode(y) & sQueryString
                Next
            End If
        Next
        sUrl = Request.ServerVariables("URL") & "?" & IIf(IsBlank(sQueryString), "", Mid(sQueryString, 2) & "&") & sPageParam & "="
    End Function

    '//函数、方法 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//输入属性 开始

    '//定义连接对象

    Public Property Set ActiveConnection(o)
        Set oConn = o
        sDbType = GetDbType(oConn)
    End Property

    '//连接字符串

    Public Property Let ConnectionString(s)
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionString = s
        oConn.Open()
        sDbType = GetDbType(oConn)
    End Property

    '//定义数据库类型

    Public Property Let DBType(s)
        sDBType = UCase(s)
        Select Case sDBType
            Case "ACCESS", "ACC", "AC"
                sDBType = "ACCESS"
            Case "MSSQL", "SQL"
                sDBType = "MSSQL"
            Case "MYSQL"
                sDBType = "MYSQL"
            Case "ORACLE"
                sDBType = "ORACLE"
            Case "PGSQL"
                sDBType = "PGSQL"
            Case "MSSQLPRODUCE", "MSSQLPR", "MSSQL_PR", "PR"
                sDBType = "MSSQLPRODUCE"
            Case Else
                If TypeName(oConn) = "Connection" Then
                    sDBType = GetDbType(oConn)
                End If
        End Select
    End Property

    '//定义 首页 样式

    Public Property Let FirstPage(s)
        sFirstPage = s
    End Property

    '//定义 上一页 样式

    Public Property Let PreviewPage(s)
        sPreviewPage = s
    End Property

    '//定义 当前页 样式

    Public Property Let CurrentPage(s)
        sCurrentPage = s
    End Property

    '//定义 分页列表页 样式

    Public Property Let ListPage(s)
        sListPage = s
    End Property

    '//定义 下一页 样式

    Public Property Let NextPage(s)
        sNextPage = s
    End Property

    '//定义 末页 样式

    Public Property Let LastPage(s)
        sLastPage = s
    End Property

    '//定义间隔符，默认半角空格

    Public Property Let SpaceMark(s)
        sSpaceMark = s
    End Property

    '//定义 列表前后多加几页

    Public Property Let PagerTop(n)
        iPagerTop = Bint(n)
    End Property

    '//定义查询表名

    Public Property Let TableName(s)
        sTableName = s
        '//如果发现表名包含 ([. ，那么就不要用 []
        If InStr(s, "(")>0 Then iTableKind = 1
        If InStr(s, "[")>0 Then iTableKind = 1
        If InStr(s, ".")>0 Then iTableKind = 1
    End Property

    '//定义需要输出的字段名

    Public Property Let Fields(s)
        sFields = s
    End Property

    '//定义主键

    Public Property Let PKey(s)
        If Not IsBlank(s) Then sPKey = s
    End Property

    '//定义排序规则

    Public Property Let OrderBy(s)
        If Not IsBlank(s) Then sOrderBy = " ORDER BY " & s & " "
    End Property

    '//定义每页的记录条数

    Public Property Let PageSize(s)
        iPageSize = Bint(s)
        iPageSize = IIf(iPageSize<1, 1, iPageSize)
    End Property

    '//定义当前页码

    Public Property Let Page(n)
        setPage Bint(n)
    End Property

    '//定义当前页码(同Property Page)

    Public Property Let AbsolutePage(n)
        setPage Bint(n)
    End Property

    '//自定义查询语句

    Public Property Let Sql(s)
        sSqlString = s
    End Property

    '//是否DISTINCT

    Public Property Let Distinct(b)
        bDistinct = b
    End Property

    '//设定分页变量的名称

    Public Property Let PageParam(s)
        sPageParam = LCase(s)
        If IsBlank(sPageParam) Then sPageParam = "page"
        setPageParam(sPageParam)
    End Property

    '//选择分页的样式,可以后面自己添加新的

    Public Property Let Style(s)
        iStyle = Bint(s)
    End Property

    '//分页列表显示数量

    Public Property Let PagerSize(n)
        iPagerSize = Bint(n)
    End Property

    '//自定义ISAPI_REWRITE路径 * 将被替换为当前页数
    '//使用Javascript时请注意本分页类用双引号引用字符串,请先处理.

    Public Property Let ReWritePath(s)
        sReWrite = s
    End Property

    '//强制TABLE类型

    Public Property Let TableKind(n)
        iTableKind = n
    End Property

    '//自定义分页信息

    Public Property Let PageInfo(s)
        sPageInfo = s
    End Property

    '//定义页面跳转类型

    Public Property Let JumpPageType(s)
        sJumpPageType = UCase(s)
        Select Case sJumpPageType
            Case "INPUT", "SELECT"
            Case Else
                sJumpPageType = "SELECT"
        End Select
    End Property

    '//定义页面跳转链接其他HTML属性

    Public Property Let JumpPageAttr(s)
        sJumpPageAttr = s
    End Property

    '//输入属性 结束
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//输出属性 开始

    '//输出连接语句

    Public Property Get ConnectionString()
        ConnectionString = oConn.ConnectionString
    End Property

    '//输出连接对象

    Public Property Get Conn()
        Set Conn = oConn
    End Property

    '//输出数据库类型

    Public Property Get DBType()
        DBType = sDBType
    End Property

    '//输出查询表名

    Public Property Get TableName()
        TableName = sTableName
    End Property

    '//输出需要输出的字段名

    Public Property Get Fields()
        Fields = sFields
    End Property

    '//输出主键

    Public Property Get PKey()
        PKey = sPKey
    End Property

    '//输出排序规则

    Public Property Get OrderBy()
        OrderBy = sOrderBy
    End Property

    '//取得当前条件下的记录数

    Public Property Get RecordCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        RecordCount = iRecordCount
    End Property

    '//取得每页记录数

    Public Property Get PageSize()
        PageSize = iPageSize
    End Property

    '//取得当前查询的条件

    Public Property Get Condition()
        If IsBlank(sCondition) Then makeCondition()
        Condition = sCondition
    End Property

    '//取得当前页码

    Public Property Get Page()
        Page = iPage
    End Property

    '//取得当前页码

    Public Property Get AbsolutePage()
        AbsolutePage = iPage
    End Property

    '//取得总页数

    Public Property Get PageCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        PageCount = iPageCount
    End Property

    '//取得当前页记录数

    Public Property Get CurrentPageSize()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        CurrentPageSize = IIf(iRecordCount>0, IIf(iPage = iPageCount, iRecordCount - (iPage -1) * iPageSize, iPageSize), 0)
    End Property

    '//得到分页后的记录集

    Public Property Get RecordSet()
        On Error Resume Next
        Select Case sDbType
            Case "MSSQL" '// MSSQL2000
                sSql = getSql()
                Set RecordSet = oConn.Execute( sSql )
            Case "MSSQLPRODUCE" '// SqlServer2000数据库存储过程版, 可使用叶子的SQL。
                Set oRs = Server.CreateObject("ADODB.RecordSet")
                Set oCommand = Server.CreateObject("ADODB.Command")
                oCommand.CommandType = 4
                oCommand.ActiveConnection = oConn
                oCommand.CommandText = "sp_Util_Page"
                oCommand.Parameters(1) = 0
                oCommand.Parameters(2) = iPage
                oCommand.Parameters(3) = iPageSize
                oCommand.Parameters(4) = sPkey
                oCommand.Parameters(5) = sFields
                oCommand.Parameters(6) = sTableName
                oCommand.Parameters(7) = Join(aCondition, " AND ")
                oCommand.Parameters(8) = Mid(sOrderBy, 11)
                oRs.CursorLocation = 3
                oRs.LockType = 1
                oRs.Open oCommand
            Case "MYSQL" 'MYSQL数据库，不会，暂时空着。
                sSql = getSql()
                Set oRs = oConn.Execute(sSql)
            Case Else '其他情况按最原始的ADO方法处理，包括ACCESS。
                sSql = getSql()
                Set RecordSet = Server.CreateObject ("ADODB.RecordSet")
                RecordSet.Open sSql, oConn, 1, 1, &H0001
                RecordSet.PageSize = iPageSize
                If RecordSet.AbsolutePage <> -1 Then
                    iPage = IIf(iPage > RecordSet.PageCount, RecordSet.PageCount, iPage)
                    RecordSet.AbsolutePage = iPage
                End If
        End Select
        If Err Then
            doError Err.Description
            If Not IsBlank(sSql) Then
                Set RecordSet = oConn.Execute( sSql )
                If Err Then doError Err.Description
            Else
                doError Err.Description
            End If
        End If
        Err.Clear()
    End Property

    '//版本信息

    Public Property Get Version()
        Version = sVersion
    End Property

    '//输出页码及记录数等信息

    Public Property Get PageInfo()
        CaculatePageCount()
        PageInfo = Replace(sPageInfo, sRecordCount, iRecordCount)
        PageInfo = Replace(PageInfo, sPageCount, iPageCount)
        PageInfo = Replace(PageInfo, sPage, iPage)
    End Property

    '//输出分页样式

    Public Property Get Style()
        Style = iStyle
    End Property

    '//输出分页变量

    Public Property Get PageParam()
        PageParam = sPageParam
    End Property

    '//输出翻页按钮

    Public Property Get Pager()
        Dim ii, iStart, iEnd
        Pager = ""
        ii = (iPagerSize \ 2)
        iEnd = iPage + ii
        iStart = iPage - (ii + (iPagerSize Mod 2)) + 1
        If iEnd > iPageCount Then
            iEnd = iPageCount
            iStart = iPageCount - iPagerSize + 1
        End If
        If iStart < 1 Then
            iStart = 1
            iEnd = iStart + iPagerSize -1
        End If
        If iEnd > iPageCount Then
            iEnd = iPageCount
        End If

        Select Case iStyle
            Case 0
                If iPageCount>0 Then
                    If iPage>1 Then
                        Pager = Pager & IIf(IsBlank(sFirstPage), "", "<a href=""" & Rewrite(1) & """>" & sFirstPage & "</a>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sPreviewPage), "", "<a href=""" & Rewrite((iPage -1)) & """>" & sPreviewPage & "</a>" & sSpaceMark)
                    Else
                        Pager = Pager & IIf(IsBlank(sFirstPage), "", "<span class=""disabled"">" & sFirstPage & "</span>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sPreviewPage), "", "<span class=""disabled"">" & sPreviewPage & "</span>" & sSpaceMark)
                    End If
                    If iPagerTop > 0 Then
                        If iPagerTop < iStart Then
                            ii = iPagerTop
                        Else
                            ii = iStart - 1
                        End If
                        For i = 1 To ii
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                        If iPagerTop < iStart -1 Then Pager = Pager & "..." & sSpaceMark
                    End If
                    If iPagerSize >0 Then
                        For i = iStart To iEnd
                            If i = iPage Then
                                Pager = Pager & "<span class=""current"">" & Replace(sCurrentPage, "{$Currentpage}", i, 1, -1, 1) & "</span>" & sSpaceMark
                            Else
                                Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                            End If
                        Next
                    End If
                    If iPagerTop > 0 Then
                        If iPageCount - iPagerTop > iEnd Then Pager = Pager & "..." & sSpaceMark
                        If iPageCount - iPagerTop > iEnd Then
                            ii = iPageCount - iPagerTop + 1
                        Else
                            ii = iEnd + 1
                        End If
                        For i = ii To iPageCount
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                    End If
                    If iPageCount>iPage Then
                        Pager = Pager & IIf(IsBlank(sNextPage), "", "<a href=""" & Rewrite(iPage + 1) & """>" & sNextPage & "</a>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sLastPage), "", "<a href=""" & Rewrite(iPageCount) & """>" & sLastPage & "</a>" & sSpaceMark)
                    Else
                        Pager = Pager & IIf(IsBlank(sNextPage), "", "<span class=""disabled"">" & sNextPage & "</span>" & sSpaceMark)
                        Pager = Pager & IIf(IsBlank(sLastPage), "", "<span class=""disabled"">" & sLastPage & "</span>")
                    End If
                End If
            Case 1
                If iPageCount>0 Then
                    If iPage>1 Then
                        Pager = Pager & "<a href=""" & Rewrite(1) & """>" & sFirstPage & "</a>" & sSpaceMark
                        Pager = Pager & "<a href=""" & Rewrite((iPage -1)) & """>" & sPreviewPage & "</a>" & sSpaceMark
                    Else
                        Pager = Pager & sFirstPage & sSpaceMark
                        Pager = Pager & sPreviewPage & sSpaceMark
                    End If
                    If iPagerTop > 0 Then
                        If iPagerTop < iStart Then
                            ii = iPagerTop
                        Else
                            ii = iStart - 1
                        End If
                        For i = 1 To ii
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                        If iPagerTop < iStart -1 Then Pager = Pager & "..." & sSpaceMark
                    End If
                    If iPagerSize >0 Then
                        For i = iStart To iEnd
                            If i = iPage Then
                                Pager = Pager & Replace(sCurrentPage, "{$Currentpage}", i, 1, -1, 1) & sSpaceMark
                            Else
                                Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                            End If
                        Next
                    End If
                    If iPagerTop > 0 Then
                        If iPageCount - iPagerTop > iEnd Then Pager = Pager & "..." & sSpaceMark
                        If iPageCount - iPagerTop > iEnd Then
                            ii = iPageCount - iPagerTop + 1
                        Else
                            ii = iEnd + 1
                        End If
                        For i = ii To iPageCount
                            Pager = Pager & "<a href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Listpage}", i, 1, -1, 1) & "</a>" & sSpaceMark
                        Next
                    End If
                    If iPageCount>iPage Then
                        Pager = Pager & "<a href=""" & Rewrite(iPage + 1) & """>" & sNextPage & "</a>" & sSpaceMark
                        Pager = Pager & "<a href=""" & Rewrite(iPageCount) & """>" & sLastPage & "</a>" & sSpaceMark
                    Else
                        Pager = Pager & sNextPage & sSpaceMark
                        Pager = Pager & sLastPage
                    End If
                End If
            Case Else
                If iPageCount>0 Then
                    If iPage>1 Then
                        Pager = Pager & "<a href=""" & Rewrite(1) & """><font color=""#220282"">[&#39318;&#39029;]</font></a>" & sSpaceMark
                        Pager = Pager & "<a href=""" & Rewrite((iPage -1)) & """><font color=""#220282"">[&#19978;&#19968;&#39029;]</font></a>" & sSpaceMark
                    Else
                        Pager = Pager & "<font color=""#220282"">[&#39318;&#39029;]</font>" & sSpaceMark
                        Pager = Pager & "<font color=""#220282"">[&#19978;&#19968;&#39029;]</font>" & sSpaceMark
                    End If
                    If iPagerTop > 0 Then
                        If iPagerTop < iStart Then
                            ii = iPagerTop
                        Else
                            ii = iStart - 1
                        End If
                        For i = 1 To ii
                            Pager = Pager & "<a href=""" & ReWrite(i) & """><font color=""#220282"">" & Replace("[{$Listpage}]", "{$Listpage}", i, 1, -1, 1) & "</font></a>" & sSpaceMark
                        Next
                        If iPagerTop < iStart -1 Then Pager = Pager & "..." & sSpaceMark
                    End If
                    If iPagerSize >0 Then
                        For i = iStart To iEnd
                            If i = iPage Then
                                Pager = Pager & "<font color=""#820222"">" & Replace("[{$Currentpage}]", "{$Currentpage}", i, 1, -1, 1) & "</font>" & sSpaceMark
                            Else
                                Pager = Pager & "<a href=""" & ReWrite(i) & """><font color=""#220282"">" & Replace("[{$Listpage}]", "{$Listpage}", i, 1, -1, 1) & "</font></a>" & sSpaceMark
                            End If
                        Next
                    End If
                    If iPagerTop > 0 Then
                        If iPageCount - iPagerTop > iEnd Then Pager = Pager & "..." & sSpaceMark
                        If iPageCount - iPagerTop > iEnd Then
                            ii = iPageCount - iPagerTop + 1
                        Else
                            ii = iEnd + 1
                        End If
                        For i = ii To iPageCount
                            Pager = Pager & "<a href=""" & ReWrite(i) & """><font color=""#220282"">" & Replace("[{$Listpage}]", "{$Listpage}", i, 1, -1, 1) & "</font></a>" & sSpaceMark
                        Next
                    End If
                    If iPageCount>iPage Then
                        Pager = Pager & "<a href=""" & Rewrite(iPage + 1) & """><font color=""#220282"">[&#19979;&#19968;&#39029;]</font></a>" & sSpaceMark
                        Pager = Pager & "<a href=""" & Rewrite(iPageCount) & """><font color=""#220282"">[&#23614;&#39029;]</font></a>" & sSpaceMark
                    Else
                        Pager = Pager & "<font color=""#220282"">[&#19979;&#19968;&#39029;]</font>" & sSpaceMark
                        Pager = Pager & "<font color=""#220282"">[&#23614;&#39029;]</font>"
                    End If
                End If
        End Select
    End Property

    '//生成页面跳转

    Public Property Get JumpPage()
        Dim x, sQueryString, aQueryString
        sJumpPage = vbCrLf
        Select Case sJumpPageType
            Case "INPUT"
                sJumpPage = "<input type=""text"" title=""&#35831;&#36755;&#20837;&#25968;&#23383;&#10;&#13;&#22238;&#36710;&#36339;&#36716;"" size=""3"" onKeyDown=""if(event.keyCode==13){if(!isNaN(this.value)){document.location.href=" & IIf(IsBlank(sRewrite), "'" & ReWrite(0) & "'+this.value", Replace("'" & sRewrite & "'", "*", "' + this.value + '")) & "}return false}"" " & sJumpPageAttr & " />"
            Case "SELECT"
                sJumpPage = sJumpPage & "<select onChange=""javascript:window.location.href=this.options[this.selectedIndex].value;"" " & sJumpPageAttr & "　>" & vbCrLf
                iStart = iPage - 50
                iEnd = iPage + 50
                If iEnd > iPageCount Then
                    iEnd = iPageCount
                    iStart = iPageCount - 100 + 1
                End If
                If iStart < 1 Then
                    iStart = 1
                    iEnd = iStart + 100 -1
                End If
                sJumpPage = sJumpPage & "<option value=""javascript:void(0)"">--</option>" & vbCrLf
                For i = iStart To IIf(iEnd > iPageCount, iPageCount, iEnd)
                    sJumpPage = sJumpPage & "<option value=""" & ReWrite(i) & """" & IIf(i = iPage, " selected=""selected"" ", "") & ">"&i&"</option>" & vbCrLf
                Next
                sJumpPage = sJumpPage & "</select>"
            Case Else
                sJumpPage = ""
        End Select
        JumpPage = sJumpPage
    End Property

    '//输出属性 结束
    '//-------------------------------------------------------------------------

End Class
%>
<%
Sub Eg()
    With Response
        .Write("<p style=""text-align:left;padding:22px;border:1px solid #820222;font-size:12px"" id=""eg"">")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// 定义Eg()样例相关变量 如果未使用Option Explicit可省略<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'Dim i, iCols, iColsPercent, iPageSize<br />")
        .Write("'Dim iCurrPage, iRecordCount, iPageCount<br />")
        .Write("'Dim sPageInfo, sPager, sJumpPage<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// Kin_Db_Pager分页类开始<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'OpenConn()<br />")
        .Write("iPageSize = 20<br />")
        .Write("Dim oDbPager<br />")
        .Write("Set oDbPager = New Kin_Db_Pager<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// 进行数据库查询前的相关参数设置<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//指定数据库连接<br />")
        .Write("oDbPager.Connect(oConn) '//方法一(推荐)<br />")
        .Write("'Set oDbPager.ActiveConnection = oConn '//方法二<br />")
        .Write("'oDbPager.ConnectionString = oConn.ConnectionString '//方法三<br />")
        .Write("'//指定数据库类型.默认值:""MSSQL""<br />")
        .Write("'oDbPager.DbType = ""ACCESS""<br />")
        .Write("'//指定目标表 可用临时表""(Select * From [Table]) t""<br />")
        .Write("oDbPager.TableName = ""Kin_Article""<br />")
        .Write("'//选择列 用逗号分隔 默认为*<br />")
        .Write("oDbPager.Fields = ""*""<br />")
        .Write("'//指定该表的主键<br />")
        .Write("oDbPager.PKey = ""Article_ID""<br />")
        .Write("'//指定每页记录集数量<br />")
        .Write("oDbPager.PageSize = iPageSize<br />")
        .Write("'//指定表示页数的URL变量 默认值:""page""<br />")
        .Write("'oDbPager.PageParam = ""page""<br />")
        .Write("'//指定当前页数<br />")
        .Write("oDbPager.Page = Request.QueryString(""page"") '//也可以直接用Request.QueryString(oDbPager.PageParam)<br />")
        .Write("'//指定排序条件<br />")
        .Write("oDbPager.OrderBy = ""Article_ID DESC""<br />")
        .Write("'//添加条件 可多次使用.如果用Or条件需要(条件1 Or 条件2 Or ...)<br />")
        .Write("oDbPager.AddCondition ""Article_Status &gt; 0""<br />")
        .Write("If Day(Date) Mod 2 = 0 Then<br />")
        .Write("&nbsp; &nbsp; oDbPager.AddCondition ""(Article_ID &lt; 104 Or Article_ID &gt; 222)""<br />")
        .Write("End If<br />")
        .Write("'GetCondition """","""",""""<br />")
        .Write("'//输出SQL语句 方便调试<br />")
        .Write("'Response.Write(oDbPager.GetSql()) : Response.Flush()<br />")
        .Write("Set oRs = oDbPager.Recordset<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// 对该记录集的分页样式及模板进行设置(不设置则使用默认样式)<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//选择 分页链接 输出的样式<br />")
        .Write("'//为0: 可以使用样式表对分页链接进行美化(http://jorkin.reallydo.com/Kin_Db_Pager/?page=10)<br />")
        .Write("'//为1: 可使用&lt;font&gt;等HTML代码进行颜色设置<br />")
        .Write("'oDbPager.Style = 0<br />")
        .Write("'//定义 首页/上一页/下一页/末页 链接样式(支持HTML)<br />")
        .Write("'oDbPager.FirstPage = ""&amp;lt;&amp;lt;""<br />")
        .Write("'oDbPager.PreviewPage = ""&amp;lt;""<br />")
        .Write("'oDbPager.NextPage = ""&amp;gt;""<br />")
        .Write("'oDbPager.LastPage = ""&amp;gt;&amp;gt;""<br />")
        .Write("'//定义 当前页/列表页 链接样式 {$CurrentPage}{$ListPage}将被替换成 当前页/列表页 的数字<br />")
        .Write("'oDbPager.CurrentPage = ""{$CurrentPage}""<br />")
        .Write("'oDbPager.ListPage = ""{$ListPage}""<br />")
        .Write("'//定义分页列表前后要显示几个链接 如12...456...78 默认为0<br />")
        .Write("'oDbPager.PagerTop = 2<br />")
        .Write("'//定义分页列表最大数量 默认为7<br />")
        .Write("'oDbPager.PagerSize = 5<br />")
        .Write("'//定义记录集综合信息<br />")
        .Write("'oDbPager.PageInfo = &quot;共有 {$Kin_RecordCount} 记录 页次:{$Kin_Page}/{$Kin_PageCount}&quot;<br />")
        .Write("'//自定义ISAPI_REWRITE路径 * 号 将被替换为当前页数<br />")
        .Write("'oDbPager.RewritePath = ""Article/*.html""<br />")
        .Write("'//定义跳转列表为&lt;INPUT&gt;文本框 默认为""SELECT""<br />")
        .Write("'oDbPager.JumpPageType = ""INPUT""<br />")
        .Write("'//定义页面跳的SELECT/INPUT的样式(HTML代码)<br />")
        .Write("'oDbPager.JumpPageAttr = ""class=""""reallydo"""" style=""""color:#820222""""""<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// 获取所需要变量以便进行输出<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//获取当前页码<br />")
        .Write("'iCurrPage = oDbPager.Page<br />")
        .Write("'//获取记录集数量<br />")
        .Write("'iRecordCount = oDbPager.RecordCount<br />")
        .Write("'//获取页面总计数量<br />")
        .Write("'iPageCount = oDbPager.PageCount<br />")
        .Write("'//获取记录集信息<br />")
        .Write("sPageInfo = oDbPager.PageInfo<br />")
        .Write("'//获取分页信息<br />")
        .Write("sPager = oDbPager.Pager<br />")
        .Write("'//获取跳转列表<br />")
        .Write("sJumpPage = oDbPager.JumpPage<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// 例子: 动态输出M行N列, 多行多列, 循环行列, 循环TABLE<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//初始化i准备循环<br />")
        .Write("i = 0<br />")
        .Write("'//定义一行最多有几列(正整数)<br />")
        .Write("iCols = 4<br />")
        .Write("iColsPercent = FormatPercent(1 / iCols, 0)<br />")
        .Write("'//输出TABLE表头<br />")
        .Write("Response.Write(""&lt;table width=""""100%"""" border=""""0"""" cellspacing=""""1"""" cellpadding=""""2"""" bgcolor=""""#000000""""&gt;&lt;tr&gt;"")<br />")
        .Write("'//方法一:记录集循环开始<br />")
        .Write("Do While Not oRs.EOF<br />")
        .Write(" &nbsp; &nbsp;'//每行例满了就加一个新行<br />")
        .Write(" &nbsp; &nbsp;If i &gt; 0 And i Mod iCols = 0 Then Response.Write(""&lt;/tr&gt;&lt;tr&gt;"")<br />")
        .Write(" &nbsp; &nbsp;i = i + 1<br />")
        .Write(" &nbsp; &nbsp;Response.Write(""&lt;td width="""""" & iColsPercent & """""" bgcolor=""""#CCE8CF""""&gt;&lt;font color=""""#000000""""&gt;"" & Server.HTMLEncode(oRs(2) & """") & ""&lt;/font&gt;&lt;/td&gt;"")<br />")
        .Write(" &nbsp; &nbsp;oRs.MoveNext<br />")
        .Write("Loop<br />")
        .Write("'//方法二:游标循环开始<br />")
        .Write("'//获取当前页面总记录数量<br />")
        .Write("'iCurrentPageSize = oDbPager.CurrentPageSize<br />")
        .Write("'For i = 0 To iCurrentPageSize - 1<br />")
        .Write("' &nbsp; &nbsp;'//每行例满了就加一个新行<br />")
        .Write("' &nbsp; &nbsp;If i &gt; 0 And i Mod iCols = 0 Then Response.Write(""&lt;/tr&gt;&lt;tr&gt;"")<br />")
        .Write("' &nbsp; &nbsp;Response.Write(""&lt;td width="""""" & iColsPercent & """""" bgcolor=""""#CCE8CF""""&gt;&lt;font color=""""#000000""""&gt;"" & Server.HTMLEncode(oRs(2) & """") & ""&lt;/font&gt;&lt;/td&gt;"")<br />")
        .Write("' &nbsp; &nbsp;oRs.MoveNext<br />")
        .Write("'Next<br />")
        .Write("'//循环结束 开始补空缺的列<br />")
        .Write("Do While i &lt; iPageSize<br />")
        .Write(" &nbsp; &nbsp;'//以下两个条件二选一<br />")
        .Write(" &nbsp; &nbsp;If i Mod iCols = 0 Then<br />")
        .Write(" &nbsp; &nbsp; &nbsp; &nbsp;Response.Write(""&lt;/tr&gt;&lt;tr&gt;"") '//如果要补满整个表格就继续输出&lt;tr&gt;&lt;/tr&gt;<br />")
        .Write(" &nbsp; &nbsp; &nbsp; &nbsp;'Exit Do '//如果只补满最后一行就直接结束<br />")
        .Write(" &nbsp; &nbsp;End If<br />")
        .Write(" &nbsp; &nbsp;i = i + 1<br />")
        .Write(" &nbsp; &nbsp;Response.Write(""&lt;td width=""""""&FormatPercent(1 / iCols, 0)&"""""" bgcolor=""""#CCCCCC""""&gt;&amp;nbsp;&lt;/td&gt;"")<br />")
        .Write("Loop<br />")
        .Write("'//输出分页信息/样式/TABLE尾<br />")
        .Write("Response.Write(""&lt;/tr&gt;&lt;tr&gt;&lt;td colspan="""""" & iCols & """""" bgcolor=""""#CCE8CF""""&gt;&lt;div class=""""kindbpager""""&gt;"" & sPager & "" 跳至: "" & sJumpPage & "" 页&lt;/div&gt;"" & sPageInfo & ""&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;"")<br />")
        .Write("oRs.Close<br />")
        .Write("Set oDbPager = Nothing<br />")
        .Write("</p>")
        .Flush()
    End With
End Sub

Sub GetCondition(valChoose, valOperator, valKeyWord)
    If IsBlank(BStr(valChoose)) Then valChoose = "Choose"
    If IsBlank(BStr(valOperator)) Then valOperator = "Operator"
    If IsBlank(BStr(valKeyWord)) Then valKeyWord = "KeyWord"
    Dim aChoose, aOperator, aKeyWord, x
    Set aChoose = Request(valChoose)
    Set aOperator = Request(valOperator)
    Set aKeyWord = Request(valKeyWord)
    If aChoose.Count = aOperator.Count And aOperator.Count = aKeyWord.Count Then
        For x = 1 To aChoose.Count
            If aChoose(x)<>"" And aChoose(x)<>"" And aKeyWord(x)<>"" Then
                Select Case aOperator(x)
                    Case "<", "=", ">", "<=", ">=", "<>", "!=", "!<", "!>"
                        If InStr(aChoose(x), "[int]")>0 Then
                            oDbPager.AddCondition Str4Sql(Replace(aChoose(x), "[int]", "")) & " " & aOperator(x) & " " & Bint(aKeyWord(x)) & ""
                        Else
                            oDbPager.AddCondition Str4Sql(aChoose(x)) & " " & aOperator(x) & " '" & Str4Sql(BStr(aKeyWord(x))) & "'"
                        End If
                    Case Else
                        If InStr(aChoose(x), "[int]")>0 Then
                            oDbPager.AddCondition " " & Str4Sql(Replace(aChoose(x), "[int]", "")) & " like '%" & Bint(aKeyWord(x)) & "%'"
                        Else
                            If LCase(aKeyWord(x)) = "null" Then
                                oDbPager.AddCondition "(" & Str4Sql(aChoose(x)) & " is null Or " & Str4Sql(aChoose(x)) & " = '')"
                            ElseIf InStr(LCase(aChoose(x)), "not")>0 Then
                                oDbPager.AddCondition Str4Sql(aChoose(x)) & " not like '%" & Str4Like(BStr(aKeyWord(x))) & "%'"
                            Else
                                oDbPager.AddCondition Str4Sql(aChoose(x)) & " like '" & Replace(Replace(Replace(Str4Like(BStr(aKeyWord(x))), "*", "%"), "?", "_"), "？", "_") & "'"
                            End If
                        End If
                End Select
            End If
        Next
    End If
End Sub
%>