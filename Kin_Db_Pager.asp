<%
'/*------------------------------------------------ Jorkin �Զ����� ��ҳ�Ż�����
' ********************************************************************* ˵    ��
' * ��Դ: KinJAVA��־ (http://jorkin.reallydo.com/article.asp?id=534)
' * ������: 2009-03-22
' * ��ǰ�汾: Ver: 1.09
' ********************************************************************* ��ɫ����
' * �з���� Eg() ʵ��,����Ҫ��סÿ������������,Ϊ�������ṩ����.
' * ISAPI_REWRITE ���ܿ�������ʵ�־�̬(α��̬)��ҳ�Լ�ajax��ҳ.
' ********************************************************************* ������ʷ
' * 2009-03-22
' *   ����qsliuliu�����Զ���PageParam���ҳ����BUG��
' * 2009-03-19
' *   ��Ҫ���£��Ż�������ٶȡ�
' * 2009-02-14
' *   �Ż�ReWrite()��������������Ч�ʡ�
' *   �����º���GetCondition()��������������������������ҪJorkin_Function.asp�⡣
' *   ����Ϊ��GetCondition("���ֶ���", "���Ƚ������", "���ؼ���")��
' *   �Ƚ��������ѡ(<, =, >, <=, >=, <>, !=, !<, !>, like, not like)���ؼ��ֿ���?��ʾ���ַ�����*��ʾ���������ַ���
' *   ���棡����ǰ̨ҳ�����ã�����¶���ݿ��ֶ����ƣ��ɲο�GetCondition()�����޸ġ�
' * 2009-02-11
' *   С��С�ģ���ǿ�˼����жϣ�����һЩ˵�������ô洢���̵�������Ҷ�ӷ�ҳ���sp_Util_Page.sql��������ȫ���ݡ�
' * 2008-11-26
' *   �������ݿ������Զ��ж����ݿ����͡�
' * 2008-09-09
' *   ���� Eg() ʵ��BUG��
' *   ��������һֱ��β��Select Where In�����ܣ�����δȫ����ɡ�(����������ΪKin_Db_Pager.OrderIn(�ֶ���,����))
' *   �����˼��������µ�ƴд���� (-_-#)
' *   �������Զ��巭ҳ��ʽʱ�趨��ֵ���жϡ�
' * 2008-08-28
' *   ���� Connect() �����������ݿ����ӣ��� ActiveConnection �� ConnectionString ����ȫ��Ч��
' *   ����Bug: ���������תINPUT����һ��FORM��س�ʱ��������ύ��������
' *   ����Bug: ������תINPUT��֧���Զ���ISAPI_REWRITE·����
' *   ����Bug: ʹ���Զ���SQL���ʱ��ҳ����
' *   ��д Eg() ����,ʹ������ױ���⡣
' *   ɾ���������ô���.
' * Ver: 1.03֮ǰ
' *   һ�д��뼴��ʵ�ְ���,����Ҫ��ס���е������趨��
' *   ����ʹ�� Eg() �鿴���ɵĴ���,����ȫѡ���Ʒ���ASP������ڼ�Ϊ����ҳ��Ĳ���ģ�塣
' ********************************************************************* ��    л
' * ��л���´��ķ�ҳ��˼�뼰����:
' * Sunrise_Chen (http://www.ccopus.com)
' * ���� (http://www.54caizi.org)
' * ���� (http://www.fonshen.com)
' * Ҷ�� (http://www.yeeh.org)
'*/-----------------------------------------------------------------------------

Class Kin_Db_Pager

    '//-------------------------------------------------------------------------
    '// ������� ��ʼ

    Private oConn '//���Ӷ���
    Private sDbType '//���ݿ�����
    Private sTableName '//����
    Private sPKey '//����
    Private sFields '//������ֶ���
    Private sOrderBy '//�����ַ���
    Private sSql '//��ǰ�Ĳ�ѯ���
    Private sSqlString '//�Զ���Sql���
    Private aCondition() '//��ѯ����(����)
    Private sCondition '//��ѯ����(�ַ���)
    Private iPage '//��ǰҳ��
    Private iPageSize '//ÿҳ��¼��
    Private iPageCount '//��ҳ��
    Private iRecordCount '//��ǰ��ѯ�����µļ�¼��
    Private sPage '//��ǰҳ �滻�ַ���
    Private sPageCount '//��ҳ�� �滻�ַ���
    Private sRecordCount '//��ǰ��ѯ�����µļ�¼�� �滻�ַ���
    Private sProjectName '//��Ŀ��
    Private sVersion '//�汾��
    Private bShowError '//�Ƿ���ʾ������Ϣ
    Private bDistinct '//�Ƿ���ʾΨһ��¼
    Private sPageInfo '//��¼����ҳ�����Ϣ
    Private sPageParam '//page��������
    Private iStyle '//��ҳ����ʽ
    Private iPagerSize '//��ҳ��ť����ֵ
    Private iCurrentPageSize '//��ǰҳ���¼����
    Private sReWrite '//��ISAP REWRITE����·��,����Javascript����ʵ��AJAX��ҳ
    Private iTableKind '//�������, �Ƿ���Ҫǿ�Ƽ� [ ]
    Private sFirstPage '//��ҳ���� ��ʽ
    Private sPreviewPage '//��һҳ���� ��ʽ
    Private sCurrentPage '//��ǰҳ���� ��ʽ
    Private sListPage '//��ҳ�б����� ��ʽ
    Private sNextPage '//��һҳ���� ��ʽ
    Private sLastPage '//ĩҳ���� ��ʽ
    Private iPagerTop '//��ҳ�б�ͷβ����
    Private iPagerGroup '//����ҳ��Ϊһ��
    Private sJumpPage '//��ҳ��ת����
    Private sJumpPageType '//��ҳ��ת����(��ѡSELECT��INPUT)
    Private sJumpPageAttr '//��ҳ��ת����HTML����
    Private sUrl, sQueryString, x, y
    Private sSpaceMark '//����֮ǰ�����

    '//������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//�¼�������: ���ʼ���¼� ��ʼ

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

    '//������¼�

    Private Sub Class_Terminate()
        Set oConn = Nothing
    End Sub

    '//�¼�������: ���ʼ���¼� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//���������� ��ʼ

    '����:ASP���IIF
    '��Դ:http://jorkin.reallydo.com/article.asp?id=26

    Private Function IIf(bExp1, sVal1, sVal2)
        If (bExp1) Then
            IIf = sVal1
        Else
            IIf = sVal2
        End If
    End Function

    '����:ֻȡ����
    '��Դ:http://jorkin.reallydo.com/article.asp?id=395

    Private Function Bint(sValue)
        On Error Resume Next
        Bint = 0
        Bint = Fix(CDbl(sValue))
    End Function

    '����:�ж��Ƿ��ǿ�ֵ
    '��Դ:http://jorkin.reallydo.com/article.asp?id=386

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

    '//������ݿ������Ƿ����

    Public Function Connect(o)
        If TypeName(o) <> "Connection" Then
            doError "��Ч�����ݿ����ӡ�"
        Else
            If o.State = 1 Then
                Set oConn = o
                sDbType = GetDbType(oConn)
            Else
                doError "���ݿ������ѹرա�"
            End If
        End If
    End Function

    '//���������Ϣ

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
            .Write "<span style=""background-color:820222;color:#FFFFFF;height:23px;font-size:14px;"">�� Kin_Db_Pager &#25552;&#31034;&#20449;&#24687;  ERROR ��</span><br />"
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

    '//������ҳ��SQL���

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

    '//���������ַ���

    Private Sub makeCondition()
        If Not IsBlank(sCondition) Then Exit Sub
        If UBound(aCondition)>= 0 Then
            sCondition = " WHERE " & Join(aCondition, " AND ")
        End If
    End Sub

    '//�����¼��

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

    '//����ҳ��

    Private Sub CaculatePageCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If iRecordCount = 0 Then
            iPageCount = 0
            Exit Sub
        End If
        iPageCount = Abs( Int( 0 - (iRecordCount / iPageSize) ) )
    End Sub

    '//����ҳ��

    Private Function setPage(n)
        iPage = Bint(n)
        If iPage < 1 Then iPage = 1
    End Function

    '//��������

    Public Sub AddCondition(s)
        If IsBlank(s) Then Exit Sub
        ReDim Preserve aCondition(UBound(aCondition) + 1)
        aCondition(UBound(aCondition)) = s
    End Sub

    '//�ж�ҳ������

    Private Function ReWrite(n)
        n = Bint(n)
        If Not IsBlank(sRewrite) Then
            ReWrite = Replace(sReWrite, "*", n)
        Else
            ReWrite = sUrl & IIf(n>0, n, "")
        End If
    End Function

    '//���ݿ��� []

    Private Function TableFormat(s)
        Select Case iTableKind
            Case 0
                TableFormat = "[" & s & "]"
            Case 1
                TableFormat = " " & s & " "
        End Select
    End Function

    '//��Where In˳���������

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

    '//�������ݿ������ж����ݿ�����

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

    '//�趨��ҳ����������

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

    '//���������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//�������� ��ʼ

    '//�������Ӷ���

    Public Property Set ActiveConnection(o)
        Set oConn = o
        sDbType = GetDbType(oConn)
    End Property

    '//�����ַ���

    Public Property Let ConnectionString(s)
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionString = s
        oConn.Open()
        sDbType = GetDbType(oConn)
    End Property

    '//�������ݿ�����

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

    '//���� ��ҳ ��ʽ

    Public Property Let FirstPage(s)
        sFirstPage = s
    End Property

    '//���� ��һҳ ��ʽ

    Public Property Let PreviewPage(s)
        sPreviewPage = s
    End Property

    '//���� ��ǰҳ ��ʽ

    Public Property Let CurrentPage(s)
        sCurrentPage = s
    End Property

    '//���� ��ҳ�б�ҳ ��ʽ

    Public Property Let ListPage(s)
        sListPage = s
    End Property

    '//���� ��һҳ ��ʽ

    Public Property Let NextPage(s)
        sNextPage = s
    End Property

    '//���� ĩҳ ��ʽ

    Public Property Let LastPage(s)
        sLastPage = s
    End Property

    '//����������Ĭ�ϰ�ǿո�

    Public Property Let SpaceMark(s)
        sSpaceMark = s
    End Property

    '//���� �б�ǰ���Ӽ�ҳ

    Public Property Let PagerTop(n)
        iPagerTop = Bint(n)
    End Property

    '//�����ѯ����

    Public Property Let TableName(s)
        sTableName = s
        '//������ֱ������� ([. ����ô�Ͳ�Ҫ�� []
        If InStr(s, "(")>0 Then iTableKind = 1
        If InStr(s, "[")>0 Then iTableKind = 1
        If InStr(s, ".")>0 Then iTableKind = 1
    End Property

    '//������Ҫ������ֶ���

    Public Property Let Fields(s)
        sFields = s
    End Property

    '//��������

    Public Property Let PKey(s)
        If Not IsBlank(s) Then sPKey = s
    End Property

    '//�����������

    Public Property Let OrderBy(s)
        If Not IsBlank(s) Then sOrderBy = " ORDER BY " & s & " "
    End Property

    '//����ÿҳ�ļ�¼����

    Public Property Let PageSize(s)
        iPageSize = Bint(s)
        iPageSize = IIf(iPageSize<1, 1, iPageSize)
    End Property

    '//���嵱ǰҳ��

    Public Property Let Page(n)
        setPage Bint(n)
    End Property

    '//���嵱ǰҳ��(ͬProperty Page)

    Public Property Let AbsolutePage(n)
        setPage Bint(n)
    End Property

    '//�Զ����ѯ���

    Public Property Let Sql(s)
        sSqlString = s
    End Property

    '//�Ƿ�DISTINCT

    Public Property Let Distinct(b)
        bDistinct = b
    End Property

    '//�趨��ҳ����������

    Public Property Let PageParam(s)
        sPageParam = LCase(s)
        If IsBlank(sPageParam) Then sPageParam = "page"
        setPageParam(sPageParam)
    End Property

    '//ѡ���ҳ����ʽ,���Ժ����Լ�����µ�

    Public Property Let Style(s)
        iStyle = Bint(s)
    End Property

    '//��ҳ�б���ʾ����

    Public Property Let PagerSize(n)
        iPagerSize = Bint(n)
    End Property

    '//�Զ���ISAPI_REWRITE·�� * �����滻Ϊ��ǰҳ��
    '//ʹ��Javascriptʱ��ע�Ȿ��ҳ����˫���������ַ���,���ȴ���.

    Public Property Let ReWritePath(s)
        sReWrite = s
    End Property

    '//ǿ��TABLE����

    Public Property Let TableKind(n)
        iTableKind = n
    End Property

    '//�Զ����ҳ��Ϣ

    Public Property Let PageInfo(s)
        sPageInfo = s
    End Property

    '//����ҳ����ת����

    Public Property Let JumpPageType(s)
        sJumpPageType = UCase(s)
        Select Case sJumpPageType
            Case "INPUT", "SELECT"
            Case Else
                sJumpPageType = "SELECT"
        End Select
    End Property

    '//����ҳ����ת��������HTML����

    Public Property Let JumpPageAttr(s)
        sJumpPageAttr = s
    End Property

    '//�������� ����
    '//-------------------------------------------------------------------------

    '//-------------------------------------------------------------------------
    '//������� ��ʼ

    '//����������

    Public Property Get ConnectionString()
        ConnectionString = oConn.ConnectionString
    End Property

    '//������Ӷ���

    Public Property Get Conn()
        Set Conn = oConn
    End Property

    '//������ݿ�����

    Public Property Get DBType()
        DBType = sDBType
    End Property

    '//�����ѯ����

    Public Property Get TableName()
        TableName = sTableName
    End Property

    '//�����Ҫ������ֶ���

    Public Property Get Fields()
        Fields = sFields
    End Property

    '//�������

    Public Property Get PKey()
        PKey = sPKey
    End Property

    '//����������

    Public Property Get OrderBy()
        OrderBy = sOrderBy
    End Property

    '//ȡ�õ�ǰ�����µļ�¼��

    Public Property Get RecordCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        RecordCount = iRecordCount
    End Property

    '//ȡ��ÿҳ��¼��

    Public Property Get PageSize()
        PageSize = iPageSize
    End Property

    '//ȡ�õ�ǰ��ѯ������

    Public Property Get Condition()
        If IsBlank(sCondition) Then makeCondition()
        Condition = sCondition
    End Property

    '//ȡ�õ�ǰҳ��

    Public Property Get Page()
        Page = iPage
    End Property

    '//ȡ�õ�ǰҳ��

    Public Property Get AbsolutePage()
        AbsolutePage = iPage
    End Property

    '//ȡ����ҳ��

    Public Property Get PageCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        PageCount = iPageCount
    End Property

    '//ȡ�õ�ǰҳ��¼��

    Public Property Get CurrentPageSize()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        If IsNull(iPageCount) Then CaculatePageCount()
        CurrentPageSize = IIf(iRecordCount>0, IIf(iPage = iPageCount, iRecordCount - (iPage -1) * iPageSize, iPageSize), 0)
    End Property

    '//�õ���ҳ��ļ�¼��

    Public Property Get RecordSet()
        On Error Resume Next
        Select Case sDbType
            Case "MSSQL" '// MSSQL2000
                sSql = getSql()
                Set RecordSet = oConn.Execute( sSql )
            Case "MSSQLPRODUCE" '// SqlServer2000���ݿ�洢���̰�, ��ʹ��Ҷ�ӵ�SQL��
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
            Case "MYSQL" 'MYSQL���ݿ⣬���ᣬ��ʱ���š�
                sSql = getSql()
                Set oRs = oConn.Execute(sSql)
            Case Else '�����������ԭʼ��ADO������������ACCESS��
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

    '//�汾��Ϣ

    Public Property Get Version()
        Version = sVersion
    End Property

    '//���ҳ�뼰��¼������Ϣ

    Public Property Get PageInfo()
        CaculatePageCount()
        PageInfo = Replace(sPageInfo, sRecordCount, iRecordCount)
        PageInfo = Replace(PageInfo, sPageCount, iPageCount)
        PageInfo = Replace(PageInfo, sPage, iPage)
    End Property

    '//�����ҳ��ʽ

    Public Property Get Style()
        Style = iStyle
    End Property

    '//�����ҳ����

    Public Property Get PageParam()
        PageParam = sPageParam
    End Property

    '//�����ҳ��ť

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

    '//����ҳ����ת

    Public Property Get JumpPage()
        Dim x, sQueryString, aQueryString
        sJumpPage = vbCrLf
        Select Case sJumpPageType
            Case "INPUT"
                sJumpPage = "<input type=""text"" title=""&#35831;&#36755;&#20837;&#25968;&#23383;&#10;&#13;&#22238;&#36710;&#36339;&#36716;"" size=""3"" onKeyDown=""if(event.keyCode==13){if(!isNaN(this.value)){document.location.href=" & IIf(IsBlank(sRewrite), "'" & ReWrite(0) & "'+this.value", Replace("'" & sRewrite & "'", "*", "' + this.value + '")) & "}return false}"" " & sJumpPageAttr & " />"
            Case "SELECT"
                sJumpPage = sJumpPage & "<select onChange=""javascript:window.location.href=this.options[this.selectedIndex].value;"" " & sJumpPageAttr & "��>" & vbCrLf
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

    '//������� ����
    '//-------------------------------------------------------------------------

End Class
%>
<%
Sub Eg()
    With Response
        .Write("<p style=""text-align:left;padding:22px;border:1px solid #820222;font-size:12px"" id=""eg"">")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// ����Eg()������ر��� ���δʹ��Option Explicit��ʡ��<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'Dim i, iCols, iColsPercent, iPageSize<br />")
        .Write("'Dim iCurrPage, iRecordCount, iPageCount<br />")
        .Write("'Dim sPageInfo, sPager, sJumpPage<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// Kin_Db_Pager��ҳ�࿪ʼ<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'OpenConn()<br />")
        .Write("iPageSize = 20<br />")
        .Write("Dim oDbPager<br />")
        .Write("Set oDbPager = New Kin_Db_Pager<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// �������ݿ��ѯǰ����ز�������<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//ָ�����ݿ�����<br />")
        .Write("oDbPager.Connect(oConn) '//����һ(�Ƽ�)<br />")
        .Write("'Set oDbPager.ActiveConnection = oConn '//������<br />")
        .Write("'oDbPager.ConnectionString = oConn.ConnectionString '//������<br />")
        .Write("'//ָ�����ݿ�����.Ĭ��ֵ:""MSSQL""<br />")
        .Write("'oDbPager.DbType = ""ACCESS""<br />")
        .Write("'//ָ��Ŀ��� ������ʱ��""(Select * From [Table]) t""<br />")
        .Write("oDbPager.TableName = ""Kin_Article""<br />")
        .Write("'//ѡ���� �ö��ŷָ� Ĭ��Ϊ*<br />")
        .Write("oDbPager.Fields = ""*""<br />")
        .Write("'//ָ���ñ������<br />")
        .Write("oDbPager.PKey = ""Article_ID""<br />")
        .Write("'//ָ��ÿҳ��¼������<br />")
        .Write("oDbPager.PageSize = iPageSize<br />")
        .Write("'//ָ����ʾҳ����URL���� Ĭ��ֵ:""page""<br />")
        .Write("'oDbPager.PageParam = ""page""<br />")
        .Write("'//ָ����ǰҳ��<br />")
        .Write("oDbPager.Page = Request.QueryString(""page"") '//Ҳ����ֱ����Request.QueryString(oDbPager.PageParam)<br />")
        .Write("'//ָ����������<br />")
        .Write("oDbPager.OrderBy = ""Article_ID DESC""<br />")
        .Write("'//������� �ɶ��ʹ��.�����Or������Ҫ(����1 Or ����2 Or ...)<br />")
        .Write("oDbPager.AddCondition ""Article_Status &gt; 0""<br />")
        .Write("If Day(Date) Mod 2 = 0 Then<br />")
        .Write("&nbsp; &nbsp; oDbPager.AddCondition ""(Article_ID &lt; 104 Or Article_ID &gt; 222)""<br />")
        .Write("End If<br />")
        .Write("'GetCondition """","""",""""<br />")
        .Write("'//���SQL��� �������<br />")
        .Write("'Response.Write(oDbPager.GetSql()) : Response.Flush()<br />")
        .Write("Set oRs = oDbPager.Recordset<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// �Ըü�¼���ķ�ҳ��ʽ��ģ���������(��������ʹ��Ĭ����ʽ)<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//ѡ�� ��ҳ���� �������ʽ<br />")
        .Write("'//Ϊ0: ����ʹ����ʽ��Է�ҳ���ӽ�������(http://jorkin.reallydo.com/Kin_Db_Pager/?page=10)<br />")
        .Write("'//Ϊ1: ��ʹ��&lt;font&gt;��HTML���������ɫ����<br />")
        .Write("'oDbPager.Style = 0<br />")
        .Write("'//���� ��ҳ/��һҳ/��һҳ/ĩҳ ������ʽ(֧��HTML)<br />")
        .Write("'oDbPager.FirstPage = ""&amp;lt;&amp;lt;""<br />")
        .Write("'oDbPager.PreviewPage = ""&amp;lt;""<br />")
        .Write("'oDbPager.NextPage = ""&amp;gt;""<br />")
        .Write("'oDbPager.LastPage = ""&amp;gt;&amp;gt;""<br />")
        .Write("'//���� ��ǰҳ/�б�ҳ ������ʽ {$CurrentPage}{$ListPage}�����滻�� ��ǰҳ/�б�ҳ ������<br />")
        .Write("'oDbPager.CurrentPage = ""{$CurrentPage}""<br />")
        .Write("'oDbPager.ListPage = ""{$ListPage}""<br />")
        .Write("'//�����ҳ�б�ǰ��Ҫ��ʾ�������� ��12...456...78 Ĭ��Ϊ0<br />")
        .Write("'oDbPager.PagerTop = 2<br />")
        .Write("'//�����ҳ�б�������� Ĭ��Ϊ7<br />")
        .Write("'oDbPager.PagerSize = 5<br />")
        .Write("'//�����¼���ۺ���Ϣ<br />")
        .Write("'oDbPager.PageInfo = &quot;���� {$Kin_RecordCount} ��¼ ҳ��:{$Kin_Page}/{$Kin_PageCount}&quot;<br />")
        .Write("'//�Զ���ISAPI_REWRITE·�� * �� �����滻Ϊ��ǰҳ��<br />")
        .Write("'oDbPager.RewritePath = ""Article/*.html""<br />")
        .Write("'//������ת�б�Ϊ&lt;INPUT&gt;�ı��� Ĭ��Ϊ""SELECT""<br />")
        .Write("'oDbPager.JumpPageType = ""INPUT""<br />")
        .Write("'//����ҳ������SELECT/INPUT����ʽ(HTML����)<br />")
        .Write("'oDbPager.JumpPageAttr = ""class=""""reallydo"""" style=""""color:#820222""""""<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// ��ȡ����Ҫ�����Ա�������<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//��ȡ��ǰҳ��<br />")
        .Write("'iCurrPage = oDbPager.Page<br />")
        .Write("'//��ȡ��¼������<br />")
        .Write("'iRecordCount = oDbPager.RecordCount<br />")
        .Write("'//��ȡҳ���ܼ�����<br />")
        .Write("'iPageCount = oDbPager.PageCount<br />")
        .Write("'//��ȡ��¼����Ϣ<br />")
        .Write("sPageInfo = oDbPager.PageInfo<br />")
        .Write("'//��ȡ��ҳ��Ϣ<br />")
        .Write("sPager = oDbPager.Pager<br />")
        .Write("'//��ȡ��ת�б�<br />")
        .Write("sJumpPage = oDbPager.JumpPage<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'// ����: ��̬���M��N��, ���ж���, ѭ������, ѭ��TABLE<br />")
        .Write("'//-----------------------------------------------------------------------------<br />")
        .Write("'//��ʼ��i׼��ѭ��<br />")
        .Write("i = 0<br />")
        .Write("'//����һ������м���(������)<br />")
        .Write("iCols = 4<br />")
        .Write("iColsPercent = FormatPercent(1 / iCols, 0)<br />")
        .Write("'//���TABLE��ͷ<br />")
        .Write("Response.Write(""&lt;table width=""""100%"""" border=""""0"""" cellspacing=""""1"""" cellpadding=""""2"""" bgcolor=""""#000000""""&gt;&lt;tr&gt;"")<br />")
        .Write("'//����һ:��¼��ѭ����ʼ<br />")
        .Write("Do While Not oRs.EOF<br />")
        .Write(" &nbsp; &nbsp;'//ÿ�������˾ͼ�һ������<br />")
        .Write(" &nbsp; &nbsp;If i &gt; 0 And i Mod iCols = 0 Then Response.Write(""&lt;/tr&gt;&lt;tr&gt;"")<br />")
        .Write(" &nbsp; &nbsp;i = i + 1<br />")
        .Write(" &nbsp; &nbsp;Response.Write(""&lt;td width="""""" & iColsPercent & """""" bgcolor=""""#CCE8CF""""&gt;&lt;font color=""""#000000""""&gt;"" & Server.HTMLEncode(oRs(2) & """") & ""&lt;/font&gt;&lt;/td&gt;"")<br />")
        .Write(" &nbsp; &nbsp;oRs.MoveNext<br />")
        .Write("Loop<br />")
        .Write("'//������:�α�ѭ����ʼ<br />")
        .Write("'//��ȡ��ǰҳ���ܼ�¼����<br />")
        .Write("'iCurrentPageSize = oDbPager.CurrentPageSize<br />")
        .Write("'For i = 0 To iCurrentPageSize - 1<br />")
        .Write("' &nbsp; &nbsp;'//ÿ�������˾ͼ�һ������<br />")
        .Write("' &nbsp; &nbsp;If i &gt; 0 And i Mod iCols = 0 Then Response.Write(""&lt;/tr&gt;&lt;tr&gt;"")<br />")
        .Write("' &nbsp; &nbsp;Response.Write(""&lt;td width="""""" & iColsPercent & """""" bgcolor=""""#CCE8CF""""&gt;&lt;font color=""""#000000""""&gt;"" & Server.HTMLEncode(oRs(2) & """") & ""&lt;/font&gt;&lt;/td&gt;"")<br />")
        .Write("' &nbsp; &nbsp;oRs.MoveNext<br />")
        .Write("'Next<br />")
        .Write("'//ѭ������ ��ʼ����ȱ����<br />")
        .Write("Do While i &lt; iPageSize<br />")
        .Write(" &nbsp; &nbsp;'//��������������ѡһ<br />")
        .Write(" &nbsp; &nbsp;If i Mod iCols = 0 Then<br />")
        .Write(" &nbsp; &nbsp; &nbsp; &nbsp;Response.Write(""&lt;/tr&gt;&lt;tr&gt;"") '//���Ҫ�����������ͼ������&lt;tr&gt;&lt;/tr&gt;<br />")
        .Write(" &nbsp; &nbsp; &nbsp; &nbsp;'Exit Do '//���ֻ�������һ�о�ֱ�ӽ���<br />")
        .Write(" &nbsp; &nbsp;End If<br />")
        .Write(" &nbsp; &nbsp;i = i + 1<br />")
        .Write(" &nbsp; &nbsp;Response.Write(""&lt;td width=""""""&FormatPercent(1 / iCols, 0)&"""""" bgcolor=""""#CCCCCC""""&gt;&amp;nbsp;&lt;/td&gt;"")<br />")
        .Write("Loop<br />")
        .Write("'//�����ҳ��Ϣ/��ʽ/TABLEβ<br />")
        .Write("Response.Write(""&lt;/tr&gt;&lt;tr&gt;&lt;td colspan="""""" & iCols & """""" bgcolor=""""#CCE8CF""""&gt;&lt;div class=""""kindbpager""""&gt;"" & sPager & "" ����: "" & sJumpPage & "" ҳ&lt;/div&gt;"" & sPageInfo & ""&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;"")<br />")
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
                                oDbPager.AddCondition Str4Sql(aChoose(x)) & " like '" & Replace(Replace(Replace(Str4Like(BStr(aKeyWord(x))), "*", "%"), "?", "_"), "��", "_") & "'"
                            End If
                        End If
                End Select
            End If
        Next
    End If
End Sub
%>