<%
Session.CodePage = "65001"
'Kin_Db_Pager 2.0.2 beta

Class Kin_Db_Pager
    Private iHaltStatus, sVersion
    Private oConn, iTimeOut
    Private sDbType, sTableName, bQuoteTable, iSpeed
    Private bDistinct, sPKey, sFields, sCondition
    Private sOrderByString, sOrderByStringRev, bOrderByPKey, bIsPKeySorted, sPKeyDefaultOrder
    Private aCondition(), aOrderBy(), aFields()
    Private iMaxRecords, iPage, iPageSize, iPageCount, iRecordCount, iMaxPageCount, iMaxRecordCount, iLastPageCount, iStartPosition, iEndPosition, iPageRecordCount
    Private sQueryString, sUrl, sFormMethod
    Private iPagerSize, iPagerTop, iPagerStyle
    Private iPage1Size, iPageOffSet
    Private sReWrite, sSeparator, sEllipsis, sPageParam, bPagerGroup, bLinkEllipsis
    Private sFirstPage, sPreviousGroup, sPreviousPage, sCurrentPage, sListPage, sNextPage, sNextGroup, sLastPage
    Private sPage, sPageCount, sRecordCount, sPagerInfo, sPagerExt
    Private sSqlString, oRecordSet, oCommand
    Private iCacheType, iCacheTimeOut
    Private Err

    Private Sub Class_Initialize()
        On Error Resume Next
        Randomize()
        sVersion = "Jorkin's Classic ASP Pagination Class Ver: 2.0.2 beta Build:20101010 Coding By Jorkin"
        iHaltStatus = -1
        sDbType = "MSSQL"
        bQuoteTable = True
        sTableName = Null
        sPKey = "ID"
        sFields = "*"
        sCondition = ""
        ReDim aCondition( -1)
        ReDim aOrderBy( -1)
        ReDim aFields(1)
        bOrderByPKey = False
        bIsPKeySorted = False
        iMaxRecords = -1
        iPage = 1
        iPageSize = 20
		iPage1Size = Null
		iPageOffSet = 0
        iPageCount = Null
        iMaxPageCount = Null
        iRecordCount = Null
        iMaxRecordCount = Null
        sFormMethod = "GET"
        sPageParam = "page"
        SetPageParam(sPageParam)
        iPagerStyle = 0
        sPKeyDefaultOrder = "ASC"
        iSpeed = 1
        sOrderByString = ""
        sFirstPage = "首页"
        sPreviousPage = "上一页"
        sPreviousGroup = "上一组"
        sCurrentPage = "{$Kin_CurrentPage}"
        sListPage = "{$Kin_ListPage}"
        sNextPage = "下一页"
        sNextGroup = "下一组"
        sLastPage = "末页"
        iPagerTop = 0
        iPagerSize = 9
        sPage = "{$Kin_Page}"
        sPageCount = "{$Kin_PageCount}"
        sRecordCount = "{$Kin_RecordCount}"
		sPagerList = "{$Kin_FirstPage} {$Kin_PreviousPage} {$Kin_ListPage} {$Kin_NextPage} {$Kin_LastPage}"
        sPagerInfo = "共有 {$Kin_RecordCount} 条记录 第 {$Kin_Page} 页/共 {$Kin_PageCount} 页"
        sSeparator = ""
        sEllipsis = "…"
        iTimeOut = 222
        bPagerGroup = False
        bLinkEllipsis = False
        iCacheType = -1
        iCacheTimeOut = 777
        bDistinct = False
        sPagerExt = ""
		iStartPosition = Null
    End Sub

    Private Sub Class_Terminate()
        If IsObject(oConn) Then
			oConn.Close()
	        Set oConn = Nothing
		End If
    End Sub

    Private Function IIf(bExp1, sVal1, sVal2)
        If (bExp1) Then
            IIf = sVal1
        Else
            IIf = sVal2
        End If
    End Function

    Private Function to_Int(v)
        On Error Resume Next
        to_Int = 0
        to_Int = Int(CDbl(v))
    End Function

    Public Function to_Str(v)
        On Error Resume Next
        to_Str = Trim(to_String(v))
    End Function

    Public Function to_String(v)
        On Error Resume Next
        If IsArray(v) Then v = Join(v)
        to_String = ""
        to_String = CStr(v)
    End Function

    Private Function to_Bool(v)
        On Error Resume Next
        to_Bool = False
        If UCase(to_Str(v)) = "TRUE" Then
            to_Bool = True
        Else
            to_Bool = CBool(v)
        End If
    End Function

    Private Function IsBlank(v)
        IsBlank = False
        Select Case VarType(v)
            Case 0, 1 : IsBlank = True
            Case 8 : If Len(v) = 0 Then IsBlank = True
            Case 9 : tmpType = TypeName(v) : If (tmpType = "Nothing") Or (tmpType = "Empty") Then IsBlank = True
            Case 8192, 8204, 8209 : If UBound(v) = -1 Then IsBlank = True
        End Select
    End Function

    Private Function InArray( sValue, aArray )
        Dim i
        InArray = False
        For i = 0 To UBound( aArray )
            If StrComp(aArray(i), sValue, 1) = 0 Then
                InArray = True
                Exit For
            End If
        Next
    End Function

    Private Function Str4Sql(v)
        Str4Sql = Replace(to_Str(v), "'", "''")
    End Function

    Private Sub doError(s)
        On Error Resume Next
		Select Case iHaltStatus
		Case -1, 0
			Randomize()
			Dim nRnd, i
			nRnd = to_Int(Rnd() * 29252888)
			With Response
				.Clear
				.Write "<div style=""width:98%;margin:auto;font-size:12px; cursor:pointer;line-height:150%"">"
				.Write "<label onClick=""ERRORDIV" & nRnd & ".style.display=(ERRORDIV" & nRnd & ".style.display=='none'?'':'none')"">"
				.Write "<span style=""background-color:#820222;color:#FFFFFF;height:23px;font-size:14px;"">〖 Kin_Pagination &#25552;&#31034;&#20449;&#24687;  ERROR 〗</span>"
				.Write "</label>"
				.Write "<div id=""ERRORDIV" & nRnd & """ style=""width:100%;border:1px solid #820222;padding:5px;overflow:hidden;"">"
				If Not IsBlank(s) Then
					.Write "<span style=""color:#FF0000;"">Kin_Pagination.Error:</span> " & Server.HTMLEncode(s) & "<br />"
				End If
				If IsObject(s) Then
					If s Is Err Then
						.Write "<span style=""color:#FF0000;"">Err.Description:</span> " & Err.Description & "<br />"
						.Write "<span style=""color:#FF0000;"">Err.Source:</span> " & Err.Source & "<br />"
					End If
				End If
				For i = 0 To oConn.Errors.Count -1
					.Write "<span style=""color:#FF0000;"">Connection.Error:</span> " & oConn.Errors.Item(i) & "<br />"
				Next
				If Not IsBlank(sSqlString) Then
					.Write "<span style=""color:#FF0000;"">Kin_Pagination.GetSql:</span> " & sSqlString & "<br />"
				End If
				.Write "<span style=""color:#FF0000;"">Kin_Pagination.Information:</span> " & sVersion & "<br />"
				.Write "<img width=""0"" height=""0"" src=""http://img.users.51.la/2782986.asp"" style=""display:none"" /></div>"
				.Write "</div><br />"
				If iHaltStatus = -1 Then .End()
			End With
	

		Case Else
		
		End Select

    End Sub

    Property Let TimeOut(v)
        iTimeOut = to_Int(v)
    End Property

    Public Sub Connect(v)
        On Error Resume Next
        If TypeName(v) = "Connection" Then
            Set oConn = v
            If oConn.CommandTimeout < iTimeOut Or oConn.ConnectionTimeout < iTimeOut Then
                If oConn.State = 1 Then v.Close()
                oConn.CommandTimeout = iTimeOut
                oConn.ConnectionTimeout = iTimeOut
                oConn.Open()
            End If
            SetDbType(GetDbType(oConn))
        Else
            doError "[Connection.Closed]"
        End If
        If Err Then Err.Clear
    End Sub

    Property Set ActiveConnection(v)
        Set oConn = v
        SetDbType(GetDbType(oConn))
    End Property

    Property Let ConnectionString(v)
        On Error Resume Next
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.ConnectionString = v
        oConn.CommandTimeout = iTimeOut
        oConn.ConnectionTimeout = iTimeOut
        oConn.Open()
        If Err Then doError Err.Description
        SetDbType(GetDbType(oConn))
    End Property

    Property Get ConnectionString()
        ConnectionString = oConn.ConnectionString
    End Property

    Property Get Conn()
        Set Conn = oConn
    End Property

    Private Function GetDbType(v)
        On Error Resume Next
        Dim sConnectionProvider : sConnectionProvider = v.Provider
        Select Case (sConnectionProvider)
            Case "MSDASQL.1", "SQLOLEDB.1", "SQLOLEDB" : GetDbType = "MSSQL"
            Case "MSDAORA.1", "OraOLEDB.Oracle" : GetDbType = "ORACLE"
            Case "Microsoft.Jet.OLEDB.4.0" : GetDbType = "ACCESS"
            Case Else : GetDbType = "ACCESS"
        End Select
    End Function

    Private Sub SetDbType(v)
        sDbType = UCase(to_Str(v))
        Select Case sDbType
            Case "MSSQL", "SQL", "SQLSERVER" : sDBType = "MSSQL"
            Case "MYSQL" : sDBType = "MYSQL" : bQuoteTable = False
            Case "ORACLE" : sDBType = "ORACLE" : bQuoteTable = False
            Case "PGSQL" : sDBType = "PGSQL"
            Case "MSSQLPRODUCE", "MSSQLPR", "MSSQL_PR", "PR", "PRODUCE" : sDBType = "MSSQLPRODUCE"
            Case Else : sDBType = "ACCESS"
        End Select
    End Sub

    Property Let DbType(v)
        SetDbType(v)
    End Property

    Property Get DbType()
        DbType = sDbType
    End Property

    Property Let Distinct(v)
        bDistinct = to_Bool(v)
    End Property

    Property Let TableName(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If CheckColumn(sTMPString)(0) Then
            sTableName = sTMPString
            If InStr(v, "[") > 0 Then bQuoteTable = False
            If InStr(v, "(") > 0 Then bQuoteTable = False
            If InStr(v, ".") > 0 Then bQuoteTable = False
            If InStr(v, ",") > 0 Then bQuoteTable = False
            If InStr(v, " ") > 0 Then bQuoteTable = False
        Else
            doError ".TableName = """ & sTMPString & """ 可能有符号匹配错误，请重新检查。"
        End If
    End Property

    Property Let QuoteTable(v)
        bQuoteTable = to_Bool(v)
    End Property

    Private Function QuoteName(v)
        QuoteName = IIf(bQuoteTable, "[" & v & "]", " " & v & " ")
    End Function

    Private Function FormatColumn(v)
        FormatColumn = to_String(Replace(Replace(Replace(to_Str(v), vbTab, " "), vbCr, " "), vbLf, " "))
    End Function

    Public Function CheckColumn(v)
        Dim bMatching, bMatchingAll, sTmpString1, sTmpString2, i, j
        Dim aCommaArray, aSingleQuoteArray, iUboundCommaArray, iUboundSingleQuoteArray
        ReDim aFieldsArray( -1)
        v = to_String(v)
        If Not IsBlank(to_Str(v)) Then
            aCommaArray = Split(v, ",")
            iUboundCommaArray = UBound(aCommaArray)
            bMatchingAll = True
            For i = 0 To iUboundCommaArray
                If Not IsBlank(to_Str(sTmpString1)) Then sTmpString1 = sTmpString1 & ","
                sTmpString1 = sTmpString1 & aCommaArray(i)
                aSingleQuoteArray = Split(sTmpString1, "'")
                iUboundSingleQuoteArray = UBound(aSingleQuoteArray)
                sTmpString2 = ""
                bMatchingAll = True
                bMatching = False
                If iUboundSingleQuoteArray Mod 2 = 0 Then
                    For j = 0 To iUboundSingleQuoteArray Step 2
                        sTmpString2 = sTmpString2 & aSingleQuoteArray(j)
                    Next
                    If (InStr(sTmpString2, "(") <= InStr(sTmpString2, ")")) And (UBound(Split(sTmpString2, "(")) = UBound(Split(sTmpString2, ")"))) Then bMatching = True
                End If
                bMatchingAll = bMatchingAll And bMatching
                If bMatchingAll And Not IsBlank(to_Str(sTmpString1)) Then
                    j = UBound(aFieldsArray) + 1
                    ReDim Preserve aFieldsArray(j)
                    aFieldsArray(j) = sTmpString1
                    sTmpString1 = ""
                    bMatchingAll = False
                    If i = iUboundCommaArray Then bMatchingAll = True
                End If
            Next
            CheckColumn = Array(bMatchingAll, aFieldsArray)
        Else
            CheckColumn = Array(False, Array())
        End If
    End Function

    Property Get TableName()
        TableName = sTableName
    End Property

	Property Let Pkey(v)
		Call setPKey(v)
	End Property

	Property Let PrimaryKey(v)
		Call setPKey(v)
	End Property
	
    Private Sub setPKey(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If CheckColumn(sTMPString)(0) Then
            sPKey = sTMPString
            If IsBlank(sPKey) Then sPKey = "ID"
            If Left(sPKey, 1) = "[" And Right(sPKey, 1) = "]" Then sPKey = to_Str(Replace(Replace(sPKey, "[", ""), "]", ""))
        Else
            doError ".PKey = """ & sTMPString & """ 可能有符号匹配错误，请重新检查。"
        End If
    End Sub

    Property Get PKey()
        PKey = sPKey
    End Property

    Property Let Fields(v)
        If IsBlank(to_Str(v)) Then v = "*"
        ReDim aFields( -1)
        ReDim aFieldString( -1)
        Call AddFields(v)
    End Property

    Private Sub AddFields(v)
        Dim a, a1, a2, a3, i, j, k, l
        v = FormatColumn(v)
        a = CheckColumn(FormatColumn(v))
        If Not a(0) Then
            doError ".Fields = """ & v & """ 可能有符号匹配错误，请重新检查。"
        Else
            a1 = a(1)
            k = UBound(a1)
            For i = 0 To k
                a2 = Split(a1(i), "AS ", -1, 1)
                j = UBound(a2)
                If j > 0 Then
                    l = Left(a1(i), InStrRev(a1(i), "AS ", -1 , 1) -1)
                    If CheckColumn(l)(0) Then
                        AddFieldsArray l, a2(j)
                    Else
                        AddFieldsArray a1(i), ""
                    End If
                Else
                    AddFieldsArray a1(i), ""
                End If
            Next
            k = UBound(aFields)
            ReDim aFieldString(k)
            For i = 0 To k
                If Not IsBlank(aFields(i)(1)) Then
                    aFieldString(i) = aFields(i)(0) & " AS " & aFields(i)(1)
                Else
                    aFieldString(i) = aFields(i)(0)
                End If
            Next
            sFields = Join(aFieldString, ", ")
        End If
    End Sub

    Private Sub AddFieldsArray(ByVal x, ByVal y)
        Dim iFields
        x = to_Str(x)
        If CheckColumn(x)(0) Then
            iFields = UBound(aFields) + 1
            ReDim Preserve aFields(iFields)
            aFields(iFields) = Array(x, to_Str(y))
        End If
    End Sub

    Property Get Fields()
        Fields = sFields
    End Property

    Property Let Condition(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If IsBlank(sTMPString) Then
            ReDim aCondition( -1)
        Else
            If CheckColumn(sTMPString)(0) Then
                ReDim aCondition(0)
                aCondition(0) = sTMPString
            Else
                doError ".Condition = """ & sTMPString & """ 可能有符号匹配错误，请重新检查。"
            End If
        End If
    End Property

    Public Sub AddCondition(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If IsBlank(sTMPString) Then Exit Sub
        If CheckColumn(sTMPString)(0) Then
            Dim iCondition : iCondition = UBound(aCondition) + 1
            ReDim Preserve aCondition(iCondition)
            aCondition(iCondition) = sTMPString
        Else
            doError ".AddCondition(""" & sTMPString & """) 可能有符号匹配错误，请重新检查。"
        End If
    End Sub

    Private Function Sorted(v)
        Sorted = IIf(UCase(to_Str(v)) = "DESC", "DESC", "ASC")
    End Function

    Private Function SortedRev(v)
        SortedRev = IIf(Sorted(v) = "DESC", "ASC", "DESC")
    End Function

    Private Function IsSorted(v)
        Select Case UCase(to_Str(v))
            Case "ASC", "DESC" : IsSorted = True
            Case Else : IsSorted = False
        End Select
    End Function

    Property Let OrderBy(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If CheckColumn(sTMPString)(0) Then
            ReDim aOrderBy( -1)
            Call AddOrderBy(sTMPString)
        Else
            doError ".OrderBy = """ & sTMPString & """ 可能有符号匹配错误，请重新检查。"
        End If
    End Property

    Public Sub AddOrderBy(v)
        Dim sTMPString : sTMPString = FormatColumn(v)
        If IsBlank(sTMPString) Then Exit Sub
        If CheckColumn(sTMPString)(0) Then
            Dim a1, a2, i, j, k
            a1 = Split(sTMPString, ",")
            k = UBound(a1)
            For i = 0 To k
                a2 = Split(a1(i), " ")
                j = UBound(a2)
                If j > 0 Then
                    If IsSorted(a2(j)) Then
                        Call AddOrderByArray(Left(a1(i), InStrRev(a1(i), " ") -1), Sorted(a2(j)))
                    Else
                        Call AddOrderByArray(a1(i), "ASC")
                    End If
                Else
                    Call AddOrderByArray(a1(i), "ASC")
                End If
            Next
        Else
            doError ".AddOrderBy(""" & sTMPString & """) 可能有符号匹配错误，请重新检查。"
        End If
    End Sub

    Public Sub AddOrderByArray(x, y)
        Dim i, j
        x = to_Str(x)
        If Not bOrderByPKey Then
            If CheckColumn(x)(0) Then
                i = UBound(aOrderBy) + 1
                ReDim Preserve aOrderBy(i)
                If StrComp(x, sPKey, 1) = 0 Or StrComp(x, "[" & sPKey & "]", 1) = 0 Then
                    bOrderByPKey = True
                    x = sPKey
				Else
					For j = 0 To UBound(aFields)
						If StrComp(sPKey, aFields(j)(0), 1) = 0 Or StrComp("[" & aFields(j)(0) & "]", sPKey, 1) = 0 Then
							If StrComp(x, aFields(j)(1), 1) = 0 Or StrComp("[" & aFields(j)(1) & "]", x, 1) = 0 Then
								bOrderByPKey = True
								x = sPKey
							End If
						End If
					Next
                End If
                aOrderBy(i) = Array(x, y)
            End If
        End If
    End Sub

    Private Sub CreateOrderByString()
        Dim i, j, k, aOrderByFields, iOrderByFields, sOrderByFieldsString, iFieldsString, bOrderByDistinct, bOrderByDistinctAll
        If Not bDistinct Then
            If iSpeed = 1 Then
                sOrderByString = " ORDER BY " & sPKey & " " & sPKeyDefaultOrder
                sOrderByStringRev = " ORDER BY " & sPKey & " " & SortedRev(sPKeyDefaultOrder)
            End If
        Else
            bDisTinct = True
            sOrderByString = ""
            sOrderByStringRev = ""
        End If
        k = UBound(aOrderBy)
        bOrderByDistinctAll = True
        If k > -1 Then
            ReDim aOrderByString(k), aOrderByStringRev(k)
            For i = 0 To k
                aOrderByString(i) = aOrderBy(i)(0) & " " & aOrderBy(i)(1)
                aOrderByStringRev(i) = aOrderBy(i)(0) & " " & SortedRev(aOrderBy(i)(1))
                If bDistinct And bOrderByDistinctAll Then
                    aOrderByFields = Split(sFields, aOrderBy(i)(0), -1, 1)
                    iOrderByFields = UBound(aOrderByFields)
                    If iOrderByFields > 0 Then
                        bOrderByDistinct = False
                        For j = 0 To iOrderByFields - 1
                            sOrderByFieldsString = sOrderByFieldsString & aOrderByFields(j)
                            iFieldsString = UBound(Split(sOrderByFieldsString, "'"))
                            If iFieldsString = -1 Or iFieldsString Mod 2 = 0 Then
                                bOrderByDistinct = True
                                Exit For
                            End If
                        Next
                        bOrderByDistinctAll = bOrderByDistinctAll And bOrderByDistinct
                    Else
                        bOrderByDistinctAll = False
                    End If
                End If
            Next
            If bDistinct And Not bOrderByDistinctAll Then
                doError "如果指定了 .DISTINCT = True，那么 ORDER BY 子句中的项就必须出现在选择列表中。"
            Else
                sOrderByString = " ORDER BY " & Join(aOrderByString, ", ")
                sOrderByStringRev = " ORDER BY " & Join(aOrderByStringRev, ", ")
                If Not bDistinct And Not bOrderByPKey Then
                    sOrderByString = sOrderByString & ", " & sPkey & " " & sPKeyDefaultOrder
                    sOrderByStringRev = sOrderByStringRev & ", " & sPkey & " " & SortedRev(sPKeyDefaultOrder)
                End If
            End If
        End If
    End Sub

    Property Get OrderBy()
        Call CreateOrderByString()
        OrderBy = sOrderByString
    End Property

    Property Let PKeyOrder(v)
        sPKeyDefaultOrder = Sorted(v)
    End Property

    Property Let PageParam(v)
        sPageParam = to_Str(v)
        If IsBlank(sPageParam) Then sPageParam = "page"
        SetPageParam(sPageParam)
    End Property

    Property Get PageParam()
        PageParam = sPageParam
    End Property

    Private Function SetPageParam(v)
        Dim x, y, sForm
        sQueryString = ""
        For Each x In Request.QueryString
            If StrComp(x, sPageParam, 1) <> 0 Then
                For Each y In Request.QueryString(x)
                    sQueryString = sQueryString & "&" & x & "=" & Server.URLEncode(y)
                Next
            End If
        Next
        sUrl = Request.ServerVariables("URL") & "?" & IIf(IsBlank(sQueryString), "", Mid(sQueryString, 2) & "&")
        If sFormMethod = "GET" Then
            sUrl = sUrl & sPageParam & "="
            SetPage(Request.QueryString(v))
        Else
            SetPage(Request.Form(v))
            For Each x In Request.Form
                If StrComp(x, sPageParam, 1) <> 0 Then
                    For Each y In Request.Form(x)
                        If InStr(y, vbCr) + InStr(y, vbLf) > 0 Then
                            sForm = sForm & "<textarea name=""" & x & """>" & Server.HTMLEncode(y) & "</textarea>" & vbCrLf
                        Else
                            sForm = sForm & "<input name=""" & x & """ value=""" & Server.HTMLEncode(y) & """>" & vbCrLf
                        End If
                    Next
                End If
            Next
            Trace sForm
        End If
    End Function

    Property Let FormMethod(v)
        sFormMethod = UCase(to_Str(v))
        If sFormMethod <> "POST" Then sFormMethod = "GET"
    End Property

    Property Let PageSize(v)
        iPageSize = to_Int(v)
        If iPageSize < 1 Then iPageSize = -1
    End Property

	Property Let Page1Size(v)
		iPage1Size = to_Int(v)
		If iPage1Size < 1 Then iPage1Size = Null
	End Property

    Private Function SetPage(v)
        iPage = to_Int(v)
        If iPage < 1 Then iPage = 1
    End Function

    Property Let Page(v)
        SetPage(v)
    End Property

    Property Let AbsolutePage(v)
        SetPage(v)
    End Property

    Property Get Page()
        Page = iPage
    End Property

    Property Get AbsolutePage()
        AbsolutePage = iPage
    End Property

	Property Let PageOffSet(v)
		iPageOffSet = to_Int(v)
	End Property

	Property Get PageOffSet()
		PageOffSet = iPageOffSet
	End Property

    Property Let CacheType(v)
        Select Case UCase(to_Str(v))
		Case "APPLICATION", 0
			iCacheType = 0
		Case "SESSION", 1
			iCacheType = 1
        Case "COOKIE", 2
			iCacheType = 2
		Case Else
			iCacheType = -1
        End Select
    End Property

    Property Let CacheTimeOut(v)
        iCacheTimeOut = to_Int(v)
    End Property

    Private Function setRecordCountCache(v)
		On Error Resume Next
		v = "{KIN_PAGINATION_CACHE:" & UCase(to_Str(v)) & "}"
		Dim iCacheDateTimeOut
		iCacheDateTimeOut = Now() + TimeSerial(0,iCacheTimeOut, 0)
        Select Case UCase(iCacheType)
		Case 0
			Application.Lock()
			Application.Contents.Remove(v)
			Application(v) = Array(iRecordCount, iCacheDateTimeOut)
			Application.UnLock()
		Case 1
			Session.Contents.Remove(v)
			Session(v) = Array(iRecordCount, iCacheDateTimeOut)
        Case 2
			Response.Cookies(v) = iRecordCount
			Response.Cookies(v).Expires = iCacheDateTimeOut
        End Select
		If Err Then doError Err.Description
    End Function

    Private Function getRecordCountCache(v)
		On Error Resume Next
		v = "{KIN_PAGINATION_CACHE:" & UCase(to_Str(v)) & "}"
		iRecordCount = Null
        Select Case UCase(iCacheType)
		Case 0
			If IsArray(Application(v)) Then
				If Application(v)(1) < Now() Then Exit Function
				iRecordCount = to_Int(Application(v)(0))
			End If
		Case 1
			If IsArray(Session(v)) Then
				If Session(v)(1) < Now() Then Exit Function
				iRecordCount = to_Int(Session(v)(0))
			End If
        Case 2
			If Not IsBlank(Request.Cookies(v)) Then
				iRecordCount = to_Int(Request.Cookies(v))
			End If
        End Select
		If Err Then doError Err.Description
    End Function

    Private Sub CalculateRecordCount()
        On Error Resume Next
		If IsBlank(oConn) Then doError "必须设定Connection。"
        Call CreateCondition()
        Dim sSql
        sSql = "SELECT COUNT(*) FROM " & IIf(bDistinct, "(SELECT DISTINCT " & sFields & " FROM " & QuoteName(sTableName) & " " & sCondition & ") KIN_PAGINATION_TABLE", QuoteName(sTableName) & " " & sCondition)
        getRecordCountCache(sSql)
        If IsNull(iRecordCount) Then
            If sDbType = "MSSQLPRODUCE" Then
                Call CreateCommand(0)
            Else
                Dim oRs
                Set oRs = oConn.Execute(sSql)
                iRecordCount = to_Int(oRs(0))
                Set oRs = Nothing
            End If
        	setRecordCountCache(sSql)
        End If
        If Err Then doError Err.Description
    End Sub

    Private Sub CalculatePageCount()
        If IsNull(iRecordCount) Then CalculateRecordCount()
        iMaxRecordCount = iRecordCount
        If iMaxRecords > 0 And iMaxRecordCount > iMaxRecords Then iMaxRecordCount = iMaxRecords
		If IsNull(iPage1Size) Then iPage1Size = iPageSize
        If iRecordCount = 0 Or iPageSize = -1 Then
            iPageCount = 1
            iMaxPageCount = 1
        Else
			iPageCount = to_Int((iRecordCount - iPage1Size) / iPageSize) + Sgn(ABS((iRecordCount - iPage1Size) Mod iPageSize)) + 1
			iMaxPageCount = to_Int((iMaxRecordCount - iPage1Size) / iPageSize) + Sgn(ABS((iMaxRecordCount - iPage1Size) Mod iPageSize)) + 1
        End If
        If iPage > iMaxPageCount Then iPage = iMaxPageCount
        iLastPageCount = iMaxRecordCount - (iPageSize * (iMaxPageCount -2)) - iPage1Size
        iStartPosition = (iPage -1) * iPageSize - iPageSize + iPage1Size '//计算开始位置
		If iPageOffSet < 0 then iStartPosition = iStartPosition + iPageOffSet '//开始页码偏移量
		If iStartPosition < 0 Then iStartPosition = 0
        iEndPosition = iPage * iPageSize - iPageSize + iPage1Size '//计算结束位置
		If iPageOffSet > 0 then iEndPosition = iEndPosition + iPageOffSet '//结束页码偏移
        If iEndPosition > iMaxRecordCount Then iEndPosition = iMaxRecordCount
        iPageRecordCount = iPageSize
        If iPage = iMaxPageCount Then iPageRecordCount = iLastPageCount
		If iPage = 1 Then iPageRecordCount = iPage1Size
    End Sub

    Property Let Speed(v)
        iSpeed = Sgn(Abs(to_Int(v)))
    End Property

    Public Function getSql()
        If IsObject(oRecordSet) Then
            getSql = oRecordSet.Source
			If sDbType = "MSSQLPRODUCE" Then getSql = getSql & vbCrlf & "{ " & sSqlString & " }"
        Else
            Call CreateOrderByString()
            Call CreateSqlString()
            getSql = sSqlString
        End If
    End Function

    Private Sub CreateCondition()
        If Not IsBlank(sCondition) Then Exit Sub
        If UBound(aCondition) >= 0 Then
            sCondition = " WHERE ((" & Join(aCondition, ") AND (") & ")) "
        End If
    End Sub

    Private Sub CreateSqlString()
        If IsNull(iPageCount) Then CalculatePageCount()
        If bDisTinct Then sDbType = "ACCESS" '//偷懒，谁家Distinct还上百万数据的。。。
        If sDbType = "MSSQLPRODUCE" And Not IsBlank(sSqlString) Then Exit Sub
        If iPageSize > 0 Then
            Select Case sDbType
                Case "MSSQL"
                    If iSpeed = 1 Then
                        If iPage = 1 Then
                            sSqlString = "SELECT " & sFields & " FROM " & _
                                         "( SELECT TOP " & iPageRecordCount & " * " & _
                                         "FROM " & QuoteName(sTableName) & sCondition & sOrderByString & ") AS KIN_PAGINATION_TABLE"
                        ElseIf iPage = iPageCount Then
                            sSqlString = "SELECT " & sFields & " FROM (" & _
                                         "SELECT TOP " & iPageRecordCount & " * " & _
                                         "FROM " & QuoteName(sTableName) & sCondition & sOrderByStringRev & ") AS KIN_PAGINATION_TABLE1" & sOrderByString
                        ElseIf sOrderByString = " ORDER BY " & sPKey & " ASC" Then
                            sSqlString = "SELECT " & "TOP " & iPageRecordCount & " " & sFields & " FROM " & QuoteName(sTableName) & " WHERE " & sPKey & " > " & _
                                         "( SELECT MAX(" & sPKey & ") FROM " & _
                                         "( SELECT TOP " & (iStartPosition) & " " & sPKey & " FROM " & QuoteName(sTableName) & sCondition & sOrderByString & _
                                         ") AS KIN_PAGINATION_TABLE1 )" & sOrderByString
                        ElseIf sOrderByString = " ORDER BY " & sPKey & " DESC" Then
                            sSqlString = "SELECT " & sFields & " FROM (" & _
                                         "SELECT TOP " & iPageRecordCount & " * FROM " & QuoteName(sTableName) & " WHERE " & sPKey & " > " & _
                                         "( SELECT MAX(" & sPKey & ") FROM " & _
                                         "( SELECT TOP " & (iRecordCount - (iStartPosition + iPageRecordCount)) & " " & sPKey & " FROM " & QuoteName(sTableName) & sCondition & sOrderByStringRev & _
                                         ") AS KIN_PAGINATION_TABLE1 )" & sOrderByStringRev & _
                                         ") AS KIN_PAGINATION_TABLE2" & sOrderByString
                        Else
                            If iPage * 2 > iPageCount Then
                                sSqlString = "SELECT " & sFields & " FROM " & QuoteName(sTableName) & " WHERE " & sPKey & " IN (" &_
                                             "SELECT TOP " & iPageRecordCount & " " & sPKey & " FROM (" &_
                                             "SELECT TOP " &(iRecordCount - iStartPosition) & " * FROM " & QuoteName(sTableName) & sCondition & sOrderByStringRev &_
                                             ") KIN_PAGINATION_TABLE1" & sOrderByString & ")" & sOrderByString
                            Else
                                sSqlString = "SELECT " & sFields & " FROM " & QuoteName(sTableName) & " WHERE " & sPKey & " IN (" &_
                                             "SELECT TOP " & iPageRecordCount & " " & sPKey & " FROM (" &_
                                             "SELECT TOP " & iEndPosition & " * FROM " & QuoteName(sTableName) & sCondition & sOrderByString &_
                                             ") KIN_PAGINATION_TABLE1" & sOrderByStringRev & ")" & sOrderByString
                            End If
                        End If
                    Else
                        sSqlString = "SELECT " & sFields & " FROM " & QuoteName(sTableName) & " " & _
                                     "WHERE " & sPKey & " IN (" & _
                                     "SELECT TOP " & iEndPosition & " " & sPKey & " FROM " & QuoteName(sTableName) & " " & sCondition & sOrderByString & ")"
                        If iPage>1 Then
                            sSqlString = sSqlString & " AND " & sPKey & " NOT IN (" & _
                                         "SELECT TOP " & iStartPosition & " " & sPKey & " FROM " & QuoteName(sTableName) & " " & sCondition & sOrderByString & ")"
                        End If
                        sSqlString = sSqlString & sOrderByString
                    End If
                Case "MYSQL"
                    sSqlString = "SELECT " & IIf(bDisTinct, "DISTINCT ", "") & sFields & " FROM " & QuoteName(sTableName) & sCondition & sOrderByString & IIf(iPageSize > 0, " LIMIT " & iStartPosition & "," & IIf(iPage = 1, iPage1Size, iPageSize) , "")
                Case "ORACLE"
                    If bDistinct Then
                        sSqlString = "SELECT * " & _
                                     "FROM (SELECT KIN_PAGINATION_TABLE1.*, ROWNUM KIN_PAGINATION_PKEY " & _
                                     "FORM (SELECT DISTINCT " & sFields & " FROM " & QuoteName(sTableName) & sCondition & sOrderByString & ") KIN_PAGINATION_TABLE1 " & _
                                     "WHERE ROWNUM <= " & iEndPosition & ") " & _
                                     "WHERE KIN_PAGINATION_PKEY > " & iStartPosition
                    Else
                        sSqlString = "SELECT " & sFields & " " & _
                                     "FROM (SELECT KIN_PAGINATION_TABLE1.*, ROWNUM KIN_PAGINATION_PKEY " & _
                                     "FORM (SELECT * FROM " & QuoteName(sTableName) & sCondition & sOrderByString & ") KIN_PAGINATION_TABLE1 " & _
                                     "WHERE ROWNUM <= " & iEndPosition & ") " & _
                                     "WHERE KIN_PAGINATION_PKEY > " & iStartPosition
                    End If
                Case Else
                    sSqlString = "SELECT " & IIf(bDisTinct, "DISTINCT ", "") & IIf(iSpeed = 1, "TOP " & iEndPosition & " ", "") & sFields & " FROM " & QuoteName(sTableName) & sCondition & sOrderByString
            End Select
        Else
            sSqlString = "SELECT " & IIf(bDisTinct, "DISTINCT ", "") & sFields & " FROM " & QuoteName(sTableName) & sCondition & sOrderByString
        End If
    End Sub

    Private Sub CreateCommand(v)
        Call CreateOrderByString()
		v = Sgn(to_Int(v))
		If IsObject(oCommand) And v = 1 Then
			With oCommand
				.Parameters.Delete "@bReturn"
				.Parameters.Append .CreateParameter("@bReturn", 3, 1, 10, v)
			End With
		Else
			Set oCommand = Server.CreateObject("ADODB.Command")
			With oCommand
				.CommandType = 4
				.Prepared = True
				.ActiveConnection = oConn
				.CommandText = "sp_Kin_Pagination"
				.CommandTimeout = iTimeOut
				.Parameters.Append .CreateParameter("@getSQL", 200, 2, 4000)
				.Parameters.Append .CreateParameter("@iRecordCount", 3, 2, 4)
				'-------// 出参结束, 入参开始 //------
				.Parameters.Append .CreateParameter("@sTableName", 200, 1, 256, QuoteName(sTableName))
				.Parameters.Append .CreateParameter("@sPKey", 200, 1, 128, sPKey)
				.Parameters.Append .CreateParameter("@sFields", 200, 1, 1024, sFields)
				.Parameters.Append .CreateParameter("@sCondition", 200, 1, 2048, sCondition)
				.Parameters.Append .CreateParameter("@sOrderByString", 200, 1, 1024, sOrderByString)
				.Parameters.Append .CreateParameter("@sOrderByStringRev", 200, 1, 1024, sOrderByStringRev)
				.Parameters.Append .CreateParameter("@iPage", 3, 1, 10, iPage)
				.Parameters.Append .CreateParameter("@iPageSize", 3, 1, 10, iPageSize)
				.Parameters.Append .CreateParameter("@iPage1Size", 3, 1, 10, IIF(IsNull(iPage1Size), iPageSize, iPage1Size))
				.Parameters.Append .CreateParameter("@iPageOffSet", 3, 1, 10, iPageOffSet)
				.Parameters.Append .CreateParameter("@iMaxRecords", 3, 1, 10, iMaxRecords)
				.Parameters.Append .CreateParameter("@iSpeed", 3, 1, 10, iSpeed)
				.Parameters.Append .CreateParameter("@bDistinct", 3, 1, 10, to_Int(bDistinct))
				.Parameters.Append .CreateParameter("@bReturn", 3, 1, 10, v)
				.Execute()
				sSqlString = oCommand.Parameters("@getSQL").Value
				iRecordCount = oCommand.Parameters("@iRecordCount").Value
			End With
		End If
    End Sub

    Private Sub CreateRecordSet()
        On Error Resume Next
        Server.ScriptTimeOut = iTimeOut
        Call getSql()
        Set oRecordSet = Server.CreateObject("ADODB.RecordSet")
        Select Case sDbType
            Case "MSSQL"
                oRecordSet.Open sSqlString, oConn, 1, 1
            Case "MSSQLPRODUCE"
				Call CreateCommand(1)
                oRecordSet.CursorLocation = 3
                oRecordSet.LockType = 1
                oRecordSet.Open oCommand
            Case "ORACLE"
                Set oRecordSet = oConn.CreateDynaset(sSqlString, 0)
            Case "MYSQL"
                oRecordSet.Open sSqlString, oConn, 1, 1
            Case Else
                oRecordSet.Open sSqlString, oConn, 1, 1
                'If iPageSize > 0 Then oRecordSet.PageSize = iPageSize
                If Not oRecordSet.EOF Then oRecordSet.AbsolutePosition = iStartPosition + 1
        End Select
        If oRecordSet.State = 0 Then
			doError "[记录集未打开或找开失败。]"
			oRecordSet.Open "SELECT " & IIf(bDistinct, "DISTINCT ", "") & sFields & " FROM " & QuoteName(sTableName) & " WHERE 1 = 0", oConn, 1, 1
        End If
        If Err Then doError Err.Description
    End Sub

    Property Get RecordSet()
        On Error Resume Next
        CreateRecordSet()
        Set RecordSet = oRecordSet
    End Property

    Property Get GetRows()
        On Error Resume Next
        CreateRecordSet()
        Dim aGetRows
        If Not oRecordSet.EOF Then
            aGetRows = oRecordSet.GetRows(iPageSize)
        Else
            ReDim aGetRows(oRecordSet.Fields.Count -1, -1)
        End If
        GetRows = aGetRows
    End Property

	Property Get GetJSArray()
		'要注意大小写
        On Error Resume Next
        CreateRecordSet()
		GetJSArray = "[["
        If Not oRecordSet.EOF Then
			GetJSArray = GetJSArray & oRecordSet.GetString(, , "],[","]],[[", "")
		End If
		GetJSArray = Left(GetJSArray, Len(GetJSArray)-3)
		trace GetJSArray
	End Property

	Property Get GetJSON()
		'要注意大小写
	End Property

    Property Get PagerStyle()
        PagerStyle = iPagerStyle
    End Property

    Property Let PagerGroup(b)
        bPagerGroup = to_Bool(b)
    End Property

    Property Let LinkEllipsis(b)
        bLinkEllipsis = to_Bool(b)
    End Property

    Property Let PagerSize(n)
        iPagerSize = to_Int(n)
    End Property

    Property Let Halt(b)
        iHaltStatus = Sgn(to_Int(b))
    End Property

    Public Sub CreateIndex()
        On Error Resume Next
        If sDBType <> "MSSQL" Then Exit Sub
        Trace aOrderBy
    End Sub

    Property Get PagerInfo()
        If IsNull(iPageCount) Then CalculatePageCount()
        PagerInfo = sPagerInfo
        PagerInfo = Replace(PagerInfo, sRecordCount, to_Int(iMaxRecordCount))
        PagerInfo = Replace(PagerInfo, sPageCount, to_Int(iMaxPageCount))
        PagerInfo = Replace(PagerInfo, sPage, to_Int(iPage))
    End Property

    Private Function ReWrite(n)
        n = to_Int(n)
        If Not IsBlank(sRewrite) Then
            ReWrite = IIf(n > 0, Replace(sReWrite, "*", n), sReWrite)
        Else
            ReWrite = sUrl & IIf(n > 0, n, "*")
        End If
    End Function

    Property Let ReWritePath(v)
        Call SetReWritePath(v)
    End Property

    Private Sub SetReWritePath(v)
        Dim x, y, z
        sReWrite = v
        For Each x In Request.QueryString
            z = ""
            For Each y In Request.QueryString(x)
                If Not IsBlank(y) Then z = z & "-" & Server.URLEncode(y)
            Next
            z = Mid(z, 2)
            If IsBlank(z) Then z = "-"
            sReWrite = Replace(sReWrite, "{" & x & "}" , z, 1, -1, 1)
        Next
    End Sub

    Property Let Separator(v)
        sSeparator = v
    End Property

    Property Let Ellipsis(v)
        sEllipsis = v
    End Property

    Property Let PagerTop(v)
        iPagerTop = to_Int(v)
    End Property

    Property Let FirstPage(v)
        sFirstPage = v
    End Property

    Property Let LastPage(v)
        sLastPage = v
    End Property

    Property Let PreviousPage(v)
        sPreviousPage = v
    End Property

    Property Let NextPage(v)
        sNextPage = v
    End Property

    Property Let CurrentPage(v)
        sCurrentPage = v
    End Property

    Property Let ListPage(v)
        sListPage = v
    End Property

    Property Let MaxRecords(v)
        iMaxRecords = to_Int(v)
        If iMaxRecords < 0 Then iMaxRecords = -1
    End Property

    Property Get RecordCount()
        If IsNull(iRecordCount) Then CaculateRecordCount()
        RecordCount = to_Int(iMaxRecordCount)
    End Property

    Property Get PageSize()
        PageSize = to_Int(iPageSize)
    End Property

    Property Get Condition()
        If IsBlank(sCondition) Then CreateCondition()
        Condition = sCondition
    End Property

    Property Get ReWritePath
        ReWritePath = sReWrite
    End Property

    Property Get PageCount()
        If IsNull(iPageCount) Then CalculatePageCount()
        PageCount = to_Int(iMaxPageCount)
    End Property

    Property Let Style(v)
        Call SetPagerStyle(v)
    End Property

    Property Let PagerStyle(v)
        Call SetPagerStyle(v)
    End Property

    Private Sub SetPagerStyle(v)
        iPagerStyle = v
        Select Case iPagerStyle
            Case 0
            Case 1
                iPagerSize = 7
                iPagerTop = 2
                sFirstPage = "&lt;&lt;"
                sPreviousPage = "&lt;"
                sNextPage = "&gt;"
                sLastPage = "&gt;&gt;"
            Case 2
                sFirstPage = "&laquo; First"
                sPreviousPage = "&#8249; Previous"
                sNextPage = "Next &#8250;"
                sLastPage = "Last &raquo;"
                sPagerInfo = "Page {$Kin_Page} of {$Kin_PageCount} ({$Kin_RecordCount} items)"
            Case 3
                iPagerSize = 10
                bPagerGroup = True
                sPreviousGroup = "前十页"
                sNextGroup = "后十页"
            Case 4 '//Only IE
                iPagerTop = 1
                iPagerSize = 10
                bPagerGroup = True
                sEllipsis = "<font face=""webdings"">`</font>" '//"<font face=""webdings"">q</font>"
                sCurrentPage = "<font face=""webdings"">;</font>"
                sListPage = "<font>{$Kin_ListPage}</font>"
                sFirstPage = "<font face=""webdings"">9</font>"
                sPreviousGroup = "<font face=""webdings"">7</font>"
                sPreviousPage = "<font face=""webdings"">3</font>"
                sNextPage = "<font face=""webdings"">4</font>"
                sNextGroup = "<font face=""webdings"">8</font>"
                sLastPage = "<font face=""webdings"">:</font>"
                sPagerExt = "<style>.listpage{display:none;}</style>"
                bLinkEllipsis = True
            Case 5 '//为JS调用
                sPagerInfo = "{$Kin_Page},{$Kin_PageCount},{$Kin_RecordCount}"
            Case Else
                sEllipsis = "<font color=""#220282"">[……]</font>"
                sSeparator = " "
                sFirstPage = "<font color=""#220282"">[首页]</font>"
                sPreviousPage = "<font color=""#220282"">[上一页]</font>"
                sNextPage = "<font color=""#220282"">[下一页]</font>"
                sLastPage = "<font color=""#220282"">[末页]</font>"
                sListPage = "<font color=""#220282"">[{$Kin_ListPage}]</font>"
                sCurrentPage = "<font color=""#820222"">[{$Kin_CurrentPage}]</font>"
                bLinkEllipsis = True
        End Select
    End Sub

    Property Get Pager()
        If IsNull(iPageCount) Then CalculatePageCount()
        Dim i, ii, iPagerStart, iPagerEnd, sPager
        If bPagerGroup Then
            iPagerEnd = iPagerSize * Abs(Int( -1 * (iPage / iPagerSize)))
        Else
            iPagerEnd = iPage + to_Int(iPagerSize / 2)
            If iPagerEnd > iMaxPageCount Then iPagerEnd = iMaxPageCount
        End If
        iPagerStart = iPagerEnd - iPagerSize + 1
        If iPagerStart < 1 Then iPagerStart = 1
        iPagerEnd = iPagerStart + iPagerSize -1
        If iPagerEnd > iMaxPageCount Then iPagerEnd = iMaxPageCount
        ReDim aPager(4), aPager0(0), aPager1( -1), aPager2( -1), aPager3( -1), aPager4(0)

        If sFormMethod = "GET" Then
            If iPage > 1 Then
                sPager = IIf(IsBlank(sFirstPage), "", "<a class=""firstpage"" href=""" & Rewrite(1) & """>" & sFirstPage & "</a>")
                If bPagerGroup Then sPager = sPager & IIf(iPage > iPagerSize, IIf(IsBlank(sPreviousGroup), "", sSeparator & "<a class=""previousgroup"" href=""" & Rewrite(iPage - iPagerSize) & """>" & sPreviousGroup & "</a>"), IIf(IsBlank(sPreviousGroup), "", sSeparator & "<span class=""disabled"">" & sPreviousGroup & "</span>"))
                sPager = sPager & IIf(IsBlank(sPreviousPage), "", sSeparator & "<a class=""previouspage"" href=""" & Rewrite(iPage -1) & """>" & sPreviousPage & "</a>")
            Else
                sPager = IIf(IsBlank(sFirstPage), "", "<span class=""disabled"">" & sFirstPage & "</span>")
                If bPagerGroup Then sPager = sPager & IIf(IsBlank(sPreviousGroup), "", sSeparator & "<span class=""disabled"">" & sPreviousGroup & "</span>")
                sPager = sPager & IIf(IsBlank(sPreviousPage), "", sSeparator & "<span class=""disabled"">" & sPreviousPage & "</span>")
            End If
            aPager0(0) = sPager
            If iPagerTop > 0 Then
                ii = IIf(iPagerTop < iPagerStart, iPagerTop, iPagerStart - 1)
                ReDim aPager1(ii)
                For i = 1 To ii
                    aPager1(i -1) = "<a class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</a>"
                Next
                If iPagerTop < iPagerStart -1 Then
                    If bLinkEllipsis Then
                        aPager1(ii) = "<a class=""ellipsis"" href=""" & Rewrite(to_Int((iPagerStart + 1 + ii) / 2)) & """>" & sEllipsis & "</a>"
                    Else
                        aPager1(ii) = "<span class=""ellipsis"">" & sEllipsis & "</span>"
                    End If
                End If
            End If
            If iPagerSize > 0 Then
                ReDim aPager2(iPagerEnd - iPagerStart)
                For i = iPagerStart To iPagerEnd
                    If i = iPage Then
                        aPager2(i - iPagerStart) = "<span class=""current"">" & Replace(sCurrentPage, "{$Kin_CurrentPage}", i, 1, -1, 1) & "</span>"
                    Else
                        aPager2(i - iPagerStart) = "<a class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</a>"
                    End If
                Next
            End If
            If iPagerTop > 0 Then
                ii = IIf(iMaxPageCount - iPagerTop > iPagerEnd, iMaxPageCount - iPagerTop + 1, iPagerEnd + 1)
                ReDim aPager3(iMaxPageCount - ii + 1)
                If iMaxPageCount - iPagerTop > iPagerEnd Then
                    If bLinkEllipsis Then
                        aPager3(0) = "<a class=""ellipsis"" href=""" & ReWrite((iMaxPageCount - iPagerTop + iPagerEnd + 1) / 2) & """>" & sEllipsis & "</a>"
                    Else
                        aPager3(0) = "<span class=""ellipsis"">" & sEllipsis & "</span>"
                    End If
                End If
                For i = ii To iMaxPageCount
                    aPager3(i - ii + 1) = "<a class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</a>"
                Next
            End If
            If iMaxPageCount > iPage Then
                sPager = IIf(IsBlank(sNextPage), "", "<a class=""nextpage"" href=""" & Rewrite(iPage + 1) & """>" & sNextPage & "</a>")
                If bPagerGroup Then sPager = sPager & IIf(iMaxPageCount > iPagerEnd, IIf(IsBlank(sNextGroup), "", sSeparator & "<a class=""nextgroup"" href=""" & Rewrite(IIf(iPage + iPagerSize > iMaxPageCount, iMaxPageCount, iPage + iPagerSize)) & """>" & sNextGroup & "</a>"), IIf(IsBlank(sNextGroup), "", sSeparator & "<span class=""disabled"">" & sNextGroup & "</span>"))
                sPager = sPager & IIf(IsBlank(sLastPage), "", sSeparator & "<a class=""lastpage"" href=""" & Rewrite(iMaxPageCount) & """>" & sLastPage & "</a>")
            Else
                sPager = IIf(IsBlank(sNextPage), "", "<span class=""disabled"">" & sNextPage & "</span>")
                If bPagerGroup Then sPager = sPager & IIf(IsBlank(sNextGroup), "", sSeparator & "<span class=""disabled"">" & sNextGroup & "</span>")
                sPager = sPager & IIf(IsBlank(sLastPage), "", sSeparator & "<span class=""disabled"">" & sLastPage & "</span>")
            End If
            aPager4(0) = sPager
            aPager(0) = Join(aPager0, sSeparator)
            aPager(1) = Join(aPager1, sSeparator)
            aPager(2) = Join(aPager2, sSeparator)
            aPager(3) = Join(aPager3, sSeparator)
            aPager(4) = Join(aPager4, sSeparator)
            Pager = IIf(IsBlank(aPager(0)), "", aPager(0)) & IIf(IsBlank(aPager(1)), "", sSeparator & aPager(1)) & IIf(IsBlank(aPager(2)), "", sSeparator & aPager(2)) & IIf(IsBlank(aPager(3)), "", sSeparator & aPager(3)) & IIf(IsBlank(aPager(4)), "", sSeparator & aPager(4))
            Pager = Pager & sPagerExt
        Else
            If iPage > 1 Then
                sPager = IIf(IsBlank(sFirstPage), "", "<button type=""button"" class=""firstpage"" href=""" & Rewrite(1) & """>" & sFirstPage & "</button>")
                If bPagerGroup Then sPager = sPager & IIf(iPage > iPagerSize, IIf(IsBlank(sPreviousGroup), "", sSeparator & "<button class=""previousgroup"" href=""" & Rewrite(iPage - iPagerSize) & """>" & sPreviousGroup & "</button>"), IIf(IsBlank(sPreviousGroup), "", sSeparator & "<span class=""disabled"">" & sPreviousGroup & "</span>"))
                sPager = sPager & IIf(IsBlank(sPreviousPage), "", sSeparator & "<button class=""previouspage"" href=""" & Rewrite(iPage -1) & """>" & sPreviousPage & "</button>")
            Else
                sPager = IIf(IsBlank(sFirstPage), "", "<span class=""disabled"">" & sFirstPage & "</span>")
                If bPagerGroup Then sPager = sPager & IIf(IsBlank(sPreviousGroup), "", sSeparator & "<span class=""disabled"">" & sPreviousGroup & "</span>")
                sPager = sPager & IIf(IsBlank(sPreviousPage), "", sSeparator & "<span class=""disabled"">" & sPreviousPage & "</span>")
            End If
            aPager0(0) = sPager
            If iPagerTop > 0 Then
                ii = IIf(iPagerTop < iPagerStart, iPagerTop, iPagerStart - 1)
                ReDim aPager1(ii)
                For i = 1 To ii
                    aPager1(i -1) = "<button class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</button>"
                Next
                If iPagerTop < iPagerStart -1 Then
                    If bLinkEllipsis Then
                        aPager1(ii) = "<button class=""ellipsis"" href=""" & Rewrite(to_Int((iPagerStart + 1 + ii) / 2)) & """>" & sEllipsis & "</button>"
                    Else
                        aPager1(ii) = "<span class=""ellipsis"">" & sEllipsis & "</span>"
                    End If
                End If
            End If
            If iPagerSize > 0 Then
                ReDim aPager2(iPagerEnd - iPagerStart)
                For i = iPagerStart To iPagerEnd
                    If i = iPage Then
                        aPager2(i - iPagerStart) = "<span class=""current"">" & Replace(sCurrentPage, "{$Kin_CurrentPage}", i, 1, -1, 1) & "</span>"
                    Else
                        aPager2(i - iPagerStart) = "<button class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</button>"
                    End If
                Next
            End If
            If iPagerTop > 0 Then
                ii = IIf(iMaxPageCount - iPagerTop > iPagerEnd, iMaxPageCount - iPagerTop + 1, iPagerEnd + 1)
                ReDim aPager3(iMaxPageCount - ii + 1)
                If iMaxPageCount - iPagerTop > iPagerEnd Then
                    If bLinkEllipsis Then
                        aPager3(0) = "<button class=""ellipsis"" href=""" & ReWrite((iMaxPageCount - iPagerTop + iPagerEnd + 1) / 2) & """>" & sEllipsis & "</button>"
                    Else
                        aPager3(0) = "<span class=""ellipsis"">" & sEllipsis & "</span>"
                    End If
                End If
                For i = ii To iMaxPageCount
                    aPager3(i - ii + 1) = "<button class=""listpage"" href=""" & ReWrite(i) & """>" & Replace(sListPage, "{$Kin_ListPage}", i, 1, -1, 1) & "</button>"
                Next
            End If
            If iMaxPageCount > iPage Then
                sPager = IIf(IsBlank(sNextPage), "", "<button class=""nextpage"" href=""" & Rewrite(iPage + 1) & """>" & sNextPage & "</button>")
                If bPagerGroup Then sPager = sPager & IIf(iMaxPageCount > iPagerEnd, IIf(IsBlank(sNextGroup), "", sSeparator & "<button class=""nextgroup"" href=""" & Rewrite(IIf(iPage + iPagerSize > iMaxPageCount, iMaxPageCount, iPage + iPagerSize)) & """>" & sNextGroup & "</button>"), IIf(IsBlank(sNextGroup), "", sSeparator & "<span class=""disabled"">" & sNextGroup & "</span>"))
                sPager = sPager & IIf(IsBlank(sLastPage), "", sSeparator & "<button class=""lastpage"" href=""" & Rewrite(iMaxPageCount) & """>" & sLastPage & "</button>")
            Else
                sPager = IIf(IsBlank(sNextPage), "", "<span class=""disabled"">" & sNextPage & "</span>")
                If bPagerGroup Then sPager = sPager & IIf(IsBlank(sNextGroup), "", sSeparator & "<span class=""disabled"">" & sNextGroup & "</span>")
                sPager = sPager & IIf(IsBlank(sLastPage), "", sSeparator & "<span class=""disabled"">" & sLastPage & "</span>")
            End If
            aPager4(0) = sPager
            aPager(0) = Join(aPager0, sSeparator)
            aPager(1) = Join(aPager1, sSeparator)
            aPager(2) = Join(aPager2, sSeparator)
            aPager(3) = Join(aPager3, sSeparator)
            aPager(4) = Join(aPager4, sSeparator)
            Pager = IIf(IsBlank(aPager(0)), "", aPager(0)) & IIf(IsBlank(aPager(1)), "", sSeparator & aPager(1)) & IIf(IsBlank(aPager(2)), "", sSeparator & aPager(2)) & IIf(IsBlank(aPager(3)), "", sSeparator & aPager(3)) & IIf(IsBlank(aPager(4)), "", sSeparator & aPager(4))
            Pager = Pager & sPagerExt
			Pager = "<form method=""post"" target=""" & sUrl & """>" & Pager
			Pager = Pager & sForm
			Pager = Pager & ""
			Pager = Pager & "</form>"
        End If
    End Property

	Function JumpInput()
	
	End Function

	Function JumpButton()
	
	End Function
	
	Function JumpMenu()
	
	End Function
	
    Property Get JumpPager(arglist, sJumpPagerAttr)
        If IsNull(iPageCount) Then CalculatePageCount()
        Dim iStart, iEnd, sJumpPager, i, j, sRandomize, iUBound
        If Not IsArray(arglist) Then arglist = Split(UCase(to_Str(arglist)), ",")
        iUBound = UBound(arglist)
        sRandomize = "Kin_Pagination" & to_Int(Rnd() * 29252888)
        sJumpPager = vbCrLf & "<scr" & "ipt type=""text/javascr" & "ipt"">function " & sRandomize & "(o){if(!(o)){alert('err');return false}var s" & sRandomize & "=o.value.split('/')[0];if(!isNaN(s" & sRandomize & ")&&s" & sRandomize & ".length>0){document.location.href=" & Replace("'" & Rewrite(0) & "'", "*", "'+s" & sRandomize & "+'") & "}return false}</script>" & vbCrLf
        For i = 0 To iUBound
            Select Case UCase(to_Str(arglist(i)))
                Case "INPUT", "BUTTON"
                    j = Len(iMaxPageCount)
                    sJumpPager = sJumpPager & "<input id=""" & sRandomize & "_input"" onkeydown=""if(event.keyCode==13){" & sRandomize & "(this)}"" type=""text"" title=""&#35831;&#36755;&#20837;&#25968;&#23383;&#10;&#13;&#22238;&#36710;&#36339;&#36716;"" size=""" & IIf(j < 3, 3, j) & """ maxlength=""" & j * 2 & """ value=""" & iPage & """ " & sJumpPagerAttr & " />" & vbCrLf
                    If UCase(to_Str(arglist(i))) = "BUTTON" Then sJumpPager = sJumpPager & "<button id=""" & sRandomize & "_button"" onclick=""" & sRandomize & "(document.getElementById('" & sRandomize & "_input'))"" " & sJumpPagerAttr & " >GO</button>" & vbCrLf
                Case Else
                    sJumpPager = sJumpPager & "<select id=""" & sRandomize & "_select"" onChange=""" & sRandomize & "(this)"" " & sJumpPagerAttr & "　>" & vbCrLf
                    iEnd = iPage + 100
                    If iEnd > iMaxPageCount Then iEnd = iMaxPageCount
                    iStart = iPage - 100
                    If iStart < 1 Then iStart = 1
                    For j = iStart To iEnd
                        sJumpPager = sJumpPager & "<option value=""" & j & """" & IIf(j = iPage, " selected=""selected"" ", "") & ">"&j&"</option>" & vbCrLf
                    Next
                    sJumpPager = sJumpPager & "</select>" & vbCrLf
            End Select
        Next
        JumpPager = sJumpPager
    End Property

    Public Function Clone()
        If Not IsObject(oRecordSet) Then doError "需要使用.RecordSet()或.GetRows()或.getSQL()方法之后才可使用。"
        Set Clone = New Kin_Db_Pager
        Clone.Connect(oConn)
        Clone.TableName = sTableName
        Clone.PKey = sPKey
        Clone.Fields = sFields
        Clone.OrderBy = sOrderByString
        Clone.Distinct = bDistinct
        Clone.Distinct = bDistinct
        Clone.FirstPage = sFirstPage
        Clone.PreviousPage = sPreviousPage
        Clone.NextPage = sNextPage
        Clone.LastPage = sLastPage
        Clone.PagerTop = iPagerTop
        Clone.PageList = sPageList
        Clone.LinkEllipsis = bLinkEllipsis
        Clone.Ellipsis = sEllipsis
        Clone.Separator = sSeparator
        Clone.OrderBy = sOrderByString
        Clone.Condition = Mid(sCondition, 7)
    End Function

    Property Let PagerExt(v)
        sPagerExt = to_Str(v)
    End Property

	Public Function Template()
		Dim oDictionary
		oDictionary("$pagesize") = iPageSize
		oDictionary("$pagecount") = iPageCount
		oDictionary("$recordcount") = iRecordCount
		oDictionary("$firstpage") = sFirstPage
		oDictionary("$previouspage") = sPreviousPage
		oDictionary("$nextpage") = sNextPage
		oDictionary("$lastpage") = sLastPage
		oDictionary("$lastpage") = sLastPage
		oDictionary("$lastpage") = sLastPage
		oDictionary("$lastpage") = sLastPage
		oDictionary("$lastpage") = sLastPage
		oDictionary("$lastpage") = sLastPage
		Template = oDictionary
	End Function

	Public Sub Help()
		With Response
			.Write("<div style=""text-align:left"">")
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
			.Write("'//指定表示页数的URL变量 默认值:""page""<br />")
			.Write("'oDbPager.PageParam = ""page""<br />")
			.Write("'//指定数据库类型.默认值:""MSSQL""<br />")
			.Write("'oDbPager.DbType = ""ACCESS""<br />")
			.Write("'//指定目标表 可用临时表""(Select * From [Table])""<br />")
			.Write("oDbPager.TableName = ""Kin_Article""<br />")
			.Write("'//选择列 用逗号分隔 默认值:""*""<br />")
			.Write("oDbPager.Fields = ""*""<br />")
			.Write("'//指定该表的主键<br />")
			.Write("oDbPager.PKey = ""Article_ID""<br />")
			.Write("'//指定每页记录集数量<br />")
			.Write("oDbPager.PageSize = iPageSize<br />")
			.Write("'//指定当前页数<br />")
			.Write("oDbPager.Page = Request.QueryString(""page"")<br />")
			.Write("'//指定排序条件<br />")
			.Write("oDbPager.OrderBy = ""Article_ID DESC""<br />")
			.Write("'//添加条件 可多次使用.如果用Or条件需要(条件1 Or 条件2 Or ...)<br />")
			.Write("oDbPager.AddCondition ""Article_Status &gt; 0""<br />")
			.Write("If Day(Date) Mod 2 = 0 Then<br />")
			.Write("&nbsp; &nbsp; oDbPager.AddCondition ""(Article_ID &lt; 104 Or Article_ID &gt; 222)""<br />")
			.Write("End If<br />")
			.Write("'GetCondition """","""",""""<br />")
			.Write("'//也可以直接使用自定义的SQL语句选取记录集(!!!不能分页)<br />")
			.Write("'oDbPager.Sql = ""Select * From Kin_Article Where Article_ID &lt; 222 Order By Article_ID Desc""<br />")
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
			.Write("</div>")
			.Flush()
		End With
	End Sub

End Class




'// .TemplateHead = "<table>"
'// .TemplateBody = "<tr><td>{$News_ID}</td><td>{$News_Title}</td><td>{$News_DateTime}</td></tr>"
'// .TemplateFoot = "</table>"
'// .TemplateAssign("News_DateTime", FormatDate("News_Article", "aaaaaa"))




%>