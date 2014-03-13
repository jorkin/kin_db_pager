<%
Const PI = 3.1415926535897932
Dim post_, get_
Set post_ = Request.Form
Set get_ = Request.QueryString

'功能:将参数转为Bool格式
Public Function to_Bit(v)
    On Error Resume Next
    to_Bit = 0
	to_Bit = ABS(to_Bool(v))
End Function

'功能:将参数转为Bool格式
Public Function to_Bool(v)
    On Error Resume Next
    to_Bool = False
	Select Case UCASE(to_Str(v))
	Case "TRUE"
		to_Bool = True
	Case Else
		to_Bool = CBool(v)
	End Select
End Function

'为FormatTxt。
Public Function to_Chr(v)
    Dim iCodePage
    iCodePage = Session.CodePage
    ReDim aArray(247)
    to_Chr = ""
    Select Case to_Int(iCodePage)
        Case 65001
            aArray(153) = ChrW(153)
            aArray(162) = ChrW(162)
            aArray(163) = ChrW(163)
            aArray(169) = ChrW(169)
            aArray(174) = ChrW(174)
            aArray(176) = ChrW(176)
            aArray(177) = ChrW(177)
            aArray(178) = ChrW(178)
            aArray(179) = ChrW(179)
            aArray(185) = ChrW(185)
            aArray(186) = ChrW(186)
            aArray(188) = ChrW(188)
            aArray(189) = ChrW(189)
            aArray(190) = ChrW(190)
            aArray(247) = ChrW(247)
        Case Else
            aArray(153) = "&#153;"
            aArray(162) = "&#162;"
            aArray(163) = "&#163;"
            aArray(169) = "&#169;"
            aArray(174) = "&#174;"
            aArray(176) = "&#176;"
            aArray(177) = "&#177;"
            aArray(178) = "&#178;"
            aArray(179) = "&#179;"
            aArray(185) = "&#185;"
            aArray(186) = "&#186;"
            aArray(188) = "&#188;"
            aArray(189) = "&#189;"
            aArray(190) = "&#190;"
            aArray(247) = "&#247;"
    End Select
    If Not IsEmpty(aArray(v)) Then to_Chr = aArray(v)
End Function

'功能:将参数转为日期格式
Public Function to_Date(v)
    If IsDate(v) Then
        to_Date = CDate(v)
    Else
        to_Date = CDate(0)
    End If
End Function

'功能:只取数字
Public Function to_Dbl(v)
    On Error Resume Next
    to_Dbl = 0.0
    to_Dbl = CDBL(FormatNumber(sValue,to_Int(Len(v)),-1,0,0))
End Function

Public Function to_Double(v)
	Dim regEx, aDouble
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = "([^\d\.])"
    v = regEx.Replace(v, "")
	aDouble = Split(v, ".")
	v = to_Int(aDouble(0))
	aDouble(0) = ""
	to_Double = v & "." & to_Int(Join(aDouble, ""))
End Function

'功能:把TEXTAREA的文字转为HTML输出
Public Function to_HTML(v)
    v = Server.HTMLEncode(to_String(v))
    v = Replace(v, vbTab, "    ")
    v = Replace(v, vbNewLine, "<br />")
    v = Replace(v, vbCr, "<br />")
    v = Replace(v, "  ", "&nbsp; ")
    v = Replace(v, "  ", "&nbsp; ")
    to_HTML = JDecode(v)
End Function

'功能:只取整数
Public Function to_Int(v)
    On Error Resume Next
	Dim i
	to_Int = 0
	to_Int = Int(CDbl(v))
End Function

Public Function to_Integer(v)

End Function
'to_IntArray格式化数组数列字符串
Public Function to_IntArray(v)
	Dim i
	If Not IsArray(v) Then v = Split(to_String(v), ",")
	For i = 0 To UBound(v)
		v(i) = to_Int(v(i))
	Next
    to_IntArray = Join(v, ",")
End Function

'功能:强制转为字符串(去空格)
Public Function to_Str(v)
    On Error Resume Next
	v = to_String(v)
    to_Str = Trim(v)
End Function

Public Function to_StrArray(v)
	Dim i
	If Not IsArray(v) Then v = Array(v)
	For i = 0 To UBound(v)
		v(i) = to_Str(v(i))
	Next
	to_StrArray = v
End Function

'功能:强制转为字符串，保留前后空格。
Public Function to_String(v)
    On Error Resume Next
    to_String = ""
	If IsArray(v) Then
		to_String = Join(v, "")
	Else
    	to_String = CStr(v)
	End If
End Function

Public Function to_StringArray(v)
	Dim i
	If Not IsArray(v) Then v = Array(v)
	For i = 0 To UBound(v)
		v(i) = to_String(v(i))
	Next
	to_StringArray = v
End Function

'格式化Tags{tag1},{tag2}...
Public Function to_Tags(s)
    Dim a, i
    s = to_Str(s)
    a = Split(s, ",")
    For i = 0 To UBound(a)
        a(i) = to_Str(a(i))
	    a(i) = Replace(a(i), "}", "")
	    a(i) = Replace(a(i), "{", "")
    Next
    to_Tags = "{" & Join(a, "},{") & "}"
End Function

'反转数组
Public Function ArrayRev(aArray)
    Dim i, j
    j = UBound(aArray)
    Redim aNewArray(j)
    For i = 0 to UBound(aArray)
        aNewArray(j-i)=arrayinput(i)
    Next
    ArrayRev = aNewArray
End Function

'功能：使用 &# 将 HTML 中的特殊字符进行 Unicode 编码
Public Function ASCII(v)
    Dim i
    v = to_String(v)
    For i = 1 To Len(v)
        ASCII = ASCII & "&#x" & Hex(AscW(Mid(v, i, 1))) & ";"
    Next
End Function

Function bytesToBstr(BODY, CSET)
    Dim oStream
    Set oStream = Server.CreateObject("ADO"&"DB.STR"&"EAM")
    oStream.Type = 1
    oStream.Mode = 3
    oStream.Open
    oStream.Write BODY
    oStream.Position = 0
    oStream.Type = 2
    oStream.Charset = CSET
    bytesToBstr = oStream.ReadText
    oStream.Close
    Set oStream = Nothing
End Function

'功能:用来在指定CheckBox的哪几个值上打勾
Public Function CheckBoxScript(ByVal FormElement , ByVal ElementValue)
    CheckBoxScript = "<scr" & "ipt language=""javascript"" type=""text/javascript"">" & vbCrLf & "String.prototype."
    CheckBoxScript = CheckBoxScript & "ReallyDo=function(){return this.replace(/(^\s*)|(\s*$)/g,"""");}" & vbCrLf
    CheckBoxScript = CheckBoxScript & "var Jorkin = """ & ElementValue & """.split("","");" & vbCrLf
    CheckBoxScript = CheckBoxScript & "for (i = 0; i < " & FormElement & ".length; i++){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "  for (j = 0; j < Jorkin.length; j++){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "    if (Jorkin[j].ReallyDo() == " & FormElement & "[i].value.ReallyDo()){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "      " & FormElement & "[i].checked = true;break; } } }</scr" & "ipt>" & vbCrLf
End Function

'功能:用来在选择<input type="radio">
Public Function RadioScript(ByVal FormElement, ByVal ElementValue)
    RadioScript = "<scr" & "ipt type=""text/javascript"">" & vbCrLf
    RadioScript = RadioScript & "var Jorkin='" & ElementValue & "';" & vbCrLf
    RadioScript = RadioScript & "  for(i = 0; i < " & FormElement & ".length; i++){" & vbCrLf
    RadioScript = RadioScript & "    if (Jorkin == " & FormElement & "[i].value){" & vbCrLf
    RadioScript = RadioScript & "      " & FormElement & "[i].checked = true}}</scr" & "ipt>" & vbCrLf
End Function

'功能:生成一个RGB颜色代码
Public Function ColorRGB(r, g, b)
    ColorRGB = "#" & Right(String(6, "0") & Hex(RGB(b, g, r)), 6)
End Function

'功能:判断发言是否来自外部
Public Function ChkPost()
    Dim server_v1, server_v2
    Chkpost = False
    server_v1 = LCase(to_Str(Request.ServerVariables("HTTP_REFERER")))
    server_v1 = Mid(server_v1, InStr(server_v1, "://") + 3)
    server_v2 = LCase(to_Str(Request.ServerVariables("SERVER_NAME")))
    If Left(server_v1,Len(server_v2)) = server_v2 Then Chkpost = True
End Function

Private Sub Clear(ByRef v)
	On Error Resume Next
    If IsObject(v) Then 
        Set v = Nothing
    ElseIf IsArray(v) Then
		Erase v
	Else
        v = Empty
    End If
End Sub

'//<c:out></c:out>
Sub Cout(sValue, sDefault, bEscapeXML)
	If IsBlank(sValue) Then
		echo(sDefault)
	Else
		If to_Bool(bEscapeXML) Then
			echo(to_HTML(sValue))
		Else
			echo(sValue)
		End If
	End If
End Sub

Sub ClearAll()
    Dim Cookie, Subkey
    For Each Cookie in Request.Cookies
        If Request.Cookies(Cookie).HasKeys Then
            For Each Subkey in Request.Cookies(Cookie)
                Response.Cookies(Cookie)(Subkey) = Empty
                Response.Cookies(Cookie)(Subkey).Expires = Now() -1
            Next
        Else
            Response.Cookies(Cookie) = Empty
            Response.Cookies(Cookie).Expires = Now() -1
        End If
    Next
    Session.Contents.RemoveAll()
    Application.Contents.RemoveAll()
End Sub

'功能:截取中英文混合字符串前N个字符.
Function CutStr(sReallyDo, iReallyDo)
    CutStr = KutStr(sReallyDo, iReallydo, "...")
End Function

Function KutStr(sReallyDo, iReallyDo, sJorkin)
    If IsBlank(sReallyDo) Then Exit Function
    If IsBlank(sJorkin) Then sJorkin = ""
    sReallyDo = Replace(sReallyDo, vbNewLine, " ")
    sReallyDo = Replace(sReallyDo, vbCr, " ")
	sReallyDo = ReplaceAll(sReallyDo, "  ", " ", True)
    Dim i, sJorkinLength, sReallyDoLength, CharASCW
    sReallyDoLength = 0
    sJorkinLength = StrLength(sJorkin)
    For i = 1 To Len(sReallyDo)
        CharASCW = ASCW(Mid(sReallyDo, i, 1))
        If CharASCW<0 Or CharASCW>255 Then
            sReallyDoLength = sReallyDoLength + 2
        Else
            sReallyDoLength = sReallyDoLength + 1
        End If
        If sReallyDoLength + sJorkinLength <= iReallyDo Then
            KutStr = Left(sReallyDo, i)
        End If
    Next
    If sReallyDoLength > iReallyDo Then
        KutStr = KutStr & sJorkin
    Else
        KutStr = sReallyDo
    End If
End Function

Public Function doHref(sUrl)
	sUrl = to_Str(sUrl)
	If Not IsBlank(sUrl) Then
		echo("try{")
		If Left(LCase(sUrl), 11) = "javascript:" Then 
			echo(Mid(sUrl, 12))	
		Else
			Select Case LCase(sUrl)
				Case "back"
            		echo("history.back();")
				Case "referer", "referrer"
					echo("var dReferrer=document.referrer||" & Str4Js(Request.ServerVariables("HTTP_REFERER")) & "||'';")
					echo("if(dReferrer!=''){location.href=dReferrer}else{history.back();location.reload();}")
				Case "close"
					echo("try{window.open('','_self');window.close()}catch(e){}")
					echo("try{window.opener=null;window.close()}catch(e){}")
					echo("try{document.body.insertAdjacentHTML(""beforeEnd"", ""<object id='noTipClose' classid='clsid:ADB880A6-D8FF-11CF-9377-00AA003B7A11'><param name='Command' value='Close' /></object>"");document.all.noTipClose.Click()}catch(e){}")
				Case "reload"
					echo("try{opener.reload()}catch(e){}")
					echo("try{parent.reload()}catch(e){}")
				Case "reloadopener"
					echo("try{opener.reload()}catch(e){}")
				Case "reloadparent"
					echo("try{parent.reload()}catch(e){}")
				Case "closereloadopener"
					doHref("reloadopener")
					doHref("close")
				Case "closereloadparent"
					doHref("reloadparent")
					doHref("close")
				Case "void(0)", "void", ""
				Case Else
                echo("location.href=" & Str4Js(sUrl) & ";")
			End Select
		End If
		echo("}catch(e){alert(e)}")
		Response.Flush()
	End If
End Function

'功能:输出alert信息并实现页面跳转
Public Function doAlert(sInfo, sUrl)
    On Error Resume Next
    sUrl = sUrl(sUrl)
    sInfo = sUrl(sInfo)
    If IsBlank(sUrl & sInfo) Then Exit Function
    echo("<scr" & "ipt type=""text/javascr" & "ipt"">" )
    If Not IsBlank(sInfo) Then echo("alert(" & Str4Js(sInfo) & ");")
    doHref(sUrl)
    echo("</scr" & "ipt>" )
    If Not IsBlank(sUrl) Then Response.End
End Function

Function doConfirm(sMessage, sTrue, sFalse)
	sMessage = to_Str(sMessage)
	sTrue = to_Str(sTrue)
	sFalse = to_Str(sFalse)
    echo("<scr" & "ipt type=""text/javascr" & "ipt"">")
    echo("if(window.confirm(" & Str4Js(sMessage) & ")){")
    doHref(sTrue)
    echo("}else{")
    doHref(sFalse)
    echo("}")
    echo("</scr" & "ipt>")
    Response.Flush()
End Function

Sub doError(Err)
	On Error Resume Next
	If IsObject(Err) Then
		If Not CONST_IsDebug Then Exit Sub
        echo "</scr" & "ipt>"
        echo "Error : # " & Err.Number & " <br />"
        echo "Description : " & Err.Description & "<br />"
        echo "Source : " & Err.Source & "<br />"
	Else
        echo "Error : " & to_HTML(Err) & "<br />"
	End If
	Err.Clear
	Response.End
End Sub

'功能:把字符串中的半角字符转为全角
Public Function SBC2DBC(v)
    '只有Chr(32)到Chr(126)转为全角才有意义
    '可以根据需要改为: For i = 1 To 256
    v = to_String(v)
    For i = 32 To 126
        v = Replace(v, Chr(i), Chr(i -23680))
    Next
    SBC2DBC = v
End Function

'功能:把字符串中的全角字符转为半角
Public Function DBC2SBC(v)
    '只有Chr(-23648)到Chr(-23554)转为半角才有意义
    '可以根据需要改为: For i = -23679 To -23424
    v = to_String(v)
    For i = -23648 To -23554
        v = Replace(v, Chr(i), Chr(i + 23680))
    Next
    DBC2SBC = v
End Function

'功能:删除二维数组中的iColumn列上值为sValue的行,当iColumn为-1时删除第sValue行
Public Function DelArray2(aArray, iColumn, sValue, bCompare)
    If Not IsArray(aArray) Then Exit Function
    Dim i, j, k, iUBound1, iUBound2
    iUBound1 = UBound(aArray)
    iUBound2 = UBound(aArray, 2)
    k = -1
    Dim aTmpArray()
    ReDim aTmpArray(iUBound1, k)
    If iColumn < 0 Or iColumn > iUBound1 Then
        For i = 0 To iUBound2
            If StrComp(i, sValue, bCompare)<>0 Then
                k = k + 1
                ReDim Preserve aTmpArray(iUBound1, k)
                For j = 0 To iUBound1
                    aTmpArray(j, k) = aArray(j, i)
                Next
            End If
        Next
    Else
        For i = 0 To iUBound2
            If StrComp(aArray(iColumn, i), sValue, bCompare)<>0 Then
                k = k + 1
                ReDim Preserve aTmpArray(iUBound1, k)
                For j = 0 To iUBound1
                    aTmpArray(j, k) = aArray(j, i)
                Next
            End If
        Next
    End If
    DelArray2 = aTmpArray
End Function

Public Sub echo(v)
	If IsArray(v) Then
		Response.Write(Join(v, ""))
	Else
    	Response.Write(to_String(v))
	End If
    Response.Flush()
End Sub

Public Sub echol(v)
	echo(v)
	echo(vbCrlf)
    Response.Flush()
End Sub

Sub die(v)
	echo(v)
	Response.End()
End Sub

'功能:执行SQL语句
Public Function Exec(sCommand)
    On Error Resume Next
    Server.ScriptTimeOut = 29252888
    OpenConn()
	Select Case TypeName(oConn)
		Case "Connection"
			Set Exec = oConn.Execute(sCommand)
		Case "IOraDatabase"
			Set Exec = oConn.CreateDynaset(sCommand, 0)
		Case Else
			Set Exec = oConn.Execute(sCommand)
	End Select
    If Err Then doError(Err)
End Function

'功能:在指定的记录集对象上设置筛选操作并打开一个新的记录集对象。
Public Function FilterField(tmpRs, tmpFilter)
    Set FilterField = tmpRs.Clone
    FilterField.Filter = tmpFilter
End Function

Function FIND_IN_SET(f,i)
	
End Function

'功能:多功能日期格式化函数
Public Function FormatDate(sDateTime, sReallyDo)
    Dim sTmpString, sJorkin, i
    sTmpString = ""
    FormatDate = ""
    sReallyDo = to_Str(sReallyDo)
    If Not IsDate(sDateTime) Then
        sDateTime = FormatDateTime(0, 1)
        Exit Function
    End If
    sDateTime = CDate(sDateTime)
    Select Case UCase(sReallyDo)
        Case "0", "1", "2", "3", "4"
            FormatDate = FormatDateTime(sDateTime, sReallyDo)
        Case "ATOM", "ISO8601", "W3C", "SITEMAP"
            FormatDate = Replace(FormatDate(sDateTime, "YYYY-MM-DD|HH:NN:SS+08:00"), "|", "T")
        Case "COOKIE", "RFC822", "RFC1123", "RFC2822", "RSS"
            FormatDate = FormatDate(sDateTime, "W, DD MMM YYYY HH:NN:SS +0800")
        Case "RFC850", "RFC1036"
            FormatDate = FormatDate(sDateTime, "WW, DD-MMM-YY HH:NN:SS +0800")
        Case "RND", "RAND", "RANDOMIZE" '//随机字符串
            Randomize
            sJorkin = Rnd()
            FormatDate = FormatDate(sDateTime, "YYYYMMDDHHNNSS") & _
                         to_Int((9 * 10^7 -1) * sJorkin) + 10^7
        Case Else
            For i = 1 To Len(sReallyDo)
                sJorkin = Mid(sReallyDo, i, 1)
                If Right(sTmpString, 1) = sJorkin Or Right(sTmpString, 1) = "C" Or Right(sTmpString, 1) = "Z" Or Right(sTmpString, 1) = "U" Or Right(sTmpString, 1) = "E" Or Right(sTmpString, 1) = "J" Then
                    sTmpString = sTmpString & sJorkin
                Else
                    FormatDate = FormatDate & FormatDateTimeString(sDateTime, sTmpString)
                    sTmpString = sJorkin
                End If
            Next
            FormatDate = FormatDate & FormatDateTimeString(sDateTime, sTmpString)
    End Select
End Function

Public Function FormatDateTimeString(sDateTime, sReallyDo)
    Dim sLocale
    sLocale = GetLocale()
    SetLocale("en-gb")
    Select Case sReallyDo
        Case "YYYY" '// 4位数年
            FormatDateTimeString = Year(sDateTime)
        Case "YY" '// 2位数年
            FormatDateTimeString = Right(Year(sDateTime), 2)
        Case "CMMMM", "CMMM", "CMM", "CM" '// 中文月份名
            SetLocale("zh-cn")
            FormatDateTimeString = MonthName(Month(sDateTime))
        Case "MMMM" '// 英文月份名(全)
            FormatDateTimeString = MonthName(Month(sDateTime), False)
        Case "MMM" '//英文月份名(缩)
            FormatDateTimeString = MonthName(Month(sDateTime), True)
        Case "MM" '// 月(补0)
            FormatDateTimeString = Right("0" & Month(sDateTime), 2)
        Case "M" '// 月
            FormatDateTimeString = Month(sDateTime)
        Case "JD" '// 日(加st nd rd th)
            FormatDateTimeString = Day(sDateTime)
            Select Case FormatDateTimeString
            Case 1, 21, 31
                FormatDateTimeString = FormatDateTimeString & "st"
            Case 2, 22
                FormatDateTimeString = FormatDateTimeString & "nd"
            Case 3, 23
                FormatDateTimeString = FormatDateTimeString & "rd"
            Case Else
                FormatDateTimeString = FormatDateTimeString & "th"
        End Select
        Case "DD" '// 日(补0)
            FormatDateTimeString = Right("0" & Day(sDateTime), 2)
        Case "D" '// 日
            FormatDateTimeString = Day(sDateTime)
        Case "HH" '// 小时(补0)
            FormatDateTimeString = Right("0" & Hour(sDateTime), 2)
        Case "H" '// 小时
            FormatDateTimeString = Hour(sDateTime)
        Case "NN" '// 分(补0)
            FormatDateTimeString = Right("0" & Minute(sDateTime), 2)
        Case "N" '// 分
            FormatDateTimeString = Minute(sDateTime)
        Case "SS" '//秒(补0)
            FormatDateTimeString = Right("0" & Second(sDateTime), 2)
        Case "S" '//秒
            FormatDateTimeString = Second(sDateTime)
        Case "CWW", "CW" '// 中文星期
            SetLocale("zh-cn")
            FormatDateTimeString = WeekdayName(Weekday(sDateTime))
        Case "WW" '// 英文星期(全)
            FormatDateTimeString = WeekdayName(Weekday(sDateTime), False)
        Case "W" '// 英文星期(缩)
            FormatDateTimeString = WeekdayName(Weekday(sDateTime), True)
        Case "CT" '// 12小时制(上午/下午)
            SetLocale("zh-tw")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "UT" '// 12小时制(AM/PM)
            SetLocale("en-us")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "ET" '// 12小时制(a.m./p.m.)
            SetLocale("es-ar")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "ZT" '// 12小时制(AM/PM)(补0)
            SetLocale("en-za")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "T" '// 24小时制时间
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case Else
            FormatDateTimeString = sReallyDo
    End Select
    SetLocale(sLocale)
End Function

'功能:储存单位转换函数
Public Function FormatFileSize(iSize)
    Dim aUnits, I
    aUnits = Array("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    I = Log(Abs(iSize)) \ 7
    If I > UBound(aUnits) Then I = UBound(aUnits)
    FormatFileSize = FormatNumber(iSize / (1024 ^ I), 2, -1) & " " & aUnits(I)
End Function

'返回表达式，此表达式已被格式化为中文数值。
'来源:http://jorkin.reallydo.com/article.asp?id=605

Function FormatNumber2(iNumber)
    On Error Resume Next
    If Not IsNumeric(iNumber) Then Exit Function
    Dim i, j, k, iNumberLength, sString, aNumber3, aNumber2, aNumber1
    aNumber1 = Array("零", "一", "二", "三", "四", "五", "六", "七", "八", "九")
    aNumber2 = Array("", "十", "百", "千")
    aNumber3 = Array("", "万", "亿", "万亿")
    iNumber = to_Int(iNumber)
    iNumberLength = Len(iNumber)
    If to_Int(iNumber / 10^15) <> 0 Then
        FormatNumber2 = "数值过大"
        Exit Function
    End If
    For i = 1 To iNumberLength
        j = Mid(iNumber, i, 1)
        sString = sString & aNumber1(j)
        If j > 0 Then sString = sString & aNumber2((iNumberLength - i) Mod 4)
        sString = Replace(sString, aNumber1(0) & aNumber1(0), aNumber1(0))
        If (iNumberLength - i) Mod 4 = 0 Then
            If i> 1 And Right(sString, 1) = aNumber1(0) Then sString = Left(sString, Len(sString) -1)
            sString = sString & aNumber3(to_Int((iNumberLength - i) / 4))
        End If
    Next
    If Left(sString, Len(aNumber1(1) & aNumber2(1))) = aNumber1(1) & aNumber2(1) Then sString = Mid(sString, Len(aNumber1(1) & aNumber2(1)))
    FormatNumber2 = sString
End Function

'按位们替换#为字符
Public Function FormatString(sString, sPattern)
    Dim i
    sString = to_String(sString)
    sPattern = to_String(sPattern)
	For i = 1 To Len(sString)
		sPattern = Replace(sPattern, "#", Mid(sString, i, 1), 1, 1)
	Next
	FormatString = sPattern
End Function

'转特殊符号
Public Function FormatTxt(v)
    v = to_String(v)
    v = Replace(v, "[2]", to_Chr(178))
    v = Replace(v, "[0]", to_Chr(186))
    v = Replace(v, "[1/2]", to_Chr(189))
    v = Replace(v, "[1/4]", to_Chr(188))
    v = Replace(v, "[1]", to_Chr(185))
    v = Replace(v, "[3/4]", to_Chr(190))
    v = Replace(v, "[3]", to_Chr(179))
    v = ReplaceX(v, "[c]", to_Chr(169))
    v = ReplaceX(v, "[cents]", to_Chr(162))
    v = ReplaceX(v, "[deg]", to_Chr(176))
    v = ReplaceX(v, "[div]", to_Chr(247))
    v = ReplaceX(v, "[plusminus]", to_Chr(177))
    v = ReplaceX(v, "[pounds]", to_Chr(163))
    v = ReplaceX(v, "[r]", to_Chr(174))
    v = ReplaceX(v, "[tm]", to_Chr(153))
    FormatTxt = v
End Function

Sub GetCondition(valChoose, valOperator, valKeyWord)
    If IsBlank(to_Str(valChoose)) Then valChoose = "Choose"
    If IsBlank(to_Str(valOperator)) Then valOperator = "Operator"
    If IsBlank(to_Str(valKeyWord)) Then valKeyWord = "KeyWord"
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
                            oDbPager.AddCondition Str4Sql(Replace(aChoose(x), "[int]", "")) & " " & aOperator(x) & " " & to_Int(aKeyWord(x)) & ""
                        Else
                            oDbPager.AddCondition Str4Sql(aChoose(x)) & " " & aOperator(x) & " '" & Str4Sql(to_Str(aKeyWord(x))) & "'"
                        End If
                    Case Else
                        If InStr(aChoose(x), "[int]")>0 Then
                            oDbPager.AddCondition " " & Str4Sql(Replace(aChoose(x), "[int]", "")) & " like '%" & to_Int(aKeyWord(x)) & "%'"
                        Else
                            If LCase(aKeyWord(x)) = "null" Then
                                oDbPager.AddCondition "(" & Str4Sql(aChoose(x)) & " is null Or " & Str4Sql(aChoose(x)) & " = '')"
                            Else
                                oDbPager.AddCondition Str4Sql(aChoose(x)) & " like '%" & Str4Like(to_Str(aKeyWord(x))) & "%'"
                            End If
                        End If
                End Select
            End If
        Next
    End If
End Sub

Public Function getFileName(v)
	v = to_String(v)
	v = Replace(v, "\", "/")
	aFileName = Split(v, "/")
	v = aFileName(UBound(aFileName))
	v = Replace(v, "<>", "")
	v = Replace(v, "|", "")
	v = Replace(v, ":", "")
	v = Replace(v, """", "")
	v = Replace(v, "*", "")
	v = Replace(v, "?", "")
	v = Replace(v, "/", "")
	v = Replace(v, "\", "")
	v = Replace(v, vbCr, "")
	v = Replace(v, vbLf, "")
	v = Replace(v, Chr(0), "")
	getFileName = v
End Function

'功能:获取全部图片地址,保存到一个数组.
Public Function getIMG(sString)
    Dim sReallyDo, regEx, iReallyDo
    Dim oMatches, cMatch
    '//定义空数组
    iReallyDo = -1
    ReDim aReallyDo(iReallyDo)
    If IsNull(sString) Then
        getIMG = aReallyDo
        Exit Function
    End If
    sReallyDo = sString
    '//格式化HTML代码
    sReallyDo = Replace(sReallyDo, vbCr, " ")
    sReallyDo = Replace(sReallyDo, vbLf, " ")
    sReallyDo = Replace(sReallyDo, vbTab, " ")
    sReallyDo = ReplaceX(sReallyDo, "<img", vbCrLf & "<img")
    sReallyDo = Replace(sReallyDo, "/>", " />" & vbCrLf)
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    '//正则匹配图片SRC地址
    regEx.Pattern = "<img.*?\ssrc\s*?=\s*?(([^\""\'\s][^\""\'\s>]*)|([\""\'])([^\""\']+?)\3)"
    Set oMatches = regEx.Execute(sReallyDo)
    '//将图片地址存入数组
    For Each cMatch in oMatches
        iReallyDo = iReallyDo + 1
        ReDim Preserve aReallyDo(iReallyDo)
        aReallyDo(iReallyDo) = regEx.Replace(cMatch.Value, "$2$4")
    Next
    getIMG = aReallyDo
End Function

'功能:如果不能取客户端真实IP，就会取客户端的代理IP。
Public Function getIP()
    Dim sIPAddress, sHTTP_X_FORWARDED_FOR
    sHTTP_X_FORWARDED_FOR = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If sHTTP_X_FORWARDED_FOR = "" Or InStr(sHTTP_X_FORWARDED_FOR, "unknown") > 0 Then
        sIPAddress = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(sHTTP_X_FORWARDED_FOR, ",") > 0 Then
        sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ",") -1)
    ElseIf InStr(sHTTP_X_FORWARDED_FOR, ";") > 0 Then
        sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ";") -1)
    Else
        sIPAddress = sHTTP_X_FORWARDED_FOR
    End If
    getIP = Replace(Mid(sIPAddress, 1, 30), " ", "")
End Function

'功能:就是把ADO的GetRows()封装了一下
Public Function GetRsRows(oRs)
    On Error Resume Next
    Dim aArray
    ReDim aArray(oRs.Fields.Count -1, -1)
    If TypeName(oRs) = "Recordset" Or TypeName(oRs) = "IOraDynaset" Then
        If Not oRs.EOF Then aArray = oRs.GetRows()
    End If
    GetRsRows = aArray
End Function

Public Function getUrl(v)
	If Not IsArray(v) Then v = Split(v, ",")
	
End Function

Function HTMLDecode(v)
    Dim I
	v = to_String(v)
    v = Replace(v, "&quot;", Chr(34))
    v = Replace(v, "&lt;" , Chr(60))
    v = Replace(v, "&gt;" , Chr(62))
    v = Replace(v, "&nbsp;", Chr(32))
    For I = 1 To 255
        v = Replace(v, "&#" & I & ";", Chr(I))
    Next
    v = Replace(v, "&amp;" , Chr(38))
    v = Replace(v, "<br>" , vbCrLf)
    v = Replace(v, "<br />" , vbCrLf)
    HTMLDecode = v
End Function

'过滤HTML各种标签样式脚本
Public Function HTMLFilter(sHTML, sFilters)
    If to_Str(sHTML) = "" Then Exit Function
    If to_Str(sFilters) = "" Then sFilters = "JORKIN,SCRIPT,OBJECT"
    Dim aFilters : aFilters = Split(UCase(sFilters), ",")
    For i = 0 To UBound(aFilters)
        Select Case UCase(Trim(aFilters(i)))
            Case "JORKIN"
                While InStr(sHTML, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") >0
                    sHTML = Replace(sHTML, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "&nbsp;&nbsp;&nbsp;&nbsp;")
                Wend
            Case "SCRIPT"
                '// 去除脚本<scr ipt></scr ipt>及 onload 等
                sHTML = RegReplace(sHTML, "<SCRIPT[\s\S]*?</SCRIPT>", "")
                sHTML = RegReplace(sHTML, "(JAVASCRIPT|JSCRIPT|VBSCRIPT|VBS):", "$1：")
                sHTML = RegReplace(sHTML, "\s[on].+?=\s+?([\""|\'])(.*?)\1", "")
            Case "FIXIMG"
                sHTML = RegReplace(sHTML, "<IMG.*?\sSRC=([^\""\'\s][^\""\'\s>]*).*?>", "<img src=$2 border=0>")
                sHTML = RegReplace(sHTML, "<IMG.*SRC=([\""\']?)(.\1\S+).*?>", "<img src=$2 border=0>")
            Case "TABLE"
                '// 去除表格<table><tr><td><th>
                sHTML = RegReplace(sHTML, "</?TABLE[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?TBODY[^>]*>", "")
                sHTML = RegReplace(sHTML, "<(/?)TR[^>]*>", "{1p>")
                sHTML = RegReplace(sHTML, "</?TH[^>]*>", " ")
                sHTML = RegReplace(sHTML, "</?TD[^>]*>", " ")
            Case "CLASS"
                '// 去除样式类class=""
                sHTML = RegReplace(sHTML, "(<[^>]+) CLASS=[^ |^>]+([^>]*>)", "$1 $2")
                sHTML = RegReplace(sHTML, "\sCLASS\s*?=\s*?([\""|\'])(.*?)\1", "")
            Case "STYLE"
                '// 去除样式style=""
                sHTML = RegReplace(sHTML, "(<[^>]+) STYLE=[^ |^>]+([^>]*>)", "$1 $2")
                sHTML = RegReplace(sHTML, "\sSTYLE\s*?=\s*?([\""|\'])(.*?)\1", "")
            Case "XML"
                '// 去除XML<?xml>
                sHTML = RegReplace(sHTML, "<\\?XML[^>]*>", "")
            Case "NAMESPACE"
                '// 去除命名空间<o:p></o:p>
                sHTML = RegReplace(sHTML, "<\/?[a-z]+:[^>]*>", "")
            Case "FONT"
                '// 去除字体<font></font>
                sHTML = RegReplace(sHTML, "</?FONT[^>]*>", "")
            Case "MARQUEE"
                '// 去除字幕<marquee></marquee>
                sHTML = RegReplace(sHTML, "</?MARQUEE[^>]*>", "")
            Case "OBJECT"
                '// 去除对象<object><param><embed></object>
                sHTML = RegReplace(sHTML, "</?OBJECT[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?PARAM[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?EMBED[^>]*>", "")
            Case "COMMENT"
                '// 去除HTML注释, 会处理<script>和<style>内注释, 慎用
                sHTML = RegReplace(sHTML, "<!--[\s\S]*?-->", "")
            Case Else
                '// 去除其它标签
                sHTML = RegReplace(sHTML, "</?" & aFilters(i) & "[^>]*?>", "")
        End Select
    Next
    HTMLFilter = sHTML
End Function

'功能:ASP里的IIF
Public Function IIf(bExp1, sVal1, sVal2)
	On Error Resume Next
    If (bExp1) Then
        IIf = sVal1
    Else
        IIf = sVal2
    End If
End Function

'功能:判断一个值是否存在于数组
Function InArray(sValue, aArray, bCompare)
    Dim i
    InArray = False
	bCompare = to_Bit(bCompare)
    For i = 0 To UBound(aArray)
		If IsObject(sValue) Then
			If aArray(i) Is sValue Then InArray = True
		Else
			If StrComp(aArray(i), sValue, bCompare) = 0 Then InArray = True
		End If
		If InArray Then Exit For
    Next
End Function


'功能:判断一个变量是否在于一个二维数据的某列
Public Function InArray2(sValue, aArray, iColumn, bCompare)
    On Error Resume Next
    Dim i, j, k
    InArray2 = False
    i = to_Int(iColumn)
	If Not IsArray(aArray) Then aArray = Array(aArray)
	bCompare = to_Bit(bCompare)
    If (i < 0 Or i > UBound(A)) Then
		For k = 0 To UBound(A)
			For j = 0 To UBound(A, 2)
				If StrComp(sValue, A(k, j), bCompare) = 0 Then
					echo(A(k, j) & "<br />")
					InArray2 = True
					Exit Function
				End If
			Next
		Next
	Else
		For j = 0 To UBound(A, 2)
			If StrComp(sValue, A(i, j), bCompare) = 0 Then
				echo(A(i, j) & "<br />")
				InArray2 = True
				Exit Function
			End If
		Next
    End If
End Function

Function Include(v)
	Dim sPATH_INFO
    sPATH_INFO = Request.ServerVariables("PATH_INFO")
    sPATH_INFO = Left(sPATH_INFO, InStrRev(sPATH_INFO, "/"))
    Include = IncludeFile(v, sPATH_INFO)
    ExecuteGlobal(Include)
    'Trace Include
End Function

Function IncludeFile(v, p)
    Dim oFSO, oFile, sInclude
    IncludeFile = ""
    If IsBlank(v) Then Exit Function
    If InStr(v, ":") = 2 Then
        v = Replace(v, "/", "\")
        If Right(v, 1)<> "\" Then v = v & "\"
        p = Left(v, InStrRev(v, "\"))
    Else
        v = Replace(v, "\", "/")
        v = ReplaceAll(v, "...", "../..", False)
        If InStr(v, "/") = 1 Then
            p = Left(v, InStrRev(v, "/"))
        Else
            p = p & Left(v, InStrRev(v, "/"))
        End If
		v = Server.MapPath(v)
    End If
    Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    If oFSO.FileExists(v) Then
        Set oFile = oFSO.OpenTextFile(v)
        If Not oFile.AtEndOfStream Then
            sInclude = to_Str(oFile.ReadAll)
            If Not IsBlank(sInclude) Then IncludeFile = IncludeASP(sInclude, p)
        End If
        Set oFile = Nothing
	Else
		'Trace Array(v, p)
    End If
End Function

Function IncludeASP(v, p)
    Dim aInclude, sTmpString, aTrimX, sInclude, i, k
    aTrimX = Array(" ", vbCr, vbLf, vbTab)
    IncludeASP = ""
    sInclude = v
    If Not IsBlank(sInclude) Then
        aInclude = Split(sInclude, "%" & ">")
        If UBound(aInclude) = 0 Then
            IncludeASP = IncludeASP & IncludeHTML(sInclude, p)
        Else
            For i = 0 To UBound(aInclude) -1
                k = InStr(aInclude(i), "<" & "%")
                sTmpString = TrimX(Left(aInclude(i), k - 1), aTrimX)
                If Not IsBlank(sTmpString) Then IncludeASP = IncludeASP & IncludeHTML(sTmpString, p)
                sTmpString = Mid(aInclude(i), k + 2)
                IncludeASP = IncludeASP & vbCrLf
                Select Case Left(TrimX(sTmpString, aTrimX), 1)
                    Case "@"
                    Case "="
                        sTmpString = Mid(sTmpString, 2)
                        IncludeASP = IncludeASP & "Response.Write(" & sTmpString & ") '//="
                    Case Else
                        If Not IsBlank(sTmpString) Then IncludeASP = IncludeASP & sTmpString & " '//~"
                End Select
            Next
            IncludeASP = IncludeASP & IncludeHTML(aInclude(i), p)
        End If
    End If
End Function

Function IncludeHTML(v, p)
    Dim oRegEx, oMatches, i, j, k
	Dim sTmpString, sIncludeString, sIncludeFileName, sIncludeFileType, sIncludePath
    sTmpString = v
    Set oRegEx = New RegExp
    With oRegEx
        .IgnoreCase = True
        .Global = True
        .Pattern = "(<!--.*#include.+(file|virtual)\s*=\s*)(""|')(.+?)(\3.*?-->)"
        .Multiline = False
        Set oMatches = .Execute(sTmpString)
        If oMatches.Count > 0 Then
            For i = 0 To oMatches.Count - 1
                sIncludeString = to_Str(oMatches(i))
                sIncludeFileName = Replace(to_Str(oMatches(i).SubMatches(3)), "\", "/")
                sIncludeFileType = LCase(to_Str(oMatches(i).SubMatches(1)))
                Select Case sIncludeFileType
                    Case "file"
                        If Left(sIncludeFileName, 1) = "/" Then
                            sIncludeFileName = ""
                        Else
                            sIncludeFileName = p & sIncludeFileName
                        End If
                    Case "virtual"
                        If Left(sIncludeFileName, 1) <> "/" Then sIncludeFileName = "/" & sIncludeFileName
                End Select
                sIncludePath = Left(sIncludeFileName, InStrRev(sIncludeFileName, "/"))
                If Not IsBlank(sIncludeFileName) Then
                    sTmpString = Replace(sTmpString, sIncludeString, "<" & "%" & vbCrLf & IncludeFile(sIncludeFileName, "./") & vbCrLf & "%" & ">")
                End If
            Next
            IncludeHTML = IncludeASP(sTmpString, p)
        Else
            sTmpString = Replace(sTmpString, """", """""")
            sTmpString = Replace(sTmpString, vbCrLf, """ & vbCrLf & """)
            If Not IsBlank(sTmpString) Then IncludeHTML = vbCrLf & "Response.Write(vbCrlf & """ & sTmpString & """)" & " '//HTML"
        End If
    End With
    Set oRegEx = Nothing
End Function

'功能:将字符串中每个单词的首字母都变为大写
Public Function InitCap(byVal sString)
    InitCap = InitCap2(sString, Array(vbCrlf, vbTab, "(", ")", ",", " ", "_", ".", "!", ";"))
	InitCap = Replace(InitCap,"_id","_ID",1,-1,1)
End Function

Public Function InitCap2(sString, aSplit)
    Dim i, j
    InitCap2 = to_String(sString)
	If Not IsArray(aSplit) Then aSplit = Array(aSplit)
	For j = 0 to UBound(aSplit)
		sSplit = aSplit(j)
		If Not IsEmpty(sSplit) Then
		    aStr = Split(InitCap2, sSplit)
		    For i = 0 To UBound(aStr)
		        aStr(i) = UCase(Left(aStr(i), 1)) & Mid(aStr(i), 2)
		    Next
		    InitCap2 = Join(aStr, sSplit)
	    End If
	Next
End Function

'功能:返回 Boolean 值指明表达式的值是否为字母。
Public Function IsAlpha(byVal sString)
    Dim regExp, oMatch, i, sStr
    For i = 1 To Len(to_Str(sString))
        sStr = Mid(sString, i, 1)
        Set regExp = New RegExp
        regExp.Global = True
        regExp.IgnoreCase = True
        regExp.Pattern = "[A-Z]|[a-z]|\s|[_]"
        Set oMatch = regExp.Execute(sStr)
        If oMatch.Count = 0 Then
            IsAlpha = False
            Exit Function
        End If
        Set regExp = Nothing
    Next
    IsAlpha = True
End Function

'功能:判断是否是空值
Public Function IsBlank(v)
    On Error Resume Next
    IsBlank = False
    Select Case VarType(v)
        Case 0, 1
            IsBlank = True
        Case 8
            If Len(v) = 0 Then IsBlank = True
        Case 9
            Select Case TypeName(v)
                Case "Nothing", "Empty"
                    IsBlank = True
                Case "RecordSet"
                    If v.State = 0 Then
                        IsBlank = True
                    Else
                        If v.Bof And v.EOF Then IsBlank = True
                    End If
                Case "Connection"
                    If v.State = 0 Then IsBlank = True
                Case "Dictionary"
                    If v.Count = 0 Then IsBlank = True
            End Select
        Case 8192, 8204, 8209
            If UBound(v) = -1 Then IsBlank = True
    End Select
End Function

'判断是否为数字
Function IsNaN(byval n)
    On Error Resume Next
    Dim d
    IsNaN = False
    If IsNumeric(n) Then
        d = CDbl(n)
        If Err.Number <> 0 Then IsNaN = True
    End If
End Function

'功能:检查是否存在系统组件或组件是否安装成功
Public Function IsObjInstalled(v)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(v)
    If 0 = Err Then IsObjInstalled = True
	If -2147352567 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

'功能:检查是否为Form的Post
Public Function IsPostBack()
    IsPostBack = False
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then IsPostBack = True
End Function

'功能:检查Email格式
Public Function IsValidEmail(email)
    Dim names, Name, i, c
    IsValidEmail = True
    names = Split(email, "@")
    If UBound(names) <> 1 Then
        IsValidEmail = False
        Exit Function
    End If
    For Each Name In names
        If Len(Name) <= 0 Then
            IsValidEmail = False
            Exit Function
        End If
        For i = 1 To Len(Name)
            c = LCase(Mid(Name, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                IsValidEmail = False
                Exit Function
            End If
        Next
        If Left(Name, 1) = "." Or Right(Name, 1) = "." Then
            IsValidEmail = False
            Exit Function
        End If
    Next
    If InStr(names(1), ".") <= 0 Then
        IsValidEmail = False
        Exit Function
    End If
    i = Len(names(1)) - InStrRev(names(1), ".")
    If i <> 2 And i <> 3 Then
        IsValidEmail = False
        Exit Function
    End If
    If InStr(email, "..") > 0 Then
        IsValidEmail = False
    End If
End Function

'80040e14/内存溢出编码
Function JEncode(byVal iStr)
    Dim i, aKatakana80040e14
    JEncode = to_String(iStr)
	aKatakana80040e14 = Array(ChrW(12468), ChrW(12460), ChrW(12462), ChrW(12464), ChrW(12466), ChrW(12470), ChrW(12472), ChrW(12474), ChrW(12485), ChrW(12487), ChrW(12489), ChrW(12509), ChrW(12505), ChrW(12503), ChrW(12499), ChrW(12497), ChrW(12532), ChrW(12508), ChrW(12506), ChrW(12502), ChrW(12500), ChrW(12496), ChrW(12482), ChrW(12480), ChrW(12478), ChrW(12476))
    For i = 0 To UBound(aKatakana80040e14)
        JEncode = Replace(JEncode, aKatakana80040e14(i), "{$Jn" & i & "}")
    Next
End Function

'80040e14/内存溢出解码
Function JDecode(byVal iStr)
    Dim i, aKatakana80040e14
    JDecode = to_String(iStr)
	aKatakana80040e14 = Array(ChrW(12468), ChrW(12460), ChrW(12462), ChrW(12464), ChrW(12466), ChrW(12470), ChrW(12472), ChrW(12474), ChrW(12485), ChrW(12487), ChrW(12489), ChrW(12509), ChrW(12505), ChrW(12503), ChrW(12499), ChrW(12497), ChrW(12532), ChrW(12508), ChrW(12506), ChrW(12502), ChrW(12500), ChrW(12496), ChrW(12482), ChrW(12480), ChrW(12478), ChrW(12476))
    For i = 0 To UBound(aKatakana80040e14)
        JDecode = Replace(JDecode, "{$Jn" & i & "}", aKatakana80040e14(i))
    Next
End Function

'功能:在一个字符串前面补全另一字符串
Public Function LFill(sString, sStr)
    Dim i, iStrLen : iStrLen = Len(to_Str(sStr))
    For i = iStrLen To 1 Step -1
        If Right(sStr, i ) = Left(sString, i ) Then Exit For
    Next
    LFill = Left(sStr, iStrLen - i) & sString
End Function

'功能:在一个字符串后面补全另一字符串
Public Function RFill(sString, sStr)
    Dim i, iStrLen : iStrLen = Len(to_Str(sStr))
    For i = iStrLen To 1 Step -1
        If Left(sStr, i) = Right(sString, i) Then Exit For
    Next
    RFill = sString & Mid(sStr, i + 1)
End Function


'功能:读取Form值，生成相应代码。
Public Function OPS()
    Dim Key, sTableName
    echo("<div style=""text-align:left;"">")
    echo("sID = to_Int(Request.QueryString(""ID""))<br /><br />")
    echo("If isPostBack() And ChkPost() And Request.Form(""editinfo"")=""editinfo"" Then<br />")
    echo("&nbsp; &nbsp; OpenConn()<br /><br />")
    For Each Key in Request.Form
        echo("&nbsp; &nbsp; s" & InitCap2(Key, "_") & " = " & IIf(IsNumeric(Request.Form(Key)), "to_Int", "to_Str") & "(Request.Form("""&Key&"""))<br />")
    Next
    echo("<br />")
    For Each Key in Request.Form
        echo("&nbsp; &nbsp; If IsBlank(s" & InitCap2(Key, "_") & ") Then doAlert """ & Key & "不能为空。"",""back""<br />")
    Next
    For Each Key in Request.Form
        If InStr(Key, "_") > 0 Then
            sTableName = InitCap(Split(Key, "_")(0))
            Exit For
        End If
    Next
    echo("&nbsp; &nbsp; sSql = ""Select * From " & sTableName & " Where " & sTableName & "_ID = "" & sID" & "<br />")
    echo("&nbsp; &nbsp; Set oRs = Server.CreateObject(""ADODB.RecordSet"")<br />")
    echo("&nbsp; &nbsp; oRs.Open sSql, oConn, 1, 3<br />")
    echo("&nbsp; &nbsp; If oRs.EOF Then<br />")
    echo("&nbsp; &nbsp; &nbsp; &nbsp; oRs.AddNew<br />")
    echo("&nbsp; &nbsp; End If<br />")
    For Each Key in Request.Form
        echo("&nbsp; &nbsp; oRs(""" & InitCap2(Key, "_") & """) = s" & InitCap2(Key, "_") & "<br />")
    Next
    echo("&nbsp; &nbsp; oRs.Update<br />")
    echo("&nbsp; &nbsp; oRs.Close<br />")
    echo("&nbsp; &nbsp; doAlert ""保存成功"",""./""<br />")
    echo("End If<br /><br /><br />")
    echo("If sId > 0 Then<br />")
    echo("&nbsp; &nbsp; ShowInfo sId<br />")
    echo("End If<br /><br />")
    echo("Function ShowInfo(sID)<br />")
    echo("&nbsp; &nbsp; echo(""&lt;scr""&""ipt language=""""javascript""""&gt;"")<br />")
    echo("&nbsp; &nbsp; echo(""var oForm = form1;"")<br />")
    echo("&nbsp; &nbsp; echo(""with(oForm){"")<br />")
    echo("&nbsp; &nbsp; Set oRs = Exec(""SELECT * FROM [Kin_" & sTableName & "] WHERE " & sTableName & "_Id = "" & sId)<br />")
    echo("&nbsp; &nbsp; If Not oRs.Eof Then <br /><br />")
    For Each Key in Request.Form
        echo("&nbsp; &nbsp; &nbsp; &nbsp; s" & InitCap2(Key, "_") & " = " & IIf(IsNumeric(Request.Form(Key)), "to_Int", "to_Str") & "(oRs(""" & InitCap2(Key, "_") & """))<br />")
    Next
    echo("<br />")
    For Each Key in Request.Form
        echo("&nbsp; &nbsp; &nbsp; &nbsp; echo(""" & Key & ".value="" & Str4Js(s" & InitCap2(Key, "_") & ") & "";"")<br />")
    Next
    echo("<br />")
    echo("&nbsp; &nbsp; End If<br />")
    echo("&nbsp; &nbsp; echo(""}"")<br />")
    echo("&nbsp; &nbsp; echo(""&lt;/scr""&""ipt&gt;"")<br />")
    echo("End Func" & "tion<br />")
    echo("<br />")
    echo("</div>")
    die()
End Function

Public Function MapPathEx(ByVal sMapPath)
    On Error Resume Next
    If IsBlank(sMapPath) Then sMapPath = "./"
    If Instr(sMapPath, ":") > 0 Then sMapPath = sMapPath & "\"
    sMapPath = Replace(sMapPath, "/", "\")
    sMapPath = ReplaceAll(sMapPath, "\\", "\", False)
    sMapPath = ReplaceAll(sMapPath, "...", "..\..", False)
    If (InStr(sMapPath, ":") > 0) Then
        If (Right(sMapPath, 2) <> ":\") And (Right(sMapPath, 1) = "\") Then
        	sMapPath = Mid(sMapPath, 1, Len(sMapPath) -1)
        End If
    Else
        sMapPath = Server.MapPath(sMapPath)
    End If
    MapPathEx = sMapPath
End Function

Public Function UnMapPath(byVal sRelPath)
    Dim sRootPath, aMapPath
    sRootPath = Server.Mappath("/")
	sFolderPath = Server.Mappath("./")
    sTranslatedPath = MapPathEx(sRelPath)
    If InStr(1, sTranslatedPath, sRootPath, 1) > 0 Then
	    UnMapPath = Replace(LCase(sTranslatedPath), LCase(sRootPath), "" )
	    UnMapPath = Right(sTranslatedPath, Len(UnMapPath))
	ElseIf InStr(1, sFolderPath, sTranslatedPath, 1) > 0 Then
		UnMapPath = Replace(LCase(sFolderPath), LCase(sTranslatedPath), "" )
		aMapPath = Split(UnMapPath, "\")
		UnMapPath = Repeat(UBound(aMapPath), "../")
    Else
	    aMapPath = Split(sTranslatedPath, "\")
		bMapPath = False
	    For i = Ubound(aMapPath) To 1 Step -1
	    	UnMapPath = "/" & aMapPath(i) & UnMapPath
	    	If MapPathEx(UnMapPath) = sTranslatedPath Then
				bMapPath = True
				Exit For
			End If
	    Next
		If Not bMapPath Then
			aFolderMapPath = Split(sFolderPath, "\")
			UnMapPath = ""
			For i = 0 To Ubound(aFolderMapPath)
				UnMapPath = UnMapPath & aFolderMapPath(i) & "\"
				If InStr(1, sTranslatedPath, UnMapPath, 1) = 0 Then
					i = i - 1
					Exit For
				End If
			Next
			UnMapPath = Repeat(UBound(aFolderMapPath) - i , "../")
			ReDim Preserve aFolderMapPath(i)
			UnMapPath = UnMapPath & ReplaceX(sTranslatedPath, Join(aFolderMapPath, "\") & "\", "")
		End If
    End IF
    UnMapPath = Replace( UnMapPath, "\", "/" )
	If IsBlank(UnMapPath) Then UnMapPath = "/"
End Function

'功能:真正实现ACCESS随机选取记录功能
Function NewID(PKey)
    'NewID = " Sin(" & Timer & "*" & PKey & ") "
    NewID = " Rnd(-" & Timer & "*" & PKey & ") "
End Function

Function NL2Br(sString)
	NL2Br = to_String(sString)
	NL2Br = Replace(NL2Br,vbNewLine,"<br />")
	NL2Br = Replace(NL2Br,vbCr,"<br />")
End Function

'功能:随机生成指定长度字符串
Public Function Rand(iLength, FromWords)
    Dim i, j, k, a(3)
    Dim p
    a(0) = "123456789"
    a(1) = "abcdefghijklmnopqrstuvwxyz"
    a(2) = UCase(a(1))
    a(3) = Left(a(2), 6)
    FromWords = to_Str(FromWords)
    If Len(FromWords) = 0 Then
        Rand = ""
        Exit Function
    End If
    Select Case FromWords
        Case "alpha"
            FromWords = a(1) & a(2)
        Case "0-9"
            FromWords = "0" & a(0)
        Case "1-9"
            FromWords = a(0)
        Case "a-z"
            FromWords = a(1)
        Case "A-Z"
            FromWords = a(2)
        Case "word"
            FromWords = "_" & a(0) & a(1) & a(2)
        Case "hex"
            FromWords = "0" & a(0) & a(3)
    End Select
    Randomize()
    k = Len(FromWords)
    For i = 1 To to_Int(iLength)
        j = to_Int(k * Rnd()) + 1
        Rand = Rand & Mid(FromWords, j, 1)
    Next
End Function

'功能:生成一个0到MaxNumber的数字
Public Function RandNumber(MaxNumber)
    Randomize()
    RandNumber = to_Int((MaxNumber + 1) * Rnd())
End Function

'功能:使用正则表示式对字符串进行替换
Public Function RegReplace(Str, PatternStr, RepStr)
    Str = to_String(Str)
    If IsBlank(NewStr) Then Exit Function
    Dim regEx
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = PatternStr
    Str = regEx.Replace(Str, RepStr)
    RegReplace = Str
End Function

'功能:可以按前缀清理Application的东东
Public Function RemoveApplication(sString, WriteRemoveApplication)
    On Error Resume Next
    sString = to_Str(sString)
    Dim aApplicationArray, x, i, j
    i = -1
    ReDim aApplicationArray(i)
    For Each x in Application.Contents
        If Left(x, Len(sString)) = sString Then
            i = i + 1
            ReDim Preserve aApplicationArray(i)
            aApplicationArray(i) = x
        End If
    Next
    Application.Lock
    For j = 0 To i
        Application.Contents.Remove(aApplicationArray(j))
        If WriteRemoveApplication Then echo "<br />释放 <strong>" & aApplicationArray(j) & "</strong> 完成<br />"
    Next
    Application.UnLock
    If WriteRemoveApplication Then echo "<br />所有对象已经更新,共释放了 <strong>" & j & "</strong> 个缓存对象.<br />"
End Function

'功能:重复 N次 特定字符串
'来源:http://jorkin.reallydo.com/article.asp?id=413
'需要to_Int函数:http://jorkin.reallydo.com/article.asp?id=395

Public Function Repeat(nTimes, sStr)
    nTimes = to_Int(nTimes)
    sStr = to_Str(sStr)
    Repeat = Replace(Space(nTimes), Space(1), sStr)
End Function


'功能：返回字符串，其中指定数目的某子字符串 全部 被替换为另一个子字符串。
'来源：http://jorkin.reallydo.com/article.asp?id=406
'需要to_Int函数:http://jorkin.reallydo.com/article.asp?id=395

Public Function ReplaceAll(sExpression, sFind, sReplaceWith, bAll)
    If IsBlank(sFind) Then ReplaceAll = sExpression : Exit Function
	If InStr( 1, sReplaceWith, sFind, to_Bit(bAll)) > 0 Then bAll = False
	If to_Bool(bAll) Then
		While InStr( 1, sExpression, sFind, 1) > 0
			sExpression = ReplaceX(sExpression, sFind, sReplaceWith)
		Wend
	Else
		While InStr(sExpression, sFind) > 0
			sExpression = Replace(sExpression, sFind, sReplaceWith)
		Wend
	End If
    ReplaceAll = sExpression
End Function

'功能:去掉全部HTML标记(Jorkin加强版)
Public Function ReplaceHTML(Textstr)
    Dim sStr, regEx
    sStr = Textstr
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Multiline = True
    regEx.Pattern = "<!--[\s\S]*?-->" '//想用就把注释去掉
    sStr = regEx.Replace(sStr, "")
    regEx.Pattern = "<script[\s\S]*?</script>"
    sStr = regEx.Replace(sStr, "")
    regEx.Pattern = "<style[\s\S]*?</style>"
    sStr = regEx.Replace(sStr, "")
    regEx.Pattern = "\s[on].+?=([\""|\'])(.*?)\1"
    sStr = regEx.Replace(sStr, "")
    regEx.Pattern = "<(.[^>]*)>"
    sStr = regEx.Replace(sStr, "")
    Set regEx = Nothing
    ReplaceHTML = sStr
End Function

'功能:不区分大小写替换
Public Function ReplaceX(sExpression, sFind, sReplaceWith)
	ReplaceX = Replace(to_String(sExpression), to_String(sFind), to_String(sReplaceWith), 1, -1, 1)
End Function

Private Function Root()
    Root = Path("/")
End Function

Sub RE()
    Response.End()
End Sub

Sub RR(s)
    Response.Redirect(s)
End Sub

Function RQ(v)
	Select Case Server.URLEncode("麒麟")
	Case "%E9%BA%92%E9%BA%9F"
		Dim aQueryString, iQueryString, sQueryString, bQueryString, QUERY_STRING_ARRAY, i
		QUERY_STRING_ARRAY = Split(Request.ServerVariables("QUERY_STRING"), "&")
		bQueryString = False
		Trace QUERY_STRING_ARRAY
		For i = 0 To UBound(QUERY_STRING_ARRAY)
			If InStr(QUERY_STRING_ARRAY(i), "=") > 0 Then
				If StrComp(Split(QUERY_STRING_ARRAY(i), "=")(0), v, 1) = 0 Then
					sQueryString = Mid(QUERY_STRING_ARRAY(i), InStr(QUERY_STRING_ARRAY(i), "=") + 1)
					Trace Array(Split(QUERY_STRING_ARRAY(i), "=")(0), sQueryString,IsValidHex(sQueryString))
					bQueryString = Not IsValidHex(sQueryString)
					If bQueryString Then Exit For
				End If
			End If
		Next
		If bQueryString Then
			Trace "中文"
			Session.CodePage = 936
			Set RQ = Request.QueryString(v)
			Session.CodePage = 65001
		Else
			Trace "Server.URLEncode"
			Set RQ = Request.QueryString(v)
		End If
	Case Else
		Set RQ =  Request.QueryString(v)
	End Select
End Function

'功能:判断该计划任务是否到期,一般定时更新Application的时候使用.
'sInterval, iNumber, dStartTime参数同 DateAdd 函数的参数
'来源:http://jorkin.reallydo.com/article.asp?id=423
Public Function ScheduleTask(sTaskName, sInterval, iNumber, dStartTime)
    Dim sApplicationName, sLastUpdate, sNextUpdate
    Select Case UCase(sInterval)
        Case "YYYY", "Q", "M", "Y", "D", "W", "WW", "H", "N", "S"
            sApplicationName = "ScheduleTask_" & sTaskName & "_LastUpdate"
            sLastUpdate = Trim(Application(sApplicationName))
            dStartTime = to_Date(dStartTime)
            ScheduleTask = False
            If sLastUpdate = "" Then
                sLastUpdate = DateAdd(sInterval, to_Int(DateDiff(sInterval, dStartTime, Now()) / iNumber -1) * iNumber, dStartTime)
                Application(sApplicationName) = sLastUpdate
            End If
            sNextUpdate = DateAdd(sInterval, to_Int(DateDiff(sInterval, dStartTime, Now()) / iNumber) * iNumber, dStartTime)
            If Now() > sNextUpdate Then
                ScheduleTask = True
                Application(sApplicationName) = sNextUpdate
            End If
        Case Else
            ScheduleTask = False
    End Select
End Function

'显示Webdings
Sub ShowWebdings()
	Response.Write("<table border=""1"" class=""webdings"">")
	For i = 33 To 126
	    Response.Write("<tr><td>" & i & "</td><td>" & Chr(I) & "</td><td><font face=""webdings"">" & Chr(i) & "</font></td></tr>")
	Next
	Response.Write("</table><style>.webdings td{font-size:64px}</style>")
End Sub

'功能:对一个一维数组进行排序
Public Function SortArray(UnSortedArray)
    Dim Front, Back, Current
    Dim Temp, ArraySize
    ArraySize = UBound(UnSortedArray)
    For Front = 0 To ArraySize - 1
        Current = Front
        For Back = Front To ArraySize
            If UnSortedArray(Current) > UnSortedArray(Back) Then
                Current = Back
            End If
        Next
        Temp = UnSortedArray(Current)
        UnSortedArray(Current) = UnSortedArray(Front)
        UnSortedArray(Front) = Temp
    Next
    SortArray = UnSortedArray
End Function

Public Function Swap(ByRef a, ByRef b)
	Dim c
	c = a
	a = b
	b = c
End Function

'功能:定义转换为Javascript字符串输出
Public Function to_Js(sString, quote)
    On Error Resume Next
    Dim iUBound1, iUBound2, i, j
    If IsArray(sString) Then
        iUBound1 = UBound(sString,1)
        iUBound2 = UBound(sString,2)
        to_Js = "["
        If Err Then
            Err.Clear
            For i = 0 To UBound(sString)
                If i > 0 Then to_Js = to_Js & ","
                to_Js = to_Js & to_Js(sString(i), quote)
            Next
        Else
            For i = 0 To iUBound2
                If i > 0 Then to_Js = to_Js & ","
                to_Js =  to_Js & "["
                    For j = 0 To iUBound1
                        If j > 0 Then to_Js = to_Js & ","
                        to_Js = to_Js & to_Js(sString(j, i), quote)
                    Next
                to_Js =  to_Js & "]"
            Next
        End If
        to_Js = to_Js & "]"
    Else
        Select Case VarType(sString)
        Case 0
            to_Js = quote & "undefined" & quote
        Case 1
            to_Js = "null"
        Case 11
            If sString Then to_Js = "true" Else to_Js = "false"
        Case 2, 3, 4, 5
            to_Js = sString
		Case 7
			to_Js = "(new Date(" & Year(sString) & ", " & Month(sString) -1 & ", " & Day(sString) & ", " & Hour(sString) & ", " & Minute(sString) & ", " & Second(sString) & "))"
        Case 9
            If TypeName(sString) = "RecordSet" Then
            	to_Js = to_Js(sString.GetRows(), quote)
            Else
            	to_Js = to_Js("Object:" & TypeName(sString), quote)
            End If
        Case Else
            to_Js = quote & JsEncode(sString, quote) & quote
        End Select
    End If
End Function

Function jsEncode(str,quote)
    str = to_String(str)
    Dim charmap(127), haystack()
    charmap(8)  = "\b"
    charmap(9)  = "\t"
    charmap(10) = "\n"
    charmap(12) = "\f"
    charmap(13) = "\r"
    charmap(34) = "\""" '//约定Javascript中，用双引号引用字符串
    charmap(39) = "\'" '//约定Javascript中，用单引号引用字符串
    charmap(47) = "\/"
    charmap(92) = "\\"
	If Not IsBlank(quote) Then
		If ASCW(quote) = 34 Then charmap(39) = Empty
		If ASCW(quote) = 39 Then charmap(34) = Empty
	End If
    Dim strlen : strlen = Len(str)
    ReDim haystack(strlen)
    Dim i, charcode
    For i = 1 To strlen
        haystack(i) = Mid(str, i, 1)
        charcode = AscW(haystack(i))
		If charcode < 0 Then charcode = charcode + 65536
        If charcode < 127 Then
            If Not IsEmpty(charmap(charcode)) Then
                haystack(i) = charmap(charcode)
            ElseIf charcode < 32 Then
                haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
            End If
        Else
            haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
        End If
    Next
    jsEncode = Join(haystack, "")
End Function

'功能:定义转换为Javascript字符串输出(约定Javascript中，用单引号引用字符串)
Public Function Str4Js(sString)
    Str4Js = to_Js(sString,"'")
End Function

'功能:在用 SQL 模糊搜索 Like 之前套一下
Public Function Str4Like(byVal v)
    v = to_String(v)
    v = Replace(v, "'" , "''")
    v = Replace(v, "[" , "[[]")
    v = Replace(v, "%" , "[%]")
    v = Replace(v, "_" , "[_]")
    Str4Like = v
End Function

Public Function String4Like(v)
    v = Str4Like(v)
    v = Replace(v, "?", "_")
    v = Replace(v, "？", "_")
    v = Replace(v, "*", "%")
    v = Replace(v, "×", "%")
    v = Replace(v, "＊", "%")
    String4Like = v
End Function

Public Function Str4RegExp(v)
	v = to_String(v)
	v = Replace(v, "\", "\\")
	v = Replace(v, "$", "\$")
	v = Replace(v, "(", "\(")
	v = Replace(v, ")", "\)")
	v = Replace(v, "*", "\*")
	v = Replace(v, "+", "\+")
	v = Replace(v, ".", "\.")
	v = Replace(v, "[", "\[")
	v = Replace(v, "|", "\|")
	v = Replace(v, "{", "\{")
	v = Replace(v, "^", "\^")
	Str4RegExp = v
End Function


'功能:在用 SQL 的 "等于" 查询前过滤单引号
Public Function Str4Sql(v)
    Str4Sql = Replace(to_String(v), "'", "''")
End Function

Function Suicide(v)
	v = to_Str(v)
	If IsNull(v) Then v = Request.ServerVariables("SCRIPT_NAME")
	v = Server.MapPath(v) 
	Set oFSO = CreateObject("Scripting.FileSystemObject") 
	If oFSO.FileExists(f) Then oFSO.Deletefile(f)
	Set oFSO = Nothing 
End Function

'功能:生成多关键字查询SQL语句
Public Function Str4Search(sField, sString)
	If Not IsArray(sField) Then sField = Array(sField)
	Const sIgnore = "@#$^&()=+]\"
	Dim i, j, iAllWords, aAllWords(), iExactWording, aExactWording(), aString1, aString2
	Dim sTmpString, bStatus
	iAllWords = -1
	iExactWording = -1
	bStatus = 1
	aString1 = Split(sString, """")
	For i = 0 To UBound(aString1)
		If i Mod 2 = 0 Then
			sTmpString = to_Str(aString1(i))
			For j = 1 To Len(sIgnore)
				sTmpString = Replace(sTmpString, Mid(sIgnore, j, 1), "")
			Next
			sTmpString = to_Str(ReplaceAll(sTmpString, "  ", " ", True))
			aString2 = Split(sTmpString, " ")
			For j = 0 To UBound(aString2)
				Select Case True
				Case aString2(j) = "AND"
					bStatus = 1
				Case aString2(j) = "NOT"
					bStatus = -1
				Case aString2(j) = "OR"
					bStatus = 0
				Case Left(aString2(j),1) = "-" Or Left(aString2(j),1) = "–" Or bStatus = -1
					If bStatus = -1 Then
						sTmpString = aString2(j)
					Else
						sTmpString = Mid(aString2(j), 2)
					End If
					If Not IsBlank(sTmpString) Then
						iExactWording = iExactWording + 1
						ReDim PreServe aExactWording(iExactWording)
						aExactWordingFields = sField
						For k = 0 To UBound(aExactWordingFields)
							aExactWordingFields(k) = Replace("({$Fields} NOT LIKE '%" & String4Like(sTmpString) & "%')", "{$Fields}", aExactWordingFields(k))
						Next
						aExactWording(iExactWording) = Join(aExactWordingFields, " AND ")
					End If
					bStatus = 1
				Case Else
					iAllWords = iAllWords + 1
					If iAllWords > 0 Then Str4Search = Str4Search & IIf(bStatus = 0, " OR ", ") AND (")
					Str4Search = Str4Search & "{$Fields} LIKE '%" & String4Like(aString2(j)) & "%'"
					bStatus = 1
				End Select
			Next
		Else
			iAllWords = iAllWords + 1
			If iAllWords > 0 Then Str4Search = Str4Search & IIf(bStatus = 0, " OR ", ") AND (")
			Str4Search = Str4Search & "{$Fields} LIKE '%" & aString1(i) & "%'"
			bStatus = 1
		End If
	Next
	If Not IsBlank(Str4Search) Then
		Str4Search = "(" & Str4Search & ")"
		For i = 0 To UBound(sField)
			sField(i) = "(" & Replace(Str4Search, "{$Fields}", sField(i)) & ")"
		Next
		Str4Search = Join(sField, " OR ")
	End If
	If iExactWording > -1 Then
		If Not IsBlank(Str4Search) Then Str4Search = Str4Search & " AND "
		Str4Search = Str4Search & "(" & Join(aExactWording, " AND ") & ")" 
	End If
End Function

'带中文的字符串长度
Function StrLength(v)
    Dim sStr, i, iLength
    v = to_String(v)
    If Len("麒麟") = 2 Then
    	iLength = 0
        For i = 1 To Len(v)
            sStr = AscW(Mid(v, i, 1))
            If sStr < 0 Then sStr = sStr + 65536
            If sStr < 255 Then
                iLength = iLength + 1
            Else
                iLength = iLength + 2
            End If
        Next
    Else
        iLength = Len(Str)
    End If
    StrLength = iLength
End Function

Public Function StrToArray(v)
    Dim l, i
    l = Len(v) - 1
    Redim aArray(l)
    For i = 0 To l
        aArray(i) = Mid(v, i + 1, 1)
    Next
    StrToArray = aArray
End Function

'功能:调试输出变量值,支持记录集/字符串/一维数组/二维数组/各种服务器变量
Public Function Trace(ByVal s)
    'On Error Resume Next
	ExecuteGlobal "If Not TraceStyle Then" & vbCrlf & "Response.Write(""</scr" & "ipt></body><style>.tracediv{color:#000;font:14px;margin:0px;padding:0px;text-align:left; width:100%}.tracediv table,.tracediv th,.tracediv td,.tracediv hr,.tracediv fieldset,.tracediv legend{background:#CCE8CF;border-collapse:collapse;padding:3px;margin:3px;color:000;border:1px solid #820222}</style>"")" & vbCrlf & "TraceStyle = True" & vbCrlf & "End If"
    echo("<div class=""tracediv""><fieldset>")
    Dim i, j, k, iUBound1, iUBound2
	Dim sTypeName,sVarType
	sTypeName = TypeName(s)
	sVarType = VarType(s)
    If IsArray(s) Then
    	On Error Resume Next
        iUBound1 = UBound(s)
        iUBound2 = UBound(s, 2)
        If Err Then
			Err.Clear
            echo("<legend style=""color:red;"">Array1 :</legend><table>")
            echo("<tr><td>&#21015;</td><td>&#20540;</td></tr>")
            For i = 0 To iUBound1
                echo("<tr><td>" & i & "</td><td>")
                If IsArray(s(i)) Or IsObject(s(i)) Then
                    Trace(s(i))
                Else
                    echo(to_HTML(s(i)))
                End If
                echo("</td></tr>")
            Next
            echo("</table>")
        Else
            echo("<legend style=""color:red;"">Array2 :</legend><table>")
            echo("<tr><td>&#20108;&#32500;/&#19968;&#32500;</td>")
            For j = 0 To iUBound1
                echo("<td>" & j & "</td>")
            Next
            echo("</tr>")
            For i = 0 To iUBound2
                echo("<tr><td>" & i & "</td>")
                For j = 0 To iUBound1
                    echo("<td>")
                    If IsArray(s(j, i)) Or IsObject(s(j, i)) Then
						Trace(s(j, i))
					Else
                        echo(to_HTML(s(j, i)))
                    End If
                    echo("</td>")
                Next
                echo("</tr>")
            Next
            echo("</table>")
        End If
    ElseIf IsObject(s) Then
    	Select Case TypeName(s)
	    	Case "Recordset", "IOraDynaset"
		        echo("<legend style=""color:red;"">" & TypeName(s) & " :</legend>")
				Do Until s Is Nothing
					If s.State = 1 Then
						echo("<table><tr><th><font color=""red""><nobr>rownum<nobr></font></th>")
						For i = 0 To s.Fields.Count - 1
							echo("<th>" & s(i).Name & "</th>")
						Next
						Do Until s.EOF
							j = j + 1
							If j*i > 2925 Then Exit Do
							echo("<tr><td>" & j & "</td>")
							For i = 0 To s.Fields.Count - 1
								If IsNull(s(i)) Then
									echo("<td><font color=""red"">&lt;NULL&gt;</font></td>")
								ElseIf IsBlank(s(i)) Then
									echo("<td><font color=""blue"">&lt;BLANK&gt;</font></td>")
								Else
									echo("<td>" & to_HTML(s(i)) & "</td>")
								End If
							Next
							echo("</tr>")
							s.MoveNext
						Loop
						echo("</table>")
					Else
						echo("<table><tr><td><font color=""red"">[" & TypeName(s) & ".Closed]</font></td></tr></table>")
					End If
					On Error Resume Next
					Set s = s.NextRecordSet
					If Err Then
						Err.Clear
						Exit Do
					End If
					On Error Goto 0
				Loop
	    	Case "Errors"
				echo("<legend class=""tracediv"" style=""color:red;"">&#20849; " & s.Count & " &#20010;Errors&#21464;&#37327;</legend>")
				For i = 0 To s.Count -1
					echo("<fieldset><legend style=""color:red"">Errors.Item("& i &")</legend>")
					echo("<strong>Errors.Item("& i &").Description" & " = " & s.Item(i).Description & "</strong><br>")
					echo("<strong>Errors.Item("& i &").Number" & " = " & s.Item(i).Number & "</strong><br>")
					echo("<strong>Errors.Item("& i &").Source" & " = " & s.Item(i).Source & "</strong><br>")
					echo("<strong>Errors.Item("& i &").SQLState" & " = " & s.Item(i).SQLState & "</strong><br>")
					echo("<strong>Errors.Item("& i &").NativeError" & " = " & s.Item(i).NativeError & "</strong><br>")
					echo("</fieldset>")
				Next
			Case "Error2"
					echo("<fieldset><legend style=""color:red"">Error</legend>")
					echo("<strong>Errors.Item("& i &").Description" & " = " & s.Description & "</strong><br>")
					echo("<strong>Errors.Item("& i &").Number" & " = " & s.Number & "</strong><br>")
					echo("<strong>Errors.Item("& i &").Source" & " = " & s.Source & "</strong><br>")
					echo("<strong>Errors.Item("& i &").SQLState" & " = " & s.SQLState & "</strong><br>")
					echo("<strong>Errors.Item("& i &").NativeError" & " = " & s.NativeError & "</strong><br>")
					echo("</fieldset>")
			Case "IRequest"
				Trace("Request")
			Case "Dictionary"
				echo("<legend class=""tracediv"" style=""color:red;"">&#20849; " & s.Count & " &#20010;Dictionary&#21464;&#37327;</legend>")
				iUBound1 = s.Keys
				iUBound2 = s.Items
				echo("<table>")
				For i = 0 To s.Count -1
					echo("<tr><td><strong>Dictionary(""" & iUBound1(i) & """)" & "</strong></td><td>")
					Trace(iUBound2(i))
					echo("</td></tr>")
				Next
				echo("</table>")
			Case "Connection"
                echo("<legend class=""tracediv"" style=""font:bold;"">" & typename(s) & "</legend>")
				Trace to_Str(s)
				'If s.Errors.Count > 0 Then trace s.Errors
				trace s.Properties
			Case "Catalog"
				echo("<legend class=""tracediv"" style=""color:red;"">" & TypeName(s) & "</legend>")
				Set xCatUser = s.Users
				trace typename(xcatusers)
				For Each i In s.Users
				trace i
				Next
				Trace s.Tables
				Trace s.Groups
				Trace s.Users
				Trace s.Procedures
				Trace s.Views 
			Case "Tables"
				echo("<legend class=""tracediv"" style=""color:red;"">" & TypeName(s) & "</legend>")
				For i = 0 To s.Count -1
					If s.Item(i).Type = "TABLE" Or s.Item(i).Type = "VIEW" Then
						trace(s.Item(i))
					End If
				Next
			Case "Table"
					echo("<legend class=""tracediv"" style=""color:red;"">[" & s.Type & "] " & s.Name & "</legend>")
					trace(s.Columns)
			Case "DOMDocument"
 	    			echo("<legend class=""tracediv"" style=""color:red;"">DOMDocument</legend>")
					trace(s.xml)
 	    	Case Else
				If s Is Request.QueryString Then
					Trace("Request.QueryString")
				ElseIf s Is Request.Form Then
					Trace("Request.Form")
				ElseIf s Is Request.Cookies Then
					Trace("Request.Cookies")
				ElseIf s Is Application Then
					Trace("Application")
				ElseIf s Is Session Then
					Trace("Session")
				ElseIf s Is Request.ServerVariables Then
					Trace("REQUEST.SERVERVARIABLES")
				Else
					echo("<fieldset><legend style=""color:red"">" & TypeName(s) & "</legend>")
					echo(to_String(s))
					echo("</fieldset>")
				End If
    	End Select
    Else
        iUBound1 = UCase(s)
        Select Case iUBound1
            Case "APPLICATION"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Application.Contents.Count & " &#20010;Application&#21464;&#37327;</legend>")
                For Each i in Application.Contents
                    echo("<strong>Application(""" & i & """)" & " = </strong>")
                    Trace(Application(i))
                Next
            Case "COOKIES", "REQUEST.COOKIES"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.Cookies.Count & " &#20010;Request.Cookies&#21464;&#37327;</legend>")
                For Each i in Request.Cookies
                    echo("<strong>Request.Cookies(""" & i & """)" & " = </strong>")
                    If Request.Cookies(i).HasKeys Then
                        echo("<fieldset><legend style=""color:red;"">" & TypeName(Request.Cookies(i)) & " :</legend>")
                        echo("<strong>&#20849; " & Request.Cookies(i).Count & " &#20010;Request.Cookies(""" & i & """)&#23376; &#21464;&#37327;</strong><br />")
                        For Each j in Request.Cookies(i)
                            echo("Request.Cookies(""" & i & """)(""" & j & """) = ")
                            Trace(Request.Cookies(i)(j))
                        Next
                        echo("</fieldset>")
                    Else
                        Trace(Request.Cookies(i))
                    End If
                Next
            Case "SESSION"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Session.Contents.Count & " &#20010;Session&#21464;&#37327;</legend>")
                For Each i in Session.Contents
                    echo("<strong>Session(""" & i & """)" & " = </strong>")
                    Trace(Session(i))
                Next
            Case "QUERYSTRING", "REQUEST.QUERYSTRING"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.QueryString.Count & " &#20010;Request.QueryString&#21464;&#37327;</legend>")
                For Each i in Request.QueryString
                    echo("<strong>Request.QueryString(""" & i & """)" & " = </strong>")
                    For Each j In Request.QueryString(i)
                        Trace(j)
                    Next
                Next
            Case "FORM", "REQUEST.FORM"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.Form.Count & " &#20010;Request.Form&#21464;&#37327;</legend>")
                For Each i in Request.Form
                    echo("<strong>Request.Form(""" & i & """)" & " = </strong>")
                    For Each j In Request.Form(i)
                        Trace(j)
                    Next
                Next
            Case "SERVERVARIABLES", "REQUEST.SERVERVARIABLES", "REQUEST.SERVER"
                echo("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.ServerVariables.Count & " &#20010;Request.ServerVariables&#21464;&#37327;</legend>")
                For Each i in Request.ServerVariables
                    echo("<strong>Request.ServerVariables(""" & i & """)" & " = </strong>")
                    Trace(Request.ServerVariables(i))
                Next
			Case "REQUEST"
				Trace("Request.QueryString")
				Trace("Request.Form")
            Case Else
                echo("<legend style=""color:red;"">" & TypeName(s) & " :</legend>")
                If IsNull(s) Then
                    echo("<font color=""red"">IsNull</font>")
                ElseIf IsEmpty(s) Then
                    echo("<font color=""red"">IsEmpty</font>")
                ElseIf s = "" Then
                    echo("<font color=""blue"">IsBlank</font>")
                Else
                    echo(to_HTML(s))
                End If
        End Select
    End If
    echo("</fieldset></div>")
    Err.Clear()
    Response.Flush()
End Function

Function dTrace(v)
	Trace(v)
	die()
End Function

Public Function Trace2(s, v)
	echo("<div class=""tracediv""><fieldset>")
	echo("<legend style=""color:orange;"">" & to_String(s) & "</legend>")
	Trace v
	echo("</fieldset>")
	echo("</div>")
End Function

'仿ORACLE的Translate
Function Translate(str1, src, dest)
    If IsNull(str1) Then Exit Function
    Dim i
    src = to_String(src)
    dest = to_String(dest)
    For i = 1 To Len(src)
        str1 = Replace(str1, Mid(src, i, 1), Mid(dest, i, 1))
    Next
    Translate = str1
End Function

'功能:返回去除前sStr、后sStr 或前后sStr 的字符串sString副本。
Public Function TrimEx(sString, sStr, bCompare, sLeftRight)
	sString = to_String(sString)
	If IsBlank(sString) Then Exit Function
	If Not IsArray(sStr) Then sStr = Array(sStr)
	bCompare = to_Bit(bCompare)
	Dim iStrLen, bTrimNext, i
	bTrimNext = True
    Select Case UCase(sLeftRight)
        Case "L", "LEFT"
			While bTrimNext
				For i = 0 To UBound(sStr)
					If Not IsBlank(sStr(i)) Then
						iStrLen = Len(sStr(i))
						While StrComp(Left(sString, iStrLen), sStr(i), bCompare) = 0
							sString = Mid(sString, iStrLen + 1)
						Wend
					End If
				Next
				bTrimNext = False
				For i = 0 To UBound(sStr)
					If Not IsBlank(sStr(i)) Then
						iStrLen = Len(sStr(i))
						If StrComp(Left(sString, iStrLen), sStr(i), bCompare) = 0 Then
							bTrimNext = True
							Exit For
						End If
					End If
				Next
			Wend
        Case "R", "RIGHT"
			While bTrimNext
				For i = 0 To UBound(sStr)
					If Not IsBlank(sStr(i)) Then
						iStrLen = Len(sStr(i))
						While StrComp(Right(sString, iStrLen), sStr(i), bCompare) = 0
							sString = Mid(sString, 1, Len(sString) - iStrLen)
						Wend
					End If
				Next
				bTrimNext = False
				For i = 0 To UBound(sStr)
					If Not IsBlank(sStr(i)) Then
						iStrLen = Len(sStr(i))
						If StrComp(Right(sString, iStrLen), sStr(i), bCompare) = 0 Then
							bTrimNext = True
							Exit For
						End If
					End If
				Next
			Wend
        Case Else
            sString = TrimEx(sString, sStr, bCompare, "L")
            sString = TrimEx(sString, sStr, bCompare, "R")
    End Select
    TrimEx = sString
End Function

Public Function TrimL(sString, sStr)
	TrimL = TrimEx(sString, sStr, 0, "L")
End Function

Public Function TrimR(sString, sStr)
	TrimR = TrimEx(sString, sStr, 0, "R")
End Function

Public Function TrimX(sString, sStr)
	TrimX = TrimEx(sString, sStr, 0, Null)
End Function

Public Function UTF8URLEncode(v)
	Dim SessionCodepage
	v = to_String(v)
	If Server.URLEncode("麒麟") = "%E9%BA%92%E9%BA%9F" Then
		UTF8URLEncode = Server.URLEncode(v)
	Else
		SessionCodepage = Session.CodePage
		Session.CodePage = 65001
		UTF8URLEncode = Server.URLEncode(v)
		Session.CodePage = SessionCodePage
	End If
End Function

'功能:URLencode解码函数,可以解码生僻双字节文字
Public Function URLDecode(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = Eval("&h" + Mid(enStr, i + 1, 2))
            If v<128 Then
                deStr = deStr&Chr(v)
                i = i + 2
            Else
                If IsValidHex(Mid(enstr, i, 3)) Then
                    If IsValidHex(Mid(enstr, i + 3, 3)) Then
                        v = Eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr = deStr&Chr(v)
                        i = i + 5
                    Else
                        v = Eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr = deStr&Chr(v)
                        i = i + 3
                    End If
                Else
                    destr = destr&c
                End If
            End If
        Else
            If c = "+" Then
                deStr = deStr&" "
            Else
                deStr = deStr&c
            End If
        End If
    Next
    URLDecode = deStr
End Function

Public Function IsValidHex(v)
	Dim av, iv, i
	v = to_String(v)
	If InStr(v, "%") = 1 Then
		IsValidHex = True
		av = Split(v, "%")
		iv = UBound(av)
		For i = 1 To iv
			If Not IsNumeric("&H" & av(i)) OR Len(av(i)) <> 2 Then
				IsValidHex = False
				Exit Function
			End If
		Next
	Else
		IsValidHex = False
	End If
End Function

'功能:输出字符<br />
Public Function PrintLn(v)
    echo(v)
    echo("<br />")
End Function

'功能:计算目录绝对路径。
Function Path(ByVal sPath)
    Path = MapPathEx(sPath)
End Function

'判断s是否是有效数据
Public Function IsValid(s, av)
    Dim i
    IsValid = False
    If IsArray(av) Then
        For i = 0 To UBound(av)
            If StrComp(to_Str(s), to_Str(av(i)), 1) = 0 Then
                IsValid = True
                Exit Function
            End If
        Next
    Else
        IsValid = IsValid(s, Split(av, ","))
    End If
End Function

'根据字典取名称
Public Function GetDictMC(sDM, sTABLE, sDefault)
    Dim aDict, oDictRs, i, j
    ReDim aDict(0, -1)
    If Not IsArray(Application(sTable)) Then
        Set oDictRs = Exec("Select DM,MC From " & sTable)
        aDict = GetRsRows(oDictRs)
        Application.Lock()
        Application(sTable) = aDict
        Application.UnLock()
    Else
        aDict = Application(sTable)
        If UBound(aDict) <> 1 Then
            Application.Lock()
            Application(sTable) = Empty
            Application.UnLock()
            GetDictMC = GetDictMC(sDM, sTABLE, -1)
            Exit Function
        End If
    End If
    GetDictMC = sDefault
    For i = 0 To UBound(aDict, 2)
        If StrComp(sDM, aDict(0, i), 1) = 0 Then
            GetDictMC = aDict(1, i)
            Exit Function
        End If
    Next
End Function

'根据名称取字典
Public Function GetDictDM(sDM, sTABLE, sDefault)
    Dim aDict, oDictRs, i, j
    ReDim aDict(0, -1)
    If Not IsArray(Application(sTable)) Then
        Set oDictRs = Exec("Select DM,MC From " & sTable)
        aDict = GetRsRows(oDictRs)
        Application.Lock()
        Application(sTable) = aDict
        Application.UnLock()
    Else
        aDict = Application(sTable)
        If UBound(aDict) <> 1 Then
            Application.Lock()
            Application(sTable) = Empty
            Application.UnLock()
            GetDictDM = GetDictDM(sDM, sTABLE, -1)
            Exit Function
        End If
    End If
    GetDictDM = sDefault
    For i = 0 To UBound(aDict, 2)
        If StrComp(sDM, aDict(1, i), 1) = 0 Then
            GetDictDM = aDict(0, i)
            Exit Function
        End If
    Next
End Function

'功能:判断字典是否有效
Public Function IsValidDM(sDM, sTABLE)
    Dim aDict, oDictRs, i, j
    ReDim aDict(0, -1)
    If Not IsArray(Application(sTable)) Then
        Set oDictRs = Exec("Select DM,MC From " & sTable)
        aDict = GetRsRows(oDictRs)
        Application.Lock()
        Application(sTable) = aDict
        Application.UnLock()
    Else
        aDict = Application(sTable)
        If UBound(aDict) <> 1 Then
            Application.Lock()
            Application(sTable) = Empty
            Application.UnLock()
            IsValidDM = IsValidDM(sDM, sTABLE)
            Exit Function
        End If
    End If
    IsValidDM = False
    For i = 0 To UBound(aDict, 2)
        If StrComp(sDM, aDict(0, i), 1) = 0 Then
            IsValidDM = True
            Exit Function
        End If
    Next
End Function


'功能:二进制转十进制
Function CDec(expression)
    CDec = 0
    expression = StrReverse(expression)
    For i = 1 To Len(expression)
        CDec = CDec + Mid(expression, i, 1) * 2^(i -1)
    Next
End Function

'功能:十进制转二进制
Function Bin(expression)
    Bin = ""
    If Not IsNumeric(expression) Then Exit Function
    Do Until expression < 2
        Bin = expression - to_Int(expression / 2) * 2 & Bin
        expression = to_Int(expression / 2)
    Loop
    Bin = expression & Bin
End Function

'功能:MOD函数的增强版
Function XMOD(number1 , number2)
    On Error Resume Next
    XMOD = number1 Mod number2
    If Err Then
        XMOD = number1 - to_Int(number1 / 2) * 2
        Err.Clear
    End If
End Function

Private Function Tables(v)
    Dim oAdox, i, j, k, l
    Set oAdox = Server.CreateObject("ADOX.Catalog")
	If IsObject(oConn) Then
		oAdox.ActiveConnection = v.ConnectionString
	Else
		oAdox.ActiveConnection = v
	End If
	ReDim aTables(oAdox.tables.count)
	j = -1
    For i = 0 To oAdox.tables.count - 1
		If UCase(oAdox.tables(i).type) = "TABLE" Then
			j = j + 1
			aTables(j) = oAdox.tables(i).name
		End If
    Next
	ReDim Preserve aTables(j)
    Tables = aTables
    Set oAdox = Nothing
End Function

Private Function Procs(ByVal connstring)
    Dim adox, i, strProcs
    Set adox = Server.CreateObject("ADOX.Catalog")
    adox.ActiveConnection = connstring
    For i = 0 To adox.procedures.count - 1
        strProcs = strProcs & adox.procedures(i).name & vbCrLf
    Next
    Set adox = Nothing
    Procs = Split( strProcs, vbCrLf )
End Function


Private Sub SqlExec(ByVal ConnString, ByVal SQL)
    Dim objCn, bErr1, bErr2, strErrDesc
    On Error Resume Next
    Set objCn = Server.CreateObject("ADODB.Connection")
    objCn.Open ConnString
    If Err Then 
        bErr1 = True
    Else
        objCn.Execute SQL
        If Err Then 
            bErr2 = True
            strErrDesc = Err.description
        End If
    End If
    objCn.Close
    Set objCn = Nothing
    On Error GoTo 0
    If bErr1 Then
        Err.Raise 5109, "SqlExec Statement", "Bad connection " & _
                "string. Database cannot be accessed."
    ElseIf bErr2 Then
        Err.Raise 5109, "SqlExec Statement", strErrDesc
    End If
End Sub

Private Sub MkDatabase(byVal pathname)
    Dim objAccess, objFSO
    If LCase( Right( pathname, 4 ) ) <> ".mdb" Then
        Err.Raise 5155, "MkDatabase Statement", _
              "Database name must end with '.mdb'"
        Exit Sub
    End If

    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists( pathname ) Then
        Set objFSO = Nothing
        Err.Raise 5155, "MkDatabase Statement", _
              "Specified MS Access database already exists."
        Exit Sub
    End If
    Set objFSO = Nothing
    On Error Resume Next
    Set objAccess = CreateObject("Access.Application")
    If Err Then 
        On Error GoTo 0
        Err.Raise 5155, "MkDatabase Statement", _
              "MS Access is not installed on this server."
        Exit Sub
    End If
    With objAccess
        .echo False
        .NewCurrentDatabase pathname
        .CloseCurrentDatabase
        .Quit
    End With
    Set objAccess = Nothing
    On Error GoTo 0
End Sub

Public Function NumericToRoman(ByVal Value)
    Dim iPos, sBuffer, iReference
    Dim sLowChar, sMidChar, sHighChar
    sBuffer = String(Value \ 1000, "M")
    Value = Value Mod 1000
    iReference = 100
    Do Until iReference = 0
        If iReference = 100 Then
            sHighChar = "M"
            sMidChar = "D"
            sLowChar = "C"
        ElseIf iReference = 10 Then
            sHighChar = "C"
            sMidChar = "L"
            sLowChar = "X"
        Else
            sHighChar = "X"
            sMidChar = "V"
            sLowChar = "I"
        End If
        iPos = Value \ iReference
        If (iPos > 0) And (iPos < 4) Then
            sBuffer = sBuffer & string(iPos, sLowChar)
        ElseIf iPos = 4 Then
            sBuffer = sBuffer & sLowChar & sMidChar
        ElseIf iPos = 5 Then
            sBuffer = sBuffer & sMidChar
        ElseIf (iPos > 5) And (iPos < 9) Then
            sBuffer = sBuffer & sMidChar & string(iPos - 5, sLowChar)
        ElseIf iPos = 9 Then
            sBuffer = sBuffer & sLowChar & sHighChar
        End If
        Value = Value - iReference * iPos
        iReference = iReference \ 10
    Loop
    NumericToRoman = sBuffer
End Function

Public Function FormatTxt(v)
    v = to_String(v)
    v = Replace(v, "[2]", Chr(178))
    v = Replace(v, "[0]", Chr(186))
    v = Replace(v, "[1/2]", Chr(189))
    v = Replace(v, "[1/4]", Chr(188))
    v = Replace(v, "[1]", Chr(185))
    v = Replace(v, "[3/4]", Chr(190))
    v = Replace(v, "[3]", Chr(179))
    v = Replace(v, "[c]", Chr(169))
    v = ReplaceX(v, "[cents]", Chr(162))
    v = ReplaceX(v, "[deg]", Chr(176))
    v = ReplaceX(v, "[div]", Chr(247))
    v = ReplaceX(v, "[plusminus]", Chr(177))
    v = ReplaceX(v, "[pounds]", Chr(163))
    v = ReplaceX(v, "[r]", Chr(174))
    v = ReplaceX(v, "[tm]", Chr(153))
    FormatTxt = v
End Function

Function GetNameFromDictByCode(sCode, sDictName, sDefault)
    GetNameFromDictByCode = GetFromBy(sCode, "MC", sDictName, "DM", sDefault)
End Function

Function GetNamesFromDictByCodes(sCode, sDictName, sDefault)
    If Not IsArray(sCode) Then
        aArray = Split(sCode, ",")
    End If
    GetNamesFromDictByCodes = sDefault
    If UBound(aArray) > -1 Then
        For i = 0 To UBound(aArray)
            If i > 0 Then GetNamesFromDictByCodes = GetNamesFromDictByCodes & ","
            sGetNameFromDictByCode = GetNameFromDictByCode(aArray(i), sDictName, Null)
            If Not IsNull(sGetNameFromDictByCode) Then
                GetNamesFromDictByCodes = GetNamesFromDictByCodes & sGetNameFromDictByCode
            End If
        Next
    End If
End Function

Function GetFromBy(sValue, sGetField, sTableName, sByField, sDefault) 'get_A_from_B_by_C
    Dim aDict, oDictRs, i, j, sSql
    ReDim aDict(0, -1)
    If Not IsArray(Application(sTableName)) Then
        sSql = "Select " & sByField & ", " & sGetField & " From " & sTableName
        Set oDictRs = Exec(sSql)
        If Not oDictRs.EOF Then aDict = oDictRs.GetRows()
        Application.Lock()
        Application(sTable) = aDict
        Application.UnLock()
    Else
        aDict = Application(sTable)
        If UBound(aDict) <> 1 Then
            Application.Lock()
            Application.Contents.Remove(B)
            Application.UnLock()
            GetFromBy = GetFromBy(sValue, sGetField, sTableName, sByField, sDefault)
            Exit Function
        End If
    End If
    GetFromBy = sDefault
    For i = 0 To UBound(aDict, 2)
        If StrComp(sValue, aDict(0, i), 1) = 0 Then
            GetFromBy = aDict(1, i)
            Exit Function
        End If
    Next
End Function
%>

<%
'********************************************************************************
'    Function（公有）
'    名称 ：    盛飞字符串截取函数
'    作用 ：    按指定首尾字符串截取内容(本函数为从左向右截取)
'    参数 ：    sContent ---- 被截取的内容
'        sStart ------ 首字符串
'        iStartNo ---- 当首字符串不是唯一时取第几个
'        bIncStart --- 是否包含首字符串(1/True为包含，0/False为不包含)
'        iStartCusor - 首偏移值(指针单位为字符数量,左偏用负值,右偏用正值,不偏为0)
'        sOver ------- 尾字符串
'        iOverNo ----- 当尾字符串不是唯一时取第几个
'        bIncOver ---- 是否包含尾字符串((1/True为包含，0/False为不包含)
'        iOverCusor -- 尾偏移值(指针单位为字符数量,左偏用负值,右偏用正值,不偏为0)
'********************************************************************************
Public Function SenFe_Cut(sContent, sStart, iStartNo, bIncStart, iStartCusor, sOver, iOverNo, bIncOver, iOverCusor)
    If sContent<>"" Then
        Dim iStartLen, iOverLen, iStart, iOver, iStartCount, iOverCount, I
        iStartLen = Len(sStart)    '首字符串长度
        iOverLen  = Len(sOver)    '尾字符串长度
        '首字符串第一次出现的位置
        iStart = InStr(sContent, sStart)
        '尾字符串在首字符串的右边第一次出现的位置
        iOver = InStr(iStart + iStartLen, sContent, sOver)
        If iStart>0 And iOver>0 Then
            If iStartNo < 1 Or IsNumeric(iStartNo)=False Then iStartNo = 1
            If iOverNo < 1 Or IsNumeric(iOverNo)=False Then iOverNo  = 1
            '取得首字符串出现的次数
            iStartCount = UBound(Split(sContent, sStart))
            If iStartNo>1 And iStartCount>0 Then
                If iStartNo>iStartCount Then iStartNo = iStartCount
                For I = 1 To iStartNo
                    iStart = InStr(iStart, sContent, sStart) + iStartLen
                Next
                iOver = InStr(iStart, sContent, sOver)
                iStart = iStart - iStartLen    '还原默认状态：包含首字符串
            End If
            '取得尾字符串出现的次数
            iOverCount = UBound(Split(Mid(sContent, iStart + iStartLen), sOver))
            If iOverNo>1 And iOverCount>0 Then
                If iOverNo>iOverCount Then iOverNo = iOverCount
                For I=1 To iOverNo
                    iOver = InStr(iOver, sContent, sOver) + iOverLen
                Next
                iOver = iOver - iOverLen    '还原默认状态：不包含尾字符串
            End If
            If CBool(bIncStart)=False Then iStart = iStart + iStartLen    '不包含首字符串
            If CBool(bIncOver)  Then iOver = iOver + iOverLen        '包含尾字符串
            iStart = iStart + iStartCusor    '加上首偏移值
            iOver  = iOver + iOverCusor    '加上尾偏移值
            If iStart<1 Then iStart = 1
            If iOver<=iStart Then iOver = iStart + 1
            '按指定的开始和结束位置截取内容
            SenFe_Cut = Mid(sContent, iStart, iOver - iStart)
        Else
            'SenFe_Cut = sContent
            SenFe_Cut = "没有找到您想要的内容，可能您设定的首尾字符串不存在！"
        End If
    Else
        SenFe_Cut = "没有内容！"
    End If
End Function
%>
<%
Function Del(sPath)
	Del = DeleteFile(sPath) OR DeleteFolder(sPath)
End Function
Function DeleteFile(sPath)
	On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sPath = MapPathEx(sPath)
    If oFSO.FileExists(sPath) Then
        oFSO.DeleteFile sPath, True
        DeleteFile = True
    End If
	If Err.Number <> 0 Then
		Err.Clear()
        DeleteFile = False
    End If
	Set oFSO = Nothing
End Function
Function DeleteFolder(sPath)
	On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sPath = MapPathEx(sPath)
    If oFSO.FolderExists(sPath) Then
        oFSO.DeleteFolder sPath, True
        DeleteFolder = True
    End If
	If Err.Number <> 0 Then
		Err.Clear()
		DeleteFolder = False
	End If
	Set oFSO = Nothing
End Function
Function MKDIR(sPath)
    On Error Resume Next
	Dim oFSO
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    Dim aPath, i
    sPath = MapPathEx(sPath)
    If oFSO.FolderExists(sPath) Then MKDIR = True : Exit Function
    aPath = Split(sPath, "\")
    sPath = ""
    For i = 0 To UBound(aPath)
        sPath = sPath & aPath(i) & "\"
        If Not oFSO.FolderExists(sPath) Then
            oFSO.CreateFolder(sPath)
        End If
    Next
    If Err.Number > 0 Then
        Err.Clear()
        MKDIR = False
    Else
        MKDIR = True
    End If
    Set oFSO = Nothing
End Function
Function getDestinationFolder(sDestination)
	Dim aDestinationFolder, b_IsFolder
	sDestination = Replace(to_Str(sDestination), "/", "\")
	b_IsFolder = (Right(sDestination, 1) = "\")
	getDestinationFolder = MapPathEx(sDestination) & IIf(b_IsFolder, "\", "")
	aDestinationFolder = Split(getDestinationFolder, "\")
	If InStr(aDestinationFolder(UBound(aDestinationFolder)), ".") > 0 Then b_IsFolder = True
	ReDim Preserve aDestinationFolder(UBound(aDestinationFolder)-1)
	getDestinationFolder = Join(aDestinationFolder, "\")
	getDestinationFolder = MapPathEx(getDestinationFolder) & IIf(b_IsFolder, "\", "")
End Function
Function XCopy(sSource,sDestination)
	XCopy = CopyFile(sSource,sDestination) OR CopyFolder(sSource,sDestination)
End Function
Public Function CopyFile(sSource,sDestination)
	On Error Resume Next
	Dim oFSO, b_IsFolder
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sSource = MapPathEx(sSource)
	sFileName = oFSO.getFileName(sSource)
	If InStr(sFileName, "*") > 0 OR InStr(sFileName, "?") > 0 Then
		sDestinationFolder = getDestinationFolder(sDestination & "\")
		MKDIR(sDestinationFolder)
		oFSO.CopyFile sSource,sDestination, True 
	Else
	    If oFSO.FileExists(sSource) Then
	    	sDestinationFolder = getDestinationFolder(sDestination)
	    	MKDIR(sDestinationFolder)
	        oFSO.CopyFile sSource,sDestination, True 
	    End If
    End If
	If Err.Number <> 0 Then
		Err.Clear()
		CopyFile = False
	Else
	    CopyFile = True
	End If
	Set oFSO = Nothing
End Function
Public Function CopyFolder(sSource,sDestination)
	On Error Resume Next
	Dim oFSO, b_IsFolder
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sSource = MapPathEx(sSource)
	sFolderName = oFSO.getFileName(sSource)
	If InStr(sFolderName, "*") > 0 OR InStr(sFolderName, "?") > 0 Then
		sDestinationFolder = getDestinationFolder(sDestination & "\")
		MKDIR(sDestinationFolder)
		oFSO.CopyFolder sSource,sDestinationFolder, True 
	Else
		sDestinationFolder = getDestinationFolder(sDestination)
		MKDIR(sDestinationFolder)
	    If oFSO.FolderExists(sSource) Then
	        oFSO.CopyFolder sSource,sDestination, True 
	    End If
    End If
	If Err.Number <> 0 Then
		Err.Clear()
		CopyFolder = False
	Else
	    CopyFolder = True
	End If
	Set oFSO = Nothing
End Function
%>