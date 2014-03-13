<%
'����:ȡ�ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function BStr(sValue)
    On Error Resume Next
    BStr = ""
    BStr = CStr(sValue)
    BStr = Trim(BStr)
End Function
%>
<%
'����:ֻȡ����
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function Bint(sValue)
    On Error Resume Next
    Bint = 0
    Bint = Fix(CDbl(sValue))
End Function
%>
<%
'����:ֻȡ����
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function BDbl(sValue)
    On Error Resume Next
    BDbl = 0
    BDbl = CDbl(sValue)
End Function
%>
<%
'����:���� SQL �� "����" ��ѯǰ���˵�����
'��Դ:http://jorkin.reallydo.com/article.asp?id=389

Function Str4Sql(sText)
    Str4Sql = Replace(BStr(sText), "'", "''")
End Function
%>
<%
'����:���� SQL ģ������ Like ֮ǰ��һ��
'��Դ:http://jorkin.reallydo.com/article.asp?id=389

Function Str4Like(byVal sText)
    sText = Replace(sText, "'" , "''")
    sText = Replace(sText, "[" , "[[]")
    sText = Replace(sText, "%" , "[%]")
    sText = Replace(sText, "_" , "[_]")
    Str4Like = sText
End Function
%>
<%
'����:����ַ����ӻ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=467

Function WriteLn(sString)
    Response.Write( sString & vbCrLf )
End Function
%>
<%
'����:����ַ�����HTML����<br />
'��Դ:http://jorkin.reallydo.com/article.asp?id=467

Function PrintLn(sString)
    Response.Write( sString & vbCrLf & "<br />" & vbCrLf)
End Function
%>
<%
'����:�жϸüƻ������Ƿ���,һ�㶨ʱ����Application��ʱ��ʹ��.
'sInterval, iNumber, dStartTime����ͬ DateAdd �����Ĳ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=423
'������ʱ�䣺2009/01/14

Function ScheduleTask(sTaskName, sInterval, iNumber, dStartTime)
    Dim sApplicationName, sLastUpdate, sNextUpdate
    Select Case UCase(sInterval)
        Case "YYYY", "Q", "M", "Y", "D", "W", "WW", "H", "N", "S"
            sApplicationName = "ScheduleTask_" & sTaskName & "_LastUpdate"
            sLastUpdate = Trim(Application(sApplicationName))
            dStartTime = BDate(dStartTime)
            ScheduleTask = False
            If sLastUpdate = "" Then
                sLastUpdate = DateAdd(sInterval, Fix(DateDiff(sInterval, dStartTime, Now()) / iNumber -1) * iNumber, dStartTime)
                Application(sApplicationName) = sLastUpdate
            End If
            sNextUpdate = DateAdd(sInterval, Fix(DateDiff(sInterval, dStartTime, Now()) / iNumber) * iNumber, dStartTime)
            If Now() > sNextUpdate Then
                ScheduleTask = True
                Application(sApplicationName) = sNextUpdate
            End If
        Case Else
            ScheduleTask = False
    End Select
End Function
%>
<%
'����:����Ŀ¼���·����
'UNPATH [Drive:]Path
'���Լ���༶��Ŀ¼������֧�����·���;���·����
'��Դ��http://jorkin.reallydo.com/article.asp?id=419
'��ҪPath������http://jorkin.reallydo.com/article.asp?id=401
'��ҪReplaceAll������http://jorkin.reallydo.com/article.asp?id=406

Function UnPath(byVal sPath)
    On Error Resume Next
    Dim cPath : cPath = ""
    While InStr (1, Path(sPath), Path(cPath), 1) = 0
        cPath = "../" & cPath
    Wend
    UnPath = Replace( Path(sPath) , Path(cPath), "", 1, -1, 1 )
    UnPath = Replace( cPath & UnPath, "\", "/" )
    UnPath = ReplaceAll(UnPath, "//", "/", False)
    UnPath = ReplaceAll(UnPath, "...", "../..", False)
End Function
%>
<%
'����:�ظ� N�� �ض��ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=413
'��ҪBint����:http://jorkin.reallydo.com/article.asp?id=395

Function Repeat(nTimes, sStr)
    nTimes = Bint(nTimes)
    sStr = BStr(sStr)
    Repeat = Replace(Space(nTimes), Space(1), sStr)
End Function
%>
<%
'AInt��ʽ�����������ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function AInt(s)
    Dim a, i
    a = Split(Bstr(s), ",")
    If UBound(a) < 0 Then
        AInt = "0"
        Exit Function
    End If
    For i = 0 To UBound(a)
        a(i) = Bint(a(i))
    Next
    AInt = Join(a, ",")
End Function
%>
<%
'����:������תΪ���ڸ�ʽ
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function BDate(sDate)
    If IsDate(sDate) Then
        BDate = CDate(sDate)
    Else
        BDate = Date()
    End If
End Function
%>
<%
'���ܣ�ʹ�� &# �� HTML �е������ַ����� Unicode ����
'��Դ��http://jorkin.reallydo.com/article.asp?id=408

Function ASCII(sStr)
    If IsBlank(sStr) Then ASCII="" : Exit Function
    Dim i
    ASCII = ""
    For i = 1 To Len(sStr)
        ASCII = ASCII & "&#x" & Hex(AscW(Mid(sStr, i, 1))) & ";"
    Next
End Function
%>
<%
'���ܣ������ַ���������ָ����Ŀ��ĳ���ַ��� ȫ�� ���滻Ϊ��һ�����ַ�����
'��Դ��http://jorkin.reallydo.com/article.asp?id=406
'��ҪBint����:http://jorkin.reallydo.com/article.asp?id=395

Function ReplaceAll(sExpression, sFind, sReplaceWith, bAll)
    If IsBlank(sExpression) Then ReplaceAll = "" : Exit Function
    If (StrComp(bAll, "True", 1) = 0) Or (CBool(Bint(bAll)) = True) Then
        Do While InStr( 1, sExpression, sFind, 1) > 0
            sExpression = Replace(sExpression, sFind, sReplaceWith, 1, -1, 1)
            If InStr( 1, sReplaceWith , sFind , 1) >0 Then Exit Do
        Loop
    Else
        Do While InStr(sExpression, sFind) > 0
            sExpression = Replace(sExpression, sFind, sReplaceWith)
            If InStr(sReplaceWith, sFind ) > 0 Then Exit Do
        Loop
    End If
    ReplaceAll = sExpression
End Function
%>
<%
'����:����Ŀ¼����·����
'PATH [Drive:]Path
'֧�ֶ༶Ŀ¼��֧�����·���;���·����
'֧���á�...��ָ����Ŀ¼�ĸ�Ŀ¼��
'��Դ��http://jorkin.reallydo.com/article.asp?id=401
'��ҪReplaceAll������http://jorkin.reallydo.com/article.asp?id=406

Function Path(ByVal sPath)
    On Error Resume Next
    If BStr(sPath) = "" Then sPath = "./"
    If Right(sPath, 1) = ":" Then sPath = sPath & "\"
    sPath = Replace(sPath, "/", "\")
    sPath = ReplaceAll(sPath, "\\", "\", False)
    sPath = ReplaceAll(sPath, "...", "..\..", False)
    If (InStr(sPath, ":") > 0) Then
        sPath = sPath
    Else
        sPath = Server.Mappath(sPath)
    End If
    Path = sPath
End Function
%>
<%
'����:ʹ�������ʾʽ���ַ��������滻
'��Դ:http://jorkin.reallydo.com/article.asp?id=345

Function RegReplace(Str, PatternStr, RepStr)
    Dim NewStr, regEx
    NewStr = Str
    If IsBlank(NewStr) Then
        RegReplace = ""
        Exit Function
    End If
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = PatternStr
    NewStr = regEx.Replace(NewStr, RepStr)
    RegReplace = NewStr
End Function
%>
<%
'����:���ַ����еİ���ַ�תΪȫ��
'��Դ:http://jorkin.reallydo.com/article.asp?id=339

Function SBC2DBC(sStr)
    'ֻ��Chr(32)��Chr(126)תΪȫ�ǲ�������
    '���Ը�����Ҫ��Ϊ: For i = 1 To 256
    For i = 32 To 126
        If InStr(sStr, Chr(i))>0 Then
            sStr = Replace(sStr, Chr(i), Chr(i -23680))
        End If
    Next
    SBC2DBC = sStr
End Function
%>
<%
'����:���ַ����е�ȫ���ַ�תΪ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=339

Function DBC2SBC(sStr)
    'ֻ��Chr(-23648)��Chr(-23554)תΪ��ǲ�������
    '���Ը�����Ҫ��Ϊ: For i = -23679 To -23424
    For i = -23648 To -23554
        If InStr(sStr, Chr(i))>0 Then
            sStr = Replace(sStr, Chr(i), Chr(i + 23680))
        End If
    Next
    DBC2SBC = sStr
End Function
%>
<%
'����:��TEXTAREA������תΪHTML���
'��Դ:http://jorkin.reallydo.com/article.asp?id=190

Function HTMLEncode2(s)
    s = Server.HTMLEncode(BStr(s))
    s = Replace(s, vbTab, "    ")
    s = Replace(s, vbNewLine, "<br />")
    s = Replace(s, vbCr, "<br />")
    s = Replace(s, "  ", "&nbsp; ")
    HTMLEncode2 = s
End Function
%>
<%
'����:����Ƿ����ϵͳ���������Ƿ�װ�ɹ�
'��Դ:http://jorkin.reallydo.com/article.asp?id=163

Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function
%>
<%
'����:����Ƿ�ΪForm��Post
'��Դ:http://jorkin.reallydo.com/article.asp?id=47

Function isPostBack()
    isPostBack = False
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then isPostBack = True
End Function
%>
<%
'����:�жϷ����Ƿ������ⲿJorkin�޸İ�
'��Դ:http://jorkin.reallydo.com/article.asp?id=31

Public Function ChkPost()
    Dim server_v1, server_v2
    Chkpost = False
    server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
    server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
    If Mid(server_v1, InStr(server_v1, "://") + 3, Len(server_v2)) = server_v2 Then Chkpost = True
End Function
%>
<%
'����:��ȡ��Ӣ�Ļ���ַ���ǰstrlen���ַ�.
'�÷�:CutStr(str,strlen)
'��Դ:http://jorkin.reallydo.com/article.asp?id=28

Function CutStr(Str, strlen)
    Dim l, t, c, i
    l = Len(Str & "")
    t = 0
    For i = 1 To l
        c = Abs(Asc(Mid(Str, i, 1)))
        If c>255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t>= strlen Then
            CutStr = Left(Str, i)&"..."
            Exit For
        Else
            CutStr = Str
        End If
    Next
    CutStr = Replace(CutStr, Chr(10), "")
End Function
%>
<%
'����:ASP���IIF
'�÷�:IIF(�������ʽ,Ϊ��ʱ����ֵ,Ϊ��ʱ����ֵ)
'��Դ:http://jorkin.reallydo.com/article.asp?id=26

Function IIf(bExp1, sVal1, sVal2)
    If (bExp1) Then
        IIf = sVal1
    Else
        IIf = sVal2
    End If
End Function
%>
<%
'����:ȥ��ȫ��HTML���(Jorkin��ǿ��)
'��Դ:http://jorkin.reallydo.com/article.asp?id=32

Public Function ReplaceHTML(Textstr)
    Dim sStr, regEx
    sStr = Textstr
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Multiline = True
    regEx.Pattern = "<!--[\s\S]*?-->" '//���þͰ�ע��ȥ��
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
%>
<%
'����:��һ���ַ���ǰ�油ȫ��һ�ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=452

Public Function LFill(ByVal sString, ByVal sStr)
    Dim i, iStrLen : iStrLen = Len(BStr(sStr))
    For i = iStrLen To 1 Step -1
        If Right(sStr, i ) = Left(sString, i ) Then Exit For
    Next
    LFill = Left(sStr, iStrLen - i) & sString
End Function
%>
<%
'����:��һ���ַ������油ȫ��һ�ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=452

Public Function RFill(ByVal sString, ByVal sStr)
    Dim i, iStrLen : iStrLen = Len(BStr(sStr))
    For i = iStrLen To 1 Step -1
        If Left(sStr, i) = Right(sString, i) Then Exit For
    Next
    RFill = sString & Mid(sStr, i + 1)
End Function
%>
<%
'����:�ж��Ƿ��ǿ�ֵ
'��Դ:http://jorkin.reallydo.com/article.asp?id=386

Public Function IsBlank(byref TempVar)
    IsBlank = False
    Select Case VarType(TempVar)
        Case 0, 1 '--- Empty & Null
            IsBlank = True
        Case 8 '--- String
            If Len(TempVar) = 0 Then
                IsBlank = True
            End If
        Case 9 '--- Object
            tmpType = TypeName(TempVar)
            If (tmpType = "Nothing") Or (tmpType = "Empty") Then
                IsBlank = True
            End If
        Case 8192, 8204, 8209 '--- Array
            If UBound(TempVar) = -1 Then
                IsBlank = True
            End If
    End Select
End Function
%>
<%
'����:���alert��Ϣ��ʵ��ҳ����ת
'��Դ:http://jorkin.reallydo.com/article.asp?id=470
'��ҪStr4Js����:http://jorkin.reallydo.com/article.asp?id=466

Function doAlert(sInfo, sUrl)
    On Error Resume Next
    sUrl = BStr(sUrl)
    sInfo = BStr(sInfo)
    Select Case LCase(sUrl)
        Case "back"
            sUrl = "javascript:history.back()"
        Case "referer"
            sUrl = Request.ServerVariables("HTTP_REFERER")
        Case "close"
            sUrl = "javascript:window.close();"
    End Select
    WriteLn("</scr" & "ipt>" & vbCrLf & "<scr" & "ipt language=""javascript"">" )
    If sInfo<>"" Then WriteLn("  alert('" & Str4Js(sInfo) & "');" )
    If sUrl<>"" Then
        Closeconn()
        WriteLn("  window.location.href='" & Str4Js(sUrl) & "';" & vbCrLf & "</scr" & "ipt>" )
        Response.End()
    Else
        WriteLn("</scr" & "ipt>" )
        Response.Flush()
    End If
End Function
%>
<%
'����:����ת��ΪJavascript�ַ������(Լ��Javascript�У��õ����������ַ���)
'��Դ:http://jorkin.reallydo.com/article.asp?id=465

Function Str4Js(sString)
    Str4Js = sString
    If IsBlank(sString) Then Str4Js = "" : Exit Function
    Str4Js = Replace(Str4Js, "\", "\\")
    Str4Js = Replace(Str4Js, "'", "\'")
    Str4Js = Replace(Str4Js, vbCrLf, "\n")
    Str4Js = Replace(Str4Js, vbCr, "\n")
    Str4Js = Replace(Str4Js, vbLf, "\n")
    Str4Js = Replace(Str4Js, vbTab, "\t")
    Str4Js = Replace(Str4Js, "script", "scr'+'ipt", 1, -1 , 1)
End Function
%>
<%
'����:�ж�һ��ֵ�Ƿ����������
'��Դ:http://jorkin.reallydo.com/article.asp?id=462

Function InArray( sValue, aArray )
    Dim x
    InArray = False
    For Each x In aArray
        If x = sValue Then
            InArray = True
            Exit For
        End If
    Next
End Function
%>
<%
'����:���ɶ�ؼ��ֲ�ѯSQL���(������)
'��������Google�ÿո�ָ�ؼ���,�ݲ�֧�� And or | ��.
'��Դ:http://jorkin.reallydo.com/article.asp?id=416
'��ҪReplaceAll����:http://jorkin.reallydo.com/article.asp?id=406
'��ҪStr4Like����:http://jorkin.reallydo.com/article.asp?id=389

Function Key4Search(sString, sFields)
    On Error Resume Next
    sFields = BStr(sFields)
    sString = Trim(ReplaceAll(sString, "  ", " ", True))
    aString = Split(sString, " ", -1, 1)
    iLenString = UBound(aString)
    Key4Search = Key4Search & " ( 1=1 "
    For i = 0 To iLenString
        Key4Search = Key4Search & " And " & sFields & " Like '%" & Str4Like(aString(i)) & "%' "
    Next
    Key4Search = Key4Search & " ) "
End Function
%>
<%
Function Del(sFileName)
    On Error Resume Next
    sFileName = Path(sFileName)
    Dim oFSO
    Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    If oFSO.FileExists(sFileName) Then
        oFSO.DeleteFile sFileName, True
        Del = True
    End If
    Set oFSO = Nothing
    If Err.Number <> 0 Then
        Del = False
    End If
End Function
%>
<%
'Rem Check For valid syntax in an email address.

Function IsValidEmail(email)
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
%>
<%
'����:�๦�����ڸ�ʽ������
'��Դ:http://jorkin.reallydo.com/article.asp?id=477

Function FormatDate(sDateTime, sReallyDo)
    Dim sTmpString, sJorkin, i
    sTmpString = ""
    FormatDate = ""
    sReallyDo = BStr(sReallyDo)
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
        Case "RND", "RAND", "RANDOMIZE" '//����ַ���
            Randomize
            sJorkin = Rnd()
            FormatDate = FormatDate(sDateTime, "YYYYMMDDHHNNSS") & _
                         Fix((9 * 10^7 -1) * sJorkin) + 10^7
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
%>
<%
Function FormatDateTimeString(sDateTime, sReallyDo)
    Dim sLocale
    sLocale = GetLocale()
    SetLocale("en-gb")
    Select Case sReallyDo
        Case "YYYY" '// 4λ����
            FormatDateTimeString = Year(sDateTime)
        Case "YY" '// 2λ����
            FormatDateTimeString = Right(Year(sDateTime), 2)
        Case "CMMMM", "CMMM", "CMM", "CM" '// �����·���
            SetLocale("zh-cn")
            FormatDateTimeString = MonthName(Month(sDateTime))
        Case "MMMM" '// Ӣ���·���(ȫ)
            FormatDateTimeString = MonthName(Month(sDateTime), False)
        Case "MMM" '//Ӣ���·���(��)
            FormatDateTimeString = MonthName(Month(sDateTime), True)
        Case "MM" '// ��(��0)
            FormatDateTimeString = Right("0" & Month(sDateTime), 2)
        Case "M" '// ��
            FormatDateTimeString = Month(sDateTime)
        Case "JD" '// ��(��st nd rd th)
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
        Case "DD" '// ��(��0)
            FormatDateTimeString = Right("0" & Day(sDateTime), 2)
        Case "D" '// ��
            FormatDateTimeString = Day(sDateTime)
        Case "HH" '// Сʱ(��0)
            FormatDateTimeString = Right("0" & Hour(sDateTime), 2)
        Case "H" '// Сʱ
            FormatDateTimeString = Hour(sDateTime)
        Case "NN" '// ��(��0)
            FormatDateTimeString = Right("0" & Minute(sDateTime), 2)
        Case "N" '// ��
            FormatDateTimeString = Minute(sDateTime)
        Case "SS" '//��(��0)
            FormatDateTimeString = Right("0" & Second(sDateTime), 2)
        Case "S" '//��
            FormatDateTimeString = Second(sDateTime)
        Case "CWW", "CW" '// ��������
            SetLocale("zh-cn")
            FormatDateTimeString = WeekdayName(Weekday(sDateTime))
        Case "WW" '// Ӣ������(ȫ)
            FormatDateTimeString = WeekdayName(Weekday(sDateTime), False)
        Case "W" '// Ӣ������(��)
            FormatDateTimeString = WeekdayName(Weekday(sDateTime), True)
        Case "CT" '// 12Сʱ��(����/����)
            SetLocale("zh-tw")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "UT" '// 12Сʱ��(AM/PM)
            SetLocale("en-us")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "ET" '// 12Сʱ��(a.m./p.m.)
            SetLocale("es-ar")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "ZT" '// 12Сʱ��(AM/PM)(��0)
            SetLocale("en-za")
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case "T" '// 24Сʱ��ʱ��
            FormatDateTimeString = FormatDateTime(sDateTime, 3)
        Case Else
            FormatDateTimeString = sReallyDo
    End Select
    SetLocale(sLocale)
End Function
%>
<%
'����:���ַ�����ÿ�����ʵ�����ĸ����Ϊ��д
'��Դ:http://jorkin.reallydo.com/article.asp?id=481

Private Function PCase(byVal sString)
    Dim Tmp, Word, Tmp1, Tmp2, firstCt, a, sSentence, c, i
    sString = BStr( sString )
    a = Split( sString, vbCrLf )
    c = UBound(a)
    i = 0
    For Each sSentence In a
        Tmp = Trim( sSentence )
        Tmp = Split( sSentence, " " )
        For Each Word In Tmp
            Word = Trim( Word )
            Tmp1 = UCase( Left( Word, 1 ) )
            Tmp2 = LCase( Mid( Word, 2 ) )
            PCase = PCase & Tmp1 & Tmp2 & " "
        Next
        PCase = Left( PCase, Len(PCase) - 1 )
        If i < c Then
            PCase = PCase & vbCrLf
        End If
        i = i + 1
    Next
End Function
%>
<%
'����:�ж�һ�������Ƿ�����һ����ά���ݵ�ĳ��
'��Դ:http://jorkin.reallydo.com/article.asp?id=481
'��ҪBint����:http://jorkin.reallydo.com/article.asp?id=395

Function InArray2(ByVal sValue, ByVal aArray(), ByVal iColumn, ByVal bCompare)
    On Error Resume Next
    Dim i, j
    InArray2 = False
    i = Bint(iColumn)
    If i < 0 Or i > UBound(A) Then
        Exit Function
    End If
    For j = 0 To UBound(A, 2)
        If StrComp(sValue, A(i, j), bCompare) = 0 Then
            WriteLn(A(i, j) & "<br />")
            InArray2 = True
            Exit Function
        End If
    Next
End Function
%>
<%
'����:��һ��һά�����������
'��Դ:http://jorkin.reallydo.com/article.asp?id=394

Private Function SortArray(byVal UnSortedArray)
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
%>
<%
'����:ɾ����ά�����е�iColumn����ֵΪsValue����,��iColumnΪ-1ʱɾ����sValue��
'��Դ:http://jorkin.reallydo.com/article.asp?id=483

Function DelArray2(aArray, iColumn, sValue, bCompare)
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
%>
<%
'����:ִ��SQL���
'��Դ��http://jorkin.reallydo.com/article.asp?id=487

Public Function Exec(sCommand)
    On Error Resume Next
    Server.ScriptTimeOut = 29252888
    OpenConn()
    Set Exec = oConn.Execute(sCommand)
    If Err Then
        WriteLn Err.Source & "���������Ĳ�ѯ�����Ƿ���ȷ��<br />"
        WriteLn "Error : # " & Err.Number & " <br />"
        WriteLn "Description : " & Err.Description & "<br />"
        WriteLn "Command : " & Server.HTMLEncode(sCommand) & "<br />"
        Err.Clear
        Response.End
    End If
End Function
%>
<%
'����:������ָ��CheckBox���ļ���ֵ�ϴ�
'��Դ:http://jorkin.reallydo.com/article.asp?id=465

Function CheckBoxScript(ByVal FormElement , ByVal ElementValue)
    CheckBoxScript = "<scr" & "ipt language=""javascript"" type=""text/javascript"">" & vbCrLf & "String.prototype."
    CheckBoxScript = CheckBoxScript & "ReallyDo=function(){return this.replace(/(^\s*)|(\s*$)/g,"""");}" & vbCrLf
    CheckBoxScript = CheckBoxScript & "var Jorkin = """ & ElementValue & """.split("","");" & vbCrLf
    CheckBoxScript = CheckBoxScript & "for (i = 0; i < " & FormElement & ".length; i++){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "  for (j = 0; j < Jorkin.length; j++){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "    if (Jorkin[j].ReallyDo() == " & FormElement & "[i].value.ReallyDo()){" & vbCrLf
    CheckBoxScript = CheckBoxScript & "      " & FormElement & "[i].checked = true } } }</scr" & "ipt>" & vbCrLf
End Function
%>
<%
'����:���� Boolean ֵָ�����ʽ��ֵ�Ƿ�Ϊ��ĸ��
'��Դ:http://jorkin.reallydo.com/article.asp?id=525

Private Function IsAlpha(byVal sString)
    Dim regExp, oMatch, i, sStr
    For i = 1 To Len(BStr(sString))
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
%>
<%
'����HTML���ֱ�ǩ��ʽ�ű�
'��Դ:http://jorkin.reallydo.com/article.asp?id=521
'��ҪRegReplace����: http://jorkin.reallydo.com/article.asp?id=345

Function HTMLFilter(sHTML, sFilters)
    If BStr(sHTML) = "" Then Exit Function
    If BStr(sFilters) = "" Then sFilters = "JORKIN,SCRIPT,OBJECT"
    Dim aFilters : aFilters = Split(UCase(sFilters), ",")
    For i = 0 To UBound(aFilters)
        Select Case UCase(Trim(aFilters(i)))
            Case "JORKIN"
                Do While InStr(sHTML, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;") >0
                    sHTML = Replace(sHTML, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "&nbsp;&nbsp;&nbsp;&nbsp;")
                Loop
            Case "SCRIPT"
                '// ȥ���ű�<scr ipt></scr ipt>�� onload ��
                sHTML = RegReplace(sHTML, "<SCRIPT[\s\S]*?</SCRIPT>", "")
                sHTML = RegReplace(sHTML, "(JAVASCRIPT|JSCRIPT|VBSCRIPT|VBS):", "$1��")
                sHTML = RegReplace(sHTML, "\s[on].+?=\s+?([\""|\'])(.*?)\1", "")
            Case "FIXIMG"
                sHTML = RegReplace(sHTML, "<IMG.*?\sSRC=([^\""\'\s][^\""\'\s>]*).*?>", "<img src=$2 border=0>")
                sHTML = RegReplace(sHTML, "<IMG.*SRC=([\""\']?)(.\1\S+).*?>", "<img src=$2 border=0>")
            Case "TABLE"
                '// ȥ�����<table><tr><td><th>
                sHTML = RegReplace(sHTML, "</?TABLE[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?TBODY[^>]*>", "")
                sHTML = RegReplace(sHTML, "<(/?)TR[^>]*>", "<$1p>")
                sHTML = RegReplace(sHTML, "</?TH[^>]*>", " ")
                sHTML = RegReplace(sHTML, "</?TD[^>]*>", " ")
            Case "CLASS"
                '// ȥ����ʽ��class=""
                sHTML = RegReplace(sHTML, "(<[^>]+) CLASS=[^ |^>]+([^>]*>)", "$1 $2")
                sHTML = RegReplace(sHTML, "\sCLASS\s*?=\s*?([\""|\'])(.*?)\1", "")
            Case "STYLE"
                '// ȥ����ʽstyle=""
                sHTML = RegReplace(sHTML, "(<[^>]+) STYLE=[^ |^>]+([^>]*>)", "$1 $2")
                sHTML = RegReplace(sHTML, "\sSTYLE\s*?=\s*?([\""|\'])(.*?)\1", "")
            Case "XML"
                '// ȥ��XML<?xml>
                sHTML = RegReplace(sHTML, "<\\?XML[^>]*>", "")
            Case "NAMESPACE"
                '// ȥ�������ռ�<o:p></o:p>
                sHTML = RegReplace(sHTML, "<\/?[a-z]+:[^>]*>", "")
            Case "FONT"
                '// ȥ������<font></font>
                sHTML = RegReplace(sHTML, "</?FONT[^>]*>", "")
            Case "MARQUEE"
                '// ȥ����Ļ<marquee></marquee>
                sHTML = RegReplace(sHTML, "</?MARQUEE[^>]*>", "")
            Case "OBJECT"
                '// ȥ������<object><param><embed></object>
                sHTML = RegReplace(sHTML, "</?OBJECT[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?PARAM[^>]*>", "")
                sHTML = RegReplace(sHTML, "</?EMBED[^>]*>", "")
            Case "COMMENT"
                '// ȥ��HTMLע��, �ᴦ��<script>��<style>��ע��, ����
                sHTML = RegReplace(sHTML, "<!--[\s\S]*?-->", "")
            Case Else
                '// ȥ��������ǩ
                sHTML = RegReplace(sHTML, "</?" & aFilters(i) & "[^>]*?>", "")
        End Select
    Next
    HTMLFilter = sHTML
End Function
%>
<%
'����:�����������ֵ,֧�ּ�¼��/�ַ���/һά����/��ά����/���ַ���������
'��Դ:http://jorkin.reallydo.com/article.asp?id=217
'��ҪBStr����:http://jorkin.reallydo.com/article.asp?id=395
'��ҪIIf����:http://jorkin.reallydo.com/article.asp?id=26
'��ҪWriteLn����:http://jorkin.reallydo.com/article.asp?id=467

Public TraceStyle

Function Trace(ByVal s)
    On Error Resume Next
    If Not TraceStyle Then
        WriteLn("<style>.tracediv{background:#CCE8CF;color:#000;font:14px;margin:5px;text-align:left;}.tracediv table,.tracediv td,.tracediv hr,.tracediv fieldset{font:12px;border-collapse:collapse;border:1px solid #820222;padding:3px;margin:3px;color:000}</style>")
        TraceStyle = True
    End If
    WriteLn("<div class=""tracediv""><fieldset>")
    Dim iUBound1, iUBound2
    Dim i, j
    If TypeName(s) = "Recordset" Then
			WriteLn("<legend style=""color:red;"">Recordset :</legend>")
			If oRs.State = 1 Then
				Set s = s.Clone
				If Not s.BOF Then s.MoveFirst
				WriteLn("<table><tr><td><nobr>���<nobr></td>")
				For i = 0 To s.Fields.Count - 1
					WriteLn("<td>" & s(i).Name & "</td>")
				Next
				Do While Not s.EOF
					j = j + 1
					If j > 50 Then Exit Do
					WriteLn("<tr><td>" & j & "</td>")
					For i = 0 To s.Fields.Count - 1
						If IsNull(s(i)) Then
							WriteLn("<td><font color=""red"">&lt;NULL&gt;</font></td>")
						Else
							WriteLn("<td>" & HTMLEncode2(s(i)) & "</td>")
						End If
					Next
					WriteLn("</tr>")
					s.MoveNext
				Loop
			Else
				WriteLn("<tr><td><font color=""red"">ָʾ�����ѹرա�</font></td></tr>")
			End If
			WriteLn("</table>")

    ElseIf IsArray(s) Then
        iUBound1 = UBound(s)
        iUBound2 = UBound(s, 2)
        If Err Then
            WriteLn("<legend style=""color:red;"">Array1 :</legend><table>")
            WriteLn("<tr><td>&#21015;</td><td>&#20540;</td></tr>")
            For i = 0 To iUBound1
                WriteLn("<tr><td>" & i & "</td>")
				If IsArray(s(i)) Then
					WriteLn("<td>")
					Trace(s(i))
					WriteLn("</td>")
				Else
	                WriteLn("<td>" & HTMLEncode2(s(i)) & "</td></tr>")
				End If
            Next
            WriteLn("</table>")
        Else
            WriteLn("<legend style=""color:red;"">Array2 :</legend><table>")
            WriteLn("<tr><td>&#20108;&#32500;/&#19968;&#32500;</td>")
            For j = 0 To iUBound1
                WriteLn("<td>" & j & "</td>")
            Next
            WriteLn("</tr>")
            For i = 0 To iUBound2
                WriteLn("<tr><td>" & i & "</td>")
                For j = 0 To iUBound1
					If IsArray(s(j, i)) Then
					WriteLn("<td>")
					Trace(s(j, i))
					WriteLn("</td>")
					Else
	                    WriteLn("<td>" & HTMLEncode2(s(j, i)) & "</td>")
					End If
                Next
                WriteLn("</tr>")
            Next
            WriteLn("</table>")
        End If
    ElseIf IsObject(s) Then
		WriteLn("<legend style=""color:red;"">" & TypeName(s) & " " & s.Version & " :</legend>")
		WriteLn(HTMLEncode2(s))
    Else
        If TypeName(s) = "String" Then s = BStr(s)
        Select Case UCase(s)
            Case "APPLICATION"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Application.Contents.Count & " &#20010;Application&#21464;&#37327;</legend>")
                For Each i in Application.Contents
                    WriteLn("<strong>Application(""" & i & """)" & " = </strong>")
                    Trace(Application(i))
                Next
            Case "COOKIES", "REQUEST.COOKIES"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.Cookies.Count & " &#20010;Request.Cookies&#21464;&#37327;</legend>")
                For Each i in Request.Cookies
                    WriteLn("<strong>Request.Cookies(""" & i & """)" & " = </strong>")
                    If Request.Cookies(i).HasKeys Then
                        WriteLn("<fieldset><legend style=""color:red;"">" & TypeName(Request.Cookies(i)) & " :</legend>")
                        WriteLn("<strong>&#20849; " & Request.Cookies(i).Count & " &#20010;Request.Cookies(""" & i & """)��&#21464;&#37327;</strong><br />")
                        For Each j in Request.Cookies(i)
                            WriteLn("Request.Cookies(""" & i & """)(""" & j & """) = ")
                            Trace(Request.Cookies(i)(j))
                        Next
                        WriteLn
                        WriteLn("</fieldset>")
                    Else
                        Trace(Request.Cookies(i))
                    End If
                Next
            Case "SESSION"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Session.Contents.Count & " &#20010;Session&#21464;&#37327;</legend>")
                For Each i in Session.Contents
                    WriteLn("<strong>Session(""" & i & """)" & " = </strong>")
                    Trace(Session(i))
                Next
            Case "QUERYSTRING", "REQUEST.QUERYSTRING"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.QueryString.Count & " &#20010;Request.QueryString&#21464;&#37327;</legend>")
                For Each i in Request.QueryString
                    WriteLn("<strong>Request.QueryString(""" & i & """)" & " = </strong>")
                    For Each j In Request.QueryString(i)
                        Trace(j)
                    Next
                Next
            Case "FORM", "REQUEST.FORM"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.Form.Count & " &#20010;Request.Form&#21464;&#37327;</legend>")
                For Each i in Request.Form
                    WriteLn("<strong>Request.Form(""" & i & """)" & " = </strong>")
                    For Each j In Request.Form(i)
                        Trace(j)
                    Next
                Next
            Case "SERVERVARIABLES", "REQUEST.SERVERVARIABLES"
                WriteLn("<legend class=""tracediv"" style=""font:bold;"">&#20849; " & Request.ServerVariables.Count & " &#20010;Request.ServerVariables&#21464;&#37327;</legend>")
                For Each i in Request.ServerVariables
                    WriteLn("<strong>Request.ServerVariables(""" & i & """)" & " = </strong>")
                    Trace(Request.ServerVariables(i))
                Next
            Case "REQUEST"
                WriteLn("<legend class=""tracediv"" style=""font:bold;color:#820222;"">ȫ��Request&#21464;&#37327;</legend>")
                Trace("COOKIES")
                Trace("QUERYSTRING")
                Trace("FORM")
                Trace("SERVERVARIABLES")
            Case Else
                WriteLn("<legend style=""color:red;"">" & TypeName(s) & " :</legend>")
                If s = "" Then
                    WriteLn("<font color=""blue"">IsBlank</font>")
                Else
                    WriteLn(HTMLEncode2(s))
                End If
        End Select
    End If
    WriteLn("</fieldset></div>")
    Response.Flush()
End Function
%>
<%
'����:���ǰ�ADO��GetRows()��װ��һ��
'��Դ:http://jorkin.reallydo.com/article.asp?id=537

Function GetRsRows(ByVal oRs)
    On Error Resume Next
    Dim aArray
    ReDim aArray(0, -1)
    If TypeName(oRs) = "Recordset" Then
        If Not oRs.BOF Then oRs.MoveFirst
        If Not oRs.EOF Then aArray = oRs.GetRows()
    End If
    GetRsRows = aArray
End Function
%>
<%
'����:����ȥ��ǰsStr����sStr ��ǰ��sStr ���ַ���sString������(���ִ�Сд)
'��Դ:http://jorkin.reallydo.com/article.asp?id=443

Public Function TrimLR(ByVal sString, ByVal sStr, ByVal sLeftOrRight)
    Dim iStrLen : iStrLen = Len(sStr)
    Select Case UCase(sLeftOrRight)
        Case "L", "LEFT"
            Do While Left(sString, iStrLen) = sStr
                sString = Mid(sString, iStrLen + 1)
            Loop
        Case "R", "RIGHT"
            Do While Right(sString, iStrLen) = sStr
                sString = Mid(sString, 1, Len(sString) - iStrLen)
            Loop
        Case Else
            sString = TrimLR(sString, sStr, "L")
            sString = TrimLR(sString, sStr, "R")
    End Select
    TrimLR = sString
End Function
%>
<%
'����:���԰�ǰ׺����Application�Ķ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=543

Public WriteRemoveApplication

Function RemoveApplication(sString)
    On Error Resume Next
    sString = BStr(sString)
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
    For j = 0 To i
        Application.Lock
        Application.Contents.Remove(aApplicationArray(j))
        If WriteRemoveApplication Then WriteLn "<br />�ͷ� <strong>" & aApplicationArray(j) & "</strong> ���<br />"
        Application.unLock
    Next
    If WriteRemoveApplication Then WriteLn "<br />���ж����Ѿ�����,���ͷ��� <strong>" & j & "</strong> ���������.<br />"
End Function
%>
<%
'����:��ȡȫ��ͼƬ��ַ,���浽һ������.
'��Դ:http://jorkin.reallydo.com/article.asp?id=448
'��ҪReplaceAll����:http://jorkin.reallydo.com/article.asp?id=406

Function getIMG(sString)
    Dim sReallyDo, regEx, iReallyDo
    Dim oMatches, cMatch
    '//����һ��������
    iReallyDo = -1
    ReDim aReallyDo(iReallyDo)
    If IsNull(sString) Then
        getIMG = aReallyDo
        Exit Function
    End If
    '//��ʽ��HTML����
    '//��ÿ�� <img ���� ���������滻
    sReallyDo = sString
    On Error Resume Next
    sReallyDo = Replace(sReallyDo, vbCr, " ")
    sReallyDo = Replace(sReallyDo, vbLf, " ")
    sReallyDo = Replace(sReallyDo, vbTab, " ")
    sReallyDo = Replace(sReallyDo, "<img ", vbCrLf & "<img ", 1, -1, 1)
    sReallyDo = Replace(sReallyDo, "/>", " />", 1, -1, 1)
    sReallyDo = ReplaceAll(sReallyDo, "= ", "=", True)
    sReallyDo = ReplaceAll(sReallyDo, "> ", ">", True)
    sReallyDo = Replace(sReallyDo, "><", ">" & vbCrLf & "<")
    sReallyDo = Trim(sReallyDo)
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True
    '//ȥ��onclick,onload�Ƚű�
    regEx.Pattern = "\s[on].+?=([\""|\'])(.*?)\1"
    sReallyDo = regEx.Replace(sReallyDo, "")
    '//��SRC�������ŵ�ͼƬ��ַ��������
    regEx.Pattern = "<img.*?\ssrc=([^\""\'\s][^\""\'\s>]*).*?>"
    sReallyDo = regEx.Replace(sReallyDo, "<img src=""$1"" />")
    '//����ƥ��ͼƬSRC��ַ
    regEx.Pattern = "<img.*?\ssrc=([\""\'])([^\""\']+?)\1.*?>"
    Set oMatches = regEx.Execute(sReallyDo)
    '//��ͼƬ��ַ��������
    For Each cMatch in oMatches
        iReallyDo = iReallyDo + 1
        ReDim Preserve aReallyDo(iReallyDo)
        aReallyDo(iReallyDo) = regEx.Replace(cMatch.Value, "$2")
    Next
    getIMG = aReallyDo
End Function
%>
<%
'����:��ָ���ļ�¼������������ɸѡ��������һ���µļ�¼������
'��Դ:http://jorkin.reallydo.com/article.asp?id=385

Public Function FilterField(tmpRs, tmpFilter)
    Set FilterField = tmpRs.Clone
    FilterField.Filter = tmpFilter
End Function
%>
<%
'����:��ȡFormֵ��������Ӧ���롣
'��Դ:http://jorkin.reallydo.com/article.asp?id=555

Function OPS(fStyle)
    Dim Key, sTableName
    WriteLn("<div style=""text-align:left;"">")
    Select Case fStyle
        Case 1
            For Each Key in Request.Form
                WriteLn("&nbsp; &nbsp; s" & PCase2(Key, "_") & " = " & IIf(IsNumeric(Request.Form(Key)), "Bint", "BStr") & "(Request.Form("""&Key&"""))<br />")
            Next
        Case 2
            For Each Key in Request.Form
                WriteLn("&nbsp; &nbsp; If IsBlank(s" & PCase2(Key, "_") & ") Then doAlert """ & Key & "����Ϊ�ա�"",""back""<br />")
            Next
        Case 3
            For Each Key in Request.Form
                If InStr(Key, "_") > 0 Then
                    sTableName = PCase(Split(Key, "_")(0))
                    Exit For
                End If
            Next
            WriteLn("&nbsp; &nbsp; sSql = ""Select * From [" & sTableName & "] Where " & sTableName & "_ID = "" & sID" & "<br />")
            WriteLn("&nbsp; &nbsp; Set oRs = Server.CreateObject(""ADODB.RecordSet"")<br />")
            WriteLn("&nbsp; &nbsp; oRs.Open sSql, oConn, 1, 3<br />")
            WriteLn("&nbsp; &nbsp; If oRs.EOF Then<br />")
            WriteLn("&nbsp; &nbsp; &nbsp; &nbsp; oRs.AddNew<br />")
            WriteLn("&nbsp; &nbsp; End If<br />")
            For Each Key in Request.Form
                WriteLn("&nbsp; &nbsp; oRs(""" & PCase2(Key, "_") & """) = s" & PCase2(Key, "_") & "<br />")
            Next
            WriteLn("&nbsp; &nbsp; oRs.Update<br />")
            WriteLn("&nbsp; &nbsp; oRs.Close<br />")
            WriteLn("&nbsp; &nbsp; doAlert ""����ɹ�"",""./""<br />")
        Case 4
            For Each Key in Request.Form
                If InStr(Key, "_") > 0 Then
                    sTableName = PCase(Split(Key, "_")(0))
                    Exit For
                End If
            Next
            WriteLn("If sId > 0 Then<br />")
            WriteLn("&nbsp; &nbsp; ShowInfo sId<br />")
            WriteLn("End If<br /><br />")
            WriteLn("Function ShowInfo(sID)<br />")
            WriteLn("&nbsp; &nbsp; WriteLn(""&lt;scr""&""ipt language=""""javascript""""&gt;"")<br />")
            WriteLn("&nbsp; &nbsp; WriteLn(""var oForm = form1;"")<br />")
            WriteLn("&nbsp; &nbsp; WriteLn(""with(oForm){"")<br />")
            WriteLn("&nbsp; &nbsp; Set oRs = Exec( ""SELECT * FROM [" & sTableName & "] WHERE " & sTableName & "_Id = "" & sId )<br />")
            WriteLn("&nbsp; &nbsp; If Not oRs.Eof Then <br />")
            WriteLn("<br />")
            For Each Key in Request.Form
                WriteLn("&nbsp; &nbsp; &nbsp; &nbsp; s" & PCase2(Key, "_") & " = " & IIf(IsNumeric(Request.Form(Key)), "Bint", "BStr") & "(oRs(""" & PCase2(Key, "_") & """))<br />")
            Next
            WriteLn("<br />")
            For Each Key in Request.Form
                WriteLn("&nbsp; &nbsp; &nbsp; &nbsp; WriteLn(""" & Key & ".value='"" & Str4Js(s" & PCase2(Key, "_") & ") & ""';"")<br />")
            Next
            WriteLn("<br />")
            WriteLn("&nbsp; &nbsp; End If<br />")
            WriteLn("&nbsp; &nbsp; WriteLn(""}"")<br />")
            WriteLn("&nbsp; &nbsp; WriteLn(""&lt;/scr""&""ipt&gt;"")<br />")
            WriteLn("End Function<br />")
        Case 12
            OPS(1)
            OPS(2)
        Case 13
            OPS(1)
            OPS(3)
        Case 123
            OPS(1)
            OPS(2)
            OPS(3)
        Case Else
            WriteLn("sID = Bint(Request.QueryString(""ID""))<br /><br />")
            WriteLn("If isPostBack() And ChkPost() And Request.Form(""editinfo"")=""editinfo"" Then<br />")
            WriteLn("&nbsp; &nbsp; OpenConn()<br />")
            OPS(1)
            OPS(2)
            OPS(3)
            WriteLn("End If<br />")
            OPS(4)
    End Select
    WriteLn("<br />")
    WriteLn("</div>")
End Function
%>
<%
Function PCase2(sString, sSplit)
    Dim i
    sString = BStr(sString)
    sSplit = BStr(sSplit)
    If IsBlank(sSplit) Then sSplit = "_"
    aStr = Split(sString, sSplit)
    For i = 0 To UBound(aStr)
        aStr(i) = UCase(Left(aStr(i), 1)) & Mid(aStr(i), 2)
    Next
    PCase2 = Join(aStr, "_")
End Function
%>
<%
'����:�������ȡ�ͻ�����ʵIP���ͻ�ȡ�ͻ��˵Ĵ���IP, ���ܲ���ǿ�Ĵ����������
'��Դ:http://jorkin.reallydo.com/article.asp?id=165

Private Function getIP()
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
    getIP = Trim(Mid(sIPAddress, 1, 15))
End Function
%>
<%
'����:������ѡ��<input type="radio">
'��Դ:http://jorkin.reallydo.com/article.asp?id=563

Function RadioScript(ByVal FormElement, ByVal ElementValue)
    RadioScript = "<scr" & "ipt language=""javascript"" type=""text/javascript"">" & vbCrLf
    RadioScript = RadioScript & "var Jorkin='" & ElementValue & "';" & vbCrLf
    RadioScript = RadioScript & "  for(i = 0; i < " & FormElement & ".length; i++){" & vbCrLf
    RadioScript = RadioScript & "    if (Jorkin == " & FormElement & "[i].value){" & vbCrLf
    RadioScript = RadioScript & "      " & FormElement & "[i].checked = true}}</scr" & "ipt>" & vbCrLf
End Function
%>
<%
'����:����Ŀ¼��
'MD [Drive:]Path
'֧�ִ����༶Ŀ¼��֧�����·���;���·����
'֧���á�...��ָ����Ŀ¼�ĸ�Ŀ¼��
'��Դ��http://jorkin.reallydo.com/article.asp?id=402
'��ҪPATH����:    http://jorkin.reallydo.com/article.asp?id=401

Function MD(sPath)
    On Error Resume Next
    Dim aPath, iPath, i, sTmpPath
    Dim oFSO
    sPath = Path(sPath) '//�˴���ҪPATH����
    Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
    If oFSO.FolderExists(sPath) Then MD = True : Exit Function
    aPath = Split(sPath, "\")
    iPath = UBound(aPath)
    sTmpPath = ""
    For i = 0 To iPath
        sTmpPath = sTmpPath & aPath(i) & "\"
        If Not oFSO.FolderExists(sTmpPath) Then
            oFSO.CreateFolder(sTmpPath)
        End If
    Next
    Set oFSO = Nothing
    If Err.Number > 0 Then
        Err.Clear()
        MD = False
    Else
        MD = True
    End If
End Function
%>

<%
'AInt��ʽ�����������ַ���
'��Դ:http://jorkin.reallydo.com/article.asp?id=395

Function BTags(s)
    Dim a, i
    s = Bstr(Bstr(s))
    s = Replace(s, "}", "")
    s = Replace(s, "{", "")
    s = Replace(s, "��", ",")
    a = Split(s, ",")
    For i = 0 To UBound(a)
        a(i) = BStr(a(i))
    Next
    BTags = "{" & Join(a, "},{") & "}"
End Function
%>
<%
Sub RE()
    Response.End()
End Sub
%>
<%
Sub RR(s)
    Response.Redirect(s)
End Sub
%>
<%
'�ж�s�Ƿ�����Ч����

Function IsValid(s, av)
    Dim i
    IsValid = False
    If IsArray(av) Then
        For i = 0 To UBound(av)
            If StrComp(BStr(s), BStr(av(i)), 1) = 0 Then
                IsValid = True
                Exit Function
            End If
        Next
    Else
        IsValid = IsValid(s, Split(av, ","))
    End If
End Function
%>
<%
'�����ֵ�ȡ����

Function GetDictMC(sDM, sTABLE, sDefault)
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
%>
<%
'��������ȡ�ֵ�

Function GetDictDM(sDM, sTABLE, sDefault)
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
%>
<%
'�ֵ��Ƿ���Ч

Function IsValidDM(sDM, sTABLE)
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

%>
<%
Function SelectOptions(sDict)
    On Error Resume Next
    Dim oRs, aRs, i
    Set oRs = Exec("Select Count(0) From " & sDict & " Where IsNumeric(DM) = 0 And DM Is Not NUll And Dm <> ''")
    sSql = "Select DM,MC From " & sDict & " Order By "
    If Not oRs(0) > 0 Then sSql = sSql & " Convert(int,DM) ASC, "
    sSql = sSql & "Dm ASC"
    Set oRs = Exec(sSql)
    aRs = GetRsRows(oRs)
    For i = 0 To UBound(aRs, 2)
        WriteLn("&lt;option value=&quot;" & aRs(0, i) & "&quot;" & IIf(i = 0, " selected=&quot;selected&quot;", "") & "&gt;" & aRs(1, i) & "&lt;/option&gt;<br />")
    Next
End Function
%>