<%
p_name = " - 智利红酒" '免费版程序，请保留此字符，去掉会报错！

Const OBJ_RST = "ADODB.Recordset"
Const OBJ_CONN = "ADODB.Connection"
Const OBJ_STRM = "ADODB.Stream"
Const OBJ_FSO = "Scripting.FilesyStemObject"
Const OBJ_XHTP = "MSXML2.XMLHTTP"
Const OBJ_DOM = "MSXML2.DOMDocument"
rqid = Request.QueryString("id")
rqact = Request.QueryString("act")

'/////////基础操作函数部分

'过程：输出字符串[代替Response.Write]

Sub echo(Str)
    response.Write(Str)
End Sub

'过程：结束页面并输出字符串

Sub die(Str)
    response.Write(Str)
    response.End()
End Sub

'函数：十进制转二进制

Function cbit(byval num)
    Dim base64
    Set base64 = New base64_class
    num = base64.cbit(num)
    Set base64 = Nothing
    cbit = num
End Function

'函数：二进制转十进制

Function cdec(byval num)
    Dim base64
    Set base64 = New base64_class
    num = base64.cdec(num)
    Set base64 = Nothing
    cdec = num
End Function

'函数：毫秒数转换为时间长度

Function ctime(byval num, n)
    Dim tmp
    tmp = 0
    If Not IsNumeric(num) Then
        ctime = tmp
        Exit Function
    End If
    If n = "" Or Not IsNumeric(n) Then n = 2
    If num >= 1000 * 60 * 60 * 24 * 30 Then
        tmp = round(num / (1000 * 60 * 60 * 24 * 30), n) & "月"
    ElseIf num >= 1000 * 60 * 60 * 24 Then
        tmp = round(num / (1000 * 60 * 60 * 24), n) & "天"
    ElseIf num >= 1000 * 60 * 60 Then
        tmp = round(num / (1000 * 60 * 60), n) & "小时"
    ElseIf num >= 1000 * 60 Then
        tmp = round(num / (1000 * 60), n) & "分钟"
    ElseIf num >= 1000 Then
        tmp = round(num / 1000, n) & "秒"
    Else
        tmp = round(num, n) & "毫秒"
    End If
    ctime = tmp
End Function

'函数：字节数转换为文件大小

Function csize(byval num, n)
    Dim tmp
    tmp = 0
    If Not IsNumeric(num) Then
        csize = tmp
        Exit Function
    End If
    If n = "" Or Not IsNumeric(n) Then n = 2
    If num >= 1024 * 1024 * 1024 * 1024 Then
        tmp = round(num / (1024 * 1024 * 1024 * 1024), n) & "TB"
    ElseIf num >= 1024 * 1024 * 1024 Then
        tmp = round(num / (1024 * 1024 * 1024), n) & "GB"
    ElseIf num >= 1024 * 1024 Then
        tmp = round(num / (1024 * 1024), n) & "MB"
    ElseIf num >= 1024 Then
        tmp = round(num / 1024, n) & "KB"
    Else
        tmp = round(num, n) & "Byte"
    End If
    csize = tmp
End Function

'函数：将ASP文件运行结果返回为字串

Function ob_get_contents(Path)
    Dim tmp, a, b, t, matches, m
    Dim Str
    Str = file_iread(Path)
    tmp = "dim htm : htm = """""&vbCrLf
    a = 1
    b = InStr(a, Str, "<%") + 2
    While b > a + 1
        t = Mid(Str, a, b - a -2)
        t = Replace(t, vbCrLf, "{::vbcrlf}")
        t = Replace(t, vbCr, "{::vbcr}")
        t = Replace(t, """", """""")
        tmp = tmp & "htm = htm & """ & t & """" & vbCrLf
        a = InStr(b, Str, "%\>") + 2
        tmp = tmp & str_replace("^\s*=", Mid(Str, b, a - b -2), "htm = htm & ") & vbCrLf
        b = InStr(a, Str, "<%") + 2
    Wend
    t = Mid(Str, a)
    t = Replace(t, vbCrLf, "{::vbcrlf}")
    t = Replace(t, vbCr, "{::vbcr}")
    t = Replace(t, """", """""")
    tmp = tmp & "htm = htm & """ & t & """" & vbCrLf
    tmp = Replace(tmp, "response.write", "htm = htm & ", 1, -1, 1)
    tmp = Replace(tmp, "echo", "htm = htm & ", 1, -1, 1)
    'execute(tmp)
    executeglobal(tmp)
    htm = Replace(htm, "{::vbcrlf}", vbCrLf)
    htm = Replace(htm, "{::vbcr}", vbCr)
    ob_get_contents = htm
End Function

'过程：动态包含文件

Sub include(Path)
    echo ob_get_contents(Path)
End Sub

'函数：base64加密

Function base64encode(byval Str)
    If IsNull(Str) Then Exit Function
    Dim base64
    Set base64 = New base64_class
    Str = base64.encode(Str)
    Set base64 = Nothing
    base64encode = Str
End Function

'函数：base64解密

Function base64decode(byval Str)
    If IsNull(Str) Then Exit Function
    Dim base64
    Set base64 = New base64_class
    Str = base64.decode(Str)
    Set base64 = Nothing
    base64decode = Str
End Function

'函数：URL加密

Function urlencode(byval Str)
    If IsNull(Str) Then Exit Function
    Str = server.URLEncode(Str)
    urlencode = Str
End Function

'函数：Escape加密

Function escape(byval Str)
    If IsNull(Str) Then Exit Function
    Dim i, c, a, tmp
    tmp = ""
    For i = 1 To Len(Str)
        c = Mid(Str, i, 1)
        a = ascw(c)
        If (a>= 48 And a<= 57) Or (a>= 65 And a<= 90) Or (a>= 97 And a<= 122) Then
            tmp = tmp & c
        ElseIf InStr("@*_+-./", c) > 0 Then
            tmp = tmp & c
        ElseIf a>0 And a<16 Then
            tmp = tmp & "%0" & Hex(a)
        ElseIf a>= 16 And a<256 Then
            tmp = tmp & "%" & Hex(a)
        Else
            tmp = tmp & "%u" & Hex(a)
        End If
    Next
    escape = tmp
End Function

'函数：Escape解密

Function unescape(byval Str)
    If IsNull(Str) Then Exit Function
    Dim i, c, tmp
    tmp = ""
    For i = 1 To Len(Str)
        c = Mid(Str, i, 1)
        If Mid(Str, i, 2) = "%u" And i<= Len(Str) -5 Then
            If IsNumeric("&H" & Mid(Str, i + 2, 4)) Then
                tmp = tmp & chrw(CInt("&H" & Mid(Str, i + 2, 4)))
                i = i + 5
            Else
                tmp = tmp & c
            End If
        ElseIf c = "%" And i<= Len(Str) -2 Then
            If IsNumeric("&H" & Mid(Str, i + 1, 2)) Then
                tmp = tmp & chrw(CInt("&H" & Mid(Str, i + 1, 2)))
                i = i + 2
            Else
                tmp = tmp & c
            End If
        Else
            tmp = tmp & c
        End If
    Next
    unescape = tmp
End Function

'函数：md5加密

Function md5(byval Str)
    If IsNull(Str) Then Exit Function
    Dim md5_cls
    Set md5_cls = New md5_class
    Str = md5_cls.md5(Str)
    Set md5_cls = Nothing
    md5 = Str
End Function

'函数：三元IF

Function IIf(Exp, v1, v2)
    Dim tmp
    tmp = v2
    If Exp Then tmp = v1
    IIf = tmp
End Function

'函数：空值测试

Function inull(Val)
    Dim tmp
    tmp = False
    If IsNull(Val) Then
        tmp = True
    ElseIf IsEmpty(Val) Then
        tmp = True
    ElseIf Trim(Val) = "" Then
        tmp = True
    End If
    inull = tmp
End Function

'全启变量：客户端IP
Dim ip
ip = request.ServerVariables("REMOTE_ADDR")


'函数：返回客户端真实IP

Function realip
    Dim tmp
    tmp = request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If Trim(tmp) = "" Then tmp = request.ServerVariables("REMOTE_ADDR")
    realip = tmp
End Function

'函数：邮件发送[Jamil-Message]

Function sendmail(fromname, sendto, subject, body, from, serveraddress, username, password)
    Dim jmail, Return
    Set jmail = server.CreateObject("JMAIL.Message")
    jmail.silent = True
    jmail.logging = True
    jmail.charset = "gbk"
    jmail.contenttype = "text/html; charset=gbk"
    jmail.addrecipient sendto
    jmail.fromname = fromname
    jmail.from = from
    jmail.mailserverusername = username
    jmail.MailServerPassword = password
    jmail.subject = subject
    jmail.body = body
    jmail.priority = 3
    Return = jmail.send(serveraddress)
    jmail.Close()
    Set jmail = Nothing
    sendmail = Return
End Function

'函数：检测组件是否安装

Function install(Str)
    Dim tmp
    tmp = False
    Dim obj_test
    On Error Resume Next
    Err.Clear()
    Set obj_test = server.CreateObject(Str)
    If Err.Number = 0 Then tmp = True
    Set obj_test = Nothing
    Err.Clear()
    install = tmp
End Function

'/////////字符串操作函数部分

'函数：正则验证

Function str_test(Pattern, Str)
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        tmp = .Test(Str)
    End With
    Set reg = Nothing
    str_test = tmp
End Function

'函数：正则替换[不区分大小写]

Function str_replace(Pattern, byval Str, s)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        tmp = .Replace(Str, s)
    End With
    Set reg = Nothing
    str_replace = tmp
End Function

'函数：正则替换[区分大小写]

Function str_ireplace(Pattern, byval Str, s)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = False
        .Global = True
        .Pattern = Pattern
        tmp = .Replace(Str, s)
    End With
    Set reg = Nothing
    str_ireplace = tmp
End Function

'函数：执行正则搜索并返回结果集[不区分大小写]

Function str_execute(Pattern, byval Str)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = Pattern
        Set tmp = .Execute(Str)
    End With
    Set reg = Nothing
    Set str_execute = tmp
End Function

'函数：执行正则搜索并返回结果集[区分大小写]

Function str_iexecute(Pattern, byval Str)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = False
    Dim reg
    Set reg = New regexp
    With reg
        .IgnoreCase = False
        .Global = True
        .Pattern = Pattern
        Set tmp = .Execute(Str)
    End With
    Set reg = Nothing
    Set str_iexecute = tmp
End Function

'函数：精确计算字符串长度

Function str_len(byval Str)
    Str = str_replace("[^\x00-\xff]", Str, "@@")
    str_len = Len(Str)
End Function

'函数：截断字串

Function str_left(byval Str, slen, ext)
    If IsNull(Str) Then Exit Function
    Dim tmp
    tmp = "&quot;=""|&amp;=&|&lt;=<|&gt;=>|&euro;=�|&nbsp;= |&laquo;=?|&raquo;=?|&hellip;=…|&copy;=?"
    Dim arr, a, v
    arr = Split(tmp, "|")
    For Each v in arr
        a = Split(v, "=")
        Str = Replace(Str, a(0), a(1))
    Next
    'die str
    Dim i, c, s, n
    n = 0
    tmp = ""
    For i = 1 To Len(Str)
        s = Mid(Str, i, 1)
        c = Abs(ascw(s))
        If c>255 Then n = n + 2 Else n = n + 1
        tmp = tmp & s
        If n >= slen Then Exit For
    Next
    If tmp = Str Then ext = ""
    str_left = tmp & ext
End Function

'函数：返回可安全地用于SQL操作的字符串

Function str_safe(byval Str)
    If IsNull(Str) Then Exit Function
    Str = str_isafe(Str)
    Str = Replace(Str, "<", "&lt;")
    Str = Replace(Str, ">", "&gt;")
    Str = Replace(Str, """", "&quot;")
    str_safe = Str
End Function

'函数：SQL关键词过滤 用于获取含HTML标签的内容

Function str_isafe(byval Str)
    If IsNull(Str) Then Exit Function
    Str = Replace(Str, "select ", "sel&#101;ct ", 1, -1, 1)
    Str = Replace(Str, "insert ", "ins&#101;rt ", 1, -1, 1)
    Str = Replace(Str, "update ", "up&#100;ate ", 1, -1, 1)
    Str = Replace(Str, "delete ", "del&#101;te ", 1, -1, 1)
    Str = Replace(Str, " and", " an&#100; ", 1, -1, 1)
    Str = Replace(Str, "drop table", "dro&#112; table", 1, -1, 1)
    Str = Replace(Str, "script", "&#115;cript")
    Str = Replace(Str, "*", "&#42;")
    Str = Replace(Str, "%", "&#37;")
    Str = Replace(Str, "'", "''")
    str_isafe = Str
End Function

'函数：替换简单HTML格式字符为控制字符

Function str_htmldecode(byval Str)
    If IsNull(Str) Then Exit Function
    Str = Replace(Str, "&nbsp;", " ")
    Str = Replace(Str, "<br />", Chr(10))
    str_htmldecode = Str
End Function

'函数：替换字符串中的控制字符为HTML代码。

Function str_htmlencode(byval Str)
    If IsNull(Str) Then Exit Function
    Str = Replace(Str, " ", "&nbsp;")
    Str = Replace(Str, Chr(10), "<br />")
    str_htmlencode = Str
End Function

'函数：清除HTML标签

Function str_htmlclear(byval Str)
    If IsNull(Str) Then Exit Function
    Str = Replace(Str, "&nbsp;", " ")
    Dim Pattern
    Pattern = "<[^>]+?>"
    Str = str_replace(Pattern, Str, "")
    str_htmlclear = Str
End Function

'函数：清除所有格式及空格 压缩字符串

Function str_trim(byval Str)
    If IsNull(Str) Then Exit Function
    Str = Replace(Str, Chr(10), "")
    Str = Replace(Str, Chr(13), "")
    Dim Pattern
    Pattern = "<[^>]+?>"
    Str = str_replace(Pattern, Str, "")
    Str = Replace(Str, "&nbsp;", "")
    Str = Replace(Str, " ", "")
    str_trim = Str
End Function

'函数：返回一个不重复的随机字串

Function str_rnd()
    Dim ran_num, dt_now, tmp
    dt_now = Now()
    Randomize
    ran_num = Int( (90000 * Rnd) + 10000 )
    tmp = Year(dt_now) & Right("0"&Month(dt_now), 2) & Right("0"&Day(dt_now), 2) & Right("0"&Hour(dt_now), 2) &_
    Right("0"&Minute(dt_now), 2) & Right("0"&Second(dt_now), 2) & ran_num
    str_rnd = base64encode(tmp)
End Function

'函数：返回格式化的时间字串

Function str_time(Format, byval Str)
    If Trim(Str) = "" Or Not IsDate(Str) Then Exit Function
    Dim tmp
    tmp = Format
    tmp = Replace(tmp, "yy", Right("0"&Year(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "y", Year(Str), 1, -1, 1)
    tmp = Replace(tmp, "mm", Right("0"&Month(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "m", Month(Str), 1, -1, 1)
    tmp = Replace(tmp, "dd", Right("0"&Day(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "d", Day(Str), 1, -1, 1)
    tmp = Replace(tmp, "hh", Right("0"&Hour(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "h", Hour(Str), 1, -1, 1)
    tmp = Replace(tmp, "ii", Right("0"&Minute(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "i", Minute(Str), 1, -1, 1)
    tmp = Replace(tmp, "ss", Right("0"&Second(Str), 2), 1, -1, 1)
    tmp = Replace(tmp, "s", Second(Str), 1, -1, 1)
    str_time = tmp
End Function

'函数：从字串中分离出远程文件URL

Function str_geturl(byval Str, ext)
    If IsNull(Str) Then Exit Function
    Dim exts
    exts = Split(ext, ",")
    Dim Pattern, e, s
    Pattern = ""
    s = ""
    For Each e in exts
        Pattern = Pattern & s & "http://[\S]+?\."&e
        s = "|"
    Next
    Dim matches
    Set matches = str_execute(Pattern, Str)
    Dim m, urls
    urls = ""
    s = ""
    For Each m in matches
        urls = urls & s & m.Value
        s = "#"
    Next
    str_geturl = Split(urls, "#")
End Function

'函数：获取URL参数串

Function str_query(del)
    Dim tmp
    tmp = request.ServerVariables("QUERY_STRING")
    If Trim(del) = "" Then
        str_query = tmp
        Exit Function
    End If
    Dim arr
    arr = Split(tmp, "&")
    Dim q, a, t
    t = ""
    tmp = ""
    For Each q in arr
        If Trim(q) <> "" Then
            a = Split(q, "=")
            If UBound(a) = 0 Then arr_push a, ""
            If Not arr_in(Split(del, ","), a(0)) Then
                tmp = tmp&t&a(0)&"="&a(1)
                t = "&"
            End If
        End If
    Next
    str_query = tmp
End Function

'函数：字符串加密

Function str_encode(byval Str)
    If IsNull(Str) Then Exit Function
    Dim base64
    Set base64 = New base64_class
    base64.bstr = "ABCDEF1234GHIJKLMnopqrs+tuvwxyz09abcdef!ghijklmNOPQRS5678TUVWXYZ"
    base64.blen = 16
    Str = base64.encode(Str)
    Set base64 = Nothing
    str_encode = Str
End Function

'函数：字符串解密

Function str_decode(byval Str)
    If IsNull(Str) Then Exit Function
    Dim base64
    Set base64 = New base64_class
    base64.bstr = "ABCDEF1234GHIJKLMnopqrs+tuvwxyz09abcdef!ghijklmNOPQRS5678TUVWXYZ"
    base64.blen = 16
    Str = base64.decode(Str)
    Set base64 = Nothing
    str_decode = Str
End Function

'/////////文件操作函数部分

'函数：获取当前脚本执行文件的文件名

Function file_self()
    Dim tmp
    tmp = request.ServerVariables("SCRIPT_NAME")
    tmp = Split(tmp, "/")
    file_self = tmp(UBound(tmp))
End Function

'函数：获取当前脚本执行文件所在的磁盘目录

Function file_dir()
    Dim tmp, arr
    tmp = request.ServerVariables("SCRIPT_NAME")
    arr = Split(tmp, "/")
    tmp = arr(UBound(arr))
    arr = Split(server.MapPath(tmp), "\")
    file_dir = arr(UBound(arr) -1)
End Function

'函数：检测文件/文件夹是否存在

Function file_exists(Path)
    Dim tmp
    tmp = False
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(Path)) Then tmp = True
    If fso.FolderExists(server.MapPath(Path)) Then tmp = True
    Set fso = Nothing
    file_exists = tmp
End Function

'函数：删除文件/文件夹

Function file_delete(Path)
    Dim tmp
    tmp = False
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(Path)) Then'目标是文件
        fso.DeleteFile(server.MapPath(Path))
        If Not fso.FileExists(server.MapPath(Path)) Then tmp = True
    End If
    If fso.FolderExists(server.MapPath(Path)) Then'目标是文件夹
        fso.DeleteFolder(server.MapPath(Path))
        If Not fso.FolderExists(server.MapPath(Path)) Then tmp = True
    End If
    Set fso = Nothing
    file_delete = tmp
End Function

'函数：获取文件/文件夹信息

Function file_info(Path)
    Dim tmp(4)
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(Path)) Then '目标是文件
        Dim fl
        Set fl = fso.GetFile(server.MapPath(Path))
        tmp(0) = fl.Type'类型
        tmp(1) = fl.Attributes'属性
        tmp(2) = csize(fl.Size, 4)'大小
        tmp(3) = fl.DateCreated'创建时间
        tmp(4) = fl.DateLastModified'最后修改时间
    ElseIf fso.FolderExists(server.MapPath(Path)) Then '目标是文件夹
        Dim fd
        Set fd = fso.GetFolder(server.MapPath(Path))
        tmp(0) = "folder"'类型
        tmp(1) = fd.Attributes'属性
        tmp(2) = csize(fd.Size, 4)'大小
        tmp(3) = fd.DateCreated'创建时间
        tmp(4) = fd.DateLastModified'最后修改时间
    End If
    Set fso = Nothing
    file_info = tmp
End Function

'函数：复制文件/文件夹

Function file_copy(file_start, file_end, model)
    If model<>0 And model<>1 Then model = False Else model = CBool(model)
    Dim tmp
    tmp = False
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(file_start)) Then '目标是文件
        fso.CopyFile server.MapPath(file_start), server.MapPath(file_end), model
        If fso.FileExists(server.MapPath(file_end)) Then tmp = True
    End If
    If fso.FolderExists(server.MapPath(file_start)) Then '目标是文件夹
        fso.CopyFolder server.MapPath(file_start), server.MapPath(file_end), model
        If fso.FolderExists(server.MapPath(file_end)) Then tmp = True
    End If
    Set fso = Nothing
    file_copy = tmp
End Function

'函数：创建文件夹

Function file_create(Path, model)
    If model<>0 And model<>1 Then model = False Else model = CBool(model)
    Dim tmp
    tmp = False
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FolderExists(server.MapPath(Path)) Then
        If model Then
            fso.DeleteFolder(server.MapPath(Path))
            fso.CreateFolder server.MapPath(Path)
        End If
    Else
        fso.CreateFolder server.MapPath(Path)
    End If
    If fso.FolderExists(server.MapPath(Path)) Then tmp = True
    Set fso = Nothing
    file_create = tmp
End Function

'函数：获取指定目录下所有文件及文件夹列表

Function file_list(Path)
    If Not file_exists(Path) Then
        file_list = Array("", "")
        Exit Function
    End If
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    Dim fdr
    Set fdr = fso.GetFolder( server.MapPath(Path) )
    Dim Folders
    Set Folders = fdr.SubFolders
    Dim f, t, tmp
    t = ""
    tmp = ""
    For Each f in Folders
        tmp = tmp & t & f.Name
        t = "|"
    Next
    tmp = tmp & "*"
    t = ""
    Dim Files
    Set Files = fdr.Files
    For Each f in Files
        tmp = tmp & t & f.Name
        t = "|"
    Next
    Set fso = Nothing
    file_list = Split(tmp, "*")'返回长度为二的字符数组
End Function

'函数：返回图片类型及尺寸

Function file_imginfo(Path)
    Dim tmp
    tmp = Array("", 0, 0)
    Dim fso
    Set fso = server.CreateObject(OBJ_FSO)
    If fso.FileExists(server.MapPath(Path)) Then
        Dim img
        Set img = LoadPicture(server.MapPath(Path))
        Select Case img.Type
            Case 0
                tmp(0) = "none"'类型
            Case 1
                tmp(0) = "bitmap"
            Case 2
                tmp(0) = "metafile"
            Case 3
                tmp(0) = "ico"
            Case 4
                tmp(0) = "win32-enhanced metafile"
        End Select
        tmp(1) = round(img.Width / 26.4583)'宽度
        tmp(2) = round(img.height / 26.4583)'高度
        Set img = Nothing
        Set fso = Nothing
    End If
    file_imginfo = tmp
End Function


'函数：检测图片文件合法性

Function file_isimg(Path)
    Dim tmp
    tmp = False
    If Not file_exists(Path) Then
        file_isimg = tmp
        Exit Function
    End If
    Dim jpg(1)
    jpg(0) = CByte(&HFF)
    jpg(1) = CByte(&HD8)
    Dim bmp(1)
    bmp(0) = CByte(&H42)
    bmp(1) = CByte(&H4D)
    Dim png(3)
    png(0) = CByte(&H89)
    png(1) = CByte(&H50)
    png(2) = CByte(&H4E)
    png(3) = CByte(&H47)
    Dim gif(5)
    gif(0) = CByte(&H47)
    gif(1) = CByte(&H49)
    gif(2) = CByte(&H46)
    gif(3) = CByte(&H39)
    gif(4) = CByte(&H38)
    gif(5) = CByte(&H61)
    Dim fstream, fext, stamp, i
    fext = Mid(Path, instrrev(Path, ".") + 1)
    Set fstream = server.CreateObject(OBJ_STRM)
    fstream.Open
    fstream.Type = 1
    fstream.loadfromfile server.MapPath(Path)
    fstream.position = 0
    Select Case fext
Case "jpg", "jpeg":
        stamp = fstream.Read(2)
        For i = 0 To 1
            If ascb(midb(stamp, i + 1, 1)) = jpg(i) Then tmp = True Else tmp = False
        Next
Case "gif":
        stamp = fstream.Read(6)
        For i = 0 To 5
            If ascb(midb(stamp, i + 1, 1)) = gif(i) Then tmp = True Else tmp = False
        Next
Case "png":
        stamp = fstream.Read(4)
        For i = 0 To 3
            If ascb(midb(stamp, i + 1, 1)) = png(i) Then tmp = True Else tmp = False
        Next
Case "bmp":
        stamp = fstream.Read(2)
        For i = 0 To 1
            If ascb(midb(stamp, i + 1, 1)) = bmp(i) Then tmp = True Else tmp = False
        Next
    End Select
    fstream.Close
    Set fstream = Nothing
    file_isimg = tmp
End Function

'函数：采集远程文件并保存到本地磁盘

Function file_savefromurl(fileurl, savepath, savetype)
    If savetype<>1 And savetype<>2 Then savetype = 2
    Dim xmlhttp
    Set xmlhttp = server.CreateObject(OBJ_XHTP)
    With xmlhttp
        .Open "get", fileurl, False
        .send()
        Dim fl
        fl = .responsebody
    End With
    Set xmlhttp = Nothing
    Dim stream
    Set stream = server.CreateObject(OBJ_STRM)
    With stream
        .Type = savetype
        .Open
        .Write fl
        .savetofile server.MapPath(savepath), 2
        .cancel()
        .Close()
    End With
    Set stream = Nothing
    file_savefromurl = file_exists(savepath)
End Function

'函数：读取文件内容到字符串

Function file_read(Path)
    Dim tmp
    tmp = ""
    If Left(Path, 7) = "http://" Then '读取远程文件
        Dim xmlhttp
        Set xmlhttp = server.CreateObject(OBJ_XHTP)
        With xmlhttp
            .Open "get", Path, False
            .send()
            tmp = .responsetext
        End With
        Set xmlhttp = Nothing
    Else '读取本地文件
        If Not file_exists(Path) Then
            file_read = tmp
            Exit Function
        End If
        Dim stream
        Set stream = server.CreateObject(OBJ_STRM)
        With stream
            .Type = 2 '文本类型
            .mode = 3 '读写模式
            .charset = "gbk"
            .Open
            .loadfromfile(server.MapPath(Path))
            tmp = .readtext()
        End With
        stream.Close
        Set stream = Nothing
    End If
    file_read = tmp
End Function

'函数：保存字符串到文件

Function file_save(Str, Path, model)
    If model<>0 And model<>1 Then model = 1
    If model = 0 And file_exists(Path) Then
        file_save = True
        Exit Function
    End If
    Dim stream
    Set stream = server.CreateObject(OBJ_STRM)
    With stream
        .Type = 2 '文本类型
        .charset = "gbk"
        .Open
        .writetext Str
        .savetofile(server.MapPath(Path)), model + 1
    End With
    stream.Close
    Set stream = Nothing
    file_save = file_exists(Path)
End Function

'函数:读取ASP类型文件的全部内容

Function file_iread(Path)
    Dim Str
    Str = file_read(Path)
    Dim Pattern
    Pattern = "<\!--#include[ ]+?file[ ]*?=[ ]*?""(\S+?)""--\>"
    Dim matches
    Set matches = str_execute(Pattern, Str)
    Dim m, f, tmp
    For Each m in matches
        f = Mid(Path, 1, instrrev(Path, "/"))&m.submatches(0)
        tmp = file_read(f)
        If str_test(Pattern, tmp) Then tmp = file_iread(f) '处理子包含
        Str = Replace(Str, m.Value, tmp)
    Next
    Pattern = "<%@[ ]*?LANGUAGE[ ]*?=[ ]*?""[a-zA-Z]+?""[ ]+?CODEPAGE[ ]*?=[ ]*?""[0-9]+?""[ ]*?%\>"
    Str = str_replace(Pattern, Str, "")
    file_iread = Str
End Function

'/////////数组操作函数部分

'函数：检测元素是否是指定数组的元素成员

Function arr_in(arr, Val)
    Dim a, tmp
    tmp = False
    For Each a in arr
        If Trim(a) = Trim(Val) Then
            tmp = True
            Exit For
        End If
    Next
    arr_in = tmp
End Function

'函数：指定字串数组的元素是否含有指定字串

Function arr_strin(arr, Str)
    Dim a, tmp
    tmp = False
    For Each a in arr
        If InStr(1, a, Str, 1)<>0 Then
            tmp = True
            Exit For
        End If
    Next
    arr_strin = tmp
End Function

'函数：动态向数组中添加新元素

Function arr_push(arr, Val)
    ReDim preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = Val
    arr_push = arr
End Function

'函数：获取元素在数组中首次出现时的索引值

Function arr_getindex(arr, Str)
    Dim i, tmp
    tmp = -1
    For i = 0 To UBound(arr)
        If arr(i) = Str Then
            tmp = i
            Exit For
        End If
    Next
    arr_getindex = tmp
End Function

'/////////XML解析操作函数部分

'函数：载入xml文件并返回操作对象

Function xml_load(Path)
    Dim obj_xml
    Set obj_xml = server.CreateObject(OBJ_DOM)
    obj_xml.load Server.MapPath(Path)
    Set xml_load = obj_xml
End Function

'/////////数据操作函数部分

'函数：执行SQL语句,只读方式。

Function ado_query(byval sql)
    Set ado_query = ado_iquery(sql, conn, 3, 1)
End Function

'函数：执行SQL语句,可修改方式。

Function ado_query_modify(byval sql)
    Set ado_query_modify = ado_iquery(sql, conn, 3, 2)
End Function


'函数：执行SQL语句

Function ado_iquery(byval sql, conn, cursortype, locktype)
    If Trim(sql) = "" Then Exit Function
    If Trim(n) = "" Or Not IsNumeric(n) Then n = 1
    Dim rs
    If LCase(Left(LTrim(sql), 6)) = "select" Then
        Set rs = server.CreateObject(OBJ_RST)
        rs.cursorlocation = 3
        rs.Open sql, conn, cursortype, locktype
    Else
        Set rs = conn.Execute(sql)
    End If
    Set ado_iquery = rs
End Function

'/////////翻页操作函数部分

'函数：翻页预处理

Function pageturner_handle(byval sql, field_id, page_size)
    pageturner_handle = pageturner_ihandle(sql, field_id, page_size, conn)
End Function

'函数：翻页预处理

Function pageturner_ihandle(sql, field_id, page_size, conn)
    '获取总记录数：page_sum
    Dim rs, page_sum, page_num
    Set rs = ado_iquery(sql, conn, 3, 1)
    page_sum = rs.recordcount
    '计算总页数：page_num
    rs.pagesize = page_size
    page_num = rs.pagecount
    '获取翻页参数
    Dim page
    page = request.QueryString("page")
    If IsEmpty(page) Or Not IsNumeric(page) Then page = 1
    If CDbl(page) <= 0 Then page = 1
    If CDbl(page) > CDbl(page_num) Then page = page_num
    '获取当前页ID列表
    Dim i, s, Filter
    s = ""
    Filter = field_id&"="
    If Not rs.EOF Then rs.absolutepage = page
    For i = 1 To page_size
        If Not rs.EOF Then
            Filter = Filter & s & rs(field_id)
            s = " or "&field_id&"="
            rs.movenext
        End If
    Next
    'die filter
    If page_sum>0 Then rs.Filter = Filter
    '返回数组
    pageturner_ihandle = Array(rs, page, page_num, page_sum)
End Function

'函数：返回翻页条

Function pageturner_show(page, page_num, page_sum, page_size, page_len)
    Dim page_start, page_end, page_link, tmp, p
    '起始页、结束页
    page_start = page - page_len
    page_end = page + page_len
    If CDbl(page_start) <= 0 Then
        page_end = page_end + Abs(page_start)
        page_start = 1
    End If
    If CDbl(page_end) > CDbl(page_num) Then page_end = page_num
    '翻页链接
    'page_link="?" : if str_query("page")<>"" then page_link = "?" & str_query("page") & "&"
    page_link = "?"
    tmp = str_query("page")
    If tmp<>"" Then page_link = "?"&tmp&"&"
    '翻页条开始
    Dim page_back, page_next
    tmp = "<div class=""page_turner"">"
    If CDbl(page) = 1 Then
        page_back = "<a title=""上一页"" href=""javascript:void(0)"">上一页</a>"
    Else
        page_back = "<a title=""上一页"" href="""& page_link & "page="& (page -1) &""">上一页</a>"
    End If'上一页
    If CDbl(page) > page_len + 1 Then tmp = tmp & "<a title=""首页"" href="""& page_link & "page=1"">1...</a>"'首页
    For p = page_start To page_end
        If CDbl(p) = CDbl(page) Then
            tmp = tmp & "<a title=""第"& p &"页"" class=""c"">"& p &"</a>"
        Else
            tmp = tmp & "<a title=""第"& p &"页"" href="""& page_link &"page="& p &""">"& p &"</a>"
        End If
    Next'第_页
    If CDbl(page) = CDbl(page_num) Then
        page_next = "<a title=""下一页"" href=""javascript:void(0)"">下一页</a>"
    Else
        page_next = "<a title=""下一页"" href="""& page_link & "page="& (page+1) &""">下一页</a>"
    End If'下一页
    If CDbl(page)<CDbl(page_num) - page_len Then tmp = tmp&"<a title=""末页"" href="""&page_link&"page="& page_num &""">..."&page_num&"</a>"'末页
    tmp = tmp & page_back & page_next
    tmp = tmp & "<span>每页"& page_size &"条&nbsp;共<label id=""total"">"& page_sum &"</label>条</span>"
    tmp = tmp & "</div>"
    pageturner_show = tmp
End Function

'/////////base6 class for VBs

Class base64_class
    Private blen_
    Private bstr_

    Public Property Get bstr
        bstr = bstr_
    End Property

    Public Property Let bstr(Val)
        bstr_ = Val
    End Property

    Public Property Get blen
        blen = blen_
    End Property

    Public Property Let blen(Val)
        blen_ = Val
    End Property

    Private Sub class_initialize
        bstr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        blen = 8
    End Sub

    'private sub class_terminate

    'end sub

    Public Function cbit(num)
        Dim cbitstr
        cbitstr = ""
        If Len(num)>0 And IsNumeric(num) Then
            Do While Not num \ 2 < 1
                cbitstr = (num Mod 2) & cbitstr
                num = num \ 2
            Loop
        End If
        cbit = num & cbitstr
    End Function

    Public Function cdec(num)
        Dim inum, cdecstr
        cdecstr = 0
        If Len(num)>0 And IsNumeric(num) Then
            For inum = 0 To Len(num) -1
                cdecstr = cdecstr + 2^inum * CInt(Mid(num, Len(num) - inum, 1))
            Next
        End If
        cdec = cdecstr
    End Function

    Public Function encode(Str)
        If Not Len(Str)>0 Then Exit Function
        Dim i, t, s, encodestr
        t = ""
        s = ""
        encodestr = ""
        For i = 1 To Len(Str)
            't = abs(ascw(mid(str,i,1)))
            t = ascw(Mid(Str, i, 1))
            If t<0 Then t = t + 65536
            t = cbit(t)
            If Len(t)<blen Then t = String(blen - Len(t), "0") & t
            s = s & t
        Next
        If Len(s) Mod 6 <> 0 Then s = s & String(6 - (Len(s) Mod 6), "0")
        t = ""
        For i = 1 To Len(s) \ 6
            t = cdec(Mid(s, i * 6 -6 + 1, 6))
            encodestr = encodestr & Mid(bstr, t + 1, 1)
        Next
        If Len(encodestr)<4 Then encodestr = encodestr & String(4 - Len(encodestr), "=")
        encode = encodestr
    End Function

    Public Function decode(Str)
        If Not Len(Str)>0 Then Exit Function
        Dim i, t, s, decodestr
        t = ""
        s = ""
        decodestr = ""
        Str = Replace(Str, "=", "")
        For i = 1 To Len(Str)
            t = cbit(InStr(bstr, Mid(Str, i, 1)) - 1)
            If Len(t)<6 Then t = String(6 - Len(t), "0") & t
            s = s & t
        Next
        If Len(s) Mod blen <> 0 Then s = Left(s, Len(s) - (Len(s) Mod blen))
        t = ""
        For i = 1 To Len(s) \ blen
            t = cdec(Mid(s, i * blen - blen + 1, blen))
            decodestr = decodestr & chrw(t)
        Next
        decode = decodestr
    End Function

End Class

'/////////md5 class for VBs
'If p_name = " - Powered By DKCMS" Then
'Else
'die""
'End If
Class md5_class
    Private BITS_TO_A_BYTE
    Private BYTES_TO_A_WORD
    Private BITS_TO_A_WORD
    Private m_lOnBits(30)
    Private m_l2Power(30)

    Private Sub class_initialize
        BITS_TO_A_BYTE = 8
        BYTES_TO_A_WORD = 4
        BITS_TO_A_WORD = 32
    End Sub

    Private Function LShift(lValue, iShiftBits)
        If iShiftBits = 0 Then
            LShift = lValue
            Exit Function
        ElseIf iShiftBits = 31 Then
            If lValue And 1 Then
                LShift = &H80000000
            Else
                LShift = 0
            End If
            Exit Function
        ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
            Err.Raise 6
        End If
        If (lValue And m_l2Power(31 - iShiftBits)) Then
            LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
        Else
            LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        End If
    End Function

    Private Function RShift(lValue, iShiftBits)
        If iShiftBits = 0 Then
            RShift = lValue
            Exit Function
        ElseIf iShiftBits = 31 Then
            If lValue And &H80000000 Then
                RShift = 1
            Else
                RShift = 0
            End If
            Exit Function
        ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
            Err.Raise 6
        End If
        RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
        If (lValue And &H80000000) Then
            RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
        End If
    End Function

    Private Function RotateLeft(lValue, iShiftBits)
        RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
    End Function

    Private Function AddUnsigned(lX, lY)
        Dim lX4
        Dim lY4
        Dim lX8
        Dim lY8
        Dim lResult

        lX8 = lX And &H80000000
        lY8 = lY And &H80000000
        lX4 = lX And &H40000000
        lY4 = lY And &H40000000

        lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

        If lX4 And lY4 Then
            lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
        ElseIf lX4 Or lY4 Then
            If lResult And &H40000000 Then
                lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
            Else
                lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
            End If
        Else
            lResult = lResult Xor lX8 Xor lY8
        End If
        AddUnsigned = lResult
    End Function

    Private Function md5_F(x, y, z)
        md5_F = (x And y) Or ((Not x) And z)
    End Function

    Private Function md5_G(x, y, z)
        md5_G = (x And z) Or (y And (Not z))
    End Function

    Private Function md5_H(x, y, z)
        md5_H = (x Xor y Xor z)
    End Function

    Private Function md5_I(x, y, z)
        md5_I = (y Xor (x Or (Not z)))
    End Function

    Private Sub md5_FF(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub md5_GG(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub md5_HH(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Sub md5_II(a, b, c, d, x, s, ac)
        a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
        a = RotateLeft(a, s)
        a = AddUnsigned(a, b)
    End Sub

    Private Function ConvertToWordArray(sMessage)
        Dim lMessageLength
        Dim lNumberOfWords
        Dim lWordArray()
        Dim lBytePosition
        Dim lByteCount
        Dim lWordCount
        Dim MODULUS_BITS
        MODULUS_BITS = 512
        Dim CONGRUENT_BITS
        CONGRUENT_BITS = 448
        lMessageLength = Len(sMessage)
        lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
        ReDim lWordArray(lNumberOfWords - 1)
        lBytePosition = 0
        lByteCount = 0
        Do Until lByteCount >= lMessageLength
            lWordCount = lByteCount \ BYTES_TO_A_WORD
            lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
            lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
            lByteCount = lByteCount + 1
        Loop
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
        lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
        lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
        ConvertToWordArray = lWordArray
    End Function

    Private Function WordToHex(lValue)
        Dim lByte
        Dim lCount
        For lCount = 0 To 3
            lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
            WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
        Next
    End Function

    Public Function MD5(sMessage)
        m_lOnBits(0) = CLng(1)
        m_lOnBits(1) = CLng(3)
        m_lOnBits(2) = CLng(7)
        m_lOnBits(3) = CLng(15)
        m_lOnBits(4) = CLng(31)
        m_lOnBits(5) = CLng(63)
        m_lOnBits(6) = CLng(127)
        m_lOnBits(7) = CLng(255)
        m_lOnBits(8) = CLng(511)
        m_lOnBits(9) = CLng(1023)
        m_lOnBits(10) = CLng(2047)
        m_lOnBits(11) = CLng(4095)
        m_lOnBits(12) = CLng(8191)
        m_lOnBits(13) = CLng(16383)
        m_lOnBits(14) = CLng(32767)
        m_lOnBits(15) = CLng(65535)
        m_lOnBits(16) = CLng(131071)
        m_lOnBits(17) = CLng(262143)
        m_lOnBits(18) = CLng(524287)
        m_lOnBits(19) = CLng(1048575)
        m_lOnBits(20) = CLng(2097151)
        m_lOnBits(21) = CLng(4194303)
        m_lOnBits(22) = CLng(8388607)
        m_lOnBits(23) = CLng(16777215)
        m_lOnBits(24) = CLng(33554431)
        m_lOnBits(25) = CLng(67108863)
        m_lOnBits(26) = CLng(134217727)
        m_lOnBits(27) = CLng(268435455)
        m_lOnBits(28) = CLng(536870911)
        m_lOnBits(29) = CLng(1073741823)
        m_lOnBits(30) = CLng(2147483647)
        m_l2Power(0) = CLng(1)
        m_l2Power(1) = CLng(2)
        m_l2Power(2) = CLng(4)
        m_l2Power(3) = CLng(8)
        m_l2Power(4) = CLng(16)
        m_l2Power(5) = CLng(32)
        m_l2Power(6) = CLng(64)
        m_l2Power(7) = CLng(128)
        m_l2Power(8) = CLng(256)
        m_l2Power(9) = CLng(512)
        m_l2Power(10) = CLng(1024)
        m_l2Power(11) = CLng(2048)
        m_l2Power(12) = CLng(4096)
        m_l2Power(13) = CLng(8192)
        m_l2Power(14) = CLng(16384)
        m_l2Power(15) = CLng(32768)
        m_l2Power(16) = CLng(65536)
        m_l2Power(17) = CLng(131072)
        m_l2Power(18) = CLng(262144)
        m_l2Power(19) = CLng(524288)
        m_l2Power(20) = CLng(1048576)
        m_l2Power(21) = CLng(2097152)
        m_l2Power(22) = CLng(4194304)
        m_l2Power(23) = CLng(8388608)
        m_l2Power(24) = CLng(16777216)
        m_l2Power(25) = CLng(33554432)
        m_l2Power(26) = CLng(67108864)
        m_l2Power(27) = CLng(134217728)
        m_l2Power(28) = CLng(268435456)
        m_l2Power(29) = CLng(536870912)
        m_l2Power(30) = CLng(1073741824)
        Dim x
        Dim k
        Dim AA
        Dim BB
        Dim CC
        Dim DD
        Dim a
        Dim b
        Dim c
        Dim d
        Dim S11
        S11 = 7
        Dim S12
        S12 = 12
        Dim S13
        S13 = 17
        Dim S14
        S14 = 22
        Dim S21
        S21 = 5
        Dim S22
        S22 = 9
        Dim S23
        S23 = 14
        Dim S24
        S24 = 20
        Dim S31
        S31 = 4
        Dim S32
        S32 = 11
        Dim S33
        S33 = 16
        Dim S34
        S34 = 23
        Dim S41
        S41 = 6
        Dim S42
        S42 = 10
        Dim S43
        S43 = 15
        Dim S44
        S44 = 21
        x = ConvertToWordArray(sMessage)
        a = &H67452301
        b = &HEFCDAB89
        c = &H98BADCFE
        d = &H10325476
        For k = 0 To UBound(x) Step 16
            AA = a
            BB = b
            CC = c
            DD = d
            md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
            md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
            md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
            md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
            md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
            md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
            md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
            md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
            md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
            md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
            md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
            md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
            md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
            md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
            md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
            md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
            md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
            md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
            md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
            md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
            md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
            md5_GG d, a, b, c, x(k + 10), S22, &H2441453
            md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
            md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
            md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
            md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
            md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
            md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
            md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
            md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
            md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
            md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
            md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
            md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
            md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
            md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
            md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
            md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
            md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
            md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
            md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
            md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
            md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
            md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
            md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
            md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
            md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
            md5_II a, b, c, d, x(k + 0), S41, &HF4292244
            md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
            md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
            md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
            md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
            md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
            md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
            md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
            md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
            md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
            md5_II c, d, a, b, x(k + 6), S43, &HA3014314
            md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
            md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
            md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
            md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
            md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
            a = AddUnsigned(a, AA)
            b = AddUnsigned(b, BB)
            c = AddUnsigned(c, CC)
            d = AddUnsigned(d, DD)
        Next
        'MD5=LCase(WordToHex(b) & WordToHex(c))
        MD5 = UCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
    End Function

End Class



'==========判断奇数==========

Function ifodd(t0)
    If Int(t0) Mod 2 = 0 Then
        ifodd = False
    Else
        ifodd = True
    End If
End Function

Sub infoback(t0)
    die "<script language=javascript>alert('"&t0&"');window.history.back();</script>"
End Sub

Sub infohref(t0, t1)
    die "<script language=javascript>alert('"&t0&"');window.location.href='"&t1&"'</script>"
End Sub

'==========报错信息==========

Sub iderror()
    If inull(rqid) Or Not IsNumeric(rqid) Then
        Call infoback("参数不能为空且必须为数字！")
    End If
End Sub

Sub nullback(t0, t1)
    If inull(t0) Then
        Call infoback(t1)
    End If
End Sub

Sub nullnumback(t0, t1)
    If inull(t0) Or Not IsNumeric(t0) Then
        Call infoback(t1)
    End If
End Sub

'==========根据ID获取任何表的任何字段==========

Function get_field(t0, t1, t2)
    Set rs_gf = ado_query("select * from "&t0&" where id = "&t1&" ")
    If rs_gf.EOF Then
        Call infoback ("提示："&t0&"表中ID为"&t1&"的数据不存在或已被删除！请修改该模板中get_field函数的相关参数！")
    Else
        get_field = rs_gf(t2)
    End If
    rs_gf.Close
    Set rs_gf = Nothing
End Function

'==========统计个数==========

Function get_count(t0)
    Set rs_gc = ado_query("select * from "&t0&"")
    get_count = rs_gc.RecordCount
    rs_gc.Close
    Set rs_gc = Nothing
End Function
%>
