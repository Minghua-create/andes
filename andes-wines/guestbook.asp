<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
If Request.Form("submit") = "提交留言" Then
    Call nullback(Request.Form("pf_content"), "留言内容不能为空！")
    Call nullback(Request.Form("pf_verifycode"), "验证码不能为空！")
    If CStr(Session("CheckCode"))<>CStr(Request.Form("pf_verifycode")) Then
        infoback "验证码错误！"
    End If
    Set rs = ado_query_modify("select * from pf_guestbook")
    rs.addnew
    If request.Form("pf_name") = "" Then
        rs("pf_name") = "匿名游客"
    Else
        rs("pf_name") = str_safe(request.Form("pf_name"))
    End If
    rs("pf_contact") = str_safe(request.Form("pf_contact"))
    rs("pf_content") = str_safe(request.Form("pf_content"))
    rs("pf_date") = Now
    rs("pf_audit") = pf_guestbook_audit
    rs.update
    rs.Close
    Set rs = Nothing
    Call infohref ("留言成功！请等待我们的回复！", "guestbook.asp")
End If
current_nav = "guestbook.asp"
echo ob_get_contents(""&skin&"guestbook.asp")
%>
