<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
call iderror()
If request.Form("submit") = "提交简历" Then
    Call nullback(Request.Form("pf_name"), "姓名不能为空！")
    Call nullback(Request.Form("pf_verifycode"), "验证码不能为空！")
    If CStr(Session("CheckCode"))<>CStr(Request.Form("pf_verifycode")) Then
        infoback "验证码错误！"
    End If
	Set rs = ado_query_modify("select * from pf_resume")
    rs.addnew
    rs("pf_name") = str_safe(Request.Form("pf_name"))
    rs("pf_position") = get_field("pf_recruitment", rqid, "pf_name")
    rs("pf_gender") = str_safe(Request.Form("pf_gender"))
    rs("pf_age") = str_safe(Request.Form("pf_age"))
    rs("pf_degree") = str_safe(Request.Form("pf_degree"))
    rs("pf_mobile") = str_safe(Request.Form("pf_mobile"))
    rs("pf_email") = str_safe(Request.Form("pf_email"))
    rs("pf_address") = str_safe(Request.Form("pf_address"))
    rs("pf_content") = str_safe(Request.Form("pf_content"))
    rs("pf_date") = Now
    rs.update
    rs.Close
    Set rs = Nothing
    call infohref ("简历提交成功！请等待我们和您联系！", "recruitment.asp")
End If
echo ob_get_contents(""&skin&"resume.asp")
%>
