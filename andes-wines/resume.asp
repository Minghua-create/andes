<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
call iderror()
If request.Form("submit") = "�ύ����" Then
    Call nullback(Request.Form("pf_name"), "��������Ϊ�գ�")
    Call nullback(Request.Form("pf_verifycode"), "��֤�벻��Ϊ�գ�")
    If CStr(Session("CheckCode"))<>CStr(Request.Form("pf_verifycode")) Then
        infoback "��֤�����"
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
    call infohref ("�����ύ�ɹ�����ȴ����Ǻ�����ϵ��", "recruitment.asp")
End If
echo ob_get_contents(""&skin&"resume.asp")
%>
