<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
If Request.Form("submit") = "�ύ����" Then
    Call nullback(Request.Form("pf_content"), "�������ݲ���Ϊ�գ�")
    Call nullback(Request.Form("pf_verifycode"), "��֤�벻��Ϊ�գ�")
    If CStr(Session("CheckCode"))<>CStr(Request.Form("pf_verifycode")) Then
        infoback "��֤�����"
    End If
    Set rs = ado_query_modify("select * from pf_guestbook")
    rs.addnew
    If request.Form("pf_name") = "" Then
        rs("pf_name") = "�����ο�"
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
    Call infohref ("���Գɹ�����ȴ����ǵĻظ���", "guestbook.asp")
End If
current_nav = "guestbook.asp"
echo ob_get_contents(""&skin&"guestbook.asp")
%>
