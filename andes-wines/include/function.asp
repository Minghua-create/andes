<%

'==========��վ������Ϣ==========
Set rs = ado_query("select * from pf_site")
skin = "skins/"&rs("pf_skin")&"/"
site_name = rs("pf_name")'&p_name
site_address = rs("pf_address")
site_url = rs("pf_url")
site_tel = rs("pf_tel")
site_mobile = rs("pf_mobile")
site_fax = rs("pf_fax")
site_email = rs("pf_email")
site_contact = rs("pf_contact")
site_qq = rs("pf_qq")
site_ww = rs("pf_ww")
site_msn = rs("pf_msn")
site_skype = rs("pf_skype")
site_icp = rs("pf_icp")
site_keywords = rs("pf_keywords")
site_description = rs("pf_description")
site_copyright = rs("pf_copyright")
site_code = rs("pf_code")
pf_customer_support = rs("pf_customer_support")
pf_member_audit = rs("pf_member_audit")
pf_guestbook_audit = rs("pf_guestbook_audit")
pf_comment_audit = rs("pf_comment_audit")
rs.Close
Set rs = Nothing

'==========��Ա�ȼ�==========

Function member_level()
    If session("pf_member_check") = "" Then
        member_level = 1
    Else
        Set rs = ado_query("select * from pf_member where pf_name = '"&session("pf_member_check")&"'")
        member_level = rs("pf_level")
    End If
End Function

'==========�жϵ�ǰ�����Ƿ�������==========

Function yn_sub(t0)
    yn_sub = True
    Set rs_ys = ado_query("select * from pf_category where pf_parent = "&t0&"")
    If rs_ys.EOF Then
        yn_sub = False
    End If
    rs_ys.Close
    Set rs_ys = Nothing
End Function

'==========�õ�==========

Function banner()
    Set rs = ado_query("select * from pf_banner order by pf_order desc")
    Do While Not rs.EOF
        banner = banner&"files+='|"&rs("pf_pic")&"';links+='|"&rs("pf_link")&"';texts+='|"&rs("pf_name")&"';"
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
End Function

'==========����ID��ȡ�������κ��ֶ�==========

Function get_category(t0, t1)
    Set rs_gc = ado_query("select * from pf_category where id = "&t0&" ")
    If rs_gc.EOF Then
        Call infoback ("��ʾ��������IDΪ"&t0&"�����ݲ����ڻ��ѱ�ɾ�������޸ĸ�ģ����get_category��������ز�����")
    Else
        get_category = rs_gc(t1)
    End If
    rs_gc.Close
    Set rs_gc = Nothing
End Function

'==========���µ���������==========

Function get_category_name(t0)
    get_category_name = "<a href=""category.asp?id="&t0&""" target=""_blank"">"&get_category(t0, "pf_name")&"</a>"
End Function


'==========�����Ƿ�Ϊ����==========

Function if_bold(t0)
    if_bold = IIf(t0 = 0, " style=""font-weight:bold;""", "")
End Function

'==========�����Ƿ�Ϊ�ȵ�==========

Function if_hot(t0)
    if_hot = IIf(t0 = 0, "��", "")
End Function

'==========�����Ƿ�Ϊ�Ƽ�==========

Function if_rec(t0)
    if_rec = IIf(t0 = 0, "��", "")
End Function

'==========�����Ƿ�ΪͼƬ==========

Function if_pic(t0)
    if_pic = IIf(t0 = 0, "ͼ", "")
End Function


'==========����ID��ȡ���==========

Function ad(t0)
    Set rs_a = ado_query("select * from pf_advertisement where id = "&t0&" ")
    If rs_a.EOF Then
        ad = "IDΪ"&t0&"�Ĺ��λ�����ڻ����ѱ�ɾ����"
    Else
        ad = rs_a("pf_content")
    End If
    rs_a.Close
    Set rs_a = Nothing
End Function

'==========���൱ǰλ��==========

Function pf_category_current_location(t0)
    Set rs_ccl = ado_query("select * from pf_category where id = "&t0&"")
    If rs_ccl("id") = Int(category_current_id) Then
        current_a = " class=""current_a"""
    Else
        current_a = ""
    End If
    pf_category_current_location = "<a"&current_a&" href=""category.asp?id="&rs_ccl("id")&""">"&rs_ccl("pf_name")&"</a>"
    If rs_ccl("pf_parent") <> 0 Then
        pf_category_current_location = pf_category_current_location(rs_ccl("pf_parent"))&" >> "&pf_category_current_location
    End If
    rs_ccl.Close
    Set rs_ccl = Nothing
End Function

'==========�����б� - ���޼�����==========

Function pf_category_list_infinite(t0, t1, t2, t3)
    Set rs_cl = ado_query("select * from pf_category where pf_parent = "&t0&" order by pf_order asc")
    For i = 1 To t1
        separator = separator&"��"
    Next
    Do While Not rs_cl.EOF
        If rs_cl("id") = Int(t3) Then
            current_category = " class=""current_category_infinite"""
        Else
            current_category = ""
        End If
        pf_category_list_infinite = pf_category_list_infinite&"<li"&current_category&"><a href=""category.asp?id="&rs_cl("id")&""">"&separator&t2&rs_cl("pf_name")&"</a></li>"&vbCrLf&pf_category_list_infinite(rs_cl("id"), i, t2, t3)
        rs_cl.movenext
    Loop
    rs_cl.Close
    Set rs_cl = Nothing
End Function

'==========�����б� - ����һ������==========

Function pf_category_list_recursion(t0, t1)
    If yn_sub(t0) Then
        sql_cl = "select * from pf_category where pf_parent = "&t0&" order by pf_order asc"
    Else
        sql_cl = "select * from pf_category where pf_parent = "&t1&" order by pf_order asc"
    End If
    Set rs_cl = ado_query(sql_cl)
    Do While Not rs_cl.EOF
        If rs_cl("id") = Int(t0) Then
            current_category = " class=""current_category_recursion"""
        Else
            current_category = ""
        End If
        pf_category_list_recursion = pf_category_list_recursion&"<li"&current_category&"><a href=""category.asp?id="&rs_cl("id")&""">"&rs_cl("pf_name")&"</a></li>"
        rs_cl.movenext
    Loop
    rs_cl.Close
    Set rs_cl = Nothing
End Function

'==========�����б� - ���޼����� - ���չ��==========

Function category_list_tree(t0)
    Set rs_clt = ado_query("select * from pf_category where pf_parent = "&t0&" order by pf_order asc")
    Do While Not rs_clt.EOF
        If yn_sub(rs_clt("id")) Then
            if_close = " class=""Closed"""
        Else
            if_close = " class=""Child"""
        End If
        category_list_tree_a = category_list_tree_a&"<li"&if_close&"><a href=""category.asp?id="&rs_clt("id")&""">"&rs_clt("pf_name")&"</a>"&category_list_tree(rs_clt("id"))&"</li>"
        rs_clt.movenext
    Loop
    If category_list_tree_a = "" Then
        category_list_tree = category_list_tree_a
    Else
        category_list_tree = "<ul>"&category_list_tree_a&"</ul>"
    End If
    rs_clt.Close
    Set rs_clt = Nothing
End Function
%>
