<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
Call iderror()
sql = "update pf_info set pf_hits = pf_hits + 1 where id="&rqid&""
conn.Execute sql

'==========信息详情==========

Set rs = ado_query("select * from pf_info where id = "&rqid&"")
If rs.EOF Then
    Call infoback ("不存在此信息或者此信息已经被删除！")
End If
info_name = rs("pf_name")
info_author = rs("pf_author")
info_source = rs("pf_source")
info_ispic = rs("pf_ispic")
info_pic = rs("pf_pic")
info_parent = rs("pf_parent")
info_keywords = rs("pf_keywords")
info_description = rs("pf_description")
info_short_content = rs("pf_short_content")
info_content = rs("pf_content")
info_hits = rs("pf_hits")
info_date = rs("pf_date")
info_order = rs("pf_order")
info_modify_date = rs("pf_modify_date")
rs.Close
Set rs = Nothing

'==========所属分类相关信息==========
Set rs = ado_query("select * from pf_category where id = "&info_parent&"")
category_current_id = rs("id")
category_current_name = rs("pf_name")
category_current_short_name = rs("pf_short_name")
category_current_type = rs("pf_type")
category_current_info_type = rs("pf_info_type")
category_current_parent = rs("pf_parent")
category_current_sub = rs("pf_sub")
category_current_main = rs("pf_main")
category_current_keywords = rs("pf_keywords")
category_current_description = rs("pf_description")
category_current_content = rs("pf_content")
category_current_short_content = rs("pf_short_content")
category_current_pagecount = rs("pf_pagecount")
category_current_order = rs("pf_order")
category_current_date = rs("pf_date")
rs.Close
Set rs = Nothing


'==========当前分类的顶级相关信息==========

Set rs = ado_query("select * from pf_category where id = "&category_current_main&"")
category_main_name = rs("pf_name")
category_main_short_name = rs("pf_short_name")
category_main_type = rs("pf_type")
category_main_info_type = rs("pf_info_type")
category_main_parent = rs("pf_parent")
category_main_sub = rs("pf_sub")
category_main_keywords = rs("pf_keywords")
category_main_description = rs("pf_description")
category_main_content = rs("pf_content")
category_main_short_content = rs("pf_short_content")
category_main_pagecount = rs("pf_pagecount")
category_main_order = rs("pf_order")
category_main_date = rs("pf_date")
rs.Close
Set rs = Nothing

'==========分类当前位置==========

Function category_current_location()
    category_current_location = pf_category_current_location(info_parent)
End Function

'==========分类列表 - 无限级分类==========

Function category_list_infinite()
    category_list_infinite = pf_category_list_infinite(category_current_main, 0, "└", info_parent)
End Function

'==========分类列表 - 递推一级分类==========

Function category_list_recursion()
    category_list_recursion = pf_category_list_recursion(category_current_parent, category_parent_parent)
End Function

'==========分类名称 - 递推一级分类==========

If yn_sub(category_current_parent) Then
    category_recursion_name = category_current_name
Else
    category_recursion_name = category_parent_name
End If

'==========上一信息==========

Function info_previous(t0)
	Set rs = ado_query("select * from pf_info where pf_parent in ("&category_main_sub&") and pf_order > "&info_order&" order by pf_order asc")
    If rs.EOF Then
        info_previous = "暂无！"
    Else
        info_previous = "<a href=""info.asp?id="&rs("id")&""" title="""&rs("pf_name")&""">"&str_left(rs("pf_name"), t0, "...")&"</a>"
    End If
    rs.Close
    Set rs = Nothing
End Function

'==========下一信息==========

Function info_next(t0)
	Set rs = ado_query("select * from pf_info where pf_parent in ("&category_main_sub&") and pf_order < "&info_order&" order by pf_order desc")
    If rs.EOF Then
        info_next = "暂无！"
    Else
        info_next = "<a href=""info.asp?id="&rs("id")&""" title="""&rs("pf_name")&""">"&str_left(rs("pf_name"), t0, "...")&"</a>"
    End If
    rs.Close
    Set rs = Nothing
End Function

current_nav = "category.asp?id="&category_current_main&""
echo ob_get_contents(skin&category_current_info_type)
%>
