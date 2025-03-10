<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="dkcms.asp"-->
<%
Call iderror()

'==========当前分类相关信息==========

Set rs = ado_query("select * from pf_category where id = "&rqid&"")
If rs.EOF Then
    Call infoback ("无效参数！")
End If
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

'==========当前分类的上级相关信息==========

If category_current_parent <> 0 Then
    sql = "select * from pf_category where id = "&category_current_parent&""
Else
    sql = "select * from pf_category where id = "&rqid&""
End If
Set rs = ado_query(sql)
category_parent_name = rs("pf_name")
category_parent_short_name = rs("pf_short_name")
category_parent_type = rs("pf_type")
category_parent_info_type = rs("pf_info_type")
category_parent_parent = rs("pf_parent")
category_parent_sub = rs("pf_sub")
category_parent_keywords = rs("pf_keywords")
category_parent_description = rs("pf_description")
category_parent_content = rs("pf_content")
category_parent_short_content = rs("pf_short_content")
category_parent_pagecount = rs("pf_pagecount")
category_parent_order = rs("pf_order")
category_parent_date = rs("pf_date")
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
    category_current_location = pf_category_current_location(rqid)
End Function

'==========分类列表 - 无限级分类==========

Function category_list_infinite()
    category_list_infinite = pf_category_list_infinite(category_current_main, 0, "└", rqid)
End Function

'==========分类列表 - 递推一级分类==========

Function category_list_recursion()
    category_list_recursion = pf_category_list_recursion(rqid,category_current_parent)
End Function

'==========分类名称 - 递推一级分类==========

If yn_sub(rqid) Then
    category_recursion_name = category_current_name
Else
    category_recursion_name = category_parent_name
End If
current_nav = "category.asp?id="&category_current_main&""
echo ob_get_contents(skin&category_current_type)
%>
