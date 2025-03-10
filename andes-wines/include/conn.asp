<%
data_path = "/data/#andesverygoodwine2011.mdb"
connstr = "DBQ=" + Server.MapPath(""&data_path&"") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn = server.CreateObject("ADODB.CONNECTION")
conn.Open connstr
'On Error Resume Next
%>
