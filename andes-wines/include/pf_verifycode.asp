<%
Option Explicit
Response.buffer = True
NumCode
Function NumCode()
    Response.Expires = -1
    Response.AddHeader "Pragma", "no-cache"
    Response.AddHeader "cache-ctrol", "no-cache"
    On Error Resume Next
    Dim zNum, i, j
    Dim Ados, Ados1
    Randomize Timer
    zNum = CInt(8999 * Rnd + 1000)
    Session("CheckCode") = zNum
    Dim zimg(4), NStr
    NStr = CStr(zNum)
    For i = 0 To 3
        zimg(i) = CInt(Mid(NStr, i + 1, 1))
    Next
    Dim Pos
    Set Ados = Server.CreateObject("Adodb.Stream")
    Ados.Mode = 3
    Ados.Type = 1
    Ados.Open
    Set Ados1 = Server.CreateObject("Adodb.Stream")
    Ados1.Mode = 3
    Ados1.Type = 1
    Ados1.Open
    Ados.LoadFromFile(Server.mappath("pf_body.Fix"))
    Ados1.Write Ados.Read(1280)
    For i = 0 To 3
        Ados.Position = (9 - zimg(i)) * 320
        Ados1.Position = i * 320
        Ados1.Write ados.Read(320)
    Next
    Ados.LoadFromFile(Server.mappath("pf_head.fix"))
    Pos = lenb(Ados.Read())
    Ados.Position = Pos
    For i = 0 To 9 Step 1
        For j = 0 To 3
            Ados1.Position = i * 32 + j * 320
            Ados.Position = Pos + 30 * j + i * 120
            Ados.Write ados1.Read(30)
        Next
    Next
    Response.ContentType = "image/BMP"
    Ados.Position = 0
    Response.BinaryWrite Ados.Read()
    Ados.Close
    Set Ados = Nothing
    Ados1.Close
    Set Ados1 = Nothing
    If Err Then Session("CheckCode") = 9999
End Function
%>
