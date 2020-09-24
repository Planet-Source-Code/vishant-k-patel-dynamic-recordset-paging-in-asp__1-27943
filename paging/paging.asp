<html>
<head>
<title>Dynamic Paging by Vishant K. Patel [ vishantpatel@hotmail.com ]</title>
<style>
.bodytext{font-family:verdana;tahoma;font-size:8pt}
A{text-decoration:none;}
</style>
</head>

<body>
<%
	Dim objConn,objRs,TableName,ConnString,intPageSize

'Just Modify following three parameters ONLY, that's it ! :-)
'********************************************************************************************

	'Give the path to your db file
	ConnString  = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("mydb.mdb")
	'Name of Table
	TableName   = "ques"  
	' No of Records per page	
	intPageSize = 10	  

'********************************************************************************************
	
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open ConnString
	Set objRs   = Server.CreateObject("ADODB.Recordset")
	objRs.CursorLocation = 3 'adUseClient
	objRs.Open TableName,objConn
	
	Dim intPageCount
	objRs.PageSize = intPageSize
	intPageCount   = objRs.PageCount
	intPageNo      = CInt(Request.QueryString("PageNo"))
	
	If IsNull(intPageNo) or intPageNo=0 Then
		objRs.AbsolutePage = 1
	ElseIf intPageNo <= intPageCount Then
		objRs.AbsolutePage = intPageNo
	End If
	
	Response.Write "<table border=0 class=bodytext><tr><td width=380 align=left>"
    Response.write "Total " & objRs.recordcount & " Records Found :&nbsp;&nbsp;"
    Response.write  "Page " & objRs.absolutepage & " of " & objRs.pagecount & "</td><td width=380 align=right>"
    
    If objRs.absolutepage = 1 Then 
      Response.Write "[<span style='color:silver'>Previous Page</span>"
    Else  
      Response.Write "[<b>&#171;&nbsp;</b><a href='" & Request.ServerVariables("SCRIPT_NAME") & "?PageNo=" & objRs.absolutepage - 1 & "'>Previous Page</a>"
    End If
    
    Response.Write "] ["
    
    If objRs.absolutepage < objRs.pagecount Then 
      Response.Write "<a href='" & Request.ServerVariables("SCRIPT_NAME") & "?PageNo=" & objRs.absolutepage + 1 & "'>Next Page</a><b>&nbsp;&#187;</b>]"
    Else  
      Response.Write "<span style='color:silver'>Next Page </span>]"
    End If
    
    Response.Write "</td></tr></table></center><hr>"
    Response.Write "<table border=0 class=bodytext cellspacing=1 cellpadding=2><tr>"
    
	Dim i,j,VishValue
	For i=0 to objRs.Fields.Count-1
		Response.Write "<td bgcolor=#ffcc99 align=center><b>" & objRs(i).Name & "</b></td>"	
	Next
	Response.Write "</tr>"
	
	For i=1 to objRs.PageSize
		Response.Write "<tr bgcolor=#eeeeee>"
		For j=0 to objRs.Fields.Count-1
			VishValue = objRs(j)
			If IsNull(VishValue) Then VishValue = "Not Entered"
			Response.Write "<td>" & VishValue & "</td>"	
		Next
		Response.Write "</tr>"
		objRs.MoveNext
	If objRs.EOF Then Exit For
	Next

%>
</table>
</body>
</html>
