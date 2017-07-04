<%
'On Error Resume Next
   dim connstr,datapath,conn
   datapath="/DataBase/xmfishBBS_analysis.mdb"
   connstr="Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Server.mappath(datapath)
   Set conn=Server.CreateObject("ADODB.Connection")
   conn.open connstr
%>