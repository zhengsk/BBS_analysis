<!-- #include file="Include/conn.asp" -->
<!-- #include file="Include/XMLHttp.asp" -->

<%

Function reply_Add()
	
	SQL="Select * from reply where reply_ID=" & reply_ID
	rs.open SQL,conn,1,3


	If Rs.Eof and Rs.Bof then
		rs.AddNew
		rs("reply_ID")=reply_ID
		response.write "<span style='color:red'>��ӳɹ���</span>"  
		response.write "<br><br><br><br>"
	Else
		response.write "<span style='color:red'>�޸ĳɹ���</span>"  
		response.write "<br><br><br><br>"  
	End If

	
	rs("reply_articleID")=reply_articleID
	rs("reply_userID")=reply_userID
	rs("reply_userName")=reply_userName
	rs("reply_time")=reply_time
	rs("reply_content")=reply_content

	rs.Update
	
	rs.Close

End Function

%>



<%'����  
Response.CharSet = "GB2312"   


dim article_rs, article_sql, article_ID,article_replyTimes
article_sql = "select article_ID,article_replyTimes from article"

	set article_rs=server.CreateObject("ADODB.RECORDSET")
		article_rs.open article_sql,conn,1,1

		if not article_rs.eof then
			article_rs.pageSize = 1 '����ÿҳ��ʾ�ļ�¼��
			allPages = article_rs.pageCount'����һ���ֶܷ���ҳ
			articlePage = Request.QueryString("articlePage")'ͨ����������ݵ�ҳ��
			if isEmpty(articlePage) or Cint(articlePage) < 1 then
				articlePage = 1
			elseif Cint(articlePage) > allPages then
				response.write("���ӵ����һ���ˣ������ˣ�")
				response.end()
			end if
			article_rs.AbsolutePage = Cint(articlePage)
		else
			response.write("���ݿ���û�����ӣ�")
			response.end()
		End If
		
		article_ID = article_rs("article_ID")
		article_replyTimes = article_rs("article_replyTimes")

	article_rs.Close
	set article_rs = nothing










dim mySearch   
set mySearch=new EngineerSearch


dim pageNum
pageNum = Request("pageNum")
if isEmpty(pageNum) or Cint(pageNum) < 1 then pageNum = 1


dim GetUrl
GetUrl = "http://bbs.xmfish.com/read-htm-tid-"& article_ID &"-page-"& pageNum &".html"

response.write(GetUrl)
response.write("<br/><br/><br/>")

dim RegStr
RegStr = "<a name=(\d+)></a>[\s\S]+?<a href=""u\.php\?uid=(\d+)"">([\s\S]*?)<\/a>[\s\S]+?<span title=""([\s\S]*?)"">������:[\s\S]+?<div class=""read_h1"" style=""margin-bottom:10px;"" id=""subject_[\s\S]+?"">([\s\S]+?)</div>[\s\S]+?<!--content_read-->([\s\S]+?)<!--content_read-->"


'URLһ���ǰ����ļ���չ����������ַ,����Ǽ��ϣ������е�ÿ����Ŀ������,Ӧ�����������Ӳ�ѯ:myMatches(0).subMatches(0)  
set myMatches=mySearch.engineer(GetUrl,RegStr)  



if myMatches.count=0 Then  
	response.write "û����������ַ���"
end if

set rs=server.createobject("adodb.recordset")
dim reply_ID,reply_articleID,reply_userID,reply_userName,reply_time,reply_content


if myMatches.count>0 then  
	response.write myMatches.count&"<br><br>"  
	for each key in myMatches  

		reply_ID = key.SubMatches.Item(0)
		reply_userID = key.SubMatches.Item(1)
		reply_userName = key.SubMatches.Item(2)
		reply_time = key.SubMatches.Item(3)
		reply_content = key.SubMatches.Item(4) & key.SubMatches.Item(5)
		reply_articleID = article_ID


		response.write "reply_articleID:---"&article_ID&"<br>"  
		response.write "reply_ID:---"&reply_ID&"<br>"  
		response.write "reply_userID:---"&reply_userID&"<br>"  
		response.write "reply_userName:---"&reply_userName&"<br>"  
		response.write "reply_time:---"&reply_time&"<br>"  
		response.write "reply_content:---"&reply_content&"<br>"  
	
		
		reply_Add()
	next  
end if  

mySearch.class_terminate()

response.write("-------------")
response.write(int(pageNum))
response.write("-------------")
response.write(int((article_replyTimes+(30-1))/30))
response.write("-------------")
response.write(int(pageNum) = int((article_replyTimes+(30-1))/30))
response.write("-------------")
response.write(int(pageNum))
response.write("-------------")

%> 



<%if int(pageNum) = int((article_replyTimes+(30-1))/30) or article_replyTimes = 0 then %> 

	<script type="text/javascript">
		var articlePage = <%=articlePage%>

		//if(articlePage >= 1000){alert("��1000ҳ�ˣ�")}

		setTimeout(function(){
			window.location = "?articlePage="+ (articlePage+1) +"&pageNum=1"
		},1000)
	</script> 

<%else%> 

	<script type="text/javascript">
		var pageNum = <%=pageNum%>
		var articlePage = <%=articlePage%>


		setTimeout(function(){
			window.location = "?articlePage="+ articlePage +"&pageNum=" + (pageNum+1)
		},1000)
	</script> 

<%end if%> 





