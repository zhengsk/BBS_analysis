<!-- #include file="Include/conn.asp" -->

<!-- #include file="Include/XMLHttp.asp" -->

<%

Function article_Add()
	
	SQL="Select * from article where article_ID=" & article_ID
	rs.open SQL,conn,1,3


	If Rs.Eof and Rs.Bof then
		rs.AddNew
		rs("article_ID")=article_ID
		response.write "<span style='color:red'>添加成功！</span>"  
		response.write "<br><br><br><br>"
	Else
		response.write "<span style='color:red'>修改成功！</span>"  
		response.write "<br><br><br><br>"  
	End If
	
	rs("article_title")=article_title
	rs("article_authorID")=article_authorID
	rs("article_authorName")=article_authorName
	rs("article_time")=article_time
	rs("article_replyTimes")=article_replyTimes
	rs("article_readTimes")=article_readTimes
	rs.Update
	
	rs.Close

End Function


%>


<%'例子  
Response.CharSet = "GB2312" 

dim pageNum
pageNum = Request("pageNum")


dim GetUrl
GetUrl = "http://bbs.xmfish.com/thread-htm-fid-160-page-"& pageNum &".html"

response.write(GetUrl)
response.write("<br/><br/><br/>")

dim RegStr
RegStr = "name=""readlinkt"" target=""_blank"" id=""a_ajax_(\d+)"" class=""subject_t f14"">([\s\S]+?)</a>[\s\S]+?<td class=""author""><a href=""u\.php\?uid=(\d+)"" rel=""external nofollow"">([\s\S]+?)</a><p>([\s\S]+?)</p></td>[\s\S]+?<td class=""num""><em>(\d+)</em>/(\d+)</td>"


dim mySearch   
set mySearch=new EngineerSearch  
'URL一定是包含文件扩展名的完整地址,结果是集合，集合中的每个项目是数组,应该这样引用子查询:myMatches(0).subMatches(0)  
set myMatches=mySearch.engineer(GetUrl,RegStr)  

if myMatches.count=0 Then  
	response.write "没有你正则的字符串"
end if


set rs=server.createobject("adodb.recordset")
dim article_ID,article_title,article_authorID,article_authorName,article_time,article_replyTimes,article_readTimes

if myMatches.count>0 then  
	response.write myMatches.count&"<br><br>"  
	for each key in myMatches  

		article_ID = key.SubMatches.Item(0)
		article_title = key.SubMatches.Item(1)
		article_authorID = key.SubMatches.Item(2)
		article_authorName = key.SubMatches.Item(3)
		article_time = key.SubMatches.Item(4)
		article_replyTimes = key.SubMatches.Item(5)
		article_readTimes = key.SubMatches.Item(6)

		response.write "article_ID:---"&article_ID&"<br>"  
		response.write "article_title:---"&article_title&"<br>"  
		response.write "article_authorID:---"&article_authorID&"<br>"  
		response.write "article_authorName:---"&article_authorName&"<br>"  
		response.write "article_time:---"&article_time&"<br>"  
		response.write "article_replyTimes:---"&article_replyTimes&"<br>"  
		response.write "article_readTimes:---"&article_readTimes&"<br>"  
		
		article_Add()

	next  
end if  

Set rs=nothing


mySearch.class_terminate()



%>


<script type="text/javascript">
	var pageNum = <%=pageNum%>

	if(pageNum >25){
		alert("到25页了，结束了！")
	}else{
		setTimeout(function(){
			window.location = "?pageNum=" + (pageNum+1)
		},5000)
	}

</script> 