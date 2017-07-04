<%  
'========================================  
class EngineerSearch  
'老龙:laolong9999@sina.com  
':模拟XML获取http标记资源（用过之后就知道为什么XML有用：））  
'利用引擎搜索(显示引擎信息或其超连接网站上的信息或直接一个指定页面的相关信息，利用正则和xmlHttp,  
'程序的使用需要会构造正则)  
'---------------------------------------------------------------  
private oReg,oxmlHttp'一个正则，一个微软xmlhttp  
'---------------------------------------------------------------  
	
	public sub class_initialize()'对象建立触发  
		set oReg=new regExp  
		oReg.Global=true  
		oReg.IgnoreCase=true  
		set oXmlHttp=server.createobject("Microsoft.XmlHttp")  
	end sub  
	'---------------------------------------------------------------  
	
	public sub class_terminate()'对象销毁触发  
		set oReg=nothing'必须手动释放class内的自建对象，asp只自动释放由class定义的对象  
		set oXmlHttp=nothing  
		If typename(tempReg)<>"nothing" then'方法体内的对象释放资源  
			set tempReg=nothing  
		end if  
	end sub  
	'---------------------------------------------------------------  
		
	'引擎级搜索  
	public function engineer(url,EngineerReg)  
	'功能介绍：获得url的返回信息(通常用于引擎查找)，提取其中的EngineerReg的特定信息，返回matches集合到  
	'函数名。获得url查询结果,搜寻出用engineerReg正则定义的结果,生成一个matches集合,  
	'由于无法建立集合及操作集合个数（vbscript）,最好再自己遍历集合，也可以考虑二维数组  
		dim strConent  
		strContent=oXmlHttp.open("get",url,false)  
		on error resume next  
		oXmlHttp.send()  
		if err.number<>0 then  
			exit function  
		end if  
		
		'strContent=bytes2BSTR(oXmlHttp.responseBody)  
		'strContent=oXmlHttp.responseBody
		strContent=bytesToBSTR(oXmlHttp.responseBody,"GB2312")

		if isnull(EngineerReg) then  
			engineer=strContent 
		else  
			oReg.Pattern=EngineerReg  
			set engineer=oReg.Execute(strContent)   
		end if  
	end function  
	'---------------------------------------------------------------  


	public Function BytesToBstr(body,Cset)
		dim objstream
		set objstream = Server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = Cset
		BytesToBstr = objstream.ReadText
		objstream.Close
		set objstream = nothing
	End Function
	'---------------------------------------------------------------  



	'汉字编码,(网人)  
	public Function bytes2BSTR(vIn)   
		strReturn = ""   
		For i = 1 To LenB(vIn)   
			ThisCharCode = AscB(MidB(vIn,i,1))   
			If ThisCharCode < &H80 Then   
				strReturn = strReturn & Chr(ThisCharCode)   
			Else   
				NextCharCode = AscB(MidB(vIn,i+1,1))   
				strReturn = strReturn & Chr (CLng(ThisCharCode) * &H100 + CInt(NextCharCode))   
				i = i + 1   
			End If   
		Next   
		bytes2BSTR = strReturn   
	End Function  


	'---------------------------------------------------------------  
	public Function SearchReplace(strContent,ReplaceReg,ResultReg)  
	'替换，将strContent中的replaceReg描述的字符串用resultReg描述的替换，返回到searchReplace去  
	'将正则的replace封装了。  
		oReg.Pattern=ReplaceReg  
		SearchReplace=oReg.replace(strContent,ResultReg)  
	End Function
	'---------------------------------------------------------------  


	public Function AbsoluteURL(strContent,byval url)  
	'将strContent中的相对URL变成oXmlHttp中指定的url的绝对地址(http/https/ftp/mailto:)  
	'正则可以修改修改。  
		dim tempReg  
		set tempReg=new RegExp  
		tempReg.IgnoreCase=true  
		tempReg.Global=true  
		tempReg.Pattern="(^.*\/).*$"'含文件名的标准路径http://www.wrclub.net/default.aspx  
		Url=tempReg.replace(url,"$1")  
		tempReg.Pattern="((?:src|href).*?=[\'\u0022](?!ftp|http|https|mailto))"  
		AbsoluteURL=tempReg.replace(strContent,"$1"+Url)  
		set tempReg=nothing  
	end Function  
	'---------------------------------------------------------------  


end class  
'========================================  
%>
