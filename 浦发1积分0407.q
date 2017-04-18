[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=27fec89a-a1fd-44fb-8b6c-c5a943fb314a
Description=浦发1积分成品0407每次修改goodsid
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]
修改手机时间到抢购时间后，点击开始抢兑，把链接保存为saz，解压，将1_c,2_c等文件分别拷贝到远程vps，一台对应一个文件，将文件更名为1.txt,识别插件用948kb的

[Script]
Function getQueryStr(Src, Str)
	Dim matches, match, submatches
	set regEx = New RegExp
	regEx.[Global] = TRUE
	regEx.IgnoreCase = FALSE
	regEx.pattern = "(\&|\?)" & Str & "=([^\&#]*)(\&|$|#)"
	Set matches = regEx.execute(Src)
	//For Each Match In matches
	//TracePrint Match.SubMatches(0)
	//TracePrint Match.SubMatches(1)
	//TracePrint Match.SubMatches(2)
	//TracePrint Match.SubMatches(3)
	//next
	Set match = matches.Item(0)
	Set submatches = match.SubMatches
	getQueryStr = submatches(1)
End Function

Function UpdateCookieSecKill(responseText)
	JSESSIONID_old = GetStrAB(cookieSecKill, "JSESSIONID=", ";")
	JSESSIONID_new = GetStrAB(responseText, "JSESSIONID=", ";")
	cookieSecKill = Replace(cookieSecKill, JSESSIONID_old, JSESSIONID_new)
End Function

Function UpdateCookieCaptcha(responseText)
	JSESSIONID_old = GetStrAB(cookieCaptcha, "JSESSIONID=", ";")
	JSESSIONID_new = GetStrAB(responseText, "JSESSIONID=", ";")
	cookieCaptcha = Replace(cookieCaptcha, JSESSIONID_old, JSESSIONID_new)
End Function

Rem subSecKill
Text1 = Plugin.File.ReadFileEx("C:\raw\0413.txt")
href = GetStrAB(text1, "GET ", " HTTP/1.1")
cookieSecKill = GetStrAB(text1, "Cookie: ", "X-Requested-With")
cookieSecKill = Replace(cookieSecKill, "\r", "")
cookieSecKill = Replace(cookieSecKill, "\n", "")
refererSecKill = href
paramAppId = getQueryStr(href, "appId")
paramSectionId = getQueryStr(href, "sectionId")
paramGoodsId = getQueryStr(href, "goodsId")
paramChannelId = getQueryStr(href, "channelId")
cookieAuthTicket = GetStrAB(text1, "UserAuth=", ";")
paramUserId = getQueryStr(href, "userId")
paramSign = getQueryStr(href, "sign")
paramUrl = "http://campaign.e-pointchina.com.cn/campaign/cloudDataService.do?ttd="
timestampBase = timestamp()

//findSectionGoodsDetailInfo
body = "channelId=" &paramChannelId& "&appId=" &paramAppId& "&authTicket=" &cookieAuthTicket& "&userId=" &paramUserId& "&sign=" &paramSign& "&sectionId=" &paramSectionId& "&goodsId=" &paramGoodsId& "&serviceType=com.ebuy.o2o.campaign.service.CampaignService&serviceMethod=findSectionGoodsDetailInfo"
Do
	Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
	With http
		.Setproxy 2,"127.0.0.1:8888",0
		.open "POST", paramUrl & timestamp(), False
		.setrequestheader "Host", "campaign.e-pointchina.com.cn" 
		.setrequestheader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
		.setrequestheader "Connection", "keep-alive"
		.setrequestheader "Content-Length", len(body)
		.setrequestheader "Referer",refererSecKill
		.setrequestheader "Accept-Language", "zh-CN,en-US;q=0.8"
		.setrequestheader "Accept-Encoding", "gzip,deflate"
		.setrequestheader "Origin","http://campaign.e-pointchina.com.cn"
		.setrequestheader "Cookie", cookieSecKill
		.send body
	End with

	If isEmpty(http.responsetext) Then 
		TracePrint "findSectionGoodsDetailInfo 失败，重新发送"
	Else 
		response = http.responsetext
		errorCode = GetStrAB(response, "<errorCode>", "</errorCode>")
		UpdateCookieSecKill(response)
		If errorCode <> "0" Then
			errorMsg = GetStrAB(response, "<errorMsg>", "</errorMsg>")
			TracePrint "findSectionGoodsDetailInfo (" &errorMsg& ")，重新发送"
		Else
			Exit Do
		End If
	End If
Loop


// findSeckillResultBySectionId
body = "channelId=" &paramChannelId& "&appId=" &paramAppId& "&authTicket=" &cookieAuthTicket& "&userId=" &paramUserId& "&sign=" &paramSign& "&sectionId=" &paramSectionId& "&serviceType=com.ebuy.o2o.campaign.service.SeckillService&serviceMethod=findSeckillResultBySectionId"
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
With http
	.Setproxy 2,"127.0.0.1:8888",0
	.open "POST", paramUrl & timestamp(), False
	.setrequestheader "Host", "campaign.e-pointchina.com.cn" 
	.setrequestheader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
	.setrequestheader "Connection", "keep-alive"
	.setrequestheader "Content-Length", len(body)
	.setrequestheader "Referer",refererSecKill
	.setrequestheader "Accept-Language", "zh-CN,en-US;q=0.8"
	.setrequestheader "Accept-Encoding", "gzip,deflate"
	.setrequestheader "Origin","http://campaign.e-pointchina.com.cn"
	.setrequestheader "Cookie", cookieSecKill
	.send body
End with

If Not isEmpty(http.responsetext) Then 
	response = http.responsetext
	errorCode = GetStrAB(response, "<errorCode>", "</errorCode>")
	If errorCode <> "0" Then
		errorMsg = GetStrAB(response, "<errorMsg>", "</errorMsg>")
		TracePrint "findSeckillResultBySectionId (" &errorMsg& ")"
	End If
End If

// prepareSeckill
body = "channelId=" &paramChannelId& "&appId=" &paramAppId& "&authTicket=" &cookieAuthTicket& "&userId=" &paramUserId& "&sign=" &paramSign& "&sectionId=" &paramSectionId& "&goodsId=" &paramGoodsId& "&serviceType=com.ebuy.o2o.campaign.service.SeckillService&serviceMethod=prepareSeckill"
Do
	Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
	With http
		.Setproxy 2,"127.0.0.1:8888",0
		.open "POST", paramUrl & timestamp(), False
		.setrequestheader "Host", "campaign.e-pointchina.com.cn" 
		.setrequestheader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
		.setrequestheader "Connection", "keep-alive"
		.setrequestheader "Content-Length", len(body)
		.setrequestheader "Referer",refererSecKill
		.setrequestheader "Accept-Language", "zh-CN,en-US;q=0.8"
		.setrequestheader "Accept-Encoding", "gzip,deflate"
		.setrequestheader "Origin","http://campaign.e-pointchina.com.cn"
		.setrequestheader "Cookie", cookieSecKill
		.send body
	End with

	If isEmpty(http.responsetext) Then 
		TracePrint "prepareSeckill 失败，重新发送"
	Else 
		response = http.responsetext
		errorCode = GetStrAB(response, "<errorCode>", "</errorCode>")
		If errorCode <> "0" Then
			errorMsg = GetStrAB(response, "<errorMsg>", "</errorMsg>")
			TracePrint "prepareSeckill (" &errorMsg& ")，重新发送"
		Else
			app = paramAppId
			userId = paramUserId
			appTimestamp = GetStrAB(response, "<createTime>", "</createTime>")
			appToken = GetStrAB(response, "<accessToken>", "</accessToken>")
			captchaType = "1"
			seckillInterface = GetStrAB(response, "<seckillInterface>", "</seckillInterface>")
			counterReturn = 0
			counterCallback = 1
			Exit Do
		End If
	End If
Loop


// captcha
paramUrlCaptcha = "http://captcha.e-pointchina.com/captcha/cloudDataService.do?"
cookieCaptcha = "JSESSIONID=84F8A3E84EBD9FB6F83B32A33444FB38"

Rem subcaptcha

// newcaptcha
counterReturn = counterReturn + 1
counterCallback = counterCallback + 1
timestampBase = timestampBase + 1
jsonpReturn = "salama_ws_jsonp_val_" & counterReturn
body = "serviceType=captcha.ws.service.CaptchaService&serviceMethod=newCaptcha&app=" &app& "&userId=" &userId& "&captchaType=" &captchaType& "&appTimestamp=" &appTimestamp& "&appToken=" &appToken& "&timestamp=" &timestamp()& "&jsonpReturn=" &jsonpReturn& "&responseType=xml.jsonp&jsoncallback=jsonp_callback" &counterCallback& "&_=" &timestampBase
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Setproxy 2,"127.0.0.1:8888",0
http.open "Get", paramUrlCaptcha & body, false
http.setrequestheader "Host", "captcha.e-pointchina.com" 
http.setrequestheader "Connection", "keep-alive"
http.setrequestheader "Accept","*/*"
http.setrequestheader "User-Agent", "Mozilla/5.0 (Linux; Android 4.4.4; M463C Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Mobile Safari/537.36 MicroMessenger/6.3.18.800 NetType/WIFI Language/zh_CN"
http.setrequestheader "Accept-Encoding", "gzip, deflate"
http.setrequestheader "Referer", refererSecKill
http.setrequestheader "Accept-Language","zh-CN,zh;q=0.8"
http.setrequestheader "Cookie",cookieCaptcha
http.send
If IsEmpty(http.responseText) Then
	TracePrint "newcaptcha 失败，重新发送"
	Goto subcaptcha
End If

response = http.responsetext
UpdateCookieCaptcha(response)
responseNewCaptcha = GetStrAB(response, jsonpReturn & ' = "', '"')
responseNewCaptcha = decodeURI(responseNewCaptcha)
result = GetStrAB(responseNewCaptcha, "<result>", "</result>")
If result <> "success" Then
	TracePrint "newcaptcha result(" &result& ")，重新发送"
	Goto subcaptcha
End If

// downloadCaptchaImage
captchaId = GetStrAB(responseNewCaptcha, "<captchaId>", "</captchaId>")
body = "serviceType=captcha.ws.service.CaptchaService&serviceMethod=downloadCaptchaImage&captchaId=" &captchaId& "&app=" &app& "&userId=" &userId
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Setproxy 2,"127.0.0.1:8888",0
http.open "Get", paramUrlCaptcha & body, false
http.setrequestheader "Host", "captcha.e-pointchina.com" 
http.setrequestheader "Connection", "keep-alive"
http.setrequestheader "Accept","image/webp,image/wxpic,image/sharpp,image/*,*/*;q=0.8"
http.setrequestheader "User-Agent", "Mozilla/5.0 (Linux; Android 4.4.4; M463C Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Mobile Safari/537.36 MicroMessenger/6.3.18.800 NetType/WIFI Language/zh_CN"
http.setrequestheader "Accept-Encoding", "gzip, deflate"
http.setrequestheader "Referer", refererSecKill
http.setrequestheader "Accept-Language","zh-CN,zh;q=0.8"
http.setrequestheader "Cookie",cookieCaptcha
http.send
if isEmpty(http.ResponseBody) Then
	TracePrint "downloadCaptchaImage 失败，重新发送"
	Goto subcaptcha
End If

verify_bit = http.ResponseBody
Set ObjStream = CreateObject("Adodb.Stream")
With ObjStream
	.Type = 1
	.Mode = 3
	.Open
	.Write verify_bit
	.SaveToFile "C:\raw\verify.jpg", 2'
end with
Dim Var
Var = Plugin.Sunday.GetCodeFromFile("C:\raw\verify.jpg", "qq3432872")
Var = Var + 0
set regEx = New RegExp
regEx.[Global] = TRUE
regEx.IgnoreCase = FALSE
regEx.pattern = "<String>(.+)</String>"
Set matches = regEx.execute(str)
Dim count
count = 0
For Each match In matches
	If count = Var Then
		answerId = match.SubMatches(0)
		Exit For
	End If
	count = count + 1
Next


// verifyCaptcha
counterReturn = counterReturn + 1
counterCallback = counterCallback + 1
timestampBase = timestampBase + 1
jsonpReturn = "salama_ws_jsonp_val_" & counterReturn
body = "serviceType=captcha.ws.service.CaptchaService&serviceMethod=verifyCaptcha&captchaId=" &captchaId& "&app=" &app& "&userId=" &userId& "&answerId=" &answerId& "&timestamp=" &timestamp()& "&jsonpReturn=" &jsonpReturn& "&responseType=xml.jsonp&jsoncallback=jsonp_callback" &counterCallback& "&_=" &timestampBase
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Setproxy 2,"127.0.0.1:8888",0
http.open "Get", paramUrlCaptcha & body, false
http.setrequestheader "Host", "captcha.e-pointchina.com" 
http.setrequestheader "Connection", "keep-alive"
http.setrequestheader "Accept","*/*"
http.setrequestheader "User-Agent", "Mozilla/5.0 (Linux; Android 4.4.4; M463C Build/KTU84P) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/33.0.0.0 Mobile Safari/537.36 MicroMessenger/6.3.18.800 NetType/WIFI Language/zh_CN"
http.setrequestheader "Accept-Encoding", "gzip, deflate"
http.setrequestheader "Referer", refererSecKill
http.setrequestheader "Accept-Language","zh-CN,zh;q=0.8"
http.setrequestheader "Cookie",cookieCaptcha
http.send
if isEmpty(http.responsetext) Then
	TracePrint "verifyCaptcha 失败，重新发送"
	Goto subcaptcha
End If

response = http.responsetext
responseVerifyCaptcha = GetStrAB(response, jsonpReturn & ' = "', '"')
responseVerifyCaptcha = decodeURI(responseVerifyCaptcha)
result = GetStrAB(responseVerifyCaptcha, "<result>", "</result>")
If result <> "success" Then
	TracePrint "verifyCaptcha result(" &result& ")，重新发送"
	Goto subcaptcha
End If

captchaPass = GetStrAB(responseVerifyCaptcha, "<captchaPass>", "</captchaPass>")


// doSecKill
body = "channelId=" &paramChannelId& "&appId=" &paramAppId& "&authTicket=" &cookieAuthTicket& "&userId=" &paramUserId& "&sign=" &paramSign& "&sectionId=" &paramSectionId& "&goodsId=" &paramGoodsId& "&goodsSku=&lastCardNo=&captchaId=" &captchaId& "&captchaPass=" &captchaPass& "&appTimestamp=" &appTimestamp& "&appToken=" &accessToken& "&serviceType=com.ebuy.o2o.campaign.service.SeckillService&serviceMethod=doSeckill"
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
With http
	.Setproxy 2,"127.0.0.1:8888",0
	.open "POST", paramUrl & timestamp(), False
	.setrequestheader "Host", "campaign.e-pointchina.com.cn" 
	.setrequestheader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
	.setrequestheader "Connection", "keep-alive"
	.setrequestheader "Content-Length", len(body)
	.setrequestheader "Referer",refererSecKill
	.setrequestheader "Accept-Language", "zh-CN,en-US;q=0.8"
	.setrequestheader "Accept-Encoding", "gzip,deflate"
	.setrequestheader "Origin","http://campaign.e-pointchina.com.cn"
	.setrequestheader "Cookie", cookieSecKill
	.send body
End with

If isEmpty(http.responsetext) Then
	TracePrint "doSecKill 失败，重新发送"
	Goto subSecKill
End
response = http.responsetext
errorCode = GetStrAB(response, "<errorCode>", "</errorCode>")
If errorCode <> "0" Then
	errorMsg = GetStrAB(response, "<errorMsg>", "</errorMsg>")
	TracePrint "doSecKill (" &errorMsg& ")，重新发送"
	Goto subSecKill
End If
orderId = GetStrAB(response, "<orderId>", "</orderId>")
status = GetStrAB(response, "<status>", "</status>")
If orderId = '1' Then
	TracePrint "抢兑成功"
End If


Function GetStrAB(Str, StrA, StrB)
	If InStr(Str,StrA)>0 And InStr(Str,StrB)>0 Then GetStrAB=Split(Split(Str,StrA)(1),StrB)(0)
End Function

Function timestamp()
    Dim js:Set js = CreateObject("ScriptControl")
    js.language = "JScript.encode"
    timestamp = js.EVAL("#@~^FAAAAA==c	+A,fmY+*R7CV!+60v#igYAAA==^#~@ ")
End Function

Function decodeURI(Str)
	Dim js:Set js = CreateObject("ScriptControl")
    js.language = "JScript.encode"
    decodeURI = js.EVAL("decodeURIComponent('" &Str& "')")
End Function
