<%
	for i=1 to 7
	 	message=message + "<strong>"&Request("field_"&i&"_descr")&"</strong>&nbsp;&nbsp;&nbsp;"&Request("field_"&i)&"<br>"
	next
	 	message=message + Request("message")	
		smtpServer = "enter your SMTP SERVER HERE"
		smtpPort = 25
		

		name = Request("your_name")
		Set myMail = CreateObject("CDO.Message") 
		myMail.Subject = "from " & name
		myMail.From = Request("your_email")
		myMail.To = Request("recipient")
		myMail.HTMLBody = "<html><head><title>Contact letter</title></head><body><br>" & message & "</body></html>"
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
		myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
		myMail.Configuration.Fields.Update 
		myMail.Send
	
%>



