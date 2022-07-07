<%
' *** Set this to False when debugging is complete ***
Const DEBUG_ENABLED = True  ' display full error messages

' The e-mailing code detects when it is on the Internet or intranet and chooses which 
' of these two SMTP hosts to use. You do NOT need to comment out one of the host names 
' for this to work in both environments.
' Update: SMPT HOST HAME NOT NEEDED ANYMORE - THIS IS INCLUDED IN THE NationalArchives.Email.Message COMPONENT jg, 22-Jun-04
'Const LIVE_MAIL_HOST = "mail.nationalarchives.gov.uk"  ' live SMTP server name
'Const DEVT_MAIL_HOST = "pro06a"           ' development SMTP server name

' The e-mail must come from a PRO address, otherwise the mail server will block it.
' However, the reply-to address is set to the user's address (as entered by them in the
' web page form), so e-mail replies should go to the user, not to this address.
Const MAIL_FROM_ADDR = "education@nationalarchives.gov.uk"  ' the e-mails come from this address
Const MAIL_FROM_NAME = "Secrets and Spies"     ' the from name displayed in a user's inbox

Const NUM_MESSAGES = 9	' number of messages (rows)
Const NUM_FIELDS   = 5  ' number of data fields per message (columns)

' These are the array indicies of each data field as returned by these functions
Const MESSAGE_ID        = 0
Const MESSAGE_TEXT      = 1
Const MESSAGE_IMAGE_URL = 2
Const MESSAGE_NOTE      = 3
Const MESSAGE_HELP_URL  = 4

' Set up the data array with the information about each encrypted message
' (Would like to store this in an Application variable,
' but the virtual directory is not set up as an ASP application)
Dim garrMessages()
ReDim garrMessages(NUM_MESSAGES - 1, NUM_FIELDS - 1)

' ------------------------------------------------------------

garrMessages(0, MESSAGE_ID)        = 1
garrMessages(0, MESSAGE_TEXT)      = "Silence is of the Gods... only monkeys chatter"
garrMessages(0, MESSAGE_IMAGE_URL) = "images/message01.gif"
garrMessages(0, MESSAGE_NOTE)      = "This phrase was written on the wall at Station XV, where SOE equipment was made and tested. ('Secret Agent's Handbook', M. Seaman ed., PRO)"
garrMessages(0, MESSAGE_HELP_URL)  = "help/help01"

' ------------------------------------------------------------

garrMessages(1, MESSAGE_ID)        = 2
garrMessages(1, MESSAGE_TEXT)      = "A short report in time is worth an encyclopaedia out of date."
garrMessages(1, MESSAGE_IMAGE_URL) = "images/message02.gif"
garrMessages(1, MESSAGE_NOTE)      = "This wise advice for spies comes from Lt-Col RWG Stephens, Commandant of Camp 020, MI5's wartime interrogation centre for enemy agents. (Stephens, 'A Digest of Ham', reproduced in 'Camp 020', O. Hoare ed., PRO)"
garrMessages(1, MESSAGE_HELP_URL)  = "help/help02"

' ------------------------------------------------------------

garrMessages(2, MESSAGE_ID)        = 3
garrMessages(2, MESSAGE_TEXT)      = "As with a man, so with a woman. There is no room for chivalry in modern espionage."
garrMessages(2, MESSAGE_IMAGE_URL) = "images/message03.gif"
garrMessages(2, MESSAGE_NOTE)      = "This wise advice for spies comes from Lt-Col RWG Stephens, Commandant of Camp 020, MI5's wartime interrogation centre for enemy agents. (Stephens, 'A Digest of Ham', reproduced in 'Camp 020', O. Hoare ed., PRO)"
garrMessages(2, MESSAGE_HELP_URL)  = "help/help03"

' ------------------------------------------------------------

garrMessages(3, MESSAGE_ID)        = 4
garrMessages(3, MESSAGE_TEXT)      = "Never promise, never bargain. The man's neck is in your grasp. Never forget it; never let him forget it."
garrMessages(3, MESSAGE_IMAGE_URL) = "images/message04.gif"
garrMessages(3, MESSAGE_NOTE)      = "This wise advice for spies comes from Lt-Col RWG Stephens, Commandant of Camp 020, MI5's wartime interrogation centre for enemy agents. (Stephens, 'A Digest of Ham', reproduced in 'Camp 020', O. Hoare ed., PRO)"
garrMessages(3, MESSAGE_HELP_URL)  = "help/help04"

' ------------------------------------------------------------

garrMessages(4, MESSAGE_ID)        = 5
garrMessages(4, MESSAGE_TEXT)      = "Do you want a shipment of chocolate? My pyjamas are blue"
garrMessages(4, MESSAGE_IMAGE_URL) = "images/message05.gif"
garrMessages(4, MESSAGE_NOTE)      = "Spies were instructed to use bizzare question and answer phrases such as this to identify each other in radio comunications during the Second World War. ('Secret warfare, the Arms and Techniques of the Resistance', P. Lorain)"
garrMessages(4, MESSAGE_HELP_URL)  = "help/help05"

' ------------------------------------------------------------

garrMessages(5, MESSAGE_ID)        = 6
garrMessages(5, MESSAGE_TEXT)      = "Aux chevaux maigres vont les mouches"
garrMessages(5, MESSAGE_IMAGE_URL) = "images/message06.gif"
garrMessages(5, MESSAGE_NOTE)      = "This phrase, from a French novel, was used as the basis for WWII double agent Treasure's secret Ahbwer code. Translated into English, it means ""Flies go to thin horses."" (Catalogue reference: KV 2/465)"
garrMessages(5, MESSAGE_HELP_URL)  = "help/help06"

' ------------------------------------------------------------

garrMessages(6, MESSAGE_ID)        = 7
garrMessages(6, MESSAGE_TEXT)      = "Three may keep a secret if two of them are dead."
garrMessages(6, MESSAGE_IMAGE_URL) = "images/message07.gif"
garrMessages(6, MESSAGE_NOTE)      = "Benjamin Franklin coined this phrase in his 'Poor Richard's Almanac', 1753."
garrMessages(6, MESSAGE_HELP_URL)  = "help/help07"

' ------------------------------------------------------------

garrMessages(7, MESSAGE_ID)        = 8
garrMessages(7, MESSAGE_TEXT)      = "There are no secrets better kept than the secrets that everybody guesses."
garrMessages(7, MESSAGE_IMAGE_URL) = "images/message08.gif"
garrMessages(7, MESSAGE_NOTE)      = "This is a line from George Bernard Shaw's 1893 play, 'Mrs Warren's Profession'."
garrMessages(7, MESSAGE_HELP_URL)  = "help/help08"

' ------------------------------------------------------------

garrMessages(8, MESSAGE_ID)        = 9
garrMessages(8, MESSAGE_TEXT)      = "Those who know the enemy as well as they know themselves will never suffer defeat."
garrMessages(8, MESSAGE_IMAGE_URL) = "images/message09.gif"
garrMessages(8, MESSAGE_NOTE)      = "This is advice from Sun Tzu's 4th Century BC book 'Art of War'."
garrMessages(8, MESSAGE_HELP_URL)  = "help/help09"

' ------------------------------------------------------------

Function GetAllMessages()

	GetAllMessages = garrMessages

End Function  ' GetAllMessages()

' ------------------------------------------------------------

Function GetMessage(ByVal lngMessageID)

	Dim arrMessage()
	Dim i

	ReDim arrMessage(NUM_FIELDS - 1)

	For i = 0 To UBound(garrMessages)
		If garrMessages(i, MESSAGE_ID) = lngMessageID Then
			arrMessage(MESSAGE_ID)        = garrMessages(i, MESSAGE_ID)
			arrMessage(MESSAGE_TEXT)      = garrMessages(i, MESSAGE_TEXT)
			arrMessage(MESSAGE_IMAGE_URL) = garrMessages(i, MESSAGE_IMAGE_URL)
			arrMessage(MESSAGE_NOTE)      = garrMessages(i, MESSAGE_NOTE)
			arrMessage(MESSAGE_HELP_URL)  = garrMessages(i, MESSAGE_HELP_URL)
			Exit For
		End If
	Next

	GetMessage = arrMessage

End Function  ' GetMessage()

' ------------------------------------------------------------

Function IsValidEmailFormat(ByVal emailAddress)
	' Email addresses are difficult to validate because they can include comments, quoted strings and whitespace as well as the 'normal' part of the address. This function is more restrictive, and only expects the 'name@domain' part of an Internet e-mail address without quoted strings. It also assumes that the domain consists of at least two parts separated by a dot.

	' See RFC 2822 for definition of an atom. The dash is quoted by the backslash. The backslash itself is not a valid character.

	Const cstrAtom = "[A-Za-z0-9!#$%&'*+\-/=?^_`{|}~]+"
	Dim strPattern

	strPattern = _
		"^"   & cstrAtom & _
		"(\." & cstrAtom & ")*" & _
		"@"   & cstrAtom & _
		"(\." & cstrAtom & ")+$"

	IsValidEmailFormat = IsValid(emailAddress, strPattern, True, 128)

End Function  ' IsValidEmailFormat()

' ------------------------------------------------------------

Function IsValid(ByVal value, ByVal pattern, ByVal ignoreCase, ByVal maxLength)
	'------------------------------------------------------------------------------
	' Function:    IsValid
	' System:      PRO Web Application Shared Components
	' Copyright:   Public Record Office, 2002
	' Description: Match a value against a regular expression pattern.
	'              Return True is the value matches the pattern, otherwise return False.
	'              The match can be case-sensitive or case-insensitive.
	'              The function can also check for a maximum length.
	'              If the pattern also implies a length (e.g. ^\d{1,4}$), then the most restrictive is applied.
	' Arguments:   value (String) - the string to validate
	'              pattern (String) - the regular expression pattern
	'              ignoreCase (Boolean) - if ignoreCase is True then do a case-insensitive match
	'              maxLength (Long) - the maximum length of value
	'                  If the length of value is greater than maxLength then the function returns False.
	' History:
	' Auth  Date           Comments
	' JG    20-Jan-2003    Created
	'------------------------------------------------------------------------------

	Const cstrErrSource = "IsValid()"

	Dim re

	IsValid = False  ' Default return value is False

	' Validate function arguments
	On Error Resume Next

		value = CStr(value)
		If Err.Number <> 0 Then
			On Error Goto 0
			Err.Clear
			Err.Raise vbObjectError + 1, cstrErrSource, "The value to be validated must be a string"
		End If

		pattern = CStr(pattern)
		If Err.Number <> 0 Then
			On Error Goto 0
			Err.Clear
			Err.Raise vbObjectError + 2, cstrErrSource, "The pattern to validate against must be a string"
		End If

		ignoreCase = CBool(ignoreCase)
		If Err.Number <> 0 Then
			On Error Goto 0
			Err.Clear
			Err.Raise vbObjectError + 3, cstrErrSource, "The ""ignore case"" flag must be True or False"
		End If

		maxLength = CLng(maxLength)
		If Err.Number <> 0 Or maxLength < 0 Then
			On Error Goto 0
			Err.Clear
			Err.Raise vbObjectError + 4, cstrErrSource, "The maximum length must be an integer between -1 and 2,147,483,647"
		End If

	On Error Goto 0

	If Len(value) > maxLength Then
		' Input is too long
		IsValid = False
		Exit Function
	End If

	Set re = New RegExp
	re.Pattern = pattern
	re.IgnoreCase = ignoreCase
	IsValid = re.Test(value)
	Set re = Nothing

End Function  ' IsValid()

' ------------------------------------------------------------

Sub HandleError(ByVal message)

	Dim strHTML

	If IsNull(message) Then
		message = "Sorry, there is a problem with this Web site which means we could not process your request."
	End If

	' DEBUG >>
	If DEBUG_ENABLED Then
		message = message & "<br><br>(" & Err.Number & " : " & Err.Source & " : " & Err.Description & ")"
	End If
	' << DEBUG

	strHTML = ReadFile(Server.MapPath("error-screen.htm"))
	strHTML = Replace(strHTML, "$$MESSAGE$$", message)

	Response.Clear
	Response.Write strHTML
	Response.Flush
	Response.End

End Sub  ' HandleError()

' ------------------------------------------------------------

Sub SendEmail(messageID, messageImageUrl, senderName, senderEmail, recipientName, recipientEmail)

	Dim strHtmlEmailBody
	Dim strTextEmailBody
	Dim strAppPath
	Dim objMail

	strAppPath = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	strAppPath = Left(strAppPath, InStrRev(strAppPath, "/"))

	' Create and send HTML e-mail with link to image, link back to site and alternative text body - read in the template file and replace the place holders
	strHtmlEmailBody = ReadFile(Server.MapPath("includes/email-template.htm"))
	strHtmlEmailBody = Replace(strHtmlEmailBody, "$$SENDER_NAME$$", Server.HTMLEncode(senderName))
	strHtmlEmailBody = Replace(strHtmlEmailBody, "$$RECIPIENT_NAME$$", Server.HTMLEncode(recipientName))
	strHtmlEmailBody = Replace(strHtmlEmailBody, "$$ENCRYPTED_MESSAGE_URL$$", strAppPath & messageImageUrl)
	strHtmlEmailBody = Replace(strHtmlEmailBody, "$$RETURN_URL$$", strAppPath & "message.asp?mid=" & Server.URLEncode(messageID) & "&nm=" & Server.URLEncode(recipientName))

	strTextEmailBody = ReadFile(Server.MapPath("includes/email-template.txt"))
	strTextEmailBody = Replace(strTextEmailBody, "$$SENDER_NAME$$", Server.HTMLEncode(senderName))
	strTextEmailBody = Replace(strTextEmailBody, "$$RECIPIENT_NAME$$", Server.HTMLEncode(recipientName))
	strTextEmailBody = Replace(strTextEmailBody, "$$ENCRYPTED_MESSAGE_URL$$", strAppPath & messageImageUrl)
	strTextEmailBody = Replace(strTextEmailBody, "$$RETURN_URL$$", strAppPath & "message.asp?mid=" & Server.URLEncode(messageID) & "&nm=" & Server.URLEncode(recipientName))

	Set objMail = Server.CreateObject("NationalArchives.Email.Message")
	With objMail
' NO NEED TO SET HOST WITH NationalArchives.Email.Message COMPONENT
'		If InStr(Request.ServerVariables("SERVER_NAME"), ".") > 0 Then
'			' assume running in Internet domain
'			.Host = LIVE_MAIL_HOST
'		Else
'			' assume running in intranet domain
'			.Host = DEVT_MAIL_HOST
'		End If
		.From     = """" & MAIL_FROM_NAME & """ <" & MAIL_FROM_ADDR & ">"
		.To       = """" & recipientName & """ <" & recipientEmail & ">"
		.ReplyTo  = """" & senderName & """ <" & senderEmail & ">"
		.Subject  = "A secret message for you to decode"
		.HTMLBody = strHtmlEmailBody
		.TextBody = strTextEmailBody
		.Send
		.Close
	End With
	Set objMail = Nothing

End Sub  ' SendEmail()

' ------------------------------------------------------------

Function ReadFile(ByVal path)
	' --------------------------------------------------------------------------
	' Function to read the contents of a file, given by path, and return them
	' as a string.
	' --------------------------------------------------------------------------

	Dim objFileSystem  ' FileSystem object
	Dim objTextStream  ' TextStream object

	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFileSystem.OpenTextFile(path)

	ReadFile = objTextStream.ReadAll

	objTextStream.Close
	Set objTextStream = Nothing
	Set objFileSystem = Nothing

End Function  ' ReadFile()

' ------------------------------------------------------------
%>
