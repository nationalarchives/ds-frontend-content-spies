<!-- #include virtual="/spies/codemaster/includes/ss-header.asp" -->
<!-- #include virtual="/spies/codemaster/includes/ss-functions.asp" -->
<%
Dim lngMessageID
Dim arrMessage
Dim strMessageText
Dim strMessageImageUrl
Dim strMessageHelpUrl
Dim strMessageNote

On Error Resume Next

' Validate txtMessageID (input)
lngMessageID = CLng(Request("txtMessageID"))  ' runtime error if cannot be converted - can be in form or querystring
If lngMessageID < 1 Or lngMessageID > NUM_MESSAGES Then
	Err.Raise vbObjectError + 1, "Secrets and Spies", "Message ID out of range"
End If

If Err.Number <> 0 Then Response.Redirect "default.asp"

arrMessage         = GetMessage(lngMessageID)
strMessageText     = arrMessage(MESSAGE_TEXT)
strMessageImageUrl = arrMessage(MESSAGE_IMAGE_URL)
strMessageNote     = arrMessage(MESSAGE_NOTE)
strMessageHelpUrl  = arrMessage(MESSAGE_HELP_URL)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>The National Archives | Research, education & online exhibitions | Exhibitions | Secrets and Spies</title>
<!--#include virtual="/includes/google-analytics-gtm-head.inc" -->

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="DC.Publisher" content="Public Record Office">
<meta name="DC.Type" content="text;image">
<meta name="DC.Format" content="text/html">
<meta name="DC.Source" content="Public Record Office, The National Archives">
<meta name="DC.Rights" content="http://www.nationalarchives.gov.uk/legal/copyright.htm">
<meta name="DC.Creator" content="Public Record Office">
<meta name="DC.Language" CONTENT="en-UK">
<meta name="DC.Identifier" content="http://www.pro.gov.uk/virtualmuseum/spies/default.htm">

<script language="JavaScript" type="text/JavaScript" src="../scripts/roll_pop.js"></script>
<link href="../css/codemaster_ns4.css" rel="stylesheet" type="text/css">
<STYLE type="text/css">
@import url(../css/codemaster_ie56ns67.css);
</STYLE>
<link href="/css/menusmicrosites/breadcrumb.css" rel="stylesheet" type="text/css">
</head>
<body  onLoad="MM_preloadImages('../scripts/pixelwhite.gif','../images/nav_codeciper_d.gif','../images/nav_spies_d.gif')">
<!--#include virtual="/includes/google-analytics-gtm-body.inc" -->
<!--#include virtual="/includes/menusmicrosites/ss_breadcrumb.inc" -->
<!-- banner start -->
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="20"><a name="top"><img src="../images/pixeltrans.gif" width="20" height="20" alt="*"></a></td>
    <td width="330"><img src="../images/pixeltrans.gif" width="330" height="20" alt="*"></td>
    <td width="285" align="right"><img src="../images/pixeltrans.gif" width="285" height="20" alt="*"></td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="20">&nbsp;</td>
    <td width="330"><a href="../default.htm"><img src="../images/sas_logo_olive.jpg" alt="secrets and spies" width="295" height="42" border="0"></a></td>
    <td width="285" align="right" valign="top"><!--#include file="../eventsinclude.inc" --></td>
  </tr>
</table>
<!-- banner end -->
<!-- navbar start -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
   <tr align="left" valign="top"> 
    <td width="650"><img src="../images/pixeltrans.gif" height="2" width="650" alt="*"></td>    
    <td width="100%"><img src="../images/pixeltrans.gif" height="2" width="130" alt="*"></td>
   </tr>
   <tr align="left" valign="top"> 
    <td width="650"><img src="../images/pixel_redline650.gif" height="1" width="650" alt="*"></td>    
    <td width="100%" bgcolor="#ED1C24"><img src="../images/pixel_ED1C24.gif" width="130" height="1" alt="*"></td>
   </tr>
   <tr align="left" valign="top"> 
    <td width="650" bgcolor="#000000"><img src="../images/pixel_000000.gif" height="10" width="650" alt="*"></td>    
    <td width="100%" bgcolor="#000000"><img src="../images/pixel_000000.gif" width="130" height="10" alt="*"></td>
   </tr>
   <tr align="left" valign="top"> 
    <td width="650" align="right" valign="bottom" bgcolor="#000000"><img src="../images/pixel_000000.gif" height="16" width="312" alt="*"><a href="../ciphers/default.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('codes','','../images/nav_codeciper_d.gif',1)"><img src="../images/nav_codeciper_u.gif" alt="codes and cipher" name="codes" width="163" height="16" border="0"></a><img src="../images/pixel_000000.gif" height="16" width="13" alt="*"><a href="../spies/default.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('spies','','../images/nav_spies_d.gif',1)"><img src="../images/nav_spies_u.gif" alt="spies" name="spies" width="45" height="16" border="0"></a><img src="../images/pixel_000000.gif" height="16" width="13" alt="*"><a href="default.asp"><img src="../images/nav_codemaster_d.gif" alt="codemaster" width="103" height="16" border="0"></a></td>    
    <td width="100%" bgcolor="#000000"><img src="../images/pixel_000000.gif" width="130" height="16" alt="*"></td>
   </tr>
   <tr align="left" valign="top"> 
    <td width="650" bgcolor="#000000"><img src="../images/pixel_000000.gif" height="2" width="650" alt="*"></td>    
    <td width="100%" bgcolor="#000000"><img src="../images/pixel_000000.gif" width="130" height="2" alt="*"></td>
   </tr>     
</table>
<!-- navbar end -->
<!-- breadcrumb start -->
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="615"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
    <td width="615"><a class="breadcrumb" href="../default.htm">Home</a> 
      <span class="breadcrumb"> > </span><a class="breadcrumb" href="default.asp">Codemaster</a> 
      <span class="breadcrumb"> > Send a Message</span></td>
  </tr>
</table>
<!-- breadcrumb end -->
<!-- body start -->
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="20" alt="*"></td>
    <td width="625"><img src="../images/pixeltrans.gif" width="25" height="20" alt="*"></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><span class="bodyblack">You have selected the message: "<b><%= Server.HTMLEncode(strMessageText) %></b>"</span></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><span class="bodyblack"><%= Server.HTMLEncode(strMessageNote) %></span></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><span class="bodyblack">Here it is encrypted into the <a href="<%= strMessageHelpURL %>.htm" target="_blank" onClick="PRO_openPopupWindow('<%= strMessageHelpURL %>.htm','popup','591','400','scrollbars=yes,resizable=yes','yes');return document.MM_returnValue">Pigpen 
      Cipher</a>:</span></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
    <td width="625"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
</tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="625"><img src="<%= strMessageImageUrl %>" alt="encrypted message"></td>
</tr>
</table>
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="20" alt="*"></td>
    <td width="330"><img src="../images/pixeltrans.gif" width="330" height="20" alt="*"></td>
    <td width="30"><img src="../images/pixeltrans.gif" width="30" height="20" alt="*"></td>
    <td width="255"><img src="../images/pixeltrans.gif" width="255" height="20" alt="*"></td>
  </tr>

</table>
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="15" alt="*"></td>
      <td width="625"><img src="../images/pixeltrans.gif" width="25" height="15" alt="*"></td>
  </tr>	
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625"><span class="bodyblackbold">To e-mail to a friend, please complete this form.</span></td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
      <td width="625"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
  </tr>
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      
    <td width="625"><span class="bodyblack">(fields marked with an asterisk "*" 
      are compulsory). Please read our <a href="terms.htm" target="_blank" onClick="PRO_openPopupWindow('terms.htm','popup','591','400','scrollbars=yes,resizable=yes','yes');return document.MM_returnValue">terms of use</a>.</span></td>
  </tr>
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
  </tr>    
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625">
	  <form action="sent_ok.asp" method="post">
	    <table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="240" align="right"><span class="notes">Your name: *</span></td>
            <td width="495"><input class="formbox" type="text" name="txtSenderName" id="txtSenderName" value="">
            </td>
          </tr>
          <tr> 
            <td width="240" align="right"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
          </tr>
          <tr> 
            <td width="240" align="right"><span class="notes">Your e-mail address: 
              *</span></td>
            <td width="495"><input class="formbox" type="text" name="txtSenderEmail" id="txtSenderEmail" value=""></td>
          </tr>
          <tr> 
            <td width="240" align="right"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
          </tr>
          <tr> 
            <td width="240" align="right"><span class="notes">Your friend's name: 
              *</span></td>
            <td width="495"><input class="formbox" type="text" name="txtRecipientName" id="txtRecipientName" value=""></td>
          </tr>
          <tr> 
            <td width="240" align="right"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
          </tr>
          <tr> 
            <td width="240" align="right"><span class="notes">Your friend's e-mail 
              address: *</span></td>
            <td width="495"><input class="formbox" type="text" name="txtRecipientEmail" id="txtRecipientEmail" value=""></td>
          </tr>
          <tr> 
            <td width="240" align="right"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
          </tr>
          <tr> 
            <td width="240" align="right"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495" valign="top"> <input type="submit" name="btnSend" id="btnSend" value="Send"> 
              <input type="hidden" name="txtMessageID" id="txtMessageID" value="<%= lngMessageID %>"> 
            </td>
          </tr>
          <tr> 
            <td width="240"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
            <td width="495"><img src="../images/pixeltrans.gif" width="5" height="5" alt="*"></td>
          </tr>
        </table>
</form>
</td>
  </tr>
    <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      
    <td width="625"><span class="bodyblackbold">To select a different message, return 
      to the <a href="default.asp">previous page</a>.</span></td>
  </tr> 
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
    <td width="625"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
</tr>  
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625">&nbsp;</td>
</tr>  
</table>
<!-- body end -->
</body>
</html>
