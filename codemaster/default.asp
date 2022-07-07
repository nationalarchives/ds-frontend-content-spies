<!-- #include virtual="/spies/codemaster/includes/ss-header.asp" -->
<!-- #include virtual="/spies/codemaster/includes/ss-functions.asp" -->
<%
Dim arrMessages
Dim i

On Error Resume Next

arrMessages = GetAllMessages()  ' from the database table

If Err.Number <> 0 Then Call HandleError(Null)
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
<meta name="DC.Identifier" content="http://www.nationalarchives.gov.uk/online/spies/default.asp">

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
      <span class="breadcrumb"> > </span><span class="breadcrumb">Codemaster</span></td>
  </tr>
</table>
<!-- breadcrumb end -->
<!-- body start -->

<table width="650" border="0" cellspacing="0" cellpadding="0">
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
    <td width="260"><img src="../images/pixeltrans.gif" width="260" height="10" alt="*"></td>
    <td width="30"><img src="../images/pixeltrans.gif" width="30" height="10" alt="*"></td>
    <td width="335"><img src="../images/pixeltrans.gif" width="335" height="10" alt="*"></td>
  </tr>
  <tr align="left" valign="top"> 
    <td width="25"><img src="../images/pixeltrans.gif" width="25" height="20" alt="*"></td>
    <td width="260"><a href="cipher_symbol_x.htm" target="_blank" onClick="PRO_openPopupWindow('cipher_symbol_x.htm','popup','591','400','scrollbars=yes,resizable=yes','yes');return document.MM_returnValue"><img src="images/interactive.jpg" alt="click here for enlarged image of cipher symbols" width="260" height="116" border="0"></a></td>
    <td width="30">&nbsp;</td>
    <td width="335"><table width="335" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/title_codemaster.gif" width="208" height="23" alt="codemaster"></td>
        </tr>
        <tr>
          <td><img src="../images/pixeltrans.gif" width="255" height="10" alt="*"></td>
        </tr>		
        <tr>
          <td align="left" valign="top"><span class="bodyblack"><strong>Send and 
            receive messages using this original French cipher, which we have 
            adapted for English.</strong></span></td>
        </tr>
        <tr>
          <td><img src="../images/pixeltrans.gif" width="255" height="10" alt="*"></td>
        </tr>
        <tr>
          <td align="left" valign="top"><span class="bodyblack">This is called 
            the 'Pigpen' Cipher and has been used to send secret messages by a 
            number of groups, including Napoleon's spies. (Find out more in our 
            <a href="../ciphers/scovell/default.htm">Codes &amp; Ciphers</a> gallery.)</span></td>
        </tr>
        <tr>
          <td><img src="../images/pixeltrans.gif" width="255" height="10" alt="*"></td>
        </tr>
        <tr>
          <td><img src="../images/pixeltrans.gif" width="255" height="10" alt="*"></td>
        </tr>		
      </table></td>
  </tr>

</table>
<table width="650" border="0" cellspacing="0" cellpadding="0">
  <form action="send.asp" method="post">
    <tr align="left" valign="top"> 
      <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625"><span class="bodyblack">Here are some real messages from spies and spy-related sayings from times past. Select a message, encrypt it and e-mail it to a friend.</span></td>
  </tr>
    <tr align="right" valign="top"> 
      <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
      <td width="625"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
  </tr>	
    <tr align="right" valign="top"> 
      <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625">
  <!-- Message list -->
  <select class="dropbox" name="txtMessageID" id="txtMessageID">
  <option value="0" selected="selected">-- Pick a message --</option>
<%
For i = 0 to UBound(arrMessages)
	Response.Write "<option value=""" & Server.HTMLEncode(arrMessages(i, MESSAGE_ID)) & """>" & Server.HTMLEncode(arrMessages(i, MESSAGE_TEXT)) & "</option>" & vbNewLine
Next
%>
  </select>
  <!-- Select button passes message ID to step 2 -->
  </td>
    </tr>
    <tr align="right" valign="top"> 
      <td width="25"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
      <td width="625"><img src="../images/pixeltrans.gif" width="25" height="5" alt="*"></td>
  </tr>  
    <tr align="right" valign="top"> 
      <td width="25"><img src="../images/pixeltrans.gif" width="25" height="10" alt="*"></td>
      <td width="625"><input type="submit" name="btnSelect" id="btnSelect" value="Select"></td>
  </tr></form>  
</table>
<!-- body end -->
</body>
</html>
