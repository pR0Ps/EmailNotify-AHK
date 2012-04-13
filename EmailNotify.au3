;##############################################################
;# EmailNotify v1.0
;# By pR0Ps
;# 
;# Pass it a subject and message and it will send an email
;# The message will be rendered as raw HTML
;# Remember to escape special characters in the message
;# Specfify server details and such in config.ini
;##############################################################

;Used for error handling
Global $oMyRet
Global $oMyError = ObjEvent("AutoIt.Error", "OnError")

Const $CONFIG_FILE = @ScriptDir & "\config.ini"
Const $MSG_TIMEOUT = 10

;##################################
;# Config
;##################################
If $CmdLine[0] < 2 Or $CmdLine[0] > 3 Then
   MsgBox(0, "Incorrect Usage","Usage: " & @ScriptName & " <subject> <line> [conditionText]", $MSG_TIMEOUT)
   Exit (1)
EndIf
$Subject = $CmdLine[1]
$Body = $CmdLine[2]
if $CmdLine[0] == 3 Then
   $ConditionIn = $CmdLine[3]
Else
   $ConditionIn = IniRead($CONFIG_FILE, "Condition", "condText", "")
EndIf

If Not FileExists($CONFIG_FILE) Then
   MsgBox(0, "Error: ","No configuration file found, making a template" & @CRLF & "This program won't work until you configure it", $MSG_TIMEOUT)
   IniWriteSection($CONFIG_FILE, "Server", "SMTPServer=" & @LF & "Username=" & @LF & "Password=" & @LF & "Port=25" & @LF & "SSL=0")
   IniWriteSection($CONFIG_FILE, "Message", "FromAddress=" & @LF & "ToAddress=" & @LF & "CcAddress=" & @LF & "BccAddress=" & @LF & "Importance=Normal"  & @LF & "FromName=")
   Exit (1)
EndIf

$SmtpServer = IniRead($CONFIG_FILE, "Server", "SMTPServer", "")
$Username = IniRead($CONFIG_FILE, "Server", "Username", "")
$Password = IniRead($CONFIG_FILE, "Server", "Password", "")
$Port = IniRead($CONFIG_FILE, "Server", "Port", "25")
$SSL = IniRead($CONFIG_FILE, "Server", "SSL", "0")
$FromAddress = IniRead($CONFIG_FILE, "Message", "FromAddress", "")
$ToAddress = IniRead($CONFIG_FILE, "Message", "ToAddress", "")
$CcAddress = IniRead($CONFIG_FILE, "Message", "CcAddress", "")
$BccAddress = IniRead($CONFIG_FILE, "Message", "BccAddress", "")
$Importance = IniRead($CONFIG_FILE, "Message", "Importance", "Normal")
$FromName = IniRead($CONFIG_FILE, "Message", "FromName", "")
$ConditionText = IniRead($CONFIG_FILE, "Condition", "condText", "")

;##################################
;# Mail Script
;##################################
If $ConditionIn == $ConditionText Then
   $msg = _INetSmtpMailCom($SmtpServer, $FromAddress, $ToAddress, $FromName, $Subject, $Body, $CcAddress, $BccAddress, $Importance, $Username, $Password, $Port, $SSL)
   If @error Then
	  MsgBox(0, "Error sending message", "Error: " & $msg, $MSG_TIMEOUT)
   Else
	  MsgBox(0, "Success!", "Email(s) sent!", $MSG_TIMEOUT)
   EndIf
Else
   MsgBox(0, "Screened", "Email did not pass screening (conditionText not matched)", $MSG_TIMEOUT)
EndIf

Func _INetSmtpMailCom($s_SmtpServer, $s_FromAddress, $s_ToAddress, $s_FromName, $s_Subject = "", $as_Body = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance="Normal", $s_Username = "", $s_Password = "", $Port = 25, $SSL = 0)
   ;Minimum required params
   If $s_SmtpServer == "" or $s_FromAddress == "" or ($s_ToAddress == "" and $s_CcAddress == "" and $s_BccAddress == "") Then
	  MsgBox(0, "Error", "Insufficient parameters. " & @CRLF & "Check " & $CONFIG_FILE & " for missing inforation", $MSG_TIMEOUT)
	  Exit(1)
   EndIf
   
   Local $objEmail = ObjCreate("CDO.Message")
   
   ;Sender
   $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
   
   ;Recipients
   $objEmail.To = $s_ToAddress
   If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
   If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
   
   $objEmail.Subject = $s_Subject
   
   ;HTML vs. plaintext
   If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
	  $objEmail.HTMLBody = $as_Body
   Else
	  $objEmail.Textbody = $as_Body & @CRLF
   EndIf
   
   ;SMTP Server
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $Port
   
   ;Authenticated SMTP
   If $s_Username <> "" Then
      $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
      $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
      $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
   EndIf
   
   ;SSL
   If $SSL Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
   EndIf
   
   ;Set Email Importance
   If $s_Importance == "High" or $s_Importance == "Normal" or $s_Importance == "Low" Then
	  $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = $s_Importance
   Else
	  $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
   EndIf
   
   ;Update settings
   $objEmail.Configuration.Fields.Update
   $objEmail.Fields.Update
   
   ;Send the Message
   $objEmail.Send
   
   If @error Then
	  SetError(2)
	  Return $oMyRet
   EndIf
   
   ;Clean up
   $objEmail=""
EndFunc

;Catches errors in sending
Func OnError()
   $oMyRet = StringStripWS($oMyError.description, 3)
   SetError(1)
   Return
EndFunc