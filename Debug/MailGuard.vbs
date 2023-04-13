'MG Version 7
'This version uses ODBC to talk to MySQL stored proc MGCheck which does all the work
'MGChecker not required.
'Also added failedMessageLogging by tapping into OnDeliveryFailed event
'MG now sends emails after removing blocked recipients, modifies CC and TO headers
'also sends reject email to the sender explaining why the message was rejected

'Option Explicit

' Global Settings
Dim MGLogPath,FailedDeliveryLogPath
MGLogPath = "D:\MailGuard\Log\log.txt"
FailedDeliveryLogPath = "D:\MailGuard\Log\failed_Delivery.txt"
Private Const setting_username = "Administrator"
' Set your hmailserver password here!
Private Const setting_password = "YOUR HMAIL SERVER PASSWORD"
Private Const LogLevel = 1 '0- Nothing, 1-Only Errors, 2-Everything
'End of global settings.

	
Sub OnAcceptMessage(oClient, oMessage)

On Error Resume Next



if lcase(oMessage.From) = "hmailserver" THEN 
	Result.Value = 0
	Exit Sub 'Do not run for messages from hMailServer that are sent from CreateFailedDeliveryLogMessage
End IF

if lcase(oClient.UserName) <> ""  THEN 'Run only for authenticated users sending outbound email 

	Dim objConnection,objRecordSet
	
	Set objConnection = CreateObject("ADODB.Connection")

		'Here is the connection. Create a SYSTEM DSN called MailGuard and give access to hmail database
	objConnection.Open "DSN=MailGuard"
	
	If Err <> 0 then
		Result.Message = Err.Description & "-" & "Unable to open the DSN. Try sending your mail again."
		Result.value = 2
		If LogLevel>0 THEN WriteLog MGLogPath,Err.Description & "-" & "DSN Error!"
		Exit Sub
	End If
	
	Dim InputString,ResultString
			Dim MailGuardID : MailGuardID = "MailGuard <mailguard@YOURDOMAIN.com>"
	
	If LogLevel>1 THEN WriteLog MgLogPath, "MG Starting"

	'Section 1: Anonymizing the sender IP
	oMessage.HeaderValue("Received") = "from a friendly server; " & now

	If LogLevel>1 THEN  WriteLog MGLogPath,"Finished anonymizing"

	'Section 2: Checking recipients
	Dim Recipients 
	Recipients = ""

	For i = 0 To oMessage.Recipients.Count-1
		   Recipients = Recipients & oMessage.Recipients(i).OriginalAddress & ","
	Next

	Recipients = truncate_one(Recipients) 'Remove the trailing comma if it exists
		  
	InputString = oClient.ipaddress & ":" & oMessage.fromAddress & ":" & Recipients
	ResultString = ""
	Err.Clear

	If LogLevel>1 THEN WriteLog MGlogPath,  InputString
	Dim MGQuery
	MGQuery = "SELECT hmaildb.MGCheck(""" & oClient.userName & """,""" & Recipients & """) as `C`;"
	If LogLevel>1 THEN WriteLog MGlogPath,  "SQL=" & MGQuery
	
	Set objRecordset = objConnection.Execute(MGQuery)
	ResultString = objRecordSet.Fields("C")	
	
	objConnection.close
		
	If LogLevel>1 THEN  WriteLog MGLogPath,"Result Received:" & ResultString
	
	If Err.Number <> 0 Then
		If LogLevel>0 THEN WriteLog MGLogPath,"Error: " & Err.Number & "-" & Err.Description
		If LogLevel>0 THEN WriteLog MgLogPath,"Error: " & ResultString
		ResultString="Z" 'Locally generated ResultString, handle it
		Err.Clear
	End If
	
	if ResultString<>"" Then 
		'Section 3: Handling the exception
			If LogLevel>1 THEN  WriteLog MGLogPath, "Received result:" & ResultString
			Dim nl
			nl = Chr(13) & Chr(10)
			'Return codes: 
			'1. X - not defined in MG Rules
			'2. IP - Sending IP not defined ub MGAllowedIPsTable
			'3. List of addresses that are not allowed as per rules
			'4. Blank - everything OK
			'5 - Z - MailGuard not running
		Select case ResultString
			case "Z"
				Result.Message = "MailGuard is not running properly on the server! Please contact IT immediately"
				Result.Value = 2
			Case "X"
				SendMessage MailGuardID,oMessage.FromAddress,"Could not send: " & oMessage.Subject,"Your email could not be sent as your email address (" & oClient.UserName  & ") has not been configured in MailGuard." & nl & "Please contact IT support."
				Result.Message = "Your mail could not be sent as your address (" & oClient.UserName & ") has not been configured in MailGuard."
				Result.Value = 2
				
			Case "IP"
				SendMessage MailGuardID,oMessage.FromAddress,"Could not send: " & oMessage.Subject,"Your email could not be sent as your IP address (" & oClient.ipaddress & ") is not configured in MailGuard. " & nl & "Please reply to this email to have it configured."
				Result.Message = ResultString
				Result.Value = 2
				
			Case Else 'A list of banned addresses has been received from
					
				Dim FinalRecipients
				FinalRecipients=""
				Dim i,x
				
				Dim ToHeader,CCHeader
				Dim emailAddress,fullAddress
				Dim newToHeader,newCCHeader

				ToHeader = oMessage.TO
				CCHeader = oMessage.CC
				newToHeader=""
				newCCHeader=""
				
				'Loop through the recipients and see if they are contained in the returned list of disallowed IDs. Create the list of FinalRecipients who will be added back later
				For i = 0 To oMessage.Recipients.Count-1
					If instr(1,LCase(ResultString),LCase(oMessage.Recipients(i).OriginalAddress),1)=0 Then 'If the recipient is not in the list of disallowed IDs
						FinalRecipients = FinalRecipients & oMessage.Recipients(i).OriginalAddress & ","
					End If
				Next
				'Loop through the To and CC headers to find any disallowed Ids. We are re-creating the TO and CC headers after removing the disallowed IDs entirely (name and email ID)
				for each fullAddress in split(ToHeader,",")
					'Get the email address portion from "Name <EmailID>"
					emailAddress = MID(fullAddress,InstrRev(fullAddress,"<")+1,Len(fullAddress)-InstrRev(fullAddress,"<")-1)
					If instr(1,LCase(ResultString),LCase(emailAddress))=0 Then 'If the disallowed ID list does not contain the ID from the ToList
						newToHeader = newToHeader & fullAddress & ", "
					End If
				Next 
				'Loop thru the CC header
				for each fullAddress in split(CCHeader,",")
				'Get the email address portion from "Name <EmailID>"
					emailAddress = MID(fullAddress,InstrRev(fullAddress,"<")+1,Len(fullAddress)-InstrRev(fullAddress,"<")-1)
					If instr(1,LCase(ResultString),LCase(emailAddress))=0 Then 'If the disallowed ID list does not contain the ID from the ToList
						newCCHeader = newCCHeader & fullAddress & ", "
					End If
				Next 
				'Remove any trailing commas
				newToHeader=truncate_one(newToHeader)
				newCCHeader=truncate_one(newCCHeader)
				
				'Clear all recipients from the message
				oMessage.ClearRecipients()
				
				'Add back the final recipients
				for each x in split(FinalRecipients,",")
					if Trim(x)<>"" then
						'EventLog.Write("Now Adding: " & x)	  
						oMessage.addRecipient x,x
					End if
				next
				
				'Add back the modified CC and TO headers
				oMessage.HeaderValue("To") = newToHeader
				oMessage.HeaderValue("CC") = newCCHeader
				
				'Save the message
				oMessage.Save
				
				'Send the reject email stating that the mail was blocked for some recipients
		
				Dim RejectMessage 
				
				RejectMessage  = "Your mail was not sent to the following recipients: " & ResultString  & " as your address (" & oClient.UserName & ") is not allowed to send mails to them. " & nl & "Please send a reply to this message if you need these IDs to be unblocked."
				SendMessage MailGuardID,oMessage.FromAddress,"MailGuard blocked - " & oMessage.Subject, RejectMessage
				
				'Result.Message = "You cannot send mails to the following IDs as they are blocked by MailGuard: " & ResultString
				Result.Value = 0
		End Select
		
	Else 'blank, everything ok
		If LogLevel>1 THEN WriteLog MGLogPath,  "MGResult:Ok"
		Result.Value = 0
	End If

Else
	Result.Message = "You need to be authenticated to send email through this system."
	Result.value = 2
End If 'Run only for authenticated outbound email
End Sub

Function truncate_one(s)
  If Right(s, 1) = "," Then 
    truncate_one = Left(s, Len(s) - 1) 
  Else 
    truncate_one = s
  End If
End Function

Function ErrorHandler(msg)
	If Err.Number <> 0 Then
		WriteLog MGLogPath, msg
		Err.Clear
	End If
End Function

Function WriteLog(filename,StrTxt)
      Dim fso2
      Set fso2 = CreateObject("Scripting.FileSystemObject")
      fso2.OpenTextFile(filename, 8, True).WriteLine(now & "-" & StrTxt)
	  Set fso2 = Nothing
End Function

Function SendMessage(sender,recipient,subject,body)
	dim oMessage
	Set oMessage=CreateObject("hMailServer.Message")
	oMessage.From=sender
	oMessage.FromAddress=sender
	oMessage.AddRecipient recipient, recipient
	oMessage.Subject = subject
	oMessage.Body = body
	oMessage.Save 
	'oMessage = nothing
End Function

Sub OnDeliveryFailed(oMessage, sRecipient, sErrorMessage)
	'Dim msg
	'msg = now & vbTab & oMessage.FromAddress & vbTab & oMessage.Subject & vbtab & sRecipient & vbTab & sErrorMessage
	'WriteLog FailedDeliveryLogPath, msg
	'Call the next script to log the entry into the database
	If lcase(oMessage.from) <> "hmailserver" Then 'The CreateFailedDeliveryLogEntry sub sends a mail informing non-delivery. If that fails, then we have looped messages. Stop them
		CreateFailedDeliveryLogEntry oMessage, sRecipient, sErrorMessage
	End If
End Sub

Sub CreateFailedDeliveryLogEntry(oMessage, sRecipient, sErrorMessage)
	Dim nl
	nl = Chr(13) & Chr(10)

   Dim obDatabase
   Set obDatabase = GetDatabaseObject
   Dim sFrom, sFilename, sSentOn, sSubject
   
   sFrom = oMessage.From
   sFilename = oMessage.Filename
   sSentOn = GetTimestamp(obDatabase)
   sSubject = oMessage.Subject
   
   sFrom = Mid(sFrom, 1, 255)
   sFilename = Mid(sFilename, 1, 255)
   sSubject = Mid(sSubject, 1, 255)
   
   if (Len(sErrorMessage) > 1000000) Then
      sErrorMessage = Mid(sErrorMessage, 1, 1000000)
   End If
   
   sFrom = Escape(obDatabase, sFrom)
   sFilename = Escape(obDatabase, sFilename)
   sSubject = Escape(obDatabase, sSubject)
   sErrorMessage = Escape(obDatabase, sErrorMessage)
   
   Dim sSQL
   sSQL = "insert into hm_Faileddeliverylog (`from`, `to`, `filename`, `senton`, `subject`, `errormsg`) values ('" & sFrom & "', '" & sRecipient & "', '" & sFilename & "', " & sSentOn & ", '" & sSubject & "', '" & sErrorMessage & "')"
   'SendMessage "hMailServer",oMessage.From,"Mail not sent - " & sSubject, "Your message could not be sent by our email server to " & sRecipient & nl & nl & "Error from remote server: " & sErrorMessage
   'dim iID
   'iID = obDatabase.ExecuteSQLWithReturn(sSQL)
   Call obDatabase.ExecuteSQL(sSQL)
   
    
   If Err.Number <> 0 Then
		WriteLog FailedDeliveryLogPath, now & Err.message
		Err.Clear
	End If
End Sub

Function GetDatabaseObject()
	Dim obApp
	Set obApp = CreateObject("hMailServer.Application")
	Call obApp.Authenticate(setting_username, setting_password)
    Set GetDatabaseObject = obApp.Database
End Function

Function Escape(obDatabase, value)

   value = Replace(value, "'", "''")

   Select Case obDatabase.DatabaseType
      Case 1: ' MySQL
         value = Replace(value, "\", "\\")
      Case 3: ' PGSQL
         value = Replace(value, "\", "\\")
   End Select
   
   Escape = value
End Function

Function GetTimestamp(obDatabase)
   Select Case obDatabase.DatabaseType
      Case 1: ' MySQL
         GetTimestamp = "NOW()"
      Case 2: ' MSSQL
         GetTimestamp = "GETDATE()"
      Case 3: ' PGSQL
         GetTimestamp = "current_timestamp"
      Case 4: ' SQL CE
         GetTimestamp = "GETDATE()"
   End Select
End Function

'   Sub OnDeliveryStart(oMessage)
'   End Sub

'   Sub OnDeliverMessage(oMessage)
'   End Sub

'   Sub OnBackupFailed(sReason)
'   End Sub

'   Sub OnBackupCompleted()
'   End Sub

'   Sub OnError(iSeverity, iCode, sSource, sDescription)
'   End Sub

'   Sub OnExternalAccountDownload(oMessage, sRemoteUID)
'   End Sub