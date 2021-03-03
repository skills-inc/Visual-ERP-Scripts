#=============================================
# - Script Name: SendEmailFromPowerShell-Tasks.ps1
#-- Author:		 Ed Hammond - Skills Inc
#-- Create date: 3/20/2020
#-- Edited by:   Ed Hammond
#-- Edit date:   3/3/2021
#-- Description: Query the TASK table in SQL and run send an email for every new ECN, every 15 min.  Run this from Task Scheduler on your a server with SQL Tools like invoke-sqlcmd installed.
#-- Description: Change the variables at the top and then change the path to your logo and any supporting QMS/Help documents at the bottom in the HTML
#-- Rev A: Published to GIT
#-- =============================================
$SQLServer = "Your-SQL-SERVER"
$db1 = "YourVisualDB"
$SMTP_RELAY_SERVER = "smtp.yourserver.com"
$Email_From = "relay@myemaildomain.com"
$Last15Min = (get-date).AddMinutes(-15)
$ECNquery = "SELECT e.FIRST_NAME,e.EMAIL_ADDR,t.EC_ID as ECID,t.TYPE,t.TASK_NO,t.SEQ_NO
,CASE WHEN t.TYPE = 'ECN' AND t.SUB_TYPE = 'A' THEN ISNULL(CONVERT(VARCHAR(30), s.ECN_AUTH_LABEL) , 'Authorization')
WHEN t.TYPE = 'ECN' AND t.SUB_TYPE = 'AT' THEN 'Assigned To'
WHEN t.TYPE = 'ECN' AND t.SUB_TYPE = 'I' THEN ISNULL(CONVERT(VARCHAR(30),s.ECN_IMPL_LABEL), 'Implementation')
WHEN t.TYPE = 'ECN' AND t.SUB_TYPE = 'AP' THEN ISNULL(s.ECN_APRV_LABEL, 'Approval')
WHEN t.TYPE = 'ECN' AND t.SUB_TYPE = 'D' THEN ISNULL(CONVERT(VARCHAR(30),s.ECN_DIST_LABEL), 'Distribution')
END as SubType
,CAST(CAST(ecr.BITS AS varbinary(Max)) AS nvarchar(Max)) as Request
,t.CREATE_DATE
FROM TASK t
JOIN EMPLOYEE e 
ON t.USER_ID = e.USER_ID
JOIN USER_SITE us
ON us.USER_ID = t.USER_ID
JOIN SITE s
ON us.SITE_ID = s.ID
JOIN ACCOUNTING_ENTITY ae
ON s.ENTITY_ID = ae.ID
inner Join EC_REQUEST ecr
ON t.EC_ID = ecr.EC_ID
WHERE (t.TYPE = 'ECN') AND (t.CREATE_DATE > '$Last15Min') AND (t.COMPLETED_DATE IS NULL) AND (t.STATUS = 'P')"



$Alerts = invoke-sqlcmd -ServerInstance $SQLServer -Database $db1 -Query $ECNquery

Foreach ($t in $alerts)
{
$ECID1 = $t.ECID
$EDTYPE = $t.TYPE
$ECTASK = $t.TASK_NO
$ECSEQ = $t.SEQ_NO
$ECDATE = $t.CREATE_DATE
$ECREQUEST = $t.Request
$ECSUB = $t.SubType


$Body1 = "<!DOCTYPE html>
<html>
<head>
<style>body {font-family: arial, sans-serif;font-size: 12pt;font-weight: normal;} 
div.socialmedia {margin-top: 10px;} div.socialmedia a {padding-left: 6px;padding-right: 6px;} p.details {color: black;padding-top: 8px;} 
p.download {float: right;margin-top: 30px;margin-bottom: 0px;text-align: right;} p.download a {text-decoration: none;}' + CHAR(10)+CHAR(13) 
hr {clear: both;height: 1px;background-color: #9a9a9a;border-color: #9a9a9a;color: #9a9a9a;border: 0px;margin: 0px;margin-bottom: 8px;} a {color: #0321ff;}
a img {border: none;} img {outline: none;text-decoration: none;max-width: 100%;-ms-interpolation-mode: bicubic;}
h2 {font-size: 18pt;color: #1e598f;margin-bottom: 0px;font-weight: normal;text-align: right;} h3 {font-size: 12pt;font-weight: bold;margin-bottom: 10px;} 
table {border: none;cellpadding: 0px;margin-left: 50px;width: 90%;} td {vertical-align: bottom;} td span {margin-left: 15px;margin-right: 15px;} table.header {margin-top: 30px;}</style></head><body> 
<table class=""header""><tr><td><img style=""margin-bottom: 15px;"" src=""https://www.skillsinc.com/wp-content/themes/skills/images/skills-incorporated-logo.png"" border=""0""></a></td><td><h2>
ECN ID # $ECID1 </div>
</h2></td></tr><tr><td colspan=""2""><hr></td></tr></table><table><tr><td><h3>You have been assigned a task in Visual</h3>
<div><b>Task:</b> $EDTYPE/$ECTASK.$ECSEQ</div>
<div><b>Created On:</b> $ECDATE </div>
<div><b>SUB-TYPE:</b> $ECSUB </div>
<div><b>Request:</b><br>
<table border=""1"" cellpadding=""4""><tr><td>
<pre>$ECREQUEST </pre></div><br>
</td></tr></table>
</td></tr></table><div><table><tr><td colspan=""3""><hr></td></tr></table><table><tr><td colspan=""3"">
<p class=""details"">This is an automated email please do not reply to this message.</p>
<p>Open Visual and choose Eng/MFG + Task Maintenance to see your assigned tasks.</p>
<br>
<h3>The following documents in the QMS may help if you have questions about using ECNs.</h3>
<p><a href=""https://mywebsitegoeshere.com/doc.doc"">Document Change Notice and Review Process</a></p>
<p><a href=""https://mywebsitegoeshere.com/doc.doc"">ECNs for Engineering Revision Rolls Standard Work</a></p>
</td></tr></table></div>"
Send-MailMessage -To $t.EMAIL_ADDR -From $Email_From -BodyAsHtml $Body1 -Subject "A Visual ECN Task has been assigned to you" -SmtpServer $SMTP_RELAY_SERVER
}

