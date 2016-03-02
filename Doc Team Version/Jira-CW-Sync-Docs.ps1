####################################################################
#Jira Functions

Function Get-Issue
{
  	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[INT]$IssueID
	)

    $RestApiURI = $JiraServerRoot + "rest/api/2/issue/$IssueID"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -method Get

    If($JSONResponse.fields)
    {
        Return $JSONResponse.fields
    }

    Else
    {
        Return $False
    }
}

Function Get-Worklogs
{
	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[String]$dateFrom,
		[Parameter(Mandatory = $true,Position = 1)]
		[String]$dateTo,
		[Parameter(Mandatory = $true,Position = 2)]
		[String]$username
	)


    $RestApiURI = $JiraServerRoot + "rest/tempo-timesheets/3/worklogs/"

        $Body = @{
    "dateFrom" = "$dateFrom"
    "dateTo" = "$dateTo"
    "username" = "$username"
    }

    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -Body $Body -ContentType application/json -method get

    If($JSONResponse)
    {
        Return $JSONResponse
    }

    Else
    {
        Return $False
    }
}

Function Get-JiraUserInfo
{
    param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[String]$Name
	)


    $Body = @{
    "username" = "$Name"
    }

    $RestApiURI = $JiraServerRoot + "rest/api/2/user"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -body $Body -ContentType application/json -method Get

    If($JSONResponse.displayname)
    {
        Return $JSONResponse.displayname
    }

    Else
    {
        Return $False
    }

}

Function Set-JiraCreds
{
    $BinaryString = [System.Runtime.InteropServices.marshal]::StringToBSTR($($Jirainfo.password))
    $JPassword = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BinaryString)
    $JLogin = $Jirainfo.user
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes("$jLogin`:$jPassword")
    $JiraCredentials = [System.Convert]::ToBase64String($bytes)
    Return $JiraCredentials
}

Function Edit-JiraIssue
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[String]$IssueID,
		[Parameter(Mandatory = $True,Position = 1)]
		[String]$CWTicketID
	)

$Body= @"
{
"fields":
	{
	"customfield_10313" : "$CWTicketID"
	}
}
"@

    $RestApiURI = $JiraServerRoot + "rest/api/2/issue/$IssueID"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Body $Body -method Put
}

####################################################################
#ConnectWise Retrieval Functions

Function Get-CWTicket
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
     }
    
    Process
    {   
        Try
        {   
            $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
        }

        Catch
        {
            $ErrorMessage = $_.exception.message
        }

    }
    
    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Get-CWTimeEntries
{
    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[INT]$TicketID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID/timeentries"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
     }
    
    Process
    {   
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
    }

    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Get-TimeEntryDetails
{
    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[INT]$TimeEntryID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/Time/Entries/$TimeEntryID"
    [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
    [string]$ContentType = "application/json"

    $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }
    
    Process
    {   
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers  -ContentType $ContentType -Method Get
    }

    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Get-CWMember
{
    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$EmailAddress
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/system/members"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"

        $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }

    Process
    {
        $Body = @{
    "conditions" = "emailaddress = '$EmailAddress'"
    
    }
       
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
    }

    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Get-CWContact
{
    [cmdletbinding()]

    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$First,
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$Last,
    	[Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
		[String]$CompanyID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/company/contacts"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"

        $Headers=@{
        'X-cw-overridessl' = "True"
        "Authorization"="Basic $encodedAuth"
        }
    }

    Process
    {
        $Body = @{
    "conditions" = "firstname = '$First' AND lastname = '$Last' AND company/id =$CompanyID"
    }
      
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
    }

    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse.id
        }

        Else
        {
            Return $False
        }
    }
}

####################################################################
#ConnectWise Post/Edit Functions

function New-CWTicket 
{
    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Ticket,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$BoardName,
        [Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$Key
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
    }

    Process
    {
        Write-Log -Message "[DEBUG] Initial Summary = $($Ticket.summary)"

        If($($Ticket.summary) -eq $Null -or $($Ticket.summary) -eq "")
        {
            Write-Log "$($Ticket | Format-Table | Out-String)"
            Write-log "$($Error | Out-string)"
        }

        #Making sure the summary field is formatted properly
        ###############################################################
        If([INT]$($Ticket.Summary.length) -gt 90)
        {
            $Summary = $($Ticket.Summary.substring(0,75))
        }

        $Summary = Format-SanitizedString -InputString $($Ticket.Summary)
        $Summary = $Summary.Replace('"', "'")
        ###############################################################
    
        If(!$($Ticket.assignee.emailaddress))
        {
            Write-Log "WARNING!! No Assignee was present on this Issue in JIRA. $DefaultContactEmail has been assigned."
            $UserInfo = Get-ProperUserInfo -JiraEmail $DefaultContactEmail
        }

        Else
        {
            $UserInfo = Get-ProperUserInfo -JiraEmail $($Ticket.assignee.emailaddress)
        }

        Write-log "[DEBUG]Summary = [JIRA][$Key] - $Summary"

        $Body= @"
{
    "summary"   :    "[JIRA][$Key] - $Summary",
    "board"     :    {"name": "$BoardName"},
    "status"    :    {"name": "New"},
    "company"   :    {"id": "$($UserInfo.CompanyID)"},
    "contact"   :    {"id": "$($UserInfo.ContactID)"}
}
"@
        $Headers=@{
'X-cw-overridessl' = "True"
"Authorization"="Basic $encodedAuth"
}
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Method Post -Body $Body
    }

    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Close-CWTicket
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"

        $Headers=@{
            'X-cw-overridessl' = "True"
            "Authorization"="Basic $encodedAuth"
            }

        $Body= @"
        [
        {
            "op" : "replace", "path": "/status/id", "value": "$Global:ClosedStatusValue"
        }
        ]
"@
     }
    
    Process
    {      
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -Body $Body -ContentType $ContentType -Method Patch
    }
    
    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

Function Open-CWTicket
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/service/tickets/$ticketID"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"

        $Headers=@{
            'X-cw-overridessl' = "True"
            "Authorization"="Basic $encodedAuth"
            }

        $Body= @"
        [
        {
            "op" : "replace", "path": "/status/id", "value": "$Global:OpenStatusValue"
        }
        ]
"@
     }
    
    Process
    {      
        $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -Body $Body -ContentType $ContentType -Method Patch
    }
    
    End
    {
        If($JSONResponse)
        {
            Return $JSONResponse
        }

        Else
        {
            Return $False
        }
    }
}

function New-CWTimeEntry
{

    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [PSCustomObject]$WorkLog,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Int]$CWTicketID
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "v4_6_Release/apis/3.0/time/entries"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
    }

    Process
    {
        #Date Magic
        $StartedUniversal = (get-date ($WorkLog.datestarted)).ToUniversalTime()
        [string]$Ended = ($StartedUniversal).AddSeconds($Worklog.timeSpentSeconds)
        [String]$Created = Get-Date ($StartedUniversal) -format "yyyy-MM-ddTHH:mm:ssZ"
        [String]$Ended = Get-Date ($Ended) -format "yyyy-MM-ddTHH:mm:ssZ"

        $MemberInfo = Get-ProperUserInfo -JiraEmail "$($Worklog.author.name)@labtechsoftware.com" -MemberCheck '1'
        $SanitizedComment = Format-sanitizedstring -InputString $($WorkLog.comment)

        $Body= @"
{
    "chargeToType"   : "ServiceTicket",
    "chargeToId"     : "$CWTicketID",
    "timeStart"      : "$Created",
    "timeend"        : "$Ended",
    "internalnotes"  : "[JiraID!!$($Worklog.id)!!] $($SanitizedComment)",
    "company"        : {"id": "$($Memberinfo.CompanyID)"},
    "member"         : {"id": "$($Memberinfo.MemberID)"},
    "billableOption" : "DoNotBill"
    
}
"@
        $Headers=@{
'X-cw-overridessl' = "True"
"Authorization"="Basic $encodedAuth"
}

        Try 
        {

            $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Method Post -Body $Body
        
        }
    
        Catch [Exception]
        {
            $ErrorMessage = $_.exception.message
            Return 'Something went wrong.'
        }
    }

    End
    {
        Return $JSONResponse
    }
}

function Invoke-TicketProcess
{
    [cmdletbinding()]

    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [PSObject]$Issue,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$Boardname,
        [Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$Key,
        [Parameter(Mandatory = $true,Position = 3,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Worklog
    )

    Process
    {
        If ($Issue.customfield_10313 -eq $Null)
        {
            Write-Log "No CW Ticket # found for this Jira issue."
            $Ticket = New-CWTicket -Ticket $Issue -BoardName "$BoardName" -Key $Key

            If($Ticket -eq $False)
            {
                Write-Log "Failed to Create CW Ticket for Jira Issue $($Issue.id)"
            }

            Write-Log "CW Ticket #$($ticket.id) created."
            Edit-JiraIssue -IssueID "$($Worklog.issue.id)" -CWTicketID "$($Ticket.id)"
            $issue.customfield_10313 = $Ticket.id 
            Write-Log "CW Ticket #$($ticket.id) mapped in JIRA."
        }

        Else
        {
            $CurrentTicket = Get-CWticket -TicketID $Issue.customfield_10313

            If($CurrentTicket.id -ne $Issue.customfield_10313)
            {
                Write-Log "CW Ticket ID #$($Issue.customfield_10313) does not exist."
                $Ticket = New-CWTicket -Ticket $Issue -BoardName "$BoardName" -Key $Key
                Write-Log "CW Ticket #$($ticket.id) created."
                Edit-JiraIssue -IssueID "$($Worklog.issue.id)" -CWTicketID "$($Ticket.id)"
                $issue.customfield_10313 = $Ticket.id 
                Write-Log "CW Ticket #$($ticket.id) mapped in JIRA." 
            }

            Else
            {
                Write-Log "CW Ticket #$($Issue.customfield_10313) is already correctly mapped."
            }
        }
    }
}

function Invoke-WorklogProcess
{
    param
    (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Issue,
        [Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Object]$Worklog,
        [Parameter(Mandatory = $true,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String]$ClosedStatus
    )

    Process
    {
        [ARRAY]$NewTimeEntries = @()

        $Ticket = Get-cwticket -TicketID $($Issue.customfield_10313)

        If($Ticket.status.name -eq $ClosedStatus)
        {
            $Closed = $True
        }

        [Array]$TimeEntryIDs = (Get-CWTimeEntries -TicketID "$($Issue.customfield_10313)").id

        Foreach($TimeEntry in $TimeEntryIDs)
        {
            [Array]$TEDetails += Get-TimeEntryDetails -TimeEntryID $TimeEntry
        }
  

        Foreach($Detail in $TEDetails)
        {
                [INT]$Present = 0
                $ErrorActionPreference = 'SilentlyContinue'
                $RegCheck = ([regex]::matches($Detail.internalnotes, "(?:\[JiraID!!)(.*)(?:!!)")).groups[1].value   
                $ErrorActionPreference = 'Continue'

                If($($Worklog.id) -eq $RegCheck)
                {
                    [INT]$Present = 1
                    break;
                }  
        }

        If($Present -ne 1)
        {
            If($Closed)
            {
                $OpenIt = Open-CWTicket -TicketID $($Issue.customfield_10313)
                    
                If ($OpenIt.status.name -eq $Global:ReopenStatusName)
                {
                    Write-Log "CW Ticket #$($Issue.customfield_10313) has been re-opened for posting time."
                }

                Else
                {
                    Write-Log "Failed to re-open CW Ticket #$($Issue.customfield_10313)"
                    break;
                }
            }
            
            $TimeEntry = New-CWTimeEntry -WorkLog $Worklog -CWTicketID "$($Issue.customfield_10313)"

                If($TimeEntry -eq 'Something went wrong.')
                {
                    Write-Log "Jira Time Entry ID #$($Worklog.id) occurred in a previous time period."
                }

                Else
                {
                    Write-Log "New Time Entry Created."
                    Write-Log "Jira Time Entry ID #$($Worklog.id) | Time Logged = $($Worklog.timespentseconds/60/60) Hours"
                }
     
        }

        Else
        {
            Write-Log "Jira Time Entry ID #$($Worklog.id) is already logged." 
        }

}
}

####################################################################
#Data Manipulation Functions

Function Format-SanitizedString
{
	[CmdLetBinding()]
	Param
		(
		[Parameter(Mandatory = $False)]
		[String]$InputString
	)
	
	$SanitizedString = "";
	If ($InputString -ne $null -and $InputString.Trim().Length -gt 0)
	{
		$SanitizedString = $InputString.Trim();
		$SanitizedString = $SanitizedString.Replace("\", "\\");
		$SanitizedString = $SanitizedString.Replace("'", "\'");
		$SanitizedString = $SanitizedString.Replace("`"", "\`"");
	}
	
	Return $SanitizedString
}

Function Get-ProperUserInfo
{

    param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[String]$JiraEmail,
		[Parameter(Mandatory = $False,Position = 1)]
		[String]$MemberCheck
	)

    $JiraUser = ($JiraEmail.split('@'))[0]
    $JiraName = Get-JiraUserInfo -Name $JiraUser

    If($JiraName -eq $False)
    {
        $ContactID = '255093'
        $Firstname = 'Phillip'
        $Lastname = 'Marshall'
    }

    Else
    {
       $SplitName = $JiraName.split(' ')
       $FirstName = $Splitname[0]
       $LastName  = $SplitName[1] 
    }

    If($JiraEmail -like '*@Labtechsoftware*')
    {
        $CompanyID = '49804'
    }

    Else
    {
        $CompanyID = '250'
    }

    $ContactID = Get-CWContact -First $FirstName -Last $LastName -CompanyID $CompanyID

    If($ContactID -eq $False)
    {
        Write-Log "Unable to retrieve a contact for Firstname: $Firstname Lastname: $Lastname CompanyID: $CompanyID"
        $Contactid = ''
    }

    If($MemberCheck)
    {
        $MemberInfo = Get-CWMember -EmailAddress $JiraEmail
        $MemberID = $MemberInfo.id

        If(!$MemberID)
        {
            $MemberID = ''
        }
    }

    Else
    {
        $MemberID = ''
    }

    

    $UserInfo += New-Object PSObject -Property @{

        First     = $FirstName;
        Last      = $LastName;
        Email     = $JiraEmail;
        CompanyID = $CompanyID;
        ContactID = $ContactID;
        MemberID  = $MemberID;
    }

    Return $UserInfo
}

Function Get-Week
{ 
    param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[Datetime]$Weekday
	)

    $DoW = $Weekday.DayOfWeek

    switch ($DoW)
        {
            “Saturday”  {[int]$Offset = -6} 
            “Sunday”    {[int]$Offset = -1} 
            “Monday”    {[int]$Offset = -3} 
            “Tuesday”   {[int]$Offset = -5} 
            “Wednesday” {[int]$Offset = -7} 
            “Thursday”  {[int]$Offset = -9} 
            “Friday”    {[int]$Offset = -11} 
        }

$DaysSince = $Weekday.DayOfWeek.value__ + $Offset
$WeekBegin = $Weekday.AddDays($DaysSince)
$StartOfWeek = $WeekBegin.addhours(- $($WeekBegin.Hour)).addminutes(- $($WeekBegin.minute)).addseconds(- $($WeekBegin.second))
$EndOfWeek = ((($StartOfWeek.AddDays(6)).addhours(23)).addminutes(59)).addseconds(59)

    $CurrentWeek += New-Object PSObject -Property @{
        Start = $StartOfWeek
        End = $EndOfWeek
        Span = (New-TimeSpan -Start $StartOfWeek -End $EndOfWeek)
        CurDay = $Weekday
        }

Return $CurrentWeek
}

Function Write-Log
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send timestamped messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER StrMessage
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -StrMessage 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    
	add-content -path $LogFilePath -value ($Message)
    Write-Output $Message
}

####################################################################
#Variable Declarations
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
[Array]$arrUsernames = @('cvalentine','sbakan','bwhitmire')
[String]$CWServerRoot = "https://cw.connectwise.net/"
[String]$JiraServerRoot = "https://jira-dev.labtechsoftware.com/"
[String]$ImpersonationMember = 'jira'
[String]$DefaultContactEmail = 'bwhitmire@labtechsoftware.com'
[String]$Boardname = 'LT-Documentation'
[String]$ClosedStatus = '>Complete'
[String]$Global:OpenStatusValue = '6952'
[String]$Global:ClosedStatusValue = '6951'
[String]$Global:ReopenStatusName = 'New'
[Int]$MaxResults = '250'
[String]$LogFilePath = "C:\Scheduled Tasks\Logs\Jira-CW-Doc-Team.txt"

Remove-Item $LogFilePath -Force -ErrorAction 'SilentlyContinue'

#Jira Credentials
$Global:JiraInfo = New-Object PSObject -Property @{
User = 'cwintegration'
Password = '@#WE23we4'
}
$JiraCredentials = Set-JiraCreds

#CW Credentials
$Global:CWInfo = New-Object PSObject -Property @{
Company = 'connectwise'
PublicKey = '4hc35v3aNRTjib9W'
PrivateKey = 'yLubF4Kfz4gWKBzU'
}
[string]$Authstring  = $CWInfo.company + '+' + $CWInfo.publickey + ':' + $CWInfo.privatekey
$encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

$StartRange = (get-date).AddDays(-7)
$EndRange = (Get-Date)
[String]$WeekStart = "$($StartRange.Year)`-$($StartRange.month)`-$($StartRange.day)"
[String]$WeekEnd = "$($EndRange.Year)`-$($EndRange.month)`-$($EndRange.day)"

Write-Log "This Week is $Weekstart - $Weekend"

Foreach($User in $arrUsernames)
{
    Write-Log "-----------------------------------------------"
    Write-Log "Beginning Processing User: $User"
    $UserWorklogs = Get-Worklogs -username $User -dateFrom $WeekStart -dateTo $WeekEnd

    If($UserWorklogs -eq $False)
    {
        Write-Log "No Time Entries for User: $User"
    }

    Else
    {
        Write-Log "Time Entries Found: $(($Userworklogs | measure-object).count)"
        [INT]$Counter = '1'
        Foreach($Worklog in $UserWorklogs)
        {
            Write-Log "-----------------------------------------------"
            Write-Log "Processing $Counter of $(($Userworklogs | measure-object).count)"
            $Issue = Get-Issue -IssueID "$($worklog.issue.id)"

            If ($Issue -eq $False -or $Issue -eq $Null)
            {
                Write-log "This damn issue didnt exist somehow."
            }

            Invoke-TicketProcess -Issue $Issue -Boardname $Boardname -Key $($Worklog.issue.key) -Worklog $Worklog
            Invoke-WorklogProcess -Issue $Issue -Worklog $Worklog -ClosedStatus $ClosedStatus
            
            #Close the ticket in CW if its closed in Jira
            If($Issue.status.name -eq 'Closed')
            {
                Write-Log "Jira Issue is closed."
                $ISClosed = Get-cwticket -TicketID $($Issue.customfield_10313)
            
                If($ISClosed.status.name -eq $ClosedStatus)
                {
                    Write-Log "CW Ticket #$($Issue.customfield_10313) is closed."
                }
            
                Else
                {
                    $CloseIt = Close-CWTicket -TicketID $($IsClosed.id)

                    If ($CloseIt.status.name -eq $ClosedStatus)
                    {
                        Write-Log "CW Ticket #$($IsClosed.id) has been closed."
                    }

                    Else
                    {
                        Write-Log "Failed to close CW Ticket #$($IsClosed.id)"
                    }
                }
            }

            $Counter++
        }

      }
}