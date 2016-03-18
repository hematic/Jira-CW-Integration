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

    Try
    {
        $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -method Get
    }

    Catch
    {
        Output-Exception
        Return "TIMEOUT"
    }


    If($JSONResponse.fields)
    {
        Return $JSONResponse.fields
    }

    Else
    {
        Return "UNKNOWN ERROR"
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

    Try
    {
        $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -Body $Body -ContentType application/json -method get
    }

    Catch
    {
        Output-Exception
        Return "Exception Caught"       
    }

    If($JSONResponse.id)
    {
        Return $JSONResponse
    }

    Else
    {
        Return "No Worklogs"
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

    Try
    {
        $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -body $Body -ContentType application/json -method Get
    }

    Catch
    {
        Output-Exception
    }

    If($JSONResponse)
    {
        Return $JSONResponse
    }

    Else
    {
        Return "No User"
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

    $RestApiURI = $JiraServerRoot + "rest/api/latest/issue/$IssueID"
    
    Try
    {    
        $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Body $Body -method Put
        $JsonResponse2 = Get-Issue -IssueID $IssueID
    }

    Catch
    {
        Output-Exception
    }

    Write-Verbose "$($JsonResponse2.customfield_10313)"

    If($($JsonResponse2.customfield_10313) -eq $CWTicketID)
    {
        Return "Success"    
    }

    Else
    {
        Return "Failed to Set CustomField_10313"
    }
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
		[String]$TicketID
    )

    Begin
    {
    [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/service/tickets/$ticketID"
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
            Output-Exception
        }

    }
    
    End
    {
        If($JSONResponse.id -eq $TicketID)
        {
            Return $JSONResponse
        }

        Else
        {
            Return 'Bad Request'
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
    [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/service/tickets/$ticketID/timeentries"
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
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse
        }

        Else
        {
            Return "NO TIME ENTRIES"
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
    [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/Time/Entries/$TimeEntryID"
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
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse
        }

        Else
        {
            Return "NO TIME ENTRY DETAILS"
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
        [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/system/members"
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
        
        Try
        {
            $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
        }

        Catch
        {
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse
        }

        Else
        {
            Return "NO CW MEMBER DATA"
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
        [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/company/contacts"
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
        
        Try
        {
            $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -ContentType $ContentType -Body $Body -Method Get
        }

        Catch
        {
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse.id
        }

        Else
        {
            Return "NO CW CONTACT DATA"
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
        [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/service/tickets"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"
    }

    Process
    {
        If($($Ticket.summary) -eq $Null -or $($Ticket.summary) -eq "")
        {
            Write-Log "$($Ticket | Format-List | Out-String)"
            Return "CW TICKET CREATION FAILED"
        }

        #Making sure the summary field is formatted properly
        ###############################################################
        If([INT]$($Ticket.Summary.length) -gt 70)
        {
            $Summary = $($Ticket.Summary.substring(0,69))
        }
        Else
        {
            $Summary = $($Ticket.summary)
        }

        $Summary = Format-SanitizedString -InputString $Summary
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

        $Body= @"
{
    "summary"   :    "[JIRA][$Key] - $Summary",
    "board"     :    {"name": "$BoardName"},
    "status"    :    {"name": "$Global:NewTicketStatusName"},
    "company"   :    {"id": "$($UserInfo.CompanyID)"},
    "contact"   :    {"id": "$($UserInfo.ContactID)"}
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

        Catch
        {
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse
        }

        Else
        {
            Return "CW TICKET CREATION FAILED"
        }
    }
}

Function Change-CWTicketStatus
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID,
    	[Parameter(Mandatory = $false,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[Bool]$Open,
    	[Parameter(Mandatory = $false,Position = 2,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[Bool]$Close
    )

    Begin
    {
        [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/service/tickets/$ticketID"
        [string]$Accept      = "application/vnd.connectwise.com+json; version=v2015_3"
        [string]$ContentType = "application/json"

        $Headers=@{
            'X-cw-overridessl' = "True"
            "Authorization"="Basic $encodedAuth"
            }

        If($Open -eq $True)
        {
            $Body= @"
            [
            {
                "op" : "replace", "path": "/status/id", "value": "$Global:OpenStatusValue"
            }
            ]
"@
        }
        
        ElseIf($Close -eq $True)
        {
            $Body= @"
            [
            {
                "op" : "replace", "path": "/status/id", "value": "$Global:ClosedStatusValue"
            }
            ]
"@
        }

        Else
        {
            Write-Log "Ambiguous parameters passed to function. Please specify Open or Close"
            Return "Ambiguous Parameters"
        }


     }
    
    Process
    {   
        Try
        {   
            $JSONResponse = Invoke-RestMethod -URI $BaseURI -Headers $Headers -Body $Body -ContentType $ContentType -Method Patch
        }

        Catch
        {
            Output-Exception         
        }
    }
    
    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne '')
        {
            Return $JSONResponse
        }

        Else
        {
            Return "Status Change Failed"
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
        [string]$BaseUri     = "$CWServerRoot" + "$Codebase" + "apis/3.0/time/entries"
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
    
        Catch
        {
            Output-Exception
        }
    }

    End
    {
        If($JSONResponse -ne $Null -and $JSONResponse -ne "")
        {
            Return $JSONResponse
        }

        Else
        {
            Return "CW TIME ENTRY CREATION FAILED"
        }
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
        If ($($Issue.customfield_10313) -eq $Null -or $($Issue.customfield_10313) -eq "None")
        {
            Write-Log "No CW Ticket # found for this Jira issue."
            $Ticket = New-CWTicket -Ticket $Issue -BoardName "$BoardName" -Key $Key

            If($Ticket -eq "CW TICKET CREATION FAILED")
            {
                Write-Log "Failed to Create CW Ticket for Jira Issue $($Issue.id)"
                Return "Process Failure"
            }

            Write-Log "CW Ticket #$($ticket.id) created."
            $EditResults = Edit-JiraIssue -IssueID "$($Worklog.issue.id)" -CWTicketID "$($Ticket.id)"

            If($EditResults -eq "Failed to Set CustomField_10313")
            {
                Write-Log "[*Error*]Failed to set customfield_10313 for Jira Issue $($Issue.id). Value should have been $($Ticket.id)."
                Return "Process Failure"
            }

            Else
            {
                $issue.customfield_10313 = $Ticket.id 
                Write-Log "CW Ticket #$($ticket.id) mapped in JIRA."
                Return "Process Success"
            }


        }

        Else
        {
            If($($Issue.customfield_10313) -notmatch "^[\d\.]+$")
            {
                Write-Log @"
[*ERROR*] : CW Ticket Field contains a value that is not all numeric.
Jira Issue: $($Worklog.issue.key)
Customfield_10313 Value: $($Issue.customfield_10313)
"@ 
                Return "BAD CW TICKET VALUE"
            }

            $CurrentTicket = Get-CWticket -TicketID $($Issue.customfield_10313)

            If($CurrentTicket -eq "Bad-Request")
            {
                Write-log "[*Error*]Unable to retrieve information for CW Ticket $($Issue.customfield_10313)"
                Return "Process Failure"
            }

            ElseIf($CurrentTicket.id -ne $Issue.customfield_10313)
            {
                Write-Log "CW Ticket ID #$($Issue.customfield_10313) does not exist."
                $Ticket = New-CWTicket -Ticket $Issue -BoardName "$BoardName" -Key $Key

                If($Ticket -eq "CW TICKET CREATION FAILED")
                {
                    Write-Log "Failed to Create CW Ticket for Jira Issue $($Issue.id)"
                    Return "Process Failure"
                }

                Write-Log "CW Ticket #$($ticket.id) created."

                $EditResults = Edit-JiraIssue -IssueID "$($Worklog.issue.id)" -CWTicketID "$($Ticket.id)"

                If($EditResults -eq "Failed to Set CustomField_10313")
                {
                    Write-Log "[*Error*]Failed to set customfield_10313 for Jira Issue $($Issue.id). Value should have been $($Ticket.id)."
                    Return "Process Failure"
                }

                Else
                {
                    $issue.customfield_10313 = $Ticket.id 
                    Write-Log "CW Ticket #$($ticket.id) mapped in JIRA."
                    Return "Process Success"
                }

            }

            Else
            {
                Write-Log "CW Ticket #$($Issue.customfield_10313) is already correctly mapped."
                Return "Process Success"
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

        If($Ticket -eq "Bad-Request")
        {
            Write-log "[*Error*]Unable to retrieve information for CW Ticket $($Issue.customfield_10313)"
            continue;
        }

        #Check if the starting status of the CW Ticket is closed and set a flag for later so we set it back.
        If($Ticket.status.name -eq $ClosedStatus)
        {
            $Closed = $True
        }

        #Check the result of the Get-CWTimeEntries function.
        $TimeEntryResponse = Get-CWTimeEntries -TicketID "$($Issue.customfield_10313)"

        #If there are already time entries to check against....
        If($TimeEntryResponse -ne "NO TIME ENTRIES")
        {
            #Set an array of TimeEntryID's
            [Array]$TimeEntryIDs = $($TimeEntryResponse.id)

            #Get information on each of those time entries.
            Foreach($TimeEntry in $TimeEntryIDs)
            {
                $TimeEntryResponse = Get-TimeEntryDetails -TimeEntryID $TimeEntry

                If($TimeEntryResponse -ne "NO TIME ENTRY DETAILS")
                {
                    [Array]$TEDetails += $TimeEntryResponse
                }
    
            }
            
            Foreach($Detail in $TEDetails)
            {
                [Bool]$Present = $False
                $ErrorActionPreference = 'SilentlyContinue'
                $RegCheck = ([regex]::matches($Detail.internalnotes, "(?:\[JiraID!!)(.*)(?:!!)")).groups[1].value   
                $ErrorActionPreference = 'Continue'

                If($($Worklog.id) -eq $RegCheck)
                {
                    [Bool]$Present = $True
                    break;
                }
         
            }
            
                          
        }

        Else
        {
            [Bool]$Present = $False
        }
            
        If([Bool]$Present -eq $False)
        {
            If($Closed)
            {
                $OpenIt = Change-CWTicketStatus -TicketID $($Issue.customfield_10313) -Open $True

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

            If($TimeEntry -eq "CW TIME ENTRY CREATION FAILED")
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

    If($JiraName -eq "No User")
    {
        $ContactID = '255093'
        $Firstname = 'Phillip'
        $Lastname = 'Marshall'
    }

    Else
    {
       $SplitName = $JiraName.displayname.split(' ')
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

    If($ContactID -eq "NO CW CONTACT DATA")
    {
        Write-Log "Unable to retrieve a contact for Firstname: $Firstname Lastname: $Lastname CompanyID: $CompanyID"
        $Contactid = ''
    }

    If($MemberCheck)
    {
        $MemberInfo = Get-CWMember -EmailAddress $JiraEmail

        If($MemberInfo -eq "NO CW MEMBER DATA")
        {
            $MemberID = ''
        }

        Else
        {
            $MemberID = $MemberInfo.id
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
	
	.PARAMETER Message
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> Write-Log -Message 'This is the message being written out to the log.' 
	
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

Function Output-Exception
{
    $Output = $_.exception | Format-List -force | Out-String
    $result = $_.Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($result)
    $Reader.BaseStream.Position = 0
    $UsefulData = $reader.ReadToEnd();

    Write-log "[*ERROR*] : `n$Output `n$Usefuldata "  
}

####################################################################
#Variable Declarations
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
[Array]$arrUsernames = @('pmarshall','mduren','mbastian','dmiller','cswain','aquenneville')
[String]$CWServerRoot = "https://api-na.myconnectwise.net/"
[String]$CodeBase = (Invoke-RestMethod -uri 'http://api-na.myconnectwise.net/login/companyinfo/connectwise').codebase
[String]$JiraServerRoot = "https://jira.labtechsoftware.com/"
[String]$DefaultContactEmail = 'pmarshall@labtechsoftware.com'
[String]$Boardname = 'LT-Infrastructure'
[String]$ClosedStatus = 'Completed Contact Confirmed'
[String]$Global:ClosedStatusValue = '7315'
[String]$Global:OpenStatusValue = '5800'
[String]$Global:ReopenStatusName = 'New (Re-Open)'
[String]$Global:NewTicketStatusName = "New"
[Int]$MaxResults = '250'
[String]$LogFilePath = "C:\Scheduled Tasks\Logs\Jira-CW-Platform-Team.txt"

Remove-Item $LogFilePath -Force -ErrorAction 'SilentlyContinue'

#Jira Credentials
$Global:JiraInfo = New-Object PSObject -Property @{
User = 'cwintegrator'
Password = 'kaRnFYpCYEZ9LQQ'
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

    If($UserWorklogs -eq "No Worklogs")
    {
        Write-Log "No Time Entries for User: $User";
        continue;
    }

    ElseIf($UserWorklogs -eq "Exception Caught")
    {
        Write-Log "[*Error*]Exception caught and logged. This user will not be processed this run.";
        continue;    
    }

    Else
    {
        Write-Log "Time Entries Found: $(($Userworklogs | measure-object).count)"
        [INT]$Counter = '1'
        Foreach($Worklog in $UserWorklogs)
        {
            [Int]$Break = 0
            Write-Log "-----------------------------------------------"
            Write-Log "Processing $Counter of $(($Userworklogs | measure-object).count)"
            $Issue = Get-Issue -IssueID "$($worklog.issue.id)"

            If($Issue -eq "TIMEOUT")
            {
                Write-Log "Timeout Detected. Sleeping for 1 Second."
                Start-Sleep -Seconds 1
                $Issue = Get-Issue -IssueID "$($worklog.issue.id)"
                
                If($Issue -eq "Timeout")
                {
                    Write-Log "2nd Timeout Detected. Sleeping for 5 seconds."
                    Start-Sleep -Seconds 5
                    $Issue = Get-Issue -IssueID "$($worklog.issue.id)"

                    If($Issue -eq "Timeout")
                    {
                        Write-Log "Third Timeout Detected. Breaking this loop."
                        [Int]$Break++
                    }
                } 
            }

            ElseIf($Issue -eq "UNKNOWN ERROR")
            {
                Write-Log "Encountered Error Retrieving Issue information"
                Write-Log "This worklog will not be processed this run."
                Write-Log "Worklog: $($Worklog | Format-List -force | Out-String)"
                [Int]$Break++
            }

            ElseIf ($Issue -eq $False -or $Issue -eq $Null)
            {
                Write-log "This damn issue didnt exist somehow."
                Write-Log "This worklog will not be processed this run."
                Write-Log "Worklog: $($Worklog | Format-List -Force)"
                [Int]$Break++
            }

            If($Break -eq 0)
            {
                $ITP_Result = Invoke-TicketProcess -Issue $Issue -Boardname $Boardname -Key $($Worklog.issue.key) -Worklog $Worklog

                If($ITP_Result -eq 'Process Failure' -or $ITP_Result -eq "BAD CW TICKET VALUE")
                {
                    Write-Log "[*Error*]Failure processing Worklog for User: $User | Issue : $($worklog.issue.id)"
                    continue;                
                }

                Else
                {
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
                            $CloseIt = Change-CWTicketStatus -TicketID $($IsClosed.id) -Close $True

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
                }

            }

            $Counter++
      }
}
}