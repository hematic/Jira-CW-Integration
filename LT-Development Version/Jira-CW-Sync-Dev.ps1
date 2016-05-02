#Function Declarations
####################################################################
Function Set-JiraCreds
{
    $BinaryString = [System.Runtime.InteropServices.marshal]::StringToBSTR($($Jirainfo.password))
    $JPassword = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BinaryString)
    $JLogin = $Jirainfo.user
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes("$jLogin`:$jPassword")
    $JiraCredentials = [System.Convert]::ToBase64String($bytes)
    Return $JiraCredentials
}

Function Get-FilterResults
{
    #$RestApiURI = $JiraServerRoot + "rest/api/2/search?jql=project%20in%20(SDT%2C%20PTC%2C%20LTC)%20AND%20status%20in%20(Closed%2C%20`"Dev%20Queue`"%2C%20`"QA%20Queue`"%2C%20Passed)%20AND%20updated%20>%20-1d%20ORDER%20BY%20updated%20ASC&maxResults=$Global:MaxResults"
    $RestApiURI = $JiraServerRoot + "rest/api/2/search?jql=project%20in%20(SDT%2C%20PTC%2C%20LTC)%20AND%20status%20in%20(Closed%2C%20`"Dev%20Queue`"%2C%20`"QA%20Queue`"%2C%20Passed)&maxResults=$Global:MaxResults"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Method Get

    #If there are an extremely large number of tickets returned PowerShell will fail to parse them to a PSCustomObject.
    #This line checks to see if that happened and if it does, passes all of the data through a conversion function that
    #will make us the object we are looking for.
    
    If($JSONResponse)
    {
        If($($JSONResponse.GetType().fullname -ne 'System.Management.Automation.PSCustomObject'))
        {
            $ParsedData = ConvertFrom-Json2 -InputObject $JSONResponse
            Write-log "Automatic Conversion to PSCustomObject failed."
            Write-Log "Data has been parsed instead."
            Write-Log "MaxResults : $($RawData.maxresults)"
            Write-log "Number of issues returned: $($RawData.total)"
            [Bool]$Global:Parsed = $True
            Return $ParsedData
        }

        Else
        {
            Write-log "Automatic Conversion to PSCustomObject Successful."
            Write-Log "MaxResults : $($Rawdata.maxresults)"
            Write-log "Number of issues returned: $($Rawdata.total)"
            Return $JSONResponse
        }
    }

    Else
    {
        Return $False
    }
}

Function Get-CWTicket
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[Int]$TicketID
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

Function Change-CWTicketStatus
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $True,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID,
    	[Parameter(Mandatory = $True,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[Int]$StatusID
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
                "op" : "replace", "path": "/status/id", "value": "$StatusID"
            }
            ]
"@
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
        If($JSONResponse.status.name)
        {
            Write-Log "Status Change Succeeded"
            Return "Status Change Succeeded"
        }

        Else
        {
            Write-Log "Status Change Failed"
            Return "Status Change Failed"
        }
    }
}

Function Check-Mapping
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[String]$CWStatus,
		[Parameter(Mandatory = $True,Position = 1)]
        [String]$JiraStatus,
        [Parameter(Mandatory = $True,Position = 2)]
        [Int]$CWTicketID,
		[Parameter(Mandatory = $False,Position = 3)]
        [String]$Resolution

	)

        If($JiraStatus -eq "SDT Queue" -and $CWStatus -ne "SDT-In Progress")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $SDT_In_Progress
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Check-in Ready" -and $CWStatus -ne "SDT-In Progress")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $SDT_In_Progress
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Dev Queue" -and $CWStatus -ne "DEV-Pending Fix")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $Dev_Pending_Fix
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Build Ready" -and $CWStatus -ne "DEV-Pending Fix")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $Dev_Pending_Fix
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "QA Assign" -and $CWStatus -ne "QA-Pending Fix Validation")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $QA_Pending_Fix_Validation
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "QA Queue" -and $CWStatus -ne "QA-Testing")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $QA_Testing
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Passed" -and $CWStatus -ne "QA-Fix Passed")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $QA_Fix_Passed
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Failed" -and $CWStatus -ne "QA-Fix Failed")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $QA_Fix_Failed
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Closed" -and $Resolution -eq "Done" -and $CWStatus -ne "Released")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $Released
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Closed" -and $Resolution -ne "Done" -and $CWStatus -ne "SDT-Closed Unapproved")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $SDT_Closed_Unapproved
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Reopened" -and $CWStatus -ne "Dev-Rework Fix")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $Dev_Rework_Fix
            Return $StatusChangeResult
        }

        ElseIf($JiraStatus -eq "Confirm Resolution" -and $CWStatus -ne "SDT-In Progress")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            $StatusChangeResult = Change-CWTicketStatus -TicketID $CWTicketID -StatusID $SDT_In_Progress
            Return $StatusChangeResult
        }

        Else
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Write-log "Ticket was already mapped correctly."
            Return "Ticket was already mapped correctly."
        }
    
}

Function Change-JiraPilotFlag
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[String]$IssueID,
		[Parameter(Mandatory = $True,Position = 1)]
		[String]$PilotValue
	)

$Body= @"
{
"fields":
	{
	"customfield_11409" : "$PilotValue"
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

    Write-Verbose "$($JsonResponse2.customfield_11409)"

    If($($JsonResponse2.customfield_11409) -eq $PilotValue)
    {
        Return "Success"    
    }

    Else
    {
        Return "Failed to Set PilotValue"
    }
}

Function Process-Issues
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[PSCustomObject]$Issue,
		[Parameter(Mandatory = $True,Position = 1)]
        [Bool]$Parsed
	)

    Write-Log "-----------------------------------------------"
    Write-Log "Issue: $($Issue.key)"
    Write-log "CW Ticket ID: $($Issue.fields.customfield_10313)"

    #Handled the parsed version of the object...
    If($Parsed -eq $True)
    {
        #Make sure there is a customfield value.
        If($($Issue.fields.Item("customfield_10313")) -eq $Null -or $($Issue.fields.Item("customfield_10313")) -eq '')
        {
            Write-log "[*ERROR*]No CW Ticket is populated in the Customfield!"
            Return "Empty CustomField"
        }

        #Verifies that the customfield value contains only numbers.
        If($($Issue.fields.Item("customfield_10313")) -as [Int] -isnot [Int])
        {
            Write-log "[*ERROR*]BAD CW Ticket value is populated in the Customfield!"
            Return "Bad CustomField"
        }

        #Set the values we need to make the actual checks and calls easier to read.
        [Int]$Jira_CWTicket_CF_Value = $Issue.fields.Item("customfield_10313") #This is the value of the CWTicket Customfield in Jira
        [String]$Jira_Pilot_Flag_Value = $Issue.fields.Item("customfield_11409").value #This is the value of the Pilot Flag CustomField in Jira
        [String]$JiraStatus = $Issue.fields.Item("status").name #This is the status of the JIRA Ticket
        [String]$Resolution = $Issue.fields.Item("resolution").name #This is the Resolution of the JIRA Ticket
        $CWTicketInfo = Get-CWTicket -TicketID $Jira_CWTicket_CF_Value
        If($CWTicketInfo -eq 'Bad Request')
        {
            Return 'Bad Request"'
        }
        [String]$CWStatus = $CWTicketInfo.status.name #This is the status of the CW Ticket
        [String]$CW_Pilot_Flag_Value = $CWTicketInfo.customfields[15].value
    }

    Else
    {
        #Make sure there is a customfield value.
        If($($Issue.fields.customfield_10313) -eq $Null -or $($Issue.fields.customfield_10313) -eq '')
        {
            Write-log "[*ERROR*]No CW Ticket is populated in the Customfield!"
            Return "Empty CustomField"
        }

        #Verifies that the customfield value contains only numbers.
        If($($Issue.fields.customfield_10313) -as [Int] -isnot [Int])
        {
            Write-log "[*ERROR*]BAD CW Ticket value is populated in the Customfield!"
            Return "Bad CustomField"
        }

        #Set the values we need to make the actual checks and calls easier to read.
        [Int]$Jira_CWTicket_CF_Value = $Issue.fields.customfield_10313
        [String]$Jira_Pilot_Flag_Value = $Issue.fields.customfield_11409
        [String]$JiraStatus = $Issue.fields.status.name
        [String]$Resolution = $Issue.fields.resolution.name
        $CWTicketInfo = Get-CWTicket -TicketID $($Issue.fields.customfield_10313)
        If($CWTicketInfo -eq 'Bad Request')
        {
            Return 'Bad Request"'
        }
        [String]$CWStatus = $CWTicketInfo.status.name
        [String]$CW_Pilot_Flag_Value = $CWTicketInfo.customfields[15].value 
    }

    If($CWStatus -eq "Bad Request")
    {
        Write-Log "Unable to Retrieve ConnectWise Ticket Information for CW Ticket : $($Issue.fields.customfield_10313)"
        Return "Failed to Retrieve CW Ticket Data"
    }
    If($Resolution)
    {
        $MappingResults = Check-Mapping -CWStatus $CWStatus -JiraStatus $JiraStatus -CWTicketID $Jira_CWTicket_CF_Value -Resolution $Resolution
        If($MappingResults -eq "Status Change Failed")
        {
            Return $MappingResults
        }
    }

    Else
    {
        $MappingResults = Check-Mapping -CWStatus $CWStatus -JiraStatus $JiraStatus -CWTicketID $Jira_CWTicket_CF_Value
        If($MappingResults -eq "Status Change Failed")
        {
            Return $MappingResults
        }    
    }

    If($Jira_Pilot_Flag_Value = 'yes')
    {
        $Jira_Pilot_Flag_Value = $True
    }

    Else
    {
        $Jira_Pilot_Flag_Value = $False
    }

    If($CW_Pilot_Flag_Value -eq $True -and $Jira_Pilot_Flag_Value -ne $True )
    {
            Write-Log "JIRA Pilot Value = $Jira_Pilot_Flag_Value | CW Pilot Value = $CW_Pilot_Flag_Value. Changing the value in JIRA to match."
            $ChangeAttempt = Change-JiraPilotFlag -IssueID $($Issue.id) -CW_Pilot_Flag_Value $Jira_Pilot_Flag_Value
            If($ChangeAttempt.customfields[15])
            {
                If($($ChangeAttempt.fields.customfield_11409) -eq $CW_Pilot_Flag_Value)
                {
                    Write-Log "Change Successful."    
                }

                Else
                {
                    Write-Log "Change FAILED."
                }
            }

            Else
            {
                Write-Log "Change FAILED."
            }
        }


}

Function Write-Log
{
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    
	add-content -path $LogFilePath -value ($Message)
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
[String]$CWServerRoot = "https://idev.connectwisedev.com/"
[String]$JiraServerRoot = "https://jira-dev-clone.labtechsoftware.com/"
[String]$LogFilePath = "C:\Scheduled Tasks\Logs\Jira-CW-Dev-Team.txt"
Remove-Item $LogFilePath -Force -ErrorAction 'SilentlyContinue'

#This is the list of statuses for the LT-Dev board that we are currently using.
[Int]$SDT_In_Progress = '5408'
[Int]$Dev_Pending_Fix = '5385'
[Int]$QA_Pending_Fix_Validation = '5387'
[Int]$QA_Testing = '7785'
[Int]$QA_Fix_Passed = '5390'
[Int]$QA_Fix_Failed = '7336'
[Int]$Released = '6912'
[Int]$SDT_Closed_Unapproved = '5434'
[Int]$Dev_Rework_Fix = '5388'
[Int]$Global:MaxResults = '750'
[Bool]$Global:Parsed = $False
[Array]$ProcessingErrors = @("Empty CustomField",
                             "Failed to Retrieve CW Ticket Data",
                             "Bad CustomField",
                             "Status Change Failed")

#Credentials used for JIRA and CW
$Global:JiraInfo = New-Object PSObject -Property @{
User = 'cwintegration'
Password = '@#WE23we4'
}
$JiraCredentials = Set-JiraCreds

$Global:CWInfo = New-Object PSObject -Property @{
Company = 'idev'
PublicKey = '4hc35v3aNRTjib9W'
PrivateKey = 'yLubF4Kfz4gWKBzU'
}
[string]$Authstring  = $CWInfo.company + '+' + $CWInfo.publickey + ':' + $CWInfo.privatekey
$encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

#Gather the list of tickets to act upon from the JIRA API.
$JiraData = Get-FilterResults

If($($RawData.GetType().fullname -ne 'System.Management.Automation.PSCustomObject'))
{
    $Jiradata = ConvertFrom-Json2 -InputObject $RawData
    [Bool]$Global:Parsed = $True
}

#We loop through each issue and process them as needed. Some filtering could be done here to ONLY loop through
#the issues with the statuses we know we need to touch, but i didn't put in the effort. I feel bad.

Foreach($Issue in $($JiraData.issues))
{
    $ProcessResult = Process-Issues -Issue $Issue -Parsed $Global:Parsed

    If($ProcessingErrors -contains $ProcessResult)
    {
        Write-Log "Processing Issue $($Issue.key) : Failed."
    }

    Else
    {
        Write-Log "Processing Issue $($Issue.key) : Succeeded."
    }
}