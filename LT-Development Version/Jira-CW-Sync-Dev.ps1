####################################################################
#Jira Functions

Function Get-FilterResults
{
    #$RestApiURI = $JiraServerRoot + "rest/api/2/search?jql=project%20in%20(SDT%2C%20PTC%2C%20LTC)%20AND%20status%20in%20(Closed%2C%20`"Dev%20Queue`"%2C%20`"QA%20Queue`"%2C%20Passed)%20AND%20updated%20>%20-1d%20ORDER%20BY%20updated%20ASC&maxResults=$Global:MaxResults"
    $RestApiURI = $JiraServerRoot + "rest/api/2/search?jql=project%20in%20(SDT%2C%20PTC%2C%20LTC)%20AND%20status%20in%20(Closed%2C%20`"Dev%20Queue`"%2C%20`"QA%20Queue`"%2C%20Passed)&maxResults=$Global:MaxResults"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Method Get

    If($JSONResponse)
    {
        Return $JSONResponse
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

####################################################################
#ConnectWise Post/Edit Functions

Function Update-CWTicketStatus
{
    [cmdletbinding()]
    
    param
    (
    	[Parameter(Mandatory = $true,Position = 0,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$TicketID,
    	[Parameter(Mandatory = $true,Position = 1,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [ValidateNotNullorEmpty()]
		[INT]$StatusID
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

####################################################################
#Data Manipulation Functions

Function Process-IssuesObj
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[PSCustomObject]$Issue
	)

        Write-Log "-----------------------------------------------"
        Write-Log "Issue: $($Issue.key)"
        Write-log "CW Ticket ID: $($Issue.fields.customfield_10313)"
        $CWStatus = (get-cwticket -TicketID $($Issue.fields.customfield_10313)).status.name

        If($($Issue.fields.status.name) -eq "Dev Queue" -and $CWStatus -ne "DEV-Pending fix (Core)")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.customfield_10313) -StatusID $Dev_Pending_Fix_Core
        }

        ElseIf($($Issue.fields.status.name) -eq "QA Queue" -and $CWStatus -ne "QA-Pending Fix Validation")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.customfield_10313) -StatusID $QA_Pending_Fix_Validation
        }

        ElseIf($($Issue.fields.status.name) -eq "Passed" -and $CWStatus -ne "QA-Fix Passed")
        {
            write-log "Status: $($Issue.fields.status.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.customfield_10313) -StatusID $QA_Fix_Passed
        }

        ElseIf($($Issue.fields.status.name) -eq "Closed" -and $($Issue.fields.resolution.name) -eq "Done" -and $CWStatus -ne "Released")
        {
            write-log "Status: $($Issue.fields.status.name)"
            write-log "Resolution: $($Issue.fields.resolution.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.customfield_10313) -StatusID $Released
        }

        ElseIf($($Issue.fields.status.name) -eq "Closed" -and $($Issue.fields.resolution.name) -ne "Done" -and $CWStatus -ne "SDT-Closed Unapproved")
        {
            write-log "Status: $($Issue.fields.status.name)"
            write-log "Resolution: $($Issue.fields.resolution.name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.customfield_10313) -StatusID $SDT_Closed_Unapproved
        }
}

Function Process-IssuesParsed
{
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[PSCustomObject]$Issue
	)

        Write-Log "-----------------------------------------------"
        Write-Log "Issue: $($Issue.key)"
        Write-log "CW Ticket ID: $($Issue.fields.customfield_10313)"
        $CWStatus = (get-cwticket -TicketID $($Issue.fields.Item("customfield_10313"))).status.name

        If($($Issue.fields.Item("status").name) -eq "Dev Queue" -and $CWStatus -ne "DEV-Pending fix (Core)")
        {
            write-log "Status: $($Issue.fields.Item("status").name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.Item("customfield_10313")) -StatusID $Dev_Pending_Fix_Core
        }

        ElseIf($($Issue.fields.Item("status").name) -eq "QA Queue" -and $CWStatus -ne "QA-Pending Fix Validation")
        {
            write-log "Status: $($Issue.fields.Item("status").name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.Item("customfield_10313")) -StatusID $QA_Pending_Fix_Validation
        }

        ElseIf($($Issue.fields.Item("status").name) -eq "Passed" -and $CWStatus -ne "QA-Fix Passed")
        {
            write-log "Status: $($Issue.fields.Item("status").name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.Item("customfield_10313")) -StatusID $QA_Fix_Passed
        }

        ElseIf($($Issue.fields.Item("status").name) -eq "Closed" -and $($Issue.fields.Item("resolution").name) -eq "Done" -and $CWStatus -ne "Released")
        {
            write-log "Status: $($Issue.fields.Item("status").name)"
            write-log "Resolution: $($Issue.fields.Item("resolution").name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.Item("customfield_10313")) -StatusID $Released
        }

        ElseIf($($Issue.fields.Item("status").name) -eq "Closed" -and $($Issue.fields.Item("resolution").name) -ne "Done" -and $CWStatus -ne "SDT-Closed Unapproved")
        {
            write-log "Status: $($Issue.fields.Item("status").name)"
            write-log "Resolution: $($Issue.fields.Item("resolution").name)"
            Write-log "CW Ticket Status: $CWStatus"
            Update-CWTicketStatus -TicketID $($Issue.fields.Item("customfield_10313")) -StatusID $SDT_Closed_Unapproved
        }
}

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

function ConvertFrom-Json2{
<#
    .SYNOPSIS
        The ConvertFrom-Json cmdlet converts a JSON-formatted string to a custom object (PSCustomObject) that has a property for each field in the JSON 

    .DESCRIPTION
        The ConvertFrom-Json cmdlet converts a JSON-formatted string to a custom object (PSCustomObject) that has a property for each field in the JSON 

    .PARAMETER InputObject
        Specifies the JSON strings to convert to JSON objects. Enter a variable that contains the string, or type a command or expression that gets the string. You can also pipe a string to ConvertFrom-Json.

    .PARAMETER MaxJsonLength
        Specifies the MaxJsonLength, can be used to extend the size of strings that are converted.  This is the main feature of this cmdlet vs the native ConvertFrom-Json2

    .EXAMPLE
        Get-Date | Select-Object -Property * | ConvertTo-Json | ConvertFrom-Json

        DisplayHint : 2

        DateTime    : Friday, January 13, 2012 8:06:31 PM

        Date        : 1/13/2012 8:00:00 AM

        Day         : 13

        DayOfWeek   : 5

        DayOfYear   : 13

        Hour        : 20

        Kind        : 2

        Millisecond : 400

        Minute      : 6

        Month       : 1

        Second      : 31

        Ticks       : 634620819914009002

        TimeOfDay   : @{Ticks=723914009002; Days=0; Hours=20; Milliseconds=400; Minutes=6; Seconds=31; TotalDays=0.83786343634490734; TotalHours=20.108722472277776; TotalMilliseconds=72391400.900200009; TotalMinutes=1206.5233483366667;TotalSeconds=72391.4009002}

        Year        : 2012

        This command uses the ConvertTo-Json and ConvertFrom-Json cmdlets to convert a DateTime object from the Get-Date cmdlet to a JSON object.

        The command uses the Select-Object cmdlet to get all of the properties of the DateTime object. It uses the ConvertTo-Json cmdlet to convert the DateTime object to a JSON-formatted string and the ConvertFrom-Json cmdlet to convert the JSON-formatted string to a JSON object..

    .EXAMPLE
        PS C:\>$j = Invoke-WebRequest -Uri http://search.twitter.com/search.json?q=PowerShell | ConvertFrom-Json

        This command uses the Invoke-WebRequest cmdlet to get JSON strings from a web service and then it uses the ConvertFrom-Json cmdlet to convert JSON content to objects that can be  managed in Windows PowerShell.

        You can also use the Invoke-RestMethod cmdlet, which automatically converts JSON content to objects.
        Example 3
        PS C:\>(Get-Content JsonFile.JSON) -join "`n" | ConvertFrom-Json

        This example shows how to use the ConvertFrom-Json cmdlet to convert a JSON file to a Windows PowerShell custom object.

        The command uses Get-Content cmdlet to get the strings in a JSON file. It uses the Join operator to join the strings in the file into a single string that is delimited by newline characters (`n). Then it uses the pipeline operator to send the delimited string to the ConvertFrom-Json cmdlet, which converts it to a custom object.

        The Join operator is required, because the ConvertFrom-Json cmdlet expects a single string.

    .NOTES
        Author: Reddit
        Version History:
            1.0 - Initial release
        Known Issues:
            1.0 - Does not convert nested objects to psobjects
    .LINK
#>

[CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='Low')]

param
(  
    [parameter(
        ParameterSetName='object',
        ValueFromPipeline=$true,
        Mandatory=$true)]
        [string]
        $InputObject,
    [parameter(
        ParameterSetName='object',
        ValueFromPipeline=$true,
        Mandatory=$false)]
        [int]
        $MaxJsonLength = 67108864

)

BEGIN 
{ 

    #Configure json deserializer to handle larger then average json conversion
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Web.Extensions')        
    $jsonserial= New-Object -TypeName System.Web.Script.Serialization.JavaScriptSerializer 
    $jsonserial.MaxJsonLength  = $MaxJsonLength

} #End BEGIN

PROCESS
{
    if ($PSCmdlet.ParameterSetName -eq 'object')
    {
        $deserializedJson = $jsonserial.DeserializeObject($InputObject)

        # Convert resulting dictionary objects to psobjects
        foreach($desJsonObj in $deserializedJson){
            $psObject = New-Object -TypeName psobject -Property $desJsonObj

            $dicMembers = $psObject | Get-Member -MemberType NoteProperty

            # Need to recursively go through members of the originating psobject that have a .GetType() Name of 'Dictionary`2' 
            # and convert to psobjects and replace the current member in the $psObject tree

            $psObject
        }
    }


}#end PROCESS

END
{
}#end END

}

####################################################################
#Variable Declarations
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'
[String]$CWServerRoot = "https://cw.connectwise.net/"
[String]$JiraServerRoot = "https://jira-dev.labtechsoftware.com/"
[Int]$Dev_Pending_Fix_Core = '6012'
[Int]$QA_Pending_Fix_Validation = '5387'
[Int]$QA_Fix_Passed = '5390'
[Int]$Released = '6912'
[Int]$SDT_Closed_Unapproved = '5434'
[Int]$Global:MaxResults = '750'
[String]$LogFilePath = "C:\Scheduled Tasks\Logs\Jira-CW-Doc-Team.txt"

Remove-Item $LogFilePath -Force -ErrorAction 'SilentlyContinue'

#Credentials
$Global:JiraInfo = New-Object PSObject -Property @{
User = 'cwintegration'
Password = '@#WE23we4'
}
$JiraCredentials = Set-JiraCreds
$Global:CWInfo = New-Object PSObject -Property @{
Company = 'connectwise'
PublicKey = '4hc35v3aNRTjib9W'
PrivateKey = 'yLubF4Kfz4gWKBzU'
}
[string]$Authstring  = $CWInfo.company + '+' + $CWInfo.publickey + ':' + $CWInfo.privatekey
$encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(($Authstring)));

$RawData = Get-FilterResults

If($($RawData.GetType().fullname -eq 'System.Management.Automation.PSCustomObject'))
{
    Write-log "Automatic Conversion to PSCustomObject Successful."
    Write-Log "MaxResults : $($Rawdata.maxresults)"
    Write-log "Number of issues returned: $($Rawdata.total)"

    Foreach($Issue in $($RawData.issues))
    {
        Process-IssuesObj -Issue $Issue
    }
}

Else
{
    $ParsedData = ConvertFrom-Json2 -InputObject $RawData
    Write-log "Automatic Conversion to PSCustomObject failed."
    Write-Log "MaxResults : $($ParsedData.maxresults)"
    Write-log "Number of issues returned: $($ParsedData.total)"

    Foreach($Issue in $($RawData.issues))
    {
        Process-IssuesParsed -Issue $Issue
    }
}

