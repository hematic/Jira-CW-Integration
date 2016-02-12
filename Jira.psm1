####################################################################
#Jira Retrieval Functions

Function Get-Issue
{
  	<#
	.SYNOPSIS
		Gets an object representation of a JIRA Issue..
	
	.DESCRIPTION
		Pass this function a JIRA issueid and it will return all
        the information about that issue in JIRA.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Get-Issue -IssueID '8675309'
#>
	
	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[INT]$IssueID
	)

    $RestApiURI = $JiraServerRoot + "rest/agile/1.0/issue/$IssueID"
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

Function Get-JiraUserInfo
{
  	<#
	.SYNOPSIS
		Gets an object representation of a JIRA Issue..
	
	.DESCRIPTION
		Pass this function a JIRA issueid and it will return all
        the information about that issue in JIRA.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Get-Issue -IssueID '8675309'
#>
	
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

Function Get-ActiveSprints
{
  	<#
	.SYNOPSIS
		Gets a list of active sprints for a given board.
	
	.DESCRIPTION
		Makes an API call to get all sprints for a given
        board and then selects only the active ones to return.
        If no sprints are active on that board it will return
        $False.
	
	.NOTES
		If the user making the API call doesn't have permission
        to access a sprint on the board, it returns a 400 error.
        Found this out the hard way.
	
	.EXAMPLE
		Get-ActiveSprints $BoardID
#>
	
	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[INT]$BoardID
	)

    $RestApiURI = $JiraServerRoot + "rest/agile/1.0/board/$BoardiD/sprint?maxResults=$MaxResults"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Method Get
    $ActiveSprints = ($JSONResponse.values | Where-Object {$_.state -eq 'Active'})

    If($ActiveSprints)
    {
        Return $ActiveSprints
    }

    Else
    {
        Return $False
    }

}

Function Get-BoardList
{
	<#
	.SYNOPSIS
		Gets a list of all boards the user has access to see.
	
	.DESCRIPTION
		Returns an object of boards from Jira. This object has
        4 properties: ID, Self, Name, and Type.
	
	.NOTES
        This function has no paramaters. It will return every
        board the user has permissions to access in Jira.
	
	.EXAMPLE
		Get-BoardList
#>

    $RestApiURI = $JiraServerRoot + "rest/agile/1.0/board?maxResults=$MaxResults"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -method Get
    
    If($JSONResponse)
    {
        Return $JSONResponse.values
    }

    Else
    {
        Return $False
    }
}

function Get-SprintInfo
{
	<#
	.SYNOPSIS
		Get details of a specific Jira Sprint
	
	.DESCRIPTION
		Returns all related information of a sprint in Jira.
        Requires you to pass it a sprint id number.
	
	.PARAMETER SprintID
		A description of the SprintID parameter.
	
	.NOTES
		This function does not do error checking, that needs
        to be done where the function is called.
	
	.EXAMPLE
		Get-Sprint '35'
#>
	
	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[INT]$SprintID
	)
	
	$RestApiURI = $JiraServerRoot + "rest/agile/1.0/sprint/$Sprintid/issue?maxResults=$MaxResults"
	$JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Body $Body -method Get

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
  	<#
	.SYNOPSIS
		Builds a Jira Credential String
	
	.DESCRIPTION
		Builds a Jira credential string. Used for API calls.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Get-JiraCreds
#>
	
    $BinaryString = [System.Runtime.InteropServices.marshal]::StringToBSTR($($Jirainfo.password))
    $JPassword = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BinaryString)
    $JLogin = $Jirainfo.user
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes("$jLogin`:$jPassword")
    $JiraCredentials = [System.Convert]::ToBase64String($bytes)
    Return $JiraCredentials
}

Function Get-ActiveProjects
{
  	<#
	.SYNOPSIS
		Gets a list of active sprints for a given board.
	
	.DESCRIPTION
		Makes an API call to get all sprints for a given
        board and then selects only the active ones to return.
        If no sprints are active on that board it will return
        $False.
	
	.NOTES
		If the user making the API call doesn't have permission
        to access a sprint on the board, it returns a 400 error.
        Found this out the hard way.
	
	.EXAMPLE
		Get-ActiveSprints $BoardID
#>
	

    $RestApiURI = $JiraServerRoot + "rest/api/2/project?maxResults=$MaxResults"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Method Get
    $ActiveProjects = $JSONResponse

    If($ActiveProjects)
    {
        Return $ActiveProjects
    }

    Else
    {
        Return $False
    }

}

Function Get-ProjectInfo
{
  	<#
	.SYNOPSIS
		Gets a list of active sprints for a given board.
	
	.DESCRIPTION
		Makes an API call to get all sprints for a given
        board and then selects only the active ones to return.
        If no sprints are active on that board it will return
        $False.
	
	.NOTES
		If the user making the API call doesn't have permission
        to access a sprint on the board, it returns a 400 error.
        Found this out the hard way.
	
	.EXAMPLE
		Get-ActiveSprints $BoardID
#>
	param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[INT]$ProjectID
	)

    $RestApiURI = $JiraServerRoot + "rest/api/2/project/$ProjectID"
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Method Get
    $ActiveProjects = $JSONResponse

    If($ActiveProjects)
    {
        Return $ActiveProjects
    }

    Else
    {
        Return $False
    }

}

####################################################################
#Jira Edit Functions

Function Edit-JiraIssue
{
  	<#
	.SYNOPSIS
		Makes edits to a specific JIRA issue.
	
	.DESCRIPTION
		Pass this function the JIRA IssueID and the 
        CW TicketID and it will update that custom field in JIRA.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Edit-JiraIssue -IssueID '8675309' -CWTicketID '1234567'
#>
	
	param
	(
		[Parameter(Mandatory = $True,Position = 0)]
		[INT]$IssueID,
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
    $JSONResponse = Invoke-RestMethod -Uri $restapiuri -Headers @{ "Authorization" = "Basic $JiraCredentials" } -ContentType application/json -Body $Body -method Put
}

Function Set-JiraCreds
{
  	<#
	.SYNOPSIS
		Builds a Jira Credential String
	
	.DESCRIPTION
		Builds a Jira credential string. Used for API calls.
	
	.NOTES
		N/A
	
	.EXAMPLE
		Get-JiraCreds
#>
	
    $BinaryString = [System.Runtime.InteropServices.marshal]::StringToBSTR($($Jirainfo.password))
    $JPassword = [System.Runtime.InteropServices.marshal]::PtrToStringAuto($BinaryString)
    $JLogin = $Jirainfo.user
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes("$jLogin`:$jPassword")
    $JiraCredentials = [System.Convert]::ToBase64String($bytes)
    Return $JiraCredentials
}