####################################################################
#Data Manipulation Functions

Function Update-CWTicketValue
{
	<#
	.SYNOPSIS
		Creates an insert statement for the 'jira_cw_issues' table.
	
	.DESCRIPTION
        An object containing information about an issue gets passed
        to this function and it creates an insert statement for the 
        'jira_cw_issues' table and then calls the Add-SQLInsert function.
	
	.PARAMETER Issue
		The object containing issue information to be inserted.

	.NOTES
		Some of this data has to be specially formatted to get into the
        mysql database properly. All Date/Time objects must be converted
        from ISo-8601 format to MySQL Date Literals. The attachments are
        parsed to just grab the links to them and concat them with a '|'.
        The subtasks have their ID's joined by a '|' as well to make a
        useable list.
	
	.EXAMPLE
        Create-IssueInsertStatement $Issue
    #>
	
	[CmdLetBinding()]
	Param
		(
		[Parameter(Mandatory = $False)]
		[Int]$CWTicketID,
		[Parameter(Mandatory = $False)]
		[String]$JiraIssueKey

	)

 
    $MySQLInsert = "UPDATE `jira_cw_issues` SET `CWTicketID`=`'$CWTicketID`' WHERE `key`=`'$JiraIssueKey`'"
    Add-SQLInsert $MySQLInsert

}

Function New-WorklogInsertStatement
{

	<#
	.SYNOPSIS
		Creates an insert statement for the 'jira_cw_worklogs' table.
	
	.DESCRIPTION
        An object containing information about a worklog gets passed
        to this function and it creates an insert statement for the 
        'jira_cw_worklogs' table and then calls the Add-SQLInsert function.
	
	.PARAMETER WorkLog
		The object containing worklog information to be inserted.

	.PARAMETER IssueID
		The numerical ID assigned to the issue this worklog is a part of.

	.NOTES
		Some of this data has to be specially formatted to get into the
        mysql database properly. All Date/Time objects must be converted
        from ISo-8601 format to MySQL Date Literals.
	
	.EXAMPLE
        Create-IssueInsertStatement $Issue
	#>

	[CmdLetBinding()]
	Param
		(
		[Parameter(Mandatory = $True)]
		[Object]$Worklog,
        [Parameter(Mandatory = $True)]
        [Int]$IssueID

	)

    $ID = $Worklog.id
    $Author = $Worklog.author.name
    $Comment = Format-SanitizedString $Worklog.comment
    $Created = Get-Date $($Worklog.created) -Format "yyyy-MM-dd HH:mm:ss"
    $Updated = Get-Date $($Worklog.updated) -Format "yyyy-MM-dd HH:mm:ss"
    $Started = Get-Date $($Worklog.started) -Format "yyyy-MM-dd HH:mm:ss"
    $TimeSpent = $Worklog.timespent
    $TimeSpentSeconds = $Worklog.timespentseconds

    $Columns = @"
USE jiraintegration;
Insert IGNORE into `jira_cw_worklogs`
(`ID`,
`IssueID`,
`Author`,
`Comment`,
`Created`,
`Updated`,
`Started`,
`TimeSpent`,
`TimeSpentSeconds`) VALUES 
('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}');
"@
 
    $MySQLInsert = $Columns -f $ID, $IssueID, $Author, $Comment, $Created, $Updated, $Started, $TimeSpent, $TimeSpentSeconds
    Add-SQLInsert $MySQLInsert

}

Function Format-SanitizedString
{
	
	<#
	.SYNOPSIS
		Takes a string and sanitizes it for insert into a MySQL DB.
	
	.DESCRIPTION
        Once a string is passed, the script verifies that any special
        characters that don't play nice with inserts are escaped properly. 
	
	.PARAMETER InputString
		The string to be sanitized.

	.NOTES
        Currently sanitizes the following characters:
            Backslashes
            Single Quotes
            Double Quotes

	
	.EXAMPLE
        Format-SanitizedString $InputString
    #>

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

Function Format-Sprintissues
{
	<#
	.SYNOPSIS
		Creates an object from a Jira issue.
	
	.DESCRIPTION
        A issue or array of issues are passed to
        the function and it parses out the data we
        need and prunes the rest. This leaves a cleaner
        more useable object for us later on.
	
	.PARAMETER Issues
		Either one issue or an array of issues to be parsed.

	.PARAMETER SprintID
		The SprintID this issue is a part of.

	.NOTES
        This function eliminates a lot of unnecessary information
        that is returned for each issue from Jira. If later on a custom
        field needs to be retained that is not currently being used,
        this is what will need to be edited to capture that data.
	
	.EXAMPLE
		Format-Sprintissues -Issues $Sprintinfo.issues -SprintID $Sprint.id
#>

    param
	(
		[Parameter(Mandatory = $true,Position = 0)]
		[Object]$Issues,
        [Parameter(Mandatory = $true,Position = 0)]
		[Object]$SprintID
	)
        Foreach($Issue in $Issues)
        {
            $Global:ObjSprintIssues += New-Object PSObject -Property @{
            Parent = $Issue.fields.parent;
		    Key = $Issue.key;
		    ID = $Issue.id;
            TimeEstimate = $Issue.fields.timeestimate;
            Assignee = $Issue.fields.assignee;
            Status = $Issue.fields.status.name;
            SubTasks = $Issue.fields.subtasks;
            Progress = $issue.fields.progress;
            WorkLog = $Issue.fields.worklog;
            IssueType = $Issue.fields.issuetype.name;
            ProjectKey = $Issue.fields.project.key;
            TimeSpent = $Issue.fields.timespent;
            CreatedDate = $Issue.fields.created;
            UpdatedDate = $Issue.fields.Updated;
            OriginalEstimate = $Issue.fields.timeoriginalestimate;
            Description = $Issue.fields.description;
            TimeTracking = $Issue.fields.Timetracking;
            Attachments = $Issue.fields.attachment;
            Summary = $Issue.fields.summary;
            CWTicketID = $Issue.fields.customfield_10313;
            Sprint = $SprintID
            ParentEpic = $Issue.fields.customfield_10005;
            }
        }
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
        Write-Output "Unable to retrieve a contact for Firstname: $Firstname Lastname: $Lastname CompanyID: $CompanyID"
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

Function Add-SQLInsert
{

	<#
	.SYNOPSIS
		Takes a insert statement and puts it in a MySQL DB.
	
	.DESCRIPTION
        A MySQL insert statement gets passed and this function
        calls mysql.exe to insert it.
	
	.PARAMETER SqlInsert
		The insert statement to be put in the DB.

	.NOTES
        This function currently has the mysql user and password
        hard set. Since the password is being called in plain text
        each time it runs it kicks a watning from mysql out to
        $Error about how its not best practice. This means that you
        can't check $SqlResult for a warning to see if something went
        wrong. I need a better method for this.

	
	.EXAMPLE
        Add-SQLInsert $SqlInsert
    #>

   	Param
		(
		[Parameter(Mandatory = $False)]
		[String]$SqlInsert
	) 
    
    set-location 'C:\Program Files\MySQL\MySQL Server 5.7\bin';

    $SqlResult = .\mysql.exe --user=root --password=wowpass3 -e "$SqlInsert;" --batch -N 2>&1 | Out-Null

}

Function Get-SQLData
{

	<#
	.SYNOPSIS
		Takes a insert statement and puts it in a MySQL DB.
	
	.DESCRIPTION
        A MySQL insert statement gets passed and this function
        calls mysql.exe to insert it.
	
	.PARAMETER SqlInsert
		The insert statement to be put in the DB.

	.NOTES
        This function currently has the mysql user and password
        hard set. Since the password is being called in plain text
        each time it runs it kicks a watning from mysql out to
        $Error about how its not best practice. This means that you
        can't check $SqlResult for a warning to see if something went
        wrong. I need a better method for this.

	
	.EXAMPLE
        Add-SQLInsert $SqlInsert
    #>

   	Param
		(
		[Parameter(Mandatory = $False)]
		[String]$SqlQuery
	) 
    
    set-location 'C:\Program Files\MySQL\MySQL Server 5.7\bin';

    $SqlResult = .\mysql.exe --user=root --password=$MySqlPass "$SqlQuery;" --batch -N

    Return $SqlResult

}

Function Get-Week
{ 
<#
	.SYNOPSIS
		Function to get you a timespan of the current week.
	
	.DESCRIPTION
        This function takes the name of the day of a week ie: 'Saturday'
        that you want the week to start on. It goes backwards and finds
        that previous day then adds 6 days and returns you the start of
        the week, end of the week, and the timespan in an object.
	
	.PARAMETER Weekday
		The day of the week you want the week to START with.

	.EXAMPLE
        Get-Week -WeekDay 'Saturday'
    #>
    param
        (
            [Parameter(Mandatory = $true,Position = 0)]
            [STRING]$WeekDay
        )


    switch ($Weekday)
        {
            “Saturday”  {[int]$Offset = 1} 
            “Sunday”    {[int]$Offset = 0} 
            “Monday”    {[int]$Offset = -1} 
            “Tuesday”   {[int]$Offset = -2} 
            “Wednesday” {[int]$Offset = -3} 
            “Thursday”  {[int]$Offset = -4} 
            “Friday”    {[int]$Offset = -5} 
        }

$Today = $(get-date)
$DaysSince = $Today.DayOfWeek.value__ + $Offset
$WeekBegin = $Today.AddDays(– $DaysSince)
$StartOfWeek = $WeekBegin.addhours(- $($WeekBegin.Hour)).addminutes(- $($WeekBegin.minute)).addseconds(- $($WeekBegin.second))
$EndOfWeek = ((($StartOfWeek.AddDays(6)).addhours(23)).addminutes(59)).addseconds(59)

    $CurrentWeek += New-Object PSObject -Property @{
        Start = $StartOfWeek
        End = $EndOfWeek
        Span = (New-TimeSpan -Start $StartOfWeek -End $EndOfWeek)
        }

Return $CurrentWeek
}

