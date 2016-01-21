﻿$ModuleParentDirectory = 'C:\Users\pmarshall\Dropbox\Work\Projects\Jira Integration\New\'
Import-Module -Name $($ModuleParentDirectory + "ConnectWise.psm1") -force
Import-Module -Name $($ModuleParentDirectory + "DataManipulation.psm1") -force
Import-Module -Name $($ModuleParentDirectory + "Jira.psm1") -force

####################################################################
#Variable Declarations
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'SilentlyContinue'

#Arrays
[Array]$Global:objActiveSprints    = @()
[Array]$Global:objSprintIssues     = @()

#Strings
[String]$CWServerRoot = "https://cw.connectwise.net/"
[String]$JiraServerRoot = "https://jira.labtechsoftware.com/"
[String]$ImpersonationMember = 'jira'
[String]$MySQlPass = 'wowpass3'

#Credentials
$Global:JiraInfo = New-Object PSObject -Property @{
User = 'cwintegrator'
Password = 'kaRnFYpCYEZ9LQQ'
}
$JiraCredentials = Set-JiraCreds

$Global:CWInfo = New-Object PSObject -Property @{
Company = 'connectwise'
User = 'jira'
Password = 'jira'
}
$CWCredentials = Get-CWKeys -ImpersonationMember 'Jira'

#Get a list of all the boards we need to check and insert them into the DB.
$ListofBoards = Get-BoardList

#Pull Active Sprints for the infrastructure board.
foreach($Board in $Listofboards)
{
    If ($Board.name -eq "Infrastructure")
    {
        $Sprints = Get-ActiveSprints -BoardID $Board.id

        If($Sprints -ne $False)
        {
            $objActiveSprints += $Sprints  
        }
    }
}

#Pull all of the issues from each sprint
foreach($Sprint in $objActiveSprints)
{
    $Sprintinfo = Get-SprintInfo -SprintID $Sprint.id
    Format-Sprintissues -Issues $Sprintinfo.issues -SprintID $Sprint.id
}

#Create missing CW Tickets and Map them to Jira
[Int]$Count = 0 | Out-Null
Foreach($Issue in $objSprintIssues)
{
    [Int]$Count++ | Out-Null
    Write-Output "-----------------------------------------------"
    Write-Output "Beginning Issue $($Issue.key)"
    Write-Output "Issue #$Count of $($objsprintissues.count)"
    Invoke-TicketProcess -Issue $Issue -BoardName "LT-Infrastructure"
    Invoke-WorklogProcess -Issue $Issue
    #Close the ticket in CW if its closed in Jira
    If($Issue.status -eq 'Closed')
    {
        $CloseIt = Close-CWTicket -TicketID $($Issue.CWTicketID)
    }
}                                                                                                                                                                                         