
<#PSScriptInfo

.VERSION 2.1.0

.GUID 7daec28b-fcd0-423d-93f6-157a3156f1d3

.AUTHOR Brandon Egbert

.COMPANYNAME 

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES ActiveDirectory, PSLogging

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES

V2.1.0
--Fixed Multiple returns bug
--Various small improvements in error handling, so that you can read it before it closes
--Better logging information
--Added Dev, Prod, and Auto modes to be more flexible
--Changed the hashtable to ordered, and let it be the export
--Changed New-Object -Typename Arraylist to [List[datatype]]::new() to improve performance and match Microsoft standards
--Changed the way the script handled the filters so it was easier to read and understand
--Added formatting functions
--Added Exit-Script function to keep things uniform
--Added capability to move emails to a different folder if successful
--Added capability to send an email to you if for whatever reason an attribute can't be set
--Various other changes for compatiblity for the above changes
#> 


<# 

.DESCRIPTION 
Pulls information from emails to update AD accounts accordingly, with logging

#> 

#region *****************************************************************	settings	*************************************************************
#Requires -Module ActiveDirectory

using namespace System.Collections.Generic;
using namespace Microsoft.ActiveDirectory.Management;

$PSScriptVersion = '2.1.0'
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
#endregion

Function Exit-Script {
    [CmdletBinding()]

    Param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)][string]$Message,
        [Parameter(Mandatory = $false, Position = 1, ValueFromPipeline = $false)][string]$PSErrorMessage
        #     [Parameter(Mandatory = $false, Position = 3)][switch]$TimeStamp,
        #     [Parameter(Mandatory = $false, Position = 4)][switch]$ExitGracefully,
        #     [Parameter(Mandatory = $false, Position = 5)][switch]$ToScreen
    )

    Write-Output ""
    Write-Error -Message $Message
    Write-Output ""
    if ($PSErrorMessage) {
        Write-Error $PSErrorMessage
        Write-Output ""
    }
    Stop-Transcript
    Read-Host "Press 'Enter' to exit"

    exit 1
}
Function Send-AcientEmail {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$SmtpServer = "{0}-com.mail.protection.outlook.com" -f $env:USERDOMAIN,
        [Parameter(Mandatory = $false, Position = 1)][string]$From = ("{0}@{1}.com" -f $env:USERNAME, $env:USERDOMAIN),
        [Parameter(Mandatory = $false, Position = 2)][string]$To = ("{0}@{1}.com" -f $env:USERNAME, $env:USERDOMAIN),
        [Parameter(Mandatory = $true, Position = 3)][string]$Subject,
        [Parameter(Mandatory = $true, Position = 4)][string]$Body
    )

    $mailParams = @{
        SmtpServer                 = $SmtpServer
        Port                       = '25'
        UseSSL                     = $true
        From                       = $From
        To                         = $To
        Subject                    = $Subject
        Body                       = $Body
        DeliveryNotificationOption = 'OnFailure'
    }
    
    Send-MailMessage @mailParams
}

function Write-Line {
    [CmdletBinding()]

    Param (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true)][string]$Message,
        [Parameter(Mandatory = $false, Position = 1)][char]$Character = " ",
        [Parameter(Mandatory = $false, Position = 2)][switch]$ToScreen,
        [Parameter(Mandatory = $false, Position = 3)][switch]$AsHeading,
        [Parameter(Mandatory = $false, Position = 4)][string]$ForegroundColor,
        [Parameter(Mandatory = $false, Position = 5)][string]$BackgroundColor
    )

    $ConsoleWidth = $Host.UI.RawUI.WindowSize.Width

    if ($ForegroundColor) {
        $CurrentFColor = $Host.UI.RawUI.ForegroundColor
    
        switch ($ForegroundColor) {
            "Black" { $Host.UI.RawUI.ForegroundColor = "Black"; break }
            "DarkBlue" { $Host.UI.RawUI.ForegroundColor = "DarkBlue"; break }
            "DarkGreen" { $Host.UI.RawUI.ForegroundColor = "DarkGreen"; break }
            "DarkCyan" { $Host.UI.RawUI.ForegroundColor = "DarkCyan"; break }
            "DarkRed" { $Host.UI.RawUI.ForegroundColor = "DarkRed"; break }
            "DarkMagenta" { $Host.UI.RawUI.ForegroundColor = "DarkMagenta"; break }
            "DarkYellow" { $Host.UI.RawUI.ForegroundColor = "DarkYellow"; break }
            "Gray" { $Host.UI.RawUI.ForegroundColor = "Gray"; break }
            "DarkGray" { $Host.UI.RawUI.ForegroundColor = "DarkGray"; break }
            "Blue" { $Host.UI.RawUI.ForegroundColor = "Blue"; break }
            "Green" { $Host.UI.RawUI.ForegroundColor = "Green"; break }
            "Cyan" { $Host.UI.RawUI.ForegroundColor = "Cyan"; break }
            "Red" { $Host.UI.RawUI.ForegroundColor = "Red"; break }
            "Magenta" { $Host.UI.RawUI.ForegroundColor = "Magenta"; break }
            "Yellow" { $Host.UI.RawUI.ForegroundColor = "Yellow"; break }
            "White" { $Host.UI.RawUI.ForegroundColor = "White"; break }
            default { Write-Error "$ForegroundColor is not a valid selection. Please enter '[Enum]::GetValues([ConsoleColor])' to see the available options" }
        }
    }
    
    
    if ($BackgroundColor) {
        $CurrentBColor = $Host.UI.RawUI.BackgroundColor
        switch ($BackgroundColor) {
            "Black" { $Host.UI.RawUI.BackgroundColor = "Black"; break }
            "DarkBlue" { $Host.UI.RawUI.BackgroundColor = "DarkBlue"; break }
            "DarkGreen" { $Host.UI.RawUI.BackgroundColor = "DarkGreen"; break }
            "DarkCyan" { $Host.UI.RawUI.BackgroundColor = "DarkCyan"; break }
            "DarkRed" { $Host.UI.RawUI.BackgroundColor = "DarkRed"; break }
            "DarkMagenta" { $Host.UI.RawUI.BackgroundColor = "DarkMagenta"; break }
            "DarkYellow" { $Host.UI.RawUI.BackgroundColor = "DarkYellow"; break }
            "Gray" { $Host.UI.RawUI.BackgroundColor = "Gray"; break }
            "DarkGray" { $Host.UI.RawUI.BackgroundColor = "DarkGray"; break }
            "Blue" { $Host.UI.RawUI.BackgroundColor = "Blue"; break }
            "Green" { $Host.UI.RawUI.BackgroundColor = "Green"; break }
            "Cyan" { $Host.UI.RawUI.BackgroundColor = "Cyan"; break }
            "Red" { $Host.UI.RawUI.BackgroundColor = "Red"; break }
            "Magenta" { $Host.UI.RawUI.BackgroundColor = "Magenta"; break }
            "Yellow" { $Host.UI.RawUI.BackgroundColor = "Yellow"; break }
            "White" { $Host.UI.RawUI.BackgroundColor = "White"; break }
            default { Write-Error "$BackgroundColor is not a valid selection. Please enter '[Enum]::GetValues([ConsoleColor])' to see the available options" }
        }
    }
    

    for ([int] $n = 0; $n -lt $ConsoleWidth; $n++) {
        $Line += $Character
    }
    
    #Determines how many characters would fit on the screen
    $CharLength = [int](($ConsoleWidth - $Message.Length - 8) / 2)

    if ($Message) {

        $ReturnMessage = "    $Message     "

        for ([int] $n = 0; $n -lt $CharLength; $n++) {
            $ReturnMessage = Format-String -String $Message -Text $Character -Surround
        }
    }


    if ($AsHeading) {
        $Output = $Line + $ReturnMessage + $Line
    }
    else {
        $Output = $Line
    }
    
        

    if ($ToScreen) {
        Write-Output $Output
    }
    else {
        Write-Verbose $Output
    }

    if ($ForegroundColor) { $Host.UI.RawUI.ForegroundColor = $CurrentFColor }
    if ($BackgroundColor) { $Host.UI.RawUI.BackgroundColor = $CurrentBColor }
}

function Format-String {
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)][string]$String,
        [Parameter(Mandatory = $false, Position = 1)][string]$Text,
        [Parameter(Mandatory = $false, Position = 2)][switch]$Surround,
        [Parameter(Mandatory = $false, Position = 3)][switch]$Prefix,
        [Parameter(Mandatory = $false, Position = 4)][switch]$Suffix

    )

    if ($Surround) {
        return $Text + $String + $Text
    }
    elseif ($Prefix) {
        return $Text + $String
    }
    elseif ($Suffix) {
        return $String + "$Text"
    }
    
}


function Write-HashTable {
    [CmdletBinding()]

    param (	[System.Collections.Specialized.OrderedDictionary]$HashTable)
	
    Write-Verbose -Message  ""
    Write-Verbose -Message  $([pscustomobject]$HashTable | Out-String)
    Write-Verbose -Message  ""
}

#endregion PSLogging


#region ************************************************	Functions	**********************************************************************
function Import-Outlook {

    Write-Line "Entered into Import-Outlook function" -Character " " -AsHeading
    Write-Verbose -Message  "Creating Outlook's Variables"

    #$EmailData, $FilteredDate, $FilteredSender, $FilteredSubject = [List[object]]::new()

    try {
        Write-Verbose -Message  "Starting COM Object"
        $outlookCOM = new-object -comobject outlook.application
        $namespace = $outlookCOM.GetNameSpace('MAPI')
        Write-Verbose -Message "Object created"
    }
    catch {
        #Throws a fatal error in case it can't open outlook
        Exit-Script "Unable to get Outlook data. Please try to run the script again, and make sure that you are not running it as an administrator." -PSErrorMessage $Error[0]
    }
    Write-Line "Exiting Import-Outlook function" -Character " " -AsHeading
    $namespace
}
function Get-EmailData {
    
    param(
        [Parameter(Mandatory = $true, Position = 0)]$NameSpace,
        [Parameter(Mandatory = $false, Position = 1)][string]$MailboxName,
        [Parameter(Mandatory = $false, Position = 2)][string]$MailboxFolder
    )
    Write-Line -Message "Entered Get-EmailData function" -Character " " -AsHeading

    Write-Debug $MailboxName
    Write-Debug $MailboxFolder

    try {
        Write-Verbose -Message  "Checking if a mailbox name was selected"
        if ($MailboxName) {
            Write-Verbose -Message  "Searching Inbox of $MailboxName"
            $Mailbox = $NameSpace.Stores[$MailboxName].GetRootFolder()
            $Inbox = $Mailbox.Folders["Inbox"]
            Write-Verbose -Message  "Checking Inbox for the folder $MailboxFolder"
            $TargetFolder = $Inbox.Folders($MailboxFolder)
        }
        elseif ($MailboxFolder) {
            Write-Verbose -Message  "Checking default mailbox for the folder $MailboxFolder"
            Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
            $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
            $Inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)
            $TargetFolder = $Inbox.Folders($MailboxFolder)
        }
        else {
            Write-Verbose -Message  "Importing default mailbox"
            Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
            $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
            $TargetFolder = $namespace.getDefaultFolder($olFolders::olFolderInBox)
        }
    }
    catch {
        Exit-Script "Unable to locate specified mailbox or folder. Please verify that the names provided properly match the structure" -PSErrorMessage $Error[0]
    }
    Write-Line -Message "Exiting Get-EmailData function" -Character " " -AsHeading
    $TargetFolder
}
function Search-EmailData {
    param (
        [Parameter(Mandatory = $true, Position = 0)][System.__ComObject]$TargetFolder,
        [Parameter(Mandatory = $false, Position = 1)][string]$Sender,
        [Parameter(Mandatory = $false, Position = 2)][string]$Subject,
        [Parameter(Mandatory = $false, Position = 3)][int]$WithinDays
    )
    
    Write-Line -Message "Started Search-EmailData function" -Character " " -AsHeading

    Write-Debug $Sender
    Write-Debug $Subject
    Write-Debug $WithinDays
    $EmailList = $TargetFolder.Items
    if ($EmailList.Count -eq 0) { 
        Exit-Script "No emails were found in the specified folder" -PSErrorMessage $Error[0]
    }
    else {
        Write-Verbose -Message  $("Initial size = " + $EmailList.Count)
    }
    
    
    #region ************************************* Time Filter ****************************************
    try {
        Write-Verbose -Message  "Calculating Days ago"
        $DaysAgo = Get-Date (Get-Date).AddDays(-$WithinDays) -Format "M/d/yyyy HH:mm"
        Write-Verbose -Message  $("Oldest email date accepted: " + $DaysAgo)


        Write-Verbose -Message  "Filtering emails by date"
        $FilteredDate = $EmailList.Restrict("[ReceivedTime] >= '$DaysAgo'")

        $EmailCount = $FilteredDate.Count
        if ($EmailCount -eq 0) { Exit-Script "No emails were found in the specified time range" }
        else { Write-Verbose -Message  $("Emails within timeframe = " + $EmailCount) }
    }
    catch {
        Exit-Script "Unable to filter by date. Please ensure you have the correct information" -PSErrorMessage $Error[0]
    }
    #endregion ************************************* Time Filter ****************************************


    #region ************************************* Sender Filter *****************************************
    try {
        Write-Verbose -Message  "Filtering by sender"
        $FilteredSender = $FilteredDate.Restrict("[SenderName] = '$Sender'")
        $EmailCount = $FilteredSender.Count
        if ($EmailCount -eq 0) { Exit-Script "No emails were found in the specified time range and sender" }
        else { Write-Verbose -Message  $("Time filtered emails with specified sender = " + $EmailCount) }
    }
    catch {
        Exit-Script "Unable to filter by sender. Please ensure you have the correct information" -PSErrorMessage $Error[0]
    }    
    #endregion ********************************** Sender Filter ******************************************	
    

    #region ************************************** subject filter ****************************************
    try {
        Write-Verbose -Message  "Filtering by subject"
        #$EmailData = $FilteredSender | Where-Object Subject -like $Subject
        $EmailData = $FilteredSender.Restrict("[Subject] > '$Subject'") 
        
        if ($EmailData.Count -eq 0) { Exit-Script "No emails were found in the specified time range and sender" }
        else { Write-Verbose -Message  $("Time and sender filtered emails with specified subject = " + $EmailData.Count) }
    }
    catch {
        Exit-Script "Unable to filter by subject. Please ensure you have the correct info" -PSErrorMessage $Error[0]
    }
    #endregion ************************************** subject filter ****************************************
    
    Write-Line -Message  $("There are " + $EmailData.Count + " emails that are being returned from Search-EmailData function") -Character " " -AsHeading
    $($EmailData | Sort-Object -Property "ReceivedTime" -Descending)
}



function Set-LogPath {
    param (
        #[Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 0)][string]$FileType,
        [Parameter(Mandatory = $true, Position = 1)][string]$Path,
        [Parameter(Mandatory = $false, Position = 2)][datetime]$Date
    )

    if (!($Date)) {
        [datetime]$Date = Get-Date
    }

    Write-Verbose "Checking validity of $Path"

    if (!(Test-Path $Path)) {
        Write-Warning "Unable to use selected path. Setting log path to local folder"
        $Path = Convert-Path -Path "."
    }
    else {
        Write-Verbose "$Path is valid"
    }


    switch ($FileType) {
        { $_ -eq 'Csv' } { $ReturnPath = Join-Path -Path $Path -ChildPath $("AdAttribute-DataFile-" + $Date.Year + "-" + $Date.Month + ".csv") }
        { $_ -eq 'Log' } { $ReturnPath = Join-Path -Path $Path -ChildPath $("AdAttribute-Log-" + $Date.ToString("s").Replace(":", ";") + ".log") }
        Default { 
            Exit-Script "Error creating log file path. Please make sure that the FileType you are sending is either 'Csv' or 'Log'"	-PSErrorMessage $Error[0]
        }
    }

    Write-Verbose -Message "File log is $ReturnPath"
    $ReturnPath
}



function Get-Manager {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][string]$Attribute,
        [Parameter(Mandatory = $true, Position = 2)][string]$ID
	
    )
    Write-Line -Message "Entered Get-Manager function" -Character " " -AsHeading
    Write-Verbose -Message  "Searching for manager..."

    Write-Verbose -Message  "Manager's $Attribute should be $ID"
    $Manager = [Microsoft.ActiveDirectory.Management.ADAccount]::new()
	
    if ($ID) {

        $Manager = Get-ADUser -Filter '$Attribute -eq $ID' -Properties DistinguishedName

        if ($Manager) {
            Write-Verbose -Message $("$Manager was matched to $ID")
        }
        else {
            Write-Warning -LogPath $LogPath "No manager was found with $Attribute = $ID"
            $Manager = ""
        }

    }
    else {
        Write-Warning "No $Attribute was found for the manager. Attempting to continue script..."
        $Manager = ""
    }

    return $Manager
}



function Search-Managers {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][List[ADUser]]$UserList,
        [Parameter(Mandatory = $true, Position = 2)][string]$Manager
    )
    Write-Verbose -Message "Filtering list by manager. Correct manager is $Manager"
    $ReturnList = [List[ADUser]]::new()

	
    foreach ($user in $UserList) {
        Write-Verbose -Message $($User.SamAccountName + "'s manager is " + $user.manager)
        Write-Verbose -Message  $($user.SamAccountName)
        Write-Verbose -Message  $("Manager " + $user.Manager)
        if ($user.Manager -eq $Manager) {
            $ReturnList.Add($user) | Out-Null
        }
		
    }
    Write-Verbose -Message  "Returning filtered list"
    
    , $ReturnList
}



function Search-Location {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][List[ADUser]]$UserList,
        [Parameter(Mandatory = $true, Position = 2)][string]$Office
    )
	

    $ReturnList = [List[ADUser]]::new()

    Write-Verbose -Message  "Searching for the $Office location in the list..."
    $ReturnList = $UserList | Where-Object { $_.Office -like $("*" + $Office + "*") }

    , $ReturnList
}



function Search-DisplayName {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][string]$Name,
        [Parameter(Mandatory = $true, Position = 2)]$Properties
    )
    Write-Line "Started Search-DisplayName function" -Character " " -AsHeading

    Write-Verbose -Message  ("Searching display names for $Name...")
    $ReturnList = [List[ADUser]]::new()
	
    $ReturnData = Get-ADUser -Filter { DisplayName -like $Name } -Properties $Properties
	
    foreach ($item in $ReturnData) {
        Write-Verbose -Message  $("User: " + $item.SamAccountName + " is type " + $item.GetType())
        $ReturnList.Add($item) | Out-Null
    }
	
    Write-Line -Message $("Returned " + $ReturnList.Count + " users of type " + $ReturnList.GetType() + " from first check") -Character " " -AsHeading

    , $ReturnList
}




function Search-BothNames {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][string[]]$Properties,
        [Parameter(Mandatory = $true, Position = 2)][System.Collections.IDictionary]$AttributeHash
    )
    Write-Line -Message "Started Search-BothNames function" -Character " " -AsHeading

    $ReturnList = [List[ADUser]]::new()
    $gn = "*" + $AttributeHash.FN + "*"
    $sn = "*" + $AttributeHash.LN + "*"
    $Count = 0
    Write-Verbose -Message $("Searching for users with firstname " + $AttributeHash.FN + " and lastname " + $AttributeHash.LN)
    $data = Get-ADUser -Filter { GivenName -like $gn -and Surname -like $sn } -Properties $Properties
	

    foreach ($pers in $data) {
        Write-Verbose -Message  $("User: " + $pers.SamAccountName + " is type " + $pers.GetType())
        $ReturnList.Add($pers) | Out-Null
        $Count++
    }

	
    Write-Line -Message $("Returning " + $ReturnList.Count + " users of type " + $ReturnList.GetType() + " from First check") -AsHeading -Character " "
	

    , $ReturnList
}



function Start-Search {
    param (
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)][string[]]$Properties,
        [Parameter(Mandatory = $true, Position = 2)][System.Collections.IDictionary]$AttributeHash,
        [Parameter(Mandatory = $false, Position = 3)][bool]$HighPrecision = $false
    )

    Write-Line -Message "Entering Search-BothNames function" -Character " " -AsHeading

    $UserList = [List[ADUser]]::new()
    $Reason = ""
	
    Write-Verbose -Message  $("Data was returned with type " + $UserList.GetType())
    $UserList.AddRange((Search-BothNames -Properties $Properties -AttributeHash $AttributeHash))
	
    Write-Verbose -Message  $("Checking number of accounts returned..." + $UserList.Count)
    #TODO This section cleanup
    #IF more than 2, then if more than 2, then if morethan2
	

    if ($UserList.Count -eq 0) {
		
        Write-Verbose -Message  "No users were found with the matching names. Initiating display name search"
        $Reason += "F&L Name: Failed - Display Name: "
        $UserList.AddRange((Search-DisplayName -Name $AttributeHash.Name -Properties $Properties)) | Out-Null
    }

	

    #Separate
    if ($UserList.Count -gt 1) {

        if ($null -ne $AttributeHash.ManagerDN) {
            Write-Verbose -Message "Calling Search-Managers function"
            $UserList = Search-Managers -UserList $UserList -Manager $AttributeHash.ManagerDN
			
            $Reason += "multiple returns - Manager:"
        }
        else {
            Write-Verbose -Message "Calling Search-Location function because there was no manager"
            $UserList = Search-Location -UserList $UserList -Office $AttributeHash.Location

            $Reason += "multiple returns - Location: "
        }
		
    }


    if ($UserList.Count -eq 0) {
        $Reason += "Failed"
	
        Write-Error -Message $("No users were found. Please manually update " + $AttributeHash.Name)
    }

    elseif ($UserList.Count -eq 1) {
        Write-Verbose -Message "One user was returned"
        foreach ($prop in $Properties) {

            if ($UserList[0].$prop) {
                Write-Verbose -Message $("Adding $prop = " + $UserList[0].$prop + " to hashtable")
                $AttributeHash.Add($prop, $UserList[0].$prop) | Out-Null
            }
            else { 
                Write-Verbose -Message $("Adding $prop as empty value to hashtable")
                $AttributeHash.Add($prop, "") | Out-Null
            }
        }
    }

    else {
        Write-Verbose -Message "Multiple accounts were returned, and contain identical attributes"
        $Reason += "; Multiple accounts"
        foreach ($prop in $Properties) {
            Write-Verbose -Message $("Adding $prop as empty value to hashtable")
            $AttributeHash.Add($prop, "") | Out-Null
        }
    }

    $AttributeHash.Add($($Json.property + " set by script"), $false)
    $AttributeHash.Add("Reason", $Reason)

    Write-Line "Returning from Start-Search function" -Character " " -AsHeading

}


function Read-Email {
    param (
        [Parameter(Mandatory = $true, Position = 0)]$Email, 
        [Parameter(Mandatory = $true, Position = 1)]$Json, 
        [Parameter(Mandatory = $false, Position = 2)][datetime]$Date = $(Get-Date)
    )

    Write-Verbose -Message  ""
    Write-Verbose -Message  "Reading email"
    Write-Verbose -Message  ""
	
    $AttributeHash = [ordered]@{}
	
    $AttributeHash.Add("Date Processed", $Date.ToString("s").Replace(":", ";")) | Out-Null


    #splits each email body up into an array of lines from the single massive string
    Write-Verbose -Message  "Splitting up the email into an array"
    $EmailBodyArray = $Email.Body -split $([System.Environment]::NewLine)
	


    foreach ($Line in $EmailBodyArray) {
        #takes the input, $line, and checks it to see if it contains one of these phrases; Then splits it up to save only the unique data
        #Write-Verbose -Message  $($Line -match $Json.delimiter)
		
        if ($Line -match $Json.delimiter) {
            Write-Verbose -Message  "Splitting $Line on the delimiter"
            $AttributeKey, $AttributeValue = $Line -split $Json.delimiter

            if ($Json.attributeArray -match $AttributeKey) {
                $AttributeHash.Add("Email-$AttributeKey", $AttributeValue) | Out-Null
            }

            switch -wildcard ($AttributeKey) {

                $("*" + $Json.attributes.Name + "*") {
                    Write-Verbose -Message  "Setting name to $AttributeValue"
                    $AttributeHash.Add("Name", $AttributeValue) | Out-Null
                    break
                }
					
                $("*" + $Json.attributes.FN + "*") {
                    Write-Verbose -Message  "Setting first name to $AttributeValue"
                    $AttributeHash.Add("FN", $AttributeValue) | Out-Null
                    break
                }
					
                $("*" + $Json.attributes.LN + "*") {
                    Write-Verbose -Message  "Setting last name to $AttributeValue"
                    $AttributeHash.Add("LN", $AttributeValue) | Out-Null
                    break
                }
					
                $("*" + $Json.attributes.ID + "*") {
                    Write-Verbose -Message  "Setting ID to $AttributeValue"
                    $AttributeHash.Add("ID", $AttributeValue) | Out-Null
                    break
                }
                $("*" + $Json.attributes.Location + "*") { 
                    $AttributeValue, $null = $AttributeValue -split ", "
                    Write-Verbose -Message  "Setting location to $AttributeValue"
                    $AttributeHash.Add("Location", $AttributeValue) | Out-Null
                    break
                }
                $("*" + $Json.attributes.ManagerID + "*") {
                    Write-Verbose -Message  "Setting manager ID to $AttributeValue"
                    $AttributeHash.Add("ManagerID", $AttributeValue) | Out-Null
                    break
                }
            }
        }
    }

    Write-Verbose -Message  $("Adding 'EmailReceivedTime = " + $Email.ReceivedTime + " to the hashtable")
    $AttributeHash += @{
        "EmailReceivedTime" = $Email.ReceivedTime
        " "                 = ""
    }
	

    Write-Verbose -Message  ""
    Write-Verbose -Message  "Returning data from email"
    Write-Verbose -Message  ""

    return , $AttributeHash
}

function Move-Email {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]$Email,
        [Parameter(Mandatory = $true, Position = 1)][System.__ComObject]$Destination
    )

    #TODO figure out how to make this intuitive but work. 
    Write-Line -Message "Entered Move-Email function" -Character " " -AsHeading
    # Write-Debug $Subject
    # Write-Debug $MailboxName
    # Write-Debug $CurrentFolder
    # Write-Debug $NewFolder

    try {
        Write-Verbose "Moving email..."
        $Email.Move($Destination)
    }
    catch {
        Write-Error $("Unable to move email to the specified folder: `n`n" + $Error[0])
    }



    # try {
    #     Write-Verbose -Message  "Checking if a mailbox name or folder was selected"
    #     if ($MailboxName) {

    #         $Mailbox = $NameSpace.Stores[$MailboxName].GetRootFolder()
    #         Write-Debug $Mailbox.GetType()
    #         $Inbox = $Mailbox.Folders["Inbox"]
    #         Write-Debug $Inbox.GetType()

            
            
    #         $ThisFolder = $Inbox.Folders["$CurrentFolder"]
    #         Write-Debug $ThisFolder.GetType()
    #         $CurrentFolderItems = $ThisFolder.Items
    #         $EmailItem = $CurrentFolderItems.Find("[Subject] = '$Subject'")

    #         $TargetNewFolder = $Inbox.Folders["$NewFolder"]
    #         Write-Verbose -Message  "Moving email to $NewFolder"
            
            
            
    #     }
    #     elseif ($CurrentFolder) {
    #         Write-Verbose -Message  "Checking default mailbox for the folder $MailboxFolder"
    #         Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
    #         $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
    #         $Inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)

    #         $ThisFolder = $Inbox.Folders["$CurrentFolder"]
    #         Write-Debug $ThisFolder.GetType()
    #         $CurrentFolderItems = $ThisFolder.Items
    #         $EmailItem = $CurrentFolderItems.Find("[Subject] = '$Subject'")

    #         $TargetNewFolder = $Inbox.Folders["$NewFolder"]
    #         Write-Verbose -Message  "Moving email to $NewFolder"
            
            
    #         $EmailItem.Move($TargetNewFolder)
    #     }
    #     else {
    #         Write-Verbose "Finding the correct email to move"
    #         Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
    #         $olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
    #         $Inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)
    #         $CurrentFolderItems = $Inbox.Items
    #         $EmailItem = $CurrentFolderItems.Find("[Subject] = '$Subject'")

    #         $TargetNewFolder = $Inbox.Folders["$NewFolder"]
    #         Write-Verbose -Message  "Moving email to $NewFolder"
            
            
    #         $EmailItem.Move($TargetNewFolder)
    #     }
    # }
    # catch {
    #     Write-Error $("Unable to locate specified mailbox or folder. Please verify that the names provided properly matches the structure" + $Error[0])
    # }

    Write-Line -Message "Exiting Move-Email function" -Character " " -AsHeading

}
function Test-JSONData {
    param(
        [Parameter(Mandatory = $false, Position = 0)][string]$LogPath,	
        [Parameter(Mandatory = $true, Position = 1)][Object]$JsonData
    )

    try {
        if (!($JsonData.throttleLimit)) {
            $JsonData.throttleLimit = ([int]$env:NUMBER_OF_PROCESSORS + 1)
            Write-Warning $("No throttle limit specified. Proceeding with default limit of " + ([int]$env:NUMBER_OF_PROCESSORS + 1))
		
        }
        if (!($JsonData.daysAgo)) {
            $JsonData.daysAgo = 30
            Write-Warning "No date filter specified. Proceeding with default date range of 30 days"
		
        }
        if (!($JsonData.searchSubject)) {
            $JsonData.searchSubject = $null
            Write-Warning "No 'Subject' filter specified"
		
        }
        if (!($JsonData.searchSender)) {
            $JsonData.searchSender = $null
            Write-Warning "No 'From' filter specified"
		
        }
        if (!($JsonData.property)) {
            $JsonData.property = $null
            Write-Error "No property was specified. Please specify the property in the config.json file" -ErrorAction stop
        }
        if (!($JsonData.delimiter)) {
            $JsonData.delimiter = ": "
            Write-Warning "No delimiter was specified. Proceeding with the default ': '"
		
        }
        if (!($JsonData.searchBase)) {
            $JsonData.searchBase = "*"
            Write-Warning "No searchbase specified. The default searchbase will be used"
		
        }
    }
    catch {
        Exit-Script "Error validating data"  -PSErrorMessage $Error[0]
    }

    Write-Verbose -Message  $("")
    Write-Verbose -Message  $("Throttle Limit: " + $JsonData.throttleLimit)
    Write-Verbose -Message  $("Days Ago: " + $JsonData.daysAgo)
    Write-Verbose -Message  $("Email Subject: " + $JsonData.searchSubject)
    Write-Verbose -Message  $("Email Sender: " + $JsonData.searchSender)
    Write-Verbose -Message  $("Property: " + $JsonData.property)
    Write-Verbose -Message  $("Delimiter: " + $JsonData.delimiter)
    Write-Verbose -Message  $("Searchbase: " + $JsonData.searchBase)
    Write-Verbose -Message  $("Mailbox Name: " + $JsonData.mailboxName)
    Write-Verbose -Message  $("Mailbox Folder: " + $JsonData.mailboxFolder)
    Write-Verbose -Message  $("Log Path: " + $JsonData.logPath)
    Write-Verbose -Message  $("")

    Write-Line -Message "Returning from Test-JSONData function" -Character " " -AsHeading
    return $JsonData
}



function Set-ScriptMode {
    param (
        [Parameter(Mandatory = $true, Position = 0)][string]$Mode
    )

    Write-Verbose "Entered Set-ScriptMode function"

    Write-Information "Setting mode to $mode"
    
    Write-Verbose "Exiting Set-ScriptMode function"
}

#endregion

#region ************************************************************	initialization		*******************************************************
# #TODO figure out how to prioritize everything
# Start-Log -LogPath "." -LogName "initiallog.log" -scriptversion 2.1.0
# $LogPath = ".\initiallog.log"

try {
    Write-Information "Importing ActiveDirectory module"
    Import-Module -Name ActiveDirectory -Force
}
catch {
    Exit-Script -Message "Unable to import ActiveDirectory module. Please make sure it is installed before proceeding" -PSErrorMessage $Error[0]
}


try {
    Write-Information "Getting config.json" -InformationAction Continue
    [Object]$JsonData = Get-Content "config.json" | ConvertFrom-Json
}
catch {
    Exit-Script "Unable to get config.json. Please make sure it is located in the same location as the script"  -PSErrorMessage $Error[0]
}


try {
    $Mode = $JsonData.mode
    switch ($Mode) {
        'prod' {
            $VerbosePreference = 'silentlycontinue'
            $InformationPreference = 'continue'
        }
        'dev' {
            $DebugPreference = 'continue'
            $VerbosePreference = 'continue'
            $InformationPreference = 'continue'
        }
        'auto' {
            $VerbosePreference = 'silentlycontinue'
            $InformationPreference = 'silentlycontinue'
        }
		
        Default {
            Exit-Script "Unknown mode. Please select 'prod', 'dev', or 'auto'" -PSErrorMessage $Error[0]
        }
    }
}
catch {
    Exit-Script "Script failed with unknown error when trying to set the mode" -PSErrorMessage $Error[0]
}

Write-Information "Initializing Script"
#endregion


#region transcript
try {
    $LogPath = Set-LogPath -Path $JsonData.logPath -FileType 'Log'
    $CsvPath = Set-LogPath -Path $JsonData.logPath -FileType 'Csv'
}
catch {
    Exit-Script "Unable to retrieve logpath for some reason" -PSErrorMessage $Error[0]
}


try {
    Start-Transcript -Path $LogPath -Force
    Write-Output "Script Version: $PSScriptVersion"
}
catch {
    Exit-Script "Unable to create log" -PSErrorMessage $Error[0]
}

#endregion


#region data verification
[object]$Json = Test-JSONData -JsonData $JsonData -LogPath $LogPath

if ($Mode -eq 'Dev') {
    try {
        Write-Debug -Message  "Removing csv"
        Remove-Item $CsvPath -Force
    }
    catch { 
        Exit-Script "Unable to remove previous csv" -PSErrorMessage $Error[0]
    }
}

#endregion


#region variables
Write-Verbose -Message  "Initializing variables..."
$Failures = [List[object]]::new()
$CsvData = [List[hashtable]]::new()
$AttributeHash = [ordered]@{}
$ProgCount, $EmailCount = [int]::new()
$Date = Get-Date
$PropertyArray = $($Json.'property'), 'DisplayName', 'SamAccountName', 'GivenName', 'Surname', 'Manager', 'Office', 'Created'
#endregion


#region assignments
Write-Output "Importing Outlook data..."

$OutlookNamespace = Import-Outlook

$TargetFolder = Get-EmailData -NameSpace $OutlookNamespace -MailboxName $Json.mailboxName -MailboxFolder $Json.mailboxFolder

$splat = @{
    TargetFolder = $TargetFolder
    Sender       = $Json.searchSender
    Subject      = $Json.searchSubject
    WithinDays   = $Json.daysAgo
}

$EmailData = Search-EmailData @splat
$EmailCount = $EmailData.Count
Write-Verbose -Message  "Final email count: $EmailCount"
$ProgCount = 0
#endregion


#starts checking each email one at a time
foreach ($item in $EmailData) {


    #region ********************	Progress Tracking	********************
    Write-Debug -Message  "Incrementing $ProgCount by one"
    $ProgCount++

    $MoveEmail = [bool]::new()
    $Email = $item | Select-Object ReceivedTime, SenderName, Subject, Body

    switch ($Mode) {
        "Dev" {	Write-Verbose -Message $("Started processing email $ProgCount out of $EmailCount"); break }
        "Prod" { Write-Progress $("Started processing email $ProgCount out of $EmailCount"); break }
        "Auto" { Write-Output -Message $("Started processing email $ProgCount out of $EmailCount"); break }
        Default { }#Add-Content -Path $LogPath -Value $("Started processing email $ProgCount out of $EmailCount") }
    }
    #endregion ********************		Progress Tracking	********************


    #region ********************	Hash Table	********************
    Write-Verbose -Message  "Clearing hashtable"
    $AttributeHash.Clear()

    Write-Verbose -Message  $("Calling Read-Email function")
    $AttributeHash = $(Read-Email -Email $Email -Json $Json)
	
    Write-HashTable $AttributeHash
    #endregion ********************	Hash Table	********************


    #region ********************	Manager		********************
    Write-Verbose -Message  "Calling Get-Manager function"
    $Manager = [Microsoft.ActiveDirectory.Management.ADAccount]::new()
    $Manager = Get-Manager -Attribute $Json.'property' -ID $AttributeHash.ManagerID
	
    if ($Manager) {
        Write-Verbose -Message  $("Adding manager Distinguished Name to hashtable as " + $Manager.DistinguishedName)
        $AttributeHash.Add("ManagerDN", $Manager.DistinguishedName) | Out-Null
    }
    else {
        Write-Warning $("No manager was found with " + $Json.property + " " + $AttributeHash.ManagerID)
        $AttributeHash.Add("ManagerDN", "") | Out-Null
    }
    #endregion ********************		Manager		********************

	
    #region ********************	search		********************
    Write-Verbose -Message  "Calling Start-Search function"
    $AttributeHash.Add("Matched", $false) | Out-Null
    Start-Search -AttributeHash $AttributeHash -Properties $PropertyArray
    Write-Line -Character "*" -ForegroundColor "Magenta"
    Write-HashTable -HashTable $AttributeHash -Verbose
    Write-Line -Character "*" -ForegroundColor "Magenta"
    #endregion ********************	search		********************

	
    if ($AttributeHash.SamAccountName) {

        Write-Verbose -Message $("Matched " + $AttributeHash.Name + " to " + $AttributeHash.SamAccountName)
        $AttributeHash.Matched = $true


        if ($AttributeHash.($Json.'property')) {
            if ($AttributeHash.($Json.'property') -eq $AttributeHash.ID) {
                $AttributeHash.Reason += $("; " + $Json.'property' + " already set for " + $AttributeHash.SamAccountName)
                Write-Verbose -Message $($Json.'property' + " already set for " + $AttributeHash.SamAccountName)
                $MoveEmail = $true
            }
            else {
                $Message = $($Json.'property' + " for " + $AttributeHash.SamAccountName + " is currently set as " + $AttributeHash.($Json.'property') + " but the email says it should be " + $AttributeHash.ID)
                Write-Error $Message
                $Failures.Add($AttributeHash)
                $AttributeHash.Reason = $Message
            }
        }

        elseif (!($AttributeHash.($Json.'property'))) {

            Write-Verbose -Message  $("Setting " + $AttributeHash.SamAccountName + " " + $Json.'property' + " to " + $AttributeHash.ID )

            try {
                if ($Mode = 'Dev') {
                    Set-ADUser $AttributeHash.SamAccountName -Add @{$($Json.'property') = $AttributeHash.ID } -WhatIf
                }
                else {
                    Set-ADUser $AttributeHash.SamAccountName -Add @{$($Json.'property') = $AttributeHash.ID }
                    $MoveEmail = $true
                }
                $AttributeHash[($Json.property + " set by script")] = $true
                Write-Verbose -Message $("Set " + $Json.'property' + " to " + $AttributeHash.ID + " for " + $AttributeHash.SamAccountName)
            }
            catch {
                
                $AttributeHash.Reason += $("; Unable to set " + $Json.'property' + " for " + $AttributeHash.SamAccountName)
                $Failures.Add($AttributeHash) | Out-Null
                Write-Error -Message $("Unable to set " + $Json.'property' + " for " + $AttributeHash.SamAccountName)
            }
        }
    }

    else {
        $AttributeHash.Matched = $false
        $Failures.Add($AttributeHash) | Out-Null
    }


    if (($MoveEmail -eq $true) -and ($null -ne $Json.moveToFolder)) {
    
        $Subject = $Email.Subject
        $MoveSplat = @{
            Email   = $item
            Destination = Get-EmailData -NameSpace $OutlookNamespace -MailboxName $Json.mailboxName -MailboxFolder $Json.moveToFolder
        }
        Move-Email @MoveSplat
    }


    Write-HashTable $AttributeHash -Verbose

    $CsvData += [PSCustomObject]$AttributeHash
}

#region ****************** Cleanup *******************************
try {
    Write-Output "All emails processed. Exporting data to Csv"

    $CsvData | Export-Csv -Path $CsvPath -NoTypeInformation -Append -force
    Stop-Transcript
}
catch {
    Write-Error -Message $("Error exporting to Csv" + $Error[0])
}

try {
    if ($Failures.Count -gt 0) {
        $EmailProps = @{
            $SmtpServer = $Json.emailSmtpServer
            $From       = $Json.emailSender
            $To         = $Json.emailTo
            $Subject    = $Json.emailSubject + " - $Date"
            $Body       = "Failed to set attribute for: `n$Failures"
        }
        Send-AcientEmail @EmailProps
    }
}
catch {
    Write-Error -Message $("Unable to send failure log.`n" + $Error[0])
}

switch ($Mode) {
    "Dev" { 
        #Add-Content -Path $LogPath -Value $("Total time: " + $(Stop-TimeLog -Time $Start))
        Invoke-Item $CsvPath; break 
    }
    "Prod" { 
        #Add-Content -Path $LogPath -Value $("Total time: " + $(Stop-TimeLog -Time $Start))
        Read-Host "Done ("; break 
    }
    "Automation" { 
        #Write-LogInfo $("Total time: " + $(Stop-TimeLog -Time $Start)) -LogPath $LogPath
        Invoke-Item -Path $LogPath; break 
    }
    Default {}
}
#endregion