
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

.EXTERNALMODULEDEPENDENCIES ActiveDirectory

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
Fixed Multiple returns bug

#> 

#Requires -Module ActiveDirectory

<# 

.DESCRIPTION 
Pulls information from email to update AD accounts and provide logging

#> 

#region settings
Set-StrictMode -Version Latest
$DebugPreference = 'continue'
$VerbosePreference = 'continue'
#endregion

#region functions
function Import-OutlookData {

	param(
		[System.Object]$SearchData 
	)

	Write-Verbose "Creating Outlook's Variables"
	$EmailData, $FilteredDate, $FilteredSender, $FilteredSubject = New-Object -TypeName System.Collections.ArrayList
	[string]$MailboxName = $SearchData.mailboxName
	[string]$MailboxFolder = $SearchData.mailboxFolder

	try {
		Write-Verbose "Starting COM Object"
		$outlookCOM = new-object -comobject outlook.application
		$namespace = $outlookCOM.GetNameSpace('MAPI')
		Add-type -assembly 'Microsoft.Office.Interop.Outlook' | out-null
		$olFolders = 'Microsoft.Office.Interop.Outlook.olDefaultFolders' -as [type]
		Write-Verbose "Checking if a mailbox name was selected"
	}
	catch {
		#Throws a fatal error in case it can't open outlook
		Write-Error "Unable to get Outlook data. Please try to run the script again, and make sure that you are not running it as an administrator. `nError: $_.exception.message" -ErrorAction Stop
	}
	try {
		if ($MailboxName) {
			Write-Verbose "Searching Inbox of $MailboxName"
			$Mailbox = $NameSpace.Stores[$MailboxName].GetRootFolder()
			$Inbox = $Mailbox.Folders["Inbox"]
			Write-Verbose "Checking Inbox for the folder $MailboxFolder"
			$TargetFolder = $Inbox.Folders($MailboxFolder)
		}
		elseif ($MailboxFolder) {
			Write-Verbose "Checking default mailbox for the folder $MailboxFolder"
			$Inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)
			$TargetFolder = $Inbox.Folders($MailboxFolder)
		}
		else {
			Write-Verbose "Importing default mailbox"
			$TargetFolder = $namespace.getDefaultFolder($olFolders::olFolderInBox)
		}
	}
	catch {
		Write-Error "Unable to locate specified mailbox or folder. Please verify that the names provided properly match the structure `n
		Error: $_.exception.message" -ErrorAction Stop
	}
	
	Write-Debug $("Initial size = " + $TargetFolder.Items.Count)

	Write-Verbose "Calculating Days ago"
	$DaysAgo = Get-Date (Get-Date).AddDays(-$SearchData.daysAgo) -Format "M/d/yyyy HH:mm"
	Write-Debug $DaysAgo

	Write-Verbose "Filtering emails by date"
	$FilteredDate = $TargetFolder.Items.Restrict("[ReceivedTime] >= '$DaysAgo'")
	Write-Debug $("Time emails = " + $FilteredDate.Count)

	Write-Verbose "Filtering by sender"
	$senderemail = $SearchData.emailSender
	$FilteredSender = $FilteredDate.Restrict("[SenderName] = '$senderemail'")
	Write-Debug $("Sender emails = " + $FilteredSender.Count)
		
	Write-Verbose "Filtering by subject"
	#$FilteredEmails = $FilteredSender.Restrict("[Subject] = $($SearchData.emailSubject + "*")")
	$EmailData = $FilteredSender | Where-Object Subject -like $SearchData.emailSubject
	Write-Debug $("Subject emails = " + $EmailData.Count)

	Write-Verbose "Returning emails"
	Write-Debug $("There are " + $EmailData.Count + " emails that were returned")
	return $EmailData | Select-Object ReceivedTime, SenderName, Subject, Body | Sort-Object -Property "ReceivedTime" -Descending
}

function Set-LogPath {
	param (
		[string]$Path,
		[string]$FileType
	)

	[datetime]$currentDate = Get-Date

	Write-Verbose "Checking path $Path validity"

	if (!(Test-Path $Path)) {
		Write-Warning "Unable to use selected path. Setting log path to local folder"
		$Path = "."
	}
	else {
		Write-Verbose "$Path is valid"
	}

	switch ($FileType) {
		{ $_ -eq 'Csv' } { $ReturnPath = $Path + "\AdAttribute-DataFile-" + $currentDate.Year + "-" + $currentDate.Month + ".csv" }
		{ $_ -eq 'Log' } { $ReturnPath = $Path + "\AdAttribute-Log-" + $currentDate.ToString("s").Replace(":", ";") + ".log" }
		Default { 
			Write-Error "Error creating log file. Please make sure that the FileType you are sending is either 'Csv' or 'Log'"
			$Error 
			Read-Host
			exit
		}
	}

	Write-Verbose "File log is $ReturnPath"
	return $ReturnPath
}

function Get-Manager {
	param (
		[string]$Attribute,
		[string]$ID
	)
	
	Write-Verbose "Searching for manager..."

	Write-Debug "Manager's $Attribute should be $ID"
	$Manager = New-Object -TypeName Microsoft.ActiveDirectory.Management.ADAccount

	$Manager = Get-ADUser -Filter '$Attribute -eq $ID' -Properties DistinguishedName

	if ($Manager) {
		Write-Verbose $("$Manager was matched to $ID")
	}

	return $Manager
}

function Search-Managers {
	param (
		[System.Collections.ArrayList]$UserList,
		[string]$Manager
	)
	Write-Verbose "Filtering list by manager..."
	Write-Debug "Correct manager is $Manager"
	Write-Debug "Creating arraylist"
	$ReturnList = New-Object -TypeName System.Collections.ArrayList

	
	foreach ($user in $UserList) {
		Write-Debug $($user.SamAccountName)
		Write-Debug $("Manager " + $user.Manager)
		if ($user.Manager -eq $Manager) {
			$ReturnList.Add($user) | Out-Null
		}
		
	}
	# foreach ($item in $user.Keys) {
	# 	Write-Debug "Key: $item"
	# 	Write-Debug $("Value: " + $user[$item])
	# 	if ($user[$item] -eq $Manager)
	# }
	# $FilteredList = $UserList | Where-Object { Manager -eq $Manager }
	# foreach ($user in $FilteredList) {
	# 	Write-Debug $user.SamAccountName
		
	# }
	Write-Debug "Returning filtered list"
	return , $ReturnList
}

function Search-Location {
	param (
		[System.Collections.ArrayList]$UserList,
		[string]$Office
	)
	

	$ReturnList = New-Object -TypeName System.Collections.ArrayList

	Write-Verbose "Searching for the $Office location in the list..."
	$ReturnList = $UserList | Where-Object { $_.Office -like $("*" + $Office + "*") }
	return , $ReturnList
}

function Search-DisplayName {
	param (
		[string]$Name,
		[System.Collections.ArrayList]$Properties
	)
	
	Write-Verbose "Searching display names for $Name..."
	$hash = New-Object -TypeName hashtable
	$ReturnList = New-Object System.Collections.ArrayList

	$UserList = Get-ADUser -Filter { DisplayName -like $Name } -Properties $Properties | Select-Object $Properties
	
	foreach ($user in $UserList) {
		Write-Debug $("Adding user " + $user.SamAccountName + " to the return list")
		
		foreach ($prop in $Properties) {
			Write-Debug $("Adding $prop = " + $user.$prop + " to hashtable")
			$hash.Add($prop, $user.$prop)
		}
		$ReturnList.Add($hash) | Out-Null
	}
	Write-Debug $("Returning " + $ReturnList.Count + " users of type " + $ReturnList.GetType() + " from Search-DisplayName")
	
	
	return , [System.Collections.ArrayList]$ReturnList
	
}

function Search-BothNames {
	param (
		[System.Collections.ArrayList]$Properties,
		[string]$FirstName,
		[string]$LastName
	)

	$ReturnList = New-Object -TypeName System.Collections.ArrayList
	$gn = "*" + $FirstName + "*"
	$sn = "*" + $LastName + "*"

	
	Write-Verbose $("Searching for users with firstname " + $FirstName + " and lastname " + $LastName)
	$ReturnList += Get-ADUser -Filter { GivenName -like $gn -and Surname -like $sn } -Properties $Properties | Select-Object $Properties


	foreach ($pers in $ReturnList) {
		Write-Debug $pers.SamAccountName
	}

	Write-Debug ""
	Write-Debug $("Returning " + $ReturnList.Count + " users of type " + $ReturnList.GetType() + " from Search-BothNames")
	
	
	return , [System.Collections.ArrayList]$ReturnList
}


function Start-BasicSearch {
	param (
		[System.Collections.ArrayList]$Properties,
		[hashtable]$Attributes
	)

	Write-Debug "Calling Search-BothNames function"
	[System.Collections.ArrayList]$UserList = @()
	$ReturnHash = New-Object -TypeName hashtable
	$ReturnHash.Add("Reason", "")
	$UserList = ($(Search-BothNames -Properties $Properties -FirstName $Attributes.FN -LastName $Attributes.LN))
	Write-Debug $("Userlist returned with type " + $UserList.GetType())
	
	
	Write-Verbose "Checking number of accounts returned..."

	#Check 1 - FNLN
	switch ($UserList.Count) {

		1 {
			Write-Verbose "One user was returned"
			
			foreach ($prop in $Properties) {

				if ($UserList[0].$prop) {
					Write-Debug $("Adding $prop = " + $UserList[0].$prop + " to hashtable")
					$ReturnHash.Add($prop, $UserList[0].$prop)
				}
				else { 
					Write-Debug $("Adding $prop as empty value to hashtable")
					$ReturnHash.Add($prop, "") 
				}
			}

			return , [hashtable]$ReturnHash
		}


		0 { 
			Write-Verbose "No users were found with the matching names. Initiating display name search"
			$ReturnHash.Reason += "F&L Name: Failed - "
			Write-Debug "Calling Search-DisplayName function"
			[System.Collections.ArrayList]$FilteredList = @()
			$FilteredList.AddRange($(Search-DisplayName -Name $Attributes.Name -Properties $Properties))
			
			
			break
		}


		Default {
			Write-Verbose $($UserList.Count.ToString() + " users found")

			if ($Attributes.ManagerDN) {
				Write-Debug "Calling Search-Managers function"
				[System.Collections.ArrayList]$FilteredList = @()
				$FilteredList.AddRange($(Search-Managers -UserList $UserList -Manager $Attributes.ManagerDN))
				$ReturnHash.Reason += "Manager: "
			}
			else {
				Write-Debug "Calling Search-Location function"
				[System.Collections.ArrayList]$FilteredList = @()
				$FilteredList.AddRange($(Search-Location -Office $Attributes.Location -UserList $UserList))
				$ReturnHash.Reason += "Location: "
			}
		}
	}

	

	
	Write-Debug $("Starting second filter with " + $FilteredList.Count + " results")
	#Check 2 - FNLN to Manager/Location
	switch ($FilteredList.Count) {
		1 {
			Write-Verbose $("Filtered to " + $FilteredList.Count + " user")
			foreach ($prop in $Properties) {
				Write-Debug $("Current property $prop")
				if ($FilteredList[0].$prop -and $prop -notlike "*Manager*") {

					Write-Debug $("Adding $prop = " + $FilteredList[0].$prop + " to hashtable")
					$ReturnHash.Add($prop, $FilteredList[0].$prop)
				}
				elseif ($prop -notlike "*Manager*") { 
					$ReturnHash.Add($prop, "") 
				}
			}
			Write-Debug $("Hash is actually a " + $ReturnHash.GetType())
			break
		}

		0 {
			
			$ReturnHash.Reason += "Failed"
		
			Write-Warning $("No users were found with the matching names and Manager/Location. Please manually update " + $Attributes.Name)

			foreach ($prop in $Properties) {
				Write-Debug $("Adding blank $prop to hashtable")
				
				$ReturnHash.Add($prop, "") 
			}
			break
			
		}
		
		Default { 
			Write-Warning $("By some miracle, two people have the same name doing the same thing. Please manually update " + $Attributes.Name)
			foreach ($prop in $Properties) {
				Write-Debug $("Adding blank $prop to hashtable")
								
				$ReturnHash.Add($prop, "")
			}
			$ReturnHash.Reason = "Duplicate accounts"
		}
	}
	return , [hashtable]$ReturnHash
}

function Read-Email {
	param (
		$Email, $Json
	)
	
	Write-Verbose "`n"
	Write-Verbose "Reading email"
	$AttributeHash = New-Object -TypeName hashtable
	$EmailBodyArray = New-Object -TypeName System.Collections.ArrayList
	
	# foreach ($att in $Json.attributeArray) {
	# 	$AttributeHash.Add($att, "")
	# }
	
	#splits each email body up into an array of lines from the single massive string
	Write-Debug "Splitting up the email into an array"
	$EmailBodyArray = $Email.Body -split $([System.Environment]::NewLine)

	
	foreach ($Line in $EmailBodyArray) {
		#takes the input, $line, and checks it to see if it contains one of these phrases; Then splits it up to save only the unique data
		#Write-Debug $($Line -match $Json.delimiter)
		if ($Line -match $Json.delimiter) {
			Write-Debug "Splitting $Line on the delimiter"
			$AttributeKey, $AttributeValue = $Line -split $Json.delimiter

			#Write-Host $($Json.attributeArray -contains $AttributeKey) -ForegroundColor Cyan
			#Write-Debug "Attempting to match $AttributeKey"
			switch -wildcard ($AttributeKey) {

				$("*" + $Json.attributes.Name + "*") {
					Write-Debug "Setting name to $AttributeValue"
					$AttributeHash.Name = $AttributeValue
					#TODO - Should just add to the thing
					#$AttributeHash.Add($AttributeKey, $AttributeValue) | Out-Null
					break
				}
			
				$("*" + $Json.attributes.FN + "*") {
					Write-Debug "Setting first name to $AttributeValue"
					$AttributeHash.FN = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.LN + "*") {
					Write-Debug "Setting last name to $AttributeValue"
					$AttributeHash.LN = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.ID + "*") {
					Write-Debug "Setting ID to $AttributeValue"
					$AttributeHash.ID = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.PN + "*") {
					Write-Debug "Setting personnel number to $AttributeValue"
					$AttributeHash.PN = $AttributeValue
					break
				}

				$("*" + $Json.attributes.Title + "*") {
					Write-Debug "Setting title to $AttributeValue"
					$AttributeHash.Title = $AttributeValue
					break
				}

				$("*" + $Json.attributes.Type + "*") {
					Write-Debug "Setting type to $AttributeValue"
					$AttributeHash.Type = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.Location + "*") { 
					$AttributeValue, $null = $AttributeValue -split ", "
					Write-Debug "Setting location to $AttributeValue"
					$AttributeHash.Location = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.ManagerName + "*") {
					Write-Debug "Setting manager name to $AttributeValue"
					$AttributeHash.ManagerName = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.ManagerTitle + "*") {
					Write-Debug "Setting manager title to $AttributeValue"
					$AttributeHash.ManagerTitle = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.ManagerPN + "*") {
					Write-Debug "Setting manager personnel number to $AttributeValue"
					$AttributeHash.ManagerPN = $AttributeValue
					break
				}
			
				$("*" + $Json.attributes.ManagerID + "*") {
					Write-Debug "Setting manager ID to $AttributeValue"
					$AttributeHash.ManagerID = $AttributeValue
					break
				}
			
				Default {}
			}
		}
	}
	Write-Debug "Returning data from email `n"
	return , $AttributeHash
}
function Test-JSONData {
	param([Object]$JsonData)

	try {
		if (!($JsonData.throttleLimit)) {
			$JsonData.throttleLimit = ([int]$env:NUMBER_OF_PROCESSORS + 1)
			Write-Warning $("No throttle limit specified. Proceeding with default limit of" + ([int]$env:NUMBER_OF_PROCESSORS + 1))
		
		}
		if (!($JsonData.daysAgo)) {
			$JsonData.daysAgo = 30
			Write-Warning "No date filter specified. Proceeding with default date range of 30 days"
		
		}
		if (!($JsonData.emailSubject)) {
			$JsonData.emailSubject = $null
			Write-Warning "No 'Subject' filter specified"
		
		}
		if (!($JsonData.emailSender)) {
			$JsonData.emailSender = $null
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
		$Error
		Write-Error "Error validating data"
		Read-Host
		exit
	}

	Write-Debug $("")
	Write-Debug $("Throttle Limit: " + $JsonData.throttleLimit)
	Write-Debug $("Days Ago: " + $JsonData.daysAgo)
	Write-Debug $("Email Subject: " + $JsonData.emailSubject)
	Write-Debug $("Email Sender: " + $JsonData.emailSender)
	Write-Debug $("Property: " + $JsonData.property)
	Write-Debug $("Delimiter: " + $JsonData.delimiter)
	Write-Debug $("Searchbase: " + $JsonData.searchBase)
	Write-Debug $("Mailbox Name: " + $JsonData.mailboxName)
	Write-Debug $("Mailbox Folder: " + $JsonData.mailboxFolder)
	Write-Debug $("Log Path: " + $JsonData.logPath)
	Write-Debug $("")


	return $JsonData
}

function Set-ScriptMode {
	param (
		[string]$Mode
	)
	$Dev = New-Object -TypeName bool

	switch ($Mode) {
		'prod' {
			$DebugPreference = 'silentlycontinue'
			$VerbosePreference = 'silentlycontinue'
			$WarningPreference = 'continue'
			$InformationPreference = 'continue'
			$ProgressPreference = 'continue'
		}
		'dev' {
			$DebugPreference = 'continue'
			$VerbosePreference = 'continue'
			$WarningPreference = 'continue'
			$InformationPreference = 'continue'
			$ProgressPreference = 'continue'
			$Dev = $true
		}
		
		Default {
			$DebugPreference = 'silentlycontinue'
			$VerbosePreference = 'silentlycontinue'
			$WarningPreference = 'continue'
			$InformationPreference = 'silentlycontinue'
			$ProgressPreference = 'continue'
		}
	}
	return $Dev
}
#endregion

#region initialization
$Error.Clear()
try {
	Import-Module -Name ActiveDirectory -Force
}
catch {
	Write-Error "Unable to import ActiveDirectory module. Please make sure it is installed before proceeding"
	$Error
	Read-Host
	exit
}


Write-Information "Getting config.json"
try {
	[Object]$JsonData = Get-Content "config.json" | ConvertFrom-Json
}
catch {
	Write-Error "Unable to get config.json. Please make sure it is located in the same location as the script"
	$Error
	Read-Host
	exit
}

try {
	[bool]$Dev = Set-ScriptMode $JsonData.mode
	if ($Dev) {
		Write-Information "Dev mode initiated. Getting content from dev.json"
		$JsonData = Get-Content "dev.json" | ConvertFrom-Json
	}
}
catch {
	$Error
	Write-Error "Dev mode initiation failed" -ErrorAction Suspend
	$Error.Clear()
}


Write-Information "Initializing Script"
#endregion

#region transcript
try {
	Start-Transcript -Path $(Set-LogPath -Path $JsonData.logPath -FileType 'Log') -Force -NoClobber
}
catch {
	$Error
	Write-Error "Unable to create log"
	Read-Host
	exit
}

Write-Verbose "Transcript started"

#endregion

#region data verification
[object]$Json = Test-JSONData -JsonData $JsonData
if ($Dev) {
	try {
		Write-Verbose "Removing csv"
		Remove-Item $(Set-LogPath -Path $JsonData.logPath -FileType 'Csv') -Force
	}
	catch {  }
}
#endregion


#region variables
Write-Verbose "Initializing variables..."
$EmailData, $PropertyArray = New-Object -TypeName System.Collections.ArrayList
$AttributeHash, $ReturnHash = New-Object -TypeName hashtable
$ProgCount, $EmailCount = New-Object -TypeName int
$PropertyArray = $($Json.'property'), 'DisplayName', 'SamAccountName', 'GivenName', 'Surname', 'Manager', 'Office', 'Created'
#endregion


#region assignments
Write-Output "Importing Outlook data..."
$EmailData = Import-OutlookData -SearchData $Json
$EmailCount = $EmailData.Count
Write-Debug "Final email count: $EmailCount"
$ProgCount = 0
#endregion


#starts checking each email one at a time
$CsvData = foreach ($Email in $EmailData) {

	Write-Verbose ""
	Write-Verbose "************************************************************************************************************************************"
	Write-Verbose ""

	Write-Debug "Clearing hashtable"
	$AttributeHash.Clear()
	$attSet, $matchFound = New-Object -TypeName bool

	Write-Debug "Incrementing $ProgCount by one"
	$ProgCount++

	if ($Dev) {
		Write-Information $("Started processing email $ProgCount out of $EmailCount")
	}
 else {
		Write-Progress $("Started processing email $ProgCount out of $EmailCount")
	}


	Write-Verbose $("Calling Read-Email function")
	$AttributeHash = Read-Email -Email $Email -Json $Json
	
	foreach ($item in $AttributeHash.Keys) {
		Write-Debug $("Key: $item `t|`tValue: " + $AttributeHash[$item])
	}
	

	Write-Debug $("Adding 'EmailReceivedTime = " + $Email.ReceivedTime + " to the hashtable")
	$AttributeHash.Add("EmailReceivedTime", $Email.ReceivedTime)


	Write-Debug "Calling Get-Manager function"
	[Microsoft.ActiveDirectory.Management.ADAccount]$Manager = Get-Manager -Attribute $Json.'property' -ID $AttributeHash.ManagerID
	
	if ($Manager) {
		Write-Debug $("Adding manager Distinguished Name to hashtable as " + $Manager.DistinguishedName)
		$AttributeHash.Add("ManagerDN", $Manager.DistinguishedName)
	}
	else {
		Write-Warning $("No manager was found with " + $Json.property + " " + $AttributeHash.ManagerID)
		$AttributeHash.Add("ManagerDN", "")
	}
	
	#IDEAS Maybe do $i++ for each single check, then proceed if it passes multiple stages?
	#$ADUser = New-Object -TypeName Microsoft.ActiveDirectory.Management.ADAccount

	Write-Debug "Calling Start-BasicSearch function"
	$ReturnHash = [hashtable]($(Start-BasicSearch -Attributes $AttributeHash -Properties $PropertyArray))


	Write-Debug $("Received type " + $ReturnHash.GetType() + " from the search function")
		
		
	foreach ($line in $ReturnHash.Keys) {
		Write-Debug $("Adding $line = " + $ReturnHash[$line] + " to AttributeHash")
		
		$AttributeHash.Add($line, $ReturnHash[$line])
	}

	
	if ($AttributeHash.SamAccountName) {

		Write-Verbose $("Matched " + $AttributeHash.Name + " to " + $AttributeHash.SamAccountName)
		$matchFound = $true
		Write-Debug $($AttributeHash.($Json.'property'))

		if ($AttributeHash.($Json.'property')) {
			$AttributeHash.Reason = $Json.'property' + " already set for " + $AttributeHash.SamAccountName
			Write-Host $AttributeHash.Reason -ForegroundColor Cyan
			
		}

		elseif (!($AttributeHash.($Json.'property'))) {

			Write-Debug $("Setting " + $AttributeHash.SamAccountName + " " + $Json.'property' + " to " + $AttributeHash.ID )

			try {
				if($Dev){
					Set-ADUser $AttributeHash.SamAccountName -Add @{$($Json.'property') = $AttributeHash.ID } -WhatIf
				} else {
					Set-ADUser $AttributeHash.SamAccountName -Add @{$($Json.'property') = $AttributeHash.ID }
				}
				

				$attSet = $true
				$AttributeHash.Reason = ""
				Write-Verbose $("Set " + $Json.'property' + " to " + $AttributeHash.ID + " for " + $AttributeHash.SamAccountName)
			}
			catch {
				$AttributeHash.Reason = $("Unable to set " + $Json.'property' + " for " + $AttributeHash.SamAccountName)
				Write-Error $AttributeHash.Reason
				
			}
		}
	}
	else {
		$matchFound = $false
	}
	$AttributeHash.Add("Matched", $matchFound)
	$AttributeHash.Add("Modified", $attSet)


	$PSCustomObject = [PSCustomObject]@{}
	
	foreach ($item in $AttributeHash.Keys) {
		Write-Debug $("Key: $item `t|`tValue: " + $AttributeHash[$item])
		
	}
	#Write-Debug $($AttributeHash | Format-Table)
	#TODO 
	$PSCustomObject = [pscustomobject]@{
		"Email Date"                         = $AttributeHash.EmailReceivedTime
		"Email Employee Name"                = $AttributeHash.Name
		"Email Employee First Name"          = $AttributeHash.FN
		"Email Employee Last Name"           = $AttributeHash.LN
		"Email Employee ID"                  = $AttributeHash.ID
		"Email Employee Personnel Number"    = $AttributeHash.PN
		"Email Employee Job Title"           = $AttributeHash.Title
		"Email Employee Type"                = $AttributeHash.Type
		"Email Employee Location"            = $AttributeHash.Location
		"Email Manager Name"                 = $AttributeHash.ManagerName
		"Email Manager Job Title"            = $AttributeHash.ManagerTitle
		"Email Manager Personnel Number"     = $AttributeHash.ManagerPN
		"Email Manager ID"                   = $AttributeHash.ManagerID
		" "                                  = " "
		"AD Account Matched"                 = $AttributeHash.Matched
		"AD Display Name"                    = $AttributeHash.DisplayName
		"AD Given Name"                      = $AttributeHash.GivenName
		"AD Surname"                         = $AttributeHash.Surname
		"AD Location"                        = $AttributeHash.Office
		"AD Username"                        = $AttributeHash.SamAccountName
		"AD Created Date"                    = $AttributeHash.Created
		"Manager DistinguishedName"          = $AttributeHash.ManagerDN
		$($Json.property + " set by script") = $AttributeHash.Modified
		"Reason"                             = $AttributeHash.Reason
	}
	$PSCustomObject
}

try {
	Write-Information "All emails processed. Exporting data to Csv"

	Write-Verbose "Calling Set-LogPath function"
	$CsvData | Select-Object -Property * | Export-Csv -Path $(Set-LogPath -FileType Csv -Path $Json.'logPath') -NoTypeInformation -Append
	Stop-Transcript
} catch {
	$Error
	Write-Error "Unable to export to Csv" -ErrorAction Suspend
	Stop-Transcript
}


if (!($Dev)) {
	Read-Host "Done ("
}