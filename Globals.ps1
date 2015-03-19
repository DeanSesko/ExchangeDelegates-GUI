function Get-ScriptDirectory {
	if ($hostinvocation -ne $null) {
		Split-Path $hostinvocation.MyCommand.path
	}
	else {
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}
[string]$ScriptDir = Get-ScriptDirectory
[string]$ScriptDir += "\"

$ExcludedUsersFile = "exclude.txt"
$Global:records = @()
$Global:mailboxData = @()

# Export File for Mailbox and Delegate relationship
$delegateOutputfileName = "Delegates.csv"

#Export File used to import mailbox information
$mailboxExportFileName = "Mailboxinfo.csv"

#DataTable setup
$TempTable = New-Object System.Data.DataTable
$dataTable = New-Object System.Data.DataTable
$MigrationTable = New-Object System.Data.DataTable
$StatsTable = New-Object System.Data.DataTable
$mailboxinfoTable = New-Object System.Data.DataTable
$DelegateTable = New-Object System.Data.DataTable
$Sdate = get-date
$index = 0
$Global:output = @()
$global:FAdelegateCount = 0
$global:SADelegateCount = 0
$global:SOBDelegateCount = 0

#Setup Logs
$DelegateOutput = "$ScriptDir$delegateOutputfileName"
$MailboxOutPut = "$ScriptDir$mailboxExportFileName"
[int]$Global:groupcountName = 1
$MigrationGroupPrefix = "MigrationGroup"
$DataOutputFile = "MigrationGroups.csv"
$DataOutput = "$ScriptDir$DataOutputFile"
$DBInput = "Delegates.csv"
$mailboxinput = "Mailboxinfo.csv"


#Test For input Files



#Import mailbox information



Function LoganEntry($LogEntry) {
	$MainForm.Controls['outputbox'].lines += $LogEntry
	[System.Windows.Forms.Application]::DoEvents()
}
Function SetupDataTables {
	
	$mailboxinfoTable.Columns.Add("SamAccountName") | Out-Null
	$mailboxinfoTable.Columns.Add("PrimarySMTPAddress") | Out-Null
	$mailboxinfoTable.Columns.Add("DN") | Out-Null
	$mailboxinfoTable.Columns.Add("ItemCount") | Out-Null
	$mailboxinfoTable.Columns.Add("Size") | Out-Null
	$mailboxinfoTable.Columns.Add("DB") | Out-Null
	$mailboxinfoTable.Columns.Add("Server") | Out-Null
	$pk = $mailboxinfoTable.Columns["SamAccountName"]
	$pk1 = $mailboxinfoTable.Columns["DN"]
	$mailboxinfoTable.PrimaryKey = $pk, $pk1
	$DelegateTable.Columns.Add("MailBox") | Out-Null
	$DelegateTable.Columns.Add("Delegate") | Out-Null
	$DelegateTable.Columns.Add("Permission") | Out-Null
	$MigrationTable.Columns.Add("Mailbox") | Out-Null
	$MigrationTable.Columns.Add("MigrationGroup") | Out-Null
	$MigrationTable.Columns.add("MailboxSize") | Out-Null
	$pk2 = $MigrationTable.Columns["Mailbox"]
	$MigrationTable.PrimaryKey = $pk2
	$dataTable.Columns.Add("Mailbox") | Out-Null
	$dataTable.Columns.Add("Delegate") | Out-Null
	$StatsTable.Columns.Add("Mailbox") | Out-Null
	$statsTable.Columns.Add("ItemCount") | Out-Null
	$statsTable.Columns.Add("Size") | Out-Null
	$statsTable.Columns.Add("DB") | Out-Null
	$statsTable.Columns.Add("Server") | Out-Null
	$pk3 = $statsTable.Columns["Mailbox"]
	$statsTable.PrimaryKey = $pk3
}
Function GetnewCreds {
	
	$OnPremAdmin = $MainForm.Controls['AdminTextBox'].Text
	$OnPrempassword = convertto-securestring  $MainForm.Controls['AdminPasswordTextBox'].text -asplaintext -force
	$global:ExchServer = $MainForm.Controls['ExchangeServerTextBox'].Text
	$global:OnPremcred = new-object -typename System.Management.Automation.PSCredential -argumentlis $OnPremAdmin, $OnPrempassword
	
}
Function ConnectLocalService {
	cls
	Try {
		$LocalSession = New-PSSession -Name "ON-Prem-Exchange" -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchServer/PowerShell/" -Authentication kerberos -Credential $OnPremcred -ErrorAction stop
		Import-PSSession $LocalSession -DisableNameChecking -WarningAction silentlyContinue | Out-Null
		return $true
	}
	Catch {
		cls
		LoganEntry ""
		LoganEntry ""
		LoganEntry "Access Denied.  Check your UserName / Password and Try Again"
		LoganEntry ""
		LoganEntry ""
		LoganEntry ""
	}
	Finally { }
}
Function CheckExclude($User) {
	
	if ($ExcludePerms -contains $user) {
		Return $True
	}
	else {
		Return $false
		
	}
	
	#End Function
}
Function FillDelegatetable($mailbox, $delegate, $perm) {
	
	$row = $DelegateTable.NewRow()
	$row["MailBox"] = $mailbox
	$row["Delegate"] = $delegate
	$row["Permission"] = $perm
	$DelegateTable.Rows.Add($row)
	
	#End Function
}
Function FillProcessDelegatesDataTables {
	$dataTable.Clear()
	FOREACH ($RECORD IN $Global:records) {
		$row = $dataTable.NewRow()
		$row["Mailbox"] = $RECORD.Mailbox
		$row["Delegate"] = $RECORD.Delegate
		$dataTable.Rows.Add($row)
		
	}
	
	FOREACH ($MBXRecord IN $Global:mailboxData) {
		
		$row = $statsTable.NewRow()
		$row["Mailbox"] = $MBXRecord.PrimarySMTPAddress
		$row["ItemCount"] = $MBXRecord.ItemCount
		$row["Size"] = $MBXRecord.size
		$row["DB"] = $MBXRecord.DB
		$row["Server"] = $MBXRecord.Server
		$statsTable.Rows.Add($row)
		
	}
	
	#end Function
}
function GetImportFiles {
	$ImportSuccess = $false
	if (test-path "$ScriptDir$DBInput") {
		$Global:records = import-csv "$ScriptDir$DBInput "
	}
	if (test-path "$ScriptDir$mailboxinput") {
		$Global:mailboxData = import-csv "$ScriptDir$mailboxinput"
	}
	if ($records -ne $null -and $mailboxData -ne $null) {
		$ImportSuccess = $True
	}
	Else {
		LoganEntry " Missing DB Input File. Please FInd Delegates Again"
		$ImportSuccess = $false
	}
	
	return $ImportSuccess
	
}
Function FillMailboxinfoDataTable {
	
	foreach ($MBX in $MyMailboxes) {
		if ($ProcessStats) {
			#Gather Mailbox Statistics
			$Stat = GetStats($mbx)
		}
		if ($stat) {
			$global:IC = [string]$Stat.ItemCount.tostring()
			$global:size = [string]$Stat.Size.tostring()
		}
		
		Else {
			$global:IC = "0"
			$global:size = "0"
		}
		
		$MBXAddress = GetSmtpAddress($mbx)
		$row = $mailboxinfoTable.NewRow()
		$row["SamAccountName"] = $MBX.SamAccountName
		$row["PrimarySMTPAddress"] = $MBXAddress
		$row["DN"] = [string]$path = $MBX.OrganizationalUnit + "/" + $MBX.Name
		$row["ItemCount"] = $global:IC
		$row["Size"] = $global:size
		$row["DB"] = $MBX.Database
		$row["Server"] = $MBX.ServerName
		$mailboxinfoTable.Rows.Add($row)
	}
	
	#End Function
}
Function GetSmtpAddress($mbx) {
	if ($mbx.primarySMTPAddress -eq $null) {
		$mbx.EmailAddresses | ForEach {
			If ($_.Prefix -match "SMTP") {
				$SmtpAddress = $_.SmtpAddress
			}
		}
	}
	
	Else {
		$SmtpAddress = $mbx.primarysmtpAddress.tostring()
	}
	return $SmtpAddress
	#End Function
}
Function GetFullAccessDelegate($mbx) {
	$results = Get-MailboxPermission $MBX.Name -erroraction Silentlycontinue -WarningAction silentlyContinue | Where { ($_.AccessRights -eq “FullAccess”) -and -not ($_.isInherited -like "True") } | Select User
	if ($results) {
		$global:FAdelegateCount++
		foreach ($result in $results) {
			$ExludedUser = CheckExclude($result.user.tostring())
			if (!($ExludedUser)) {
				$RowDelegate = $null
				$Delegate = $null
				$SamAccountDelegate = $result.user.tostring().split("\")[1]
				[string]$filter = " SamAccountName = '" + $SamAccountDelegate + "'"
				Try {
					$RowDelegate = $mailboxinfoTable.Select($filter)
					$Delegate = $RowDelegate[0]["PrimarySMTPAddress"].ToString()
					FillDelegatetable $MBXAddress $Delegate "FullAccess" $global:ic  $global:size  $db $server
				}
				Catch { }
				Finally { }
			}
		}
	}
	
	else {
		$Delegate = "NA"
		FillDelegatetable $MBXAddress $Delegate "FullAccess" $global:ic  $global:size  $db $server
	}
	
	#End FUnction
}
Function GetSendONDelegate($mbx) {
	
	$delegateINfo = @()
	$global:SOBDelegateCount++
	foreach ($delegate in $mbx.GrantSendOnBehalfTo) {
		$RowDelegate = $null
		$MyDelegate = $null
		[string]$filter = " DN = '" + $delegate + "'"
		Try {
			$RowDelegate = $mailboxinfoTable.Select($filter)
			$MyDelegate = $RowDelegate[0]["PrimarySMTPAddress"].ToString()
		}
		Catch { }
		Finally { }
		FillDelegatetable $MBXAddress $MyDelegate "SendonBeHalfOf"
	}
	#End Function
}
fUNCTION GetSendAsDelegate($SamAcct) {
	
	$FoundDelegate = $false
	$strFilter = "(&(objectCategory=User)(SamAccountName=$SamAcct))"
	$objDomain = New-Object System.DirectoryServices.DirectoryEntry
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objSearcher.SearchRoot = $objDomain
	$objSearcher.PageSize = 1000
	$objSearcher.Filter = $strFilter
	$objSearcher.SearchScope = "Subtree"
	$colResults = $objSearcher.Findone()
	$usercount = $colResults.Count
	foreach ($objResult in $colResults) {
		$objItem = $objResult.Properties
		ForEach ($address in $objItem.proxyaddresses) {
			If ($address.StartsWith("SMTP:")) {
				$SmtpAddress = $address.tostring().split(":")[1]
			}
		}
		$Name = $objItem.distinguishedname
		$ADObject = [ADSI]"LDAP://$Name"
		$aclObject = $ADObject.psbase.ObjectSecurity
		$aclList = $aclObject.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])
		foreach ($acl in $aclList) {
			
			$objSID = New-Object System.Security.Principal.SecurityIdentifier($acl.IdentityReference)
			if (($acl.ObjectType -eq "ab721a54-1e2f-11d0-9819-00aa0040529b") -and -not ($acl.IdentityReference -like "S-1-5-10") -and -not ($acl.IdentityReference -like "S-1-5-0") -and -not ($acl.IdentityReference -like "S-1-5-7")) {
				$FoundDelegate = $true
				Try {
					[string]$obj = $objSID.Translate([System.Security.Principal.NTAccount])
					$SamAccountDelegate = $obj.tostring().split("\")[1]
					[string]$filter = " SamAccountName = '" + $SamAccountDelegate + "'"
					$EnterDB = $true
				}
				Catch { $EnterDB = $False }
				Finally { }
				
				Try {
					if ($EnterDB) {
						$ExludedUser = CheckExclude($obj.tostring())
						if (!($ExludedUser)) {
							$RowDelegate = $mailboxinfoTable.Select($filter)
							$Delegate = $RowDelegate[0]["PrimarySMTPAddress"].ToString()
							FillDelegatetable $MBXAddress $Delegate "SendAS"
							
						}
					}
				}
				Catch { }
				Finally { }
			}
		}
		
	}
	
	if ($FoundDelegate) {
		$global:SADelegateCount++
	}
	
	Return $nothing
	#End Function
}
Function GetAllMailboxes {
	$Mailboxes = get-mailbox -resultsize unlimited -WarningAction silentlyContinue | select primarySMTPAddress, Database, Name, ServerName, GrantSendOnBehalfTo, SamAccountName, OrganizationalUnit
	Return $Mailboxes
	#End Function
}
Function ProcessMigrationFile {
	[String]$Global:MigrationGroupName = "$MigrationGroupPrefix $Global:groupcountName"
	foreach ($rec in $mailboxes) {
		
		if ($rec.contains("'")) {
			$myuser = $rec.Replace("'", "''")
		}
		Else { $myuser = $rec }
		[String]$StatString = "Mailbox = '$myuser'"
		foreach ($Statuser in $StatsTable.Select($StatString)) { $Statsize = $Statuser.size }
		$row = $MigrationTable.NewRow()
		$row["Mailbox"] = $myuser
		$row["MigrationGroup"] = $Global:MigrationGroupName
		$row["MailboxSize"] = $Statsize
		try {
			$MigrationTable.Rows.Add($row)
		}
		Catch { }
		Finally { }
		
		
		foreach ($row in $dataTable.Select("Mailbox like '$myuser' or Delegate like '$myuser'")) {
			If ($row.Item("Delegate") -ne "NA") {
				SearchMyDT $row.Delegate
			}
		}
	}
	
	
	
	do {
		$TempTable = $MigrationTable.Copy()
		foreach ($Row in $temptable) {
			SearchMyDT $row.Mailbox
		}
		
	}
	while ($temptable.rows.count -ne $MigrationTable.rows.Count)
	
	
	#End function
}
Function GetMigrationGroups {
	[String]$Global:MigrationGroupName = "$MigrationGroupPrefix $Global:groupcountName"
	foreach ($user in $mailboxData) {
		$MBXAddress = $user.primarysmtpAddress
		
		if ($MBXAddress.contains("'")) {
			$myuser = $MBXAddress.Replace("'", "''")
		}
		Else { $myuser = $MBXAddress }
		try {
			$row = $MigrationTable.NewRow()
			$row["Mailbox"] = $MBXAddress
			$row["MigrationGroup"] = $Global:MigrationGroupName
			$row["MailboxSize"] = $user.size
			$MigrationTable.Rows.Add($row)
			
		}
		
		Catch { }
		
		Finally { }
		
		
		foreach ($row in $dataTable.Select("Mailbox like '$myuser' or Delegate like '$myuser'")) {
			
			If ($row.Item("Delegate") -ne "NA") {
				SearchMyDT $row.Delegate
			}
			
		}
		
		do {
			$TempTable = $MigrationTable.Copy()
			foreach ($Row in $temptable) { SearchMyDT $row.Mailbox }
		}
		
		while ($temptable.rows.count -ne $MigrationTable.rows.Count)
		
		$RowCounter = 0
		foreach ($RC in $MigrationTable.Select("MigrationGroup like '$Global:MigrationGroupName' ")) { $RowCounter++ }
		
		if ($RowCounter -ge $GroupMaxCount) {
			$Global:groupcountName++
			[String]$Global:MigrationGroupName = "$MigrationGroupPrefix $Global:groupcountName"
		}
	}
	#End Function
}
Function SearchMyDT ($User) {
	[System.Windows.Forms.Application]::DoEvents()
	if ($user.contains("'")) {
		$myuser = $User.Replace("'", "''")
	}
	Else { $myuser = $user }
	
	[string]$mystring = "Mailbox like '$myuser'  or Delegate like '$myuser'"
	
	ForEach ($row3 In $dataTable.Select($mystring)) {
		
		
		Try {
			$tempuser = $row3.Item("Mailbox")
			[String]$StatString = "Mailbox = '$tempuser'"
			foreach ($Statuser in $StatsTable.Select($StatString)) { $Statsize = $Statuser.size }
			$Myrow = $MigrationTable.NewRow()
			$Myrow["Mailbox"] = $row3.Item("Mailbox")
			$Myrow["MigrationGroup"] = $Global:MigrationGroupName
			$Myrow["MailboxSize"] = $Statsize
			$MigrationTable.Rows.Add($Myrow)
			
			
		}
		Catch { }
		Finally { }
		
		
		If ($row3.Item("Delegate").ToString() -ne "NA") {
			
			
			Try {
				$tempuser = $row3.Item("Delegate")
				[String]$StatString = "Mailbox = '$tempuser'"
				foreach ($Statuser in $StatsTable.Select($StatString)) { $Statsize = $Statuser.size }
				$Myrow = $MigrationTable.NewRow()
				$Myrow["Mailbox"] = $row3.Item("Delegate")
				$Myrow["MigrationGroup"] = $Global:MigrationGroupName
				$Myrow["MailboxSize"] = $Statsize
				$MigrationTable.Rows.Add($Myrow)
				
			}
			
			
			
			Catch { }
			Finally { }
		}
		
	}
	
}
Function GetMailboxesWithNoDelegates {
	
	foreach ($user in $mailboxData) {
		$MBXAddress = $user.primarysmtpAddress
		if ($MBXAddress.contains("'")) {
			$myuser = $MBXAddress.Replace("'", "''")
		}
		Else { $myuser = $MBXAddress }
		
		[string]$mystring = "Mailbox like '$myuser' and Delegate like 'NA'"
		$MyRecordSet = $dataTable.Select($mystring)
		$AnyMigUser = @()
		
		if ($MyRecordSet.Count -eq 1) {
			
			$AnyMigUser = $dataTable.Select("Delegate = '$myuser'")
			
			if ($AnyMigUser.Count -eq 0) {
				
				$row = $MigrationTable.NewRow()
				$row["Mailbox"] = $MBXAddress
				$row["MigrationGroup"] = "Any Migration Group"
				$row["MailboxSize"] = $user.size
				$MigrationTable.Rows.Add($row)
				
				
			}
			
		}
		
	}
	
	#End Function
}
Function GetStats($mbx) {
	$MBXSTAT = Get-MailboxStatistics $MBX.Name -erroraction Silentlycontinue -WarningAction SilentlyContinue | select   @{ label = "Size"; expression = { $_.TotalItemSize } }, itemcount
	
	if ($mbxstat -ne $null) {
		return $MBXSTAT
	}
	
	#End Function
}
Function ProcessUsers {
	foreach ($rec in $MigrationUsers) {
		$row = $MigTable.NewRow()
		$row["Mailbox"] = $Rec
		
		try {
			$MigTable.Rows.Add($row)
		}
		
		Catch { }
		Finally { }
		foreach ($row2 in $dataTable.Select("Mailbox like '$rec' or Delegate like '$rec'")) {
			If ($row2.Item("Delegate") -ne "NA") { SearchMyDT $row2.Delegate }
		}
	}

}
Function GetMailboxPerf($MBXUsercOUNT) {
	
	$MBXseconds = (New-TimeSpan -Start $sdate -End $MbxProcStartTime).totalseconds
	$MBXPTime = New-TimeSpan -Seconds $MBXseconds
	$MBXPRoxTime = '{0:00}:{1:00}:{2:00}' -f $MBXPTime.Hours, $MBXPTime.Minutes, $MBXPTime.Seconds
	
	
	LoganEntry "-------------------------------------------------------------"
	LoganEntry "Found $MBXUsercOUNT Mailboxes in $MBXPRoxTime  "
	LoganEntry "-------------------------------------------------------------"
	
	LoganEntry "Started Proccesing Mailboxes at: $MbxProcStartTime "
	
	#End Function
}
Function UpdateProgressBar($Users, $ProgressIndex, $Progress) {
	
	$progress = ($ProgressIndex/$Users) * 100
	$FP = "{0:N2}" -f $progress
	$EDate = get-date
	$seconds = (New-TimeSpan -Start $MbxProcStartTime -End $EDate).totalseconds
	$MBXPerSec = $seconds / $ProgressIndex
	$MBXPS = "{0:N2}" -f $MBXPerSec
	$MainForm.Controls['MainProgressBar'].Maximum = 100
	$MainForm.Controls['MainProgressBar'].Value = $FP
	Return $MBXPS
	
}
Function ProcessDelegates {
	$Progress = 0
	$ProgressIndex = 0
	
	foreach ($MBX in $MyMailboxes) {
		$MBXAddress = GetSmtpAddress($mbx)
		$db = $mbx.Database.tostring()
		$server = $mbx.ServerName.tostring()
		$FullAccessDelegateSMTP = GetFullAccessDelegate($mbx)
		if ($mbx.GrantSendOnBehalfTo) {
			$SendonBehalfoDelegateSMTP = GetSendONDelegate($mbx)
		}
		
		GetSendAsDelegate($mbx.SamAccountName)
		$ProgressIndex++
		$ProgressUpdate = UpdateProgressBar $MyMailboxes.count $Progressindex $Progress
		#End Function
	}
	
	Return $ProgressUpdate
}
Function FindDelegatesCloseOut {
	$EDate = get-date
	$mbxTime = $MyMBXPERSEC.trim()
	LoganEntry ""
	LoganEntry "Finsihed Gathering Delegate Information"
	LoganEntry "---------------------------------------------------------------"
	LoganEntry "Process Started at:  $sdate  "
	LoganEntry "Process ended at:  $EDate  "
	LoganEntry ""
	LoganEntry "Found $global:FAdelegateCount Delegates with Full Access Permissions  "
	LoganEntry "Found $SAdelegateCount Delegates with Send As Permissions"
	LoganEntry "Found $global:SOBDelegateCount Delegates with Send on Be Half Of Permissions  "
	LoganEntry ""
	LoganEntry "Mailboxes Per Second $mbxTime"
	LoganEntry "---------------------------------------------------------------"
	LoganEntry ""
	$DelegateTable | Export-csv $DelegateOutput -NoTypeInformation
	$mailboxinfoTable | Export-csv $MailboxOutPut -NoTypeInformation
	#End Function
}
Function ProcessDelegatesStartup {
	$Sdate = get-date
	LoganEntry "Importing Data"
	LoganEntry ""
	LoganEntry "Connecting to Databases and Analyzing Environment"
	LoganEntry ""
	LoganEntry " Start Time: $sdate"
	LoganEntry ""
}
Function ProcessDelegatesCloseout {
	$EDate = get-date
	LoganEntry ""
	LoganEntry "Finished Gathering  Information"
	LoganEntry ""
	LoganEntry "End Time: $EDate"
	LoganEntry ""
	LoganEntry ""
	$MainForm.Controls['datagridview1'].dataSource = $MigrationTable
	$MigrationTable | export-csv $DataOutput -NoTypeInformation
	if (Test-Path $DataOutput) {
		$MainForm.Controls['exportlabel'].text = "$DataOutput"
		
	}
	
	
}