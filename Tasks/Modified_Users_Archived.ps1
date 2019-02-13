<#
Created by Victor Sanchez on April 28 2016
Query all users in the Archive database that are not disabled and query if they are enabled or disabled in AD
Query all users in OU=Archive and move if not not in Archive Database
Query all users in OU-Dead Accounts and modify user, Set to hidden and Set out of office message
#>

Import-Module activedirectory

if ((get-pssession |% {$_.name}) -notcontains "Exch2013"){
Write-Host "Load Exchange 2013 Session"
$connectionUri = "http://prodexchapp20.repsrv.com/powershell?serializationLevel=Full"
New-PSSession -name Exch2013 -ConnectionURI "$connectionUri" -ConfigurationName Microsoft.Exchange
$session = Get-PSSession -Name Exch2013
Import-PSSession $session -AllowClobber
}


# New-EventLog -LogName AD_Changes -Source AD_Changes
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Start AD Import!!!"
#    $today=(Get-date).AddDays(-1).ToString('yyyyMMdd')
$today=Get-date -Format yyyyMMdd
$file="\\repsharedfs\Share\axwayftp\adreports\Outgoing\" + $today + "-ADExport.csv"
# $file="\\repsharedfs\Share\axwayftp\adreports\Outgoing\20180214-ADExport.csv"
$adusers = Import-Csv -Path $file # | Where-Object {$_.dn -like "*archive*"}

<#
Query Users where OU=Archive, if not in the Year OU then move to OU=2018
#>
$Archived = $ADusers | where {$_.distinguishedname -like "*archive,DC*"}
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in Archive OU"
$NotinYearFolder = $Archived | where{$_.distinguishedname -notlike "*OU=20*" -and $_.distinguishedname -notlike "*OU=archiveusers*"}
#$NotinYearFolder = $Archived | where{$_.distinguishedname -like "*OU=2013*"}
foreach ($NotInOU in $NotinYearFolder)    {   
    $deadsam = $NotInOU.samaccountname
    $DeadSam # ="espeybe"
    Set-Mailbox $deadsam  -MaxReceiveSize 10KB
    Set-ADUser –Identity $Deadsam -Clear "extensionattribute3" #Area
    Set-ADUser –Identity $Deadsam -Clear "extensionattribute12" #Region
    Set-ADUser –Identity $Deadsam -Clear "extensionattribute13" #Job Code
    Set-ADUser –Identity $Deadsam -Clear "extensionattribute14" #BU
    Set-ADUser –Identity $Deadsam -Clear "extensionattribute11" # Division
    Set-User $deadsam -Fax $null
    Set-User $deadsam -Title $null
    Move-ADObject $NotInOU.distinguishedname -TargetPath "OU=2019,OU=Archive,DC=repsrv,DC=com"
        }

# Cleanup of attributes of Dead Acounts/Users, 2/12/2018 2,752 accounts not modified and moved to OU=cleared
$DeadAccounts = $ADusers | where {$_.distinguishedname -like "*ou=users,ou=dead*"}
$DeadAccounts = $DeadAccounts | where{$_.distinguishedname -notlike "*OU=Cleared*"}
$DeadAccounts = $DeadAccounts | where{$_.distinguishedname -notlike "*OU=AdminLeave*"}
$DeadAccounts.count
$DeadAccounts | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\Engineering\PS\data\Changes\$today"_DeadAccount_CleanupAttribute.csv"
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in OU=Dead Accounts"
foreach ($Dead in $DeadAccounts)    {   
    $deadsam = $Dead.samaccountname
    $DeadSam
    $dead.DistinguishedName
    #Set-Mailbox $deadsam  -MaxReceiveSize 1KB
    Set-Mailbox $deadsam -CustomAttribute3 $null #Area
    Set-Mailbox $deadsam -CustomAttribute12 $null #Region
    Set-Mailbox $deadsam -CustomAttribute13 $null #Job Code
    Set-Mailbox $deadsam -CustomAttribute14 $null #BU
    Set-Mailbox $deadsam -CustomAttribute11 $null #Division
    Set-User $deadsam -Fax $null
    Set-User $deadsam -Title $null
    Move-ADObject $Dead.distinguishedname -TargetPath "OU=Cleared,OU=Users,OU=Dead Accounts,DC=repsrv,DC=com"
    #Start-Sleep 1
           }
#>

foreach($arc in $archived){
    $arcsam=$arc.samaccountname
    $notArcDB=Get-Mailbox $arcsam | where{$_.database -notlike "*repsrvarc0*"} | select name, alias, database 
    if ($notArcDB -ne $null) {Write-Host $arcsam "User in the ARchive OU but not in Archive DB, move to repsrvarc01"
        Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message $arcsam" User in the ARchive OU but not in Archive DB, move to repsrvarc01"
        $notarcDB | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\Engineering\PS\data\Changes\$today"_MovedToArcDB.csv"
        $NotarcDBcount++
        Set-Mailbox $arcsam  -MaxReceiveSize 1KB
        Set-Mailbox $deadsam -CustomAttribute3 $null #Area
        Set-Mailbox $deadsam -CustomAttribute12 $null #Region
        Set-Mailbox $deadsam -CustomAttribute13 $null #Job Code
        Set-Mailbox $deadsam -CustomAttribute14 $null #BU
        Set-Mailbox $deadsam -CustomAttribute11 $null #Division
        Set-User $deadsam -Fax $null
        Set-User $deadsam -Title $null
        New-MoveRequest -Identity $arcsam -TargetDatabase repsrvarc02
        }}
#region Repsrvarc02 accounts are truly Disabled
<#
Query if they are not Disabled in Repsrvarc02, if not disabled then query AD to confirm if Enable or Disabled
If Enabled move out of the ARC DB and move ro REPSRVDAG01DB1
If Disabled then reset/refresh the account so it changes it's status to Disabled
#>
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in archive db to see if enabled or disabled, if Enabled move out of archvie DB"
$NotDis =Get-Mailbox -Database repsrvarc02 -resultsize unlimited | where{$_.ExchangeUserAccountControl -notmatch 'AccountDisabled'} #| ft name, alias, data*, exchangeuseraccountcontrol
if($NotDis -ne $null) {$notdis |Export-Csv \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today"_ArcUsers_notDis.csv"
foreach ($not in $NotDis){
$account=$not.alias
$ADuser=Get-ADUser $account -properties name, samaccountname, enabled| select name, samaccountname, enabled
$adsam=$aduser.samaccountname
if ($aduser.enabled -eq $True){#Write-Host $adsam "not disabled in ARCDB but is Enabled in AD"
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "users were not disabled, moving users to repsrvdag01db24"
$aduser | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today+"_InArcDB_ButEnabledAD.csv"
New-MoveRequest -Identity $adsam -BatchName $adsam -TargetDatabase repsrvdag01db1
$MoveEnabledCount++
}
if ($aduser.enabled -eq $False){#Write-Host $adsam "not disabled in ARC_DB but is Disabled in AD" #If this is true then set-mail to hidden
Set-mailbox $adsam
$UpdateDisCount++
}}}
#endregion

#region Repsrvarc06 accounts are truly Disabled
<#
Query if they are not Disabled in Repsrvarc02, if not disabled then query AD to confirm if Enable or Disabled
If Enabled move out of the ARC DB and move ro REPSRVDAG01DB1
If Disabled then reset/refresh the account so it changes it's status to Disabled
#>
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in archive db to see if enabled or disabled, if Enabled move out of archvie DB"
$NotDis =Get-Mailbox -Database repsrvarc06 -resultsize unlimited | where{$_.ExchangeUserAccountControl -notmatch 'AccountDisabled'} #| ft name, alias, data*, exchangeuseraccountcontrol
if($NotDis -ne $null) {$notdis |Export-Csv \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today"_ArcUsers_notDis.csv"
foreach ($not in $NotDis){
    $account=$not.alias
    $ADuser=Get-ADUser $account -properties name, samaccountname, enabled| select name, samaccountname, enabled
    $adsam=$aduser.samaccountname
    if ($aduser.enabled -eq $True){Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "users were not disabled, moving users to repsrvdag01db24"
        $aduser | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today+"_InArcDB_ButEnabledAD.csv"
        New-MoveRequest -Identity $adsam -BatchName $adsam -TargetDatabase repsrvdag01db1
        $MoveEnabledCount++
    }
    if ($aduser.enabled -eq $False){Set-mailbox $adsam
        $UpdateDisCount++
}}}
#endregion

#region Repsrvarc07 accounts are truly Disabled
<#
Query if they are not Disabled in Repsrvarc02, if not disabled then query AD to confirm if Enable or Disabled
If Enabled move out of the ARC DB and move ro REPSRVDAG01DB1
If Disabled then reset/refresh the account so it changes it's status to Disabled
#
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in archive db to see if enabled or disabled, if Enabled move out of archvie DB"
$NotDis =Get-Mailbox -Database repsrvarc07 -resultsize unlimited | where{$_.ExchangeUserAccountControl -notmatch 'AccountDisabled'} #| ft name, alias, data*, exchangeuseraccountcontrol
if($NotDis -ne $null) {$notdis |Export-Csv \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today"_ArcUsers_notDis.csv"
foreach ($not in $NotDis){
    $account=$not.alias
    $ADuser=Get-ADUser $account -properties name, samaccountname, enabled| select name, samaccountname, enabled
    $adsam=$aduser.samaccountname
    if ($aduser.enabled -eq $True){Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "users were not disabled, moving users to repsrvdag01db24"
        $aduser | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\Engineering\PS\Data\Changes\$today+"_InArcDB_ButEnabledAD.csv"
        New-MoveRequest -Identity $adsam -BatchName $adsam -TargetDatabase repsrvdag01db1
        $MoveEnabledCount++
    }
    if ($aduser.enabled -eq $False){Set-mailbox $adsam
        $UpdateDisCount++
}}}
#endregion
#>

<#
Users where OU=Dead Accounts
Query these users to see if they are marked to be hidden from address book and if oof is enalbed
if not hide from address book, enable oof office with a standard message and remove all job codes
The out of office spreadsheet will get smaller as oof is enabled for all accounts
#
$DeadAccounts=$adusers | where {$_.distinguishedname -like "*dead accounts*"}
Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Query users in the Dead Accounts OU"
Write-Host "Query users in the Dead Accounts OU"
foreach($dead in $DeadAccounts){
    Start-Sleep 4
    $deadsam=$dead.samaccountname
    $deadname=$dead.displayname
    #$deadsam
    $MB=Get-MailboxAutoReplyConfiguration $deadsam
    $OOF=$MB.AutoReplyState
    $MB | fl identity, autoreplystate | Out-File -Append c:\temp\oof.txt
    if (($notHidden=Get-Mailbox $deadsam | select name, alias, HiddenFromAddressListsEnabled | where{$_.HiddenFromAddressListsEnabled -eq $False}) -ne $null) {
        Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message $deadsam" User not set to Hide from Address book, set to hdie from address book"
        Write-Host "Users not set to Hide from Address book, set to hdie from address book"
        $notHidden | Export-Csv -Append -NoTypeInformation \\repsharedfs\techsupport\ADreports\Changes\$today"_DeadAccountsHidden.csv"
        $NotHiddencount++
        Set-mailbox $deadsam -HiddenFromAddressListsEnabled $true
        Set-Mailbox $deadsam -CustomAttribute3 $null #Area
        Set-Mailbox $deadsam -CustomAttribute12 $null #Region
        Set-Mailbox $deadsam -CustomAttribute13 $null #Job Code
        Set-Mailbox $deadsam -CustomAttribute14 $null #BU
        }
    if ($oof -like "*disabled*") {$mb | select Identity, autoreplystate | Export-Csv -Append \\repsharedfs\techsupport\ADreports\Changes\$today"_DeadAccount_OOF.csv"
        Write-EventLog -LogName AD_Changes -Source AD_Changes -EventId 18500 -EntryType Information -Message "Users out of office was not set, Enable out of office and configure message"
        Write-Host "Users out of office was not set, Enable out of office and configure message"
        Set-Mailbox $deadsam -CustomAttribute13 $null
        Clear-Variable managerdetails
        $oofenableCount++
        $managerDetails = Get-ADUser (Get-ADUser $deadsam -properties manager).manager
        if ($managerDetails -ne $null) {#Write-Host "---enable oof and set message"
            $managerDetails2= Get-Mailbox $managerDetails.samaccountname | select DisplayName, PrimarySmtpAddress
            $mngrname = $managerDetails2.displayname
            $mngremail = $managerDetails2.PrimarySmtpAddress.ToString()
            $OOFMessage = "This is an automated response that needs your immediate attention, '"+$deadname +"' is no longer employed by the company. If your email needs attention or review, you must send it to '"+$mngrname +"' at the following email address: " +$mngremail
            Set-MailboxAutoReplyConfiguration -Identity $deadsam -AutoReplyState Enabled -InternalMessage $oofmessage -ExternalMessage $OOFmessage
            Clear-Variable mngrname,mngremail}
        Else {#Write-Host "Has no manager"
        $OOFMessageNomgr ="This is an automated response that needs your immediate attention. '"+$deadname +"' is no longer employed by the company.  This message will not be read."
        Set-MailboxAutoReplyConfiguration -Identity $deadsam -AutoReplyState Enabled -InternalMessage $OOFMessageNomgr -ExternalMessage $OOFMessageNomgr}
}
}
#>
#endregion

$Done=Get-Date
# sending message
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = "Prodpsapp02@republicservices.com"
    $message.to.Add("vsanchez2@republicservices.com")
    #, ETuchband@republicservices.com, LGonzalez@republicservices.com, RGlonek@republicservices.com")
    $message.IsBodyHTML = $true
    $smtp = new-object Net.Mail.SmtpClient
    $smtp.Host = "relay.repsrv.com"
    $smtp.UseDefaultCredentials = $true
    $Start=Get-Date
    $message.Subject = ("Job Completed.. Modified accounts when Disabled/Enabled!!!  "+$today)
        $message.Body = @"
        <html>
        <body>

        <p>Jobs Started at: $Start</p> 
        <p>Number of Disabled users moved to ARCSRVDB: $NotarcDBcount</p>
        <p>Number of Users modified to Hide from Address book: $NotHiddencount</p>
        <p>Number of Users out of office set and enabled: $oofenabledCount</p>
        <p>Number of Enabled Users moved from ARCSRVDB to Repsrvdag01db24: $MoveEnabledCount</p>
        <p>Number of Disabled Users not showing Disabled, force to update: $UpdateDisCount</p> 
        <p>Jobs completed at: $Done</p>
               
        <b>Please do not reply to this message. This message was system-generated and e-mails sent by reply are not monitored.</b>
        </body>
"@
$smtp.Send($message)

#Clear-Variable UpdateDisCount, NotHiddencount, NotarcDBcount, MoveEnabledCount
Clear-Variable ADusers, DeadAccounts, Archived

$Day = Get-Date (Get-Date).AddDays(2) -format yyyyMMdd
"Modified_Users_Archived" | Out-File \\repsharedfs\techsupport\Engineering\PS\Data\Tasks\TasksCompletedWeekend_$Day.txt -Append