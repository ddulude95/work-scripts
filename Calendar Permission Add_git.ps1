#
# Initial commands procured by James Sargent
# Modified into a script by Devin Dulude 10/19/2020
#
#


# ============================== STEP 1 READ THIS!!!! ==========================
# 1. replace the users in the $providers and $psrs arrays to your needed users. use their full email addresses.
# 2. also modify the type of access you want to give on lines 19 and 20

# array of users whose calendar needs modifying (typically providers).
$providers = @("email@domain.com") 

# array of users who are going to gain access to those calendars (typically PSRs)
$psrs = @("email@domain.com")

#type of access you want to give
$accessType = "Editor"
$mailboxFolder = ":\Calendar"

# prompts user (you) for credentials. Enter your email
$Credential = Get-Credential

# creating a connection with O365 exchange and giving it our credentials
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication "Basic" -AllowRedirection

# importing the session so all further commands will be executed in the exchange session
Import-PSSession $ExchangeSession

# ============================== STEP 2 READ THIS!!!! ==========================
# nested loop that goes through each psr and checks to see if any permissions already exist.
# If so, overwrite with $accessType. If they're already editor, then skip them. If no permissions exist, then add them.


foreach ($psr in $psrs){
    
    foreach ($provider in $providers){
        Write-Host "`nChecking for existing permissons for" $psr "on" ($provider + $mailboxFolder)
        $result = Get-MailboxFolderPermission -Identity ($provider + $mailboxFolder) -user $psr -ea silentlycontinue | Select-Object -ExpandProperty AccessRights

        if (!$result){
            Write-Host "no existing permissions found..."
            Write-Host "`nAdding" $psr "to" ($provider + $mailboxFolder) "AS" $accessType
            Add-MailboxFolderPermission -Identity ($provider + $mailboxFolder) -user $psr -AccessRights $accessType
        }
        elseif ($result -eq "Editor") {
            Write-Host $psr "already" $accessType "for" ($provider + $mailboxFolder) ", moving on to next"
        }
        else {
            Write-Host $psr "is" $result ", over writing to" $accessType
            Set-MailboxFolderPermission -Identity ($provider + $mailboxFolder) -user $psr -AccessRights $accessType
        }          
    }

}
Write-Host "Script has ended."

# EXTRA CODE ( NOT NEEDED )
# nested loop to go through and verify access
# Not necessary to run, only used to look at what current permissions are
<#
foreach ($psr in $psrs){
    Write-Host "`nVerifying" $psr "access for the following calendars:"
    foreach ($provider in $providers){
        Get-MailboxFolderPermission -Identity ($provider + ":\calendar") -user $psr | ft Identity,User,AccessRights            
    }

}
#>