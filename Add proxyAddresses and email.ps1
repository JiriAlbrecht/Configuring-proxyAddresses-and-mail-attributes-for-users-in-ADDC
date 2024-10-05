#################################
# Add proxyAddresses and email  #
#################################
###
###
###############
# version 1.1 #
###############
###
###
###################################
# Author: Jiri Albrecht           #
# https://github.com/JiriAlbrecht #
###################################
###
###
###################################
# Run PowerShell as administrator #
###################################
###
###
##################################################################################################
# Set the OU where changes to proxyAddresses will be made! Changes apply to all Users in the OU. #
##################################################################################################
###
#################################################################################################
# Nastavte OU, kde se budou provadet zmeny proxyAddresses! Zmeny se tykaji vsech uzivatelu v OU #
#################################################################################################
###
###
#################################################################################################################################
# The script will perform bulk setting of the proxyAddresses attribute and the mail attribute for all Users in the selected OU. #
# In Exchange Online, there are no default policies for setting a unified email address.                                        #
#################################################################################################################################
###
########################################################################################################
# Skript udela hromadné nastavení atributu proxyAddresses a atributu mail u vsech Users ve vybrané OU. #
# V Exchange Online nejsou vychozi politiky pro nastaveni jednotne emailove adresy.                    #
########################################################################################################
###
###
################################################################################################################################################################
# In the selected OU, the script will delete the proxyAddresses and mail settings for all users. Then, it will set the proxyAddresses in the following format: #
# Primary email: Surname.Givenname@Your domain                                                                                                                 #
# Secondary email: Givenname.Surname@Your domain, sAMAccountName@Your domain                                                                                   #
# Mail: Surname.Givenname@Your domain                                                                                                                          #
# "Your domain" is derived from the domain.                                                                                                                    #
################################################################################################################################################################
###
#######################################################################################################################
# Ve zvolene OU skrypt vymaze nastaveni proxyAddresses a mail u všech Users. Potom nastavi proxiAddresses ve formatu: #
# primarní email: Surname.Givenname@Your domain                                                                       #
# sekundarní email: Givenname.Surname@Your domain a sAMAccountName@Your domain                                        #
# mail: Surname.Givenname@Your domain                                                                                 #
# "Your domain" vypise z domeny.                                                                                      #
#######################################################################################################################
###
###
###################################################################################################################################################
# The script records how many users were changed and outputs this information upon completion, along with whether the script ended with an error. #
###################################################################################################################################################
###
##################################################################################################################################################
# Skript zaznamenava, kolik uzivatelu bylo zmeneno. Vypisuje tuto informaci po dokoncení skriptu spolu s informaci, zda skript skoncil s chybou. # 
##################################################################################################################################################
###
###
###
#########################################
# Description of changes in version 1.1 #
# - Email Conflict Check                #
# - Formatting Improvements             #
#########################################
###
################################
# Popis zmen verze 1.1         #
# - Kontrola konfliktu e-mailu #
# - Vylepseni formatovani      #
################################
###
###
###
# Import-Module ActiveDirectory # (if you’re not running it on a DC, make sure you have installed the Active Directory module for PowerShell via RSAT)

# OU in which proxyAddresses are changed
$OU = "OU=Test,DC=contoso,DC=com"

# Function to remove diacritics (accents) from characters
function Remove-Diacritics {
    param(
        [string]$inputString
    )

    $normalizedString = $inputString.Normalize([Text.NormalizationForm]::FormD)
    $diacriticFreeString = $normalizedString -creplace '\p{M}', ''

    return $diacriticFreeString
}

# Function to check if email exists in the entire domain (proxyAddresses or mail attribute)
function Check-EmailExists {
    param(
        [string]$email
    )

    # Search in the whole domain for both proxyAddresses and mail attribute
    $existingUser = Get-ADUser -Filter { proxyAddresses -eq $email -or mail -eq $email } -Properties proxyAddresses, mail
    return $existingUser -ne $null
}

# Highlight confirmation message at the beginning
Write-Host "`nAre you sure you want to make changes in the OU '$OU'? (Y/N)" -ForegroundColor Cyan

$ConfirmOU = Read-Host
if ($ConfirmOU -ne "Y") {
    Write-Host "Operation canceled. Exiting script."
    Exit
}

$ErrorOccurred = $false
$UsersChanged = 0
$EmailConflicts = @()

# List users from the specified OU
$Users = Get-ADUser -Filter * -SearchBase $OU 

# Loop through each user in the OU
Foreach ($User in $Users) {
    Write-Host "Editing user: $($User.Name)"
    
    $UserModified = $false  # Flag to track if the user has been modified
    $ConflictingEmails = @()  # Array to store conflicting emails for the user
    
    # Clear proxyAddresses
    $User.proxyAddresses.Clear()
    Set-ADUser -Instance $User -ErrorAction SilentlyContinue
    if ($?) {
        # Clear mail
        $User.mail = $null
        Set-ADUser -Instance $User -ErrorAction SilentlyContinue

        # Prepare proxyAddresses
        $GivenName = $User.GivenName
        $Surname = $User.Surname

        # Remove diacritics from names
        $GivenNameClean = Remove-Diacritics -inputString $GivenName
        $SurnameClean = Remove-Diacritics -inputString $Surname
      
        $Domain = $User.UserPrincipalName.Split('@')[1] # Extract domain from User Logon Name
        
        # Construct email addresses without diacritics
        $EmailAddress4 = $GivenNameClean + "." + $SurnameClean + "@" + $Domain
        $EmailAddress1 = "SMTP:" + $GivenNameClean + "." + $SurnameClean + "@" + $Domain
        $EmailAddress2 = "smtp:" + $SurnameClean + "." + $GivenNameClean + "@" + $Domain
        $EmailAddress3 = "smtp:" + $User.sAMAccountName + "@" + $Domain

        # Add email to mail property if it does not already exist in the domain
        if (-not (Check-EmailExists $EmailAddress4)) {
            $User.mail = $EmailAddress4
            $UserModified = $true
        } else {
            $ConflictingEmails += $EmailAddress4
            Write-Host "Skipping e-mail $EmailAddress4 for user $($User.Name) due to conflict."
        }

        # Try to add proxyAddresses if they do not already exist in the domain
        if (-not (Check-EmailExists $EmailAddress1)) {
            try {
                $User.proxyAddresses.Add($EmailAddress1) | Out-Null
                Set-ADUser -Instance $User -ErrorAction Stop
                $UserModified = $true
            } catch {
                Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress1"
                $ErrorOccurred = $true
            }
        } else {
            $ConflictingEmails += $EmailAddress1
            Write-Host "Skipping e-mail $EmailAddress1 for user $($User.Name) due to conflict."
        }

        if (-not (Check-EmailExists $EmailAddress2)) {
            try {
                $User.proxyAddresses.Add($EmailAddress2) | Out-Null
                Set-ADUser -Instance $User -ErrorAction Stop
                $UserModified = $true
            } catch {
                Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress2"
                $ErrorOccurred = $true
            }
        } else {
            $ConflictingEmails += $EmailAddress2
            Write-Host "Skipping e-mail $EmailAddress2 for user $($User.Name) due to conflict."
        }

        if (-not (Check-EmailExists $EmailAddress3)) {
            try {
                $User.proxyAddresses.Add($EmailAddress3) | Out-Null
                Set-ADUser -Instance $User -ErrorAction Stop
                $UserModified = $true
            } catch {
                Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress3"
                $ErrorOccurred = $true
            }
        } else {
            $ConflictingEmails += $EmailAddress3
            Write-Host "Skipping e-mail $EmailAddress3 for user $($User.Name) due to conflict."
        }

        # Only count the user if they have been modified (email added or changed)
        if ($UserModified) {
            $UsersChanged++
        }

        # If there are any conflicting emails, store the user and their conflicting emails
        if ($ConflictingEmails.Count -gt 0) {
            $EmailConflicts += [pscustomobject]@{
                Name = $User.Name
                sAMAccountName = $User.sAMAccountName
                Emails = $ConflictingEmails -join ", "
            }
        }

        if (-not $ErrorOccurred) {
            # Display the edited user and their proxyAddresses and mail settings
            Write-Host "User $($User.Name) settings:"
            Write-Host "ProxyAddresses: $($User.proxyAddresses -join ", ")"
            Write-Host "Mail: $($User.mail)"
            Write-Host ""
        }
    } else {
        $ErrorOccurred = $true
    }
}

# Display email conflicts if any
if ($EmailConflicts.Count -gt 0) {
    Write-Host -ForegroundColor Yellow "The following users had email conflicts ($($EmailConflicts.Count) users skipped):"
    $EmailConflicts | ForEach-Object {
        Write-Host -ForegroundColor Yellow "User: $($_.Name) (sAMAccountName: $($_.sAMAccountName)), Conflicting Emails: $($_.Emails)"
    }
}

# Add an empty line before the final message
Write-Host ""

if ($ErrorOccurred) {
    Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Script finished with an error."
} else {
    Write-Host -ForegroundColor Cyan "Script finished successfully. Number of users changed: $UsersChanged"
}

