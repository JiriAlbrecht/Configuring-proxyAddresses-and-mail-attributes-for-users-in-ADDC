#Requires -RunAsAdministrator

<#
.SYNOPSIS
Add proxyAddresses and email

.VERSION
2.1

.AUTHOR
Jiri Albrecht

.LINK
https://github.com/JiriAlbrecht

.Requires
Set the OU where the proxyAddresses changes will be applied! The changes affect all users in the OU. Configure the email address format in the script.
Nastavte OU, kde se budou provadet zmeny proxyAddresses! Zmeny se tykaji vsech uzivatelu v OU. Ve skriptu nastavte format emailovych adres.

.DESCRIPTION
The script performs a bulk configuration of the proxyAddresses and mail attributes for all users in the selected OU. 
There are no default policies in Exchange Online for setting a unified email address.

In the selected OU, the script will delete the proxyAddresses and mail settings for all Users, unless it is chosen to add only a secondary domain.
Then, it sets the proxyAddresses and mail in the chosen format.

The script allows selecting the primary address and individual email aliases:
    FirstName.LastName@YourDomain
    LastName.FirstName@YourDomain
    LastName@YourDomain
    sAMAccountName@YourDomain

"YourDomain" is retrieved from the domain.

The script also allows selecting the secondary domain address and individual email aliases: 
    FirstName.LastName@SecondaryDomain
    LastName.FirstName@SecondaryDomain
    LastName@SecondaryDomain

Add only email for secondary domain
Keeps email settings and adds a secondary email address
    FirstName.LastName@SecondaryDomain
    LastName.FirstName@SecondaryDomain
    LastName@SecondaryDomain


Skript udela hromadne nastaveni atributu proxyAddresses a atributu mail u vsech Users ve vybrane OU.
V Exchange Online nejsou vychozi politiky pro nastaveni jednotne emailove adresy.
    
Ve zvolenem OU skript vymaze nastaveni proxyAddresses a mail u vsech Users, pokud neni zvoleno pridat pouze sekundarni domenu.
Potom nastavi proxiAddresses a mail ve zvolenem formatu.
                         
Skrypt umoznuje zvolit primarni adresu a jednotlive emailove aliasy.
    FirstName.LastName@Your domain
    LastName.FirstName@Your domain
    Lastname@Your domain
    sAMAccountName@Your domain
    
"Your domain" vypise z domeny.

Skrypt umoznuje zvolit sekundarni domenovou adresu a jednotlive emailove aliasy.
    FirstName.LastName@Secondary Domain
    LastName.FirstName@Secondary Domain
    Lastname@Secondary Domain 

Pridat pouze sekundarni domenu.
Zachova nastaveni emailu a prida sekundarni emailovou adresu
    FirstName.LastName@Secondary Domain
    LastName.FirstName@Secondary Domain
    Lastname@Secondary Domain 
#>


# --- CONFIGURATION ---

# OU in which proxyAddresses are changed
$OU = "OU=Test,DC=contoso,DC=com"

# Specify email format and primary email
$EnableEmailOption1 = 1 # FirstName.LastName@contoso.com (Enable email = 1; Disable email = 0; Recommended - 1)
$EnableEmailOption2 = 1 # LastName.FirstName@contoso.com (Enable email = 1; Disable email = 0; Recommended - 1)
$EnableEmailOption3 = 0 # Lastname@contoso.com (Enable email = 1; Disable email = 0; Recommended - 0)
$EnableEmailOption4 = 0 # sAMAccountName@contoso.com (Enable email = 1; Disable email = 0; Recommended - 0)

# Primary email - 1 - 4 (Recommended - 1) - Number corresponds to previous numbers
$EnablePrimaryEmail = 1 

# End Specify email format and primary email


# Specify secondary domain in email addresses

# Add only email for secondary domain (Add Only Secondary Domain = 1; Disable Add Only Secondary Domain = 0). Sets only the secondary email.
$AddOnlySecondaryDomain = 0

# Add email for secondary domain (Enable Secondary Domain = 1; Disable Secondary Email = 0). It will be applied after setting up the primary email.
$EnableSecondaryDomain = 0

# Specify Secondary Domain (Example: adatum.com)
$SecondaryDomain = "adata.com"

# End Specify secondary domain in email addresses


# Options for secondary domain
$EnableEmailOption5 = 1 # FirstName.LastName@adatum.com (Enable email = 1; Disable email = 0; Recommended - 1)
$EnableEmailOption6 = 1 # LastName.FirstName@adatum.com (Enable email = 1; Disable email = 0; Recommended - 1)
$EnableEmailOption7 = 0 # Lastname@adatum.com (Enable email = 1; Disable email = 0; Recommended - 0)

### End Specify secondary domain in email addresses ###

# --- END CONFIGURATION ---


# --- FUNCTIONS ---

# Function to remove diacritics (accents) from characters
function Remove-Diacritics {    
    param(
        [string]$inputString
    )
    $normalizedString = $inputString.Normalize([Text.NormalizationForm]::FormD)
    $diacriticFreeString = $normalizedString -creplace '\p{M}', ''
 #   if (-not $inputString) { return "Unknown" } # Value filling test
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

# --- END FUNCTIONS ---


# --- VALIDATION & CONFIRMATION ---

# Validation of the path to the OU
try {
    Get-ADOrganizationalUnit -Identity $OU -ErrorAction Stop
    Write-Host -ForegroundColor Cyan "`n[OK] OU path verified: "$OU
    Write-Host "`n"
} catch {
    Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "`n[ERROR] OU not found: "$OU
    Write-Host "`n"
    return
}

# Highlight confirmation message at the beginning
Write-Host "`nAre you sure you want to make changes in the OU '$OU'? (Y/N)" -ForegroundColor Yellow
Write-Host ""

$ConfirmOU = Read-Host
if ($ConfirmOU -ne "Y") {
    Write-Host "`nOperation canceled. Exiting script."
    Exit
}
Write-Host "`n"

# --- END VALIDATION & CONFIRMATION ---


# --- MAIN PROCESSING ---

$ErrorOccurred = $false
$UsersChanged = 0
$SkippingNumber = 0
$EmailConflicts = @()

# List users from the specified OU
$Users = Get-ADUser -Filter * -SearchBase $OU -Properties proxyAddresses, mail, GivenName, Surname, sAMAccountName, UserPrincipalName

# Check if primary email option is disabled
if ($AddOnlySecondaryDomain -eq 0) {
    if (($EnablePrimaryEmail -eq 1 -and $EnableEmailOption1 -eq 0) -or
        ($EnablePrimaryEmail -eq 2 -and $EnableEmailOption2 -eq 0) -or
        ($EnablePrimaryEmail -eq 3 -and $EnableEmailOption3 -eq 0) -or
        ($EnablePrimaryEmail -eq 4 -and $EnableEmailOption4 -eq 0)) {
        Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error: The selected primary email option ($EnablePrimaryEmail) is disabled. Enable it or choose a different primary email." 
        Exit
    }
}

# Loop through each user in the OU
Foreach ($User in $Users) {
    Write-Host ("-" * 50)"`n"
    Write-Host "Editing user: $($User.Name)" -ForegroundColor Cyan
    
    $UserModified = $false  # Flag to track if the user has been modified
    $ConflictingEmails = @()  # Array to store conflicting emails for the user
    $EmailAddressPrimary = $null

    # Prepare proxyAddresses
    $GivenName = $User.GivenName
    $Surname = $User.Surname

    # Value filling test in AD
    if ([string]::IsNullOrWhiteSpace($GivenName) -or [string]::IsNullOrWhiteSpace($Surname)) {
        Write-Host "[SKIP] User '$($User.Name)' has missing GivenName or Surname in AD!" -ForegroundColor Cyan -BackgroundColor DarkRed
        $ErrorOccurred = $true
        $SkippingNumber++
        continue # Skip to next user
    }

    # Remove diacritics from names
    $GivenNameClean = Remove-Diacritics -inputString $GivenName
    $SurnameClean = Remove-Diacritics -inputString $Surname
    $Domain = $User.UserPrincipalName.Split('@')[1] # Extract domain from User Logon Name

# --- END MAIN PROCESSING ---


# --- SECTION 1: PRIMARY DOMAIN LOGIC ---

    # Only runs if AddOnlySecondaryDomain is NOT 1
    if ($AddOnlySecondaryDomain -eq 0) {    
    
        $User.proxyAddresses.Clear() # Clear proxyAddresses
        $User.mail = $null # Clear mail

        try {
            Set-ADUser -Instance $User -ErrorAction Stop
        } catch {
            Write-Host "Failed to update primary addresses for $($User.Name)" -ForegroundColor Red
            $ErrorOccurred = $true
            continue # Jump to next user on fatal error
        }
   
        # Construct email addresses without diacritics
        $EmailAddress1 = "$GivenNameClean.$SurnameClean@$Domain"
        $EmailAddress2 = "$SurnameClean.$GivenNameClean@$Domain"
        $EmailAddress3 = "$SurnameClean@$Domain" # Alias for last name only
        $EmailAddress4 = "$($User.sAMAccountName)@$Domain"
       
        # Try to add proxyAddresses if they do not already exist in the domain - Primary Domain        
        for ($i = 1; $i -le 4; $i++) {
            $currentEmail = Get-Variable -Name ("EmailAddress" + $i) -ValueOnly
            $enableOption = Get-Variable -Name ("EnableEmailOption" + $i) -ValueOnly
   
            if ($EnablePrimaryEmail -eq $i) { # Add primary email name
                $EmailAddressPrimary = $currentEmail   
                $currentEmailWithPrefix = "SMTP:" + $currentEmail
            } else {
                $currentEmailWithPrefix = "smtp:" + $currentEmail
            }

            if ($enableOption -eq 1) {
                if (-not (Check-EmailExists $currentEmailWithPrefix)) {
                    try {
                        $User.proxyAddresses.Add($currentEmailWithPrefix) | Out-Null
                        Set-ADUser -Instance $User -ErrorAction Stop
                        $UserModified = $true
                    } catch {
                        Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $currentEmailWithPrefix"
                        $ErrorOccurred = $true
                        $SkippingNumber++
                    }
                } else {
                    $ConflictingEmails += $currentEmailWithPrefix
                    Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Skipping email from proxyAddresses for user '$($User.Name)' due to conflict: $currentEmailWithPrefix"
                    $SkippingNumber++
                }
            }
        }
    } 

# --- END SECTION 1: PRIMARY DOMAIN LOGIC ---


# --- SECTION 2: SECONDARY DOMAIN LOGIC ---

    # Construct email addresses without diacritics in secondary domain
    if ($EnableSecondaryDomain -eq 1 -or $AddOnlySecondaryDomain -eq 1) {
        $EmailAddress5 = "$GivenNameClean.$SurnameClean@$SecondaryDomain"
        $EmailAddress6 = "$SurnameClean.$GivenNameClean@$SecondaryDomain"
        $EmailAddress7 = "$SurnameClean@$SecondaryDomain"

        # Add Secondary Domain in proxyAddresses 
        for ($i = 5; $i -le 7; $i++) {
            $currentEmail = Get-Variable -Name ("EmailAddress" + $i) -ValueOnly
            $enableOption = Get-Variable -Name ("EnableEmailOption" + $i) -ValueOnly
            $currentEmailWithPrefix = "smtp:" + $currentEmail

            if ($enableOption -eq 1) {
                if (-not (Check-EmailExists $currentEmailWithPrefix)) {
                    try {
                        $User.proxyAddresses.Add($currentEmailWithPrefix) | Out-Null
                        Set-ADUser -Instance $User -ErrorAction Stop
                        $UserModified = $true
                    } catch {
                        Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $currentEmailWithPrefix"
                        $ErrorOccurred = $true
                        $SkippingNumber++
                    }
                } else {
                    $ConflictingEmails += $currentEmailWithPrefix
                    Write-Host -ForegroundColor Yellow -BackgroundColor DarkRed "Skipping email from proxyAddresses for user '$($User.Name)' due to conflict: $currentEmailWithPrefix"
                    $SkippingNumber++
                }
            }
        }
    }

# --- END SECTION 2: SECONDARY DOMAIN LOGIC ---


# --- SECTION 3: EMAIL IN MAIL PROPERTY ---

    # Add email in mail property
    if ($AddOnlySecondaryDomain -eq 0) {

        if (-not (Check-EmailExists $EmailAddressPrimary)) {
            $User.mail = $EmailAddressPrimary  # Set email to the Mail property directly
            Set-ADUser -Instance $User -ErrorAction Stop
            $UserModified = $true
        } else {
            $ConflictingEmails += $EmailAddressPrimary
            Write-Host -ForegroundColor Yellow -BackgroundColor DarkRed "Skipping e-mail for user '$($User.Name)' due to conflict: $EmailAddressPrimary"
            $SkippingNumber++
            $ErrorOccurred = $true
        }
    }

# --- END SECTION 3: EMAIL IN MAIL PROPERTY ---


# --- REPORT AND TESTING ---

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

    # Display the edited user and their proxyAddresses and mail settings
    Write-Host "User $($User.Name) settings:"

    $CurrentMode = if ($AddOnlySecondaryDomain -eq 1) { "Add Only Secondary (Keeping existing addresses)" } else { "Full Update (Clearing existing addresses)" }
    Write-Host "Mode: $CurrentMode" -ForegroundColor Gray
    Write-Host "ProxyAddresses: $($User.proxyAddresses -join ", ")"
    Write-Host "Mail: $($User.mail)"
    Write-Host ""
}

# --- REPORT AND TESTING ---


# --- FINAL REPORT ---

# Display email conflicts if any
Write-Host "`n"("=" * 50)"`n"
if ($EmailConflicts.Count -gt 0) {
    Write-Host -ForegroundColor Yellow "The following users had email conflicts ($($EmailConflicts.Count) users skipped):"
    $EmailConflicts | ForEach-Object {
        Write-Host -ForegroundColor Yellow "User: $($_.Name) (sAMAccountName: $($_.sAMAccountName)), Conflicting Emails: $($_.Emails)"
    }
}

if ($ErrorOccurred) {
    Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Script finished with an error."
    Write-Host -ForegroundColor Cyan "Number of users changed: $UsersChanged"
    Write-Host -ForegroundColor Cyan "Number skipping email: $SkippingNumber"
} else {
    Write-Host -ForegroundColor Cyan "Script finished successfully."
    Write-Host -ForegroundColor Cyan "Number of users changed: $UsersChanged"
    Write-Host -ForegroundColor Cyan "Number skipping email: $SkippingNumber"
}

# --- END FINAL REPORT ---