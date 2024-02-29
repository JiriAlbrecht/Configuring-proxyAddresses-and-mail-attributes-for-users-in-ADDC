#################################
# Add proxyAddresses and email  #
#################################
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

# Zvýraznění dotazu na začátku
Write-Host "`nAre you sure you want to make changes in the OU '$OU'? (Y/N)" -ForegroundColor Cyan

$ConfirmOU = Read-Host
if ($ConfirmOU -ne "Y") {
    Write-Host "Operation canceled. Exiting script."
    Exit
}

$ErrorOccurred = $false
$UsersChanged = 0

# List Users from the specified OU
$Users = Get-ADUser -Filter * -SearchBase $OU 

# Loop through each user in the OU
Foreach ($User in $Users) {
    Write-Host "Editing user: $($User.Name)"
    
    # Delete proxyAddresses
    $User.proxyAddresses.Clear()
    Set-ADUser -Instance $User -ErrorAction SilentlyContinue
    if ($?) {
        # Delete mail
        $User.mail = $null
        Set-ADUser -Instance $User -ErrorAction SilentlyContinue

        # Add proxyAddresses
        $GivenName = $User.GivenName
        $Surname = $User.Surname

        # Remove diacritics (accents) from names
        $GivenNameClean = Remove-Diacritics -inputString $GivenName
        $SurnameClean = Remove-Diacritics -inputString $Surname
      
        $Domain = $User.UserPrincipalName.Split('@')[1] # Extract domain from User Logon Name
        
        # Construct email addresses without diacritics
        $EmailAddress4 = $GivenNameClean + "." + $SurnameClean + "@" + $Domain
        $EmailAddress1 = "SMTP:" + $GivenNameClean + "." + $SurnameClean + "@" + $Domain
        $EmailAddress2 = "smtp:" + $SurnameClean + "." + $GivenNameClean + "@" + $Domain
        $EmailAddress3 = "smtp:" + $User.sAMAccountName + "@" + $Domain

        # Add e-mail to mail property
        $User.mail = $EmailAddress4
        
        # Add proxyAddresses
        try {
            $User.proxyAddresses.Add($EmailAddress1)
            Set-ADUser -Instance $User -ErrorAction Stop
            $UsersChanged++
        } catch {
            Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress1"
            $ErrorOccurred = $true
        }
        
        try {
            $User.proxyAddresses.Add($EmailAddress2)
            Set-ADUser -Instance $User -ErrorAction Stop
            $UsersChanged++
        } catch {
            Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress2"
            $ErrorOccurred = $true
        }
        
        try {
            $User.proxyAddresses.Add($EmailAddress3)
            Set-ADUser -Instance $User -ErrorAction Stop
            $UsersChanged++
        } catch {
            Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Error occurred while adding address: $EmailAddress3"
            $ErrorOccurred = $true
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

if ($ErrorOccurred) {
    Write-Host -ForegroundColor Cyan -BackgroundColor DarkRed "Script finished with an error."
} else {
    Write-Host -ForegroundColor Cyan "Script finished successfully. Number of users changed: $UsersChanged"
}

