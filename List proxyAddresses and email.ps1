#################################
# List proxyAddresses and email #
#################################
###
###
###################################
# Author: Jiri Albrecht           #
# https://github.com/JiriAlbrecht #
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
################################################################################################################################################
# The script creates a list of users in the specified OU and prints their proxyAddresses and mail attributes. Everything is saved into a file. #
################################################################################################################################################
###
###########################################################################################################################################
# Skript vytvoří seznam uživatelů v zadané organizační jednotce a vytiskne jejich proxyAddresses a atribut mail. Vše se uloží do souboru. #
###########################################################################################################################################
###
###
###
# Import-Module ActiveDirectory # (if you’re not running it on a DC, make sure you have installed the Active Directory module for PowerShell via RSAT)

# OU in which proxyAddresses are changed
$OU = "OU=Test,DC=contoso,DC=com"

# Create or clear the content of the backup file
"DN;smtpAddress;mail" | Out-File ".\List proxyAddresses.txt" -Encoding utf8

# Get objects with the proxyAddresses and mail properties in the specified OU
Get-ADObject -SearchBase $OU -LDAPFilter "(&(proxyAddresses=*)(mail=*))" -Properties proxyAddresses, mail | ForEach-Object {
    
    # Process each object
    $proxyAddresses = $_.proxyAddresses -join ';'
    
    # Create the output string
    $Output = $_.distinguishedName + ";" + $proxyAddresses + ";" + $_.mail
    Write-Host $Output
        
    # Directly write to the file with closing the file after writing
    $Output | Out-File -FilePath ".\List proxyAddresses.txt" -Append -Encoding utf8
}