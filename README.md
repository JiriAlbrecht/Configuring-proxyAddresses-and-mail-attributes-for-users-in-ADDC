# Configuring proxyAddresses and mail attributes for users in ADDC

## Configuring proxyAddresses and mail attributes for users in ADDC

#### Run PowerShell as administrator!

Set the OU where the proxyAddresses changes will be applied! The changes affect all users in the OU. Configure the email address format in the script.
Nastavte OU, kde se budou provadet zmeny proxyAddresses! Zmeny se tykaji vsech uzivatelu v OU. Ve skriptu nastavte format emailovych adres.

The script performs a bulk configuration of the proxyAddresses and mail attributes for all users in the selected OU. 
There are no default policies in Exchange Online for setting a unified email address.

In the selected OU, the script will delete the proxyAddresses and mail settings for all Users, unless it is chosen to add only a secondary domain.
Then, it sets the proxyAddresses and mail in the chosen format.

The script allows selecting the primary address and individual email aliases:
   * FirstName.LastName@YourDomain
   * LastName.FirstName@YourDomain
   * LastName@YourDomain
   * sAMAccountName@YourDomain

"YourDomain" is retrieved from the domain.

The script also allows selecting the secondary domain address and individual email aliases: 
   * FirstName.LastName@SecondaryDomain
   * LastName.FirstName@SecondaryDomain
   * LastName@SecondaryDomain

Add only email for secondary domain
Keeps email settings and adds a secondary email address
   * FirstName.LastName@SecondaryDomain
   * LastName.FirstName@SecondaryDomain
   * LastName@SecondaryDomain


Skript udela hromadne nastaveni atributu proxyAddresses a atributu mail u vsech Users ve vybrane OU.
V Exchange Online nejsou vychozi politiky pro nastaveni jednotne emailove adresy.
    
Ve zvolenem OU skript vymaze nastaveni proxyAddresses a mail u vsech Users, pokud neni zvoleno pridat pouze sekundarni domenu.
Potom nastavi proxiAddresses a mail ve zvolenem formatu.
                         
Skrypt umoznuje zvolit primarni adresu a jednotlive emailove aliasy.
   * FirstName.LastName@Your domain
   * LastName.FirstName@Your domain
   * Lastname@Your domain
   * sAMAccountName@Your domain
    
"Your domain" vypise z domeny.

Skrypt umoznuje zvolit sekundarni domenovou adresu a jednotlive emailove aliasy.
   * FirstName.LastName@Secondary Domain
   * LastName.FirstName@Secondary Domain
   * Lastname@Secondary Domain 

Pridat pouze sekundarni domenu.
Zachova nastaveni emailu a prida sekundarni emailovou adresu
   * FirstName.LastName@Secondary Domain
   * LastName.FirstName@Secondary Domain
   * Lastname@Secondary Domain 
    
## List proxyAddresses and email
The script creates a list of users in the specified OU and prints their proxyAddresses and mail attributes. Everything is saved into a file.

Skript vytvoří seznam uživatelů v zadané organizační jednotce a vytiskne jejich proxyAddresses a atribut mail. Vše se uloží do souboru.

