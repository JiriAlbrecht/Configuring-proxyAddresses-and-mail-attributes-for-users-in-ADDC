# Configuring proxyAddresses and mail attributes for users in ADDC

## Configuring proxyAddresses and mail attributes for users in ADDC

#### Run PowerShell as administrator!
The script will perform bulk setting of the proxyAddresses attribute and the mail attribute for all Users in the selected OU.

In Exchange Online, there are no default policies for setting a unified email address.

In the selected OU, the script will delete the proxyAddresses and mail settings for all users. Then, it will set the proxyAddresses in the following format:
* Primary email: Givenname.Surname@Your domain
* Secondary email: Surname.Givenname@Your domain, sAMAccountName@Your domain
* Mail: Givenname.Surname@Your domain
* "Your domain" is derived from the domain. 

Set the OU where changes to proxyAddresses will be made! Changes apply to all Users in the OU.




## List proxyAddresses and email

The script creates a list of users in the specified OU and prints their proxyAddresses and mail attributes. Everything is saved into a file.

Set the OU where changes to proxyAddresses will be made! Changes apply to all Users in the OU.
