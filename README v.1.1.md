Description of Changes:
=======================

Email Conflict Check:
---------------------
Implemented a function (Check-EmailExists) to check if an email address already exists in the entire domain (not just in the specified OU).
When a conflict is detected, the script skips adding the conflicting email and logs the details.

User Modification Tracking:
---------------------------
Introduced a flag ($UserModified) to track whether any modifications were made to a user’s email or proxy addresses.
Only users who had their email or proxy addresses modified are counted in the final summary.

Email Conflict Logging:
-----------------------
Created an array ($EmailConflicts) to store details of users who encountered email conflicts, including their names, sAMAccountNames, and conflicting email addresses.

Detailed Output for Conflicts:
------------------------------
Enhanced the output to display all users with email conflicts, formatted with user names and their conflicting email addresses.
The count of skipped users due to email conflicts is included in the output.

General Email Field Check:
--------------------------
Ensured that if the "General - E-mail" email field is filled for any user, it is checked against existing emails in the domain before allowing updates.

Formatting Improvements:
------------------------
Added colored output for user messages to improve readability.
Included a blank line before the final summary message to visually separate it from the previous output.


Popis změn:
===========

Kontrola konfliktů e-mailů:
---------------------------
Implementována funkce (Check-EmailExists), která kontroluje, zda e-mailová adresa již existuje v celé doméně (nejen v určeném OU).
Když dojde k detekci konfliktu, skript přeskočí přidání konfliktujícího e-mailu a zaznamená detaily.

Sledování změn uživatelů:
-------------------------
Zaveden příznak ($UserModified), který sleduje, zda došlo k jakýmkoli změnám v e-mailu nebo proxy adresách uživatele.
Počítají se pouze uživatelé, kteří měli upravený e-mail nebo proxy adresy, do konečného shrnutí.

Záznam konfliktů e-mailů:
-------------------------
Vytvořeno pole ($EmailConflicts), které ukládá detaily uživatelů, kteří se setkali s konflikty e-mailů, včetně jejich jmen, sAMAccountNames a konfliktujících e-mailových adres.

Podrobný výstup pro konflikty:
------------------------------
Vylepšený výstup, který zobrazuje všechny uživatele s konflikty e-mailů, formátovaný se jmény uživatelů a jejich konfliktujícími e-mailovými adresami.
Zahrnuto je také počítání přeskočených uživatelů kvůli konfliktům e-mailů.

Kontrola e-mailového pole General - E-mail:
-------------------------------------------
Zajištěno, že pokud je vyplněno pole "General - E-mail" pro jakéhokoli uživatele, je zkontrolováno proti existujícím e-mailům v doméně před povolením aktualizací.

Vylepšení formátování:
----------------------
Přidán barevný výstup pro zprávy uživatelů pro zlepšení čitelnosti.
Zahrnut prázdný řádek před závěrečnou zprávou pro vizuální oddělení od předchozího výstupu.

