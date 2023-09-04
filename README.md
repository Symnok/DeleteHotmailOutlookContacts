# DeleteHotmailOutlookContacts

Web UI of outllok.com does not allow to delete more than 50 contacts at ones (in some cases this limit is even 10 contacts)

This python3 program removes all contacts at ones.

1. Install python3
2. Install exchangelib:
pip install exchangelib

4. This program uses Basic Authentication (not OAUTH2) that still (as of Sept 1 2023) supported by Microsoft for outlook.com/hotmail.com
   The code uses MS Exchange server m.hotmail.com but you can modify the code to use any of the following URLs as of September 1 2023:

m.outlook.com

outlook.office365.com

eas.outlook.com

s.outlook.com


In the code, set your hotmail.com/outlook.com credentials in

ACCOUNT_USERNAME_AND_EMAIL

and

ACCOUNT_PASSWORD

4. 2FA supported! If you use 2FA, generate application password and set its value to ACCOUNT_PASSWORD
5. You must use modern OS that supports TLS 1.2. TLS 1.1 and below is not supported anymore by the corresponding Exchange servers
*****************************************************************************
6. Use this program on your own risk. It may contain bugs that can cause data loss, email loss, contacts loss,
   Microsoft/hotmail.com/outlook.com and any other account suspension, account deletion, 
   permanent ban by Microsoft and other unexpected consequences. 
   USE THIS PROGRAM ON YOUR OWN RISK!!!
*****************************************************************************
7. This program may not work if your account registered using non-Microsoft email address. Add Microsoft email address *@outlook.com) to your account, make it Default Alias and use this email with this program.
8. Run this program as

   python3 deleteHotmailOutlookContacts.py
9. On UI https://outlook.live.com/people/0/?actSwt=true turn ON New Contacts option. The program will work without this option too but you will not see if some contacts are stuck and not deleted
10. You may need to run this program few times before all contacts are deleted. Wait 3-5 mins before second and additional runs. You may need to remove a few remaining contacts manually.
