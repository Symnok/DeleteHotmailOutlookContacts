"""

Install exchangelib:
pip install exchangelib

You can use following servers to connect to MS Exchange Server of
hotmail.com/outlook.com/live.com as of September 1 2023:

m.outlook.com
outlook.office365.com
eas.outlook.com
s.outlook.com

Set your hotmail.com/outlook.com credentials in
ACCOUNT_USERNAME_AND_EMAIL
and
ACCOUNT_PASSWORD

If you use 2FA, generate application password and set its
value to ACCOUNT_PASSWORD
"""

from exchangelib import (
    Credentials,
    Configuration,
    OAUTH2,
    BASIC,
    Account,
    DELEGATE,
    IMPERSONATION,
    OAuth2AuthorizationCodeCredentials,
)
from exchangelib.indexed_properties import EmailAddress


SERVER_URL = "m.hotmail.com"
ACCOUNT_USERNAME_AND_EMAIL = "xxxxxxxxx@hotmail.com"
ACCOUNT_PASSWORD = "my very strong password"


#config = Configuration(server="m.hotmail.com", max_connections=10)

def connectToExchangeServer(server, USERNAME_AND_EMAIL, password):
    """
    Get Exchange account cconnection with server
    """
    creds = Credentials(username=USERNAME_AND_EMAIL, password=password)
    config = Configuration(server=SERVER_URL, credentials=creds, auth_type = BASIC)
    return Account(primary_smtp_address=USERNAME_AND_EMAIL, autodiscover=False, config=config,  access_type=DELEGATE)

def printNumberOfMessagesInAllFolders(account):
    print("*** printNumberOfMessagesInAllFolders started ***")

    all_folders = account.root.glob('**/*')
    for folder in all_folders:
        if (not str(folder).startswith("Application")) and (not str(folder).startswith("Folder")):
            print(str(folder) + " " + str(folder.all().count()))    
    print("*** printNumberOfMessagesInAllFolders ended ***\n")

def printNumberOfMessagesInSupportedFolders(account):
    print("*** printNumberOfMessagesInSupportedFolders started ***\n")

    all_folders = account.root.glob('**/*')
    for folder in all_folders:
        if (str(folder).startswith("AllContacts")) or (str(folder).startswith("RecoverableItemsDeletions") or (str(folder).startswith("Recipient Cache"))):
            print(str(folder) + " " + str(folder.all().count()))    
    print("*** printNumberOfMessagesInSupportedFolders ended ***\n")

def printNumberOfMessagesInContactsFolders(account):
    print("*** printNumberOfMessagesInContactsFolders started ***\n")

    all_folders = account.contacts.glob('**/*')
    for folder in all_folders:
        if (not str(folder).startswith("Application")):
          print(str(folder) + " " + str(folder.all().count()))    
    print("*** printNumberOfMessagesInContactsFolders ended ***\n")
    

def deleteItemsFromFolder(all_folders, visibleName, logName):
    for folder in all_folders: 
        
        if (str(folder).startswith(visibleName)):
            contacts = folder.all()
            print("contact folder: " + str(folder) + str (contacts.count()))

            print("Deletion " + logName + " started")
            count = 0
            for contact in contacts:
                print(contact)
                print ('\n')
                contact.delete()
                print (logName + ' item deleted\n')

                count = count + 1
            print("Deletion " + logName + " ended,  deleted " + str(count) + " contacts")

def deleteMainContactsFolder(account):
    contacts = account.contacts
    allContacts = contacts.all()
    print("number of contact: " + str(allContacts.count()))

    count = 0
    print("Deletion AllContacts (main folder) started")
    for contact in allContacts:

        print(contact)
        print ('\n')
        contact.delete()
        print ('AllContacts (main folder) item deleted\n')
        count = count +1
    print("Deletion AllContacts (main folder) ended, deleted " + str(count) + " contacts")

def main():           

    account = connectToExchangeServer(SERVER_URL, ACCOUNT_USERNAME_AND_EMAIL, ACCOUNT_PASSWORD)
    print (account)

    printNumberOfMessagesInAllFolders(account)

    folder = account.root / "AllContacts"
    print("number of people on account: " + str(folder.people().count()))

    print("*******************************\n")

    """
    This section purges your contact folder. You may or may not recover you address book, backup it before you run this program.
    """

    all_folders = account.root.glob('**/*')
    deleteItemsFromFolder(all_folders, "AllContacts (AllContacts)", "AllContactsAsFolder")
    print("*******************************\n")
   
    """
    # Uncomment this section if you want to delete all purged contacts completely. Remember: It may delete also your Deleted folder and you will
    # not be able to recover deleted messages

    all_folders = account.root.glob('**/*') # Do it again because folder structure after previous deletions could change
    deleteItemsFromFolder(all_folders, "RecoverableItemsDeletions", "RecoverableItemsDeletions")
    """

    """
    The following section deleted collected email addresses that collected from incoming mail.
    You can comment it if you do not want to purge these addresses
    """
    all_folders = account.contacts.glob('**/*') # Do it again because folder structure after previous deletions could change
    deleteItemsFromFolder(all_folders, "Recipient Cache", "Recipient Cache")

    printNumberOfMessagesInContactsFolders(account)
    printNumberOfMessagesInSupportedFolders(account)

if __name__ == "__main__":
    main()
