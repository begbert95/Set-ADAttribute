Created by Brandon Egbert

This powershell script attempts to automatically set an attribute in Active Directory based on emails sent from a particular email address, with a specified subject, and was received within the last x amount of time. It searches for users by First Name, Last Name, and Manager/Location. It also logs the script data in the local folder, as well as adds to a master CSV to look at every processed user

This script must be run normally, by a user with the access to change the attribute. Running it as an administrator will cause the script to fail.

