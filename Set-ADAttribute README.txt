Created by Brandon Egbert

This powershell script attempts to automatically set an attribute in Active Directory based on emails sent from a particular email address, with a specified subject, and was received within the last x amount of time. It searches for users by First Name, Last Name, and Manager/Location. It also logs the script data in the local folder, as well as adds to a master CSV to look at every processed user

This script must be run normally, by a user with the access to change the attribute. Running it as an administrator will cause the script to fail.

Configuration settings

1. logDetail - determines what is output to the console and logs
    a. Success - Successes and progress
    b. Error - Outputs success, progress, and failure information
    d. Warnings - Displays warnings, errors, progress, and successes
    e. Information - Information to help understand what a script is doing as well as warnings, errors, and successes
    f. Verbose - Displays output that is useful for understanding the internal workings of the script, information, warnings, errors, and successes
    g. Debug - Debug messages can contain internal details necessary for deep troubleshooting, verbose, information, warnings, errors, and successes
    h. Progress - Only displays the progress of the script