# Office 365 and Azure AD Auditing

## Merging data from two sources

### Problem

Microsoft keeps useful Office 365 user data in two separate areas with two separate connection methods.

### Solution

This script creates separate connection objects with both the Microsoft Exchange Online service and the Microsoft Azure AD tenancy associated with it.
It then retrieves the full user list and full associated mailbox list from both and performs an inner join. It returns the resulting merged data as new Office 365 User data objects.
Depending on your command line arguments, you can choose to receive the object list returned in standard out to be piped elsewhere as you see fit in a script.
*OR* you can pass the $EZ argument which writes the values to a .csv file named after the default domain of the tenancy.
