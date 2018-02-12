
<#
.Synopsis
Connects to Office 365 and Azure AD and retrieves list of user objects.

.Description
Connects to the MS Office 365 PowerShell and the MS Azure AD PowerShell. It then joins together the user list data via UserPrincipalName and returns the list of objects


.Parameter O365UserName
Microsoft Office 365 Tenant UserName
.Parameter Office365Password
Microsoft Office 365 Tenant Password
.Parameter EZout
Do not connect to AD Connect even if it is installed.

.Notes
Only connects to MS Azure if it is installed.
Only connects to AD Connect if it is installed.
Default O365UserName can be specified in the script PARAM section.

#>
[CmdletBinding()]
#Accept input parameters 
param( 
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $Office365Username, 
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)] 
    [string] $Office365Password,
    [Parameter(Mandatory=$false)]
	[switch]$EZ

) 
#Main 
Function Main {

 
    #Remove all existing Powershell sessions 
    Get-PSSession | Remove-PSSession 

    try{
        #Call function to connect to MSOL service
        Write-Progress -Activity "Connecting to Office 365..."
        ConnectTo-365Tenancy -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password

        #Call ConnectTo-ExchangeOnline function with correct credentials 
        Write-Progress -Activity "Connecting to Exchange Online..."
        ConnectTo-ExchangeOnline -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password
       
        #gather all mailboxes from Office 365
        Write-Progress -Activity "Receiving User Lists..." -PercentComplete 0
        $objUsers = get-mailbox -ResultSize Unlimited
    
        #gather all users in 365 tenancy
        Write-Progress -Activity "Receiving User Lists..." -PercentComplete 50 -CurrentOperation "1 of 2 completed."
        $objMsolUsers = Get-MsolUser -All
        Write-Progress -Activity "Receiving User Lists..." -PercentComplete 100 -Completed -Status "2 of 2 completed."

        #Grab default domain name and save in a variable
        $default365Domain = (Get-MsolDomain | ?{$_.isDefault -eq $true}).Name

        #call Knitting Function if EZ switch argument is pulled
        if($EZ -eq $false){ Knit-UserData}
        else {
         Write-Progress -Activity "Writing List to file..." -id 1
         Knit-UserData | Export-Csv ($default365Domain+'_userlist.csv') -NoTypeInformation
         }

    }
    
    Finally {
        #Clean up session 
        Get-PSSession | Remove-PSSession
    }
} 

function Knit-UserData {
    #Iterate through all users     
    ForEach ($objUser in $objUsers) {   
        $indexofProgress = ($objUsers.IndexOf($objUser) / $objUsers.Length) * 100
        $indexofProgress = [math]::Round($indexOfProgress,2)

        $365Account = $objMsolUsers | ?{$_.UserPrincipalName -eq $objUser.UserPrincipalName}
        
        New-Object -TypeName PSObject -Property @{
            Alias = $objUser.Alias
            DisplayName = $objUser.DisplayName
            EmailAddress = $objUser.PrimarySMTPAddress
            LastLogon = $(get-mailboxstatistics -Identity $objUser.UserPrincipalName -warningaction SilentlyContinue).LastLogonTime
            MailboxType = $objUser.RecipientTypeDetails
            isLicensed = $365Account.isLicensed
            }
        Write-Progress -Activity "Fetching and knitting the user data together…" `
            -PercentComplete $indexofProgress -CurrentOperation "$indexOfProgress% through the userlist" -Status "Hang Tight"
         
    }
}


############################################################################### 
# 
# Function ConnectTo-ExchangeOnline 
# 
# PURPOSE 
#    Connects to Exchange Online Remote PowerShell using the tenant credentials 
# 
# INPUT 
#    Tenant Admin username and password. 
# 
# RETURN 
#    None. 
# 
############################################################################### 
function ConnectTo-ExchangeOnline 
{    
    Param(  
        [Parameter( 
        Mandatory=$true, 
        Position=0)] 
        [String]$Office365AdminUsername, 
        [Parameter( 
        Mandatory=$true, 
        Position=1)] 
        [String]$Office365AdminPassword 
 
    ) 
         
    #Encrypt password for transmission to Office365 
    $SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force     
     
    #Build credentials object 
    $Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365AdminUsername, $SecureOffice365Password 
     
    #Create remote Powershell session 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Office365credentials -Authentication Basic –AllowRedirection         
 
    #Import the session 
    Import-PSSession $Session -AllowClobber | Out-Null 
} 

############################################################################### 
# 
# Function ConnectTo-365Tenancy 
# 
# PURPOSE 
#    Connects to Microsoft Online Remote PowerShell using the tenant credentials 
# 
# INPUT 
#    Tenant Admin username and password. 
# 
# RETURN 
#    None. 
#
# REQUIRED SOFTWARE
#    Microsoft ONline Services Sign-in Assistant
#    Windows Azure Active Directory Module for Windows Powershell
############################################################################### 
function ConnectTo-365Tenancy 
{    
    Param(  
        [Parameter( 
        Mandatory=$true, 
        Position=0)] 
        [String]$Office365AdminUsername, 
        [Parameter( 
        Mandatory=$true, 
        Position=1)] 
        [String]$Office365AdminPassword 
 
    ) 
         
    #Encrypt password for transmission to Office365 
    $SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force     
     
    #Build credentials object 
    $Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365AdminUsername, $SecureOffice365Password 
     
    #Create remote Powershell session 
    Connect-MsolService -Credential $Office365Credentials | Out-Null 
} 


# call main
. Main
