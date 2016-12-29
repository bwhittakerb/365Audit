<#
.Synopsis
Connects to Office 365 and Azure AD

.Description
Connects to the MS Office 365 PowerShell and the MS Azure AD PowerShell.
Will also connect to AD Connect (but should not be used with DirSync, suggest upgrading)
For issuing remote PowerShell commands to these Online Services.
NOT for Microsoft Windows Azure IaaS and PaaS services (VMs, etc).

.Parameter O365UserName
Microsoft Office 365 Tenant UserName
.Parameter Disconnect
Should we Disconnect an existing O365 session?
.Parameter NoAzure
Do not connect to MS Azure even if it is installed.
.Parameter NoADConnect
Do not connect to AD Connect even if it is installed.

.Notes
Only connects to MS Azure if it is installed.
Only connects to AD Connect if it is installed.
Default O365UserName can be specified in the script PARAM section.

Can make a short cut to this with something like:
 %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe -NoExit c:\tools\Connect-Office365.ps1
Then just need to double-click to run and be prompted to log in.

Last Updated Dec 20, 2016 by Saul Ansbacher / CompuVision Systems Inc.
#>

# Default O365 login is listed below in PARAM!!!

# param needs to be first executable line
param (
	#O365UserName is the name to use to connect to Office 365
	# If NO name provided on command line you can add a default one here, or it can prompt if left blank
	[Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
	[string]$O365UserName = "gwkwadmin@greatwestkenworth.com",		# CHANGE this OR remove for no default username!
	# Should we Disconnect?
	[Parameter(Mandatory=$false)]
	[switch]$Disconnect,
	# Skip Azure AD?
	[Parameter(Mandatory=$false)]
	[switch]$NoAzure,
	# Skip AD Connect?
	[Parameter(Mandatory=$false)]
	[switch]$NoADConnect

)

# All variables must be declared before evaluation
Set-PSDebug -Strict 

# To connect via PowerShell 2.0+, needs .NET 4.5: [Do NOT use Exchange's EMS as some commands will conflict, use regular PowerShell!]
# Test to confirm this is NOT being run in an Exchange Management Shell session
$EMSCommands = gcm *ExchangeServer 		# These command would exist in EMS but NOT in O365 or regular PS
if ($EMSCommands) {
	write-host
	write-host -foreground red 'ERROR!!! Do NOT run this in Exchange Management Shell(EMS)!' 
	write-host @"

Please use a regular, non-Exchange, PowerShell session.

"@
	exit  #exit script
}

# Disconnect function, persists after the script ends.
function global:Disconnect-Office365 {
	# Can't test $O355Session.state because if it times out it will be Broken, but remains Broken even if reconnected
	write "Disconnecting Office 365 session"
	write-host
	try {
		Remove-PSSession $O365Session
	} catch {
		Write-host "WARNING: Disconnecting the Office 365 Session failed for some reason."
		write-host "The Session state is: $($O365Session.State)"
	}
	# Reset the Windowtitle to something generic:
	$host.UI.RawUI.WindowTitle = ("$($env:username): Microsoft Windows PowerShell")
	# reset the prompt to what it (probably) was before:
	function Global:prompt {$(if (test-path variable:/PSDebugContext) { '[DBG]: ' } else { '' }) + 'PS ' + $(Get-Location) + $(if ($nestedpromptlevel -ge 1) { '>>' }) + '> '}
}

if  ($Disconnect) {
	Disconnect-Office365
	exit    # exit script
}  

# If needed you could embed the username AND password in the script (which is insecure of course) to bypass prompting:
# $O365UserName = "user@domain.com"
# $SecPassword = ConvertTo-SecureString "S0mePassw0rd" -AsPlainText -force
# $O365Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecPassword
# (May also be possible to encode the SecurePassword (using convertfrom-securestring) and just embed that, which would be better) [untested]

write-host
write-host -foreground magenta "Please provide the Office 365 password in the pop-up box..."
write-host
try {
	# Making this global in scope so the creds can be used in other commands later if needed.
	$Global:O365Credentials = Get-Credential  $O365UserName
	# Note, if ever needed: $PlainPassword = $O365Credentials.GetNetworkCredential().Password 
} catch {
	write-host -foreground red 'ERROR!!! You MUST enter a Username and Password!' 
	write-host
	# Clear the disconnect function
	function global:Disconnect-Office365 {}
	exit  #exit script
}

# This is for Office 365 / Exchange Online, for EOP see below.
write "Please wait, connecting to Office 365"
write-host 
# Use $global:varname scope so variable perists outside of this script so the session can be ended later. Most people forget this in their O365 scripts!

# For Office 365 use:
$Global:O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365Credentials -Authentication Basic -AllowRedirection
# Another URI could be: https://outlook.office365.com/powershell-liveid/

# For EOP use:
# $Global:O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ -Credential $O365Credentials -Authentication Basic -AllowRedirection

if ($O365Session.State -eq "Opened") {
	write-host	
	write-host "SUCCESS! Connected to Office 365" -foreground Green
} else {
	write-host "FAILURE! The Connection failed for some reason!" -Foreground yellow -Background red
	write-host " (probably a wrong password...)"
	write "The Session state is: $($O365Session.State)"
	# Clear the disconnect function
	function global:Disconnect-Office365 {}
	exit  #exit script
}

Import-PSSession $O365Session -AllowClobber
write-host
# if it worked then the following should work:
#Get-AcceptedDomain

# Attemp to connect to Azure AD and import those commands using the SAME creds: 
if ( (test-path "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\MSOnline\Microsoft.Online.Administration.Automation.PSModule.dll") -and !$NoAzure ) {
	write-host
	write-host "Connecting to Microsoft Azure AD PowerShell. Specify -NoAzure to skip this."
	if ($PSVersionTable.PSCompatibleVersions -contains '3.0') {
		import-module MSOnline
		connect-msolservice -credential $O365Credentials
		# If it worked then the following should work:
		# Get-MsolAccountSku
		write-host
		write-host "Microsoft Azure AD (MS-Online) PowerShell commands loaded."  -foreground "green"
		write-host
	} else {
		write-host "WARNING: PowerShell 3.0 is NOT available!" -foreground Yellow
		write-host "MS Online / Azure AD PowerShell is installed but can't be loaded! Skipping..."
	}
}

# Attemp to import AD Connect PS (the new DirSync)
# List all modules available: get-module -ListAvailable
if ( ( Get-Module -ListAvailable ADSync) -and !$NoADConnect ) {
	# May need to add the same $PSVersionTable.PSCompatibleVersions check here, except I think it is a requirement for AD Connect anyway...
	write-host
	write-host "Importing AD Connect PowerShell. Specify -NoADConnect to skip this."
	import-module ADSync
	# If it worked then the following should work:
	# Start-ADSyncSyncCycle Delta
	write-host
	write-host "AD Connect PowerShell commands loaded."  -foreground "green"
	write-host
}

write-host @"

Please remember to disconnect with:

"@
write-host 'Disconnect-Office365'  -foreground "yellow"
write-host
# this will pull the full path to this script:
write-host "OR run: $($MyInvocation.MyCommand.Definition) -Disconnect"


write-host @"

If you leave this window long enough it will time out and you will be prompted for the password again.
Do NOT import the DirSync module in this session, doing so will block access to the
Office 365 and/or Azure AD PowerShell commands. Use a fresh PowerShell session (or upgrade to AD Connect).
AD Connect works fine, auto-loading the module if available.
"@ # This is because Import-Module DirSync actually starts a NEW PowerShell instance (if 2.0). It may be OK on newer PS.
# Alternatively you could import DirSync FIRST and then run this script to connect. Or just upgrade to AD Connect!

# Change the Window title
$host.UI.RawUI.WindowTitle = ("Microsoft Office 365 PowerShell: $($O365Credentials.username)")
# Change the prompt so we know we're connected to Office 365
function Global:prompt {$cwd = (get-location).Path;$host.UI.Write("Green", $host.UI.RawUI.BackGroundColor, "[PS O365]");" $cwd>"}
# To see all available colours use: [enum]::GetValues([System.ConsoleColor]) | Foreach-Object {Write-Host $_ -ForegroundColor $_}
# To see what the prompt is now use: cat function:prompt
