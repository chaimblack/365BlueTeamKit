# 365BlueTeamKit

### Disclaimer:
These scripts have not been created by Microsoft and are being published AS IS. Use these scripts at your own risk and verify the data is accurate. Complete testing on all operating systems and PowerShell versions have not been done.

## Introduction:
Scripts in this repository include many scripts for producing reports on user and tenant settings for Office 365. A good portion of these scripts are designed to aid an investigation on business email compromises and produce security related reports.

A couple things to note about me and these scripts:

* These scripts are always under development as Microsoft changes things at a pace quicker than they can change all their licensing names. Having created these scripts and can recognize any results which do not seem accurate, I trust the output of these scripts, but things may break due to changes and I take NO responsibility for the validity of any of data these scripts output.
* These scripts typically export all output in Excel files using the 'ImportExcel' module created by Doug Fink (he is my hero). Using Excel files as output, many of the spreadsheets exporting data includes data in multiple spreadsheets and formatting.
* When many of the scripts saves the Excel file, it will save it in read-only mode, capture a hash, and export the hash in a read-only csv file in the location of the file.
* All scripts performing audit log searching uses a script to lookup IP addresses to geolocation. As including the internet service provider is critical to reviewing geo-lookups, I use an API from ipgeolocation.io which is free for 1K lookups per day and 30K per month. To reduce on IP lookups, the Get-IPLookup uses a built-in cache file which saves 150 IP addresses to prevent you from burning up all searches in several minutes (more on this later). 
* The output of all reports save to the following location by default but can typically be modified: C:\PSOutPut\Module Name\Company Name
* With my scripts utilizing data from several Office 365 services, these scripts require a connection to the Exchange Online Managment, AzureAD and MSOnline PowerShell modules. The Connect-Tenant script handles this connection.
## Required Modules
Many of the scripts will automatically check for these following modules and will automatically install them if they are not found.
### ImportExcel
PowerShell Module for exporting to Excel files
Download path https://github.com/dfinke/ImportExcel

To Install from the PowerShell Gallery:
```
Install-Module -Name ImportExcel
```
### ExchangeOnlineManagement
PowerShell Module for connecting to Exchange Online. Minimum version is 2.0.5

To Install from the PowerShell Gallery:
```
Install-Module -Name ExchangeOnlineManagement 
```
## MSOnline
PowerShell Module for connecting to MSOnline. 

To Install from the PowerShell Gallery:
```
Install-Module -Name MSOnline  
```
## AzureAD
PowerShell Module for connecting to AzureAD. 

To Install from the PowerShell Gallery:
```
Install-Module -Name AzureAD
```
## Geo-Lookups
Utilizing any scripts which search the audit logs, it is REQUIRED to first obtain an API key from ipgeolocation.io to obtain ISP information. The "NoLookup" switch has been built into several scripts to allow for searching without any geolookups. 

I chose this service for the following reasons:

* This service offers a free solution which allows for low usage of 1K lookups per day and 30K per month. To reduce on IP lookups, the Get-IPLookup script uses a built-in cache file which saves 150 IP addresses to prevent you from burning up all searches in several minutes. For users requiring more lookups, there are paid options as well and can be found at https://ipgeolocation.io/pricing.html
* This service includes the Internet Service Provider (ISP) information. When looking at an IP lookup while investigating a security incident, it is critical to know the ISP of the IP address since it will provide more context to if it is legitimate. Simply know the location of an IP address without knowing the IP address can through off an investigation. As an example, if the IP address is owned by a mobile carrier, the location may likely not be accurate, or if the IP address is owned by a major company, it can provide insight if the login was expected.

As stated previously, if you are not trained in cyber security incident response and are familiar with how normal audit logs look like, looking at the audit log geolocation information can cause incorrect conclusions.

Once you obtain an API key, you will need to modify the "Get-IPLookup" script using the instructions in the next section.

## Importing and running these scripts:
To run these scripts, you can either import one of these scripts manually, or importing all of them in the following manner:
1. Download the scripts.
2. Verify Execution Policy allows for running these scripts (Restricted will not allow it).
3. Place the folder with the scripts in the following location: C:\PSModules\
4. Import the scripts using the information below, by placing path of folder containing scripts in the quotes on the next line.:

```
#Example where script location is "C:\PSModules\365BlueTeamKit"
$psdir="C:\PSModules\365BlueTeamKit"  
Get-ChildItem "${psdir}\*.ps1" | %{.$_} 
```
5. Add the API key the Get-IPLookup script. Insert API key on line 53 inside of the quotes.

# Scripts
## Get-365InboxRules.ps1

Get-365InboxRules was created to audit user inbox rules in Office 365. 

Attackers typically create malicious inbox rules after compromising an account, often redirecting, forwarding, moving, or deleting mail. 
This script captures all user inbox rules, looks at several attributes which are often misused by attackers, and exports it to a xlsx (default) or CSV. It also outputs the raw results to json.

As a reminder, always check user inbox rules when investigating an email compromise. When running this script in larger companies, I often find either active email compromises, or remnants of prior compromises which was not fully cleaned up.

### Usage ##

 Parameter | Possible Values 
--- | --- |
-NoLaunch | [Switch] (Optional) - Prevents the output folder from launching after completion of command. Default is set to open folder.
-Path | [String] (Optional) - Sets the path of the output. Default path is C:\PSOutput\Get-365InboxRules\
-ExportCSV | [Switch] (Optional) - By default, this script exports as an Excel file. Setting this switch exports as CSV instead.
-OutObject | [Switch] (Optional) -  Outputs all rules. Can be used if you want to pipe it to something else.
-OutRawObject | [Switch] (Optional) - Outputs all raw rules. Can be used if you want to pipe it to something else.
-Username | [Switch] (Optional) -By default, this script searches an entire Office 365 tenant. You can use this to specify a single user to search instead.

### Output
The script either outputs an XLSX file (default) or two CSV files, and a JSON file with the raw output.

In its default output of XLSX file, it generates a main worksheet with results, and if it detects known suspicious rules, it will put them in another worksheet named "Suspicious" 

If the "ExportCSV" switch is defined, it will instead output a main worksheet with the results, and if it detects known suspicious rules, it will put them in another CSV file named "Suspicious" 

It will also export the raw unfiltered data as a JSON file.

Column explanation:
Column | Explanation 
--- | --- |
User | Username
Name | DisplayName
Enabled | Defines if rule is enabled
Delete | Defines if rule can delete mail
ApplyAll | Defines if rule can apply to all incoming mail
Date | Defines if rule can apply to emails before or after a date
Size | Defines if rule can be dependent to size of mail
Move | Defines if rule can move mail to another folder
MarkAsRead | Defines if rule can mark mail as read
FwdorRedir | Defines if rule can forward or redirect mail
FwdorRedirExt | Defines if rule can forward or redirect mail to an external domain
Rule | Defines rule
FwdorRedirTo | Defines where mail rule is forwarding to (if applicable)
FwdorRedirExtTo |Defines where mail rule is forwarding to if recipient is an external domain (if applicable)
UserObjectID | Defines User ObjectID

### Issues

Known Issues:
This script uses 'Get-InboxRule' which requires a mailbox to be specified. This script uses a foreach loop to capture the data for all mailboxes but can run into throttling issues if there are too many mailboxes.

## Other great resoures

Hawk: https://github.com/T0pCyber/hawk

DFIR-O365RC: https://github.com/ANSSI-FR/DFIR-O365RC