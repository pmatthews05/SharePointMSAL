# Connect to SharePoint using MSAL and KeyVault

This is a demo project for the blog at cann0nf0dder.wordpress.com.

A console application using MSAL and .NET Core to connect to SharePoint Online.

## Setup

### Prepare Azure
- Create an Azure AD app registration (with the Client Certificate)
- Create a KeyVault
- Store the Certificate in the KeyVault
- Store the ClientID in the KeyVault Secrets 
- Grant Application permissions for SharePoint > Sites.FullControl.All

This has been automated in a PowerShell Script.
In the [PowerShell folder](../Powershell/Install-AzureEnvironment.ps1) run the .\install-AzureEnvironment.ps1 

```ps1
az login
.\Install-AzureEnvironment.ps1 -Environment:<TenantName> -Name:SharePointMSAL
```
This will create the following in your environment if your TenantName is Contso
- <b>App Registration</b>: Contso-SharePointMSAL, granted with SharePoint > Sites.FullControl.All
- <b>Key Vault</b>: Contso-SharePointMSAL (<i>Note:Will be truncated to 24 characters if longer</i>)
- <b>CertificateName stored in KeyVault as</b>: Contso-SharePointMSAL
- <b>ClientId stored in KeyVault Secret as</b>: ConstoSharePointMSAL


### Prepare the Console Application
Update the appsettings.json file for your environment

```json
{
    "environment": "<tenantName>",
    "name": "SharePointMSAL",
    "site": "<relative URL e.g, /sites/teamsite>"
}
```