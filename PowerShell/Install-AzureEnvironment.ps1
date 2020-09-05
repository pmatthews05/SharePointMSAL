<#
.SYNOPSIS
- You need to be already signed in with AZ CLI
- Creates:
	-Resource Group
	-KeyVault
	-Self Signed Certificate
	-Application Registration
	-Grants Access
 Defaults to the UKSouth location. This does not check if the name already exists.
.EXAMPLE
.\Install-AzureEnvironment.ps1 -Environment "cfcodedev" -Name:"sharepointMSAL"
.EXAMPLE
.\Install-AzureEnvironment.ps1 -Environment "cfcodedev" -Name:"sharepointMSAL" -Location:"westus"
#>
param(
	[Parameter(Mandatory)]
	[string]
	$Environment,
	[Parameter(Mandatory)]
	[string]
	$Name,
	[string]
	$Location = "uksouth"
)

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

[string]$Identifier = "$Environment-$Name"


Import-Module -Name:"$PSScriptRoot\MSAL" -Force -ArgumentList:@(
	$ErrorActionPreference,
	$InformationPreference,
	$VerbosePreference
)

Write-Information "Setting the Azure CLI defaults..."
az configure --defaults location="$Location"

Write-Information -Message:"Creating the $Identifier resource group..."
az group create --name "$Identifier"

Write-Information "Setting the Azure CLI defaults..."
az configure --defaults location=$Location group=$Identifier

[string]$KeyVaultName = ConvertTo-KeyVaultAccountName -name:$Identifier
Write-Information "Create a KeyVault..."
az keyvault create --name "$KeyVaultName"

Write-Information "Checking if the App already exists"
$AppReg = az ad app list --all --display-name "$Identifier" | ConvertFrom-Json
if($AppReg.length -eq 0) {
    Write-Information "Create an Application Registration"
    $AppReg = az ad app create --display-name "$Identifier" | ConvertFrom-Json
}

Write-Information "Store the Application Client ID in Keyvault"
Set-ApplicationIdToKeyVault -ApplicationName:"$($AppReg.DisplayName)" -ApplicationAppId:"$($AppReg.appId)" -KeyVaultName:"$KeyVaultName"

Write-Information -MessageData:"Checking if the Service Principal exists for the $($AppReg.appId)..."
$servicePrincipal = az ad sp list --spn $($AppReg.appId) | ConvertFrom-Json
if ($servicePrincipal.Length -eq 0) {
    Write-Information -MessageData:"Creating the Service Principal for the $($AppReg.appId)..."
        $servicePrincipal = az ad sp create --id $($AppReg.appId) | ConvertFrom-Json
}

Write-Information "Create a self signed Certificate and put in KeyVault"
Set-SelfSignedCertificate -ApplicationRegistration:$AppReg -Identifier:$Identifier -KeyVaultName:$KeyVaultName

#Only seems to work with SharePoint Only. 
#See blog for correct way with any permission - https://cann0nf0dder.wordpress.com/2020/06/24/grant-application-and-delegate-permissions-using-an-app-registration/
Write-Information "Giving Application SharePoint Full Control Permission"
Set-AppPermissionsAndGrant -ApplicationRegistration:$AppReg -RequestedPermissions:"Application.SharePoint.FullControl.All"


