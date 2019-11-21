Register-PSRepository `
-Name GraphPowerShell `
-SourceLocation https://graphpowershellrepository.azurewebsites.net/nuget

# Installing the Graph PowerShell module for the Beta API 
Install-module Microsoft.Graph.Beta -Repository GraphPowerShell

Import-Module Microsoft.Graph.Beta -Prefix "Graph"


Connect-Graph -Scopes "User.Read.All" -ForceRefresh -Verbose

Get-User -Top 10 -Select Id, DisplayName, BusinessPhones | Format-Table Id, DisplayName, BusinessPhones



get
dir variable:

Get-Command -Module Microsoft.Graph.Beta* | out-gridview

#run in windows PowerShell

install-module AzureAD
Import-module AzureAD
Connect-AzureAD


Get-AzureADDirectoryRole