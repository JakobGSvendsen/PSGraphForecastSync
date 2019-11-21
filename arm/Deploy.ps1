
cd $psscriptroot
<#
Connect-AzAccount
Select-AzSubScription -SubscriptionName "Microsoft Azure Sponsorship"
#>
$ResourceGroupName = "testSync"
New-AzResourceGroup -Location "WestEurope" -Name $ResourceGroupName -Force
New-AzResourceGroupDeployment `
-ResourceGroupName $ResourceGroupName `
-TemplateFile '.\azuredeploy.json' `
-TemplateParameterFile '.\azuredeploy.parameters.json' -Verbose

