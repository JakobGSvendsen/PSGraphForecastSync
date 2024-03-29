#
# Module manifest for module 'MicrosoftGraphAPI'
#
# Generated by: Jakob Gottlieb Svendsen
#
# Generated on: 26-04-2016
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'MicrosoftGraphAPI.psm1'

# Version number of this module.
ModuleVersion = '0.3.0'

# ID used to uniquely identify this module
GUID = '621fafe4-cf0f-4d53-8c40-1b70d12b56f4'

# Author of this module
Author = 'Jakob Gottlieb Svendsen'

# Company or vendor of this module
CompanyName = 'CTGlobal'

# Copyright statement for this module
Copyright = '(c) 2018 Jakob Gottlieb Svendsen. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Module for accessing Microsoft Graph API'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
 #RequiredAssemblies = @('Microsoft.IdentityModel.Clients.ActiveDirectory.dll','Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll')

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
ModuleList = @('MicrosoftGraphAPI.psm1')

# List of all files packaged with this module
FileList = 'Microsoft.IdentityModel.Clients.ActiveDirectory.dll', 
               'Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll', 
               'MicrosoftGraphAPI-Automation.json'

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

