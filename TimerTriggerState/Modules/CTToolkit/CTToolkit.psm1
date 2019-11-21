$clientAppInsights = $null

Class CTReturnObject {
    [String] $Status
    [String] $ErrorMessage
    [System.Collections.ArrayList] $Tracelog
    [System.Collections.ArrayList] $Input = ([Hashtable] $PSBoundParameters)
    [System.Collections.ArrayList] $Output = @()
}

Function Enable-CTApplicationInsight {
    param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $false,
            Position = 0)]    
        [String] $Key
    )
    #TODO REMOVE in Module
    $AssemblyPath = "C:\TFS\CT Script Repository\Prototypes\Application Insights\Microsoft.ApplicationInsights.dll"
    [System.Reflection.Assembly]::LoadFrom($AssemblyPath)
    
    $script:LastActionTime = get-date    
    $script:clientAppInsights = new-object Microsoft.ApplicationInsights.TelemetryClient
    $script:clientAppInsights.InstrumentationKey = $Key

}
#Not in use, because of issue with $PSBoundparameters as default value
Function Get-CTReturnObject {
    param(
        [String] $Status = "Unknown",
        [String] $ErrorMessage,
        [System.Collections.ArrayList] $Tracelog = $Script:Tracelog,
        [Hashtable] $InputParameters = ([Hashtable] $PSBoundParameters),
        [System.Collections.ArrayList] $Output = [System.Collections.ArrayList] @()
    )
    return [CTReturnObject] @{
        Status       = $Status
        ErrorMessage = $ErrorMessage
        Tracelog     = $Tracelog
        Input        = $InputParameters
        Output       = $Output
    }
}
Function Get-TraceLogMessage {
    <#
    .Synopsis
    Generates shorter strings from a long string.
    .DESCRIPTION
    Used to limit a tracelog file and split into multiple pieces, for use with a limited field, such as windows event log.
    #>
    param($RunbookName, $Message, $maxLength)
    $Message += "Runbook Name: $RunbookName`n"
    $Message += "Runas domain: $($env:userdomain)`n"
    #add last part of tracelog.
    $Message += "Last part of Trace Log:`n"

    #   $maxLength = 30000
    $NumberOfMessages = [Math]::Ceiling(($Script:Tracelog.Length / $maxLength))

    For ($i = 0; $i -lt $NumberOfMessages; $i++) {
        $start = (0 + ($maxLength * $i))
        $end = $maxLength

        if ($end -gt $Script:Tracelog.Length - $start) { $end = $Script:Tracelog.Length - $start }


        $Message = $Script:Tracelog.Substring($start, $end)

        if ($Message.Length -eq $maxLength) {
            $Message = $Message + "`ncontinued.."
        }

        $Message 
    }
}
Function Write-EventLogTraceLog {
    <#
    .Synopsis
    Splits tracelog into multiple strings and then writes them to the windows event log
    .DESCRIPTION
    uses the Get-TraceLogMessages function to split the tracelog and write it to event log.
    #>
    param($LogName, $SourceName, $RunbookName, $EntryType, $EventId, $Message)
        
    $Messages = Get-TraceLogMessage -RunbookName $RunbookName -Message $Message -maxLength 30000

    #write event log
    if (!([System.Diagnostics.EventLog]::SourceExists($SourceName))) {
        New-EventLog -LogName $LogName -Source $SourceName
    }

    foreach ($message in $messages) {
        Write-EventLog -LogName $LogName -Source $SourceName -EntryType $EntryType -EventId $EventId -Message $Message
    }
}
Function Add-Tracelog {
    <#
    .Synopsis
    Adds a tracelog message to tracelog 
    .DESCRIPTION
    Uses a script scope tracelog variable to have alle scopes write to a sinlg etracelog.
    Outputs each message to the Verbose stream.
    #>
    param($Message)

    $Message = "$(get-date) - $Message`n"
    Write-Verbose $Message
    if ([String]::IsNullOrEmpty($script:Tracelog)) {
        [System.Collections.ArrayList] $script:Tracelog = @()
    }
    $script:Tracelog.Add($Message)

    if ($null -ne $script:clientAppInsights) {
        $script:clientAppInsights.TrackEvent($Message);
        $script:clientAppInsights.Flush()
    }
}

Function Get-TraceLog {
    [System.Collections.ArrayList] $script:TraceLog
}
Function Write-TextLog {
    <#
    .Synopsis
    Writes the tracelog to a text file.
    .DESCRIPTION
    Writes the tracelog to a text file.
    If the text file looks "weird" in notepad, it is becausenotepad cannot alway list the line breaks correctly.
    Use another editor to view the log.
    #>
    param($Message, $LogFolderPath, $LogName, $StartTime)    
    #Write to file
    $Message + "`n" + ($script:TraceLog -join "`n") | Out-file $LogFolderPath\$($LogName)_$($StartTime.ToString("yyyyMMddhhmmss")).log -Encoding ASCII
}
Function ConvertTo-IndexedTable {
    <#
    .Synopsis
    Converts an array of objects to a array of hashtables for performance
    .DESCRIPTION
    Converts an array of objects to a hashtable of hashtables for performance
    The hashtable has takes one field as index, for example name 
    and the index field can be used to filter/sort much quicker than and array of objects. 
    (see example for details)
    .Example
    Get-CMDevice -CollectionName "All Systems" -Fast | ConvertTo-IndexedTable Name

    Name                           Value
    ----                           -----
    Object                         ...
    Name                           NCOP-CCMP-CCM01

    the index field is available directly, while the complete object is available in the object key of the hash table.
    #>
    param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $false,
            Position = 0)]
        $IndexFieldName,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromPipeline = $true,
            Position = 1)]
        [Object]$Object
    )

    #begin { $ReturnArray = @() } #New-Object System.Collections.ArrayList }
    process {
        @{ 
            $IndexFieldName = $Object.$IndexFieldName
            Object          = $Object
        }
    }
    #end { $ReturnArray }
}
Function Clear-Tracelog {
    <#
    .Synopsis
    Clears the tracelog 
    .DESCRIPTION
    Clears the tracelog
    #>

    $script:Tracelog = [System.Collections.ArrayList]  @()

}

Function Write-CTError {
    param($PSError)
    $CurrentError = $PSError
    $ErrorMessage = $CurrentError.Exception.Message + "`n $($CurrentError.InvocationInfo.PositionMessage)"
    Add-Tracelog -Message "Error in control runbook: $ErrorMessage"
    Write-Error -Message $ErrorMessage -ErrorAction Continue #output error without stopping, for logging purposes.

    if ($null -ne $script:clientAppInsights) {
        $telemtryException = New-Object "Microsoft.ApplicationInsights.DataContracts.ExceptionTelemetry"  
        $telemtryException.Exception = $CurrentError.Exception  
        $script:clientAppInsights.TrackException($telemtryException) 
        $script:clientAppInsights.Flush()
    }
}

$stateStores = @{ }
function Get-CTStateStore {

    param(
        [Parameter(Mandatory=$true)]
        [string]
        $storeName
        )
    $ctx = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage
    $tableName = $storeName
    $storageTable = Get-AzStorageTable -Name $tableName -Context $ctx -ErrorAction "silentlycontinue"
    if ($null -eq $storageTable) {
     
        $storageTable = New-AzStorageTable -Name $tableName -Context $ctx 
    }
    $property = @{ }
    $partitionKey = "configuration"
    $rowKey = "configuration"
    $Row = Get-AzTableRow -table $storageTable.CloudTable -partitionKey $partitionKey -rowKey $rowKey
    if ($null -eq $Row) {
        $Row = Add-AzTableRow `
            -table $storageTable.CloudTable `
            -partitionKey $partitionKey `
            -rowKey $rowKey -property $property
    } 

    $stateStores[$storeName] = $Row
    return $Row
}

Function Get-CTState {
    param(
        $Name,
        [Parameter(Mandatory=$true)]
        [string]
        $storeName,
        $DefaultValue
    )
    if($null -eq $storeName){
        throw "Store Name not provided"
    }
    $store = Get-CTStateStore -storeName $storeName
    if($null -eq $store.$Name){
        if($null -eq $DefaultValue){
            throw "State not found on name '$Name' and no default value is set"
        }

        Set-CTState -Name $Name -Value $DefaultValue -StoreName $storeName
        return $DefaultValue
    }
    return $store.$Name
} 

Function Set-CTState {
    param(
        $Name,
        $Value,
        [Parameter(Mandatory=$true)]
        [string]
        $storeName
    )
    $store = Get-CTStateStore -storeName $storeName
    $store | Add-Member -MemberType NoteProperty -Name $Name -Value $value -Force
    $result = Set-CTStateStore -storeName $StoreName -store $store
} 

function Set-CTStateStore {
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $storeName,
        $store
    )
    $ctx = New-AzStorageContext -ConnectionString $env:AzureWebJobsStorage
    $tableName = $storeName
    $storageTable = Get-AzStorageTable -Name $tableName -Context $ctx -ErrorAction "silentlycontinue"
    $storeUpdated = $store | Update-AzTableRow -table $storageTable.cloudTable

    return $storeUpdated

}
