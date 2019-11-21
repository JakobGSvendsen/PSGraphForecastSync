# Input bindings are passed in via param block.
param($Timer, $TriggerMetadata)
<#
  Forecast Outlook sync script by Jakob Gottlieb Svendsen - www.ctglobalservices.com
  Goto https://id.getharvest.com/developers
  Create Personal Access Token
  Insert Access token in $Token below
  Optionally set day start time.

  Notice:
  Set the notes in forecast to start with time to make the appointment in outlook appear at that time.
  in format xx:xx etc. 13:00 (24 hour clock)
 #>

 #Config
$CategoryName = "ForecastV2"  
$DayStartTime = "09:00"
$AddReminder = $false
$ReminderMinutesBeforeStart = 15
$propertyExtName = "extquv76e3t_forecast"
$AADTenant = $env:GraphAADTenant
$clientId = $env:GraphClientId
$clientSecret = $env:GraphClientSecret 
$TokenForecast = $env:GraphTokenForecast 
$ForecastAccountId = $env:GraphForecastAccountId

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()
$FunctionName = $TriggerMetadata.FunctionName
$PSDefaultParameterValues = @{
    "*CTState*:storeName" = $FunctionName
}

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

try {
    $ErrorActionPreference = "stop"
    #Add local azure function modules path first in list
    $env:PSModulePath = "$pwd\$($TriggerMetadata.FunctionName)\Modules;" + $env:PSModulePath

    Write-Information ("pwd:" + $pwd)
    Import-Module CTToolkit -Force
    Import-Module MicrosoftGraphAPI -Force

    $updated_since = Get-CTState -Name "updated_datetime" -DefaultValue (Get-Date).AddDays(-90)
    write-information "updated_datetime: $updated_since"

    #Get appointments
    $start_datetime = (Get-Date).AddMonths(-3)
    $end_datetime = (Get-Date).AddYears(1)

    #Get forecast Data
    $headersForecast = @{
        'Forecast-Account-ID' = $ForecastAccountId
        'Authorization'       = "Bearer $TokenForecast"
        'User-Agent'          = "PowerShell"
    }

    #get info about users
    $response = invoke-webrequest -Uri  "https://api.forecastapp.com/people" -Headers $headersForecast -ContentType "application/xml" -Method get
    $people = ConvertFrom-JSON $response.Content
    $peopleFiltered = $people.people | Where-Object archived -eq $false

    #get projects
    $response = invoke-webrequest -Uri  "https://api.forecastapp.com/projects" -Headers $headersForecast -ContentType "application/xml" -Method get
    $projects = (ConvertFrom-JSON $response.Content).projects

    #get clients
    $response = invoke-webrequest -Uri  "https://api.forecastapp.com/clients" -Headers $headersForecast -ContentType "application/xml" -Method get
    $clients = (ConvertFrom-JSON $response.Content).clients

    #Get Token
    #Connect-Graph -Scopes "Calendars.ReadWrite","User.ReadWrite.All"
    $OAuthResult = . Get-GraphAuthToken -AADTenant $AADTenant -ClientId $clientId -ClientSecret $clientSecret
    $TokenGraph = $OAuthResult.access_token


    #Get users
    $url = "https://graph.microsoft.com/beta/users" ;
    $result = $null
    $usersGraph = Invoke-GraphGet -url $url  -Token $TokenGraph -ErrorAction "continue" -All 
    $peopleGraphForecast = $peopleFiltered | Where-Object { $_.email -in $usersGraph.userPrincipalName }

    write-output "Processing $($peopleGraphForecast.count) users"

    #pilot test
    #$peopleGraphForecast = $peopleGraphForecast | Where-Object Email -eq "jgs@ctglobalservices.com"
    #Foreach Person
    foreach ($person in $peopleGraphForecast) {   
        $UserPrincipalName = $person.email
        write-output "$UserPrincipalName"
        $url = "https://graph.microsoft.com/beta/users/$UserPrincipalName/calendar/calendarView?startDateTime={0:yyyy-MM-ddTHH:mm:ss.fffffff}&endDateTime={1:yyyy-MM-ddTHH:mm:ss.fffffff}&`$top=10000&`$filter=categories/any(c: c eq '$CategoryName')&`$select=*,$propertyExtName" -f $start_datetime, $end_datetime ;
        
        $eventsForecast = @()
        $eventsForecast += Invoke-GraphGet -url $url  -Token $TokenGraph  -ErrorAction "continue" -All
        write-output "Total Number of calendar events in outlook: $($eventsForecast.Count) $CategoryName"
        
        #get persons updated assignments
        $response = invoke-webrequest -Uri  "https://api.forecastapp.com/assignments?person_id=$($Person.id)&start_date=$($start_datetime.ToString("yyyy-MM-dd"))&end_date=$($end_datetime.ToString("yyyy-MM-dd"))" -Headers $headersForecast -ContentType "application/xml" -Method get
        $assignments = (ConvertFrom-JSON $response.Content).assignments;
        $assignmentsFormatted = $assignments | select-object *, @{n = "project"; e = { $projects | ? id -eq $_.project_Id } } | select *, @{n = "client"; e = { $clients | ? id -eq $_.project.client_id } } 
        $assingmentsFiltered = $assignmentsFormatted | where-object { $_.project.name -ne "Time Off" -and $_.project.archived -eq $False } | Sort-Object -Property Id
        $assingmentsUpdated = $assingmentsFiltered | where-object { $_.updated_at -gt $updated_since } | Sort-Object -Property Id

        write-output "Total Number of assignments in Forecast: $($assingmentsFiltered.Count)"
        write-output "Processing $($assingmentsUpdated.Count) updated assignments"
    
        Foreach ($ass in $assingmentsUpdated) {
            $forecastId = $ass.id
            $forecastProjectId = $ass.project.id

            #Prep content
            if ($ass.notes -like "??:??*") {
                #if notes start with xx:xx etc. 13:00
                $starttimeHour = $ass.notes.SubString(0, 5)
            }
            else {
                $starttimeHour = $DayStartTime
            }

            $starttime = ([Datetime] "$($ass.start_date) $starttimeHour").ToString("yyyy-MM-ddTHH:mm:ss.fffffff")
            $endtime = ([Datetime] "$($ass.end_date) $starttimeHour").AddSeconds($ass.allocation).ToString("yyyy-MM-ddTHH:mm:ss.fffffff")

            if ($ass.notes) {
                $Subject = $ass.project.name + "- " + ($ass.notes -split "`n")[0]
            }
            else {
                $Subject = $ass.project.name
            }

            $Body = $ass.notes
            if ($AddReminder) {
                $ReminderMinutesBeforeStart = $ReminderMinutesBeforeStart
            }

            $CreateEventBody = @"
{
    "subject": "$Subject",
    "body": {
      "contentType": "HTML",
      "content": "$Body"
    },
    "categories": [ "$CategoryName" ],
    "start": {
        "dateTime": "$StartTime",
        "timeZone": "Europe/Berlin"
    },
    "end": {
        "dateTime": "$EndTime",
        "timeZone": "Europe/Berlin"
    },
    "location":{
        "displayName":"$($ass.client.name)"
    }
  }
"@
            $bFound = $false
            $eventExisting = $eventsForecast.Where( { $_.$propertyExtName.forecastId -eq $forecastId })
            if ($eventExisting.Count -eq 1) {
                if ([DateTime]$ass.updated_at -gt $eventExisting[0].lastModifiedDateTime) {
                    "Updated: $forecastId"   
                    $Url = "https://graph.microsoft.com/beta/users/$UserPrincipalName/events/$($eventExisting[0].id)"
                    $result = Invoke-GraphRequest -url $url  -Token $TokenGraph -Method Patch -Body $CreateEventBody
                    continue
                }
                else {  
                    write-verbose "Skipped: $forecastId"   
                    continue
                }

            }
            elseif ($eventExisting.Count -eq 0) {
                "New: $($ass.Id)"  
            
                $Url = "https://graph.microsoft.com/beta/users/$UserPrincipalName/events"
                $result = Invoke-GraphRequest -url $url  -Token $TokenGraph -Method Post -Body $CreateEventBody
                $eventId = $result.id
                $ForecastEventBody = @"
{
    "$propertyExtName": {
        "forecastId":"$forecastId",
        "forecastProjectId":"$forecastProjectId"
    }
}
"@
                $Url = "https://graph.microsoft.com/beta/users/$UserPrincipalName/events/$eventId"
            
            
                $retryLimit = 10
                $retryCount = 0
                do {
                    try {
                        $retry = $false
                        "Add forecast props - Attempt: $retryCount"
                        #Start-Sleep -Milliseconds 500
                        $result = Invoke-GraphRequest -url $url  -Token $TokenGraph -Method Patch -Body $ForecastEventBody
                     
                    }
                    catch {
                        if ($_.Exception.Message -like "*Please retry the request*") {
                            if ($retryCount -ge $retryLimit) {
                                throw $_   
                            }
                            $retry = $true
                            $retryCount++
                            Start-Sleep -Milliseconds 500
                        }
                        write-output $_
                    } #catch

                } while ($retry) #While
            } #elseif count -eq 0


        } #Foreach ($ass in $assingmentsFiltered) {

        #Cleanup Future events that have been removed from forecast
        $eventsToDelete = $eventsForecast.Where( { $_."$propertyExtName".forecastId -notin $assingmentsFiltered.id } )
        "Found $($eventsToDelete.Count) for deletion"
        foreach ($eventDelete in $eventsToDelete) {
            #Delete event
            $eventDeleteId = $eventDelete.id
            $Url = "https://graph.microsoft.com/beta/users/$UserPrincipalName/events/$eventDeleteId"
            write-verbose "Deleting event: $($eventDelete.Id) - $($eventDelete.subject) - StartDate: $($eventDelete.start.dateTime)"
            "Deleting event: $($eventDelete.subject) - StartDate: $($eventDelete.start.dateTime)"
            Invoke-GraphRequest -url $url  -Token $TokenGraph -Method Delete
        }
    } #foreach person

    Set-CTState -Name "updated_datetime" -value (get-date)
}
catch {
    throw $_
    #    [Environment]::Exit(1)


}
#>
