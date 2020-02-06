param(
#------------------------------------------------
# Static Parameters
#------------------------------------------------
    [string]$EaseApiUrl,
    [string]$EaseUser, 
    [string]$EasePassword, 
    [string]$ScheduleName,
#------------------------------------------------
# Dynamic Parameters
#------------------------------------------------
    [string]$JobName, 
    [string]$Prompt,
    [string]$Response
)
 
$easePROMPTJobName = "PROMPT"
$frequency = "OnRequest"
$instancePropertyName = "JOBNAME"
$instancePropertyName2 = "PROMPT"
$instancePropertyName3 = "RESPONSE"
$reason = "EASE Agent"

#------------------------------------------------
# Ignores Self Signed certificates across domains
#------------------------------------------------
function Ignore-SelfSignedCerts {
    add-type -TypeDefinition  @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;
        public class TrustAllCertsPolicy : ICertificatePolicy {
            public bool CheckValidationResult(
                ServicePoint srvPoint, X509Certificate certificate,
                WebRequest request, int certificateProblem) {
                return true;
            }
        }
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}
#----------------
# Fetch the token
#----------------
$tokensUri = ($EaseApiUrl + "/api/tokens")
$tokenObject = @{
    user = @{
        loginName = $EaseUser
        password = $EasePassword
    }
    tokenType = @{
        type = "User"
    }
}
try
{
    Ignore-SelfSignedCerts
    $token = Invoke-RestMethod -Method Post -Uri $tokensUri -Body (ConvertTo-Json $tokenObject) -ContentType "application/json"
}
catch
{
    Write-Host ("Unable to fetch token for user '" + $EaseUser + "'")
    Write-Host ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
    Write-Host ("StatusDescription: " + $_.Exception.Response.StatusDescription)

    if(!($_.Exception.Response.StatusCode.value__))
    { 
        Write-Host $_.Exception
        Exit 499 
    }
    else
    { 
        Exit $_.Exception.Response.StatusCode.value__ 
    }
}
Write-Host ("Token received with id: " + $token.id)

#------------------------------------------------------
# Create authentication header for subsequent API calls
#------------------------------------------------------
$authHeader = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$authHeader.Add("Authorization", ("Token " + $token.id))

#-----------------------------
# Get schedule id and instance
#-----------------------------
Write-Host ("Fetching schedule id from date and name provided.")

$scheduleDate = Get-Date -Format "yyyy-MM-dd"

$dailySchedulesUri = ($EaseApiUrl + "/api/dailySchedules?dates=" + $scheduleDate + "&name=" + $ScheduleName)
try
{
    $dailySchedules = Invoke-RestMethod -Method Get -Uri $dailySchedulesUri -Headers $authHeader
}
catch
{
    Write-Host ("Unable to fetch daily schedule. URI: " + $dailySchedulesUri)
    Write-Host ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
    Write-Host ("StatusDescription: " + $_.Exception.Response.StatusDescription)
    Exit $_.Exception.Response.StatusCode.value__
}

if ($dailySchedules.Count -eq 0)
{
    Write-Host ("No schedules built for date: " + $scheduleDate + ". URI: " + $dailySchedulesUri)
    Exit 600
}
elseIf($dailySchedules.status.category -eq "HELD")
{
    Write-Host "$scheduleName schedule ON HOLD, exiting..."
    Exit 601
}
elseIf($dailySchedules.status.category -eq "WAITING")
{
    Write-Host "$scheduleName schedule WAITING, exiting..."
    Exit 602
}
elseIf($dailySchedules.status.category -eq "FINISHED OK")
{
    Write-Host "$scheduleName schedule COMPLETE, exiting..."
    Exit 603
}
$scheduleId = $dailySchedules[0].masterId
$scheduleInstance = $dailySchedules[0].instance
$scheduleDate = ((New-TimeSpan -Start '1899-12-30' -End (Get-Date).Date).Days).ToString()
$id = ($scheduleDate + "|" + $scheduleId + "|" + $scheduleInstance)
Write-Host ("Fetched schedule id from name: " + $id)

#------------
# Add the job
#------------
Write-Host ("Attempting to add the job '" + $easePROMPTJobName + "' to schedule '" + $id + "'.")

$addJobUri = ($EaseApiUrl + "/api/scheduleActions")
$addJobJson = "{`"action`": `"addJobs`", `"scheduleActionItems`": [{`"id`": `"" + $id + "`", `"jobs`": [{`"id`": `"" + $easePROMPTJobName + "`", `"instanceProperties`": [{ `"name`": `"" + $instancePropertyName + "`", `"value`": `"" + $JobName + "`" }],`"instanceProperties`": [{ `"name`": `"" + $instancePropertyName2 + "`", `"value`": `"" + $Prompt + "`" }],`"instanceProperties`": [{ `"name`": `"" + $instancePropertyName3 + "`", `"value`": `"" + $Response + "`" }], `"frequency`": `"" + $Frequency + "`"}]}], `"reason`": `"" + $Reason + "`"}"
try
{
    $scheduleAction = Invoke-RestMethod -Method Post -Uri $addJobUri -Headers $authHeader -Body $addJobJson -ContentType "application/json" -ErrorVariable RespErr
}
catch
{
    Write-Host ("Unable to post request to add job: " + $addJobJson)
    Write-Host ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
    Write-Host ("StatusDescription: " + $_.Exception.Response.StatusDescription)
    Write-Host ("Content: " + $RespErr)
    Exit $_.Exception.Response.StatusCode.value__
}
Write-Host ("Add job request posted with id: " + $scheduleAction.id)

#--------------------------------------------------------
# Loop until we fetch the final result of the add request
#--------------------------------------------------------
$addJobResultUri = ($EaseApiUrl + "/api/scheduleActions/" + $scheduleAction.id)
$timeOut = 0
While ($scheduleAction.result -eq "submitted")
{
    try
    {
        $scheduleAction = Invoke-RestMethod -Method Get -Uri $addJobResultUri -Headers $authHeader
    }
    catch
    {
        Write-Host ("Unable to fetch status of job add request. URI: " + $addJobResultUri)
        Write-Host ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
        Write-Host ("StatusDescription: " + $_.Exception.Response.StatusDescription)
        Exit $_.Exception.Response.StatusCode.value__
    }

    if($scheduleAction.result -eq "submitted" -and $timeOut -lt 20)
    {   
        Start-Sleep -Seconds 3
        $timeOut++
    }
    elseif($timeOut -ge 20)
    {
        Write-Host "Timeout exceeded, check OpCon permissions and configuration"
        Exit 999
    }
}
if ($scheduleAction.result -eq "failed")
{
    Write-Host "Failed to add job to schedule:"
    Write-Host $scheduleAction.scheduleActionItems[0].jobs[0].message
    Exit 650
}

Write-Host "Returned Schedule Action:"
Write-Host ConvertTo-Json $scheduleAction.scheduleActionItems[0].jobs[0]
Write-Host ""

#-------------------------------------------------------------------------
# Successfully added job. Now fetch the final result of the job execution.
#-------------------------------------------------------------------------
Write-Host ("Successfully added job '" + $easePROMPTJobName + "' to schedule '" + $id + "'.")
Write-Host ("Waiting until the job finishes running...")

$dailyJobsUri = ($EaseApiUrl + "/api/dailyJobs?ids=" + $scheduleAction.scheduleActionItems[0].jobs[0].id)
Do
{
    try
    {
        Start-Sleep -Seconds 5
        $dailyJobs = Invoke-RestMethod -Method Get -Uri $dailyJobsUri -Headers $authHeader
        if ($dailyJobs.Count -eq 0)
        {
            Write-Host ("Added job is not found. URI: " + $dailyJobsUri)
            Exit 651
        }
        $dailyJob = $dailyJobs[0]
    }
    catch
    {
        Write-Host ("Unable to fetch status of job execution. URI: " + $dailyJobsUri)
        Write-Host ("StatusCode: " + $_.Exception.Response.StatusCode.value__)
        Write-Host ("StatusDescription: " + $_.Exception.Response.StatusDescription)
        Exit $_.Exception.Response.StatusCode.value__
    }
}
While (($dailyJob.status.id -ne 900) -and ($dailyJob.status.id -ne 910))

#-------
# Failed
#-------
if ($dailyJob.status.id -eq 910)
{
    Write-Host ("Job '" + $easePROMPTJobName + "' failed: " + $dailyJob.terminationDescription)
	$errorCodeDefinition = $dailyJob.terminationDescription
	$errorCodeDefinition = $errorCodeDefinition.SubString(1,9)
    Exit $errorCodeDefinition
}

#------------
# Finished OK
#------------
Write-Host ("Job '" + $dailyJob.id + "' finished OK: " + $dailyJob.terminationDescription)
Exit 0 