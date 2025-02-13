<#
    .SYNOPSIS
        Emergency script to resubmit any emails in the Phish queue from 5:44pm ET Feb 10 to 3:20am ET Feb 11.

    .DESCRIPTION
        Script written to handle recovery from Feb 10 2025 Proofpoint Incident

#>
Import-Module PSPas

$pphost = '<Admin POD URL>' 
$adminuri = '/rest/v1/quarantine'
$APIParms = @{}

# Admin API password file created 
## 'password' | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File apipw.txt
$apipw = Get-Content apipw.txt | ConvertTo-SecureString
$apiuser = '<apiusername>'

#Resubmit all emails that are in the Phish queue that arrived between 5:44pm 2/10 and 3:20am 2/11
$startDate = '2025-02-10 22:44:00' #UTCTime
$startDateDT = Get-Date "$($startdate)Z"
$endDateStepDT = $startdatedt.AddMinutes(30)
$enddatestep = Get-Date ($enddatestepDT.ToUniversalTime()) -format "yyyy-MM-dd HH:mm:ss"
$endDateDT = Get-Date '2025-02-11 08:20:00Z' #UTC Time
$rcpt = '*@MyDomain.com'
$folder = 'Phish'

#Get the emails in 30 minute chunks older to newer.  API always returns them newer to older, so first call will get emails from 23:14 to 22:44.  Second call will then
#   get 23:44 to 23:14
while ($endDateStepDT -le $endDateDT) {
    $uriparms = "startdate=$([URI]::EscapeDataString($startdate))&enddate=$([URI]::EscapeDataString($enddatestep))&rcpt=$([URI]::EscapeDataString($rcpt))&folder=$([URI]::EscapeDataString($folder))"
    $APIParms['Credential'] = New-Object -TypeName PSCredential -ArgumentList $apiuser, $apipw
    $APIParms['Uri'] = 'https://' + $pphost + $adminuri + '?' + $uriparms
    $APIParms['UserAgent'] = 'Powershell'
    $APIParms['Method'] = 'GET'
    $APIParms['ContentType'] = 'application/json'
    $APIParms['Headers'] = @{'Host' = $pphost; 'Accept' = 'application/json' }
    try {
        $response = Invoke-WebRequest @APIParms -ErrorAction Stop
        if ($response.StatusCode -eq 200) {
            $appResponse = $Response.Content | ConvertFrom-Json
            $recordCount = $appResponse.count
            $meta = $appResponse.meta
            $queryID = $meta.$queryID
            $query_params = $meta.$query_params
            $limit = $meta.limit
            $records = $appResponse.records
        }
        else {
            throw "HTTP Response $($response.StatusCode): $($response.StatusDescription)"
        }
    }
    catch {
        Throw "Error: $($Error[0])"
    }

    if ($records.count -gt 0) { Write-Host "Found $($records.count) messages Starting at $($records[0].date)" }

    $APIParms['Method'] = 'POST'
    $APIParms['Uri'] = 'https://' + $pphost + $adminuri
    
    
    foreach ($rec in $records) {
        $body = New-Object -Typename PSObject -Property @{  "localguid" = $rec.localguid; "folder" = "Phish"; "action" = "resubmit"; }
        $bodyjson = $body | ConvertTo-Json

        $response = $bodyjson | Invoke-WebRequest @APIParms
        if ($response.StatusCode -ne '202') {
            Write-Host "$($rec.date) $($rec.from) $($rec.rcpts) $($rec.subject) STATUS:\n $($Response)"
        }
    }
    
    $endDateStepDT = $endDateStepDT.AddMinutes(30)
    if ($endDateStepDT -gt $endDateDT -and -not $last) {
        #Set the flag so this is the last cycle
        $endDateStepDT = $endDateDT
        $last = $true
    }
    $enddatestep = Get-Date ($enddatestepDT.ToUniversalTime()) -format "yyyy-MM-dd HH:mm:ss"
}