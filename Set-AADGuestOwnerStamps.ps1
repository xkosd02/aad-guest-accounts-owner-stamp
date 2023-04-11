#how many days past yesterday should we look
$daysBackOffset = 0

#max allowed time difference between audit log event and AAD guest account creation (in seconds)
$timeDiffTolerance = 60

#AAD app authentication details  
$ClientId = 'xyz123'
$Thumbprint = '123456789'
$ClientCertificate = Get-Item "Cert:\LocalMachine\My\$($Thumbprint)"

$Yesterday 			= $(Get-Date).AddDays(-1-$daysBackOffset).ToString("yyyy-MM-dd")
$YesterdayUTCStart 	= $Yesterday + "T00:00:00Z"

#wrapper function for reading from Graph
function GetGraphOutputRM {
	param ([string]$prmURI, [string]$prmContentType, [string]$prmText)
	$URI = $prmURI
	Write-Host -NoNewline "Reading $($prmText) from Graph" -ForegroundColor Cyan
	$Headers = @{Authorization = "Bearer $accessToken"}
	Try {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Do {
            if ($prmText) { 
                Write-Host -NoNewline "." -ForegroundColor Cyan 
            }
            $Query = Invoke-RestMethod -Headers $Headers -Uri $URI -Method GET -ContentType $prmContentType -ErrorAction Stop
            if ($Query.Value) {
                    $Result += $Query.Value
                }	
            Else {
                $Result += $Query
            }
            $URI = $Query.'@odata.nextlink'
        }
        Until ($null -eq $URI)
	}
	Catch {
        Write-Host "ERR: $($URI) $($_.Exception.Message)" -ForegroundColor Red
        Exit
	}
	Write-Host "..done ($($Result.Count))" -ForegroundColor Cyan
	Return ,$result
}

#get access token from Graph (uses MSAL.PS module)
Try {
    $Token = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientCertificate $ClientCertificate -ErrorAction Stop
}
Catch {
    Write-Host "ERR: $($_.Exception.Message)" -ForegroundColor Red
    Exit
}
if ($Token) {
    $AccessToken = $Token.AccessToken
    Write-Host "MSAL access token requested - expiration $($Token.ExpiresOn)"
}

#get audit events from Graph
$Uri= "https://graph.microsoft.com/beta/auditLogs/directoryAudits?`$Filter=activityDisplayName+eq+'Invite external user'+and+result+eq+'success'+and+activityDateTime+ge+$($YesterdayUTCStart)&`$orderby=activityDateTime"
$RecentAuditLogEventsInvite = GetGraphOutputRM $URI $ContentTypeJSON "AAD guest account invite log events"

#get recently created AAD guest accounts
$Uri = "https://graph.microsoft.com/beta/users?`$filter=UserType+eq+'Guest'+and+createdDateTime+ge+$($YesterdayUTCStart)"
$RecentGuests = GetGraphOutputRM $URI $ContentTypeJSON "AAD guest accounts"

#fill the audit log event database
$AuditLogEvents_DB = @{}
foreach ($AuditLogEvent in $RecentAuditLogEventsInvite) {
    $ownerUpn = $AuditLogEvent.initiatedBy.user.userPrincipalName
    if ($ownerUpn) {
        if ($AuditLogEvent.additionalDetails.Count -ge 1) {
            $InvitedUserMail = $($AuditLogEvent.additionalDetails | Where-Object { $_.key -eq "invitedUserEmailAddress" })[0].value
        }
        if ($InvitedUserMail) {
            $AuditRecord = [pscustomobject]@{
                activityDateTime    = $AuditLogEvent.activityDateTime;
                ownerUpn            = $ownerUpn;
                invitedUserMail     = $InvitedUserMail
            }
            if ($AuditLogEvents_DB.Contains($InvitedUserMail)) {
                $diff = New-TimeSpan $AuditLogEvents_DB[$InvitedUserMail].activityDateTime $AuditLogEvent.activityDateTime
                $AuditLogEvents_DB.Set_Item($InvitedUserMail,$AuditRecord)
            }
            else {
                $AuditLogEvents_DB.Add($InvitedUserMail,$AuditRecord)
            }
        }
	}
}
Remove-Variable RecentAuditLogEventsInvite

#process AAD guest accounts
$stampedGuestsCount = 0
foreach ($Guest in $RecentGuests) {
    if (($Guest.employeeType -eq "") -or ($null -eq $Guest.employeeType)) {
        $AuditLogDBRecord = $AuditLogEvents_DB[$Guest.mail]
        if ($AuditLogDBRecord) {
            $diff = New-TimeSpan $AuditLogDBRecord.activityDateTime $Guest.createdDateTime
            if ([Math]::Abs($diff.Seconds) -le $timeDiffTolerance) {
                $stampString = $AuditLogDBRecord.ownerUpn + ";" + $AuditLogDBRecord.activityDateTime   
                $Uri= "https://graph.microsoft.com/beta/users/$($Guest.id)"
                $GraphBody = @{employeeType = $stampString} | ConvertTo-Json
                Try {
                    $ResultPATCH = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $Uri -Body $GraphBody -Method PATCH -ContentType "application/json"
                    Write-Host "$($AuditLogDBRecord.activityDateTime) $($Guest.mail) $($AuditLogDBRecord.ownerUpn) $($Guest.createdDateTime) $($diff.Seconds)" -ForegroundColor Green
                    $stampedGuestsCount ++
                }
                Catch {
                    Write-Host "ERR: $($_.ErrorDetails.Message)"
                }
            }
            else {
                Write-Host "$($AuditLogDBRecord.activityDateTime) $($Guest.mail) $($Guest.createdDateTime) $($diff.Seconds) - max allowed diff $($timeDiffTolerance) sec" -ForegroundColor Red
            }
        }
        else {
            write-host "No audit record for $($Guest.mail) found" -ForegroundColor Yellow
        }
    }
}

if ($stampedGuestsCount -gt 0) {
    Write-Host "$($stampedGuestsCount) AAD guest accounts stamped" -ForegroundColor Green
}
else {
    Write-Host "No AAD guest accounts found to stamp" -ForegroundColor Gray
}

