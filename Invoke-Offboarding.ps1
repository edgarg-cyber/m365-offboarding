#Requires -Version 5.1
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, ExchangeOnlineManagement

<#
.SYNOPSIS
    Offboards a Microsoft 365 user in an order that cannot lose the mailbox.

.DESCRIPTION
    Steps, in this order:
      1. Block sign-in
      2. Revoke sign-in sessions
      3. Convert the mailbox to shared, then confirm the conversion took effect
      4. Hide the mailbox from the global address list (GAL)
      5. Remove group memberships
      6. Remove licenses, only if step 3 was confirmed and the mailbox is safe without one

    If the mailbox cannot be confirmed as shared, the script stops before steps 5 and 6.
    Group removal waits too, because a license assigned through a group disappears
    when the membership does.

    Every action is written to a CSV log. Use -WhatIf to see what would happen.

.PARAMETER UserPrincipalName
    Sign-in name of the user to offboard.

.PARAMETER LogPath
    Folder for the action log. Created if missing.

.PARAMETER VerifyTimeoutSeconds
    How long to wait for Exchange Online to report the mailbox as shared. Default 120.

.PARAMETER KeepLicenses
    Skip license removal.

.PARAMETER SkipConnect
    Use Graph and Exchange Online sessions that are already open.

.EXAMPLE
    ./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com -WhatIf

.EXAMPLE
    ./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com -Confirm:$false

.NOTES
    Graph scopes: User.ReadWrite.All, User.RevokeSessions.All, GroupMember.ReadWrite.All
    Exchange Online: a role that can run Set-Mailbox and Remove-DistributionGroupMember
    (Recipient Management covers both).
    Accounts synced from on-premises Active Directory are refused: disable those on-premises.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$UserPrincipalName,

    [string]$LogPath = (Join-Path $PSScriptRoot 'logs'),

    [ValidateRange(10, 900)]
    [int]$VerifyTimeoutSeconds = 120,

    [switch]$KeepLicenses,

    [switch]$SkipConnect
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'OffboardingFunctions.ps1')

if (-not (Test-Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -WhatIf:$false | Out-Null
}
$logFile = Join-Path $LogPath ("offboarding-{0}.csv" -f (Get-Date -Format 'yyyyMMdd'))

function Write-ActionLog {
    # One row per action, so the log can be handed to whoever asks "what did you do to this account?"
    param(
        [string]$Step,
        [ValidateSet('Done', 'Skipped', 'WhatIf', 'Failed', 'Info')][string]$Result,
        [string]$Detail
    )
    $row = [pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        User      = $UserPrincipalName
        Step      = $Step
        Result    = $Result
        Detail    = $Detail
    }
    $row | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8 -WhatIf:$false
    Write-Information ("[{0}] {1}: {2}" -f $Result, $Step, $Detail) -InformationAction Continue
}

function Get-MailboxOrNull {
    # Returns the mailbox, or $null when the user has none.
    try { return Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop }
    catch { return $null }
}

try {
    # ---------- Connect ----------
    if (-not $SkipConnect) {
        Connect-MgGraph -Scopes 'User.ReadWrite.All', 'User.RevokeSessions.All', 'GroupMember.ReadWrite.All' -NoWelcome
        Connect-ExchangeOnline -ShowBanner:$false
    }

    # ---------- Find the user ----------
    $user = Get-MgUser -UserId $UserPrincipalName -Property 'id', 'displayName', 'userPrincipalName', 'accountEnabled', 'assignedLicenses', 'onPremisesSyncEnabled'
    Write-ActionLog -Step 'Lookup' -Result Info -Detail "Found $($user.DisplayName) ($($user.Id))"

    if ($user.OnPremisesSyncEnabled) {
        # Cloud-side edits to a synced account are rejected or overwritten at the next sync.
        Write-ActionLog -Step 'Lookup' -Result Failed -Detail 'Account is synced from on-premises Active Directory. Disable it there.'
        throw "$UserPrincipalName is synced from on-premises Active Directory. Offboard it there."
    }

    # ---------- 1. Block sign-in ----------
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Block sign-in')) {
        Update-MgUser -UserId $user.Id -AccountEnabled:$false
        Write-ActionLog -Step 'Block sign-in' -Result Done -Detail 'accountEnabled set to false'
    }
    else { Write-ActionLog -Step 'Block sign-in' -Result WhatIf -Detail 'Would set accountEnabled to false' }

    # ---------- 2. Revoke sessions ----------
    # Blocking sign-in stops new sign-ins. Revoking kills the refresh tokens already issued.
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Revoke sign-in sessions')) {
        Revoke-MgUserSignInSession -UserId $user.Id | Out-Null
        Write-ActionLog -Step 'Revoke sessions' -Result Done -Detail 'Refresh tokens invalidated'
    }
    else { Write-ActionLog -Step 'Revoke sessions' -Result WhatIf -Detail 'Would revoke sessions' }

    # ---------- 3. Convert mailbox to shared, then verify ----------
    $mailbox = Get-MailboxOrNull
    $hasMailbox = ($null -ne $mailbox)
    $confirmedShared = $false

    if (-not $hasMailbox) {
        Write-ActionLog -Step 'Convert mailbox' -Result Skipped -Detail 'User has no mailbox'
    }
    elseif ($mailbox.RecipientTypeDetails -eq 'SharedMailbox') {
        $confirmedShared = $true
        Write-ActionLog -Step 'Convert mailbox' -Result Skipped -Detail 'Already a shared mailbox'
    }
    elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Convert mailbox to shared')) {
        Set-Mailbox -Identity $UserPrincipalName -Type Shared

        # Do not trust the command returning without error. Ask Exchange what the mailbox is now.
        $deadline = (Get-Date).AddSeconds($VerifyTimeoutSeconds)
        while ((Get-Date) -lt $deadline -and -not $confirmedShared) {
            Start-Sleep -Seconds 10
            $mailbox = Get-MailboxOrNull
            if ($mailbox -and $mailbox.RecipientTypeDetails -eq 'SharedMailbox') { $confirmedShared = $true }
        }

        if ($confirmedShared) {
            Write-ActionLog -Step 'Convert mailbox' -Result Done -Detail 'Verified as SharedMailbox'
        }
        else {
            Write-ActionLog -Step 'Convert mailbox' -Result Failed -Detail "Not confirmed as shared after $VerifyTimeoutSeconds seconds"
        }
    }
    else { Write-ActionLog -Step 'Convert mailbox' -Result WhatIf -Detail 'Would convert to shared and verify' }

    # The gate. With a mailbox that is not confirmed shared, nothing below is allowed to run for real.
    if ($hasMailbox -and -not $confirmedShared -and -not $WhatIfPreference) {
        Write-ActionLog -Step 'Stop' -Result Failed -Detail 'Mailbox not confirmed shared. Groups and licenses left untouched.'
        throw 'Mailbox was not confirmed as shared. Groups and licenses were left in place. Check the mailbox and run again.'
    }

    # ---------- 4. Hide from the address list ----------
    if ($hasMailbox) {
        if ($PSCmdlet.ShouldProcess($UserPrincipalName, 'Hide from global address list')) {
            Set-Mailbox -Identity $UserPrincipalName -HiddenFromAddressListsEnabled $true
            Write-ActionLog -Step 'Hide from GAL' -Result Done -Detail 'HiddenFromAddressListsEnabled set to true'
        }
        else { Write-ActionLog -Step 'Hide from GAL' -Result WhatIf -Detail 'Would hide from address lists' }
    }

    # ---------- 5. Remove group memberships ----------
    $memberships = Get-MgUserMemberOf -UserId $user.Id -All
    foreach ($membership in $memberships) {
        $props  = $membership.AdditionalProperties
        $name   = $props['displayName']
        $method = Get-GroupRemovalMethod -Properties $props

        switch ($method) {
            'SkipNotGroup' { Write-ActionLog -Step 'Remove membership' -Result Skipped -Detail "$name is a directory role or admin unit. Remove by hand." }
            'SkipSynced'   { Write-ActionLog -Step 'Remove membership' -Result Skipped -Detail "$name is synced from on-premises. Remove there." }
            'SkipDynamic'  { Write-ActionLog -Step 'Remove membership' -Result Skipped -Detail "$name has dynamic membership." }
            default {
                if ($PSCmdlet.ShouldProcess("$UserPrincipalName in $name", 'Remove group membership')) {
                    try {
                        if ($method -eq 'Exchange') {
                            Remove-DistributionGroupMember -Identity $props['mail'] -Member $UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false
                        }
                        else {
                            # Documented Graph call: DELETE /groups/{id}/members/{id}/$ref
                            $uri = "https://graph.microsoft.com/v1.0/groups/$($membership.Id)/members/$($user.Id)/`$ref"
                            Invoke-MgGraphRequest -Method DELETE -Uri $uri | Out-Null
                        }
                        Write-ActionLog -Step 'Remove membership' -Result Done -Detail "$name ($method)"
                    }
                    catch {
                        # One stubborn group should not stop the rest of the offboarding.
                        Write-ActionLog -Step 'Remove membership' -Result Failed -Detail "$name : $($_.Exception.Message)"
                    }
                }
                else { Write-ActionLog -Step 'Remove membership' -Result WhatIf -Detail "Would remove from $name ($method)" }
            }
        }
    }

    # ---------- 6. Remove licenses ----------
    $skuIds = @($user.AssignedLicenses | ForEach-Object { $_.SkuId })

    if ($KeepLicenses) {
        Write-ActionLog -Step 'Remove licenses' -Result Skipped -Detail '-KeepLicenses was set'
    }
    elseif ($skuIds.Count -eq 0) {
        Write-ActionLog -Step 'Remove licenses' -Result Skipped -Detail 'No licenses assigned'
    }
    elseif ($WhatIfPreference) {
        Write-ActionLog -Step 'Remove licenses' -Result WhatIf -Detail "Would remove $($skuIds.Count) license(s) after the mailbox checks pass"
    }
    else {
        $bytes = $null
        $archiveActive = $false
        $hold = $false
        if ($hasMailbox) {
            $stats = Get-MailboxStatistics -Identity $UserPrincipalName
            $bytes = ConvertTo-ByteCount -SizeString ([string]$stats.TotalItemSize)
            $archiveActive = ($mailbox.ArchiveStatus -eq 'Active')
            $hold = [bool]$mailbox.LitigationHoldEnabled
        }

        $check = Test-LicenseRemovalSafe -HasMailbox $hasMailbox -RecipientTypeDetails ([string]$mailbox.RecipientTypeDetails) `
            -ArchiveActive $archiveActive -LitigationHoldEnabled $hold -MailboxBytes $bytes

        if (-not $check.Safe) {
            Write-ActionLog -Step 'Remove licenses' -Result Skipped -Detail $check.Reason
        }
        elseif ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove $($skuIds.Count) license(s)")) {
            try {
                Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $skuIds | Out-Null
                Write-ActionLog -Step 'Remove licenses' -Result Done -Detail "$($skuIds.Count) license(s) removed. $($check.Reason)"
            }
            catch {
                # Licenses inherited from a group cannot be removed directly; they go when the membership goes.
                Write-ActionLog -Step 'Remove licenses' -Result Failed -Detail $_.Exception.Message
            }
        }
    }

    Write-ActionLog -Step 'Finish' -Result Info -Detail "Log: $logFile"
}
catch {
    Write-ActionLog -Step 'Error' -Result Failed -Detail $_.Exception.Message
    throw
}
