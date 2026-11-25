# Pure helper functions for Invoke-Offboarding.ps1.
# No calls to Microsoft Graph or Exchange Online here, so Pester can test them offline.

function ConvertTo-ByteCount {
    <#
    .SYNOPSIS
        Pulls the byte count out of an Exchange size string such as "1.2 GB (1,288,490,188 bytes)".
    .OUTPUTS
        Int64, or $null if the string has no byte count.
    #>
    param([string]$SizeString)

    if ($SizeString -match '\(([\d,]+) bytes\)') {
        return [int64]($Matches[1] -replace ',', '')
    }
    return $null
}

function Test-LicenseRemovalSafe {
    <#
    .SYNOPSIS
        Decides whether removing the user's licenses could cost them their mailbox.
    .DESCRIPTION
        A shared mailbox needs no license only while it is under 50 GB, has no
        archive, and is not on litigation hold. Anything else keeps its license.
    .OUTPUTS
        Object with Safe (bool) and Reason (string).
    #>
    param(
        [bool]$HasMailbox,
        [string]$RecipientTypeDetails,
        [bool]$ArchiveActive,
        [bool]$LitigationHoldEnabled,
        [Nullable[int64]]$MailboxBytes
    )

    $limitBytes = 50GB

    if (-not $HasMailbox) {
        return [pscustomobject]@{ Safe = $true; Reason = 'User has no mailbox.' }
    }
    if ($RecipientTypeDetails -ne 'SharedMailbox') {
        return [pscustomobject]@{ Safe = $false; Reason = "Mailbox type is '$RecipientTypeDetails', not SharedMailbox. Removing the license would delete it." }
    }
    if ($ArchiveActive) {
        return [pscustomobject]@{ Safe = $false; Reason = 'Mailbox has an archive, which needs a license.' }
    }
    if ($LitigationHoldEnabled) {
        return [pscustomobject]@{ Safe = $false; Reason = 'Mailbox is on litigation hold, which needs a license.' }
    }
    if ($null -eq $MailboxBytes) {
        return [pscustomobject]@{ Safe = $false; Reason = 'Mailbox size could not be read.' }
    }
    if ($MailboxBytes -ge $limitBytes) {
        return [pscustomobject]@{ Safe = $false; Reason = 'Mailbox is 50 GB or larger, which needs a license.' }
    }
    return [pscustomobject]@{ Safe = $true; Reason = 'Shared mailbox under 50 GB with no archive or hold.' }
}

function Get-GroupRemovalMethod {
    <#
    .SYNOPSIS
        Works out how (or whether) a membership can be removed.
    .PARAMETER Properties
        The AdditionalProperties dictionary of an object returned by Get-MgUserMemberOf.
    .OUTPUTS
        One of: Graph, Exchange, SkipDynamic, SkipSynced, SkipNotGroup.
    #>
    param([System.Collections.IDictionary]$Properties)

    if ($Properties['@odata.type'] -ne '#microsoft.graph.group') {
        # Directory roles and administrative units also come back from memberOf.
        return 'SkipNotGroup'
    }
    if ($Properties['onPremisesSyncEnabled'] -eq $true) {
        # Synced from on-premises Active Directory. Changes must be made there.
        return 'SkipSynced'
    }

    $groupTypes = @($Properties['groupTypes'])
    if ($groupTypes -contains 'DynamicMembership') {
        # Membership is calculated from a rule and cannot be edited by hand.
        return 'SkipDynamic'
    }
    if ($Properties['mailEnabled'] -eq $true -and $groupTypes -notcontains 'Unified') {
        # Distribution lists and mail-enabled security groups are managed by Exchange, not Graph.
        return 'Exchange'
    }
    return 'Graph'
}
