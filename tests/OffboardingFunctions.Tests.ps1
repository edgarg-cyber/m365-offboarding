BeforeAll {
    . "$PSScriptRoot/../OffboardingFunctions.ps1"
}

Describe 'ConvertTo-ByteCount' {
    It 'reads the byte count from an Exchange size string' {
        ConvertTo-ByteCount -SizeString '1.2 GB (1,288,490,188 bytes)' | Should -Be 1288490188
    }
    It 'returns null when there is no byte count' {
        ConvertTo-ByteCount -SizeString 'Unlimited' | Should -BeNullOrEmpty
    }
}

Describe 'Test-LicenseRemovalSafe' {
    It 'refuses when the mailbox is still a user mailbox' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'UserMailbox' -MailboxBytes 1GB
        $r.Safe | Should -BeFalse
    }
    It 'allows a small shared mailbox with no archive or hold' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'SharedMailbox' -MailboxBytes 5GB
        $r.Safe | Should -BeTrue
    }
    It 'refuses a shared mailbox of 50 GB or more' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'SharedMailbox' -MailboxBytes 60GB
        $r.Safe | Should -BeFalse
    }
    It 'refuses when an archive is active' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'SharedMailbox' -ArchiveActive $true -MailboxBytes 1GB
        $r.Safe | Should -BeFalse
    }
    It 'refuses when litigation hold is on' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'SharedMailbox' -LitigationHoldEnabled $true -MailboxBytes 1GB
        $r.Safe | Should -BeFalse
    }
    It 'refuses when the size is unknown' {
        $r = Test-LicenseRemovalSafe -HasMailbox $true -RecipientTypeDetails 'SharedMailbox' -MailboxBytes $null
        $r.Safe | Should -BeFalse
    }
    It 'allows a user with no mailbox' {
        (Test-LicenseRemovalSafe -HasMailbox $false).Safe | Should -BeTrue
    }
}

Describe 'Get-GroupRemovalMethod' {
    It 'uses Graph for a security group' {
        $p = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @(); mailEnabled = $false }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'Graph'
    }
    It 'uses Graph for a Microsoft 365 group' {
        $p = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @('Unified'); mailEnabled = $true }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'Graph'
    }
    It 'uses Exchange for a distribution list' {
        $p = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @(); mailEnabled = $true }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'Exchange'
    }
    It 'skips dynamic groups' {
        $p = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @('DynamicMembership'); mailEnabled = $false }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'SkipDynamic'
    }
    It 'skips groups synced from on-premises' {
        $p = @{ '@odata.type' = '#microsoft.graph.group'; groupTypes = @(); onPremisesSyncEnabled = $true }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'SkipSynced'
    }
    It 'skips directory roles' {
        $p = @{ '@odata.type' = '#microsoft.graph.directoryRole' }
        Get-GroupRemovalMethod -Properties $p | Should -Be 'SkipNotGroup'
    }
}
