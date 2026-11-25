# Microsoft 365 offboarding

A PowerShell script that offboards a user in an order that can't lose their mailbox.

Personal project. I built it from a problem I've handled as an administrator; it is not something I deployed at my employer. All sample data is fake.

## The problem

I once cleaned up an offboarding where the license had been removed from the account before the mailbox was converted to shared. In Exchange Online, the license is what keeps a user mailbox alive. Take it away and the mailbox is disconnected, then deleted when the retention window runs out. Getting it back depends on soft-delete timing and whether a backup exists, and neither is something to count on.

The fix is an order of operations: convert the mailbox to shared first, confirm it really is shared, and only then touch the license. A checklist can say that. A script can enforce it.

## What it does

1. Blocks sign-in
2. Revokes sign-in sessions, so tokens already issued stop working
3. Converts the mailbox to shared, then asks Exchange Online to confirm the new type
4. Hides the mailbox from the global address list
5. Removes group memberships
6. Removes licenses

If step 3 can't be confirmed, the script stops. Groups are held back as well as licenses, because a license assigned through a group disappears when the membership does.

Step 6 also refuses when the shared mailbox still needs a license: 50 GB or larger, an active archive, or a litigation hold.

Every action goes to a CSV log. `-WhatIf` shows the plan without changing anything.

## Sample output

```
[Info] Lookup: Found Adele Vance (00000000-0000-0000-0000-000000000001)
[Done] Block sign-in: accountEnabled set to false
[Done] Revoke sessions: Refresh tokens invalidated
[Done] Convert mailbox: Verified as SharedMailbox
[Done] Hide from GAL: HiddenFromAddressListsEnabled set to true
[Done] Remove membership: Sales Team (Graph)
[Done] Remove membership: All Staff (Exchange)
[Skipped] Remove membership: Global Reader is a directory role or admin unit. Remove by hand.
[Done] Remove licenses: 1 license(s) removed. Shared mailbox under 50 GB with no archive or hold.
```

A full sample log is in [`sample-output/offboarding-sample.csv`](sample-output/offboarding-sample.csv).

## Requirements

- PowerShell 5.1 or 7
- `Install-Module Microsoft.Graph -Scope CurrentUser`
- `Install-Module ExchangeOnlineManagement -Scope CurrentUser`

## Permissions (least privilege)

| Where | Permission | Used for |
|---|---|---|
| Graph | `User.ReadWrite.All` | Block sign-in, remove licenses |
| Graph | `User.RevokeSessions.All` | Revoke sessions |
| Graph | `GroupMember.ReadWrite.All` | Read and remove memberships |
| Exchange Online | Recipient Management role group | `Set-Mailbox`, `Remove-DistributionGroupMember` |

No credentials are stored. Both modules prompt for an interactive sign-in. For unattended use, switch to certificate-based app authentication.

## Usage

```powershell
# See the plan, change nothing
./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com -WhatIf

# Run it, confirming each step
./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com

# Run it without prompts
./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com -Confirm:$false

# Keep the licenses for now
./Invoke-Offboarding.ps1 -UserPrincipalName adelev@contoso.com -KeepLicenses

Get-Help ./Invoke-Offboarding.ps1 -Full
```

## Tests

```powershell
Invoke-Pester ./tests
```

## Limits

- Accounts synced from on-premises Active Directory are refused. Disable those on-premises.
- Directory roles are reported, not removed. Taking away admin rights should be a deliberate act.
- OneDrive files, mail forwarding and delegate access are outside this script.
