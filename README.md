# Azure Key Vault Role Management Guide

Use this guide to run Manage-KeyVaultRoles.ps1 safely.

## Operator SOP (Quick Run)

1. Activate PIM Owner role for the Resource Group that contains the Key Vault (if your tenant uses PIM).
2. Open PowerShell and run:

```powershell
# Go to the folder where Manage-KeyVaultRoles.ps1 is located
cd "<path-to-script-folder>"
Connect-AzAccount
Set-AzContext -SubscriptionId <your-subscription-id>
```

3. Ensure users_and_roles.csv is in the same folder as the script and has columns UserID,Role.
4. Run:

```powershell
.\Manage-KeyVaultRoles.ps1
```

5. First run recommendation:
- Choose preview mode = yes
- Review output
- Run again with preview mode = no to apply changes

6. If lock prompts appear:
- Confirm lock details are saved
- Allow temporary lock removal
- Restore lock at the end

7. Keep PIM role active until the script completes and lock restore is done.

## Prerequisites

### Access and Roles
You need permission to:
- Add/remove RBAC assignments
- Remove/create Key Vault locks

Recommended role: Owner on the Resource Group containing the Key Vault.

If using PIM, activate Owner before the run. User Access Administrator alone is not enough for lock remove/restore operations.

### PowerShell Modules
Required modules:
- Az.Accounts
- Az.KeyVault
- Az.Resources

The script auto-checks and can install missing modules. Manual install command:

```powershell
Install-Module -Name Az.Accounts,Az.KeyVault,Az.Resources -Scope CurrentUser
```

## Input File

File name must be: users_and_roles.csv

Required format:

```csv
UserID,Role
user1@contoso.com,Key Vault Administrator
user2@contoso.com,Key Vault Secrets Officer
```

Valid UserID values:
- User principal name (email style)
- Azure AD object ID (GUID)

## Run Flow

The script prompts for:
- Subscription ID
- Resource Group name
- Key Vault name
- Operation (Add or Remove)
- Preview mode (yes or no)

If locks exist, the script can remove and later restore them.

## Important Behavior

- Re-running Add is safe: existing assignments are skipped.
- Remove only removes direct Key Vault scope assignments.
- If a role is inherited from Resource Group/Subscription scope, it is skipped and reported.
- If lock-related conflict appears, wait 30-60 seconds and rerun (lock propagation delay).

## Quick Troubleshooting

- User not found in Azure AD:
  - Verify UserID spelling/tenant
  - Use correct UPN or object ID

- Insufficient permissions:
  - Activate/assign Owner at Resource Group scope

- Key Vault not found:
  - Verify subscription, resource group, and vault name

- Required file not found: users_and_roles.csv:
  - Place file in script folder
  - Confirm header is exactly UserID,Role
