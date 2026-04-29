#############################################################################
# Script: Manage-KeyVaultRoles.ps1
# Description: Add or remove role assignments from Azure Key Vault based on CSV
#              with automatic lock management
# Author: Magomedbashir Kushtov
# Version: 1.0
#
# Copyright (c) 2026 Magomedbashir Kushtov
#
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#############################################################################

# Write-Host is used intentionally throughout this script for colored interactive console output.
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '')]
param()

# Enable strict error handling
$ErrorActionPreference = "Stop"

# Color codes for output
$Colors = @{
    Success = "Green"
    Warning = "Yellow"
    Error   = "Red"
    Info    = "Cyan"
}

#############################################################################
# FUNCTIONS
#############################################################################

function Write-ColorOutput {
    param(
        [string]$Message,
        [ValidateSet("Success", "Warning", "Error", "Info")]
        [string]$Type = "Info"
    )
    Write-Host $Message -ForegroundColor $Colors[$Type]
}

function Test-LockManagementSupport {
    <#
    .SYNOPSIS
    Checks whether lock management cmdlets are available
    #>
    $requiredCmdlets = @(
        "Get-AzResourceLock",
        "Remove-AzResourceLock",
        "New-AzResourceLock"
    )

    $missingCmdlets = @(
        $requiredCmdlets | Where-Object {
            -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        }
    )

    return @{
        IsAvailable    = ($missingCmdlets.Count -eq 0)
        MissingCmdlets = $missingCmdlets
    }
}

function Get-LockLevelValue {
    <#
    .SYNOPSIS
    Returns lock level from different Az lock object shapes
    #>
    param(
        [object]$Lock
    )

    if ($null -eq $Lock) {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($Lock.LockLevel)) {
        return $Lock.LockLevel
    }

    if (-not [string]::IsNullOrWhiteSpace($Lock.Level)) {
        return $Lock.Level
    }

    if ($Lock.Properties -and -not [string]::IsNullOrWhiteSpace($Lock.Properties.Level)) {
        return $Lock.Properties.Level
    }

    return $null
}

function Get-LockTypeDisplay {
    <#
    .SYNOPSIS
    Converts Azure lock level values to Azure Portal lock type labels
    #>
    param(
        [string]$LockLevel
    )

    switch ($LockLevel) {
        "CanNotDelete" { return "Delete" }
        "ReadOnly" { return "Read-only" }
        default { return $LockLevel }
    }
}

function Test-ScriptPrerequisite {
    <#
    .SYNOPSIS
    Validates required Azure PowerShell modules are installed
    Checks if modules exist on disk, not just in current session
    #>
    $requiredModules = @("Az.Accounts", "Az.KeyVault", "Az.Resources")

    $missingModules = @(
        $requiredModules | Where-Object {
            -not (Get-Module -Name $_ -ListAvailable -ErrorAction SilentlyContinue)
        }
    )

    return @{
        MissingModules     = $missingModules
        AllModulesPresent  = ($missingModules.Count -eq 0)
    }
}

function Test-RequiredCmdlet {
    <#
    .SYNOPSIS
    Validates required Azure cmdlets are available after modules are imported
    #>
    $requiredCmdlets = @(
        "Get-AzContext",
        "Set-AzContext",
        "Get-AzKeyVault",
        "Get-AzADUser",
        "Get-AzRoleAssignment",
        "New-AzRoleAssignment",
        "Remove-AzRoleAssignment",
        "Get-AzResourceLock",
        "Remove-AzResourceLock",
        "New-AzResourceLock"
    )

    $moduleInstallMap = @{
        "Get-AzContext"           = "Az.Accounts"
        "Set-AzContext"           = "Az.Accounts"
        "Get-AzKeyVault"          = "Az.KeyVault"
        "Get-AzADUser"            = "Az.Resources"
        "Get-AzRoleAssignment"    = "Az.Resources"
        "New-AzRoleAssignment"    = "Az.Resources"
        "Remove-AzRoleAssignment" = "Az.Resources"
        "Get-AzResourceLock"      = "Az.Resources"
        "Remove-AzResourceLock"   = "Az.Resources"
        "New-AzResourceLock"      = "Az.Resources"
    }

    $missingCmdlets = @(
        $requiredCmdlets | Where-Object {
            -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        }
    )

    $missingModules = @(
        $missingCmdlets |
        ForEach-Object { $moduleInstallMap[$_] } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )

    return @{
        MissingCmdlets    = $missingCmdlets
        MissingModules    = $missingModules
        AllCmdletsPresent = ($missingCmdlets.Count -eq 0)
    }
}

function Read-YesNoResponse {
    param(
        [string]$Prompt
    )

    do {
        $response = Read-Host $Prompt
        if ($response.ToLower() -in @("yes", "y")) {
            return $true
        }

        if ($response.ToLower() -in @("no", "n")) {
            return $false
        }

        Write-ColorOutput "Please answer yes or no." "Error"
    } while ($true)
}

function Install-MissingModule {
    <#
    .SYNOPSIS
    Offers to install missing Azure PowerShell modules
    #>
    param(
        [string[]]$MissingModules
    )

    Write-ColorOutput "`n⚠ Missing required modules: $($MissingModules -join ', ')" "Warning"
    Write-Host "`nThese modules are required to run this script:" -ForegroundColor Cyan
    $MissingModules | ForEach-Object {
        Write-Host "  - $_" -ForegroundColor Cyan
    }

    $installChoice = Read-YesNoResponse -Prompt "`nWould you like to install these modules now? (yes/no)"

    if (-not $installChoice) {
        Write-ColorOutput "`nScript execution cancelled. Please install the required modules manually." "Warning"
        Write-Host "To install manually, run:" -ForegroundColor Cyan
        Write-Host "Install-Module -Name $($MissingModules -join ',') -Scope CurrentUser" -ForegroundColor White
        return $false
    }

    Write-ColorOutput "`nStarting module installation..." "Info"
    Write-Host "PowerShell may ask for repository trust. Answer 'Y' (Yes) to continue." -ForegroundColor Yellow

    try {
        foreach ($module in $MissingModules) {
            Write-Host "`nInstalling module: $module" -ForegroundColor Cyan
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-ColorOutput "✓ Successfully installed: $module" "Success"
        }

        Write-ColorOutput "`n✓ All required modules installed successfully!" "Success"
        Write-Host "`nModules are ready to use. Script will now import them..." -ForegroundColor Cyan
        return $true
    }
    catch {
        Write-ColorOutput "`n✗ Error during module installation: $_" "Error"
        Write-Host "`nPlease try installing modules manually:" -ForegroundColor Yellow
        Write-Host "Install-Module -Name $($MissingModules -join ',') -Scope CurrentUser" -ForegroundColor White
        return $false
    }
}

function Get-UserInput {
    <#
    .SYNOPSIS
    Prompts user for required Azure parameters
    #>
    Write-Host "`n========== Azure Key Vault Configuration ==========" -ForegroundColor Magenta

    # Get Subscription ID
    do {
        $subscriptionId = Read-Host "Enter your Azure Subscription ID"
        if (-not [guid]::TryParse($subscriptionId, [ref][guid]::Empty)) {
            Write-ColorOutput "Invalid GUID format for Subscription ID. Please try again." "Error"
            $subscriptionId = $null
        }
    } while (-not $subscriptionId)

    # Get Resource Group Name
    $resourceGroupName = Read-Host "Enter the Resource Group name"
    while ([string]::IsNullOrWhiteSpace($resourceGroupName)) {
        Write-ColorOutput "Resource Group name cannot be empty." "Error"
        $resourceGroupName = Read-Host "Enter the Resource Group name"
    }

    # Get Key Vault Name
    $keyVaultName = Read-Host "Enter the Key Vault name"
    while ([string]::IsNullOrWhiteSpace($keyVaultName)) {
        Write-ColorOutput "Key Vault name cannot be empty." "Error"
        $keyVaultName = Read-Host "Enter the Key Vault name"
    }

    # Get Operation Type
    Write-Host "`nSelect operation:" -ForegroundColor Cyan
    Write-Host "  1 - Add roles" -ForegroundColor Cyan
    Write-Host "  2 - Remove roles" -ForegroundColor Cyan
    $operation = $null
    do {
        $operationChoice = Read-Host "Enter your choice (1 or 2)"
        if ($operationChoice -eq "1") {
            $operation = "Add"
        }
        elseif ($operationChoice -eq "2") {
            $operation = "Remove"
        }
        else {
            Write-ColorOutput "Invalid selection. Please enter 1 or 2." "Error"
        }
    } while (-not $operation)

    $previewOnly = Read-YesNoResponse -Prompt "Run in preview mode only (no changes will be made)? (yes/no)"
    if ($previewOnly) {
        Write-ColorOutput "Preview mode enabled. The script will report what would change without making updates." "Warning"
    }

    return @{
        SubscriptionId      = $subscriptionId
        ResourceGroupName   = $resourceGroupName
        KeyVaultName        = $keyVaultName
        Operation           = $operation
        PreviewOnly         = $previewOnly
    }
}

function Wait-ForCSVFile {
    <#
    .SYNOPSIS
    Checks if users_and_roles.csv exists in the script directory
    If not found, prompts user to place it and waits for confirmation
    #>
    $scriptDirectory = Split-Path -Parent $MyInvocation.PSCommandPath
    if ([string]::IsNullOrEmpty($scriptDirectory)) {
        $scriptDirectory = Get-Location
    }
    $csvFilePath = Join-Path -Path $scriptDirectory -ChildPath "users_and_roles.csv"

    Write-Host "`n========== CSV File Setup ==========" -ForegroundColor Magenta

    # Check if file exists
    while (-not (Test-Path -Path $csvFilePath)) {
        Write-ColorOutput "`n⚠ Required file not found: users_and_roles.csv" "Warning"
        Write-Host "`nTo proceed, you need to place the file in the script directory:" -ForegroundColor Yellow
        Write-Host "  Location: $scriptDirectory" -ForegroundColor White
        Write-Host "  Filename: users_and_roles.csv" -ForegroundColor White

        Write-Host "`nThe CSV file must contain two columns:" -ForegroundColor Cyan
        Write-Host "  Column 1: UserID (email or Object ID)" -ForegroundColor Cyan
        Write-Host "  Column 2: Role (Key Vault role name)" -ForegroundColor Cyan

        Write-Host "`nExample CSV format:" -ForegroundColor Cyan
        Write-Host "  UserID,Role" -ForegroundColor White
        Write-Host "  user1@contoso.com,Key Vault Administrator" -ForegroundColor White
        Write-Host "  user2@contoso.com,Key Vault Secrets Officer" -ForegroundColor White

        $ready = Read-YesNoResponse -Prompt "`nHave you placed users_and_roles.csv in the script folder? (yes/no)"

        if (-not $ready) {
            $cancel = Read-YesNoResponse -Prompt "Do you want to cancel script execution? (yes/no)"
            if ($cancel) {
                Write-ColorOutput "`nScript execution cancelled." "Warning"
                exit 0
            }
        }
        else {
            if (-not (Test-Path -Path $csvFilePath)) {
                Write-ColorOutput "File still not found. Please check the location and try again." "Error"
            }
        }
    }

    Write-ColorOutput "✓ CSV file found: $csvFilePath" "Success"
    return $csvFilePath
}

function Connect-ToAzure {
    <#
    .SYNOPSIS
    Connects to Azure subscription
    #>
    param([string]$SubscriptionId)

    Write-ColorOutput "Connecting to Azure subscription..." "Info"
    try {
        $context = Get-AzContext
        if ($context.Subscription.Id -ne $SubscriptionId) {
            Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
        }
        Write-ColorOutput "Successfully connected to subscription: $SubscriptionId" "Success"
    }
    catch {
        Write-ColorOutput "Failed to connect to Azure. Error: $_" "Error"
        exit 1
    }
}

function Test-KeyVault {
    <#
    .SYNOPSIS
    Validates that the Key Vault exists and is accessible
    #>
    param(
        [string]$ResourceGroupName,
        [string]$KeyVaultName
    )

    Write-ColorOutput "Validating Key Vault..." "Info"
    try {
        $kv = Get-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $KeyVaultName -ErrorAction SilentlyContinue
        if ($null -eq $kv) {
            throw "Key Vault '$KeyVaultName' not found in resource group '$ResourceGroupName'"
        }
        Write-ColorOutput "Key Vault validated successfully." "Success"
        return $kv
    }
    catch {
        Write-ColorOutput "Error validating Key Vault: $_" "Error"
        exit 1
    }
}

function Get-KeyVaultLock {
    <#
    .SYNOPSIS
    Retrieves all management locks on the Key Vault
    #>
    param(
        [string]$ResourceGroupName,
        [string]$KeyVaultName
    )

    Write-ColorOutput "`nChecking for management locks..." "Info"
    try {
        $lockSupport = Test-LockManagementSupport
        if (-not $lockSupport.IsAvailable) {
            throw "Lock management cmdlets are not installed. Install module Az.Resources to continue."
        }

        # Locks directly on the Key Vault resource.
        $directLocks = @(
            Get-AzResourceLock -ResourceGroupName $ResourceGroupName -ResourceName $KeyVaultName -ResourceType "Microsoft.KeyVault/vaults" -ErrorAction SilentlyContinue
        )

        # Locks inherited from parent scopes.
        $subscriptionId = (Get-AzContext).Subscription.Id
        $resourceGroupScope = "/subscriptions/$subscriptionId/resourceGroups/$ResourceGroupName"
        $subscriptionScope = "/subscriptions/$subscriptionId"

        $inheritedLocks = @(
            (Get-AzResourceLock -Scope $resourceGroupScope -AtScope -ErrorAction SilentlyContinue),
            (Get-AzResourceLock -Scope $subscriptionScope -AtScope -ErrorAction SilentlyContinue)
        ) | Where-Object { $_ }

        if ($inheritedLocks.Count -gt 0) {
            Write-ColorOutput "Found inherited lock(s) at parent scope that affect this Key Vault:" "Warning"
            $inheritedLocks | ForEach-Object {
                $lockLevel = Get-LockLevelValue -Lock $_
                $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
                Write-Host "  - Lock Name: $($_.Name), Lock Type: $lockTypeDisplay, Level: $lockLevel, Scope: $($_.ResourceId)" -ForegroundColor Yellow
            }

            throw "Inherited lock(s) detected at Resource Group or Subscription scope. Remove them manually or run this script with permission to manage parent-scope locks."
        }

        if ($directLocks.Count -gt 0) {
            Write-ColorOutput "Found $($directLocks.Count) lock(s) on the Key Vault:" "Warning"
            $directLocks | ForEach-Object {
                $lockLevel = Get-LockLevelValue -Lock $_
                $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
                Write-Host "  - Lock Name: $($_.Name), Lock Type: $lockTypeDisplay, Level: $lockLevel"
            }
            return $directLocks
        }
        else {
            Write-ColorOutput "No locks found on the Key Vault." "Success"
            return $null
        }
    }
    catch {
        throw "Error checking locks: $_"
    }
}

function Remove-KeyVaultLock {
    <#
    .SYNOPSIS
    Removes management locks from the Key Vault
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param([array]$Locks)

    if ($null -eq $Locks -or $Locks.Count -eq 0) {
        return $true
    }

    $lockSupport = Test-LockManagementSupport
    if (-not $lockSupport.IsAvailable) {
        Write-ColorOutput "Lock removal cmdlets are not installed. Cannot remove locks." "Error"
        return $false
    }

    Write-ColorOutput "Removing locks..." "Info"
    try {
        foreach ($lock in $Locks) {
            Remove-AzResourceLock -LockId $lock.LockId -Force | Out-Null
            Write-ColorOutput "  - Removed lock: $($lock.Name)" "Success"
        }
        return $true
    }
    catch {
        Write-ColorOutput "Failed to remove locks: $_" "Error"
        return $false
    }
}

function Wait-ForKeyVaultUnlock {
    <#
    .SYNOPSIS
    Waits until direct Key Vault locks are no longer present
    #>
    param(
        [string]$ResourceGroupName,
        [string]$KeyVaultName,
        [int]$MaxAttempts = 10,
        [int]$DelaySeconds = 3
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $remainingDirectLocks = @(
            Get-AzResourceLock -ResourceGroupName $ResourceGroupName -ResourceName $KeyVaultName -ResourceType "Microsoft.KeyVault/vaults" -ErrorAction SilentlyContinue
        )

        if ($remainingDirectLocks.Count -eq 0) {
            return $true
        }

        Write-ColorOutput "Waiting for lock removal to propagate (attempt $attempt/$MaxAttempts)..." "Warning"
        [System.Threading.Thread]::Sleep($DelaySeconds * 1000)
    }

    return $false
}

function Get-EffectiveLockSummary {
    <#
    .SYNOPSIS
    Returns a readable summary of direct/inherited locks affecting the Key Vault
    #>
    param(
        [string]$ResourceGroupName,
        [string]$KeyVaultName
    )

    $summary = @()
    $subscriptionId = (Get-AzContext).Subscription.Id

    $directLocks = @(
        Get-AzResourceLock -ResourceGroupName $ResourceGroupName -ResourceName $KeyVaultName -ResourceType "Microsoft.KeyVault/vaults" -ErrorAction SilentlyContinue
    )

    foreach ($lock in $directLocks) {
        $lockLevel = Get-LockLevelValue -Lock $lock
        $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
        $summary += "Direct Key Vault lock: $($lock.Name) [$lockTypeDisplay/$lockLevel]"
    }

    $resourceGroupScope = "/subscriptions/$subscriptionId/resourceGroups/$ResourceGroupName"
    $subscriptionScope = "/subscriptions/$subscriptionId"

    $rgLocks = @(
        Get-AzResourceLock -Scope $resourceGroupScope -AtScope -ErrorAction SilentlyContinue
    )

    foreach ($lock in $rgLocks) {
        $lockLevel = Get-LockLevelValue -Lock $lock
        $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
        $summary += "Resource Group lock: $($lock.Name) [$lockTypeDisplay/$lockLevel]"
    }

    $subLocks = @(
        Get-AzResourceLock -Scope $subscriptionScope -AtScope -ErrorAction SilentlyContinue
    )

    foreach ($lock in $subLocks) {
        $lockLevel = Get-LockLevelValue -Lock $lock
        $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
        $summary += "Subscription lock: $($lock.Name) [$lockTypeDisplay/$lockLevel]"
    }

    if ($summary.Count -eq 0) {
        return "No direct or parent-scope locks were detected during retry checks."
    }

    return ($summary -join "; ")
}

function Restore-KeyVaultLock {
    <#
    .SYNOPSIS
    Restores management locks on the Key Vault
    #>
    param(
        [string]$ResourceGroupName,
        [string]$KeyVaultName,
        [array]$Locks
    )

    if ($null -eq $Locks -or $Locks.Count -eq 0) {
        return
    }

    $lockSupport = Test-LockManagementSupport
    if (-not $lockSupport.IsAvailable) {
        Write-ColorOutput "Lock creation cmdlets are not installed. Cannot restore locks." "Error"
        return
    }

    Write-ColorOutput "Restoring locks..." "Info"
    try {
        foreach ($lock in $Locks) {
            $lockLevel = Get-LockLevelValue -Lock $lock
            if ([string]::IsNullOrWhiteSpace($lockLevel)) {
                throw "Could not determine lock level for lock '$($lock.Name)'."
            }

            # Recreate lock directly on the Key Vault resource.
            New-AzResourceLock -LockName $lock.Name -LockLevel $lockLevel -ResourceGroupName $ResourceGroupName -ResourceName $KeyVaultName -ResourceType "Microsoft.KeyVault/vaults" -Force | Out-Null
            Write-ColorOutput "  - Restored lock: $($lock.Name)" "Success"
        }
    }
    catch {
        Write-ColorOutput "Failed to restore locks: $_" "Error"
    }
}

function Import-CSVUser {
    <#
    .SYNOPSIS
    Imports and validates CSV file with UserID and Role columns
    #>
    param([string]$CsvPath)

    Write-ColorOutput "Importing CSV file..." "Info"
    try {
        $users = Import-Csv -Path $CsvPath -ErrorAction Stop

        # Validate columns
        if ($users.Count -eq 0) {
            throw "CSV file is empty"
        }

        $firstObject = $users | Select-Object -First 1
        if (-not $firstObject.PSObject.Properties.Name.Contains("UserID")) {
            throw "CSV must contain 'UserID' column"
        }
        if (-not $firstObject.PSObject.Properties.Name.Contains("Role")) {
            throw "CSV must contain 'Role' column"
        }

        Write-ColorOutput "Successfully imported $($users.Count) user(s)" "Success"
        return $users
    }
    catch {
        Write-ColorOutput "Error importing CSV: $_" "Error"
        exit 1
    }
}

function Invoke-RoleAssignment {
    <#
    .SYNOPSIS
    Adds or removes role assignments for users
    #>
    param(
        [string]$KeyVaultName,
        [string]$ResourceGroupName,
        [string]$UserID,
        [string]$Role,
        [string]$Operation,
        [bool]$PreviewOnly = $false
    )

    try {
        $kv = Get-AzKeyVault -ResourceGroupName $ResourceGroupName -VaultName $KeyVaultName

        # Convert UserID to ObjectID if it's an email or UPN
        $objectId = $null

        # Try as object ID first
        if ([guid]::TryParse($UserID, [ref][guid]::Empty)) {
            $objectId = $UserID
        }

        # If not an object ID, try to resolve from Azure AD
        if (-not $objectId) {
            try {
                $adUser = Get-AzADUser -UserPrincipalName $UserID -ErrorAction SilentlyContinue
                if ($adUser) {
                    $objectId = $adUser.Id
                }
                else {
                    $adUser = Get-AzADUser -Filter "mail eq '$UserID'" -ErrorAction SilentlyContinue
                    if ($adUser) {
                        $objectId = $adUser.Id
                    }
                    else {
                        throw "User '$UserID' not found in Azure AD"
                    }
                }
            }
            catch {
                throw "Could not resolve user '$UserID': $_"
            }
        }

        if ($Operation -eq "Add") {
            # Check if assignment already exists
            $existing = Get-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName $Role -Scope $kv.ResourceId -ErrorAction SilentlyContinue

            if ($existing) {
                return @{
                    Status  = "Skipped"
                    Message = "User already has role '$Role'"
                }
            }
            else {
                if ($PreviewOnly) {
                    return @{
                        Status  = "Preview"
                        Message = "Would add role '$Role'"
                    }
                }

                New-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName $Role -Scope $kv.ResourceId | Out-Null
                return @{
                    Status  = "Success"
                    Message = "Role '$Role' added successfully"
                }
            }
        }
        else {
            # Remove role assignment
            $kvScopeNormalized = $kv.ResourceId.TrimEnd('/')
            $allAssignments = @(
                Get-AzRoleAssignment -ObjectId $objectId -RoleDefinitionName $Role -Scope $kv.ResourceId -ErrorAction SilentlyContinue
            )

            $directAssignments = @(
                $allAssignments |
                Where-Object { $_.Scope -and $_.Scope.TrimEnd('/') -eq $kvScopeNormalized }
            )

            if ($directAssignments.Count -gt 0) {
                if ($PreviewOnly) {
                    return @{
                        Status  = "Preview"
                        Message = "Would remove role '$Role'"
                    }
                }

                foreach ($assignment in $directAssignments) {
                    $maxAttempts = 6
                    $removed = $false

                    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                        try {
                            Remove-AzRoleAssignment -ObjectId $objectId -RoleDefinitionId $assignment.RoleDefinitionId -Scope $assignment.Scope -ErrorAction Stop | Out-Null
                            $removed = $true
                            break
                        }
                        catch {
                            $removeErrorText = $_.ToString()
                            $isConflict = ($removeErrorText -match "ScopeLocked") -or ($removeErrorText -match "invalid status code 'Conflict'")

                            if ($isConflict -and $attempt -lt $maxAttempts) {
                                Write-ColorOutput "  ! Role removal conflict detected, waiting for lock propagation (attempt $attempt/$maxAttempts)..." "Warning"
                                [System.Threading.Thread]::Sleep(10000)
                                continue
                            }

                            if ($isConflict) {
                                $lockSummary = Get-EffectiveLockSummary -ResourceGroupName $ResourceGroupName -KeyVaultName $KeyVaultName
                                throw "Role removal conflict after $maxAttempts attempts. Effective lock check: $lockSummary"
                            }

                            throw
                        }
                    }

                    if (-not $removed) {
                        throw "Role assignment removal did not complete after retry attempts."
                    }
                }

                return @{
                    Status  = "Success"
                    Message = "Role '$Role' removed successfully"
                }
            }
            else {
                $anyAssignment = $allAssignments

                if ($anyAssignment.Count -gt 0) {
                    $inheritedScopes = @(
                        $anyAssignment |
                        Where-Object { $_.Scope -and $_.Scope.TrimEnd('/') -ne $kvScopeNormalized } |
                        Select-Object -ExpandProperty Scope -Unique
                    )

                    if ($inheritedScopes.Count -gt 0) {
                        $scopeDetails = @(
                            $inheritedScopes | ForEach-Object {
                                if ($_ -match "^/subscriptions/[^/]+$") {
                                    "Subscription ($_)"
                                }
                                elseif ($_ -match "/resourceGroups/[^/]+$") {
                                    "Resource Group ($_)"
                                }
                                else {
                                    "Parent scope ($_)"
                                }
                            }
                        ) -join "; "

                        return @{
                            Status  = "Skipped"
                            Message = "Role '$Role' is inherited from parent scope(s): $scopeDetails. Remove it at that parent scope instead of Key Vault scope"
                        }
                    }

                    return @{
                        Status  = "Skipped"
                        Message = "Role '$Role' is inherited from a parent scope and cannot be removed at Key Vault scope"
                    }
                }

                return @{
                    Status  = "Skipped"
                    Message = "User does not have role '$Role'"
                }
            }
        }
    }
    catch {
        $errorText = $_.ToString()

        if ($Operation -eq "Remove" -and ($errorText -match "ScopeLocked" -or $errorText -match "invalid status code 'Conflict'")) {
            return @{
                Status  = "Error"
                Message = "Role removal was blocked because the Key Vault scope is still locked. This can happen while lock removal is propagating or if an inherited lock exists. Wait 30-60 seconds and retry. Details: $errorText"
            }
        }

        return @{
            Status  = "Error"
            Message = $errorText
        }
    }
}

function Invoke-RoleAssignmentBatch {
    <#
    .SYNOPSIS
    Processes all role assignments from CSV
    #>
    param(
        [array]$Users,
        [string]$KeyVaultName,
        [string]$ResourceGroupName,
        [string]$Operation,
        [bool]$PreviewOnly = $false
    )

    Write-Host "`n========== Processing Role Assignments ==========" -ForegroundColor Magenta

    $results = @{
        Success = 0
        Skipped = 0
        Preview = 0
        Error   = 0
        Details = @()
    }

    foreach ($user in $Users) {
        $userId = $user.UserID.Trim()
        $role = $user.Role.Trim()

        Write-Host "`nProcessing: User ID: $userId, Role: $role, Operation: $Operation" -ForegroundColor Cyan

        $result = Invoke-RoleAssignment -KeyVaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -UserID $userId -Role $role -Operation $Operation -PreviewOnly $PreviewOnly

        switch ($result.Status) {
            "Success" {
                Write-ColorOutput "  ✓ $($result.Message)" "Success"
                $results.Success++
            }
            "Skipped" {
                Write-ColorOutput "  ⊘ $($result.Message)" "Warning"
                $results.Skipped++
            }
            "Preview" {
                Write-ColorOutput "  i $($result.Message)" "Info"
                $results.Preview++
            }
            "Error" {
                Write-ColorOutput "  ✗ $($result.Message)" "Error"
                $results.Error++
            }
        }

        $results.Details += @{
            UserID  = $userId
            Role    = $role
            Status  = $result.Status
            Message = $result.Message
        }
    }

    return $results
}

function Show-Summary {
    <#
    .SYNOPSIS
    Displays operation summary
    #>
    param([hashtable]$Results)

    Write-Host "`n========== Operation Summary ==========" -ForegroundColor Magenta
    Write-ColorOutput "Successful: $($Results.Success)" "Success"
    Write-ColorOutput "Skipped: $($Results.Skipped)" "Warning"
    if ($null -ne $Results.Preview) {
        Write-ColorOutput "Preview actions: $($Results.Preview)" "Info"
    }
    Write-ColorOutput "Errors: $($Results.Error)" "Error"

    if ($Results.Error -gt 0) {
        Write-Host "`nFailed Operations:" -ForegroundColor Red
        $Results.Details | Where-Object { $_.Status -eq "Error" } | ForEach-Object {
            Write-Host "  - User: $($_.UserID), Role: $($_.Role) - $($_.Message)" -ForegroundColor Red
        }
    }
}

#############################################################################
# MAIN EXECUTION
#############################################################################

function Main {
    try {
        # Welcome message
        Write-Host "`n" -ForegroundColor Magenta
        Write-Host "╔════════════════════════════════════════════════════════╗" -ForegroundColor Magenta
        Write-Host "║   Azure Key Vault Role Management Script              ║" -ForegroundColor Magenta
        Write-Host "║   Add/Remove user roles with automatic lock mgmt      ║" -ForegroundColor Magenta
        Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Magenta

        # Check if required modules are installed
        $prerequisites = Test-ScriptPrerequisite
        if (-not $prerequisites.AllModulesPresent) {
            $missingModules = $prerequisites.MissingModules
            $installSuccess = Install-MissingModule -MissingModules $missingModules

            if (-not $installSuccess) {
                exit 1
            }
        }

        # Import modules (silent import if already loaded)
        Write-Host "`nImporting required modules..." -ForegroundColor Cyan
        try {
            Import-Module -Name Az.Accounts -ErrorAction Stop | Out-Null
            Import-Module -Name Az.KeyVault -ErrorAction Stop | Out-Null
            Import-Module -Name Az.Resources -ErrorAction Stop | Out-Null
            Write-ColorOutput "✓ All modules imported successfully!" "Success"
        }
        catch {
            Write-ColorOutput "✗ Error importing modules: $_" "Error"
            Write-Host "Please restart PowerShell and run the script again." -ForegroundColor Yellow
            exit 1
        }

        # Validate required cmdlets after import (module may be present but outdated/incomplete)
        $cmdletCheck = Test-RequiredCmdlet
        if (-not $cmdletCheck.AllCmdletsPresent) {
            Write-ColorOutput "`n⚠ Some required Azure cmdlets are missing: $($cmdletCheck.MissingCmdlets -join ', ')" "Warning"

            $installSuccess = Install-MissingModule -MissingModules $cmdletCheck.MissingModules
            if (-not $installSuccess) {
                exit 1
            }

            # Try to reload modules in current session.
            # If module versions are currently in use, PowerShell may require a session restart.
            Remove-Module -Name Az.Resources -Force -ErrorAction SilentlyContinue
            Remove-Module -Name Az.KeyVault -Force -ErrorAction SilentlyContinue
            Remove-Module -Name Az.Accounts -Force -ErrorAction SilentlyContinue

            Import-Module -Name Az.Accounts -Force -ErrorAction SilentlyContinue | Out-Null
            Import-Module -Name Az.KeyVault -Force -ErrorAction SilentlyContinue | Out-Null
            Import-Module -Name Az.Resources -Force -ErrorAction SilentlyContinue | Out-Null

            $cmdletCheck = Test-RequiredCmdlet
            if (-not $cmdletCheck.AllCmdletsPresent) {
                Write-ColorOutput "`n✗ Required Azure cmdlets are still unavailable in this PowerShell session: $($cmdletCheck.MissingCmdlets -join ', ')" "Error"
                Write-Host "A newer module version may have been installed, but the current session is still using older loaded modules." -ForegroundColor Yellow
                Write-Host "Please close this PowerShell window, open a new one, and run the script again." -ForegroundColor Yellow
                exit 1
            }
        }

        # Get user inputs
        $inputs = Get-UserInput

        # Check for CSV file in script directory
        $csvFilePath = Wait-ForCSVFile

        # Connect to Azure
        Connect-ToAzure -SubscriptionId $inputs.SubscriptionId

        # Validate Key Vault exists
        Test-KeyVault -ResourceGroupName $inputs.ResourceGroupName -KeyVaultName $inputs.KeyVaultName

        # Get existing locks
        $existingLocks = Get-KeyVaultLock -ResourceGroupName $inputs.ResourceGroupName -KeyVaultName $inputs.KeyVaultName

        if ($inputs.PreviewOnly) {
            Write-ColorOutput "Preview mode: lock removal/restoration steps are skipped." "Info"
        }

        # Handle lock removal if needed
        $locksRemoved = $false
        if ($existingLocks -and -not $inputs.PreviewOnly) {
            Write-Host "`nExisting lock details:" -ForegroundColor Yellow
            $existingLocks | ForEach-Object {
                $lockLevel = Get-LockLevelValue -Lock $_
                $lockTypeDisplay = Get-LockTypeDisplay -LockLevel $lockLevel
                Write-Host "  - Lock Name: $($_.Name), Lock Type: $lockTypeDisplay" -ForegroundColor Yellow
            }

            $savedLockDetails = Read-YesNoResponse -Prompt "Have you saved the lock name and lock type exactly as shown above? (yes/no)"
            if (-not $savedLockDetails) {
                throw "Operation cancelled. Save the lock name and lock type before continuing."
            }

            $removeLocks = Read-YesNoResponse -Prompt "Do you want to remove the existing lock(s) now so the role operation can continue? (yes/no)"
            if (-not $removeLocks) {
                throw "Operation cancelled. Existing lock(s) were not removed."
            }

            if (Remove-KeyVaultLock -Locks $existingLocks) {
                if (-not (Wait-ForKeyVaultUnlock -ResourceGroupName $inputs.ResourceGroupName -KeyVaultName $inputs.KeyVaultName)) {
                    throw "Lock removal did not fully propagate. The Key Vault still appears locked. Please wait a minute and retry."
                }

                $locksRemoved = $true
            }
            else {
                throw "Could not remove existing lock(s). Role changes were not started."
            }
        }

        # Import CSV
        $users = Import-CSVUser -CsvPath $csvFilePath

        # Process role assignments
        $results = Invoke-RoleAssignmentBatch -Users $users -KeyVaultName $inputs.KeyVaultName -ResourceGroupName $inputs.ResourceGroupName -Operation $inputs.Operation -PreviewOnly $inputs.PreviewOnly

        # Restore locks if they were removed
        if ($locksRemoved) {
            $restoreResponse = Read-YesNoResponse -Prompt "Do you want to restore the saved lock(s) exactly as they were before? (yes/no)"
            if ($restoreResponse) {
                Restore-KeyVaultLock -ResourceGroupName $inputs.ResourceGroupName -KeyVaultName $inputs.KeyVaultName -Locks $existingLocks
            }
            else {
                Write-ColorOutput "Locks were not restored. Key Vault remains unlocked." "Warning"
            }
        }

        # Show summary
        Show-Summary -Results $results

        Write-ColorOutput "`n✓ Script execution completed." "Success"
    }
    catch {
        Write-ColorOutput "`n✗ Fatal error: $_" "Error"
        exit 1
    }
}

# Execute main function
Main
