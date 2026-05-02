#############################################################################
# Script: Launch-KeyVaultManager.ps1
# Description: Windows Forms GUI for Azure Key Vault role management.
#              Self-contained; runs Azure operations in a background runspace
#              so the window stays fully responsive during long-running tasks.
# Author: Magomedbashir Kushtov
# Version: 2.0
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

param()

#region ── Bootstrap ────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
#endregion

#region ── Colour palette ───────────────────────────────────────────────────
$C = @{
    Bg        = [System.Drawing.Color]::FromArgb(28,  28,  28)
    Panel     = [System.Drawing.Color]::FromArgb(40,  40,  40)
    GroupBg   = [System.Drawing.Color]::FromArgb(46,  46,  46)
    Input     = [System.Drawing.Color]::FromArgb(55,  55,  55)
    Accent    = [System.Drawing.Color]::FromArgb(0,  120, 215)
    AccentHov = [System.Drawing.Color]::FromArgb(16, 137, 227)
    Danger    = [System.Drawing.Color]::FromArgb(196,  43,  28)
    Text      = [System.Drawing.Color]::FromArgb(240, 240, 240)
    Muted     = [System.Drawing.Color]::FromArgb(150, 150, 150)
    LogBg     = [System.Drawing.Color]::FromArgb(18,  18,  18)
    Success   = [System.Drawing.Color]::FromArgb( 87, 202,  89)
    Warning   = [System.Drawing.Color]::FromArgb(234, 179,   8)
    Error     = [System.Drawing.Color]::FromArgb(239,  68,  68)
    Info      = [System.Drawing.Color]::FromArgb( 56, 189, 248)
    Purple    = [System.Drawing.Color]::FromArgb(167, 139, 250)
    White     = [System.Drawing.Color]::FromArgb(240, 240, 240)
}
#endregion

#region ── Shared state for background-runspace ↔ UI communication ──────────
$sync = [hashtable]::Synchronized(@{
    Messages       = [System.Collections.Generic.List[psobject]]::new()
    Done           = $false
    NeedsLockReply = $false   # runspace sets to ask main thread for lock confirm
    LockNames      = ''
    LockReply      = $null    # main thread sets to $true/$false
    NeedLockRestore= $false
    RestoreReply   = $null
    LockDetected   = 'Not checked yet.'
    LockCurrent    = 'Waiting to run.'
    LockEnd        = 'Pending.'
})
#endregion

#region ── Helper: append a coloured line to the log RichTextBox ────────────
function Add-LogLine {
    param(
        [System.Windows.Forms.RichTextBox]$Rtb,
        [string]$Text,
        [System.Drawing.Color]$Color
    )
    $Rtb.SelectionStart  = $Rtb.TextLength
    $Rtb.SelectionLength = 0
    $Rtb.SelectionColor  = $Color
    $Rtb.AppendText("$Text`r`n")
    $Rtb.SelectionColor  = $Rtb.ForeColor
    $Rtb.ScrollToCaret()
}
#endregion

#region ── Background work: all Azure operations ────────────────────────────
# This scriptblock runs in a separate runspace. It communicates back to the UI
# only via $sync — never touches WinForms controls directly.
$AzureWork = {
    param($sync, $subId, $rg, $kvName, $op, $preview, $csvPath)

    function Log {
        param([string]$Msg, [string]$Kind = 'Info')
        $sync.Messages.Add([pscustomobject]@{ Text = $Msg; Kind = $Kind })
    }

    function Write-LockState {
        param(
            [string]$Detected,
            [string]$Current,
            [string]$End
        )
        if ($null -ne $Detected) { $sync.LockDetected = $Detected }
        if ($null -ne $Current)  { $sync.LockCurrent  = $Current }
        if ($null -ne $End)      { $sync.LockEnd      = $End }
    }

    function Get-LockLevelValue {
        param([object]$Lock)
        if ($null -eq $Lock) { return $null }
        if (-not [string]::IsNullOrWhiteSpace($Lock.LockLevel)) { return $Lock.LockLevel }
        if (-not [string]::IsNullOrWhiteSpace($Lock.Level)) { return $Lock.Level }
        if ($Lock.Properties -and -not [string]::IsNullOrWhiteSpace($Lock.Properties.Level)) { return $Lock.Properties.Level }
        return $null
    }

    function Get-LockTypeDisplay {
        param([string]$LockLevel)
        switch ($LockLevel) {
            'CanNotDelete' { return 'Delete' }
            'ReadOnly' { return 'Read-only' }
            default { return $LockLevel }
        }
    }

    function Format-LockSummary {
        param([object[]]$Locks)
        if ($null -eq $Locks -or $Locks.Count -eq 0) {
            return 'None'
        }

        return (($Locks | ForEach-Object {
            $lvl = Get-LockLevelValue -Lock $_
            $type = Get-LockTypeDisplay -LockLevel $lvl

            if ([string]::IsNullOrWhiteSpace($lvl)) {
                "Lock Name: $($_.Name)"
            }
            elseif ([string]::IsNullOrWhiteSpace($type) -or $type -eq $lvl) {
                "Lock Name: $($_.Name), Level: $lvl"
            }
            else {
                "Lock Name: $($_.Name), Lock Type: $type, Level: $lvl"
            }
        }) -join '; ')
    }

    function Write-LogSeparator { Log '─────────────────────────────────────────' 'Purple' }

    try {
        Log ''
        Log '  Azure Key Vault Role Manager  ' 'Purple'
        Write-LogSeparator
        Log "  Operation : $op$(if ($preview) { '  [PREVIEW — no changes]' })" 'Purple'
        Write-LogSeparator
        Log ''
        Write-LockState -Detected 'Not checked yet.' -Current 'Checking lock state...' -End 'Pending.'

        # ── 1. Module check ──────────────────────────────────────────────
        Log 'Checking Az modules...' 'Info'
        $required = @('Az.Accounts', 'Az.KeyVault', 'Az.Resources')
        $missing  = $required | Where-Object { -not (Get-Module -Name $_ -ListAvailable -EA SilentlyContinue) }

        if ($missing) {
            Log "Installing missing module(s): $($missing -join ', ')" 'Warning'
            foreach ($m in $missing) {
                Log "  Installing $m ..." 'Info'
                Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -EA Stop
                Log "  ✓ Installed $m" 'Success'
            }
        }

        # ── 2. Import modules ────────────────────────────────────────────
        Import-Module Az.Accounts -EA Stop | Out-Null
        Import-Module Az.KeyVault  -EA Stop | Out-Null
        Import-Module Az.Resources -EA Stop | Out-Null
        Log '✓ Az modules ready.' 'Success'

        # ── 3. Connect ───────────────────────────────────────────────────
        Log "Connecting to subscription $subId ..." 'Info'
        $ctx = Get-AzContext -EA SilentlyContinue
        if ($null -eq $ctx -or $ctx.Subscription.Id -ne $subId) {
            Connect-AzAccount -SubscriptionId $subId -EA Stop | Out-Null
        }
        Set-AzContext -SubscriptionId $subId -EA Stop | Out-Null
        Log '✓ Connected to Azure.' 'Success'

        # ── 4. Validate Key Vault ────────────────────────────────────────
        Log "Validating Key Vault '$kvName'..." 'Info'
        $kv = Get-AzKeyVault -ResourceGroupName $rg -VaultName $kvName -EA SilentlyContinue
        if ($null -eq $kv) {
            Log "✗ Key Vault '$kvName' not found in resource group '$rg'." 'Error'
            return
        }
        Log '✓ Key Vault found.' 'Success'

        # ── 5. Lock detection ────────────────────────────────────────────
        Log 'Checking for management locks...' 'Info'

        $subIdCurrent = (Get-AzContext).Subscription.Id
        $rgScope  = "/subscriptions/$subIdCurrent/resourceGroups/$rg"
        $subScope = "/subscriptions/$subIdCurrent"

        $inheritedLocks = @(
            (Get-AzResourceLock -Scope $rgScope  -AtScope -EA SilentlyContinue),
            (Get-AzResourceLock -Scope $subScope -AtScope -EA SilentlyContinue)
        ) | Where-Object { $_ }

        if ($inheritedLocks.Count -gt 0) {
            $names = Format-LockSummary -Locks $inheritedLocks
            Write-LockState -Detected "Inherited lock(s): $names" -Current 'Blocked by inherited parent-scope lock(s).' -End 'Unchanged (operation stopped).'
            Log "✗ Inherited lock(s) at parent scope: $names" 'Error'
            Log '  Remove them manually at the Resource Group or Subscription level.' 'Warning'
            return
        }

        $directLocks = @(
            Get-AzResourceLock -ResourceGroupName $rg -ResourceName $kvName `
                -ResourceType 'Microsoft.KeyVault/vaults' -EA SilentlyContinue
        )

        $locksRemoved   = $false
        $locksToRestore = $null

        if ($directLocks.Count -gt 0 -and -not $preview) {
            $lockNames = Format-LockSummary -Locks $directLocks
            Write-LockState -Detected $lockNames -Current 'Direct lock(s) found. Awaiting user confirmation to remove.' -End $null
            Log "⚠ Direct lock(s) found: $lockNames" 'Warning'

            # Ask UI thread for confirmation
            $sync.LockNames      = $lockNames
            $sync.NeedsLockReply = $true

            $deadline = (Get-Date).AddSeconds(120)
            while ($null -eq $sync.LockReply -and (Get-Date) -lt $deadline) {
                [System.Threading.Thread]::Sleep(200)
            }

            if ($sync.LockReply -ne $true) {
                Write-LockState -Detected $null -Current 'Direct locks were kept (user cancelled removal).' -End 'Unchanged (cancelled by user).'
                Log 'Run cancelled — locks were not removed.' 'Warning'
                return
            }
            $sync.LockReply = $null

            # Remove locks
            foreach ($lock in $directLocks) {
                Remove-AzResourceLock -LockId $lock.LockId -Force -EA Stop | Out-Null
                Log "  Removed lock: $($lock.Name)" 'Info'
            }

            # Wait for propagation
            Log 'Waiting for lock removal to propagate...' 'Warning'
            Write-LockState -Detected $null -Current 'Removing locks and waiting for propagation...' -End $null
            $propagated = $false
            for ($i = 1; $i -le 10; $i++) {
                $remaining = @(
                    Get-AzResourceLock -ResourceGroupName $rg -ResourceName $kvName `
                        -ResourceType 'Microsoft.KeyVault/vaults' -EA SilentlyContinue
                )
                if ($remaining.Count -eq 0) { $propagated = $true; break }
                Log "  Still propagating (attempt $i/10)..." 'Warning'
                [System.Threading.Thread]::Sleep(3000)
            }
            if (-not $propagated) {
                Write-LockState -Detected $null -Current 'Lock removal still propagating (timed out).' -End 'Unknown (recheck lock state manually).'
                Log '✗ Locks did not clear in time. Please wait and retry.' 'Error'
                return
            }

            $locksRemoved   = $true
            $locksToRestore = $directLocks
            Write-LockState -Detected $null -Current 'No direct locks currently on Key Vault (removed for operation).' -End 'Pending restore decision.'
            Log '✓ Locks removed.' 'Success'
        }
        elseif ($directLocks.Count -gt 0 -and $preview) {
            $lockNames = Format-LockSummary -Locks $directLocks
            Write-LockState -Detected $lockNames -Current 'Direct lock(s) present (preview mode keeps locks unchanged).' -End 'Unchanged (preview mode).'
            Log "ℹ Lock(s) present but skipped in preview mode: $lockNames" 'Info'
        }
        else {
            Write-LockState -Detected 'None' -Current 'No direct locks found.' -End 'No lock changes required.'
            Log '✓ No direct locks on Key Vault.' 'Success'
        }

        # ── 6. Import CSV ────────────────────────────────────────────────
        Log "Reading CSV: $csvPath" 'Info'
        if (-not (Test-Path $csvPath)) {
            Log "✗ CSV not found: $csvPath" 'Error'
            return
        }

        $users = Import-Csv -Path $csvPath -EA Stop
        if ($users.Count -eq 0) {
            Log '✗ CSV file is empty.' 'Error'
            return
        }
        $first = $users | Select-Object -First 1
        if (-not $first.PSObject.Properties.Name.Contains('UserID')) {
            Log "✗ CSV must have a 'UserID' column." 'Error'
            return
        }
        if (-not $first.PSObject.Properties.Name.Contains('Role')) {
            Log "✗ CSV must have a 'Role' column." 'Error'
            return
        }
        Log "✓ Loaded $($users.Count) user(s) from CSV." 'Success'
        Log ''

        # ── 7. Process assignments ───────────────────────────────────────
        Log '─── Processing role assignments ───' 'Purple'
        $cntSuccess = 0; $cntSkipped = 0; $cntPreview = 0; $cntError = 0

        foreach ($row in $users) {
            $userId = $row.UserID.Trim()
            $role   = $row.Role.Trim()
            Log "  $userId  │  $role  │  $op" 'White'

            try {
                # Resolve UserID → objectId
                $objectId = $null
                if ([guid]::TryParse($userId, [ref][guid]::Empty)) {
                    $objectId = $userId
                }
                else {
                    $adUser = Get-AzADUser -UserPrincipalName $userId -EA SilentlyContinue
                    if (-not $adUser) {
                        $adUser = Get-AzADUser -Filter "mail eq '$userId'" -EA SilentlyContinue
                    }
                    if ($adUser) { $objectId = $adUser.Id }
                    else { throw "User '$userId' not found in Azure AD." }
                }

                if ($op -eq 'Add') {
                    $existing = Get-AzRoleAssignment -ObjectId $objectId `
                        -RoleDefinitionName $role -Scope $kv.ResourceId -EA SilentlyContinue

                    if ($existing) {
                        Log "    ⊘ Already has role '$role' — skipped." 'Warning'
                        $cntSkipped++
                    }
                    elseif ($preview) {
                        Log "    i Would add role '$role'." 'Info'
                        $cntPreview++
                    }
                    else {
                        New-AzRoleAssignment -ObjectId $objectId `
                            -RoleDefinitionName $role -Scope $kv.ResourceId | Out-Null
                        Log "    ✓ Added role '$role'." 'Success'
                        $cntSuccess++
                    }
                }
                else {
                    $kvScope = $kv.ResourceId.TrimEnd('/')
                    $assignments = @(
                        Get-AzRoleAssignment -ObjectId $objectId `
                            -RoleDefinitionName $role -Scope $kv.ResourceId -EA SilentlyContinue
                    )
                    $direct = @($assignments | Where-Object {
                        $_.Scope -and $_.Scope.TrimEnd('/') -eq $kvScope
                    })

                    if ($direct.Count -eq 0) {
                        $inherited = @($assignments | Where-Object {
                            $_.Scope -and $_.Scope.TrimEnd('/') -ne $kvScope
                        })
                        if ($inherited.Count -gt 0) {
                            $scopes = ($inherited.Scope | Select-Object -Unique) -join '; '
                            Log "    ⊘ Role '$role' is inherited from: $scopes" 'Warning'
                        }
                        else {
                            Log "    ⊘ User does not have role '$role' — skipped." 'Warning'
                        }
                        $cntSkipped++
                    }
                    elseif ($preview) {
                        Log "    i Would remove role '$role'." 'Info'
                        $cntPreview++
                    }
                    else {
                        foreach ($a in $direct) {
                            $removed = $false
                            for ($attempt = 1; $attempt -le 6; $attempt++) {
                                try {
                                    Remove-AzRoleAssignment -ObjectId $objectId `
                                        -RoleDefinitionId $a.RoleDefinitionId `
                                        -Scope $a.Scope -EA Stop | Out-Null
                                    $removed = $true; break
                                }
                                catch {
                                    $errTxt = $_.ToString()
                                    if (($errTxt -match 'ScopeLocked' -or $errTxt -match "Conflict") -and $attempt -lt 6) {
                                        Log "    ! Conflict — waiting 10 s (attempt $attempt/6)..." 'Warning'
                                        [System.Threading.Thread]::Sleep(10000)
                                    }
                                    else { throw }
                                }
                            }
                            if (-not $removed) { throw 'Removal did not complete after retries.' }
                        }
                        Log "    ✓ Removed role '$role'." 'Success'
                        $cntSuccess++
                    }
                }
            }
            catch {
                Log "    ✗ $_" 'Error'
                $cntError++
            }
        }

        # ── 8. Restore locks ─────────────────────────────────────────────
        if ($locksRemoved) {
            $sync.NeedLockRestore = $true
            Write-LockState -Detected $null -Current 'Role operation complete. Awaiting restore decision.' -End $null

            $deadline = (Get-Date).AddSeconds(120)
            while ($null -eq $sync.RestoreReply -and (Get-Date) -lt $deadline) {
                [System.Threading.Thread]::Sleep(200)
            }

            if ($sync.RestoreReply -eq $true) {
                Write-LockState -Detected $null -Current 'Restoring previously removed lock(s)...' -End $null
                foreach ($lock in $locksToRestore) {
                    $lockLvl = $null
                    if (-not [string]::IsNullOrWhiteSpace($lock.LockLevel)) { $lockLvl = $lock.LockLevel }
                    elseif (-not [string]::IsNullOrWhiteSpace($lock.Level))  { $lockLvl = $lock.Level }
                    elseif ($lock.Properties -and -not [string]::IsNullOrWhiteSpace($lock.Properties.Level)) {
                        $lockLvl = $lock.Properties.Level
                    }

                    if (-not $lockLvl) { Log "  ⚠ Could not determine lock level for '$($lock.Name)' — skipped." 'Warning'; continue }

                    New-AzResourceLock -LockName $lock.Name -LockLevel $lockLvl `
                        -ResourceGroupName $rg -ResourceName $kvName `
                        -ResourceType 'Microsoft.KeyVault/vaults' -Force | Out-Null
                    Log "  ✓ Restored lock: $($lock.Name)" 'Success'
                }
                Write-LockState -Detected $null -Current 'Lock restore completed.' -End 'Restored to pre-run lock set.'
            }
            else {
                Write-LockState -Detected $null -Current 'Restore skipped by user.' -End 'Not restored (Key Vault remains unlocked).'
                Log '⚠ Locks were NOT restored. Key Vault remains unlocked.' 'Warning'
            }
            $sync.RestoreReply   = $null
            $sync.NeedLockRestore = $false
        }

        $finalDirectLocks = @(
            Get-AzResourceLock -ResourceGroupName $rg -ResourceName $kvName `
                -ResourceType 'Microsoft.KeyVault/vaults' -EA SilentlyContinue
        )
        if ($finalDirectLocks.Count -gt 0) {
            $finalNames = Format-LockSummary -Locks $finalDirectLocks
            Write-LockState -Detected $null -Current "Current lock state: $finalNames" -End "Post-run direct lock(s): $finalNames"
        }
        else {
            Write-LockState -Detected $null -Current 'Current lock state: none (no direct Key Vault locks).' -End 'Post-run direct lock(s): none.'
        }

        # ── 9. Summary ───────────────────────────────────────────────────
        Log ''
        Log '─────────── Summary ───────────' 'Purple'
        Log "  ✓ Success : $cntSuccess" 'Success'
        Log "  ⊘ Skipped : $cntSkipped" 'Warning'
        Log "  i Preview : $cntPreview" 'Info'
        Log "  ✗ Errors  : $cntError"  'Error'
        Log '───────────────────────────────' 'Purple'
    }
    catch {
        Log "Fatal error: $_" 'Error'
    }
    finally {
        $sync.Done = $true
    }
}
#endregion

#region ── Build the main window ────────────────────────────────────────────

$Form                  = New-Object System.Windows.Forms.Form
$Form.Text             = 'Azure Key Vault Role Manager'
$Form.Size             = New-Object System.Drawing.Size(920, 840)
$Form.MinimumSize      = New-Object System.Drawing.Size(900, 720)
$Form.StartPosition    = 'CenterScreen'
$Form.BackColor        = $C.Bg
$Form.ForeColor        = $C.Text
$Form.Font             = New-Object System.Drawing.Font('Segoe UI', 10)

$MainSplit                    = New-Object System.Windows.Forms.SplitContainer
$MainSplit.Dock               = 'Fill'
$MainSplit.Orientation        = 'Horizontal'
$MainSplit.SplitterWidth      = 7
$MainSplit.BackColor          = $C.Bg
$Form.Controls.Add($MainSplit)

# ── Single scrollable config panel ──────────────────────────────────────────
$PnlTop             = New-Object System.Windows.Forms.Panel
$PnlTop.Dock        = 'Fill'
$PnlTop.AutoScroll  = $true
$PnlTop.BackColor   = $C.Panel
$MainSplit.Panel1.Controls.Add($PnlTop)

# Layout grid constants for consistent spacing and alignment.
$LayoutMargin       = 10
$LayoutGap          = 12
$LayoutColumnWidth  = 430
$LayoutLeftX        = $LayoutMargin
$LayoutRightX       = $LayoutLeftX + $LayoutColumnWidth + $LayoutGap
$LayoutFullWidth    = ($LayoutColumnWidth * 2) + $LayoutGap
$LayoutInputLabelX  = 150
$LayoutInputWidth   = 266

function Write-Label {
    param([string]$Text, [int]$X, [int]$Y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.AutoSize = $true; $l.ForeColor = $C.Text
    return $l
}
function Write-TextBox {
    param([int]$X, [int]$Y, [int]$W = 500, [string]$PlaceHolder = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = New-Object System.Drawing.Point($X, $Y)
    $t.Size      = New-Object System.Drawing.Size($W, 24)
    $t.BackColor = $C.Input; $t.ForeColor = $C.Text
    $t.BorderStyle = 'FixedSingle'
    if ($PlaceHolder) {
        $t.Tag = $PlaceHolder
        $t.ForeColor = $C.Muted; $t.Text = $PlaceHolder
        $t.Add_Enter({
            param($textBoxControl)
            if ($textBoxControl.ForeColor -eq $C.Muted) {
                $textBoxControl.Text = ''
                $textBoxControl.ForeColor = $C.Text
            }
        })
        $t.Add_Leave({
            param($textBoxControl)
            if ([string]::IsNullOrWhiteSpace($textBoxControl.Text)) {
                $textBoxControl.ForeColor = $C.Muted
                $textBoxControl.Text = [string]$textBoxControl.Tag
            }
        })
    }
    return $t
}

# ── Left col: Azure Connection ────────────────────────────────────────────────
$GrpConn            = New-Object System.Windows.Forms.GroupBox
$GrpConn.Text       = 'Azure Connection'
$GrpConn.Location   = New-Object System.Drawing.Point($LayoutLeftX, 10)
$GrpConn.Size       = New-Object System.Drawing.Size($LayoutColumnWidth, 118)
$GrpConn.ForeColor  = $C.Muted
$GrpConn.BackColor  = $C.GroupBg
$PnlTop.Controls.Add($GrpConn)

$GrpConn.Controls.Add((Write-Label -Text 'Subscription ID' -X 14 -Y 26))
$TxtSubId  = Write-TextBox -X $LayoutInputLabelX -Y 22 -W $LayoutInputWidth -PlaceHolder 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$GrpConn.Controls.Add($TxtSubId)

$GrpConn.Controls.Add((Write-Label -Text 'Resource Group'  -X 14 -Y 56))
$TxtRG     = Write-TextBox -X $LayoutInputLabelX -Y 52 -W $LayoutInputWidth
$GrpConn.Controls.Add($TxtRG)

$GrpConn.Controls.Add((Write-Label -Text 'Key Vault Name'  -X 14 -Y 86))
$TxtKV     = Write-TextBox -X $LayoutInputLabelX -Y 82 -W $LayoutInputWidth
$GrpConn.Controls.Add($TxtKV)

# ── Right col: Role Operation ─────────────────────────────────────────────────
$GrpOp             = New-Object System.Windows.Forms.GroupBox
$GrpOp.Text        = 'Role Operation'
$GrpOp.Location    = New-Object System.Drawing.Point($LayoutRightX, 10)
$GrpOp.Size        = New-Object System.Drawing.Size($LayoutColumnWidth, 96)
$GrpOp.ForeColor   = $C.Muted
$GrpOp.BackColor   = $C.GroupBg
$PnlTop.Controls.Add($GrpOp)

$RbAdd             = New-Object System.Windows.Forms.RadioButton
$RbAdd.Text        = 'Add roles to users'
$RbAdd.Location    = New-Object System.Drawing.Point(16, 26)
$RbAdd.AutoSize    = $true
$RbAdd.ForeColor   = $C.Text
$RbAdd.Checked     = $true
$GrpOp.Controls.Add($RbAdd)

$RbRemove          = New-Object System.Windows.Forms.RadioButton
$RbRemove.Text     = 'Remove roles from users'
$RbRemove.Location = New-Object System.Drawing.Point(16, 54)
$RbRemove.AutoSize = $true
$RbRemove.ForeColor = $C.Text
$GrpOp.Controls.Add($RbRemove)

# ── Right col: Execution Mode ────────────────────────────────────────────────
$GrpMode           = New-Object System.Windows.Forms.GroupBox
$GrpMode.Text      = 'Execution Mode'
$GrpMode.Location  = New-Object System.Drawing.Point($LayoutRightX, 118)
$GrpMode.Size      = New-Object System.Drawing.Size($LayoutColumnWidth, 80)
$GrpMode.ForeColor = $C.Muted
$GrpMode.BackColor = $C.GroupBg
$PnlTop.Controls.Add($GrpMode)

$ChkPreview        = New-Object System.Windows.Forms.CheckBox
$ChkPreview.Text   = 'Preview mode  (no changes will be made)'
$ChkPreview.Location = New-Object System.Drawing.Point(16, 26)
$ChkPreview.AutoSize = $true
$ChkPreview.ForeColor = $C.Text
$GrpMode.Controls.Add($ChkPreview)

$LblPreviewHint    = New-Object System.Windows.Forms.Label
$LblPreviewHint.Text = 'Shows what would change without applying anything.'
$LblPreviewHint.Location = New-Object System.Drawing.Point(16, 52)
$LblPreviewHint.AutoSize = $true
$LblPreviewHint.ForeColor = $C.Muted
$GrpMode.Controls.Add($LblPreviewHint)

# ── Right col: Input File ─────────────────────────────────────────────────────
$GrpCsv             = New-Object System.Windows.Forms.GroupBox
$GrpCsv.Text        = 'Input File'
$GrpCsv.Location    = New-Object System.Drawing.Point($LayoutRightX, 210)
$GrpCsv.Size        = New-Object System.Drawing.Size($LayoutColumnWidth, 108)
$GrpCsv.ForeColor   = $C.Muted
$GrpCsv.BackColor   = $C.GroupBg
$PnlTop.Controls.Add($GrpCsv)

$LblCsvPath = New-Object System.Windows.Forms.Label
$LblCsvPath.Text = 'CSV File'
$LblCsvPath.Location = New-Object System.Drawing.Point(16, 22)
$LblCsvPath.AutoSize = $true
$LblCsvPath.ForeColor = $C.Text
$GrpCsv.Controls.Add($LblCsvPath)

$TxtCSV     = Write-TextBox -X 16 -Y 42 -W 378
$TxtCSV.Text = (Join-Path $ScriptDir 'users_and_roles.csv')
$TxtCSV.ForeColor = $C.Text
$GrpCsv.Controls.Add($TxtCSV)

$BtnBrowse             = New-Object System.Windows.Forms.Button
$BtnBrowse.Text        = 'Browse...'
$BtnBrowse.Location    = New-Object System.Drawing.Point(328, 72)
$BtnBrowse.Size        = New-Object System.Drawing.Size(88, 26)
$BtnBrowse.FlatStyle   = 'Flat'
$BtnBrowse.BackColor   = $C.Panel
$BtnBrowse.ForeColor   = $C.Text
$BtnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $ofd.InitialDirectory = $ScriptDir
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtCSV.Text = $ofd.FileName
        $TxtCSV.ForeColor = $C.Text
    }
})
$GrpCsv.Controls.Add($BtnBrowse)

$LblCsvHint = New-Object System.Windows.Forms.Label
$LblCsvHint.Text = 'Expected columns: UserID, Role'
$LblCsvHint.Location = New-Object System.Drawing.Point(16, 82)
$LblCsvHint.AutoSize = $true
$LblCsvHint.ForeColor = $C.Muted
$GrpCsv.Controls.Add($LblCsvHint)

# ── Left col: Azure Login ─────────────────────────────────────────────────────
$GrpAuthStatus             = New-Object System.Windows.Forms.GroupBox
$GrpAuthStatus.Text        = 'Azure Login'
$GrpAuthStatus.Location    = New-Object System.Drawing.Point($LayoutLeftX, 140)
$GrpAuthStatus.Size        = New-Object System.Drawing.Size($LayoutColumnWidth, 178)
$GrpAuthStatus.ForeColor   = $C.Muted
$GrpAuthStatus.BackColor   = $C.GroupBg
$PnlTop.Controls.Add($GrpAuthStatus)

$LblAuthState            = New-Object System.Windows.Forms.Label
$LblAuthState.Text       = 'State'
$LblAuthState.Location   = New-Object System.Drawing.Point(14, 28)
$LblAuthState.AutoSize   = $true
$LblAuthState.ForeColor  = $C.Text
$GrpAuthStatus.Controls.Add($LblAuthState)

$PnlAuthDot              = New-Object System.Windows.Forms.Panel
$PnlAuthDot.Location     = New-Object System.Drawing.Point(82, 28)
$PnlAuthDot.Size         = New-Object System.Drawing.Size(14, 14)
$PnlAuthDot.BackColor    = $C.Error
$GrpAuthStatus.Controls.Add($PnlAuthDot)

$LblAuthStateValue            = New-Object System.Windows.Forms.Label
$LblAuthStateValue.Text       = 'Not signed in'
$LblAuthStateValue.Location   = New-Object System.Drawing.Point(104, 26)
$LblAuthStateValue.AutoSize   = $true
$LblAuthStateValue.ForeColor  = $C.Warning
$GrpAuthStatus.Controls.Add($LblAuthStateValue)

$LblAuthAccount            = New-Object System.Windows.Forms.Label
$LblAuthAccount.Text       = 'Account'
$LblAuthAccount.Location   = New-Object System.Drawing.Point(14, 54)
$LblAuthAccount.AutoSize   = $true
$LblAuthAccount.ForeColor  = $C.Text
$GrpAuthStatus.Controls.Add($LblAuthAccount)

$LblAuthAccountValue            = New-Object System.Windows.Forms.Label
$LblAuthAccountValue.Text       = '-'
$LblAuthAccountValue.Location   = New-Object System.Drawing.Point(104, 54)
$LblAuthAccountValue.Size       = New-Object System.Drawing.Size(312, 20)
$LblAuthAccountValue.ForeColor  = $C.Muted
$GrpAuthStatus.Controls.Add($LblAuthAccountValue)

$LblAuthTenant            = New-Object System.Windows.Forms.Label
$LblAuthTenant.Text       = 'Tenant'
$LblAuthTenant.Location   = New-Object System.Drawing.Point(14, 78)
$LblAuthTenant.AutoSize   = $true
$LblAuthTenant.ForeColor  = $C.Text
$GrpAuthStatus.Controls.Add($LblAuthTenant)

$LblAuthTenantValue            = New-Object System.Windows.Forms.Label
$LblAuthTenantValue.Text       = '-'
$LblAuthTenantValue.Location   = New-Object System.Drawing.Point(104, 78)
$LblAuthTenantValue.Size       = New-Object System.Drawing.Size(312, 20)
$LblAuthTenantValue.ForeColor  = $C.Muted
$GrpAuthStatus.Controls.Add($LblAuthTenantValue)

$LblAuthSub            = New-Object System.Windows.Forms.Label
$LblAuthSub.Text       = 'Subscription'
$LblAuthSub.Location   = New-Object System.Drawing.Point(14, 102)
$LblAuthSub.AutoSize   = $true
$LblAuthSub.ForeColor  = $C.Text
$GrpAuthStatus.Controls.Add($LblAuthSub)

$LblAuthSubValue            = New-Object System.Windows.Forms.Label
$LblAuthSubValue.Text       = '-'
$LblAuthSubValue.Location   = New-Object System.Drawing.Point(104, 102)
$LblAuthSubValue.Size       = New-Object System.Drawing.Size(312, 20)
$LblAuthSubValue.ForeColor  = $C.Muted
$GrpAuthStatus.Controls.Add($LblAuthSubValue)

$BtnLoginAzure             = New-Object System.Windows.Forms.Button
$BtnLoginAzure.Text        = 'Sign In'
$BtnLoginAzure.Location    = New-Object System.Drawing.Point(14, 134)
$BtnLoginAzure.Size        = New-Object System.Drawing.Size(126, 30)
$BtnLoginAzure.FlatStyle   = 'Flat'
$BtnLoginAzure.BackColor   = $C.Accent
$BtnLoginAzure.ForeColor   = $C.Text
$BtnLoginAzure.Cursor      = [System.Windows.Forms.Cursors]::Hand
$GrpAuthStatus.Controls.Add($BtnLoginAzure)

$BtnRefreshAuth             = New-Object System.Windows.Forms.Button
$BtnRefreshAuth.Text        = 'Refresh'
$BtnRefreshAuth.Location    = New-Object System.Drawing.Point(152, 134)
$BtnRefreshAuth.Size        = New-Object System.Drawing.Size(126, 30)
$BtnRefreshAuth.FlatStyle   = 'Flat'
$BtnRefreshAuth.BackColor   = $C.Panel
$BtnRefreshAuth.ForeColor   = $C.Text
$BtnRefreshAuth.Cursor      = [System.Windows.Forms.Cursors]::Hand
$GrpAuthStatus.Controls.Add($BtnRefreshAuth)

$BtnLogoutAzure             = New-Object System.Windows.Forms.Button
$BtnLogoutAzure.Text        = 'Sign Out'
$BtnLogoutAzure.Location    = New-Object System.Drawing.Point(290, 134)
$BtnLogoutAzure.Size        = New-Object System.Drawing.Size(126, 30)
$BtnLogoutAzure.FlatStyle   = 'Flat'
$BtnLogoutAzure.BackColor   = $C.Panel
$BtnLogoutAzure.ForeColor   = $C.Text
$BtnLogoutAzure.Cursor      = [System.Windows.Forms.Cursors]::Hand
$GrpAuthStatus.Controls.Add($BtnLogoutAzure)

# ── Full-width: Lock Information (y=330) ──────────────────────────────────────
$GrpLockInfo                = New-Object System.Windows.Forms.GroupBox
$GrpLockInfo.Text           = 'Lock Information'
$GrpLockInfo.Location       = New-Object System.Drawing.Point($LayoutLeftX, 330)
$GrpLockInfo.Size           = New-Object System.Drawing.Size($LayoutFullWidth, 124)
$GrpLockInfo.ForeColor      = $C.Muted
$GrpLockInfo.BackColor      = $C.GroupBg
$PnlTop.Controls.Add($GrpLockInfo)

$LblLockDetected            = New-Object System.Windows.Forms.Label
$LblLockDetected.Text       = 'Detected before run'
$LblLockDetected.Location   = New-Object System.Drawing.Point(14, 28)
$LblLockDetected.AutoSize   = $true
$LblLockDetected.ForeColor  = $C.Text
$GrpLockInfo.Controls.Add($LblLockDetected)

$TxtLockDetected               = New-Object System.Windows.Forms.TextBox
$TxtLockDetected.Location      = New-Object System.Drawing.Point(176, 24)
$TxtLockDetected.Size          = New-Object System.Drawing.Size(682, 24)
$TxtLockDetected.ReadOnly      = $true
$TxtLockDetected.BackColor     = $C.Input
$TxtLockDetected.ForeColor     = $C.Text
$TxtLockDetected.BorderStyle   = 'FixedSingle'
$TxtLockDetected.Text          = $sync.LockDetected
$GrpLockInfo.Controls.Add($TxtLockDetected)

$LblLockCurrent            = New-Object System.Windows.Forms.Label
$LblLockCurrent.Text       = 'Current status'
$LblLockCurrent.Location   = New-Object System.Drawing.Point(14, 58)
$LblLockCurrent.AutoSize   = $true
$LblLockCurrent.ForeColor  = $C.Text
$GrpLockInfo.Controls.Add($LblLockCurrent)

$TxtLockCurrent               = New-Object System.Windows.Forms.TextBox
$TxtLockCurrent.Location      = New-Object System.Drawing.Point(176, 54)
$TxtLockCurrent.Size          = New-Object System.Drawing.Size(682, 24)
$TxtLockCurrent.ReadOnly      = $true
$TxtLockCurrent.BackColor     = $C.Input
$TxtLockCurrent.ForeColor     = $C.Text
$TxtLockCurrent.BorderStyle   = 'FixedSingle'
$TxtLockCurrent.Text          = $sync.LockCurrent
$GrpLockInfo.Controls.Add($TxtLockCurrent)

$LblLockEnd            = New-Object System.Windows.Forms.Label
$LblLockEnd.Text       = 'Post-run status'
$LblLockEnd.Location   = New-Object System.Drawing.Point(14, 88)
$LblLockEnd.AutoSize   = $true
$LblLockEnd.ForeColor  = $C.Text
$GrpLockInfo.Controls.Add($LblLockEnd)

$TxtLockEnd               = New-Object System.Windows.Forms.TextBox
$TxtLockEnd.Location      = New-Object System.Drawing.Point(176, 84)
$TxtLockEnd.Size          = New-Object System.Drawing.Size(682, 24)
$TxtLockEnd.ReadOnly      = $true
$TxtLockEnd.BackColor     = $C.Input
$TxtLockEnd.ForeColor     = $C.Text
$TxtLockEnd.BorderStyle   = 'FixedSingle'
$TxtLockEnd.Text          = $sync.LockEnd
$GrpLockInfo.Controls.Add($TxtLockEnd)

# ── Full-width: Required Permissions (y=466) ──────────────────────────────────
$GrpPerms             = New-Object System.Windows.Forms.GroupBox
$GrpPerms.Text        = 'Required Permissions'
$GrpPerms.Location    = New-Object System.Drawing.Point($LayoutLeftX, 466)
$GrpPerms.Size        = New-Object System.Drawing.Size($LayoutFullWidth, 160)
$GrpPerms.ForeColor   = $C.Muted
$GrpPerms.BackColor   = $C.GroupBg
$PnlTop.Controls.Add($GrpPerms)

$LblPerm1           = New-Object System.Windows.Forms.Label
$LblPerm1.Text      = '1)  Role assignments on Key Vault scope:'
$LblPerm1.Location  = New-Object System.Drawing.Point(14, 24)
$LblPerm1.AutoSize  = $true
$LblPerm1.ForeColor = $C.Text
$GrpPerms.Controls.Add($LblPerm1)

$LblPerm1b           = New-Object System.Windows.Forms.Label
$LblPerm1b.Text      = '     Microsoft.Authorization/roleAssignments/read|write|delete  (Owner or User Access Administrator)'
$LblPerm1b.Location  = New-Object System.Drawing.Point(14, 42)
$LblPerm1b.AutoSize  = $true
$LblPerm1b.ForeColor = $C.Muted
$GrpPerms.Controls.Add($LblPerm1b)

$LblPerm2           = New-Object System.Windows.Forms.Label
$LblPerm2.Text      = '2)  Lock management (also at RG/Subscription scope for inherited locks):'
$LblPerm2.Location  = New-Object System.Drawing.Point(14, 66)
$LblPerm2.AutoSize  = $true
$LblPerm2.ForeColor = $C.Text
$GrpPerms.Controls.Add($LblPerm2)

$LblPerm2b           = New-Object System.Windows.Forms.Label
$LblPerm2b.Text      = '     Microsoft.Authorization/locks/read|write|delete'
$LblPerm2b.Location  = New-Object System.Drawing.Point(14, 84)
$LblPerm2b.AutoSize  = $true
$LblPerm2b.ForeColor = $C.Muted
$GrpPerms.Controls.Add($LblPerm2b)

$PnlPermDot           = New-Object System.Windows.Forms.Panel
$PnlPermDot.Location  = New-Object System.Drawing.Point(14, 120)
$PnlPermDot.Size      = New-Object System.Drawing.Size(14, 14)
$PnlPermDot.BackColor = $C.Muted
$GrpPerms.Controls.Add($PnlPermDot)

$LblPermStatus           = New-Object System.Windows.Forms.Label
$LblPermStatus.Text      = 'Not checked'
$LblPermStatus.Location  = New-Object System.Drawing.Point(36, 118)
$LblPermStatus.AutoSize  = $true
$LblPermStatus.ForeColor = $C.Muted
$GrpPerms.Controls.Add($LblPermStatus)

$BtnCheckPerms           = New-Object System.Windows.Forms.Button
$BtnCheckPerms.Text      = 'Check Permissions'
$BtnCheckPerms.Location  = New-Object System.Drawing.Point(720, 112)
$BtnCheckPerms.Size      = New-Object System.Drawing.Size(140, 30)
$BtnCheckPerms.FlatStyle = 'Flat'
$BtnCheckPerms.BackColor = $C.Panel
$BtnCheckPerms.ForeColor = $C.Text
$BtnCheckPerms.Cursor    = [System.Windows.Forms.Cursors]::Hand
$GrpPerms.Controls.Add($BtnCheckPerms)

# ── Action buttons (below tabs) ──────────────────────────────────────────────
$BtnRun            = New-Object System.Windows.Forms.Button
$BtnRun.Text       = '▶  Run'
$BtnRun.Location   = New-Object System.Drawing.Point(10, 320)
$BtnRun.Size       = New-Object System.Drawing.Size(130, 34)
$BtnRun.FlatStyle  = 'Flat'
$BtnRun.BackColor  = $C.Accent
$BtnRun.ForeColor  = $C.Text
$BtnRun.Font       = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$BtnRun.Cursor     = [System.Windows.Forms.Cursors]::Hand
$BtnRun.Location   = New-Object System.Drawing.Point(10, 8)
$MainSplit.Panel2.Controls.Add($BtnRun)

$BtnClear          = New-Object System.Windows.Forms.Button
$BtnClear.Text     = 'Clear Log'
$BtnClear.Location = New-Object System.Drawing.Point(155, 320)
$BtnClear.Size     = New-Object System.Drawing.Size(100, 34)
$BtnClear.FlatStyle = 'Flat'
$BtnClear.BackColor = $C.Panel
$BtnClear.ForeColor = $C.Text
$BtnClear.Cursor    = [System.Windows.Forms.Cursors]::Hand
$BtnClear.Add_Click({ $RtbLog.Clear() })
$BtnClear.Location = New-Object System.Drawing.Point(155, 8)
$MainSplit.Panel2.Controls.Add($BtnClear)

# ── Status label ─────────────────────────────────────────────────────────────
$LblStatus         = New-Object System.Windows.Forms.Label
$LblStatus.Text    = 'Ready'
$LblStatus.Location = New-Object System.Drawing.Point(270, 16)
$LblStatus.Size    = New-Object System.Drawing.Size(470, 20)
$LblStatus.ForeColor = $C.Muted
$MainSplit.Panel2.Controls.Add($LblStatus)

# ── Output log ────────────────────────────────────────────────────────────────
$LblLog            = New-Object System.Windows.Forms.Label
$LblLog.Text       = 'Output Log'
$LblLog.Location   = New-Object System.Drawing.Point(10, 48)
$LblLog.AutoSize   = $true
$LblLog.ForeColor  = $C.Muted
$MainSplit.Panel2.Controls.Add($LblLog)

$RtbLog            = New-Object System.Windows.Forms.RichTextBox
$RtbLog.Location   = New-Object System.Drawing.Point(10, 68)
$RtbLog.Size       = New-Object System.Drawing.Size(862, 230)
$RtbLog.Anchor     = 'Top,Left,Right,Bottom'
$RtbLog.BackColor  = $C.LogBg
$RtbLog.ForeColor  = $C.Text
$RtbLog.Font       = New-Object System.Drawing.Font('Consolas', 10)
$RtbLog.ReadOnly   = $true
$RtbLog.BorderStyle = 'None'
$RtbLog.ScrollBars = 'Vertical'
$MainSplit.Panel2.Controls.Add($RtbLog)

#endregion

#region ── Azure login status helpers ───────────────────────────────────────
function Write-AzureAuthIndicator {
    param(
        [bool]$IsSignedIn,
        [string]$StateText,
        [string]$AccountText,
        [string]$TenantText,
        [string]$SubscriptionText
    )

    if ($IsSignedIn) {
        $PnlAuthDot.BackColor = $C.Success
        $LblAuthStateValue.ForeColor = $C.Success
    }
    else {
        $PnlAuthDot.BackColor = $C.Error
        $LblAuthStateValue.ForeColor = $C.Warning
    }

    $LblAuthStateValue.Text = $StateText
    $LblAuthAccountValue.Text = $AccountText
    $LblAuthTenantValue.Text = $TenantText
    $LblAuthSubValue.Text = $SubscriptionText
}

function Sync-AzureLoginStatus {
    try {
        Import-Module Az.Accounts -ErrorAction Stop | Out-Null
        $ctx = Get-AzContext -ErrorAction SilentlyContinue

        if ($null -ne $ctx -and $null -ne $ctx.Account -and $null -ne $ctx.Subscription) {
            $subDisplay = "$($ctx.Subscription.Name) ($($ctx.Subscription.Id))"
            Write-AzureAuthIndicator -IsSignedIn $true `
                -StateText 'Signed in' `
                -AccountText ([string]$ctx.Account.Id) `
                -TenantText ([string]$ctx.Tenant.Id) `
                -SubscriptionText $subDisplay
        }
        else {
            Write-AzureAuthIndicator -IsSignedIn $false `
                -StateText 'Not signed in' `
                -AccountText '-' `
                -TenantText '-' `
                -SubscriptionText '-'
        }
    }
    catch {
        Write-AzureAuthIndicator -IsSignedIn $false `
            -StateText 'Az.Accounts not available' `
            -AccountText '-' `
            -TenantText '-' `
            -SubscriptionText '-'
    }
}

$BtnRefreshAuth.Add_Click({
    Sync-AzureLoginStatus
})

$BtnLoginAzure.Add_Click({
    try {
        Import-Module Az.Accounts -ErrorAction Stop | Out-Null
        Connect-AzAccount -ErrorAction Stop | Out-Null

        $subFromSettings = $TxtSubId.Text.Trim()
        if ([guid]::TryParse($subFromSettings, [ref][guid]::Empty)) {
            Set-AzContext -SubscriptionId $subFromSettings -ErrorAction SilentlyContinue | Out-Null
        }

        Sync-AzureLoginStatus
        [System.Windows.Forms.MessageBox]::Show(
            'Successfully signed in to Azure.',
            'Azure Login',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Sync-AzureLoginStatus
        [System.Windows.Forms.MessageBox]::Show(
            "Azure sign-in failed: $_",
            'Azure Login',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

$BtnLogoutAzure.Add_Click({
    try {
        Import-Module Az.Accounts -ErrorAction Stop | Out-Null

        Disconnect-AzAccount -Scope Process -ErrorAction SilentlyContinue | Out-Null
        Clear-AzContext -Scope Process -Force -ErrorAction SilentlyContinue | Out-Null

        Sync-AzureLoginStatus
        [System.Windows.Forms.MessageBox]::Show(
            'Signed out from Azure for this PowerShell session.',
            'Azure Login',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Sync-AzureLoginStatus
        [System.Windows.Forms.MessageBox]::Show(
            "Azure sign-out failed: $_",
            'Azure Login',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
})

Sync-AzureLoginStatus

$BtnCheckPerms.Add_Click({
    $sub = $TxtSubId.Text.Trim()
    $rg  = $TxtRG.Text.Trim()
    $kv  = $TxtKV.Text.Trim()
    $ph  = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

    if ([string]::IsNullOrWhiteSpace($sub) -or $sub -eq $ph -or
        [string]::IsNullOrWhiteSpace($rg)  -or
        [string]::IsNullOrWhiteSpace($kv)) {
        $PnlPermDot.BackColor    = $C.Warning
        $LblPermStatus.ForeColor = $C.Warning
        $LblPermStatus.Text      = 'Fill in Subscription ID, Resource Group and Key Vault Name first.'
        return
    }

    $PnlPermDot.BackColor    = $C.Muted
    $LblPermStatus.ForeColor = $C.Muted
    $LblPermStatus.Text      = 'Checking...'
    [System.Windows.Forms.Application]::DoEvents()

    try {
        Import-Module Az.Accounts  -ErrorAction Stop | Out-Null
        Import-Module Az.Resources -ErrorAction Stop | Out-Null

        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        if ($null -eq $ctx) {
            $PnlPermDot.BackColor    = $C.Error
            $LblPermStatus.ForeColor = $C.Error
            $LblPermStatus.Text      = 'Not signed in — sign in first.'
            return
        }

        Set-AzContext -SubscriptionId $sub -ErrorAction Stop | Out-Null
        $kvScope = "/subscriptions/$sub/resourceGroups/$rg/providers/Microsoft.KeyVault/vaults/$kv"
        $raOk   = $false
        $lockOk = $false
        try { Get-AzRoleAssignment -Scope $kvScope -ErrorAction Stop | Out-Null; $raOk   = $true } catch { $null = $_ }
        try { Get-AzResourceLock  -ResourceGroupName $rg -ResourceName $kv -ResourceType 'Microsoft.KeyVault/vaults' -ErrorAction Stop | Out-Null; $lockOk = $true } catch { $null = $_ }

        if ($raOk -and $lockOk) {
            $PnlPermDot.BackColor    = $C.Success
            $LblPermStatus.ForeColor = $C.Success
            $LblPermStatus.Text      = 'Role assignment & lock permissions confirmed.'
        } elseif ($raOk) {
            $PnlPermDot.BackColor    = $C.Warning
            $LblPermStatus.ForeColor = $C.Warning
            $LblPermStatus.Text      = 'Role assignment OK — lock management permission could not be confirmed.'
        } elseif ($lockOk) {
            $PnlPermDot.BackColor    = $C.Warning
            $LblPermStatus.ForeColor = $C.Warning
            $LblPermStatus.Text      = 'Lock management OK — role assignment permission could not be confirmed.'
        } else {
            $PnlPermDot.BackColor    = $C.Error
            $LblPermStatus.ForeColor = $C.Error
            $LblPermStatus.Text      = 'Could not confirm permissions — check your role assignments.'
        }
    }
    catch {
        $PnlPermDot.BackColor    = $C.Error
        $LblPermStatus.ForeColor = $C.Error
        $LblPermStatus.Text      = "Error: $_"
    }
})
#endregion

#region ── Input validation ──────────────────────────────────────────────────
function Get-InputError {
    $e = @()
    $sub = $TxtSubId.Text.Trim()
    $rg  = $TxtRG.Text.Trim()
    $kv  = $TxtKV.Text.Trim()
    $csv = $TxtCSV.Text.Trim()
    $ph  = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

    if ([string]::IsNullOrWhiteSpace($sub) -or $sub -eq $ph)       { $e += 'Subscription ID is required.' }
    elseif (-not [guid]::TryParse($sub, [ref][guid]::Empty))        { $e += 'Subscription ID must be a valid GUID.' }
    if ([string]::IsNullOrWhiteSpace($rg))                          { $e += 'Resource Group is required.' }
    if ([string]::IsNullOrWhiteSpace($kv))                          { $e += 'Key Vault Name is required.' }
    if ([string]::IsNullOrWhiteSpace($csv))                         { $e += 'CSV file path is required.' }
    elseif (-not (Test-Path $csv))                                  { $e += "CSV file not found: $csv" }
    return $e
}
#endregion

#region ── Background-runspace polling timer ─────────────────────────────────
$colorMap = @{
    Success = $C.Success; Warning = $C.Warning; Error = $C.Error
    Info    = $C.Info;    Purple  = $C.Purple;  White = $C.White
}

$Timer          = New-Object System.Windows.Forms.Timer
$Timer.Interval = 150
$Timer.Add_Tick({
    $TxtLockDetected.Text = [string]$sync.LockDetected
    $TxtLockCurrent.Text  = [string]$sync.LockCurrent
    $TxtLockEnd.Text      = [string]$sync.LockEnd

    # Flush pending log messages
    while ($sync.Messages.Count -gt 0) {
        $msg = $null
        if ($sync.Messages.Count -gt 0) {
            $msg = $sync.Messages[0]
            $sync.Messages.RemoveAt(0)
        }
        if ($null -ne $msg) {
            $col = if ($colorMap.ContainsKey($msg.Kind)) { $colorMap[$msg.Kind] } else { $C.Text }
            Add-LogLine -Rtb $RtbLog -Text $msg.Text -Color $col
        }
    }

    # Lock-removal confirmation needed
    if ($sync.NeedsLockReply) {
        $sync.NeedsLockReply = $false
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "The following lock(s) must be removed to continue:`n`n$($sync.LockNames)`n`nThey will be restored after the operation. Proceed?",
            'Locks Detected',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $sync.LockReply = ($ans -eq [System.Windows.Forms.DialogResult]::Yes)
    }

    # Lock-restore confirmation needed
    if ($sync.NeedLockRestore -and $null -eq $sync.RestoreReply) {
        $sync.NeedLockRestore = $false
        $ans = [System.Windows.Forms.MessageBox]::Show(
            'Restore the lock(s) that were removed before the operation?',
            'Restore Locks',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        $sync.RestoreReply = ($ans -eq [System.Windows.Forms.DialogResult]::Yes)
    }

    # Job finished
    if ($sync.Done) {
        $Timer.Stop()
        $BtnRun.Enabled  = $true
        $BtnRun.Text     = '▶  Run'
        $BtnRun.BackColor = $C.Accent
        $LblStatus.Text  = 'Completed.'
        $LblStatus.ForeColor = $C.Success
    }
})
#endregion

#region ── Run button handler ─────────────────────────────────────────────────
$BtnRun.Add_Click({
    $errors = Get-InputError
    if ($errors.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            ($errors -join "`n"),
            'Please fix the following',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $subId   = $TxtSubId.Text.Trim()
    $rg      = $TxtRG.Text.Trim()
    $kvName  = $TxtKV.Text.Trim()
    $csvPath = $TxtCSV.Text.Trim()
    $op      = if ($RbAdd.Checked) { 'Add' } else { 'Remove' }
    $preview = $ChkPreview.Checked

    $RtbLog.Clear()
    $BtnRun.Enabled  = $false
    $BtnRun.Text     = '⏳ Running…'
    $BtnRun.BackColor = $C.Panel
    $LblStatus.Text  = 'Running — please wait…'
    $LblStatus.ForeColor = $C.Warning

    # Reset sync state
    $sync.Done           = $false
    $sync.NeedsLockReply = $false
    $sync.LockReply      = $null
    $sync.NeedLockRestore= $false
    $sync.RestoreReply   = $null
    $sync.LockDetected   = 'Checking lock state...'
    $sync.LockCurrent    = 'Run started.'
    $sync.LockEnd        = 'Pending.'
    $sync.Messages.Clear()

    $TxtLockDetected.Text = [string]$sync.LockDetected
    $TxtLockCurrent.Text  = [string]$sync.LockCurrent
    $TxtLockEnd.Text      = [string]$sync.LockEnd

    # Launch background runspace
    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($AzureWork)
    [void]$ps.AddArgument($sync)
    [void]$ps.AddArgument($subId)
    [void]$ps.AddArgument($rg)
    [void]$ps.AddArgument($kvName)
    [void]$ps.AddArgument($op)
    [void]$ps.AddArgument($preview)
    [void]$ps.AddArgument($csvPath)

    [void]$ps.BeginInvoke()

    $Timer.Start()
})
#endregion

# Set splitter position after layout is realized
$Form.Add_Load({
    $MainSplit.Panel1MinSize    = 250
    $MainSplit.Panel2MinSize    = 180
    $MainSplit.SplitterDistance = 460
})

# Launch
[System.Windows.Forms.Application]::Run($Form)
