<#
.SYNOPSIS
    VAM-VarHandler: Pack, Unpack, Unify, Sanitize, Repair, Restore, View .var packages.
    Includes High-Performance GUI and CLI support.

.DESCRIPTION
    - Version: 2.0 (Renamed to VAM-VarHandler)
    - Feature: Repair can replace corrupt files with valid empty stubs.
    - Feature: Viewer with search/highlight.
    - Safety: Full backup conflict management and non-destructive default behavior.

.AUTHOR
    Co-authored by User & Gemini (Google AI)
#>

# Configuration
$ScriptVersion = "2.0"
$ErrorActionPreference = "Stop" 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Global State
$global:GuiLogBox = $null
$global:IsGuiMode = $false
$global:CancelRequested = $false

# ------------------- Infrastructure Functions -------------------

function Log-Message {
    param([string]$Message, [string]$Color="White")
    
    $consoleColor = "Gray"
    switch ($Color) {
        "Red" { $consoleColor = "Red" }
        "Green" { $consoleColor = "Green" }
        "Yellow" { $consoleColor = "Yellow" }
        "Cyan" { $consoleColor = "Cyan" }
        "Magenta" { $consoleColor = "Magenta" }
    }
    Write-Host $Message -ForegroundColor $consoleColor

    if ($global:IsGuiMode -and $global:GuiLogBox) {
        $prefix = ""
        if ($Color -eq "Red") { $prefix = "[ERROR] " }
        if ($Color -eq "Yellow") { $prefix = "[WARN] " }
        
        $global:GuiLogBox.AppendText("$prefix$Message`r`n")
        $global:GuiLogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents() 
    }
}

function Check-Cancel {
    [System.Windows.Forms.Application]::DoEvents()
    return $global:CancelRequested
}

function Create-ZipStub {
    param([string]$Path)
    # Hex Signature for an empty ZIP file (End of Central Directory Record)
    $bytes = [byte[]](0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

function Show-DecisionDialog {
    param([string]$Message, [string]$Title="Decision Needed", [bool]$ShowAllOptions=$true)

    $dForm = New-Object System.Windows.Forms.Form
    $dForm.Text = $Title
    $dForm.Width = 500
    $dForm.StartPosition = "CenterParent"
    $dForm.FormBorderStyle = "FixedDialog"
    $dForm.ControlBox = $false
    $dForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Message
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.AutoSize = $true
    $lbl.MaximumSize = New-Object System.Drawing.Size(440, 0) 
    $dForm.Controls.Add($lbl)

    $dForm.PerformLayout() 
    $buttonY = $lbl.Location.Y + $lbl.Height + 25
    $dForm.Height = $buttonY + 80

    $btnYes = New-Object System.Windows.Forms.Button
    $btnYes.Text = "Yes"
    $btnYes.Location = New-Object System.Drawing.Point(20, $buttonY)
    $btnYes.DialogResult = "Yes"
    $dForm.Controls.Add($btnYes)

    $btnNo = New-Object System.Windows.Forms.Button
    $btnNo.Text = "No"
    $btnNo.Location = New-Object System.Drawing.Point(100, $buttonY)
    $btnNo.DialogResult = "No"
    $dForm.Controls.Add($btnNo)

    if ($ShowAllOptions) {
        $btnYesAll = New-Object System.Windows.Forms.Button
        $btnYesAll.Text = "Yes to All"
        $btnYesAll.Location = New-Object System.Drawing.Point(180, $buttonY)
        $btnYesAll.Width = 80
        $dForm.Controls.Add($btnYesAll)
        $btnYesAll.Add_Click({ $dForm.Tag = "YesAll"; $dForm.Close() })

        $btnNoAll = New-Object System.Windows.Forms.Button
        $btnNoAll.Text = "No to All"
        $btnNoAll.Location = New-Object System.Drawing.Point(270, $buttonY)
        $btnNoAll.Width = 80
        $dForm.Controls.Add($btnNoAll)
        $btnNoAll.Add_Click({ $dForm.Tag = "NoAll"; $dForm.Close() })
    }

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(360, $buttonY)
    $btnCancel.DialogResult = "Cancel"
    $dForm.Controls.Add($btnCancel)

    $dForm.AcceptButton = $btnYes
    $dForm.CancelButton = $btnCancel

    $result = $dForm.ShowDialog()
    if ($dForm.Tag -ne $null) { return $dForm.Tag }
    return $result.ToString() 
}

function Show-TextWindow {
    param([string]$Content, [string]$Title="Viewer")
    
    $vForm = New-Object System.Windows.Forms.Form
    $vForm.Text = $Title
    $vForm.Size = New-Object System.Drawing.Size(700, 800)
    $vForm.StartPosition = "CenterParent"
    $vForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $pnl = New-Object System.Windows.Forms.Panel; $pnl.Dock = "Top"; $pnl.Height = 40; $vForm.Controls.Add($pnl)
    $lblS = New-Object System.Windows.Forms.Label; $lblS.Text = "Search:"; $lblS.Location = New-Object System.Drawing.Point(10, 10); $lblS.AutoSize = $true; $pnl.Controls.Add($lblS)
    $txtFind = New-Object System.Windows.Forms.TextBox; $txtFind.Location = New-Object System.Drawing.Point(60, 8); $txtFind.Size = New-Object System.Drawing.Size(400, 25); $pnl.Controls.Add($txtFind)
    $btnFind = New-Object System.Windows.Forms.Button; $btnFind.Text = "Find Next"; $btnFind.Location = New-Object System.Drawing.Point(470, 7); $btnFind.Size = New-Object System.Drawing.Size(150, 26); $pnl.Controls.Add($btnFind)

    $rtb = New-Object System.Windows.Forms.RichTextBox; $rtb.Dock = "Fill"; $rtb.ScrollBars = "Both"; $rtb.Font = New-Object System.Drawing.Font("Consolas", 10)
    $rtb.Text = $Content; $rtb.ReadOnly = $true; $rtb.BackColor = "White"; $rtb.WordWrap = $false; $vForm.Controls.Add($rtb)
    
    $doSearch = {
        $term = $txtFind.Text; if ([string]::IsNullOrWhiteSpace($term)) { return }
        $vForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $rtb.SelectAll(); $rtb.SelectionBackColor = [System.Drawing.Color]::White; $rtb.DeselectAll()
        $index = 0; $firstMatch = -1
        while ($index -lt $rtb.TextLength) {
            $found = $rtb.Find($term, $index, [System.Windows.Forms.RichTextBoxFinds]::None)
            if ($found -eq -1) { break }
            if ($firstMatch -eq -1) { $firstMatch = $found }
            $rtb.Select($found, $term.Length); $rtb.SelectionBackColor = [System.Drawing.Color]::Yellow
            $index = $found + $term.Length
        }
        if ($firstMatch -ne -1) { $rtb.Select($firstMatch, 0); $rtb.ScrollToCaret() } else { [System.Windows.Forms.MessageBox]::Show("Text not found.", "Search") }
        $vForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
    $btnFind.Add_Click($doSearch)
    $txtFind.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Enter') { & $doSearch; $e.SuppressKeyPress = $true } })
    $vForm.ShowDialog()
}

function Ask-UserAction-CLI {
    param([string]$Message, [string[]]$Options, [string]$PromptText)
    Log-Message "[!] WARNING: $Message" "Yellow"
    do {
        $choice = Read-Host $PromptText
        $choice = $choice.ToLower()
    } until ($choice -in $Options)
    if ($choice -eq 'c') { Log-Message "[-] Cancelled." "Red"; Exit }
    return $choice
}

function Handle-OrigConflict {
    param([string]$OrigPath, [ref]$OverwriteAllRef, [ref]$KeepOldAllRef)
    if ($KeepOldAllRef.Value) { return "KeepOld" }
    if ($OverwriteAllRef.Value) { return "Overwrite" }

    $baseName = Split-Path $OrigPath -Leaf
    $msg = "Backup file '$baseName' already exists.`nOverwrite old backup?`n(YES = Overwrite, NO = Keep old backup)"
    
    if ($global:IsGuiMode) {
        $res = Show-DecisionDialog $msg "Backup Conflict" $true
        switch ($res) {
            "Cancel" { return "Cancel" }
            "No"     { return "KeepOld" }
            "NoAll"  { $KeepOldAllRef.Value = $true; return "KeepOld" }
            "YesAll" { $OverwriteAllRef.Value = $true; return "Overwrite" }
            "Yes"    { return "Overwrite" }
        }
    } else {
        $act = Ask-UserAction-CLI $msg @('y','n','ya','na','c') "Overwrite? [y]es, [n]o, [ya]ll, [na]ll, [c]ancel"
        if ($act -eq 'c') { return "Cancel" }
        if ($act -eq 'n') { return "KeepOld" }
        if ($act -eq 'na') { $KeepOldAllRef.Value = $true; return "KeepOld" }
        if ($act -eq 'ya') { $OverwriteAllRef.Value = $true; return "Overwrite" }
        if ($act -eq 'y') { return "Overwrite" }
    }
    return "KeepOld"
}

# ------------------- Core Functions -------------------

Function UnpackVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Starting Unpack ---" "Cyan"
    
    if (-not (Test-Path -LiteralPath $addonPackages)) { Log-Message "Dir not found." "Red"; return }
    $varfiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File
    
    $overwriteFolderAll = $false; $skipFolderAll = $false
    $overwriteOrigAll = $false; $keepOrigAll = $false

    foreach ($varfile in $varfiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($varfile.Name)) { continue }

        $baseName = $varfile.Name
        $fullPath = $varfile.FullName
        $extractDir = Join-Path -Path $addonPackages -ChildPath $baseName
        $zipName    = $fullPath + ".zip"
        $origName   = $fullPath + ".orig"
        
        Log-Message "Processing: $baseName"

        if (Test-Path -LiteralPath $extractDir) {
            $skip = $false
            if ($skipFolderAll) { Log-Message " -> Skipped (Auto)." "Gray"; continue }
            
            if (-not $overwriteFolderAll) {
                if ($global:IsGuiMode) {
                    $res = Show-DecisionDialog "Folder '$baseName' exists.`nOverwrite?" "Conflict" $true
                    switch ($res) { "Cancel" {return} "No" {$skip=$true} "NoAll" {$skip=$true; $skipFolderAll=$true} "YesAll" {$overwriteFolderAll=$true} }
                } else {
                    $act = Ask-UserAction-CLI "Folder exists." @('o','s','c'); if ($act -eq 's') {$skip=$true}
                }
            }
            if ($skip) { Log-Message " -> Skipped." "Gray"; continue }
        }

        $doBackup = $true
        if (Test-Path -LiteralPath $origName) {
            $action = Handle-OrigConflict $origName ([ref]$overwriteOrigAll) ([ref]$keepOrigAll)
            if ($action -eq "Cancel") { return }
            if ($action -eq "KeepOld") { Log-Message " -> Keeping old backup." "Yellow"; $doBackup = $false }
            if ($action -eq "Overwrite") { Remove-Item -LiteralPath $origName -Force }
        }

        Try {
            Rename-Item -LiteralPath $fullPath -NewName ($varfile.Name + ".zip")
            Expand-Archive -LiteralPath $zipName -DestinationPath $extractDir -Force
            
            if (Check-Cancel) {
                Log-Message "  ! Cancelling... Rolling back." "Red"
                if (Test-Path -LiteralPath $extractDir) { Remove-Item -LiteralPath $extractDir -Recurse -Force }
                Rename-Item -LiteralPath $zipName -NewName $varfile.Name
                break
            }

            if ($doBackup) {
                Rename-Item -LiteralPath $zipName -NewName ($varfile.Name + ".orig")
                Log-Message " -> [OK] (Backup created)" "Green"
            } else {
                Rename-Item -LiteralPath $zipName -NewName $varfile.Name
                Log-Message " -> [OK] (Old backup kept)" "Green"
            }
        }
        Catch {
            Log-Message "Error: $($_.Exception.Message)" "Red"
            if (Test-Path -LiteralPath $zipName) { Rename-Item -LiteralPath $zipName -NewName $baseName -ErrorAction SilentlyContinue }
        }
    }
}

Function PackVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Starting Pack ---" "Cyan"
    $directories = Get-ChildItem -LiteralPath $addonPackages -Directory | Where-Object { $_.Name -like "*.var" }
    $overwriteOrigAll = $false; $keepOrigAll = $false

    foreach ($dir in $directories) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($dir.Name)) { continue }

        $dirPath = $dir.FullName; $dirName = $dir.Name
        $zipTarget = Join-Path -Path $addonPackages -ChildPath ($dirName + ".zip")
        $origBackup = Join-Path -Path $addonPackages -ChildPath ($dirName + ".orig")
        
        Log-Message "Packing: $dirName"
        
        if (Test-Path -LiteralPath $origBackup) {
            $action = Handle-OrigConflict $origBackup ([ref]$overwriteOrigAll) ([ref]$keepOrigAll)
            if ($action -eq "Cancel") { return }
            if ($action -eq "KeepOld") { Log-Message " -> Keeping old backup." "Yellow" }
            if ($action -eq "Overwrite") { Remove-Item -LiteralPath $origBackup -Force }
        }
        
        Try {
            $originalDate = $null
            if (Test-Path -LiteralPath $origBackup) { $originalDate = (Get-Item -LiteralPath $origBackup).LastWriteTime }

            $filesToCompress = Get-ChildItem -LiteralPath $dirPath -Recurse -Force
            if (-not $filesToCompress) { Log-Message "Empty dir." "Yellow"; continue }

            Compress-Archive -LiteralPath $dirPath -DestinationPath $zipTarget -Force
            
            if (Check-Cancel) {
                Log-Message "  ! Cancelling..." "Red"
                if (Test-Path -LiteralPath $zipTarget) { Remove-Item -LiteralPath $zipTarget -Force }
                break
            }

            if ($originalDate) { (Get-Item -LiteralPath $zipTarget).LastWriteTime = $originalDate }
            Remove-Item -LiteralPath $dirPath -Recurse -Force
            Rename-Item -LiteralPath $zipTarget -NewName $dirName
            Log-Message " -> [OK]" "Green"
        } Catch { Log-Message "Error: $($_.Exception.Message)" "Red" }
    }
}

Function SanitizeVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Starting Sanitize ---" "Cyan"
    $varfiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File
    $autoYesList = New-Object System.Collections.Generic.HashSet[string]
    $autoNoList  = New-Object System.Collections.Generic.HashSet[string]
    $overwriteOrigAll = $false; $keepOrigAll = $false

    foreach ($varfile in $varfiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($varfile.Name)) { continue }

        $baseName = $varfile.Name; $fullPath = $varfile.FullName; $tempZip = $fullPath + ".zip"; $origBackup = $fullPath + ".orig"
        Log-Message "Checking: $baseName"
        
        Try {
            $originalDate = $varfile.LastWriteTime
            Rename-Item -LiteralPath $fullPath -NewName ($baseName + ".zip")
            
            $archive = [System.IO.Compression.ZipFile]::OpenRead($tempZip)
            $foundEntries = $archive.Entries | Where-Object { $_.FullName -like "*.dll" -and $_.Name -notlike "_*" }
            
            $unpackRequired = $false
            if ($foundEntries) {
                 foreach ($entry in $foundEntries) {
                    if (Check-Cancel) { $archive.Dispose(); Rename-Item $tempZip $baseName; return }
                    $fName = $entry.Name
                    if ($autoNoList.Contains($fName)) { continue }
                    if ($autoYesList.Contains($fName)) { $unpackRequired = $true; continue }
                    
                    if ($global:IsGuiMode) {
                        $res = Show-DecisionDialog "Found: $fName`nRename this DLL?" "Sanitize" $true
                        switch ($res) { "Cancel" {$archive.Dispose(); Rename-Item $tempZip $baseName; return} "YesAll" {$autoYesList.Add($fName); $unpackRequired=$true} "Yes" {$unpackRequired=$true} "NoAll" {$autoNoList.Add($fName)} }
                    } else {
                        $act = Ask-UserAction-CLI "Rename $fName?" @('y','n','ya','na','c') "Rename?"; if ($act -eq 'ya' -or $act -eq 'y') {$unpackRequired=$true}
                    }
                 }
            }
            $archive.Dispose()

            if (-not $unpackRequired) {
                Rename-Item -LiteralPath $tempZip -NewName $baseName
                (Get-Item -LiteralPath $fullPath).LastWriteTime = $originalDate
                continue
            }

            # Logic Required -> Backup Check
            $makeBackup = $true
            if (Test-Path -LiteralPath $origBackup) {
                $action = Handle-OrigConflict $origBackup ([ref]$overwriteOrigAll) ([ref]$keepOrigAll)
                if ($action -eq "Cancel") { Rename-Item $tempZip $baseName; return }
                if ($action -eq "KeepOld") { $makeBackup = $false; Log-Message " -> Keeping old backup." "Yellow" }
                if ($action -eq "Overwrite") { Remove-Item -LiteralPath $origBackup -Force }
            }
            if ($makeBackup -and -not (Test-Path -LiteralPath $origBackup)) { Copy-Item -LiteralPath $tempZip -Destination $origBackup }

            Log-Message " -> Processing..." "Yellow"
            $extractDir = Join-Path -Path $addonPackages -ChildPath $baseName
            Expand-Archive -LiteralPath $tempZip -DestinationPath $extractDir -Force
            
            if (Check-Cancel) {
                Log-Message "  ! Cancelling..." "Red"
                if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
                Rename-Item $tempZip $baseName
                break
            }

            $dllFiles = Get-ChildItem -LiteralPath $extractDir -Recurse -Filter "*.dll" | Where-Object { $_.Name -notlike "_*" }
            foreach ($file in $dllFiles) { if (-not $autoNoList.Contains($file.Name)) { Rename-Item -LiteralPath $file.FullName -NewName ("_" + $file.Name) } }

            Compress-Archive -LiteralPath $extractDir -DestinationPath $tempZip -Force
            Remove-Item -LiteralPath $extractDir -Recurse -Force
            (Get-Item -LiteralPath $tempZip).LastWriteTime = $originalDate
            Rename-Item -LiteralPath $tempZip -NewName $baseName
            Log-Message " -> Done." "Green"
        } Catch { Log-Message "Error: $($_.Exception.Message)" "Red"; if (Test-Path $tempZip) { Rename-Item $tempZip $baseName -ErrorAction SilentlyContinue } }
    }
}

Function DesanitizeVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Starting Desanitize ---" "Cyan"
    $varfiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File
    $autoNoList  = New-Object System.Collections.Generic.HashSet[string]
    $overwriteOrigAll = $false; $keepOrigAll = $false

    foreach ($varfile in $varfiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($varfile.Name)) { continue }
        $baseName = $varfile.Name; $fullPath = $varfile.FullName; $tempZip = $fullPath + ".zip"; $origBackup = $fullPath + ".orig"
        
        Try {
            $originalDate = $varfile.LastWriteTime
            Rename-Item -LiteralPath $fullPath -NewName ($baseName + ".zip")
            
            $archive = [System.IO.Compression.ZipFile]::OpenRead($tempZip)
            $foundEntries = $archive.Entries | Where-Object { $_.Name -like "_*" -and $_.FullName -like "*_*.dll" }
            $hasEntries = ($foundEntries -ne $null)
            $archive.Dispose()
            
            if (-not $hasEntries) {
                Rename-Item -LiteralPath $tempZip -NewName $baseName
                (Get-Item -LiteralPath $fullPath).LastWriteTime = $originalDate
                continue
            }
            
            # Action Needed -> Backup Check
            $makeBackup = $true
            if (Test-Path -LiteralPath $origBackup) {
                $action = Handle-OrigConflict $origBackup ([ref]$overwriteOrigAll) ([ref]$keepOrigAll)
                if ($action -eq "Cancel") { Rename-Item $tempZip $baseName; return }
                if ($action -eq "KeepOld") { $makeBackup = $false; Log-Message " -> Keeping old backup." "Yellow" }
                if ($action -eq "Overwrite") { Remove-Item -LiteralPath $origBackup -Force }
            }
            if ($makeBackup -and -not (Test-Path -LiteralPath $origBackup)) { Copy-Item -LiteralPath $tempZip -Destination $origBackup }

            Log-Message "Restoring..." "Yellow"
            $extractDir = Join-Path -Path $addonPackages -ChildPath $baseName
            Expand-Archive -LiteralPath $tempZip -DestinationPath $extractDir -Force
            
            if (Check-Cancel) {
                Log-Message "  ! Cancelling..." "Red"
                if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
                Rename-Item $tempZip $baseName
                break
            }

            Get-ChildItem -LiteralPath $extractDir -Recurse -Filter "*.dll" | Where-Object { $_.Name -like "_*" } | ForEach-Object {
                if (-not $autoNoList.Contains($_.Name)) { Rename-Item -LiteralPath $_.FullName -NewName ($_.Name.Substring(1)) }
            }

            Compress-Archive -LiteralPath $extractDir -DestinationPath $tempZip -Force
            Remove-Item -LiteralPath $extractDir -Recurse -Force
            (Get-Item -LiteralPath $tempZip).LastWriteTime = $originalDate
            Rename-Item -LiteralPath $tempZip -NewName $baseName
            Log-Message " -> OK" "Green"
        } Catch { Log-Message "Error: $($_.Exception.Message)" "Red"; if (Test-Path $tempZip) { Rename-Item $tempZip $baseName -ErrorAction SilentlyContinue } }
    }
}

Function RepairVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Repair Mode ---" "Cyan"
    $varfiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File
    
    foreach ($varfile in $varfiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($varfile.Name)) { continue }
        
        $baseName = $varfile.Name; $tempZip = $varfile.FullName + ".zip"
        Try {
            Rename-Item -LiteralPath $varfile.FullName -NewName ($baseName + ".zip")
            $archive = [System.IO.Compression.ZipFile]::OpenRead($tempZip); $null = $archive.Entries; $archive.Dispose()
            Rename-Item -LiteralPath $tempZip -NewName $baseName
        } Catch {
            Log-Message "$baseName CORRUPT. Rescuing..." "Red"
            $rescueDir = Join-Path -Path $addonPackages -ChildPath ("RESCUE_" + $baseName)
            Try {
                Expand-Archive -LiteralPath $tempZip -DestinationPath $rescueDir -Force -ErrorAction Stop
                if (Check-Cancel) {
                     if (Test-Path $rescueDir) { Remove-Item $rescueDir -Recurse -Force }
                     Rename-Item $tempZip $baseName; break
                }
                Compress-Archive -LiteralPath $rescueDir -DestinationPath $tempZip -Force -ErrorAction Stop
                Remove-Item -LiteralPath $rescueDir -Recurse -Force
                Rename-Item -LiteralPath $tempZip -NewName $baseName
                Log-Message "Repaired!" "Green"
            } Catch {
                Log-Message "Rescue failed." "Red"
                if (Test-Path $rescueDir) { Remove-Item $rescueDir -Recurse -Force -ErrorAction SilentlyContinue }
                
                # STUBBING
                $createStub = $false; $deleteFile = $false
                if ($global:IsGuiMode) {
                    $r = Show-DecisionDialog "Repair failed.`nReplace with empty Stub?" "Stub" $false
                    if ($r -eq "Yes") { $createStub = $true } elseif ($r -eq "No") {
                        $d = Show-DecisionDialog "Delete corrupt file?" "Delete" $false
                        if ($d -eq "Yes") { $deleteFile = $true }
                    }
                } else {
                    $a = Ask-UserAction-CLI "Repair Failed. Stub?" @('y','n','c') "Stub?"; if ($a -eq 'y') { $createStub=$true }
                    else { $a2 = Ask-UserAction-CLI "Delete?" @('y','n') "Delete?"; if ($a2 -eq 'y') { $deleteFile=$true } }
                }

                if ($createStub) {
                    Remove-Item -LiteralPath $tempZip -Force
                    Create-ZipStub -Path (Join-Path $addonPackages $baseName)
                    Log-Message " -> Replaced with Stub." "Green"
                } elseif ($deleteFile) {
                    Remove-Item -LiteralPath $tempZip -Force
                    Log-Message " -> Deleted." "Red"
                } else {
                    Rename-Item -LiteralPath $tempZip -NewName $baseName
                    Log-Message " -> Skipped (Corrupt kept)." "Yellow"
                }
            }
        }
    }
}

Function UnifyVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Unifying ---" "Cyan"
    $MergedName = "anonymous.merged.1.var"; $mergedDir = Join-Path -Path $addonPackages -ChildPath $MergedName
    if (-not (Test-Path $mergedDir)) { New-Item -Path $mergedDir -ItemType Directory | Out-Null }
    
    $varfiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File | Where-Object { $_.Name -ne $MergedName }
    $overwriteOrigAll = $false; $keepOrigAll = $false

    foreach ($varfile in $varfiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($varfile.Name)) { continue }

        Log-Message "Merging $($varfile.Name)"
        $fullPath = $varfile.FullName; $tempZip = $fullPath + ".zip"; $origName = $fullPath + ".orig"

        $doRenameToOrig = $true
        if (Test-Path -LiteralPath $origName) {
            $action = Handle-OrigConflict $origName ([ref]$overwriteOrigAll) ([ref]$keepOrigAll)
            if ($action -eq "Cancel") { return }
            if ($action -eq "KeepOld") { $doRenameToOrig = $false }
            if ($action -eq "Overwrite") { Remove-Item -LiteralPath $origName -Force }
        }

        Try {
            Rename-Item -LiteralPath $fullPath -NewName ($varfile.Name + ".zip")
            Expand-Archive -LiteralPath $tempZip -DestinationPath $mergedDir -Force
            if (Check-Cancel) { Rename-Item $tempZip $varfile.Name; break }

            if ($doRenameToOrig) { Rename-Item $tempZip ($varfile.Name + ".orig") }
            else { Rename-Item $tempZip $varfile.Name }
        } Catch { Log-Message "Error: $($_.Exception.Message)" "Red"; if (Test-Path $tempZip) { Rename-Item $tempZip $varfile.Name -ErrorAction SilentlyContinue } }
    }
}

Function RestoreVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- Restore ---" "Cyan"
    $origFiles = Get-ChildItem -LiteralPath $addonPackages -Filter "*.orig" -File
    
    foreach ($orig in $origFiles) {
        if (Check-Cancel) { break }
        if ($TargetList -and -not $TargetList.Contains($orig.Name)) { continue }
        
        $targetName = $orig.Name -replace ".orig$", ""
        $targetDir = Join-Path -Path $addonPackages -ChildPath $targetName
        $targetFile = Join-Path -Path $addonPackages -ChildPath $targetName
        
        Log-Message "Restoring: $targetName"
        Try {
            if (Test-Path -LiteralPath $targetDir -PathType Container) { Remove-Item -LiteralPath $targetDir -Recurse -Force }
            if (Test-Path -LiteralPath $targetFile -PathType Leaf) { Remove-Item -LiteralPath $targetFile -Force }
            Rename-Item -LiteralPath $orig.FullName -NewName $targetName
            Log-Message "  -> Restored." "Green"
        } Catch { Log-Message "Error: $($_.Exception.Message)" "Red" }
    }
}

Function ViewVars {
    param([System.Collections.Generic.HashSet[string]]$TargetList = $null)
    Log-Message "--- View Mode ---" "Cyan"
    $items = @(Get-ChildItem -LiteralPath $addonPackages -Filter "*.var")
    $sb = [System.Text.StringBuilder]::new()
    $count = 0

    foreach ($item in $items) {
        if ($TargetList -and -not $TargetList.Contains($item.Name)) { continue }
        $count++
        $sb.AppendLine("======================================================================")
        $sb.AppendLine(" CONTENT OF: $($item.Name)")
        $sb.AppendLine("======================================================================")
        
        if ($item.PSIsContainer) {
            try {
                $files = Get-ChildItem -LiteralPath $item.FullName -Recurse
                foreach ($f in $files) { $rel = $f.FullName.Substring($item.FullName.Length + 1); $sb.AppendLine(" [DIR] $rel") }
            } catch { $sb.AppendLine("Error reading directory.") }
        } else {
            try {
                Rename-Item -LiteralPath $item.FullName -NewName ($item.Name + ".zip")
                $zipPath = $item.FullName + ".zip"
                $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
                foreach ($entry in $archive.Entries) { $sb.AppendLine(" [ZIP] $($entry.FullName)") }
                $archive.Dispose()
                Rename-Item -LiteralPath $zipPath -NewName $item.Name
            } catch {
                $sb.AppendLine("Error reading archive: $($_.Exception.Message)")
                if (Test-Path ($item.FullName + ".zip")) { Rename-Item ($item.FullName + ".zip") $item.Name -ErrorAction SilentlyContinue }
            }
        }
        $sb.AppendLine("")
    }
    if ($count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No files selected.", "Viewer") }
    else { Show-TextWindow $sb.ToString() "Viewer ($count items)" }
}

# ------------------- GUI Launcher -------------------

Function Show-GUI {
    $global:IsGuiMode = $true
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "VAM-VarHandler v$ScriptVersion"
    $form.Size = New-Object System.Drawing.Size(800, 750)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Location = New-Object System.Drawing.Point(10, 15); $lblPath.Size = New-Object System.Drawing.Size(600, 20); $lblPath.Text = "Path: $addonPackages"
    $form.Controls.Add($lblPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Change Folder"; $btnBrowse.Location = New-Object System.Drawing.Point(620, 10); $btnBrowse.Size = New-Object System.Drawing.Size(150, 25)
    $form.Controls.Add($btnBrowse)

    $cbMode = New-Object System.Windows.Forms.ComboBox
    $cbMode.Location = New-Object System.Drawing.Point(10, 45); $cbMode.Size = New-Object System.Drawing.Size(150, 25); $cbMode.DropDownStyle = "DropDownList"
    $cbMode.Items.AddRange(@("Unpack", "Pack", "Sanitize", "Desanitize", "Repair", "Unify", "Restore", "View"))
    $cbMode.SelectedIndex = 0
    $form.Controls.Add($cbMode)

    $btnHelp = New-Object System.Windows.Forms.Button
    $btnHelp.Text = "?"; $btnHelp.Location = New-Object System.Drawing.Point(165, 44); $btnHelp.Size = New-Object System.Drawing.Size(25, 25)
    $btnHelp.FlatStyle = "Popup"; $btnHelp.BackColor = "WhiteSmoke"; $form.Controls.Add($btnHelp)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(200, 45); $txtSearch.Size = New-Object System.Drawing.Size(570, 25)
    $form.Controls.Add($txtSearch)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(10, 75); $lblInfo.Size = New-Object System.Drawing.Size(760, 20)
    $lblInfo.Text = "Select a mode and the files to process."
    $form.Controls.Add($lblInfo)

    $checkList = New-Object System.Windows.Forms.CheckedListBox
    $checkList.Location = New-Object System.Drawing.Point(10, 100); $checkList.Size = New-Object System.Drawing.Size(760, 350); $checkList.CheckOnClick = $true
    $form.Controls.Add($checkList)

    $global:GuiLogBox = New-Object System.Windows.Forms.TextBox
    $global:GuiLogBox.Location = New-Object System.Drawing.Point(10, 460); $global:GuiLogBox.Size = New-Object System.Drawing.Size(760, 180)
    $global:GuiLogBox.Multiline = $true; $global:GuiLogBox.ScrollBars = "Vertical"; $global:GuiLogBox.ReadOnly = $true
    $global:GuiLogBox.BackColor = "Black"; $global:GuiLogBox.ForeColor = "LightGray"; $global:GuiLogBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $form.Controls.Add($global:GuiLogBox)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Start Operation"; $btnRun.Location = New-Object System.Drawing.Point(10, 650); $btnRun.Size = New-Object System.Drawing.Size(150, 30); $btnRun.BackColor = "LightBlue"
    $form.Controls.Add($btnRun)

    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Text = "STOP"; $btnStop.Location = New-Object System.Drawing.Point(170, 650); $btnStop.Size = New-Object System.Drawing.Size(100, 30); $btnStop.BackColor = "Salmon"; $btnStop.Enabled = $false 
    $form.Controls.Add($btnStop)
    
    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All Visible"; $btnSelectAll.Location = New-Object System.Drawing.Point(620, 650); $btnSelectAll.Size = New-Object System.Drawing.Size(150, 30)
    $form.Controls.Add($btnSelectAll)

    $script:AllItems = New-Object System.Collections.Generic.List[string]
    $script:SelectedSet = New-Object System.Collections.Generic.HashSet[string]
    
    $Action_LoadData = {
        $mode = $cbMode.SelectedItem; $lblInfo.Text = "Loading data for mode: $mode ..."; $form.Refresh()
        $script:AllItems.Clear(); $script:SelectedSet.Clear(); [string[]]$items = @()
        
        if ($mode -eq "Pack") {
             $items = @(Get-ChildItem -LiteralPath $addonPackages -Directory | Where-Object { $_.Name -like "*.var" } | Select-Object -ExpandProperty Name)
        } elseif ($mode -eq "Restore") {
             $items = @(Get-ChildItem -LiteralPath $addonPackages -Filter "*.orig" -File -Name)
        } elseif ($mode -eq "View") {
             $f = @(Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File -Name)
             $d = @(Get-ChildItem -LiteralPath $addonPackages -Directory | Where-Object { $_.Name -like "*.var" } | Select-Object -ExpandProperty Name)
             $items = $f + $d
        } else {
             $items = @(Get-ChildItem -LiteralPath $addonPackages -Filter "*.var" -File -Name)
        }
        if ($items.Count -gt 0) { $script:AllItems.AddRange($items) }
        & $Action_UpdateView ""
    }

    $Action_UpdateView = {
        param($filter); $checkList.BeginUpdate(); $checkList.Items.Clear()
        $limit = 100; if ($filter) { $limit = 500 }
        $count = 0
        foreach ($item in $script:AllItems) {
            if ($count -ge $limit) { break }
            if (-not $filter -or $item -like "*$filter*") {
                $isChecked = $script:SelectedSet.Contains($item); $checkList.Items.Add($item, $isChecked) | Out-Null; $count++
            }
        }
        $lblInfo.Text = "Loaded: $($script:AllItems.Count) | Visible: $count | Selected: $($script:SelectedSet.Count)"; $checkList.EndUpdate()
    }

    $cbMode.Add_SelectedIndexChanged({ & $Action_LoadData })
    $txtSearch.Add_TextChanged({ & $Action_UpdateView $txtSearch.Text })
    
    $checkList.Add_ItemCheck({
        param($s, $e); $val = $checkList.Items[$e.Index]
        if ($e.NewValue -eq 'Checked') { $script:SelectedSet.Add($val) | Out-Null } else { $script:SelectedSet.Remove($val) | Out-Null }
        $lblInfo.Text = "Loaded: $($script:AllItems.Count) | Visible: $($checkList.Items.Count) | Selected: $($script:SelectedSet.Count)"
    })

    $btnBrowse.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.SelectedPath = $addonPackages
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:addonPackages = $dialog.SelectedPath; $lblPath.Text = "Path: $script:addonPackages"; & $Action_LoadData
        }
    })

    $btnStop.Add_Click({
        $global:CancelRequested = $true; Log-Message "!!! STOP REQUESTED BY USER !!!" "Red"; $btnStop.Text = "Stopping..."; $btnStop.Enabled = $false
    })

    $btnRun.Add_Click({
        $mode = $cbMode.SelectedItem
        $btnRun.Enabled = $false; $btnStop.Enabled = $true; $btnStop.Text = "STOP"; $cbMode.Enabled = $false; $btnBrowse.Enabled = $false
        $global:CancelRequested = $false
        Log-Message "--- Starting $mode ---" "Magenta"
        
        $targets = $null; if ($script:SelectedSet.Count -gt 0) { $targets = $script:SelectedSet } else { Log-Message "No selection - Processing ALL." "Yellow" }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        
        switch ($mode) {
            "Unpack"     { UnpackVars $targets }
            "Pack"       { PackVars $targets }
            "Sanitize"   { SanitizeVars $targets }
            "Desanitize" { DesanitizeVars $targets }
            "Repair"     { RepairVars $targets }
            "Unify"      { UnifyVars $targets }
            "Restore"    { RestoreVars $targets }
            "View"       { ViewVars $targets }
        }
        
        if ($global:CancelRequested) {
            Log-Message "--- ABORTED ---" "Red"
            $btnStop.Text = "STOPPED"
            $btnStop.BackColor = "DarkRed"; $btnStop.ForeColor = "White"
            Start-Sleep -Seconds 1
        } else { Log-Message "--- Finished ---" "Magenta" }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled = $true; $btnStop.Enabled = $false; $btnStop.BackColor = "Salmon"; $btnStop.ForeColor = "Black"; $btnStop.Text = "STOP"
        $cbMode.Enabled = $true; $btnBrowse.Enabled = $true
        if ($mode -ne "View") { & $Action_LoadData }
    })
    
    $btnSelectAll.Add_Click({ $count = $checkList.Items.Count; for ($i=0; $i -lt $count; $i++) { $checkList.SetItemChecked($i, $true) } })
    
    $btnHelp.Add_Click({
        $currMode = $cbMode.SelectedItem; $helpText = ""
        switch ($currMode) {
            "Pack"       { $helpText = "Compresses folders back to .var files.`nRestores original timestamps from .orig files." }
            "Unpack"     { $helpText = "Extracts .var files.`nKeeps original as .orig." }
            "Sanitize"   { $helpText = "Disables DLLs inside .var files." }
            "Desanitize" { $helpText = "Restores disabled DLLs inside .var files." }
            "Repair"     { $helpText = "Extracts/Repacks corrupt archives." }
            "Unify"      { $helpText = "Merges content to 'anonymous.merged.1.var'." }
            "Restore"    { $helpText = "UNDO: Restores .orig files, deletes changes." }
            "View"       { $helpText = "Lists file content without extracting." }
        }
        [System.Windows.Forms.MessageBox]::Show($helpText, "Help: $currMode", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    
    $form.Add_Shown({ Log-Message "Welcome to VAM-VarHandler v$ScriptVersion" "Cyan"; & $Action_LoadData })
    $form.ShowDialog()
}

# ------------------- Entry Point -------------------

if (-not $PSScriptRoot) { $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
$addonPackages = Join-Path -Path $PSScriptRoot -ChildPath "AddonPackages"
if (-not $addonPackages) { $addonPackages = $PWD }

if ($args.Count -gt 0) {
    if (-not (Test-Path -LiteralPath $addonPackages)) { Write-Host "Error: Folder not found." -ForegroundColor Red; Exit }
    Log-Message "--- VAM-VarHandler v$ScriptVersion [CLI Mode] ---" "Cyan"
    $cmd = $args[0]
    switch ($cmd.ToLower()) {
        "pack"     { PackVars }
        "unpack"   { UnpackVars }
        "unify"    { UnifyVars }
        "sanitize" { SanitizeVars }
        "repair"   { RepairVars }
        "restore"  { RestoreVars }
        default    { Write-Host "Unknown Command." }
    }
} else {
    Show-GUI
}