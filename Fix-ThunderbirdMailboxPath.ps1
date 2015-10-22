#Initialize script variables:
[bool]$tRunning = $false      # Boolean - indicates if Thunderbird process is running.
[array]$tProcs = @()          # Array   - contains process objects for any running instances of Thunderbird.
[array]$tProfs = @()          # Array   - file system directory objects for any existing Thunderbird profiles.
[array]$newContent = @()      # Array   - Content of the modified prefs.js file
[bool]$hasPathPrefix = $false # Boolean - Indicates if the mailbox path prefix was discovered in the prefs.js file 

function reportError {
    param ($errText,$errCode)
    #What we really want here is a modal dialog that reports the error.
    write-host $errText
    Exit $errCode
}

$tProcs += get-process -name 'thunderbird' -ea SilentlyContinue
if ($tProcs.count -gt 0) {
    $tRunning = $true
    foreach ($t in $tProcs) {
        try {
            $t | Stop-Process -Force -Confirm:$false -ea SilentlyContinue
        } catch {
            [string]$errText = "Could not stop Thunderbird.  Please make sure that Thunderbird has been stopped and try again."
            [int32]$errCode = 100
            reportError $errText $errCode
        }
    }
}

try {
    [string]$tRoam = join-path -Path $env:APPDATA -ChildPath 'Thunderbird\Profiles' -Resolve -ea Stop
} catch {
    [string]$errText = "Could not find Thunderbird 'Profiles' directory: $tRoam. This script cannot continue."
    [int32]$errCode = 200
    reportError $errText $errCode
}

$tProfs = gci -Path $tRoam -att d -ea SilentlyContinue
if ($tProfs.count -gt 0) {
    foreach ($tProf in $tProfs) {
        try {
            $prefsJS = join-path -Path $tProf.FullName -ChildPath prefs.js -Resolve
        } catch {
            [string]$errText = "Could not find Thunderbird user preferences file at $prefsJS. This script cannot continue."
            [int32]$errCode = 202
            reportError $errText $errCode
        }

        $oldContent = get-content $prefsJS
        $newContent += $oldContent | % { 
            if ($_ -match '^user_pref\("mail\.server\.server[2-9]\.server_sub_directory"') {
                $hasPathPrefix = $true
            } else { $_ }
        }
        if ($hasPathPrefix) {
            $prefsJSBackup = join-path -Path $tProf.FullName -ChildPath prefs-backup.js
            cp -Path $prefsJS -Destination $prefsJSBackup -Force -Confirm:$false
            try {
                test-path -Path $prefsJSBackup -ea SilentlyContinue
            } catch {
                [string]$errText = "Could not backup Thunderbird user preferences file at $prefsJS. This script cannot continue."
                [int32]$errCode = 203
                reportError $errText $errCode
            }
            del $prefsJS -Force -Confirm:$false
            $newContent | Out-File -FilePath $prefsJS -Encoding utf8 -Append
        }
    }
} elseif ($tProfs.count -eq 0) {
    [string]$errText = "Thunderbird 'Profiles' directory at $tRoam is empty. This script cannot continue."
    [int32]$errCode = 201
    reportError $errText $errCode
}

#Now try to restart Thunderbird.