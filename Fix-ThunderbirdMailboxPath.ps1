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

try {
    [string]$tRoam = join-path -Path $env:APPDATA -ChildPath 'Thunderbird\Profiles' -Resolve -ea Stop
} catch {
    [string]$errText = "Could not find Thunderbird 'Profiles' directory: $tRoam. This script cannot continue."
    [int32]$errCode = 200
    reportError $errText $errCode
}

$tProfs = gci -Path $tRoam -att d -ea Stop
if ($tProfs.count -gt 0) {
    foreach ($tProf in $tProfs) {
        try {
            $prefsJS = join-path -Path $tProf.FullName -ChildPath prefs.js -Resolve
        } catch {
            [string]$errText = "Could not find Thunderbird user preferences file at $prefsJS. This script cannot continue."
            [int32]$errCode = 202
            reportError $errText $errCode
        }
        #Capture the content of the original prefs.js file into memory:
        $oldContent = get-content $prefsJS
        #Check the file for the mailbox path prefix setting.  Save everything other than this line to a new object:
        $newContent += $oldContent | % { 
            if ($_ -match '^user_pref\("mail\.server\.server[2-9]\.server_sub_directory"') {
                $hasPathPrefix = $true
            } else { $_ }
        }
        #If we detected the mailbox path prefix, kill T-Bird, backup prefs.js, delete it, and write the new one.
        if ($hasPathPrefix) {
            #Kill T-bird if it is running:
            $tProcs += get-process -name 'thunderbird' -ea Stop
            if ($tProcs.count -gt 0) {
                $tRunning = $true
                foreach ($t in $tProcs) {
                    try {
                        $t | Stop-Process -Force -Confirm:$false -ea Stop
                    } catch {
                        [string]$errText = "Could not stop Thunderbird.  Please make sure that Thunderbird has been stopped and try again."
                        [int32]$errCode = 100
                        reportError $errText $errCode
                    }
                }
            }
            #Backup prefs.js:
            $prefsJSBackup = join-path -Path $tProf.FullName -ChildPath prefs-backup.js
            cp -Path $prefsJS -Destination $prefsJSBackup -Force -Confirm:$false
            try {
                test-path -Path $prefsJSBackup -ea Stop
            } catch {
                [string]$errText = "Could not backup Thunderbird user preferences file at $prefsJS. This script cannot continue."
                [int32]$errCode = 203
                reportError $errText $errCode
            }
            #Delete prefs.js
            del $prefsJS -Force -Confirm:$false
            #Write out a new prefs.js
            $newContent | Out-File -FilePath $prefsJS -Encoding utf8 -Append
        }
    }
} elseif ($tProfs.count -eq 0) {
    [string]$errText = "Thunderbird 'Profiles' directory at $tRoam is empty. This script cannot continue."
    [int32]$errCode = 201
    reportError $errText $errCode
}

#Now try to restart Thunderbird.
if ($tRunning) {
    $arch = (Get-WmiObject -Class Win32_OperatingSystem).OSArchitecture
    if ($arch -eq '64-bit') {
        $RootPath = ${env:ProgramFiles(x86)}
    } else {
        $RootPath = $env:ProgramFiles
    }
    try {
        $tPath = Join-Path -Path $RootPath -ChildPath '\Mozilla Thunderbird\thunderbird.exe' -Resolve -ea Stop
    } catch {
        [string]$errText = "Thunderbird Files updated successfully, but could not restart Thunderbird.  Please start up Thunderbird manually."
        [int32]$errCode = 300
        reportError 
    }
    Start-Process -FilePath $tPath
}
# SIG # Begin signature block
# MIIY5gYJKoZIhvcNAQcCoIIY1zCCGNMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUmLYKDhPwMukMOesMRgr4TOlf
# whGgghQPMIID7jCCA1egAwIBAgIQfpPr+3zGTlnqS5p31Ab8OzANBgkqhkiG9w0B
# AQUFADCBizELMAkGA1UEBhMCWkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIG
# A1UEBxMLRHVyYmFudmlsbGUxDzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhh
# d3RlIENlcnRpZmljYXRpb24xHzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcg
# Q0EwHhcNMTIxMjIxMDAwMDAwWhcNMjAxMjMwMjM1OTU5WjBeMQswCQYDVQQGEwJV
# UzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xMDAuBgNVBAMTJ1N5bWFu
# dGVjIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EgLSBHMjCCASIwDQYJKoZIhvcN
# AQEBBQADggEPADCCAQoCggEBALGss0lUS5ccEgrYJXmRIlcqb9y4JsRDc2vCvy5Q
# WvsUwnaOQwElQ7Sh4kX06Ld7w3TMIte0lAAC903tv7S3RCRrzV9FO9FEzkMScxeC
# i2m0K8uZHqxyGyZNcR+xMd37UWECU6aq9UksBXhFpS+JzueZ5/6M4lc/PcaS3Er4
# ezPkeQr78HWIQZz/xQNRmarXbJ+TaYdlKYOFwmAUxMjJOxTawIHwHw103pIiq8r3
# +3R8J+b3Sht/p8OeLa6K6qbmqicWfWH3mHERvOJQoUvlXfrlDqcsn6plINPYlujI
# fKVOSET/GeJEB5IL12iEgF1qeGRFzWBGflTBE3zFefHJwXECAwEAAaOB+jCB9zAd
# BgNVHQ4EFgQUX5r1blzMzHSa1N197z/b7EyALt0wMgYIKwYBBQUHAQEEJjAkMCIG
# CCsGAQUFBzABhhZodHRwOi8vb2NzcC50aGF3dGUuY29tMBIGA1UdEwEB/wQIMAYB
# Af8CAQAwPwYDVR0fBDgwNjA0oDKgMIYuaHR0cDovL2NybC50aGF3dGUuY29tL1Ro
# YXd0ZVRpbWVzdGFtcGluZ0NBLmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAOBgNV
# HQ8BAf8EBAMCAQYwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0y
# MDQ4LTEwDQYJKoZIhvcNAQEFBQADgYEAAwmbj3nvf1kwqu9otfrjCR27T4IGXTdf
# plKfFo3qHJIJRG71betYfDDo+WmNI3MLEm9Hqa45EfgqsZuwGsOO61mWAK3ODE2y
# 0DGmCFwqevzieh1XTKhlGOl5QGIllm7HxzdqgyEIjkHq3dlXPx13SYcqFgZepjhq
# IhKjURmDfrYwggSjMIIDi6ADAgECAhAOz/Q4yP6/NW4E2GqYGxpQMA0GCSqGSIb3
# DQEBBQUAMF4xCzAJBgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3Jh
# dGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBD
# QSAtIEcyMB4XDTEyMTAxODAwMDAwMFoXDTIwMTIyOTIzNTk1OVowYjELMAkGA1UE
# BhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMTQwMgYDVQQDEytT
# eW1hbnRlYyBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIFNpZ25lciAtIEc0MIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAomMLOUS4uyOnREm7Dv+h8GEKU5Ow
# mNutLA9KxW7/hjxTVQ8VzgQ/K/2plpbZvmF5C1vJTIZ25eBDSyKV7sIrQ8Gf2Gi0
# jkBP7oU4uRHFI/JkWPAVMm9OV6GuiKQC1yoezUvh3WPVF4kyW7BemVqonShQDhfu
# ltthO0VRHc8SVguSR/yrrvZmPUescHLnkudfzRC5xINklBm9JYDh6NIipdC6Anqh
# d5NbZcPuF3S8QYYq3AhMjJKMkS2ed0QfaNaodHfbDlsyi1aLM73ZY8hJnTrFxeoz
# C9Lxoxv0i77Zs1eLO94Ep3oisiSuLsdwxb5OgyYI+wu9qU+ZCOEQKHKqzQIDAQAB
# o4IBVzCCAVMwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAO
# BgNVHQ8BAf8EBAMCB4AwcwYIKwYBBQUHAQEEZzBlMCoGCCsGAQUFBzABhh5odHRw
# Oi8vdHMtb2NzcC53cy5zeW1hbnRlYy5jb20wNwYIKwYBBQUHMAKGK2h0dHA6Ly90
# cy1haWEud3Muc3ltYW50ZWMuY29tL3Rzcy1jYS1nMi5jZXIwPAYDVR0fBDUwMzAx
# oC+gLYYraHR0cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vdHNzLWNhLWcyLmNy
# bDAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMjAdBgNV
# HQ4EFgQURsZpow5KFB7VTNpSYxc/Xja8DeYwHwYDVR0jBBgwFoAUX5r1blzMzHSa
# 1N197z/b7EyALt0wDQYJKoZIhvcNAQEFBQADggEBAHg7tJEqAEzwj2IwN3ijhCcH
# bxiy3iXcoNSUA6qGTiWfmkADHN3O43nLIWgG2rYytG2/9CwmYzPkSWRtDebDZw73
# BaQ1bHyJFsbpst+y6d0gxnEPzZV03LZc3r03H0N45ni1zSgEIKOq8UvEiCmRDoDR
# EfzdXHZuT14ORUZBbg2w6jiasTraCXEQ/Bx5tIB7rGn0/Zy2DBYr8X9bCT2bW+IW
# yhOBbQAuOA2oKY8s4bL0WqkBrxWcLC9JG9siu8P+eJRRw4axgohd8D20UaF5Mysu
# e7ncIAkTcetqGVvP6KUwVyyJST+5z3/Jvz4iaGNTmr1pdKzFHTx/kuDDvBzYBHUw
# ggWDMIIEa6ADAgECAhBswRTjkn1j/N3jIBkUFIdjMA0GCSqGSIb3DQEBCwUAMHwx
# CzAJBgNVBAYTAlVTMQswCQYDVQQIEwJNSTESMBAGA1UEBxMJQW5uIEFyYm9yMRIw
# EAYDVQQKEwlJbnRlcm5ldDIxETAPBgNVBAsTCEluQ29tbW9uMSUwIwYDVQQDExxJ
# bkNvbW1vbiBSU0EgQ29kZSBTaWduaW5nIENBMB4XDTE1MDkxMjAwMDAwMFoXDTE2
# MDkxMTIzNTk1OVowgaIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQRDAUwNTQwNTELMAkG
# A1UECAwCVlQxEzARBgNVBAcMCkJ1cmxpbmd0b24xITAfBgNVBAkMGDg1IFNvdXRo
# IFByb3NwZWN0IFN0cmVldDEeMBwGA1UECgwVVW5pdmVyc2l0eSBvZiBWZXJtb250
# MR4wHAYDVQQDDBVVbml2ZXJzaXR5IG9mIFZlcm1vbnQwggEiMA0GCSqGSIb3DQEB
# AQUAA4IBDwAwggEKAoIBAQCm7FkT0jiP2Vil57acMaGej6XPnp0GD3K+kuvrh/vK
# TiIlweRiA0lEO1DcwAegc2YE5KC5R+OU2B9hWDqSwT686ytjuWLireceKmQ0jDvY
# 8XAo1DAIFaKhd6n8TWT2mJG37u0BOBSM7nXmCYx0SS5Ien3B4U4A29AyJCR+TZCJ
# aBYoK62gqXa0kQh+0TNWIS65+SQVUQm1I6nuHW+AtN7Mg8/FG60VHoAkXuz4+WsH
# Vu+OOihBKesw0YzeyAoQjYCLo6Mw1gFkBipp6I3qUxmpd52PUhIpRJOeA6TglS3V
# fEa/i9izhOv7Ro1J+1GD5smc3WGNZBEXK8mnkX9nkEAVAgMBAAGjggHYMIIB1DAf
# BgNVHSMEGDAWgBSuNSMX//8GPZxQ4IwkZTMecBCIojAdBgNVHQ4EFgQUADFudteh
# qMzvzjV8NcXVwUnoynIwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMGYGA1UdIARfMF0w
# WwYMKwYBBAGuIwEEAwIBMEswSQYIKwYBBQUHAgEWPWh0dHBzOi8vd3d3LmluY29t
# bW9uLm9yZy9jZXJ0L3JlcG9zaXRvcnkvY3BzX2NvZGVfc2lnbmluZy5wZGYwSQYD
# VR0fBEIwQDA+oDygOoY4aHR0cDovL2NybC5pbmNvbW1vbi1yc2Eub3JnL0luQ29t
# bW9uUlNBQ29kZVNpZ25pbmdDQS5jcmwwfgYIKwYBBQUHAQEEcjBwMEQGCCsGAQUF
# BzAChjhodHRwOi8vY3J0LmluY29tbW9uLXJzYS5vcmcvSW5Db21tb25SU0FDb2Rl
# U2lnbmluZ0NBLmNydDAoBggrBgEFBQcwAYYcaHR0cDovL29jc3AuaW5jb21tb24t
# cnNhLm9yZzAZBgNVHREEEjAQgQ5zYWEtYWRAdXZtLmVkdTANBgkqhkiG9w0BAQsF
# AAOCAQEAV/JhnPh/JVazACCiWtE+mtK2YgKTW2hA5d7tVmHRwK8m4lyOGmFUOz/7
# RQWyFmk/tqCehQgY8D/QiwHB/3mPJXbRx2uU04tBl26qw4vKAxdqwzOBvMbzr14d
# 1uCQGS8Jqyn0xjNm5G0EqbsanQupAhpsrbzktANQ1qI+ntrr0WnhG8/d6S4SfUW9
# 4d/EcEE5RGp5KlNgyjqafSaf4syFW/D50sOb1xozvk0rJ0KZb1qw5qfjvm7tRRRV
# 6YrhDltAFovUmhHeB6hD02nrPvKyVLgaXgiDfeJCLb+akOW5piqVsNTPP6qkp0Us
# j9+2w0n+i2ZXQ6c3N7ZJYRIXvXacozCCBeswggPToAMCAQICEGXh4uPV3lBFhfMm
# JIAF4tQwDQYJKoZIhvcNAQENBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpO
# ZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkgQ2l0eTEeMBwGA1UEChMVVGhlIFVT
# RVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVVU0VSVHJ1c3QgUlNBIENlcnRpZmlj
# YXRpb24gQXV0aG9yaXR5MB4XDTE0MDkxOTAwMDAwMFoXDTI0MDkxODIzNTk1OVow
# fDELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAk1JMRIwEAYDVQQHEwlBbm4gQXJib3Ix
# EjAQBgNVBAoTCUludGVybmV0MjERMA8GA1UECxMISW5Db21tb24xJTAjBgNVBAMT
# HEluQ29tbW9uIFJTQSBDb2RlIFNpZ25pbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQDAoC+LHnq7anWs+D7co7o5Isrzo3bkv30wJ+a605gyViNc
# BoaXDYDo7aKBNesL9l5+qT5oc/2d1Gd5zqrqaLcZ2xx2OlmHXV6Zx6GyuKmEcwzM
# q4dGHGrH7zklvqfd2iw1cDYdIi4gO93jHA4/NJ/lff5VgFsGfIJXhFXzOPvyDDap
# uV6yxYFHI30SgaDAASg+A/k4l6OtAvICaP3VAav11VFNUNMXIkblcxjgOuQ3d1HI
# nn1Sik+A3Ca5wEzK/FH6EAkRelcqc8TgISpswlS9HD6D+FupLPH623jP2YmabaP/
# Dac/fkxWI9YJvuGlHYsHxb/j31iq76SvgssF+AoJAgMBAAGjggFaMIIBVjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQUrjUjF///Bj2c
# UOCMJGUzHnAQiKIwDgYDVR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAw
# EwYDVR0lBAwwCgYIKwYBBQUHAwMwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FD
# ZXJ0aWZpY2F0aW9uQXV0aG9yaXR5LmNybDB2BggrBgEFBQcBAQRqMGgwPwYIKwYB
# BQUHMAKGM2h0dHA6Ly9jcnQudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FBZGRU
# cnVzdENBLmNydDAlBggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNlcnRydXN0LmNv
# bTANBgkqhkiG9w0BAQ0FAAOCAgEARiy2f2pOJWa9nGqmqtCevQ+uTjX88Dgnwced
# BMmCNNuG4RP3wZaNMEQT0jXtefdXXJOmEldtq3mXwSZk38lcy8M2om2TI6HbqjAC
# a+q4wIXWkqJBbK4MOWXFH0wQKnrEXjCcfUxyzhZ4s6tA/L4LmRYTmCD/srpz0bVU
# 3AuSX+mj05E+WPEop4WE+D35OLcnMcjFbst3KWN99xxaK40VHnX8EkcBkipQPDcu
# yt1hbOCDjHTq2Ay84R/SchN6WkVPGpW8y0mGc59lul1dlDmjVOynF9MRU5ACynTk
# dQ0JfKHOeVUuvQlo2Qzt52CTn3OZ1NtIZ0yrxm267pXKuK86UxI9aZrLkyO/BPO4
# 2itvAG/QMv7tzJkGns1hmi74OgZ3WUVk3SNTkixAqCbf7TSmecnrtyt0XB/P/xur
# cyFOIo5YRvTgVPc5lWn6PO9oKEdYtDyBsI5GAKVpmrUfdqojsl5GRYQQSnpO/hYB
# Wyv+LsuhdTvaA5vwIDM8WrAjgTFx2vGnQjg5dsQIeUOpTixMierCUzCh+bF47i73
# jX3qoiolCX7xLKSXTpWS2oy7HzgjDdlAsfTwnwton5YNTJxzg6NjrUjsUbEIORtJ
# B/eeld5EWbQgGfwaJb5NEOTonZckUtYS1VmaFugWUEuhSWodQIq7RA6FT/4AQ6qd
# j3yPbNExggRBMIIEPQIBATCBkDB8MQswCQYDVQQGEwJVUzELMAkGA1UECBMCTUkx
# EjAQBgNVBAcTCUFubiBBcmJvcjESMBAGA1UEChMJSW50ZXJuZXQyMREwDwYDVQQL
# EwhJbkNvbW1vbjElMCMGA1UEAxMcSW5Db21tb24gUlNBIENvZGUgU2lnbmluZyBD
# QQIQbMEU45J9Y/zd4yAZFBSHYzAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEK
# MAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUbZVQ2rycuniosoox
# vOPJhgEifAQwDQYJKoZIhvcNAQEBBQAEggEAH4S9YeeZQixI0RmJ0A1NPLp62bwK
# Qee3qnywEKAKZPaFz863lS2LG0J27cXGAjGqH5vUyf4xoikKOfqYjuNg5VIPBzGj
# 91ZTqiAgR11ZrXLLTffHzsJ7pWBfG7iXAwD0RQi3QHd7P1XBjA/Pq4QDzCP7PyHD
# urRdePdeF5Q59usu4yhvKI50p7cpEbRWCf9kVzuyBVA/TMOe8mRtS93UV/HBP1ZO
# 0XuLD+syxdOUAu/kyeXe6VvDKNZ4h5TJ5DvT+hjcoalEKBd1TUrk9L8E66+PsPHC
# alnbbAPHQ5SPrf4h5E2rvwE2zqyFR8054jPYPvGY6psWaL/lNhBLkger4qGCAgsw
# ggIHBgkqhkiG9w0BCQYxggH4MIIB9AIBATByMF4xCzAJBgNVBAYTAlVTMR0wGwYD
# VQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjEwMC4GA1UEAxMnU3ltYW50ZWMgVGlt
# ZSBTdGFtcGluZyBTZXJ2aWNlcyBDQSAtIEcyAhAOz/Q4yP6/NW4E2GqYGxpQMAkG
# BSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJ
# BTEPFw0xNTEwMjMxMjM1NDdaMCMGCSqGSIb3DQEJBDEWBBQaf8gKpd6FXFro/LDl
# kYOEgtfhJDANBgkqhkiG9w0BAQEFAASCAQBuZFvU4P2yfItkp9QGfjY3GLttKNR/
# betHz7wCJRiLVyB3PKaGlMD4PFE0/kM3UTrhSrv8I3ilE6aZ3eE5NSYFJggJ1ONd
# igtS0+evAdv4YzajvDoCxy92nYzsFHN0baWikQtv59atEXICs6D0AgJCDhDoWAv/
# J8NWNPifwb6AP1zjao4QJ9wm+AyMBDRz37WI/zRZkVXrCbtQONxzlxtes2QR7yH8
# 7HU4lUGi0MxxPT1bII4ZXCTUOTCqKj9KF9Dri2jYDuyXQBQzzO+E9jgsCywBLQ2j
# ScVY+Iodu8SKTRDeV1pZNceVxucoammYN2vwh7YoRjUkw4srcqzOnMey
# SIG # End signature block
