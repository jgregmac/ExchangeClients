$tproc = @()
$tproc += get-process -name 'thunderbird'

if ($tproc.count -gt 0) {
    foreach ($t in $tproc) {
        try {
            $t | Stop-Process -Force -Confirm:$false -ea SilentlyContinue
        } catch {
            
            $errText = "Could not stop Thunderbird.  Please make sure that Thunderbird has been stopped and try again."
        }
    }
}
