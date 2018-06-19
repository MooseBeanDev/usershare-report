Import-Module ActiveDirectory

# This little section registers an external powershell script that I downloaded from the internet
# The external powershell script, Get-FolderSize.ps1 will check the directory size and give us a nice looking output
$env:path += ";c:\scripts\FolderSize\"
Unblock-File "c:\scripts\FolderSize\Get-FolderSize.ps1"
. "c:\scripts\FolderSize\Get-FolderSize.ps1"

$computers = @()
$computers += Get-AdComputer -Filter {Name -like "*-lt-*"}
$computers += Get-AdComputer -Filter {Name -like "*-pc-*"}

$pccount = $computers.Count
Write-Host "Total PCs to target: $pccount" -fore white -back black

$total = 0
$maxsize = 0;
$minthreshold = 10;
$counter = 0;
$pctargetcounter = 0;
$profilecounter = 0;
$average = 0;

$stopwatch = New-object -TypeName System.Diagnostics.Stopwatch
$stopwatch.Start();

$cursor = 3
$lock = 3
$ado = New-Object -ComObject ADODB.Connection
$recordset = New-Object -ComObject ADODB.Recordset

$ado.open("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=C:\Scripts\FolderSize\ProfileSize.accdb")

$query = "SELECT * FROM ProfileSize"

ForEach ($targetcomputer in $computers) {
    $pctargetcounter++
    $percentcomplete = [math]::Round($pctargetcounter * 100 / $pccount,2)
    
    $targetcomputername = $targetcomputer.Name
    Write-Host "Trying to grab info from $targetcomputername ($pctargetcounter / $pccount or %$percentcomplete. Elapsed:" + $stopwatch.Elapsed -fore yellow -back black
    if (Test-Connection -ComputerName $targetcomputername -BufferSize 16 -Count 1 -Quiet) {
        Write-Host "Grabbing info from $targetcomputername" -fore yellow -back black
        if (Test-Path -Path \\$targetcomputername\c$\Users\) {
            $userfolders = Get-ChildItem -Path \\$targetcomputername\c$\Users\
        }
        ForEach ($profile in $userfolders) {
                $profiletotal = 0;

                if (Test-Path -Path \\$targetcomputername\c$\users\$profile\Documents) {
                    $result = Get-FolderSize \\$targetcomputername\c$\users\$profile\Documents -RoboOnly | Select "TotalMBytes"
                    $profiletotal += [double]$result.TotalMBytes
                }

                if (Test-Path -Path \\$targetcomputername\c$\users\$profile\Favorites) {
                    $result = Get-FolderSize \\$targetcomputername\c$\users\$profile\Favorites -RoboOnly | Select "TotalMBytes"
                    $profiletotal += [double]$result.TotalMBytes
                }

                if (Test-Path -Path \\$targetcomputername\c$\users\$profile\Links) {
                    $result = Get-FolderSize \\$targetcomputername\c$\users\$profile\Links -RoboOnly | Select "TotalMBytes"
                    $profiletotal += [double]$result.TotalMBytes
                }

                if (Test-Path -Path \\$targetcomputername\c$\users\$profile\Pictures) {
                    $result = Get-FolderSize \\$targetcomputername\c$\users\$profile\Pictures -RoboOnly | Select "TotalMBytes"
                    $profiletotal += [double]$result.TotalMBytes
                }

                if ($profiletotal -gt $minthreshold) {
                    Write-Host "$profile total size is $profiletotal MB" -fore white -back black

                    $recordset.open($query,$ado,$cursor,$lock)

                    $recordset.AddNew()
                    $recordset.Fields.Item("ComputerName") = "$targetcomputername"
                    $recordset.Fields.Item("Profile") = "$profile"
                    $recordset.Fields.Item("SizeMB") = "$profiletotal"
                    $recordset.Update()
                    $recordset.Close()

                    $total += $profiletotal
                    $profilecounter++
                }
        }
        $counter++
    }
}

$ado.Close()
$stopwatch.Stop();

$date = Get-Date -Format FileDate
Start-Transcript -Path "C:\Scripts\FolderSize\Log$date.txt"

$datetime = Get-Date
$pcaverage = $total / $counter
$profileaverage = $total / $profilecounter
$pctotalcounter = $computers.Count

Write-Host ""
Write-Host "Script finished on $datetime" -fore white -back black
Write-Host $stopwatch.Elapsed -fore white -back black

Write-Host ""
Write-Host "Total PCs Targeted: $pctotalcounter" -ForegroundColor white -back black
Write-Host "PCs hit: $counter" -fore green -back black

Write-Host "Total size: $total MB" -fore green -back black

Write-Host "Profiles counted: $profilecounter" -fore green -back black

Write-Host "PC Average: $pcaverage MB" -fore green -back black
Write-Host "Profile Average: $profileaverage MB" -fore green -back black

Stop-Transcript
