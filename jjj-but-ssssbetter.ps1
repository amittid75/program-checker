# Get computer name and department
$computername = $env:COMPUTERNAME
$agaf = $computername.Substring(0, 3)
switch ($agaf) {
    "giz" { $result = "name1" }
    "giv" { $result = "name2" }
    "hnd" { $result = "name3" }
    "hin" { $result = "name4" }
    "roa" { $result = "name5" }
    "mok" { $result = "name6" }
    "bit" { $result = "name7" }
    "vet" { $result = "name8" }
    "sph" { $result = "name7" }
    "rev" { $result = "name8" }
    "din" { $result = "name9" }
    "mon" { $result = "name10" }
    default { $result = "diffrent" }
}

# Get program details without publisher
$installedPrograms = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate | Where-Object { $_.DisplayName -match '\w+' }

# Write program list to CSV file
$path = "\\127.0.0.1\Archion\prog\$result\prog.csv"  # Change file extension to .csv
$installedPrograms | Export-Csv -Path $path -NoTypeInformation -Append

# Optional filtering (uncomment below to only include programs installed after a specific date)
# $dateFilter = "2023-10-01"
# $installedPrograms = $installedPrograms | Where-Object { $_.InstallDate -ge $dateFilter }