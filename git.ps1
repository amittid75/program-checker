Import-Module ActiveDirectory

# Get the current user's OU path (excluding unwanted strings and commas)
$modifiedOU = ((Get-ADUser -Identity $env:USERNAME).DistinguishedName -replace "OU=name", "" -replace "DC=name", "" -replace ",", "").Substring(3)  # Adjusted substring for cleaner extraction

# Get installed programs (more detailed properties, without publisher)
$installedPrograms = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate | Where-Object { $_.DisplayName -match '\w+' }

# Append programs to CSV file (replaces existing "progg.csv")
$installedPrograms | Export-Csv -Path ".\progg.csv" -NoTypeInformation

# Construct the destination path
$destinationPath = "\\127.0.0.1\$modifiedOU"

# Copy the updated CSV file to the destination
Copy-Item -Path ".\progg.csv" -Destination $destinationPath

# Display optional messages
Write-Host "Your current OU path (cleaned): $modifiedOU"
Write-Host "File progg.csv (with appended programs) sent to $destinationPath successfully!"
