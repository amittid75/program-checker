Import-Module ActiveDirectory

# Get the current user's OU path (excluding DCs)
$currentOU = (Get-ADUser -Identity $env:USERNAME).DistinguishedName.Substring($currentUserDN.IndexOf("OU=")).Substring(0, $currentUserDN.IndexOf(",DC="))

# Create a new variable without "OU=Synced", "DC=name",D", and commas
$modifiedOU = $currentOU -replace "OU=Synced", "" -replace "DC=name",D", "" -replace ",", ""

# Get all installed programs (no exclusions)
$ComputerName = $env:COMPUTERNAME
$installedPrograms = Get-WmiObject -Class Win32_Product | Select-Object -Property Caption

# Append programs to the existing "progg.txt" file
$installedPrograms | Out-File -FilePath ".\progg.txt" -Append

# Construct the destination path with the modified OU
$destinationPath = "\\127.0.0.1\Archion\prog\$modifiedOU"

# Send the updated "progg.txt" file to the destination path
Copy-Item -Path ".\progg.txt" -Destination $destinationPath

# Display the OU path and confirmation message
Write-Host "Your current OU path (excluding DCs, synced, specified text, and commas): $modifiedOU"
Write-Host "File progg.txt (with appended programs) sent to $destinationPath successfully!"