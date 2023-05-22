import subprocess

def set_office_trusted_locations():
    # PowerShell script
    script = r"""
        $officeVersions = @("15.0", "16.0")  # Add or modify versions as needed

        # Get the list of user IDs from HKEY_USERS
        $userIds = Get-ChildItem "Registry::HKEY_USERS" |
                Where-Object { $_.PSChildName -notlike ".DEFAULT" -and $_.PSChildName -notlike "S-1-5-18" -and $_.PSChildName -notlike "S-1-5-19" -and $_.PSChildName -notlike "S-1-5-20" -and $_.PSChildName -notlike "*_Classes" }

        # Loop through each user ID and set the registry entries
        foreach ($userId in $userIds) {
            foreach ($version in $officeVersions) {
                $regPath = "Registry::HKEY_USERS\$userId\Software\Policies\Microsoft\Office\$version\Common\Security\Trusted Locations"

                # Create the "LocationX" key
                ##$locationKey = New-Item -Path $regPath -Name "Location1" -Force

                # Set the "Path" value
                ##Set-ItemProperty -Path "$regPath\Location1" -Name "Path" -Value "\\network\location"

                # Set the "AllowSubfolders" value
                ##Set-ItemProperty -Path "$regPath\Location1" -Name "AllowSubfolders" -Value 1
            }
        }
    """

    # Execute PowerShell script using subprocess
    subprocess.call(["powershell", "-Command", script], shell=True)

# Call the function to set Office trusted locations
set_office_trusted_locations()
