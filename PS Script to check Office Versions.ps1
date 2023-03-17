# Define the function to get the Office version
function Get-OfficeVersion {
    # Get the registry key for Office
    $officeRegKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    # Get all subkeys under the office registry key
    $subKeys = Get-ChildItem $officeRegKey | Where-Object { $_.Name -match "Winword.exe" }
    # Loop through each subkey and get the version information
    foreach ($subKey in $subKeys) {
        # Get the path to winword.exe from the registry value
        $winwordPath = (Get-ItemProperty -Path "$($subKey.PSPath)\")."(Default)"
        # Check if winword.exe exists at that path
        if (Test-Path $winwordPath) {
            # Get the file version information for winword.exe
            $versionInfo = (Get-Item $winwordPath).VersionInfo.ProductVersion
            # Output the version information
            Write-Output "Office Version: $($versionInfo)"
        }
    }
}

# Define the function to get the Office 365 version
function Get-Office365Version {
    # Get all installed products from WMI class Win32_Product
    $products = gwmi Win32_Product | Where-Object { $_.Name -like "*Microsoft 365*" }
    foreach ($product in $products) {
        Write-Output "Office 365 Version: $($product.Version)"
    }
}

# Call the functions to get both Office and Office 365 versions
Get-OfficeVersion
Get-Office365Version