#requires -version 3.0
<#
    For SCCM 2012 and above:
    Creates an application, deployment type and optionally distributes the package to a named distribution point group
#>

<#
.SYNOPSIS
For SCCM 2012 and above: Creates an application, files it according to software vendor and package name, creates a deployment type and optionally distributes the application to a named distribution point group

.PARAMETER sitecode
The three letter SCCM site code

.PARAMETER msipath
The full UNC path to the MSI package to be distributed

.PARAMETER transform
Optional: The path to the MSI transform file to be applied. The path should be relative to the location of the MSI package

.PARAMETER installOptions
Optional: Any additional installation options to be passed to msiexec

.PARAMETER distribute
Should the package be distributed after creation. Must be specified with the -dpgroup parameter

.PARAMETER dpgroup
The distribution point group to distribute the package to. Must be combined with the -distribute switch parameter

.EXAMPLE
& .\SCCM-CreateApplicationFromMSI.ps1 -sitecode ORG -msipath '\\server\share\SomeVendor\SomePackage\1.0\VendorSomePackage-v1.0.msi' transforms='orgSettings.mst' -distribute -dpgroup 'US-SouthEast-DistributionPoints'
This will create an application for SomeVendor SomePackage v1.0, including creating a Vendor\Package folder structure under Applications, create the deployment type and apply the specified organisation settings transform file, and distribute the application to the US-SouthEast-DistributionPoints distribution point group.

#>

Param(
    [Parameter(Mandatory=$true,HelpMessage='Three letter SCCM site code')][string]$sitecode,
    [Parameter(Mandatory=$true)][string]$msipath,
    [string]$transform,
    [string]$installOptions,
    [switch]$distribute,
    [string]$dpgroup
)
function GetMSIProperties(){
    Param( [string]$msipath )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $installer.GetType().InvokeMember(
        "OpenDatabase",
        "InvokeMethod",
        $null,
        $installer,
        @($msipath, 0)
    )
    $view = $database.GetType().InvokeMember(
        "OpenView",
        "InvokeMethod",
        $null,
        $database,
        "SELECT * FROM Property"
    )
    $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) 
 
    $record = $View.GetType().InvokeMember( 
        "Fetch", 
        "InvokeMethod", 
        $Null, 
        $View, 
        $Null 
    ) 
 
    $msi_props = @{} 
    while ($record -ne $null) { 
        $msi_props[$record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)] = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 2) 
        $record = $View.GetType().InvokeMember( 
            "Fetch", 
            "InvokeMethod", 
            $Null, 
            $View, 
            $Null 
        ) 
    }

    Return $msi_props
}

if ((Test-Path -Path "Microsoft.PowerShell.Core\FileSystem::$msipath") -eq $false) {
    Throw 'Cannot find MSI installer'
}

if ($distribute -and $dpgroup -eq $null) {
    Throw 'To distribute the application, a distribution point group name must be specified using the -dpgroup parameter'
}

Import-Module ConfigurationManager
Set-Location -Path "$($sitecode):"

# Get MSI details
Write-Output 'Getting package details'
$msiprops = GetMSIProperties -msipath $msipath
$msifilename = Split-Path -Path "Microsoft.PowerShell.Core\FileSystem::$msipath" -Leaf
$packageName = ("$($msiprops.Manufacturer) $($msiprops.ProductName) $($msiprops.ProductVersion)")
$deployComment = ("$($env:USERNAME) - " + [System.DateTime]::Today.ToShortDateString()) 

# Create package
Write-Output "Creating package"
$package = New-CMApplication -Name $packageName -Description $deployComment -Publisher $msiprops.Manufacturer -SoftwareVersion $msiprops.ProductVersion

if ($transform -ne $null) {
    $installCommand = "msiexec.exe /i ""$msifilename"" transforms=""$transform"" /q $installOptions"
} else {
    $installCommand = "msiexec.exe /i ""$msifilename"" /q $installOptions"
}

# Adding MSI deployment type
Write-Output "Creating deployment type"
$deploymentType = Add-CMMsiDeploymentType -ContentLocation $msipath -InputObject $package -Comment $deployComment -DeploymentTypeName "Install MSI" -EstimatedRuntimeMins 15 -InstallCommand $installCommand -LogonRequirementType OnlyWhenNoUserLoggedOn -ProductCode $msiprops.ProductCode -UninstallCommand "msiexec.exe /x $($msiprops.ProductCode)"

# Setting deployment type properties
Set-CMDeploymentType -InputObject $deploymentType -InstallationBehaviorType InstallForSystem -MsiOrScriptInstaller -ProductCode $msiprops.ProductCode

# Moving package to correct path 
if ((Test-Path -path "$($sitecode):\Application\$($msiprops.Manufacturer)\$($msiprops.ProductName)") -eq $false) {
    if ((Test-Path -path "$($sitecode):\Application\$($msiprops.Manufacturer)") -eq $false) {
        Write-Output "Creating parent (manufacturer) folder"
        New-Item -Name $msiprops.Manufacturer -Path "AUS:\Application"
    }
    Write-Output "Creating child (product) folder"
    New-Item -Name $msiprops.ProductName -Path "$($sitecode):\Application\$($msiprops.Manufacturer)"
}
Move-CMObject -InputObject $package -FolderPath "$($sitecode):\Application\$($msiprops.Manufacturer)\$($msiprops.ProductName)"

# Get package details
Write-Output 'Getting package details'
$package = Get-CMApplication -Name $packageName

if ($package -eq $null) {
    Throw "No application of that name could be found!"
}

if ($distribute) {
    Write-Output 'Distributing package'
    Try {
        Start-CMContentDistribution -Application $package -DistributionPointGroupName $dpgroup
    } Catch {
        Write-Output 'Package is already distributed. Skipping...'
    }
}

Write-Host -ForegroundColor Yellow "All done!"


