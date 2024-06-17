<#
.SYNOPSIS
    Extracts Package Data for Use in Auto Recipe Creation.
.DESCRIPTION
    Extracts Package Data for Use in Auto Recipe Creation.
.PARAMETER PackageLocation
    Name of Application as found in the Evergreen or Nevergreen Find- functions.
.REQUIRES PowerShell Version 5.0, Cloudpaging Studio and Completed Cloudpaging Package (.stp in a project directory)
 .EXAMPLE
    >GrabTokData.ps1 -PackageLocation C:\7-Zip
#>

Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$PackageLocation
   )

# Get the first .stp file in the current directory (or specify a path)
Set-Location $PackageLocation

$stpFile = Get-ChildItem -Filter *.stp | Select-Object -First 1
$ExtractDir = "C:\Users\Public\Extract"

# Check if an .stp file was found
if ($stpFile -ne $null) {
    # Copy the file and rename the copy to Extract.zip
    write-output "Found $stpFile"

    $zipFile = "Extract.zip"
    Copy-Item $stpFile.FullName -Destination $zipFile

    # Load the necessary assembly for extraction
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # Extract the zip file
    Expand-Archive -Path $zipFile -DestinationPath $ExtractDir

    Write-Output "Extracting Package Metadata."
} else {
    Write-Output "No .stp file found."
}

Write-Output "Extracted Package Metadata."

Set-Location $ExtractDir

$tokFile = Get-ChildItem -Filter *.tok | Select-Object -First 1

# Define the path to the executable
$exePath = "C:\Program Files\Numecent\Cloudpaging Studio\JukeboxStudio.exe"

# Define the arguments to pass
$arguments = " -d " + [char]34 + "$ExtractDir\$tokFile" + [char]34 + " -f C:\Users\Public\token_output.txt"

# Start the process
Start-Process -FilePath $exePath -ArgumentList $arguments

Write-Output "Generating Token Output"

Start-Sleep -Seconds 10

Copy-Item -Path "C:\Users\Public\token_output.txt" -Destination $PackageLocation

Write-Output "Copying Token Output"

Start-Sleep -Seconds 10

"Deleting Token Output from Source"

Remove-Item "C:\Users\Public\token_output.txt" -Force

Set-Location $PackageLocation

"Deleting Zip File"

Start-Sleep -Seconds 10

Remove-Item $zipFile -Force

Remove-Item "C:\Users\Public\Extract" -Force

Write-Output "Please copy package files (including token_output.txt to a machine with Office installed and run Recipe Creation Script."
