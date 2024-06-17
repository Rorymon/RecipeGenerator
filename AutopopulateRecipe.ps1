<#
.SYNOPSIS
    Populates Recipe document with data extracted from Package.
.DESCRIPTION
    Populates Recipe document with data extracted from Package.
.PARAMETER documentPath
    Full path to the Recipe Template file on your local machine. This script uses the User Group Template but can work in other documents as long as the default values are the same as our template.
.PARAMETER tokFile
    Full path to the Tok file that was generated on the Packaging Machine.
.REQUIRES PowerShell Version 5.0, Microsoft Word, a Cloudpaging Recipe Template and a Cloudpaging .tok file
 .EXAMPLE
    >AutoPopulateRecipe.ps1 -documentPath C:\Users\Public\Documents\CloudpagingRecipeTemplate.docx -tokFile "C:\Users\Public\Chrome"
#>

Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$documentPath,

   [Parameter(Mandatory=$True,Position=1)]
   [string]$tokFile
   )

$TodaysDate = Get-Date -Format "D"
$PackagerName = [Environment]::UserName

Function Replace-WordText {
    param (
        [string]$documentPath,
        [string]$textToFind
    )

    # Enclose the placeholder text in angle brackets
    $placeholder = "<" + $textToFind + ">"
    
    # Read and parse the replacement value from token_output.txt
    $fileContent = Get-Content -Path $tokFile
    $docvar = $fileContent | Where-Object { $_ -match "$textToFind=" } | ForEach-Object { $_ -replace "$textToFind=", "" }

    # If no replacement is found, exit the function
    if (-not $docvar) {
        Write-Host "No replacement found for $placeholder"
        return
    }

    # The replacement text
    $textToReplace = $docvar

    Write-Host "Updating $placeholder to $textToReplace"

    # Create a new Word application object and ensure it's visible for debugging
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true

    try {
        # Open the document
        $document = $word.Documents.Open($documentPath)

        # Perform the find and replace operation across the whole document
        $searchRange = $document.Content
        $searchRange.Find.Execute($placeholder, $true, $true, $false, $false, $false, $true, 1, $false, $textToReplace, 2)

        # Save the changes
        $document.Save()
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        # Clean up
        $document.Close()
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}


Function Replace-UniqueWordText {
    param (
        [string]$documentPath,
        [string]$textToFind,
        [string]$textToReplace
    )

    # Create a new Word application object
    $word = New-Object -ComObject Word.Application

    try {
        # Make Word visible (optional, for debugging purposes)
        $word.Visible = $false

        # Open the document
        $document = $word.Documents.Open($documentPath)

        # Perform the find and replace operation across the whole document
        $searchRange = $document.Content
        $searchRange.Find.Execute($textToFind, $true, $true, $false, $false, $false, $true, 1, $false, $textToReplace, 2)

        # Save the changes
        $document.Save()
    }
    catch {
        Write-Host "An error occurred: $_"
    }
    finally {
        # Clean up
        $document.Close()
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

Replace-UniqueWordText -documentPath $documentPath -textToFind "<PackagerName>" -textToReplace "$PackagerName"
Replace-UniqueWordText -documentPath $documentPath -textToFind "<Date>" -textToReplace "$TodaysDate"
Replace-WordText -documentPath $documentPath -textToFind "ApplicationID"
Replace-WordText -documentPath $documentPath -textToFind "ProjectName"
Replace-WordText -documentPath $documentPath -textToFind "ProjectDescription"
Replace-WordText -documentPath $documentPath -textToFind "CommandLine"
Replace-WordText -documentPath $documentPath -textToFind "FolderExclusions"
Replace-WordText -documentPath $documentPath -textToFind "KeyExclusions"

Write-Output "The template has now been populated, be sure to complete the remaining fields and add any packaging and pre-packaging steps and screenshots."
