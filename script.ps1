
$CLDIR = ".\Cover letter"
$RDIR = ".\Resume"
$Dest = ".\Package"
$TDIR = ".\Transcript"
$Archive = "Archive"

<# Files that need to exist before usage #>
$Rtmp = "$($RDIR)\R-Template"
$CLtmp = "$($CLDIR)\CL-Template"
$LOG = ".\log"
$Userfile = ".\username"

$CLsig = "CL"
$Rsig = "R"

# Checks if all the file paths are created, if not created it will initialize them
@($CLDIR, $RDIR, $Dest, $TDIR, "$($RDIR)\$($Archive)", "$($CLDIR)\$($Archive)") | ForEach-Object { if (!(Test-Path -Path $_)) { 
        try {
            New-Item -ItemType "directory" -Path $_ | Out-Null
            Write-Host "Created directory $_" -ForegroundColor green
        }
        catch {
            Write-Host "Failed to create directory $_" -ForegroundColor red
        }
    } }

# Check if csv log file exists, create them if not
@("$Rtmp.docx", "$CLtmp.docx", "$LOG.csv", "$Userfile.txt") | ForEach-Object { if (!(Test-Path -Path $_)) { 
        try {
            New-Item -ItemType "file" -Path $_ | Out-Null
            Write-Host "Created file $_" -ForegroundColor green
        }
        catch {
            Write-Host "Failed to create file $_" -ForegroundColor red
        }
    } }

:exitLabel while ($true) {
    # User Input
    $Company = (Read-Host "Enter the company name").TrimStart().TrimEnd().replace(' ', '_')
    $Position = (Read-Host "Enter the position to apply").TrimStart().TrimEnd().replace(' ', '_')  



    $err = $false
    # If username file is empty ask for username and write to file
    if ([String]::IsNullOrWhiteSpace((Get-content "$Userfile.txt"))) {
        $User = (Read-Host "Enter your name").TrimStart().TrimEnd().replace(' ', '_')
        Add-Content -Path "$Userfile.txt" -Value $User
        Set-ItemProperty -Path "$Userfile.txt" -Name IsReadOnly -Value $true 
    }
    else {
        $User = Get-Content -Path "$Userfile.txt" -TotalCount 1
    }
    
    $CLfilename = "$($CLDIR)\$($CLsig)-$($Position)-$($Company)"
    $Rfilename = "$($RDIR)\$($Rsig)-$($Position)-$($Company)"

    $CLNewfilename = "$($Dest)\$($User)-$($Position)-$($Company)-Coverletter"
    $RNewfilename = "$($Dest)\$($User)-$($Position)-$($Company)-Resume"
    $TNewfilename = "$($Dest)\$($User)-$($Position)-$($Company)-Transcript"

    if ((Read-Host "`nCreate a new resume and coverletter from template? [Y] / [N]") -eq 'Y') {
        Copy-Item -Path "$Rtmp.docx" -Destination "$Rfilename.docx" 
        ii "$Rfilename.docx" ; Write-Host "Created the initial resume file" -ForegroundColor green

        Copy-Item -Path "$CLtmp.docx" -Destination "$CLfilename.docx" 
        ii "$CLfilename.docx" ; Write-Host "Created the initial coverletter file" -ForegroundColor green

        Read-Host -Prompt "`nPress any key to continue once ready with the pdfs or CTRL+C to quit" 
    }

    Remove-Item -path "$($Dest)/*" -Include *.pdf -Exclude *"$($Position)-$($Company)"* 

    # Copy Cover letter
    try {
        Get-ChildItem "$CLfilename.pdf" -ErrorAction stop | Copy-Item -Destination "$CLNewfilename.pdf" ; ii "$CLNewfilename.pdf" ; Write-Host "Copied Coverletter file to destination" -ForegroundColor green
        Get-ChildItem $CLfilename* -ErrorAction stop | Move-Item -Destination "$($CLDIR)\$($Archive)" 
    }
    catch {
        Write-Host "Failed to copy $($CLfilename)" -ForegroundColor red
        $err = $true
    }

    # Copy Resume
    try {
        Get-ChildItem "$Rfilename.pdf" -ErrorAction stop | Copy-Item -Destination "$RNewfilename.pdf" ; ii "$RNewfilename.pdf" ; Write-Host "Copied Resume file to destination" -ForegroundColor green
        Get-ChildItem $Rfilename* | Move-Item -Destination "$($RDIR)\$($Archive)" 
    }
    catch {
        Write-Host "Failed to copy $($Rfilename)" -ForegroundColor red
        $err = $true
    }

    # Copy Current Transcript
    try {
        Get-ChildItem "$($TDIR)\Current_Transcript.pdf" -ErrorAction stop | Copy-Item -Destination "$TNewfilename.pdf" ; ii "$TNewfilename.pdf" ; Write-Host "Copied Transcript file to destination" -ForegroundColor green
    }
    catch {
        Write-Host "Failed to copy Current_Transcript.pdf" -ForegroundColor red
        $err = $true
    }

    $date = Get-Date -Format "MM/dd/yyyy"
    $time = Get-Date -Format "HH:mm K"

    # If log file is empty create header 
    if ([String]::IsNullOrWhiteSpace((Get-content "$LOG.csv")), !$err) {
        "{0}, {1}, {2}, {3}" -f "Company", "Position", "Submision Date mm/dd/yyyy", "Time" | add-content "$LOG.csv"
    } 
    "{0}, {1}, {2}, {3}" -f $Company, $Position, $date, $time | add-content "$LOG.csv"

    while ($true) {
        $in = Read-Host "`nSome files failed be copy, Try again? [Y] / [N]" 
        if ($err) { if ($in -eq 'Y') { break skipuserinput } 
        elseif ($in -eq 'N') {
            $in = Read-Host "`nStart a new application [Y] / [N]"
            if ($in -eq 'N') {
                break exitLabel
            }
            elseif ($in -eq "Y") {
                break 
            }
        }
	}
    }
}

# Print lenght of the csv file excluding header
Write-Host "`nNumber of postion applied to: $(((Get-Content "$LOG.csv").Length) - 1)"

Read-Host -Prompt "Press ENTER to exit program" 


