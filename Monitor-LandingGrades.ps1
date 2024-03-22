<#
.SYNOPSIS
    Monitors a directory for new DCS log files, then extracts and displays the landing grades for the pilot.

.DESCRIPTION
    This script monitors a specified directory for new DCS log files. 
    Once a new file is detected, the script processes the file and extracts the landing grades for the initiator pilot.
    The script then displays the landing grades and the corresponding LSOP phrases to the user.

    User can press the Escape key twice to exit the monitoring loop (or normally just stop the script with Ctrl+C).

.PARAMETER LandingLogFolder
    Specifies the folder path where the landing grade files are located. 
    The default value is "$($env:UserProfile)\Saved Games\DCS\Missions".

.PARAMETER ParseLogFile
    Specifies the path to a log file that should be parsed immediately. 
    If this parameter is provided, the script will only parse this file and then exit.

.PARAMETER IncludeFilesNewerThan
    Specifies the age of the files ([timespan]) to include at the start. 
    If this parameter is provided, the script will also parse files that are newer than the given value at the start.
#>
[CmdletBinding(DefaultParameterSetName='None')]
param (
    [string]$LandingLogFolder = "$($env:UserProfile)\Saved Games\DCS\Missions",
    [string]$ParseLogFile,
    [timespan]$IncludeFilesNewerThan
)

. "$PSScriptRoot\LSOPhraseParser.ps1"

<#
.SYNOPSIS
    Retrieves landing grades for the pilot from a log DCS file.

.DESCRIPTION
    The Get-LandingGrade function reads a DCS log file and extracts landing grades for the pilot. 
    It searches for sections in the log file that represent landing quality marks and retrieves the pilot's name and LSO phrase associated with each landing grade.

.PARAMETER filePath
    The path to the log file.

.OUTPUTS
    System.Collections.Generic.List[object]
    A list of landing grades for the specified pilot. Each landing grade is represented as a PSCustomObject with properties "Pilot" and "LSOPhrase". 

.NOTES
    This function assumes that the log file follows a specific format where 
        - The flying pilot name is stored in the log file as a key-value pair with the key "callsign".
        - Landing quality marks are stored in sections with the key "type" set to "landing quality mark".
        - Each landing quality mark section contains the pilot's name and the LSO phrase associated with the landing grade

    For example:
        debriefing = {
            ["callsign"] = "Wainamoinen",
            ["events"] = {
                ...
                [4] = {
                    ["type"] = "landing quality mark"
                    ["initiatorPilotName"] = "Wainamoinen"
                    ["comment"] = "LSO: GRADE:WO  (DLX)  _LULX_  _LULIM_  _LULIC_  WO(AFU)IC 
                }
                ...
                },
                [7] = {
                    ["type"] = "landing quality mark",
                    ["initiatorPilotName"] = "Wainamoinen",
                    ["comment"] = "LSO: GRADE:B  _FX_   _LOIM_ _LULIM_  LOIC  _LULIC_  WO(AFU)IC  (LLIW)  BIW ",		
                    ...
                },
                ...
            }
        }
#>
Function Get-LandingGrade([string]$filePath)
{
    if (-not (Test-Path $filePath)) {
        Write-Error "The file $filePath does not exist."
        return
    }

    # Read the log file
    $logLines = Get-Content -Path $filePath

    # Initialize variables to store the current section and the values to extract
    $currentLogSection = $null
    $callsign = $null
    $initiatorPilotName = $null
    $lsoPhrase = $null

    $landingGrades = New-Object System.Collections.Generic.List[object]

    # Iterate over the lines in the log file
    foreach ($line in $logLines) {

        # Try to extract the callsign from the current line
        if ($line -match "\[\`"callsign\`"\] = `"?(.+?)`"?,?\s*$") {
            $callsign = $matches[1]
        }

        # If the line starts a new section, update the current section
        if ($line -match "\[\d+\] = \{") {
            $currentLogSection = @{}
        }

        # If the line ends the current section, check if it's a "landing quality mark" section
        elseif ($null -ne $currentLogSection -and $line -match "\}," -and $currentLogSection["type"] -eq "landing quality mark") {
            $landingGrades.Add([PSCustomObject]@{
                Pilot = $initiatorPilotName
                LSOPhrase = $lsoPhrase
            })
        }
        # Try to extract the key and value from the line
        elseif ($null -ne $currentLogSection -and $line -match ".*\[\`"(.*)\`"\] = `"?(.+?)`"?,?\s*$") {
            $key = $matches[1]
            $value = $matches[2]
            $currentLogSection[$key] = $value
            if ($key -eq "initiatorPilotName") {
                $initiatorPilotName = $value
            }
            elseif ($key -eq "comment") {
                $lsoPhrase = $value
            }
        }
    }

    # Only return the landing grades for the pilot who was flying the aircraft, not for the other pilots in the mission
    return $landingGrades | Where-Object{ $_.Pilot -eq $callsign}
}

function Start-SleepButExitWithEsc($seconds) {
    $endTime = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $endTime) {
        if ([System.Console]::KeyAvailable) {
            $key = [System.Console]::ReadKey($true)
            if ($key.Key -eq "Escape") {
                # Ask the user to confirm the exit
                Write-Host ""
                Write-Host "Press Esc again to exit, or any other key to continue."
                while ($true) {
                    if ([System.Console]::KeyAvailable) {
                        $key = [System.Console]::ReadKey($true)
                        if ($key.Key -eq "Escape") {
                            return $true
                        } else {
                            Write-Host "Continuing monitoring..."
                            break
                        }
                    }
                    start-sleep -Milliseconds 100
                }
            }
        }
        Start-Sleep -Milliseconds 500
    }

    return $false
}

function Print-LandingGrades {
    param (
        [Parameter(Mandatory=$true)][array]$Grades,
        [Parameter(Mandatory=$true)][datetime]$gradeTime
    )

    if ($Grades.count -gt 0)
    {
        Write-Host "$gradeTime Landing grades for pilot: $($Grades[0].Pilot)"  -ForegroundColor Green
        Write-Host ""
        $i = 1
        foreach ($grade in $Grades) {
            if ($result.count -gt 1) {
                Write-Host "Landing #$($i): $($grade.LSOPhrase)" -ForegroundColor Magenta
            }

            try {
                $result = Get-LSOPhraseMeaning $grade.LSOPhrase
                Write-LSOPhraseToHost $result
            }
            catch {
                Write-Host "Error when handling LSO Phrase '$($grade.LSOPhrase)': $($_.Exception.Message)" -ForegroundColor Red
                Write-Host $_.Exception.StackTrace
            }

            Write-Host ""
            $i++
        }
    }
}

if ($ParseLogFile) {
    if (-not (Test-Path $ParseLogFile)) {
        Write-Host "The file $ParseLogFile does not exist."
        return
    }

    $logFile = Get-Item -Path $ParseLogFile
    $grades = Get-LandingGrade -filePath $ParseLogFile
    Print-LandingGrades -Grades $grades -gradeTime ($logFile.LastWriteTime)
    return
}

if (-not (Test-Path $LandingLogFolder)) {
    Write-Host "The directory $LandingLogFolder does not exist."
    exit
}

# Get the list of files that are initially in the directory
$existingFiles = Get-ChildItem -Path $LandingLogFolder -File -Recurse -Include *.log

if ($IncludeFilesNewerThan) {
    $cutoffTime = (Get-Date).Add(-$IncludeFilesNewerThan)
    $oldFiles = $existingFiles | Where-Object { $_.LastWriteTime -gt $cutoffTime } | Sort-Object LastWriteTime

    if ($oldFiles.count -gt 0) {
        Write-Host "Parsing files newer than $cutoffTime"
    
        foreach ($file in $oldFiles) {
            $grades = Get-LandingGrade -filePath $file.FullName
            if ($grades.count -gt 0)
            {
                Write-Host ""
                Print-LandingGrades -Grades $grades  ($file.LastWriteTime)
            }
        }
    }
}

$foundNewGradesAfterLastPing = $false
Write-Host "Monitoring for new landing grades.." -NoNewline
while ($true) {
    # Get the current list of files in the directory
    $currentFiles = Get-ChildItem -Path $landingLogFolder -File

    # Find any new files
    $newFiles = $currentFiles | Where-Object { $_.FullName -notin $existingFiles.FullName }

    # Process each new file
    foreach ($file in $newFiles) {
        $grades = Get-LandingGrade -filePath $file.FullName
        if ($grades.count -gt 0)
        {
            $foundNewGradesAfterLastPing = $true
            Write-Host ""
            Write-Host ""
            Print-LandingGrades -Grades $grades ($file.LastWriteTime)
        }
    }

    # Update the list of existing files
    $existingFiles = $currentFiles

    if ($foundNewGradesAfterLastPing -eq $true) {
        Write-Host "Monitoring for new landing grades..." -NoNewline
        $nextPingTime = (Get-Date).AddSeconds(15)
        $foundNewGradesAfterLastPing = $false
    } else {
        $now = Get-Date
        if ($nextPingTime -le $now) {
            Write-Host "." -NoNewline
            $nextPingTime = $now.AddSeconds(10)
        }
    }

    $shouldExit = Start-SleepButExitWithEsc -seconds 4
    if ($shouldExit -eq $true) {
        return
    }
}