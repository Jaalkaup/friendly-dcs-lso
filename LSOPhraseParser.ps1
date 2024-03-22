
# All accents used in the DCS LSO phrase
$script:LSOAccentsMarks___ = @(
    [PSCustomObject]@{beginAccent = "`_"; endAccent = "`_"; meaning = "A lot"; color="Red"}
    [PSCustomObject]@{beginAccent = "("; endAccent = ")"; meaning = "A little"; color="Yellow"}
    [PSCustomObject]@{beginAccent = "["; endAccent = "]"; meaning = "Ignored signal LSO"; color="Red"}
)

# All distance marks used in the DCS LSO phrase
$script:LSODistanceMarks___ = @(
    [PSCustomObject]@{code = "BC"; meaning = "Ball call (before first 1/3 of glideslope)"} 
    [PSCustomObject]@{code = "X"; meaning = "At the start (first 1/3 of glideslope)"} 
    [PSCustomObject]@{code = "IM"; meaning = "In the middle (middle 1/3 of the glideslope)"}
    [PSCustomObject]@{code = "IC"; meaning = "In close (last 1/3 of glideslope)"}
    [PSCustomObject]@{code = "AR"; meaning = "At Ramp" }
    [PSCustomObject]@{code = "TL"; meaning = "To land" }
    [PSCustomObject]@{code = "IW"; meaning = "In the Wires" }
    [PSCustomObject]@{code = "AW"; meaning = "After wires"  }
)

# All landed grades used in the DCS LSO phrase
$script:LSOLandedGrades___ = @(
    [PSCustomObject]@{code = "_OK_"; meaning = "Perfect pass: rarely awarded"; points=5.0; color="Green"}
    [PSCustomObject]@{code = "OK"; meaning = "Pass with very minor deviations"; points=4.0; color="Green"}
    [PSCustomObject]@{code = "(OK)"; meaning = "Pass with one or more safe deviations"; points=3.0; color="Green"}
    [PSCustomObject]@{code = "---"; meaning = "No-grade. Below average but safe pass"; points=2.0}
    [PSCustomObject]@{code = "WO"; meaning = "Waveoff"; points=[Double]::NaN; color="Yellow"}
    [PSCustomObject]@{code = "OWO"; meaning = "Own Waveoff"; points=[Double]::NaN; color="Yellow"}
    [PSCustomObject]@{code = "C"; meaning = "Cut. Unsafe, gross deviations inside waveoff window"; points=0.0; color="Red"}
    [PSCustomObject]@{code = "B"; meaning = "Bolter: Safe pass where aircraft fails to hook a wire"; points=[Double]::NaN}
)

# All main errors used in the DCS LSO phrase
$script:LSOMainErrors___ = @(
    [PSCustomObject]@{code = "AFU"; meaning = "All 'fouled' up"}
    [PSCustomObject]@{code = "DL"; meaning = "Drifted left"}
    [PSCustomObject]@{code = "DR"; meaning = "Drifted right"}
    [PSCustomObject]@{code = "EG"; meaning = "Eased gun (did not advance throttles to MIL/AB after touchdown)"}
    [PSCustomObject]@{code = "F"; meaning = " Fast"}
    [PSCustomObject]@{code = "FD"; meaning = "Fouled deck"}
    [PSCustomObject]@{code = "H"; meaning = "High"}
    [PSCustomObject]@{code = "LL"; meaning = "Landed left"}
    [PSCustomObject]@{code = "LO"; meaning = "Low"}
    [PSCustomObject]@{code = "LR"; meaning = "Landed right"}
    [PSCustomObject]@{code = "LUL"; meaning = "Lined up left"}
    [PSCustomObject]@{code = "LUR"; meaning = "Lined up right"}
    [PSCustomObject]@{code = "N"; meaning = "Nose"}
    [PSCustomObject]@{code = "NERD"; meaning = "Not enough rate of descent"}
    [PSCustomObject]@{code = "NSU"; meaning = "Not set up"}
    [PSCustomObject]@{code = "P"; meaning = "Power"}
    [PSCustomObject]@{code = "SLO"; meaning = "Slow"}
    [PSCustomObject]@{code = "TMRD"; meaning = "Too much rate of descent"}
    [PSCustomObject]@{code = "LLWD"; meaning = "Landed left wing down"}
    [PSCustomObject]@{code = "LRWD"; meaning = "Landed right wing down"}
    [PSCustomObject]@{code = "LNF"; meaning = " Landed nose"}
    [PSCustomObject]@{code = "3PTS"; meaning = " Landed 3 points"; color="Green"}
)

Write-Host "Friendly DCS LSO"
Write-Host "================"

# Returns the worst color from a list of colors
function Get-WorstColor([string[]]$colors) {
    if ($colors -contains "Red") {
        return "Red"
    } elseif ($colors -contains "Yellow") {
        return "Yellow"
    } else {
        return "Gray"
    }
}

# Returns the worst accent color from a list of accents
function Get-WorstAccentColor($accents){
    $colors = $accents | ForEach-Object { $_.color }
    return Get-WorstColor $colors
}


<#
.SYNOPSIS
    Parses a grading word and extracts codes, meanings, and accents.

.DESCRIPTION
    The Get-LSOGradingWord function takes a grading word as input and parses it to extract codes, meanings, and accents.
    It iterates over each character in the grading word and checks if it matches any distance marks, grades, or main errors. 
    It also handles accents and ensures that they are properly opened and closed. The function returns an object containing the extracted codes, meanings, and accents.

.PARAMETER gradingWord
    The grading word to be parsed.

.EXAMPLE
    $gradingWord = "_DLIM_"
    $result = Get-LSOGradingWord -gradingWord $gradingWord
    $result.Codes
    $result.DistanceMark

    This example parses the grading word "_DLIM_" and stores the result in the $result variable. 
    The extracted codes are accessed using $result.Codes, and the distance mark (if present) is accessed using $result.DistanceMark.
#>
function Get-LSOGradingWord {
    param(
        [Parameter(Mandatory=$true)]
        [string]$gradingWord
    )

    $accentsStack = [System.Collections.Generic.Stack[PSCustomObject]]::new() # Stack to store opened accents
    $codes = @() # Stores the extracted codes
    $currentCodebuffer = "" # Stores the current code being parsed
    $foundDistanceMark = $null # Stores the found distance mark, null if none is found

    # Pre-filter the possible distance marks, grades, and main errors to speed up the search
    $possibleDistanceMarks = @($script:LSODistanceMarks___ | Where-Object { $gradingWord -match $_.code })
    $possibleGrades = @($script:LSOLandedGrades___ | Where-Object { $gradingWord -match $_.code })  
    $possibleMainErrors = @($script:LSOMainErrors___ | Where-Object { $gradingWord -match $_.code })

    # Iterate over each character in the grading word
    for ($i = 0; $i -lt $gradingWord.Length; $i++) {
        $char = $gradingWord[$i]

        # Check if the character is a beginning or ending accent
        $accent = $script:LSOAccentsMarks___ | Where-Object { $_.beginAccent -eq $char -or $_.endAccent -eq $char }
        if ($null -ne $accent) {
            if ($currentCodebuffer -ne "") {
                # Accent cannot be found in the middle of a code, so this would be somekind of syntax or parse error
                throw "Syntax error: Accent '$accent' found in the middle of a code: $currentCodebuffer"
            }

            if ($accentsStack.Count -gt 0 -and $accent.endAccent -eq $char) {
                # Ending accent found, check if it matches the last opened accent as it should be
                $lastOpenedAccent = $accentsStack.Peek()
                if ($lastOpenedAccent.meaning -ne $accent.meaning) {
                    throw "Mismatched accents: opened with $($lastOpenedAccent.meaning), but closed with $($accent.meaning)"
                }
                $accentsStack.Pop() | Out-Null
            } elseif ($accent.beginAccent -eq $char) {
                # Beginning accent found, add it to the stack
                $accentsStack.Push($accent)
            }
            continue
        }

        # Add the character to the buffer
        $currentCodebuffer += $char

        # Check if the buffer matches a distance mark, grade, or main error
        $possibleMatches = @(
            $possibleDistanceMarks | Where-Object { $_.code.StartsWith($currentCodebuffer) }
            $possibleGrades  | Where-Object { $_.code.StartsWith($currentCodebuffer) }
            $possibleMainErrors | Where-Object { $_.code.StartsWith($currentCodebuffer) }
        )

        
        if ($possibleMatches.Count -eq 1 -and $possibleMatches[0].code -eq $currentCodebuffer) {
            # The buffer matches a distance mark, grade, or main error exactly
            $code = $possibleMatches[0]
            
            if ($code -in $script:LSODistanceMarks___) {
                if ($null -ne $foundDistanceMark) {
                    # Only one distance mark is allowed per code
                    throw "Multiple distance marks found"
                }
            
                # Add the distance mark to the found distance mark with the accents
                $foundDistanceMark = [PSCustomObject]@{
                    Code = $code.code
                    Meaning = $code.meaning
                    Accents = if ($accentsStack.Count -gt 0) { $accentsStack.ToArray() } else { $null }
                }
            } elseif ($null -ne $foundDistanceMark) {
                # Distance mark must be the last code in the grading word
                # So if a distance mark is found but there were still codes after it, it's an error
                throw "Code found after distance mark"
            } else {
                # Add the code to the list of codes with the accents
                $codes += [PSCustomObject]@{
                    Code = $code.code
                    Meaning = $code.meaning
                    Accents = if ($accentsStack.Count -gt 0) { $accentsStack.ToArray() } else { $null }
                }
            }

            $currentCodebuffer = ""
        } elseif ($possibleMatches.Count -eq 0) {
            # There are no possible matches, so the code being parsed is invalid
            throw "Unrecognized code: $currentCodebuffer"
        }
    }

    # Return all codes and distance mark found
    return [PSCustomObject]@{ 
        Codes = $codes
        DistanceMark = $foundDistanceMark
    }
}

<#
.SYNOPSIS
    Get-LSOGrading function parses a sentence and returns parse LSO grading.

.DESCRIPTION
    The Get-LSOGrading function takes a sentence as input and splits it into individual words. 
    It then trims any empty words and parses each word using the Get-LSOGradingWord function. 
    The parsed words are then grouped based on their distance mark for easier access.

.PARAMETER sentence
    The sentence to be parsed and graded.

.OUTPUTS
    System.Object[]
    An array of grouped words based on distance marks.
#>
function Get-LSOGrading([string]$sentence) {
    # Split the sentence into words and trimg empty words out
    $words = $sentence -split " " | Where-Object { $_.Trim() -ne "" }

    # Parse each word
    $parsedWords = $words | ForEach-Object { Get-LSOGradingWord -gradingWord $_ }

    # Group the parsed words by distance mark for easier access
    $groupedWords = $parsedWords | Group-Object -Property {$_.DistanceMark.Code}
    return $groupedWords
}


<#
.SYNOPSIS
    Parses the LSO Grade text and extracts the meaning of the LSO phrase.

.DESCRIPTION
    This function takes an LSO Grade text as input and parses it to extract the meaning of the LSO phrase. 
    It checks the syntax of the LSO Grade text and extracts the landed grade, LSO grading, and wire number (if present).

.PARAMETER LSOText
    The LSO Grade text to be parsed.

.EXAMPLE
    $LSOText = "LSO: GRADE: 3.5: OK: 3-WIRE"
    $result = Get-LSOPhraseMeaning -LSOText $LSOText

    $result.LandedGrade    # Output: OK
    $result.LSOGrading     # Output: 3.5
    $result.Wire           # Output: 3

.NOTES
    - The LSO Grade text must start with "LSO: GRADE:".
    - The LSO Grade text can have the following format:
        - "LSO: GRADE: <LSO Grading>: <Landed Grade>: <Wire Number>"
        - "LSO: GRADE: <LSO Grading>: <Landed Grade>"
        - "LSO: GRADE: <LSO Grading>"
    - The Landed Grade must be a valid grade code.
    - The Wire Number must be a positive integer.
#>
function Get-LSOPhraseMeaning([string]$LSOText) {
    if (-not $LSOText.startsWith("LSO: GRADE:")) {
        throw "Invalid LSO Grade text: $LSOText"
    }

    if ($LSOText -match "WIRE#\s*(\d+)") {
        $wireNumber = [int]$matches[1]
        $LSOText = $LSOText -replace "WIRE#\s*\d+", ""
    } else {
        $wireNumber = $null
    }

    $sections = $LSOText -split ":"
    if ($sections.Count -lt 2) {
        throw "Invalid syntax in LSO Grade text: $LSOText"
    }

    $landedGrade = $null
    
    if ($sections.count -eq 4){
        $landedGradeStr = $sections[2].Trim()
        $landedGrade = $script:LSOLandedGrades___ | Where-Object { $_.code.StartsWith($landedGradeStr) }
        if ($null -eq $landedGrade) {
            throw "Invalid landed grade: $landedGradeStr"
        } elseif ($landedGrade.count -ne 1) {
            throw "Multiple landed grades found: $landedGradeStr"
        }

        $LSOGradingStr = $sections[3].Trim()
    } else {
        $LSOGradingStr = $sections[2].Trim()
    }

    $groupedLSOWords = Get-LSOGrading $LSOGradingStr

    return [PSCustomObject]@{
        LandedGrade = $landedGrade
        LSOGrading = $groupedLSOWords
        Wire = $wireNumber
    }
}

Function Write-LSOPhraseToHost([PSCustomObject]$LSOGrade) {
    if ($LSOGrade.LandedGrade -and $LSOGrade.Wire)
    {
        Write-Host "Landed Grade: $($LSOGrade.LandedGrade.Meaning)" -ForegroundColor $(if($LSOGrade.LandedGrade.color) { $LSOGrade.LandedGrade.color } else { "Gray" } )
        Write-Host "Wire: $($LSOGrade.Wire)" -ForegroundColor $(if ($LSOGrade.Wire -eq 1 ) { "Red" } elseif ($LSOGrade.Wire -eq 3) {  "Green" } elseif  ($LSOGrade.Wire -eq 5) { "Yellow" } else { "Gray"})
        Write-host ""
    } elseif ($LSOGrade.LandedGrade) {
        Write-Host "Landed Grade: $($LSOGrade.LandedGrade.Meaning)" -ForegroundColor $(if($LSOGrade.LandedGrade.color) { $LSOGrade.LandedGrade.color } else { "Gray" } )
        Write-host ""
    } else {
        Write-Host "Landed Grade: N/A"
        Write-host ""
    }

    $orderedGroups = $script:LSODistanceMarks___.Code | ForEach-Object {
        $distanceMarkCode = $_
        $group = $LSOGrade.LSOGrading | Where-Object { $_.Name -eq $distanceMarkCode }
        $group
    }

    $emptyGroup = $LSOGrade.LSOGrading | Where-Object { $_.Name -eq "" }
    $orderedGroups = @($emptyGroup) + $orderedGroups
    
    foreach($distanceMark in $orderedGroups) {
        if ($distanceMark.Name -eq "")
        {
            $GroupMeaning = "Global grade:"
        } else {
            $GroupMeaning = $distanceMark.Group[0].DistanceMark.Meaning
        }

        Write-Host "  => $GroupMeaning <="


        foreach($word in $distanceMark.Group) {
            $codeTexts = @()
            
            if ($word.Codes.Count -eq 0 -and $null -ne $word.DistanceMark.Accents) {
                $accentsText = ($word.DistanceMark.Accents |  ForEach-Object{ $_.meaning }) -join " "
                Write-Host "    $accentsText" -ForegroundColor $(Get-WorstAccentColor $word.DistanceMark.Accents)
                continue
            }

            $color = "Gray"
            foreach($code in $word.Codes) {
                if ($null -ne $code.Accents){
                    $accentsText = ($code.Accents |  ForEach-Object{ $_.meaning }) -join " "
                    $accentsText += " -> "
                    $worseAccentColor = Get-WorstAccentColor $code.Accents
                    $color =  Get-WorstColor @($color, $worseAccentColor)
                } else {
                    $accentsText = ""
                }
                $codeTexts += "$accentsText$($code.Meaning)"
                $color = if ($code.color) { Get-WorstColor $($code.color, $Color) } else { $color }
            }
            Write-Host "    $($codeTexts -join "; ")" -ForegroundColor $color
        } 
        Write-Host ""
    }
}
