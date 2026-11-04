param(
    [string]$ZipPath = "",
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string]$OutputPath = "",
    [string]$ModulePath = "",
    [int]$MaxEntries = 0
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName System.Web

if ([string]::IsNullOrWhiteSpace($ZipPath)) {
    $candidate = Get-ChildItem -LiteralPath (Join-Path $env:USERPROFILE "Downloads") -Filter "*_xml_*.zip" |
        Where-Object { $_.Name -like "*20260819.zip" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($candidate) {
        $ZipPath = $candidate.FullName
    }
}

if (-not (Test-Path -LiteralPath $ZipPath)) {
    throw "KRDict zip not found: $ZipPath"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $ProjectRoot "data\krdict_english_loanwords.tsv"
}

if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    $ModulePath = Join-Path $ProjectRoot "src\CompanyNameTools.bas"
}

function ConvertFrom-HtmlAttribute {
    param([string]$Value)
    return [System.Web.HttpUtility]::HtmlDecode([string]$Value).Trim()
}

function Test-HangulChar {
    param([char]$Char)
    $code = [int]$Char
    return (($code -ge 0xAC00 -and $code -le 0xD7A3) -or
            ($code -ge 0x1100 -and $code -le 0x11FF) -or
            ($code -ge 0x3130 -and $code -le 0x318F))
}

function Test-ContainsHangul {
    param([string]$Value)
    foreach ($char in $Value.ToCharArray()) {
        if (Test-HangulChar $char) { return $true }
    }
    return $false
}

function Test-HangulText {
    param(
        [string]$Value,
        [bool]$AllowSpace
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    foreach ($char in $Value.ToCharArray()) {
        if ($AllowSpace -and [char]::IsWhiteSpace($char)) { continue }
        if (-not (Test-HangulChar $char)) { return $false }
    }
    return $true
}

function Test-EnglishOrigin {
    param([string]$Origin)
    $value = ConvertFrom-HtmlAttribute $Origin
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    if ($value -notmatch "[A-Za-z]") { return $false }
    if (Test-ContainsHangul $value) { return $false }
    if ($value -match "[^\x00-\x7F]") { return $false }
    return $true
}

function Test-KoreanForm {
    param([string]$Form)
    $value = ConvertFrom-HtmlAttribute $Form
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    if (-not (Test-HangulText -Value $value -AllowSpace $true)) { return $false }
    if ($value.EndsWith(([string][char]0xD558 + [string][char]0xB2E4))) { return $false }
    if ($value.EndsWith(([string][char]0xB418 + [string][char]0xB2E4))) { return $false }
    if ($value.EndsWith(([string][char]0xC2DC + [string][char]0xD0A4 + [string][char]0xB2E4))) { return $false }
    return $true
}

function Get-OriginSegments {
    param([string]$Origin)
    $value = ConvertFrom-HtmlAttribute $Origin
    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($value -split "[,;]")) {
        $part = $part.Trim()
        if ($part.Length -gt 0) { [void]$segments.Add($part) }
    }
    if ($segments.Count -eq 0 -and $value.Length -gt 0) {
        [void]$segments.Add($value)
    }
    return $segments
}

function Get-EnglishTokens {
    param([string]$Text)
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($match in [regex]::Matches($Text.ToLowerInvariant(), "[a-z]+")) {
        [void]$tokens.Add($match.Value)
    }
    return $tokens
}

function Test-EnglishEquivalentMatch {
    param(
        [string]$Origin,
        [string]$EnglishEquivalent
    )

    if ([string]::IsNullOrWhiteSpace($EnglishEquivalent)) { return $false }

    $equivalentTokens = @(Get-EnglishTokens (ConvertFrom-HtmlAttribute $EnglishEquivalent))
    if ($equivalentTokens.Count -eq 0) { return $false }
    $equivalentSet = @{}
    foreach ($token in $equivalentTokens) { $equivalentSet[$token] = $true }
    $equivalentCompact = ($equivalentTokens -join "")

    foreach ($segment in (Get-OriginSegments $Origin)) {
        $originTokens = @(Get-EnglishTokens $segment)
        if ($originTokens.Count -eq 0) { continue }

        $originCompact = ($originTokens -join "")
        if ($equivalentCompact.Contains($originCompact)) { return $true }

        $allTokensFound = $true
        foreach ($token in $originTokens) {
            if (-not $equivalentSet.ContainsKey($token)) {
                $allTokensFound = $false
                break
            }
        }
        if ($allTokensFound) { return $true }
    }

    $originValue = (ConvertFrom-HtmlAttribute $Origin).Trim()
    if ($originValue -cmatch "^[A-Z]{2,}$") { return $true }

    return $false
}

function Add-Candidate {
    param(
        [hashtable]$Map,
        [string]$English,
        [string]$Korean,
        [string]$SourceForm,
        [string]$SourceOrigin
    )

    $key = $English.ToLowerInvariant().Trim()
    $value = $Korean.Trim()
    if ($key.Length -lt 2 -or $key.Length -gt 32) { return }
    if ($key -notmatch "^[a-z]+$") { return }
    if (-not (Test-HangulText -Value $value -AllowSpace $false)) { return }

    $row = [pscustomobject]@{
        English = $key
        Korean = $value
        SourceWrittenForm = (ConvertFrom-HtmlAttribute $SourceForm)
        SourceOrigin = (ConvertFrom-HtmlAttribute $SourceOrigin)
    }

    if (-not $Map.ContainsKey($key)) {
        $Map[$key] = $row
        return
    }

    $old = $Map[$key]
    if ($old.Korean -eq $value) { return }

    if ($value.Length -lt $old.Korean.Length) {
        $Map[$key] = $row
    }
}

function Add-EntryCandidates {
    param(
        [hashtable]$Map,
        [string]$WrittenForm,
        [string]$Origin,
        [string]$EnglishEquivalent
    )

    if (-not (Test-EnglishOrigin $Origin)) { return }
    if (-not (Test-KoreanForm $WrittenForm)) { return }
    if (-not (Test-EnglishEquivalentMatch -Origin $Origin -EnglishEquivalent $EnglishEquivalent)) { return }

    $form = ConvertFrom-HtmlAttribute $WrittenForm
    $formWords = @($form -split "\s+" | Where-Object { $_.Length -gt 0 })
    if ($formWords.Count -eq 0) { return }

    $formCompact = ($formWords -join "")
    foreach ($segment in (Get-OriginSegments $Origin)) {
        $tokens = @(Get-EnglishTokens $segment)
        if ($tokens.Count -eq 0) { continue }

        if ($tokens.Count -eq 1) {
            Add-Candidate -Map $Map -English $tokens[0] -Korean $formCompact -SourceForm $WrittenForm -SourceOrigin $Origin
            continue
        }

        $compactEnglish = ($tokens -join "")
        Add-Candidate -Map $Map -English $compactEnglish -Korean $formCompact -SourceForm $WrittenForm -SourceOrigin $Origin

        if ($formWords.Count -eq $tokens.Count) {
            for ($i = 0; $i -lt $tokens.Count; $i++) {
                Add-Candidate -Map $Map -English $tokens[$i] -Korean $formWords[$i] -SourceForm $WrittenForm -SourceOrigin $Origin
            }
        }
    }
}

function Format-VbaString {
    param([string]$Value)
    return '"' + $Value.Replace('"', '""') + '"'
}

function New-VbaDictionaryBlock {
    param([object[]]$Rows)

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("' BEGIN KRDICT LOANWORD DICTIONARY")
    [void]$lines.Add("' Source: National Institute of Korean Language Korean Basic Dictionary XML.")
    [void]$lines.Add("' Generated by scripts\update_krdict_loanword_dictionary.ps1.")
    [void]$lines.Add("Private Function CompanyName_KRDictLoanwordKorean(ByVal wordText As String) As String")
    [void]$lines.Add("    Select Case Left`$(LCase`$(wordText), 1)")
    for ($code = 97; $code -le 122; $code++) {
        $letter = [string][char]$code
        [void]$lines.Add("        Case $(Format-VbaString $letter): CompanyName_KRDictLoanwordKorean = CompanyName_KRDictLoanwordKorean_$($letter.ToUpperInvariant())(wordText)")
    }
    [void]$lines.Add("        Case Else: CompanyName_KRDictLoanwordKorean = wordText")
    [void]$lines.Add("    End Select")
    [void]$lines.Add("End Function")

    $rowsByLetter = @{}
    foreach ($row in $Rows) {
        $letter = $row.English.Substring(0, 1)
        if (-not $rowsByLetter.ContainsKey($letter)) {
            $rowsByLetter[$letter] = New-Object System.Collections.Generic.List[object]
        }
        [void]$rowsByLetter[$letter].Add($row)
    }

    for ($code = 97; $code -le 122; $code++) {
        $letter = [string][char]$code
        [void]$lines.Add("")
        [void]$lines.Add("Private Function CompanyName_KRDictLoanwordKorean_$($letter.ToUpperInvariant())(ByVal wordText As String) As String")
        [void]$lines.Add("    Select Case LCase`$(wordText)")
        if ($rowsByLetter.ContainsKey($letter)) {
            foreach ($row in @($rowsByLetter[$letter] | Sort-Object English)) {
                [void]$lines.Add("        Case $(Format-VbaString $row.English): CompanyName_KRDictLoanwordKorean_$($letter.ToUpperInvariant()) = $(Format-VbaString $row.Korean)")
            }
        }
        [void]$lines.Add("        Case Else: CompanyName_KRDictLoanwordKorean_$($letter.ToUpperInvariant()) = wordText")
        [void]$lines.Add("    End Select")
        [void]$lines.Add("End Function")
    }
    [void]$lines.Add("' END KRDICT LOANWORD DICTIONARY")
    return ($lines -join "`r`n")
}

$map = @{}
$writtenFormPattern = [regex]::new('<feat\s+att="writtenForm"\s+val="([^"]*)"')
$originPattern = [regex]::new('<feat\s+att="origin"\s+val="([^"]*)"')
$languagePattern = [regex]::new('<feat\s+att="language"\s+val="([^"]*)"')
$lemmaPattern = [regex]::new('<feat\s+att="lemma"\s+val="([^"]*)"')
$englishLanguageLabel = [string][char]0xC601 + [string][char]0xC5B4

$zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
try {
    foreach ($entry in $zip.Entries) {
        if (-not $entry.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $stream = $entry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
            $inEntry = $false
            $inEquivalent = $false
            $inEnglishEquivalent = $false
            $writtenForm = $null
            $origin = $null
            $englishEquivalent = New-Object System.Collections.Generic.List[string]

            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()

                if ($line.Contains("<LexicalEntry")) {
                    $inEntry = $true
                    $inEquivalent = $false
                    $inEnglishEquivalent = $false
                    $writtenForm = $null
                    $origin = $null
                    $englishEquivalent.Clear()
                    continue
                }

                if ($inEntry) {
                    if ($line.Contains("<Equivalent")) {
                        $inEquivalent = $true
                        $inEnglishEquivalent = $false
                    }

                    if ($inEquivalent) {
                        $match = $languagePattern.Match($line)
                        if ($match.Success -and (ConvertFrom-HtmlAttribute $match.Groups[1].Value) -eq $englishLanguageLabel) {
                            $inEnglishEquivalent = $true
                        }

                        if ($inEnglishEquivalent) {
                            $match = $lemmaPattern.Match($line)
                            if ($match.Success) {
                                [void]$englishEquivalent.Add((ConvertFrom-HtmlAttribute $match.Groups[1].Value))
                            }
                        }

                        if ($line.Contains("</Equivalent>")) {
                            $inEquivalent = $false
                            $inEnglishEquivalent = $false
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($writtenForm)) {
                        $match = $writtenFormPattern.Match($line)
                        if ($match.Success) {
                            $writtenForm = $match.Groups[1].Value
                        }
                    }

                    $match = $originPattern.Match($line)
                    if ($match.Success) {
                        $origin = $match.Groups[1].Value
                    }
                }

                if ($line.Contains("</LexicalEntry>")) {
                    if ($inEntry -and $writtenForm -and $origin) {
                        Add-EntryCandidates -Map $map -WrittenForm $writtenForm -Origin $origin -EnglishEquivalent ($englishEquivalent -join " ")
                    }
                    $inEntry = $false
                }
            }
            $reader.Dispose()
        } finally {
            $stream.Dispose()
        }
    }
} finally {
    $zip.Dispose()
}

$rows = @($map.Values | Sort-Object English)
if ($MaxEntries -gt 0 -and $rows.Count -gt $MaxEntries) {
    $rows = @($rows | Select-Object -First $MaxEntries)
}

$outputDir = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$tsv = New-Object System.Collections.Generic.List[string]
[void]$tsv.Add("# Source: Korean Basic Dictionary XML, National Institute of Korean Language")
[void]$tsv.Add("# Generated by scripts\update_krdict_loanword_dictionary.ps1")
[void]$tsv.Add("english`tkorean`tsourceWrittenForm`tsourceOrigin")
foreach ($row in $rows) {
    [void]$tsv.Add("$($row.English)`t$($row.Korean)`t$($row.SourceWrittenForm)`t$($row.SourceOrigin)")
}
Set-Content -LiteralPath $OutputPath -Encoding UTF8 -Value $tsv

$moduleText = Get-Content -LiteralPath $ModulePath -Raw -Encoding UTF8
$moduleText = $moduleText.Replace(
    "Case Else: CompanyName_EnglishWordKorean = wordText",
    "Case Else: CompanyName_EnglishWordKorean = CompanyName_KRDictLoanwordKorean(wordText)"
)

$block = New-VbaDictionaryBlock -Rows $rows
$pattern = "(?s)' BEGIN KRDICT LOANWORD DICTIONARY.*?' END KRDICT LOANWORD DICTIONARY"
if ([regex]::IsMatch($moduleText, $pattern)) {
    $moduleText = [regex]::Replace($moduleText, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $block })
} else {
    $insertBefore = "Private Function CompanyName_SpellAlphabetKorean"
    $index = $moduleText.IndexOf($insertBefore, [System.StringComparison]::Ordinal)
    if ($index -lt 0) { throw "Insert location not found: $insertBefore" }
    $moduleText = $moduleText.Substring(0, $index) + $block + "`r`n`r`n" + $moduleText.Substring($index)
}

$moduleText = $moduleText.TrimEnd() + "`r`n"
Set-Content -LiteralPath $ModulePath -Encoding UTF8 -Value $moduleText

Write-Output "KRDict English loanword entries: $($rows.Count)"
Write-Output "Dictionary TSV: $OutputPath"
Write-Output "Updated VBA module: $ModulePath"
