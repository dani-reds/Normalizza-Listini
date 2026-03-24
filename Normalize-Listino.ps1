param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$Carrier,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Export', 'Import')]
    [string]$Direction,

    [Parameter(Mandatory = $false)]
    [string]$Reference = '',

    [Parameter(Mandatory = $false)]
    [string]$ValidityStartDate = '',

    [Parameter(Mandatory = $false)]
    [string]$RulesPath = '',

    [Parameter(Mandatory = $false)]
    [string]$UnlocodePath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:CurrentScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-TempDirectory {
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
    [System.IO.Directory]::CreateDirectory($path) | Out-Null
    return $path
}

function Remove-DirectorySafe {
    param([string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function Expand-XlsxPackage {
    param(
        [string]$Path,
        [string]$DestinationPath
    )

    [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $DestinationPath)
}

function Load-XmlDocument {
    param([string]$Path)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.Load($Path)
    return $doc
}

function Resolve-OpenXmlTargetPath {
    param(
        [string]$PackageRoot,
        [string]$Target
    )

    if (-not $Target) {
        throw 'OpenXML relationship target is empty.'
    }

    $normalized = $Target.Replace('/', '\').Trim()
    $normalized = $normalized.TrimStart('\')
    if ($normalized -like 'xl\*') {
        return (Join-Path $PackageRoot $normalized)
    }

    return (Join-Path $PackageRoot (Join-Path 'xl' $normalized))
}

function Get-FirstVisibleWorksheetPath {
    param([string]$PackageRoot)

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $pkgNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

    $workbookDoc = Load-XmlDocument (Join-Path $PackageRoot 'xl\workbook.xml')
    $workbookMgr = New-Object System.Xml.XmlNamespaceManager($workbookDoc.NameTable)
    $workbookMgr.AddNamespace('a', $mainNs)
    $workbookMgr.AddNamespace('r', $relNs)

    $relsDoc = Load-XmlDocument (Join-Path $PackageRoot 'xl\_rels\workbook.xml.rels')
    $relsMgr = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
    $relsMgr.AddNamespace('r', $pkgNs)

    $targetById = @{}
    foreach ($rel in $relsDoc.SelectNodes('//r:Relationship', $relsMgr)) {
        $targetById[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
    }

    foreach ($sheet in $workbookDoc.SelectNodes('//a:sheets/a:sheet', $workbookMgr)) {
        $state = $sheet.GetAttribute('state')
        if (-not $state -or $state -eq 'visible') {
            $relId = $sheet.GetAttribute('id', $relNs)
            $target = $targetById[$relId]
            if (-not $target) {
                throw "Worksheet relationship not found for $relId."
            }
            return (Resolve-OpenXmlTargetPath -PackageRoot $PackageRoot -Target $target)
        }
    }

    throw 'No visible worksheet found in workbook.'
}

function Get-WorksheetInfos {
    param([string]$PackageRoot)

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $pkgNs = 'http://schemas.openxmlformats.org/package/2006/relationships'

    $workbookDoc = Load-XmlDocument (Join-Path $PackageRoot 'xl\workbook.xml')
    $workbookMgr = New-Object System.Xml.XmlNamespaceManager($workbookDoc.NameTable)
    $workbookMgr.AddNamespace('a', $mainNs)
    $workbookMgr.AddNamespace('r', $relNs)

    $relsDoc = Load-XmlDocument (Join-Path $PackageRoot 'xl\_rels\workbook.xml.rels')
    $relsMgr = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
    $relsMgr.AddNamespace('r', $pkgNs)

    $targetById = @{}
    foreach ($rel in $relsDoc.SelectNodes('//r:Relationship', $relsMgr)) {
        $targetById[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target')
    }

    $infos = @()
    foreach ($sheet in $workbookDoc.SelectNodes('//a:sheets/a:sheet', $workbookMgr)) {
        $state = $sheet.GetAttribute('state')
        if (-not $state) {
            $state = 'visible'
        }

        $relId = $sheet.GetAttribute('id', $relNs)
        $target = $targetById[$relId]
        if (-not $target) {
            throw "Worksheet relationship not found for $relId."
        }

        $infos += [pscustomobject]@{
            Name = $sheet.GetAttribute('name')
            State = $state
            Path = (Resolve-OpenXmlTargetPath -PackageRoot $PackageRoot -Target $target)
        }
    }

    return $infos
}

function Get-SharedStrings {
    param([string]$PackageRoot)

    $path = Join-Path $PackageRoot 'xl\sharedStrings.xml'
    if (-not (Test-Path -LiteralPath $path)) {
        return @()
    }

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $doc = Load-XmlDocument $path
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $mgr.AddNamespace('a', $mainNs)

    $values = @()
    foreach ($si in $doc.SelectNodes('//a:si', $mgr)) {
        $parts = @()
        foreach ($t in $si.SelectNodes('.//a:t', $mgr)) {
            $parts += $t.InnerText
        }
        $values += ($parts -join '')
    }

    return $values
}

function Get-WorksheetRows {
    param(
        [string]$WorksheetPath,
        [string[]]$SharedStrings
    )

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $doc = Load-XmlDocument $WorksheetPath
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $mgr.AddNamespace('a', $mainNs)

    $rows = @()
    foreach ($rowNode in $doc.SelectNodes('//a:sheetData/a:row', $mgr)) {
        $cells = @{}
        foreach ($cell in $rowNode.SelectNodes('./a:c', $mgr)) {
            $ref = $cell.GetAttribute('r')
            $value = ''
            $inlineNode = $cell.SelectSingleNode('./a:is/a:t', $mgr)
            if ($inlineNode) {
                $value = $inlineNode.InnerText
            } else {
                $valueNode = $cell.SelectSingleNode('./a:v', $mgr)
                if ($valueNode) {
                    $raw = $valueNode.InnerText
                    switch ($cell.GetAttribute('t')) {
                        's' {
                            $index = [int]$raw
                            if ($index -ge 0 -and $index -lt $SharedStrings.Count) {
                                $value = $SharedStrings[$index]
                            } else {
                                $value = $raw
                            }
                        }
                        default {
                            $value = $raw
                        }
                    }
                }
            }
            $cells[$ref] = $value
        }

        $rows += [pscustomobject]@{
            RowNumber = [int]$rowNode.GetAttribute('r')
            Cells = $cells
        }
    }

    return $rows
}

function Get-Cell {
    param(
        [hashtable]$Cells,
        [string]$Column,
        [int]$RowNumber
    )

    $key = "$Column$RowNumber"
    if ($Cells.ContainsKey($key)) {
        return [string]$Cells[$key]
    }
    return ''
}

function Normalize-Whitespace {
    param([string]$Text)
    if ($null -eq $Text) {
        return ''
    }
    return (($Text -replace '\s+', ' ').Trim())
}

function Remove-Diacritics {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    $normalized = $Text.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($char in $normalized.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
        if ($category -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }
    return $builder.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Normalize-Key {
    param([string]$Text)
    $value = Normalize-Whitespace $Text
    $value = Remove-Diacritics $value
    return $value.ToUpperInvariant()
}

function Add-UniqueString {
    param(
        [System.Collections.Generic.List[string]]$List,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return
    }

    $normalized = Normalize-Whitespace $Value
    if (-not $List.Contains($normalized)) {
        $List.Add($normalized)
    }
}

function Normalize-LocationText {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    return $Text.Replace(([string][char]0xFF08), '(').Replace(([string][char]0xFF09), ')')
}

function Get-LocationIndexKeysForName {
    param([string]$Text)

    $keys = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    $sourceText = Normalize-LocationText $Text
    $normalized = Normalize-Key $sourceText
    Add-UniqueString -List $keys -Value $normalized

    $beforeParenthesis = Normalize-Key ($sourceText -replace '\(.*$', ' ')
    Add-UniqueString -List $keys -Value $beforeParenthesis

    $punctuationAsSpaces = Normalize-Key (($sourceText -replace '[\(\)\[\]（）]', ' ') -replace '[,;:/\\]+', ' ' -replace '[-_]+', ' ')
    Add-UniqueString -List $keys -Value $punctuationAsSpaces

    $compact = Normalize-Whitespace ($punctuationAsSpaces -replace '\s+', '')
    Add-UniqueString -List $keys -Value $compact

    $segments = @($sourceText -split '\s*,\s*' | Where-Object { $_ })
    if ($segments.Count -gt 0) {
        Add-UniqueString -List $keys -Value (Normalize-Key $segments[0])

        if ($segments.Count -gt 1) {
            Add-UniqueString -List $keys -Value (Normalize-Key (($segments[0..1]) -join ', '))
            Add-UniqueString -List $keys -Value (Normalize-Key (($segments[0..1]) -join ' '))
        }

        if ($segments[-1] -match '^[A-Za-z]{2}$' -and $segments.Count -gt 1) {
            $withoutCountry = @($segments[0..($segments.Count - 2)])
            Add-UniqueString -List $keys -Value (Normalize-Key ($withoutCountry -join ', '))
            Add-UniqueString -List $keys -Value (Normalize-Key ($withoutCountry -join ' '))
        }
    }

    $slashSegments = @($sourceText -split '\s*/\s*' | Where-Object { $_ })
    if ($slashSegments.Count -gt 1) {
        Add-UniqueString -List $keys -Value (Normalize-Key $slashSegments[0])
        Add-UniqueString -List $keys -Value (Normalize-Key (($slashSegments[0..1]) -join '/'))
        Add-UniqueString -List $keys -Value (Normalize-Key (($slashSegments[0..1]) -join ' '))
    }

    return $keys.ToArray()
}

function Get-LocationLookupCandidates {
    param([string]$RawName)

    $sourceText = Normalize-LocationText $RawName
    $cleanWithDetails = Normalize-Whitespace ((($sourceText -replace '[\(\)]', ' ') -replace '\*.*$', ' ') -replace '\s+', ' ')
    $clean = Normalize-Whitespace ((($sourceText -replace '\(.*?\)', ' ') -replace '\*.*$', ' ') -replace '\s+', ' ')
    $phrases = New-Object System.Collections.Generic.List[string]
    Add-UniqueString -List $phrases -Value $sourceText
    Add-UniqueString -List $phrases -Value $cleanWithDetails
    Add-UniqueString -List $phrases -Value $clean

    foreach ($match in [regex]::Matches($sourceText, '\((.*?)\)')) {
        $detail = Normalize-Whitespace $match.Groups[1].Value
        Add-UniqueString -List $phrases -Value $detail
        if ($clean) {
            Add-UniqueString -List $phrases -Value ($clean + ' ' + $detail)
        }
    }

    $segments = @($clean -split '\s*,\s*' | Where-Object { $_ })
    if ($segments.Count -gt 0) {
        Add-UniqueString -List $phrases -Value ($segments -join ' ')
        Add-UniqueString -List $phrases -Value $segments[0]

        if ($segments.Count -gt 1) {
            Add-UniqueString -List $phrases -Value (($segments[0..1]) -join ', ')
            Add-UniqueString -List $phrases -Value (($segments[0..1]) -join ' ')
        }

        if ($segments[-1] -match '^[A-Za-z]{2}$' -and $segments.Count -gt 1) {
            $withoutCountry = @($segments[0..($segments.Count - 2)])
            Add-UniqueString -List $phrases -Value ($withoutCountry -join ', ')
            Add-UniqueString -List $phrases -Value ($withoutCountry -join ' ')
        }
    }

    return $phrases.ToArray()
}

function Get-LocationCountryHint {
    param([string]$RawName)

    $sourceText = Normalize-LocationText $RawName
    $clean = Normalize-Whitespace ($sourceText -replace '\(.*?\)', '')
    $segments = @($clean -split '\s*,\s*' | Where-Object { $_ })
    if ($segments.Count -gt 0 -and $segments[-1] -match '^[A-Za-z]{2}$') {
        return (Normalize-Key $segments[-1])
    }

    return ''
}

function Resolve-UnlocodeLookupPath {
    param(
        [string]$ExplicitPath,
        [string]$InputPath
    )

    $candidates = New-Object System.Collections.Generic.List[string]

    if ($ExplicitPath) {
        if (-not (Test-Path -LiteralPath $ExplicitPath)) {
            throw "UNLOCODE lookup file not found: $ExplicitPath"
        }
        return $ExplicitPath
    }

    if ($env:UNLOCODE_LOOKUP_PATH) {
        Add-UniqueString -List $candidates -Value $env:UNLOCODE_LOOKUP_PATH
    }

    if ($script:CurrentScriptDirectory) {
        Add-UniqueString -List $candidates -Value (Join-Path $script:CurrentScriptDirectory 'UNLOCODE.txt')
    }

    if ($InputPath) {
        $inputDirectory = Split-Path -Path $InputPath -Parent
        if ($inputDirectory) {
            Add-UniqueString -List $candidates -Value (Join-Path $inputDirectory 'UNLOCODE.txt')
        }
    }

    if ($env:USERPROFILE) {
        Add-UniqueString -List $candidates -Value (Join-Path $env:USERPROFILE 'Desktop\n8n\n8n BA Extractor\UNLOCODE.txt')
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return ''
}

function Resolve-UnlocodeCachePath {
    param([string]$LookupPath)

    if (-not $LookupPath) {
        return ''
    }

    $fileName = ([System.IO.Path]::GetFileNameWithoutExtension($LookupPath) + '.lookup.clixml')
    if ($script:CurrentScriptDirectory) {
        return (Join-Path $script:CurrentScriptDirectory $fileName)
    }

    return (Join-Path (Split-Path -Path $LookupPath -Parent) $fileName)
}

function Import-UnlocodeLookup {
    param(
        [string]$Path,
        [string[]]$RawNames
    )

    if (-not $Path -or -not $RawNames -or $RawNames.Count -eq 0) {
        return $null
    }

    $cachePath = Resolve-UnlocodeCachePath -LookupPath $Path
    if ($cachePath -and (Test-Path -LiteralPath $cachePath)) {
        $cacheItem = Get-Item -LiteralPath $cachePath
        $sourceItem = Get-Item -LiteralPath $Path
        if ($cacheItem.LastWriteTimeUtc -ge $sourceItem.LastWriteTimeUtc) {
            return (Import-Clixml -LiteralPath $cachePath)
        }
    }

    $targetKeys = @{}
    foreach ($rawName in $RawNames) {
        foreach ($candidate in (Get-LocationLookupCandidates -RawName $rawName)) {
            foreach ($key in (Get-LocationIndexKeysForName -Text $candidate)) {
                $targetKeys[$key] = $true
            }
        }
    }

    if ($targetKeys.Count -eq 0) {
        return $null
    }

    $index = @{}
    $reader = [System.IO.File]::OpenText($Path)
    try {
        while (($line = $reader.ReadLine()) -ne $null) {
            $match = [regex]::Match($line, '^\s*(.+?)\s*:\s*([A-Z]{2}[A-Z0-9]{3})\s*$')
            if (-not $match.Success) {
                continue
            }

            $name = $match.Groups[1].Value
            $code = $match.Groups[2].Value.ToUpperInvariant()
            foreach ($key in (Get-LocationIndexKeysForName -Text $name)) {
                if (-not $targetKeys.ContainsKey($key)) {
                    continue
                }

                if (-not $index.ContainsKey($key)) {
                    $index[$key] = New-Object System.Collections.Generic.List[string]
                }
                if (-not $index[$key].Contains($code)) {
                    $index[$key].Add($code)
                }
            }
        }
    }
    finally {
        $reader.Dispose()
    }

    return $index
}

function Resolve-UnlocodesFromLookup {
    param(
        [string]$RawName,
        [hashtable]$Lookup,
        [string]$CountryHint = ''
    )

    if (-not $Lookup) {
        return @()
    }

    $countryHint = if ($CountryHint) { Normalize-Key $CountryHint } else { Get-LocationCountryHint -RawName $RawName }
    foreach ($candidate in (Get-LocationLookupCandidates -RawName $RawName)) {
        $candidateCodes = New-Object System.Collections.Generic.List[string]
        foreach ($key in (Get-LocationIndexKeysForName -Text $candidate)) {
            if (-not $Lookup.ContainsKey($key)) {
                continue
            }

            foreach ($code in $Lookup[$key]) {
                if ($countryHint -and (-not $code.StartsWith($countryHint))) {
                    continue
                }
                if (-not $candidateCodes.Contains($code)) {
                    $candidateCodes.Add($code)
                }
            }
        }

        if ($candidateCodes.Count -eq 1) {
            return $candidateCodes.ToArray()
        }

    }

    return @()
}

function Parse-ListinoDate {
    param([string]$Text)

    $formats = @('d/M/yy', 'd/M/yyyy', 'dd/MM/yy', 'dd/MM/yyyy', 'd MMM yyyy', 'dd MMM yyyy')
    foreach ($format in $formats) {
        try {
            return [DateTime]::ParseExact($Text, $format, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
        }
    }

    throw "Unable to parse date value '$Text'."
}

function Get-ValidityWindow {
    param([string]$HeaderText)

    $match = [regex]::Match($HeaderText, 'etd\s+(\d{1,2}/\d{1,2}/\d{2,4}).*?up to\s+(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'Unable to extract validity dates from header row.'
    }

    return [pscustomobject]@{
        Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
        End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
    }
}

function Format-OutputDateText {
    param(
        [string]$DateText,
        [string]$OutputFormat = 'dd/MM/yyyy'
    )

    if (-not $DateText) {
        return ''
    }

    $formats = @('yyyy-MM-dd', 'yyyy-M-d', 'dd/MM/yyyy', 'd/M/yyyy')
    foreach ($format in $formats) {
        try {
            return [DateTime]::ParseExact($DateText, $format, [System.Globalization.CultureInfo]::InvariantCulture).ToString($OutputFormat)
        } catch {
        }
    }

    throw "Unable to format output date '$DateText'."
}

function Resolve-PdfToTextPath {
    $commandNames = @('pdftotext.exe', 'pdftotext')
    foreach ($commandName in $commandNames) {
        try {
            $command = Get-Command $commandName -ErrorAction Stop | Select-Object -First 1
            if ($command.Source) {
                return $command.Source
            }
            if ($command.Path) {
                return $command.Path
            }
        } catch {
        }
    }

    $candidates = @(
        'C:\Program Files\Git\mingw64\bin\pdftotext.exe',
        'C:\Program Files\poppler\Library\bin\pdftotext.exe',
        'C:\Program Files\poppler\bin\pdftotext.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw 'pdftotext executable not found. Install Poppler or Git for Windows, or provide the tool in PATH.'
}

function Get-PdfText {
    param(
        [string]$InputPath,
        [ValidateSet('layout', 'table', 'raw', 'lineprinter')]
        [string]$Mode = 'layout'
    )

    $tempRoot = New-TempDirectory
    try {
        $textPath = Join-Path $tempRoot 'document.txt'
        $stdoutPath = Join-Path $tempRoot 'document.stdout.txt'
        $stderrPath = Join-Path $tempRoot 'document.stderr.txt'
        $pdfToTextPath = Resolve-PdfToTextPath
        $modeFlag = switch ($Mode) {
            'table' { '-table' }
            'raw' { '-raw' }
            'lineprinter' { '-lineprinter' }
            default { '-layout' }
        }
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $pdfToTextPath
        $startInfo.Arguments = ('{0} -enc UTF-8 "{1}" "{2}"' -f $modeFlag, $InputPath, $textPath)
        $startInfo.UseShellExecute = $false
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        [System.IO.File]::WriteAllText($stdoutPath, $stdout, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($stderrPath, $stderr, [System.Text.Encoding]::UTF8)
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $textPath)) {
            throw 'Failed to extract text from PDF with pdftotext.'
        }

        return [System.IO.File]::ReadAllText($textPath, [System.Text.Encoding]::UTF8)
    }
    finally {
        Remove-DirectorySafe $tempRoot
    }
}

function Convert-LocalizedNumberText {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    if (-not $value) {
        return ''
    }

    if ($value -match '^\d{1,3}(?:\.\d{3})+(?:,\d+)?$') {
        return ($value.Replace('.', '').Replace(',', '.'))
    }

    if ($value -match '^\d+,\d+$') {
        return ($value.Replace(',', '.'))
    }

    return ($value.Replace(',', '.'))
}

function Get-LineLeftSegment {
    param([string]$Line)

    $trimmed = ($Line | ForEach-Object { $_.TrimEnd() })
    if (-not $trimmed) {
        return ''
    }

    $parts = [regex]::Split($trimmed, '\s{2,}')
    if ($parts.Count -eq 0) {
        return ''
    }

    return Normalize-Whitespace $parts[0]
}

function Get-HapagQuotePages {
    param([string]$Text)

    $pages = @()
    foreach ($page in ($Text -split "`f")) {
        if (-not (Normalize-Whitespace $page)) {
            continue
        }

        if ($page -match 'Freight Charges' -and $page -match 'Lumpsum' -and $page -match 'Estimated Transportation Days') {
            $pages += $page
        }
    }

    return $pages
}

function Get-HapagRouteText {
    param([string[]]$Lines)

    $startIndex = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Lines[$i] -match '^\s*From\s+') {
            $startIndex = $i
            break
        }
    }

    if ($startIndex -lt 0) {
        throw 'Unable to locate route block in Hapag-Lloyd PDF page.'
    }

    $parts = New-Object System.Collections.Generic.List[string]
    for ($i = $startIndex; $i -lt $Lines.Count; $i++) {
        $left = Get-LineLeftSegment $Lines[$i]
        if (-not $left) {
            if ($parts.Count -gt 0) {
                break
            }
            continue
        }

        if ($parts.Count -gt 0 -and $left -match '^(Freight Charges|Lumpsum|Unless otherwise specified|Export Surcharges|Import Surcharges|Terminal Security Charge Orig\.)') {
            break
        }

        if ($parts.Count -eq 0 -and $left -notmatch '^From\s+') {
            continue
        }

        $parts.Add($left)
    }

    if ($parts.Count -eq 0) {
        throw 'Unable to compose route text from Hapag-Lloyd PDF page.'
    }

    return Normalize-Whitespace ($parts -join ' ')
}

function Parse-HapagRoute {
    param([string]$RouteText)

    $text = Normalize-Whitespace $RouteText
    if ($text -notmatch '^From\s+') {
        throw "Unsupported Hapag route text '$RouteText'."
    }

    $afterFrom = $text.Substring(5)
    $lastToIndex = $afterFrom.LastIndexOf(' to ')
    if ($lastToIndex -lt 0) {
        throw "Unable to split origin and destination from route text '$RouteText'."
    }

    $beforeDestination = $afterFrom.Substring(0, $lastToIndex)
    $destination = $afterFrom.Substring($lastToIndex + 4)
    $origin = ($beforeDestination -split '\s+via\s+', 2)[0]

    $origin = Normalize-Whitespace ($origin -replace '\s+\(.*$', '')
    $destination = Normalize-Whitespace ($destination -replace '\s+\(.*$', '')

    return [pscustomobject]@{
        Origin = $origin
        Destination = $destination
    }
}

function Get-HapagValidityWindow {
    param([string]$PageText)

    $match = [regex]::Match($PageText, '(\d{2} [A-Za-z]{3} \d{4})\s+(\d{2} [A-Za-z]{3} \d{4})')
    if (-not $match.Success) {
        throw 'Unable to extract validity dates from Hapag-Lloyd PDF page.'
    }

    return [pscustomobject]@{
        Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
        End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
    }
}

function Get-HapagTransitTime {
    param([string]$PageText)

    $lines = @($PageText -split "`r?`n")
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -notmatch 'Estimated Transportation Days') {
            continue
        }

        $limit = [Math]::Min($i + 6, $lines.Count - 1)
        for ($j = $i + 1; $j -le $limit; $j++) {
            $candidate = ($lines[$j]).Trim()
            if ($candidate -match '^\d+$') {
                return $candidate
            }
        }
    }

    return ''
}

function Get-HapagReference {
    param([string]$Text)

    $match = [regex]::Match($Text, 'QUOTAZIONE\s+Nr\.\:\s*(Q[0-9A-Z]+)', 'IgnoreCase')
    if (-not $match.Success) {
        return ''
    }

    return $match.Groups[1].Value.ToUpperInvariant()
}

function Get-HapagOceanFreightDetails {
    param([string]$PageText)

    $match = [regex]::Match($PageText, 'Lumpsum\s+USD\s+(-?\d+(?:[.,]\d+)?)\s+(-?\d+(?:[.,]\d+)?)', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'Unable to extract Lumpsum values from Hapag-Lloyd PDF page.'
    }

    $price20 = $match.Groups[1].Value.Replace(',', '.')
    $price40 = $match.Groups[2].Value.Replace(',', '.')

    return @(
        (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' 'LUMPSUM' (Get-ContainerEvaluation "20'RE") $price20),
        (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' 'LUMPSUM' (Get-ContainerEvaluation "40'RE") $price40)
    )
}

function Parse-HapagAdditionalLine {
    param(
        [string]$Line,
        [string]$CanonicalName
    )

    $normalizedLine = Normalize-Whitespace $Line
    $match = [regex]::Match($normalizedLine, '^.+?\s+(USD|EUR|JPY)\s+(-?\d+(?:[.,]\d+)?)\s+(-?\d+(?:[.,]\d+)?)\b', 'IgnoreCase')
    if (-not $match.Success) {
        return @()
    }

    $currency = $match.Groups[1].Value.ToUpperInvariant()
    $price20 = $match.Groups[2].Value.Replace(',', '.')
    $price40 = $match.Groups[3].Value.Replace(',', '.')

    return @(
        (New-PriceDetail $CanonicalName $currency $normalizedLine (Get-ContainerEvaluation "20'RE") $price20),
        (New-PriceDetail $CanonicalName $currency $normalizedLine (Get-ContainerEvaluation "40'RE") $price40)
    )
}

function Get-HapagAdditionalDetails {
    param(
        [string]$PageText,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals (Get-ExpectedAdditionals -Rules $Rules -Carrier $Carrier -Direction $Direction) -Rules $Rules
    $details = @()
    $lines = @($PageText -split "`r?`n")

    foreach ($definition in $definitions) {
        foreach ($line in $lines) {
            $normalizedLine = Normalize-Key $line
            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedLine -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            $details += Parse-HapagAdditionalLine -Line $line -CanonicalName $definition.Name
            break
        }
    }

    return $details
}

function Test-CoscoCanadaPdfText {
    param([string]$Text)

    $normalized = Normalize-Key $Text
    return ($normalized -match '\bCOSCO\b' -and $normalized -match '\bCANADA\b' -and $normalized -match 'RAMPS FROM HALIFAX')
}

function Get-CoscoCanadaValidityWindow {
    param([string]$PdfText)

    $match = [regex]::Match($PdfText, 'Valid from\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+till\s+(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'Unable to extract validity dates from COSCO Canada PDF.'
    }

    return [pscustomobject]@{
        Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
        End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
    }
}

function Get-CoscoCanadaOceanRouteEntries {
    param([string]$PdfText)

    $pages = @($PdfText -split "`f")
    $page = $pages | Where-Object { $_ -match 'Service' -and $_ -match 'Halifax' -and $_ -match 'La Spezia' } | Select-Object -First 1
    if (-not $page) {
        throw 'Unable to locate main service page in COSCO Canada PDF.'
    }

    $serviceBlockMatch = [regex]::Match($page, 'Service.+?Offer subject', 'Singleline,IgnoreCase')
    $serviceBlock = if ($serviceBlockMatch.Success) { $serviceBlockMatch.Value } else { $page }

    $originMatches = [regex]::Matches($serviceBlock, '(La Spezia|Genova)\s+\([^)]+\)', 'IgnoreCase')
    if ($originMatches.Count -eq 0) {
        throw 'Unable to extract origin ports from COSCO Canada PDF.'
    }

    $origins = New-Object System.Collections.Generic.List[string]
    foreach ($originMatch in $originMatches) {
        Add-UniqueString -List $origins -Value $originMatch.Groups[1].Value
    }

    $transitTimes = New-Object System.Collections.Generic.List[string]
    foreach ($line in ($serviceBlock -split "`r?`n")) {
        $normalizedLine = Normalize-Whitespace $line
        if (-not $normalizedLine) {
            continue
        }

        if ($normalizedLine -match 'Genova\s+\([^)]+\)\s+(?<days>\d{1,2})\b') {
            $transitTimes.Add($Matches['days'])
            continue
        }

        if ($normalizedLine -match '^(?<days>\d{1,2})$') {
            $transitTimes.Add($Matches['days'])
        }
    }

    $rateMatch = [regex]::Match($serviceBlock, "days\s+(?<price20>\d+(?:[.,]\d+)?)\s+USD\s+(?<price40>\d+(?:[.,]\d+)?)\s+USD", 'IgnoreCase')
    if (-not $rateMatch.Success) {
        throw 'Unable to extract dry ocean freight values from COSCO Canada PDF.'
    }

    $reeferMatch = [regex]::Match($serviceBlock, '(?<price>\d+(?:[.,]\d+)?)\s+USD(?!.*\d+\s+USD)', 'Singleline,IgnoreCase')
    if (-not $reeferMatch.Success) {
        throw 'Unable to extract reefer ocean freight value from COSCO Canada PDF.'
    }

    $price20 = Convert-LocalizedNumberText $rateMatch.Groups['price20'].Value
    $price40 = Convert-LocalizedNumberText $rateMatch.Groups['price40'].Value
    $priceReefer = Convert-LocalizedNumberText $reeferMatch.Groups['price'].Value

    $baseDetails = @(
        (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' 'Direct service to Halifax' "Cntr 20' Box" $price20),
        (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' 'Direct service to Halifax' "Cntr 40' Box" $price40),
        (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' 'Direct service to Halifax' "Cntr 40' Reefer" $priceReefer)
    )

    $entries = @()
    for ($i = 0; $i -lt $origins.Count; $i++) {
        $entries += [pscustomobject]@{
            Origin = $origins[$i]
            Destination = 'Halifax'
            TransitTime = if ($i -lt $transitTimes.Count) { $transitTimes[$i] } else { '' }
            PriceDetails = @($baseDetails)
        }
    }

    return $entries
}

function Get-CoscoCanadaRampEntries {
    param([string]$PdfText)

    $pages = @($PdfText -split "`f")
    $page = $pages | Where-Object { $_ -match 'RAMPS from Halifax' } | Select-Object -First 1
    if (-not $page) {
        return @()
    }

    $lines = @($page -split "`r?`n")
    $headerLine = $lines | Where-Object { $_ -match '^\s*RAMP\s+' } | Select-Object -First 1
    if (-not $headerLine) {
        throw 'Unable to locate ramp destination header in COSCO Canada PDF.'
    }

    $headerParts = @([regex]::Split((Normalize-Whitespace $headerLine), '\s{2,}') | Where-Object { $_ })
    if ($headerParts.Count -lt 3) {
        $headerParts = @([regex]::Split($headerLine.TrimEnd(), '\s{2,}') | Where-Object { $_ })
    }
    if ($headerParts.Count -lt 3) {
        throw 'Unable to parse ramp destinations from COSCO Canada PDF.'
    }

    $facilityLine = $lines | Where-Object { $_ -match 'CN Tashereau Intermodal Terminal' -and $_ -match 'CN Rail Brampton Facility' } | Select-Object -First 1
    if (-not $facilityLine) {
        throw 'Unable to parse ramp facilities from COSCO Canada PDF.'
    }

    $facilityParts = @([regex]::Split($facilityLine.TrimEnd(), '\s{2,}') | Where-Object { $_ })
    if ($facilityParts.Count -lt 3) {
        throw 'Unable to split ramp facility line in COSCO Canada PDF.'
    }

    $priceLines = @()
    $capture = $false
    foreach ($line in $lines) {
        if ($line -match '40HQ\s+REEFER') {
            $capture = $true
        }

        if (-not $capture) {
            continue
        }

        if ($line -match 'Pls note Weight limitation') {
            break
        }

        $matches = [regex]::Matches($line, '(?<price>\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)\s*USD', 'IgnoreCase')
        if ($matches.Count -ge 2) {
            $priceLines += ,@($matches | ForEach-Object { Convert-LocalizedNumberText $_.Groups['price'].Value })
        }
    }

    if ($priceLines.Count -lt 3) {
        throw 'Unable to extract ramp prices from COSCO Canada PDF.'
    }

    $evaluations = @(
        "Cntr 20' Box",
        "Cntr 40' Box",
        "Cntr 40' Reefer"
    )

    $destinations = @(
        [pscustomobject]@{ Name = 'Montreal'; Header = $headerParts[1]; Facility = $facilityParts[1]; Prices = @($priceLines[0][0], $priceLines[1][0], $priceLines[2][0]) },
        [pscustomobject]@{ Name = 'Toronto'; Header = $headerParts[2]; Facility = $facilityParts[2]; Prices = @($priceLines[0][1], $priceLines[1][1], $priceLines[2][1]) }
    )

    $entries = @()
    foreach ($destination in $destinations) {
        $details = @()
        for ($i = 0; $i -lt $evaluations.Count; $i++) {
            $details += (New-PriceDetail 'INLAND FREIGHT' 'USD' ("Ramp from Halifax - " + $destination.Facility) $evaluations[$i] $destination.Prices[$i])
        }

        $entries += [pscustomobject]@{
            Destination = $destination.Name
            Transshipment = 'Halifax'
            Comment = 'Ramp from Halifax'
            PriceDetails = @($details)
        }
    }

    return $entries
}

function Get-CoscoCanadaAdditionalDetails {
    param([string]$PdfText)

    $details = @()

    $dryMatch = [regex]::Match($PdfText, 'Notes and Surcharges\s+(?<currency>EUR)\s*(?<price>\d+(?:[.,]\d+)?)\/TEU[^\r\n]*\r?\nTHC\/THD\/ORS\/BUC\/EIS\/LWS included\s*\r?\n\s*Surcharges\s*\r?\nETS DRY', 'IgnoreCase')
    if ($dryMatch.Success) {
        $dryComment = ('ETS DRY (European Union Emissions Trading System): {0} {1}/TEU Q1 2026' -f $dryMatch.Groups['currency'].Value.ToUpperInvariant(), $dryMatch.Groups['price'].Value)
        $details += Parse-AdditionalText -Text $dryComment -CanonicalName 'ETS'
    }

    $reeferMatch = [regex]::Match($PdfText, 'ETS REEFER\s*\(European Union Emissions Trading System\)\s*:\s*(?<currency>EUR)\s*(?<price>\d+(?:[.,]\d+)?)\/TEU[^\r\n]*', 'IgnoreCase')
    if ($reeferMatch.Success) {
        $reeferComment = ('ETS REEFER (European Union Emissions Trading System): {0} {1}/TEU Q1 2026' -f $reeferMatch.Groups['currency'].Value.ToUpperInvariant(), $reeferMatch.Groups['price'].Value)
        $details += Parse-AdditionalText -Text $reeferComment -CanonicalName 'ETS'
    }

    return $details
}

function Convert-CoscoCanadaPdfToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = '',
        [string]$PdfText = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }

    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO Canada PDF adapter expects carrier COSCO. Received '$Carrier'."
    }

    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO Canada PDF adapter expects Export direction. Received '$Direction'."
    }

    if (-not $PdfText) {
        $PdfText = Get-PdfText -InputPath $InputPath
    }

    if (-not (Test-CoscoCanadaPdfText -Text $PdfText)) {
        throw 'COSCO Canada PDF markers not found in the PDF text.'
    }

    $validity = Get-CoscoCanadaValidityWindow -PdfText $PdfText
    $oceanRoutes = Get-CoscoCanadaOceanRouteEntries -PdfText $PdfText
    $rampEntries = Get-CoscoCanadaRampEntries -PdfText $PdfText
    $additionalDetails = Get-CoscoCanadaAdditionalDetails -PdfText $PdfText

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($route in $oceanRoutes) {
        Add-UniqueString -List $rawLocationNames -Value $route.Origin
        Add-UniqueString -List $rawLocationNames -Value $route.Destination
    }
    foreach ($rampEntry in $rampEntries) {
        Add-UniqueString -List $rawLocationNames -Value $rampEntry.Destination
        Add-UniqueString -List $rawLocationNames -Value $rampEntry.Transshipment
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1
    $halifaxCodes = @(Get-LocationCodes -RawName 'Halifax' -Rules $Rules -UnlocodeLookup $unlocodeLookup)

    foreach ($route in $oceanRoutes) {
        $originCodes = @(Get-LocationCodes -RawName $route.Origin -Rules $Rules -UnlocodeLookup $unlocodeLookup)
        $destinationCodes = @(Get-LocationCodes -RawName $route.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup)
        foreach ($originCode in $originCodes) {
            foreach ($destinationCode in $destinationCodes) {
                $details = @()
                $details += @($route.PriceDetails)
                $details += @($additionalDetails)
                $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $details -TransitTime $route.TransitTime
                $rowIndex++
            }
        }

        foreach ($rampEntry in $rampEntries) {
            $destinationCodes = @(Get-LocationCodes -RawName $rampEntry.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup)
            foreach ($originCode in $originCodes) {
                foreach ($destinationCode in $destinationCodes) {
                    $details = @()
                    $details += @($route.PriceDetails)
                    $details += @($rampEntry.PriceDetails)
                    $details += @($additionalDetails)
                    $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $details -TransshipmentAddress $halifaxCodes[0] -Comment $rampEntry.Comment
                    $rowIndex++
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Test-CoscoSouthAmericaPdfText {
    param([string]$Text)

    $normalized = Normalize-Key $Text
    return ($normalized -match 'SOUTH AMERICA' -and $normalized -match 'ETS DRY' -and $normalized -match 'ASUNCION')
}

function Get-CoscoSouthAmericaValidityWindow {
    param([string]$PdfText)

    $match = [regex]::Match($PdfText, 'Validity:\s*From\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+to\s+(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'Unable to extract validity dates from COSCO South America PDF.'
    }

    return [pscustomobject]@{
        Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
        End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
    }
}

function Parse-CoscoSouthAmericaRouteLine {
    param([string]$Line)

    $normalizedLine = Normalize-Whitespace $Line
    if (-not $normalizedLine) {
        return $null
    }

    $priceMatch = [regex]::Match($normalizedLine, '^(?<prefix>.+?)\s+[^\d]*(?<price20>\d{1,3}(?:\.\d{3})*(?:,\d+)?)\s+[^\d]*(?<price40>\d{1,3}(?:\.\d{3})*(?:,\d+)?)$', 'IgnoreCase')
    if (-not $priceMatch.Success) {
        return $null
    }

    $prefix = Normalize-Whitespace $priceMatch.Groups['prefix'].Value
    $price20 = Convert-LocalizedNumberText $priceMatch.Groups['price20'].Value
    $price40 = Convert-LocalizedNumberText $priceMatch.Groups['price40'].Value

    $multiOriginMatch = [regex]::Match($prefix, '^(?<destination>.+?)\s+GOA\s+(?<goa>\d{1,2})\s+SPE\s+(?<spe>\d{1,2})\s+SAL\s+(?<sal>\d{1,2})\s+via\s+(?<routing>.+?)\s+(?<service>[A-Za-z0-9+ ]+)$', 'IgnoreCase')
    if ($multiOriginMatch.Success) {
        return [pscustomobject]@{
            Destination = Normalize-Whitespace $multiOriginMatch.Groups['destination'].Value
            Routing = Normalize-Whitespace ('via ' + $multiOriginMatch.Groups['routing'].Value)
            Service = Normalize-Whitespace $multiOriginMatch.Groups['service'].Value
            Price20 = $price20
            Price40 = $price40
            TransitByOrigin = [ordered]@{
                'Genova' = $multiOriginMatch.Groups['goa'].Value
                'La Spezia' = $multiOriginMatch.Groups['spe'].Value
                'Salerno' = $multiOriginMatch.Groups['sal'].Value
            }
            TransitTime = ''
            HasTransitDelayNote = $false
        }
    }

    $singleOriginMatch = [regex]::Match($prefix, '^(?<destination>.+?)\s+(?<tt>\d{1,2})\s*\*\s+via\s+(?<routing>.+?)\s+(?<service>[A-Za-z0-9+ ]+)$', 'IgnoreCase')
    if ($singleOriginMatch.Success) {
        return [pscustomobject]@{
            Destination = Normalize-Whitespace $singleOriginMatch.Groups['destination'].Value
            Routing = Normalize-Whitespace ('via ' + $singleOriginMatch.Groups['routing'].Value)
            Service = Normalize-Whitespace $singleOriginMatch.Groups['service'].Value
            Price20 = $price20
            Price40 = $price40
            TransitByOrigin = @{}
            TransitTime = $singleOriginMatch.Groups['tt'].Value
            HasTransitDelayNote = $true
        }
    }

    return $null
}

function Get-CoscoSouthAmericaRouteEntries {
    param([string]$RawPdfText)

    $lines = @($RawPdfText -split "`r?`n")
    $entries = @()
    $pendingOrigins = New-Object System.Collections.Generic.List[string]
    $activeOrigins = @()
    $blockStarted = $false
    $knownOrigins = @{
        'GENOVA' = 'Genova'
        'LA SPEZIA' = 'La Spezia'
        'SALERNO' = 'Salerno'
        'VENEZIA' = 'Venezia'
        'ANCONA' = 'Ancona'
        'RAVENNA' = 'Ravenna'
        'BARI (IPM)' = 'Bari'
        'BARI' = 'Bari'
    }

    foreach ($line in $lines) {
        $normalizedLine = Normalize-Whitespace $line
        if (-not $normalizedLine) {
            continue
        }

        if ($normalizedLine -match '^POD \(Terminal\) TT - days Service 20'' 40''/HQ$') {
            $blockStarted = $true
            continue
        }

        if (-not $blockStarted) {
            continue
        }

        if ($normalizedLine -match '^\* TT subject to delay due to double transshipment\.$') {
            break
        }

        $originKey = Normalize-Key $normalizedLine
        if ($knownOrigins.ContainsKey($originKey)) {
            if ($activeOrigins.Count -gt 0) {
                $activeOrigins = @()
                $pendingOrigins.Clear()
            }

            Add-UniqueString -List $pendingOrigins -Value $knownOrigins[$originKey]
            continue
        }

        if ($normalizedLine -match '^\(.*\)\s*/?$') {
            continue
        }

        $routeEntry = Parse-CoscoSouthAmericaRouteLine -Line $normalizedLine
        if (-not $routeEntry) {
            continue
        }

        if ($pendingOrigins.Count -gt 0) {
            $activeOrigins = @($pendingOrigins.ToArray())
            $pendingOrigins.Clear()
        }

        if ($activeOrigins.Count -eq 0) {
            throw "Unable to determine the active origin block for route line '$normalizedLine'."
        }

        $entries += [pscustomobject]@{
            Origins = @($activeOrigins)
            Destination = $routeEntry.Destination
            Routing = $routeEntry.Routing
            Service = $routeEntry.Service
            Price20 = $routeEntry.Price20
            Price40 = $routeEntry.Price40
            TransitByOrigin = $routeEntry.TransitByOrigin
            TransitTime = $routeEntry.TransitTime
            HasTransitDelayNote = $routeEntry.HasTransitDelayNote
        }
    }

    if ($entries.Count -eq 0) {
        throw 'Unable to extract route rows from COSCO South America PDF.'
    }

    return $entries
}

function Get-CoscoSouthAmericaAdditionalDetails {
    param(
        [string]$PdfText,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals (Get-ExpectedAdditionals -Rules $Rules -Carrier $Carrier -Direction $Direction) -Rules $Rules
    $details = @()
    $lines = @($PdfText -split "`r?`n")

    foreach ($definition in $definitions) {
        foreach ($line in $lines) {
            $normalizedLine = Normalize-Key $line
            if ($normalizedLine -match 'INCLUDED') {
                continue
            }

            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedLine -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            $parsedDetails = @(Parse-AdditionalText -Text $line -CanonicalName $definition.Name)
            if ($parsedDetails.Count -gt 0) {
                $details += $parsedDetails
            }
            break
        }
    }

    return $details
}

function Convert-CoscoSouthAmericaPdfToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = '',
        [string]$PdfText = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }

    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO South America PDF adapter expects carrier COSCO. Received '$Carrier'."
    }

    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO South America PDF adapter expects Export direction. Received '$Direction'."
    }

    if (-not $PdfText) {
        $PdfText = Get-PdfText -InputPath $InputPath
    }

    if (-not (Test-CoscoSouthAmericaPdfText -Text $PdfText)) {
        throw 'COSCO South America PDF markers not found in the PDF text.'
    }

    $rawPdfText = Get-PdfText -InputPath $InputPath -Mode 'raw'
    $validity = Get-CoscoSouthAmericaValidityWindow -PdfText $PdfText
    $routeEntries = Get-CoscoSouthAmericaRouteEntries -RawPdfText $rawPdfText
    $additionalDetails = Get-CoscoSouthAmericaAdditionalDetails -PdfText $PdfText -Carrier $normalizedCarrier -Direction $normalizedDirection -Rules $Rules

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($routeEntry in $routeEntries) {
        foreach ($origin in $routeEntry.Origins) {
            Add-UniqueString -List $rawLocationNames -Value $origin
        }
        Add-UniqueString -List $rawLocationNames -Value $routeEntry.Destination
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($routeEntry in $routeEntries) {
        $destinationCodes = @(Get-LocationCodes -RawName $routeEntry.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup)
        $detailCommentParts = @()
        if ($routeEntry.Service) {
            $detailCommentParts += $routeEntry.Service
        }
        if ($routeEntry.Routing) {
            $detailCommentParts += $routeEntry.Routing
        }
        $detailComment = Normalize-Whitespace ($detailCommentParts -join ' | ')

        $baseDetails = @(
            (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'EUR' $detailComment (Get-ContainerEvaluation "20'") $routeEntry.Price20),
            (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'EUR' $detailComment (Get-ContainerEvaluation "40'") $routeEntry.Price40)
        )

        foreach ($origin in $routeEntry.Origins) {
            $originCodes = @(Get-LocationCodes -RawName $origin -Rules $Rules -UnlocodeLookup $unlocodeLookup)
            $transitTime = ''
            if ($routeEntry.TransitByOrigin.Count -gt 0 -and $routeEntry.TransitByOrigin.Contains($origin)) {
                $transitTime = $routeEntry.TransitByOrigin[$origin]
            } elseif ($routeEntry.TransitTime) {
                $transitTime = $routeEntry.TransitTime
            }

            $rowComment = if ($routeEntry.HasTransitDelayNote) {
                'TT subject to delay due to double transshipment.'
            } else {
                ''
            }

            foreach ($originCode in $originCodes) {
                foreach ($destinationCode in $destinationCodes) {
                    $details = @()
                    $details += @($baseDetails)
                    $details += @($additionalDetails)
                    $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $details -TransitTime $transitTime -Comment $rowComment
                    $rowIndex++
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Get-ContainerEvaluation {
    param([string]$Label)

    $normalized = Normalize-Key $Label
    switch ($normalized) {
        "20'" { return "Cntr 20' Box" }
        "20" { return "Cntr 20' Box" }
        '20BOX' { return "Cntr 20' Box" }
        '20GP' { return "Cntr 20' Box" }
        "20'RE" { return "Cntr 20' Reefer" }
        '20RE' { return "Cntr 20' Reefer" }
        "40'" { return "Cntr 40' Box" }
        "40" { return "Cntr 40' Box" }
        '40BOX/HQ' { return "Cntr 40' Box" }
        '40GP' { return "Cntr 40' Box" }
        '40GP/40HQ' { return "Cntr 40' Box" }
        "40'RE" { return "Cntr 40' Reefer" }
        '40RE' { return "Cntr 40' Reefer" }
        "20DV" { return "Cntr 20' Box" }
        "40DV" { return "Cntr 40' Box" }
        "40HC" { return "Cntr 40' HC" }
        'TEU' { return 'TEUS' }
        'TEUS' { return 'TEUS' }
        'CNTR' { return 'CNTR' }
        default { return $normalized }
    }
}

function Get-CurrencyCode {
    param([string]$Symbol)

    $trimmed = Normalize-Key $Symbol
    if ($trimmed -eq '$') {
        return 'USD'
    }
    if (($trimmed -eq ([string][char]0x20AC)) -or ($trimmed -eq 'EUR')) {
        return 'EUR'
    }
    return $trimmed
}

function Get-RateRows {
    param([object[]]$Rows)

    $result = @()
    foreach ($row in $Rows) {
        if ($row.RowNumber -lt 5) {
            continue
        }

        $aValue = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if ($aValue -like 'All above*') {
            break
        }

        $numericCount = 0
        foreach ($col in @('B','C','D','E','F','G','H','I','J','K','L','M')) {
            $value = Normalize-Whitespace (Get-Cell $row.Cells $col $row.RowNumber)
            if ($value -match '^\d+(\.\d+)?$') {
                $numericCount++
            }
        }

        if ($aValue -and $numericCount -ge 2) {
            $result += $row
        }
    }

    return $result
}

function Get-DestinationMap {
    param([object[]]$Rows)

    $row3 = $Rows | Where-Object { $_.RowNumber -eq 3 } | Select-Object -First 1
    $row4 = $Rows | Where-Object { $_.RowNumber -eq 4 } | Select-Object -First 1
    if (-not $row3 -or -not $row4) {
        throw 'Destination header rows not found.'
    }

    $pairs = @(
        @('B', 'C'),
        @('D', 'E'),
        @('F', 'G'),
        @('H', 'I'),
        @('J', 'K'),
        @('L', 'M')
    )

    $destinations = @()
    foreach ($pair in $pairs) {
        $destinationName = Normalize-Whitespace (Get-Cell $row3.Cells $pair[0] 3)
        if (-not $destinationName) {
            continue
        }

        $firstSize = Normalize-Whitespace (Get-Cell $row4.Cells $pair[0] 4)
        $secondSize = Normalize-Whitespace (Get-Cell $row4.Cells $pair[1] 4)

        $destinations += [pscustomobject]@{
            Destination = $destinationName
            PriceColumns = @(
                [pscustomobject]@{
                    Column = $pair[0]
                    Evaluation = Get-ContainerEvaluation $firstSize
                },
                [pscustomobject]@{
                    Column = $pair[1]
                    Evaluation = Get-ContainerEvaluation $secondSize
                }
            )
        }
    }

    return $destinations
}

function Test-GenericClassicPairRateMatrixLayout {
    param(
        [object[]]$Rows,
        [object]$HeaderRow,
        [object[]]$Destinations,
        [object[]]$NoteRows
    )

    if (-not $HeaderRow) {
        return $false
    }

    $headerText = Normalize-Whitespace (Get-Cell $HeaderRow.Cells 'A' $HeaderRow.RowNumber)
    if ($headerText -notmatch 'etd\s+\d{1,2}/\d{1,2}/\d{2,4}.*up to\s+\d{1,2}/\d{1,2}/\d{2,4}') {
        return $false
    }

    if ((@($Destinations)).Count -ne 6) {
        return $false
    }

    $row4 = $Rows | Where-Object { $_.RowNumber -eq 4 } | Select-Object -First 1
    if (-not $row4) {
        return $false
    }

    $expectedLabels = @("20'", "40'", "20'", "40'", "20'", "40'", "20'", "40'", "20'", "40'", "20'", "40'")
    $actualLabels = @()
    foreach ($column in @('B','C','D','E','F','G','H','I','J','K','L','M')) {
        $actualLabels += Normalize-Whitespace (Get-Cell $row4.Cells $column 4)
    }

    if ($actualLabels.Count -ne $expectedLabels.Count) {
        return $false
    }

    for ($i = 0; $i -lt $expectedLabels.Count; $i++) {
        if ($actualLabels[$i] -ne $expectedLabels[$i]) {
            return $false
        }
    }

    $markerCount = 0
    foreach ($pattern in @('^\+ BRC', '^\+ ECA', '^\+ Emissions Trading System', '^\+ Fuel EU')) {
        if (Find-RowByText $NoteRows $pattern) {
            $markerCount++
        }
    }

    return ($markerCount -ge 3)
}

function Get-LocationCodes {
    param(
        [string]$RawName,
        [hashtable]$Rules,
        [hashtable]$UnlocodeLookup = $null,
        [string]$CountryHint = ''
    )

    $key = Normalize-Key $RawName
    if ($Rules.UnlocodesByName.ContainsKey($key)) {
        return @($Rules.UnlocodesByName[$key])
    }

    foreach ($candidate in (Get-LocationLookupCandidates -RawName $RawName)) {
        $ruleCodes = New-Object System.Collections.Generic.List[string]
        foreach ($candidateKey in (Get-LocationIndexKeysForName -Text $candidate)) {
            if (-not $Rules.UnlocodesByName.ContainsKey($candidateKey)) {
                continue
            }

            foreach ($code in @($Rules.UnlocodesByName[$candidateKey])) {
                if (-not $ruleCodes.Contains($code)) {
                    $ruleCodes.Add($code)
                }
            }
        }

        if ($ruleCodes.Count -gt 0) {
            return $ruleCodes.ToArray()
        }
    }

    $lookupCodes = @(Resolve-UnlocodesFromLookup -RawName $RawName -Lookup $UnlocodeLookup -CountryHint $CountryHint)
    if ($lookupCodes.Count -gt 0) {
        return $lookupCodes
    }

    throw "UNLOCODE mapping not configured for '$RawName'."
}

function Get-ExpectedAdditionals {
    param(
        [hashtable]$Rules,
        [string]$Carrier,
        [string]$Direction
    )

    $normalizedDirection = Normalize-Key $Direction
    $normalizedCarrier = Normalize-Key $Carrier
    if (-not $Rules.CarrierAdditionals.ContainsKey($normalizedDirection)) {
        throw "Direction '$Direction' not configured in rules."
    }
    if (-not $Rules.CarrierAdditionals[$normalizedDirection].ContainsKey($normalizedCarrier)) {
        throw "Carrier '$Carrier' not configured for direction '$Direction'."
    }
    return @($Rules.CarrierAdditionals[$normalizedDirection][$normalizedCarrier])
}

function Get-AdditionalPatterns {
    param([string]$CanonicalName)

    $name = Normalize-Key $CanonicalName
    switch ($name) {
        'ETS' { return @('\bETS\b', 'EMISSIONS TRADING SYSTEM') }
        'FUEL EU' { return @('FUEL EU') }
        'WAR RISK' { return @('WAR RISK', '\bWR\b') }
        'WR' { return @('\bWR\b', 'WAR RISK') }
        'PSS ISRAEL' { return @('PSS.*ISRAEL') }
        'WRS ISRAEL' { return @('WRS.*ISRAEL', 'WAR RISK.*ISRAEL') }
        'TAC (MERSIN)' { return @('TAC.*MERSIN') }
        'ORS (CASABLANCA)' { return @('ORS.*CASABLANCA') }
        'WAR RISK PERSIAN GULF' { return @('WAR RISK.*PERSIAN GULF') }
        'WAR RISK JEDDAH' { return @('WAR RISK.*JEDDAH') }
        'C.S. AT GRPIR' { return @('C\.?S\.?\s+AT\s+GRPIR', 'C\.?S\.?.*GRPIR') }
        'PSS TO ISRAEL' { return @('PSS.*ISRAEL') }
        'EIS TO ISRAEL' { return @('EIS.*ISRAEL') }
        'OCC PER DJBUTI' { return @('OCC.*DJ(B|I)OUTI') }
        'DTHC AT PORT SUDAN PREPAID' { return @('DTHC.*PORT SUDAN.*PREPAID') }
        'DTHC AT CHATTOGRAM PREPAID' { return @('DTHC.*CHATTOGRAM.*PREPAID') }
        'PAD 75 USD/TEU + DOF 45 USD/TEU FOR MALE PREPAID' { return @('PAD.*DOF.*MALE.*PREPAID') }
        'DTHC UMM QASR PREPAID' { return @('DTHC.*UMM QASR.*PREPAID') }
        'SUEZ CANAL' { return @('SUEZ CANAL', '\bSUEZ\b') }
        'ADEN GULF' { return @('ADEN GULF', '\bADEN\b') }
        'EUETS' { return @('EUETS', 'EU ETS') }
        default { return @([regex]::Escape($name)) }
    }
}

function Get-AdditionalDefinitions {
    param(
        [string[]]$ExpectedAdditionals,
        [hashtable]$Rules
    )

    $definitions = @()
    foreach ($name in $ExpectedAdditionals) {
        $normalizedName = Normalize-Key $name
        $patterns = $null
        $applyTargets = @()

        if ($Rules.ContainsKey('AdditionalRuleDetails') -and $Rules.AdditionalRuleDetails.ContainsKey($normalizedName)) {
            $ruleDetails = $Rules.AdditionalRuleDetails[$normalizedName]
            if ($ruleDetails.ContainsKey('Patterns')) {
                $patterns = @($ruleDetails.Patterns)
            }
            if ($ruleDetails.ContainsKey('ApplyTargets')) {
                $applyTargets = @($ruleDetails.ApplyTargets)
            }
        }

        if (-not $patterns) {
            $patterns = @(Get-AdditionalPatterns $name)
        }

        $definitions += [pscustomobject]@{
            Name = $normalizedName
            Patterns = @($patterns)
            ApplyTargets = @($applyTargets)
        }
    }
    return $definitions
}

function Resolve-CarrierDirection {
    param(
        [object[]]$Rows,
        [hashtable]$Rules,
        [string]$Carrier,
        [string]$Direction
    )

    if ($Carrier -and $Direction) {
        return [pscustomobject]@{
            Carrier = Normalize-Key $Carrier
            Direction = Normalize-Key $Direction
        }
    }

    $combined = Normalize-Key (($Rows | ForEach-Object { Get-RowComment $_ }) -join ' ')
    $candidates = @()

    foreach ($directionKey in $Rules.CarrierAdditionals.Keys) {
        if ($Direction -and $directionKey -ne (Normalize-Key $Direction)) {
            continue
        }

        foreach ($carrierKey in $Rules.CarrierAdditionals[$directionKey].Keys) {
            if ($Carrier -and $carrierKey -ne (Normalize-Key $Carrier)) {
                continue
            }

            $score = 0
            foreach ($definition in (Get-AdditionalDefinitions -ExpectedAdditionals $Rules.CarrierAdditionals[$directionKey][$carrierKey] -Rules $Rules)) {
                foreach ($pattern in $definition.Patterns) {
                    if ($combined -match $pattern) {
                        $score++
                        break
                    }
                }
            }

            $candidates += [pscustomobject]@{
                Carrier = $carrierKey
                Direction = $directionKey
                Score = $score
            }
        }
    }

    $ordered = $candidates | Sort-Object Score -Descending
    if (-not $ordered -or $ordered[0].Score -le 0) {
        throw 'Unable to infer carrier/direction from the listino. Please pass -Carrier and -Direction explicitly.'
    }

    $topScore = $ordered[0].Score
    $best = @($ordered | Where-Object { $_.Score -eq $topScore })
    if ((@($best)).Count -gt 1) {
        $pairs = ($best | ForEach-Object { "$($_.Carrier) $($_.Direction)" }) -join ', '
        throw "Carrier/direction inference is ambiguous ($pairs). Please pass -Carrier and -Direction explicitly."
    }

    return $best[0]
}

function Get-OriginName {
    param([string]$RawName)

    $name = Normalize-Whitespace $RawName
    if ($name -eq 'Termini Imerese/Augusta/Trapani/Pozzallo') {
        return 'SICILY PORTS'
    }
    return $name
}

function Find-RowByText {
    param(
        [object[]]$Rows,
        [string]$Pattern
    )

    foreach ($row in $Rows) {
        $text = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if ($text -match $Pattern) {
            return $row
        }
    }

    return $null
}

function Get-RowComment {
    param([object]$Row)

    $a = Normalize-Whitespace (Get-Cell $Row.Cells 'A' $Row.RowNumber)
    $b = Normalize-Whitespace (Get-Cell $Row.Cells 'B' $Row.RowNumber)
    $c = Normalize-Whitespace (Get-Cell $Row.Cells 'C' $Row.RowNumber)
    return (Normalize-Whitespace (($a, $b, $c -join ' ')))
}

function Get-EvaluationFromUnitToken {
    param([string]$Token)

    $normalized = Normalize-Key $Token
    if ($normalized -match '^20''?RE') {
        return "Cntr 20' Reefer"
    }
    if ($normalized -match '^40''?RE') {
        return "Cntr 40' Reefer"
    }
    if ($normalized -match '^40(?:HC|HQ|H)$') {
        return "Cntr 40' HC"
    }
    if ($normalized -match '^20(?:BOX|GP)') {
        return "Cntr 20' Box"
    }
    if ($normalized -match '^40(?:BOX/HQ|GP/40HQ|GP|HQ)') {
        return "Cntr 40' Box"
    }
    if ($normalized -match '^20') {
        return "Cntr 20' Box"
    }
    if ($normalized -match '^40') {
        return "Cntr 40' Box"
    }
    if ($normalized -match '^TEU') {
        return 'TEUS'
    }
    if ($normalized -match '^CNTR') {
        return 'CNTR'
    }
    return $normalized
}

function Get-AdditionalApplyTargets {
    param(
        [string]$CanonicalName,
        [string]$Comment
    )

    $text = Normalize-Key ($CanonicalName + ' ' + $Comment)
    $targets = @()
    $locationPatterns = @(
        @{ Pattern = 'DJ(B|I)OUTI'; Target = 'DJIBOUTI' },
        @{ Pattern = 'ADEN'; Target = 'ADEN' },
        @{ Pattern = 'MUKALLAH'; Target = 'MUKALLAH' },
        @{ Pattern = 'PORT SUDAN'; Target = 'PORT SUDAN' },
        @{ Pattern = 'CHATTOGRAM'; Target = 'CHATTOGRAM' },
        @{ Pattern = 'MALE'; Target = 'MALE' },
        @{ Pattern = 'UMM QASR'; Target = 'UMM QASR' },
        @{ Pattern = 'CASABLANCA'; Target = 'CASABLANCA' },
        @{ Pattern = 'MERSIN'; Target = 'MERSIN' },
        @{ Pattern = 'JEDDAH'; Target = 'JEDDAH' },
        @{ Pattern = 'PERSIAN GULF'; Target = 'PERSIAN GULF' },
        @{ Pattern = 'ISRAEL'; Target = 'ISRAEL' },
        @{ Pattern = 'YANGON'; Target = 'YANGON' }
    )

    foreach ($item in $locationPatterns) {
        if ($text -match $item.Pattern) {
            $targets += $item.Target
        }
    }

    return @($targets | Select-Object -Unique)
}

function Resolve-AdditionalCurrency {
    param(
        [object]$Row,
        [string]$Comment
    )

    $rowCurrency = Get-CurrencyCode (Get-Cell $Row.Cells 'B' $Row.RowNumber)
    if ($rowCurrency) {
        return $rowCurrency
    }

    $normalized = Normalize-Key $Comment
    if ($normalized -match '\bUSD\b' -or $Comment.Contains('$')) {
        return 'USD'
    }
    if ($normalized -match '\bEUR\b' -or $Comment.Contains([char]0x20AC)) {
        return 'EUR'
    }
    return ''
}

function New-PriceDetail {
    param(
        [string]$Name,
        [string]$Currency,
        [string]$Comment,
        [string]$Evaluation,
        [object]$Price
    )

    return [pscustomobject]@{
        Name = $Name
        Currency = $Currency
        Comment = $Comment
        Evaluation = $Evaluation
        Minimum = ''
        Maximum = ''
        TiersType = ''
        Price = [string]$Price
    }
}

function Copy-PriceDetail {
    param(
        [pscustomobject]$Detail,
        [string]$Name = '',
        [string]$Comment = '',
        [string]$Evaluation = ''
    )

    return [pscustomobject]@{
        Name = if ($Name) { $Name } else { $Detail.Name }
        Currency = $Detail.Currency
        Comment = if ($PSBoundParameters.ContainsKey('Comment')) { $Comment } else { $Detail.Comment }
        Evaluation = if ($Evaluation) { $Evaluation } else { $Detail.Evaluation }
        Minimum = $Detail.Minimum
        Maximum = $Detail.Maximum
        TiersType = $Detail.TiersType
        Price = [string]$Detail.Price
    }
}

function Clear-PriceDetailComments {
    param([object[]]$Details)

    $normalizedDetails = New-Object System.Collections.Generic.List[object]
    foreach ($detail in @($Details)) {
        $normalizedDetails.Add((Copy-PriceDetail -Detail $detail -Comment ''))
    }

    return $normalizedDetails.ToArray()
}

function Add-Missing40HcFallbackDuplicates {
    param(
        [object[]]$Details,
        [object[]]$PriceColumns
    )

    $hasExplicit40Hc = @($PriceColumns | Where-Object { $_.Evaluation -eq "Cntr 40' HC" }).Count -gt 0
    if ($hasExplicit40Hc) {
        return @($Details)
    }

    $hasFortyBox = @($PriceColumns | Where-Object { $_.Evaluation -eq "Cntr 40' Box" }).Count -gt 0
    if (-not $hasFortyBox) {
        return @($Details)
    }

    $expandedDetails = New-Object System.Collections.Generic.List[object]
    foreach ($detail in @($Details)) {
        $expandedDetails.Add($detail)
        if ($detail.Evaluation -eq "Cntr 40' Box") {
            $expandedDetails.Add((Copy-PriceDetail -Detail $detail -Evaluation "Cntrs 40' HC"))
        }
    }

    return $expandedDetails.ToArray()
}

function Parse-AdditionalText {
    param(
        [string]$Text,
        [string]$CanonicalName,
        [string]$DefaultCurrency = ''
    )

    $comment = Normalize-Whitespace $Text
    if (-not $comment) {
        return @()
    }

    $currency = $DefaultCurrency
    if (-not $currency) {
        $normalized = Normalize-Key $comment
        if ($normalized -match '\bUSD\b' -or $comment.Contains('$')) {
            $currency = 'USD'
        } elseif ($normalized -match '\bEUR\b' -or $comment.Contains([char]0x20AC)) {
            $currency = 'EUR'
        } elseif ($normalized -match '\bJPY\b') {
            $currency = 'JPY'
        }
    }

    $unitPattern = '(?<unit>20(?:BOX|GP|DV|HC|RE|''|)|40(?:BOX/HQ|GP/40HQ|GP|HQ|DV|HC|RE|''|)|TEU(?:S)?|CNTR|UNIT)'
    $patterns = @(
        "(?<currency>USD|EUR|JPY)\s*(?<price>\d+(?:[.,]\d+)?)\s*(?:/|PER)\s*$unitPattern",
        "(?<price>\d+(?:[.,]\d+)?)\s*(?<currency>USD|EUR|JPY)\s*(?:/|PER)\s*$unitPattern",
        "(?<price>\d+(?:[.,]\d+)?)\s*(?:/|PER)\s*$unitPattern"
    )

    $details = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($pattern in $patterns) {
        foreach ($match in [regex]::Matches($comment, $pattern, 'IgnoreCase')) {
            $resolvedCurrency = if ($match.Groups['currency'].Success) {
                $match.Groups['currency'].Value.ToUpperInvariant()
            } else {
                $currency
            }

            if (-not $resolvedCurrency) {
                continue
            }

            $price = $match.Groups['price'].Value.Replace(',', '.')
            $unit = $match.Groups['unit'].Value
            $evaluation = Get-EvaluationFromUnitToken $unit
            $key = "$CanonicalName|$resolvedCurrency|$evaluation|$price"
            if ($seen.ContainsKey($key)) {
                continue
            }

            $seen[$key] = $true
            $details.Add((New-PriceDetail $CanonicalName $resolvedCurrency $comment $evaluation $price))
        }
    }

    return $details.ToArray()
}

function Parse-AdditionalRow {
    param(
        [object]$Row,
        [string]$CanonicalName
    )

    $comment = Get-RowComment $Row
    $currency = Resolve-AdditionalCurrency -Row $Row -Comment $comment
    return ,@(Parse-AdditionalText -Text $comment -CanonicalName $CanonicalName -DefaultCurrency $currency)
}

function Get-ExpectedAdditionalDetails {
    param(
        [object[]]$Rows,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $expected = Get-ExpectedAdditionals -Rules $Rules -Carrier $Carrier -Direction $Direction
    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals $expected -Rules $Rules
    $details = @()

    foreach ($definition in $definitions) {
        foreach ($row in $Rows) {
            $comment = Get-RowComment $row
            $normalizedComment = Normalize-Key $comment
            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedComment -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            $parsed = Parse-AdditionalRow -Row $row -CanonicalName $definition.Name
            if ((@($parsed)).Count -eq 0) {
                continue
            }

            $applyTargets = @($definition.ApplyTargets)
            if ((@($applyTargets)).Count -eq 0) {
                $applyTargets = Get-AdditionalApplyTargets -CanonicalName $definition.Name -Comment $comment
            }
            foreach ($detail in $parsed) {
                $details += [pscustomobject]@{
                    AppliesTo = @($applyTargets)
                    Detail = $detail
                }
            }
        }
    }

    return $details
}

function Should-ApplyAdditionalToDestination {
    param(
        [string[]]$ApplyTargets,
        [string]$DestinationName
    )

    if (-not $ApplyTargets -or (@($ApplyTargets)).Count -eq 0) {
        return $true
    }

    $normalizedDestination = Normalize-Key $DestinationName
    foreach ($target in $ApplyTargets) {
        if ($normalizedDestination -like "*$target*") {
            return $true
        }
    }
    return $false
}

function Get-AdditionalDetailTemplates {
    param([object[]]$Rows)

    $templates = @()

    $brcRow = Find-RowByText $Rows '^\+ BRC'
    if ($brcRow) {
        $comment = Get-RowComment $brcRow
        $currency = Get-CurrencyCode (Get-Cell $brcRow.Cells 'B' $brcRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'BRC' $currency $comment "Cntr 20' Box" 280) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'BRC' $currency $comment "Cntr 40' Box" 560) }
        )
    }

    $ecaRow = Find-RowByText $Rows '^\+ ECA'
    if ($ecaRow) {
        $comment = Get-RowComment $ecaRow
        $currency = Get-CurrencyCode (Get-Cell $ecaRow.Cells 'B' $ecaRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ECA' $currency $comment "Cntr 20' Box" 16) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ECA' $currency $comment "Cntr 40' Box" 32) }
        )
    }

    $etsRow = Find-RowByText $Rows '^\+ Emissions Trading System'
    if ($etsRow) {
        $comment = Get-RowComment $etsRow
        $currency = Get-CurrencyCode (Get-Cell $etsRow.Cells 'B' $etsRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ETS' $currency $comment "Cntr 20' Box" 32) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ETS' $currency $comment "Cntr 40' Box" 64) }
        )
    }

    $fuelRow = Find-RowByText $Rows '^\+ Fuel EU'
    if ($fuelRow) {
        $comment = Get-RowComment $fuelRow
        $currency = Get-CurrencyCode (Get-Cell $fuelRow.Cells 'B' $fuelRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'FUEL EU' $currency $comment "Cntr 20' Box" 10) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'FUEL EU' $currency $comment "Cntr 40' Box" 20) }
        )
    }

    $selfHeatingRow = Find-RowByText $Rows '^\+ Self Heating Commodities'
    if ($selfHeatingRow) {
        $comment = Get-RowComment $selfHeatingRow
        $currency = Get-CurrencyCode (Get-Cell $selfHeatingRow.Cells 'B' $selfHeatingRow.RowNumber)
        $templates += [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'SELF HEATING' $currency $comment 'CNTR' 50) }
    }

    $fqsRow = Find-RowByText $Rows '^\+FQS'
    if ($fqsRow) {
        $comment = Get-RowComment $fqsRow
        $currency = Get-CurrencyCode (Get-Cell $fqsRow.Cells 'B' $fqsRow.RowNumber)
        $templates += [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'FQS' $currency $comment 'CNTR' 100) }
    }

    $openTopRow = Find-RowByText $Rows '^Open top surcharge'
    if ($openTopRow) {
        $comment = Get-RowComment $openTopRow
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'OPEN TOP' 'EUR' $comment 'TEUS' 500) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'OPEN TOP' 'EUR' $comment 'TEUS' 500) }
        )
    }

    $imoStorageRow = Find-RowByText $Rows '^IMO storage'
    if ($imoStorageRow) {
        $comment = Get-RowComment $imoStorageRow
        $currency = Get-CurrencyCode (Get-Cell $imoStorageRow.Cells 'B' $imoStorageRow.RowNumber)
        $templates += [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'IMO STORAGE' $currency $comment 'TEUS' 25) }
    }

    $thcRow = Find-RowByText $Rows '^Loading THC$'
    if ($thcRow) {
        $comment = Get-RowComment $thcRow
        $currency = Get-CurrencyCode (Get-Cell $thcRow.Cells 'B' $thcRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'THC' $currency $comment 'CNTR' 190) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'THC' $currency $comment 'CNTR' 380) }
        )
    }

    $ispsRow = Find-RowByText $Rows '^ISPS$'
    if ($ispsRow) {
        $comment = Get-RowComment $ispsRow
        $currency = Get-CurrencyCode (Get-Cell $ispsRow.Cells 'B' $ispsRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ISPS' $currency $comment 'CNTR' 18) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'ISPS' $currency $comment 'CNTR' 25) }
        )
    }

    $loloRow = Find-RowByText $Rows '^Lift on - Lift off'
    if ($loloRow) {
        $comment = Get-RowComment $loloRow
        $currency = Get-CurrencyCode (Get-Cell $loloRow.Cells 'B' $loloRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'LIFT ON/LIFT OFF' $currency $comment 'CNTR' 68) },
            [pscustomobject]@{ AppliesTo = '*'; Detail = (New-PriceDetail 'LIFT ON/LIFT OFF' $currency $comment 'CNTR' 88) }
        )
    }

    $occRow = Find-RowByText $Rows '^Operation Cost Contribution \(OCC\)'
    if ($occRow) {
        $comment = Normalize-Whitespace (Get-Cell $occRow.Cells 'A' $occRow.RowNumber)
        $templates += @(
            [pscustomobject]@{ AppliesTo = 'BERBERA'; Detail = (New-PriceDetail 'OCC' 'USD' $comment "Cntr 20' Box" 240) },
            [pscustomobject]@{ AppliesTo = 'BERBERA'; Detail = (New-PriceDetail 'OCC' 'USD' $comment "Cntr 40' Box" 345) },
            [pscustomobject]@{ AppliesTo = 'BERBERA'; Detail = (New-PriceDetail 'OCC' 'USD' $comment "Cntr 20' Box" 305) },
            [pscustomobject]@{ AppliesTo = 'BERBERA'; Detail = (New-PriceDetail 'OCC' 'USD' $comment "Cntr 40' Box" 443) }
        )
    }

    return $templates
}

function Get-OutputHeaders {
    $headers = @(
        'Row Index',
        'From Geo Area',
        'To Geo Area',
        'From Address',
        'To Address',
        'From Country',
        'To Country',
        'Validity Start Date',
        'Expiration Date',
        'Transshipment Address',
        'Supplier',
        'Carrier',
        'Reference',
        'Comment',
        'External Note',
        'NAC',
        'Is Dangerous',
        'Transit Time'
    )

    for ($i = 1; $i -le 25; $i++) {
        $headers += @(
            "Price Detail $i",
            "Price Detail $i - Currency",
            "Price Detail $i - Comment",
            "Price Detail $i - Evaluation",
            "Price Detail $i - Minimum",
            "Price Detail $i - Maximum",
            "Price Detail $i - Tiers Type",
            "Price Detail $i - Price"
        )
    }

    return $headers
}

function Convert-RouteToRow {
    param(
        [int]$Index,
        [string]$FromAddress,
        [string]$ToAddress,
        [string]$ValidityStart,
        [string]$ValidityEnd,
        [string]$Carrier,
        [object[]]$PriceDetails,
        [string]$TransitTime = '',
        [string]$Reference = '',
        [string]$TransshipmentAddress = '',
        [string]$Comment = ''
    )

    $row = [ordered]@{
        'Row Index' = $Index
        'From Geo Area' = ''
        'To Geo Area' = ''
        'From Address' = $FromAddress
        'To Address' = $ToAddress
        'From Country' = ''
        'To Country' = ''
        'Validity Start Date' = $ValidityStart
        'Expiration Date' = $ValidityEnd
        'Transshipment Address' = $TransshipmentAddress
        'Supplier' = ''
        'Carrier' = $Carrier
        'Reference' = $Reference
        'Comment' = $Comment
        'External Note' = ''
        'NAC' = ''
        'Is Dangerous' = ''
        'Transit Time' = $TransitTime
    }

    for ($i = 1; $i -le 25; $i++) {
        $detail = $null
        if ($i -le (@($PriceDetails)).Count) {
            $detail = $PriceDetails[$i - 1]
        }

        $row["Price Detail $i"] = if ($detail) { $detail.Name } else { '' }
        $row["Price Detail $i - Currency"] = if ($detail) { $detail.Currency } else { '' }
        $row["Price Detail $i - Comment"] = if ($detail) { $detail.Comment } else { '' }
        $row["Price Detail $i - Evaluation"] = if ($detail) { $detail.Evaluation } else { '' }
        $row["Price Detail $i - Minimum"] = if ($detail) { $detail.Minimum } else { '' }
        $row["Price Detail $i - Maximum"] = if ($detail) { $detail.Maximum } else { '' }
        $row["Price Detail $i - Tiers Type"] = if ($detail) { $detail.TiersType } else { '' }
        $row["Price Detail $i - Price"] = if ($detail) { $detail.Price } else { '' }
    }

    return [pscustomobject]$row
}

function Convert-IndexToColumnName {
    param([int]$Index)

    $value = $Index
    $name = ''
    while ($value -gt 0) {
        $remainder = ($value - 1) % 26
        $name = ([string][char]([int](65 + $remainder))) + $name
        $value = [Math]::Floor(($value - 1) / 26)
    }

    return $name
}

function Escape-XmlText {
    param([string]$Text)
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-CellXml {
    param(
        [string]$Reference,
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrEmpty($text)) {
        return $null
    }

    if ($text -match '^-?\d+(\.\d+)?$') {
        return "<c r=""$Reference""><v>$text</v></c>"
    }

    $escaped = Escape-XmlText $text
    return "<c r=""$Reference"" t=""inlineStr""><is><t xml:space=""preserve"">$escaped</t></is></c>"
}

function Write-NormalizedWorkbook {
    param(
        [string]$OutputPath,
        [string[]]$Headers,
        [object[]]$DataRows
    )

    $tempRoot = New-TempDirectory
    try {
        $sheetPath = Join-Path $tempRoot 'xl\worksheets'
        $relsPath = Join-Path $tempRoot 'xl\_rels'
        $rootRelsPath = Join-Path $tempRoot '_rels'

        [System.IO.Directory]::CreateDirectory($sheetPath) | Out-Null
        [System.IO.Directory]::CreateDirectory($relsPath) | Out-Null
        [System.IO.Directory]::CreateDirectory($rootRelsPath) | Out-Null

        $rowXmlList = New-Object System.Collections.Generic.List[string]

        $headerCells = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $ref = "{0}1" -f (Convert-IndexToColumnName ($i + 1))
            $cellXml = New-CellXml $ref $Headers[$i]
            if ($cellXml) {
                $headerCells.Add($cellXml)
            }
        }
        $rowXmlList.Add("<row r=""1"">$($headerCells -join '')</row>")

        for ($rowIndex = 0; $rowIndex -lt $DataRows.Count; $rowIndex++) {
            $excelRow = $rowIndex + 2
            $cells = New-Object System.Collections.Generic.List[string]
            $data = $DataRows[$rowIndex].PSObject.Properties
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                $header = $Headers[$i]
                $value = $DataRows[$rowIndex].$header
                $ref = "{0}{1}" -f (Convert-IndexToColumnName ($i + 1)), $excelRow
                $cellXml = New-CellXml $ref $value
                if ($cellXml) {
                    $cells.Add($cellXml)
                }
            }
            $rowXmlList.Add("<row r=""$excelRow"">$($cells -join '')</row>")
        }

        $lastColumn = Convert-IndexToColumnName $Headers.Count
        $lastRow = $DataRows.Count + 1
        $sheetXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:$lastColumn$lastRow"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetData>
    $($rowXmlList -join "`n    ")
  </sheetData>
</worksheet>
"@

        $contentTypesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
"@

        $rootRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

        $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"@

        $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"@

        [System.IO.File]::WriteAllText((Join-Path $tempRoot '[Content_Types].xml'), $contentTypesXml, [System.Text.UTF8Encoding]::new($false))
        [System.IO.File]::WriteAllText((Join-Path $rootRelsPath '.rels'), $rootRelsXml, [System.Text.UTF8Encoding]::new($false))
        [System.IO.File]::WriteAllText((Join-Path $tempRoot 'xl\workbook.xml'), $workbookXml, [System.Text.UTF8Encoding]::new($false))
        [System.IO.File]::WriteAllText((Join-Path $relsPath 'workbook.xml.rels'), $workbookRelsXml, [System.Text.UTF8Encoding]::new($false))
        [System.IO.File]::WriteAllText((Join-Path $sheetPath 'sheet1.xml'), $sheetXml, [System.Text.UTF8Encoding]::new($false))

        if (Test-Path -LiteralPath $OutputPath) {
            Remove-Item -LiteralPath $OutputPath -Force
        }

        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempRoot, $OutputPath)
    }
    finally {
        Remove-DirectorySafe $tempRoot
    }
}

function Test-CoscoFarEastWorkbook {
    param([object[]]$WorksheetInfos)

    $normalizedNames = @($WorksheetInfos | ForEach-Object { Normalize-Key $_.Name })
    return ($normalizedNames -contains 'ADDIZIONALI') -and ($normalizedNames -contains 'POLS GENOVA-LA SPEZIA')
}

function Get-CoscoFarEastValidityWindow {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        $text = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if (-not $text) {
            continue
        }

        $plain = Remove-Diacritics $text
        $match = [regex]::Match($plain, 'validita.*?dal\s+(\d{1,2}/\d{1,2}/\d{2,4}).*?al\s+(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
        if (-not $match.Success) {
            continue
        }

        return [pscustomobject]@{
            Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
            End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
        }
    }

    throw 'Unable to extract validity dates from COSCO Far East workbook.'
}

function Get-CoscoFarEastOriginNames {
    param([string]$TitleText)

    $text = Normalize-Key $TitleText
    $text = $text -replace '^PORT OF LOADING\s+', ''
    $text = $text -replace '\s*&\s*', ','
    $origins = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($text -split '\s*,\s*')) {
        $value = Normalize-Whitespace $part
        if ($value) {
            Add-UniqueString -List $origins -Value $value
        }
    }

    return $origins.ToArray()
}

function Get-CoscoFarEastCountryHintFromSection {
    param([string]$SectionLabel)

    $label = Normalize-Key $SectionLabel
    switch -Regex ($label) {
        '^CAMBODIA BASE PORT' { return 'KH' }
        '^INDONESIA BASE PORT' { return 'ID' }
        '^JAPAN BASE PORT' { return 'JP' }
        '^MALAYSIA BASE PORT' { return 'MY' }
        '^MYANMAR BASE PORT' { return 'MM' }
        '^PHILIPPINES BASE PORT' { return 'PH' }
        '^SOUTH KOREA BASE PORT' { return 'KR' }
        '^TAIWAN BASE PORT' { return 'TW' }
        '^THAILAND BASE PORT' { return 'TH' }
        '^VIETNAM BASE PORT' { return 'VN' }
        '^CHINA ' { return 'CN' }
        default { return '' }
    }
}

function Get-CoscoFarEastCountryHintForDestination {
    param([string]$DestinationName)

    $sourceText = Normalize-LocationText $DestinationName
    $normalized = Normalize-Key (($sourceText -replace '\(.*?\)', ' ') -replace '\*.*$', ' ')
    switch -Regex ($normalized) {
        '^BUSAN(?:\s|$)' { return 'KR' }
        '^DALIAN(?:\s|$)' { return 'CN' }
        '^HONG KONG(?:\s|$)' { return 'HK' }
        '^KAOHSIUNG(?:\s|$)' { return 'TW' }
        '^NANSHA(?:\s|$)' { return 'CN' }
        '^NINGBO(?:\s|$)' { return 'CN' }
        '^PORT KELANG(?:\s|$)' { return 'MY' }
        '^QINGDAO(?:\s|$)' { return 'CN' }
        '^SHANGHAI(?:\s|$)' { return 'CN' }
        '^SHEKOU(?:\s|$)' { return 'CN' }
        '^SINGAPORE(?:\s|$)' { return 'SG' }
        '^XIAMEN(?:\s|$)' { return 'CN' }
        '^XINGANG(?:\s|$)' { return 'CN' }
        '^YANTIAN(?:\s|$)' { return 'CN' }
        default { return '' }
    }
}

function Normalize-CoscoFarEastDestination {
    param([string]$Text)

    $sourceText = Normalize-LocationText $Text
    $value = Normalize-Whitespace ((($sourceText -replace '\*.*$', ' ') -replace '\s+', ' '))
    return $value.Trim()
}

function Get-CoscoFarEastRateEntries {
    param([object[]]$Rows)

    $entries = @()
    $currentCountryHint = ''
    foreach ($row in $Rows) {
        $aValue = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        $bValue = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $cValue = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        $dValue = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
        $eValue = Normalize-Whitespace (Get-Cell $row.Cells 'E' $row.RowNumber)

        $sectionCountryHint = Get-CoscoFarEastCountryHintFromSection -SectionLabel $aValue
        if ($sectionCountryHint) {
            $currentCountryHint = $sectionCountryHint
            continue
        }

        if (-not $aValue -or $bValue -notmatch '^[A-Z]{3}$') {
            continue
        }

        if ($cValue -notmatch '^-?\d+(?:[.,]\d+)?$' -and $dValue -notmatch '^-?\d+(?:[.,]\d+)?$') {
            continue
        }

        $destination = Normalize-CoscoFarEastDestination -Text $aValue
        $countryHint = if ($currentCountryHint) {
            $currentCountryHint
        } else {
            Get-CoscoFarEastCountryHintForDestination -DestinationName $destination
        }

        $entries += [pscustomobject]@{
            Destination = $destination
            CountryHint = $countryHint
            Currency = $bValue.ToUpperInvariant()
            Rate20 = $cValue
            Rate40 = $dValue
            Notes = $eValue
        }
    }

    return $entries
}

function Get-CoscoFarEastAdditionalDetails {
    param(
        [object[]]$Rows,
        [hashtable]$Rules
    )

    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals (Get-ExpectedAdditionals -Rules $Rules -Carrier 'COSCO' -Direction 'EXPORT') -Rules $Rules
    $details = @()
    $currentSection = ''

    foreach ($row in $Rows) {
        $text = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if (-not $text) {
            continue
        }

        if ((Normalize-Key $text) -eq 'OCEANIA ETS') {
            $currentSection = 'OCEANIA ETS'
            continue
        }

        $normalizedText = Normalize-Key $text
        foreach ($definition in $definitions) {
            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedText -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            if ($normalizedText -match '\bINCLUDED\b') {
                continue
            }

            if ($definition.Name -eq 'ETS' -and $currentSection -eq 'OCEANIA ETS') {
                continue
            }

            $parsed = Parse-AdditionalText -Text $text -CanonicalName $definition.Name
            if ((@($parsed)).Count -eq 0) {
                continue
            }

            $applyTargets = @($definition.ApplyTargets)
            if ((@($applyTargets)).Count -eq 0) {
                $applyTargets = Get-AdditionalApplyTargets -CanonicalName $definition.Name -Comment $text
            }

            foreach ($detail in $parsed) {
                $details += [pscustomobject]@{
                    AppliesTo = @($applyTargets)
                    Detail = $detail
                }
            }
        }
    }

    return $details
}

function Convert-CoscoFarEastWorkbook {
    param(
        [string]$PackageRoot,
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }
    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO Far East workbook adapter expects carrier COSCO. Received '$Carrier'."
    }
    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO Far East workbook adapter expects Export direction. Received '$Direction'."
    }

    $additionalsSheet = $WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -eq 'ADDIZIONALI' } | Select-Object -First 1
    if (-not $additionalsSheet) {
        throw 'COSCO Far East workbook is missing the addizionali sheet.'
    }

    $additionalsRows = Get-WorksheetRows -WorksheetPath $additionalsSheet.Path -SharedStrings $SharedStrings
    $validity = Get-CoscoFarEastValidityWindow -Rows $additionalsRows
    $additionalTemplates = Get-CoscoFarEastAdditionalDetails -Rows $additionalsRows -Rules $Rules

    $sheetPlans = @()
    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($worksheetInfo in ($WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -ne 'ADDIZIONALI' })) {
        $rows = Get-WorksheetRows -WorksheetPath $worksheetInfo.Path -SharedStrings $SharedStrings
        $title = ''
        $titleRow = $rows | Where-Object { $_.RowNumber -eq 1 } | Select-Object -First 1
        if ($titleRow) {
            $title = Get-Cell $titleRow.Cells 'A' 1
        }

        $origins = Get-CoscoFarEastOriginNames -TitleText $title
        foreach ($origin in $origins) {
            Add-UniqueString -List $rawLocationNames -Value $origin
        }

        $entries = Get-CoscoFarEastRateEntries -Rows $rows
        foreach ($entry in $entries) {
            Add-UniqueString -List $rawLocationNames -Value $entry.Destination
        }

        $sheetPlans += [pscustomobject]@{
            Origins = $origins
            Entries = $entries
        }
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($sheetPlan in $sheetPlans) {
        foreach ($originName in $sheetPlan.Origins) {
            $originCodes = Get-LocationCodes -RawName $originName -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint 'IT'
            foreach ($entry in $sheetPlan.Entries) {
                $destinationCodes = Get-LocationCodes -RawName $entry.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint $entry.CountryHint
                foreach ($originCode in $originCodes) {
                    foreach ($destinationCode in $destinationCodes) {
                        $details = @()

                        if ($entry.Rate20 -match '^-?\d+(?:[.,]\d+)?$') {
                            $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency $entry.Notes "Cntr 20' Box" ($entry.Rate20.Replace(',', '.')))
                        }
                        if ($entry.Rate40 -match '^-?\d+(?:[.,]\d+)?$') {
                            $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency $entry.Notes "Cntr 40' Box" ($entry.Rate40.Replace(',', '.')))
                        }

                        foreach ($template in $additionalTemplates) {
                            if (Should-ApplyAdditionalToDestination -ApplyTargets $template.AppliesTo -DestinationName $entry.Destination) {
                                $details += $template.Detail
                            }
                        }

                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $details
                        $rowIndex++
                    }
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Convert-ExcelSerialDateToIsoString {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    if ($value -notmatch '^-?\d+(?:[.,]\d+)?$') {
        return ''
    }

    try {
        return [DateTime]::FromOADate([double]($value.Replace(',', '.'))).ToString('yyyy-MM-dd')
    } catch {
        return ''
    }
}

function Normalize-CoscoOriginAlias {
    param([string]$Text)

    $value = Normalize-Key (($Text -replace '\(.*?\)', ' ') -replace '\s+', ' ')
    switch -Regex ($value) {
        '^GOA$|^GENOVA$|^GENOA$' { return 'Genova' }
        '^SPE$|^LA SPEZIA$' { return 'La Spezia' }
        '^TRIESTE$' { return 'Trieste' }
        '^NAPOLI$|^NAPLES$' { return 'Napoli' }
        '^RAVENNA$' { return 'Ravenna' }
        '^VENEZIA$|^VENICE$' { return 'Venezia' }
        '^ANCONA$' { return 'Ancona' }
        '^BARI$' { return 'Bari' }
        '^AUGUSTA$' { return 'Augusta' }
        default {
            return (Normalize-Whitespace (($Text -replace '\(.*?\)', ' ') -replace '\s+', ' '))
        }
    }
}

function Split-CoscoOriginField {
    param([string]$Text)

    $clean = Normalize-Whitespace (($Text -replace '\(.*?\)', ' ') -replace '\s+', ' ')
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($clean -split '\s*/\s*')) {
        $value = Normalize-CoscoOriginAlias $part
        if ($value) {
            Add-UniqueString -List $items -Value $value
        }
    }

    return $items.ToArray()
}

function Split-CoscoDestinationField {
    param([string]$Text)

    $clean = Normalize-Whitespace (($Text -replace '\s+', ' '))
    if ($clean -match 'NORTH/SOUTH') {
        return @($clean)
    }

    $items = New-Object System.Collections.Generic.List[string]
    foreach ($part in ($clean -split '\s*/\s*')) {
        $value = Normalize-Whitespace $part
        if ($value) {
            Add-UniqueString -List $items -Value $value
        }
    }

    return $items.ToArray()
}

function Get-TransitTimeValue {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    $match = [regex]::Match($value, '(\d{1,3})')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return ''
}

function Test-CoscoStructuredWorkbook {
    param([object[]]$WorksheetInfos)

    $normalizedNames = @($WorksheetInfos | ForEach-Object { Normalize-Key $_.Name })
    return ($normalizedNames -contains 'TARIFFE') -and ($normalizedNames -contains 'SOVRAPPREZZI') -and ($normalizedNames -contains 'POL_ADDON')
}

function Get-CoscoStructuredWorkbookValidityWindow {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        foreach ($column in @('A', 'B', 'C', 'D', 'E')) {
            $text = Normalize-Whitespace (Get-Cell $row.Cells $column $row.RowNumber)
            if (-not $text) {
                continue
            }

            $plain = Remove-Diacritics $text
            $match = [regex]::Match($plain, '(\d{1,2}/\d{1,2}/\d{2,4})\s*-\s*(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
            if ($match.Success) {
                return [pscustomobject]@{
                    Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
                    End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
                }
            }
        }
    }

    $validFromRow = $Rows | Where-Object { $_.RowNumber -eq 6 } | Select-Object -First 1
    if ($validFromRow) {
        $start = Convert-ExcelSerialDateToIsoString (Get-Cell $validFromRow.Cells 'B' 6)
        $end = Convert-ExcelSerialDateToIsoString (Get-Cell $validFromRow.Cells 'E' 6)
        if ($start -and $end) {
            return [pscustomobject]@{
                Start = $start
                End = $end
            }
        }
    }

    throw 'Unable to extract validity dates from COSCO structured workbook.'
}

function Get-CoscoStructuredWorkbookReference {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        foreach ($column in @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H')) {
            $text = Normalize-Whitespace (Get-Cell $row.Cells $column $row.RowNumber)
            if (-not $text) {
                continue
            }

            $match = [regex]::Match($text, '([A-Z]{2,}\d{4}[A-Z]\d+\s*/\s*AF\s*\d+|FM\d+[A-Z]\d+\s*/\s*AF\s*\d+)', 'IgnoreCase')
            if ($match.Success) {
                return (Normalize-Whitespace $match.Groups[1].Value)
            }
        }
    }

    return ''
}

function Get-CoscoStructuredWorkbookAddOnRules {
    param([object[]]$Rows)

    $rules = @()
    foreach ($row in $Rows) {
        $origin = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        $amount = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $baseRate = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        if (-not $origin -or $amount -notmatch '^-?\d+(?:[.,]\d+)?$') {
            continue
        }

        $baseOrigins = Split-CoscoOriginField -Text $baseRate
        if ((@($baseOrigins)).Count -eq 0) {
            continue
        }

        $rules += [pscustomobject]@{
            Origin = (Normalize-CoscoOriginAlias $origin)
            BaseOrigins = @($baseOrigins)
            AddOn = (Convert-LocalizedNumberText $amount)
        }
    }

    return ,@($rules)
}

function Get-CoscoStructuredWorkbookRateEntries {
    param(
        [object[]]$Rows,
        [object[]]$AddOnRules
    )

    $entries = @()
    $headerRowNumber = 0
    foreach ($row in $Rows) {
        $aValue = Normalize-Key (Get-Cell $row.Cells 'A' $row.RowNumber)
        $bValue = Normalize-Key (Get-Cell $row.Cells 'B' $row.RowNumber)
        if ($aValue -eq 'RATE ID' -and $bValue -eq 'PORT OF LOADING') {
            $headerRowNumber = $row.RowNumber
            break
        }
    }

    if (-not $headerRowNumber) {
        throw 'Unable to locate tariff header row in COSCO structured workbook.'
    }

    foreach ($row in ($Rows | Where-Object { $_.RowNumber -gt $headerRowNumber })) {
        $rateId = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        $originField = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $destinationField = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        if ($rateId -notmatch '^\d+$' -or -not $originField -or -not $destinationField) {
            continue
        }

        $baseOrigins = @(Split-CoscoOriginField -Text $originField)
        $destinations = @(Split-CoscoDestinationField -Text $destinationField)
        $entry = [pscustomobject]@{
            Origins = $baseOrigins
            Destinations = $destinations
            Transshipment = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
            TransitTime = Get-TransitTimeValue (Get-Cell $row.Cells 'E' $row.RowNumber)
            Currency = 'USD'
            Rate20 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'F' $row.RowNumber)
            Rate40 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'G' $row.RowNumber)
            Reference = Normalize-Whitespace (Get-Cell $row.Cells 'H' $row.RowNumber)
            Comment = ''
        }
        $entries += $entry

        foreach ($rule in $AddOnRules) {
            $baseIntersection = @($rule.BaseOrigins | Where-Object { $baseOrigins -contains $_ })
            if ($baseIntersection.Count -eq 0 -or $baseOrigins -contains $rule.Origin) {
                continue
            }

            $entries += [pscustomobject]@{
                Origins = @($rule.Origin)
                Destinations = $destinations
                Transshipment = $entry.Transshipment
                TransitTime = $entry.TransitTime
                Currency = $entry.Currency
                Rate20 = if ($entry.Rate20) { [string]([double]$entry.Rate20 + [double]$rule.AddOn) } else { '' }
                Rate40 = if ($entry.Rate40) { [string]([double]$entry.Rate40 + [double]$rule.AddOn) } else { '' }
                Reference = $entry.Reference
                Comment = ("Derived from {0} base rate with POL add-on {1} USD/CNTR." -f ($baseIntersection -join '/'), $rule.AddOn)
            }
        }
    }

    return $entries
}

function Get-CoscoStructuredWorkbookAdditionalTemplates {
    param(
        [object[]]$Rows,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals (Get-ExpectedAdditionals -Rules $Rules -Carrier $Carrier -Direction $Direction) -Rules $Rules
    $details = @()

    foreach ($row in $Rows) {
        $category = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        $scope = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $amount = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        $currency = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
        $unit = Normalize-Whitespace (Get-Cell $row.Cells 'E' $row.RowNumber)
        $note = Normalize-Whitespace (Get-Cell $row.Cells 'F' $row.RowNumber)
        $effective = Normalize-Whitespace (Get-Cell $row.Cells 'G' $row.RowNumber)

        if (-not $category) {
            continue
        }

        $combinedText = Normalize-Whitespace (($category, $scope, $note) -join ' ')
        $normalizedText = Normalize-Key $combinedText
        if ($normalizedText -match '\bINCLUDED\b') {
            continue
        }
        if ($amount -notmatch '^-?\d+(?:[.,]\d+)?$') {
            continue
        }

        foreach ($definition in $definitions) {
            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedText -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            $commentParts = @($scope, $note)
            if ($effective) {
                $commentParts += $effective
            }
            $comment = Normalize-Whitespace (($commentParts | Where-Object { $_ }) -join ' | ')
            $applyTargets = @($definition.ApplyTargets)
            if ((@($applyTargets)).Count -eq 0) {
                $applyTargets = @(Get-AdditionalApplyTargets -CanonicalName $definition.Name -Comment $combinedText)
            }

            $details += [pscustomobject]@{
                AppliesTo = @($applyTargets)
                Detail = (New-PriceDetail $definition.Name (Get-CurrencyCode $currency) $comment (Get-EvaluationFromUnitToken $unit) (Convert-LocalizedNumberText $amount))
            }
            break
        }
    }

    return $details
}

function Get-PreferredLocationCodes {
    param(
        [string]$RawName,
        [string]$ExplicitCode,
        [hashtable]$Rules,
        [hashtable]$UnlocodeLookup = $null,
        [string]$CountryHint = ''
    )

    if ($RawName) {
        try {
            return @(Get-LocationCodes -RawName $RawName -Rules $Rules -UnlocodeLookup $UnlocodeLookup -CountryHint $CountryHint)
        } catch {
        }
    }

    if ($ExplicitCode -match '^[A-Z]{5}$') {
        return @($ExplicitCode.ToUpperInvariant())
    }

    if ($RawName) {
        throw "UNLOCODE mapping not configured for '$RawName'."
    }

    return @()
}

function Resolve-OptionalTransshipmentAddress {
    param(
        [string]$Text,
        [hashtable]$Rules,
        [hashtable]$UnlocodeLookup = $null
    )

    $value = Normalize-Whitespace $Text
    if (-not $value -or $value -eq '-') {
        return ''
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    Add-UniqueString -List $candidates -Value $value
    Add-UniqueString -List $candidates -Value (Normalize-Whitespace ($value -replace '\s+and\s+.*$', ''))
    Add-UniqueString -List $candidates -Value (Normalize-Whitespace (($value -split '\s*\+\s*')[0]))
    Add-UniqueString -List $candidates -Value (Normalize-Whitespace (($value -split '\s*,\s*')[0]))
    Add-UniqueString -List $candidates -Value (Normalize-Whitespace (($value -split '\s*/\s*')[0]))

    foreach ($candidate in $candidates) {
        if (-not $candidate) {
            continue
        }

        try {
            return (@(Get-LocationCodes -RawName $candidate -Rules $Rules -UnlocodeLookup $UnlocodeLookup))[0]
        } catch {
        }
    }

    return ''
}

function Convert-CoscoStructuredEntriesToWorkbook {
    param(
        [object[]]$Entries,
        [object[]]$AdditionalTemplates,
        [pscustomobject]$Validity,
        [string]$OutputPath,
        [string]$Carrier,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $Entries) {
        foreach ($origin in @($entry.Origins)) {
            Add-UniqueString -List $rawLocationNames -Value $origin
        }
        foreach ($destination in @($entry.Destinations)) {
            Add-UniqueString -List $rawLocationNames -Value $destination
        }
        Add-UniqueString -List $rawLocationNames -Value $entry.Transshipment
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($entry in $Entries) {
        $transshipmentAddress = Resolve-OptionalTransshipmentAddress -Text $entry.Transshipment -Rules $Rules -UnlocodeLookup $unlocodeLookup
        foreach ($originName in @($entry.Origins)) {
            $originCodes = @(Get-LocationCodes -RawName $originName -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint 'IT')
            foreach ($destinationName in @($entry.Destinations)) {
                $countryHint = if ($entry.PSObject.Properties.Name -contains 'CountryHint') { $entry.CountryHint } else { '' }
                $destinationCodeHint = if ($entry.PSObject.Properties.Name -contains 'DestinationCodeHint') { $entry.DestinationCodeHint } else { '' }
                $destinationCodes = @(Get-PreferredLocationCodes -RawName $destinationName -ExplicitCode $destinationCodeHint -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint $countryHint)
                foreach ($originCode in $originCodes) {
                    foreach ($destinationCode in $destinationCodes) {
                        $details = @()
                        if ($entry.Rate20) {
                            $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency '' "Cntr 20' Box" $entry.Rate20)
                        }
                        if ($entry.Rate40) {
                            $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency '' "Cntr 40' Box" $entry.Rate40)
                        }
                        if ($entry.PSObject.Properties.Name -contains 'Rate40H' -and $entry.Rate40H) {
                            $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency '' "Cntr 40' HC" $entry.Rate40H)
                        }

                        foreach ($template in $AdditionalTemplates) {
                            if (Should-ApplyAdditionalToDestination -ApplyTargets $template.AppliesTo -DestinationName $destinationName) {
                                $details += $template.Detail
                            }
                        }

                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $Validity.Start -ValidityEnd $Validity.End -Carrier $Carrier -PriceDetails $details -TransitTime $entry.TransitTime -Reference $entry.Reference -TransshipmentAddress $transshipmentAddress -Comment $entry.Comment
                        $rowIndex++
                    }
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Convert-CoscoStructuredWorkbook {
    param(
        [string]$PackageRoot,
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }
    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO structured workbook adapter expects carrier COSCO. Received '$Carrier'."
    }
    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO structured workbook adapter expects Export direction. Received '$Direction'."
    }

    $tariffeSheet = $WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -eq 'TARIFFE' } | Select-Object -First 1
    $sovrapprezziSheet = $WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -eq 'SOVRAPPREZZI' } | Select-Object -First 1
    $polAddOnSheet = $WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -eq 'POL_ADDON' } | Select-Object -First 1
    if (-not $tariffeSheet -or -not $sovrapprezziSheet -or -not $polAddOnSheet) {
        throw 'COSCO structured workbook is missing one or more required sheets.'
    }

    $tariffeRows = Get-WorksheetRows -WorksheetPath $tariffeSheet.Path -SharedStrings $SharedStrings
    $sovrapprezziRows = Get-WorksheetRows -WorksheetPath $sovrapprezziSheet.Path -SharedStrings $SharedStrings
    $polAddOnRows = Get-WorksheetRows -WorksheetPath $polAddOnSheet.Path -SharedStrings $SharedStrings

    $validity = Get-CoscoStructuredWorkbookValidityWindow -Rows $tariffeRows
    $reference = Get-CoscoStructuredWorkbookReference -Rows $tariffeRows
    $entries = Get-CoscoStructuredWorkbookRateEntries -Rows $tariffeRows -AddOnRules (Get-CoscoStructuredWorkbookAddOnRules -Rows $polAddOnRows)
    foreach ($entry in $entries) {
        if (-not $entry.Reference) {
            $entry.Reference = $reference
        }
    }
    $additionalTemplates = Get-CoscoStructuredWorkbookAdditionalTemplates -Rows $sovrapprezziRows -Carrier $normalizedCarrier -Direction $normalizedDirection -Rules $Rules

    Convert-CoscoStructuredEntriesToWorkbook -Entries $entries -AdditionalTemplates $additionalTemplates -Validity $validity -OutputPath $OutputPath -Carrier $normalizedCarrier -Rules $Rules -UnlocodePath $UnlocodePath
}

function Test-CoscoIpakPdfText {
    param([string]$Text)

    $normalized = Normalize-Key $Text
    return (
        $normalized -match 'OFFICIAL RATE REFERENCE TO BE SHOWN ON BOOKING REQUEST' -and
        $normalized -match 'GENOVA \(VTE\)' -and
        $normalized -match 'NOTES AND SURCHARGES' -and
        ($normalized -match 'ON TOP OF GENOA RATE' -or $normalized -match '\+\s*USD\s*50\s*/\s*CNTR')
    )
}

function Get-CoscoIpakPdfValidityWindow {
    param([string]$PdfText)

    $match = [regex]::Match($PdfText, 'Validity:\s*from\s+(\d{1,2}/\d{1,2}/\d{2,4})\s+to\s+(\d{1,2}/\d{1,2}/\d{2,4})', 'IgnoreCase')
    if (-not $match.Success) {
        throw 'Unable to extract validity dates from COSCO IPAK PDF.'
    }

    return [pscustomobject]@{
        Start = (Parse-ListinoDate $match.Groups[1].Value).ToString('yyyy-MM-dd')
        End = (Parse-ListinoDate $match.Groups[2].Value).ToString('yyyy-MM-dd')
    }
}

function Get-CoscoIpakPdfReference {
    param([string]$PdfText)

    $match = [regex]::Match($PdfText, 'Official rate reference to be shown on booking request:\s*(.+)$', 'IgnoreCase,Multiline')
    if (-not $match.Success) {
        return ''
    }

    return (Normalize-Whitespace $match.Groups[1].Value)
}

function Get-CoscoIpakPdfEntries {
    param([string]$PdfText)

    $entries = @()
    $blockMatch = [regex]::Match($PdfText, 'Port of Loading.+?Notes and Surcharges', 'Singleline,IgnoreCase')
    if (-not $blockMatch.Success) {
        throw 'Unable to locate rate table in COSCO IPAK PDF.'
    }

    $reference = Get-CoscoIpakPdfReference -PdfText $PdfText
    foreach ($line in ($blockMatch.Value -split "`r?`n")) {
        $normalizedLine = Normalize-Whitespace $line
        if ($normalizedLine -notmatch '^Genova \(VTE\) ') {
            continue
        }

        $match = [regex]::Match($normalizedLine, '^Genova \(VTE\)\s+(?<destination>.+?)\s+(?<transshipment>Singapore|-)\s+About\s+(?<days>\d+)\s+days\s+(?<rate20>\d+(?:[.,]\d+)?)\s+USD\s+(?<rate40>\d+(?:[.,]\d+)?)\s+USD$', 'IgnoreCase')
        if (-not $match.Success) {
            continue
        }

        $entries += [pscustomobject]@{
            Origins = @('Genova')
            Destinations = @(Split-CoscoDestinationField -Text $match.Groups['destination'].Value)
            DestinationCodeHint = ''
            CountryHint = ''
            Transshipment = $match.Groups['transshipment'].Value
            TransitTime = $match.Groups['days'].Value
            Currency = 'USD'
            Rate20 = (Convert-LocalizedNumberText $match.Groups['rate20'].Value)
            Rate40 = (Convert-LocalizedNumberText $match.Groups['rate40'].Value)
            Rate40H = ''
            Reference = $reference
            Comment = ''
        }
    }

    $addOnOrigins = @('Ancona', 'Venezia', 'Bari', 'Napoli', 'Ravenna', 'La Spezia')
    $expandedEntries = @()
    foreach ($entry in $entries) {
        $expandedEntries += $entry
        foreach ($origin in $addOnOrigins) {
            $expandedEntries += [pscustomobject]@{
                Origins = @($origin)
                Destinations = $entry.Destinations
                DestinationCodeHint = ''
                CountryHint = ''
                Transshipment = $entry.Transshipment
                TransitTime = $entry.TransitTime
                Currency = $entry.Currency
                Rate20 = [string]([double]$entry.Rate20 + 50)
                Rate40 = [string]([double]$entry.Rate40 + 50)
                Rate40H = ''
                Reference = $entry.Reference
                Comment = 'Derived from Genoa base rate with PDF add-on 50 USD/CNTR.'
            }
        }
    }

    return ,@($expandedEntries)
}

function Get-CoscoIpakPdfAdditionalTemplates {
    param(
        [string]$PdfText,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $definitions = Get-AdditionalDefinitions -ExpectedAdditionals (Get-ExpectedAdditionals -Rules $Rules -Carrier $Carrier -Direction $Direction) -Rules $Rules
    $details = @()
    foreach ($line in ($PdfText -split "`r?`n")) {
        $normalizedLine = Normalize-Key $line
        if (-not $normalizedLine -or $normalizedLine -match '\bINCLUDED\b') {
            continue
        }

        foreach ($definition in $definitions) {
            $matched = $false
            foreach ($pattern in $definition.Patterns) {
                if ($normalizedLine -match $pattern) {
                    $matched = $true
                    break
                }
            }

            if (-not $matched) {
                continue
            }

            $parsed = @(Parse-AdditionalText -Text $line -CanonicalName $definition.Name)
            foreach ($detail in $parsed) {
                $details += [pscustomobject]@{
                    AppliesTo = @($definition.ApplyTargets)
                    Detail = $detail
                }
            }
            break
        }
    }

    return $details
}

function Convert-CoscoIpakPdfToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = '',
        [string]$PdfText = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }
    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO IPAK PDF adapter expects carrier COSCO. Received '$Carrier'."
    }
    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO IPAK PDF adapter expects Export direction. Received '$Direction'."
    }

    if (-not $PdfText) {
        $PdfText = Get-PdfText -InputPath $InputPath -Mode 'raw'
    }

    if (-not (Test-CoscoIpakPdfText -Text $PdfText)) {
        throw 'COSCO IPAK PDF markers not found in the PDF text.'
    }

    $validity = Get-CoscoIpakPdfValidityWindow -PdfText $PdfText
    $entries = Get-CoscoIpakPdfEntries -PdfText $PdfText
    $additionalTemplates = Get-CoscoIpakPdfAdditionalTemplates -PdfText $PdfText -Carrier $normalizedCarrier -Direction $normalizedDirection -Rules $Rules
    Convert-CoscoStructuredEntriesToWorkbook -Entries $entries -AdditionalTemplates $additionalTemplates -Validity $validity -OutputPath $OutputPath -Carrier $normalizedCarrier -Rules $Rules -UnlocodePath $UnlocodePath
}

function Test-EvergreenRvsWorkbook {
    param(
        [string]$InputPath,
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings
    )

    $normalizedNames = @($WorksheetInfos | ForEach-Object { Normalize-Key $_.Name })
    if (-not ($normalizedNames -contains 'FAK')) {
        return $false
    }

    if ((Normalize-Key ([System.IO.Path]::GetFileName($InputPath))) -like '*FAK RVS*') {
        return $true
    }

    $combined = Normalize-Key ($SharedStrings -join ' ')
    return ($combined -match 'RATES ARE SUBJECT TO ISOCC' -and $combined -match 'GENOVA' -and $combined -match 'LA SPEZIA')
}

function Get-EvergreenRvsValidityWindow {
    param(
        [string]$InputPath,
        [string[]]$SharedStrings
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $match = [regex]::Match($fileName, '(\d{8})\s*-\s*(\d{8})')
    if ($match.Success) {
        return [pscustomobject]@{
            Start = ([DateTime]::ParseExact($match.Groups[1].Value, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
            End = ([DateTime]::ParseExact($match.Groups[2].Value, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
        }
    }

    $combined = $SharedStrings -join ' '
    $textMatch = [regex]::Match($combined, 'Validity:\s*([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4}).+?([A-Za-z]{3}\s+\d{1,2}(?:st|nd|rd|th),\s+\d{4})', 'IgnoreCase')
    if ($textMatch.Success) {
        $startText = ($textMatch.Groups[1].Value -replace '(st|nd|rd|th),', ',')
        $endText = ($textMatch.Groups[2].Value -replace '(st|nd|rd|th),', ',')
        return [pscustomobject]@{
            Start = ([DateTime]::ParseExact($startText, 'MMM d, yyyy', [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
            End = ([DateTime]::ParseExact($endText, 'MMM d, yyyy', [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
        }
    }

    throw 'Unable to extract validity dates from Evergreen RVS workbook.'
}

function Get-EvergreenRvsOriginBlocks {
    param([object[]]$Rows)

    $row1 = $Rows | Where-Object { $_.RowNumber -eq 1 } | Select-Object -First 1
    $row2 = $Rows | Where-Object { $_.RowNumber -eq 2 } | Select-Object -First 1
    if (-not $row1 -or -not $row2) {
        throw 'Evergreen RVS workbook is missing header rows.'
    }

    $blocks = @()
    for ($index = 4; $index -le 24; $index += 3) {
        $originName = Normalize-Whitespace (Get-Cell $row1.Cells (Convert-IndexToColumnName $index) 1)
        if (-not $originName) {
            continue
        }

        $blocks += [pscustomobject]@{
            Origin = (Normalize-CoscoOriginAlias $originName)
            PriceColumns = @(
                [pscustomobject]@{ Column = (Convert-IndexToColumnName $index); Evaluation = (Get-EvaluationFromUnitToken (Get-Cell $row2.Cells (Convert-IndexToColumnName $index) 2)) },
                [pscustomobject]@{ Column = (Convert-IndexToColumnName ($index + 1)); Evaluation = (Get-EvaluationFromUnitToken (Get-Cell $row2.Cells (Convert-IndexToColumnName ($index + 1)) 2)) },
                [pscustomobject]@{ Column = (Convert-IndexToColumnName ($index + 2)); Evaluation = (Get-EvaluationFromUnitToken (Get-Cell $row2.Cells (Convert-IndexToColumnName ($index + 2)) 2)) }
            )
        }
    }

    return ,@($blocks)
}

function Get-EvergreenRvsAdditionalTemplates {
    param([string[]]$SharedStrings)

    $combined = $SharedStrings -join ' '
    $templates = @()

    $isoccMatch = [regex]::Match($combined, 'ISOCC\s+Usd\s+(\d+(?:[.,]\d+)?)\/teu', 'IgnoreCase')
    if ($isoccMatch.Success) {
        $templates += [pscustomobject]@{ Name = 'ISOCC'; Currency = 'USD'; Price = (Convert-LocalizedNumberText $isoccMatch.Groups[1].Value); ExcludeChina = $false }
    }

    $euisMatch = [regex]::Match($combined, 'EUIS\s+Eur\s+(\d+(?:[.,]\d+)?)\/teu', 'IgnoreCase')
    if ($euisMatch.Success) {
        $templates += [pscustomobject]@{ Name = 'EUIS'; Currency = 'EUR'; Price = (Convert-LocalizedNumberText $euisMatch.Groups[1].Value); ExcludeChina = $true }
    }

    $lssMatch = [regex]::Match($combined, 'LSS\s+Usd\s+(\d+(?:[.,]\d+)?)\/Teu', 'IgnoreCase')
    if ($lssMatch.Success) {
        $templates += [pscustomobject]@{ Name = 'LSS'; Currency = 'USD'; Price = (Convert-LocalizedNumberText $lssMatch.Groups[1].Value); ExcludeChina = $true }
    }

    return ,@($templates)
}

function Convert-EvergreenRvsWorkbook {
    param(
        [string]$InputPath,
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $sheetInfo = $WorksheetInfos | Where-Object { (Normalize-Key $_.Name) -eq 'FAK' } | Select-Object -First 1
    if (-not $sheetInfo) {
        throw 'Evergreen RVS workbook is missing the FAK sheet.'
    }

    $rows = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
    $validity = Get-EvergreenRvsValidityWindow -InputPath $InputPath -SharedStrings $SharedStrings
    $originBlocks = Get-EvergreenRvsOriginBlocks -Rows $rows
    $additionalTemplates = Get-EvergreenRvsAdditionalTemplates -SharedStrings $SharedStrings

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($block in $originBlocks) {
        Add-UniqueString -List $rawLocationNames -Value $block.Origin
    }
    foreach ($row in $rows) {
        if ((Get-Cell $row.Cells 'B' $row.RowNumber) -match '^[A-Z]{5}$') {
            Add-UniqueString -List $rawLocationNames -Value (Get-Cell $row.Cells 'C' $row.RowNumber)
        }
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($row in $rows | Where-Object { (Get-Cell $_.Cells 'B' $_.RowNumber) -match '^[A-Z]{5}$' }) {
        $destinationCodeHint = (Get-Cell $row.Cells 'B' $row.RowNumber)
        $destinationName = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        $resolvedDestinationCodes = @(Get-PreferredLocationCodes -RawName $destinationName -ExplicitCode $destinationCodeHint -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint $destinationCodeHint.Substring(0,2))

        foreach ($originBlock in $originBlocks) {
            $originCodes = @(Get-LocationCodes -RawName $originBlock.Origin -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint 'IT')
            $details = @()
            foreach ($priceColumn in $originBlock.PriceColumns) {
                $rateValue = Convert-LocalizedNumberText (Get-Cell $row.Cells $priceColumn.Column $row.RowNumber)
                if ($rateValue -match '^-?\d+(?:[.,]\d+)?$') {
                    $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' 'USD' '' $priceColumn.Evaluation $rateValue)
                }
            }

            foreach ($template in $additionalTemplates) {
                if ($template.ExcludeChina) {
                    if ($destinationCodeHint.StartsWith('CN') -and -not $destinationCodeHint.StartsWith('HK') -and -not $destinationCodeHint.StartsWith('TW')) {
                        continue
                    }
                }

                $details += (New-PriceDetail $template.Name $template.Currency '' 'TEUS' $template.Price)
            }

            foreach ($originCode in $originCodes) {
                foreach ($destinationCode in $resolvedDestinationCodes) {
                    $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier 'EVERGREEN' -PriceDetails $details -Reference ([System.IO.Path]::GetFileNameWithoutExtension($InputPath))
                    $rowIndex++
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Test-CoscoIetWorkbook {
    param(
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings
    )

    $normalizedNames = @($WorksheetInfos | ForEach-Object { Normalize-Key $_.Name })
    $hasExpectedSheets = ($normalizedNames -contains 'GVA -SPE- VEN-RAV') -and ($normalizedNames -contains 'NAPOLI - ANCONA') -and ($normalizedNames -contains 'BARI - AUGUSTA')
    if (-not $hasExpectedSheets) {
        return $false
    }

    $combined = Normalize-Key ($SharedStrings -join ' ')
    return ($combined -match 'COSCO INTRAMED' -and $combined -match 'INTRAEUROPE RATES')
}

function Get-CoscoIetReference {
    param([string[]]$SharedStrings)

    $combined = $SharedStrings -join ' '
    $match = [regex]::Match($combined, 'ref:\s*([A-Z0-9-]+)', 'IgnoreCase')
    if ($match.Success) {
        return (Normalize-Whitespace $match.Groups[1].Value)
    }

    return ''
}

function Get-CoscoIetValidityWindow {
    param(
        [string]$InputPath,
        [string[]]$SharedStrings,
        [string]$Reference
    )

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $startMatch = [regex]::Match($fileName, 'dal\s+(\d{2})-(\d{2})-(\d{4})', 'IgnoreCase')
    if (-not $startMatch.Success) {
        throw 'Unable to extract IET validity start date from file name.'
    }

    $start = [DateTime]::ParseExact(('{0}-{1}-{2}' -f $startMatch.Groups[1].Value, $startMatch.Groups[2].Value, $startMatch.Groups[3].Value), 'dd-MM-yyyy', [Globalization.CultureInfo]::InvariantCulture)
    $endText = ''

    if ($Reference -match '(\d{8})$') {
        $endText = $matches[1]
    } else {
        $combined = $SharedStrings -join ' '
        $endMatch = [regex]::Match($combined, 'ref:[A-Z0-9-]+-(\d{8})', 'IgnoreCase')
        if ($endMatch.Success) {
            $endText = $endMatch.Groups[1].Value
        }
    }

    if (-not $endText) {
        throw 'Unable to extract IET validity end date from reference.'
    }

    $end = [DateTime]::ParseExact($endText, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)
    return [pscustomobject]@{
        Start = $start.ToString('yyyy-MM-dd')
        End = $end.ToString('yyyy-MM-dd')
    }
}

function Get-CoscoIetOriginsFromTitle {
    param([string]$Title)

    $text = Normalize-Whitespace $Title
    $match = [regex]::Match($text, 'POL\s+(.+?)(?:\s*\(|$)', 'IgnoreCase')
    if (-not $match.Success) {
        throw "Unable to extract origin group from IET title '$Title'."
    }

    $origins = New-Object System.Collections.Generic.List[string]
    foreach ($part in ([regex]::Split($match.Groups[1].Value, '\s*&\s*'))) {
        $value = Normalize-CoscoOriginAlias $part
        if ($value) {
            Add-UniqueString -List $origins -Value $value
        }
    }

    return $origins.ToArray()
}

function Add-CoscoIetDestinationCodeHint {
    param(
        [hashtable]$Map,
        [string]$Name,
        [string]$Code
    )

    $normalizedCode = Normalize-Key $Code
    if ($normalizedCode -notmatch '^[A-Z]{5}$') {
        return
    }

    foreach ($key in @(
        (Normalize-Key $Name),
        (Normalize-Key (Normalize-Whitespace (($Name -replace '\(.*?\)', ' ') -replace '\s+', ' ')))
    )) {
        if ($key -and (-not $Map.ContainsKey($key))) {
            $Map[$key] = $normalizedCode
        }
    }
}

function Get-CoscoIetDestinationCodeHint {
    param(
        [hashtable]$Map,
        [string]$Name
    )

    foreach ($key in @(
        (Normalize-Key $Name),
        (Normalize-Key (Normalize-Whitespace (($Name -replace '\(.*?\)', ' ') -replace '\s+', ' ')))
    )) {
        if ($key -and $Map.ContainsKey($key)) {
            return $Map[$key]
        }
    }

    return ''
}

function Normalize-CoscoIetViaText {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    if (-not $value) {
        return ''
    }

    $value = Normalize-Whitespace ($value -replace '^\s*via\s+', '')
    if ((Normalize-Key $value) -match '^(DIRECT|SUSPENDED.*)$') {
        return ''
    }

    return $value
}

function Normalize-CoscoIetDestinationName {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    if (-not $value) {
        return ''
    }

    $value = $value -replace '\(\s*(?:Pol|Fm)\b.*?\)', ''
    $value = $value -replace '\s+SOLO\s+DA.*$', ''
    $value = $value -replace '\s+FM\s+.*$', ''
    return (Normalize-Whitespace $value)
}

function Get-CoscoIetCountryHintFromSection {
    param([string]$Section)

    $text = Normalize-Key $Section
    switch -Regex ($text) {
        'ALBANIA' { return 'AL' }
        'ALGERIA' { return 'DZ' }
        'BELGIO|BELGIUM' { return 'BE' }
        'BULGARIA' { return 'BG' }
        'CIPRO|CYPRUS' { return 'CY' }
        'DANIMARCA|DENMARK' { return 'DK' }
        'EGITTO|EGYPT' { return 'EG' }
        'ESTONIA' { return 'EE' }
        'FINLANDIA|FINLAND' { return 'FI' }
        'FRANCIA|FRANCE' { return 'FR' }
        'GEORGIA' { return 'GE' }
        'GERMANIA|GERMANY' { return 'DE' }
        'GRAN BRETAGNA|UNITED KINGDOM|GREAT BRITAIN' { return 'GB' }
        'GRECIA|GREECE' { return 'GR' }
        'IRLANDA|IRELAND' { return 'IE' }
        'ISRAELE|ISRAEL' { return 'IL' }
        'LIBANO|LEBANON' { return 'LB' }
        'LIBIA|LIBYA' { return 'LY' }
        'LETTONIA|LATVIA' { return 'LV' }
        'LITUANIA|LITHUANIA' { return 'LT' }
        'MALTA' { return 'MT' }
        'MAROCCO|MOROCCO' { return 'MA' }
        'NORVEGIA|NORWAY' { return 'NO' }
        'OLANDA|HOLLAND|NETHERLANDS' { return 'NL' }
        'POLONIA|POLAND' { return 'PL' }
        'PORTUGAL' { return 'PT' }
        'ROMANIA' { return 'RO' }
        'SPAGNA|SPAIN' { return 'ES' }
        'SVEZIA|SWEDEN' { return 'SE' }
        'TURCHIA|TURKEY' { return 'TR' }
        default { return '' }
    }
}

function Get-CoscoIetHalfEntries {
    param(
        [object[]]$Rows,
        [string[]]$Origins,
        [string]$NameColumn,
        [string]$CodeColumn,
        [string]$ViaColumn,
        [string]$CurrencyColumn,
        [string]$Rate20Column,
        [string]$Rate40Column,
        [string]$ServiceColumn,
        [hashtable]$NameCodeHints
    )

    $entries = @()
    $currentSection = ''
    foreach ($row in ($Rows | Where-Object { $_.RowNumber -ge 17 })) {
        $name = Normalize-CoscoIetDestinationName (Get-Cell $row.Cells $NameColumn $row.RowNumber)
        if (-not $name) {
            continue
        }

        $codeHint = if ($CodeColumn) { Normalize-Whitespace (Get-Cell $row.Cells $CodeColumn $row.RowNumber) } else { '' }
        $via = if ($ViaColumn) { Normalize-Whitespace (Get-Cell $row.Cells $ViaColumn $row.RowNumber) } else { '' }
        $currencyToken = if ($CurrencyColumn) { Normalize-Whitespace (Get-Cell $row.Cells $CurrencyColumn $row.RowNumber) } else { '' }
        $rate20 = if ($Rate20Column) { Convert-LocalizedNumberText (Get-Cell $row.Cells $Rate20Column $row.RowNumber) } else { '' }
        $rate40 = if ($Rate40Column) { Convert-LocalizedNumberText (Get-Cell $row.Cells $Rate40Column $row.RowNumber) } else { '' }
        $service = if ($ServiceColumn) { Normalize-Whitespace (Get-Cell $row.Cells $ServiceColumn $row.RowNumber) } else { '' }

        $hasRates = ($rate20 -match '^-?\d+(?:\.\d+)?$') -or ($rate40 -match '^-?\d+(?:\.\d+)?$')
        if (-not $hasRates) {
            if (-not $codeHint -and -not $currencyToken -and -not $service) {
                $currentSection = $name
            }
            continue
        }

        $combinedMarker = Normalize-Key (($name, $via, $service) -join ' ')
        if ($combinedMarker -match 'SUSPENDED') {
            continue
        }

        if ($codeHint) {
            Add-CoscoIetDestinationCodeHint -Map $NameCodeHints -Name $name -Code $codeHint
        } else {
            $codeHint = Get-CoscoIetDestinationCodeHint -Map $NameCodeHints -Name $name
        }

        $commentParts = @()
        if ($currentSection) {
            $commentParts += $currentSection
        }
        if ($service) {
            $commentParts += ("Service: {0}" -f $service)
        }

        $entries += [pscustomobject]@{
            Origins = @($Origins)
            DestinationName = $name
            DestinationCodeHint = $codeHint
            CountryHint = (Get-CoscoIetCountryHintFromSection $currentSection)
            Transshipment = (Normalize-CoscoIetViaText $via)
            Currency = (Get-CurrencyCode $currencyToken)
            Rate20 = if ($rate20 -match '^-?\d+(?:\.\d+)?$') { $rate20 } else { '' }
            Rate40 = if ($rate40 -match '^-?\d+(?:\.\d+)?$') { $rate40 } else { '' }
            Comment = (Normalize-Whitespace ($commentParts -join ' | '))
        }
    }

    return $entries
}

function Get-CoscoIetEntries {
    param(
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos
    )

    $rowsBySheet = @{}
    foreach ($sheetInfo in $WorksheetInfos) {
        $rowsBySheet[$sheetInfo.Name] = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
    }

    $nameCodeHints = @{}
    foreach ($sheetName in $rowsBySheet.Keys) {
        foreach ($row in $rowsBySheet[$sheetName]) {
            $name = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
            $code = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
            if ($name -and $code) {
                Add-CoscoIetDestinationCodeHint -Map $nameCodeHints -Name $name -Code $code
            }
        }
    }

    $entries = @()
    foreach ($sheetInfo in $WorksheetInfos) {
        $rows = $rowsBySheet[$sheetInfo.Name]
        $titleRow = $rows | Where-Object { $_.RowNumber -eq 12 } | Select-Object -First 1
        if (-not $titleRow) {
            continue
        }

        $leftTitle = Normalize-Whitespace (Get-Cell $titleRow.Cells 'A' 12)
        $rightTitle = Normalize-Whitespace (Get-Cell $titleRow.Cells 'I' 12)
        if ($leftTitle) {
            $entries += Get-CoscoIetHalfEntries -Rows $rows -Origins (Get-CoscoIetOriginsFromTitle -Title $leftTitle) -NameColumn 'A' -CodeColumn 'B' -ViaColumn 'C' -CurrencyColumn 'D' -Rate20Column 'E' -Rate40Column 'F' -ServiceColumn 'G' -NameCodeHints $nameCodeHints
        }
        if ($rightTitle) {
            $entries += Get-CoscoIetHalfEntries -Rows $rows -Origins (Get-CoscoIetOriginsFromTitle -Title $rightTitle) -NameColumn 'I' -CodeColumn '' -ViaColumn 'J' -CurrencyColumn 'K' -Rate20Column 'L' -Rate40Column 'M' -ServiceColumn 'N' -NameCodeHints $nameCodeHints
        }
    }

    return $entries
}

function Get-CoscoIetAdditionalDetails {
    param(
        [string]$DestinationCode,
        [string]$DestinationName
    )

    $details = @()
    $normalizedCode = Normalize-Key $DestinationCode
    $normalizedName = Normalize-Key $DestinationName

    $details += (New-PriceDetail 'FAF' 'EUR' 'March 2026 intramed rate sheet.' 'TEUS' '107')

    $ets59Prefixes = @('AL', 'GR', 'MA', 'IL', 'EG', 'DZ', 'TR', 'LB', 'MT', 'CY', 'ES', 'RO', 'BG', 'GE')
    $ets87Prefixes = @('GB', 'DE', 'BE', 'NL', 'NO', 'SE', 'IE', 'DK', 'FI', 'PT', 'PL', 'LT', 'LV', 'EE')

    $etsValue = ''
    if ($normalizedCode -eq 'FRFOS') {
        $etsValue = '59'
    } elseif ($normalizedCode.Length -ge 2 -and ($ets59Prefixes -contains $normalizedCode.Substring(0,2))) {
        $etsValue = '59'
    } elseif ($normalizedCode.Length -ge 2 -and ($ets87Prefixes -contains $normalizedCode.Substring(0,2))) {
        $etsValue = '87'
    } elseif ($normalizedCode -like 'FR*' -or $normalizedName -like '*FRANCE*') {
        $etsValue = '87'
    }

    if ($etsValue) {
        $details += (New-PriceDetail 'ETS' 'EUR' 'March 2026 ETS surcharge per destination cluster.' 'TEUS' $etsValue)
    }

    if ($normalizedCode -like 'IL*') {
        $details += (New-PriceDetail 'WRS ISRAEL' 'USD' 'Israel prepaid basis.' 'TEUS' '100')
    }

    if ($normalizedCode -eq 'TRMER') {
        $details += (New-PriceDetail 'TAC (MERSIN)' 'USD' 'Terminal Additional Charge for Mersin.' 'TEUS' '31.5')
    }

    if ($normalizedCode -eq 'MACAS') {
        $details += (New-PriceDetail 'PCS' 'EUR' 'Casablanca port congestion surcharge.' 'TEUS' '50')
        $details += (New-PriceDetail 'ORS (CASABLANCA)' 'EUR' 'Casablanca operational recovery charge.' 'TEUS' '100')
    }

    return $details
}

function Convert-CoscoIetWorkbook {
    param(
        [string]$InputPath,
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'COSCO' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }
    if ($normalizedCarrier -ne 'COSCO') {
        throw "COSCO IET workbook adapter expects carrier COSCO. Received '$Carrier'."
    }
    if ($normalizedDirection -ne 'EXPORT') {
        throw "COSCO IET workbook adapter expects Export direction. Received '$Direction'."
    }

    $reference = Get-CoscoIetReference -SharedStrings $SharedStrings
    $validity = Get-CoscoIetValidityWindow -InputPath $InputPath -SharedStrings $SharedStrings -Reference $reference
    $entries = Get-CoscoIetEntries -SharedStrings $SharedStrings -WorksheetInfos $WorksheetInfos

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $entries) {
        foreach ($origin in @($entry.Origins)) {
            Add-UniqueString -List $rawLocationNames -Value $origin
        }
        Add-UniqueString -List $rawLocationNames -Value $entry.DestinationName
        Add-UniqueString -List $rawLocationNames -Value $entry.Transshipment
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($entry in $entries) {
        $transshipmentAddress = Resolve-OptionalTransshipmentAddress -Text $entry.Transshipment -Rules $Rules -UnlocodeLookup $unlocodeLookup
        foreach ($originName in @($entry.Origins)) {
            $originCodes = @(Get-LocationCodes -RawName $originName -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint 'IT')
            $destinationCountryHint = if ($entry.DestinationCodeHint) { $entry.DestinationCodeHint.Substring(0, [Math]::Min(2, $entry.DestinationCodeHint.Length)) } else { $entry.CountryHint }
            $destinationCodes = @(Get-PreferredLocationCodes -RawName $entry.DestinationName -ExplicitCode $entry.DestinationCodeHint -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint $destinationCountryHint)

            foreach ($originCode in $originCodes) {
                foreach ($destinationCode in $destinationCodes) {
                    $details = @()
                    if ($entry.Rate20) {
                        $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency '' "Cntr 20' Box" $entry.Rate20)
                    }
                    if ($entry.Rate40) {
                        $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $entry.Currency '' "Cntr 40' Box" $entry.Rate40)
                    }
                    $details += Get-CoscoIetAdditionalDetails -DestinationCode $destinationCode -DestinationName $entry.DestinationName

                    $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $details -Reference $reference -TransshipmentAddress $transshipmentAddress -Comment $entry.Comment
                    $rowIndex++
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Test-HmmCitWorkbook {
    param(
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings
    )

    $requiredSheets = @('Head', 'Freight', 'Arb Addon', 'Subject to')
    $sheetNames = @($WorksheetInfos | ForEach-Object { $_.Name })
    foreach ($requiredSheet in $requiredSheets) {
        if (-not ($sheetNames -contains $requiredSheet)) {
            return $false
        }
    }

    return (@($SharedStrings | Where-Object { $_ -eq 'Contract Number' }).Count -gt 0)
}

function Get-HmmCitWorksheetRows {
    param(
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings,
        [string]$SheetName
    )

    $sheetInfo = $WorksheetInfos | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
    if (-not $sheetInfo) {
        throw "Worksheet '$SheetName' not found in HMM CIT workbook."
    }

    return Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
}

function Convert-HmmCitIsoDateToString {
    param([string]$Text)

    $value = Normalize-Whitespace $Text
    if (-not $value) {
        return ''
    }

    return ([DateTime]::ParseExact($value, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).ToString('yyyy-MM-dd')
}

function Get-HmmCitHeadInfo {
    param([object[]]$Rows)

    $row = $Rows | Where-Object { $_.RowNumber -eq 3 } | Select-Object -First 1
    if (-not $row) {
        throw 'Unable to read Head sheet row 3 from HMM CIT workbook.'
    }

    $contractNumber = Normalize-Whitespace (Get-Cell $row.Cells 'A' 3)
    $amendNumber = Normalize-Whitespace (Get-Cell $row.Cells 'B' 3)
    $effectiveDate = Convert-HmmCitIsoDateToString (Get-Cell $row.Cells 'C' 3)
    $contractStart = Convert-HmmCitIsoDateToString (Get-Cell $row.Cells 'D' 3)
    $contractEnd = Convert-HmmCitIsoDateToString (Get-Cell $row.Cells 'E' 3)
    $reference = if ($contractNumber -and $amendNumber) {
        "$contractNumber-amd$amendNumber"
    } elseif ($contractNumber) {
        $contractNumber
    } else {
        ''
    }

    $coverageList = @()
    foreach ($coverage in @((Get-Cell $row.Cells 'I' 3) -split '\s*,\s*')) {
        $normalizedCoverage = Normalize-Key $coverage
        if ($normalizedCoverage) {
            $coverageList += $normalizedCoverage
        }
    }

    return [pscustomobject]@{
        ContractNumber = $contractNumber
        AmendNumber = $amendNumber
        Reference = $reference
        Start = if ($effectiveDate) { $effectiveDate } else { $contractStart }
        End = $contractEnd
        CoverageList = @($coverageList | Select-Object -Unique)
    }
}

function Split-HmmCitCodeList {
    param([string]$Text)

    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($segment in @($Text -split '\s*,\s*')) {
        $code = Normalize-Key $segment
        if ($code -match '^[A-Z]{5}$' -and (-not $codes.Contains($code))) {
            $codes.Add($code)
        }
    }

    return $codes.ToArray()
}

function Get-HmmCitFreightEntries {
    param([object[]]$Rows)

    $entries = @()
    $currentCoverage = ''
    $currentBulletSeq = ''
    $currentCommodity = ''
    $currentCurrency = ''

    foreach ($row in ($Rows | Where-Object { $_.RowNumber -gt 2 } | Sort-Object RowNumber)) {
        $coverage = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if ($coverage) {
            $currentCoverage = Normalize-Key $coverage
        }

        $bulletSeq = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        if ($bulletSeq) {
            $currentBulletSeq = $bulletSeq
        }

        $commodity = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        if ($commodity) {
            $currentCommodity = $commodity
        }

        $currency = Get-CurrencyCode (Get-Cell $row.Cells 'K' $row.RowNumber)
        if ($currency) {
            $currentCurrency = $currency
        }

        $originCodes = @(Split-HmmCitCodeList (Get-Cell $row.Cells 'E' $row.RowNumber))
        $destinationCodes = @(Split-HmmCitCodeList (Get-Cell $row.Cells 'G' $row.RowNumber))
        if (-not $currentCoverage -or $originCodes.Count -eq 0 -or $destinationCodes.Count -eq 0) {
            continue
        }

        $entries += [pscustomobject]@{
            Coverage = $currentCoverage
            BulletSeq = $currentBulletSeq
            Commodity = $currentCommodity
            Loop = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
            OriginCodes = $originCodes
            DestinationCodes = $destinationCodes
            OriginTerm = Normalize-Whitespace (Get-Cell $row.Cells 'F' $row.RowNumber)
            DestinationTerm = Normalize-Whitespace (Get-Cell $row.Cells 'H' $row.RowNumber)
            ContainerType = Normalize-Whitespace (Get-Cell $row.Cells 'I' $row.RowNumber)
            CargoType = Normalize-Whitespace (Get-Cell $row.Cells 'J' $row.RowNumber)
            Currency = $currentCurrency
            Rate20 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'L' $row.RowNumber)
            Rate40 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'M' $row.RowNumber)
            Rate40H = Convert-LocalizedNumberText (Get-Cell $row.Cells 'N' $row.RowNumber)
        }
    }

    return $entries
}

function Get-HmmCitArbEntries {
    param([object[]]$Rows)

    $entries = @()
    $currentCoverage = ''

    foreach ($row in ($Rows | Where-Object { $_.RowNumber -gt 2 } | Sort-Object RowNumber)) {
        $coverage = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if ($coverage) {
            $currentCoverage = Normalize-Key $coverage
        }

        $item = Normalize-Key (Get-Cell $row.Cells 'B' $row.RowNumber)
        if (-not $currentCoverage -or -not $item) {
            continue
        }

        $outportCode = Normalize-Key (Get-Cell $row.Cells 'C' $row.RowNumber)
        $baseRateCode = Normalize-Key (Get-Cell $row.Cells 'F' $row.RowNumber)
        if ($outportCode -notmatch '^[A-Z]{5}$' -or $baseRateCode -notmatch '^[A-Z]{5}$') {
            continue
        }

        $entries += [pscustomobject]@{
            Coverage = $currentCoverage
            Item = $item
            OutportCode = $outportCode
            OutportDescription = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
            ServiceTerm = Normalize-Whitespace (Get-Cell $row.Cells 'E' $row.RowNumber)
            BaseRateCode = $baseRateCode
            BaseRateDescription = Normalize-Whitespace (Get-Cell $row.Cells 'G' $row.RowNumber)
            ViaCode = Normalize-Key (Get-Cell $row.Cells 'H' $row.RowNumber)
            DirectFlag = Normalize-Key (Get-Cell $row.Cells 'I' $row.RowNumber)
            Loop = Normalize-Whitespace (Get-Cell $row.Cells 'J' $row.RowNumber)
            Currency = Get-CurrencyCode (Get-Cell $row.Cells 'M' $row.RowNumber)
            Rate20 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'N' $row.RowNumber)
            Rate40 = Convert-LocalizedNumberText (Get-Cell $row.Cells 'O' $row.RowNumber)
            Rate40H = Convert-LocalizedNumberText (Get-Cell $row.Cells 'P' $row.RowNumber)
        }
    }

    return $entries
}

function Get-HmmCitSubjectContentsByCoverage {
    param([object[]]$Rows)

    $contentsByCoverage = @{}
    $currentCoverage = ''

    foreach ($row in ($Rows | Where-Object { $_.RowNumber -gt 2 } | Sort-Object RowNumber)) {
        $coverage = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        if ($coverage) {
            $currentCoverage = Normalize-Key $coverage
        }

        $content = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        if (-not $currentCoverage -or -not $content) {
            continue
        }

        if (-not $contentsByCoverage.ContainsKey($currentCoverage)) {
            $contentsByCoverage[$currentCoverage] = New-Object System.Collections.Generic.List[string]
        }
        $contentsByCoverage[$currentCoverage].Add($content)
    }

    return $contentsByCoverage
}

function Get-HmmCitAdditionalCurrencyFromText {
    param([string]$Text)

    $normalized = Normalize-Key $Text
    if ($normalized -match '\bEUR\b') {
        return 'EUR'
    }
    if ($Text.Contains('$') -or $normalized -match '\bUSD\b') {
        return 'USD'
    }
    return ''
}

function Add-HmmCitTemplateIfMissing {
    param(
        [System.Collections.Generic.List[object]]$Templates,
        [hashtable]$Seen,
        [string]$Name,
        [string]$Currency,
        [string]$Comment,
        [string]$Evaluation,
        [string]$Price,
        [string]$AppliesTo
    )

    if (-not $Name -or -not $Currency -or -not $Evaluation -or -not $Price) {
        return
    }

    $key = "$Name|$Currency|$Evaluation|$Price|$AppliesTo"
    if ($Seen.ContainsKey($key)) {
        return
    }

    $Seen[$key] = $true
    $Templates.Add([pscustomobject]@{
        Name = $Name
        Currency = $Currency
        Comment = $Comment
        Evaluation = $Evaluation
        Price = $Price
        AppliesTo = $AppliesTo
    })
}

function Get-HmmCitAdditionalTemplatesByCoverage {
    param(
        [hashtable]$SubjectContentsByCoverage,
        [hashtable]$Rules
    )

    $allowedAdditionals = @(Get-ExpectedAdditionals -Rules $Rules -Carrier 'HMM' -Direction 'IMPORT')
    $templatesByCoverage = @{}

    foreach ($coverage in $SubjectContentsByCoverage.Keys) {
        $templates = New-Object System.Collections.Generic.List[object]
        $seen = @{}

        foreach ($content in @($SubjectContentsByCoverage[$coverage])) {
            $comment = Normalize-Whitespace $content
            if (-not $comment) {
                continue
            }

            if (($allowedAdditionals -contains 'ECC') -and $comment -match '\bECC\b\s*\$?(\d+(?:[.,]\d+)?)\s*/T') {
                Add-HmmCitTemplateIfMissing -Templates $templates -Seen $seen -Name 'ECC' -Currency (Get-HmmCitAdditionalCurrencyFromText $comment) -Comment $comment -Evaluation 'TEUS' -Price (Convert-LocalizedNumberText $matches[1]) -AppliesTo '*'
            }

            if (($allowedAdditionals -contains 'STF') -and $comment -match '\bSTF\b\s*:?\s*\$?(\d+(?:[.,]\d+)?)\s*/T') {
                Add-HmmCitTemplateIfMissing -Templates $templates -Seen $seen -Name 'STF' -Currency (Get-HmmCitAdditionalCurrencyFromText $comment) -Comment $comment -Evaluation 'TEUS' -Price (Convert-LocalizedNumberText $matches[1]) -AppliesTo '*'
            }

            if (($allowedAdditionals -contains 'ETS') -and $comment -match 'Q1 EES \(ETS\).*?\+\$?(\d+(?:[.,]\d+)?)\s*/\s*\+\$?(\d+(?:[.,]\d+)?).*?ORIGIN\s*=\s*CHINA') {
                Add-HmmCitTemplateIfMissing -Templates $templates -Seen $seen -Name 'ETS' -Currency 'USD' -Comment $comment -Evaluation 'TEUS' -Price (Convert-LocalizedNumberText $matches[1]) -AppliesTo 'CN'
            } elseif (($allowedAdditionals -contains 'ETS') -and $comment -match 'Q1 EES \(ETS\):\s*EUR\s+(\d+(?:[.,]\d+)?)\s*/\s*(\d+(?:[.,]\d+)?)') {
                Add-HmmCitTemplateIfMissing -Templates $templates -Seen $seen -Name 'ETS' -Currency 'EUR' -Comment $comment -Evaluation 'TEUS' -Price (Convert-LocalizedNumberText $matches[1]) -AppliesTo 'NON_CN'
            }
        }

        $templatesByCoverage[$coverage] = $templates.ToArray()
    }

    return $templatesByCoverage
}

function Get-HmmCitAdditionalDetails {
    param(
        [string]$Coverage,
        [string]$OriginCode,
        [hashtable]$TemplatesByCoverage
    )

    $coverageKey = Normalize-Key $Coverage
    if (-not $TemplatesByCoverage.ContainsKey($coverageKey)) {
        return @()
    }

    $originPrefix = if ($OriginCode -and $OriginCode.Length -ge 2) {
        $OriginCode.Substring(0, 2).ToUpperInvariant()
    } else {
        ''
    }

    $details = @()
    foreach ($template in @($TemplatesByCoverage[$coverageKey])) {
        $shouldApply = $false
        switch ($template.AppliesTo) {
            '*' { $shouldApply = $true }
            'CN' { $shouldApply = ($originPrefix -eq 'CN') }
            'NON_CN' { $shouldApply = ($originPrefix -and $originPrefix -ne 'CN') }
            default { $shouldApply = ($originPrefix -eq $template.AppliesTo) }
        }

        if ($shouldApply) {
            $details += (New-PriceDetail $template.Name $template.Currency $template.Comment $template.Evaluation $template.Price)
        }
    }

    return $details
}

function Get-HmmCitOceanFreightDetails {
    param([pscustomobject]$Entry)

    $details = @()
    if ($Entry.Rate20) {
        $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $Entry.Currency '' "Cntr 20' Box" $Entry.Rate20)
    }
    if ($Entry.Rate40) {
        $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $Entry.Currency '' "Cntr 40' Box" $Entry.Rate40)
    }
    if ($Entry.Rate40H) {
        $details += (New-PriceDetail 'OCEAN FREIGHT - CONTAINERS' $Entry.Currency '' "Cntr 40' HC" $Entry.Rate40H)
    }

    return $details
}

function Get-HmmCitArbitraryDetails {
    param([pscustomobject]$Entry)

    $details = @()
    $itemComment = if ($Entry.Item -eq 'DEST ARB') { 'Destination arbitrary' } else { 'Origin arbitrary' }
    $commentParts = New-Object System.Collections.Generic.List[string]
    $commentParts.Add($itemComment)
    if ($Entry.OutportDescription) {
        $commentParts.Add($Entry.OutportDescription)
    }
    if ($Entry.BaseRateDescription) {
        $commentParts.Add("base $($Entry.BaseRateDescription)")
    } elseif ($Entry.BaseRateCode) {
        $commentParts.Add("base $($Entry.BaseRateCode)")
    }
    if ($Entry.ViaCode -match '^[A-Z]{5}$') {
        $commentParts.Add("via $($Entry.ViaCode)")
    }
    $comment = Normalize-Whitespace (($commentParts -join ' | '))

    if ($Entry.Rate20) {
        $details += (New-PriceDetail 'INLAND FREIGHT' $Entry.Currency $comment "Cntr 20' Box" $Entry.Rate20)
    }
    if ($Entry.Rate40) {
        $details += (New-PriceDetail 'INLAND FREIGHT' $Entry.Currency $comment "Cntr 40' Box" $Entry.Rate40)
    }
    if ($Entry.Rate40H) {
        $details += (New-PriceDetail 'INLAND FREIGHT' $Entry.Currency $comment "Cntr 40' HC" $Entry.Rate40H)
    }

    return $details
}

function Get-HmmCitTransshipmentCode {
    param([pscustomobject]$Entry)

    if ($Entry.ViaCode -match '^[A-Z]{5}$') {
        return $Entry.ViaCode
    }

    if ($Entry.BaseRateCode -match '^[A-Z]{5}$') {
        return $Entry.BaseRateCode
    }

    return ''
}

function Get-HmmCitRouteComment {
    param(
        [pscustomobject]$FreightEntry,
        [string]$Item = ''
    )

    $parts = New-Object System.Collections.Generic.List[string]
    if ($FreightEntry.Coverage) {
        $parts.Add("Coverage $($FreightEntry.Coverage)")
    }
    if ($FreightEntry.Loop) {
        $parts.Add("Loop $($FreightEntry.Loop)")
    }
    if ($Item) {
        $parts.Add($Item)
    }

    return Normalize-Whitespace (($parts -join ' | '))
}

function Convert-HmmCitWorkbook {
    param(
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'HMM' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'IMPORT' }

    if ($normalizedCarrier -ne 'HMM') {
        throw "HMM CIT workbook adapter expects carrier HMM. Received '$Carrier'."
    }
    if ($normalizedDirection -ne 'IMPORT') {
        throw "HMM CIT workbook adapter expects Import direction. Received '$Direction'."
    }

    $headRows = Get-HmmCitWorksheetRows -WorksheetInfos $WorksheetInfos -SharedStrings $SharedStrings -SheetName 'Head'
    $freightRows = Get-HmmCitWorksheetRows -WorksheetInfos $WorksheetInfos -SharedStrings $SharedStrings -SheetName 'Freight'
    $arbRows = Get-HmmCitWorksheetRows -WorksheetInfos $WorksheetInfos -SharedStrings $SharedStrings -SheetName 'Arb Addon'
    $subjectRows = Get-HmmCitWorksheetRows -WorksheetInfos $WorksheetInfos -SharedStrings $SharedStrings -SheetName 'Subject to'

    $headInfo = Get-HmmCitHeadInfo -Rows $headRows
    $freightEntries = @(Get-HmmCitFreightEntries -Rows $freightRows)
    $arbEntries = @(Get-HmmCitArbEntries -Rows $arbRows)
    $templatesByCoverage = Get-HmmCitAdditionalTemplatesByCoverage -SubjectContentsByCoverage (Get-HmmCitSubjectContentsByCoverage -Rows $subjectRows) -Rules $Rules

    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($freightEntry in $freightEntries) {
        foreach ($originCode in @($freightEntry.OriginCodes)) {
            $details = @()
            $details += Get-HmmCitOceanFreightDetails -Entry $freightEntry
            $details += Get-HmmCitAdditionalDetails -Coverage $freightEntry.Coverage -OriginCode $originCode -TemplatesByCoverage $templatesByCoverage

            foreach ($destinationCode in @($freightEntry.DestinationCodes)) {
                $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $headInfo.Start -ValidityEnd $headInfo.End -Carrier $normalizedCarrier -PriceDetails $details -Reference $headInfo.Reference -Comment (Get-HmmCitRouteComment -FreightEntry $freightEntry)
                $rowIndex++
            }
        }
    }

    foreach ($arbEntry in $arbEntries) {
        switch ($arbEntry.Item) {
            'ORIGIN ARB' {
                $matchingFreightEntries = @($freightEntries | Where-Object { $_.Coverage -eq $arbEntry.Coverage -and (@($_.OriginCodes) -contains $arbEntry.BaseRateCode) })
                foreach ($freightEntry in $matchingFreightEntries) {
                    $details = @()
                    $details += Get-HmmCitOceanFreightDetails -Entry $freightEntry
                    $details += Get-HmmCitArbitraryDetails -Entry $arbEntry
                    $details += Get-HmmCitAdditionalDetails -Coverage $freightEntry.Coverage -OriginCode $arbEntry.OutportCode -TemplatesByCoverage $templatesByCoverage

                    foreach ($destinationCode in @($freightEntry.DestinationCodes)) {
                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $arbEntry.OutportCode -ToAddress $destinationCode -ValidityStart $headInfo.Start -ValidityEnd $headInfo.End -Carrier $normalizedCarrier -PriceDetails $details -Reference $headInfo.Reference -TransshipmentAddress (Get-HmmCitTransshipmentCode -Entry $arbEntry) -Comment (Get-HmmCitRouteComment -FreightEntry $freightEntry -Item 'Origin Arb')
                        $rowIndex++
                    }
                }
            }
            'DEST ARB' {
                $matchingFreightEntries = @($freightEntries | Where-Object { $_.Coverage -eq $arbEntry.Coverage -and (@($_.DestinationCodes) -contains $arbEntry.BaseRateCode) })
                foreach ($freightEntry in $matchingFreightEntries) {
                    foreach ($originCode in @($freightEntry.OriginCodes)) {
                        $details = @()
                        $details += Get-HmmCitOceanFreightDetails -Entry $freightEntry
                        $details += Get-HmmCitArbitraryDetails -Entry $arbEntry
                        $details += Get-HmmCitAdditionalDetails -Coverage $freightEntry.Coverage -OriginCode $originCode -TemplatesByCoverage $templatesByCoverage

                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $arbEntry.OutportCode -ValidityStart $headInfo.Start -ValidityEnd $headInfo.End -Carrier $normalizedCarrier -PriceDetails $details -Reference $headInfo.Reference -TransshipmentAddress (Get-HmmCitTransshipmentCode -Entry $arbEntry) -Comment (Get-HmmCitRouteComment -FreightEntry $freightEntry -Item 'Dest Arb')
                        $rowIndex++
                    }
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Convert-ColumnNameToIndex {
    param([string]$ColumnName)

    $name = Normalize-Key $ColumnName
    $index = 0
    foreach ($char in $name.ToCharArray()) {
        if ($char -lt 'A' -or $char -gt 'Z') {
            continue
        }
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }
    return $index
}

function Get-BaselineWorkbookTargetSheetInfos {
    param([object[]]$WorksheetInfos)

    $targetNames = @(
        'ISRAEL',
        'ISRAEL REEFER',
        'USA',
        'USA REEFER',
        'CANADA',
        'CANADA REEFER',
        'MEXICO',
        'CARIBB'
    )

    return @(
        $WorksheetInfos |
            Where-Object { $_.State -eq 'visible' -and ($targetNames -contains (Normalize-Key $_.Name)) }
    )
}

function Find-BaselineSheetValidityRow {
    param([object[]]$Rows)

    return ($Rows | Where-Object {
            (Normalize-Key (Get-Cell $_.Cells 'B' $_.RowNumber)) -match '^VALIDITY:?$' -and
            (Normalize-Whitespace (Get-Cell $_.Cells 'C' $_.RowNumber))
        } | Select-Object -First 1)
}

function Find-BaselineSheetPolRow {
    param([object[]]$Rows)

    return ($Rows | Where-Object {
            (Normalize-Key (Get-Cell $_.Cells 'B' $_.RowNumber)) -eq 'POL' -and
            (Normalize-Key (Get-Cell $_.Cells 'C' $_.RowNumber)) -eq 'POD' -and
            (Normalize-Key (Get-Cell $_.Cells 'D' $_.RowNumber)) -like 'CURRENCY*'
        } | Select-Object -First 1)
}

function Find-BaselineSheetAddizionaliRow {
    param(
        [object[]]$Rows,
        [int]$AfterRowNumber = 0
    )

    return ($Rows | Where-Object {
            $_.RowNumber -gt $AfterRowNumber -and
            (Normalize-Key (Get-Cell $_.Cells 'B' $_.RowNumber)) -eq 'ADDIZIONALI'
        } | Select-Object -First 1)
}

function Parse-BaselineSheetDate {
    param([string]$Text)

    $formats = @('dd.MM.yyyy', 'd.M.yyyy', 'dd/MM/yyyy', 'd/M/yyyy')
    foreach ($format in $formats) {
        try {
            return [DateTime]::ParseExact((Normalize-Whitespace $Text), $format, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
        }
    }

    throw "Unable to parse Baseline workbook date '$Text'."
}

function Format-BaselineWorkbookOutputDateText {
    param([string]$DateText)

    if (-not $DateText) {
        return ''
    }

    $formats = @('yyyy-MM-dd', 'yyyy-M-d', 'dd/MM/yyyy', 'd/M/yyyy', 'dd.MM.yyyy', 'd.M.yyyy')
    foreach ($format in $formats) {
        try {
            return [DateTime]::ParseExact((Normalize-Whitespace $DateText), $format, [System.Globalization.CultureInfo]::InvariantCulture).ToString('dd/MM/yyyy')
        } catch {
        }
    }

    throw "Unable to format Baseline workbook date '$DateText'."
}

function Get-BaselineSheetValidityWindow {
    param(
        [object[]]$Rows,
        [string]$ExplicitValidityStartDate = ''
    )

    $validityRow = Find-BaselineSheetValidityRow -Rows $Rows
    if (-not $validityRow) {
        throw 'Baseline workbook validity row not found.'
    }

    $dateText = Normalize-Whitespace (Get-Cell $validityRow.Cells 'C' $validityRow.RowNumber)
    $date = Parse-BaselineSheetDate $dateText
    return [pscustomobject]@{
        Start = if ($ExplicitValidityStartDate) { Format-BaselineWorkbookOutputDateText $ExplicitValidityStartDate } else { '' }
        End = $date.ToString('dd/MM/yyyy')
    }
}

function Get-BaselineSheetCurrency {
    param([object]$HeaderRow)

    $headerText = Normalize-Key (Get-Cell $HeaderRow.Cells 'D' $HeaderRow.RowNumber)
    if ($headerText -match '\bEUR\b' -or $headerText.Contains(([string][char]0x20AC))) {
        return 'EUR'
    }
    return 'USD'
}

function Get-BaselineContainerEvaluation {
    param([string]$Token)

    $normalized = Normalize-Key ($Token -replace '''', '')
    switch -Regex ($normalized) {
        '^20(?:DV|GP|BOX)?$' { return "Cntr 20' Box" }
        '^40(?:DV|GP|BOX)$' { return "Cntr 40' Box" }
        '^40(?:HC|HQ|H)$' { return "Cntr 40' HC" }
        '^20(?:RH|RE)$' { return "Cntr 20' Reefer" }
        '^40(?:RH|RE)$' { return "Cntr 40' Reefer" }
        default { return (Get-EvaluationFromUnitToken $Token) }
    }
}

function Get-BaselineSheetPriceColumns {
    param([object[]]$Rows)

    $polRow = Find-BaselineSheetPolRow -Rows $Rows
    if (-not $polRow) {
        throw 'Baseline workbook POL/POD header row not found.'
    }

    $equipmentRow = $Rows | Where-Object { $_.RowNumber -eq ($polRow.RowNumber + 1) } | Select-Object -First 1
    if (-not $equipmentRow) {
        throw 'Baseline workbook equipment row not found.'
    }

    $priceColumns = @()
    foreach ($cellRef in ($equipmentRow.Cells.Keys | Sort-Object { Convert-ColumnNameToIndex ($_ -replace '\d+', '') })) {
        $column = ($cellRef -replace '\d+', '')
        if ((Convert-ColumnNameToIndex $column) -lt (Convert-ColumnNameToIndex 'D')) {
            continue
        }

        $token = Normalize-Whitespace (Get-Cell $equipmentRow.Cells $column $equipmentRow.RowNumber)
        if (-not $token) {
            continue
        }

        $priceColumns += [pscustomobject]@{
            Column = $column
            Evaluation = Get-BaselineContainerEvaluation $token
            Token = $token
        }
    }

    if ((@($priceColumns)).Count -eq 0) {
        throw 'Baseline workbook equipment columns not found.'
    }

    return [pscustomobject]@{
        HeaderRow = $polRow
        EquipmentRow = $equipmentRow
        PriceColumns = @($priceColumns)
    }
}

function Test-BaselineWorkbookSheetLayout {
    param([object[]]$Rows)

    try {
        $validityRow = Find-BaselineSheetValidityRow -Rows $Rows
        $sheetColumns = Get-BaselineSheetPriceColumns -Rows $Rows
        $addizionaliRow = Find-BaselineSheetAddizionaliRow -Rows $Rows -AfterRowNumber $sheetColumns.EquipmentRow.RowNumber
        if (-not $validityRow -or -not $addizionaliRow) {
            return $false
        }

        [void](Parse-BaselineSheetDate (Get-Cell $validityRow.Cells 'C' $validityRow.RowNumber))
        return ((@($sheetColumns.PriceColumns)).Count -ge 1)
    } catch {
        return $false
    }
}

function Test-BaselineWorkbookFamily {
    param(
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings
    )

    $targetSheets = @(Get-BaselineWorkbookTargetSheetInfos -WorksheetInfos $WorksheetInfos)
    if ($targetSheets.Count -lt 7) {
        return $false
    }

    $normalizedNames = @($targetSheets | ForEach-Object { Normalize-Key $_.Name })
    foreach ($requiredName in @('ISRAEL', 'USA', 'CANADA', 'MEXICO', 'CARIBB')) {
        if (-not ($normalizedNames -contains $requiredName)) {
            return $false
        }
    }

    $reeferCount = @($normalizedNames | Where-Object { $_ -like '*REEFER' }).Count
    if ($reeferCount -lt 2) {
        return $false
    }

    $layoutMatches = 0
    foreach ($sheetInfo in ($targetSheets | Select-Object -First 3)) {
        $rows = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
        if (Test-BaselineWorkbookSheetLayout -Rows $rows) {
            $layoutMatches++
        }
    }

    return ($layoutMatches -ge 2)
}

function Get-BaselineSheetCountryHint {
    param([string]$SheetName)

    $normalized = Normalize-Key ($SheetName -replace '\s+REEFER$', '')
    switch ($normalized) {
        'ISRAEL' { return 'IL' }
        'USA' { return 'US' }
        'CANADA' { return 'CA' }
        'MEXICO' { return 'MX' }
        default { return '' }
    }
}

function Get-BaselineLocationAliasMap {
    return @{
        'GENOVA' = @('ITGOA')
        'LIVORNO' = @('ITLIV')
        'ASHDOD' = @('ILASH')
        'HAIFA' = @('ILHFA')
        'NEW YORK' = @('USNYC')
        'NORFOLK' = @('USORF')
        'SAVANNAH' = @('USSAV')
        'PORT EVERGLADES' = @('USPEF')
        'HALIFAX' = @('CAHAL')
        'TORONTO' = @('CATOR')
        'MONTREAL' = @('CAMTR')
        'VANCOUVER' = @('CAVAN')
        'VERACRUZ' = @('MXVER')
        'ALTAMIRA' = @('MXATM')
        'KINGSTON' = @('JMKIN')
        'PORT OF SPAIN' = @('TTPOS')
        'POINT LISAS' = @('TTPTS')
        'PARAMARIBO' = @('SRPBM')
        'SAN JUAN' = @('PRSJU')
        'SAN JUAN, PUERTO RICO' = @('PRSJU')
        'RIO HAINA' = @('DOHAI')
        'CAUCEDO' = @('DOCAU')
        'PUERTO CORTES' = @('HNPCR')
        'GEORGETOWN, GRAND CAYMAN' = @('KYGEC')
        'GEORGETOWN, GUYANA' = @('GYGEO')
        'BRIDGETOWN' = @('BBBGI')
        'MOIN' = @('CRPMN')
        'MANZANILLO' = @('PAMIT')
        'PANAMA MANZANILLO' = @('PAMIT')
        'LA GUAIRA' = @('VELAG')
        'PUERTO CABELLO' = @('VEPBL')
        'SANTO TOMAS DE CASTILLA' = @('GTSTC')
        'PUERTO SANTO TOMAS DE CASTILLA' = @('GTSTC')
        'SAN ANTONIO' = @('CLSAI')
        'CARTAGENA' = @('COSPC')
        'GUAYAQUIL' = @('ECGYE')
        'CALLAO' = @('PEPUE')
        'PUERTO CALLAO' = @('PEPUE')
        'HOUSTON' = @('USHOU')
        'CHICAGO' = @('USCHI')
    }
}

function Normalize-BaselineLocationText {
    param(
        [string]$Text,
        [switch]$IsDestination
    )

    $value = Normalize-Whitespace $Text
    if (-not $value) {
        return ''
    }

    if ($IsDestination) {
        if ((Normalize-Key $value) -eq 'GEORGETOWN GRAND CAYMAN') {
            return 'Georgetown, Grand Cayman'
        }
        if ((Normalize-Key $value) -eq 'GEORGETOWN GUYANNA') {
            return 'Georgetown, Guyana'
        }
        if ((Normalize-Key $value) -eq 'BRIDGETOWN - BARBADOS') {
            return 'Bridgetown'
        }
        if ((Normalize-Key $value) -eq 'SAN JUAN -PUERTO RICO') {
            return 'San Juan, Puerto Rico'
        }
        if ((Normalize-Key $value) -eq 'SAN JUAN - PUERTO RICO') {
            return 'San Juan, Puerto Rico'
        }
        if ((Normalize-Key $value) -eq 'PANAMA MANZANILLO') {
            return 'Panama Manzanillo'
        }
        if ((Normalize-Key $value) -eq 'SANTO TOMAS DE CASTILLA') {
            return 'Puerto Santo Tomas de Castilla'
        }

        $value = [regex]::Replace($value, '\((?i:\s*via.*?)\)', '')
        $value = [regex]::Replace($value, '(?i)\bvia\b.*$', '')
        $value = [regex]::Replace($value, '(?i)\bRAMP\b', '')
        $value = [regex]::Replace($value, '(?i)\bGUYANNA\b', 'Guyana')
    }

    return (Normalize-Whitespace $value)
}

function Split-BaselineLocationField {
    param(
        [string]$Text,
        [switch]$IsDestination
    )

    $source = Normalize-BaselineLocationText -Text $Text -IsDestination:$IsDestination
    if (-not $source) {
        return @()
    }

    $segments = @($source -split '\s*/\s*' | Where-Object { $_ })
    if ($segments.Count -eq 0) {
        return @()
    }

    $results = New-Object System.Collections.Generic.List[string]
    foreach ($segment in $segments) {
        Add-UniqueString -List $results -Value (Normalize-BaselineLocationText -Text $segment -IsDestination:$IsDestination)
    }

    return $results.ToArray()
}

function Resolve-BaselineLocationCodes {
    param(
        [string]$RawName,
        [hashtable]$Rules,
        [hashtable]$UnlocodeLookup,
        [string]$CountryHint = ''
    )

    $aliases = Get-BaselineLocationAliasMap
    $normalized = Normalize-Key $RawName
    if ($aliases.ContainsKey($normalized)) {
        return @($aliases[$normalized])
    }

    return @(Get-LocationCodes -RawName $RawName -Rules $Rules -UnlocodeLookup $UnlocodeLookup -CountryHint $CountryHint)
}

function Get-BaselineSheetRateRows {
    param(
        [object[]]$Rows,
        [int]$StartRowNumber,
        [int]$EndRowNumber,
        [object[]]$PriceColumns
    )

    $result = @()
    foreach ($row in ($Rows | Where-Object { $_.RowNumber -ge $StartRowNumber -and $_.RowNumber -lt $EndRowNumber } | Sort-Object RowNumber)) {
        $originText = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $destinationText = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        if (-not $originText -or -not $destinationText) {
            continue
        }

        $hasNumericRate = $false
        foreach ($priceColumn in $PriceColumns) {
            $rateText = Normalize-Whitespace (Get-Cell $row.Cells $priceColumn.Column $row.RowNumber)
            if ($rateText -match '^\d+(?:[.,]\d+)?$') {
                $hasNumericRate = $true
                break
            }
        }

        if ($hasNumericRate) {
            $result += $row
        }
    }

    return $result
}

function Get-BaselineSheetAdditionalRows {
    param(
        [object[]]$Rows,
        [int]$StartRowNumber
    )

    $result = @()
    $started = $false
    foreach ($row in ($Rows | Where-Object { $_.RowNumber -gt $StartRowNumber } | Sort-Object RowNumber)) {
        $bText = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $cText = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        $aText = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)

        if (-not $started) {
            if ($bText -or $cText) {
                $started = $true
            } else {
                continue
            }
        }

        $normalizedB = Normalize-Key $bText
        $normalizedA = Normalize-Key $aText
        if ((-not $bText -and -not $cText) -or
            $normalizedB -eq 'PORT' -or
            $normalizedB -like 'CURRENT WHARFAGE CHARGES*' -or
            $normalizedA -like '*FREE TIME*' -or
            $normalizedA -like '*THANK ADDITIONAL PLUG IN*') {
            break
        }

        $result += $row
    }

    return $result
}

function Convert-BaselineAdditionalRowsToGenericRows {
    param([object[]]$Rows)

    $converted = @()
    foreach ($row in $Rows) {
        $name = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $amount = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)
        $scope = Normalize-Whitespace (Get-Cell $row.Cells 'D' $row.RowNumber)
        if (-not $name -and -not $amount -and -not $scope) {
            continue
        }

        $cells = @{}
        $cells["A$($row.RowNumber)"] = $name
        $cells["B$($row.RowNumber)"] = $amount
        $cells["C$($row.RowNumber)"] = $scope
        $converted += [pscustomobject]@{
            RowNumber = $row.RowNumber
            Cells = $cells
        }
    }

    return $converted
}

function Get-BaselineAdditionalDetails {
    param(
        [object[]]$Rows,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$DestinationName
    )

    if (-not $Carrier -or -not $Direction) {
        return @()
    }

    $convertedRows = @(Convert-BaselineAdditionalRowsToGenericRows -Rows $Rows)
    if ($convertedRows.Count -eq 0) {
        return @()
    }

    try {
        $templates = @(Get-ExpectedAdditionalDetails -Rows $convertedRows -Carrier $Carrier -Direction $Direction -Rules $Rules)
    } catch {
        return @()
    }

    $details = @()
    foreach ($template in $templates) {
        if (Should-ApplyAdditionalToDestination -ApplyTargets $template.AppliesTo -DestinationName $DestinationName) {
            $details += $template.Detail
        }
    }

    return $details
}

function Convert-BaselineWorkbookFamily {
    param(
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [string]$Reference = '',
        [string]$ValidityStartDate = '',
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $sheetInfos = @(Get-BaselineWorkbookTargetSheetInfos -WorksheetInfos $WorksheetInfos)
    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { '' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { '' }

    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    $sheetContexts = @()
    foreach ($sheetInfo in $sheetInfos) {
        $rows = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
        $validity = Get-BaselineSheetValidityWindow -Rows $rows -ExplicitValidityStartDate $ValidityStartDate
        $sheetColumns = Get-BaselineSheetPriceColumns -Rows $rows
        $priceColumns = @($sheetColumns.PriceColumns)
        $addizionaliRow = Find-BaselineSheetAddizionaliRow -Rows $rows -AfterRowNumber $sheetColumns.EquipmentRow.RowNumber
        if (-not $addizionaliRow) {
            throw "Baseline workbook ADDIZIONALI block not found on sheet '$($sheetInfo.Name)'."
        }

        $rateRows = @(Get-BaselineSheetRateRows -Rows $rows -StartRowNumber ($sheetColumns.EquipmentRow.RowNumber + 1) -EndRowNumber $addizionaliRow.RowNumber -PriceColumns $priceColumns)
        $additionalRows = @(Get-BaselineSheetAdditionalRows -Rows $rows -StartRowNumber $addizionaliRow.RowNumber)
        $countryHint = Get-BaselineSheetCountryHint -SheetName $sheetInfo.Name
        $isReefer = ((Normalize-Key $sheetInfo.Name) -like '*REEFER') -or (@($priceColumns | Where-Object { $_.Evaluation -like "*Reefer" }).Count -gt 0)
        $currency = Get-BaselineSheetCurrency -HeaderRow $sheetColumns.HeaderRow

        foreach ($rateRow in $rateRows) {
            foreach ($originName in @(Split-BaselineLocationField -Text (Get-Cell $rateRow.Cells 'B' $rateRow.RowNumber))) {
                Add-UniqueString -List $rawLocationNames -Value $originName
            }
            foreach ($destinationName in @(Split-BaselineLocationField -Text (Get-Cell $rateRow.Cells 'C' $rateRow.RowNumber) -IsDestination)) {
                Add-UniqueString -List $rawLocationNames -Value $destinationName
            }
        }

        $sheetContexts += [pscustomobject]@{
            SheetInfo = $sheetInfo
            Rows = $rows
            Validity = $validity
            PriceColumns = $priceColumns
            RateRows = $rateRows
            AdditionalRows = $additionalRows
            CountryHint = $countryHint
            IsReefer = $isReefer
            Currency = $currency
        }
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($sheetContext in $sheetContexts) {
        foreach ($rateRow in $sheetContext.RateRows) {
            $originNames = @(Split-BaselineLocationField -Text (Get-Cell $rateRow.Cells 'B' $rateRow.RowNumber))
            $destinationNames = @(Split-BaselineLocationField -Text (Get-Cell $rateRow.Cells 'C' $rateRow.RowNumber) -IsDestination)
            foreach ($originName in $originNames) {
                $originCodes = @(Resolve-BaselineLocationCodes -RawName $originName -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint 'IT')
                foreach ($destinationName in $destinationNames) {
                    $destinationCodes = @(Resolve-BaselineLocationCodes -RawName $destinationName -Rules $Rules -UnlocodeLookup $unlocodeLookup -CountryHint $sheetContext.CountryHint)
                    $details = @()
                    foreach ($priceColumn in $sheetContext.PriceColumns) {
                        $rateValue = Convert-LocalizedNumberText (Get-Cell $rateRow.Cells $priceColumn.Column $rateRow.RowNumber)
                        if ($rateValue -match '^-?\d+(?:[.,]\d+)?$') {
                            $details += (New-PriceDetail 'Ocean Freight - Containers' $sheetContext.Currency '' $priceColumn.Evaluation $rateValue)
                        }
                    }

                    $details += @(Get-BaselineAdditionalDetails -Rows $sheetContext.AdditionalRows -Carrier $normalizedCarrier -Direction $normalizedDirection -Rules $Rules -DestinationName $destinationName)

                    foreach ($originCode in $originCodes) {
                        foreach ($destinationCode in $destinationCodes) {
                            $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $sheetContext.Validity.Start -ValidityEnd $sheetContext.Validity.End -Carrier $normalizedCarrier -PriceDetails $details -Reference $Reference
                            $rowIndex++
                        }
                    }
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Get-Baseline2WorkbookTariffSheetInfos {
    param([object[]]$WorksheetInfos)

    return @(
        $WorksheetInfos |
            Where-Object {
                $_.State -eq 'visible' -and
                $_.Name -ne 'ADDITIONAL & LOCAL CHARGES'
            } |
            Sort-Object Name
    )
}

function Get-Baseline2ExpectedPolGroups {
    return @(
        [pscustomobject]@{ Name = 'GENOVA'; Column = 'E'; PriceColumns = @('E', 'F', 'G') },
        [pscustomobject]@{ Name = 'LA SPEZIA'; Column = 'H'; PriceColumns = @('H', 'I', 'J') },
        [pscustomobject]@{ Name = 'VENEZIA'; Column = 'K'; PriceColumns = @('K', 'L', 'M') },
        [pscustomobject]@{ Name = 'ANCONA'; Column = 'N'; PriceColumns = @('N', 'O', 'P') },
        [pscustomobject]@{ Name = 'LIVORNO'; Column = 'Q'; PriceColumns = @('Q', 'R', 'S') },
        [pscustomobject]@{ Name = 'NAPOLI'; Column = 'T'; PriceColumns = @('T', 'U', 'V') },
        [pscustomobject]@{ Name = 'TRIESTE'; Column = 'W'; PriceColumns = @('W', 'X', 'Y') }
    )
}

function Get-Baseline2SheetValidityWindow {
    param([object[]]$Rows)

    $validityRow = $Rows | Where-Object {
        $_.RowNumber -eq 4 -and
        (Normalize-Key (Get-Cell $_.Cells 'M' $_.RowNumber)) -like "RATE'S VALIDITY*"
    } | Select-Object -First 1
    if (-not $validityRow) {
        throw 'Baseline2 workbook validity row not found.'
    }

    $startDate = Parse-BaselineSheetDate (Get-Cell $validityRow.Cells 'T' $validityRow.RowNumber)
    $endDate = Parse-BaselineSheetDate (Get-Cell $validityRow.Cells 'W' $validityRow.RowNumber)
    return [pscustomobject]@{
        Start = $startDate.ToString('dd/MM/yyyy')
        End = $endDate.ToString('dd/MM/yyyy')
    }
}

function Get-Baseline2SheetPolGroups {
    param([object[]]$Rows)

    $headerRow = $Rows | Where-Object { $_.RowNumber -eq 12 } | Select-Object -First 1
    $equipmentRow = $Rows | Where-Object { $_.RowNumber -eq 13 } | Select-Object -First 1
    if (-not $headerRow -or -not $equipmentRow) {
        throw 'Baseline2 workbook tariff header rows not found.'
    }

    if ((Normalize-Key (Get-Cell $headerRow.Cells 'A' 12)) -ne 'COUNTRY' -or
        (Normalize-Key (Get-Cell $headerRow.Cells 'B' 12)) -ne 'POD' -or
        (Normalize-Key (Get-Cell $headerRow.Cells 'C' 12)) -ne 'TERM' -or
        (Normalize-Key (Get-Cell $headerRow.Cells 'D' 12)) -ne 'SERVICE VIA' -or
        (Normalize-Key (Get-Cell $headerRow.Cells 'Z' 12)) -ne 'REMARKS') {
        throw 'Baseline2 workbook tariff header labels not recognized.'
    }

    $groups = @()
    foreach ($group in (Get-Baseline2ExpectedPolGroups)) {
        $headerName = Normalize-Whitespace (Get-Cell $headerRow.Cells $group.Column 12)
        if ((Normalize-Key $headerName) -ne $group.Name) {
            throw "Baseline2 workbook POL header '$($group.Name)' not found in column $($group.Column)."
        }

        $priceColumns = @()
        foreach ($column in $group.PriceColumns) {
            $token = Normalize-Whitespace (Get-Cell $equipmentRow.Cells $column 13)
            if (-not $token) {
                throw "Baseline2 workbook equipment header missing in column $column."
            }

            $tokenKey = Normalize-Key ($token -replace '''', '' -replace '\s+', '')
            $evaluation = switch ($tokenKey) {
                '20BOX' { "Cntr 20' Box"; break }
                '40BOX' { "Cntr 40' Box"; break }
                '40HC'  { "Cntrs 40' HC"; break }
                default { Get-BaselineContainerEvaluation $token }
            }

            $priceColumns += [pscustomobject]@{
                Column = $column
                Evaluation = $evaluation
                Token = $token
            }
        }

        $groups += [pscustomobject]@{
            OriginName = $group.Name
            PriceColumns = $priceColumns
        }
    }

    return @($groups)
}

function Test-Baseline2WorkbookSheetLayout {
    param([object[]]$Rows)

    try {
        [void](Get-Baseline2SheetValidityWindow -Rows $Rows)
        [void](Get-Baseline2SheetPolGroups -Rows $Rows)
        $titleRow = $Rows | Where-Object { $_.RowNumber -eq 7 } | Select-Object -First 1
        if (-not $titleRow) {
            return $false
        }

        $title = Normalize-Key (Get-Cell $titleRow.Cells 'A' 7)
        if ($title -notlike 'EXPORT F.A.K. RATES FROM ORIGIN ITALY TO *') {
            return $false
        }

        return ((Normalize-Key (Get-Cell ($Rows | Where-Object { $_.RowNumber -eq 13 } | Select-Object -First 1).Cells 'AD' 13)) -like '*ADDITIONAL*LOCAL*CHARGE*')
    } catch {
        return $false
    }
}

function Test-Baseline2WorkbookFamily {
    param(
        [object[]]$WorksheetInfos,
        [string[]]$SharedStrings
    )

    $visibleSheets = @($WorksheetInfos | Where-Object { $_.State -eq 'visible' })
    if ($visibleSheets.Count -ne 3) {
        return $false
    }

    $additionalSheet = $visibleSheets | Where-Object { $_.Name -eq 'ADDITIONAL & LOCAL CHARGES' } | Select-Object -First 1
    if (-not $additionalSheet) {
        return $false
    }

    $tariffSheets = @(Get-Baseline2WorkbookTariffSheetInfos -WorksheetInfos $WorksheetInfos)
    if ($tariffSheets.Count -ne 2) {
        return $false
    }

    foreach ($sheetInfo in $tariffSheets) {
        if ($sheetInfo.Name -notlike 'F.A.K.*') {
            return $false
        }

        $rows = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
        if (-not (Test-Baseline2WorkbookSheetLayout -Rows $rows)) {
            return $false
        }
    }

    return $true
}

function Get-Baseline2CountryHint {
    param([string]$CountryText)

    switch -Regex (Normalize-Key $CountryText) {
        '^AUSTRALIA$' { return 'AU' }
        '^BANGLADESH$' { return 'BD' }
        '^CAMBODIA$' { return 'KH' }
        '^CHINA' { return 'CN' }
        '^HONG KONG$' { return 'HK' }
        '^INDIA$' { return 'IN' }
        '^INDONESIA$' { return 'ID' }
        '^JAPAN$' { return 'JP' }
        '^KOREA' { return 'KR' }
        '^MALAYSIA$' { return 'MY' }
        '^MYANMAR$' { return 'MM' }
        '^NEW ZEALAND$' { return 'NZ' }
        '^PAKISTAN$' { return 'PK' }
        '^PHILIPPINES$' { return 'PH' }
        '^SINGAPORE$' { return 'SG' }
        '^SRI LANKA$' { return 'LK' }
        '^TAIWAN$' { return 'TW' }
        '^THAILAND$' { return 'TH' }
        '^VIETNAM$' { return 'VN' }
        default { return '' }
    }
}

function Resolve-Baseline2OriginCodes {
    param([string]$OriginName)

    switch (Normalize-Key $OriginName) {
        'GENOVA' { return @('ITGOA') }
        'LA SPEZIA' { return @('ITSPE') }
        'VENEZIA' { return @('ITVCE') }
        'ANCONA' { return @('ITAOI') }
        'LIVORNO' { return @('ITLIV') }
        'NAPOLI' { return @('ITNAP') }
        'TRIESTE' { return @('ITTRS') }
        default { throw "Baseline2 origin mapping not configured for '$OriginName'." }
    }
}

function Resolve-Baseline2DestinationCodes {
    param(
        [string]$DestinationName,
        [string]$CountryHint,
        [hashtable]$Rules,
        [hashtable]$UnlocodeLookup
    )

    $aliases = @{
        'ADELAIDE' = @('AUADL')
        'BANGKOK PAT, BMT, BBT' = @('THBKK')
        'BELAWAN' = @('IDBLW')
        'BRISBANE' = @('AUBNE')
        'CHITTAGONG' = @('BDCGP')
        'HO-CHI-MINH (CAI MEP)' = @('VNCMT')
        'HO-CHI-MINH (CAT LAI)' = @('VNSGN')
        'INCHON' = @('KRINC')
        'JAKARTA' = @('IDJKT')
        'KEELUNG' = @('TWKEL')
        'KUANTAN' = @('MYKUA')
        'LYTTELTON' = @('NZLYT')
        'NAGOYA' = @('JPNGO')
        'NAHA' = @('JPNAH')
        'NANHAI (SANSHAN)' = @('CNNAH')
        'NAPIER' = @('NZNPE')
        'NHAVA SHEVA' = @('INNSA')
        'PALEMBANG' = @('IDPLM')
        'PASIR GUDANG' = @('MYPGU')
        'PORT KELANG' = @('MYPKG')
        'RONGQI' = @('CNROQ')
        'SIHANOUKVILLE' = @('KHKOS')
        'SYDNEY' = @('AUSYD')
        'TAIZHOU' = @('CNTAZ')
        'XINHUI' = @('CNXIN')
        'ZHONGSHAN' = @('CNZSN')
    }

    $normalized = Normalize-Key $DestinationName
    if ($aliases.ContainsKey($normalized)) {
        return @($aliases[$normalized])
    }

    return @(Get-LocationCodes -RawName $DestinationName -Rules $Rules -UnlocodeLookup $UnlocodeLookup -CountryHint $CountryHint)
}

function Get-Baseline2SheetRateRows {
    param(
        [object[]]$Rows,
        [object[]]$PolGroups
    )

    $result = @()
    foreach ($row in ($Rows | Where-Object { $_.RowNumber -ge 14 } | Sort-Object RowNumber)) {
        $countryText = Normalize-Whitespace (Get-Cell $row.Cells 'A' $row.RowNumber)
        $destinationText = Normalize-Whitespace (Get-Cell $row.Cells 'B' $row.RowNumber)
        $termText = Normalize-Whitespace (Get-Cell $row.Cells 'C' $row.RowNumber)

        $hasTariffToken = $false
        foreach ($group in $PolGroups) {
            foreach ($priceColumn in $group.PriceColumns) {
                $rateText = Normalize-Whitespace (Get-Cell $row.Cells $priceColumn.Column $row.RowNumber)
                if ($rateText -match '^\d+(?:[.,]\d+)?$' -or (Normalize-Key $rateText) -eq 'NO SERVICE OPTION') {
                    $hasTariffToken = $true
                    break
                }
            }
            if ($hasTariffToken) {
                break
            }
        }

        if ($countryText -and $destinationText -and $termText -and $hasTariffToken) {
            $result += $row
            continue
        }

        if ($result.Count -gt 0) {
            break
        }
    }

    return @($result)
}

function Convert-Baseline2WorkbookFamily {
    param(
        [string[]]$SharedStrings,
        [object[]]$WorksheetInfos,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [string]$Reference = '',
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $sheetInfos = @(Get-Baseline2WorkbookTariffSheetInfos -WorksheetInfos $WorksheetInfos)
    $resolvedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { '' }
    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($group in (Get-Baseline2ExpectedPolGroups)) {
        Add-UniqueString -List $rawLocationNames -Value $group.Name
    }

    $sheetContexts = @()
    foreach ($sheetInfo in $sheetInfos) {
        $rows = Get-WorksheetRows -WorksheetPath $sheetInfo.Path -SharedStrings $SharedStrings
        $validity = Get-Baseline2SheetValidityWindow -Rows $rows
        $polGroups = @(Get-Baseline2SheetPolGroups -Rows $rows)
        $rateRows = @(Get-Baseline2SheetRateRows -Rows $rows -PolGroups $polGroups)
        foreach ($rateRow in $rateRows) {
            Add-UniqueString -List $rawLocationNames -Value (Normalize-Whitespace (Get-Cell $rateRow.Cells 'B' $rateRow.RowNumber))
        }

        $sheetContexts += [pscustomobject]@{
            SheetInfo = $sheetInfo
            Validity = $validity
            PolGroups = $polGroups
            RateRows = $rateRows
        }
    }

    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $headers = Get-OutputHeaders
    $outputRows = @()
    $rowIndex = 1

    foreach ($sheetContext in $sheetContexts) {
        foreach ($rateRow in $sheetContext.RateRows) {
            $countryText = Normalize-Whitespace (Get-Cell $rateRow.Cells 'A' $rateRow.RowNumber)
            $destinationName = Normalize-Whitespace (Get-Cell $rateRow.Cells 'B' $rateRow.RowNumber)
            $remarks = Normalize-Whitespace (Get-Cell $rateRow.Cells 'Z' $rateRow.RowNumber)
            $countryHint = Get-Baseline2CountryHint -CountryText $countryText
            $destinationCodes = @(Resolve-Baseline2DestinationCodes -DestinationName $destinationName -CountryHint $countryHint -Rules $Rules -UnlocodeLookup $unlocodeLookup)

            foreach ($group in $sheetContext.PolGroups) {
                $details = @()
                foreach ($priceColumn in $group.PriceColumns) {
                    $rateText = Normalize-Whitespace (Get-Cell $rateRow.Cells $priceColumn.Column $rateRow.RowNumber)
                    if (-not $rateText -or (Normalize-Key $rateText) -eq 'NO SERVICE OPTION') {
                        continue
                    }

                    $rateValue = Convert-LocalizedNumberText $rateText
                    if ($rateValue -match '^-?\d+(?:[.,]\d+)?$') {
                        $details += (New-PriceDetail 'Ocean Freight - Containers' 'USD' '' $priceColumn.Evaluation $rateValue)
                    }
                }

                if ($details.Count -eq 0) {
                    continue
                }

                $originCodes = @(Resolve-Baseline2OriginCodes -OriginName $group.OriginName)
                foreach ($originCode in $originCodes) {
                    foreach ($destinationCode in $destinationCodes) {
                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $sheetContext.Validity.Start -ValidityEnd $sheetContext.Validity.End -Carrier $resolvedCarrier -PriceDetails $details -Reference $Reference -Comment $remarks
                        $rowIndex++
                    }
                }
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Convert-ListinoToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [string]$Reference = '',
        [string]$ValidityStartDate = '',
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $packageRoot = New-TempDirectory
    try {
        Expand-XlsxPackage -Path $InputPath -DestinationPath $packageRoot
        $sharedStrings = Get-SharedStrings -PackageRoot $packageRoot
        $worksheetInfos = Get-WorksheetInfos -PackageRoot $packageRoot

        if (Test-CoscoFarEastWorkbook -WorksheetInfos $worksheetInfos) {
            Convert-CoscoFarEastWorkbook -PackageRoot $packageRoot -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        if (Test-CoscoStructuredWorkbook -WorksheetInfos $worksheetInfos) {
            Convert-CoscoStructuredWorkbook -PackageRoot $packageRoot -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        if (Test-EvergreenRvsWorkbook -InputPath $InputPath -WorksheetInfos $worksheetInfos -SharedStrings $sharedStrings) {
            Convert-EvergreenRvsWorkbook -InputPath $InputPath -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        if (Test-CoscoIetWorkbook -WorksheetInfos $worksheetInfos -SharedStrings $sharedStrings) {
            Convert-CoscoIetWorkbook -InputPath $InputPath -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        if (Test-HmmCitWorkbook -WorksheetInfos $worksheetInfos -SharedStrings $sharedStrings) {
            Convert-HmmCitWorkbook -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Rules $Rules
            return
        }

        if (Test-BaselineWorkbookFamily -WorksheetInfos $worksheetInfos -SharedStrings $sharedStrings) {
            Convert-BaselineWorkbookFamily -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Reference $Reference -ValidityStartDate $ValidityStartDate -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        if (Test-Baseline2WorkbookFamily -WorksheetInfos $worksheetInfos -SharedStrings $sharedStrings) {
            Convert-Baseline2WorkbookFamily -SharedStrings $sharedStrings -WorksheetInfos $worksheetInfos -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Reference $Reference -Rules $Rules -UnlocodePath $UnlocodePath
            return
        }

        $worksheetPath = Get-FirstVisibleWorksheetPath -PackageRoot $packageRoot
        $rows = Get-WorksheetRows -WorksheetPath $worksheetPath -SharedStrings $sharedStrings

        $headerRow = $rows | Where-Object { $_.RowNumber -eq 1 } | Select-Object -First 1
        $validity = Get-ValidityWindow (Get-Cell $headerRow.Cells 'A' 1)
        $destinations = Get-DestinationMap $rows
        $rateRows = Get-RateRows $rows
        $noteRows = @($rows | Where-Object { $_.RowNumber -gt $rateRows[-1].RowNumber })
        $isClassicPairMatrixLayout = Test-GenericClassicPairRateMatrixLayout -Rows $rows -HeaderRow $headerRow -Destinations $destinations -NoteRows $noteRows
        $formattedValidityStart = if ($isClassicPairMatrixLayout) { Format-OutputDateText $validity.Start } else { $validity.Start }
        $formattedValidityEnd = if ($isClassicPairMatrixLayout) { Format-OutputDateText $validity.End } else { $validity.End }
        $resolvedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { '' }
        $resolvedDirection = if ($Direction) { Normalize-Key $Direction } else { '' }
        if ($resolvedCarrier -and -not $resolvedDirection) {
            $resolvedDirection = (Resolve-CarrierDirection -Rows $noteRows -Rules $Rules -Carrier $resolvedCarrier -Direction '').Direction
        }

        $templates = @()
        if ($resolvedCarrier -and $resolvedDirection) {
            $templates = Get-ExpectedAdditionalDetails -Rows $noteRows -Carrier $resolvedCarrier -Direction $resolvedDirection -Rules $Rules
        }
        $rawLocationNames = New-Object System.Collections.Generic.List[string]
        foreach ($rateRow in $rateRows) {
            Add-UniqueString -List $rawLocationNames -Value (Get-Cell $rateRow.Cells 'A' $rateRow.RowNumber)
        }
        foreach ($destination in $destinations) {
            Add-UniqueString -List $rawLocationNames -Value $destination.Destination
        }
        $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
        $headers = Get-OutputHeaders

        $outputRows = @()
        $rowIndex = 1

        foreach ($rateRow in $rateRows) {
            $originCodes = Get-LocationCodes -RawName (Get-Cell $rateRow.Cells 'A' $rateRow.RowNumber) -Rules $Rules -UnlocodeLookup $unlocodeLookup
            foreach ($destination in $destinations) {
                $destinationCodes = Get-LocationCodes -RawName $destination.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup
                foreach ($originCode in $originCodes) {
                    foreach ($destinationCode in $destinationCodes) {
                        $details = @()

                        foreach ($priceColumn in $destination.PriceColumns) {
                            $rate = Normalize-Whitespace (Get-Cell $rateRow.Cells $priceColumn.Column $rateRow.RowNumber)
                            if ($rate) {
                                $oceanFreightLabel = if ($isClassicPairMatrixLayout) { 'Ocean Freight - Containers' } else { 'OCEAN FREIGHT - CONTAINERS' }
                                $details += (New-PriceDetail $oceanFreightLabel 'USD' '' $priceColumn.Evaluation $rate)
                            }
                        }

                        foreach ($template in $templates) {
                            if (Should-ApplyAdditionalToDestination -ApplyTargets $template.AppliesTo -DestinationName $destination.Destination) {
                                $details += $template.Detail
                            }
                        }

                        if ($isClassicPairMatrixLayout) {
                            $details = Add-Missing40HcFallbackDuplicates -Details $details -PriceColumns $destination.PriceColumns
                            $details = Clear-PriceDetailComments -Details $details
                        }

                        $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $formattedValidityStart -ValidityEnd $formattedValidityEnd -Carrier $resolvedCarrier -PriceDetails $details -Reference $Reference
                        $rowIndex++
                    }
                }
            }
        }

        Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
    }
    finally {
        Remove-DirectorySafe $packageRoot
    }
}

function Convert-HapagPdfToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = '',
        [string]$PdfText = ''
    )

    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { 'HAPAG-LLOYD' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { 'EXPORT' }

    if ($normalizedCarrier -ne 'HAPAG-LLOYD') {
        throw "PDF adapter currently supports HAPAG-LLOYD only. Received carrier '$Carrier'."
    }

    if ($normalizedDirection -ne 'EXPORT') {
        throw "PDF adapter currently supports Export quotations only. Received direction '$Direction'."
    }

    if (-not $PdfText) {
        $PdfText = Get-PdfText -InputPath $InputPath
    }

    $pdfText = $PdfText
    $pages = Get-HapagQuotePages -Text $pdfText
    if ((@($pages)).Count -eq 0) {
        throw 'No Hapag-Lloyd quotation pages found in the PDF.'
    }

    $headers = Get-OutputHeaders
    $reference = Get-HapagReference -Text $pdfText
    $pageRoutes = @()
    $rawLocationNames = New-Object System.Collections.Generic.List[string]
    foreach ($page in $pages) {
        $pageLines = @($page -split "`r?`n")
        $routeText = Get-HapagRouteText -Lines $pageLines
        $route = Parse-HapagRoute -RouteText $routeText
        Add-UniqueString -List $rawLocationNames -Value $route.Origin
        Add-UniqueString -List $rawLocationNames -Value $route.Destination
        $pageRoutes += [pscustomobject]@{
            Page = $page
            Route = $route
        }
    }
    $unlocodeLookup = Import-UnlocodeLookup -Path $UnlocodePath -RawNames $rawLocationNames.ToArray()
    $outputRows = @()
    $rowIndex = 1

    foreach ($pageRoute in $pageRoutes) {
        $page = $pageRoute.Page
        $route = $pageRoute.Route
        $validity = Get-HapagValidityWindow -PageText $page
        $transitTime = Get-HapagTransitTime -PageText $page
        $priceDetails = @()
        $priceDetails += Get-HapagOceanFreightDetails -PageText $page
        $priceDetails += Get-HapagAdditionalDetails -PageText $page -Carrier $normalizedCarrier -Direction $normalizedDirection -Rules $Rules

        foreach ($originCode in (Get-LocationCodes -RawName $route.Origin -Rules $Rules -UnlocodeLookup $unlocodeLookup)) {
            foreach ($destinationCode in (Get-LocationCodes -RawName $route.Destination -Rules $Rules -UnlocodeLookup $unlocodeLookup)) {
                $outputRows += Convert-RouteToRow -Index $rowIndex -FromAddress $originCode -ToAddress $destinationCode -ValidityStart $validity.Start -ValidityEnd $validity.End -Carrier $normalizedCarrier -PriceDetails $priceDetails -TransitTime $transitTime -Reference $reference
                $rowIndex++
            }
        }
    }

    Write-NormalizedWorkbook -OutputPath $OutputPath -Headers $headers -DataRows $outputRows
}

function Convert-PdfToNormalizedWorkbook {
    param(
        [string]$InputPath,
        [string]$OutputPath,
        [string]$Carrier,
        [string]$Direction,
        [hashtable]$Rules,
        [string]$UnlocodePath = ''
    )

    $pdfText = Get-PdfText -InputPath $InputPath
    $normalizedCarrier = if ($Carrier) { Normalize-Key $Carrier } else { '' }
    $normalizedDirection = if ($Direction) { Normalize-Key $Direction } else { '' }

    if ((-not $normalizedCarrier -or $normalizedCarrier -eq 'COSCO') -and (Test-CoscoCanadaPdfText -Text $pdfText)) {
        $resolvedDirection = if ($normalizedDirection) { $normalizedDirection } else { 'EXPORT' }
        Convert-CoscoCanadaPdfToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier 'COSCO' -Direction $resolvedDirection -Rules $Rules -UnlocodePath $UnlocodePath -PdfText $pdfText
        return
    }

    if ((-not $normalizedCarrier -or $normalizedCarrier -eq 'COSCO') -and (Test-CoscoSouthAmericaPdfText -Text $pdfText)) {
        $resolvedDirection = if ($normalizedDirection) { $normalizedDirection } else { 'EXPORT' }
        Convert-CoscoSouthAmericaPdfToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier 'COSCO' -Direction $resolvedDirection -Rules $Rules -UnlocodePath $UnlocodePath -PdfText $pdfText
        return
    }

    if ((-not $normalizedCarrier -or $normalizedCarrier -eq 'COSCO') -and (Test-CoscoIpakPdfText -Text $pdfText)) {
        $resolvedDirection = if ($normalizedDirection) { $normalizedDirection } else { 'EXPORT' }
        Convert-CoscoIpakPdfToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier 'COSCO' -Direction $resolvedDirection -Rules $Rules -UnlocodePath $UnlocodePath -PdfText $pdfText
        return
    }

    if ((-not $normalizedCarrier -or $normalizedCarrier -eq 'HAPAG-LLOYD') -and (@(Get-HapagQuotePages -Text $pdfText)).Count -gt 0) {
        $resolvedDirection = if ($normalizedDirection) { $normalizedDirection } else { 'EXPORT' }
        Convert-HapagPdfToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier 'HAPAG-LLOYD' -Direction $resolvedDirection -Rules $Rules -UnlocodePath $UnlocodePath -PdfText $pdfText
        return
    }

    throw 'No supported PDF adapter matched this file. Pass -Carrier explicitly or extend the PDF parser.'
}

if (-not $RulesPath) {
    $RulesPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'NormalizationRules.psd1'
}

if (-not (Test-Path -LiteralPath $RulesPath)) {
    throw "Rules file not found: $RulesPath"
}

$rules = Import-PowerShellDataFile -Path $RulesPath
$resolvedUnlocodePath = Resolve-UnlocodeLookupPath -ExplicitPath $UnlocodePath -InputPath $InputPath

if (-not $OutputPath) {
    $directory = Split-Path -Path $InputPath -Parent
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $OutputPath = Join-Path $directory ($fileName + '_normalized.xlsx')
}

$extension = [System.IO.Path]::GetExtension($InputPath).ToLowerInvariant()
switch ($extension) {
    '.xlsx' {
Convert-ListinoToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Reference $Reference -ValidityStartDate $ValidityStartDate -Rules $rules -UnlocodePath $resolvedUnlocodePath
        break
    }
    '.pdf' {
        Convert-PdfToNormalizedWorkbook -InputPath $InputPath -OutputPath $OutputPath -Carrier $Carrier -Direction $Direction -Rules $rules -UnlocodePath $resolvedUnlocodePath
        break
    }
    default {
        throw "Unsupported input format '$extension'. Supported formats: .xlsx, .pdf"
    }
}

Write-Output $OutputPath
