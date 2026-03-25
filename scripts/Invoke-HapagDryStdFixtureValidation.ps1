param(
    [Parameter(Mandatory = $true)]
    [string]$InputPdfPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = '',

    [Parameter(Mandatory = $false)]
    [int[]]$PageNumbers = @(1, 2, 3, 4, 5, 6, 7, 8)
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RepoRoot {
    return (Split-Path -Parent $PSScriptRoot)
}

function Join-Pages {
    param([string[]]$Pages)

    return ($Pages -join "`f")
}

$repoRoot = Get-RepoRoot
$normalizeScriptPath = Join-Path $repoRoot 'Normalize-Listino.ps1'
$rulesPath = Join-Path $repoRoot 'NormalizationRules.psd1'
$fixtureOutputDirectory = Join-Path $repoRoot 'output\fixture-validation'

[System.IO.Directory]::CreateDirectory($fixtureOutputDirectory) | Out-Null

if (-not $OutputPath) {
    $OutputPath = Join-Path $fixtureOutputDirectory 'hapag-dry-std-port-to-port-candidate.xlsx'
}

$normalizeScriptText = [System.IO.File]::ReadAllText($normalizeScriptPath)
$bodyStart = $normalizeScriptText.IndexOf('Set-StrictMode -Version Latest')
$bodyEnd = $normalizeScriptText.IndexOf("if (-not `$RulesPath)")

if ($bodyStart -lt 0 -or $bodyEnd -le $bodyStart) {
    throw 'Unable to isolate reusable function block from Normalize-Listino.ps1.'
}

$reusableBlock = $normalizeScriptText.Substring($bodyStart, $bodyEnd - $bodyStart)
$reusableBlock = $reusableBlock.Replace('$script:CurrentScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path', ('$script:CurrentScriptDirectory = ' + "'" + $repoRoot.Replace("'", "''") + "'"))

. ([scriptblock]::Create($reusableBlock))
$script:CurrentScriptDirectory = $repoRoot

$layoutText = Get-PdfText -InputPath $InputPdfPath
$rawText = Get-PdfText -InputPath $InputPdfPath -Mode 'raw'
$layoutPages = @($layoutText -split "`f")
$rawPages = @($rawText -split "`f")

$selectedLayoutPages = New-Object System.Collections.Generic.List[string]
$selectedRawPages = New-Object System.Collections.Generic.List[string]

foreach ($pageNumber in $PageNumbers) {
    if ($pageNumber -le 0) {
        throw "Invalid page number '$pageNumber'."
    }

    $pageIndex = $pageNumber - 1
    if ($pageIndex -ge $rawPages.Count -or $pageIndex -ge $layoutPages.Count) {
        throw "Requested page $pageNumber is outside extracted PDF page range."
    }

    $rawPage = [string]$rawPages[$pageIndex]
    $layoutPage = [string]$layoutPages[$pageIndex]

    if (-not (Test-HapagDryStdTariffPage -PageText $rawPage)) {
        throw "Requested page $pageNumber is not a supported Hapag dry/std tariff page."
    }

    $unsupportedReason = Get-HapagDryStdUnsupportedPageReason -PageText $rawPage
    if ($unsupportedReason) {
        throw "Requested page $pageNumber is outside v1 scope: $unsupportedReason."
    }

    $selectedLayoutPages.Add($layoutPage)
    $selectedRawPages.Add($rawPage)
}

$selectedLayoutText = Join-Pages -Pages $selectedLayoutPages.ToArray()
$selectedRawText = Join-Pages -Pages $selectedRawPages.ToArray()

$layoutFixturePath = Join-Path $fixtureOutputDirectory 'hapag-dry-std-port-to-port.layout.txt'
$rawFixturePath = Join-Path $fixtureOutputDirectory 'hapag-dry-std-port-to-port.raw.txt'
[System.IO.File]::WriteAllText($layoutFixturePath, $selectedLayoutText, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($rawFixturePath, $selectedRawText, [System.Text.Encoding]::UTF8)

function Get-PdfText {
    param(
        [string]$InputPath,
        [ValidateSet('layout', 'table', 'raw', 'lineprinter')]
        [string]$Mode = 'layout'
    )

    switch ($Mode) {
        'raw' { return $selectedRawText }
        default { return $selectedLayoutText }
    }
}

$rules = Import-PowerShellDataFile -Path $rulesPath
$resolvedUnlocodePath = Resolve-UnlocodeLookupPath -ExplicitPath '' -InputPath $InputPdfPath

Convert-HapagDryStdPdfToNormalizedWorkbook `
    -InputPath $InputPdfPath `
    -OutputPath $OutputPath `
    -Carrier 'HAPAG-LLOYD' `
    -Direction 'Export' `
    -Rules $rules `
    -UnlocodePath $resolvedUnlocodePath `
    -PdfText $selectedLayoutText

[pscustomobject]@{
    OutputPath = $OutputPath
    LayoutFixturePath = $layoutFixturePath
    RawFixturePath = $rawFixturePath
    PageNumbers = @($PageNumbers)
}
