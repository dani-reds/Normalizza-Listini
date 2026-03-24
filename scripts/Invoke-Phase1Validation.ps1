param(
    [Parameter(Mandatory = $false)]
    [string]$ManifestPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RepoRoot {
    return (Split-Path -Parent $PSScriptRoot)
}

function Resolve-RepoPath {
    param(
        [string]$RepoRoot,
        [string]$PathText
    )

    if ([string]::IsNullOrWhiteSpace($PathText)) {
        return ''
    }

    if ([System.IO.Path]::IsPathRooted($PathText)) {
        return $PathText
    }

    return (Join-Path $RepoRoot $PathText)
}

function Get-SafeFileStem {
    param([string]$Text)

    $safe = [string]$Text
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, '_')
    }

    return ($safe -replace '\s+', '_')
}

function New-RunnerResult {
    param(
        [string]$SampleId,
        [string]$InputRelativePath,
        [string]$GeneratedOutputPath,
        [string]$ExpectedOutputRelativePath,
        [string]$Status,
        [int]$ErrorCount,
        [int]$DifferenceCount,
        [string[]]$FailureCategories,
        [object[]]$Errors = @(),
        [string]$GeneratedWorksheetName = '',
        [string]$ExpectedWorksheetName = '',
        [int]$GeneratedDataRowCount = 0,
        [int]$ExpectedWorkbookDataRowCount = 0,
        [int]$ManifestExpectedDataRowCount = 0,
        [string]$DiffReportPath = ''
    )

    return [pscustomobject]@{
        SampleId = $SampleId
        InputRelativePath = $InputRelativePath
        GeneratedOutputPath = $GeneratedOutputPath
        ExpectedOutputRelativePath = $ExpectedOutputRelativePath
        Status = $Status
        ErrorCount = $ErrorCount
        DifferenceCount = $DifferenceCount
        FailureCategories = @($FailureCategories)
        GeneratedWorksheetName = $GeneratedWorksheetName
        ExpectedWorksheetName = $ExpectedWorksheetName
        GeneratedDataRowCount = $GeneratedDataRowCount
        ExpectedWorkbookDataRowCount = $ExpectedWorkbookDataRowCount
        ManifestExpectedDataRowCount = $ManifestExpectedDataRowCount
        DiffReportPath = $DiffReportPath
        Errors = @($Errors)
    }
}

function New-ManifestFailureResult {
    param(
        [object]$Sample,
        [string]$GeneratedOutputRelativePath,
        [string]$Message
    )

    return (New-RunnerResult `
        -SampleId ([string]$Sample.Id) `
        -InputRelativePath ([string]$Sample.InputRelativePath) `
        -GeneratedOutputPath $GeneratedOutputRelativePath `
        -ExpectedOutputRelativePath ([string]$Sample.ExpectedOutputRelativePath) `
        -Status 'FAIL' `
        -ErrorCount 1 `
        -DifferenceCount 0 `
        -FailureCategories @('Manifest') `
        -ManifestExpectedDataRowCount ([int]$Sample.ExpectedDataRowCount) `
        -Errors @($Message))
}

function New-GenerationFailureResult {
    param(
        [object]$Sample,
        [string]$GeneratedOutputRelativePath,
        [string]$Message
    )

    return (New-RunnerResult `
        -SampleId ([string]$Sample.Id) `
        -InputRelativePath ([string]$Sample.InputRelativePath) `
        -GeneratedOutputPath $GeneratedOutputRelativePath `
        -ExpectedOutputRelativePath ([string]$Sample.ExpectedOutputRelativePath) `
        -Status 'FAIL' `
        -ErrorCount 1 `
        -DifferenceCount 0 `
        -FailureCategories @('Generation') `
        -ManifestExpectedDataRowCount ([int]$Sample.ExpectedDataRowCount) `
        -Errors @($Message))
}

function Convert-ValidationResultToRunnerResult {
    param(
        [object]$Sample,
        [string]$GeneratedOutputRelativePath,
        [object]$ValidationResult
    )

    return (New-RunnerResult `
        -SampleId ([string]$Sample.Id) `
        -InputRelativePath ([string]$Sample.InputRelativePath) `
        -GeneratedOutputPath $GeneratedOutputRelativePath `
        -ExpectedOutputRelativePath ([string]$Sample.ExpectedOutputRelativePath) `
        -Status ([string]$ValidationResult.Status) `
        -ErrorCount ([int]$ValidationResult.ErrorCount) `
        -DifferenceCount ([int]$ValidationResult.DifferenceCount) `
        -FailureCategories @($ValidationResult.FailureCategories) `
        -GeneratedWorksheetName ([string]$ValidationResult.GeneratedWorksheetName) `
        -ExpectedWorksheetName ([string]$ValidationResult.ExpectedWorksheetName) `
        -GeneratedDataRowCount ([int]$ValidationResult.GeneratedDataRowCount) `
        -ExpectedWorkbookDataRowCount ([int]$ValidationResult.ExpectedDataRowCount) `
        -ManifestExpectedDataRowCount ([int]$Sample.ExpectedDataRowCount) `
        -DiffReportPath ([string]$ValidationResult.DiffReportPath) `
        -Errors @($ValidationResult.Errors))
}

function Write-HumanReadableSummary {
    param(
        [string]$Path,
        [datetime]$StartedAt,
        [datetime]$FinishedAt,
        [string]$ManifestPath,
        [object[]]$Results
    )

    $passCount = @($Results | Where-Object { $_.Status -eq 'PASS' }).Count
    $failCount = @($Results | Where-Object { $_.Status -ne 'PASS' }).Count

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Phase 1 validation started : $($StartedAt.ToString('s'))")
    $lines.Add("Phase 1 validation finished: $($FinishedAt.ToString('s'))")
    $lines.Add("Manifest: $ManifestPath")
    $lines.Add("Samples : $(@($Results).Count)")
    $lines.Add("PASS    : $passCount")
    $lines.Add("FAIL    : $failCount")
    $lines.Add('')

    foreach ($result in $Results) {
        $lines.Add("SampleId: $($result.SampleId)")
        $lines.Add("InputRelativePath: $($result.InputRelativePath)")
        $lines.Add("GeneratedOutputPath: $($result.GeneratedOutputPath)")
        $lines.Add("ExpectedOutputRelativePath: $($result.ExpectedOutputRelativePath)")
        $lines.Add("Status: $($result.Status)")
        $lines.Add("ErrorCount: $($result.ErrorCount)")
        $lines.Add("DifferenceCount: $($result.DifferenceCount)")
        $lines.Add("FailureCategories: $(if (@($result.FailureCategories).Count -gt 0) { (@($result.FailureCategories) -join ', ') } else { 'none' })")
        if (@($result.Errors).Count -gt 0) {
            $lines.Add("Errors: $((@($result.Errors)) -join ' | ')")
        }
        if ($result.DiffReportPath) {
            $lines.Add("DiffReportPath: $($result.DiffReportPath)")
        }
        $lines.Add('')
    }

    [System.IO.File]::WriteAllLines($Path, $lines)
}

$repoRoot = Get-RepoRoot
$manifestFilePath = if ($ManifestPath) { Resolve-RepoPath -RepoRoot $repoRoot -PathText $ManifestPath } else { Join-Path $repoRoot 'samples\SampleManifest.psd1' }
$normalizeScriptPath = Join-Path $repoRoot 'Normalize-Listino.ps1'
$validatorScriptPath = Join-Path $PSScriptRoot 'Test-NormalizedWorkbook.ps1'
$generatedOutputDirectory = Join-Path $repoRoot 'output\validation'
$logDirectory = Join-Path $repoRoot 'logs\validation'

[System.IO.Directory]::CreateDirectory($generatedOutputDirectory) | Out-Null
[System.IO.Directory]::CreateDirectory($logDirectory) | Out-Null

$startedAt = Get-Date
$timestamp = $startedAt.ToString('yyyyMMdd-HHmmss-fff')
$jsonLogPath = Join-Path $logDirectory "phase1-validation-run-$timestamp.json"
$summaryLogPath = Join-Path $logDirectory "phase1-validation-summary-$timestamp.txt"

$manifest = Import-PowerShellDataFile -Path $manifestFilePath
$samples = @($manifest.Samples)
$results = New-Object System.Collections.Generic.List[object]

foreach ($sample in $samples) {
    $generatedFileName = ('{0}_generated.xlsx' -f (Get-SafeFileStem ([string]$sample.Id)))
    $generatedOutputRelativePath = ('output/validation/{0}' -f $generatedFileName)
    $generatedOutputAbsolutePath = Join-Path $generatedOutputDirectory $generatedFileName

    $missingFields = New-Object System.Collections.Generic.List[string]
    foreach ($fieldName in @('Id', 'InputRelativePath', 'ExpectedOutputRelativePath', 'SourceType')) {
        if (-not $sample.ContainsKey($fieldName) -or [string]::IsNullOrWhiteSpace([string]$sample[$fieldName])) {
            $missingFields.Add($fieldName)
        }
    }

    if ($missingFields.Count -gt 0) {
        $results.Add((New-ManifestFailureResult -Sample $sample -GeneratedOutputRelativePath $generatedOutputRelativePath -Message ("Missing required manifest field(s): " + ($missingFields -join ', '))))
        continue
    }

    $inputPath = Resolve-RepoPath -RepoRoot $repoRoot -PathText ([string]$sample.InputRelativePath)
    $expectedPath = Resolve-RepoPath -RepoRoot $repoRoot -PathText ([string]$sample.ExpectedOutputRelativePath)

    if (-not (Test-Path -LiteralPath $inputPath)) {
        $results.Add((New-ManifestFailureResult -Sample $sample -GeneratedOutputRelativePath $generatedOutputRelativePath -Message "Input file not found: $inputPath"))
        continue
    }

    if (-not (Test-Path -LiteralPath $expectedPath)) {
        $results.Add((New-ManifestFailureResult -Sample $sample -GeneratedOutputRelativePath $generatedOutputRelativePath -Message "Expected output file not found: $expectedPath"))
        continue
    }

    $normalizeParams = @{
        InputPath = $inputPath
        OutputPath = $generatedOutputAbsolutePath
    }

    if ($sample.ContainsKey('InvokeParameters') -and $sample.InvokeParameters) {
        foreach ($key in $sample.InvokeParameters.Keys) {
            $normalizeParams[$key] = $sample.InvokeParameters[$key]
        }
    }

    try {
        & $normalizeScriptPath @normalizeParams | Out-Null
        if (-not (Test-Path -LiteralPath $generatedOutputAbsolutePath)) {
            throw "Generated workbook not found after normalization: $generatedOutputAbsolutePath"
        }
    }
    catch {
        $results.Add((New-GenerationFailureResult -Sample $sample -GeneratedOutputRelativePath $generatedOutputRelativePath -Message $_.Exception.Message))
        continue
    }

    try {
        $validationResult = & $validatorScriptPath -GeneratedWorkbookPath $generatedOutputAbsolutePath -ExpectedWorkbookPath $expectedPath
        $results.Add((Convert-ValidationResultToRunnerResult -Sample $sample -GeneratedOutputRelativePath $generatedOutputRelativePath -ValidationResult $validationResult))
    }
    catch {
        $results.Add((New-RunnerResult `
            -SampleId ([string]$sample.Id) `
            -InputRelativePath ([string]$sample.InputRelativePath) `
            -GeneratedOutputPath $generatedOutputRelativePath `
            -ExpectedOutputRelativePath ([string]$sample.ExpectedOutputRelativePath) `
            -Status 'FAIL' `
            -ErrorCount 1 `
            -DifferenceCount 0 `
            -FailureCategories @('ValidationExecution') `
            -ManifestExpectedDataRowCount ([int]$sample.ExpectedDataRowCount) `
            -Errors @($_.Exception.Message)))
    }
}

$finishedAt = Get-Date
$resultArray = $results.ToArray()
$passCount = @($resultArray | Where-Object { $_.Status -eq 'PASS' }).Count
$failCount = @($resultArray | Where-Object { $_.Status -ne 'PASS' }).Count

$runReport = [pscustomobject]@{
    StartedAt = $startedAt.ToString('s')
    FinishedAt = $finishedAt.ToString('s')
    ManifestPath = $manifestFilePath
    SampleCount = @($resultArray).Count
    PassCount = $passCount
    FailCount = $failCount
    Results = @($resultArray)
}

$runReport | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonLogPath -Encoding UTF8
Write-HumanReadableSummary -Path $summaryLogPath -StartedAt $startedAt -FinishedAt $finishedAt -ManifestPath $manifestFilePath -Results @($resultArray)

foreach ($result in $resultArray) {
    $categoryText = if (@($result.FailureCategories).Count -gt 0) { (@($result.FailureCategories) -join ',') } else { 'none' }
    Write-Output ("[{0}] {1} errors={2} diff={3} categories={4}" -f $result.Status, $result.SampleId, $result.ErrorCount, $result.DifferenceCount, $categoryText)
}
Write-Output ("Phase 1 validation completed: PASS={0} FAIL={1}" -f $passCount, $failCount)
Write-Output ("Machine-readable log: {0}" -f $jsonLogPath)
Write-Output ("Summary log: {0}" -f $summaryLogPath)

if ($failCount -gt 0) {
    exit 1
}

exit 0
