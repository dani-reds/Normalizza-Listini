param(
    [Parameter(Mandatory = $true)]
    [string]$GeneratedWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedWorkbookPath,

    [Parameter(Mandatory = $false)]
    [string[]]$AllowedTransshipmentValues = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-ExpectedOutputHeaders {
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

function Get-FirstWorksheetInfo {
    param([string]$PackageRoot)

    $workbookPath = Join-Path $PackageRoot 'xl\workbook.xml'
    $relsPath = Join-Path $PackageRoot 'xl\_rels\workbook.xml.rels'

    if (-not (Test-Path -LiteralPath $workbookPath)) {
        throw "Workbook file not found: $workbookPath"
    }

    if (-not (Test-Path -LiteralPath $relsPath)) {
        throw "Workbook relationships file not found: $relsPath"
    }

    $workbookDoc = Load-XmlDocument $workbookPath
    $relsDoc = Load-XmlDocument $relsPath

    $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookDoc.NameTable)
    $workbookNs.AddNamespace('main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

    $relsNs = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
    $relsNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')

    $firstSheet = $workbookDoc.SelectSingleNode('/main:workbook/main:sheets/main:sheet[1]', $workbookNs)
    if ($null -eq $firstSheet) {
        throw 'No worksheet found in workbook.'
    }

    $relationshipId = $firstSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    if (-not $relationshipId) {
        throw 'First worksheet is missing relationship id.'
    }

    $relationship = $relsDoc.SelectSingleNode("/rel:Relationships/rel:Relationship[@Id='$relationshipId']", $relsNs)
    if ($null -eq $relationship) {
        throw "Relationship '$relationshipId' not found for first worksheet."
    }

    return [pscustomobject]@{
        Name = [string]$firstSheet.GetAttribute('name')
        Path = Resolve-OpenXmlTargetPath -PackageRoot $PackageRoot -Target ([string]$relationship.GetAttribute('Target'))
    }
}

function Convert-ColumnNameToIndex {
    param([string]$ColumnName)

    $name = ([string]$ColumnName).ToUpperInvariant()
    if ($name -notmatch '^[A-Z]+$') {
        throw "Invalid column name '$ColumnName'."
    }

    $index = 0
    foreach ($char in $name.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return $index
}

function Normalize-CellText {
    param([object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value) -replace "`r`n?", "`n"
}

function Get-SharedStrings {
    param([string]$PackageRoot)

    $sharedStringsPath = Join-Path $PackageRoot 'xl\sharedStrings.xml'
    if (-not (Test-Path -LiteralPath $sharedStringsPath)) {
        return @()
    }

    $doc = Load-XmlDocument $sharedStringsPath
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $values = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($doc.SelectNodes('/main:sst/main:si', $ns))) {
        $textNodes = @($item.SelectNodes('.//main:t', $ns))
        if ($textNodes.Count -eq 0) {
            $values.Add('')
            continue
        }

        $values.Add((($textNodes | ForEach-Object { $_.InnerText }) -join ''))
    }

    return $values.ToArray()
}

function Get-CellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [string[]]$SharedStrings,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $cellType = [string]$Cell.GetAttribute('t')
    if ($cellType -eq 'inlineStr') {
        $textNodes = @($Cell.SelectNodes('main:is//main:t', $NamespaceManager))
        if ($textNodes.Count -eq 0) {
            return ''
        }

        return (($textNodes | ForEach-Object { $_.InnerText }) -join '')
    }

    $valueNode = $Cell.SelectSingleNode('main:v', $NamespaceManager)
    $value = if ($null -ne $valueNode) { $valueNode.InnerText } else { '' }

    if ($cellType -eq 's') {
        if ([string]::IsNullOrWhiteSpace($value)) {
            return ''
        }

        $index = [int]$value
        if ($index -lt 0 -or $index -ge $SharedStrings.Count) {
            throw "Shared string index '$index' is out of range."
        }

        return $SharedStrings[$index]
    }

    return $value
}

function Convert-WorksheetRowToArray {
    param(
        [System.Xml.XmlElement]$Row,
        [int]$ColumnCount,
        [string[]]$SharedStrings,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    $values = New-Object string[] $ColumnCount
    for ($i = 0; $i -lt $ColumnCount; $i++) {
        $values[$i] = ''
    }

    $maxColumnIndex = 0
    $nextColumnIndex = 1

    foreach ($cell in @($Row.SelectNodes('main:c', $NamespaceManager))) {
        $columnIndex = $nextColumnIndex
        $reference = [string]$cell.GetAttribute('r')
        if ($reference -match '^([A-Z]+)\d+$') {
            $columnIndex = Convert-ColumnNameToIndex $Matches[1]
        }

        if ($columnIndex -gt $maxColumnIndex) {
            $maxColumnIndex = $columnIndex
        }

        if ($columnIndex -ge 1 -and $columnIndex -le $ColumnCount) {
            $values[$columnIndex - 1] = Normalize-CellText (Get-CellText -Cell $cell -SharedStrings $SharedStrings -NamespaceManager $NamespaceManager)
        }

        $nextColumnIndex = $columnIndex + 1
    }

    return [pscustomobject]@{
        Values = $values
        MaxColumnIndex = $maxColumnIndex
        ExcelRowNumber = [int]$Row.GetAttribute('r')
    }
}

function Get-WorkbookTable {
    param(
        [string]$WorkbookPath,
        [int]$ExpectedColumnCount
    )

    $packageRoot = New-TempDirectory
    try {
        Expand-XlsxPackage -Path $WorkbookPath -DestinationPath $packageRoot

        $worksheetInfo = Get-FirstWorksheetInfo -PackageRoot $packageRoot
        if (-not (Test-Path -LiteralPath $worksheetInfo.Path)) {
            throw "First worksheet path not found: $($worksheetInfo.Path)"
        }

        $sharedStrings = Get-SharedStrings -PackageRoot $packageRoot
        $sheetDoc = Load-XmlDocument $worksheetInfo.Path
        $ns = New-Object System.Xml.XmlNamespaceManager($sheetDoc.NameTable)
        $ns.AddNamespace('main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

        $rowNodes = @($sheetDoc.SelectNodes('/main:worksheet/main:sheetData/main:row', $ns))
        if ($rowNodes.Count -eq 0) {
            throw 'Worksheet has no rows.'
        }

        $headerRow = Convert-WorksheetRowToArray -Row $rowNodes[0] -ColumnCount $ExpectedColumnCount -SharedStrings $sharedStrings -NamespaceManager $ns

        $dataRows = New-Object System.Collections.Generic.List[object]
        $maxColumnIndexSeen = $headerRow.MaxColumnIndex
        $rowsWithExtraColumns = New-Object System.Collections.Generic.List[int]

        if ($headerRow.MaxColumnIndex -gt $ExpectedColumnCount) {
            $rowsWithExtraColumns.Add(1)
        }

        for ($i = 1; $i -lt $rowNodes.Count; $i++) {
            $rowData = Convert-WorksheetRowToArray -Row $rowNodes[$i] -ColumnCount $ExpectedColumnCount -SharedStrings $sharedStrings -NamespaceManager $ns
            if ($rowData.MaxColumnIndex -gt $ExpectedColumnCount) {
                $excelRowNumber = if ($rowData.ExcelRowNumber -gt 0) { $rowData.ExcelRowNumber } else { $i + 1 }
                $rowsWithExtraColumns.Add($excelRowNumber)
            }

            if ($rowData.MaxColumnIndex -gt $maxColumnIndexSeen) {
                $maxColumnIndexSeen = $rowData.MaxColumnIndex
            }

            $dataRows.Add($rowData.Values)
        }

        return [pscustomobject]@{
            WorksheetName = $worksheetInfo.Name
            HeaderValues = $headerRow.Values
            HeaderActualColumnCount = $headerRow.MaxColumnIndex
            DataRows = $dataRows.ToArray()
            DataRowCount = $dataRows.Count
            MaxColumnIndexSeen = $maxColumnIndexSeen
            RowsWithExtraColumns = $rowsWithExtraColumns.ToArray()
        }
    }
    finally {
        Remove-DirectorySafe -Path $packageRoot
    }
}

function Test-AddressColumns {
    param(
        [object]$Table,
        [string]$WorkbookRole,
        [string[]]$AllowedTransshipmentValues
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $allowed = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($value in @($AllowedTransshipmentValues)) {
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            [void]$allowed.Add((Normalize-CellText $value))
        }
    }

    $fromIndex = 3
    $toIndex = 4
    $transshipmentIndex = 9
    $pattern = '^[A-Z]{5}$'

    for ($rowIndex = 0; $rowIndex -lt $Table.DataRows.Count; $rowIndex++) {
        $excelRowNumber = $rowIndex + 2
        $row = $Table.DataRows[$rowIndex]

        $fromValue = Normalize-CellText $row[$fromIndex]
        if ([string]::IsNullOrEmpty($fromValue) -or $fromValue -notmatch $pattern) {
            $errors.Add("$WorkbookRole row $excelRowNumber has invalid From Address '$fromValue'.")
        }

        $toValue = Normalize-CellText $row[$toIndex]
        if ([string]::IsNullOrEmpty($toValue) -or $toValue -notmatch $pattern) {
            $errors.Add("$WorkbookRole row $excelRowNumber has invalid To Address '$toValue'.")
        }

        $transshipmentValue = Normalize-CellText $row[$transshipmentIndex]
        if (-not [string]::IsNullOrEmpty($transshipmentValue)) {
            if (($transshipmentValue -notmatch $pattern) -and (-not $allowed.Contains($transshipmentValue))) {
                $errors.Add("$WorkbookRole row $excelRowNumber has invalid Transshipment Address '$transshipmentValue'.")
            }
        }
    }

    return $errors.ToArray()
}

function New-DifferencePreview {
    param(
        [string]$Section,
        [int]$ExcelRow,
        [int]$ColumnNumber,
        [string]$ColumnName,
        [string]$GeneratedValue,
        [string]$ExpectedValue
    )

    return [pscustomobject]@{
        Section = $Section
        ExcelRow = $ExcelRow
        ColumnNumber = $ColumnNumber
        ColumnName = $ColumnName
        GeneratedValue = $GeneratedValue
        ExpectedValue = $ExpectedValue
    }
}

function Compare-WorkbookTables {
    param(
        [object]$GeneratedTable,
        [object]$ExpectedTable,
        [string[]]$SchemaHeaders
    )

    $headerErrors = New-Object System.Collections.Generic.List[string]
    $structuralErrors = New-Object System.Collections.Generic.List[string]
    $differencePreview = New-Object System.Collections.Generic.List[object]
    $storedDifferenceLimit = 200
    $differenceCount = 0
    $schemaDifferenceCount = 0
    $contentDifferenceCount = 0

    if ($GeneratedTable.HeaderActualColumnCount -ne $SchemaHeaders.Count) {
        $headerErrors.Add("Generated workbook header count is $($GeneratedTable.HeaderActualColumnCount); expected $($SchemaHeaders.Count).")
    }

    if ($ExpectedTable.HeaderActualColumnCount -ne $SchemaHeaders.Count) {
        $headerErrors.Add("Expected workbook header count is $($ExpectedTable.HeaderActualColumnCount); expected $($SchemaHeaders.Count).")
    }

    if ($GeneratedTable.RowsWithExtraColumns.Count -gt 0) {
        $structuralErrors.Add("Generated workbook contains cells beyond column $($SchemaHeaders.Count) on rows: $($GeneratedTable.RowsWithExtraColumns -join ', ').")
    }

    if ($ExpectedTable.RowsWithExtraColumns.Count -gt 0) {
        $structuralErrors.Add("Expected workbook contains cells beyond column $($SchemaHeaders.Count) on rows: $($ExpectedTable.RowsWithExtraColumns -join ', ').")
    }

    for ($columnIndex = 0; $columnIndex -lt $SchemaHeaders.Count; $columnIndex++) {
        $generatedHeader = $GeneratedTable.HeaderValues[$columnIndex]
        $expectedHeader = $ExpectedTable.HeaderValues[$columnIndex]
        $schemaHeader = $SchemaHeaders[$columnIndex]

        if ($generatedHeader -ne $expectedHeader) {
            $differenceCount++
            $schemaDifferenceCount++
            if ($differencePreview.Count -lt $storedDifferenceLimit) {
                $differencePreview.Add((New-DifferencePreview -Section 'Schema' -ExcelRow 1 -ColumnNumber ($columnIndex + 1) -ColumnName $schemaHeader -GeneratedValue $generatedHeader -ExpectedValue $expectedHeader))
            }
        }

        if ($generatedHeader -ne $schemaHeader) {
            $headerErrors.Add("Generated workbook header mismatch at column $($columnIndex + 1): found '$generatedHeader', expected schema '$schemaHeader'.")
        }

        if ($expectedHeader -ne $schemaHeader) {
            $headerErrors.Add("Expected workbook header mismatch at column $($columnIndex + 1): found '$expectedHeader', expected schema '$schemaHeader'.")
        }
    }

    if ($GeneratedTable.DataRowCount -ne $ExpectedTable.DataRowCount) {
        $structuralErrors.Add("Data row count mismatch: generated=$($GeneratedTable.DataRowCount), expected=$($ExpectedTable.DataRowCount).")
    }

    $rowCountToCompare = [Math]::Min($GeneratedTable.DataRowCount, $ExpectedTable.DataRowCount)
    for ($rowIndex = 0; $rowIndex -lt $rowCountToCompare; $rowIndex++) {
        for ($columnIndex = 0; $columnIndex -lt $SchemaHeaders.Count; $columnIndex++) {
            $generatedValue = Normalize-CellText $GeneratedTable.DataRows[$rowIndex][$columnIndex]
            $expectedValue = Normalize-CellText $ExpectedTable.DataRows[$rowIndex][$columnIndex]

            if ($generatedValue -ne $expectedValue) {
                $differenceCount++
                $contentDifferenceCount++
                if ($differencePreview.Count -lt $storedDifferenceLimit) {
                    $differencePreview.Add((New-DifferencePreview -Section 'Content' -ExcelRow ($rowIndex + 2) -ColumnNumber ($columnIndex + 1) -ColumnName $SchemaHeaders[$columnIndex] -GeneratedValue $generatedValue -ExpectedValue $expectedValue))
                }
            }
        }
    }

    return [pscustomobject]@{
        HeaderErrors = $headerErrors.ToArray()
        StructuralErrors = $structuralErrors.ToArray()
        Errors = @($headerErrors.ToArray() + $structuralErrors.ToArray())
        DifferenceCount = $differenceCount
        SchemaDifferenceCount = $schemaDifferenceCount
        ContentDifferenceCount = $contentDifferenceCount
        DifferencePreview = $differencePreview.ToArray()
    }
}

function Write-DiffReport {
    param(
        [string]$GeneratedPath,
        [string]$ExpectedPath,
        [object]$GeneratedTable,
        [object]$ExpectedTable,
        [string[]]$HeaderErrors,
        [string[]]$StructuralErrors,
        [string[]]$AddressErrors,
        [object[]]$DifferencePreview,
        [int]$DifferenceCount,
        [int]$SchemaDifferenceCount,
        [int]$ContentDifferenceCount
    )

    $repoRoot = Split-Path -Parent $PSScriptRoot
    $reportDirectory = Join-Path $repoRoot 'logs\validation'
    [System.IO.Directory]::CreateDirectory($reportDirectory) | Out-Null

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $reportPath = Join-Path $reportDirectory "normalized-workbook-diff-$timestamp.txt"

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Generated Workbook: $GeneratedPath")
    $lines.Add("Expected Workbook: $ExpectedPath")
    $lines.Add("Generated Worksheet: $($GeneratedTable.WorksheetName)")
    $lines.Add("Expected Worksheet: $($ExpectedTable.WorksheetName)")
    $lines.Add("Generated Header Count: $($GeneratedTable.HeaderActualColumnCount)")
    $lines.Add("Expected Header Count: $($ExpectedTable.HeaderActualColumnCount)")
    $lines.Add("Generated Data Rows: $($GeneratedTable.DataRowCount)")
    $lines.Add("Expected Data Rows: $($ExpectedTable.DataRowCount)")
    $lines.Add("Difference Count: $DifferenceCount")
    $lines.Add("Schema Difference Count: $SchemaDifferenceCount")
    $lines.Add("Content Difference Count: $ContentDifferenceCount")
    $lines.Add('')

    $schemaPreview = @($DifferencePreview | Where-Object { $_.Section -eq 'Schema' })
    $contentPreview = @($DifferencePreview | Where-Object { $_.Section -eq 'Content' })

    $lines.Add('Schema:')
    if ($HeaderErrors.Count -gt 0) {
        foreach ($errorText in $HeaderErrors) {
            $lines.Add(" - $errorText")
        }
    }
    if ($schemaPreview.Count -gt 0) {
        foreach ($difference in $schemaPreview) {
            $lines.Add(" - [$($difference.Section)] row=$($difference.ExcelRow) col=$($difference.ColumnNumber) '$($difference.ColumnName)'")
            $lines.Add("   generated: '$($difference.GeneratedValue)'")
            $lines.Add("   expected : '$($difference.ExpectedValue)'")
        }
    }
    if ($HeaderErrors.Count -eq 0 -and $schemaPreview.Count -eq 0) {
        $lines.Add(' - none')
    }
    $lines.Add('')

    $lines.Add('Structural:')
    if ($StructuralErrors.Count -gt 0) {
        foreach ($errorText in $StructuralErrors) {
            $lines.Add(" - $errorText")
        }
    }
    else {
        $lines.Add(' - none')
    }
    $lines.Add('')

    $lines.Add('Content:')
    if ($contentPreview.Count -gt 0) {
        foreach ($difference in $contentPreview) {
            $lines.Add(" - [$($difference.Section)] row=$($difference.ExcelRow) col=$($difference.ColumnNumber) '$($difference.ColumnName)'")
            $lines.Add("   generated: '$($difference.GeneratedValue)'")
            $lines.Add("   expected : '$($difference.ExpectedValue)'")
        }
    }
    else {
        $lines.Add(' - none')
    }
    $lines.Add('')

    $lines.Add('AddressValidation:')
    if ($AddressErrors.Count -gt 0) {
        foreach ($errorText in $AddressErrors) {
            $lines.Add(" - $errorText")
        }
    }
    else {
        $lines.Add(' - none')
    }

    [System.IO.File]::WriteAllLines($reportPath, $lines)
    return $reportPath
}

$result = $null

try {
    $generatedPath = (Resolve-Path -LiteralPath $GeneratedWorkbookPath).Path
    $expectedPath = (Resolve-Path -LiteralPath $ExpectedWorkbookPath).Path

    $schemaHeaders = Get-ExpectedOutputHeaders
    $expectedHeaderCount = $schemaHeaders.Count

    $generatedTable = Get-WorkbookTable -WorkbookPath $generatedPath -ExpectedColumnCount $expectedHeaderCount
    $expectedTable = Get-WorkbookTable -WorkbookPath $expectedPath -ExpectedColumnCount $expectedHeaderCount

    $comparison = Compare-WorkbookTables -GeneratedTable $generatedTable -ExpectedTable $expectedTable -SchemaHeaders $schemaHeaders
    $addressErrors = @(
        @(Test-AddressColumns -Table $generatedTable -WorkbookRole 'Generated workbook' -AllowedTransshipmentValues $AllowedTransshipmentValues)
        @(Test-AddressColumns -Table $expectedTable -WorkbookRole 'Expected workbook' -AllowedTransshipmentValues $AllowedTransshipmentValues)
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $addressErrors = @($addressErrors)

    $allErrors = @($comparison.Errors + $addressErrors)
    $errorCount = $allErrors.Count + $comparison.ContentDifferenceCount
    $isSuccess = ($allErrors.Count -eq 0 -and $comparison.DifferenceCount -eq 0)
    $diffReportPath = ''
    $failureCategories = New-Object System.Collections.Generic.List[string]

    if ($comparison.HeaderErrors.Count -gt 0 -or $comparison.SchemaDifferenceCount -gt 0) {
        $failureCategories.Add('Schema')
    }

    if ($comparison.StructuralErrors.Count -gt 0) {
        $failureCategories.Add('Structural')
    }

    if ($comparison.ContentDifferenceCount -gt 0) {
        $failureCategories.Add('Content')
    }

    if ($addressErrors.Count -gt 0) {
        $failureCategories.Add('AddressValidation')
    }

    if (-not $isSuccess) {
        $diffReportPath = Write-DiffReport -GeneratedPath $generatedPath -ExpectedPath $expectedPath -GeneratedTable $generatedTable -ExpectedTable $expectedTable -HeaderErrors $comparison.HeaderErrors -StructuralErrors $comparison.StructuralErrors -AddressErrors $addressErrors -DifferencePreview $comparison.DifferencePreview -DifferenceCount $comparison.DifferenceCount -SchemaDifferenceCount $comparison.SchemaDifferenceCount -ContentDifferenceCount $comparison.ContentDifferenceCount
    }

    $result = [pscustomobject]@{
        Status = if ($isSuccess) { 'PASS' } else { 'FAIL' }
        IsMatch = $isSuccess
        GeneratedWorkbookPath = $generatedPath
        ExpectedWorkbookPath = $expectedPath
        GeneratedWorksheetName = $generatedTable.WorksheetName
        ExpectedWorksheetName = $expectedTable.WorksheetName
        ExpectedHeaderCount = $expectedHeaderCount
        GeneratedHeaderCount = $generatedTable.HeaderActualColumnCount
        ExpectedWorkbookHeaderCount = $expectedTable.HeaderActualColumnCount
        GeneratedDataRowCount = $generatedTable.DataRowCount
        ExpectedDataRowCount = $expectedTable.DataRowCount
        HeaderValidationPassed = ($comparison.HeaderErrors.Count -eq 0)
        AddressValidationPassed = ($addressErrors.Count -eq 0)
        WorkbookComparisonPassed = ($comparison.DifferenceCount -eq 0 -and $comparison.StructuralErrors.Count -eq 0)
        StructuralValidationPassed = ($comparison.StructuralErrors.Count -eq 0)
        ErrorCount = $errorCount
        DifferenceCount = $comparison.DifferenceCount
        PreviewDifferenceCount = $comparison.DifferencePreview.Count
        FailureCategories = @($failureCategories)
        DiffReportPath = $diffReportPath
        AllowedTransshipmentValues = @($AllowedTransshipmentValues)
        Errors = @($allErrors)
        DifferencePreview = @($comparison.DifferencePreview)
    }
}
catch {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $reportDirectory = Join-Path $repoRoot 'logs\validation'
    [System.IO.Directory]::CreateDirectory($reportDirectory) | Out-Null

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $reportPath = Join-Path $reportDirectory "normalized-workbook-diff-$timestamp.txt"
    $message = $_.Exception.Message
    [System.IO.File]::WriteAllLines($reportPath, @(
        "Generated Workbook: $GeneratedWorkbookPath",
        "Expected Workbook: $ExpectedWorkbookPath",
        '',
        'Validator failed before comparison completed.',
        "Error: $message"
    ))

    $result = [pscustomobject]@{
        Status = 'FAIL'
        IsMatch = $false
        GeneratedWorkbookPath = $GeneratedWorkbookPath
        ExpectedWorkbookPath = $ExpectedWorkbookPath
        GeneratedWorksheetName = ''
        ExpectedWorksheetName = ''
        ExpectedHeaderCount = (Get-ExpectedOutputHeaders).Count
        GeneratedHeaderCount = 0
        ExpectedWorkbookHeaderCount = 0
        GeneratedDataRowCount = 0
        ExpectedDataRowCount = 0
        HeaderValidationPassed = $false
        AddressValidationPassed = $false
        WorkbookComparisonPassed = $false
        StructuralValidationPassed = $false
        ErrorCount = 1
        DifferenceCount = 0
        PreviewDifferenceCount = 0
        FailureCategories = @('Execution')
        DiffReportPath = $reportPath
        AllowedTransshipmentValues = @($AllowedTransshipmentValues)
        Errors = @($message)
        DifferencePreview = @()
    }
}

$result
