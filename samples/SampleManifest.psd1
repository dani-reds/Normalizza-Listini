@{
    Samples = @(
        @{
            Id = 'altro-listino-1'
            Description = 'Approved normalized baseline for ALTRO_LISTINO_1 generic workbook path.'
            InputRelativePath = 'samples/input/ALTRO_LISTINO_1.xlsx'
            ExpectedOutputRelativePath = 'samples/expected-output/ALTRO_LISTINO_1_normalized.xlsx'
            SourceType = 'xlsx'
            InvokeParameters = @{
                Carrier = 'MSC'
                Direction = 'Export'
                Reference = 'ALTRO-REF-001'
            }
            ExpectedDataRowCount = 96
            Notes = 'Direction intentionally omitted from InvokeParameters because it is not explicitly recoverable from the approved baseline files alone.'
        }
        @{
            Id = 'baseline1'
            Description = 'Approved normalized baseline for the Baseline1 workbook-family adapter.'
            InputRelativePath = 'samples/input/Baseline1.xlsx'
            ExpectedOutputRelativePath = 'samples/expected-output/Baseline1_normalized.xlsx'
            SourceType = 'xlsx'
            InvokeParameters = @{
                Carrier = 'ZIM'
                Direction = 'Export'
                Reference = 'BASELINE1-CANDIDATE'
                ValidityStartDate = '01/08/2025'
            }
            ExpectedDataRowCount = 104
            Notes = 'Direction intentionally omitted from InvokeParameters because it is not explicitly recoverable from the approved baseline files alone.'
        }
        @{
            Id = 'baseline2'
            Description = 'Approved normalized baseline for the Baseline2 workbook-family adapter.'
            InputRelativePath = 'samples/input/Baseline2.xlsx'
            ExpectedOutputRelativePath = 'samples/expected-output/Baseline2_normalized.xlsx'
            SourceType = 'xlsx'
            InvokeParameters = @{
                Carrier = 'ONE'
                Direction = 'Export'
                Reference = 'GOAN00967A'
            }
            ExpectedDataRowCount = 767
            Notes = 'Approved Baseline2 behavior: 2 tariff sheets only, ADDITIONAL & LOCAL CHARGES excluded from tariff generation, validity 28/07/2024 -> 31/08/2024, Ocean Freight - Containers, 20/40/40HC preserved as Cntr 20'' Box / Cntr 40'' Box / Cntrs 40'' HC, NO SERVICE OPTION skipped, remarks kept only as row Comment.'
        }
        @{
            Id = 'hapag-pdf-q2603goa03287-casasc-003'
            Description = 'Approved normalized baseline for the dedicated Hapag dry/std PDF adapter on the CASASC quotation sample.'
            InputRelativePath = 'samples/input/Quotation_Q2603GOA03287_CASASC_003.pdf'
            ExpectedOutputRelativePath = 'samples/expected-output/Quotation_Q2603GOA03287_CASASC_003_normalized.xlsx'
            SourceType = 'pdf'
            InvokeParameters = @{
                Carrier = 'HAPAG-LLOYD'
                Direction = 'Export'
                Reference = 'Q2603GOA03287'
                ValidityStartDate = '2026-04-01'
            }
            ExpectedDataRowCount = 8
            Notes = 'Approved Hapag dry/std PDF baseline for the dedicated adapter, limited to the validated port-to-port behavior represented by the approved expected workbook.'
        }
        @{
            Id = 'hapag-pdf-q2603goa02143-gdt-002-1'
            Description = 'Approved normalized baseline for the dedicated Hapag dry/std PDF adapter on the first GDT quotation sample.'
            InputRelativePath = 'samples/input/Quotation_Q2603GOA02143_GDT_002 (1).pdf'
            ExpectedOutputRelativePath = 'samples/expected-output/Quotation_Q2603GOA02143_GDT_002 (1)_normalized.xlsx'
            SourceType = 'pdf'
            InvokeParameters = @{
                Carrier = 'HAPAG-LLOYD'
                Direction = 'Export'
                Reference = 'Q2603GOA02143'
                ValidityStartDate = '2026-04-01'
            }
            ExpectedDataRowCount = 36
            Notes = 'Approved Hapag dry/std PDF baseline for the dedicated adapter after the approved endpoint-based inclusion rule for port-to-port processing.'
        }
        @{
            Id = 'hapag-pdf-q2603goa02149-gdt-002-1'
            Description = 'Approved normalized baseline for the Hapag PDF reefer quotation sample on the second GDT document.'
            InputRelativePath = 'samples/input/Quotation_Q2603GOA02149_GDT_002 (1).pdf'
            ExpectedOutputRelativePath = 'samples/expected-output/Quotation_Q2603GOA02149_GDT_002 (1)_normalized.xlsx'
            SourceType = 'pdf'
            InvokeParameters = @{
                Carrier = 'HAPAG-LLOYD'
                Direction = 'Export'
                Reference = 'Q2603GOA02149'
                ValidityStartDate = '2026-04-01'
            }
            ExpectedDataRowCount = 35
            Notes = 'Approved Hapag PDF reefer baseline represented by the approved expected workbook for quotation Q2603GOA02149.'
        }
    )
}
