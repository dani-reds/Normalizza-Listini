# Workflow

## Project Purpose

Normalize ocean freight listini from `xlsx` and `pdf` into a fixed Excel structure for TMS import.

This is not generic extraction. It is controlled normalization with business rules.

Core files:

- [Normalize-Listino.ps1](C:\Users\DanieleRossi\Documents\New project\Normalize-Listino.ps1)
- [NormalizationRules.psd1](C:\Users\DanieleRossi\Documents\New project\NormalizationRules.psd1)

Related context:

- [CHATGPT_HANDOFF_LISTINI.md](C:\Users\DanieleRossi\Documents\New project\CHATGPT_HANDOFF_LISTINI.md)
- [task-prompts.md](C:\Users\DanieleRossi\Documents\New project\codex\listini-normalization\references\task-prompts.md)

## Existing Supported Adapters

- generic Excel listino
- COSCO Far East workbook
- COSCO structured workbook
- COSCO IPAK PDF
- COSCO Canada PDF
- COSCO South America PDF
- COSCO IET workbook
- EVERGREEN FAK RVS workbook
- HAPAG-LLOYD quotation PDF
- HMM CIT workbook

## High-Risk Areas

- adapter detection order
- validity derivation
- direction inference
- UNLOCODE resolution
- arbitrary origin/destination handling
- transshipment assignment
- surcharge filtering by carrier + direction
- TEU/container evaluation mapping
- final TMS output column mapping

## Validation Checklist

Always check:

- Is carrier correct?
- Is direction correct?
- Are validity start/end dates coming from the intended fields?
- Are origin/destination codes trusted, mapped, or inferred?
- Are UNLOCODE overrides still winning over dirty raw names when needed?
- Are arbitrary charges intentionally separate or intentionally merged?
- Are only allowed surcharge names written to output?
- Are inclusive charges excluded intentionally?
- Are TEU-based charges aligned with TMS logic?
- Is transshipment populated only when logically justified?
- Did duplicate rows appear unexpectedly?
- Are required output columns still present?
- Did adapter routing order become unsafe?

## Task Templates

### New Adapter

First inspect:

1. adapter detection and dispatch path
2. closest existing adapter
3. reusable functions
4. minimum new functions needed

Then implement localized changes only and report:

- detection rule
- extraction rule
- mapping rule
- validation checklist
- regression risks

### Bug Fix

Follow this order:

1. locate the exact producing code path
2. explain current behavior
3. show why it fails on the target case
4. apply the smallest safe fix
5. check whether shared logic affects other adapters
6. summarize verification steps

### Review Without Edits

Review only:

- validity logic
- direction inference
- surcharge filtering
- arbitrary handling
- TEU/container evaluation mapping
- transshipment assignment

Return:

1. current behavior
2. possible errors or ambiguities
3. hidden regression risks
4. recommended tests
5. only then, suggested code changes if needed

If you want a ready-to-paste task prompt, use:

- [task-prompts.md](C:\Users\DanieleRossi\Documents\New project\codex\listini-normalization\references\task-prompts.md)

## HMM CIT Current Assumptions

Current implementation choices:

- Carrier forced to `HMM`
- Direction forced to `IMPORT`
- Reference = `ContractNumber-amdAmendNo`
- Validity = `Effective -> Contract Duration To`
- Freight sheet origin/destination codes treated as UNLOCODE
- `Origin Arb` and `Dest Arb` written as `INLAND FREIGHT`
- `Transshipment Address = Via`, otherwise `Base Rate Code`
- Subject-to surcharges written only if coherent with `HMM Import`
- Allowed CIT additionals currently: `ECC`, `STF`, `ETS`
- `BUC` excluded when marked inclusive
- `PEK` and `KLS` excluded from import logic

If changing any of the above:

1. point to the exact code section
2. explain why the current assumption is risky
3. propose the smallest safe alternative
4. describe how to validate it on the existing normalized output
