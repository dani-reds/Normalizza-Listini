You are working on a local freight tariff normalization engine.

Goal:
Normalize ocean freight rate lists from xlsx/pdf into a fixed Excel output for TMS import, with maximum reliability and minimal regressions.

Project context:
- Main parser: Normalize-Listino.ps1
- Rules file: NormalizationRules.psd1
- The engine already supports multiple carrier/layout adapters.
- The output is a normalized Excel workbook with fixed TMS structure.
- UNLOCODE resolution, surcharge filtering, carrier/direction logic, and TMS output mapping are already part of the system.
- Existing working adapters must not be broken.

Operating rules:
1. Do NOT freely reinterpret tariff meaning if a deterministic rule is missing.
2. Always separate:
   - raw extraction/parsing
   - business rules
   - TMS mapping/output
   - validation
3. Make the smallest safe change possible.
4. Preserve backward compatibility unless explicitly told otherwise.
5. Before editing, inspect the relevant functions and explain:
   - current behavior
   - proposed change
   - regression risks
6. After editing, always run or propose validation steps.
7. If data is ambiguous, do not guess silently. Surface the ambiguity explicitly.
8. Never mix new carrier-specific logic into generic logic unless clearly justified.
9. Prefer adapter-specific logic over broad heuristics when the file layout is carrier-specific.
10. If you touch surcharge logic, validity logic, UNLOCODE resolution, arbitrary handling, or output column mapping, treat that as high-risk and verify carefully.

Business constraints:
- Input files can be xlsx or pdf.
- Layouts differ by carrier and sometimes by trade.
- The final output is not a generic extraction: it is a constrained TMS import format.
- When a file contains more commercial notes than the TMS model allows, keep only mapped or allowed items in the final output.
- Avoid silent changes in validity dates, direction inference, arbitrary treatment, surcharge filtering, transshipment assignment, and price detail naming.

Current HMM CIT assumptions:
- Carrier forced to HMM
- Direction forced to IMPORT
- Reference = ContractNumber-amdAmendNo
- Validity = Effective -> Contract Duration To
- Freight origin/destination codes treated as UNLOCODE
- Origin Arb / Dest Arb written as INLAND FREIGHT, not summed into ocean freight
- Transshipment Address = Via if present, otherwise Base Rate Code
- Subject-to surcharges mapped only if coherent with HMM Import matrix
- Allowed HMM CIT additionals: ECC, STF, ETS
- BUC excluded when marked inclusive
- Export-only items such as PEK/KLS excluded from import logic

If you believe one of these assumptions is wrong, do not silently change it.
Instead:
1. point to the code section,
2. explain why it is risky,
3. propose the smallest safe alternative,
4. describe how to validate it on existing normalized output.

Definition of done:
A task is complete only if:
- the targeted behavior is implemented,
- existing supported adapters are not obviously broken,
- the logic is readable and localized,
- validation steps are documented,
- any unresolved ambiguity is explicitly listed.

When working:
- First inspect the code paths involved.
- Then propose a short plan.
- Then implement in small steps.
- Then validate.
- Then summarize:
  1. files changed
  2. behavior changed
  3. tests/checks performed
  4. possible edge cases still open
