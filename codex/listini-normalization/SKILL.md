---
name: listini-normalization
description: Use when working on the local freight tariff normalization engine for xlsx/pdf ocean rate lists, especially when adding carrier-specific adapters, changing surcharge or validity logic, handling UNLOCODE resolution, arbitrary charges, or TMS Excel output mapping.
---

# Listini Normalization

Use this skill when touching the tariff normalization engine in:

- [Normalize-Listino.ps1](C:\Users\DanieleRossi\Documents\New project\Normalize-Listino.ps1)
- [NormalizationRules.psd1](C:\Users\DanieleRossi\Documents\New project\NormalizationRules.psd1)

## Quick Start

Before editing:

1. Inspect the exact adapter or shared code path involved.
2. Separate parsing, business rules, TMS mapping, and validation.
3. Make the smallest safe change possible.
4. Treat surcharge logic, validity logic, arbitrary handling, UNLOCODE mapping, and output column mapping as high-risk.

## Non-Negotiable Rules

- Do not reinterpret tariff meaning freely when no deterministic rule exists.
- Do not silently widen surcharge scope.
- Do not silently change direction inference, validity logic, TEU evaluation, transshipment assignment, or arbitrary treatment.
- Prefer adapter-specific logic over broad heuristics for carrier-specific layouts.
- Preserve existing adapters unless the task explicitly requires shared refactoring.
- If ambiguity remains, surface it explicitly instead of guessing.

## Required Workflow

1. Explain current behavior.
2. Identify the smallest safe change.
3. Implement localized edits only.
4. Validate the target case.
5. Check for obvious regressions on at least one existing supported adapter when shared code was touched.
6. Summarize changed files, checks performed, and residual risks.

## What To Read Next

- For a compact everyday reminder: [CHEATSHEET_LISTINI.md](C:\Users\DanieleRossi\Documents\New project\codex\CHEATSHEET_LISTINI.md)
- For business constraints, validation checklist, and task templates: [workflow.md](C:\Users\DanieleRossi\Documents\New project\codex\listini-normalization\references\workflow.md)
- For ready-to-use prompts for new adapters, bug fixes, and review work: [task-prompts.md](C:\Users\DanieleRossi\Documents\New project\codex\listini-normalization\references\task-prompts.md)
- For current project state and HMM CIT assumptions: [CHATGPT_HANDOFF_LISTINI.md](C:\Users\DanieleRossi\Documents\New project\CHATGPT_HANDOFF_LISTINI.md)

## Done Criteria

A task is done only if:

- the target behavior is implemented,
- existing supported adapters are not obviously broken,
- the logic remains readable and localized,
- validation was run or clearly described,
- open ambiguities are listed explicitly.
