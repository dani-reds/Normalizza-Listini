# Ready-To-Use Task Prompts

Use these prompts with Codex when working on the tariff normalization engine.

## 1. New Adapter

```text
Use $listini-normalization.

Task:
Add support for a new carrier/layout adapter for this input file:
[insert file path]

Do not start coding immediately.
First:
1. inspect how adapter detection and dispatch currently work,
2. identify the closest existing adapter,
3. list the functions that should be mirrored, reused, or kept isolated,
4. explain the minimum set of new functions needed,
5. identify business ambiguities that cannot be inferred safely.

Then implement the adapter with localized changes only.
Do not refactor unrelated code.
Do not change existing adapter behavior unless strictly necessary.

When designing the adapter, keep parsing, business rules, TMS mapping, and validation clearly separated.
Prefer adapter-specific logic over broad heuristics.

At the end, return:
1. files changed
2. detection rule
3. extraction rule
4. mapping rule
5. validation steps run
6. regression risks
7. open ambiguities
```

## 2. Bug Fix

```text
Use $listini-normalization.

Task:
Investigate and fix this suspected bug in normalized output:
[describe the exact problem]

Workflow:
1. locate the exact code path producing the wrong field or value,
2. explain current behavior,
3. show why the current logic fails on this case,
4. propose the smallest safe fix,
5. check whether shared logic could affect other adapters,
6. implement only localized changes,
7. validate the target case,
8. report possible regressions explicitly.

Do not make speculative broad refactors.
Do not silently change business assumptions such as direction, validity, surcharge scope, arbitrary handling, TEU logic, or transshipment assignment.

At the end, return:
1. files changed
2. bug cause
3. behavior changed
4. validation steps run
5. residual risks or edge cases still open
```

## 3. Review Or Validation Only

```text
Use $listini-normalization.

Do not edit code yet.

Review the current implementation for:
- validity logic
- direction inference
- surcharge filtering
- arbitrary handling
- TEU/container evaluation mapping
- transshipment assignment
- adapter detection order
- TMS output mapping

Scope:
[insert adapter name, file path, or function area]

Return:
1. current behavior
2. possible errors or ambiguities
3. hidden regression risks
4. recommended validation tests
5. only after that, suggest code changes if really needed

Do not propose wide refactors unless the current design clearly creates repeated risk.
If an assumption looks wrong, point to the exact code section and explain the smallest safe alternative.
```
