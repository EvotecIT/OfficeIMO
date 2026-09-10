# Project XML qualification evidence — 2026-09-10

This directory retains compact results from the synthetic Project 2024 corpus and the opt-in [verification tools](../../README.md). `context.json` records the measured source revision, environment, input hashes, budgets, and validation lanes. `scale-summary.json` contains all eight scenarios, with one warmup and three measured iterations each. All scenarios passed correctness, memory, and median-runtime budgets.

The [native reading and scheduling evidence](scheduling/README.md) qualifies the modern reader, explicit schedule calculation, assignment analysis, and their additional producer fixtures.

Application readbacks identify the Microsoft Project build and input hash. They cover authored and edited XML, a newly added root task, calendar exceptions, resource types, and the bounded native rewrite/name-edit experiments. The paired resource readbacks expose the application's cost-resource XML import discrepancy: native retains 300 while its own exported XML reopens with zero. These results qualify selected records and fixtures; they do not prove every application warning was absent.

`cancellation.json` records cooperative load/save cancellation and preservation of the destination and modified model. Its timings are single-run response checks, not benchmark distributions. Offline schema validation passed for authored and resource-authored XML against the Microsoft Project 2013 SDK schema using the verifier's explicit namespace alias. The vendor schema is not redistributed.

The seed-free minimal native-container experiment was rejected by Microsoft Project. Native MPP authoring, general native editing, legacy producer families, and cost-resource application amount fidelity remain unqualified. See the [support matrix](../../../../OfficeIMO.Project/SUPPORT.md) for the current public contract.

Bulky scale inputs, reexports, local package feeds, SDK extraction files, and intermediate binaries were removed after retaining these results. Reproduction creates new output directories through the verification tools.
