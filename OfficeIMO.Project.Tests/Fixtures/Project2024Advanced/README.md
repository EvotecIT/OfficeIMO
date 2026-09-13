# Advanced Project fixtures

These synthetic schedules were created for OfficeIMO with Microsoft Project 16.0, build 16.0.20326.20132. The manifest records XML hashes and redistribution terms. `context.json` retains the original automated producer context. Fixtures contain no customer data.

| Files | Preparation and tested contract |
| --- | --- |
| `calendars.xml`, `rates.xml` | `Build/Project/New-ProjectAdvancedFixtures.ps1`: differing resource calendars, delayed assignments, availability, dated rates, material usage, and cost accrual |
| `effort-before.xml`, `effort.xml` | Same script: fixed-unit/work/duration tasks before and after adding another resource |
| `progress.xml`, `contours.xml`, `leveling.xml` | Same script: actual/overtime work, status date and baseline, named contours, and explicit application leveling |
| `fields.xml`, `rollups.xml` | Same script with `Set-ProjectCustomFieldFixture.ps1`: local formulas, modern lookup values, and summary rollups |
| `outline.xml` | `New-ProjectOutlineFixture.ps1`: task Outline Code 1, two-level `Design.01` path, masks, shared value IDs/GUIDs, and complete-path/table restrictions |
| `split.xml` | Application split of an OfficeIMO-authored two-task schedule: task 2 is interrupted from 2026-10-06 08:00 to 2026-10-07 08:00, then exported by Microsoft Project. Tests use the producer's remaining-work intervals, including the zero-work gap |
| `recurrence.xml` | Application UI: weekly review, one-day duration, every Monday, three occurrences from 2026-09-11. Project creates occurrences on September 14, 21, and 28, then the interoperability tool exports XML |
| `external-owner.xml`, `external-consumer.xml` | Application-created linked projects with a local dependency on an external task. The consumer's external project path is sanitized to relative `owner.mpp`; task identities and dates remain producer values |

The recurrence XML contains expanded children and recurring markers, not the native editable recurrence rule. The children in this fixture are manual tasks with explicit dates and start-no-earlier-than constraints. External fixtures are inert: tests supply documents through a resolver and never open the named external file automatically.

Regenerating a fixture can change GUIDs, producer timestamps, caches, and application defaults. Review semantic changes and update the manifest hashes deliberately. Application readback and the supported operation boundaries are documented in [the Project matrix](../../../OfficeIMO.Project/SUPPORT.md).
