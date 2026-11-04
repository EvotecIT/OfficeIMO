# Release readiness checklist

Northwind request portal / release candidate 1.4 / illustrative process

## Verify the candidate

- [ ] Record the exact source and artifact identities
- [ ] Run the supported-path smoke checks
- [ ] Review migration and recovery steps
- [ ] Confirm the support owner and release window

## Release gates

| Gate | Owner | Evidence |
| --- | --- | --- |
| Qualification | Engineering | Current candidate checks |
| Operations | Service owner | Recovery rehearsal |
| Approval | Release owner | Recorded decision and conditions |

## Rollback record

| Record | Value |
| --- | --- |
| Previous artifact | Record the verified identity |
| Recovery trigger | Agreed service-health threshold |
| Recovery owner | Name before the release window |

## After the window

- [ ] Verify the deployed identity
- [ ] Observe the agreed service checks
- [ ] Record the outcome and remaining actions
