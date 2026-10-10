# Release readiness report

**Release 4.2.0**  · candidate  `4.2.0-rc.3`  · decision due Friday

## Readiness summary

| Gate | Owner | Result | Evidence |
| --- | --- | --- | --- |
| Automated checks | Engineering | Passed | 1,284 tests, 0 failures |
| Package inspection | Release manager | Passed | Contents and metadata match |
| Upgrade rehearsal | Platform | Passed | 4.1.x to 4.2.0 on staging |
| Security review | Security | Passed | No open findings |
| Production approval | Service owner | Pending | Awaiting sign-off |

## Release checklist

- [x] Versioned artifacts are reproducible.
- [x] Upgrade and rollback notes are published.
- [x] Monitoring ownership is confirmed.
- [ ] Service owner records the approval.

## Rollout

Window: Tuesday 06:00-08:00 UTC
Rollback: Redeploy 4.1.8, about 15 minutes
Owner: Release manager, on call until noon

> [!WARNING] Decision required
> Production publication starts only after the service owner records approval.
