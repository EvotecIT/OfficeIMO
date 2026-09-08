# Service incident runbook

Owner: Operations \| Applies to: the Northwind request portal

> [!WARNING] Before changing the service
> Confirm impact and name the incident owner. Record evidence before restarting components.

## 1. Establish the situation

- Record the first observed failure and the affected user groups.
- Check whether a deployment or configuration change preceded the failure.
- Open an incident record and agree the next update time.

## 2. Choose the next check

| Observation | Next check | Owner |
| --- | --- | --- |
| All users affected | Service health and dependencies | Operations |
| One group affected | Access and routing | Support |
| Failure after a change | Deployment evidence and rollback readiness | Engineering |

## 3. Communicate

```text
Impact: <what users cannot do>
Owner: <named coordinator>
Next action: <specific check>
Next update: <time>
```

## 4. Close with evidence

Confirm recovery with an affected user, record the timeline, and assign follow-up actions.
