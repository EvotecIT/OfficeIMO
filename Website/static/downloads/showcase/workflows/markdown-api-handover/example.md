# Request API handover

A sample integration guide for the fictional Northwind request service.

## Create a request

Send a title, category, and requester reference. The service returns the request identifier and initial status.

```http
POST /requests
Content-Type: application/json

{
  "title": "Prepare a new workspace",
  "category": "equipment",
  "requester": "team-delivery"
}
```

## Field contract

| Field | Required | Meaning |
| --- | --- | --- |
| title | Yes | Short description of the requested outcome |
| category | Yes | Routing category agreed with the service owner |
| requester | Yes | Stable caller-owned reference |

## Read the response

```json
{
  "id": "REQ-1042",
  "status": "submitted"
}
```

> [!NOTE] Integration checklist
> Agree authentication, retry behavior, validation errors, and rate limits with the service owner. The payloads here describe a fictional API.

## Support handover

Record the integration owner, an escalation route, and a representative successful request before enabling a production caller.

## Handle an unsuccessful request

| Response | Caller action |
| --- | --- |
| 400 - Invalid payload | Correct the named field before sending another request |
| 401 - Authentication required | Check the configured identity and credential lifetime |
| 429 - Rate limited | Respect the agreed retry delay and avoid parallel retries |
| 503 - Service unavailable | Use a bounded retry policy and contact the service owner |

## Record a handover example

```text
Environment: test
Caller: workspace-intake
Request: REQ-1042
Result: submitted
Integration owner: Workplace team
Service contact: Operations queue
```

## Before enabling the caller

- Verify one accepted request and one rejected request.
- Agree how to detect and prevent duplicate submissions.
- Keep credentials and personal data out of diagnostic logs.
- Confirm the support route and the first review date.
