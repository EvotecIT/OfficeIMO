# Service request export dictionary

Schema 1.0 / one record per request / UTF-8 JSON Lines

## Fields

| Field | Type | Constraint | Meaning |
| --- | --- | --- | --- |
| request_id | string | Required; unique | Stable identifier, for example SR-2048 |
| status | string | new \| active \| waiting \| closed | Current workflow state |
| owner | string | Required while active | Team responsible for the next action |
| opened_at | datetime | ISO 8601 UTC | When the request entered the service |
| age_days | integer | Zero or greater | Whole elapsed days at export time |

## Sample record

```json
{
  "request_id": "SR-2048",
  "status": "active",
  "owner": "Support",
  "opened_at": "2026-09-01T09:00:00Z",
  "age_days": 2
}
```

## Consumer rules

- Treat request\_id as an identifier, not a number.
- Preserve unknown fields when forwarding records.
- Reject negative age\_days and report the source record.
- Use opened\_at for time-based analysis; age\_days is a snapshot.

> [!NOTE] Example contract
> These fields describe a fictional service export. Replace the catalog with your application's actual schema.
