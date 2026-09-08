# Configuration migration guide

Move an illustrative report job from separate output settings to a named delivery block.

Before you start, back up the configuration and use a temporary output folder for the first run.

## Map the settings

| Version 1 | Version 2 | Meaning |
| --- | --- | --- |
| outputPath | delivery.folder | Destination for generated files |
| filePrefix | delivery.prefix | Prefix applied to each report |
| overwrite | delivery.conflict | Use replace or fail explicitly |

## Before

```json
{
  "outputPath": "./reports",
  "filePrefix": "weekly-",
  "overwrite": false
}
```

## After

```json
{
  "delivery": {
    "folder": "./reports",
    "prefix": "weekly-",
    "conflict": "fail"
  }
}
```

## Verify the migration

- Run once with a small known input.
- Compare filenames, row counts and totals with the previous job.
- Repeat the run to confirm the chosen conflict behavior.
- Keep the original configuration until the scheduled run succeeds.
