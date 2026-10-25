# Saved conversion job examples

Copy this complete directory to a writable location before running the examples.
It contains three small job definitions and their inputs. Outputs and checkpoints
are created beside the jobs, outside `Inputs`. Install `OfficeIMO.Tool` for the
terminal commands and PSWriteOffice for the PowerShell commands. The commands use
the installed public entry points.

## Convert a folder and reuse verified outputs

```text
officeimo workflow job plan ./folder.officeimojob.json
officeimo workflow job run ./folder.officeimojob.json --format ndjson
officeimo workflow job run ./folder.officeimojob.json
```

The plan selects the two TXT files and reports the other inputs as skipped. It
creates no output or checkpoint folders. The first run writes
`Outputs/Folder/brief.txt.pdf` and `actions.txt.pdf`. Reopen both PDFs and check
their text. The second run reports two reused artifacts after verifying the
sources, rendering settings and saved output hashes. It does not overwrite them.

In Studio's Convert workbench, expand the local-folder batch section, choose
**Load job**, select `folder.officeimojob.json`, inspect **Plan**, then run it.
**Copy CLI** and **Copy PowerShell** refer to that saved definition. Edit the
settings and save before copying a command for the changed job. Local-folder jobs
require an edition with ongoing filesystem access.

PowerShell 7.4 or newer executes the same definition:

```powershell
Import-Module PSWriteOffice
Invoke-OfficeConversionJob -Path ./folder.officeimojob.json -Plan
Invoke-OfficeConversionJob -Path ./folder.officeimojob.json -ItemResults -SummaryVariable run
$run
```

## Prepare a delivery PDF

```text
officeimo workflow job plan ./delivery.officeimojob.json
officeimo workflow job run ./delivery.officeimojob.json
```

This job converts the self-contained HTML input with its supported `Faithful` profile.
Expected output is `Outputs/Delivery/delivery.html.pdf`, with the delivery title,
table and review instruction. Inspect the rendered PDF and any conversion
diagnostics. A successful save/reopen does not establish visual fidelity or a
PDF conformance certification. The job's relative paths work even when it is
invoked from another working directory using its absolute filename.

## Recover a partial batch

```text
officeimo workflow job plan ./retry.officeimojob.json
officeimo workflow job run ./retry.officeimojob.json --format ndjson
```

`Inputs/repair.docx` deliberately contains invalid document bytes. Planning
selects its known route but does not decode or certify the document. Execution
keeps the successful `Outputs/Retry/brief.txt.pdf`, reports a failed DOCX item and
returns a failure exit code. Fix only the failed source. For a self-contained
repair using PowerShell:

```powershell
Import-Module PSWriteOffice
Remove-Item -LiteralPath ./Inputs/repair.docx
New-OfficeWord -Path ./Inputs/repair.docx {
    WordParagraph 'The failed input has been repaired.'
}
Invoke-OfficeConversionJob -Path ./retry.officeimojob.json -RetryFailed -ItemResults -SummaryVariable run
$run
```

The retry reports one verified reused PDF and one newly completed PDF. Reopen
both files. Replacing the deliberately invalid fixture above is only appropriate
in the writable example copy. For real jobs, preserve source versions according
to the application's normal document policy.

Cancellation retains completed outputs. Restart the same checkpointed job to
verify and reuse them; add `--retry-failed` or `-RetryFailed` for recorded failed
items. If an output has no verified receipt or its hash changed, inspect it and
choose separate output/checkpoint locations rather than deleting evidence to
force reuse. Ordinary jobs without checkpoints obey their selected conflict
policy. An empty selection is an error unless explicitly accepted with
`--allow-empty` or `-AllowEmpty`.

Job definitions contain paths and settings, so review them before execution.
Keep password values outside JSON and use environment-variable references when
needed. The portable contract rejects unknown settings and native options it
cannot reproduce.
