# Review a mixed-format folder before distribution

Use [Example-MixedFormat-ConversionReview.ps1](Example-MixedFormat-ConversionReview.ps1) to convert an offline handover folder to PDFs and receive a file-by-file review. The sample contains Markdown, RTF and text documents describing the same synthetic service handover.

This workflow requires **PSWriteOffice 3.0.9** and **PowerShell 7.4 or newer**. The single-file conversion command has a different runtime boundary; Windows PowerShell 5.1 does not run this batch recipe. The website's public 3.0.8 snapshot documents its own available operations separately.

```powershell
./Examples/Workflows/Example-MixedFormat-ConversionReview.ps1 -OutputDirectory ./ConversionReview

# Use your own source folder. Keep output outside that folder.
./Examples/Workflows/Example-MixedFormat-ConversionReview.ps1 `
    -InputDirectory ./Handover -OutputDirectory ./ConversionReview
```

The recipe calls the module's conversion discovery and batch workflow. It uses two concurrent conversions, at most 250 discovered files, 16 MiB per source and 32 MiB per output by default. Change the named limit parameters deliberately for larger inputs. Unsupported files are reported as skipped.

The output has three parts:

| Path | Purpose |
| --- | --- |
| `Pdf/` | Distribution PDFs, retaining each source filename and extension before `.pdf` |
| `State/` | Engine-owned checkpoints for source/output hash verification and restart |
| `Reports/<run>/` | A fresh `Review.md`, `outcomes.csv`, `review.json` and, when present, `diagnostics.csv` for that run |

Read the failure summaries and warning/error diagnostics, then inspect every PDF. `Completed` means the converter produced an output; it does not establish acceptable fidelity. Review-required flags make skipped files and reported warnings visible. The full JSON includes the selected module and host versions, the assembly's source revision, the generating script hash, configured limits, available route limitations and structured per-file diagnostics. The engine's checkpoint receipts retain the source/output hashes used for restart verification.

Rerun the same command to verify completed outputs and reuse them. Correct a failed source, then add `-RetryFailed` to retry recorded failures. Completed files remain protected. A changed completed source, changed PDF or different conversion configuration requires a new reviewed output directory. Preserve the original files and the previous review; checkpoints are not a backup.

Use the reviewed `Completed` entries in `outcomes.csv` to select distribution files. An interrupted job can retain dot-prefixed staging files for recovery; the engine owns those files and completes their publication on restart.

When the engine returns a summary, the recipe writes the review before throwing for failed files, a cancelled batch or no completed PDFs. An empty folder still receives a review with zero counts. A fatal setup or input-limit error can stop the batch before its final summary exists. Ctrl+C stops the PowerShell pipeline and can prevent the run review from being written. Retain the completed PDFs and state, investigate the error, and rerun to verify the outputs and create a fresh review. Cancellation during report writing may leave an incomplete review directory.

The sample runs offline and does not send mail, publish content or install an optional provider. Other formats depend on the installed engine's executable routes and provider requirements. Review the route catalog and diagnostics for those inputs. These PDFs do not establish PDF/A, preservation of executable Office features, or archival compliance.
