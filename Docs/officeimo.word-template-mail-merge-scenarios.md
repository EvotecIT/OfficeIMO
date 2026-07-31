# OfficeIMO.Word Template And Mail-Merge Scenario Matrix

This matrix tracks the public template and mail-merge workflows that OfficeIMO.Word can prove today, where the proof lives, and what remains partial. It intentionally focuses on non-PDF `.docx` automation.

## How To Run The Proof

```powershell
dotnet run --project OfficeIMO.Examples\OfficeIMO.Examples.csproj -f net8.0 -- --word-mail-merge-workflows
dotnet run --project OfficeIMO.Examples\OfficeIMO.Examples.csproj -f net8.0 -- --word-market-readiness
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net8.0 --filter "FullyQualifiedName~Test_MailMerge"
```

The workflow runner writes invoice, grouped table report, proposal, review letter, header/footer approval package, and form-fill documents to the examples output folder. The form-fill workflow also writes content-control validation diagnostics in JSON and Markdown. The market-readiness gallery writes clean and blocked template preflight reports.

## Scenario Matrix

| Scenario | Public API | Proof | Status | Current limit |
| --- | --- | --- | --- | --- |
| Merge fields | `WordMailMerge.Execute`, `WordMailMerge.ExecuteBatch`, `WordMailMerge.PreflightTemplate` | `Test_MailMerge_ReplacesFields`, `Test_MailMerge_ExecuteBatchCreatesOutputsAndKeepsTemplateUnchanged`, `Test_MailMerge_ComplexSplitRunFieldsPreserveResultFormattingWhenKeepingFields`, `MailMergeInvoiceWorkflow.docx` | Covered | Imported-template proof does not cover every unusual field-code shape. |
| Conditional blocks | `WordMailMerge.ExecuteConditionalBlocks`, `WordMailMerge.PreflightTemplate` | `Test_MailMerge_ConditionalBlocksCanIncludeBodyContentAndMergeFields`, `Test_MailMerge_ConditionalBlocksCanRunInsideHeadersAndFooters`, `Test_MailMerge_ConditionalBlocksCanKeepOrRemoveSectionRegions`, `MailMergeProposalWorkflow.docx`, `MailMergeReviewLetterWorkflow.docx` | Covered | Imported Word-authored conditional-template proof is bounded to the listed fixtures. |
| Repeated table rows | `WordMailMerge.ExecuteTableRows` | `Test_MailMerge_RepeatsTableRowsAndRemovesTemplateRow`, `MailMergeInvoiceWorkflow.docx` | Covered | Imported table formatting proof is bounded to the listed fixtures. |
| Grouped table rows | `WordMailMerge.ExecuteTableRowGroups` | `Test_MailMerge_RepeatsGroupedTableRowsAndPreservesFormatting`, `MailMergeGroupedTableWorkflow.docx`, `MailMergeGroupedTableWorkflow.Preflight.md` | Covered | Imported Word-authored grouped-table proof is bounded to the listed fixtures. |
| Repeated body blocks | `WordMailMerge.ExecuteRepeatingBlocks` | `Test_MailMerge_RepeatingBlocksCloneBodyContentTablesAndFormatting`, `MailMergeReviewLetterWorkflow.docx` | Covered | Mixed imported body-content proof is bounded to the listed fixtures. |
| Nested regions | `WordMailMerge.ExecuteRepeatingBlockRegions`, `WordMailMergeBlockData` | `Test_MailMerge_RepeatingBlockRegionsBindNestedData`, `Test_MailMerge_NestedRegionsPreserveTableCellFieldFormatting`, `MailMergeProposalWorkflow.docx` | Covered | Deeply nested error-report behavior is not comprehensively fixture-backed. |
| Section regions | `WordMailMerge.ExecuteConditionalBlocks`, `WordMailMerge.ExecuteRepeatingBlocks`, `WordMailMerge.PreflightTemplate` | `Test_MailMerge_RepeatingBlockRegionsPreserveSectionBreakProperties`, `Test_MailMerge_ConditionalBlocksCanKeepOrRemoveSectionRegions`, `Test_MailMerge_WordAuthoredMultiSectionTemplateCanBePreflightedAndBound`, `word-authored-multi-section-template.docx` | Covered | Imported section and page-setup proof is bounded to the listed shapes. |
| Headers and footers | `WordMailMerge.Execute`, `WordMailMerge.ExecuteConditionalBlocks`, `WordMailMerge.InspectTemplate`, `WordMailMerge.PreflightTemplate` | `Test_MailMerge_ConditionalBlocksCanRunInsideHeadersAndFooters`, `Test_MailMerge_PreflightTemplateSeesHeaderFooterTemplateMarkersAfterSaveLoad`, `MailMergeHeaderFooterWorkflow.docx`, `MailMergeHeaderFooterWorkflow.Preflight.md` | Covered | Imported header/footer proof is bounded to the listed fixtures. |
| Table cells | `WordMailMerge.Execute`, `WordMailMerge.ExecuteConditionalBlocks`, `WordMailMerge.ExecuteTableRows`, `WordMailMerge.ExecuteRepeatingBlockRegions`, `FillContentControlValues` | `Test_MailMerge_ConditionalBlocksCanIncludeTableCellContentAndMergeFields`, `Test_MailMerge_ConditionalBlocksCanRemoveTableCellContent`, `Test_MailMerge_NestedRegionsPreserveTableCellFieldFormatting`, `Test_ContentControlForm_WordAuthoredFixtureCanValidateFillAndExtractValues`, `word-authored-content-control-form.docx`, `MailMergeInvoiceWorkflow.docx` | Covered | Imported table-cell proof is bounded to the listed Word-authored shapes. |
| Content controls | `ValidateContentControlValues`, `FillContentControlValues`, `ExtractContentControlValues`, `WordMailMerge.RefreshContentControlDataBindings`, `WordMailMerge.ExecuteContentControlDataBindings` | `Test_MailMerge_ConditionalBlocksCanRunInsideBlockContentControls`, `Test_MailMerge_RefreshesContentControlDataBindingsFromCustomXml`, `Test_MailMerge_ExecutesContentControlDataBindingsAndUpdatesCustomXml`, `Test_ContentControlFormValidationReportsMissingInvalidAndUnusedValues`, `Test_ContentControlForm_WordAuthoredFixtureCanValidateFillAndExtractValues`, `Test_MailMerge_ContentControlFormFillPreservesTextRunFormatting`, `word-authored-content-control-form.docx`, `MailMergeFormFillWorkflow.docx`, `MailMergeFormFillWorkflow.Validation.json`, `MailMergeFormFillWorkflow.Validation.md` | Covered | SDT mapping and bound-content-control proof is bounded to the listed fixtures. |
| Template diagnostics | `WordMailMerge.InspectTemplate`, `WordMailMerge.PreflightTemplate`, `WordTemplatePreflightReport` | `Test_MailMerge_PreflightTemplateReportsCapabilitiesAndSerializes`, `Test_MailMerge_PreflightTemplateSeparatesCapabilityDiagnostics`, `template-preflight.md`, `template-preflight-blocked.md` | Covered | Repair hints are generic capability diagnostics rather than scenario-specific guidance. |

## Public Workflow Examples

| Workflow | Output | What it demonstrates |
| --- | --- | --- |
| Invoice | `MailMergeInvoiceWorkflow.docx` | Merge fields, repeated table rows, template preflight, standard save path. |
| Grouped table report | `MailMergeGroupedTableWorkflow.docx`, `MailMergeGroupedTableWorkflow.Preflight.md` | Group/detail table rows, grouped totals, merge-field preflight, and final body-field binding. |
| Proposal | `MailMergeProposalWorkflow.docx` | Merge fields, conditional blocks, nested repeated regions, template preflight. |
| Review letter | `MailMergeReviewLetterWorkflow.docx` | Merge fields, conditional blocks, repeated body blocks, generated comment context. |
| Header/footer approval package | `MailMergeHeaderFooterWorkflow.docx`, `MailMergeHeaderFooterWorkflow.Preflight.md` | Header/footer-hosted merge fields, conditional header block, repeated footer block, and template preflight. |
| Form fill | `MailMergeFormFillWorkflow.docx`, `MailMergeFormFillWorkflow.Validation.json`, `MailMergeFormFillWorkflow.Validation.md` | Content-control validation, reusable JSON/Markdown diagnostics, fill, extraction, and generated diagnostics. |

## Current Limits

- Section-region support has fixture proof for conditional include/remove flows, repeated regions that preserve section break, orientation, and margin properties, and a first Word-authored multi-section conditional template with Word-created merge fields. Broader imported section and page-setup shapes remain open.
- Content-control form support has fixture proof for OfficeIMO-authored forms and a Word-authored body/table-cell form with text, rich text, checkbox, date, dropdown, combo box, picture, and table-cell block SDTs. Broader imported SDT mapping and binding scenarios remain open.
- Formatting preservation has focused proof for simple and complex merge fields, split-run complex fields, repeated table rows, grouped table rows with a public workflow output, repeated body blocks, section-shaped regions, nested table-cell regions, header/footer-hosted template preflight and public workflow output, and content-control form fill. Broader imported-template fixtures remain open.
- PowerShell-friendly wrappers belong in PSWriteOffice and stay thin over the reusable OfficeIMO.Word behavior.
