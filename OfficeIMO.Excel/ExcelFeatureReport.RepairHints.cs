namespace OfficeIMO.Excel {
    public sealed partial class ExcelFeatureReport {
        /// <summary>
        /// Returns actionable repair or routing hints for a blocked preflight capability.
        /// </summary>
        /// <param name="capability">Capability to explain.</param>
        public IReadOnlyList<ExcelPreflightRepairHint> GetRepairHints(ExcelPreflightCapability capability) {
            if (Can(capability)) {
                return Array.Empty<ExcelPreflightRepairHint>();
            }

            var hints = new List<ExcelPreflightRepairHint>();
            foreach (ExcelFeatureFinding finding in GetCapabilityFindings(capability)) {
                AddRepairHints(hints, capability, finding);
            }

            return hints
                .GroupBy(hint => hint.FeatureName + "\u001f" + hint.Action + "\u001f" + (hint.Command ?? string.Empty), StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
        }

        private IEnumerable<ExcelFeatureFinding> GetCapabilityFindings(ExcelPreflightCapability capability) {
            switch (capability) {
                case ExcelPreflightCapability.ReadWorkbookData:
                    return UnsupportedFeatures.Concat(FindFeatures("Non-worksheet sheets"));
                case ExcelPreflightCapability.EditCellValues:
                    return UnsupportedFeatures.Concat(FindFeatures("Digital signatures"));
                case ExcelPreflightCapability.EditWorkbookStructure:
                case ExcelPreflightCapability.BindTemplate:
                    return UnsupportedFeatures.Concat(PreservedFeatures);
                case ExcelPreflightCapability.UseCachedFormulaValues:
                    return FindFeatures("Missing formula caches", "Dirty formula caches", "Workbook recalculation requests");
                case ExcelPreflightCapability.CalculateFormulas:
                    return FindFeatures("Formula calculation blockers", "Formula dependency issues");
                case ExcelPreflightCapability.ExportPdfReport:
                    return UnsupportedFeatures
                        .Concat(PreservedFeatures.Where(IsPdfExportWorkbookBlocker))
                        .Concat(FindFeatures(
                            "PDF-missing formula caches",
                            "PDF-dirty formula caches",
                            "PDF-workbook recalculation requests",
                            "PDF-unsupported charts",
                            "PDF-unreadable charts",
                            "PDF-unsupported images",
                            "PDF-unsupported hyperlinks",
                            "PDF-unrendered drawing shapes",
                            "PDF-unsupported print areas",
                            "PDF-unsupported print titles",
                            "PDF-unsupported header/footer formatting",
                            "PDF-unrendered pivot tables",
                            "PDF-unrendered sparklines",
                            "Non-worksheet sheets"));
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported Excel preflight capability.");
            }
        }

        private static void AddRepairHints(List<ExcelPreflightRepairHint> hints, ExcelPreflightCapability capability, ExcelFeatureFinding finding) {
            switch (finding.Name) {
                case "Missing formula caches":
                case "Dirty formula caches":
                case "Workbook recalculation requests":
                case "PDF-missing formula caches":
                case "PDF-dirty formula caches":
                case "PDF-workbook recalculation requests":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Refresh cached formula values before trusting reads or exports.",
                        "Calculate(), InvalidateFormulas(), or save with ForceFullCalculationOnOpen",
                        "Use OfficeIMO calculation when every formula is supported; otherwise open the workbook in Excel-compatible software and save after recalculation."));
                    break;
                case "Formula calculation blockers":
                case "Formula dependency issues":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Route calculation to Excel-compatible software or simplify unsupported formulas before using OfficeIMO calculation.",
                        "InspectFormulas()",
                        "Clean cached values may still be usable for read/export workflows when caches are present and current."));
                    break;
                case "Digital signatures":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Work on an unsigned copy or avoid edits that would invalidate the signature.",
                        null,
                        "OfficeIMO preserves signature metadata but cannot keep a signature valid after package mutation."));
                    break;
                case "PDF-unrendered pivot tables":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Materialize the pivot output as ordinary worksheet cells before first-party PDF export.",
                        "AddPivotTable(...), then refresh/open in Excel-compatible software and save",
                        "OfficeIMO preserves and authors pivot metadata, but the first-party PDF path renders worksheet cells, not pivot UI metadata."));
                    break;
                case "PDF-unrendered sparklines":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Replace sparklines with rendered chart/image/table indicators for first-party PDF export.",
                        "AddDashboardChart(...)",
                        null));
                    break;
                case "PDF-unsupported images":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Convert worksheet/header/footer images to PNG or JPEG before PDF export.",
                        null,
                        "Unsupported or invalid image bytes are skipped by the first-party PDF image writer."));
                    break;
                case "PDF-unsupported hyperlinks":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Use absolute external hyperlinks or visible exported targets.",
                        "SetHyperlink(...)",
                        "Relative links and links to skipped targets are not emitted by the first-party PDF hyperlink writer."));
                    break;
                case "PDF-unsupported print areas":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Use a single contiguous print area or export sheets/ranges separately.",
                        "SetPrintArea(...)",
                        null));
                    break;
                case "PDF-unsupported print titles":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Use repeat-title rows only or pre-render repeated columns into the export range.",
                        "SetPrintTitles(...)",
                        null));
                    break;
                case "PDF-unsupported header/footer formatting":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Simplify header/footer formatting to supported text, font, and size tokens.",
                        "SetHeaderFooter(...)",
                        null));
                    break;
                case "PDF-unsupported charts":
                case "PDF-unreadable charts":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Use supported chart types or export the chart through Excel-compatible software.",
                        "AddDashboardChart(...)",
                        null));
                    break;
                case "Non-worksheet sheets":
                    hints.Add(new ExcelPreflightRepairHint(
                        capability,
                        finding.Name,
                        "Move needed content to normal worksheets before worksheet-only read/export workflows.",
                        null,
                        "Chartsheets and similar sheet parts are preserve-only for worksheet workflows."));
                    break;
                default:
                    if (finding.SupportLevel == ExcelFeatureSupportLevel.Preserved || finding.SupportLevel == ExcelFeatureSupportLevel.Unsupported) {
                        hints.Add(new ExcelPreflightRepairHint(
                            capability,
                            finding.Name,
                            "Use a preserve-only workflow, remove the feature, or route this workbook through Excel-compatible software.",
                            "CopyPackage(...)",
                            finding.Note));
                    }
                    break;
            }
        }
    }
}
