namespace OfficeIMO.Excel {
    /// <summary>
    /// Workflow-level OfficeIMO.Excel operation categories covered by feature preflight checks.
    /// </summary>
    public enum ExcelPreflightCapability {
        /// <summary>Read worksheet data, tables, and typed rows from the workbook.</summary>
        ReadWorkbookData,

        /// <summary>Edit existing cell values while preserving package parts OfficeIMO does not fully author yet.</summary>
        EditCellValues,

        /// <summary>Perform structure-changing edits such as adding/removing sheets, tables, drawings, pivots, or relationships.</summary>
        EditWorkbookStructure,

        /// <summary>Use cached formula values for reads, export, or downstream reporting.</summary>
        UseCachedFormulaValues,

        /// <summary>Calculate workbook formulas through OfficeIMO's lightweight evaluator.</summary>
        CalculateFormulas,

        /// <summary>Bind data into an Excel template and save the generated workbook.</summary>
        BindTemplate,

        /// <summary>Export a report workbook through the first-party OfficeIMO Excel-to-PDF path.</summary>
        ExportPdfReport
    }

    public sealed partial class ExcelFeatureReport {
        /// <summary>
        /// True when OfficeIMO can attempt read-oriented workbook data operations.
        /// </summary>
        public bool CanReadWorkbookData =>
            UnsupportedFeatures.Count == 0 &&
            FindFeatureCount("Non-worksheet sheets") == 0;

        /// <summary>
        /// True when OfficeIMO can attempt cell-value edits without known unsupported package features.
        /// </summary>
        public bool CanEditCellValues =>
            UnsupportedFeatures.Count == 0 &&
            FindFeatureCount("Digital signatures") == 0;

        /// <summary>
        /// True when OfficeIMO can attempt structure-changing workbook edits without preserve-only or unsupported feature blockers.
        /// </summary>
        public bool CanEditWorkbookStructure => !HasAdvancedFeatures;

        /// <summary>
        /// True when cached formula values can be trusted for read/export workflows.
        /// </summary>
        public bool CanUseCachedFormulaValues =>
            FindFeatureCount("Missing formula caches") == 0 &&
            FindFeatureCount("Dirty formula caches") == 0 &&
            FindFeatureCount("Workbook recalculation requests") == 0;

        /// <summary>
        /// True when OfficeIMO's lightweight evaluator can calculate all discovered formulas without known dependency issues.
        /// </summary>
        public bool CanCalculateFormulas =>
            FindFeatureCount("Formula calculation blockers") == 0 &&
            CountFormulaDependencyIssues(includeMissingCachedDependencyIssues: false) == 0;

        /// <summary>
        /// True when template binding can be attempted without preserve-only or unsupported advanced package features.
        /// </summary>
        public bool CanBindTemplate => !HasAdvancedFeatures;

        /// <summary>
        /// True when the workbook is suitable for the first-party report-grade Excel-to-PDF export path.
        /// </summary>
        public bool CanExportPdfReport =>
            !HasPdfExportWorkbookBlockers() &&
            FindFeatureCount("PDF-missing formula caches") == 0 &&
            FindFeatureCount("PDF-dirty formula caches") == 0 &&
            FindFeatureCount("PDF-workbook recalculation requests") == 0 &&
            FindFeatureCount("PDF-unsupported charts") == 0 &&
            FindFeatureCount("PDF-unreadable charts") == 0 &&
            FindFeatureCount("PDF-unsupported images") == 0 &&
            FindFeatureCount("PDF-unsupported hyperlinks") == 0 &&
            FindFeatureCount("PDF-unrendered drawing shapes") == 0 &&
            FindFeatureCount("PDF-unsupported print areas") == 0 &&
            FindFeatureCount("PDF-unsupported print titles") == 0 &&
            FindFeatureCount("PDF-unsupported header/footer formatting") == 0 &&
            FindFeatureCount("PDF-unrendered pivot tables") == 0 &&
            FindFeatureCount("PDF-unrendered sparklines") == 0 &&
            FindFeatureCount("Non-worksheet sheets") == 0;

        /// <summary>
        /// Returns true when the requested workflow-level capability can be attempted for this workbook.
        /// </summary>
        /// <param name="capability">The OfficeIMO.Excel workflow capability to check.</param>
        public bool Can(ExcelPreflightCapability capability) {
            switch (capability) {
                case ExcelPreflightCapability.ReadWorkbookData:
                    return CanReadWorkbookData;
                case ExcelPreflightCapability.EditCellValues:
                    return CanEditCellValues;
                case ExcelPreflightCapability.EditWorkbookStructure:
                    return CanEditWorkbookStructure;
                case ExcelPreflightCapability.UseCachedFormulaValues:
                    return CanUseCachedFormulaValues;
                case ExcelPreflightCapability.CalculateFormulas:
                    return CanCalculateFormulas;
                case ExcelPreflightCapability.BindTemplate:
                    return CanBindTemplate;
                case ExcelPreflightCapability.ExportPdfReport:
                    return CanExportPdfReport;
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported Excel preflight capability.");
            }
        }

        /// <summary>
        /// Throws with workflow-specific diagnostics when the requested capability cannot be attempted for this workbook.
        /// </summary>
        /// <param name="capability">The OfficeIMO.Excel workflow capability that must be available.</param>
        /// <returns>The current feature report for fluent guard usage.</returns>
        public ExcelFeatureReport EnsureCan(ExcelPreflightCapability capability) {
            if (Can(capability)) {
                return this;
            }

            IReadOnlyList<string> diagnostics = GetCapabilityDiagnostics(capability);
            string detail = diagnostics.Count == 0
                ? "No additional diagnostics were reported."
                : string.Join("; ", diagnostics);
            throw new InvalidOperationException($"Excel preflight capability '{capability}' is not available: {detail}");
        }

        /// <summary>
        /// Returns operation-specific diagnostics explaining why a workflow-level capability is blocked, or an empty list when it can be attempted.
        /// </summary>
        /// <param name="capability">The OfficeIMO.Excel workflow capability to explain.</param>
        public IReadOnlyList<string> GetCapabilityDiagnostics(ExcelPreflightCapability capability) {
            if (Can(capability)) {
                return Array.Empty<string>();
            }

            var messages = new List<string>();
            switch (capability) {
                case ExcelPreflightCapability.ReadWorkbookData:
                    AddUnsupportedDiagnostics(messages, "Workbook data reads are blocked by unsupported workbook features.");
                    AddFeatureDiagnostics(messages, FindFeatures("Non-worksheet sheets"));
                    break;
                case ExcelPreflightCapability.EditCellValues:
                    AddUnsupportedDiagnostics(messages, "Cell-value edits are blocked by unsupported workbook features.");
                    AddFeatureDiagnostics(messages, FindFeatures("Digital signatures"));
                    break;
                case ExcelPreflightCapability.EditWorkbookStructure:
                    AddUnsupportedDiagnostics(messages, "Structure-changing edits are blocked by unsupported workbook features.");
                    AddPreservedDiagnostics(messages, "Structure-changing edits should not be attempted while preserve-only workbook features are present.");
                    break;
                case ExcelPreflightCapability.UseCachedFormulaValues:
                    AddFormulaDiagnostics(messages, requireCachedValues: true, requireSupportedFormulas: false);
                    break;
                case ExcelPreflightCapability.CalculateFormulas:
                    AddFormulaDiagnostics(messages, requireCachedValues: false, requireSupportedFormulas: true);
                    break;
                case ExcelPreflightCapability.BindTemplate:
                    AddUnsupportedDiagnostics(messages, "Template binding is blocked by unsupported workbook features.");
                    AddPreservedDiagnostics(messages, "Template binding should not be attempted while preserve-only workbook features are present unless the workflow has explicit preservation coverage for those parts.");
                    break;
                case ExcelPreflightCapability.ExportPdfReport:
                    AddPdfExportWorkbookDiagnostics(messages);
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-missing formula caches"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-dirty formula caches"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-workbook recalculation requests"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported charts"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unreadable charts"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported images"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported hyperlinks"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unrendered drawing shapes"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported print areas"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported print titles"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unsupported header/footer formatting"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unrendered pivot tables"));
                    AddFeatureDiagnostics(messages, FindFeatures("PDF-unrendered sparklines"));
                    AddFeatureDiagnostics(messages, FindFeatures("Non-worksheet sheets"));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability, "Unsupported Excel preflight capability.");
            }

            if (messages.Count == 0) {
                AddDistinct(messages, "The requested Excel workflow is not available for this workbook.");
            }

            return messages.AsReadOnly();
        }

        private bool HasPdfExportWorkbookBlockers() {
            if (UnsupportedFeatures.Count > 0) {
                return true;
            }

            return PreservedFeatures.Any(IsPdfExportWorkbookBlocker);
        }

        private void AddPdfExportWorkbookDiagnostics(List<string> messages) {
            AddUnsupportedDiagnostics(messages, "Excel-to-PDF report export is blocked by unsupported workbook features.");
            var preservedBlockers = PreservedFeatures
                .Where(IsPdfExportWorkbookBlocker)
                .ToArray();
            if (preservedBlockers.Length == 0) {
                return;
            }

            AddDistinct(messages, "Excel-to-PDF report export does not render preserve-only workbook features.");
            AddFeatureDiagnostics(messages, preservedBlockers);
        }

        private static bool IsPdfExportWorkbookBlocker(ExcelFeatureFinding finding) {
            return !IsFormulaCachedValueFinding(finding);
        }

        private static bool IsFormulaCachedValueFinding(ExcelFeatureFinding finding) {
            if (!string.Equals(finding.Category, "Calculation", StringComparison.Ordinal)) {
                return false;
            }

            switch (finding.Name) {
                case "Unsupported formulas":
                case "Missing formula caches":
                case "Dirty formula caches":
                case "Workbook recalculation requests":
                case "PDF-missing formula caches":
                case "PDF-dirty formula caches":
                case "PDF-workbook recalculation requests":
                case "Formula dependency issues":
                case "Formula calculation blockers":
                    return true;
                default:
                    return false;
            }
        }

        private void AddUnsupportedDiagnostics(List<string> messages, string fallbackMessage) {
            if (UnsupportedFeatures.Count == 0) {
                return;
            }

            AddDistinct(messages, fallbackMessage);
            AddFeatureDiagnostics(messages, UnsupportedFeatures);
        }

        private void AddPreservedDiagnostics(List<string> messages, string fallbackMessage) {
            if (PreservedFeatures.Count == 0) {
                return;
            }

            AddDistinct(messages, fallbackMessage);
            AddFeatureDiagnostics(messages, PreservedFeatures);
        }

        private void AddFormulaDiagnostics(List<string> messages, bool requireCachedValues, bool requireSupportedFormulas) {
            if (requireSupportedFormulas) {
                AddFeatureDiagnostics(messages, FindFeatures("Formula calculation blockers"));
            }

            if (requireCachedValues) {
                AddFeatureDiagnostics(messages, FindFeatures("Missing formula caches"));
                AddFeatureDiagnostics(messages, FindFeatures("Dirty formula caches"));
                AddFeatureDiagnostics(messages, FindFeatures("Workbook recalculation requests"));
                return;
            }

            AddFormulaDependencyDiagnostics(messages, includeMissingCachedDependencyIssues: false);
        }

        private void AddFormulaDependencyDiagnostics(List<string> messages, bool includeMissingCachedDependencyIssues) {
            foreach (ExcelFeatureFinding finding in FindFeatures("Formula dependency issues")) {
                IReadOnlyList<string> details = FilterFormulaDependencyIssueDetails(finding.Details, includeMissingCachedDependencyIssues);
                if (details.Count == 0) {
                    continue;
                }

                AddDistinct(messages, FormatCapabilityFinding(new ExcelFeatureFinding(
                    finding.Category,
                    finding.Name,
                    finding.SupportLevel,
                    details.Count,
                    finding.Scope,
                    finding.Note,
                    details)));
            }
        }

        private int CountFormulaDependencyIssues(bool includeMissingCachedDependencyIssues) {
            int count = 0;
            foreach (ExcelFeatureFinding finding in FindFeatures("Formula dependency issues")) {
                count += includeMissingCachedDependencyIssues
                    ? finding.Count
                    : FilterFormulaDependencyIssueDetails(finding.Details, includeMissingCachedDependencyIssues).Count;
            }

            return count;
        }

        private static IReadOnlyList<string> FilterFormulaDependencyIssueDetails(IReadOnlyList<string> details, bool includeMissingCachedDependencyIssues) {
            if (includeMissingCachedDependencyIssues) {
                return details;
            }

            var filtered = new List<string>();
            foreach (string detail in details) {
                int separator = detail.IndexOf(": ", StringComparison.Ordinal);
                if (separator < 0) {
                    if (detail.IndexOf("without a cached result", StringComparison.OrdinalIgnoreCase) < 0) {
                        filtered.Add(detail);
                    }
                    continue;
                }

                string prefix = detail.Substring(0, separator);
                string[] issues = detail.Substring(separator + 2)
                    .Split(new[] { "; " }, StringSplitOptions.RemoveEmptyEntries);
                var calculationIssues = issues
                    .Where(issue => issue.IndexOf("without a cached result", StringComparison.OrdinalIgnoreCase) < 0)
                    .ToArray();
                if (calculationIssues.Length > 0) {
                    filtered.Add(prefix + ": " + string.Join("; ", calculationIssues));
                }
            }

            return filtered;
        }

        private static void AddFeatureDiagnostics(List<string> messages, IEnumerable<ExcelFeatureFinding> findings) {
            foreach (ExcelFeatureFinding finding in findings) {
                AddDistinct(messages, FormatCapabilityFinding(finding));
            }
        }

        private static string FormatCapabilityFinding(ExcelFeatureFinding finding) {
            string message = $"{finding.Name} ({finding.Count}, {finding.SupportLevel}): {finding.Note}";
            if (finding.Details.Count == 0) {
                return message;
            }

            const int maxDetails = 3;
            string details = string.Join("; ", finding.Details.Take(maxDetails));
            if (finding.Details.Count > maxDetails) {
                details += $"; +{finding.Details.Count - maxDetails} more";
            }

            return message + " [" + details + "]";
        }

        private int FindFeatureCount(string featureName) {
            return FindFeatures(featureName).Sum(feature => feature.Count);
        }

        private static void AddDistinct(List<string> messages, string message) {
            if (string.IsNullOrWhiteSpace(message)) {
                return;
            }

            for (int i = 0; i < messages.Count; i++) {
                if (string.Equals(messages[i], message, StringComparison.Ordinal)) {
                    return;
                }
            }

            messages.Add(message);
        }
    }
}
