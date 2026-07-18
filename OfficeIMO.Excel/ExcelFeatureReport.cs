using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Utilities;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// OfficeIMO's current author/edit/preserve status for a workbook feature discovered during inspection.
    /// </summary>
    public enum ExcelFeatureSupportLevel {
        /// <summary>OfficeIMO can author or edit this feature directly.</summary>
        Editable,

        /// <summary>OfficeIMO can author or edit common cases, but not the full Excel feature surface.</summary>
        PartiallyEditable,

        /// <summary>OfficeIMO should preserve the feature during round-trip saves, but does not expose a rich authoring API.</summary>
        Preserved,

        /// <summary>OfficeIMO has no meaningful support for the feature yet.</summary>
        Unsupported
    }

    /// <summary>
    /// Workbook-level feature and compatibility report.
    /// </summary>
    public sealed partial class ExcelFeatureReport {
        private readonly List<ExcelFeatureFinding> _features = new List<ExcelFeatureFinding>();

        internal ExcelFeatureReport(IReadOnlyList<ExcelFeatureFinding> features) {
            if (features == null) throw new ArgumentNullException(nameof(features));
            _features.AddRange(features);
        }

        /// <summary>
        /// Features discovered in the workbook.
        /// </summary>
        public IReadOnlyList<ExcelFeatureFinding> Features => _features;

        /// <summary>
        /// Features OfficeIMO can author or edit directly.
        /// </summary>
        public IReadOnlyList<ExcelFeatureFinding> EditableFeatures => _features
            .Where(feature => feature.SupportLevel == ExcelFeatureSupportLevel.Editable)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO can partly author or edit.
        /// </summary>
        public IReadOnlyList<ExcelFeatureFinding> PartiallyEditableFeatures => _features
            .Where(feature => feature.SupportLevel == ExcelFeatureSupportLevel.PartiallyEditable)
            .ToArray();

        /// <summary>
        /// Advanced features OfficeIMO should preserve but cannot fully author or edit yet.
        /// </summary>
        public IReadOnlyList<ExcelFeatureFinding> PreservedFeatures => _features
            .Where(feature => feature.SupportLevel == ExcelFeatureSupportLevel.Preserved)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO does not meaningfully support yet.
        /// </summary>
        public IReadOnlyList<ExcelFeatureFinding> UnsupportedFeatures => _features
            .Where(feature => feature.SupportLevel == ExcelFeatureSupportLevel.Unsupported)
            .ToArray();

        /// <summary>
        /// Whether the workbook contains advanced features that should be checked before edit-heavy round trips.
        /// </summary>
        public bool HasAdvancedFeatures => PreservedFeatures.Count > 0 || UnsupportedFeatures.Count > 0;

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>External workbook links</c>.</param>
        public IReadOnlyList<ExcelFeatureFinding> FindFeatures(params string[] featureNames) {
            return FindFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>External workbook links</c>.</param>
        public IReadOnlyList<ExcelFeatureFinding> FindFeatures(IEnumerable<string> featureNames) {
            if (featureNames == null) throw new ArgumentNullException(nameof(featureNames));
            var names = new HashSet<string>(featureNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (names.Count == 0) return Array.Empty<ExcelFeatureFinding>();

            return _features
                .Where(feature => names.Contains(feature.Name))
                .ToArray();
        }

        /// <summary>
        /// Returns discovered features with one of the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to match.</param>
        public IReadOnlyList<ExcelFeatureFinding> FindFeatures(params ExcelFeatureSupportLevel[] supportLevels) {
            if (supportLevels == null) throw new ArgumentNullException(nameof(supportLevels));
            if (supportLevels.Length == 0) return Array.Empty<ExcelFeatureFinding>();

            var levels = new HashSet<ExcelFeatureSupportLevel>(supportLevels);
            return _features
                .Where(feature => levels.Contains(feature.SupportLevel))
                .ToArray();
        }

        /// <summary>
        /// Throws when the workbook contains unsupported features.
        /// </summary>
        public ExcelFeatureReport EnsureNoUnsupportedFeatures() {
            if (UnsupportedFeatures.Count > 0) {
                ThrowBlockedFeatures("Unsupported workbook features", UnsupportedFeatures);
            }

            return this;
        }

        /// <summary>
        /// Throws when the workbook contains preserve-only or unsupported advanced features.
        /// </summary>
        public ExcelFeatureReport EnsureNoAdvancedFeatures() {
            if (HasAdvancedFeatures) {
                ThrowBlockedFeatures("Advanced workbook features need review before edit-heavy round trips", PreservedFeatures.Concat(UnsupportedFeatures));
            }

            return this;
        }

        /// <summary>
        /// Throws when the workbook contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>External workbook links</c>.</param>
        public ExcelFeatureReport EnsureNoFeatures(params string[] featureNames) {
            return EnsureNoFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Throws when the workbook contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>External workbook links</c>.</param>
        public ExcelFeatureReport EnsureNoFeatures(IEnumerable<string> featureNames) {
            var matches = FindFeatures(featureNames);
            if (matches.Count > 0) {
                ThrowBlockedFeatures("Workbook contains blocked features", matches);
            }

            return this;
        }

        /// <summary>
        /// Throws when the workbook contains any features with the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to reject.</param>
        public ExcelFeatureReport EnsureNoFeatures(params ExcelFeatureSupportLevel[] supportLevels) {
            var matches = FindFeatures(supportLevels);
            if (matches.Count > 0) {
                ThrowBlockedFeatures("Workbook contains blocked feature support levels", matches);
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of discovered workbook features and support status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Excel Feature Report");
            builder.AppendLine();
            builder.AppendLine($"Total findings: {Features.Count}");
            builder.AppendLine($"Editable features: {EditableFeatures.Count}");
            builder.AppendLine($"Partially editable features: {PartiallyEditableFeatures.Count}");
            builder.AppendLine($"Preserved features: {PreservedFeatures.Count}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatures.Count}");
            builder.AppendLine();
            builder.AppendLine("## Capability Preflight");
            builder.AppendLine();
            builder.AppendLine("| Capability | Can attempt | Diagnostics |");
            builder.AppendLine("| --- | --- | --- |");
            foreach (ExcelPreflightCapability capability in Enum.GetValues(typeof(ExcelPreflightCapability))) {
                builder.Append("| ");
                builder.Append(capability);
                builder.Append(" | ");
                builder.Append(Can(capability) ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", GetCapabilityDiagnostics(capability))));
                builder.AppendLine(" |");
            }

            builder.AppendLine();
            builder.AppendLine("## Repair Hints");
            builder.AppendLine();
            builder.AppendLine("| Capability | Feature | Action | Command | Details |");
            builder.AppendLine("| --- | --- | --- | --- | --- |");
            foreach (ExcelPreflightCapability capability in Enum.GetValues(typeof(ExcelPreflightCapability))) {
                foreach (ExcelPreflightRepairHint hint in GetRepairHints(capability)) {
                    builder.Append("| ");
                    builder.Append(EscapeMarkdownCell(capability.ToString()));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(hint.FeatureName));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(hint.Action));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(hint.Command ?? string.Empty));
                    builder.Append(" | ");
                    builder.Append(EscapeMarkdownCell(hint.Details ?? string.Empty));
                    builder.AppendLine(" |");
                }
            }

            builder.AppendLine();
            builder.AppendLine("| Category | Feature | Count | Support | Scope | Note | Details |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (ExcelFeatureFinding feature in Features) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(feature.Category));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(feature.Name));
                builder.Append(" | ");
                builder.Append(feature.Count);
                builder.Append(" | ");
                builder.Append(feature.SupportLevel);
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(feature.Scope ?? string.Empty));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(feature.Note));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(FormatDetails(feature.Details)));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string FormatDetails(IReadOnlyList<string> details) {
            if (details.Count == 0) return string.Empty;
            const int maxDetails = 8;
            if (details.Count <= maxDetails) {
                return string.Join("; ", details);
            }

            return string.Join("; ", details.Take(maxDetails)) + $"; +{details.Count - maxDetails} more";
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }

        private static void ThrowBlockedFeatures(string message, IEnumerable<ExcelFeatureFinding> findings) {
            var formatted = findings
                .OrderBy(feature => feature.Name, StringComparer.OrdinalIgnoreCase)
                .Select(FormatBlockedFeature)
                .ToArray();
            throw new InvalidOperationException(message + ": " + string.Join(", ", formatted));
        }

        private static string FormatBlockedFeature(ExcelFeatureFinding feature) {
            string summary = $"{feature.Name} ({feature.Count}, {feature.SupportLevel})";
            if (feature.Details.Count == 0) {
                return summary;
            }

            const int maxDetails = 3;
            string details = string.Join("; ", feature.Details.Take(maxDetails));
            if (feature.Details.Count > maxDetails) {
                details += $"; +{feature.Details.Count - maxDetails} more";
            }

            return summary + " [" + details + "]";
        }
    }

    /// <summary>
    /// One feature discovered in a workbook.
    /// </summary>
    public sealed class ExcelFeatureFinding {
        internal ExcelFeatureFinding(string category, string name, ExcelFeatureSupportLevel supportLevel, int count, string? scope, string note,
            IReadOnlyList<string>? details = null) {
            Category = string.IsNullOrWhiteSpace(category) ? throw new ArgumentNullException(nameof(category)) : category;
            Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentNullException(nameof(name)) : name;
            SupportLevel = supportLevel;
            Count = count;
            Scope = string.IsNullOrWhiteSpace(scope) ? null : scope;
            Note = string.IsNullOrWhiteSpace(note) ? string.Empty : note;
            Details = details ?? Array.Empty<string>();
        }

        /// <summary>
        /// Broad feature area, for example calculation, visualization, or compatibility.
        /// </summary>
        public string Category { get; }

        /// <summary>
        /// Feature name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// OfficeIMO support status for this feature.
        /// </summary>
        public ExcelFeatureSupportLevel SupportLevel { get; }

        /// <summary>
        /// Number of matching items discovered.
        /// </summary>
        public int Count { get; }

        /// <summary>
        /// Optional workbook or worksheet scope.
        /// </summary>
        public string? Scope { get; }

        /// <summary>
        /// Short explanation of what OfficeIMO can do with this feature today.
        /// </summary>
        public string Note { get; }

        /// <summary>
        /// Optional package, relationship, or worksheet details that explain where the feature was found.
        /// </summary>
        public IReadOnlyList<string> Details { get; }
    }

    public partial class ExcelDocument {
        /// <summary>
        /// Inspects workbook features and reports which ones OfficeIMO can edit, partially edit, preserve, or does not support yet.
        /// </summary>
        public ExcelFeatureReport InspectFeatures() {
            var features = new List<ExcelFeatureFinding>();
            WorkbookPart workbookPart = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            var sheetElements = workbook.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList()
                ?? new List<DocumentFormat.OpenXml.Spreadsheet.Sheet>();

            Add(features, "Workbook", "Worksheets", ExcelFeatureSupportLevel.Editable, sheetElements.Count, null,
                "Worksheets can be created, renamed, moved, hidden, inspected, copied, and removed.");

            int nonWorksheetSheetCount = 0;
            var nonWorksheetSheetDetails = new List<string>();
            int namedRangeCount = workbook.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Count() ?? 0;
            Add(features, "Workbook", "Named ranges", ExcelFeatureSupportLevel.Editable, namedRangeCount, null,
                "Workbook and sheet-local named ranges are editable.");

            int tableCount = 0;
            int chartCount = 0;
            int pivotCount = 0;
            int dataValidationCount = 0;
            int conditionalFormattingCount = 0;
            int sparklineCount = 0;
            int pdfUnrenderedPivotCount = 0;
            int pdfUnrenderedSparklineCount = 0;
            int legacyCommentCount = 0;
            int threadedCommentPartCount = 0;
            int imagePartCount = 0;
            int oleObjectCount = 0;
            int formControlCount = 0;
            int externalHyperlinkCount = 0;
            int pdfUnsupportedChartCount = 0;
            int pdfUnreadableChartCount = 0;
            int pdfUnsupportedImageCount = 0;
            int pdfUnsupportedHyperlinkCount = 0;
            int pdfUnrenderedDrawingShapeCount = 0;
            int pdfUnsupportedPrintAreaCount = 0;
            int pdfUnsupportedPrintTitleCount = 0;
            int pdfUnsupportedHeaderFooterFormattingCount = 0;
            var threadedCommentDetails = new List<string>();
            var oleObjectDetails = new List<string>();
            var formControlDetails = new List<string>();
            var externalHyperlinkDetails = new List<string>();
            var pdfUnsupportedChartDetails = new List<string>();
            var pdfUnreadableChartDetails = new List<string>();
            var pdfUnsupportedImageDetails = new List<string>();
            var pdfUnsupportedHyperlinkDetails = new List<string>();
            var pdfUnsupportedPrintAreaDetails = new List<string>();
            var pdfUnsupportedPrintTitleDetails = new List<string>();
            var pdfUnsupportedHeaderFooterFormattingDetails = new List<string>();
            var pivotDetails = new List<string>();
            var pdfUnrenderedPivotDetails = new List<string>();
            var sparklineDetails = new List<string>();
            var pdfUnrenderedSparklineDetails = new List<string>();
            var pdfUnrenderedDrawingShapeDetails = new List<string>();
            var threadedCommentPeople = ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);
            var defaultPdfExportSheetsByName = new Dictionary<string, ExcelSheet>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheet in sheetElements) {
                if (string.IsNullOrWhiteSpace(sheet.Id?.Value)) {
                    continue;
                }

                OpenXmlPart sheetPart = workbookPart.GetPartById(sheet.Id!.Value!);
                if (sheetPart is not WorksheetPart worksheetPart) {
                    nonWorksheetSheetCount++;
                    nonWorksheetSheetDetails.Add($"{sheet.Name}: {sheetPart.GetType().Name} ({sheetPart.Uri})");
                    continue;
                }

                var excelSheet = new ExcelSheet(this, _spreadSheetDocument!, sheet);
                var worksheet = worksheetPart.Worksheet;
                bool isVisibleForDefaultPdfExport = !IsHiddenSheet(sheet);
                string sheetName = sheet.Name?.Value ?? string.Empty;
                if (isVisibleForDefaultPdfExport && !string.IsNullOrWhiteSpace(sheetName)) {
                    defaultPdfExportSheetsByName[sheetName] = excelSheet;
                }
                tableCount += worksheetPart.TableDefinitionParts.Count();
                int sheetPivotCount = worksheetPart.PivotTableParts.Count();
                pivotCount += sheetPivotCount;
                dataValidationCount += worksheet?.Descendants<DocumentFormat.OpenXml.Spreadsheet.DataValidation>().Count() ?? 0;
                conditionalFormattingCount += worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>().Count() ?? 0;
                IReadOnlyList<ExcelWorksheetSparklineInfo> sheetSparklines = ExcelWorksheetSparklineResolver.FindSparklines(worksheetPart);
                int sheetSparklineCount = sheetSparklines.Count;
                sparklineCount += sheetSparklineCount;
                if (sheetPivotCount > 0) pivotDetails.Add($"{sheet.Name}: {sheetPivotCount} pivot table(s)");
                if (sheetSparklineCount > 0) {
                    sparklineDetails.AddRange(sheetSparklines.Select(sparkline => $"{sheet.Name}!{sparkline.CellReference}: sparkline"));
                }
                if (isVisibleForDefaultPdfExport) {
                    pdfUnrenderedPivotCount += sheetPivotCount;
                    pdfUnrenderedSparklineCount += sheetSparklineCount;
                    if (sheetPivotCount > 0) pdfUnrenderedPivotDetails.Add($"{sheet.Name}: {sheetPivotCount} pivot table(s)");
                    if (sheetSparklineCount > 0) {
                        pdfUnrenderedSparklineDetails.AddRange(sheetSparklines.Select(sparkline => $"{sheet.Name}!{sparkline.CellReference}: sparkline"));
                    }
                }
                legacyCommentCount += worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<DocumentFormat.OpenXml.Spreadsheet.Comment>().Count() ?? 0;
                var threadedComments = ExcelWorksheetCommentResolver.BuildThreadedCommentMap(worksheetPart, threadedCommentPeople, sheetName)
                    .Values
                    .SelectMany(comments => comments)
                    .ToList();
                threadedCommentPartCount += threadedComments.Count;
                foreach (var threadedComment in threadedComments) {
                    string author = string.IsNullOrWhiteSpace(threadedComment.Author) ? threadedComment.PersonId ?? "unknown author" : threadedComment.Author!;
                    threadedCommentDetails.Add($"{sheet.Name}: {threadedComment.CellReference} by {author}");
                }

                var images = excelSheet.Images.ToList();
                imagePartCount += images.Count;
                foreach (ExcelImage image in images) {
                    if (isVisibleForDefaultPdfExport && !IsPdfSupportedWorksheetImage(image, out string reason)) {
                        pdfUnsupportedImageCount++;
                        pdfUnsupportedImageDetails.Add($"{sheet.Name}!{A1.CellReference(image.RowIndex, image.ColumnIndex)}: {reason}");
                    }
                }

                var charts = excelSheet.Charts.ToList();
                chartCount += charts.Count;
                foreach (ExcelChart chart in charts.Where(_ => isVisibleForDefaultPdfExport)) {
                    if (!chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                        pdfUnreadableChartCount++;
                        pdfUnreadableChartDetails.Add($"{sheet.Name}: {GetSafeChartDisplayName(chart)} data could not be read into a PDF snapshot.");
                        continue;
                    }

                    if (!HasRenderablePdfChartData(snapshot)) {
                        pdfUnreadableChartCount++;
                        pdfUnreadableChartDetails.Add($"{sheet.Name}: {GetChartDisplayName(snapshot)} does not contain renderable chart categories and series.");
                    } else if (HasMixedPdfChartTypes(snapshot)) {
                        pdfUnsupportedChartCount++;
                        pdfUnsupportedChartDetails.Add($"{sheet.Name}: mixed per-series chart types ({GetChartDisplayName(snapshot)})");
                    } else if (!IsPdfSupportedChartType(snapshot.ChartType)) {
                        pdfUnsupportedChartCount++;
                        pdfUnsupportedChartDetails.Add($"{sheet.Name}: {snapshot.ChartType} ({GetChartDisplayName(snapshot)})");
                    }
                }
                if (isVisibleForDefaultPdfExport) {
                    ExcelSheet.HeaderFooterSnapshot headerFooter = excelSheet.GetHeaderFooter();
                    AddUnsupportedHeaderFooterImages(headerFooter, sheetName, ref pdfUnsupportedImageCount, pdfUnsupportedImageDetails);
                    AddUnsupportedHeaderFooterFormatting(headerFooter, sheetName, ref pdfUnsupportedHeaderFooterFormattingCount, pdfUnsupportedHeaderFooterFormattingDetails);
                    AddUnsupportedPrintArea(excelSheet, sheetName, ref pdfUnsupportedPrintAreaCount, pdfUnsupportedPrintAreaDetails);
                    AddUnsupportedPrintTitles(excelSheet, sheetName, ref pdfUnsupportedPrintTitleCount, pdfUnsupportedPrintTitleDetails);
                    AddUnrenderedDrawingShapes(worksheetPart, sheetName, ref pdfUnrenderedDrawingShapeCount, pdfUnrenderedDrawingShapeDetails);
                    AddUnsupportedWorksheetHyperlinks(workbookPart, sheetElements, worksheetPart, sheetName, ref pdfUnsupportedHyperlinkCount, pdfUnsupportedHyperlinkDetails);
                    AddUnsupportedDrawingHyperlinks(worksheetPart, sheetName, ref pdfUnsupportedHyperlinkCount, pdfUnsupportedHyperlinkDetails);
                }
                int sheetOleObjects = CountDescendantsByLocalName(worksheet, "oleObject");
                int sheetFormControls = CountDescendantsByLocalName(worksheet, "control") + CountDescendantsByLocalName(worksheet, "formControl");
                oleObjectCount += sheetOleObjects;
                formControlCount += sheetFormControls;
                externalHyperlinkCount += worksheetPart.HyperlinkRelationships.Count();
                if (sheetOleObjects > 0) oleObjectDetails.Add($"{sheet.Name}: {sheetOleObjects} OLE object(s)");
                if (sheetFormControls > 0) formControlDetails.Add($"{sheet.Name}: {sheetFormControls} form control marker(s)");
                foreach (var relationship in worksheetPart.HyperlinkRelationships) {
                    externalHyperlinkDetails.Add($"{sheet.Name}: {relationship.Id} -> {relationship.Uri}");
                    if (isVisibleForDefaultPdfExport && !IsSupportedPdfExternalHyperlink(relationship.Uri)) {
                        pdfUnsupportedHyperlinkCount++;
                        pdfUnsupportedHyperlinkDetails.Add($"{sheet.Name}: {relationship.Id} -> {relationship.Uri} is not an absolute URI supported by the first-party PDF hyperlink writer.");
                    }
                }
            }

            Add(features, "Data", "Tables", ExcelFeatureSupportLevel.Editable, tableCount, null,
                "Tables can be authored and inspected, including AutoFilter metadata.");
            Add(features, "Data", "Data validations", ExcelFeatureSupportLevel.Editable, dataValidationCount, null,
                "List, numeric, date, time, text-length, custom formula, prompt, and error metadata are editable.");
            Add(features, "Formatting", "Conditional formatting", ExcelFeatureSupportLevel.PartiallyEditable, conditionalFormattingCount, null,
                "Common rule types are editable; full Excel conditional-formatting parity remains a roadmap item.");
            Add(features, "Visualization", "Charts", ExcelFeatureSupportLevel.PartiallyEditable, chartCount, null,
                "Common chart authoring and updates are supported, including stacked/100% stacked column/bar/line/area variants, 3-D area/line/column/bar/pie, pie-of-pie/bar-of-pie, radar, stock, and filled/wireframe/contour surface charts; advanced chart families remain partial.");
            Add(features, "Visualization", "PDF-unsupported charts", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedChartCount, null,
                "These charts can be authored or preserved in the workbook but are skipped by the first-party Excel-to-PDF chart snapshot renderer.",
                pdfUnsupportedChartDetails);
            Add(features, "Visualization", "PDF-unreadable charts", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnreadableChartCount, null,
                "These charts are present in the workbook but cannot be read into the first-party Excel-to-PDF chart snapshot model.",
                pdfUnreadableChartDetails);
            Add(features, "Visualization", "Pivot tables", ExcelFeatureSupportLevel.PartiallyEditable, pivotCount, null,
                "Source-range pivot creation and inspection are supported, including composable fluent field sort/subtotal/layout/display/number-format helpers with built-in/custom id/code readback, field item/page filters with fluent helpers plus hidden, visible, and selected-item readback, common label/value filters, negated filter variants, fixed and dynamic date filters, top/bottom count/percent/sum filters, formula-backed calculated fields with number-format id/code readback, date/number grouping metadata, generated multi-level date hierarchy fields with base/parent relationships, and explicit grouped-cache item metadata; slicers, deeper Excel interoperability checks, and advanced filters remain partial.");
            Add(features, "Visualization", "PDF-unrendered pivot tables", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnrenderedPivotCount, null,
                "Pivot table metadata is not rendered by the first-party Excel-to-PDF path unless the pivot output is already materialized as ordinary worksheet cells.",
                pdfUnrenderedPivotDetails);
            Add(features, "Visualization", "Sparklines", ExcelFeatureSupportLevel.Editable, sparklineCount, null,
                "Line, column, and win/loss sparklines can be authored.");
            Add(features, "Visualization", "PDF-unrendered sparklines", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnrenderedSparklineCount, null,
                "Sparkline metadata is authored and preserved in worksheets, but the first-party Excel-to-PDF path does not render sparkline visuals.",
                pdfUnrenderedSparklineDetails);
            Add(features, "Collaboration", "Legacy comments", ExcelFeatureSupportLevel.PartiallyEditable, legacyCommentCount, null,
                "Legacy comments can be authored and inspected, including rich-text runs for authored comments.");
            Add(features, "Collaboration", "Threaded comments", ExcelFeatureSupportLevel.PartiallyEditable, threadedCommentPartCount, null,
                "Plain-text threaded comments and replies can be authored, inspected, updated, resolved, reopened, and removed while maintaining workbook person metadata; mentions and richer Microsoft 365 identity metadata remain preserve-only.",
                threadedCommentDetails);
            Add(features, "Media", "Images", ExcelFeatureSupportLevel.PartiallyEditable, imagePartCount, null,
                "Images can be inserted in common worksheet/header/footer scenarios; advanced drawing behaviors remain partial.");
            Add(features, "Media", "PDF-unsupported images", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedImageCount, null,
                "Worksheet images are present but are skipped by the first-party Excel-to-PDF image writer because only valid PNG and JPEG images are rendered.",
                pdfUnsupportedImageDetails);
            Add(features, "Media", "PDF-unrendered drawing shapes", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnrenderedDrawingShapeCount, null,
                "Worksheet drawing shapes and text boxes are present but are skipped by the first-party Excel-to-PDF path.",
                pdfUnrenderedDrawingShapeDetails);
            Add(features, "Layout", "PDF-unsupported print areas", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedPrintAreaCount, null,
                "Worksheet print-area settings are present but the first-party Excel-to-PDF path falls back to the worksheet used range for multi-area print areas.",
                pdfUnsupportedPrintAreaDetails);
            Add(features, "Layout", "PDF-unsupported print titles", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedPrintTitleCount, null,
                "Worksheet print-title columns are configured but the first-party Excel-to-PDF path currently repeats print-title rows only.",
                pdfUnsupportedPrintTitleDetails);
            Add(features, "Layout", "PDF-unsupported header/footer formatting", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedHeaderFooterFormattingCount, null,
                "Header or footer text uses formatting that is simplified by the first-party Excel-to-PDF path.",
                pdfUnsupportedHeaderFooterFormattingDetails);
            Add(features, "Compatibility", "OLE objects", ExcelFeatureSupportLevel.PartiallyEditable, oleObjectCount, null,
                "Embedded OLE payloads can be inventoried, hash-checked, extracted with byte limits, replaced, and removed; creating complete worksheet OLE presentation markup remains unsupported.", oleObjectDetails);
            Add(features, "Compatibility", "Form controls", ExcelFeatureSupportLevel.Preserved, formControlCount, null,
                "Form controls are preserve-only worksheet metadata.", formControlDetails);
            Add(features, "Compatibility", "External hyperlinks", ExcelFeatureSupportLevel.PartiallyEditable, externalHyperlinkCount, null,
                "Worksheet external hyperlinks can be authored and are rendered by PDF export when they target absolute URIs.",
                externalHyperlinkDetails);
            Add(features, "Compatibility", "PDF-unsupported hyperlinks", ExcelFeatureSupportLevel.PartiallyEditable, pdfUnsupportedHyperlinkCount, null,
                "These external hyperlinks are present but are skipped by the first-party Excel-to-PDF hyperlink writer.",
                pdfUnsupportedHyperlinkDetails);
            Add(features, "Compatibility", "Non-worksheet sheets", ExcelFeatureSupportLevel.Preserved, nonWorksheetSheetCount, null,
                "Chartsheets and other non-worksheet sheet parts are preserve-only and cannot be materialized by worksheet-only workflows.",
                nonWorksheetSheetDetails);

            var formulas = InspectFormulas();
            var pdfExportFormulas = formulas.Formulas
                .Where(formula => IsDefaultPdfExportedFormulaCell(formula, defaultPdfExportSheetsByName))
                .ToArray();
            bool hasWorkbookRecalculationRequest = formulas.Formulas.Count > 0 && HasWorkbookRecalculationRequest(workbook);
            bool hasPdfWorkbookRecalculationRequest = pdfExportFormulas.Length > 0 && HasWorkbookRecalculationRequest(workbook);
            Add(features, "Calculation", "Supported formulas", ExcelFeatureSupportLevel.PartiallyEditable, formulas.SupportedFormulas, null,
                "Simple supported formulas can be recalculated by OfficeIMO.");
            Add(features, "Calculation", "Unsupported formulas", ExcelFeatureSupportLevel.Preserved, formulas.UnsupportedFormulas, null,
                "Unsupported formulas are preserved and should be recalculated by Excel or read from cached values.");
            Add(features, "Calculation", "Missing formula caches", ExcelFeatureSupportLevel.Preserved, formulas.MissingCachedResults, null,
                "Formulas without cached results need OfficeIMO calculation support or Excel recalculation before cached-value reads are reliable.");
            Add(features, "Calculation", "Dirty formula caches", ExcelFeatureSupportLevel.PartiallyEditable, formulas.DirtyFormulas, null,
                "Dirty formulas have cached results that are explicitly awaiting recalculation before cached-value reads are reliable.",
                formulas.Formulas
                    .Where(formula => formula.IsDirty)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray());
            Add(features, "Calculation", "Workbook recalculation requests", ExcelFeatureSupportLevel.Preserved, hasWorkbookRecalculationRequest ? 1 : 0, null,
                "The workbook requests full recalculation on open, so cached formula values should be refreshed before cached-value reads are trusted.",
                hasWorkbookRecalculationRequest ? DescribeWorkbookRecalculationRequest(workbook).ToArray() : Array.Empty<string>());
            Add(features, "Calculation", "PDF-missing formula caches", ExcelFeatureSupportLevel.Preserved, pdfExportFormulas.Count(formula => !formula.HasCachedValue), null,
                "Visible worksheet formulas without cached results need OfficeIMO calculation support or Excel recalculation before PDF export can trust cached values.",
                pdfExportFormulas
                    .Where(formula => !formula.HasCachedValue)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray());
            Add(features, "Calculation", "PDF-dirty formula caches", ExcelFeatureSupportLevel.PartiallyEditable, pdfExportFormulas.Count(formula => formula.IsDirty), null,
                "Visible worksheet formulas with dirty cached results need recalculation before PDF export can trust cached values.",
                pdfExportFormulas
                    .Where(formula => formula.IsDirty)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray());
            Add(features, "Calculation", "PDF-workbook recalculation requests", ExcelFeatureSupportLevel.Preserved, hasPdfWorkbookRecalculationRequest ? 1 : 0, null,
                "The workbook requests full recalculation on open while visible worksheet formulas are exported to PDF, so cached values should be refreshed first.",
                hasPdfWorkbookRecalculationRequest ? DescribeWorkbookRecalculationRequest(workbook).ToArray() : Array.Empty<string>());
            var formulaCalculationBlockers = formulas.Formulas
                .SelectMany(GetFormulaCalculationBlockers)
                .ToArray();
            Add(features, "Calculation", "Formula calculation blockers", ExcelFeatureSupportLevel.Preserved, formulaCalculationBlockers.Length, null,
                "These formula cells need Excel recalculation or broader evaluator support before OfficeIMO can calculate the workbook.",
                formulaCalculationBlockers);
            string[] formulaDependencyDetails = formulas.Formulas
                .Where(formula => formula.DependencyIssues.Count > 0)
                .Select(formula => $"{formula.SheetName}!{formula.CellReference}: {string.Join("; ", formula.DependencyIssues)}")
                .Concat(formulas.DependencyGraph.CircularReferences.Select(circular =>
                    $"Circular reference: {string.Join(" -> ", circular.References)}"))
                .ToArray();
            Add(features, "Calculation", "Formula dependency issues", ExcelFeatureSupportLevel.Preserved, formulas.DependencyIssueCount, null,
                "Formula dependencies need review before OfficeIMO calculation can be trusted; clean cached values remain usable when every formula cache is present and current.",
                formulaDependencyDetails);

            var allParts = EnumeratePackageParts(_spreadSheetDocument).ToList();
            var vbaDetails = DescribePartsByUriOrContentType(allParts, "vbaProject");
            var slicerDetails = DescribeParts(allParts, IsNativeSlicerPackagePart);
            var timelineDetails = DescribeParts(allParts, IsNativeTimelinePackagePart);
            var slicerBindingMetadataDetails = DescribeParts(allParts, IsSlicerBindingMetadataPart);
            var timelineBindingMetadataDetails = DescribeParts(allParts, IsTimelineBindingMetadataPart);
            var externalLinkDetails = DescribePartsByUriOrContentType(allParts, "externalLink");
            var connectionDetails = DescribePartsByUriOrContentType(allParts, "connection")
                .Concat(DescribePartsByUriOrContentType(allParts, "queryTable"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var customXmlDetails = DescribePartsByUri(allParts, "/customXml/");
            var embeddedPackageDetails = DescribePartsByUri(allParts, "/embeddings/");
            var signatureDetails = DescribePartsByUriOrContentType(allParts, "signature")
                .Concat(DescribePartsByUriOrContentType(allParts, "xmlsignatures"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (_spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                signatureDetails.Add("Extended application properties contain digital signature metadata.");
            }
            var externalRelationshipDetails = DescribeExternalRelationships(allParts, includeHyperlinks: false)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            Add(features, "Compatibility", "VBA macros", ExcelFeatureSupportLevel.PartiallyEditable, vbaDetails.Count, null,
                "VBA projects can be attached, hash-checked, inspected, extracted with byte limits, and removed; OfficeIMO does not edit VBA source or sign macro projects.", vbaDetails);
            Add(features, "Compatibility", "Slicers", ExcelFeatureSupportLevel.PartiallyEditable, slicerDetails.Count, null,
                "Native Excel slicer parts can be inspected and preserved; native cache and UI authoring remains partial.", slicerDetails);
            Add(features, "Compatibility", "Timelines", ExcelFeatureSupportLevel.PartiallyEditable, timelineDetails.Count, null,
                "Native Excel timeline parts can be inspected and preserved; native cache and UI authoring remains partial.", timelineDetails);
            Add(features, "Compatibility", "Slicer binding metadata", ExcelFeatureSupportLevel.Editable, slicerBindingMetadataDetails.Count, null,
                "OfficeIMO-owned pivot slicer binding metadata can be authored and read back, but it is not a native Excel slicer cache or UI object.", slicerBindingMetadataDetails);
            Add(features, "Compatibility", "Timeline binding metadata", ExcelFeatureSupportLevel.Editable, timelineBindingMetadataDetails.Count, null,
                "OfficeIMO-owned pivot timeline binding metadata can be authored and read back, but it is not a native Excel timeline cache or UI object.", timelineBindingMetadataDetails);
            Add(features, "Compatibility", "External workbook links", ExcelFeatureSupportLevel.Preserved, externalLinkDetails.Count + externalRelationshipDetails.Count, null,
                "External relationships and workbook-link parts should be treated carefully during round trips.",
                externalLinkDetails.Concat(externalRelationshipDetails).ToArray());
            Add(features, "Compatibility", "Connections and query tables", ExcelFeatureSupportLevel.Preserved, connectionDetails.Count, null,
                "Connections and query-table metadata are preserve-only.", connectionDetails);
            Add(features, "Compatibility", "Custom XML parts", ExcelFeatureSupportLevel.Preserved, customXmlDetails.Count, null,
                "Custom XML parts are preserve-only package metadata.", customXmlDetails);
            Add(features, "Compatibility", "Digital signatures", ExcelFeatureSupportLevel.Preserved, signatureDetails.Count, null,
                "Digital signature parts are preserve-only package metadata.", signatureDetails);
            Add(features, "Compatibility", "Embedded packages", ExcelFeatureSupportLevel.PartiallyEditable, embeddedPackageDetails.Count, null,
                "Embedded package payloads can be inventoried, hash-checked, extracted with byte limits, replaced, and removed; complete visual OLE authoring remains unsupported.", embeddedPackageDetails);

            AddLegacyXlsImportFeatures(features);

            return new ExcelFeatureReport(features);
        }

        private void AddLegacyXlsImportFeatures(List<ExcelFeatureFinding> features) {
            if (SourceFormat != ExcelFileFormat.Xls) {
                return;
            }

            Add(features, "Compatibility", "Legacy XLS source", ExcelFeatureSupportLevel.Editable, 1, null,
                "Legacy binary XLS content was projected into the normal OfficeIMO Excel model. Saving writes Open XML .xlsx content; native .xls writing is not supported.");

            LegacyXlsImportDiagnostic[] warnings = _legacyXlsImportDiagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning)
                .ToArray();
            Add(features, "Compatibility", "Legacy XLS import diagnostics", ExcelFeatureSupportLevel.Preserved, warnings.Length, null,
                "Legacy XLS import warnings should be reviewed before relying on full fidelity.",
                warnings.Select(FormatLegacyXlsDiagnostic).ToArray());

            Add(features, "Compatibility", "Legacy XLS unsupported features", ExcelFeatureSupportLevel.Unsupported, _legacyXlsUnsupportedFeatures.Length, null,
                "Some legacy XLS features were detected as unsupported import metadata and are not written to the converted .xlsx package.",
                _legacyXlsUnsupportedFeatures.Select(FormatLegacyXlsUnsupportedFeature).ToArray());

            Add(features, "Compatibility", "Legacy XLS preserved records", ExcelFeatureSupportLevel.Preserved, _legacyXlsPreservedFeatures.Length, null,
                "Preserve-only BIFF records were detected as import metadata but are not projected into the workbook model or written to converted output.",
                _legacyXlsPreservedFeatures.Select(FormatLegacyXlsPreservedFeature).ToArray());

            Add(features, "Compatibility", "Legacy XLS unsupported sheets", ExcelFeatureSupportLevel.Unsupported, _legacyXlsUnsupportedSheets.Length, null,
                "Some legacy XLS sheet entries were discovered but not projected as normal worksheets and are not written to the converted .xlsx package.",
                _legacyXlsUnsupportedSheets.Select(FormatLegacyXlsUnsupportedSheet).ToArray());

            Add(features, "Compatibility", "Legacy XLS chart sheets", ExcelFeatureSupportLevel.Preserved, _legacyXlsChartSheets.Length, null,
                "Legacy XLS chart sheets are decoded as import metadata, but are not projected as normal worksheets or written by native XLS save.",
                _legacyXlsChartSheets.Select(FormatLegacyXlsChartSheet).ToArray());

            Add(features, "Compatibility", "Legacy XLS compound features", ExcelFeatureSupportLevel.Preserved, _legacyXlsCompoundFeatures.Length, null,
                "Legacy XLS compound-container features are decoded as import metadata, but are not projected into the normal workbook package or written by native XLS save.",
                _legacyXlsCompoundFeatures.Select(FormatLegacyXlsCompoundFeature).ToArray());
        }

        private static string FormatLegacyXlsDiagnostic(LegacyXlsImportDiagnostic diagnostic) {
            string scope = diagnostic.SheetName == null ? "(workbook)" : diagnostic.SheetName;
            string detail = string.IsNullOrWhiteSpace(diagnostic.DetailCode) ? diagnostic.Code : diagnostic.DetailCode!;
            return $"{scope}: {diagnostic.Code} ({detail}) - {diagnostic.Message}";
        }

        private static string FormatLegacyXlsUnsupportedFeature(LegacyXlsUnsupportedFeature feature) {
            string scope = feature.SheetName == null ? "(workbook)" : feature.SheetName;
            string detail = string.IsNullOrWhiteSpace(feature.DetailCode) ? feature.Code : feature.DetailCode!;
            return $"{scope}: {feature.Kind} ({detail}) - {feature.Description}";
        }

        private static string FormatLegacyXlsPreservedFeature(LegacyXlsPreservedFeatureRecord feature) {
            string scope = feature.SheetName == null ? "(workbook)" : feature.SheetName;
            string detail = string.IsNullOrWhiteSpace(feature.DetailCode) ? feature.Code : feature.DetailCode!;
            return $"{scope}: {feature.Code} / {feature.Kind} ({detail}) - {feature.Description}";
        }

        private static string FormatLegacyXlsUnsupportedSheet(LegacyXlsUnsupportedSheet sheet) {
            return $"{sheet.Name}: {sheet.Kind} ({sheet.VisibilityName})";
        }

        private static string FormatLegacyXlsChartSheet(LegacyXlsChartSheet sheet) {
            return $"{sheet.Name}: ChartSheet ({sheet.VisibilityName})";
        }

        private static string FormatLegacyXlsCompoundFeature(LegacyXlsCompoundFeatureRecord feature) {
            return $"{feature.Kind}: Entries:{feature.Entries.Count}";
        }

        private static void Add(List<ExcelFeatureFinding> features, string category, string name, ExcelFeatureSupportLevel supportLevel, int count,
            string? scope, string note, IReadOnlyList<string>? details = null) {
            if (count <= 0 && supportLevel != ExcelFeatureSupportLevel.Editable) {
                return;
            }

            features.Add(new ExcelFeatureFinding(category, name, supportLevel, count, scope, note, details));
        }

        private static IEnumerable<OpenXmlPart> EnumeratePackageParts(OpenXmlPackage package) {
            var seen = new HashSet<Uri>();
            foreach (var pair in package.Parts) {
                foreach (var part in EnumeratePartAndChildren(pair.OpenXmlPart, seen)) {
                    yield return part;
                }
            }
        }

        private static IEnumerable<OpenXmlPart> EnumeratePartAndChildren(OpenXmlPart part, HashSet<Uri> seen) {
            if (!seen.Add(part.Uri)) {
                yield break;
            }

            yield return part;

            foreach (var child in EnumerateParts(part, seen)) {
                yield return child;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPartContainer container, HashSet<Uri> seen) {
            foreach (var pair in container.Parts) {
                var part = pair.OpenXmlPart;
                foreach (var child in EnumeratePartAndChildren(part, seen)) {
                    yield return child;
                }
            }
        }

        private static int CountPartsByUri(IEnumerable<OpenXmlPart> parts, string uriFragment) {
            return parts.Count(part => part.Uri.OriginalString.IndexOf(uriFragment, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static int CountPartsByUriOrContentType(IEnumerable<OpenXmlPart> parts, string fragment) {
            return parts.Count(part =>
                part.Uri.OriginalString.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0
                || part.ContentType.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static List<string> DescribePartsByUri(IEnumerable<OpenXmlPart> parts, string uriFragment) {
            return parts
                .Where(part => part.Uri.OriginalString.IndexOf(uriFragment, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribePartsByUriOrContentType(IEnumerable<OpenXmlPart> parts, string fragment) {
            return parts
                .Where(part =>
                    part.Uri.OriginalString.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0
                    || part.ContentType.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeParts(IEnumerable<OpenXmlPart> parts, Func<OpenXmlPart, bool> predicate) {
            if (predicate == null) throw new ArgumentNullException(nameof(predicate));

            return parts
                .Where(predicate)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeExternalRelationships(IEnumerable<OpenXmlPart> parts, bool includeHyperlinks = true) {
            return parts
                .SelectMany(part => part.ExternalRelationships
                    .Where(relationship => includeHyperlinks || !IsHyperlinkRelationship(relationship))
                    .Select(relationship => $"{part.Uri}: {relationship.Id} -> {relationship.Uri}"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IEnumerable<string> GetFormulaCalculationBlockers(ExcelFormulaCellInfo formula) {
            foreach (string issue in formula.DependencyIssues) {
                if (issue.IndexOf("without a cached result", StringComparison.OrdinalIgnoreCase) >= 0) {
                    continue;
                }

                yield return $"{formula.SheetName}!{formula.CellReference}: {issue}";
            }

            if (!formula.IsSupportedByOfficeIMO && !IsMissingCacheOnlyFormulaCalculationGap(formula)) {
                string reason = string.IsNullOrWhiteSpace(formula.UnsupportedReason)
                    ? "Formula is outside OfficeIMO's lightweight evaluator support."
                    : formula.UnsupportedReason!;
                yield return $"{formula.SheetName}!{formula.CellReference}: {reason}";
            }
        }

        private static bool IsMissingCacheOnlyFormulaCalculationGap(ExcelFormulaCellInfo formula) {
            if (formula.DependencyIssues.Count == 0
                || formula.DependencyIssues.Any(issue => issue.IndexOf("without a cached result", StringComparison.OrdinalIgnoreCase) < 0)) {
                return false;
            }

            string reason = formula.UnsupportedReason ?? string.Empty;
            return reason.Length == 0
                || reason.IndexOf("Formula is outside OfficeIMO's lightweight evaluator support", StringComparison.OrdinalIgnoreCase) >= 0
                || reason.IndexOf("Formula uses supported function", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string DescribePart(OpenXmlPart part) {
            return $"{part.Uri} ({part.ContentType})";
        }

        private static int CountDescendantsByLocalName(DocumentFormat.OpenXml.OpenXmlElement? root, string localName) {
            if (root == null) return 0;
            return root.Descendants().Count(element => string.Equals(element.LocalName, localName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
