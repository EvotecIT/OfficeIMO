using DocumentFormat.OpenXml.Packaging;
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
    public sealed class ExcelFeatureReport {
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
        /// Throws when the workbook contains unsupported features.
        /// </summary>
        public ExcelFeatureReport EnsureNoUnsupportedFeatures() {
            if (UnsupportedFeatures.Count > 0) {
                throw new InvalidOperationException("Unsupported workbook features: " + string.Join(", ", UnsupportedFeatures.Select(feature => feature.Name)));
            }

            return this;
        }

        /// <summary>
        /// Throws when the workbook contains preserve-only or unsupported advanced features.
        /// </summary>
        public ExcelFeatureReport EnsureNoAdvancedFeatures() {
            if (HasAdvancedFeatures) {
                var advanced = PreservedFeatures
                    .Concat(UnsupportedFeatures)
                    .Select(feature => feature.Name)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                throw new InvalidOperationException("Advanced workbook features need review before edit-heavy round trips: " + string.Join(", ", advanced));
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
            builder.AppendLine("| Category | Feature | Count | Support | Scope | Note |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- |");

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
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }

    /// <summary>
    /// One feature discovered in a workbook.
    /// </summary>
    public sealed class ExcelFeatureFinding {
        internal ExcelFeatureFinding(string category, string name, ExcelFeatureSupportLevel supportLevel, int count, string? scope, string note) {
            Category = string.IsNullOrWhiteSpace(category) ? throw new ArgumentNullException(nameof(category)) : category;
            Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentNullException(nameof(name)) : name;
            SupportLevel = supportLevel;
            Count = count;
            Scope = string.IsNullOrWhiteSpace(scope) ? null : scope;
            Note = string.IsNullOrWhiteSpace(note) ? string.Empty : note;
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

            int namedRangeCount = workbook.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Count() ?? 0;
            Add(features, "Workbook", "Named ranges", ExcelFeatureSupportLevel.Editable, namedRangeCount, null,
                "Workbook and sheet-local named ranges are editable.");

            int tableCount = 0;
            int chartCount = 0;
            int pivotCount = 0;
            int dataValidationCount = 0;
            int conditionalFormattingCount = 0;
            int sparklineCount = 0;
            int legacyCommentCount = 0;
            int threadedCommentPartCount = 0;
            int imagePartCount = 0;
            int oleObjectCount = 0;

            foreach (var sheet in sheetElements) {
                if (string.IsNullOrWhiteSpace(sheet.Id?.Value)) {
                    continue;
                }

                if (workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart) {
                    continue;
                }

                var worksheet = worksheetPart.Worksheet;
                tableCount += worksheetPart.TableDefinitionParts.Count();
                pivotCount += worksheetPart.PivotTableParts.Count();
                dataValidationCount += worksheet?.Descendants<DocumentFormat.OpenXml.Spreadsheet.DataValidation>().Count() ?? 0;
                conditionalFormattingCount += worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>().Count() ?? 0;
                sparklineCount += CountDescendantsByLocalName(worksheet, "sparkline");
                legacyCommentCount += worksheetPart.WorksheetCommentsPart?.Comments?.CommentList?.Elements<DocumentFormat.OpenXml.Spreadsheet.Comment>().Count() ?? 0;
                threadedCommentPartCount += CountPartsByUri(worksheetPart.Parts.Select(pair => pair.OpenXmlPart), "threadedComment");
                imagePartCount += worksheetPart.DrawingsPart?.ImageParts.Count() ?? 0;
                chartCount += worksheetPart.DrawingsPart?.ChartParts.Count() ?? 0;
                oleObjectCount += CountDescendantsByLocalName(worksheet, "oleObject");
            }

            Add(features, "Data", "Tables", ExcelFeatureSupportLevel.Editable, tableCount, null,
                "Tables can be authored and inspected, including AutoFilter metadata.");
            Add(features, "Data", "Data validations", ExcelFeatureSupportLevel.Editable, dataValidationCount, null,
                "List, numeric, date, time, text-length, custom formula, prompt, and error metadata are editable.");
            Add(features, "Formatting", "Conditional formatting", ExcelFeatureSupportLevel.PartiallyEditable, conditionalFormattingCount, null,
                "Common rule types are editable; full Excel conditional-formatting parity remains a roadmap item.");
            Add(features, "Visualization", "Charts", ExcelFeatureSupportLevel.PartiallyEditable, chartCount, null,
                "Common chart authoring and updates are supported; advanced chart families remain partial.");
            Add(features, "Visualization", "Pivot tables", ExcelFeatureSupportLevel.PartiallyEditable, pivotCount, null,
                "Source-range pivot creation and inspection are supported; grouping, slicers, and advanced filters remain partial.");
            Add(features, "Visualization", "Sparklines", ExcelFeatureSupportLevel.Editable, sparklineCount, null,
                "Line, column, and win/loss sparklines can be authored.");
            Add(features, "Collaboration", "Legacy comments", ExcelFeatureSupportLevel.PartiallyEditable, legacyCommentCount, null,
                "Legacy comments can be authored and inspected; rich/threaded comment workflows remain limited.");
            Add(features, "Collaboration", "Threaded comments", ExcelFeatureSupportLevel.Preserved, threadedCommentPartCount, null,
                "Threaded comments are advanced Excel metadata and should be treated as preserve-only.");
            Add(features, "Media", "Images", ExcelFeatureSupportLevel.PartiallyEditable, imagePartCount, null,
                "Images can be inserted in common worksheet/header/footer scenarios; advanced drawing behaviors remain partial.");
            Add(features, "Compatibility", "OLE objects", ExcelFeatureSupportLevel.Preserved, oleObjectCount, null,
                "Embedded OLE objects are advanced package content and should be treated as preserve-only.");

            var formulas = InspectFormulas();
            Add(features, "Calculation", "Supported formulas", ExcelFeatureSupportLevel.PartiallyEditable, formulas.SupportedFormulas, null,
                "Simple supported formulas can be recalculated by OfficeIMO.");
            Add(features, "Calculation", "Unsupported formulas", ExcelFeatureSupportLevel.Preserved, formulas.UnsupportedFormulas, null,
                "Unsupported formulas are preserved and should be recalculated by Excel or read from cached values.");
            Add(features, "Calculation", "Missing formula caches", ExcelFeatureSupportLevel.Preserved, formulas.MissingCachedResults, null,
                "Formulas without cached results need OfficeIMO calculation support or Excel recalculation before cached-value reads are reliable.");

            var allParts = EnumerateParts(workbookPart).ToList();
            int vbaPartCount = CountPartsByUriOrContentType(allParts, "vbaProject");
            int slicerPartCount = CountPartsByUriOrContentType(allParts, "slicer");
            int timelinePartCount = CountPartsByUriOrContentType(allParts, "timeline");
            int externalLinkPartCount = CountPartsByUriOrContentType(allParts, "externalLink");
            int connectionPartCount = CountPartsByUriOrContentType(allParts, "connection") + CountPartsByUriOrContentType(allParts, "queryTable");
            int customXmlPartCount = CountPartsByUri(allParts, "/customXml/");
            int embeddedPackagePartCount = CountPartsByUri(allParts, "/embeddings/");
            int externalRelationshipCount = allParts.Sum(part => part.ExternalRelationships.Count());

            Add(features, "Compatibility", "VBA macros", ExcelFeatureSupportLevel.Preserved, vbaPartCount, null,
                "Macro projects are preserve-only; OfficeIMO.Excel does not author or edit VBA modules.");
            Add(features, "Compatibility", "Slicers", ExcelFeatureSupportLevel.Preserved, slicerPartCount, null,
                "Slicer metadata is preserve-only; authoring slicers remains a roadmap item.");
            Add(features, "Compatibility", "Timelines", ExcelFeatureSupportLevel.Preserved, timelinePartCount, null,
                "Timeline metadata is preserve-only; authoring timelines remains a roadmap item.");
            Add(features, "Compatibility", "External workbook links", ExcelFeatureSupportLevel.Preserved, externalLinkPartCount + externalRelationshipCount, null,
                "External relationships and workbook-link parts should be treated carefully during round trips.");
            Add(features, "Compatibility", "Connections and query tables", ExcelFeatureSupportLevel.Preserved, connectionPartCount, null,
                "Connections and query-table metadata are preserve-only.");
            Add(features, "Compatibility", "Custom XML parts", ExcelFeatureSupportLevel.Preserved, customXmlPartCount, null,
                "Custom XML parts are preserve-only package metadata.");
            Add(features, "Compatibility", "Embedded packages", ExcelFeatureSupportLevel.Preserved, embeddedPackagePartCount, null,
                "Embedded packages are preserve-only package content.");

            return new ExcelFeatureReport(features);
        }

        private static void Add(List<ExcelFeatureFinding> features, string category, string name, ExcelFeatureSupportLevel supportLevel, int count, string? scope, string note) {
            if (count <= 0 && supportLevel != ExcelFeatureSupportLevel.Editable) {
                return;
            }

            features.Add(new ExcelFeatureFinding(category, name, supportLevel, count, scope, note));
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPartContainer container) {
            foreach (var pair in container.Parts) {
                var part = pair.OpenXmlPart;
                yield return part;

                foreach (var child in EnumerateParts(part)) {
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

        private static int CountDescendantsByLocalName(DocumentFormat.OpenXml.OpenXmlElement? root, string localName) {
            if (root == null) return 0;
            return root.Descendants().Count(element => string.Equals(element.LocalName, localName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
