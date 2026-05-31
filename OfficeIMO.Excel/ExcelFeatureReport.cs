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
            int formControlCount = 0;
            int externalHyperlinkCount = 0;
            var threadedCommentDetails = new List<string>();
            var oleObjectDetails = new List<string>();
            var formControlDetails = new List<string>();
            var externalHyperlinkDetails = new List<string>();
            var threadedCommentPeople = BuildThreadedCommentPersonMap(workbookPart);

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
                var threadedComments = BuildThreadedCommentMap(worksheetPart, threadedCommentPeople)
                    .Values
                    .SelectMany(comments => comments)
                    .ToList();
                threadedCommentPartCount += threadedComments.Count;
                foreach (var threadedComment in threadedComments) {
                    string author = string.IsNullOrWhiteSpace(threadedComment.Author) ? threadedComment.PersonId ?? "unknown author" : threadedComment.Author!;
                    threadedCommentDetails.Add($"{sheet.Name}: {threadedComment.CellReference} by {author}");
                }
                imagePartCount += worksheetPart.DrawingsPart?.ImageParts.Count() ?? 0;
                chartCount += worksheetPart.DrawingsPart?.ChartParts.Count() ?? 0;
                int sheetOleObjects = CountDescendantsByLocalName(worksheet, "oleObject");
                int sheetFormControls = CountDescendantsByLocalName(worksheet, "control") + CountDescendantsByLocalName(worksheet, "formControl");
                oleObjectCount += sheetOleObjects;
                formControlCount += sheetFormControls;
                externalHyperlinkCount += worksheetPart.HyperlinkRelationships.Count();
                if (sheetOleObjects > 0) oleObjectDetails.Add($"{sheet.Name}: {sheetOleObjects} OLE object(s)");
                if (sheetFormControls > 0) formControlDetails.Add($"{sheet.Name}: {sheetFormControls} form control marker(s)");
                foreach (var relationship in worksheetPart.HyperlinkRelationships) {
                    externalHyperlinkDetails.Add($"{sheet.Name}: {relationship.Id} -> {relationship.Uri}");
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
            Add(features, "Visualization", "Pivot tables", ExcelFeatureSupportLevel.PartiallyEditable, pivotCount, null,
                "Source-range pivot creation and inspection are supported, including composable fluent field sort/subtotal/layout/display/number-format helpers with built-in/custom id/code readback, field item/page filters with fluent helpers plus hidden, visible, and selected-item readback, common label/value filters, negated filter variants, fixed and dynamic date filters, top/bottom count/percent/sum filters, formula-backed calculated fields with number-format id/code readback, date/number grouping metadata, generated multi-level date hierarchy fields with base/parent relationships, and explicit grouped-cache item metadata; slicers, deeper Excel interoperability checks, and advanced filters remain partial.");
            Add(features, "Visualization", "Sparklines", ExcelFeatureSupportLevel.Editable, sparklineCount, null,
                "Line, column, and win/loss sparklines can be authored.");
            Add(features, "Collaboration", "Legacy comments", ExcelFeatureSupportLevel.PartiallyEditable, legacyCommentCount, null,
                "Legacy comments can be authored and inspected, including rich-text runs for authored comments; threaded comment workflows remain preserve-only.");
            Add(features, "Collaboration", "Threaded comments", ExcelFeatureSupportLevel.Preserved, threadedCommentPartCount, null,
                "Threaded comments can be inspected and round-trip preserved, but authoring/editing modern conversations remains preserve-only.",
                threadedCommentDetails);
            Add(features, "Media", "Images", ExcelFeatureSupportLevel.PartiallyEditable, imagePartCount, null,
                "Images can be inserted in common worksheet/header/footer scenarios; advanced drawing behaviors remain partial.");
            Add(features, "Compatibility", "OLE objects", ExcelFeatureSupportLevel.Preserved, oleObjectCount, null,
                "Embedded OLE objects are advanced package content and should be treated as preserve-only.", oleObjectDetails);
            Add(features, "Compatibility", "Form controls", ExcelFeatureSupportLevel.Preserved, formControlCount, null,
                "Form controls are preserve-only worksheet metadata.", formControlDetails);

            var formulas = InspectFormulas();
            Add(features, "Calculation", "Supported formulas", ExcelFeatureSupportLevel.PartiallyEditable, formulas.SupportedFormulas, null,
                "Simple supported formulas can be recalculated by OfficeIMO.");
            Add(features, "Calculation", "Unsupported formulas", ExcelFeatureSupportLevel.Preserved, formulas.UnsupportedFormulas, null,
                "Unsupported formulas are preserved and should be recalculated by Excel or read from cached values.");
            Add(features, "Calculation", "Missing formula caches", ExcelFeatureSupportLevel.Preserved, formulas.MissingCachedResults, null,
                "Formulas without cached results need OfficeIMO calculation support or Excel recalculation before cached-value reads are reliable.");

            var allParts = EnumeratePackageParts(_spreadSheetDocument).ToList();
            var vbaDetails = DescribePartsByUriOrContentType(allParts, "vbaProject");
            var slicerDetails = DescribePartsByUriOrContentType(allParts, "slicer");
            var timelineDetails = DescribePartsByUriOrContentType(allParts, "timeline");
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
            var externalRelationshipDetails = DescribeExternalRelationships(allParts)
                .Concat(externalHyperlinkDetails)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            Add(features, "Compatibility", "VBA macros", ExcelFeatureSupportLevel.Preserved, vbaDetails.Count, null,
                "Macro projects are preserve-only; OfficeIMO.Excel does not author or edit VBA modules.", vbaDetails);
            Add(features, "Compatibility", "Slicers", ExcelFeatureSupportLevel.Preserved, slicerDetails.Count, null,
                "Slicer metadata is preserve-only; authoring slicers remains a roadmap item.", slicerDetails);
            Add(features, "Compatibility", "Timelines", ExcelFeatureSupportLevel.Preserved, timelineDetails.Count, null,
                "Timeline metadata is preserve-only; authoring timelines remains a roadmap item.", timelineDetails);
            Add(features, "Compatibility", "External workbook links", ExcelFeatureSupportLevel.Preserved, externalLinkDetails.Count + externalRelationshipDetails.Count, null,
                "External relationships, external hyperlinks, and workbook-link parts should be treated carefully during round trips.",
                externalLinkDetails.Concat(externalRelationshipDetails).ToArray());
            Add(features, "Compatibility", "Connections and query tables", ExcelFeatureSupportLevel.Preserved, connectionDetails.Count, null,
                "Connections and query-table metadata are preserve-only.", connectionDetails);
            Add(features, "Compatibility", "Custom XML parts", ExcelFeatureSupportLevel.Preserved, customXmlDetails.Count, null,
                "Custom XML parts are preserve-only package metadata.", customXmlDetails);
            Add(features, "Compatibility", "Digital signatures", ExcelFeatureSupportLevel.Preserved, signatureDetails.Count, null,
                "Digital signature parts are preserve-only package metadata.", signatureDetails);
            Add(features, "Compatibility", "Embedded packages", ExcelFeatureSupportLevel.Preserved, embeddedPackageDetails.Count, null,
                "Embedded packages are preserve-only package content.", embeddedPackageDetails);

            return new ExcelFeatureReport(features);
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

        private static List<string> DescribeExternalRelationships(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part => part.ExternalRelationships.Select(relationship =>
                    $"{part.Uri}: {relationship.Id} -> {relationship.Uri}"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
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
