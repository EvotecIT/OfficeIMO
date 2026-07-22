using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// OfficeIMO's current author/edit/preserve status for a Word feature discovered during inspection.
    /// </summary>
    public enum WordFeatureSupportLevel {
        /// <summary>OfficeIMO can author or edit this feature directly.</summary>
        Editable,

        /// <summary>OfficeIMO can author or edit common cases, but not the full Word feature surface.</summary>
        PartiallyEditable,

        /// <summary>OfficeIMO should preserve the feature during round-trip saves, but does not expose a rich authoring API.</summary>
        Preserved,

        /// <summary>OfficeIMO has no meaningful support for the feature yet.</summary>
        Unsupported
    }

    /// <summary>
    /// Document-level feature and compatibility report.
    /// </summary>
    public sealed partial class WordFeatureReport {
        private readonly List<WordFeatureFinding> _features = new List<WordFeatureFinding>();

        internal WordFeatureReport(IReadOnlyList<WordFeatureFinding> features) {
            if (features == null) throw new ArgumentNullException(nameof(features));
            _features.AddRange(features);
        }

        /// <summary>
        /// Features discovered in the document.
        /// </summary>
        public IReadOnlyList<WordFeatureFinding> Features => _features;

        /// <summary>
        /// Features OfficeIMO can author or edit directly.
        /// </summary>
        public IReadOnlyList<WordFeatureFinding> EditableFeatures => _features
            .Where(feature => feature.SupportLevel == WordFeatureSupportLevel.Editable)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO can partly author or edit.
        /// </summary>
        public IReadOnlyList<WordFeatureFinding> PartiallyEditableFeatures => _features
            .Where(feature => feature.SupportLevel == WordFeatureSupportLevel.PartiallyEditable)
            .ToArray();

        /// <summary>
        /// Advanced features OfficeIMO should preserve but cannot fully author or edit yet.
        /// </summary>
        public IReadOnlyList<WordFeatureFinding> PreservedFeatures => _features
            .Where(feature => feature.SupportLevel == WordFeatureSupportLevel.Preserved)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO does not meaningfully support yet.
        /// </summary>
        public IReadOnlyList<WordFeatureFinding> UnsupportedFeatures => _features
            .Where(feature => feature.SupportLevel == WordFeatureSupportLevel.Unsupported)
            .ToArray();

        /// <summary>
        /// Whether the document contains non-editable advanced features that should be checked before edit-heavy round trips.
        /// </summary>
        public bool HasAdvancedFeatures => PreservedFeatures.Count > 0 || UnsupportedFeatures.Count > 0;

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public IReadOnlyList<WordFeatureFinding> FindFeatures(params string[] featureNames) {
            return FindFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public IReadOnlyList<WordFeatureFinding> FindFeatures(IEnumerable<string> featureNames) {
            if (featureNames == null) throw new ArgumentNullException(nameof(featureNames));
            var names = new HashSet<string>(featureNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (names.Count == 0) return Array.Empty<WordFeatureFinding>();

            return _features
                .Where(feature => names.Contains(feature.Name))
                .ToArray();
        }

        /// <summary>
        /// Returns discovered features with one of the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to match.</param>
        public IReadOnlyList<WordFeatureFinding> FindFeatures(params WordFeatureSupportLevel[] supportLevels) {
            if (supportLevels == null) throw new ArgumentNullException(nameof(supportLevels));
            if (supportLevels.Length == 0) return Array.Empty<WordFeatureFinding>();

            var levels = new HashSet<WordFeatureSupportLevel>(supportLevels);
            return _features
                .Where(feature => levels.Contains(feature.SupportLevel))
                .ToArray();
        }

        /// <summary>
        /// Throws when the document contains unsupported features.
        /// </summary>
        public WordFeatureReport EnsureNoUnsupportedFeatures() {
            if (UnsupportedFeatures.Count > 0) {
                ThrowBlockedFeatures("Unsupported document features", UnsupportedFeatures);
            }

            return this;
        }

        /// <summary>
        /// Throws when the document contains preserve-only or unsupported advanced features.
        /// </summary>
        public WordFeatureReport EnsureNoAdvancedFeatures() {
            if (HasAdvancedFeatures) {
                ThrowBlockedFeatures("Advanced document features need review before edit-heavy round trips", PreservedFeatures.Concat(UnsupportedFeatures));
            }

            return this;
        }

        /// <summary>
        /// Throws when the document contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public WordFeatureReport EnsureNoFeatures(params string[] featureNames) {
            return EnsureNoFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Throws when the document contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public WordFeatureReport EnsureNoFeatures(IEnumerable<string> featureNames) {
            var matches = FindFeatures(featureNames);
            if (matches.Count > 0) {
                ThrowBlockedFeatures("Document contains blocked features", matches);
            }

            return this;
        }

        /// <summary>
        /// Throws when the document contains any features with the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to reject.</param>
        public WordFeatureReport EnsureNoFeatures(params WordFeatureSupportLevel[] supportLevels) {
            var matches = FindFeatures(supportLevels);
            if (matches.Count > 0) {
                ThrowBlockedFeatures("Document contains blocked feature support levels", matches);
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of discovered document features and support status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Word Feature Report");
            builder.AppendLine();
            builder.AppendLine($"Total findings: {Features.Count}");
            builder.AppendLine($"Editable features: {EditableFeatures.Count}");
            builder.AppendLine($"Partially editable features: {PartiallyEditableFeatures.Count}");
            builder.AppendLine($"Preserved features: {PreservedFeatures.Count}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatures.Count}");
            builder.AppendLine();
            builder.AppendLine("| Category | Feature | Count | Support | Scope | Note | Details |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (WordFeatureFinding feature in Features) {
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

            builder.AppendLine();
            builder.AppendLine("## Capability Preflight");
            builder.AppendLine();
            builder.AppendLine("| Capability | Available | Diagnostics |");
            builder.AppendLine("| --- | --- | --- |");
            foreach (WordPreflightCapability capability in global::OfficeIMO.Internal.EnumCompat.GetValues<WordPreflightCapability>()) {
                builder.Append("| ");
                builder.Append(capability);
                builder.Append(" | ");
                builder.Append(Can(capability) ? "Yes" : "No");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(string.Join("; ", GetCapabilityDiagnostics(capability))));
                builder.AppendLine(" |");
            }

            WordPreflightRepairHint[] repairHints = global::OfficeIMO.Internal.EnumCompat.GetValues<WordPreflightCapability>()
                .SelectMany(GetRepairHints)
                .GroupBy(hint => hint.Capability + "\u001f" + hint.FeatureName + "\u001f" + hint.Action,
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            if (repairHints.Length > 0) {
                builder.AppendLine();
                builder.AppendLine("## Repair And Routing Hints");
                builder.AppendLine();
                builder.AppendLine("| Capability | Feature | Action | API | Details |");
                builder.AppendLine("| --- | --- | --- | --- | --- |");
                foreach (WordPreflightRepairHint hint in repairHints) {
                    builder.Append("| ");
                    builder.Append(hint.Capability);
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

        private static void ThrowBlockedFeatures(string message, IEnumerable<WordFeatureFinding> findings) {
            var formatted = findings
                .OrderBy(feature => feature.Name, StringComparer.OrdinalIgnoreCase)
                .Select(FormatBlockedFeature)
                .ToArray();
            throw new InvalidOperationException(message + ": " + string.Join(", ", formatted));
        }

        private static string FormatBlockedFeature(WordFeatureFinding feature) {
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
    /// One feature discovered in a document.
    /// </summary>
    public sealed class WordFeatureFinding {
        internal WordFeatureFinding(string category, string name, WordFeatureSupportLevel supportLevel, int count, string? scope, string note,
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
        /// Broad feature area, for example content, collaboration, or compatibility.
        /// </summary>
        public string Category { get; }

        /// <summary>
        /// Feature name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// OfficeIMO support status for this feature.
        /// </summary>
        public WordFeatureSupportLevel SupportLevel { get; }

        /// <summary>
        /// Number of matching items discovered.
        /// </summary>
        public int Count { get; }

        /// <summary>
        /// Optional document or package scope.
        /// </summary>
        public string? Scope { get; }

        /// <summary>
        /// Short explanation of what OfficeIMO can do with this feature today.
        /// </summary>
        public string Note { get; }

        /// <summary>
        /// Optional package, relationship, or document details that explain where the feature was found.
        /// </summary>
        public IReadOnlyList<string> Details { get; }
    }

    public partial class WordDocument {
        /// <summary>
        /// Inspects document features and reports which ones OfficeIMO can edit, partially edit, preserve, or does not support yet.
        /// </summary>
        public WordFeatureReport InspectFeatures() {
            var features = new List<WordFeatureFinding>();
            MainDocumentPart mainPart = _wordprocessingDocument.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            var document = mainPart.Document ?? throw new InvalidOperationException("Document is missing.");
            var allParts = EnumerateWordPartsAndRoot(mainPart).ToList();
            var bibliographyDetails = DescribeBibliographyParts(allParts);

            Add(features, "Content", "Paragraphs", WordFeatureSupportLevel.Editable, Paragraphs.Count, null,
                "Paragraphs, runs, common formatting, styles, lists, bookmarks, fields, and hyperlinks can be authored and inspected.");
            Add(features, "Content", "Tables", WordFeatureSupportLevel.Editable, TablesIncludingNestedTables.Count, null,
                "Tables, rows, cells, borders, widths, merges, and nested table structures can be authored and inspected.");
            Add(features, "Content", "Sections", WordFeatureSupportLevel.Editable, Sections.Count, null,
                "Sections, page settings, margins, columns, headers, and footers can be authored and inspected.");
            Add(features, "Media", "Images", WordFeatureSupportLevel.PartiallyEditable, Images.Count, null,
                "Images can be inserted and inspected in common document/header/footer scenarios; advanced drawing behaviors remain partial.");
            IReadOnlyList<WordFieldInfo> fieldInventory = InspectFields();
            WordReviewInfo reviewInfo = InspectReview();
            Add(features, "Content", "Fields", WordFeatureSupportLevel.PartiallyEditable, Math.Max(Fields.Count, fieldInventory.Count), null,
                "Common field codes can be authored, updated, and inventoried with InspectFields(); full Word field evaluation remains partial.",
                DescribeFieldInventory(fieldInventory));
            Add(features, "Content", "Bookmarks", WordFeatureSupportLevel.Editable, Bookmarks.Count, null,
                "Bookmarks can be authored, inspected, and used as hyperlink anchors.");
            Add(features, "Content", "Document variables", WordFeatureSupportLevel.Editable, DocumentVariables.Count, null,
                "Document variables can be authored, inspected, updated, and removed.");
            Add(features, "References", "Bibliography sources", WordFeatureSupportLevel.Editable, Math.Max(BibliographySources.Count, bibliographyDetails.Count), null,
                "Bibliography sources can be authored, loaded, and used by citation and bibliography fields.",
                bibliographyDetails);
            Add(features, "Content", "Footnotes", WordFeatureSupportLevel.PartiallyEditable, FootNotes.Count, null,
                "Footnotes can be authored, inspected, and removed; advanced note numbering and cross-format workflows remain partial.");
            Add(features, "Content", "Endnotes", WordFeatureSupportLevel.PartiallyEditable, EndNotes.Count, null,
                "Endnotes can be authored, inspected, and removed; advanced note numbering and cross-format workflows remain partial.");
            Add(features, "Content", "External hyperlinks", WordFeatureSupportLevel.PartiallyEditable, CountExternalHyperlinks(), null,
                "External hyperlinks can be authored and edited; the report exposes external relationships for round-trip review.",
                DescribeExternalRelationships(EnumerateWordPartsAndRoot(mainPart)));
            Add(features, "Content", "Content controls", WordFeatureSupportLevel.PartiallyEditable, StructuredDocumentTags.Count, null,
                "Common content controls such as check boxes, combo boxes, dropdown lists, date pickers, picture controls, and repeating sections are editable; the full SDT surface remains partial.");
            Add(features, "Content", "Text boxes", WordFeatureSupportLevel.PartiallyEditable, TextBoxes.Count, null,
                "Text boxes can be authored and inspected in common scenarios; advanced layout behaviors remain partial.");
            Add(features, "Content", "Shapes", WordFeatureSupportLevel.PartiallyEditable, Shapes.Count, null,
                "Basic shapes can be authored and inspected; complex drawing behaviors remain partial.");
            var chartDetails = DescribePartsByType<ChartPart>(allParts);
            Add(features, "Visualization", "Charts", WordFeatureSupportLevel.PartiallyEditable, Math.Max(Charts.Count, chartDetails.Count), null,
                "Common chart authoring is supported; advanced chart editing remains partial.",
                chartDetails);
            var smartArtDataDetails = DescribePartsByType<DiagramDataPart>(allParts);
            var smartArtDetails = DescribeDiagramParts(allParts);
            Add(features, "Visualization", "SmartArt", WordFeatureSupportLevel.Preserved, Math.Max(SmartArts.Count, smartArtDataDetails.Count), null,
                "SmartArt diagrams are detected with related diagram package parts and should be treated as preserve-only advanced drawing content.",
                smartArtDetails);
            var equationDetails = DescribeElementsByLocalName(allParts, "oMath");
            Add(features, "Math", "Equations", WordFeatureSupportLevel.PartiallyEditable, Math.Max(Equations.Count, CountElementsByLocalName(allParts, "oMath")), null,
                "Equations can be discovered across document parts; rich equation authoring and editing remains partial.",
                equationDetails);
            Add(features, "Review", "Comments", WordFeatureSupportLevel.PartiallyEditable, reviewInfo.CommentCount, null,
                "Comments, replies, resolved state, target text, and authors can be inspected through InspectReview(); edit operations remain partial.",
                DescribeReviewComments(reviewInfo));

            Add(features, "Review", "Revisions", WordFeatureSupportLevel.PartiallyEditable, reviewInfo.RevisionCount, null,
                "Inserted, deleted, move, and common formatting revisions can be inspected through InspectReview(); accept/reject operations remain broader but less granular.",
                DescribeReviewRevisions(reviewInfo));

            int protectionCount = CountDescendantsByLocalName(mainPart.DocumentSettingsPart?.Settings, "documentProtection");
            Add(features, "Protection", "Document protection", WordFeatureSupportLevel.PartiallyEditable, protectionCount, null,
                "Document protection metadata can be inspected through settings; complete protection workflows remain partial.");

            var vbaDetails = DescribePartsByUriOrContentType(allParts, "vbaProject");
            var legacyDocPreservedDetails = LegacyDocPreservedFeatures
                .Select(feature => $"{feature.Kind}: {feature.DetailCode}")
                .ToArray();
            var legacyDocCompoundDetails = LegacyDocCompoundFeatures
                .Select(feature => $"{feature.Kind}: {feature.DetailCode}; Entry={feature.EntryPath ?? string.Empty}; Entries={feature.EntryCount}; Bytes={feature.TotalBytes}")
                .ToArray();
            var altChunkDetails = DescribePartsByType<AlternativeFormatImportPart>(allParts);
            var externalImageDetails = DescribeExternalRelationshipsByType(allParts, "relationships/image");
            var attachedTemplateDetails = DescribeExternalRelationshipsByType(allParts, "attachedTemplate");
            var contentControlDataBindingDetails = DescribeContentControlDataBindings(allParts);
            var glossaryDetails = DescribePartsByUriOrContentType(allParts, "glossary");
            var modernCommentDetails = DescribePartsByUriOrContentType(allParts, "commentsExtended")
                .Concat(DescribePartsByUriOrContentType(allParts, "commentsIds"))
                .Concat(DescribePartsByUriOrContentType(allParts, "people"))
                .Concat(reviewInfo.UnsupportedMetadata)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var webExtensionDetails = DescribePartsByUriOrContentType(allParts, "webextension")
                .Concat(DescribePartsByUriOrContentType(allParts, "taskpane"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var embeddedPackageDetails = DescribePartsByType<EmbeddedPackagePart>(allParts)
                .Concat(DescribePartsByUri(allParts, "/embeddings/"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var activeXControlDetails = DescribePartsByUriOrContentType(allParts, "activeX")
                .Concat(DescribeExternalRelationshipsByType(allParts, "activeX"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var customXmlDetails = DescribePartsByUri(allParts, "/customXml/");
            WordSignatureValidationReport signatureValidation = ValidateSignatures();

            Add(features, "Compatibility", "Alternative format imports", WordFeatureSupportLevel.PartiallyEditable, altChunkDetails.Count, null,
                "Alternative-format imports can be authored, extracted, and removed through embedded document APIs; imported content remains package-backed until Word processes it.",
                altChunkDetails);
            Add(features, "Media", "External linked images", WordFeatureSupportLevel.PartiallyEditable, externalImageDetails.Count, null,
                "Externally linked images can be authored, inspected, and removed; extracting or saving image bytes requires embedded image data.",
                externalImageDetails);
            Add(features, "Compatibility", "Attached templates", WordFeatureSupportLevel.Preserved, attachedTemplateDetails.Count, null,
                "Attached template relationships are detected as preserve-only package metadata before edit-heavy workflows.",
                attachedTemplateDetails);
            Add(features, "Compatibility", "Legacy DOC preserved metadata", WordFeatureSupportLevel.Preserved, legacyDocPreservedDetails.Length, "Legacy DOC",
                "Legacy binary DOC picture and revision-tracking indicators are detected as preserve-only import metadata.",
                legacyDocPreservedDetails);
            Add(features, "Compatibility", "Legacy DOC compound storage", WordFeatureSupportLevel.Preserved, legacyDocCompoundDetails.Length, "Legacy DOC",
                "Legacy binary DOC compound storage such as macros, embedded objects, ActiveX controls, embedded packages, and binary payload streams is detected as preserve-only import metadata.",
                legacyDocCompoundDetails);
            Add(features, "Content", "Content-control data bindings", WordFeatureSupportLevel.PartiallyEditable, contentControlDataBindingDetails.Count, null,
                "Bound content controls can be refreshed from backing Custom XML or filled from supplied values with backing XML updates; broader SDT mapping workflows remain partial.",
                contentControlDataBindingDetails);
            Add(features, "Compatibility", "Building blocks and glossary", WordFeatureSupportLevel.Preserved, glossaryDetails.Count, null,
                "Glossary/building-block package parts are detected as preserve-only document metadata.",
                glossaryDetails);
            Add(features, "Review", "Modern comment metadata", WordFeatureSupportLevel.Preserved, modernCommentDetails.Count, null,
                "Modern threaded/resolved comment metadata is detected as preserve-only review metadata.",
                modernCommentDetails);
            Add(features, "Compatibility", "Web extensions and task panes", WordFeatureSupportLevel.Preserved, webExtensionDetails.Count, null,
                "Office add-in and task-pane package metadata is detected as preserve-only advanced content.",
                webExtensionDetails);
            Add(features, "Compatibility", "Embedded packages", WordFeatureSupportLevel.PartiallyEditable, embeddedPackageDetails.Count, null,
                "Embedded package and OLE payloads can be inventoried, hash-checked, extracted with byte limits, replaced, and removed; authoring remains available through the embedded-object API.",
                embeddedPackageDetails);
            Add(features, "Compatibility", "ActiveX controls", WordFeatureSupportLevel.Preserved, activeXControlDetails.Count, null,
                "ActiveX control package metadata is detected as preserve-only advanced document content.",
                activeXControlDetails);
            Add(features, "Compatibility", "VBA macros", WordFeatureSupportLevel.PartiallyEditable, vbaDetails.Count, null,
                "VBA projects can be attached, hash-checked, extracted with byte limits, enumerated, and removed; OfficeIMO does not edit VBA source or sign macro projects.",
                vbaDetails);
            Add(features, "Compatibility", "Custom XML parts", WordFeatureSupportLevel.Preserved, customXmlDetails.Count, null,
                "Custom XML parts are preserve-only package metadata.",
                customXmlDetails);
            Add(features, "Compatibility", "Digital signatures", WordFeatureSupportLevel.Unsupported, signatureValidation.SignatureInfo.FindingCount, null,
                "Digital signature package metadata, XML signature structure, reference digest method/value metadata, and signed package-part references can be inspected; cryptographic validation, digest verification, certificate trust, revocation, timestamp, and package signing are not implemented.",
                signatureValidation.SignatureInfo.Details.Concat(signatureValidation.Findings).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());

            return new WordFeatureReport(features);
        }

        private static void Add(List<WordFeatureFinding> features, string category, string name, WordFeatureSupportLevel supportLevel, int count,
            string? scope, string note, IReadOnlyList<string>? details = null) {
            if (count <= 0 && supportLevel != WordFeatureSupportLevel.Editable) {
                return;
            }

            features.Add(new WordFeatureFinding(category, name, supportLevel, count, scope, note, details));
        }

        private int CountExternalHyperlinks() {
            return EnumerateWordPartsAndRoot(_wordprocessingDocument.MainDocumentPart!)
                .SelectMany(part => part.HyperlinkRelationships)
                .Count(relationship => relationship.IsExternal);
        }

        private static IEnumerable<OpenXmlPart> EnumerateWordPartsAndRoot(OpenXmlPart root) {
            yield return root;

            foreach (var part in EnumerateWordParts(root)) {
                yield return part;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateWordParts(OpenXmlPartContainer container) {
            foreach (var pair in container.Parts) {
                var part = pair.OpenXmlPart;
                yield return part;

                foreach (var child in EnumerateWordParts(part)) {
                    yield return child;
                }
            }
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

        private static List<string> DescribePartsByType<TPart>(IEnumerable<OpenXmlPart> parts)
            where TPart : OpenXmlPart {
            return parts
                .OfType<TPart>()
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeDiagramParts(IEnumerable<OpenXmlPart> parts) {
            return DescribePartsByType<DiagramDataPart>(parts)
                .Concat(DescribePartsByType<DiagramLayoutDefinitionPart>(parts))
                .Concat(DescribePartsByType<DiagramColorsPart>(parts))
                .Concat(DescribePartsByType<DiagramStylePart>(parts))
                .Concat(DescribePartsByType<DiagramPersistLayoutPart>(parts))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeFieldInventory(IReadOnlyList<WordFieldInfo> fields) {
            if (fields.Count == 0) {
                return new List<string>();
            }

            var refreshableFields = fields
                .Where(field => field.FieldType != null && IsDeterministicFieldRefreshCandidate(field.FieldType.Value))
                .ToArray();
            var queuedOrManualFields = fields
                .Where(field => field.FieldType != null && IsQueuedOrManualFieldRefresh(field.FieldType.Value))
                .ToArray();
            var knownUnsupportedFields = fields
                .Where(field =>
                    field.FieldType != null
                    && !IsDeterministicFieldRefreshCandidate(field.FieldType.Value)
                    && !IsQueuedOrManualFieldRefresh(field.FieldType.Value))
                .ToArray();

            var details = new List<string> {
                $"Simple fields: {fields.Count(field => field.Representation == WordFieldRepresentation.Simple)}",
                $"Complex fields: {fields.Count(field => field.Representation == WordFieldRepresentation.Complex)}",
                $"Deterministic refresh candidates: {refreshableFields.Length}"
            };

            AddGroupedFieldTypes(details, "Refreshable field types", refreshableFields);
            AddGroupedFieldTypes(details, "Queued/manual refresh fields", queuedOrManualFields);
            AddGroupedFieldTypes(details, "Known unsupported refresh fields", knownUnsupportedFields);

            string locations = string.Join(", ",
                fields
                    .GroupBy(field => field.LocationKind)
                    .OrderBy(group => group.Key.ToString(), StringComparer.Ordinal)
                    .Select(group => $"{group.Key}: {group.Count()}"));
            if (!string.IsNullOrWhiteSpace(locations)) {
                details.Add("Locations: " + locations);
            }

            string parsedTypes = string.Join(", ",
                fields
                    .Where(field => field.FieldType != null)
                    .Select(field => field.FieldType!.Value.ToString())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(parsedTypes)) {
                details.Add("Parsed field types: " + parsedTypes);
            }

            int unsupportedCount = fields.Count(field => field.UnsupportedParseDetails.Count > 0);
            if (unsupportedCount > 0) {
                details.Add($"Field parser diagnostics: {unsupportedCount}");
            }

            int dirtyCount = fields.Count(field => field.IsDirty);
            if (dirtyCount > 0) {
                details.Add($"Dirty fields: {dirtyCount}");
            }

            int lockedCount = fields.Count(field => field.IsLocked);
            if (lockedCount > 0) {
                details.Add($"Locked fields: {lockedCount}");
            }

            var containerDetails = new[] {
                new { Name = "tables", Count = fields.Count(field => field.IsInTable) },
                new { Name = "content controls", Count = fields.Count(field => field.IsInContentControl) },
                new { Name = "text boxes", Count = fields.Count(field => field.IsInTextBox) }
            }
                .Where(item => item.Count > 0)
                .Select(item => $"{item.Name}: {item.Count}")
                .ToArray();
            if (containerDetails.Length > 0) {
                details.Add("Container fields: " + string.Join(", ", containerDetails));
            }

            return details;
        }

        private static void AddGroupedFieldTypes(List<string> details, string label, IReadOnlyList<WordFieldInfo> fields) {
            if (fields.Count == 0) {
                return;
            }

            string grouped = string.Join(", ",
                fields
                    .Where(field => field.FieldType != null)
                    .GroupBy(field => field.FieldType!.Value)
                    .OrderBy(group => group.Key.ToString(), StringComparer.Ordinal)
                    .Select(group => $"{group.Key}: {group.Count()}"));
            if (!string.IsNullOrWhiteSpace(grouped)) {
                details.Add(label + ": " + grouped);
            }
        }

        private static bool IsDeterministicFieldRefreshCandidate(WordFieldType fieldType) {
            switch (fieldType) {
                case WordFieldType.Author:
                case WordFieldType.Comments:
                case WordFieldType.CreateDate:
                case WordFieldType.Date:
                case WordFieldType.DocProperty:
                case WordFieldType.DocVariable:
                case WordFieldType.FileName:
                case WordFieldType.FileSize:
                case WordFieldType.Formula:
                case WordFieldType.Info:
                case WordFieldType.Keywords:
                case WordFieldType.LastSavedBy:
                case WordFieldType.NumChars:
                case WordFieldType.NumPages:
                case WordFieldType.NumWords:
                case WordFieldType.Page:
                case WordFieldType.PageRef:
                case WordFieldType.PrintDate:
                case WordFieldType.Quote:
                case WordFieldType.Ref:
                case WordFieldType.RevNum:
                case WordFieldType.SaveDate:
                case WordFieldType.Section:
                case WordFieldType.SectionPages:
                case WordFieldType.Seq:
                case WordFieldType.Subject:
                case WordFieldType.Time:
                case WordFieldType.Title:
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsQueuedOrManualFieldRefresh(WordFieldType fieldType) {
            switch (fieldType) {
                case WordFieldType.Index:
                case WordFieldType.TC:
                case WordFieldType.TOC:
                case WordFieldType.XE:
                    return true;
                default:
                    return false;
            }
        }

        private static List<string> DescribeReviewComments(WordReviewInfo reviewInfo) {
            if (reviewInfo.CommentCount == 0) {
                return new List<string>();
            }

            var details = new List<string> {
                $"Replies: {reviewInfo.ReplyCount}",
                $"Resolved: {reviewInfo.ResolvedCommentCount}",
                $"Known unresolved: {reviewInfo.UnresolvedCommentCount}"
            };

            string authors = string.Join(", ",
                reviewInfo.Comments
                    .Select(comment => comment.Author)
                    .Where(author => !string.IsNullOrWhiteSpace(author))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(author => author, StringComparer.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(authors)) {
                details.Add("Authors: " + authors);
            }

            string locations = string.Join(", ",
                reviewInfo.Comments
                    .Where(comment => comment.TargetLocationKind != null)
                    .GroupBy(comment => comment.TargetLocationKind!.Value)
                    .OrderBy(group => group.Key.ToString(), StringComparer.Ordinal)
                    .Select(group => $"{group.Key}: {group.Count()}"));
            if (!string.IsNullOrWhiteSpace(locations)) {
                details.Add("Locations: " + locations);
            }

            return details;
        }

        private static List<string> DescribeReviewRevisions(WordReviewInfo reviewInfo) {
            if (reviewInfo.RevisionCount == 0) {
                return new List<string>();
            }

            var details = new List<string>();
            string types = string.Join(", ",
                reviewInfo.Revisions
                    .GroupBy(revision => revision.RevisionType)
                    .OrderBy(group => group.Key.ToString(), StringComparer.Ordinal)
                    .Select(group => $"{group.Key}: {group.Count()}"));
            if (!string.IsNullOrWhiteSpace(types)) {
                details.Add("Types: " + types);
            }

            string authors = string.Join(", ",
                reviewInfo.Revisions
                    .Select(revision => revision.Author)
                    .Where(author => !string.IsNullOrWhiteSpace(author))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(author => author, StringComparer.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(authors)) {
                details.Add("Authors: " + authors);
            }

            string locations = string.Join(", ",
                reviewInfo.Revisions
                    .GroupBy(revision => revision.LocationKind)
                    .OrderBy(group => group.Key.ToString(), StringComparer.Ordinal)
                    .Select(group => $"{group.Key}: {group.Count()}"));
            if (!string.IsNullOrWhiteSpace(locations)) {
                details.Add("Locations: " + locations);
            }

            return details;
        }

        private static List<string> DescribeElementsByLocalName(IEnumerable<OpenXmlPart> parts, string localName) {
            return parts
                .Select(part => new {
                    Part = part,
                    Count = CountDescendantsByLocalName(part.RootElement, localName)
                })
                .Where(item => item.Count > 0)
                .Select(item => $"{item.Part.Uri}: {item.Count} {localName} element(s)")
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static int CountElementsByLocalName(IEnumerable<OpenXmlPart> parts, string localName) {
            return parts.Sum(part => CountDescendantsByLocalName(part.RootElement, localName));
        }

        private static List<string> DescribeBibliographyParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(IsBibliographyPart)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsBibliographyPart(OpenXmlPart part) {
            if (part.Uri.OriginalString.IndexOf("bibliography", StringComparison.OrdinalIgnoreCase) >= 0
                || part.ContentType.IndexOf("bibliography", StringComparison.OrdinalIgnoreCase) >= 0) {
                return true;
            }

            if (part is not CustomXmlPart) {
                return false;
            }

            try {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                XDocument document = XDocument.Load(stream);
                XElement? root = document.Root;
                return root != null
                    && string.Equals(root.Name.LocalName, "Sources", StringComparison.OrdinalIgnoreCase)
                    && root.Name.NamespaceName.IndexOf("bibliography", StringComparison.OrdinalIgnoreCase) >= 0;
            } catch {
                return false;
            }
        }

        private static List<string> DescribeContentControlDataBindings(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part => (part.RootElement?.Descendants<DataBinding>() ?? Enumerable.Empty<DataBinding>())
                    .Select(binding => DescribeContentControlDataBinding(part, binding)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string DescribeContentControlDataBinding(OpenXmlPart part, DataBinding binding) {
            string storeItemId = binding.StoreItemId?.Value ?? "(no store item id)";
            string xpath = binding.XPath?.Value ?? "(no xpath)";
            return $"{part.Uri}: storeItemId={storeItemId}, xpath={xpath}";
        }

        private static List<string> DescribeExternalRelationships(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part =>
                    part.ExternalRelationships.Select(relationship =>
                        $"{part.Uri}: {relationship.Id} -> {relationship.Uri}")
                    .Concat(part.HyperlinkRelationships
                        .Where(relationship => relationship.IsExternal)
                        .Select(relationship => $"{part.Uri}: {relationship.Id} -> {relationship.Uri}")))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeExternalRelationshipsByType(IEnumerable<OpenXmlPart> parts, string relationshipTypeFragment) {
            return parts
                .SelectMany(part => part.ExternalRelationships
                    .Where(relationship => relationship.RelationshipType.IndexOf(relationshipTypeFragment, StringComparison.OrdinalIgnoreCase) >= 0)
                    .Select(relationship => $"{part.Uri}: {relationship.Id} ({relationship.RelationshipType}) -> {relationship.Uri}"))
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
