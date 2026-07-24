using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// OfficeIMO's current author/edit/preserve status for a PowerPoint feature discovered during inspection.
    /// </summary>
    public enum PowerPointFeatureSupportLevel {
        /// <summary>OfficeIMO can author or edit this feature directly.</summary>
        Editable,

        /// <summary>OfficeIMO can author or edit common cases, but not the full PowerPoint feature surface.</summary>
        PartiallyEditable,

        /// <summary>OfficeIMO should preserve the feature during round-trip saves, but does not expose a rich authoring API.</summary>
        Preserved,

        /// <summary>OfficeIMO has no meaningful support for the feature yet.</summary>
        Unsupported
    }

    /// <summary>
    /// Presentation-level feature and compatibility report.
    /// </summary>
    public sealed class PowerPointFeatureReport {
        private readonly List<PowerPointFeatureFinding> _features = new();

        internal PowerPointFeatureReport(IReadOnlyList<PowerPointFeatureFinding> features) {
            if (features == null) throw new ArgumentNullException(nameof(features));
            _features.AddRange(features);
        }

        /// <summary>
        /// Features discovered in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointFeatureFinding> Features => _features;

        /// <summary>
        /// Features OfficeIMO can author or edit directly.
        /// </summary>
        public IReadOnlyList<PowerPointFeatureFinding> EditableFeatures => _features
            .Where(feature => feature.SupportLevel == PowerPointFeatureSupportLevel.Editable)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO can partly author or edit.
        /// </summary>
        public IReadOnlyList<PowerPointFeatureFinding> PartiallyEditableFeatures => _features
            .Where(feature => feature.SupportLevel == PowerPointFeatureSupportLevel.PartiallyEditable)
            .ToArray();

        /// <summary>
        /// Advanced features OfficeIMO should preserve but cannot fully author or edit yet.
        /// </summary>
        public IReadOnlyList<PowerPointFeatureFinding> PreservedFeatures => _features
            .Where(feature => feature.SupportLevel == PowerPointFeatureSupportLevel.Preserved)
            .ToArray();

        /// <summary>
        /// Features OfficeIMO does not meaningfully support yet.
        /// </summary>
        public IReadOnlyList<PowerPointFeatureFinding> UnsupportedFeatures => _features
            .Where(feature => feature.SupportLevel == PowerPointFeatureSupportLevel.Unsupported)
            .ToArray();

        /// <summary>
        /// Whether the presentation contains non-editable advanced features that should be checked before edit-heavy round trips.
        /// </summary>
        public bool HasAdvancedFeatures => PreservedFeatures.Count > 0 || UnsupportedFeatures.Count > 0;

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public IReadOnlyList<PowerPointFeatureFinding> FindFeatures(params string[] featureNames) {
            return FindFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Returns discovered features matching the provided feature names.
        /// </summary>
        /// <param name="featureNames">Feature names to match, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public IReadOnlyList<PowerPointFeatureFinding> FindFeatures(IEnumerable<string> featureNames) {
            if (featureNames == null) throw new ArgumentNullException(nameof(featureNames));
            var names = new HashSet<string>(featureNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (names.Count == 0) return Array.Empty<PowerPointFeatureFinding>();

            return _features
                .Where(feature => names.Contains(feature.Name))
                .ToArray();
        }

        /// <summary>
        /// Returns discovered features with one of the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to match.</param>
        public IReadOnlyList<PowerPointFeatureFinding> FindFeatures(params PowerPointFeatureSupportLevel[] supportLevels) {
            if (supportLevels == null) throw new ArgumentNullException(nameof(supportLevels));
            if (supportLevels.Length == 0) return Array.Empty<PowerPointFeatureFinding>();

            var levels = new HashSet<PowerPointFeatureSupportLevel>(supportLevels);
            return _features
                .Where(feature => levels.Contains(feature.SupportLevel))
                .ToArray();
        }

        /// <summary>
        /// Throws when the presentation contains unsupported features.
        /// </summary>
        public PowerPointFeatureReport EnsureNoUnsupportedFeatures() {
            if (UnsupportedFeatures.Count > 0) {
                ThrowBlockedFeatures("Unsupported presentation features", UnsupportedFeatures);
            }

            return this;
        }

        /// <summary>
        /// Throws when the presentation contains preserve-only or unsupported advanced features.
        /// </summary>
        public PowerPointFeatureReport EnsureNoAdvancedFeatures() {
            if (HasAdvancedFeatures) {
                ThrowBlockedFeatures("Advanced presentation features need review before edit-heavy round trips", PreservedFeatures.Concat(UnsupportedFeatures));
            }

            return this;
        }

        /// <summary>
        /// Throws when the presentation contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public PowerPointFeatureReport EnsureNoFeatures(params string[] featureNames) {
            return EnsureNoFeatures((IEnumerable<string>)featureNames);
        }

        /// <summary>
        /// Throws when the presentation contains any of the named features.
        /// </summary>
        /// <param name="featureNames">Feature names to reject, for example <c>VBA macros</c> or <c>Digital signatures</c>.</param>
        public PowerPointFeatureReport EnsureNoFeatures(IEnumerable<string> featureNames) {
            var matches = FindFeatures(featureNames)
                .Where(feature => feature.Count > 0)
                .ToArray();
            if (matches.Length > 0) {
                ThrowBlockedFeatures("Presentation contains blocked features", matches);
            }

            return this;
        }

        /// <summary>
        /// Throws when the presentation contains any features with the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to reject.</param>
        public PowerPointFeatureReport EnsureNoFeatures(params PowerPointFeatureSupportLevel[] supportLevels) {
            var matches = FindFeatures(supportLevels)
                .Where(feature => feature.Count > 0)
                .ToArray();
            if (matches.Length > 0) {
                ThrowBlockedFeatures("Presentation contains blocked feature support levels", matches);
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of discovered presentation features and support status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# PowerPoint Feature Report");
            builder.AppendLine();
            builder.AppendLine($"Total findings: {Features.Count}");
            builder.AppendLine($"Editable features: {EditableFeatures.Count}");
            builder.AppendLine($"Partially editable features: {PartiallyEditableFeatures.Count}");
            builder.AppendLine($"Preserved features: {PreservedFeatures.Count}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatures.Count}");
            builder.AppendLine();
            builder.AppendLine("| Category | Feature | Count | Support | Scope | Note | Details |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (PowerPointFeatureFinding feature in Features) {
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

        private static void ThrowBlockedFeatures(string message, IEnumerable<PowerPointFeatureFinding> findings) {
            var formatted = findings
                .OrderBy(feature => feature.Name, StringComparer.OrdinalIgnoreCase)
                .Select(FormatBlockedFeature)
                .ToArray();
            throw new InvalidOperationException(message + ": " + string.Join(", ", formatted));
        }

        private static string FormatBlockedFeature(PowerPointFeatureFinding feature) {
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
    /// One feature discovered in a presentation.
    /// </summary>
    public sealed class PowerPointFeatureFinding {
        internal PowerPointFeatureFinding(string category, string name, PowerPointFeatureSupportLevel supportLevel, int count, string? scope,
            string note, IReadOnlyList<string>? details = null) {
            Category = string.IsNullOrWhiteSpace(category) ? throw new ArgumentNullException(nameof(category)) : category;
            Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentNullException(nameof(name)) : name;
            SupportLevel = supportLevel;
            Count = count;
            Scope = string.IsNullOrWhiteSpace(scope) ? null : scope;
            Note = string.IsNullOrWhiteSpace(note) ? string.Empty : note;
            Details = details ?? Array.Empty<string>();
        }

        /// <summary>
        /// Broad feature area, for example content, visualization, or compatibility.
        /// </summary>
        public string Category { get; }

        /// <summary>
        /// Feature name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// OfficeIMO support status for this feature.
        /// </summary>
        public PowerPointFeatureSupportLevel SupportLevel { get; }

        /// <summary>
        /// Number of matching items discovered.
        /// </summary>
        public int Count { get; }

        /// <summary>
        /// Optional presentation or package scope.
        /// </summary>
        public string? Scope { get; }

        /// <summary>
        /// Short explanation of what OfficeIMO can do with this feature today.
        /// </summary>
        public string Note { get; }

        /// <summary>
        /// Optional package, relationship, or slide details that explain where the feature was found.
        /// </summary>
        public IReadOnlyList<string> Details { get; }
    }

    public sealed partial class PowerPointPresentation {
        private const int MaxFeatureInspectionParts = 100_000;
        private const int MaxFeatureInspectionRelationships = 500_000;
        private const int MaxFeatureInspectionDepth = 128;

        /// <summary>
        /// Inspects presentation features and reports which ones OfficeIMO can edit, partially edit, preserve, or does not support yet.
        /// </summary>
        public PowerPointFeatureReport InspectFeatures() {
            ThrowIfDisposed();

            var features = new List<PowerPointFeatureFinding>();
            var allParts = EnumeratePowerPointPartsAndPackage(_document!, _presentationPart).ToList();
            var packageFoundationDetails = DescribePackageFoundationParts(allParts);
            var tableMetadataDetails = DescribeTableMetadata();
            var chartDetails = DescribePartsByType<ChartPart>(allParts);
            var diagramDetails = DescribeDiagramParts(allParts);
            var safeChartWorkbookParts = GetSafeChartWorkbookParts(allParts);
            var chartWorkbookDetails = DescribeChartWorkbookParts(safeChartWorkbookParts);
            var chartCompanionDetails = DescribePartsByUriOrContentType(allParts, "chartStyle")
                .Concat(DescribePartsByUriOrContentType(allParts, "chartColorStyle"))
                .Concat(chartWorkbookDetails)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var externalHyperlinkDetails = DescribeExternalHyperlinkRelationships(allParts);
            var externalPackageRelationshipDetails = DescribeExternalPackageRelationships(allParts);

            Add(features, "Structure", "Slides", PowerPointFeatureSupportLevel.Editable, Slides.Count, null,
                "Slides can be authored, duplicated, imported, hidden, reordered, and removed.");
            Add(features, "Structure", "PowerPoint package scaffolding", PowerPointFeatureSupportLevel.Editable, packageFoundationDetails.Count, null,
                "Presentation, layout, master, theme, view, table-style, notes-master, and thumbnail parts are created or preserved as package foundation.",
                packageFoundationDetails);
            Add(features, "Structure", "Sections", PowerPointFeatureSupportLevel.Editable, GetSections().Count, null,
                "Sections can be authored, inspected, renamed, moved, and synchronized with slides.");
            Add(features, "Content", "Text boxes", PowerPointFeatureSupportLevel.Editable, Slides.Sum(CountSlideTextBoxes), null,
                "Text boxes, runs, common formatting, markdown import, hyperlinks, and replacement are editable.");
            Add(features, "Content", "Tables", PowerPointFeatureSupportLevel.Editable, Slides.Sum(CountSlideTables), null,
                "Tables, cells, text, merges, dimensions, fills, borders, alignment, banding flags, and typed object binding are editable.");
            Add(features, "Content", "Table style metadata", PowerPointFeatureSupportLevel.Editable, tableMetadataDetails.Count, null,
                "Generated tables include PowerPoint-style table IDs, banding metadata, row and column IDs, and language-aware cell text defaults.",
                tableMetadataDetails);
            Add(features, "Visualization", "Charts", PowerPointFeatureSupportLevel.PartiallyEditable, Math.Max(Slides.Sum(slide => slide.Charts.Count()), chartDetails.Count), null,
                "Common chart authoring and data updates are supported; advanced chart editing remains partial.",
                chartDetails.Concat(chartCompanionDetails).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
            Add(features, "Media", "Images", PowerPointFeatureSupportLevel.PartiallyEditable, Slides.Sum(CountSlideImages), null,
                "Images can be inserted and inspected in common slide scenarios; advanced drawing behaviors remain partial.");
            Add(features, "Media", "Audio and video", PowerPointFeatureSupportLevel.PartiallyEditable, Slides.Sum(CountSlideMedia), null,
                "Embedded audio and video can be authored with poster frames and playback timing; rich media editing remains partial.",
                DescribePartsByUriOrContentType(allParts, "media"));
            Add(features, "Media", "Legacy media metadata",
                PowerPointFeatureSupportLevel.Preserved,
                LegacyPptMediaDetails.Count, null,
                "Linked, device-based, or legacy-only playback metadata from binary PowerPoint remains available for exact binary preservation but is not editable in the Open XML media surface.",
                LegacyPptMediaDetails);
            Add(features, "Visualization", "SmartArt", PowerPointFeatureSupportLevel.PartiallyEditable, Math.Max(Slides.Sum(CountSlideSmartArt), diagramDetails.Count), null,
                "SmartArt diagrams can be generated and discovered; rich diagram editing remains partial.",
                diagramDetails);
            var richNotesDetails = DescribeRichNotesContent();
            Add(features, "Presentation", "Speaker notes", PowerPointFeatureSupportLevel.Editable, Slides.Count(slide => slide.SlidePart.NotesSlidePart != null), null,
                "Speaker notes can be authored, inspected, updated, and preserved.");
            Add(features, "Presentation", "Rich notes content", PowerPointFeatureSupportLevel.Preserved, richNotesDetails.Count, null,
                "Notes-page drawings beyond speaker text are detected as preserve-only presentation content.",
                richNotesDetails);
            Add(features, "Presentation", "Slide transitions", PowerPointFeatureSupportLevel.Editable, Slides.Count(HasTransitionMarkup), null,
                "Common transitions, Morph fallback markup, speed, duration, advance timing, and transition sound actions can be authored and round-tripped.");
            Add(features, "Media", "Transition and action sounds",
                PowerPointFeatureSupportLevel.Editable,
                Slides.Sum(CountTransitionAndActionSounds), null,
                "Embedded sounds referenced by slide transitions and shape or text actions are detected and round-tripped.");
            var unsupportedTransitionDetails = DescribeUnsupportedTransitionMarkup();
            Add(features, "Presentation", "Unsupported transition markup", PowerPointFeatureSupportLevel.Preserved, unsupportedTransitionDetails.Count, null,
                "Transition markup not mapped by OfficeIMO is detected as preserve-only slide metadata.",
                unsupportedTransitionDetails);
            int classicAnimationCount = Slides
                .Where(slide => slide.HasOnlyClassicAnimationTiming())
                .Sum(slide => slide.ClassicAnimations.Count);
            Add(features, "Presentation", "Classic animations",
                PowerPointFeatureSupportLevel.Editable,
                classicAnimationCount, null,
                "Classic shape and text entrance effects, paragraph builds, order, automatic advance, after-effects, and sounds can be authored, inspected, and round-tripped.");
            var advancedTimingDetails = DescribeAdvancedTimingMarkup();
            Add(features, "Presentation", "Animations and timing", PowerPointFeatureSupportLevel.Preserved, advancedTimingDetails.Count, null,
                "Timing trees beyond OfficeIMO's classic-animation and media-playback helpers are detected as preserve-only advanced animation metadata.",
                advancedTimingDetails);
            Add(features, "Content", "External relationships", PowerPointFeatureSupportLevel.PartiallyEditable, externalHyperlinkDetails.Count, null,
                "External hyperlinks can be authored and inspected.",
                externalHyperlinkDetails);
            Add(features, "Compatibility", "External package relationships", PowerPointFeatureSupportLevel.Preserved, externalPackageRelationshipDetails.Count, null,
                "Linked package relationships outside hyperlink markup are detected as preserve-only package metadata.",
                externalPackageRelationshipDetails);

            var commentDetails = DescribeCommentParts(allParts);
            var customXmlDetails = DescribePartsByUri(allParts, "/customXml/");
            PowerPointOleObject[] editableOleObjects = Slides
                .SelectMany(slide => slide.OleObjects).ToArray();
            var editableOleParts = new HashSet<OpenXmlPart>(
                editableOleObjects.Select(ole =>
                    (OpenXmlPart)ole.EmbeddedPart));
            var embeddedPackageDetails = DescribeNonChartEmbeddedPackageParts(
                allParts, editableOleParts, safeChartWorkbookParts);
            var linkedOleDetails = LegacyPptLinkedOleDetails.ToList();
            var activeXControlDetails = DescribeActiveXControlParts(allParts)
                .Concat(LegacyPptActiveXDetails)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var vbaDetails = DescribeVbaProjectParts(allParts);
            var webExtensionDetails = DescribeWebExtensionParts(allParts);
            var signatureDetails = DescribeDigitalSignatureParts(allParts);
            if (_document?.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null) {
                signatureDetails.Add("Extended application properties contain digital signature metadata.");
            }

            Add(features, "Review", "Comments", PowerPointFeatureSupportLevel.Preserved, commentDetails.Count, null,
                "Comment package metadata is detected as preserve-only review content.",
                commentDetails);
            Add(features, "Compatibility", "Custom XML parts", PowerPointFeatureSupportLevel.Preserved, customXmlDetails.Count, null,
                "Custom XML parts are preserve-only package metadata.",
                customXmlDetails);
            Add(features, "Compatibility", "Embedded OLE objects",
                PowerPointFeatureSupportLevel.Editable,
                editableOleObjects.Length, null,
                "Embedded OLE compound objects expose their ProgID, display mode, color-follow setting, exact storage bytes, geometry, duplication, and removal through the normal slide model.",
                editableOleObjects.Select(ole =>
                    $"{ole.Name ?? "OLE object"}: {ole.ProgId ?? "Package"}, {ole.ContentType}").ToList());
            Add(features, "Compatibility", "Embedded packages", PowerPointFeatureSupportLevel.Preserved, embeddedPackageDetails.Count, null,
                "Unreferenced or unrecognized embedded package parts remain preserve-only package content.",
                embeddedPackageDetails);
            Add(features, "Compatibility", "Linked OLE objects",
                PowerPointFeatureSupportLevel.Preserved,
                linkedOleDetails.Count, null,
                "Binary linked OLE metadata and cached compound storage are typed and retained exactly, but are not projected to an editable Open XML object.",
                linkedOleDetails);
            Add(features, "Compatibility", "ActiveX controls", PowerPointFeatureSupportLevel.Preserved, activeXControlDetails.Count, null,
                "ActiveX metadata and native control storage are detected and retained as preserve-only advanced presentation content.",
                activeXControlDetails);
            Add(features, "Compatibility", "VBA macros", PowerPointFeatureSupportLevel.Preserved, vbaDetails.Count, null,
                "VBA project parts are detected as preserve-only macro content; OfficeIMO does not edit or sign VBA modules.",
                vbaDetails);
            Add(features, "Compatibility", "Web extensions and task panes", PowerPointFeatureSupportLevel.Preserved, webExtensionDetails.Count, null,
                "Office add-in and task-pane package metadata is detected as preserve-only advanced content.",
                webExtensionDetails);
            Add(features, "Compatibility", "Digital signatures", PowerPointFeatureSupportLevel.Unsupported, signatureDetails.Count, null,
                "Package signing and signature validation are not implemented; editing signed presentations may invalidate signatures.",
                signatureDetails);

            return new PowerPointFeatureReport(features);
        }

        private static void Add(List<PowerPointFeatureFinding> features, string category, string name, PowerPointFeatureSupportLevel supportLevel,
            int count, string? scope, string note, IReadOnlyList<string>? details = null) {
            if (count <= 0 && supportLevel != PowerPointFeatureSupportLevel.Editable) {
                return;
            }

            features.Add(new PowerPointFeatureFinding(category, name, supportLevel, count, scope, note, details));
        }

        private static IEnumerable<OpenXmlPart> EnumeratePowerPointPartsAndPackage(PresentationDocument document, OpenXmlPart root) {
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pending = new Stack<(OpenXmlPart Part, int Depth)>();
            int relationshipCount = 0;
            foreach (var pair in document.Parts) {
                relationshipCount++;
                if (relationshipCount > MaxFeatureInspectionRelationships) {
                    throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported relationship limit of {MaxFeatureInspectionRelationships}.");
                }

                pending.Push((pair.OpenXmlPart, 0));
            }

            relationshipCount++;
            if (relationshipCount > MaxFeatureInspectionRelationships) {
                throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported relationship limit of {MaxFeatureInspectionRelationships}.");
            }

            pending.Push((root, 0));

            if (document.DigitalSignatureOriginPart != null) {
                relationshipCount++;
                if (relationshipCount > MaxFeatureInspectionRelationships) {
                    throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported relationship limit of {MaxFeatureInspectionRelationships}.");
                }

                pending.Push((document.DigitalSignatureOriginPart, 0));
            }

            while (pending.Count > 0) {
                (OpenXmlPart part, int depth) = pending.Pop();
                if (depth > MaxFeatureInspectionDepth) {
                    throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported relationship depth of {MaxFeatureInspectionDepth}.");
                }

                string key = part.Uri.OriginalString + "|" + part.ContentType;
                if (!visited.Add(key)) {
                    continue;
                }

                if (visited.Count > MaxFeatureInspectionParts) {
                    throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported part limit of {MaxFeatureInspectionParts}.");
                }

                yield return part;

                foreach (IdPartPair pair in part.Parts) {
                    relationshipCount++;
                    if (relationshipCount > MaxFeatureInspectionRelationships) {
                        throw new InvalidDataException($"PowerPoint feature inspection exceeded the supported relationship limit of {MaxFeatureInspectionRelationships}.");
                    }

                    pending.Push((pair.OpenXmlPart, depth + 1));
                }
            }
        }

        private static List<string> DescribePackageFoundationParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(part => part is SlideMasterPart
                    || part is SlideLayoutPart
                    || part is ThemePart
                    || part is TableStylesPart
                    || part is NotesMasterPart
                    || part is ThumbnailPart
                    || part.Uri.OriginalString.IndexOf("/viewProps", StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> DescribeTableMetadata() {
            return Slides
                .SelectMany((slide, slideIndex) => EnumerateSlideTables(slide)
                    .Select((table, tableIndex) => DescribeTableMetadata(slideIndex, tableIndex, table))
                    ?? Enumerable.Empty<string?>())
                .Where(detail => !string.IsNullOrWhiteSpace(detail))
                .Select(detail => detail!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string? DescribeTableMetadata(int slideIndex, int tableIndex, A.Table table) {
            A.TableProperties? properties = table.TableProperties;
            string? styleId = properties?.GetFirstChild<A.TableStyleId>()?.Text;
            int columnIds = CountDescendantsByLocalName(table, "colId");
            int rowIds = CountDescendantsByLocalName(table, "rowId");
            string flags = string.Join(", ", new[] {
                    properties?.FirstRow?.Value == true ? "firstRow" : null,
                    properties?.LastRow?.Value == true ? "lastRow" : null,
                    properties?.FirstColumn?.Value == true ? "firstCol" : null,
                    properties?.LastColumn?.Value == true ? "lastCol" : null,
                    properties?.BandRow?.Value == true ? "bandRow" : null,
                    properties?.BandColumn?.Value == true ? "bandCol" : null
                }
                .Where(flag => flag != null));
            if (string.IsNullOrWhiteSpace(styleId) && string.IsNullOrWhiteSpace(flags) && columnIds == 0 && rowIds == 0) {
                return null;
            }

            return $"slide {slideIndex + 1}, table {tableIndex + 1}: style={styleId ?? "(none)"}, flags={flags}, colIds={columnIds}, rowIds={rowIds}";
        }

        private static bool HasTransitionMarkup(PowerPointSlide slide) {
            return slide.Transition != SlideTransition.None
                || slide.SlidePart.Slide?.Transition != null;
        }

        private static int CountTransitionAndActionSounds(PowerPointSlide slide) =>
            (slide.SlidePart.Slide?.Transition?
                .Descendants<DocumentFormat.OpenXml.Presentation.Sound>().Count() ?? 0)
            + (slide.SlidePart.Slide?
                .Descendants<A.HyperlinkSound>().Count() ?? 0);

        private static int CountSlideTextBoxes(PowerPointSlide slide) {
            int slideTextBoxes = slide.SlidePart.Slide?.Descendants<Shape>().Count(shape => shape.TextBody != null) ?? 0;
            return slideTextBoxes + slide.GetInheritedShapesForExport().Sum(shape => CountTextBoxElements(shape.Element));
        }

        private static int CountSlideImages(PowerPointSlide slide) {
            return CountSlidePictures(slide) + CountSlideBackgroundImages(slide) + CountSlideShapeFillImages(slide);
        }

        private static int CountSlidePictures(PowerPointSlide slide) {
            int slidePictures = slide.SlidePart.Slide?
                .Descendants<Picture>()
                .Count(picture => !PowerPointMedia.TryGetMediaKind(picture, out _))
                ?? 0;

            return slidePictures + slide.GetInheritedShapesForExport().Sum(shape => CountPictureElements(shape.Element));
        }

        private static int CountSlideBackgroundImages(PowerPointSlide slide) {
            return EnumerateResolvedBackgroundProperties(slide)
                .Any(properties => properties.GetFirstChild<A.BlipFill>()?.Blip != null)
                ? 1
                : 0;
        }

        private static IEnumerable<BackgroundProperties> EnumerateResolvedBackgroundProperties(PowerPointSlide slide) {
            Background? slideBackground = slide.SlidePart.Slide?.CommonSlideData?.Background;
            if (slideBackground != null) {
                BackgroundProperties? slideProperties = slideBackground.BackgroundProperties;
                if (slideProperties?.HasChildren == true) {
                    yield return slideProperties;
                    yield break;
                }

                if (slideBackground.BackgroundStyleReference != null) {
                    yield break;
                }
            }

            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            Background? layoutBackground = layoutPart?.SlideLayout?.CommonSlideData?.Background;
            if (layoutBackground != null) {
                BackgroundProperties? layoutProperties = layoutBackground.BackgroundProperties;
                if (layoutProperties?.HasChildren == true) {
                    yield return layoutProperties;
                    yield break;
                }

                if (layoutBackground.BackgroundStyleReference != null) {
                    yield break;
                }
            }

            BackgroundProperties? masterProperties = layoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background?.BackgroundProperties;
            if (masterProperties != null) {
                yield return masterProperties;
            }
        }

        private static int CountSlideMedia(PowerPointSlide slide) {
            int slideMedia = slide.SlidePart.Slide?
                .Descendants<Picture>()
                .Count(picture => PowerPointMedia.TryGetMediaKind(picture, out _))
                ?? 0;

            return slideMedia + slide.GetInheritedShapesForExport().Sum(shape => CountMediaElements(shape.Element));
        }

        private static int CountTextBoxElements(OpenXmlElement element) {
            int count = element is Shape shape && HasVisibleText(shape) ? 1 : 0;
            return count + element.Descendants<Shape>().Count(HasVisibleText);
        }

        private static bool HasVisibleText(Shape shape) {
            return !string.IsNullOrWhiteSpace(shape.TextBody?.InnerText);
        }

        private static int CountPictureElements(OpenXmlElement element) {
            int count = element is Picture picture && !PowerPointMedia.TryGetMediaKind(picture, out _) ? 1 : 0;
            return count + element.Descendants<Picture>().Count(picture => !PowerPointMedia.TryGetMediaKind(picture, out _));
        }

        private static int CountSlideShapeFillImages(PowerPointSlide slide) {
            int slideFillImages = CountShapeFillImageElements(slide.SlidePart.Slide);
            return slideFillImages + slide.GetInheritedShapesForExport().Sum(shape => CountShapeFillImageElements(shape.Element));
        }

        private static int CountShapeFillImageElements(OpenXmlElement? element) {
            if (element == null) {
                return 0;
            }

            return element.Descendants<A.BlipFill>()
                .Count(blipFill => blipFill.Blip != null && !blipFill.Ancestors<BackgroundProperties>().Any());
        }

        private static int CountMediaElements(OpenXmlElement element) {
            int count = element is Picture picture && PowerPointMedia.TryGetMediaKind(picture, out _) ? 1 : 0;
            return count + element.Descendants<Picture>().Count(picture => PowerPointMedia.TryGetMediaKind(picture, out _));
        }

        private static int CountSlideTables(PowerPointSlide slide) {
            int slideTables = slide.SlidePart.Slide?.Descendants<A.Table>().Count() ?? 0;
            return slideTables + slide.GetInheritedShapesForExport().Sum(shape => shape.Element.Descendants<A.Table>().Count());
        }

        private static int CountSlideSmartArt(PowerPointSlide slide) {
            return slide.SlidePart.Slide?
                .Descendants<GraphicFrame>()
                .Count(frame => frame.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>() != null)
                ?? 0;
        }

        private static IEnumerable<A.Table> EnumerateSlideTables(PowerPointSlide slide) {
            IEnumerable<A.Table> slideTables = slide.SlidePart.Slide?.Descendants<A.Table>() ?? Enumerable.Empty<A.Table>();
            IEnumerable<A.Table> inheritedTables = slide.GetInheritedShapesForExport()
                .SelectMany(shape => shape.Element.Descendants<A.Table>());
            return slideTables.Concat(inheritedTables);
        }

        private List<string> DescribeUnsupportedTransitionMarkup() {
            return Slides
                .SelectMany((slide, index) => EnumerateSlideTransitions(slide)
                    .Where(HasUnsupportedTransitionMarkup)
                    .Select(transition => $"slide {index + 1}: unsupported transition markup {transition.OuterXml}"))
                .ToList();
        }

        private static IEnumerable<Transition> EnumerateSlideTransitions(PowerPointSlide slide) {
            Slide? root = slide.SlidePart.Slide;
            if (root?.Transition != null) {
                yield return root.Transition;
            }

            if (root == null) {
                yield break;
            }

            foreach (AlternateContent alternateContent in root.Elements<AlternateContent>()) {
                foreach (AlternateContentChoice choice in alternateContent.Elements<AlternateContentChoice>()) {
                    foreach (Transition transition in choice.Elements<Transition>()) {
                        yield return transition;
                    }
                }

                AlternateContentFallback? fallback = alternateContent.GetFirstChild<AlternateContentFallback>();
                if (fallback == null) {
                    continue;
                }

                foreach (Transition transition in fallback.Elements<Transition>()) {
                    yield return transition;
                }
            }
        }

        private static bool HasUnsupportedTransitionMarkup(Transition? transition) {
            if (transition == null) {
                return false;
            }

            return transition.ChildElements.Any(element =>
                    !IsMappedTransitionEffectElement(element)
                    && element is not SoundAction)
                || transition.Elements<SoundAction>().Any(
                    sound => !IsSupportedTransitionSoundAction(sound))
                || transition.GetAttributes().Any(IsUnsupportedTransitionAttribute)
                || transition.ChildElements
                    .Where(IsMappedTransitionEffectElement)
                    .Any(element => element.GetAttributes().Any(attribute => IsUnsupportedTransitionEffectAttribute(element, attribute)));
        }

        private static bool IsSupportedTransitionSoundAction(SoundAction sound) {
            StartSoundAction[] starts = sound.Elements<StartSoundAction>().ToArray();
            EndSoundAction[] ends = sound.Elements<EndSoundAction>().ToArray();
            if (sound.ChildElements.Count != 1
                || starts.Length + ends.Length != 1) return false;
            if (ends.Length == 1) return true;
            return starts[0].ChildElements.Count == 1
                && starts[0].Elements<DocumentFormat.OpenXml.Presentation.Sound>()
                    .Count() == 1;
        }

        private static bool IsMappedTransitionEffectElement(OpenXmlElement element) {
            if (string.Equals(element.LocalName, "prstTrans", StringComparison.OrdinalIgnoreCase)) {
                return string.Equals(element.NamespaceUri, "http://schemas.microsoft.com/office/powerpoint/2012/main", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(element.GetAttribute("prst", string.Empty).Value, "morph", StringComparison.OrdinalIgnoreCase);
            }

            var supportedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "blinds",
                "checker",
                "circle",
                "comb",
                "cover",
                "cut",
                "diamond",
                "dissolve",
                "fade",
                "ferris",
                "flash",
                "morph",
                "newsflash",
                "plus",
                "prism",
                "pull",
                "push",
                "random",
                "randomBar",
                "split",
                "strips",
                "wedge",
                "wheel",
                "warp",
                "wipe",
                "zoom"
            };

            return supportedNames.Contains(element.LocalName);
        }

        private static bool IsUnsupportedTransitionAttribute(OpenXmlAttribute attribute) {
            if (IsNamespaceDeclaration(attribute)) {
                return false;
            }

            var supportedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "advClick",
                "advTm",
                "dur",
                "spd"
            };

            return !supportedNames.Contains(attribute.LocalName);
        }

        private static bool IsUnsupportedTransitionEffectAttribute(OpenXmlElement element, OpenXmlAttribute attribute) {
            if (IsNamespaceDeclaration(attribute)) {
                return false;
            }

            var supportedNames = GetSupportedTransitionEffectAttributes(element);
            return !supportedNames.Contains(attribute.LocalName);
        }

        private static HashSet<string> GetSupportedTransitionEffectAttributes(OpenXmlElement element) {
            switch (element.LocalName) {
                case "blinds":
                case "checker":
                case "comb":
                case "cover":
                case "ferris":
                case "pull":
                case "push":
                case "randomBar":
                case "strips":
                case "warp":
                case "wipe":
                case "zoom":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "dir" };
                case "split":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "dir", "orient" };
                case "cut":
                case "fade":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "thruBlk" };
                case "wheel":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "spokes" };
                case "prism":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "isContent" };
                case "morph":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "option" };
                case "prstTrans":
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "prst" };
                default:
                    return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static bool IsNamespaceDeclaration(OpenXmlAttribute attribute) {
            return attribute.Prefix == "xmlns"
                || (string.IsNullOrEmpty(attribute.Prefix) && attribute.LocalName == "xmlns");
        }

        private List<string> DescribeAdvancedTimingMarkup() {
            return Slides
                .SelectMany((slide, index) => EnumerateSlideTimingMarkup(
                        slide, index + 1)
                    .Where(item => item.Scope != "slide"
                        || !slide.HasOnlyClassicAnimationTiming()))
                .Where(item => !ContainsOnlyMediaPlaybackTiming(item.Timing, item.MediaShapeIds))
                .Select(item => $"slide {item.Index} {item.Scope}: {item.Timing.Descendants().Count()} timing descendant(s)")
                .ToList();
        }

        private static IEnumerable<(int Index, string Scope, DocumentFormat.OpenXml.Presentation.Timing Timing, HashSet<string> MediaShapeIds)> EnumerateSlideTimingMarkup(PowerPointSlide slide, int index) {
            HashSet<string> mediaShapeIds = GetSlideMediaShapeIds(slide);
            if (slide.SlidePart.Slide?.Timing != null) {
                yield return (index, "slide", slide.SlidePart.Slide.Timing, mediaShapeIds);
            }

            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            if (layoutPart?.SlideLayout?.Timing != null) {
                yield return (index, "layout", layoutPart.SlideLayout.Timing, mediaShapeIds);
            }

            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
            if (masterPart?.SlideMaster?.Timing != null) {
                yield return (index, "master", masterPart.SlideMaster.Timing, mediaShapeIds);
            }
        }

        private static bool ContainsOnlyMediaPlaybackTiming(DocumentFormat.OpenXml.Presentation.Timing timing, HashSet<string> mediaShapeIds) {
            if (mediaShapeIds.Count == 0) {
                return false;
            }

            ShapeTarget[] targets = timing
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTarget>()
                .ToArray()!;
            return ContainsOnlyMediaPlaybackTimingMarkup(timing)
                && targets.Length > 0
                && targets.All(target =>
                !string.IsNullOrWhiteSpace(target.ShapeId?.Value)
                && mediaShapeIds.Contains(target.ShapeId!.Value!)
                && IsMediaPlaybackTarget(target));
        }

        private static bool ContainsOnlyMediaPlaybackTimingMarkup(DocumentFormat.OpenXml.Presentation.Timing timing) {
            var allowedLocalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "audio",
                "cMediaNode",
                "cTn",
                "childTnLst",
                "cond",
                "par",
                "spTgt",
                "stCondLst",
                "tgtEl",
                "tnLst",
                "video"
            };

            return timing.Descendants().All(element => allowedLocalNames.Contains(element.LocalName));
        }

        private static bool IsMediaPlaybackTarget(ShapeTarget target) {
            return target.Ancestors<CommonMediaNode>().Any()
                && (target.Ancestors<Audio>().Any() || target.Ancestors<Video>().Any());
        }

        private static HashSet<string> GetSlideMediaShapeIds(PowerPointSlide slide) {
            return new HashSet<string>(
                EnumerateSlideMediaPictures(slide)
                    .Select(picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value.ToString())
                    .Where(id => !string.IsNullOrWhiteSpace(id))
                    .Select(id => id!),
                StringComparer.OrdinalIgnoreCase);
        }

        private static IEnumerable<Picture> EnumerateSlideMediaPictures(PowerPointSlide slide) {
            IEnumerable<Picture> slideMedia = slide.SlidePart.Slide?
                .Descendants<Picture>()
                .Where(picture => PowerPointMedia.TryGetMediaKind(picture, out _))
                ?? Enumerable.Empty<Picture>();
            IEnumerable<Picture> inheritedMedia = slide.GetInheritedShapesForExport()
                .SelectMany(shape => EnumerateMediaPictures(shape.Element));
            return slideMedia.Concat(inheritedMedia);
        }

        private static IEnumerable<Picture> EnumerateMediaPictures(OpenXmlElement element) {
            if (element is Picture picture && PowerPointMedia.TryGetMediaKind(picture, out _)) {
                yield return picture;
            }

            foreach (Picture descendant in element.Descendants<Picture>().Where(picture => PowerPointMedia.TryGetMediaKind(picture, out _))) {
                yield return descendant;
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

        private static List<string> DescribeVbaProjectParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(part => string.Equals(part.ContentType, "application/vnd.ms-office.vbaProject", StringComparison.OrdinalIgnoreCase))
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeActiveXControlParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(part =>
                    string.Equals(part.ContentType, "application/vnd.ms-office.activeX+xml", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(part.ContentType, "application/vnd.ms-office.activeX", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(part.ContentType, "application/vnd.ms-office.activeX.bin", StringComparison.OrdinalIgnoreCase))
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeWebExtensionParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(part =>
                    string.Equals(part.ContentType, "application/vnd.ms-office.webextension+xml", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(part.ContentType, "application/vnd.ms-office.webextensiontaskpanes+xml", StringComparison.OrdinalIgnoreCase))
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private List<string> DescribeRichNotesContent() {
            List<string> details = Slides
                .SelectMany((slide, index) => DescribeRichNotesContent(slide, index + 1))
                .ToList();
            details.AddRange(DescribeRichNotesContent(_presentationPart.NotesMasterPart?.NotesMaster, "notes master"));
            return details;
        }

        private static IEnumerable<string> DescribeRichNotesContent(PowerPointSlide slide, int slideIndex) {
            NotesSlide? notesSlide = slide.SlidePart.NotesSlidePart?.NotesSlide;
            return DescribeRichNotesContent(notesSlide, $"slide {slideIndex} notes");
        }

        private static IEnumerable<string> DescribeRichNotesContent(OpenXmlElement? notesRoot, string scope) {
            if (notesRoot == null) {
                yield break;
            }

            int pictures = notesRoot.Descendants<Picture>()
                .Count(picture => !PowerPointMedia.TryGetMediaKind(picture, out _));
            int shapeFillImages = CountShapeFillImageElements(notesRoot);
            int tables = notesRoot.Descendants<A.Table>().Count();
            int charts = notesRoot
                .Descendants<GraphicFrame>()
                .Count(frame => frame.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>() != null);
            int smartArt = notesRoot
                .Descendants<GraphicFrame>()
                .Count(frame => frame.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>() != null);
            int textShapes = notesRoot
                .Descendants<Shape>()
                .Count(shape => HasNonEmptyShapeText(shape) && !HasPlaceholder(shape));

            if (pictures > 0) {
                yield return $"{scope}: {pictures} picture(s)";
            }

            if (shapeFillImages > 0) {
                yield return $"{scope}: {shapeFillImages} shape fill image(s)";
            }

            if (tables > 0) {
                yield return $"{scope}: {tables} table(s)";
            }

            if (charts > 0) {
                yield return $"{scope}: {charts} chart(s)";
            }

            if (smartArt > 0) {
                yield return $"{scope}: {smartArt} SmartArt diagram(s)";
            }

            if (textShapes > 0) {
                yield return $"{scope}: {textShapes} extra text shape(s)";
            }
        }

        private static bool HasNonEmptyShapeText(Shape shape) {
            return !string.IsNullOrWhiteSpace(shape.TextBody?.InnerText);
        }

        private static bool HasPlaceholder(Shape shape) {
            return shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>() != null;
        }

        private static List<string> DescribeDigitalSignatureParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(part =>
                    part.Uri.OriginalString.IndexOf("/_xmlsignatures/", StringComparison.OrdinalIgnoreCase) >= 0
                    || part.ContentType.IndexOf("digital-signature", StringComparison.OrdinalIgnoreCase) >= 0
                    || part.ContentType.IndexOf("xmlsignature", StringComparison.OrdinalIgnoreCase) >= 0)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeCommentParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .Where(IsCommentPart)
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsCommentPart(OpenXmlPart part) {
            string uri = part.Uri.OriginalString;
            string contentType = part.ContentType;
            return uri.IndexOf("/ppt/comments/", StringComparison.OrdinalIgnoreCase) >= 0
                || uri.IndexOf("/ppt/commentAuthors", StringComparison.OrdinalIgnoreCase) >= 0
                || uri.IndexOf("/ppt/persons/", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("presentationml.comments", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("presentationml.commentAuthors", StringComparison.OrdinalIgnoreCase) >= 0
                || contentType.IndexOf("presentationml.person", StringComparison.OrdinalIgnoreCase) >= 0;
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

        private static List<string> DescribeChartWorkbookParts(ISet<OpenXmlPart> safeChartWorkbookParts) {
            return safeChartWorkbookParts
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeNonChartEmbeddedPackageParts(
            IEnumerable<OpenXmlPart> parts,
            ISet<OpenXmlPart> editableOleParts,
            ISet<OpenXmlPart> safeChartWorkbookParts) {
            return parts
                .Where(part => part is EmbeddedPackagePart || part is EmbeddedObjectPart)
                .Where(part => !safeChartWorkbookParts.Contains(part))
                .Where(part => !editableOleParts.Contains(part))
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static HashSet<OpenXmlPart> GetSafeChartWorkbookParts(IEnumerable<OpenXmlPart> parts) {
            var result = new HashSet<OpenXmlPart>();
            foreach (ChartPart chartPart in parts.OfType<ChartPart>()) {
                foreach (EmbeddedPackagePart packagePart in chartPart.GetPartsOfType<EmbeddedPackagePart>()) {
                    if (IsSafeChartWorkbookPart(chartPart, packagePart)) {
                        result.Add(packagePart);
                    }
                }
            }

            return result;
        }

        private static bool IsSafeChartWorkbookPart(ChartPart chartPart, EmbeddedPackagePart part) {
            if (!string.Equals(
                    part.ContentType,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string relationshipId = chartPart.GetIdOfPart(part);
            bool isReferenced = chartPart.ChartSpace?
                .Descendants<C.ExternalData>()
                .Any(externalData => string.Equals(externalData.Id?.Value, relationshipId, StringComparison.Ordinal)) == true;
            if (!isReferenced) {
                return false;
            }

            try {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                byte[] workbookBytes = ReadChartWorkbookBytes(stream);
                OfficePackageSecurityInspector.Validate(workbookBytes, CreateChartWorkbookSecurityOptions());
                using var workbookStream = new MemoryStream(workbookBytes, writable: false);
                using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookStream, false);
                return IsSafeGeneratedChartWorkbook(workbook);
            } catch (FileFormatException) {
                return false;
            } catch (OpenXmlPackageException) {
                return false;
            } catch (InvalidDataException) {
                return false;
            } catch (IOException) {
                return false;
            }
        }

        private static OfficePackageSecurityOptions CreateChartWorkbookSecurityOptions() =>
            new OfficePackageSecurityOptions {
                MaxPackageBytes = 8L * 1024L * 1024L,
                MaxPartCount = 64,
                MaxPartUncompressedBytes = 2L * 1024L * 1024L,
                MaxTotalUncompressedBytes = 8L * 1024L * 1024L,
                MaxCompressionRatio = 100D,
                Macros = OfficePackageContentPolicy.Reject,
                EmbeddedPayloads = OfficePackageContentPolicy.Reject,
                ActiveX = OfficePackageContentPolicy.Reject,
                ExternalRelationships = OfficePackageContentPolicy.Reject
            };

        private static byte[] ReadChartWorkbookBytes(Stream stream) {
            const int maximumBytes = 8 * 1024 * 1024;
            using var buffer = new MemoryStream();
            var chunk = new byte[81920];
            int totalBytes = 0;
            while (true) {
                int read = stream.Read(chunk, 0, Math.Min(chunk.Length, maximumBytes + 1 - totalBytes));
                if (read == 0) {
                    return buffer.ToArray();
                }

                totalBytes = checked(totalBytes + read);
                if (totalBytes > maximumBytes) {
                    throw new InvalidDataException(
                        $"Embedded chart workbook exceeds the configured maximum of {maximumBytes} bytes.");
                }

                buffer.Write(chunk, 0, read);
            }
        }

        private static bool IsSafeGeneratedChartWorkbook(SpreadsheetDocument workbook) {
            WorkbookPart? workbookPart = workbook.WorkbookPart;
            if (workbook.DocumentType != SpreadsheetDocumentType.Workbook
                || workbookPart?.Workbook == null
                || workbookPart.VbaProjectPart != null
                || workbook.Parts.Count() != 1
                || workbook.Parts.Any(pair => pair.OpenXmlPart is not WorkbookPart)
                || workbook.ExternalRelationships.Any()) {
                return false;
            }

            WorksheetPart[] worksheets = workbookPart.GetPartsOfType<WorksheetPart>().ToArray();
            SharedStringTablePart[] sharedStrings = workbookPart.GetPartsOfType<SharedStringTablePart>().ToArray();
            if (worksheets.Length != 1
                || sharedStrings.Length != 1
                || workbookPart.Parts.Count() != 2
                || workbookPart.Parts.Any(pair => pair.OpenXmlPart is not WorksheetPart && pair.OpenXmlPart is not SharedStringTablePart)
                || HasUnsupportedChartWorkbookRelationships(workbookPart)) {
                return false;
            }

            WorksheetPart worksheetPart = worksheets[0];
            S.Worksheet? worksheet = worksheetPart.Worksheet;
            if (worksheet == null
                || worksheetPart.Parts.Any()
                || sharedStrings[0].Parts.Any()
                || HasUnsupportedChartWorkbookRelationships(worksheetPart)
                || HasUnsupportedChartWorkbookRelationships(sharedStrings[0])
                || worksheet.Descendants<S.CellFormula>().Any()
                || worksheet.Descendants<S.Hyperlinks>().Any()
                || worksheet.Descendants<S.OleObjects>().Any()
                || worksheet.Descendants<S.Controls>().Any()) {
                return false;
            }

            S.Sheets? sheets = workbookPart.Workbook.Sheets;
            S.Sheet? sheet = sheets?.Elements<S.Sheet>().SingleOrDefault();
            return sheet?.Id?.Value != null
                && string.Equals(sheet.Id.Value, workbookPart.GetIdOfPart(worksheetPart), StringComparison.Ordinal)
                && workbookPart.Workbook.DefinedNames == null
                && workbookPart.Workbook.ExternalReferences == null;
        }

        private static bool HasUnsupportedChartWorkbookRelationships(OpenXmlPart part) =>
            part.ExternalRelationships.Any()
            || part.HyperlinkRelationships.Any()
            || part.DataPartReferenceRelationships.Any();

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

        private static List<string> DescribeExternalHyperlinkRelationships(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part =>
                    part.HyperlinkRelationships
                        .Where(relationship => relationship.IsExternal)
                        .Where(relationship => IsRelationshipReferenced(part, relationship.Id))
                        .Select(relationship => $"{part.Uri}: {relationship.Id} -> {relationship.Uri}"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsRelationshipReferenced(OpenXmlPart part, string relationshipId) {
            OpenXmlElement? root = part.RootElement;
            if (root == null) {
                return false;
            }

            return EnumerateSelfAndDescendants(root)
                .Any(element => element.GetAttributes().Any(attribute => IsRelationshipReference(attribute, relationshipId)));
        }

        private static IEnumerable<OpenXmlElement> EnumerateSelfAndDescendants(OpenXmlElement root) {
            yield return root;
            foreach (OpenXmlElement descendant in root.Descendants()) {
                yield return descendant;
            }
        }

        private static bool IsRelationshipReference(OpenXmlAttribute attribute, string relationshipId) {
            return string.Equals(attribute.LocalName, "id", StringComparison.OrdinalIgnoreCase)
                && string.Equals(attribute.NamespaceUri, "http://schemas.openxmlformats.org/officeDocument/2006/relationships", StringComparison.OrdinalIgnoreCase)
                && string.Equals(attribute.Value, relationshipId, StringComparison.OrdinalIgnoreCase);
        }

        private static List<string> DescribeExternalPackageRelationships(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part =>
                    part.ExternalRelationships.Select(relationship =>
                        $"{part.Uri}: {relationship.Id} ({relationship.RelationshipType}) -> {relationship.Uri}"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string DescribePart(OpenXmlPart part) {
            return $"{part.Uri} ({part.ContentType})";
        }

        private static int CountDescendantsByLocalName(OpenXmlElement? root, string localName) {
            if (root == null) return 0;
            return root.Descendants().Count(element => string.Equals(element.LocalName, localName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
