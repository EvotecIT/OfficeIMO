using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

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
            var matches = FindFeatures(featureNames);
            if (matches.Count > 0) {
                ThrowBlockedFeatures("Presentation contains blocked features", matches);
            }

            return this;
        }

        /// <summary>
        /// Throws when the presentation contains any features with the provided support levels.
        /// </summary>
        /// <param name="supportLevels">Support levels to reject.</param>
        public PowerPointFeatureReport EnsureNoFeatures(params PowerPointFeatureSupportLevel[] supportLevels) {
            var matches = FindFeatures(supportLevels);
            if (matches.Count > 0) {
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
            var chartWorkbookDetails = DescribeChartWorkbookParts(allParts);
            var chartCompanionDetails = DescribePartsByUriOrContentType(allParts, "chartStyle")
                .Concat(DescribePartsByUriOrContentType(allParts, "chartColorStyle"))
                .Concat(chartWorkbookDetails)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var externalRelationshipDetails = DescribeExternalRelationships(allParts);

            Add(features, "Structure", "Slides", PowerPointFeatureSupportLevel.Editable, Slides.Count, null,
                "Slides can be authored, duplicated, imported, hidden, reordered, and removed.");
            Add(features, "Structure", "PowerPoint package scaffolding", PowerPointFeatureSupportLevel.Editable, packageFoundationDetails.Count, null,
                "Presentation, layout, master, theme, view, table-style, notes-master, and thumbnail parts are created or preserved as package foundation.",
                packageFoundationDetails);
            Add(features, "Structure", "Sections", PowerPointFeatureSupportLevel.Editable, GetSections().Count, null,
                "Sections can be authored, inspected, renamed, moved, and synchronized with slides.");
            Add(features, "Content", "Text boxes", PowerPointFeatureSupportLevel.Editable, Slides.Sum(slide => slide.TextBoxes.Count()), null,
                "Text boxes, runs, common formatting, markdown import, hyperlinks, and replacement are editable.");
            Add(features, "Content", "Tables", PowerPointFeatureSupportLevel.Editable, Slides.Sum(slide => slide.Tables.Count()), null,
                "Tables, cells, text, merges, dimensions, fills, borders, alignment, banding flags, and typed object binding are editable.");
            Add(features, "Content", "Table style metadata", PowerPointFeatureSupportLevel.Editable, tableMetadataDetails.Count, null,
                "Generated tables include PowerPoint-style table IDs, banding metadata, row and column IDs, and language-aware cell text defaults.",
                tableMetadataDetails);
            Add(features, "Visualization", "Charts", PowerPointFeatureSupportLevel.PartiallyEditable, Math.Max(Slides.Sum(slide => slide.Charts.Count()), chartDetails.Count), null,
                "Common chart authoring and data updates are supported; advanced chart editing remains partial.",
                chartDetails.Concat(chartCompanionDetails).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
            Add(features, "Media", "Images", PowerPointFeatureSupportLevel.PartiallyEditable, Slides.Sum(slide => slide.Pictures.Count()), null,
                "Images can be inserted and inspected in common slide scenarios; advanced drawing behaviors remain partial.");
            Add(features, "Media", "Audio and video", PowerPointFeatureSupportLevel.PartiallyEditable, Slides.Sum(slide => slide.Media.Count()), null,
                "Embedded audio and video can be authored with poster frames and playback timing; rich media editing remains partial.",
                DescribePartsByUriOrContentType(allParts, "media"));
            Add(features, "Visualization", "SmartArt", PowerPointFeatureSupportLevel.PartiallyEditable, Slides.Sum(slide => slide.SmartArts.Count()), null,
                "SmartArt diagrams can be generated and discovered; rich diagram editing remains partial.",
                DescribeDiagramParts(allParts));
            Add(features, "Presentation", "Speaker notes", PowerPointFeatureSupportLevel.Editable, Slides.Count(slide => slide.SlidePart.NotesSlidePart != null), null,
                "Speaker notes can be authored, inspected, updated, and preserved.");
            Add(features, "Presentation", "Slide transitions", PowerPointFeatureSupportLevel.Editable, Slides.Count(HasTransitionMarkup), null,
                "Common transitions, Morph fallback markup, speed, duration, and advance timing can be authored and round-tripped.");
            var unsupportedTransitionDetails = DescribeUnsupportedTransitionMarkup();
            Add(features, "Presentation", "Unsupported transition markup", PowerPointFeatureSupportLevel.Preserved, unsupportedTransitionDetails.Count, null,
                "Transition markup not mapped by OfficeIMO is detected as preserve-only slide metadata.",
                unsupportedTransitionDetails);
            var advancedTimingDetails = DescribeAdvancedTimingMarkup();
            Add(features, "Presentation", "Animations and timing", PowerPointFeatureSupportLevel.Preserved, advancedTimingDetails.Count, null,
                "Timing trees beyond OfficeIMO's media playback helpers are detected as preserve-only advanced animation metadata.",
                advancedTimingDetails);
            Add(features, "Content", "External relationships", PowerPointFeatureSupportLevel.PartiallyEditable, externalRelationshipDetails.Count, null,
                "External hyperlinks can be authored and inspected; other external package relationships are surfaced for round-trip review.",
                externalRelationshipDetails);

            var commentDetails = DescribePartsByUriOrContentType(allParts, "comment")
                .Concat(DescribePartsByUriOrContentType(allParts, "person"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var customXmlDetails = DescribePartsByUri(allParts, "/customXml/");
            var embeddedPackageDetails = DescribeNonChartEmbeddedPackageParts(allParts);
            var vbaDetails = DescribePartsByUriOrContentType(allParts, "vbaProject");
            var webExtensionDetails = DescribePartsByUriOrContentType(allParts, "webextension")
                .Concat(DescribePartsByUriOrContentType(allParts, "taskpane"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var signatureDetails = DescribePartsByUriOrContentType(allParts, "signature")
                .Concat(DescribePartsByUriOrContentType(allParts, "xmlsignatures"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Add(features, "Review", "Comments", PowerPointFeatureSupportLevel.Preserved, commentDetails.Count, null,
                "Comment package metadata is detected as preserve-only review content.",
                commentDetails);
            Add(features, "Compatibility", "Custom XML parts", PowerPointFeatureSupportLevel.Preserved, customXmlDetails.Count, null,
                "Custom XML parts are preserve-only package metadata.",
                customXmlDetails);
            Add(features, "Compatibility", "Embedded packages", PowerPointFeatureSupportLevel.Preserved, embeddedPackageDetails.Count, null,
                "Embedded package parts and OLE payloads are advanced package content and should be treated as preserve-only.",
                embeddedPackageDetails);
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
            foreach (var pair in document.Parts) {
                foreach (var part in EnumeratePowerPointParts(pair.OpenXmlPart, visited)) {
                    yield return part;
                }
            }

            foreach (var part in EnumeratePowerPointParts(root, visited)) {
                yield return part;
            }

            if (document.DigitalSignatureOriginPart != null) {
                foreach (var part in EnumeratePowerPointParts(document.DigitalSignatureOriginPart, visited)) {
                    yield return part;
                }
            }
        }

        private static IEnumerable<OpenXmlPart> EnumeratePowerPointParts(OpenXmlPart part, HashSet<string> visited) {
            string key = part.Uri.OriginalString + "|" + part.ContentType;
            if (!visited.Add(key)) {
                yield break;
            }

            yield return part;

            foreach (var child in EnumeratePowerPointParts((OpenXmlPartContainer)part, visited)) {
                yield return child;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumeratePowerPointParts(OpenXmlPartContainer container, HashSet<string> visited) {
            foreach (var pair in container.Parts) {
                var part = pair.OpenXmlPart;
                foreach (var child in EnumeratePowerPointParts(part, visited)) {
                    yield return child;
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
                .SelectMany((slide, slideIndex) => slide.SlidePart.Slide?.Descendants<A.Table>()
                    .Select((table, tableIndex) => DescribeTableMetadata(slideIndex, tableIndex, table))
                    ?? Enumerable.Empty<string>())
                .Where(detail => !string.IsNullOrWhiteSpace(detail))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string DescribeTableMetadata(int slideIndex, int tableIndex, A.Table table) {
            A.TableProperties? properties = table.TableProperties;
            string styleId = properties?.GetFirstChild<A.TableStyleId>()?.Text ?? "(none)";
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
            return $"slide {slideIndex + 1}, table {tableIndex + 1}: style={styleId}, flags={flags}, colIds={columnIds}, rowIds={rowIds}";
        }

        private static bool HasTransitionMarkup(PowerPointSlide slide) {
            return slide.Transition != SlideTransition.None;
        }

        private List<string> DescribeUnsupportedTransitionMarkup() {
            return Slides
                .Select((slide, index) => new {
                    Index = index + 1,
                    Transition = slide.SlidePart.Slide?.Transition,
                    MappedTransition = slide.Transition
                })
                .Where(item => item.Transition != null && item.MappedTransition == SlideTransition.None)
                .Select(item => $"slide {item.Index}: unsupported transition markup {item.Transition!.OuterXml}")
                .ToList();
        }

        private List<string> DescribeAdvancedTimingMarkup() {
            return Slides
                .Select((slide, index) => new {
                    Index = index + 1,
                    Timing = slide.SlidePart.Slide?.Timing,
                    MediaShapeIds = new HashSet<string>(
                        slide.Media
                            .Select(media => media.Id?.ToString())
                            .Where(id => !string.IsNullOrWhiteSpace(id))!,
                        StringComparer.OrdinalIgnoreCase)
                })
                .Where(item => item.Timing != null)
                .Where(item => !ContainsOnlyMediaPlaybackTiming(item.Timing!, item.MediaShapeIds))
                .Select(item => $"slide {item.Index}: {item.Timing!.Descendants().Count()} timing descendant(s)")
                .ToList();
        }

        private static bool ContainsOnlyMediaPlaybackTiming(DocumentFormat.OpenXml.Presentation.Timing timing, HashSet<string> mediaShapeIds) {
            if (mediaShapeIds.Count == 0) {
                return false;
            }

            ShapeTarget[] targets = timing
                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTarget>()
                .ToArray()!;
            return targets.Length > 0 && targets.All(target =>
                !string.IsNullOrWhiteSpace(target.ShapeId?.Value)
                && mediaShapeIds.Contains(target.ShapeId!.Value!)
                && IsMediaPlaybackTarget(target));
        }

        private static bool IsMediaPlaybackTarget(ShapeTarget target) {
            return target.Ancestors<CommonMediaNode>().Any()
                && (target.Ancestors<Audio>().Any() || target.Ancestors<Video>().Any());
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

        private static List<string> DescribeChartWorkbookParts(IEnumerable<OpenXmlPart> parts) {
            return parts
                .OfType<ChartPart>()
                .SelectMany(chartPart => chartPart.GetPartsOfType<EmbeddedPackagePart>())
                .Select(DescribePart)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(detail => detail, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> DescribeNonChartEmbeddedPackageParts(IEnumerable<OpenXmlPart> parts) {
            var chartWorkbooks = new HashSet<OpenXmlPart>(
                parts.OfType<ChartPart>().SelectMany(chartPart => chartPart.GetPartsOfType<EmbeddedPackagePart>()));

            return parts
                .Where(part => part is EmbeddedPackagePart || part is EmbeddedObjectPart)
                .Where(part => !chartWorkbooks.Contains(part))
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

        private static List<string> DescribeExternalRelationships(IEnumerable<OpenXmlPart> parts) {
            return parts
                .SelectMany(part =>
                    part.ExternalRelationships.Select(relationship =>
                        $"{part.Uri}: {relationship.Id} ({relationship.RelationshipType}) -> {relationship.Uri}")
                    .Concat(part.HyperlinkRelationships
                        .Where(relationship => relationship.IsExternal)
                        .Select(relationship => $"{part.Uri}: {relationship.Id} -> {relationship.Uri}")))
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
