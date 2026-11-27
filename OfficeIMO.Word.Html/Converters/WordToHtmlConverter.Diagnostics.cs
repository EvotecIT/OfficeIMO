using OfficeIMO.Html;
using System.Text;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static ExportInspection InspectExport(WordDocument document, WordHtmlExportOptions options) {
            HtmlConversionProfileContracts.Get(options.SharedProfile);
            options.DocumentOutput.Validate();
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxDocumentElements, nameof(options.MaxDocumentElements));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxEmbeddedImageBytes, nameof(options.MaxEmbeddedImageBytes));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxTotalEmbeddedImageBytes, nameof(options.MaxTotalEmbeddedImageBytes));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxOutputCharacters, nameof(options.MaxOutputCharacters));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxEquationNestingDepth, nameof(options.MaxEquationNestingDepth));
            if (options.MaxEquationNestingDepth > WordMath.DefaultMaximumProjectionDepth) {
                throw new ArgumentOutOfRangeException(
                    nameof(options.MaxEquationNestingDepth),
                    "MaxEquationNestingDepth cannot exceed the supported OMML projection depth of " +
                    WordMath.DefaultMaximumProjectionDepth + ".");
            }
            if (!Enum.IsDefined(typeof(WordTrackedChangeExportPolicy), options.TrackedChangePolicy)) {
                throw new ArgumentOutOfRangeException(nameof(options.TrackedChangePolicy));
            }
            if (!Enum.IsDefined(typeof(WordFieldExportPolicy), options.FieldPolicy)) {
                throw new ArgumentOutOfRangeException(nameof(options.FieldPolicy));
            }

            long elements = 0;
            bool hasFields = false;
            bool hasRevisions = false;
            bool hasComments = false;
            long outputConstructionCharacters = 0;
            foreach ((OfficeXmlNode Root, string PartUri, long StreamedElements) root in document.HtmlExportRoots(
                options.IncludeCustomProperties)) {
                if (root.Root.IsNone) {
                    // A styles part that was never loaded is counted from its stream; it holds no exported content.
                    elements = SaturatingAdd(elements, root.StreamedElements);
                    if (elements > options.MaxDocumentElements) {
                        ThrowExportLimitExceeded(options, "word-element-limit-exceeded", "The Word document exceeds the configured HTML export element limit.", root.PartUri, options.MaxDocumentElements + 1, options.MaxDocumentElements);
                    }
                    continue;
                }
                bool countOutputContent = IsOutputContentRoot(root.Root, options);
                bool inspectExportedStory = IsSelectedStoryRoot(root.Root, options);
                foreach ((OfficeXmlNode Element, bool OmitOutputContent) inspected in EnumerateRootAndDescendants(
                             root.Root,
                             options.TrackedChangePolicy)) {
                    OfficeXmlNode element = inspected.Element;
                    elements++;
                    if (elements > options.MaxDocumentElements) {
                        ThrowExportLimitExceeded(options, "word-element-limit-exceeded", "The Word document exceeds the configured HTML export element limit.", root.PartUri, elements, options.MaxDocumentElements);
                    }
                    if (countOutputContent && !inspected.OmitOutputContent) {
                        long elementCharacters = GetOutputContentCharacters(element);
                        outputConstructionCharacters = SaturatingAdd(outputConstructionCharacters, elementCharacters);
                        if (outputConstructionCharacters > options.MaxOutputCharacters) {
                            ThrowExportLimitExceeded(
                                options,
                                "word-html-output-limit-exceeded",
                                "Word text and attribute content exceeds the configured HTML output-character limit before DOM construction.",
                                root.PartUri,
                                outputConstructionCharacters,
                                options.MaxOutputCharacters);
                        }
                    }
                    if (!hasFields && inspectExportedStory && !inspected.OmitOutputContent &&
                        (element.IsW("fldSimple") || element.IsW("fldChar") || element.IsW("instrText")) && element.IsModelElement()) {
                        hasFields = true;
                    }
                    if (!hasRevisions && inspectExportedStory && IsRevisionElement(element.LocalName)) {
                        hasRevisions = true;
                    }
                    if (!hasComments && element.IsW("comment") && element.Parent.IsW("comments") && element.IsModelElement()) {
                        hasComments = true;
                    }
                }
            }
            return new ExportInspection(
                hasFields,
                hasRevisions,
                hasComments,
                outputConstructionCharacters);
        }

        private static long SaturatingAdd(long left, long right) =>
            left > long.MaxValue - right ? long.MaxValue : left + right;

        private static long GetHtmlEncodedLength(string? value, bool attributeValue) {
            if (value == null || value.Length == 0) return 0;

            long length = 0;
            foreach (char character in value) {
                length = SaturatingAdd(length, character switch {
                    '&' => 5,
                    '<' => 4,
                    '>' => 4,
                    '"' when attributeValue => 6,
                    _ => 1
                });
            }
            return length;
        }

        private static long MeasureOutputContentCharacters(OfficeXmlNode root,
            WordTrackedChangeExportPolicy policy = WordTrackedChangeExportPolicy.Markup) {
            long characters = 0;
            foreach ((OfficeXmlNode Element, bool OmitOutputContent) inspected in EnumerateRootAndDescendants(
                         root,
                         policy)) {
                if (inspected.OmitOutputContent) continue;
                characters = SaturatingAdd(characters, GetOutputContentCharacters(inspected.Element));
            }
            return characters;
        }

        private static long GetOutputContentCharacters(OfficeXmlNode element) {
            long characters = element.LeafText() is { } text && ShouldCountOutputLeafText(element)
                ? GetHtmlEncodedLength(text, attributeValue: false)
                : 0;
            if (IsVisibleRunArtifact(element)) {
                characters = SaturatingAdd(characters, 1);
            }
            return characters;
        }

        // A w:tab in a run is a tab character; in w:tabs it is a tab stop. Untyped markup is no run content.
        private static bool IsVisibleRunArtifact(OfficeXmlNode element) =>
            element.NamespaceId == WordXml.W && element.IsModelElement() && element.LocalName switch {
                "tab" => !element.Parent.IsW("tabs"),
                "cr" or "br" or "noBreakHyphen" or "softHyphen" => true,
                _ => false
            };

        private static bool ShouldCountOutputLeafText(OfficeXmlNode element) =>
            !((element.IsW("instrText") || element.Is(WordXml.M, "t")) && element.IsModelElement());

        // Header and footer parts are intentionally excluded here. A shared part can be emitted
        // once per section, so AppendHeaderFooterRegion measures and reserves its source content
        // immediately before every emitted region rather than charging the package part once.
        private static bool IsOutputContentRoot(OfficeXmlNode root, WordHtmlExportOptions options) {
            if (root.NamespaceId == WordXml.W) {
                switch (root.LocalName) {
                    case "hdr": case "ftr": case "footnotes": case "endnotes": case "comments": case "styles":
                    case "numbering": case "fonts": case "settings": case "webSettings":
                        return false;
                }
            }
            return !root.Is("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties", "Properties") &&
                !root.Is(WordXml.A, "theme") &&
                !root.Is(WordXml.C, "chartSpace") &&
                root.NamespaceUri != "http://schemas.microsoft.com/office/drawing/2014/chartex";
        }

        private static IEnumerable<(OfficeXmlNode Element, bool OmitOutputContent)> EnumerateRootAndDescendants(
            OfficeXmlNode root,
            WordTrackedChangeExportPolicy policy) {
            var pending = new Stack<(OfficeXmlNode Element, bool InOmittedRevision)>();
            pending.Push((root, false));
            while (pending.Count > 0) {
                (OfficeXmlNode element, bool inOmittedRevision) = pending.Pop();
                bool omitOutputContent = inOmittedRevision || IsRevisionContentExcluded(element, policy);
                yield return (element, omitOutputContent);
                // Retain one sibling continuation per depth rather than allocating a
                // stack entry for every child before the element budget can stop traversal.
                if (!element.Equals(root) && element.NextSibling is { Exists: true } sibling)
                    pending.Push((sibling, inOmittedRevision));
                if (element.FirstChild is { Exists: true } child)
                    pending.Push((child, omitOutputContent));
            }
        }

        private static bool IsRevisionContentExcluded(OfficeXmlNode element, WordTrackedChangeExportPolicy policy) {
            if (element.LocalName is not ("ins" or "del" or "moveFrom" or "moveTo")) return false;
            bool insertedView = element.LocalName is "ins" or "moveTo";
            return policy switch {
                WordTrackedChangeExportPolicy.Final => !insertedView,
                WordTrackedChangeExportPolicy.Original => insertedView,
                WordTrackedChangeExportPolicy.Markup => false,
                _ => throw new ArgumentOutOfRangeException(nameof(policy), policy, "Word tracked-change policy is not supported.")
            };
        }

        private static bool IsSelectedStoryRoot(OfficeXmlNode root, WordHtmlExportOptions options) =>
            root.NamespaceId == WordXml.W && root.LocalName switch {
                "body" => true,
                "hdr" or "ftr" => options.ExportHeadersAndFooters,
                "footnotes" => options.ExportFootnotes,
                "endnotes" => options.ExportEndnotes,
                "comments" => options.ExportComments,
                _ => false
            };

        private static void ReportKnownExportLimitations(WordDocument document, WordHtmlExportOptions options, ExportInspection inspection) {
            Dictionary<WordParagraph, WordDocumentTraversal.ResolvedListMarker> resolvedMarkers =
                WordDocumentTraversal.BuildResolvedListMarkers(document);
            IEnumerable<WordDocumentTraversal.ResolvedListMarker> renderedMarkers = options.ExportHeadersAndFooters
                ? resolvedMarkers.Values
                : resolvedMarkers.Where(pair => pair.Key.Node().Ancestors().Any(ancestor => ancestor.IsW("body"))).Select(pair => pair.Value);
            foreach (int pictureBulletId in WordDocumentTraversal.GetPictureBulletFallbackIds(renderedMarkers)) {
                AddExportDiagnostic(options, "picture-bullet-text-fallback",
                    "Picture bullet " + pictureBulletId + " is represented by a text bullet in HTML output.",
                    OfficeConversionLossKind.Approximation);
            }
            if (!options.EmitDocumentShell &&
                (options.IncludeCustomProperties || options.AdditionalMetaTags.Count > 0 || options.AdditionalLinkTags.Count > 0)) {
                AddExportDiagnostic(options, "document-head-metadata-omitted-for-fragment",
                    "Document-head metadata and links were omitted because fragment output was requested.",
                    OfficeConversionLossKind.Omission);
            }
            if (inspection.HasRevisions) {
                string policy = options.TrackedChangePolicy.ToString();
                AddExportDiagnostic(options, "tracked-revisions-projected",
                    "Tracked revisions were projected through the explicit " + policy + " static HTML policy; live accept/reject behavior remains in Word.",
                    OfficeConversionLossKind.Approximation);
            }
            if (inspection.HasComments && !options.ExportComments) {
                AddExportDiagnostic(options, "comments-omitted", "Word comments were omitted because ExportComments is false.", OfficeConversionLossKind.Omission);
            }
            if (inspection.HasFields) {
                AddExportDiagnostic(options, "field-instructions-flattened",
                    options.FieldPolicy == WordFieldExportPolicy.VisibleResultWithReviewMetadata
                        ? "Word fields use their stored visible results and inert instruction metadata; live field evaluation is not represented in HTML."
                        : "Word field instructions are exported through their visible results; live field behavior is not represented in HTML.",
                    OfficeConversionLossKind.Approximation);
            }
            if (document.HasMacros) {
                AddExportDiagnostic(options, "macro-project-omitted", "The VBA project is package metadata and is not represented in HTML.", OfficeConversionLossKind.Omission);
            }
            if (document.HasDigitalSignatureOriginPart() ||
                document.ApplicationProperties.HasDigitalSignatureMetadata) {
                AddExportDiagnostic(options, "package-signatures-omitted", "OPC package signature metadata is not represented in HTML.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportHeadersAndFooters && document.Sections.Any(section =>
                section.PageSetup.GetHeader(WordHeaderFooterType.Default) != null || section.PageSetup.GetHeader(WordHeaderFooterType.Even) != null || section.PageSetup.GetHeader(WordHeaderFooterType.First) != null ||
                section.PageSetup.GetFooter(WordHeaderFooterType.Default) != null || section.PageSetup.GetFooter(WordHeaderFooterType.Even) != null || section.PageSetup.GetFooter(WordHeaderFooterType.First) != null)) {
                AddExportDiagnostic(options, "headers-footers-omitted", "Section headers or footers were omitted because ExportHeadersAndFooters is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportFootnotes && document.FootNotes.Count > 0) {
                AddExportDiagnostic(options, "footnotes-omitted", "Footnotes were omitted because ExportFootnotes is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportEndnotes && document.EndNotes.Count > 0) {
                AddExportDiagnostic(options, "endnotes-omitted", "Endnotes were omitted because ExportEndnotes is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.IncludeSectionMetadata &&
                (document.Sections.Count > 1 || document.Sections.Any(section => HasNonDefaultPageGeometry(section.Node())))) {
                AddExportDiagnostic(options, "section-layout-flattened", "Word section page geometry is exported without page metadata because IncludeSectionMetadata is false.", OfficeConversionLossKind.Approximation);
            }
        }

        private static void ReportTableBorderLimitations(WordTableCell cell, WordHtmlExportOptions options, HashSet<string> reported) {
            const string patternCode = "word-table-border-pattern-approximated";
            if (!reported.Contains(patternCode) &&
                (IsApproximate(WordTableBorderSide.Top) || IsApproximate(WordTableBorderSide.Bottom) ||
                 IsApproximate(WordTableBorderSide.Left) || IsApproximate(WordTableBorderSide.Right))) {
                reported.Add(patternCode);
                AddExportDiagnostic(options, patternCode,
                    "Word table border patterns without a CSS equivalent are represented by solid borders.",
                    OfficeConversionLossKind.Approximation);
            }
            const string diagonalCode = "word-table-diagonal-border-omitted";
            if (!reported.Contains(diagonalCode) &&
                (IsVisible(WordConverterNodes.ResolveTableCellDiagonalBorderNode(cell, true)) ||
                 IsVisible(WordConverterNodes.ResolveTableCellDiagonalBorderNode(cell, false)))) {
                reported.Add(diagonalCode);
                AddExportDiagnostic(options, diagonalCode,
                    "Diagonal Word table-cell borders have no CSS table-border equivalent and are omitted.",
                    OfficeConversionLossKind.Omission);
            }

            bool IsApproximate(WordTableBorderSide side) {
                OfficeXmlNode border = WordConverterNodes.ResolveTableCellBorderNode(cell, side);
                string? value = border.WVal();
                return IsVisible(border) && value != "single" && value != "dashed" &&
                    value != "dotted" && value != "double";
            }
            static bool IsVisible(OfficeXmlNode border) => border.WVal() is { } value &&
                value != "nil" && value != "none";
        }

        private static bool HasNonDefaultPageGeometry(OfficeXmlNode sectionProperties) {
            OfficeXmlNode pageSize = sectionProperties.WChild("pgSz");
            if (pageSize.WAttr("orient") == "landscape") return true;
            if (pageSize.WUInt32("w") is uint width && width != 12240U) return true;
            if (pageSize.WUInt32("h") is uint height && height != 15840U) return true;

            OfficeXmlNode margin = sectionProperties.WChild("pgMar");
            return (margin.WInt32("top") is int top && top != 1440) ||
                   (margin.WUInt32("right") is uint right && right != 1440U) ||
                   (margin.WInt32("bottom") is int bottom && bottom != 1440) ||
                   (margin.WUInt32("left") is uint left && left != 1440U) ||
                   (margin.WUInt32("header") is uint header && header != 720U) ||
                   (margin.WUInt32("footer") is uint footer && footer != 720U) ||
                   (margin.WUInt32("gutter") is uint gutter && gutter != 0U);
        }

        private static bool IsRevisionElement(string localName) {
            switch (localName) {
                case "ins":
                case "del":
                case "moveFrom":
                case "moveTo":
                case "pPrChange":
                case "rPrChange":
                case "tblPrChange":
                case "tblGridChange":
                case "tblPrExChange":
                case "trPrChange":
                case "tcPrChange":
                case "sectPrChange":
                    return true;
                default:
                    return false;
            }
        }

        private readonly struct ExportInspection {
            internal ExportInspection(
                bool hasFields,
                bool hasRevisions,
                bool hasComments,
                long outputConstructionCharacters) {
                HasFields = hasFields;
                HasRevisions = hasRevisions;
                HasComments = hasComments;
                OutputConstructionCharacters = outputConstructionCharacters;
            }

            internal bool HasFields { get; }
            internal bool HasRevisions { get; }
            internal bool HasComments { get; }
            internal long OutputConstructionCharacters { get; }
        }

        private static readonly HtmlUrlPolicy ExportLinkPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();

        // Exported HTML is opened in browsers, so a document link must not carry a script or embedded page.
        private static bool IsExportableLink(string? href, WordHtmlExportOptions options, string source) {
            if (string.IsNullOrEmpty(href)) return false;
            if (HtmlUrlPolicyEvaluator.IsAllowed(href, ExportLinkPolicy)) return true;
            AddExportDiagnostic(options, "unsafe-link-dropped",
                source + " used a URL the HTML hyperlink policy rejects, so the link was removed and its text kept.",
                OfficeConversionLossKind.Omission);
            return false;
        }

        private static void AddExportDiagnostic(WordHtmlExportOptions options, string code, string message, OfficeConversionLossKind lossKind) {
            options.ConversionReport.Add(
                "OfficeIMO.Word.Html",
                code,
                message,
                HtmlDiagnosticSeverity.Warning,
                "word:document",
                null,
                lossKind);
        }

        private static void ThrowExportLimitExceeded(
            WordHtmlExportOptions options,
            string code,
            string message,
            string source,
            long actual,
            long limit) {
            string detail = "Actual=" + actual + "; Limit=" + limit;
            options.ConversionReport.Add(
                "OfficeIMO.Word.Html",
                code,
                message,
                HtmlDiagnosticSeverity.Error,
                source,
                detail,
                OfficeConversionLossKind.Failure);
            throw new HtmlConversionLimitException(code, message, source, actual, limit, detail);
        }

        private sealed class BoundedHtmlWriter : TextWriter {
            private readonly StringBuilder _builder = new StringBuilder();
            private readonly long _maxCharacters;
            private readonly Action<long> _limitExceeded;

            internal BoundedHtmlWriter(long maxCharacters, Action<long> limitExceeded) {
                _maxCharacters = maxCharacters;
                _limitExceeded = limitExceeded;
            }

            public override Encoding Encoding => Encoding.UTF8;

            public override void Write(char value) {
                EnsureCapacity(1);
                _builder.Append(value);
            }

            public override void Write(char[] buffer, int index, int count) {
                EnsureCapacity(count);
                _builder.Append(buffer, index, count);
            }

            public override void Write(string? value) {
                if (string.IsNullOrEmpty(value)) return;
                EnsureCapacity(value!.Length);
                _builder.Append(value);
            }

            public override string ToString() => _builder.ToString();

            private void EnsureCapacity(int additionalCharacters) {
                long requested = (long)_builder.Length + additionalCharacters;
                if (requested > _maxCharacters) _limitExceeded(requested);
            }
        }

        private sealed class CountingHtmlWriter : TextWriter {
            public override Encoding Encoding => Encoding.UTF8;

            internal long CharacterCount { get; private set; }

            public override void Write(char value) => AddCharacters(1);

            public override void Write(char[] buffer, int index, int count) => AddCharacters(count);

            public override void Write(string? value) {
                if (!string.IsNullOrEmpty(value)) AddCharacters(value!.Length);
            }

            private void AddCharacters(int count) {
                CharacterCount = CharacterCount > long.MaxValue - count
                    ? long.MaxValue
                    : CharacterCount + count;
            }
        }
    }
}
