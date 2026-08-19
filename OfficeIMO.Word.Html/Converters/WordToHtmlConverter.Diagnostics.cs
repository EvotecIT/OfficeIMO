using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using System.IO;
using System.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static ExportInspection InspectExport(WordDocument document, WordToHtmlOptions options) {
            HtmlConversionProfileContracts.Get(options.SharedProfile);
            options.DocumentOutput.Validate();
            if (options.MaxDocumentElements <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDocumentElements));
            if (options.MaxEmbeddedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedImageBytes));
            if (options.MaxTotalEmbeddedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalEmbeddedImageBytes));
            if (options.MaxOutputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxOutputCharacters));
            if (options.MaxEquationNestingDepth <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEquationNestingDepth));
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
            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart != null) {
                foreach ((OpenXmlElement Root, string PartUri) root in EnumerateExportRoots(
                    mainPart,
                    options.IncludeCustomProperties ? document._wordprocessingDocument.CustomFilePropertiesPart : null)) {
                    bool countOutputContent = IsOutputContentRoot(root.Root, options);
                    bool inspectExportedStory = IsSelectedStoryRoot(root.Root, options);
                    foreach ((OpenXmlElement Element, bool OmitOutputContent) inspected in EnumerateRootAndDescendants(
                                 root.Root,
                                 options.TrackedChangePolicy)) {
                        OpenXmlElement element = inspected.Element;
                        elements++;
                        if (elements > options.MaxDocumentElements) {
                            ThrowExportLimitExceeded(options, "WordElementLimitExceeded", "The Word document exceeds the configured HTML export element limit.", root.PartUri, elements, options.MaxDocumentElements);
                        }
                        if (countOutputContent && !inspected.OmitOutputContent) {
                            long elementCharacters = GetOutputContentCharacters(element);
                            outputConstructionCharacters = SaturatingAdd(outputConstructionCharacters, elementCharacters);
                            if (outputConstructionCharacters > options.MaxOutputCharacters) {
                                ThrowExportLimitExceeded(
                                    options,
                                    "WordHtmlOutputLimitExceeded",
                                    "Word text and attribute content exceeds the configured HTML output-character limit before DOM construction.",
                                    root.PartUri,
                                    outputConstructionCharacters,
                                    options.MaxOutputCharacters);
                            }
                        }
                        if (!hasFields && inspectExportedStory && !inspected.OmitOutputContent &&
                            (element is DocumentFormat.OpenXml.Wordprocessing.SimpleField or
                                DocumentFormat.OpenXml.Wordprocessing.FieldChar or
                                DocumentFormat.OpenXml.Wordprocessing.FieldCode)) {
                            hasFields = true;
                        }
                        if (!hasRevisions && inspectExportedStory && IsRevisionElement(element.LocalName)) {
                            hasRevisions = true;
                        }
                        if (!hasComments && element is DocumentFormat.OpenXml.Wordprocessing.Comment) {
                            hasComments = true;
                        }
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

        private static long MeasureOutputContentCharacters(OpenXmlElement root) {
            long characters = 0;
            foreach ((OpenXmlElement Element, bool OmitOutputContent) inspected in EnumerateRootAndDescendants(
                         root,
                         WordTrackedChangeExportPolicy.Markup)) {
                if (inspected.OmitOutputContent) continue;
                characters = SaturatingAdd(characters, GetOutputContentCharacters(inspected.Element));
            }
            return characters;
        }

        private static long GetOutputContentCharacters(OpenXmlElement element) {
            long characters = element is OpenXmlLeafTextElement textElement && ShouldCountOutputLeafText(element)
                ? GetHtmlEncodedLength(textElement.Text, attributeValue: false)
                : 0;
            if (IsVisibleRunArtifact(element)) {
                characters = SaturatingAdd(characters, 1);
            }
            return characters;
        }

        private static bool IsVisibleRunArtifact(OpenXmlElement element) =>
            element is W.TabChar or W.CarriageReturn or W.Break or W.NoBreakHyphen or W.SoftHyphen;

        private static bool ShouldCountOutputLeafText(OpenXmlElement element) =>
            element is not DocumentFormat.OpenXml.Wordprocessing.FieldCode &&
            element is not DocumentFormat.OpenXml.Math.Text;

        // Header and footer parts are intentionally excluded here. A shared part can be emitted
        // once per section, so AppendHeaderFooterRegion measures and reserves its source content
        // immediately before every emitted region rather than charging the package part once.
        private static bool IsOutputContentRoot(OpenXmlElement root, WordToHtmlOptions options) =>
            root is not DocumentFormat.OpenXml.Wordprocessing.Header &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Footer &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Footnotes &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Endnotes &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Comments &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Styles &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Numbering &&
            root is not DocumentFormat.OpenXml.CustomProperties.Properties &&
            root is not DocumentFormat.OpenXml.Drawing.Theme &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Fonts &&
            root is not DocumentFormat.OpenXml.Wordprocessing.Settings &&
            root is not DocumentFormat.OpenXml.Wordprocessing.WebSettings;

        private static IEnumerable<(OpenXmlElement Element, bool OmitOutputContent)> EnumerateRootAndDescendants(
            OpenXmlElement root,
            WordTrackedChangeExportPolicy policy) {
            var pending = new Stack<(OpenXmlElement Element, bool InOmittedRevision)>();
            pending.Push((root, false));
            while (pending.Count > 0) {
                (OpenXmlElement element, bool inOmittedRevision) = pending.Pop();
                bool omitOutputContent = inOmittedRevision || IsRevisionContentExcluded(element, policy);
                yield return (element, omitOutputContent);
                for (int index = element.ChildElements.Count - 1; index >= 0; index--) {
                    pending.Push((element.ChildElements[index], omitOutputContent));
                }
            }
        }

        private static bool IsRevisionContentExcluded(OpenXmlElement element, WordTrackedChangeExportPolicy policy) {
            if (element.LocalName is not ("ins" or "del" or "moveFrom" or "moveTo")) return false;
            bool insertedView = element.LocalName is "ins" or "moveTo";
            return policy switch {
                WordTrackedChangeExportPolicy.Final => !insertedView,
                WordTrackedChangeExportPolicy.Original => insertedView,
                WordTrackedChangeExportPolicy.Markup => false,
                _ => throw new ArgumentOutOfRangeException(nameof(policy), policy, "Word tracked-change policy is not supported.")
            };
        }

        private static bool IsSelectedStoryRoot(OpenXmlElement root, WordToHtmlOptions options) => root switch {
            DocumentFormat.OpenXml.Wordprocessing.Body => true,
            DocumentFormat.OpenXml.Wordprocessing.Header or DocumentFormat.OpenXml.Wordprocessing.Footer => options.ExportHeadersAndFooters,
            DocumentFormat.OpenXml.Wordprocessing.Footnotes => options.ExportFootnotes,
            DocumentFormat.OpenXml.Wordprocessing.Endnotes => options.ExportEndnotes,
            DocumentFormat.OpenXml.Wordprocessing.Comments => options.ExportComments,
            _ => false
        };

        private static IEnumerable<(OpenXmlElement Root, string PartUri)> EnumerateExportRoots(
            MainDocumentPart mainPart,
            CustomFilePropertiesPart? customPropertiesPart) {
            if (customPropertiesPart?.Properties is OpenXmlElement customProperties) {
                yield return (customProperties, customPropertiesPart.Uri.ToString());
            }

            foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
                yield return (root.Root, root.PartUri);
            }
            if (mainPart.WordprocessingCommentsPart?.Comments is OpenXmlElement comments) {
                yield return (comments, mainPart.WordprocessingCommentsPart.Uri.ToString());
            }
            if (mainPart.StyleDefinitionsPart?.Styles is OpenXmlElement styles) {
                yield return (styles, mainPart.StyleDefinitionsPart.Uri.ToString());
            }
            if (mainPart.StylesWithEffectsPart?.Styles is OpenXmlElement stylesWithEffects) {
                yield return (stylesWithEffects, mainPart.StylesWithEffectsPart.Uri.ToString());
            }
            if (mainPart.NumberingDefinitionsPart?.Numbering is OpenXmlElement numbering) {
                yield return (numbering, mainPart.NumberingDefinitionsPart.Uri.ToString());
            }
            if (mainPart.ThemePart?.Theme is OpenXmlElement theme) {
                yield return (theme, mainPart.ThemePart.Uri.ToString());
            }
            if (mainPart.FontTablePart?.Fonts is OpenXmlElement fonts) {
                yield return (fonts, mainPart.FontTablePart.Uri.ToString());
            }
            if (mainPart.DocumentSettingsPart?.Settings is OpenXmlElement settings) {
                yield return (settings, mainPart.DocumentSettingsPart.Uri.ToString());
            }
            if (mainPart.WebSettingsPart?.WebSettings is OpenXmlElement webSettings) {
                yield return (webSettings, mainPart.WebSettingsPart.Uri.ToString());
            }
        }

        private static void ReportKnownExportLimitations(WordDocument document, WordToHtmlOptions options, ExportInspection inspection) {
            if (!options.EmitDocumentShell &&
                (options.IncludeCustomProperties || options.AdditionalMetaTags.Count > 0 || options.AdditionalLinkTags.Count > 0)) {
                AddExportDiagnostic(options, "DocumentHeadMetadataOmittedForFragment",
                    "Document-head metadata and links were omitted because fragment output was requested.",
                    OfficeConversionLossKind.Omission);
            }
            if (inspection.HasRevisions) {
                string policy = options.TrackedChangePolicy.ToString();
                AddExportDiagnostic(options, "TrackedRevisionsProjected",
                    "Tracked revisions were projected through the explicit " + policy + " static HTML policy; live accept/reject behavior remains in Word.",
                    OfficeConversionLossKind.Approximation);
            }
            if (inspection.HasComments && !options.ExportComments) {
                AddExportDiagnostic(options, "CommentsOmitted", "Word comments were omitted because ExportComments is false.", OfficeConversionLossKind.Omission);
            }
            if (inspection.HasFields) {
                AddExportDiagnostic(options, "FieldInstructionsFlattened",
                    options.FieldPolicy == WordFieldExportPolicy.VisibleResultWithReviewMetadata
                        ? "Word fields use their stored visible results and inert instruction metadata; live field evaluation is not represented in HTML."
                        : "Word field instructions are exported through their visible results; live field behavior is not represented in HTML.",
                    OfficeConversionLossKind.Approximation);
            }
            if (document.HasMacros) {
                AddExportDiagnostic(options, "MacroProjectOmitted", "The VBA project is package metadata and is not represented in HTML.", OfficeConversionLossKind.Omission);
            }
            if (document._wordprocessingDocument.DigitalSignatureOriginPart != null ||
                document.ApplicationProperties.HasDigitalSignatureMetadata) {
                AddExportDiagnostic(options, "PackageSignaturesOmitted", "OPC package signature metadata is not represented in HTML.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportHeadersAndFooters && document.Sections.Any(section =>
                section.Header.Default != null || section.Header.Even != null || section.Header.First != null ||
                section.Footer.Default != null || section.Footer.Even != null || section.Footer.First != null)) {
                AddExportDiagnostic(options, "HeadersFootersOmitted", "Section headers or footers were omitted because ExportHeadersAndFooters is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportFootnotes && document.FootNotes.Count > 0) {
                AddExportDiagnostic(options, "FootnotesOmitted", "Footnotes were omitted because ExportFootnotes is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.ExportEndnotes && document.EndNotes.Count > 0) {
                AddExportDiagnostic(options, "EndnotesOmitted", "Endnotes were omitted because ExportEndnotes is false.", OfficeConversionLossKind.Omission);
            }
            if (!options.IncludeSectionMetadata &&
                (document.Sections.Count > 1 || document.Sections.Any(section => HasNonDefaultPageGeometry(section._sectionProperties)))) {
                AddExportDiagnostic(options, "SectionLayoutFlattened", "Word section page geometry is exported without page metadata because IncludeSectionMetadata is false.", OfficeConversionLossKind.Approximation);
            }
        }

        private static bool HasNonDefaultPageGeometry(DocumentFormat.OpenXml.Wordprocessing.SectionProperties sectionProperties) {
            DocumentFormat.OpenXml.Wordprocessing.PageSize? pageSize = sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>();
            if (pageSize?.Orient?.Value == DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Landscape) return true;
            if (pageSize?.Width?.Value is uint width && width != 12240U) return true;
            if (pageSize?.Height?.Value is uint height && height != 15840U) return true;

            DocumentFormat.OpenXml.Wordprocessing.PageMargin? margin = sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageMargin>();
            return (margin?.Top?.Value is int top && top != 1440) ||
                   (margin?.Right?.Value is uint right && right != 1440U) ||
                   (margin?.Bottom?.Value is int bottom && bottom != 1440) ||
                   (margin?.Left?.Value is uint left && left != 1440U) ||
                   (margin?.Header?.Value is uint header && header != 720U) ||
                   (margin?.Footer?.Value is uint footer && footer != 720U) ||
                   (margin?.Gutter?.Value is uint gutter && gutter != 0U);
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

        private static void AddExportDiagnostic(WordToHtmlOptions options, string code, string message, OfficeConversionLossKind lossKind) {
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
            WordToHtmlOptions options,
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
