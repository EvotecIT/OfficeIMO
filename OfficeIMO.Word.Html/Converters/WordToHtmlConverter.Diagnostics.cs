using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static ExportInspection InspectExport(WordDocument document, WordToHtmlOptions options) {
            if (options.MaxDocumentElements <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxDocumentElements));
            if (options.MaxEmbeddedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEmbeddedImageBytes));
            if (options.MaxTotalEmbeddedImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalEmbeddedImageBytes));
            if (options.MaxOutputCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxOutputCharacters));

            long elements = 0;
            bool hasFields = false;
            bool hasRevisions = false;
            bool hasComments = false;
            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            if (mainPart != null) {
                foreach ((OpenXmlElement Root, string PartUri) root in EnumerateExportRoots(mainPart)) {
                    foreach (OpenXmlElement element in EnumerateRootAndDescendants(root.Root)) {
                        elements++;
                        if (elements > options.MaxDocumentElements) {
                            ThrowExportLimitExceeded(options, "WordElementLimitExceeded", "The Word document exceeds the configured HTML export element limit.", root.PartUri, elements, options.MaxDocumentElements);
                        }
                        if (!hasFields && element is DocumentFormat.OpenXml.Wordprocessing.SimpleField or DocumentFormat.OpenXml.Wordprocessing.FieldChar or DocumentFormat.OpenXml.Wordprocessing.FieldCode) {
                            hasFields = true;
                        }
                        if (!hasRevisions && IsRevisionElement(element.LocalName)) {
                            hasRevisions = true;
                        }
                        if (!hasComments && element is DocumentFormat.OpenXml.Wordprocessing.Comment) {
                            hasComments = true;
                        }
                    }
                }
            }
            return new ExportInspection(hasFields, hasRevisions, hasComments);
        }

        private static IEnumerable<OpenXmlElement> EnumerateRootAndDescendants(OpenXmlElement root) {
            yield return root;
            foreach (OpenXmlElement descendant in root.Descendants()) yield return descendant;
        }

        private static IEnumerable<(OpenXmlElement Root, string PartUri)> EnumerateExportRoots(MainDocumentPart mainPart) {
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
            if (inspection.HasRevisions) {
                AddExportDiagnostic(options, "TrackedRevisionsFlattened", "Tracked revisions are exported as visible document content without Word revision semantics.", HtmlConversionLossKind.Approximation);
            }
            if (inspection.HasComments && !options.ExportComments) {
                AddExportDiagnostic(options, "CommentsOmitted", "Word comments were omitted because ExportComments is false.", HtmlConversionLossKind.Omission);
            }
            if (inspection.HasFields) {
                AddExportDiagnostic(options, "FieldInstructionsFlattened", "Word field instructions are exported through their visible results; live field behavior is not represented in HTML.", HtmlConversionLossKind.Approximation);
            }
            if (document.HasMacros) {
                AddExportDiagnostic(options, "MacroProjectOmitted", "The VBA project is package metadata and is not represented in HTML.", HtmlConversionLossKind.Omission);
            }
            if (document._wordprocessingDocument.DigitalSignatureOriginPart != null) {
                AddExportDiagnostic(options, "PackageSignaturesOmitted", "OPC package signatures are not represented in HTML.", HtmlConversionLossKind.Omission);
            }
            if (!options.ExportHeadersAndFooters && document.Sections.Any(section => section.Header != null || section.Footer != null)) {
                AddExportDiagnostic(options, "HeadersFootersOmitted", "Section headers or footers were omitted because ExportHeadersAndFooters is false.", HtmlConversionLossKind.Omission);
            }
            if (!options.ExportFootnotes && document.FootNotes.Count > 0) {
                AddExportDiagnostic(options, "FootnotesOmitted", "Footnotes were omitted because ExportFootnotes is false.", HtmlConversionLossKind.Omission);
            }
            if (!options.ExportEndnotes && document.EndNotes.Count > 0) {
                AddExportDiagnostic(options, "EndnotesOmitted", "Endnotes were omitted because ExportEndnotes is false.", HtmlConversionLossKind.Omission);
            }
            if (!options.IncludeSectionMetadata && document.Sections.Count > 1) {
                AddExportDiagnostic(options, "SectionLayoutFlattened", "Multiple Word sections are exported without page geometry metadata because IncludeSectionMetadata is false.", HtmlConversionLossKind.Approximation);
            }
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
            internal ExportInspection(bool hasFields, bool hasRevisions, bool hasComments) {
                HasFields = hasFields;
                HasRevisions = hasRevisions;
                HasComments = hasComments;
            }

            internal bool HasFields { get; }
            internal bool HasRevisions { get; }
            internal bool HasComments { get; }
        }

        private static void AddExportDiagnostic(WordToHtmlOptions options, string code, string message, HtmlConversionLossKind lossKind) {
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
                HtmlConversionLossKind.Failure);
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
    }
}
