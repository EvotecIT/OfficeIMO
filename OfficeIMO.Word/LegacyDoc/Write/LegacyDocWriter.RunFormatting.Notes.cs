using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendFootnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, Run run, LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFormatting formatting = ReadSupportedNoteReferenceFormatting(run, inheritedFormatting);
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case FootnoteReference footnoteReference:
                        AppendFootnoteReference(text, runs, footnotes, footnoteReference, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports footnote reference runs only when they contain footnote references. Unsupported footnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendEndnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, Run run, LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFormatting formatting = ReadSupportedNoteReferenceFormatting(run, inheritedFormatting);
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case LastRenderedPageBreak:
                        break;
                    case EndnoteReference endnoteReference:
                        AppendEndnoteReference(text, runs, endnotes, endnoteReference, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports endnote reference runs only when they contain endnote references. Unsupported endnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableFormatting ReadSupportedNoteReferenceFormatting(Run run, LegacyDocWritableFormatting inheritedFormatting) {
            // Native note style slots cannot carry the source character-style
            // definitions. Materialize their typography before direct overrides.
            LegacyDocWritableFormatting direct = ReadSupportedRunFormatting(run.RunProperties,
                allowHyperlinkRunStyle: false, allowNoteReferenceRunStyle: true);
            OpenXmlPartRootElement? root = run.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
            Styles? styles = (root?.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            string? styleId = run.RunProperties?.RunStyle?.Val?.Value;
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            LegacyDocWritableFormatting styleFormatting = LegacyDocWritableFormatting.Plain;
            while (styles != null && !string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                Style? style = styles.Elements<Style>().FirstOrDefault(style =>
                    string.Equals(style.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase));
                if (style == null) break;
                styleFormatting = styleFormatting.WithInheritedFormatting(
                    ReadSupportedRunFormatting(style.StyleRunProperties, allowHyperlinkRunStyle: false));
                styleId = style.BasedOn?.Val?.Value;
            }
            return direct.WithInheritedFormatting(styleFormatting.WithInheritedFormatting(inheritedFormatting));
        }

        private static void AppendFootnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, FootnoteReference footnoteReference, LegacyDocWritableFormatting formatting) {
            long? id = footnoteReference.Id?.Value;
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports footnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            footnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), formatting.WithSpecialCharacter());
        }

        private static void AppendEndnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, EndnoteReference endnoteReference, LegacyDocWritableFormatting formatting) {
            long? id = endnoteReference.Id?.Value;
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports endnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            endnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), formatting.WithSpecialCharacter());
        }

    }
}
