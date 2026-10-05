using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendFootnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, Run run, LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties,
                allowHyperlinkRunStyle: false, allowNoteReferenceRunStyle: true).WithInheritedFormatting(inheritedFormatting);
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
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties,
                allowHyperlinkRunStyle: false, allowNoteReferenceRunStyle: true).WithInheritedFormatting(inheritedFormatting);
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
