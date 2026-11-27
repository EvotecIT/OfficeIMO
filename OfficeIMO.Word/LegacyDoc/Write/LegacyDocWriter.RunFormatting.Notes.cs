using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendFootnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, LegacyDocElement run, LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFormatting formatting = ReadSupportedNoteReferenceFormatting(run, inheritedFormatting);
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "footnoteReference":
                        AppendFootnoteReference(text, runs, footnotes, child, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports footnote reference runs only when they contain footnote references. Unsupported footnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendEndnoteReferenceRun(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, LegacyDocElement run, LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocWritableFormatting formatting = ReadSupportedNoteReferenceFormatting(run, inheritedFormatting);
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                        break;
                    case "lastRenderedPageBreak":
                        break;
                    case "endnoteReference":
                        AppendEndnoteReference(text, runs, endnotes, child, formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports endnote reference runs only when they contain endnote references. Unsupported endnote reference run element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableFormatting ReadSupportedNoteReferenceFormatting(LegacyDocElement run, LegacyDocWritableFormatting inheritedFormatting) {
            // Native note style slots cannot carry the source character-style
            // definitions. Materialize their typography before direct overrides.
            LegacyDocElement runProperties = run.Child("rPr");
            LegacyDocWritableFormatting direct = ReadSupportedRunFormatting(runProperties,
                allowHyperlinkRunStyle: false, allowNoteReferenceRunStyle: true);
            LegacyDocSourcePackage? source = t_writingSource;
            string? styleId = runProperties.Child("rStyle").Val;
            var visited = new HashSet<string>(StringComparer.Ordinal);
            LegacyDocStyleCatalog? definitions = source == null || !source.Styles.Exists ? null : source.Styles;
            LegacyDocWritableFormatting styleFormatting = LegacyDocWritableFormatting.Plain;
            while (definitions != null && !string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                if (visited.Count > MaxInheritedStyleDepth) throw new NotSupportedException("Native DOC saving does not support note reference style chains beyond 256 levels.");
                LegacyDocElement style = definitions.FindById(styleId);
                if (style.IsNone) break;
                styleFormatting = styleFormatting.WithInheritedFormatting(
                    ReadSupportedRunFormatting(style.Child("rPr"), allowHyperlinkRunStyle: false));
                styleId = style.Child("basedOn").Val;
            }
            return direct.WithInheritedFormatting(styleFormatting.WithInheritedFormatting(inheritedFormatting));
        }

        private static void AppendFootnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableFootnotes footnotes, LegacyDocElement footnoteReference, LegacyDocWritableFormatting formatting) {
            long? id = LegacyDocValues.Int64(footnoteReference.Attr("id"));
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports footnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            footnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), formatting.WithSpecialCharacter());
        }

        private static void AppendEndnoteReference(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableEndnotes endnotes, LegacyDocElement endnoteReference, LegacyDocWritableFormatting formatting) {
            long? id = LegacyDocValues.Int64(endnoteReference.Attr("id"));
            if (id == null || id.Value <= 0) {
                throw new NotSupportedException("Native DOC saving supports endnote references only when they use a positive identifier.");
            }

            int referencePosition = text.Length;
            endnotes.AddReference(id.Value, referencePosition);
            AppendFormattedText(text, runs, LegacyDocFootnoteReader.FootnoteReferenceCharacter.ToString(), formatting.WithSpecialCharacter());
        }

    }
}
