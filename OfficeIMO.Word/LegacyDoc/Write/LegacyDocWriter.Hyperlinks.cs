using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            Hyperlink hyperlink,
            OpenXmlPartContainer relationshipOwner,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes) {
            AppendSupportedHyperlinkText(text, runs, hyperlink, relationshipOwner, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            Hyperlink hyperlink,
            OpenXmlPartContainer relationshipOwner,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting) {
            string instruction = CreateSupportedHyperlinkInstruction(hyperlink, relationshipOwner);
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, instruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);

            int displayStart = text.Length;
            foreach (OpenXmlElement child in hyperlink.ChildElements) {
                switch (child) {
                    case Run run:
                        EnsureSupportedHyperlinkRun(run);
                        AppendSupportedRunText(text, runs, run, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: true);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple hyperlinks only when they contain text runs. Unsupported hyperlink element: {child.LocalName}.");
                }
            }

            if (text.Length == displayStart) {
                throw new NotSupportedException("Native DOC saving supports hyperlinks only when they contain display text.");
            }

            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static string CreateSupportedHyperlinkInstruction(Hyperlink hyperlink, OpenXmlPartContainer relationshipOwner) {
            if (string.IsNullOrEmpty(hyperlink.Id)) {
                if (string.IsNullOrWhiteSpace(hyperlink.Anchor?.Value)) {
                    throw new NotSupportedException("Native DOC saving supports hyperlinks only when they target an external relationship or an internal bookmark anchor.");
                }

                return " HYPERLINK \\l \"" + EscapeFieldString(hyperlink.Anchor!.Value!) + "\" ";
            }

            HyperlinkRelationship? relationship = relationshipOwner.HyperlinkRelationships.FirstOrDefault(item => item.Id == hyperlink.Id);
            if (relationship == null) {
                throw new NotSupportedException($"Native DOC saving could not find hyperlink relationship '{hyperlink.Id}'.");
            }

            if (!relationship.Uri.IsAbsoluteUri) {
                throw new NotSupportedException("Native DOC saving supports external hyperlinks only when their target URI is absolute.");
            }

            return " HYPERLINK \"" + EscapeFieldString(relationship.Uri.ToString()) + "\" ";
        }

        private static void EnsureSupportedHyperlinkRun(Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case Text:
                    case TabChar:
                    case Break:
                        break;
                    case FootnoteReference:
                    case EndnoteReference:
                        throw new NotSupportedException("Native DOC saving supports hyperlink display text only as regular text runs. Footnote and endnote references inside hyperlinks are not supported yet.");
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple hyperlinks only when their display text contains text, tabs, and supported breaks. Unsupported hyperlink run element: {child.LocalName}.");
                }
            }
        }

        private static string EscapeFieldString(string value) {
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
