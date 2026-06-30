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
            MainDocumentPart mainPart,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes) {
            AppendSupportedHyperlinkText(text, runs, hyperlink, mainPart, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            Hyperlink hyperlink,
            MainDocumentPart mainPart,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting) {
            Uri uri = ReadSupportedExternalHyperlinkUri(hyperlink, mainPart);
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, " HYPERLINK \"" + EscapeFieldString(uri.ToString()) + "\" ", LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);

            int displayStart = text.Length;
            foreach (OpenXmlElement child in hyperlink.ChildElements) {
                switch (child) {
                    case Run run:
                        EnsureSupportedHyperlinkRun(run);
                        AppendSupportedRunText(text, runs, run, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: true);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple external hyperlinks only when they contain text runs. Unsupported hyperlink element: {child.LocalName}.");
                }
            }

            if (text.Length == displayStart) {
                throw new NotSupportedException("Native DOC saving supports hyperlinks only when they contain display text.");
            }

            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static Uri ReadSupportedExternalHyperlinkUri(Hyperlink hyperlink, MainDocumentPart mainPart) {
            if (string.IsNullOrEmpty(hyperlink.Id)) {
                throw new NotSupportedException("Native DOC saving supports external hyperlinks only. Internal bookmark hyperlinks are not supported yet.");
            }

            HyperlinkRelationship? relationship = mainPart.HyperlinkRelationships.FirstOrDefault(item => item.Id == hyperlink.Id);
            if (relationship == null) {
                throw new NotSupportedException($"Native DOC saving could not find hyperlink relationship '{hyperlink.Id}'.");
            }

            if (!relationship.Uri.IsAbsoluteUri) {
                throw new NotSupportedException("Native DOC saving supports external hyperlinks only when their target URI is absolute.");
            }

            return relationship.Uri;
        }

        private static void EnsureSupportedHyperlinkRun(Run run) {
            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                    case Text:
                        break;
                    case FootnoteReference:
                    case EndnoteReference:
                        throw new NotSupportedException("Native DOC saving supports hyperlink display text only as regular text runs. Footnote and endnote references inside hyperlinks are not supported yet.");
                    case TabChar:
                    case Break:
                        throw new NotSupportedException("Native DOC saving supports simple external hyperlinks only when their display text is plain text. Tabs and breaks inside hyperlinks are not supported yet.");
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple external hyperlinks only when their display text is plain text. Unsupported hyperlink run element: {child.LocalName}.");
                }
            }
        }

        private static string EscapeFieldString(string value) {
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
