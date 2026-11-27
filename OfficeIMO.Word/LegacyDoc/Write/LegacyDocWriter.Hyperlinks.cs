using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement hyperlink,
            LegacyDocPart relationshipOwner,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes) {
            AppendSupportedHyperlinkText(text, runs, bookmarks, hyperlink, relationshipOwner, footnotes, endnotes, LegacyDocWritableFormatting.Plain);
        }

        private static void AppendSupportedHyperlinkText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement hyperlink,
            LegacyDocPart relationshipOwner,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting) {
            string instruction = CreateSupportedHyperlinkInstruction(hyperlink, relationshipOwner);
            AppendFormattedText(text, runs, LegacyDocField.Begin.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
            AppendFormattedText(text, runs, instruction, LegacyDocWritableFormatting.Plain);
            AppendFormattedText(text, runs, LegacyDocField.Separator.ToString(), LegacyDocWritableFormatting.SpecialCharacter);

            int displayStart = text.Length;
            foreach (LegacyDocElement child in hyperlink.Elements()) {
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationField(text, runs, child, inheritedFormatting);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        EnsureSupportedHyperlinkRun(child);
                        AppendSupportedRunText(text, runs, child, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: true);
                        break;
                    case "ins":
                        AppendSupportedHyperlinkRevisionText(text, runs, child, LegacyDocRevisionKind.Inserted, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "del":
                        AppendSupportedHyperlinkRevisionText(text, runs, child, LegacyDocRevisionKind.Deleted, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "sdt":
                        AppendSupportedHyperlinkInlineContentControlText(text, runs, bookmarks, child, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "fldSimple":
                        AppendSupportedPageNumberFieldFromSimpleField(
                            text,
                            runs,
                            bookmarks,
                            child,
                            inheritedFormatting,
                            allowHyperlinkRunStyle: true);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple hyperlinks only when they contain text runs, supported fields, equations, and inline content controls. Unsupported hyperlink element: {child.LocalName}.");
                }
            }

            if (text.Length == displayStart) {
                throw new NotSupportedException("Native DOC saving supports hyperlinks only when they contain display text.");
            }

            AppendFormattedText(text, runs, LegacyDocField.End.ToString(), LegacyDocWritableFormatting.SpecialCharacter);
        }

        private static string CreateSupportedHyperlinkInstruction(LegacyDocElement hyperlink, LegacyDocPart relationshipOwner) {
            string? id = hyperlink.Attr(OfficeXmlNamespaces.Relationships, "id");
            if (string.IsNullOrEmpty(id)) {
                string? anchor = hyperlink.Attr("anchor");
                if (string.IsNullOrWhiteSpace(anchor)) {
                    throw new NotSupportedException("Native DOC saving supports hyperlinks only when they target an external relationship or an internal bookmark anchor.");
                }

                return AppendHyperlinkMetadata(" HYPERLINK \\l \"" + EscapeFieldString(anchor!) + "\" ", hyperlink);
            }

            LegacyDocRelationship? relationship = LegacyDocWordHelpers.FindHyperlinkRelationship(relationshipOwner, id);
            if (relationship == null) {
                throw new NotSupportedException($"Native DOC saving could not find hyperlink relationship '{id}'.");
            }

            // The SDK's relationship URI: the raw target as a relative-or-absolute URI.
            var uri = new Uri(relationship.Target!, UriKind.RelativeOrAbsolute);
            if (!uri.IsAbsoluteUri) {
                throw new NotSupportedException("Native DOC saving supports external hyperlinks only when their target URI is absolute.");
            }

            return AppendHyperlinkMetadata(" HYPERLINK \"" + EscapeFieldString(uri.ToString()) + "\" ", hyperlink);
        }

        private static string AppendHyperlinkMetadata(string instruction, LegacyDocElement hyperlink) {
            string? tooltipValue = hyperlink.Attr("tooltip");
            string? targetFrame = hyperlink.Attr("tgtFrame");
            EnsureSupportedHyperlinkMetadata(tooltipValue);
            EnsureSupportedHyperlinkMetadata(targetFrame);
            if (tooltipValue is string tooltip) {
                instruction += "\\o \"" + EscapeFieldString(tooltip) + "\" ";
            }
            if (!string.IsNullOrEmpty(targetFrame)) {
                instruction += "\\t \"" + EscapeFieldString(targetFrame!) + "\" ";
            }
            return instruction;
        }

        private static void EnsureSupportedHyperlinkMetadata(string? value) {
            if (value != null && (value.IndexOf('\r') >= 0 || value.IndexOf('\n') >= 0)) {
                throw new NotSupportedException("Native DOC saving cannot preserve line breaks in hyperlink tooltip or target-frame metadata.");
            }
        }

        private static void EnsureSupportedHyperlinkRun(LegacyDocElement run) {
            foreach (LegacyDocElement child in run.Elements()) {
                switch (child.WName) {
                    case "rPr":
                    case "lastRenderedPageBreak":
                    case "t":
                    case "delText":
                    case "tab":
                    case "cr":
                    case "noBreakHyphen":
                    case "softHyphen":
                    case "br":
                        break;
                    case "footnoteReference":
                    case "endnoteReference":
                        throw new NotSupportedException("Native DOC saving supports hyperlink display text only as regular text runs. Footnote and endnote references inside hyperlinks are not supported yet.");
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple hyperlinks only when their display text contains text, tabs, carriage returns, soft/no-break hyphens, and supported breaks. Unsupported hyperlink run element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedHyperlinkRevisionText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocElement revisionElement,
            LegacyDocRevisionKind revisionKind,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocRevision revision = ReadSupportedRevision(revisionElement, revisionKind);
            LegacyDocWritableFormatting revisionFormatting = inheritedFormatting.WithRevision(revision);
            foreach (LegacyDocElement child in revisionElement.Elements()) {
                if (child.Is("r")) {
                    EnsureSupportedHyperlinkRun(child);
                    AppendSupportedRunText(text, runs, child, footnotes, endnotes, revisionFormatting, allowHyperlinkRunStyle: true);
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports tracked insertions and deletions in hyperlinks only when they contain text runs. Unsupported revision element: {child.LocalName}.");
            }
        }

        private static void AppendSupportedHyperlinkInlineContentControlText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, "hyperlink inline content control");
            foreach (LegacyDocElement child in children) {
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationField(text, runs, child, inheritedFormatting);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        EnsureSupportedHyperlinkRun(child);
                        AppendSupportedRunText(text, runs, child, footnotes, endnotes, inheritedFormatting, allowHyperlinkRunStyle: true);
                        break;
                    case "ins":
                        AppendSupportedHyperlinkRevisionText(text, runs, child, LegacyDocRevisionKind.Inserted, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "del":
                        AppendSupportedHyperlinkRevisionText(text, runs, child, LegacyDocRevisionKind.Deleted, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "sdt":
                        AppendSupportedHyperlinkInlineContentControlText(text, runs, bookmarks, child, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "fldSimple":
                        AppendSupportedPageNumberFieldFromSimpleField(
                            text,
                            runs,
                            bookmarks,
                            child,
                            inheritedFormatting,
                            allowHyperlinkRunStyle: true);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, text.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports hyperlink inline content controls only when they contain supported text runs, supported fields, equations, and nested inline content controls. Unsupported hyperlink inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static string EscapeFieldString(string value) {
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
