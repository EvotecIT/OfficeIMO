using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedInlineContentControlText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting,
            string context) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, context);
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationField(text, runs, child, inheritedFormatting);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, inheritedFormatting);
                        } else {
                            AppendSupportedRunText(
                                text,
                                runs,
                                child,
                                footnotes,
                                endnotes,
                                inheritedFormatting,
                                allowHyperlinkRunStyle: false,
                                pictures,
                                relationshipOwner);
                        }

                        break;
                    case "ins":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Inserted, footnotes, endnotes, inheritedFormatting, pictures, relationshipOwner);
                        break;
                    case "del":
                        AppendSupportedRevisionText(text, runs, child, LegacyDocRevisionKind.Deleted, footnotes, endnotes, inheritedFormatting, pictures, relationshipOwner);
                        break;
                    case "hyperlink":
                        AppendSupportedHyperlinkText(text, runs, bookmarks, child, relationshipOwner, footnotes, endnotes, inheritedFormatting);
                        break;
                    case "fldSimple":
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, child, inheritedFormatting);
                        break;
                    case "sdt":
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, child, relationshipOwner, pictures, footnotes, endnotes, inheritedFormatting, context);
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

                        throw new NotSupportedException($"Native DOC saving supports {context}s only when they contain supported text runs, {SupportedFieldNames} fields, bookmarks, nested inline content controls, and simple hyperlinks. Unsupported inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static void AppendFormattedHeaderFooterInlineContentControl(
            StringBuilder storyText,
            List<LegacyDocWritableRun> formattedRuns,
            StringBuilder paragraphText,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, $"{kind} inline content control");
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendFormattedHeaderFooterMathEquation(storyText, formattedRuns, paragraphText, child);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendFormattedHeaderFooterComplexPageNumberField(storyText, formattedRuns, paragraphText, bookmarks, children, ref index, kind);
                        } else {
                            AppendFormattedHeaderFooterRun(storyText, formattedRuns, paragraphText, child, relationshipOwner, pictures, kind);
                        }

                        break;
                    case "ins":
                        AppendFormattedHeaderFooterRevision(storyText, formattedRuns, paragraphText, child, LegacyDocRevisionKind.Inserted, relationshipOwner, pictures, kind);
                        break;
                    case "del":
                        AppendFormattedHeaderFooterRevision(storyText, formattedRuns, paragraphText, child, LegacyDocRevisionKind.Deleted, relationshipOwner, pictures, kind);
                        break;
                    case "hyperlink":
                        AppendFormattedHeaderFooterHyperlink(storyText, formattedRuns, paragraphText, bookmarks, child, relationshipOwner, kind);
                        break;
                    case "fldSimple":
                        AppendFormattedHeaderFooterPageNumberField(storyText, formattedRuns, paragraphText, bookmarks, child, kind);
                        break;
                    case "sdt":
                        AppendFormattedHeaderFooterInlineContentControl(storyText, formattedRuns, paragraphText, bookmarks, child, relationshipOwner, pictures, kind);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyText.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyText.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports {kind} inline content controls only when they contain supported text runs, {SupportedFieldNames} fields, bookmarks, nested inline content controls, and simple hyperlinks. Unsupported {kind} inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedFootnoteInlineContentControl(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            long id,
            int storyStart) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, $"footnote id '{id}' inline content control");
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationNoteField(builder, runs, bookmarks, child, storyStart);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleFootnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case "ins":
                        AppendSimpleFootnoteRevision(builder, runs, child, LegacyDocRevisionKind.Inserted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "del":
                        AppendSimpleFootnoteRevision(builder, runs, child, LegacyDocRevisionKind.Deleted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "hyperlink":
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, child, relationshipOwner, id, "footnote", storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, child, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedFootnoteInlineContentControl(builder, runs, bookmarks, child, relationshipOwner, pictures, id, storyStart);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyStart + builder.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyStart + builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports footnote id '{id}' inline content controls only when they contain supported text runs, {SupportedFieldNames} fields, bookmarks, nested inline content controls, and simple hyperlinks. Unsupported footnote inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static void AppendSupportedEndnoteInlineContentControl(
            StringBuilder builder,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            LegacyDocElement sdtRun,
            LegacyDocPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            long id,
            int storyStart) {
            LegacyDocElement[] children = GetInlineContentControlChildren(sdtRun, $"endnote id '{id}' inline content control");
            for (int index = 0; index < children.Length; index++) {
                LegacyDocElement child = children[index];
                if (child.Is(OfficeXmlNamespaces.Math, "oMath") || child.Is(OfficeXmlNamespaces.Math, "oMathPara")) {
                    AppendMathEquationNoteField(builder, runs, bookmarks, child, storyStart);
                    continue;
                }

                switch (child.WName) {
                    case "r":
                        if (IsComplexFieldBeginRun(child)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleEndnoteRun(builder, runs, child, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case "ins":
                        AppendSimpleEndnoteRevision(builder, runs, child, LegacyDocRevisionKind.Inserted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "del":
                        AppendSimpleEndnoteRevision(builder, runs, child, LegacyDocRevisionKind.Deleted, id, storyStart, relationshipOwner, pictures);
                        break;
                    case "hyperlink":
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, child, relationshipOwner, id, "endnote", storyStart);
                        break;
                    case "fldSimple":
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, child, storyStart);
                        break;
                    case "sdt":
                        AppendSupportedEndnoteInlineContentControl(builder, runs, bookmarks, child, relationshipOwner, pictures, id, storyStart);
                        break;
                    case "bookmarkStart":
                        bookmarks.AddStart(child, storyStart + builder.Length);
                        break;
                    case "bookmarkEnd":
                        bookmarks.AddEnd(child, storyStart + builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports endnote id '{id}' inline content controls only when they contain supported text runs, {SupportedFieldNames} fields, bookmarks, nested inline content controls, and simple hyperlinks. Unsupported endnote inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocElement[] GetInlineContentControlChildren(LegacyDocElement sdtRun, string context) {
            LegacyDocElement contentRun = sdtRun.Child("sdtContent");
            if (contentRun.IsNone) {
                throw new NotSupportedException($"Native DOC saving supports {context}s only when they contain run-level content.");
            }

            return contentRun.Elements().ToArray();
        }
    }
}
