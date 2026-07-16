using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendSupportedInlineContentControlText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            LegacyDocWritableBookmarksBuilder bookmarks,
            SdtRun sdtRun,
            MainDocumentPart mainPart,
            LegacyDocWritablePictures pictures,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            LegacyDocWritableFormatting inheritedFormatting,
            string context) {
            OpenXmlElement[] children = GetInlineContentControlChildren(sdtRun, context);
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, bookmarks, inheritedFormatting);
                        } else {
                            AppendSupportedRunText(
                                text,
                                runs,
                                run,
                                footnotes,
                                endnotes,
                                inheritedFormatting,
                                allowHyperlinkRunStyle: false,
                                pictures,
                                mainPart);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedHyperlinkText(text, runs, bookmarks, hyperlink, mainPart, footnotes, endnotes, inheritedFormatting);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, bookmarks, simpleField, inheritedFormatting);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendMathEquationField(text, runs, officeMath, inheritedFormatting);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendMathEquationField(text, runs, mathParagraph, inheritedFormatting);
                        break;
                    case SdtRun nestedSdtRun:
                        AppendSupportedInlineContentControlText(text, runs, bookmarks, nestedSdtRun, mainPart, pictures, footnotes, endnotes, inheritedFormatting, context);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, text.Length);
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
            SdtRun sdtRun,
            OpenXmlPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            string kind) {
            OpenXmlElement[] children = GetInlineContentControlChildren(sdtRun, $"{kind} inline content control");
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendFormattedHeaderFooterComplexPageNumberField(storyText, formattedRuns, paragraphText, bookmarks, children, ref index, kind);
                        } else {
                            AppendFormattedHeaderFooterRun(storyText, formattedRuns, paragraphText, run, relationshipOwner, pictures, kind);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendFormattedHeaderFooterHyperlink(storyText, formattedRuns, paragraphText, bookmarks, hyperlink, relationshipOwner, kind);
                        break;
                    case SimpleField simpleField:
                        AppendFormattedHeaderFooterPageNumberField(storyText, formattedRuns, paragraphText, bookmarks, simpleField, kind);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendFormattedHeaderFooterMathEquation(storyText, formattedRuns, paragraphText, officeMath);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendFormattedHeaderFooterMathEquation(storyText, formattedRuns, paragraphText, mathParagraph);
                        break;
                    case SdtRun nestedSdtRun:
                        AppendFormattedHeaderFooterInlineContentControl(storyText, formattedRuns, paragraphText, bookmarks, nestedSdtRun, relationshipOwner, pictures, kind);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyText.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyText.Length);
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
            SdtRun sdtRun,
            FootnotesPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            long id,
            int storyStart) {
            OpenXmlElement[] children = GetInlineContentControlChildren(sdtRun, $"footnote id '{id}' inline content control");
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleFootnoteRun(builder, runs, run, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, hyperlink, relationshipOwner, id, "footnote", storyStart);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, simpleField, storyStart);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendMathEquationNoteField(builder, runs, bookmarks, officeMath, storyStart);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendMathEquationNoteField(builder, runs, bookmarks, mathParagraph, storyStart);
                        break;
                    case SdtRun nestedSdtRun:
                        AppendSupportedFootnoteInlineContentControl(builder, runs, bookmarks, nestedSdtRun, relationshipOwner, pictures, id, storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + builder.Length);
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
            SdtRun sdtRun,
            EndnotesPart relationshipOwner,
            LegacyDocWritablePictures pictures,
            long id,
            int storyStart) {
            OpenXmlElement[] children = GetInlineContentControlChildren(sdtRun, $"endnote id '{id}' inline content control");
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedNoteComplexPageNumberField(children, ref index, builder, runs, bookmarks, storyStart);
                        } else {
                            AppendSimpleEndnoteRun(builder, runs, run, id, storyStart, relationshipOwner, pictures);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedNoteHyperlinkText(builder, runs, bookmarks, hyperlink, relationshipOwner, id, "endnote", storyStart);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedNoteFieldFromSimpleField(builder, runs, bookmarks, simpleField, storyStart);
                        break;
                    case DocumentFormat.OpenXml.Math.OfficeMath officeMath:
                        AppendMathEquationNoteField(builder, runs, bookmarks, officeMath, storyStart);
                        break;
                    case DocumentFormat.OpenXml.Math.Paragraph mathParagraph:
                        AppendMathEquationNoteField(builder, runs, bookmarks, mathParagraph, storyStart);
                        break;
                    case SdtRun nestedSdtRun:
                        AppendSupportedEndnoteInlineContentControl(builder, runs, bookmarks, nestedSdtRun, relationshipOwner, pictures, id, storyStart);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, storyStart + builder.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, storyStart + builder.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports endnote id '{id}' inline content controls only when they contain supported text runs, {SupportedFieldNames} fields, bookmarks, nested inline content controls, and simple hyperlinks. Unsupported endnote inline content-control element: {child.LocalName}.");
                }
            }
        }

        private static OpenXmlElement[] GetInlineContentControlChildren(SdtRun sdtRun, string context) {
            SdtContentRun? contentRun = sdtRun.SdtContentRun;
            if (contentRun == null) {
                throw new NotSupportedException($"Native DOC saving supports {context}s only when they contain run-level content.");
            }

            return contentRun.ChildElements.ToArray();
        }
    }
}
