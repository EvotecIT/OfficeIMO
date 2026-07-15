using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using M = DocumentFormat.OpenXml.Math;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SdtContentPicture = DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        private const string NormalizedLineFeed = "\n";

        // Non text-wrapping breaks (for example, page or column breaks) are surfaced as the Unicode line
        // separator so that they remain discoverable during text operations and can be restored exactly.
        private const char NonTextBreakPlaceholder = '\u2028';

        private static bool IsTextWrappingBreak(Break breakNode) {
            return breakNode.Type is null || breakNode.Type.Value == BreakValues.TextWrapping;
        }

        private static bool ShouldEmitTextNode(string segment, int totalSegments, bool isLastSegment) {
            // Emit text nodes for non-empty segments, retain a single empty segment when clearing text,
            // and keep the trailing empty segment created by a terminal newline so that round-trips preserve it.
            if (segment.Length > 0) {
                return true;
            }

            if (totalSegments == 1) {
                return true;
            }

            return isLastSegment;
        }

        /// <summary>
        /// Gets or sets the text for this run.
        /// Text-wrapping breaks (<c>null</c> or <c>BreakValues.TextWrapping</c>) are surfaced as <c>"\n"</c>
        /// in the returned string so callers receive the same representation on every platform. The setter
        /// accepts any mix of <c>"\r\n"</c>, <c>"\r"</c>, or <c>"\n"</c> and normalizes them to
        /// <c>"\n"</c> before updating the OpenXML elements. Non text-wrapping breaks—such as page or column
        /// breaks—are represented using the Unicode line separator character (<c>'\u2028'</c>) so that text
        /// operations (for example find/replace) can preserve their positions. When the text is modified those
        /// breaks are re-inserted at their original locations.
        /// </summary>
        public string Text {
            get {
                if (_run != null) {
                    return ReadVisibleText(_run);
                }

                if (_hyperlink != null) {
                    return ReadVisibleText(_hyperlink);
                }

                if (_simpleField != null) {
                    return ReadVisibleText(_simpleField);
                }

                if (_runs != null) {
                    return ReadComplexFieldResultText(_runs);
                }

                if (_stdRun != null) {
                    if (_stdRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().Any() == true) {
                        return string.Empty;
                    }

                    return ReadVisibleText(_stdRun);
                }

                if (_officeMath != null) {
                    return ReadVisibleText(_officeMath);
                }

                if (_mathParagraph != null) {
                    return ReadVisibleText(_mathParagraph);
                }

                return string.Empty;
            }
            set {
                if (_officeMath != null) {
                    WordMath.SetText(_officeMath, value);
                    return;
                }

                if (_mathParagraph != null) {
                    WordMath.SetText(_mathParagraph, value);
                    return;
                }

                Run run = ResolveTextSetterRun(out OpenXmlElement textContainer, out IReadOnlyList<Run>? contentRuns);

                var preservedBreaks = new List<(int ContentIndex, Break Break)>();
                int contentNodesEncountered = 0;

                static int CountContentUnits(string? text) {
                    var segment = text ?? string.Empty;
                    if (segment.Length == 0) {
                        return 1;
                    }

                    int units = 0;
                    var parts = segment.Split('\t');
                    for (int partIndex = 0; partIndex < parts.Length; partIndex++) {
                        if (parts[partIndex].Length > 0) {
                            units++;
                        }

                        if (partIndex < parts.Length - 1) {
                            units++;
                        }
                    }

                    return units;
                }

                IEnumerable<Run> runsToClear = contentRuns ?? textContainer.Descendants<Run>().Prepend(textContainer).OfType<Run>().Distinct();
                foreach (Run contentRun in runsToClear.ToList()) {
                    foreach (var child in contentRun.ChildElements.ToList()) {
                        switch (child) {
                            case Text textNode:
                                textNode.Remove();
                                contentNodesEncountered += CountContentUnits(textNode.Text);
                                break;
                            case TabChar tabChar:
                                tabChar.Remove();
                                contentNodesEncountered++;
                                break;
                            case Break breakNode:
                                if (IsTextWrappingBreak(breakNode)) {
                                    breakNode.Remove();
                                } else {
                                    preservedBreaks.Add((contentNodesEncountered, breakNode));
                                    breakNode.Remove();
                                }
                                break;
                        }
                    }
                }

                var normalized = (value ?? string.Empty)
                    .Replace("\r\n", NormalizedLineFeed)
                    .Replace("\r", NormalizedLineFeed);

                static List<(string Text, bool EndsWithTextWrappingBreak)> BuildSegments(string source) {
                    var result = new List<(string Text, bool EndsWithTextWrappingBreak)>();
                    var blocks = source.Split(NonTextBreakPlaceholder);

                    foreach (var block in blocks) {
                        var lines = block.Split('\n');
                        if (lines.Length == 0) {
                            result.Add((string.Empty, false));
                            continue;
                        }

                        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
                            var line = lines[lineIndex];
                            bool endsWithBreak = lineIndex < lines.Length - 1;
                            result.Add((line, endsWithBreak));
                        }
                    }

                    return result;
                }

                var segments = BuildSegments(normalized);
                int emittedContentCount = 0;
                int preservedIndex = 0;

                void AppendPreservedBreaksForContentIndex(int contentIndex) {
                    while (preservedIndex < preservedBreaks.Count && preservedBreaks[preservedIndex].ContentIndex == contentIndex) {
                        run.Append(preservedBreaks[preservedIndex].Break);
                        preservedIndex++;
                    }
                }

                void AppendTextWithTabs(string segment) {
                    if (segment.Length == 0) {
                        run.Append(new Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve });
                        emittedContentCount++;
                        AppendPreservedBreaksForContentIndex(emittedContentCount);
                        return;
                    }

                    var parts = segment.Split('\t');
                    for (int partIndex = 0; partIndex < parts.Length; partIndex++) {
                        if (parts[partIndex].Length > 0) {
                            run.Append(new Text(parts[partIndex]) { Space = SpaceProcessingModeValues.Preserve });
                            emittedContentCount++;
                            AppendPreservedBreaksForContentIndex(emittedContentCount);
                        }

                        if (partIndex < parts.Length - 1) {
                            run.Append(new TabChar());
                            emittedContentCount++;
                            AppendPreservedBreaksForContentIndex(emittedContentCount);
                        }
                    }
                }

                AppendPreservedBreaksForContentIndex(0);

                for (int i = 0; i < segments.Count; i++) {
                    var (segment, endsWithTextWrappingBreak) = segments[i];
                    bool isLast = i == segments.Count - 1;
                    bool shouldAddText = ShouldEmitTextNode(segment, segments.Count, isLast);

                    if (shouldAddText) {
                        AppendTextWithTabs(segment);
                    }

                    if (endsWithTextWrappingBreak) {
                        run.Append(new Break());
                    }
                }

                while (preservedIndex < preservedBreaks.Count) {
                    run.Append(preservedBreaks[preservedIndex].Break);
                    preservedIndex++;
                }
            }
        }

        private Run ResolveTextSetterRun(out OpenXmlElement textContainer, out IReadOnlyList<Run>? contentRuns) {
            contentRuns = null;
            if (_run != null) {
                textContainer = _run;
                return _run;
            }

            if (_hyperlink != null) {
                textContainer = _hyperlink;
                return EnsureTextRun(_hyperlink);
            }

            if (_simpleField != null) {
                textContainer = _simpleField;
                return EnsureTextRun(_simpleField);
            }

            if (_runs != null && _runs.Count > 0) {
                contentRuns = GetComplexFieldResultRuns(_runs);
                textContainer = contentRuns.Count > 0 ? contentRuns[0] : _runs[0];
                return contentRuns.Count > 0 ? contentRuns[0] : _runs[0];
            }

            if (_stdRun != null) {
                SdtContentRun content = _stdRun.SdtContentRun ??= new SdtContentRun();
                textContainer = content;
                return EnsureTextRun(content);
            }

            textContainer = VerifyRun();
            return (Run)textContainer;
        }

        private static Run EnsureTextRun(OpenXmlCompositeElement container) {
            Run? run = container.Descendants<Run>().FirstOrDefault();
            if (run != null) {
                return run;
            }

            run = new Run();
            container.Append(run);
            return run;
        }

        private static IReadOnlyList<Run> GetComplexFieldResultRuns(IReadOnlyList<Run> runs) {
            var resultRuns = new List<Run>();
            bool inOuterResult = false;
            int fieldDepth = 0;

            foreach (Run run in runs) {
                bool includeRun = false;
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (child is FieldChar fieldChar) {
                        FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                        if (fieldCharType == FieldCharValues.Begin) {
                            fieldDepth++;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.Separate && fieldDepth == 1) {
                            inOuterResult = true;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.End && fieldDepth > 0) {
                            if (fieldDepth == 1) {
                                inOuterResult = false;
                            }

                            fieldDepth--;
                            if (fieldDepth == 0) {
                                break;
                            }
                        }

                        continue;
                    }

                    if (inOuterResult && fieldDepth == 1) {
                        includeRun = true;
                    }
                }

                if (includeRun) {
                    resultRuns.Add(run);
                }
            }

            return resultRuns.Count > 0 ? resultRuns : runs;
        }

        private static string ReadVisibleText(OpenXmlElement element) {
            var builder = new StringBuilder();
            AppendVisibleText(builder, element);
            return builder.ToString();
        }

        private static void AppendVisibleText(StringBuilder builder, OpenXmlElement element) {
            switch (element) {
                case Run run when IsHiddenCommentReferenceRun(run):
                    return;
                case Text text:
                    builder.Append(text.Text);
                    return;
                case M.Text mathText:
                    builder.Append(mathText.Text);
                    return;
                case TabChar:
                    builder.Append('\t');
                    return;
                case CarriageReturn:
                    builder.Append(NormalizedLineFeed);
                    return;
                case NoBreakHyphen:
                    builder.Append('\u2011');
                    return;
                case SoftHyphen:
                    builder.Append('\u00ad');
                    return;
                case Break breakNode:
                    builder.Append(IsTextWrappingBreak(breakNode) ? NormalizedLineFeed : NonTextBreakPlaceholder);
                    return;
                case CommentReference:
                    return;
                case FieldCode:
                case FieldChar:
                case DeletedRun:
                case MoveFromRun:
                    return;
            }

            if (element.NamespaceUri == WordMath.MathNamespace) {
                builder.Append(WordMath.GetText(element));
                return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendVisibleText(builder, child);
            }
        }

        private static bool IsHiddenCommentReferenceRun(Run run) {
            Vanish? vanish = run.RunProperties?.GetFirstChild<Vanish>();
            return vanish != null &&
                   (vanish.Val == null || vanish.Val.Value) &&
                   run.Elements<CommentReference>().Any();
        }

        private static string ReadComplexFieldResultText(IReadOnlyList<Run> runs) {
            var builder = new StringBuilder();
            bool sawSeparator = false;
            int fieldDepth = 0;

            foreach (Run run in runs) {
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (child is FieldChar fieldChar) {
                        FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                        if (fieldCharType == FieldCharValues.Begin) {
                            fieldDepth++;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.Separate && fieldDepth > 0) {
                            sawSeparator = true;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.End && fieldDepth > 0) {
                            fieldDepth--;
                            if (fieldDepth == 0) {
                                return builder.ToString();
                            }
                        }

                        continue;
                    }

                    if (sawSeparator && fieldDepth > 0) {
                        AppendVisibleText(builder, child);
                    }
                }
            }

            return builder.ToString();
        }
    }
}
