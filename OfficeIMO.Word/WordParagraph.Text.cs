using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
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
                    var builder = new StringBuilder();
                    foreach (var child in _run.ChildElements) {
                        switch (child) {
                            case Text text:
                                builder.Append(text.Text);
                                break;
                            case TabChar:
                                builder.Append('\t');
                                break;
                            case Break breakNode:
                                if (IsTextWrappingBreak(breakNode)) {
                                    builder.Append(NormalizedLineFeed);
                                } else {
                                    builder.Append(NonTextBreakPlaceholder);
                                }
                                break;
                        }
                    }

                    return builder.ToString();
                }

                return string.Empty;
            }
            set {
                var run = VerifyRun();

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

                foreach (var child in run.ChildElements.ToList()) {
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
    }
}
