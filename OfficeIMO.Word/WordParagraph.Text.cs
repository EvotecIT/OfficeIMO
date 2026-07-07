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
        private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

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
                    SetMathElementText(_officeMath, value);
                    return;
                }

                if (_mathParagraph != null) {
                    SetMathElementText(_mathParagraph, value);
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
                textContainer = _stdRun;
                return EnsureTextRun(_stdRun);
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
            bool sawSeparator = false;
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

                        if (fieldCharType == FieldCharValues.Separate && fieldDepth > 0) {
                            sawSeparator = true;
                            continue;
                        }

                        if (fieldCharType == FieldCharValues.End && fieldDepth > 0) {
                            fieldDepth--;
                            if (fieldDepth == 0) {
                                break;
                            }
                        }

                        continue;
                    }

                    if (sawSeparator && fieldDepth > 0) {
                        includeRun = true;
                    }
                }

                if (includeRun) {
                    resultRuns.Add(run);
                }
            }

            return resultRuns.Count > 0 ? resultRuns : runs;
        }

        private static void SetMathElementText(OpenXmlElement element, string? value) {
            var normalized = (value ?? string.Empty)
                .Replace("\r\n", NormalizedLineFeed)
                .Replace("\r", NormalizedLineFeed);
            List<M.Text> textNodes = element.Descendants<M.Text>().ToList();
            if (textNodes.Count == 0) {
                M.Run? mathRun = element.Descendants<M.Run>().FirstOrDefault();
                if (mathRun == null) {
                    mathRun = new M.Run();
                    if (element is OpenXmlCompositeElement composite) {
                        composite.Append(mathRun);
                    } else {
                        return;
                    }
                }

                mathRun.Append(new M.Text(normalized));
                return;
            }

            textNodes[0].Text = normalized;
            for (int i = 1; i < textNodes.Count; i++) {
                textNodes[i].Remove();
            }
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
                case CommentReference commentReference:
                    builder.Append("[c");
                    builder.Append(commentReference.Id?.Value ?? "?");
                    builder.Append(']');
                    return;
                case FieldCode:
                case FieldChar:
                    return;
            }

            if (element.NamespaceUri == MathNamespace && TryAppendMathText(builder, element)) {
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

        private static bool TryAppendMathText(StringBuilder builder, OpenXmlElement element) {
            switch (element.LocalName) {
                case "f":
                    AppendDelimitedMath(builder, element, "num", "den", "(", ")/(", ")");
                    return true;
                case "sSup":
                    AppendMathChildText(builder, element, "e");
                    AppendMathScript(builder, "^", ReadMathChildText(element, "sup"));
                    return true;
                case "sSub":
                    AppendMathChildText(builder, element, "e");
                    AppendMathScript(builder, "_", ReadMathChildText(element, "sub"));
                    return true;
                case "sSubSup":
                    AppendMathChildText(builder, element, "e");
                    AppendMathScript(builder, "_", ReadMathChildText(element, "sub"));
                    AppendMathScript(builder, "^", ReadMathChildText(element, "sup"));
                    return true;
                case "sPre":
                    AppendMathScript(builder, "^", ReadMathChildText(element, "sup"));
                    AppendMathScript(builder, "_", ReadMathChildText(element, "sub"));
                    AppendMathChildText(builder, element, "e");
                    return true;
                case "rad":
                    string degree = ReadMathChildText(element, "deg");
                    string radicand = ReadMathChildText(element, "e");
                    if (degree.Length == 0) {
                        builder.Append("sqrt(");
                        builder.Append(radicand);
                        builder.Append(')');
                    } else {
                        builder.Append("root(");
                        builder.Append(degree);
                        builder.Append(',');
                        builder.Append(radicand);
                        builder.Append(')');
                    }

                    return true;
                case "nary":
                case "int":
                    AppendNaryMathText(builder, element);
                    return true;
                case "func":
                    string functionName = ReadMathChildText(element, "fName");
                    string argument = ReadMathChildText(element, "e");
                    if (functionName.Length > 0) {
                        builder.Append(functionName);
                        builder.Append('(');
                        builder.Append(argument);
                        builder.Append(')');
                    } else {
                        builder.Append(argument);
                    }

                    return true;
                case "acc":
                    AppendAccentMathText(builder, element);
                    return true;
                case "bar":
                    AppendFunctionMathText(builder, "bar", ReadMathChildText(element, "e"));
                    return true;
                case "d":
                    AppendDelimiterMathText(builder, element);
                    return true;
                case "groupChr":
                    AppendGroupCharMathText(builder, element);
                    return true;
                case "m":
                    AppendMatrixMathText(builder, element);
                    return true;
                case "eqArr":
                    AppendEquationArrayMathText(builder, element);
                    return true;
                case "limLow":
                    AppendMathChildText(builder, element, "e");
                    AppendMathScript(builder, "_", ReadMathChildText(element, "lim"));
                    return true;
                case "limUpp":
                    AppendMathChildText(builder, element, "e");
                    AppendMathScript(builder, "^", ReadMathChildText(element, "lim"));
                    return true;
            }

            return false;
        }

        private static void AppendDelimitedMath(StringBuilder builder, OpenXmlElement element, string leftChild, string rightChild, string prefix, string separator, string suffix) {
            builder.Append(prefix);
            AppendMathChildText(builder, element, leftChild);
            builder.Append(separator);
            AppendMathChildText(builder, element, rightChild);
            builder.Append(suffix);
        }

        private static void AppendAccentMathText(StringBuilder builder, OpenXmlElement element) {
            string expression = ReadMathChildText(element, "e");
            string accent = ReadMathCharacterValue(element, "chr");
            string functionName = accent switch {
                "^" => "hat",
                "\u0302" => "hat",
                "~" => "tilde",
                "\u0303" => "tilde",
                "." => "dot",
                "\u0307" => "dot",
                "\u00a8" => "ddot",
                "\u0308" => "ddot",
                _ => string.Empty
            };

            if (functionName.Length > 0) {
                AppendFunctionMathText(builder, functionName, expression);
                return;
            }

            builder.Append("accent(");
            builder.Append(accent);
            builder.Append(',');
            builder.Append(expression);
            builder.Append(')');
        }

        private static void AppendDelimiterMathText(StringBuilder builder, OpenXmlElement element) {
            string begin = ReadMathCharacterValue(element, "begChr");
            string end = ReadMathCharacterValue(element, "endChr");
            builder.Append(begin.Length == 0 ? "(" : begin);
            AppendJoinedMathChildText(builder, element, "e", ",");
            builder.Append(end.Length == 0 ? ")" : end);
        }

        private static void AppendGroupCharMathText(StringBuilder builder, OpenXmlElement element) {
            string expression = ReadMathChildText(element, "e");
            string character = ReadMathCharacterValue(element, "chr");
            string functionName = character switch {
                "\u23de" => "overbrace",
                "\u23df" => "underbrace",
                "\u23b4" => "overbracket",
                "\u23b5" => "underbracket",
                _ => "group"
            };
            AppendFunctionMathText(builder, functionName, expression);
        }

        private static void AppendMatrixMathText(StringBuilder builder, OpenXmlElement element) {
            builder.Append("matrix(");
            bool firstRow = true;
            foreach (OpenXmlElement row in FindMathChildren(element, "mr")) {
                if (!firstRow) {
                    builder.Append(';');
                }

                bool firstCell = true;
                foreach (OpenXmlElement cell in FindMathChildren(row, "e")) {
                    if (!firstCell) {
                        builder.Append(',');
                    }

                    AppendVisibleText(builder, cell);
                    firstCell = false;
                }

                firstRow = false;
            }

            builder.Append(')');
        }

        private static void AppendEquationArrayMathText(StringBuilder builder, OpenXmlElement element) {
            builder.Append("eqarray(");
            AppendJoinedMathChildText(builder, element, "e", ";");
            builder.Append(')');
        }

        private static void AppendFunctionMathText(StringBuilder builder, string functionName, string expression) {
            builder.Append(functionName);
            builder.Append('(');
            builder.Append(expression);
            builder.Append(')');
        }

        private static void AppendNaryMathText(StringBuilder builder, OpenXmlElement element) {
            string operatorText = element.LocalName == "int" ? "int" : ReadNaryOperatorText(element);
            string subscript = ReadMathChildText(element, "sub");
            string superscript = ReadMathChildText(element, "sup");
            string expression = ReadMathChildText(element, "e");
            builder.Append(operatorText);
            AppendMathScript(builder, "_", subscript);
            AppendMathScript(builder, "^", superscript);
            if (expression.Length > 0) {
                builder.Append('(');
                builder.Append(expression);
                builder.Append(')');
            }
        }

        private static void AppendJoinedMathChildText(StringBuilder builder, OpenXmlElement element, string localName, string separator) {
            bool first = true;
            foreach (OpenXmlElement child in FindMathChildren(element, localName)) {
                if (!first) {
                    builder.Append(separator);
                }

                AppendVisibleText(builder, child);
                first = false;
            }
        }

        private static string ReadMathCharacterValue(OpenXmlElement element, string localName) {
            OpenXmlElement? character = FindFirstMathDescendant(element, localName);
            string? value = character?.GetAttribute("val", MathNamespace).Value;
            if (string.IsNullOrEmpty(value)) {
                value = character?.GetAttribute("val", string.Empty).Value;
            }

            return value ?? string.Empty;
        }

        private static string ReadNaryOperatorText(OpenXmlElement element) {
            OpenXmlElement? chr = FindFirstMathDescendant(element, "chr");
            string? value = chr?.GetAttribute("val", MathNamespace).Value;
            if (string.IsNullOrEmpty(value)) {
                value = chr?.GetAttribute("val", string.Empty).Value;
            }

            if (string.IsNullOrEmpty(value)) {
                return "sum";
            }

            string operatorText = value ?? "sum";
            return operatorText switch {
                "\u2211" => "sum",
                "\u220F" => "prod",
                "\u222B" => "int",
                _ => operatorText
            };
        }

        private static void AppendMathScript(StringBuilder builder, string marker, string value) {
            if (value.Length == 0) {
                return;
            }

            builder.Append(marker);
            builder.Append('(');
            builder.Append(value);
            builder.Append(')');
        }

        private static void AppendMathChildText(StringBuilder builder, OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstMathChild(element, localName);
            if (child != null) {
                AppendVisibleText(builder, child);
            }
        }

        private static string ReadMathChildText(OpenXmlElement element, string localName) {
            OpenXmlElement? child = FindFirstMathChild(element, localName);
            if (child == null) {
                return string.Empty;
            }

            var builder = new StringBuilder();
            AppendVisibleText(builder, child);
            return builder.ToString();
        }

        private static IEnumerable<OpenXmlElement> FindMathChildren(OpenXmlElement element, string localName) {
            foreach (OpenXmlElement child in element.ChildElements) {
                if (child.NamespaceUri == MathNamespace && child.LocalName == localName) {
                    yield return child;
                }
            }
        }

        private static OpenXmlElement? FindFirstMathChild(OpenXmlElement element, string localName) {
            foreach (OpenXmlElement child in element.ChildElements) {
                if (child.NamespaceUri == MathNamespace && child.LocalName == localName) {
                    return child;
                }
            }

            return null;
        }

        private static OpenXmlElement? FindFirstMathDescendant(OpenXmlElement element, string localName) {
            foreach (OpenXmlElement child in element.ChildElements) {
                if (child.NamespaceUri == MathNamespace && child.LocalName == localName) {
                    return child;
                }

                OpenXmlElement? descendant = FindFirstMathDescendant(child, localName);
                if (descendant != null) {
                    return descendant;
                }
            }

            return null;
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
