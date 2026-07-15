using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Identifies the native representation that backs a Word equation.</summary>
    public enum WordEquationRepresentation {
        /// <summary>Office Math Markup Language used by DOCX.</summary>
        Omml,
        /// <summary>Word EQ field used by legacy DOC and RTF.</summary>
        EquationField
    }

    /// <summary>
    /// Encapsulates an OMML equation or a legacy Word <c>EQ</c> field.
    /// </summary>
    public class WordEquation : WordElement {
        private readonly WordDocument _document;
        private readonly Paragraph _paragraph;
        private readonly DocumentFormat.OpenXml.Math.OfficeMath? _officeMath;
        private readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;
        private readonly SimpleField? _simpleField;
        private readonly List<Run>? _runs;

        /// <summary>Initializes an equation backed by an Office Math element.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;
            _officeMath = officeMath;
        }

        /// <summary>Initializes an equation backed by an Office Math paragraph.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _mathParagraph = mathParagraph;
        }

        /// <summary>Initializes an equation backed by an Office Math element and its paragraph.</summary>
        public WordEquation(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _officeMath = officeMath;
            _mathParagraph = mathParagraph;
        }

        internal WordEquation(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;
            _simpleField = simpleField;
        }

        internal WordEquation(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;
            _runs = runs;
        }

        /// <summary>Gets the native equation representation.</summary>
        public WordEquationRepresentation Representation => MathElement == null
            ? WordEquationRepresentation.EquationField
            : WordEquationRepresentation.Omml;

        /// <summary>
        /// Gets or sets deterministic visible equation text. Setting OMML text replaces the structured
        /// expression with one math text run; setting an EQ field updates its cached display result.
        /// </summary>
        public string Text {
            get {
                OpenXmlElement? math = MathElement;
                if (math != null) return WordMath.GetText(math);
                return FieldParagraph?.Text ?? string.Empty;
            }
            set {
                OpenXmlElement? math = MathElement;
                if (math != null) {
                    WordMath.SetText(math, value);
                    return;
                }
                if (FieldParagraph != null) FieldParagraph.Text = value;
            }
        }

        /// <summary>Gets raw OMML when the equation is backed by DOCX math markup; otherwise <c>null</c>.</summary>
        public string? Omml => MathElement?.OuterXml;

        /// <summary>Gets the raw EQ field instruction when the equation is field-backed; otherwise <c>null</c>.</summary>
        public string? FieldInstruction => FieldParagraph?.Field?.Field;

        /// <summary>Projects the equation to LaTeX. Field-backed equations fall back to their cached display text.</summary>
        public string ToLatex() => MathElement != null ? WordMath.ToLatex(MathElement) : Text;

        /// <summary>Projects the equation to MathML. Field-backed equations use a safe <c>mtext</c> fallback.</summary>
        public string ToMathMl() => MathElement != null
            ? WordMath.ToMathMl(MathElement)
            : $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mtext>{EscapeXml(Text)}</mtext></math>";

        /// <summary>Projects OMML to a legacy Word EQ instruction, or returns the existing EQ instruction.</summary>
        public string ToEquationFieldInstruction() => MathElement != null
            ? WordMath.ToEquationFieldInstruction(MathElement)
            : FieldInstruction ?? " EQ ";

        /// <summary>Removes the equation and its backing markup from the document.</summary>
        public void Remove() {
            _officeMath?.Remove();
            _mathParagraph?.Remove();
            _simpleField?.Remove();
            if (_runs != null) {
                foreach (Run run in _runs.ToList()) run.Remove();
            }
        }

        internal OpenXmlElement? MathElement => (OpenXmlElement?)_officeMath ?? _mathParagraph;

        internal static IReadOnlyList<WordEquationOccurrence> GetOccurrences(WordDocument document, Paragraph paragraph) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));

            List<OpenXmlElement> children = paragraph.ChildElements.ToList();
            var occurrences = new List<WordEquationOccurrence>();
            foreach (WordEquation equation in GetVisibleEquations(document, paragraph, paragraph)) {
                if (equation.TryGetChildRange(children, out int startIndex, out int endIndex)) {
                    occurrences.Add(new WordEquationOccurrence(equation, startIndex, endIndex));
                }
            }

            return occurrences
                .OrderBy(occurrence => occurrence.StartChildIndex)
                .ToList();
        }

        internal static IReadOnlyList<WordEquationContentSegment> GetVisibleContentSegments(
            OpenXmlElement container,
            IReadOnlyList<WordEquationOccurrence> occurrences) {
            if (container == null) throw new ArgumentNullException(nameof(container));
            if (occurrences == null) throw new ArgumentNullException(nameof(occurrences));

            var segments = new List<WordEquationContentSegment>();
            var emittedEquations = new HashSet<WordEquation>();
            AppendVisibleContentSegments(segments, container, occurrences, emittedEquations);
            return segments;
        }

        internal static string GetVisibleTextOutsideEquations(
            OpenXmlElement container,
            IReadOnlyList<WordEquationOccurrence> occurrences) =>
            string.Concat(GetVisibleContentSegments(container, occurrences)
                .Where(segment => segment.Equation == null)
                .Select(segment => segment.Text));

        internal static string GetVisibleTextWithEquations(
            OpenXmlElement container,
            IReadOnlyList<WordEquationOccurrence> occurrences) =>
            string.Concat(GetVisibleContentSegments(container, occurrences)
                .Select(segment => segment.Equation?.Text ?? segment.Text));

        private static void AppendVisibleContentSegments(
            List<WordEquationContentSegment> segments,
            OpenXmlElement element,
            IReadOnlyList<WordEquationOccurrence> occurrences,
            HashSet<WordEquation> emittedEquations) {
            if (element is DeletedRun || element is MoveFromRun) return;

            WordEquation? backedEquation = occurrences
                .Select(occurrence => occurrence.Equation)
                .FirstOrDefault(equation => equation.IsBackingElement(element));
            if (backedEquation != null) {
                if (emittedEquations.Add(backedEquation)) {
                    segments.Add(WordEquationContentSegment.FromEquation(backedEquation));
                }
                return;
            }

            if (element is Text text) {
                AppendVisibleTextSegment(segments, text.Text);
                return;
            }
            if (element is TabChar) {
                AppendVisibleTextSegment(segments, "\t");
                return;
            }
            if (element is Break || element is CarriageReturn) {
                AppendVisibleTextSegment(segments, "\n");
                return;
            }
            if (element is NoBreakHyphen) {
                AppendVisibleTextSegment(segments, "\u2011");
                return;
            }
            if (element is SoftHyphen) {
                AppendVisibleTextSegment(segments, "\u00ad");
                return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendVisibleContentSegments(segments, child, occurrences, emittedEquations);
            }
        }

        private static void AppendVisibleTextSegment(List<WordEquationContentSegment> segments, string? text) {
            if (string.IsNullOrEmpty(text)) return;
            if (segments.Count > 0 && segments[segments.Count - 1].Equation == null) {
                WordEquationContentSegment previous = segments[segments.Count - 1];
                segments[segments.Count - 1] = WordEquationContentSegment.FromText((previous.Text ?? string.Empty) + text);
                return;
            }

            segments.Add(WordEquationContentSegment.FromText(text!));
        }

        private bool IsBackingElement(OpenXmlElement element) =>
            ReferenceEquals(element, _mathParagraph) ||
            ReferenceEquals(element, _officeMath) ||
            ReferenceEquals(element, _simpleField) ||
            (_runs?.Any(run => ReferenceEquals(run, element)) ?? false);

        private static IEnumerable<WordEquation> GetVisibleEquations(
            WordDocument document,
            Paragraph paragraph,
            OpenXmlElement container) {
            var candidates = new List<(int Order, WordEquation Equation)>();
            var complexFields = new List<(int Order, List<Run> Runs)>();
            int order = 0;

            foreach (OpenXmlElement element in EnumerateVisibleInlineContent(container)) {
                int elementOrder = order++;
                if (element is DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
                    candidates.Add((elementOrder, new WordEquation(document, paragraph, mathParagraph)));
                    continue;
                }

                if (element is DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
                    candidates.Add((elementOrder, new WordEquation(document, paragraph, officeMath)));
                    continue;
                }

                if (element is SimpleField simpleField) {
                    var field = new WordField(document, paragraph, simpleField, null);
                    if (field.FieldType == WordFieldType.EQ) {
                        candidates.Add((elementOrder, new WordEquation(document, paragraph, simpleField)));
                    }
                    continue;
                }

                if (element is not Run run) continue;

                foreach ((int _, List<Run> runs) in complexFields) runs.Add(run);
                FieldChar? fieldCharacter = run.Elements<FieldChar>().FirstOrDefault();
                if (fieldCharacter?.FieldCharType?.Value == FieldCharValues.Begin) {
                    complexFields.Add((elementOrder, new List<Run> { run }));
                } else if (fieldCharacter?.FieldCharType?.Value == FieldCharValues.End && complexFields.Count > 0) {
                    (int fieldOrder, List<Run> fieldRuns) = complexFields[complexFields.Count - 1];
                    complexFields.RemoveAt(complexFields.Count - 1);
                    var field = new WordField(document, paragraph, null, fieldRuns);
                    if (field.FieldType == WordFieldType.EQ) {
                        candidates.Add((fieldOrder, new WordEquation(document, paragraph, fieldRuns)));
                    }
                }
            }

            return candidates
                .OrderBy(candidate => candidate.Order)
                .Select(candidate => candidate.Equation);
        }

        private static IEnumerable<OpenXmlElement> EnumerateVisibleInlineContent(OpenXmlElement container) {
            foreach (OpenXmlElement child in container.ChildElements) {
                if (child is DeletedRun || child is MoveFromRun) continue;
                if (child is Run ||
                    child is SimpleField ||
                    child is DocumentFormat.OpenXml.Math.OfficeMath ||
                    child is DocumentFormat.OpenXml.Math.Paragraph) {
                    yield return child;
                    continue;
                }

                foreach (OpenXmlElement nested in EnumerateVisibleInlineContent(child)) {
                    yield return nested;
                }
            }
        }

        private bool TryGetChildRange(IReadOnlyList<OpenXmlElement> children, out int startIndex, out int endIndex) {
            OpenXmlElement? start = (OpenXmlElement?)_mathParagraph
                ?? _officeMath
                ?? _simpleField
                ?? (OpenXmlElement?)_runs?.FirstOrDefault();
            OpenXmlElement? end = (OpenXmlElement?)_runs?.LastOrDefault() ?? start;
            start = GetDirectParagraphChild(start);
            end = GetDirectParagraphChild(end);
            startIndex = start == null ? -1 : IndexOfReference(children, start);
            endIndex = end == null ? -1 : IndexOfReference(children, end);
            return startIndex >= 0 && endIndex >= startIndex;
        }

        private OpenXmlElement? GetDirectParagraphChild(OpenXmlElement? element) {
            while (element != null && !ReferenceEquals(element.Parent, _paragraph)) {
                element = element.Parent;
            }

            return element;
        }

        private static int IndexOfReference(IReadOnlyList<OpenXmlElement> elements, OpenXmlElement target) {
            for (int i = 0; i < elements.Count; i++) {
                if (ReferenceEquals(elements[i], target)) return i;
            }

            return -1;
        }

        private WordParagraph? FieldParagraph => _simpleField != null
            ? new WordParagraph(_document, _paragraph, _simpleField)
            : _runs != null ? new WordParagraph(_document, _paragraph, _runs) : null;

        private static string EscapeXml(string value) => value
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
    }

    internal sealed class WordEquationOccurrence {
        internal WordEquationOccurrence(WordEquation equation, int startChildIndex, int endChildIndex) {
            Equation = equation;
            StartChildIndex = startChildIndex;
            EndChildIndex = endChildIndex;
        }

        internal WordEquation Equation { get; }
        internal int StartChildIndex { get; }
        internal int EndChildIndex { get; }

        internal bool ContainsChildIndex(int childIndex) =>
            childIndex >= StartChildIndex && childIndex <= EndChildIndex;
    }

    internal sealed class WordEquationContentSegment {
        private WordEquationContentSegment(string? text, WordEquation? equation) {
            Text = text;
            Equation = equation;
        }

        internal string? Text { get; }
        internal WordEquation? Equation { get; }

        internal static WordEquationContentSegment FromText(string text) => new(text, null);
        internal static WordEquationContentSegment FromEquation(WordEquation equation) => new(null, equation);
    }
}
