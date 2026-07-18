using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

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
        private DocumentFormat.OpenXml.Math.OfficeMath? _officeMath;
        private readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;
        private SimpleField? _simpleField;
        private List<Run>? _runs;

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

        /// <summary>Projects this equation into the reusable OfficeIMO math expression tree.</summary>
        public OfficeMathExpression ToExpression() => MathElement != null
            ? WordMath.ToExpression(MathElement)
            : OfficeIMO.Drawing.OfficeMath.Text(Text);

        /// <summary>Renders this equation through the shared dependency-free math renderer.</summary>
        public OfficeDrawing ToDrawing(OfficeMathRenderOptions? options = null) => OfficeMathRenderer.Render(ToExpression(), options);

        /// <summary>Replaces this equation with an editable OMML representation of a shared expression.</summary>
        public WordEquation SetExpression(OfficeMathExpression expression) {
            if (expression == null) throw new ArgumentNullException(nameof(expression));
            var replacement = new DocumentFormat.OpenXml.Math.OfficeMath(WordMathMarkup.ToOmml(expression));
            if (_officeMath != null) {
                _officeMath.RemoveAllChildren();
                foreach (OpenXmlElement child in replacement.ChildElements.ToList()) _officeMath.Append(child.CloneNode(true));
                CopyCompatibilityMetadata(replacement, _officeMath);
                return this;
            }
            if (_mathParagraph != null) {
                OpenXmlElement[] existingEquations = _mathParagraph.ChildElements
                    .Where(child => child.LocalName == "oMath")
                    .ToArray();
                if (existingEquations.Length > 0) {
                    _mathParagraph.InsertBefore(replacement, existingEquations[0]);
                    foreach (OpenXmlElement existing in existingEquations) existing.Remove();
                } else {
                    OpenXmlElement? paragraphProperties = _mathParagraph.ChildElements
                        .FirstOrDefault(child => child.LocalName == "oMathParaPr");
                    if (paragraphProperties != null) _mathParagraph.InsertAfter(replacement, paragraphProperties);
                    else _mathParagraph.Append(replacement);
                }
                return this;
            }

            OpenXmlElement? firstBacking = (OpenXmlElement?)_simpleField ?? _runs?.FirstOrDefault();
            OpenXmlCompositeElement insertionParent = ResolveReplacementParent(firstBacking);
            OpenXmlElement insertionAnchor = ResolveReplacementAnchor(insertionParent, firstBacking);
            insertionParent.InsertBefore(replacement, insertionAnchor);
            _simpleField?.Remove();
            if (_runs != null) foreach (Run run in _runs.ToList()) run.Remove();
            _officeMath = replacement;
            _simpleField = null;
            _runs = null;
            return this;
        }

        private static void CopyCompatibilityMetadata(OpenXmlElement source, OpenXmlElement target) {
            foreach (KeyValuePair<string, string> declaration in source.NamespaceDeclarations) {
                if (target.NamespaceDeclarations.Any(existing => existing.Key == declaration.Key)) {
                    target.RemoveNamespaceDeclaration(declaration.Key);
                }
                target.AddNamespaceDeclaration(declaration.Key, declaration.Value);
            }
            string? sourceIgnorable = source.MCAttributes?.Ignorable?.Value;
            if (string.IsNullOrWhiteSpace(sourceIgnorable)) return;
            MarkupCompatibilityAttributes attributes = target.MCAttributes ?? new MarkupCompatibilityAttributes();
            attributes.Ignorable = string.Join(" ",
                ((attributes.Ignorable?.Value ?? string.Empty) + " " + sourceIgnorable)
                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Distinct(StringComparer.Ordinal));
            target.MCAttributes = attributes;
        }

        private OpenXmlCompositeElement ResolveReplacementParent(OpenXmlElement? firstBacking) {
            if (firstBacking?.Parent is OpenXmlCompositeElement directParent &&
                (_runs == null || _runs.All(run => ReferenceEquals(run.Parent, directParent)))) {
                return directParent;
            }
            return _paragraph;
        }

        private OpenXmlElement ResolveReplacementAnchor(OpenXmlCompositeElement parent, OpenXmlElement? firstBacking) {
            OpenXmlElement? anchor = firstBacking;
            while (anchor != null && !ReferenceEquals(anchor.Parent, parent)) anchor = anchor.Parent;
            return anchor ?? throw new InvalidOperationException("The EQ field is detached from its paragraph.");
        }

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
            IReadOnlyList<WordEquationOccurrence> occurrences,
            Func<OpenXmlElement, bool>? includeElement = null) {
            if (container == null) throw new ArgumentNullException(nameof(container));
            if (occurrences == null) throw new ArgumentNullException(nameof(occurrences));

            var segments = new List<WordEquationContentSegment>();
            var emittedEquations = new HashSet<WordEquation>();
            var emittedControlArtifacts = new HashSet<SdtRun>();
            AppendVisibleContentSegments(
                segments,
                container,
                occurrences,
                emittedEquations,
                emittedControlArtifacts,
                includeElement,
                null,
                null);
            return segments;
        }

        internal static string GetVisibleTextOutsideEquations(
            OpenXmlElement container,
            IReadOnlyList<WordEquationOccurrence> occurrences) =>
            string.Concat(GetVisibleContentSegments(container, occurrences)
                .Where(segment => segment.Equation == null)
                .Select(segment => segment.VisibleText));

        internal static string GetVisibleTextWithEquations(
            OpenXmlElement container,
            IReadOnlyList<WordEquationOccurrence> occurrences,
            Func<OpenXmlElement, bool>? includeElement = null) =>
            string.Concat(GetVisibleContentSegments(container, occurrences, includeElement)
                .Select(segment => segment.Equation?.Text ?? segment.VisibleText));

        internal static WordEquation? GetFirstOccurrenceForContainer(
            WordDocument document,
            Paragraph paragraph,
            OpenXmlElement container) {
            OpenXmlElement? directChild = GetDirectParagraphChild(paragraph, container);
            if (directChild == null) return null;

            int childIndex = IndexOfReference(paragraph.ChildElements.ToList(), directChild);
            return GetOccurrences(document, paragraph)
                .FirstOrDefault(occurrence => occurrence.ContainsChildIndex(childIndex))?
                .Equation;
        }

        private static void AppendVisibleContentSegments(
            List<WordEquationContentSegment> segments,
            OpenXmlElement element,
            IReadOnlyList<WordEquationOccurrence> occurrences,
            HashSet<WordEquation> emittedEquations,
            HashSet<SdtRun> emittedControlArtifacts,
            Func<OpenXmlElement, bool>? includeElement,
            OpenXmlElement? sourceElement,
            SdtRun? artifactControl) {
            if (element is DeletedRun || element is MoveFromRun) return;
            if (includeElement != null && !includeElement(element)) return;
            if (element is Hyperlink || element is SdtRun) sourceElement = element;
            if (element is Run run) {
                sourceElement = run;
            } else if (element is SdtRun sdtRun && HasSupportedSdtArtifact(sdtRun)) {
                artifactControl = sdtRun;
            }

            WordEquation? backedEquation = occurrences
                .Select(occurrence => occurrence.Equation)
                .FirstOrDefault(equation => equation.IsBackingElement(element));
            if (backedEquation != null) {
                if (emittedEquations.Add(backedEquation)) {
                    segments.Add(WordEquationContentSegment.FromEquation(backedEquation, sourceElement));
                }
                return;
            }

            if (IsSupportedRunArtifactElement(element) && sourceElement is Run artifactRun) {
                segments.Add(WordEquationContentSegment.FromRunArtifact(artifactRun, element));
                return;
            }

            if (element is Text text) {
                AppendVisibleTextOrControlArtifact(
                    segments,
                text.Text,
                sourceElement,
                artifactControl,
                emittedControlArtifacts,
                occurrences);
                return;
            }
            if (element is TabChar) {
                AppendVisibleTextOrControlArtifact(
                    segments,
                    "\t",
                    sourceElement,
                    artifactControl,
                    emittedControlArtifacts,
                    occurrences);
                return;
            }
            if (element is NoBreakHyphen) {
                AppendVisibleTextOrControlArtifact(
                    segments,
                    "\u2011",
                    sourceElement,
                    artifactControl,
                    emittedControlArtifacts,
                    occurrences);
                return;
            }
            if (element is SoftHyphen) {
                AppendVisibleTextOrControlArtifact(
                    segments,
                    "\u00ad",
                    sourceElement,
                    artifactControl,
                    emittedControlArtifacts,
                    occurrences);
                return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendVisibleContentSegments(
                    segments,
                    child,
                    occurrences,
                    emittedEquations,
                    emittedControlArtifacts,
                    includeElement,
                    sourceElement,
                    artifactControl);
            }
        }

        private static void AppendVisibleTextOrControlArtifact(
            List<WordEquationContentSegment> segments,
            string? text,
            OpenXmlElement? sourceElement,
            SdtRun? artifactControl,
            HashSet<SdtRun> emittedControlArtifacts,
            IReadOnlyList<WordEquationOccurrence> occurrences) {
            if (string.IsNullOrEmpty(text)) return;
            if (artifactControl != null) {
                if (!emittedControlArtifacts.Add(artifactControl)) return;
                string visibleArtifactText = string.Concat(artifactControl
                    .Descendants<Text>()
                    .Where(value => !IsInsideEquationBackingElement(value, occurrences))
                    .Select(value => value.Text));
                segments.Add(WordEquationContentSegment.FromRunArtifact(
                    artifactControl,
                    artifactControl,
                    string.IsNullOrEmpty(visibleArtifactText) ? text : visibleArtifactText));
                return;
            }

            AppendVisibleTextSegment(segments, text, sourceElement);
        }

        private static bool IsSupportedRunArtifactElement(OpenXmlElement element) =>
            element is Break ||
            element is CarriageReturn ||
            element is FootnoteReference ||
            element is EndnoteReference ||
            element is CommentReference ||
            element is DocumentFormat.OpenXml.Wordprocessing.Drawing ||
            element is DocumentFormat.OpenXml.Vml.ImageData;

        private static bool HasSupportedSdtArtifact(SdtRun sdtRun) =>
            sdtRun.SdtProperties?.Elements<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>().Any() == true ||
            sdtRun.SdtProperties?.Elements<SdtContentDropDownList>().Any() == true ||
            sdtRun.SdtProperties?.Elements<SdtContentComboBox>().Any() == true ||
            sdtRun.SdtProperties?.Elements<SdtContentDate>().Any() == true;

        internal static bool IsVisibleEquationContentContainer(OpenXmlElement element) =>
            element is Hyperlink ||
            element is SdtRun ||
            element is InsertedRun ||
            element is MoveToRun;

        private static bool IsInsideEquationBackingElement(
            OpenXmlElement element,
            IReadOnlyList<WordEquationOccurrence> occurrences) {
            for (OpenXmlElement? current = element.Parent; current != null; current = current.Parent) {
                if (occurrences.Any(occurrence => occurrence.Equation.IsBackingElement(current))) {
                    return true;
                }
            }
            return false;
        }

        private static void AppendVisibleTextSegment(List<WordEquationContentSegment> segments, string? text, OpenXmlElement? sourceElement) {
            if (string.IsNullOrEmpty(text)) return;
            if (segments.Count > 0 &&
                segments[segments.Count - 1].Equation == null &&
                !segments[segments.Count - 1].IsRunArtifact &&
                ReferenceEquals(segments[segments.Count - 1].SourceElement, sourceElement)) {
                WordEquationContentSegment previous = segments[segments.Count - 1];
                segments[segments.Count - 1] = WordEquationContentSegment.FromText((previous.Text ?? string.Empty) + text, sourceElement);
                return;
            }

            segments.Add(WordEquationContentSegment.FromText(text!, sourceElement));
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
            start = GetDirectParagraphChild(_paragraph, start);
            end = GetDirectParagraphChild(_paragraph, end);
            startIndex = start == null ? -1 : IndexOfReference(children, start);
            endIndex = end == null ? -1 : IndexOfReference(children, end);
            return startIndex >= 0 && endIndex >= startIndex;
        }

        internal static OpenXmlElement? GetDirectParagraphChild(Paragraph paragraph, OpenXmlElement? element) {
            while (element != null && !ReferenceEquals(element.Parent, paragraph)) {
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
        private WordEquationContentSegment(
            string? text,
            WordEquation? equation,
            OpenXmlElement? sourceElement,
            OpenXmlElement? artifactElement,
            string? artifactVisibleText = null) {
            Text = text;
            Equation = equation;
            SourceElement = sourceElement;
            ArtifactElement = artifactElement;
            ArtifactVisibleText = artifactVisibleText;
        }

        internal string? Text { get; }
        internal WordEquation? Equation { get; }
        internal OpenXmlElement? SourceElement { get; }
        internal Run? SourceRun => SourceElement as Run;
        internal OpenXmlElement? ArtifactElement { get; }
        internal string? ArtifactVisibleText { get; }
        internal bool IsRunArtifact => ArtifactElement != null;
        internal string VisibleText {
            get {
                if (Text != null) return Text;
                if (!IsRunArtifact || SourceElement == null) return string.Empty;
                if (ArtifactElement is Break || ArtifactElement is CarriageReturn) {
                    return "\n";
                }
                return ArtifactVisibleText ?? string.Empty;
            }
        }

        internal WordParagraph CreateSourceParagraph(WordDocument document, Paragraph paragraph, WordParagraph fallback) {
            if (SourceElement == null) return fallback;

            WordParagraph source;
            if (SourceElement is Hyperlink sourceHyperlink) {
                source = new WordParagraph(document, paragraph, sourceHyperlink);
            } else if (SourceElement is SdtRun sourceSdtRun) {
                source = new WordParagraph(document, paragraph, sourceSdtRun);
            } else if (SourceElement is Run sourceRun) {
                source = new WordParagraph(document, paragraph, sourceRun);
            } else {
                return fallback;
            }

            for (OpenXmlElement? ancestor = SourceElement.Parent; ancestor != null; ancestor = ancestor.Parent) {
                if (ancestor is Hyperlink hyperlink) {
                    source._hyperlink = hyperlink;
                    break;
                }
                if (ReferenceEquals(ancestor, paragraph)) break;
            }
            return source;
        }

        internal static WordEquationContentSegment FromText(string text, OpenXmlElement? sourceElement) => new(text, null, sourceElement, null);
        internal static WordEquationContentSegment FromEquation(WordEquation equation, OpenXmlElement? sourceElement) => new(null, equation, sourceElement, null);
        internal static WordEquationContentSegment FromRunArtifact(
            OpenXmlElement sourceElement,
            OpenXmlElement artifactElement,
            string? visibleText = null) => new(null, null, sourceElement, artifactElement, visibleText);
    }
}
