using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>Why a piece of visible text cannot be edited as plain run text.</summary>
    [Flags]
    internal enum WordTextHazard {
        None = 0,
        /// <summary>Text is a field result; a field refresh would overwrite an edit.</summary>
        Field = 1,
        /// <summary>Text belongs to a tracked insertion or move.</summary>
        TrackedChange = 2,
        /// <summary>Text comes from an equation, drawing, embedded object or other non-run content.</summary>
        Opaque = 4
    }

    /// <summary>One visible piece of paragraph text and the element that produces it.</summary>
    internal sealed class WordTextAtom {
        internal WordTextAtom(OpenXmlElement element, Run? run, int start, int length, WordTextHazard hazard) {
            Element = element;
            Run = run;
            Start = start;
            Length = length;
            Hazard = hazard;
        }

        /// <summary>The <c>w:t</c>, tab, break or opaque element producing the text.</summary>
        internal OpenXmlElement Element { get; }

        /// <summary>The run holding <see cref="Element"/>; <c>null</c> for content outside runs, such as equations.</summary>
        internal Run? Run { get; }

        internal int Start { get; }
        internal int Length { get; }
        internal int End => Start + Length;
        internal WordTextHazard Hazard { get; }

        /// <summary>The run, or the element itself when it is not inside a run.</summary>
        internal OpenXmlElement InlineNode => (OpenXmlElement?)Run ?? Element;
    }

    /// <summary>
    /// The visible text of one paragraph, with each character traced to the element that produces it.
    /// Deleted text, field instructions and hidden comment references are not visible text.
    /// </summary>
    internal sealed class WordTextMap {
        private WordTextMap(Paragraph paragraph, string text, List<WordTextAtom> atoms) {
            Paragraph = paragraph;
            Text = text;
            Atoms = atoms;
        }

        internal Paragraph Paragraph { get; }
        internal string Text { get; }
        internal IReadOnlyList<WordTextAtom> Atoms { get; }

        /// <summary>Returns the atom holding the character at <paramref name="offset"/>.</summary>
        internal WordTextAtom? AtomAt(int offset) {
            foreach (WordTextAtom atom in Atoms) {
                if (atom.Length > 0 && atom.Start <= offset && offset < atom.End) return atom;
            }
            return null;
        }

        /// <summary>Builds the map from the paragraph alone; a field opened in an earlier paragraph is seen only through this paragraph's own markers.</summary>
        internal static WordTextMap Build(Paragraph paragraph) => Build(paragraph, null);

        /// <summary>Builds the map for a paragraph that starts inside <paramref name="enteringFieldDepth"/> open fields.</summary>
        /// <param name="paragraph">Paragraph to map.</param>
        /// <param name="enteringFieldDepth">Fields still open at the paragraph start, or <c>null</c> to judge from the paragraph alone.</param>
        internal static WordTextMap Build(Paragraph paragraph, int? enteringFieldDepth) {
            var builder = new Builder();
            if (enteringFieldDepth is int depth) {
                for (int level = 0; level < depth; level++) builder.FieldStack.Push(true);
            } else if (StartsInsideFieldResult(paragraph)) {
                builder.FieldStack.Push(true);
            }
            builder.Walk(paragraph, WordTextHazard.None);
            return new WordTextMap(paragraph, builder.Text.ToString(), builder.Atoms);
        }

        /// <summary>
        /// Counts the fields still open where <paramref name="paragraph"/> starts. A field result can run through several
        /// paragraphs, such as a table of contents, so a paragraph without field markers can still be field text.
        /// Walks the story up to the paragraph; use <see cref="GetEnteringFieldDepths"/> when mapping a whole story.
        /// </summary>
        internal static int GetEnteringFieldDepth(Paragraph paragraph) {
            int depth = 0;
            foreach (OpenXmlElement element in WordRange.GetStoryRoot(paragraph).Descendants()) {
                if (ReferenceEquals(element, paragraph)) return depth;
                if (element is FieldChar fieldChar) depth = Apply(fieldChar, depth);
            }
            return depth;
        }

        /// <summary>Counts the fields open at the start of every paragraph of a story in one pass.</summary>
        internal static void GetEnteringFieldDepths(OpenXmlElement storyRoot, Dictionary<Paragraph, int> depths) {
            int depth = 0;
            foreach (OpenXmlElement element in storyRoot.Descendants()) {
                if (element is Paragraph paragraph) depths[paragraph] = depth;
                else if (element is FieldChar fieldChar) depth = Apply(fieldChar, depth);
            }
        }

        private static int Apply(FieldChar fieldChar, int depth) {
            FieldCharValues? type = fieldChar.FieldCharType?.Value;
            if (type == FieldCharValues.Begin) return depth + 1;
            if (type == FieldCharValues.End && depth > 0) return depth - 1;
            return depth;
        }

        /// <summary>
        /// A field can begin in an earlier paragraph. Text before an unmatched separator or end marker is then a field result.
        /// </summary>
        private static bool StartsInsideFieldResult(Paragraph paragraph) {
            int open = 0;
            foreach (FieldChar fieldChar in paragraph.Descendants<FieldChar>()) {
                FieldCharValues? type = fieldChar.FieldCharType?.Value;
                if (type == FieldCharValues.Begin) {
                    open++;
                } else if (type == FieldCharValues.End) {
                    if (open == 0) return true;
                    open--;
                } else if (type == FieldCharValues.Separate && open == 0) {
                    return true;
                }
            }
            return false;
        }

        internal static bool IsPlainRunContent(OpenXmlElement child) =>
            child is TabChar or PositionalTab or Break or CarriageReturn or NoBreakHyphen or SoftHyphen or LastRenderedPageBreak;

        private sealed class Builder {
            internal readonly StringBuilder Text = new();
            internal readonly List<WordTextAtom> Atoms = new();
            // true once the field on that level has passed its separator.
            internal readonly Stack<bool> FieldStack = new();

            internal void Walk(OpenXmlElement parent, WordTextHazard inherited) {
                foreach (OpenXmlElement child in parent.ChildElements) {
                    switch (child) {
                        case ParagraphProperties:
                        case DeletedRun:
                        case MoveFromRun:
                            continue;
                        case Run run:
                            AddRun(run, inherited);
                            continue;
                        case SimpleField:
                            Walk(child, inherited | WordTextHazard.Field);
                            continue;
                        case InsertedRun:
                        case MoveToRun:
                            Walk(child, inherited | WordTextHazard.TrackedChange);
                            continue;
                        case SdtRun sdtRun:
                            if (sdtRun.SdtProperties?.Elements<W14.SdtContentCheckBox>().Any() == true) continue;
                            if (sdtRun.SdtContentRun != null) Walk(sdtRun.SdtContentRun, inherited);
                            continue;
                        case Hyperlink:
                        case CustomXmlRun:
                            Walk(child, inherited);
                            continue;
                    }

                    if (child is AlternateContent || child.NamespaceUri == WordMath.MathNamespace) {
                        AddAtom(child, null, WordParagraph.ReadVisibleText(child), inherited | WordTextHazard.Opaque);
                    }
                }
            }

            private void AddRun(Run run, WordTextHazard inherited) {
                if (IsHiddenCommentReferenceRun(run)) return;
                bool hasFieldChar = false;
                foreach (OpenXmlElement child in run.ChildElements) {
                    if (child is FieldChar) {
                        hasFieldChar = true;
                        break;
                    }
                }

                foreach (OpenXmlElement child in run.ChildElements) {
                    switch (child) {
                        case RunProperties:
                        case FieldCode:
                        case DeletedFieldCode:
                        case DeletedText:
                        case CommentReference:
                            continue;
                        case FieldChar fieldChar:
                            UpdateFieldState(fieldChar);
                            continue;
                    }

                    WordTextHazard hazard = inherited;
                    if (hasFieldChar || FieldStack.Count > 0) hazard |= WordTextHazard.Field;
                    if (child is Text text) {
                        AddAtom(text, run, WordParagraph.ReadVisibleText(text), hazard);
                    } else if (child is LastRenderedPageBreak) {
                        continue;
                    } else {
                        if (!IsPlainRunContent(child)) hazard |= WordTextHazard.Opaque;
                        AddAtom(child, run, WordParagraph.ReadVisibleText(child), hazard);
                    }
                }
            }

            private void UpdateFieldState(FieldChar fieldChar) {
                FieldCharValues? type = fieldChar.FieldCharType?.Value;
                if (type == FieldCharValues.Begin) {
                    FieldStack.Push(false);
                } else if (type == FieldCharValues.Separate) {
                    if (FieldStack.Count > 0) {
                        FieldStack.Pop();
                        FieldStack.Push(true);
                    }
                } else if (type == FieldCharValues.End && FieldStack.Count > 0) {
                    FieldStack.Pop();
                }
            }

            private void AddAtom(OpenXmlElement element, Run? run, string text, WordTextHazard hazard) {
                Atoms.Add(new WordTextAtom(element, run, Text.Length, text.Length, hazard));
                Text.Append(text);
            }

            private static bool IsHiddenCommentReferenceRun(Run run) {
                Vanish? vanish = run.RunProperties?.GetFirstChild<Vanish>();
                return vanish != null &&
                       (vanish.Val == null || vanish.Val.Value) &&
                       run.Elements<CommentReference>().Any();
            }
        }
    }
}
