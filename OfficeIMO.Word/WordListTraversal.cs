using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents events emitted while traversing document lists.
    /// </summary>
    public readonly struct WordListEvent {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordListEvent"/> struct.
        /// </summary>
        /// <param name="eventType">Type of event being reported.</param>
        /// <param name="paragraph">Paragraph associated with the event when applicable.</param>
        /// <param name="ordered"><c>true</c> when the current list is ordered.</param>
        /// <param name="level">Zero-based nesting level of the list.</param>
        public WordListEvent(WordListEventType eventType, Paragraph? paragraph, bool ordered, int level) {
            EventType = eventType;
            Paragraph = paragraph;
            Ordered = ordered;
            Level = level;
        }

        /// <summary>
        /// Type of the emitted event.
        /// </summary>
        public WordListEventType EventType { get; }

        /// <summary>
        /// Paragraph associated with the event when applicable.
        /// </summary>
        public Paragraph? Paragraph { get; }

        /// <summary>
        /// Indicates whether the current list is ordered.
        /// </summary>
        public bool Ordered { get; }

        /// <summary>
        /// Zero-based nesting level of the list item.
        /// </summary>
        public int Level { get; }
    }

    /// <summary>
    /// Types of list traversal events.
    /// </summary>
    public enum WordListEventType {
        /// <summary>Signals the start of a new list.</summary>
        StartList,
        /// <summary>Signals the end of the current list.</summary>
        EndList,
        /// <summary>Signals the start of a new list item.</summary>
        StartItem,
        /// <summary>Signals the end of the current list item.</summary>
        EndItem,
        /// <summary>Represents a standalone paragraph outside of lists.</summary>
        Paragraph
    }

    /// <summary>
    /// Provides utilities to traverse list structures within a document.
    /// </summary>
    public static class WordListTraversal {
        /// <summary>
        /// Traverses the body of the provided document emitting events for list structures
        /// and standalone paragraphs. This helper enables converters to reuse list handling
        /// logic without duplicating stack management code.
        /// </summary>
        /// <param name="document">Document to traverse.</param>
        /// <returns>Sequence of events describing lists and paragraphs.</returns>
        public static IEnumerable<WordListEvent> Traverse(WordprocessingDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            Dictionary<int, bool> listTypes = GetListTypes(document);
            Stack<(int numId, bool ordered)> listStack = new Stack<(int numId, bool ordered)>();
            var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("The document does not contain a main document part.");
            var body = mainPart.Document?.Body;
            if (body == null) {
                yield break;
            }

            static bool CurrentOrdered(Stack<(int numId, bool ordered)> stack) => stack.Count > 0 && stack.Peek().ordered;

            List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();
            int previousLevel = -1;

            foreach (Paragraph paragraph in paragraphs) {
                NumberingProperties? numProps = paragraph.ParagraphProperties?.NumberingProperties;
                if (numProps != null) {
                    int level = numProps.NumberingLevelReference?.Val ?? 0;
                    int numId = numProps.NumberingId?.Val ?? 0;
                    bool ordered = listTypes.ContainsKey(numId) && listTypes[numId];

                    if (previousLevel == -1) {
                        for (int lvl = 0; lvl <= level; lvl++) {
                            yield return new WordListEvent(WordListEventType.StartList, null, ordered, lvl);
                            listStack.Push((numId, ordered));
                        }
                    } else if (level > previousLevel) {
                        for (int lvl = previousLevel + 1; lvl <= level; lvl++) {
                            yield return new WordListEvent(WordListEventType.StartList, null, ordered, lvl);
                            listStack.Push((numId, ordered));
                        }
                    } else {
                        yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), previousLevel);
                        for (int lvl = previousLevel; lvl > level; lvl--) {
                            var closing = listStack.Pop();
                            yield return new WordListEvent(WordListEventType.EndList, null, closing.ordered, lvl);
                            if (listStack.Count > 0) {
                                yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), lvl - 1);
                            }
                        }
                        if (listStack.Count > 0 && listStack.Peek().numId != numId) {
                            var closing = listStack.Pop();
                            yield return new WordListEvent(WordListEventType.EndList, null, closing.ordered, level);
                            if (listStack.Count > 0) {
                                yield return new WordListEvent(WordListEventType.EndItem, null, listStack.Peek().ordered, level - 1);
                            }
                        }
                        if (listStack.Count <= level) {
                            yield return new WordListEvent(WordListEventType.StartList, null, ordered, level);
                            listStack.Push((numId, ordered));
                        }
                }

                yield return new WordListEvent(WordListEventType.StartItem, paragraph, ordered, level);
                previousLevel = level;
                continue;
            }

            if (previousLevel != -1) {
                yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), previousLevel);
                while (listStack.Count > 0) {
                    var closing = listStack.Pop();
                    yield return new WordListEvent(WordListEventType.EndList, null, closing.ordered, previousLevel);
                    if (listStack.Count > 0) {
                        yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), previousLevel - 1);
                    }
                }
                previousLevel = -1;
            }

            yield return new WordListEvent(WordListEventType.Paragraph, paragraph, false, 0);
        }

        if (previousLevel != -1) {
            yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), previousLevel);
            while (listStack.Count > 0) {
                var closing = listStack.Pop();
                yield return new WordListEvent(WordListEventType.EndList, null, closing.ordered, previousLevel);
                if (listStack.Count > 0) {
                    yield return new WordListEvent(WordListEventType.EndItem, null, CurrentOrdered(listStack), previousLevel - 1);
                }
            }
        }
        }

        /// <summary>
        /// Builds a map of numbering identifiers to list type (ordered or bullet).
        /// </summary>
        /// <param name="document">Document containing numbering definitions.</param>
        /// <returns>Dictionary mapping numbering IDs to <c>true</c> when ordered, <c>false</c> when bullet.</returns>
        private static Dictionary<int, bool> GetListTypes(WordprocessingDocument document) {
            Dictionary<int, bool> listTypes = new Dictionary<int, bool>();

            var mainPart = document.MainDocumentPart;
            var numberingPart = mainPart?.NumberingDefinitionsPart;
            if (numberingPart?.Numbering == null) {
                return listTypes;
            }

            foreach (NumberingInstance instance in numberingPart.Numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value is not int id) {
                    continue;
                }

                if (instance.AbstractNumId?.Val?.Value is not int absId) {
                    continue;
                }

                AbstractNum? abs = numberingPart.Numbering.Elements<AbstractNum>()
                    .FirstOrDefault(a => a.AbstractNumberId?.Value == absId);
                if (abs == null) {
                    continue;
                }

                bool ordered = true;
                Level? lvl = abs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
                NumberFormatValues? format = lvl?.NumberingFormat?.Val?.Value;
                if (format == NumberFormatValues.Bullet) {
                    ordered = false;
                }
                listTypes[id] = ordered;
            }

            return listTypes;
        }
    }
}