using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Provides a fluent API wrapper around <see cref="WordDocument"/>.
    /// </summary>
    public class WordFluentDocument {
        internal WordDocument Document { get; }

        public WordFluentDocument(WordDocument document) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
        }

        public InfoBuilder Info => new InfoBuilder(this);
        public SectionBuilder Sections => new SectionBuilder(this);
        public PageBuilder Pages => new PageBuilder(this);
        public ParagraphBuilder Paragraphs => new ParagraphBuilder(this);
        public RunBuilder Runs => new RunBuilder(this);
        public ListBuilder Lists => new ListBuilder(this);
        public TableBuilder Tables => new TableBuilder(this);
        public ImageBuilder Images => new ImageBuilder(this);
        public HeadersBuilder Headers => new HeadersBuilder(this);
        public FootersBuilder Footers => new FootersBuilder(this);

        public WordFluentDocument ForEachParagraph(Action<ParagraphBuilder> action) {
            Document.ForEachParagraph(p => action(new ParagraphBuilder(this, p)));
            return this;
        }

        public WordFluentDocument Find(string text, Action<ParagraphBuilder> action, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            foreach (var paragraph in Document.FindParagraphs(text, stringComparison)) {
                action(new ParagraphBuilder(this, paragraph));
            }
            return this;
        }

        public IEnumerable<ParagraphBuilder> Select(Func<ParagraphBuilder, bool> predicate) {
            foreach (var paragraph in Document.SelectParagraphs(p => predicate(new ParagraphBuilder(this, p)))) {
                yield return new ParagraphBuilder(this, paragraph);
            }
        }
    }
}
