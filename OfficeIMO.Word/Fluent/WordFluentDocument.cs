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

        public WordFluentDocument Info(Action<InfoBuilder> action) {
            action(new InfoBuilder(this));
            return this;
        }

        public WordFluentDocument Section(Action<SectionBuilder> action) {
            action(new SectionBuilder(this));
            return this;
        }

        public WordFluentDocument Page(Action<PageBuilder> action) {
            action(new PageBuilder(this));
            return this;
        }

        public WordFluentDocument Paragraph(Action<ParagraphBuilder> action) {
            action(new ParagraphBuilder(this));
            return this;
        }

        public WordFluentDocument Run(Action<RunBuilder> action) {
            action(new RunBuilder(this));
            return this;
        }

        public WordFluentDocument List(Action<ListBuilder> action) {
            action(new ListBuilder(this));
            return this;
        }

        public WordFluentDocument Table(Action<TableBuilder> action) {
            action(new TableBuilder(this));
            return this;
        }

        public WordFluentDocument Image(Action<ImageBuilder> action) {
            action(new ImageBuilder(this));
            return this;
        }

        public WordFluentDocument Header(Action<HeadersBuilder> action) {
            action(new HeadersBuilder(this));
            return this;
        }

        public WordFluentDocument Footer(Action<FootersBuilder> action) {
            action(new FootersBuilder(this));
            return this;
        }

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
