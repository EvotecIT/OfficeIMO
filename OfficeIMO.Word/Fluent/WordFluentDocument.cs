using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Provides a fluent API wrapper around <see cref="WordDocument"/>.
    /// </summary>
    public class WordFluentDocument {
        internal WordDocument Document { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordFluentDocument"/> class.
        /// </summary>
        /// <param name="document">The underlying <see cref="WordDocument"/>.</param>
        public WordFluentDocument(WordDocument document) {
            Document = document ?? throw new ArgumentNullException(nameof(document));
        }

        /// <summary>
        /// Provides fluent access to document information.
        /// </summary>
        /// <param name="action">Action that receives an <see cref="InfoBuilder"/>.</param>
        public WordFluentDocument Info(Action<InfoBuilder> action) {
            action(new InfoBuilder(this));
            return this;
        }

        /// <summary>
        /// Configures document sections.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="SectionBuilder"/>.</param>
        public WordFluentDocument Section(Action<SectionBuilder> action) {
            action(new SectionBuilder(this));
            return this;
        }

        /// <summary>
        /// Configures page settings.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="PageBuilder"/>.</param>
        public WordFluentDocument Page(Action<PageBuilder> action) {
            action(new PageBuilder(this));
            return this;
        }

        /// <summary>
        /// Adds a new paragraph and allows fluent configuration of its contents.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="ParagraphBuilder"/>.</param>
        public WordFluentDocument Paragraph(Action<ParagraphBuilder> action) {
            var paragraph = Document.AddParagraph();
            action(new ParagraphBuilder(this, paragraph));
            return this;
        }

        /// <summary>
        /// Adds or modifies a list.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="ListBuilder"/>.</param>
        public WordFluentDocument List(Action<ListBuilder> action) {
            action(new ListBuilder(this));
            return this;
        }

        /// <summary>
        /// Adds or modifies a table.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="TableBuilder"/>.</param>
        public WordFluentDocument Table(Action<TableBuilder> action) {
            action(new TableBuilder(this));
            return this;
        }

        /// <summary>
        /// Adds or modifies an image.
        /// </summary>
        /// <param name="action">Action that receives an <see cref="ImageBuilder"/>.</param>
        public WordFluentDocument Image(Action<ImageBuilder> action) {
            action(new ImageBuilder(this));
            return this;
        }

        /// <summary>
        /// Configures document headers.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="HeadersBuilder"/>.</param>
        public WordFluentDocument Header(Action<HeadersBuilder> action) {
            action(new HeadersBuilder(this));
            return this;
        }

        /// <summary>
        /// Configures document footers.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="FootersBuilder"/>.</param>
        public WordFluentDocument Footer(Action<FootersBuilder> action) {
            action(new FootersBuilder(this));
            return this;
        }

        /// <summary>
        /// Executes an action for each paragraph in the document.
        /// </summary>
        /// <param name="action">Action to execute for every paragraph.</param>
        public WordFluentDocument ForEachParagraph(Action<ParagraphBuilder> action) {
            Document.ForEachParagraph(p => action(new ParagraphBuilder(this, p)));
            return this;
        }

        /// <summary>
        /// Finds paragraphs containing the specified text.
        /// </summary>
        /// <param name="text">Text to search for.</param>
        /// <param name="action">Action executed for each matching paragraph.</param>
        /// <param name="stringComparison">String comparison option.</param>
        public WordFluentDocument Find(string text, Action<ParagraphBuilder> action, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            foreach (var paragraph in Document.FindParagraphs(text, stringComparison)) {
                action(new ParagraphBuilder(this, paragraph));
            }
            return this;
        }

        /// <summary>
        /// Selects paragraphs that match the specified predicate.
        /// </summary>
        /// <param name="predicate">Filter predicate.</param>
        public IEnumerable<ParagraphBuilder> Select(Func<ParagraphBuilder, bool> predicate) {
            foreach (var paragraph in Document.SelectParagraphs(p => predicate(new ParagraphBuilder(this, p)))) {
                yield return new ParagraphBuilder(this, paragraph);
            }
        }
    }
}
