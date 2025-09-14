using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeIMO.Word;

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
        /// Configures document-wide page setup defaults.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="PageSetupBuilder"/>.</param>
        public WordFluentDocument PageSetup(Action<PageSetupBuilder> action) {
            action(new PageSetupBuilder(this));
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
        /// Adds or modifies an image asynchronously.
        /// </summary>
        /// <param name="action">Async action that receives an <see cref="ImageBuilder"/>.</param>
        public async Task<WordFluentDocument> ImageAsync(Func<ImageBuilder, Task> action) {
            await action(new ImageBuilder(this));
            return this;
        }

        /// <summary>
        /// Adds a plain paragraph (Markdown-style alias of <see cref="Paragraph(Action{ParagraphBuilder})"/>).
        /// </summary>
        /// <param name="text">Text to insert.</param>
        public WordFluentDocument P(string text) {
            var p = Document.AddParagraph(text);
            return this;
        }

        /// <summary>
        /// Adds a heading paragraph styled as Heading 1 (alias mirrors MarkdownDoc.H1).
        /// </summary>
        public WordFluentDocument H1(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading1); return this; }
        /// <summary>
        /// Adds a heading paragraph styled as Heading 2 (alias mirrors MarkdownDoc.H2).
        /// </summary>
        public WordFluentDocument H2(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading2); return this; }
        /// <summary>
        /// Adds a heading paragraph styled as Heading 3 (alias mirrors MarkdownDoc.H3).
        /// </summary>
        public WordFluentDocument H3(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading3); return this; }
        /// <summary>
        /// Adds a heading paragraph styled as Heading 4 (alias mirrors MarkdownDoc.H4).
        /// </summary>
        public WordFluentDocument H4(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading4); return this; }
        /// <summary>
        /// Adds a heading paragraph styled as Heading 5 (alias mirrors MarkdownDoc.H5).
        /// </summary>
        public WordFluentDocument H5(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading5); return this; }
        /// <summary>
        /// Adds a heading paragraph styled as Heading 6 (alias mirrors MarkdownDoc.H6).
        /// </summary>
        public WordFluentDocument H6(string text) { var p = Document.AddParagraph(text); p.SetStyle(WordParagraphStyles.Heading6); return this; }

        /// <summary>
        /// Adds a bulleted list using the list builder (alias mirrors MarkdownDoc.Ul).
        /// </summary>
        public WordFluentDocument Ul(Action<ListBuilder> action) {
            var lb = new ListBuilder(this).Bulleted();
            action(lb);
            return this;
        }

        /// <summary>
        /// Adds a numbered list using the list builder (alias mirrors MarkdownDoc.Ol).
        /// </summary>
        public WordFluentDocument Ol(Action<ListBuilder> action) {
            var lb = new ListBuilder(this).Numbered();
            action(lb);
            return this;
        }

        /// <summary>
        /// Adds a monospace code paragraph (alias mirrors MarkdownDoc.Code).
        /// </summary>
        /// <param name="language">Optional language hint (not used for styling).</param>
        /// <param name="content">Code text.</param>
        public WordFluentDocument Code(string language, string content) {
            var p = Document.AddParagraph(content);
            var mono = Helpers.FontResolver.Resolve("monospace") ?? "Consolas";
            p.SetFontFamily(mono);
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
        /// Inserts a Table of Contents near the top of the document (Markdown parity: TocAtTop).
        /// </summary>
        /// <param name="title">Heading shown above the TOC.</param>
        /// <param name="minLevel">Minimum heading level to include.</param>
        /// <param name="maxLevel">Maximum heading level to include.</param>
        /// <param name="titleLevel">Heading level for the title (1..6).</param>
        public WordFluentDocument TocAtTop(string title = "Contents", int minLevel = 1, int maxLevel = 3, int titleLevel = 2) {
            // Insert heading at top
            WordParagraph heading;
            if (Document.Paragraphs.Count > 0) {
                heading = Document.Paragraphs[0].AddParagraphBeforeSelf();
            } else {
                heading = Document.AddParagraph();
            }
            heading.Text = title;
            heading.SetStyle((WordParagraphStyles)Math.Max((int)WordParagraphStyles.Heading1, Math.Min((int)WordParagraphStyles.Heading9, (int)WordParagraphStyles.Heading1 + (titleLevel - 1))));

            // Add TOC field right after heading with desired levels
            var tocPara = heading.AddParagraphAfterSelf();
            var builder = new WordFieldBuilder(WordFieldType.TOC)
                .AddSwitch($"\\o \"{minLevel}-{maxLevel}\"")
                .AddSwitch("\\h")
                .AddSwitch("\\z")
                .AddSwitch("\\u");
            tocPara.AddField(builder);
            Document.Settings.UpdateFieldsOnOpen = true;
            return this;
        }

        /// <summary>
        /// Appends a Table of Contents at the end (Markdown parity: TocHere).
        /// </summary>
        public WordFluentDocument TocHere(string title = "Contents", int minLevel = 1, int maxLevel = 3, int titleLevel = 3) {
            var heading = Document.AddParagraph(title);
            heading.SetStyle((WordParagraphStyles)Math.Max((int)WordParagraphStyles.Heading1, Math.Min((int)WordParagraphStyles.Heading9, (int)WordParagraphStyles.Heading1 + (titleLevel - 1))));
            var tocPara = Document.AddParagraph();
            var builder = new WordFieldBuilder(WordFieldType.TOC)
                .AddSwitch($"\\o \"{minLevel}-{maxLevel}\"")
                .AddSwitch("\\h")
                .AddSwitch("\\z")
                .AddSwitch("\\u");
            tocPara.AddField(builder);
            Document.Settings.UpdateFieldsOnOpen = true;
            return this;
        }

        /// <summary>
        /// Creates a simple callout block (Markdown parity: Callout(kind,title,body)).
        /// </summary>
        public WordFluentDocument Callout(string kind, string title, string body) {
            var color = kind?.ToLowerInvariant() switch {
                "info" => SixLabors.ImageSharp.Color.LightBlue,
                "warning" => SixLabors.ImageSharp.Color.Khaki,
                "danger" => SixLabors.ImageSharp.Color.MistyRose,
                "tip" => SixLabors.ImageSharp.Color.Honeydew,
                _ => SixLabors.ImageSharp.Color.LightGray
            };
            var p1 = Document.AddParagraph();
            p1.ShadingFillColor = color;
            p1.AddFormattedText(title, bold: true);
            var p2 = Document.AddParagraph(body);
            p2.ShadingFillColor = color;
            return this;
        }

        /// <summary>
        /// Executes an action for each section in the document.
        /// </summary>
        /// <param name="action">Action to execute for every section with its 1-based index.</param>
        public WordFluentDocument ForEachSection(Action<int, WordSection> action) {
            for (int i = 0; i < Document.Sections.Count; i++) {
                action(i + 1, Document.Sections[i]);
            }
            return this;
        }

        /// <summary>
        /// Returns all sections in the document.
        /// </summary>
        public IEnumerable<WordSection> Sections() {
            return Document.Sections;
        }

        /// <summary>
        /// Returns all paragraphs in the document.
        /// </summary>
        public IEnumerable<WordParagraph> Paragraphs() {
            return Document.Paragraphs;
        }

        /// <summary>
        /// Returns all tables in the document.
        /// </summary>
        public IEnumerable<WordTable> Tables() {
            return Document.Tables;
        }

        /// <summary>
        /// Ends fluent configuration and returns the underlying <see cref="WordDocument"/>.
        /// </summary>
        /// <returns>The wrapped <see cref="WordDocument"/> for further processing.</returns>
        public WordDocument End() {
            return Document;
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
        /// Executes an action for each run in the document.
        /// </summary>
        /// <param name="action">Action to execute for every run.</param>
        public WordFluentDocument ForEachRun(Action<RunBuilder> action) {
            Document.ForEachRun(r => action(new RunBuilder(r)));
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
        /// Finds runs matching the specified regular expression pattern.
        /// </summary>
        /// <param name="pattern">Regular expression pattern.</param>
        /// <param name="action">Action executed for each matching run.</param>
        public WordFluentDocument FindRegex(string pattern, Action<ParagraphBuilder> action) {
            foreach (var run in Document.FindRunsRegex(pattern)) {
                action(new ParagraphBuilder(this, run));
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
