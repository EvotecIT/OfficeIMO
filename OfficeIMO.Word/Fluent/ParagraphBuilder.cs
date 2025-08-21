using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for paragraphs.
    /// </summary>
    public class ParagraphBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph _paragraph;

        internal ParagraphBuilder(WordFluentDocument fluent, WordParagraph paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph Paragraph => _paragraph;

        public ParagraphBuilder Text(string text, Action<TextBuilder>? configure = null) {
            var run = _paragraph.AddText(text);
            configure?.Invoke(new TextBuilder(run));
            return this;
        }

        public ParagraphBuilder Run(string text, Action<TextBuilder>? configure = null) => Text(text, configure);

        public ParagraphBuilder InlineImage(string path, double? widthPx = null, double? heightPx = null, string alt = "") {
            _paragraph.AddImage(path, widthPx, heightPx, WrapTextImage.InLineWithText, alt);
            return this;
        }

        public ParagraphBuilder Link(string url, string? text = null, bool style = false) {
            _paragraph.AddHyperLink(text ?? url, new Uri(url), style);
            return this;
        }

        public ParagraphBuilder Break(BreakValues? breakType = null) {
            _paragraph.AddBreak(breakType);
            return this;
        }

        public ParagraphBuilder Tab() {
            _paragraph.AddTab();
            return this;
        }

        /// <summary>
        /// Adds or modifies a list within the document context.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="ListBuilder"/>.</param>
        public ParagraphBuilder List(Action<ListBuilder> action) {
            action(new ListBuilder(_fluent));
            return this;
        }

        /// <summary>
        /// Adds or modifies a table within the document context.
        /// </summary>
        /// <param name="action">Action that receives a <see cref="TableBuilder"/>.</param>
        public ParagraphBuilder Table(Action<TableBuilder> action) {
            action(new TableBuilder(_fluent));
            return this;
        }
      
        public ParagraphBuilder Align(HorizontalAlignment alignment) {
            _paragraph.ParagraphAlignment = alignment switch {
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                _ => JustificationValues.Left,
            };
            return this;
        }

        public ParagraphBuilder Justify() {
            _paragraph.ParagraphAlignment = JustificationValues.Both;
            return this;
        }

        public ParagraphBuilder SpacingBefore(double points) {
            _paragraph.LineSpacingBeforePoints = points;
            return this;
        }

        public ParagraphBuilder SpacingAfter(double points) {
            _paragraph.LineSpacingAfterPoints = points;
            return this;
        }

        public ParagraphBuilder LineSpacing(double points) {
            _paragraph.LineSpacingPoints = points;
            return this;
        }

        public ParagraphBuilder Indentation(double? left = null, double? firstLine = null, double? right = null) {
            if (left != null) {
                _paragraph.IndentationBeforePoints = left.Value;
            }

            if (firstLine != null) {
                _paragraph.IndentationFirstLinePoints = firstLine.Value;
            }

            if (right != null) {
                _paragraph.IndentationAfterPoints = right.Value;
            }

            return this;
        }

        public ParagraphBuilder Style(WordParagraphStyles style) {
            _paragraph.SetStyle(style);
            return this;
        }

        public ParagraphBuilder Style(string styleId) {
            _paragraph.SetStyleId(styleId);
            return this;
        }

        public ListBuilder AddList(WordListStyle style) {
            var list = _paragraph.AddList(style);
            return new ListBuilder(_fluent, list);
        }

        public TableBuilder AddTableAfter(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            var table = _paragraph.AddTableAfter(rows, columns, tableStyle);
            return new TableBuilder(_fluent, table);
        }

        public TableBuilder AddTableBefore(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            var table = _paragraph.AddTableBefore(rows, columns, tableStyle);
            return new TableBuilder(_fluent, table);
        }
    }
}
