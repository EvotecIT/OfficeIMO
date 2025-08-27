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

        /// <summary>
        /// Gets the underlying paragraph.
        /// </summary>
        public WordParagraph Paragraph => _paragraph;

        /// <summary>
        /// Adds a text run to the paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <param name="configure">Optional configuration for the run.</param>
        public ParagraphBuilder Text(string text, Action<TextBuilder>? configure = null) {
            var run = _paragraph.AddText(text);
            configure?.Invoke(new TextBuilder(run));
            return this;
        }

        /// <summary>
        /// Adds a text run to the paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <param name="configure">Optional configuration for the run.</param>
        public ParagraphBuilder Run(string text, Action<TextBuilder>? configure = null) => Text(text, configure);

        /// <summary>
        /// Inserts an inline image into the paragraph.
        /// </summary>
        /// <param name="path">Path to the image file.</param>
        /// <param name="widthPx">Optional width in pixels.</param>
        /// <param name="heightPx">Optional height in pixels.</param>
        /// <param name="alt">Alternative text.</param>
        public ParagraphBuilder InlineImage(string path, double? widthPx = null, double? heightPx = null, string alt = "") {
            _paragraph.AddImage(path, widthPx, heightPx, WrapTextImage.InLineWithText, alt);
            return this;
        }

        /// <summary>
        /// Adds a hyperlink to the paragraph.
        /// </summary>
        /// <param name="url">Destination URL.</param>
        /// <param name="text">Optional text to display.</param>
        /// <param name="style">True to apply hyperlink style.</param>
        public ParagraphBuilder Link(string url, string? text = null, bool style = false) {
            _paragraph.AddHyperLink(text ?? url, new Uri(url), style);
            return this;
        }

        /// <summary>
        /// Inserts a break into the paragraph.
        /// </summary>
        /// <param name="breakType">Optional break type.</param>
        public ParagraphBuilder Break(BreakValues? breakType = null) {
            _paragraph.AddBreak(breakType);
            return this;
        }

        /// <summary>
        /// Inserts a tab character.
        /// </summary>
        public ParagraphBuilder Tab() {
            _paragraph.AddTab();
            return this;
        }

        /// <summary>
        /// Sets the paragraph alignment.
        /// </summary>
        /// <param name="alignment">Desired alignment.</param>
        public ParagraphBuilder Align(HorizontalAlignment alignment) {
            _paragraph.ParagraphAlignment = alignment switch {
                HorizontalAlignment.Center => JustificationValues.Center,
                HorizontalAlignment.Right => JustificationValues.Right,
                _ => JustificationValues.Left,
            };
            return this;
        }

        /// <summary>
        /// Applies justified alignment to the paragraph.
        /// </summary>
        public ParagraphBuilder Justify() {
            _paragraph.ParagraphAlignment = JustificationValues.Both;
            return this;
        }

        /// <summary>
        /// Sets spacing before the paragraph.
        /// </summary>
        /// <param name="points">Spacing in points.</param>
        public ParagraphBuilder SpacingBefore(double points) {
            _paragraph.LineSpacingBeforePoints = points;
            return this;
        }

        /// <summary>
        /// Sets spacing after the paragraph.
        /// </summary>
        /// <param name="points">Spacing in points.</param>
        public ParagraphBuilder SpacingAfter(double points) {
            _paragraph.LineSpacingAfterPoints = points;
            return this;
        }

        /// <summary>
        /// Sets line spacing for the paragraph.
        /// </summary>
        /// <param name="points">Spacing in points.</param>
        public ParagraphBuilder LineSpacing(double points) {
            _paragraph.LineSpacingPoints = points;
            return this;
        }

        /// <summary>
        /// Sets indentation values for the paragraph.
        /// </summary>
        /// <param name="left">Left indentation in points.</param>
        /// <param name="firstLine">First-line indentation in points.</param>
        /// <param name="right">Right indentation in points.</param>
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

        /// <summary>
        /// Applies a built-in style to the paragraph.
        /// </summary>
        /// <param name="style">Built-in style.</param>
        public ParagraphBuilder Style(WordParagraphStyles style) {
            _paragraph.SetStyle(style);
            return this;
        }

        /// <summary>
        /// Applies a style by its identifier.
        /// </summary>
        /// <param name="styleId">Style identifier.</param>
        public ParagraphBuilder Style(string styleId) {
            _paragraph.SetStyleId(styleId);
            return this;
        }
    }
}
