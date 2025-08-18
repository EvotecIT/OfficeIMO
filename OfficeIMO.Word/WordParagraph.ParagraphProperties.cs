using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides accessors for paragraph formatting.
    /// </summary>
    public partial class WordParagraph {
        /// <summary>
        /// Provides access to the borders applied to this paragraph.
        /// </summary>
        public WordParagraphBorders Borders {
            get {
                return new WordParagraphBorders(_document, this);
            }
        }

        /// <summary>
        /// Gets the identifier of the paragraph style, if any.
        /// </summary>
        public string? StyleId {
            get {
                if (_paragraphProperties != null && _paragraphProperties.ParagraphStyleId != null) {
                    return _paragraphProperties.ParagraphStyleId.Val;
                }
                return null;
            }
        }

        /// <summary>
        /// Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.
        /// </summary>
        public JustificationValues? ParagraphAlignment {
            get {
                if (_paragraphProperties != null)
                    if (_paragraphProperties.Justification != null)
                        return _paragraphProperties.Justification.Val?.Value;
                return null;
            }
            set {
                if (_paragraphProperties == null) {
                _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.Justification = new Justification {
                    Val = value
                };
            }
        }

        /// <summary>
        /// Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.
        /// </summary>
        public VerticalTextAlignmentValues? VerticalCharacterAlignmentOnLine {
            get {
                if (_paragraphProperties != null)
                    if (_paragraphProperties.TextAlignment != null)
                        return _paragraphProperties.TextAlignment.Val?.Value;
                return null;
            }
            set {
                var textAlignment = new TextAlignment();
                textAlignment.Val = value;
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextAlignment = textAlignment;
            }
        }
        /// <summary>
        /// Gets or sets the indentation before the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? IndentationBefore {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Left)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Left, out var left)) {
                            return left;
                        }
                    }
                }
                return null;
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Left = value?.ToString();
                _paragraphProperties.Indentation = indentation;
            }
        }

        /// <summary>
        /// Gets or sets the indentation before in points.
        /// </summary>
        public double? IndentationBeforePoints {
            get {
                if (IndentationBefore != null) {
                    return Helpers.ConvertTwipsToPoints(IndentationBefore.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    IndentationBefore = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }
        /// <summary>
        /// Gets or sets the indentation after the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? IndentationAfter {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Right)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Right, out var right)) {
                            return right;
                        }
                    }
                }
                return null;
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Right = value?.ToString();
                _paragraphProperties.Indentation = indentation;
            }
        }

        /// <summary>
        /// Gets or sets the indentation after in points.
        /// </summary>
        public double? IndentationAfterPoints {
            get {
                if (IndentationAfter != null) {
                    return Helpers.ConvertTwipsToPoints(IndentationAfter.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    IndentationAfter = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }

        /// <summary>
        /// The property which puts a paragraph on the beginning of a next side without add a page break to the document
        /// </summary>
        /// <value>bool</value>
        public bool PageBreakBefore {
            get {
                return _paragraphProperties != null && _paragraphProperties.PageBreakBefore is not null;
            }
            set {
                if (value == true) {
                    var pageBreakBefore = new PageBreakBefore();
                    _paragraphProperties.PageBreakBefore = pageBreakBefore;
                } else {
                    _paragraphProperties.PageBreakBefore = null;
                }
            }
        }
        /// <summary>
        /// Gets or sets the first line indentation in twips (1/20 of a point).
        /// </summary>
        public int? IndentationFirstLine {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.FirstLine)) {
                        if (int.TryParse(_paragraphProperties.Indentation.FirstLine, out var firstLine)) {
                            return firstLine;
                        }
                    }
                }
                return null;
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.FirstLine = value?.ToString();
                _paragraphProperties.Indentation = indentation;
            }
        }

        /// <summary>
        /// Gets or sets the first line indentation in points.
        /// </summary>
        public double? IndentationFirstLinePoints {
            get {
                if (IndentationFirstLine != null) {
                    return Helpers.ConvertTwipsToPoints(IndentationFirstLine.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    IndentationFirstLine = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }
        /// <summary>
        /// Gets or sets the hanging indentation in twips (1/20 of a point).
        /// </summary>
        public int? IndentationHanging {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.Indentation.Hanging)) {
                        if (int.TryParse(_paragraphProperties.Indentation.Hanging, out var hanging)) {
                            return hanging;
                        }
                    }
                }
                return null;
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Hanging = value?.ToString();
                _paragraphProperties.Indentation = indentation;
            }
        }

        /// <summary>
        /// Gets or sets the hanging indentation in points.
        /// </summary>
        public double? IndentationHangingPoints {
            get {
                if (IndentationHanging != null) {
                    return Helpers.ConvertTwipsToPoints(IndentationHanging.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    IndentationHanging = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }
        /// <summary>
        /// Gets or sets the text flow direction for the paragraph.
        /// </summary>
        public TextDirectionValues? TextDirection {
            get {
                if (_paragraphProperties != null && _paragraphProperties.TextDirection != null) {
                    return _paragraphProperties.TextDirection.Val?.Value;
                }
                return null;
            }
            set {
                var textDirection = new TextDirection { Val = value };
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraph.ParagraphProperties!.TextDirection = textDirection;
            }
        }

        /// <summary>
        /// Indicates that paragraph text should be displayed from right to left.
        /// </summary>
        public bool BiDi {
            get {
                return _paragraphProperties != null && _paragraphProperties.BiDi is not null;
            }
            set {
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }

                var paragraphProps = _paragraph.ParagraphProperties!;

                if (value) {
                    paragraphProps.BiDi ??= new BiDi();
                } else {
                    paragraphProps.BiDi?.Remove();
                }
            }
        }
        /// <summary>
        /// Gets or sets the rule used to calculate line spacing for the paragraph.
        /// </summary>
        public LineSpacingRuleValues? LineSpacingRule {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    if (_paragraphProperties.SpacingBetweenLines.LineRule != null) {
                        return _paragraphProperties.SpacingBetweenLines.LineRule;
                    }
                }
                return null;
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }
                spacing.LineRule = value;
                _paragraphProperties.SpacingBetweenLines = spacing;
            }
        }
        /// <summary>
        /// Gets or sets the line spacing for the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? LineSpacing {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.SpacingBetweenLines.Line)) {
                        if (int.TryParse(_paragraphProperties.SpacingBetweenLines.Line, out var line)) {
                            return line;
                        }
                    }
                }
                return null;
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.Line = value?.ToString();
                _paragraphProperties.SpacingBetweenLines = spacing;
            }
        }

        /// <summary>
        /// Gets or sets the line spacing in points.
        /// </summary>
        public double? LineSpacingPoints {
            get {
                if (LineSpacing != null) {
                    return Helpers.ConvertTwipsToPoints(LineSpacing.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    LineSpacing = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }
        /// <summary>
        /// Gets or sets the spacing before the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? LineSpacingBefore {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {                    if (!string.IsNullOrEmpty(_paragraphProperties.SpacingBetweenLines.Before)) {
                        if (int.TryParse(_paragraphProperties.SpacingBetweenLines.Before, out var before)) {
                            return before;
                        }
                    }
                }
                return null;
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.Before = value?.ToString();
                _paragraphProperties.SpacingBetweenLines = spacing;
            }
        }

        /// <summary>
        /// Gets or sets the line spacing before in points.
        /// </summary>
        public double? LineSpacingBeforePoints {
            get {
                if (LineSpacingBefore != null) {
                    return Helpers.ConvertTwipsToPoints(LineSpacingBefore.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    LineSpacingBefore = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }
        /// <summary>
        /// Gets or sets the spacing after the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? LineSpacingAfter {
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    if (!string.IsNullOrEmpty(_paragraphProperties.SpacingBetweenLines.After)) {
                        if (int.TryParse(_paragraphProperties.SpacingBetweenLines.After, out var after)) {
                            return after;
                        }
                    }
                }
                return null;
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.After = value?.ToString();
                _paragraphProperties.SpacingBetweenLines = spacing;
            }
        }

        /// <summary>
        /// Gets or sets the line spacing after in points.
        /// </summary>
        public double? LineSpacingAfterPoints {
            get {
                if (LineSpacingAfter != null) {
                    return Helpers.ConvertTwipsToPoints(LineSpacingAfter.Value);
                }
                return null;
            }
            set {
                if (value != null) {
                    LineSpacingAfter = Helpers.ConvertPointsToTwips(value.Value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the vertical text alignment - the alignment of the text in the paragraph with respect to the line height.
        /// </summary>
        public VerticalPositionValues? VerticalTextAlignment {
            get {
                if (_runProperties != null && _runProperties.VerticalTextAlignment != null) {
                    return _runProperties.VerticalTextAlignment.Val?.Value;
                }
                return null;
            }
            set {
                _runProperties ??= new RunProperties();
                if (value == null) {
                    if (_runProperties.VerticalTextAlignment == null) {
                        return;
                    }
                    _runProperties.VerticalTextAlignment = null;
                } else {
                    _runProperties.VerticalTextAlignment = new VerticalTextAlignment { Val = value };
                }
            }
        }
    }
}
