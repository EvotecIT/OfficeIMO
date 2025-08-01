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
        /// Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.
        /// </summary>
        public JustificationValues? ParagraphAlignment {
            get {
                if (_paragraphProperties != null)
                    if (_paragraphProperties.Justification != null)
                        return _paragraphProperties.Justification.Val;
                return null;
            }
            set {
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraphProperties.Justification = new Justification {
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
                        return _paragraphProperties.TextAlignment.Val;
                return null;
            }
            set {
                DocumentFormat.OpenXml.Wordprocessing.TextAlignment textAlignment = new TextAlignment();
                textAlignment.Val = value;
                if (_paragraphProperties == null) {
                    _paragraph.ParagraphProperties = new ParagraphProperties();
                }
                _paragraphProperties.TextAlignment = textAlignment;
            }
        }
        /// <summary>
        /// Gets or sets the indentation before the paragraph in twips (1/20 of a point).
        /// </summary>
        public int? IndentationBefore {
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {                    if (_paragraphProperties.Indentation.Left != "") {
                        return int.Parse(_paragraphProperties.Indentation.Left);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Left = value.ToString();
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
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {                    if (_paragraphProperties.Indentation.Right != "") {
                        return int.Parse(_paragraphProperties.Indentation.Right);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Right = value.ToString();
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
                    if (_paragraphProperties.Indentation.FirstLine != "") {
                        return int.Parse(_paragraphProperties.Indentation.FirstLine);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.FirstLine = value.ToString();
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
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {                    if (_paragraphProperties.Indentation.Hanging != "") {
                        return int.Parse(_paragraphProperties.Indentation.Hanging);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                Indentation indentation;
                if (_paragraphProperties.Indentation == null) {
                    indentation = new Indentation();
                } else {
                    indentation = _paragraphProperties.Indentation;
                }

                indentation.Hanging = value.ToString();
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
                    if (_paragraphProperties.TextDirection != null) {
                        return _paragraphProperties.TextDirection.Val;
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                TextDirection textDirection = new TextDirection();
                textDirection.Val = value;
                _paragraphProperties.TextDirection = textDirection;
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

                if (value) {
                    if (_paragraphProperties.BiDi == null) {
                        _paragraphProperties.BiDi = new BiDi();
                    }
                } else {
                    if (_paragraphProperties.BiDi != null) {
                        _paragraphProperties.BiDi.Remove();
                    }
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
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
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
                    if (_paragraphProperties.SpacingBetweenLines.Line != "") {
                        return int.Parse(_paragraphProperties.SpacingBetweenLines.Line);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.Line = value.ToString();
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
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {                    if (_paragraphProperties.SpacingBetweenLines.Before != "") {
                        return int.Parse(_paragraphProperties.SpacingBetweenLines.Before);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.Before = value.ToString();
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
                    if (_paragraphProperties.SpacingBetweenLines.After != "") {
                        return int.Parse(_paragraphProperties.SpacingBetweenLines.After);
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                SpacingBetweenLines spacing;
                if (_paragraphProperties.SpacingBetweenLines == null) {
                    spacing = new SpacingBetweenLines();
                } else {
                    spacing = _paragraphProperties.SpacingBetweenLines;
                }

                spacing.After = value.ToString();
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
                    return _runProperties.VerticalTextAlignment.Val;
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
