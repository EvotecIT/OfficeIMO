using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
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

        public int? IndentationBefore {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    //new Indentation() { Left = "720", Right = "0", FirstLine = "0" };
                    if (_paragraphProperties.Indentation.Left != "") {
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

        public int? IndentationAfter {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    //new Indentation() { Left = "720", Right = "0", FirstLine = "0" };
                    if (_paragraphProperties.Indentation.Right != "") {
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

        public int? IndentationFirstLine {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
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

        public int? IndentationHanging {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties != null && _paragraphProperties.Indentation != null) {
                    //new Indentation() { Left = "720", Right = "0", FirstLine = "0" };
                    if (_paragraphProperties.Indentation.Hanging != "") {
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

        public TextDirectionValues? TextDirection {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
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

        public int? LineSpacing {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
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

        public int? LineSpacingBefore {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties != null && _paragraphProperties.SpacingBetweenLines != null) {
                    //new Indentation() { Left = "720", Right = "0", FirstLine = "0" };
                    if (_paragraphProperties.SpacingBetweenLines.Before != "") {
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

        public int? LineSpacingAfter {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
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
    }
}
