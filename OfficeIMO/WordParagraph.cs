using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OfficeIMO {
    public class WordParagraph {
        internal WordDocument _document = null;
        internal Paragraph _paragraph = null;
        internal RunProperties _runProperties = null;
        internal Text _text = null;
        internal Run _run = null;
        internal ParagraphProperties _paragraphProperties;

        public string Text {
            get { return _text.Text; }
            set { _text.Text = value; }
        }

        public bool Bold {
            get {
                if (_runProperties.Bold != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.Bold = null;
                } else {
                    _runProperties.Bold = new Bold();
                }
            }
        }

        public bool Italic {
            get {
                if (_runProperties.Italic != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.Italic = null;
                } else {
                    _runProperties.Italic = new Italic();
                }
            }
        }

        public UnderlineValues? Underline {
            get {
                if (_runProperties.Underline.Val != null) {
                    return _runProperties.Underline.Val;
                } else {
                    return null;
                }
            }
            set {
                if (_runProperties.Underline == null) {
                    _runProperties.Underline = new Underline();
                } else {
                }

                _runProperties.Underline.Val = value;
            }
        }

        public bool DoNotCheckSpellingOrGrammar {
            get {
                if (_runProperties.NoProof != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.NoProof = null;
                } else {
                    _runProperties.NoProof = new NoProof();
                }
            }
        }

        public int? Spacing {
            get {
                if (_runProperties.Spacing != null) {
                    return _runProperties.Spacing.Val;
                } else {
                    return null;
                }
            }
            set {
                if (value != null) {
                    Spacing spacing = new Spacing();
                    spacing.Val = value;
                    _runProperties.Spacing = spacing;
                } else {
                    _runProperties.Spacing = null;
                }
            }
        }

        public bool Strike {
            get {
                if (_runProperties.Strike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.Strike = null;
                } else {
                    _runProperties.Strike = new Strike();
                }
            }
        }

        public bool DoubleStrike {
            get {
                if (_runProperties.DoubleStrike != null) {
                    return true;
                } else {
                    return false;
                }
            }
            set {
                if (value != true) {
                    _runProperties.DoubleStrike = null;
                } else {
                    _runProperties.DoubleStrike = new DoubleStrike();
                }
            }
        }

        public int? FontSize {
            get {
                if (_runProperties.FontSize != null) {
                    var fontSizeInHalfPoint = int.Parse(_runProperties.FontSize.Val);
                    return fontSizeInHalfPoint / 2;
                } else {
                    return null;
                }
            }
            set {
                if (value != null) {
                    FontSize fontSize = new FontSize();
                    fontSize.Val = (value * 2).ToString();
                    _runProperties.FontSize = fontSize;
                } else {
                    _runProperties.FontSize = null;
                }
            }
        }

        public string Color {
            get {
                if (_runProperties.Color != null) {
                    return _runProperties.Color.Val;
                } else {
                    return "";
                }
            }
            set {
                //string stringColor = value;
                // var color = System.Drawing.Color.FromArgb(Convert.ToInt32(stringColor.Substring(0, 2), 16), Convert.ToInt32(stringColor.Substring(2, 2), 16), Convert.ToInt32(stringColor.Substring(4, 2), 16));
                if (value != "") {
                    var color = new DocumentFormat.OpenXml.Wordprocessing.Color();
                    color.Val = value;
                    _runProperties.Color = color;
                } else {
                    _runProperties.Color = null;
                }
            }
        }

        public HighlightColorValues? Highlight {
            get {
                if (_runProperties.Highlight != null) {
                    return _runProperties.Highlight.Val;
                } else {
                    return null;
                }
            }
            set {
                var highlight = new Highlight();
                highlight.Val = value;
                _runProperties.Highlight = highlight;
            }
        }

        public CapsStyle CapsStyle {
            get {
                if (_runProperties.Caps != null) {
                    return CapsStyle.Caps;
                } else if (_runProperties.SmallCaps != null) {
                    return CapsStyle.SmallCaps;
                } else {
                    return CapsStyle.None;
                }
            }
            set {
                if (value == CapsStyle.None) {
                    _runProperties.Caps = null;
                    _runProperties.SmallCaps = null;
                } else if (value == CapsStyle.Caps) {
                    _runProperties.Caps = new Caps();
                } else if (value == CapsStyle.SmallCaps) {
                    _runProperties.SmallCaps = new SmallCaps();
                }
            }
        }

        public string FontFamily {
            get {
                if (_runProperties.RunFonts != null) {
                    return _runProperties.RunFonts.Ascii;
                } else {
                    return null;
                }
            }
            set {
                var runFonts = new RunFonts();
                runFonts.Ascii = value;
                _runProperties.RunFonts = runFonts;
            }
        }

        /// <summary>
        /// Alignment aka Paragraph Alignment. This element specifies the paragraph alignment which shall be applied to text in this paragraph.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then no alignment is applied to the paragraph.
        /// </summary>
        public JustificationValues ParagraphAlignment {
            get { return _paragraphProperties.Justification.Val; }
            set {
                DocumentFormat.OpenXml.Wordprocessing.Justification justification = new Justification();
                justification.Val = value;
                _paragraphProperties.Justification = justification;
            }
        }

        /// <summary>
        /// Text Alignment aka Vertical Character Alignment on Line. This element specifies the vertical alignment of all text on each line displayed within a paragraph. If the line height (before any added spacing) is larger than one or more characters on the line, all characters are aligned to each other as specified by this element.
        /// If this element is omitted on a given paragraph, its value is determined by the setting previously set at any level of the style hierarchy (i.e.that previous setting remains unchanged). If this setting is never specified in the style hierarchy, then the vertical alignment of all characters on the line shall be automatically determined by the consumer.
        /// </summary>
        public VerticalTextAlignmentValues VerticalCharacterAlignmentOnLine {
            get { return _paragraphProperties.TextAlignment.Val; }
            set {
                DocumentFormat.OpenXml.Wordprocessing.TextAlignment textAlignment = new TextAlignment();
                textAlignment.Val = value;
                _paragraphProperties.TextAlignment = textAlignment;
            }
        }

        public int? IndentationBefore {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties.Indentation != null) {
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
                if (_paragraphProperties.Indentation != null) {
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

        public int? IndentationFirstLine {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties.Indentation != null) {
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
                if (_paragraphProperties.Indentation != null) {
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
                if (_paragraphProperties.TextDirection != null) {
                    //new Indentation() { Left = "720", Right = "0", FirstLine = "0" };
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
                if (_paragraphProperties.SpacingBetweenLines != null) {
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

                spacing.After = value.ToString();
                _paragraphProperties.SpacingBetweenLines = spacing;
            }
        }

        public int? LineSpacing {
            // TODO: probably needs calculated values instead of just values
            //https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            get {
                if (_paragraphProperties.SpacingBetweenLines != null) {
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
                if (_paragraphProperties.SpacingBetweenLines != null) {
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
                if (_paragraphProperties.SpacingBetweenLines != null) {
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


        public WordParagraph(WordDocument document = null, bool newParagraph = true) {
            this._document = document;
            this._run = new Run();
            this._runProperties = new RunProperties();
            this._text = new Text {
                // this ensures spaces are preserved between runs
                Space = SpaceProcessingModeValues.Preserve
            };
            this._paragraphProperties = new ParagraphProperties();
            this._run.AppendChild(_runProperties);
            this._run.AppendChild(_text);
            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(_paragraphProperties);
                this._paragraph.AppendChild(_run);
            }

            if (document != null) {
                document.Paragraphs.Add(this);
            }
        }

        public WordParagraph(string text) {
            WordParagraph paragraph = new WordParagraph(this._document);
            paragraph.Text = text;
        }


        //public List<WordParagraph> GetParagraphs(List<Paragraph> list)
        //{
        //    var listWord = new List<WordParagraph>();
        //    //var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().ToList();
        //    foreach (Paragraph paragraph in list)
        //    {

        //        WordParagraph wordParagraph = new WordParagraph();

        //        //listWord.Add(wordParagraph);
        //        // foreach (var element in paragraph.ChildElements.OfType<Run>())
        //        // {
        //        //     
        //        //    }
        //    }

        //    return listWord;
        //}

        public WordParagraph(WordDocument document, Paragraph paragraph) {
            //_paragraph = paragraph;
            int count = 0;
            foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                RunProperties runProperties = run.RunProperties;
                Text text = run.ChildElements.OfType<Text>().First();

                if (count > 0) {
                    WordParagraph wordParagraph = new WordParagraph(this._document);
                    wordParagraph._run = run;
                    wordParagraph._text = text;
                    wordParagraph._paragraph = paragraph;
                    wordParagraph._runProperties = runProperties;
                    document.Paragraphs.Add(wordParagraph);
                } else {
                    this._run = run;
                    this._text = text;
                    this._paragraph = paragraph;
                    this._runProperties = runProperties;
                    document.Paragraphs.Add(this);
                }

                count++;
            }
        }

        public WordParagraph InsertText(string text) {
            //DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            //DocumentFormat.OpenXml.Wordprocessing.Text textProp = new DocumentFormat.OpenXml.Wordprocessing.Text();
            //textProp.Text = text;
            //run.AppendChild(textProp);
            //this.paragraph.Append(run);
            //return this;

            // WordParagraph paragraph = new WordParagraph();
            //this.paragraph.Text = text;
            //return this.paragraph;
            this._text.Text = text;
            return this;
        }

        public WordParagraph AppendText(string text) {
            WordParagraph wordParagraph = new WordParagraph(this._document, false);
            wordParagraph.Text = text;
            this._paragraph.Append(wordParagraph._run);
            return wordParagraph;
        }
    }
}