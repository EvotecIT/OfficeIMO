using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;
using SdtContentPicture = DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Linq;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a paragraph within a Word document.
    /// </summary>
    public partial class WordParagraph : WordElement {
        internal WordDocument _document = null!;
        internal Paragraph _paragraph = null!;

        /// <summary>
        /// This allows to know where the paragraph is located. Useful for hyperlinks or other stuff.
        /// </summary>
        internal string TopParent {
            get {
                var test = _paragraph.Parent;
                if (test == null) {
                    throw new InvalidOperationException($"Paragraph with text '{Text}' has no parent.");
                }
                if (test is Body) {
                    return "body";
                }
                if (test is Header) {
                    return "header";
                }
                if (test is Footer) {
                    return "footer";
                }
                var parent = test;
                while (!(parent is Header) && !(parent is Footer) && !(parent is Body)) {
                    parent = parent.Parent;
                    if (parent == null) {
                        throw new InvalidOperationException($"Unsupported parent chain for paragraph with text '{Text}'.");
                    }
                }
                if (parent is Body) {
                    return "body";
                }
                if (parent is Footer) {
                    return "footer";
                }
                if (parent is Header) {
                    return "header";
                }
                throw new InvalidOperationException($"Unsupported parent chain for paragraph with text '{Text}'.");
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is the last run within its parent container.
        /// </summary>
        public bool IsLastRun {
            get {
                if (_run is not null) {
                    var parent = _run.Parent;
                    if (parent != null) {
                        var runs = parent.ChildElements.OfType<Run>();
                        return runs.LastOrDefault() == _run;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is the first run within its parent container.
        /// </summary>
        public bool IsFirstRun {
            get {
                if (_run is not null) {
                    var parent = _run.Parent;
                    if (parent != null) {
                        var runs = parent.ChildElements.OfType<Run>();
                        return runs.FirstOrDefault() == _run;
                    }
                }
                return false;
            }
        }

        internal RunProperties? _runProperties {
            get {
                if (_run is not null) {
                    return _run.RunProperties;
                }

                return null;
            }
            set {
                if (_run != null) {
                    _run.RunProperties = value;
                }
            }
        }

        internal Text? _text {
            get {
                if (_run != null) {
                    return _run.ChildElements.OfType<Text>().FirstOrDefault();
                }

                return null;
            }
        }
        internal Run? _run;

        internal ParagraphProperties? _paragraphProperties {
            get {
                if (_paragraph != null && _paragraph.ParagraphProperties != null) {
                    return _paragraph.ParagraphProperties;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the first image associated with this run, if any.
        /// </summary>
        public WordImage? Image {
            get {
                if (_run != null) {
                    // DrawingML pictures
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing != null) {
                        if (drawing.Inline != null) {
                            if (drawing.Inline.Graphic != null && drawing.Inline.Graphic.GraphicData != null) {
                                var picture = drawing.Inline.Graphic.GraphicData.ChildElements
                                    .OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
                                    .FirstOrDefault();
                                if (picture != null) {
                                    return new WordImage(_document, drawing);
                                }
                            }
                        } else if (drawing.Anchor != null) {
                            var anchorGraphic = drawing.Anchor.OfType<Graphic>().FirstOrDefault();
                            if (anchorGraphic != null && anchorGraphic.GraphicData != null) {
                                var picture = anchorGraphic.GraphicData
                                    .ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
                                    .FirstOrDefault();
                                if (picture != null) {
                                    return new WordImage(_document, drawing);
                                }
                            }
                        }
                    }

                    // VML pictures
                    var vmlImage = _run.Descendants<V.ImageData>().FirstOrDefault();
                    if (vmlImage != null) {
                        var shape = vmlImage.Ancestors<V.Shape>().FirstOrDefault();
                        if (shape != null) {
                            return new WordImage(_document, _paragraph, _run, shape);
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the embedded object associated with this run, if any.
        /// </summary>
        public WordEmbeddedObject? EmbeddedObject {
            get {
                if (_run != null) {
                    var ole = _run.Descendants<Ovml.OleObject>().FirstOrDefault();
                    if (ole != null) {
                        return new WordEmbeddedObject(_document, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this paragraph is part of a numbered or bulleted list.
        /// </summary>
        public bool IsListItem {
            get {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    return true;
                } else {
                    return false;
                }
            }
        }

        /// <summary>
        /// Gets or sets the indentation level for the paragraph when it belongs to a list.
        /// </summary>
        public int? ListItemLevel {
            get {
                var val = _paragraphProperties?.NumberingProperties?.NumberingLevelReference?.Val;
                return val?.Value;
            }
            set {
                var levelRef = _paragraphProperties?.NumberingProperties?.NumberingLevelReference;
                if (levelRef != null) {
                    levelRef.Val = value;
                } else {
                    // should throw?
                }
            }
        }

        internal int? _listNumberId {
            get {
                var val = _paragraphProperties?.NumberingProperties?.NumberingId?.Val;
                return val?.Value;
            }
        }

        /// <summary>
        /// Gets the list style when this paragraph is part of a list.
        /// </summary>
        public WordListStyle? GetListStyle() {
            if (!IsListItem) return null;

            int? numberId = _listNumberId;
            if (numberId == null || _document == null) return null;

            var list = _document.Lists.FirstOrDefault(l => l._numberId == numberId);
            return list?.Style;
        }

        /// <summary>
        /// Gets the list style when this paragraph is part of a list.
        /// </summary>
        public WordListStyle? ListStyle {
            get {
                return GetListStyle();
            }
        }


        /// <summary>
        /// Gets or sets the paragraph style. Updating this to a heading style will flag the document to update the table of contents on open.
        /// </summary>
        public WordParagraphStyles? Style {
            get {
                var styleId = _paragraphProperties?.ParagraphStyleId?.Val;
                return styleId != null ? WordParagraphStyle.GetStyle(styleId.Value!) : null;
            }
            set {
                if (value != null) {
                    if (_paragraphProperties == null) {
                        _paragraph.ParagraphProperties = new ParagraphProperties();
                    }
                    var paragraphProperties = _paragraphProperties;
                    if (paragraphProperties != null) {
                        if (paragraphProperties.ParagraphStyleId == null) {
                            paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
                        }
                        paragraphProperties.ParagraphStyleId.Val = value.Value.ToStringStyle();
                        if (value.Value >= WordParagraphStyles.Heading1 && value.Value <= WordParagraphStyles.Heading9) {
                            _document?.HeadingModified();
                        }
                    }
                }
            }
        }


        internal WordList? _list;
        internal List<Run>? _runs;
        internal Hyperlink? _hyperlink;
        internal SimpleField? _simpleField;
        internal BookmarkStart? _bookmarkStart;
        internal readonly OfficeMath? _officeMath;
        internal readonly SdtRun? _stdRun;
        internal readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;

        /// <summary>
        /// Get or set a text within Paragraph
        /// </summary>
        public string Text {
            get {
                if (_text == null) {
                    return "";
                }

                return _text.Text;
            }
            set {
                var text = VerifyText();
                text.Text = value;
            }
        }

        /// <summary>
        /// Get PageBreaks within Paragraph
        /// </summary>
        public WordBreak? PageBreak {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null && brake.Type != null && brake.Type.Value == BreakValues.Page) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Get Breaks within Paragraph
        /// </summary>
        public WordBreak? Break {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the <see cref="WordTabChar"/> representing a tab character in the current run, or <c>null</c> if none is present.
        /// </summary>
        public WordTabChar? Tab {
            get {
                if (_run != null) {
                    var tabChar = _run.ChildElements.OfType<TabChar>().FirstOrDefault();
                    if (tabChar != null) {
                        return new WordTabChar(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Initializes a new paragraph.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="newParagraph">Create a new paragraph element.</param>
        /// <param name="newRun">Create a new run inside the paragraph.</param>
        public WordParagraph(WordDocument? document = null, bool newParagraph = true, bool newRun = true) {
            this._document = document!;

            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(new ParagraphProperties());

                if (newRun) {
                    this._run = new Run();
                    this._paragraph.AppendChild(_run);
                }
            }
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, bool newRun = true) {
            this._document = document;
            this._paragraph = paragraph;

            if (newRun) {
                this._run = new Run();
                this._paragraph.AppendChild(_run);
            }
        }

        /// <summary>
        /// Wraps an existing paragraph from the document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph to wrap.</param>
        public WordParagraph(WordDocument document, Paragraph paragraph) {
            this._document = document;
            this._paragraph = paragraph;
        }

        /// <summary>
        /// Wraps an existing paragraph and run from the document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph to wrap.</param>
        /// <param name="run">Run within the paragraph.</param>
        public WordParagraph(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _paragraph = paragraph;
            _run = run;
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;

            //this.Hyperlink = new WordHyperLink(document, paragraph, hyperlink);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;
            _runs = runs;
            //this.Field = new WordField(document, paragraph, runs);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;

            _simpleField = simpleField;

            //  this.Field = new WordField(document, paragraph, simpleField);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            _document = document;
            _paragraph = paragraph;

            _bookmarkStart = bookmarkStart;

            // this.Bookmark = new WordBookmark(document, paragraph, bookmarkStart);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;

            _officeMath = officeMath;

            //this.Equation = new WordEquation(document, paragraph, officeMath);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            _document = document;
            _paragraph = paragraph;
            _stdRun = stdRun;
            //this.StructuredDocumentTag = new WordStructuredDocumentTag(document, paragraph, stdRun);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _mathParagraph = mathParagraph;
            //  this.Equation = new WordEquation(document, paragraph, mathParagraph);
        }

        internal WordStructuredDocumentTag? StructuredDocumentTag {
            get {
                if (_stdRun != null) {
                    return new WordStructuredDocumentTag(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the checkbox contained in this paragraph, if present.
        /// </summary>
        public WordCheckBox? CheckBox {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>().Any() == true) {
                    return new WordCheckBox(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }


        /// <summary>
        /// Gets the date picker contained in this paragraph, if present.
        /// </summary>
        public WordDatePicker? DatePicker {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentDate>().Any() == true) {
                    return new WordDatePicker(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the dropdown list contained in this paragraph, if present.
        /// </summary>
        public WordDropDownList? DropDownList {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentDropDownList>().Any() == true) {
                    return new WordDropDownList(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the combo box contained in this paragraph, if present.
        /// </summary>
        public WordComboBox? ComboBox {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentComboBox>().Any() == true) {
                    return new WordComboBox(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the picture content control contained in this paragraph, if present.
        /// </summary>
        public WordPictureControl? PictureControl {
            get {
                if (_stdRun != null && _stdRun.SdtProperties?.Elements<SdtContentPicture>().Any() == true) {
                    return new WordPictureControl(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the repeating section contained in this paragraph, if present.
        /// </summary>
        public WordRepeatingSection? RepeatingSection =>
            _stdRun is not null && _stdRun.SdtProperties?.Elements<W15.SdtRepeatedSection>().Any() is true
                ? new WordRepeatingSection(_document, _paragraph, _stdRun)
                : null;
        /// <summary>
        /// Gets the bookmark associated with this paragraph, if present.
        /// </summary>
        public WordBookmark? Bookmark =>
            _bookmarkStart is not null ? new WordBookmark(_document, _paragraph, _bookmarkStart) : null;

        /// <summary>
        /// Gets the mathematical equation contained in this paragraph, if any.
        /// </summary>
        public WordEquation? Equation =>
            _officeMath is not null && _mathParagraph is not null ? new WordEquation(_document, _paragraph, _officeMath, _mathParagraph) :
            _officeMath is not null ? new WordEquation(_document, _paragraph, _officeMath) :
            _mathParagraph is not null ? new WordEquation(_document, _paragraph, _mathParagraph) :
            null;

        /// <summary>
        /// Gets the field contained in this paragraph, if any.
        /// </summary>
        public WordField? Field =>
            _simpleField is not null || _runs is not null ? new WordField(_document, _paragraph, _simpleField, _runs) : null;

        /// <summary>
        /// Gets the chart contained in this paragraph, if present.
        /// </summary>
        public WordChart? Chart {
            get {
                if (_run is not null) {
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing is not null) {
                        if (drawing.Inline is not null) {
                            if (drawing.Inline.Graphic is not null) {
                                if (drawing.Inline.Graphic.GraphicData is not null) {
                                    var chart = drawing.Inline.Graphic.GraphicData.ChildElements.OfType<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();
                                    if (chart is not null) {
                                        return new WordChart(_document, this, drawing);
                                    }
                                }
                            }
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the SmartArt diagram contained in this paragraph, if present.
        /// </summary>
        public WordSmartArt? SmartArt {
            get {
                if (_run is not null) {
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing is not null) {
                        var data = drawing.Descendants<GraphicData>().FirstOrDefault();
                        if (data is not null && data.Uri == "http://schemas.openxmlformats.org/drawingml/2006/diagram") {
                            return new WordSmartArt(_document, this, drawing);
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the hyperlink contained in this paragraph, if present.
        /// </summary>
        public WordHyperLink? Hyperlink =>
            _hyperlink is not null ? new WordHyperLink(_document, _paragraph, _hyperlink) : null;

        /// <summary>
        /// Gets the footnote associated with this paragraph, if any.
        /// </summary>
        public WordFootNote? FootNote {
            get {
                if (_run is not null && _runProperties is not null) {
                    var footReference = _run.ChildElements.OfType<FootnoteReference>().FirstOrDefault();
                    if (footReference is not null) {
                        return new WordFootNote(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the endnote associated with this paragraph, if any.
        /// </summary>
        public WordEndNote? EndNote {
            get {
                if (_run is not null && _runProperties is not null) {
                    var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
                    if (endNoteReference is not null) {
                        return new WordEndNote(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a hyperlink.
        /// </summary>
        public bool IsHyperLink => Hyperlink is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph hosts a field code.
        /// </summary>
        public bool IsField {
            get {
                var field = Field;
                return field is not null && field.Field is not null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph includes a bookmark start tag.
        /// </summary>
        public bool IsBookmark {
            get {
                var bookmark = Bookmark;
                return bookmark is not null && bookmark.Name is not null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains an equation.
        /// </summary>
        public bool IsEquation => Equation is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph holds a structured document tag.
        /// </summary>
        public bool IsStructuredDocumentTag => StructuredDocumentTag is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a checkbox control.
        /// </summary>
        public bool IsCheckBox => CheckBox is not null;


        /// <summary>
        /// Gets a value indicating whether the paragraph contains a date picker control.
        /// </summary>
        public bool IsDatePicker => DatePicker is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a dropdown list control.
        /// </summary>
        public bool IsDropDownList => DropDownList is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a combo box control.
        /// </summary>
        public bool IsComboBox => ComboBox is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a picture control.
        /// </summary>
        public bool IsPictureControl => PictureControl is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a repeating section control.
        /// </summary>
        public bool IsRepeatingSection => RepeatingSection is not null;
        /// <summary>
        /// Gets a value indicating whether an image is found in the paragraph.
        /// </summary>
        public bool IsImage => Image is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph has an embedded object.
        /// </summary>
        public bool IsEmbeddedObject => EmbeddedObject is not null;

        /// <summary>
        /// Gets a value indicating whether the run within the paragraph contains a tab character.
        /// </summary>
        public bool IsTab => Tab is not null;

        /// <summary>
        /// Gets a value indicating whether a chart is associated with the paragraph.
        /// </summary>
        public bool IsChart => Chart is not null;

        /// <summary>
        /// Gets a value indicating whether SmartArt is present in the paragraph.
        /// </summary>
        public bool IsSmartArt => SmartArt is not null;

        /// <summary>
        /// Gets a value indicating whether an endnote reference is present in the paragraph.
        /// </summary>
        public bool IsEndNote => EndNote is not null;

        /// <summary>
        /// Gets a value indicating whether a footnote reference is present in the paragraph.
        /// </summary>
        public bool IsFootNote => FootNote is not null;

        /// <summary>
        /// Gets all tab stops defined on the paragraph.
        /// </summary>
        public List<WordTabStop> TabStops {
            get {
                List<WordTabStop> list = new List<WordTabStop>();
            if (_paragraph is not null && _paragraphProperties is not null) {
                if (_paragraphProperties.Tabs is not null) {
                        foreach (TabStop tab in _paragraphProperties.Tabs) {
                            list.Add(new WordTabStop(this, tab));
                        }
                    }
                }
                return list;
            }
        }

        /// <summary>
        /// Gets the <see cref="WordTextBox"/> contained within the paragraph, if any.
        /// </summary>
        public WordTextBox? TextBox {
            get {
                if (_run is not null) {
                    // DrawingML text boxes
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing is not null) {
                        if (drawing.Descendants<Wps.TextBoxInfo2>().Any()) {
                            return new WordTextBox(_document, _paragraph, _run);
                        }
                    }

                    // Legacy text boxes wrapped in AlternateContent (Word 2007)
                      bool choiceHasOnlyShape = false;
                      foreach (var ac in _run.ChildElements.OfType<AlternateContent>()) {
                          var choice = ac.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                          if (choice is not null) {
                              bool choiceHasTextBox = choice.Descendants<Wps.TextBoxInfo2>().Any() || choice.Descendants<V.TextBox>().Any();
                              if (choiceHasTextBox) {
                                  return new WordTextBox(_document, _paragraph, _run);
                              }
                              bool hasShape = choice.Descendants<Wps.WordprocessingShape>().Any() ||
                                  choice.Descendants<V.Shape>().Any(s => !s.Descendants<V.ImageData>().Any() && !s.Descendants<V.TextBox>().Any());
                              if (hasShape) {
                                  choiceHasOnlyShape = true;
                                  continue;
                              }
                          }
                          var fallback = ac.ChildElements.OfType<AlternateContentFallback>().FirstOrDefault();
                          if (fallback is not null) {
                              if (fallback.Descendants<Wps.TextBoxInfo2>().Any() || fallback.Descendants<V.TextBox>().Any()) {
                                  return new WordTextBox(_document, _paragraph, _run);
                              }
                          }
                      }
                    if (choiceHasOnlyShape) {
                        return null;
                    }

                    // VML text boxes
                      if (_run.Descendants<V.TextBox>().Any()) {
                        return new WordTextBox(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Returns a <see cref="WordShape"/> instance when the paragraph contains shapes.
        /// </summary>
        public WordShape? Shape {
            get {
                if (_run is not null) {
                    if (TextBox is not null) {
                        return null;
                    }
                    // VML shapes
                    if (_run.Descendants<V.Rectangle>().Any() ||
                        _run.Descendants<V.RoundRectangle>().Any() ||
                        _run.Descendants<V.Oval>().Any() ||
                        _run.Descendants<V.Line>().Any() ||
                        _run.Descendants<V.PolyLine>().Any() ||
                        _run.Descendants<V.Shape>().Any(s => !s.Descendants<V.ImageData>().Any() && !s.Descendants<V.TextBox>().Any())) {
                        return new WordShape(_document, _paragraph, _run);
                    }

                    // DrawingML shapes (non-pictures and not text boxes)
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing is null) {
                        foreach (var ac in _run.ChildElements.OfType<AlternateContent>()) {
                            var choice = ac.ChildElements.OfType<AlternateContentChoice>().FirstOrDefault();
                            if (choice is not null) {
                                drawing = choice.Descendants<Drawing>().FirstOrDefault();
                                if (drawing is not null) {
                                    break;
                                }
                            }
                            var fallback = ac.ChildElements.OfType<AlternateContentFallback>().FirstOrDefault();
                            if (fallback is not null) {
                                drawing = fallback.Descendants<Drawing>().FirstOrDefault();
                                if (drawing is not null) {
                                    break;
                                }
                            }
                        }
                    }
                    if (drawing is not null) {
                        bool hasPicture = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().Any();
                        bool hasTextBox = drawing.Descendants<Wps.TextBoxInfo2>().Any();
                        bool hasShape = drawing.Descendants<Wps.WordprocessingShape>().Any();
                        if (!hasPicture && !hasTextBox && hasShape) {
                            return new WordShape(_document, _paragraph, _run, drawing);
                        }
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the line shape contained in this paragraph, if present.
        /// </summary>
        public WordLine? Line {
            get {
                if (_run is not null) {
                    var line = _run.Descendants<V.Line>().FirstOrDefault();
                    if (line is not null) {
                        return new WordLine(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a text box.
        /// </summary>
        public bool IsTextBox => TextBox is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a shape.
        /// </summary>
        public bool IsShape => Shape is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a VML line shape.
        /// </summary>
        public bool IsLine => Line is not null;
    }
}
